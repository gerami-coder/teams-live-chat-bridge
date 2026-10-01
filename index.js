import express from 'express';
import crypto from 'node:crypto';
import { App, ExpressAdapter } from '@microsoft/teams.apps';
import { MessageActivityInput } from '@microsoft/teams.api';
import { createClient } from '@supabase/supabase-js';

const server = express();
server.disable('x-powered-by');
server.use(express.json({ limit: '256kb' }));

const requiredEnv = [
  'CLIENT_ID',
  'CLIENT_SECRET',
  'TENANT_ID',
  'SUPABASE_URL',
  'SUPABASE_SERVICE_ROLE_KEY',
];

for (const key of requiredEnv) {
  if (!process.env[key]) {
    console.warn(`[config] Missing environment variable: ${key}`);
  }
}

const supabase = createClient(
  process.env.SUPABASE_URL || 'https://invalid.local',
  process.env.SUPABASE_SERVICE_ROLE_KEY || 'invalid',
  {
    auth: {
      persistSession: false,
      autoRefreshToken: false,
    },
  },
);

const app = new App({
  clientId: process.env.CLIENT_ID,
  clientSecret: process.env.CLIENT_SECRET,
  tenantId: process.env.TENANT_ID,
  manifest: {
    name: {
      short: process.env.APP_NAME || 'Web Chat',
      full: process.env.APP_NAME || 'Website Live Chat',
    },
  },
  httpServerAdapter: new ExpressAdapter(server),
});

function normalizeHost(value) {
  try {
    return new URL(value).hostname.toLowerCase().replace(/^www\./, '');
  } catch {
    return '';
  }
}

function safeEqual(a, b) {
  const aa = Buffer.from(String(a || ''));
  const bb = Buffer.from(String(b || ''));
  return aa.length === bb.length && crypto.timingSafeEqual(aa, bb);
}

function slugify(value) {
  return String(value || 'agent')
    .toLowerCase()
    .trim()
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/^-|-$/g, '')
    .slice(0, 40) || 'agent';
}

async function resolveSite(pageUrl, siteKey) {
  if (siteKey) {
    const { data } = await supabase
      .from('sites')
      .select('*')
      .eq('site_key', siteKey)
      .eq('is_active', true)
      .maybeSingle();

    if (data) return data;
  }

  const host = normalizeHost(pageUrl);
  if (!host) return null;

  const { data } = await supabase
    .from('sites')
    .select('*')
    .eq('is_active', true);

  return (data || []).find((site) => normalizeHost(site.base_url) === host) || null;
}

async function upsertAgent(activity) {
  const aadObjectId = activity.from?.aadObjectId || null;
  const conversationId = activity.conversation?.id || null;
  const displayName = activity.from?.name || 'Support Agent';

  if (!conversationId) return null;

  let existing = null;

  if (aadObjectId) {
    existing = (
      await supabase
        .from('agents')
        .select('*')
        .eq('aad_object_id', aadObjectId)
        .maybeSingle()
    ).data;
  }

  if (!existing) {
    existing = (
      await supabase
        .from('agents')
        .select('*')
        .eq('teams_conversation_id', conversationId)
        .maybeSingle()
    ).data;
  }

  if (existing) {
    const { data, error } = await supabase
      .from('agents')
      .update({
        display_name: displayName,
        aad_object_id: aadObjectId || existing.aad_object_id,
        teams_conversation_id: conversationId,
        status: existing.status === 'offline' ? 'offline' : 'online',
        last_seen_at: new Date().toISOString(),
        updated_at: new Date().toISOString(),
      })
      .eq('id', existing.id)
      .select('*')
      .single();

    if (error) throw error;
    return data;
  }

  const suffix = (aadObjectId || conversationId)
    .replace(/[^a-zA-Z0-9]/g, '')
    .slice(-6)
    .toLowerCase();

  const { data, error } = await supabase
    .from('agents')
    .insert({
      agent_key: `${slugify(displayName)}-${suffix}`,
      display_name: displayName,
      aad_object_id: aadObjectId,
      teams_conversation_id: conversationId,
      status: 'online',
      last_seen_at: new Date().toISOString(),
    })
    .select('*')
    .single();

  if (error) throw error;
  return data;
}

async function getMappedAgentIds(siteId, department = 'general') {
  const { data } = await supabase
    .from('site_agents')
    .select('agent_id,department,priority')
    .eq('site_id', siteId)
    .eq('is_active', true)
    .order('priority', { ascending: true });

  const mappings = data || [];
  const exact = mappings.filter((m) => m.department === department);
  const general = mappings.filter((m) => m.department === 'general');
  const selected = exact.length ? exact : general;

  return [...new Set(selected.map((m) => m.agent_id))];
}

async function chooseAgent(siteId, department = 'general') {
  const mappedIds = await getMappedAgentIds(siteId, department);

  let query = supabase
    .from('agents')
    .select('*')
    .eq('is_active', true)
    .in('status', ['online', 'away'])
    .order('last_assigned_at', { ascending: true, nullsFirst: true })
    .limit(1);

  if (mappedIds.length) query = query.in('id', mappedIds);

  let { data } = await query;

  if (!data?.length) {
    let fallback = supabase
      .from('agents')
      .select('*')
      .eq('is_active', true)
      .order('last_assigned_at', { ascending: true, nullsFirst: true })
      .limit(1);

    if (mappedIds.length) fallback = fallback.in('id', mappedIds);
    data = (await fallback).data;
  }

  const agent = data?.[0] || null;

  if (agent) {
    await supabase
      .from('agents')
      .update({ last_assigned_at: new Date().toISOString() })
      .eq('id', agent.id);
  }

  return agent;
}

function buildChatCard({
  siteName,
  visitorName,
  visitorEmail,
  message,
  pageUrl,
  conversationUuid,
}) {
  return {
    $schema: 'https://adaptivecards.io/schemas/adaptive-card.json',
    type: 'AdaptiveCard',
    version: '1.5',
    msteams: { width: 'Full' },
    body: [
      {
        type: 'TextBlock',
        text: siteName,
        size: 'Medium',
        weight: 'Bolder',
        wrap: true,
      },
      {
        type: 'TextBlock',
        text: 'New website conversation',
        spacing: 'None',
        isSubtle: true,
        wrap: true,
      },
      {
        type: 'FactSet',
        separator: true,
        facts: [
          { title: 'Visitor', value: visitorName },
          ...(visitorEmail ? [{ title: 'Email', value: visitorEmail }] : []),
          { title: 'Conversation', value: conversationUuid },
        ],
      },
      {
        type: 'TextBlock',
        text: message,
        wrap: true,
        separator: true,
        size: 'Medium',
        weight: 'Bolder',
      },
      {
        type: 'TextBlock',
        text: 'Reply directly in this Teams chat.',
        isSubtle: true,
        wrap: true,
      },
    ],
    actions: pageUrl
      ? [
          {
            type: 'Action.OpenUrl',
            title: 'Open website',
            url: pageUrl,
          },
        ]
      : [],
  };
}

app.on('install.add', async ({ activity, send }) => {
  try {
    const agent = await upsertAgent(activity);
    await send(
      agent
        ? `Website Live Chat is connected for ${agent.display_name}. Use /online, /away or /offline to change your availability.`
        : 'Website Live Chat is connected.',
    );
  } catch (error) {
    console.error('[install.add]', error);
    await send('Website Live Chat connected, but agent registration failed.');
  }
});

app.on('message', async ({ activity, send }) => {
  try {
    const text = String(activity.text || '').trim();
    const agent = await upsertAgent(activity);

    if (!agent) {
      await send('I could not register this Teams user.');
      return;
    }

    const command = text.toLowerCase();

    if (['/online', '/away', '/offline'].includes(command)) {
      const status = command.slice(1);
      await supabase
        .from('agents')
        .update({
          status,
          last_seen_at: new Date().toISOString(),
          updated_at: new Date().toISOString(),
        })
        .eq('id', agent.id);

      await send(`Your website-chat status is now ${status}.`);
      return;
    }

    if (!text) return;

    const { data: conversation, error } = await supabase
      .from('conversations')
      .select('*, sites(*)')
      .eq('assigned_agent_id', agent.id)
      .in('status', ['open', 'pending'])
      .order('last_message_at', { ascending: false })
      .limit(1)
      .maybeSingle();

    if (error) throw error;

    if (!conversation) {
      await send(
        'No active website conversation is assigned to you. Commands: /online, /away, /offline',
      );
      return;
    }

    const site = conversation.sites;

    if (!site?.shared_secret) {
      await send('This website is missing its callback secret.');
      return;
    }

    const response = await fetch(conversation.reply_url, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'X-TLC-Secret': site.shared_secret,
      },
      body: JSON.stringify({
        conversation_uuid: conversation.conversation_uuid,
        message: text,
        sender_name: activity.from?.name || 'Support',
        external_id: activity.id,
      }),
    });

    if (!response.ok) {
      throw new Error(`WordPress replied ${response.status}`);
    }

    await supabase.from('messages').insert({
      conversation_id: conversation.id,
      sender_type: 'agent',
      sender_name: activity.from?.name || 'Support',
      body: text,
      external_id: activity.id,
    });

    await supabase
      .from('conversations')
      .update({
        last_message_at: new Date().toISOString(),
        updated_at: new Date().toISOString(),
      })
      .eq('id', conversation.id);
  } catch (error) {
    console.error('[message]', error);
    await send('I could not deliver that reply back to the website.');
  }
});

server.get('/', (_req, res) => {
  res.json({
    ok: true,
    service: 'teams-live-chat-bridge',
  });
});

server.get('/health', async (_req, res) => {
  try {
    const { count, error } = await supabase
      .from('sites')
      .select('*', { count: 'exact', head: true });

    if (error) throw error;

    res.json({
      ok: true,
      database: 'connected',
      sites: count ?? 0,
    });
  } catch (error) {
    console.error('[health]', error);
    res.status(503).json({
      ok: false,
      database: 'unavailable',
    });
  }
});

server.post('/wp/message', async (req, res) => {
  try {
    const {
      conversation_uuid: conversationUuid,
      site_key: siteKey,
      visitor,
      page_url: pageUrl,
      message,
      reply_url: replyUrl,
      department = 'general',
    } = req.body || {};

    if (!conversationUuid || !message || !replyUrl || !pageUrl) {
      return res.status(400).json({ error: 'invalid_payload' });
    }

    const site = await resolveSite(String(pageUrl), siteKey ? String(siteKey) : undefined);

    if (!site) {
      return res.status(404).json({ error: 'site_not_registered' });
    }

    const suppliedSecret = req.header('X-TLC-Secret') || '';

    if (
      !site.shared_secret ||
      !suppliedSecret ||
      !safeEqual(site.shared_secret, suppliedSecret)
    ) {
      return res.status(403).json({ error: 'forbidden' });
    }

    let { data: conversation } = await supabase
      .from('conversations')
      .select('*')
      .eq('conversation_uuid', conversationUuid)
      .maybeSingle();

    let assignedAgent = null;

    if (!conversation) {
      assignedAgent = await chooseAgent(site.id, department);

      const { data, error } = await supabase
        .from('conversations')
        .insert({
          conversation_uuid: conversationUuid,
          site_id: site.id,
          assigned_agent_id: assignedAgent?.id || null,
          visitor_name: visitor?.name || 'Visitor',
          visitor_email: visitor?.email || null,
          visitor_phone: visitor?.phone || null,
          page_url: pageUrl,
          department,
          status: assignedAgent ? 'open' : 'missed',
          reply_url: replyUrl,
          last_message_at: new Date().toISOString(),
        })
        .select('*')
        .single();

      if (error) throw error;
      conversation = data;
    } else {
      if (conversation.assigned_agent_id) {
        assignedAgent = (
          await supabase
            .from('agents')
            .select('*')
            .eq('id', conversation.assigned_agent_id)
            .maybeSingle()
        ).data;
      }

      await supabase
        .from('conversations')
        .update({
          visitor_name: visitor?.name || conversation.visitor_name,
          visitor_email: visitor?.email || conversation.visitor_email,
          page_url: pageUrl,
          reply_url: replyUrl,
          last_message_at: new Date().toISOString(),
          updated_at: new Date().toISOString(),
        })
        .eq('id', conversation.id);
    }

    await supabase.from('messages').insert({
      conversation_id: conversation.id,
      sender_type: 'visitor',
      sender_name: visitor?.name || 'Visitor',
      body: String(message),
    });

    if (!assignedAgent?.teams_conversation_id) {
      return res.status(202).json({
        ok: true,
        queued: true,
        reason: 'no_agent_available',
      });
    }

    const card = buildChatCard({
      siteName: site.name,
      visitorName: visitor?.name || 'Visitor',
      visitorEmail: visitor?.email || undefined,
      message: String(message),
      pageUrl: String(pageUrl),
      conversationUuid: String(conversationUuid),
    });

    const outbound = new MessageActivityInput();
    outbound.addCard('adaptive', card);

    await app.send(assignedAgent.teams_conversation_id, outbound);

    return res.json({
      ok: true,
      assigned_agent: assignedAgent.display_name,
    });
  } catch (error) {
    console.error('[wp/message]', error);
    return res.status(500).json({
      error: 'bridge_error',
      detail: error instanceof Error ? error.message : String(error),
    });
  }
});

await app.initialize();

export default server;
