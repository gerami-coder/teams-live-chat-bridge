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
    const incomingIsPersonal = String(conversationId).startsWith('a:');
    const existingIsPersonal = String(existing.teams_conversation_id || '').startsWith('a:');
    const safeConversationId =
      incomingIsPersonal || !existingIsPersonal
        ? conversationId
        : existing.teams_conversation_id;

    const { data, error } = await supabase
      .from('agents')
      .update({
        display_name: displayName,
        aad_object_id: aadObjectId || existing.aad_object_id,
        teams_conversation_id: safeConversationId,
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
    actions: [
      {
        type: 'Action.OpenUrl',
        title: 'Open in Inbox',
        url: `https://teams-live-chat-bridge.vercel.app/tab?conversation=${encodeURIComponent(conversationUuid)}`,
      },
      ...(pageUrl
        ? [
            {
              type: 'Action.OpenUrl',
              title: 'Open website',
              url: pageUrl,
            },
          ]
        : []),
    ],
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

    const { data: activeConversations, error } = await supabase
      .from('conversations')
      .select('*, sites(*)')
      .eq('assigned_agent_id', agent.id)
      .in('status', ['open', 'pending'])
      .order('last_message_at', { ascending: false })
      .limit(3);

    if (error) throw error;

    if (!activeConversations?.length) {
      await send(
        'No active website conversation is assigned to you. Open the Inbox tab to view all conversations.',
      );
      return;
    }

    if (activeConversations.length > 1) {
      await send(
        'You have multiple active website conversations. Use the Inbox tab to choose the visitor before replying.',
      );
      return;
    }

    const conversation = activeConversations[0];
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


server.post('/wp/register-site', async (req, res) => {
  try {
    const suppliedSecret = req.header('X-TLC-Secret') || '';
    const {
      site_url: siteUrl,
      site_name: siteName,
      site_key: requestedSiteKey,
    } = req.body || {};

    if (!siteUrl || !suppliedSecret) {
      return res.status(400).json({ error: 'invalid_payload' });
    }

    const host = normalizeHost(String(siteUrl));
    if (!host) return res.status(400).json({ error: 'invalid_site_url' });

    const siteKey = slugify(requestedSiteKey || host.replace(/\./g, '-'));
    const { data: allSites, error: siteError } = await supabase
      .from('sites')
      .select('*');

    if (siteError) throw siteError;

    let site = (allSites || []).find((row) => normalizeHost(row.base_url) === host) || null;

    if (site) {
      if (site.shared_secret && !safeEqual(site.shared_secret, suppliedSecret)) {
        return res.status(403).json({ error: 'forbidden' });
      }

      const { data, error } = await supabase
        .from('sites')
        .update({
          name: String(siteName || site.name || host),
          base_url: String(siteUrl),
          site_key: site.site_key || siteKey,
          shared_secret: suppliedSecret,
          secret_hash: crypto.createHash('sha256').update(suppliedSecret).digest('hex'),
          is_active: true,
          updated_at: new Date().toISOString(),
        })
        .eq('id', site.id)
        .select('*')
        .single();

      if (error) throw error;
      site = data;
    } else {
      const { data, error } = await supabase
        .from('sites')
        .insert({
          site_key: siteKey,
          name: String(siteName || host),
          base_url: String(siteUrl),
          shared_secret: suppliedSecret,
          secret_hash: crypto.createHash('sha256').update(suppliedSecret).digest('hex'),
          is_active: true,
        })
        .select('*')
        .single();

      if (error) throw error;
      site = data;
    }

    return res.json({
      ok: true,
      site: {
        id: site.id,
        site_key: site.site_key,
        name: site.name,
        base_url: site.base_url,
      },
    });
  } catch (error) {
    console.error('[wp/register-site]', error);
    return res.status(500).json({
      error: 'bridge_error',
      detail: error instanceof Error ? error.message : String(error),
    });
  }
});

server.get('/wp/bootstrap', async (req, res) => {
  try {
    const siteUrl = String(req.query.site_url || '');
    const siteKey = req.query.site_key ? String(req.query.site_key) : undefined;
    const suppliedSecret = req.header('X-TLC-Secret') || '';

    const site = await resolveSite(siteUrl, siteKey);
    if (!site) return res.status(404).json({ error: 'site_not_registered' });

    if (!site.shared_secret || !suppliedSecret || !safeEqual(site.shared_secret, suppliedSecret)) {
      return res.status(403).json({ error: 'forbidden' });
    }

    const [{ data: agents, error: agentsError }, { data: sites, error: sitesError }, { data: mappings, error: mappingsError }] = await Promise.all([
      supabase
        .from('agents')
        .select('id,agent_key,display_name,status,is_active,last_seen_at')
        .eq('is_active', true)
        .order('display_name', { ascending: true }),
      supabase
        .from('sites')
        .select('id,site_key,name,base_url,is_active')
        .eq('is_active', true)
        .order('name', { ascending: true }),
      supabase
        .from('site_agents')
        .select('agent_id,department,priority,is_active')
        .eq('site_id', site.id)
        .eq('is_active', true)
        .order('priority', { ascending: true }),
    ]);

    if (agentsError) throw agentsError;
    if (sitesError) throw sitesError;
    if (mappingsError) throw mappingsError;

    const activeAgents = agents || [];
    const onlineCount = activeAgents.filter((a) => a.status === 'online').length;
    const availableCount = activeAgents.filter((a) => a.status === 'online' || a.status === 'away').length;

    return res.json({
      ok: true,
      site: {
        id: site.id,
        site_key: site.site_key,
        name: site.name,
        base_url: site.base_url,
      },
      agents: activeAgents,
      assigned_agent_ids: (mappings || []).map((m) => m.agent_id),
      sites: sites || [],
      availability: {
        online: onlineCount,
        available: availableCount,
        total: activeAgents.length,
      },
    });
  } catch (error) {
    console.error('[wp/bootstrap]', error);
    return res.status(500).json({
      error: 'bridge_error',
      detail: error instanceof Error ? error.message : String(error),
    });
  }
});

server.post('/wp/assign-agents', async (req, res) => {
  try {
    const suppliedSecret = req.header('X-TLC-Secret') || '';
    const {
      site_url: siteUrl,
      site_key: siteKey,
      agent_ids: agentIds,
      department = 'general',
    } = req.body || {};

    if (!siteUrl || !Array.isArray(agentIds)) {
      return res.status(400).json({ error: 'invalid_payload' });
    }

    const site = await resolveSite(String(siteUrl), siteKey ? String(siteKey) : undefined);
    if (!site) return res.status(404).json({ error: 'site_not_registered' });

    if (!site.shared_secret || !suppliedSecret || !safeEqual(site.shared_secret, suppliedSecret)) {
      return res.status(403).json({ error: 'forbidden' });
    }

    const cleanIds = [...new Set(agentIds.map(String).filter(Boolean))];

    const { error: deleteError } = await supabase
      .from('site_agents')
      .delete()
      .eq('site_id', site.id)
      .eq('department', String(department));

    if (deleteError) throw deleteError;

    if (cleanIds.length) {
      const rows = cleanIds.map((agentId, index) => ({
        site_id: site.id,
        agent_id: agentId,
        department: String(department),
        priority: index + 1,
        is_active: true,
      }));

      const { error: insertError } = await supabase
        .from('site_agents')
        .insert(rows);

      if (insertError) throw insertError;
    }

    return res.json({ ok: true, assigned_agent_ids: cleanIds });
  } catch (error) {
    console.error('[wp/assign-agents]', error);
    return res.status(500).json({
      error: 'bridge_error',
      detail: error instanceof Error ? error.message : String(error),
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
