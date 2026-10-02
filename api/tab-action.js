import crypto from 'node:crypto';
import { noStore, readSession, supabase } from '../lib/tab-auth.js';

export default async function handler(req, res) {
  noStore(res);
  if (req.method !== 'POST') return res.status(405).json({ error: 'method_not_allowed' });

  try {
    const currentAgent = await readSession(req);
    const { action, conversation_uuid: conversationUuid } = req.body || {};

    if (!action) return res.status(400).json({ error: 'invalid_payload' });

    if (action === 'agent_status') {
      const status = String(req.body.status || '');
      if (!['online', 'away', 'offline'].includes(status)) {
        return res.status(400).json({ error: 'invalid_agent_status' });
      }
      const { error: updateError } = await supabase
        .from('agents')
        .update({
          status,
          last_seen_at: new Date().toISOString(),
          updated_at: new Date().toISOString(),
        })
        .eq('id', currentAgent.id);

      if (updateError) throw updateError;
      return res.json({ ok: true });
    }

    if (!conversationUuid) {
      return res.status(400).json({ error: 'missing_conversation_uuid' });
    }

    const { data: conversation, error } = await supabase
      .from('conversations')
      .select('*')
      .eq('conversation_uuid', conversationUuid)
      .maybeSingle();

    if (error) throw error;
    if (!conversation) return res.status(404).json({ error: 'conversation_not_found' });

    if (action === 'reply') {
      const message = String(req.body.message || '').trim();
      if (!message) return res.status(400).json({ error: 'empty_message' });

      const { data: site, error: siteError } = await supabase
        .from('sites')
        .select('*')
        .eq('id', conversation.site_id)
        .maybeSingle();

      if (siteError) throw siteError;
      if (!site?.shared_secret) throw new Error('site_secret_missing');

      const externalId = `tab:${crypto.randomUUID()}`;
      const response = await fetch(conversation.reply_url, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'X-TLC-Secret': site.shared_secret,
        },
        body: JSON.stringify({
          conversation_uuid: conversation.conversation_uuid,
          message,
          sender_name: currentAgent.display_name,
          external_id: externalId,
        }),
      });

      if (!response.ok) throw new Error(`wordpress_reply_${response.status}`);

      await supabase.from('messages').insert({
        conversation_id: conversation.id,
        sender_type: 'agent',
        sender_name: currentAgent.display_name,
        body: message,
        external_id: externalId,
      });

      await supabase
        .from('conversations')
        .update({
          assigned_agent_id: currentAgent.id,
          status: 'pending',
          last_message_at: new Date().toISOString(),
          updated_at: new Date().toISOString(),
        })
        .eq('id', conversation.id);

      return res.json({ ok: true });
    }

    if (action === 'assign') {
      const agentId = String(req.body.agent_id || '');
      if (!agentId) return res.status(400).json({ error: 'missing_agent_id' });

      const { error: updateError } = await supabase
        .from('conversations')
        .update({
          assigned_agent_id: agentId,
          status: 'open',
          updated_at: new Date().toISOString(),
        })
        .eq('id', conversation.id);

      if (updateError) throw updateError;
      return res.json({ ok: true });
    }

    if (action === 'status') {
      const status = String(req.body.status || '');
      if (!['open', 'pending', 'closed', 'missed'].includes(status)) {
        return res.status(400).json({ error: 'invalid_status' });
      }

      const { error: updateError } = await supabase
        .from('conversations')
        .update({
          status,
          updated_at: new Date().toISOString(),
        })
        .eq('id', conversation.id);

      if (updateError) throw updateError;
      return res.json({ ok: true });
    }

    return res.status(400).json({ error: 'unknown_action' });
  } catch (error) {
    console.error('[tab-action]', error);
    return res.status(500).json({ error: error instanceof Error ? error.message : String(error) });
  }
}
