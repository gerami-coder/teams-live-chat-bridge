import { noStore, readSession, supabase } from '../lib/tab-auth.js';

export default async function handler(req, res) {
  noStore(res);
  if (req.method !== 'GET') return res.status(405).json({ error: 'method_not_allowed' });

  try {
    const agent = await readSession(req);
    const conversationUuid = String(req.query.conversation || '');

    const [{ data: sites, error: sitesError }, { data: agents, error: agentsError }] = await Promise.all([
      supabase.from('sites').select('id,site_key,name,base_url,is_active').eq('is_active', true).order('name'),
      supabase.from('agents').select('id,agent_key,display_name,status,is_active,last_seen_at').eq('is_active', true).order('display_name'),
    ]);
    if (sitesError) throw sitesError;
    if (agentsError) throw agentsError;

    if (conversationUuid) {
      const { data: conversation, error } = await supabase
        .from('conversations')
        .select('*')
        .eq('conversation_uuid', conversationUuid)
        .maybeSingle();

      if (error) throw error;
      if (!conversation) return res.status(404).json({ error: 'conversation_not_found' });

      const { data: messages, error: messagesError } = await supabase
        .from('messages')
        .select('*')
        .eq('conversation_id', conversation.id)
        .order('created_at', { ascending: true });

      if (messagesError) throw messagesError;

      return res.json({
        ok: true,
        agent: publicAgent(agent),
        sites: sites || [],
        agents: agents || [],
        conversation,
        messages: messages || [],
      });
    }

    const { data: conversations, error } = await supabase
      .from('conversations')
      .select('*')
      .order('last_message_at', { ascending: false })
      .limit(150);

    if (error) throw error;

    const ids = (conversations || []).map((c) => c.id);
    const latestMessages = [];

    if (ids.length) {
      const { data: msgs, error: msgError } = await supabase
        .from('messages')
        .select('id,conversation_id,sender_type,sender_name,body,created_at')
        .in('conversation_id', ids)
        .order('created_at', { ascending: false });

      if (msgError) throw msgError;

      const seen = new Set();
      for (const msg of msgs || []) {
        if (!seen.has(msg.conversation_id)) {
          latestMessages.push(msg);
          seen.add(msg.conversation_id);
        }
      }
    }

    return res.json({
      ok: true,
      agent: publicAgent(agent),
      sites: sites || [],
      agents: agents || [],
      conversations: conversations || [],
      latest_messages: latestMessages,
    });
  } catch (error) {
    console.error('[tab-data]', error);
    return res.status(401).json({ error: error instanceof Error ? error.message : String(error) });
  }
}

function publicAgent(agent) {
  return {
    id: agent.id,
    agent_key: agent.agent_key,
    display_name: agent.display_name,
    status: agent.status,
  };
}
