import { noStore, readSession, supabase } from '../lib/tab-auth.js';
import { assignedConversations, transferAgents } from '../lib/inbox-access.js';
export default async function handler(req, res) {
  noStore(res);
  if (req.method !== 'GET') return res.status(405).json({ error: 'method_not_allowed' });
  try {
    const agent = await readSession(req);
    const uuid = String(req.query.conversation || '');
    const publicAgent = { id: agent.id, display_name: agent.display_name, status: agent.status, is_admin: agent.is_admin === true, is_manager: agent.is_manager === true };
    if (uuid) {
      const { data: conversation, error } = await assignedConversations(supabase, agent).eq('conversation_uuid', uuid).maybeSingle();
      if (error) throw error;
      if (!conversation) return res.status(404).json({ error: 'conversation_not_found' });
      const [messages, site, agents] = await Promise.all([
        supabase.from('messages').select('*').eq('conversation_id', conversation.id).order('created_at', { ascending: true }),
        supabase.from('sites').select('id,name,base_url').eq('id', conversation.site_id).maybeSingle(),
        transferAgents(supabase, conversation.site_id, agent),
      ]);
      if (messages.error) throw messages.error;
      if (site.error) throw site.error;
      return res.json({ ok: true, agent: publicAgent, conversation, site: site.data, agents, messages: messages.data || [] });
    }
    const { data, error } = await assignedConversations(supabase, agent).order('last_message_at', { ascending: false }).limit(150);
    if (error) throw error;
    const conversations = data || [];
    const ids = conversations.map(c => c.id);
    const siteIds = [...new Set(conversations.map(c => c.site_id))];
    const sites = siteIds.length ? await supabase.from('sites').select('id,name,base_url').in('id', siteIds).order('name') : { data: [] };
    if (sites.error) throw sites.error;
    const messages = ids.length ? await supabase.from('messages').select('id,conversation_id,sender_type,sender_name,body,created_at').in('conversation_id', ids).order('created_at', { ascending: false }) : { data: [] };
    if (messages.error) throw messages.error;
    const seen = new Set();
    const latest = (messages.data || []).filter(m => { if (seen.has(m.conversation_id)) return false; seen.add(m.conversation_id); return true; });
    return res.json({ ok: true, agent: publicAgent, sites: sites.data || [], agents: [], conversations, latest_messages: latest });
  } catch (error) {
    console.error('[tab-data]', error);
    return res.status(401).json({ error: error instanceof Error ? error.message : String(error) });
  }
}
