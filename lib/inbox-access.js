import { realAgentName } from './agent-registration.js';
export function assignedConversations(db, agent) {
  const query = db.from('conversations').select('*');
  return agent.is_admin === true ? query : query.eq('assigned_agent_id', agent.id);
}
export async function transferAgents(db, siteId, currentAgent) {
  if (currentAgent.is_admin === true) {
    const { data, error } = await db.from('agents').select('id,display_name,status').eq('is_active', true).order('display_name');
    if (error) throw error;
    return (data || []).filter(a => realAgentName(a.display_name));
  }
  const { data: mappings, error } = await db.from('site_agents').select('agent_id').eq('site_id', siteId).eq('is_active', true);
  if (error) throw error;
  const ids = [...new Set([currentAgent.id, ...(mappings || []).map(m => m.agent_id)])];
  const result = await db.from('agents').select('id,display_name,status').eq('is_active', true).in('id', ids).order('display_name');
  if (result.error) throw result.error;
  return (result.data || []).filter(a => realAgentName(a.display_name));
}
