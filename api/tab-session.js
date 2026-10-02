import { createSession, noStore, supabase, verifyTeamsSso } from '../lib/tab-auth.js';

export default async function handler(req, res) {
  noStore(res);
  if (req.method !== 'POST') return res.status(405).json({ error: 'method_not_allowed' });
  try {
    const auth = String(req.headers.authorization || '');
    const token = auth.startsWith('Bearer ') ? auth.slice(7) : '';
    if (!token) return res.status(401).json({ error: 'missing_teams_token' });

    const identity = await verifyTeamsSso(token);
    const { data: agent, error } = await supabase
      .from('agents')
      .select('*')
      .eq('aad_object_id', identity.oid)
      .eq('is_active', true)
      .maybeSingle();

    if (error) throw error;
    if (!agent) {
      return res.status(403).json({
        error: 'agent_not_registered',
        message: 'Open the Website Live Chat bot and send hello once, then reopen Inbox.',
      });
    }

    await supabase.from('agents').update({
      last_seen_at: new Date().toISOString(),
      status: agent.status === 'offline' ? 'offline' : 'online',
    }).eq('id', agent.id);

    return res.json({
      ok: true,
      session: createSession(agent),
      agent: {
        id: agent.id,
        agent_key: agent.agent_key,
        display_name: agent.display_name,
        status: agent.status,
      },
    });
  } catch (error) {
    console.error('[tab-session]', error);
    return res.status(401).json({ error: error instanceof Error ? error.message : String(error) });
  }
}
