import { noStore, readSession, supabase } from '../lib/tab-auth.js';
import { realAgentName } from '../lib/agent-registration.js';

export default async function handler(req, res) {
  noStore(res);
  if (!['GET','POST'].includes(req.method)) return res.status(405).json({error:'method_not_allowed'});
  try {
    const actor = await readSession(req);
    if (!actor.is_admin && !actor.is_manager) return res.status(403).json({error:'management_access_required'});
    if (req.method === 'GET') {
      const results = await Promise.all([
        supabase.from('sites').select('id,name,base_url,enforce_agent_assignments').eq('is_active',true).order('name'),
        supabase.from('agents').select('id,display_name,status').eq('is_active',true).order('display_name'),
        supabase.from('site_agents').select('site_id,agent_id').eq('is_active',true).eq('department','general'),
        actor.is_admin ? supabase.from('inbox_managers').select('agent_id') : Promise.resolve({data:[]}),
        actor.is_admin ? supabase.from('inbox_administrators').select('agent_id') : Promise.resolve({data:[]}),
      ]);
      for(const result of results) if(result.error) throw result.error;
      return res.json({ok:true,sites:results[0].data||[],agents:(results[1].data||[]).filter(a=>realAgentName(a.display_name)),assignments:results[2].data||[],managers:results[3].data||[],admins:results[4].data||[],is_admin:actor.is_admin===true});
    }
    const body=req.body||{};
    if (body.action === 'manager_role') {
      if(!actor.is_admin) return res.status(403).json({error:'admin_required'});
      if(typeof body.enabled !== 'boolean') return res.status(400).json({error:'invalid_role'});
      const {data:target,error} = await supabase.from('agents').select('id,display_name').eq('id',String(body.agent_id||'')).eq('is_active',true).maybeSingle();
      if(error) throw error;
      if(!target||!realAgentName(target.display_name)) return res.status(400).json({error:'invalid_agent'});
      const result=body.enabled ? await supabase.from('inbox_managers').upsert({agent_id:target.id},{onConflict:'agent_id',ignoreDuplicates:true}) : await supabase.from('inbox_managers').delete().eq('agent_id',target.id);
      if(result.error) throw result.error;
      return res.json({ok:true});
    }
    if(body.action!=='website_agents'||!Array.isArray(body.agent_ids)||!Array.isArray(body.expected_ids)) return res.status(400).json({error:'invalid_payload'});
    const {error} = await supabase.rpc('set_website_agents',{p_site_id:String(body.site_id||''),p_agent_ids:[...new Set(body.agent_ids)],p_expected_ids:[...new Set(body.expected_ids)]});
    if(error){if(error.code==='40001')return res.status(409).json({error:'assignments_changed',message:'Another manager changed this website. Reload assignments before saving.'});throw error;}
    return res.json({ok:true});
  } catch(error) {
    console.error('[tab-manage]',error);
    return res.status(400).json({error:'management_request_failed'});
  }
}
