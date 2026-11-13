// Never manufacture a person from a name-less Teams lifecycle event.
export function realAgentName(value) {
  const name = typeof value === 'string' ? value.trim().replace(/\s+/g, ' ') : '';
  return name && !/^(support(?: agent)?|agent|unknown|undefined|null)$/i.test(name) ? name : '';
}

export async function registerAgent(supabase, activity, slugify) {
  const conversationId = activity.conversation?.id;
  const type = activity.conversation?.conversationType;
  const personal = type ? type === 'personal' : String(conversationId || '').startsWith('a:');
  if (!conversationId || !personal || activity.from?.role === 'bot') return null;

  const aadObjectId = activity.from?.aadObjectId || null;
  const name = realAgentName(activity.from?.name);
  let existing = null;
  if (aadObjectId) {
    const result = await supabase.from('agents').select('*').eq('aad_object_id', aadObjectId).maybeSingle();
    if (result.error) throw result.error;
    existing = result.data;
  }
  if (!existing) {
    const result = await supabase.from('agents').select('*').eq('teams_conversation_id', conversationId).maybeSingle();
    if (result.error) throw result.error;
    existing = result.data;
    // A conversation match must never rebind a different directory identity.
    if (existing?.aad_object_id && aadObjectId && existing.aad_object_id !== aadObjectId) return null;
  }

  if (existing) {
    const displayName = name || realAgentName(existing.display_name);
    if (!displayName) return null;
    const patch = {
      display_name: displayName,
      aad_object_id: aadObjectId || existing.aad_object_id,
      teams_conversation_id: conversationId,
      updated_at: new Date().toISOString(),
    };
    // Install events do not prove availability. Only an actual user message does.
    if (activity.type === 'message') {
      patch.status = existing.status === 'offline' ? 'offline' : 'online';
      patch.last_seen_at = new Date().toISOString();
    }
    const { data, error } = await supabase.from('agents').update(patch).eq('id', existing.id).select('*').single();
    if (error) throw error;
    return data;
  }

  // A new agent needs a real name and a stable Microsoft user identity.
  if (!aadObjectId || !name) return null;
  const suffix = aadObjectId.replace(/[^a-zA-Z0-9]/g, '').slice(-6).toLowerCase();
  const { data, error } = await supabase.from('agents').insert({
    agent_key: `${slugify(name)}-${suffix}`,
    display_name: name,
    aad_object_id: aadObjectId,
    teams_conversation_id: conversationId,
    status: activity.type === 'message' ? 'online' : 'offline',
    ...(activity.type === 'message' ? { last_seen_at: new Date().toISOString() } : {}),
  }).select('*').single();
  if (error) throw error;
  return data;
}
