import { test } from 'node:test';
import assert from 'node:assert/strict';
import { registerAgent, realAgentName } from '../lib/agent-registration.js';

function database(rows = []) {
  const writes = [];
  return { rows, writes, from(table) {
    assert.equal(table, 'agents');
    let filters = [], value, update = false;
    const q = {
      select() { return q; },
      eq(k, v) { filters.push([k,v]); return q; },
      update(v) { value = v; update = true; return q; },
      insert(v) { value = v; return q; },
      async maybeSingle() { return { data: rows.find(r => filters.every(([k,v]) => r[k] === v)) || null }; },
      async single() {
        writes.push(value);
        if (update) {
          const row = rows.find(r => filters.every(([k,v]) => r[k] === v));
          Object.assign(row, value); return { data: row };
        }
        const row = { id: 'new-id', ...value }; rows.push(row); return { data: row };
      },
    }; return q;
  }};
}
const activity = (name, extra = {}) => ({ type: 'message', from: { aadObjectId: 'person-123', name }, conversation: { id: 'a:personal', conversationType: 'personal' }, ...extra });
const register = (db, a) => registerAgent(db, a, s => s.toLowerCase().replaceAll(' ', '-'));

test('installation with a person ID but no name creates no placeholder', async () => {
  const db = database(); assert.equal(await register(db, activity(undefined, { type: 'installationUpdate' })), null); assert.equal(db.writes.length, 0);
});
test('missing identity, generic name, bot and group activities cannot create agents', async () => {
  for (const a of [activity('Support Agent'), activity('Agent'), activity('  '), activity('Real Name',{from:{name:'Real Name'}}), activity('Real Name',{from:{role:'bot',name:'Real Name'}}), activity('Real Name',{conversation:{id:'group',conversationType:'groupChat'}})]) {
    const db=database(); assert.equal(await register(db,a),null); assert.equal(db.writes.length,0);
  }
});
test('named personal message registers once and retains identity on subsequent messages', async () => {
  const db=database(); const first=await register(db,activity('Real Person')); assert.equal(first.status,'online');
  const next=await register(db,activity('Real Person')); assert.equal(next.id,first.id); assert.equal(db.rows.length,1);
});
test('name-less install cannot overwrite a real name or turn away into online', async () => {
  const db=database([{id:'existing',aad_object_id:'person-123',display_name:'Real Person',status:'away',teams_conversation_id:'a:old'}]);
  const row=await register(db,activity(undefined,{type:'installationUpdate'})); assert.equal(row.display_name,'Real Person'); assert.equal(row.status,'away'); assert.equal(db.writes[0].last_seen_at,undefined);
});
test('existing real agent can reply without a name and offline preference survives', async () => {
  const db=database([{id:'existing',aad_object_id:'person-123',display_name:'Real Person',status:'offline'}]);
  const row=await register(db,activity(undefined)); assert.equal(row.id,'existing'); assert.equal(row.status,'offline'); assert.equal(row.display_name,'Real Person');
});
test('a real name repairs a placeholder in place without replacing assignments', async () => {
  const db=database([{id:'existing',aad_object_id:'person-123',display_name:'Support Agent',status:'online'}]);
  const row=await register(db,activity(' Actual   Person ')); assert.equal(row.id,'existing'); assert.equal(row.display_name,'Actual Person'); assert.equal(db.rows.length,1);
});
test('named install does not claim the user is online', async () => {
  const row=await register(database(),activity('Real Person',{type:'installationUpdate'})); assert.equal(row.status,'offline');
});
test('conversation fallback cannot rebind a different Microsoft identity', async () => {
  const db=database([{id:'existing',aad_object_id:'other-person',teams_conversation_id:'a:personal',display_name:'Other Person'}]); assert.equal(await register(db,activity('Real Person')),null); assert.equal(db.writes.length,0);
});
test('database lookup failures stop registration instead of inserting duplicates', async () => {
  const db={from:()=>({select(){return this},eq(){return this},async maybeSingle(){return {error:new Error('lookup failed')}}})};
  await assert.rejects(register(db,activity('Real Person')),/lookup failed/);
});
test('real names include non-Latin names while placeholder labels are rejected', () => {
  assert.equal(realAgentName('ሰላም'),'ሰላም'); assert.equal(realAgentName(' Support Agent '),'');
});
