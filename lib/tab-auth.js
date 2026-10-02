import crypto from 'node:crypto';
import { createClient } from '@supabase/supabase-js';
import { createRemoteJWKSet, decodeJwt, jwtVerify } from 'jose';

export const supabase = createClient(
  process.env.SUPABASE_URL || 'https://invalid.local',
  process.env.SUPABASE_SERVICE_ROLE_KEY || 'invalid',
  { auth: { persistSession: false, autoRefreshToken: false } },
);

const tenantId = process.env.TENANT_ID || '';
const clientId = process.env.CLIENT_ID || '';
const sessionSecret = process.env.TAB_SESSION_SECRET || process.env.CLIENT_SECRET || '';
const JWKS = createRemoteJWKSet(new URL(`https://login.microsoftonline.com/${tenantId}/discovery/v2.0/keys`));

function b64url(input) {
  return Buffer.from(input).toString('base64url');
}

function sign(payload) {
  return crypto.createHmac('sha256', sessionSecret).update(payload).digest('base64url');
}

export function createSession(agent) {
  const payload = JSON.stringify({
    agent_id: agent.id,
    aad_object_id: agent.aad_object_id,
    display_name: agent.display_name,
    exp: Math.floor(Date.now() / 1000) + 60 * 60 * 8,
  });
  const encoded = b64url(payload);
  return `${encoded}.${sign(encoded)}`;
}

export async function readSession(req) {
  if (!sessionSecret) throw new Error('missing_session_secret');
  const token = String(req.headers['x-tlc-session'] || '');
  const [encoded, signature] = token.split('.');
  if (!encoded || !signature) throw new Error('missing_session');
  const expected = sign(encoded);
  const a = Buffer.from(signature);
  const b = Buffer.from(expected);
  if (a.length !== b.length || !crypto.timingSafeEqual(a, b)) throw new Error('invalid_session');
  const payload = JSON.parse(Buffer.from(encoded, 'base64url').toString('utf8'));
  if (!payload.exp || payload.exp < Math.floor(Date.now() / 1000)) throw new Error('expired_session');
  const { data: agent, error } = await supabase.from('agents').select('*').eq('id', payload.agent_id).eq('is_active', true).maybeSingle();
  if (error || !agent) throw new Error('agent_not_found');
  return agent;
}

export async function verifyTeamsSso(token) {
  if (!tenantId || !clientId) throw new Error('missing_teams_configuration');
  const decoded = decodeJwt(token);
  if (decoded.tid !== tenantId) throw new Error('wrong_tenant');
  const resourceAudience = `api://teams-live-chat-bridge.vercel.app/${clientId}`;\n  const { payload } = await jwtVerify(token, JWKS, { audience: [clientId, resourceAudience] });
  const oid = payload.oid || payload.sub;
  if (!oid) throw new Error('missing_object_id');
  return { oid: String(oid), name: String(payload.name || payload.preferred_username || 'Agent') };
}

export function noStore(res) {
  res.setHeader('Cache-Control', 'no-store, max-age=0');
  res.setHeader('Pragma', 'no-cache');
}
