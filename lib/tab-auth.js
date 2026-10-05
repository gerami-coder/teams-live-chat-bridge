import crypto from 'node:crypto';
import { createClient } from '@supabase/supabase-js';

export const supabase = createClient(
  process.env.SUPABASE_URL || 'https://invalid.local',
  process.env.SUPABASE_SERVICE_ROLE_KEY || 'invalid',
  { auth: { persistSession: false, autoRefreshToken: false } },
);

const tenantId = process.env.TENANT_ID || '';
const clientId = process.env.CLIENT_ID || '';
const sessionSecret = process.env.TAB_SESSION_SECRET || process.env.CLIENT_SECRET || '';

let jwksCache = { keys: [], expiresAt: 0 };

function b64url(input) {
  return Buffer.from(input).toString('base64url');
}

function sign(payload) {
  return crypto.createHmac('sha256', sessionSecret).update(payload).digest('base64url');
}

function decodePart(part) {
  return JSON.parse(Buffer.from(part, 'base64url').toString('utf8'));
}

async function getJwks() {
  if (Date.now() < jwksCache.expiresAt && jwksCache.keys.length) return jwksCache.keys;

  const url = `https://login.microsoftonline.com/${tenantId}/discovery/v2.0/keys`;
  const response = await fetch(url, { headers: { Accept: 'application/json' } });
  if (!response.ok) throw new Error(`jwks_fetch_${response.status}`);
  const data = await response.json();
  const keys = Array.isArray(data.keys) ? data.keys : [];
  if (!keys.length) throw new Error('jwks_empty');

  jwksCache = { keys, expiresAt: Date.now() + 60 * 60 * 1000 };
  return keys;
}

export function createSession(agent) {
  if (!sessionSecret) throw new Error('missing_session_secret');
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

  const { data: agent, error } = await supabase
    .from('agents')
    .select('*')
    .eq('id', payload.agent_id)
    .eq('is_active', true)
    .maybeSingle();

  if (error || !agent) throw new Error('agent_not_found');
  return agent;
}

export async function verifyTeamsSso(token) {
  if (!tenantId || !clientId) throw new Error('missing_teams_configuration');

  const parts = String(token || '').split('.');
  if (parts.length !== 3) throw new Error('invalid_token_format');

  const [encodedHeader, encodedPayload, encodedSignature] = parts;
  const header = decodePart(encodedHeader);
  const payload = decodePart(encodedPayload);

  if (header.alg !== 'RS256' || !header.kid) throw new Error('unsupported_token');
  if (payload.tid !== tenantId) throw new Error('wrong_tenant');

  const expectedAudiences = [
    clientId,
    `api://teams-live-chat-bridge.vercel.app/${clientId}`,
  ];
  if (!expectedAudiences.includes(String(payload.aud || ''))) throw new Error('wrong_audience');

  const now = Math.floor(Date.now() / 1000);
  if (payload.exp && Number(payload.exp) < now) throw new Error('token_expired');
  if (payload.nbf && Number(payload.nbf) > now + 60) throw new Error('token_not_yet_valid');

  const keys = await getJwks();
  const jwk = keys.find((key) => key.kid === header.kid);
  if (!jwk) throw new Error('signing_key_not_found');

  const publicKey = crypto.createPublicKey({ key: jwk, format: 'jwk' });
  const signed = Buffer.from(`${encodedHeader}.${encodedPayload}`);
  const signature = Buffer.from(encodedSignature, 'base64url');
  const valid = crypto.verify('RSA-SHA256', signed, publicKey, signature);
  if (!valid) throw new Error('invalid_token_signature');

  const oid = payload.oid || payload.sub;
  if (!oid) throw new Error('missing_object_id');

  return {
    oid: String(oid),
    name: String(payload.name || payload.preferred_username || 'Agent'),
  };
}

export function noStore(res) {
  res.setHeader('Cache-Control', 'no-store, max-age=0');
  res.setHeader('Pragma', 'no-cache');
}
