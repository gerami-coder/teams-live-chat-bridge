import { createClient } from '@supabase/supabase-js';

export default async function handler(_req, res) {
  try {
    const url = process.env.SUPABASE_URL;
    const key = process.env.SUPABASE_SERVICE_ROLE_KEY;

    if (!url || !key) {
      return res.status(500).json({
        ok: false,
        error: 'missing_supabase_environment_variables'
      });
    }

    const supabase = createClient(url, key, {
      auth: { persistSession: false, autoRefreshToken: false }
    });

    const { count, error } = await supabase
      .from('sites')
      .select('*', { count: 'exact', head: true });

    if (error) throw error;

    return res.status(200).json({
      ok: true,
      database: 'connected',
      sites: count ?? 0
    });
  } catch (error) {
    console.error('[health]', error);
    return res.status(500).json({
      ok: false,
      error: error instanceof Error ? error.message : String(error)
    });
  }
}
