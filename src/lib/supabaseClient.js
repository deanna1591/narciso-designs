import { createClient } from "@supabase/supabase-js";

// These come from Supabase → Project Settings → API.
// They are PUBLIC (anon) values — safe to ship in the front end.
// In Vite, env vars must be prefixed with VITE_ to be exposed.
const url = import.meta.env.VITE_SUPABASE_URL;
const anonKey = import.meta.env.VITE_SUPABASE_ANON_KEY;

if (!url || !anonKey) {
  console.error("Missing VITE_SUPABASE_URL / VITE_SUPABASE_ANON_KEY env vars.");
}

export const supabase = createClient(url, anonKey, {
  auth: { persistSession: true, autoRefreshToken: true },
});
