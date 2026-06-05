import { supabase } from "./supabaseClient";

// Drop-in replacement for the sandbox `store` (window.storage).
// Same interface the app already uses everywhere:
//   await store.load(key)  -> parsed value or null
//   await store.save(key, value)
// Difference: data now lives in Supabase, scoped to the signed-in
// user (the designer), so it syncs across devices and is durable.

async function currentUserId() {
  const { data } = await supabase.auth.getUser();
  return data?.user?.id || null;
}

export const store = {
  async load(key) {
    try {
      const user_id = await currentUserId();
      if (!user_id) return null;
      const { data, error } = await supabase
        .from("kv")
        .select("v")
        .eq("user_id", user_id)
        .eq("k", key)
        .maybeSingle();
      if (error) { console.error("load failed", error); return null; }
      return data ? data.v : null;       // jsonb comes back already parsed
    } catch (e) {
      console.error("load failed", e);
      return null;
    }
  },

  async save(key, val) {
    try {
      const user_id = await currentUserId();
      if (!user_id) return;
      const { error } = await supabase
        .from("kv")
        .upsert(
          { user_id, k: key, v: val, updated_at: new Date().toISOString() },
          { onConflict: "user_id,k" }
        );
      if (error) console.error("save failed", error);
    } catch (e) {
      console.error("save failed", e);
    }
  },

  // optional helper if you ever need it
  async remove(key) {
    try {
      const user_id = await currentUserId();
      if (!user_id) return;
      await supabase.from("kv").delete().eq("user_id", user_id).eq("k", key);
    } catch (e) { console.error("remove failed", e); }
  },
};
