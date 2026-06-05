-- ============================================================
--  Narciso Designs — Furniture & Procurement Tracker
--  Supabase schema  (paste into Supabase → SQL Editor → Run)
-- ============================================================

-- Key/value store: one row per (user, key).
-- This mirrors the app's existing storage model, so the React
-- code barely changes — every project's items, discounts,
-- contractors, payments, presentations, etc. live here as JSON,
-- scoped to the signed-in designer.
create table if not exists public.kv (
  user_id    uuid        not null references auth.users (id) on delete cascade,
  k          text        not null,
  v          jsonb,
  updated_at timestamptz not null default now(),
  primary key (user_id, k)
);

-- Row Level Security: a user can only ever touch their own rows.
alter table public.kv enable row level security;

drop policy if exists "kv select own" on public.kv;
drop policy if exists "kv insert own" on public.kv;
drop policy if exists "kv update own" on public.kv;
drop policy if exists "kv delete own" on public.kv;

create policy "kv select own" on public.kv
  for select using (auth.uid() = user_id);
create policy "kv insert own" on public.kv
  for insert with check (auth.uid() = user_id);
create policy "kv update own" on public.kv
  for update using (auth.uid() = user_id) with check (auth.uid() = user_id);
create policy "kv delete own" on public.kv
  for delete using (auth.uid() = user_id);

-- ============================================================
--  PHASE 2 (client portal) — not required for the first deploy.
--  When you're ready for external client logins, a serverless
--  function will use the SERVICE ROLE key to read/write the
--  client-visible slice of a project, validated by access code.
--  Table below stores which project an access code maps to.
-- ============================================================
-- create table if not exists public.client_access (
--   code        text primary key,
--   owner_id    uuid not null references auth.users(id) on delete cascade,
--   project_id  text not null,
--   created_at  timestamptz not null default now()
-- );
-- (No RLS policies = locked to service role only. The /api/portal
--  function reads it server-side and never exposes the DB to the client.)
