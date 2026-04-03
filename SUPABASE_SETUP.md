# Supabase Setup (HRBD Dashboard)

## 1) Create table

Run in Supabase SQL Editor:

```sql
create table if not exists public.dashboard_state (
  id text primary key,
  payload jsonb not null,
  updated_at timestamptz not null default now()
);

create index if not exists dashboard_state_updated_at_idx
  on public.dashboard_state (updated_at desc);
```

## 2) Enable RLS + policies (simple team mode)

```sql
alter table public.dashboard_state enable row level security;

-- Allow read for anon/authenticated
drop policy if exists "dashboard_state_read" on public.dashboard_state;
create policy "dashboard_state_read"
on public.dashboard_state
for select
to anon, authenticated
using (true);

-- Allow insert for anon/authenticated
drop policy if exists "dashboard_state_insert" on public.dashboard_state;
create policy "dashboard_state_insert"
on public.dashboard_state
for insert
to anon, authenticated
with check (true);

-- Allow update for anon/authenticated
drop policy if exists "dashboard_state_update" on public.dashboard_state;
create policy "dashboard_state_update"
on public.dashboard_state
for update
to anon, authenticated
using (true)
with check (true);
```

> Note: this is the fastest setup for internal sharing. If needed, we can harden this later with login + user-based policies.

## 3) Collect credentials

From Supabase project settings:
- Project URL
- anon public key

## 4) Configure in Dashboard

On page **Trang Import** -> block **Supabase Cloud Sync**:
1. Fill `Supabase URL`
2. Fill `Supabase Anon Key`
3. Table = `dashboard_state`
4. State key = `global`
5. Click `Lưu cấu hình`
6. Click `Test kết nối`
7. Click `Đẩy lên cloud` (first publish)
8. Turn on `Auto sync`

## 5) Multi-user usage

- Share the dashboard URL with teammates.
- Each teammate inputs the same Supabase URL + anon key once.
- Click `Tải từ cloud` first time to load latest shared data.

## 6) Auto config for everyone (Vercel, no manual input)

If deployed on Vercel, app can auto-load Supabase config for all users.

In Vercel project -> **Settings** -> **Environment Variables**, add:

- `SUPABASE_URL` = your project URL
- `SUPABASE_ANON_KEY` = your anon public key
- `SUPABASE_TABLE` = `dashboard_state`
- `SUPABASE_STATE_KEY` = `global`
- `SUPABASE_AUTO_SYNC` = `true`

Then redeploy the project. After that, teammates open the dashboard link and cloud sync fields are prefilled automatically.
