-- 在 Supabase 的 SQL Editor 中完整执行一次。
create table if not exists public.fortress_state (
  id text primary key,
  members jsonb not null default '[]'::jsonb,
  queues jsonb not null default '{}'::jsonb,
  last_sweep text not null default '',
  contest boolean not null default false,
  updated_at timestamptz not null default now()
);

alter table public.fortress_state enable row level security;
grant select, insert, update on public.fortress_state to anon, authenticated;

drop policy if exists "shared state can be read" on public.fortress_state;
drop policy if exists "shared state can be created" on public.fortress_state;
drop policy if exists "shared state can be updated" on public.fortress_state;

create policy "shared state can be read"
  on public.fortress_state for select to anon, authenticated
  using (id = 'main');

create policy "shared state can be created"
  on public.fortress_state for insert to anon, authenticated
  with check (id = 'main');

create policy "shared state can be updated"
  on public.fortress_state for update to anon, authenticated
  using (id = 'main')
  with check (id = 'main');
