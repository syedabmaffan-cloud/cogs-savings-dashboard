-- ============================================================================
-- COGS Dashboard — shared Material Planning ETA register (Supabase / Postgres)
-- Run this once in Supabase: SQL Editor -> New query -> paste -> Run.
--
-- Model: anyone with the dashboard link can read and write ETA values.
-- The anon key shipped in the page is public by design; RLS below allows
-- select/insert/update and deliberately allows NO delete, so a bad edit can
-- always be recovered from mp_eta_history.
-- ============================================================================

create table if not exists public.mp_eta (
  sbu         text        not null,
  item_code   text        not null,
  eta1        date,
  eta_qty1    numeric,
  eta2        date,
  eta_qty2    numeric,
  eta3        date,
  eta_qty3    numeric,
  updated_by  text        not null default 'anonymous',
  updated_at  timestamptz not null default now(),
  primary key (sbu, item_code)
);

create table if not exists public.mp_eta_history (
  id          bigint generated always as identity primary key,
  sbu         text,
  item_code   text,
  eta1        date,
  eta_qty1    numeric,
  eta2        date,
  eta_qty2    numeric,
  eta3        date,
  eta_qty3    numeric,
  updated_by  text,
  updated_at  timestamptz not null default now(),
  action      text
);

create index if not exists mp_eta_history_key_idx on public.mp_eta_history (sbu, item_code, updated_at desc);

-- keep updated_at authoritative (server clock), even if a client sends one
create or replace function public.mp_eta_touch()
returns trigger
language plpgsql
as $$
begin
  new.updated_at := now();
  return new;
end;
$$;

drop trigger if exists mp_eta_touch_trg on public.mp_eta;
create trigger mp_eta_touch_trg
before insert or update on public.mp_eta
for each row execute function public.mp_eta_touch();

-- append-only audit trail
create or replace function public.mp_eta_log()
returns trigger
language plpgsql
security definer
set search_path = public
as $$
begin
  insert into public.mp_eta_history
    (sbu, item_code, eta1, eta_qty1, eta2, eta_qty2, eta3, eta_qty3, updated_by, updated_at, action)
  values
    (new.sbu, new.item_code, new.eta1, new.eta_qty1, new.eta2, new.eta_qty2, new.eta3, new.eta_qty3, new.updated_by, now(), tg_op);
  return new;
end;
$$;

drop trigger if exists mp_eta_log_trg on public.mp_eta;
create trigger mp_eta_log_trg
after insert or update on public.mp_eta
for each row execute function public.mp_eta_log();

-- access control: anonymous, link-only editing (no deletes)
alter table public.mp_eta enable row level security;
alter table public.mp_eta_history enable row level security;

drop policy if exists mp_eta_read   on public.mp_eta;
drop policy if exists mp_eta_insert on public.mp_eta;
drop policy if exists mp_eta_update on public.mp_eta;
drop policy if exists mp_eta_hist_read on public.mp_eta_history;

create policy mp_eta_read   on public.mp_eta for select using (true);
create policy mp_eta_insert on public.mp_eta for insert with check (true);
create policy mp_eta_update on public.mp_eta for update using (true) with check (true);
create policy mp_eta_hist_read on public.mp_eta_history for select using (true);

grant select, insert, update on public.mp_eta to anon, authenticated;
grant select on public.mp_eta_history to anon, authenticated;

-- ============================================================================
-- Recovery helpers (run manually when someone wipes or mis-edits values)
--
-- Latest value per item:
--   select distinct on (sbu, item_code) *
--   from public.mp_eta_history
--   order by sbu, item_code, updated_at desc;
--
-- Undo one item to its previous value:
--   update public.mp_eta e
--      set eta1 = h.eta1, eta_qty1 = h.eta_qty1,
--          eta2 = h.eta2, eta_qty2 = h.eta_qty2,
--          eta3 = h.eta3, eta_qty3 = h.eta_qty3
--     from (select * from public.mp_eta_history
--            where sbu = 'AAFL' and item_code = '11113855'
--            order by updated_at desc offset 1 limit 1) h
--    where e.sbu = 'AAFL' and e.item_code = '11113855';
-- ============================================================================
