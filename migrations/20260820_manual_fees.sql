alter table public.monthly_settlement
  add column if not exists replacement_fee double precision default 0;

alter table public.monthly_settlement
  add column if not exists mechanism_fee_manual double precision default 0;
