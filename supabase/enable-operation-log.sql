-- 已啟用單一管理員模式的專案：新增操作紀錄。可重複執行。
begin;

create table if not exists public.operation_log (
  id bigint generated always as identity primary key,
  occurred_at timestamptz not null default clock_timestamp(),
  actor_id uuid,
  entity text not null,
  operation text not null check (operation in ('INSERT', 'UPDATE', 'DELETE', 'RESTORE')),
  target text not null,
  before_data jsonb,
  after_data jsonb
);

alter table public.operation_log enable row level security;
revoke all on public.operation_log from public, anon, authenticated;
grant select on public.operation_log to authenticated;
create index if not exists operation_log_timeline_idx on public.operation_log (occurred_at desc, id desc);

drop policy if exists operation_log_owner_read on public.operation_log;
create policy operation_log_owner_read on public.operation_log for select to authenticated
  using ((select auth.uid()) = '5568706d-caaa-459b-b842-13cdbc3929fc'::uuid);

-- 私有表只在還原交易進行時標記，避免逐筆還原內容淹沒時間表。
create table if not exists private.operation_log_restore (
  transaction_id bigint primary key
);
revoke all on private.operation_log_restore from public, anon, authenticated;

create or replace function private.record_operation_log()
returns trigger language plpgsql security definer set search_path = ''
as $$
declare
  v_before jsonb;
  v_after jsonb;
  v_row jsonb;
  v_target text;
begin
  if exists (select 1 from private.operation_log_restore where transaction_id = txid_current()) then
    return null;
  end if;

  if tg_op <> 'INSERT' then v_before := to_jsonb(old); end if;
  if tg_op <> 'DELETE' then v_after := to_jsonb(new); end if;
  v_row := coalesce(v_after, v_before);
  if v_row->>'user_id' is distinct from '5568706d-caaa-459b-b842-13cdbc3929fc' then return null; end if;

  if tg_table_name = 'stock_movements' then
    select code into v_target from public.products where id = (v_row->>'product_id')::uuid;
  end if;
  v_target := coalesce(v_target, v_row->>'code', v_row->>'product_code', v_row->>'value', v_row->>'product_id', v_row->>'id');

  v_before := v_before - array['id', 'user_id', 'created_at', 'updated_at', 'request_key', 'product_id', 'total_minutes'];
  v_after := v_after - array['id', 'user_id', 'created_at', 'updated_at', 'request_key', 'product_id', 'total_minutes'];
  if v_before is not distinct from v_after then return null; end if;

  insert into public.operation_log (actor_id, entity, operation, target, before_data, after_data)
  values (auth.uid(), tg_table_name, tg_op, v_target, v_before, v_after);
  return null;
end;
$$;
revoke all on function private.record_operation_log() from public, anon, authenticated;

drop trigger if exists record_operation_log on public.products;
create trigger record_operation_log after insert or update or delete on public.products
for each row execute function private.record_operation_log();
drop trigger if exists record_operation_log on public.stock_movements;
create trigger record_operation_log after insert or update or delete on public.stock_movements
for each row execute function private.record_operation_log();
drop trigger if exists record_operation_log on public.production_records;
create trigger record_operation_log after insert or update or delete on public.production_records
for each row execute function private.record_operation_log();
drop trigger if exists record_operation_log on public.production_options;
create trigger record_operation_log after insert or update or delete on public.production_options
for each row execute function private.record_operation_log();

create or replace function public.admin_import_all_with_operation_log(p_payload jsonb)
returns jsonb language plpgsql security definer set search_path = ''
as $$
declare
  v_user_id uuid := private.require_app_user();
  v_transaction bigint := txid_current();
  v_result jsonb;
begin
  insert into private.operation_log_restore values (v_transaction);
  v_result := public.admin_import_all(p_payload);
  delete from private.operation_log_restore where transaction_id = v_transaction;
  insert into public.operation_log (actor_id, entity, operation, target, after_data)
  values (v_user_id, 'admin_data', 'RESTORE', 'Excel', v_result->'counts');
  return v_result;
end;
$$;
revoke all on function public.admin_import_all_with_operation_log(jsonb) from public, anon, authenticated;
grant execute on function public.admin_import_all_with_operation_log(jsonb) to authenticated;

commit;
