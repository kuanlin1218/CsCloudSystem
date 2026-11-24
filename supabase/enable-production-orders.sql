-- 在現有單一管理員、產品編號編輯及條碼備份設定之後執行。
begin;

alter table public.production_options
  add column if not exists order_quantity integer,
  add column if not exists due_date date,
  add column if not exists completed_on date;

alter table public.production_options drop constraint if exists production_options_order_settings_check;
alter table public.production_options add constraint production_options_order_settings_check check (
  (order_quantity is null or order_quantity > 0)
  and (option_type = 'product' or (order_quantity is null and due_date is null and completed_on is null))
);

create or replace function public.update_production_order(
  p_current_value text, p_new_value text, p_order_quantity integer, p_due_date date
)
returns void
language plpgsql
security definer
set search_path = ''
as $$
declare
  v_user_id uuid := private.require_app_user();
  v_current text := upper(btrim(coalesce(p_current_value, '')));
  v_next text := upper(btrim(coalesce(p_new_value, '')));
begin
  if v_current = '' or v_next = '' or (p_order_quantity is not null and p_order_quantity <= 0) then
    raise exception '訂單編號或數量不正確' using errcode = '22023';
  end if;

  if v_current <> v_next then
    perform public.rename_production_product_code(v_current, v_next);
  end if;

  update public.production_options
    set order_quantity = p_order_quantity, due_date = p_due_date
    where user_id = v_user_id and option_type = 'product' and upper(btrim(value)) = v_next;
  if not found then raise exception '找不到原本的訂單' using errcode = 'P0002'; end if;
end;
$$;

create or replace function public.save_completed_production_record(
  p_input jsonb, p_completed_on date, p_expected jsonb default null
)
returns void
language plpgsql
security definer
set search_path = ''
as $$
declare
  v_user_id uuid := private.require_app_user();
  v_input public.production_records := jsonb_populate_record(null::public.production_records, p_input);
  v_expected public.production_records := jsonb_populate_record(null::public.production_records, p_expected);
  v_product text := upper(btrim(coalesce(v_input.product_code, '')));
  v_machine text := upper(btrim(coalesce(v_input.machine, '')));
begin
  if p_completed_on is null or p_completed_on > (now() at time zone 'Asia/Taipei')::date
     or v_input.work_date is null or v_product = '' or v_machine = ''
     or coalesce(v_input.hours, -1) < 0 or coalesce(v_input.minutes, -1) not between 0 and 59
     or v_input.hours::bigint * 60 + v_input.minutes not between 1 and 2147483647
     or coalesce(v_input.quantity, 0) <= 0 then
    raise exception '加工日期、工時、數量或完工日期不正確' using errcode = '22023';
  end if;

  -- 先更新訂單以鎖定該列；下方加工紀錄若寫入失敗，完工日期也會一併回復。
  update public.production_options set completed_on = p_completed_on
    where user_id = v_user_id and option_type = 'product' and upper(btrim(value)) = v_product;
  if not found then raise exception '找不到要登記完工的訂單' using errcode = 'P0002'; end if;

  if p_expected is null then
    insert into public.production_records (user_id, work_date, product_code, machine, hours, minutes, quantity, notes)
    values (v_user_id, v_input.work_date, v_product, v_machine, v_input.hours, v_input.minutes, v_input.quantity, btrim(coalesce(v_input.notes, '')));
  else
    update public.production_records as r
      set hours = v_input.hours, minutes = v_input.minutes, quantity = v_input.quantity, notes = btrim(coalesce(v_input.notes, ''))
      where r.user_id = v_user_id and r.id = v_expected.id
        and r.work_date = v_input.work_date and upper(btrim(r.product_code)) = v_product and upper(btrim(r.machine)) = v_machine
        and r.work_date = v_expected.work_date and r.product_code = v_expected.product_code and r.machine = v_expected.machine
        and r.hours = v_expected.hours and r.minutes = v_expected.minutes
        and r.quantity = v_expected.quantity and r.notes is not distinct from v_expected.notes;
    if not found then raise exception '原登記已變更或刪除，請取消後重新送出。' using errcode = 'P0002'; end if;
  end if;
end;
$$;

-- 沿用原本的完整備份、還原及驗證；保留公開 RPC 名稱供既有前端和操作紀錄呼叫。
do $$
begin
  if to_regprocedure('private.admin_export_all_before_orders()') is null then
    alter function public.admin_export_all() set schema private;
    alter function private.admin_export_all() rename to admin_export_all_before_orders;
  end if;
  if to_regprocedure('private.admin_import_all_before_orders(jsonb)') is null then
    alter function public.admin_import_all(jsonb) set schema private;
    alter function private.admin_import_all(jsonb) rename to admin_import_all_before_orders;
  end if;
end;
$$;

create or replace function public.admin_export_all()
returns jsonb
language plpgsql
security definer
set search_path = ''
as $$
declare
  v_user_id uuid := private.require_app_user();
  v_export jsonb := private.admin_export_all_before_orders();
  v_options jsonb;
begin
  select coalesce(jsonb_agg(jsonb_build_object(
    'option_type', o.option_type, 'value', o.value,
    'order_quantity', o.order_quantity, 'due_date', o.due_date, 'completed_on', o.completed_on
  ) order by o.option_type, o.value), '[]'::jsonb)
  into v_options from public.production_options o where o.user_id = v_user_id;
  return jsonb_set(v_export, '{data,production_options}', v_options);
end;
$$;

create or replace function public.admin_import_all(p_payload jsonb)
returns jsonb
language plpgsql
security definer
set search_path = ''
as $$
declare
  v_user_id uuid := private.require_app_user();
  v_result jsonb := private.admin_import_all_before_orders(p_payload);
begin
  -- 舊備份沒有訂單欄位時保留空值；與原還原流程共用交易及操作紀錄範圍。
  update public.production_options o
    set order_quantity = restored.order_quantity, due_date = restored.due_date, completed_on = restored.completed_on
    from jsonb_to_recordset(p_payload->'production_options') as restored(
      option_type text, value text, order_quantity integer, due_date date, completed_on date
    )
    where o.user_id = v_user_id and o.option_type = 'product' and restored.option_type = 'product'
      and o.value = upper(btrim(restored.value))
      and (restored.order_quantity is not null or restored.due_date is not null or restored.completed_on is not null);
  return v_result;
end;
$$;

revoke all on function private.admin_export_all_before_orders() from public, anon, authenticated;
revoke all on function private.admin_import_all_before_orders(jsonb) from public, anon, authenticated;
revoke all on function public.update_production_order(text, text, integer, date),
  public.save_completed_production_record(jsonb, date, jsonb), public.admin_export_all(), public.admin_import_all(jsonb)
  from public, anon, authenticated;
grant execute on function public.update_production_order(text, text, integer, date),
  public.save_completed_production_record(jsonb, date, jsonb), public.admin_export_all(), public.admin_import_all(jsonb)
  to authenticated;

commit;
