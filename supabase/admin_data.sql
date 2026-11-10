-- 詮栩 CNC 雲端系統：管理員 Excel 備份／還原
-- 先執行 schema.sql，再在 Supabase SQL Editor 的「新查詢」執行本檔。
-- 注意：若已執行 secure_single_user.sql，請勿再執行本檔；安全版已改用 Supabase Auth 密碼。

create schema if not exists extensions;
create extension if not exists pgcrypto with schema extensions;
create schema if not exists private;
revoke all on schema private from public, anon, authenticated;

create table if not exists private.admin_settings (
  id smallint primary key default 1 check (id = 1),
  password_hash text not null,
  failed_attempts integer not null default 0,
  locked_until timestamptz,
  updated_at timestamptz not null default now()
);

revoke all on private.admin_settings from public, anon, authenticated;

create or replace function private.set_admin_password(p_password text)
returns void
language plpgsql
security definer
set search_path = ''
as $$
begin
  if coalesce(p_password, '') !~ '^[0-9]{8}$' then
    raise exception '管理員密碼必須是 8 位數字';
  end if;

  insert into private.admin_settings (id, password_hash, failed_attempts, locked_until, updated_at)
  values (1, extensions.crypt(p_password, extensions.gen_salt('bf', 12)), 0, null, now())
  on conflict (id) do update set
    password_hash = excluded.password_hash,
    failed_attempts = 0,
    locked_until = null,
    updated_at = now();
end;
$$;

revoke all on function private.set_admin_password(text) from public, anon, authenticated;

create or replace function private.check_admin_password(p_password text)
returns jsonb
language plpgsql
security definer
set search_path = ''
as $$
declare
  v_settings private.admin_settings%rowtype;
  v_failed integer;
begin
  select * into v_settings
  from private.admin_settings
  where id = 1
  for update;

  if not found then
    return jsonb_build_object('ok', false, 'error', '尚未設定管理員密碼，請先在 SQL Editor 執行設定指令');
  end if;

  if v_settings.locked_until is not null and v_settings.locked_until > now() then
    return jsonb_build_object('ok', false, 'error', '密碼錯誤次數過多，請稍後再試');
  end if;

  if extensions.crypt(coalesce(p_password, ''), v_settings.password_hash) <> v_settings.password_hash then
    v_failed := case when v_settings.locked_until is not null and v_settings.locked_until <= now() then 1 else v_settings.failed_attempts + 1 end;
    update private.admin_settings
    set failed_attempts = v_failed,
        locked_until = case when v_failed >= 5 then now() + interval '5 minutes' else null end,
        updated_at = now()
    where id = 1;
    return jsonb_build_object('ok', false, 'error', case when v_failed >= 5 then '密碼錯誤次數過多，已暫停 5 分鐘' else '管理員密碼不正確' end);
  end if;

  update private.admin_settings
  set failed_attempts = 0, locked_until = null, updated_at = now()
  where id = 1;
  return jsonb_build_object('ok', true);
end;
$$;

revoke all on function private.check_admin_password(text) from public, anon, authenticated;

create or replace function public.admin_export_all(p_password text)
returns jsonb
language plpgsql
security definer
set search_path = ''
as $$
declare
  v_auth jsonb;
begin
  v_auth := private.check_admin_password(p_password);
  if not coalesce((v_auth->>'ok')::boolean, false) then return v_auth; end if;

  return jsonb_build_object(
    'ok', true,
    'exported_at', now(),
    'data', jsonb_build_object(
      'products', coalesce((
        select jsonb_agg(jsonb_build_object(
          'code', p.code, 'barcode', p.barcode, 'category', p.category, 'specification', p.specification,
          'size', p.size, 'chipbreaker', p.chipbreaker, 'material', p.material,
          'name', p.name, 'brand', p.brand, 'supplier', p.supplier, 'unit', p.unit,
          'stock_quantity', p.stock_quantity, 'unit_price', p.unit_price,
          'purchased_at', p.purchased_at, 'notes', p.notes
        ) order by p.code)
        from public.products p
      ), '[]'::jsonb),
      'movements', coalesce((
        select jsonb_agg(jsonb_build_object(
          'sequence', row_number, 'product_code', code, 'movement_type', movement_type,
          'movement_date', movement_date, 'quantity', quantity, 'unit_price', unit_price,
          'supplier', supplier, 'notes', notes
        ) order by row_number)
        from (
          select row_number() over (order by m.movement_date, m.created_at, m.id) as row_number,
                 p.code, m.movement_type, m.movement_date, m.quantity, m.unit_price, m.supplier, m.notes
          from public.stock_movements m join public.products p on p.id = m.product_id
        ) ordered_movements
      ), '[]'::jsonb),
      'production_records', coalesce((
        select jsonb_agg(jsonb_build_object(
          'work_date', r.work_date, 'product_code', r.product_code, 'machine', r.machine,
          'hours', r.hours, 'minutes', r.minutes, 'quantity', r.quantity, 'notes', r.notes
        ) order by r.work_date, r.created_at, r.id)
        from public.production_records r
      ), '[]'::jsonb),
      'production_options', coalesce((
        select jsonb_agg(jsonb_build_object('option_type', o.option_type, 'value', o.value)
          order by o.option_type, o.value)
        from public.production_options o
      ), '[]'::jsonb)
    )
  );
end;
$$;

create or replace function public.admin_import_all(p_password text, p_payload jsonb)
returns jsonb
language plpgsql
security definer
set search_path = ''
as $$
declare
  v_auth jsonb;
  v_row record;
  v_item jsonb;
  v_code text;
  v_product_id uuid;
  v_current integer;
  v_target integer;
  v_quantity integer;
  v_type text;
  v_product_count integer := 0;
  v_movement_count integer := 0;
  v_record_count integer := 0;
  v_option_count integer := 0;
begin
  v_auth := private.check_admin_password(p_password);
  if not coalesce((v_auth->>'ok')::boolean, false) then return v_auth; end if;

  if jsonb_typeof(p_payload) <> 'object'
     or jsonb_typeof(p_payload->'products') <> 'array'
     or jsonb_typeof(p_payload->'movements') <> 'array'
     or jsonb_typeof(p_payload->'production_records') <> 'array'
     or jsonb_typeof(p_payload->'production_options') <> 'array' then
    raise exception '匯入資料格式不正確';
  end if;

  if jsonb_array_length(p_payload->'products') > 20000
     or jsonb_array_length(p_payload->'movements') > 100000
     or jsonb_array_length(p_payload->'production_records') > 100000
     or jsonb_array_length(p_payload->'production_options') > 20000 then
    raise exception '匯入資料筆數超過安全上限';
  end if;

  -- Supabase 的 safeupdate 會拒絕沒有 WHERE 的 DELETE；id 為主鍵且不可為空，
  -- 因此這些條件仍明確代表完整清除，並可在同一交易內安全還原。
  delete from public.stock_movements where id is not null;
  delete from public.products where id is not null;
  delete from public.production_records where id is not null;
  delete from public.production_options where id is not null;

  for v_row in select value as item from jsonb_array_elements(p_payload->'products') loop
    v_item := v_row.item;
    v_code := upper(btrim(coalesce(v_item->>'code', '')));
    if v_code = '' then raise exception '刀具資料中有空白編號'; end if;
    if coalesce(v_item->>'category', '') not in ('刀片', '鑽銑刀', '刀柄', '其他') then
      raise exception '刀具 % 的分類不正確', v_code;
    end if;
    v_target := coalesce((v_item->>'stock_quantity')::integer, 0);
    if v_target < 0 then raise exception '刀具 % 的庫存不可小於 0', v_code; end if;

    insert into public.products (
      code, barcode, category, specification, size, chipbreaker, material, brand, supplier,
      unit, stock_quantity, unit_price, purchased_at, notes
    ) values (
      v_code, btrim(coalesce(v_item->>'barcode', '')), v_item->>'category', btrim(coalesce(v_item->>'specification', '')),
      btrim(coalesce(v_item->>'size', '')), btrim(coalesce(v_item->>'chipbreaker', '')),
      btrim(coalesce(v_item->>'material', '')), btrim(coalesce(v_item->>'brand', '')),
      btrim(coalesce(v_item->>'supplier', '')), coalesce(nullif(btrim(v_item->>'unit'), ''), '個'),
      0, greatest(coalesce((v_item->>'unit_price')::numeric, 0), 0),
      nullif(v_item->>'purchased_at', '')::date, btrim(coalesce(v_item->>'notes', ''))
    );
    v_product_count := v_product_count + 1;
  end loop;

  for v_row in select value as item from jsonb_array_elements(p_payload->'movements') with ordinality order by ordinality loop
    v_item := v_row.item;
    v_code := upper(btrim(coalesce(v_item->>'product_code', '')));
    select id, stock_quantity into v_product_id, v_current from public.products where code = v_code;
    if not found then raise exception '出入庫紀錄找不到刀具編號 %', v_code; end if;
    v_type := lower(coalesce(v_item->>'movement_type', ''));
    if v_type not in ('in', 'out') then raise exception '刀具 % 的出入庫類型不正確', v_code; end if;
    v_quantity := coalesce((v_item->>'quantity')::integer, 0);
    if v_quantity <= 0 then raise exception '刀具 % 的出入庫數量必須大於 0', v_code; end if;
    if v_type = 'out' and v_current < v_quantity then
      raise exception '刀具 % 的歷史紀錄會造成負庫存（目前 %，出庫 %）', v_code, v_current, v_quantity;
    end if;
    v_current := v_current + case when v_type = 'in' then v_quantity else -v_quantity end;

    insert into public.stock_movements (
      product_id, movement_type, quantity, movement_date, unit_price, supplier,
      notes, stock_after, request_key
    ) values (
      v_product_id, v_type, v_quantity, coalesce(nullif(v_item->>'movement_date', '')::date, current_date),
      case when v_type = 'in' then nullif(v_item->>'unit_price', '')::numeric else null end,
      btrim(coalesce(v_item->>'supplier', '')), btrim(coalesce(v_item->>'notes', '')),
      v_current, pg_catalog.gen_random_uuid()
    );
    update public.products set stock_quantity = v_current where id = v_product_id;
    v_movement_count := v_movement_count + 1;
  end loop;

  -- Excel 的「目前庫存」是最終值；若與歷史加總不同，自動建立盤點調整紀錄。
  for v_row in select value as item from jsonb_array_elements(p_payload->'products') loop
    v_item := v_row.item;
    v_code := upper(btrim(v_item->>'code'));
    select id, stock_quantity into v_product_id, v_current from public.products where code = v_code;
    v_target := coalesce((v_item->>'stock_quantity')::integer, 0);
    if v_target <> v_current then
      insert into public.stock_movements (
        product_id, movement_type, quantity, movement_date, unit_price, supplier,
        notes, stock_after, request_key
      ) values (
        v_product_id, case when v_target > v_current then 'in' else 'out' end,
        abs(v_target - v_current), current_date,
        case when v_target > v_current then greatest(coalesce((v_item->>'unit_price')::numeric, 0), 0) else null end,
        case when v_target > v_current then btrim(coalesce(v_item->>'supplier', '')) else '' end,
        'Excel 匯入盤點調整', v_target, pg_catalog.gen_random_uuid()
      );
      v_movement_count := v_movement_count + 1;
    end if;
    update public.products set
      stock_quantity = v_target,
      unit_price = greatest(coalesce((v_item->>'unit_price')::numeric, 0), 0),
      supplier = btrim(coalesce(v_item->>'supplier', '')),
      purchased_at = nullif(v_item->>'purchased_at', '')::date
    where id = v_product_id;
  end loop;

  for v_row in select value as item from jsonb_array_elements(p_payload->'production_records') loop
    v_item := v_row.item;
    v_code := upper(btrim(coalesce(v_item->>'product_code', '')));
    if v_code = '' or btrim(coalesce(v_item->>'machine', '')) = '' then
      raise exception '加工紀錄的產品編號或機台不可空白';
    end if;
    if coalesce((v_item->>'hours')::integer, 0) < 0
       or coalesce((v_item->>'minutes')::integer, 0) not between 0 and 59
       or coalesce((v_item->>'hours')::integer, 0) * 60 + coalesce((v_item->>'minutes')::integer, 0) <= 0
       or coalesce((v_item->>'quantity')::integer, 0) <= 0 then
      raise exception '產品 % 的加工時間或數量不正確', v_code;
    end if;
    insert into public.production_records (work_date, product_code, machine, hours, minutes, quantity, notes)
    values (coalesce(nullif(v_item->>'work_date', '')::date, current_date), v_code,
      upper(btrim(v_item->>'machine')), (v_item->>'hours')::integer, (v_item->>'minutes')::integer,
      (v_item->>'quantity')::integer, btrim(coalesce(v_item->>'notes', '')));
    v_record_count := v_record_count + 1;
  end loop;

  for v_row in select value as item from jsonb_array_elements(p_payload->'production_options') loop
    v_item := v_row.item;
    if coalesce(v_item->>'option_type', '') not in ('product', 'machine') or btrim(coalesce(v_item->>'value', '')) = '' then
      raise exception '管理選項格式不正確';
    end if;
    insert into public.production_options (option_type, value)
    values (v_item->>'option_type', upper(btrim(v_item->>'value')))
    on conflict do nothing;
  end loop;

  insert into public.production_options (option_type, value)
  select 'product', product_code from public.production_records on conflict do nothing;
  insert into public.production_options (option_type, value)
  select 'machine', machine from public.production_records on conflict do nothing;
  select count(*) into v_option_count from public.production_options;

  return jsonb_build_object('ok', true, 'counts', jsonb_build_object(
    'products', v_product_count, 'movements', v_movement_count,
    'production_records', v_record_count, 'production_options', v_option_count
  ));
end;
$$;

-- 所有刪除操作都必須經過管理員密碼驗證。
create or replace function public.admin_delete_product(p_password text, p_product_id uuid)
returns jsonb
language plpgsql
security definer
set search_path = ''
as $$
declare
  v_auth jsonb;
begin
  v_auth := private.check_admin_password(p_password);
  if not coalesce((v_auth->>'ok')::boolean, false) then return v_auth; end if;
  delete from public.products where id = p_product_id and user_id is null;
  if not found then raise exception '找不到這筆刀具資料'; end if;
  return jsonb_build_object('ok', true);
end;
$$;

create or replace function public.admin_delete_stock_movement(p_password text, p_movement_id uuid)
returns jsonb
language plpgsql
security definer
set search_path = ''
as $$
declare
  v_auth jsonb;
begin
  v_auth := private.check_admin_password(p_password);
  if not coalesce((v_auth->>'ok')::boolean, false) then return v_auth; end if;
  perform public.delete_stock_movement(p_movement_id);
  return jsonb_build_object('ok', true);
end;
$$;

create or replace function public.admin_delete_production_record(p_password text, p_record_id uuid)
returns jsonb
language plpgsql
security definer
set search_path = ''
as $$
declare
  v_auth jsonb;
begin
  v_auth := private.check_admin_password(p_password);
  if not coalesce((v_auth->>'ok')::boolean, false) then return v_auth; end if;
  delete from public.production_records where id = p_record_id and user_id is null;
  if not found then raise exception '找不到這筆加工紀錄'; end if;
  return jsonb_build_object('ok', true);
end;
$$;

create or replace function public.admin_delete_production_option(p_password text, p_option_type text, p_value text)
returns jsonb
language plpgsql
security definer
set search_path = ''
as $$
declare
  v_auth jsonb;
begin
  v_auth := private.check_admin_password(p_password);
  if not coalesce((v_auth->>'ok')::boolean, false) then return v_auth; end if;
  if p_option_type not in ('product', 'machine') then raise exception '管理選項類型不正確'; end if;
  delete from public.production_options
  where user_id is null and option_type = p_option_type and lower(value) = lower(btrim(p_value));
  if not found then raise exception '找不到這筆管理選項'; end if;
  return jsonb_build_object('ok', true);
end;
$$;

create or replace function public.admin_delete_production_product_data(p_password text, p_product_code text)
returns jsonb
language plpgsql
security definer
set search_path = ''
as $$
declare
  v_auth jsonb;
  v_deleted integer;
begin
  v_auth := private.check_admin_password(p_password);
  if not coalesce((v_auth->>'ok')::boolean, false) then return v_auth; end if;
  v_deleted := public.delete_production_product_data(p_product_code);
  return jsonb_build_object('ok', true, 'deleted', v_deleted);
end;
$$;

-- 移除可繞過密碼的舊刪除權限；新增與修改權限不受影響。
revoke delete on table public.products from anon, authenticated;
revoke delete on table public.production_records from anon, authenticated;
revoke delete on table public.production_options from anon, authenticated;
revoke execute on function public.delete_stock_movement(uuid) from anon, authenticated;
revoke execute on function public.delete_production_product_data(text) from anon, authenticated;

revoke all on function public.admin_export_all(text) from public;
revoke all on function public.admin_import_all(text, jsonb) from public;
grant execute on function public.admin_export_all(text) to anon, authenticated;
grant execute on function public.admin_import_all(text, jsonb) to anon, authenticated;
revoke all on function public.admin_delete_product(text, uuid) from public;
revoke all on function public.admin_delete_stock_movement(text, uuid) from public;
revoke all on function public.admin_delete_production_record(text, uuid) from public;
revoke all on function public.admin_delete_production_option(text, text, text) from public;
revoke all on function public.admin_delete_production_product_data(text, text) from public;
grant execute on function public.admin_delete_product(text, uuid) to anon, authenticated;
grant execute on function public.admin_delete_stock_movement(text, uuid) to anon, authenticated;
grant execute on function public.admin_delete_production_record(text, uuid) to anon, authenticated;
grant execute on function public.admin_delete_production_option(text, text, text) to anon, authenticated;
grant execute on function public.admin_delete_production_product_data(text, text) to anon, authenticated;

comment on function public.admin_export_all(text) is '管理員密碼驗證後匯出完整 Excel 備份資料';
comment on function public.admin_import_all(text, jsonb) is '管理員密碼驗證後，以單一交易整批還原 Excel 資料';
comment on function public.admin_delete_product(text, uuid) is '管理員密碼驗證後刪除刀具與其出入庫歷史';
comment on function public.admin_delete_stock_movement(text, uuid) is '管理員密碼驗證後刪除出入庫紀錄並重算庫存';
comment on function public.admin_delete_production_record(text, uuid) is '管理員密碼驗證後刪除單筆加工紀錄';
comment on function public.admin_delete_production_option(text, text, text) is '管理員密碼驗證後刪除產品編號或機台選項';
comment on function public.admin_delete_production_product_data(text, text) is '管理員密碼驗證後永久刪除產品編號及全部加工紀錄';

-- 本檔執行成功後，另外執行以下一行設定 8 位數字管理員密碼：
-- select private.set_admin_password('請改成8位數字密碼');
