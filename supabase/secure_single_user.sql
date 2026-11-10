-- 詮栩 CNC 雲端系統：單一帳號安全升級
-- 執行前請先從網站下載完整 Excel 備份。
-- 本檔只需要在 Supabase SQL Editor 的「New query」完整執行一次。

begin;

-- 固定授權帳號：kuanlin / kuanlin@chyuansheu.com
do $$
declare
  v_email text;
begin
  select email into v_email
  from auth.users
  where id = '5568706d-caaa-459b-b842-13cdbc3929fc'::uuid;

  if not found then
    raise exception '找不到指定的 Supabase 使用者 UID，已停止升級，資料尚未變更';
  end if;
  if lower(coalesce(v_email, '')) <> 'kuanlin@chyuansheu.com' then
    raise exception '指定 UID 的 Email 不相符，已停止升級，資料尚未變更';
  end if;

  if exists (select 1 from public.products where user_id is not null and user_id <> '5568706d-caaa-459b-b842-13cdbc3929fc'::uuid)
     or exists (select 1 from public.stock_movements where user_id is not null and user_id <> '5568706d-caaa-459b-b842-13cdbc3929fc'::uuid)
     or exists (select 1 from public.production_records where user_id is not null and user_id <> '5568706d-caaa-459b-b842-13cdbc3929fc'::uuid)
     or exists (select 1 from public.production_options where user_id is not null and user_id <> '5568706d-caaa-459b-b842-13cdbc3929fc'::uuid) then
    raise exception '資料庫內含其他帳號的資料，為避免誤覆蓋已停止升級';
  end if;
end;
$$;

-- 將原本免登入共用資料全部歸戶給指定帳號。
update public.products set user_id = '5568706d-caaa-459b-b842-13cdbc3929fc'::uuid where user_id is null;
update public.stock_movements set user_id = '5568706d-caaa-459b-b842-13cdbc3929fc'::uuid where user_id is null;
update public.production_records set user_id = '5568706d-caaa-459b-b842-13cdbc3929fc'::uuid where user_id is null;
update public.production_options set user_id = '5568706d-caaa-459b-b842-13cdbc3929fc'::uuid where user_id is null;

alter table public.products alter column user_id set not null;
alter table public.stock_movements alter column user_id set not null;
alter table public.production_records alter column user_id set not null;
alter table public.production_options alter column user_id set not null;
alter table public.products alter column user_id set default auth.uid();
alter table public.stock_movements alter column user_id set default auth.uid();
alter table public.production_records alter column user_id set default auth.uid();
alter table public.production_options alter column user_id set default auth.uid();

drop index if exists public.products_shared_code_unique;
drop index if exists public.stock_movements_shared_request_unique;
drop index if exists public.production_options_shared_value_unique;
create unique index if not exists production_options_user_value_unique
  on public.production_options(user_id, option_type, lower(value));

create schema if not exists private;
revoke all on schema private from public, anon, authenticated;

-- 所有安全交易函式都先經過此檢查；前端公開 UID 並不等於取得登入權限。
create or replace function private.require_app_user()
returns uuid
language plpgsql
stable
security definer
set search_path = ''
as $$
declare
  v_user_id uuid := auth.uid();
begin
  if v_user_id is null then
    raise exception '請先登入系統' using errcode = '42501';
  end if;
  if v_user_id <> '5568706d-caaa-459b-b842-13cdbc3929fc'::uuid then
    raise exception '這個帳號沒有系統使用權限' using errcode = '42501';
  end if;
  return v_user_id;
end;
$$;
revoke all on function private.require_app_user() from public, anon, authenticated;

-- 清除舊共用政策，改為只允許目前登入者讀寫自己的資料。
do $$
declare
  v_policy record;
begin
  for v_policy in
    select schemaname, tablename, policyname
    from pg_policies
    where schemaname = 'public'
      and tablename in ('products', 'stock_movements', 'production_records', 'production_options')
  loop
    execute format('drop policy if exists %I on %I.%I', v_policy.policyname, v_policy.schemaname, v_policy.tablename);
  end loop;
end;
$$;

create policy products_owner_select on public.products for select to authenticated
  using ((select auth.uid()) = user_id and user_id = '5568706d-caaa-459b-b842-13cdbc3929fc'::uuid);
create policy products_owner_update on public.products for update to authenticated
  using ((select auth.uid()) = user_id and user_id = '5568706d-caaa-459b-b842-13cdbc3929fc'::uuid)
  with check ((select auth.uid()) = user_id and user_id = '5568706d-caaa-459b-b842-13cdbc3929fc'::uuid);
create policy movements_owner_select on public.stock_movements for select to authenticated
  using ((select auth.uid()) = user_id and user_id = '5568706d-caaa-459b-b842-13cdbc3929fc'::uuid);
create policy production_records_owner_select on public.production_records for select to authenticated
  using ((select auth.uid()) = user_id and user_id = '5568706d-caaa-459b-b842-13cdbc3929fc'::uuid);
create policy production_records_owner_insert on public.production_records for insert to authenticated
  with check ((select auth.uid()) = user_id and user_id = '5568706d-caaa-459b-b842-13cdbc3929fc'::uuid);
create policy production_records_owner_update on public.production_records for update to authenticated
  using ((select auth.uid()) = user_id and user_id = '5568706d-caaa-459b-b842-13cdbc3929fc'::uuid)
  with check ((select auth.uid()) = user_id and user_id = '5568706d-caaa-459b-b842-13cdbc3929fc'::uuid);
create policy production_options_owner_select on public.production_options for select to authenticated
  using ((select auth.uid()) = user_id and user_id = '5568706d-caaa-459b-b842-13cdbc3929fc'::uuid);
create policy production_options_owner_insert on public.production_options for insert to authenticated
  with check ((select auth.uid()) = user_id and user_id = '5568706d-caaa-459b-b842-13cdbc3929fc'::uuid);

revoke all on table public.products from anon, authenticated;
revoke all on table public.stock_movements from anon, authenticated;
revoke all on table public.production_records from anon, authenticated;
revoke all on table public.production_options from anon, authenticated;
grant select on table public.products to authenticated;
grant update (specification, size, chipbreaker, material, name, brand, supplier, category, unit, notes) on table public.products to authenticated;
grant select on table public.stock_movements to authenticated;
grant select, insert on table public.production_records to authenticated;
grant update (work_date, product_code, machine, hours, minutes, quantity, notes) on table public.production_records to authenticated;
grant select, insert on table public.production_options to authenticated;

-- 安全建立刀具與期初入庫。
create or replace function public.create_product_with_initial_stock(
  p_specification text, p_size text default '', p_chipbreaker text default null,
  p_material text default null, p_brand text default null, p_supplier text default null,
  p_category text default '刀片', p_unit text default '片', p_notes text default null,
  p_initial_quantity integer default 0, p_initial_unit_price numeric default 0,
  p_purchased_at date default current_date, p_request_key uuid default gen_random_uuid()
)
returns uuid
language plpgsql
security definer
set search_path = ''
as $$
declare
  v_user_id uuid := private.require_app_user();
  v_product_id uuid;
  v_code text;
  v_prefix text;
  v_next_number integer;
begin
  if p_specification is null or btrim(p_specification) = '' then raise exception '規格為必填欄位'; end if;
  if p_category not in ('刀片', '鑽銑刀', '刀柄', '其他') then raise exception '商品分類不正確'; end if;
  if p_category in ('刀片', '鑽銑刀') and btrim(coalesce(p_size, '')) = '' then raise exception '此分類需要輸入尺寸'; end if;
  if p_category in ('刀片', '鑽銑刀') and btrim(coalesce(p_material, '')) = '' then raise exception '此分類需要輸入材質'; end if;
  if coalesce(p_initial_quantity, 0) < 0 then raise exception '期初庫存不能小於 0'; end if;
  if coalesce(p_initial_unit_price, 0) < 0 then raise exception '單價不能小於 0'; end if;

  perform pg_advisory_xact_lock(hashtextextended(v_user_id::text || '-inventory', 73001));
  v_prefix := case p_category when '刀片' then 'T' when '鑽銑刀' then 'M' when '刀柄' then 'H' else 'O' end;
  select coalesce(max(substring(code from 2)::integer), 0) + 1 into v_next_number
  from public.products
  where user_id = v_user_id and category = p_category and code ~ ('^' || v_prefix || '[0-9]+$');
  v_code := v_prefix || lpad(v_next_number::text, 4, '0');

  insert into public.products (
    user_id, code, name, specification, size, chipbreaker, material, brand,
    supplier, category, unit, stock_quantity, unit_price, purchased_at, notes
  ) values (
    v_user_id, v_code, upper(btrim(p_specification)), upper(btrim(p_specification)),
    case when p_category in ('刀片', '鑽銑刀') then upper(btrim(coalesce(p_size, ''))) else '' end,
    case when p_category = '刀片' then upper(btrim(coalesce(p_chipbreaker, ''))) else '' end,
    case when p_category in ('刀片', '鑽銑刀') then upper(btrim(coalesce(p_material, ''))) else '' end,
    btrim(coalesce(p_brand, '')), btrim(coalesce(p_supplier, '')), p_category,
    coalesce(nullif(btrim(p_unit), ''), '片'), p_initial_quantity,
    case when p_initial_quantity > 0 then p_initial_unit_price else 0 end,
    case when p_initial_quantity > 0 then p_purchased_at else null end,
    btrim(coalesce(p_notes, ''))
  ) returning id into v_product_id;

  if p_initial_quantity > 0 then
    insert into public.stock_movements (
      user_id, product_id, movement_type, quantity, movement_date, unit_price,
      supplier, notes, stock_after, request_key
    ) values (
      v_user_id, v_product_id, 'in', p_initial_quantity, coalesce(p_purchased_at, current_date),
      p_initial_unit_price, btrim(coalesce(p_supplier, '')), '建立商品時的期初庫存',
      p_initial_quantity, p_request_key
    );
  end if;
  return v_product_id;
end;
$$;

-- 安全登記出入庫。
create or replace function public.register_stock_movement(
  p_product_id uuid, p_movement_type text, p_quantity integer,
  p_movement_date date default current_date, p_unit_price numeric default null,
  p_supplier text default null, p_notes text default null,
  p_request_key uuid default gen_random_uuid()
)
returns uuid
language plpgsql
security definer
set search_path = ''
as $$
declare
  v_user_id uuid := private.require_app_user();
  v_product public.products%rowtype;
  v_next_stock integer;
  v_movement_id uuid;
begin
  if p_movement_type not in ('in', 'out') then raise exception '出入庫類型必須是 in 或 out'; end if;
  if p_quantity is null or p_quantity <= 0 then raise exception '數量必須大於 0'; end if;
  if p_movement_type = 'in' and coalesce(p_unit_price, 0) < 0 then raise exception '單價不能小於 0'; end if;

  select id into v_movement_id from public.stock_movements
  where user_id = v_user_id and request_key = p_request_key;
  if v_movement_id is not null then return v_movement_id; end if;

  select * into v_product from public.products
  where id = p_product_id and user_id = v_user_id for update;
  if not found then raise exception '找不到商品或沒有權限'; end if;

  v_next_stock := v_product.stock_quantity + case when p_movement_type = 'in' then p_quantity else -p_quantity end;
  if v_next_stock < 0 then raise exception '庫存不足，目前只有 % %', v_product.stock_quantity, v_product.unit; end if;

  update public.products set
    stock_quantity = v_next_stock,
    unit_price = case when p_movement_type = 'in' then coalesce(p_unit_price, 0) else unit_price end,
    supplier = case when p_movement_type = 'in' and btrim(coalesce(p_supplier, '')) <> '' then btrim(p_supplier) else supplier end,
    purchased_at = case when p_movement_type = 'in' then coalesce(p_movement_date, current_date) else purchased_at end
  where id = p_product_id and user_id = v_user_id;

  insert into public.stock_movements (
    user_id, product_id, movement_type, quantity, movement_date, unit_price,
    supplier, notes, stock_after, request_key
  ) values (
    v_user_id, p_product_id, p_movement_type, p_quantity, coalesce(p_movement_date, current_date),
    case when p_movement_type = 'in' then coalesce(p_unit_price, 0) else null end,
    case when p_movement_type = 'in' then btrim(coalesce(p_supplier, '')) else '' end,
    btrim(coalesce(p_notes, '')), v_next_stock, p_request_key
  ) returning id into v_movement_id;
  return v_movement_id;
end;
$$;

-- 盤點以差額建立歷史，不覆寫既有單價。
create or replace function public.set_stocktake_quantity(
  p_product_id uuid, p_actual_quantity integer, p_notes text default null,
  p_request_key uuid default gen_random_uuid()
)
returns uuid
language plpgsql
security definer
set search_path = ''
as $$
declare
  v_user_id uuid := private.require_app_user();
  v_product public.products%rowtype;
  v_difference integer;
  v_movement_id uuid;
  v_note text;
begin
  if p_actual_quantity is null or p_actual_quantity < 0 then raise exception '實際盤點數量必須是 0 以上的整數'; end if;
  select id into v_movement_id from public.stock_movements where user_id = v_user_id and request_key = p_request_key;
  if v_movement_id is not null then return v_movement_id; end if;
  select * into v_product from public.products where id = p_product_id and user_id = v_user_id for update;
  if not found then raise exception '找不到刀具或沒有權限'; end if;
  v_difference := p_actual_quantity - v_product.stock_quantity;
  if v_difference = 0 then return null; end if;
  v_note := format('庫存盤點調整：%s → %s', v_product.stock_quantity, p_actual_quantity);
  if btrim(coalesce(p_notes, '')) <> '' then v_note := v_note || '｜' || btrim(p_notes); end if;
  insert into public.stock_movements (
    user_id, product_id, movement_type, quantity, movement_date, unit_price,
    supplier, notes, stock_after, request_key
  ) values (
    v_user_id, v_product.id, case when v_difference > 0 then 'in' else 'out' end,
    abs(v_difference), current_date, case when v_difference > 0 then v_product.unit_price else null end,
    case when v_difference > 0 then v_product.supplier else '' end, v_note,
    p_actual_quantity, p_request_key
  ) returning id into v_movement_id;
  update public.products set stock_quantity = p_actual_quantity where id = v_product.id and user_id = v_user_id;
  return v_movement_id;
end;
$$;

-- 刪除一筆出入庫後，重算該刀具的全部歷史餘額與最新單價。
create or replace function public.delete_stock_movement(p_movement_id uuid)
returns void
language plpgsql
security definer
set search_path = ''
as $$
declare
  v_user_id uuid := private.require_app_user();
  v_product_id uuid;
  v_min_stock integer;
  v_final_stock integer := 0;
  v_latest_price numeric(14,2);
  v_latest_date date;
  v_latest_supplier text;
begin
  select product_id into v_product_id from public.stock_movements
  where id = p_movement_id and user_id = v_user_id;
  if not found then raise exception '找不到出入庫紀錄或沒有權限'; end if;
  perform 1 from public.products where id = v_product_id and user_id = v_user_id for update;
  if not found then raise exception '找不到商品或沒有權限'; end if;
  delete from public.stock_movements where id = p_movement_id and user_id = v_user_id;

  with balances as (
    select sum(case when movement_type = 'in' then quantity else -quantity end)
      over (order by movement_date, created_at, id) as balance
    from public.stock_movements where product_id = v_product_id and user_id = v_user_id
  ) select coalesce(min(balance), 0) into v_min_stock from balances;
  if v_min_stock < 0 then raise exception '無法刪除：這會使後續某筆歷史庫存變成負數'; end if;

  with balances as (
    select id, sum(case when movement_type = 'in' then quantity else -quantity end)
      over (order by movement_date, created_at, id) as balance
    from public.stock_movements where product_id = v_product_id and user_id = v_user_id
  )
  update public.stock_movements m set stock_after = b.balance from balances b where m.id = b.id;

  select coalesce(sum(case when movement_type = 'in' then quantity else -quantity end), 0)
  into v_final_stock from public.stock_movements where product_id = v_product_id and user_id = v_user_id;
  select unit_price, movement_date, supplier into v_latest_price, v_latest_date, v_latest_supplier
  from public.stock_movements
  where product_id = v_product_id and user_id = v_user_id and movement_type = 'in'
  order by movement_date desc, created_at desc, id desc limit 1;
  update public.products set stock_quantity = v_final_stock, unit_price = coalesce(v_latest_price, 0),
    purchased_at = v_latest_date,
    supplier = case when btrim(coalesce(v_latest_supplier, '')) <> '' then v_latest_supplier else supplier end
  where id = v_product_id and user_id = v_user_id;
end;
$$;

create or replace function public.delete_production_product_data(p_product_code text)
returns integer
language plpgsql
security definer
set search_path = ''
as $$
declare
  v_user_id uuid := private.require_app_user();
  v_product_code text := upper(btrim(coalesce(p_product_code, '')));
  v_deleted_count integer := 0;
begin
  if v_product_code = '' then raise exception '產品編號不能空白'; end if;
  delete from public.production_records where user_id = v_user_id and upper(product_code) = v_product_code;
  get diagnostics v_deleted_count = row_count;
  delete from public.production_options
  where user_id = v_user_id and option_type = 'product' and upper(value) = v_product_code;
  return v_deleted_count;
end;
$$;

-- 舊 8 位數管理密碼函式停用；刪除改由 Supabase 登入密碼重新驗證後呼叫。
drop function if exists public.admin_delete_product(text, uuid);
drop function if exists public.admin_delete_stock_movement(text, uuid);
drop function if exists public.admin_delete_production_record(text, uuid);
drop function if exists public.admin_delete_production_option(text, text, text);
drop function if exists public.admin_delete_production_product_data(text, text);

create or replace function public.admin_delete_product(p_product_id uuid)
returns jsonb language plpgsql security definer set search_path = '' as $$
declare v_user_id uuid := private.require_app_user();
begin
  delete from public.products where id = p_product_id and user_id = v_user_id;
  if not found then raise exception '找不到這筆刀具資料'; end if;
  return jsonb_build_object('ok', true);
end; $$;

create or replace function public.admin_delete_stock_movement(p_movement_id uuid)
returns jsonb language plpgsql security definer set search_path = '' as $$
begin
  perform private.require_app_user();
  perform public.delete_stock_movement(p_movement_id);
  return jsonb_build_object('ok', true);
end; $$;

create or replace function public.admin_delete_production_record(p_record_id uuid)
returns jsonb language plpgsql security definer set search_path = '' as $$
declare v_user_id uuid := private.require_app_user();
begin
  delete from public.production_records where id = p_record_id and user_id = v_user_id;
  if not found then raise exception '找不到這筆加工紀錄'; end if;
  return jsonb_build_object('ok', true);
end; $$;

create or replace function public.admin_delete_production_option(p_option_type text, p_value text)
returns jsonb language plpgsql security definer set search_path = '' as $$
declare v_user_id uuid := private.require_app_user();
begin
  if p_option_type not in ('product', 'machine') then raise exception '管理選項類型不正確'; end if;
  delete from public.production_options
  where user_id = v_user_id and option_type = p_option_type and lower(value) = lower(btrim(p_value));
  if not found then raise exception '找不到這筆管理選項'; end if;
  return jsonb_build_object('ok', true);
end; $$;

create or replace function public.admin_delete_production_product_data(p_product_code text)
returns jsonb language plpgsql security definer set search_path = '' as $$
declare v_deleted integer;
begin
  perform private.require_app_user();
  v_deleted := public.delete_production_product_data(p_product_code);
  return jsonb_build_object('ok', true, 'deleted', v_deleted);
end; $$;

-- 完整備份只讀取登入者的資料。
drop function if exists public.admin_export_all(text);
create or replace function public.admin_export_all()
returns jsonb
language plpgsql
security definer
set search_path = ''
as $$
declare v_user_id uuid := private.require_app_user();
begin
  return jsonb_build_object(
    'ok', true, 'exported_at', now(),
    'data', jsonb_build_object(
      'products', coalesce((select jsonb_agg(jsonb_build_object(
        'code', p.code, 'category', p.category, 'specification', p.specification,
        'size', p.size, 'chipbreaker', p.chipbreaker, 'material', p.material,
        'name', p.name, 'brand', p.brand, 'supplier', p.supplier, 'unit', p.unit,
        'stock_quantity', p.stock_quantity, 'unit_price', p.unit_price,
        'purchased_at', p.purchased_at, 'notes', p.notes
      ) order by p.code) from public.products p where p.user_id = v_user_id), '[]'::jsonb),
      'movements', coalesce((select jsonb_agg(jsonb_build_object(
        'sequence', row_number, 'product_code', code, 'movement_type', movement_type,
        'movement_date', movement_date, 'quantity', quantity, 'unit_price', unit_price,
        'supplier', supplier, 'notes', notes
      ) order by row_number) from (
        select row_number() over (order by m.movement_date, m.created_at, m.id) as row_number,
          p.code, m.movement_type, m.movement_date, m.quantity, m.unit_price, m.supplier, m.notes
        from public.stock_movements m join public.products p on p.id = m.product_id and p.user_id = v_user_id
        where m.user_id = v_user_id
      ) x), '[]'::jsonb),
      'production_records', coalesce((select jsonb_agg(jsonb_build_object(
        'work_date', r.work_date, 'product_code', r.product_code, 'machine', r.machine,
        'hours', r.hours, 'minutes', r.minutes, 'quantity', r.quantity, 'notes', r.notes
      ) order by r.work_date, r.created_at, r.id)
        from public.production_records r where r.user_id = v_user_id), '[]'::jsonb),
      'production_options', coalesce((select jsonb_agg(jsonb_build_object(
        'option_type', o.option_type, 'value', o.value) order by o.option_type, o.value)
        from public.production_options o where o.user_id = v_user_id), '[]'::jsonb)
    )
  );
end;
$$;

-- 完整還原：限制於登入者，並在單一交易內完成。
drop function if exists public.admin_import_all(text, jsonb);
create or replace function public.admin_import_all(p_payload jsonb)
returns jsonb
language plpgsql
security definer
set search_path = ''
as $$
declare
  v_user_id uuid := private.require_app_user();
  v_row record; v_item jsonb; v_code text; v_product_id uuid;
  v_current integer; v_target integer; v_quantity integer; v_type text;
  v_product_count integer := 0; v_movement_count integer := 0;
  v_record_count integer := 0; v_option_count integer := 0;
begin
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

  delete from public.stock_movements where user_id = v_user_id;
  delete from public.products where user_id = v_user_id;
  delete from public.production_records where user_id = v_user_id;
  delete from public.production_options where user_id = v_user_id;

  for v_row in select value as item from jsonb_array_elements(p_payload->'products') loop
    v_item := v_row.item; v_code := upper(btrim(coalesce(v_item->>'code', '')));
    if v_code = '' then raise exception '刀具資料中有空白編號'; end if;
    if coalesce(v_item->>'category', '') not in ('刀片', '鑽銑刀', '刀柄', '其他') then raise exception '刀具 % 的分類不正確', v_code; end if;
    v_target := coalesce((v_item->>'stock_quantity')::integer, 0);
    if v_target < 0 then raise exception '刀具 % 的庫存不可小於 0', v_code; end if;
    insert into public.products (
      user_id, code, category, specification, size, chipbreaker, material, brand,
      supplier, unit, stock_quantity, unit_price, purchased_at, notes, name
    ) values (
      v_user_id, v_code, v_item->>'category', btrim(coalesce(v_item->>'specification', '')),
      btrim(coalesce(v_item->>'size', '')), btrim(coalesce(v_item->>'chipbreaker', '')),
      btrim(coalesce(v_item->>'material', '')), btrim(coalesce(v_item->>'brand', '')),
      btrim(coalesce(v_item->>'supplier', '')), coalesce(nullif(btrim(v_item->>'unit'), ''), '個'),
      0, greatest(coalesce((v_item->>'unit_price')::numeric, 0), 0),
      nullif(v_item->>'purchased_at', '')::date, btrim(coalesce(v_item->>'notes', '')),
      coalesce(nullif(btrim(v_item->>'name'), ''), v_code)
    );
    v_product_count := v_product_count + 1;
  end loop;

  for v_row in select value as item from jsonb_array_elements(p_payload->'movements') with ordinality order by ordinality loop
    v_item := v_row.item; v_code := upper(btrim(coalesce(v_item->>'product_code', '')));
    select id, stock_quantity into v_product_id, v_current from public.products where user_id = v_user_id and code = v_code;
    if not found then raise exception '出入庫紀錄找不到刀具編號 %', v_code; end if;
    v_type := lower(coalesce(v_item->>'movement_type', ''));
    if v_type not in ('in', 'out') then raise exception '刀具 % 的出入庫類型不正確', v_code; end if;
    v_quantity := coalesce((v_item->>'quantity')::integer, 0);
    if v_quantity <= 0 then raise exception '刀具 % 的出入庫數量必須大於 0', v_code; end if;
    if v_type = 'out' and v_current < v_quantity then raise exception '刀具 % 的歷史紀錄會造成負庫存', v_code; end if;
    v_current := v_current + case when v_type = 'in' then v_quantity else -v_quantity end;
    insert into public.stock_movements (
      user_id, product_id, movement_type, quantity, movement_date, unit_price,
      supplier, notes, stock_after, request_key
    ) values (
      v_user_id, v_product_id, v_type, v_quantity,
      coalesce(nullif(v_item->>'movement_date', '')::date, current_date),
      case when v_type = 'in' then nullif(v_item->>'unit_price', '')::numeric else null end,
      btrim(coalesce(v_item->>'supplier', '')), btrim(coalesce(v_item->>'notes', '')),
      v_current, gen_random_uuid()
    );
    update public.products set stock_quantity = v_current where id = v_product_id and user_id = v_user_id;
    v_movement_count := v_movement_count + 1;
  end loop;

  for v_row in select value as item from jsonb_array_elements(p_payload->'products') loop
    v_item := v_row.item; v_code := upper(btrim(v_item->>'code'));
    select id, stock_quantity into v_product_id, v_current from public.products where user_id = v_user_id and code = v_code;
    v_target := coalesce((v_item->>'stock_quantity')::integer, 0);
    if v_target <> v_current then
      insert into public.stock_movements (
        user_id, product_id, movement_type, quantity, movement_date, unit_price,
        supplier, notes, stock_after, request_key
      ) values (
        v_user_id, v_product_id, case when v_target > v_current then 'in' else 'out' end,
        abs(v_target - v_current), current_date,
        case when v_target > v_current then greatest(coalesce((v_item->>'unit_price')::numeric, 0), 0) else null end,
        case when v_target > v_current then btrim(coalesce(v_item->>'supplier', '')) else '' end,
        'Excel 匯入盤點調整', v_target, gen_random_uuid()
      );
      v_movement_count := v_movement_count + 1;
    end if;
    update public.products set stock_quantity = v_target,
      unit_price = greatest(coalesce((v_item->>'unit_price')::numeric, 0), 0),
      supplier = btrim(coalesce(v_item->>'supplier', '')),
      purchased_at = nullif(v_item->>'purchased_at', '')::date
    where id = v_product_id and user_id = v_user_id;
  end loop;

  for v_row in select value as item from jsonb_array_elements(p_payload->'production_records') loop
    v_item := v_row.item; v_code := upper(btrim(coalesce(v_item->>'product_code', '')));
    if v_code = '' or btrim(coalesce(v_item->>'machine', '')) = '' then raise exception '加工紀錄的產品編號或機台不可空白'; end if;
    if coalesce((v_item->>'hours')::integer, 0) < 0
       or coalesce((v_item->>'minutes')::integer, 0) not between 0 and 59
       or coalesce((v_item->>'hours')::integer, 0) * 60 + coalesce((v_item->>'minutes')::integer, 0) <= 0
       or coalesce((v_item->>'quantity')::integer, 0) <= 0 then raise exception '產品 % 的加工時間或數量不正確', v_code; end if;
    insert into public.production_records (user_id, work_date, product_code, machine, hours, minutes, quantity, notes)
    values (v_user_id, coalesce(nullif(v_item->>'work_date', '')::date, current_date), v_code,
      upper(btrim(v_item->>'machine')), (v_item->>'hours')::integer, (v_item->>'minutes')::integer,
      (v_item->>'quantity')::integer, btrim(coalesce(v_item->>'notes', '')));
    v_record_count := v_record_count + 1;
  end loop;

  for v_row in select value as item from jsonb_array_elements(p_payload->'production_options') loop
    v_item := v_row.item;
    if coalesce(v_item->>'option_type', '') not in ('product', 'machine') or btrim(coalesce(v_item->>'value', '')) = '' then raise exception '管理選項格式不正確'; end if;
    insert into public.production_options (user_id, option_type, value)
    values (v_user_id, v_item->>'option_type', upper(btrim(v_item->>'value'))) on conflict do nothing;
  end loop;
  insert into public.production_options (user_id, option_type, value)
    select distinct v_user_id, 'product', product_code from public.production_records where user_id = v_user_id on conflict do nothing;
  insert into public.production_options (user_id, option_type, value)
    select distinct v_user_id, 'machine', machine from public.production_records where user_id = v_user_id on conflict do nothing;
  select count(*) into v_option_count from public.production_options where user_id = v_user_id;
  return jsonb_build_object('ok', true, 'counts', jsonb_build_object(
    'products', v_product_count, 'movements', v_movement_count,
    'production_records', v_record_count, 'production_options', v_option_count
  ));
end;
$$;

-- 舊管理員密碼儲存與檢查不再需要。
drop function if exists private.check_admin_password(text);
drop function if exists private.set_admin_password(text);
drop table if exists private.admin_settings;

-- 所有交易函式只開放給已登入帳號；匿名角色完全沒有資料與 RPC 權限。
revoke all on function public.create_product_with_initial_stock(text,text,text,text,text,text,text,text,text,integer,numeric,date,uuid) from public, anon, authenticated;
revoke all on function public.register_stock_movement(uuid,text,integer,date,numeric,text,text,uuid) from public, anon, authenticated;
revoke all on function public.set_stocktake_quantity(uuid,integer,text,uuid) from public, anon, authenticated;
revoke all on function public.delete_stock_movement(uuid) from public, anon, authenticated;
revoke all on function public.delete_production_product_data(text) from public, anon, authenticated;
revoke all on function public.admin_delete_product(uuid) from public, anon, authenticated;
revoke all on function public.admin_delete_stock_movement(uuid) from public, anon, authenticated;
revoke all on function public.admin_delete_production_record(uuid) from public, anon, authenticated;
revoke all on function public.admin_delete_production_option(text,text) from public, anon, authenticated;
revoke all on function public.admin_delete_production_product_data(text) from public, anon, authenticated;
revoke all on function public.admin_export_all() from public, anon, authenticated;
revoke all on function public.admin_import_all(jsonb) from public, anon, authenticated;

grant execute on function public.create_product_with_initial_stock(text,text,text,text,text,text,text,text,text,integer,numeric,date,uuid) to authenticated;
grant execute on function public.register_stock_movement(uuid,text,integer,date,numeric,text,text,uuid) to authenticated;
grant execute on function public.set_stocktake_quantity(uuid,integer,text,uuid) to authenticated;
grant execute on function public.admin_delete_product(uuid) to authenticated;
grant execute on function public.admin_delete_stock_movement(uuid) to authenticated;
grant execute on function public.admin_delete_production_record(uuid) to authenticated;
grant execute on function public.admin_delete_production_option(text,text) to authenticated;
grant execute on function public.admin_delete_production_product_data(text) to authenticated;
grant execute on function public.admin_export_all() to authenticated;
grant execute on function public.admin_import_all(jsonb) to authenticated;

comment on table public.production_records is '僅限指定 Supabase 帳號存取的加工數量、機台與加工時間紀錄。';
comment on table public.production_options is '僅限指定 Supabase 帳號存取的產品編號與機台選單。';

commit;

-- 成功時 Results 會顯示：Success. No rows returned
