-- 刀片庫存管理系統 / Supabase PostgreSQL schema
-- 可重複執行。請在 Supabase Dashboard > SQL Editor 貼上並執行整份檔案。

create extension if not exists pgcrypto;

create table if not exists public.products (
  id uuid primary key default gen_random_uuid(),
  user_id uuid references auth.users(id) on delete cascade,
  code text not null,
  name text not null,
  specification text not null,
  size text not null default '',
  chipbreaker text not null default '',
  material text not null default '',
  brand text not null default '',
  supplier text not null default '',
  category text not null default '刀片',
  unit text not null default '片',
  stock_quantity integer not null default 0 check (stock_quantity >= 0),
  unit_price numeric(14, 2) not null default 0 check (unit_price >= 0),
  purchased_at date,
  notes text not null default '',
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  constraint products_user_code_unique unique (user_id, code),
  constraint products_specification_not_blank check (btrim(specification) <> ''),
  constraint products_name_not_blank check (btrim(name) <> ''),
  constraint products_unit_not_blank check (btrim(unit) <> '')
);

create table if not exists public.stock_movements (
  id uuid primary key default gen_random_uuid(),
  user_id uuid references auth.users(id) on delete cascade,
  product_id uuid not null references public.products(id) on delete cascade,
  movement_type text not null check (movement_type in ('in', 'out')),
  quantity integer not null check (quantity > 0),
  movement_date date not null default current_date,
  unit_price numeric(14, 2) check (unit_price is null or unit_price >= 0),
  supplier text not null default '',
  notes text not null default '',
  stock_after integer not null check (stock_after >= 0),
  request_key uuid not null,
  created_at timestamptz not null default now(),
  constraint stock_movements_request_unique unique (user_id, request_key)
);

create table if not exists public.production_records (
  id uuid primary key default gen_random_uuid(),
  user_id uuid references auth.users(id) on delete cascade,
  work_date date not null default current_date,
  product_code text not null check (btrim(product_code) <> ''),
  machine text not null check (btrim(machine) <> ''),
  hours integer not null default 0 check (hours >= 0),
  minutes integer not null default 0 check (minutes between 0 and 59),
  total_minutes integer generated always as (hours * 60 + minutes) stored,
  quantity integer not null check (quantity > 0),
  notes text not null default '',
  created_at timestamptz not null default now(),
  constraint production_time_positive check (hours * 60 + minutes > 0)
);

create table if not exists public.production_options (
  id uuid primary key default gen_random_uuid(),
  user_id uuid references auth.users(id) on delete cascade,
  option_type text not null check (option_type in ('product', 'machine')),
  value text not null check (btrim(value) <> ''),
  created_at timestamptz not null default now()
);

create index if not exists products_user_id_idx on public.products(user_id);
create index if not exists products_user_name_idx on public.products(user_id, name);
create index if not exists stock_movements_user_date_idx on public.stock_movements(user_id, movement_date desc, created_at desc);
create index if not exists stock_movements_product_idx on public.stock_movements(product_id);
create index if not exists production_records_product_date_idx on public.production_records(product_code, work_date desc, created_at desc);
create unique index if not exists production_options_shared_value_unique
  on public.production_options(option_type, lower(value)) where user_id is null;

-- 免登入共用模式：所有網站訪客使用 user_id = null 的同一套資料。
alter table public.products alter column user_id drop not null;
alter table public.stock_movements alter column user_id drop not null;
alter table public.production_records alter column user_id drop not null;
alter table public.production_options alter column user_id drop not null;
update public.stock_movements set user_id = null where user_id is not null;
update public.products set user_id = null where user_id is not null;
update public.production_records set user_id = null where user_id is not null;
update public.production_options set user_id = null where user_id is not null;

-- 將既有加工紀錄中的產品編號與機台補進下拉選單資料。
insert into public.production_options (user_id, option_type, value)
select null, 'product', source.value
from (
  select distinct upper(btrim(product_code)) as value
  from public.production_records
  where btrim(product_code) <> ''
) as source
where not exists (
  select 1 from public.production_options as option
  where option.user_id is null and option.option_type = 'product' and lower(option.value) = lower(source.value)
);

insert into public.production_options (user_id, option_type, value)
select null, 'machine', source.value
from (
  select distinct upper(btrim(machine)) as value
  from public.production_records
  where btrim(machine) <> ''
) as source
where not exists (
    select 1 from public.production_options as option
    where option.user_id is null and option.option_type = 'machine' and lower(option.value) = lower(source.value)
  );

-- 將舊版自由輸入分類遷移為四種固定商品類別，再加上資料庫檢查。
update public.products
set category = '刀片'
where category not in ('刀片', '鑽銑刀', '刀柄', '其他');

alter table public.products alter column category set default '刀片';

do $$
begin
  if not exists (
    select 1 from pg_constraint
    where conname = 'products_category_valid' and conrelid = 'public.products'::regclass
  ) then
    alter table public.products
      add constraint products_category_valid check (category in ('刀片', '鑽銑刀', '刀柄', '其他'));
  end if;
end;
$$;

-- 四種分類使用獨立前綴與流水號：刀片 T、鑽銑刀 M、刀柄 H、其他 O。
-- 先暫存為 UUID 形式避免既有編號在重新編排時撞到 unique constraint。
update public.products set code = 'TMP-' || id::text;
with numbered as (
  select id,
         case category when '刀片' then 'T' when '鑽銑刀' then 'M' when '刀柄' then 'H' else 'O' end as prefix,
         row_number() over (partition by user_id, category order by created_at, id) as sequence_number
  from public.products
)
update public.products as product
set code = numbered.prefix || lpad(numbered.sequence_number::text, 4, '0')
from numbered
where product.id = numbered.id;

create unique index if not exists products_shared_code_unique
  on public.products(code) where user_id is null;
create unique index if not exists stock_movements_shared_request_unique
  on public.stock_movements(request_key) where user_id is null;

create or replace function public.set_updated_at()
returns trigger
language plpgsql
set search_path = public, pg_temp
as $$
begin
  new.updated_at := now();
  return new;
end;
$$;

drop trigger if exists products_set_updated_at on public.products;
create trigger products_set_updated_at
before update on public.products
for each row execute function public.set_updated_at();

create or replace function public.sync_product_name()
returns trigger
language plpgsql
set search_path = public, pg_temp
as $$
begin
  new.specification := upper(btrim(new.specification));
  new.size := upper(btrim(coalesce(new.size, '')));
  new.chipbreaker := upper(btrim(coalesce(new.chipbreaker, '')));
  new.material := upper(btrim(coalesce(new.material, '')));
  new.size := case when new.category in ('刀片', '鑽銑刀') then new.size else '' end;
  new.chipbreaker := case when new.category = '刀片' then new.chipbreaker else '' end;
  new.material := case when new.category in ('刀片', '鑽銑刀') then new.material else '' end;
  new.name := case
    when new.category = '刀片' then concat_ws('-', nullif(new.specification, ''), nullif(new.size, ''), nullif(new.chipbreaker, ''), nullif(new.material, ''))
    when new.category = '鑽銑刀' then concat_ws('-', nullif(new.specification, ''), nullif(new.size, ''), nullif(new.material, ''))
    else new.specification
  end;
  return new;
end;
$$;

drop trigger if exists products_sync_name on public.products;
create trigger products_sync_name
before insert or update of category, specification, size, chipbreaker, material on public.products
for each row execute function public.sync_product_name();

alter table public.products enable row level security;
alter table public.stock_movements enable row level security;
alter table public.production_records enable row level security;
alter table public.production_options enable row level security;
alter table public.products force row level security;
alter table public.stock_movements force row level security;
alter table public.production_records force row level security;
alter table public.production_options force row level security;

drop policy if exists "products_select_own" on public.products;
drop policy if exists "products_shared_select" on public.products;
create policy "products_shared_select" on public.products
for select to anon, authenticated
using (user_id is null);

drop policy if exists "products_update_own" on public.products;
drop policy if exists "products_shared_update" on public.products;
create policy "products_shared_update" on public.products
for update to anon, authenticated
using (user_id is null)
with check (user_id is null);

drop policy if exists "products_delete_own" on public.products;
drop policy if exists "products_shared_delete" on public.products;
create policy "products_shared_delete" on public.products
for delete to anon, authenticated
using (user_id is null);

drop policy if exists "movements_select_own" on public.stock_movements;
drop policy if exists "movements_shared_select" on public.stock_movements;
create policy "movements_shared_select" on public.stock_movements
for select to anon, authenticated
using (user_id is null);

drop policy if exists "production_records_shared_select" on public.production_records;
create policy "production_records_shared_select" on public.production_records
for select to anon, authenticated
using (user_id is null);

drop policy if exists "production_records_shared_insert" on public.production_records;
create policy "production_records_shared_insert" on public.production_records
for insert to anon, authenticated
with check (user_id is null);

drop policy if exists "production_records_shared_delete" on public.production_records;
create policy "production_records_shared_delete" on public.production_records
for delete to anon, authenticated
using (user_id is null);

drop policy if exists "production_options_shared_select" on public.production_options;
create policy "production_options_shared_select" on public.production_options
for select to anon, authenticated
using (user_id is null);

drop policy if exists "production_options_shared_insert" on public.production_options;
create policy "production_options_shared_insert" on public.production_options
for insert to anon, authenticated
with check (user_id is null);

drop policy if exists "production_options_shared_delete" on public.production_options;
create policy "production_options_shared_delete" on public.production_options
for delete to anon, authenticated
using (user_id is null);

-- 安全建立商品及選填的第一筆入庫。商品編號在每位使用者範圍內依序產生。
create or replace function public.create_product_with_initial_stock(
  p_specification text,
  p_size text default '',
  p_chipbreaker text default null,
  p_material text default null,
  p_brand text default null,
  p_supplier text default null,
  p_category text default '刀片',
  p_unit text default '片',
  p_notes text default null,
  p_initial_quantity integer default 0,
  p_initial_unit_price numeric default 0,
  p_purchased_at date default current_date,
  p_request_key uuid default gen_random_uuid()
)
returns uuid
language plpgsql
security definer
set search_path = public, pg_temp
as $$
declare
  v_user_id uuid := null;
  v_product_id uuid;
  v_code text;
  v_prefix text;
  v_next_number integer;
begin
  if p_specification is null or btrim(p_specification) = '' then
    raise exception '規格為必填欄位' using errcode = '22023';
  end if;
  if p_category not in ('刀片', '鑽銑刀', '刀柄', '其他') then
    raise exception '商品分類不正確' using errcode = '22023';
  end if;
  if p_category in ('刀片', '鑽銑刀') and (p_size is null or btrim(p_size) = '') then
    raise exception '此分類需要輸入尺寸' using errcode = '22023';
  end if;
  if p_category in ('刀片', '鑽銑刀') and (p_material is null or btrim(p_material) = '') then
    raise exception '此分類需要輸入材質' using errcode = '22023';
  end if;
  if p_initial_quantity < 0 then
    raise exception '期初庫存不能小於 0' using errcode = '22023';
  end if;
  if p_initial_unit_price < 0 then
    raise exception '單價不能小於 0' using errcode = '22023';
  end if;

  -- 同一帳戶同時新增商品時，也只會取到不同的下一號。
  perform pg_advisory_xact_lock(hashtextextended('shared-inventory', 73001));
  v_prefix := case p_category when '刀片' then 'T' when '鑽銑刀' then 'M' when '刀柄' then 'H' else 'O' end;
  select coalesce(max(substring(code from 2)::integer), 0) + 1
  into v_next_number
  from public.products
  where user_id is not distinct from v_user_id and category = p_category and code ~ ('^' || v_prefix || '[0-9]+$');
  v_code := v_prefix || lpad(v_next_number::text, 4, '0');

  insert into public.products (
    user_id, code, name, specification, size, chipbreaker, material, brand,
    supplier, category, unit, stock_quantity, unit_price, purchased_at, notes
  ) values (
    v_user_id,
    v_code,
    upper(btrim(p_specification)),
    upper(btrim(p_specification)),
    case when p_category in ('刀片', '鑽銑刀') then upper(btrim(coalesce(p_size, ''))) else '' end,
    case when p_category = '刀片' then upper(btrim(coalesce(p_chipbreaker, ''))) else '' end,
    case when p_category in ('刀片', '鑽銑刀') then upper(btrim(coalesce(p_material, ''))) else '' end,
    btrim(coalesce(p_brand, '')),
    btrim(coalesce(p_supplier, '')),
    p_category,
    coalesce(nullif(btrim(p_unit), ''), '片'),
    p_initial_quantity,
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

-- 安全刪除一筆出入庫並重建該商品的完整庫存軌跡。
create or replace function public.delete_stock_movement(p_movement_id uuid)
returns void
language plpgsql
security definer
set search_path = public, pg_temp
as $$
declare
  v_user_id uuid := null;
  v_product_id uuid;
  v_min_stock integer;
  v_final_stock integer := 0;
  v_latest_price numeric(14, 2);
  v_latest_date date;
  v_latest_supplier text;
begin
  select product_id into v_product_id
  from public.stock_movements
  where id = p_movement_id and user_id is not distinct from v_user_id;
  if not found then
    raise exception '找不到出入庫紀錄或沒有權限' using errcode = '42501';
  end if;

  -- 鎖定商品，避免刪除與另一筆出入庫同時改動庫存。
  perform 1 from public.products
  where id = v_product_id and user_id is not distinct from v_user_id
  for update;
  if not found then
    raise exception '找不到商品或沒有權限' using errcode = '42501';
  end if;

  delete from public.stock_movements
  where id = p_movement_id and user_id is not distinct from v_user_id;

  with balances as (
    select sum(case when movement_type = 'in' then quantity else -quantity end)
      over (order by movement_date, created_at, id) as balance
    from public.stock_movements
    where product_id = v_product_id and user_id is not distinct from v_user_id
  )
  select coalesce(min(balance), 0) into v_min_stock from balances;

  if v_min_stock < 0 then
    raise exception '無法刪除：這會使後續某筆歷史庫存變成負數' using errcode = '23514';
  end if;

  with balances as (
    select id,
           sum(case when movement_type = 'in' then quantity else -quantity end)
             over (order by movement_date, created_at, id) as balance
    from public.stock_movements
    where product_id = v_product_id and user_id is not distinct from v_user_id
  )
  update public.stock_movements as movement
  set stock_after = balances.balance
  from balances
  where movement.id = balances.id;

  select coalesce(sum(case when movement_type = 'in' then quantity else -quantity end), 0)
  into v_final_stock
  from public.stock_movements
  where product_id = v_product_id and user_id is not distinct from v_user_id;

  select unit_price, movement_date, supplier
  into v_latest_price, v_latest_date, v_latest_supplier
  from public.stock_movements
  where product_id = v_product_id and user_id is not distinct from v_user_id and movement_type = 'in'
  order by movement_date desc, created_at desc, id desc
  limit 1;

  update public.products
  set stock_quantity = v_final_stock,
      unit_price = coalesce(v_latest_price, 0),
      purchased_at = v_latest_date,
      supplier = case when nullif(btrim(coalesce(v_latest_supplier, '')), '') is not null then v_latest_supplier else supplier end
  where id = v_product_id and user_id is not distinct from v_user_id;
end;
$$;

-- 安全出入庫：資料列鎖避免競爭、request_key 防止重送，並在同一交易中更新庫存與歷史。
create or replace function public.register_stock_movement(
  p_product_id uuid,
  p_movement_type text,
  p_quantity integer,
  p_movement_date date default current_date,
  p_unit_price numeric default null,
  p_supplier text default null,
  p_notes text default null,
  p_request_key uuid default gen_random_uuid()
)
returns uuid
language plpgsql
security definer
set search_path = public, pg_temp
as $$
declare
  v_user_id uuid := null;
  v_product public.products%rowtype;
  v_next_stock integer;
  v_movement_id uuid;
begin
  if p_movement_type not in ('in', 'out') then
    raise exception '出入庫類型必須是 in 或 out' using errcode = '22023';
  end if;
  if p_quantity is null or p_quantity <= 0 then
    raise exception '數量必須大於 0' using errcode = '22023';
  end if;
  if p_movement_type = 'in' and coalesce(p_unit_price, 0) < 0 then
    raise exception '單價不能小於 0' using errcode = '22023';
  end if;

  select id into v_movement_id
  from public.stock_movements
  where user_id is not distinct from v_user_id and request_key = p_request_key;
  if v_movement_id is not null then
    return v_movement_id;
  end if;

  select * into v_product
  from public.products
  where id = p_product_id and user_id is not distinct from v_user_id
  for update;
  if not found then
    raise exception '找不到商品或沒有權限' using errcode = '42501';
  end if;

  -- 取得鎖後再檢查一次，涵蓋同商品的同時重送。
  select id into v_movement_id
  from public.stock_movements
  where user_id is not distinct from v_user_id and request_key = p_request_key;
  if v_movement_id is not null then
    return v_movement_id;
  end if;

  v_next_stock := v_product.stock_quantity + case when p_movement_type = 'in' then p_quantity else -p_quantity end;
  if v_next_stock < 0 then
    raise exception '庫存不足，目前只有 % %', v_product.stock_quantity, v_product.unit using errcode = '23514';
  end if;

  update public.products
  set stock_quantity = v_next_stock,
      unit_price = case when p_movement_type = 'in' then coalesce(p_unit_price, 0) else unit_price end,
      supplier = case when p_movement_type = 'in' and nullif(btrim(coalesce(p_supplier, '')), '') is not null then btrim(p_supplier) else supplier end,
      purchased_at = case when p_movement_type = 'in' then coalesce(p_movement_date, current_date) else purchased_at end
  where id = p_product_id;

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

-- 庫存盤點：不要求新價格，鎖定刀具後以差異建立調整紀錄並保留原單價。
create or replace function public.set_stocktake_quantity(
  p_product_id uuid,
  p_actual_quantity integer,
  p_notes text default null,
  p_request_key uuid default gen_random_uuid()
)
returns uuid
language plpgsql
security definer
set search_path = public, pg_temp
as $$
declare
  v_user_id uuid := null;
  v_product public.products%rowtype;
  v_difference integer;
  v_movement_id uuid;
  v_note text;
begin
  if p_actual_quantity is null or p_actual_quantity < 0 then
    raise exception '實際盤點數量必須是 0 以上的整數' using errcode = '22023';
  end if;

  select id into v_movement_id
  from public.stock_movements
  where user_id is not distinct from v_user_id and request_key = p_request_key;
  if v_movement_id is not null then
    return v_movement_id;
  end if;

  select * into v_product
  from public.products
  where id = p_product_id and user_id is not distinct from v_user_id
  for update;
  if not found then
    raise exception '找不到刀具或沒有權限' using errcode = '42501';
  end if;

  select id into v_movement_id
  from public.stock_movements
  where user_id is not distinct from v_user_id and request_key = p_request_key;
  if v_movement_id is not null then
    return v_movement_id;
  end if;

  v_difference := p_actual_quantity - v_product.stock_quantity;
  if v_difference = 0 then
    return null;
  end if;

  v_note := format('庫存盤點調整：%s → %s', v_product.stock_quantity, p_actual_quantity);
  if nullif(btrim(coalesce(p_notes, '')), '') is not null then
    v_note := v_note || '｜' || btrim(p_notes);
  end if;

  insert into public.stock_movements (
    user_id, product_id, movement_type, quantity, movement_date, unit_price,
    supplier, notes, stock_after, request_key
  ) values (
    v_user_id, v_product.id, case when v_difference > 0 then 'in' else 'out' end,
    abs(v_difference), current_date,
    case when v_difference > 0 then v_product.unit_price else null end,
    case when v_difference > 0 then v_product.supplier else '' end,
    v_note, p_actual_quantity, p_request_key
  ) returning id into v_movement_id;

  update public.products
  set stock_quantity = p_actual_quantity
  where id = v_product.id and user_id is not distinct from v_user_id;

  return v_movement_id;
end;
$$;

-- 永久刪除一個加工產品：加工紀錄與「管理產品編號」選項在同一交易內一起移除。
create or replace function public.delete_production_product_data(p_product_code text)
returns integer
language plpgsql
security definer
set search_path = public, pg_temp
as $$
declare
  v_product_code text := upper(btrim(coalesce(p_product_code, '')));
  v_deleted_count integer := 0;
begin
  if v_product_code = '' then
    raise exception '產品編號不能空白' using errcode = '22023';
  end if;

  delete from public.production_records
  where user_id is null and upper(product_code) = v_product_code;
  get diagnostics v_deleted_count = row_count;

  delete from public.production_options
  where user_id is null and option_type = 'product' and upper(value) = v_product_code;

  return v_deleted_count;
end;
$$;

revoke all on table public.products from anon, authenticated;
revoke all on table public.stock_movements from anon, authenticated;
grant select, delete on table public.products to anon, authenticated;
grant update (specification, size, chipbreaker, material, name, brand, supplier, category, unit, notes) on table public.products to anon, authenticated;
grant select on table public.stock_movements to anon, authenticated;
revoke all on table public.production_records from anon, authenticated;
grant select, insert, delete on table public.production_records to anon, authenticated;
revoke all on table public.production_options from anon, authenticated;
grant select, insert, delete on table public.production_options to anon, authenticated;

revoke all on function public.create_product_with_initial_stock(text, text, text, text, text, text, text, text, text, integer, numeric, date, uuid) from public, anon;
revoke all on function public.register_stock_movement(uuid, text, integer, date, numeric, text, text, uuid) from public, anon;
revoke all on function public.delete_stock_movement(uuid) from public, anon;
revoke all on function public.set_stocktake_quantity(uuid, integer, text, uuid) from public, anon;
revoke all on function public.delete_production_product_data(text) from public, anon;
grant execute on function public.create_product_with_initial_stock(text, text, text, text, text, text, text, text, text, integer, numeric, date, uuid) to anon, authenticated;
grant execute on function public.register_stock_movement(uuid, text, integer, date, numeric, text, text, uuid) to anon, authenticated;
grant execute on function public.delete_stock_movement(uuid) to anon, authenticated;
grant execute on function public.set_stocktake_quantity(uuid, integer, text, uuid) to anon, authenticated;
grant execute on function public.delete_production_product_data(text) to anon, authenticated;

comment on function public.register_stock_movement is '以資料列鎖、負庫存檢查及防重送鍵原子化處理出入庫。';
comment on function public.delete_stock_movement is '安全刪除出入庫紀錄，拒絕負庫存並重算商品庫存、歷史餘額及最新單價。';
comment on function public.set_stocktake_quantity is '以盤點實際數量建立差異紀錄並更新庫存，不要求或覆寫單價。';
comment on function public.delete_production_product_data is '在同一交易內永久刪除指定產品的加工紀錄與產品編號選項。';
comment on column public.stock_movements.request_key is '由前端每次送出時產生，重試相同請求不會重複扣加庫存。';
comment on table public.production_records is '免登入共用的加工數量、機台與加工時間紀錄。';
comment on table public.production_options is '加工數量登記使用的產品編號與機台下拉選單資料。';
