-- 詮栩 CNC 雲端系統：啟用刀具條碼欄位與完整備份
-- 請在 Supabase SQL Editor 建立 New query，貼上整份內容並按 Run。
-- 可重複執行；不會刪除既有刀具、出入庫或加工資料。

begin;

alter table public.products add column if not exists barcode text not null default '';
update public.products set barcode = btrim(barcode) where barcode <> btrim(barcode);

create unique index if not exists products_barcode_unique
  on public.products(lower(barcode))
  where btrim(barcode) <> '';

grant update (barcode) on table public.products to authenticated;

-- Excel 完整備份加入條碼欄位。
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
        'code', p.code, 'barcode', p.barcode, 'category', p.category, 'specification', p.specification,
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

-- Excel 完整還原同時還原條碼；舊版沒有條碼欄位的 Excel 仍可匯入。
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
      user_id, code, barcode, category, specification, size, chipbreaker, material, brand,
      supplier, unit, stock_quantity, unit_price, purchased_at, notes, name
    ) values (
      v_user_id, v_code, btrim(coalesce(v_item->>'barcode', '')), v_item->>'category', btrim(coalesce(v_item->>'specification', '')),
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

revoke all on function public.admin_export_all() from public, anon, authenticated;
revoke all on function public.admin_import_all(jsonb) from public, anon, authenticated;
grant execute on function public.admin_export_all() to authenticated;
grant execute on function public.admin_import_all(jsonb) to authenticated;

comment on column public.products.barcode is '手機相機掃描用的刀具條碼；空白代表尚未綁定。';

commit;

select column_name, data_type, is_nullable
from information_schema.columns
where table_schema = 'public' and table_name = 'products' and column_name = 'barcode';
