-- 加工管理系統：安全修改產品編號。
-- 執行一次即可。修改會在同一筆交易內同步更新「管理產品編號」與全部既有加工紀錄。

begin;

create or replace function public.rename_production_product_code(
  p_current_value text,
  p_new_value text
)
returns integer
language plpgsql
security definer
set search_path = public, pg_temp
as $$
declare
  v_user_id uuid := auth.uid();
  v_current_value text := upper(btrim(coalesce(p_current_value, '')));
  v_new_value text := upper(btrim(coalesce(p_new_value, '')));
  v_option_count integer := 0;
  v_record_count integer := 0;
begin
  if v_user_id is distinct from '5568706d-caaa-459b-b842-13cdbc3929fc'::uuid then
    raise exception '只有指定管理員帳號可以修改產品編號' using errcode = '42501';
  end if;

  if v_current_value = '' or v_new_value = '' then
    raise exception '產品編號不能空白' using errcode = '22023';
  end if;

  if v_current_value = v_new_value then
    return 0;
  end if;

  -- 避免兩台裝置同時修改產品編號造成重複或資料合併。
  perform pg_advisory_xact_lock(hashtextextended('production-product-code-rename', 73003));

  if exists (
    select 1
    from public.production_options
    where user_id = v_user_id
      and option_type = 'product'
      and upper(btrim(value)) = v_new_value
  ) or exists (
    select 1
    from public.production_records
    where user_id = v_user_id
      and upper(btrim(product_code)) = v_new_value
  ) then
    raise exception '新的產品編號已經存在' using errcode = '23505';
  end if;

  update public.production_options
  set value = v_new_value
  where user_id = v_user_id
    and option_type = 'product'
    and upper(btrim(value)) = v_current_value;
  get diagnostics v_option_count = row_count;

  if v_option_count = 0 then
    raise exception '找不到原本的產品編號' using errcode = 'P0002';
  end if;

  update public.production_records
  set product_code = v_new_value
  where user_id = v_user_id
    and upper(btrim(product_code)) = v_current_value;
  get diagnostics v_record_count = row_count;

  return v_record_count;
end;
$$;

revoke all on function public.rename_production_product_code(text, text) from public, anon, authenticated;
grant execute on function public.rename_production_product_code(text, text) to authenticated;

comment on function public.rename_production_product_code(text, text) is
  '登入管理員可在同一交易內修改產品編號選項，並同步取代全部既有加工紀錄中的原編號。';

commit;
