-- 在啟用單一管理員及 enable-production-orders.sql 後套用機台管理。
begin;

do $$
begin
  if to_regprocedure('private.admin_import_all_before_orders(jsonb)') is null then
    raise exception '請先套用 enable-production-orders.sql';
  end if;
end;
$$;

alter table public.production_options add column if not exists machine_active boolean not null default true;

create table if not exists public.machine_profiles (
  machine_id uuid primary key references public.production_options(id) on delete cascade,
  user_id uuid not null references auth.users(id),
  name text not null default '',
  brand text not null default '',
  model text not null default '',
  serial_number text not null default '',
  purchased_on date,
  started_on date,
  notes text not null default ''
);

create table if not exists public.machine_service_records (
  id uuid primary key default gen_random_uuid(),
  machine_id uuid not null references public.production_options(id) on delete restrict,
  user_id uuid not null references auth.users(id),
  service_date date not null,
  service_type text not null check (service_type in ('maintenance', 'repair')),
  description text not null check (btrim(description) <> ''),
  technician text not null default '',
  cost numeric(14, 2) check (cost >= 0 and cost <> 'NaN'::numeric),
  downtime_minutes integer check (downtime_minutes >= 0),
  created_at timestamptz not null default now()
);

create index if not exists machine_service_records_history_idx
  on public.machine_service_records (machine_id, service_date desc, created_at desc, id desc);

alter table public.machine_profiles enable row level security;
alter table public.machine_service_records enable row level security;
revoke all on public.machine_profiles, public.machine_service_records from public, anon, authenticated;
grant select on public.machine_profiles, public.machine_service_records to authenticated;

drop policy if exists machine_profiles_owner_read on public.machine_profiles;
create policy machine_profiles_owner_read on public.machine_profiles for select to authenticated
  using (user_id = (select auth.uid()));
drop policy if exists machine_service_records_owner_read on public.machine_service_records;
create policy machine_service_records_owner_read on public.machine_service_records for select to authenticated
  using (user_id = (select auth.uid()));

create or replace function public.save_machine_profile(
  p_machine_id uuid, p_current_code text, p_code text, p_active boolean, p_profile jsonb
)
returns void language plpgsql security definer set search_path = ''
as $$
declare
  v_user_id uuid := private.require_app_user();
  v_current text;
  v_code text := upper(btrim(coalesce(p_code, '')));
  v_profile public.machine_profiles := jsonb_populate_record(null::public.machine_profiles, p_profile);
begin
  if v_code = '' or p_active is null or jsonb_typeof(p_profile) is distinct from 'object' then
    raise exception '請輸入有效的機台編號、狀態與基本資料';
  end if;
  perform pg_advisory_xact_lock(hashtextextended('production-machine-profile', 73003));
  select value into v_current from public.production_options
    where id = p_machine_id and user_id = v_user_id and option_type = 'machine' for update;
  if not found or v_current is distinct from p_current_code then
    raise exception '機台編號已變更或不存在，請重新開啟。';
  end if;
  if upper(v_current) <> v_code and (
    exists (select 1 from public.production_options where user_id = v_user_id and option_type = 'machine' and upper(value) = v_code)
    or exists (select 1 from public.production_records where user_id = v_user_id and upper(machine) = v_code)
  ) then
    raise exception '這個機台編號已在使用或已有加工紀錄';
  end if;

  update public.production_options set value = v_code, machine_active = p_active where id = p_machine_id;
  if v_current <> v_code then
    update public.production_records set machine = v_code where user_id = v_user_id and upper(machine) = upper(v_current);
  end if;

  insert into public.machine_profiles (machine_id, user_id, name, brand, model, serial_number, purchased_on, started_on, notes)
  values (p_machine_id, v_user_id, btrim(coalesce(v_profile.name, '')), btrim(coalesce(v_profile.brand, '')),
    btrim(coalesce(v_profile.model, '')), btrim(coalesce(v_profile.serial_number, '')), v_profile.purchased_on,
    v_profile.started_on, btrim(coalesce(v_profile.notes, '')))
  on conflict (machine_id) do update set name = excluded.name, brand = excluded.brand, model = excluded.model,
    serial_number = excluded.serial_number, purchased_on = excluded.purchased_on, started_on = excluded.started_on, notes = excluded.notes;
end;
$$;

create or replace function public.save_machine_service_record(p_machine_id uuid, p_record_id uuid, p_record jsonb)
returns jsonb language plpgsql security definer set search_path = ''
as $$
declare
  v_user_id uuid := private.require_app_user();
  v_input public.machine_service_records := jsonb_populate_record(null::public.machine_service_records, p_record);
  v_saved public.machine_service_records;
begin
  perform pg_advisory_xact_lock(hashtextextended('production-machine-profile', 73003));
  perform 1 from public.production_options where id = p_machine_id and user_id = v_user_id and option_type = 'machine' for key share;
  if not found then raise exception '找不到這台機台'; end if;
  if v_input.service_date is null or v_input.service_date > (now() at time zone 'Asia/Taipei')::date then
    raise exception '請輸入不晚於今天的日期';
  end if;
  if p_record_id is null then
    insert into public.machine_service_records (machine_id, user_id, service_date, service_type, description, technician, cost, downtime_minutes)
      values (p_machine_id, v_user_id, v_input.service_date, v_input.service_type, btrim(v_input.description),
        btrim(coalesce(v_input.technician, '')), v_input.cost, v_input.downtime_minutes)
      returning * into v_saved;
  else
    update public.machine_service_records set service_date = v_input.service_date, service_type = v_input.service_type,
      description = btrim(v_input.description), technician = btrim(coalesce(v_input.technician, '')),
      cost = v_input.cost, downtime_minutes = v_input.downtime_minutes
      where id = p_record_id and machine_id = p_machine_id and user_id = v_user_id returning * into v_saved;
    if not found then raise exception '找不到這筆保養維修紀錄'; end if;
  end if;
  return to_jsonb(v_saved) - 'user_id';
end;
$$;

-- 同一份 Excel 的舊有匯出、還原與操作紀錄流程仍沿用原 RPC。
do $$
begin
  if to_regprocedure('private.admin_export_all_before_machines()') is null then
    alter function public.admin_export_all() set schema private;
    alter function private.admin_export_all() rename to admin_export_all_before_machines;
  end if;
  if to_regprocedure('private.admin_import_all_before_machines(jsonb)') is null then
    alter function public.admin_import_all(jsonb) set schema private;
    alter function private.admin_import_all(jsonb) rename to admin_import_all_before_machines;
  end if;
end;
$$;

create or replace function public.admin_export_all()
returns jsonb language plpgsql security definer set search_path = ''
as $$
declare
  v_user_id uuid := private.require_app_user();
  v_export jsonb := private.admin_export_all_before_machines();
  v_profiles jsonb;
  v_records jsonb;
begin
  select coalesce(jsonb_agg(jsonb_build_object(
    'machine_code', o.value, 'machine_active', o.machine_active, 'name', p.name, 'brand', p.brand, 'model', p.model,
    'serial_number', p.serial_number, 'purchased_on', p.purchased_on, 'started_on', p.started_on, 'notes', p.notes
  ) order by o.value), '[]'::jsonb)
    into v_profiles from public.production_options o left join public.machine_profiles p on p.machine_id = o.id and p.user_id = v_user_id
    where o.user_id = v_user_id and o.option_type = 'machine';
  select coalesce(jsonb_agg((to_jsonb(r) - array['id','machine_id','user_id','created_at'])
    || jsonb_build_object('machine_code', o.value) order by o.value, r.service_date, r.created_at, r.id), '[]'::jsonb)
    into v_records from public.machine_service_records r join public.production_options o on o.id = r.machine_id and o.user_id = v_user_id
    where r.user_id = v_user_id;
  return jsonb_set(jsonb_set(v_export, '{data,machine_profiles}', v_profiles), '{data,machine_service_records}', v_records);
end;
$$;

create or replace function public.admin_import_all(p_payload jsonb)
returns jsonb language plpgsql security definer set search_path = ''
as $$
declare
  v_user_id uuid := private.require_app_user();
  v_profiles jsonb := coalesce(p_payload->'machine_profiles', '[]'::jsonb);
  v_records jsonb := coalesce(p_payload->'machine_service_records', '[]'::jsonb);
  v_result jsonb;
  v_item jsonb;
  v_id uuid;
  v_code text;
begin
  if jsonb_typeof(v_profiles) is distinct from 'array' or jsonb_typeof(v_records) is distinct from 'array' then
    raise exception '機台備份資料格式不正確';
  end if;
  if jsonb_array_length(v_profiles) > 20000 or jsonb_array_length(v_records) > 100000 then
    raise exception '機台備份資料筆數超過安全上限';
  end if;
  perform pg_advisory_xact_lock(hashtextextended('production-machine-profile', 73003));
  -- 原還原會重新建立機台選項；保養維修依機台編號重新連結，整筆交易可回復。
  delete from public.machine_service_records where user_id = v_user_id;
  v_result := private.admin_import_all_before_machines(p_payload);

  for v_item in select value from jsonb_array_elements(v_profiles) loop
    v_code := upper(btrim(v_item->>'machine_code'));
    select id into v_id from public.production_options where user_id = v_user_id and option_type = 'machine' and value = v_code;
    if not found then raise exception '機台基本資料使用了不存在的機台「%」', v_code; end if;
    perform public.save_machine_profile(v_id, v_code, v_code, coalesce((v_item->>'machine_active')::boolean, true), v_item);
  end loop;
  for v_item in select value from jsonb_array_elements(v_records) loop
    v_code := upper(btrim(v_item->>'machine_code'));
    select id into v_id from public.production_options where user_id = v_user_id and option_type = 'machine' and value = v_code;
    if not found then raise exception '保養維修紀錄使用了不存在的機台「%」', v_code; end if;
    perform public.save_machine_service_record(v_id, null, v_item);
  end loop;
  return jsonb_set(jsonb_set(v_result, '{counts,machine_profiles}', to_jsonb(jsonb_array_length(v_profiles))),
    '{counts,machine_service_records}', to_jsonb(jsonb_array_length(v_records)));
end;
$$;

-- 含機台資料的檔案會先確認伺服器支援完整還原。
create or replace function public.admin_import_all_with_machine_data(p_payload jsonb)
returns jsonb language plpgsql security definer set search_path = ''
as $$
begin
  perform private.require_app_user();
  if to_regprocedure('public.admin_import_all_with_operation_log(jsonb)') is not null then
    return public.admin_import_all_with_operation_log(p_payload);
  end if;
  return public.admin_import_all(p_payload);
end;
$$;

-- 已啟用操作紀錄時記下機台編號與變更內容，整批還原仍只留一筆摘要。
create or replace function private.record_machine_operation_log()
returns trigger language plpgsql security definer set search_path = ''
as $$
declare
  v_before jsonb;
  v_after jsonb;
  v_row jsonb;
  v_target text;
begin
  if exists (select 1 from private.operation_log_restore where transaction_id = txid_current()) then return null; end if;
  if tg_op <> 'INSERT' then v_before := to_jsonb(old); end if;
  if tg_op <> 'DELETE' then v_after := to_jsonb(new); end if;
  v_row := coalesce(v_after, v_before);
  select value into v_target from public.production_options where id = (v_row->>'machine_id')::uuid;
  v_before := v_before - array['id','machine_id','user_id','created_at'];
  v_after := v_after - array['id','machine_id','user_id','created_at'];
  if v_before is distinct from v_after then
    insert into public.operation_log (actor_id, entity, operation, target, before_data, after_data)
      values (auth.uid(), tg_table_name, tg_op, coalesce(v_target, v_row->>'machine_id'), v_before, v_after);
  end if;
  return null;
end;
$$;

do $$
begin
  if to_regclass('public.operation_log') is not null and to_regclass('private.operation_log_restore') is not null then
    drop trigger if exists record_machine_operation_log on public.machine_profiles;
    create trigger record_machine_operation_log after insert or update or delete on public.machine_profiles
      for each row execute function private.record_machine_operation_log();
    drop trigger if exists record_machine_operation_log on public.machine_service_records;
    create trigger record_machine_operation_log after insert or update or delete on public.machine_service_records
      for each row execute function private.record_machine_operation_log();
  end if;
end;
$$;

revoke all on function private.admin_export_all_before_machines(), private.admin_import_all_before_machines(jsonb),
  private.record_machine_operation_log() from public, anon, authenticated;
revoke all on function public.save_machine_profile(uuid, text, text, boolean, jsonb),
  public.save_machine_service_record(uuid, uuid, jsonb), public.admin_export_all(), public.admin_import_all(jsonb), public.admin_import_all_with_machine_data(jsonb)
  from public, anon, authenticated;
grant execute on function public.save_machine_profile(uuid, text, text, boolean, jsonb),
  public.save_machine_service_record(uuid, uuid, jsonb), public.admin_export_all(), public.admin_import_all(jsonb), public.admin_import_all_with_machine_data(jsonb)
  to authenticated;

commit;
