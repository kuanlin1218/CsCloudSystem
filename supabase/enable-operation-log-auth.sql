-- 先套用 enable-operation-log.sql，並在 Supabase Auth 開啟將稽核紀錄寫入資料庫。
begin;

create or replace function public.read_operation_timeline(
  p_before_time timestamptz default null,
  p_before_id text default null
)
returns table (
  id text,
  occurred_at timestamptz,
  actor_id uuid,
  entity text,
  operation text,
  target text,
  before_data jsonb,
  after_data jsonb
)
language plpgsql stable security definer set search_path = ''
as $$
declare
  v_owner uuid := private.require_app_user();
begin
  return query
  select timeline.*
  from (
    select 'data:' || lpad(entry.id::text, 19, '0') as id,
      entry.occurred_at, entry.actor_id, entry.entity, entry.operation,
      entry.target, entry.before_data, entry.after_data
    from public.operation_log as entry
    union all
    select 'auth:' || entry.id::text, entry.created_at, v_owner,
      'auth'::text, upper(entry.payload->>'action'), '系統帳號'::text,
      null::jsonb, jsonb_build_object('auth_event_id', entry.id)
    from auth.audit_log_entries as entry
    where entry.created_at is not null
      and coalesce(entry.payload->>'actor_id', entry.payload->>'user_id') = v_owner::text
      and entry.payload->>'action' in ('login', 'logout')
  ) as timeline
  where p_before_time is null
    or (timeline.occurred_at, timeline.id) < (p_before_time, p_before_id)
  order by timeline.occurred_at desc, timeline.id desc
  limit 51;
end;
$$;

revoke all on function public.read_operation_timeline(timestamptz, text) from public, anon, authenticated;
grant execute on function public.read_operation_timeline(timestamptz, text) to authenticated;

commit;
