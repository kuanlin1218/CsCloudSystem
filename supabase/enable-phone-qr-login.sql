begin;

-- 手機掃碼登入只經由 Edge Function 存取，訪客與一般登入工作階段皆無權直接讀寫。
create table if not exists public.phone_qr_login_requests (
  desktop_hash text primary key check (desktop_hash ~ '^[0-9a-f]{64}$'),
  phone_hash text not null unique check (phone_hash ~ '^[0-9a-f]{64}$'),
  confirmation_hash text not null check (confirmation_hash ~ '^[0-9a-f]{64}$'),
  expires_at timestamptz not null,
  approved_user_id uuid
);

alter table public.phone_qr_login_requests enable row level security;
revoke all on table public.phone_qr_login_requests from public, anon, authenticated;
grant select, insert, update, delete on table public.phone_qr_login_requests to service_role;

-- 建立新請求時清除逾期資料，限制目前有效的請求數，避免公開入口累積無限資料。
create or replace function public.create_phone_qr_login_request(
  p_desktop_hash text,
  p_phone_hash text,
  p_confirmation_hash text
)
returns timestamptz
language plpgsql
security invoker
set search_path = ''
as $$
declare
  cutoff timestamptz;
begin
  perform pg_catalog.pg_advisory_xact_lock(1096338249, 593721);
  delete from public.phone_qr_login_requests where expires_at <= pg_catalog.clock_timestamp();

  if (select count(*) from public.phone_qr_login_requests) >= 60 then
    return null;
  end if;

  cutoff := pg_catalog.clock_timestamp() + interval '3 minutes';
  insert into public.phone_qr_login_requests (desktop_hash, phone_hash, confirmation_hash, expires_at)
  values (p_desktop_hash, p_phone_hash, p_confirmation_hash, cutoff);
  return cutoff;
end;
$$;

-- 只允許原分頁持有的憑證兌換。DELETE RETURNING 將領取與刪除合併成一次操作。
create or replace function public.redeem_phone_qr_login_request(p_desktop_hash text, p_owner_id uuid)
returns uuid
language sql
security invoker
set search_path = ''
as $$
  delete from public.phone_qr_login_requests
  where desktop_hash = p_desktop_hash
    and approved_user_id = p_owner_id
    and expires_at > pg_catalog.clock_timestamp()
  returning approved_user_id;
$$;

revoke execute on function public.create_phone_qr_login_request(text, text, text) from public, anon, authenticated;
revoke execute on function public.redeem_phone_qr_login_request(text, uuid) from public, anon, authenticated;
grant execute on function public.create_phone_qr_login_request(text, text, text) to service_role;
grant execute on function public.redeem_phone_qr_login_request(text, uuid) to service_role;

commit;
