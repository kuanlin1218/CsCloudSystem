-- 詮栩 CNC 雲端系統：啟用免登入訪客唯讀模式
-- 請在 Supabase SQL Editor 建立 New query，貼上整份內容並按 Run。
-- 此檔可重複執行；只開放 SELECT，不開放新增、修改、刪除或業務 RPC。

begin;

-- 四張前端資料表都必須維持 RLS，避免資料表權限繞過政策。
alter table public.products enable row level security;
alter table public.stock_movements enable row level security;
alter table public.production_records enable row level security;
alter table public.production_options enable row level security;
alter table public.products force row level security;
alter table public.stock_movements force row level security;
alter table public.production_records force row level security;
alter table public.production_options force row level security;

-- 訪客只能取得指定帳號所屬資料。
drop policy if exists products_guest_select on public.products;
create policy products_guest_select
on public.products for select to anon
using (user_id = '5568706d-caaa-459b-b842-13cdbc3929fc'::uuid);

drop policy if exists movements_guest_select on public.stock_movements;
create policy movements_guest_select
on public.stock_movements for select to anon
using (user_id = '5568706d-caaa-459b-b842-13cdbc3929fc'::uuid);

drop policy if exists production_records_guest_select on public.production_records;
create policy production_records_guest_select
on public.production_records for select to anon
using (user_id = '5568706d-caaa-459b-b842-13cdbc3929fc'::uuid);

drop policy if exists production_options_guest_select on public.production_options;
create policy production_options_guest_select
on public.production_options for select to anon
using (user_id = '5568706d-caaa-459b-b842-13cdbc3929fc'::uuid);

-- 權限層只給 SELECT；即使訪客繞過網站按鈕直接呼叫 API，也不能寫入資料。
grant usage on schema public to anon;
revoke all on table public.products from anon;
revoke all on table public.stock_movements from anon;
revoke all on table public.production_records from anon;
revoke all on table public.production_options from anon;
grant select on table public.products to anon;
grant select on table public.stock_movements to anon;
grant select on table public.production_records to anon;
grant select on table public.production_options to anon;

-- 訪客不需要任何 RPC；管理員既有的 authenticated 明確授權不受影響。
revoke execute on all functions in schema public from public, anon;

commit;

-- 成功後應列出四筆 anon SELECT policy。
select tablename, policyname, roles, cmd
from pg_policies
where schemaname = 'public'
  and policyname in (
    'products_guest_select',
    'movements_guest_select',
    'production_records_guest_select',
    'production_options_guest_select'
  )
order by tablename;
