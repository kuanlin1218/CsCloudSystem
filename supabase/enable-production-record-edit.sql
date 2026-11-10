-- 啟用「加工數量查詢－編輯單筆加工紀錄」。
-- 可重複執行；請在 Supabase SQL Editor 建立 New query、貼上整份內容並按 Run。

begin;

drop policy if exists production_records_owner_update on public.production_records;
create policy production_records_owner_update
on public.production_records
for update
to authenticated
using (
  (select auth.uid()) = user_id
  and user_id = '5568706d-caaa-459b-b842-13cdbc3929fc'::uuid
)
with check (
  (select auth.uid()) = user_id
  and user_id = '5568706d-caaa-459b-b842-13cdbc3929fc'::uuid
);

grant update (work_date, product_code, machine, hours, minutes, quantity, notes)
on table public.production_records
to authenticated;

commit;
