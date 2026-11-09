-- 詮栩 CNC 雲端系統：啟用手機與電腦即時同步
-- 請將本檔案全部內容貼到 Supabase SQL Editor 的 New query，按 Run。
-- 此程式可重複執行，不會重複加入資料表。

do $$
declare
  realtime_table text;
begin
  if not exists (select 1 from pg_publication where pubname = 'supabase_realtime') then
    execute 'create publication supabase_realtime';
  end if;

  foreach realtime_table in array array['products', 'stock_movements', 'production_records', 'production_options']
  loop
    if not exists (
      select 1
      from pg_publication_tables
      where pubname = 'supabase_realtime'
        and schemaname = 'public'
        and tablename = realtime_table
    ) then
      execute format('alter publication supabase_realtime add table public.%I', realtime_table);
    end if;
  end loop;
end;
$$;

-- 執行完成後應列出四個資料表。
select schemaname, tablename
from pg_publication_tables
where pubname = 'supabase_realtime'
  and schemaname = 'public'
  and tablename in ('products', 'stock_movements', 'production_records', 'production_options')
order by tablename;
