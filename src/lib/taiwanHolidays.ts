/*
 * 台灣國定假日與補假日期。
 * 資料依行政院人事行政總處公布的政府行政機關辦公日曆表整理。
 * 2026：https://www.dgpa.gov.tw/information?pid=12574&uid=41
 * 2027：https://www.dgpa.gov.tw/information?pid=12983&uid=2
 * 每年官方日曆公布後，請在這裡加入新年度日期並更新支援年份。
 */

export const TAIWAN_HOLIDAY_CALENDAR_YEARS = [2026, 2027] as const

const TAIWAN_PUBLIC_HOLIDAYS = new Set([
  /* 2026 年國定假日及補假；相鄰的一般星期六、日不在此清單，由排程選項另外判斷。 */
  '2026-01-01',
  '2026-02-15', '2026-02-16', '2026-02-17', '2026-02-18', '2026-02-19', '2026-02-20',
  '2026-02-27', '2026-02-28',
  '2026-04-03', '2026-04-04', '2026-04-05', '2026-04-06',
  '2026-05-01',
  '2026-06-19',
  '2026-09-25', '2026-09-28',
  '2026-10-09', '2026-10-10',
  '2026-10-25', '2026-10-26',
  '2026-12-25',

  /* 2027 年國定假日及補假；相鄰的一般星期六、日不在此清單，由排程選項另外判斷。 */
  '2027-01-01',
  '2027-02-04', '2027-02-05', '2027-02-06', '2027-02-07', '2027-02-08', '2027-02-09', '2027-02-10',
  '2027-02-28', '2027-03-01',
  '2027-04-04', '2027-04-05', '2027-04-06',
  '2027-04-30', '2027-05-01',
  '2027-06-09',
  '2027-09-15', '2027-09-28',
  '2027-10-10', '2027-10-11',
  '2027-10-25',
  '2027-12-24', '2027-12-25', '2027-12-31',
])

export const isTaiwanPublicHoliday = (dateValue: string) => TAIWAN_PUBLIC_HOLIDAYS.has(dateValue)

export const hasTaiwanHolidayCalendar = (dateValue: string) => {
  const year = Number(dateValue.slice(0, 4))
  return TAIWAN_HOLIDAY_CALENDAR_YEARS.includes(year as (typeof TAIWAN_HOLIDAY_CALENDAR_YEARS)[number])
}

