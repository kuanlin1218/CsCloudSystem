import { read, utils, writeFileXLSX } from 'xlsx'
import type { AdminDataPayload, AdminMovementRow, AdminOptionRow, AdminProductRow, AdminProductionRow } from './adminData'

const SHEETS = {
  guide: '使用說明', products: '刀具資料', movements: '出入庫歷史', production: '加工數量', options: '管理選項',
} as const

type CellRow = Record<string, unknown>

const text = (value: unknown) => String(value ?? '').trim()
const upper = (value: unknown) => text(value).toUpperCase()
const number = (value: unknown, field: string, row: number, integer = false) => {
  const parsed = typeof value === 'number' ? value : Number(String(value ?? '').replaceAll(',', '').trim())
  if (!Number.isFinite(parsed) || (integer && !Number.isInteger(parsed))) throw new Error(`${field}第 ${row} 列不是有效數字`)
  return parsed
}
const optionalNumber = (value: unknown, field: string, row: number) => text(value) === '' ? null : number(value, field, row)
const localDate = (value: unknown, field: string, row: number, optional = false) => {
  if (value instanceof Date && !Number.isNaN(value.valueOf())) {
    const y = value.getFullYear(); const m = String(value.getMonth() + 1).padStart(2, '0'); const d = String(value.getDate()).padStart(2, '0')
    return `${y}-${m}-${d}`
  }
  const raw = text(value).replaceAll('/', '-')
  if (optional && !raw) return null
  const match = raw.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/)
  if (!match) throw new Error(`${field}第 ${row} 列日期格式不正確，請使用 YYYY-MM-DD`)
  const result = `${match[1]}-${match[2].padStart(2, '0')}-${match[3].padStart(2, '0')}`
  const checked = new Date(`${result}T00:00:00`)
  if (Number.isNaN(checked.valueOf()) || checked.getFullYear() !== Number(match[1]) || checked.getMonth() + 1 !== Number(match[2]) || checked.getDate() !== Number(match[3])) throw new Error(`${field}第 ${row} 列日期不存在`)
  return result
}

function rows(workbook: ReturnType<typeof read>, sheetName: string) {
  const sheet = workbook.Sheets[sheetName]
  if (!sheet) throw new Error(`找不到「${sheetName}」工作表，請使用系統匯出的 Excel 範本`)
  return utils.sheet_to_json<CellRow>(sheet, { defval: '', raw: true })
}

function addSheet(workbook: ReturnType<typeof utils.book_new>, name: string, data: CellRow[], widths: number[]) {
  const sheet = utils.json_to_sheet(data)
  sheet['!cols'] = widths.map((wch) => ({ wch }))
  if (sheet['!ref']) sheet['!autofilter'] = { ref: sheet['!ref'] }
  utils.book_append_sheet(workbook, sheet, name)
}

export function downloadAdminWorkbook(payload: AdminDataPayload, exportedAt: string) {
  const workbook = utils.book_new()
  const guide = utils.aoa_to_sheet([
    ['詮栩 CNC 雲端系統｜完整資料備份'],
    ['匯出時間', new Date(exportedAt).toLocaleString('zh-TW')],
    [],
    ['使用方式'],
    ['1', '保留工作表名稱與第一列欄位名稱。'],
    ['2', '整理資料後，回到首頁「資料備份／還原」選擇本檔匯入。'],
    ['3', '匯入會完整取代雲端資料，請先保留一份未修改的備份。'],
    ['4', '「品名（自動組合）」只供閱讀，匯入時會依規格等欄位重新產生。'],
    ['5', '修改「目前庫存」時，系統會自動建立一筆「Excel 匯入盤點調整」紀錄。'],
  ])
  guide['!cols'] = [{ wch: 12 }, { wch: 78 }]
  utils.book_append_sheet(workbook, guide, SHEETS.guide)

  addSheet(workbook, SHEETS.products, payload.products.map((p) => ({
    '刀具編號': p.code, '分類': p.category, '規格': p.specification, '尺寸': p.size,
    '斷屑槽': p.chipbreaker, '材質': p.material, '品名（自動組合）': p.name || '',
    '廠牌': p.brand, '供應商': p.supplier, '單位': p.unit, '目前庫存': p.stock_quantity,
    '最新單價': p.unit_price, '最近購入日期': p.purchased_at || '', '備註': p.notes,
  })), [14, 10, 18, 14, 12, 14, 34, 16, 16, 9, 12, 12, 14, 30])

  addSheet(workbook, SHEETS.movements, payload.movements.map((m, index) => ({
    '紀錄順序': m.sequence || index + 1, '刀具編號': m.product_code,
    '類型': m.movement_type === 'in' ? '入庫' : '出庫', '日期': m.movement_date,
    '數量': m.quantity, '單價': m.unit_price ?? '', '供應商': m.supplier, '備註': m.notes,
  })), [12, 14, 10, 14, 10, 12, 18, 34])

  addSheet(workbook, SHEETS.production, payload.production_records.map((r) => ({
    '日期': r.work_date, '產品編號': r.product_code, '機台': r.machine,
    '時': r.hours, '分': r.minutes, '數量': r.quantity, '備註': r.notes,
  })), [14, 24, 12, 9, 9, 12, 34])

  addSheet(workbook, SHEETS.options, payload.production_options.map((o) => ({
    '類型': o.option_type === 'product' ? '產品編號' : '機台', '值': o.value,
  })), [16, 32])

  const day = new Date().toLocaleDateString('sv-SE')
  writeFileXLSX(workbook, `詮栩CNC完整備份_${day}.xlsx`, { compression: true })
}

export interface WorkbookPreview {
  payload: AdminDataPayload
  summary: { products: number; movements: number; productionRecords: number; options: number }
}

export async function parseAdminWorkbook(file: File): Promise<WorkbookPreview> {
  const workbook = read(await file.arrayBuffer(), { type: 'array', cellDates: true })
  const errors: string[] = []
  let products: AdminProductRow[] = []
  let movements: AdminMovementRow[] = []
  let productionRecords: AdminProductionRow[] = []
  let options: AdminOptionRow[] = []

  try {
    products = rows(workbook, SHEETS.products).filter((r) => Object.values(r).some((v) => text(v))).map((r, i) => {
      const row = i + 2; const code = upper(r['刀具編號']); const category = text(r['分類'])
      if (!code) throw new Error(`刀具資料第 ${row} 列缺少刀具編號`)
      if (!['刀片', '鑽銑刀', '刀柄', '其他'].includes(category)) throw new Error(`刀具資料第 ${row} 列分類不正確`)
      const stock = number(r['目前庫存'], '刀具資料「目前庫存」', row, true)
      const price = number(r['最新單價'], '刀具資料「最新單價」', row)
      if (stock < 0 || price < 0) throw new Error(`刀具資料第 ${row} 列的庫存或單價不可小於 0`)
      return { code, category, specification: text(r['規格']), size: text(r['尺寸']), chipbreaker: text(r['斷屑槽']), material: text(r['材質']), brand: text(r['廠牌']), supplier: text(r['供應商']), unit: text(r['單位']) || '個', stock_quantity: stock, unit_price: price, purchased_at: localDate(r['最近購入日期'], '刀具資料「最近購入日期」', row, true), notes: text(r['備註']) }
    })
    const duplicate = products.find((p, index) => products.findIndex((q) => q.code === p.code) !== index)
    if (duplicate) throw new Error(`刀具編號「${duplicate.code}」重複`)
  } catch (error) { errors.push(error instanceof Error ? error.message : '刀具資料格式錯誤') }

  try {
    movements = rows(workbook, SHEETS.movements).filter((r) => Object.values(r).some((v) => text(v))).map((r, i) => {
      const row = i + 2; const productCode = upper(r['刀具編號']); const type = text(r['類型'])
      if (!productCode) throw new Error(`出入庫歷史第 ${row} 列缺少刀具編號`)
      if (!['入庫', '出庫', 'in', 'out'].includes(type)) throw new Error(`出入庫歷史第 ${row} 列類型必須是入庫或出庫`)
      const quantity = number(r['數量'], '出入庫歷史「數量」', row, true)
      if (quantity <= 0) throw new Error(`出入庫歷史第 ${row} 列數量必須大於 0`)
      return { sequence: optionalNumber(r['紀錄順序'], '出入庫歷史「紀錄順序」', row) || i + 1, product_code: productCode, movement_type: (type === '入庫' || type === 'in') ? 'in' as const : 'out' as const, movement_date: localDate(r['日期'], '出入庫歷史「日期」', row)!, quantity, unit_price: optionalNumber(r['單價'], '出入庫歷史「單價」', row), supplier: text(r['供應商']), notes: text(r['備註']) }
    }).sort((a, b) => (a.sequence || 0) - (b.sequence || 0))
  } catch (error) { errors.push(error instanceof Error ? error.message : '出入庫歷史格式錯誤') }

  try {
    productionRecords = rows(workbook, SHEETS.production).filter((r) => Object.values(r).some((v) => text(v))).map((r, i) => {
      const row = i + 2; const productCode = upper(r['產品編號']); const machine = upper(r['機台'])
      if (!productCode || !machine) throw new Error(`加工數量第 ${row} 列缺少產品編號或機台`)
      const hours = number(r['時'], '加工數量「時」', row, true); const minutes = number(r['分'], '加工數量「分」', row, true); const quantity = number(r['數量'], '加工數量「數量」', row, true)
      if (hours < 0 || minutes < 0 || minutes > 59 || quantity <= 0) throw new Error(`加工數量第 ${row} 列的時間或數量不正確`)
      return { work_date: localDate(r['日期'], '加工數量「日期」', row)!, product_code: productCode, machine, hours, minutes, quantity, notes: text(r['備註']) }
    })
  } catch (error) { errors.push(error instanceof Error ? error.message : '加工數量格式錯誤') }

  try {
    options = rows(workbook, SHEETS.options).filter((r) => Object.values(r).some((v) => text(v))).map((r, i) => {
      const type = text(r['類型']); const value = upper(r['值'])
      if (!['產品編號', '機台', 'product', 'machine'].includes(type) || !value) throw new Error(`管理選項第 ${i + 2} 列格式不正確`)
      return { option_type: (type === '產品編號' || type === 'product') ? 'product' as const : 'machine' as const, value }
    })
    options = options.filter((item, index) => options.findIndex((other) => other.option_type === item.option_type && other.value === item.value) === index)
  } catch (error) { errors.push(error instanceof Error ? error.message : '管理選項格式錯誤') }

  if (!errors.length) {
    const productCodes = new Set(products.map((p) => p.code))
    const missing = movements.find((m) => !productCodes.has(m.product_code))
    if (missing) errors.push(`出入庫歷史使用了不存在的刀具編號「${missing.product_code}」`)
  }
  if (errors.length) throw new Error(errors.slice(0, 5).join('；'))

  const payload = { products, movements, production_records: productionRecords, production_options: options }
  return { payload, summary: { products: products.length, movements: movements.length, productionRecords: productionRecords.length, options: options.length } }
}
