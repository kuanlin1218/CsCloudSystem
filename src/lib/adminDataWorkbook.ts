import ExcelJS from 'exceljs'
import { read, utils, write } from 'xlsx'
import type { AdminDataPayload, AdminMovementRow, AdminOptionRow, AdminProductRow, AdminProductionRow } from './adminData'
import { createQrCodeDataUrl } from './barcode'

const SHEETS = {
  guide: '使用說明', products: '刀具資料', movements: '出入庫歷史', production: '加工數量', options: '管理選項', qrLabels: 'QR Code 標籤',
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

async function addQrLabelSheet(workbook: ExcelJS.Workbook, products: AdminProductRow[]) {
  const sheet = workbook.addWorksheet(SHEETS.qrLabels, {
    properties: { defaultRowHeight: 21 },
    pageSetup: {
      paperSize: 9,
      orientation: 'landscape',
      fitToPage: true,
      fitToWidth: 1,
      fitToHeight: 0,
      margins: { left: 0.25, right: 0.25, top: 0.35, bottom: 0.35, header: 0.1, footer: 0.1 },
    },
  })
  sheet.views = [{ showGridLines: false }]
  sheet.columns = [9, 9, 16, 16, 3, 9, 9, 16, 16].map((width) => ({ width }))
  sheet.mergeCells('A1:I1')
  const title = sheet.getCell('A1')
  title.value = '詮栩 CNC 雲端系統｜刀具 QR Code 標籤'
  title.font = { name: 'Microsoft JhengHei', size: 16, bold: true, color: { argb: 'FF172126' } }
  title.alignment = { horizontal: 'center', vertical: 'middle' }
  title.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE9F3F1' } }
  sheet.getRow(1).height = 30

  const labelProducts = products.filter((product) => product.barcode.trim())
  if (!labelProducts.length) {
    sheet.mergeCells('A3:I5')
    const empty = sheet.getCell('A3')
    empty.value = '目前沒有已綁定條碼的刀具'
    empty.font = { name: 'Microsoft JhengHei', size: 14, color: { argb: 'FF68777E' } }
    empty.alignment = { horizontal: 'center', vertical: 'middle' }
    return
  }

  const qrImages = await Promise.all(labelProducts.map((product) => createQrCodeDataUrl(product.barcode, 320)))
  const thinBorder: Partial<ExcelJS.Borders> = {
    top: { style: 'thin', color: { argb: 'FFB7C3C8' } },
    left: { style: 'thin', color: { argb: 'FFB7C3C8' } },
    bottom: { style: 'thin', color: { argb: 'FFB7C3C8' } },
    right: { style: 'thin', color: { argb: 'FFB7C3C8' } },
  }

  labelProducts.forEach((product, index) => {
    const blockRow = 3 + Math.floor(index / 2) * 5
    const leftColumn = index % 2 === 0 ? 1 : 6
    const textStart = leftColumn + 2
    const textEnd = leftColumn + 3

    for (let row = blockRow; row < blockRow + 5; row += 1) {
      sheet.getRow(row).height = 21
      for (let column = leftColumn; column <= leftColumn + 3; column += 1) {
        const cell = sheet.getCell(row, column)
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFFFF' } }
        cell.border = thinBorder
      }
      sheet.mergeCells(row, textStart, row, textEnd)
    }
    sheet.mergeCells(blockRow, leftColumn, blockRow + 4, leftColumn + 1)

    const values = [
      { value: product.code, size: 14, bold: true, color: 'FFC68A00' },
      { value: product.name || product.specification, size: 11, bold: true, color: 'FF172126' },
      { value: `條碼：${product.barcode}`, size: 9, bold: false, color: 'FF34454D' },
      { value: `${product.category}${product.brand ? `｜${product.brand}` : ''}`, size: 9, bold: false, color: 'FF52636B' },
      { value: product.supplier ? `供應商：${product.supplier}` : '供應商：—', size: 9, bold: false, color: 'FF52636B' },
    ]
    values.forEach((item, offset) => {
      const cell = sheet.getCell(blockRow + offset, textStart)
      cell.value = item.value
      cell.font = { name: 'Microsoft JhengHei', size: item.size, bold: item.bold, color: { argb: item.color } }
      cell.alignment = { vertical: 'middle', wrapText: true }
    })

    const imageId = workbook.addImage({ base64: qrImages[index], extension: 'png' })
    sheet.addImage(imageId, {
      tl: { col: leftColumn - 1 + 0.25, row: blockRow - 1 + 0.2 },
      ext: { width: 112, height: 112 },
      editAs: 'oneCell',
    })
  })

  sheet.pageSetup.printArea = `A1:I${3 + Math.ceil(labelProducts.length / 2) * 5}`
}

export async function downloadAdminWorkbook(payload: AdminDataPayload, exportedAt: string) {
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
    ['6', '「QR Code 標籤」會列出所有已綁定條碼的刀具，可直接列印；匯入時只讀取「刀具資料」中的條碼文字。'],
  ])
  guide['!cols'] = [{ wch: 12 }, { wch: 78 }]
  utils.book_append_sheet(workbook, guide, SHEETS.guide)

  addSheet(workbook, SHEETS.products, payload.products.map((p) => ({
    '刀具編號': p.code, '條碼': p.barcode || '', '分類': p.category, '規格': p.specification, '尺寸': p.size,
    '斷屑槽': p.chipbreaker, '材質': p.material, '品名（自動組合）': p.name || '',
    '廠牌': p.brand, '供應商': p.supplier, '單位': p.unit, '目前庫存': p.stock_quantity,
    '最新單價': p.unit_price, '最近購入日期': p.purchased_at || '', '備註': p.notes,
  })), [14, 20, 10, 18, 14, 12, 14, 34, 16, 16, 9, 12, 12, 14, 30])

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

  const baseWorkbook = write(workbook, { bookType: 'xlsx', type: 'array', compression: true })
  const printableWorkbook = new ExcelJS.Workbook()
  await printableWorkbook.xlsx.load(baseWorkbook)
  await addQrLabelSheet(printableWorkbook, payload.products)
  const finalWorkbook = await printableWorkbook.xlsx.writeBuffer()
  const downloadUrl = URL.createObjectURL(new Blob([finalWorkbook], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }))
  const link = document.createElement('a')
  const day = new Date().toLocaleDateString('sv-SE')
  link.href = downloadUrl
  link.download = `詮栩CNC完整備份_${day}.xlsx`
  document.body.appendChild(link)
  link.click()
  link.remove()
  window.setTimeout(() => URL.revokeObjectURL(downloadUrl), 1000)
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
      return { code, barcode: text(r['條碼']), category, specification: text(r['規格']), size: text(r['尺寸']), chipbreaker: text(r['斷屑槽']), material: text(r['材質']), brand: text(r['廠牌']), supplier: text(r['供應商']), unit: text(r['單位']) || '個', stock_quantity: stock, unit_price: price, purchased_at: localDate(r['最近購入日期'], '刀具資料「最近購入日期」', row, true), notes: text(r['備註']) }
    })
    const duplicate = products.find((p, index) => products.findIndex((q) => q.code === p.code) !== index)
    if (duplicate) throw new Error(`刀具編號「${duplicate.code}」重複`)
    const duplicateBarcode = products.find((p, index) => p.barcode && products.findIndex((q) => q.barcode.toLowerCase() === p.barcode.toLowerCase()) !== index)
    if (duplicateBarcode) throw new Error(`條碼「${duplicateBarcode.barcode}」重複`)
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
