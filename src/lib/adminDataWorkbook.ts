import ExcelJS from 'exceljs'
import { read, utils } from 'xlsx'
import type { AdminDataPayload, AdminMachineProfileRow, AdminMachineServiceRow, AdminMovementRow, AdminOptionRow, AdminProductRow, AdminProductionRow } from './adminData'
import { createQrCodeDataUrl } from './barcode'

const SHEETS = {
  guide: '使用說明', products: '刀具資料', movements: '出入庫歷史', production: '加工數量', options: '管理選項',
  machineProfiles: '機台基本資料', machineService: '保養維修紀錄', qrLabels: 'QR Code 標籤',
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

function addDataSheet(workbook: ExcelJS.Workbook, name: string, data: Record<string, string | number>[], columns: { header: string; key: string; width: number }[]) {
  const sheet = workbook.addWorksheet(name)
  sheet.columns = columns
  sheet.addRows(data)
  sheet.views = [{ state: 'frozen', ySplit: 1 }]
  sheet.autoFilter = {
    from: { row: 1, column: 1 },
    to: { row: Math.max(2, data.length + 1), column: columns.length },
  }
  const header = sheet.getRow(1)
  header.height = 25
  header.font = { name: 'Microsoft JhengHei', bold: true, color: { argb: 'FFFFFFFF' } }
  header.alignment = { vertical: 'middle', horizontal: 'center' }
  header.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF176B8F' } }
  sheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return
    row.height = 22
    row.font = { name: 'Microsoft JhengHei', size: 10, color: { argb: 'FF172126' } }
    row.alignment = { vertical: 'middle' }
    row.eachCell((cell) => {
      cell.border = { bottom: { style: 'thin', color: { argb: 'FFDCE3E6' } } }
    })
  })
  return sheet
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

/* 「資料備份／還原」的「完整 Excel 與 QR Code 標籤匯出」：全程使用 ExcelJS，避免不同 Excel 套件轉換後產生損壞的篩選器定義。 */
export async function downloadAdminWorkbook(payload: AdminDataPayload, exportedAt: string) {
  const workbook = new ExcelJS.Workbook()
  workbook.creator = '詮栩 CNC 雲端系統'
  workbook.created = new Date(exportedAt)
  workbook.modified = new Date()
  const guide = workbook.addWorksheet(SHEETS.guide)
  guide.addRows([
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
  guide.columns = [{ width: 12 }, { width: 78 }]
  guide.mergeCells('A1:B1')
  guide.getRow(1).height = 32
  guide.getCell('A1').font = { name: 'Microsoft JhengHei', size: 17, bold: true, color: { argb: 'FF172126' } }
  guide.getCell('A1').alignment = { vertical: 'middle', horizontal: 'center' }
  guide.getCell('A1').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE9F3F1' } }
  guide.eachRow((row, rowNumber) => {
    if (rowNumber > 1) row.font = { name: 'Microsoft JhengHei', size: 11, color: { argb: 'FF27363D' } }
    row.alignment = { vertical: 'middle', wrapText: true }
  })

  addDataSheet(workbook, SHEETS.products, payload.products.map((p) => ({
    code: p.code, barcode: p.barcode || '', category: p.category, specification: p.specification, size: p.size,
    chipbreaker: p.chipbreaker, material: p.material, name: p.name || '', brand: p.brand, supplier: p.supplier,
    unit: p.unit, stock: p.stock_quantity, price: p.unit_price, purchasedAt: p.purchased_at || '', notes: p.notes,
  })), [
    { header: '刀具編號', key: 'code', width: 14 }, { header: '條碼', key: 'barcode', width: 20 },
    { header: '分類', key: 'category', width: 10 }, { header: '規格', key: 'specification', width: 18 },
    { header: '尺寸', key: 'size', width: 14 }, { header: '斷屑槽', key: 'chipbreaker', width: 12 },
    { header: '材質', key: 'material', width: 14 }, { header: '品名（自動組合）', key: 'name', width: 34 },
    { header: '廠牌', key: 'brand', width: 16 }, { header: '供應商', key: 'supplier', width: 16 },
    { header: '單位', key: 'unit', width: 9 }, { header: '目前庫存', key: 'stock', width: 12 },
    { header: '最新單價', key: 'price', width: 12 }, { header: '最近購入日期', key: 'purchasedAt', width: 14 },
    { header: '備註', key: 'notes', width: 30 },
  ])

  addDataSheet(workbook, SHEETS.movements, payload.movements.map((m, index) => ({
    sequence: m.sequence || index + 1, productCode: m.product_code, type: m.movement_type === 'in' ? '入庫' : '出庫',
    date: m.movement_date, quantity: m.quantity, price: m.unit_price ?? '', supplier: m.supplier, notes: m.notes,
  })), [
    { header: '紀錄順序', key: 'sequence', width: 12 }, { header: '刀具編號', key: 'productCode', width: 14 },
    { header: '類型', key: 'type', width: 10 }, { header: '日期', key: 'date', width: 14 },
    { header: '數量', key: 'quantity', width: 10 }, { header: '單價', key: 'price', width: 12 },
    { header: '供應商', key: 'supplier', width: 18 }, { header: '備註', key: 'notes', width: 34 },
  ])

  addDataSheet(workbook, SHEETS.production, payload.production_records.map((r) => ({
    date: r.work_date, productCode: r.product_code, machine: r.machine, hours: r.hours,
    minutes: r.minutes, quantity: r.quantity, notes: r.notes,
  })), [
    { header: '日期', key: 'date', width: 14 }, { header: '產品編號', key: 'productCode', width: 24 },
    { header: '機台', key: 'machine', width: 12 }, { header: '時', key: 'hours', width: 9 },
    { header: '分', key: 'minutes', width: 9 }, { header: '數量', key: 'quantity', width: 12 },
    { header: '備註', key: 'notes', width: 34 },
  ])

  addDataSheet(workbook, SHEETS.options, payload.production_options.map((o) => ({
    type: o.option_type === 'product' ? '產品編號' : '機台', value: o.value,
    orderQuantity: o.order_quantity ?? '', dueDate: o.due_date ?? '', completedOn: o.completed_on ?? '',
  })), [
    { header: '類型', key: 'type', width: 16 }, { header: '值', key: 'value', width: 32 },
    { header: '訂單數量', key: 'orderQuantity', width: 16 }, { header: '交期', key: 'dueDate', width: 16 },
    { header: '完工日期', key: 'completedOn', width: 16 },
  ])

  addDataSheet(workbook, SHEETS.machineProfiles, (payload.machine_profiles ?? []).map((machine) => ({
    code: machine.machine_code, active: machine.machine_active ? '使用中' : '停用', name: machine.name ?? '',
    brand: machine.brand ?? '', model: machine.model ?? '', serial: machine.serial_number ?? '',
    purchased: machine.purchased_on ?? '', started: machine.started_on ?? '', notes: machine.notes ?? '',
  })), [
    { header: '機台編號', key: 'code', width: 16 }, { header: '狀態', key: 'active', width: 12 },
    { header: '機台名稱', key: 'name', width: 22 }, { header: '廠牌', key: 'brand', width: 16 },
    { header: '型號', key: 'model', width: 18 }, { header: '出廠序號', key: 'serial', width: 18 },
    { header: '購入日期', key: 'purchased', width: 14 }, { header: '啟用日期', key: 'started', width: 14 },
    { header: '備註', key: 'notes', width: 34 },
  ])

  addDataSheet(workbook, SHEETS.machineService, (payload.machine_service_records ?? []).map((record) => ({
    code: record.machine_code, date: record.service_date, type: record.service_type === 'maintenance' ? '保養' : '維修',
    description: record.description, technician: record.technician, cost: record.cost ?? '', downtime: record.downtime_minutes ?? '',
  })), [
    { header: '機台編號', key: 'code', width: 16 }, { header: '日期', key: 'date', width: 14 },
    { header: '類型', key: 'type', width: 12 }, { header: '處理內容', key: 'description', width: 40 },
    { header: '處理人員／廠商', key: 'technician', width: 24 }, { header: '費用', key: 'cost', width: 14 },
    { header: '停機分鐘', key: 'downtime', width: 14 },
  ])

  await addQrLabelSheet(workbook, payload.products)
  const finalWorkbook = await workbook.xlsx.writeBuffer()
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
  summary: { products: number; movements: number; productionRecords: number; options: number; machineProfiles: number; machineServiceRecords: number }
}

export async function parseAdminWorkbook(file: File): Promise<WorkbookPreview> {
  const workbook = read(await file.arrayBuffer(), { type: 'array', cellDates: true })
  const errors: string[] = []
  let products: AdminProductRow[] = []
  let movements: AdminMovementRow[] = []
  let productionRecords: AdminProductionRow[] = []
  let options: AdminOptionRow[] = []
  let machineProfiles: AdminMachineProfileRow[] = []
  let machineService: AdminMachineServiceRow[] = []

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
      if (type === '機台' || type === 'machine') return { option_type: 'machine' as const, value }
      const row = i + 2
      const orderQuantity = text(r['訂單數量']) ? number(r['訂單數量'], '管理選項「訂單數量」', row, true) : null
      if (orderQuantity !== null && (orderQuantity < 1 || orderQuantity > 2147483647)) throw new Error(`管理選項第 ${row} 列訂單數量超出有效範圍`)
      return {
        option_type: 'product' as const, value, order_quantity: orderQuantity,
        due_date: localDate(r['交期'], '管理選項「交期」', row, true),
        completed_on: localDate(r['完工日期'], '管理選項「完工日期」', row, true),
      }
    })
    options = options.filter((item, index) => options.findIndex((other) => other.option_type === item.option_type && other.value === item.value) === index)
  } catch (error) { errors.push(error instanceof Error ? error.message : '管理選項格式錯誤') }

  const knownMachineCodes = new Set([...options.filter((option) => option.option_type === 'machine').map((option) => option.value), ...productionRecords.map((record) => record.machine)])
  if (workbook.Sheets[SHEETS.machineProfiles]) {
    try {
      const seen = new Set<string>()
      machineProfiles = rows(workbook, SHEETS.machineProfiles).filter((r) => Object.values(r).some((v) => text(v))).map((r, i) => {
        const row = i + 2
        const machineCode = upper(r['機台編號'])
        const status = text(r['狀態'])
        if (!knownMachineCodes.has(machineCode) || seen.has(machineCode)) throw new Error(`機台基本資料第 ${row} 列的機台編號不存在或重複`)
        if (status && status !== '使用中' && status !== '停用') throw new Error(`機台基本資料第 ${row} 列狀態不正確`)
        seen.add(machineCode)
        return {
          machine_code: machineCode, machine_active: status !== '停用',
          name: text(r['機台名稱']), brand: text(r['廠牌']), model: text(r['型號']), serial_number: text(r['出廠序號']),
          purchased_on: localDate(r['購入日期'], '機台基本資料「購入日期」', row, true),
          started_on: localDate(r['啟用日期'], '機台基本資料「啟用日期」', row, true), notes: text(r['備註']),
        }
      })
    } catch (error) { errors.push(error instanceof Error ? error.message : '機台基本資料格式錯誤') }
  }
  if (workbook.Sheets[SHEETS.machineService]) {
    try {
      machineService = rows(workbook, SHEETS.machineService).filter((r) => Object.values(r).some((v) => text(v))).map((r, i) => {
        const row = i + 2
        const machineCode = upper(r['機台編號'])
        const type = text(r['類型'])
        const description = text(r['處理內容'])
        if (!knownMachineCodes.has(machineCode) || (type !== '保養' && type !== '維修') || !description) {
          throw new Error(`保養維修紀錄第 ${row} 列的機台、類型或處理內容不正確`)
        }
        const cost = optionalNumber(r['費用'], '保養維修紀錄「費用」', row)
        const downtime = text(r['停機分鐘']) ? number(r['停機分鐘'], '保養維修紀錄「停機分鐘」', row, true) : null
        if ((cost !== null && (cost < 0 || cost > 999999999999.99)) || (downtime !== null && (downtime < 0 || downtime > 2147483647))) {
          throw new Error(`保養維修紀錄第 ${row} 列的費用或停機分鐘超出範圍`)
        }
        return {
          machine_code: machineCode, service_date: localDate(r['日期'], '保養維修紀錄「日期」', row)!,
          service_type: type === '保養' ? 'maintenance' : 'repair',
          description, technician: text(r['處理人員／廠商']), cost, downtime_minutes: downtime,
        }
      })
    } catch (error) { errors.push(error instanceof Error ? error.message : '保養維修紀錄格式錯誤') }
  }

  if (!errors.length) {
    const productCodes = new Set(products.map((p) => p.code))
    const missing = movements.find((m) => !productCodes.has(m.product_code))
    if (missing) errors.push(`出入庫歷史使用了不存在的刀具編號「${missing.product_code}」`)
  }
  if (errors.length) throw new Error(errors.slice(0, 5).join('；'))

  const payload = { products, movements, production_records: productionRecords, production_options: options, machine_profiles: machineProfiles, machine_service_records: machineService }
  return { payload, summary: {
    products: products.length, movements: movements.length, productionRecords: productionRecords.length, options: options.length,
    machineProfiles: machineProfiles.length, machineServiceRecords: machineService.length,
  } }
}
