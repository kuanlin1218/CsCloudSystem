import type { ProductionRecord } from '../types'

type ZipEntry = { name: string; data: Uint8Array; crc: number; offset: number }

const encoder = new TextEncoder()
const crcTable = Array.from({ length: 256 }, (_, index) => {
  let value = index
  for (let bit = 0; bit < 8; bit += 1) value = (value & 1) ? 0xEDB88320 ^ (value >>> 1) : value >>> 1
  return value >>> 0
})

function crc32(data: Uint8Array) {
  let crc = 0xFFFFFFFF
  for (const byte of data) crc = crcTable[(crc ^ byte) & 0xFF] ^ (crc >>> 8)
  return (crc ^ 0xFFFFFFFF) >>> 0
}

function concat(parts: Uint8Array[]) {
  const output = new Uint8Array(parts.reduce((sum, part) => sum + part.length, 0))
  let offset = 0
  for (const part of parts) { output.set(part, offset); offset += part.length }
  return output
}

function set16(view: DataView, offset: number, value: number) { view.setUint16(offset, value, true) }
function set32(view: DataView, offset: number, value: number) { view.setUint32(offset, value >>> 0, true) }

function dosTimestamp(date = new Date()) {
  const year = Math.max(date.getFullYear(), 1980)
  return {
    time: (date.getHours() << 11) | (date.getMinutes() << 5) | Math.floor(date.getSeconds() / 2),
    date: ((year - 1980) << 9) | ((date.getMonth() + 1) << 5) | date.getDate(),
  }
}

function createZip(files: Array<{ name: string; content: string }>) {
  const timestamp = dosTimestamp()
  const localParts: Uint8Array[] = []
  const entries: ZipEntry[] = []
  let offset = 0
  for (const file of files) {
    const name = encoder.encode(file.name)
    const data = encoder.encode(file.content)
    const crc = crc32(data)
    const header = new Uint8Array(30)
    const view = new DataView(header.buffer)
    set32(view, 0, 0x04034B50); set16(view, 4, 20); set16(view, 6, 0x0800); set16(view, 8, 0)
    set16(view, 10, timestamp.time); set16(view, 12, timestamp.date); set32(view, 14, crc)
    set32(view, 18, data.length); set32(view, 22, data.length); set16(view, 26, name.length); set16(view, 28, 0)
    localParts.push(header, name, data)
    entries.push({ name: file.name, data, crc, offset })
    offset += header.length + name.length + data.length
  }
  const centralParts: Uint8Array[] = []
  for (const entry of entries) {
    const name = encoder.encode(entry.name)
    const header = new Uint8Array(46)
    const view = new DataView(header.buffer)
    set32(view, 0, 0x02014B50); set16(view, 4, 20); set16(view, 6, 20); set16(view, 8, 0x0800); set16(view, 10, 0)
    set16(view, 12, timestamp.time); set16(view, 14, timestamp.date); set32(view, 16, entry.crc)
    set32(view, 20, entry.data.length); set32(view, 24, entry.data.length); set16(view, 28, name.length)
    set16(view, 30, 0); set16(view, 32, 0); set16(view, 34, 0); set16(view, 36, 0); set32(view, 38, 0); set32(view, 42, entry.offset)
    centralParts.push(header, name)
  }
  const central = concat(centralParts)
  const end = new Uint8Array(22)
  const endView = new DataView(end.buffer)
  set32(endView, 0, 0x06054B50); set16(endView, 4, 0); set16(endView, 6, 0)
  set16(endView, 8, entries.length); set16(endView, 10, entries.length); set32(endView, 12, central.length); set32(endView, 16, offset); set16(endView, 20, 0)
  return concat([...localParts, central, end])
}

function escapeXml(value: string | number) {
  return String(value).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&apos;')
}

function excelDate(value: string) {
  const [year, month, day] = value.split('-').map(Number)
  return `${year}/${month}/${day}`
}

function columnName(index: number) {
  let name = ''
  let current = index
  while (current > 0) { current -= 1; name = String.fromCharCode(65 + current % 26) + name; current = Math.floor(current / 26) }
  return name
}

function cellXml(row: number, column: number, value: string | number, style: number) {
  const address = `${columnName(column)}${row}`
  if (typeof value === 'number') return `<c r="${address}" s="${style}" t="n"><v>${value}</v></c>`
  return `<c r="${address}" s="${style}" t="inlineStr"><is><t xml:space="preserve">${escapeXml(value)}</t></is></c>`
}

function rowXml(row: number, values: Array<string | number>, style: number) {
  return `<row r="${row}">${Array.from({ length: 7 }, (_, index) => cellXml(row, index + 1, values[index] ?? '', style)).join('')}</row>`
}

function safeFileName(value: string) {
  return value.replace(/[<>:"/\\|?*]/g, '_').trim() || '未命名產品'
}

function buildWorkbookFiles(productCode: string, sourceRecords: ProductionRecord[]) {
  const records = [...sourceRecords].sort((a, b) =>
    `${a.work_date}-${a.machine}-${a.created_at}-${a.id}`.localeCompare(`${b.work_date}-${b.machine}-${b.created_at}-${b.id}`, 'zh-Hant'),
  )
  const totalDays = new Set(records.map((record) => record.work_date)).size
  const totalQuantity = records.reduce((sum, record) => sum + record.quantity, 0)
  const totalMinutes = records.reduce((sum, record) => sum + record.total_minutes, 0)
  const dailyAverage = totalDays ? Math.round(totalQuantity / totalDays) : 0
  const pieceSeconds = totalQuantity ? Math.round(totalMinutes * 60 / totalQuantity) : 0
  const dataRows = records.map((record, index) => rowXml(index + 3, [
    excelDate(record.work_date), record.machine, record.hours, record.minutes,
    record.quantity, record.total_minutes, record.notes || '',
  ], 3)).join('')
  const summaryRow = records.length + 3
  const lastRow = summaryRow + 1
  const sheetXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:G${lastRow}"/><sheetViews><sheetView showGridLines="1" workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/><cols><col min="1" max="1" width="12.7109375" customWidth="1"/><col min="2" max="6" width="10.7109375" customWidth="1"/><col min="7" max="7" width="20.7109375" customWidth="1"/></cols><sheetData>${rowXml(1, [`產品編號：${productCode.trim().toUpperCase()}`], 1)}${rowXml(2, ['日期', '機台', '時', '分', '數量', '總時間', '備註'], 2)}${dataRows}${rowXml(summaryRow, ['合計', '總天數', '總數量', '總時間', '1天平均(顆)', '1顆平均(s)', ''], 4)}${rowXml(lastRow, ['', totalDays, totalQuantity, totalMinutes, dailyAverage, pieceSeconds, ''], 4)}</sheetData><mergeCells count="1"><mergeCell ref="A1:G1"/></mergeCells><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/><pageSetup orientation="landscape" fitToWidth="1" fitToHeight="0"/></worksheet>`
  const timestamp = new Date().toISOString()
  return [
    { name: '[Content_Types].xml', content: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>` },
    { name: '_rels/.rels', content: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>` },
    { name: 'docProps/app.xml', content: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>詮栩 CNC 雲端系統</Application></Properties>` },
    { name: 'docProps/core.xml', content: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:creator>詮栩 CNC 雲端系統</dc:creator><dcterms:created xsi:type="dcterms:W3CDTF">${timestamp}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${timestamp}</dcterms:modified></cp:coreProperties>` },
    { name: 'xl/workbook.xml', content: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="加工數量" sheetId="1" r:id="rId1"/></sheets><calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>` },
    { name: 'xl/_rels/workbook.xml.rels', content: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>` },
    { name: 'xl/styles.xml', content: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="12"/><name val="Calibri"/><family val="2"/></font><font><b/><sz val="12"/><name val="Calibri"/><family val="2"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFEEEEEE"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="5"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="0" fillId="2" borderId="0" xfId="0" applyFill="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf><xf numFmtId="0" fontId="0" fillId="2" borderId="0" xfId="0" applyFill="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>` },
    { name: 'xl/worksheets/sheet1.xml', content: sheetXml },
  ]
}

export function createProductionXlsxData(productCode: string, records: ProductionRecord[]) {
  return createZip(buildWorkbookFiles(productCode, records))
}

export async function exportProductionXlsx(productCode: string, records: ProductionRecord[]) {
  const zip = createProductionXlsxData(productCode, records)
  const data = zip.buffer.slice(zip.byteOffset, zip.byteOffset + zip.byteLength) as ArrayBuffer
  const blob = new Blob([data], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' })
  const url = URL.createObjectURL(blob)
  const link = document.createElement('a')
  link.href = url
  link.download = `${safeFileName(productCode)}_加工數量查詢.xlsx`
  document.body.appendChild(link)
  link.click()
  link.remove()
  window.setTimeout(() => URL.revokeObjectURL(url), 1000)
}
