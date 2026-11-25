import { dataStorage, isLocalDataMode, supabase } from './supabase'
import { composeProductName, today, todayInTaipei } from './format'
import { PRODUCT_CATEGORIES, type Movement, type MovementInput, type Product, type ProductCategory, type ProductInput, type ProductionInput, type ProductionOption, type ProductionOptionType, type ProductionRecord, type StocktakeInput } from '../types'
import { requireAuthenticatedUserId } from './auth'
import type { MachineProfile, MachineServiceInput, MachineServiceRecord, OperationLogEntry } from '../types'

type StoredOperationLogEntry = Omit<OperationLogEntry, 'id'> & { id: number }

export async function loadOperationLog(cursor?: Pick<OperationLogEntry, 'id' | 'occurred_at'>) {
  if (!supabase) throw new Error('請連線至 Supabase 後查看操作紀錄')
  const timeline = await supabase.rpc('read_operation_timeline', cursor
    ? { p_before_time: cursor.occurred_at, p_before_id: cursor.id }
    : {})
  if (!timeline.error) {
    const data = timeline.data as OperationLogEntry[]
    return { entries: data.slice(0, 50), hasMore: data.length > 50, authIncluded: true }
  }
  if (timeline.error.code !== 'PGRST202') throw timeline.error

  // 先上傳網站時，仍可查看舊版資料異動紀錄，直到套用登入登出 SQL。
  let query = supabase.from('operation_log').select('*')
  if (cursor?.id.startsWith('data:')) {
    const legacyId = cursor.id.slice(5)
    query = query.or(`occurred_at.lt.${cursor.occurred_at},and(occurred_at.eq.${cursor.occurred_at},id.lt.${legacyId})`)
  }
  const { data, error } = await query.order('occurred_at', { ascending: false }).order('id', { ascending: false }).limit(51)
  if (error?.code === 'PGRST205' || error?.code === '42P01') throw new Error('操作紀錄尚未啟用，請先套用 enable-operation-log.sql。')
  if (error) throw error
  const legacyEntries = data as StoredOperationLogEntry[]
  const entries = legacyEntries.slice(0, 50).map((entry) => ({ ...entry, id: `data:${String(entry.id).padStart(19, '0')}` }))
  return { entries, hasMore: data.length > 50, authIncluded: false }
}

const PRODUCT_KEY = 'blade_inventory_products_v3'
const MOVEMENT_KEY = 'blade_inventory_movements_v3'
const PRODUCTION_KEY = 'cnc_production_records_v1'
const PRODUCTION_OPTION_KEY = 'cnc_production_options_v1'
const MACHINE_PROFILE_KEY = 'cnc_machine_profiles_v1'
const MACHINE_SERVICE_KEY = 'cnc_machine_service_records_v1'
const previousDay = (() => { const date = new Date(`${today()}T12:00:00`); date.setDate(date.getDate() - 1); return date.toISOString().slice(0, 10) })()
const codePrefix = (category: ProductCategory) => ({ '刀片': 'T', '鑽銑刀': 'M', '刀柄': 'H', '其他': 'O' }[category])

const seedProducts: Product[] = [
  { id: 'demo-p1', code: 'T0001', barcode: '4710000000011', name: 'WNMG-080408-MS-PR1535', specification: 'WNMG', size: '080408', chipbreaker: 'MS', material: 'PR1535', brand: '京瓷', supplier: '福燕', category: '刀片', unit: '片', stock_quantity: 190, unit_price: 200, purchased_at: today(), notes: '鋼件・通用加工', created_at: new Date().toISOString(), updated_at: new Date().toISOString() },
  { id: 'demo-p2', code: 'T0002', barcode: '', name: 'TNMG-160404-無-VP15TF', specification: 'TNMG', size: '160404', chipbreaker: '無', material: 'VP15TF', brand: '三菱', supplier: '怡立', category: '刀片', unit: '片', stock_quantity: 100, unit_price: 100, purchased_at: today(), notes: '', created_at: new Date().toISOString(), updated_at: new Date().toISOString() },
  { id: 'demo-p3', code: 'T0003', barcode: '', name: 'VNMG-160404-GU-AC8025P', specification: 'VNMG', size: '160404', chipbreaker: 'GU', material: 'AC8025P', brand: '住友', supplier: '益詮', category: '刀片', unit: '片', stock_quantity: 100, unit_price: 165, purchased_at: today(), notes: '', created_at: new Date().toISOString(), updated_at: new Date().toISOString() },
  { id: 'demo-p4', code: 'M0001', barcode: '', name: 'EMM4100-10-鎢鋼', specification: 'EMM4100', size: '10', chipbreaker: '', material: '鎢鋼', brand: '益詮', supplier: '益詮', category: '鑽銑刀', unit: '支', stock_quantity: 5, unit_price: 800, purchased_at: today(), notes: '立銑刀', created_at: new Date().toISOString(), updated_at: new Date().toISOString() },
]

const seedMovements: Movement[] = [
  { id: 'demo-m8', product_id: 'demo-p4', movement_type: 'in', quantity: 3, movement_date: today(), unit_price: 800, supplier: '益詮', notes: '', stock_after: 6, request_key: 'seed-8', created_at: new Date().toISOString(), product: { code: 'M0001', name: 'EMM4100-10-鎢鋼', unit: '支', brand: '益詮' } },
  { id: 'demo-m7', product_id: 'demo-p4', movement_type: 'out', quantity: 1, movement_date: today(), unit_price: null, supplier: '', notes: '試刀', stock_after: 5, request_key: 'seed-7', created_at: new Date().toISOString(), product: { code: 'M0001', name: 'EMM4100-10-鎢鋼', unit: '支', brand: '益詮' } },
  { id: 'demo-m6', product_id: 'demo-p4', movement_type: 'in', quantity: 3, movement_date: previousDay, unit_price: 780, supplier: '益詮', notes: '', stock_after: 3, request_key: 'seed-6', created_at: new Date(Date.now() - 86_400_000).toISOString(), product: { code: 'M0001', name: 'EMM4100-10-鎢鋼', unit: '支', brand: '益詮' } },
  { id: 'demo-m1', product_id: 'demo-p3', movement_type: 'in', quantity: 100, movement_date: today(), unit_price: 165, supplier: '益詮', notes: '', stock_after: 100, request_key: 'seed-1', created_at: new Date().toISOString(), product: { code: 'T0003', name: 'VNMG-160404-GU-AC8025P', unit: '片', brand: '住友' } },
  { id: 'demo-m2', product_id: 'demo-p2', movement_type: 'in', quantity: 100, movement_date: today(), unit_price: 100, supplier: '怡立', notes: '', stock_after: 100, request_key: 'seed-2', created_at: new Date().toISOString(), product: { code: 'T0002', name: 'TNMG-160404-無-VP15TF', unit: '片', brand: '三菱' } },
  { id: 'demo-m5', product_id: 'demo-p1', movement_type: 'in', quantity: 100, movement_date: today(), unit_price: 200, supplier: '福燕', notes: '', stock_after: 200, request_key: 'seed-5', created_at: new Date().toISOString(), product: { code: 'T0001', name: 'WNMG-080408-MS-PR1535', unit: '片', brand: '京瓷' } },
  { id: 'demo-m3', product_id: 'demo-p1', movement_type: 'in', quantity: 100, movement_date: previousDay, unit_price: 180, supplier: '福燕', notes: '期初庫存', stock_after: 100, request_key: 'seed-3', created_at: new Date(Date.now() - 86_400_000).toISOString(), product: { code: 'T0001', name: 'WNMG-080408-MS-PR1535', unit: '片', brand: '京瓷' } },
  { id: 'demo-m4', product_id: 'demo-p1', movement_type: 'out', quantity: 10, movement_date: today(), unit_price: null, supplier: '', notes: '機台領用', stock_after: 190, request_key: 'seed-4', created_at: new Date().toISOString(), product: { code: 'T0001', name: 'WNMG-080408-MS-PR1535', unit: '片', brand: '京瓷' } },
]

const productionDate = (daysAgo: number) => { const date = new Date(`${today()}T12:00:00`); date.setDate(date.getDate() - daysAgo); return date.toISOString().slice(0, 10) }
const productionSeed = (idValue: string, daysAgo: number, machine: string, hours: number, minutes: number, quantity: number, notes = ''): ProductionRecord => ({
  id: idValue, work_date: productionDate(daysAgo), product_code: 'OAS-025<SP263735>', machine,
  hours, minutes, total_minutes: hours * 60 + minutes, quantity, notes,
  created_at: new Date(Date.now() - daysAgo * 86_400_000 + (machine === 'MT03' ? 1000 : 0)).toISOString(),
})
const seedProductionRecords: ProductionRecord[] = [
  productionSeed('demo-r1', 0, 'MT03', 11, 37, 532),
  productionSeed('demo-r2', 0, 'MT02', 11, 43, 544),
  productionSeed('demo-r3', 1, 'MT03', 8, 19, 381),
  productionSeed('demo-r4', 1, 'MT02', 8, 12, 381),
  productionSeed('demo-r5', 2, 'MT03', 11, 1, 507),
  productionSeed('demo-r6', 2, 'MT02', 11, 25, 530),
  { ...productionSeed('demo-r7', 3, 'MT01', 7, 35, 420), product_code: 'AX-100<AL6061>' },
]
const seedProductionOptions: ProductionOption[] = [
  { id: 'demo-option-p1', option_type: 'product', value: 'OAS-025<SP263735>', created_at: new Date().toISOString() },
  { id: 'demo-option-p2', option_type: 'product', value: 'AX-100<AL6061>', created_at: new Date().toISOString() },
  { id: 'demo-option-m1', option_type: 'machine', value: 'MT01', created_at: new Date().toISOString() },
  { id: 'demo-option-m2', option_type: 'machine', value: 'MT02', created_at: new Date().toISOString() },
  { id: 'demo-option-m3', option_type: 'machine', value: 'MT03', created_at: new Date().toISOString() },
]

const normalizeCategory = (value: string): ProductCategory => PRODUCT_CATEGORIES.includes(value as ProductCategory) ? value as ProductCategory : '刀片'

/* 「庫存盤點」的「盤點備註日期與時間」：固定使用台灣時區，避免不同裝置時區造成紀錄時間不同。 */
const formatTaipeiDateTime = () => {
  const parts = new Intl.DateTimeFormat('zh-TW', {
    timeZone: 'Asia/Taipei', year: 'numeric', month: '2-digit', day: '2-digit',
    hour: '2-digit', minute: '2-digit', second: '2-digit', hourCycle: 'h23',
  }).formatToParts(new Date())
  const part = (type: Intl.DateTimeFormatPartTypes) => parts.find((item) => item.type === type)?.value || ''
  return `${part('year')}/${part('month')}/${part('day')} ${part('hour')}:${part('minute')}:${part('second')}`
}

function readLocal<T>(key: string, fallback: T): T {
  const value = dataStorage.getItem(key)
  if (!value) { dataStorage.setItem(key, JSON.stringify(fallback)); return fallback }
  try { return JSON.parse(value) as T } catch { return fallback }
}

function writeLocal(key: string, value: unknown) { dataStorage.setItem(key, JSON.stringify(value)) }
const id = () => crypto.randomUUID()

async function runAdminDelete(functionName: string, values: Record<string, unknown>): Promise<void> {
  const { data, error } = await supabase!.rpc(functionName, values)
  if (error) throw error
  const result = data as { ok?: boolean; error?: string } | null
  if (!result?.ok) throw new Error(result?.error || '刪除失敗')
}

export async function loadInventory(): Promise<{ products: Product[]; movements: Movement[] }> {
  if (isLocalDataMode) {
    const products = readLocal(PRODUCT_KEY, seedProducts).map((product) => ({ ...product, barcode: product.barcode || '', category: normalizeCategory(product.category) }))
    writeLocal(PRODUCT_KEY, products)
    return { products, movements: readLocal(MOVEMENT_KEY, seedMovements) }
  }
  const [productsResult, movementsResult] = await Promise.all([
    supabase!.from('products').select('*').order('code'),
    supabase!.from('stock_movements').select('*, product:products(code,barcode,name,unit,brand)').order('movement_date', { ascending: false }).order('created_at', { ascending: false }),
  ])
  if (productsResult.error) throw productsResult.error
  if (movementsResult.error) throw movementsResult.error
  return { products: (productsResult.data as Product[]).map((product) => ({ ...product, barcode: product.barcode || '' })), movements: movementsResult.data as unknown as Movement[] }
}

export async function loadProductMovements(productId: string): Promise<Movement[]> {
  if (isLocalDataMode) return readLocal(MOVEMENT_KEY, seedMovements).filter((movement) => movement.product_id === productId)
  const movements: Movement[] = []
  let lastId: string | undefined
  while (true) {
    const query = supabase!.from('stock_movements').select('*').eq('product_id', productId).order('id').limit(1000)
    if (lastId) query.gt('id', lastId)
    const { data, error } = await query
    if (error) throw error
    if (!data.length) return movements
    const page = data as Movement[]
    movements.push(...page)
    lastId = page[page.length - 1].id
  }
}

export async function loadProductChanges(productCode: string, createdAt: string): Promise<OperationLogEntry[]> {
  if (isLocalDataMode) return []
  const changes: OperationLogEntry[] = []
  let lastId: number | undefined
  while (true) {
    // 操作紀錄只保留刀具編號，以建立時間區分重新使用同一編號的刀具。
    const query = supabase!.from('operation_log').select('*').eq('target', productCode).gte('occurred_at', createdAt)
      .in('entity', ['products', 'stock_movements']).in('operation', ['UPDATE', 'DELETE']).order('id').limit(1000)
    if (lastId !== undefined) query.gt('id', lastId)
    const { data, error } = await query
    if (error) throw error
    if (!data.length) return changes.sort((a, b) => Date.parse(b.occurred_at) - Date.parse(a.occurred_at) || b.id.localeCompare(a.id, 'en', { numeric: true }))
    const page = data as StoredOperationLogEntry[]
    changes.push(...page.map((entry) => ({ ...entry, id: String(entry.id) })))
    lastId = page[page.length - 1].id
  }
}

export async function loadProductionRecords(): Promise<ProductionRecord[]> {
  if (isLocalDataMode) return readLocal(PRODUCTION_KEY, seedProductionRecords)
  const records: ProductionRecord[] = []
  let lastId: string | undefined
  // 由上一批最後的紀錄 ID 繼續讀取，避免分頁期間增刪資料造成位置位移。
  while (true) {
    const query = supabase!.from('production_records').select('*').order('id').limit(1000)
    if (lastId) query.gt('id', lastId)
    const { data, error } = await query
    if (error) throw error
    if (!data.length) break
    const page = data as ProductionRecord[]
    records.push(...page)
    lastId = page[page.length - 1].id
  }
  return records.sort((a, b) => b.work_date.localeCompare(a.work_date) || Date.parse(b.created_at) - Date.parse(a.created_at))
}

export async function loadProductionOptions(): Promise<ProductionOption[]> {
  if (isLocalDataMode) return readLocal(PRODUCTION_OPTION_KEY, seedProductionOptions)
  const { data, error } = await supabase!.from('production_options').select('*').order('option_type').order('value')
  if (error) throw error
  return data as ProductionOption[]
}

export async function createProductionOption(optionType: ProductionOptionType, rawValue: string, settings?: Pick<ProductionOption, 'order_quantity' | 'due_date'>): Promise<void> {
  const value = rawValue.trim().toUpperCase()
  if (!value) throw new Error(optionType === 'product' ? '產品編號不能空白' : '機台名稱不能空白')
  if (isLocalDataMode) {
    const options = readLocal(PRODUCTION_OPTION_KEY, seedProductionOptions)
    if (options.some((option) => option.option_type === optionType && option.value.toLowerCase() === value.toLowerCase())) throw new Error('這筆資料已經存在')
    writeLocal(PRODUCTION_OPTION_KEY, [...options, { id: id(), option_type: optionType, value, created_at: new Date().toISOString(), ...settings }])
    return
  }
  const userId = await requireAuthenticatedUserId()
  const { error } = await supabase!.from('production_options').insert({ user_id: userId, option_type: optionType, value, ...settings })
  if (error) {
    if (error.code === '23505') throw new Error('這筆資料已經存在')
    throw error
  }
}

export async function deleteProductionOption(optionType: ProductionOptionType, value: string): Promise<void> {
  if (isLocalDataMode) {
    const options = readLocal(PRODUCTION_OPTION_KEY, seedProductionOptions)
    const option = options.find((item) => item.option_type === optionType && item.value === value)
    if (optionType === 'machine' && option && readLocal<MachineServiceRecord[]>(MACHINE_SERVICE_KEY, []).some((record) => record.machine_id === option.id)) {
      throw new Error('機台已有保養維修紀錄，請到基本資料設為停用。')
    }
    writeLocal(PRODUCTION_OPTION_KEY, options.filter((option) => !(option.option_type === optionType && option.value === value)))
    if (optionType === 'machine' && option) {
      const profiles = readLocal<Record<string, MachineProfile>>(MACHINE_PROFILE_KEY, {})
      if (profiles[option.id]) {
        delete profiles[option.id]
        try { writeLocal(MACHINE_PROFILE_KEY, profiles) }
        catch (error) { writeLocal(PRODUCTION_OPTION_KEY, options); throw error }
      }
    }
    return
  }
  try {
    await runAdminDelete('admin_delete_production_option', { p_option_type: optionType, p_value: value })
  } catch (error) {
    if (optionType === 'machine' && (error as { code?: string })?.code === '23503') {
      throw new Error('機台已有保養維修紀錄，請到基本資料設為停用。')
    }
    throw error
  }
}

const machineProfileFields = 'name,brand,model,serial_number,purchased_on,started_on,notes'
const blankMachineProfile: MachineProfile = { name: '', brand: '', model: '', serial_number: '', purchased_on: null, started_on: null, notes: '' }

export async function loadMachineProfile(machineId: string): Promise<MachineProfile> {
  if (isLocalDataMode) return readLocal<Record<string, MachineProfile>>(MACHINE_PROFILE_KEY, {})[machineId] ?? blankMachineProfile
  const { data, error } = await supabase!.from('machine_profiles').select(machineProfileFields).eq('machine_id', machineId).maybeSingle()
  if (error) throw error
  return data ?? blankMachineProfile
}

export async function saveMachineProfile(machine: ProductionOption, rawCode: string, active: boolean, profile: MachineProfile): Promise<void> {
  const code = rawCode.trim().toUpperCase()
  if (!code) throw new Error('請輸入機台編號')
  if (isLocalDataMode) {
    const options = readLocal(PRODUCTION_OPTION_KEY, seedProductionOptions)
    const current = options.find((item) => item.id === machine.id && item.option_type === 'machine')
    if (!current || current.value !== machine.value) throw new Error('機台編號已變更或不存在，請重新開啟。')
    const records = readLocal(PRODUCTION_KEY, seedProductionRecords)
    const renamed = current.value !== code
    if (current.value.toUpperCase() !== code && (options.some((item) => item.option_type === 'machine' && item.value.toUpperCase() === code)
      || records.some((record) => record.machine.toUpperCase() === code))) throw new Error('這個機台編號已在使用或已有加工紀錄')
    const profiles = readLocal<Record<string, MachineProfile>>(MACHINE_PROFILE_KEY, {})
    writeLocal(PRODUCTION_OPTION_KEY, options.map((item) => item.id === machine.id ? { ...item, value: code, machine_active: active } : item))
    try {
      if (renamed) writeLocal(PRODUCTION_KEY, records.map((record) => record.machine.toUpperCase() === current.value.toUpperCase() ? { ...record, machine: code } : record))
      writeLocal(MACHINE_PROFILE_KEY, { ...profiles, [machine.id]: profile })
    } catch (error) {
      if (renamed) writeLocal(PRODUCTION_KEY, records)
      writeLocal(PRODUCTION_OPTION_KEY, options)
      throw error
    }
    return
  }
  const { error } = await supabase!.rpc('save_machine_profile', {
    p_machine_id: machine.id, p_current_code: machine.value, p_code: code, p_active: active, p_profile: profile,
  })
  if (error) throw error
}

export async function loadMachineServiceRecords(machineId: string): Promise<MachineServiceRecord[]> {
  if (isLocalDataMode) {
    return readLocal<MachineServiceRecord[]>(MACHINE_SERVICE_KEY, [])
      .filter((record) => record.machine_id === machineId)
      .sort((a, b) => b.service_date.localeCompare(a.service_date) || b.created_at.localeCompare(a.created_at) || b.id.localeCompare(a.id))
  }
  const records: MachineServiceRecord[] = []
  const pageSize = 500
  for (let start = 0; ; start += pageSize) {
    const { data, error } = await supabase!.from('machine_service_records')
      .select('id,machine_id,service_date,service_type,description,technician,cost,downtime_minutes,created_at')
      .eq('machine_id', machineId).order('service_date', { ascending: false }).order('created_at', { ascending: false })
      .order('id', { ascending: false }).range(start, start + pageSize - 1)
    if (error) throw error
    const page = data as MachineServiceRecord[]
    records.push(...page)
    if (page.length < pageSize) return records
  }
}

export async function saveMachineServiceRecord(machineId: string, input: MachineServiceInput, recordId: string | null): Promise<MachineServiceRecord> {
  if (!input.service_date || input.service_date > todayInTaipei() || !input.description.trim() ||
    (input.cost !== null && (!Number.isFinite(input.cost) || input.cost < 0 || input.cost > 999999999999.99)) ||
    (input.downtime_minutes !== null && (!Number.isInteger(input.downtime_minutes) || input.downtime_minutes < 0 || input.downtime_minutes > 2147483647))) {
    throw new Error('請檢查保養維修日期、處理內容、費用及停機分鐘。')
  }
  if (isLocalDataMode) {
    if (!readLocal(PRODUCTION_OPTION_KEY, seedProductionOptions).some((item) => item.id === machineId && item.option_type === 'machine')) {
      throw new Error('找不到這台機台')
    }
    const records = readLocal<MachineServiceRecord[]>(MACHINE_SERVICE_KEY, [])
    const existing = recordId ? records.find((record) => record.id === recordId && record.machine_id === machineId) : null
    if (recordId && !existing) throw new Error('找不到這筆保養維修紀錄')
    const record = { id: existing?.id ?? id(), machine_id: machineId, created_at: existing?.created_at ?? new Date().toISOString(), ...input }
    writeLocal(MACHINE_SERVICE_KEY, existing ? records.map((item) => item.id === record.id ? record : item) : [...records, record])
    return record
  }
  const { data, error } = await supabase!.rpc('save_machine_service_record', { p_machine_id: machineId, p_record_id: recordId, p_record: input })
  if (error) throw error
  return data as MachineServiceRecord
}

export async function updateProductionOrder(currentValue: string, nextValue: string, settings: Pick<ProductionOption, 'order_quantity' | 'due_date'>): Promise<void> {
  const current = currentValue.trim().toUpperCase()
  const next = nextValue.trim().toUpperCase()
  if (!current || !next) throw new Error('產品編號不能空白')
  if (isLocalDataMode) {
    const options = readLocal(PRODUCTION_OPTION_KEY, seedProductionOptions)
    const order = options.find((option) => option.option_type === 'product' && option.value.toUpperCase() === current)
    if (!order) throw new Error('找不到原本的訂單')
    const records = readLocal(PRODUCTION_KEY, seedProductionRecords)
    if (current !== next && (options.some((option) => option.option_type === 'product' && option.value.toUpperCase() === next)
      || records.some((record) => record.product_code.toUpperCase() === next))) throw new Error('新的產品編號已經存在')
    writeLocal(PRODUCTION_OPTION_KEY, options.map((option) => option.id === order.id ? { ...option, value: next, ...settings } : option))
    if (current !== next) {
      try {
        writeLocal(PRODUCTION_KEY, records.map((record) => record.product_code.toUpperCase() === current ? { ...record, product_code: next } : record))
      } catch (error) {
        writeLocal(PRODUCTION_OPTION_KEY, options)
        throw error
      }
    }
    return
  }
  const { error } = await supabase!.rpc('update_production_order', {
    p_current_value: current, p_new_value: next, p_order_quantity: settings.order_quantity ?? null, p_due_date: settings.due_date ?? null,
  })
  if (error) throw error
}

export async function saveCompletedProductionRecord(input: ProductionInput, completedOn: string, overwrite?: ProductionRecord): Promise<void> {
  const date = new Date(`${completedOn}T00:00:00Z`)
  if (!/^\d{4}-\d{2}-\d{2}$/.test(completedOn) || Number.isNaN(date.valueOf()) || date.toISOString().slice(0, 10) !== completedOn || completedOn > todayInTaipei()) {
    throw new Error('請選擇有效的完工日期，且不得晚於今天')
  }
  const productCode = input.product_code.trim().toUpperCase()
  if (isLocalDataMode) {
    const options = readLocal(PRODUCTION_OPTION_KEY, seedProductionOptions)
    const order = options.find((option) => option.option_type === 'product' && option.value.toUpperCase() === productCode)
    if (!order) throw new Error('找不到要登記完工的訂單')
    writeLocal(PRODUCTION_OPTION_KEY, options.map((option) => option.id === order.id ? { ...option, completed_on: completedOn } : option))
    try {
      if (overwrite) await updateProductionRecord(overwrite.id, input, overwrite)
      else await createProductionRecord(input)
    } catch (error) {
      writeLocal(PRODUCTION_OPTION_KEY, options)
      throw error
    }
    return
  }
  const { error } = await supabase!.rpc('save_completed_production_record', {
    p_input: input, p_completed_on: completedOn, p_expected: overwrite ?? null,
  })
  if (error) throw error
}

export async function findMatchingProductionRecord(input: Pick<ProductionInput, 'work_date' | 'product_code' | 'machine'>): Promise<ProductionRecord | null> {
  if (isLocalDataMode) {
    return readLocal(PRODUCTION_KEY, seedProductionRecords)
      .filter((record) => record.work_date === input.work_date
        && record.product_code.trim().toUpperCase() === input.product_code
        && record.machine.trim().toUpperCase() === input.machine)
      .sort((a, b) => Date.parse(b.created_at) - Date.parse(a.created_at) || b.id.localeCompare(a.id))[0] ?? null
  }
  const userId = await requireAuthenticatedUserId()
  const { data, error } = await supabase!.from('production_records').select('*')
    .eq('user_id', userId).eq('work_date', input.work_date).eq('product_code', input.product_code).eq('machine', input.machine)
    .order('created_at', { ascending: false }).order('id', { ascending: false }).limit(1).maybeSingle()
  if (error) throw error
  return data as ProductionRecord | null
}

export async function createProductionRecord(input: ProductionInput): Promise<void> {
  const record: ProductionRecord = {
    id: id(), ...input, product_code: input.product_code.trim().toUpperCase(), machine: input.machine.trim().toUpperCase(),
    total_minutes: input.hours * 60 + input.minutes, notes: input.notes.trim(), created_at: new Date().toISOString(),
  }
  if (isLocalDataMode) {
    writeLocal(PRODUCTION_KEY, [record, ...readLocal(PRODUCTION_KEY, seedProductionRecords)])
    return
  }
  const { id: _id, created_at: _createdAt, total_minutes: _totalMinutes, ...values } = record
  const userId = await requireAuthenticatedUserId()
  const { error } = await supabase!.from('production_records').insert({ ...values, user_id: userId })
  if (error) throw error
}

export async function updateProductionRecord(recordId: string, input: ProductionInput, expected?: ProductionInput): Promise<void> {
  const values = {
    ...input,
    product_code: input.product_code.trim().toUpperCase(),
    machine: input.machine.trim().toUpperCase(),
    notes: input.notes.trim(),
  }
  const expectedFields = expected && {
    work_date: expected.work_date,
    product_code: expected.product_code,
    machine: expected.machine,
    hours: expected.hours,
    minutes: expected.minutes,
    quantity: expected.quantity,
    notes: expected.notes,
  }
  const conflictMessage = '原登記已變更或刪除，請取消後重新送出。'
  if (isLocalDataMode) {
    const records = readLocal(PRODUCTION_KEY, seedProductionRecords)
    const current = records.find((record) => record.id === recordId)
    if (expectedFields && (!current || Object.entries(expectedFields).some(([key, value]) => current[key as keyof ProductionInput] !== value))) {
      throw new Error(conflictMessage)
    }
    writeLocal(PRODUCTION_KEY, records.map((record) => record.id === recordId
      ? { ...record, ...values, total_minutes: values.hours * 60 + values.minutes }
      : record))
    return
  }
  const userId = await requireAuthenticatedUserId()
  const query = supabase!
    .from('production_records')
    .update(values)
    .eq('id', recordId)
    .eq('user_id', userId)
  if (expectedFields) query.match(expectedFields)
  const { data, error } = await query.select('id').maybeSingle()
  if (error) throw error
  if (!data) throw new Error(expectedFields ? conflictMessage : '找不到這筆加工紀錄或沒有修改權限')
}

export async function deleteProductionRecord(recordId: string): Promise<void> {
  if (isLocalDataMode) {
    writeLocal(PRODUCTION_KEY, readLocal(PRODUCTION_KEY, seedProductionRecords).filter((record) => record.id !== recordId))
    return
  }
  await runAdminDelete('admin_delete_production_record', { p_record_id: recordId })
}

export async function deleteProductionProductData(productCode: string): Promise<void> {
  const normalized = productCode.trim().toUpperCase()
  if (!normalized) throw new Error('請先選擇產品編號')
  if (isLocalDataMode) {
    writeLocal(PRODUCTION_KEY, readLocal(PRODUCTION_KEY, seedProductionRecords).filter((record) => record.product_code.toUpperCase() !== normalized))
    writeLocal(PRODUCTION_OPTION_KEY, readLocal(PRODUCTION_OPTION_KEY, seedProductionOptions).filter((option) => !(option.option_type === 'product' && option.value.toUpperCase() === normalized)))
    return
  }
  await runAdminDelete('admin_delete_production_product_data', { p_product_code: normalized })
}

export async function createProduct(input: ProductInput): Promise<void> {
  if (isLocalDataMode) {
    const products = readLocal(PRODUCT_KEY, seedProducts)
    const movements = readLocal(MOVEMENT_KEY, seedMovements)
    const now = new Date().toISOString()
    const prefix = codePrefix(input.category)
    const nextNumber = products.filter((item) => item.category === input.category).reduce((max, item) => Math.max(max, Number(item.code.slice(1)) || 0), 0) + 1
    const product: Product = { id: id(), code: `${prefix}${String(nextNumber).padStart(4, '0')}`, name: composeProductName(input.category, input.specification, input.size, input.chipbreaker, input.material), ...input, barcode: input.barcode.trim(), stock_quantity: input.initial_quantity || 0, unit_price: input.initial_unit_price || 0, purchased_at: input.purchased_at || null, created_at: now, updated_at: now }
    const { initial_quantity, initial_unit_price, ...stored } = product as Product & { initial_quantity?: number; initial_unit_price?: number }
    writeLocal(PRODUCT_KEY, [...products, stored])
    if (input.initial_quantity > 0) movements.unshift({ id: id(), product_id: product.id, movement_type: 'in', quantity: input.initial_quantity, movement_date: input.purchased_at, unit_price: input.initial_unit_price, supplier: input.supplier, notes: '建立時初庫存', stock_after: input.initial_quantity, request_key: id(), created_at: now, product: { code: product.code, name: product.name, unit: product.unit, brand: product.brand } })
    writeLocal(MOVEMENT_KEY, movements)
    return
  }
  const barcode = input.barcode.trim()
  if (barcode) {
    const { data: existingBarcode, error: barcodeLookupError } = await supabase!.from('products').select('id').ilike('barcode', barcode).maybeSingle()
    if (barcodeLookupError) throw barcodeLookupError
    if (existingBarcode) throw new Error('這個條碼已綁定其他刀具')
  }
  const { data: productId, error } = await supabase!.rpc('create_product_with_initial_stock', {
    p_specification: input.specification, p_size: input.size, p_chipbreaker: input.chipbreaker || null,
    p_material: input.material || null, p_brand: input.brand || null, p_supplier: input.supplier || null,
    p_category: input.category, p_unit: input.unit, p_notes: input.notes || null,
    p_initial_quantity: input.initial_quantity || 0, p_initial_unit_price: input.initial_unit_price || 0,
    p_purchased_at: input.purchased_at || today(), p_request_key: id(),
  })
  if (error) throw error
  if (barcode && productId) {
    const { error: barcodeError } = await supabase!.from('products').update({ barcode }).eq('id', productId)
    if (barcodeError) throw barcodeError.code === '23505' ? new Error('這個條碼已綁定其他刀具') : barcodeError
  }
}

export async function updateProduct(productId: string, input: ProductInput): Promise<void> {
  const changes = { barcode: input.barcode.trim(), specification: input.specification, size: input.size, chipbreaker: input.chipbreaker, material: input.material, name: composeProductName(input.category, input.specification, input.size, input.chipbreaker, input.material), brand: input.brand, supplier: input.supplier, category: input.category, unit: input.unit, notes: input.notes }
  if (isLocalDataMode) {
    const products = readLocal(PRODUCT_KEY, seedProducts).map((p) => p.id === productId ? { ...p, ...changes, updated_at: new Date().toISOString() } : p)
    writeLocal(PRODUCT_KEY, products); return
  }
  const { error } = await supabase!.from('products').update(changes).eq('id', productId)
  if (error) throw error.code === '23505' ? new Error('這個條碼已綁定其他刀具') : error
}

export async function updateProductBarcode(productId: string, rawBarcode: string, expectedBarcode: string): Promise<void> {
  const barcode = rawBarcode.trim()
  if (!barcode) throw new Error('條碼不能空白')
  const conflictMessage = '刀具條碼已變更或刀具已移除，請重新整理後再掃描綁定。'
  if (isLocalDataMode) {
    const products = readLocal(PRODUCT_KEY, seedProducts)
    const product = products.find((item) => item.id === productId)
    if (!product || (product.barcode || '') !== expectedBarcode) throw new Error(conflictMessage)
    if (products.some((product) => product.id !== productId && product.barcode?.toLowerCase() === barcode.toLowerCase())) throw new Error('這個條碼已綁定其他刀具')
    writeLocal(PRODUCT_KEY, products.map((product) => product.id === productId ? { ...product, barcode, updated_at: new Date().toISOString() } : product))
    return
  }
  const userId = await requireAuthenticatedUserId()
  const { data, error } = await supabase!.from('products').update({ barcode })
    .eq('id', productId).eq('user_id', userId).eq('barcode', expectedBarcode).select('id').maybeSingle()
  if (error) throw error.code === '23505' ? new Error('這個條碼已綁定其他刀具') : error
  if (!data) throw new Error(conflictMessage)
}

export async function deleteProduct(productId: string): Promise<void> {
  if (isLocalDataMode) {
    writeLocal(PRODUCT_KEY, readLocal(PRODUCT_KEY, seedProducts).filter((p) => p.id !== productId))
    writeLocal(MOVEMENT_KEY, readLocal(MOVEMENT_KEY, seedMovements).filter((m) => m.product_id !== productId)); return
  }
  await runAdminDelete('admin_delete_product', { p_product_id: productId })
}

export async function registerMovement(input: MovementInput): Promise<void> {
  if (isLocalDataMode) {
    const products = readLocal(PRODUCT_KEY, seedProducts)
    const product = products.find((p) => p.id === input.product_id)
    if (!product) throw new Error('找不到商品')
    const next = product.stock_quantity + (input.movement_type === 'in' ? input.quantity : -input.quantity)
    if (next < 0) throw new Error(`庫存不足，目前只有 ${product.stock_quantity} ${product.unit}`)
    const updated = { ...product, stock_quantity: next, unit_price: input.movement_type === 'in' ? input.unit_price : product.unit_price, supplier: input.movement_type === 'in' && input.supplier ? input.supplier : product.supplier, purchased_at: input.movement_type === 'in' ? input.movement_date : product.purchased_at, updated_at: new Date().toISOString() }
    writeLocal(PRODUCT_KEY, products.map((p) => p.id === product.id ? updated : p))
    const movement: Movement = { id: id(), ...input, unit_price: input.movement_type === 'in' ? input.unit_price : null, stock_after: next, created_at: new Date().toISOString(), product: { code: product.code, name: product.name, unit: product.unit, brand: product.brand } }
    writeLocal(MOVEMENT_KEY, [movement, ...readLocal(MOVEMENT_KEY, seedMovements)]); return
  }
  const { error } = await supabase!.rpc('register_stock_movement', {
    p_product_id: input.product_id, p_movement_type: input.movement_type, p_quantity: input.quantity,
    p_movement_date: input.movement_date, p_unit_price: input.movement_type === 'in' ? input.unit_price : null,
    p_supplier: input.supplier || null, p_notes: input.notes || null, p_request_key: input.request_key,
  })
  if (error) throw error
}

export async function setStocktakeQuantity(input: StocktakeInput): Promise<void> {
  if (!Number.isInteger(input.actual_quantity) || input.actual_quantity < 0) throw new Error('實際盤點數量必須是 0 以上的整數')
  const userNotes = input.notes.trim()
  const timestampedNotes = `盤點時間：${formatTaipeiDateTime()}${userNotes ? `｜${userNotes}` : ''}`
  if (isLocalDataMode) {
    const products = readLocal(PRODUCT_KEY, seedProducts)
    const product = products.find((item) => item.id === input.product_id)
    if (!product) throw new Error('找不到刀具資料')
    const difference = input.actual_quantity - product.stock_quantity
    if (difference === 0) return
    const now = new Date().toISOString()
    const movement: Movement = {
      id: id(), product_id: product.id, movement_type: difference > 0 ? 'in' : 'out', quantity: Math.abs(difference),
      movement_date: today(), unit_price: difference > 0 ? product.unit_price : null,
      supplier: difference > 0 ? product.supplier : '',
      notes: `庫存盤點調整：${product.stock_quantity} → ${input.actual_quantity}｜${timestampedNotes}`,
      stock_after: input.actual_quantity, request_key: input.request_key, created_at: now,
      product: { code: product.code, name: product.name, unit: product.unit, brand: product.brand },
    }
    writeLocal(PRODUCT_KEY, products.map((item) => item.id === product.id ? { ...item, stock_quantity: input.actual_quantity, updated_at: now } : item))
    writeLocal(MOVEMENT_KEY, [movement, ...readLocal(MOVEMENT_KEY, seedMovements)])
    return
  }
  const { error } = await supabase!.rpc('set_stocktake_quantity', {
    p_product_id: input.product_id,
    p_actual_quantity: input.actual_quantity,
    p_notes: timestampedNotes,
    p_request_key: input.request_key,
  })
  if (error) throw error
}

export async function deleteMovement(movementId: string): Promise<void> {
  if (isLocalDataMode) {
    const products = readLocal(PRODUCT_KEY, seedProducts)
    const movements = readLocal(MOVEMENT_KEY, seedMovements)
    const target = movements.find((movement) => movement.id === movementId)
    if (!target) throw new Error('找不到這筆出入庫紀錄')
    const remaining = movements.filter((movement) => movement.id !== movementId)
    const productMovements = remaining
      .filter((movement) => movement.product_id === target.product_id)
      .sort((a, b) => `${a.movement_date}-${a.created_at}-${a.id}`.localeCompare(`${b.movement_date}-${b.created_at}-${b.id}`))
    let balance = 0
    for (const movement of productMovements) {
      balance += movement.movement_type === 'in' ? movement.quantity : -movement.quantity
      if (balance < 0) throw new Error('無法刪除：這會使後續某筆歷史庫存變成負數')
      movement.stock_after = balance
    }
    const latestInbound = [...productMovements].reverse().find((movement) => movement.movement_type === 'in')
    const updatedProducts = products.map((product) => product.id === target.product_id ? {
      ...product,
      stock_quantity: balance,
      unit_price: latestInbound?.unit_price || 0,
      purchased_at: latestInbound?.movement_date || null,
      supplier: latestInbound?.supplier || product.supplier,
      updated_at: new Date().toISOString(),
    } : product)
    writeLocal(PRODUCT_KEY, updatedProducts)
    writeLocal(MOVEMENT_KEY, remaining)
    return
  }
  await runAdminDelete('admin_delete_stock_movement', { p_movement_id: movementId })
}
