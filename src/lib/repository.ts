import { supabase, isDemoMode } from './supabase'
import { composeProductName, today } from './format'
import { PRODUCT_CATEGORIES, type Movement, type MovementInput, type Product, type ProductCategory, type ProductInput, type ProductionInput, type ProductionOption, type ProductionOptionType, type ProductionRecord, type StocktakeInput } from '../types'

const PRODUCT_KEY = 'blade_inventory_products_v3'
const MOVEMENT_KEY = 'blade_inventory_movements_v3'
const PRODUCTION_KEY = 'cnc_production_records_v1'
const PRODUCTION_OPTION_KEY = 'cnc_production_options_v1'
const previousDay = (() => { const date = new Date(`${today()}T12:00:00`); date.setDate(date.getDate() - 1); return date.toISOString().slice(0, 10) })()
const codePrefix = (category: ProductCategory) => ({ '刀片': 'T', '鑽銑刀': 'M', '刀柄': 'H', '其他': 'O' }[category])

const seedProducts: Product[] = [
  { id: 'demo-p1', code: 'T0001', name: 'WNMG-080408-MS-PR1535', specification: 'WNMG', size: '080408', chipbreaker: 'MS', material: 'PR1535', brand: '京瓷', supplier: '福燕', category: '刀片', unit: '片', stock_quantity: 190, unit_price: 200, purchased_at: today(), notes: '鋼件・通用加工', created_at: new Date().toISOString(), updated_at: new Date().toISOString() },
  { id: 'demo-p2', code: 'T0002', name: 'TNMG-160404-無-VP15TF', specification: 'TNMG', size: '160404', chipbreaker: '無', material: 'VP15TF', brand: '三菱', supplier: '怡立', category: '刀片', unit: '片', stock_quantity: 100, unit_price: 100, purchased_at: today(), notes: '', created_at: new Date().toISOString(), updated_at: new Date().toISOString() },
  { id: 'demo-p3', code: 'T0003', name: 'VNMG-160404-GU-AC8025P', specification: 'VNMG', size: '160404', chipbreaker: 'GU', material: 'AC8025P', brand: '住友', supplier: '益詮', category: '刀片', unit: '片', stock_quantity: 100, unit_price: 165, purchased_at: today(), notes: '', created_at: new Date().toISOString(), updated_at: new Date().toISOString() },
  { id: 'demo-p4', code: 'M0001', name: 'EMM4100-10-鎢鋼', specification: 'EMM4100', size: '10', chipbreaker: '', material: '鎢鋼', brand: '益詮', supplier: '益詮', category: '鑽銑刀', unit: '支', stock_quantity: 5, unit_price: 800, purchased_at: today(), notes: '立銑刀', created_at: new Date().toISOString(), updated_at: new Date().toISOString() },
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

function readLocal<T>(key: string, fallback: T): T {
  const value = localStorage.getItem(key)
  if (!value) { localStorage.setItem(key, JSON.stringify(fallback)); return fallback }
  try { return JSON.parse(value) as T } catch { return fallback }
}

function writeLocal(key: string, value: unknown) { localStorage.setItem(key, JSON.stringify(value)) }
const id = () => crypto.randomUUID()

export async function loadInventory(): Promise<{ products: Product[]; movements: Movement[] }> {
  if (isDemoMode) {
    const products = readLocal(PRODUCT_KEY, seedProducts).map((product) => ({ ...product, category: normalizeCategory(product.category) }))
    writeLocal(PRODUCT_KEY, products)
    return { products, movements: readLocal(MOVEMENT_KEY, seedMovements) }
  }
  const [productsResult, movementsResult] = await Promise.all([
    supabase!.from('products').select('*').order('code'),
    supabase!.from('stock_movements').select('*, product:products(code,name,unit,brand)').order('movement_date', { ascending: false }).order('created_at', { ascending: false }),
  ])
  if (productsResult.error) throw productsResult.error
  if (movementsResult.error) throw movementsResult.error
  return { products: productsResult.data as Product[], movements: movementsResult.data as unknown as Movement[] }
}

export async function loadProductionRecords(): Promise<ProductionRecord[]> {
  if (isDemoMode) return readLocal(PRODUCTION_KEY, seedProductionRecords)
  const { data, error } = await supabase!.from('production_records').select('*').order('work_date', { ascending: false }).order('created_at', { ascending: false })
  if (error) throw error
  return data as ProductionRecord[]
}

export async function loadProductionOptions(): Promise<ProductionOption[]> {
  if (isDemoMode) return readLocal(PRODUCTION_OPTION_KEY, seedProductionOptions)
  const { data, error } = await supabase!.from('production_options').select('id,option_type,value,created_at').order('option_type').order('value')
  if (error) throw error
  return data as ProductionOption[]
}

export async function createProductionOption(optionType: ProductionOptionType, rawValue: string): Promise<void> {
  const value = rawValue.trim().toUpperCase()
  if (!value) throw new Error(optionType === 'product' ? '產品編號不能空白' : '機台名稱不能空白')
  if (isDemoMode) {
    const options = readLocal(PRODUCTION_OPTION_KEY, seedProductionOptions)
    if (options.some((option) => option.option_type === optionType && option.value.toLowerCase() === value.toLowerCase())) throw new Error('這筆資料已經存在')
    writeLocal(PRODUCTION_OPTION_KEY, [...options, { id: id(), option_type: optionType, value, created_at: new Date().toISOString() }])
    return
  }
  const { error } = await supabase!.from('production_options').insert({ user_id: null, option_type: optionType, value })
  if (error) {
    if (error.code === '23505') throw new Error('這筆資料已經存在')
    throw error
  }
}

export async function deleteProductionOption(optionType: ProductionOptionType, value: string): Promise<void> {
  if (isDemoMode) {
    const options = readLocal(PRODUCTION_OPTION_KEY, seedProductionOptions)
    writeLocal(PRODUCTION_OPTION_KEY, options.filter((option) => !(option.option_type === optionType && option.value === value)))
    return
  }
  const { error } = await supabase!.from('production_options').delete().eq('option_type', optionType).eq('value', value)
  if (error) throw error
}

export async function createProductionRecord(input: ProductionInput): Promise<void> {
  const record: ProductionRecord = {
    id: id(), ...input, product_code: input.product_code.trim().toUpperCase(), machine: input.machine.trim().toUpperCase(),
    total_minutes: input.hours * 60 + input.minutes, notes: input.notes.trim(), created_at: new Date().toISOString(),
  }
  if (isDemoMode) {
    writeLocal(PRODUCTION_KEY, [record, ...readLocal(PRODUCTION_KEY, seedProductionRecords)])
    return
  }
  const { id: _id, created_at: _createdAt, total_minutes: _totalMinutes, ...values } = record
  const { error } = await supabase!.from('production_records').insert({ ...values, user_id: null })
  if (error) throw error
}

export async function deleteProductionRecord(recordId: string): Promise<void> {
  if (isDemoMode) {
    writeLocal(PRODUCTION_KEY, readLocal(PRODUCTION_KEY, seedProductionRecords).filter((record) => record.id !== recordId))
    return
  }
  const { error } = await supabase!.from('production_records').delete().eq('id', recordId)
  if (error) throw error
}

export async function deleteProductionProductData(productCode: string): Promise<void> {
  const normalized = productCode.trim().toUpperCase()
  if (!normalized) throw new Error('請先選擇產品編號')
  if (isDemoMode) {
    writeLocal(PRODUCTION_KEY, readLocal(PRODUCTION_KEY, seedProductionRecords).filter((record) => record.product_code.toUpperCase() !== normalized))
    writeLocal(PRODUCTION_OPTION_KEY, readLocal(PRODUCTION_OPTION_KEY, seedProductionOptions).filter((option) => !(option.option_type === 'product' && option.value.toUpperCase() === normalized)))
    return
  }
  const { error } = await supabase!.rpc('delete_production_product_data', { p_product_code: normalized })
  if (error) throw error
}

export async function createProduct(input: ProductInput): Promise<void> {
  if (isDemoMode) {
    const products = readLocal(PRODUCT_KEY, seedProducts)
    const movements = readLocal(MOVEMENT_KEY, seedMovements)
    const now = new Date().toISOString()
    const prefix = codePrefix(input.category)
    const nextNumber = products.filter((item) => item.category === input.category).reduce((max, item) => Math.max(max, Number(item.code.slice(1)) || 0), 0) + 1
    const product: Product = { id: id(), code: `${prefix}${String(nextNumber).padStart(4, '0')}`, name: composeProductName(input.category, input.specification, input.size, input.chipbreaker, input.material), ...input, stock_quantity: input.initial_quantity || 0, unit_price: input.initial_unit_price || 0, purchased_at: input.purchased_at || null, created_at: now, updated_at: now }
    const { initial_quantity, initial_unit_price, ...stored } = product as Product & { initial_quantity?: number; initial_unit_price?: number }
    writeLocal(PRODUCT_KEY, [...products, stored])
    if (input.initial_quantity > 0) movements.unshift({ id: id(), product_id: product.id, movement_type: 'in', quantity: input.initial_quantity, movement_date: input.purchased_at, unit_price: input.initial_unit_price, supplier: input.supplier, notes: '建立商品時的期初庫存', stock_after: input.initial_quantity, request_key: id(), created_at: now, product: { code: product.code, name: product.name, unit: product.unit, brand: product.brand } })
    writeLocal(MOVEMENT_KEY, movements)
    return
  }
  const { error } = await supabase!.rpc('create_product_with_initial_stock', {
    p_specification: input.specification, p_size: input.size, p_chipbreaker: input.chipbreaker || null,
    p_material: input.material || null, p_brand: input.brand || null, p_supplier: input.supplier || null,
    p_category: input.category, p_unit: input.unit, p_notes: input.notes || null,
    p_initial_quantity: input.initial_quantity || 0, p_initial_unit_price: input.initial_unit_price || 0,
    p_purchased_at: input.purchased_at || today(), p_request_key: id(),
  })
  if (error) throw error
}

export async function updateProduct(productId: string, input: ProductInput): Promise<void> {
  const changes = { specification: input.specification, size: input.size, chipbreaker: input.chipbreaker, material: input.material, name: composeProductName(input.category, input.specification, input.size, input.chipbreaker, input.material), brand: input.brand, supplier: input.supplier, category: input.category, unit: input.unit, notes: input.notes }
  if (isDemoMode) {
    const products = readLocal(PRODUCT_KEY, seedProducts).map((p) => p.id === productId ? { ...p, ...changes, updated_at: new Date().toISOString() } : p)
    writeLocal(PRODUCT_KEY, products); return
  }
  const { error } = await supabase!.from('products').update(changes).eq('id', productId)
  if (error) throw error
}

export async function deleteProduct(productId: string): Promise<void> {
  if (isDemoMode) {
    writeLocal(PRODUCT_KEY, readLocal(PRODUCT_KEY, seedProducts).filter((p) => p.id !== productId))
    writeLocal(MOVEMENT_KEY, readLocal(MOVEMENT_KEY, seedMovements).filter((m) => m.product_id !== productId)); return
  }
  const { error } = await supabase!.from('products').delete().eq('id', productId)
  if (error) throw error
}

export async function registerMovement(input: MovementInput): Promise<void> {
  if (isDemoMode) {
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
  if (isDemoMode) {
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
      notes: `庫存盤點調整：${product.stock_quantity} → ${input.actual_quantity}${input.notes.trim() ? `｜${input.notes.trim()}` : ''}`,
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
    p_notes: input.notes.trim() || null,
    p_request_key: input.request_key,
  })
  if (error) throw error
}

export async function deleteMovement(movementId: string): Promise<void> {
  if (isDemoMode) {
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
  const { error } = await supabase!.rpc('delete_stock_movement', { p_movement_id: movementId })
  if (error) throw error
}
