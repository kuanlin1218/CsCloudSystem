export type MovementType = 'in' | 'out'
export type Page = 'home' | 'production-entry' | 'production-query' | 'dashboard' | 'movements' | 'stocktake' | 'products'
export type ProductCategory = '刀片' | '鑽銑刀' | '刀柄' | '其他'
export const PRODUCT_CATEGORIES: ProductCategory[] = ['刀片', '鑽銑刀', '刀柄', '其他']

export interface Product {
  id: string
  code: string
  barcode: string
  name: string
  specification: string
  size: string
  chipbreaker: string
  material: string
  brand: string
  supplier: string
  category: ProductCategory
  unit: string
  stock_quantity: number
  unit_price: number
  purchased_at: string | null
  notes: string
  created_at: string
  updated_at: string
}

export interface Movement {
  id: string
  product_id: string
  movement_type: MovementType
  quantity: number
  movement_date: string
  unit_price: number | null
  supplier: string
  notes: string
  stock_after: number
  request_key: string
  created_at: string
  product?: Pick<Product, 'code' | 'name' | 'unit' | 'brand'>
}

export interface ProductInput {
  barcode: string
  specification: string
  size: string
  chipbreaker: string
  material: string
  brand: string
  supplier: string
  category: ProductCategory
  unit: string
  notes: string
  initial_quantity: number
  initial_unit_price: number
  purchased_at: string
}

export interface MovementInput {
  product_id: string
  movement_type: MovementType
  quantity: number
  movement_date: string
  unit_price: number
  supplier: string
  notes: string
  request_key: string
}

export interface StocktakeInput {
  product_id: string
  actual_quantity: number
  notes: string
  request_key: string
}

export interface ProductionRecord {
  id: string
  work_date: string
  product_code: string
  machine: string
  hours: number
  minutes: number
  total_minutes: number
  quantity: number
  notes: string
  created_at: string
}

export interface ProductionInput {
  work_date: string
  product_code: string
  machine: string
  hours: number
  minutes: number
  quantity: number
  notes: string
}

export type ProductionOptionType = 'product' | 'machine'

export interface ProductionOption {
  id: string
  option_type: ProductionOptionType
  value: string
  created_at: string
}
