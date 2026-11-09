import { isDemoMode, supabase } from './supabase'

export interface AdminProductRow {
  code: string
  category: string
  specification: string
  size: string
  chipbreaker: string
  material: string
  name?: string
  brand: string
  supplier: string
  unit: string
  stock_quantity: number
  unit_price: number
  purchased_at: string | null
  notes: string
}

export interface AdminMovementRow {
  sequence?: number
  product_code: string
  movement_type: 'in' | 'out'
  movement_date: string
  quantity: number
  unit_price: number | null
  supplier: string
  notes: string
}

export interface AdminProductionRow {
  work_date: string
  product_code: string
  machine: string
  hours: number
  minutes: number
  quantity: number
  notes: string
}

export interface AdminOptionRow {
  option_type: 'product' | 'machine'
  value: string
}

export interface AdminDataPayload {
  products: AdminProductRow[]
  movements: AdminMovementRow[]
  production_records: AdminProductionRow[]
  production_options: AdminOptionRow[]
}

export interface AdminExportResult {
  ok: boolean
  error?: string
  exported_at?: string
  data?: AdminDataPayload
}

export interface AdminImportResult {
  ok: boolean
  error?: string
  counts?: {
    products: number
    movements: number
    production_records: number
    production_options: number
  }
}

function requireCloud() {
  if (isDemoMode || !supabase) throw new Error('此功能只適用於已連接 Supabase 的雲端資料')
  return supabase
}

export async function exportAllAdminData(password: string): Promise<Required<Pick<AdminExportResult, 'exported_at' | 'data'>>> {
  const client = requireCloud()
  const { data, error } = await client.rpc('admin_export_all', { p_password: password })
  if (error) throw error
  const result = data as AdminExportResult
  if (!result?.ok || !result.data) throw new Error(result?.error || '無法匯出雲端資料')
  return { exported_at: result.exported_at || new Date().toISOString(), data: result.data }
}

export async function importAllAdminData(password: string, payload: AdminDataPayload): Promise<AdminImportResult['counts']> {
  const client = requireCloud()
  const { data, error } = await client.rpc('admin_import_all', { p_password: password, p_payload: payload })
  if (error) throw error
  const result = data as AdminImportResult
  if (!result?.ok || !result.counts) throw new Error(result?.error || '無法匯入雲端資料')
  return result.counts
}
