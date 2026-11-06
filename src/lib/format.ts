export const today = () => {
  const now = new Date()
  const offset = now.getTimezoneOffset()
  return new Date(now.getTime() - offset * 60_000).toISOString().slice(0, 10)
}

export const formatDate = (value: string) => value ? value.replaceAll('-', '.') : '—'
export const formatMoney = (value: number) => new Intl.NumberFormat('zh-TW', { maximumFractionDigits: 0 }).format(value)
export const formatNumber = (value: number) => new Intl.NumberFormat('zh-TW').format(value)
import type { ProductCategory } from '../types'

export const composeName = (parts: Array<string | undefined>) => parts.map((v) => v?.trim()).filter(Boolean).join('-')

export const composeProductName = (
  category: ProductCategory,
  specification: string,
  size: string,
  chipbreaker: string,
  material: string,
) => {
  if (category === '刀片') return composeName([specification, size, chipbreaker, material])
  if (category === '鑽銑刀') return composeName([specification, size, material])
  return specification.trim()
}
