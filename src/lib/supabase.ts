import { createClient } from '@supabase/supabase-js'

const url = import.meta.env.VITE_SUPABASE_URL?.trim()
const key = import.meta.env.VITE_SUPABASE_ANON_KEY?.trim()

export const isDemoMode = !url || !key || url.includes('YOUR_PROJECT') || key.includes('YOUR_SUPABASE')
export const TEST_MODE_KEY = 'cnc-cloud-test-mode'
export const isTestMode = window.sessionStorage.getItem(TEST_MODE_KEY) === '1'
export const isLocalDataMode = isDemoMode || isTestMode
export const dataStorage = isTestMode ? window.sessionStorage : window.localStorage

// 登入只保留在目前分頁：重新整理不會登出，但關閉分頁／瀏覽器後必須重新登入。
const sessionStorageKey = url ? `sb-${new URL(url).hostname.split('.')[0]}-auth-token` : 'cnc-cloud-auth-token'
if (!isLocalDataMode && typeof window !== 'undefined') {
  // 清除舊版曾寫入 localStorage 的長期登入資料，避免更新後仍沿用永久登入。
  window.localStorage.removeItem(sessionStorageKey)
}

// 測試畫面使用分頁資料，整頁載入時即停用 Supabase，包括登入與即時同步。
export const supabase = isLocalDataMode ? null : createClient(url!, key!, {
  auth: {
    experimental: { passkey: true },
    persistSession: true,
    storage: window.sessionStorage,
    storageKey: sessionStorageKey,
    autoRefreshToken: true,
    detectSessionInUrl: true,
  },
})
