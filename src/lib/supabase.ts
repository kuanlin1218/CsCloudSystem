import { createClient } from '@supabase/supabase-js'

const url = import.meta.env.VITE_SUPABASE_URL?.trim()
const key = import.meta.env.VITE_SUPABASE_ANON_KEY?.trim()

export const isDemoMode = !url || !key || url.includes('YOUR_PROJECT') || key.includes('YOUR_SUPABASE')

// 登入只保留在目前分頁：重新整理不會登出，但關閉分頁／瀏覽器後必須重新登入。
const sessionStorageKey = url ? `sb-${new URL(url).hostname.split('.')[0]}-auth-token` : 'cnc-cloud-auth-token'
if (!isDemoMode && typeof window !== 'undefined') {
  // 清除舊版曾寫入 localStorage 的長期登入資料，避免更新後仍沿用永久登入。
  window.localStorage.removeItem(sessionStorageKey)
}

export const supabase = isDemoMode ? null : createClient(url!, key!, {
  auth: {
    persistSession: true,
    storage: window.sessionStorage,
    storageKey: sessionStorageKey,
    autoRefreshToken: true,
    detectSessionInUrl: true,
  },
})
