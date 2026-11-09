import { isDemoMode, supabase } from './supabase'

export const OWNER_USER_ID = '5568706d-caaa-459b-b842-13cdbc3929fc'
export const OWNER_USERNAME = 'kuanlin'
export const OWNER_EMAIL = 'kuanlin@chyuansheu.com'

export function isOwnerUser(userId?: string | null) {
  return userId === OWNER_USER_ID
}

export async function requireAuthenticatedUserId(): Promise<string> {
  if (isDemoMode || !supabase) throw new Error('目前未連接 Supabase 雲端資料')
  const { data, error } = await supabase.auth.getUser()
  if (error || !data.user || !isOwnerUser(data.user.id)) throw new Error('登入狀態已失效，請重新登入')
  return data.user.id
}

export async function verifyOwnerPassword(password: string): Promise<void> {
  if (isDemoMode || !supabase) throw new Error('目前未連接 Supabase 雲端資料')
  if (!/^\d{4,}$/.test(password)) throw new Error('登入密碼必須是至少 4 位數字')
  const { data, error } = await supabase.auth.signInWithPassword({ email: OWNER_EMAIL, password })
  if (error) throw new Error('登入密碼不正確')
  if (!isOwnerUser(data.user?.id)) {
    await supabase.auth.signOut()
    throw new Error('這個帳號沒有系統使用權限')
  }
}
