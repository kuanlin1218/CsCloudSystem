import { isDemoMode, supabase } from './supabase'

export const OWNER_USER_ID = '5568706d-caaa-459b-b842-13cdbc3929fc'
export const OWNER_USERNAME = 'kuanlin'
export const OWNER_EMAIL = 'kuanlin@chyuansheu.com'

export function isOwnerUser(userId?: string | null) {
  return userId === OWNER_USER_ID
}

export function supportsPasskeys() {
  return Boolean(supabase && window.isSecureContext && window.PublicKeyCredential && navigator.credentials)
}

function passkeyErrorMessage(error: { code?: string; name: string; message: string }) {
  switch (error.code) {
    case 'passkey_disabled':
      return 'Supabase 尚未啟用通行密鑰，請先使用帳密登入。'
    case 'webauthn_credential_exists':
      return '這組通行密鑰已經註冊。'
    case 'webauthn_credential_not_found':
      return '找不到此通行密鑰，請先用帳密登入並重新設定。'
    case 'webauthn_challenge_expired':
    case 'webauthn_challenge_not_found':
      return '通行密鑰驗證已逾時，請重試。'
  }
  if (error.name === 'NotAllowedError' || error.name === 'AbortError') {
    return '通行密鑰操作已取消、逾時，或找不到可用的通行密鑰。'
  }
  return '通行密鑰操作失敗，請確認裝置設定或使用帳密登入。'
}

export async function signInWithPasskey(signal?: AbortSignal) {
  if (!supabase || !supportsPasskeys()) throw new Error('此瀏覽器或網址不支援通行密鑰。')
  const { error } = await supabase.auth.signInWithPasskey({ options: { signal } })
  if (error) throw new Error(passkeyErrorMessage(error))
}

export async function registerOwnerPasskey() {
  if (!supabase || !supportsPasskeys()) throw new Error('此瀏覽器或網址不支援通行密鑰。')
  await requireAuthenticatedUserId()
  const { error } = await supabase.auth.registerPasskey()
  if (error) throw new Error(passkeyErrorMessage(error))
}

export async function requireAuthenticatedUserId(): Promise<string> {
  if (isDemoMode || !supabase) throw new Error('目前未連接 Supabase 雲端資料')
  const { data, error } = await supabase.auth.getUser()
  if (error || !data.user || !isOwnerUser(data.user.id)) throw new Error('登入狀態已失效，請重新登入')
  return data.user.id
}

export async function verifyOwnerPassword(password: string): Promise<void> {
  if (isDemoMode || !supabase) throw new Error('目前未連接 Supabase 雲端資料')
  if (!/^\d{6,}$/.test(password)) throw new Error('登入密碼必須是至少 6 位數字')
  const { data, error } = await supabase.auth.signInWithPassword({ email: OWNER_EMAIL, password })
  if (error) throw new Error('登入密碼不正確')
  if (!isOwnerUser(data.user?.id)) {
    await supabase.auth.signOut()
    throw new Error('這個帳號沒有系統使用權限')
  }
}
