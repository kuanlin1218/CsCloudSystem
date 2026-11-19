import { createClient, type VerifyPasskeyAuthenticationParams } from 'npm:@supabase/supabase-js@2.119.0'

const siteSetting = Deno.env.get('QR_LOGIN_SITE_URL')
const ownerId = Deno.env.get('QR_LOGIN_OWNER_USER_ID')?.toLowerCase()
const supabaseUrl = Deno.env.get('SUPABASE_URL')
const serviceKey = Deno.env.get('SUPABASE_SERVICE_ROLE_KEY')
const publicKey = Deno.env.get('SUPABASE_ANON_KEY')
const site = siteSetting ? new URL(siteSetting) : null
const tokenPattern = /^[0-9a-f]{64}$/
const ownerPattern = /^[0-9a-f]{8}(?:-[0-9a-f]{4}){3}-[0-9a-f]{12}$/i
const authOptions = { auth: { persistSession: false, autoRefreshToken: false, detectSessionInUrl: false } }

function hex(bytes: Uint8Array) {
  return Array.from(bytes, (value) => value.toString(16).padStart(2, '0')).join('')
}

async function digest(value: string) {
  return hex(new Uint8Array(await crypto.subtle.digest('SHA-256', new TextEncoder().encode(value))))
}

function randomToken() {
  return hex(crypto.getRandomValues(new Uint8Array(32)))
}

function verifiedAssertion(value: unknown): VerifyPasskeyAuthenticationParams['credential'] | null {
  const assertion = value as { type?: unknown; response?: { authenticatorData?: unknown } } | null
  const authenticatorData = assertion?.response?.authenticatorData
  if (assertion?.type !== 'public-key' || typeof authenticatorData !== 'string') return null
  try {
    const bytes = atob(authenticatorData.replace(/-/g, '+').replace(/_/g, '/'))
    // UV 位元屬於通行密鑰簽署的資料，後續仍由 Supabase Auth 驗證整份簽章。
    return bytes.length >= 37 && (bytes.charCodeAt(32) & 4) !== 0
      ? value as VerifyPasskeyAuthenticationParams['credential'] : null
  } catch { return null }
}

Deno.serve(async (request) => {
  const headers = new Headers({ 'cache-control': 'no-store', vary: 'Origin' })
  const respond = (body: object, status = 200) => Response.json(body, { status, headers })
  if (!site || site.protocol !== 'https:' || !site.pathname.endsWith('/') || site.search || site.hash ||
      site.username || site.password || !ownerId || !ownerPattern.test(ownerId) || !supabaseUrl || !serviceKey || !publicKey) {
    return respond({ error: '手機掃碼登入尚未設定完成。' }, 503)
  }
  if (request.headers.get('origin') !== site.origin) return respond({ error: '來源不符合設定。' }, 403)

  headers.set('access-control-allow-origin', site.origin)
  headers.set('access-control-allow-methods', 'POST, OPTIONS')
  headers.set('access-control-allow-headers', 'authorization, apikey, content-type, x-client-info')
  if (request.method === 'OPTIONS') return new Response(null, { status: 204, headers })
  if (request.method !== 'POST') return respond({ error: '不支援此操作。' }, 405)

  let input: { action?: unknown; token?: unknown; challengeId?: unknown; credential?: unknown } | null
  try {
    const body = await request.text()
    if (body.length > 16_384) return respond({ error: '請求格式錯誤。' }, 400)
    input = JSON.parse(body)
  } catch { return respond({ error: '請求格式錯誤。' }, 400) }
  if (!input || typeof input !== 'object') return respond({ error: '請求格式錯誤。' }, 400)

  const admin = createClient(supabaseUrl, serviceKey, authOptions)
  try {
    if (input.action === 'start') {
      const desktopToken = randomToken()
      const phoneToken = randomToken()
      const [desktopHash, phoneHash] = await Promise.all([digest(desktopToken), digest(phoneToken)])
      const { data: expiresAt, error } = await admin.rpc('create_phone_qr_login_request', {
        p_desktop_hash: desktopHash, p_phone_hash: phoneHash, p_confirmation_hash: randomToken(),
      })
      if (error) throw error
      if (!expiresAt) return respond({ error: '目前掃碼登入請求過多，請稍後重試。' }, 429)
      return respond({ desktopToken, qrPayload: `cnc-phone-login:v2:${phoneToken}`, expiresAt })
    }

    if (typeof input.token !== 'string' || !tokenPattern.test(input.token)) {
      return respond({ error: '登入請求無效，請在電腦重新產生 QR Code。' }, 400)
    }
    const tokenHash = await digest(input.token)
    if (input.action === 'prepare' || input.action === 'approve') {
      const pendingRequest = () => admin.from('phone_qr_login_requests').select('confirmation_hash')
        .eq('phone_hash', tokenHash).is('approved_user_id', null).gt('expires_at', new Date().toISOString())
      const passkeyClient = createClient(supabaseUrl, publicKey, authOptions)
      if (input.action === 'prepare') {
        const { data: pending, error } = await pendingRequest().maybeSingle()
        if (error) throw error
        if (!pending) return respond({ error: '此 QR Code 已失效，請在電腦重新產生。' }, 409)
        const { data: challenge, error: challengeError } = await passkeyClient.auth.passkey.startAuthentication()
        if (challengeError || !challenge) return respond({ error: '無法啟動通行密鑰，請確認 Supabase 通行密鑰已啟用。' }, 503)
        const challengeHash = await digest(`passkey:${input.token}:${challenge.challenge_id}`)
        const { data: updated, error: updateError } = await admin.from('phone_qr_login_requests')
          .update({ confirmation_hash: challengeHash }).eq('phone_hash', tokenHash)
          .is('approved_user_id', null).gt('expires_at', new Date().toISOString())
          .select('phone_hash').maybeSingle()
        if (updateError) throw updateError
        if (!updated) return respond({ error: '此 QR Code 已失效，請在電腦重新產生。' }, 409)
        return respond({ challengeId: challenge.challenge_id, options: challenge.options })
      }

      const credential = verifiedAssertion(input.credential)
      if (typeof input.challengeId !== 'string' || !input.challengeId || input.challengeId.length > 128 || !credential) {
        return respond({ error: '請使用手機通行密鑰確認登入。' }, 400)
      }
      const confirmationHash = await digest(`passkey:${input.token}:${input.challengeId}`)
      const { data: pending, error: checkError } = await pendingRequest().eq('confirmation_hash', confirmationHash).maybeSingle()
      if (checkError) throw checkError
      if (!pending) return respond({ error: '通行密鑰請求已失效，請重新確認。' }, 409)
      const { data: verified, error: verifyError } = await passkeyClient.auth.passkey.verifyAuthentication({
        challengeId: input.challengeId, credential,
      })
      if (verifyError || !verified.session) return respond({ error: '通行密鑰驗證未完成，請重新確認。' }, 401)
      // 驗證工作階段僅供後端確認身分。只撤銷這次臨時工作階段，保留手機既有登入。
      const { error: signOutError } = await passkeyClient.auth.admin.signOut(verified.session.access_token, 'local')
      if (signOutError) throw signOutError
      if (verified.user?.id !== ownerId) return respond({ error: '請選擇詮栩 CNC 管理員的通行密鑰。' }, 403)
      const { data: approved, error } = await admin.from('phone_qr_login_requests')
        .update({ approved_user_id: ownerId })
        .eq('phone_hash', tokenHash).eq('confirmation_hash', confirmationHash)
        .is('approved_user_id', null).gt('expires_at', new Date().toISOString())
        .select('phone_hash').maybeSingle()
      if (error) throw error
      if (!approved) return respond({ error: '此 QR Code 已逾時或使用。請在電腦重新產生。' }, 409)
      return respond({ status: 'approved' })
    }

    if (input.action === 'cancel') {
      const { error } = await admin.from('phone_qr_login_requests').delete().eq('desktop_hash', tokenHash)
      if (error) throw error
      return respond({ status: 'cancelled' })
    }

    if (input.action !== 'poll') return respond({ error: '不支援此操作。' }, 400)
    const { data: attempt, error: lookupError } = await admin.from('phone_qr_login_requests')
      .select('approved_user_id').eq('desktop_hash', tokenHash)
      .gt('expires_at', new Date().toISOString()).maybeSingle()
    if (lookupError) throw lookupError
    if (!attempt) return respond({ status: 'expired' })
    if (!attempt.approved_user_id) return respond({ status: 'pending' })

    const { data: redeemed, error: redeemError } = await admin.rpc('redeem_phone_qr_login_request', {
      p_desktop_hash: tokenHash, p_owner_id: ownerId,
    })
    if (redeemError) throw redeemError
    if (redeemed !== ownerId) return respond({ status: 'expired' })

    // 透過 Supabase Auth 發行獨立工作階段。手機的 JWT、密碼和工作階段均不傳給電腦。
    const { data: owner, error: ownerError } = await admin.auth.admin.getUserById(ownerId)
    if (ownerError || !owner.user?.email) throw new Error('Unable to load QR login owner')
    const { data: link, error: linkError } = await admin.auth.admin.generateLink({ type: 'magiclink', email: owner.user.email })
    if (linkError || !link.properties?.hashed_token || link.user?.id !== ownerId) throw new Error('Unable to issue QR login')
    const sessionClient = createClient(supabaseUrl, publicKey, authOptions)
    const { data: verified, error: sessionError } = await sessionClient.auth.verifyOtp({ token_hash: link.properties.hashed_token, type: 'magiclink' })
    if (sessionError || verified.user?.id !== ownerId || !verified.session) throw new Error('Unable to verify QR login')
    return respond({ status: 'complete', access_token: verified.session.access_token, refresh_token: verified.session.refresh_token })
  } catch {
    return respond({ error: '手機掃碼登入暫時無法完成，請重新產生 QR Code 或用帳密登入。' }, 503)
  }
})
