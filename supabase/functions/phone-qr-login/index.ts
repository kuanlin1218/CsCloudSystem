import { createClient } from 'npm:@supabase/supabase-js@2.119.0'

const siteSetting = Deno.env.get('QR_LOGIN_SITE_URL')
const ownerId = Deno.env.get('QR_LOGIN_OWNER_USER_ID')?.toLowerCase()
const supabaseUrl = Deno.env.get('SUPABASE_URL')
const serviceKey = Deno.env.get('SUPABASE_SERVICE_ROLE_KEY')
const publicKey = Deno.env.get('SUPABASE_ANON_KEY')
const site = siteSetting ? new URL(siteSetting) : null
const tokenPattern = /^[0-9a-f]{64}$/
const codePattern = /^[0-9]{6}$/
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

  let input: { action?: unknown; token?: unknown; code?: unknown } | null
  try {
    const body = await request.text()
    if (body.length > 1024) return respond({ error: '請求格式錯誤。' }, 400)
    input = JSON.parse(body)
  } catch { return respond({ error: '請求格式錯誤。' }, 400) }
  if (!input || typeof input !== 'object') return respond({ error: '請求格式錯誤。' }, 400)

  const admin = createClient(supabaseUrl, serviceKey, authOptions)
  try {
    if (input.action === 'start') {
      const desktopToken = randomToken()
      const phoneToken = randomToken()
      const code = String(crypto.getRandomValues(new Uint32Array(1))[0] % 1_000_000).padStart(6, '0')
      const [desktopHash, phoneHash, confirmationHash] = await Promise.all([
        digest(desktopToken), digest(phoneToken), digest(`${phoneToken}:${code}`),
      ])
      const { data: expiresAt, error } = await admin.rpc('create_phone_qr_login_request', {
        p_desktop_hash: desktopHash, p_phone_hash: phoneHash, p_confirmation_hash: confirmationHash,
      })
      if (error) throw error
      if (!expiresAt) return respond({ error: '目前掃碼登入請求過多，請稍後重試。' }, 429)
      const approvalUrl = new URL(site)
      approvalUrl.hash = `phone-login=${phoneToken}`
      return respond({ desktopToken, approvalUrl: approvalUrl.href, code, expiresAt })
    }

    if (typeof input.token !== 'string' || !tokenPattern.test(input.token)) {
      return respond({ error: '登入請求無效，請在電腦重新產生 QR Code。' }, 400)
    }
    const tokenHash = await digest(input.token)
    if (input.action === 'approve') {
      if (typeof input.code !== 'string' || !codePattern.test(input.code)) {
        return respond({ error: '請輸入電腦顯示的六位確認碼。' }, 400)
      }
      const jwt = request.headers.get('authorization')?.match(/^Bearer (\S+)$/i)?.[1]
      if (!jwt) return respond({ error: '請先用管理員帳號登入手機。' }, 401)
      const { data, error: authError } = await admin.auth.getUser(jwt)
      if (authError || data.user?.id !== ownerId) return respond({ error: '請用管理員帳號登入手機。' }, 403)
      const confirmationHash = await digest(`${input.token}:${input.code}`)
      const { data: approved, error } = await admin.from('phone_qr_login_requests')
        .update({ approved_user_id: data.user.id })
        .eq('phone_hash', tokenHash).eq('confirmation_hash', confirmationHash)
        .is('approved_user_id', null).gt('expires_at', new Date().toISOString())
        .select('phone_hash').maybeSingle()
      if (error) throw error
      if (!approved) return respond({ error: '確認碼不符，或請求已逾時或使用。請核對電腦畫面。' }, 409)
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
