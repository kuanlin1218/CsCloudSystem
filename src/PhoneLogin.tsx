import { useCallback, useEffect, useRef, useState } from 'react'
import { Camera, QrCode } from 'lucide-react'
import { isOwnerUser } from './lib/auth'
import { createQrCodeDataUrl } from './lib/barcode'
import { supabase } from './lib/supabase'

const loginQrPattern = /^cnc-phone-login:v2:([0-9a-f]{64})$/
type StartedLogin = { desktopToken: string; qrPayload: string; expiresAt: string }
type LoginStatus = { status: 'pending' | 'expired' } | {
  status: 'complete'; access_token: string; refresh_token: string
}

async function requestPhoneLogin<T>(body: object, signal?: AbortSignal): Promise<T> {
  if (!supabase) throw new Error('手機掃碼登入需要連線到雲端。')
  const { data, error } = await supabase.functions.invoke<T>('phone-qr-login', { body, signal })
  if (!error && data) return data
  const response = error?.context
  if (response instanceof Response) {
    const details = await response.json().catch(() => null)
    if (typeof details?.error === 'string') throw new Error(details.error)
  }
  throw new Error('手機掃碼登入暫時無法使用，請稍後重試或改用帳密登入。')
}

export function PhoneLoginButton({ disabled }: { disabled: boolean }) {
  const [attempt, setAttempt] = useState<(StartedLogin & { image: string }) | null>(null)
  const [starting, setStarting] = useState(false)
  const [completing, setCompleting] = useState(false)
  const [message, setMessage] = useState('')

  useEffect(() => {
    if (!attempt || !supabase) return
    const client = supabase
    let active = true
    let timer = 0
    const poll = async () => {
      try {
        if (Date.now() >= Date.parse(attempt.expiresAt)) {
          setAttempt(null)
          setMessage('QR Code 已到期，請重新產生。')
          return
        }
        const result = await requestPhoneLogin<LoginStatus>({ action: 'poll', token: attempt.desktopToken })
        if (!active) return
        if (result.status === 'pending') {
          timer = window.setTimeout(() => { void poll() }, 3000)
        } else if (result.status === 'expired') {
          setAttempt(null)
          setMessage('這次登入請求已失效，請重新產生 QR Code。')
        } else if (result.status === 'complete') {
          setCompleting(true)
          const { data, error } = await client.auth.setSession({
            access_token: result.access_token, refresh_token: result.refresh_token,
          })
          if (error || !isOwnerUser(data.user?.id)) throw new Error('登入未完成，請重新產生 QR Code。')
        }
      } catch (error) {
        if (!active) return
        setAttempt(null)
        setCompleting(false)
        setMessage(error instanceof Error ? error.message : '登入未完成，請重試。')
      }
    }
    timer = window.setTimeout(() => { void poll() }, 3000)
    return () => { active = false; window.clearTimeout(timer) }
  }, [attempt])

  const start = async () => {
    setStarting(true)
    setMessage('')
    try {
      const result = await requestPhoneLogin<StartedLogin>({ action: 'start' })
      if (!loginQrPattern.test(result.qrPayload) || !validExpiry(result.expiresAt)) {
        throw new Error('無法取得新版掃碼登入請求，請確認 Edge Function 已更新。')
      }
      const image = await createQrCodeDataUrl(result.qrPayload, 240)
      setAttempt({ ...result, image })
    } catch (error) { setMessage(error instanceof Error ? error.message : '無法產生 QR Code。') }
    finally { setStarting(false) }
  }

  const cancel = () => {
    if (!attempt) return
    setAttempt(null)
    // 本分頁立即停止等待。若離線無法通知後端，原請求會自行到期。
    void requestPhoneLogin({ action: 'cancel', token: attempt.desktopToken }).catch(() => {})
  }

  return <>
    <button className="button auth-qr-button" type="button" onClick={() => void start()} disabled={disabled || starting || Boolean(attempt)}>
      <QrCode aria-hidden="true" />{starting ? '正在產生 QR Code…' : '手機掃碼登入'}
    </button>
    {attempt && <section className="auth-qr-panel" aria-label="掃碼登入電腦">
      <p>在手機開啟詮栩 CNC，點「掃描電腦登入 QR Code」。</p>
      <img src={attempt.image} alt="手機掃碼登入 QR Code" />
      <p role="status">{completing ? '正在登入…' : '等待手機通行密鑰確認，QR Code 有效 3 分鐘。'}</p>
      <button className="button" type="button" disabled={completing} onClick={cancel}>取消掃碼登入</button>
    </section>}
    {message && <p className="form-error" role="alert">{message}</p>}
  </>
}

function validExpiry(value: string) {
  const deadline = Date.parse(value)
  return Number.isFinite(deadline) && deadline > Date.now()
}

type PasskeyChallenge = { challengeId: string; options: PublicKeyCredentialRequestOptionsJSON }

function decodeBase64Url(value: string) {
  return Uint8Array.from(atob(value.replace(/-/g, '+').replace(/_/g, '/')), (byte) => byte.charCodeAt(0))
}

function encodeBase64Url(value: ArrayBuffer) {
  return btoa(String.fromCharCode(...new Uint8Array(value))).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '')
}

async function confirmPasskey(options: PublicKeyCredentialRequestOptionsJSON, signal: AbortSignal) {
  // 舊版 iOS 可直接接收二進位挑戰與通行密鑰識別碼。
  const publicKey: PublicKeyCredentialRequestOptions = typeof PublicKeyCredential.parseRequestOptionsFromJSON === 'function'
    ? PublicKeyCredential.parseRequestOptionsFromJSON(options)
    : {
      rpId: options.rpId,
      timeout: options.timeout,
      challenge: decodeBase64Url(options.challenge),
      allowCredentials: options.allowCredentials?.map((entry) => ({ type: 'public-key', id: decodeBase64Url(entry.id) })),
    }
  publicKey.userVerification = 'required'
  const credential = await navigator.credentials.get({ publicKey, signal })
  if (!(credential instanceof PublicKeyCredential)) throw new Error('未完成通行密鑰確認。')
  if (typeof credential.toJSON === 'function') return credential.toJSON()
  const response = credential.response as AuthenticatorAssertionResponse
  return {
    id: credential.id,
    rawId: encodeBase64Url(credential.rawId),
    type: credential.type,
    authenticatorAttachment: credential.authenticatorAttachment,
    clientExtensionResults: credential.getClientExtensionResults(),
    response: {
      authenticatorData: encodeBase64Url(response.authenticatorData),
      clientDataJSON: encodeBase64Url(response.clientDataJSON),
      signature: encodeBase64Url(response.signature),
      userHandle: response.userHandle ? encodeBase64Url(response.userHandle) : undefined,
    },
  }
}

function LoginQrCamera({ onScan }: { onScan: (token: string) => void }) {
  const video = useRef<HTMLVideoElement>(null)
  const [message, setMessage] = useState('')

  useEffect(() => {
    const preview = video.current
    if (!preview) return
    let active = true
    let stream: MediaStream | undefined
    let controls: { stop: () => void } | undefined
    const stop = () => {
      active = false
      controls?.stop()
      stream?.getTracks().forEach((track) => track.stop())
    }
    const start = async () => {
      try {
        if (!navigator.mediaDevices?.getUserMedia) throw new Error('此裝置無法開啟相機。')
        stream = await navigator.mediaDevices.getUserMedia({ audio: false, video: { facingMode: { ideal: 'environment' } } })
        if (!active) { stop(); return }
        const { BrowserQRCodeReader } = await import('@zxing/browser')
        if (!active) return
        const reader = new BrowserQRCodeReader()
        controls = await reader.decodeFromStream(stream, preview, (result) => {
          if (!active || !result) return
          const token = result.getText().match(loginQrPattern)?.[1]
          if (!token) { setMessage('請對準電腦登入頁顯示的 QR Code。'); return }
          stop()
          onScan(token)
        })
        if (!active) stop()
      } catch (error) {
        if (!active) return
        setMessage(error instanceof DOMException && error.name === 'NotAllowedError'
          ? '相機權限未開啟。請允許詮栩 CNC 使用相機，再重新開啟掃描。'
          : '無法開啟掃碼相機。請關閉占用相機的程式後重試。')
        stop()
      }
    }
    void Promise.resolve().then(() => { if (active) return start() })
    return stop
  }, [onScan])

  return <>
    <p className="auth-copy">對準電腦登入頁顯示的 QR Code。</p>
    <video className="auth-qr-camera" ref={video} muted playsInline aria-label="電腦登入 QR Code 相機預覽" />
    {message && <p className="form-error" role="alert">{message}</p>}
  </>
}

export function PhoneLoginScanner({ onClose }: { onClose: () => void }) {
  const [token, setToken] = useState('')
  const [challenge, setChallenge] = useState<PasskeyChallenge | null>(null)
  const [busy, setBusy] = useState(false)
  const [approved, setApproved] = useState(false)
  const [message, setMessage] = useState('')
  const operation = useRef<AbortController | null>(null)

  useEffect(() => () => operation.current?.abort(), [])

  const prepare = useCallback(async (scannedToken: string) => {
    const controller = new AbortController()
    operation.current?.abort()
    operation.current = controller
    setToken(scannedToken)
    setChallenge(null)
    setMessage('')
    setBusy(true)
    try {
      const result = await requestPhoneLogin<PasskeyChallenge>({ action: 'prepare', token: scannedToken }, controller.signal)
      if (!controller.signal.aborted) setChallenge(result)
    } catch (error) {
      if (!controller.signal.aborted) setMessage(error instanceof Error ? error.message : '無法讀取登入請求。')
    } finally { if (!controller.signal.aborted) setBusy(false) }
  }, [])

  const approve = async () => {
    if (!challenge || busy || approved) return
    const controller = new AbortController()
    operation.current = controller
    setBusy(true)
    setMessage('')
    try {
      const credential = await confirmPasskey(challenge.options, controller.signal)
      await requestPhoneLogin({ action: 'approve', token, challengeId: challenge.challengeId, credential }, controller.signal)
      if (!controller.signal.aborted) setApproved(true)
    } catch (error) {
      if (controller.signal.aborted) return
      setChallenge(null)
      setMessage(error instanceof DOMException && error.name === 'NotAllowedError'
        ? '通行密鑰確認已取消或逾時。'
        : error instanceof Error ? error.message : '通行密鑰確認未完成。')
    } finally { if (!controller.signal.aborted) setBusy(false) }
  }

  return <main className="auth-screen">
    <span className="auth-grid" />
    <section className="auth-card" aria-labelledby="phone-login-title">
      <h1 id="phone-login-title">{token ? '確認電腦登入' : '掃描電腦登入 QR Code'}</h1>
      {!token ? <LoginQrCamera onScan={prepare} /> : approved ? <p className="auth-copy" role="status">已允許登入，電腦會自動進入系統。</p> : <>
        <p className="auth-copy">請只批准你剛才在電腦發出的登入請求。</p>
        {message && <p className="form-error" role="alert">{message}</p>}
        {busy && <p className="auth-copy" role="status">{challenge ? '正在驗證通行密鑰…' : '正在準備通行密鑰…'}</p>}
        {challenge && <button className="button primary auth-submit" type="button" disabled={busy} onClick={() => void approve()}>用 Face ID 確認</button>}
        {!challenge && !busy && <button className="button primary auth-submit" type="button" onClick={() => void prepare(token)}>重試通行密鑰</button>}
      </>}
      <button className="button auth-qr-button" type="button" onClick={onClose} disabled={busy}>{approved ? '完成' : '取消'}</button>
    </section>
  </main>
}

export function OpenPhoneLoginScanner({ onOpen, disabled }: { onOpen: () => void; disabled?: boolean }) {
  return <button className="button auth-qr-button" type="button" onClick={onOpen} disabled={disabled}>
    <Camera aria-hidden="true" />掃描電腦登入 QR Code
  </button>
}
