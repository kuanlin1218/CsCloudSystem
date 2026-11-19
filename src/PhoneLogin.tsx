import { useCallback, useEffect, useRef, useState } from 'react'
import { createPortal } from 'react-dom'
import { Camera } from 'lucide-react'
import { isOwnerUser } from './lib/auth'
import { createQrCodeDataUrl } from './lib/barcode'
import { supabase } from './lib/supabase'

const loginQrPattern = /^cnc-phone-login:v2:([0-9a-f]{64})$/
type StartedLogin = { desktopToken: string; qrPayload: string; expiresAt: string }
type LoginStatus = { status: 'pending' | 'expired' } | {
  status: 'complete'; access_token: string; refresh_token: string
}

export function LoginFeedback({ message, success = false, target }: {
  message: string; success?: boolean; target?: HTMLElement | null
}) {
  if (!message) return null
  const notice = <p className="auth-login-feedback" data-tone={success ? 'success' : 'error'}
    role={success ? 'status' : 'alert'}>{message}</p>
  return target ? createPortal(notice, target) : notice
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

export function PhoneLoginPanel({ feedbackTarget }: { feedbackTarget: HTMLElement | null }) {
  const [image, setImage] = useState('')
  const [generation, setGeneration] = useState(0)
  const [completing, setCompleting] = useState(false)
  const [message, setMessage] = useState('')

  useEffect(() => {
    if (!supabase) return
    const client = supabase
    let active = true
    let timer = 0
    let pendingToken: string | undefined
    const cancelPending = () => {
      if (!pendingToken) return
      const token = pendingToken
      pendingToken = undefined
      void requestPhoneLogin({ action: 'cancel', token }).catch(() => {})
    }
    const fail = (text: string) => {
      cancelPending()
      if (!active) return
      setImage('')
      setCompleting(false)
      setMessage(text)
    }
    const poll = async (attempt: StartedLogin) => {
      try {
        if (Date.now() >= Date.parse(attempt.expiresAt)) {
          fail('QR Code 已到期，請重新產生。')
          return
        }
        const result = await requestPhoneLogin<LoginStatus>({ action: 'poll', token: attempt.desktopToken })
        if (!active) return
        if (result.status === 'pending') {
          timer = window.setTimeout(() => { void poll(attempt) }, 3000)
        } else if (result.status === 'expired') {
          fail('這次登入請求已失效，請重新產生 QR Code。')
        } else if (result.status === 'complete') {
          pendingToken = undefined
          setCompleting(true)
          const { data, error } = await client.auth.setSession({
            access_token: result.access_token, refresh_token: result.refresh_token,
          })
          if (error || !isOwnerUser(data.user?.id)) throw new Error('登入未完成，請重新產生 QR Code。')
        }
      } catch (error) {
        fail(error instanceof Error ? error.message : '登入未完成，請重試。')
      }
    }
    const start = async () => {
      if (!active) return
      try {
        const attempt = await requestPhoneLogin<StartedLogin>({ action: 'start' })
        pendingToken = attempt.desktopToken
        if (!active) { cancelPending(); return }
        const expiresAt = Date.parse(attempt.expiresAt)
        if (!loginQrPattern.test(attempt.qrPayload) || !Number.isFinite(expiresAt) || expiresAt <= Date.now()) {
          throw new Error('無法取得新版掃碼登入請求，請確認 Edge Function 已更新。')
        }
        const nextImage = await createQrCodeDataUrl(attempt.qrPayload, 240)
        if (!active) return
        setImage(nextImage)
        timer = window.setTimeout(() => { void poll(attempt) }, 3000)
      } catch (error) {
        fail(error instanceof Error ? error.message : '無法產生 QR Code。')
      }
    }
    // 切換登入方式時停止輪詢，並取消已建立的登入請求。
    void Promise.resolve().then(start)
    return () => { active = false; window.clearTimeout(timer); cancelPending() }
  }, [generation])

  return <section className="auth-qr-panel" aria-label="掃碼登入電腦">
    {image ? <>
      <p>在手機開啟詮栩管理系統，點「掃描電腦登入 QR Code」。</p>
      <img src={image} alt="手機掃碼登入 QR Code" />
      <p role="status">{completing ? '正在登入…' : '等待手機通行密鑰確認，QR Code 有效 3 分鐘。'}</p>
    </> : message ? <>
      <LoginFeedback message={message} target={feedbackTarget} />
      <button className="button auth-qr-retry" type="button" onClick={() => { setMessage(''); setGeneration((value) => value + 1) }}>重新產生 QR Code</button>
    </> : <p role="status">正在產生 QR Code…</p>}
  </section>
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

function LoginQrCamera({ onScan, feedbackTarget }: {
  onScan: (token: string) => void; feedbackTarget?: HTMLElement | null
}) {
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
    <p className="auth-copy">對準你剛才在電腦登入頁產生的 QR Code，掃到後會啟動通行密鑰驗證。</p>
    <video className="auth-qr-camera" ref={video} muted playsInline aria-label="電腦登入 QR Code 相機預覽" />
    <LoginFeedback message={message} target={feedbackTarget} />
  </>
}

export function PhoneLoginScanner({ onClose, feedbackTarget }: {
  onClose?: () => void; feedbackTarget?: HTMLElement | null
}) {
  const [token, setToken] = useState('')
  const [challenge, setChallenge] = useState<PasskeyChallenge | null>(null)
  const [busy, setBusy] = useState(false)
  const [approved, setApproved] = useState(false)
  const [message, setMessage] = useState('')
  const operation = useRef<AbortController | null>(null)

  useEffect(() => () => operation.current?.abort(), [])

  const finish = useCallback(async (scannedToken: string, request: PasskeyChallenge, signal: AbortSignal, automatic = false) => {
    let credential: unknown
    try { credential = await confirmPasskey(request.options, signal) }
    catch (error) {
      // 部分 iPhone 需要先點一下才能顯示系統驗證，保留同一挑戰供按鈕重試。
      if (automatic && error instanceof DOMException && ['NotAllowedError', 'SecurityError'].includes(error.name)) return
      throw error
    }
    if (signal.aborted) return
    await requestPhoneLogin({ action: 'approve', token: scannedToken, challengeId: request.challengeId, credential }, signal)
    if (!signal.aborted) setApproved(true)
  }, [])

  const prepare = useCallback(async (scannedToken: string, automatic = false) => {
    const controller = new AbortController()
    operation.current?.abort()
    operation.current = controller
    setToken(scannedToken)
    setChallenge(null)
    setMessage('')
    setBusy(true)
    try {
      const result = await requestPhoneLogin<PasskeyChallenge>({ action: 'prepare', token: scannedToken }, controller.signal)
      if (controller.signal.aborted) return
      setChallenge(result)
      if (automatic) await finish(scannedToken, result, controller.signal, true)
    } catch (error) {
      if (!controller.signal.aborted) {
        setChallenge(null)
        setMessage(error instanceof Error ? error.message : '無法讀取登入請求。')
      }
    } finally { if (!controller.signal.aborted) setBusy(false) }
  }, [finish])
  const scan = useCallback((scannedToken: string) => { void prepare(scannedToken, true) }, [prepare])

  const approve = async () => {
    if (!challenge || busy || approved) return
    const controller = new AbortController()
    operation.current = controller
    setBusy(true)
    setMessage('')
    try {
      await finish(token, challenge, controller.signal)
    } catch (error) {
      if (controller.signal.aborted) return
      setChallenge(null)
      setMessage(error instanceof DOMException && error.name === 'NotAllowedError'
        ? '通行密鑰確認已取消或逾時。'
        : error instanceof Error ? error.message : '通行密鑰確認未完成。')
    } finally { if (!controller.signal.aborted) setBusy(false) }
  }

  const content = <>
      {!token ? <LoginQrCamera onScan={scan} feedbackTarget={feedbackTarget} />
        : approved ? <LoginFeedback message="已允許登入，電腦會自動進入系統。" success target={feedbackTarget} /> : <>
        <LoginFeedback message={message} target={feedbackTarget} />
        {busy && <p className="auth-copy" role="status">{challenge ? '正在驗證通行密鑰…' : '正在準備通行密鑰…'}</p>}
        {!busy && <button className="button primary auth-method-action" type="button"
          onClick={() => void (challenge ? approve() : prepare(token))}>
          {challenge ? '用 Face ID 確認' : '重試通行密鑰'}
        </button>}
      </>}
  </>

  if (!onClose) return <section aria-label="掃描電腦登入 QR Code">{content}</section>
  return <main className="auth-screen auth-login-screen">
    <span className="auth-grid" />
    <section className="auth-card" aria-labelledby="phone-login-title">
      <div className="auth-method-heading">
        <h1 id="phone-login-title">掃碼登入</h1>
        <button className="button auth-method-back" type="button" onClick={onClose}>{approved ? '完成' : '返回'}</button>
      </div>
      {content}
    </section>
  </main>
}

export function OpenPhoneLoginScanner({ onOpen, disabled }: { onOpen: () => void; disabled?: boolean }) {
  return <button className="button auth-qr-button" type="button" onClick={onOpen} disabled={disabled}>
    <Camera aria-hidden="true" />掃描電腦登入 QR Code
  </button>
}
