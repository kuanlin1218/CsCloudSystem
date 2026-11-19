import { useEffect, useState, type FormEvent } from 'react'
import { QrCode } from 'lucide-react'
import { isOwnerUser } from './lib/auth'
import { createQrCodeDataUrl } from './lib/barcode'
import { supabase } from './lib/supabase'

type StartedLogin = { desktopToken: string; approvalUrl: string; code: string; expiresAt: string }
type LoginStatus = { status: 'pending' | 'expired' } | {
  status: 'complete'; access_token: string; refresh_token: string
}

export function readPhoneLoginToken() {
  return window.location.hash.match(/^#phone-login=([0-9a-f]{64})$/)?.[1] ?? null
}

async function requestPhoneLogin<T>(body: object): Promise<T> {
  if (!supabase) throw new Error('手機掃碼登入需要連線到雲端。')
  const { data, error } = await supabase.functions.invoke<T>('phone-qr-login', { body })
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
      const url = new URL(result.approvalUrl)
      if (url.origin !== window.location.origin || !validExpiry(result.expiresAt)) {
        throw new Error('手機掃碼登入網址或有效時間設定錯誤。')
      }
      const image = await createQrCodeDataUrl(url.href, 240)
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
      <p>用手機相機掃描，並在手機輸入下方確認碼。</p>
      <img src={attempt.image} alt="手機掃碼登入 QR Code" />
      <strong className="auth-qr-code" aria-label={`確認碼 ${attempt.code}`}>{attempt.code}</strong>
      <p role="status">{completing ? '正在登入…' : '等待手機批准，QR Code 有效 3 分鐘。'}</p>
      <button className="button" type="button" disabled={completing} onClick={cancel}>取消掃碼登入</button>
    </section>}
    {message && <p className="form-error" role="alert">{message}</p>}
  </>
}

function validExpiry(value: string) {
  const deadline = Date.parse(value)
  return Number.isFinite(deadline) && deadline > Date.now()
}

export function PhoneLoginApproval({ token, onClose }: { token: string; onClose: () => void }) {
  const [code, setCode] = useState('')
  const [busy, setBusy] = useState(false)
  const [approved, setApproved] = useState(false)
  const [message, setMessage] = useState('')

  const approve = async (event: FormEvent<HTMLFormElement>) => {
    event.preventDefault()
    if (busy || approved || !/^\d{6}$/.test(code)) return
    setBusy(true)
    setMessage('')
    try {
      await requestPhoneLogin<{ status: 'approved' }>({ action: 'approve', token, code })
      setApproved(true)
    } catch (error) { setMessage(error instanceof Error ? error.message : '無法批准此次登入。') }
    finally { setBusy(false) }
  }

  return <main className="auth-screen">
    <span className="auth-grid" />
    <section className="auth-card" aria-labelledby="phone-login-title">
      <div className="auth-brand"><h1 id="phone-login-title">確認電腦登入</h1></div>
      {approved ? <p className="auth-copy" role="status">已允許登入，請回到電腦確認。</p> : <form onSubmit={approve}>
        <p className="auth-copy">請只批准你剛才在電腦發出的登入請求。</p>
        <label className="field wide">
          <span>電腦上顯示的六位確認碼</span>
          <input type="text" inputMode="numeric" autoComplete="off" maxLength={6} required
            pattern="[0-9]{6}" value={code} disabled={busy} onChange={(event) => setCode(event.target.value.replace(/\D/g, ''))} />
        </label>
        {message && <p className="form-error" role="alert">{message}</p>}
        <button className="button primary auth-submit" type="submit" disabled={busy || code.length !== 6}>{busy ? '正在批准…' : '允許登入'}</button>
      </form>}
      <button className="button auth-qr-button" type="button" onClick={onClose} disabled={busy}>{approved ? '完成' : '取消'}</button>
    </section>
  </main>
}
