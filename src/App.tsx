import { useEffect, useMemo, useRef, useState, type FormEvent, type KeyboardEvent as ReactKeyboardEvent, type PointerEvent as ReactPointerEvent, type ReactNode } from 'react'
import type { Session, User } from '@supabase/supabase-js'
import {
  ArrowDownToLine, ArrowLeftRight, ArrowUpFromLine, Boxes, CalendarDays, ChevronDown, ChevronUp,
  Calculator, Camera, Clock3, ClipboardCheck, ClipboardList, Database, Download, History, Home, LayoutGrid,
  FileSpreadsheet, FileUp, Keyboard, ListPlus, LockKeyhole, LogOut, Menu, Pencil, Plus, QrCode, Save, Search, ShieldCheck, Trash2, Triangle, UserRound, WandSparkles, Wrench, X,
  ScanBarcode,Factory,LayoutDashboard,Layers,DatabaseBackup,Combine,ChartColumnBig,CalendarClock,LaptopMinimal
} from 'lucide-react'
import { isDemoMode, supabase } from './lib/supabase'
import { isOwnerUser, OWNER_EMAIL, OWNER_USERNAME, registerOwnerPasskey, signInWithPasskey, supportsPasskeys, verifyOwnerPassword } from './lib/auth'
import { composeProductName, formatDate, formatMoney, formatNumber, today } from './lib/format'
import { hasTaiwanHolidayCalendar, isTaiwanPublicHoliday, TAIWAN_HOLIDAY_CALENDAR_YEARS } from './lib/taiwanHolidays'
import { exportProductionXlsx } from './lib/exportProductionXlsx'
import { exportAllAdminData, importAllAdminData } from './lib/adminData'
import { createQrCodeDataUrl, generateUniqueToolBarcode } from './lib/barcode'
import type { WorkbookPreview } from './lib/adminDataWorkbook'
import { createProduct, createProductionOption, createProductionRecord, deleteMovement, deleteProduct, deleteProductionOption, deleteProductionProductData, deleteProductionRecord, loadInventory, loadOperationLog, loadProductionOptions, loadProductionRecords, registerMovement, renameProductionProductCode, setStocktakeQuantity, updateProduct, updateProductBarcode, updateProductionRecord } from './lib/repository'
import { PRODUCT_CATEGORIES, type Movement, type MovementInput, type MovementType, type OperationLogEntry, type Page, type Product, type ProductCategory, type ProductInput, type ProductionInput, type ProductionOption, type ProductionOptionType, type ProductionRecord } from './types'

const blankProduct = (): ProductInput => ({
  barcode: '', specification: '', size: '', chipbreaker: '', material: '', brand: '', supplier: '',
  category: '刀片', unit: '片', notes: '', initial_quantity: 0, initial_unit_price: 0, purchased_at: today(),
})

const previousLocalDate = () => {
  const date = new Date()
  date.setDate(date.getDate() - 1)
  const year = date.getFullYear()
  const month = String(date.getMonth() + 1).padStart(2, '0')
  const day = String(date.getDate()).padStart(2, '0')
  return `${year}-${month}-${day}`
}

/* 庫存管理系統｜出入庫與刀具資料的紀錄時間：固定以台灣時區顯示，避免裝置時區不同。 */
const formatTaipeiRecordTime = (value: string, includeDate = false, includeSeconds = false) => {
  const date = new Date(value)
  if (!value || Number.isNaN(date.getTime())) return '—'
  const parts = new Intl.DateTimeFormat('zh-TW', {
    timeZone: 'Asia/Taipei', year: 'numeric', month: '2-digit', day: '2-digit',
    hour: '2-digit', minute: '2-digit', second: '2-digit', hourCycle: 'h23',
  }).formatToParts(date)
  const part = (type: Intl.DateTimeFormatPartTypes) => parts.find((item) => item.type === type)?.value || ''
  const time = `${part('hour')}:${part('minute')}${includeSeconds ? `:${part('second')}` : ''}`
  return includeDate ? `${part('year')}.${part('month')}.${part('day')} ${time}` : time
}

/* 「加工管理系統總覽」的「長條圖日期與星期」：圖表刻度顯示月日，日期範圍顯示完整年月日。 */
const formatProductionChartDate = (value: string, includeYear = false) => {
  const [year, month, day] = value.split('-').map(Number)
  const weekday = year && month && day ? ['日', '一', '二', '三', '四', '五', '六'][new Date(year, month - 1, day).getDay()] : ''
  const dateLabel = includeYear ? formatDate(value) : formatDate(value).slice(5)
  return weekday ? `${dateLabel}（${weekday}）` : dateLabel
}

type ForecastCapacityScope = 'all' | 'range'

type ForecastDraft = {
  orderQuantity: number
  dueDate: string
  capacityScope: ForecastCapacityScope
  historyFrom: string
  historyTo: string
  machineCount: number
}

/* 「訂單完工預估」的「未完成輸入資料」：保存在目前瀏覽器，切換功能或重新整理後仍可還原。 */
const FORECAST_PRODUCT_STORAGE_KEY = 'cnc-production-forecast-product-v1'
const forecastDraftStorageKey = (productCode: string) => `cnc-production-forecast-draft-v1:${productCode.trim().toUpperCase()}`
const isForecastCapacityScope = (value: unknown): value is ForecastCapacityScope => value === 'all' || value === 'range'
const readForecastProductSelection = () => {
  try { return window.localStorage.getItem(FORECAST_PRODUCT_STORAGE_KEY) || '' }
  catch { return '' }
}
const saveForecastProductSelection = (productCode: string) => {
  try {
    if (productCode) window.localStorage.setItem(FORECAST_PRODUCT_STORAGE_KEY, productCode)
    else window.localStorage.removeItem(FORECAST_PRODUCT_STORAGE_KEY)
  } catch { /* 瀏覽器禁止本機儲存時，仍維持目前畫面操作。 */ }
}
const readForecastDraft = (productCode: string): ForecastDraft | null => {
  try {
    const raw = window.localStorage.getItem(forecastDraftStorageKey(productCode))
    if (!raw) return null
    const value = JSON.parse(raw) as Partial<ForecastDraft>
    const legacy = value as Partial<ForecastDraft> & { scheduleMode?: string; dailyHours?: number }
    if (!Number.isFinite(value.orderQuantity) || !Number.isFinite(value.machineCount)) return null
    return {
      orderQuantity: Math.max(0, Number(value.orderQuantity)),
      dueDate: typeof value.dueDate === 'string' ? value.dueDate : '',
      capacityScope: isForecastCapacityScope(value.capacityScope) ? value.capacityScope : legacy.scheduleMode === 'history-range' ? 'range' : 'all',
      historyFrom: typeof value.historyFrom === 'string' ? value.historyFrom : '',
      historyTo: typeof value.historyTo === 'string' ? value.historyTo : '',
      machineCount: Math.max(0, Number(value.machineCount)),
    }
  } catch { return null }
}
const saveForecastDraft = (productCode: string, draft: ForecastDraft) => {
  try { window.localStorage.setItem(forecastDraftStorageKey(productCode), JSON.stringify(draft)) }
  catch { /* 瀏覽器禁止本機儲存時，仍維持目前畫面操作。 */ }
}

/* 「訂單完工預估」的「日期排程」：以台灣本地日期運算，避免 UTC 時差造成日期提前或延後。 */
const parseLocalDate = (value: string) => {
  const [year, month, day] = value.split('-').map(Number)
  return new Date(year, month - 1, day, 12, 0, 0)
}
const localDateValue = (date: Date) => {
  const year = date.getFullYear()
  const month = String(date.getMonth() + 1).padStart(2, '0')
  const day = String(date.getDate()).padStart(2, '0')
  return `${year}-${month}-${day}`
}
const countForecastCalendarDays = (start: string, end: string) => {
  if (!end || end <= start) return 0
  const date = parseLocalDate(start)
  const lastDate = parseLocalDate(end)
  let count = 0
  while (date < lastDate) { date.setDate(date.getDate() + 1); count += 1 }
  return count
}

const PAGE_VALUES: Page[] = ['home', 'production-dashboard', 'production-entry', 'production-query', 'production-forecast', 'production-settings', 'production-delete', 'dashboard', 'movements', 'stocktake', 'products', 'scanner-settings', 'operation-log']
const preventInputEnterSubmit = (event: ReactKeyboardEvent<HTMLFormElement>) => {
  if (event.key !== 'Enter' || event.nativeEvent.isComposing || !(event.target instanceof HTMLInputElement)) return
  event.preventDefault()
  event.target.blur()
}
const pageFromLocation = (): Page => {
  const value = window.location.hash.replace(/^#/, '') as Page
  return PAGE_VALUES.includes(value) ? value : 'home'
}
const pageUrl = (page: Page) => `${window.location.pathname}${window.location.search}#${page}`
const GUEST_SESSION_KEY = 'cnc-cloud-guest-mode'
const INITIAL_STOCK_NOTE = '建立時初庫存'
const isInitialStockNote = (value: string) => value === INITIAL_STOCK_NOTE || value === '建立商品時的期初庫存'

type StockModalState = { type: MovementType; productId?: string }
type AdminDeleteRequest = {
  title: string
  description: string
  execute: () => Promise<void>
  resolve: (result: boolean) => void
}

type BarcodeCameraCapabilities = MediaTrackCapabilities & {
  focusMode?: string[]
  zoom?: { min: number; max: number; step: number }
}
type BarcodeCameraConstraintSet = MediaTrackConstraintSet & {
  focusMode?: 'continuous'
  zoom?: number
}

const messageFor = (error: unknown) => error instanceof Error ? error.message : (error as { message?: string })?.message || '發生未知錯誤'

export default function App() {
  const [session, setSession] = useState<Session | null>(null)
  const [guestMode, setGuestMode] = useState(() => window.sessionStorage.getItem(GUEST_SESSION_KEY) === '1')
  const [authReady, setAuthReady] = useState(isDemoMode)
  const [authMessage, setAuthMessage] = useState('')
  const [forcePasswordSetup, setForcePasswordSetup] = useState(false)

  useEffect(() => {
    const client = supabase
    if (isDemoMode || !client) return
    const acceptSession = async (nextSession: Session | null) => {
      if (nextSession && !isOwnerUser(nextSession.user.id)) {
        await client.auth.signOut()
        setSession(null)
        setAuthMessage('這個帳號沒有系統使用權限')
      } else {
        if (nextSession) {
          window.sessionStorage.removeItem(GUEST_SESSION_KEY)
          setGuestMode(false)
        }
        setSession(nextSession)
      }
      setAuthReady(true)
    }
    void client.auth.getSession().then(({ data }) => acceptSession(data.session))
    const { data: listener } = client.auth.onAuthStateChange((event, nextSession) => {
      if (event === 'PASSWORD_RECOVERY') setForcePasswordSetup(true)
      if (nextSession && !isOwnerUser(nextSession.user.id)) {
        // Supabase releases its auth lock after the callback returns.
        window.setTimeout(() => { void acceptSession(nextSession) }, 0)
        return
      }
      void acceptSession(nextSession)
    })
    return () => listener.subscription.unsubscribe()
  }, [])

  const enterGuestMode = () => {
    window.sessionStorage.setItem(GUEST_SESSION_KEY, '1')
    setAuthMessage('')
    setGuestMode(true)
  }
  const exitGuestMode = () => {
    window.sessionStorage.removeItem(GUEST_SESSION_KEY)
    window.history.replaceState({ page: 'home' }, '', pageUrl('home'))
    setGuestMode(false)
  }

  if (isDemoMode) return <AuthenticatedApp session={null} />
  if (guestMode) return <AuthenticatedApp session={null} guestMode exitGuest={exitGuestMode} />
  if (!authReady) return <LoadingScreen label="正在確認安全登入狀態…" />
  if (!session) return <LoginScreen initialMessage={authMessage} onGuest={enterGuestMode} />
  if (forcePasswordSetup || session.user.user_metadata?.password_configured !== true) {
    return <PasswordSetupScreen session={session} complete={(user) => { setForcePasswordSetup(false); setSession({ ...session, user }) }} />
  }
  return <AuthenticatedApp session={session} />
}

function LoadingScreen({ label }: { label: string }) {
  return <main className="loading-screen"><span className="loader"/><p>{label}</p></main>
}

function LoginScreen({ initialMessage, onGuest }: { initialMessage?: string; onGuest: () => void }) {
  const [username, setUsername] = useState('')
  const [password, setPassword] = useState('')
  const [message, setMessage] = useState(initialMessage || '')
  const [busy, setBusy] = useState(false)
  const [resetOpen, setResetOpen] = useState(false)
  const [resetEmail, setResetEmail] = useState('')
  const [resetMessage, setResetMessage] = useState('')
  const [resetSuccess, setResetSuccess] = useState(false)

  useEffect(() => { setMessage(initialMessage || '') }, [initialMessage])

  const submitPasskey = async () => {
    setMessage('')
    setBusy(true)
    try { await signInWithPasskey() }
    catch (error) { setMessage(messageFor(error)) }
    finally { setBusy(false) }
  }

  const submit = async (event: FormEvent) => {
    event.preventDefault(); setMessage('')
    const normalizedUsername = username.trim().toLowerCase()
    if (!/^[a-z0-9]+$/.test(normalizedUsername) || normalizedUsername !== OWNER_USERNAME) { setMessage('登入名稱或密碼不正確'); return }
    if (!/^\d{6,}$/.test(password)) { setMessage('密碼必須是至少 6 位數字'); return }
    setBusy(true)
    try { await verifyOwnerPassword(password) }
    catch (error) { setMessage(messageFor(error)) }
    finally { setBusy(false) }
  }

  const sendPasswordEmail = async (event: FormEvent) => {
    event.preventDefault()
    if (!supabase) return
    const normalizedEmail = resetEmail.trim().toLowerCase()
    setResetMessage(''); setResetSuccess(false)
    if (!normalizedEmail) { setResetMessage('請先輸入管理員 Email'); return }
    if (normalizedEmail !== OWNER_EMAIL.toLowerCase()) { setResetMessage('Email 與後台管理員帳號不符'); return }
    setBusy(true)
    try {
      const redirectTo = `${window.location.origin}${import.meta.env.BASE_URL}`
      const { error } = await supabase.auth.resetPasswordForEmail(normalizedEmail, { redirectTo })
      if (error) throw error
      setResetSuccess(true); setResetMessage('重設密碼郵件已寄出，請使用最新一封郵件中的連結。')
    } catch (error) { setResetMessage(messageFor(error)) }
    finally { setBusy(false) }
  }

  return <main className="auth-screen">
    <span className="auth-grid"/>
    <section className="auth-card" aria-label="登入系統">
      <div className="auth-brand">
        <img className="auth-logo" src={`${import.meta.env.BASE_URL}company-logo.png`} alt="詮栩公司 Logo"/>
        <p className="auth-copy">詮栩股份有限公司 Chyuan Sheu Co.,Ltd.</p>
        <h1>登入雲端系統</h1>
        <p className="auth-eyebrow">2026</p>
      </div>
    <form onSubmit={submit}>
        {/* 「登入畫面」的「登入名稱輸入框」：允許桌面中文輸入法正常組字，送出時再驗證只能使用英文字母與數字。 */}
        <Field label="登入名稱" wide><input autoFocus type="text" inputMode="email" pattern="[A-Za-z0-9]*" lang="en" autoCapitalize="none" autoCorrect="off" spellCheck={false} autoComplete="username" enterKeyHint="next" value={username} onChange={(event) => setUsername(event.target.value)} placeholder="輸入英文或數字登入名稱" disabled={busy}/></Field>
        <Field label="密碼" wide><input type="password" inputMode="numeric" pattern="[0-9]{6,}" autoComplete="current-password" value={password} onChange={(event) => setPassword(event.target.value.replace(/\D/g, ''))} placeholder="至少 6 位數字" disabled={busy}/></Field>
        {message && <p className="form-error">{message}</p>}
        <button className="button primary auth-submit" type="submit" disabled={busy}>{busy ? '正在驗證…' : '登入'}</button>
      </form>
      {supportsPasskeys() && <button className="button auth-passkey-button" type="button" onClick={() => void submitPasskey()} disabled={busy}><LockKeyhole />使用通行密鑰登入 Face ID</button>}
      <button className="button auth-guest-button" type="button" onClick={onGuest} disabled={busy}><UserRound />訪客查看</button>
      <p className="auth-guest-note">訪客僅能查看，不能編輯資料。</p>
      <button className="auth-switch" type="button" onClick={() => { setResetOpen((value) => !value); setResetMessage(''); setResetSuccess(false) }} disabled={busy}>{resetOpen ? '取消重設密碼' : '忘記密碼或尚未設定密碼'}</button>
      {resetOpen && <form className="auth-reset-form" onSubmit={sendPasswordEmail}>
        <Field label="管理員 Email" wide><input type="email" inputMode="email" autoCapitalize="none" autoCorrect="off" spellCheck={false} autoComplete="email" value={resetEmail} onChange={(event) => setResetEmail(event.target.value)} placeholder="輸入後台設定的 Email" disabled={busy}/></Field>
        {resetMessage && <p className={resetSuccess ? 'admin-success' : 'form-error'}>{resetMessage}</p>}
        <button className="button auth-reset-submit" type="submit" disabled={busy}>{busy ? '正在寄送…' : '寄送重設密碼郵件'}</button>
      </form>}
      <p className="auth-guest-note">© 2026 KUAN LIN CHEN. All rights reserved.</p>
    </section>
  </main>
}

function PasswordSetupScreen({ session, complete }: { session: Session; complete: (user: User) => void }) {
  const [password, setPassword] = useState('')
  const [confirmation, setConfirmation] = useState('')
  const [message, setMessage] = useState('')
  const [busy, setBusy] = useState(false)

  const submit = async (event: FormEvent) => {
    event.preventDefault(); setMessage('')
    if (!/^\d{6,}$/.test(password)) { setMessage('密碼必須是至少 6 位數字'); return }
    if (password !== confirmation) { setMessage('兩次輸入的密碼不同'); return }
    if (!supabase) return
    setBusy(true)
    try {
      const { data, error } = await supabase.auth.updateUser({
        password,
        data: { ...session.user.user_metadata, username: OWNER_USERNAME, password_configured: true },
      })
      if (error) throw error
      complete(data.user)
    } catch (error) { setMessage(messageFor(error)) }
    finally { setBusy(false) }
  }

  return <main className="auth-screen">
    <span className="auth-grid"/>
    <section className="auth-card" aria-label="設定登入密碼">
      <div className="auth-brand">
        <img className="auth-logo" src={`${import.meta.env.BASE_URL}company-logo.png`} alt="詮栩公司 Logo"/>
        <p className="auth-eyebrow">FIRST-TIME SECURITY SETUP</p>
        <h1>設定登入密碼</h1>
        <p className="auth-copy">登入名稱固定為 <b>{OWNER_USERNAME}</b>。請設定至少 6 位數字的新密碼。</p>
      </div>
      <form onSubmit={submit}>
        <Field label="新密碼" wide><input autoFocus type="password" inputMode="numeric" pattern="[0-9]{6,}" autoComplete="new-password" value={password} onChange={(event) => setPassword(event.target.value.replace(/\D/g, ''))} placeholder="至少 6 位數字" disabled={busy}/></Field>
        <Field label="再次輸入新密碼" wide><input type="password" inputMode="numeric" pattern="[0-9]{6,}" autoComplete="new-password" value={confirmation} onChange={(event) => setConfirmation(event.target.value.replace(/\D/g, ''))} placeholder="再次輸入相同數字密碼" disabled={busy}/></Field>
        {message && <p className="form-error">{message}</p>}
        <button className="button primary auth-submit" type="submit" disabled={busy}>{busy ? '正在安全儲存…' : '儲存密碼並進入系統'}</button>
      </form>
    </section>
  </main>
}

function AuthenticatedApp({ session, guestMode = false, exitGuest }: { session: Session | null; guestMode?: boolean; exitGuest?: () => void }) {
  const [page, setPage] = useState<Page>(pageFromLocation)
  const [products, setProducts] = useState<Product[]>([])
  const [movements, setMovements] = useState<Movement[]>([])
  const [productionRecords, setProductionRecords] = useState<ProductionRecord[]>([])
  const [productionOptions, setProductionOptions] = useState<ProductionOption[]>([])
  const [loading, setLoading] = useState(true)
  const [error, setError] = useState('')
  const [toast, setToast] = useState('')
  const [stockModal, setStockModal] = useState<StockModalState | null>(null)
  const [productModal, setProductModal] = useState<Product | 'new' | null>(null)
  const [scannerBarcode, setScannerBarcode] = useState<string | null>(null)
  const [pendingBarcode, setPendingBarcode] = useState('')
  const [stocktakeFocusProductId, setStocktakeFocusProductId] = useState('')
  const [priceHistoryProduct, setPriceHistoryProduct] = useState<Product | null>(null)
  const [productionOptionModal, setProductionOptionModal] = useState<ProductionOptionType | null>(null)
  const [adminDataModal, setAdminDataModal] = useState(false)
  const [adminDelete, setAdminDelete] = useState<AdminDeleteRequest | null>(null)
  const [mobileNav, setMobileNav] = useState(false)
  const [realtimeConnected, setRealtimeConnected] = useState(false)

  const refresh = async (showLoading = true) => {
    if (showLoading) setLoading(true)
    setError('')
    try {
      const [data, productionData, optionData] = await Promise.all([loadInventory(), loadProductionRecords(), loadProductionOptions()])
      setProducts(data.products); setMovements(data.movements)
      setProductionRecords(productionData)
      setProductionOptions(optionData)
    } catch (e) { setError(messageFor(e)) }
    finally { if (showLoading) setLoading(false) }
  }

  useEffect(() => { void refresh() }, [])
  useEffect(() => {
    const client = supabase
    if (!client) return
    let refreshTimer: number | undefined
    const queueRealtimeRefresh = () => {
      if (refreshTimer) window.clearTimeout(refreshTimer)
      refreshTimer = window.setTimeout(() => void refresh(false), 180)
    }
    const realtimeChannel = client
      .channel('cnc-cloud-data-sync')
      .on('postgres_changes', { event: '*', schema: 'public', table: 'products' }, queueRealtimeRefresh)
      .on('postgres_changes', { event: '*', schema: 'public', table: 'stock_movements' }, queueRealtimeRefresh)
      .on('postgres_changes', { event: '*', schema: 'public', table: 'production_records' }, queueRealtimeRefresh)
      .on('postgres_changes', { event: '*', schema: 'public', table: 'production_options' }, queueRealtimeRefresh)
      .subscribe((status) => {
        setRealtimeConnected(status === 'SUBSCRIBED')
        if (status === 'SUBSCRIBED') queueRealtimeRefresh()
      })
    return () => {
      if (refreshTimer) window.clearTimeout(refreshTimer)
      void client.removeChannel(realtimeChannel)
    }
  }, [])
  useEffect(() => {
    if (window.location.hash !== `#${pageFromLocation()}` || !window.location.hash) window.history.replaceState({ page: pageFromLocation() }, '', pageUrl(pageFromLocation()))
    const followBrowserHistory = () => {
      const nextPage = pageFromLocation()
      setPage(nextPage)
      if (nextPage !== 'stocktake') setStocktakeFocusProductId('')
      setMobileNav(false)
      setStockModal(null); setProductModal(null); setScannerBarcode(null); setPendingBarcode(''); setPriceHistoryProduct(null)
      setProductionOptionModal(null); setAdminDataModal(false)
      setAdminDelete((request) => { request?.resolve(false); return null })
      void refresh()
    }
    window.addEventListener('popstate', followBrowserHistory)
    return () => window.removeEventListener('popstate', followBrowserHistory)
  }, [])
  useEffect(() => {
    const refreshWhenVisible = () => {
      if (document.visibilityState === 'visible') void refresh()
    }
    document.addEventListener('visibilitychange', refreshWhenVisible)
    return () => document.removeEventListener('visibilitychange', refreshWhenVisible)
  }, [])
  useEffect(() => {
    if (!toast) return
    const timer = window.setTimeout(() => setToast(''), 2600)
    return () => window.clearTimeout(timer)
  }, [toast])

  const nav = (next: Page) => {
    if (guestMode && (next === 'production-entry' || next === 'production-settings' || next === 'production-delete' || next === 'stocktake')) return
    if (window.location.hash !== `#${next}`) window.history.pushState({ page: next }, '', pageUrl(next))
    if (next !== 'stocktake') setStocktakeFocusProductId('')
    setPage(next)
    setMobileNav(false)
    void refresh()
  }
  const afterSave = async (message: string) => { await refresh(false); setToast(message) }
  const requestAdminDelete = (title: string, description: string, execute: () => Promise<void>) => new Promise<boolean>((resolve) => {
    setAdminDelete({ title, description, execute, resolve })
  })
  const canReadOperationLog = !guestMode && isOwnerUser(session?.user.id)
  const activePage: Page = page === 'operation-log' && !canReadOperationLog ? 'home' : guestMode && (page === 'production-entry' || page === 'production-settings' || page === 'production-delete') ? 'production-dashboard' : guestMode && page === 'stocktake' ? 'dashboard' : page
  const isHome = activePage === 'home'
  const isInventory = activePage === 'dashboard' || activePage === 'movements' || activePage === 'stocktake' || activePage === 'products' || activePage === 'scanner-settings'
  const isProduction = activePage === 'production-dashboard' || activePage === 'production-entry' || activePage === 'production-query' || activePage === 'production-forecast' || activePage === 'production-settings' || activePage === 'production-delete'
  const isSystem = isInventory || isProduction
  const showTableScrollControls = activePage === 'production-query' || activePage === 'production-delete' || activePage === 'movements' || activePage === 'stocktake' || activePage === 'products'
  const productionEntryProductCodes = useMemo(() => productionOptions.filter((option) => option.option_type === 'product').map((option) => option.value).sort(), [productionOptions])
  const productionQueryProductCodes = useMemo(() => [...new Set([
    ...productionEntryProductCodes,
    ...productionRecords.map((record) => record.product_code),
  ].filter(Boolean))].sort(), [productionEntryProductCodes, productionRecords])
  const productionMachines = useMemo(() => productionOptions.filter((option) => option.option_type === 'machine').map((option) => option.value).sort(), [productionOptions])

  return (
    <div className={`app-shell ${isHome ? 'portal-shell' : ''} ${isSystem && !isProduction ? 'inventory-shell' : ''}`}>
      {!isDemoMode && !isSystem && <button className="portal-logout" type="button" onClick={() => guestMode ? exitGuest?.() : void supabase?.auth.signOut()}>{guestMode ? <UserRound /> : <LogOut />}{guestMode ? '返回登入' : '登出'}</button>}
      {isSystem && <aside className={`sidebar ${isProduction ? 'production-sidebar' : ''} ${mobileNav ? 'is-open' : ''}`}>
        <div className={`brand ${isProduction ? 'production-brand' : ''}`}>
          <span className="brand-mark">{isProduction? <Factory />: <><Triangle fill="currentColor" /><i /></>}</span>
          <span><strong>{isProduction ? '加工管理系統' : '庫存管理系統'}</strong><small>{isProduction ? 'PRODUCTION SYSTEM' : 'INVENTORY SYSTEM'}</small></span>
          <button className="icon-button mobile-close" onClick={() => setMobileNav(false)} aria-label="關閉選單"><X /></button>
        </div>
        <nav>
          <NavItem active={false} icon={<Home />} onClick={() => nav('home')}>系統首頁</NavItem>
          {isProduction ? <>
            <NavItem active={activePage === 'production-dashboard'} icon={<ChartColumnBig  />} onClick={() => nav('production-dashboard')}>總覽</NavItem>
            {!guestMode && <NavItem active={activePage === 'production-entry'} icon={<ListPlus />} onClick={() => nav('production-entry')}>加工數量登記</NavItem>}
            <NavItem active={activePage === 'production-query'} icon={<Search />} onClick={() => nav('production-query')}>加工數量查詢</NavItem>
            <NavItem active={activePage === 'production-forecast'} icon={<CalendarClock />} onClick={() => nav('production-forecast')}>訂單完工預估</NavItem>
            {!guestMode && <NavItem active={activePage === 'production-settings'} icon={<Wrench />} onClick={() => nav('production-settings')}>基礎資料設定</NavItem>}
            {!guestMode && <NavItem active={activePage === 'production-delete'} icon={<Combine />} onClick={() => nav('production-delete')}>資料匯出與刪除</NavItem>}
          </> : <>
            <NavItem active={activePage === 'dashboard'} icon={<LayoutDashboard />} onClick={() => nav('dashboard')}>總覽</NavItem>
            <NavItem active={activePage === 'movements'} icon={<ArrowLeftRight />} onClick={() => nav('movements')}>出入庫</NavItem>
            {!guestMode && <NavItem active={activePage === 'stocktake'} icon={<ClipboardCheck />} onClick={() => nav('stocktake')}>庫存盤點</NavItem>}
            <NavItem active={activePage === 'products'} icon={<LaptopMinimal />} onClick={() => nav('products')}>刀具資料</NavItem>
            <NavItem active={activePage === 'scanner-settings'} icon={<ScanBarcode />} onClick={() => nav('scanner-settings')}>掃描槍設定</NavItem>
          </>}
        </nav>
        <div className="sidebar-foot">
          <span className="status-line">{guestMode ? <ShieldCheck /> : isDemoMode ? <ShieldCheck /> : <Database />} {guestMode ? '訪客唯讀模式' : isDemoMode ? '示範模式' : realtimeConnected ? '雲端即時同步' : '雲端連線中'}</span>
          <span className="status-line"><UserRound /> {guestMode ? '僅供查看' : isDemoMode ? '本機示範資料' : session?.user.user_metadata?.username || OWNER_USERNAME}</span>
          {!guestMode && isOwnerUser(session?.user.id) && supportsPasskeys() && <PasskeySetup />}
          {!isDemoMode && <button className="sidebar-logout" type="button" onClick={() => guestMode ? exitGuest?.() : void supabase?.auth.signOut()}>{guestMode ? <UserRound /> : <LogOut />}{guestMode ? '返回登入' : '登出系統'}</button>}
        </div>
      </aside>}

      {isSystem && mobileNav && <button className="nav-scrim" onClick={() => setMobileNav(false)} aria-label="關閉選單" />}

      <main className={`main-panel ${!isSystem ? 'wide-panel' : ''} ${isHome ? 'portal-panel' : ''}`}>
        {isSystem && <header className={`mobile-header ${isProduction ? 'production-mobile-header' : ''}`}>
          <button className="icon-button" onClick={() => setMobileNav(true)} aria-label="開啟選單"><Menu /></button>
          <strong>{isProduction ? '加工管理系統' : '庫存管理系統'}</strong>
          <span className="mobile-dot" />
        </header>}
        {isSystem && isDemoMode && <div className="demo-banner"><ShieldCheck /> 目前是可操作的示範模式，資料儲存在這台裝置。設定 Supabase 後會自動切換為雲端模式。</div>}
        {guestMode && <div className="guest-banner"><ShieldCheck /> 訪客模式僅供查看，無法新增、修改、刪除、盤點、匯出或還原資料。</div>}
        {error && <div className="error-banner">{error}<button onClick={() => void refresh()}>重試</button></div>}

        {activePage === 'home' && <HomePage navigate={nav} openAdmin={() => setAdminDataModal(true)} guestMode={guestMode} showOperationLog={canReadOperationLog} />}
        {activePage === 'operation-log' && <OperationLogPage goHome={() => nav('home')} />}
        {activePage === 'production-dashboard' && <ProductionDashboard records={productionRecords} productCodes={productionQueryProductCodes} loading={loading} />}
        {activePage === 'production-entry' && !guestMode && <ProductionEntryPage records={productionRecords} productCodes={productionEntryProductCodes} machines={productionMachines} loading={loading} onSave={async (input) => {
          await createProductionRecord(input); await afterSave('加工數量已登記')
        }} onDeleteLast={async (record) => {
          const deleted = await requestAdminDelete('刪除最後一筆加工紀錄', `${record.product_code}・${record.machine}・${record.quantity}，刪除後無法復原。`, () => deleteProductionRecord(record.id))
          if (deleted) await afterSave('最後一筆加工紀錄已刪除')
        }} />}
        {activePage === 'production-query' && <ProductionQueryPage records={productionRecords} productCodes={productionQueryProductCodes} machines={productionMachines} loading={loading} readOnly={guestMode} onUpdateRecord={async (record, input) => {
          await updateProductionRecord(record.id, input)
          await afterSave('加工紀錄已更新')
        }} onDeleteRecord={async (record) => {
          const deleted = await requestAdminDelete('刪除加工紀錄', `${formatDate(record.work_date)}・${record.product_code}・${record.machine}・${record.quantity}，刪除後無法復原。`, () => deleteProductionRecord(record.id))
          if (deleted) await afterSave('所選加工紀錄已刪除')
          return deleted
        }} />}
        {activePage === 'production-forecast' && <ProductionForecastPage records={productionRecords} productCodes={productionQueryProductCodes} loading={loading} />}
        {activePage === 'production-settings' && !guestMode && <ProductionSettingsPage openOption={setProductionOptionModal} />}
        {activePage === 'production-delete' && !guestMode && <ProductionDataManagementPage records={productionRecords} productCodes={productionQueryProductCodes} loading={loading} onDeleteProductRecords={async (productCode, count) => {
          const deleted = await requestAdminDelete('永久刪除產品加工資料', `將刪除產品編號「${productCode}」、管理清單選項與全部 ${count} 筆加工紀錄。`, () => deleteProductionProductData(productCode))
          if (deleted) await afterSave(`${productCode} 的產品編號與全部加工紀錄已永久刪除`)
          return deleted
        }} />}
        {activePage === 'dashboard' && <Dashboard products={products} movements={movements} loading={loading} openStock={(type) => setStockModal({ type })} openScanner={() => setScannerBarcode('')} readOnly={guestMode} />}
        {activePage === 'movements' && <MovementsPage products={products} movements={movements} loading={loading} openStock={(type) => setStockModal({ type })} openScanner={() => setScannerBarcode('')} readOnly={guestMode} onDelete={async (movement) => {
          const product = movement.product || products.find((item) => item.id === movement.product_id)
          const deleted = await requestAdminDelete('刪除出入庫紀錄', `${formatDate(movement.movement_date)} 的${movement.movement_type === 'in' ? '入庫' : '出庫'}紀錄「${product?.name || ''}」。庫存會自動重新計算。`, () => deleteMovement(movement.id))
          if (deleted) await afterSave('出入庫紀錄已刪除，庫存已重新計算')
        }} />}
        {activePage === 'stocktake' && !guestMode && <StocktakePage products={products} loading={loading} add={() => setProductModal('new')} openScanner={() => setScannerBarcode('')} focusProductId={stocktakeFocusProductId} clearFocusProduct={() => setStocktakeFocusProductId('')} onAdjust={async (product, actualQuantity, notes) => {
          if (actualQuantity === product.stock_quantity) return true
          if (!window.confirm(`確定將 ${product.code}「${product.name}」的庫存由 ${product.stock_quantity} ${product.unit}調整為 ${actualQuantity} ${product.unit}？`)) return false
          try {
            await setStocktakeQuantity({ product_id: product.id, actual_quantity: actualQuantity, notes, request_key: crypto.randomUUID() })
            await afterSave(`${product.code} 的盤點數量已更新`)
            return true
          } catch (e) { setError(messageFor(e)); return false }
        }} />}
        {activePage === 'products' && <ProductsPage products={products} movements={movements} loading={loading} add={() => setProductModal('new')} edit={setProductModal} openStock={(type, productId) => setStockModal({ type, productId })} openScanner={() => setScannerBarcode('')} viewHistory={setPriceHistoryProduct} readOnly={guestMode} onDelete={async (p) => {
          const deleted = await requestAdminDelete('刪除刀具資料', `將刪除 ${p.code}「${p.name}」及其全部出入庫紀錄。`, () => deleteProduct(p.id))
          if (deleted) await afterSave('商品已刪除')
        }} />}
        {activePage === 'scanner-settings' && <ScannerSettingsPage />}
      </main>

      {showTableScrollControls && <TableScrollControls />}

      {!guestMode && stockModal && <StockModal type={stockModal.type} initialProductId={stockModal.productId} products={products} close={() => setStockModal(null)} switchType={(type) => setStockModal((current) => current ? { ...current, type } : null)} save={async (input) => {
        await registerMovement(input); setStockModal(null); await afterSave(input.movement_type === 'in' ? '入庫完成' : '出庫完成')
      }} />}
      {!guestMode && productModal && <ProductModal product={productModal === 'new' ? null : productModal} products={products} initialBarcode={productModal === 'new' ? pendingBarcode : ''} close={() => {
        if (productModal === 'new' && pendingBarcode) setScannerBarcode(pendingBarcode)
        setProductModal(null); setPendingBarcode('')
      }} save={async (input) => {
        if (productModal === 'new') await createProduct(input); else await updateProduct(productModal.id, input)
        setProductModal(null); setPendingBarcode(''); await afterSave(productModal === 'new' ? '商品已建立' : '商品已更新')
      }} />}
      {!guestMode && scannerBarcode !== null && <BarcodeScannerModal products={products} initialBarcode={scannerBarcode} close={() => setScannerBarcode(null)} createWithBarcode={(barcode) => {
        setScannerBarcode(null); setPendingBarcode(barcode); setProductModal('new')
      }} linkBarcode={async (productId, barcode) => {
        await updateProductBarcode(productId, barcode); await afterSave('條碼已綁定刀具')
      }} openStock={(type, productId) => {
        setScannerBarcode(null); setStockModal({ type, productId })
      }} openStocktake={(productId) => {
        setScannerBarcode(null); setStocktakeFocusProductId(productId); nav('stocktake')
      }} editProduct={(product) => {
        setScannerBarcode(null); setProductModal(product)
      }} />}
      {priceHistoryProduct && <PriceHistoryModal product={priceHistoryProduct} movements={movements} close={() => setPriceHistoryProduct(null)} />}
      {!guestMode && productionOptionModal && <ProductionOptionModal type={productionOptionModal} values={(productionOptionModal === 'product' ? productionEntryProductCodes : productionMachines)} close={() => setProductionOptionModal(null)} save={async (value) => {
        await createProductionOption(productionOptionModal, value); await afterSave(productionOptionModal === 'product' ? '產品編號已新增' : '機台已新增')
      }} remove={async (value) => {
        const label = productionOptionModal === 'product' ? '產品編號' : '機台'
        const deleted = await requestAdminDelete(`刪除${label}`, `確定刪除${label}「${value}」？過去已登記的加工紀錄會保留。`, () => deleteProductionOption(productionOptionModal, value))
        if (deleted) await afterSave(`${label}已刪除`)
      }} rename={productionOptionModal === 'product' ? async (currentValue, nextValue) => {
        await renameProductionProductCode(currentValue, nextValue)
        await afterSave(`產品編號已由 ${currentValue} 更新為 ${nextValue.trim().toUpperCase()}`)
      } : undefined} />}
      {!guestMode && adminDataModal && <AdminDataModal close={() => setAdminDataModal(false)} afterImport={async () => {
        await refresh(); setToast('Excel 資料已完整匯入')
      }} />}
      {!guestMode && adminDelete && <AdminDeleteModal request={adminDelete} close={(result) => {
        adminDelete.resolve(result); setAdminDelete(null)
      }} />}
      {toast && <div className="toast"><ShieldCheck />{toast}</div>}
    </div>
  )
}

function NavItem({ active, icon, onClick, children }: { active: boolean; icon: ReactNode; onClick: () => void; children: ReactNode }) {
  return <button className={`nav-item ${active ? 'active' : ''}`} onClick={onClick}>{icon}<span>{children}</span></button>
}

function HomePage({ navigate, openAdmin, guestMode, showOperationLog }: { navigate: (page: Page) => void; openAdmin: () => void; guestMode: boolean; showOperationLog: boolean }) {
  return <section className="page-content home-page">
    <div className="home-hero">
      <img
          className="home-company-logo"
          src={`${import.meta.env.BASE_URL}company-logo.png`}
          alt="詮栩公司 Logo"
        />
      {/* <span className="home-eyebrow">CHYUAN SHEU CNC CLOUD</span> */}
      <h1>詮 栩 雲 端 系 統</h1>
      <p>加工資訊與庫存管理整合平台</p>
    </div>
    <div className="home-section-head"><div><span>SYSTEM FUNCTIONS</span><h2>功能選擇</h2></div><small className="home-eyebrow"></small></div>
    <div className="function-grid portal-functions">
      <FunctionCard icon={<Layers/>} title="加工管理系統" description={guestMode ? '查看加工總覽與歷史紀錄' : '加工數量登記、查詢與產量總覽'} tone="pink" onClick={() => navigate('production-dashboard')}/>
      <FunctionCard icon={<Boxes/>} title="庫存管理系統" description="刀具庫存、出入庫與價格管理" tone="green" onClick={() => navigate('dashboard')}/>
      <FunctionCard icon={<DatabaseBackup/>} title="資料備份／還原" description={guestMode ? '訪客模式無法匯出或還原資料' : '匯出或整批匯入 Supabase Excel'} tone="purple" onClick={openAdmin} disabled={guestMode} disabledLabel="管理員限定"/>
      {showOperationLog && <FunctionCard icon={<History/>} title="操作紀錄" description="查看資料異動與登入登出" tone="yellow" onClick={() => navigate('operation-log')}/>}
    </div>
  </section>
}

const operationEntities: Record<string, string> = {
  products: '刀具', stock_movements: '出入庫紀錄', production_records: '加工紀錄', production_options: '加工選項',
}

const operationFields: Record<string, string> = {
  code: '編號', barcode: '條碼', name: '品名', specification: '規格', size: '尺寸', chipbreaker: '斷屑槽', material: '材質',
  brand: '廠牌', supplier: '供應商', category: '分類', unit: '單位', stock_quantity: '庫存', unit_price: '單價',
  purchased_at: '購入日期', notes: '備註', movement_type: '出入庫類型', quantity: '數量', movement_date: '出入庫日期',
  stock_after: '庫存餘額', work_date: '加工日期', product_code: '產品編號', machine: '機台', hours: '小時', minutes: '分鐘',
  option_type: '選項類型', value: '選項', products: '刀具', movements: '出入庫紀錄', production_records: '加工紀錄', production_options: '加工選項',
  auth_event_id: '驗證事件編號',
}

function displayOperationValue(value: string | number | boolean | null | undefined) {
  return value == null || value === '' ? '—' : String(value)
}

function describeOperation(entry: OperationLogEntry, fields: string[]) {
  const before = entry.before_data || {}
  const after = entry.after_data || {}
  const row = entry.operation === 'DELETE' ? before : after
  const value = (key: string) => displayOperationValue(row[key])
  const action = { INSERT: '新增', UPDATE: '修改', DELETE: '刪除', RESTORE: '還原', LOGIN: '登入', LOGOUT: '登出' }[entry.operation]
  const changes = fields.map((key) => `${operationFields[key] || key}：${displayOperationValue(before[key])} → ${displayOperationValue(after[key])}`)

  if (entry.entity === 'auth') return { title: `${action}系統`, lines: [] }
  if (entry.operation === 'RESTORE') {
    return { title: '從 Excel 整批還原資料', lines: [fields.map((key) => `${operationFields[key] || key} ${value(key)} 筆`).join('、')].filter(Boolean) }
  }
  if (entry.entity === 'stock_movements') {
    const stocktake = typeof row.notes === 'string' ? row.notes.match(/^庫存盤點調整：\s*(\d+)\s*→\s*(\d+)/) : null
    const movement = stocktake ? '庫存盤點' : row.movement_type === 'in' ? '入庫' : row.movement_type === 'out' ? '出庫' : '出入庫'
    const title = entry.operation === 'INSERT' ? `為 ${entry.target} 登記${movement}` : `${action} ${entry.target} 的${movement}紀錄`
    const quantity = stocktake
      ? `盤點庫存：${stocktake[1]} → ${stocktake[2]}`
      : `${movement}數量：${value('quantity')}${row.stock_after == null ? '' : `・當時庫存餘額：${value('stock_after')}`}`
    return { title, lines: entry.operation === 'UPDATE' ? changes : [`${quantity}・紀錄日期 ${value('movement_date')}`] }
  }
  if (entry.entity === 'production_records') {
    return {
      title: `${action} ${entry.target} 的加工紀錄`,
      lines: [`加工日期 ${value('work_date')}・機台 ${value('machine')}`, ...(entry.operation === 'UPDATE' ? changes : [`加工數量：${value('quantity')}`])],
    }
  }
  if (entry.entity === 'production_options') {
    const kind = row.option_type === 'machine' ? '機台' : '產品編號'
    return { title: `${action}加工${kind}選項 ${entry.target}`, lines: entry.operation === 'UPDATE' ? changes : [] }
  }
  const entity = operationEntities[entry.entity] || entry.entity
  const name = entry.entity === 'products' && row.name ? `｜${value('name')}` : ''
  const details = entry.operation === 'UPDATE' ? changes : [row.stock_quantity == null ? '' : `庫存：${value('stock_quantity')} ${row.unit || ''}\u3000單價：${value('unit_price')}`].filter(Boolean)
  return { title: `${action}${entity} ${entry.target}${name}`, lines: details }
}

function OperationLogPage({ goHome }: { goHome: () => void }) {
  const [entries, setEntries] = useState<OperationLogEntry[]>([])
  const [hasMore, setHasMore] = useState(false)
  const [authIncluded, setAuthIncluded] = useState<boolean | null>(null)
  const [loading, setLoading] = useState(true)
  const [error, setError] = useState('')
  const [refreshKey, setRefreshKey] = useState(0)

  useEffect(() => {
    let cancelled = false
    setLoading(true)
    setError('')
    void loadOperationLog().then((result) => {
      if (cancelled) return
      setEntries(result.entries)
      setHasMore(result.hasMore)
      setAuthIncluded(result.authIncluded)
    }).catch((err) => {
      if (!cancelled) setError(messageFor(err))
    }).finally(() => { if (!cancelled) setLoading(false) })
    return () => { cancelled = true }
  }, [refreshKey])

  const loadMore = async () => {
    setLoading(true)
    setError('')
    try {
      const result = await loadOperationLog(entries.at(-1))
      setEntries((current) => [...current, ...result.entries])
      setHasMore(result.hasMore)
      setAuthIncluded(result.authIncluded)
    } catch (err) { setError(messageFor(err)) }
    finally { setLoading(false) }
  }

  return <section className="page-content operation-log-page">
    <PageTitle eyebrow="OPERATION LOG" title="操作紀錄" actions={<button className="button scan-button" type="button" onClick={goHome}><Home/>返回首頁</button>} />
    <div className="operation-log-header"><span>台灣時間・最新在上</span><button className="button scan-button" type="button" disabled={loading} onClick={() => setRefreshKey((key) => key + 1)}>重新整理</button></div>
    {error && <p className="form-error" role="alert">{error}</p>}
    {authIncluded === false && <p className="form-error">尚未接入登入登出紀錄，請套用 enable-operation-log-auth.sql。</p>}
    {loading && !entries.length && <p role="status">讀取中…</p>}
    {!loading && !error && !entries.length && <p>尚無操作紀錄。</p>}
    {entries.length > 0 && <ol className="operation-log-list">{entries.map((entry) => <OperationLogItem key={entry.id} entry={entry}/>)}</ol>}
    {hasMore && <button className="button scan-button" type="button" disabled={loading} onClick={() => void loadMore()}>{loading ? '讀取中…' : '載入較早紀錄'}</button>}
  </section>
}

function OperationLogItem({ entry }: { entry: OperationLogEntry }) {
  const before = entry.before_data || {}
  const after = entry.after_data || {}
  const fields = [...new Set([...Object.keys(before), ...Object.keys(after)])].filter((key) => before[key] !== after[key])
  const { title, lines } = describeOperation(entry, fields)
  const actor = isOwnerUser(entry.actor_id ?? undefined) ? OWNER_USERNAME : entry.actor_id || '系統／未識別'

  return <li className="operation-log-item">
    <details>
      <summary>
        <span className="operation-log-meta"><time dateTime={entry.occurred_at}>{formatTaipeiRecordTime(entry.occurred_at, true, true)}</time><span>操作者 {actor}</span></span>
        <strong>{title}</strong>
        {lines.length > 0 && <span className="operation-log-preview">{lines.join('\n')}</span>}
      </summary>
      <dl className="operation-log-details">
        {entry.actor_id && <div><dt>操作者 ID</dt><dd>{entry.actor_id}</dd></div>}
        {fields.map((key) => <div key={key}>
          <dt>{operationFields[key] || key}</dt>
          <dd>{entry.operation === 'UPDATE'
            ? <><span className="operation-log-value"><b>修改前</b> {displayOperationValue(before[key])}</span><span className="operation-log-value"><b>修改後</b> {displayOperationValue(after[key])}</span></>
            : displayOperationValue((entry.operation === 'DELETE' ? before : after)[key])}</dd>
        </div>)}
      </dl>
    </details>
  </li>
}

function PasskeySetup() {
  const [busy, setBusy] = useState(false)
  const [feedback, setFeedback] = useState({ text: '', success: false })

  const register = async () => {
    setBusy(true)
    setFeedback({ text: '', success: false })
    try {
      await registerOwnerPasskey()
      setFeedback({ text: '通行密鑰已建立。下次可在登入頁選擇「使用通行密鑰登入」。', success: true })
    } catch (error) {
      setFeedback({ text: messageFor(error), success: false })
    } finally {
      setBusy(false)
    }
  }

  return <>
    <button className="sidebar-passkey" type="button" onClick={() => void register()} disabled={busy}>
      <LockKeyhole aria-hidden="true" />{busy ? '正在設定通行密鑰…' : '設定通行密鑰 Face ID'}
    </button>
    {feedback.text && <p role={feedback.success ? 'status' : 'alert'} className={`sidebar-passkey-feedback ${feedback.success ? 'form-message' : 'form-error'}`}>{feedback.text}</p>}
  </>
}

function ProductionDashboard({ records, productCodes, loading }: { records: ProductionRecord[]; productCodes: string[]; loading: boolean }) {
  /* 「加工管理系統總覽」的「每個產品獨立圖表資料」：新增產品編號後會自動建立圖表，不需要另外設定。 */
  const productCharts = useMemo(() => productCodes.map((productCode) => {
    const productRecords = records.filter((record) => record.product_code === productCode)
    const dailyTotals = new Map<string, number>()
    productRecords.forEach((record) => dailyTotals.set(record.work_date, (dailyTotals.get(record.work_date) || 0) + record.quantity))
    const days = [...dailyTotals.entries()]
      .sort(([left], [right]) => left.localeCompare(right))
      .map(([date, quantity]) => ({ date, quantity }))
    return {
      productCode,
      days,
      total: productRecords.reduce((sum, record) => sum + record.quantity, 0),
    }
  }), [productCodes, records])

  return <section className="page-content production-dashboard-page">
    <PageTitle eyebrow="PRODUCTION OVERVIEW" title="加工總覽" />
    {loading ? <PageSkeleton /> : <div className="production-chart-list">
      {productCharts.length
        ? productCharts.map((chart) => <ProductionQuantityChart key={chart.productCode} productCode={chart.productCode} days={chart.days} total={chart.total} />)
        : <article className="chart-card production-chart-card"><div className="production-chart-empty"><ClipboardList />尚未建立產品編號</div></article>}
    </div>}
  </section>
}

function ProductionQuantityChart({ productCode, days, total }: { productCode: string; days: { date: string; quantity: number }[]; total: number }) {
  const chartScrollRef = useRef<HTMLDivElement>(null)
  const maximum = Math.max(1, ...days.map((day) => day.quantity))
  const roundedMaximum = Math.ceil(maximum / 10) * 10 || 10
  /* 「加工管理系統總覽」的「全部日期長條圖寬度」：日期越多，圖表越寬並由外層提供左右捲動。 */
  const chartWidth = Math.max(760, 82 + days.length * 92)
  const step = (chartWidth - 82) / Math.max(days.length, 1)
  const barWidth = Math.min(42, Math.max(20, step * .48))
  const dateRange = days.length ? ` ${formatProductionChartDate(days[0].date, true)} ～ ${formatProductionChartDate(days[days.length - 1].date, true)}` : '尚無加工日期'
  const latestChartDate = days[days.length - 1]?.date || ''
  /* 「加工總覽」的「每個產品長條圖」：初次顯示或新增日期後，自動捲到最右側的最新一筆。 */
  useEffect(() => {
    const frame = window.requestAnimationFrame(() => {
      const chart = chartScrollRef.current
      if (chart) chart.scrollLeft = chart.scrollWidth
    })
    return () => window.cancelAnimationFrame(frame)
  }, [days.length, latestChartDate])
  return <article className="chart-card production-chart-card">
    <div className="chart-head production-chart-head"><span>{dateRange}</span><div><strong>{productCode}</strong><small className="production-chart-total"><span>總加工數量</span><b>{formatNumber(total)}</b></small></div></div>
    {days.length ? <div ref={chartScrollRef} className="chart-wrap production-chart-wrap">
      <svg viewBox={`0 0 ${chartWidth} 270`} style={{ width: `${chartWidth}px`, minWidth: '100%' }} role="img" aria-label={`${productCode} 依日期統計的加工數量長條圖`}>
        {[0, 1, 2, 3, 4].map((index) => { const y = 22 + index * 44; const value = Math.round(roundedMaximum * (1 - index / 4)); return <g key={index}><line x1="52" x2={chartWidth - 18} y1={y} y2={y}/><text x="43" y={y + 5} textAnchor="end">{formatNumber(value)}</text></g> })}
        {days.map((day, index) => {
          const x = 58 + step * index + (step - barWidth) / 2
          const height = day.quantity / roundedMaximum * 176
          return <g key={day.date}><rect className="production-bar" x={x} y={198 - height} width={barWidth} height={height} rx="3"/><text className="production-bar-value" x={x + barWidth / 2} y={Math.max(16, 190 - height)} textAnchor="middle">{formatNumber(day.quantity)}</text><text x={x + barWidth / 2} y="224" textAnchor="middle">{formatProductionChartDate(day.date)}</text></g>
        })}
      </svg>
    </div> : <div className="production-chart-empty"><ClipboardList />此產品尚無加工紀錄</div>}
  </article>
}

function ProductionEntryPage({ records, productCodes, machines, loading, onSave, onDeleteLast }: { records: ProductionRecord[]; productCodes: string[]; machines: string[]; loading: boolean; onSave: (input: ProductionInput) => Promise<void>; onDeleteLast: (record: ProductionRecord) => Promise<void> }) {
  const [form, setForm] = useState<ProductionInput>({ work_date: previousLocalDate(), product_code: '', machine: '', hours: 0, minutes: 0, quantity: 0, notes: '' })
  const [showTimeCalculator, setShowTimeCalculator] = useState(false)
  const [formError, setFormError] = useState('')
  const [saving, setSaving] = useState(false)
  const latest = useMemo(() => [...records].sort((a, b) => `${b.created_at}-${b.id}`.localeCompare(`${a.created_at}-${a.id}`))[0], [records])
  useEffect(() => {
    setForm((current) => ({
      ...current,
      product_code: current.product_code && !productCodes.includes(current.product_code) ? '' : current.product_code,
      machine: current.machine && !machines.includes(current.machine) ? '' : current.machine,
    }))
  }, [productCodes, machines])

  const submit = async (event: FormEvent) => {
    event.preventDefault(); setFormError('')
    if (!form.product_code.trim()) { setFormError('請輸入或選擇產品編號'); return }
    if (!form.machine.trim()) { setFormError('請輸入或選擇機台'); return }
    if (!Number.isInteger(form.hours) || form.hours < 0 || !Number.isInteger(form.minutes) || form.minutes < 0 || form.minutes > 59) { setFormError('工時格式不正確，分鐘需介於 0 到 59'); return }
    if (form.hours * 60 + form.minutes <= 0) { setFormError('加工時間必須大於 0 分鐘'); return }
    if (!Number.isInteger(form.quantity) || form.quantity <= 0) { setFormError('加工數量必須是大於 0 的整數'); return }
    setSaving(true)
    try {
      await onSave({ ...form, product_code: form.product_code.trim().toUpperCase(), machine: form.machine.trim().toUpperCase() })
      setForm((current) => ({ ...current, hours: 0, minutes: 0, quantity: 0, notes: '' }))
    } catch (error) { setFormError(messageFor(error)) }
    finally { setSaving(false) }
  }

  return <section className="page-content production-system-page production-entry-page">
    <PageTitle eyebrow="PRODUCTION ENTRY" title="加工數量登記" />
    <form className="production-card entry-card" onSubmit={submit} onKeyDown={preventInputEnterSubmit}>
      <div className="production-field-row"><label>日期</label><input type="date" value={form.work_date} onChange={(e) => setForm({ ...form, work_date: e.target.value })} required /></div>
      <div className="production-field-row"><label>產品編號</label><select value={form.product_code} onChange={(e) => setForm({ ...form, product_code: e.target.value })} required><option value="">請選擇產品編號</option>{productCodes.map((value) => <option key={value} value={value}>{value}</option>)}</select></div>
      <div className="production-field-row"><label>機台</label><select value={form.machine} onChange={(e) => setForm({ ...form, machine: e.target.value })} required><option value="">請選擇機台</option>{machines.map((value) => <option key={value} value={value}>{value}</option>)}</select></div>

      <div className="time-entry-panel">
        <div className="duration-fields compact">
          <label className="actual-hours-field"><span>加工時間（時）</span><input className="actual-hours-input" type="number" inputMode="numeric" pattern="[0-9]*" min="0" step="1" value={form.hours || ''} onChange={(e) => setForm({ ...form, hours: Math.max(0, Number(e.target.value)) })} aria-label="加工時數" /></label>
          <label className="actual-minutes-field"><span>加工時間（分）</span><input className="actual-minutes-input" type="number" inputMode="numeric" pattern="[0-9]*" min="0" max="59" step="1" value={form.minutes || ''} onChange={(e) => setForm({ ...form, minutes: Math.max(0, Number(e.target.value)) })} aria-label="加工分鐘" /></label>
          <div><span>實際總時間</span><strong>{formatNumber(form.hours * 60 + form.minutes)} 分鐘</strong></div>
        </div>
      </div>

      <div className="production-field-row"><label>數量</label><input type="number" inputMode="numeric" pattern="[0-9]*" min="1" step="1" value={form.quantity || ''} onChange={(e) => setForm({ ...form, quantity: Math.max(0, Number(e.target.value)) })} placeholder="輸入完成數量" required /></div>
      <div className="production-field-row note-row"><label>備註</label><textarea value={form.notes} onChange={(e) => setForm({ ...form, notes: e.target.value })} placeholder="選填" /></div>
      {formError && <div className="form-error production-error">{formError}</div>}
      <button className="production-action save-action" disabled={saving} type="submit"><Save />{saving ? '儲存中…' : '送出登記'}</button>
      <button className="production-action delete-action" disabled={loading || !latest} type="button" onClick={() => latest && void onDeleteLast(latest)}><Trash2 />刪除最後一筆</button>
      <button className="production-action time-tool-action" type="button" onClick={() => setShowTimeCalculator(true)}><Calculator />加工時間計算</button>
    </form>
    {showTimeCalculator && <ProductionTimeCalculatorModal initialHours={form.hours} initialMinutes={form.minutes} close={() => setShowTimeCalculator(false)} apply={(hours, minutes) => { setForm((current) => ({ ...current, hours, minutes })); setShowTimeCalculator(false) }} />}
  </section>
}

type ProductionQueryMode = 'range' | 'latest' | 'all'

/* 「加工數量查詢」與「資料匯出與刪除」共用的加工統計，數值與下載 Excel 的合計欄位一致。 */
function calculateProductionTotals(records: ProductionRecord[]) {
  const totalQuantity = records.reduce((sum, record) => sum + record.quantity, 0)
  const totalMinutes = records.reduce((sum, record) => sum + record.total_minutes, 0)
  const totalDays = new Set(records.map((record) => record.work_date)).size
  return {
    totalQuantity,
    totalMinutes,
    totalDays,
    dailyAverage: totalDays ? Math.round(totalQuantity / totalDays) : 0,
    pieceSeconds: totalQuantity ? Math.round(totalMinutes * 60 / totalQuantity) : 0,
  }
}

function ProductionTotals({ totals }: { totals: ReturnType<typeof calculateProductionTotals> }) {
  return <div className="production-totals">
    <div><span>總天數</span><strong>{formatNumber(totals.totalDays)}</strong></div>
    <div><span>總時間</span><strong>{formatNumber(totals.totalMinutes)} <small>分</small></strong></div>
    <div><span>1天平均</span><strong>{formatNumber(totals.dailyAverage)} <small>顆</small></strong></div>
    <div><span>1顆平均</span><strong>{formatNumber(totals.pieceSeconds)} <small>秒</small></strong></div>
    <div className="production-total-quantity"><span>總數量</span><strong>{formatNumber(totals.totalQuantity)}</strong></div>
  </div>
}

function ProductionQueryPage({ records, productCodes, machines, loading, readOnly, onUpdateRecord, onDeleteRecord }: { records: ProductionRecord[]; productCodes: string[]; machines: string[]; loading: boolean; readOnly: boolean; onUpdateRecord: (record: ProductionRecord, input: ProductionInput) => Promise<void>; onDeleteRecord: (record: ProductionRecord) => Promise<boolean> }) {
  const [productCode, setProductCode] = useState('')
  const [from, setFrom] = useState(`${today().slice(0, 8)}01`)
  const [to, setTo] = useState(today())
  const [mode, setMode] = useState<ProductionQueryMode | null>(null)
  const [editingRecord, setEditingRecord] = useState<ProductionRecord | null>(null)
  const [queryError, setQueryError] = useState('')
  const [deletingRecordId, setDeletingRecordId] = useState('')
  const selectedRecords = useMemo(() => records
    .filter((record) => record.product_code.toLowerCase() === productCode.trim().toLowerCase())
    .sort((a, b) => `${b.work_date}-${b.created_at}-${b.id}`.localeCompare(`${a.work_date}-${a.created_at}-${a.id}`)), [records, productCode])
  const results = useMemo(() => {
    if (!mode) return []
    if (mode === 'range') return selectedRecords.filter((record) => (!from || record.work_date >= from) && (!to || record.work_date <= to))
    if (mode === 'latest') {
      const latestDate = selectedRecords[0]?.work_date
      return latestDate ? selectedRecords.filter((record) => record.work_date === latestDate) : []
    }
    return selectedRecords
  }, [mode, selectedRecords, from, to])
  const totals = useMemo(() => calculateProductionTotals(results), [results])
  const runQuery = (nextMode: ProductionQueryMode) => {
    setQueryError('')
    if (!productCode.trim()) { setMode(null); setQueryError('請先輸入或選擇產品編號'); return }
    if (nextMode === 'range' && from && to && from > to) { setMode(null); setQueryError('開始日期不能晚於結束日期'); return }
    setMode(nextMode)
  }

  const deleteOneRecord = async (record: ProductionRecord) => {
    setDeletingRecordId(record.id); setQueryError('')
    await onDeleteRecord(record)
    setDeletingRecordId('')
  }

  return <section className="page-content production-system-page production-query-page">
    <PageTitle eyebrow="PRODUCTION LOOKUP" title="加工數量查詢" />
    <div className="production-card query-card">
      <div className="query-fields">
        <div className="production-field-row"><label>開始日期</label><input type="date" value={from} onChange={(e) => setFrom(e.target.value)} /></div>
        <div className="production-field-row"><label>結束日期</label><input type="date" value={to} onChange={(e) => setTo(e.target.value)} /></div>
        <div className="production-field-row"><label>產品編號</label><select value={productCode} onChange={(e) => { setProductCode(e.target.value); setMode(null); setQueryError('') }}><option value="">請選擇產品編號</option>{productCodes.map((value) => <option key={value} value={value}>{value}</option>)}</select></div>
      </div>
      <div className="query-actions">
        <button className="query-button latest" onClick={() => runQuery('latest')}><Clock3 />查詢最後一筆</button>
        <button className="query-button all" onClick={() => runQuery('all')}><Search />查詢全部</button>
        <button className="query-button range" onClick={() => runQuery('range')}><CalendarDays />日期查詢</button>
      </div>
      {queryError && <div className="form-error production-error">{queryError}</div>}
    </div>

    {loading ? <PageSkeleton /> : mode && <div className="production-results">
      <div className="result-title"><span>產品編號</span><strong>{productCode.trim().toUpperCase()}</strong><small>{mode === 'range' ? `${formatDate(from)} ～ ${formatDate(to)}` : mode === 'latest' ? `最後日期 ${formatDate(results[0]?.work_date || '')}・全部機台` : '全部紀錄'}</small></div>
      {results.length === 0 ? <div className="production-empty"><ClipboardList />沒有符合條件的加工紀錄</div> : <>
        <div className="production-table-wrap"><table className={`production-table ${readOnly ? 'guest-read-only-table' : ''}`}><thead><tr><th>日期</th><th>機台</th><th>時</th><th>分</th><th>數量</th><th>總時間</th><th>備註</th>{!readOnly && <th className="record-action-column" aria-label="操作" />}</tr></thead><tbody>{results.map((record) => <tr key={record.id}><td>{formatDate(record.work_date)}</td><td>{record.machine}</td><td>{record.hours}</td><td>{record.minutes}</td><td>{formatNumber(record.quantity)}</td><td>{formatNumber(record.total_minutes)}</td><td>{record.notes || '—'}</td>{!readOnly && <td><div className="production-row-actions"><button className="production-row-edit" onClick={() => setEditingRecord(record)} disabled={Boolean(deletingRecordId)} aria-label={`編輯 ${formatDate(record.work_date)} ${record.machine} 加工紀錄`}><Pencil />編輯</button><button className="production-row-delete" onClick={() => void deleteOneRecord(record)} disabled={Boolean(deletingRecordId)} aria-label={`刪除 ${formatDate(record.work_date)} ${record.machine} 加工紀錄`}><Trash2 />{deletingRecordId === record.id ? '刪除中' : '刪除'}</button></div></td>}</tr>)}</tbody></table></div>
        <ProductionTotals totals={totals} />
      </>}
    </div>}
    {!readOnly && editingRecord && <ProductionRecordEditModal record={editingRecord} productCodes={productCodes} machines={machines} close={() => setEditingRecord(null)} save={async (input) => { await onUpdateRecord(editingRecord, input); setEditingRecord(null) }} />}
  </section>
}

function ProductionSettingsPage({ openOption }: { openOption: (type: ProductionOptionType) => void }) {
  return <section className="page-content production-system-page production-settings-page">
    <PageTitle eyebrow="PRODUCTION SETTINGS" title="基礎資料設定" />
    <div className="production-settings-grid">
      <button className="production-setting-card product-setting" type="button" onClick={() => openOption('product')}>
        <Wrench /><span><strong>管理產品編號</strong><small>新增、編輯或刪除加工產品編號</small></span><Pencil />
      </button>
      <button className="production-setting-card machine-setting" type="button" onClick={() => openOption('machine')}>
        <Wrench /><span><strong>管理機台</strong><small>新增或刪除加工機台名稱</small></span><Pencil />
      </button>
    </div>
  </section>
}

function ProductionForecastPage({ records, productCodes, loading }: { records: ProductionRecord[]; productCodes: string[]; loading: boolean }) {
  const [productCode, setProductCode] = useState(readForecastProductSelection)
  const selectedRecords = useMemo(() => records.filter((record) => record.product_code.toLowerCase() === productCode.toLowerCase()), [records, productCode])
  /* 「訂單完工預估」的「產品編號」：切換功能、瀏覽器分頁或重新整理後保留上次選擇。 */
  useEffect(() => { saveForecastProductSelection(productCode) }, [productCode])
  useEffect(() => {
    if (!loading && productCode && !productCodes.some((value) => value.toLowerCase() === productCode.toLowerCase())) setProductCode('')
  }, [loading, productCode, productCodes])
  return <section className="page-content production-system-page production-forecast-page">
    <PageTitle eyebrow="ORDER FORECAST" title="訂單完工預估" />
    <div className="production-card production-tool-page-card">
      <div className="production-field-row"><label>產品編號</label><select value={productCode} onChange={(event) => setProductCode(event.target.value)} disabled={loading}><option value="">請選擇產品編號</option>{productCodes.map((value) => <option key={value} value={value}>{value}</option>)}</select></div>
      {loading ? <PageSkeleton /> : productCode ? <OrderForecastPanel key={productCode} productCode={productCode} records={selectedRecords} /> : <div className="production-tool-placeholder"><Calculator />請先選擇要估算的產品編號</div>}
    </div>
  </section>
}

function ProductionDataManagementPage({ records, productCodes, loading, onDeleteProductRecords }: { records: ProductionRecord[]; productCodes: string[]; loading: boolean; onDeleteProductRecords: (productCode: string, count: number) => Promise<boolean> }) {
  const [productCode, setProductCode] = useState('')
  const [deleting, setDeleting] = useState(false)
  const [exporting, setExporting] = useState(false)
  const [actionError, setActionError] = useState('')
  const selectedRecords = useMemo(() => records
    .filter((record) => record.product_code.toLowerCase() === productCode.toLowerCase())
    .sort((a, b) => `${a.work_date}-${a.machine}-${a.created_at}-${a.id}`.localeCompare(`${b.work_date}-${b.machine}-${b.created_at}-${b.id}`, 'zh-Hant')), [records, productCode])
  const recordCount = selectedRecords.length
  const totals = useMemo(() => calculateProductionTotals(selectedRecords), [selectedRecords])
  const downloadExcel = async () => {
    if (!productCode || !selectedRecords.length) return
    setExporting(true); setActionError('')
    try { await exportProductionXlsx(productCode, selectedRecords) }
    catch (error) { setActionError(`Excel 產生失敗：${messageFor(error)}`) }
    finally { setExporting(false) }
  }
  const removeSelectedProduct = async () => {
    if (!productCode) return
    setDeleting(true); setActionError('')
    try {
      if (await onDeleteProductRecords(productCode, recordCount)) setProductCode('')
    } finally { setDeleting(false) }
  }
  return <section className="page-content production-system-page production-delete-page">
    <PageTitle eyebrow="DATA EXPORT & DELETE" title="資料匯出與刪除" />
    <div className="production-card production-tool-page-card production-delete-card">
      <div className="production-field-row"><label>產品編號</label><select value={productCode} onChange={(event) => { setProductCode(event.target.value); setActionError('') }} disabled={loading || deleting || exporting}><option value="">請選擇產品編號</option>{productCodes.map((value) => <option key={value} value={value}>{value}</option>)}</select></div>
      <div className="production-data-actions">
        <button className="query-button export" type="button" onClick={() => void downloadExcel()} disabled={!productCode || !recordCount || loading || deleting || exporting}><Download />{exporting ? '產生中…' : '下載 EXCEL'}</button>
        <button className="query-button danger" type="button" onClick={() => void removeSelectedProduct()} disabled={!productCode || loading || deleting || exporting}><Trash2 />{deleting ? '刪除中…' : `永久刪除${productCode ? `（${recordCount} 筆）` : ''}`}</button>
      </div>
      {actionError && <div className="form-error production-error">{actionError}</div>}
      <div className="production-delete-warning"><Trash2 /><div><strong>刪除會永久移除產品編號與全部加工紀錄</strong><span>「下載 EXCEL」只會下載資料，不會修改雲端內容；「永久刪除」仍需輸入登入密碼確認。</span></div></div>
    </div>
    {loading ? <PageSkeleton /> : productCode ? <div className="production-results production-export-preview">
      <div className="result-title"><span>產品編號</span><strong>{productCode.trim().toUpperCase()}</strong><small>全部加工紀錄・共 {formatNumber(recordCount)} 筆</small></div>
      {selectedRecords.length === 0 ? <div className="production-empty"><ClipboardList />此產品目前沒有加工紀錄，仍可永久刪除產品編號</div> : <>
        <div className="production-table-wrap"><table className="production-table production-export-preview-table"><thead><tr><th>日期</th><th>機台</th><th>時</th><th>分</th><th>數量</th><th>總時間</th><th>備註</th></tr></thead><tbody>{selectedRecords.map((record) => <tr key={record.id}><td>{formatDate(record.work_date)}</td><td>{record.machine}</td><td>{record.hours}</td><td>{record.minutes}</td><td>{formatNumber(record.quantity)}</td><td>{formatNumber(record.total_minutes)}</td><td>{record.notes || '—'}</td></tr>)}</tbody></table></div>
        <ProductionTotals totals={totals} />
      </>}
    </div> : <div className="production-card production-tool-placeholder production-data-placeholder"><FileSpreadsheet />請先選擇產品編號，即可預覽全部資料</div>}
  </section>
}

function ProductionRecordEditModal({ record, productCodes, machines, close, save }: { record: ProductionRecord; productCodes: string[]; machines: string[]; close: () => void; save: (input: ProductionInput) => Promise<void> }) {
  const [form, setForm] = useState<ProductionInput>({
    work_date: record.work_date,
    product_code: record.product_code,
    machine: record.machine,
    hours: record.hours,
    minutes: record.minutes,
    quantity: record.quantity,
    notes: record.notes,
  })
  const [busy, setBusy] = useState(false)
  const [formError, setFormError] = useState('')
  const availableProductCodes = productCodes.includes(form.product_code) ? productCodes : [form.product_code, ...productCodes]
  const availableMachines = machines.includes(form.machine) ? machines : [form.machine, ...machines]

  const submit = async (event: FormEvent) => {
    event.preventDefault(); setFormError('')
    if (!form.product_code.trim() || !form.machine.trim()) { setFormError('請選擇產品編號與機台'); return }
    if (!Number.isInteger(form.hours) || form.hours < 0 || !Number.isInteger(form.minutes) || form.minutes < 0 || form.minutes > 59) { setFormError('工時格式不正確，分鐘需介於 0 到 59'); return }
    if (form.hours * 60 + form.minutes <= 0) { setFormError('加工時間必須大於 0 分鐘'); return }
    if (!Number.isInteger(form.quantity) || form.quantity <= 0) { setFormError('加工數量必須是大於 0 的整數'); return }
    setBusy(true)
    try { await save({ ...form, product_code: form.product_code.trim().toUpperCase(), machine: form.machine.trim().toUpperCase() }) }
    catch (error) { setFormError(messageFor(error)); setBusy(false) }
  }

  return <Modal title="編輯加工紀錄" close={close} size="production-edit-modal">
    <form onSubmit={submit} onKeyDown={preventInputEnterSubmit}>
      <div className="form-grid production-edit-grid">
        <Field label="日期"><input required type="date" value={form.work_date} onChange={(event) => setForm({ ...form, work_date: event.target.value })} /></Field>
        <Field label="產品編號"><select required value={form.product_code} onChange={(event) => setForm({ ...form, product_code: event.target.value })}>{availableProductCodes.map((value) => <option key={value} value={value}>{value}</option>)}</select></Field>
        <Field label="機台"><select required value={form.machine} onChange={(event) => setForm({ ...form, machine: event.target.value })}>{availableMachines.map((value) => <option key={value} value={value}>{value}</option>)}</select></Field>
        <Field label="工時（時）"><input required type="number" inputMode="numeric" pattern="[0-9]*" min="0" step="1" value={form.hours} onChange={(event) => setForm({ ...form, hours: Math.max(0, Number(event.target.value)) })} /></Field>
        <Field label="工時（分）"><input required type="number" inputMode="numeric" pattern="[0-9]*" min="0" max="59" step="1" value={form.minutes} onChange={(event) => setForm({ ...form, minutes: Math.max(0, Number(event.target.value)) })} /></Field>
        <Field label="數量"><input required type="number" inputMode="numeric" pattern="[0-9]*" min="1" step="1" value={form.quantity} onChange={(event) => setForm({ ...form, quantity: Math.max(0, Number(event.target.value)) })} /></Field>
      </div>
      <Field label="備註" wide><textarea value={form.notes} onChange={(event) => setForm({ ...form, notes: event.target.value })} placeholder="選填" /></Field>
      <div className="production-edit-total"><span>總加工時間</span><strong>{formatNumber(form.hours * 60 + form.minutes)} 分鐘</strong></div>
      {formError && <p className="form-error">{formError}</p>}
      <button className="button modal-submit primary" disabled={busy}>{busy ? '儲存中…' : '儲存修改'}</button>
    </form>
  </Modal>
}

function ProductionTimeCalculatorModal({ initialHours, initialMinutes, close, apply }: { initialHours: number; initialMinutes: number; close: () => void; apply: (hours: number, minutes: number) => void }) {
  const [values, setValues] = useState({ totalHours: initialHours, totalMinutes: initialMinutes, deductHours: 0, deductMinutes: 0 })
  const [result, setResult] = useState<{ hours: number; minutes: number } | null>(null)
  const [error, setError] = useState('')
  const calculate = () => {
    setError('')
    const inputs = [values.totalHours, values.totalMinutes, values.deductHours, values.deductMinutes]
    if (inputs.some((value) => !Number.isInteger(value) || value < 0)) { setResult(null); setError('工時只能輸入 0 以上的整數'); return }
    if (values.totalMinutes > 59 || values.deductMinutes > 59) { setResult(null); setError('分鐘需介於 0 到 59'); return }
    const difference = values.totalHours * 60 + values.totalMinutes - values.deductHours * 60 - values.deductMinutes
    if (difference <= 0) { setResult(null); setError('總工時必須大於扣除工時'); return }
    setResult({ hours: Math.floor(difference / 60), minutes: difference % 60 })
  }
  return <Modal title="加工時間計算" close={close} size="time-tool-modal">
    <p className="modal-intro">這裡計算的是工時長度相減，例如 11 小時 5 分減去 3 小時 10 分，結果是 7 小時 55 分。</p>
    <div className="duration-calculation modal-calculation">
      <div className="duration-row"><strong>總工時</strong><label><input type="number" inputMode="numeric" pattern="[0-9]*" min="0" step="1" value={values.totalHours || ''} onChange={(e) => { setValues({ ...values, totalHours: Math.max(0, Number(e.target.value)) }); setResult(null) }} aria-label="總工時小時" /><span>時</span></label><label><input type="number" inputMode="numeric" pattern="[0-9]*" min="0" max="59" step="1" value={values.totalMinutes || ''} onChange={(e) => { setValues({ ...values, totalMinutes: Math.max(0, Number(e.target.value)) }); setResult(null) }} aria-label="總工時分鐘" /><span>分</span></label></div>
      <div className="duration-operator">減去</div>
      <div className="duration-row"><strong>扣除工時</strong><label><input type="number" inputMode="numeric" pattern="[0-9]*" min="0" step="1" value={values.deductHours || ''} onChange={(e) => { setValues({ ...values, deductHours: Math.max(0, Number(e.target.value)) }); setResult(null) }} aria-label="扣除工時小時" /><span>時</span></label><label><input type="number" inputMode="numeric" pattern="[0-9]*" min="0" max="59" step="1" value={values.deductMinutes || ''} onChange={(e) => { setValues({ ...values, deductMinutes: Math.max(0, Number(e.target.value)) }); setResult(null) }} aria-label="扣除工時分鐘" /><span>分</span></label></div>
      <button className="button calculator-button" type="button" onClick={calculate}><Calculator />計算實際工時</button>
    </div>
    {error && <div className="form-error production-error modal-tool-error">{error}</div>}
    <div className={`time-result ${result ? 'has-result' : ''}`}><span>計算結果</span><strong>{result ? `${result.hours} 小時 ${result.minutes} 分` : '尚未計算'}</strong></div>
    <div className="modal-tool-actions"><button className="button" type="button" onClick={close}>取消</button><button className="button primary" type="button" disabled={!result} onClick={() => result && apply(result.hours, result.minutes)}><Save />套用實際工時</button></div>
  </Modal>
}

function OrderForecastPanel({ productCode, records }: { productCode: string; records: ProductionRecord[] }) {
  const historyDateBounds = useMemo(() => {
    const dates = records.map((record) => record.work_date).filter(Boolean).sort()
    return { first: dates[0] || '', last: dates[dates.length - 1] || '' }
  }, [records])
  const totalCompletedQuantity = useMemo(() => records.reduce((sum, record) => sum + record.quantity, 0), [records])
  const storedDraft = useMemo(() => readForecastDraft(productCode), [productCode])
  const [orderQuantity, setOrderQuantity] = useState(storedDraft?.orderQuantity ?? 0)
  const [dueDate, setDueDate] = useState(storedDraft?.dueDate ?? '')
  const [capacityScope, setCapacityScope] = useState<ForecastCapacityScope>(storedDraft?.capacityScope ?? 'all')
  const [historyFrom, setHistoryFrom] = useState(storedDraft?.historyFrom || historyDateBounds.first)
  const [historyTo, setHistoryTo] = useState(storedDraft?.historyTo || historyDateBounds.last)
  const historyRangeInvalid = capacityScope === 'range' && Boolean(historyFrom && historyTo && historyFrom > historyTo)
  const scopedRecords = useMemo(() => capacityScope === 'range'
    ? records.filter((record) => (!historyFrom || record.work_date >= historyFrom) && (!historyTo || record.work_date <= historyTo) && !historyRangeInvalid)
    : records, [records, capacityScope, historyFrom, historyTo, historyRangeInvalid])
  /* 產能只採用未來排程也會使用的日期：星期一至星期六，並排除國定假日與補假。 */
  const capacityRecords = useMemo(() => scopedRecords.filter((record) => {
    const weekday = parseLocalDate(record.work_date).getDay()
    return !isTaiwanPublicHoliday(record.work_date) && weekday >= 1 && weekday <= 6
  }), [scopedRecords])
  const capacityHistory = useMemo(() => {
    const dailyMap = new Map<string, { quantity: number; machines: Set<string> }>()
    capacityRecords.forEach((record) => {
      const current = dailyMap.get(record.work_date) || { quantity: 0, machines: new Set<string>() }
      current.quantity += record.quantity
      current.machines.add(record.machine)
      dailyMap.set(record.work_date, current)
    })
    const dailySamples = [...dailyMap.entries()].map(([date, value]) => ({ date, quantity: value.quantity, machineCount: value.machines.size }))
    const workDates = dailySamples.map((sample) => sample.date)
    const capacitySampleQuantity = capacityRecords.reduce((sum, record) => sum + record.quantity, 0)
    const totalMachineDays = dailySamples.reduce((sum, sample) => sum + sample.machineCount, 0)
    const historyDays = workDates.length
    const machineCountFrequency = new Map<number, number>()
    dailySamples.forEach((sample) => machineCountFrequency.set(sample.machineCount, (machineCountFrequency.get(sample.machineCount) || 0) + 1))
    const defaultMachineCount = [...machineCountFrequency.entries()].sort((left, right) => right[1] - left[1] || left[0] - right[0])[0]?.[0] || 1
    return {
      historyDays,
      capacitySampleQuantity,
      dailySamples,
      averageQuantityPerMachineDay: totalMachineDays ? capacitySampleQuantity / totalMachineDays : 0,
      saturdayDays: workDates.filter((value) => parseLocalDate(value).getDay() === 6).length,
      includesSaturday: workDates.some((value) => parseLocalDate(value).getDay() === 6),
      defaultMachineCount,
    }
  }, [capacityRecords])
  const [machineCount, setMachineCount] = useState(storedDraft?.machineCount || capacityHistory.defaultMachineCount)
  const previousCapacityBasis = useRef(`${capacityScope}|${historyFrom}|${historyTo}`)
  useEffect(() => {
    const basis = `${capacityScope}|${historyFrom}|${historyTo}`
    if (basis === previousCapacityBasis.current) return
    previousCapacityBasis.current = basis
    setMachineCount(capacityHistory.defaultMachineCount)
  }, [capacityScope, historyFrom, historyTo, capacityHistory.defaultMachineCount])
  useEffect(() => {
    saveForecastDraft(productCode, { orderQuantity, dueDate, capacityScope, historyFrom, historyTo, machineCount })
  }, [productCode, orderQuantity, dueDate, capacityScope, historyFrom, historyTo, machineCount])

  const remainingQuantity = Math.max(orderQuantity - totalCompletedQuantity, 0)
  const matchingMachineSamples = capacityHistory.dailySamples.filter((sample) => sample.machineCount === machineCount)
  const estimatedDailyQuantity = matchingMachineSamples.length
    ? matchingMachineSamples.reduce((sum, sample) => sum + sample.quantity, 0) / matchingMachineSamples.length
    : capacityHistory.averageQuantityPerMachineDay * Math.max(machineCount, 0)
  const forecastResult = useMemo(() => {
    if (estimatedDailyQuantity <= 0) return { requiredWorkDays: 0, estimatedDate: '', availableWorkDays: 0, availableQuantity: 0, canMeetDeadline: null as boolean | null, quantityDifference: null as number | null }
    const isScheduledDate = (date: Date) => {
      const dateValue = localDateValue(date)
      const weekday = date.getDay()
      return !isTaiwanPublicHoliday(dateValue) && (weekday >= 1 && weekday <= 5 || capacityHistory.includesSaturday && weekday === 6)
    }
    let pending = remainingQuantity
    let requiredWorkDays = 0
    let estimatedDate = remainingQuantity <= 0 ? today() : ''
    const cursor = parseLocalDate(today())
    for (let guard = 0; pending > 0 && guard < 3660; guard += 1) {
      cursor.setDate(cursor.getDate() + 1)
      if (!isScheduledDate(cursor)) continue
      pending -= estimatedDailyQuantity
      requiredWorkDays += 1
      if (pending <= 0) estimatedDate = localDateValue(cursor)
    }
    let availableWorkDays = 0
    if (dueDate && dueDate >= today()) {
      const dueCursor = parseLocalDate(today())
      const due = parseLocalDate(dueDate)
      while (dueCursor < due) {
        dueCursor.setDate(dueCursor.getDate() + 1)
        if (isScheduledDate(dueCursor)) availableWorkDays += 1
      }
    }
    const availableQuantity = availableWorkDays * estimatedDailyQuantity
    const canMeetDeadline = dueDate ? remainingQuantity <= availableQuantity : null
    const quantityDifference = dueDate ? Math.round(availableQuantity - remainingQuantity) : null
    return { requiredWorkDays, estimatedDate, availableWorkDays, availableQuantity, canMeetDeadline, quantityDifference }
  }, [remainingQuantity, estimatedDailyQuantity, capacityHistory.includesSaturday, dueDate])
  const historyRangeText = capacityScope === 'range' ? `採用區間 ${formatDate(historyFrom)} ～ ${formatDate(historyTo)}` : '採用全部歷史紀錄'
  const holidayCalendarIncomplete = [dueDate, forecastResult.estimatedDate].some((value) => value && !hasTaiwanHolidayCalendar(value))
  const holidayCalendarYears = `${TAIWAN_HOLIDAY_CALENDAR_YEARS[0]}～${TAIWAN_HOLIDAY_CALENDAR_YEARS[TAIWAN_HOLIDAY_CALENDAR_YEARS.length - 1]}`
  return <div className="order-forecast production-forecast-panel">
    <div className="forecast-modal-product"><span>產品編號</span><strong>{productCode}</strong></div>
    <div className="forecast-head modal-forecast-head"><div><Calculator /><span><strong>依歷史產能預估完工日期</strong><small>依選定歷史資料的實際產量與機台數進行預估</small></span></div><label><span>訂單總數量</span><input className="order-quantity-input" type="number" inputMode="numeric" pattern="[0-9]*" min="1" step="1" value={orderQuantity || ''} onChange={(event) => setOrderQuantity(Math.max(0, Number(event.target.value)))} placeholder="輸入訂單數量" /></label></div>
    <div className="forecast-capacity-controls forecast-single-control">
      <label className="forecast-setting"><span>歷史產能依據</span><select value={capacityScope} onChange={(event) => setCapacityScope(event.target.value as ForecastCapacityScope)}><option value="all">全部歷史資料</option><option value="range">指定歷史日期區間</option></select></label>
    </div>
    {capacityScope === 'range' && <div className="forecast-history-range"><label className="forecast-setting"><span>歷史開始日期</span><input type="date" min={historyDateBounds.first || undefined} max={historyTo || historyDateBounds.last || undefined} value={historyFrom} onChange={(event) => setHistoryFrom(event.target.value)} /></label><label className="forecast-setting"><span>歷史結束日期</span><input type="date" min={historyFrom || historyDateBounds.first || undefined} max={historyDateBounds.last || undefined} value={historyTo} onChange={(event) => setHistoryTo(event.target.value)} /></label></div>}
    <div className="forecast-history-basis"><span>產能依據</span><strong>{capacityHistory.historyDays} 個加工日</strong><strong>總產量 {formatNumber(capacityHistory.capacitySampleQuantity)} 只</strong><strong>平均產量 {formatNumber(Math.round(estimatedDailyQuantity))} 只／日</strong><small>{historyRangeText}；{matchingMachineSamples.length ? `採用歷史中實際使用 ${machineCount} 台的 ${matchingMachineSamples.length} 個加工日` : `歷史沒有使用 ${machineCount} 台的加工日，暫依每台平均產量換算`}；{capacityHistory.includesSaturday ? `資料包含星期六（共 ${capacityHistory.saturdayDays} 天），未來排程會計入星期六` : '資料沒有星期六，未來排程只計算星期一至星期五'}</small></div>
    <div className="forecast-future-config">
      <label className="forecast-setting"><span>訂單交期</span><input type="date" min={today()} value={dueDate} onChange={(event) => setDueDate(event.target.value)} /></label>
      <label className="forecast-setting"><span>預計使用機台數</span><input type="number" inputMode="numeric" pattern="[0-9]*" min="1" step="1" value={machineCount || ''} onChange={(event) => setMachineCount(Math.max(0, Math.trunc(Number(event.target.value))))} /></label>
    </div>
    {historyRangeInvalid ? <p className="forecast-placeholder warning">歷史開始日期不能晚於歷史結束日期</p> : orderQuantity <= 0 ? <p className="forecast-placeholder">輸入訂單總數量後顯示預估結果</p> : !capacityHistory.historyDays || !capacityHistory.averageQuantityPerMachineDay ? <p className="forecast-placeholder warning">所選歷史產能依據沒有有效的加工數量與機台資料，暫時無法評估</p> : !machineCount ? <p className="forecast-placeholder warning">預計使用機台數必須大於 0</p> : <>
      <div className="forecast-grid forecast-result-grid"><div className="accent"><span>剩餘數量</span><strong>{formatNumber(remainingQuantity)}</strong><small>訂單總數量扣除全部歷史完成量</small></div><div><span>預估日產量</span><strong>{formatNumber(Math.round(estimatedDailyQuantity))}</strong><small>{formatNumber(machineCount)} 台機台</small></div><div><span>所需加工日</span><strong>{formatNumber(forecastResult.requiredWorkDays)}</strong><small>不含星期日與國定假日</small></div><div><span>預計完工日期</span><strong>{formatDate(forecastResult.estimatedDate)}</strong><small>{capacityHistory.includesSaturday ? '排程包含星期六' : '排程不包含星期六'}</small></div></div>
      {dueDate ? <div className={`forecast-deadline-assessment ${forecastResult.canMeetDeadline ? 'can-meet' : 'will-be-late'}`}>{forecastResult.canMeetDeadline ? <ShieldCheck /> : <Triangle />}<div><strong>{forecastResult.canMeetDeadline ? '可以在交期內完成' : '無法在交期內完成'}</strong><span>交期前可加工 {formatNumber(forecastResult.availableWorkDays)} 日；{forecastResult.canMeetDeadline ? `預計可多完成 ${formatNumber(Math.max(forecastResult.quantityDifference || 0, 0))} 只` : `預估尚缺 ${formatNumber(Math.abs(Math.min(forecastResult.quantityDifference || 0, 0)))} 只`}</span></div></div> : <p className="forecast-placeholder">輸入訂單交期後，系統會評估是否能在交期內完成</p>}
      {holidayCalendarIncomplete && <p className="forecast-placeholder warning">訂單交期或預估完工日超出目前載入的 {holidayCalendarYears} 官方行事曆，尚未公布年度的國定假日可能無法完整排除。</p>}
    </>}
    <p className="forecast-note">計算順序：預計使用機台數只接受整數。系統優先採用歷史中實際使用相同機台數的加工日平均產量；沒有相同機台數的歷史日期時，才依每台平均產量換算。若歷史產能依據中有星期六紀錄，未來排程會自動包含星期六；星期日、國定假日與補假不排產。訂單交期用來評估交期前可加工日數、可完成數量與尚缺數量；官方行事曆使用 {holidayCalendarYears} 年資料。</p>
  </div>
}

function ProductionOptionModal({ type, values, close, save, remove, rename }: { type: ProductionOptionType; values: string[]; close: () => void; save: (value: string) => Promise<void>; remove: (value: string) => Promise<void>; rename?: (currentValue: string, nextValue: string) => Promise<void> }) {
  const [value, setValue] = useState('')
  const [error, setError] = useState('')
  const [saving, setSaving] = useState(false)
  const [removing, setRemoving] = useState('')
  const [editing, setEditing] = useState('')
  const [editValue, setEditValue] = useState('')
  const [updating, setUpdating] = useState(false)
  const label = type === 'product' ? '產品編號' : '機台'
  const submit = async (event: FormEvent) => {
    event.preventDefault(); setError('')
    if (!value.trim()) { setError(`請輸入${label}`); return }
    setSaving(true)
    try { await save(value); setValue('') } catch (submitError) { setError(messageFor(submitError)) }
    finally { setSaving(false) }
  }
  const removeValue = async (target: string) => {
    setError(''); setRemoving(target)
    try { await remove(target) } catch (removeError) { setError(messageFor(removeError)) }
    finally { setRemoving('') }
  }
  const updateValue = async () => {
    if (!rename || !editing) return
    const normalized = editValue.trim().toUpperCase()
    setError('')
    if (!normalized) { setError('產品編號不能空白'); return }
    if (normalized === editing.trim().toUpperCase()) { setEditing(''); setEditValue(''); return }
    setUpdating(true)
    try { await rename(editing, normalized); setEditing(''); setEditValue('') }
    catch (updateError) { setError(messageFor(updateError)) }
    finally { setUpdating(false) }
  }
  return <Modal title={`管理${label}`} close={close}>
    <form onSubmit={submit} onKeyDown={preventInputEnterSubmit}>
      <Field label={label} wide><input autoFocus value={value} onChange={(event) => setValue(event.target.value)} placeholder={type === 'product' ? '產品編號<訂單編號>' : '例如 MT04'} /></Field>
      <p className="option-help">{type === 'product' ? '新增後會出現在加工登記選單；編輯會同步取代資料庫內全部既有加工紀錄的產品編號。' : '新增後會出現在加工登記選單；刪除不會影響過去的加工紀錄。'}</p>
      {error && <div className="form-error">{error}</div>}
      <button className="button primary option-add-button" disabled={saving || !!removing || updating} type="submit"><Plus />{saving ? '新增中…' : `新增${label}`}</button>
    </form>
    <div className="option-list-head"><span>目前的{label}</span><small>{values.length} 筆</small></div>
    <div className="option-list">
      {values.length === 0 ? <div className="option-empty">尚未新增任何{label}</div> : values.map((item) => <div className={`option-list-item ${editing === item ? 'is-editing' : ''}`} key={item}>
        {editing === item
          ? <input className="option-edit-input" autoFocus value={editValue} onChange={(event) => setEditValue(event.target.value)} onKeyDown={(event) => { if (event.key === 'Enter') { event.preventDefault(); void updateValue() } }} aria-label={`編輯產品編號 ${item}`} />
          : <span>{item}</span>}
        <div className="option-list-actions">
          {type === 'product' && (editing === item ? <><button className="option-save-edit" disabled={updating} onClick={() => void updateValue()} aria-label={`儲存產品編號 ${item}`} title="儲存修改">{updating ? '…' : <Save />}</button><button disabled={updating} onClick={() => { setEditing(''); setEditValue(''); setError('') }} aria-label="取消編輯" title="取消編輯"><X /></button></> : <button className="option-start-edit" disabled={!!removing || updating || !!editing} onClick={() => { setEditing(item); setEditValue(item); setError('') }} aria-label={`編輯產品編號 ${item}`} title="編輯產品編號"><Pencil /></button>)}
          <button className="danger-icon" disabled={!!removing || updating || !!editing} onClick={() => void removeValue(item)} aria-label={`刪除${label} ${item}`} title={`刪除${label}`}>{removing === item ? '…' : <Trash2 />}</button>
        </div>
      </div>)}
    </div>
  </Modal>
}

function FunctionCard({ icon, title, description, tone, onClick, disabled = false, disabledLabel = '即將推出' }: { icon: ReactNode; title: string; description: string; tone: string; onClick?: () => void; disabled?: boolean; disabledLabel?: string }) {
  return <button type="button" className={`function-card ${tone} ${disabled && disabledLabel ? 'is-future' : ''}`} onClick={onClick} disabled={disabled}><span className="function-icon">{icon}</span><span className="function-copy"><strong>{title}</strong><small>{description}</small></span>{disabled ? (disabledLabel && <em>{disabledLabel}</em>) : <ChevronDown className="function-arrow"/>}</button>
}

function PageTitle({ eyebrow, title, actions }: { eyebrow: string; title: string; actions?: ReactNode }) {
  return <div className="page-title"><div><span>{eyebrow}</span><h1>{title}</h1></div>{actions && <div className="title-actions">{actions}</div>}</div>
}

function StockButtons({ openStock }: { openStock: (type: MovementType) => void }) {
  return <><button className="button inbound" onClick={() => openStock('in')}><ArrowDownToLine />入庫</button><button className="button outbound" onClick={() => openStock('out')}><ArrowUpFromLine />出庫</button></>
}

function ScanButton({ openScanner }: { openScanner: () => void }) {
  return <button className="button scan-button" type="button" onClick={openScanner} aria-label="掃描條碼" title="掃描條碼"><ScanBarcode />掃描條碼</button>
}

function Dashboard({ products, movements, loading, openStock, openScanner, readOnly }: { products: Product[]; movements: Movement[]; loading: boolean; openStock: (t: MovementType) => void; openScanner: () => void; readOnly: boolean }) {
  const month = today().slice(0, 7)
  const inbound = movements.filter((m) => m.movement_type === 'in' && m.movement_date.startsWith(month))
  const outbound = movements.filter((m) => m.movement_type === 'out' && m.movement_date.startsWith(month))
  const stock = products.reduce((sum, p) => sum + p.stock_quantity, 0)
  const value = products.reduce((sum, p) => sum + p.stock_quantity * p.unit_price, 0)
  const inboundQty = inbound.reduce((sum, m) => sum + m.quantity, 0)
  const outboundQty = outbound.reduce((sum, m) => sum + m.quantity, 0)
  const purchased = inbound.reduce((sum, m) => sum + m.quantity * (m.unit_price || 0), 0)

  return <section className="page-content">
    <PageTitle eyebrow={formatDate(today())} title="庫存總覽" actions={!readOnly ? <><ScanButton openScanner={openScanner}/><StockButtons openStock={openStock} /></> : undefined} />
    {loading ? <PageSkeleton /> : <>
      <div className="stats-grid">
        <StatCard label="庫存總數" value={formatNumber(stock)} note={`${products.length} 項刀具`} />
        <StatCard label="庫存價值" value={`$${formatMoney(value)}`} note="以最新單價計算" tone="yellow" />
        <StatCard label="本月入庫" value={formatNumber(inboundQty)} note={`採購 $${formatMoney(purchased)}`} tone="green" />
        <StatCard label="本月出庫" value={formatNumber(outboundQty)} note={outboundQty ? `${outbound.length} 筆領用紀錄` : '本月尚無出庫'} tone="red" />
      </div>
      <MovementChart movements={movements} />
    </>}
  </section>
}

function StatCard({ label, value, note, tone = '' }: { label: string; value: string; note: string; tone?: string }) {
  return <article className={`stat-card ${tone}`}><span>{label}</span><strong>{value}</strong><small>{note}</small></article>
}

function MovementChart({ movements }: { movements: Movement[] }) {
  const days = useMemo(() => {
    const result = Array.from({ length: 14 }, (_, index) => {
      const date = new Date(`${today()}T12:00:00`); date.setDate(date.getDate() - (13 - index))
      const key = date.toISOString().slice(0, 10)
      return { key, label: `${date.getMonth() + 1}/${date.getDate()}`, in: 0, out: 0 }
    })
    movements.forEach((m) => { const item = result.find((d) => d.key === m.movement_date); if (item) item[m.movement_type] += m.quantity })
    return result
  }, [movements])
  const max = Math.max(10, ...days.flatMap((d) => [d.in, d.out]))
  const roundedMax = Math.ceil(max / 100) * 100 || 100
  return <article className="chart-card">
    <div className="chart-head"><span>近 14 天出入庫</span><div className="legend"><i className="green-dot" />入庫<i className="red-dot" />出庫</div></div>
    <div className="chart-wrap">
      <svg viewBox="0 0 760 255" role="img" aria-label="近十四天入庫與出庫數量長條圖">
        {[0, 1, 2, 3, 4].map((i) => { const y = 20 + i * 44; const value = Math.round(roundedMax * (1 - i / 4)); return <g key={i}><line x1="45" x2="744" y1={y} y2={y} /><text x="36" y={y + 5} textAnchor="end">{value}</text></g> })}
        {days.map((day, index) => {
          const x = 52 + index * 49.4; const inH = day.in / roundedMax * 176; const outH = day.out / roundedMax * 176
          return <g key={day.key}><rect className="bar-in" x={x} y={196 - inH} width="14" height={inH} rx="2"/><rect className="bar-out" x={x + 16} y={196 - outH} width="14" height={outH} rx="2"/><text x={x + 15} y="221" textAnchor="middle">{day.label}</text></g>
        })}
      </svg>
    </div>
  </article>
}

// 庫存管理的品名搜尋可省略連字號，條碼與其他欄位沿用原本的比對方式。
function matchesProductName(name: string | undefined, query: string) {
  const withoutHyphens = query.replaceAll('-', '')
  return Boolean(withoutHyphens && name?.toLowerCase().replaceAll('-', '').includes(withoutHyphens))
}

function MovementsPage({ products, movements, loading, openStock, openScanner, readOnly, onDelete }: { products: Product[]; movements: Movement[]; loading: boolean; openStock: (t: MovementType) => void; openScanner: () => void; readOnly: boolean; onDelete: (movement: Movement) => void }) {
  const [query, setQuery] = useState('')
  const [type, setType] = useState<'all' | MovementType>('all')
  const [from, setFrom] = useState('')
  const [to, setTo] = useState('')
  const normalizedQuery = query.trim().toLowerCase()
  const filtered = movements.filter((m) => {
    const product = products.find((p) => p.id === m.product_id) || m.product
    /* 「出入庫」的「搜尋輸入框」：條碼掃描或手動輸入後，同時比對條碼、編號、品名、廠牌、供應商與備註。 */
    const haystack = `${product?.barcode || ''} ${product?.code || ''} ${product?.name || ''} ${product?.brand || ''} ${m.supplier || ''} ${m.notes || ''} ${isInitialStockNote(m.notes) ? INITIAL_STOCK_NOTE : ''}`.toLowerCase()
    return (haystack.includes(normalizedQuery) || matchesProductName(product?.name, normalizedQuery)) && (type === 'all' || m.movement_type === type) && (!from || m.movement_date >= from) && (!to || m.movement_date <= to)
  })
  const inQty = filtered.filter((m) => m.movement_type === 'in').reduce((s, m) => s + m.quantity, 0)
  const outQty = filtered.filter((m) => m.movement_type === 'out').reduce((s, m) => s + m.quantity, 0)
  const purchased = filtered.filter((m) => m.movement_type === 'in').reduce((s, m) => s + m.quantity * (m.unit_price || 0), 0)
  return <section className="page-content">
    <PageTitle eyebrow="STOCK MOVEMENTS" title="出入庫" actions={!readOnly ? <><ScanButton openScanner={openScanner}/><StockButtons openStock={openStock} /></> : undefined} />
    <div className="filter-card">
      <SearchField value={query} setValue={setQuery} placeholder="掃描條碼或搜尋編號、品名、廠牌、供應商、備註" scanProducts={products} />
      <div className="filter-row">
        <label className="select-wrap"><select value={type} onChange={(e) => setType(e.target.value as 'all' | MovementType)}><option value="all">全部類型</option><option value="in">入庫</option><option value="out">出庫</option></select><ChevronDown /></label>
        <label className="date-wrap"><input type="date" value={from} onChange={(e) => setFrom(e.target.value)} aria-label="開始日期"/><CalendarDays /></label>
        <label className="date-wrap"><input type="date" value={to} onChange={(e) => setTo(e.target.value)} aria-label="結束日期"/><CalendarDays /></label>
      </div>
    </div>
    <div className="movement-summary"><span>{filtered.length} 筆</span><b className="green-text">入 +{formatNumber(inQty)}</b><b className="red-text">出 -{formatNumber(outQty)}</b><b className="yellow-text">採購 ${formatMoney(purchased)}</b></div>
    {loading ? <PageSkeleton /> : <MovementTable movements={filtered} products={products} readOnly={readOnly} onDelete={onDelete} />}
  </section>
}

function MovementTable({ movements, products, readOnly, onDelete }: { movements: Movement[]; products: Product[]; readOnly: boolean; onDelete: (movement: Movement) => void }) {
  return <div className="table-card"><table className={`data-table movement-table ${readOnly ? 'guest-read-only-table' : ''}`}><thead><tr><th>類型</th><th>日期</th><th>編號</th><th>品名</th><th className="number">數量</th><th className="number">單價</th><th className="number">金額</th><th>供應商</th><th>備註</th>{!readOnly && <th aria-label="操作" />}</tr></thead><tbody>
    {movements.length === 0 ? <EmptyRow colSpan={readOnly ? 9 : 10} label="沒有符合條件的出入庫紀錄" /> : movements.map((m) => {
      const p = m.product || products.find((item) => item.id === m.product_id)
      const initialStock = isInitialStockNote(m.notes)
      return <tr key={m.id}><td><span className={`tag ${m.movement_type}`}>{m.movement_type === 'in' ? '入庫' : '出庫'}</span></td><td className="muted nowrap">{formatDate(m.movement_date)}</td><td className="code">{p?.code}</td><td className="product-name">{p?.name}</td><td className={`number ${m.movement_type === 'in' ? 'green-text' : 'red-text'}`}>{m.movement_type === 'in' ? '+' : '-'}{formatNumber(m.quantity)} <small>{p?.unit}</small></td><td className="number">{m.unit_price ? formatMoney(m.unit_price) : '—'}</td><td className="number muted">{m.unit_price ? formatMoney(m.quantity * m.unit_price) : '—'}</td><td>{m.supplier || '—'}</td><td className="muted note-cell">{initialStock ? INITIAL_STOCK_NOTE : m.notes || '—'}{initialStock && <small className="movement-record-time">{formatTaipeiRecordTime(m.created_at, true)}</small>}</td>{!readOnly && <td><div className="row-actions"><button className="danger-icon" onClick={() => onDelete(m)} aria-label={`刪除 ${formatDate(m.movement_date)} ${p?.name || ''} ${m.movement_type === 'in' ? '入庫' : '出庫'}紀錄`} title="刪除紀錄並重算庫存"><Trash2 /></button></div></td>}</tr>
    })}
  </tbody></table></div>
}

const priceEntriesFor = (productId: string, movements: Movement[]) => movements
  .filter((movement) => movement.product_id === productId && movement.movement_type === 'in' && movement.unit_price !== null)
  .sort((a, b) => `${b.movement_date}-${b.created_at}`.localeCompare(`${a.movement_date}-${a.created_at}`))

function StocktakePage({ products, loading, add, openScanner, focusProductId, clearFocusProduct, onAdjust }: { products: Product[]; loading: boolean; add: () => void; openScanner: () => void; focusProductId: string; clearFocusProduct: () => void; onAdjust: (product: Product, actualQuantity: number, notes: string) => Promise<boolean> }) {
  const [query, setQuery] = useState('')
  const [category, setCategory] = useState('全部分類')
  const [drafts, setDrafts] = useState<Record<string, string>>({})
  const [notes, setNotes] = useState<Record<string, string>>({})
  const [savingId, setSavingId] = useState('')
  const [pageError, setPageError] = useState('')
  const categories = ['全部分類', ...PRODUCT_CATEGORIES]
  useEffect(() => {
    const product = products.find((item) => item.id === focusProductId)
    if (product) {
      setQuery(product.code)
      setCategory('全部分類')
      clearFocusProduct()
    }
  }, [focusProductId, products])
  const normalizedQuery = query.toLowerCase()
  /* 「庫存盤點」的「表格資料順序」：依刀具最後更新時間由新到舊排列；時間相同時再依編號由新到舊。 */
  const filtered = products
    .filter((product) => {
      const haystack = `${product.code} ${product.barcode || ''} ${product.name} ${product.brand} ${product.supplier}`.toLowerCase()
      return (haystack.includes(normalizedQuery) || matchesProductName(product.name, normalizedQuery)) && (category === '全部分類' || product.category === category)
    })
    .sort((a, b) => b.updated_at.localeCompare(a.updated_at) || b.code.localeCompare(a.code))
  const actualFor = (product: Product) => drafts[product.id] ?? String(product.stock_quantity)
  const differenceFor = (product: Product) => {
    const actual = Number(actualFor(product))
    return Number.isInteger(actual) && actual >= 0 ? actual - product.stock_quantity : null
  }
  const pendingCount = filtered.filter((product) => differenceFor(product) !== null && differenceFor(product) !== 0).length
  const noPriceCount = filtered.filter((product) => product.unit_price <= 0).length

  const save = async (product: Product) => {
    setPageError('')
    const actual = Number(actualFor(product))
    if (!Number.isInteger(actual) || actual < 0) { setPageError(`${product.code} 的實際數量必須是 0 以上的整數`); return }
    if (actual === product.stock_quantity) return
    setSavingId(product.id)
    const saved = await onAdjust(product, actual, notes[product.id] || '')
    if (saved) {
      setDrafts((current) => ({ ...current, [product.id]: String(actual) }))
      setNotes((current) => ({ ...current, [product.id]: '' }))
    }
    setSavingId('')
  }

  return <section className="page-content stocktake-page">
    <PageTitle eyebrow="STOCKTAKE" title="庫存盤點" actions={<><ScanButton openScanner={openScanner}/><button className="button primary" onClick={add}><Plus />新增盤點品項</button></>} />
    <div className="stocktake-guide"><ClipboardCheck /><div><strong>直接填入現場實際數量</strong><p>不需要輸入價格。系統只以盤點差異調整庫存，既有單價會保留；沒有價格的刀具也能正常盤點。</p></div></div>
    <div className="stocktake-summary">
      <div><span>目前顯示</span><strong>{formatNumber(filtered.length)}</strong><small>個品項</small></div>
      <div><span>未設定單價</span><strong>{formatNumber(noPriceCount)}</strong><small>仍可盤點</small></div>
      <div className={pendingCount ? 'attention' : ''}><span>待儲存差異</span><strong>{formatNumber(pendingCount)}</strong><small>個品項</small></div>
    </div>
    <div className="filter-card product-filter"><SearchField value={query} setValue={setQuery} placeholder="掃描條碼或搜尋編號、品名、廠牌、供應商" scanProducts={products}/><label className="select-wrap"><select value={category} onChange={(event) => setCategory(event.target.value)}>{categories.map((value) => <option key={value}>{value}</option>)}</select><ChevronDown /></label></div>
    {pageError && <div className="form-error stocktake-error">{pageError}</div>}
    {loading ? <PageSkeleton /> : <div className="table-card"><table className="data-table stocktake-table"><thead><tr><th>編號</th><th>品名</th><th>分類</th><th>單價</th><th className="number">庫存</th><th className="number">實際數量</th><th className="number">差異</th><th aria-label="更新" /><th>盤點備註</th></tr></thead><tbody>
      {filtered.length === 0 ? <EmptyRow colSpan={9} label="沒有符合條件的刀具"/> : filtered.map((product) => {
        const difference = differenceFor(product)
        const invalid = difference === null
        return <tr key={product.id} className={difference ? 'has-stocktake-difference' : ''}>
          <td className="code">{product.code}</td>
          <td className="product-name">{product.name}</td>
          <td><span className="tag neutral">{product.category}</span></td>
          <td className="number">{product.unit_price > 0 ? <span className="known-price">{formatMoney(product.unit_price)}</span> : <span className="unknown-price">未設定</span>}</td>
          <td className="number quantity">{formatNumber(product.stock_quantity)} <small>{product.unit}</small></td>
          <td><label className={`stocktake-quantity ${invalid ? 'invalid' : ''}`}><input type="number" inputMode="numeric" pattern="[0-9]*" min="0" step="1" value={actualFor(product)} onFocus={(event) => event.currentTarget.select()} onChange={(event) => { setDrafts((current) => ({ ...current, [product.id]: event.target.value })); setPageError('') }} aria-label={`${product.name} 實際盤點數量`}/><span>{product.unit}</span></label></td>
          <td className={`number stocktake-difference ${difference === null ? 'invalid' : difference > 0 ? 'positive' : difference < 0 ? 'negative' : ''}`}>{difference === null ? '格式錯誤' : difference === 0 ? '0' : `${difference > 0 ? '+' : ''}${formatNumber(difference)}`}</td>
          <td><button className="stocktake-save" onClick={() => void save(product)} disabled={savingId !== '' || invalid || difference === 0}><Save />{savingId === product.id ? '儲存中' : '更新'}</button></td>
          <td><input className="stocktake-note" value={notes[product.id] || ''} onChange={(event) => setNotes((current) => ({ ...current, [product.id]: event.target.value }))} placeholder="選填" /></td>
        </tr>
      })}
    </tbody></table></div>}
  </section>
}

function ProductsPage({ products, movements, loading, add, edit, openStock, openScanner, viewHistory, readOnly, onDelete }: { products: Product[]; movements: Movement[]; loading: boolean; add: () => void; edit: (p: Product) => void; openStock: (type: MovementType, productId: string) => void; openScanner: () => void; viewHistory: (p: Product) => void; readOnly: boolean; onDelete: (p: Product) => void }) {
  const [query, setQuery] = useState('')
  const [category, setCategory] = useState('全部分類')
  const categories = ['全部分類', ...PRODUCT_CATEGORIES]
  const normalizedQuery = query.toLowerCase()
  const filtered = products
    .filter((product) => {
      const haystack = `${product.code} ${product.barcode || ''} ${product.name} ${product.brand} ${product.supplier}`.toLowerCase()
      return (haystack.includes(normalizedQuery) || matchesProductName(product.name, normalizedQuery)) && (category === '全部分類' || product.category === category)
    })
    .sort((a, b) => `${b.purchased_at || '0000-00-00'}-${b.updated_at}-${b.code}`.localeCompare(`${a.purchased_at || '0000-00-00'}-${a.updated_at}-${a.code}`))
  return <section className="page-content">
    <PageTitle eyebrow="TOOLS" title="刀具資料" actions={!readOnly ? <><ScanButton openScanner={openScanner}/><button className="button primary" onClick={add}><Plus />新增刀具</button></> : undefined} />
    <div className="filter-card product-filter"><SearchField value={query} setValue={setQuery} placeholder="掃描條碼或搜尋編號、品名、廠牌、供應商" scanProducts={products}/><label className="select-wrap"><select value={category} onChange={(e) => setCategory(e.target.value)}>{categories.map((c) => <option key={c}>{c}</option>)}</select><ChevronDown /></label></div>
    {loading ? <PageSkeleton /> : <div className="table-card"><table className="data-table product-table"><thead><tr><th>編號</th><th>品名</th><th>廠牌</th><th>供應商</th><th>分類</th><th className="number">數量</th><th className="number">單價</th><th className="number" aria-label="漲幅" /><th>購入日期</th><th aria-label="操作" /></tr></thead><tbody>
      {filtered.length === 0 ? <EmptyRow colSpan={10} label="沒有符合條件的刀具"/> : filtered.map((p) => {
        const prices = priceEntriesFor(p.id, movements)
        const latest = prices[0]?.unit_price || p.unit_price
        const previous = prices[1]?.unit_price
        const percent = previous ? Math.round(((latest - previous) / previous) * 100) : null
        const percentClass = percent === null || percent === 0 ? 'price-flat' : percent > 0 ? 'price-up' : 'price-down'
        const percentLabel = percent === null ? '—' : percent === 0 ? '0%' : `${percent > 0 ? '▲' : '▼'}${Math.abs(percent)}%`
        const purchaseRecord = prices.find((entry) => entry.movement_date === p.purchased_at)
        const recordedAt = purchaseRecord?.created_at || p.created_at
        return <tr key={p.id}><td className="code">{p.code}</td><td className="product-name">{p.name}{p.barcode && <small className="product-barcode"><ScanBarcode />{p.barcode}</small>}</td><td><strong>{p.brand || '—'}</strong></td><td className="muted"><strong>{p.supplier || '—'}</strong></td><td><span className="tag neutral">{p.category}</span></td><td className="number quantity"><strong>{formatNumber(p.stock_quantity)}</strong> <small>{p.unit}</small></td><td className="number price-cell"><span>{formatMoney(p.unit_price)}</span><span className={`mobile-price-change ${percentClass}`}>{percentLabel}</span></td><td className="number price-change-cell"><span className={percentClass}>{percentLabel}</span></td><td className="muted nowrap"><span className="product-purchase-datetime"><span>{p.purchased_at ? formatDate(p.purchased_at) : '—'}</span><span>{formatTaipeiRecordTime(recordedAt, false, true)}</span></span></td><td><div className={`row-actions product-actions ${readOnly ? 'guest-view-actions' : ''}`}>{!readOnly && <><button className="in-icon" onClick={() => openStock('in', p.id)} aria-label={`${p.name} 入庫`} title="入庫"><ArrowDownToLine /></button><button className="out-icon" onClick={() => openStock('out', p.id)} aria-label={`${p.name} 出庫`} title="出庫"><ArrowUpFromLine /></button></>}<button className="history-icon" onClick={() => viewHistory(p)} aria-label={`${p.name} 歷史價格`} title="歷史價格"><History /></button>{!readOnly && <><button onClick={() => edit(p)} aria-label={`編輯 ${p.name}`} title="編輯"><Pencil /></button><button className="danger-icon" onClick={() => onDelete(p)} aria-label={`刪除 ${p.name}`} title="刪除"><Trash2 /></button></>}</div></td></tr>
      })}
    </tbody></table></div>}
  </section>
}

function SearchField({ value, setValue, placeholder, scanProducts = [] }: { value: string; setValue: (v: string) => void; placeholder: string; scanProducts?: Product[] }) {
  const inputRef = useRef<HTMLInputElement>(null)
  const clearTimerRef = useRef<number | null>(null)
  const [inputValue, setInputValue] = useState(value)
  const [scanResult, setScanResult] = useState<{ success: boolean; message: string } | null>(null)

  useEffect(() => { setInputValue(value) }, [value])
  useEffect(() => () => {
    if (clearTimerRef.current !== null) window.clearTimeout(clearTimerRef.current)
  }, [])

  /* 「出入庫、庫存盤點、刀具資料」的「條碼搜尋」：清空可輸入文字但保留篩選條件，下一筆掃描會直接取代上一筆。 */
  const finishBarcodeScan = (event: ReactKeyboardEvent<HTMLInputElement>) => {
    if (event.key !== 'Enter' || event.nativeEvent.isComposing || !scanProducts.length) return
    event.preventDefault()
    const barcode = inputValue.trim()
    if (!barcode) return
    const matchedProduct = scanProducts.find((product) => product.barcode?.trim().toLowerCase() === barcode.toLowerCase())
    setScanResult(matchedProduct
      ? { success: true, message: `掃碼成功：${matchedProduct.code} ${matchedProduct.name}` }
      : { success: false, message: `掃碼失敗：找不到條碼 ${barcode}` })
    if (clearTimerRef.current !== null) window.clearTimeout(clearTimerRef.current)
    clearTimerRef.current = window.setTimeout(() => {
      setInputValue('')
      setScanResult(null)
      inputRef.current?.focus()
      clearTimerRef.current = null
    }, 1800)
  }

  const clearSearch = () => {
    if (clearTimerRef.current !== null) window.clearTimeout(clearTimerRef.current)
    clearTimerRef.current = null
    setScanResult(null)
    setInputValue('')
    setValue('')
    inputRef.current?.focus()
  }

  const inputPlaceholder = !inputValue && value ? `目前保留 ${value} 的搜尋結果，可直接掃描下一筆` : placeholder
  return <label className="search-field"><Search /><input ref={inputRef} value={inputValue} onChange={(e) => { setInputValue(e.target.value); setValue(e.target.value); setScanResult(null) }} onKeyDown={finishBarcodeScan} placeholder={inputPlaceholder}/>{(inputValue || value) && <button type="button" onClick={clearSearch} aria-label="清除搜尋文字與目前結果"><X /></button>}{scanResult && <span className={`search-scan-status ${scanResult.success ? 'success' : 'failure'}`} role="status">{scanResult.success ? <ShieldCheck/> : <Triangle/>}<strong>{scanResult.message}</strong><small>輸入文字即將清空，搜尋結果會保留，可繼續掃描下一筆</small></span>}</label>
}

function EmptyRow({ colSpan, label }: { colSpan: number; label: string }) { return <tr><td colSpan={colSpan}><div className="empty-state"><ClipboardList />{label}</div></td></tr> }

function StockModal({ type, initialProductId, products, close, switchType, save }: { type: MovementType; initialProductId?: string; products: Product[]; close: () => void; switchType: (t: MovementType) => void; save: (input: MovementInput) => Promise<void> }) {
  const [productId, setProductId] = useState(initialProductId || '')
  const [quantity, setQuantity] = useState('')
  const [date, setDate] = useState(today())
  const [price, setPrice] = useState('')
  const [supplier, setSupplier] = useState('')
  const [notes, setNotes] = useState('')
  const [busy, setBusy] = useState(false)
  const [formError, setFormError] = useState('')
  const selected = products.find((p) => p.id === productId)
  const supplierOptions = Array.from(new Set(products.map((product) => product.supplier.trim()).filter(Boolean))).sort((a, b) => a.localeCompare(b, 'zh-Hant'))
  const totalPrice = Number(quantity || 0) * Number(price || 0)
  useEffect(() => { if (selected && type === 'in') { setPrice(String(selected.unit_price || '')); setSupplier(selected.supplier || '') } }, [productId, type])
  const submit = async (e: FormEvent) => {
    e.preventDefault(); setFormError('')
    if (!productId || Number(quantity) <= 0) { setFormError('請選擇商品並輸入大於 0 的數量'); return }
    if (type === 'in' && Number(price) < 0) { setFormError('單價不能小於 0'); return }
    setBusy(true)
    try { await save({ product_id: productId, movement_type: type, quantity: Number(quantity), movement_date: date, unit_price: Number(price) || 0, supplier, notes, request_key: crypto.randomUUID() }) }
    catch (err) { setFormError(messageFor(err)); setBusy(false) }
  }
  return <Modal close={close} title={`登記${type === 'in' ? '入庫' : '出庫'}`}>
    <form className="stock-movement-form" onSubmit={submit} onKeyDown={preventInputEnterSubmit}>
      <div className="segmented"><button type="button" className={type === 'in' ? 'selected-in' : ''} onClick={() => switchType('in')}>入庫</button><button type="button" className={type === 'out' ? 'selected-out' : ''} onClick={() => switchType('out')}>出庫</button></div>
      <Field label="刀具" wide><ProductPicker products={products} value={productId} onChange={setProductId}/></Field>
      {selected && <div className="stock-hint">目前庫存 <strong>{formatNumber(selected.stock_quantity)} {selected.unit}</strong></div>}
      <div className="form-grid"><Field label="數量"><input required min="1" step="1" type="number" inputMode="numeric" pattern="[0-9]*" value={quantity} onChange={(e) => setQuantity(e.target.value)}/></Field><Field label={type === 'in' ? '購入日期' : '出庫日期'}><input required type="date" value={date} onChange={(e) => setDate(e.target.value)}/></Field>
      {type === 'in' && <><Field label="單價"><input required min="0" step="0.01" type="number" inputMode="numeric" pattern="[0-9]*" value={price} onChange={(e) => setPrice(e.target.value)}/></Field><Field label="供應商"><EditableCombo options={supplierOptions} value={supplier} onChange={setSupplier} placeholder="選擇或輸入新供應商"/></Field></>}</div>
      {type === 'in' && <TotalPreview quantity={Number(quantity || 0)} unitPrice={Number(price || 0)} total={totalPrice}/>} 
      <Field label="備註" wide><textarea value={notes} onChange={(e) => setNotes(e.target.value)} placeholder={type === 'in' ? '選填' : '選填'}/></Field>
      {formError && <p className="form-error">{formError}</p>}
      <button disabled={busy || products.length === 0} className={`button modal-submit ${type === 'in' ? 'inbound' : 'outbound'}`}>{busy ? '處理中…' : `確認${type === 'in' ? '入庫' : '出庫'}`}</button>
    </form>
  </Modal>
}

function ProductPicker({ products, value, onChange, inline = false, disabled = false }: { products: Product[]; value: string; onChange: (value: string) => void; inline?: boolean; disabled?: boolean }) {
  const selected = products.find((product) => product.id === value)
  const display = (product: Product) => `${product.code}｜${product.name}`
  const [query, setQuery] = useState(selected ? display(selected) : '')
  const [open, setOpen] = useState(inline && !selected)
  const inputRef = useRef<HTMLInputElement>(null)
  useEffect(() => { const product = products.find((item) => item.id === value); if (product) setQuery(display(product)) }, [value, products])
  const normalized = query.trim().toLowerCase()
  const filtered = products.filter((product) => {
    const haystack = `${product.code} ${product.name} ${product.brand} ${product.supplier} ${product.category}`.toLowerCase()
    return haystack.includes(normalized) || matchesProductName(product.name, normalized)
  }).slice(0, 10)
  const choose = (product: Product) => { onChange(product.id); setQuery(display(product)); setOpen(false); inputRef.current?.blur() }
  const openPicker = () => {
    if (inline && selected && !open) setQuery('')
    setOpen(true)
  }
  return <div className={`product-picker${inline ? ' product-picker-inline' : ''}`} onBlur={inline ? undefined : () => window.setTimeout(() => setOpen(false), 120)}>
    <div className="product-picker-input">
      <Search/>
      <input ref={inputRef} id={inline ? 'barcode-binding-product' : undefined} required disabled={disabled} role="combobox" aria-expanded={open} aria-controls="product-picker-options" value={query} onFocus={openPicker} onClick={openPicker} onChange={(event) => { setQuery(event.target.value); onChange(''); setOpen(true) }} placeholder="輸入編號、品名、廠牌或供應商搜尋"/>
      {inline ? <button className="product-picker-toggle" type="button" disabled={disabled} aria-expanded={open} aria-controls="product-picker-options" aria-label={open ? '收起刀具清單' : '展開刀具清單'} onClick={() => {
        if (open) { setOpen(false); inputRef.current?.blur() }
        else openPicker()
      }}>{open ? <ChevronUp/> : <ChevronDown/>}</button> : <ChevronDown/>}
    </div>
    {inline && !open && selected && <p className="barcode-binding-selection"><strong>{selected.code}</strong><span>{selected.name}</span></p>}
    {open && <div className="product-picker-menu" id="product-picker-options" role="listbox">
      {filtered.length === 0 ? <div className="product-picker-empty">找不到符合的刀具</div> : filtered.map((product) => <button type="button" role="option" disabled={disabled} aria-selected={product.id === value} key={product.id} onMouseDown={inline ? undefined : (event) => event.preventDefault()} onClick={() => choose(product)}><span><b>{product.code}</b>{product.name}</span><small>庫存 {formatNumber(product.stock_quantity)} {product.unit}</small></button>)}
    </div>}
  </div>
}

function TotalPreview({ quantity, unitPrice, total }: { quantity: number; unitPrice: number; total: number }) {
  return <div className="total-preview"><span>採購總價</span><strong>${formatMoney(total)}</strong><small>{formatNumber(quantity)} × ${formatMoney(unitPrice)}</small></div>
}

// 設定碼取自 XB-6208RB 說明書。頁面一次只顯示一張，避免掃到相鄰選項。
const SCANNER_SETTINGS = [
  { title: '藍牙配對操作', options: [
    { id: 'pair-1', label: '配對碼 I' },
    { id: 'pair-2', label: '配對碼 II' },
  ] },
  { title: '休眠時間設定', options: [
    { id: 'sleep-1', label: '1 分鐘' },
    { id: 'sleep-5', label: '5 分鐘' },
    { id: 'sleep-10', label: '10 分鐘' },
    { id: 'sleep-off', label: '不休眠' },
  ] },
  { title: '聲音設定', options: [
    { id: 'sound-on', label: '開啟聲音' },
    { id: 'sound-off', label: '關閉聲音' },
    { id: 'sound-high', label: '高音' },
    { id: 'sound-medium', label: '中音' },
    { id: 'sound-low', label: '低音' },
  ] },
]

function ScannerSettingsPage() {
  const [selection, setSelection] = useState({ groupIndex: 0, optionIndex: 0 })
  const group = SCANNER_SETTINGS[selection.groupIndex]
  const option = group.options[selection.optionIndex]

  return <section className="page-content">
    <PageTitle eyebrow="XB-6208-RB" title="掃描槍設定" />
    <div className="scanner-settings-card">
      <div>
        <div className="scanner-settings-buttons" role="group" aria-label="設定類別">
          {SCANNER_SETTINGS.map(({ title }, groupIndex) => <button key={title} type="button" className="button" aria-pressed={selection.groupIndex === groupIndex} onClick={() => setSelection({ groupIndex, optionIndex: 0 })}>{title}</button>)}
        </div>
        {selection.groupIndex === 0 && <ol className="scanner-pairing-steps">
          <li>用掃描槍依序掃描「配對碼 I」和「配對碼 II」，直到藍燈閃爍。</li>
          <li>手機或平板請到藍牙設定選擇「R&amp;B40」。若使用 USB 接收器，此時插上接收器。藍燈恆亮表示配對完成。</li>
        </ol>}
        <div className="scanner-settings-buttons" role="group" aria-label={group.title}>
          {group.options.map(({ id, label }, optionIndex) => <button key={id} type="button" className="button" aria-pressed={selection.optionIndex === optionIndex} onClick={() => setSelection({ ...selection, optionIndex })}>{label}</button>)}
        </div>
      </div>
      <figure className="scanner-settings-code">
        <img src={`${import.meta.env.BASE_URL}scanner-settings/${option.id}.png`} width={256} height={256} alt={`${option.label} QR Code`} />
        <figcaption aria-live="polite">{option.label}</figcaption>
        <small>請用 XB-6208-RB 掃描畫面上的設定碼</small>
      </figure>
    </div>
  </section>
}

function BarcodeScannerModal({ products, initialBarcode, close, createWithBarcode, linkBarcode, openStock, openStocktake, editProduct }: {
  products: Product[]
  initialBarcode: string
  close: () => void
  createWithBarcode: (barcode: string) => void
  linkBarcode: (productId: string, barcode: string) => Promise<void>
  openStock: (type: MovementType, productId: string) => void
  openStocktake: (productId: string) => void
  editProduct: (product: Product) => void
}) {
  const videoRef = useRef<HTMLVideoElement>(null)
  const targetRef = useRef<HTMLSpanElement>(null)
  const resultRef = useRef<HTMLElement>(null)
  const controlsRef = useRef<{ stop: () => void } | null>(null)
  const audioContextRef = useRef<AudioContext | null>(null)
  const [cameraActive, setCameraActive] = useState(false)
  const [barcode, setBarcode] = useState(initialBarcode)
  const [manualBarcode, setManualBarcode] = useState('')
  const [selectedProductId, setSelectedProductId] = useState('')
  const [error, setError] = useState('')
  const [linking, setLinking] = useState(false)
  const [bindingOpen, setBindingOpen] = useState(false)
  const bindingRef = useRef<HTMLElement>(null)
  const [keyboardQrOpen, setKeyboardQrOpen] = useState(false)
  const normalizedBarcode = barcode.trim().toLowerCase()
  const matchedProduct = products.find((product) => product.barcode?.trim().toLowerCase() === normalizedBarcode)
    || products.find((product) => product.code.toLowerCase() === normalizedBarcode)

  // 「掃描刀具條碼」的「掃描成功提示音」：相機按鈕或查詢按鈕觸發後才啟用，避免 iPhone 阻擋自動播放。
  const prepareScanAudio = async () => {
    try {
      const AudioContextConstructor = window.AudioContext
        || (window as typeof window & { webkitAudioContext?: typeof AudioContext }).webkitAudioContext
      if (!AudioContextConstructor) return
      if (!audioContextRef.current) audioContextRef.current = new AudioContextConstructor()
      if (audioContextRef.current.state === 'suspended') await audioContextRef.current.resume()
    } catch { /* 瀏覽器不支援提示音時仍保留視覺提示。 */ }
  }

  const playScanFeedback = () => {
    const audioContext = audioContextRef.current
    if (audioContext?.state === 'running') {
      const startAt = audioContext.currentTime
      ;[880, 1175].forEach((frequency, index) => {
        const oscillator = audioContext.createOscillator()
        const gain = audioContext.createGain()
        const toneStart = startAt + index * .09
        oscillator.frequency.value = frequency
        oscillator.type = 'sine'
        gain.gain.setValueAtTime(.0001, toneStart)
        gain.gain.exponentialRampToValueAtTime(.12, toneStart + .01)
        gain.gain.exponentialRampToValueAtTime(.0001, toneStart + .08)
        oscillator.connect(gain).connect(audioContext.destination)
        oscillator.start(toneStart)
        oscillator.stop(toneStart + .085)
      })
    }
    navigator.vibrate?.(70)
  }

  // 「掃描刀具條碼」的「iPhone 後鏡頭連續對焦與適度變焦參數」：裝置不支援時會略過，不影響相機開啟。
  const optimizeBarcodeCamera = async () => {
    const stream = videoRef.current?.srcObject
    if (!(stream instanceof MediaStream)) return
    const track = stream.getVideoTracks()[0]
    if (!track?.getCapabilities) return
    const capabilities = track.getCapabilities() as BarcodeCameraCapabilities
    const applyOptionalConstraint = async (constraint: BarcodeCameraConstraintSet) => {
      try { await track.applyConstraints({ advanced: [constraint] }) }
      catch { /* iPhone／瀏覽器未提供這項相機能力時維持預設設定。 */ }
    }
    if (capabilities.focusMode?.includes('continuous')) await applyOptionalConstraint({ focusMode: 'continuous' })
    if (capabilities.zoom && Number.isFinite(capabilities.zoom.min) && Number.isFinite(capabilities.zoom.max)) {
      const preferredZoom = Math.min(capabilities.zoom.max, Math.max(capabilities.zoom.min, 1.35))
      await applyOptionalConstraint({ zoom: preferredZoom })
    }
  }

  const stopCamera = () => {
    controlsRef.current?.stop()
    controlsRef.current = null
    if (videoRef.current?.srcObject instanceof MediaStream) videoRef.current.srcObject.getTracks().forEach((track) => track.stop())
    if (videoRef.current) videoRef.current.srcObject = null
    setCameraActive(false)
  }

  useEffect(() => () => {
    controlsRef.current?.stop()
    controlsRef.current = null
    if (videoRef.current?.srcObject instanceof MediaStream) videoRef.current.srcObject.getTracks().forEach((track) => track.stop())
    if (audioContextRef.current) void audioContextRef.current.close().catch(() => undefined)
  }, [])

  useEffect(() => {
    if (bindingOpen) {
      bindingRef.current?.closest('.modal')?.scrollTo(0, 0)
      return
    }
    if (!barcode) return
    const frame = window.requestAnimationFrame(() => resultRef.current?.scrollIntoView({ behavior: 'smooth', block: 'nearest' }))
    return () => window.cancelAnimationFrame(frame)
  }, [barcode, bindingOpen])

  const startCamera = async () => {
    setError(''); setBarcode(''); setSelectedProductId('')
    if (!window.isSecureContext || !navigator.mediaDevices?.getUserMedia) {
      setError('此瀏覽器無法開啟相機。請使用正式 HTTPS 網址，或先用下方手動輸入測試。')
      return
    }
    stopCamera()
    let stream: MediaStream | undefined
    let timer: number | undefined
    const controls = {
      stop: () => {
        window.clearTimeout(timer)
        stream?.getTracks().forEach((track) => track.stop())
      },
    }
    controlsRef.current = controls
    try {
      await prepareScanAudio()
      const [{ BrowserMultiFormatReader }, { DecodeHintType }] = await Promise.all([
        import('@zxing/browser'),
        import('@zxing/library'),
      ])
      const video = videoRef.current
      if (!video || controlsRef.current !== controls) return
      // 刀具掃碼視窗｜更密集地搜尋 QR 定位圖案，並縮短失敗後等待下一次取像的時間。
      const hints = new Map([[DecodeHintType.TRY_HARDER, true]])
      const reader = new BrowserMultiFormatReader(hints)
      const fullCanvas = document.createElement('canvas')
      const fullContext = fullCanvas.getContext('2d', { willReadFrequently: true })
      if (!fullContext) throw new Error('無法讀取相機影像')
      const qrCanvas = document.createElement('canvas')
      const qrContext = qrCanvas.getContext('2d', { willReadFrequently: true })
      // 刀具掃碼視窗｜將可見黃框換算成原始影像座標，包含 object-fit: cover 裁掉的畫面偏移。
      const readTargetQr = async () => {
        const target = targetRef.current
        const { videoWidth, videoHeight } = video
        if (!target || !qrContext || video.readyState < HTMLMediaElement.HAVE_CURRENT_DATA || !videoWidth || !videoHeight) return ''
        const preview = video.getBoundingClientRect()
        if (!preview.width || !preview.height) return ''
        const frame = target.getBoundingClientRect()
        const scale = Math.max(preview.width / videoWidth, preview.height / videoHeight)
        const imageLeft = preview.left + (preview.width - videoWidth * scale) / 2
        const imageTop = preview.top + (preview.height - videoHeight * scale) / 2
        const left = Math.max(0, Math.floor((frame.left - imageLeft) / scale))
        const top = Math.max(0, Math.floor((frame.top - imageTop) / scale))
        const sourceWidth = Math.min(videoWidth, Math.ceil((frame.right - imageLeft) / scale)) - left
        const sourceHeight = Math.min(videoHeight, Math.ceil((frame.bottom - imageTop) / scale)) - top
        if (sourceWidth < 32 || sourceHeight < 32) return ''

        const outputScale = Math.min(1, 960 / Math.max(sourceWidth, sourceHeight))
        const outputWidth = Math.round(sourceWidth * outputScale)
        const outputHeight = Math.round(sourceHeight * outputScale)
        const padding = Math.ceil(Math.min(outputWidth, outputHeight) / 16)
        const canvasWidth = outputWidth + padding * 2
        const canvasHeight = outputHeight + padding * 2
        if (qrCanvas.width !== canvasWidth || qrCanvas.height !== canvasHeight) {
          qrCanvas.width = canvasWidth
          qrCanvas.height = canvasHeight
        }
        try {
          qrContext.fillStyle = '#fff'
          qrContext.fillRect(0, 0, canvasWidth, canvasHeight)
          qrContext.drawImage(video, left, top, sourceWidth, sourceHeight, padding, padding, outputWidth, outputHeight)
          const image = qrContext.getImageData(0, 0, canvasWidth, canvasHeight)
          const { scanQr } = await import('./lib/scanQr')
          if (controlsRef.current !== controls) return ''
          return await scanQr(image)
        } catch { return '' }
      }
      stream = await navigator.mediaDevices.getUserMedia({
        audio: false,
        video: {
          facingMode: { ideal: 'environment' },
          width: { ideal: 1920 },
          height: { ideal: 1080 },
          frameRate: { ideal: 30, max: 30 },
          aspectRatio: { ideal: 16 / 9 },
        },
      })
      if (videoRef.current !== video || controlsRef.current !== controls) {
        controls.stop()
        return
      }
      video.srcObject = stream
      await video.play()
      if (controlsRef.current !== controls) return
      await optimizeBarcodeCamera()
      if (controlsRef.current !== controls) return
      setCameraActive(true)

      // 刀具掃碼視窗｜每次依目前影像尺寸取像，單一解碼器失敗時仍繼續掃描。
      const scanFrame = async () => {
        if (controlsRef.current !== controls) return
        const { videoWidth, videoHeight } = video
        let value = ''
        if (video.readyState >= HTMLMediaElement.HAVE_CURRENT_DATA && videoWidth && videoHeight) {
          if (fullCanvas.width !== videoWidth || fullCanvas.height !== videoHeight) {
            fullCanvas.width = videoWidth
            fullCanvas.height = videoHeight
          }
          try {
            fullContext.drawImage(video, 0, 0)
            value = reader.decodeFromCanvas(fullCanvas).getText().trim()
          } catch { /* 全畫面未解碼，改試黃框內的 QR Code。 */ }
          value ||= await readTargetQr()
        }
        if (controlsRef.current !== controls) return
        if (!value) {
          timer = window.setTimeout(scanFrame, 250)
          return
        }
        stopCamera()
        setBarcode(value)
        setManualBarcode('')
        playScanFeedback()
      }
      timer = window.setTimeout(scanFrame, 250)
    } catch (cameraError) {
      if (controlsRef.current !== controls) return
      stopCamera()
      const name = (cameraError as { name?: string })?.name
      setError(name === 'NotAllowedError'
        ? '相機權限被拒絕。請在 iPhone 的網站設定中允許相機後再試一次。'
        : name === 'NotFoundError'
          ? '找不到可使用的相機。'
          : name === 'NotReadableError'
            ? '相機目前被其他程式占用，請關閉其他相機程式後再試。'
            : `無法啟動條碼掃描：${messageFor(cameraError)}`)
    }
  }

  const useManualBarcode = async () => {
    const value = manualBarcode.trim()
    if (!value) { setError('請輸入條碼內容'); return }
    await prepareScanAudio()
    stopCamera(); setError(''); setBarcode(value); setManualBarcode(''); setSelectedProductId('')
    playScanFeedback()
  }

  const bindExistingProduct = async () => {
    if (!selectedProductId || !barcode) { setError('請先選擇要綁定的刀具'); return }
    setLinking(true); setError('')
    try {
      await linkBarcode(selectedProductId, barcode)
      setBindingOpen(false)
    }
    catch (linkError) { setError(messageFor(linkError)) }
    finally { setLinking(false) }
  }

  const returnToScan = () => {
    if (linking) return
    setError('')
    setBindingOpen(false)
  }

  return <Modal title={bindingOpen ? '綁定現有刀具' : '掃描刀具條碼'} close={bindingOpen ? returnToScan : () => { stopCamera(); close() }} size="barcode-scanner-modal">
    {bindingOpen ? <section ref={bindingRef} className="barcode-binding-step">
      <p className="barcode-binding-code">待綁定條碼 <strong>{barcode}</strong></p>
      <div className="field wide">
        <span><label htmlFor="barcode-binding-product">搜尋刀具</label></span>
        <ProductPicker products={products} value={selectedProductId} onChange={setSelectedProductId} inline disabled={linking} />
      </div>
      {error && <p className="form-error barcode-scanner-error" role="alert">{error}</p>}
      <div className="barcode-binding-actions">
        <button className="button" type="button" disabled={linking} onClick={returnToScan}>返回掃描結果</button>
        <button className="button scan-button" type="button" disabled={!selectedProductId || linking} onClick={() => void bindExistingProduct()}><Save/>{linking ? '綁定中…' : '確認綁定'}</button>
      </div>
    </section> : <>
    {/* <p className="barcode-scanner-intro">將條碼放在框線中央，建議先保持約 20～30 公分再慢慢靠近。第一次使用時，請允許網站使用 iPhone 相機。</p> */}
    <button className="ios-keyboard-toggle-inline" type="button" onClick={() => setKeyboardQrOpen((current) => !current)} aria-expanded={keyboardQrOpen}>
      <Keyboard/><span>XB-6208-RB：顯示／隱藏 iOS 鍵盤</span>{keyboardQrOpen ? <ChevronUp/> : <ChevronDown/>}
    </button>
    {keyboardQrOpen && <section className="ios-keyboard-toggle-panel">
      <p className="ios-keyboard-toggle-intro">使用 XB-6208-RB 掃描下方設定碼，即可切換 iPhone 鍵盤。</p>
      <figure className="ios-keyboard-toggle-figure">
        <img src={`${import.meta.env.BASE_URL}ios-keyboard-toggle-qr.png`} alt="XB-6208-RB 的 iOS 鍵盤顯示與隱藏設定 QR Code"/>
        <figcaption>鍵盤隱藏時掃描會顯示；鍵盤顯示時再次掃描會隱藏。</figcaption>
      </figure>
      <p className="ios-keyboard-toggle-help"><ShieldCheck/>這只是掃描器設定碼，不會新增、修改或刪除網站資料。掃描後若鍵盤沒有立刻出現，請重新點一下輸入框。</p>
    </section>}
    <div className={`barcode-camera ${cameraActive ? 'is-active' : ''} ${barcode && !cameraActive ? matchedProduct ? 'is-success' : 'is-unmatched' : ''}`}>
      <video ref={videoRef} muted playsInline aria-label="條碼掃描相機畫面" />
      {cameraActive && <><span ref={targetRef} className="barcode-target" aria-hidden="true"/><span className="barcode-scanning-label"><ScanBarcode/>掃描中</span></>}
      {!cameraActive && barcode && <div className={`barcode-scan-feedback ${matchedProduct ? 'success' : 'unmatched'}`}>
        {matchedProduct ? <ShieldCheck/> : <Triangle/>}
        <strong>{matchedProduct ? '掃描成功' : '已掃到條碼'}</strong>
        <span>{barcode}</span>
        <small>{matchedProduct ? `${matchedProduct.code}・${matchedProduct.name}` : '此條碼尚未綁定刀具'}</small>
      </div>}
      {!cameraActive && !barcode && <div className="barcode-camera-placeholder"><ScanBarcode/><span>點擊下方按鈕開始掃描</span></div>}
    </div>
    <button className={`button barcode-camera-button ${cameraActive ? 'outbound' : 'scan-button'}`} type="button" onClick={() => cameraActive ? stopCamera() : void startCamera()}>{cameraActive ? <X/> : barcode ? <History/> : <Camera/>}{cameraActive ? '停止相機' : barcode ? '掃描下一筆' : '開啟相機掃描'}</button>

    <div className="barcode-manual-test">
      <span>電腦測試／相機無法辨識時</span>
      <div><input value={manualBarcode} onChange={(event) => setManualBarcode(event.target.value)} onKeyDown={(event) => event.key === 'Enter' && (event.preventDefault(), void useManualBarcode())} placeholder="輸入條碼後按查詢" autoComplete="off"/><button type="button" onClick={() => void useManualBarcode()}><Search/>查詢</button></div>
    </div>

    {error && <p className="form-error barcode-scanner-error" role="alert">{error}</p>}
    {barcode && <section ref={resultRef} className={`barcode-result ${matchedProduct ? 'matched' : 'unmatched'}`}>
      <div className="barcode-result-code"><span>掃描結果</span><strong>{barcode}</strong></div>
      {matchedProduct ? <>
        <div className="barcode-product-card">
          <div><span>{matchedProduct.code}</span><strong>{matchedProduct.name}</strong><small>{matchedProduct.category}・{matchedProduct.brand || '未設定廠牌'}</small></div>
          <p>目前庫存 <b>{formatNumber(matchedProduct.stock_quantity)} {matchedProduct.unit}</b></p>
        </div>
        <div className="barcode-result-actions">
          <button className="button inbound" type="button" onClick={() => openStock('in', matchedProduct.id)}><ArrowDownToLine/>入庫</button>
          <button className="button outbound" type="button" onClick={() => openStock('out', matchedProduct.id)}><ArrowUpFromLine/>出庫</button>
          <button className="button stocktake-barcode-action" type="button" onClick={() => openStocktake(matchedProduct.id)}><ClipboardCheck/>盤點</button>
          <button className="button edit-barcode-action" type="button" onClick={() => editProduct(matchedProduct)}><Pencil/>編輯刀具</button>
        </div>
      </> : <>
        <div className="barcode-unmatched-message"><ScanBarcode/><div><strong>尚未綁定任何刀具</strong><span>可以建立新刀具，或從現有刀具中選擇一筆綁定。</span></div></div>
        <button className="button primary barcode-create-product" type="button" onClick={() => createWithBarcode(barcode)}><Plus/>以此條碼建立新刀具</button>
        <button className="button scan-button barcode-link-existing" type="button" onClick={() => { setError(''); setBindingOpen(true) }}>綁定現有刀具</button>
      </>}
    </section>}
    </>}
  </Modal>
}

function ProductModal({ product, products, initialBarcode = '', close, save }: { product: Product | null; products: Product[]; initialBarcode?: string; close: () => void; save: (input: ProductInput) => Promise<void> }) {
  const [data, setData] = useState<ProductInput>(() => product ? { barcode: product.barcode || '', specification: product.specification, size: product.size, chipbreaker: product.chipbreaker, material: product.material, brand: product.brand, supplier: product.supplier, category: product.category, unit: product.unit, notes: product.notes, initial_quantity: 0, initial_unit_price: 0, purchased_at: product.purchased_at || today() } : { ...blankProduct(), barcode: initialBarcode })
  const [busy, setBusy] = useState(false)
  const [formError, setFormError] = useState('')
  const [qrPreview, setQrPreview] = useState('')
  const [qrPreviewError, setQrPreviewError] = useState('')
  const set = <K extends keyof ProductInput>(key: K, value: ProductInput[K]) => setData((old) => ({ ...old, [key]: value }))
  const name = composeProductName(data.category, data.specification, data.size, data.chipbreaker, data.material)
  const initialTotalPrice = data.initial_quantity * data.initial_unit_price
  const categoryProducts = products.filter((item) => item.category === data.category)
  const uniqueValues = (items: Product[], key: 'specification' | 'size' | 'chipbreaker' | 'material' | 'brand' | 'supplier' | 'unit') => Array.from(new Set(items.map((item) => item[key].trim()).filter(Boolean))).sort((a, b) => a.localeCompare(b, 'zh-Hant', { numeric: true }))
  const specificationOptions = uniqueValues(categoryProducts, 'specification')
  const sizeOptions = uniqueValues(categoryProducts, 'size')
  const chipbreakerOptions = uniqueValues(categoryProducts, 'chipbreaker')
  const materialOptions = uniqueValues(categoryProducts, 'material')
  const brandOptions = uniqueValues(products, 'brand')
  const supplierOptions = uniqueValues(products, 'supplier')
  const unitOptions = uniqueValues(products, 'unit')
  useEffect(() => {
    let cancelled = false
    const barcode = data.barcode.trim()
    setQrPreview(''); setQrPreviewError('')
    if (!barcode) return () => { cancelled = true }
    void createQrCodeDataUrl(barcode).then((image) => {
      if (!cancelled) setQrPreview(image)
    }).catch(() => {
      if (!cancelled) setQrPreviewError('QR Code 產生失敗，請重新產生或重新輸入條碼')
    })
    return () => { cancelled = true }
  }, [data.barcode])
  const generateBarcode = () => {
    setFormError('')
    try { set('barcode', generateUniqueToolBarcode(products.map((item) => item.barcode))) }
    catch (error) { setFormError(messageFor(error)) }
  }
  const changeCategory = (category: ProductCategory) => setData((old) => ({
    ...old,
    category,
    size: category === '刀片' || category === '鑽銑刀' ? old.size : '',
    chipbreaker: category === '刀片' ? old.chipbreaker : '',
    material: category === '刀片' || category === '鑽銑刀' ? old.material : '',
    unit: category === '刀片' ? '片' : category === '其他' ? '件' : '支',
  }))
  const submit = async (e: FormEvent) => {
    e.preventDefault(); setFormError('')
    if (!data.specification.trim()) { setFormError('規格為必填欄位'); return }
    if (data.barcode.trim() && products.some((item) => item.id !== product?.id && item.barcode?.trim().toLowerCase() === data.barcode.trim().toLowerCase())) { setFormError('這個條碼已綁定其他刀具'); return }
    if ((data.category === '刀片' || data.category === '鑽銑刀') && !data.size.trim()) { setFormError('此分類需要輸入尺寸'); return }
    if ((data.category === '刀片' || data.category === '鑽銑刀') && !data.material.trim()) { setFormError('此分類需要輸入材質'); return }
    setBusy(true)
    try { await save(data) } catch (err) { setFormError(messageFor(err)); setBusy(false) }
  }
  return <Modal close={() => { if (!busy) close() }} title={product ? '編輯刀具' : '新增刀具'} size="large">
    <form onSubmit={submit} onKeyDown={preventInputEnterSubmit}>
      <div className="form-grid">
        <Field label="編號"><input disabled value={product?.code || '儲存後自動產生'} /></Field>
        <Field label="分類 *"><select required value={data.category} onChange={(e) => { changeCategory(e.target.value as ProductCategory); e.currentTarget.blur() }}>{PRODUCT_CATEGORIES.map((category) => <option key={category} value={category}>{category}</option>)}</select></Field>
      </div>
      <Field label="條碼（可掃描、手動輸入或自動產生）" wide>
        <div className="barcode-generator">
          <div className="barcode-generator-entry">
            <div className="barcode-input"><ScanBarcode/><input value={data.barcode} onChange={(event) => set('barcode', event.target.value.trimStart())} placeholder="尚未綁定條碼，可手動輸入" autoComplete="off" /></div>
            <button className="button barcode-generate-button" type="button" onClick={generateBarcode}><WandSparkles/>{data.barcode.trim() ? '重新產生條碼' : '自動產生條碼'}</button>
          </div>
          {data.barcode.trim() && <figure className="product-qr-preview">
            <div>{qrPreview ? <img src={qrPreview} alt={`條碼 ${data.barcode.trim()} 的 QR Code`}/> : <QrCode aria-hidden="true"/>}</div>
            <figcaption><strong>QR Code 預覽</strong><span>{data.barcode.trim()}</span><small>QR Code 與條碼號碼使用相同內容</small></figcaption>
          </figure>}
          {qrPreviewError && <p className="form-error barcode-qr-error">{qrPreviewError}</p>}
        </div>
      </Field>
      <Field label="品名（自動組合）" wide><input className="composed-name" disabled value={name || (data.category === '刀片' ? '規格-尺寸-斷屑槽-材質' : data.category === '鑽銑刀' ? '規格-尺寸-材質' : '規格')} /></Field>
      <div className="form-grid dynamic-fields">
        <Field label="規格 *"><EditableCombo options={specificationOptions} required value={data.specification} onChange={(value) => set('specification', value.toUpperCase())} placeholder={data.category === '刀片' ? 'ex. VNMG' : data.category === '鑽銑刀' ? 'ex. EMM4100' : '選擇或輸入新規格'}/></Field>
        {(data.category === '刀片' || data.category === '鑽銑刀') && <Field label="尺寸 *"><EditableCombo options={sizeOptions} required value={data.size} onChange={(value) => set('size', value.toUpperCase())} placeholder={data.category === '刀片' ? 'ex. 160404' : '輸入切削直徑值'}/></Field>}
        {data.category === '刀片' && <Field label="斷屑槽（無則留空）"><EditableCombo options={chipbreakerOptions} value={data.chipbreaker} onChange={(value) => set('chipbreaker', value.toUpperCase())} placeholder="ex. MS"/></Field>}
        {(data.category === '刀片' || data.category === '鑽銑刀') && <Field label="材質 *"><EditableCombo options={materialOptions} required value={data.material} onChange={(value) => set('material', value.toUpperCase())} placeholder={data.category === '刀片' ? 'ex. PR1535' : 'ex. 鎢鋼'}/></Field>}
      </div>
      <div className="form-grid common-fields">
        <Field label="廠牌"><EditableCombo options={brandOptions} value={data.brand} onChange={(value) => set('brand', value)} placeholder="選擇或輸入新廠牌"/></Field>
        <Field label="供應商"><EditableCombo options={supplierOptions} value={data.supplier} onChange={(value) => set('supplier', value)} placeholder="選擇或輸入新供應商"/></Field>
        <Field label="單位"><EditableCombo options={unitOptions} required value={data.unit} onChange={(value) => set('unit', value)} placeholder="片"/></Field>
      </div>
      {!product && <fieldset className="initial-stock"><legend>目前庫存（可只填數量，單價找不到時可留空）</legend><div className="form-grid three"><Field label="數量"><input min="0" step="1" type="number" inputMode="numeric" pattern="[0-9]*" value={data.initial_quantity || ''} onChange={(e) => set('initial_quantity', Number(e.target.value))}/></Field><Field label="單價"><input min="0" step="0.01" type="number" inputMode="numeric" pattern="[0-9]*" value={data.initial_unit_price || ''} onChange={(e) => set('initial_unit_price', Number(e.target.value))} placeholder="可留空"/></Field><Field label="購入日期"><input type="date" value={data.purchased_at} onChange={(e) => set('purchased_at', e.target.value)}/></Field></div><TotalPreview quantity={data.initial_quantity} unitPrice={data.initial_unit_price} total={initialTotalPrice}/></fieldset>}
      <Field label="備註" wide><textarea value={data.notes} onChange={(e) => set('notes', e.target.value)} placeholder="選填"/></Field>
      {formError && <p className="form-error">{formError}</p>}
      <button disabled={busy} className="button modal-submit primary">{busy ? '儲存中…' : '儲存'}</button>
      {initialBarcode && <button disabled={busy} className="button modal-submit scan-button" type="button" onClick={close}>返回條碼掃描</button>}
    </form>
  </Modal>
}

function EditableCombo({ options, value, onChange, placeholder, required = false }: { options: string[]; value: string; onChange: (value: string) => void; placeholder: string; required?: boolean }) {
  const [open, setOpen] = useState(false)
  const inputRef = useRef<HTMLInputElement>(null)
  const normalized = value.trim().toLocaleLowerCase('zh-Hant')
  const filtered = options.filter((option) => !normalized || option.toLocaleLowerCase('zh-Hant').includes(normalized)).slice(0, 20)
  const choose = (option: string) => { onChange(option); setOpen(false); inputRef.current?.blur() }
  return <span className="editable-combo" onBlur={() => window.setTimeout(() => setOpen(false), 120)}>
    <input ref={inputRef} autoComplete="off" required={required} role="combobox" aria-expanded={open} value={value} onFocus={() => setOpen(true)} onClick={() => setOpen(true)} onChange={(event) => { onChange(event.target.value); setOpen(true) }} placeholder={placeholder}/><ChevronDown aria-hidden="true"/>
    {open && <span className="product-picker-menu editable-combo-menu" role="listbox">
      {filtered.length === 0 ? <span className="product-picker-empty">可直接輸入新資料</span> : filtered.map((option) => <button type="button" role="option" aria-selected={option === value} key={option} onMouseDown={(event) => event.preventDefault()} onClick={() => choose(option)}><span>{option}</span></button>)}
    </span>}
  </span>
}

function PriceHistoryModal({ product, movements, close }: { product: Product; movements: Movement[]; close: () => void }) {
  const entries = priceEntriesFor(product.id, movements)
  const prices = entries.map((entry) => entry.unit_price || 0)
  const latest = prices[0] ?? product.unit_price
  const minimum = prices.length ? Math.min(...prices) : product.unit_price
  const maximum = prices.length ? Math.max(...prices) : product.unit_price
  return <Modal close={close} title={`歷史價格 ${product.code}`} size="price-history">
    <div className="price-stats">
      <StatCard label="最新" value={formatMoney(latest)} note="目前採購單價" tone="yellow" />
      <StatCard label="最低" value={formatMoney(minimum)} note="歷史最低" />
      <StatCard label="最高" value={formatMoney(maximum)} note="歷史最高" />
    </div>
    <div className="price-history-table"><table className="data-table"><thead><tr><th>購入日期</th><th className="number">單價</th><th className="number">數量</th><th>供應商</th></tr></thead><tbody>
      {entries.length === 0 ? <EmptyRow colSpan={4} label="尚無入庫價格紀錄"/> : entries.map((entry, index) => {
        const olderPrice = entries[index + 1]?.unit_price
        const difference = olderPrice === null || olderPrice === undefined ? null : (entry.unit_price || 0) - olderPrice
        return <tr key={entry.id}><td className="nowrap">{formatDate(entry.movement_date)}</td><td className="number history-price">{formatMoney(entry.unit_price || 0)}{difference !== null && difference !== 0 && <small className={difference > 0 ? 'price-up' : 'price-down'}>{difference > 0 ? '▲' : '▼'}{formatMoney(Math.abs(difference))}</small>}</td><td className="number muted">{formatNumber(entry.quantity)}</td><td>{entry.supplier || product.supplier || '—'}</td></tr>
      })}
    </tbody></table></div>
  </Modal>
}

function AdminDeleteModal({ request, close }: { request: AdminDeleteRequest; close: (result: boolean) => void }) {
  const [password, setPassword] = useState('')
  const [error, setError] = useState('')
  const [deleting, setDeleting] = useState(false)
  const submit = async (event: FormEvent) => {
    event.preventDefault(); setError('')
    if (!/^\d{6,}$/.test(password)) { setError('請輸入至少 6 位數字登入密碼'); return }
    setDeleting(true)
    try { await verifyOwnerPassword(password); await request.execute(); close(true) }
    catch (submitError) { setError(messageFor(submitError)) }
    finally { setDeleting(false) }
  }
  return <Modal title={request.title} close={() => !deleting && close(false)}>
    <div className="admin-delete-warning"><Trash2/><div><strong>此操作無法復原</strong><span>{request.description}</span></div></div>
    <form onSubmit={submit} onKeyDown={preventInputEnterSubmit}>
      <Field label="登入密碼" wide><input autoFocus type="password" inputMode="numeric" pattern="[0-9]{6,}" autoComplete="current-password" value={password} onChange={(event) => setPassword(event.target.value.replace(/\D/g, ''))} placeholder="至少 6 位數字" disabled={deleting}/></Field>
      {error && <div className="form-error">{error}</div>}
      <button className="button outbound admin-action" type="submit" disabled={deleting || password.length < 4}><Trash2/>{deleting ? '正在驗證並刪除…' : '驗證登入密碼並刪除'}</button>
    </form>
  </Modal>
}

function AdminDataModal({ close, afterImport }: { close: () => void; afterImport: () => Promise<void> }) {
  const [password, setPassword] = useState('')
  const [preview, setPreview] = useState<WorkbookPreview | null>(null)
  const [fileName, setFileName] = useState('')
  const [confirmation, setConfirmation] = useState('')
  const [busy, setBusy] = useState<'export' | 'parse' | 'import' | null>(null)
  const [message, setMessage] = useState('')
  const [success, setSuccess] = useState(false)

  const exportExcel = async () => {
    setMessage(''); setSuccess(false)
    if (!/^\d{6,}$/.test(password)) { setMessage('請先輸入至少 6 位數字登入密碼'); return }
    setBusy('export')
    try {
      await verifyOwnerPassword(password)
      const result = await exportAllAdminData()
      const { downloadAdminWorkbook } = await import('./lib/adminDataWorkbook')
      await downloadAdminWorkbook(result.data, result.exported_at)
      setSuccess(true); setMessage('完整備份已下載，請保留一份未修改的原始檔')
    } catch (error) { setMessage(messageFor(error)) }
    finally { setBusy(null) }
  }

  const selectFile = async (file?: File) => {
    setPreview(null); setFileName(file?.name || ''); setConfirmation(''); setMessage(''); setSuccess(false)
    if (!file) return
    setBusy('parse')
    try {
      const { parseAdminWorkbook } = await import('./lib/adminDataWorkbook')
      const parsed = await parseAdminWorkbook(file)
      setPreview(parsed); setSuccess(true); setMessage('Excel 格式檢查完成，請確認筆數後再匯入')
    } catch (error) { setMessage(messageFor(error)) }
    finally { setBusy(null) }
  }

  const importExcel = async () => {
    setMessage(''); setSuccess(false)
    if (!/^\d{6,}$/.test(password)) { setMessage('請輸入至少 6 位數字登入密碼'); return }
    if (!preview) { setMessage('請先選擇並檢查 Excel 檔案'); return }
    if (confirmation !== '覆蓋資料') { setMessage('請輸入「覆蓋資料」確認執行'); return }
    setBusy('import')
    try {
      await verifyOwnerPassword(password)
      const counts = await importAllAdminData(preview.payload)
      setSuccess(true)
      setMessage(`匯入完成：刀具 ${counts?.products || 0}、出入庫 ${counts?.movements || 0}、加工紀錄 ${counts?.production_records || 0} 筆`)
      await afterImport()
      setPreview(null); setFileName(''); setConfirmation('')
    } catch (error) { setMessage(messageFor(error)) }
    finally { setBusy(null) }
  }

  return <Modal title="資料備份／還原" close={() => !busy && close()} size="admin-data">
    <div className="admin-security-note"><LockKeyhole/><div><strong>帳號本人限定功能</strong><span>密碼由 Supabase Auth 重新驗證，不會儲存在瀏覽器或網站程式中。</span></div></div>
    <Field label="登入密碼"><input type="password" inputMode="numeric" pattern="[0-9]{6,}" autoComplete="current-password" value={password} onChange={(event) => setPassword(event.target.value.replace(/\D/g, ''))} placeholder="至少 6 位數字" disabled={Boolean(busy)}/></Field>

    <section className="admin-data-section">
      <div className="admin-data-heading"><Download/><div><strong>1. 匯出完整備份</strong><span>包含刀具、出入庫、加工數量與管理選項</span></div></div>
      <button className="button primary admin-action" type="button" onClick={() => void exportExcel()} disabled={Boolean(busy) || isDemoMode}>{busy === 'export' ? '正在製作 Excel…' : '下載完整 Excel'}</button>
    </section>

    <section className="admin-data-section danger-zone">
      <div className="admin-data-heading"><FileUp/><div><strong>2. 整理後匯入</strong><span>會完整取代目前的雲端資料，請先下載並保留備份</span></div></div>
      <label className="admin-file-picker"><input type="file" accept=".xlsx,.xls" disabled={Boolean(busy)} onChange={(event) => void selectFile(event.target.files?.[0])}/><FileSpreadsheet/><span>{busy === 'parse' ? '正在檢查 Excel…' : fileName || '選擇系統匯出的 Excel'}</span></label>
      {preview && <div className="admin-preview">
        <span><b>{preview.summary.products}</b> 刀具</span><span><b>{preview.summary.movements}</b> 出入庫</span><span><b>{preview.summary.productionRecords}</b> 加工紀錄</span><span><b>{preview.summary.options}</b> 管理選項</span>
      </div>}
      {preview && <Field label="輸入「覆蓋資料」確認"><input value={confirmation} onChange={(event) => setConfirmation(event.target.value)} placeholder="覆蓋資料" disabled={Boolean(busy)}/></Field>}
      <button className="button outbound admin-action" type="button" onClick={() => void importExcel()} disabled={Boolean(busy) || !preview || isDemoMode}>{busy === 'import' ? '正在安全匯入…' : '覆蓋並匯入雲端資料'}</button>
    </section>
    {isDemoMode && <p className="form-message">目前是示範模式，連接 Supabase 後才能使用完整備份。</p>}
    {message && <p className={success ? 'admin-success' : 'form-error'}>{message}</p>}
  </Modal>
}

function Field({ label, wide, children }: { label: string; wide?: boolean; children: ReactNode }) { return <label className={`field ${wide ? 'wide' : ''}`}><span>{label}</span>{children}</label> }

function TableScrollControls() {
  const scrollToTop = () => window.scrollTo({ top: 0, behavior: 'smooth' })
  const scrollToBottom = () => window.scrollTo({ top: document.documentElement.scrollHeight, behavior: 'smooth' })
  return <nav className="table-scroll-controls" aria-label="表格快速捲動">
    <button type="button" onClick={scrollToTop} aria-label="回到頁面最頂端" title="回到最頂端"><ChevronUp /></button>
    <button type="button" onClick={scrollToBottom} aria-label="前往頁面最底端" title="前往最底端"><ChevronDown /></button>
  </nav>
}

function Modal({ title, close, size = '', children }: { title: string; close: () => void; size?: string; children: ReactNode }) {
  useEffect(() => { const fn = (e: KeyboardEvent) => e.key === 'Escape' && close(); window.addEventListener('keydown', fn); return () => window.removeEventListener('keydown', fn) }, [close])
  const dismissInput = () => { if (document.activeElement instanceof HTMLElement) document.activeElement.blur() }
  const dismissFromBlankArea = (event: ReactPointerEvent<HTMLElement>) => {
    const target = event.target as HTMLElement
    if (!target.closest('input, select, textarea, button, a, label, [role="option"]')) dismissInput()
  }
  return <div className="modal-backdrop" role="presentation" onMouseDown={(e) => e.target === e.currentTarget && close()}>
    <section className={`modal ${size}`} role="dialog" aria-modal="true" aria-label={title} onPointerDown={dismissFromBlankArea}>
      <div className="modal-head"><h2>{title}</h2><button className="icon-button" onClick={close} aria-label="關閉"><X /></button></div>
      {/* 暫停顯示手機彈窗的鍵盤收合按鈕，待確認是否恢復。
      <button type="button" className="modal-input-dismiss" onClick={dismissInput}><ChevronDown />完成輸入／收起選單</button>
      */}
      {children}
    </section>
  </div>
}

function PageSkeleton() { return <div className="skeleton"><i/><i/><i/></div> }
