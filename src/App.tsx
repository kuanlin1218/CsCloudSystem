import { useEffect, useMemo, useRef, useState, type FormEvent, type KeyboardEvent as ReactKeyboardEvent, type PointerEvent as ReactPointerEvent, type ReactNode } from 'react'
import type { Session, User } from '@supabase/supabase-js'
import {
  ArrowDownToLine, ArrowLeftRight, ArrowUpFromLine, Boxes, CalendarDays, ChevronDown,
  Calculator, ChevronLeft, Clock3, ClipboardCheck, ClipboardList, Database, Download, History, Home, LayoutGrid,
  FileSpreadsheet, FileUp, ListPlus, LockKeyhole, LogOut, Menu, Pencil, Plus, Save, Search, ShieldCheck, Trash2, Triangle, UserRound, Wrench, X,
} from 'lucide-react'
import { isDemoMode, supabase } from './lib/supabase'
import { isOwnerUser, OWNER_EMAIL, OWNER_USERNAME, verifyOwnerPassword } from './lib/auth'
import { composeProductName, formatDate, formatMoney, formatNumber, today } from './lib/format'
import { exportProductionXlsx } from './lib/exportProductionXlsx'
import { exportAllAdminData, importAllAdminData } from './lib/adminData'
import type { WorkbookPreview } from './lib/adminDataWorkbook'
import { createProduct, createProductionOption, createProductionRecord, deleteMovement, deleteProduct, deleteProductionOption, deleteProductionProductData, deleteProductionRecord, loadInventory, loadProductionOptions, loadProductionRecords, registerMovement, setStocktakeQuantity, updateProduct } from './lib/repository'
import { PRODUCT_CATEGORIES, type Movement, type MovementInput, type MovementType, type Page, type Product, type ProductCategory, type ProductInput, type ProductionInput, type ProductionOption, type ProductionOptionType, type ProductionRecord } from './types'

const blankProduct = (): ProductInput => ({
  specification: '', size: '', chipbreaker: '', material: '', brand: '', supplier: '',
  category: '刀片', unit: '片', notes: '', initial_quantity: 0, initial_unit_price: 0, purchased_at: today(),
})

const PAGE_VALUES: Page[] = ['home', 'production-entry', 'production-query', 'dashboard', 'movements', 'stocktake', 'products']
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

type StockModalState = { type: MovementType; productId?: string }
type AdminDeleteRequest = {
  title: string
  description: string
  execute: () => Promise<void>
  resolve: (result: boolean) => void
}

const messageFor = (error: unknown) => error instanceof Error ? error.message : (error as { message?: string })?.message || '發生未知錯誤'

export default function App() {
  const [session, setSession] = useState<Session | null>(null)
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
        setSession(nextSession)
      }
      setAuthReady(true)
    }
    void client.auth.getSession().then(({ data }) => acceptSession(data.session))
    const { data: listener } = client.auth.onAuthStateChange((event, nextSession) => {
      if (event === 'PASSWORD_RECOVERY') setForcePasswordSetup(true)
      void acceptSession(nextSession)
    })
    return () => listener.subscription.unsubscribe()
  }, [])

  if (isDemoMode) return <AuthenticatedApp session={null} />
  if (!authReady) return <LoadingScreen label="正在確認安全登入狀態…" />
  if (!session) return <LoginScreen initialMessage={authMessage} />
  if (forcePasswordSetup || session.user.user_metadata?.password_configured !== true) {
    return <PasswordSetupScreen session={session} complete={(user) => { setForcePasswordSetup(false); setSession({ ...session, user }) }} />
  }
  return <AuthenticatedApp session={session} />
}

function LoadingScreen({ label }: { label: string }) {
  return <main className="loading-screen"><span className="loader"/><p>{label}</p></main>
}

function LoginScreen({ initialMessage }: { initialMessage?: string }) {
  const [username, setUsername] = useState('')
  const [password, setPassword] = useState('')
  const [message, setMessage] = useState(initialMessage || '')
  const [success, setSuccess] = useState(false)
  const [busy, setBusy] = useState(false)

  const submit = async (event: FormEvent) => {
    event.preventDefault(); setMessage(''); setSuccess(false)
    if (username.trim().toLowerCase() !== OWNER_USERNAME) { setMessage('登入名稱或密碼不正確'); return }
    if (!/^\d{4,}$/.test(password)) { setMessage('密碼必須是至少 4 位數字'); return }
    setBusy(true)
    try { await verifyOwnerPassword(password) }
    catch (error) { setMessage(messageFor(error)) }
    finally { setBusy(false) }
  }

  const sendPasswordEmail = async () => {
    if (!supabase) return
    setBusy(true); setMessage(''); setSuccess(false)
    try {
      const redirectTo = `${window.location.origin}${import.meta.env.BASE_URL}`
      const { error } = await supabase.auth.resetPasswordForEmail(OWNER_EMAIL, { redirectTo })
      if (error) throw error
      setSuccess(true); setMessage('設定密碼郵件已寄出，請使用最新一封郵件中的連結。')
    } catch (error) { setMessage(messageFor(error)) }
    finally { setBusy(false) }
  }

  return <main className="auth-screen">
    <span className="auth-grid"/>
    <section className="auth-card" aria-label="登入系統">
      <div className="auth-brand">
        <img className="auth-logo" src={`${import.meta.env.BASE_URL}company-logo.png`} alt="詮栩公司 Logo"/>
        <p className="auth-eyebrow">CHYUAN SHEU CNC CLOUD</p>
        <h1>登入雲端系統</h1>
        <p className="auth-copy">本系統僅限指定帳號使用。</p>
      </div>
      <form onSubmit={submit}>
        <Field label="登入名稱" wide><input autoFocus autoCapitalize="none" autoCorrect="off" autoComplete="username" value={username} onChange={(event) => setUsername(event.target.value)} placeholder="輸入登入名稱" disabled={busy}/></Field>
        <Field label="密碼" wide><input type="password" inputMode="numeric" pattern="[0-9]{4,}" autoComplete="current-password" value={password} onChange={(event) => setPassword(event.target.value.replace(/\D/g, ''))} placeholder="至少 4 位數字" disabled={busy}/></Field>
        {message && <p className={success ? 'admin-success' : 'form-error'}>{message}</p>}
        <button className="button primary auth-submit" type="submit" disabled={busy}>{busy ? '正在驗證…' : '登入'}</button>
      </form>
      <button className="auth-switch" type="button" onClick={() => void sendPasswordEmail()} disabled={busy}>忘記密碼或尚未設定密碼</button>
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
    if (!/^\d{4,}$/.test(password)) { setMessage('密碼必須是至少 4 位數字'); return }
    if (password === '16866319') { setMessage('這組舊管理密碼已經公開使用過，請設定不同的新密碼'); return }
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
        <p className="auth-copy">登入名稱固定為 <b>{OWNER_USERNAME}</b>。請設定至少 4 位數字的新密碼。</p>
      </div>
      <form onSubmit={submit}>
        <Field label="新密碼" wide><input autoFocus type="password" inputMode="numeric" pattern="[0-9]{4,}" autoComplete="new-password" value={password} onChange={(event) => setPassword(event.target.value.replace(/\D/g, ''))} placeholder="至少 4 位數字" disabled={busy}/></Field>
        <Field label="再次輸入新密碼" wide><input type="password" inputMode="numeric" pattern="[0-9]{4,}" autoComplete="new-password" value={confirmation} onChange={(event) => setConfirmation(event.target.value.replace(/\D/g, ''))} placeholder="再次輸入相同數字密碼" disabled={busy}/></Field>
        {message && <p className="form-error">{message}</p>}
        <button className="button primary auth-submit" type="submit" disabled={busy}>{busy ? '正在安全儲存…' : '儲存密碼並進入系統'}</button>
      </form>
    </section>
  </main>
}

function AuthenticatedApp({ session }: { session: Session | null }) {
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
      setPage(pageFromLocation())
      setMobileNav(false)
      setStockModal(null); setProductModal(null); setPriceHistoryProduct(null)
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
    if (window.location.hash !== `#${next}`) window.history.pushState({ page: next }, '', pageUrl(next))
    setPage(next)
    setMobileNav(false)
    void refresh()
  }
  const afterSave = async (message: string) => { await refresh(false); setToast(message) }
  const requestAdminDelete = (title: string, description: string, execute: () => Promise<void>) => new Promise<boolean>((resolve) => {
    setAdminDelete({ title, description, execute, resolve })
  })
  const isHome = page === 'home'
  const isInventory = page === 'dashboard' || page === 'movements' || page === 'stocktake' || page === 'products'
  const productionEntryProductCodes = useMemo(() => productionOptions.filter((option) => option.option_type === 'product').map((option) => option.value).sort(), [productionOptions])
  const productionQueryProductCodes = useMemo(() => [...new Set([
    ...productionEntryProductCodes,
    ...productionRecords.map((record) => record.product_code),
  ].filter(Boolean))].sort(), [productionEntryProductCodes, productionRecords])
  const productionMachines = useMemo(() => productionOptions.filter((option) => option.option_type === 'machine').map((option) => option.value).sort(), [productionOptions])

  return (
    <div className={`app-shell ${isHome ? 'portal-shell' : ''}`}>
      {!isDemoMode && !isInventory && <button className="portal-logout" type="button" onClick={() => void supabase?.auth.signOut()}><LogOut />登出</button>}
      {isInventory && <aside className={`sidebar ${mobileNav ? 'is-open' : ''}`}>
        <div className="brand">
          <span className="brand-mark"><Triangle fill="currentColor" /><i /></span>
          <span><strong>庫存管理系統</strong><small>INVENTORY SYSTEM</small></span>
          <button className="icon-button mobile-close" onClick={() => setMobileNav(false)} aria-label="關閉選單"><X /></button>
        </div>
        <nav>
          <NavItem active={false} icon={<Home />} onClick={() => nav('home')}>系統首頁</NavItem>
          <NavItem active={page === 'dashboard'} icon={<LayoutGrid />} onClick={() => nav('dashboard')}>總覽</NavItem>
          <NavItem active={page === 'movements'} icon={<ArrowLeftRight />} onClick={() => nav('movements')}>出入庫</NavItem>
          <NavItem active={page === 'stocktake'} icon={<ClipboardCheck />} onClick={() => nav('stocktake')}>庫存盤點</NavItem>
          <NavItem active={page === 'products'} icon={<Boxes />} onClick={() => nav('products')}>刀具資料</NavItem>
        </nav>
        <div className="sidebar-foot">
          <span className="status-line">{isDemoMode ? <ShieldCheck /> : <Database />} {isDemoMode ? '示範模式' : realtimeConnected ? '雲端即時同步' : '雲端連線中'}</span>
          <span className="status-line"><UserRound /> {isDemoMode ? '本機示範資料' : session?.user.user_metadata?.username || OWNER_USERNAME}</span>
          {!isDemoMode && <button className="sidebar-logout" type="button" onClick={() => void supabase?.auth.signOut()}><LogOut />登出系統</button>}
        </div>
      </aside>}

      {isInventory && mobileNav && <button className="nav-scrim" onClick={() => setMobileNav(false)} aria-label="關閉選單" />}

      <main className={`main-panel ${!isInventory ? 'wide-panel' : ''} ${isHome ? 'portal-panel' : ''}`}>
        {isInventory && <header className="mobile-header">
          <button className="icon-button" onClick={() => setMobileNav(true)} aria-label="開啟選單"><Menu /></button>
          <strong>庫存管理系統</strong>
          <span className="mobile-dot" />
        </header>}
        {isInventory && isDemoMode && <div className="demo-banner"><ShieldCheck /> 目前是可操作的示範模式，資料儲存在這台裝置。設定 Supabase 後會自動切換為雲端模式。</div>}
        {error && <div className="error-banner">{error}<button onClick={() => void refresh()}>重試</button></div>}

        {page === 'home' && <HomePage navigate={nav} openAdmin={() => setAdminDataModal(true)} />}
        {page === 'production-entry' && <ProductionEntryPage records={productionRecords} productCodes={productionEntryProductCodes} machines={productionMachines} loading={loading} goHome={() => nav('home')} addOption={setProductionOptionModal} onSave={async (input) => {
          await createProductionRecord(input); await afterSave('加工數量已登記')
        }} onDeleteLast={async (record) => {
          const deleted = await requestAdminDelete('刪除最後一筆加工紀錄', `${record.product_code}・${record.machine}・${record.quantity}，刪除後無法復原。`, () => deleteProductionRecord(record.id))
          if (deleted) await afterSave('最後一筆加工紀錄已刪除')
        }} />}
        {page === 'production-query' && <ProductionQueryPage records={productionRecords} productCodes={productionQueryProductCodes} loading={loading} goHome={() => nav('home')} onDeleteRecord={async (record) => {
          const deleted = await requestAdminDelete('刪除加工紀錄', `${formatDate(record.work_date)}・${record.product_code}・${record.machine}・${record.quantity}，刪除後無法復原。`, () => deleteProductionRecord(record.id))
          if (deleted) await afterSave('所選加工紀錄已刪除')
          return deleted
        }} onDeleteProductRecords={async (productCode, count) => {
          const deleted = await requestAdminDelete('永久刪除產品加工資料', `將刪除產品編號「${productCode}」、管理清單選項與全部 ${count} 筆加工紀錄。`, () => deleteProductionProductData(productCode))
          if (deleted) await afterSave(`${productCode} 的產品編號與全部加工紀錄已永久刪除`)
          return deleted
        }} />}
        {page === 'dashboard' && <Dashboard products={products} movements={movements} loading={loading} openStock={(type) => setStockModal({ type })} />}
        {page === 'movements' && <MovementsPage products={products} movements={movements} loading={loading} openStock={(type) => setStockModal({ type })} onDelete={async (movement) => {
          const product = movement.product || products.find((item) => item.id === movement.product_id)
          const deleted = await requestAdminDelete('刪除出入庫紀錄', `${formatDate(movement.movement_date)} 的${movement.movement_type === 'in' ? '入庫' : '出庫'}紀錄「${product?.name || ''}」。庫存會自動重新計算。`, () => deleteMovement(movement.id))
          if (deleted) await afterSave('出入庫紀錄已刪除，庫存已重新計算')
        }} />}
        {page === 'stocktake' && <StocktakePage products={products} loading={loading} add={() => setProductModal('new')} onAdjust={async (product, actualQuantity, notes) => {
          if (actualQuantity === product.stock_quantity) return true
          if (!window.confirm(`確定將 ${product.code}「${product.name}」的庫存由 ${product.stock_quantity} ${product.unit}調整為 ${actualQuantity} ${product.unit}？`)) return false
          try {
            await setStocktakeQuantity({ product_id: product.id, actual_quantity: actualQuantity, notes, request_key: crypto.randomUUID() })
            await afterSave(`${product.code} 的盤點數量已更新`)
            return true
          } catch (e) { setError(messageFor(e)); return false }
        }} />}
        {page === 'products' && <ProductsPage products={products} movements={movements} loading={loading} add={() => setProductModal('new')} edit={setProductModal} openStock={(type, productId) => setStockModal({ type, productId })} viewHistory={setPriceHistoryProduct} onDelete={async (p) => {
          const deleted = await requestAdminDelete('刪除刀具資料', `將刪除 ${p.code}「${p.name}」及其全部出入庫紀錄。`, () => deleteProduct(p.id))
          if (deleted) await afterSave('商品已刪除')
        }} />}
      </main>

      {stockModal && <StockModal type={stockModal.type} initialProductId={stockModal.productId} products={products} close={() => setStockModal(null)} switchType={(type) => setStockModal((current) => current ? { ...current, type } : null)} save={async (input) => {
        await registerMovement(input); setStockModal(null); await afterSave(input.movement_type === 'in' ? '入庫完成' : '出庫完成')
      }} />}
      {productModal && <ProductModal product={productModal === 'new' ? null : productModal} products={products} close={() => setProductModal(null)} save={async (input) => {
        if (productModal === 'new') await createProduct(input); else await updateProduct(productModal.id, input)
        setProductModal(null); await afterSave(productModal === 'new' ? '商品已建立' : '商品已更新')
      }} />}
      {priceHistoryProduct && <PriceHistoryModal product={priceHistoryProduct} movements={movements} close={() => setPriceHistoryProduct(null)} />}
      {productionOptionModal && <ProductionOptionModal type={productionOptionModal} values={(productionOptionModal === 'product' ? productionEntryProductCodes : productionMachines)} close={() => setProductionOptionModal(null)} save={async (value) => {
        await createProductionOption(productionOptionModal, value); await afterSave(productionOptionModal === 'product' ? '產品編號已新增' : '機台已新增')
      }} remove={async (value) => {
        const label = productionOptionModal === 'product' ? '產品編號' : '機台'
        const deleted = await requestAdminDelete(`刪除${label}`, `確定刪除${label}「${value}」？過去已登記的加工紀錄會保留。`, () => deleteProductionOption(productionOptionModal, value))
        if (deleted) await afterSave(`${label}已刪除`)
      }} />}
      {adminDataModal && <AdminDataModal close={() => setAdminDataModal(false)} afterImport={async () => {
        await refresh(); setToast('Excel 資料已完整匯入')
      }} />}
      {adminDelete && <AdminDeleteModal request={adminDelete} close={(result) => {
        adminDelete.resolve(result); setAdminDelete(null)
      }} />}
      {toast && <div className="toast"><ShieldCheck />{toast}</div>}
    </div>
  )
}

function NavItem({ active, icon, onClick, children }: { active: boolean; icon: ReactNode; onClick: () => void; children: ReactNode }) {
  return <button className={`nav-item ${active ? 'active' : ''}`} onClick={onClick}>{icon}<span>{children}</span></button>
}

function HomePage({ navigate, openAdmin }: { navigate: (page: Page) => void; openAdmin: () => void }) {
  return <section className="page-content home-page">
    <div className="home-hero">
      <img
          className="home-company-logo"
          src={`${import.meta.env.BASE_URL}company-logo.png`}
          alt="詮栩公司 Logo"
        />
      <span className="home-eyebrow">CHYUAN SHEU CNC CLOUD</span>
      <h1>詮栩 CNC 雲端系統</h1>
      <p>加工資訊與庫存管理整合平台</p>
    </div>
    <div className="home-section-head"><div><span>SYSTEM FUNCTIONS</span><h2>功能選擇</h2></div><small>選擇要使用的系統</small></div>
    <div className="function-grid portal-functions">
      <FunctionCard icon={<ListPlus/>} title="加工數量登記" description="登記數量並自動計算加工時間" tone="pink" onClick={() => navigate('production-entry')}/>
      <FunctionCard icon={<Search/>} title="加工數量查詢" description="依產品編號查詢與匯出紀錄" tone="sunset" onClick={() => navigate('production-query')}/>
      <FunctionCard icon={<Boxes/>} title="庫存管理系統" description="刀具庫存、出入庫與價格管理" tone="green" onClick={() => navigate('dashboard')}/>
      <FunctionCard icon={<FileSpreadsheet/>} title="資料備份／還原" description="匯出或整批匯入 Supabase Excel" tone="purple" onClick={openAdmin}/>
    </div>
  </section>
}

function ProductionHeader({ title, eyebrow, icon, goHome }: { title: string; eyebrow: string; icon: ReactNode; goHome: () => void }) {
  return <div className="production-head">
    <button className="production-back" onClick={goHome}><ChevronLeft />系統首頁</button>
    <div><span className="production-head-icon">{icon}</span><span>{eyebrow}</span><h1>{title}</h1></div>
    <span className={`production-mode ${isDemoMode ? 'demo' : ''}`}>{isDemoMode ? '示範資料' : '雲端資料'}</span>
  </div>
}

function ProductionEntryPage({ records, productCodes, machines, loading, goHome, addOption, onSave, onDeleteLast }: { records: ProductionRecord[]; productCodes: string[]; machines: string[]; loading: boolean; goHome: () => void; addOption: (type: ProductionOptionType) => void; onSave: (input: ProductionInput) => Promise<void>; onDeleteLast: (record: ProductionRecord) => Promise<void> }) {
  const [form, setForm] = useState<ProductionInput>({ work_date: today(), product_code: '', machine: '', hours: 0, minutes: 0, quantity: 0, notes: '' })
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

  return <section className="production-page production-entry-page">
    <ProductionHeader title="加工數量登記" eyebrow="PRODUCTION ENTRY" icon={<ListPlus />} goHome={goHome} />
    <form className="production-card entry-card" onSubmit={submit} onKeyDown={preventInputEnterSubmit}>
      <div className="production-data-tools">
        <div><span>基礎資料</span><small>管理產品編號與機台，再從登記表單選擇</small></div>
        <div><button type="button" onClick={() => addOption('product')}><Wrench />管理產品編號</button><button type="button" onClick={() => addOption('machine')}><Wrench />管理機台</button></div>
      </div>
      <div className="production-field-row"><label>日期</label><input type="date" value={form.work_date} onChange={(e) => setForm({ ...form, work_date: e.target.value })} required /></div>
      <div className="production-field-row"><label>產品編號</label><select value={form.product_code} onChange={(e) => setForm({ ...form, product_code: e.target.value })} required><option value="">請選擇產品編號</option>{productCodes.map((value) => <option key={value} value={value}>{value}</option>)}</select></div>
      <div className="production-field-row"><label>機台</label><select value={form.machine} onChange={(e) => setForm({ ...form, machine: e.target.value })} required><option value="">請選擇機台</option>{machines.map((value) => <option key={value} value={value}>{value}</option>)}</select></div>

      <div className="time-entry-panel">
        <div className="time-entry-head"><div><Clock3 /><span><strong>實際加工時間</strong><small>可直接輸入；需要相減時再開啟計算工具</small></span></div><button type="button" onClick={() => setShowTimeCalculator(true)}><Calculator />加工時間計算</button></div>
        <div className="duration-fields compact">
          <label className="actual-hours-field"><span>實際工時（時）</span><input className="actual-hours-input" type="number" inputMode="numeric" pattern="[0-9]*" min="0" step="1" value={form.hours || ''} onChange={(e) => setForm({ ...form, hours: Math.max(0, Number(e.target.value)) })} aria-label="加工時數" /></label>
          <label className="actual-minutes-field"><span>實際工時（分）</span><input className="actual-minutes-input" type="number" inputMode="numeric" pattern="[0-9]*" min="0" max="59" step="1" value={form.minutes || ''} onChange={(e) => setForm({ ...form, minutes: Math.max(0, Number(e.target.value)) })} aria-label="加工分鐘" /></label>
          <div><span>實際總時間</span><strong>{formatNumber(form.hours * 60 + form.minutes)} 分鐘</strong></div>
        </div>
      </div>

      <div className="production-field-row"><label>數量</label><input type="number" inputMode="numeric" pattern="[0-9]*" min="1" step="1" value={form.quantity || ''} onChange={(e) => setForm({ ...form, quantity: Math.max(0, Number(e.target.value)) })} placeholder="輸入完成數量" required /></div>
      <div className="production-field-row note-row"><label>備註</label><textarea value={form.notes} onChange={(e) => setForm({ ...form, notes: e.target.value })} placeholder="選填" /></div>
      {formError && <div className="form-error production-error">{formError}</div>}
      <button className="production-action save-action" disabled={saving} type="submit"><Save />{saving ? '儲存中…' : '送出登記'}</button>
      <button className="production-action delete-action" disabled={loading || !latest} type="button" onClick={() => latest && void onDeleteLast(latest)}><Trash2 />刪除最後一筆{latest ? `｜${latest.product_code}・${latest.machine}` : ''}</button>
    </form>
    {showTimeCalculator && <ProductionTimeCalculatorModal initialHours={form.hours} initialMinutes={form.minutes} close={() => setShowTimeCalculator(false)} apply={(hours, minutes) => { setForm((current) => ({ ...current, hours, minutes })); setShowTimeCalculator(false) }} />}
  </section>
}

type ProductionQueryMode = 'range' | 'latest' | 'all'

function ProductionQueryPage({ records, productCodes, loading, goHome, onDeleteRecord, onDeleteProductRecords }: { records: ProductionRecord[]; productCodes: string[]; loading: boolean; goHome: () => void; onDeleteRecord: (record: ProductionRecord) => Promise<boolean>; onDeleteProductRecords: (productCode: string, count: number) => Promise<boolean> }) {
  const [productCode, setProductCode] = useState('')
  const [from, setFrom] = useState(`${today().slice(0, 8)}01`)
  const [to, setTo] = useState(today())
  const [mode, setMode] = useState<ProductionQueryMode | null>(null)
  const [showForecast, setShowForecast] = useState(false)
  const [queryError, setQueryError] = useState('')
  const [deleting, setDeleting] = useState(false)
  const [deletingRecordId, setDeletingRecordId] = useState('')
  const [exporting, setExporting] = useState(false)
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
  const totals = useMemo(() => {
    const totalQuantity = results.reduce((sum, record) => sum + record.quantity, 0)
    const totalMinutes = results.reduce((sum, record) => sum + record.total_minutes, 0)
    const totalDays = new Set(results.map((record) => record.work_date)).size
    return {
      totalQuantity, totalMinutes, totalDays,
      dailyAverage: totalDays ? Math.round(totalQuantity / totalDays) : 0,
      pieceSeconds: totalQuantity ? Math.round(totalMinutes * 60 / totalQuantity) : 0,
    }
  }, [results])
  const runQuery = (nextMode: ProductionQueryMode) => {
    setQueryError('')
    if (!productCode.trim()) { setMode(null); setQueryError('請先輸入或選擇產品編號'); return }
    if (nextMode === 'range' && from && to && from > to) { setMode(null); setQueryError('開始日期不能晚於結束日期'); return }
    setMode(nextMode)
  }

  const downloadExcel = async () => {
    if (mode !== 'all') { setQueryError('請先執行「查詢全部」後再下載 Excel'); return }
    if (!results.length) { setQueryError('目前沒有可下載的查詢結果'); return }
    setExporting(true); setQueryError('')
    try { await exportProductionXlsx(productCode, results) }
    catch (error) { setQueryError(`Excel 產生失敗：${messageFor(error)}`) }
    finally { setExporting(false) }
  }

  const deleteAllRecords = async () => {
    if (!productCode) return
    setDeleting(true); setQueryError('')
    const deleted = await onDeleteProductRecords(productCode, selectedRecords.length)
    if (deleted) { setMode(null); setProductCode(''); setShowForecast(false) }
    setDeleting(false)
  }

  const deleteOneRecord = async (record: ProductionRecord) => {
    setDeletingRecordId(record.id); setQueryError('')
    await onDeleteRecord(record)
    setDeletingRecordId('')
  }

  return <section className="production-page production-query-page">
    <ProductionHeader title="加工數量查詢" eyebrow="PRODUCTION LOOKUP" icon={<Search />} goHome={goHome} />
    <div className="production-card query-card">
      <div className="query-fields">
        <div className="production-field-row"><label>開始日期</label><input type="date" value={from} onChange={(e) => setFrom(e.target.value)} /></div>
        <div className="production-field-row"><label>結束日期</label><input type="date" value={to} onChange={(e) => setTo(e.target.value)} /></div>
        <div className="production-field-row"><label>產品編號</label><select value={productCode} onChange={(e) => { setProductCode(e.target.value); setMode(null); setShowForecast(false); setQueryError('') }}><option value="">請選擇產品編號</option>{productCodes.map((value) => <option key={value} value={value}>{value}</option>)}</select></div>
      </div>
      <div className="query-utility-bar"><div><Calculator /><span><strong>訂單完工預估</strong><small>低頻工具，需要時再開啟</small></span></div><button onClick={() => setShowForecast(true)} disabled={!productCode}><Calculator />開啟預估工具</button></div>
      <div className="query-actions">
        <button className="query-button range" onClick={() => runQuery('range')}><CalendarDays />日期區間查詢</button>
        <button className="query-button latest" onClick={() => runQuery('latest')}><Clock3 />查詢最後一筆</button>
        <button className="query-button all" onClick={() => runQuery('all')}><Search />查詢全部</button>
        <button className="query-button export" onClick={() => void downloadExcel()} disabled={mode !== 'all' || !results.length || exporting} title={mode === 'all' ? '下載全部查詢結果' : '請先按「查詢全部」'}><Download />{exporting ? '產生中…' : '下載 EXCEL'}</button>
      </div>
      <div className="query-danger-zone"><div><strong>永久刪除產品資料</strong><small>會同時刪除全部加工紀錄，以及「管理產品編號」中的產品編號，無法復原。</small></div><button onClick={() => void deleteAllRecords()} disabled={!productCode || deleting}><Trash2 />{deleting ? '刪除中…' : `永久刪除產品與資料（${selectedRecords.length} 筆）`}</button></div>
      {queryError && <div className="form-error production-error">{queryError}</div>}
    </div>

    {loading ? <PageSkeleton /> : mode && <div className="production-results">
      <div className="result-title"><span>產品編號</span><strong>{productCode.trim().toUpperCase()}</strong><small>{mode === 'range' ? `${formatDate(from)} ～ ${formatDate(to)}` : mode === 'latest' ? `最後日期 ${formatDate(results[0]?.work_date || '')}・全部機台` : '全部紀錄'}</small></div>
      {results.length === 0 ? <div className="production-empty"><ClipboardList />沒有符合條件的加工紀錄</div> : mode === 'latest' ?
        <div className="production-table-wrap"><table className="production-table compact"><thead><tr><th>日期</th><th>產品・機台</th><th>數量</th><th className="record-action-column" aria-label="操作" /></tr></thead><tbody>{results.map((record) => <tr key={record.id}><td>{formatDate(record.work_date)}</td><td>{record.product_code}・{record.machine}</td><td>{formatNumber(record.quantity)}</td><td><button className="production-row-delete" onClick={() => void deleteOneRecord(record)} disabled={Boolean(deletingRecordId)} aria-label={`刪除 ${record.machine} 加工紀錄`}><Trash2 />{deletingRecordId === record.id ? '刪除中' : '刪除'}</button></td></tr>)}</tbody><tfoot><tr><td colSpan={2}>總數量</td><td>{formatNumber(totals.totalQuantity)}</td><td /></tr></tfoot></table></div>
        : <>
          <div className="production-table-wrap"><table className="production-table"><thead><tr><th>日期</th><th>機台</th><th>時</th><th>分</th><th>數量</th><th>總時間</th><th>備註</th><th className="record-action-column" aria-label="操作" /></tr></thead><tbody>{results.map((record) => <tr key={record.id}><td>{formatDate(record.work_date)}</td><td>{record.machine}</td><td>{record.hours}</td><td>{record.minutes}</td><td>{formatNumber(record.quantity)}</td><td>{formatNumber(record.total_minutes)}</td><td>{record.notes || '—'}</td><td><button className="production-row-delete" onClick={() => void deleteOneRecord(record)} disabled={Boolean(deletingRecordId)} aria-label={`刪除 ${formatDate(record.work_date)} ${record.machine} 加工紀錄`}><Trash2 />{deletingRecordId === record.id ? '刪除中' : '刪除'}</button></td></tr>)}</tbody></table></div>
          <div className="production-totals">
            <div><span>總天數</span><strong>{formatNumber(totals.totalDays)}</strong></div>
            <div><span>總數量</span><strong>{formatNumber(totals.totalQuantity)}</strong></div>
            <div><span>總時間</span><strong>{formatNumber(totals.totalMinutes)} <small>分</small></strong></div>
            <div><span>1天平均</span><strong>{formatNumber(totals.dailyAverage)} <small>顆</small></strong></div>
            <div><span>1顆平均</span><strong>{formatNumber(totals.pieceSeconds)} <small>秒</small></strong></div>
          </div>
        </>}
    </div>}
    {showForecast && productCode && <OrderForecastModal productCode={productCode} records={selectedRecords} close={() => setShowForecast(false)} />}
  </section>
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

function OrderForecastModal({ productCode, records, close }: { productCode: string; records: ProductionRecord[]; close: () => void }) {
  const [orderQuantity, setOrderQuantity] = useState(0)
  const forecast = useMemo(() => {
    const completedQuantity = records.reduce((sum, record) => sum + record.quantity, 0)
    const historyDays = new Set(records.map((record) => record.work_date)).size
    const dailyAverage = historyDays ? completedQuantity / historyDays : 0
    const remainingQuantity = Math.max(orderQuantity - completedQuantity, 0)
    const estimatedDays = dailyAverage > 0 ? Math.ceil(remainingQuantity / dailyAverage) : null
    let estimatedDate = ''
    if (estimatedDays !== null) {
      const date = new Date(`${today()}T12:00:00`)
      date.setDate(date.getDate() + estimatedDays)
      estimatedDate = date.toISOString().slice(0, 10)
    }
    return { completedQuantity, historyDays, dailyAverage, remainingQuantity, estimatedDays, estimatedDate }
  }, [records, orderQuantity])
  return <Modal title="訂單完工預估" close={close} size="large forecast-modal">
    <div className="forecast-modal-product"><span>產品編號</span><strong>{productCode}</strong></div>
    <div className="forecast-head modal-forecast-head"><div><Calculator /><span><strong>依全部歷史資料估算</strong><small>輸入訂單總數量後自動計算</small></span></div><label><span>訂單總數量</span><input className="order-quantity-input" autoFocus type="number" inputMode="numeric" pattern="[0-9]*" min="1" step="1" value={orderQuantity || ''} onChange={(event) => setOrderQuantity(Math.max(0, Number(event.target.value)))} placeholder="輸入訂單數量" /></label></div>
    {orderQuantity <= 0 ? <p className="forecast-placeholder">輸入訂單總數量後，會顯示剩餘數量、天數與預估完工日期</p> : !forecast.historyDays ? <p className="forecast-placeholder warning">此產品尚無歷史加工資料，暫時無法估算完工時間</p> : <div className="forecast-grid">
      <div><span>歷史完成量</span><strong>{formatNumber(forecast.completedQuantity)}</strong></div>
      <div><span>平均每日產量</span><strong>{formatNumber(Math.round(forecast.dailyAverage))}</strong></div>
      <div className="accent"><span>剩餘數量</span><strong>{formatNumber(forecast.remainingQuantity)}</strong></div>
      <div><span>預估剩餘天數</span><strong>{formatNumber(forecast.estimatedDays || 0)} <small>天</small></strong></div>
      <div className="accent"><span>預估完工日期</span><strong>{formatDate(forecast.estimatedDate)}</strong></div>
    </div>}
    <p className="forecast-note">公式：歷史總產量 ÷ 實際加工天數＝每日平均；剩餘數量 ÷ 每日平均＝預估曆日天數。</p>
    <div className="modal-tool-actions"><button className="button primary" type="button" onClick={close}>完成</button></div>
  </Modal>
}

function ProductionOptionModal({ type, values, close, save, remove }: { type: ProductionOptionType; values: string[]; close: () => void; save: (value: string) => Promise<void>; remove: (value: string) => Promise<void> }) {
  const [value, setValue] = useState('')
  const [error, setError] = useState('')
  const [saving, setSaving] = useState(false)
  const [removing, setRemoving] = useState('')
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
  return <Modal title={`管理${label}`} close={close}>
    <form onSubmit={submit} onKeyDown={preventInputEnterSubmit}>
      <Field label={label} wide><input autoFocus value={value} onChange={(event) => setValue(event.target.value)} placeholder={type === 'product' ? '例如 OAS-025<SP263735>' : '例如 MT04'} /></Field>
      <p className="option-help">新增後會出現在加工登記選單；刪除不會影響過去的加工紀錄。</p>
      {error && <div className="form-error">{error}</div>}
      <button className="button primary option-add-button" disabled={saving || !!removing} type="submit"><Plus />{saving ? '新增中…' : `新增${label}`}</button>
    </form>
    <div className="option-list-head"><span>目前的{label}</span><small>{values.length} 筆</small></div>
    <div className="option-list">
      {values.length === 0 ? <div className="option-empty">尚未新增任何{label}</div> : values.map((item) => <div className="option-list-item" key={item}><span>{item}</span><button className="danger-icon" disabled={!!removing} onClick={() => void removeValue(item)} aria-label={`刪除${label} ${item}`} title={`刪除${label}`}>{removing === item ? '…' : <Trash2 />}</button></div>)}
    </div>
  </Modal>
}

function FunctionCard({ icon, title, description, tone, onClick, disabled = false }: { icon: ReactNode; title: string; description: string; tone: string; onClick?: () => void; disabled?: boolean }) {
  return <button className={`function-card ${tone} ${disabled ? 'is-future' : ''}`} onClick={onClick} disabled={disabled}><span className="function-icon">{icon}</span><span className="function-copy"><strong>{title}</strong><small>{description}</small></span>{disabled ? <em>即將推出</em> : <ChevronDown className="function-arrow"/>}</button>
}

function PageTitle({ eyebrow, title, actions }: { eyebrow: string; title: string; actions?: ReactNode }) {
  return <div className="page-title"><div><span>{eyebrow}</span><h1>{title}</h1></div>{actions && <div className="title-actions">{actions}</div>}</div>
}

function StockButtons({ openStock }: { openStock: (type: MovementType) => void }) {
  return <><button className="button inbound" onClick={() => openStock('in')}><ArrowDownToLine />入庫</button><button className="button outbound" onClick={() => openStock('out')}><ArrowUpFromLine />出庫</button></>
}

function Dashboard({ products, movements, loading, openStock }: { products: Product[]; movements: Movement[]; loading: boolean; openStock: (t: MovementType) => void }) {
  const month = today().slice(0, 7)
  const inbound = movements.filter((m) => m.movement_type === 'in' && m.movement_date.startsWith(month))
  const outbound = movements.filter((m) => m.movement_type === 'out' && m.movement_date.startsWith(month))
  const stock = products.reduce((sum, p) => sum + p.stock_quantity, 0)
  const value = products.reduce((sum, p) => sum + p.stock_quantity * p.unit_price, 0)
  const inboundQty = inbound.reduce((sum, m) => sum + m.quantity, 0)
  const outboundQty = outbound.reduce((sum, m) => sum + m.quantity, 0)
  const purchased = inbound.reduce((sum, m) => sum + m.quantity * (m.unit_price || 0), 0)

  return <section className="page-content">
    <PageTitle eyebrow={formatDate(today())} title="庫存總覽" actions={<StockButtons openStock={openStock} />} />
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

function MovementsPage({ products, movements, loading, openStock, onDelete }: { products: Product[]; movements: Movement[]; loading: boolean; openStock: (t: MovementType) => void; onDelete: (movement: Movement) => void }) {
  const [query, setQuery] = useState('')
  const [type, setType] = useState<'all' | MovementType>('all')
  const [from, setFrom] = useState('')
  const [to, setTo] = useState('')
  const filtered = movements.filter((m) => {
    const product = m.product || products.find((p) => p.id === m.product_id)
    const haystack = `${product?.code} ${product?.name} ${product?.brand} ${m.supplier} ${m.notes}`.toLowerCase()
    return (!query || haystack.includes(query.toLowerCase())) && (type === 'all' || m.movement_type === type) && (!from || m.movement_date >= from) && (!to || m.movement_date <= to)
  })
  const inQty = filtered.filter((m) => m.movement_type === 'in').reduce((s, m) => s + m.quantity, 0)
  const outQty = filtered.filter((m) => m.movement_type === 'out').reduce((s, m) => s + m.quantity, 0)
  const purchased = filtered.filter((m) => m.movement_type === 'in').reduce((s, m) => s + m.quantity * (m.unit_price || 0), 0)
  return <section className="page-content">
    <PageTitle eyebrow="STOCK MOVEMENTS" title="出入庫" actions={<StockButtons openStock={openStock} />} />
    <div className="filter-card">
      <SearchField value={query} setValue={setQuery} placeholder="搜尋編號、品名、廠牌、供應商、備註" />
      <div className="filter-row">
        <label className="select-wrap"><select value={type} onChange={(e) => setType(e.target.value as 'all' | MovementType)}><option value="all">全部類型</option><option value="in">入庫</option><option value="out">出庫</option></select><ChevronDown /></label>
        <label className="date-wrap"><input type="date" value={from} onChange={(e) => setFrom(e.target.value)} aria-label="開始日期"/><CalendarDays /></label>
        <label className="date-wrap"><input type="date" value={to} onChange={(e) => setTo(e.target.value)} aria-label="結束日期"/><CalendarDays /></label>
      </div>
    </div>
    <div className="movement-summary"><span>{filtered.length} 筆</span><b className="green-text">入 +{formatNumber(inQty)}</b><b className="red-text">出 -{formatNumber(outQty)}</b><b className="yellow-text">採購 ${formatMoney(purchased)}</b></div>
    {loading ? <PageSkeleton /> : <MovementTable movements={filtered} products={products} onDelete={onDelete} />}
  </section>
}

function MovementTable({ movements, products, onDelete }: { movements: Movement[]; products: Product[]; onDelete: (movement: Movement) => void }) {
  return <div className="table-card"><table className="data-table movement-table"><thead><tr><th>類型</th><th>日期</th><th>編號</th><th>品名</th><th className="number">數量</th><th className="number">單價</th><th className="number">金額</th><th>供應商</th><th>備註</th><th aria-label="操作" /></tr></thead><tbody>
    {movements.length === 0 ? <EmptyRow colSpan={10} label="沒有符合條件的出入庫紀錄" /> : movements.map((m) => {
      const p = m.product || products.find((item) => item.id === m.product_id)
      return <tr key={m.id}><td><span className={`tag ${m.movement_type}`}>{m.movement_type === 'in' ? '入庫' : '出庫'}</span></td><td className="muted nowrap">{formatDate(m.movement_date)}</td><td className="code">{p?.code}</td><td className="product-name">{p?.name}</td><td className={`number ${m.movement_type === 'in' ? 'green-text' : 'red-text'}`}>{m.movement_type === 'in' ? '+' : '-'}{formatNumber(m.quantity)} <small>{p?.unit}</small></td><td className="number">{m.unit_price ? formatMoney(m.unit_price) : '—'}</td><td className="number muted">{m.unit_price ? formatMoney(m.quantity * m.unit_price) : '—'}</td><td>{m.supplier || '—'}</td><td className="muted note-cell">{m.notes || '—'}</td><td><div className="row-actions"><button className="danger-icon" onClick={() => onDelete(m)} aria-label={`刪除 ${formatDate(m.movement_date)} ${p?.name || ''} ${m.movement_type === 'in' ? '入庫' : '出庫'}紀錄`} title="刪除紀錄並重算庫存"><Trash2 /></button></div></td></tr>
    })}
  </tbody></table></div>
}

const priceEntriesFor = (productId: string, movements: Movement[]) => movements
  .filter((movement) => movement.product_id === productId && movement.movement_type === 'in' && movement.unit_price !== null)
  .sort((a, b) => `${b.movement_date}-${b.created_at}`.localeCompare(`${a.movement_date}-${a.created_at}`))

function StocktakePage({ products, loading, add, onAdjust }: { products: Product[]; loading: boolean; add: () => void; onAdjust: (product: Product, actualQuantity: number, notes: string) => Promise<boolean> }) {
  const [query, setQuery] = useState('')
  const [category, setCategory] = useState('全部分類')
  const [drafts, setDrafts] = useState<Record<string, string>>({})
  const [notes, setNotes] = useState<Record<string, string>>({})
  const [savingId, setSavingId] = useState('')
  const [pageError, setPageError] = useState('')
  const categories = ['全部分類', ...PRODUCT_CATEGORIES]
  const filtered = products.filter((product) => `${product.code} ${product.name} ${product.brand} ${product.supplier}`.toLowerCase().includes(query.toLowerCase()) && (category === '全部分類' || product.category === category))
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
    <PageTitle eyebrow="STOCKTAKE" title="庫存盤點" actions={<button className="button primary" onClick={add}><Plus />新增盤點品項</button>} />
    <div className="stocktake-guide"><ClipboardCheck /><div><strong>直接填入現場實際數量</strong><p>不需要輸入價格。系統只以盤點差異調整庫存，既有單價會保留；沒有價格的刀具也能正常盤點。</p></div></div>
    <div className="stocktake-summary">
      <div><span>目前顯示</span><strong>{formatNumber(filtered.length)}</strong><small>個品項</small></div>
      <div><span>未設定單價</span><strong>{formatNumber(noPriceCount)}</strong><small>仍可盤點</small></div>
      <div className={pendingCount ? 'attention' : ''}><span>待儲存差異</span><strong>{formatNumber(pendingCount)}</strong><small>個品項</small></div>
    </div>
    <div className="filter-card product-filter"><SearchField value={query} setValue={setQuery} placeholder="搜尋編號、品名、廠牌、供應商"/><label className="select-wrap"><select value={category} onChange={(event) => setCategory(event.target.value)}>{categories.map((value) => <option key={value}>{value}</option>)}</select><ChevronDown /></label></div>
    {pageError && <div className="form-error stocktake-error">{pageError}</div>}
    {loading ? <PageSkeleton /> : <div className="table-card"><table className="data-table stocktake-table"><thead><tr><th>編號</th><th>品名</th><th>分類</th><th>單價狀態</th><th className="number">帳面庫存</th><th>實際盤點數量</th><th className="number">差異</th><th>更新</th><th>盤點備註</th></tr></thead><tbody>
      {filtered.length === 0 ? <EmptyRow colSpan={9} label="沒有符合條件的刀具"/> : filtered.map((product) => {
        const difference = differenceFor(product)
        const invalid = difference === null
        return <tr key={product.id} className={difference ? 'has-stocktake-difference' : ''}>
          <td className="code">{product.code}</td>
          <td className="product-name">{product.name}</td>
          <td><span className="tag neutral">{product.category}</span></td>
          <td>{product.unit_price > 0 ? <span className="known-price">{formatMoney(product.unit_price)}</span> : <span className="unknown-price">未設定價格</span>}</td>
          <td className="number quantity">{formatNumber(product.stock_quantity)} <small>{product.unit}</small></td>
          <td><label className={`stocktake-quantity ${invalid ? 'invalid' : ''}`}><input type="number" inputMode="numeric" pattern="[0-9]*" min="0" step="1" value={actualFor(product)} onFocus={(event) => event.currentTarget.select()} onChange={(event) => { setDrafts((current) => ({ ...current, [product.id]: event.target.value })); setPageError('') }} aria-label={`${product.name} 實際盤點數量`}/><span>{product.unit}</span></label></td>
          <td className={`number stocktake-difference ${difference === null ? 'invalid' : difference > 0 ? 'positive' : difference < 0 ? 'negative' : ''}`}>{difference === null ? '格式錯誤' : difference === 0 ? '一致' : `${difference > 0 ? '+' : ''}${formatNumber(difference)}`}</td>
          <td><button className="stocktake-save" onClick={() => void save(product)} disabled={savingId !== '' || invalid || difference === 0}><Save />{savingId === product.id ? '儲存中' : '更新數量'}</button></td>
          <td><input className="stocktake-note" value={notes[product.id] || ''} onChange={(event) => setNotes((current) => ({ ...current, [product.id]: event.target.value }))} placeholder="選填，例如盤點位置" /></td>
        </tr>
      })}
    </tbody></table></div>}
  </section>
}

function ProductsPage({ products, movements, loading, add, edit, openStock, viewHistory, onDelete }: { products: Product[]; movements: Movement[]; loading: boolean; add: () => void; edit: (p: Product) => void; openStock: (type: MovementType, productId: string) => void; viewHistory: (p: Product) => void; onDelete: (p: Product) => void }) {
  const [query, setQuery] = useState('')
  const [category, setCategory] = useState('全部分類')
  const categories = ['全部分類', ...PRODUCT_CATEGORIES]
  const filtered = products
    .filter((p) => `${p.code} ${p.name} ${p.brand} ${p.supplier}`.toLowerCase().includes(query.toLowerCase()) && (category === '全部分類' || p.category === category))
    .sort((a, b) => `${b.purchased_at || '0000-00-00'}-${b.updated_at}-${b.code}`.localeCompare(`${a.purchased_at || '0000-00-00'}-${a.updated_at}-${a.code}`))
  return <section className="page-content">
    <PageTitle eyebrow="TOOLS" title="刀具資料" actions={<button className="button primary" onClick={add}><Plus />新增刀具</button>} />
    <div className="filter-card product-filter"><SearchField value={query} setValue={setQuery} placeholder="搜尋編號、品名、廠牌、供應商"/><label className="select-wrap"><select value={category} onChange={(e) => setCategory(e.target.value)}>{categories.map((c) => <option key={c}>{c}</option>)}</select><ChevronDown /></label></div>
    {loading ? <PageSkeleton /> : <div className="table-card"><table className="data-table product-table"><thead><tr><th>編號</th><th>品名</th><th>廠牌</th><th>供應商</th><th>分類</th><th className="number">數量</th><th className="number">單價</th><th>購入日期</th><th aria-label="操作" /></tr></thead><tbody>
      {filtered.length === 0 ? <EmptyRow colSpan={9} label="沒有符合條件的刀具"/> : filtered.map((p) => {
        const prices = priceEntriesFor(p.id, movements)
        const latest = prices[0]?.unit_price || p.unit_price
        const previous = prices[1]?.unit_price
        const percent = previous ? Math.round(((latest - previous) / previous) * 100) : null
        return <tr key={p.id}><td className="code">{p.code}</td><td className="product-name">{p.name}</td><td><strong>{p.brand || '—'}</strong></td><td className="muted">{p.supplier || '—'}</td><td><span className="tag neutral">{p.category}</span></td><td className="number quantity">{formatNumber(p.stock_quantity)} <small>{p.unit}</small></td><td className="number price-cell"><span>{formatMoney(p.unit_price)}</span>{percent !== null && percent !== 0 && <small className={percent > 0 ? 'price-up' : 'price-down'}>{percent > 0 ? '▲' : '▼'}{Math.abs(percent)}%</small>}</td><td className="muted nowrap">{p.purchased_at ? formatDate(p.purchased_at) : '—'}</td><td><div className="row-actions product-actions"><button className="in-icon" onClick={() => openStock('in', p.id)} aria-label={`${p.name} 入庫`} title="入庫"><ArrowDownToLine /></button><button className="out-icon" onClick={() => openStock('out', p.id)} aria-label={`${p.name} 出庫`} title="出庫"><ArrowUpFromLine /></button><button className="history-icon" onClick={() => viewHistory(p)} aria-label={`${p.name} 歷史價格`} title="歷史價格"><History /></button><button onClick={() => edit(p)} aria-label={`編輯 ${p.name}`} title="編輯"><Pencil /></button><button className="danger-icon" onClick={() => onDelete(p)} aria-label={`刪除 ${p.name}`} title="刪除"><Trash2 /></button></div></td></tr>
      })}
    </tbody></table></div>}
  </section>
}

function SearchField({ value, setValue, placeholder }: { value: string; setValue: (v: string) => void; placeholder: string }) {
  return <label className="search-field"><Search /><input value={value} onChange={(e) => setValue(e.target.value)} placeholder={placeholder}/>{value && <button type="button" onClick={() => setValue('')} aria-label="清除搜尋"><X /></button>}</label>
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
    <form onSubmit={submit} onKeyDown={preventInputEnterSubmit}>
      <div className="segmented"><button type="button" className={type === 'in' ? 'selected-in' : ''} onClick={() => switchType('in')}>入庫</button><button type="button" className={type === 'out' ? 'selected-out' : ''} onClick={() => switchType('out')}>出庫</button></div>
      <Field label="刀具" wide><ProductPicker products={products} value={productId} onChange={setProductId}/></Field>
      {selected && <div className="stock-hint">目前庫存 <strong>{formatNumber(selected.stock_quantity)} {selected.unit}</strong></div>}
      <div className="form-grid"><Field label="數量"><input required min="1" step="1" type="number" inputMode="numeric" pattern="[0-9]*" value={quantity} onChange={(e) => setQuantity(e.target.value)}/></Field><Field label={type === 'in' ? '購入日期' : '出庫日期'}><input required type="date" value={date} onChange={(e) => setDate(e.target.value)}/></Field>
      {type === 'in' && <><Field label="單價"><input required min="0" step="0.01" type="number" inputMode="numeric" pattern="[0-9]*" value={price} onChange={(e) => setPrice(e.target.value)}/></Field><Field label="供應商"><EditableCombo options={supplierOptions} value={supplier} onChange={setSupplier} placeholder="選擇既有供應商或輸入新資料"/></Field></>}</div>
      {type === 'in' && <TotalPreview quantity={Number(quantity || 0)} unitPrice={Number(price || 0)} total={totalPrice}/>} 
      <Field label="備註" wide><textarea value={notes} onChange={(e) => setNotes(e.target.value)} placeholder={type === 'in' ? '例如：發票號碼' : '例如：機台編號、工單號碼'}/></Field>
      {formError && <p className="form-error">{formError}</p>}
      <button disabled={busy || products.length === 0} className={`button modal-submit ${type === 'in' ? 'inbound' : 'outbound'}`}>{busy ? '處理中…' : `確認${type === 'in' ? '入庫' : '出庫'}`}</button>
    </form>
  </Modal>
}

function ProductPicker({ products, value, onChange }: { products: Product[]; value: string; onChange: (value: string) => void }) {
  const selected = products.find((product) => product.id === value)
  const display = (product: Product) => `${product.code}｜${product.name}`
  const [query, setQuery] = useState(selected ? display(selected) : '')
  const [open, setOpen] = useState(false)
  const inputRef = useRef<HTMLInputElement>(null)
  useEffect(() => { const product = products.find((item) => item.id === value); if (product) setQuery(display(product)) }, [value, products])
  const normalized = query.trim().toLowerCase()
  const filtered = products.filter((product) => `${product.code} ${product.name} ${product.brand} ${product.supplier} ${product.category}`.toLowerCase().includes(normalized)).slice(0, 10)
  const choose = (product: Product) => { onChange(product.id); setQuery(display(product)); setOpen(false); inputRef.current?.blur() }
  return <div className="product-picker" onBlur={() => window.setTimeout(() => setOpen(false), 120)}>
    <div className="product-picker-input"><Search/><input ref={inputRef} required role="combobox" aria-expanded={open} aria-controls="product-picker-options" value={query} onFocus={() => setOpen(true)} onClick={() => setOpen(true)} onChange={(event) => { setQuery(event.target.value); onChange(''); setOpen(true) }} placeholder="輸入編號、品名、廠牌或供應商搜尋"/><ChevronDown/></div>
    {open && <div className="product-picker-menu" id="product-picker-options" role="listbox">
      {filtered.length === 0 ? <div className="product-picker-empty">找不到符合的刀具</div> : filtered.map((product) => <button type="button" role="option" aria-selected={product.id === value} key={product.id} onMouseDown={(event) => event.preventDefault()} onClick={() => choose(product)}><span><b>{product.code}</b>{product.name}</span><small>庫存 {formatNumber(product.stock_quantity)} {product.unit}</small></button>)}
    </div>}
  </div>
}

function TotalPreview({ quantity, unitPrice, total }: { quantity: number; unitPrice: number; total: number }) {
  return <div className="total-preview"><span>採購總價</span><strong>${formatMoney(total)}</strong><small>{formatNumber(quantity)} × ${formatMoney(unitPrice)}</small></div>
}

function ProductModal({ product, products, close, save }: { product: Product | null; products: Product[]; close: () => void; save: (input: ProductInput) => Promise<void> }) {
  const [data, setData] = useState<ProductInput>(() => product ? { specification: product.specification, size: product.size, chipbreaker: product.chipbreaker, material: product.material, brand: product.brand, supplier: product.supplier, category: product.category, unit: product.unit, notes: product.notes, initial_quantity: 0, initial_unit_price: 0, purchased_at: product.purchased_at || today() } : blankProduct())
  const [busy, setBusy] = useState(false)
  const [formError, setFormError] = useState('')
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
    if ((data.category === '刀片' || data.category === '鑽銑刀') && !data.size.trim()) { setFormError('此分類需要輸入尺寸'); return }
    if ((data.category === '刀片' || data.category === '鑽銑刀') && !data.material.trim()) { setFormError('此分類需要輸入材質'); return }
    setBusy(true)
    try { await save(data) } catch (err) { setFormError(messageFor(err)); setBusy(false) }
  }
  return <Modal close={close} title={product ? '編輯刀具' : '新增刀具'} size="large">
    <form onSubmit={submit} onKeyDown={preventInputEnterSubmit}>
      <div className="form-grid">
        <Field label="編號"><input disabled value={product?.code || '儲存後自動產生'} /></Field>
        <Field label="分類 *"><select required value={data.category} onChange={(e) => { changeCategory(e.target.value as ProductCategory); e.currentTarget.blur() }}>{PRODUCT_CATEGORIES.map((category) => <option key={category} value={category}>{category}</option>)}</select></Field>
      </div>
      <Field label="品名（自動組合）" wide><input className="composed-name" disabled value={name || (data.category === '刀片' ? '規格-尺寸-斷屑槽-材質' : data.category === '鑽銑刀' ? '規格-尺寸-材質' : '規格')} /></Field>
      <div className="form-grid dynamic-fields">
        <Field label="規格 *"><EditableCombo options={specificationOptions} required value={data.specification} onChange={(value) => set('specification', value.toUpperCase())} placeholder={data.category === '刀片' ? 'WNMG' : data.category === '鑽銑刀' ? 'EMM4100' : '選擇既有資料或輸入新規格'}/></Field>
        {(data.category === '刀片' || data.category === '鑽銑刀') && <Field label="尺寸 *"><EditableCombo options={sizeOptions} required value={data.size} onChange={(value) => set('size', value.toUpperCase())} placeholder={data.category === '刀片' ? '080408' : '例如 10'}/></Field>}
        {data.category === '刀片' && <Field label="斷屑槽（無則留空）"><EditableCombo options={chipbreakerOptions} value={data.chipbreaker} onChange={(value) => set('chipbreaker', value.toUpperCase())} placeholder="選擇既有資料或輸入新值"/></Field>}
        {(data.category === '刀片' || data.category === '鑽銑刀') && <Field label="材質 *"><EditableCombo options={materialOptions} required value={data.material} onChange={(value) => set('material', value.toUpperCase())} placeholder={data.category === '刀片' ? 'PR1535' : '例如 鎢鋼'}/></Field>}
      </div>
      <div className="form-grid common-fields">
        <Field label="廠牌"><EditableCombo options={brandOptions} value={data.brand} onChange={(value) => set('brand', value)} placeholder="選擇既有廠牌或輸入新資料"/></Field>
        <Field label="供應商"><EditableCombo options={supplierOptions} value={data.supplier} onChange={(value) => set('supplier', value)} placeholder="選擇既有供應商或輸入新資料"/></Field>
        <Field label="單位"><EditableCombo options={unitOptions} required value={data.unit} onChange={(value) => set('unit', value)} placeholder="片"/></Field>
      </div>
      {!product && <fieldset className="initial-stock"><legend>目前庫存（可只填數量，單價找不到時可留空）</legend><div className="form-grid three"><Field label="數量"><input min="0" step="1" type="number" inputMode="numeric" pattern="[0-9]*" value={data.initial_quantity || ''} onChange={(e) => set('initial_quantity', Number(e.target.value))}/></Field><Field label="單價"><input min="0" step="0.01" type="number" inputMode="numeric" pattern="[0-9]*" value={data.initial_unit_price || ''} onChange={(e) => set('initial_unit_price', Number(e.target.value))} placeholder="沒有價格可留空"/></Field><Field label="購入日期"><input type="date" value={data.purchased_at} onChange={(e) => set('purchased_at', e.target.value)}/></Field></div><TotalPreview quantity={data.initial_quantity} unitPrice={data.initial_unit_price} total={initialTotalPrice}/></fieldset>}
      <Field label="備註" wide><textarea value={data.notes} onChange={(e) => set('notes', e.target.value)} placeholder="材質、適用工件等"/></Field>
      {formError && <p className="form-error">{formError}</p>}
      <button disabled={busy} className="button modal-submit primary">{busy ? '儲存中…' : '儲存'}</button>
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
    if (!/^\d{4,}$/.test(password)) { setError('請輸入至少 4 位數字登入密碼'); return }
    setDeleting(true)
    try { await verifyOwnerPassword(password); await request.execute(); close(true) }
    catch (submitError) { setError(messageFor(submitError)) }
    finally { setDeleting(false) }
  }
  return <Modal title={request.title} close={() => !deleting && close(false)}>
    <div className="admin-delete-warning"><Trash2/><div><strong>此操作無法復原</strong><span>{request.description}</span></div></div>
    <form onSubmit={submit} onKeyDown={preventInputEnterSubmit}>
      <Field label="登入密碼" wide><input autoFocus type="password" inputMode="numeric" pattern="[0-9]{4,}" autoComplete="current-password" value={password} onChange={(event) => setPassword(event.target.value.replace(/\D/g, ''))} placeholder="至少 4 位數字" disabled={deleting}/></Field>
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
    if (!/^\d{4,}$/.test(password)) { setMessage('請先輸入至少 4 位數字登入密碼'); return }
    setBusy('export')
    try {
      await verifyOwnerPassword(password)
      const result = await exportAllAdminData()
      const { downloadAdminWorkbook } = await import('./lib/adminDataWorkbook')
      downloadAdminWorkbook(result.data, result.exported_at)
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
    if (!/^\d{4,}$/.test(password)) { setMessage('請輸入至少 4 位數字登入密碼'); return }
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
    <Field label="登入密碼"><input type="password" inputMode="numeric" pattern="[0-9]{4,}" autoComplete="current-password" value={password} onChange={(event) => setPassword(event.target.value.replace(/\D/g, ''))} placeholder="至少 4 位數字" disabled={Boolean(busy)}/></Field>

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

function Modal({ title, close, size = '', children }: { title: string; close: () => void; size?: string; children: ReactNode }) {
  useEffect(() => { const fn = (e: KeyboardEvent) => e.key === 'Escape' && close(); window.addEventListener('keydown', fn); return () => window.removeEventListener('keydown', fn) }, [close])
  const dismissInput = () => { if (document.activeElement instanceof HTMLElement) document.activeElement.blur() }
  const dismissFromBlankArea = (event: ReactPointerEvent<HTMLElement>) => {
    const target = event.target as HTMLElement
    if (!target.closest('input, select, textarea, button, a, label, [role="option"]')) dismissInput()
  }
  return <div className="modal-backdrop" role="presentation" onMouseDown={(e) => e.target === e.currentTarget && close()}><section className={`modal ${size}`} role="dialog" aria-modal="true" aria-label={title} onPointerDown={dismissFromBlankArea}><div className="modal-head"><h2>{title}</h2><button className="icon-button" onClick={close} aria-label="關閉"><X /></button></div><button type="button" className="modal-input-dismiss" onClick={dismissInput}><ChevronDown />完成輸入／收起選單</button>{children}</section></div>
}

function PageSkeleton() { return <div className="skeleton"><i/><i/><i/></div> }
