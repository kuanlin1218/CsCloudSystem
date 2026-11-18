import React from 'react'
import ReactDOM from 'react-dom/client'
import App from './App'
import './styles.css'

const root = document.getElementById('root')!

// 全站數字輸入框：滾輪捲動時先解除焦點，避免瀏覽器自動增減數值。
root.addEventListener('wheel', (event) => {
  const field = document.activeElement
  if (field instanceof HTMLInputElement && field.type === 'number' && event.target === field) {
    field.blur()
  }
}, { capture: true, passive: true })

ReactDOM.createRoot(root).render(
  <React.StrictMode><App /></React.StrictMode>,
)
