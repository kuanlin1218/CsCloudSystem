import { createClient } from '@supabase/supabase-js'

const url = import.meta.env.VITE_SUPABASE_URL?.trim()
const key = import.meta.env.VITE_SUPABASE_ANON_KEY?.trim()

export const isDemoMode = !url || !key || url.includes('YOUR_PROJECT') || key.includes('YOUR_SUPABASE')
export const supabase = isDemoMode ? null : createClient(url!, key!, {
  auth: { persistSession: true, autoRefreshToken: true, detectSessionInUrl: true },
})
