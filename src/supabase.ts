import { createClient } from '@supabase/supabase-js'

const projectUrl = import.meta.env.VITE_SUPABASE_URL
const publishableKey = import.meta.env.VITE_SUPABASE_PUBLISHABLE_KEY

export const supabase: any = projectUrl && publishableKey
  ? createClient(projectUrl, publishableKey)
  : null
