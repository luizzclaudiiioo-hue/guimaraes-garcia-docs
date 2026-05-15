import { createClient } from '@supabase/supabase-js'

const SUPABASE_URL = 'https://ybewjqkaeyslbzujcpyz.supabase.co'
const SUPABASE_ANON_KEY = 'sb_publishable_WH1_aAVryPEcanp7UUgsKw_OWyLoRsc'

export const supabase = createClient(SUPABASE_URL, SUPABASE_ANON_KEY)
