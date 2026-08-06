import { createClient } from '@supabase/supabase-js'

// Reemplazá con la URL que copiaste de Supabase (Paso 1.4)
const supabaseUrl = 'https://twfgzehevxqmlbonfhof.supabase.co' 

// Reemplazá con la clave anon/public que copiaste de Supabase
const supabaseAnonKey = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InR3Zmd6ZWhldnhxbWxib25maG9mIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODU4NzQ2ODQsImV4cCI6MjEwMTQ1MDY4NH0.Ippx84SqXmvUqdzCN45uqZRNdcRKGNRXFL1wYJoqd4U'


export const supabase = createClient(supabaseUrl, supabaseAnonKey)