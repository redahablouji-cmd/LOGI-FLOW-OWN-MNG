/// <reference types="vite/client" />
import { createClient } from '@supabase/supabase-js';

const rawUrl      = import.meta.env.VITE_SUPABASE_URL as string | undefined;
const rawKey      = import.meta.env.VITE_SUPABASE_ANON_KEY as string | undefined;

const supabaseUrl = rawUrl
  ? rawUrl.replace(/\/rest\/v1\/?$/, '').replace(/\/$/, '')
  : '';

const supabaseAnonKey = rawKey ?? '';

export const isSupabaseConfigured =
  !!supabaseUrl &&
  !!supabaseAnonKey &&
  supabaseUrl !== 'https://your_project.supabase.co' &&
  supabaseUrl !== 'https://placeholder-project.supabase.co';

export const supabase = isSupabaseConfigured
  ? createClient(supabaseUrl, supabaseAnonKey)
  : createClient('https://placeholder-project.supabase.co', 'placeholder-anon-key');