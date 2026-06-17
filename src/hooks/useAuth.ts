import { useState, useEffect, useCallback } from 'react';
import { supabase, isSupabaseConfigured } from '../lib/supabase';
import { signOut as authSignOut } from '../lib/auth';

export interface UseAuthReturn {
  user: any | null;
  profile: any | null;
  role: string | null;
  company_id: string | null;
  loading: boolean;
  signOut: () => Promise<void>;
  refreshAuth: () => Promise<void>;
}

export function useAuth(): UseAuthReturn {
  const [user, setUser]       = useState<any | null>(null);
  const [loading, setLoading] = useState<boolean>(true);

  const fetchUser = useCallback(async () => {
    try {
      if (isSupabaseConfigured) {
        const { data: { session } } = await supabase.auth.getSession();
        setUser(session?.user || null);
      } else {
        const raw = localStorage.getItem('logiflow_mock_session');
        setUser(raw ? JSON.parse(raw) : null);
      }
    } catch (err) {
      console.error('useAuth error:', err);
      setUser(null);
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => {
    fetchUser();

    if (isSupabaseConfigured) {
      const { data: { subscription } } = supabase.auth.onAuthStateChange((_event, session) => {
        setUser(session?.user || null);
        // Don't touch loading here — fetchUser already handles it once
      });
      return () => subscription.unsubscribe();
    }
  }, [fetchUser]);

  const signOut = useCallback(async () => {
    await authSignOut();
    setUser(null);
  }, []);

  const refreshAuth = useCallback(async () => {
    // Don't set loading=true on refresh — prevents flash redirects
    try {
      if (isSupabaseConfigured) {
        const { data: { session } } = await supabase.auth.getSession();
        setUser(session?.user || null);
      }
    } catch (err) {
      console.error('refreshAuth error:', err);
    }
  }, []);

  return {
    user,
    profile:    null,
    role:       null,
    company_id: null,
    loading,
    signOut,
    refreshAuth,
  };
}