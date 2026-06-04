import { useState, useEffect, useCallback } from 'react';
import { supabase, isSupabaseConfigured } from '../lib/supabase';
import { getSession, getCurrentProfile, signOut as authSignOut, StaffProfile } from '../lib/auth';

export interface UseAuthReturn {
  user: any | null;
  profile: StaffProfile | null;
  role: string | null;
  company_id: string | null;
  loading: boolean;
  signOut: () => Promise<void>;
  refreshAuth: () => Promise<void>;
}

export function useAuth(): UseAuthReturn {
  const [user, setUser] = useState<any | null>(null);
  const [profile, setProfile] = useState<StaffProfile | null>(null);
  const [loading, setLoading] = useState<boolean>(true);

  const fetchAuthData = useCallback(async () => {
    try {
      setLoading(true);
      const currentUser = await getSession();
      setUser(currentUser);

      if (currentUser) {
        const { profile: userProfile } = await getCurrentProfile(currentUser.id);
        setProfile(userProfile);
      } else {
        setProfile(null);
      }
    } catch (err) {
      console.error('Error fetching auth data in hook:', err);
      setUser(null);
      setProfile(null);
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => {
    // Initial fetch
    fetchAuthData();

    // Listen for auth events if Supabase is set up
    if (isSupabaseConfigured) {
      const { data: { subscription } } = supabase.auth.onAuthStateChange(async (event, session) => {
        if (event === 'SIGNED_IN' || event === 'TOKEN_REFRESHED') {
          const authUser = session?.user || null;
          setUser(authUser);
          if (authUser) {
            const { profile: userProfile } = await getCurrentProfile(authUser.id);
            setProfile(userProfile);
          }
        } else if (event === 'SIGNED_OUT') {
          setUser(null);
          setProfile(null);
        }
        setLoading(false);
      });

      return () => {
        subscription.unsubscribe();
      };
    } else {
      // In Demo/Mock mode, listen to window storage adjustments for seamless tab changes
      const handleStorageChange = () => {
        fetchAuthData();
      };
      window.addEventListener('storage', handleStorageChange);
      return () => {
        window.removeEventListener('storage', handleStorageChange);
      };
    }
  }, [fetchAuthData]);

  const signOut = useCallback(async () => {
    setLoading(true);
    await authSignOut();
    setUser(null);
    setProfile(null);
    setLoading(false);

    // Trigger local storage storage event for local updates
    if (!isSupabaseConfigured) {
      window.dispatchEvent(new Event('storage'));
    }
  }, []);

  return {
    user,
    profile,
    role: profile ? profile.role : null,
    company_id: profile ? profile.company_id : null,
    loading,
    signOut,
    refreshAuth: fetchAuthData
  };
}
