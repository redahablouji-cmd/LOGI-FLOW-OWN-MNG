/// <reference types="vite/client" />
import { supabase, isSupabaseConfigured } from './supabase';
import { createClient } from '@supabase/supabase-js';

export { supabase, isSupabaseConfigured };

// Types for internal models
export interface Company {
  id: string;
  name: string;
  business_code: string;
  created_at: string;
}

export interface StaffProfile {
  id: string;
  company_id: string;
  full_name: string;
  role: string; // 'owner' | 'manager' | 'accountant' | 'dispatcher' | ...
  employee_code?: string;
  vehicle_plate?: string;
  auth_user_id: string;
  is_active: boolean;
  created_at: string;
}

// Low-profile LocalStorage Mock Database for Demo Mode
const getMockData = () => {
  const compRaw = localStorage.getItem('logiflow_mock_companies');
  const staffRaw = localStorage.getItem('logiflow_mock_staff_profiles');
  const sessionRaw = localStorage.getItem('logiflow_mock_session');

  const companies: Company[] = compRaw ? JSON.parse(compRaw) : [];
  const staff_profiles: StaffProfile[] = staffRaw ? JSON.parse(staffRaw) : [];
  const sessionUser = sessionRaw ? JSON.parse(sessionRaw) : null;

  // Pre-seed some demo data if empty so the preview starts with rich context
  if (companies.length === 0) {
    const demoCompanyId = 'd89b3f36-9a28-4c12-88d4-5fe5c86fb2e6';
    const demoOwnerAuthId = 'owner-auth-id-123';
    const demoManagerAuthId = 'manager-auth-id-456';

    const defaultCompany: Company = {
      id: demoCompanyId,
      name: 'LogiFlow Transports',
      business_code: 'LF-LOG-4821',
      created_at: new Date(Date.now() - 30 * 24 * 3600 * 1000).toISOString()
    };

    const defaultOwner: StaffProfile = {
      id: 'owner-prof-111',
      company_id: demoCompanyId,
      full_name: 'Jean Renaud',
      role: 'owner',
      employee_code: 'EMP-OWNER-01',
      auth_user_id: demoOwnerAuthId,
      is_active: true,
      created_at: defaultCompany.created_at
    };

    const defaultManager: StaffProfile = {
      id: 'manager-prof-222',
      company_id: demoCompanyId,
      full_name: 'Amélie Martin',
      role: 'manager',
      employee_code: 'EMP-MGR-02',
      auth_user_id: demoManagerAuthId,
      is_active: true,
      created_at: new Date(Date.now() - 15 * 24 * 3600 * 1000).toISOString()
    };

    const defaultDriver: StaffProfile = {
      id: 'driver-prof-333',
      company_id: demoCompanyId,
      full_name: 'Pierre Dubois',
      role: 'driver',
      employee_code: 'DRV-102',
      vehicle_plate: 'AA-123-BB',
      auth_user_id: 'driver-auth-999',
      is_active: true,
      created_at: new Date(Date.now() - 5 * 24 * 3600 * 1000).toISOString()
    };

    companies.push(defaultCompany);
    staff_profiles.push(defaultOwner, defaultManager, defaultDriver);

    localStorage.setItem('logiflow_mock_companies', JSON.stringify(companies));
    localStorage.setItem('logiflow_mock_staff_profiles', JSON.stringify(staff_profiles));
  }

  return { companies, staff_profiles, sessionUser };
};

const saveMockData = (companies: Company[], staff_profiles: StaffProfile[], sessionUser: any) => {
  localStorage.setItem('logiflow_mock_companies', JSON.stringify(companies));
  localStorage.setItem('logiflow_mock_staff_profiles', JSON.stringify(staff_profiles));
  if (sessionUser) {
    localStorage.setItem('logiflow_mock_session', JSON.stringify(sessionUser));
  } else {
    localStorage.removeItem('logiflow_mock_session');
  }
};

/**
 * 1. Sign In with email & password
 */
export async function signIn(email: string, password: string): Promise<{ user: any; error: any }> {
  if (isSupabaseConfigured) {
    try {
      const { data, error } = await supabase.auth.signInWithPassword({
        email,
        password,
      });
      if (error) throw error;
      return { user: data.user, error: null };
    } catch (err: any) {
      return { user: null, error: err };
    }
  } else {
    // Mock Mode
    const { staff_profiles } = getMockData();
    // In mock mode, check if email corresponds to one of our personas or standard emails
    const cleanEmail = email.toLowerCase().trim();
    
    let simulatedUser = null;
    let foundProfile = null;

    if (cleanEmail === 'owner@logiflow.com' || cleanEmail === 'redahablouji@gmail.com') {
      foundProfile = staff_profiles.find(p => p.role === 'owner');
    } else if (cleanEmail === 'manager@logiflow.com') {
      foundProfile = staff_profiles.find(p => p.role === 'manager');
    } else {
      // Find by matching email as simulated auth_user_id (in mock we store email or auth ID)
      foundProfile = staff_profiles.find(p => p.auth_user_id === cleanEmail || p.employee_code?.toLowerCase() === cleanEmail);
    }

    if (!foundProfile && cleanEmail.includes('@')) {
      // Create on the fly to simulate seamless testing
      return { user: null, error: new Error("Identifiants incorrects ou utilisateur non trouvé en mode Démo.") };
    }

    if (foundProfile) {
      simulatedUser = {
        id: foundProfile.auth_user_id,
        email: cleanEmail,
        user_metadata: { full_name: foundProfile.full_name }
      };
      saveMockData(getMockData().companies, staff_profiles, simulatedUser);
      return { user: simulatedUser, error: null };
    }

    return { user: null, error: new Error("Identifiants de démo non reconnus. Essayez 'owner@logiflow.com' ou 'manager@logiflow.com'") };
  }
}

/**
 * 2. Sign Out
 */
export async function signOut(): Promise<void> {
  if (isSupabaseConfigured) {
    await supabase.auth.signOut();
  } else {
    localStorage.removeItem('logiflow_mock_session');
  }
}

/**
 * 3. Get Current Session Auth User
 */
export async function getSession(): Promise<any | null> {
  if (isSupabaseConfigured) {
    const { data: { session } } = await supabase.auth.getSession();
    return session?.user || null;
  } else {
    const { sessionUser } = getMockData();
    return sessionUser;
  }
}

/**
 * 4. Create Staff Account (without disconnecting current user)
 * Uses a non-persisting secondary client in Supabase so the admin's session remains untouched.
 */
export async function createStaffAccount(email: string, password: string): Promise<{ user: any; error: any }> {
  if (isSupabaseConfigured) {
    try {
      const supabaseUrl = import.meta.env.VITE_SUPABASE_URL;
      const supabaseAnonKey = import.meta.env.VITE_SUPABASE_ANON_KEY;
      
      // Secondary client preventing session hijacking
      const secondaryClient = createClient(supabaseUrl, supabaseAnonKey, {
        auth: {
          persistSession: false,
          autoRefreshToken: false,
        }
      });
      
      const { data, error } = await secondaryClient.auth.signUp({
        email,
        password,
      });

      if (error) throw error;
      return { user: data.user, error: null };
    } catch (err: any) {
      return { user: null, error: err };
    }
  } else {
    // Mock Mode: Registering an auth session or email
    const simulatedUserId = email.toLowerCase().trim();
    return {
      user: { id: simulatedUserId, email: email },
      error: null
    };
  }
}

/**
 * 5. Get Profile in staff_profiles where auth_user_id = user.id
 */
export async function getCurrentProfile(auth_user_id: string): Promise<{ profile: StaffProfile | null; company: Company | null; error: any }> {
  if (isSupabaseConfigured) {
    try {
      // Fetch staff details
      const { data: profileData, error: profileErr } = await supabase
        .from('staff_profiles')
        .select('*')
        .eq('auth_user_id', auth_user_id)
        .maybeSingle();

      if (profileErr) throw profileErr;
      if (!profileData) {
        return { profile: null, company: null, error: null };
      }

      // Fetch company details
      const { data: companyData, error: companyErr } = await supabase
        .from('companies')
        .select('*')
        .eq('id', profileData.company_id)
        .maybeSingle();

      if (companyErr) throw companyErr;

      return { profile: profileData, company: companyData, error: null };
    } catch (err: any) {
      return { profile: null, company: null, error: err };
    }
  } else {
    // Mock Mode
    const { staff_profiles, companies } = getMockData();
    const profile = staff_profiles.find(p => p.auth_user_id === auth_user_id) || null;
    let company = null;

    if (profile) {
      company = companies.find(c => c.id === profile.company_id) || null;
    }

    return { profile, company, error: null };
  }
}

// -------------------------------------------------------------
// HELPER WRITERS FOR REGISTERING / DB INTERACTION
// -------------------------------------------------------------

/**
 * Create a brand new company and its owner profile
 */
export async function registerCompanyAndOwner(params: {
  companyName: string;
  ownerName: string;
  email: string;
  authUserId: string;
  businessCode: string;
}): Promise<{ company: Company | null; profile: StaffProfile | null; error: any }> {
  if (isSupabaseConfigured) {
    try {
      // 1. Insert Company
      const { data: companyData, error: companyErr } = await supabase
        .from('companies')
        .insert({
          name: params.companyName,
          business_code: params.businessCode,
        })
        .select()
        .single();

      if (companyErr) throw companyErr;

      // 2. Insert Staff Profile (Owner)
      const { data: profileData, error: profileErr } = await supabase
        .from('staff_profiles')
        .insert({
          company_id: companyData.id,
          full_name: params.ownerName,
          role: 'owner',
          auth_user_id: params.authUserId,
          is_active: true,
        })
        .select()
        .single();

      if (profileErr) throw profileErr;

      return { company: companyData, profile: profileData, error: null };
    } catch (err: any) {
      return { company: null, profile: null, error: err };
    }
  } else {
    // Mock Mode
    const { companies, staff_profiles } = getMockData();
    
    // Create new company
    const newCompany: Company = {
      id: `comp-id-${Math.random().toString(36).substr(2, 9)}`,
      name: params.companyName,
      business_code: params.businessCode,
      created_at: new Date().toISOString()
    };

    // Create owner profile
    const newOwner: StaffProfile = {
      id: `owner-prof-${Math.random().toString(36).substr(2, 9)}`,
      company_id: newCompany.id,
      full_name: params.ownerName,
      role: 'owner',
      employee_code: 'EMP-OWNER-NEW',
      auth_user_id: params.authUserId,
      is_active: true,
      created_at: new Date().toISOString()
    };

    companies.push(newCompany);
    staff_profiles.push(newOwner);

    // Save
    saveMockData(companies, staff_profiles, { id: params.authUserId, email: params.email, user_metadata: { full_name: params.ownerName } });

    return { company: newCompany, profile: newOwner, error: null };
  }
}

/**
 * Fetch all managers or staff for a company
 */
export async function getCompanyStaff(companyId: string, options?: { role?: string; excludeRoles?: string[] }): Promise<StaffProfile[]> {
  if (isSupabaseConfigured) {
    try {
      let query = supabase
        .from('staff_profiles')
        .select('*')
        .eq('company_id', companyId);

      if (options?.role) {
        query = query.eq('role', options.role);
      }
      if (options?.excludeRoles && options.excludeRoles.length > 0) {
        query = query.not('role', 'in', `(${options.excludeRoles.join(',')})`);
      }

      // Sort by creation date
      query = query.order('created_at', { ascending: false });

      const { data, error } = await query;
      if (error) throw error;
      return data || [];
    } catch (err) {
      console.error('Error fetching company staff:', err);
      return [];
    }
  } else {
    // Mock Mode
    const { staff_profiles } = getMockData();
    let filtered = staff_profiles.filter(p => p.company_id === companyId);

    if (options?.role) {
      filtered = filtered.filter(p => p.role === options.role);
    }
    if (options?.excludeRoles && options.excludeRoles.length > 0) {
      filtered = filtered.filter(p => !options.excludeRoles?.includes(p.role));
    }

    // Sort by created at descending
    return [...filtered].sort((a,b) => new Date(b.created_at).getTime() - new Date(a.created_at).getTime());
  }
}

/**
 * Toggle active status of a staff member
 */
export async function toggleStaffActive(profileId: string, currentStatus: boolean): Promise<boolean> {
  if (isSupabaseConfigured) {
    try {
      const { error } = await supabase
        .from('staff_profiles')
        .update({ is_active: !currentStatus })
        .eq('id', profileId);

      if (error) throw error;
      return !currentStatus;
    } catch (err) {
      console.error('Error toggling staff active status:', err);
      return currentStatus;
    }
  } else {
    // Mock Mode
    const { companies, staff_profiles, sessionUser } = getMockData();
    const idx = staff_profiles.findIndex(p => p.id === profileId);
    if (idx !== -1) {
      staff_profiles[idx].is_active = !currentStatus;
      saveMockData(companies, staff_profiles, sessionUser);
      return !currentStatus;
    }
    return currentStatus;
  }
}

/**
 * Create a staff / manager record in DB
 */
export async function insertStaffProfile(profile: Omit<StaffProfile, 'id' | 'created_at'>): Promise<{ profile: StaffProfile | null; error: any }> {
  if (isSupabaseConfigured) {
    try {
      const { data, error } = await supabase
        .from('staff_profiles')
        .insert({
          ...profile,
        })
        .select()
        .single();

      if (error) throw error;
      return { profile: data, error: null };
    } catch (err: any) {
      return { profile: null, error: err };
    }
  } else {
    // Mock Mode
    const { companies, staff_profiles, sessionUser } = getMockData();
    const newProf: StaffProfile = {
      id: `prof-${Math.random().toString(36).substr(2, 9)}`,
      ...profile,
      created_at: new Date().toISOString()
    };

    // Before inserting, if it's manager or staff, check if employee_code is unique within company
    if (profile.employee_code) {
      const isDuplicate = staff_profiles.some(p => p.company_id === profile.company_id && p.employee_code === profile.employee_code);
      if (isDuplicate) {
        return { profile: null, error: new Error(`Code employé '${profile.employee_code}' déjà utilisé dans cette entreprise.`) };
      }
    }

    staff_profiles.push(newProf);
    saveMockData(companies, staff_profiles, sessionUser);
    return { profile: newProf, error: null };
  }
}
