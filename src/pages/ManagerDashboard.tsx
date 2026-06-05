import { useState, useEffect } from 'react';
import { useNavigate } from 'react-router-dom';
import { Loader2, LogOut, RefreshCw, BadgeCheck } from 'lucide-react';
import { useAuth } from '../hooks/useAuth';
import { supabase, isSupabaseConfigured } from '../lib/supabase';
import { Company } from '../lib/auth';
import CreateStaffForm from '../components/manager/CreateStaffForm';
import { toast } from 'sonner';

export default function ManagerDashboard() {
  const navigate = useNavigate();
  const { user, loading, signOut } = useAuth();

  const [activeCompany, setActiveCompany]     = useState<Company | null>(null);
  const [managerName, setManagerName]         = useState<string>('');
  const [loadingCompany, setLoadingCompany]   = useState(true);

  const fetchCompany = async () => {
    if (!user) return;
    try {
      setLoadingCompany(true);

      if (isSupabaseConfigured) {
        // Manager is in staff_profiles — get their company_id from there
        const { data: profileData, error: profileErr } = await supabase
          .from('staff_profiles')
          .select('company_id, full_name, role')
          .eq('auth_user_id', user.id)
          .maybeSingle();

        if (profileErr) throw profileErr;

        if (!profileData) {
          toast.error("Aucun profil manager trouvé.");
          await signOut();
          navigate('/login');
          return;
        }

        if (profileData.role !== 'manager') {
          toast.error("Accès refusé. Ce portail est réservé aux managers.");
          await signOut();
          navigate('/login');
          return;
        }

        setManagerName(profileData.full_name);

        // Now fetch the company
        const { data: companyData, error: companyErr } = await supabase
          .from('companies')
          .select('*')
          .eq('id', profileData.company_id)
          .maybeSingle();

        if (companyErr) throw companyErr;
        setActiveCompany(companyData);

      } else {
        // Demo mode
        const staffRaw = localStorage.getItem('logiflow_mock_staff_profiles');
        const compRaw  = localStorage.getItem('logiflow_mock_companies');
        if (staffRaw && compRaw) {
          const profiles = JSON.parse(staffRaw);
          const companies = JSON.parse(compRaw);
          const profile = profiles.find((p: any) => p.auth_user_id === user.id && p.role === 'manager');
          if (profile) {
            setManagerName(profile.full_name);
            const company = companies.find((c: any) => c.id === profile.company_id);
            if (company) setActiveCompany(company);
          }
        }
      }
    } catch (err) {
      console.error('Error fetching manager company:', err);
      toast.error("Erreur lors du chargement.");
    } finally {
      setLoadingCompany(false);
    }
  };

  useEffect(() => {
    if (!loading) {
      if (!user) {
        navigate('/login');
      } else {
        fetchCompany();
      }
    }
  }, [user, loading]);

  if (loading || loadingCompany) {
    return (
      <div className="min-h-screen bg-slate-50 flex flex-col items-center justify-center">
        <Loader2 className="w-8 h-8 text-blue-600 animate-spin mb-3" />
        <span className="text-sm font-semibold text-slate-500">Chargement de votre espace manager...</span>
      </div>
    );
  }

  if (!activeCompany) {
    return (
      <div className="min-h-screen bg-slate-50 flex flex-col items-center justify-center">
        <p className="text-slate-500 text-sm">Aucune entreprise trouvée.</p>
        <button onClick={() => navigate('/login')} className="mt-4 text-blue-600 text-sm font-bold">
          Retour à la connexion
        </button>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-slate-50 font-sans pb-12">

      {/* Navbar */}
      <nav className="bg-slate-900 border-b border-slate-800 sticky top-0 z-45 shadow-md px-4 sm:px-6 lg:px-8 text-white">
        <div className="max-w-7xl mx-auto flex items-center justify-between h-16">
          <div className="flex items-center space-x-8">
            <div className="flex items-center space-x-2">
              <div className="bg-blue-600 h-8 w-8 rounded flex items-center justify-center">
                <span className="text-white font-bold text-lg">LF</span>
              </div>
              <span className="text-xl font-bold tracking-tight text-white">LOGI-FLOW</span>
            </div>
            <div className="h-6 w-px bg-slate-800 hidden sm:block" />
            <div className="hidden sm:flex items-center space-x-2">
              <span className="text-xs font-bold text-slate-400 uppercase tracking-wider">Entreprise :</span>
              <span className="text-xs font-bold text-white uppercase tracking-widest">{activeCompany.name}</span>
            </div>
          </div>

          <div className="flex items-center gap-4">
            <div className="text-right hidden sm:block">
              <span className="text-[10px] text-slate-400 block uppercase font-bold tracking-wider">Manager Connecté</span>
              <span className="text-sm font-bold text-white">{managerName}</span>
            </div>
            <button
              onClick={() => signOut().then(() => { toast.success("Déconnexion réussie"); navigate('/login'); })}
              className="inline-flex items-center gap-2 px-3.5 py-1.5 text-xs font-bold uppercase tracking-wider rounded-lg text-rose-400 bg-rose-500/10 hover:bg-rose-500/20 transition-colors cursor-pointer border border-rose-500/10"
            >
              <LogOut className="w-3.5 h-3.5" />
              <span className="hidden sm:inline">Déconnexion</span>
            </button>
          </div>
        </div>
      </nav>

      {/* Main */}
      <main className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 mt-8">
        <div className="mb-8 flex flex-col md:flex-row md:items-center justify-between gap-4 bg-slate-900 text-white rounded-xl p-6 border border-slate-800">
          <div>
            <span className="inline-flex items-center gap-1.5 px-2 py-0.5 rounded text-[10px] font-extrabold uppercase tracking-widest bg-emerald-500 text-slate-950 mb-2">
              <BadgeCheck className="w-3.5 h-3.5" /> Espace Gestionnaire
            </span>
            <h1 className="text-2xl sm:text-3xl font-extrabold tracking-tight">Supervision des Opérations du Staff</h1>
            <p className="text-sm text-slate-400 mt-1 max-w-xl">
              Ajoutez les comptes du personnel logistique et gérez leurs statuts d'accès.
            </p>
          </div>
          <button
            onClick={fetchCompany}
            className="self-start md:self-center bg-white/10 hover:bg-white/15 px-3 py-2 rounded-lg text-xs font-bold uppercase tracking-wider flex items-center gap-1.5 transition-all cursor-pointer border border-white/5"
          >
            <RefreshCw className="w-3.5 h-3.5" /> Actualiser
          </button>
        </div>

        <div className="max-w-5xl mx-auto">
          <CreateStaffForm companyId={activeCompany.id} />
        </div>
      </main>
    </div>
  );
}