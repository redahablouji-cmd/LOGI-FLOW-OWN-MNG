import { useState, useEffect } from 'react';
import { useNavigate } from 'react-router-dom';
import { motion } from 'motion/react';
import { 
  Compass, LogOut, Loader2, Navigation, ShieldAlert, RefreshCw, BadgeCheck 
} from 'lucide-react';
import { useAuth } from '../hooks/useAuth';
import { getCurrentProfile, Company } from '../lib/auth';
import CreateStaffForm from '../components/manager/CreateStaffForm';
import { toast } from 'sonner';

export default function ManagerDashboard() {
  const navigate = useNavigate();
  const { user, profile, role, company_id, loading, signOut } = useAuth();
  
  const [activeCompany, setActiveCompany] = useState<Company | null>(null);
  const [loadingCompany, setLoadingCompany] = useState(true);

  // Security guard for manager role
  useEffect(() => {
    if (!loading) {
      if (!user) {
        navigate('/login');
      } else if (role && role !== 'manager') {
        // Redirect non-manager to correct page
        if (role === 'owner') navigate('/owner');
        else {
          alert("Accès refusé. Ce portail est réservé aux managers habilités.");
          signOut().then(() => navigate('/login'));
        }
      }
    }
  }, [user, role, loading, navigate, signOut]);

  // Fetch company details
  const fetchCompanyDetails = async () => {
    if (user) {
      try {
        setLoadingCompany(true);
        const { company } = await getCurrentProfile(user.id);
        setActiveCompany(company);
      } catch (err) {
        console.error('Error fetching company details:', err);
      } finally {
        setLoadingCompany(false);
      }
    }
  };

  useEffect(() => {
    fetchCompanyDetails();
  }, [user]);

  if (loading || loadingCompany || !user || role !== 'manager' || !activeCompany) {
    return (
      <div className="min-h-screen bg-slate-50 flex flex-col items-center justify-center font-sans">
        <Loader2 className="w-8 h-8 text-blue-600 animate-spin mb-3" />
        <span className="text-sm font-semibold text-slate-500">Chargement de votre espace manager...</span>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-slate-50 font-sans pb-12" id="manager-dashboard-page-root">
      
      {/* Top Professional Navbar consistent with design wireframe */}
      <nav className="bg-slate-900 border-b border-slate-800 sticky top-0 z-45 shadow-md px-4 sm:px-6 lg:px-8 text-white">
        <div className="max-w-7xl mx-auto flex items-center justify-between h-16">
          
          {/* Logo & Company Label */}
          <div className="flex items-center space-x-8">
            <div className="flex items-center space-x-2">
              <div className="bg-blue-600 h-8 w-8 rounded flex items-center justify-center">
                <span className="text-white font-bold text-lg">LF</span>
              </div>
              <span className="text-xl font-bold tracking-tight text-white">LOGI-FLOW</span>
            </div>
            <div className="h-6 w-px bg-slate-800 hidden sm:block"></div>
            <div className="hidden sm:flex items-center space-x-2">
              <span className="text-xs font-bold text-slate-400 uppercase tracking-wider">Entreprise :</span>
              <span className="text-xs font-bold text-white uppercase tracking-widest">{activeCompany.name}</span>
            </div>
          </div>

          {/* Right: Manager welcome & Logout */}
          <div className="flex items-center gap-4">
            <div className="text-right hidden sm:block">
              <span className="text-[10px] text-slate-400 block uppercase font-bold tracking-wider">Manager Connecté</span>
              <span className="text-sm font-bold text-white leading-tight">{profile?.full_name}</span>
            </div>
            
            <button
              onClick={() => {
                signOut().then(() => {
                  toast.success("Déconnexion réussie");
                  navigate('/login');
                });
              }}
              className="inline-flex items-center gap-2 px-3.5 py-1.5 text-xs font-bold uppercase tracking-wider rounded-lg text-rose-450 hover:text-rose-350 bg-rose-500/10 hover:bg-rose-500/20 transition-colors cursor-pointer border border-rose-500/10"
              id="manager-logout-btn"
            >
              <LogOut className="w-3.5 h-3.5" />
              <span className="hidden sm:inline">Déconnexion</span>
            </button>
          </div>

        </div>
      </nav>

      {/* Main Content Area */}
      <main className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 mt-8">
        
        {/* Header summary info banner */}
        <div className="mb-8 flex flex-col md:flex-row md:items-center justify-between gap-4 bg-slate-900 text-white rounded-xl p-6 shadow-xs border border-slate-800">
          <div>
            <span className="inline-flex items-center gap-1.5 px-2 py-0.5 rounded text-[10px] font-extrabold uppercase tracking-widest bg-emerald-500 text-slate-950 mb-2">
              <BadgeCheck className="w-3.5 h-3.5" /> Espace Gestionnaire
            </span>
            <h1 className="text-2xl sm:text-3xl font-extrabold tracking-tight">Supervision des Opérations du Staff</h1>
            <p className="text-sm text-slate-400 mt-1 max-w-xl">
              Ajoutez les comptes du personnel logistique (chauffeurs, dispatchers, mécaniciens, comptables) et gérez leurs statuts d'accès.
            </p>
          </div>
          <button 
            onClick={fetchCompanyDetails}
            className="self-start md:self-center bg-white/10 hover:bg-white/15 px-3 py-2 rounded-lg text-xs font-bold uppercase tracking-wider flex items-center gap-1.5 backdrop-blur-xs transition-all cursor-pointer border border-white/5"
          >
            <RefreshCw className="w-3.5 h-3.5" /> Actualiser
          </button>
        </div>

        {/* Manager Single-screen Central Staff Form component */}
        <div className="max-w-5xl mx-auto">
          <CreateStaffForm companyId={activeCompany.id} />
        </div>

      </main>

    </div>
  );
}
