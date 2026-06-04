import { useState, useEffect } from 'react';
import { useNavigate } from 'react-router-dom';
import { motion } from 'motion/react';
import { 
  Building2, LogOut, Copy, Check, ShieldCheck, Loader2, RefreshCw 
} from 'lucide-react';
import { useAuth } from '../hooks/useAuth';
import { getCurrentProfile, Company } from '../lib/auth';
import CompanyProfile from '../components/owner/CompanyProfile';
import CreateManagerForm from '../components/owner/CreateManagerForm';
import { toast } from 'sonner';

export default function OwnerDashboard() {
  const navigate = useNavigate();
  const { user, profile, role, company_id, loading, signOut, refreshAuth } = useAuth();
  
  const [activeCompany, setActiveCompany] = useState<Company | null>(null);
  const [loadingCompany, setLoadingCompany] = useState(true);
  const [copiedCode, setCopiedCode] = useState(false);

  // Security guard for owner role
  useEffect(() => {
    if (!loading) {
      if (!user) {
        navigate('/login');
      } else if (role && role !== 'owner') {
        // Redirect non-owner to correct page
        if (role === 'manager') navigate('/manager');
        else {
          alert("Accès refusé. Ce portail est réservé aux propriétaires.");
          signOut().then(() => navigate('/login'));
        }
      }
    }
  }, [user, role, loading, navigate, signOut]);

  // Fetch company details
  const fetchAndRegisterCompany = async () => {
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
    fetchAndRegisterCompany();
  }, [user]);

  const handleCopyCode = async () => {
    if (!activeCompany) return;
    try {
      await navigator.clipboard.writeText(activeCompany.business_code);
      setCopiedCode(true);
      setTimeout(() => setCopiedCode(false), 2000);
      toast.success("Code copié !");
    } catch (err) {
      console.error(err);
    }
  };

  if (loading || loadingCompany || !user || role !== 'owner' || !activeCompany) {
    return (
      <div className="min-h-screen bg-slate-50 flex flex-col items-center justify-center font-sans">
        <Loader2 className="w-8 h-8 text-blue-600 animate-spin mb-3" />
        <span className="text-sm font-semibold text-slate-500">Chargement de votre espace propriétaire...</span>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-slate-50 font-sans pb-12" id="owner-dashboard-page-root">
      
      {/* Top Professional Navbar */}
      <nav className="bg-white border-b border-slate-200 sticky top-0 z-45 shadow-xs px-4 sm:px-6 lg:px-8">
        <div className="max-w-7xl mx-auto flex items-center justify-between h-16">
          
          {/* Logo & Company Label: Matches the precise design specifications */}
          <div className="flex items-center space-x-8">
            <div className="flex items-center space-x-2">
              <div className="bg-blue-600 h-8 w-8 rounded flex items-center justify-center">
                <span className="text-white font-bold text-lg">LF</span>
              </div>
              <span className="text-xl font-bold tracking-tight text-slate-900">LOGI-FLOW</span>
            </div>
            <div className="h-6 w-px bg-slate-200 hidden sm:block"></div>
            <div className="hidden sm:flex items-center space-x-2">
              <span className="text-xs font-bold text-slate-400 uppercase tracking-wider">Entreprise :</span>
              <span className="text-xs font-bold text-slate-900 uppercase tracking-widest">{activeCompany.name}</span>
            </div>
          </div>

          {/* Center: Business code quick monitor with copy */}
          <div className="hidden md:flex items-center gap-2 px-3 py-1.5 rounded-lg bg-slate-50 border border-slate-100">
            <span className="text-[10px] font-bold text-slate-400 uppercase">Code Entreprise</span>
            <code className="text-xs font-mono font-bold text-slate-800 bg-slate-100 px-1.5 py-0.5 rounded">{activeCompany.business_code}</code>
            <button
              onClick={handleCopyCode}
              className="p-1 hover:bg-slate-200 rounded text-slate-500 hover:text-slate-700 transition-colors cursor-pointer"
              title="Copier le code"
              id="nav-copy-code-btn"
            >
              {copiedCode ? <Check className="w-3.5 h-3.5 text-emerald-600" /> : <Copy className="w-3.5 h-3.5" />}
            </button>
          </div>

          {/* Right: Owner welcome & Logout */}
          <div className="flex items-center gap-4">
            <div className="text-right hidden sm:block">
              <span className="text-xs text-slate-400 block">Bienvenue,</span>
              <span className="text-sm font-bold text-slate-700">{profile?.full_name}</span>
            </div>
            
            <button
              onClick={() => {
                signOut().then(() => {
                  toast.success("Déconnexion réussie");
                  navigate('/login');
                });
              }}
              className="inline-flex items-center gap-2 px-3.5 py-1.5 text-xs font-bold uppercase tracking-wider rounded-lg text-rose-600 hover:text-rose-700 bg-rose-50 hover:bg-rose-100/70 transition-colors cursor-pointer"
              id="owner-logout-btn"
            >
              <LogOut className="w-3.5 h-3.5" />
              <span className="hidden sm:inline">Déconnexion</span>
            </button>
          </div>

        </div>
      </nav>

      {/* Main Content Area */}
      <main className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 mt-8">
        
        {/* Header summary info banner with sleek Slate-900 / blue tones */}
        <div className="mb-8 flex flex-col md:flex-row md:items-center justify-between gap-4 bg-slate-900 border border-slate-850 text-white rounded-xl p-6 shadow-xs">
          <div>
            <span className="inline-flex items-center gap-1.5 px-2 py-0.5 rounded text-[10px] font-extrabold uppercase tracking-widest bg-blue-600 text-white mb-2">
              <ShieldCheck className="w-3 h-3" /> Espace Propriétaire
            </span>
            <h1 className="text-2xl sm:text-3xl font-extrabold tracking-tight">Pilotez LOGI-FLOW ERP</h1>
            <p className="text-sm text-slate-400 mt-1 max-w-xl">
              Modifiez la structure de votre entreprise d'un côté, et recrutez des responsables managers de l'autre pour diviser le travail opérationnel.
            </p>
          </div>
          <button 
            onClick={fetchAndRegisterCompany}
            className="self-start md:self-center bg-white/10 hover:bg-white/15 px-3 py-2 rounded-lg text-xs font-bold uppercase tracking-wider flex items-center gap-1.5 backdrop-blur-xs transition-all cursor-pointer"
          >
            <RefreshCw className="w-3.5 h-3.5" /> Actualiser
          </button>
        </div>

        {/* Dashboard Side-by-Side Bento Grid */}
        <div className="grid grid-cols-1 lg:grid-cols-12 gap-8 items-start">
          
          {/* LEFT: Company Profile */}
          <div className="lg:col-span-5 space-y-6">
            <CompanyProfile 
              company={activeCompany} 
              ownerProfile={profile} 
              ownerEmail={user.email || ''} 
              onCompanyUpdate={(updated) => {
                setActiveCompany(updated);
                toast.success("Nom de l'entreprise enregistré avec succès !");
              }}
            />
          </div>

          {/* RIGHT: CreateManagerForm */}
          <div className="lg:col-span-7">
            <CreateManagerForm companyId={activeCompany.id} />
          </div>

        </div>

      </main>

    </div>
  );
}
