import { useState, useEffect } from 'react';
import { useNavigate } from 'react-router-dom';
import { Loader2, LogOut, Copy, Check, ShieldCheck, RefreshCw, Users } from 'lucide-react';
import { supabase } from '../lib/supabase';
import { signOut as authSignOut } from '../lib/auth';
import { Company } from '../lib/auth';
import CompanyProfile from '../components/owner/CompanyProfile';
import CreateManagerForm from '../components/owner/CreateManagerForm';
import { toast } from 'sonner';

const ROLE_PORTALS: Record<string, { label: string; portal: string; color: string }> = {
  owner:                { label: 'Propriétaire',       portal: 'Owner Portal',                   color: 'bg-blue-100 text-blue-800' },
  manager:              { label: 'Manager',            portal: 'Manager Portal',                 color: 'bg-violet-100 text-violet-800' },
  accountant:           { label: 'Comptable',          portal: 'Caisse & Chauffeurs Portal',     color: 'bg-emerald-100 text-emerald-800' },
  dispatcher:           { label: 'Dispatcher',         portal: 'Caisse & Chauffeurs Portal',     color: 'bg-teal-100 text-teal-800' },
  mechanic:             { label: 'Mécanicien',         portal: 'FleetFix Mobile Portal',         color: 'bg-amber-100 text-amber-800' },
  purchase_responsible: { label: 'Responsable Achats', portal: 'Invoice & Purchase Portal',      color: 'bg-cyan-100 text-cyan-800' },
  driver:               { label: 'Chauffeur',          portal: 'Aucun portail (données flotte)', color: 'bg-slate-100 text-slate-600' },
};

export default function OwnerDashboard() {
  const navigate = useNavigate();
  const [activeCompany, setActiveCompany] = useState<Company | null>(null);
  const [loadingCompany, setLoadingCompany] = useState(true);
  const [copiedCode, setCopiedCode] = useState(false);
  const [ownerEmail, setOwnerEmail] = useState('');
  const [staffList, setStaffList] = useState<any[]>([]);
  const [loadingStaff, setLoadingStaff] = useState(false);

  const fetchStaff = async (compId: string) => {
    setLoadingStaff(true);
    const { data } = await supabase
      .from('staff_profiles')
      .select('*')
      .eq('company_id', compId)
      .order('created_at', { ascending: false });
    setStaffList(data || []);
    setLoadingStaff(false);
  };

  const fetchCompany = async () => {
    try {
      setLoadingCompany(true);
      let email = sessionStorage.getItem('owner_email');
      if (!email) {
        const { data: { session } } = await supabase.auth.getSession();
        email = session?.user?.email || null;
      }
      if (!email) {
        setLoadingCompany(false);
        return;
      }
      setOwnerEmail(email);
      const { data, error } = await supabase
        .from('companies')
        .select('*')
        .eq('owner_email', email)
        .maybeSingle();
      if (error) throw error;
      if (data) {
        setActiveCompany(data);
        fetchStaff(data.id);
      }
    } catch (err) {
      console.error('Error fetching company:', err);
      toast.error("Erreur lors du chargement de l'entreprise.");
    } finally {
      setLoadingCompany(false);
    }
  };

  useEffect(() => { fetchCompany(); }, []);

  const handleSignOut = async () => {
    sessionStorage.removeItem('owner_email');
    sessionStorage.removeItem('owner_session');
    try { await authSignOut(); } catch {}
    toast.success("Déconnexion réussie");
    navigate('/login');
  };

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

  if (loadingCompany) {
    return (
      <div className="min-h-screen bg-slate-50 flex flex-col items-center justify-center">
        <Loader2 className="w-8 h-8 text-blue-600 animate-spin mb-3" />
        <span className="text-sm font-semibold text-slate-500">Chargement de votre espace propriétaire...</span>
      </div>
    );
  }

  if (!activeCompany) {
    return (
      <div className="min-h-screen bg-slate-50 flex flex-col items-center justify-center gap-4">
        <div className="bg-white rounded-xl border border-slate-200 p-8 max-w-md text-center shadow-sm">
          <div className="bg-blue-600 h-10 w-10 rounded flex items-center justify-center mx-auto mb-4">
            <span className="text-white font-bold text-lg">LF</span>
          </div>
          <h2 className="text-lg font-extrabold text-slate-900 mb-2">Session Expirée</h2>
          <p className="text-slate-500 text-sm mb-6">Veuillez vous reconnecter avec votre code entreprise.</p>
          <button onClick={() => navigate('/login')}
            className="w-full py-3 bg-blue-600 hover:bg-blue-700 text-white font-bold rounded-lg text-sm cursor-pointer transition-colors">
            Retour à la connexion
          </button>
        </div>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-slate-50 font-sans pb-12">
      <nav className="bg-white border-b border-slate-200 sticky top-0 z-40 shadow-xs px-4 sm:px-6 lg:px-8">
        <div className="max-w-7xl mx-auto flex items-center justify-between h-16">
          <div className="flex items-center space-x-8">
            <div className="flex items-center space-x-2">
              <div className="bg-blue-600 h-8 w-8 rounded flex items-center justify-center">
                <span className="text-white font-bold text-lg">LF</span>
              </div>
              <span className="text-xl font-bold tracking-tight text-slate-900">LOGI-FLOW</span>
            </div>
            <div className="h-6 w-px bg-slate-200 hidden sm:block" />
            <div className="hidden sm:flex items-center space-x-2">
              <span className="text-xs font-bold text-slate-400 uppercase tracking-wider">Entreprise :</span>
              <span className="text-xs font-bold text-slate-900 uppercase tracking-widest">{activeCompany.name}</span>
            </div>
          </div>
          <div className="hidden md:flex items-center gap-2 px-3 py-1.5 rounded-lg bg-slate-50 border border-slate-100">
            <span className="text-[10px] font-bold text-slate-400 uppercase">Code Entreprise</span>
            <code className="text-xs font-mono font-bold text-slate-800 bg-slate-100 px-1.5 py-0.5 rounded">{activeCompany.business_code}</code>
            <button onClick={handleCopyCode} className="p-1 hover:bg-slate-200 rounded text-slate-500 transition-colors cursor-pointer">
              {copiedCode ? <Check className="w-3.5 h-3.5 text-emerald-600" /> : <Copy className="w-3.5 h-3.5" />}
            </button>
          </div>
          <div className="flex items-center gap-4">
            <div className="text-right hidden sm:block">
              <span className="text-xs text-slate-400 block">Bienvenue,</span>
              <span className="text-sm font-bold text-slate-700">{activeCompany.owner_name}</span>
            </div>
            <button onClick={handleSignOut}
              className="inline-flex items-center gap-2 px-3.5 py-1.5 text-xs font-bold uppercase tracking-wider rounded-lg text-rose-600 hover:text-rose-700 bg-rose-50 hover:bg-rose-100/70 transition-colors cursor-pointer">
              <LogOut className="w-3.5 h-3.5" />
              <span className="hidden sm:inline">Déconnexion</span>
            </button>
          </div>
        </div>
      </nav>

      <main className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 mt-8">
        <div className="mb-8 flex flex-col md:flex-row md:items-center justify-between gap-4 bg-slate-900 text-white rounded-xl p-6 shadow-xs">
          <div>
            <span className="inline-flex items-center gap-1.5 px-2 py-0.5 rounded text-[10px] font-extrabold uppercase tracking-widest bg-blue-600 text-white mb-2">
              <ShieldCheck className="w-3 h-3" /> Espace Propriétaire
            </span>
            <h1 className="text-2xl sm:text-3xl font-extrabold tracking-tight">Pilotez LOGI-FLOW ERP</h1>
            <p className="text-sm text-slate-400 mt-1 max-w-xl">Modifiez la structure de votre entreprise et recrutez des managers pour diviser le travail opérationnel.</p>
          </div>
          <button onClick={fetchCompany}
            className="self-start md:self-center bg-white/10 hover:bg-white/15 px-3 py-2 rounded-lg text-xs font-bold uppercase tracking-wider flex items-center gap-1.5 transition-all cursor-pointer">
            <RefreshCw className="w-3.5 h-3.5" /> Actualiser
          </button>
        </div>

        <div className="grid grid-cols-1 lg:grid-cols-12 gap-8 items-start">
          <div className="lg:col-span-5 space-y-6">
            <CompanyProfile
              company={activeCompany}
              ownerProfile={null}
              ownerEmail={ownerEmail}
              onCompanyUpdate={(updated) => { setActiveCompany(updated); toast.success("Nom de l'entreprise enregistré !"); }}
            />
          </div>
          <div className="lg:col-span-7">
            <CreateManagerForm companyId={activeCompany.id} />
          </div>
        </div>

        <div className="mt-8 bg-white rounded-xl border border-slate-200 shadow-xs overflow-hidden">
          <div className="px-6 py-4 bg-slate-50 border-b border-slate-200 flex items-center justify-between">
            <div className="flex items-center gap-3">
              <div className="w-10 h-10 rounded-lg bg-slate-900 flex items-center justify-center text-white">
                <Users size={20} />
              </div>
              <div>
                <h3 className="text-sm font-black text-slate-900 uppercase tracking-widest">Équipe Complète</h3>
                <p className="text-xs text-slate-500 mt-0.5">{staffList.length} membre(s) — accès par portail</p>
              </div>
            </div>
            <button onClick={() => activeCompany && fetchStaff(activeCompany.id)}
              className="bg-slate-100 hover:bg-slate-200 px-3 py-2 rounded-lg text-xs font-bold uppercase tracking-wider flex items-center gap-1.5 cursor-pointer text-slate-600">
              <RefreshCw size={14} /> Actualiser
            </button>
          </div>
          {loadingStaff ? (
            <div className="flex items-center justify-center py-12"><Loader2 className="w-5 h-5 text-blue-600 animate-spin" /></div>
          ) : staffList.length === 0 ? (
            <div className="py-12 text-center text-slate-400 text-sm">Aucun membre du staff trouvé.</div>
          ) : (
            <div className="overflow-x-auto">
              <table className="w-full text-left">
                <thead className="bg-slate-50 border-b border-slate-200">
                  <tr>
                    {['Nom Complet','Rôle','Portail d\'Accès','Code Employé','Plaque','Statut','Date Création'].map(h => (
                      <th key={h} className="px-4 py-3 text-[10px] font-black text-slate-400 uppercase tracking-widest">{h}</th>
                    ))}
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {staffList.map((s: any) => {
                    const info = ROLE_PORTALS[s.role] || { label: s.role, portal: '—', color: 'bg-slate-100 text-slate-600' };
                    return (
                      <tr key={s.id} className="hover:bg-slate-50 transition-colors">
                        <td className="px-4 py-3 text-sm font-semibold text-slate-800">{s.full_name}</td>
                        <td className="px-4 py-3">
                          <span className={`text-[10px] font-black px-2 py-0.5 rounded uppercase ${info.color}`}>{info.label}</span>
                        </td>
                        <td className="px-4 py-3 text-xs text-slate-600">{info.portal}</td>
                        <td className="px-4 py-3 font-mono text-xs text-slate-500">{s.employee_code || '—'}</td>
                        <td className="px-4 py-3 font-mono text-xs text-blue-600">{s.vehicle_plate || '—'}</td>
                        <td className="px-4 py-3">
                          <span className={`text-[10px] font-black px-2 py-0.5 rounded ${s.is_active ? 'bg-emerald-50 text-emerald-700' : 'bg-rose-50 text-rose-700'}`}>
                            {s.is_active ? 'Actif' : 'Inactif'}
                          </span>
                        </td>
                        <td className="px-4 py-3 text-xs text-slate-500">{new Date(s.created_at).toLocaleDateString('fr-MA')}</td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          )}
        </div>
      </main>
    </div>
  );
}