import { useState, useEffect } from 'react';
import { useNavigate } from 'react-router-dom';
import { Loader2, LogOut, Users, ShoppingBag, Wrench, Menu, X, BadgeCheck, RefreshCw, Plus, Eye, Download, FileText, Pencil, Trash2 } from 'lucide-react';
import { useAuth } from '../hooks/useAuth';
import { supabase } from '../lib/supabase';
import { Company } from '../lib/auth';
import CreateStaffForm from '../components/manager/CreateStaffForm';
import { toast } from 'sonner';
import { motion, AnimatePresence } from 'motion/react';

// ── XLS Export ─────────────────────────────────────────────────────────────
const exportToXLS = (data: any[], filename: string) => {
  if (!data.length) return;
  const headers = Object.keys(data[0]);
  const rows    = data.map(row => headers.map(h => row[h] ?? '').join('\t'));
  const blob    = new Blob([[headers.join('\t'), ...rows].join('\n')], { type: 'application/vnd.ms-excel' });
  const a       = document.createElement('a');
  a.href        = URL.createObjectURL(blob);
  a.download    = `${filename}.xls`;
  a.click();
};

type ManagerTab = 'staff' | 'purchases' | 'fleetfix' | 'suivi';

// ── Types ──────────────────────────────────────────────────────────────────
interface Purchase {
  id: string; category: string; fournisseur: string; numero_facture: string;
  date_achat: string; montant_ht: number; tva_rate: number; tva_amount: number;
  montant_ttc: number; banque: string; mode_paiement: string; notes: string; created_at: string;
}
interface MaintenanceRecord {
  id: string; truck_plate: string; type: string; part_fixed: string; garage_name: string;
  total_cost: number; notes: string; receipt_url: string; date: string; mechanic_id: string; created_at: string;
}
interface MechanicProfile { id: string; full_name: string; employee_code: string; }
interface FundTopup { id: string; mechanic_id: string; amount: number; notes: string; created_at: string; }
interface SuiviRecord {
  id: string; date: string; matricule: string; type: string; facture: string;
  bon_commande: string; ot_bl_bs_be: string; client: string; depart: string; arrivee: string;
  manutention: number; immobilisation: number; prix_ht: number; prix_ttc: number;
  cout_revient: number; benefice: number; created_at: string;
}

const emptySuivi = {
  date: new Date().toISOString().split('T')[0],
  matricule: '', type: '', facture: '', bon_commande: '', ot_bl_bs_be: '',
  client: '', depart: '', arrivee: '',
  manutention: '', immobilisation: '', prix_ht: '', prix_ttc: '',
  cout_revient: '', benefice: '',
};

export default function ManagerDashboard() {
  const navigate = useNavigate();
  const { user, loading, signOut } = useAuth();

  const [activeCompany,    setActiveCompany]    = useState<Company | null>(null);
  const [managerName,      setManagerName]      = useState('');
  const [companyId,        setCompanyId]        = useState('');
  const [managerId,        setManagerId]        = useState('');
  const [loadingCompany,   setLoadingCompany]   = useState(true);
  const [activeTab,        setActiveTab]        = useState<ManagerTab>('staff');
  const [sidebarOpen,      setSidebarOpen]      = useState(false);

  // Purchases
  const [purchases,        setPurchases]        = useState<Purchase[]>([]);
  const [loadingPurchases, setLoadingPurchases] = useState(false);
  const [editingPurchase,  setEditingPurchase]  = useState<Purchase | null>(null);

  // FleetFix
  const [mechanics,        setMechanics]        = useState<MechanicProfile[]>([]);
  const [selectedMechanic, setSelectedMechanic] = useState<MechanicProfile | null>(null);
  const [maintenance,      setMaintenance]      = useState<MaintenanceRecord[]>([]);
  const [topups,           setTopups]           = useState<FundTopup[]>([]);
  const [loadingFleet,     setLoadingFleet]     = useState(false);
  const [topupAmount,      setTopupAmount]      = useState('');
  const [topupNotes,       setTopupNotes]       = useState('');
  const [viewingReceipt,   setViewingReceipt]   = useState<string | null>(null);
  const [editingRecord,    setEditingRecord]    = useState<MaintenanceRecord | null>(null);

  // Suivi Facturation
  const [suiviList,        setSuiviList]        = useState<SuiviRecord[]>([]);
  const [loadingSuivi,     setLoadingSuivi]     = useState(false);
  const [showSuiviForm,    setShowSuiviForm]    = useState(false);
  const [suiviForm,        setSuiviForm]        = useState(emptySuivi);
  const [editingSuivi,     setEditingSuivi]     = useState<SuiviRecord | null>(null);

  // ── Fetch company ──────────────────────────────────────────────────────
  const fetchCompany = async () => {
    if (!user) return;
    try {
      setLoadingCompany(true);
      const { data: profileData } = await supabase
        .from('staff_profiles').select('id, company_id, full_name, role')
        .eq('auth_user_id', user.id).maybeSingle();
      if (!profileData || profileData.role !== 'manager') {
        await signOut(); navigate('/login'); return;
      }
      setManagerName(profileData.full_name);
      setCompanyId(profileData.company_id);
      setManagerId(profileData.id);
      const { data: companyData } = await supabase
        .from('companies').select('*').eq('id', profileData.company_id).maybeSingle();
      setActiveCompany(companyData);
    } catch { toast.error("Erreur lors du chargement."); }
    finally { setLoadingCompany(false); }
  };

  // ── Fetch purchases ────────────────────────────────────────────────────
  const fetchPurchases = async () => {
    setLoadingPurchases(true);
    const { data } = await supabase.from('purchases').select('*').order('created_at', { ascending: false });
    setPurchases(data || []);
    setLoadingPurchases(false);
  };

  // ── Fetch FleetFix ─────────────────────────────────────────────────────
  const fetchMechanics = async () => {
    if (!companyId) return;
    const { data } = await supabase.from('staff_profiles')
      .select('id, full_name, employee_code').eq('company_id', companyId).eq('role', 'mechanic');
    setMechanics(data || []);
  };

  const fetchMechanicData = async (mechanicId: string) => {
    setLoadingFleet(true);
    const [{ data: recs }, { data: tops }] = await Promise.all([
      supabase.from('maintenance_records').select('*').eq('mechanic_id', mechanicId).order('created_at', { ascending: false }),
      supabase.from('fund_topups').select('*').eq('mechanic_id', mechanicId).order('created_at', { ascending: false }),
    ]);
    setMaintenance(recs || []);
    setTopups(tops || []);
    setLoadingFleet(false);
  };

  const mechanicBalance = () =>
    topups.reduce((s, t) => s + t.amount, 0) - maintenance.reduce((s, r) => s + r.total_cost, 0);

  const handleTopup = async () => {
    if (!selectedMechanic || !topupAmount) return;
    const amount = parseFloat(topupAmount);
    if (isNaN(amount) || amount <= 0) return;
    const { error } = await supabase.from('fund_topups').insert({
      company_id: companyId, mechanic_id: selectedMechanic.id,
      amount, notes: topupNotes || 'Fonds ajoutés par le manager',
    });
    if (!error) {
      toast.success(`${amount.toLocaleString('fr-MA')} MAD ajoutés`);
      setTopupAmount(''); setTopupNotes('');
      fetchMechanicData(selectedMechanic.id);
    } else toast.error(`Erreur: ${error.message}`);
  };

  // ── Fetch Suivi Facturation ────────────────────────────────────────────
  const fetchSuivi = async () => {
    setLoadingSuivi(true);
    const { data } = await supabase.from('suivi_facturation')
      .select('*').order('created_at', { ascending: false });
    setSuiviList(data || []);
    setLoadingSuivi(false);
  };

  const handleSaveSuivi = async () => {
    const payload = {
      company_id:    companyId    || null,
      manager_id:    managerId    || null,
      date:          suiviForm.date          || null,
      matricule:     suiviForm.matricule     || null,
      type:          suiviForm.type          || null,
      facture:       suiviForm.facture       || null,
      bon_commande:  suiviForm.bon_commande  || null,
      ot_bl_bs_be:   suiviForm.ot_bl_bs_be   || null,
      client:        suiviForm.client        || null,
      depart:        suiviForm.depart        || null,
      arrivee:       suiviForm.arrivee       || null,
      manutention:   parseFloat(suiviForm.manutention  as string) || 0,
      immobilisation:parseFloat(suiviForm.immobilisation as string) || 0,
      prix_ht:       parseFloat(suiviForm.prix_ht      as string) || 0,
      prix_ttc:      parseFloat(suiviForm.prix_ttc     as string) || 0,
      cout_revient:  parseFloat(suiviForm.cout_revient as string) || 0,
      benefice:      parseFloat(suiviForm.benefice     as string) || 0,
    };

    if (editingSuivi) {
      const { error } = await supabase.from('suivi_facturation').update(payload).eq('id', editingSuivi.id);
      if (!error) { toast.success("Enregistrement modifié."); setEditingSuivi(null); fetchSuivi(); }
      else toast.error(`Erreur: ${error.message}`);
    } else {
      const { error } = await supabase.from('suivi_facturation').insert(payload);
      if (!error) { toast.success("Enregistrement ajouté."); setShowSuiviForm(false); fetchSuivi(); }
      else toast.error(`Erreur: ${error.message}`);
    }
    setSuiviForm(emptySuivi);
  };

  const handleDeleteSuivi = async (id: string) => {
    if (!confirm('Supprimer cet enregistrement ?')) return;
    const { error } = await supabase.from('suivi_facturation').delete().eq('id', id);
    if (!error) { setSuiviList(prev => prev.filter(s => s.id !== id)); toast.success("Supprimé."); }
    else toast.error(`Erreur: ${error.message}`);
  };

  // ── Purchases edit/save ────────────────────────────────────────────────
  const handleSavePurchase = async () => {
    if (!editingPurchase) return;
    const { error } = await supabase.from('purchases').update({
      fournisseur: editingPurchase.fournisseur, numero_facture: editingPurchase.numero_facture,
      date_achat: editingPurchase.date_achat, montant_ht: editingPurchase.montant_ht,
      tva_amount: editingPurchase.tva_amount, montant_ttc: editingPurchase.montant_ttc,
      notes: editingPurchase.notes,
    }).eq('id', editingPurchase.id);
    if (!error) { toast.success("Modifié."); setEditingPurchase(null); fetchPurchases(); }
    else toast.error(`Erreur: ${error.message}`);
  };

  // ── Maintenance edit/save ──────────────────────────────────────────────
  const handleSaveRecord = async () => {
    if (!editingRecord) return;
    const { error } = await supabase.from('maintenance_records').update({
      truck_plate: editingRecord.truck_plate, part_fixed: editingRecord.part_fixed,
      garage_name: editingRecord.garage_name, total_cost: editingRecord.total_cost, notes: editingRecord.notes,
    }).eq('id', editingRecord.id);
    if (!error) { toast.success("Modifié."); setEditingRecord(null); if (selectedMechanic) fetchMechanicData(selectedMechanic.id); }
    else toast.error(`Erreur: ${error.message}`);
  };

  useEffect(() => {
    if (!loading) { if (!user) navigate('/login'); else fetchCompany(); }
  }, [user, loading]);

  useEffect(() => {
    if (activeTab === 'purchases') fetchPurchases();
    if (activeTab === 'fleetfix' && companyId) fetchMechanics();
    if (activeTab === 'suivi') fetchSuivi();
  }, [activeTab, companyId]);

  useEffect(() => { if (selectedMechanic) fetchMechanicData(selectedMechanic.id); }, [selectedMechanic]);

  if (loading || loadingCompany) {
    return (
      <div className="min-h-screen bg-slate-50 flex flex-col items-center justify-center">
        <Loader2 className="w-8 h-8 text-blue-600 animate-spin mb-3" />
        <span className="text-sm font-semibold text-slate-500">Chargement...</span>
      </div>
    );
  }

  if (!activeCompany) {
    return (
      <div className="min-h-screen bg-slate-50 flex flex-col items-center justify-center">
        <p className="text-slate-500 text-sm">Aucune entreprise trouvée.</p>
        <button onClick={() => navigate('/login')} className="mt-4 text-blue-600 text-sm font-bold">Retour</button>
      </div>
    );
  }

  const CATEGORY_LABELS: Record<string, string> = {
    insurance: 'Assurance', toll: 'Autoroute', accounting: 'Comptable',
    office_supplies: 'Fourniture', fuel: 'Carburant', repair: 'Réparation',
    tires: 'Pneumatique', spare_parts: 'Pièce rechange', telecom: 'Telecom',
    leasing: 'Leasing', transport: 'Transport',
  };

  const navItems = [
    { id: 'staff',     label: 'Staff',               icon: Users },
    { id: 'purchases', label: 'Achats & Factures',    icon: ShoppingBag },
    { id: 'fleetfix',  label: 'FleetFix',             icon: Wrench },
    { id: 'suivi',     label: 'Suivi Facturation',    icon: FileText },
  ] as const;

  // Suivi form fields mapping
  const suiviFields = [
    { label: 'Date',                key: 'date',           type: 'date'   },
    { label: 'Matricule',           key: 'matricule',      type: 'text'   },
    { label: 'Type',                key: 'type',           type: 'text'   },
    { label: 'Factures',            key: 'facture',        type: 'text'   },
    { label: 'N° Bon de Commande',  key: 'bon_commande',   type: 'text'   },
    { label: 'OT / BL-BS-BE-Booking', key: 'ot_bl_bs_be', type: 'text'   },
    { label: 'Clients',             key: 'client',         type: 'text'   },
    { label: 'Départ',              key: 'depart',         type: 'text'   },
    { label: 'Arrivée',             key: 'arrivee',        type: 'text'   },
    { label: 'Manutention (MAD)',   key: 'manutention',    type: 'number' },
    { label: 'Immobilisation (MAD)',key: 'immobilisation', type: 'number' },
    { label: 'Prix HT (MAD)',       key: 'prix_ht',        type: 'number' },
    { label: 'Prix TTC (MAD)',      key: 'prix_ttc',       type: 'number' },
    { label: 'Coût de Revient (MAD)', key: 'cout_revient', type: 'number' },
    { label: 'Bénéfice (MAD)',      key: 'benefice',       type: 'number' },
  ];

  return (
    <div className="min-h-screen bg-slate-50 font-sans flex flex-col">

      {/* Navbar */}
      <nav className="bg-slate-900 border-b border-slate-800 sticky top-0 z-50 shadow-md px-4 sm:px-6 lg:px-8 text-white">
        <div className="max-w-full mx-auto flex items-center justify-between h-16">
          <div className="flex items-center gap-4">
            <button onClick={() => setSidebarOpen(!sidebarOpen)}
              className="p-2 rounded-lg hover:bg-slate-800 transition-colors">
              {sidebarOpen ? <X size={20} /> : <Menu size={20} />}
            </button>
            <div className="flex items-center space-x-2">
              <div className="bg-blue-600 h-8 w-8 rounded flex items-center justify-center">
                <span className="text-white font-bold text-lg">LF</span>
              </div>
              <span className="text-xl font-bold tracking-tight text-white">LOGI-FLOW</span>
            </div>
            <div className="h-6 w-px bg-slate-800 hidden sm:block" />
            <span className="text-xs font-bold text-slate-400 uppercase tracking-wider hidden sm:block">
              {activeCompany.name}
            </span>
          </div>
          <div className="flex items-center gap-4">
            <div className="text-right hidden sm:block">
              <span className="text-[10px] text-slate-400 block uppercase font-bold tracking-wider">Manager</span>
              <span className="text-sm font-bold text-white">{managerName}</span>
            </div>
            <button
              onClick={() => signOut().then(() => { toast.success("Déconnexion réussie"); navigate('/login'); })}
              className="inline-flex items-center gap-2 px-3.5 py-1.5 text-xs font-bold uppercase tracking-wider rounded-lg text-rose-400 bg-rose-500/10 hover:bg-rose-500/20 transition-colors cursor-pointer border border-rose-500/10">
              <LogOut className="w-3.5 h-3.5" />
              <span className="hidden sm:inline">Déconnexion</span>
            </button>
          </div>
        </div>
      </nav>

      <div className="flex flex-1 overflow-hidden">

        {/* Sidebar */}
        <AnimatePresence>
          {sidebarOpen && (
            <>
              <motion.div initial={{ opacity: 0 }} animate={{ opacity: 1 }} exit={{ opacity: 0 }}
                className="fixed inset-0 bg-black/40 z-40 lg:hidden"
                onClick={() => setSidebarOpen(false)} />
              <motion.aside
                initial={{ x: -280 }} animate={{ x: 0 }} exit={{ x: -280 }}
                transition={{ type: 'spring', damping: 25, stiffness: 200 }}
                className="fixed left-0 top-16 bottom-0 w-64 bg-slate-900 z-50 flex flex-col border-r border-slate-800 shadow-xl">
                <div className="p-4 border-b border-slate-800">
                  <p className="text-[10px] font-black text-slate-500 uppercase tracking-widest">Navigation</p>
                </div>
                <nav className="flex-1 p-3 space-y-1">
                  {navItems.map(item => {
                    const Icon = item.icon;
                    const isActive = activeTab === item.id;
                    return (
                      <button key={item.id}
                        onClick={() => { setActiveTab(item.id); setSidebarOpen(false); }}
                        className={`w-full flex items-center gap-3 px-4 py-3 rounded-xl text-sm font-bold transition-all cursor-pointer ${
                          isActive ? 'bg-blue-600 text-white shadow-lg shadow-blue-600/20'
                            : 'text-slate-400 hover:bg-slate-800 hover:text-white'
                        }`}>
                        <Icon size={18} />
                        {item.label}
                      </button>
                    );
                  })}
                </nav>
                <div className="p-4 border-t border-slate-800">
                  <p className="text-[10px] text-slate-600 font-bold uppercase tracking-widest">Espace Gestionnaire</p>
                </div>
              </motion.aside>
            </>
          )}
        </AnimatePresence>

        {/* Main Content */}
        <main className="flex-1 overflow-auto p-4 sm:p-6 lg:p-8">

          {/* TAB: STAFF */}
          {activeTab === 'staff' && (
            <div>
              <div className="mb-6 flex flex-col md:flex-row md:items-center justify-between gap-4 bg-slate-900 text-white rounded-xl p-6 border border-slate-800">
                <div>
                  <span className="inline-flex items-center gap-1.5 px-2 py-0.5 rounded text-[10px] font-extrabold uppercase tracking-widest bg-emerald-500 text-slate-950 mb-2">
                    <BadgeCheck className="w-3.5 h-3.5" /> Espace Gestionnaire
                  </span>
                  <h1 className="text-2xl font-extrabold tracking-tight">Supervision du Staff</h1>
                  <p className="text-sm text-slate-400 mt-1">Créez et gérez les comptes du personnel logistique.</p>
                </div>
                <button onClick={fetchCompany} className="self-start bg-white/10 hover:bg-white/15 px-3 py-2 rounded-lg text-xs font-bold uppercase tracking-wider flex items-center gap-1.5 transition-all cursor-pointer">
                  <RefreshCw className="w-3.5 h-3.5" /> Actualiser
                </button>
              </div>
              <div className="max-w-5xl mx-auto">
                <CreateStaffForm companyId={activeCompany.id} />
              </div>
            </div>
          )}

          {/* TAB: PURCHASES */}
          {activeTab === 'purchases' && (
            <div>
              <div className="mb-6 bg-slate-900 text-white rounded-xl p-6 border border-slate-800">
                <div className="flex items-center justify-between flex-wrap gap-3">
                  <div>
                    <span className="inline-flex items-center gap-1.5 px-2 py-0.5 rounded text-[10px] font-extrabold uppercase tracking-widest bg-blue-500 text-white mb-2">
                      <ShoppingBag className="w-3.5 h-3.5" /> Achats & Facturation
                    </span>
                    <h1 className="text-2xl font-extrabold tracking-tight">Historique des Achats</h1>
                    <p className="text-sm text-slate-400 mt-1">{purchases.length} enregistrements</p>
                  </div>
                  <div className="flex gap-2">
                    <button onClick={fetchPurchases} className="bg-white/10 hover:bg-white/15 px-3 py-2 rounded-lg text-xs font-bold uppercase tracking-wider flex items-center gap-1.5 transition-all cursor-pointer">
                      <RefreshCw className="w-3.5 h-3.5" /> Actualiser
                    </button>
                    <button onClick={() => exportToXLS(purchases.map(p => ({
                      'Date': p.date_achat || '', 'Catégorie': CATEGORY_LABELS[p.category] || p.category,
                      'Fournisseur': p.fournisseur || '', 'N° Facture': p.numero_facture || '',
                      'Montant HT': p.montant_ht, 'TVA': p.tva_amount, 'Montant TTC': p.montant_ttc,
                    })), 'achats_historique')}
                      className="bg-emerald-600 hover:bg-emerald-700 text-white px-3 py-2 rounded-lg text-xs font-black uppercase tracking-wider flex items-center gap-1.5 cursor-pointer">
                      <Download size={14} /> Export XLS
                    </button>
                  </div>
                </div>
              </div>
              {loadingPurchases ? (
                <div className="flex items-center justify-center py-20"><Loader2 className="w-6 h-6 text-blue-600 animate-spin" /></div>
              ) : (
                <div className="bg-white rounded-xl border border-slate-200 overflow-hidden">
                  <div className="overflow-x-auto">
                    <table className="w-full text-left">
                      <thead className="bg-slate-50 border-b border-slate-200">
                        <tr>{['Date','Catégorie','Fournisseur','N° Facture','HT','TVA','TTC','Actions'].map(h => (
                          <th key={h} className="px-4 py-3 text-[10px] font-black text-slate-400 uppercase tracking-widest">{h}</th>
                        ))}</tr>
                      </thead>
                      <tbody className="divide-y divide-slate-100">
                        {purchases.length === 0 ? (
                          <tr><td colSpan={8} className="px-4 py-10 text-center text-sm text-slate-400">Aucun achat enregistré.</td></tr>
                        ) : purchases.map(p => (
                          <tr key={p.id} className="hover:bg-slate-50 transition-colors">
                            <td className="px-4 py-3 text-xs text-slate-700">{p.date_achat || '—'}</td>
                            <td className="px-4 py-3"><span className="text-[10px] font-black bg-slate-100 text-slate-600 px-2 py-0.5 rounded uppercase">{CATEGORY_LABELS[p.category] || p.category}</span></td>
                            <td className="px-4 py-3 text-xs font-semibold text-slate-700">{p.fournisseur || '—'}</td>
                            <td className="px-4 py-3 font-mono text-xs text-blue-600">{p.numero_facture || '—'}</td>
                            <td className="px-4 py-3 font-mono text-xs text-slate-700">{p.montant_ht?.toLocaleString('fr-MA', { minimumFractionDigits: 2 })}</td>
                            <td className="px-4 py-3 font-mono text-xs text-amber-700">{p.tva_amount?.toLocaleString('fr-MA', { minimumFractionDigits: 2 })}</td>
                            <td className="px-4 py-3 font-mono text-xs font-bold text-slate-900">{p.montant_ttc?.toLocaleString('fr-MA', { minimumFractionDigits: 2 })}</td>
                            <td className="px-4 py-3">
                              <button onClick={() => setEditingPurchase(p)} className="text-[10px] font-bold text-blue-600 hover:text-blue-700 uppercase tracking-wider">Modifier</button>
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>
              )}
            </div>
          )}

          {/* TAB: FLEETFIX */}
          {activeTab === 'fleetfix' && (
            <div>
              <div className="mb-6 bg-slate-900 text-white rounded-xl p-6 border border-slate-800">
                <div className="flex items-center justify-between flex-wrap gap-3">
                  <div>
                    <span className="inline-flex items-center gap-1.5 px-2 py-0.5 rounded text-[10px] font-extrabold uppercase tracking-widest bg-amber-500 text-slate-950 mb-2">
                      <Wrench className="w-3.5 h-3.5" /> FleetFix
                    </span>
                    <h1 className="text-2xl font-extrabold tracking-tight">Gestion des Mécaniciens</h1>
                    <p className="text-sm text-slate-400 mt-1">Fonds de travail et fiches d'entretien en temps réel.</p>
                  </div>
                  <div className="flex gap-2">
                    <button onClick={() => { fetchMechanics(); if (selectedMechanic) fetchMechanicData(selectedMechanic.id); }}
                      className="bg-white/10 hover:bg-white/15 px-3 py-2 rounded-lg text-xs font-bold uppercase tracking-wider flex items-center gap-1.5 cursor-pointer">
                      <RefreshCw className="w-3.5 h-3.5" /> Actualiser
                    </button>
                    <button onClick={() => exportToXLS(maintenance.map(r => ({
                      'Date': r.date, 'Camion': r.truck_plate, 'Catégorie': r.type,
                      'Pièce': r.part_fixed, 'Garage': r.garage_name, 'Coût MAD': r.total_cost,
                      'Mécanicien': selectedMechanic?.full_name || '',
                    })), `fleetfix_${selectedMechanic?.full_name || 'export'}`)}
                      disabled={!selectedMechanic || maintenance.length === 0}
                      className="bg-emerald-600 hover:bg-emerald-700 disabled:opacity-40 text-white px-3 py-2 rounded-lg text-xs font-black uppercase tracking-wider flex items-center gap-1.5 cursor-pointer">
                      <Download size={14} /> Export XLS
                    </button>
                  </div>
                </div>
              </div>

              <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
                <div className="bg-white rounded-xl border border-slate-200 p-4 space-y-2">
                  <p className="text-[10px] font-black text-slate-400 uppercase tracking-widest mb-3">Mécaniciens</p>
                  {mechanics.length === 0 ? (
                    <p className="text-sm text-slate-400 text-center py-4">Aucun mécanicien trouvé.</p>
                  ) : mechanics.map(m => (
                    <button key={m.id} onClick={() => setSelectedMechanic(m)}
                      className={`w-full text-left px-4 py-3 rounded-lg border transition-all cursor-pointer ${
                        selectedMechanic?.id === m.id ? 'bg-blue-600 border-blue-700 text-white' : 'border-slate-200 hover:border-blue-300 hover:bg-blue-50/40'
                      }`}>
                      <p className={`text-sm font-bold ${selectedMechanic?.id === m.id ? 'text-white' : 'text-slate-800'}`}>{m.full_name}</p>
                      <p className={`text-[10px] font-mono ${selectedMechanic?.id === m.id ? 'text-blue-200' : 'text-slate-400'}`}>{m.employee_code}</p>
                    </button>
                  ))}
                </div>

                <div className="lg:col-span-2 space-y-6">
                  {!selectedMechanic ? (
                    <div className="bg-white rounded-xl border border-slate-200 flex items-center justify-center h-48">
                      <p className="text-slate-400 text-sm">Sélectionnez un mécanicien</p>
                    </div>
                  ) : (
                    <>
                      <div className="bg-slate-900 text-white rounded-xl p-6">
                        <p className="text-[10px] font-black text-slate-400 uppercase tracking-widest mb-1">Solde — {selectedMechanic.full_name}</p>
                        <p className={`text-3xl font-black ${mechanicBalance() >= 0 ? 'text-emerald-400' : 'text-red-400'}`}>
                          {mechanicBalance().toLocaleString('fr-MA', { minimumFractionDigits: 2 })} MAD
                        </p>
                        <p className="text-[10px] text-slate-500 mt-1">
                          Entrées: {topups.reduce((s,t) => s+t.amount,0).toLocaleString('fr-MA')} MAD —
                          Sorties: {maintenance.reduce((s,r) => s+r.total_cost,0).toLocaleString('fr-MA')} MAD
                        </p>
                      </div>
                      <div className="bg-white rounded-xl border border-slate-200 p-5">
                        <p className="text-[10px] font-black text-slate-400 uppercase tracking-widest mb-3">Ajouter des Fonds</p>
                        <div className="flex gap-3">
                          <input type="number" placeholder="Montant (MAD)" value={topupAmount}
                            onChange={e => setTopupAmount(e.target.value)}
                            className="flex-1 h-10 rounded-lg border-2 border-slate-200 px-3 text-sm font-bold focus:outline-none focus:border-blue-500" />
                          <input type="text" placeholder="Note (optionnel)" value={topupNotes}
                            onChange={e => setTopupNotes(e.target.value)}
                            className="flex-1 h-10 rounded-lg border-2 border-slate-200 px-3 text-sm focus:outline-none focus:border-blue-500" />
                          <button onClick={handleTopup}
                            className="px-4 h-10 bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-black uppercase rounded-lg flex items-center gap-1.5 cursor-pointer">
                            <Plus size={14} /> Ajouter
                          </button>
                        </div>
                      </div>
                      <div className="bg-white rounded-xl border border-slate-200 overflow-hidden">
                        <div className="p-4 border-b border-slate-100 bg-slate-50">
                          <p className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Fiches d'Entretien — {maintenance.length}</p>
                        </div>
                        {loadingFleet ? (
                          <div className="flex items-center justify-center py-10"><Loader2 className="w-5 h-5 text-blue-600 animate-spin" /></div>
                        ) : maintenance.length === 0 ? (
                          <div className="py-10 text-center text-slate-400 text-sm">Aucune fiche.</div>
                        ) : (
                          <div className="overflow-x-auto">
                            <table className="w-full text-left">
                              <thead className="bg-slate-50 border-b border-slate-100">
                                <tr>{['Date','Camion','Catégorie','Pièce','Garage','Coût','Reçu','Actions'].map(h => (
                                  <th key={h} className="px-4 py-2 text-[9px] font-black text-slate-400 uppercase tracking-widest">{h}</th>
                                ))}</tr>
                              </thead>
                              <tbody className="divide-y divide-slate-50">
                                {maintenance.map(r => (
                                  <tr key={r.id} className="hover:bg-slate-50">
                                    <td className="px-4 py-3 text-xs text-slate-600">{r.date}</td>
                                    <td className="px-4 py-3 font-mono text-xs font-bold text-blue-600">{r.truck_plate}</td>
                                    <td className="px-4 py-3 text-xs text-slate-600">{r.type}</td>
                                    <td className="px-4 py-3 text-xs text-slate-700 font-semibold">{r.part_fixed}</td>
                                    <td className="px-4 py-3 text-xs text-slate-600">{r.garage_name}</td>
                                    <td className="px-4 py-3 font-mono text-xs font-bold text-red-600">-{r.total_cost?.toLocaleString('fr-MA', { minimumFractionDigits: 2 })}</td>
                                    <td className="px-4 py-3">
                                      {r.receipt_url ? (
                                        <button onClick={() => setViewingReceipt(r.receipt_url)} className="text-blue-600"><Eye size={16} /></button>
                                      ) : <span className="text-slate-300 text-xs">—</span>}
                                    </td>
                                    <td className="px-4 py-3">
                                      <button onClick={() => setEditingRecord(r)} className="text-[10px] font-bold text-blue-600 uppercase tracking-wider">Modifier</button>
                                    </td>
                                  </tr>
                                ))}
                              </tbody>
                            </table>
                          </div>
                        )}
                      </div>
                    </>
                  )}
                </div>
              </div>
            </div>
          )}

          {/* TAB: SUIVI FACTURATION */}
          {activeTab === 'suivi' && (
            <div>
              <div className="mb-6 bg-slate-900 text-white rounded-xl p-6 border border-slate-800">
                <div className="flex items-center justify-between flex-wrap gap-3">
                  <div>
                    <span className="inline-flex items-center gap-1.5 px-2 py-0.5 rounded text-[10px] font-extrabold uppercase tracking-widest bg-violet-500 text-white mb-2">
                      <FileText className="w-3.5 h-3.5" /> Suivi Facturation
                    </span>
                    <h1 className="text-2xl font-extrabold tracking-tight">Suivi des Prestations</h1>
                    <p className="text-sm text-slate-400 mt-1">{suiviList.length} enregistrements</p>
                  </div>
                  <div className="flex gap-2 flex-wrap">
                    <button onClick={fetchSuivi}
                      className="bg-white/10 hover:bg-white/15 px-3 py-2 rounded-lg text-xs font-bold uppercase tracking-wider flex items-center gap-1.5 cursor-pointer">
                      <RefreshCw className="w-3.5 h-3.5" /> Actualiser
                    </button>
                    <button onClick={() => exportToXLS(suiviList.map(s => ({
                      'Date': s.date, 'Matricule': s.matricule, 'Type': s.type,
                      'Factures': s.facture, 'N° Bon Commande': s.bon_commande,
                      'OT/BL-BS-BE': s.ot_bl_bs_be, 'Clients': s.client,
                      'Départ': s.depart, 'Arrivée': s.arrivee,
                      'Manutention': s.manutention, 'Immobilisation': s.immobilisation,
                      'Prix HT': s.prix_ht, 'Prix TTC': s.prix_ttc,
                      'Coût Revient': s.cout_revient, 'Bénéfice': s.benefice,
                    })), 'suivi_facturation')}
                      className="bg-emerald-600 hover:bg-emerald-700 text-white px-3 py-2 rounded-lg text-xs font-black uppercase tracking-wider flex items-center gap-1.5 cursor-pointer">
                      <Download size={14} /> Export XLS
                    </button>
                    <button onClick={() => { setSuiviForm(emptySuivi); setEditingSuivi(null); setShowSuiviForm(true); }}
                      className="bg-blue-600 hover:bg-blue-700 text-white px-3 py-2 rounded-lg text-xs font-black uppercase tracking-wider flex items-center gap-1.5 cursor-pointer">
                      <Plus size={14} /> Nouveau
                    </button>
                  </div>
                </div>
              </div>

              {loadingSuivi ? (
                <div className="flex items-center justify-center py-20"><Loader2 className="w-6 h-6 text-blue-600 animate-spin" /></div>
              ) : (
                <div className="bg-white rounded-xl border border-slate-200 overflow-hidden">
                  <div className="overflow-x-auto">
                    <table className="w-full text-left min-w-[1200px]">
                      <thead className="bg-slate-50 border-b border-slate-200">
                        <tr>
                          {['Date','Matricule','Type','Facture','N° Bon Cmd','OT/BL','Client','Départ','Arrivée','Manut.','Immob.','HT','TTC','Coût Rev.','Bénéfice','Actions'].map(h => (
                            <th key={h} className="px-3 py-3 text-[9px] font-black text-slate-400 uppercase tracking-widest whitespace-nowrap">{h}</th>
                          ))}
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-100">
                        {suiviList.length === 0 ? (
                          <tr><td colSpan={16} className="px-4 py-10 text-center text-sm text-slate-400">Aucune prestation enregistrée. Cliquez sur "Nouveau" pour commencer.</td></tr>
                        ) : suiviList.map(s => (
                          <tr key={s.id} className="hover:bg-slate-50 transition-colors">
                            <td className="px-3 py-3 text-xs text-slate-700 whitespace-nowrap">{s.date}</td>
                            <td className="px-3 py-3 font-mono text-xs font-bold text-blue-600">{s.matricule || '—'}</td>
                            <td className="px-3 py-3 text-xs text-slate-600">{s.type || '—'}</td>
                            <td className="px-3 py-3 font-mono text-xs text-slate-700">{s.facture || '—'}</td>
                            <td className="px-3 py-3 text-xs text-slate-600">{s.bon_commande || '—'}</td>
                            <td className="px-3 py-3 text-xs text-slate-600">{s.ot_bl_bs_be || '—'}</td>
                            <td className="px-3 py-3 text-xs font-semibold text-slate-700">{s.client || '—'}</td>
                            <td className="px-3 py-3 text-xs text-slate-600">{s.depart || '—'}</td>
                            <td className="px-3 py-3 text-xs text-slate-600">{s.arrivee || '—'}</td>
                            <td className="px-3 py-3 font-mono text-xs text-slate-700">{s.manutention?.toLocaleString('fr-MA')}</td>
                            <td className="px-3 py-3 font-mono text-xs text-slate-700">{s.immobilisation?.toLocaleString('fr-MA')}</td>
                            <td className="px-3 py-3 font-mono text-xs text-slate-700">{s.prix_ht?.toLocaleString('fr-MA')}</td>
                            <td className="px-3 py-3 font-mono text-xs font-bold text-slate-900">{s.prix_ttc?.toLocaleString('fr-MA')}</td>
                            <td className="px-3 py-3 font-mono text-xs text-amber-700">{s.cout_revient?.toLocaleString('fr-MA')}</td>
                            <td className="px-3 py-3 font-mono text-xs font-bold text-emerald-600">{s.benefice?.toLocaleString('fr-MA')}</td>
                            <td className="px-3 py-3">
                              <div className="flex items-center gap-1">
                                <button onClick={() => { setEditingSuivi(s); setSuiviForm({
                                  date: s.date, matricule: s.matricule, type: s.type,
                                  facture: s.facture, bon_commande: s.bon_commande, ot_bl_bs_be: s.ot_bl_bs_be,
                                  client: s.client, depart: s.depart, arrivee: s.arrivee,
                                  manutention: String(s.manutention), immobilisation: String(s.immobilisation),
                                  prix_ht: String(s.prix_ht), prix_ttc: String(s.prix_ttc),
                                  cout_revient: String(s.cout_revient), benefice: String(s.benefice),
                                }); setShowSuiviForm(true); }}
                                  className="p-1.5 rounded hover:bg-blue-50 text-slate-400 hover:text-blue-600 transition-colors">
                                  <Pencil size={13} />
                                </button>
                                <button onClick={() => handleDeleteSuivi(s.id)}
                                  className="p-1.5 rounded hover:bg-rose-50 text-slate-400 hover:text-rose-600 transition-colors">
                                  <Trash2 size={13} />
                                </button>
                              </div>
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>
              )}
            </div>
          )}
        </main>
      </div>

      {/* Suivi Form Modal */}
      <AnimatePresence>
        {showSuiviForm && (
          <div className="fixed inset-0 bg-black/50 z-50 flex items-center justify-center p-4">
            <motion.div initial={{ scale: 0.95, opacity: 0 }} animate={{ scale: 1, opacity: 1 }}
              exit={{ scale: 0.95, opacity: 0 }}
              className="bg-white rounded-xl p-6 max-w-2xl w-full shadow-xl max-h-[90vh] overflow-y-auto">
              <div className="flex items-center justify-between mb-5">
                <h3 className="font-black text-slate-900 uppercase tracking-widest text-sm">
                  {editingSuivi ? 'Modifier la Prestation' : 'Nouvelle Prestation'}
                </h3>
                <button onClick={() => { setShowSuiviForm(false); setEditingSuivi(null); setSuiviForm(emptySuivi); }}
                  className="p-1.5 rounded hover:bg-slate-100 text-slate-400">
                  <X size={16} />
                </button>
              </div>
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                {suiviFields.map(({ label, key, type }) => (
                  <div key={key}>
                    <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">{label}</label>
                    <input type={type} value={(suiviForm as any)[key] || ''}
                      onChange={e => setSuiviForm(prev => ({ ...prev, [key]: e.target.value }))}
                      className="w-full mt-1 h-9 rounded-lg border-2 border-slate-200 px-3 text-sm focus:outline-none focus:border-blue-500" />
                  </div>
                ))}
              </div>
              <div className="flex gap-3 pt-5">
                <button onClick={handleSaveSuivi}
                  className="flex-1 h-10 bg-blue-600 hover:bg-blue-700 text-white text-sm font-bold rounded-lg cursor-pointer">
                  {editingSuivi ? 'Enregistrer les modifications' : 'Ajouter la prestation'}
                </button>
                <button onClick={() => { setShowSuiviForm(false); setEditingSuivi(null); setSuiviForm(emptySuivi); }}
                  className="flex-1 h-10 bg-slate-100 hover:bg-slate-200 text-slate-700 text-sm font-bold rounded-lg cursor-pointer">
                  Annuler
                </button>
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>

      {/* Edit Purchase Modal */}
      <AnimatePresence>
        {editingPurchase && (
          <div className="fixed inset-0 bg-black/50 z-50 flex items-center justify-center p-4">
            <motion.div initial={{ scale: 0.95, opacity: 0 }} animate={{ scale: 1, opacity: 1 }} exit={{ scale: 0.95, opacity: 0 }}
              className="bg-white rounded-xl p-6 max-w-lg w-full shadow-xl space-y-4">
              <h3 className="font-black text-slate-900 uppercase tracking-widest text-sm">Modifier l'Achat</h3>
              {[
                { label: 'Fournisseur', key: 'fournisseur', type: 'text' },
                { label: 'N° Facture', key: 'numero_facture', type: 'text' },
                { label: 'Date', key: 'date_achat', type: 'date' },
                { label: 'Montant HT', key: 'montant_ht', type: 'number' },
                { label: 'Montant TTC', key: 'montant_ttc', type: 'number' },
                { label: 'Notes', key: 'notes', type: 'text' },
              ].map(({ label, key, type }) => (
                <div key={key}>
                  <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">{label}</label>
                  <input type={type} value={(editingPurchase as any)[key] || ''}
                    onChange={e => setEditingPurchase({ ...editingPurchase, [key]: type === 'number' ? parseFloat(e.target.value) : e.target.value })}
                    className="w-full mt-1 h-9 rounded-lg border-2 border-slate-200 px-3 text-sm focus:outline-none focus:border-blue-500" />
                </div>
              ))}
              <div className="flex gap-3 pt-2">
                <button onClick={handleSavePurchase} className="flex-1 h-10 bg-blue-600 hover:bg-blue-700 text-white text-sm font-bold rounded-lg cursor-pointer">Enregistrer</button>
                <button onClick={() => setEditingPurchase(null)} className="flex-1 h-10 bg-slate-100 hover:bg-slate-200 text-slate-700 text-sm font-bold rounded-lg cursor-pointer">Annuler</button>
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>

      {/* Edit Maintenance Modal */}
      <AnimatePresence>
        {editingRecord && (
          <div className="fixed inset-0 bg-black/50 z-50 flex items-center justify-center p-4">
            <motion.div initial={{ scale: 0.95, opacity: 0 }} animate={{ scale: 1, opacity: 1 }} exit={{ scale: 0.95, opacity: 0 }}
              className="bg-white rounded-xl p-6 max-w-lg w-full shadow-xl space-y-4">
              <h3 className="font-black text-slate-900 uppercase tracking-widest text-sm">Modifier la Fiche</h3>
              {[
                { label: 'Plaque', key: 'truck_plate', type: 'text' },
                { label: 'Pièce', key: 'part_fixed', type: 'text' },
                { label: 'Garage', key: 'garage_name', type: 'text' },
                { label: 'Coût MAD', key: 'total_cost', type: 'number' },
                { label: 'Notes', key: 'notes', type: 'text' },
              ].map(({ label, key, type }) => (
                <div key={key}>
                  <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">{label}</label>
                  <input type={type} value={(editingRecord as any)[key] || ''}
                    onChange={e => setEditingRecord({ ...editingRecord, [key]: type === 'number' ? parseFloat(e.target.value) : e.target.value })}
                    className="w-full mt-1 h-9 rounded-lg border-2 border-slate-200 px-3 text-sm focus:outline-none focus:border-blue-500" />
                </div>
              ))}
              <div className="flex gap-3 pt-2">
                <button onClick={handleSaveRecord} className="flex-1 h-10 bg-blue-600 hover:bg-blue-700 text-white text-sm font-bold rounded-lg cursor-pointer">Enregistrer</button>
                <button onClick={() => setEditingRecord(null)} className="flex-1 h-10 bg-slate-100 hover:bg-slate-200 text-slate-700 text-sm font-bold rounded-lg cursor-pointer">Annuler</button>
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>

      {/* Receipt Viewer */}
      <AnimatePresence>
        {viewingReceipt && (
          <div className="fixed inset-0 bg-black/80 z-50 flex items-center justify-center p-4" onClick={() => setViewingReceipt(null)}>
            <motion.div initial={{ scale: 0.9, opacity: 0 }} animate={{ scale: 1, opacity: 1 }} exit={{ scale: 0.9, opacity: 0 }}
              className="max-w-lg w-full" onClick={e => e.stopPropagation()}>
              <img src={viewingReceipt} alt="Reçu" className="w-full rounded-xl shadow-2xl" />
              <button onClick={() => setViewingReceipt(null)} className="mt-4 w-full py-3 bg-white text-slate-900 font-bold rounded-xl cursor-pointer">Fermer</button>
            </motion.div>
          </div>
        )}
      </AnimatePresence>
    </div>
  );
}