import { useState, useEffect } from 'react';
import { useNavigate } from 'react-router-dom';
import { Loader2, LogOut, Users, ShoppingBag, Wrench, Menu, X, BadgeCheck, RefreshCw, Plus, Eye, Download, FileText, Pencil, Trash2, Truck, Upload, Receipt, Settings } from 'lucide-react';
import { useAuth } from '../hooks/useAuth';
import { supabase } from '../lib/supabase';
import { Company } from '../lib/auth';
import CreateStaffForm from '../components/manager/CreateStaffForm';
import * as XLSX from 'xlsx';
import { toast } from 'sonner';
import { motion, AnimatePresence } from 'motion/react';

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

type ManagerTab = 'staff' | 'purchases' | 'fleetfix' | 'suivi' | 'chauffeurs' | 'clients' | 'facturation' | 'settings';

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

const emptyFactForm = {
  date: new Date().toISOString().split('T')[0],
  numero_facture: '', client: '', depart: '', arrivee: '',
  montant_ht: '', tva: '', montant_ttc: '', tva_rate: '',
  bl_ot: '', bc: '', delai_paiement: '60',
  date_paiement: '', reglement_banque_type: '',
  reglement_numero: '', echeances: '', mode_paiement: '',
  statut: 'impayé',
};

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
  const [fleetDrivers,     setFleetDrivers]     = useState<any[]>([]);
  const [loadingDrivers,   setLoadingDrivers]   = useState(false);
  const [editingDriver,    setEditingDriver]     = useState<any | null>(null);
  const [driverEditForm,   setDriverEditForm]   = useState<any>({});
  const [uploadingXLS,     setUploadingXLS]     = useState(false);
  const [clientsList,      setClientsList]      = useState<any[]>([]);
  const [loadingClients,   setLoadingClients]   = useState(false);
  const [editingClient,    setEditingClient]    = useState<any | null>(null);
  const [clientEditForm,   setClientEditForm]   = useState<any>({});
  const [uploadingClients, setUploadingClients] = useState(false);

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
  const [editingTopup,     setEditingTopup]     = useState<FundTopup | null>(null);
  const [topupEditAmount,  setTopupEditAmount]  = useState('');
  const [topupEditNotes,   setTopupEditNotes]   = useState('');

  // Suivi Facturation
  const [suiviList,        setSuiviList]        = useState<SuiviRecord[]>([]);
  const [loadingSuivi,     setLoadingSuivi]     = useState(false);
  const [showSuiviForm,    setShowSuiviForm]    = useState(false);
  const [suiviForm,        setSuiviForm]        = useState(emptySuivi);
  const [editingSuivi,     setEditingSuivi]     = useState<SuiviRecord | null>(null);
// Suivi Facturation
const [facturationList,    setFacturationList]    = useState<any[]>([]);
const [loadingFacturation, setLoadingFacturation] = useState(false);
const [showFactForm,       setShowFactForm]       = useState(false);
const [editingFact,        setEditingFact]        = useState<any | null>(null);
const [factForm,           setFactForm]           = useState<any>({});
const [selectedFacts,      setSelectedFacts]      = useState<string[]>([]);
const [factFilter,         setFactFilter]         = useState({ client: '', dateFrom: '', dateTo: '', statut: '' });
const [uploadingFacts, setUploadingFacts] = useState(false);
const [invoiceSettings,  setInvoiceSettings]  = useState<any>({
  company_name: '', address: '', phone: '', email: '',
  ice: '', rc: '', rib: '', bank_name: '',
  invoice_title: 'FACTURE',
  primary_color: '#1e40af', accent_color: '#f59e0b',
  footer_text: '', signature_label: 'Signature & Cachet',
  show_bl_ot: true, show_bc: true, show_manut: false,
  show_immob: false, show_ecart: false,
  logo_url: '', logo_storage_path: '',
  skip_header: false, skip_footer: false,
  margin_top: 12, margin_bottom: 12,
  margin_left: 15, margin_right: 15,
  font_size: 11,
  col_width_date: 10, col_width_fact: 10,
  col_width_client: 15, col_width_route: 15,
  col_width_amounts: 12,
  rows_per_page: 15,
});
const [savingSettings,   setSavingSettings]   = useState(false);
const [settingsLoaded,   setSettingsLoaded]   = useState(false);
const [showPreview,      setShowPreview]      = useState(false);
const [uploadingLogo,    setUploadingLogo]    = useState(false);
const [logoPreviewUrl,   setLogoPreviewUrl]   = useState('');
const [prestationPickerOpen, setPrestationPickerOpen] = useState(false);
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

  // ── Fund topup ─────────────────────────────────────────────────────────
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

  const handleDeleteTopup = async (id: string) => {
    if (!confirm('Supprimer ce versement ?')) return;
    const { error } = await supabase.from('fund_topups').delete().eq('id', id);
    if (!error) {
      setTopups(prev => prev.filter(t => t.id !== id));
      toast.success("Versement supprimé.");
    } else toast.error(`Erreur: ${error.message}`);
  };

  const handleEditTopupSave = async () => {
    if (!editingTopup) return;
    const { error } = await supabase.from('fund_topups').update({
      amount: parseFloat(topupEditAmount) || 0,
      notes:  topupEditNotes,
    }).eq('id', editingTopup.id);
    if (!error) {
      toast.success("Versement modifié.");
      setEditingTopup(null);
      if (selectedMechanic) fetchMechanicData(selectedMechanic.id);
    } else toast.error(`Erreur: ${error.message}`);
  };

  // ── Purchases ──────────────────────────────────────────────────────────
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

  const handleDeletePurchase = async (id: string) => {
    if (!confirm('Supprimer cet achat ?')) return;
    const { error } = await supabase.from('purchases').delete().eq('id', id);
    if (!error) { setPurchases(prev => prev.filter(p => p.id !== id)); toast.success("Supprimé."); }
    else toast.error(`Erreur: ${error.message}`);
  };

  // ── Maintenance ────────────────────────────────────────────────────────
  const handleSaveRecord = async () => {
    if (!editingRecord) return;
    const { error } = await supabase.from('maintenance_records').update({
      truck_plate: editingRecord.truck_plate, part_fixed: editingRecord.part_fixed,
      garage_name: editingRecord.garage_name, total_cost: editingRecord.total_cost, notes: editingRecord.notes,
    }).eq('id', editingRecord.id);
    if (!error) { toast.success("Modifié."); setEditingRecord(null); if (selectedMechanic) fetchMechanicData(selectedMechanic.id); }
    else toast.error(`Erreur: ${error.message}`);
  };

  const handleDeleteRecord = async (id: string) => {
    if (!confirm('Supprimer cette fiche ?')) return;
    const { error } = await supabase.from('maintenance_records').delete().eq('id', id);
    if (!error) { setMaintenance(prev => prev.filter(r => r.id !== id)); toast.success("Supprimé."); }
    else toast.error(`Erreur: ${error.message}`);
  };

  // ── Suivi Facturation ──────────────────────────────────────────────────
  const fetchSuivi = async () => {
    setLoadingSuivi(true);
    const { data } = await supabase.from('suivi_prestation').select('*').order('created_at', { ascending: false });
    setSuiviList(data || []);
    setLoadingSuivi(false);
  };

  const handleSaveSuivi = async () => {
  const payload = {
    company_id:     companyId || null,
    manager_id:     managerId || null,
    date:           suiviForm.date           || null,
    matricule:      suiviForm.matricule      || null,
    type:           suiviForm.type           || null,
    facture:        suiviForm.facture        || null,
    bon_commande:   suiviForm.bon_commande   || null,
    ot_bl_bs_be:    suiviForm.ot_bl_bs_be    || null,
    client:         suiviForm.client         || null,
    depart:         suiviForm.depart         || null,
    arrivee:        suiviForm.arrivee        || null,
    manutention:    parseFloat(suiviForm.manutention    as string) || 0,
    immobilisation: parseFloat(suiviForm.immobilisation as string) || 0,
    prix_ht:        parseFloat(suiviForm.prix_ht        as string) || 0,
    prix_ttc:       parseFloat(suiviForm.prix_ttc       as string) || 0,
    cout_revient:   parseFloat(suiviForm.cout_revient   as string) || 0,
    benefice:       parseFloat(suiviForm.benefice       as string) || 0,
  };

  if (editingSuivi) {
    const { error } = await supabase
      .from('suivi_prestation')
      .update(payload)
      .eq('id', editingSuivi.id);

    if (!error) {
      toast.success("Prestation modifiée.");
      setEditingSuivi(null);

      // Sync shared fields to all linked factures
      const sharedFields = {
        client:      payload.client,
        depart:      payload.depart,
        arrivee:     payload.arrivee,
        montant_ht:  payload.prix_ht,
        montant_ttc: payload.prix_ttc,
        bl_ot:       payload.ot_bl_bs_be,
        bc:          payload.bon_commande,
      };
      await supabase
        .from('suivi_facturation')
        .update(sharedFields)
        .eq('prestation_id', editingSuivi.id);

      fetchSuivi();
      if (activeTab === 'facturation') fetchFacturation();
    } else toast.error(`Erreur: ${error.message}`);
  } else {
    const { error } = await supabase
      .from('suivi_prestation')
      .insert(payload);
    if (!error) {
      toast.success("Prestation ajoutée.");
      setShowSuiviForm(false);
      fetchSuivi();
    } else toast.error(`Erreur: ${error.message}`);
  }
  setSuiviForm(emptySuivi);
};

  const handleDeleteSuivi = async (id: string) => {
    if (!confirm('Supprimer ?')) return;
    const { error } = await supabase.from('suivi_prestation').delete().eq('id', id);
    if (!error) { setSuiviList(prev => prev.filter(s => s.id !== id)); toast.success("Supprimé."); }
    else toast.error(`Erreur: ${error.message}`);
  };
// ── Fleet Drivers ──────────────────────────────────────────────────────
const fetchFleetDrivers = async () => {
  if (!companyId) return;
  setLoadingDrivers(true);
  const { data } = await supabase
    .from('fleet_drivers')
    .select('*')
    .eq('company_id', companyId)
    .order('created_at', { ascending: false });
  setFleetDrivers(data || []);
  setLoadingDrivers(false);
};

const handleXLSUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
  const file = e.target.files?.[0];
  if (!file || !companyId) return;
  setUploadingXLS(true);
  try {
    const buffer = await file.arrayBuffer();
    const wb = XLSX.read(buffer, { type: 'array' });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const rows: any[] = XLSX.utils.sheet_to_json(ws, { header: 1 });

    const dataRows = rows.slice(1).filter((r: any[]) => r.length > 0 && r[1]);

    const records = dataRows.map((r: any[]) => ({
      company_id:          companyId,
      code:                String(r[0] || ''),
      nom_prenom:          String(r[1] || ''),
      immatriculation:     String(r[2] || ''),
      type_vehicule:       String(r[3] || ''),
      cin:                 String(r[4] || ''),
      imm_cnss:            String(r[5] || ''),
      fonction:            String(r[6] || ''),
      date_naissance:      String(r[7] || ''),
      situation_familiale: String(r[8] || ''),
      nb_deduction:        parseInt(r[9]) || 0,
      date_embauche:       String(r[10] || ''),
      adresse:             String(r[11] || ''),
    }));

    if (records.length === 0) {
      toast.error("Aucune donnée trouvée dans le fichier.");
      return;
    }

    await supabase.from('fleet_drivers').delete().eq('company_id', companyId);
    const { error } = await supabase.from('fleet_drivers').insert(records);

    if (!error) {
      toast.success(`${records.length} chauffeurs importés avec succès.`);
      fetchFleetDrivers();
    } else {
      toast.error(`Erreur: ${error.message}`);
    }
  } catch (err: any) {
    toast.error(`Erreur lecture fichier: ${err.message}`);
  } finally {
    setUploadingXLS(false);
    e.target.value = '';
  }
};

const handleDeleteDriver = async (id: string) => {
  if (!confirm('Supprimer ce chauffeur ?')) return;
  const { error } = await supabase.from('fleet_drivers').delete().eq('id', id);
  if (!error) { setFleetDrivers(prev => prev.filter(d => d.id !== id)); toast.success("Supprimé."); }
  else toast.error(`Erreur: ${error.message}`);
};

const handleEditDriverSave = async () => {
  const { error } = await supabase.from('fleet_drivers').update({
    code:                driverEditForm.code,
    nom_prenom:          driverEditForm.nom_prenom,
    immatriculation:     driverEditForm.immatriculation,
    type_vehicule:       driverEditForm.type_vehicule,
    cin:                 driverEditForm.cin,
    imm_cnss:            driverEditForm.imm_cnss,
    fonction:            driverEditForm.fonction,
    date_naissance:      driverEditForm.date_naissance,
    situation_familiale: driverEditForm.situation_familiale,
    nb_deduction:        parseInt(driverEditForm.nb_deduction) || 0,
    date_embauche:       driverEditForm.date_embauche,
    adresse:             driverEditForm.adresse,
  }).eq('id', editingDriver.id);
  if (!error) {
    toast.success("Chauffeur modifié.");
    setEditingDriver(null);
    fetchFleetDrivers();
  } else toast.error(`Erreur: ${error.message}`);
};
const fetchClients = async () => {
  if (!companyId) return;
  setLoadingClients(true);
  const { data } = await supabase
    .from('clients')
    .select('*')
    .eq('company_id', companyId)
    .order('created_at', { ascending: false });
  setClientsList(data || []);
  setLoadingClients(false);
};

const handleClientsXLSUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
  const file = e.target.files?.[0];
  if (!file || !companyId) return;
  setUploadingClients(true);
  try {
    const buffer = await file.arrayBuffer();
    const wb = XLSX.read(buffer, { type: 'array' });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const rows: any[] = XLSX.utils.sheet_to_json(ws, { header: 1 });
    const dataRows = rows.slice(1).filter((r: any[]) => r.length > 0 && r[0]);
    const records = dataRows.map((r: any[]) => ({
      company_id:    companyId,
      nom:           String(r[0] || '').trim(),
      adresse:       String(r[1] || '').trim(),
      ice:           String(r[2] || '').trim(),
      delai_paiement: parseInt(r[3]) || 60,
    }));
    if (records.length === 0) { toast.error("Aucune donnée trouvée."); return; }
    await supabase.from('clients').delete().eq('company_id', companyId);
    const { error } = await supabase.from('clients').insert(records);
    if (!error) { toast.success(`${records.length} clients importés.`); fetchClients(); }
    else toast.error(`Erreur: ${error.message}`);
  } catch (err: any) {
    toast.error(`Erreur: ${err.message}`);
  } finally {
    setUploadingClients(false);
    e.target.value = '';
  }
};

const handleDeleteClient = async (id: string) => {
  if (!confirm('Supprimer ce client ?')) return;
  const { error } = await supabase.from('clients').delete().eq('id', id);
  if (!error) { setClientsList(prev => prev.filter(c => c.id !== id)); toast.success("Supprimé."); }
  else toast.error(`Erreur: ${error.message}`);
};

const handleEditClientSave = async () => {
  const { error } = await supabase.from('clients').update({
    nom:           clientEditForm.nom,
    adresse:       clientEditForm.adresse,
    ice:           clientEditForm.ice,
    delai_paiement: parseInt(clientEditForm.delai_paiement) || 60,
  }).eq('id', editingClient.id);
  if (!error) { toast.success("Client modifié."); setEditingClient(null); fetchClients(); }
  else toast.error(`Erreur: ${error.message}`);
};
// ── Suivi Facturation ──────────────────────────────────────────────────

const calcEcartDelai = (dateFacture: string, datePaiement: string, delai: number): number => {
  if (!dateFacture || !delai) return 0;
  const base = new Date(dateFacture);
  const echeance = new Date(base);
  echeance.setDate(echeance.getDate() + delai);
  const paiement = datePaiement ? new Date(datePaiement) : new Date();
  return Math.round((paiement.getTime() - echeance.getTime()) / (1000 * 60 * 60 * 24));
};
const fetchFacturation = async () => {
  if (!companyId) return;
  setLoadingFacturation(true);
  const { data } = await supabase
    .from('suivi_facturation')
    .select('*, suivi_prestation(client, depart, arrivee, prix_ht, prix_ttc, ot_bl_bs_be, bon_commande)')
    .eq('company_id', companyId)
    .order('created_at', { ascending: false });
  const merged = (data || []).map((f: any) => ({
    ...f,
    client:  f.suivi_prestation?.client  || f.client,
    depart:  f.suivi_prestation?.depart  || f.depart,
    arrivee: f.suivi_prestation?.arrivee || f.arrivee,
  }));
  setFacturationList(merged);
  setLoadingFacturation(false);
};

const handleAutoFillFromPrestation = (prestation: any) => {
  const client = clientsList.find(c => c.nom === prestation.client);
  setFactForm({
    ...emptyFactForm,
    client:          prestation.client       || '',
    depart:          prestation.depart       || '',
    arrivee:         prestation.arrivee      || '',
    montant_ht:      String(prestation.prix_ht  || ''),
    tva:             String((prestation.prix_ttc - prestation.prix_ht) || ''),
    montant_ttc:     String(prestation.prix_ttc  || ''),
    bl_ot:           prestation.ot_bl_bs_be  || '',
    bc:              prestation.bon_commande || '',
    delai_paiement:  String(client?.delai_paiement || 60),
    prestation_id:   prestation.id,
  });
  setPrestationPickerOpen(false);
  setShowFactForm(true);
};

const handleSaveFacturation = async () => {
  const payload = {
    company_id:            companyId || null,
    manager_id:            managerId || null,
    prestation_id:         factForm.prestation_id || null,
    date:                  factForm.date                  || null,
    numero_facture:        factForm.numero_facture        || null,
    client:                factForm.client                || null,
    depart:                factForm.depart                || null,
    arrivee:               factForm.arrivee               || null,
    montant_ht:            parseFloat(factForm.montant_ht)  || 0,
    tva:                   parseFloat(factForm.tva)         || 0,
    montant_ttc:           parseFloat(factForm.montant_ttc) || 0,
    bl_ot:                 factForm.bl_ot                   || null,
    bc:                    factForm.bc                      || null,
    delai_paiement:        parseInt(factForm.delai_paiement) || 60,
    date_paiement:         factForm.date_paiement           || null,
    reglement_banque_type: factForm.reglement_banque_type   || null,
    reglement_numero:      factForm.reglement_numero        || null,
    echeances:             factForm.echeances               || null,
    mode_paiement:         factForm.mode_paiement           || null,
    statut:                factForm.statut                  || 'impayé',
    ecart_delai:           calcEcartDelai(
                             factForm.date,
                             factForm.date_paiement,
                             parseInt(factForm.delai_paiement) || 60
                           ),
  };

  if (editingFact) {
    const { error } = await supabase
      .from('suivi_facturation')
      .update(payload)
      .eq('id', editingFact.id);

    if (!error) {
      toast.success("Facture modifiée.");
      setEditingFact(null);

      // Sync shared fields back to linked prestation
      if (editingFact.prestation_id) {
        const sharedBack = {
          client:       payload.client,
          depart:       payload.depart,
          arrivee:      payload.arrivee,
          prix_ht:      payload.montant_ht,
          prix_ttc:     payload.montant_ttc,
          ot_bl_bs_be:  payload.bl_ot,
          bon_commande: payload.bc,
        };
        await supabase
          .from('suivi_prestation')
          .update(sharedBack)
          .eq('id', editingFact.prestation_id);

        if (activeTab === 'suivi') fetchSuivi();
      }

      fetchFacturation();
    } else toast.error(`Erreur: ${error.message}`);
  } else {
    const { error } = await supabase
      .from('suivi_facturation')
      .insert(payload);
    if (!error) {
      toast.success("Facture ajoutée.");
      setShowFactForm(false);
      fetchFacturation();
    } else toast.error(`Erreur: ${error.message}`);
  }
  setFactForm(emptyFactForm);
};

const handleDeleteFact = async (id: string) => {
  if (!confirm('Supprimer cette facture ?')) return;
  const { error } = await supabase.from('suivi_facturation').delete().eq('id', id);
  if (!error) { setFacturationList(prev => prev.filter(f => f.id !== id)); toast.success("Supprimé."); }
  else toast.error(`Erreur: ${error.message}`);
};
const handleFactXLSUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
  const file = e.target.files?.[0];
  if (!file || !companyId) return;
  setUploadingFacts(true);
  try {
    const buffer = await file.arrayBuffer();
    const wb = XLSX.read(buffer, { type: 'array' });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const rows: any[] = XLSX.utils.sheet_to_json(ws, { header: 1 });
    const dataRows = rows.slice(1).filter((r: any[]) => r.length > 0 && r[0]);
    const records = dataRows.map((r: any[]) => ({
      company_id:           companyId,
      manager_id:           managerId || null,
      date:                 r[0] ? String(r[0]) : null,
      numero_facture:       String(r[1] || '').trim(),
      client:               String(r[2] || '').trim(),
      depart:               String(r[3] || '').trim(),
      arrivee:              String(r[4] || '').trim(),
      montant_ht:           parseFloat(r[5]) || 0,
      tva:                  parseFloat(r[6]) || 0,
      montant_ttc:          parseFloat(r[7]) || 0,
      bl_ot:                String(r[8] || '').trim(),
      bc:                   String(r[9] || '').trim(),
      delai_paiement:       parseInt(r[10]) || 60,
      date_paiement:        r[11] ? String(r[11]) : null,
      statut:               String(r[12] || 'impayé').trim().toLowerCase(),
      mode_paiement:        String(r[13] || '').trim(),
    }));
    if (records.length === 0) { toast.error("Aucune donnée trouvée."); return; }
    const { error } = await supabase.from('suivi_facturation').insert(records);
    if (!error) { toast.success(`${records.length} factures importées.`); fetchFacturation(); }
    else toast.error(`Erreur: ${error.message}`);
  } catch (err: any) {
    toast.error(`Erreur: ${err.message}`);
  } finally {
    setUploadingFacts(false);
    e.target.value = '';
  }
};
// ── Invoice Settings ───────────────────────────────────────────────────
const fetchInvoiceSettings = async () => {
  if (!companyId) return;
  const { data } = await supabase
    .from('invoice_settings')
    .select('*')
    .eq('company_id', companyId)
    .maybeSingle();
  if (data) {
    setInvoiceSettings(data);
    if (data.logo_url) setLogoPreviewUrl(data.logo_url);
  }
  setSettingsLoaded(true);
};

const handleLogoUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
  const file = e.target.files?.[0];
  if (!file || !companyId) return;
  setUploadingLogo(true);
  try {
    const ext = file.name.split('.').pop();
    const path = `${companyId}/logo.${ext}`;
    const { error: upErr } = await supabase.storage
      .from('logos')
      .upload(path, file, { upsert: true });
    if (upErr) throw upErr;
    const { data: urlData } = supabase.storage.from('logos').getPublicUrl(path);
    const url = urlData.publicUrl;
    setLogoPreviewUrl(url);
    setInvoiceSettings((p: any) => ({ ...p, logo_url: url, logo_storage_path: path }));
    toast.success("Logo uploadé !");
  } catch (err: any) {
    toast.error(`Erreur logo: ${err.message}`);
  } finally {
    setUploadingLogo(false);
    e.target.value = '';
  }
};

const handleSaveInvoiceSettings = async () => {
  if (!companyId) return;
  setSavingSettings(true);
  const payload = { ...invoiceSettings, company_id: companyId, updated_at: new Date().toISOString() };
  const { error } = await supabase
    .from('invoice_settings')
    .upsert(payload, { onConflict: 'company_id' });
  if (!error) toast.success("Modèle sauvegardé !");
  else toast.error(`Erreur: ${error.message}`);
  setSavingSettings(false);
};
const filteredFacts = facturationList.filter(f => {
  if (factFilter.client  && !f.client?.toLowerCase().includes(factFilter.client.toLowerCase())) return false;
  if (factFilter.statut  && f.statut !== factFilter.statut) return false;
  if (factFilter.dateFrom && f.date < factFilter.dateFrom) return false;
  if (factFilter.dateTo   && f.date > factFilter.dateTo)   return false;
  return true;
});
const today = new Date();
const dueIn15 = facturationList.filter(f => {
  if (f.statut === 'payé' || !f.date || !f.delai_paiement) return false;
  const echeance = new Date(f.date);
  echeance.setDate(echeance.getDate() + f.delai_paiement);
  const daysLeft = Math.round((echeance.getTime() - today.getTime()) / (1000 * 60 * 60 * 24));
  return daysLeft >= 0 && daysLeft <= 15;
});
const overdue = facturationList.filter(f => {
  if (f.statut === 'payé' || !f.date || !f.delai_paiement) return false;
  const echeance = new Date(f.date);
  echeance.setDate(echeance.getDate() + f.delai_paiement);
  return echeance < today;
});
const numberToWords = (n: number): string => {
  // Simple French number to words for invoice
  if (n === 0) return 'Zéro dirham';
  const ones = ['','un','deux','trois','quatre','cinq','six','sept','huit','neuf',
    'dix','onze','douze','treize','quatorze','quinze','seize','dix-sept','dix-huit','dix-neuf'];
  const tens = ['','','vingt','trente','quarante','cinquante','soixante','soixante','quatre-vingt','quatre-vingt'];
  const convert = (num: number): string => {
    if (num < 20) return ones[num];
    if (num < 100) {
      const t = Math.floor(num/10), r = num%10;
      if (t === 7 || t === 9) return tens[t] + (r > 0 ? '-' + ones[10+r] : (t===9?'-dix':''));
      return tens[t] + (r > 0 ? '-' + ones[r] : '');
    }
    if (num < 1000) return (num === 100 ? 'cent' : ones[Math.floor(num/100)] + ' cent') + (num%100 > 0 ? ' ' + convert(num%100) : '');
    if (num < 1000000) {
      const m = Math.floor(num/1000);
      return (m === 1 ? 'mille' : convert(m) + ' mille') + (num%1000 > 0 ? ' ' + convert(num%1000) : '');
    }
    return n.toLocaleString('fr-MA') + ' dirhams';
  };
  const intPart = Math.floor(n);
  const decPart = Math.round((n - intPart) * 100);
  return convert(intPart) + ' dirham' + (intPart > 1 ? 's' : '') +
    (decPart > 0 ? ' et ' + convert(decPart) + ' centime' + (decPart > 1 ? 's' : '') : '');
};
const handleGenerateInvoicePDF = async () => {
  const selected = facturationList.filter(f => selectedFacts.includes(f.id));
  if (selected.length === 0) { toast.error("Sélectionnez au moins une facture."); return; }

  const s = invoiceSettings;
  console.log('Invoice settings logo_url:', s.logo_url);
  console.log('Full settings:', s);
  // Convert logo to base64 so it works in the PDF print window
  let logoBase64 = '';
  if (s.logo_url) {
    try {
      const resp = await fetch(s.logo_url, { mode: 'cors', cache: 'no-cache' });
      const blob = await resp.blob();
      logoBase64 = await new Promise<string>((resolve) => {
        const reader = new FileReader();
        reader.onloadend = () => resolve(reader.result as string);
        reader.readAsDataURL(blob);
      });
      console.log('Logo base64 length:', logoBase64.length, logoBase64.substring(0, 50));
    } catch (err) {
      console.error('Logo fetch failed:', err);
      // Fallback — use URL directly
      logoBase64 = s.logo_url;
    }
  }
  const ROWS_PER_PAGE = s.rows_per_page || 15;
  const pages: any[][] = [];
  for (let i = 0; i < selected.length; i += ROWS_PER_PAGE) {
    pages.push(selected.slice(i, i + ROWS_PER_PAGE));
  }

  const totalHT  = selected.reduce((sum, f) => sum + (f.montant_ht  || 0), 0);
  const totalTVA = selected.reduce((sum, f) => sum + (f.tva         || 0), 0);
  const totalTTC = selected.reduce((sum, f) => sum + (f.montant_ttc || 0), 0);
  const clientName = selected[0]?.client || '';
  const primary = s.primary_color || '#1e40af';
  const accent  = s.accent_color  || '#f59e0b';
  const fs      = s.font_size     || 11;

  // All 18 columns — filtered by user selection
  const allCols = [
    { label: 'Date',            key: 'date',                  show: s.col_show_date,             num: false },
    { label: 'N° Fact.',        key: 'numero_facture',        show: s.col_show_fact,             num: false },
    { label: 'Client',          key: 'client',                show: s.col_show_client,           num: false },
    { label: 'Départ',          key: 'depart',                show: s.col_show_depart,           num: false },
    { label: 'Arrivée',         key: 'arrivee',               show: s.col_show_arrivee,          num: false },
    { label: 'HT (MAD)',        key: 'montant_ht',            show: s.col_show_ht,               num: true  },
    { label: 'TVA (MAD)',       key: 'tva',                   show: s.col_show_tva,              num: true  },
    { label: 'TTC (MAD)',       key: 'montant_ttc',           show: s.col_show_ttc,              num: true  },
    { label: 'BL/OT',          key: 'bl_ot',                 show: s.col_show_bl_ot,            num: false },
    { label: 'BC',              key: 'bc',                    show: s.col_show_bc,               num: false },
    { label: 'Délai (J)',       key: 'delai_paiement',        show: s.col_show_delai,            num: false },
    { label: 'Date Paiement',   key: 'date_paiement',         show: s.col_show_date_paiement,    num: false },
    { label: 'Écart Délai',     key: 'ecart_delai',           show: s.col_show_ecart,            num: false, ecart: true },
    { label: 'Règl. Banque',    key: 'reglement_banque_type', show: s.col_show_reglement_banque, num: false },
    { label: 'Règl. N°',        key: 'reglement_numero',      show: s.col_show_reglement_num,    num: false },
    { label: 'Échéances',       key: 'echeances',             show: s.col_show_echeances,        num: false },
    { label: 'Mode Paiement',   key: 'mode_paiement',         show: s.col_show_mode,             num: false },
    { label: 'Statut',          key: 'statut',                show: s.col_show_statut,           num: false, statut: true },
  ].filter(c => c.show !== false);

  const mt = s.margin_top    || 12;
  const mb = s.margin_bottom || 12;
  const ml = s.margin_left   || 15;
  const mr = s.margin_right  || 15;
  const colW = Math.floor(100 / allCols.length);

  // If HTML template exists use it, otherwise generate
  if (s.invoice_template_html) {
    // Build rows table from selected cols
    const tableHTML = selected.map((f, idx) => `
      <tr>
        <td class="italic">${f.date || '—'}</td>
        <td class="italic">${[f.depart, f.arrivee].filter(Boolean).join(' → ') || f.client || '—'}</td>
        <td class="center italic">${f.tva && f.montant_ht ? Math.round((f.tva/f.montant_ht)*100)+'%' : '—'}</td>
        <td class="center italic">F</td>
        <td class="center italic">1</td>
        <td class="right italic">${Number(f.montant_ht||0).toLocaleString('fr-MA',{minimumFractionDigits:2})}</td>
        <td class="right italic">${Number(f.montant_ht||0).toLocaleString('fr-MA',{minimumFractionDigits:2})}</td>
      </tr>
    `).join('');

    let finalHtml = s.invoice_template_html
      .replaceAll('{{company_name}}',    s.company_name    || activeCompany?.name || '')
      .replaceAll('{{company_address}}', s.address         || '')
      .replaceAll('{{company_phone}}',   s.phone           || '')
      .replaceAll('{{company_email}}',   s.email           || '')
      .replaceAll('{{company_ice}}',     s.ice             || '')
      .replaceAll('{{company_rc}}',      s.rc              || '')
      .replaceAll('{{company_logo}}',    logoBase64 ? `<img src="${logoBase64}" style="max-height:60px;max-width:150px;object-fit:contain"/>` : '')
      .replaceAll('{{invoice_title}}',   s.invoice_title   || 'FACTURE')
      .replaceAll('{{numero_facture}}',  selected[0]?.numero_facture || '')
      .replaceAll('{{date}}',            new Date().toLocaleDateString('fr-MA'))
      .replaceAll('{{client}}',          clientName)
      .replaceAll('{{delai_paiement}}',  String(selected[0]?.delai_paiement || 60))
      .replaceAll('{{date_paiement}}',   selected[0]?.date_paiement || '')
      .replaceAll('{{montant_ht}}',      totalHT.toLocaleString('fr-MA'))
      .replaceAll('{{tva}}',             totalTVA.toLocaleString('fr-MA'))
      .replaceAll('{{montant_ttc}}',     totalTTC.toLocaleString('fr-MA'))
      .replaceAll('{{total_ht}}',        totalHT.toLocaleString('fr-MA'))
      .replaceAll('{{total_tva}}',       totalTVA.toLocaleString('fr-MA'))
      .replaceAll('{{total_ttc}}',       totalTTC.toLocaleString('fr-MA'))
      .replaceAll('{{bl_ot}}',           selected[0]?.bl_ot  || '')
      .replaceAll('{{bc}}',              selected[0]?.bc      || '')
      .replaceAll('{{rib}}',             s.rib             || '')
      .replaceAll('{{bank_name}}',       s.bank_name       || '')
      .replaceAll('{{footer_text}}',     s.footer_text     || '')
      .replaceAll('{{signature_label}}', s.signature_label || 'Signature & Cachet')
      .replaceAll('{{rows}}',            tableHTML)
      .replaceAll('{{client_ice}}',      '')
      .replaceAll('{{montant_lettres}}',  numberToWords(totalTTC))
      .replaceAll('{{page_num}}',        `1 / ${pages.length}`);

    const win = window.open('', '_blank');
    if (win) { win.document.write(finalHtml); win.document.close(); win.focus(); setTimeout(() => win.print(), 600); }
    return;
  }

  // Generated template
  const html = `<!DOCTYPE html>
<html lang="fr">
<head>
<meta charset="UTF-8"/>
<style>
  * { margin:0; padding:0; box-sizing:border-box; }
  body { font-family:Arial,sans-serif; font-size:${fs}px; color:#1e293b; }
  .page { width:210mm; min-height:297mm; padding:${mt}mm ${mr}mm ${mb}mm ${ml}mm; page-break-after:always; position:relative; }
  .page:last-child { page-break-after:auto; }
  .header { display:flex; justify-content:space-between; align-items:flex-start; margin-bottom:14px; padding-bottom:10px; border-bottom:3px solid ${primary}; }
  .logo-img { max-height:55px; max-width:130px; object-fit:contain; }
  .company-name { font-size:${fs+7}px; font-weight:900; color:${primary}; }
  .company-sub { font-size:${fs-2}px; color:#64748b; margin-top:3px; line-height:1.6; }
  .invoice-title { font-size:${fs+11}px; font-weight:900; color:${primary}; letter-spacing:2px; }
  .invoice-num { font-size:${fs+2}px; font-weight:700; color:${accent}; margin-top:3px; }
  .invoice-meta { font-size:${fs-2}px; color:#64748b; margin-top:3px; line-height:1.6; }
  .client-box { background:#f8fafc; border-left:4px solid ${accent}; padding:7px 12px; margin-bottom:12px; border-radius:0 6px 6px 0; }
  .client-name { font-size:${fs+1}px; font-weight:700; color:#1e293b; }
  .client-sub { font-size:${fs-2}px; color:#64748b; margin-top:2px; }
  table { width:100%; border-collapse:collapse; margin-bottom:10px; }
  th { background:${primary}; color:white; padding:5px 6px; text-align:left; font-size:${fs-3}px; text-transform:uppercase; letter-spacing:0.3px; }
  td { padding:4px 6px; border-bottom:1px solid #f1f5f9; font-size:${fs-1}px; }
  tr:nth-child(even) td { background:#f8fafc; }
  .num { text-align:right; }
  .totals-wrap { display:flex; justify-content:flex-end; margin-top:8px; }
  .totals-box { border:2px solid ${primary}; border-radius:6px; padding:8px 14px; min-width:200px; }
  .t-row { display:flex; justify-content:space-between; padding:2px 0; font-size:${fs}px; }
  .t-total { font-weight:700; font-size:${fs+2}px; color:${primary}; border-top:1px solid #e2e8f0; margin-top:4px; padding-top:5px; }
  .payment-box { margin-top:10px; padding:7px 12px; background:#f8fafc; border:1px solid #e2e8f0; border-radius:6px; font-size:${fs-2}px; }
  .payment-title { font-weight:700; color:${primary}; font-size:${fs-1}px; margin-bottom:3px; }
  .sig-wrap { display:flex; justify-content:flex-end; margin-top:14px; }
  .sig-box { border:1px solid #cbd5e1; border-radius:6px; padding:8px 16px; min-width:160px; text-align:center; }
  .sig-label { font-size:${fs-3}px; font-weight:700; color:#64748b; text-transform:uppercase; }
  .sig-space { height:38px; }
  .footer-area { margin-top:14px; border-top:1px solid #e2e8f0; padding-top:5px; }
  .footer-text { font-size:${fs-3}px; color:#94a3b8; text-align:center; }
  .page-num { font-size:${fs-3}px; color:#94a3b8; text-align:right; margin-top:3px; }
  .badge { display:inline-block; padding:1px 5px; border-radius:3px; font-size:${fs-3}px; font-weight:700; }
  .badge-red { background:#fee2e2; color:#991b1b; }
  .badge-green { background:#dcfce7; color:#166534; }
  .statut-paye { background:#dcfce7; color:#166534; padding:1px 6px; border-radius:3px; font-size:${fs-3}px; font-weight:700; text-transform:uppercase; }
  .statut-impaye { background:#fee2e2; color:#991b1b; padding:1px 6px; border-radius:3px; font-size:${fs-3}px; font-weight:700; text-transform:uppercase; }
  @media print { body { print-color-adjust:exact; -webkit-print-color-adjust:exact; } }
</style>
</head>
<body>
${pages.map((pageRows, pageIdx) => `
<div class="page">
  ${!s.skip_header ? `
  <div class="header">
    <div style="display:flex;align-items:center;gap:10px">
      ${logoBase64 ? `<img src="${logoBase64}" class="logo-img"/>` : ''}
      <div>
        <div class="company-name">${s.company_name || activeCompany?.name || ''}</div>
        <div class="company-sub">
          ${s.address ? s.address + '<br/>' : ''}
          ${s.phone   ? 'Tél: ' + s.phone + (s.email ? ' — ' : '<br/>') : ''}
          ${s.email   ? s.email + '<br/>' : ''}
          ${s.ice     ? 'ICE: ' + s.ice + (s.rc ? ' — RC: ' + s.rc : '') : ''}
        </div>
      </div>
    </div>
    <div style="text-align:right">
      <div class="invoice-title">${s.invoice_title || 'FACTURE'}</div>
      ${selected[0]?.numero_facture ? `<div class="invoice-num">N° ${selected[0].numero_facture}</div>` : ''}
      <div class="invoice-meta">Date: ${new Date().toLocaleDateString('fr-MA')}<br/>Page ${pageIdx+1} / ${pages.length}</div>
    </div>
  </div>
  <div class="client-box">
    <div class="client-name">${clientName}</div>
    <div class="client-sub">${selected.length} prestation(s) — Délai: ${selected[0]?.delai_paiement || 60}j</div>
  </div>
  ` : `<div style="text-align:right;font-size:${fs-2}px;color:#64748b;margin-bottom:8px;">Page ${pageIdx+1} / ${pages.length}</div>`}

  <table>
    <thead>
      <tr>${allCols.map(c => `<th style="width:${colW}%">${c.label}</th>`).join('')}</tr>
    </thead>
    <tbody>
      ${pageRows.map(f => `
        <tr>
          ${allCols.map(c => {
            const v = f[c.key];
            if (c.num)    return `<td class="num">${Number(v||0).toLocaleString('fr-MA')} MAD</td>`;
            if (c.ecart)  { const e = v??0; return `<td class="num"><span class="badge ${e>0?'badge-red':'badge-green'}">${e>0?'+':''}${e}j</span></td>`; }
            if (c.statut) return `<td><span class="${v==='payé'?'statut-paye':'statut-impaye'}">${v||'impayé'}</span></td>`;
            return `<td>${v||'—'}</td>`;
          }).join('')}
        </tr>
      `).join('')}
    </tbody>
  </table>

  ${pageIdx === pages.length - 1 ? `
  <div class="totals-wrap">
    <div class="totals-box">
      <div class="t-row"><span>Total HT</span><span>${totalHT.toLocaleString('fr-MA')} MAD</span></div>
      <div class="t-row"><span>Total TVA</span><span>${totalTVA.toLocaleString('fr-MA')} MAD</span></div>
      <div class="t-row t-total"><span>Total TTC</span><span>${totalTTC.toLocaleString('fr-MA')} MAD</span></div>
    </div>
  </div>
  ${(s.rib||s.bank_name) ? `
  <div class="payment-box">
    <div class="payment-title">Coordonnées Bancaires</div>
    ${s.bank_name ? 'Banque: <strong>'+s.bank_name+'</strong>  ' : ''}
    ${s.rib       ? 'RIB: <strong>'+s.rib+'</strong>' : ''}
  </div>` : ''}
  <div class="sig-wrap">
    <div class="sig-box">
      <div class="sig-label">${s.signature_label||'Signature & Cachet'}</div>
      <div class="sig-space"></div>
    </div>
  </div>
  ` : ''}

  ${!s.skip_footer ? `
  <div class="footer-area">
    ${s.footer_text ? `<div class="footer-text">${s.footer_text}</div>` : ''}
    <div class="page-num">Page ${pageIdx+1} / ${pages.length}</div>
  </div>` : ''}
</div>
`).join('')}
</body>
</html>`;

  const win = window.open('', '_blank');
  if (win) { win.document.write(html); win.document.close(); win.focus(); setTimeout(() => win.print(), 600); }
};
  useEffect(() => {
    if (!loading) { if (!user) navigate('/login'); else fetchCompany(); }
  }, [user, loading]);

  useEffect(() => {
    if (activeTab === 'purchases') fetchPurchases();
    if (activeTab === 'fleetfix' && companyId) fetchMechanics();
    if (activeTab === 'suivi') fetchSuivi();
    if (activeTab === 'chauffeurs' && companyId) fetchFleetDrivers();
    if (activeTab === 'clients' && companyId) fetchClients();
    if (activeTab === 'facturation' && companyId) { fetchFacturation(); fetchSuivi(); fetchClients(); fetchInvoiceSettings(); }
    if (activeTab === 'settings' && companyId) fetchInvoiceSettings();
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
  { id: 'staff',       label: 'Staff',               icon: Users },
  { id: 'purchases',   label: 'Achats & Factures',    icon: ShoppingBag },
  { id: 'fleetfix',    label: 'FleetFix',             icon: Wrench },
  { id: 'suivi',       label: 'Suivi Prestation',     icon: FileText },
  { id: 'chauffeurs',  label: 'Chauffeurs',           icon: Truck },
  { id: 'clients',     label: 'Clients',              icon: Users },
  { id: 'facturation', label: 'Suivi Facturation',    icon: Receipt },
  { id: 'settings', label: 'Paramètres Facture', icon: Settings },
] as const;

  const suiviFields = [
    { label: 'Date',                  key: 'date',           type: 'date'   },
    { label: 'Matricule',             key: 'matricule',      type: 'text'   },
    { label: 'Type',                  key: 'type',           type: 'text'   },
    { label: 'Factures',              key: 'facture',        type: 'text'   },
    { label: 'N° Bon de Commande',    key: 'bon_commande',   type: 'text'   },
    { label: 'OT / BL-BS-BE-Booking', key: 'ot_bl_bs_be',   type: 'text'   },
    { label: 'Clients',               key: 'client',         type: 'text'   },
    { label: 'Départ',                key: 'depart',         type: 'text'   },
    { label: 'Arrivée',               key: 'arrivee',        type: 'text'   },
    { label: 'Manutention (MAD)',      key: 'manutention',    type: 'number' },
    { label: 'Immobilisation (MAD)',   key: 'immobilisation', type: 'number' },
    { label: 'Prix HT (MAD)',          key: 'prix_ht',        type: 'number' },
    { label: 'Prix TTC (MAD)',         key: 'prix_ttc',       type: 'number' },
    { label: 'Coût de Revient (MAD)',  key: 'cout_revient',   type: 'number' },
    { label: 'Bénéfice (MAD)',         key: 'benefice',       type: 'number' },
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
            <span className="text-xs font-bold text-slate-400 uppercase tracking-wider hidden sm:block">{activeCompany.name}</span>
          </div>
          <div className="flex items-center gap-4">
            <div className="text-right hidden sm:block">
              <span className="text-[10px] text-slate-400 block uppercase font-bold tracking-wider">Manager</span>
              <span className="text-sm font-bold text-white">{managerName}</span>
            </div>
            <button onClick={() => signOut().then(() => { toast.success("Déconnexion réussie"); navigate('/login'); })}
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
              <motion.aside initial={{ x: -280 }} animate={{ x: 0 }} exit={{ x: -280 }}
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
                        className={`w-full flex items-center gap-3 px-4 py-3 rounded-xl text-sm font-bold transition-all cursor-pointer ${isActive ? 'bg-blue-600 text-white shadow-lg shadow-blue-600/20' : 'text-slate-400 hover:bg-slate-800 hover:text-white'}`}>
                        <Icon size={18} />{item.label}
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

        {/* Main */}
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
                <button onClick={fetchCompany} className="self-start bg-white/10 hover:bg-white/15 px-3 py-2 rounded-lg text-xs font-bold uppercase tracking-wider flex items-center gap-1.5 cursor-pointer">
                  <RefreshCw className="w-3.5 h-3.5" /> Actualiser
                </button>
              </div>
              <div className="max-w-5xl mx-auto"><CreateStaffForm companyId={activeCompany.id} /></div>
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
                    <button onClick={fetchPurchases} className="bg-white/10 hover:bg-white/15 px-3 py-2 rounded-lg text-xs font-bold uppercase tracking-wider flex items-center gap-1.5 cursor-pointer">
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
                              <div className="flex items-center gap-1">
                                <button onClick={() => setEditingPurchase(p)} className="p-1.5 rounded hover:bg-blue-50 text-slate-400 hover:text-blue-600 transition-colors"><Pencil size={13} /></button>
                                <button onClick={() => handleDeletePurchase(p.id)} className="p-1.5 rounded hover:bg-rose-50 text-slate-400 hover:text-rose-600 transition-colors"><Trash2 size={13} /></button>
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
                {/* Mechanic selector */}
                <div className="bg-white rounded-xl border border-slate-200 p-4 space-y-2">
                  <p className="text-[10px] font-black text-slate-400 uppercase tracking-widest mb-3">Mécaniciens</p>
                  {mechanics.length === 0 ? (
                    <p className="text-sm text-slate-400 text-center py-4">Aucun mécanicien trouvé.</p>
                  ) : mechanics.map(m => (
                    <button key={m.id} onClick={() => setSelectedMechanic(m)}
                      className={`w-full text-left px-4 py-3 rounded-lg border transition-all cursor-pointer ${selectedMechanic?.id === m.id ? 'bg-blue-600 border-blue-700 text-white' : 'border-slate-200 hover:border-blue-300 hover:bg-blue-50/40'}`}>
                      <p className={`text-sm font-bold ${selectedMechanic?.id === m.id ? 'text-white' : 'text-slate-800'}`}>{m.full_name}</p>
                      <p className={`text-[10px] font-mono ${selectedMechanic?.id === m.id ? 'text-blue-200' : 'text-slate-400'}`}>{m.employee_code}</p>
                    </button>
                  ))}
                </div>

                {/* Right panel */}
                <div className="lg:col-span-2 space-y-6">
                  {!selectedMechanic ? (
                    <div className="bg-white rounded-xl border border-slate-200 flex items-center justify-center h-48">
                      <p className="text-slate-400 text-sm">Sélectionnez un mécanicien</p>
                    </div>
                  ) : (
                    <>
                      {/* Balance */}
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

                      {/* Add funds */}
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

                      {/* Fund topups history */}
                      <div className="bg-white rounded-xl border border-slate-200 overflow-hidden">
                        <div className="p-4 border-b border-slate-100 bg-emerald-50">
                          <p className="text-[10px] font-black text-emerald-700 uppercase tracking-widest">
                            Fonds Accordés — {topups.length} versements
                          </p>
                        </div>
                        {topups.length === 0 ? (
                          <div className="py-6 text-center text-slate-400 text-sm">Aucun fonds versé.</div>
                        ) : (
                          <div className="overflow-x-auto">
                            <table className="w-full text-left">
                              <thead className="bg-slate-50 border-b border-slate-100">
                                <tr>
                                  {['Date','Montant','Note','Actions'].map(h => (
                                    <th key={h} className="px-4 py-2 text-[9px] font-black text-slate-400 uppercase tracking-widest">{h}</th>
                                  ))}
                                </tr>
                              </thead>
                              <tbody className="divide-y divide-slate-50">
                                {topups.map(t => (
                                  <tr key={t.id} className="hover:bg-slate-50">
                                    <td className="px-4 py-3 text-xs text-slate-600">{t.created_at?.split('T')[0]}</td>
                                    <td className="px-4 py-3 font-mono text-xs font-bold text-emerald-600">
                                      +{Number(t.amount).toLocaleString('fr-MA', { minimumFractionDigits: 2 })} MAD
                                    </td>
                                    <td className="px-4 py-3 text-xs text-slate-600">{t.notes || '—'}</td>
                                    <td className="px-4 py-3">
                                      <div className="flex items-center gap-1">
                                        <button onClick={() => { setEditingTopup(t); setTopupEditAmount(String(t.amount)); setTopupEditNotes(t.notes || ''); }}
                                          className="p-1.5 rounded hover:bg-blue-50 text-slate-400 hover:text-blue-600 transition-colors">
                                          <Pencil size={13} />
                                        </button>
                                        <button onClick={() => handleDeleteTopup(t.id)}
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
                        )}
                      </div>

                      {/* Maintenance records */}
                      <div className="bg-white rounded-xl border border-slate-200 overflow-hidden">
                        <div className="p-4 border-b border-slate-100 bg-rose-50">
                          <p className="text-[10px] font-black text-rose-700 uppercase tracking-widest">
                            Fiches d'Entretien — {maintenance.length} réparations
                          </p>
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
                                    <td className="px-4 py-3 font-mono text-xs font-bold text-red-600">
                                      -{r.total_cost?.toLocaleString('fr-MA', { minimumFractionDigits: 2 })}
                                    </td>
                                    <td className="px-4 py-3">
                                      {r.receipt_url ? (
                                        <button onClick={() => setViewingReceipt(r.receipt_url)} className="text-blue-600"><Eye size={16} /></button>
                                      ) : <span className="text-slate-300 text-xs">—</span>}
                                    </td>
                                    <td className="px-4 py-3">
                                      <div className="flex items-center gap-1">
                                        <button onClick={() => setEditingRecord(r)} className="p-1.5 rounded hover:bg-blue-50 text-slate-400 hover:text-blue-600 transition-colors"><Pencil size={13} /></button>
                                        <button onClick={() => handleDeleteRecord(r.id)} className="p-1.5 rounded hover:bg-rose-50 text-slate-400 hover:text-rose-600 transition-colors"><Trash2 size={13} /></button>
                                      </div>
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
                    <button onClick={fetchSuivi} className="bg-white/10 hover:bg-white/15 px-3 py-2 rounded-lg text-xs font-bold uppercase tracking-wider flex items-center gap-1.5 cursor-pointer">
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
                          <tr><td colSpan={20} className="px-4 py-10 text-center text-sm text-slate-400">Aucune prestation. Cliquez sur "Nouveau".</td></tr>
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
          {/* TAB: CHAUFFEURS */}
{activeTab === 'chauffeurs' && (
  <div>
    <div className="mb-6 bg-slate-900 text-white rounded-xl p-6 border border-slate-800">
      <div className="flex items-center justify-between flex-wrap gap-3">
        <div>
          <span className="inline-flex items-center gap-1.5 px-2 py-0.5 rounded text-[10px] font-extrabold uppercase tracking-widest bg-cyan-500 text-slate-950 mb-2">
            <Truck className="w-3.5 h-3.5" /> Flotte Chauffeurs
          </span>
          <h1 className="text-2xl font-extrabold tracking-tight">Gestion des Chauffeurs</h1>
          <p className="text-sm text-slate-400 mt-1">{fleetDrivers.length} chauffeurs enregistrés</p>
        </div>
        <div className="flex gap-2 flex-wrap items-center">
          <button onClick={fetchFleetDrivers}
            className="bg-white/10 hover:bg-white/15 px-3 py-2 rounded-lg text-xs font-bold uppercase tracking-wider flex items-center gap-1.5 cursor-pointer">
            <RefreshCw className="w-3.5 h-3.5" /> Actualiser
          </button>
          <button onClick={() => exportToXLS(fleetDrivers.map(d => ({
            'Code': d.code, 'Nom/Prénom': d.nom_prenom, 'Immatriculation': d.immatriculation,
            'Type': d.type_vehicule, 'CIN': d.cin, 'IMM CNSS': d.imm_cnss,
            'Fonction': d.fonction, 'Date Naissance': d.date_naissance,
            'Situation Familiale': d.situation_familiale, 'Nb Déduction': d.nb_deduction,
            'Date Embauche': d.date_embauche, 'Adresse': d.adresse,
          })), 'chauffeurs_export')}
            className="bg-emerald-600 hover:bg-emerald-700 text-white px-3 py-2 rounded-lg text-xs font-black uppercase tracking-wider flex items-center gap-1.5 cursor-pointer">
            <Download size={14} /> Export XLS
          </button>
          {/* XLS Upload */}
          <label className={`flex items-center gap-1.5 px-3 py-2 rounded-lg text-xs font-black uppercase tracking-wider cursor-pointer transition-all ${uploadingXLS ? 'bg-slate-600 opacity-60' : 'bg-blue-600 hover:bg-blue-700'} text-white`}>
            <Upload size={14} />
            {uploadingXLS ? 'Importation...' : 'Importer XLS'}
            <input type="file" accept=".xlsx,.xls" onChange={handleXLSUpload} className="hidden" disabled={uploadingXLS} />
          </label>
        </div>
      </div>
    </div>

    {/* Template download hint */}
    <div className="mb-4 p-4 bg-blue-50 border border-blue-100 rounded-xl text-xs text-blue-700 font-medium">
      📋 Format attendu du fichier XLS — colonnes dans l'ordre :
      <span className="font-black ml-1">Code | Nom/Prénom | Immatriculation | Type | CIN | IMM CNSS | Fonction | Date Naissance | Situation Familiale | Nb Déduction | Date Embauche | Adresse</span>
    </div>

    {loadingDrivers ? (
      <div className="flex items-center justify-center py-20"><Loader2 className="w-6 h-6 text-blue-600 animate-spin" /></div>
    ) : (
      <div className="bg-white rounded-xl border border-slate-200 overflow-hidden">
        <div className="overflow-x-auto">
          <table className="w-full text-left min-w-[1100px]">
            <thead className="bg-slate-50 border-b border-slate-200">
              <tr>
                {['Code','Nom / Prénom','Immat.','Type','CIN','IMM CNSS','Fonction','Naissance','Situation','Déductions','Embauche','Adresse','Actions'].map(h => (
                  <th key={h} className="px-3 py-3 text-[9px] font-black text-slate-400 uppercase tracking-widest whitespace-nowrap">{h}</th>
                ))}
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100">
              {fleetDrivers.length === 0 ? (
                <tr>
                  <td colSpan={13} className="px-4 py-12 text-center">
                    <div className="flex flex-col items-center gap-3">
                      <Truck size={32} className="text-slate-300" />
                      <p className="text-sm text-slate-400 font-medium">Aucun chauffeur importé.</p>
                      <p className="text-xs text-slate-400">Cliquez sur "Importer XLS" pour charger votre liste de chauffeurs.</p>
                    </div>
                  </td>
                </tr>
              ) : fleetDrivers.map(d => (
                <tr key={d.id} className="hover:bg-slate-50 transition-colors">
                  <td className="px-3 py-3 font-mono text-xs font-bold text-blue-600">{d.code || '—'}</td>
                  <td className="px-3 py-3 text-xs font-semibold text-slate-800">{d.nom_prenom}</td>
                  <td className="px-3 py-3 font-mono text-xs font-bold text-slate-700">{d.immatriculation || '—'}</td>
                  <td className="px-3 py-3 text-xs text-slate-600">{d.type_vehicule || '—'}</td>
                  <td className="px-3 py-3 font-mono text-xs text-slate-600">{d.cin || '—'}</td>
                  <td className="px-3 py-3 font-mono text-xs text-slate-600">{d.imm_cnss || '—'}</td>
                  <td className="px-3 py-3 text-xs text-slate-600">{d.fonction || '—'}</td>
                  <td className="px-3 py-3 text-xs text-slate-500">{d.date_naissance || '—'}</td>
                  <td className="px-3 py-3 text-xs text-slate-500">{d.situation_familiale || '—'}</td>
                  <td className="px-3 py-3 text-xs text-center text-slate-600">{d.nb_deduction ?? '—'}</td>
                  <td className="px-3 py-3 text-xs text-slate-500">{d.date_embauche || '—'}</td>
                  <td className="px-3 py-3 text-xs text-slate-500 max-w-[150px] truncate">{d.adresse || '—'}</td>
                  <td className="px-3 py-3">
                    <div className="flex items-center gap-1">
                      <button onClick={() => { setEditingDriver(d); setDriverEditForm({ ...d }); }}
                        className="p-1.5 rounded hover:bg-blue-50 text-slate-400 hover:text-blue-600 transition-colors">
                        <Pencil size={13} />
                      </button>
                      <button onClick={() => handleDeleteDriver(d.id)}
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
{activeTab === 'clients' && (
  <div>
    <div className="mb-6 bg-slate-900 text-white rounded-xl p-6 border border-slate-800">
      <div className="flex items-center justify-between flex-wrap gap-3">
        <div>
          <span className="inline-flex items-center gap-1.5 px-2 py-0.5 rounded text-[10px] font-extrabold uppercase tracking-widest bg-emerald-500 text-slate-950 mb-2">
            <Users className="w-3.5 h-3.5" /> Base Clients
          </span>
          <h1 className="text-2xl font-extrabold tracking-tight">Gestion des Clients</h1>
          <p className="text-sm text-slate-400 mt-1">{clientsList.length} clients enregistrés</p>
        </div>
        <div className="flex gap-2 flex-wrap items-center">
          <button onClick={fetchClients}
            className="bg-white/10 hover:bg-white/15 px-3 py-2 rounded-lg text-xs font-bold uppercase tracking-wider flex items-center gap-1.5 cursor-pointer">
            <RefreshCw className="w-3.5 h-3.5" /> Actualiser
          </button>
          <button onClick={() => exportToXLS(clientsList.map(c => ({
            'Client': c.nom, 'Adresse': c.adresse, 'ICE': c.ice, 'Délai Paiement': c.delai_paiement,
          })), 'clients_export')}
            className="bg-emerald-600 hover:bg-emerald-700 text-white px-3 py-2 rounded-lg text-xs font-black uppercase tracking-wider flex items-center gap-1.5 cursor-pointer">
            <Download size={14} /> Export XLS
          </button>
          <label className={`flex items-center gap-1.5 px-3 py-2 rounded-lg text-xs font-black uppercase tracking-wider cursor-pointer transition-all ${uploadingClients ? 'bg-slate-600 opacity-60' : 'bg-blue-600 hover:bg-blue-700'} text-white`}>
            <Upload size={14} />
            {uploadingClients ? 'Importation...' : 'Importer XLS'}
            <input type="file" accept=".xlsx,.xls" onChange={handleClientsXLSUpload} className="hidden" disabled={uploadingClients} />
          </label>
        </div>
      </div>
    </div>

    <div className="mb-4 p-4 bg-blue-50 border border-blue-100 rounded-xl text-xs text-blue-700 font-medium">
      📋 Format attendu : <span className="font-black">Clients | Adresse | ICE | Délai de paiement/Jour</span>
    </div>

    {loadingClients ? (
      <div className="flex items-center justify-center py-20"><Loader2 className="w-6 h-6 text-blue-600 animate-spin" /></div>
    ) : (
      <div className="bg-white rounded-xl border border-slate-200 overflow-hidden">
        <div className="overflow-x-auto">
          <table className="w-full text-left">
            <thead className="bg-slate-50 border-b border-slate-200">
              <tr>
                {['Client','Adresse','ICE','Délai (J)','Actions'].map(h => (
                  <th key={h} className="px-4 py-3 text-[10px] font-black text-slate-400 uppercase tracking-widest">{h}</th>
                ))}
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100">
              {clientsList.length === 0 ? (
                <tr>
                  <td colSpan={5} className="px-4 py-12 text-center">
                    <div className="flex flex-col items-center gap-3">
                      <Users size={32} className="text-slate-300" />
                      <p className="text-sm text-slate-400 font-medium">Aucun client importé.</p>
                      <p className="text-xs text-slate-400">Cliquez sur "Importer XLS" pour charger votre base clients.</p>
                    </div>
                  </td>
                </tr>
              ) : clientsList.map(c => (
                <tr key={c.id} className="hover:bg-slate-50 transition-colors">
                  <td className="px-4 py-3 text-xs font-semibold text-slate-800">{c.nom}</td>
                  <td className="px-4 py-3 text-xs text-slate-600 max-w-[300px] truncate">{c.adresse || '—'}</td>
                  <td className="px-4 py-3 font-mono text-xs text-blue-600">{c.ice || '—'}</td>
                  <td className="px-4 py-3 text-xs text-center font-bold text-slate-700">{c.delai_paiement} J</td>
                  <td className="px-4 py-3">
                    <div className="flex items-center gap-1">
                      <button onClick={() => { setEditingClient(c); setClientEditForm({ ...c }); }}
                        className="p-1.5 rounded hover:bg-blue-50 text-slate-400 hover:text-blue-600 transition-colors">
                        <Pencil size={13} />
                      </button>
                      <button onClick={() => handleDeleteClient(c.id)}
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
{activeTab === 'facturation' && (
  <div>
  {/* Payment Due Notifications */}
    {(dueIn15.length > 0 || overdue.length > 0) && (
      <div className="mb-4 space-y-2">
        {overdue.length > 0 && (
          <div className="bg-rose-50 border border-rose-200 rounded-xl p-4 flex items-center gap-3">
            <div className="w-8 h-8 bg-rose-100 rounded-full flex items-center justify-center shrink-0">
              <Receipt className="w-4 h-4 text-rose-600" />
            </div>
            <div className="flex-1">
              <p className="text-sm font-black text-rose-800">⚠ {overdue.length} facture(s) en retard de paiement</p>
              <p className="text-xs text-rose-600 mt-0.5">
                {overdue.map(f => `${f.client} — N°${f.numero_facture || '?'}`).join(' • ')}
              </p>
            </div>
          </div>
        )}
        {dueIn15.length > 0 && (
          <div className="bg-amber-50 border border-amber-200 rounded-xl p-4 flex items-center gap-3">
            <div className="w-8 h-8 bg-amber-100 rounded-full flex items-center justify-center shrink-0">
              <Receipt className="w-4 h-4 text-amber-600" />
            </div>
            <div className="flex-1">
              <p className="text-sm font-black text-amber-800">🔔 {dueIn15.length} facture(s) à échéance dans 15 jours</p>
              <p className="text-xs text-amber-600 mt-0.5">
                {dueIn15.map(f => {
                  const echeance = new Date(f.date);
                  echeance.setDate(echeance.getDate() + f.delai_paiement);
                  const daysLeft = Math.round((echeance.getTime() - today.getTime()) / (1000 * 60 * 60 * 24));
                  return `${f.client} — ${daysLeft}j restants`;
                }).join(' • ')}
              </p>
            </div>
          </div>
        )}
      </div>
    )}
    {/* Header */}
    <div className="mb-6 bg-slate-900 text-white rounded-xl p-6 border border-slate-800">
      <div className="flex items-center justify-between flex-wrap gap-3">
        <div>
          <span className="inline-flex items-center gap-1.5 px-2 py-0.5 rounded text-[10px] font-extrabold uppercase tracking-widest bg-blue-500 text-white mb-2">
            <Receipt className="w-3.5 h-3.5" /> Suivi Facturation
          </span>
          <h1 className="text-2xl font-extrabold tracking-tight">Gestion des Factures</h1>
          <p className="text-sm text-slate-400 mt-1">{facturationList.length} factures — {selectedFacts.length} sélectionnées</p>
        </div>
        <div className="flex gap-2 flex-wrap">
          <button onClick={fetchFacturation}
            className="bg-white/10 hover:bg-white/15 px-3 py-2 rounded-lg text-xs font-bold uppercase tracking-wider flex items-center gap-1.5 cursor-pointer">
            <RefreshCw className="w-3.5 h-3.5" /> Actualiser
          </button>
          <button onClick={() => exportToXLS(filteredFacts.map(f => ({
            'Date': f.date,
            'N° Facture': f.numero_facture,
            'Client': f.client,
            'Départ': f.depart,
            'Arrivée': f.arrivee,
            'Montant HT': f.montant_ht,
            'TVA': f.tva,
            'Montant TTC': f.montant_ttc,
            'BL/OT': f.bl_ot,
            'BC': f.bc,
            'Délai Paiement (J)': f.delai_paiement,
            'Date Paiement': f.date_paiement,
            'Écart Délai (J)': f.ecart_delai ?? calcEcartDelai(f.date, f.date_paiement, f.delai_paiement),
            'Règlement Banque/Type': f.reglement_banque_type,
            'Règlement N°': f.reglement_numero,
            'Échéances': f.echeances,
            'Mode Paiement': f.mode_paiement,
            'Statut': f.statut,
          })), 'suivi_facturation')}
            className="bg-emerald-600 hover:bg-emerald-700 text-white px-3 py-2 rounded-lg text-xs font-black uppercase tracking-wider flex items-center gap-1.5 cursor-pointer">
            <Download size={14} /> Export XLS
          </button>
          {selectedFacts.length > 0 && (
            <button onClick={handleGenerateInvoicePDF}
              className="bg-violet-600 hover:bg-violet-700 text-white px-3 py-2 rounded-lg text-xs font-black uppercase tracking-wider flex items-center gap-1.5 cursor-pointer">
              <FileText size={14} /> Générer PDF ({selectedFacts.length})
            </button>
          )}
          <label className={`flex items-center gap-1.5 px-3 py-2 rounded-lg text-xs font-black uppercase tracking-wider cursor-pointer transition-all ${uploadingFacts ? 'bg-slate-600 opacity-60' : 'bg-amber-600 hover:bg-amber-700'} text-white`}>
            <Upload size={14} />
            {uploadingFacts ? 'Importation...' : 'Importer XLS'}
            <input type="file" accept=".xlsx,.xls" onChange={handleFactXLSUpload} className="hidden" disabled={uploadingFacts} />
          </label>
          <button onClick={() => { setPrestationPickerOpen(true); setEditingFact(null); setFactForm(emptyFactForm); }}
            className="bg-blue-600 hover:bg-blue-700 text-white px-3 py-2 rounded-lg text-xs font-black uppercase tracking-wider flex items-center gap-1.5 cursor-pointer">
            <Plus size={14} /> Nouvelle Facture
          </button>
        </div>
      </div>
    </div>

    {/* Filters */}
    <div className="mb-4 bg-white rounded-xl border border-slate-200 p-4 flex flex-wrap gap-3 items-end">
      <div>
        <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Client</label>
        <input type="text" placeholder="Filtrer par client..."
          value={factFilter.client}
          onChange={e => setFactFilter(p => ({ ...p, client: e.target.value }))}
          className="block mt-1 h-8 rounded-lg border-2 border-slate-200 px-3 text-xs focus:outline-none focus:border-blue-500 w-48" />
      </div>
      <div>
        <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Date de</label>
        <input type="date" value={factFilter.dateFrom}
          onChange={e => setFactFilter(p => ({ ...p, dateFrom: e.target.value }))}
          className="block mt-1 h-8 rounded-lg border-2 border-slate-200 px-3 text-xs focus:outline-none focus:border-blue-500" />
      </div>
      <div>
        <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Date à</label>
        <input type="date" value={factFilter.dateTo}
          onChange={e => setFactFilter(p => ({ ...p, dateTo: e.target.value }))}
          className="block mt-1 h-8 rounded-lg border-2 border-slate-200 px-3 text-xs focus:outline-none focus:border-blue-500" />
      </div>
      <div>
        <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Statut</label>
        <select value={factFilter.statut}
          onChange={e => setFactFilter(p => ({ ...p, statut: e.target.value }))}
          className="block mt-1 h-8 rounded-lg border-2 border-slate-200 px-3 text-xs focus:outline-none focus:border-blue-500">
          <option value="">Tous</option>
          <option value="payé">Payé</option>
          <option value="impayé">Impayé</option>
        </select>
      </div>
      <button onClick={() => setFactFilter({ client: '', dateFrom: '', dateTo: '', statut: '' })}
        className="h-8 px-3 bg-slate-100 hover:bg-slate-200 text-slate-600 text-xs font-bold rounded-lg cursor-pointer">
        Réinitialiser
      </button>
      {filteredFacts.length > 0 && (
        <button onClick={() => {
          const allIds = filteredFacts.map(f => f.id);
          setSelectedFacts(prev => prev.length === allIds.length ? [] : allIds);
        }}
          className="h-8 px-3 bg-blue-50 hover:bg-blue-100 text-blue-700 text-xs font-bold rounded-lg cursor-pointer border border-blue-200">
          {selectedFacts.length === filteredFacts.length ? 'Tout désélectionner' : 'Tout sélectionner'}
        </button>
      )}
    </div>
    <div className="mb-4 p-3 bg-amber-50 border border-amber-100 rounded-xl text-xs text-amber-700 font-medium">
      📋 Format import XLS — colonnes dans l'ordre :
      <span className="font-black ml-1">Date | N° Facture | Client | Départ | Arrivée | Montant HT | TVA | Montant TTC | BL/OT | BC | Délai (J) | Date Paiement | Statut | Mode Paiement</span>
    </div>

    {/* Table */}
    {loadingFacturation ? (
      <div className="flex items-center justify-center py-20"><Loader2 className="w-6 h-6 text-blue-600 animate-spin" /></div>
    ) : (
      <div className="bg-white rounded-xl border border-slate-200 overflow-hidden">
        <div className="overflow-x-auto">
          <table className="w-full text-left min-w-[1400px]">
            <thead className="bg-slate-50 border-b border-slate-200">
              <tr>
                <th className="px-3 py-3 w-10">
                  <input type="checkbox"
                    checked={selectedFacts.length === filteredFacts.length && filteredFacts.length > 0}
                    onChange={() => {
                      const allIds = filteredFacts.map(f => f.id);
                      setSelectedFacts(prev => prev.length === allIds.length ? [] : allIds);
                    }} />
                </th>
                {['Date','N° Fact.','Client','Départ','Arrivée','HT','TVA','TTC','BL/OT','BC','Délai','Date Paie.','Écart Délai','Règl. Banque','Règl. N°','Échéances','Mode','Statut','Actions'].map(h => (
                  <th key={h} className="px-3 py-3 text-[9px] font-black text-slate-400 uppercase tracking-widest whitespace-nowrap">{h}</th>
                ))}
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100">
              {filteredFacts.length === 0 ? (
                <tr><td colSpan={16} className="px-4 py-10 text-center text-sm text-slate-400">
                  Aucune facture. Cliquez sur "Nouvelle Facture".
                </td></tr>
              ) : filteredFacts.map(f => (
                <tr key={f.id} className={`hover:bg-slate-50 transition-colors ${selectedFacts.includes(f.id) ? 'bg-blue-50/50' : ''}`}>
                  <td className="px-3 py-3">
                    <input type="checkbox" checked={selectedFacts.includes(f.id)}
                      onChange={() => setSelectedFacts(prev =>
                        prev.includes(f.id) ? prev.filter(id => id !== f.id) : [...prev, f.id]
                      )} />
                  </td>
                  <td className="px-3 py-3 text-xs text-slate-700 whitespace-nowrap">{f.date}</td>
                  <td className="px-3 py-3 font-mono text-xs font-bold text-blue-600">{f.numero_facture || '—'}</td>
                  <td className="px-3 py-3 text-xs font-semibold text-slate-800">{f.client || '—'}</td>
                  <td className="px-3 py-3 text-xs text-slate-600">{f.depart || '—'}</td>
                  <td className="px-3 py-3 text-xs text-slate-600">{f.arrivee || '—'}</td>
                  <td className="px-3 py-3 font-mono text-xs text-slate-700">{Number(f.montant_ht).toLocaleString('fr-MA')}</td>
                  <td className="px-3 py-3 font-mono text-xs text-amber-700">{Number(f.tva).toLocaleString('fr-MA')}</td>
                  <td className="px-3 py-3 font-mono text-xs font-bold text-slate-900">{Number(f.montant_ttc).toLocaleString('fr-MA')}</td>
                  <td className="px-3 py-3 text-xs text-slate-600">{f.bl_ot || '—'}</td>
                  <td className="px-3 py-3 text-xs text-slate-600">{f.bc || '—'}</td>
                  <td className="px-3 py-3 text-xs text-center text-slate-600">{f.delai_paiement} J</td>
                  <td className="px-3 py-3 text-xs text-slate-600 whitespace-nowrap">{f.date_paiement || '—'}</td>
                  <td className="px-3 py-3">
                    {(() => {
                      const ecart = f.ecart_delai ?? calcEcartDelai(f.date, f.date_paiement, f.delai_paiement);
                      return (
                        <span className={`px-2 py-0.5 rounded text-[9px] font-bold font-mono ${ecart > 0 ? 'bg-rose-50 text-rose-700' : ecart < 0 ? 'bg-emerald-50 text-emerald-700' : 'bg-slate-100 text-slate-600'}`}>
                          {ecart > 0 ? `+${ecart}` : ecart} j
                        </span>
                      );
                    })()}
                  </td>
                  <td className="px-3 py-3 text-xs text-slate-600">{f.reglement_banque_type || '—'}</td>
                  <td className="px-3 py-3 font-mono text-xs text-slate-600">{f.reglement_numero || '—'}</td>
                  <td className="px-3 py-3 text-xs text-slate-600">{f.echeances || '—'}</td>
                  <td className="px-3 py-3 text-xs text-slate-600">{f.mode_paiement || '—'}</td>
                  <td className="px-3 py-3">
                    <span className={`px-2 py-0.5 rounded text-[9px] font-bold uppercase ${f.statut === 'payé' ? 'bg-emerald-50 text-emerald-700' : 'bg-rose-50 text-rose-700'}`}>
                      {f.statut || 'impayé'}
                    </span>
                  </td>
                  <td className="px-3 py-3">
                    <div className="flex items-center gap-1">
                      <button onClick={() => {
                        setEditingFact(f);
                        const ht = Number(f.montant_ht) || 0;
                        const tvaAmt = Number(f.tva) || 0;
                        const rate = ht > 0 ? String(Math.round(tvaAmt / ht * 100)) : '';
                        setFactForm({
                          date:                  f.date                  || '',
                          numero_facture:        f.numero_facture        || '',
                          client:                f.client                || '',
                          depart:                f.depart                || '',
                          arrivee:               f.arrivee               || '',
                          montant_ht:            String(f.montant_ht     || ''),
                          tva:                   String(f.tva            || ''),
                          montant_ttc:           String(f.montant_ttc    || ''),
                          tva_rate:              rate,
                          bl_ot:                 f.bl_ot                 || '',
                          bc:                    f.bc                    || '',
                          delai_paiement:        String(f.delai_paiement || 60),
                          date_paiement:         f.date_paiement         || '',
                          reglement_banque_type: f.reglement_banque_type || '',
                          reglement_numero:      f.reglement_numero      || '',
                          echeances:             f.echeances             || '',
                          mode_paiement:         f.mode_paiement         || '',
                          statut:                f.statut                || 'impayé',
                        });
                        setShowFactForm(true);
                      }} className="p-1.5 rounded hover:bg-blue-50 text-slate-400 hover:text-blue-600 transition-colors">
                        <Pencil size={13} />
                      </button>
                      <button onClick={() => handleDeleteFact(f.id)}
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
{activeTab === 'settings' && (
  <div>
    {/* Header */}
    <div className="mb-6 bg-slate-900 text-white rounded-xl p-6 border border-slate-800">
      <div className="flex items-center justify-between flex-wrap gap-3">
        <div>
          <span className="inline-flex items-center gap-1.5 px-2 py-0.5 rounded text-[10px] font-extrabold uppercase tracking-widest bg-slate-500 text-white mb-2">
            <Settings className="w-3.5 h-3.5" /> Paramètres
          </span>
          <h1 className="text-2xl font-extrabold tracking-tight">Modèle de Facture</h1>
          <p className="text-sm text-slate-400 mt-1">Choisissez comment configurer votre modèle de facture.</p>
        </div>
        {invoiceSettings.setup_mode && (
          <button onClick={() => setInvoiceSettings((p: any) => ({ ...p, setup_mode: '' }))}
            className="bg-white/10 hover:bg-white/20 text-white px-4 py-2 rounded-lg text-xs font-black uppercase tracking-wider flex items-center gap-2 cursor-pointer">
            <X size={14} /> Changer de mode
          </button>
        )}
      </div>
    </div>

    {/* MODE SELECTION */}
    {!invoiceSettings.setup_mode && (
      <div className="grid grid-cols-1 md:grid-cols-2 gap-6 max-w-3xl mx-auto mt-8">
        <div onClick={() => setInvoiceSettings((p: any) => ({ ...p, setup_mode: 'import' }))}
          className="bg-white border-2 border-slate-200 hover:border-blue-500 hover:shadow-lg rounded-2xl p-8 cursor-pointer transition-all group flex flex-col items-center text-center gap-4">
          <div className="w-16 h-16 bg-blue-50 group-hover:bg-blue-600 rounded-2xl flex items-center justify-center transition-colors">
            <Upload size={28} className="text-blue-600 group-hover:text-white transition-colors" />
          </div>
          <div>
            <h2 className="text-lg font-black text-slate-800 uppercase tracking-tight">Importer ma facture</h2>
            <p className="text-xs text-slate-500 mt-2 leading-relaxed">
              Uploadez votre template HTML existant. Le système injecte automatiquement vos données dedans à chaque génération PDF.
            </p>
          </div>
          <div className="mt-2 px-4 py-2 bg-blue-50 group-hover:bg-blue-600 text-blue-600 group-hover:text-white rounded-lg text-xs font-black uppercase tracking-wider transition-colors">
            Choisir ce mode →
          </div>
        </div>
        <div onClick={() => setInvoiceSettings((p: any) => ({ ...p, setup_mode: 'create' }))}
          className="bg-white border-2 border-slate-200 hover:border-violet-500 hover:shadow-lg rounded-2xl p-8 cursor-pointer transition-all group flex flex-col items-center text-center gap-4">
          <div className="w-16 h-16 bg-violet-50 group-hover:bg-violet-600 rounded-2xl flex items-center justify-center transition-colors">
            <FileText size={28} className="text-violet-600 group-hover:text-white transition-colors" />
          </div>
          <div>
            <h2 className="text-lg font-black text-slate-800 uppercase tracking-tight">Créer mon modèle</h2>
            <p className="text-xs text-slate-500 mt-2 leading-relaxed">
              Créez votre facture depuis zéro. Remplissez vos informations, couleurs, colonnes — aperçu en temps réel.
            </p>
          </div>
          <div className="mt-2 px-4 py-2 bg-violet-50 group-hover:bg-violet-600 text-violet-600 group-hover:text-white rounded-lg text-xs font-black uppercase tracking-wider transition-colors">
            Choisir ce mode →
          </div>
        </div>
      </div>
    )}

    {/* ── SHARED COLUMN SELECTOR (used in both modes) ── */}
    {invoiceSettings.setup_mode && (() => {
      const allCols = [
        { label: 'Date',            key: 'col_show_date',             default: true  },
        { label: 'N° Facture',      key: 'col_show_fact',             default: true  },
        { label: 'Client',          key: 'col_show_client',           default: true  },
        { label: 'Départ',          key: 'col_show_depart',           default: true  },
        { label: 'Arrivée',         key: 'col_show_arrivee',          default: true  },
        { label: 'Montant HT',      key: 'col_show_ht',               default: true  },
        { label: 'TVA',             key: 'col_show_tva',              default: true  },
        { label: 'Montant TTC',     key: 'col_show_ttc',              default: true  },
        { label: 'BL / OT',         key: 'col_show_bl_ot',            default: true  },
        { label: 'BC',              key: 'col_show_bc',               default: true  },
        { label: 'Délai Paiement',  key: 'col_show_delai',            default: false },
        { label: 'Date Paiement',   key: 'col_show_date_paiement',    default: false },
        { label: 'Écart Délai',     key: 'col_show_ecart',            default: false },
        { label: 'Règl. Banque',    key: 'col_show_reglement_banque', default: false },
        { label: 'Règl. N°',        key: 'col_show_reglement_num',    default: false },
        { label: 'Échéances',       key: 'col_show_echeances',        default: false },
        { label: 'Mode Paiement',   key: 'col_show_mode',             default: false },
        { label: 'Statut',          key: 'col_show_statut',           default: true  },
      ];
      return (
        <div className="mb-6 bg-white rounded-xl border border-slate-200 overflow-hidden">
          <div className="px-5 py-3 bg-slate-50 border-b border-slate-200 flex items-center justify-between">
            <div className="flex items-center gap-2">
              <div className="w-6 h-6 rounded-full bg-blue-600 flex items-center justify-center">
                <span className="text-[10px] font-black text-white">✓</span>
              </div>
              <span className="text-xs font-black text-slate-700 uppercase tracking-widest">Colonnes à inclure dans le PDF</span>
              <span className="text-[10px] text-slate-400 ml-1">— s'applique aux deux modes</span>
            </div>
            <span className="text-[10px] text-slate-400">
              {allCols.filter(c => invoiceSettings[c.key] !== false && (invoiceSettings[c.key] === true || c.default)).length} / {allCols.length} sélectionnées
            </span>
          </div>
          <div className="p-5">
            <div className="grid grid-cols-2 sm:grid-cols-3 md:grid-cols-4 lg:grid-cols-6 gap-2">
              {allCols.map(({ label, key, default: def }) => {
                const isOn = invoiceSettings[key] !== undefined ? invoiceSettings[key] : def;
                return (
                  <div key={key}
                    onClick={() => setInvoiceSettings((p: any) => ({ ...p, [key]: !isOn }))}
                    className={`flex items-center gap-2 p-2.5 rounded-xl border-2 cursor-pointer transition-all ${isOn ? 'border-blue-400 bg-blue-50' : 'border-slate-200 bg-slate-50 opacity-60'}`}>
                    <div className={`w-4 h-4 rounded flex items-center justify-center shrink-0 border-2 transition-all ${isOn ? 'bg-blue-600 border-blue-600' : 'border-slate-300'}`}>
                      {isOn && <span className="text-white text-[9px] font-black">✓</span>}
                    </div>
                    <span className={`text-[10px] font-bold leading-tight ${isOn ? 'text-blue-800' : 'text-slate-400'}`}>{label}</span>
                  </div>
                );
              })}
            </div>
            <p className="text-[10px] text-slate-400 mt-3">
              💡 Ces colonnes seront extraites de vos données de facturation et insérées dans le PDF généré.
            </p>
          </div>
        </div>
      );
    })()}

    {/* MODE: IMPORT */}
    {invoiceSettings.setup_mode === 'import' && (
      <div className="max-w-3xl mx-auto space-y-6">

        {/* Step 1: Download starter template */}
        <div className="bg-white rounded-xl border border-slate-200 overflow-hidden">
          <div className="px-5 py-3 bg-slate-50 border-b border-slate-200 flex items-center gap-2">
            <div className="w-6 h-6 rounded-full bg-slate-600 flex items-center justify-center text-[10px] font-black text-white">0</div>
            <span className="text-xs font-black text-slate-700 uppercase tracking-widest">Télécharger le modèle de départ (optionnel)</span>
          </div>
          <div className="p-5 flex items-center justify-between gap-4">
            <div>
              <p className="text-xs text-slate-600 font-bold">Vous n'avez pas de template HTML ?</p>
              <p className="text-[10px] text-slate-400 mt-1">
                Téléchargez notre modèle de base, modifiez-le dans Notepad ou Word, puis re-uploadez-le.
                Les balises <code className="bg-slate-100 px-1 rounded text-blue-600">{`{{placeholders}}`}</code> seront remplacées par vos vraies données.
              </p>
            </div>
            <button onClick={() => {
              const starter = `<!DOCTYPE html>
<html lang="fr">
<head>
<meta charset="UTF-8"/>
<style>
  body { font-family: Arial, sans-serif; font-size: 11px; color: #1e293b; margin: 0; padding: 0; }
  .page { width: 210mm; min-height: 297mm; padding: 15mm; box-sizing: border-box; }
  .header { display: flex; justify-content: space-between; align-items: flex-start; margin-bottom: 20px; padding-bottom: 12px; border-bottom: 3px solid #1e40af; }
  .company-name { font-size: 20px; font-weight: 900; color: #1e40af; }
  .company-info { font-size: 9px; color: #64748b; margin-top: 4px; line-height: 1.6; }
  .invoice-title { font-size: 24px; font-weight: 900; color: #1e40af; text-align: right; }
  .invoice-num { font-size: 13px; font-weight: 700; color: #f59e0b; text-align: right; }
  .invoice-meta { font-size: 9px; color: #64748b; text-align: right; }
  .client-box { background: #f8fafc; border-left: 4px solid #f59e0b; padding: 8px 14px; margin-bottom: 16px; border-radius: 0 6px 6px 0; }
  table { width: 100%; border-collapse: collapse; margin-bottom: 14px; }
  th { background: #1e40af; color: white; padding: 6px 8px; font-size: 8px; text-transform: uppercase; text-align: left; }
  td { padding: 5px 8px; border-bottom: 1px solid #f1f5f9; font-size: 10px; }
  tr:nth-child(even) td { background: #f8fafc; }
  .totals { display: flex; justify-content: flex-end; margin-top: 10px; }
  .totals-box { border: 2px solid #1e40af; border-radius: 6px; padding: 10px 16px; min-width: 220px; }
  .t-row { display: flex; justify-content: space-between; font-size: 11px; padding: 2px 0; }
  .t-total { font-weight: 900; font-size: 14px; color: #1e40af; border-top: 1px solid #e2e8f0; margin-top: 4px; padding-top: 5px; }
  .bank-box { margin-top: 12px; padding: 8px 14px; background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 6px; font-size: 9px; }
  .sig { display: flex; justify-content: flex-end; margin-top: 16px; }
  .sig-box { border: 1px solid #cbd5e1; border-radius: 6px; padding: 8px 16px; min-width: 160px; text-align: center; }
  .sig-label { font-size: 8px; font-weight: 700; color: #64748b; text-transform: uppercase; }
  .sig-space { height: 40px; }
  .footer { margin-top: 16px; border-top: 1px solid #e2e8f0; padding-top: 6px; font-size: 8px; color: #94a3b8; text-align: center; }
  @media print { body { print-color-adjust: exact; -webkit-print-color-adjust: exact; } }
</style>
</head>
<body>
<div class="page">
  <div class="header">
    <div>
      {{company_logo}}
      <div class="company-name">{{company_name}}</div>
      <div class="company-info">
        {{company_address}}<br/>
        Tél: {{company_phone}} — {{company_email}}<br/>
        ICE: {{company_ice}} — RC: {{company_rc}}
      </div>
    </div>
    <div>
      <div class="invoice-title">{{invoice_title}}</div>
      <div class="invoice-num">N° {{numero_facture}}</div>
      <div class="invoice-meta">Date: {{date}}<br/>Délai: {{delai_paiement}} jours</div>
    </div>
  </div>

  <div class="client-box">
    <strong style="font-size:13px">{{client}}</strong>
  </div>

  {{rows}}

  <div class="totals">
    <div class="totals-box">
      <div class="t-row"><span>Total HT</span><span>{{total_ht}} MAD</span></div>
      <div class="t-row"><span>TVA</span><span>{{total_tva}} MAD</span></div>
      <div class="t-row t-total"><span>Total TTC</span><span>{{total_ttc}} MAD</span></div>
    </div>
  </div>

  <div class="bank-box">
    <strong>Coordonnées Bancaires</strong><br/>
    Banque: {{bank_name}} — RIB: {{rib}}
  </div>

  <div class="sig">
    <div class="sig-box">
      <div class="sig-label">{{signature_label}}</div>
      <div class="sig-space"></div>
    </div>
  </div>

  <div class="footer">{{footer_text}}</div>
</div>
</body>
</html>`;
              const blob = new Blob([starter], { type: 'text/html' });
              const a = document.createElement('a');
              a.href = URL.createObjectURL(blob);
              a.download = 'modele_facture_logi-flow.html';
              a.click();
              toast.success("Modèle téléchargé !");
            }}
              className="shrink-0 bg-slate-800 hover:bg-slate-700 text-white px-4 py-2.5 rounded-lg text-xs font-black uppercase tracking-wider flex items-center gap-2 cursor-pointer whitespace-nowrap">
              <Download size={14} /> Télécharger le modèle
            </button>
          </div>
        </div>

        {/* Step 1: Upload HTML template */}
        <div className="bg-white rounded-xl border border-slate-200 overflow-hidden">
          <div className="px-5 py-3 bg-slate-50 border-b border-slate-200 flex items-center gap-2">
            <div className="w-6 h-6 rounded-full bg-blue-600 flex items-center justify-center text-[10px] font-black text-white">1</div>
            <span className="text-xs font-black text-slate-700 uppercase tracking-widest">Uploader votre template HTML</span>
          </div>
          <div className="p-5">
            <p className="text-xs text-slate-500 mb-4">
              Uploadez votre fichier <strong>.html</strong>. Le système remplacera automatiquement les balises
              <code className="bg-slate-100 px-1 mx-1 rounded text-blue-600 text-[10px]">{`{{placeholder}}`}</code>
              par vos vraies données à chaque génération PDF.
            </p>

            {/* Placeholder reference */}
            <div className="mb-4 p-3 bg-blue-50 border border-blue-100 rounded-xl">
              <p className="text-[10px] font-black text-blue-700 uppercase tracking-widest mb-2">Balises disponibles dans votre HTML :</p>
              <div className="grid grid-cols-2 sm:grid-cols-3 gap-1">
                {[
                  '{{company_name}}', '{{company_address}}', '{{company_phone}}',
                  '{{company_email}}', '{{company_ice}}', '{{company_rc}}',
                  '{{company_logo}}', '{{invoice_title}}', '{{numero_facture}}',
                  '{{date}}', '{{client}}', '{{delai_paiement}}',
                  '{{date_paiement}}', '{{montant_ht}}', '{{tva}}',
                  '{{montant_ttc}}', '{{total_ht}}', '{{total_tva}}',
                  '{{total_ttc}}', '{{bl_ot}}', '{{bc}}',
                  '{{rib}}', '{{bank_name}}', '{{footer_text}}',
                  '{{signature_label}}', '{{rows}}',
                ].map(tag => (
                  <code key={tag} className="text-[9px] bg-white border border-blue-200 text-blue-600 px-1.5 py-0.5 rounded font-mono">{tag}</code>
                ))}
              </div>
              <p className="text-[10px] text-blue-500 mt-2">
                ⚡ <strong>{'{{rows}}'}</strong> = le tableau complet avec les colonnes sélectionnées ci-dessus.
              </p>
            </div>

            <label className={`flex items-center justify-center h-24 border-2 border-dashed rounded-xl cursor-pointer transition-all group ${invoiceSettings.invoice_template_html ? 'border-emerald-400 bg-emerald-50' : 'border-slate-300 hover:border-blue-400 hover:bg-blue-50'}`}>
              {invoiceSettings.invoice_template_html ? (
                <div className="flex items-center gap-3">
                  <div className="w-10 h-10 bg-emerald-100 rounded-xl flex items-center justify-center">
                    <FileText size={20} className="text-emerald-600" />
                  </div>
                  <div>
                    <p className="text-sm font-black text-emerald-700">Template HTML chargé ✓</p>
                    <p className="text-[10px] text-emerald-500">Cliquer pour remplacer</p>
                  </div>
                </div>
              ) : (
                <div className="flex flex-col items-center gap-2">
                  <Upload size={20} className="text-slate-400 group-hover:text-blue-500 transition-colors" />
                  <span className="text-sm font-black text-slate-600">Cliquer pour uploader le fichier .html</span>
                </div>
              )}
              <input type="file" accept=".html,.htm"
                onChange={async (e) => {
  const file = e.target.files?.[0];
  if (!file) return;
  let text = await file.text();

  // Auto-inject placeholders by replacing common static patterns
  // Client name area — replace whatever is in .client-name
  text = text.replace(
    /(<[^>]*class="[^"]*client-name[^"]*"[^>]*>)[^<]*/,
    '$1{{client}}'
  );
  // ICE client value
  text = text.replace(
    /(<[^>]*class="[^"]*info-val[^"]*"[^>]*>)\s*\[?\s*ICE\s*CLIENT\s*\]?\s*/i,
    '$1{{client_ice}}'
  );
  // BC number
  text = text.replace(
    /(<[^>]*class="[^"]*info-val[^"]*"[^>]*>)\s*\[?\s*N°?\s*BON\s*DE\s*COMMANDE\s*\]?\s*/i,
    '$1{{bc}}'
  );
  // BL/OT
  text = text.replace(
    /(<[^>]*class="[^"]*info-val[^"]*"[^>]*>)\s*\[?\s*N°?\s*BL[^<]*\]?\s*/i,
    '$1{{bl_ot}}'
  );
  // Règlement mode
  text = text.replace(
    /(<[^>]*class="[^"]*info-val[^"]*"[^>]*>)\s*\[?\s*VIREMENT[^<]*\]?\s*/i,
    '$1{{mode_paiement}}'
  );
  // Invoice number meta-value
  text = text.replace(
    /(<[^>]*class="[^"]*meta-value[^"]*"[^>]*>)\s*\[?\s*AUTO[^<]*\]?\s*/i,
    '$1{{numero_facture}}'
  );
  // Date meta-value
  text = text.replace(
    /(<[^>]*class="[^"]*meta-value[^"]*"[^>]*>)\s*\[?\s*DATE[^<]*\]?\s*/i,
    '$1{{date}}'
  );
  // Total TTC value
  text = text.replace(
    /(<[^>]*class="[^"]*ttc-value[^"]*"[^>]*>)[^<]*/,
    '$1{{total_ttc}} MAD'
  );
  // Total HT
  text = text.replace(
    /(<td[^>]*>)\s*=\s*Σ\s*lignes\s*(<\/td>)/i,
    '$1{{total_ht}} MAD$2'
  );
  // TVA rows — replace formula text
  text = text.replace(
    /(<td[^>]*>)\s*=\s*HT\s*×\s*0,10\s*(<\/td>)/i,
    '$1{{tva}} MAD$2'
  );
  text = text.replace(
    /(<td[^>]*>)\s*=\s*HT\s*×\s*0,20\s*(<\/td>)/i,
    '$1$2'
  );
  // Lettres value
  text = text.replace(
    /(<[^>]*class="[^"]*lettres-value[^"]*"[^>]*>)[^<]*/,
    '$1{{montant_lettres}}'
  );
  // Replace entire tbody with {{rows}}
  text = text.replace(
    /<tbody>[\s\S]*?<\/tbody>/i,
    '<tbody>{{rows}}</tbody>'
  );
  // Page number
  text = text.replace(
    /Page\s*<strong>[^<]*<\/strong>/i,
    'Page <strong>{{page_num}}</strong>'
  );

  setInvoiceSettings((p: any) => ({ ...p, invoice_template_html: text }));
  toast.success("Template HTML importé et converti automatiquement !");
  e.target.value = '';
}}
                className="hidden" />
            </label>

            {invoiceSettings.invoice_template_html && (
              <div className="mt-3 space-y-2">
                <div className="flex items-center justify-between">
                  <p className="text-[10px] font-black text-slate-500 uppercase tracking-widest">Aperçu du code HTML</p>
                  <button onClick={() => setInvoiceSettings((p: any) => ({ ...p, invoice_template_html: '' }))}
                    className="text-[10px] text-rose-500 font-bold hover:underline cursor-pointer">
                    × Supprimer le template
                  </button>
                </div>
                <textarea
                  value={invoiceSettings.invoice_template_html}
                  onChange={e => setInvoiceSettings((p: any) => ({ ...p, invoice_template_html: e.target.value }))}
                  rows={10}
                  className="w-full rounded-lg border-2 border-slate-200 px-3 py-2 text-[10px] font-mono focus:outline-none focus:border-blue-500 resize-y bg-slate-50"
                  placeholder="Votre code HTML ici..."
                />
                <p className="text-[10px] text-slate-400">Vous pouvez modifier directement le HTML ici.</p>
              </div>
            )}
          </div>
        </div>

        {/* Step 2: Logo */}
        <div className="bg-white rounded-xl border border-slate-200 overflow-hidden">
          <div className="px-5 py-3 bg-slate-50 border-b border-slate-200 flex items-center gap-2">
            <div className="w-6 h-6 rounded-full bg-blue-600 flex items-center justify-center text-[10px] font-black text-white">2</div>
            <span className="text-xs font-black text-slate-700 uppercase tracking-widest">Logo de l'entreprise</span>
          </div>
          <div className="p-5">
            <p className="text-xs text-slate-500 mb-4">Format <strong>PNG fond transparent</strong> recommandé. Utilisez la balise <code className="bg-slate-100 px-1 rounded text-blue-600 text-[10px]">{'{{company_logo}}'}</code> dans votre HTML.</p>
            <label className={`flex flex-col items-center justify-center h-32 border-2 border-dashed rounded-xl cursor-pointer transition-all group ${logoPreviewUrl ? 'border-emerald-400 bg-emerald-50' : 'border-slate-300 hover:border-emerald-400 hover:bg-emerald-50'}`}>
              {logoPreviewUrl ? (
                <div className="relative w-full h-full p-2">
                  <img src={logoPreviewUrl} alt="Logo" className="w-full h-full object-contain" />
                  <div className="absolute inset-0 bg-black/40 rounded-xl opacity-0 group-hover:opacity-100 transition-opacity flex items-center justify-center">
                    <span className="text-white text-xs font-black bg-emerald-600 px-3 py-1 rounded-lg">Changer</span>
                  </div>
                </div>
              ) : (
                <div className="flex flex-col items-center gap-2 text-center p-4">
                  {uploadingLogo ? <Loader2 size={24} className="text-emerald-500 animate-spin" />
                    : <div className="w-12 h-12 bg-slate-100 rounded-xl flex items-center justify-center group-hover:bg-emerald-100 transition-colors">
                        <Upload size={20} className="text-slate-400 group-hover:text-emerald-500 transition-colors" />
                      </div>}
                  <span className="text-sm font-black text-slate-600">{uploadingLogo ? 'Upload...' : 'Uploader le logo'}</span>
                  <span className="text-xs text-slate-400">PNG, JPG, SVG</span>
                </div>
              )}
              <input type="file" accept="image/*" onChange={handleLogoUpload} className="hidden" disabled={uploadingLogo} />
            </label>
            {logoPreviewUrl && (
              <button onClick={() => { setLogoPreviewUrl(''); setInvoiceSettings((p: any) => ({ ...p, logo_url: '', logo_storage_path: '' })); }}
                className="mt-2 text-[10px] text-rose-500 font-bold hover:underline cursor-pointer w-full text-center">× Supprimer</button>
            )}
          </div>
        </div>

        {/* Step 3: Pre-printed */}
        <div className="bg-white rounded-xl border border-slate-200 overflow-hidden">
          <div className="px-5 py-3 bg-slate-50 border-b border-slate-200 flex items-center gap-2">
            <div className="w-6 h-6 rounded-full bg-amber-500 flex items-center justify-center text-[10px] font-black text-white">3</div>
            <span className="text-xs font-black text-slate-700 uppercase tracking-widest">Papier pré-imprimé ?</span>
          </div>
          <div className="p-5 space-y-3">
            <p className="text-xs text-slate-500">Si votre papier a déjà un en-tête ou pied de page imprimé, activez ci-dessous pour éviter les doublons.</p>
            {[
              { label: 'En-tête déjà imprimé',      sub: 'Logo, nom, adresse déjà sur le papier en haut',        key: 'skip_header' },
              { label: 'Pied de page déjà imprimé', sub: 'Mentions légales, signature déjà sur le papier en bas', key: 'skip_footer' },
            ].map(({ label, sub, key }) => (
              <div key={key} onClick={() => setInvoiceSettings((p: any) => ({ ...p, [key]: !p[key] }))}
                className={`flex items-center justify-between p-4 rounded-xl border-2 cursor-pointer transition-all ${invoiceSettings[key] ? 'border-amber-400 bg-amber-50' : 'border-slate-200 hover:border-amber-300'}`}>
                <div>
                  <p className={`text-sm font-black ${invoiceSettings[key] ? 'text-amber-800' : 'text-slate-700'}`}>{label}</p>
                  <p className="text-xs text-slate-400 mt-0.5">{sub}</p>
                </div>
                <div className={`w-12 h-7 rounded-full flex items-center shrink-0 ml-4 transition-all ${invoiceSettings[key] ? 'bg-amber-500' : 'bg-slate-300'}`}>
                  <div className={`w-5 h-5 bg-white rounded-full shadow mx-1 transition-all ${invoiceSettings[key] ? 'translate-x-5' : 'translate-x-0'}`} />
                </div>
              </div>
            ))}
          </div>
        </div>

        {/* Test PDF button */}
        {invoiceSettings.invoice_template_html && (
          <div className="bg-blue-50 border border-blue-100 rounded-xl p-4 flex items-center justify-between">
            <div>
              <p className="text-sm font-black text-blue-800">Tester le template</p>
              <p className="text-[10px] text-blue-500 mt-0.5">Génère un PDF de test avec des données fictives pour vérifier le rendu.</p>
            </div>
            <button onClick={() => {
              const testData = {
                company_name: invoiceSettings.company_name || 'VOTRE ENTREPRISE',
                company_address: invoiceSettings.address || 'Rue X, Casablanca',
                company_phone: invoiceSettings.phone || '+212 6 00 00 00 00',
                company_email: invoiceSettings.email || 'contact@entreprise.ma',
                company_ice: invoiceSettings.ice || '001234567000000',
                company_rc: invoiceSettings.rc || '123456',
                company_logo: invoiceSettings.logo_url ? `<img src="${invoiceSettings.logo_url}" style="max-height:60px;max-width:150px;object-fit:contain"/>` : '',
                invoice_title: invoiceSettings.invoice_title || 'FACTURE',
                numero_facture: 'TEST-001',
                date: new Date().toLocaleDateString('fr-MA'),
                client: 'CLIENT TEST SARL',
                delai_paiement: '60',
                date_paiement: '',
                montant_ht: '10,000',
                tva: '1,000',
                montant_ttc: '11,000',
                total_ht: '10,000',
                total_tva: '1,000',
                total_ttc: '11,000',
                bl_ot: 'BL-001',
                bc: 'BC-123',
                rib: invoiceSettings.rib || 'XXX XXX XXXXXXXXXXXXXXX XX',
                bank_name: invoiceSettings.bank_name || 'Attijariwafa Bank',
                footer_text: invoiceSettings.footer_text || 'Merci de votre confiance.',
                signature_label: invoiceSettings.signature_label || 'Signature & Cachet',
                rows: `<table style="width:100%;border-collapse:collapse">
                  <thead><tr style="background:#1e40af;color:white">
                    <th style="padding:6px;font-size:9px;text-align:left">Date</th>
                    <th style="padding:6px;font-size:9px;text-align:left">N° Fact</th>
                    <th style="padding:6px;font-size:9px;text-align:left">Client</th>
                    <th style="padding:6px;font-size:9px;text-align:right">HT (MAD)</th>
                    <th style="padding:6px;font-size:9px;text-align:right">TTC (MAD)</th>
                  </tr></thead>
                  <tbody>
                    <tr><td style="padding:5px 6px;border-bottom:1px solid #f1f5f9;font-size:10px">01/06/2026</td><td style="padding:5px 6px;border-bottom:1px solid #f1f5f9;font-size:10px">F-001</td><td style="padding:5px 6px;border-bottom:1px solid #f1f5f9;font-size:10px">CLIENT TEST</td><td style="padding:5px 6px;border-bottom:1px solid #f1f5f9;font-size:10px;text-align:right">5,000 MAD</td><td style="padding:5px 6px;border-bottom:1px solid #f1f5f9;font-size:10px;text-align:right">5,500 MAD</td></tr>
                    <tr style="background:#f8fafc"><td style="padding:5px 6px;border-bottom:1px solid #f1f5f9;font-size:10px">02/06/2026</td><td style="padding:5px 6px;border-bottom:1px solid #f1f5f9;font-size:10px">F-002</td><td style="padding:5px 6px;border-bottom:1px solid #f1f5f9;font-size:10px">CLIENT TEST</td><td style="padding:5px 6px;border-bottom:1px solid #f1f5f9;font-size:10px;text-align:right">5,000 MAD</td><td style="padding:5px 6px;border-bottom:1px solid #f1f5f9;font-size:10px;text-align:right">5,500 MAD</td></tr>
                  </tbody>
                </table>`,
              };
              let html = invoiceSettings.invoice_template_html;
              Object.entries(testData).forEach(([k, v]) => {
                html = html.replaceAll(`{{${k}}}`, v as string);
              });
              const win = window.open('', '_blank');
              if (win) { win.document.write(html); win.document.close(); win.focus(); setTimeout(() => win.print(), 500); }
            }}
              className="bg-blue-600 hover:bg-blue-700 text-white px-4 py-2 rounded-lg text-xs font-black uppercase cursor-pointer flex items-center gap-2">
              <Eye size={14} /> Tester le PDF
            </button>
          </div>
        )}

        {/* Save */}
        <div className="flex justify-end pb-6">
          <button onClick={handleSaveInvoiceSettings} disabled={savingSettings}
            className="bg-blue-600 hover:bg-blue-700 disabled:opacity-60 text-white px-8 py-3 rounded-xl text-sm font-black uppercase tracking-wider flex items-center gap-2 cursor-pointer shadow-lg">
            {savingSettings ? <Loader2 size={16} className="animate-spin" /> : null}
            {savingSettings ? 'Sauvegarde...' : '✓ Sauvegarder'}
          </button>
        </div>
      </div>
    )}

    {/* MODE: CREATE */}
    {invoiceSettings.setup_mode === 'create' && (
      <div className="flex gap-6">
        <div className="flex-1 min-w-0 space-y-5">

          {/* BLOCK 1: Company Identity */}
          <div className="bg-white rounded-xl border border-slate-200 overflow-hidden">
            <div className="px-5 py-3 bg-slate-50 border-b border-slate-200 flex items-center gap-2">
              <div className="w-6 h-6 rounded-full bg-blue-600 flex items-center justify-center text-[10px] font-black text-white">1</div>
              <span className="text-xs font-black text-slate-700 uppercase tracking-widest">Identité de l'entreprise</span>
            </div>
            <div className="p-5">
              <div className="grid grid-cols-2 gap-3">
                {[
                  { label: 'Raison sociale', key: 'company_name', placeholder: 'FOTRAL SARL'       },
                  { label: 'Adresse',        key: 'address',      placeholder: 'Rue X, Casablanca' },
                  { label: 'Téléphone',      key: 'phone',        placeholder: '+212 6...'          },
                  { label: 'Email',          key: 'email',        placeholder: 'contact@...'        },
                  { label: 'ICE',            key: 'ice',          placeholder: '001234567000000'    },
                  { label: 'RC',             key: 'rc',           placeholder: '123456'             },
                ].map(({ label, key, placeholder }) => (
                  <div key={key}>
                    <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">{label}</label>
                    <input type="text" placeholder={placeholder} value={invoiceSettings[key] || ''}
                      onChange={e => setInvoiceSettings((p: any) => ({ ...p, [key]: e.target.value }))}
                      className="w-full mt-1 h-8 rounded-lg border-2 border-slate-200 px-3 text-xs focus:outline-none focus:border-blue-500" />
                  </div>
                ))}
              </div>
              <div className="mt-4 p-3 bg-slate-50 rounded-xl border-2 border-dashed border-slate-300 flex items-center gap-4">
                {logoPreviewUrl ? (
                  <div className="relative shrink-0">
                    <img src={logoPreviewUrl} alt="Logo" className="h-12 w-auto object-contain rounded border border-slate-200 bg-white p-1" />
                    <button onClick={() => { setLogoPreviewUrl(''); setInvoiceSettings((p: any) => ({ ...p, logo_url: '', logo_storage_path: '' })); }}
                      className="absolute -top-1.5 -right-1.5 w-4 h-4 bg-rose-500 text-white rounded-full flex items-center justify-center cursor-pointer">
                      <X size={8} />
                    </button>
                  </div>
                ) : (
                  <div className="w-12 h-12 bg-slate-200 rounded-lg flex items-center justify-center text-slate-400 text-[9px] font-bold shrink-0">LOGO</div>
                )}
                <div>
                  <p className="text-xs font-black text-slate-700">Logo</p>
                  <p className="text-[10px] text-slate-400">PNG fond transparent recommandé</p>
                  <label className={`mt-1 inline-flex items-center gap-1 px-2 py-1 rounded text-[10px] font-black uppercase cursor-pointer ${uploadingLogo ? 'bg-slate-200 text-slate-400' : 'bg-blue-600 text-white hover:bg-blue-700'}`}>
                    {uploadingLogo ? <Loader2 size={9} className="animate-spin" /> : <Upload size={9} />}
                    {uploadingLogo ? 'Upload...' : 'Choisir'}
                    <input type="file" accept="image/*" onChange={handleLogoUpload} className="hidden" disabled={uploadingLogo} />
                  </label>
                </div>
              </div>
            </div>
          </div>

          {/* BLOCK 2: Style */}
          <div className="bg-white rounded-xl border border-slate-200 overflow-hidden">
            <div className="px-5 py-3 bg-slate-50 border-b border-slate-200 flex items-center gap-2">
              <div className="w-6 h-6 rounded-full bg-violet-600 flex items-center justify-center text-[10px] font-black text-white">2</div>
              <span className="text-xs font-black text-slate-700 uppercase tracking-widest">Style</span>
            </div>
            <div className="p-5 space-y-4">
              <div>
                <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Titre</label>
                <div className="grid grid-cols-3 gap-2 mt-1">
                  {['FACTURE', 'AVOIR', 'DEVIS'].map(t => (
                    <button key={t} onClick={() => setInvoiceSettings((p: any) => ({ ...p, invoice_title: t }))}
                      className={`h-8 rounded-lg text-xs font-black uppercase border-2 transition-all cursor-pointer ${invoiceSettings.invoice_title === t ? 'bg-blue-600 border-blue-600 text-white' : 'border-slate-200 text-slate-600 hover:border-blue-300'}`}>
                      {t}
                    </button>
                  ))}
                </div>
              </div>
              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Couleur principale</label>
                  <div className="flex items-center gap-2 mt-1">
                    <input type="color" value={invoiceSettings.primary_color || '#1e40af'}
                      onChange={e => setInvoiceSettings((p: any) => ({ ...p, primary_color: e.target.value }))}
                      className="h-8 w-10 rounded cursor-pointer border-2 border-slate-200" />
                    <span className="text-[10px] font-mono text-slate-500">{invoiceSettings.primary_color}</span>
                  </div>
                </div>
                <div>
                  <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Couleur accent</label>
                  <div className="flex items-center gap-2 mt-1">
                    <input type="color" value={invoiceSettings.accent_color || '#f59e0b'}
                      onChange={e => setInvoiceSettings((p: any) => ({ ...p, accent_color: e.target.value }))}
                      className="h-8 w-10 rounded cursor-pointer border-2 border-slate-200" />
                    <span className="text-[10px] font-mono text-slate-500">{invoiceSettings.accent_color}</span>
                  </div>
                </div>
              </div>
              <div>
                <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">
                  Taille police — <span className="text-slate-600">{invoiceSettings.font_size || 11}px</span>
                </label>
                <input type="range" min="8" max="14" value={invoiceSettings.font_size || 11}
                  onChange={e => setInvoiceSettings((p: any) => ({ ...p, font_size: parseInt(e.target.value) }))}
                  className="w-full mt-1 accent-blue-600" />
              </div>
            </div>
          </div>

          {/* BLOCK 3: Pre-printed */}
          <div className="bg-white rounded-xl border border-slate-200 overflow-hidden">
            <div className="px-5 py-3 bg-slate-50 border-b border-slate-200 flex items-center gap-2">
              <div className="w-6 h-6 rounded-full bg-amber-500 flex items-center justify-center text-[10px] font-black text-white">3</div>
              <span className="text-xs font-black text-slate-700 uppercase tracking-widest">Papier pré-imprimé</span>
            </div>
            <div className="p-5 space-y-3">
              {[
                { label: 'En-tête déjà imprimé',      sub: 'PDF sans logo/nom/adresse en haut',          key: 'skip_header' },
                { label: 'Pied de page déjà imprimé', sub: 'PDF sans mentions légales/signature en bas', key: 'skip_footer' },
              ].map(({ label, sub, key }) => (
                <div key={key} onClick={() => setInvoiceSettings((p: any) => ({ ...p, [key]: !p[key] }))}
                  className={`flex items-center justify-between p-3 rounded-xl border-2 cursor-pointer transition-all ${invoiceSettings[key] ? 'border-amber-400 bg-amber-50' : 'border-slate-200 hover:border-amber-300'}`}>
                  <div>
                    <p className={`text-xs font-black ${invoiceSettings[key] ? 'text-amber-800' : 'text-slate-700'}`}>{label}</p>
                    <p className="text-[10px] text-slate-400 mt-0.5">{sub}</p>
                  </div>
                  <div className={`w-10 h-6 rounded-full flex items-center shrink-0 ml-3 transition-all ${invoiceSettings[key] ? 'bg-amber-500' : 'bg-slate-300'}`}>
                    <div className={`w-4 h-4 bg-white rounded-full shadow mx-1 transition-all ${invoiceSettings[key] ? 'translate-x-4' : 'translate-x-0'}`} />
                  </div>
                </div>
              ))}
            </div>
          </div>

          {/* BLOCK 4: Margins */}
          <div className="bg-white rounded-xl border border-slate-200 overflow-hidden">
            <div className="px-5 py-3 bg-slate-50 border-b border-slate-200 flex items-center gap-2">
              <div className="w-6 h-6 rounded-full bg-emerald-600 flex items-center justify-center text-[10px] font-black text-white">4</div>
              <span className="text-xs font-black text-slate-700 uppercase tracking-widest">Marges & Mise en page</span>
            </div>
            <div className="p-5 space-y-4">
              <div className="grid grid-cols-2 gap-3">
                {[
                  { label: 'Marge Haut',   key: 'margin_top'    },
                  { label: 'Marge Bas',    key: 'margin_bottom' },
                  { label: 'Marge Gauche', key: 'margin_left'   },
                  { label: 'Marge Droite', key: 'margin_right'  },
                ].map(({ label, key }) => (
                  <div key={key}>
                    <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">
                      {label} — <span className="text-slate-600">{invoiceSettings[key] || 15}mm</span>
                    </label>
                    <input type="range" min="5" max="40" value={invoiceSettings[key] || 15}
                      onChange={e => setInvoiceSettings((p: any) => ({ ...p, [key]: parseInt(e.target.value) }))}
                      className="w-full mt-1 accent-emerald-600" />
                  </div>
                ))}
              </div>
              <div>
                <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">
                  Lignes par page — <span className="text-slate-600">{invoiceSettings.rows_per_page || 15}</span>
                </label>
                <input type="range" min="5" max="30" value={invoiceSettings.rows_per_page || 15}
                  onChange={e => setInvoiceSettings((p: any) => ({ ...p, rows_per_page: parseInt(e.target.value) }))}
                  className="w-full mt-1 accent-emerald-600" />
              </div>
              <div>
                <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest mb-2 block">Largeur colonnes (%)</label>
                <div className="space-y-2">
                  {[
                    { label: 'Date',     key: 'col_width_date'    },
                    { label: 'N° Fact',  key: 'col_width_fact'    },
                    { label: 'Client',   key: 'col_width_client'  },
                    { label: 'Route',    key: 'col_width_route'   },
                    { label: 'Montants', key: 'col_width_amounts' },
                  ].map(({ label, key }) => (
                    <div key={key} className="flex items-center gap-3">
                      <span className="text-[10px] text-slate-500 w-16 shrink-0">{label}</span>
                      <input type="range" min="5" max="30" value={invoiceSettings[key] || 12}
                        onChange={e => setInvoiceSettings((p: any) => ({ ...p, [key]: parseInt(e.target.value) }))}
                        className="flex-1 accent-blue-500" />
                      <span className="text-[10px] font-black text-slate-600 w-7 text-right">{invoiceSettings[key] || 12}%</span>
                    </div>
                  ))}
                </div>
              </div>
            </div>
          </div>

          {/* BLOCK 5: Payment & Footer */}
          <div className="bg-white rounded-xl border border-slate-200 overflow-hidden">
            <div className="px-5 py-3 bg-slate-50 border-b border-slate-200 flex items-center gap-2">
              <div className="w-6 h-6 rounded-full bg-rose-500 flex items-center justify-center text-[10px] font-black text-white">5</div>
              <span className="text-xs font-black text-slate-700 uppercase tracking-widest">Paiement & Pied de page</span>
            </div>
            <div className="p-5 space-y-3">
              {[
                { label: 'RIB',             key: 'rib',             placeholder: '123 456 789...'    },
                { label: 'Banque',          key: 'bank_name',       placeholder: 'Attijariwafa Bank' },
                { label: 'Label Signature', key: 'signature_label', placeholder: 'Signature & Cachet'},
              ].map(({ label, key, placeholder }) => (
                <div key={key}>
                  <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">{label}</label>
                  <input type="text" placeholder={placeholder} value={invoiceSettings[key] || ''}
                    onChange={e => setInvoiceSettings((p: any) => ({ ...p, [key]: e.target.value }))}
                    className="w-full mt-1 h-8 rounded-lg border-2 border-slate-200 px-3 text-xs focus:outline-none focus:border-blue-500" />
                </div>
              ))}
              <div>
                <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Mentions légales</label>
                <textarea value={invoiceSettings.footer_text || ''}
                  onChange={e => setInvoiceSettings((p: any) => ({ ...p, footer_text: e.target.value }))}
                  rows={3} placeholder="Ex: Tout retard de paiement entraîne des pénalités..."
                  className="w-full mt-1 rounded-lg border-2 border-slate-200 px-3 py-2 text-xs focus:outline-none focus:border-blue-500 resize-none" />
              </div>
            </div>
          </div>

          {/* Save */}
          <div className="flex justify-end pb-6">
            <button onClick={handleSaveInvoiceSettings} disabled={savingSettings}
              className="bg-violet-600 hover:bg-violet-700 disabled:opacity-60 text-white px-8 py-3 rounded-xl text-sm font-black uppercase tracking-wider flex items-center gap-2 cursor-pointer shadow-lg">
              {savingSettings ? <Loader2 size={16} className="animate-spin" /> : null}
              {savingSettings ? 'Sauvegarde...' : '✓ Sauvegarder le modèle'}
            </button>
          </div>
        </div>

        {/* RIGHT: Live Preview */}
        <div className="w-[400px] shrink-0">
          <div className="sticky top-20">
            <div className="flex items-center justify-between mb-3">
              <span className="text-xs font-black text-slate-500 uppercase tracking-widest">Aperçu en direct</span>
              <span className="text-[10px] text-slate-400 bg-slate-100 px-2 py-0.5 rounded">Temps réel</span>
            </div>
            <div className="rounded-xl overflow-hidden border-2 shadow-lg"
              style={{ borderColor: invoiceSettings.primary_color || '#1e40af' }}>
              {!invoiceSettings.skip_header && (
                <div className="p-3 border-b-2" style={{ borderColor: invoiceSettings.primary_color, backgroundColor: (invoiceSettings.primary_color || '#1e40af') + '12' }}>
                  <div className="flex justify-between items-start">
                    <div className="flex items-center gap-2">
                      {logoPreviewUrl && <img src={logoPreviewUrl} alt="Logo" className="h-8 w-auto object-contain" />}
                      <div>
                        <p className="font-black text-xs" style={{ color: invoiceSettings.primary_color }}>
                          {invoiceSettings.company_name || 'Nom entreprise'}
                        </p>
                        {invoiceSettings.address && <p className="text-[9px] text-slate-500">{invoiceSettings.address}</p>}
                        {invoiceSettings.ice     && <p className="text-[9px] text-slate-500">ICE: {invoiceSettings.ice}</p>}
                      </div>
                    </div>
                    <div className="text-right">
                      <p className="font-black text-sm" style={{ color: invoiceSettings.primary_color }}>
                        {invoiceSettings.invoice_title || 'FACTURE'}
                      </p>
                      <p className="text-[9px]" style={{ color: invoiceSettings.accent_color }}>N° 00001</p>
                      <p className="text-[9px] text-slate-400">{new Date().toLocaleDateString('fr-MA')}</p>
                    </div>
                  </div>
                </div>
              )}
              {invoiceSettings.skip_header && (
                <div className="p-2 text-center text-[9px] text-amber-600 bg-amber-50 font-bold border-b border-amber-200">
                  ↑ EN-TÊTE NON IMPRIMÉ
                </div>
              )}
              <div className="mx-3 my-2 p-2 bg-slate-50 rounded text-xs"
                style={{ borderLeft: `3px solid ${invoiceSettings.accent_color || '#f59e0b'}` }}>
                <p className="font-bold text-slate-700 text-[10px]">NOM DU CLIENT</p>
                <p className="text-[9px] text-slate-400">X prestation(s)</p>
              </div>
              <div className="mx-3 mb-2 rounded-lg overflow-hidden border border-slate-200">
                <div className="flex text-[8px] font-black text-white px-2 py-1 gap-1"
                  style={{ backgroundColor: invoiceSettings.primary_color }}>
                  {[
                    { label: 'Date',    key: 'col_show_date',   def: true  },
                    { label: 'N° Fact', key: 'col_show_fact',   def: true  },
                    { label: 'Client',  key: 'col_show_client', def: true  },
                    { label: 'HT',      key: 'col_show_ht',     def: true  },
                    { label: 'TTC',     key: 'col_show_ttc',    def: true  },
                    { label: 'Statut',  key: 'col_show_statut', def: true  },
                  ].filter(c => invoiceSettings[c.key] !== false && (invoiceSettings[c.key] === true || c.def))
                   .map(c => <span key={c.key} className="flex-1">{c.label}</span>)}
                </div>
                {[1,2,3].map(i => (
                  <div key={i} className={`flex text-[8px] text-slate-400 px-2 py-1 gap-1 ${i%2===0?'bg-slate-50':'bg-white'}`}>
                    <span className="flex-1">01/06</span>
                    <span className="flex-1">F-00{i}</span>
                    <span className="flex-1">Client</span>
                    <span className="flex-1">4,400</span>
                    <span className="flex-1">4,840</span>
                    <span className="flex-1">✓</span>
                  </div>
                ))}
              </div>
              <div className="flex justify-end mx-3 mb-2">
                <div className="border-2 rounded-lg p-2 min-w-[120px]" style={{ borderColor: invoiceSettings.primary_color }}>
                  <div className="flex justify-between text-[8px] text-slate-400 mb-0.5"><span>Total HT</span><span>X MAD</span></div>
                  <div className="flex justify-between text-[8px] text-slate-400 mb-0.5"><span>TVA</span><span>X MAD</span></div>
                  <div className="flex justify-between text-[9px] font-black border-t pt-0.5" style={{ color: invoiceSettings.primary_color }}>
                    <span>TTC</span><span>X MAD</span>
                  </div>
                </div>
              </div>
              {(invoiceSettings.rib || invoiceSettings.bank_name) && (
                <div className="mx-3 mb-2 p-1.5 bg-slate-50 rounded text-[8px] text-slate-500 border border-slate-200">
                  {invoiceSettings.bank_name && <span className="font-bold">{invoiceSettings.bank_name} </span>}
                  {invoiceSettings.rib && <span>RIB: {invoiceSettings.rib}</span>}
                </div>
              )}
              <div className="flex justify-end mx-3 mb-2">
                <div className="border border-slate-200 rounded p-1.5 text-center min-w-[90px]">
                  <p className="text-[8px] text-slate-400 uppercase font-bold">{invoiceSettings.signature_label || 'Signature'}</p>
                  <div className="h-5" />
                </div>
              </div>
              {!invoiceSettings.skip_footer && invoiceSettings.footer_text && (
                <div className="mx-3 mb-2 text-[8px] text-slate-400 border-t border-slate-200 pt-1.5 text-center">
                  {invoiceSettings.footer_text}
                </div>
              )}
              {invoiceSettings.skip_footer && (
                <div className="p-2 text-center text-[9px] text-amber-600 bg-amber-50 font-bold border-t border-amber-200">
                  ↓ PIED DE PAGE NON IMPRIMÉ
                </div>
              )}
            </div>
            <p className="text-[10px] text-slate-400 text-center mt-2">Aperçu mis à jour en temps réel</p>
          </div>
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
                  className="p-1.5 rounded hover:bg-slate-100 text-slate-400"><X size={16} /></button>
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
                <button onClick={handleSaveSuivi} className="flex-1 h-10 bg-blue-600 hover:bg-blue-700 text-white text-sm font-bold rounded-lg cursor-pointer">
                  {editingSuivi ? 'Enregistrer les modifications' : 'Ajouter la prestation'}
                </button>
                <button onClick={() => { setShowSuiviForm(false); setEditingSuivi(null); setSuiviForm(emptySuivi); }}
                  className="flex-1 h-10 bg-slate-100 hover:bg-slate-200 text-slate-700 text-sm font-bold rounded-lg cursor-pointer">Annuler</button>
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

      {/* Edit Topup Modal */}
      <AnimatePresence>
        {editingTopup && (
          <div className="fixed inset-0 bg-black/50 z-50 flex items-center justify-center p-4">
            <motion.div initial={{ scale: 0.95, opacity: 0 }} animate={{ scale: 1, opacity: 1 }} exit={{ scale: 0.95, opacity: 0 }}
              className="bg-white rounded-xl p-6 max-w-sm w-full shadow-xl space-y-4">
              <h3 className="font-black text-slate-900 uppercase tracking-widest text-sm">Modifier le Versement</h3>
              <div>
                <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Montant (MAD)</label>
                <input type="number" value={topupEditAmount} onChange={e => setTopupEditAmount(e.target.value)}
                  className="w-full mt-1 h-9 rounded-lg border-2 border-slate-200 px-3 text-sm focus:outline-none focus:border-blue-500" />
              </div>
              <div>
                <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Note</label>
                <input type="text" value={topupEditNotes} onChange={e => setTopupEditNotes(e.target.value)}
                  className="w-full mt-1 h-9 rounded-lg border-2 border-slate-200 px-3 text-sm focus:outline-none focus:border-blue-500" />
              </div>
              <div className="flex gap-3 pt-2">
                <button onClick={handleEditTopupSave} className="flex-1 h-10 bg-blue-600 hover:bg-blue-700 text-white text-sm font-bold rounded-lg cursor-pointer">Enregistrer</button>
                <button onClick={() => setEditingTopup(null)} className="flex-1 h-10 bg-slate-100 hover:bg-slate-200 text-slate-700 text-sm font-bold rounded-lg cursor-pointer">Annuler</button>
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
      {/* Edit Driver Modal */}
<AnimatePresence>
  {editingDriver && (
    <div className="fixed inset-0 bg-black/50 z-50 flex items-center justify-center p-4">
      <motion.div initial={{ scale: 0.95, opacity: 0 }} animate={{ scale: 1, opacity: 1 }}
        exit={{ scale: 0.95, opacity: 0 }}
        className="bg-white rounded-xl p-6 max-w-2xl w-full shadow-xl max-h-[90vh] overflow-y-auto">
        <div className="flex items-center justify-between mb-5">
          <h3 className="font-black text-slate-900 uppercase tracking-widest text-sm">Modifier le Chauffeur</h3>
          <button onClick={() => setEditingDriver(null)} className="p-1.5 rounded hover:bg-slate-100 text-slate-400"><X size={16} /></button>
        </div>
        <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
          {[
            { label: 'Code',                key: 'code',                type: 'text'   },
            { label: 'Nom / Prénom',         key: 'nom_prenom',          type: 'text'   },
            { label: 'Immatriculation',      key: 'immatriculation',     type: 'text'   },
            { label: 'Type Véhicule',        key: 'type_vehicule',       type: 'text'   },
            { label: 'CIN',                  key: 'cin',                 type: 'text'   },
            { label: 'IMM CNSS',             key: 'imm_cnss',            type: 'text'   },
            { label: 'Fonction',             key: 'fonction',            type: 'text'   },
            { label: 'Date de Naissance',    key: 'date_naissance',      type: 'text'   },
            { label: 'Situation Familiale',  key: 'situation_familiale', type: 'text'   },
            { label: 'Nombre de Déductions', key: 'nb_deduction',        type: 'number' },
            { label: 'Date d\'Embauche',     key: 'date_embauche',       type: 'text'   },
            { label: 'Adresse',              key: 'adresse',             type: 'text'   },
          ].map(({ label, key, type }) => (
            <div key={key}>
              <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">{label}</label>
              <input type={type} value={driverEditForm[key] || ''}
                onChange={e => setDriverEditForm((p: any) => ({ ...p, [key]: e.target.value }))}
                className="w-full mt-1 h-9 rounded-lg border-2 border-slate-200 px-3 text-sm focus:outline-none focus:border-blue-500" />
            </div>
          ))}
        </div>
        <div className="flex gap-3 pt-5">
          <button onClick={handleEditDriverSave}
            className="flex-1 h-10 bg-blue-600 hover:bg-blue-700 text-white text-sm font-bold rounded-lg cursor-pointer">
            Enregistrer
          </button>
          <button onClick={() => setEditingDriver(null)}
            className="flex-1 h-10 bg-slate-100 hover:bg-slate-200 text-slate-700 text-sm font-bold rounded-lg cursor-pointer">
            Annuler
          </button>
        </div>
      </motion.div>
    </div>
  )}
</AnimatePresence>
<AnimatePresence>
  {editingClient && (
    <div className="fixed inset-0 bg-black/50 z-50 flex items-center justify-center p-4">
      <motion.div initial={{ scale: 0.95, opacity: 0 }} animate={{ scale: 1, opacity: 1 }}
        exit={{ scale: 0.95, opacity: 0 }}
        className="bg-white rounded-xl p-6 max-w-lg w-full shadow-xl space-y-4">
        <div className="flex items-center justify-between mb-2">
          <h3 className="font-black text-slate-900 uppercase tracking-widest text-sm">Modifier le Client</h3>
          <button onClick={() => setEditingClient(null)} className="p-1.5 rounded hover:bg-slate-100 text-slate-400"><X size={16} /></button>
        </div>
        {[
          { label: 'Nom Client',         key: 'nom',            type: 'text'   },
          { label: 'Adresse',            key: 'adresse',        type: 'text'   },
          { label: 'ICE',                key: 'ice',            type: 'text'   },
          { label: 'Délai Paiement (J)', key: 'delai_paiement', type: 'number' },
        ].map(({ label, key, type }) => (
          <div key={key}>
            <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">{label}</label>
            <input type={type} value={clientEditForm[key] || ''}
              onChange={e => setClientEditForm((p: any) => ({ ...p, [key]: e.target.value }))}
              className="w-full mt-1 h-9 rounded-lg border-2 border-slate-200 px-3 text-sm focus:outline-none focus:border-blue-500" />
          </div>
        ))}
        <div className="flex gap-3 pt-2">
          <button onClick={handleEditClientSave}
            className="flex-1 h-10 bg-blue-600 hover:bg-blue-700 text-white text-sm font-bold rounded-lg cursor-pointer">
            Enregistrer
          </button>
          <button onClick={() => setEditingClient(null)}
            className="flex-1 h-10 bg-slate-100 hover:bg-slate-200 text-slate-700 text-sm font-bold rounded-lg cursor-pointer">
            Annuler
          </button>
        </div>
      </motion.div>
    </div>
  )}
</AnimatePresence>
{/* Prestation Picker */}
<AnimatePresence>
  {prestationPickerOpen && (
    <div className="fixed inset-0 bg-black/50 z-50 flex items-center justify-center p-4">
      <motion.div initial={{ scale: 0.95, opacity: 0 }} animate={{ scale: 1, opacity: 1 }}
        exit={{ scale: 0.95, opacity: 0 }}
        className="bg-white rounded-xl p-6 max-w-4xl w-full shadow-xl max-h-[80vh] overflow-y-auto">
        <div className="flex items-center justify-between mb-4">
          <h3 className="font-black text-slate-900 uppercase tracking-widest text-sm">Sélectionner une Prestation</h3>
          <button onClick={() => setPrestationPickerOpen(false)} className="p-1.5 rounded hover:bg-slate-100 text-slate-400"><X size={16} /></button>
        </div>
        <p className="text-xs text-slate-500 mb-4">Cliquez sur une prestation pour auto-remplir la facture.</p>
        <table className="w-full text-left text-xs">
          <thead className="bg-slate-50 border-b border-slate-200">
            <tr>
              {['Date','Matricule','Client','Départ','Arrivée','Prix HT','Prix TTC','BL/OT'].map(h => (
                <th key={h} className="px-3 py-2 text-[9px] font-black text-slate-400 uppercase tracking-widest">{h}</th>
              ))}
            </tr>
          </thead>
          <tbody className="divide-y divide-slate-100">
            {suiviList.length === 0 ? (
              <tr><td colSpan={8} className="px-4 py-6 text-center text-slate-400">Aucune prestation disponible.</td></tr>
            ) : suiviList.map(p => (
              <tr key={p.id}
                onClick={() => handleAutoFillFromPrestation(p)}
                className="hover:bg-blue-50 cursor-pointer transition-colors">
                <td className="px-3 py-2">{p.date}</td>
                <td className="px-3 py-2 font-mono text-blue-600">{p.matricule || '—'}</td>
                <td className="px-3 py-2 font-semibold">{p.client || '—'}</td>
                <td className="px-3 py-2">{p.depart || '—'}</td>
                <td className="px-3 py-2">{p.arrivee || '—'}</td>
                <td className="px-3 py-2">{Number(p.prix_ht).toLocaleString('fr-MA')}</td>
                <td className="px-3 py-2 font-bold">{Number(p.prix_ttc).toLocaleString('fr-MA')}</td>
                <td className="px-3 py-2">{p.ot_bl_bs_be || '—'}</td>
              </tr>
            ))}
          </tbody>
        </table>
        <div className="mt-4 pt-4 border-t border-slate-100">
          <button onClick={() => { setShowFactForm(true); setPrestationPickerOpen(false); setFactForm(emptyFactForm); }}
            className="text-xs text-blue-600 font-bold hover:underline cursor-pointer">
            → Créer une facture sans prestation
          </button>
        </div>
      </motion.div>
    </div>
  )}
</AnimatePresence>

{/* Facture Form Modal */}
<AnimatePresence>
  {showFactForm && (
  <div className="fixed inset-0 bg-black/50 z-50 flex items-center justify-center p-4">
    <motion.div initial={{ scale: 0.95, opacity: 0 }} animate={{ scale: 1, opacity: 1 }}
      exit={{ scale: 0.95, opacity: 0 }}
      className="bg-white rounded-xl p-6 max-w-3xl w-full shadow-xl max-h-[90vh] overflow-y-auto">
      <div className="flex items-center justify-between mb-5">
        <h3 className="font-black text-slate-900 uppercase tracking-widest text-sm">
          {editingFact ? 'Modifier la Facture' : 'Nouvelle Facture'}
        </h3>
        <button onClick={() => { setShowFactForm(false); setEditingFact(null); }}
          className="p-1.5 rounded hover:bg-slate-100 text-slate-400"><X size={16} /></button>
      </div>
      <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-4">

        {/* Text/date fields */}
        {[
          { label: 'Date',                 key: 'date',                  type: 'date'   },
          { label: 'N° Facture',           key: 'numero_facture',        type: 'text'   },
          { label: 'Client',               key: 'client',                type: 'text'   },
          { label: 'Départ',               key: 'depart',                type: 'text'   },
          { label: 'Arrivée',              key: 'arrivee',               type: 'text'   },
          { label: 'BL / OT',              key: 'bl_ot',                 type: 'text'   },
          { label: 'BC',                   key: 'bc',                    type: 'text'   },
          { label: 'Délai Paiement (J)',   key: 'delai_paiement',        type: 'number' },
          { label: 'Date de Paiement',     key: 'date_paiement',         type: 'date'   },
          { label: 'Règlement Banque/Type',key: 'reglement_banque_type', type: 'text'   },
          { label: 'Règlement N°',         key: 'reglement_numero',      type: 'text'   },
          { label: 'Échéances',            key: 'echeances',             type: 'text'   },
          { label: 'Mode de Paiement',     key: 'mode_paiement',         type: 'text'   },
        ].map(({ label, key, type }) => (
          <div key={key}>
            <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">{label}</label>
            <input type={type} value={(factForm as any)[key] || ''}
              onChange={e => setFactForm((p: any) => ({ ...p, [key]: e.target.value }))}
              className="w-full mt-1 h-9 rounded-lg border-2 border-slate-200 px-3 text-sm focus:outline-none focus:border-blue-500" />
          </div>
        ))}

        {/* Montant HT */}
        <div>
          <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Montant HT (MAD)</label>
          <input type="number" value={factForm.montant_ht || ''}
            onChange={e => {
              const ht  = parseFloat(e.target.value) || 0;
              const tva = parseFloat(factForm.tva)    || 0;
              setFactForm((p: any) => ({
                ...p,
                montant_ht:  e.target.value,
                montant_ttc: tva > 0 ? String((ht * (1 + tva / 100)).toFixed(2)) : String((ht + parseFloat(p.tva_amount || '0')).toFixed(2)),
              }));
            }}
            className="w-full mt-1 h-9 rounded-lg border-2 border-slate-200 px-3 text-sm focus:outline-none focus:border-blue-500" />
        </div>

        {/* TVA rate % */}
        <div>
          <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">TVA (%)</label>
          <div className="flex gap-2 mt-1 items-center">
            {['0', '7', '10', '14', '20'].map(rate => (
              <button key={rate} type="button"
                onClick={() => {
                  const ht = parseFloat(factForm.montant_ht) || 0;
                  const r  = parseFloat(rate);
                  const tvaAmount = parseFloat((ht * r / 100).toFixed(2));
                  const ttc = parseFloat((ht + tvaAmount).toFixed(2));
                  setFactForm((p: any) => ({
                    ...p,
                    tva:         String(tvaAmount),
                    montant_ttc: String(ttc),
                    tva_rate:    rate,
                  }));
                }}
                className={`h-9 px-2 rounded-lg text-xs font-black border-2 transition-all cursor-pointer ${factForm.tva_rate === rate ? 'bg-blue-600 border-blue-600 text-white' : 'border-slate-200 text-slate-600 hover:border-blue-300'}`}>
                {rate}%
              </button>
            ))}
            <input
              type="number"
              placeholder="Autre %"
              value={['0','7','10','14','20'].includes(factForm.tva_rate) ? '' : factForm.tva_rate || ''}
              onChange={e => {
                const rate = e.target.value;
                const ht = parseFloat(factForm.montant_ht) || 0;
                const r  = parseFloat(rate) || 0;
                const tvaAmount = parseFloat((ht * r / 100).toFixed(2));
                const ttc = parseFloat((ht + tvaAmount).toFixed(2));
                setFactForm((p: any) => ({
                  ...p,
                  tva:         String(tvaAmount),
                  montant_ttc: String(ttc),
                  tva_rate:    rate,
                }));
              }}
              className="flex-1 h-9 rounded-lg border-2 border-slate-200 px-3 text-sm focus:outline-none focus:border-blue-500"
            />
          </div>
        </div>

        {/* TVA amount — read only, auto calculated */}
        <div>
          <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">TVA Montant (MAD) — auto</label>
          <input type="number" value={factForm.tva || ''}
            readOnly
            className="w-full mt-1 h-9 rounded-lg border-2 border-slate-100 bg-slate-50 px-3 text-sm text-slate-500 cursor-not-allowed" />
        </div>

        {/* Montant TTC — auto calculated */}
        <div>
          <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">
            Montant TTC (MAD) — auto
            <span className="ml-1 text-blue-500 normal-case font-medium">= HT + TVA</span>
          </label>
          <input type="number" value={factForm.montant_ttc || ''}
            readOnly
            className="w-full mt-1 h-9 rounded-lg border-2 border-blue-100 bg-blue-50 px-3 text-sm font-bold text-blue-700 cursor-not-allowed" />
        </div>

        {/* Statut */}
        <div>
          <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Statut</label>
          <select value={factForm.statut || 'impayé'}
            onChange={e => setFactForm((p: any) => ({ ...p, statut: e.target.value }))}
            className="w-full mt-1 h-9 rounded-lg border-2 border-slate-200 px-3 text-sm focus:outline-none focus:border-blue-500">
            <option value="impayé">Impayé</option>
            <option value="payé">Payé</option>
          </select>
        </div>

      </div>
      <div className="flex gap-3 pt-5">
        <button onClick={handleSaveFacturation}
          className="flex-1 h-10 bg-blue-600 hover:bg-blue-700 text-white text-sm font-bold rounded-lg cursor-pointer">
          {editingFact ? 'Enregistrer les modifications' : 'Ajouter la facture'}
        </button>
        <button onClick={() => { setShowFactForm(false); setEditingFact(null); }}
          className="flex-1 h-10 bg-slate-100 hover:bg-slate-200 text-slate-700 text-sm font-bold rounded-lg cursor-pointer">
          Annuler
        </button>
      </div>
    </motion.div>
  </div>
)}
</AnimatePresence>
{/* Invoice Preview Modal */}
<AnimatePresence>
  {showPreview && (
    <div className="fixed inset-0 bg-black/80 z-50 flex items-center justify-center p-4">
      <motion.div initial={{ scale: 0.95, opacity: 0 }} animate={{ scale: 1, opacity: 1 }}
        exit={{ scale: 0.95, opacity: 0 }}
        className="bg-white rounded-xl w-full max-w-2xl max-h-[90vh] overflow-y-auto shadow-2xl">
        <div className="flex items-center justify-between p-4 border-b border-slate-200">
          <h3 className="font-black text-slate-900 uppercase tracking-widest text-sm">Aperçu du Modèle</h3>
          <button onClick={() => setShowPreview(false)} className="p-1.5 rounded hover:bg-slate-100 text-slate-400"><X size={16} /></button>
        </div>
        <div className="p-6">
          {/* Mini preview */}
          <div className="border-2 rounded-xl overflow-hidden" style={{ borderColor: invoiceSettings.primary_color }}>
            {/* Header preview */}
            <div className="p-4 flex justify-between items-start" style={{ backgroundColor: invoiceSettings.primary_color + '15' }}>
              <div>
                <p className="font-black text-sm" style={{ color: invoiceSettings.primary_color }}>
                  {invoiceSettings.company_name || activeCompany?.name || 'Nom de l\'entreprise'}
                </p>
                <p className="text-[10px] text-slate-500 mt-1">{invoiceSettings.address || 'Adresse...'}</p>
                <p className="text-[10px] text-slate-500">{invoiceSettings.phone} {invoiceSettings.email}</p>
                {invoiceSettings.ice && <p className="text-[10px] text-slate-500">ICE: {invoiceSettings.ice}</p>}
              </div>
              <div className="text-right">
                <p className="text-xl font-black" style={{ color: invoiceSettings.primary_color }}>
                  {invoiceSettings.invoice_title || 'FACTURE'}
                </p>
                <p className="text-[10px] text-slate-500">{new Date().toLocaleDateString('fr-MA')}</p>
              </div>
            </div>
            {/* Client box preview */}
            <div className="mx-4 my-3 p-3 bg-slate-50 rounded-lg" style={{ borderLeft: `4px solid ${invoiceSettings.accent_color}` }}>
              <p className="text-xs font-bold text-slate-700">NOM DU CLIENT</p>
              <p className="text-[10px] text-slate-500">X prestation(s) — Délai: {invoiceSettings.delai_paiement || 60}j</p>
            </div>
            {/* Table preview */}
            <div className="mx-4 mb-3">
              <div className="rounded-lg overflow-hidden border border-slate-200">
                <div className="grid text-[9px] font-black text-white uppercase px-2 py-1.5"
                  style={{ backgroundColor: invoiceSettings.primary_color,
                    gridTemplateColumns: `repeat(${4 + (invoiceSettings.show_bl_ot ? 1:0) + (invoiceSettings.show_bc ? 1:0)}, 1fr)` }}>
                  <span>Date</span><span>N° Fact.</span>
                  {invoiceSettings.show_bl_ot && <span>BL/OT</span>}
                  {invoiceSettings.show_bc    && <span>BC</span>}
                  <span>HT</span><span>TTC</span>
                </div>
                <div className="px-2 py-1.5 text-[9px] text-slate-400 bg-slate-50">
                  Lignes de facturation...
                </div>
              </div>
            </div>
            {/* Totals preview */}
            <div className="flex justify-end mx-4 mb-3">
              <div className="border-2 rounded-lg p-3 min-w-[160px]" style={{ borderColor: invoiceSettings.primary_color }}>
                <div className="flex justify-between text-[10px] text-slate-600 mb-1"><span>Total HT</span><span>X MAD</span></div>
                <div className="flex justify-between text-[10px] text-slate-600 mb-1"><span>TVA</span><span>X MAD</span></div>
                <div className="flex justify-between text-xs font-black" style={{ color: invoiceSettings.primary_color }}>
                  <span>Total TTC</span><span>X MAD</span>
                </div>
              </div>
            </div>
            {/* Footer preview */}
            {(invoiceSettings.rib || invoiceSettings.bank_name) && (
              <div className="mx-4 mb-3 p-2 bg-slate-50 rounded text-[9px] text-slate-500 border border-slate-200">
                {invoiceSettings.bank_name && <span>Banque: {invoiceSettings.bank_name} </span>}
                {invoiceSettings.rib && <span>RIB: {invoiceSettings.rib}</span>}
              </div>
            )}
            {invoiceSettings.footer_text && (
              <div className="mx-4 mb-3 text-[9px] text-slate-400 border-t border-slate-200 pt-2 text-center">
                {invoiceSettings.footer_text}
              </div>
            )}
            <div className="flex justify-end mx-4 mb-4">
              <div className="border border-slate-200 rounded p-2 text-center min-w-[120px]">
                <p className="text-[9px] text-slate-400 uppercase font-bold">{invoiceSettings.signature_label || 'Signature & Cachet'}</p>
                <div className="h-8" />
              </div>
            </div>
          </div>
          <p className="text-[10px] text-slate-400 text-center mt-3">Ceci est un aperçu — le PDF final contiendra vos vraies données de facturation.</p>
        </div>
      </motion.div>
    </div>
  )}
</AnimatePresence>
    </div>
  );
}