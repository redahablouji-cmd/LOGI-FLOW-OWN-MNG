import { useState, useEffect } from 'react';
import { useNavigate } from 'react-router-dom';
import { Loader2, LogOut, Users, ShoppingBag, Wrench, Menu, X, BadgeCheck, RefreshCw, Plus, Eye, Download, FileText, Pencil, Trash2, Truck, Upload, Receipt } from 'lucide-react';
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

type ManagerTab = 'staff' | 'purchases' | 'fleetfix' | 'suivi' | 'chauffeurs' | 'clients' | 'facturation';

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
  montant_ht: '', tva: '', montant_ttc: '',
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
      company_id: companyId || null, manager_id: managerId || null,
      date: suiviForm.date || null, matricule: suiviForm.matricule || null,
      type: suiviForm.type || null, facture: suiviForm.facture || null,
      bon_commande: suiviForm.bon_commande || null, ot_bl_bs_be: suiviForm.ot_bl_bs_be || null,
      client: suiviForm.client || null, depart: suiviForm.depart || null, arrivee: suiviForm.arrivee || null,
      manutention:    parseFloat(suiviForm.manutention    as string) || 0,
      immobilisation: parseFloat(suiviForm.immobilisation as string) || 0,
      prix_ht:        parseFloat(suiviForm.prix_ht        as string) || 0,
      prix_ttc:       parseFloat(suiviForm.prix_ttc       as string) || 0,
      cout_revient:   parseFloat(suiviForm.cout_revient   as string) || 0,
      benefice:       parseFloat(suiviForm.benefice       as string) || 0,
    };
    if (editingSuivi) {
      const { error } = await supabase.from('suivi_facturation').update(payload).eq('id', editingSuivi.id);
      if (!error) { toast.success("Modifié."); setEditingSuivi(null); fetchSuivi(); }
      else toast.error(`Erreur: ${error.message}`);
    } else {
      const { error } = await supabase.from('suivi_facturation').insert(payload);
      if (!error) { toast.success("Ajouté."); setShowSuiviForm(false); fetchSuivi(); }
      else toast.error(`Erreur: ${error.message}`);
    }
    setSuiviForm(emptySuivi);
  };

  const handleDeleteSuivi = async (id: string) => {
    if (!confirm('Supprimer ?')) return;
    const { error } = await supabase.from('suivi_facturation').delete().eq('id', id);
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
    .select('*')
    .eq('company_id', companyId)
    .order('created_at', { ascending: false });
  setFacturationList(data || []);
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
    company_id:           companyId || null,
    manager_id:           managerId || null,
    prestation_id:        factForm.prestation_id || null,
    date:                 factForm.date           || null,
    numero_facture:       factForm.numero_facture || null,
    client:               factForm.client         || null,
    depart:               factForm.depart         || null,
    arrivee:              factForm.arrivee        || null,
    montant_ht:           parseFloat(factForm.montant_ht)  || 0,
    tva:                  parseFloat(factForm.tva)         || 0,
    montant_ttc:          parseFloat(factForm.montant_ttc) || 0,
    bl_ot:                factForm.bl_ot                   || null,
    bc:                   factForm.bc                      || null,
    delai_paiement:       parseInt(factForm.delai_paiement) || 60,
    date_paiement:        factForm.date_paiement            || null,
    reglement_banque_type: factForm.reglement_banque_type  || null,
    reglement_numero:     factForm.reglement_numero        || null,
    echeances:            factForm.echeances               || null,
    mode_paiement:        factForm.mode_paiement           || null,
    statut:               factForm.statut || 'impayé',
    ecart_delai: calcEcartDelai(factForm.date, factForm.date_paiement, parseInt(factForm.delai_paiement) || 60),
  };

  if (editingFact) {
    const { error } = await supabase.from('suivi_facturation').update(payload).eq('id', editingFact.id);
    if (!error) { toast.success("Facture modifiée."); setEditingFact(null); fetchFacturation(); }
    else toast.error(`Erreur: ${error.message}`);
  } else {
    const { error } = await supabase.from('suivi_facturation').insert(payload);
    if (!error) { toast.success("Facture ajoutée."); setShowFactForm(false); fetchFacturation(); }
    else toast.error(`Erreur: ${error.message}`);
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
const handleGenerateInvoicePDF = () => {
  const selected = facturationList.filter(f => selectedFacts.includes(f.id));
  if (selected.length === 0) { toast.error("Sélectionnez au moins une facture."); return; }

  const ROWS_PER_PAGE = 15;
  const pages = [];
  for (let i = 0; i < selected.length; i += ROWS_PER_PAGE) {
    pages.push(selected.slice(i, i + ROWS_PER_PAGE));
  }

  const totalHT  = selected.reduce((s, f) => s + (f.montant_ht  || 0), 0);
  const totalTVA = selected.reduce((s, f) => s + (f.tva         || 0), 0);
  const totalTTC = selected.reduce((s, f) => s + (f.montant_ttc || 0), 0);
  const clientName = selected[0]?.client || '';

  const html = `
<!DOCTYPE html>
<html lang="fr">
<head>
  <meta charset="UTF-8"/>
  <style>
    * { margin: 0; padding: 0; box-sizing: border-box; }
    body { font-family: Arial, sans-serif; font-size: 11px; color: #1e293b; }
    .page { width: 210mm; min-height: 297mm; padding: 15mm; page-break-after: always; }
    .page:last-child { page-break-after: auto; }
    .header { display: flex; justify-content: space-between; align-items: flex-start; margin-bottom: 20px; border-bottom: 2px solid #1e40af; padding-bottom: 12px; }
    .logo { font-size: 22px; font-weight: 900; color: #1e40af; letter-spacing: -1px; }
    .logo span { color: #f59e0b; }
    .invoice-title { font-size: 18px; font-weight: 700; color: #1e293b; text-align: right; }
    .invoice-meta { text-align: right; color: #64748b; margin-top: 4px; font-size: 10px; }
    .client-box { background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 6px; padding: 10px 14px; margin-bottom: 16px; }
    .client-box strong { font-size: 13px; color: #1e293b; }
    table { width: 100%; border-collapse: collapse; margin-bottom: 16px; }
    th { background: #1e40af; color: white; padding: 7px 8px; text-align: left; font-size: 9px; text-transform: uppercase; letter-spacing: 0.5px; }
    td { padding: 6px 8px; border-bottom: 1px solid #f1f5f9; font-size: 10px; }
    tr:nth-child(even) td { background: #f8fafc; }
    .totals { display: flex; justify-content: flex-end; margin-top: 8px; }
    .totals-box { border: 2px solid #1e40af; border-radius: 6px; padding: 10px 16px; min-width: 220px; }
    .totals-row { display: flex; justify-content: space-between; padding: 3px 0; font-size: 11px; }
    .totals-row.total { font-weight: 700; font-size: 13px; color: #1e40af; border-top: 1px solid #e2e8f0; margin-top: 4px; padding-top: 6px; }
    .footer { position: fixed; bottom: 10mm; left: 15mm; right: 15mm; text-align: center; color: #94a3b8; font-size: 9px; border-top: 1px solid #e2e8f0; padding-top: 6px; }
    .page-num { position: fixed; bottom: 10mm; right: 15mm; font-size: 9px; color: #94a3b8; }
    .badge { display: inline-block; padding: 2px 8px; border-radius: 4px; font-size: 9px; font-weight: 700; text-transform: uppercase; }
    .badge-green { background: #dcfce7; color: #166534; }
    .badge-red { background: #fee2e2; color: #991b1b; }
    @media print { body { print-color-adjust: exact; -webkit-print-color-adjust: exact; } }
  </style>
</head>
<body>
${pages.map((pageRows, pageIdx) => `
  <div class="page">
    <div class="header">
      <div>
        <div class="logo">LOGI<span>-FLOW</span></div>
        <div style="color:#64748b;font-size:10px;margin-top:4px;">${activeCompany?.name || ''}</div>
      </div>
      <div>
        <div class="invoice-title">FACTURE</div>
        <div class="invoice-meta">Page ${pageIdx + 1} / ${pages.length}</div>
        <div class="invoice-meta">Généré le: ${new Date().toLocaleDateString('fr-MA')}</div>
      </div>
    </div>

    <div class="client-box">
      <strong>${clientName}</strong><br/>
      <span style="color:#64748b;font-size:10px;">
        ${selected.length} ligne(s) — Délai paiement: ${selected[0]?.delai_paiement || 60} jours
      </span>
    </div>

    <table>
      <thead>
        <tr>
          <th>Date</th>
          <th>N° Fact.</th>
          <th>BL/OT</th>
          <th>BC</th>
          <th>Départ</th>
          <th>Arrivée</th>
          <th>HT</th>
          <th>TVA</th>
          <th>TTC</th>
          <th>Statut</th>
        </tr>
      </thead>
      <tbody>
        ${pageRows.map(f => `
          <tr>
            <td>${f.date || '—'}</td>
            <td><strong>${f.numero_facture || '—'}</strong></td>
            <td>${f.bl_ot || '—'}</td>
            <td>${f.bc   || '—'}</td>
            <td>${f.depart  || '—'}</td>
            <td>${f.arrivee || '—'}</td>
            <td style="text-align:right">${Number(f.montant_ht).toLocaleString('fr-MA')} MAD</td>
            <td style="text-align:right">${Number(f.tva).toLocaleString('fr-MA')} MAD</td>
            <td style="text-align:right;font-weight:700">${Number(f.montant_ttc).toLocaleString('fr-MA')} MAD</td>
            <td><span class="badge ${f.statut === 'payé' ? 'badge-green' : 'badge-red'}">${f.statut || 'impayé'}</span></td>
          </tr>
        `).join('')}
      </tbody>
    </table>

    ${pageIdx === pages.length - 1 ? `
    <div class="totals">
      <div class="totals-box">
        <div class="totals-row"><span>Total HT</span><span>${totalHT.toLocaleString('fr-MA')} MAD</span></div>
        <div class="totals-row"><span>Total TVA</span><span>${totalTVA.toLocaleString('fr-MA')} MAD</span></div>
        <div class="totals-row total"><span>Total TTC</span><span>${totalTTC.toLocaleString('fr-MA')} MAD</span></div>
      </div>
    </div>
    ` : ''}
  </div>
`).join('')}
</body>
</html>`;

  const win = window.open('', '_blank');
  if (win) {
    win.document.write(html);
    win.document.close();
    win.focus();
    setTimeout(() => win.print(), 500);
  }
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
    if (activeTab === 'facturation' && companyId) { fetchFacturation(); fetchSuivi(); fetchClients(); }
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
                    <span className={`px-2 py-0.5 rounded text-[9px] font-bold uppercase ${f.statut === 'payé' ? 'bg-emerald-50 text-emerald-700' : 'bg-rose-50 text-rose-700'}`}>
                      {f.statut || 'impayé'}
                    </span>
                  </td>
                  <td className="px-3 py-3 text-xs text-slate-600">{f.mode_paiement || '—'}</td>
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
                  <td className="px-3 py-3">
                    <div className="flex items-center gap-1">
                      <button onClick={() => {
                        setEditingFact(f);
                        setFactForm({
                          date: f.date, numero_facture: f.numero_facture, client: f.client,
                          depart: f.depart, arrivee: f.arrivee,
                          montant_ht: String(f.montant_ht), tva: String(f.tva),
                          montant_ttc: String(f.montant_ttc), bl_ot: f.bl_ot, bc: f.bc,
                          delai_paiement: String(f.delai_paiement), date_paiement: f.date_paiement || '',
                          reglement_banque_type: f.reglement_banque_type || '',
                          reglement_numero: f.reglement_numero || '',
                          echeances: f.echeances || '', mode_paiement: f.mode_paiement || '',
                          statut: f.statut || 'impayé',
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
          {[
            { label: 'Date',                  key: 'date',                  type: 'date'   },
            { label: 'N° Facture',            key: 'numero_facture',        type: 'text'   },
            { label: 'Client',                key: 'client',                type: 'text'   },
            { label: 'Départ',                key: 'depart',                type: 'text'   },
            { label: 'Arrivée',               key: 'arrivee',               type: 'text'   },
            { label: 'Montant HT (MAD)',       key: 'montant_ht',            type: 'number' },
            { label: 'TVA (MAD)',              key: 'tva',                   type: 'number' },
            { label: 'Montant TTC (MAD)',      key: 'montant_ttc',           type: 'number' },
            { label: 'BL / OT',               key: 'bl_ot',                 type: 'text'   },
            { label: 'BC',                    key: 'bc',                    type: 'text'   },
            { label: 'Délai Paiement (J)',     key: 'delai_paiement',        type: 'number' },
            { label: 'Date de Paiement',       key: 'date_paiement',         type: 'date'   },
            { label: 'Règlement Banque/Type',  key: 'reglement_banque_type', type: 'text'   },
            { label: 'Règlement N°',           key: 'reglement_numero',      type: 'text'   },
            { label: 'Échéances',              key: 'echeances',             type: 'text'   },
            { label: 'Mode de Paiement',       key: 'mode_paiement',         type: 'text'   },
          ].map(({ label, key, type }) => (
            <div key={key}>
              <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">{label}</label>
              <input type={type} value={(factForm as any)[key] || ''}
                onChange={e => setFactForm((p: any) => ({ ...p, [key]: e.target.value }))}
                className="w-full mt-1 h-9 rounded-lg border-2 border-slate-200 px-3 text-sm focus:outline-none focus:border-blue-500" />
            </div>
          ))}
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
    </div>
  );
}