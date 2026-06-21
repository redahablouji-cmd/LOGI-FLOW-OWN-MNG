import { useState, useEffect } from 'react';
import { useNavigate } from 'react-router-dom';
import { Loader2, LogOut, Users, ShoppingBag, Wrench, Menu, X, BadgeCheck, RefreshCw, Plus, Eye, Download, FileText, Pencil, Trash2, Truck, Upload, Receipt, Settings, TrendingUp, FolderOpen } from 'lucide-react';
import { useAuth } from '../hooks/useAuth';
import { supabase } from '../lib/supabase';
import { Company } from '../lib/auth';
import CreateStaffForm from '../components/manager/CreateStaffForm';
import CoutRevientTab from '../components/manager/CoutRevientTab';
import InvoiceEngine, { generateInvoicePDF, numberToWords } from '../components/manager/InvoiceEngine';
import TruckDocumentsTab from '../components/manager/TruckDocumentsTab';
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

type ManagerTab = 'staff' | 'purchases' | 'fleetfix' | 'suivi' | 'chauffeurs' | 'cout_revient' | 'clients' | 'fournisseurs' | 'truck_docs' | 'facturation' | 'settings';

interface Purchase {
  id: string; category: string; fournisseur: string; numero_facture: string;
  date_achat: string; montant_ht: number; tva_rate: number; tva_amount: number;
  montant_ttc: number; banque: string; mode_paiement: string; notes: string; created_at: string;
  designation?: string; if_number?: string; ice_number?: string;
  affectation_immatriculation?: string; numero_ref?: string;
  echeance?: string; ecart_delai_paiement?: number;
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
  observation: '',
  designation: '',
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

  // Fournisseurs
  const [fournisseursList,      setFournisseursList]      = useState<any[]>([]);
  const [loadingFournisseurs,   setLoadingFournisseurs]   = useState(false);
  const [editingFournisseur,    setEditingFournisseur]    = useState<any | null>(null);
  const [fournisseurEditForm,   setFournisseurEditForm]   = useState<any>({});
  const [uploadingFournisseurs, setUploadingFournisseurs] = useState(false);

  // Purchases
  const [purchases,        setPurchases]        = useState<Purchase[]>([]);
  const [loadingPurchases, setLoadingPurchases] = useState(false);
  const [editingPurchase,  setEditingPurchase]  = useState<Purchase | null>(null);
  const [uploadingPurchaseXLS, setUploadingPurchaseXLS] = useState(false);
  const [purchaseFilter, setPurchaseFilter] = useState({
    fournisseur: '', dateFrom: '', dateTo: '', category: '', matricule: '', banque: '', mode: '',
  });

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
  const [clientSearch,    setClientSearch]    = useState('');
  const [matriculeSearch, setMatriculeSearch] = useState('');
  const [showClientDrop,  setShowClientDrop]  = useState(false);
  const [showMatDrop,     setShowMatDrop]     = useState(false);
// Suivi Facturation
const [facturationList,    setFacturationList]    = useState<any[]>([]);
const [loadingFacturation, setLoadingFacturation] = useState(false);
const [showFactForm,       setShowFactForm]       = useState(false);
const [editingFact,        setEditingFact]        = useState<any | null>(null);
const [factForm,           setFactForm]           = useState<any>({});
const [selectedFacts, setSelectedFacts] = useState<string[]>([]);
const [allTemplates, setAllTemplates] = useState<any[]>([]);
const [selectedTemplateId, setSelectedTemplateId] = useState('');
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
  const [selectedPrestations, setSelectedPrestations] = useState<string[]>([]);
  const [prestationFilter, setPrestationFilter] = useState({ client: '', dateFrom: '', dateTo: '', matricule: '' });
  const [wizardOpen, setWizardOpen] = useState(false);
  const [wizardPrestations, setWizardPrestations] = useState<any[]>([]);
  const [wizardForms, setWizardForms] = useState<any[]>([]);
  const [wizardStep, setWizardStep] = useState(0);
  const [wizardGroupId, setWizardGroupId] = useState('');
  const [savingWizard, setSavingWizard] = useState(false);
  const [expandedGroups, setExpandedGroups] = useState<string[]>([]);
  const [showAvoirForm, setShowAvoirForm] = useState(false);
  const [avoirForm, setAvoirForm] = useState({ date: new Date().toISOString().split('T')[0], numero_facture: '', client: '', depart: '', arrivee: '', montant_ht: '', tva: '', montant_ttc: '', observation: '' });
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
  const handlePurchaseXLSUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setUploadingPurchaseXLS(true);
    try {
      const buffer = await file.arrayBuffer();
      const wb = XLSX.read(buffer, { type: 'array' });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const rawRows: any[] = XLSX.utils.sheet_to_json(ws, { header: 1 });
      const dataRows = rawRows.slice(1).filter((r: any[]) => r.length > 0 && (r[0] || r[1] || r[2]));

      const records = dataRows.map((r: any[]) => ({
        date_achat:                  r[0] ? String(r[0]) : null,
        numero_facture:              String(r[1] || '').trim() || null,
        fournisseur:                 String(r[2] || '').trim() || null,
        category:                    String(r[3] || '').trim() || null,
        designation:                 String(r[4] || '').trim() || null,
        montant_ht:                  parseFloat(r[5]) || 0,
        tva_rate:                    parseFloat(r[6]) || 0,
        tva_amount:                  parseFloat(r[7]) || 0,
        montant_ttc:                 parseFloat(r[8]) || 0,
        ecart_delai_paiement:        parseInt(r[9]) || null,
        if_number:                   String(r[10] || '').trim() || null,
        ice_number:                  String(r[11] || '').trim() || null,
        affectation_immatriculation: String(r[12] || '').trim() || null,
        banque:                      String(r[13] || '').trim() || null,
        numero_ref:                  String(r[14] || '').trim() || null,
        echeance:                    r[15] ? String(r[15]) : null,
        mode_paiement:               String(r[16] || '').trim() || null,
      }));

      if (records.length === 0) { toast.error("Aucune donnée trouvée."); return; }
      const { error } = await supabase.from('purchases').insert(records);
      if (!error) { toast.success(`${records.length} achats importés.`); fetchPurchases(); }
      else toast.error(`Erreur: ${error.message}`);
    } catch (err: any) {
      toast.error(`Erreur: ${err.message}`);
    } finally {
      setUploadingPurchaseXLS(false);
      e.target.value = '';
    }
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
        client:         payload.client,
        depart:         payload.depart,
        arrivee:        payload.arrivee,
        montant_ht:     payload.prix_ht,
        montant_ttc:    payload.prix_ttc,
        bl_ot:          payload.ot_bl_bs_be,
        bc:             payload.bon_commande,
        numero_facture: payload.facture,
        date:           payload.date,
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
      consommation:        parseFloat(r[4]) || 0,
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
    consommation:        parseFloat(driverEditForm.consommation) || 0,
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
      company_id:     companyId,
      nom:            String(r[0] || '').trim(),
      adresse:        String(r[1] || '').trim(),
      ice:            String(r[2] || '').trim(),
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
// ── Fournisseurs ──────────────────────────────────────────────────────
const fetchFournisseurs = async () => {
  if (!companyId) return;
  setLoadingFournisseurs(true);
  const { data } = await supabase
    .from('fournisseurs')
    .select('*')
    .eq('company_id', companyId)
    .order('created_at', { ascending: false });
  setFournisseursList(data || []);
  setLoadingFournisseurs(false);
};

const handleFournisseursXLSUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
  const file = e.target.files?.[0];
  if (!file || !companyId) return;
  setUploadingFournisseurs(true);
  try {
    const buffer = await file.arrayBuffer();
    const wb = XLSX.read(buffer, { type: 'array' });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const rows: any[] = XLSX.utils.sheet_to_json(ws, { header: 1 });
    const dataRows = rows.slice(1).filter((r: any[]) => r.length > 0 && r[0]);
    const records = dataRows.map((r: any[]) => ({
      company_id:     companyId,
      nom:            String(r[0] || '').trim(),
      categorie:      String(r[1] || '').trim(),
      taux_tva:       parseFloat(r[2]) || 0.20,
      if_number:      String(r[3] || '').trim(),
      ice:            String(r[4] || '').trim(),
      adresse:        String(r[5] || '').trim(),
      telephone:      String(r[6] || '').trim(),
      banque:         String(r[7] || '').trim(),
      delai_paiement: parseInt(r[8]) || 60,
    }));
    if (records.length === 0) { toast.error("Aucune donnée trouvée."); return; }
    await supabase.from('fournisseurs').delete().eq('company_id', companyId);
    const { error } = await supabase.from('fournisseurs').insert(records);
    if (!error) { toast.success(`${records.length} fournisseurs importés.`); fetchFournisseurs(); }
    else toast.error(`Erreur: ${error.message}`);
  } catch (err: any) {
    toast.error(`Erreur: ${err.message}`);
  } finally {
    setUploadingFournisseurs(false);
    e.target.value = '';
  }
};

const handleDeleteFournisseur = async (id: string) => {
  if (!confirm('Supprimer ce fournisseur ?')) return;
  const { error } = await supabase.from('fournisseurs').delete().eq('id', id);
  if (!error) { setFournisseursList(prev => prev.filter(f => f.id !== id)); toast.success("Supprimé."); }
  else toast.error(`Erreur: ${error.message}`);
};

const handleEditFournisseurSave = async () => {
  const { error } = await supabase.from('fournisseurs').update({
    nom:            fournisseurEditForm.nom,
    categorie:      fournisseurEditForm.categorie,
    taux_tva:       parseFloat(fournisseurEditForm.taux_tva) || 0.20,
    if_number:      fournisseurEditForm.if_number,
    ice:            fournisseurEditForm.ice,
    adresse:        fournisseurEditForm.adresse,
    telephone:      fournisseurEditForm.telephone,
    banque:         fournisseurEditForm.banque,
    delai_paiement: parseInt(fournisseurEditForm.delai_paiement) || 60,
  }).eq('id', editingFournisseur.id);
  if (!error) { toast.success("Fournisseur modifié."); setEditingFournisseur(null); fetchFournisseurs(); }
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
  const ht     = parseFloat(prestation.prix_ht)  || 0;
  const ttc    = parseFloat(prestation.prix_ttc) || 0;
  const tvaAmt = parseFloat((ttc - ht).toFixed(2));
  const rate   = ht > 0 ? String(Math.round(tvaAmt / ht * 100)) : '';

  setFactForm({
    ...emptyFactForm,
    // From prestation directly
    date:                  prestation.date          || new Date().toISOString().split('T')[0],
    numero_facture:        prestation.facture        || '',
    client:                prestation.client         || '',
    depart:                prestation.depart         || '',
    arrivee:               prestation.arrivee        || '',
    montant_ht:            String(ht),
    tva:                   String(tvaAmt),
    montant_ttc:           String(ttc),
    tva_rate:              rate,
    bl_ot:                 prestation.ot_bl_bs_be   || '',
    bc:                    prestation.bon_commande   || '',
    delai_paiement:        String(client?.delai_paiement || prestation.delai_paiement || 60),
    // User fills these manually
    date_paiement:         '',
    reglement_banque_type: '',
    reglement_numero:      '',
    echeances:             '',
    mode_paiement:         '',
    statut:                'impayé',
    prestation_id:         prestation.id,
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
    is_avoir:              (factForm as any).is_avoir       || false,
    observation:           (factForm as any).observation     || null,
    designation:           (factForm as any).designation     || null,
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
          facture:      payload.numero_facture,
          date:         payload.date,
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
const handleStartWizard = () => {
    const selected = suiviList.filter((p: any) => selectedPrestations.includes(p.id));
    if (selected.length < 2) return;
    const groupId = crypto.randomUUID();
    const forms = selected.map((prestation: any) => {
      const client = clientsList.find((c: any) => c.nom === prestation.client);
      const ht = parseFloat(prestation.prix_ht) || 0;
      const ttc = parseFloat(prestation.prix_ttc) || 0;
      const tvaAmt = parseFloat((ttc - ht).toFixed(2));
      const rate = ht > 0 ? String(Math.round(tvaAmt / ht * 100)) : '';
      return {
        date: prestation.date || new Date().toISOString().split('T')[0],
        numero_facture: prestation.facture || '',
        client: prestation.client || '',
        depart: prestation.depart || '',
        arrivee: prestation.arrivee || '',
        montant_ht: String(ht),
        tva: String(tvaAmt),
        montant_ttc: String(ttc),
        tva_rate: rate,
        bl_ot: prestation.ot_bl_bs_be || '',
        bc: prestation.bon_commande || '',
        delai_paiement: String(client?.delai_paiement || 60),
        date_paiement: '',
        reglement_banque_type: '',
        reglement_numero: '',
        echeances: '',
        mode_paiement: '',
        statut: 'impayé',
        prestation_id: prestation.id,
      };
    });
    setWizardGroupId(groupId);
    setWizardPrestations(selected);
    setWizardForms(forms);
    setWizardStep(0);
    setPrestationPickerOpen(false);
    setSelectedPrestations([]);
    setWizardOpen(true);
  };

  const handleSaveWizard = async () => {
    setSavingWizard(true);
    try {
      for (const form of wizardForms) {
        const payload = {
          company_id: companyId || null,
          manager_id: managerId || null,
          prestation_id: form.prestation_id || null,
          invoice_group_id: wizardGroupId,
          date: form.date || null,
          numero_facture: form.numero_facture || null,
          client: form.client || null,
          depart: form.depart || null,
          arrivee: form.arrivee || null,
          montant_ht: parseFloat(form.montant_ht) || 0,
          tva: parseFloat(form.tva) || 0,
          montant_ttc: parseFloat(form.montant_ttc) || 0,
          bl_ot: form.bl_ot || null,
          bc: form.bc || null,
          delai_paiement: parseInt(form.delai_paiement) || 60,
          date_paiement: form.date_paiement || null,
          reglement_banque_type: form.reglement_banque_type || null,
          reglement_numero: form.reglement_numero || null,
          echeances: form.echeances || null,
          mode_paiement: form.mode_paiement || null,
          statut: form.statut || 'impayé',
          observation: form.observation || null,
          ecart_delai: calcEcartDelai(form.date, form.date_paiement, parseInt(form.delai_paiement) || 60),
        };
        await supabase.from('suivi_facturation').insert(payload);
      }
      toast.success(`${wizardForms.length} factures groupées créées !`);
      setWizardOpen(false);
      setWizardForms([]);
      setWizardPrestations([]);
      setWizardStep(0);
      fetchFacturation();
    } catch (err: any) {
      toast.error(`Erreur: ${err.message}`);
    } finally {
      setSavingWizard(false);
    }
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
  const { data } = await supabase.from('invoice_templates')
        .select('*').eq('company_id', companyId).eq('is_default', true).maybeSingle();
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
    .from('invoice_templates')
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

const handleGenerateInvoicePDF = () => {
    if (selectedFacts.length === 0) return;
    const selected = facturationList.filter((f: any) => selectedFacts.includes(f.id));
    if (selected.length === 0) return;

    // Get the selected template (or fall back to default invoiceSettings)
    const tmpl = allTemplates.find((t: any) => t.id === selectedTemplateId);
    const s = invoiceSettings;
    const ROWS_PER_PAGE_FIRST = s.rows_per_page || 18;
    const ROWS_PER_PAGE = ROWS_PER_PAGE_FIRST + 6;

    const totalHT  = selected.reduce((sum: number, f: any) => sum + (parseFloat(f.montant_ht) || 0), 0);
    const totalTVA = selected.reduce((sum: number, f: any) => sum + (parseFloat(f.tva) || 0), 0);
    const totalTTC = selected.reduce((sum: number, f: any) => sum + (parseFloat(f.montant_ttc) || 0), 0);
    const fmt = (n: number) => n.toLocaleString('fr-MA', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    // Detect if this is a facture d'avoir
    const isAvoir = selected.some((f: any) => f.is_avoir);

    // Auto-fill client from clients table
    const clientName = selected[0]?.client || '';
    const clientData = clientsList.find((c: any) => c.nom === clientName);
    const clientAddress = clientData?.adresse || '';
    const clientICE = clientData?.ice || '';

    // ─── Detect column structure from template ───
    // Parse column_mapping if available, otherwise auto-detect from template HTML
    let columns: { header: string; field: string; align: string; format?: string }[] = [];

    if (tmpl?.column_mapping && Array.isArray(tmpl.column_mapping) && tmpl.column_mapping.length > 0) {
      columns = tmpl.column_mapping;
    } else {
      // Default: detect from template name or use 9-column fallback
      const templateHtml = tmpl?.original_html || tmpl?.invoice_template_html || s.original_html || s.invoice_template_html || '';
      // Count <th> tags in the data table to determine column count
      const thMatch = templateHtml.match(/<th[^>]*>[\s\S]*?<\/th>/gi);
      const thCount = thMatch ? thMatch.length : 0;

      if (thCount >= 16) {
        // 10-column template (6 meta + 10 data)
        columns = [
          { header: 'Date', field: 'date', align: 'center' },
          { header: 'Lieu de déchargement', field: 'arrivee', align: 'left' },
          { header: 'Client', field: 'client', align: 'center' },
          { header: 'N°BL', field: 'bl_ot', align: 'center' },
          { header: 'N°EXP', field: 'numero_facture', align: 'center' },
          { header: 'Type', field: 'type', align: 'center' },
          { header: 'Quantité', field: 'quantity', align: 'center' },
          { header: 'Prix unitaire HT', field: 'montant_ht', align: 'right', format: 'number' },
          { header: 'TVA %', field: 'tva_rate', align: 'center' },
          { header: 'Montant HT', field: 'montant_ht_total', align: 'right', format: 'number' },
        ];
      } else if (thCount >= 15) {
        // 9-column template (6 meta + 9 data)
        columns = [
          { header: 'Date/Poste', field: 'date', align: 'center' },
          { header: 'Désignation', field: 'designation', align: 'left' },
          { header: 'Type', field: 'type', align: 'center' },
          { header: 'BL', field: 'bl_ot', align: 'center' },
          { header: 'BC', field: 'bc', align: 'center' },
          { header: 'Quantité', field: 'quantity', align: 'center' },
          { header: 'Prix unitaire HT', field: 'montant_ht', align: 'right', format: 'number' },
          { header: 'TVA %', field: 'tva_rate', align: 'center' },
          { header: 'Montant HT', field: 'montant_ht_total', align: 'right', format: 'number' },
        ];
      } else if (thCount >= 12) {
        // 6-column template (6 meta + 6 data)
        columns = [
          { header: 'Date/Poste', field: 'date', align: 'center' },
          { header: 'Désignation', field: 'designation', align: 'left' },
          { header: 'Quantité', field: 'quantity', align: 'center' },
          { header: 'Prix unitaire HT', field: 'montant_ht', align: 'right', format: 'number' },
          { header: 'TVA %', field: 'tva_rate', align: 'center' },
          { header: 'Montant HT', field: 'montant_ht_total', align: 'right', format: 'number' },
        ];
      } else {
        // Fallback
        columns = [
          { header: 'Date', field: 'date', align: 'center' },
          { header: 'Désignation', field: 'designation', align: 'left' },
          { header: 'Quantité', field: 'quantity', align: 'center' },
          { header: 'Prix unitaire HT', field: 'montant_ht', align: 'right', format: 'number' },
          { header: 'TVA %', field: 'tva_rate', align: 'center' },
          { header: 'Montant HT', field: 'montant_ht_total', align: 'right', format: 'number' },
        ];
      }
    }

    // ─── Dynamic row builder ───
    const getFieldValue = (f: any, field: string): string => {
      const ht = parseFloat(f.montant_ht) || 0;
      const tva = parseFloat(f.tva) || 0;
      const tvaRate = ht !== 0 ? (Math.abs(tva / ht) * 100).toFixed(0) + '%' : '0%';
      switch (field) {
        case 'date': return f.date || '';
        case 'designation': return isAvoir && f.observation ? f.observation : `${f.depart || ''} → ${f.arrivee || ''}`;
        case 'type': return f.type || '';
        case 'bl_ot': return f.bl_ot || f.ot_bl_bs_be || '';
        case 'bc': return f.bc || '';
        case 'quantity': return '1';
        case 'montant_ht': return fmt(ht);
        case 'tva_rate': return tvaRate;
        case 'montant_ht_total': return fmt(ht);
        case 'tva_amount': return fmt(tva);
        case 'montant_ttc': return fmt(parseFloat(f.montant_ttc) || 0);
        case 'numero_facture': return f.numero_facture || '';
        case 'client': return f.client || '';
        case 'depart': return f.depart || '';
        case 'arrivee': return f.arrivee || '';
        default: return f[field] || '';
      }
    };

    const buildRow = (f: any, i: number) => {
      const bg = i % 2 === 1 ? '#F2F2F2' : '#fff';
     return `<tr>${columns.map(col => {
        const val = col.format === 'number' ? getFieldValue(f, col.field) : getFieldValue(f, col.field);
        return `<td style="padding:7px 8px;border:1px solid #ddd;text-align:${col.align};font-size:11px;background:${bg}">${val}</td>`;
      }).join('')}</tr>`;
    };

    // ─── Dynamic thead ───
    const colW = Math.floor(100 / columns.length);
    const theadHTML = `<thead><tr>${columns.map(col =>
      `<th style="background:#1F3864;color:#fff;font-size:11px;font-weight:700;text-align:center;padding:8px 6px;text-transform:uppercase;border:1px solid #1F3864;width:${colW}%">${col.header}</th>`
    ).join('')}</tr></thead>`;

    // ─── Paginate ───
    const pages: { rows: any[]; isFirst: boolean; isLast: boolean; num: number }[] = [];
    const remaining = [...selected];
    let pageNum = 0;
    while (remaining.length > 0) {
      pageNum++;
      const isFirst = pageNum === 1;
      const capacity = isFirst ? ROWS_PER_PAGE_FIRST : ROWS_PER_PAGE;
      const pageRows = remaining.splice(0, capacity);
      pages.push({ rows: pageRows, isFirst, isLast: remaining.length === 0, num: pageNum });
    }
    const totalPages = pages.length;

    const clientHTML = `<div style="display:flex;justify-content:flex-end;margin-bottom:20px">
      <div style="border:1px solid #bbb;border-left:4px solid #D4A017;padding:12px 18px;min-width:320px;background:#FAFAFA;border-radius:0 4px 4px 0">
        <div style="font-size:14px;font-weight:700;color:#1e293b">${clientName}</div>
        ${clientAddress ? `<div style="font-size:11px;color:#555;margin-top:4px">${clientAddress}</div>` : ''}
        ${clientICE ? `<div style="font-size:11px;color:#555;margin-top:3px">ICE: ${clientICE}</div>` : ''}
      </div>
    </div>`;

    const metaHTML = `<table style="width:100%;border-collapse:collapse;margin-bottom:14px">
      <tr>
        <th style="background:#1F3864;color:#fff;font-size:11px;font-weight:700;text-align:center;padding:8px 8px;text-transform:uppercase;border:1px solid #1F3864">${isAvoir ? 'N° Avoir' : 'N° Facture'}</th>
        <th style="background:#1F3864;color:#fff;font-size:11px;font-weight:700;text-align:center;padding:8px 8px;text-transform:uppercase;border:1px solid #1F3864">N°OT / N°BL</th>
        <th style="background:#1F3864;color:#fff;font-size:11px;font-weight:700;text-align:center;padding:8px 8px;text-transform:uppercase;border:1px solid #1F3864">Date</th>
        <th style="background:#1F3864;color:#fff;font-size:11px;font-weight:700;text-align:center;padding:8px 8px;text-transform:uppercase;border:1px solid #1F3864">N°Commande</th>
        <th style="background:#1F3864;color:#fff;font-size:11px;font-weight:700;text-align:center;padding:8px 8px;text-transform:uppercase;border:1px solid #1F3864">Échéance</th>
        <th style="background:#1F3864;color:#fff;font-size:11px;font-weight:700;text-align:center;padding:8px 8px;text-transform:uppercase;border:1px solid #1F3864">Réf. Client</th>
      </tr>
      <tr>
        <td style="text-align:center;padding:8px 8px;font-size:12px;border:1px solid #ddd;background:#fff">${selected[0]?.numero_facture || ''}</td>
        <td style="text-align:center;padding:8px 8px;font-size:12px;border:1px solid #ddd;background:#fff">${selected[0]?.bl_ot || selected[0]?.ot_bl_bs_be || ''}</td>
        <td style="text-align:center;padding:8px 8px;font-size:12px;border:1px solid #ddd;background:#fff">${selected[0]?.date || ''}</td>
        <td style="text-align:center;padding:8px 8px;font-size:12px;border:1px solid #ddd;background:#fff">${selected[0]?.bc || ''}</td>
        <td style="text-align:center;padding:8px 8px;font-size:12px;border:1px solid #ddd;background:#fff">${selected[0]?.delai_paiement || 60} jours</td>
        <td style="text-align:center;padding:8px 8px;font-size:12px;border:1px solid #ddd;background:#fff">${clientName}</td>
      </tr>
    </table>`;

    const totalsHTML = `<div style="display:flex;justify-content:flex-end;margin-top:0">
      <div style="width:300px;border:1px solid #ddd">
        <div style="display:flex;justify-content:space-between;padding:6px 12px;font-size:12px;font-weight:700;border-bottom:1px solid #eee"><span>Sous-total HT</span><span>${fmt(totalHT)} MAD</span></div>
        <div style="display:flex;justify-content:space-between;padding:6px 12px;font-size:12px;font-weight:700;border-bottom:1px solid #eee"><span>TVA</span><span>${fmt(totalTVA)} MAD</span></div>
        <div style="display:flex;justify-content:space-between;padding:6px 12px;font-size:12px;font-weight:700;border-bottom:1px solid #eee"><span>Remise</span><span>0,00 MAD</span></div>
        <div style="display:flex;justify-content:space-between;padding:8px 12px;font-size:14px;font-weight:900;background:#1F3864;color:#fff"><span>TOTAL TTC</span><span>${fmt(totalTTC)} MAD</span></div>
      </div>
    </div>
    <div style="margin-top:16px;font-size:11px;color:#7F7F7F;line-height:1.5">
      <strong style="color:#333">Arrêtée la présente ${isAvoir ? "facture d'avoir" : 'facture'} à la somme de :</strong>
      ${numberToWords(Math.abs(totalTTC))}
    </div>`;

    const pagesHTML = pages.map(p => {
      // Observation row (first row, first page only, spans all columns)
      const obsText = p.isFirst && !isAvoir && selected[0]?.observation ? selected[0].observation : '';
      const obsRow = obsText
        ? `<tr><td colspan="${columns.length}" style="padding:8px 10px;border:1px solid #ddd;font-size:11px;font-weight:700;font-style:italic;background:#fff;text-align:center">${obsText}</td></tr>`
        : '';
      const rowsHtml = obsRow + p.rows.map((f: any, i: number) => buildRow(f, i + (obsText ? 1 : 0))).join('');

      // Fill remaining space with empty rows
      const maxRows = p.isFirst ? ROWS_PER_PAGE_FIRST : ROWS_PER_PAGE;
      const emptyCount = Math.max(0, (p.isLast ? maxRows - 4 : maxRows) - p.rows.length);
      const emptyCols = columns.length;
      const emptyRowsHtml = Array.from({ length: emptyCount }, (_, i) => {
        const idx = p.rows.length + i;
        const bg = idx % 2 === 1 ? '#F2F2F2' : '#fff';
        return `<tr>${Array.from({ length: emptyCols }, () =>
          `<td style="padding:7px 8px;border:1px solid #ddd;font-size:11px;background:${bg}">&nbsp;</td>`
        ).join('')}</tr>`;
      }).join('');

      return `<div style="width:210mm;min-height:297mm;padding:${p.isFirst ? '35mm' : '15mm'} 12mm 15mm 12mm;position:relative;font-family:Arial,sans-serif;font-size:10px;color:#000;page-break-after:${p.isLast ? 'auto' : 'always'}">
        <div style="position:absolute;top:8mm;right:12mm;font-size:8px;color:#999">Page ${p.num} / ${totalPages}</div>
        ${isAvoir && p.isFirst ? '<div style="font-size:16px;font-weight:900;color:#B91C1C;text-align:center;margin-bottom:10px;text-decoration:underline">AVOIR</div>' : ''}
        ${p.isFirst ? clientHTML : ''}
        ${p.isFirst && !isAvoir ? metaHTML : ''}
        <table style="width:100%;border-collapse:collapse">${theadHTML}<tbody>${rowsHtml}${emptyRowsHtml}</tbody></table>
        ${p.isLast ? totalsHTML : ''}
      </div>`;
    }).join('');

    const html = `<!DOCTYPE html><html lang="fr"><head><meta charset="UTF-8"/>
      <style>*{margin:0;padding:0;box-sizing:border-box}@page{margin:0;size:A4}@media print{body{print-color-adjust:exact;-webkit-print-color-adjust:exact}html,body{margin:0;padding:0}}</style>
      </head><body>${pagesHTML}</body></html>`;

    const win = window.open('', '_blank');
    if (win) { win.document.write(html); win.document.close(); win.focus(); setTimeout(() => win.print(), 600); }
  };
  const handleGenerateRelance = () => {
    const selected = facturationList.filter((f: any) => selectedFacts.includes(f.id));
    if (selected.length === 0) return;

    const impaye = selected.filter((f: any) => f.statut !== 'payé');
    if (impaye.length === 0) {
      toast.error("Aucune facture impayée dans la sélection.");
      return;
    }

    const clientName = impaye[0]?.client || '';
    const clientData = clientsList.find((c: any) => c.nom === clientName);
    const clientAddress = clientData?.adresse || '';
    const clientICE = clientData?.ice || '';

    const totalImpaye = impaye.reduce((sum: number, f: any) => sum + (parseFloat(f.montant_ttc) || 0), 0);
    const fmt = (n: number) => n.toLocaleString('fr-MA', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    const today = new Date().toLocaleDateString('fr-MA', { day: 'numeric', month: 'long', year: 'numeric' });

    // Calculate how many rows fit on one page
    // Fixed space: top margin(35mm) + client(25mm) + date(10mm) + title(12mm) + intro(35mm) + thead(10mm) + total row(10mm) + closing(35mm) + bottom(15mm) = ~187mm
    // Available for rows on A4 (297mm): 297 - 187 = ~110mm → ~15 rows at 7mm each
    const ROWS_PER_PAGE = 15;

    // Paginate if needed
    const pages: { rows: any[]; isFirst: boolean; isLast: boolean; num: number }[] = [];
    const remaining = [...impaye];
    let pageNum = 0;
    const firstPageRows = ROWS_PER_PAGE;
    const nextPageRows = ROWS_PER_PAGE + 12; // no intro/client on subsequent pages

    while (remaining.length > 0) {
      pageNum++;
      const isFirst = pageNum === 1;
      const capacity = isFirst ? firstPageRows : nextPageRows;
      const pageRows = remaining.splice(0, capacity);
      pages.push({ rows: pageRows, isFirst, isLast: remaining.length === 0, num: pageNum });
    }
    const totalPages = pages.length;

    const buildRow = (f: any, i: number) => {
      const bg = i % 2 === 1 ? '#F2F2F2' : '#fff';
      return `<tr>
        <td style="padding:6px 10px;border:1px solid #ddd;text-align:center;font-size:11px;background:${bg}">${f.numero_facture || '—'}</td>
        <td style="padding:6px 10px;border:1px solid #ddd;text-align:center;font-size:11px;background:${bg}">${f.date || '—'}</td>
        <td style="padding:6px 10px;border:1px solid #ddd;text-align:right;font-size:11px;font-weight:700;background:${bg}">${fmt(parseFloat(f.montant_ttc) || 0)}</td>
      </tr>`;
    };

    const theadHTML = `<thead><tr>
      <th style="background:#1F3864;color:#fff;font-size:11px;font-weight:700;text-align:center;padding:8px 10px;border:1px solid #1F3864;width:35%">N° Facture</th>
      <th style="background:#1F3864;color:#fff;font-size:11px;font-weight:700;text-align:center;padding:8px 10px;border:1px solid #1F3864;width:30%">Date Facture</th>
      <th style="background:#1F3864;color:#fff;font-size:11px;font-weight:700;text-align:center;padding:8px 10px;border:1px solid #1F3864;width:35%">Montant TTC</th>
    </tr></thead>`;

    const totalRow = `<tr>
      <td colspan="2" style="padding:8px 10px;border:1px solid #1F3864;font-size:12px;font-weight:900;text-align:right;background:#1F3864;color:#fff">TOTAL IMPAYÉ</td>
      <td style="padding:8px 10px;border:1px solid #1F3864;font-size:13px;font-weight:900;text-align:right;background:#1F3864;color:#fff">${fmt(totalImpaye)} DHS</td>
    </tr>`;

    const pagesHTML = pages.map(p => {
      const rowsHtml = p.rows.map((f: any, i: number) => buildRow(f, i)).join('');

      // Fill empty rows to use remaining space
      const maxOnPage = p.isFirst ? firstPageRows : nextPageRows;
      const emptyCount = Math.max(0, maxOnPage - p.rows.length);
      const emptyRowsHtml = Array.from({ length: emptyCount }, (_, i) => {
        const idx = p.rows.length + i;
        const bg = idx % 2 === 1 ? '#F2F2F2' : '#fff';
        return `<tr>
          <td style="padding:6px 10px;border:1px solid #ddd;font-size:11px;background:${bg}">&nbsp;</td>
          <td style="padding:6px 10px;border:1px solid #ddd;font-size:11px;background:${bg}">&nbsp;</td>
          <td style="padding:6px 10px;border:1px solid #ddd;font-size:11px;background:${bg}">&nbsp;</td>
        </tr>`;
      }).join('');

      return `<div style="width:210mm;min-height:297mm;padding:${p.isFirst ? '35mm' : '15mm'} 15mm 15mm 15mm;position:relative;font-family:Arial,sans-serif;page-break-after:${p.isLast ? 'auto' : 'always'}">
        <div style="position:absolute;top:8mm;right:15mm;font-size:8px;color:#999">Page ${p.num} / ${totalPages}</div>

        ${p.isFirst ? `
          <!-- Client box -->
          <div style="display:flex;justify-content:flex-end;margin-bottom:16px">
            <div style="border:2px solid #1F3864;padding:12px 18px;min-width:300px;border-radius:4px">
              <div style="font-size:14px;font-weight:900;color:#1F3864">${clientName}</div>
              ${clientAddress ? `<div style="font-size:11px;color:#555;margin-top:4px">${clientAddress}</div>` : ''}
              ${clientICE ? `<div style="font-size:11px;color:#555;margin-top:3px">ICE: ${clientICE}</div>` : ''}
            </div>
          </div>
          <div style="text-align:right;font-size:12px;color:#333;margin-bottom:20px">Casablanca le ${today}</div>
          <div style="font-size:18px;font-weight:900;color:#1F3864;text-decoration:underline;margin-bottom:16px">Relance</div>
          <div style="font-size:12px;line-height:1.8;color:#333;margin-bottom:14px">
            <p>Madame, Monsieur,</p>
            <br/>
            <p>Sauf erreur ou omission de notre part, nous constatons que votre compte client présente à ce jour un solde débiteur de <strong style="color:#1F3864;font-size:13px">${fmt(totalImpaye)} DHS</strong>.</p>
            <br/>
            <p>Ce montant correspond à nos factures suivantes restées impayées :</p>
          </div>
        ` : ''}

        <!-- Table -->
        <table style="width:100%;border-collapse:collapse">
          ${theadHTML}
          <tbody>
            ${rowsHtml}
            ${emptyRowsHtml}
            ${p.isLast ? totalRow : ''}
          </tbody>
        </table>

        ${p.isLast ? `
          <div style="font-size:12px;line-height:1.8;color:#333;margin-top:14px">
            <p>L'échéance étant dépassée, nous vous demandons de bien vouloir régulariser cette situation dans les meilleurs délais.</p>
            <br/>
            <p>Vous remerciant par avance, nous vous prions d'agréer, Monsieur, l'expression de nos salutations distinguées.</p>
          </div>
        ` : ''}
      </div>`;
    }).join('');

    const html = `<!DOCTYPE html><html lang="fr"><head><meta charset="UTF-8"/>
      <style>*{margin:0;padding:0;box-sizing:border-box}@page{margin:0;size:A4}@media print{body{print-color-adjust:exact;-webkit-print-color-adjust:exact}}</style>
      </head><body>${pagesHTML}</body></html>`;

    const win = window.open('', '_blank');
    if (win) { win.document.write(html); win.document.close(); win.focus(); setTimeout(() => win.print(), 600); }
  };
  useEffect(() => {
    if (!loading) { if (!user) navigate('/login'); else fetchCompany(); }
  }, [user, loading]);

  useEffect(() => {
    if (activeTab === 'purchases') fetchPurchases();
    if (activeTab === 'fleetfix' && companyId) fetchMechanics();
    if (activeTab === 'suivi') { fetchSuivi(); fetchClients(); if (companyId) fetchFleetDrivers(); }
    if (activeTab === 'chauffeurs' && companyId) fetchFleetDrivers();
    if (activeTab === 'clients' && companyId) fetchClients();
    if (activeTab === 'fournisseurs' && companyId) fetchFournisseurs();
    if (activeTab === 'purchases') { fetchPurchases(); if (companyId) fetchFournisseurs(); }
    if (activeTab === 'facturation' && companyId) {
      fetchFacturation(); fetchSuivi(); fetchClients(); fetchInvoiceSettings();
      supabase.from('invoice_templates').select('*')
        .eq('company_id', companyId).order('created_at').then(({ data }) => {
          setAllTemplates(data || []);
          const def = (data || []).find((t: any) => t.is_default);
          if (def && !selectedTemplateId) setSelectedTemplateId(def.id);
        });
    }
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
  { id: 'cout_revient',label: 'Coût de Revient',      icon: TrendingUp },
  { id: 'clients',     label: 'Clients',              icon: Users },
  { id: 'fournisseurs',label: 'Fournisseurs',         icon: ShoppingBag },
  { id: 'truck_docs',  label: 'Documents Camions',    icon: FolderOpen },
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
  const filteredPurchases = purchases.filter((p: any) => {
    if (purchaseFilter.fournisseur && !p.fournisseur?.toLowerCase().includes(purchaseFilter.fournisseur.toLowerCase())) return false;
    if (purchaseFilter.category && p.category !== purchaseFilter.category) return false;
    if (purchaseFilter.dateFrom && (p.date_achat || '') < purchaseFilter.dateFrom) return false;
    if (purchaseFilter.dateTo && (p.date_achat || '') > purchaseFilter.dateTo) return false;
    if (purchaseFilter.matricule && !p.affectation_immatriculation?.toLowerCase().includes(purchaseFilter.matricule.toLowerCase())) return false;
    if (purchaseFilter.banque && p.banque !== purchaseFilter.banque) return false;
    if (purchaseFilter.mode && p.mode_paiement !== purchaseFilter.mode) return false;
    return true;
  });

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
                    <label className={`flex items-center gap-1.5 px-3 py-2 rounded-lg text-xs font-black uppercase tracking-wider cursor-pointer transition-all ${uploadingPurchaseXLS ? 'bg-slate-600 opacity-60' : 'bg-amber-600 hover:bg-amber-700'} text-white`}>
                      <Upload size={14} />
                      {uploadingPurchaseXLS ? 'Importation...' : 'Importer XLS'}
                      <input type="file" accept=".xlsx,.xls" onChange={handlePurchaseXLSUpload} className="hidden" disabled={uploadingPurchaseXLS} />
                    </label>
                    <button onClick={() => exportToXLS(filteredPurchases.map((p: any) => ({
                      'Date':                        p.date_achat                  || '',
                      'Catégorie':                   CATEGORY_LABELS[p.category]   || p.category,
                      'Fournisseur':                 p.fournisseur                 || '',
                      'N° Facture':                  p.numero_facture              || '',
                      'Désignation':                 p.designation                 || '',
                      'IF':                          p.if_number                   || '',
                      'ICE':                         p.ice_number                  || '',
                      'Affectation Immatriculation': p.affectation_immatriculation || '',
                      'Montant HT':                  p.montant_ht,
                      'TVA':                         p.tva_amount,
                      'Montant TTC':                 p.montant_ttc,
                      'Banque':                      p.banque                      || '',
                      'N° Réf':                      p.numero_ref                  || '',
                      'Échéance':                    p.echeance                    || '',
                      'Mode Paiement':               p.mode_paiement               || '',
                      'Écart Délai (J)':             p.ecart_delai_paiement ?? '',
                      'Notes':                       p.notes                       || '',
                    })), 'achats_historique')}
                      className="bg-emerald-600 hover:bg-emerald-700 text-white px-3 py-2 rounded-lg text-xs font-black uppercase tracking-wider flex items-center gap-1.5 cursor-pointer">
                      <Download size={14} /> Export XLS
                    </button>
                  </div>
                </div>
              </div>
              {/* Filters */}
              <div className="mb-4 bg-white rounded-xl border border-slate-200 p-4 flex flex-wrap gap-3 items-end">
                <div>
                  <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Fournisseur</label>
                  <input type="text" placeholder="Rechercher..."
                    value={purchaseFilter.fournisseur}
                    onChange={e => setPurchaseFilter(p => ({ ...p, fournisseur: e.target.value }))}
                    className="block mt-1 h-8 rounded-lg border-2 border-slate-200 px-3 text-xs focus:outline-none focus:border-blue-500 w-44" />
                </div>
                <div>
                  <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Date de</label>
                  <input type="date" value={purchaseFilter.dateFrom}
                    onChange={e => setPurchaseFilter(p => ({ ...p, dateFrom: e.target.value }))}
                    className="block mt-1 h-8 rounded-lg border-2 border-slate-200 px-3 text-xs focus:outline-none focus:border-blue-500" />
                </div>
                <div>
                  <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Date à</label>
                  <input type="date" value={purchaseFilter.dateTo}
                    onChange={e => setPurchaseFilter(p => ({ ...p, dateTo: e.target.value }))}
                    className="block mt-1 h-8 rounded-lg border-2 border-slate-200 px-3 text-xs focus:outline-none focus:border-blue-500" />
                </div>
                <div>
                  <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Catégorie</label>
                  <select value={purchaseFilter.category}
                    onChange={e => setPurchaseFilter(p => ({ ...p, category: e.target.value }))}
                    className="block mt-1 h-8 rounded-lg border-2 border-slate-200 px-3 text-xs focus:outline-none focus:border-blue-500">
                    <option value="">Toutes</option>
                    {Object.entries(CATEGORY_LABELS).map(([k, v]) => (
                      <option key={k} value={k}>{v}</option>
                    ))}
                  </select>
                </div>
                <div>
                  <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Matricule</label>
                  <input type="text" placeholder="Rechercher..."
                    value={purchaseFilter.matricule}
                    onChange={e => setPurchaseFilter(p => ({ ...p, matricule: e.target.value }))}
                    className="block mt-1 h-8 rounded-lg border-2 border-slate-200 px-3 text-xs focus:outline-none focus:border-blue-500 w-36" />
                </div>
                <div>
                  <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Banque</label>
                  <select value={purchaseFilter.banque}
                    onChange={e => setPurchaseFilter(p => ({ ...p, banque: e.target.value }))}
                    className="block mt-1 h-8 rounded-lg border-2 border-slate-200 px-3 text-xs focus:outline-none focus:border-blue-500">
                    <option value="">Toutes</option>
                    <option value="AWB">AWB</option>
                    <option value="BMCE">BMCE</option>
                    <option value="BMCI">BMCI</option>
                    <option value="CIH">CIH</option>
                    <option value="SG">Société Générale</option>
                    <option value="Espèces">Espèces</option>
                  </select>
                </div>
                <div>
                <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Statut</label>
                <select value={factFilter?.statut || ''}
                  onChange={e => setFactFilter((p: any) => ({ ...p, statut: e.target.value }))}
                  className="block mt-1 h-8 rounded-lg border-2 border-slate-200 px-3 text-xs focus:outline-none focus:border-blue-500">
                  <option value="">Tous</option>
                  <option value="impayé">Impayé</option>
                  <option value="payé">Payé</option>
                  <option value="avoir">Avoir</option>
                </select>
              </div>
                <button onClick={() => setPurchaseFilter({ fournisseur: '', dateFrom: '', dateTo: '', category: '', matricule: '', banque: '', mode: '' })}
                  className="h-8 px-3 bg-slate-100 hover:bg-slate-200 text-slate-600 text-xs font-bold rounded-lg cursor-pointer">
                  Réinitialiser
                </button>
                <span className="text-[10px] text-slate-400 font-bold ml-auto">
                  {filteredPurchases.length} / {purchases.length} résultat(s)
                </span>
              </div>
              <div className="mb-4 p-3 bg-blue-50 border border-blue-100 rounded-xl text-xs text-blue-700 font-medium">
                📋 Format import XLS — colonnes dans l'ordre :
                <span className="font-black ml-1">Dates | Factures | Fournisseurs | Catégorie | Désignations | Montant HT | Taux | TVA | Montant TTC | Écart Délai | IF | ICE | Affectation Immat. | Banque | N° Réf | Échéance | Mode</span>
              </div>
              {loadingPurchases ? (
                <div className="flex items-center justify-center py-20"><Loader2 className="w-6 h-6 text-blue-600 animate-spin" /></div>
              ) : (
                <div className="bg-white rounded-xl border border-slate-200 overflow-hidden">
                  <div className="overflow-x-auto">
                    <table className="w-full text-left min-w-[1800px]">
                      <thead className="bg-slate-50 border-b border-slate-200">
                        <tr>{['Date','Catégorie','Fournisseur','N° Facture','Désignation','IF','ICE','Affectation','HT','TVA','TTC','Banque','N° Réf','Échéance','Mode','Écart Délai','Actions'].map(h => (
                          <th key={h} className="px-3 py-3 text-[9px] font-black text-slate-400 uppercase tracking-widest whitespace-nowrap">{h}</th>
                        ))}</tr>
                      </thead>
                      <tbody className="divide-y divide-slate-100">
                        {filteredPurchases.length === 0 ? (
                          <tr><td colSpan={17} className="px-4 py-10 text-center text-sm text-slate-400">{purchases.length === 0 ? 'Aucun achat enregistré.' : 'Aucun résultat pour ces filtres.'}</td></tr>
                        ) : filteredPurchases.map((p: any) => (
                          <tr key={p.id} className="hover:bg-slate-50 transition-colors">
                            <td className="px-3 py-3 text-xs text-slate-700 whitespace-nowrap">{p.date_achat || '—'}</td>
                            <td className="px-3 py-3"><span className="text-[10px] font-black bg-slate-100 text-slate-600 px-2 py-0.5 rounded uppercase whitespace-nowrap">{CATEGORY_LABELS[p.category] || p.category}</span></td>
                            <td className="px-3 py-3 text-xs font-semibold text-slate-700">{p.fournisseur || '—'}</td>
                            <td className="px-3 py-3 font-mono text-xs text-blue-600">{p.numero_facture || '—'}</td>
                            <td className="px-3 py-3 text-xs text-slate-600 max-w-[180px] truncate" title={p.designation || ''}>{p.designation || '—'}</td>
                            <td className="px-3 py-3 font-mono text-xs text-slate-600">{p.if_number || '—'}</td>
                            <td className="px-3 py-3 font-mono text-xs text-slate-600">{p.ice_number || '—'}</td>
                            <td className="px-3 py-3 font-mono text-xs font-bold text-cyan-700">{p.affectation_immatriculation || '—'}</td>
                            <td className="px-3 py-3 font-mono text-xs text-slate-700">{p.montant_ht?.toLocaleString('fr-MA', { minimumFractionDigits: 2 })}</td>
                            <td className="px-3 py-3 font-mono text-xs text-amber-700">{p.tva_amount?.toLocaleString('fr-MA', { minimumFractionDigits: 2 })}</td>
                            <td className="px-3 py-3 font-mono text-xs font-bold text-slate-900">{p.montant_ttc?.toLocaleString('fr-MA', { minimumFractionDigits: 2 })}</td>
                            <td className="px-3 py-3 text-xs text-slate-600">{p.banque || '—'}</td>
                            <td className="px-3 py-3 font-mono text-xs text-slate-700">{p.numero_ref || '—'}</td>
                            <td className="px-3 py-3 text-xs text-slate-600 whitespace-nowrap">{p.echeance || '—'}</td>
                            <td className="px-3 py-3 text-xs text-slate-600">{p.mode_paiement || '—'}</td>
                            <td className="px-3 py-3">
                              {p.ecart_delai_paiement !== null && p.ecart_delai_paiement !== undefined ? (
                                <span className={`px-2 py-0.5 rounded text-[9px] font-bold font-mono ${p.ecart_delai_paiement < 0 ? 'bg-rose-50 text-rose-700' : p.ecart_delai_paiement <= 15 ? 'bg-amber-50 text-amber-700' : 'bg-emerald-50 text-emerald-700'}`}>
                                  {p.ecart_delai_paiement > 0 ? `+${p.ecart_delai_paiement}` : p.ecart_delai_paiement}j
                                </span>
                              ) : '—'}
                            </td>
                            <td className="px-3 py-3">
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
            'Type': d.type_vehicule, 'Consommation': d.consommation, 'CIN': d.cin, 'IMM CNSS': d.imm_cnss,
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
      <span className="font-black ml-1">Code | Nom/Prénom | Immatriculation | Type | Consommation | CIN | IMM CNSS | Fonction | Date Naissance | Situation Familiale | Nb Déduction | Date Embauche | Adresse</span>
    </div>

    {loadingDrivers ? (
      <div className="flex items-center justify-center py-20"><Loader2 className="w-6 h-6 text-blue-600 animate-spin" /></div>
    ) : (
      <div className="bg-white rounded-xl border border-slate-200 overflow-hidden">
        <div className="overflow-x-auto">
          <table className="w-full text-left min-w-[1100px]">
            <thead className="bg-slate-50 border-b border-slate-200">
              <tr>
                {['Code','Nom / Prénom','Immat.','Type','Consommation','CIN','IMM CNSS','Fonction','Naissance','Situation','Déductions','Embauche','Adresse','Actions'].map(h => (
                  <th key={h} className="px-3 py-3 text-[9px] font-black text-slate-400 uppercase tracking-widest whitespace-nowrap">{h}</th>
                ))}
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100">
              {fleetDrivers.length === 0 ? (
                <tr>
                  <td colSpan={14} className="px-4 py-12 text-center">
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
                  <td className="px-3 py-3 font-mono text-xs text-slate-700 font-semibold">{d.consommation ? `${d.consommation} L/100km` : '—'}</td>
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
{activeTab === 'cout_revient' && (
  <CoutRevientTab companyId={companyId} />
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
{activeTab === 'fournisseurs' && (
  <div>
    <div className="mb-6 bg-slate-900 text-white rounded-xl p-6 border border-slate-800">
      <div className="flex items-center justify-between flex-wrap gap-3">
        <div>
          <span className="inline-flex items-center gap-1.5 px-2 py-0.5 rounded text-[10px] font-extrabold uppercase tracking-widest bg-orange-500 text-slate-950 mb-2">
            <ShoppingBag className="w-3.5 h-3.5" /> Base Fournisseurs
          </span>
          <h1 className="text-2xl font-extrabold tracking-tight">Gestion des Fournisseurs</h1>
          <p className="text-sm text-slate-400 mt-1">{fournisseursList.length} fournisseurs enregistrés</p>
        </div>
        <div className="flex gap-2 flex-wrap items-center">
          <button onClick={fetchFournisseurs}
            className="bg-white/10 hover:bg-white/15 px-3 py-2 rounded-lg text-xs font-bold uppercase tracking-wider flex items-center gap-1.5 cursor-pointer">
            <RefreshCw className="w-3.5 h-3.5" /> Actualiser
          </button>
          <button onClick={() => exportToXLS(fournisseursList.map(f => ({
            'Fournisseur': f.nom, 'Catégorie': f.categorie, 'Taux TVA': f.taux_tva,
            'IF': f.if_number, 'ICE': f.ice, 'Adresse': f.adresse,
            'Téléphone': f.telephone, 'Banque': f.banque, 'Délai Paiement': f.delai_paiement,
          })), 'fournisseurs_export')}
            className="bg-emerald-600 hover:bg-emerald-700 text-white px-3 py-2 rounded-lg text-xs font-black uppercase tracking-wider flex items-center gap-1.5 cursor-pointer">
            <Download size={14} /> Export XLS
          </button>
          <label className={`flex items-center gap-1.5 px-3 py-2 rounded-lg text-xs font-black uppercase tracking-wider cursor-pointer transition-all ${uploadingFournisseurs ? 'bg-slate-600 opacity-60' : 'bg-blue-600 hover:bg-blue-700'} text-white`}>
            <Upload size={14} />
            {uploadingFournisseurs ? 'Importation...' : 'Importer XLS'}
            <input type="file" accept=".xlsx,.xls" onChange={handleFournisseursXLSUpload} className="hidden" disabled={uploadingFournisseurs} />
          </label>
        </div>
      </div>
    </div>
    <div className="mb-4 p-4 bg-blue-50 border border-blue-100 rounded-xl text-xs text-blue-700 font-medium">
      📋 Format attendu : <span className="font-black">Fournisseur | Catégorie | Taux TVA | IF | ICE | Adresse | Téléphone | Banque | Délai Paiement (J)</span>
    </div>
    {loadingFournisseurs ? (
      <div className="flex items-center justify-center py-20"><Loader2 className="w-6 h-6 text-blue-600 animate-spin" /></div>
    ) : (
      <div className="bg-white rounded-xl border border-slate-200 overflow-hidden">
        <div className="overflow-x-auto">
          <table className="w-full text-left">
            <thead className="bg-slate-50 border-b border-slate-200">
              <tr>
                {['Fournisseur','Catégorie','Taux TVA','IF','ICE','Adresse','Téléphone','Banque','Délai (J)','Actions'].map(h => (
                  <th key={h} className="px-4 py-3 text-[10px] font-black text-slate-400 uppercase tracking-widest">{h}</th>
                ))}
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100">
              {fournisseursList.length === 0 ? (
                <tr>
                  <td colSpan={10} className="px-4 py-12 text-center">
                    <div className="flex flex-col items-center gap-3">
                      <ShoppingBag size={32} className="text-slate-300" />
                      <p className="text-sm text-slate-400 font-medium">Aucun fournisseur.</p>
                      <p className="text-xs text-slate-400">Cliquez sur "Importer XLS" pour charger votre base fournisseurs.</p>
                    </div>
                  </td>
                </tr>
              ) : fournisseursList.map(f => (
                <tr key={f.id} className="hover:bg-slate-50 transition-colors">
                  <td className="px-4 py-3 text-xs font-semibold text-slate-800">{f.nom}</td>
                  <td className="px-4 py-3"><span className="text-[9px] font-black bg-slate-100 text-slate-600 px-1.5 py-0.5 rounded uppercase">{f.categorie || '—'}</span></td>
                  <td className="px-4 py-3 font-mono text-xs font-bold text-amber-700">{f.taux_tva ? `${(f.taux_tva * 100).toFixed(0)}%` : '20%'}</td>
                  <td className="px-4 py-3 font-mono text-xs text-slate-600">{f.if_number || '—'}</td>
                  <td className="px-4 py-3 font-mono text-xs text-blue-600">{f.ice || '—'}</td>
                  <td className="px-4 py-3 text-xs text-slate-600 max-w-[250px] truncate">{f.adresse || '—'}</td>
                  <td className="px-4 py-3 text-xs text-slate-600">{f.telephone || '—'}</td>
                  <td className="px-4 py-3 text-xs text-slate-600">{f.banque || '—'}</td>
                  <td className="px-4 py-3 text-xs text-center font-bold text-slate-700">{f.delai_paiement} J</td>
                  <td className="px-4 py-3">
                    <div className="flex items-center gap-1">
                      <button onClick={() => { setEditingFournisseur(f); setFournisseurEditForm({ ...f }); }}
                        className="p-1.5 rounded hover:bg-blue-50 text-slate-400 hover:text-blue-600 transition-colors"><Pencil size={13} /></button>
                      <button onClick={() => handleDeleteFournisseur(f.id)}
                        className="p-1.5 rounded hover:bg-rose-50 text-slate-400 hover:text-rose-600 transition-colors"><Trash2 size={13} /></button>
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
{activeTab === 'truck_docs' && (
  <TruckDocumentsTab companyId={companyId} />
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
          <button onClick={() => {
            if (!filteredFacts.length) return;
            const groups: Record<string, any[]> = {};
            const ungrouped: any[] = [];
            filteredFacts.forEach((f: any) => {
              if (f.invoice_group_id) {
                if (!groups[f.invoice_group_id]) groups[f.invoice_group_id] = [];
                groups[f.invoice_group_id].push(f);
              } else {
                ungrouped.push(f);
              }
            });
            const rows: any[] = [];
            Object.entries(groups).forEach(([gid, items]) => {
              rows.push({
                'Date': items[0]?.date || '',
                'N° Facture': items[0]?.numero_facture || '',
                'Client': items[0]?.client || '',
                'Départ': items.map((i: any) => i.depart).filter(Boolean).join(' / '),
                'Arrivée': items.map((i: any) => i.arrivee).filter(Boolean).join(' / '),
                'Montant HT': items.reduce((s: number, i: any) => s + (parseFloat(i.montant_ht) || 0), 0),
                'TVA': items.reduce((s: number, i: any) => s + (parseFloat(i.tva) || 0), 0),
                'Montant TTC': items.reduce((s: number, i: any) => s + (parseFloat(i.montant_ttc) || 0), 0),
                'BL/OT': items.map((i: any) => i.bl_ot).filter(Boolean).join(' / '),
                'BC': items.map((i: any) => i.bc).filter(Boolean).join(' / '),
                'Délai (J)': items[0]?.delai_paiement || '',
                'Date Paiement': items[0]?.date_paiement || '',
                'Écart Délai (J)': items[0]?.ecart_delai ?? '',
                'Règlement Banque': items[0]?.reglement_banque_type || '',
                'Règlement N°': items[0]?.reglement_numero || '',
                'Échéances': items[0]?.echeances || '',
                'Mode Paiement': items[0]?.mode_paiement || '',
                'Statut': items.every((i: any) => i.statut === 'payé') ? 'payé' : 'impayé',
                'Observation': items[0]?.observation || '',
                'Type': 'Groupée (' + items.length + ')',
              });
            });
            ungrouped.forEach((f: any) => {
              rows.push({
                'Date': f.date || '',
                'N° Facture': f.numero_facture || '',
                'Client': f.client || '',
                'Départ': f.depart || '',
                'Arrivée': f.arrivee || '',
                'Montant HT': parseFloat(f.montant_ht) || 0,
                'TVA': parseFloat(f.tva) || 0,
                'Montant TTC': parseFloat(f.montant_ttc) || 0,
                'BL/OT': f.bl_ot || '',
                'BC': f.bc || '',
                'Délai (J)': f.delai_paiement || '',
                'Date Paiement': f.date_paiement || '',
                'Écart Délai (J)': f.ecart_delai ?? calcEcartDelai(f.date, f.date_paiement, f.delai_paiement),
                'Règlement Banque': f.reglement_banque_type || '',
                'Règlement N°': f.reglement_numero || '',
                'Échéances': f.echeances || '',
                'Mode Paiement': f.mode_paiement || '',
                'Statut': f.statut || '',
                'Observation': f.observation || '',
                'Type': 'Unique',
              });
            });
            if (rows.length === 0) return;
            exportToXLS(rows, 'suivi_facturation');
          }}
            className="bg-emerald-600 hover:bg-emerald-700 text-white px-3 py-2 rounded-lg text-xs font-black uppercase tracking-wider flex items-center gap-1.5 cursor-pointer">
            <Download size={14} /> Export XLS
          </button>
          {selectedFacts.length > 0 && (
                <div className="flex items-center gap-2">
                  <select value={selectedTemplateId}
                    onChange={e => setSelectedTemplateId(e.target.value)}
                    className="h-9 rounded-lg border-2 border-violet-300 bg-violet-50 px-3 text-xs font-bold text-violet-800 focus:outline-none focus:border-violet-500">
                    <option value="">— Modèle —</option>
                    {allTemplates.map((t: any) => (
                      <option key={t.id} value={t.id}>{t.template_name}{t.is_default ? ' ⭐' : ''}</option>
                    ))}
                  </select>
                  <button onClick={handleGenerateInvoicePDF}
                    className="bg-violet-600 hover:bg-violet-700 text-white px-3 py-2 rounded-lg text-xs font-black uppercase tracking-wider flex items-center gap-1.5 cursor-pointer">
                    <FileText size={14} /> Générer PDF ({selectedFacts.length})
                  </button>
                  <button onClick={handleGenerateRelance}
                    className="bg-rose-600 hover:bg-rose-700 text-white px-3 py-2 rounded-lg text-xs font-black uppercase tracking-wider flex items-center gap-1.5 cursor-pointer">
                    <FileText size={14} /> Relance Impayé
                  </button>
                </div>
              )}
          <label className={`flex items-center gap-1.5 px-3 py-2 rounded-lg text-xs font-black uppercase tracking-wider cursor-pointer transition-all ${uploadingFacts ? 'bg-slate-600 opacity-60' : 'bg-amber-600 hover:bg-amber-700'} text-white`}>
            <Upload size={14} />
            {uploadingFacts ? 'Importation...' : 'Importer XLS'}
            <input type="file" accept=".xlsx,.xls" onChange={handleFactXLSUpload} className="hidden" disabled={uploadingFacts} />
          </label>
          <button onClick={() => setShowAvoirForm(true)}
                  className="bg-rose-600 hover:bg-rose-700 text-white px-3 py-2 rounded-lg text-xs font-black uppercase tracking-wider flex items-center gap-1.5 cursor-pointer">
                  <FileText size={14} /> Facture d'Avoir
                </button>
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
           <option value="avoir">Avoir</option>
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
              {(() => {
                const groups: Record<string, any[]> = {};
                const ungrouped: any[] = [];
                filteredFacts.forEach((f: any) => {
                  if (f.invoice_group_id) {
                    if (!groups[f.invoice_group_id]) groups[f.invoice_group_id] = [];
                    groups[f.invoice_group_id].push(f);
                  } else {
                    ungrouped.push(f);
                  }
                });
                const allRows: { type: 'single' | 'group'; data: any; children?: any[] }[] = [];
                Object.entries(groups).forEach(([gid, items]) => {
                  allRows.push({ type: 'group', data: { id: gid, invoice_group_id: gid,
                    client: items[0]?.client, date: items[0]?.date, numero_facture: items[0]?.numero_facture,
                    montant_ht: items.reduce((s: number, i: any) => s + (parseFloat(i.montant_ht)||0), 0),
                    tva: items.reduce((s: number, i: any) => s + (parseFloat(i.tva)||0), 0),
                    montant_ttc: items.reduce((s: number, i: any) => s + (parseFloat(i.montant_ttc)||0), 0),
                    statut: items.every((i: any) => i.statut === 'payé') ? 'payé' : 'impayé',
                    count: items.length,
                  }, children: items });
                });
                ungrouped.forEach(f => allRows.push({ type: 'single', data: f }));

                if (allRows.length === 0) return (
                  <tr><td colSpan={16} className="px-4 py-10 text-center text-sm text-slate-400">Aucune facture. Cliquez sur "Nouvelle Facture".</td></tr>
                );

                return allRows.map((row) => {
                  if (row.type === 'group') {
                    const g = row.data;
                    const isExp = expandedGroups.includes(g.invoice_group_id);
                    const childIds = (row.children||[]).map((c:any) => c.id);
                    const allSel = childIds.every((id:string) => selectedFacts.includes(id));
                    return [
                      <tr key={`grp-${g.invoice_group_id}`}
                        onClick={() => setExpandedGroups(prev => prev.includes(g.invoice_group_id) ? prev.filter(x => x !== g.invoice_group_id) : [...prev, g.invoice_group_id])}
                        className="bg-blue-50/50 hover:bg-blue-50 cursor-pointer border-b-2 border-blue-200">
                        <td className="px-3 py-3"><input type="checkbox" checked={allSel} onClick={e => e.stopPropagation()}
                          onChange={e => setSelectedFacts(prev => e.target.checked ? [...new Set([...prev,...childIds])] : prev.filter(id => !childIds.includes(id)))} className="accent-blue-600" /></td>
                        <td className="px-3 py-3 text-xs text-slate-700">{g.date}</td>
                        <td className="px-3 py-3 font-mono text-xs text-blue-600">{g.numero_facture}</td>
                        <td className="px-3 py-3 text-xs font-semibold text-slate-700">{g.client}</td>
                        <td colSpan={2} className="px-3 py-3"><span className="text-[9px] font-black bg-blue-100 text-blue-700 px-2 py-0.5 rounded uppercase">{g.count} prestations groupées {isExp ? '▼' : '▶'}</span></td>
                        <td className="px-3 py-3 font-mono text-xs text-slate-700">{g.montant_ht.toLocaleString('fr-MA',{minimumFractionDigits:2})}</td>
                        <td className="px-3 py-3 font-mono text-xs text-amber-700">{g.tva.toLocaleString('fr-MA',{minimumFractionDigits:2})}</td>
                        <td className="px-3 py-3 font-mono text-xs font-bold text-slate-900">{g.montant_ttc.toLocaleString('fr-MA',{minimumFractionDigits:2})}</td>
                        <td colSpan={6} className="px-3 py-3"><span className={`text-[9px] font-black px-1.5 py-0.5 rounded ${g.statut==='payé'?'bg-emerald-50 text-emerald-700':'bg-rose-50 text-rose-700'}`}>{g.statut}</span></td>
                      </tr>,
                      ...(isExp ? (row.children||[]).map((f:any) => (
                        <tr key={f.id} className="bg-blue-50/20 hover:bg-blue-50/40 transition-colors">
                          <td className="px-3 py-2 pl-8"><input type="checkbox" checked={selectedFacts.includes(f.id)}
                            onChange={e => setSelectedFacts(prev => e.target.checked ? [...prev,f.id] : prev.filter(x=>x!==f.id))} className="accent-blue-600" /></td>
                          <td className="px-3 py-2 text-xs text-slate-500">{f.date||'—'}</td>
                          <td className="px-3 py-2 font-mono text-xs text-blue-400">{f.numero_facture||'—'}</td>
                          <td className="px-3 py-2 text-xs text-slate-500">{f.client||'—'}</td>
                          <td className="px-3 py-2 text-xs text-slate-500">{f.depart||'—'}</td>
                          <td className="px-3 py-2 text-xs text-slate-500">{f.arrivee||'—'}</td>
                          <td className="px-3 py-2 font-mono text-xs text-slate-500">{Number(f.montant_ht||0).toLocaleString('fr-MA',{minimumFractionDigits:2})}</td>
                          <td className="px-3 py-2 font-mono text-xs text-amber-500">{Number(f.tva||0).toLocaleString('fr-MA',{minimumFractionDigits:2})}</td>
                          <td className="px-3 py-2 font-mono text-xs text-slate-600">{Number(f.montant_ttc||0).toLocaleString('fr-MA',{minimumFractionDigits:2})}</td>
                          <td className="px-3 py-2 text-xs text-slate-500">{f.bl_ot||'—'}</td>
                          <td className="px-3 py-2 text-xs text-slate-500">{f.bc||'—'}</td>
                          <td className="px-3 py-2"><span className={`text-[9px] font-black px-1.5 py-0.5 rounded ${f.statut === 'payé' ? 'bg-emerald-50 text-emerald-700' : f.statut === 'avoir' || f.is_avoir ? 'bg-purple-50 text-purple-700' : 'bg-rose-50 text-rose-700'}`}>
                            {f.is_avoir ? 'AVOIR' : f.statut}
                          </span>
                          </td>
                          <td className="px-3 py-2">
                            <button onClick={() => { setEditingFact(f); setFactForm({...f, montant_ht: String(f.montant_ht), tva: String(f.tva), montant_ttc: String(f.montant_ttc), delai_paiement: String(f.delai_paiement || 60), is_avoir: f.is_avoir || false }); setShowFactForm(true); }}
                              className="text-[10px] font-bold text-blue-600 hover:text-blue-700 uppercase tracking-wider cursor-pointer">Modifier</button>
                          </td>
                        </tr>
                      )) : []),
                    ];
                  }
                  const f = row.data;
                  return (
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
                  <td className={`px-3 py-3 font-mono text-xs font-bold ${f.is_avoir ? 'text-rose-600' : 'text-slate-900'}`}>{Number(f.montant_ttc).toLocaleString('fr-MA')}</td>
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
                  );
                });
              })()}
            </tbody>
          </table>
        </div>
      </div>
    )}
  </div>
)}
{activeTab === 'settings' && (
          <InvoiceEngine companyId={companyId} logoPreviewUrl={logoPreviewUrl} />
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
        <button onClick={() => { setShowSuiviForm(false); setEditingSuivi(null); setSuiviForm(emptySuivi); setShowClientDrop(false); setShowMatDrop(false); }}
          className="p-1.5 rounded hover:bg-slate-100 text-slate-400"><X size={16} /></button>
      </div>

      <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">

        {/* CLIENT — searchable dropdown */}
        <div className="relative">
          <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Clients</label>
          <div className="relative mt-1">
            <input
              type="text"
              value={showClientDrop ? clientSearch : (suiviForm.client || '')}
              placeholder="Sélectionner ou rechercher..."
              onFocus={() => { setShowClientDrop(true); setClientSearch(suiviForm.client || ''); }}
              onChange={e => { setClientSearch(e.target.value); setShowClientDrop(true); }}
              className="w-full h-9 rounded-lg border-2 border-slate-200 px-3 text-sm focus:outline-none focus:border-blue-500 pr-8"
            />
            <span className="absolute right-2 top-1/2 -translate-y-1/2 text-slate-400 pointer-events-none text-xs">▾</span>
          </div>
          {showClientDrop && (
            <div className="absolute z-50 top-full left-0 right-0 bg-white border-2 border-blue-200 rounded-xl shadow-xl max-h-48 overflow-y-auto mt-1">
              {clientsList
                .filter(c => c.nom?.toLowerCase().includes(clientSearch.toLowerCase()))
                .length === 0 ? (
                <div className="px-3 py-3 text-xs text-slate-400 text-center">Aucun client trouvé</div>
              ) : clientsList
                .filter(c => c.nom?.toLowerCase().includes(clientSearch.toLowerCase()))
                .map(c => (
                  <button key={c.id} type="button"
                    onClick={() => {
                      setSuiviForm(p => ({ ...p, client: c.nom }));
                      setShowClientDrop(false);
                      setClientSearch('');
                    }}
                    className={`w-full text-left px-3 py-2 text-xs hover:bg-blue-50 transition-colors border-b border-slate-100 last:border-0 ${suiviForm.client === c.nom ? 'bg-blue-50 font-bold text-blue-700' : 'text-slate-700'}`}>
                    <p className="font-semibold">{c.nom}</p>
                    {c.ice && <p className="text-[10px] text-slate-400">ICE: {c.ice}</p>}
                  </button>
                ))
              }
              {/* Allow typing a custom client not in list */}
              {clientSearch && !clientsList.find(c => c.nom?.toLowerCase() === clientSearch.toLowerCase()) && (
                <button type="button"
                  onClick={() => {
                    setSuiviForm(p => ({ ...p, client: clientSearch }));
                    setShowClientDrop(false);
                    setClientSearch('');
                  }}
                  className="w-full text-left px-3 py-2 text-xs hover:bg-emerald-50 text-emerald-700 border-t border-slate-200 font-bold">
                  + Utiliser "{clientSearch}"
                </button>
              )}
            </div>
          )}
          {showClientDrop && (
            <div className="fixed inset-0 z-40" onClick={() => { setShowClientDrop(false); setClientSearch(''); }} />
          )}
        </div>

        {/* MATRICULE — searchable dropdown from fleet_drivers */}
        <div className="relative">
          <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Matricule</label>
          <div className="relative mt-1">
            <input
              type="text"
              value={showMatDrop ? matriculeSearch : (suiviForm.matricule || '')}
              placeholder="Sélectionner ou rechercher..."
              onFocus={() => { setShowMatDrop(true); setMatriculeSearch(suiviForm.matricule || ''); }}
              onChange={e => { setMatriculeSearch(e.target.value); setShowMatDrop(true); }}
              className="w-full h-9 rounded-lg border-2 border-slate-200 px-3 text-sm focus:outline-none focus:border-blue-500 pr-8"
            />
            <span className="absolute right-2 top-1/2 -translate-y-1/2 text-slate-400 pointer-events-none text-xs">▾</span>
          </div>
          {showMatDrop && (
            <div className="absolute z-50 top-full left-0 right-0 bg-white border-2 border-blue-200 rounded-xl shadow-xl max-h-48 overflow-y-auto mt-1">
              {fleetDrivers
                .filter(d =>
                  d.immatriculation?.toLowerCase().includes(matriculeSearch.toLowerCase()) ||
                  d.nom_prenom?.toLowerCase().includes(matriculeSearch.toLowerCase())
                )
                .length === 0 ? (
                <div className="px-3 py-3 text-xs text-slate-400 text-center">Aucun véhicule trouvé</div>
              ) : fleetDrivers
                .filter(d =>
                  d.immatriculation?.toLowerCase().includes(matriculeSearch.toLowerCase()) ||
                  d.nom_prenom?.toLowerCase().includes(matriculeSearch.toLowerCase())
                )
                .map(d => (
                  <button key={d.id} type="button"
                    onClick={() => {
                      setSuiviForm(p => ({ ...p, matricule: d.immatriculation, type: d.type_vehicule || '' }));
                      setShowMatDrop(false);
                      setMatriculeSearch('');
                    }}
                    className={`w-full text-left px-3 py-2 text-xs hover:bg-blue-50 transition-colors border-b border-slate-100 last:border-0 ${suiviForm.matricule === d.immatriculation ? 'bg-blue-50 font-bold text-blue-700' : 'text-slate-700'}`}>
                    <p className="font-semibold font-mono">{d.immatriculation || '—'}</p>
                    <p className="text-[10px] text-slate-400">{d.nom_prenom} — {d.type_vehicule}</p>
                  </button>
                ))
              }
              {/* Allow typing a custom plate not in list */}
              {matriculeSearch && !fleetDrivers.find(d => d.immatriculation?.toLowerCase() === matriculeSearch.toLowerCase()) && (
                <button type="button"
                  onClick={() => {
                    setSuiviForm(p => ({ ...p, matricule: matriculeSearch }));
                    setShowMatDrop(false);
                    setMatriculeSearch('');
                  }}
                  className="w-full text-left px-3 py-2 text-xs hover:bg-emerald-50 text-emerald-700 border-t border-slate-200 font-bold">
                  + Utiliser "{matriculeSearch}"
                </button>
              )}
            </div>
          )}
          {showMatDrop && (
            <div className="fixed inset-0 z-40" onClick={() => { setShowMatDrop(false); setMatriculeSearch(''); }} />
          )}
        </div>

        {/* All other fields — same as before but skip client and matricule */}
        {suiviFields
          .filter(f => f.key !== 'client' && f.key !== 'matricule')
          .map(({ label, key, type }) => (
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
        <button onClick={() => { setShowSuiviForm(false); setEditingSuivi(null); setSuiviForm(emptySuivi); setShowClientDrop(false); setShowMatDrop(false); }}
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
              {/* Fournisseur dropdown */}
              <div>
                <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Fournisseur</label>
                <select value={(editingPurchase as any).fournisseur || ''}
                  onChange={e => {
                    const nom = e.target.value;
                    const found = fournisseursList.find(f => f.nom === nom);
                    setEditingPurchase({ ...editingPurchase!,
                      fournisseur: nom,
                      ...(found ? { if_number: found.if_number || '', ice_number: found.ice || '' } : {}),
                    });
                  }}
                  className="w-full mt-1 h-9 rounded-lg border-2 border-slate-200 px-3 text-sm focus:outline-none focus:border-blue-500">
                  <option value="">— Sélectionner —</option>
                  {fournisseursList.map(f => (
                    <option key={f.id} value={f.nom}>{f.nom}{f.ice ? ` (ICE: ${f.ice})` : ''}</option>
                  ))}
                  {(editingPurchase as any).fournisseur && !fournisseursList.find(f => f.nom === (editingPurchase as any).fournisseur) && (
                    <option value={(editingPurchase as any).fournisseur}>{(editingPurchase as any).fournisseur} (personnalisé)</option>
                  )}
                </select>
              </div>
              {/* Other fields */}
              {[
                { label: 'N° Facture', key: 'numero_facture', type: 'text' },
                { label: 'Date', key: 'date_achat', type: 'date' },
                { label: 'Montant HT', key: 'montant_ht', type: 'number' },
                { label: 'Montant TTC', key: 'montant_ttc', type: 'number' },
                { label: 'Notes', key: 'notes', type: 'text' },
              ].map(({ label, key, type }) => (
                <div key={key}>
                  <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">{label}</label>
                  <input type={type} value={(editingPurchase as any)[key] || ''}
                    onChange={e => setEditingPurchase({ ...editingPurchase!, [key]: type === 'number' ? parseFloat(e.target.value) : e.target.value })}
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
            { label: 'Consommation (L/100km)', key: 'consommation',       type: 'number' },
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
{/* Edit Fournisseur Modal */}
<AnimatePresence>
  {editingFournisseur && (
    <div className="fixed inset-0 bg-black/50 z-50 flex items-center justify-center p-4">
      <motion.div initial={{ scale: 0.95, opacity: 0 }} animate={{ scale: 1, opacity: 1 }}
        exit={{ scale: 0.95, opacity: 0 }}
        className="bg-white rounded-xl p-6 max-w-lg w-full shadow-xl space-y-4">
        <div className="flex items-center justify-between mb-2">
          <h3 className="font-black text-slate-900 uppercase tracking-widest text-sm">Modifier le Fournisseur</h3>
          <button onClick={() => setEditingFournisseur(null)} className="p-1.5 rounded hover:bg-slate-100 text-slate-400"><X size={16} /></button>
        </div>
        {[
          { label: 'Nom Fournisseur',     key: 'nom',            type: 'text'   },
          { label: 'Catégorie',            key: 'categorie',      type: 'text'   },
          { label: 'Taux TVA (ex: 0.20)',  key: 'taux_tva',       type: 'number' },
          { label: 'IF',                   key: 'if_number',      type: 'text'   },
          { label: 'ICE',                  key: 'ice',            type: 'text'   },
          { label: 'Adresse',              key: 'adresse',        type: 'text'   },
          { label: 'Téléphone',            key: 'telephone',      type: 'text'   },
          { label: 'Banque',               key: 'banque',         type: 'text'   },
          { label: 'Délai Paiement (J)',   key: 'delai_paiement', type: 'number' },
        ].map(({ label, key, type }) => (
          <div key={key}>
            <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">{label}</label>
            <input type={type} value={fournisseurEditForm[key] || ''}
              onChange={e => setFournisseurEditForm((p: any) => ({ ...p, [key]: e.target.value }))}
              className="w-full mt-1 h-9 rounded-lg border-2 border-slate-200 px-3 text-sm focus:outline-none focus:border-blue-500" />
          </div>
        ))}
        <div className="flex gap-3 pt-2">
          <button onClick={handleEditFournisseurSave}
            className="flex-1 h-10 bg-blue-600 hover:bg-blue-700 text-white text-sm font-bold rounded-lg cursor-pointer">Enregistrer</button>
          <button onClick={() => setEditingFournisseur(null)}
            className="flex-1 h-10 bg-slate-100 hover:bg-slate-200 text-slate-700 text-sm font-bold rounded-lg cursor-pointer">Annuler</button>
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
        className="bg-white rounded-xl p-6 max-w-5xl w-full shadow-xl max-h-[85vh] overflow-y-auto">
        <div className="flex items-center justify-between mb-4">
          <div>
            <h3 className="font-black text-slate-900 uppercase tracking-widest text-sm">Sélectionner des Prestations</h3>
            <p className="text-xs text-slate-500 mt-1">Cliquez sur une ligne pour une facture simple, ou cochez plusieurs pour une facture groupée.</p>
          </div>
          <button onClick={() => { setPrestationPickerOpen(false); setSelectedPrestations([]); }}
            className="p-1.5 rounded hover:bg-slate-100 text-slate-400"><X size={16} /></button>
        </div>
        {/* Filters */}
        <div className="mb-4 flex flex-wrap gap-3 items-end">
          <div>
            <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Client</label>
            <input type="text" placeholder="Rechercher..." value={prestationFilter.client}
              onChange={e => setPrestationFilter(p => ({ ...p, client: e.target.value }))}
              className="block mt-1 h-8 rounded-lg border-2 border-slate-200 px-3 text-xs focus:outline-none focus:border-blue-500 w-44" />
          </div>
          <div>
            <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Date de</label>
            <input type="date" value={prestationFilter.dateFrom}
              onChange={e => setPrestationFilter(p => ({ ...p, dateFrom: e.target.value }))}
              className="block mt-1 h-8 rounded-lg border-2 border-slate-200 px-3 text-xs focus:outline-none focus:border-blue-500" />
          </div>
          <div>
            <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Date à</label>
            <input type="date" value={prestationFilter.dateTo}
              onChange={e => setPrestationFilter(p => ({ ...p, dateTo: e.target.value }))}
              className="block mt-1 h-8 rounded-lg border-2 border-slate-200 px-3 text-xs focus:outline-none focus:border-blue-500" />
          </div>
          <div>
            <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Matricule</label>
            <input type="text" placeholder="Rechercher..." value={prestationFilter.matricule}
              onChange={e => setPrestationFilter(p => ({ ...p, matricule: e.target.value }))}
              className="block mt-1 h-8 rounded-lg border-2 border-slate-200 px-3 text-xs focus:outline-none focus:border-blue-500 w-36" />
          </div>
          <button onClick={() => setPrestationFilter({ client: '', dateFrom: '', dateTo: '', matricule: '' })}
            className="h-8 px-3 bg-slate-100 hover:bg-slate-200 text-slate-600 text-xs font-bold rounded-lg cursor-pointer">
            Réinitialiser
          </button>
        </div>

        {/* Multi-select toolbar */}
        {selectedPrestations.length > 0 && (
          <div className="mb-4 p-3 bg-blue-50 border border-blue-200 rounded-xl flex items-center justify-between">
            <span className="text-sm font-black text-blue-800">{selectedPrestations.length} prestation(s) sélectionnée(s)</span>
            <div className="flex gap-2">
              <button onClick={() => setSelectedPrestations([])}
                className="px-3 py-1.5 bg-white border border-slate-200 text-slate-600 text-xs font-bold rounded-lg cursor-pointer">
                Tout désélectionner
              </button>
              {selectedPrestations.length >= 2 && (
                <button onClick={handleStartWizard}
                  className="px-4 py-1.5 bg-blue-600 hover:bg-blue-700 text-white text-xs font-black uppercase rounded-lg cursor-pointer flex items-center gap-1.5">
                  <FileText size={12} /> Créer facture groupée ({selectedPrestations.length})
                </button>
              )}
            </div>
          </div>
        )}

        <table className="w-full text-left text-xs">
          <thead className="bg-slate-50 border-b border-slate-200">
            <tr>
              <th className="px-3 py-2 w-8">
                <input type="checkbox"
                  checked={selectedPrestations.length > 0 && suiviList.filter((p: any) => {
                    if (prestationFilter.client && !p.client?.toLowerCase().includes(prestationFilter.client.toLowerCase())) return false;
                    if (prestationFilter.matricule && !p.matricule?.toLowerCase().includes(prestationFilter.matricule.toLowerCase())) return false;
                    if (prestationFilter.dateFrom && (p.date || '') < prestationFilter.dateFrom) return false;
                    if (prestationFilter.dateTo && (p.date || '') > prestationFilter.dateTo) return false;
                    return true;
                  }).every((p: any) => selectedPrestations.includes(p.id))}
                  onChange={e => {
                    const filtered = suiviList.filter((p: any) => {
                      if (prestationFilter.client && !p.client?.toLowerCase().includes(prestationFilter.client.toLowerCase())) return false;
                      if (prestationFilter.matricule && !p.matricule?.toLowerCase().includes(prestationFilter.matricule.toLowerCase())) return false;
                      if (prestationFilter.dateFrom && (p.date || '') < prestationFilter.dateFrom) return false;
                      if (prestationFilter.dateTo && (p.date || '') > prestationFilter.dateTo) return false;
                      return true;
                    });
                    setSelectedPrestations(e.target.checked ? filtered.map((p: any) => p.id) : []);
                  }}
                  className="accent-blue-600" />
              </th>
              {['Date','Matricule','Client','Départ','Arrivée','Prix HT','Prix TTC','BL/OT'].map(h => (
                <th key={h} className="px-3 py-2 text-[9px] font-black text-slate-400 uppercase tracking-widest">{h}</th>
              ))}
            </tr>
          </thead>
          <tbody className="divide-y divide-slate-100">
            {(() => {
              const filtered = suiviList.filter((p: any) => {
                if (prestationFilter.client && !p.client?.toLowerCase().includes(prestationFilter.client.toLowerCase())) return false;
                if (prestationFilter.matricule && !p.matricule?.toLowerCase().includes(prestationFilter.matricule.toLowerCase())) return false;
                if (prestationFilter.dateFrom && (p.date || '') < prestationFilter.dateFrom) return false;
                if (prestationFilter.dateTo && (p.date || '') > prestationFilter.dateTo) return false;
                return true;
              });
              return filtered.length === 0 ? (
                <tr><td colSpan={9} className="px-4 py-6 text-center text-slate-400">{suiviList.length === 0 ? 'Aucune prestation disponible.' : 'Aucun résultat pour ces filtres.'}</td></tr>
              ) : filtered.map((p: any) => (
              <tr key={p.id} className="hover:bg-blue-50 transition-colors">
                <td className="px-3 py-2" onClick={e => e.stopPropagation()}>
                  <input type="checkbox"
                    checked={selectedPrestations.includes(p.id)}
                    onChange={e => {
                      setSelectedPrestations(prev =>
                        e.target.checked ? [...prev, p.id] : prev.filter(id => id !== p.id)
                      );
                    }}
                    className="accent-blue-600" />
                </td>
                <td className="px-3 py-2 cursor-pointer" onClick={() => handleAutoFillFromPrestation(p)}>{p.date}</td>
                <td className="px-3 py-2 font-mono text-blue-600 cursor-pointer" onClick={() => handleAutoFillFromPrestation(p)}>{p.matricule || '—'}</td>
                <td className="px-3 py-2 font-semibold cursor-pointer" onClick={() => handleAutoFillFromPrestation(p)}>{p.client || '—'}</td>
                <td className="px-3 py-2 cursor-pointer" onClick={() => handleAutoFillFromPrestation(p)}>{p.depart || '—'}</td>
                <td className="px-3 py-2 cursor-pointer" onClick={() => handleAutoFillFromPrestation(p)}>{p.arrivee || '—'}</td>
                <td className="px-3 py-2 cursor-pointer" onClick={() => handleAutoFillFromPrestation(p)}>{Number(p.prix_ht).toLocaleString('fr-MA')}</td>
                <td className="px-3 py-2 font-bold cursor-pointer" onClick={() => handleAutoFillFromPrestation(p)}>{Number(p.prix_ttc).toLocaleString('fr-MA')}</td>
                <td className="px-3 py-2 cursor-pointer" onClick={() => handleAutoFillFromPrestation(p)}>{p.ot_bl_bs_be || '—'}</td>
              </tr>
            ));
            })()}
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
{/* Wizard Modal — step through prestations */}
<AnimatePresence>
  {wizardOpen && wizardForms.length > 0 && (
    <div className="fixed inset-0 bg-black/50 z-50 flex items-center justify-center p-4">
      <motion.div initial={{ scale: 0.95, opacity: 0 }} animate={{ scale: 1, opacity: 1 }}
        exit={{ scale: 0.95, opacity: 0 }}
        className="bg-white rounded-xl p-6 max-w-3xl w-full shadow-xl max-h-[90vh] overflow-y-auto">

        {/* Progress bar */}
        <div className="mb-5">
          <div className="flex items-center justify-between mb-2">
            <h3 className="font-black text-slate-900 uppercase tracking-widest text-sm">
              Facture Groupée — Prestation {wizardStep + 1} / {wizardForms.length}
            </h3>
            <button onClick={() => { setWizardOpen(false); setWizardForms([]); }}
              className="p-1.5 rounded hover:bg-slate-100 text-slate-400"><X size={16} /></button>
          </div>
          <div className="w-full bg-slate-200 rounded-full h-2">
            <div className="bg-blue-600 h-2 rounded-full transition-all" style={{ width: `${((wizardStep + 1) / wizardForms.length) * 100}%` }} />
          </div>
          {/* Step dots */}
          <div className="flex gap-1 mt-2 justify-center">
            {wizardForms.map((_: any, i: number) => (
              <button key={i} onClick={() => setWizardStep(i)}
                className={`w-6 h-6 rounded-full text-[9px] font-black cursor-pointer transition-all ${i === wizardStep ? 'bg-blue-600 text-white' : i < wizardStep ? 'bg-emerald-100 text-emerald-700' : 'bg-slate-100 text-slate-400'}`}>
                {i + 1}
              </button>
            ))}
          </div>
        </div>

        {/* Prestation info bar */}
        <div className="mb-4 p-3 bg-slate-50 border border-slate-200 rounded-xl flex items-center justify-between">
          <div>
            <span className="text-xs font-black text-slate-700">{wizardPrestations[wizardStep]?.client}</span>
            <span className="text-xs text-slate-400 ml-2">{wizardPrestations[wizardStep]?.depart} → {wizardPrestations[wizardStep]?.arrivee}</span>
          </div>
          <span className="text-xs font-mono font-bold text-blue-600">{wizardPrestations[wizardStep]?.matricule}</span>
        </div>

        {/* Form fields — same as existing facture form */}
        <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-4">
          {[
            { label: 'Date', key: 'date', type: 'date' },
            { label: 'N° Facture', key: 'numero_facture', type: 'text' },
            { label: 'Client', key: 'client', type: 'text' },
            { label: 'Départ', key: 'depart', type: 'text' },
            { label: 'Arrivée', key: 'arrivee', type: 'text' },
            { label: 'BL / OT', key: 'bl_ot', type: 'text' },
            { label: 'BC', key: 'bc', type: 'text' },
            { label: 'Délai Paiement (J)', key: 'delai_paiement', type: 'number' },
            { label: 'Date de Paiement', key: 'date_paiement', type: 'date' },
            { label: 'Mode Paiement', key: 'mode_paiement', type: 'text' },
          ].map(({ label, key, type }) => (
            <div key={key}>
              <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">{label}</label>
              <input type={type} value={wizardForms[wizardStep]?.[key] || ''}
                onChange={e => {
                  const updated = [...wizardForms];
                  updated[wizardStep] = { ...updated[wizardStep], [key]: e.target.value };
                  setWizardForms(updated);
                }}
                className="w-full mt-1 h-9 rounded-lg border-2 border-slate-200 px-3 text-sm focus:outline-none focus:border-blue-500" />
            </div>
          ))}

          {/* Montant HT */}
          <div>
            <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Montant HT (MAD)</label>
            <input type="number" value={wizardForms[wizardStep]?.montant_ht || ''}
              onChange={e => {
                const updated = [...wizardForms];
                const ht = parseFloat(e.target.value) || 0;
                const rate = parseFloat(updated[wizardStep]?.tva_rate) || 0;
                const tvaAmt = parseFloat((ht * rate / 100).toFixed(2));
                updated[wizardStep] = { ...updated[wizardStep], montant_ht: e.target.value, tva: String(tvaAmt), montant_ttc: String((ht + tvaAmt).toFixed(2)) };
                setWizardForms(updated);
              }}
              className="w-full mt-1 h-9 rounded-lg border-2 border-slate-200 px-3 text-sm focus:outline-none focus:border-blue-500" />
          </div>

          {/* TVA */}
          <div>
            <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">TVA (MAD) — auto</label>
            <input type="number" readOnly value={wizardForms[wizardStep]?.tva || ''}
              className="w-full mt-1 h-9 rounded-lg border-2 border-slate-100 bg-slate-50 px-3 text-sm text-slate-500 cursor-not-allowed" />
          </div>

          {/* TTC */}
          <div>
            <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Montant TTC (MAD) — auto</label>
            <input type="number" readOnly value={wizardForms[wizardStep]?.montant_ttc || ''}
              className="w-full mt-1 h-9 rounded-lg border-2 border-blue-100 bg-blue-50 px-3 text-sm font-bold text-blue-700 cursor-not-allowed" />
          </div>
          {/* Observation */}
          <div className="sm:col-span-2 lg:col-span-3">
            <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Observation (1ère ligne)</label>
            <input type="text" value={wizardForms[wizardStep]?.observation || ''}
              onChange={e => {
                const updated = [...wizardForms];
                updated[wizardStep] = { ...updated[wizardStep], observation: e.target.value };
                setWizardForms(updated);
              }}
              placeholder="ex: Transport marchandises — Lot 45B"
              className="w-full mt-1 h-9 rounded-lg border-2 border-slate-200 px-3 text-sm focus:outline-none focus:border-blue-500" />
          </div>

          {/* Statut */}
          <div>
            <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Statut</label>
            <select value={wizardForms[wizardStep]?.statut || 'impayé'}
              onChange={e => {
                const updated = [...wizardForms];
                updated[wizardStep] = { ...updated[wizardStep], statut: e.target.value };
                setWizardForms(updated);
              }}
              className="w-full mt-1 h-9 rounded-lg border-2 border-slate-200 px-3 text-sm focus:outline-none focus:border-blue-500">
              <option value="impayé">Impayé</option>
              <option value="payé">Payé</option>
            </select>
          </div>
        </div>

        {/* Navigation buttons */}
        <div className="flex gap-3 pt-5 border-t border-slate-200 mt-5">
          <button onClick={() => setWizardStep(s => Math.max(0, s - 1))}
            disabled={wizardStep === 0}
            className="flex-1 h-10 bg-slate-100 hover:bg-slate-200 disabled:opacity-40 text-slate-700 text-sm font-bold rounded-lg cursor-pointer">
            ← Précédent
          </button>

          {wizardStep < wizardForms.length - 1 ? (
            <button onClick={() => setWizardStep(s => s + 1)}
              className="flex-1 h-10 bg-blue-600 hover:bg-blue-700 text-white text-sm font-bold rounded-lg cursor-pointer">
              Suivant →
            </button>
          ) : (
            <button onClick={handleSaveWizard} disabled={savingWizard}
              className="flex-1 h-10 bg-emerald-600 hover:bg-emerald-700 disabled:opacity-60 text-white text-sm font-black rounded-lg cursor-pointer flex items-center justify-center gap-2">
              {savingWizard ? <Loader2 size={14} className="animate-spin" /> : null}
              {savingWizard ? 'Enregistrement...' : `✓ Enregistrer tout (${wizardForms.length} factures)`}
            </button>
          )}

          <button onClick={() => { setWizardOpen(false); setWizardForms([]); }}
            className="h-10 px-4 bg-white border border-slate-200 text-slate-500 text-sm font-bold rounded-lg cursor-pointer">
            Annuler
          </button>
        </div>
      </motion.div>
    </div>
  )}
</AnimatePresence>
{/* Avoir Form Modal */}
<AnimatePresence>
  {showAvoirForm && (
    <div className="fixed inset-0 bg-black/50 z-50 flex items-center justify-center p-4">
      <motion.div initial={{ scale: 0.95, opacity: 0 }} animate={{ scale: 1, opacity: 1 }}
        exit={{ scale: 0.95, opacity: 0 }}
        className="bg-white rounded-xl p-6 max-w-2xl w-full shadow-xl max-h-[90vh] overflow-y-auto">
        <div className="flex items-center justify-between mb-5">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 rounded-lg bg-rose-600 flex items-center justify-center text-white">
              <FileText size={20} />
            </div>
            <div>
              <h3 className="font-black text-slate-900 uppercase tracking-widest text-sm">Facture d'Avoir</h3>
              <p className="text-[10px] text-slate-500 mt-0.5">Les montants seront enregistrés en négatif</p>
            </div>
          </div>
          <button onClick={() => setShowAvoirForm(false)} className="p-1.5 rounded hover:bg-slate-100 text-slate-400"><X size={16} /></button>
        </div>

        <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
          <div>
            <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Date</label>
            <input type="date" value={avoirForm.date}
              onChange={e => setAvoirForm(p => ({ ...p, date: e.target.value }))}
              className="w-full mt-1 h-9 rounded-lg border-2 border-slate-200 px-3 text-sm focus:outline-none focus:border-rose-500" />
          </div>
          <div>
            <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">N° Avoir</label>
            <input type="text" value={avoirForm.numero_facture} placeholder="AV-001"
              onChange={e => setAvoirForm(p => ({ ...p, numero_facture: e.target.value }))}
              className="w-full mt-1 h-9 rounded-lg border-2 border-slate-200 px-3 text-sm focus:outline-none focus:border-rose-500" />
          </div>
          {/* Client dropdown */}
          <div>
            <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Client</label>
            <select value={avoirForm.client}
              onChange={e => setAvoirForm(p => ({ ...p, client: e.target.value }))}
              className="w-full mt-1 h-9 rounded-lg border-2 border-slate-200 px-3 text-sm focus:outline-none focus:border-rose-500">
              <option value="">— Sélectionner —</option>
              {clientsList.map((c: any) => (
                <option key={c.id} value={c.nom}>{c.nom}</option>
              ))}
            </select>
          </div>
          <div>
            <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Désignation / Motif</label>
            <input type="text" value={avoirForm.observation} placeholder="Retour marchandise, erreur facturation..."
              onChange={e => setAvoirForm(p => ({ ...p, observation: e.target.value, designation: e.target.value }))}
              className="w-full mt-1 h-9 rounded-lg border-2 border-slate-200 px-3 text-sm focus:outline-none focus:border-rose-500" />
          </div>

          {/* Montant HT → auto TVA + TTC */}
          <div>
            <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Montant HT (positif, sera inversé)</label>
            <input type="number" value={avoirForm.montant_ht} placeholder="500"
              onChange={e => {
                const ht = parseFloat(e.target.value) || 0;
                const rate = parseFloat(avoirForm.tva) || 0;
                const tvaAmount = parseFloat((ht * rate / 100).toFixed(2));
                setAvoirForm(p => ({ ...p, montant_ht: e.target.value, montant_ttc: String((ht + tvaAmount).toFixed(2)) }));
              }}
              className="w-full mt-1 h-9 rounded-lg border-2 border-slate-200 px-3 text-sm focus:outline-none focus:border-rose-500" />
          </div>
          <div>
            <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">TVA (%)</label>
            <input type="number" value={avoirForm.tva} placeholder="20"
              onChange={e => {
                const rate = parseFloat(e.target.value) || 0;
                const ht = parseFloat(avoirForm.montant_ht) || 0;
                const tvaAmount = parseFloat((ht * rate / 100).toFixed(2));
                setAvoirForm(p => ({ ...p, tva: e.target.value, montant_ttc: String((ht + tvaAmount).toFixed(2)) }));
              }}
              className="w-full mt-1 h-9 rounded-lg border-2 border-slate-200 px-3 text-sm focus:outline-none focus:border-rose-500" />
          </div>
          <div className="sm:col-span-2">
            <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Montant TTC (auto)</label>
            <input type="number" value={avoirForm.montant_ttc} readOnly
              className="w-full mt-1 h-9 rounded-lg border-2 border-rose-200 bg-rose-50 px-3 text-sm font-bold text-rose-700 cursor-not-allowed" />
          </div>
        </div>

        {/* Preview */}
        <div className="mt-4 p-3 bg-rose-50 border border-rose-200 rounded-xl text-sm text-rose-800">
          {(() => { const ht = parseFloat(avoirForm.montant_ht)||0; const rate = parseFloat(avoirForm.tva)||0; const tvaAmt = parseFloat((ht*rate/100).toFixed(2)); return `Sera enregistré comme : HT: -${ht.toFixed(2)} | TVA (${rate}%): -${tvaAmt.toFixed(2)} | TTC: -${avoirForm.montant_ttc || '0'} MAD`; })()}
        </div>

        <div className="flex gap-3 pt-5">
          <button onClick={async () => {
            const ht = parseFloat(avoirForm.montant_ht) || 0;
            const rate = parseFloat(avoirForm.tva) || 0;
              const tvaAmount = parseFloat((ht * rate / 100).toFixed(2));
              const ttc = ht + tvaAmount;
            if (ht <= 0) { toast.error("Saisissez un montant."); return; }
            const { error } = await supabase.from('suivi_facturation').insert({
              company_id: companyId,
              manager_id: managerId || null,
              date: avoirForm.date || null,
              numero_facture: avoirForm.numero_facture || null,
              client: avoirForm.client || null,
              montant_ht: -ht,
              tva: -tvaAmount,
              montant_ttc: -ttc,
              observation: avoirForm.observation || null,
              is_avoir: true,
              was_avoir: true,
              statut: 'avoir',
              designation: avoirForm.observation || null,
            });
            if (!error) {
              toast.success("Facture d'avoir enregistrée.");
              setShowAvoirForm(false);
              setAvoirForm({ date: new Date().toISOString().split('T')[0], numero_facture: '', client: '', depart: '', arrivee: '', montant_ht: '', tva: '', montant_ttc: '', observation: '' });
              fetchFacturation();
            } else toast.error(`Erreur: ${error.message}`);
          }}
            className="flex-1 h-10 bg-rose-600 hover:bg-rose-700 text-white text-sm font-bold rounded-lg cursor-pointer">
            Enregistrer l'avoir
          </button>
          <button onClick={() => setShowAvoirForm(false)}
            className="flex-1 h-10 bg-slate-100 hover:bg-slate-200 text-slate-700 text-sm font-bold rounded-lg cursor-pointer">
            Annuler
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
        {/* Désignation */}
        <div className="sm:col-span-2">
          <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Désignation</label>
          <input type="text" value={(factForm as any).designation || ''}
            onChange={e => setFactForm((p: any) => ({ ...p, designation: e.target.value }))}
            placeholder="ex: Transport matériaux, Avoir sur facture..."
            className="w-full mt-1 h-9 rounded-lg border-2 border-slate-200 px-3 text-sm focus:outline-none focus:border-blue-500" />
        </div>
        {/* Observation */}
        <div className="sm:col-span-2 lg:col-span-3">
          <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Observation (1ère ligne de la facture)</label>
          <input type="text" value={(factForm as any).observation || ''}
            onChange={e => setFactForm((p: any) => ({ ...p, observation: e.target.value }))}
            placeholder="ex: Transport de matériaux de construction — Chantier Tanger Nord"
            className="w-full mt-1 h-9 rounded-lg border-2 border-slate-200 px-3 text-sm focus:outline-none focus:border-blue-500" />
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
        {editingFact?.is_avoir ? (
          <>
            <button onClick={async () => {
              const { error } = await supabase.from('suivi_facturation').update({
                date: factForm.date || null, numero_facture: factForm.numero_facture || null,
                client: factForm.client || null, depart: factForm.depart || null, arrivee: factForm.arrivee || null,
                montant_ht: parseFloat(factForm.montant_ht) || 0, tva: parseFloat(factForm.tva) || 0,
                montant_ttc: parseFloat(factForm.montant_ttc) || 0, bl_ot: factForm.bl_ot || null, bc: factForm.bc || null,
                observation: (factForm as any).observation || null, designation: (factForm as any).designation || null,
                is_avoir: true, was_avoir: true, statut: 'avoir',
              }).eq('id', editingFact.id);
              if (!error) { toast.success("Avoir modifié."); setShowFactForm(false); setEditingFact(null); setFactForm(emptyFactForm); fetchFacturation(); }
              else toast.error(`Erreur: ${error.message}`);
            }}
              className="flex-1 h-10 bg-purple-600 hover:bg-purple-700 text-white text-sm font-bold rounded-lg cursor-pointer">
              Enregistrer comme Avoir
            </button>
            <button onClick={async () => {
              const { error } = await supabase.from('suivi_facturation').update({
                date: factForm.date || null, numero_facture: factForm.numero_facture || null,
                client: factForm.client || null, depart: factForm.depart || null, arrivee: factForm.arrivee || null,
                montant_ht: Math.abs(parseFloat(factForm.montant_ht) || 0), tva: Math.abs(parseFloat(factForm.tva) || 0),
                montant_ttc: Math.abs(parseFloat(factForm.montant_ttc) || 0), bl_ot: factForm.bl_ot || null, bc: factForm.bc || null,
                observation: (factForm as any).observation || null, designation: (factForm as any).designation || null,
                is_avoir: false, was_avoir: true, statut: 'impayé',
              }).eq('id', editingFact.id);
              if (!error) { toast.success("Converti en facture."); setShowFactForm(false); setEditingFact(null); setFactForm(emptyFactForm); fetchFacturation(); }
              else toast.error(`Erreur: ${error.message}`);
            }}
              className="flex-1 h-10 bg-blue-600 hover:bg-blue-700 text-white text-sm font-bold rounded-lg cursor-pointer">
              Convertir en Facture
            </button>
          </>
        ) : editingFact?.was_avoir && !editingFact?.is_avoir ? (
          <>
            <button onClick={handleSaveFacturation}
              className="flex-1 h-10 bg-blue-600 hover:bg-blue-700 text-white text-sm font-bold rounded-lg cursor-pointer">
              Enregistrer les modifications
            </button>
            <button onClick={async () => {
              const { error } = await supabase.from('suivi_facturation').update({
                date: factForm.date || null, numero_facture: factForm.numero_facture || null,
                client: factForm.client || null, depart: factForm.depart || null, arrivee: factForm.arrivee || null,
                montant_ht: -(Math.abs(parseFloat(factForm.montant_ht) || 0)), tva: -(Math.abs(parseFloat(factForm.tva) || 0)),
                montant_ttc: -(Math.abs(parseFloat(factForm.montant_ttc) || 0)), bl_ot: factForm.bl_ot || null, bc: factForm.bc || null,
                observation: (factForm as any).observation || null, designation: (factForm as any).designation || null,
                is_avoir: true, was_avoir: true, statut: 'avoir',
              }).eq('id', editingFact.id);
              if (!error) { toast.success("Revenu en avoir."); setShowFactForm(false); setEditingFact(null); setFactForm(emptyFactForm); fetchFacturation(); }
              else toast.error(`Erreur: ${error.message}`);
            }}
              className="flex-1 h-10 bg-purple-600 hover:bg-purple-700 text-white text-sm font-bold rounded-lg cursor-pointer">
              Revenir en Avoir
            </button>
          </>
        ) : (
          <button onClick={handleSaveFacturation}
            className="flex-1 h-10 bg-blue-600 hover:bg-blue-700 text-white text-sm font-bold rounded-lg cursor-pointer">
            {editingFact ? 'Enregistrer les modifications' : 'Ajouter la facture'}
          </button>
        )}
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