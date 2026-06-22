import { useState, useEffect } from 'react';
import { useNavigate } from 'react-router-dom';
import { Loader2, LogOut, Users, ShoppingBag, Wrench, Menu, X, BadgeCheck, RefreshCw, Plus, Eye, Download, FileText, Pencil, Trash2, Truck, Upload, Receipt, Settings, TrendingUp, FolderOpen, Check, Landmark, Search } from 'lucide-react';
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

type ManagerTab = 'staff' | 'purchases' | 'fleetfix' | 'suivi' | 'chauffeurs' | 'cout_revient' | 'clients' | 'fournisseurs' | 'truck_docs' | 'facturation' | 'settings' |'devis' | 'bon_commande' | 'reglements'| 'bank_rip' | 'bank_releve' | 'bank_base_rip' | 'bank_virement' | 'bank_etat_explicatif' | 'bank_tva';

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
  const [activeTab, setActiveTab] = useState<ManagerTab>('dashboard');
  const [bankOpen, setBankOpen] = useState(false);
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
  // Devis state
  const [devisList, setDevisList] = useState<any[]>([]);
  const [loadingDevis, setLoadingDevis] = useState(false);
  const [showDevisForm, setShowDevisForm] = useState(false);
  const [editingDevis, setEditingDevis] = useState<any>(null);
  const [selectedDevis, setSelectedDevis] = useState<string[]>([]);
  const [devisFilter, setDevisFilter] = useState({ client: '', dateFrom: '', dateTo: '', statut: '' });
  const [devisTemplateId, setDevisTemplateId] = useState('');
  const emptyDevisForm = {
    numero_devis: '', date: new Date().toISOString().split('T')[0], client: '',
    personne_contact: '', type_vehicule: '', designation: '', depart: '', arrivee: '',
    quantite: '1', prix_unitaire_ht: '', tva_rate: '20', montant_ht: '', tva_amount: '', montant_ttc: '',
    observation: '', statut: 'en_attente',
  };
  const [devisForm, setDevisForm] = useState<any>(emptyDevisForm);
  const [avoirForm, setAvoirForm] = useState({ date: new Date().toISOString().split('T')[0], numero_facture: '', client: '', depart: '', arrivee: '', montant_ht: '', tva: '', montant_ttc: '', observation: '' });
  // Bon de Commande state
  const [bcList, setBcList] = useState<any[]>([]);
  const [loadingBC, setLoadingBC] = useState(false);
  const [showBCForm, setShowBCForm] = useState(false);
  const [editingBC, setEditingBC] = useState<any>(null);
  const [selectedBC, setSelectedBC] = useState<string[]>([]);
  const [bcFilter, setBcFilter] = useState({ fournisseur: '', dateFrom: '', dateTo: '', statut: '' });
  const emptyBCForm = {
    numero_bc: '', date: new Date().toISOString().split('T')[0], fournisseur: '',
    personne_contact: '', reference: '', designation: '',
    quantite: '1', prix_unitaire_ht: '', tva_rate: '20', montant_ht: '', tva_amount: '', montant_ttc: '',
    observation: '', statut: 'en_attente',
  };
  const [bcForm, setBcForm] = useState<any>(emptyBCForm);
  // Règlement state
  const [reglementsList, setReglementsList] = useState<any[]>([]);
  const [loadingReglements, setLoadingReglements] = useState(false);
  const [showReglementForm, setShowReglementForm] = useState(false);
  const [reglementFilter, setReglementFilter] = useState({ type: '', banque: '', dateFrom: '', dateTo: '' });
  const [expandedReglement, setExpandedReglement] = useState<string | null>(null);
  const [editingReglement, setEditingReglement] = useState<any>(null);
  const [showEditReglementForm, setShowEditReglementForm] = useState(false);
  const [savingReglement, setSavingReglement] = useState(false);
  const [reglementForm, setReglementForm] = useState<any>({
    type_reglement: 'cheque', date_reglement: new Date().toISOString().split('T')[0],
    numero: '', banque: '', date_echeance: '', recu_par: '', reference_virement: '', observation: '', scanFile: null, tva_mois: '', client: '', code_reglement: '',
  });
  // Bank RIP state
  const [ripList, setRipList] = useState<any[]>([]);
  const [loadingRip, setLoadingRip] = useState(false);
  const [showRipForm, setShowRipForm] = useState(false);
  const [editingRip, setEditingRip] = useState<any>(null);
  const emptyRipForm = { banque: '', agence: '', code_banque: '', code_ville: '', numero_compte: '', cle_rip: '', code_swift: '' };
  const [ripForm, setRipForm] = useState<any>(emptyRipForm);

  // Bank Base RIP state
  const [baseRipList, setBaseRipList] = useState<any[]>([]);
  const [loadingBaseRip, setLoadingBaseRip] = useState(false);
  const [showBaseRipForm, setShowBaseRipForm] = useState(false);
  const [editingBaseRip, setEditingBaseRip] = useState<any>(null);
  const emptyBaseRipForm = { raison_social: '', rib: '' };
  const [baseRipForm, setBaseRipForm] = useState<any>(emptyBaseRipForm);

  // Bank Virement state
  const [virementList, setVirementList] = useState<any[]>([]);
  const [loadingVirement, setLoadingVirement] = useState(false);
  const [showVirementForm, setShowVirementForm] = useState(false);
  const [editingVirement, setEditingVirement] = useState<any>(null);
  const [selectedVirements, setSelectedVirements] = useState<string[]>([]);
  const [virementFilter, setVirementFilter] = useState({ raison: '', dateFrom: '', dateTo: '' });
  const emptyVirementForm = { date_virement: new Date().toISOString().split('T')[0], raison_social: '', rib: '', montant: '', justification: '', tva_mois: '', compte_id: '', banque: '', agence: '', numero_compte: '' };
  const [virementForm, setVirementForm] = useState<any>(emptyVirementForm);
  // Relevé Bancaire + État Explicatif state
  const [releveList, setReleveList] = useState<any[]>([]);
  const [loadingReleve, setLoadingReleve] = useState(false);
  const [releveMois, setReleveMois] = useState(new Date().toISOString().slice(0, 7));
  const [parsingPDF, setParsingPDF] = useState(false);
  const [showReleveForm, setShowReleveForm] = useState(false);
  const [editingReleve, setEditingReleve] = useState<any>(null);
  const [releveFilter, setReleveFilter] = useState({ mois: '', libelle: '', category: '' });
  const emptyReleveForm = { code: '', date_operation: '', libelle: '', date_valeur: '', debit: '', credit: '', category: 'virement', destination: '', note_operation: '', ref_reglement: '', observation: '', code_reglement: '' };
  const [releveForm, setReleveForm] = useState<any>(emptyReleveForm);
  const [etatExpandedSections, setEtatExpandedSections] = useState<string[]>(['virement','emission_cheque','emission_effets','remise_cheque','remise_lc','frais_bancaire']);
  const [checkedReleve, setCheckedReleve] = useState<string[]>([]);
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
  const handleSavePurchase = async () => {
    if (!editingPurchase) return;
    const p = editingPurchase as any;
    // Only send fields that have values, let Supabase ignore unknown ones
    const payload: any = {};
    Object.keys(p).forEach(key => {
      if (key !== 'id' && key !== 'created_at') payload[key] = p[key];
    });
    const { error } = await supabase.from('purchases').update(payload).eq('id', p.id);
    if (!error) { toast.success("Achat modifié."); setEditingPurchase(null); fetchPurchases(); }
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
    if ((factFilter as any).numero && !f.numero_facture?.toLowerCase().includes((factFilter as any).numero.toLowerCase())) return false;
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
  // ── Devis CRUD ──
  const fetchDevis = async () => {
    if (!companyId) return;
    setLoadingDevis(true);
    const { data } = await supabase.from('devis').select('*').eq('company_id', companyId).order('created_at', { ascending: false });
    setDevisList(data || []);
    setLoadingDevis(false);
  };

  const handleSaveDevis = async () => {
    const ht = parseFloat(devisForm.prix_unitaire_ht) || 0;
    const qty = parseFloat(devisForm.quantite) || 1;
    const rate = parseFloat(devisForm.tva_rate) || 0;
    const montantHT = ht * qty;
    const tvaAmt = parseFloat((montantHT * rate / 100).toFixed(2));
    const ttc = montantHT + tvaAmt;

    const payload = {
      company_id: companyId,
      numero_devis: devisForm.numero_devis || null,
      date: devisForm.date || null,
      client: devisForm.client || null,
      personne_contact: devisForm.personne_contact || null,
      type_vehicule: devisForm.type_vehicule || null,
      designation: devisForm.designation || null,
      depart: devisForm.depart || null,
      arrivee: devisForm.arrivee || null,
      quantite: qty,
      prix_unitaire_ht: ht,
      tva_rate: rate,
      montant_ht: montantHT,
      tva_amount: tvaAmt,
      montant_ttc: ttc,
      observation: devisForm.observation || null,
      statut: devisForm.statut || 'en_attente',
    };

    if (editingDevis) {
      const { error } = await supabase.from('devis').update(payload).eq('id', editingDevis.id);
      if (!error) { toast.success("Devis modifié."); setEditingDevis(null); }
      else { toast.error(`Erreur: ${error.message}`); return; }
    } else {
      const { error } = await supabase.from('devis').insert(payload);
      if (!error) toast.success("Devis ajouté.");
      else { toast.error(`Erreur: ${error.message}`); return; }
    }
    setShowDevisForm(false);
    setDevisForm(emptyDevisForm);
    fetchDevis();
  };

  const handleDeleteDevis = async (id: string) => {
    if (!confirm('Supprimer ce devis ?')) return;
    await supabase.from('devis').delete().eq('id', id);
    toast.success("Supprimé.");
    fetchDevis();
  };

  const filteredDevis = devisList.filter((d: any) => {
    if (devisFilter.client && !d.client?.toLowerCase().includes(devisFilter.client.toLowerCase())) return false;
    if (devisFilter.statut && d.statut !== devisFilter.statut) return false;
    if (devisFilter.dateFrom && (d.date || '') < devisFilter.dateFrom) return false;
    if (devisFilter.dateTo && (d.date || '') > devisFilter.dateTo) return false;
    return true;
  });

  const handleGenerateDevisPDF = () => {
    const selected = devisList.filter((d: any) => selectedDevis.includes(d.id));
    if (selected.length === 0) return;

    const s = invoiceSettings;
    const totalHT = selected.reduce((sum: number, d: any) => sum + (parseFloat(d.montant_ht) || 0), 0);
    const totalTVA = selected.reduce((sum: number, d: any) => sum + (parseFloat(d.tva_amount) || 0), 0);
    const totalTTC = selected.reduce((sum: number, d: any) => sum + (parseFloat(d.montant_ttc) || 0), 0);
    const fmt = (n: number) => n.toLocaleString('fr-MA', { minimumFractionDigits: 2, maximumFractionDigits: 2 });

    const clientName = selected[0]?.client || '';
    const clientData = clientsList.find((c: any) => c.nom === clientName);
    const clientAddress = clientData?.adresse || '';
    const clientICE = clientData?.ice || '';

    const ROWS_PER_PAGE = 15;
    const pages: { rows: any[]; isFirst: boolean; isLast: boolean; num: number }[] = [];
    const remaining = [...selected];
    let pageNum = 0;
    while (remaining.length > 0) {
      pageNum++;
      const isFirst = pageNum === 1;
      const pageRows = remaining.splice(0, ROWS_PER_PAGE);
      pages.push({ rows: pageRows, isFirst, isLast: remaining.length === 0, num: pageNum });
    }
    const totalPages = pages.length;

    const theadHTML = `<thead><tr>
      <th style="background:#1F3864;color:#fff;font-size:10px;font-weight:700;text-align:center;padding:7px 6px;text-transform:uppercase;border:1px solid #1F3864;width:10%">Type</th>
      <th style="background:#1F3864;color:#fff;font-size:10px;font-weight:700;text-align:center;padding:7px 6px;text-transform:uppercase;border:1px solid #1F3864;width:28%">Désignation</th>
      <th style="background:#1F3864;color:#fff;font-size:10px;font-weight:700;text-align:center;padding:7px 6px;text-transform:uppercase;border:1px solid #1F3864;width:10%">Quantité</th>
      <th style="background:#1F3864;color:#fff;font-size:10px;font-weight:700;text-align:center;padding:7px 6px;text-transform:uppercase;border:1px solid #1F3864;width:18%">Prix unitaire HT</th>
      <th style="background:#1F3864;color:#fff;font-size:10px;font-weight:700;text-align:center;padding:7px 6px;text-transform:uppercase;border:1px solid #1F3864;width:10%">TVA %</th>
      <th style="background:#1F3864;color:#fff;font-size:10px;font-weight:700;text-align:center;padding:7px 6px;text-transform:uppercase;border:1px solid #1F3864;width:16%">Montant HT</th>
    </tr></thead>`;

    const pagesHTML = pages.map(p => {
      const rowsHtml = p.rows.map((d: any, i: number) => {
        const bg = i % 2 === 1 ? '#F2F2F2' : '#fff';
        return `<tr>
          <td style="padding:6px 8px;border:1px solid #e5e5e5;text-align:center;font-size:11px;background:${bg}">${d.type_vehicule || ''}</td>
          <td style="padding:6px 8px;border:1px solid #e5e5e5;text-align:left;font-size:11px;background:${bg}">${d.designation || `${d.depart||''} → ${d.arrivee||''}`}</td>
          <td style="padding:6px 8px;border:1px solid #e5e5e5;text-align:center;font-size:11px;background:${bg}">${d.quantite || 1}</td>
          <td style="padding:6px 8px;border:1px solid #e5e5e5;text-align:right;font-size:11px;background:${bg}">${fmt(parseFloat(d.prix_unitaire_ht)||0)}</td>
          <td style="padding:6px 8px;border:1px solid #e5e5e5;text-align:center;font-size:11px;background:${bg}">${d.tva_rate||0}%</td>
          <td style="padding:6px 8px;border:1px solid #e5e5e5;text-align:right;font-size:11px;background:${bg}">${fmt(parseFloat(d.montant_ht)||0)}</td>
        </tr>`;
      }).join('');
      const maxRows = p.isFirst ? ROWS_PER_PAGE : ROWS_PER_PAGE + 6;
      const emptyCount = Math.max(0, (p.isLast ? maxRows - 4 : maxRows) - p.rows.length);
      const emptyHtml = Array.from({ length: emptyCount }, (_, i) => {
        const bg = (p.rows.length + i) % 2 === 1 ? '#F2F2F2' : '#fff';
        return `<tr>${Array(6).fill(`<td style="padding:6px 8px;border:1px solid #e5e5e5;font-size:11px;background:${bg}">&nbsp;</td>`).join('')}</tr>`;
      }).join('');

      return `<div style="width:210mm;min-height:297mm;padding:${p.isFirst?'35mm':'15mm'} 12mm 15mm 12mm;position:relative;font-family:Arial,sans-serif;page-break-after:${p.isLast?'auto':'always'}">
        <div style="position:absolute;top:8mm;right:12mm;font-size:8px;color:#999">Page ${p.num} / ${totalPages}</div>
        ${p.isFirst ? `
          <div style="display:flex;justify-content:flex-end;margin-bottom:16px">
            <div style="border:1px solid #bbb;border-left:4px solid #D4A017;padding:12px 18px;min-width:300px;background:#FAFAFA;border-radius:0 4px 4px 0">
              <div style="font-size:13px;font-weight:700;color:#1e293b">${clientName}</div>
              ${clientAddress?`<div style="font-size:10px;color:#555;margin-top:3px">${clientAddress}</div>`:''}
              ${clientICE?`<div style="font-size:10px;color:#555;margin-top:2px">ICE: ${clientICE}</div>`:''}
            </div>
          </div>
          <table style="width:100%;border-collapse:collapse;margin-bottom:12px">
            <tr>
              <th style="background:#1F3864;color:#fff;font-size:10px;font-weight:700;text-align:center;padding:7px 8px;border:1px solid #1F3864">Devis N°</th>
              <th style="background:#1F3864;color:#fff;font-size:10px;font-weight:700;text-align:center;padding:7px 8px;border:1px solid #1F3864">Date</th>
              <th style="background:#1F3864;color:#fff;font-size:10px;font-weight:700;text-align:center;padding:7px 8px;border:1px solid #1F3864">Personne à contacter</th>
            </tr>
            <tr>
              <td style="text-align:center;padding:7px 8px;font-size:11px;border:1px solid #ddd">${selected[0]?.numero_devis||''}</td>
              <td style="text-align:center;padding:7px 8px;font-size:11px;border:1px solid #ddd">${selected[0]?.date||''}</td>
              <td style="text-align:center;padding:7px 8px;font-size:11px;border:1px solid #ddd">${selected[0]?.personne_contact||''}</td>
            </tr>
          </table>
        `:''}
        <table style="width:100%;border-collapse:collapse">${theadHTML}<tbody>${rowsHtml}${emptyHtml}</tbody></table>
        ${p.isLast?`
          <div style="display:flex;justify-content:flex-end;margin-top:0">
            <div style="width:280px;border:1px solid #ddd">
              <div style="display:flex;justify-content:space-between;padding:5px 12px;font-size:11px;font-weight:700;border-bottom:1px solid #eee"><span>Sous-total HT</span><span>${fmt(totalHT)} MAD</span></div>
              <div style="display:flex;justify-content:space-between;padding:5px 12px;font-size:11px;font-weight:700;border-bottom:1px solid #eee"><span>TVA</span><span>${fmt(totalTVA)} MAD</span></div>
              <div style="display:flex;justify-content:space-between;padding:5px 12px;font-size:11px;font-weight:700;border-bottom:1px solid #eee"><span>Remise</span><span>0,00 MAD</span></div>
              <div style="display:flex;justify-content:space-between;padding:7px 12px;font-size:13px;font-weight:900;background:#1F3864;color:#fff"><span>TOTAL TTC</span><span>${fmt(totalTTC)} MAD</span></div>
            </div>
          </div>
          <div style="margin-top:14px;font-size:10px;color:#7F7F7F"><strong style="color:#333">Arrêté le présent devis à la somme de :</strong> ${numberToWords(totalTTC)}</div>
        `:''}
      </div>`;
    }).join('');

    const html = `<!DOCTYPE html><html lang="fr"><head><meta charset="UTF-8"/>
      <style>*{margin:0;padding:0;box-sizing:border-box}@page{margin:0;size:A4}@media print{body{print-color-adjust:exact;-webkit-print-color-adjust:exact}}</style>
      </head><body>${pagesHTML}</body></html>`;
    const win = window.open('', '_blank');
    if (win) { win.document.write(html); win.document.close(); win.focus(); setTimeout(() => win.print(), 600); }
  };
  // ── Bon de Commande CRUD ──
  const fetchBC = async () => {
    if (!companyId) return;
    setLoadingBC(true);
    const { data } = await supabase.from('bon_commande').select('*').eq('company_id', companyId).order('created_at', { ascending: false });
    setBcList(data || []);
    setLoadingBC(false);
  };

  const handleSaveBC = async () => {
    const ht = parseFloat(bcForm.prix_unitaire_ht) || 0;
    const qty = parseFloat(bcForm.quantite) || 1;
    const rate = parseFloat(bcForm.tva_rate) || 0;
    const montantHT = ht * qty;
    const tvaAmt = parseFloat((montantHT * rate / 100).toFixed(2));
    const payload = {
      company_id: companyId, numero_bc: bcForm.numero_bc || null, date: bcForm.date || null,
      fournisseur: bcForm.fournisseur || null, personne_contact: bcForm.personne_contact || null,
      reference: bcForm.reference || null, designation: bcForm.designation || null,
      quantite: qty, prix_unitaire_ht: ht, tva_rate: rate,
      montant_ht: montantHT, tva_amount: tvaAmt, montant_ttc: montantHT + tvaAmt,
      observation: bcForm.observation || null, statut: bcForm.statut || 'en_attente',
    };
    if (editingBC) {
      const { error } = await supabase.from('bon_commande').update(payload).eq('id', editingBC.id);
      if (!error) { toast.success("BC modifié."); setEditingBC(null); } else { toast.error(`Erreur: ${error.message}`); return; }
    } else {
      const { error } = await supabase.from('bon_commande').insert(payload);
      if (!error) toast.success("BC ajouté."); else { toast.error(`Erreur: ${error.message}`); return; }
    }
    setShowBCForm(false); setBcForm(emptyBCForm); fetchBC();
  };

  const handleDeleteBC = async (id: string) => {
    if (!confirm('Supprimer ce bon de commande ?')) return;
    await supabase.from('bon_commande').delete().eq('id', id);
    toast.success("Supprimé."); fetchBC();
  };

  const filteredBC = bcList.filter((d: any) => {
    if (bcFilter.fournisseur && !d.fournisseur?.toLowerCase().includes(bcFilter.fournisseur.toLowerCase())) return false;
    if (bcFilter.statut && d.statut !== bcFilter.statut) return false;
    if (bcFilter.dateFrom && (d.date || '') < bcFilter.dateFrom) return false;
    if (bcFilter.dateTo && (d.date || '') > bcFilter.dateTo) return false;
    return true;
  });

  const handleGenerateBCPDF = () => {
    const selected = bcList.filter((d: any) => selectedBC.includes(d.id));
    if (selected.length === 0) return;
    const s = invoiceSettings;
    const totalHT = selected.reduce((sum: number, d: any) => sum + (parseFloat(d.montant_ht) || 0), 0);
    const totalTVA = selected.reduce((sum: number, d: any) => sum + (parseFloat(d.tva_amount) || 0), 0);
    const totalTTC = selected.reduce((sum: number, d: any) => sum + (parseFloat(d.montant_ttc) || 0), 0);
    const fmt = (n: number) => n.toLocaleString('fr-MA', { minimumFractionDigits: 2, maximumFractionDigits: 2 });

    const fournName = selected[0]?.fournisseur || '';
    const fournData = fournisseursList.find((f: any) => f.nom === fournName);
    const fournAddress = fournData?.adresse || '';
    const fournICE = fournData?.ice || '';

    const ROWS_PER_PAGE = 15;
    const pages: { rows: any[]; isFirst: boolean; isLast: boolean; num: number }[] = [];
    const remaining = [...selected];
    let pageNum = 0;
    while (remaining.length > 0) {
      pageNum++;
      const capacity = pageNum === 1 ? ROWS_PER_PAGE : ROWS_PER_PAGE + 6;
      const pageRows = remaining.splice(0, capacity);
      pages.push({ rows: pageRows, isFirst: pageNum === 1, isLast: remaining.length === 0, num: pageNum });
    }
    const totalPages = pages.length;

    const theadHTML = `<thead><tr>
      <th style="background:#1F3864;color:#fff;font-size:10px;font-weight:700;text-align:center;padding:7px 6px;text-transform:uppercase;border:1px solid #1F3864;width:12%">Référence</th>
      <th style="background:#1F3864;color:#fff;font-size:10px;font-weight:700;text-align:center;padding:7px 6px;text-transform:uppercase;border:1px solid #1F3864;width:30%">Désignation</th>
      <th style="background:#1F3864;color:#fff;font-size:10px;font-weight:700;text-align:center;padding:7px 6px;text-transform:uppercase;border:1px solid #1F3864;width:10%">Quantité</th>
      <th style="background:#1F3864;color:#fff;font-size:10px;font-weight:700;text-align:center;padding:7px 6px;text-transform:uppercase;border:1px solid #1F3864;width:18%">Prix unitaire HT</th>
      <th style="background:#1F3864;color:#fff;font-size:10px;font-weight:700;text-align:center;padding:7px 6px;text-transform:uppercase;border:1px solid #1F3864;width:10%">TVA %</th>
      <th style="background:#1F3864;color:#fff;font-size:10px;font-weight:700;text-align:center;padding:7px 6px;text-transform:uppercase;border:1px solid #1F3864;width:16%">Montant HT</th>
    </tr></thead>`;

    const pagesHTML = pages.map(p => {
      const rowsHtml = p.rows.map((d: any, i: number) => {
        const bg = i % 2 === 1 ? '#F2F2F2' : '#fff';
        return `<tr>
          <td style="padding:6px 8px;border:1px solid #e5e5e5;text-align:center;font-size:11px;background:${bg}">${d.reference || ''}</td>
          <td style="padding:6px 8px;border:1px solid #e5e5e5;text-align:left;font-size:11px;background:${bg}">${d.designation || ''}</td>
          <td style="padding:6px 8px;border:1px solid #e5e5e5;text-align:center;font-size:11px;background:${bg}">${d.quantite || 1}</td>
          <td style="padding:6px 8px;border:1px solid #e5e5e5;text-align:right;font-size:11px;background:${bg}">${fmt(parseFloat(d.prix_unitaire_ht)||0)}</td>
          <td style="padding:6px 8px;border:1px solid #e5e5e5;text-align:center;font-size:11px;background:${bg}">${d.tva_rate||0}%</td>
          <td style="padding:6px 8px;border:1px solid #e5e5e5;text-align:right;font-size:11px;background:${bg}">${fmt(parseFloat(d.montant_ht)||0)}</td>
        </tr>`;
      }).join('');
      const maxRows = p.isFirst ? ROWS_PER_PAGE : ROWS_PER_PAGE + 6;
      const emptyCount = Math.max(0, (p.isLast ? maxRows - 4 : maxRows) - p.rows.length);
      const emptyHtml = Array.from({ length: emptyCount }, (_, i) => {
        const bg = (p.rows.length + i) % 2 === 1 ? '#F2F2F2' : '#fff';
        return `<tr>${Array(6).fill(`<td style="padding:6px 8px;border:1px solid #e5e5e5;font-size:11px;background:${bg}">&nbsp;</td>`).join('')}</tr>`;
      }).join('');

      return `<div style="width:210mm;min-height:297mm;padding:${p.isFirst?'35mm':'15mm'} 12mm 15mm 12mm;position:relative;font-family:Arial,sans-serif;page-break-after:${p.isLast?'auto':'always'}">
        <div style="position:absolute;top:8mm;right:12mm;font-size:8px;color:#999">Page ${p.num} / ${totalPages}</div>
        ${p.isFirst ? `
          <div style="display:flex;justify-content:flex-end;margin-bottom:16px">
            <div style="border:1px solid #bbb;border-left:4px solid #D4A017;padding:12px 18px;min-width:300px;background:#FAFAFA;border-radius:0 4px 4px 0">
              <div style="font-size:13px;font-weight:700;color:#1e293b">${fournName}</div>
              ${fournAddress?`<div style="font-size:10px;color:#555;margin-top:3px">${fournAddress}</div>`:''}
              ${fournICE?`<div style="font-size:10px;color:#555;margin-top:2px">ICE: ${fournICE}</div>`:''}
            </div>
          </div>
          <table style="width:100%;border-collapse:collapse;margin-bottom:12px">
            <tr>
              <th style="background:#1F3864;color:#fff;font-size:10px;font-weight:700;text-align:center;padding:7px 8px;border:1px solid #1F3864">Bon de commande N°</th>
              <th style="background:#1F3864;color:#fff;font-size:10px;font-weight:700;text-align:center;padding:7px 8px;border:1px solid #1F3864">Date</th>
              <th style="background:#1F3864;color:#fff;font-size:10px;font-weight:700;text-align:center;padding:7px 8px;border:1px solid #1F3864">Personne à contacter</th>
            </tr>
            <tr>
              <td style="text-align:center;padding:7px 8px;font-size:11px;border:1px solid #ddd">${selected[0]?.numero_bc||''}</td>
              <td style="text-align:center;padding:7px 8px;font-size:11px;border:1px solid #ddd">${selected[0]?.date||''}</td>
              <td style="text-align:center;padding:7px 8px;font-size:11px;border:1px solid #ddd">${selected[0]?.personne_contact||''}</td>
            </tr>
          </table>
        `:''}
        <table style="width:100%;border-collapse:collapse">${theadHTML}<tbody>${rowsHtml}${emptyHtml}</tbody></table>
        ${p.isLast?`
          <div style="display:flex;justify-content:flex-end;margin-top:0">
            <div style="width:280px;border:1px solid #ddd">
              <div style="display:flex;justify-content:space-between;padding:5px 12px;font-size:11px;font-weight:700;border-bottom:1px solid #eee"><span>Sous-total HT</span><span>${fmt(totalHT)} MAD</span></div>
              <div style="display:flex;justify-content:space-between;padding:5px 12px;font-size:11px;font-weight:700;border-bottom:1px solid #eee"><span>TVA</span><span>${fmt(totalTVA)} MAD</span></div>
              <div style="display:flex;justify-content:space-between;padding:5px 12px;font-size:11px;font-weight:700;border-bottom:1px solid #eee"><span>Remise</span><span>0,00 MAD</span></div>
              <div style="display:flex;justify-content:space-between;padding:7px 12px;font-size:13px;font-weight:900;background:#1F3864;color:#fff"><span>TOTAL TTC</span><span>${fmt(totalTTC)} MAD</span></div>
            </div>
          </div>
          <div style="margin-top:14px;font-size:10px;color:#7F7F7F"><strong style="color:#333">Arrêté le présent bon de commande à la somme de :</strong> ${numberToWords(totalTTC)}</div>
        `:''}
      </div>`;
    }).join('');

    const html = `<!DOCTYPE html><html lang="fr"><head><meta charset="UTF-8"/>
      <style>*{margin:0;padding:0;box-sizing:border-box}@page{margin:0;size:A4}@media print{body{print-color-adjust:exact;-webkit-print-color-adjust:exact}}</style>
      </head><body>${pagesHTML}</body></html>`;
    const win = window.open('', '_blank');
    if (win) { win.document.write(html); win.document.close(); win.focus(); setTimeout(() => win.print(), 600); }
  };
  // ── Règlement CRUD ──
  const fetchReglements = async () => {
    if (!companyId) return;
    setLoadingReglements(true);
    const { data } = await supabase.from('reglements').select('*').eq('company_id', companyId).order('created_at', { ascending: false });
    setReglementsList(data || []);
    setLoadingReglements(false);
  };

  const handleSaveReglement = async () => {
    const selected = facturationList.filter((f: any) => selectedFacts.includes(f.id));
    if (selected.length === 0) return;
    setSavingReglement(true);

    try {
      // Upload scan if provided
      let scanUrl = '';
      let scanPath = '';
      if (reglementForm.scanFile) {
        const ext = reglementForm.scanFile.name.split('.').pop();
        scanPath = `${companyId}/${Date.now()}.${ext}`;
        const { error: upErr } = await supabase.storage.from('reglements').upload(scanPath, reglementForm.scanFile);
        if (!upErr) {
          const { data: urlData } = supabase.storage.from('reglements').getPublicUrl(scanPath);
          scanUrl = urlData?.publicUrl || '';
        }
      }
      const clientName = [...new Set(selected.map((f: any) => f.client).filter(Boolean))].join(', ');

      const totalMontant = selected.reduce((s: number, f: any) => s + (parseFloat(f.montant_ttc) || 0), 0);
      const factIds = selected.map((f: any) => f.id);
      const factNums = selected.map((f: any) => f.numero_facture || '—');

      // Insert reglement
      const { error } = await supabase.from('reglements').insert({
        company_id: companyId,
        date_reglement: reglementForm.date_reglement || null,
        type_reglement: reglementForm.type_reglement,
        numero: reglementForm.numero || null,
        banque: reglementForm.banque || null,
        date_echeance: reglementForm.date_echeance || null,
        montant_total: totalMontant,
        recu_par: reglementForm.recu_par || null,
        reference_virement: reglementForm.reference_virement || null,
        observation: reglementForm.observation || null,
        scan_url: scanUrl || null,
        scan_path: scanPath || null,
        facture_ids: factIds,
        facture_numbers: factNums,
        client: clientName,
        tva_mois: reglementForm.tva_mois || null,
        code_reglement: reglementForm.code_reglement || null,
      });

      if (error) { toast.error(`Erreur: ${error.message}`); return; }

      // Update all selected invoices to payé
      for (const id of factIds) {
        await supabase.from('suivi_facturation').update({
          statut: 'payé',
          date_paiement: reglementForm.date_reglement || new Date().toISOString().split('T')[0],
          mode_paiement: reglementForm.type_reglement,
          reglement_banque_type: reglementForm.banque || reglementForm.type_reglement,
          reglement_numero: reglementForm.numero || reglementForm.reference_virement || '',
        }).eq('id', id);
      }

      toast.success(`${factIds.length} facture(s) réglée(s) !`);
      setShowReglementForm(false);
      setSelectedFacts([]);
      setReglementForm({ type_reglement: 'cheque', date_reglement: new Date().toISOString().split('T')[0], numero: '', banque: '', date_echeance: '', recu_par: '', reference_virement: '', observation: '', scanFile: null });
      fetchFacturation();
    } catch (err: any) {
      toast.error(`Erreur: ${err.message}`);
    } finally {
      setSavingReglement(false);
    }
  };

  const filteredReglements = reglementsList.filter((r: any) => {
    if (reglementFilter.type && r.type_reglement !== reglementFilter.type) return false;
    if ((reglementFilter as any).client && !r.client?.toLowerCase().includes((reglementFilter as any).client.toLowerCase())) return false;
    if ((reglementFilter as any).tva_mois && r.tva_mois !== (reglementFilter as any).tva_mois) return false;
    if (reglementFilter.banque && !r.banque?.toLowerCase().includes(reglementFilter.banque.toLowerCase())) return false;
    if (reglementFilter.dateFrom && (r.date_reglement || '') < reglementFilter.dateFrom) return false;
    if (reglementFilter.dateTo && (r.date_reglement || '') > reglementFilter.dateTo) return false;
    return true;
  });
  const handleDeleteReglement = async (r: any) => {
    if (!confirm('Supprimer ce règlement ? Les factures liées repasseront en impayé.')) return;
    // Revert linked invoices to impayé
    for (const fId of (r.facture_ids || [])) {
      await supabase.from('suivi_facturation').update({ statut: 'impayé', date_paiement: null, mode_paiement: null, reglement_banque_type: null, reglement_numero: null }).eq('id', fId);
    }
    await supabase.from('reglements').delete().eq('id', r.id);
    toast.success("Règlement supprimé, factures repassées en impayé.");
    fetchReglements(); fetchFacturation();
  };

  const handleUpdateReglement = async () => {
    if (!editingReglement) return;
    const { error } = await supabase.from('reglements').update({
      date_reglement: editingReglement.date_reglement || null,
      type_reglement: editingReglement.type_reglement,
      numero: editingReglement.numero || null,
      banque: editingReglement.banque || null,
      date_echeance: editingReglement.date_echeance || null,
      recu_par: editingReglement.recu_par || null,
      reference_virement: editingReglement.reference_virement || null,
      observation: editingReglement.observation || null,
      tva_mois: editingReglement.tva_mois || null,
      code_reglement: editingReglement.code_reglement || null,
    }).eq('id', editingReglement.id);
    if (!error) { toast.success("Règlement modifié."); setShowEditReglementForm(false); setEditingReglement(null); fetchReglements(); }
    else toast.error(`Erreur: ${error.message}`);
  };
  // ── Bank RIP CRUD ──
  const fetchRip = async () => {
    if (!companyId) return; setLoadingRip(true);
    const { data } = await supabase.from('bank_rip').select('*').eq('company_id', companyId).order('created_at', { ascending: false });
    setRipList(data || []); setLoadingRip(false);
  };
  const handleSaveRip = async () => {
    const payload = { company_id: companyId, ...ripForm };
    if (editingRip) {
      const { error } = await supabase.from('bank_rip').update(payload).eq('id', editingRip.id);
      if (!error) { toast.success("RIP modifié."); setEditingRip(null); } else { toast.error(`Erreur: ${error.message}`); return; }
    } else {
      const { error } = await supabase.from('bank_rip').insert(payload);
      if (!error) toast.success("RIP ajouté."); else { toast.error(`Erreur: ${error.message}`); return; }
    }
    setShowRipForm(false); setRipForm(emptyRipForm); fetchRip();
  };
  const handleDeleteRip = async (id: string) => {
    if (!confirm('Supprimer ce RIP ?')) return;
    await supabase.from('bank_rip').delete().eq('id', id);
    toast.success("Supprimé."); fetchRip();
  };

  // ── Bank Base RIP CRUD ──
  const fetchBaseRip = async () => {
    if (!companyId) return; setLoadingBaseRip(true);
    const { data } = await supabase.from('bank_base_rip').select('*').eq('company_id', companyId).order('created_at', { ascending: false });
    setBaseRipList(data || []); setLoadingBaseRip(false);
  };
  const handleSaveBaseRip = async () => {
    const payload = { company_id: companyId, ...baseRipForm };
    if (editingBaseRip) {
      const { error } = await supabase.from('bank_base_rip').update(payload).eq('id', editingBaseRip.id);
      if (!error) { toast.success("Base RIP modifié."); setEditingBaseRip(null); } else { toast.error(`Erreur: ${error.message}`); return; }
    } else {
      const { error } = await supabase.from('bank_base_rip').insert(payload);
      if (!error) toast.success("Base RIP ajouté."); else { toast.error(`Erreur: ${error.message}`); return; }
    }
    setShowBaseRipForm(false); setBaseRipForm(emptyBaseRipForm); fetchBaseRip();
  };
  const handleDeleteBaseRip = async (id: string) => {
    if (!confirm('Supprimer ?')) return;
    await supabase.from('bank_base_rip').delete().eq('id', id);
    toast.success("Supprimé."); fetchBaseRip();
  };

  // ── Bank Virement CRUD ──
  const fetchVirement = async () => {
    if (!companyId) return; setLoadingVirement(true);
    const { data } = await supabase.from('bank_virement').select('*').eq('company_id', companyId).order('created_at', { ascending: false });
    setVirementList(data || []); setLoadingVirement(false);
  };
  const handleSaveVirement = async () => {
    const payload = {
      company_id: companyId, date_virement: virementForm.date_virement || null,
      raison_social: virementForm.raison_social || null, rib: virementForm.rib || null,
      montant: parseFloat(virementForm.montant) || 0, justification: virementForm.justification || null,
      tva_mois: virementForm.tva_mois || null, compte_id: virementForm.compte_id || null,
      banque: virementForm.banque || null, agence: virementForm.agence || null, numero_compte: virementForm.numero_compte || null,
    };
    if (editingVirement) {
      const { error } = await supabase.from('bank_virement').update(payload).eq('id', editingVirement.id);
      if (!error) { toast.success("Virement modifié."); setEditingVirement(null); } else { toast.error(`Erreur: ${error.message}`); return; }
    } else {
      const { error } = await supabase.from('bank_virement').insert(payload);
      if (!error) toast.success("Virement ajouté."); else { toast.error(`Erreur: ${error.message}`); return; }
    }
    setShowVirementForm(false); setVirementForm(emptyVirementForm); fetchVirement();
  };
  const handleDeleteVirement = async (id: string) => {
    if (!confirm('Supprimer ?')) return;
    await supabase.from('bank_virement').delete().eq('id', id);
    toast.success("Supprimé."); fetchVirement();
  };
  const filteredVirements = virementList.filter((v: any) => {
    if (virementFilter.raison && !v.raison_social?.toLowerCase().includes(virementFilter.raison.toLowerCase())) return false;
    if (virementFilter.dateFrom && (v.date_virement || '') < virementFilter.dateFrom) return false;
    if (virementFilter.dateTo && (v.date_virement || '') > virementFilter.dateTo) return false;
    return true;
  });

  const handleGenerateVirementPDF = () => {
    const selected = virementList.filter((v: any) => selectedVirements.includes(v.id));
    if (selected.length === 0) return;
    const toWords = (n: number): string => {
      const a = Math.abs(Math.round(n));
      if (a === 0) return 'zéro dirham';
      const u = ['','un','deux','trois','quatre','cinq','six','sept','huit','neuf','dix','onze','douze','treize','quatorze','quinze','seize','dix-sept','dix-huit','dix-neuf'];
      const d = ['','dix','vingt','trente','quarante','cinquante','soixante','soixante','quatre-vingt','quatre-vingt'];
      const w = (n: number): string => {
        if (n === 0) return '';
        if (n < 20) return u[n];
        if (n < 70) return d[Math.floor(n/10)] + (n%10 === 1 ? ' et un' : n%10 ? '-' + u[n%10] : '');
        if (n < 80) return 'soixante' + (n === 71 ? ' et onze' : '-' + u[n-60]);
        if (n < 100) return 'quatre-vingt' + (n === 80 ? 's' : '-' + u[n-80]);
        if (n < 200) return 'cent' + (n === 100 ? '' : ' ' + w(n-100));
        if (n < 1000) return u[Math.floor(n/100)] + ' cent' + (n%100 === 0 ? 's' : ' ' + w(n%100));
        if (n < 2000) return 'mille' + (n === 1000 ? '' : ' ' + w(n-1000));
        if (n < 1000000) return w(Math.floor(n/1000)) + ' mille' + (n%1000 ? ' ' + w(n%1000) : '');
        if (n < 2000000) return 'un million' + (n%1000000 ? ' ' + w(n%1000000) : '');
        return w(Math.floor(n/1000000)) + ' millions' + (n%1000000 ? ' ' + w(n%1000000) : '');
      };
      return w(a) + ' dirhams';
    };
    const total = selected.reduce((s: number, v: any) => s + (Number(v.montant) || 0), 0);
    const fmt = (n: number) => n.toLocaleString('fr-MA', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    const acct = selected[0];

    const rowsHtml = selected.map((v: any, i: number) => {
      const bg = i % 2 === 1 ? '#F2F2F2' : '#fff';
      return `<tr>
        <td style="padding:8px 10px;border:1px solid #ddd;font-size:11px;background:${bg};font-weight:700">${v.raison_social || '—'}</td>
        <td style="padding:8px 10px;border:1px solid #ddd;font-size:11px;font-family:monospace;background:${bg}">${v.rib || '—'}</td>
        <td style="padding:8px 10px;border:1px solid #ddd;font-size:11px;text-align:right;font-weight:700;background:${bg}">${fmt(parseFloat(v.montant) || 0)}</td>
      </tr>`;
    }).join('');

    const emptyCount = Math.max(0, 10 - selected.length);
    const emptyHtml = Array.from({ length: emptyCount }, (_, i) => {
      const bg = (selected.length + i) % 2 === 1 ? '#F2F2F2' : '#fff';
      return `<tr><td style="padding:8px 10px;border:1px solid #ddd;background:${bg}">&nbsp;</td><td style="padding:8px 10px;border:1px solid #ddd;background:${bg}">&nbsp;</td><td style="padding:8px 10px;border:1px solid #ddd;background:${bg}">&nbsp;</td></tr>`;
    }).join('');

    const html = `<!DOCTYPE html><html lang="fr"><head><meta charset="UTF-8"/>
      <style>*{margin:0;padding:0;box-sizing:border-box}@page{margin:0;size:A4}@media print{body{print-color-adjust:exact;-webkit-print-color-adjust:exact}}</style>
      </head><body>
      <div style="width:210mm;min-height:297mm;padding:35mm 15mm 15mm 15mm;font-family:Arial,sans-serif;position:relative">
        <div style="position:absolute;top:8mm;right:15mm;font-size:8px;color:#999">Page 1 / 1</div>
        <div style="display:flex;justify-content:flex-end;margin-bottom:10px">
          <div style="border:2px solid #1F3864;border-radius:4px;padding:8px 16px;min-width:220px;background:#FFFDE7">
            <div style="font-size:10px;font-weight:700;color:#1F3864">BANQUE : <span style="font-weight:400">${acct.banque || '—'}</span></div>
            <div style="font-size:10px;font-weight:700;color:#1F3864;margin-top:3px">Agence : <span style="font-weight:400">${acct.agence || '—'}</span></div>
          </div>
        </div>
        <div style="font-size:18px;font-weight:900;color:#1F3864;text-align:center;margin-bottom:20px;text-decoration:underline">ORDRE DE VIREMENT</div>
        <div style="font-size:12px;line-height:1.8;color:#333;margin-bottom:10px">Madame, Monsieur,</div>
        <div style="font-size:12px;line-height:1.8;color:#333;margin-bottom:10px">Par le débit de mon compte n° <strong style="color:#1F3864">${acct.numero_compte || '—'}</strong></div>
        <div style="font-size:12px;line-height:1.8;color:#333;margin-bottom:16px">Nous vous prions de bien vouloir effectuer le virement suivant :</div>
        <table style="width:100%;border-collapse:collapse;margin-bottom:0">
          <thead><tr>
            <th style="background:#1F3864;color:#fff;font-size:11px;font-weight:700;text-align:center;padding:8px 10px;border:1px solid #1F3864;width:35%">Raison Sociale</th>
            <th style="background:#1F3864;color:#fff;font-size:11px;font-weight:700;text-align:center;padding:8px 10px;border:1px solid #1F3864;width:40%">RIB</th>
            <th style="background:#1F3864;color:#fff;font-size:11px;font-weight:700;text-align:center;padding:8px 10px;border:1px solid #1F3864;width:25%">Montant</th>
          </tr></thead>
          <tbody>
            ${rowsHtml}${emptyHtml}
            <tr>
              <td colspan="2" style="padding:8px 10px;border:1px solid #1F3864;font-size:12px;font-weight:900;text-align:right;background:#1F3864;color:#fff">TOTAL</td>
              <td style="padding:8px 10px;border:1px solid #1F3864;font-size:13px;font-weight:900;text-align:right;background:#1F3864;color:#fff">${fmt(total)} MAD</td>
            </tr>
          </tbody>
        </table>
        <div style="margin-top:20px;font-size:10px;color:#7F7F7F"><strong style="color:#333">Arrêté le présent ordre de virement à la somme de :</strong> ${toWords(total)}</div>
        <div style="margin-top:30px;text-align:right;font-size:11px;color:#333">Casablanca le ${new Date().toLocaleDateString('fr-MA', { day: 'numeric', month: 'long', year: 'numeric' })}</div>
        <div style="margin-top:20px;display:flex;justify-content:flex-end"><div style="width:200px;text-align:center;font-size:10px;color:#555;border-top:1px solid #ccc;padding-top:6px">Signature & Cachet</div></div>
      </div>
      </body></html>`;
    const win = window.open('', '_blank');
    if (win) { win.document.write(html); win.document.close(); win.focus(); setTimeout(() => win.print(), 600); }
  };
  // ── Auto-categorize bank transaction ──
  const categorizeTransaction = (libelle: string): string => {
    const l = (libelle || '').toUpperCase();
    if (/PAIEMENT CHEQUE|RETRAIT ESPECES CHQ|RETRAIT ESPECES SANS/.test(l)) return 'emission_cheque';
    if (/REGLEMENT LCN|REGLEMENT IMPAYE|VIREMENT DE MASSE/.test(l)) return 'emission_effets';
    if (/REMISE CHEQUE/.test(l)) return 'remise_cheque';
    if (/ENCAISSEMENT.*LCN|ESC REM LCN/.test(l)) return 'remise_lc';
    if (/FRAIS|COTISATION|ARRETE|COMMISSION|COM RETRAIT|TIMFISC|TIMBRE|SECCART|MODULE DOCNET/.test(l)) return 'frais_bancaire';
    return 'virement';
  };

  const isCredit = (libelle: string): boolean => {
    const l = (libelle || '').toUpperCase();
    return /VIR.*RECU|VIREMENT RECU|VERS ESP|VERSEMENT ESP|REMISE CHEQUE|ENCAISSEMENT|ESC REM LCN|AIDES AUX TRANSPORT/.test(l);
  };

  // ── PDF Parser ──
  const parseBankPDF = async (file: File) => {
    setParsingPDF(true);
    try {
      // Load pdf.js from CDN if not loaded
      if (!(window as any).pdfjsLib) {
        await new Promise<void>((resolve, reject) => {
          const s = document.createElement('script');
          s.src = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js';
          s.onload = () => resolve();
          s.onerror = () => reject(new Error('Failed to load PDF.js'));
          document.head.appendChild(s);
        });
      }
      const pdfjsLib = (window as any).pdfjsLib;
      pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';

      const buffer = await file.arrayBuffer();
      const pdf = await pdfjsLib.getDocument({ data: buffer }).promise;
      const allItems: { text: string; x: number; y: number; page: number }[] = [];

      for (let p = 1; p <= pdf.numPages; p++) {
        const page = await pdf.getPage(p);
        const content = await page.getTextContent();
        const viewport = page.getViewport({ scale: 1 });
        content.items.forEach((item: any) => {
          if (item.str.trim()) {
            allItems.push({ text: item.str.trim(), x: Math.round(item.transform[4]), y: Math.round(viewport.height - item.transform[5]), page: p });
          }
        });
      }

      // Group items by row (y-position ± 3px tolerance)
      const rows: Map<string, typeof allItems> = new Map();
      allItems.forEach(item => {
        const key = `${item.page}-${Math.round(item.y / 4) * 4}`;
        if (!rows.has(key)) rows.set(key, []);
        rows.get(key)!.push(item);
      });

      // Parse transaction rows
      const transactions: any[] = [];
      let soldeDepart = 0;
      let soldeFinal = 0;

      rows.forEach((items) => {
        items.sort((a, b) => a.x - b.x);
        const fullText = items.map(i => i.text).join(' ');

        // Extract solde
        if (/SOLDE DEPART/i.test(fullText)) {
          const m = fullText.match(/([\d\s]+,\d{2})/);
          if (m) soldeDepart = parseFloat(m[1].replace(/\s/g, '').replace(',', '.'));
        }
        if (/SOLDE FINAL/i.test(fullText)) {
          const m = fullText.match(/([\d\s]+,\d{2})/);
          if (m) soldeFinal = parseFloat(m[1].replace(/\s/g, '').replace(',', '.'));
        }

        // Match transaction line: CODE DD MM LIBELLE DD MM YYYY AMOUNT
        const txMatch = fullText.match(/^(\w{4,8})\s*(\d{2})\s*(\d{2})\s+(.+?)\s+(\d{2})\s+(\d{2})\s+(\d{4})\s+([\d\s]+,\d{2})$/);
        if (txMatch) {
          const code = txMatch[1];
          const dateOp = `${txMatch[2]}/${txMatch[3]}`;
          const libelle = txMatch[4].trim();
          const dateVal = `${txMatch[5]}/${txMatch[6]}/${txMatch[7]}`;
          const amount = parseFloat(txMatch[8].replace(/\s/g, '').replace(',', '.'));
          const credit = isCredit(libelle);

          transactions.push({
            company_id: companyId,
            mois: releveMois,
            code, date_operation: dateOp, libelle, date_valeur: dateVal,
            debit: credit ? 0 : amount,
            credit: credit ? amount : 0,
            category: categorizeTransaction(libelle),
            solde_depart: soldeDepart,
            solde_final: soldeFinal,
          });
        }
      });

      if (transactions.length === 0) {
        toast.error("Aucune transaction trouvée dans le PDF. Vérifiez le format.");
        return;
      }

      // Update solde on all rows
      transactions.forEach(t => { t.solde_depart = soldeDepart; t.solde_final = soldeFinal; });

      const { error } = await supabase.from('bank_releve').insert(transactions);
      if (!error) {
        toast.success(`${transactions.length} transactions extraites et importées !`);
        fetchReleve();
      } else toast.error(`Erreur: ${error.message}`);
    } catch (err: any) {
      toast.error(`Erreur PDF: ${err.message}`);
    } finally {
      setParsingPDF(false);
    }
  };

  // ── Relevé CRUD ──
  const fetchReleve = async () => {
    if (!companyId) return; setLoadingReleve(true);
    let query = supabase.from('bank_releve').select('*').eq('company_id', companyId);
    if (releveFilter.mois) query = query.eq('mois', releveFilter.mois);
    const { data } = await query.order('created_at');
    setReleveList(data || []); setLoadingReleve(false);
  };
  const handleSaveReleve = async () => {
    const payload = {
      company_id: companyId, mois: releveMois, code: releveForm.code || null,
      date_operation: releveForm.date_operation || null, libelle: releveForm.libelle || null,
      date_valeur: releveForm.date_valeur || null, debit: parseFloat(releveForm.debit) || 0,
      credit: parseFloat(releveForm.credit) || 0, category: releveForm.category || 'virement',
      destination: releveForm.destination || null, note_operation: releveForm.note_operation || null,
      ref_reglement: releveForm.ref_reglement || null, observation: releveForm.observation || null,
      code_reglement: releveForm.code_reglement || null,
    };
    if (editingReleve) {
      const { error } = await supabase.from('bank_releve').update(payload).eq('id', editingReleve.id);
      if (!error) { toast.success("Modifié."); setEditingReleve(null); } else { toast.error(`Erreur: ${error.message}`); return; }
    } else {
      const { error } = await supabase.from('bank_releve').insert(payload);
      if (!error) toast.success("Ajouté."); else { toast.error(`Erreur: ${error.message}`); return; }
    }
    setShowReleveForm(false); setReleveForm(emptyReleveForm); fetchReleve();
  };
  const handleDeleteReleve = async (id: string) => {
    if (!confirm('Supprimer ?')) return;
    await supabase.from('bank_releve').delete().eq('id', id);
    toast.success("Supprimé."); fetchReleve();
  };
  const filteredReleve = releveList.filter((r: any) => {
    if (releveFilter.libelle && !r.libelle?.toLowerCase().includes(releveFilter.libelle.toLowerCase())) return false;
    if (releveFilter.category && r.category !== releveFilter.category) return false;
    return true;
  });

  const ETAT_CATEGORIES = [
    { id: 'virement', label: 'Virement / Prélèvement', color: 'blue' },
    { id: 'emission_cheque', label: 'Émission de Chèque', color: 'amber' },
    { id: 'emission_effets', label: 'Émission des Effets', color: 'purple' },
    { id: 'remise_cheque', label: 'Remise de Chèque', color: 'emerald' },
    { id: 'remise_lc', label: 'Remise de LC', color: 'cyan' },
    { id: 'frais_bancaire', label: 'Frais Bancaire', color: 'rose' },
  ];
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
    if (activeTab === 'devis' && companyId) { fetchDevis(); fetchClients(); }
    if (activeTab === 'bon_commande' && companyId) { fetchBC(); fetchFournisseurs(); }
    if (activeTab === 'reglements' && companyId) { fetchReglements(); }
    if (activeTab === 'bank_rip' && companyId) fetchRip();
    if (activeTab === 'bank_base_rip' && companyId) fetchBaseRip();
    if (activeTab === 'bank_virement' && companyId) { fetchVirement(); fetchRip(); fetchBaseRip(); }
    if ((activeTab === 'bank_releve' || activeTab === 'bank_etat_explicatif') && companyId) fetchReleve();
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
  { id: 'staff',         label: 'Staff',               icon: Users },
  { id: 'purchases',     label: 'Achats & Factures',   icon: ShoppingBag },
  { id: 'fleetfix',      label: 'FleetFix',            icon: Wrench },
  { id: 'suivi',         label: 'Suivi Prestation',     icon: FileText },
  { id: 'chauffeurs',    label: 'Chauffeurs',           icon: Truck },
  { id: 'cout_revient',  label: 'Coût de Revient',     icon: TrendingUp },
  { id: 'clients',       label: 'Clients',             icon: Users },
  { id: 'fournisseurs',  label: 'Fournisseurs',        icon: ShoppingBag },
  { id: 'truck_docs',    label: 'Documents Camions',   icon: FolderOpen },
  { id: 'facturation',   label: 'Suivi Facturation',   icon: Receipt },
  { id: 'devis',         label: 'Devis',               icon: FileText },
  { id: 'bon_commande',  label: 'Bon de Commande',     icon: FileText },
  { id: 'reglements',    label: 'Règlements',          icon: Check },
  { id: 'settings',      label: 'Paramètres Facture',  icon: Settings },
] as const;

const bankSubItems: { id: ManagerTab; label: string }[] = [
  { id: 'bank_rip',             label: 'RIP' },
  { id: 'bank_releve',          label: 'Relevé Bancaire' },
  { id: 'bank_base_rip',        label: 'Base RIP FRN / PRT' },
  { id: 'bank_virement',        label: 'Virement' },
  { id: 'bank_etat_explicatif', label: 'État Explicatif' },
  { id: 'bank_tva',             label: 'TVA' },
];

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
                <nav className="flex-1 p-3 space-y-1 overflow-y-auto">
                  {navItems.map(item => {
  if (item.id === 'settings') {
    // Insert Bank group before Settings
    return (
      <div key="bank-group">
        {/* Bank group header */}
        <button onClick={() => setBankOpen(p => !p)}
          className={`w-full flex items-center justify-between gap-3 px-3 py-2.5 rounded-lg text-sm font-bold transition-all cursor-pointer ${bankOpen || activeTab.startsWith('bank_') ? 'bg-blue-50 text-blue-700' : 'text-slate-600 hover:bg-slate-50'}`}>
          <div className="flex items-center gap-3">
            <Landmark size={18} />
            <span>Banque</span>
          </div>
          <span className="text-xs">{bankOpen ? '▼' : '▶'}</span>
        </button>
        {/* Bank sub-items */}
        {(bankOpen || activeTab.startsWith('bank_')) && (
          <div className="ml-6 mt-1 space-y-0.5">
            {bankSubItems.map(sub => (
              <button key={sub.id} onClick={() => { setActiveTab(sub.id); if (sidebarOpen) setSidebarOpen(false); }}
                className={`w-full text-left px-3 py-2 rounded-lg text-xs font-bold transition-all cursor-pointer ${activeTab === sub.id ? 'bg-blue-600 text-white' : 'text-slate-500 hover:bg-slate-50 hover:text-slate-700'}`}>
                {sub.label}
              </button>
            ))}
          </div>
        )}
        {/* Settings item */}
        <button key={item.id} onClick={() => { setActiveTab(item.id); if (sidebarOpen) setSidebarOpen(false); }}
          className={`w-full flex items-center gap-3 px-3 py-2.5 rounded-lg text-sm font-bold transition-all cursor-pointer mt-1 ${activeTab === item.id ? 'bg-blue-600 text-white shadow-lg' : 'text-slate-600 hover:bg-slate-50'}`}>
          <item.icon size={18} />
          <span>{item.label}</span>
        </button>
      </div>
    );
  }
  return (
    <button key={item.id} onClick={() => { setActiveTab(item.id); if (sidebarOpen) setSidebarOpen(false); }}
      className={`w-full flex items-center gap-3 px-3 py-2.5 rounded-lg text-sm font-bold transition-all cursor-pointer ${activeTab === item.id ? 'bg-blue-600 text-white shadow-lg' : 'text-slate-600 hover:bg-slate-50'}`}>
      <item.icon size={18} />
      <span>{item.label}</span>
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
                        <tr>{['Date','Catégorie','Fournisseur','N° Facture','Désignation','IF','ICE','Affectation','HT','TVA','TTC','Banque','N° Réf','Échéance','Mode','Écart Délai','Code Règ.','Actions'].map(h => (
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
                            <td className="px-3 py-2 font-mono text-[10px] text-amber-600 font-bold">{p?.code_reglement || '—'}</td>
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
                  <button onClick={() => setShowReglementForm(true)}
                    className="bg-emerald-600 hover:bg-emerald-700 text-white px-3 py-2 rounded-lg text-xs font-black uppercase tracking-wider flex items-center gap-1.5 cursor-pointer">
                    <Check size={14} /> Réglé
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
        <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">N° Facture</label>
        <input type="text" placeholder="Rechercher..."
          value={(factFilter as any).numero || ''}
          onChange={e => setFactFilter(p => ({ ...p, numero: e.target.value }))}
          className="block mt-1 h-8 rounded-lg border-2 border-slate-200 px-3 text-xs focus:outline-none focus:border-blue-500 w-36" />
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
       <button onClick={() => setFactFilter({ client: '', numero: '', dateFrom: '', dateTo: '', statut: '' } as any)}
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
    <div className="mb-4 bg-white rounded-xl border border-slate-200 p-4 flex flex-wrap gap-6 items-center">
      {(() => {
        const source = selectedFacts.length > 0
          ? facturationList.filter((f: any) => selectedFacts.includes(f.id))
          : filteredFacts;
        const tHT = source.reduce((s: number, f: any) => s + (parseFloat(f.montant_ht) || 0), 0);
        const tTVA = source.reduce((s: number, f: any) => s + (parseFloat(f.tva) || 0), 0);
        const tTTC = source.reduce((s: number, f: any) => s + (parseFloat(f.montant_ttc) || 0), 0);
        const label = selectedFacts.length > 0 ? `${selectedFacts.length} sélectionnée(s)` : `${filteredFacts.length} facture(s)`;
        const fmt = (n: number) => n.toLocaleString('fr-MA', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
        return <>
          <div className="text-[10px] font-black text-slate-400 uppercase tracking-widest">{label}</div>
          <div className="flex gap-5">
            <div><span className="text-[9px] font-black text-slate-400 uppercase tracking-widest block">Total HT</span><span className="text-sm font-bold text-slate-800">{fmt(tHT)} MAD</span></div>
            <div><span className="text-[9px] font-black text-slate-400 uppercase tracking-widest block">TVA</span><span className="text-sm font-bold text-amber-700">{fmt(tTVA)} MAD</span></div>
            <div><span className="text-[9px] font-black text-slate-400 uppercase tracking-widest block">Total TTC</span><span className="text-sm font-black text-blue-700">{fmt(tTTC)} MAD</span></div>
          </div>
        </>;
      })()}
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
{activeTab === 'devis' && (
          <div>
            <div className="mb-6 bg-slate-900 text-white rounded-xl p-6 border border-slate-800">
              <div className="flex items-center justify-between flex-wrap gap-3">
                <div>
                  <span className="inline-flex items-center gap-1.5 px-2 py-0.5 rounded text-[10px] font-extrabold uppercase tracking-widest bg-amber-500 text-white mb-2">
                    <FileText className="w-3.5 h-3.5" /> Devis
                  </span>
                  <h1 className="text-2xl font-extrabold tracking-tight">Gestion des Devis</h1>
                  <p className="text-sm text-slate-400 mt-1">{filteredDevis.length} devis — {selectedDevis.length} sélectionnés</p>
                </div>
                <div className="flex gap-2 flex-wrap">
                  <button onClick={fetchDevis} className="bg-white/10 hover:bg-white/15 px-3 py-2 rounded-lg text-xs font-bold uppercase tracking-wider flex items-center gap-1.5 cursor-pointer"><RefreshCw size={14} /> Actualiser</button>
                  <label className="bg-amber-600 hover:bg-amber-700 text-white px-3 py-2 rounded-lg text-xs font-black uppercase tracking-wider flex items-center gap-1.5 cursor-pointer">
                    <Upload size={14} /> Importer XLS
                    <input type="file" accept=".xlsx,.xls" className="hidden" onChange={async (e) => {
                      const file = e.target.files?.[0];
                      if (!file) return;
                      try {
                        const buffer = await file.arrayBuffer();
                        const wb = XLSX.read(buffer, { type: 'array' });
                        const ws = wb.Sheets[wb.SheetNames[0]];
                        const rawRows: any[] = XLSX.utils.sheet_to_json(ws, { header: 1 });
                        const dataRows = rawRows.slice(1).filter((r: any[]) => r.length > 0 && (r[0] || r[1] || r[2]));
                        const records = dataRows.map((r: any[]) => {
                          const qty = parseFloat(r[8]) || 1;
                          const prixHT = parseFloat(r[9]) || 0;
                          const tvaRate = parseFloat(r[10]) || 0;
                          const montantHT = qty * prixHT;
                          const tvaAmt = parseFloat((montantHT * tvaRate / 100).toFixed(2));
                          return {
                            company_id: companyId,
                            date: r[0] ? String(r[0]) : null,
                            numero_devis: String(r[1] || '').trim() || null,
                            client: String(r[2] || '').trim() || null,
                            personne_contact: String(r[3] || '').trim() || null,
                            type_vehicule: String(r[4] || '').trim() || null,
                            designation: String(r[5] || '').trim() || null,
                            depart: String(r[6] || '').trim() || null,
                            arrivee: String(r[7] || '').trim() || null,
                            quantite: qty,
                            prix_unitaire_ht: prixHT,
                            tva_rate: tvaRate,
                            montant_ht: montantHT,
                            tva_amount: tvaAmt,
                            montant_ttc: montantHT + tvaAmt,
                            observation: String(r[11] || '').trim() || null,
                            statut: String(r[12] || 'en_attente').trim(),
                          };
                        });
                        if (records.length === 0) { toast.error("Aucune donnée trouvée."); return; }
                        const { error } = await supabase.from('devis').insert(records);
                        if (!error) { toast.success(`${records.length} devis importés.`); fetchDevis(); }
                        else toast.error(`Erreur: ${error.message}`);
                      } catch (err: any) { toast.error(`Erreur: ${err.message}`); }
                      e.target.value = '';
                    }} />
                  </label>
                  <button onClick={() => { if (!filteredDevis.length) return; exportToXLS(filteredDevis.map((d:any) => ({ 'Date':d.date,'N° Devis':d.numero_devis,'Client':d.client,'Contact':d.personne_contact,'Type':d.type_vehicule,'Désignation':d.designation,'Départ':d.depart,'Arrivée':d.arrivee,'Qté':d.quantite,'Prix Unit. HT':d.prix_unitaire_ht,'TVA %':d.tva_rate,'Montant HT':d.montant_ht,'TVA':d.tva_amount,'TTC':d.montant_ttc,'Observation':d.observation,'Statut':d.statut })),'devis'); }}
                    className="bg-emerald-600 hover:bg-emerald-700 text-white px-3 py-2 rounded-lg text-xs font-black uppercase tracking-wider flex items-center gap-1.5 cursor-pointer"><Download size={14} /> Export XLS</button>
                  {selectedDevis.length > 0 && (
                    <button onClick={handleGenerateDevisPDF}
                      className="bg-violet-600 hover:bg-violet-700 text-white px-3 py-2 rounded-lg text-xs font-black uppercase tracking-wider flex items-center gap-1.5 cursor-pointer"><FileText size={14} /> Générer PDF ({selectedDevis.length})</button>
                  )}
                  <button onClick={() => { setDevisForm(emptyDevisForm); setEditingDevis(null); setShowDevisForm(true); }}
                    className="bg-blue-600 hover:bg-blue-700 text-white px-3 py-2 rounded-lg text-xs font-black uppercase tracking-wider flex items-center gap-1.5 cursor-pointer"><Plus size={14} /> Nouveau Devis</button>
                </div>
              </div>
            </div>

            {/* Filters */}
            <div className="mb-4 bg-white rounded-xl border border-slate-200 p-4 flex flex-wrap gap-3 items-end">
              <div><label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Client</label>
                <input type="text" placeholder="Filtrer par client..."
                  value={factFilter.client}
                  onChange={e => setFactFilter(p => ({ ...p, client: e.target.value }))}
                  className="block mt-1 h-8 rounded-lg border-2 border-slate-200 px-3 text-xs focus:outline-none focus:border-blue-500 w-48" />
              </div>
              <div><label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Date de</label>
                <input type="date" value={devisFilter.dateFrom} onChange={e => setDevisFilter(p => ({...p, dateFrom: e.target.value}))} className="block mt-1 h-8 rounded-lg border-2 border-slate-200 px-3 text-xs focus:outline-none focus:border-blue-500" /></div>
              <div><label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Date à</label>
                <input type="date" value={devisFilter.dateTo} onChange={e => setDevisFilter(p => ({...p, dateTo: e.target.value}))} className="block mt-1 h-8 rounded-lg border-2 border-slate-200 px-3 text-xs focus:outline-none focus:border-blue-500" /></div>
              <div><label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Statut</label>
                <select value={devisFilter.statut} onChange={e => setDevisFilter(p => ({...p, statut: e.target.value}))} className="block mt-1 h-8 rounded-lg border-2 border-slate-200 px-3 text-xs focus:outline-none focus:border-blue-500">
                  <option value="">Tous</option><option value="en_attente">En attente</option><option value="accepté">Accepté</option><option value="refusé">Refusé</option>
                </select></div>
              <button onClick={() => setDevisFilter({ client: '', dateFrom: '', dateTo: '', statut: '' })} className="h-8 px-3 bg-slate-100 hover:bg-slate-200 text-slate-600 text-xs font-bold rounded-lg cursor-pointer">Réinitialiser</button>
            </div>
            <div className="mb-4 bg-white rounded-xl border border-slate-200 p-4 flex flex-wrap gap-6 items-center">
              {(() => {
                const source = selectedFacts.length > 0
                  ? facturationList.filter((f: any) => selectedFacts.includes(f.id))
                  : filteredFacts;
                const tHT = source.reduce((s: number, f: any) => s + (parseFloat(f.montant_ht) || 0), 0);
                const tTVA = source.reduce((s: number, f: any) => s + (parseFloat(f.tva) || 0), 0);
                const tTTC = source.reduce((s: number, f: any) => s + (parseFloat(f.montant_ttc) || 0), 0);
                const label = selectedFacts.length > 0 ? `${selectedFacts.length} sélectionnée(s)` : `${filteredFacts.length} facture(s)`;
                const fmt = (n: number) => n.toLocaleString('fr-MA', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
                return <>
                  <div className="text-[10px] font-black text-slate-400 uppercase tracking-widest">{label}</div>
                  <div className="flex gap-5">
                    <div><span className="text-[9px] font-black text-slate-400 uppercase tracking-widest block">Total HT</span><span className="text-sm font-bold text-slate-800">{fmt(tHT)} MAD</span></div>
                    <div><span className="text-[9px] font-black text-slate-400 uppercase tracking-widest block">TVA</span><span className="text-sm font-bold text-amber-700">{fmt(tTVA)} MAD</span></div>
                    <div><span className="text-[9px] font-black text-slate-400 uppercase tracking-widest block">Total TTC</span><span className="text-sm font-black text-blue-700">{fmt(tTTC)} MAD</span></div>
                  </div>
                </>;
              })()}
            </div>
            <div className="mb-4 p-3 bg-blue-50 border border-blue-100 rounded-xl text-xs text-blue-700 font-medium">
              📋 Format import XLS — colonnes dans l'ordre :
              <span className="font-black ml-1">Date | N° Devis | Client | Contact | Type | Désignation | Départ | Arrivée | Qté | Prix Unit. HT | TVA % | Observation | Statut</span>
            </div>

            {/* Table */}
            {loadingDevis ? <div className="flex justify-center py-20"><Loader2 className="w-6 h-6 text-blue-600 animate-spin" /></div> : (
              <div className="bg-white rounded-xl border border-slate-200 overflow-hidden">
                <div className="overflow-x-auto">
                  <table className="w-full text-left min-w-[1200px]">
                    <thead className="bg-slate-50 border-b border-slate-200">
                      <tr>
                        <th className="px-3 py-3 w-8"><input type="checkbox" checked={selectedDevis.length === filteredDevis.length && filteredDevis.length > 0} onChange={e => setSelectedDevis(e.target.checked ? filteredDevis.map((d:any)=>d.id) : [])} className="accent-blue-600" /></th>
                        {['Date','N° Devis','Client','Contact','Type','Désignation','Qté','Prix Unit. HT','TVA %','Montant HT','TTC','Statut','Actions'].map(h => (
                          <th key={h} className="px-3 py-2 text-[9px] font-black text-slate-400 uppercase tracking-widest whitespace-nowrap">{h}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-100">
                      {filteredDevis.length === 0 ? (
                        <tr><td colSpan={14} className="px-4 py-10 text-center text-sm text-slate-400">Aucun devis.</td></tr>
                      ) : filteredDevis.map((d: any) => (
                        <tr key={d.id} className={`hover:bg-slate-50 transition-colors ${selectedDevis.includes(d.id) ? 'bg-blue-50/50' : ''}`}>
                          <td className="px-3 py-3"><input type="checkbox" checked={selectedDevis.includes(d.id)} onChange={e => setSelectedDevis(prev => e.target.checked ? [...prev, d.id] : prev.filter(x=>x!==d.id))} className="accent-blue-600" /></td>
                          <td className="px-3 py-2 text-xs text-slate-700">{d.date || '—'}</td>
                          <td className="px-3 py-2 font-mono text-xs text-blue-600">{d.numero_devis || '—'}</td>
                          <td className="px-3 py-2 text-xs font-semibold text-slate-700">{d.client || '—'}</td>
                          <td className="px-3 py-2 text-xs text-slate-600">{d.personne_contact || '—'}</td>
                          <td className="px-3 py-2 text-xs text-slate-600">{d.type_vehicule || '—'}</td>
                          <td className="px-3 py-2 text-xs text-slate-600 max-w-[150px] truncate">{d.designation || `${d.depart||''} → ${d.arrivee||''}`}</td>
                          <td className="px-3 py-2 font-mono text-xs text-slate-600">{d.quantite}</td>
                          <td className="px-3 py-2 font-mono text-xs text-slate-700">{Number(d.prix_unitaire_ht||0).toLocaleString('fr-MA',{minimumFractionDigits:2})}</td>
                          <td className="px-3 py-2 font-mono text-xs text-slate-500">{d.tva_rate}%</td>
                          <td className="px-3 py-2 font-mono text-xs text-slate-700">{Number(d.montant_ht||0).toLocaleString('fr-MA',{minimumFractionDigits:2})}</td>
                          <td className="px-3 py-2 font-mono text-xs font-bold text-slate-900">{Number(d.montant_ttc||0).toLocaleString('fr-MA',{minimumFractionDigits:2})}</td>
                          <td className="px-3 py-2"><span className={`text-[9px] font-black px-1.5 py-0.5 rounded ${d.statut==='accepté'?'bg-emerald-50 text-emerald-700':d.statut==='refusé'?'bg-rose-50 text-rose-700':'bg-amber-50 text-amber-700'}`}>{d.statut}</span></td>
                          <td className="px-3 py-2 flex gap-1">
                            <button onClick={() => { setEditingDevis(d); setDevisForm({...d, quantite:String(d.quantite), prix_unitaire_ht:String(d.prix_unitaire_ht), tva_rate:String(d.tva_rate)}); setShowDevisForm(true); }} className="text-[10px] font-bold text-blue-600 hover:text-blue-700 uppercase cursor-pointer"><Pencil size={13} /></button>
                            <button onClick={() => handleDeleteDevis(d.id)} className="text-[10px] font-bold text-slate-400 hover:text-rose-600 cursor-pointer"><Trash2 size={13} /></button>
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
        {activeTab === 'bon_commande' && (
          <div>
            <div className="mb-6 bg-slate-900 text-white rounded-xl p-6 border border-slate-800">
              <div className="flex items-center justify-between flex-wrap gap-3">
                <div>
                  <span className="inline-flex items-center gap-1.5 px-2 py-0.5 rounded text-[10px] font-extrabold uppercase tracking-widest bg-cyan-500 text-white mb-2">
                    <FileText className="w-3.5 h-3.5" /> Bon de Commande
                  </span>
                  <h1 className="text-2xl font-extrabold tracking-tight">Gestion des Bons de Commande</h1>
                  <p className="text-sm text-slate-400 mt-1">{filteredBC.length} BC — {selectedBC.length} sélectionnés</p>
                </div>
                <div className="flex gap-2 flex-wrap">
                  <button onClick={fetchBC} className="bg-white/10 hover:bg-white/15 px-3 py-2 rounded-lg text-xs font-bold uppercase tracking-wider flex items-center gap-1.5 cursor-pointer"><RefreshCw size={14} /> Actualiser</button>
                  <label className="bg-amber-600 hover:bg-amber-700 text-white px-3 py-2 rounded-lg text-xs font-black uppercase tracking-wider flex items-center gap-1.5 cursor-pointer">
                    <Upload size={14} /> Importer XLS
                    <input type="file" accept=".xlsx,.xls" className="hidden" onChange={async (e) => {
                      const file = e.target.files?.[0]; if (!file) return;
                      try {
                        const buffer = await file.arrayBuffer();
                        const wb = XLSX.read(buffer, { type: 'array' });
                        const ws = wb.Sheets[wb.SheetNames[0]];
                        const rawRows: any[] = XLSX.utils.sheet_to_json(ws, { header: 1 });
                        const dataRows = rawRows.slice(1).filter((r: any[]) => r.length > 0 && (r[0]||r[1]||r[2]));
                        const records = dataRows.map((r: any[]) => {
                          const qty = parseFloat(r[6]) || 1; const pHT = parseFloat(r[7]) || 0; const rate = parseFloat(r[8]) || 0;
                          const mHT = qty * pHT; const tva = parseFloat((mHT * rate / 100).toFixed(2));
                          return { company_id: companyId, date: r[0]?String(r[0]):null, numero_bc: String(r[1]||'').trim()||null,
                            fournisseur: String(r[2]||'').trim()||null, personne_contact: String(r[3]||'').trim()||null,
                            reference: String(r[4]||'').trim()||null, designation: String(r[5]||'').trim()||null,
                            quantite: qty, prix_unitaire_ht: pHT, tva_rate: rate, montant_ht: mHT, tva_amount: tva, montant_ttc: mHT+tva,
                            observation: String(r[9]||'').trim()||null, statut: String(r[10]||'en_attente').trim() };
                        });
                        if (!records.length) { toast.error("Aucune donnée."); return; }
                        const { error } = await supabase.from('bon_commande').insert(records);
                        if (!error) { toast.success(`${records.length} BC importés.`); fetchBC(); } else toast.error(`Erreur: ${error.message}`);
                      } catch (err: any) { toast.error(`Erreur: ${err.message}`); }
                      e.target.value = '';
                    }} />
                  </label>
                  <button onClick={() => { if (!filteredBC.length) return; exportToXLS(filteredBC.map((d:any) => ({ 'Date':d.date,'N° BC':d.numero_bc,'Fournisseur':d.fournisseur,'Contact':d.personne_contact,'Référence':d.reference,'Désignation':d.designation,'Qté':d.quantite,'Prix Unit. HT':d.prix_unitaire_ht,'TVA %':d.tva_rate,'Montant HT':d.montant_ht,'TVA':d.tva_amount,'TTC':d.montant_ttc,'Observation':d.observation,'Statut':d.statut })),'bon_commande'); }}
                    className="bg-emerald-600 hover:bg-emerald-700 text-white px-3 py-2 rounded-lg text-xs font-black uppercase tracking-wider flex items-center gap-1.5 cursor-pointer"><Download size={14} /> Export XLS</button>
                  {selectedBC.length > 0 && (
                    <button onClick={handleGenerateBCPDF} className="bg-violet-600 hover:bg-violet-700 text-white px-3 py-2 rounded-lg text-xs font-black uppercase tracking-wider flex items-center gap-1.5 cursor-pointer"><FileText size={14} /> Générer PDF ({selectedBC.length})</button>
                  )}
                  <button onClick={() => { setBcForm(emptyBCForm); setEditingBC(null); setShowBCForm(true); }}
                    className="bg-blue-600 hover:bg-blue-700 text-white px-3 py-2 rounded-lg text-xs font-black uppercase tracking-wider flex items-center gap-1.5 cursor-pointer"><Plus size={14} /> Nouveau BC</button>
                </div>
              </div>
            </div>
            {/* Filters */}
            <div className="mb-4 bg-white rounded-xl border border-slate-200 p-4 flex flex-wrap gap-3 items-end">
              <div><label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Fournisseur</label>
                <input type="text" placeholder="Filtrer..." value={bcFilter.fournisseur} onChange={e => setBcFilter(p => ({...p, fournisseur: e.target.value}))} className="block mt-1 h-8 rounded-lg border-2 border-slate-200 px-3 text-xs focus:outline-none focus:border-blue-500 w-44" /></div>
              <div><label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Date de</label>
                <input type="date" value={bcFilter.dateFrom} onChange={e => setBcFilter(p => ({...p, dateFrom: e.target.value}))} className="block mt-1 h-8 rounded-lg border-2 border-slate-200 px-3 text-xs focus:outline-none focus:border-blue-500" /></div>
              <div><label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Date à</label>
                <input type="date" value={bcFilter.dateTo} onChange={e => setBcFilter(p => ({...p, dateTo: e.target.value}))} className="block mt-1 h-8 rounded-lg border-2 border-slate-200 px-3 text-xs focus:outline-none focus:border-blue-500" /></div>
              <div><label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Statut</label>
                <select value={bcFilter.statut} onChange={e => setBcFilter(p => ({...p, statut: e.target.value}))} className="block mt-1 h-8 rounded-lg border-2 border-slate-200 px-3 text-xs focus:outline-none focus:border-blue-500">
                  <option value="">Tous</option><option value="en_attente">En attente</option><option value="validé">Validé</option><option value="livré">Livré</option><option value="annulé">Annulé</option>
                </select></div>
              <button onClick={() => setBcFilter({ fournisseur: '', dateFrom: '', dateTo: '', statut: '' })} className="h-8 px-3 bg-slate-100 hover:bg-slate-200 text-slate-600 text-xs font-bold rounded-lg cursor-pointer">Réinitialiser</button>
            </div>
            <div className="mb-4 p-3 bg-blue-50 border border-blue-100 rounded-xl text-xs text-blue-700 font-medium">
              📋 Format import XLS : <span className="font-black ml-1">Date | N° BC | Fournisseur | Contact | Référence | Désignation | Qté | Prix Unit. HT | TVA % | Observation | Statut</span>
            </div>
            {/* Table */}
            {loadingBC ? <div className="flex justify-center py-20"><Loader2 className="w-6 h-6 text-blue-600 animate-spin" /></div> : (
              <div className="bg-white rounded-xl border border-slate-200 overflow-hidden">
                <div className="overflow-x-auto">
                  <table className="w-full text-left min-w-[1100px]">
                    <thead className="bg-slate-50 border-b border-slate-200">
                      <tr>
                        <th className="px-3 py-3 w-8"><input type="checkbox" checked={selectedBC.length === filteredBC.length && filteredBC.length > 0} onChange={e => setSelectedBC(e.target.checked ? filteredBC.map((d:any)=>d.id) : [])} className="accent-blue-600" /></th>
                        {['Date','N° BC','Fournisseur','Contact','Réf.','Désignation','Qté','Prix HT','TVA %','Montant HT','TTC','Statut','Actions'].map(h => (
                          <th key={h} className="px-3 py-2 text-[9px] font-black text-slate-400 uppercase tracking-widest whitespace-nowrap">{h}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-100">
                      {filteredBC.length === 0 ? (
                        <tr><td colSpan={14} className="px-4 py-10 text-center text-sm text-slate-400">Aucun bon de commande.</td></tr>
                      ) : filteredBC.map((d: any) => (
                        <tr key={d.id} className={`hover:bg-slate-50 transition-colors ${selectedBC.includes(d.id)?'bg-blue-50/50':''}`}>
                          <td className="px-3 py-3"><input type="checkbox" checked={selectedBC.includes(d.id)} onChange={e => setSelectedBC(prev => e.target.checked ? [...prev,d.id] : prev.filter(x=>x!==d.id))} className="accent-blue-600" /></td>
                          <td className="px-3 py-2 text-xs text-slate-700">{d.date||'—'}</td>
                          <td className="px-3 py-2 font-mono text-xs text-blue-600">{d.numero_bc||'—'}</td>
                          <td className="px-3 py-2 text-xs font-semibold text-slate-700">{d.fournisseur||'—'}</td>
                          <td className="px-3 py-2 text-xs text-slate-600">{d.personne_contact||'—'}</td>
                          <td className="px-3 py-2 text-xs text-slate-600">{d.reference||'—'}</td>
                          <td className="px-3 py-2 text-xs text-slate-600 max-w-[120px] truncate">{d.designation||'—'}</td>
                          <td className="px-3 py-2 font-mono text-xs text-slate-600">{d.quantite}</td>
                          <td className="px-3 py-2 font-mono text-xs text-slate-700">{Number(d.prix_unitaire_ht||0).toLocaleString('fr-MA',{minimumFractionDigits:2})}</td>
                          <td className="px-3 py-2 font-mono text-xs text-slate-500">{d.tva_rate}%</td>
                          <td className="px-3 py-2 font-mono text-xs text-slate-700">{Number(d.montant_ht||0).toLocaleString('fr-MA',{minimumFractionDigits:2})}</td>
                          <td className="px-3 py-2 font-mono text-xs font-bold text-slate-900">{Number(d.montant_ttc||0).toLocaleString('fr-MA',{minimumFractionDigits:2})}</td>
                          <td className="px-3 py-2"><span className={`text-[9px] font-black px-1.5 py-0.5 rounded ${d.statut==='validé'?'bg-emerald-50 text-emerald-700':d.statut==='livré'?'bg-blue-50 text-blue-700':d.statut==='annulé'?'bg-rose-50 text-rose-700':'bg-amber-50 text-amber-700'}`}>{d.statut}</span></td>
                          <td className="px-3 py-2 flex gap-1">
                            <button onClick={() => { setEditingBC(d); setBcForm({...d,quantite:String(d.quantite),prix_unitaire_ht:String(d.prix_unitaire_ht),tva_rate:String(d.tva_rate)}); setShowBCForm(true); }} className="text-slate-400 hover:text-blue-600 cursor-pointer"><Pencil size={13} /></button>
                            <button onClick={() => handleDeleteBC(d.id)} className="text-slate-400 hover:text-rose-600 cursor-pointer"><Trash2 size={13} /></button>
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
        {activeTab === 'reglements' && (
          <div>
            <div className="mb-6 bg-slate-900 text-white rounded-xl p-6 border border-slate-800">
              <div className="flex items-center justify-between flex-wrap gap-3">
                <div>
                  <span className="inline-flex items-center gap-1.5 px-2 py-0.5 rounded text-[10px] font-extrabold uppercase tracking-widest bg-emerald-500 text-white mb-2">
                    <Check className="w-3.5 h-3.5" /> Règlements
                  </span>
                  <h1 className="text-2xl font-extrabold tracking-tight">Suivi des Règlements</h1>
                  <p className="text-sm text-slate-400 mt-1">{filteredReglements.length} règlements</p>
                </div>
                <div className="flex gap-2 flex-wrap">
                  <button onClick={fetchReglements} className="bg-white/10 hover:bg-white/15 px-3 py-2 rounded-lg text-xs font-bold uppercase tracking-wider flex items-center gap-1.5 cursor-pointer"><RefreshCw size={14} /> Actualiser</button>
                  <button onClick={() => { if (!filteredReglements.length) return; exportToXLS(filteredReglements.map((r:any) => ({
                    'Date': r.date_reglement, 'Type': r.type_reglement, 'Client': r.client, 'TVA Mois': r.tva_mois, 'N°': r.numero,
                    'Banque': r.banque, 'Échéance': r.date_echeance, 'Montant TTC': r.montant_total,
                    'Reçu par': r.recu_par, 'Réf. Virement': r.reference_virement,
                    'Factures': (r.facture_numbers||[]).join(', '), 'Observation': r.observation,
                  })), 'reglements'); }}
                    className="bg-emerald-600 hover:bg-emerald-700 text-white px-3 py-2 rounded-lg text-xs font-black uppercase tracking-wider flex items-center gap-1.5 cursor-pointer"><Download size={14} /> Export XLS</button>
                </div>
              </div>
            </div>
            {/* Filters */}
            <div className="mb-4 bg-white rounded-xl border border-slate-200 p-4 flex flex-wrap gap-3 items-end">
              <div><label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Type</label>
                <select value={reglementFilter.type} onChange={e => setReglementFilter(p => ({...p,type:e.target.value}))} className="block mt-1 h-8 rounded-lg border-2 border-slate-200 px-3 text-xs focus:outline-none focus:border-blue-500">
                  <option value="">Tous</option><option value="cheque">Chèque</option><option value="effet">Effet</option><option value="virement">Virement</option><option value="espece">Espèce</option><option value="compensation">Compensation</option>
                </select></div>
              <div><label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Client</label>
                <input type="text" placeholder="Filtrer..." value={(reglementFilter as any).client || ''} onChange={e => setReglementFilter(p => ({...p,client:e.target.value}))} className="block mt-1 h-8 rounded-lg border-2 border-slate-200 px-3 text-xs focus:outline-none focus:border-blue-500 w-36" /></div>
              <div><label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">TVA Mois</label>
                <input type="month" value={(reglementFilter as any).tva_mois || ''} onChange={e => setReglementFilter(p => ({...p,tva_mois:e.target.value}))} className="block mt-1 h-8 rounded-lg border-2 border-slate-200 px-3 text-xs focus:outline-none focus:border-blue-500" /></div>
              <div><label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Banque</label>
                <input type="text" placeholder="Filtrer..." value={reglementFilter.banque} onChange={e => setReglementFilter(p => ({...p,banque:e.target.value}))} className="block mt-1 h-8 rounded-lg border-2 border-slate-200 px-3 text-xs focus:outline-none focus:border-blue-500 w-36" /></div>
              <div><label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Date de</label>
                <input type="date" value={reglementFilter.dateFrom} onChange={e => setReglementFilter(p => ({...p,dateFrom:e.target.value}))} className="block mt-1 h-8 rounded-lg border-2 border-slate-200 px-3 text-xs focus:outline-none focus:border-blue-500" /></div>
              <div><label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Date à</label>
                <input type="date" value={reglementFilter.dateTo} onChange={e => setReglementFilter(p => ({...p,dateTo:e.target.value}))} className="block mt-1 h-8 rounded-lg border-2 border-slate-200 px-3 text-xs focus:outline-none focus:border-blue-500" /></div>
              <button onClick={() => setReglementFilter({type:'',client:'',tva_mois:'',banque:'',dateFrom:'',dateTo:''} as any)} className="h-8 px-3 bg-slate-100 hover:bg-slate-200 text-slate-600 text-xs font-bold rounded-lg cursor-pointer">Réinitialiser</button>
            </div>
            {/* Table */}
            {loadingReglements ? <div className="flex justify-center py-20"><Loader2 className="w-6 h-6 text-blue-600 animate-spin" /></div> : (
              <div className="space-y-3">
                {filteredReglements.filter(Boolean).length === 0 ? (
                  <div className="bg-white rounded-xl border border-slate-200 p-10 text-center text-sm text-slate-400">Aucun règlement.</div>
                ) : filteredReglements.filter(Boolean).map((r: any) => (
                  <div key={r.id} className="bg-white rounded-xl border border-slate-200 overflow-hidden">
                    {/* Header row — click to expand */}
                    <div onClick={() => setExpandedReglement(prev => prev === r.id ? null : r.id)}
                      className="flex items-center justify-between px-5 py-4 cursor-pointer hover:bg-slate-50 transition-colors">
                      <div className="flex items-center gap-4">
                        <span className={`text-[9px] font-black px-2 py-0.5 rounded uppercase ${r.type_reglement==='cheque'?'bg-blue-50 text-blue-700':r.type_reglement==='effet'?'bg-amber-50 text-amber-700':r.type_reglement==='virement'?'bg-purple-50 text-purple-700':'bg-emerald-50 text-emerald-700'}`}>
                          {r.type_reglement}
                        </span>
                        <div>
                          <span className="text-sm font-bold text-slate-800">{r.numero || r.reference_virement || '—'}</span>
                        {r && r.code_reglement && <span className="text-[9px] font-mono font-bold bg-slate-100 text-slate-600 px-1.5 py-0.5 rounded">{r.code_reglement}</span>}
                          <span className="text-xs text-slate-400 ml-2">{r.banque || r.recu_par || ''}</span>
                        </div>
                        <span className="text-xs text-slate-500">{r.date_reglement}</span>
                      </div>
                      <div className="flex items-center gap-4">
                        <span className="text-sm font-black text-emerald-700">{Number(r.montant_total||0).toLocaleString('fr-MA',{minimumFractionDigits:2})} MAD</span>
                        <span className="text-xs text-slate-400">{(r.facture_numbers||[]).length} facture(s)</span>
                        <span className="text-slate-400">{expandedReglement === r.id ? '▼' : '▶'}</span>
                      </div>
                    </div>
                    {/* Expanded details */}
                    {expandedReglement === r.id && (
                      <div className="px-5 pb-5 border-t border-slate-100">
                        <div className="grid grid-cols-2 sm:grid-cols-4 gap-3 mt-4 mb-4">
                          {r.client && <div><span className="text-[9px] font-black text-slate-400 uppercase block">Client</span><span className="text-xs font-bold text-slate-700">{r.client}</span></div>}
                          {r.tva_mois && <div><span className="text-[9px] font-black text-slate-400 uppercase block">TVA Mois</span><span className="text-xs font-bold text-slate-700">{r.tva_mois}</span></div>}
                          {r.date_echeance && <div><span className="text-[9px] font-black text-slate-400 uppercase block">Échéance</span><span className="text-xs font-bold text-slate-700">{r.date_echeance}</span></div>}
                          {r.banque && <div><span className="text-[9px] font-black text-slate-400 uppercase block">Banque</span><span className="text-xs font-bold text-slate-700">{r.banque}</span></div>}
                          {r.recu_par && <div><span className="text-[9px] font-black text-slate-400 uppercase block">Reçu par</span><span className="text-xs font-bold text-slate-700">{r.recu_par}</span></div>}
                          {r.observation && <div><span className="text-[9px] font-black text-slate-400 uppercase block">Observation</span><span className="text-xs text-slate-700">{r.observation}</span></div>}
                        </div>
                        {/* Scan */}
                        {r.scan_url && (
                          <div className="mb-4">
                            <span className="text-[9px] font-black text-slate-400 uppercase block mb-2">Scan du règlement</span>
                            {r.scan_url.endsWith('.pdf') ? (
                              <a href={r.scan_url} target="_blank" rel="noopener noreferrer" className="text-xs text-blue-600 font-bold hover:underline">📄 Ouvrir le PDF</a>
                            ) : (
                              <img src={r.scan_url} alt="Scan" className="max-w-sm rounded-lg border border-slate-200 shadow-sm" />
                            )}
                          </div>
                        )}
                        {/* Actions */}
                        <div className="mb-4 flex gap-2">
                          <button onClick={(e) => { e.stopPropagation(); setEditingReglement({...r}); setShowEditReglementForm(true); }}
                            className="px-3 py-1.5 bg-blue-50 hover:bg-blue-100 text-blue-700 text-[10px] font-black uppercase rounded-lg cursor-pointer flex items-center gap-1">
                            <Pencil size={12} /> Modifier
                          </button>
                          <button onClick={(e) => { e.stopPropagation(); handleDeleteReglement(r); }}
                            className="px-3 py-1.5 bg-rose-50 hover:bg-rose-100 text-rose-700 text-[10px] font-black uppercase rounded-lg cursor-pointer flex items-center gap-1">
                            <Trash2 size={12} /> Supprimer
                          </button>
                        </div>
                        {/* Factures linked */}
                        <div>
                          <span className="text-[9px] font-black text-slate-400 uppercase block mb-2">Factures réglées</span>
                          <div className="flex flex-wrap gap-1.5">
                            {(r.facture_numbers||[]).map((num: string, i: number) => (
                              <span key={i} className="text-[10px] bg-emerald-50 border border-emerald-200 text-emerald-800 font-bold px-2 py-0.5 rounded">
                                {num}
                              </span>
                            ))}
                          </div>
                        </div>
                      </div>
                    )}
                  </div>
                ))}
              </div>
            )}
          </div>
        )}
        {activeTab === 'bank_rip' && (
          <div>
            <div className="mb-6 bg-slate-900 text-white rounded-xl p-6 border border-slate-800">
              <div className="flex items-center justify-between flex-wrap gap-3">
                <div>
                  <span className="inline-flex items-center gap-1.5 px-2 py-0.5 rounded text-[10px] font-extrabold uppercase tracking-widest bg-blue-500 text-white mb-2"><Landmark className="w-3.5 h-3.5" /> RIP</span>
                  <h1 className="text-2xl font-extrabold tracking-tight">Relevé d'Identité Bancaire</h1>
                  <p className="text-sm text-slate-400 mt-1">{ripList.length} comptes</p>
                </div>
                <div className="flex gap-2 flex-wrap">
                  <button onClick={fetchRip} className="bg-white/10 hover:bg-white/15 px-3 py-2 rounded-lg text-xs font-bold uppercase tracking-wider flex items-center gap-1.5 cursor-pointer"><RefreshCw size={14} /> Actualiser</button>
                  <label className="bg-amber-600 hover:bg-amber-700 text-white px-3 py-2 rounded-lg text-xs font-black uppercase tracking-wider flex items-center gap-1.5 cursor-pointer">
                    <Upload size={14} /> Importer XLS
                    <input type="file" accept=".xlsx,.xls" className="hidden" onChange={async (e) => {
                      const file = e.target.files?.[0]; if (!file) return;
                      try {
                        const buffer = await file.arrayBuffer(); const wb = XLSX.read(buffer, { type: 'array' });
                        const ws = wb.Sheets[wb.SheetNames[0]];
                        const rawRows: any[] = XLSX.utils.sheet_to_json(ws, { header: 1 });
                        const headerIdx = rawRows.findIndex((r: any[]) => r.some((c: any) => String(c).toLowerCase().includes('banque')));
                        const dataRows = rawRows.slice(headerIdx >= 0 ? headerIdx + 1 : 1).filter((r: any[]) => r.length > 0 && (r[0]||r[1]));
                        const records = dataRows.map((r: any[]) => ({
                          company_id: companyId, banque: String(r[0]||'').trim()||null, agence: String(r[1]||'').trim()||null,
                          code_banque: String(r[2]||'').trim()||null, code_ville: String(r[3]||'').trim()||null,
                          numero_compte: String(r[4]||'').trim()||null, cle_rip: String(r[5]||'').trim()||null, code_swift: String(r[6]||'').trim()||null,
                        }));
                        if (!records.length) { toast.error("Aucune donnée."); return; }
                        const { error } = await supabase.from('bank_rip').insert(records);
                        if (!error) { toast.success(`${records.length} RIP importés.`); fetchRip(); } else toast.error(`Erreur: ${error.message}`);
                      } catch (err: any) { toast.error(`Erreur: ${err.message}`); }
                      e.target.value = '';
                    }} />
                  </label>
                  <button onClick={() => { if (!ripList.length) return; exportToXLS(ripList.map((r:any) => ({ 'Banque':r.banque,'Agence':r.agence,'Code Banque':r.code_banque,'Code Ville':r.code_ville,'N° Compte':r.numero_compte,'Clé RIP':r.cle_rip,'Code SWIFT':r.code_swift })),'rip'); }}
                    className="bg-emerald-600 hover:bg-emerald-700 text-white px-3 py-2 rounded-lg text-xs font-black uppercase tracking-wider flex items-center gap-1.5 cursor-pointer"><Download size={14} /> Export XLS</button>
                  <button onClick={() => { setRipForm(emptyRipForm); setEditingRip(null); setShowRipForm(true); }}
                    className="bg-blue-600 hover:bg-blue-700 text-white px-3 py-2 rounded-lg text-xs font-black uppercase tracking-wider flex items-center gap-1.5 cursor-pointer"><Plus size={14} /> Nouveau</button>
                </div>
              </div>
            </div>
            {loadingRip ? <div className="flex justify-center py-20"><Loader2 className="w-6 h-6 text-blue-600 animate-spin" /></div> : (
              <div className="bg-white rounded-xl border border-slate-200 overflow-hidden">
                <div className="overflow-x-auto">
                  <table className="w-full text-left">
                    <thead className="bg-slate-50 border-b border-slate-200">
                      <tr>
                        {['Banque','Agence','Code Banque','Code Ville','N° Compte','Clé RIP','Code SWIFT','Actions'].map(h => (
                          <th key={h} className="px-4 py-2 text-[9px] font-black text-slate-400 uppercase tracking-widest">{h}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-100">
                      {ripList.length === 0 ? (
                        <tr><td colSpan={8} className="px-4 py-10 text-center text-sm text-slate-400">Aucun RIP.</td></tr>
                      ) : ripList.map((r: any) => (
                        <tr key={r.id} className="hover:bg-slate-50 transition-colors">
                          <td className="px-4 py-3 text-xs font-semibold text-slate-700">{r.banque||'—'}</td>
                          <td className="px-4 py-3 text-xs text-slate-600">{r.agence||'—'}</td>
                          <td className="px-4 py-3 font-mono text-xs text-slate-600">{r.code_banque||'—'}</td>
                          <td className="px-4 py-3 font-mono text-xs text-slate-600">{r.code_ville||'—'}</td>
                          <td className="px-4 py-3 font-mono text-xs text-blue-600 font-bold">{r.numero_compte||'—'}</td>
                          <td className="px-4 py-3 font-mono text-xs text-slate-600">{r.cle_rip||'—'}</td>
                          <td className="px-4 py-3 font-mono text-xs text-slate-600">{r.code_swift||'—'}</td>
                          <td className="px-4 py-3 flex gap-1">
                            <button onClick={() => { setEditingRip(r); setRipForm({...r}); setShowRipForm(true); }} className="text-slate-400 hover:text-blue-600 cursor-pointer"><Pencil size={13} /></button>
                            <button onClick={() => handleDeleteRip(r.id)} className="text-slate-400 hover:text-rose-600 cursor-pointer"><Trash2 size={13} /></button>
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
        {activeTab === 'bank_base_rip' && (
          <div>
            <div className="mb-6 bg-slate-900 text-white rounded-xl p-6 border border-slate-800">
              <div className="flex items-center justify-between flex-wrap gap-3">
                <div>
                  <span className="inline-flex items-center gap-1.5 px-2 py-0.5 rounded text-[10px] font-extrabold uppercase tracking-widest bg-purple-500 text-white mb-2"><Landmark className="w-3.5 h-3.5" /> Base RIP</span>
                  <h1 className="text-2xl font-extrabold tracking-tight">Base RIP Fournisseurs / Particuliers</h1>
                  <p className="text-sm text-slate-400 mt-1">{baseRipList.length} enregistrements</p>
                </div>
                <div className="flex gap-2 flex-wrap">
                  <button onClick={fetchBaseRip} className="bg-white/10 hover:bg-white/15 px-3 py-2 rounded-lg text-xs font-bold uppercase tracking-wider flex items-center gap-1.5 cursor-pointer"><RefreshCw size={14} /> Actualiser</button>
                  <label className="bg-amber-600 hover:bg-amber-700 text-white px-3 py-2 rounded-lg text-xs font-black uppercase tracking-wider flex items-center gap-1.5 cursor-pointer">
                    <Upload size={14} /> Importer XLS
                    <input type="file" accept=".xlsx,.xls" className="hidden" onChange={async (e) => {
                      const file = e.target.files?.[0]; if (!file) return;
                      try {
                        const buffer = await file.arrayBuffer(); const wb = XLSX.read(buffer, { type: 'array' });
                        const ws = wb.Sheets[wb.SheetNames[0]];
                        const rawRows: any[] = XLSX.utils.sheet_to_json(ws, { header: 1 });
                        const dataRows = rawRows.slice(1).filter((r: any[]) => r.length > 0 && (r[0]||r[2]));
                        const records = dataRows.map((r: any[]) => ({
                          company_id: companyId, raison_social: String(r[0]||'').trim()||null, rib: String(r[2]||r[1]||'').trim()||null,
                        }));
                        if (!records.length) { toast.error("Aucune donnée."); return; }
                        const { error } = await supabase.from('bank_base_rip').insert(records);
                        if (!error) { toast.success(`${records.length} Base RIP importés.`); fetchBaseRip(); } else toast.error(`Erreur: ${error.message}`);
                      } catch (err: any) { toast.error(`Erreur: ${err.message}`); }
                      e.target.value = '';
                    }} />
                  </label>
                  <button onClick={() => { if (!baseRipList.length) return; exportToXLS(baseRipList.map((r:any) => ({ 'Raison Sociale':r.raison_social, 'RIB':r.rib })),'base_rip'); }}
                    className="bg-emerald-600 hover:bg-emerald-700 text-white px-3 py-2 rounded-lg text-xs font-black uppercase tracking-wider flex items-center gap-1.5 cursor-pointer"><Download size={14} /> Export XLS</button>
                  <button onClick={() => { setBaseRipForm(emptyBaseRipForm); setEditingBaseRip(null); setShowBaseRipForm(true); }}
                    className="bg-blue-600 hover:bg-blue-700 text-white px-3 py-2 rounded-lg text-xs font-black uppercase tracking-wider flex items-center gap-1.5 cursor-pointer"><Plus size={14} /> Nouveau</button>
                </div>
              </div>
            </div>
            {loadingBaseRip ? <div className="flex justify-center py-20"><Loader2 className="w-6 h-6 text-blue-600 animate-spin" /></div> : (
              <div className="bg-white rounded-xl border border-slate-200 overflow-hidden">
                <div className="overflow-x-auto">
                  <table className="w-full text-left">
                    <thead className="bg-slate-50 border-b border-slate-200">
                      <tr>
                        {['Raison Sociale','RIB','Actions'].map(h => (
                          <th key={h} className="px-4 py-2 text-[9px] font-black text-slate-400 uppercase tracking-widest">{h}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-100">
                      {baseRipList.length === 0 ? (
                        <tr><td colSpan={3} className="px-4 py-10 text-center text-sm text-slate-400">Aucune donnée.</td></tr>
                      ) : baseRipList.map((r: any) => (
                        <tr key={r.id} className="hover:bg-slate-50 transition-colors">
                          <td className="px-4 py-3 text-xs font-semibold text-slate-700">{r.raison_social||'—'}</td>
                          <td className="px-4 py-3 font-mono text-xs text-blue-600 font-bold">{r.rib||'—'}</td>
                          <td className="px-4 py-3 flex gap-1">
                            <button onClick={() => { setEditingBaseRip(r); setBaseRipForm({...r}); setShowBaseRipForm(true); }} className="text-slate-400 hover:text-blue-600 cursor-pointer"><Pencil size={13} /></button>
                            <button onClick={() => handleDeleteBaseRip(r.id)} className="text-slate-400 hover:text-rose-600 cursor-pointer"><Trash2 size={13} /></button>
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
        {activeTab === 'bank_virement' && (
          <div>
            <div className="mb-6 bg-slate-900 text-white rounded-xl p-6 border border-slate-800">
              <div className="flex items-center justify-between flex-wrap gap-3">
                <div>
                  <span className="inline-flex items-center gap-1.5 px-2 py-0.5 rounded text-[10px] font-extrabold uppercase tracking-widest bg-indigo-500 text-white mb-2"><Landmark className="w-3.5 h-3.5" /> Virement</span>
                  <h1 className="text-2xl font-extrabold tracking-tight">Ordres de Virement</h1>
                  <p className="text-sm text-slate-400 mt-1">{filteredVirements.length} virements — {selectedVirements.length} sélectionnés</p>
                </div>
                <div className="flex gap-2 flex-wrap">
                  <button onClick={fetchVirement} className="bg-white/10 hover:bg-white/15 px-3 py-2 rounded-lg text-xs font-bold uppercase tracking-wider flex items-center gap-1.5 cursor-pointer"><RefreshCw size={14} /> Actualiser</button>
                  <label className="bg-amber-600 hover:bg-amber-700 text-white px-3 py-2 rounded-lg text-xs font-black uppercase tracking-wider flex items-center gap-1.5 cursor-pointer">
                    <Upload size={14} /> Importer XLS
                    <input type="file" accept=".xlsx,.xls" className="hidden" onChange={async (e) => {
                      const file = e.target.files?.[0]; if (!file) return;
                      try {
                        const buffer = await file.arrayBuffer(); const wb = XLSX.read(buffer, { type: 'array' });
                        const ws = wb.Sheets[wb.SheetNames[0]];
                        const rawRows: any[] = XLSX.utils.sheet_to_json(ws, { header: 1 });
                        const dataRows = rawRows.slice(2).filter((r: any[]) => r.length > 0 && (r[0]||r[1]));
                        const records = dataRows.map((r: any[]) => ({
                          company_id: companyId, date_virement: r[0]?String(r[0]):null, raison_social: String(r[1]||'').trim()||null,
                          montant: parseFloat(r[2]) || 0, justification: String(r[3]||'').trim()||null, tva_mois: String(r[4]||'').trim()||null,
                        }));
                        if (!records.length) { toast.error("Aucune donnée."); return; }
                        const { error } = await supabase.from('bank_virement').insert(records);
                        if (!error) { toast.success(`${records.length} virements importés.`); fetchVirement(); } else toast.error(`Erreur: ${error.message}`);
                      } catch (err: any) { toast.error(`Erreur: ${err.message}`); }
                      e.target.value = '';
                    }} />
                  </label>
                  <button onClick={() => { if (!filteredVirements.length) return; exportToXLS(filteredVirements.map((v:any) => ({ 'Date':v.date_virement,'Raison Sociale':v.raison_social,'RIB':v.rib,'Montant':v.montant,'Justification':v.justification,'TVA/Mois':v.tva_mois,'Banque':v.banque,'Agence':v.agence,'N° Compte':v.numero_compte })),'virements'); }}
                    className="bg-emerald-600 hover:bg-emerald-700 text-white px-3 py-2 rounded-lg text-xs font-black uppercase tracking-wider flex items-center gap-1.5 cursor-pointer"><Download size={14} /> Export XLS</button>
                  {selectedVirements.length > 0 && (
                    <button onClick={handleGenerateVirementPDF} className="bg-violet-600 hover:bg-violet-700 text-white px-3 py-2 rounded-lg text-xs font-black uppercase tracking-wider flex items-center gap-1.5 cursor-pointer"><FileText size={14} /> Générer PDF ({selectedVirements.length})</button>
                  )}
                  <button onClick={() => { setVirementForm(emptyVirementForm); setEditingVirement(null); setShowVirementForm(true); }}
                    className="bg-blue-600 hover:bg-blue-700 text-white px-3 py-2 rounded-lg text-xs font-black uppercase tracking-wider flex items-center gap-1.5 cursor-pointer"><Plus size={14} /> Nouveau</button>
                </div>
              </div>
            </div>
            {/* Filters */}
            <div className="mb-4 bg-white rounded-xl border border-slate-200 p-4 flex flex-wrap gap-3 items-end">
              <div><label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Raison Sociale</label>
                <input type="text" placeholder="Filtrer..." value={virementFilter.raison} onChange={e => setVirementFilter(p => ({...p,raison:e.target.value}))} className="block mt-1 h-8 rounded-lg border-2 border-slate-200 px-3 text-xs focus:outline-none focus:border-blue-500 w-44" /></div>
              <div><label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Date de</label>
                <input type="date" value={virementFilter.dateFrom} onChange={e => setVirementFilter(p => ({...p,dateFrom:e.target.value}))} className="block mt-1 h-8 rounded-lg border-2 border-slate-200 px-3 text-xs focus:outline-none focus:border-blue-500" /></div>
              <div><label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Date à</label>
                <input type="date" value={virementFilter.dateTo} onChange={e => setVirementFilter(p => ({...p,dateTo:e.target.value}))} className="block mt-1 h-8 rounded-lg border-2 border-slate-200 px-3 text-xs focus:outline-none focus:border-blue-500" /></div>
              <button onClick={() => setVirementFilter({raison:'',dateFrom:'',dateTo:''})} className="h-8 px-3 bg-slate-100 hover:bg-slate-200 text-slate-600 text-xs font-bold rounded-lg cursor-pointer">Réinitialiser</button>
            </div>
            {/* Table */}
            {loadingVirement ? <div className="flex justify-center py-20"><Loader2 className="w-6 h-6 text-blue-600 animate-spin" /></div> : (
              <div className="bg-white rounded-xl border border-slate-200 overflow-hidden">
                <div className="overflow-x-auto">
                  <table className="w-full text-left min-w-[900px]">
                    <thead className="bg-slate-50 border-b border-slate-200">
                      <tr>
                        <th className="px-3 py-3 w-8"><input type="checkbox" checked={selectedVirements.length === filteredVirements.length && filteredVirements.length > 0} onChange={e => setSelectedVirements(e.target.checked ? filteredVirements.map((v:any)=>v.id) : [])} className="accent-blue-600" /></th>
                        {['Date','Raison Sociale','RIB','Montant','Justification','TVA/Mois','Banque','Agence','Actions'].map(h => (
                          <th key={h} className="px-3 py-2 text-[9px] font-black text-slate-400 uppercase tracking-widest whitespace-nowrap">{h}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-100">
                      {filteredVirements.length === 0 ? (
                        <tr><td colSpan={10} className="px-4 py-10 text-center text-sm text-slate-400">Aucun virement.</td></tr>
                      ) : filteredVirements.map((v: any) => (
                        <tr key={v.id} className={`hover:bg-slate-50 transition-colors ${selectedVirements.includes(v.id)?'bg-blue-50/50':''}`}>
                          <td className="px-3 py-3"><input type="checkbox" checked={selectedVirements.includes(v.id)} onChange={e => setSelectedVirements(prev => e.target.checked ? [...prev,v.id] : prev.filter(x=>x!==v.id))} className="accent-blue-600" /></td>
                          <td className="px-3 py-2 text-xs text-slate-700">{v.date_virement||'—'}</td>
                          <td className="px-3 py-2 text-xs font-semibold text-slate-700">{v.raison_social||'—'}</td>
                          <td className="px-3 py-2 font-mono text-xs text-blue-600">{v.rib||'—'}</td>
                          <td className="px-3 py-2 font-mono text-xs font-bold text-slate-900">{Number(v.montant||0).toLocaleString('fr-MA',{minimumFractionDigits:2})}</td>
                          <td className="px-3 py-2 text-xs text-slate-600">{v.justification||'—'}</td>
                          <td className="px-3 py-2 text-xs text-slate-500">{v.tva_mois||'—'}</td>
                          <td className="px-3 py-2 text-xs text-slate-500">{v.banque||'—'}</td>
                          <td className="px-3 py-2 text-xs text-slate-500">{v.agence||'—'}</td>
                          <td className="px-3 py-2 flex gap-1">
                            <button onClick={() => { setEditingVirement(v); setVirementForm({...v, montant: String(v.montant)}); setShowVirementForm(true); }} className="text-slate-400 hover:text-blue-600 cursor-pointer"><Pencil size={13} /></button>
                            <button onClick={() => handleDeleteVirement(v.id)} className="text-slate-400 hover:text-rose-600 cursor-pointer"><Trash2 size={13} /></button>
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
        {activeTab === 'bank_releve' && (
          <div>
            <div className="mb-6 bg-slate-900 text-white rounded-xl p-6 border border-slate-800">
              <div className="flex items-center justify-between flex-wrap gap-3">
                <div>
                  <span className="inline-flex items-center gap-1.5 px-2 py-0.5 rounded text-[10px] font-extrabold uppercase tracking-widest bg-teal-500 text-white mb-2"><Landmark className="w-3.5 h-3.5" /> Relevé Bancaire</span>
                  <h1 className="text-2xl font-extrabold tracking-tight">Relevé de Compte Bancaire</h1>
                  <p className="text-sm text-slate-400 mt-1">{filteredReleve.length} transactions</p>
                </div>
                <div className="flex gap-2 flex-wrap items-center">
                  <input type="month" value={releveMois} onChange={e => setReleveMois(e.target.value)}
                    className="h-9 rounded-lg border-2 border-white/20 bg-white/10 px-3 text-xs text-white focus:outline-none" />
                  <button onClick={() => { setReleveFilter({mois:'',libelle:'',category:''}); fetchReleve(); }} className="bg-white/10 hover:bg-white/15 px-3 py-2 rounded-lg text-xs font-bold uppercase tracking-wider flex items-center gap-1.5 cursor-pointer"><RefreshCw size={14} /> Actualiser</button>
                  <label className={`${parsingPDF ? 'bg-amber-400' : 'bg-amber-600 hover:bg-amber-700'} text-white px-3 py-2 rounded-lg text-xs font-black uppercase tracking-wider flex items-center gap-1.5 cursor-pointer`}>
                    {parsingPDF ? <><Loader2 size={14} className="animate-spin" /> Extraction...</> : <><Upload size={14} /> Importer PDF</>}
                    <input type="file" accept=".pdf" className="hidden" disabled={parsingPDF} onChange={e => { const f = e.target.files?.[0]; if (f) parseBankPDF(f); e.target.value=''; }} />
                  </label>
                  <button onClick={() => { if (!filteredReleve.length) return; exportToXLS(filteredReleve.map((r:any) => ({ 'Mois':r.mois,'Code':r.code,'Date Op.':r.date_operation,'Libellé':r.libelle,'Date Valeur':r.date_valeur,'Débit':r.debit,'Crédit':r.credit,'Catégorie':r.category })),'releve_bancaire'); }}
                    className="bg-emerald-600 hover:bg-emerald-700 text-white px-3 py-2 rounded-lg text-xs font-black uppercase tracking-wider flex items-center gap-1.5 cursor-pointer"><Download size={14} /> Export XLS</button>
                  <button onClick={() => { setReleveForm(emptyReleveForm); setEditingReleve(null); setShowReleveForm(true); }}
                    className="bg-blue-600 hover:bg-blue-700 text-white px-3 py-2 rounded-lg text-xs font-black uppercase tracking-wider flex items-center gap-1.5 cursor-pointer"><Plus size={14} /> Manuel</button>
                </div>
              </div>
            </div>
            {/* Filters */}
            <div className="mb-4 bg-white rounded-xl border border-slate-200 p-4 flex flex-wrap gap-3 items-end">
              <div><label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Mois</label>
                <input type="month" value={releveFilter.mois} onChange={e => setReleveFilter(p => ({...p,mois:e.target.value}))} className="block mt-1 h-8 rounded-lg border-2 border-slate-200 px-3 text-xs focus:outline-none focus:border-blue-500" /></div>
              <div><label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Libellé</label>
                <input type="text" placeholder="Rechercher..." value={releveFilter.libelle} onChange={e => setReleveFilter(p => ({...p,libelle:e.target.value}))} className="block mt-1 h-8 rounded-lg border-2 border-slate-200 px-3 text-xs focus:outline-none focus:border-blue-500 w-44" /></div>
              <div><label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Catégorie</label>
                <select value={releveFilter.category} onChange={e => setReleveFilter(p => ({...p,category:e.target.value}))} className="block mt-1 h-8 rounded-lg border-2 border-slate-200 px-3 text-xs focus:outline-none focus:border-blue-500">
                  <option value="">Toutes</option>
                  {ETAT_CATEGORIES.map(c => (<option key={c.id} value={c.id}>{c.label}</option>))}
                </select></div>
              <button onClick={() => setReleveFilter({mois:'',libelle:'',category:''})} className="h-8 px-3 bg-slate-100 hover:bg-slate-200 text-slate-600 text-xs font-bold rounded-lg cursor-pointer">Réinitialiser</button>
            </div>
            {/* Totals */}
            {filteredReleve.length > 0 && (() => {
              const tD = filteredReleve.reduce((s: number, r: any) => s + (parseFloat(r.debit)||0), 0);
              const tC = filteredReleve.reduce((s: number, r: any) => s + (parseFloat(r.credit)||0), 0);
              const fmt2 = (n: number) => n.toLocaleString('fr-MA',{minimumFractionDigits:2});
              return <div className="mb-4 bg-white rounded-xl border border-slate-200 p-4 flex gap-6">
                <div><span className="text-[9px] font-black text-slate-400 uppercase block">Total Débit</span><span className="text-sm font-bold text-rose-700">{fmt2(tD)}</span></div>
                <div><span className="text-[9px] font-black text-slate-400 uppercase block">Total Crédit</span><span className="text-sm font-bold text-emerald-700">{fmt2(tC)}</span></div>
                <div><span className="text-[9px] font-black text-slate-400 uppercase block">Solde</span><span className="text-sm font-black text-blue-700">{fmt2(tC - tD)}</span></div>
              </div>;
            })()}
            {/* Table */}
            {loadingReleve ? <div className="flex justify-center py-20"><Loader2 className="w-6 h-6 text-blue-600 animate-spin" /></div> : (
              <div className="bg-white rounded-xl border border-slate-200 overflow-hidden">
                <div className="overflow-x-auto">
                  <table className="w-full text-left min-w-[900px]">
                    <thead className="bg-slate-50 border-b border-slate-200">
                      <tr>{['Code','Date Op.','Libellé','Valeur','Débit','Crédit','Catégorie','Actions'].map(h => (
                        <th key={h} className="px-3 py-2 text-[9px] font-black text-slate-400 uppercase tracking-widest">{h}</th>
                      ))}</tr>
                    </thead>
                    <tbody className="divide-y divide-slate-100">
                      {filteredReleve.length === 0 ? <tr><td colSpan={8} className="px-4 py-10 text-center text-sm text-slate-400">Aucune transaction. Uploadez un relevé PDF.</td></tr>
                      : filteredReleve.map((r: any) => (
                        <tr key={r.id} className="hover:bg-slate-50">
                          <td className="px-3 py-2 font-mono text-[10px] text-slate-500">{r.code||'—'}</td>
                          <td className="px-3 py-2 text-xs text-slate-700">{r.date_operation||'—'}</td>
                          <td className="px-3 py-2 text-xs text-slate-700 max-w-[250px] truncate">{r.libelle||'—'}</td>
                          <td className="px-3 py-2 text-xs text-slate-500">{r.date_valeur||'—'}</td>
                          <td className="px-3 py-2 font-mono text-xs text-rose-600 font-bold">{parseFloat(r.debit)>0 ? Number(r.debit).toLocaleString('fr-MA',{minimumFractionDigits:2}) : ''}</td>
                          <td className="px-3 py-2 font-mono text-xs text-emerald-600 font-bold">{parseFloat(r.credit)>0 ? Number(r.credit).toLocaleString('fr-MA',{minimumFractionDigits:2}) : ''}</td>
                          <td className="px-3 py-2"><span className={`text-[8px] font-black px-1.5 py-0.5 rounded uppercase ${r.category==='virement'?'bg-blue-50 text-blue-700':r.category==='emission_cheque'?'bg-amber-50 text-amber-700':r.category==='emission_effets'?'bg-purple-50 text-purple-700':r.category==='remise_cheque'?'bg-emerald-50 text-emerald-700':r.category==='remise_lc'?'bg-cyan-50 text-cyan-700':'bg-rose-50 text-rose-700'}`}>{ETAT_CATEGORIES.find(c=>c.id===r.category)?.label||r.category}</span></td>
                          <td className="px-3 py-2 flex gap-1">
                            <button onClick={() => { setEditingReleve(r); setReleveForm({...r,debit:String(r.debit),credit:String(r.credit)}); setShowReleveForm(true); }} className="text-slate-400 hover:text-blue-600 cursor-pointer"><Pencil size={12} /></button>
                            <button onClick={() => handleDeleteReleve(r.id)} className="text-slate-400 hover:text-rose-600 cursor-pointer"><Trash2 size={12} /></button>
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
        {activeTab === 'bank_etat_explicatif' && (
          <div>
            <div className="mb-6 bg-slate-900 text-white rounded-xl p-6 border border-slate-800">
              <div className="flex items-center justify-between flex-wrap gap-3">
                <div>
                  <span className="inline-flex items-center gap-1.5 px-2 py-0.5 rounded text-[10px] font-extrabold uppercase tracking-widest bg-orange-500 text-white mb-2"><FileText className="w-3.5 h-3.5" /> État Explicatif</span>
                  <h1 className="text-2xl font-extrabold tracking-tight">État Explicatif du Relevé Bancaire</h1>
                  <p className="text-sm text-slate-400 mt-1">{releveList.length} transactions — {ETAT_CATEGORIES.length} sections</p>
                </div>
                <div className="flex gap-2 flex-wrap items-center">
                  <input type="month" value={releveFilter.mois || ''} onChange={e => setReleveFilter(p => ({...p,mois:e.target.value}))}
                    className="h-9 rounded-lg border-2 border-white/20 bg-white/10 px-3 text-xs text-white focus:outline-none" />
                  <button onClick={fetchReleve} className="bg-white/10 hover:bg-white/15 px-3 py-2 rounded-lg text-xs font-bold uppercase tracking-wider flex items-center gap-1.5 cursor-pointer"><RefreshCw size={14} /> Actualiser</button>
                  <button onClick={() => {
                    const rows: any[] = [];
                    ETAT_CATEGORIES.forEach(cat => {
                      const items = filteredReleve.filter((r:any) => r.category === cat.id);
                      rows.push({ 'Date':'', 'Libellé': `═══ ${cat.label.toUpperCase()} ═══`, 'Destination':'', 'Note':'', 'Décaissement':'', 'Encaissement':'', 'Réf. Règlement':'', 'Observation':'' });
                      items.forEach((r:any) => rows.push({ 'Date':r.date_operation, 'Libellé':r.libelle, 'Destination':r.destination||'', 'Note':r.note_operation||'', 'Décaissement':parseFloat(r.debit)||'', 'Encaissement':parseFloat(r.credit)||'', 'Réf. Règlement':r.ref_reglement||'', 'Observation':r.observation||'' }));
                      const subD = items.reduce((s:number,r:any)=>s+(parseFloat(r.debit)||0),0);
                      const subC = items.reduce((s:number,r:any)=>s+(parseFloat(r.credit)||0),0);
                      rows.push({ 'Date':'', 'Libellé':'SOUS-TOTAL', 'Destination':'', 'Note':'', 'Décaissement':subD||'', 'Encaissement':subC||'', 'Réf. Règlement':'', 'Observation':'' });
                    });
                    exportToXLS(rows, 'etat_explicatif');
                  }}
                    className="bg-emerald-600 hover:bg-emerald-700 text-white px-3 py-2 rounded-lg text-xs font-black uppercase tracking-wider flex items-center gap-1.5 cursor-pointer"><Download size={14} /> Export XLS</button>
                </div>
              </div>
            </div>
            {/* Solde départ */}
            {filteredReleve.length > 0 && (
              <div className="mb-4 bg-amber-50 border border-amber-200 rounded-xl p-4 flex justify-between items-center">
                <span className="text-xs font-black text-amber-800 uppercase">Solde de départ</span>
                <span className="font-mono text-sm font-bold text-amber-900">{Number(filteredReleve[0]?.solde_depart||0).toLocaleString('fr-MA',{minimumFractionDigits:2})} DÉBITEUR</span>
              </div>
            )}
            {/* Sections */}
            {ETAT_CATEGORIES.map(cat => {
              const items = filteredReleve.filter((r: any) => r.category === cat.id);
              const isExp = etatExpandedSections.includes(cat.id);
              const subDebit = items.reduce((s: number, r: any) => s + (parseFloat(r.debit) || 0), 0);
              const subCredit = items.reduce((s: number, r: any) => s + (parseFloat(r.credit) || 0), 0);
              const fmt2 = (n: number) => n > 0 ? n.toLocaleString('fr-MA', { minimumFractionDigits: 2 }) : '';
              return (
                <div key={cat.id} className="mb-3 bg-white rounded-xl border border-slate-200 overflow-hidden">
                  <div onClick={() => setEtatExpandedSections(prev => prev.includes(cat.id) ? prev.filter(x => x !== cat.id) : [...prev, cat.id])}
                    className={`flex items-center justify-between px-5 py-3 cursor-pointer hover:bg-slate-50 border-l-4 ${cat.color==='blue'?'border-blue-500':cat.color==='amber'?'border-amber-500':cat.color==='purple'?'border-purple-500':cat.color==='emerald'?'border-emerald-500':cat.color==='cyan'?'border-cyan-500':'border-rose-500'}`}>
                    <div className="flex items-center gap-3">
                      <span className="text-sm font-black text-slate-800">{cat.label}</span>
                      <span className="text-[10px] font-bold text-slate-400">{items.length} lignes</span>
                    </div>
                    <div className="flex items-center gap-4">
                      <span className="font-mono text-xs text-rose-600">{fmt2(subDebit)}</span>
                      <span className="font-mono text-xs text-emerald-600">{fmt2(subCredit)}</span>
                      <span className="text-slate-400">{isExp ? '▼' : '▶'}</span>
                    </div>
                  </div>
                  {isExp && items.length > 0 && (
                    <div className="overflow-x-auto">
                      <table className="w-full text-left min-w-[900px]">
                        <thead className="bg-slate-50 border-b border-slate-200">
                          <tr>{['✓','Date','Libellé','Code Rèq.','Destination','Note','Décaissement','Encaissement','Réf.','Obs.',''].map(h => (
                            <th key={h} className="px-3 py-1.5 text-[8px] font-black text-slate-400 uppercase tracking-widest">{h}</th>
                          ))}</tr>
                        </thead>
                        <tbody className="divide-y divide-slate-50">
                          {items.map((r: any) => (
                            <tr key={r.id} className={`hover:bg-slate-50 ${checkedReleve.includes(r.id) ? 'bg-emerald-50' : ''}`}>
                              <td className="px-3 py-1.5">
                                <input type="checkbox" checked={checkedReleve.includes(r.id)}
                                  onChange={e => setCheckedReleve(prev => e.target.checked ? [...prev, r.id] : prev.filter(x => x !== r.id))}
                                  className="accent-emerald-600" />
                              </td>
                              <td className="px-3 py-1.5 text-[10px] text-slate-600">{r.date_operation}</td>
                              <td className="px-3 py-1.5 text-[10px] text-slate-700 max-w-[200px] truncate">{r.libelle}</td>
                              <td className="px-3 py-1.5">
                                <div className="flex items-center gap-1">
                                  <input type="text" value={r?.code_reglement || ''} placeholder="—"
                                    onChange={e => {
                                      const val = e.target.value;
                                      setReleveList(prev => prev.map(x => x.id === r.id ? { ...x, code_reglement: val } : x));
                                    }}
                                    onBlur={async (e) => {
                                      await supabase.from('bank_releve').update({ code_reglement: e.target.value || null }).eq('id', r.id);
                                    }}
                                    className="w-16 h-6 rounded border border-amber-200 bg-amber-50 px-1 text-[9px] font-mono font-bold text-amber-700 focus:outline-none focus:border-amber-500" />
                                  <button onClick={async () => {
                                    const code = r?.code_reglement?.trim();
                                    if (!code) { toast.error("Saisissez un code."); return; }
                                    // Save code first
                                    await supabase.from('bank_releve').update({ code_reglement: code }).eq('id', r.id);
                                    // Search in reglements
                                    const { data: regData } = await supabase.from('reglements').select('*').eq('company_id', companyId).eq('code_reglement', code).limit(1);
                                    if (regData && regData.length > 0) {
                                      const reg = regData[0];
                                      const factNums = (reg.facture_numbers || []).join(', ');
                                      const updates = {
                                        destination: reg.client || null,
                                        note_operation: factNums ? `Fact: ${factNums}` : null,
                                        ref_reglement: reg.numero || reg.reference_virement || null,
                                      };
                                      await supabase.from('bank_releve').update(updates).eq('id', r.id);
                                      setReleveList(prev => prev.map(x => x.id === r.id ? { ...x, ...updates, code_reglement: code } : x));
                                      toast.success(`Trouvé: ${reg.client}`);
                                      return;
                                    }
                                    // Search in purchases
                                    const { data: purchData } = await supabase.from('purchases').select('*').eq('company_id', companyId).eq('code_reglement', code);
                                    if (purchData && purchData.length > 0) {
                                      const fourn = purchData[0].fournisseur || purchData[0].supplier || '';
                                      const details = purchData.map((p: any) => p.numero_facture || p.invoice_number || '').filter(Boolean).join(', ');
                                      const updates = {
                                        destination: fourn || null,
                                        note_operation: details || null,
                                      };
                                      await supabase.from('bank_releve').update(updates).eq('id', r.id);
                                      setReleveList(prev => prev.map(x => x.id === r.id ? { ...x, ...updates, code_reglement: code } : x));
                                      toast.success(`Trouvé: ${fourn}`);
                                      return;
                                    }
                                    toast.error("Code non trouvé.");
                                  }} className="text-amber-400 hover:text-amber-600 cursor-pointer"><Search size={10} /></button>
                                </div>
                              </td>
                              <td className="px-3 py-1.5 text-[10px] text-slate-500">{r.destination||'—'}</td>
                              <td className="px-3 py-1.5 text-[10px] text-slate-500">{r.note_operation||'—'}</td>
                              <td className="px-3 py-1.5 font-mono text-[10px] text-rose-600 font-bold">{parseFloat(r.debit)>0?Number(r.debit).toLocaleString('fr-MA',{minimumFractionDigits:2}):''}</td>
                              <td className="px-3 py-1.5 font-mono text-[10px] text-emerald-600 font-bold">{parseFloat(r.credit)>0?Number(r.credit).toLocaleString('fr-MA',{minimumFractionDigits:2}):''}</td>
                              <td className="px-3 py-1.5 text-[10px] text-slate-500">{r.ref_reglement||'—'}</td>
                              <td className="px-3 py-1.5 text-[10px] text-slate-500">{r.observation||'—'}</td>
                              <td className="px-3 py-1.5"><button onClick={() => { setEditingReleve(r); setReleveForm({...r,debit:String(r.debit),credit:String(r.credit)}); setShowReleveForm(true); }} className="text-slate-300 hover:text-blue-600 cursor-pointer"><Pencil size={11} /></button></td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  )}
                </div>
              );
            })}
            {/* Totals + Solde Final */}
            {filteredReleve.length > 0 && (() => {
              const tD = filteredReleve.reduce((s:number,r:any)=>s+(parseFloat(r.debit)||0),0);
              const tC = filteredReleve.reduce((s:number,r:any)=>s+(parseFloat(r.credit)||0),0);
              const fmt2 = (n:number) => n.toLocaleString('fr-MA',{minimumFractionDigits:2});
              return <>
                <div className="mt-4 bg-slate-900 rounded-xl p-4 flex justify-between items-center">
                  <span className="text-xs font-black text-white uppercase">Total Mouvements</span>
                  <div className="flex gap-6">
                    <span className="font-mono text-sm font-bold text-rose-400">{fmt2(tD)}</span>
                    <span className="font-mono text-sm font-bold text-emerald-400">{fmt2(tC)}</span>
                  </div>
                </div>
                <div className="mt-2 bg-blue-50 border border-blue-200 rounded-xl p-4 flex justify-between items-center">
                  <span className="text-xs font-black text-blue-800 uppercase">Solde Final</span>
                  <span className="font-mono text-sm font-bold text-blue-900">{fmt2(Math.abs(tC - tD + (parseFloat(filteredReleve[0]?.solde_depart)||0)))} {tC - tD + (parseFloat(filteredReleve[0]?.solde_depart)||0) >= 0 ? 'CRÉDITEUR' : 'DÉBITEUR'}</span>
                </div>
              </>;
            })()}
          </div>
        )}

{activeTab === 'settings' && (
          <InvoiceEngine companyId={companyId} logoPreviewUrl={logoPreviewUrl} />
        )}
        </main>
      </div>
      {/* RIP Form */}
<AnimatePresence>
  {showRipForm && (
    <div className="fixed inset-0 bg-black/50 z-50 flex items-center justify-center p-4">
      <motion.div initial={{ scale: 0.95, opacity: 0 }} animate={{ scale: 1, opacity: 1 }} exit={{ scale: 0.95, opacity: 0 }}
        className="bg-white rounded-xl p-6 max-w-2xl w-full shadow-xl">
        <div className="flex items-center justify-between mb-5">
          <h3 className="font-black text-slate-900 uppercase tracking-widest text-sm">{editingRip ? 'Modifier RIP' : 'Nouveau RIP'}</h3>
          <button onClick={() => { setShowRipForm(false); setEditingRip(null); setRipForm(emptyRipForm); }} className="p-1.5 rounded hover:bg-slate-100 text-slate-400"><X size={16} /></button>
        </div>
        <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
          {[
            { label: 'Banque', key: 'banque' },
            { label: 'Agence', key: 'agence' },
            { label: 'Code Banque', key: 'code_banque' },
            { label: 'Code Ville', key: 'code_ville' },
            { label: 'N° Compte', key: 'numero_compte' },
            { label: 'Clé RIP', key: 'cle_rip' },
            { label: 'Code SWIFT', key: 'code_swift' },
          ].map(({ label, key }) => (
            <div key={key}>
              <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">{label}</label>
              <input type="text" value={ripForm[key] || ''} onChange={e => setRipForm((p: any) => ({ ...p, [key]: e.target.value }))}
                className="w-full mt-1 h-9 rounded-lg border-2 border-slate-200 px-3 text-sm focus:outline-none focus:border-blue-500" />
            </div>
          ))}
        </div>
        <div className="flex gap-3 pt-5">
          <button onClick={handleSaveRip} className="flex-1 h-10 bg-blue-600 hover:bg-blue-700 text-white text-sm font-bold rounded-lg cursor-pointer">{editingRip ? 'Enregistrer' : 'Ajouter'}</button>
          <button onClick={() => { setShowRipForm(false); setEditingRip(null); setRipForm(emptyRipForm); }} className="flex-1 h-10 bg-slate-100 hover:bg-slate-200 text-slate-700 text-sm font-bold rounded-lg cursor-pointer">Annuler</button>
        </div>
      </motion.div>
    </div>
  )}
</AnimatePresence>

{/* Base RIP Form */}
<AnimatePresence>
  {showBaseRipForm && (
    <div className="fixed inset-0 bg-black/50 z-50 flex items-center justify-center p-4">
      <motion.div initial={{ scale: 0.95, opacity: 0 }} animate={{ scale: 1, opacity: 1 }} exit={{ scale: 0.95, opacity: 0 }}
        className="bg-white rounded-xl p-6 max-w-lg w-full shadow-xl">
        <div className="flex items-center justify-between mb-5">
          <h3 className="font-black text-slate-900 uppercase tracking-widest text-sm">{editingBaseRip ? 'Modifier' : 'Nouveau Base RIP'}</h3>
          <button onClick={() => { setShowBaseRipForm(false); setEditingBaseRip(null); setBaseRipForm(emptyBaseRipForm); }} className="p-1.5 rounded hover:bg-slate-100 text-slate-400"><X size={16} /></button>
        </div>
        <div className="grid grid-cols-1 gap-4">
          <div>
            <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Raison Sociale</label>
            <input type="text" value={baseRipForm.raison_social || ''} onChange={e => setBaseRipForm((p: any) => ({ ...p, raison_social: e.target.value }))}
              className="w-full mt-1 h-9 rounded-lg border-2 border-slate-200 px-3 text-sm focus:outline-none focus:border-blue-500" />
          </div>
          <div>
            <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">RIB</label>
            <input type="text" value={baseRipForm.rib || ''} onChange={e => setBaseRipForm((p: any) => ({ ...p, rib: e.target.value }))}
              className="w-full mt-1 h-9 rounded-lg border-2 border-slate-200 px-3 text-sm focus:outline-none focus:border-blue-500" />
          </div>
        </div>
        <div className="flex gap-3 pt-5">
          <button onClick={handleSaveBaseRip} className="flex-1 h-10 bg-blue-600 hover:bg-blue-700 text-white text-sm font-bold rounded-lg cursor-pointer">{editingBaseRip ? 'Enregistrer' : 'Ajouter'}</button>
          <button onClick={() => { setShowBaseRipForm(false); setEditingBaseRip(null); setBaseRipForm(emptyBaseRipForm); }} className="flex-1 h-10 bg-slate-100 hover:bg-slate-200 text-slate-700 text-sm font-bold rounded-lg cursor-pointer">Annuler</button>
        </div>
      </motion.div>
    </div>
  )}
</AnimatePresence>
{/* Virement Form */}
<AnimatePresence>
  {showVirementForm && (
    <div className="fixed inset-0 bg-black/50 z-50 flex items-center justify-center p-4">
      <motion.div initial={{ scale: 0.95, opacity: 0 }} animate={{ scale: 1, opacity: 1 }} exit={{ scale: 0.95, opacity: 0 }}
        className="bg-white rounded-xl p-6 max-w-3xl w-full shadow-xl max-h-[90vh] overflow-y-auto">
        <div className="flex items-center justify-between mb-5">
          <h3 className="font-black text-slate-900 uppercase tracking-widest text-sm">{editingVirement ? 'Modifier Virement' : 'Nouveau Virement'}</h3>
          <button onClick={() => { setShowVirementForm(false); setEditingVirement(null); setVirementForm(emptyVirementForm); }} className="p-1.5 rounded hover:bg-slate-100 text-slate-400"><X size={16} /></button>
        </div>
        <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-4">
          <div>
            <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Date</label>
            <input type="date" value={virementForm.date_virement || ''} onChange={e => setVirementForm((p: any) => ({...p, date_virement: e.target.value}))}
              className="w-full mt-1 h-9 rounded-lg border-2 border-slate-200 px-3 text-sm focus:outline-none focus:border-blue-500" />
          </div>

          {/* Raison Sociale — auto-fill from Base RIP */}
          <div>
            <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Raison Sociale</label>
            <select value={virementForm.raison_social || ''} onChange={e => {
              const selected = baseRipList.find((b: any) => b.raison_social === e.target.value);
              setVirementForm((p: any) => ({...p, raison_social: e.target.value, rib: selected?.rib || p.rib }));
            }}
              className="w-full mt-1 h-9 rounded-lg border-2 border-slate-200 px-3 text-sm focus:outline-none focus:border-blue-500">
              <option value="">— Sélectionner —</option>
              {baseRipList.map((b: any) => (<option key={b.id} value={b.raison_social}>{b.raison_social}</option>))}
            </select>
          </div>

          {/* RIB — auto-filled */}
          <div>
            <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">RIB (auto)</label>
            <input type="text" value={virementForm.rib || ''} onChange={e => setVirementForm((p: any) => ({...p, rib: e.target.value}))}
              className="w-full mt-1 h-9 rounded-lg border-2 border-blue-100 bg-blue-50 px-3 text-sm font-mono text-blue-700 focus:outline-none focus:border-blue-500" />
          </div>

          <div>
            <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Montant (MAD)</label>
            <input type="number" value={virementForm.montant || ''} onChange={e => setVirementForm((p: any) => ({...p, montant: e.target.value}))}
              className="w-full mt-1 h-9 rounded-lg border-2 border-slate-200 px-3 text-sm focus:outline-none focus:border-blue-500" />
          </div>
          <div>
            <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Justification</label>
            <input type="text" value={virementForm.justification || ''} onChange={e => setVirementForm((p: any) => ({...p, justification: e.target.value}))}
              className="w-full mt-1 h-9 rounded-lg border-2 border-slate-200 px-3 text-sm focus:outline-none focus:border-blue-500" />
          </div>
          <div>
            <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">TVA / Mois</label>
            <input type="month" value={virementForm.tva_mois || ''} onChange={e => setVirementForm((p: any) => ({...p, tva_mois: e.target.value}))}
              className="w-full mt-1 h-9 rounded-lg border-2 border-slate-200 px-3 text-sm focus:outline-none focus:border-blue-500" />
          </div>

          {/* Compte N° — select from RIP */}
          <div className="sm:col-span-2 lg:col-span-3">
            <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Compte débiteur (RIP société)</label>
            <select value={virementForm.compte_id || ''} onChange={e => {
              const acct = ripList.find((r: any) => r.id === e.target.value);
              setVirementForm((p: any) => ({...p, compte_id: e.target.value, banque: acct?.banque || '', agence: acct?.agence || '', numero_compte: acct?.numero_compte || '' }));
            }}
              className="w-full mt-1 h-9 rounded-lg border-2 border-amber-200 bg-amber-50 px-3 text-sm font-bold text-amber-800 focus:outline-none focus:border-amber-500">
              <option value="">— Sélectionner un compte —</option>
              {ripList.map((r: any) => (<option key={r.id} value={r.id}>{r.banque} — {r.agence} — {r.numero_compte}</option>))}
            </select>
          </div>

          {/* Auto-filled from account */}
          {virementForm.banque && (
            <div className="sm:col-span-2 lg:col-span-3 p-3 bg-amber-50 border border-amber-200 rounded-xl flex gap-6">
              <div><span className="text-[9px] font-black text-amber-600 uppercase block">Banque</span><span className="text-xs font-bold text-amber-800">{virementForm.banque}</span></div>
              <div><span className="text-[9px] font-black text-amber-600 uppercase block">Agence</span><span className="text-xs font-bold text-amber-800">{virementForm.agence}</span></div>
              <div><span className="text-[9px] font-black text-amber-600 uppercase block">N° Compte</span><span className="text-xs font-mono font-bold text-amber-800">{virementForm.numero_compte}</span></div>
            </div>
          )}
        </div>
        <div className="flex gap-3 pt-5">
          <button onClick={handleSaveVirement} className="flex-1 h-10 bg-blue-600 hover:bg-blue-700 text-white text-sm font-bold rounded-lg cursor-pointer">{editingVirement ? 'Enregistrer' : 'Ajouter'}</button>
          <button onClick={() => { setShowVirementForm(false); setEditingVirement(null); setVirementForm(emptyVirementForm); }} className="flex-1 h-10 bg-slate-100 hover:bg-slate-200 text-slate-700 text-sm font-bold rounded-lg cursor-pointer">Annuler</button>
        </div>
      </motion.div>
    </div>
  )}
</AnimatePresence>
{/* Relevé Form */}
<AnimatePresence>
  {showReleveForm && (
    <div className="fixed inset-0 bg-black/50 z-50 flex items-center justify-center p-4">
      <motion.div initial={{ scale: 0.95, opacity: 0 }} animate={{ scale: 1, opacity: 1 }} exit={{ scale: 0.95, opacity: 0 }}
        className="bg-white rounded-xl p-6 max-w-3xl w-full shadow-xl max-h-[90vh] overflow-y-auto">
        <div className="flex items-center justify-between mb-5">
          <h3 className="font-black text-slate-900 uppercase tracking-widest text-sm">{editingReleve ? 'Modifier Transaction' : 'Nouvelle Transaction'}</h3>
          <button onClick={() => { setShowReleveForm(false); setEditingReleve(null); setReleveForm(emptyReleveForm); }} className="p-1.5 rounded hover:bg-slate-100 text-slate-400"><X size={16} /></button>
        </div>
        <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-4">
          {[
            { label: 'Code', key: 'code' },
            { label: 'Date Opération', key: 'date_operation' },
            { label: 'Date Valeur', key: 'date_valeur' },
          ].map(({label, key}) => (
            <div key={key}><label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">{label}</label>
              <input type="text" value={releveForm[key]||''} onChange={e => setReleveForm((p:any)=>({...p,[key]:e.target.value}))}
                className="w-full mt-1 h-9 rounded-lg border-2 border-slate-200 px-3 text-sm focus:outline-none focus:border-blue-500" /></div>
          ))}
          <div className="sm:col-span-2"><label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Libellé</label>
            <input type="text" value={releveForm.libelle||''} onChange={e => setReleveForm((p:any)=>({...p,libelle:e.target.value,category:categorizeTransaction(e.target.value)}))}
              className="w-full mt-1 h-9 rounded-lg border-2 border-slate-200 px-3 text-sm focus:outline-none focus:border-blue-500" /></div>
          <div><label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Débit</label>
            <input type="number" value={releveForm.debit||''} onChange={e => setReleveForm((p:any)=>({...p,debit:e.target.value}))}
              className="w-full mt-1 h-9 rounded-lg border-2 border-rose-200 bg-rose-50 px-3 text-sm focus:outline-none focus:border-rose-500" /></div>
          <div><label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Crédit</label>
            <input type="number" value={releveForm.credit||''} onChange={e => setReleveForm((p:any)=>({...p,credit:e.target.value}))}
              className="w-full mt-1 h-9 rounded-lg border-2 border-emerald-200 bg-emerald-50 px-3 text-sm focus:outline-none focus:border-emerald-500" /></div>
          <div><label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Catégorie</label>
            <select value={releveForm.category||'virement'} onChange={e => setReleveForm((p:any)=>({...p,category:e.target.value}))}
              className="w-full mt-1 h-9 rounded-lg border-2 border-slate-200 px-3 text-sm focus:outline-none focus:border-blue-500">
              {ETAT_CATEGORIES.map(c => (<option key={c.id} value={c.id}>{c.label}</option>))}
            </select></div>
          <div><label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Destination</label>
            <input type="text" value={releveForm.destination||''} onChange={e => setReleveForm((p:any)=>({...p,destination:e.target.value}))}
              className="w-full mt-1 h-9 rounded-lg border-2 border-slate-200 px-3 text-sm focus:outline-none focus:border-blue-500" /></div>
          <div><label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Note Opération</label>
            <input type="text" value={releveForm.note_operation||''} onChange={e => setReleveForm((p:any)=>({...p,note_operation:e.target.value}))}
              className="w-full mt-1 h-9 rounded-lg border-2 border-slate-200 px-3 text-sm focus:outline-none focus:border-blue-500" /></div>
              <div><label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Code Règlement</label>
            <input type="text" value={releveForm.code_reglement || ''} onChange={e => setReleveForm((p: any) => ({...p, code_reglement: e.target.value}))}
              className="w-full mt-1 h-9 rounded-lg border-2 border-slate-200 px-3 text-sm focus:outline-none focus:border-blue-500" /></div>
          <div><label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Réf. Règlement</label>
            <input type="text" value={releveForm.ref_reglement||''} onChange={e => setReleveForm((p:any)=>({...p,ref_reglement:e.target.value}))}
              className="w-full mt-1 h-9 rounded-lg border-2 border-slate-200 px-3 text-sm focus:outline-none focus:border-blue-500" /></div>
          <div className="sm:col-span-2"><label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Observation</label>
            <input type="text" value={releveForm.observation||''} onChange={e => setReleveForm((p:any)=>({...p,observation:e.target.value}))}
              className="w-full mt-1 h-9 rounded-lg border-2 border-slate-200 px-3 text-sm focus:outline-none focus:border-blue-500" /></div>
              <div className="sm:col-span-2 lg:col-span-3">
            <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Code Règlement (auto-remplit Destination + Note)</label>
            <div className="flex gap-2 mt-1">
              <input type="text" value={releveForm.code_reglement || ''} placeholder="Tapez le code..."
                onChange={e => setReleveForm((p: any) => ({ ...p, code_reglement: e.target.value }))}
                className="flex-1 h-9 rounded-lg border-2 border-amber-200 bg-amber-50 px-3 text-sm font-mono font-bold text-amber-800 focus:outline-none focus:border-amber-500" />
              <button type="button" onClick={async () => {
                const code = releveForm.code_reglement?.trim();
                if (!code) { toast.error("Saisissez un code."); return; }
                // Lookup in reglements
                const { data: regData } = await supabase.from('reglements').select('*').eq('company_id', companyId).eq('code_reglement', code).limit(1);
                if (regData && regData.length > 0) {
                  const reg = regData[0];
                  const factNums = (reg.facture_numbers || []).join(', ');
                  setReleveForm((p: any) => ({
                    ...p,
                    destination: reg.client || p.destination || '',
                    note_operation: factNums ? `Facture(s): ${factNums}` : p.note_operation || '',
                    ref_reglement: reg.numero || reg.reference_virement || p.ref_reglement || '',
                  }));
                  toast.success(`Trouvé: ${reg.client} — ${reg.type_reglement}`);
                  return;
                }
                // Lookup in purchases
                const { data: purchData } = await supabase.from('purchases').select('*').eq('company_id', companyId).eq('code_reglement', code);
                if (purchData && purchData.length > 0) {
                  const fournisseur = purchData[0].fournisseur || purchData[0].supplier || '';
                  const details = purchData.map((p: any) => p.numero_facture || p.invoice_number || '').filter(Boolean).join(', ');
                  setReleveForm((p: any) => ({
                    ...p,
                    destination: fournisseur || p.destination || '',
                    note_operation: details || p.note_operation || '',
                  }));
                  toast.success(`Trouvé: ${fournisseur} — ${purchData.length} achat(s)`);
                  return;
                }
                toast.error("Aucun règlement ou achat trouvé avec ce code.");
              }}
                className="h-9 px-4 bg-amber-600 hover:bg-amber-700 text-white text-xs font-black rounded-lg cursor-pointer whitespace-nowrap">
                Rechercher
              </button>
            </div>
          </div>
        </div>
        <div className="flex gap-3 pt-5">
          <button onClick={handleSaveReleve} className="flex-1 h-10 bg-blue-600 hover:bg-blue-700 text-white text-sm font-bold rounded-lg cursor-pointer">{editingReleve?'Enregistrer':'Ajouter'}</button>
          <button onClick={() => { setShowReleveForm(false); setEditingReleve(null); setReleveForm(emptyReleveForm); }} className="flex-1 h-10 bg-slate-100 hover:bg-slate-200 text-slate-700 text-sm font-bold rounded-lg cursor-pointer">Annuler</button>
        </div>
      </motion.div>
    </div>
  )}
</AnimatePresence>

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
                { label: 'Code Règlement', key: 'code_reglement', type: 'text' },
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
{/* BC Form Modal */}
<AnimatePresence>
  {showBCForm && (
    <div className="fixed inset-0 bg-black/50 z-50 flex items-center justify-center p-4">
      <motion.div initial={{ scale: 0.95, opacity: 0 }} animate={{ scale: 1, opacity: 1 }}
        exit={{ scale: 0.95, opacity: 0 }}
        className="bg-white rounded-xl p-6 max-w-3xl w-full shadow-xl max-h-[90vh] overflow-y-auto">
        <div className="flex items-center justify-between mb-5">
          <h3 className="font-black text-slate-900 uppercase tracking-widest text-sm">{editingBC ? 'Modifier le BC' : 'Nouveau Bon de Commande'}</h3>
          <button onClick={() => { setShowBCForm(false); setEditingBC(null); setBcForm(emptyBCForm); }} className="p-1.5 rounded hover:bg-slate-100 text-slate-400"><X size={16} /></button>
        </div>
        <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-4">
          {[
            { label: 'Date', key: 'date', type: 'date' },
            { label: 'N° BC', key: 'numero_bc', type: 'text' },
            { label: 'Personne à contacter', key: 'personne_contact', type: 'text' },
            { label: 'Référence', key: 'reference', type: 'text' },
            { label: 'Désignation', key: 'designation', type: 'text' },
          ].map(({ label, key, type }) => (
            <div key={key}>
              <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">{label}</label>
              <input type={type} value={bcForm[key] || ''} onChange={e => setBcForm((p: any) => ({ ...p, [key]: e.target.value }))}
                className="w-full mt-1 h-9 rounded-lg border-2 border-slate-200 px-3 text-sm focus:outline-none focus:border-blue-500" />
            </div>
          ))}
          <div>
            <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Fournisseur</label>
            <select value={bcForm.fournisseur || ''} onChange={e => setBcForm((p: any) => ({ ...p, fournisseur: e.target.value }))}
              className="w-full mt-1 h-9 rounded-lg border-2 border-slate-200 px-3 text-sm focus:outline-none focus:border-blue-500">
              <option value="">— Sélectionner —</option>
              {fournisseursList.map((f: any) => (<option key={f.id} value={f.nom}>{f.nom}</option>))}
            </select>
          </div>
          <div>
            <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Quantité</label>
            <input type="number" value={bcForm.quantite || '1'} onChange={e => setBcForm((p: any) => ({ ...p, quantite: e.target.value }))}
              className="w-full mt-1 h-9 rounded-lg border-2 border-slate-200 px-3 text-sm focus:outline-none focus:border-blue-500" />
          </div>
          <div>
            <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Prix unitaire HT</label>
            <input type="number" value={bcForm.prix_unitaire_ht || ''} onChange={e => {
                const ht = parseFloat(e.target.value)||0; const qty = parseFloat(bcForm.quantite)||1; const rate = parseFloat(bcForm.tva_rate)||0;
                const m = ht*qty; const tva = parseFloat((m*rate/100).toFixed(2));
                setBcForm((p: any) => ({...p, prix_unitaire_ht: e.target.value, montant_ht: String(m.toFixed(2)), tva_amount: String(tva), montant_ttc: String((m+tva).toFixed(2))}));
              }}
              className="w-full mt-1 h-9 rounded-lg border-2 border-slate-200 px-3 text-sm focus:outline-none focus:border-blue-500" />
          </div>
          <div>
            <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">TVA (%)</label>
            <input type="number" value={bcForm.tva_rate || ''} onChange={e => {
                const rate = parseFloat(e.target.value)||0; const ht = parseFloat(bcForm.prix_unitaire_ht)||0; const qty = parseFloat(bcForm.quantite)||1;
                const m = ht*qty; const tva = parseFloat((m*rate/100).toFixed(2));
                setBcForm((p: any) => ({...p, tva_rate: e.target.value, tva_amount: String(tva), montant_ttc: String((m+tva).toFixed(2))}));
              }}
              className="w-full mt-1 h-9 rounded-lg border-2 border-slate-200 px-3 text-sm focus:outline-none focus:border-blue-500" />
          </div>
          <div>
            <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Montant TTC (auto)</label>
            <input type="number" value={bcForm.montant_ttc || ''} readOnly className="w-full mt-1 h-9 rounded-lg border-2 border-blue-100 bg-blue-50 px-3 text-sm font-bold text-blue-700 cursor-not-allowed" />
          </div>
          <div>
            <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Statut</label>
            <select value={bcForm.statut || 'en_attente'} onChange={e => setBcForm((p: any) => ({...p, statut: e.target.value}))}
              className="w-full mt-1 h-9 rounded-lg border-2 border-slate-200 px-3 text-sm focus:outline-none focus:border-blue-500">
              <option value="en_attente">En attente</option><option value="validé">Validé</option><option value="livré">Livré</option><option value="annulé">Annulé</option>
            </select>
          </div>
          <div className="sm:col-span-2">
            <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Observation</label>
            <input type="text" value={bcForm.observation || ''} onChange={e => setBcForm((p: any) => ({...p, observation: e.target.value}))}
              className="w-full mt-1 h-9 rounded-lg border-2 border-slate-200 px-3 text-sm focus:outline-none focus:border-blue-500" />
          </div>
        </div>
        <div className="flex gap-3 pt-5">
          <button onClick={handleSaveBC} className="flex-1 h-10 bg-blue-600 hover:bg-blue-700 text-white text-sm font-bold rounded-lg cursor-pointer">{editingBC ? 'Enregistrer' : 'Ajouter le BC'}</button>
          <button onClick={() => { setShowBCForm(false); setEditingBC(null); setBcForm(emptyBCForm); }} className="flex-1 h-10 bg-slate-100 hover:bg-slate-200 text-slate-700 text-sm font-bold rounded-lg cursor-pointer">Annuler</button>
        </div>
      </motion.div>
    </div>
  )}
</AnimatePresence>
{/* Devis Form Modal */}
<AnimatePresence>
  {showDevisForm && (
    <div className="fixed inset-0 bg-black/50 z-50 flex items-center justify-center p-4">
      <motion.div initial={{ scale: 0.95, opacity: 0 }} animate={{ scale: 1, opacity: 1 }}
        exit={{ scale: 0.95, opacity: 0 }}
        className="bg-white rounded-xl p-6 max-w-3xl w-full shadow-xl max-h-[90vh] overflow-y-auto">
        <div className="flex items-center justify-between mb-5">
          <h3 className="font-black text-slate-900 uppercase tracking-widest text-sm">{editingDevis ? 'Modifier le Devis' : 'Nouveau Devis'}</h3>
          <button onClick={() => { setShowDevisForm(false); setEditingDevis(null); setDevisForm(emptyDevisForm); }} className="p-1.5 rounded hover:bg-slate-100 text-slate-400"><X size={16} /></button>
        </div>
        <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-4">
          {[
            { label: 'Date', key: 'date', type: 'date' },
            { label: 'N° Devis', key: 'numero_devis', type: 'text' },
            { label: 'Personne à contacter', key: 'personne_contact', type: 'text' },
            { label: 'Départ', key: 'depart', type: 'text' },
            { label: 'Arrivée', key: 'arrivee', type: 'text' },
            { label: 'Type véhicule', key: 'type_vehicule', type: 'text' },
            { label: 'Désignation', key: 'designation', type: 'text' },
          ].map(({ label, key, type }) => (
            <div key={key}>
              <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">{label}</label>
              <input type={type} value={devisForm[key] || ''} onChange={e => setDevisForm((p: any) => ({ ...p, [key]: e.target.value }))}
                className="w-full mt-1 h-9 rounded-lg border-2 border-slate-200 px-3 text-sm focus:outline-none focus:border-blue-500" />
            </div>
          ))}
          {/* Client dropdown */}
          <div>
            <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Client</label>
            <select value={devisForm.client || ''} onChange={e => setDevisForm((p: any) => ({ ...p, client: e.target.value }))}
              className="w-full mt-1 h-9 rounded-lg border-2 border-slate-200 px-3 text-sm focus:outline-none focus:border-blue-500">
              <option value="">— Sélectionner —</option>
              {clientsList.map((c: any) => (<option key={c.id} value={c.nom}>{c.nom}</option>))}
            </select>
          </div>
          <div>
            <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Quantité</label>
            <input type="number" value={devisForm.quantite || '1'} onChange={e => setDevisForm((p: any) => ({ ...p, quantite: e.target.value }))}
              className="w-full mt-1 h-9 rounded-lg border-2 border-slate-200 px-3 text-sm focus:outline-none focus:border-blue-500" />
          </div>
          <div>
            <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Prix unitaire HT</label>
            <input type="number" value={devisForm.prix_unitaire_ht || ''} placeholder="0"
              onChange={e => {
                const ht = parseFloat(e.target.value) || 0;
                const qty = parseFloat(devisForm.quantite) || 1;
                const rate = parseFloat(devisForm.tva_rate) || 0;
                const montant = ht * qty;
                const tva = parseFloat((montant * rate / 100).toFixed(2));
                setDevisForm((p: any) => ({ ...p, prix_unitaire_ht: e.target.value, montant_ht: String(montant.toFixed(2)), tva_amount: String(tva), montant_ttc: String((montant + tva).toFixed(2)) }));
              }}
              className="w-full mt-1 h-9 rounded-lg border-2 border-slate-200 px-3 text-sm focus:outline-none focus:border-blue-500" />
          </div>
          <div>
            <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">TVA (%)</label>
            <input type="number" value={devisForm.tva_rate || ''} placeholder="20"
              onChange={e => {
                const rate = parseFloat(e.target.value) || 0;
                const ht = parseFloat(devisForm.prix_unitaire_ht) || 0;
                const qty = parseFloat(devisForm.quantite) || 1;
                const montant = ht * qty;
                const tva = parseFloat((montant * rate / 100).toFixed(2));
                setDevisForm((p: any) => ({ ...p, tva_rate: e.target.value, tva_amount: String(tva), montant_ttc: String((montant + tva).toFixed(2)) }));
              }}
              className="w-full mt-1 h-9 rounded-lg border-2 border-slate-200 px-3 text-sm focus:outline-none focus:border-blue-500" />
          </div>
          <div>
            <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Montant TTC (auto)</label>
            <input type="number" value={devisForm.montant_ttc || ''} readOnly
              className="w-full mt-1 h-9 rounded-lg border-2 border-blue-100 bg-blue-50 px-3 text-sm font-bold text-blue-700 cursor-not-allowed" />
          </div>
          <div>
            <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Statut</label>
            <select value={devisForm.statut || 'en_attente'} onChange={e => setDevisForm((p: any) => ({ ...p, statut: e.target.value }))}
              className="w-full mt-1 h-9 rounded-lg border-2 border-slate-200 px-3 text-sm focus:outline-none focus:border-blue-500">
              <option value="en_attente">En attente</option><option value="accepté">Accepté</option><option value="refusé">Refusé</option>
            </select>
          </div>
          <div className="sm:col-span-2">
            <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Observation</label>
            <input type="text" value={devisForm.observation || ''} onChange={e => setDevisForm((p: any) => ({ ...p, observation: e.target.value }))}
              className="w-full mt-1 h-9 rounded-lg border-2 border-slate-200 px-3 text-sm focus:outline-none focus:border-blue-500" />
          </div>
        </div>
        <div className="flex gap-3 pt-5">
          <button onClick={handleSaveDevis} className="flex-1 h-10 bg-blue-600 hover:bg-blue-700 text-white text-sm font-bold rounded-lg cursor-pointer">
            {editingDevis ? 'Enregistrer' : 'Ajouter le devis'}
          </button>
          <button onClick={() => { setShowDevisForm(false); setEditingDevis(null); setDevisForm(emptyDevisForm); }}
            className="flex-1 h-10 bg-slate-100 hover:bg-slate-200 text-slate-700 text-sm font-bold rounded-lg cursor-pointer">Annuler</button>
        </div>
      </motion.div>
    </div>
  )}
</AnimatePresence>
{/* Règlement Form Modal */}
<AnimatePresence>
  {showReglementForm && (
    <div className="fixed inset-0 bg-black/50 z-50 flex items-center justify-center p-4">
      <motion.div initial={{ scale: 0.95, opacity: 0 }} animate={{ scale: 1, opacity: 1 }}
        exit={{ scale: 0.95, opacity: 0 }}
        className="bg-white rounded-xl p-6 max-w-2xl w-full shadow-xl max-h-[90vh] overflow-y-auto">
        <div className="flex items-center justify-between mb-5">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 rounded-lg bg-emerald-600 flex items-center justify-center text-white"><Check size={20} /></div>
            <div>
              <h3 className="font-black text-slate-900 uppercase tracking-widest text-sm">Règlement</h3>
              <p className="text-[10px] text-slate-500 mt-0.5">{selectedFacts.length} facture(s) sélectionnée(s) — {facturationList.filter((f:any) => selectedFacts.includes(f.id)).reduce((s:number,f:any) => s+(parseFloat(f.montant_ttc)||0), 0).toLocaleString('fr-MA',{minimumFractionDigits:2})} MAD</p>
            </div>
          </div>
          <button onClick={() => setShowReglementForm(false)} className="p-1.5 rounded hover:bg-slate-100 text-slate-400"><X size={16} /></button>
        </div>

        <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
          {/* Type de règlement */}
          <div className="sm:col-span-2">
            <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Type de règlement</label>
            <div className="flex gap-2 mt-1">
              {[
                { value: 'cheque', label: 'Chèque' },
                { value: 'effet', label: 'Effet' },
                { value: 'virement', label: 'Virement' },
                { value: 'espece', label: 'Espèce' },
                { value: 'compensation', label: 'Compensation' },
              ].map(t => (
                <button key={t.value} onClick={() => setReglementForm((p: any) => ({ ...p, type_reglement: t.value }))}
                  className={`flex-1 h-9 rounded-lg text-xs font-black uppercase tracking-wider cursor-pointer border-2 transition-all ${reglementForm.type_reglement === t.value ? 'bg-emerald-600 text-white border-emerald-600' : 'bg-white text-slate-600 border-slate-200 hover:border-emerald-300'}`}>
                  {t.label}
                </button>
              ))}
            </div>
          </div>

          {/* Date */}
          <div>
            <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Date de règlement</label>
            <input type="date" value={reglementForm.date_reglement}
              onChange={e => setReglementForm((p: any) => ({ ...p, date_reglement: e.target.value }))}
              className="w-full mt-1 h-9 rounded-lg border-2 border-slate-200 px-3 text-sm focus:outline-none focus:border-emerald-500" />
          </div>
          {/* Client — auto-filled */}
          <div>
            <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Client (auto)</label>
            <input type="text" readOnly value={(() => {
              const sel = facturationList.filter((f:any) => selectedFacts.includes(f.id));
              const clients = [...new Set(sel.map((f:any) => f.client).filter(Boolean))];
              return clients.join(', ');
            })()}
              className="w-full mt-1 h-9 rounded-lg border-2 border-slate-100 bg-slate-50 px-3 text-sm font-bold text-slate-700 cursor-not-allowed" />
          </div>

          {/* TVA Mois */}
          <div>
            <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">TVA Mois</label>
            <input type="month" value={reglementForm.tva_mois || ''}
              onChange={e => setReglementForm((p: any) => ({ ...p, tva_mois: e.target.value }))}
              className="w-full mt-1 h-9 rounded-lg border-2 border-slate-200 px-3 text-sm focus:outline-none focus:border-emerald-500" />
          </div>

          {/* Chèque / Effet fields */}
          {(reglementForm.type_reglement === 'cheque' || reglementForm.type_reglement === 'effet') && (
            <>
              <div>
                <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">
                  N° {reglementForm.type_reglement === 'cheque' ? 'Chèque' : 'Effet'}
                </label>
                <input type="text" value={reglementForm.numero} placeholder="N°..."
                  onChange={e => setReglementForm((p: any) => ({ ...p, numero: e.target.value }))}
                  className="w-full mt-1 h-9 rounded-lg border-2 border-slate-200 px-3 text-sm focus:outline-none focus:border-emerald-500" />
              </div>
              <div>
                <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Banque</label>
                <input type="text" value={reglementForm.banque} placeholder="Attijariwafa, BMCE..."
                  onChange={e => setReglementForm((p: any) => ({ ...p, banque: e.target.value }))}
                  className="w-full mt-1 h-9 rounded-lg border-2 border-slate-200 px-3 text-sm focus:outline-none focus:border-emerald-500" />
              </div>
            </>
          )}

          {/* Effet: date d'échéance */}
          {reglementForm.type_reglement === 'effet' && (
            <div>
              <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Date d'échéance</label>
              <input type="date" value={reglementForm.date_echeance}
                onChange={e => setReglementForm((p: any) => ({ ...p, date_echeance: e.target.value }))}
                className="w-full mt-1 h-9 rounded-lg border-2 border-slate-200 px-3 text-sm focus:outline-none focus:border-emerald-500" />
            </div>
          )}

          {/* Virement fields */}
          {reglementForm.type_reglement === 'virement' && (
            <>
              <div>
                <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Référence virement</label>
                <input type="text" value={reglementForm.reference_virement}
                  onChange={e => setReglementForm((p: any) => ({ ...p, reference_virement: e.target.value }))}
                  className="w-full mt-1 h-9 rounded-lg border-2 border-slate-200 px-3 text-sm focus:outline-none focus:border-emerald-500" />
              </div>
              <div>
                <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Banque</label>
                <input type="text" value={reglementForm.banque}
                  onChange={e => setReglementForm((p: any) => ({ ...p, banque: e.target.value }))}
                  className="w-full mt-1 h-9 rounded-lg border-2 border-slate-200 px-3 text-sm focus:outline-none focus:border-emerald-500" />
              </div>
            </>
          )}

          {/* Espèce fields */}
          {reglementForm.type_reglement === 'espece' && (
            <div>
              <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Reçu par</label>
              <input type="text" value={reglementForm.recu_par}
                onChange={e => setReglementForm((p: any) => ({ ...p, recu_par: e.target.value }))}
                className="w-full mt-1 h-9 rounded-lg border-2 border-slate-200 px-3 text-sm focus:outline-none focus:border-emerald-500" />
            </div>
          )}
          {/* Code Règlement */}
          <div>
            <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Code Règlement</label>
            <input type="text" value={reglementForm.code_reglement || ''} placeholder="REG-001"
              onChange={e => setReglementForm((p: any) => ({ ...p, code_reglement: e.target.value }))}
              className="w-full mt-1 h-9 rounded-lg border-2 border-slate-200 px-3 text-sm focus:outline-none focus:border-emerald-500" />
          </div>

          {/* Observation */}
          <div className="sm:col-span-2">
            <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Observation</label>
            <input type="text" value={reglementForm.observation}
              onChange={e => setReglementForm((p: any) => ({ ...p, observation: e.target.value }))}
              className="w-full mt-1 h-9 rounded-lg border-2 border-slate-200 px-3 text-sm focus:outline-none focus:border-emerald-500" />
          </div>

          {/* Scan upload */}
          <div className="sm:col-span-2">
            <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Scan du règlement (image ou PDF)</label>
            <input type="file" accept="image/*,.pdf"
              onChange={e => setReglementForm((p: any) => ({ ...p, scanFile: e.target.files?.[0] || null }))}
              className="w-full mt-1 text-sm text-slate-600 file:mr-3 file:py-1.5 file:px-3 file:rounded-lg file:border-0 file:text-xs file:font-bold file:bg-emerald-50 file:text-emerald-700 hover:file:bg-emerald-100" />
            {reglementForm.scanFile && (
              <p className="mt-1 text-[10px] text-emerald-600 font-bold">{reglementForm.scanFile.name}</p>
            )}
          </div>
        </div>

        {/* Factures being paid */}
        <div className="mt-4 p-3 bg-emerald-50 border border-emerald-200 rounded-xl">
          <p className="text-[10px] font-black text-emerald-800 uppercase tracking-widest mb-2">Factures à régler :</p>
          <div className="flex flex-wrap gap-1.5">
            {facturationList.filter((f:any) => selectedFacts.includes(f.id)).map((f:any) => (
              <span key={f.id} className="text-[10px] bg-white border border-emerald-200 text-emerald-800 font-bold px-2 py-0.5 rounded">
                {f.numero_facture || '—'} — {f.client} — {Number(f.montant_ttc||0).toLocaleString('fr-MA',{minimumFractionDigits:2})} MAD
              </span>
            ))}
          </div>
        </div>

        <div className="flex gap-3 pt-5">
          <button onClick={handleSaveReglement} disabled={savingReglement}
            className="flex-1 h-10 bg-emerald-600 hover:bg-emerald-700 disabled:opacity-60 text-white text-sm font-black rounded-lg cursor-pointer flex items-center justify-center gap-2">
            {savingReglement ? <Loader2 size={14} className="animate-spin" /> : <Check size={14} />}
            {savingReglement ? 'Enregistrement...' : 'Régler les factures'}
          </button>
          <button onClick={() => setShowReglementForm(false)}
            className="flex-1 h-10 bg-slate-100 hover:bg-slate-200 text-slate-700 text-sm font-bold rounded-lg cursor-pointer">Annuler</button>
        </div>
      </motion.div>
    </div>
  )}
</AnimatePresence>
{/* Edit Règlement Modal */}
<AnimatePresence>
  {showEditReglementForm && editingReglement && (
    <div className="fixed inset-0 bg-black/50 z-50 flex items-center justify-center p-4">
      <motion.div initial={{ scale: 0.95, opacity: 0 }} animate={{ scale: 1, opacity: 1 }}
        exit={{ scale: 0.95, opacity: 0 }}
        className="bg-white rounded-xl p-6 max-w-2xl w-full shadow-xl max-h-[90vh] overflow-y-auto">
        <div className="flex items-center justify-between mb-5">
          <h3 className="font-black text-slate-900 uppercase tracking-widest text-sm">Modifier le Règlement</h3>
          <button onClick={() => { setShowEditReglementForm(false); setEditingReglement(null); }} className="p-1.5 rounded hover:bg-slate-100 text-slate-400"><X size={16} /></button>
        </div>
        <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
          <div>
            <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Type</label>
            <select value={editingReglement.type_reglement} onChange={e => setEditingReglement((p: any) => ({...p, type_reglement: e.target.value}))}
              className="w-full mt-1 h-9 rounded-lg border-2 border-slate-200 px-3 text-sm focus:outline-none focus:border-blue-500">
              <option value="cheque">Chèque</option><option value="effet">Effet</option><option value="virement">Virement</option><option value="espece">Espèce</option><option value="compensation">Compensation</option>
            </select>
          </div>
          <div>
            <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Date</label>
            <input type="date" value={editingReglement.date_reglement || ''} onChange={e => setEditingReglement((p: any) => ({...p, date_reglement: e.target.value}))}
              className="w-full mt-1 h-9 rounded-lg border-2 border-slate-200 px-3 text-sm focus:outline-none focus:border-blue-500" />
          </div>
          <div>
            <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">N°</label>
            <input type="text" value={editingReglement.numero || ''} onChange={e => setEditingReglement((p: any) => ({...p, numero: e.target.value}))}
              className="w-full mt-1 h-9 rounded-lg border-2 border-slate-200 px-3 text-sm focus:outline-none focus:border-blue-500" />
          </div>
          <div>
            <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Banque</label>
            <input type="text" value={editingReglement.banque || ''} onChange={e => setEditingReglement((p: any) => ({...p, banque: e.target.value}))}
              className="w-full mt-1 h-9 rounded-lg border-2 border-slate-200 px-3 text-sm focus:outline-none focus:border-blue-500" />
          </div>
          <div>
            <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Date d'échéance</label>
            <input type="date" value={editingReglement.date_echeance || ''} onChange={e => setEditingReglement((p: any) => ({...p, date_echeance: e.target.value}))}
              className="w-full mt-1 h-9 rounded-lg border-2 border-slate-200 px-3 text-sm focus:outline-none focus:border-blue-500" />
          </div>
          <div>
            <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">TVA Mois</label>
            <input type="month" value={editingReglement.tva_mois || ''} onChange={e => setEditingReglement((p: any) => ({...p, tva_mois: e.target.value}))}
              className="w-full mt-1 h-9 rounded-lg border-2 border-slate-200 px-3 text-sm focus:outline-none focus:border-blue-500" />
          </div>
          <div>
            <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Réf. Virement</label>
            <input type="text" value={editingReglement.reference_virement || ''} onChange={e => setEditingReglement((p: any) => ({...p, reference_virement: e.target.value}))}
              className="w-full mt-1 h-9 rounded-lg border-2 border-slate-200 px-3 text-sm focus:outline-none focus:border-blue-500" />
          </div>
          <div>
            <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Reçu par</label>
            <input type="text" value={editingReglement.recu_par || ''} onChange={e => setEditingReglement((p: any) => ({...p, recu_par: e.target.value}))}
              className="w-full mt-1 h-9 rounded-lg border-2 border-slate-200 px-3 text-sm focus:outline-none focus:border-blue-500" />
          </div>
          <div className="sm:col-span-2">
            <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Observation</label>
            <input type="text" value={editingReglement.observation || ''} onChange={e => setEditingReglement((p: any) => ({...p, observation: e.target.value}))}
              className="w-full mt-1 h-9 rounded-lg border-2 border-slate-200 px-3 text-sm focus:outline-none focus:border-blue-500" />
          </div>
        </div>
        <div className="flex gap-3 pt-5">
          <button onClick={handleUpdateReglement} className="flex-1 h-10 bg-blue-600 hover:bg-blue-700 text-white text-sm font-bold rounded-lg cursor-pointer">Enregistrer</button>
          <button onClick={() => { setShowEditReglementForm(false); setEditingReglement(null); }} className="flex-1 h-10 bg-slate-100 hover:bg-slate-200 text-slate-700 text-sm font-bold rounded-lg cursor-pointer">Annuler</button>
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
            <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Départ</label>
            <input type="text" value={avoirForm.depart} placeholder="Casablanca"
              onChange={e => setAvoirForm(p => ({ ...p, depart: e.target.value }))}
              className="w-full mt-1 h-9 rounded-lg border-2 border-slate-200 px-3 text-sm focus:outline-none focus:border-rose-500" />
          </div>
          <div>
            <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Arrivée</label>
            <input type="text" value={avoirForm.arrivee} placeholder="Agadir"
              onChange={e => setAvoirForm(p => ({ ...p, arrivee: e.target.value }))}
              className="w-full mt-1 h-9 rounded-lg border-2 border-slate-200 px-3 text-sm focus:outline-none focus:border-rose-500" />
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
              depart: avoirForm.depart || null,
              arrivee: avoirForm.arrivee || null,
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