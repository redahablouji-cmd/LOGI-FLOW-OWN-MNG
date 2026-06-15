import { useState, useEffect } from 'react';
import {  useRef, useCallback } from 'react';
import { Loader2, RefreshCw, Download, Truck, Copy, CopyPlus, TrendingUp, Check, CloudOff } from 'lucide-react';
import { supabase } from '../../lib/supabase';
import { toast } from 'sonner';

// ───────────────────────── Field definitions ─────────────────────────
const MONTHLY_FIELDS = [
  { key: 'ca_ht_mensuel',       label: 'CA HT Mensuel (MAD)' },
  { key: 'km_parcouru_mensuel', label: 'Km Parcouru Mensuel' },
  { key: 'km_trajet',           label: 'Km Trajet (estimation)' },
] as const;

const CHARGES_FIXES = [
  { key: 'cf_salaire_conducteur',  label: 'Salaire Conducteur' },
  { key: 'cf_chs_patronales_cnss', label: 'CHS Patronales CNSS' },
  { key: 'cf_amo',                 label: 'AMO' },
  { key: 'cf_cimr',                label: 'CIMR' },
  { key: 'cf_annuite',             label: 'Annuité' },
  { key: 'cf_provision',           label: 'Provision' },
  { key: 'cf_amortissement',       label: 'Amortissement' },
  { key: 'cf_assurance',           label: 'Assurance' },
  { key: 'cf_assurance_mses',      label: 'Assurance M/ses' },
  { key: 'cf_vignette',            label: 'Vignette' },
  { key: 'cf_taxe_essieu',         label: "Taxe à l'Essieu" },
  { key: 'cf_carte_crise',         label: 'Carte Crise' },
  { key: 'cf_visite_technique',    label: 'Visite Technique' },
  { key: 'cf_leasing_remorque',    label: 'Leasing Remorque' },
  { key: 'cf_leasing_plateau',     label: 'Leasing Plateau' },
  { key: 'cf_interet_emprunt',     label: "Intérêt de l'Emprunt" },
] as const;

const CHARGES_VARIABLES = [
  { key: 'cv_consommation_dhs',  label: 'Consommation en DHS' },
  { key: 'cv_lavage_graissage',  label: 'Lavage et Graissage' },
  { key: 'cv_gasoil_externes',   label: 'Gasoil Externes' },
  { key: 'cv_lubrifiant',        label: 'Lubrifiant' },
  { key: 'cv_pneumatique',       label: 'Pneumatique' },
  { key: 'cv_entretien',         label: 'Entretien' },
  { key: 'cv_manutention',       label: 'Manutention' },
  { key: 'cv_reparation',        label: 'Réparation' },
  { key: 'cv_peage',             label: 'Péage' },
  { key: 'cv_frais_deplacement', label: 'Frais de Déplacement' },
  { key: 'cv_parking',           label: 'Parking' },
  { key: 'cv_autre',             label: 'Autre' },
] as const;

const CHARGES_STRUCTURE = [
  { key: 'cs_salaire_bureau',        label: 'Salaire Bureau' },
  { key: 'cs_charges_personnels',    label: 'Charges Personnels' },
  { key: 'cs_amo',                   label: 'AMO' },
  { key: 'cs_cimr',                  label: 'CIMR' },
  { key: 'cs_fourniture_bureau',     label: 'Fourniture de Bureau' },
  { key: 'cs_materiel_bureau',       label: 'Matériel de Bureau' },
  { key: 'cs_essence_moto',          label: 'Essence Moto' },
  { key: 'cs_gasoil_cadeaux',        label: 'Gasoil Cadeaux' },
  { key: 'cs_gasoil_direction',      label: 'Gasoil Direction Actionnaire' },
  { key: 'cs_loyer',                 label: 'Loyer' },
  { key: 'cs_electricite',           label: 'Électricité' },
  { key: 'cs_mobile',                label: 'Mobile' },
  { key: 'cs_num_bouchaib',          label: 'Num Bouchaib' },
  { key: 'cs_internet',              label: 'Internet' },
  { key: 'cs_reparation_express',    label: 'Réparation Express' },
  { key: 'cs_pneumatique_express',   label: 'Pneumatique Express' },
  { key: 'cs_fax',                   label: 'Fax' },
  { key: 'cs_fix',                   label: 'Fix' },
  { key: 'cs_femme_menage',          label: 'Femme de Ménage' },
  { key: 'cs_comission',             label: 'Comission' },
  { key: 'cs_is',                    label: 'IS' },
  { key: 'cs_pattente',              label: 'Patente' },
  { key: 'cs_taxe_habitation',       label: "Taxe d'Habitation" },
  { key: 'cs_taxe_professionnelle',  label: 'Taxe Professionnelle' },
  { key: 'cs_caisse_bureau',         label: 'Caisse Bureau' },
  { key: 'cs_voiture_berlingo',      label: 'Voiture Berlingo' },
  { key: 'cs_assurance_berlingo',    label: 'Assurance Berlingo' },
  { key: 'cs_fiduciaire',            label: 'Fiduciaire' },
  { key: 'cs_entretien_bureau',      label: 'Entretien Bureau' },
  { key: 'cs_divers',                label: 'Divers' },
] as const;

const FIXES_KEYS     = CHARGES_FIXES.map(f => f.key);
const VARIABLES_KEYS = CHARGES_VARIABLES.map(f => f.key);
const STRUCTURE_KEYS = CHARGES_STRUCTURE.map(f => f.key);
const ALL_NUM_KEYS = [
  ...MONTHLY_FIELDS.map(f => f.key),
  ...FIXES_KEYS, ...VARIABLES_KEYS, ...STRUCTURE_KEYS,
];

const emptyRow = (matricule: string) => {
  const r: any = { matricule };
  ALL_NUM_KEYS.forEach(k => { r[k] = 0; });
  return r;
};

const num = (v: any) => parseFloat(v) || 0;
const fmt = (v: number) =>
  isFinite(v) ? v.toLocaleString('fr-MA', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) : '0,00';

// ───────────────────────── KPI calculator (your Excel formulas) ─────────────────────────
const computeKPIs = (row: any) => {
  const totalFixes     = FIXES_KEYS.reduce((s, k) => s + num(row[k]), 0);
  const totalVariables = VARIABLES_KEYS.reduce((s, k) => s + num(row[k]), 0);
  const totalStructure = STRUCTURE_KEYS.reduce((s, k) => s + num(row[k]), 0);

  const ca   = num(row.ca_ht_mensuel);
  const km   = num(row.km_parcouru_mensuel);
  const kmT  = num(row.km_trajet);

  const coutRevientMensuel = totalFixes + totalVariables + totalStructure;
  const coutRevientKm      = km > 0 ? coutRevientMensuel / km : 0;
  const benefice           = ca - coutRevientMensuel;
  const beneficeKm         = km > 0 ? benefice / km : 0;
  const margeBeneficiaire  = ca > 0 ? (benefice / ca) * 100 : 0;
  const prixVenteKm        = coutRevientKm + beneficeKm;
  const estimationPrixVente = prixVenteKm * kmT;

  return {
    totalFixes, totalVariables, totalStructure,
    coutRevientMensuel, coutRevientKm, benefice,
    beneficeKm, margeBeneficiaire, prixVenteKm, estimationPrixVente,
  };
};

interface Props {
  companyId: string;
}

export default function CoutRevientTab({ companyId }: Props) {
  const [saving, setSaving]   = useState(false);
  const [trucks, setTrucks]   = useState<string[]>([]);
  const [rows, setRows]       = useState<Record<string, any>>({});
  const [loading, setLoading] = useState(false);


  // Autosave status
  const [saveStatus, setSaveStatus] = useState<'idle' | 'saving' | 'saved' | 'error'>('idle');
  const saveTimer = useRef<ReturnType<typeof setTimeout> | null>(null);
  const rowsRef   = useRef(rows);
  const trucksRef = useRef(trucks);
  rowsRef.current   = rows;
  trucksRef.current = trucks;

  // Save all trucks silently in the background
  const saveNow = useCallback(async () => {
    if (!companyId) return;
    const currentTrucks = trucksRef.current;
    const currentRows   = rowsRef.current;
    if (!currentTrucks.length) return;
    setSaveStatus('saving');
    const payload = currentTrucks.map(t => {
      const r = currentRows[t] || {};
      const clean: any = { company_id: companyId, matricule: t, updated_at: new Date().toISOString() };
      ALL_NUM_KEYS.forEach(k => { clean[k] = num(r[k]); });
      return clean;
    });
    const { error } = await supabase
      .from('cout_revient')
      .upsert(payload, { onConflict: 'company_id,matricule' });
    if (error) setSaveStatus('error');
    else setSaveStatus('saved');
  }, [companyId]);

  // Debounce: trigger save ~900ms after the last change
  const scheduleSave = useCallback(() => {
    setSaveStatus('saving');
    if (saveTimer.current) clearTimeout(saveTimer.current);
    saveTimer.current = setTimeout(() => { saveNow(); }, 900);
  }, [saveNow]);

  // Flush pending save when component unmounts (e.g. user changes tab)
  useEffect(() => {
    return () => {
      if (saveTimer.current) clearTimeout(saveTimer.current);
      saveNow();
    };
  }, [saveNow]);
  
  const [activeTruck, setActiveTruck] = useState<string>('');

  // ── Load trucks + saved cost rows ──
  const fetchAll = async () => {
    if (!companyId) return;
    setLoading(true);

    const [{ data: drivers }, { data: saved }] = await Promise.all([
      supabase.from('fleet_drivers').select('immatriculation').eq('company_id', companyId),
      supabase.from('cout_revient').select('*').eq('company_id', companyId),
    ]);

    const plates: string[] = Array.from(
      new Set((drivers || []).map((d: any) => (d.immatriculation || '').trim()).filter(Boolean))
    );

    const savedMap: Record<string, any> = {};
    (saved || []).forEach(s => { savedMap[s.matricule] = s; });

    const merged: Record<string, any> = {};
    plates.forEach(p => { merged[p] = savedMap[p] ? { ...savedMap[p] } : emptyRow(p); });
    // keep any saved rows whose truck was removed from fleet_drivers
    (saved || []).forEach(s => { if (!merged[s.matricule]) { plates.push(s.matricule); merged[s.matricule] = { ...s }; } });

    setTrucks(plates);
    setRows(merged);
    if (plates.length && !plates.includes(activeTruck)) setActiveTruck(plates[0]);
    setLoading(false);
  };

  useEffect(() => { fetchAll(); /* eslint-disable-next-line */ }, [companyId]);

  // ── Update a single field on one truck ──
  const setField = (matricule: string, key: string, value: string) => {
    setRows(prev => ({ ...prev, [matricule]: { ...prev[matricule], [key]: value } }));
    scheduleSave();
  };

  // ── Copy one field's value into every other truck ──
  // For structure fields this is the "shared overhead" behaviour (always identical across trucks).
  const copyFieldToAll = (key: string, sourceMatricule: string) => {
    const value = rows[sourceMatricule]?.[key] ?? 0;
    setRows(prev => {
      const next = { ...prev };
      trucks.forEach(t => { next[t] = { ...next[t], [key]: value }; });
      return next;
    });
   toast.success(`Valeur copiée sur ${trucks.length} camion(s).`);
    scheduleSave();
  };

  // ── Copy an ENTIRE section (Fixes / Variables / Structure / Monthly) to all trucks ──
  const copySectionToAll = (keys: readonly string[], sourceMatricule: string, label: string) => {
    setRows(prev => {
      const next = { ...prev };
      const src = prev[sourceMatricule];
      trucks.forEach(t => {
        const patch: any = { ...next[t] };
        keys.forEach(k => { patch[k] = src[k] ?? 0; });
        next[t] = patch;
      });
      return next;
    });
    toast.success(`${label} copié sur tous les camions.`);
    scheduleSave();
  };

  // ── Save everything ──
  const handleSave = async () => {
    if (!companyId) return;
    setSaving(true);
    const payload = trucks.map(t => {
      const r = rows[t];
      const clean: any = { company_id: companyId, matricule: t, updated_at: new Date().toISOString() };
      ALL_NUM_KEYS.forEach(k => { clean[k] = num(r[k]); });
      return clean;
    });
    const { error } = await supabase
      .from('cout_revient')
      .upsert(payload, { onConflict: 'company_id,matricule' });
    if (!error) { toast.success('Coûts de revient sauvegardés.'); fetchAll(); }
    else toast.error(`Erreur: ${error.message}`);
    setSaving(false);
  };

  // ── Export all trucks' KPIs to XLS ──
  const handleExport = () => {
    const data = trucks.map(t => {
      const r = rows[t];
      const k = computeKPIs(r);
      return {
        'Matricule': t,
        'CA HT Mensuel': num(r.ca_ht_mensuel),
        'Km Parcouru': num(r.km_parcouru_mensuel),
        'Km Trajet': num(r.km_trajet),
        'Total Charges Fixes': k.totalFixes,
        'Total Charges Variables': k.totalVariables,
        'Total Charges Structure': k.totalStructure,
        'Coût Revient Mensuel': k.coutRevientMensuel,
        'Coût Revient Km': k.coutRevientKm,
        'Bénéfice': k.benefice,
        'Bénéfice/Km': k.beneficeKm,
        'Marge Bénéficiaire %': k.margeBeneficiaire,
        'Prix Vente Km': k.prixVenteKm,
        'Estimation Prix Vente': k.estimationPrixVente,
      };
    });
    if (!data.length) return;
    const headers = Object.keys(data[0]);
    const lines = data.map(row => headers.map(h => (row as any)[h] ?? '').join('\t'));
    const blob = new Blob([[headers.join('\t'), ...lines].join('\n')], { type: 'application/vnd.ms-excel' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = 'cout_de_revient.xls';
    a.click();
  };

  const current = activeTruck ? rows[activeTruck] : null;
  const kpi = current ? computeKPIs(current) : null;

  // ───────────────────────── Render ─────────────────────────
  return (
    <div>
      {/* Header */}
      <div className="mb-6 bg-slate-900 text-white rounded-xl p-6 border border-slate-800">
        <div className="flex items-center justify-between flex-wrap gap-3">
          <div>
            <span className="inline-flex items-center gap-1.5 px-2 py-0.5 rounded text-[10px] font-extrabold uppercase tracking-widest bg-teal-500 text-slate-950 mb-2">
              <TrendingUp className="w-3.5 h-3.5" /> Coût de Revient
            </span>
            <h1 className="text-2xl font-extrabold tracking-tight">Calcul de Rentabilité par Camion</h1>
            <p className="text-sm text-slate-400 mt-1">{trucks.length} camion(s) — basé sur la matricule</p>
          </div>
          <div className="flex gap-2 flex-wrap">
            <button onClick={fetchAll}
              className="bg-white/10 hover:bg-white/15 px-3 py-2 rounded-lg text-xs font-bold uppercase tracking-wider flex items-center gap-1.5 cursor-pointer">
              <RefreshCw className="w-3.5 h-3.5" /> Actualiser
            </button>
            <button onClick={handleExport}
              className="bg-emerald-600 hover:bg-emerald-700 text-white px-3 py-2 rounded-lg text-xs font-black uppercase tracking-wider flex items-center gap-1.5 cursor-pointer">
              <Download size={14} /> Export XLS
            </button>
            {/* Autosave status indicator */}
            <div className="flex items-center gap-1.5 px-3 py-2 rounded-lg text-xs font-bold uppercase tracking-wider">
              {saveStatus === 'saving' && (
                <span className="flex items-center gap-1.5 text-slate-300">
                  <Loader2 size={14} className="animate-spin" /> Enregistrement...
                </span>
              )}
              {saveStatus === 'saved' && (
                <span className="flex items-center gap-1.5 text-emerald-400">
                  <Check size={14} /> Enregistré
                </span>
              )}
              {saveStatus === 'error' && (
                <span className="flex items-center gap-1.5 text-rose-400">
                  <CloudOff size={14} /> Erreur
                </span>
              )}
            </div>
          </div>
        </div>
      </div>

      {loading ? (
        <div className="flex items-center justify-center py-20"><Loader2 className="w-6 h-6 text-blue-600 animate-spin" /></div>
      ) : trucks.length === 0 ? (
        <div className="bg-white rounded-xl border border-slate-200 py-16 text-center">
          <Truck size={32} className="text-slate-300 mx-auto mb-3" />
          <p className="text-sm text-slate-400 font-medium">Aucun camion trouvé.</p>
          <p className="text-xs text-slate-400 mt-1">Importez des chauffeurs (avec immatriculation) dans l'onglet Chauffeurs.</p>
        </div>
      ) : (
        <div className="grid grid-cols-1 lg:grid-cols-4 gap-6">

          {/* Truck selector */}
          <div className="bg-white rounded-xl border border-slate-200 p-3 space-y-1.5 lg:max-h-[80vh] lg:overflow-y-auto">
            <p className="text-[10px] font-black text-slate-400 uppercase tracking-widest px-1 mb-2">Camions</p>
            {trucks.map(t => {
              const k = computeKPIs(rows[t]);
              const isActive = activeTruck === t;
              return (
                <button key={t} onClick={() => setActiveTruck(t)}
                  className={`w-full text-left px-3 py-2.5 rounded-lg border transition-all cursor-pointer ${isActive ? 'bg-blue-600 border-blue-700 text-white' : 'border-slate-200 hover:border-blue-300 hover:bg-blue-50/40'}`}>
                  <p className={`text-sm font-bold font-mono ${isActive ? 'text-white' : 'text-slate-800'}`}>{t}</p>
                  <p className={`text-[10px] ${isActive ? 'text-blue-200' : k.benefice >= 0 ? 'text-emerald-600' : 'text-rose-500'}`}>
                    Bénéfice: {fmt(k.benefice)} MAD
                  </p>
                </button>
              );
            })}
          </div>

          {/* Active truck panel */}
          <div className="lg:col-span-3 space-y-6">
            {current && kpi && (
              <>
                {/* KPI summary */}
                <div className="bg-slate-900 text-white rounded-xl p-5">
                  <div className="flex items-center justify-between mb-4">
                    <p className="text-[10px] font-black text-slate-400 uppercase tracking-widest">
                      Résultats — <span className="font-mono text-teal-400">{activeTruck}</span>
                    </p>
                  </div>
                  <div className="grid grid-cols-2 md:grid-cols-4 gap-3">
                    <KpiCard label="Coût Revient Mensuel" value={`${fmt(kpi.coutRevientMensuel)} MAD`} />
                    <KpiCard label="Coût Revient / Km" value={`${fmt(kpi.coutRevientKm)} MAD`} />
                    <KpiCard label="Bénéfice" value={`${fmt(kpi.benefice)} MAD`} accent={kpi.benefice >= 0 ? 'emerald' : 'rose'} />
                    <KpiCard label="Bénéfice / Km" value={`${fmt(kpi.beneficeKm)} MAD`} accent={kpi.beneficeKm >= 0 ? 'emerald' : 'rose'} />
                    <KpiCard label="Marge Bénéficiaire" value={`${fmt(kpi.margeBeneficiaire)} %`} accent={kpi.margeBeneficiaire >= 0 ? 'emerald' : 'rose'} />
                    <KpiCard label="Prix Vente / Km" value={`${fmt(kpi.prixVenteKm)} MAD`} accent="teal" />
                    <KpiCard label="Estimation Prix Vente" value={`${fmt(kpi.estimationPrixVente)} MAD`} accent="teal" />
                  </div>
                </div>

                {/* Monthly inputs */}
                <Section
                  title="Données Mensuelles"
                  color="slate"
                  fields={MONTHLY_FIELDS as any}
                  row={current}
                  matricule={activeTruck}
                  setField={setField}
                  copyFieldToAll={copyFieldToAll}
                  onCopySection={() => copySectionToAll(MONTHLY_FIELDS.map(f => f.key), activeTruck, 'Données mensuelles')}
                  showTotal={false}
                />

                {/* Charges fixes */}
                <Section
                  title="Charges Fixes"
                  color="blue"
                  fields={CHARGES_FIXES as any}
                  row={current}
                  matricule={activeTruck}
                  setField={setField}
                  copyFieldToAll={copyFieldToAll}
                  onCopySection={() => copySectionToAll(FIXES_KEYS, activeTruck, 'Charges fixes')}
                  total={kpi.totalFixes}
                  totalLabel="Total Charges Fixes"
                />

                {/* Charges variables */}
                <Section
                  title="Les Charges Variables"
                  color="amber"
                  fields={CHARGES_VARIABLES as any}
                  row={current}
                  matricule={activeTruck}
                  setField={setField}
                  copyFieldToAll={copyFieldToAll}
                  onCopySection={() => copySectionToAll(VARIABLES_KEYS, activeTruck, 'Charges variables')}
                  total={kpi.totalVariables}
                  totalLabel="Total Charges Variables"
                />

                {/* Charges structure (shared overhead) */}
                <Section
                  title="Charge de Structure"
                  subtitle="Frais généraux du bureau — partagés. « Copier sur tous » applique la même valeur à chaque camion."
                  color="violet"
                  fields={CHARGES_STRUCTURE as any}
                  row={current}
                  matricule={activeTruck}
                  setField={setField}
                  copyFieldToAll={copyFieldToAll}
                  onCopySection={() => copySectionToAll(STRUCTURE_KEYS, activeTruck, 'Charge de structure')}
                  total={kpi.totalStructure}
                  totalLabel="Total Charges Structure"
                />
              </>
            )}
          </div>
        </div>
      )}
    </div>
  );
}

// ───────────────────────── Sub-components ─────────────────────────
function KpiCard({ label, value, accent }: { label: string; value: string; accent?: string }) {
  const colorMap: Record<string, string> = {
    emerald: 'text-emerald-400',
    rose: 'text-rose-400',
    teal: 'text-teal-400',
  };
  return (
    <div className="bg-white/5 rounded-lg p-3 border border-white/5">
      <p className="text-[9px] font-black text-slate-400 uppercase tracking-widest leading-tight">{label}</p>
      <p className={`text-base font-black mt-1 ${accent ? colorMap[accent] : 'text-white'}`}>{value}</p>
    </div>
  );
}

interface SectionProps {
  title: string;
  subtitle?: string;
  color: string;
  fields: { key: string; label: string }[];
  row: any;
  matricule: string;
  setField: (m: string, k: string, v: string) => void;
  copyFieldToAll: (k: string, m: string) => void;
  onCopySection: () => void;
  total?: number;
  totalLabel?: string;
  showTotal?: boolean;
}

function Section({
  title, subtitle, color, fields, row, matricule,
  setField, copyFieldToAll, onCopySection, total, totalLabel, showTotal = true,
}: SectionProps) {
  const head: Record<string, string> = {
    slate:  'bg-slate-50 text-slate-700',
    blue:   'bg-blue-50 text-blue-700',
    amber:  'bg-amber-50 text-amber-700',
    violet: 'bg-violet-50 text-violet-700',
  };
  const totalBar: Record<string, string> = {
    slate:  'bg-slate-100 text-slate-800',
    blue:   'bg-blue-600 text-white',
    amber:  'bg-amber-500 text-white',
    violet: 'bg-violet-600 text-white',
  };
  return (
    <div className="bg-white rounded-xl border border-slate-200 overflow-hidden">
      <div className={`px-5 py-3 border-b border-slate-200 flex items-center justify-between gap-3 flex-wrap ${head[color]}`}>
        <div>
          <p className="text-xs font-black uppercase tracking-widest">{title}</p>
          {subtitle && <p className="text-[10px] font-medium opacity-70 mt-0.5 normal-case tracking-normal">{subtitle}</p>}
        </div>
        <button onClick={onCopySection}
          className="bg-white/70 hover:bg-white border border-current/20 px-2.5 py-1.5 rounded-lg text-[10px] font-black uppercase tracking-wider flex items-center gap-1.5 cursor-pointer">
          <CopyPlus size={12} /> Copier section sur tous
        </button>
      </div>
      <div className="p-5">
        <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
          {fields.map(({ key, label }) => (
            <div key={key}>
              <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">{label}</label>
              <div className="flex gap-1.5 mt-1">
                <input
                  type="number"
                  value={row[key] ?? ''}
                  onChange={e => setField(matricule, key, e.target.value)}
                  className="flex-1 h-9 rounded-lg border-2 border-slate-200 px-3 text-sm font-semibold focus:outline-none focus:border-blue-500"
                  placeholder="0"
                />
                <button onClick={() => copyFieldToAll(key, matricule)}
                  title="Copier cette valeur sur tous les camions"
                  className="shrink-0 w-9 h-9 rounded-lg border-2 border-slate-200 text-slate-400 hover:text-blue-600 hover:border-blue-300 flex items-center justify-center cursor-pointer transition-colors">
                  <Copy size={13} />
                </button>
              </div>
            </div>
          ))}
        </div>
        {showTotal && total !== undefined && (
          <div className={`mt-4 px-4 py-3 rounded-lg flex items-center justify-between ${totalBar[color]}`}>
            <span className="text-xs font-black uppercase tracking-widest">{totalLabel}</span>
            <span className="text-lg font-black">{fmt(total)} MAD</span>
          </div>
        )}
      </div>
    </div>
  );
}