import { useState, useEffect } from 'react';
import { Loader2, Upload, FileText, Download, Trash2, CheckCircle2, XCircle, Truck, RefreshCw, Eye } from 'lucide-react';
import { supabase } from '../../lib/supabase';
import { toast } from 'sonner';

const DOC_TYPES = [
  { key: 'carte_grise',       label: 'Carte Grise' },
  { key: 'carte_verte',       label: 'Carte Verte' },
  { key: 'visite_technique',  label: 'Visite Technique' },
  { key: 'tachygraphe',       label: 'Tachygraphe' },
  { key: 'extincteur_2kg',    label: 'Extincteur 2kg' },
  { key: 'extincteur_6kg',    label: 'Extincteur 6kg' },
  { key: 'assurance',         label: 'Assurance' },
  { key: 'vignette',          label: 'Vignette' },
  { key: 'visite_grue',       label: 'Visite Grue' },
  { key: 'permis',            label: 'Permis de Conduire' },
  { key: 'cin',               label: 'CIN' },
  { key: 'attestation_grue',  label: 'Attestation Grue' },
];

interface Props { companyId: string; }

export default function TruckDocumentsTab({ companyId }: Props) {
  const [trucks, setTrucks] = useState<any[]>([]);
  const [activeTruck, setActiveTruck] = useState('');
  const [docs, setDocs] = useState<Record<string, any>>({});
  const [allDocs, setAllDocs] = useState<any[]>([]);
  const [loading, setLoading] = useState(false);
  const [uploading, setUploading] = useState<string | null>(null);
  const [viewingDoc, setViewingDoc] = useState<string | null>(null);

  const fetchTrucks = async () => {
    if (!companyId) return;
    const { data } = await supabase.from('fleet_drivers')
      .select('immatriculation, type_vehicule').eq('company_id', companyId);
    const list = (data || []).filter((d: any) => d.immatriculation);
    setTrucks(list);
    if (list.length && !activeTruck) setActiveTruck(list[0].immatriculation);
    // Fetch all docs for overview
    const { data: allD } = await supabase.from('truck_documents').select('*').eq('company_id', companyId);
    setAllDocs(allD || []);
  };

  const fetchDocs = async () => {
    if (!companyId || !activeTruck) return;
    setLoading(true);
    const { data } = await supabase.from('truck_documents')
      .select('*').eq('company_id', companyId).eq('matricule', activeTruck);
    const map: Record<string, any> = {};
    (data || []).forEach((d: any) => { map[d.doc_type] = d; });
    setDocs(map);
    setLoading(false);
  };

  useEffect(() => { fetchTrucks(); }, [companyId]);
  useEffect(() => { fetchDocs(); }, [activeTruck, companyId]);

  const handleUpload = async (docType: string, file: File) => {
    if (!companyId || !activeTruck) return;
    setUploading(docType);
    try {
      const ext = file.name.split('.').pop() || 'pdf';
      const path = `${companyId}/${activeTruck}/${docType}.${ext}`;
      const { error: upErr } = await supabase.storage.from('truck-documents').upload(path, file, { upsert: true });
      if (upErr) throw upErr;
      const { data: urlData } = supabase.storage.from('truck-documents').getPublicUrl(path);
      const { error: dbErr } = await supabase.from('truck_documents').upsert({
        company_id: companyId, matricule: activeTruck, doc_type: docType,
        file_url: urlData.publicUrl, file_name: file.name, storage_path: path,
        uploaded_by: 'manager_portal',
      }, { onConflict: 'company_id,matricule,doc_type' });
      if (dbErr) throw dbErr;
      toast.success(`${file.name} uploadé.`);
      fetchDocs(); fetchTrucks();
    } catch (err: any) { toast.error(`Erreur: ${err.message}`); }
    finally { setUploading(null); }
  };

  const handleDelete = async (docType: string) => {
    if (!confirm('Supprimer ce document ?')) return;
    const doc = docs[docType]; if (!doc) return;
    if (doc.storage_path) await supabase.storage.from('truck-documents').remove([doc.storage_path]);
    await supabase.from('truck_documents').delete().eq('id', doc.id);
    toast.success('Supprimé.'); fetchDocs(); fetchTrucks();
  };

  const uploaded = DOC_TYPES.filter(d => docs[d.key]);
  const missing = DOC_TYPES.filter(d => !docs[d.key]);

  // Overview: count docs per truck
  const truckDocCounts: Record<string, number> = {};
  allDocs.forEach(d => { truckDocCounts[d.matricule] = (truckDocCounts[d.matricule] || 0) + 1; });

  return (
    <div>
      <div className="mb-6 bg-slate-900 text-white rounded-xl p-6 border border-slate-800">
        <div className="flex items-center justify-between flex-wrap gap-3">
          <div>
            <span className="inline-flex items-center gap-1.5 px-2 py-0.5 rounded text-[10px] font-extrabold uppercase tracking-widest bg-violet-500 text-white mb-2">
              <FileText className="w-3.5 h-3.5" /> Documents Camions
            </span>
            <h1 className="text-2xl font-extrabold tracking-tight">Archivage & Téléchargement</h1>
            <p className="text-sm text-slate-400 mt-1">{trucks.length} camion(s) — {allDocs.length} document(s) archivés</p>
          </div>
          <button onClick={() => { fetchTrucks(); fetchDocs(); }}
            className="bg-white/10 hover:bg-white/15 px-3 py-2 rounded-lg text-xs font-bold uppercase tracking-wider flex items-center gap-1.5 cursor-pointer">
            <RefreshCw className="w-3.5 h-3.5" /> Actualiser
          </button>
        </div>
      </div>

      {trucks.length === 0 ? (
        <div className="bg-white rounded-xl border border-slate-200 py-16 text-center">
          <Truck size={32} className="text-slate-300 mx-auto mb-3" />
          <p className="text-sm text-slate-400">Aucun camion trouvé.</p>
        </div>
      ) : (
        <div className="grid grid-cols-1 lg:grid-cols-4 gap-6">
          {/* Truck list */}
          <div className="bg-white rounded-xl border border-slate-200 p-3 space-y-1.5 lg:max-h-[80vh] lg:overflow-y-auto">
            <p className="text-[10px] font-black text-slate-400 uppercase tracking-widest px-1 mb-2">Camions</p>
            {trucks.map((t: any) => {
              const isActive = activeTruck === t.immatriculation;
              const count = truckDocCounts[t.immatriculation] || 0;
              return (
                <button key={t.immatriculation} onClick={() => setActiveTruck(t.immatriculation)}
                  className={`w-full text-left px-3 py-2.5 rounded-lg border transition-all cursor-pointer ${isActive ? 'bg-blue-600 border-blue-700 text-white' : 'border-slate-200 hover:border-blue-300 hover:bg-blue-50/40'}`}>
                  <div className="flex items-center justify-between gap-2">
                    <p className={`text-sm font-bold font-mono ${isActive ? 'text-white' : 'text-slate-800'}`}>{t.immatriculation}</p>
                    <span className={`text-[10px] font-black px-1.5 py-0.5 rounded ${isActive ? 'bg-white/20 text-white' : count === DOC_TYPES.length ? 'bg-emerald-100 text-emerald-700' : count > 0 ? 'bg-amber-100 text-amber-700' : 'bg-slate-100 text-slate-500'}`}>
                      {count}/{DOC_TYPES.length}
                    </span>
                  </div>
                  <p className={`text-[10px] ${isActive ? 'text-blue-200' : 'text-slate-500'}`}>{t.type_vehicule || '—'}</p>
                </button>
              );
            })}
          </div>

          {/* Documents panel */}
          <div className="lg:col-span-3 space-y-3">
            {loading ? (
              <div className="flex items-center justify-center py-20"><Loader2 className="w-6 h-6 text-blue-600 animate-spin" /></div>
            ) : (
              <>
                <div className="flex items-center justify-between mb-2">
                  <p className="text-xs font-black text-slate-500 uppercase tracking-widest">
                    <span className="font-mono text-blue-600">{activeTruck}</span>
                    <span className="ml-3 text-emerald-600">{uploaded.length} uploadés</span>
                    {missing.length > 0 && <span className="ml-2 text-rose-500">{missing.length} manquants</span>}
                  </p>
                </div>
                {DOC_TYPES.map(dt => {
                  const doc = docs[dt.key];
                  const isUploading = uploading === dt.key;
                  return (
                    <div key={dt.key} className={`rounded-xl border-2 p-4 flex items-center justify-between gap-4 transition-colors ${doc ? 'border-emerald-200 bg-emerald-50/30' : 'border-slate-200 bg-white'}`}>
                      <div className="flex items-center gap-3 min-w-0">
                        {doc ? <CheckCircle2 size={20} className="text-emerald-500 shrink-0" />
                              : <XCircle size={20} className="text-slate-300 shrink-0" />}
                        <div className="min-w-0">
                          <p className={`text-sm font-bold ${doc ? 'text-emerald-800' : 'text-slate-700'}`}>{dt.label}</p>
                          {doc ? (
                            <p className="text-[10px] text-slate-500 truncate">{doc.file_name} — {new Date(doc.created_at).toLocaleDateString('fr-MA')} — via {doc.uploaded_by === 'manager_portal' ? 'Manager' : 'Achats'}</p>
                          ) : (
                            <p className="text-[10px] text-slate-400">Aucun document</p>
                          )}
                        </div>
                      </div>
                      <div className="flex items-center gap-1.5 shrink-0">
                        {doc && (
                          <>
                            <button onClick={() => setViewingDoc(doc.file_url)}
                              className="p-2 rounded-lg hover:bg-blue-50 text-blue-600 transition-colors" title="Aperçu">
                              <Eye size={16} />
                            </button>
                            <a href={doc.file_url} download={doc.file_name}
                              className="p-2 rounded-lg hover:bg-emerald-50 text-emerald-600 transition-colors" title="Télécharger">
                              <Download size={16} />
                            </a>
                            <button onClick={() => handleDelete(dt.key)}
                              className="p-2 rounded-lg hover:bg-rose-50 text-slate-400 hover:text-rose-600 transition-colors" title="Supprimer">
                              <Trash2 size={16} />
                            </button>
                          </>
                        )}
                        <label className={`flex items-center gap-1.5 px-3 py-2 rounded-lg text-[10px] font-black uppercase tracking-wider cursor-pointer transition-all text-white ${isUploading ? 'bg-slate-400' : doc ? 'bg-amber-500 hover:bg-amber-600' : 'bg-blue-600 hover:bg-blue-700'}`}>
                          {isUploading ? <Loader2 size={12} className="animate-spin" /> : <Upload size={12} />}
                          {isUploading ? '...' : doc ? 'Remplacer' : 'Upload'}
                          <input type="file" accept=".pdf,.png,.jpg,.jpeg" className="hidden" disabled={!!uploading}
                            onChange={e => { const f = e.target.files?.[0]; if (f) handleUpload(dt.key, f); e.target.value = ''; }} />
                        </label>
                      </div>
                    </div>
                  );
                })}
              </>
            )}
          </div>
        </div>
      )}

      {/* Document viewer modal */}
      {viewingDoc && (
        <div className="fixed inset-0 bg-black/80 z-50 flex items-center justify-center p-4" onClick={() => setViewingDoc(null)}>
          <div className="max-w-4xl w-full max-h-[90vh] bg-white rounded-xl overflow-hidden" onClick={e => e.stopPropagation()}>
            {viewingDoc.match(/\.(png|jpg|jpeg|gif)$/i) ? (
              <img src={viewingDoc} alt="Document" className="w-full h-auto max-h-[85vh] object-contain" />
            ) : (
              <iframe src={viewingDoc} className="w-full h-[85vh]" />
            )}
            <div className="p-3 border-t border-slate-200 flex justify-end gap-2">
              <a href={viewingDoc} download className="px-4 py-2 bg-emerald-600 text-white text-xs font-black uppercase rounded-lg flex items-center gap-1.5">
                <Download size={14} /> Télécharger
              </a>
              <button onClick={() => setViewingDoc(null)} className="px-4 py-2 bg-slate-200 text-slate-700 text-xs font-bold rounded-lg cursor-pointer">Fermer</button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}