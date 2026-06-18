import { useState, useEffect } from 'react';
import { Loader2, Plus, FileText, Upload, Trash2, Copy, Check, Star, Eye, X, Download, Settings, Pencil } from 'lucide-react';
import { supabase } from '../../lib/supabase';
import { toast } from 'sonner';
import { motion, AnimatePresence } from 'motion/react';

// ═══════════════════════════════════════════════════════════════
// ZONE DETECTION ENGINE — auto-parses uploaded HTML into 5 zones
// ═══════════════════════════════════════════════════════════════
function detectZones(html: string) {
  const parser = new DOMParser();
  const doc = parser.parseFromString(html, 'text/html');

  // Extract <style> blocks
  const styles = Array.from(doc.querySelectorAll('style')).map(s => s.outerHTML).join('\n');

  // Find the main data table (the one with the most rows)
  const tables = Array.from(doc.querySelectorAll('table'));
  let mainTable: Element | null = null;
  let maxRows = 0;
  tables.forEach(t => {
    const rows = t.querySelectorAll('tr').length;
    if (rows > maxRows) { maxRows = rows; mainTable = t; }
  });

  if (!mainTable) {
    // No table found — store entire body as header, leave rest empty
    return {
      styles,
      header: doc.body.innerHTML,
      client: '',
      tableHead: '',
      tableRow: '',
      totals: '',
      footer: '',
    };
  }

  // Split document around the main table
  const bodyHTML = doc.body.innerHTML;
  const tableOuterHTML = (mainTable as Element).outerHTML;
  const tableIndex = bodyHTML.indexOf(tableOuterHTML);

  const beforeTable = bodyHTML.substring(0, tableIndex).trim();
  const afterTable = bodyHTML.substring(tableIndex + tableOuterHTML.length).trim();

  // Extract thead and first tbody row as template
  const thead = (mainTable as Element).querySelector('thead');
  const theadHTML = thead ? thead.innerHTML : '';
  const tbody = (mainTable as Element).querySelector('tbody');
  const firstRow = tbody?.querySelector('tr');
  const rowTemplate = firstRow ? firstRow.outerHTML : '';

  // Split beforeTable into header vs client
  // Heuristic: if there's a div/section with client-related content, split there
  let header = beforeTable;
  let client = '';

  // Look for client box patterns
  const clientPatterns = [
    /(<div[^>]*class="[^"]*client[^"]*"[^>]*>[\s\S]*?<\/div>)/i,
    /(<div[^>]*style="[^"]*border-left[^"]*"[^>]*>[\s\S]*?<\/div>)/i,
  ];
  for (const pattern of clientPatterns) {
    const match = beforeTable.match(pattern);
    if (match) {
      const idx = beforeTable.indexOf(match[0]);
      header = beforeTable.substring(0, idx).trim();
      client = match[0];
      break;
    }
  }

  // Split afterTable into totals vs footer
  // Heuristic: totals contain numbers/amounts, footer has signature/bank
  let totals = afterTable;
  let footer = '';

  const sigPatterns = [
    /(<div[^>]*class="[^"]*sig[^"]*"[^>]*>[\s\S]*?<\/div>)/i,
    /(<div[^>]*class="[^"]*footer[^"]*"[^>]*>[\s\S]*?<\/div>)/i,
    /(Signature[\s\S]*$)/i,
  ];
  for (const pattern of sigPatterns) {
    const match = afterTable.match(pattern);
    if (match) {
      const idx = afterTable.indexOf(match[0]);
      totals = afterTable.substring(0, idx).trim();
      footer = afterTable.substring(idx).trim();
      break;
    }
  }

  return { styles, header, client, tableHead: theadHTML, tableRow: rowTemplate, totals, footer };
}

// ═══════════════════════════════════════════════════════════════
// MULTI-PAGE PDF RENDERER — builds paginated HTML from zones
// ═══════════════════════════════════════════════════════════════
function generateInvoicePDF(
  template: any,
  rows: any[],
  placeholders: Record<string, string>,
  numberToWords: (n: number) => string,
) {
  const s = template;
  const mt = s.margin_top || 12;
  const mb = s.margin_bottom || 12;
  const ml = s.margin_left || 15;
  const mr = s.margin_right || 15;
  const fs = s.font_size || 11;
  const rpp = s.rows_per_page || 18;
  const primary = s.primary_color || '#1e40af';

  // Build the visible columns list
  const allCols = [
    { label: 'Date',          key: 'date',              show: s.col_show_date,             num: false },
    { label: 'N° Fact.',      key: 'numero_facture',    show: s.col_show_fact,             num: false },
    { label: 'Client',        key: 'client',            show: s.col_show_client,           num: false },
    { label: 'Départ',        key: 'depart',            show: s.col_show_depart,           num: false },
    { label: 'Arrivée',       key: 'arrivee',           show: s.col_show_arrivee,          num: false },
    { label: 'HT (MAD)',      key: 'montant_ht',        show: s.col_show_ht,               num: true  },
    { label: 'TVA (MAD)',     key: 'tva',               show: s.col_show_tva,              num: true  },
    { label: 'TTC (MAD)',     key: 'montant_ttc',       show: s.col_show_ttc,              num: true  },
    { label: 'BL/OT',        key: 'bl_ot',             show: s.col_show_bl_ot,            num: false },
    { label: 'BC',            key: 'bc',                show: s.col_show_bc,               num: false },
    { label: 'Délai',        key: 'delai_paiement',    show: s.col_show_delai,            num: false },
    { label: 'Date Paie.',    key: 'date_paiement',     show: s.col_show_date_paiement,    num: false },
    { label: 'Écart',        key: 'ecart_delai',       show: s.col_show_ecart,            num: false },
    { label: 'Mode',          key: 'mode_paiement',     show: s.col_show_mode,             num: false },
    { label: 'Statut',        key: 'statut',            show: s.col_show_statut,           num: false },
  ].filter(c => c.show !== false);

  const colW = Math.floor(100 / allCols.length);

  // Calculate totals
  const totalHT  = rows.reduce((sum, f) => sum + (f.montant_ht || 0), 0);
  const totalTVA = rows.reduce((sum, f) => sum + (f.tva || 0), 0);
  const totalTTC = rows.reduce((sum, f) => sum + (f.montant_ttc || 0), 0);

  // Replace placeholders in a string
  const fill = (html: string) => {
    let result = html || '';
    const allPlaceholders: Record<string, string> = {
      ...placeholders,
      total_ht: totalHT.toLocaleString('fr-MA', { minimumFractionDigits: 2 }) + ' MAD',
      total_tva: totalTVA.toLocaleString('fr-MA', { minimumFractionDigits: 2 }) + ' MAD',
      total_ttc: totalTTC.toLocaleString('fr-MA', { minimumFractionDigits: 2 }) + ' MAD',
      montant_lettres: numberToWords(totalTTC),
    };
    Object.entries(allPlaceholders).forEach(([k, v]) => {
      result = result.replaceAll(`{{${k}}}`, v);
    });
    return result;
  };

  // Build row HTML for a single data row
  const buildRow = (f: any) => allCols.map(c => {
    const v = f[c.key];
    if (c.num) return `<td style="text-align:right;padding:4px 6px;border-bottom:1px solid #f1f5f9;font-size:${fs-1}px">${Number(v||0).toLocaleString('fr-MA', {minimumFractionDigits:2})} MAD</td>`;
    return `<td style="padding:4px 6px;border-bottom:1px solid #f1f5f9;font-size:${fs-1}px">${v||'—'}</td>`;
  }).join('');

  // Thead HTML
  const theadHTML = s.zone_table_head
    ? `<thead>${fill(s.zone_table_head)}</thead>`
    : `<thead><tr>${allCols.map(c => `<th style="background:${primary};color:white;padding:5px 6px;text-align:left;font-size:${fs-3}px;text-transform:uppercase;width:${colW}%">${c.label}</th>`).join('')}</tr></thead>`;

  // Paginate rows
  const pages: { rows: any[]; isFirst: boolean; isLast: boolean; num: number }[] = [];
  const remaining = [...rows];
  let pageNum = 0;
  // First page has less room (client box + header)
  const firstPageRows = Math.max(5, rpp - 4);
  const lastPageRows = Math.max(5, rpp - 6); // room for totals + footer

  while (remaining.length > 0) {
    pageNum++;
    const isFirst = pageNum === 1;
    let capacity: number;
    if (isFirst && remaining.length <= firstPageRows) {
      capacity = remaining.length; // everything fits on page 1
    } else if (isFirst) {
      capacity = firstPageRows;
    } else if (remaining.length <= lastPageRows) {
      capacity = remaining.length; // last page
    } else {
      capacity = rpp;
    }
    const pageRows = remaining.splice(0, capacity);
    pages.push({ rows: pageRows, isFirst, isLast: remaining.length === 0, num: pageNum });
  }

  const totalPages = pages.length;

  // If template has imported zones, use them; otherwise generate built-in
  const hasImported = s.zone_header || s.original_html;

  // Build each page
  const pagesHTML = pages.map(p => {
    const headerHTML = (!s.skip_header && (s.zone_header || !hasImported)) ? fill(s.zone_header || `
      <div style="display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:14px;padding-bottom:10px;border-bottom:3px solid ${primary}">
        <div style="display:flex;align-items:center;gap:10px">
          ${s.logo_url ? `<img src="${s.logo_url}" style="max-height:55px;max-width:130px;object-fit:contain"/>` : ''}
          <div>
            <div style="font-size:${fs+7}px;font-weight:900;color:${primary}">${s.company_name || ''}</div>
            <div style="font-size:${fs-2}px;color:#64748b;margin-top:3px;line-height:1.6">
              ${s.address || ''}<br/>${s.phone ? 'Tél: '+s.phone : ''} ${s.email || ''}<br/>
              ${s.ice ? 'ICE: '+s.ice : ''} ${s.rc ? '— RC: '+s.rc : ''}
            </div>
          </div>
        </div>
        <div style="text-align:right">
          <div style="font-size:${fs+11}px;font-weight:900;color:${primary};letter-spacing:2px">${s.invoice_title || 'FACTURE'}</div>
          <div style="font-size:${fs+2}px;font-weight:700;color:${s.accent_color || '#f59e0b'};margin-top:3px">N° ${placeholders.numero_facture || ''}</div>
          <div style="font-size:${fs-2}px;color:#64748b;margin-top:3px">Date: ${placeholders.date || ''}<br/>Page ${p.num} / ${totalPages}</div>
        </div>
      </div>
    `) : `<div style="text-align:right;font-size:${fs-2}px;color:#64748b;margin-bottom:8px">Page ${p.num} / ${totalPages}</div>`;

    const clientHTML = p.isFirst ? fill(s.zone_client || `
      <div style="background:#f8fafc;border-left:4px solid ${s.accent_color || '#f59e0b'};padding:7px 12px;margin-bottom:12px;border-radius:0 6px 6px 0">
        <div style="font-size:${fs+1}px;font-weight:700;color:#1e293b">${placeholders.client || ''}</div>
        <div style="font-size:${fs-2}px;color:#64748b;margin-top:2px">${rows.length} prestation(s) — Délai: ${placeholders.delai_paiement || 60}j</div>
      </div>
    `) : '';

    const tableHTML = `
      <table style="width:100%;border-collapse:collapse;margin-bottom:10px">
        ${theadHTML}
        <tbody>
          ${p.rows.map((f, i) => `<tr${i%2===1?' style="background:#f8fafc"':''}>${buildRow(f)}</tr>`).join('')}
        </tbody>
      </table>
    `;

    const totalsHTML = p.isLast ? fill(s.zone_totals || `
      <div style="display:flex;justify-content:flex-end;margin-top:8px">
        <div style="border:2px solid ${primary};border-radius:6px;padding:8px 14px;min-width:200px">
          <div style="display:flex;justify-content:space-between;padding:2px 0;font-size:${fs}px"><span>Total HT</span><span>{{total_ht}}</span></div>
          <div style="display:flex;justify-content:space-between;padding:2px 0;font-size:${fs}px"><span>Total TVA</span><span>{{total_tva}}</span></div>
          <div style="display:flex;justify-content:space-between;padding:2px 0;font-size:${fs+2}px;font-weight:700;color:${primary};border-top:1px solid #e2e8f0;margin-top:4px;padding-top:5px"><span>Total TTC</span><span>{{total_ttc}}</span></div>
        </div>
      </div>
    `) : '';

    const footerHTML = (p.isLast && !s.skip_footer) ? fill(s.zone_footer || `
      ${(s.rib || s.bank_name) ? `<div style="margin-top:10px;padding:7px 12px;background:#f8fafc;border:1px solid #e2e8f0;border-radius:6px;font-size:${fs-2}px">
        ${s.bank_name ? 'Banque: <strong>'+s.bank_name+'</strong> ' : ''}${s.rib ? 'RIB: <strong>'+s.rib+'</strong>' : ''}
      </div>` : ''}
      <div style="display:flex;justify-content:flex-end;margin-top:14px">
        <div style="border:1px solid #cbd5e1;border-radius:6px;padding:8px 16px;min-width:160px;text-align:center">
          <div style="font-size:${fs-3}px;font-weight:700;color:#64748b;text-transform:uppercase">${s.signature_label || 'Signature & Cachet'}</div>
          <div style="height:38px"></div>
        </div>
      </div>
      ${s.footer_text ? `<div style="margin-top:14px;border-top:1px solid #e2e8f0;padding-top:5px;font-size:${fs-3}px;color:#94a3b8;text-align:center">${s.footer_text}</div>` : ''}
    `) : '';

    return `<div style="width:210mm;min-height:297mm;padding:${mt}mm ${mr}mm ${mb}mm ${ml}mm;page-break-after:${p.isLast ? 'auto' : 'always'};position:relative;font-family:Arial,sans-serif;font-size:${fs}px;color:#1e293b">
      ${headerHTML}${clientHTML}${tableHTML}${totalsHTML}${footerHTML}
    </div>`;
  }).join('');

  return `<!DOCTYPE html><html lang="fr"><head><meta charset="UTF-8"/>
    ${s.zone_styles || ''}<style>*{margin:0;padding:0;box-sizing:border-box}@media print{body{print-color-adjust:exact;-webkit-print-color-adjust:exact}}</style>
    </head><body>${pagesHTML}</body></html>`;
}

// ═══════════════════════════════════════════════════════════════
// numberToWords utility (French)
// ═══════════════════════════════════════════════════════════════
const numberToWords = (n: number): string => {
  if (n === 0) return 'Zéro dirham';
  const ones = ['','un','deux','trois','quatre','cinq','six','sept','huit','neuf','dix','onze','douze','treize','quatorze','quinze','seize','dix-sept','dix-huit','dix-neuf'];
  const tens = ['','','vingt','trente','quarante','cinquante','soixante','soixante','quatre-vingt','quatre-vingt'];
  const convert = (num: number): string => {
    if (num < 20) return ones[num];
    if (num < 100) { const t = Math.floor(num/10), r = num%10; if (t===7||t===9) return tens[t]+(r>0?'-'+ones[10+r]:(t===9?'-dix':'')); return tens[t]+(r>0?'-'+ones[r]:''); }
    if (num < 1000) return (num===100?'cent':ones[Math.floor(num/100)]+' cent')+(num%100>0?' '+convert(num%100):'');
    if (num < 1000000) { const m = Math.floor(num/1000); return (m===1?'mille':convert(m)+' mille')+(num%1000>0?' '+convert(num%1000):''); }
    return n.toLocaleString('fr-MA');
  };
  const intPart = Math.floor(n), decPart = Math.round((n-intPart)*100);
  return convert(intPart)+' dirham'+(intPart>1?'s':'')+(decPart>0?' et '+convert(decPart)+' centime'+(decPart>1?'s':''):'');
};

// ═══════════════════════════════════════════════════════════════
// MAIN COMPONENT — Template Library + Editor
// ═══════════════════════════════════════════════════════════════
interface Props { companyId: string; logoPreviewUrl: string; onGeneratePDF?: (templateId: string) => void; }

export default function InvoiceEngine({ companyId, logoPreviewUrl }: Props) {
  const [templates, setTemplates] = useState<any[]>([]);
  const [loading, setLoading] = useState(false);
  const [editing, setEditing] = useState<any | null>(null);
  const [uploading, setUploading] = useState(false);

  const fetchTemplates = async () => {
    if (!companyId) return;
    setLoading(true);
    const { data } = await supabase.from('invoice_templates').select('*').eq('company_id', companyId).order('created_at', { ascending: false });
    setTemplates(data || []);
    setLoading(false);
  };

  useEffect(() => { fetchTemplates(); }, [companyId]);

  const handleCreate = async () => {
    const { data, error } = await supabase.from('invoice_templates').insert({
      company_id: companyId,
      template_name: `Modèle ${templates.length + 1}`,
      is_default: templates.length === 0,
    }).select().single();
    if (!error && data) { setEditing(data); fetchTemplates(); toast.success("Nouveau modèle créé."); }
    else toast.error(`Erreur: ${error?.message}`);
  };

  const handleDuplicate = async (t: any) => {
    const { id, created_at, updated_at, ...rest } = t;
    const { error } = await supabase.from('invoice_templates').insert({ ...rest, template_name: `${t.template_name} (copie)`, is_default: false });
    if (!error) { fetchTemplates(); toast.success("Modèle dupliqué."); }
  };

  const handleDelete = async (id: string) => {
    if (!confirm('Supprimer ce modèle ?')) return;
    await supabase.from('invoice_templates').delete().eq('id', id);
    fetchTemplates(); toast.success("Supprimé.");
  };

  const handleSetDefault = async (id: string) => {
    await supabase.from('invoice_templates').update({ is_default: false }).eq('company_id', companyId);
    await supabase.from('invoice_templates').update({ is_default: true }).eq('id', id);
    fetchTemplates(); toast.success("Modèle par défaut mis à jour.");
  };

  const handleSave = async () => {
    if (!editing) return;
    const { error } = await supabase.from('invoice_templates').update({
      ...editing, updated_at: new Date().toISOString(),
    }).eq('id', editing.id);
    if (!error) { toast.success("Modèle sauvegardé."); fetchTemplates(); }
    else toast.error(`Erreur: ${error.message}`);
  };

  const handleImportHTML = async (file: File) => {
    setUploading(true);
    try {
      const html = await file.text();
      const zones = detectZones(html);
      setEditing((prev: any) => ({
        ...prev,
        original_html: html,
        zone_styles: zones.styles,
        zone_header: zones.header,
        zone_client: zones.client,
        zone_table_head: zones.tableHead,
        zone_table_row: zones.tableRow,
        zone_totals: zones.totals,
        zone_footer: zones.footer,
      }));
      toast.success(`Template importé — ${zones.header ? '✓ Header' : '✗ Header'} | ${zones.tableHead ? '✓ Table' : '✗ Table'} | ${zones.totals ? '✓ Totals' : '✗ Totals'} | ${zones.footer ? '✓ Footer' : '✗ Footer'}`);
    } catch (err: any) {
      toast.error(`Erreur: ${err.message}`);
    } finally {
      setUploading(false);
    }
  };

  const handleTestPDF = () => {
    if (!editing) return;
    const testRows = Array.from({ length: 5 }, (_, i) => ({
      date: `0${i+1}/06/2026`, numero_facture: `F-00${i+1}`, client: 'CLIENT TEST',
      depart: 'Casablanca', arrivee: 'Tanger', montant_ht: 5000 + i*1000,
      tva: 500 + i*100, montant_ttc: 5500 + i*1100, bl_ot: `BL-${i+1}`, bc: `BC-${i+1}`,
      delai_paiement: 60, date_paiement: '', ecart_delai: 0, mode_paiement: 'Virement', statut: 'impayé',
    }));
    const html = generateInvoicePDF(editing, testRows, {
      numero_facture: 'TEST-001', date: new Date().toLocaleDateString('fr-MA'),
      client: 'CLIENT TEST SARL', delai_paiement: '60',
    }, numberToWords);
    const win = window.open('', '_blank');
    if (win) { win.document.write(html); win.document.close(); win.focus(); setTimeout(() => win.print(), 600); }
  };

  // ───── EDITING VIEW ─────
  if (editing) {
    return (
      <div>
        <div className="mb-6 bg-slate-900 text-white rounded-xl p-6 border border-slate-800">
          <div className="flex items-center justify-between flex-wrap gap-3">
            <div>
              <span className="inline-flex items-center gap-1.5 px-2 py-0.5 rounded text-[10px] font-extrabold uppercase tracking-widest bg-violet-500 text-white mb-2">
                <Pencil className="w-3.5 h-3.5" /> Éditeur de Modèle
              </span>
              <h1 className="text-2xl font-extrabold tracking-tight">{editing.template_name}</h1>
            </div>
            <div className="flex gap-2">
              <button onClick={handleTestPDF} className="bg-violet-600 hover:bg-violet-700 text-white px-3 py-2 rounded-lg text-xs font-black uppercase tracking-wider flex items-center gap-1.5 cursor-pointer">
                <Eye size={14} /> Tester PDF
              </button>
              <button onClick={() => { handleSave(); setEditing(null); }} className="bg-blue-600 hover:bg-blue-700 text-white px-4 py-2 rounded-lg text-xs font-black uppercase tracking-wider flex items-center gap-1.5 cursor-pointer">
                <Check size={14} /> Sauvegarder & Retour
              </button>
              <button onClick={() => setEditing(null)} className="bg-white/10 hover:bg-white/15 px-3 py-2 rounded-lg text-xs font-bold uppercase tracking-wider flex items-center gap-1.5 cursor-pointer">
                <X size={14} /> Annuler
              </button>
            </div>
          </div>
        </div>

        <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
          {/* Left: Settings */}
          <div className="space-y-5">
            {/* Name */}
            <div className="bg-white rounded-xl border border-slate-200 p-5">
              <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Nom du Modèle</label>
              <input type="text" value={editing.template_name || ''} onChange={e => setEditing((p: any) => ({ ...p, template_name: e.target.value }))}
                className="w-full mt-1 h-9 rounded-lg border-2 border-slate-200 px-3 text-sm font-bold focus:outline-none focus:border-blue-500" />
            </div>

            {/* Import HTML */}
            <div className="bg-white rounded-xl border border-slate-200 p-5">
              <p className="text-xs font-black text-slate-700 uppercase tracking-widest mb-3">Importer un Template HTML</p>

              {/* Prompt guide */}
              <div className="mb-4 bg-slate-50 border border-slate-200 rounded-xl p-4">
                <p className="text-xs font-black text-slate-700 mb-2">📋 Vous avez un Excel ? Convertissez-le en HTML avec n'importe quel AI :</p>
                <p className="text-[10px] text-slate-500 mb-3">Copiez ce prompt, collez-le dans ChatGPT / Claude / Gemini, et uploadez votre fichier Excel à côté.</p>
                <div className="relative">
                  <pre className="bg-white border border-slate-200 rounded-lg p-3 text-[9px] text-slate-600 font-mono whitespace-pre-wrap max-h-32 overflow-y-auto leading-relaxed">{`Convert my attached invoice Excel file into a clean HTML template for a logistics invoicing system.

RULES:
1. Replicate my exact layout, colors, fonts, logo placement from the Excel
2. Use A4 page size (210mm × 297mm) with 15mm margins
3. Use only inline CSS or <style> block
4. Structure in order: HEADER, CLIENT BOX, DATA TABLE (<thead>+<tbody>), TOTALS, FOOTER
5. Replace dynamic data with: {{company_name}} {{company_address}} {{company_phone}} {{company_email}} {{company_ice}} {{company_rc}} {{company_logo}} {{invoice_title}} {{numero_facture}} {{date}} {{client}} {{delai_paiement}} {{total_ht}} {{total_tva}} {{total_ttc}} {{montant_lettres}} {{rib}} {{bank_name}} {{signature_label}} {{footer_text}}
6. Table body = <tbody>{{rows}}</tbody>
7. Add @media print CSS
8. Main data table must be the LARGEST table
9. Keep my original design
10. Output ONLY the HTML file`}</pre>
                  <button onClick={() => {
                    navigator.clipboard.writeText(`Convert my attached invoice Excel file into a clean HTML template for a logistics invoicing system.

RULES:
1. Replicate my exact layout, colors, fonts, logo placement from the Excel
2. Use A4 page size (210mm × 297mm) with 15mm margins
3. Use only inline CSS or <style> block
4. Structure in order: HEADER, CLIENT BOX, DATA TABLE (<thead>+<tbody>), TOTALS, FOOTER
5. Replace dynamic data with: {{company_name}} {{company_address}} {{company_phone}} {{company_email}} {{company_ice}} {{company_rc}} {{company_logo}} {{invoice_title}} {{numero_facture}} {{date}} {{client}} {{delai_paiement}} {{total_ht}} {{total_tva}} {{total_ttc}} {{montant_lettres}} {{rib}} {{bank_name}} {{signature_label}} {{footer_text}}
6. Table body = <tbody>{{rows}}</tbody>
7. Add @media print CSS
8. Main data table must be the LARGEST table
9. Keep my original design
10. Output ONLY the HTML file`);
                    toast.success("Prompt copié !");
                  }}
                    className="absolute top-2 right-2 bg-blue-600 hover:bg-blue-700 text-white px-2 py-1 rounded text-[9px] font-black uppercase cursor-pointer">
                    Copier
                  </button>
                </div>
              </div>

              <label className={`flex items-center justify-center h-20 border-2 border-dashed rounded-xl cursor-pointer transition-all group ${editing.original_html ? 'border-emerald-400 bg-emerald-50' : 'border-slate-300 hover:border-blue-400 hover:bg-blue-50'}`}>
                {editing.original_html ? (
                  <div className="flex items-center gap-2">
                    <FileText size={20} className="text-emerald-600" />
                    <div><p className="text-sm font-black text-emerald-700">Template importé ✓</p><p className="text-[10px] text-emerald-500">Zones auto-détectées. Cliquer pour remplacer.</p></div>
                  </div>
                ) : (
                  <div className="flex items-center gap-2"><Upload size={20} className="text-slate-400" /><span className="text-sm font-black text-slate-600">Uploader .html</span></div>
                )}
                <input type="file" accept=".html,.htm" className="hidden" onChange={e => { const f = e.target.files?.[0]; if (f) handleImportHTML(f); e.target.value = ''; }} />
              </label>
              {editing.original_html && (
                <div className="mt-3 grid grid-cols-2 gap-2 text-[10px]">
                  {[
                    { label: 'Header', zone: editing.zone_header },
                    { label: 'Client', zone: editing.zone_client },
                    { label: 'Table', zone: editing.zone_table_head },
                    { label: 'Totals', zone: editing.zone_totals },
                    { label: 'Footer', zone: editing.zone_footer },
                    { label: 'Styles', zone: editing.zone_styles },
                  ].map(z => (
                    <div key={z.label} className={`px-2 py-1.5 rounded-lg border ${z.zone ? 'border-emerald-200 bg-emerald-50 text-emerald-700' : 'border-slate-200 bg-slate-50 text-slate-400'}`}>
                      {z.zone ? '✓' : '✗'} {z.label}
                    </div>
                  ))}
                </div>
              )}
            </div>

            {/* Company Info */}
            <div className="bg-white rounded-xl border border-slate-200 p-5">
              <p className="text-xs font-black text-slate-700 uppercase tracking-widest mb-3">Identité Entreprise (pour modèle intégré)</p>
              <div className="grid grid-cols-2 gap-3">
                {[
                  { key: 'company_name', label: 'Raison sociale' }, { key: 'address', label: 'Adresse' },
                  { key: 'phone', label: 'Téléphone' }, { key: 'email', label: 'Email' },
                  { key: 'ice', label: 'ICE' }, { key: 'rc', label: 'RC' },
                  { key: 'rib', label: 'RIB' }, { key: 'bank_name', label: 'Banque' },
                  { key: 'signature_label', label: 'Label Signature' },
                ].map(f => (
                  <div key={f.key}>
                    <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">{f.label}</label>
                    <input type="text" value={editing[f.key] || ''} onChange={e => setEditing((p: any) => ({ ...p, [f.key]: e.target.value }))}
                      className="w-full mt-1 h-8 rounded-lg border-2 border-slate-200 px-3 text-xs focus:outline-none focus:border-blue-500" />
                  </div>
                ))}
              </div>
              <div className="mt-3">
                <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Mentions légales (footer)</label>
                <textarea value={editing.footer_text || ''} onChange={e => setEditing((p: any) => ({ ...p, footer_text: e.target.value }))}
                  rows={2} className="w-full mt-1 rounded-lg border-2 border-slate-200 px-3 py-2 text-xs focus:outline-none focus:border-blue-500 resize-none" />
              </div>
            </div>

            {/* Layout */}
            <div className="bg-white rounded-xl border border-slate-200 p-5">
              <p className="text-xs font-black text-slate-700 uppercase tracking-widest mb-3">Mise en Page</p>
              <div className="grid grid-cols-2 gap-3">
                {[
                  { key: 'invoice_title', label: 'Titre', type: 'text' },
                  { key: 'rows_per_page', label: 'Lignes/page', type: 'number' },
                  { key: 'font_size', label: 'Taille police', type: 'number' },
                  { key: 'primary_color', label: 'Couleur principale', type: 'color' },
                ].map(f => (
                  <div key={f.key}>
                    <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">{f.label}</label>
                    <input type={f.type} value={editing[f.key] || ''} onChange={e => setEditing((p: any) => ({ ...p, [f.key]: f.type === 'number' ? parseInt(e.target.value) : e.target.value }))}
                      className="w-full mt-1 h-8 rounded-lg border-2 border-slate-200 px-3 text-xs focus:outline-none focus:border-blue-500" />
                  </div>
                ))}
              </div>
              <div className="mt-3 flex gap-4">
                {[{ key: 'skip_header', label: 'Papier pré-imprimé (pas header)' }, { key: 'skip_footer', label: 'Pas de footer' }].map(f => (
                  <label key={f.key} className="flex items-center gap-2 cursor-pointer">
                    <input type="checkbox" checked={editing[f.key] || false} onChange={e => setEditing((p: any) => ({ ...p, [f.key]: e.target.checked }))} />
                    <span className="text-[10px] font-bold text-slate-600">{f.label}</span>
                  </label>
                ))}
              </div>
            </div>
          </div>

          {/* Right: Column selector */}
          <div className="space-y-5">
            <div className="bg-white rounded-xl border border-slate-200 p-5">
              <p className="text-xs font-black text-slate-700 uppercase tracking-widest mb-3">Colonnes visibles dans le PDF</p>
              <div className="grid grid-cols-2 gap-2">
                {[
                  { label: 'Date', key: 'col_show_date' }, { label: 'N° Facture', key: 'col_show_fact' },
                  { label: 'Client', key: 'col_show_client' }, { label: 'Départ', key: 'col_show_depart' },
                  { label: 'Arrivée', key: 'col_show_arrivee' }, { label: 'HT', key: 'col_show_ht' },
                  { label: 'TVA', key: 'col_show_tva' }, { label: 'TTC', key: 'col_show_ttc' },
                  { label: 'BL/OT', key: 'col_show_bl_ot' }, { label: 'BC', key: 'col_show_bc' },
                  { label: 'Délai', key: 'col_show_delai' }, { label: 'Date Paie.', key: 'col_show_date_paiement' },
                  { label: 'Écart', key: 'col_show_ecart' }, { label: 'Mode', key: 'col_show_mode' },
                  { label: 'Statut', key: 'col_show_statut' },
                ].map(c => (
                  <div key={c.key} onClick={() => setEditing((p: any) => ({ ...p, [c.key]: !p[c.key] }))}
                    className={`flex items-center gap-2 p-2.5 rounded-xl border-2 cursor-pointer transition-all ${editing[c.key] !== false ? 'border-blue-400 bg-blue-50' : 'border-slate-200 bg-slate-50 opacity-60'}`}>
                    <div className={`w-4 h-4 rounded flex items-center justify-center shrink-0 border-2 ${editing[c.key] !== false ? 'bg-blue-600 border-blue-600' : 'border-slate-300'}`}>
                      {editing[c.key] !== false && <span className="text-white text-[9px] font-black">✓</span>}
                    </div>
                    <span className={`text-[10px] font-bold ${editing[c.key] !== false ? 'text-blue-800' : 'text-slate-400'}`}>{c.label}</span>
                  </div>
                ))}
              </div>
            </div>

            {/* Quick test */}
            <div className="bg-violet-50 border border-violet-200 rounded-xl p-5 flex items-center justify-between">
              <div>
                <p className="text-sm font-black text-violet-800">Tester avec données fictives</p>
                <p className="text-[10px] text-violet-600 mt-0.5">Génère un PDF test pour vérifier le rendu multi-pages.</p>
              </div>
              <button onClick={handleTestPDF} className="bg-violet-600 hover:bg-violet-700 text-white px-4 py-2 rounded-lg text-xs font-black uppercase cursor-pointer flex items-center gap-1.5">
                <Eye size={14} /> Tester
              </button>
            </div>
          </div>
        </div>
      </div>
    );
  }

  // ───── LIBRARY VIEW ─────
  return (
    <div>
      <div className="mb-6 bg-slate-900 text-white rounded-xl p-6 border border-slate-800">
        <div className="flex items-center justify-between flex-wrap gap-3">
          <div>
            <span className="inline-flex items-center gap-1.5 px-2 py-0.5 rounded text-[10px] font-extrabold uppercase tracking-widest bg-slate-500 text-white mb-2">
              <Settings className="w-3.5 h-3.5" /> Moteur de Facturation
            </span>
            <h1 className="text-2xl font-extrabold tracking-tight">Modèles de Facture</h1>
            <p className="text-sm text-slate-400 mt-1">{templates.length} modèle(s) — importez ou créez vos formats</p>
          </div>
          <button onClick={handleCreate} className="bg-blue-600 hover:bg-blue-700 text-white px-4 py-2 rounded-lg text-xs font-black uppercase tracking-wider flex items-center gap-1.5 cursor-pointer">
            <Plus size={14} /> Nouveau Modèle
          </button>
        </div>
      </div>

      {loading ? (
        <div className="flex items-center justify-center py-20"><Loader2 className="w-6 h-6 text-blue-600 animate-spin" /></div>
      ) : templates.length === 0 ? (
        <div className="bg-white rounded-xl border border-slate-200 py-16 text-center">
          <FileText size={40} className="text-slate-300 mx-auto mb-3" />
          <p className="text-sm text-slate-400 font-medium">Aucun modèle de facture.</p>
          <p className="text-xs text-slate-400 mt-1">Cliquez sur "Nouveau Modèle" pour commencer.</p>
          <button onClick={handleCreate} className="mt-4 bg-blue-600 hover:bg-blue-700 text-white px-6 py-2.5 rounded-lg text-xs font-black uppercase tracking-wider cursor-pointer">
            <Plus size={14} className="inline mr-1" /> Créer mon premier modèle
          </button>
        </div>
      ) : (
        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
          {templates.map(t => (
            <div key={t.id} className={`bg-white rounded-xl border-2 overflow-hidden transition-all ${t.is_default ? 'border-blue-400 shadow-lg shadow-blue-100' : 'border-slate-200 hover:border-blue-300'}`}>
              <div className="p-5">
                <div className="flex items-start justify-between mb-3">
                  <div>
                    <p className="text-sm font-black text-slate-900">{t.template_name}</p>
                    <p className="text-[10px] text-slate-400 mt-0.5">{t.original_html ? 'Importé' : 'Intégré'} — {t.invoice_title || 'FACTURE'}</p>
                  </div>
                  {t.is_default && (
                    <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded text-[9px] font-black bg-blue-600 text-white uppercase">
                      <Star size={10} /> Défaut
                    </span>
                  )}
                </div>
                {/* Mini preview */}
                <div className="h-24 rounded-lg border border-slate-200 bg-slate-50 flex items-center justify-center text-slate-300 mb-3 overflow-hidden">
                  <div className="w-full h-full p-2" style={{ fontSize: '4px' }}>
                    <div className="h-2 rounded bg-slate-200 mb-1" style={{ backgroundColor: t.primary_color || '#1e40af', width: '60%' }} />
                    <div className="h-1 rounded bg-slate-100 mb-0.5 w-3/4" />
                    <div className="h-1 rounded bg-slate-100 mb-0.5 w-1/2" />
                    <div className="h-1 rounded bg-slate-100 mb-0.5 w-2/3" />
                    <div className="h-1 rounded bg-slate-100 mb-0.5 w-3/5" />
                    <div className="h-1.5 rounded mt-1" style={{ backgroundColor: t.primary_color || '#1e40af', width: '40%', marginLeft: 'auto' }} />
                  </div>
                </div>
                <div className="flex gap-1.5">
                  <button onClick={() => setEditing(t)} className="flex-1 h-8 bg-blue-600 hover:bg-blue-700 text-white text-[10px] font-black uppercase rounded-lg cursor-pointer flex items-center justify-center gap-1">
                    <Pencil size={11} /> Modifier
                  </button>
                  <button onClick={() => handleDuplicate(t)} className="h-8 px-2 bg-slate-100 hover:bg-slate-200 text-slate-600 text-[10px] font-bold rounded-lg cursor-pointer" title="Dupliquer">
                    <Copy size={13} />
                  </button>
                  {!t.is_default && (
                    <button onClick={() => handleSetDefault(t.id)} className="h-8 px-2 bg-amber-50 hover:bg-amber-100 text-amber-600 text-[10px] font-bold rounded-lg cursor-pointer" title="Définir par défaut">
                      <Star size={13} />
                    </button>
                  )}
                  <button onClick={() => handleDelete(t.id)} className="h-8 px-2 bg-rose-50 hover:bg-rose-100 text-rose-500 text-[10px] font-bold rounded-lg cursor-pointer" title="Supprimer">
                    <Trash2 size={13} />
                  </button>
                </div>
              </div>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

// Export the PDF generator so Facturation tab can use it
export { generateInvoicePDF, numberToWords };