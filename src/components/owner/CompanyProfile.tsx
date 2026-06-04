import { useState, useEffect } from 'react';
import { motion } from 'motion/react';
import { Building2, Copy, Check, Calendar, Users, Mail, User, Edit2, Save, X } from 'lucide-react';
import { Company, StaffProfile, getCompanyStaff, isSupabaseConfigured, supabase } from '../../lib/auth';
import { toast } from 'sonner';

interface CompanyProfileProps {
  company: Company;
  ownerProfile: StaffProfile | null;
  ownerEmail: string;
  onCompanyUpdate?: (updatedCompany: Company) => void;
}

export default function CompanyProfile({ company, ownerProfile, ownerEmail, onCompanyUpdate }: CompanyProfileProps) {
  const [isEditing, setIsEditing] = useState(false);
  const [editedName, setEditedName] = useState(company.name);
  const [isSubmitting, setIsSubmitting] = useState(false);
  const [staffCount, setStaffCount] = useState(0);
  const [copied, setCopied] = useState(false);

  useEffect(() => {
    setEditedName(company.name);
    // Fetch staff count
    const fetchStaffCount = async () => {
      try {
        const staff = await getCompanyStaff(company.id);
        setStaffCount(staff.length);
      } catch (err) {
        console.error('Error fetching staff count:', err);
      }
    };
    fetchStaffCount();
  }, [company]);

  const handleCopyCode = async () => {
    try {
      await navigator.clipboard.writeText(company.business_code);
      setCopied(true);
      setTimeout(() => setCopied(false), 2000);
    } catch (err) {
      console.error('Could not copy:', err);
    }
  };

  const handleSave = async () => {
    if (!editedName.trim()) {
      return;
    }
    setIsSubmitting(true);
    try {
      if (isSupabaseConfigured) {
        const { data, error } = await supabase
          .from('companies')
          .update({ name: editedName.trim() })
          .eq('id', company.id)
          .select()
          .single();

        if (error) throw error;
        if (onCompanyUpdate && data) onCompanyUpdate(data);
      } else {
        // Mock Update
        const compRaw = localStorage.getItem('logiflow_mock_companies');
        if (compRaw) {
          const companies: Company[] = JSON.parse(compRaw);
          const index = companies.findIndex(c => c.id === company.id);
          if (index !== -1) {
            companies[index].name = editedName.trim();
            localStorage.setItem('logiflow_mock_companies', JSON.stringify(companies));
            if (onCompanyUpdate) onCompanyUpdate(companies[index]);
          }
        }
      }
      setIsEditing(false);
    } catch (err: any) {
      console.error('Error updating company name:', err);
    } finally {
      setIsSubmitting(false);
    }
  };

  const formatDate = (isoString?: string) => {
    if (!isoString) return 'Non disponible';
    const date = new Date(isoString);
    return date.toLocaleDateString('fr-FR', {
      day: 'numeric',
      month: 'long',
      year: 'numeric'
    });
  };

  return (
    <motion.div
      initial={{ opacity: 0, y: 15 }}
      animate={{ opacity: 1, y: 0 }}
      transition={{ duration: 0.4 }}
      className="bg-white rounded-xl border border-slate-200 shadow-sm p-6"
      id="comp-profile-card"
    >
      <div className="flex items-start justify-between border-b border-slate-100 pb-5 mb-5">
        <div>
          <span className="inline-flex items-center gap-1.5 px-2.5 py-1 rounded-md text-xs font-bold uppercase tracking-wider bg-blue-50 text-blue-600 mb-2">
            <Building2 className="w-3.5 h-3.5" /> Profil de l'Entreprise
          </span>
          {isEditing ? (
            <div className="flex items-center gap-2 mt-1">
              <input
                type="text"
                value={editedName}
                onChange={(e) => setEditedName(e.target.value)}
                disabled={isSubmitting}
                className="text-xl font-bold text-slate-900 border border-blue-250 outline-hidden focus:ring-2 focus:ring-blue-100 rounded-lg px-2 py-1 bg-slate-50 transition-all w-64 text-sm"
                id="edit-company-name-input"
              />
              <button
                onClick={handleSave}
                disabled={isSubmitting}
                className="p-1 px-2.5 rounded-lg bg-emerald-500 hover:bg-emerald-600 text-white transition-colors cursor-pointer"
                title="Enregistrer"
                id="save-company-name-button"
              >
                {isSubmitting ? '...' : <Save className="w-4 h-4" />}
              </button>
              <button
                onClick={() => {
                  setEditedName(company.name);
                  setIsEditing(false);
                }}
                disabled={isSubmitting}
                className="p-1 px-2.5 rounded-lg bg-slate-100 hover:bg-slate-200 text-slate-600 transition-colors cursor-pointer"
                title="Annuler"
                id="cancel-company-name-button"
              >
                <X className="w-4 h-4" />
              </button>
            </div>
          ) : (
            <h2 className="text-xl font-extrabold text-slate-900 flex items-center gap-2 tracking-tight">
              <span>{company.name}</span>
              <button
                onClick={() => setIsEditing(true)}
                className="p-1 hover:bg-slate-50 rounded-lg text-slate-400 hover:text-slate-600 transition-colors cursor-pointer"
                title="Modifier le nom"
                id="edit-company-name-trigger"
              >
                <Edit2 className="w-3.5 h-3.5" />
              </button>
            </h2>
          )}
        </div>
      </div>

      <div className="space-y-5">
        {/* Business Code Block: Matches high-contrast slate-900 design */}
        <div>
          <label className="text-[10px] font-bold text-slate-400 uppercase tracking-widest block mb-1">
            Code d'entreprise (Business Code)
          </label>
          <div className="flex items-center justify-between bg-slate-900 text-white rounded-lg p-3 group border border-slate-800">
            <code className="font-mono font-bold text-lg">{company.business_code}</code>
            <button
              onClick={handleCopyCode}
              className={`p-1.5 rounded transition-colors text-slate-400 hover:text-white cursor-pointer ${
                copied ? 'bg-emerald-600 text-white hover:bg-emerald-600' : 'hover:bg-slate-800'
              }`}
              title="Copier le code"
              id="copy-business-code-btn"
            >
              {copied ? <Check className="w-4 h-4" /> : <Copy className="w-4 h-4" />}
            </button>
          </div>
          <p className="text-[11px] text-slate-500 mt-2 leading-tight italic">
            Partagez ce code avec vos managers pour qu'ils puissent rejoindre votre entreprise.
          </p>
        </div>

        {/* 2-column metrics blocks styled to perfectly match design wireframe */}
        <div className="pt-1 grid grid-cols-2 gap-4">
          <div className="bg-slate-50 p-3 rounded-lg border border-slate-100">
            <span className="block text-[10px] font-bold text-slate-400 uppercase">Staff Total</span>
            <span className="text-2xl font-bold text-slate-950">{staffCount}</span>
          </div>
          <div className="bg-slate-50 p-3 rounded-lg border border-slate-100">
            <span className="block text-[10px] font-bold text-slate-400 uppercase">Activité</span>
            <span className="text-2xl font-bold text-green-600">Haute</span>
          </div>
        </div>

        {/* Profile Details Block */}
        <div className="pt-2 border-t border-slate-100 space-y-3">
          <div className="flex justify-between text-sm">
            <span className="text-slate-500">Propriétaire</span>
            <span className="font-semibold text-slate-900">{ownerProfile?.full_name || 'Chargement...'}</span>
          </div>
          <div className="flex justify-between text-sm">
            <span className="text-slate-500">Inscrit le</span>
            <span className="font-semibold text-slate-900">{formatDate(company.created_at)}</span>
          </div>
          <div className="flex justify-between text-sm">
            <span className="text-slate-500">Email</span>
            <span className="font-semibold text-slate-900 truncate max-w-[200px]" title={ownerEmail}>
              {ownerEmail || 'Chargement...'}
            </span>
          </div>
        </div>
      </div>
    </motion.div>
  );
}
