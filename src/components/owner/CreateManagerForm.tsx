import { useState, useEffect, FormEvent } from 'react';
import { motion, AnimatePresence } from 'motion/react';
import { 
  UserPlus, Mail, Lock, Shield, Check, Copy, ToggleLeft, ToggleRight, 
  Calendar, Trash, Users, Eye, EyeOff, ClipboardCheck 
} from 'lucide-react';
import { 
  getCompanyStaff, createStaffAccount, insertStaffProfile, 
  toggleStaffActive, StaffProfile, isSupabaseConfigured 
} from '../../lib/auth';
import { toast } from 'sonner';

interface CreateManagerFormProps {
  companyId: string;
}

export default function CreateManagerForm({ companyId }: CreateManagerFormProps) {
  // Form State
  const [fullName, setFullName] = useState('');
  const [email, setEmail] = useState('');
  const [tempPassword, setTempPassword] = useState('');
  const [showPassword, setShowPassword] = useState(false);
  const [isSubmitting, setIsSubmitting] = useState(false);
  
  // Managers List state
  const [managers, setManagers] = useState<StaffProfile[]>([]);
  const [loadingList, setLoadingList] = useState(true);

  // Success Modal State for Credentials Display
  const [createdCredentials, setCreatedCredentials] = useState<{
    name: string;
    email: string;
    pass: string;
  } | null>(null);
  const [copiedModal, setCopiedModal] = useState(false);

  // Load managers
  const loadManagersList = async () => {
    try {
      setLoadingList(true);
      const data = await getCompanyStaff(companyId, { role: 'manager' });
      setManagers(data);
    } catch (err) {
      console.error('Error fetching managers:', err);
    } finally {
      setLoadingList(false);
    }
  };

  useEffect(() => {
    if (companyId) {
      loadManagersList();
    }
  }, [companyId]);

  // Handle password generation
  const handleGeneratePassword = () => {
    const chars = 'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789!@#$%&*';
    let password = '';
    for (let i = 0; i < 12; i++) {
      password += chars.charAt(Math.floor(Math.random() * chars.length));
    }
    setTempPassword(password);
  };

  const handleCreateManager = async (e: FormEvent) => {
    e.preventDefault();
    
    if (!fullName.trim() || !email.trim() || !tempPassword.trim()) {
      alert("Veuillez remplir tous les champs.");
      return;
    }

    if (tempPassword.length < 6) {
      alert("Le mot de passe doit comporter au moins 6 caractères.");
      return;
    }

    setIsSubmitting(true);

    try {
      // 1. Create native Auth user
      const { user: authUser, error: authErr } = await createStaffAccount(email.trim(), tempPassword.trim());
      
      if (authErr) {
        throw new Error(authErr.message || "Erreur lors de la création de l'authentification.");
      }

      const authId = authUser.id;

      // 2. Insert profile record
      // We store the email in employee_code for managers so we can display it!
      const { profile: newProfile, error: profileErr } = await insertStaffProfile({
        company_id: companyId,
        full_name: fullName.trim(),
        role: 'manager',
        employee_code: email.trim(), // Save email here for future lookups
        auth_user_id: authId,
        is_active: true
      });

      if (profileErr) {
        throw new Error(profileErr.message || "Impossible de sauvegarder le profil manager.");
      }

      // Success
      setCreatedCredentials({
        name: fullName.trim(),
        email: email.trim(),
        pass: tempPassword.trim()
      });

      // Clear Form
      setFullName('');
      setEmail('');
      setTempPassword('');
      setShowPassword(false);

      // Reload
      await loadManagersList();

    } catch (err: any) {
      console.error('Create manager error:', err);
      alert(`Erreur de création: ${err.message || "Impossible de finaliser l'opération"}`);
    } finally {
      setIsSubmitting(false);
    }
  };

  const handleToggleActive = async (manager: StaffProfile) => {
    try {
      const prevStatus = manager.is_active;
      const updatedStatus = await toggleStaffActive(manager.id, prevStatus);
      
      setManagers(prev => prev.map(m => {
        if (m.id === manager.id) {
          return { ...m, is_active: updatedStatus };
        }
        return m;
      }));
    } catch (err) {
      console.error('Toggle status error:', err);
    }
  };

  const copyCredsToClipboard = async () => {
    if (!createdCredentials) return;
    const blockText = `--- IDENTIFIANTS MANAGER LOGI-FLOW ---
Nom: ${createdCredentials.name}
Email: ${createdCredentials.email}
Mot de passe temporaire: ${createdCredentials.pass}
----------------------------------------`;
    try {
      await navigator.clipboard.writeText(blockText);
      setCopiedModal(true);
      setTimeout(() => setCopiedModal(false), 2000);
    } catch (err) {
      console.error('Failed to copy', err);
    }
  };

  const formatDate = (dateStr?: string) => {
    if (!dateStr) return '';
    return new Date(dateStr).toLocaleDateString('fr-FR', {
      day: 'numeric',
      month: 'short',
      year: 'numeric'
    });
  };

  return (
    <div className="space-y-8" id="create-manager-component-root">
      
      {/* Form Card */}
      <motion.div
        initial={{ opacity: 0, y: 15 }}
        animate={{ opacity: 1, y: 0 }}
        transition={{ duration: 0.4, delay: 0.1 }}
        className="bg-white rounded-xl border border-slate-200 shadow-sm p-6"
      >
        <div className="flex items-center gap-2 border-b border-slate-100 pb-4 mb-5">
          <div className="h-8 w-8 rounded bg-blue-50 flex items-center justify-center text-blue-600">
            <UserPlus className="w-4 h-4" />
          </div>
          <div>
            <h3 className="font-bold text-slate-900 text-base tracking-tight">Recruter un Manager</h3>
            <p className="text-xs text-slate-500">Ajouter un nouveau compte gestionnaire pour cette entreprise</p>
          </div>
        </div>

        <form onSubmit={handleCreateManager} className="space-y-4">
          <div>
            <label className="block text-xs font-bold uppercase tracking-wider text-slate-500 mb-1">
              Nom complet du Manager <span className="text-red-500">*</span>
            </label>
            <input
              type="text"
              required
              value={fullName}
              onChange={(e) => setFullName(e.target.value)}
              placeholder="Ex: Jean Dupont"
              className="w-full text-sm rounded-lg border border-slate-200 px-3.5 py-2.5 outline-hidden focus:ring-2 focus:ring-blue-100 focus:border-blue-500 transition-all text-slate-850"
              id="manager-fullname-input"
            />
          </div>

          <div>
            <label className="block text-xs font-bold uppercase tracking-wider text-slate-500 mb-1">
              Adresse Email <span className="text-red-500">*</span>
            </label>
            <div className="relative">
              <Mail className="absolute left-3.5 top-3 w-4 h-4 text-slate-400" />
              <input
                type="email"
                required
                value={email}
                onChange={(e) => setEmail(e.target.value)}
                placeholder="Ex: manager@entreprise.com"
                className="w-full text-sm rounded-lg border border-slate-200 pl-10 pr-3.5 py-2.5 outline-hidden focus:ring-2 focus:ring-blue-100 focus:border-blue-500 transition-all text-slate-850"
                id="manager-email-input"
              />
            </div>
          </div>

          <div>
            <label className="block text-xs font-bold uppercase tracking-wider text-slate-500 mb-1">
              Mot de passe temporaire <span className="text-red-500">*</span>
            </label>
            <div className="flex gap-2">
              <div className="relative flex-1">
                <Lock className="absolute left-3.5 top-3 w-4 h-4 text-slate-400" />
                <input
                  type={showPassword ? 'text' : 'password'}
                  required
                  value={tempPassword}
                  onChange={(e) => setTempPassword(e.target.value)}
                  placeholder="Minimum 6 caractères"
                  className="w-full text-sm rounded-lg border border-slate-200 pl-10 pr-10 py-2.5 outline-hidden focus:ring-2 focus:ring-blue-100 focus:border-blue-500 transition-all text-slate-850 font-mono"
                  id="manager-password-input"
                />
                <button
                  type="button"
                  onClick={() => setShowPassword(!showPassword)}
                  className="absolute right-3 top-3 text-slate-400 hover:text-slate-600 transition-colors cursor-pointer"
                  id="toggle-manager-password-view"
                >
                  {showPassword ? <EyeOff className="w-4 h-4" /> : <Eye className="w-4 h-4" />}
                </button>
              </div>
              <button
                type="button"
                onClick={handleGeneratePassword}
                className="px-3 py-2 text-xs font-semibold rounded-lg border border-slate-200 hover:bg-slate-50 text-slate-700 transition-colors cursor-pointer"
                id="generate-password-btn"
              >
                Générer
              </button>
            </div>
          </div>

          <button
            type="submit"
            disabled={isSubmitting}
            className="w-full mt-2 inline-flex items-center justify-center gap-2 bg-blue-600 hover:bg-blue-700 text-white text-sm font-semibold py-2.5 px-4 rounded-lg transition-all shadow-sm cursor-pointer disabled:opacity-50"
            id="submit-manager-form-btn"
          >
            {isSubmitting ? 'Création en cours...' : 'Créer le Compte Manager'}
          </button>
        </form>
      </motion.div>

      {/* Managers List Table */}
      <motion.div
        initial={{ opacity: 0, y: 15 }}
        animate={{ opacity: 1, y: 0 }}
        transition={{ duration: 0.4, delay: 0.2 }}
        className="bg-white rounded-xl border border-slate-200 shadow-sm flex flex-col overflow-hidden"
      >
        <div className="p-6 border-b flex justify-between items-center bg-slate-50/50">
          <div>
            <h3 className="text-lg font-bold text-slate-900">Gestion des Managers</h3>
            <p className="text-sm text-slate-500">Créez et gérez les comptes d'accès pour vos gestionnaires d'entrepôt.</p>
          </div>
          <span className="text-xs font-bold uppercase tracking-widest px-2.5 py-1 bg-slate-100 text-slate-600 border border-slate-200 rounded-md">
            {managers.length} manager{managers.length > 1 ? 's' : ''}
          </span>
        </div>

        {loadingList ? (
          <div className="py-8 text-center text-slate-400 text-sm italic font-medium">Chargement de la liste...</div>
        ) : managers.length === 0 ? (
          <div className="p-12 text-center text-slate-400 text-sm border-dashed border border-slate-200 m-6 rounded-xl">
            Aucun manager n'a été créé pour le moment.
          </div>
        ) : (
          <div className="overflow-x-auto">
            <table className="w-full text-left" id="managers-table">
              <thead className="bg-slate-50 sticky top-0 z-10 border-b border-slate-200">
                <tr>
                  <th className="px-6 py-3 text-[10px] font-bold text-slate-400 uppercase tracking-widest">Nom Complet</th>
                  <th className="px-6 py-3 text-[10px] font-bold text-slate-400 uppercase tracking-widest">Email</th>
                  <th className="px-6 py-3 text-[10px] font-bold text-slate-400 uppercase tracking-widest">Créé le</th>
                  <th className="px-6 py-3 text-[10px] font-bold text-slate-400 uppercase tracking-widest text-center">Statut</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100">
                {managers.map((manager) => (
                  <tr key={manager.id} className="hover:bg-slate-50 transition-colors">
                    <td className="px-6 py-4">
                      <div className="font-semibold text-slate-900">{manager.full_name}</div>
                    </td>
                    <td className="px-6 py-4 text-sm text-slate-600 font-mono">
                      {manager.employee_code || manager.auth_user_id}
                    </td>
                    <td className="px-6 py-4 text-sm text-slate-500 italic">
                      {formatDate(manager.created_at)}
                    </td>
                    <td className="px-6 py-4">
                      <div className="flex justify-center">
                        <button
                          onClick={() => handleToggleActive(manager)}
                          className={`px-2.5 py-1 rounded-full text-[11px] font-bold uppercase tracking-wider transition-colors cursor-pointer ${
                            manager.is_active 
                              ? 'bg-green-100 text-green-700 hover:bg-green-200' 
                              : 'bg-slate-100 text-slate-400 hover:bg-slate-200'
                          }`}
                          id={`toggle-active-${manager.id}`}
                        >
                          {manager.is_active ? 'Actif' : 'Inactif'}
                        </button>
                      </div>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </motion.div>

      {/* SECURE CREATED CREDENTIALS MODAL (Shown exactly once, then never again) */}
      <AnimatePresence>
        {createdCredentials && (
          <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-xs flex items-center justify-center p-4 z-50">
            <motion.div
              initial={{ scale: 0.9, opacity: 0 }}
              animate={{ scale: 1, opacity: 1 }}
              exit={{ scale: 0.9, opacity: 0 }}
              className="bg-white rounded-xl shadow-xl border border-slate-200 p-6 max-w-md w-full relative"
              id="credentials-success-modal"
            >
              <div className="text-center mb-5">
                <div className="mx-auto h-12 w-12 rounded bg-emerald-100 flex items-center justify-center text-emerald-600 mb-3">
                  <Shield className="w-6 h-6" />
                </div>
                <h4 className="text-lg font-bold text-slate-900">Compte Manager Créé</h4>
                <p className="text-xs text-slate-500 mt-1">
                  Les identifiants ne s'afficheront qu'<u>une seule fois</u> pour des raisons de sécurité.
                </p>
              </div>

              <div className="space-y-3 bg-slate-50 p-4 rounded-lg border border-slate-100 text-sm">
                <div>
                  <span className="block text-[10px] uppercase font-bold text-slate-400">Nom Complet</span>
                  <span className="font-semibold text-slate-800">{createdCredentials.name}</span>
                </div>
                <div>
                  <span className="block text-[10px] uppercase font-bold text-slate-400">Identifiant Email</span>
                  <span className="font-mono text-xs text-slate-800">{createdCredentials.email}</span>
                </div>
                <div>
                  <span className="block text-[10px] uppercase font-bold text-slate-400">Mot de passe temporaire</span>
                  <span className="font-mono text-sm font-bold text-blue-700 bg-blue-50 px-1.5 py-0.5 rounded-sm select-all">
                    {createdCredentials.pass}
                  </span>
                </div>
              </div>

              <div className="mt-6 flex flex-col gap-2">
                <button
                  onClick={copyCredsToClipboard}
                  className={`w-full py-2.5 rounded-lg font-semibold text-sm flex items-center justify-center gap-2 cursor-pointer transition-colors ${
                    copiedModal 
                      ? 'bg-emerald-500 text-white' 
                      : 'bg-blue-600 hover:bg-blue-700 text-white shadow-xs'
                  }`}
                  id="modal-copy-credentials-btn"
                >
                  {copiedModal ? (
                    <>
                      <ClipboardCheck className="w-4 h-4" /> Copié dans le presse-papier
                    </>
                  ) : (
                    <>
                      <Copy className="w-4 h-4" /> Copier les Identifiants
                    </>
                  )}
                </button>
                <button
                  onClick={() => {
                    setCreatedCredentials(null);
                    toast.success("Identifiants à transmettre manuellement.");
                  }}
                  className="w-full py-2 rounded-lg text-slate-550 hover:text-slate-800 hover:bg-slate-50 text-sm font-medium transition-colors cursor-pointer text-center"
                  id="modal-close-credentials-btn"
                >
                  J'ai noté les identifiants — Fermer
                </button>
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>

    </div>
  );
}
