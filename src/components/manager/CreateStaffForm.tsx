import { useState, useEffect, FormEvent } from 'react';
import { motion, AnimatePresence } from 'motion/react';
import { 
  UserPlus, Mail, Lock, Shield, Check, Copy, ToggleLeft, ToggleRight, 
  Car, Eye, EyeOff, ClipboardCheck, Briefcase, FileText, BadgeCheck 
} from 'lucide-react';
import { 
  getCompanyStaff, createStaffAccount, insertStaffProfile, 
  toggleStaffActive, StaffProfile, isSupabaseConfigured 
} from '../../lib/auth';
import { toast } from 'sonner';

interface CreateStaffFormProps {
  companyId: string;
}

const ROLE_LABELS: Record<string, string> = {
  accountant: 'Comptable',
  dispatcher: 'Dispatcher / Répartiteur',
  mechanic: 'Mécanicien',
  driver: 'Chauffeur / Conducteur',
  staff: 'Personnel de bureau',
  cleaning: 'Entretien / Nettoyage',
  security: 'Sécurité',
  local: 'Local / Logistique'
};

export default function CreateStaffForm({ companyId }: CreateStaffFormProps) {
  // Form State
  const [fullName, setFullName] = useState('');
  const [role, setRole] = useState('driver');
  const [employeeCode, setEmployeeCode] = useState('');
  const [vehiclePlate, setVehiclePlate] = useState('');
  const [email, setEmail] = useState('');
  const [tempPassword, setTempPassword] = useState('');
  
  const [showPassword, setShowPassword] = useState(false);
  const [isSubmitting, setIsSubmitting] = useState(false);

  // Staff List State
  const [staffList, setStaffList] = useState<StaffProfile[]>([]);
  const [loadingList, setLoadingList] = useState(true);

  // Security Credentials Modal
  const [createdCredentials, setCreatedCredentials] = useState<{
    name: string;
    role: string;
    code: string;
    email: string;
    pass: string;
  } | null>(null);
  const [copiedModal, setCopiedModal] = useState(false);

  // Load staff list (exclude owner and other managers as per prompt)
  const loadStaffList = async () => {
    try {
      setLoadingList(true);
      const data = await getCompanyStaff(companyId, { excludeRoles: ['owner', 'manager'] });
      setStaffList(data);
    } catch (err) {
      console.error('Error fetching staff:', err);
    } finally {
      setLoadingList(false);
    }
  };

  useEffect(() => {
    if (companyId) {
      loadStaffList();
    }
  }, [companyId]);

  // Pass generation
  const handleGeneratePassword = () => {
    const chars = 'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789!@#$%&*';
    let password = '';
    for (let i = 0; i < 12; i++) {
      password += chars.charAt(Math.floor(Math.random() * chars.length));
    }
    setTempPassword(password);
  };

  const handleCreateStaff = async (e: FormEvent) => {
    e.preventDefault();

    if (!fullName.trim() || !employeeCode.trim() || !email.trim() || !tempPassword.trim()) {
      alert("Veuillez remplir tous les champs obligatoires.");
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

      // 2. Clear out plate if role is not driver or mechanic
      const actualPlate = (role === 'driver' || role === 'mechanic') ? vehiclePlate.trim() : '';

      // 3. Insert staff profile record inside DB
      const { profile: newProfile, error: profileErr } = await insertStaffProfile({
        company_id: companyId,
        full_name: fullName.trim(),
        role: role,
        employee_code: employeeCode.trim(),
        vehicle_plate: actualPlate || undefined,
        auth_user_id: authId,
        is_active: true
      });

      if (profileErr) {
        throw new Error(profileErr.message || "Impossible de sauvegarder le profil employé.");
      }

      // Success
      setCreatedCredentials({
        name: fullName.trim(),
        role: role,
        code: employeeCode.trim(),
        email: email.trim(),
        pass: tempPassword.trim()
      });

      // Clear Form
      setFullName('');
      setEmployeeCode('');
      setVehiclePlate('');
      setEmail('');
      setTempPassword('');
      setShowPassword(false);

      // Reload list
      await loadStaffList();

    } catch (err: any) {
      console.error('Create staff error:', err);
      alert(`Erreur de création de l'employé: ${err.message || "Impossible de finaliser l'opération"}`);
    } finally {
      setIsSubmitting(false);
    }
  };

  const handleToggleActive = async (staff: StaffProfile) => {
    try {
      const prevStatus = staff.is_active;
      const updatedStatus = await toggleStaffActive(staff.id, prevStatus);

      setStaffList(prev => prev.map(s => {
        if (s.id === staff.id) {
          return { ...s, is_active: updatedStatus };
        }
        return s;
      }));
    } catch (err) {
      console.error('Toggle staff status error:', err);
    }
  };

  const copyCredsToClipboard = async () => {
    if (!createdCredentials) return;
    const blockText = `--- IDENTIFIANTS EMPLOYÉ LOGI-FLOW ---
Nom: ${createdCredentials.name}
Role: ${ROLE_LABELS[createdCredentials.role] || createdCredentials.role}
Code Employé: ${createdCredentials.code}
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

  return (
    <div className="space-y-8" id="create-staff-component-root">
      
      {/* Form Card */}
      <motion.div
        initial={{ opacity: 0, y: 15 }}
        animate={{ opacity: 1, y: 0 }}
        transition={{ duration: 0.4 }}
        className="bg-white rounded-xl border border-slate-200 shadow-xs p-6"
      >
        <div className="flex items-center gap-2 border-b border-slate-100 pb-4 mb-5">
          <div className="h-8 w-8 rounded bg-blue-50 flex items-center justify-center text-blue-600">
            <UserPlus className="w-4 h-4" />
          </div>
          <div>
            <h3 className="font-extrabold text-slate-850 tracking-tight">Recruter un Nouvel Employé</h3>
            <p className="text-xs text-slate-500">Ajouter les informations et les accès du nouveau membre d'équipage</p>
          </div>
        </div>

        <form onSubmit={handleCreateStaff} className="grid grid-cols-1 md:grid-cols-2 gap-4">
          
          {/* Left Column: Core staff profile details */}
          <div className="space-y-4">
            <div>
              <label className="block text-xs font-bold text-slate-500 uppercase tracking-wider mb-1.5">
                Nom complet de l'employé <span className="text-red-500">*</span>
              </label>
              <input
                type="text"
                required
                value={fullName}
                onChange={(e) => setFullName(e.target.value)}
                placeholder="Ex: Paul Martin"
                className="w-full text-sm rounded-lg border border-slate-200 px-3.5 py-2.5 outline-hidden focus:ring-2 focus:ring-blue-100 focus:border-blue-500 transition-all text-slate-800"
                id="staff-fullname-input"
              />
            </div>

            <div>
              <label className="block text-xs font-bold text-slate-500 uppercase tracking-wider mb-1.5">
                Rôle attribué <span className="text-red-500">*</span>
              </label>
              <div className="relative">
                <Briefcase className="absolute left-3.5 top-3 w-4 h-4 text-slate-400" />
                <select
                  value={role}
                  onChange={(e) => setRole(e.target.value)}
                  className="w-full text-sm rounded-lg border border-slate-200 pl-10 pr-3.5 py-2.5 outline-hidden focus:ring-2 focus:ring-blue-100 focus:border-blue-500 transition-all text-slate-800 bg-white cursor-pointer"
                  id="staff-role-select"
                >
                  <option value="driver">Chauffeur / Conducteur</option>
                  <option value="dispatcher">Dispatcher / Répartiteur</option>
                  <option value="mechanic">Mécanicien</option>
                  <option value="accountant">Comptable</option>
                  <option value="staff">Personnel de bureau</option>
                  <option value="cleaning">Entretien / Nettoyage</option>
                  <option value="security">Sécurité</option>
                  <option value="local">Local / Logistique</option>
                </select>
              </div>
            </div>

            <div>
              <label className="block text-xs font-bold text-slate-500 uppercase tracking-wider mb-1.5">
                Code Employé (Unique) <span className="text-red-500">*</span>
              </label>
              <div className="relative">
                <FileText className="absolute left-3.5 top-3 w-4 h-4 text-slate-400" />
                <input
                  type="text"
                  required
                  value={employeeCode}
                  onChange={(e) => setEmployeeCode(e.target.value)}
                  placeholder="Ex: DRV-482"
                  className="w-full text-sm rounded-lg border border-slate-200 pl-10 pr-3.5 py-2.5 outline-hidden focus:ring-2 focus:ring-blue-100 focus:border-blue-500 transition-all text-slate-800 font-mono"
                  id="staff-code-input"
                />
              </div>
            </div>

            {/* Condition: Show Plate only if role is Driver or Mechanic */}
            <AnimatePresence mode="popLayout">
              {(role === 'driver' || role === 'mechanic') && (
                <motion.div
                  initial={{ opacity: 0, height: 0 }}
                  animate={{ opacity: 1, height: 'auto' }}
                  exit={{ opacity: 0, height: 0 }}
                  transition={{ duration: 0.2 }}
                  className="overflow-hidden"
                >
                  <label className="block text-xs font-bold text-slate-500 uppercase tracking-wider mb-1.5">
                    Plaque d'Immatriculation du Véhicule
                  </label>
                  <div className="relative">
                    <Car className="absolute left-3.5 top-3 w-4 h-4 text-slate-400" />
                    <input
                      type="text"
                      value={vehiclePlate}
                      onChange={(e) => setVehiclePlate(e.target.value)}
                      placeholder="Ex: AA-123-BB"
                      className="w-full text-sm rounded-lg border border-slate-200 pl-10 pr-3.5 py-2.5 outline-hidden focus:ring-2 focus:ring-blue-100 focus:border-blue-500 transition-all text-slate-800 uppercase font-mono"
                      id="staff-plate-input"
                    />
                  </div>
                </motion.div>
              )}
            </AnimatePresence>
          </div>

          {/* Right Column: Connection Details */}
          <div className="space-y-4">
            <div>
              <label className="block text-xs font-bold text-slate-500 uppercase tracking-wider mb-1.5">
                Adresse Email (Connexion) <span className="text-red-500">*</span>
              </label>
              <div className="relative">
                <Mail className="absolute left-3.5 top-3 w-4 h-4 text-slate-400" />
                <input
                  type="email"
                  required
                  value={email}
                  onChange={(e) => setEmail(e.target.value)}
                  placeholder="Ex: salarie@logiflow.com"
                  className="w-full text-sm rounded-lg border border-slate-200 pl-10 pr-3.5 py-2.5 outline-hidden focus:ring-2 focus:ring-blue-100 focus:border-blue-500 transition-all text-slate-800"
                  id="staff-email-input"
                />
              </div>
            </div>

            <div>
              <label className="block text-xs font-bold text-slate-500 uppercase tracking-wider mb-1.5">
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
                    placeholder="Mots de passe"
                    className="w-full text-sm rounded-lg border border-slate-200 pl-10 pr-10 py-2.5 outline-hidden focus:ring-2 focus:ring-blue-100 focus:border-blue-500 transition-all text-slate-800 font-mono"
                    id="staff-password-input"
                  />
                  <button
                    type="button"
                    onClick={() => setShowPassword(!showPassword)}
                    className="absolute right-3 top-3 text-slate-400 hover:text-slate-600 transition-colors"
                  >
                    {showPassword ? <EyeOff className="w-4 h-4" /> : <Eye className="w-4 h-4" />}
                  </button>
                </div>
                <button
                  type="button"
                  onClick={handleGeneratePassword}
                  className="px-3 py-2 text-xs font-bold uppercase tracking-wider rounded-lg border border-slate-200 hover:bg-slate-50 text-slate-700 transition-all cursor-pointer"
                  id="staff-generate-password-btn"
                >
                  Générer
                </button>
              </div>
            </div>

            <div className="pt-2">
              <span className="block text-[11px] text-slate-400 mb-2 leading-relaxed italic">
                Après la création, l'employé pourra se connecter avec son email et ce mot de passe de départ.
              </span>
              <button
                type="submit"
                disabled={isSubmitting}
                className="w-full inline-flex items-center justify-center gap-2 bg-blue-600 hover:bg-blue-700 text-white font-bold py-2.5 px-4 rounded-lg transition-all shadow-xs cursor-pointer disabled:opacity-50"
                id="submit-staff-form-btn"
              >
                {isSubmitting ? 'Création...' : "Créer le Profil d'Équipage"}
              </button>
            </div>
          </div>
        </form>
      </motion.div>

      {/* Staff Roster Table */}
      <motion.div
        initial={{ opacity: 0, y: 15 }}
        animate={{ opacity: 1, y: 0 }}
        transition={{ duration: 0.4, delay: 0.1 }}
        className="bg-white rounded-xl border border-slate-200 shadow-xs p-6"
      >
        <div className="flex items-center justify-between border-b border-slate-100 pb-4 mb-5">
          <div className="flex items-center gap-2">
            <div className="h-8 w-8 rounded bg-blue-50 flex items-center justify-center text-blue-600">
              <BadgeCheck className="w-4 h-4" />
            </div>
            <div>
              <h3 className="font-extrabold text-slate-850 tracking-tight">Équipes Actives</h3>
              <p className="text-xs text-slate-500">Chauffeurs, mécanos et dispatchers sous votre supervision</p>
            </div>
          </div>
          <span className="text-xs font-extrabold px-2.5 py-1 bg-slate-50 text-slate-600 border border-slate-200 rounded">
            {staffList.length} équipiers
          </span>
        </div>

        {loadingList ? (
          <div className="py-8 text-center text-slate-400 text-sm">Chargement de la liste...</div>
        ) : staffList.length === 0 ? (
          <div className="py-12 border border-dashed border-slate-200 rounded-lg text-center text-slate-400 text-sm">
            Aucun membre d'équipage n'a été créé pour le moment.
          </div>
        ) : (
          <div className="overflow-x-auto rounded-lg border border-slate-200">
            <table className="w-full text-left border-collapse" id="staff-table">
              <thead>
                <tr className="bg-slate-50 text-[11px] font-bold text-slate-400 uppercase tracking-wider border-b border-slate-200">
                  <th className="p-4">Code Employé</th>
                  <th className="p-4">Nom Complet</th>
                  <th className="p-4">Rôle</th>
                  <th className="p-4">Immatriculation</th>
                  <th className="p-4 text-center">Accès</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100 text-sm text-slate-700">
                {staffList.map((staff) => (
                  <tr key={staff.id} className="hover:bg-slate-55/40 transition-colors">
                    <td className="p-4 font-mono text-xs font-bold text-blue-600">{staff.employee_code}</td>
                    <td className="p-4 font-bold text-slate-800">{staff.full_name}</td>
                    <td className="p-4 text-xs font-medium text-slate-600">
                      <span className="px-2 py-0.5 bg-slate-105 border border-slate-200 rounded">
                        {ROLE_LABELS[staff.role] || staff.role}
                      </span>
                    </td>
                    <td className="p-4">
                      {staff.vehicle_plate ? (
                        <span className="font-mono text-xs bg-amber-50 text-amber-800 border border-amber-200 px-2 py-0.5 rounded">
                          {staff.vehicle_plate}
                        </span>
                      ) : (
                        <span className="text-slate-400">-</span>
                      )}
                    </td>
                    <td className="p-4 text-center">
                      <button
                        onClick={() => handleToggleActive(staff)}
                        className={`inline-flex items-center gap-1 px-3 py-1 rounded-full text-xs font-bold uppercase tracking-wider cursor-pointer transition-all ${
                          staff.is_active 
                            ? 'bg-emerald-50 text-emerald-700 hover:bg-emerald-100/80' 
                            : 'bg-slate-100 text-slate-400 hover:bg-slate-150'
                        }`}
                        id={`toggle-staff-${staff.id}`}
                      >
                        {staff.is_active ? 'Actif' : 'Bloqué'}
                      </button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </motion.div>

      {/* SECURE CREATED CREDENTIALS MODAL */}
      <AnimatePresence>
        {createdCredentials && (
          <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-xs flex items-center justify-center p-4 z-50">
            <motion.div
              initial={{ scale: 0.95, opacity: 0 }}
              animate={{ scale: 1, opacity: 1 }}
              exit={{ scale: 0.95, opacity: 0 }}
              className="bg-white rounded-xl shadow-xl border border-slate-200 p-6 max-w-md w-full relative"
              id="staff-credentials-modal"
            >
              <div className="text-center mb-5">
                <div className="mx-auto h-12 w-12 rounded bg-emerald-100 flex items-center justify-center text-emerald-650 mb-3">
                  <Shield className="w-6 h-6" />
                </div>
                <h4 className="text-lg font-extrabold text-slate-900 tracking-tight">Compte Employé Créé</h4>
                <p className="text-xs text-slate-500 mt-1 leading-normal">
                  Les identifiants ne s'afficheront qu'<u>une seule fois</u> pour des raisons de sécurité.
                </p>
              </div>

              <div className="space-y-3 bg-slate-900 text-white p-4 rounded-lg border border-slate-800 text-xs">
                <div>
                  <span className="block text-[9px] uppercase font-bold text-slate-400 mb-0.5">Nom Complet</span>
                  <span className="font-bold text-sm text-white">{createdCredentials.name}</span>
                </div>
                <div>
                  <span className="block text-[9px] uppercase font-bold text-slate-400 mb-0.5">Rôle</span>
                  <span className="font-semibold text-slate-300">{ROLE_LABELS[createdCredentials.role] || createdCredentials.role}</span>
                </div>
                <div>
                  <span className="block text-[9px] uppercase font-bold text-slate-400 mb-0.5">Code Employé</span>
                  <span className="font-mono font-bold text-blue-400">{createdCredentials.code}</span>
                </div>
                <div>
                  <span className="block text-[9px] uppercase font-bold text-slate-400 mb-0.5">Identifiant Email (Connexion)</span>
                  <span className="font-mono text-slate-300">{createdCredentials.email}</span>
                </div>
                <div>
                  <span className="block text-[9px] uppercase font-bold text-slate-400 mb-0.5">Mot de passe temporaire</span>
                  <span className="font-mono font-black text-emerald-400 select-all block mt-1 tracking-widest text-base">
                    {createdCredentials.pass}
                  </span>
                </div>
              </div>

              <div className="mt-6 flex flex-col gap-2">
                <button
                  onClick={copyCredsToClipboard}
                  className={`w-full py-2.5 rounded-lg font-bold text-xs uppercase tracking-wider flex items-center justify-center gap-2 cursor-pointer transition-colors ${
                    copiedModal 
                      ? 'bg-emerald-600 text-white' 
                      : 'bg-blue-600 hover:bg-blue-700 text-white'
                  }`}
                  id="modal-copy-staff-btn"
                >
                  {copiedModal ? (
                    <>
                      <ClipboardCheck className="w-4 h-4" /> Copié !
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
                    toast.success("Identifiants employés transmis avec succès.");
                  }}
                  className="w-full py-2 rounded-lg text-slate-500 hover:text-slate-800 hover:bg-slate-50 text-xs font-bold uppercase tracking-wider transition-colors cursor-pointer text-center"
                  id="modal-close-staff-btn"
                >
                  J'ai Copié — Fermer
                </button>
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>

    </div>
  );
}
