import { useState } from 'react';
import { Loader2, UserPlus, Mail, User, Lock, CheckCircle2 } from 'lucide-react';
import { createStaffAccount, insertStaffProfile } from '../../lib/auth';
import { toast } from 'sonner';

interface Props {
  companyId: string;
}

export default function CreateManagerForm({ companyId }: Props) {
  const [fullName, setFullName] = useState('');
  const [email, setEmail]       = useState('');
  const [password, setPassword] = useState('');
  const [employeeCode, setEmployeeCode] = useState('');
  const [submitting, setSubmitting] = useState(false);
  const [successInfo, setSuccessInfo] = useState<{ email: string; password: string } | null>(null);

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!fullName.trim() || !email.trim() || !password.trim()) {
      toast.error("Veuillez remplir tous les champs obligatoires.");
      return;
    }
    if (password.length < 6) {
      toast.error("Le mot de passe doit faire au moins 6 caractères.");
      return;
    }
    setSubmitting(true);

    try {
      const { user, error: authErr } = await createStaffAccount(email.trim().toLowerCase(), password);
      if (authErr) throw authErr;
      if (!user) throw new Error("Création du compte échouée.");

      const { error: profErr } = await insertStaffProfile({
        company_id:    companyId,
        full_name:     fullName.trim(),
        role:          'manager',
        employee_code: employeeCode.trim() || `MGR-${Date.now().toString().slice(-4)}`,
        auth_user_id:  user.id,
        is_active:     true,
      });
      if (profErr) throw profErr;

      setSuccessInfo({ email: email.trim().toLowerCase(), password });
      toast.success(`Manager ${fullName} créé avec succès !`);
      setFullName('');
      setEmail('');
      setPassword('');
      setEmployeeCode('');
    } catch (err: any) {
      toast.error(`Erreur: ${err.message || 'Échec de la création.'}`);
    } finally {
      setSubmitting(false);
    }
  };

  return (
    <div className="bg-white rounded-xl border border-slate-200 shadow-xs overflow-hidden">
      <div className="px-6 py-4 bg-slate-50 border-b border-slate-200">
        <div className="flex items-center gap-3">
          <div className="w-10 h-10 rounded-lg bg-blue-600 flex items-center justify-center text-white">
            <UserPlus size={20} />
          </div>
          <div>
            <h3 className="text-sm font-black text-slate-900 uppercase tracking-widest">Créer un Manager</h3>
            <p className="text-xs text-slate-500 mt-0.5">Recrutez un gestionnaire pour piloter les opérations</p>
          </div>
        </div>
      </div>

      <div className="p-6">
        {successInfo && (
          <div className="mb-5 p-4 rounded-lg bg-emerald-50 border border-emerald-200 flex items-start gap-3">
            <CheckCircle2 className="w-5 h-5 text-emerald-600 shrink-0 mt-0.5" />
            <div className="flex-1">
              <p className="text-sm font-black text-emerald-800">Compte Manager créé !</p>
              <p className="text-xs text-emerald-700 mt-1">
                Partagez ces identifiants : <span className="font-mono font-bold">{successInfo.email}</span> / <span className="font-mono font-bold">{successInfo.password}</span>
              </p>
              <button onClick={() => setSuccessInfo(null)} className="text-[10px] font-bold text-emerald-600 mt-2 hover:underline cursor-pointer">
                × Fermer
              </button>
            </div>
          </div>
        )}

        <form onSubmit={handleSubmit} className="space-y-4">
          <div>
            <label className="block text-xs font-bold text-slate-500 uppercase tracking-wider mb-1.5">
              Nom Complet *
            </label>
            <div className="relative">
              <User className="absolute left-3.5 top-3 w-4 h-4 text-slate-400" />
              <input
                type="text"
                required
                placeholder="ex: Amélie Martin"
                value={fullName}
                onChange={e => setFullName(e.target.value)}
                className="w-full text-sm rounded-lg border border-slate-200 pl-10 pr-3.5 py-2.5 outline-hidden focus:ring-2 focus:ring-blue-100 focus:border-blue-500 transition-all"
              />
            </div>
          </div>

          <div>
            <label className="block text-xs font-bold text-slate-500 uppercase tracking-wider mb-1.5">
              Email Professionnel *
            </label>
            <div className="relative">
              <Mail className="absolute left-3.5 top-3 w-4 h-4 text-slate-400" />
              <input
                type="email"
                required
                placeholder="manager@entreprise.com"
                value={email}
                onChange={e => setEmail(e.target.value)}
                className="w-full text-sm rounded-lg border border-slate-200 pl-10 pr-3.5 py-2.5 outline-hidden focus:ring-2 focus:ring-blue-100 focus:border-blue-500 transition-all"
              />
            </div>
          </div>

          <div>
            <label className="block text-xs font-bold text-slate-500 uppercase tracking-wider mb-1.5">
              Mot de passe Temporaire *
            </label>
            <div className="relative">
              <Lock className="absolute left-3.5 top-3 w-4 h-4 text-slate-400" />
              <input
                type="text"
                required
                placeholder="Min. 6 caractères"
                value={password}
                onChange={e => setPassword(e.target.value)}
                className="w-full text-sm rounded-lg border border-slate-200 pl-10 pr-3.5 py-2.5 outline-hidden focus:ring-2 focus:ring-blue-100 focus:border-blue-500 transition-all font-mono"
              />
            </div>
            <p className="text-[10px] text-slate-400 mt-1">Le manager pourra le changer plus tard.</p>
          </div>

          <div>
            <label className="block text-xs font-bold text-slate-500 uppercase tracking-wider mb-1.5">
              Code Employé (optionnel)
            </label>
            <input
              type="text"
              placeholder="ex: MGR-001"
              value={employeeCode}
              onChange={e => setEmployeeCode(e.target.value)}
              className="w-full text-sm rounded-lg border border-slate-200 px-3.5 py-2.5 outline-hidden focus:ring-2 focus:ring-blue-100 focus:border-blue-500 transition-all font-mono"
            />
          </div>

          <button
            type="submit"
            disabled={submitting}
            className="w-full flex justify-center items-center gap-2 bg-blue-600 hover:bg-blue-700 disabled:opacity-60 text-white font-bold py-3 px-4 rounded-lg transition-all shadow-sm cursor-pointer mt-2"
          >
            {submitting ? (
              <><Loader2 className="w-4 h-4 animate-spin" /> Création en cours...</>
            ) : (
              <><UserPlus className="w-4 h-4" /> Créer le Manager</>
            )}
          </button>
        </form>
      </div>
    </div>
  );
}