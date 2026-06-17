import { useState, FormEvent } from 'react';
import { useNavigate, Link } from 'react-router-dom';
import { motion, AnimatePresence } from 'motion/react';
import { 
  Mail, Lock, ShieldAlert, Building2, ShieldCheck, Compass
} from 'lucide-react';
import { signIn, getCurrentProfile, isSupabaseConfigured, supabase } from '../lib/auth';
import { toast } from 'sonner';

export default function LoginPage() {
  const navigate = useNavigate();
  const [identifier, setIdentifier]     = useState('');
  const [password, setPassword]         = useState('');
  const [isSubmitting, setIsSubmitting] = useState(false);
  const [errorMessage, setErrorMessage] = useState<string | null>(null);

  // Detect if input is a business code (no @ symbol)
  const isBusinessCode = !identifier.includes('@');

 const handleLogin = async (e: FormEvent) => {
  e.preventDefault();
  setErrorMessage(null);
  if (!identifier.trim() || !password.trim()) {
    setErrorMessage("Veuillez remplir tous les champs.");
    return;
  }
  setIsSubmitting(true);
  try {
    let emailToUse = identifier.trim();
    let isOwner = false;

    // Business code → owner login
    if (isBusinessCode) {
      const { data: company, error: companyErr } = await supabase
        .from('companies')
        .select('id, owner_email')
        .ilike('business_code', identifier.trim())
        .maybeSingle();

      if (companyErr || !company) {
        throw new Error("Code entreprise invalide.");
      }
      emailToUse = company.owner_email;
      isOwner = true;
    }

    // Sign in
    const { user, error: signErr } = await signIn(emailToUse, password.trim());
    if (signErr) throw new Error("Identifiants invalides.");

    // Wait for session to be fully persisted
    await new Promise(resolve => setTimeout(resolve, 300));

    // Route based on how they logged in
    if (isOwner) {
      toast.success("Connexion réussie — Espace Propriétaire");
      navigate('/owner');
    } else {
      // Email login → check staff_profiles for managers
      const { profile, error: profileErr } = await getCurrentProfile(user.id);
      if (profileErr) throw new Error("Erreur de chargement du profil.");
      if (!profile) throw new Error("Aucun profil associé à ce compte.");

      if (profile.role === 'manager') {
        toast.success("Connexion réussie — Espace Manager");
        navigate('/manager');
      } else {
        setErrorMessage("Ce rôle n'a pas accès à ce portail.");
      }
    }
  } catch (err: any) {
    setErrorMessage(err.message || "Impossible de se connecter.");
  } finally {
    setIsSubmitting(false);
  }
};

  return (
    <div className="min-h-screen bg-slate-50 flex flex-col justify-center py-12 px-4 sm:px-6 lg:px-8 relative overflow-hidden">
      <div className="absolute inset-x-0 top-0 h-40 bg-gradient-to-b from-blue-50/50 to-transparent pointer-events-none" />
      <div className="absolute top-1/3 -right-20 w-80 h-80 bg-blue-100/20 rounded-full blur-3xl pointer-events-none" />
      <div className="absolute bottom-5 -left-20 w-80 h-80 bg-slate-100/40 rounded-full blur-3xl pointer-events-none" />

      <div className="sm:mx-auto sm:w-full sm:max-w-md relative z-10 text-center">
        <div className="inline-flex items-center justify-center space-x-2 mb-3">
          <div className="bg-blue-600 h-9 w-9 rounded flex items-center justify-center shadow-md shadow-blue-500/20">
            <span className="text-white font-extrabold text-lg">LF</span>
          </div>
          <span className="text-xl font-black tracking-tight text-slate-900">LOGI-FLOW</span>
        </div>
        <h2 className="text-2xl font-extrabold text-slate-800 tracking-tight">Portail d'Accès ERP</h2>
        <p className="mt-1 text-xs text-slate-500 max-w-sm mx-auto leading-normal">
          Système SaaS de gestion et de pilotage logistique multi-tenant
        </p>
        {!isSupabaseConfigured && (
          <div className="mt-3.5 p-3 rounded-lg bg-amber-50 border border-amber-100 text-xs flex gap-2 items-center justify-center max-w-sm mx-auto">
            <span className="flex-shrink-0 inline-block w-1.5 h-1.5 rounded-full bg-amber-500 animate-pulse" />
            <span>Mode Démo actif (local).</span>
          </div>
        )}
      </div>

      <div className="mt-6 sm:mx-auto sm:w-full sm:max-w-md relative z-10">
        <div className="bg-white py-8 px-6 sm:px-10 rounded-xl border border-slate-200 shadow-sm">
          <form onSubmit={handleLogin} className="space-y-5">

            <AnimatePresence mode="popLayout">
              {errorMessage && (
                <motion.div
                  initial={{ opacity: 0, height: 0 }}
                  animate={{ opacity: 1, height: 'auto' }}
                  exit={{ opacity: 0, height: 0 }}
                  className="p-3.5 rounded-lg bg-rose-50 border border-rose-100 text-rose-800 text-xs flex gap-2 items-start"
                >
                  <ShieldAlert className="w-4 h-4 text-rose-600 flex-shrink-0 mt-0.5" />
                  <div>
                    <span className="font-semibold block mb-0.5">Restriction d'Accès</span>
                    <span>{errorMessage}</span>
                  </div>
                </motion.div>
              )}
            </AnimatePresence>

            {/* Identifier field — email or business code */}
            <div>
              <label className="block text-xs font-bold text-slate-500 uppercase tracking-wider mb-1.5">
                Email ou Code Entreprise
              </label>
              <div className="relative">
                {isBusinessCode && identifier.length > 0
                  ? <Building2 className="absolute left-3.5 top-3.5 w-4 h-4 text-blue-500" />
                  : <Mail className="absolute left-3.5 top-3.5 w-4 h-4 text-slate-400" />
                }
                <input
                  type="text"
                  required
                  placeholder="nom@entreprise.com ou LF-FOT-1234"
                  value={identifier}
                  onChange={(e) => setIdentifier(e.target.value)}
                  className="w-full text-sm rounded-lg border border-slate-200 pl-10 pr-3.5 py-3 outline-hidden focus:ring-2 focus:ring-blue-100 focus:border-blue-500 transition-all text-slate-800"
                />
              </div>
              {/* Live hint */}
              <AnimatePresence>
                {identifier.length > 2 && (
                  <motion.p
                    initial={{ opacity: 0, y: -4 }}
                    animate={{ opacity: 1, y: 0 }}
                    exit={{ opacity: 0 }}
                    className="text-[10px] mt-1 ml-1 font-semibold"
                  >
                    {isBusinessCode
                      ? <span className="text-blue-600">🏢 Connexion Propriétaire via code entreprise</span>
                      : <span className="text-emerald-600">👤 Connexion Manager via email</span>
                    }
                  </motion.p>
                )}
              </AnimatePresence>
            </div>

            {/* Password */}
            <div>
              <label className="block text-xs font-bold text-slate-500 uppercase tracking-wider mb-1.5">
                Mot de passe
              </label>
              <div className="relative">
                <Lock className="absolute left-3.5 top-3.5 w-4 h-4 text-slate-400" />
                <input
                  type="password"
                  required
                  placeholder="••••••••"
                  value={password}
                  onChange={(e) => setPassword(e.target.value)}
                  className="w-full text-sm rounded-lg border border-slate-200 pl-10 pr-3.5 py-3 outline-hidden focus:ring-2 focus:ring-blue-100 focus:border-blue-500 transition-all text-slate-800"
                />
              </div>
            </div>

            <button
              type="submit"
              disabled={isSubmitting}
              className="w-full flex justify-center items-center gap-2 bg-blue-600 hover:bg-blue-700 disabled:opacity-60 text-white font-bold py-3 px-4 rounded-lg transition-all shadow-sm cursor-pointer mt-2"
            >
              {isSubmitting ? "Connexion en cours..." : "Se connecter"}
            </button>

            <div className="text-center pt-4 border-t border-slate-100">
              <span className="text-xs text-slate-500">Nouvelle entreprise ?</span>
              <Link to="/register" className="text-xs font-bold text-blue-600 hover:text-blue-700 ml-1.5 transition-colors">
                Créer un compte
              </Link>
            </div>
          </form>

          {/* Demo shortcuts */}
          {!isSupabaseConfigured && (
            <div className="mt-8 pt-6 border-t border-slate-100">
              <span className="block text-center text-[10px] font-bold uppercase tracking-widest text-slate-400 mb-3.5">
                Raccourcis de Test (Mode Démo)
              </span>
              <div className="grid grid-cols-2 gap-3">
                <button
                  type="button"
                  onClick={() => handleQuickLogin('owner')}
                  className="flex flex-col items-center p-3 rounded-lg border border-slate-200 bg-slate-50 hover:bg-blue-50 hover:border-blue-200 transition-all cursor-pointer"
                >
                  <ShieldCheck className="w-5 h-5 text-blue-600 mb-1" />
                  <span className="text-xs font-bold text-slate-800">Propriétaire</span>
                  <span className="text-[9px] text-slate-400 mt-0.5">Jean Renaud</span>
                </button>
                <button
                  type="button"
                  onClick={() => handleQuickLogin('manager')}
                  className="flex flex-col items-center p-3 rounded-lg border border-slate-200 bg-slate-50 hover:bg-blue-50 hover:border-blue-200 transition-all cursor-pointer"
                >
                  <Compass className="w-5 h-5 text-emerald-600 mb-1" />
                  <span className="text-xs font-bold text-slate-800">Manager</span>
                  <span className="text-[9px] text-slate-400 mt-0.5">Amélie Martin</span>
                </button>
              </div>
            </div>
          )}
        </div>
      </div>
    </div>
  );
}