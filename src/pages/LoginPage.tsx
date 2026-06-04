import { useState, FormEvent } from 'react';
import { useNavigate, Link } from 'react-router-dom';
import { motion, AnimatePresence } from 'motion/react';
import { 
  Building2, Mail, Lock, ShieldAlert, ArrowRight, Compass, ShieldCheck 
} from 'lucide-react';
import { signIn, getCurrentProfile, isSupabaseConfigured } from '../lib/auth';
import { toast } from 'sonner';

export default function LoginPage() {
  const navigate = useNavigate();
  const [email, setEmail] = useState('');
  const [password, setPassword] = useState('');
  const [isSubmitting, setIsSubmitting] = useState(false);
  const [errorMessage, setErrorMessage] = useState<string | null>(null);

  const handleLogin = async (e: FormEvent) => {
    e.preventDefault();
    setErrorMessage(null);

    if (!email.trim() || !password.trim()) {
      setErrorMessage("Veuillez remplir tous les champs.");
      return;
    }

    setIsSubmitting(true);

    try {
      // 1. Sign In
      const { user, error: signErr } = await signIn(email.trim(), password.trim());

      if (signErr) {
        throw new Error(signErr.message || "Identifiants invalides.");
      }

      // 2. Fetch staff profile and role
      const { profile, error: profileErr } = await getCurrentProfile(user.id);

      if (profileErr) {
        throw new Error(profileErr.message || "Erreur de chargement du profil professionnel.");
      }

      if (!profile) {
        throw new Error("Aucun profil professionnel associé à ce compte.");
      }

      // 3. Redirections based on role
      if (profile.role === 'owner') {
        toast.success("Connexion réussie - Espace Propriétaire");
        navigate('/owner');
      } else if (profile.role === 'manager') {
        toast.success("Connexion réussie - Espace Manager");
        navigate('/manager');
      } else {
        // Any other role is restricted
        setErrorMessage("Ce portail n'est pas accessible avec ce rôle. Contactez votre manager.");
      }

    } catch (err: any) {
      console.error(err);
      setErrorMessage(err.message || "Impossible de se connecter.");
    } finally {
      setIsSubmitting(false);
    }
  };

  // Helper shortcut for testing in Demo/Mock Mode
  const handleQuickLogin = async (demoRole: 'owner' | 'manager') => {
    setErrorMessage(null);
    setIsSubmitting(true);
    try {
      const targetEmail = demoRole === 'owner' ? 'owner@logiflow.com' : 'manager@logiflow.com';
      const { user, error } = await signIn(targetEmail, 'password123');

      if (error) throw error;

      const { profile } = await getCurrentProfile(user.id);
      if (profile) {
        if (profile.role === 'owner') {
          toast.success("Connexion Démo - Espace Propriétaire");
          navigate('/owner');
        } else if (profile.role === 'manager') {
          toast.success("Connexion Démo - Espace Manager");
          navigate('/manager');
        }
      }
    } catch (err: any) {
      setErrorMessage(err.message);
    } finally {
      setIsSubmitting(false);
    }
  };

  return (
    <div className="min-h-screen bg-slate-50 flex flex-col justify-center py-12 px-4 sm:px-6 lg:px-8 relative overflow-hidden" id="login-page-root">
      
      {/* Decors */}
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
        
        <h2 className="text-center text-2xl font-extrabold text-slate-800 tracking-tight">
          Portail d'Accès ERP
        </h2>
        <p className="mt-1 text-center text-xs text-slate-500 max-w-sm mx-auto leading-normal">
          Système SaaS de gestion et de pilotage logistique multi-tenant
        </p>

        {!isSupabaseConfigured && (
          <div className="mt-3.5 p-3 rounded-lg bg-amber-50 border border-amber-100 text-amber-850 text-xs flex gap-2 items-center justify-center max-w-sm mx-auto">
            <span className="flex-shrink-0 inline-block w-1.5 h-1.5 rounded-full bg-amber-500 animate-pulse" />
            <span>Mode Démo actif (local). Idéal pour tester immédiatement.</span>
          </div>
        )}
      </div>

      <div className="mt-6 sm:mx-auto sm:w-full sm:max-w-md relative z-10">
        <div className="bg-white py-8 px-6 sm:px-10 rounded-xl border border-slate-200 shadow-sm">
          
          <form onSubmit={handleLogin} className="space-y-5" id="login-form">
            
            {/* Warning Alarm Banner */}
            <AnimatePresence mode="popLayout">
              {errorMessage && (
                <motion.div
                  initial={{ opacity: 0, height: 0 }}
                  animate={{ opacity: 1, height: 'auto' }}
                  exit={{ opacity: 0, height: 0 }}
                  className="p-3.5 rounded-lg bg-rose-50 border border-rose-100 text-rose-800 text-xs flex gap-2 items-start"
                  id="login-error-banner"
                >
                  <ShieldAlert className="w-4 h-4 text-rose-600 flex-shrink-0 mt-0.5" />
                  <div>
                    <span className="font-semibold block mb-0.5">Restriction d'Accès</span>
                    <span>{errorMessage}</span>
                  </div>
                </motion.div>
              )}
            </AnimatePresence>

            <div>
              <label className="block text-xs font-bold text-slate-500 uppercase tracking-wider mb-1.5">
                Adresse Email
              </label>
              <div className="relative">
                <Mail className="absolute left-3.5 top-3.5 w-4 h-4 text-slate-400" />
                <input
                  type="email"
                  required
                  placeholder="nom@entreprise.com"
                  value={email}
                  onChange={(e) => setEmail(e.target.value)}
                  className="w-full text-sm rounded-lg border border-slate-200 pl-10 pr-3.5 py-3 outline-hidden focus:ring-2 focus:ring-blue-100 focus:border-blue-500 transition-all text-slate-800"
                  id="login-email-input"
                />
              </div>
            </div>

            <div>
              <div className="flex justify-between mb-1.5">
                <label className="block text-xs font-bold text-slate-500 uppercase tracking-wider">
                  Mot de passe
                </label>
              </div>
              <div className="relative">
                <Lock className="absolute left-3.5 top-3.5 w-4 h-4 text-slate-400" />
                <input
                  type="password"
                  required
                  placeholder="••••••••"
                  value={password}
                  onChange={(e) => setPassword(e.target.value)}
                  className="w-full text-sm rounded-lg border border-slate-200 pl-10 pr-3.5 py-3 outline-hidden focus:ring-2 focus:ring-blue-100 focus:border-blue-500 transition-all text-slate-800"
                  id="login-password-input"
                />
              </div>
            </div>

            <button
              type="submit"
              disabled={isSubmitting}
              className="w-full flex justify-center items-center gap-2 bg-blue-600 hover:bg-blue-700 text-white font-bold py-3 px-4 rounded-lg transition-all shadow-sm cursor-pointer mt-2"
              id="login-submit-btn"
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

          {/* Quick Login Section for Demo Mode */}
          {!isSupabaseConfigured && (
            <div className="mt-8 pt-6 border-t border-slate-150">
              <span className="block text-center text-[10px] font-bold uppercase tracking-widest text-slate-400 mb-3.5">
                Raccourcis de Test (Mode Démo)
              </span>
              <div className="grid grid-cols-2 gap-3">
                <button
                  type="button"
                  onClick={() => handleQuickLogin('owner')}
                  className="flex flex-col items-center p-3 rounded-lg border border-slate-200 bg-slate-50 hover:bg-blue-50 hover:border-blue-200 transition-all text-left cursor-pointer"
                  id="quick-owner-login-btn"
                >
                  <ShieldCheck className="w-5 h-5 text-blue-600 mb-1" />
                  <span className="text-xs font-bold text-slate-850">Propriétaire</span>
                  <span className="text-[9px] text-slate-400 mt-0.5 font-medium">Jean Renaud</span>
                </button>

                <button
                  type="button"
                  onClick={() => handleQuickLogin('manager')}
                  className="flex flex-col items-center p-3 rounded-lg border border-slate-200 bg-slate-50 hover:bg-blue-50 hover:border-blue-200 transition-all text-left cursor-pointer"
                  id="quick-manager-login-btn"
                >
                  <Compass className="w-5 h-5 text-emerald-600 mb-1" />
                  <span className="text-xs font-bold text-slate-850">Manager</span>
                  <span className="text-[9px] text-slate-400 mt-0.5 font-medium">Amélie Martin</span>
                </button>
              </div>
            </div>
          )}

        </div>
      </div>
    </div>
  );
}
