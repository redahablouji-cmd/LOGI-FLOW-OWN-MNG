import { useState, FormEvent } from 'react';
import { useNavigate, Link } from 'react-router-dom';
import { motion, AnimatePresence } from 'motion/react';
import { 
  Building2, User, Mail, Lock, CheckCircle2, Copy, Check, ArrowRight,
  Phone, MapPin, Truck, Star
} from 'lucide-react';
import { registerCompanyAndOwner, createStaffAccount, isSupabaseConfigured } from '../lib/auth';
import { toast } from 'sonner';

export default function RegisterCompanyPage() {
  const navigate = useNavigate();

  // Company fields
  const [companyName, setCompanyName]           = useState('');
  const [phone, setPhone]                       = useState('');
  const [city, setCity]                         = useState('');
  const [fleetSize, setFleetSize]               = useState('');
  const [subscriptionPlan, setSubscriptionPlan] = useState('starter');

  // Owner fields
  const [ownerName, setOwnerName]               = useState('');
  const [email, setEmail]                       = useState('');
  const [password, setPassword]                 = useState('');
  const [confirmPassword, setConfirmPassword]   = useState('');

  const [isSubmitting, setIsSubmitting]         = useState(false);
  const [registeredCode, setRegisteredCode]     = useState<string | null>(null);
  const [copiedCode, setCopiedCode]             = useState(false);

  const handleCopyCode = async () => {
    if (!registeredCode) return;
    try {
      await navigator.clipboard.writeText(registeredCode);
      setCopiedCode(true);
      setTimeout(() => setCopiedCode(false), 2000);
    } catch (err) {
      console.error('Could not copy:', err);
    }
  };

  const generateBusinessCode = (name: string): string => {
    const clean = name.replace(/[^a-zA-Z]/g, '').toUpperCase();
    const prefix = clean.substring(0, 3).padEnd(3, 'X');
    const randomDigits = Math.floor(1000 + Math.random() * 9000);
    return `LF-${prefix}-${randomDigits}`;
  };

  const handleSubmit = async (e: FormEvent) => {
    e.preventDefault();

    if (!companyName.trim() || !ownerName.trim() || !email.trim() || !password.trim()) {
      toast.error("Veuillez remplir tous les champs obligatoires.");
      return;
    }
    if (password !== confirmPassword) {
      toast.error("Les mots de passe ne correspondent pas.");
      return;
    }
    if (password.length < 6) {
      toast.error("Le mot de passe doit contenir au moins 6 caractères.");
      return;
    }

    setIsSubmitting(true);

    try {
      const code = generateBusinessCode(companyName);

      const { user: authUser, error: authErr } = await createStaffAccount(email.trim(), password.trim());
      if (authErr) throw new Error(authErr.message || "Impossible de créer le compte utilisateur.");

      const { error: dbErr } = await registerCompanyAndOwner({
        companyName:      companyName.trim(),
        ownerName:        ownerName.trim(),
        email:            email.trim(),
        authUserId:       authUser.id,
        businessCode:     code,
        phone:            phone.trim(),
        city:             city.trim(),
        fleetSize:        fleetSize ? parseInt(fleetSize) : null,
        subscriptionPlan: subscriptionPlan,
      });

      if (dbErr) throw new Error(dbErr.message || "Erreur d'écriture en base de données.");

      setRegisteredCode(code);
      toast.success("Entreprise enregistrée !");

    } catch (err: any) {
      console.error(err);
      toast.error(`Erreur: ${err.message || 'Problème technique lors de la création.'}`);
    } finally {
      setIsSubmitting(false);
    }
  };

  const inputClass = "w-full text-sm rounded-lg border border-slate-200 pl-10 pr-3.5 py-3 outline-hidden focus:ring-2 focus:ring-blue-100 focus:border-blue-500 transition-all text-slate-800";
  const labelClass = "block text-xs font-bold text-slate-500 uppercase tracking-wider mb-1.5";

  return (
    <div className="min-h-screen bg-slate-50 flex flex-col justify-center py-12 px-4 sm:px-6 lg:px-8 relative overflow-hidden">
      <div className="absolute inset-x-0 top-0 h-40 bg-gradient-to-b from-blue-50/50 to-transparent pointer-events-none" />
      <div className="absolute top-1/4 -left-20 w-80 h-80 bg-blue-100/20 rounded-full blur-3xl pointer-events-none" />
      <div className="absolute bottom-10 -right-20 w-80 h-80 bg-slate-100/40 rounded-full blur-3xl pointer-events-none" />

      <div className="sm:mx-auto sm:w-full sm:max-w-md relative z-10 text-center">
        <div className="inline-flex items-center justify-center space-x-2 mb-3">
          <div className="bg-blue-600 h-9 w-9 rounded flex items-center justify-center shadow-md shadow-blue-500/20">
            <span className="text-white font-extrabold text-lg">LF</span>
          </div>
          <span className="text-xl font-black tracking-tight text-slate-900">LOGI-FLOW</span>
        </div>
        <h2 className="text-2xl font-extrabold text-slate-900 tracking-tight">Nouveau Compte Propriétaire</h2>
        <p className="mt-1 text-xs text-slate-500 max-w-sm mx-auto leading-normal">
          Inscrivez et configurez votre entreprise logistique sur LOGI-FLOW ERP
        </p>
        {!isSupabaseConfigured && (
          <div className="mt-3.5 p-3 rounded-lg bg-amber-50 border border-amber-100 text-xs flex gap-2 items-center justify-center max-w-sm mx-auto">
            <span className="flex-shrink-0 inline-block w-1.5 h-1.5 rounded-full bg-amber-500 animate-pulse" />
            <span>Mode Démo actif (local). Création d'entreprise instantanée.</span>
          </div>
        )}
      </div>

      <div className="mt-6 sm:mx-auto sm:w-full sm:max-w-2xl relative z-10">
        <div className="bg-white py-8 px-6 sm:px-10 rounded-xl border border-slate-200 shadow-sm">
          <AnimatePresence mode="wait">
            {!registeredCode ? (
              <motion.div
                key="register-form"
                initial={{ opacity: 0, x: -10 }}
                animate={{ opacity: 1, x: 0 }}
                exit={{ opacity: 0, x: 10 }}
                transition={{ duration: 0.3 }}
              >
                <form onSubmit={handleSubmit} className="space-y-5">

                  {/* ── Section: Informations Entreprise ── */}
                  <div>
                    <p className="text-[10px] font-black text-slate-400 uppercase tracking-widest mb-3 border-b border-slate-100 pb-2">
                      Informations de l'entreprise
                    </p>
                    <div className="space-y-4">

                      {/* Company Name (full width) */}
                      <div>
                        <label className={labelClass}>Nom de l'entreprise *</label>
                        <div className="relative">
                          <Building2 className="absolute left-3.5 top-3.5 w-4 h-4 text-slate-400" />
                          <input
                            type="text" required
                            placeholder="Ex: Transit Fast"
                            value={companyName}
                            onChange={(e) => setCompanyName(e.target.value)}
                            className={inputClass}
                          />
                        </div>
                      </div>

                      {/* Phone + City */}
                      <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                        <div>
                          <label className={labelClass}>Téléphone</label>
                          <div className="relative">
                            <Phone className="absolute left-3.5 top-3.5 w-4 h-4 text-slate-400" />
                            <input
                              type="tel"
                              placeholder="Ex: 0661234567"
                              value={phone}
                              onChange={(e) => setPhone(e.target.value)}
                              className={inputClass}
                            />
                          </div>
                        </div>
                        <div>
                          <label className={labelClass}>Ville</label>
                          <div className="relative">
                            <MapPin className="absolute left-3.5 top-3.5 w-4 h-4 text-slate-400" />
                            <input
                              type="text"
                              placeholder="Ex: Casablanca"
                              value={city}
                              onChange={(e) => setCity(e.target.value)}
                              className={inputClass}
                            />
                          </div>
                        </div>
                      </div>

                      {/* Fleet Size + Subscription Plan */}
                      <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                        <div>
                          <label className={labelClass}>Taille de la flotte</label>
                          <div className="relative">
                            <Truck className="absolute left-3.5 top-3.5 w-4 h-4 text-slate-400" />
                            <input
                              type="number" min="1"
                              placeholder="Ex: 12"
                              value={fleetSize}
                              onChange={(e) => setFleetSize(e.target.value)}
                              className={inputClass}
                            />
                          </div>
                        </div>
                        <div>
                          <label className={labelClass}>Plan d'abonnement</label>
                          <div className="relative">
                            <Star className="absolute left-3.5 top-3.5 w-4 h-4 text-slate-400" />
                            <select
                              value={subscriptionPlan}
                              onChange={(e) => setSubscriptionPlan(e.target.value)}
                              className={`${inputClass} bg-white appearance-none`}
                            >
                              <option value="starter">Starter</option>
                              <option value="pro">Pro</option>
                              <option value="enterprise">Enterprise</option>
                            </select>
                          </div>
                        </div>
                      </div>
                    </div>
                  </div>

                  {/* ── Section: Informations Propriétaire ── */}
                  <div>
                    <p className="text-[10px] font-black text-slate-400 uppercase tracking-widest mb-3 border-b border-slate-100 pb-2">
                      Compte du Propriétaire
                    </p>
                    <div className="space-y-4">

                      {/* Owner Name (full width) */}
                      <div>
                        <label className={labelClass}>Nom complet du Propriétaire *</label>
                        <div className="relative">
                          <User className="absolute left-3.5 top-3.5 w-4 h-4 text-slate-400" />
                          <input
                            type="text" required
                            placeholder="Ex: Redah Ablouji"
                            value={ownerName}
                            onChange={(e) => setOwnerName(e.target.value)}
                            className={inputClass}
                          />
                        </div>
                      </div>

                      {/* Email (full width) */}
                      <div>
                        <label className={labelClass}>Adresse Email Professionnelle *</label>
                        <div className="relative">
                          <Mail className="absolute left-3.5 top-3.5 w-4 h-4 text-slate-400" />
                          <input
                            type="email" required
                            placeholder="Ex: owner@transitfast.com"
                            value={email}
                            onChange={(e) => setEmail(e.target.value)}
                            className={inputClass}
                          />
                        </div>
                      </div>

                      {/* Password + Confirm */}
                      <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                        <div>
                          <label className={labelClass}>Mot de passe *</label>
                          <div className="relative">
                            <Lock className="absolute left-3.5 top-3.5 w-4 h-4 text-slate-400" />
                            <input
                              type="password" required
                              placeholder="••••••••"
                              value={password}
                              onChange={(e) => setPassword(e.target.value)}
                              className={inputClass}
                            />
                          </div>
                        </div>
                        <div>
                          <label className={labelClass}>Confirmer le mot de passe *</label>
                          <div className="relative">
                            <Lock className="absolute left-3.5 top-3.5 w-4 h-4 text-slate-400" />
                            <input
                              type="password" required
                              placeholder="••••••••"
                              value={confirmPassword}
                              onChange={(e) => setConfirmPassword(e.target.value)}
                              className={inputClass}
                            />
                          </div>
                        </div>
                      </div>
                    </div>
                  </div>

                  <button
                    type="submit"
                    disabled={isSubmitting}
                    className="w-full flex justify-center items-center gap-2 bg-blue-600 hover:bg-blue-700 disabled:opacity-60 text-white font-bold py-3 px-4 rounded-lg transition-all shadow-sm cursor-pointer mt-2"
                  >
                    {isSubmitting ? "Création du compte..." : "Enregistrer mon entreprise"}
                  </button>

                  <div className="text-center pt-4 border-t border-slate-100">
                    <span className="text-xs text-slate-500">Déjà inscrit ?</span>
                    <Link to="/login" className="text-xs font-bold text-blue-600 hover:text-blue-700 ml-1.5 transition-colors">
                      Se connecter au portail
                    </Link>
                  </div>
                </form>
              </motion.div>
            ) : (
              <motion.div
                key="register-success"
                initial={{ opacity: 0, scale: 0.95 }}
                animate={{ opacity: 1, scale: 1 }}
                transition={{ duration: 0.4 }}
                className="text-center py-4"
              >
                <div className="mx-auto h-14 w-14 rounded bg-emerald-100 flex items-center justify-center text-emerald-600 mb-4">
                  <CheckCircle2 className="w-8 h-8" />
                </div>
                <h3 className="text-2xl font-extrabold text-slate-900 tracking-tight">Bienvenue sur LOGI-FLOW !</h3>
                <p className="text-xs text-slate-500 mt-2 max-w-sm mx-auto leading-relaxed">
                  Votre entreprise a été créée avec succès. Voici votre code entreprise confidentiel :
                </p>
                <div className="mt-6 max-w-md mx-auto text-left">
                  <label className="text-[10px] font-bold text-slate-400 uppercase tracking-widest block mb-1">
                    Code d'entreprise (Business Code)
                  </label>
                  <div className="flex items-center justify-between bg-slate-900 text-white rounded-lg p-3 border border-slate-800">
                    <code className="font-mono font-bold text-lg">{registeredCode}</code>
                    <button
                      onClick={handleCopyCode}
                      className={`p-1.5 rounded transition-colors text-slate-400 hover:text-white cursor-pointer ${copiedCode ? 'text-emerald-400' : 'hover:bg-slate-800'}`}
                    >
                      {copiedCode ? <Check className="w-4 h-4" /> : <Copy className="w-4 h-4" />}
                    </button>
                  </div>
                  <p className="text-[11px] text-slate-500 mt-2 leading-tight italic">
                    Partagez ce code avec vos managers pour qu'ils puissent rejoindre votre entreprise.
                  </p>
                </div>
                <div className="mt-8 max-w-sm mx-auto">
                  <button
                    onClick={() => navigate('/owner')}
                    className="w-full inline-flex items-center justify-center gap-2 bg-blue-600 hover:bg-blue-700 text-white font-bold py-3.5 px-4 rounded-lg transition-all cursor-pointer text-sm"
                  >
                    Accéder à mon tableau de bord <ArrowRight className="w-4 h-4" />
                  </button>
                </div>
              </motion.div>
            )}
          </AnimatePresence>
        </div>
      </div>
    </div>
  );
}