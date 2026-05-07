import React, { useState } from 'react';
import { supabase } from '../lib/supabase';
import { motion } from 'motion/react';
import { Cpu, Mail, Lock, AlertCircle } from 'lucide-react';

export function LoginPage() {
  const [email, setEmail] = useState('');
  const [password, setPassword] = useState('');
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const handleLogin = async (e: React.FormEvent) => {
    e.preventDefault();
    setLoading(true);
    setError(null);
    const { error } = await supabase.auth.signInWithPassword({ email, password });
    if (error) setError(error.message);
    setLoading(false);
  };

  const handleSignUp = async () => {
    setLoading(true);
    setError(null);
    const { error } = await supabase.auth.signUp({ email, password });
    if (error) {
      if (error.status === 400 && error.message.includes('already registered')) {
        setError("Email ini sudah terdaftar. Silakan gunakan menu Login.");
      } else {
        setError(error.message);
      }
    } else {
      setError("Check your email for verification link.");
    }
    setLoading(false);
  };

  return (
    <div className="min-h-screen bg-[#020617] flex items-center justify-center p-6 bg-[radial-gradient(circle_at_50%_50%,rgba(79,70,229,0.1),transparent)]">
      <motion.div 
        initial={{ opacity: 0, scale: 0.95 }}
        animate={{ opacity: 1, scale: 1 }}
        className="w-full max-w-md bg-slate-900 border border-slate-800 rounded-3xl overflow-hidden shadow-[0_0_50px_rgba(79,70,229,0.1)]"
      >
        <div className="p-10 border-b border-slate-800 bg-indigo-600/5 text-center">
          <div className="w-16 h-16 bg-indigo-600/10 rounded-2xl flex items-center justify-center mx-auto mb-6 border border-indigo-500/20">
            <Cpu className="w-8 h-8 text-indigo-500" />
          </div>
          <h1 className="text-3xl font-black text-white tracking-tighter uppercase italic">
            OM <span className="text-indigo-500">DEDY</span> <span className="text-slate-500 italic lowercase text-sm">v2.0</span>
          </h1>
          <p className="text-[10px] text-slate-500 mt-2 font-bold uppercase tracking-[0.2em]">Operational Management & Delivery Deployment Yield</p>
        </div>

        <form onSubmit={handleLogin} className="p-8 space-y-6">
          {error && (
            <div className="bg-rose-500/10 border border-rose-500/20 p-3 rounded-xl flex items-center gap-3 text-rose-400 text-xs">
              <AlertCircle className="w-4 h-4 shrink-0" />
              <span>{error}</span>
            </div>
          )}

          <div className="space-y-4">
            <div className="space-y-2">
              <label className="text-[10px] font-black text-slate-500 uppercase tracking-widest px-1">Access Identity</label>
              <div className="relative">
                <Mail className="absolute left-4 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-600" />
                <input 
                  type="email"
                  required
                  value={email}
                  onChange={(e) => setEmail(e.target.value)}
                  placeholder="name@company.com"
                  className="w-full bg-slate-950 border border-slate-800 rounded-xl pl-12 pr-4 py-3 text-slate-200 outline-none focus:border-indigo-500 transition-colors"
                />
              </div>
            </div>

            <div className="space-y-2">
              <label className="text-[10px] font-black text-slate-500 uppercase tracking-widest px-1">Crypto-Key</label>
              <div className="relative">
                <Lock className="absolute left-4 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-600" />
                <input 
                  type="password"
                  required
                  value={password}
                  onChange={(e) => setPassword(e.target.value)}
                  placeholder="••••••••"
                  className="w-full bg-slate-950 border border-slate-800 rounded-xl pl-12 pr-4 py-3 text-slate-200 outline-none focus:border-indigo-500 transition-colors"
                />
              </div>
            </div>
          </div>

          <div className="flex flex-col gap-3">
            <button 
              type="submit"
              disabled={loading}
              className="w-full py-4 bg-indigo-600 hover:bg-indigo-500 disabled:opacity-50 text-white font-black uppercase text-xs tracking-[0.2em] rounded-xl shadow-lg shadow-indigo-600/20 transition-all flex items-center justify-center gap-2"
            >
              {loading ? "Authenticating..." : "Authorize Access"}
            </button>
            <button 
              type="button"
              onClick={handleSignUp}
              disabled={loading}
              className="w-full py-3 text-slate-500 hover:text-slate-300 font-bold uppercase text-[10px] tracking-widest transition-colors"
            >
              Provision New Account
            </button>
          </div>
        </form>

        <div className="p-8 bg-slate-950/50 border-t border-slate-800/60 text-center">
          <p className="text-[9px] text-slate-600 leading-relaxed uppercase tracking-widest">
            Authorized Personnel Only • Dynamic Governance Protocol v2.5.1
          </p>
        </div>
      </motion.div>
    </div>
  );
}
