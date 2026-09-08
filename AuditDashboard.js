"use client";
import React from "react";
import { ArrowLeft, ClipboardCheck } from "lucide-react";

export default function AuditDashboard({ onBack }) {
  return (
    <div className="min-h-screen bg-[#0a0f1c] text-slate-300 font-sans flex items-center justify-center p-4 relative overflow-hidden">
       <div className="absolute inset-0 z-0">
         <div className="absolute top-1/2 left-1/2 transform -translate-x-1/2 -translate-y-1/2 w-[30rem] h-[30rem] bg-indigo-600/10 rounded-full blur-[120px]"></div>
       </div>
       
       <div className="relative z-10 max-w-lg w-full bg-[#121929]/80 backdrop-blur-xl border border-white/5 rounded-3xl p-10 text-center shadow-[0_0_40px_rgba(99,102,241,0.1)]">
         <div className="w-20 h-20 bg-indigo-500/10 border border-indigo-500/30 rounded-2xl flex items-center justify-center mx-auto mb-6">
            <ClipboardCheck className="w-10 h-10 text-indigo-400" />
         </div>
         <h1 className="text-3xl font-bold text-white mb-3 tracking-wide">Work in Progress</h1>
         <p className="text-slate-400 mb-8 leading-relaxed">The Audit Dashboard is currently under construction. Please check back later for updates.</p>
         
         <button onClick={onBack} className="inline-flex items-center gap-2 px-6 py-3 bg-white/5 hover:bg-white/10 text-white font-bold uppercase tracking-widest text-xs rounded-xl border border-white/10 transition-all">
            <ArrowLeft className="w-4 h-4" /> Return to Hub
         </button>
       </div>
    </div>
  );
}
