"use client";
import React from "react";
import { ArrowLeft, Users } from "lucide-react";

export default function TaskIndiaDashboard({ onBack }) {
  return (
    <div className="min-h-screen bg-[#0a0f1c] text-slate-300 font-sans flex items-center justify-center p-4 relative overflow-hidden">
       <div className="absolute inset-0 z-0">
         <div className="absolute top-1/2 left-1/2 transform -translate-x-1/2 -translate-y-1/2 w-[30rem] h-[30rem] bg-amber-600/10 rounded-full blur-[120px]"></div>
       </div>
       
       <div className="relative z-10 max-w-lg w-full bg-[#121929]/80 backdrop-blur-xl border border-white/5 rounded-3xl p-10 text-center shadow-[0_0_40px_rgba(245,158,11,0.1)]">
         <div className="w-20 h-20 bg-amber-500/10 border border-amber-500/30 rounded-2xl flex items-center justify-center mx-auto mb-6">
            <Users className="w-10 h-10 text-amber-400" />
         </div>
         <h1 className="text-3xl font-bold text-white mb-3 tracking-wide">Work in Progress</h1>
         <p className="text-slate-400 mb-8 leading-relaxed">The Task Management (Dawateislami India Zimmedaran) Dashboard is currently under construction.</p>
         
         <button onClick={onBack} className="inline-flex items-center gap-2 px-6 py-3 bg-white/5 hover:bg-white/10 text-white font-bold uppercase tracking-widest text-xs rounded-xl border border-white/10 transition-all">
            <ArrowLeft className="w-4 h-4" /> Return to Hub
         </button>
       </div>
    </div>
  );
}