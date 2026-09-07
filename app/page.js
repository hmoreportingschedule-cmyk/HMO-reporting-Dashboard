"use client";
import React, { useState, useEffect } from "react";
import { Lock, User, LogOut, Network, ArrowRight, Loader2, Building2, BookOpen, FileText } from "lucide-react";
import DeeniKaamDashboard from "../components/DeeniKaamDashboard";
import DepartmentDashboard from "../components/DepartmentDashboard";
import WeeklyRisalaDashboard from "../components/WeeklyRisalaDashboard";

const ADMIN_USER = "admin";
const ADMIN_PASS = "admin123";
// ⚠️ IMPORTANT: YAHAN APNA NAYA GOOGLE APPS SCRIPT URL PASTE KAREIN JO STEP 2 ME MILA HAI
const SCRIPT_URL = "PASTE_YOUR_NEW_SCRIPT_URL_HERE";

export default function App() {
  const [isAuthenticated, setIsAuthenticated] = useState(false);
  const [isAuthenticating, setIsAuthenticating] = useState(false);
  const [username, setUsername] = useState("");
  const [password, setPassword] = useState("");
  const [loginError, setLoginError] = useState("");
  
  const [currentView, setCurrentView] = useState("hub");
  const [accessiblePortals, setAccessiblePortals] = useState([]);
  const [risalaToken, setRisalaToken] = useState("");
  const [risalaUser, setRisalaUser] = useState({});

  useEffect(() => {
    const authType = sessionStorage.getItem("hmo_auth_type");
    if (authType === "superadmin") {
        setAccessiblePortals(["deeni", "department", "weekly_risala"]);
        setIsAuthenticated(true);
    } else if (authType === "user") {
        const token = sessionStorage.getItem("risalaToken");
        const user = JSON.parse(sessionStorage.getItem("risalaUser") || "{}");
        setRisalaToken(token);
        setRisalaUser(user);
        setAccessiblePortals(["weekly_risala"]);
        setIsAuthenticated(true);
        setCurrentView("weekly_risala");
    }
  }, []);

  const handleLogin = async (e) => {
    e.preventDefault();
    setIsAuthenticating(true);
    setLoginError("");

    // Check Super Admin First
    if (username === ADMIN_USER && password === ADMIN_PASS) {
        setTimeout(() => {
            setIsAuthenticated(true);
            sessionStorage.setItem("hmo_auth_type", "superadmin");
            setAccessiblePortals(["deeni", "department", "weekly_risala"]);
            setCurrentView("hub");
            setIsAuthenticating(false);
        }, 1000);
        return;
    }

    // Check Google Script Users (e.g., Weekly Risala Users)
    try {
        const res = await fetch(`${SCRIPT_URL}?action=login&userId=${encodeURIComponent(username)}&password=${encodeURIComponent(password)}`);
        const d = await res.json();
        if (d.status === "Success") {
            sessionStorage.setItem("hmo_auth_type", "user");
            sessionStorage.setItem("risalaToken", d.token);
            sessionStorage.setItem("risalaUser", JSON.stringify(d.user));
            setRisalaToken(d.token);
            setRisalaUser(d.user);
            setAccessiblePortals(["weekly_risala"]);
            setIsAuthenticated(true);
            setCurrentView("weekly_risala"); // Auto Redirect since they only have 1 portal
        } else {
            setLoginError(d.message || "Access Denied. Invalid credentials.");
        }
    } catch (err) {
        setLoginError("Network Error. Cannot connect to server.");
    }
    setIsAuthenticating(false);
  };

  const handleLogout = () => {
    setIsAuthenticated(false);
    sessionStorage.clear();
    setUsername("");
    setPassword("");
    setCurrentView("hub");
    setAccessiblePortals([]);
  };

  if (!isAuthenticated) {
    return (
      <div className="min-h-screen flex items-center justify-center bg-[#0a0f1c] relative overflow-hidden font-sans">
        <div className="absolute inset-0 z-0">
          <div className="absolute top-[20%] left-[20%] w-[30rem] h-[30rem] bg-cyan-600/20 rounded-full blur-[120px] mix-blend-screen"></div>
          <div className="absolute bottom-[20%] right-[20%] w-[30rem] h-[30rem] bg-emerald-600/10 rounded-full blur-[120px] mix-blend-screen"></div>
          <div className="absolute inset-0 bg-[linear-gradient(rgba(255,255,255,0.02)_1px,transparent_1px),linear-gradient(90deg,rgba(255,255,255,0.02)_1px,transparent_1px)] bg-[size:40px_40px] opacity-20"></div>
        </div>
        <div className="relative z-10 w-full max-w-lg p-8 sm:p-12 bg-[#121929]/70 backdrop-blur-2xl border border-white/5 rounded-3xl shadow-[0_0_60px_rgba(6,182,212,0.1)]">
          <div className="flex flex-col items-center justify-center mb-10 text-center">
            <div className="w-16 h-16 bg-gradient-to-br from-cyan-400 to-blue-600 rounded-2xl flex items-center justify-center mb-6 shadow-[0_0_30px_rgba(34,211,238,0.4)] border border-white/20">
              <Network className="w-8 h-8 text-white" />
            </div>
            <h1 className="text-3xl font-bold text-white tracking-wide">Dawateislami <span className="text-cyan-400 font-light">India</span></h1>
            <p className="text-slate-400 text-sm mt-2 font-medium tracking-widest uppercase">Hind Mushawarat Office</p>
          </div>
          {loginError && <div className="mb-6 p-4 bg-red-900/40 border border-red-500/50 text-red-400 text-sm rounded-xl text-center backdrop-blur-sm">{loginError}</div>}
          <form onSubmit={handleLogin} className="space-y-6">
            <div className="space-y-2">
              <label className="text-xs font-bold text-cyan-400 uppercase tracking-widest ml-1">User ID</label>
              <div className="relative group">
                <div className="absolute inset-y-0 left-0 pl-4 flex items-center pointer-events-none">
                  <User className="w-5 h-5 text-slate-500 group-focus-within:text-cyan-400 transition-colors" />
                </div>
                <input type="text" required value={username} onChange={(e) => setUsername(e.target.value)} placeholder="Enter system identity" className="w-full pl-12 pr-4 py-4 bg-[#0a0f1c]/80 border border-slate-700/50 rounded-xl text-white placeholder-slate-600 focus:bg-[#0a0f1c] focus:border-cyan-500/50 focus:ring-1 focus:ring-cyan-500/50 outline-none transition-all" />
              </div>
            </div>
            <div className="space-y-2">
              <label className="text-xs font-bold text-cyan-400 uppercase tracking-widest ml-1">Password</label>
              <div className="relative group">
                <div className="absolute inset-y-0 left-0 pl-4 flex items-center pointer-events-none">
                  <Lock className="w-5 h-5 text-slate-500 group-focus-within:text-cyan-400 transition-colors" />
                </div>
                <input type="password" required value={password} onChange={(e) => setPassword(e.target.value)} placeholder="••••••••" className="w-full pl-12 pr-4 py-4 bg-[#0a0f1c]/80 border border-slate-700/50 rounded-xl text-white placeholder-slate-600 focus:bg-[#0a0f1c] focus:border-cyan-500/50 focus:ring-1 focus:ring-cyan-500/50 outline-none transition-all" />
              </div>
            </div>
            <button type="submit" disabled={isAuthenticating} className="w-full py-4 px-4 bg-gradient-to-r from-cyan-500 to-blue-600 hover:from-cyan-400 hover:to-blue-500 text-white font-bold rounded-xl shadow-[0_0_20px_rgba(6,182,212,0.3)] transition-all duration-300 transform active:scale-[0.98] flex justify-center items-center gap-3 mt-8 border border-white/10">
              {isAuthenticating ? <><Loader2 className="w-5 h-5 animate-spin" /> AUTHENTICATING...</> : <><ArrowRight className="w-5 h-5" /> Login</>}
            </button>
          </form>
        </div>
      </div>
    );
  }

  if (currentView === "deeni") {
    return <DeeniKaamDashboard onBack={() => setCurrentView("hub")} onLogout={handleLogout} />;
  }

  if (currentView === "department") {
    return <DepartmentDashboard onBack={() => setCurrentView("hub")} />;
  }

  if (currentView === "weekly_risala") {
    const handleBack = () => {
       if(accessiblePortals.length === 1) handleLogout();
       else setCurrentView("hub");
    };
    return <WeeklyRisalaDashboard onBack={handleBack} token={risalaToken} officeUser={risalaUser} onLogout={handleLogout} />;
  }

  return (
    <div className="min-h-screen bg-[#0a0f1c] relative overflow-hidden font-sans text-slate-300 flex items-center justify-center p-4 sm:p-8">
       <div className="absolute inset-0 z-0">
         <div className="absolute top-[10%] left-[10%] w-[30rem] h-[30rem] bg-cyan-600/10 rounded-full blur-[120px] mix-blend-screen"></div>
         <div className="absolute bottom-[10%] right-[10%] w-[30rem] h-[30rem] bg-fuchsia-600/10 rounded-full blur-[120px] mix-blend-screen"></div>
       </div>
       <div className="relative z-10 w-full max-w-7xl">
         <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center mb-12 gap-4">
           <div>
             <h1 className="text-3xl font-bold text-white tracking-wide">HMO <span className="text-cyan-400 font-light">DASHBOARD HUB</span></h1>
             <p className="text-slate-400 text-sm mt-1 uppercase tracking-widest">Select a portal to continue</p>
           </div>
           <button onClick={handleLogout} className="flex items-center gap-2 text-xs font-bold tracking-wider uppercase bg-red-500/10 hover:bg-red-500/20 text-red-400 px-4 py-2.5 rounded-xl border border-red-500/20 transition-all active:scale-95">
             <LogOut className="w-4 h-4" /> Logout
           </button>
         </div>
         <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
           {accessiblePortals.includes("deeni") && (
             <button onClick={() => setCurrentView("deeni")} className="group text-left bg-[#121929]/80 backdrop-blur-md p-8 rounded-3xl border border-cyan-500/20 shadow-[0_0_20px_rgba(6,182,212,0.1)] hover:shadow-[0_0_40px_rgba(6,182,212,0.3)] hover:-translate-y-2 transition-all duration-300 relative overflow-hidden flex flex-col h-full">
                <div className="absolute top-0 right-0 w-32 h-32 bg-cyan-500/10 rounded-full blur-2xl group-hover:bg-cyan-500/20 transition-all"></div>
                <div className="w-16 h-16 bg-cyan-500/10 border border-cyan-500/30 rounded-2xl flex items-center justify-center mb-6 group-hover:scale-110 transition-transform"><BookOpen className="w-8 h-8 text-cyan-400" /></div>
                <h2 className="text-2xl font-bold text-white mb-2 tracking-wide">12 Deeni Kaam</h2>
                <p className="text-slate-400 text-sm leading-relaxed mb-6 flex-grow">Access the live analytics dashboard for 12 Deeni Kaam metrics.</p>
                <div className="flex items-center gap-2 text-cyan-400 text-sm font-bold tracking-widest uppercase mt-auto">Enter Portal <ArrowRight className="w-4 h-4 group-hover:translate-x-2 transition-transform" /></div>
             </button>
           )}
           {accessiblePortals.includes("weekly_risala") && (
             <button onClick={() => setCurrentView("weekly_risala")} className="group text-left bg-[#121929]/80 backdrop-blur-md p-8 rounded-3xl border border-emerald-500/20 shadow-[0_0_20px_rgba(16,185,129,0.1)] hover:shadow-[0_0_40px_rgba(16,185,129,0.3)] hover:-translate-y-2 transition-all duration-300 relative overflow-hidden flex flex-col h-full">
                <div className="absolute top-0 right-0 w-32 h-32 bg-emerald-500/10 rounded-full blur-2xl group-hover:bg-emerald-500/20 transition-all"></div>
                <div className="w-16 h-16 bg-emerald-500/10 border border-emerald-500/30 rounded-2xl flex items-center justify-center mb-6 group-hover:scale-110 transition-transform"><FileText className="w-8 h-8 text-emerald-400" /></div>
                <h2 className="text-xl font-bold text-white mb-2 tracking-wide">Weekly Risala Live Report</h2>
                <p className="text-slate-400 text-sm leading-relaxed mb-6 flex-grow">Monitor weekly risala distribution and performance.</p>
                <div className="flex items-center gap-2 text-emerald-400 text-sm font-bold tracking-widest uppercase mt-auto">Enter Portal <ArrowRight className="w-4 h-4 group-hover:translate-x-2 transition-transform" /></div>
             </button>
           )}
           {accessiblePortals.includes("department") && (
             <button onClick={() => setCurrentView("department")} className="group text-left bg-[#121929]/80 backdrop-blur-md p-8 rounded-3xl border border-fuchsia-500/20 shadow-[0_0_20px_rgba(217,70,239,0.1)] hover:shadow-[0_0_40px_rgba(217,70,239,0.3)] hover:-translate-y-2 transition-all duration-300 relative overflow-hidden flex flex-col h-full">
                <div className="absolute top-0 right-0 w-32 h-32 bg-fuchsia-500/10 rounded-full blur-2xl group-hover:bg-fuchsia-500/20 transition-all"></div>
                <div className="w-16 h-16 bg-fuchsia-500/10 border border-fuchsia-500/30 rounded-2xl flex items-center justify-center mb-6 group-hover:scale-110 transition-transform"><Building2 className="w-8 h-8 text-fuchsia-400" /></div>
                <h2 className="text-2xl font-bold text-white mb-2 tracking-wide">Department Report</h2>
                <p className="text-slate-400 text-sm leading-relaxed mb-6 flex-grow">Under Construction. Future module for detailed departmental analytics.</p>
                <div className="flex items-center gap-2 text-fuchsia-400 text-sm font-bold tracking-widest uppercase mt-auto">Coming Soon <ArrowRight className="w-4 h-4 group-hover:translate-x-2 transition-transform" /></div>
             </button>
           )}
         </div>
       </div>
    </div>
  );
}