"use client";
import React, { useState, useEffect, useMemo } from "react";
import { ArrowLeft, BookOpenCheck, LogOut, RefreshCw, Filter, Search, Activity, Download, Lock, User, LayoutDashboard, MapPin, Layers } from "lucide-react";

const SCRIPT_URL = "https://script.google.com/macros/s/AKfycbwtCMnX7JgApy0BLFdhs7ByumM8H9JGjLLgDbYMBMpuQjtPHuzywoDesSz1lYZ-hYwE/exec";

export default function WeeklyRisalaDashboard({ onBack }) {
  const [isLoggedIn, setIsLoggedIn] = useState(false);
  const [userId, setUserId] = useState("");
  const [password, setPassword] = useState("");
  const [authLoading, setAuthLoading] = useState(false);
  const [authError, setAuthError] = useState("");
  
  const [token, setToken] = useState("");
  const [officeUser, setOfficeUser] = useState({});
  const [dashboardData, setDashboardData] = useState({ b4Value: "", b6Value: "", totalRows: 0 });
  const [rows, setRows] = useState([]);
  const [loading, setLoading] = useState(false);
  const [fetchError, setFetchError] = useState("");

  const [search, setSearch] = useState("");
  const [region, setRegion] = useState("");
  const [state, setState] = useState("");
  const [division, setDivision] = useState("");
  const [district, setDistrict] = useState("");
  const [department, setDepartment] = useState("");
  const [chain, setChain] = useState("");

  const [currentPage, setCurrentPage] = useState(1);
  const rowsPerPage = 10;

  useEffect(() => {
    const storedToken = sessionStorage.getItem("risalaToken");
    const storedUser = sessionStorage.getItem("risalaUser");
    if (storedToken && storedUser) {
        setToken(storedToken);
        setOfficeUser(JSON.parse(storedUser));
        setIsLoggedIn(true);
        fetchDashboard(storedToken);
    }
  }, []);

  const handleLogin = async (e) => {
    e.preventDefault();
    setAuthLoading(true);
    setAuthError("");
    try {
        const res = await fetch(`${SCRIPT_URL}?action=login&userId=${encodeURIComponent(userId)}&password=${encodeURIComponent(password)}`);
        const d = await res.json();
        if (d.status === "Success") {
            sessionStorage.setItem("risalaToken", d.token);
            sessionStorage.setItem("risalaUser", JSON.stringify(d.user));
            setToken(d.token);
            setOfficeUser(d.user);
            setIsLoggedIn(true);
            fetchDashboard(d.token);
        } else {
            setAuthError(d.message || "Invalid credentials.");
        }
    } catch (err) {
        setAuthError("Network Error. Cannot connect to server.");
    }
    setAuthLoading(false);
  };

  const handleLogout = () => {
    sessionStorage.removeItem("risalaToken");
    sessionStorage.removeItem("risalaUser");
    setIsLoggedIn(false);
    setToken("");
    setOfficeUser({});
    setRows([]);
  };

  const fetchDashboard = async (tkn) => {
    setLoading(true);
    setFetchError("");
    try {
        const res = await fetch(`${SCRIPT_URL}?action=dashboard&token=${encodeURIComponent(tkn)}&_=${Date.now()}`, { cache: "no-store" });
        const d = await res.json();
        if (d.status === "Success") {
            setDashboardData({ b4Value: d.b4Value, b6Value: d.b6Value, totalRows: d.totalRows });
            setRows(Array.isArray(d.rows) ? d.rows : []);
        } else {
            if (/session|login required|invalid/i.test(d.message || "")) {
                handleLogout();
                alert("Session expired. Please login again.");
            } else {
                setFetchError(d.message || "Failed to load data.");
            }
        }
    } catch(err) {
        setFetchError("Network Error. Cannot fetch dashboard data.");
    }
    setLoading(false);
  };

  const sameClient = (a, b) => String(a || "").trim().toLowerCase() === String(b || "").trim().toLowerCase();

  const filteredRows = useMemo(() => {
    let result = rows.filter(x =>
        (!region || sameClient(x.region, region)) &&
        (!state || sameClient(x.state, state)) &&
        (!division || sameClient(x.division, division)) &&
        (!district || sameClient(x.district, district)) &&
        (!department || sameClient(x.department, department)) &&
        (!chain || sameClient(x.chain, chain)) &&
        (!search || JSON.stringify(x).toLowerCase().includes(search.toLowerCase().trim()))
    );
    return result;
  }, [rows, search, region, state, division, district, department, chain]);

  useEffect(() => { setCurrentPage(1); }, [search, region, state, division, district, department, chain]);

  const pagedRows = useMemo(() => {
    const start = (currentPage - 1) * rowsPerPage;
    return filteredRows.slice(start, start + rowsPerPage);
  }, [filteredRows, currentPage]);

  const groupCount = (data, key) => {
    const map = {};
    data.forEach(x => {
        const k = (x[key] || "Unknown").trim();
        const n = Number(String(x.report || "").replace(/,/g, ""));
        map[k] = (map[k] || 0) + (isNaN(n) ? 0 : n);
    });
    return Object.keys(map).sort((a,b) => map[b] - map[a]).map(k => ({ label: k, count: map[k] }));
  };

  const totalReportSum = useMemo(() => {
    return filteredRows.reduce((sum, x) => {
        const n = Number(String(x.report || "").replace(/,/g, ""));
        return sum + (isNaN(n) ? 0 : n);
    }, 0);
  }, [filteredRows]);

  const uniqValues = (data, key) => {
    return [...new Set(data.map(x => String(x[key] || "").trim()).filter(Boolean))].sort((a,b) => a.localeCompare(b));
  };

  const downloadCSV = () => {
    if (!filteredRows.length) return alert("No data to export");
    const headers = ["Date/Time", "Name", "Contact Number", "Chain Type", "Nigran Level", "Zimmedar Level", "Department", "Risala Report", "Pincode", "District", "Division", "State", "Region", "Country"];
    const csvRows = [headers.join(",")];
    filteredRows.forEach(x => {
        const row = [x.date, x.name, x.contact, x.chain, x.nigran||x.zimmedar||x.level, x.zimmedar, x.department, x.report, x.pincode, x.district, x.division, x.state, x.region, x.country];
        csvRows.push(row.map(v => `"${(v||"").toString().replace(/"/g, '""')}"`).join(","));
    });
    const blob = new Blob([csvRows.join("\n")], { type: "text/csv" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `Weekly_Risala_Report.csv`;
    a.click();
  };

  const MiniTable = ({ title, data }) => (
    <div className="bg-[#121929]/60 backdrop-blur-md rounded-2xl border border-emerald-500/20 overflow-hidden flex flex-col h-[320px] shadow-[0_0_20px_rgba(16,185,129,0.05)]">
      <div className="p-4 border-b border-emerald-500/20 bg-emerald-900/10 sticky top-0">
        <h3 className="text-xs font-bold text-emerald-400 uppercase tracking-widest">{title}</h3>
      </div>
      <div className="overflow-y-auto flex-1 p-2 custom-scrollbar">
        <table className="w-full text-left text-sm">
          <tbody>
            {data.length > 0 ? data.map((d, i) => (
              <tr key={i} className="border-b border-white/5 hover:bg-white/[0.02] transition-colors">
                <td className="py-3 px-3 text-slate-300 font-medium text-xs">{d.label}</td>
                <td className="py-3 px-3 text-emerald-400 font-bold text-right text-sm">{d.count.toLocaleString("en-IN")}</td>
              </tr>
            )) : <tr><td colSpan="2" className="text-center py-4 text-slate-500 text-xs">No data available</td></tr>}
          </tbody>
        </table>
      </div>
    </div>
  );

  if (!isLoggedIn) {
    return (
      <div className="min-h-screen flex items-center justify-center bg-[#0a0f1c] relative overflow-hidden font-sans">
        <div className="absolute inset-0 z-0">
          <div className="absolute top-[20%] left-[20%] w-[30rem] h-[30rem] bg-emerald-600/20 rounded-full blur-[120px] mix-blend-screen"></div>
        </div>
        <div className="relative z-10 w-full max-w-lg p-8 sm:p-12 bg-[#121929]/70 backdrop-blur-2xl border border-white/5 rounded-3xl shadow-[0_0_60px_rgba(16,185,129,0.1)]">
          <button onClick={onBack} className="absolute top-6 left-6 text-slate-400 hover:text-emerald-400 transition-colors"><ArrowLeft className="w-5 h-5" /></button>
          <div className="flex flex-col items-center justify-center mb-10 text-center mt-4">
            <div className="w-16 h-16 bg-gradient-to-br from-emerald-400 to-teal-600 rounded-2xl flex items-center justify-center mb-6 shadow-[0_0_30px_rgba(16,185,129,0.4)] border border-white/20">
              <BookOpenCheck className="w-8 h-8 text-white" />
            </div>
            <h1 className="text-2xl font-bold text-white tracking-wide">WEEKLY <span className="text-emerald-400 font-light">RISALA</span></h1>
            <p className="text-slate-400 text-xs mt-2 font-medium tracking-widest uppercase">Office Authentication</p>
          </div>
          {authError && <div className="mb-6 p-4 bg-red-900/40 border border-red-500/50 text-red-400 text-sm rounded-xl text-center backdrop-blur-sm">{authError}</div>}
          <form onSubmit={handleLogin} className="space-y-6">
            <div className="space-y-2">
              <label className="text-xs font-bold text-emerald-400 uppercase tracking-widest ml-1">User ID</label>
              <div className="relative group">
                <div className="absolute inset-y-0 left-0 pl-4 flex items-center pointer-events-none"><User className="w-5 h-5 text-slate-500 group-focus-within:text-emerald-400 transition-colors" /></div>
                <input type="text" required value={userId} onChange={(e) => setUserId(e.target.value)} placeholder="Enter assigned user ID" className="w-full pl-12 pr-4 py-4 bg-[#0a0f1c]/80 border border-slate-700/50 rounded-xl text-white placeholder-slate-600 focus:bg-[#0a0f1c] focus:border-emerald-500/50 focus:ring-1 focus:ring-emerald-500/50 outline-none transition-all" />
              </div>
            </div>
            <div className="space-y-2">
              <label className="text-xs font-bold text-emerald-400 uppercase tracking-widest ml-1">Password</label>
              <div className="relative group">
                <div className="absolute inset-y-0 left-0 pl-4 flex items-center pointer-events-none"><Lock className="w-5 h-5 text-slate-500 group-focus-within:text-emerald-400 transition-colors" /></div>
                <input type="password" required value={password} onChange={(e) => setPassword(e.target.value)} placeholder="••••••••" className="w-full pl-12 pr-4 py-4 bg-[#0a0f1c]/80 border border-slate-700/50 rounded-xl text-white placeholder-slate-600 focus:bg-[#0a0f1c] focus:border-emerald-500/50 focus:ring-1 focus:ring-emerald-500/50 outline-none transition-all" />
              </div>
            </div>
            <button type="submit" disabled={authLoading} className="w-full py-4 px-4 bg-gradient-to-r from-emerald-500 to-teal-600 hover:from-emerald-400 hover:to-teal-500 text-white font-bold rounded-xl shadow-[0_0_20px_rgba(16,185,129,0.3)] transition-all duration-300 transform active:scale-[0.98] flex justify-center items-center gap-3 mt-8 border border-white/10">
              {authLoading ? "VERIFYING..." : "ACCESS DASHBOARD"}
            </button>
          </form>
        </div>
      </div>
    );
  }

  const regionData = groupCount(filteredRows, "region");
  const maxRegionCount = regionData.length ? Math.max(...regionData.map(d => d.count)) : 0;

  return (
    <div className="min-h-screen bg-[#0a0f1c] text-slate-300 font-sans relative overflow-hidden">
      <div className="fixed inset-0 z-0 pointer-events-none">
        <div className="absolute top-0 right-0 w-[40rem] h-[40rem] bg-emerald-900/10 rounded-full blur-[150px]"></div>
      </div>
      <div className="relative z-10">
        <header className="bg-[#121929]/80 backdrop-blur-xl border-b border-white/5 sticky top-0 z-30">
          <div className="max-w-[1400px] mx-auto px-4 sm:px-6 lg:px-8 flex flex-col md:flex-row items-center justify-between py-4 gap-4">
            <div className="flex items-center gap-4 w-full md:w-auto">
              <button onClick={onBack} className="p-2.5 bg-[#1a2333] hover:bg-[#222d42] border border-white/10 text-white rounded-xl shadow-lg transition-colors"><ArrowLeft className="w-5 h-5" /></button>
              <div className="p-2.5 bg-gradient-to-br from-emerald-500 to-teal-600 text-white rounded-xl shadow-[0_0_15px_rgba(16,185,129,0.4)] border border-white/10 hidden sm:block">
                <BookOpenCheck className="w-6 h-6" />
              </div>
              <div>
                <h1 className="text-lg sm:text-xl font-bold text-white tracking-wide">WEEKLY RISALA <span className="font-light text-emerald-400">REPORT</span></h1>
                <p className="text-[10px] font-bold text-emerald-500 uppercase tracking-widest mt-0.5">{dashboardData.b4Value || "Risala Dashboard"}</p>
              </div>
            </div>
            <div className="flex flex-wrap items-center justify-center md:justify-end gap-3 w-full md:w-auto">
              <div className="text-[10px] font-bold text-slate-400 uppercase tracking-widest bg-slate-800/50 px-3 py-2 rounded-lg border border-slate-700 hidden lg:block">
                User: <span className="text-emerald-400">{officeUser.name || officeUser.userId}</span>
              </div>
              <button onClick={() => fetchDashboard(token)} disabled={loading} className="flex items-center gap-2 text-xs font-bold tracking-wider uppercase bg-[#1a2333] hover:bg-[#222d42] text-emerald-400 px-4 py-2.5 rounded-xl border border-emerald-500/20 transition-all active:scale-95">
                <RefreshCw className={`w-4 h-4 ${loading ? "animate-spin" : ""}`} /> <span className="hidden sm:inline">Sync</span>
              </button>
              <button onClick={downloadCSV} className="flex items-center gap-2 text-xs font-bold tracking-wider uppercase bg-teal-900/30 hover:bg-teal-900/50 text-teal-400 px-4 py-2.5 rounded-xl border border-teal-500/30 transition-all active:scale-95">
                <Download className="w-4 h-4" /> <span className="hidden sm:inline">Excel</span>
              </button>
              <button onClick={handleLogout} className="flex items-center gap-2 text-xs font-bold tracking-wider uppercase bg-red-500/10 hover:bg-red-500/20 text-red-400 px-4 py-2.5 rounded-xl border border-red-500/20 transition-all active:scale-95">
                <LogOut className="w-4 h-4" /> <span className="hidden sm:inline">Logout</span>
              </button>
            </div>
          </div>
        </header>

        <main className="max-w-[1400px] mx-auto px-4 sm:px-6 lg:px-8 py-8 space-y-6">
          
          {fetchError && <div className="p-4 bg-red-900/20 border border-red-500/30 text-red-400 rounded-2xl text-sm font-medium flex items-center gap-3"><div className="w-2 h-2 rounded-full bg-red-500 animate-ping"></div>{fetchError}</div>}

          {/* KPIs */}
          <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
            <div className="bg-[#121929]/80 backdrop-blur-md p-6 rounded-3xl border border-emerald-500/20 shadow-[0_0_20px_rgba(16,185,129,0.1)] relative overflow-hidden group">
              <div className="relative flex justify-between items-center z-10">
                <div className="flex items-center gap-4">
                    <div className="p-3.5 rounded-2xl bg-slate-800 text-slate-300 border border-white/5"><LayoutDashboard className="w-6 h-6" /></div>
                    <div>
                        <p className="text-[11px] font-bold text-slate-500 uppercase tracking-widest mb-1">TOTAL SUBMITTED</p>
                        <h3 className="text-4xl font-extrabold text-white tracking-tight">{filteredRows.length.toLocaleString("en-IN")}</h3>
                    </div>
                </div>
              </div>
            </div>
            <div className="bg-[#121929]/80 backdrop-blur-md p-6 rounded-3xl border border-emerald-500/20 shadow-[0_0_20px_rgba(16,185,129,0.1)] relative overflow-hidden group">
              <div className="relative flex justify-between items-center z-10">
                <div className="flex items-center gap-4">
                    <div className="p-3.5 rounded-2xl bg-emerald-900/40 text-emerald-400 border border-emerald-500/20"><Activity className="w-6 h-6" /></div>
                    <div>
                        <p className="text-[11px] font-bold text-slate-500 uppercase tracking-widest mb-1">REPORT QUANTITY</p>
                        <h3 className="text-4xl font-extrabold text-emerald-400 tracking-tight" style={{ textShadow: "0 0 15px rgba(16,185,129,0.4)" }}>{totalReportSum.toLocaleString("en-IN")}</h3>
                    </div>
                </div>
              </div>
            </div>
          </div>

          {/* Filters */}
          <div className="bg-[#121929]/60 backdrop-blur-md p-6 rounded-3xl border border-white/5 shadow-lg">
            <div className="flex items-center gap-3 mb-6">
              <Filter className="w-5 h-5 text-emerald-400" />
              <h2 className="text-sm font-bold text-white tracking-widest uppercase">Data Parameters</h2>
            </div>
            <div className="grid grid-cols-2 md:grid-cols-7 gap-4">
              <div className="col-span-2 md:col-span-1">
                <input type="text" placeholder="Search..." value={search} onChange={(e) => setSearch(e.target.value)} className="w-full p-3 rounded-xl border border-slate-700/50 bg-[#0a0f1c] text-sm font-medium text-slate-300 focus:outline-none focus:border-emerald-500/50 focus:ring-1 focus:ring-emerald-500/50" />
              </div>
              {[
                { label: "Region", val: region, set: setRegion, opts: uniqValues(rows, "region") },
                { label: "State", val: state, set: setState, opts: uniqValues(rows.filter(x=>!region||sameClient(x.region,region)), "state") },
                { label: "Division", val: division, set: setDivision, opts: uniqValues(rows.filter(x=>(!region||sameClient(x.region,region))&&(!state||sameClient(x.state,state))), "division") },
                { label: "District", val: district, set: setDistrict, opts: uniqValues(rows.filter(x=>(!region||sameClient(x.region,region))&&(!state||sameClient(x.state,state))&&(!division||sameClient(x.division,division))), "district") },
                { label: "Department", val: department, set: setDepartment, opts: uniqValues(rows, "department") },
                { label: "Chain", val: chain, set: setChain, opts: uniqValues(rows, "chain") }
              ].map((f, i) => (
                <div key={i}>
                  <select value={f.val} onChange={(e) => f.set(e.target.value)} className="w-full p-3 rounded-xl border border-slate-700/50 bg-[#0a0f1c] text-sm font-medium text-slate-400 focus:outline-none focus:border-emerald-500/50 focus:ring-1 focus:ring-emerald-500/50 appearance-none">
                    <option value="">{f.label}</option>
                    {f.opts.map(o => <option key={o} value={o}>{o}</option>)}
                  </select>
                </div>
              ))}
            </div>
          </div>

          {/* Charts Grid */}
          <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
            
            {/* Region Bar Chart */}
            <div className="bg-[#121929]/60 backdrop-blur-md rounded-2xl border border-emerald-500/20 p-5 shadow-[0_0_20px_rgba(16,185,129,0.05)] h-[320px] flex flex-col">
               <h3 className="text-xs font-bold text-emerald-400 uppercase tracking-widest mb-4">REPORTS BY REGION</h3>
               <div className="flex-1 flex items-end gap-2 pb-2">
                 {regionData.length > 0 ? regionData.slice(0, 8).map((d, i) => {
                    const h = maxRegionCount ? (d.count / maxRegionCount) * 100 : 0;
                    return (
                      <div key={i} className="flex-1 flex flex-col justify-end items-center group relative h-full">
                         <div className="opacity-0 group-hover:opacity-100 absolute -top-8 bg-[#0a0f1c] border border-emerald-500/30 text-emerald-400 text-[10px] px-2 py-1 rounded shadow-lg transition-opacity whitespace-nowrap z-10 pointer-events-none">
                           {d.count.toLocaleString("en-IN")}
                         </div>
                         <div style={{height: `${Math.max(h, 2)}%`}} className="w-full bg-emerald-600/70 group-hover:bg-emerald-400 rounded-t-sm transition-all border-t border-emerald-400/50" />
                         <span className="text-[9px] text-slate-400 mt-2 truncate w-full text-center px-1" title={d.label}>{d.label}</span>
                      </div>
                    )
                 }) : <div className="w-full text-center text-slate-500 text-xs my-auto">No data</div>}
               </div>
            </div>

            <MiniTable title="REPORTS BY STATE" data={groupCount(filteredRows, "state")} />
            <MiniTable title="REPORTS BY DEPARTMENT" data={groupCount(filteredRows, "department")} />
            <MiniTable title="REPORTS BY DIVISION" data={groupCount(filteredRows, "division")} />
          </div>

          {/* Main Table */}
          <div className="bg-[#121929]/60 backdrop-blur-md rounded-3xl border border-white/5 overflow-hidden shadow-lg mt-6">
            <div className="p-6 border-b border-white/5 flex flex-col md:flex-row justify-between items-center gap-4 bg-black/20">
              <div>
                <h2 className="text-sm font-bold text-white uppercase tracking-widest">Detailed Telemetry Output</h2>
                <p className="text-[11px] text-slate-500 mt-1 uppercase tracking-widest">Source: {dashboardData.totalRows} • Visible: {filteredRows.length}</p>
              </div>
            </div>
            
            <div className="overflow-x-auto">
              <table className="min-w-full text-left text-sm whitespace-nowrap">
                <thead className="bg-[#0a0f1c]/50">
                  <tr>
                    {["Date", "Name", "Contact", "Chain", "Level", "Department", "Report", "District", "Division", "State", "Region"].map(h => (
                      <th key={h} className={`px-4 py-4 text-[10px] font-bold text-slate-400 uppercase tracking-widest ${h==="Report"?"text-emerald-400 text-right":""}`}>{h}</th>
                    ))}
                  </tr>
                </thead>
                <tbody className="divide-y divide-white/5">
                  {pagedRows.length > 0 ? pagedRows.map((row, idx) => (
                    <tr key={idx} className="hover:bg-white/[0.02] transition-colors">
                      <td className="px-4 py-3 text-xs text-slate-500">{row.date}</td>
                      <td className="px-4 py-3 text-xs font-bold text-white">{row.name}</td>
                      <td className="px-4 py-3 text-xs text-slate-400">{row.contact}</td>
                      <td className="px-4 py-3 text-[10px] uppercase tracking-wider font-bold text-slate-300"><span className="bg-slate-800 px-2 py-1 rounded">{row.chain}</span></td>
                      <td className="px-4 py-3 text-xs text-slate-400">{row.nigran || row.zimmedar || row.level}</td>
                      <td className="px-4 py-3 text-xs text-slate-300">{row.department}</td>
                      <td className="px-4 py-3 text-sm text-right font-extrabold text-emerald-400">{row.report}</td>
                      <td className="px-4 py-3 text-xs text-slate-300">{row.district}</td>
                      <td className="px-4 py-3 text-xs text-slate-400">{row.division}</td>
                      <td className="px-4 py-3 text-xs text-slate-400">{row.state}</td>
                      <td className="px-4 py-3 text-xs text-slate-500">{row.region}</td>
                    </tr>
                  )) : (
                    <tr><td colSpan="11" className="px-6 py-12 text-center text-slate-500 text-xs uppercase tracking-widest">{loading ? "Fetching records..." : "No matching records found"}</td></tr>
                  )}
                </tbody>
              </table>
            </div>

            {/* Pagination */}
            {filteredRows.length > rowsPerPage && (
              <div className="p-4 border-t border-white/5 bg-black/20 flex justify-center items-center gap-2">
                <button disabled={currentPage===1} onClick={()=>setCurrentPage(p=>p-1)} className="px-3 py-1.5 rounded-lg border border-white/10 bg-slate-800 text-xs font-bold text-slate-300 disabled:opacity-30">PREV</button>
                <span className="text-xs font-bold text-emerald-400 px-4">PAGE {currentPage} OF {Math.ceil(filteredRows.length / rowsPerPage)}</span>
                <button disabled={currentPage===Math.ceil(filteredRows.length / rowsPerPage)} onClick={()=>setCurrentPage(p=>p+1)} className="px-3 py-1.5 rounded-lg border border-white/10 bg-slate-800 text-xs font-bold text-slate-300 disabled:opacity-30">NEXT</button>
              </div>
            )}
          </div>

        </main>
      </div>
    </div>
  );
}