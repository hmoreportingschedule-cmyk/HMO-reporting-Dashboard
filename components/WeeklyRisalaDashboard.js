"use client";
import React, { useState, useEffect, useMemo } from "react";
import Papa from "papaparse";
import { ArrowLeft, BookOpenCheck, LogOut, RefreshCw, Filter, Search, Activity, Download, LayoutDashboard, Database, Presentation } from "lucide-react";
import pptxgen from "pptxgenjs";
import * as XLSX from "xlsx";

export default function WeeklyRisalaDashboard({ onBack, officeUser, onLogout }) {
  const [sheetUrl, setSheetUrl] = useState("");
  const [dashboardData, setDashboardData] = useState({ b4Value: "Live CSV Sync", b6Value: "Online", totalRows: 0 });
  const [rows, setRows] = useState([]);
  const [loading, setLoading] = useState(false);
  const [fetchError, setFetchError] = useState("");

  const [search, setSearch] = useState("");
  const [region, setRegion] = useState(officeUser?.region && officeUser.region.toLowerCase() !== "all" ? officeUser.region : "");
  const [state, setState] = useState(officeUser?.state && officeUser.state.toLowerCase() !== "all" ? officeUser.state : "");
  const [division, setDivision] = useState(officeUser?.division && officeUser.division.toLowerCase() !== "all" ? officeUser.division : "");
  const [district, setDistrict] = useState(officeUser?.district && officeUser.district.toLowerCase() !== "all" ? officeUser.district : "");
  const [department, setDepartment] = useState(officeUser?.department && officeUser.department.toLowerCase() !== "all" ? officeUser.department : "");
  const [chain, setChain] = useState(officeUser?.chain && officeUser.chain.toLowerCase() !== "all" ? officeUser.chain : "");

  const [currentPage, setCurrentPage] = useState(1);
  const rowsPerPage = 10;

  useEffect(() => {
    const savedUrl = localStorage.getItem("risala_csv_url");
    if (savedUrl) {
      setSheetUrl(savedUrl);
      fetchData(savedUrl);
    }
  }, []);

  const fetchData = async (urlToFetch) => {
    const targetUrl = urlToFetch || sheetUrl;
    if (!targetUrl) {
      setFetchError("Please provide a valid Google Sheet CSV URL.");
      return;
    }
    
    setLoading(true);
    setFetchError("");
    localStorage.setItem("risala_csv_url", targetUrl);

    Papa.parse(targetUrl, {
      download: true,
      header: true,
      skipEmptyLines: true,
      complete: (results) => {
        if (results.data && results.data.length > 0) {
          const normalizeKey = (key) => key.toLowerCase().replace(/[^a-z]/g, '');
          const mappedRows = results.data.map(row => {
            let newRow = { original: row };
            for (let key in row) {
              let nKey = normalizeKey(key);
              if (nKey === 'date' || nKey.includes('timestamp')) newRow.date = row[key];
              else if (nKey === 'name' || nKey === 'username') newRow.name = row[key];
              else if (nKey.includes('contact') || nKey.includes('mobile')) newRow.contact = row[key];
              else if (nKey.includes('chain')) newRow.chain = row[key];
              else if (nKey.includes('level') || nKey.includes('nigran') || nKey.includes('zimmedar')) newRow.level = row[key];
              else if (nKey.includes('department')) newRow.department = row[key];
              else if (nKey.includes('report') || nKey.includes('qty')) newRow.report = row[key];
              else if (nKey === 'district') newRow.district = row[key];
              else if (nKey === 'division') newRow.division = row[key];
              else if (nKey === 'state') newRow.state = row[key];
              else if (nKey === 'region') newRow.region = row[key];
              else if (nKey.includes('pincode') || nKey.includes('pin')) newRow.pincode = row[key];
              else if (nKey === 'country') newRow.country = row[key];
            }
            newRow.nigran = newRow.level;
            newRow.zimmedar = newRow.level;
            return newRow;
          }).filter(r => Object.keys(r).length > 1); // remove empty rows
          
          setRows(mappedRows);
          setDashboardData({ b4Value: "Live CSV Sync", b6Value: "Online", totalRows: mappedRows.length });
        } else {
          setFetchError("Data stream empty or invalid CSV link.");
        }
        setLoading(false);
      },
      error: (err) => {
        setFetchError("Connection Failed. Make sure link is 'Publish to Web' as CSV.");
        setLoading(false);
      }
    });
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

  const downloadPPT = () => {
    if (!filteredRows.length) return alert("No data to export");
    let pres = new pptxgen();
    let slide1 = pres.addSlide();
    slide1.background = { color: "0a0f1c" };
    slide1.addText("WEEKLY RISALA REPORT", { x: 1, y: 2, w: 8, fontSize: 36, bold: true, color: "34d399", align: "center" });
    slide1.addText(`User: ${officeUser?.name || officeUser?.userId || "Admin"}`, { x: 1, y: 3, w: 8, fontSize: 16, color: "94a3b8", align: "center" });
    slide1.addText(`Generated on: ${new Date().toLocaleDateString()}`, { x: 1, y: 3.5, w: 8, fontSize: 12, color: "64748b", align: "center" });

    let slide2 = pres.addSlide();
    slide2.background = { color: "0a0f1c" };
    slide2.addText("Key Performance Indicators", { x: 0.5, y: 0.5, w: 9, fontSize: 24, bold: true, color: "ffffff" });
    slide2.addShape(pres.ShapeType.rect, { x: 1, y: 1.5, w: 3.5, h: 2, fill: "121929", line: {color: "34d399", width: 1} });
    slide2.addText("TOTAL SUBMITTED", { x: 1, y: 1.8, w: 3.5, fontSize: 14, color: "94a3b8", align: "center" });
    slide2.addText(filteredRows.length.toLocaleString("en-IN"), { x: 1, y: 2.3, w: 3.5, fontSize: 32, bold: true, color: "ffffff", align: "center" });
    slide2.addShape(pres.ShapeType.rect, { x: 5.5, y: 1.5, w: 3.5, h: 2, fill: "121929", line: {color: "34d399", width: 1} });
    slide2.addText("REPORT QUANTITY", { x: 5.5, y: 1.8, w: 3.5, fontSize: 14, color: "94a3b8", align: "center" });
    slide2.addText(totalReportSum.toLocaleString("en-IN"), { x: 5.5, y: 2.3, w: 3.5, fontSize: 32, bold: true, color: "34d399", align: "center" });

    const regionDataPPT = groupCount(filteredRows, "region");
    if (regionDataPPT.length > 0) {
        let slide3 = pres.addSlide();
        slide3.background = { color: "0a0f1c" };
        slide3.addText("Reports by Region", { x: 0.5, y: 0.5, w: 9, fontSize: 24, bold: true, color: "ffffff" });
        let chartData = [{ name: "Reports", labels: regionDataPPT.slice(0, 8).map(d => d.label || "Unknown"), values: regionDataPPT.slice(0, 8).map(d => d.count) }];
        slide3.addChart(pres.ChartType.bar, chartData, { x: 0.5, y: 1.2, w: 9, h: 3.5, barDir: "col", chartColors: ["34d399"], valAxisLabelColor: "94a3b8", catAxisLabelColor: "94a3b8", showLegend: false });
    }
    pres.writeFile({ fileName: `Weekly_Risala_PPT_${new Date().getTime()}.pptx` });
  };

  const downloadExcel = () => {
    if (!filteredRows.length) return alert("No data to export");
    const exportData = filteredRows.map(x => ({
        "Date/Time": x.date || "",
        "Name": x.name || "",
        "Contact Number": x.contact || "",
        "Chain Type": x.chain || "",
        "Level": x.nigran || x.zimmedar || x.level || "",
        "Department": x.department || "",
        "Risala Report": x.report || "",
        "Pincode": x.pincode || "",
        "District": x.district || "",
        "Division": x.division || "",
        "State": x.state || "",
        "Region": x.region || "",
        "Country": x.country || ""
    }));
    const worksheet = XLSX.utils.json_to_sheet(exportData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Weekly Risala Data");
    XLSX.writeFile(workbook, `Weekly_Risala_RawData_${new Date().getTime()}.xlsx`);
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
                <p className="text-[10px] font-bold text-emerald-500 uppercase tracking-widest mt-0.5">{dashboardData.b4Value}</p>
              </div>
            </div>
            <div className="flex flex-wrap items-center justify-center md:justify-end gap-3 w-full md:w-auto">
              <div className="text-[10px] font-bold text-slate-400 uppercase tracking-widest bg-slate-800/50 px-3 py-2 rounded-lg border border-slate-700 hidden lg:block">
                User: <span className="text-emerald-400">{officeUser?.name || officeUser?.userId || "Admin"}</span>
              </div>
              <button onClick={() => fetchData(sheetUrl)} disabled={loading} className="flex items-center gap-2 text-xs font-bold tracking-wider uppercase bg-[#1a2333] hover:bg-[#222d42] text-emerald-400 px-4 py-2.5 rounded-xl border border-emerald-500/20 transition-all active:scale-95">
                <RefreshCw className={`w-4 h-4 ${loading ? "animate-spin" : ""}`} /> <span className="hidden sm:inline">Sync</span>
              </button>
              <button onClick={downloadPPT} className="flex items-center gap-2 text-xs font-bold tracking-wider uppercase bg-emerald-900/30 hover:bg-emerald-900/50 text-emerald-400 px-4 py-2.5 rounded-xl border border-emerald-500/30 transition-all active:scale-95">
                <Presentation className="w-4 h-4" /> <span className="hidden sm:inline">PPT</span>
              </button>
              <button onClick={downloadExcel} className="flex items-center gap-2 text-xs font-bold tracking-wider uppercase bg-teal-900/30 hover:bg-teal-900/50 text-teal-400 px-4 py-2.5 rounded-xl border border-teal-500/30 transition-all active:scale-95">
                <Download className="w-4 h-4" /> <span className="hidden sm:inline">Excel</span>
              </button>
              <button onClick={onLogout} className="flex items-center gap-2 text-xs font-bold tracking-wider uppercase bg-red-500/10 hover:bg-red-500/20 text-red-400 px-4 py-2.5 rounded-xl border border-red-500/20 transition-all active:scale-95">
                <LogOut className="w-4 h-4" /> <span className="hidden sm:inline">Logout</span>
              </button>
            </div>
          </div>
        </header>

        <main className="max-w-[1400px] mx-auto px-4 sm:px-6 lg:px-8 py-8 space-y-6">
          
          <div className="bg-[#121929]/60 backdrop-blur-md p-5 rounded-2xl border border-white/5 flex flex-col md:flex-row gap-4 items-center justify-between">
            <div className="flex items-center gap-3 w-full md:w-auto">
              <div className="p-2 bg-slate-800/50 rounded-lg border border-slate-700">
                <Database className="w-5 h-5 text-emerald-400" />
              </div>
              <span className="font-semibold text-slate-300 text-sm tracking-wide">Responses CSV Link</span>
            </div>
            <div className="flex w-full md:w-2/3 gap-3">
              <input type="text" value={sheetUrl} onChange={(e) => setSheetUrl(e.target.value)} placeholder="Paste Weekly Risala 'Publish to Web' CSV Link..." className="w-full px-4 py-2.5 text-sm rounded-xl border border-slate-700/50 bg-[#0a0f1c] text-white focus:outline-none focus:border-emerald-500/50 focus:ring-1 focus:ring-emerald-500/50 transition-all" />
              <button onClick={() => fetchData(sheetUrl)} className="px-6 py-2.5 bg-emerald-600 hover:bg-emerald-500 text-white text-xs tracking-widest font-bold uppercase rounded-xl shrink-0 transition-colors shadow-[0_0_15px_rgba(16,185,129,0.2)]">Connect</button>
            </div>
          </div>

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
                  <label className="block text-[10px] font-bold text-slate-500 uppercase tracking-widest mb-2">{f.label}</label>
                  <select value={f.val} onChange={(e) => f.set(e.target.value)} disabled={officeUser?.[f.label.toLowerCase()] && officeUser[f.label.toLowerCase()].toLowerCase() !== "all"} className="w-full p-3 rounded-xl border border-slate-700/50 bg-[#0a0f1c] text-sm font-medium text-slate-400 focus:outline-none focus:border-emerald-500/50 focus:ring-1 focus:ring-emerald-500/50 appearance-none">
                    <option value="">{f.label}</option>
                    {f.opts.map(o => <option key={o} value={o}>{o}</option>)}
                  </select>
                </div>
              ))}
            </div>
          </div>

          {/* Charts Grid */}
          <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
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
