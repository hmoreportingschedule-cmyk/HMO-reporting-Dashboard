"use client";
import React, { useState, useEffect, useMemo } from "react";
import Papa from "papaparse";
import { LogOut, RefreshCw, Filter, Database, CheckCircle2, Layers, Calendar, MapPin, Search, Activity, Globe, ArrowLeft, User, Presentation, Download } from "lucide-react";
import * as XLSX from "xlsx";
import pptxgen from "pptxgenjs";

const DEFAULT_SHEET_URL = "";

export default function DeeniKaamDashboard({ onBack, onLogout, officeUser }) {
  const [sheetUrl, setSheetUrl] = useState(DEFAULT_SHEET_URL);
  const [rawData, setRawData] = useState([]);
  const [loading, setLoading] = useState(false);
  const [fetchError, setFetchError] = useState("");

  const [selectedYear, setSelectedYear] = useState("All");
  const [selectedMonth, setSelectedMonth] = useState("All");
  const [selectedRegion, setSelectedRegion] = useState(officeUser?.region && officeUser.region.toLowerCase() !== "all" ? officeUser.region : "All");
  const [selectedDistrict, setSelectedDistrict] = useState(officeUser?.district && officeUser.district.toLowerCase() !== "all" ? officeUser.district : "All");
  const [selectedCategory, setSelectedCategory] = useState(officeUser?.department && officeUser.department.toLowerCase() !== "all" ? officeUser.department : "All");
  const [searchTerm, setSearchTerm] = useState("");

  const fetchData = async () => {
    if (!sheetUrl) {
      setFetchError("System Error: No valid data source provided.");
      return;
    }
    setLoading(true);
    setFetchError("");

    try {
      Papa.parse(sheetUrl, {
        download: true,
        header: true,
        skipEmptyLines: true,
        complete: (results) => {
          if (results.data && results.data.length > 0) {
            setRawData(results.data);
          } else {
            setFetchError("Data stream empty or connection lost.");
          }
          setLoading(false);
        },
        error: (err) => {
          setFetchError("Connection Failed. Verify secure CSV link.");
          setLoading(false);
        },
      });
    } catch (err) {
      setFetchError("Network latency issue.");
      setLoading(false);
    }
  };

  const filterOptions = useMemo(() => {
    const years = new Set(), months = new Set(), regions = new Set(), districts = new Set(), categories = new Set();
    rawData.forEach((row) => {
      if (row["Year"]) years.add(row["Year"].trim());
      if (row["Month"]) months.add(row["Month"].trim());
      if (row["Region"]) regions.add(row["Region"].trim());
      if (row["District"]) districts.add(row["District"].trim());
      if (row["Category"]) categories.add(row["Category"].trim());
    });
    return {
      years: Array.from(years), months: Array.from(months), regions: Array.from(regions),
      districts: Array.from(districts), categories: Array.from(categories),
    };
  }, [rawData]);

  const filteredData = useMemo(() => {
    return rawData.filter((row) => {
      const matchYear = selectedYear === "All" || row["Year"]?.trim() === selectedYear;
      const matchMonth = selectedMonth === "All" || row["Month"]?.trim() === selectedMonth;
      const matchRegion = selectedRegion === "All" || row["Region"]?.trim() === selectedRegion;
      const matchDistrict = selectedDistrict === "All" || row["District"]?.trim() === selectedDistrict;
      const matchCategory = selectedCategory === "All" || row["Category"]?.trim() === selectedCategory;
      const matchSearch = !searchTerm || 
        row["Fields"]?.toLowerCase().includes(searchTerm.toLowerCase()) || 
        row["District"]?.toLowerCase().includes(searchTerm.toLowerCase());
      return matchYear && matchMonth && matchRegion && matchDistrict && matchCategory && matchSearch;
    });
  }, [rawData, selectedYear, selectedMonth, selectedRegion, selectedDistrict, selectedCategory, searchTerm]);

  const kpiStats = useMemo(() => {
    let totalMuballigh = 0, totalMasjid = 0, totalHalqe = 0;
    filteredData.forEach((item) => {
      const field = item["Fields"]?.trim().toLowerCase() || "";
      const val = parseFloat(item["Report Value"]) || 0;
      if (field.includes("total active muballigh")) totalMuballigh += val;
      if (field.includes("total masjid")) totalMasjid += val;
      if (field.includes("total zeili halqe")) totalHalqe += val;
    });
    return { totalMuballigh, totalMasjid, totalHalqe, totalReports: filteredData.length };
  }, [filteredData]);

  
  
  const downloadExcel = () => {
    if (!filteredData.length) return alert("No data to export");
    const worksheet = XLSX.utils.json_to_sheet(filteredData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Deeni Kaam Data");
    XLSX.writeFile(workbook, `12_Deeni_Kaam_RawData_${new Date().getTime()}.xlsx`);
  };

  const downloadPPT = () => {
    if (!filteredData.length) return alert("No data to export");
    
    let pres = new pptxgen();
    
    let slide1 = pres.addSlide();
    slide1.background = { color: "0a0f1c" };
    slide1.addText("12 DEENI KAAM REPORT", { x: 1, y: 2, w: 8, fontSize: 36, bold: true, color: "22d3ee", align: "center" });
    slide1.addText(`User: ${officeUser?.name || officeUser?.userId || "Admin"}`, { x: 1, y: 3, w: 8, fontSize: 16, color: "94a3b8", align: "center" });
    slide1.addText(`Generated on: ${new Date().toLocaleDateString()}`, { x: 1, y: 3.5, w: 8, fontSize: 12, color: "64748b", align: "center" });

    let slide2 = pres.addSlide();
    slide2.background = { color: "0a0f1c" };
    slide2.addText("Key Performance Indicators", { x: 0.5, y: 0.5, w: 9, fontSize: 24, bold: true, color: "ffffff" });
    
    slide2.addShape(pres.ShapeType.rect, { x: 0.5, y: 1.5, w: 4, h: 1.5, fill: "121929", line: {color: "22d3ee", width: 1} });
    slide2.addText("TOTAL MUBALLIGH", { x: 0.5, y: 1.7, w: 4, fontSize: 12, color: "94a3b8", align: "center" });
    slide2.addText(kpiStats.totalMuballigh.toLocaleString("en-IN"), { x: 0.5, y: 2.2, w: 4, fontSize: 28, bold: true, color: "ffffff", align: "center" });

    slide2.addShape(pres.ShapeType.rect, { x: 5.5, y: 1.5, w: 4, h: 1.5, fill: "121929", line: {color: "e879f9", width: 1} });
    slide2.addText("TOTAL MASAJID", { x: 5.5, y: 1.7, w: 4, fontSize: 12, color: "94a3b8", align: "center" });
    slide2.addText(kpiStats.totalMasjid.toLocaleString("en-IN"), { x: 5.5, y: 2.2, w: 4, fontSize: 28, bold: true, color: "ffffff", align: "center" });

    slide2.addShape(pres.ShapeType.rect, { x: 0.5, y: 3.5, w: 4, h: 1.5, fill: "121929", line: {color: "60a5fa", width: 1} });
    slide2.addText("ZEILI HALQE", { x: 0.5, y: 3.7, w: 4, fontSize: 12, color: "94a3b8", align: "center" });
    slide2.addText(kpiStats.totalHalqe.toLocaleString("en-IN"), { x: 0.5, y: 4.2, w: 4, fontSize: 28, bold: true, color: "ffffff", align: "center" });

    slide2.addShape(pres.ShapeType.rect, { x: 5.5, y: 3.5, w: 4, h: 1.5, fill: "121929", line: {color: "34d399", width: 1} });
    slide2.addText("SYSTEM RECORDS", { x: 5.5, y: 3.7, w: 4, fontSize: 12, color: "94a3b8", align: "center" });
    slide2.addText(kpiStats.totalReports.toLocaleString("en-IN"), { x: 5.5, y: 4.2, w: 4, fontSize: 28, bold: true, color: "ffffff", align: "center" });

    pres.writeFile({ fileName: `12_Deeni_Kaam_PPT_${new Date().getTime()}.pptx` });
  };

  return (
    <div className="min-h-screen bg-[#0a0f1c] text-slate-300 font-sans relative overflow-hidden">
      <div className="fixed inset-0 z-0 pointer-events-none">
        <div className="absolute top-0 right-0 w-[40rem] h-[40rem] bg-indigo-900/10 rounded-full blur-[150px]"></div>
        <div className="absolute bottom-0 left-0 w-[40rem] h-[40rem] bg-cyan-900/10 rounded-full blur-[150px]"></div>
      </div>

      <div className="relative z-10">
        <header className="bg-[#121929]/80 backdrop-blur-xl border-b border-white/5 sticky top-0 z-30">
          <div className="max-w-[1400px] mx-auto px-4 sm:px-6 lg:px-8 flex items-center justify-between h-20">
            <div className="flex items-center gap-4">
              <button onClick={onBack} className="p-2.5 bg-[#1a2333] hover:bg-[#222d42] border border-white/10 text-white rounded-xl shadow-lg transition-colors">
                <ArrowLeft className="w-5 h-5" />
              </button>
              <div className="hidden sm:flex p-2.5 bg-gradient-to-br from-cyan-500 to-blue-600 text-white rounded-xl shadow-[0_0_15px_rgba(6,182,212,0.4)] border border-white/10">
                <Globe className="w-6 h-6" />
              </div>
              <div>
                <h1 className="text-xl font-bold text-white tracking-wide">
                  12 DEENI KAAM <span className="font-light text-cyan-400">NODE</span>
                </h1>
                <p className="text-[10px] font-bold text-slate-500 uppercase tracking-widest mt-0.5">Live Data Synchronization</p>
              </div>
            </div>

            <div className="flex items-center gap-2 sm:gap-4">
              <button onClick={fetchData} disabled={loading} className="flex items-center gap-2 text-xs font-bold tracking-wider uppercase bg-[#1a2333] hover:bg-[#222d42] text-cyan-400 px-3 sm:px-4 py-2.5 rounded-xl border border-cyan-500/20 transition-all active:scale-95 shadow-[0_0_10px_rgba(6,182,212,0.1)]">
                <RefreshCw className={`w-4 h-4 ${loading ? "animate-spin" : ""}`} />
                <span className="hidden sm:inline">Sync Data</span>
              </button>
                            <button onClick={downloadPPT} className="flex items-center gap-2 text-xs font-bold tracking-wider uppercase bg-cyan-900/30 hover:bg-cyan-900/50 text-cyan-400 px-4 py-2.5 rounded-xl border border-cyan-500/30 transition-all active:scale-95">
                <Presentation className="w-4 h-4" /> <span className="hidden sm:inline">PPT</span>
              </button>
                            <button onClick={downloadExcel} className="flex items-center gap-2 text-xs font-bold tracking-wider uppercase bg-teal-900/30 hover:bg-teal-900/50 text-teal-400 px-4 py-2.5 rounded-xl border border-teal-500/30 transition-all active:scale-95">
                <Download className="w-4 h-4" /> <span className="hidden sm:inline">Excel</span>
              </button>
              <button onClick={onLogout} className="flex items-center gap-2 text-xs font-bold tracking-wider uppercase bg-red-500/10 hover:bg-red-500/20 text-red-400 px-3 sm:px-4 py-2.5 rounded-xl border border-red-500/20 transition-all active:scale-95">
                <LogOut className="w-4 h-4" />
                <span className="hidden sm:inline">Logout</span>
              </button>
            </div>
          </div>
        </header>

        <main className="max-w-[1400px] mx-auto px-4 sm:px-6 lg:px-8 py-8 space-y-6">
          <div className="bg-[#121929]/60 backdrop-blur-md p-5 rounded-2xl border border-white/5 flex flex-col md:flex-row gap-4 items-center justify-between">
            <div className="flex items-center gap-3 w-full md:w-auto">
              <div className="p-2 bg-slate-800/50 rounded-lg border border-slate-700">
                <Database className="w-5 h-5 text-cyan-400" />
              </div>
              <span className="font-semibold text-slate-300 text-sm tracking-wide">Data Stream URL</span>
            </div>
            <div className="flex w-full md:w-2/3 gap-3">
              <input type="text" value={sheetUrl} onChange={(e) => setSheetUrl(e.target.value)} placeholder="Paste Google Sheet CSV Published Link..." className="w-full px-4 py-2.5 text-sm rounded-xl border border-slate-700/50 bg-[#0a0f1c] text-white focus:outline-none focus:border-cyan-500/50 focus:ring-1 focus:ring-cyan-500/50 transition-all" />
              <button onClick={fetchData} className="px-6 py-2.5 bg-cyan-600 hover:bg-cyan-500 text-white text-xs tracking-widest font-bold uppercase rounded-xl shrink-0 transition-colors shadow-[0_0_15px_rgba(6,182,212,0.2)]">Connect</button>
            </div>
          </div>

          {fetchError && (
            <div className="p-4 bg-red-900/20 border border-red-500/30 text-red-400 rounded-2xl text-sm font-medium flex items-center gap-3">
              <div className="w-2 h-2 rounded-full bg-red-500 animate-ping"></div>{fetchError}
            </div>
          )}

          <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-6">
            {[
              { title: "Active Muballigh", val: kpiStats.totalMuballigh, icon: User, color: "text-cyan-400", bg: "bg-cyan-400/10", shadow: "shadow-[0_0_20px_rgba(6,182,212,0.15)]", border: "border-cyan-500/20" },
              { title: "Total Masajid", val: kpiStats.totalMasjid, icon: MapPin, color: "text-fuchsia-400", bg: "bg-fuchsia-400/10", shadow: "shadow-[0_0_20px_rgba(232,121,249,0.15)]", border: "border-fuchsia-500/20" },
              { title: "Zeili Halqe", val: kpiStats.totalHalqe, icon: Layers, color: "text-blue-400", bg: "bg-blue-400/10", shadow: "shadow-[0_0_20px_rgba(96,165,250,0.15)]", border: "border-blue-500/20" },
              { title: "System Records", val: kpiStats.totalReports, icon: Activity, color: "text-emerald-400", bg: "bg-emerald-400/10", shadow: "shadow-[0_0_20px_rgba(52,211,153,0.15)]", border: "border-emerald-500/20" }
            ].map((kpi, idx) => (
              <div key={idx} className={`bg-[#121929]/80 backdrop-blur-md p-6 rounded-3xl border ${kpi.border} ${kpi.shadow} relative overflow-hidden group hover:scale-[1.02] transition-transform duration-300`}>
                <div className="relative flex justify-between items-start z-10">
                  <div>
                    <p className="text-xs font-bold text-slate-500 uppercase tracking-widest mb-1">{kpi.title}</p>
                    <h3 className="text-4xl font-extrabold text-white tracking-tight">{kpi.val}</h3>
                  </div>
                  <div className={`p-3.5 rounded-2xl ${kpi.bg} ${kpi.color} border border-white/5`}><kpi.icon className="w-6 h-6" /></div>
                </div>
              </div>
            ))}
          </div>

          <div className="bg-[#121929]/60 backdrop-blur-md p-6 rounded-3xl border border-white/5 shadow-lg">
            <div className="flex items-center gap-3 mb-6">
              <Filter className="w-5 h-5 text-cyan-400" />
              <h2 className="text-sm font-bold text-white tracking-widest uppercase">Data Parameters</h2>
            </div>
            <div className="grid grid-cols-1 md:grid-cols-5 gap-4">
              {[{ label: "Year", value: selectedYear, setter: setSelectedYear, options: filterOptions.years },
                { label: "Month", value: selectedMonth, setter: setSelectedMonth, options: filterOptions.months },
                { label: "Region", value: selectedRegion, setter: setSelectedRegion, options: filterOptions.regions },
                { label: "District", value: selectedDistrict, setter: setSelectedDistrict, options: filterOptions.districts },
                { label: "Category", value: selectedCategory, setter: setSelectedCategory, options: filterOptions.categories }
              ].map((filter, idx) => (
                <div key={idx}>
                  <label className="block text-[10px] font-bold text-slate-500 uppercase tracking-widest mb-2">{filter.label}</label>
                  <select value={filter.value} onChange={(e) => filter.setter(e.target.value)} disabled={officeUser?.[filter.label === "Category" ? "department" : filter.label.toLowerCase()] && officeUser[filter.label === "Category" ? "department" : filter.label.toLowerCase()].toLowerCase() !== "all"} className="w-full p-3 rounded-xl border border-slate-700/50 bg-[#0a0f1c] text-sm font-medium text-slate-300 focus:outline-none focus:border-cyan-500/50 focus:ring-1 focus:ring-cyan-500/50 cursor-pointer appearance-none">
                    <option value="All">All {filter.label}s</option>
                    {filter.options.map((opt) => (<option key={opt} value={opt}>{opt}</option>))}
                  </select>
                </div>
              ))}
            </div>
          </div>

          <div className="bg-[#121929]/60 backdrop-blur-md rounded-3xl border border-white/5 overflow-hidden shadow-lg">
            <div className="p-6 border-b border-white/5 flex flex-col md:flex-row justify-between items-center gap-4 bg-black/20">
              <div>
                <h2 className="text-sm font-bold text-white uppercase tracking-widest">Telemetry Output</h2>
                <p className="text-xs text-slate-500 mt-1">Live grid data monitoring</p>
              </div>
              <div className="w-full md:w-80 relative">
                <span className="absolute inset-y-0 left-0 pl-3 flex items-center text-slate-500"><Search className="w-4 h-4" /></span>
                <input type="text" placeholder="Query fields..." value={searchTerm} onChange={(e) => setSearchTerm(e.target.value)} className="w-full pl-10 pr-4 py-2.5 text-sm bg-[#0a0f1c] border border-slate-700/50 rounded-xl text-white focus:outline-none focus:border-cyan-500/50 focus:ring-1 focus:ring-cyan-500/50 transition-all placeholder-slate-600" />
              </div>
            </div>

            <div className="overflow-x-auto">
              <table className="min-w-full text-left text-sm whitespace-nowrap">
                <thead className="bg-[#0a0f1c]/50">
                  <tr>
                    <th className="px-6 py-4 text-[11px] font-bold text-slate-400 uppercase tracking-widest">Location</th>
                    <th className="px-6 py-4 text-[11px] font-bold text-slate-400 uppercase tracking-widest">Category</th>
                    <th className="px-6 py-4 text-[11px] font-bold text-slate-400 uppercase tracking-widest">Parameter</th>
                    <th className="px-6 py-4 text-[11px] font-bold text-cyan-400 uppercase tracking-widest text-right">Value</th>
                    <th className="px-6 py-4 text-[11px] font-bold text-slate-400 uppercase tracking-widest text-center">Timestamp</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-white/5">
                  {filteredData.length > 0 ? (
                    filteredData.map((row, idx) => (
                      <tr key={idx} className="hover:bg-white/[0.02] transition-colors">
                        <td className="px-6 py-4">
                          <div className="font-bold text-white">{row["District"] || "-"}</div>
                          <div className="text-xs font-medium text-slate-500">{row["Region"] || "-"}</div>
                        </td>
                        <td className="px-6 py-4">
                          <span className="inline-flex px-3 py-1 rounded-md text-[11px] font-bold bg-slate-800 text-slate-300 border border-slate-700">
                            {row["Category"] || "-"}
                          </span>
                        </td>
                        <td className="px-6 py-4 font-medium text-slate-300">{row["Fields"] || "-"}</td>
                        <td className="px-6 py-4 text-right">
                          <span className="text-base font-extrabold text-cyan-400" style={{ textShadow: '0 0 10px rgba(6,182,212,0.4)' }}>
                            {row["Report Value"] || "0"}
                          </span>
                        </td>
                        <td className="px-6 py-4 text-center">
                          <div className="inline-flex items-center gap-1.5 px-3 py-1 bg-slate-800/50 rounded-md text-[11px] font-bold text-slate-400 border border-slate-700/50">
                            <Calendar className="w-3 h-3" /> {row["Month"]} {row["Year"]}
                          </div>
                        </td>
                      </tr>
                    ))
                  ) : (
                    <tr>
                      <td colSpan="5" className="px-6 py-12 text-center text-slate-500">
                        <div className="flex flex-col items-center justify-center gap-3">
                          <Activity className="w-8 h-8 opacity-20" />
                          <p className="text-xs uppercase tracking-widest font-bold">{loading ? "Establishing connection..." : "No matching telemetry found"}</p>
                        </div>
                      </td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>
          </div>
        </main>
      </div>
    </div>
  );
}