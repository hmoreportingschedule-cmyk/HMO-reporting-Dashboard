"use client";
import React, { useState, useEffect, useMemo } from "react";
import Papa from "papaparse";
import { LogOut, RefreshCw, Filter, Search, Activity, ArrowLeft, User, Presentation, Download, UserCheck, MapPin, Building, ShieldCheck, Image as ImageIcon , Clock} from "lucide-react";
import * as XLSX from "xlsx";
import pptxgen from "pptxgenjs";
import html2canvas from "html2canvas";

// ⚠️ YAHAN VOLUNTEERS KI GOOGLE SHEET KA CSV LINK PASTE KAREIN
const DEFAULT_CSV_URL = "PASTE_VOLUNTEERS_CSV_LINK_HERE";

const sameClient = (a, b) => String(a || "").trim().toLowerCase() === String(b || "").trim().toLowerCase();

const uniqValues = (data, key) => {
  return [...new Set(data.map(x => String(x[key] || "").trim()).filter(Boolean))].sort((a,b) => a.localeCompare(b));
};

const groupCount = (data, key) => {
  const map = {};
  data.forEach(x => {
      const k = (x[key] || "Unknown").trim();
      map[k] = (map[k] || 0) + 1;
  });
  return Object.keys(map).sort((a,b) => map[b] - map[a]).map(k => ({ label: k, count: map[k] }));
};

const MiniTable = ({ title, data }) => (
  <div className="bg-white rounded-xl shadow-sm border border-slate-200 overflow-hidden flex flex-col h-[340px]">
    <div className="p-4 border-b border-slate-100">
      <h3 className="text-teal-700 font-bold uppercase tracking-widest text-xs">{title}</h3>
    </div>
    <div className="overflow-y-auto flex-1 custom-scrollbar">
      <table className="w-full text-left text-xs">
        <thead className="bg-teal-700 text-white sticky top-0 z-10">
          <tr>
            <th className="py-2.5 px-4 font-bold uppercase tracking-wider">Name</th>
            <th className="py-2.5 px-4 font-bold uppercase tracking-wider text-right">Count</th>
          </tr>
        </thead>
        <tbody className="divide-y divide-slate-100">
          {data.length > 0 ? data.map((d, i) => (
            <tr key={i} className="hover:bg-slate-50 transition-colors">
              <td className="py-2.5 px-4 text-slate-700">{d.label}</td>
              <td className="py-2.5 px-4 text-teal-700 font-bold text-right">{d.count.toLocaleString("en-IN")}</td>
            </tr>
          )) : <tr><td colSpan="2" className="text-center py-4 text-slate-500">No data available</td></tr>}
        </tbody>
      </table>
    </div>
  </div>
);

export default function VolunteersDashboard({ onBack, officeUser, onLogout }) {
  const [sheetUrl, setSheetUrl] = useState(DEFAULT_CSV_URL);
  const [rawData, setRawData] = useState([]);
  const [loading, setLoading] = useState(false);
  const [fetchError, setFetchError] = useState("");

  const [searchQuery, setSearchQuery] = useState("");
  const [region, setRegion] = useState(officeUser?.region && officeUser.region.toLowerCase() !== "all" ? officeUser.region : "");
  const [state, setState] = useState(officeUser?.state && officeUser.state.toLowerCase() !== "all" ? officeUser.state : "");
  const [division, setDivision] = useState(officeUser?.division && officeUser.division.toLowerCase() !== "all" ? officeUser.division : "");
  const [district, setDistrict] = useState(officeUser?.district && officeUser.district.toLowerCase() !== "all" ? officeUser.district : "");
  const [zimmedari, setZimmedari] = useState("");
  const [chain, setChain] = useState("");

  const [currentPage, setCurrentPage] = useState(1);
  const rowsPerPage = 10;

  const [now, setNow] = useState(null);
  useEffect(() => {
    setNow(new Date());
    const timer = setInterval(() => setNow(new Date()), 1000);
    return () => clearInterval(timer);
  }, []);
  const dateStr = now ? now.toLocaleDateString('en-GB', { day: '2-digit', month: '2-digit', year: 'numeric' }).replace(/\//g, '-') : "";
  const timeStr = now ? now.toLocaleTimeString('en-US', { hour: '2-digit', minute: '2-digit', second: '2-digit', hour12: true }) : "";



  useEffect(() => {
    if (DEFAULT_CSV_URL !== "PASTE_VOLUNTEERS_CSV_LINK_HERE") {
      fetchData();
    }
  }, []);

  const fetchData = async () => {
    if (!sheetUrl || sheetUrl === "PASTE_VOLUNTEERS_CSV_LINK_HERE") {
      setFetchError("System Error: No valid data source provided. Please update CSV link.");
      return;
    }
    setLoading(true);
    setFetchError("");

    Papa.parse(sheetUrl, {
      download: true,
      header: true,
      skipEmptyLines: true,
      complete: (results) => {
        if (results.data && results.data.length > 0) {
          const cleanedData = results.data.filter(r => r["Name"] || r["Mob.No."] || r["Zimmedari Level"]);
          setRawData(cleanedData);
        } else {
          setFetchError("Data stream empty or invalid CSV link.");
        }
        setLoading(false);
      },
      error: () => {
        setFetchError("Connection Failed. Make sure link is 'Publish to Web' as CSV.");
        setLoading(false);
      }
    });
  };

  const handleSetRegion = (v) => {
    setRegion(v);
    if(!(officeUser?.state && officeUser.state.toLowerCase() !== "all")) setState("");
    if(!(officeUser?.division && officeUser.division.toLowerCase() !== "all")) setDivision("");
    if(!(officeUser?.district && officeUser.district.toLowerCase() !== "all")) setDistrict("");
  };
  const handleSetState = (v) => {
    setState(v);
    if(!(officeUser?.division && officeUser.division.toLowerCase() !== "all")) setDivision("");
    if(!(officeUser?.district && officeUser.district.toLowerCase() !== "all")) setDistrict("");
  };
  const handleSetDivision = (v) => {
    setDivision(v);
    if(!(officeUser?.district && officeUser.district.toLowerCase() !== "all")) setDistrict("");
  };

  const filteredData = useMemo(() => {
    return rawData.filter(row => {
      const searchMatch = !searchQuery || JSON.stringify(row).toLowerCase().includes(searchQuery.toLowerCase().trim());
      const regionMatch = !region || sameClient(row["Region"], region);
      const stateMatch = !state || sameClient(row["State"], state);
      const divMatch = !division || sameClient(row["Division"], division);
      const distMatch = !district || sameClient(row["District"], district);
      const zimmMatch = !zimmedari || sameClient(row["Zimmedari Level"], zimmedari);
      const chainMatch = !chain || sameClient(row["Chain"], chain);

      return searchMatch && regionMatch && stateMatch && divMatch && distMatch && zimmMatch && chainMatch;
    });
  }, [rawData, searchQuery, region, state, division, district, zimmedari, chain]);

  useEffect(() => { setCurrentPage(1); }, [searchQuery, region, state, division, district, zimmedari, chain]);

  const pagedRows = useMemo(() => {
    const start = (currentPage - 1) * rowsPerPage;
    return filteredData.slice(start, start + rowsPerPage);
  }, [filteredData, currentPage]);

  const kpiStats = useMemo(() => {
    const totalVolunteers = filteredData.length;
    const ajeerCount = filteredData.filter(r => String(r["Ajeer (Yes/No)"]).trim().toLowerCase() === "yes").length;
    const distinctRegions = new Set(filteredData.map(r => r["Region"]?.trim()).filter(Boolean)).size;
    const distinctDistricts = new Set(filteredData.map(r => r["District"]?.trim()).filter(Boolean)).size;

    return { totalVolunteers, ajeerCount, distinctRegions, distinctDistricts };
  }, [filteredData]);

  // Dynamic Graph Logic
  const { dynamicGraphData, dynamicGraphTitle } = useMemo(() => {
    if (division) return { dynamicGraphData: groupCount(filteredData, "District"), dynamicGraphTitle: "VOLUNTEERS BY DISTRICT" };
    if (state) return { dynamicGraphData: groupCount(filteredData, "Division"), dynamicGraphTitle: "VOLUNTEERS BY DIVISION" };
    if (region) return { dynamicGraphData: groupCount(filteredData, "State"), dynamicGraphTitle: "VOLUNTEERS BY STATE" };
    return { dynamicGraphData: groupCount(filteredData, "Region"), dynamicGraphTitle: "VOLUNTEERS BY REGION" };
  }, [filteredData, region, state, division]);

  const maxDynamicCount = dynamicGraphData.length ? Math.max(...dynamicGraphData.map(d => d.count)) : 0;

  const downloadExcel = () => {
    if (!filteredData.length) return alert("No data to export");
    const worksheet = XLSX.utils.json_to_sheet(filteredData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Volunteers Data");
    XLSX.writeFile(workbook, `Volunteers_RawData_${new Date().getTime()}.xlsx`);
  };

  const downloadPPT = () => {
    if (!filteredData.length) return alert("No data to export");
    let pres = new pptxgen();
    let slide1 = pres.addSlide();
    slide1.background = { color: "ffffff" };
    slide1.addText("VOLUNTEERS & ZIMMEDARAN REPORT", { x: 1, y: 2, w: 8, fontSize: 36, bold: true, color: "008080", align: "center" });
    slide1.addText(`User: ${officeUser?.name || officeUser?.userId || "Admin"}`, { x: 1, y: 3, w: 8, fontSize: 16, color: "334155", align: "center" });
    slide1.addText(`Generated on: ${new Date().toLocaleDateString()}`, { x: 1, y: 3.5, w: 8, fontSize: 12, color: "64748b", align: "center" });

    let slide2 = pres.addSlide();
    slide2.background = { color: "ffffff" };
    slide2.addText("Key Performance Indicators", { x: 0.5, y: 0.5, w: 9, fontSize: 24, bold: true, color: "0f766e" });
    
    slide2.addShape(pres.ShapeType.rect, { x: 0.5, y: 1.5, w: 4, h: 1.5, fill: "f8fafc", line: {color: "008080", width: 1} });
    slide2.addText("TOTAL ZIMMEDARAN", { x: 0.5, y: 1.7, w: 4, fontSize: 12, color: "64748b", align: "center" });
    slide2.addText(kpiStats.totalVolunteers.toLocaleString("en-IN"), { x: 0.5, y: 2.2, w: 4, fontSize: 28, bold: true, color: "0f766e", align: "center" });

    slide2.addShape(pres.ShapeType.rect, { x: 5.5, y: 1.5, w: 4, h: 1.5, fill: "f8fafc", line: {color: "008080", width: 1} });
    slide2.addText("TOTAL AJEER", { x: 5.5, y: 1.7, w: 4, fontSize: 12, color: "64748b", align: "center" });
    slide2.addText(kpiStats.ajeerCount.toLocaleString("en-IN"), { x: 5.5, y: 2.2, w: 4, fontSize: 28, bold: true, color: "0f766e", align: "center" });

    if (dynamicGraphData.length > 0) {
        let slide3 = pres.addSlide();
        slide3.background = { color: "ffffff" };
        slide3.addText(dynamicGraphTitle, { x: 0.5, y: 0.5, w: 9, fontSize: 24, bold: true, color: "0f766e" });
        let chartData = [{ name: "Volunteers", labels: dynamicGraphData.slice(0, 8).map(d => d.label || "Unknown"), values: dynamicGraphData.slice(0, 8).map(d => d.count) }];
        slide3.addChart(pres.ChartType.bar, chartData, { x: 0.5, y: 1.2, w: 9, h: 3.5, barDir: "col", chartColors: ["008080"], valAxisLabelColor: "475569", catAxisLabelColor: "475569", showLegend: false });
    }
    pres.writeFile({ fileName: `Volunteers_PPT_${new Date().getTime()}.pptx` });
  };

  const downloadJPEG = async () => {
    const element = document.getElementById("volunteers-report-card-view");
    if (!element) return;
    try {
        const canvas = await html2canvas(element, { 
            scale: 2, 
            backgroundColor: "#e0f2f1", 
            useCORS: true,
            ignoreElements: (node) => node.hasAttribute("data-html2canvas-ignore")
        });
        const link = document.createElement('a');
        link.download = `Volunteers_Report_Card_${new Date().getTime()}.jpg`;
        link.href = canvas.toDataURL('image/jpeg', 0.9);
        link.click();
    } catch (err) {
        console.error("Failed to capture image", err);
        alert("Image download failed.");
    }
  };

  return (
    <div className="min-h-screen bg-[#e0f2f1] text-slate-800 font-sans relative overflow-hidden">
      <div className="fixed inset-0 z-0 pointer-events-none">
        <div className="absolute top-0 right-0 w-[40rem] h-[40rem] bg-teal-100/40 rounded-full blur-[150px]"></div>
      </div>

      <div className="relative z-10">
        <div id="volunteers-report-card-view" className="bg-[#e0f2f1] pb-6 pt-4">
            <header className="max-w-[1400px] mx-auto px-4 sm:px-6 lg:px-8 relative z-30">
              <div className="bg-white border border-slate-200 shadow-sm rounded-2xl px-6 py-5 flex flex-col md:flex-row items-center justify-between gap-6">
                <div className="flex items-center gap-4 w-full md:w-auto">
                  <button data-html2canvas-ignore="true" onClick={onBack} className="p-3 bg-slate-100 hover:bg-slate-200 border border-slate-200 text-slate-700 rounded-xl shadow-sm transition-colors">
                    <ArrowLeft className="w-5 h-5" />
                  </button>
                  <div className="p-3 bg-teal-700 text-white rounded-xl shadow-sm border border-teal-700 hidden sm:block">
                    <UserCheck className="w-6 h-6" />
                  </div>
                  <div>
                    <h1 className="text-xl sm:text-2xl font-black text-slate-800 tracking-wide">VOLUNTEERS <span className="font-light text-teal-700">DASHBOARD</span></h1>
                    <p className="text-[10px] font-bold text-teal-700 uppercase tracking-widest mt-0.5">Zimmedaran Master Data</p>
                  </div>
                </div>
                
                <div data-html2canvas-ignore="true" className="flex flex-wrap items-center justify-center md:justify-end gap-3 w-full md:w-auto">
                  <div className="text-[10px] font-bold text-slate-600 uppercase tracking-widest bg-slate-50 px-3 py-2 rounded-lg border border-slate-200 hidden lg:flex items-center gap-2 whitespace-nowrap">
                    <Clock className="w-3 h-3 text-teal-600" />
                    <span>Date: {dateStr} &bull; Time: {timeStr}</span>
                  </div>
                  <div className="text-[10px] font-bold text-slate-600 uppercase tracking-widest bg-slate-50 px-3 py-2 rounded-lg border border-slate-200 hidden lg:block">
                    User: <span className="text-teal-700">{officeUser?.name || officeUser?.userId || "Admin"}</span>
                  </div>
                  <button onClick={fetchData} disabled={loading} className="flex items-center gap-2 text-xs font-bold tracking-wider uppercase bg-teal-600 hover:bg-teal-700 text-white px-4 py-2.5 rounded-lg border border-teal-600 transition-all active:scale-95">
                    <RefreshCw className={`w-4 h-4 ${loading ? "animate-spin" : ""}`} /> <span className="hidden sm:inline">Sync</span>
                  </button>
                  <button onClick={downloadJPEG} className="flex items-center gap-2 text-xs font-bold tracking-wider uppercase bg-[#0f766e] hover:bg-[#115e59] text-white px-4 py-2.5 rounded-lg border border-[#0f766e] transition-all active:scale-95 shadow-md">
                    <ImageIcon className="w-4 h-4" /> <span className="hidden sm:inline">JPEG</span>
                  </button>
                  <button onClick={downloadPPT} className="flex items-center gap-2 text-xs font-bold tracking-wider uppercase bg-teal-600 hover:bg-teal-700 text-white px-4 py-2.5 rounded-lg border border-teal-600 transition-all active:scale-95">
                    <Presentation className="w-4 h-4" /> <span className="hidden sm:inline">PPT</span>
                  </button>
                  <button onClick={downloadExcel} className="flex items-center gap-2 text-xs font-bold tracking-wider uppercase bg-[#198754] hover:bg-green-700 text-white px-4 py-2.5 rounded-lg border border-[#198754] transition-all active:scale-95">
                    <Download className="w-4 h-4" /> <span className="hidden sm:inline">Excel</span>
                  </button>
                  <button onClick={onLogout} className="flex items-center gap-2 text-xs font-bold tracking-wider uppercase bg-[#dc3545] hover:bg-red-700 text-white px-4 py-2.5 rounded-lg border border-[#dc3545] transition-all active:scale-95">
                    <LogOut className="w-4 h-4" /> <span className="hidden sm:inline">Logout</span>
                  </button>
                </div>
              </div>
            </header>

            <main className="max-w-[1400px] mx-auto px-4 sm:px-6 lg:px-8 pt-6 space-y-6">
              {fetchError && (
                <div data-html2canvas-ignore="true" className="p-4 bg-red-50 border border-red-200 text-red-600 rounded-xl text-sm font-medium flex items-center gap-3">
                  <div className="w-2 h-2 rounded-full bg-red-500 animate-ping"></div>
                  {fetchError}
                </div>
              )}

              {/* KPIs */}
              <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-6">
                {[
                  { title: "Total Zimmedaran", val: kpiStats.totalVolunteers, icon: UserCheck, color: "text-blue-500", bg: "bg-blue-50", border: "border-blue-100" },
                  { title: "Total Ajeer", val: kpiStats.ajeerCount, icon: ShieldCheck, color: "text-emerald-500", bg: "bg-emerald-50", border: "border-emerald-100" },
                  { title: "Regions Covered", val: kpiStats.distinctRegions, icon: MapPin, color: "text-purple-500", bg: "bg-purple-50", border: "border-purple-100" },
                  { title: "Districts Covered", val: kpiStats.distinctDistricts, icon: Building, color: "text-orange-500", bg: "bg-orange-50", border: "border-orange-100" }
                ].map((kpi, idx) => (
                  <div key={idx} className="bg-white p-6 rounded-xl border border-slate-200 shadow-sm relative overflow-hidden group">
                    <div className="relative flex justify-between items-center z-10">
                      <div className="flex items-center gap-4">
                          <div className={`p-3.5 rounded-xl ${kpi.bg} ${kpi.color} border ${kpi.border}`}>
                            <kpi.icon className="w-6 h-6" />
                          </div>
                          <div>
                              <p className="text-[11px] font-bold text-slate-500 uppercase tracking-widest mb-1.5">{kpi.title}</p>
                              <h3 className={`text-3xl font-black ${idx === 0 ? 'text-blue-800' : idx === 1 ? 'text-emerald-800' : 'text-slate-800'} tracking-tight`}>
                                  {kpi.val.toLocaleString("en-IN")}
                              </h3>
                          </div>
                      </div>
                    </div>
                  </div>
                ))}
              </div>

              {/* Filters */}
              <div className="bg-white p-5 rounded-xl border border-slate-200 shadow-sm">
                <div className="flex items-center gap-3 mb-6">
                  <Filter className="w-5 h-5 text-teal-700" />
                  <h2 className="text-sm font-bold text-slate-800 tracking-widest uppercase">Data Parameters</h2>
                </div>
                <div className="grid grid-cols-2 md:grid-cols-7 gap-4">
                  <div className="col-span-2 md:col-span-1" data-html2canvas-ignore="true">
                    <label className="block text-[10px] font-bold text-slate-500 uppercase tracking-widest mb-2">Search</label>
                    <div className="relative">
                      <span className="absolute inset-y-0 left-0 pl-3 flex items-center text-slate-400"><Search className="w-4 h-4" /></span>
                      <input type="text" placeholder="Search..." value={searchQuery} onChange={(e) => setSearchQuery(e.target.value)} className="w-full pl-9 pr-3 py-2.5 rounded-lg border border-slate-300 bg-white text-sm font-medium text-slate-700 focus:outline-none focus:border-teal-500 focus:ring-1 focus:ring-teal-500" />
                    </div>
                  </div>
                  {[
                    { label: "Region", val: region, set: handleSetRegion, opts: uniqValues(rawData, "Region") },
                    { label: "State", val: state, set: handleSetState, opts: uniqValues(rawData.filter(x=>!region||sameClient(x["Region"],region)), "State") },
                    { label: "Division", val: division, set: handleSetDivision, opts: uniqValues(rawData.filter(x=>(!region||sameClient(x["Region"],region))&&(!state||sameClient(x["State"],state))), "Division") },
                    { label: "District", val: district, set: setDistrict, opts: uniqValues(rawData.filter(x=>(!region||sameClient(x["Region"],region))&&(!state||sameClient(x["State"],state))&&(!division||sameClient(x["Division"],division))), "District") },
                    { label: "Chain", val: chain, set: setChain, opts: uniqValues(rawData, "Chain") },
                    { label: "Zimmedari Level", val: zimmedari, set: setZimmedari, opts: uniqValues(rawData, "Zimmedari Level") }
                  ].map((f, i) => (
                    <div key={i} className="flex flex-col flex-1 min-w-[130px]">
                      <label className="block text-[10px] font-bold text-slate-500 uppercase tracking-widest mb-2">{f.label}</label>
                      <select value={f.val} onChange={(e) => f.set(e.target.value)} disabled={officeUser?.[f.label.toLowerCase()] && officeUser[f.label.toLowerCase()].toLowerCase() !== "all"} className="w-full p-2.5 rounded-lg border border-slate-300 bg-white text-sm font-medium text-slate-700 focus:outline-none focus:border-teal-500 focus:ring-1 focus:ring-teal-500 appearance-none cursor-pointer disabled:bg-slate-50 disabled:text-slate-400">
                        <option value="">{f.val || "All"}</option>
                        {f.opts.map(o => <option key={o} value={o}>{o}</option>)}
                      </select>
                    </div>
                  ))}
                </div>
              </div>

              {/* Charts Grid */}
              <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
                 <div className="bg-white rounded-xl shadow-sm border border-slate-200 p-5 h-[340px] flex flex-col">
                    <h3 className="text-xs font-bold text-teal-700 uppercase tracking-widest mb-4">{dynamicGraphTitle}</h3>
                    <div className="flex-1 flex w-full pt-2">
                      {dynamicGraphData.length > 0 ? (
                        <>
                          {/* Y-Axis Column */}
                          <div className="flex flex-col justify-between items-end pr-3 border-r border-slate-300 pb-8 text-[10px] font-bold text-slate-500 w-12 shrink-0">
                            <span>{maxDynamicCount.toLocaleString("en-IN")}</span>
                            <span>{Math.round(maxDynamicCount / 2).toLocaleString("en-IN")}</span>
                            <span>0</span>
                          </div>
                          {/* Chart Area */}
                          <div className="flex-1 flex justify-start items-end gap-4 sm:gap-6 pl-4 pb-8 relative border-b border-slate-300 overflow-x-auto custom-scrollbar">
                            {dynamicGraphData.slice(0, 10).map((d, i) => {
                               const h = maxDynamicCount ? (d.count / maxDynamicCount) * 100 : 0;
                               const bgColor = i === 0 ? "bg-[#064e3b]" : i === 1 ? "bg-[#1e3a8a]" : i === 2 ? "bg-[#581c87]" : "bg-[#9a3412]";
                               return (
                                 <div key={i} className="flex flex-col justify-end items-center relative h-full w-10 sm:w-14 shrink-0">
                                    <span className="text-[11px] font-bold text-slate-800 mb-1.5">{d.count.toLocaleString("en-IN")}</span>
                                    <div style={{height: `${Math.max(h, 2)}%`}} className={`w-8 sm:w-12 ${bgColor} rounded-t-md transition-all hover:opacity-80`} />
                                    <span className="absolute -bottom-7 w-20 text-center text-[9px] text-slate-600 truncate px-1 font-medium">{d.label}</span>
                                 </div>
                               )
                            })}
                          </div>
                        </>
                      ) : <div className="w-full text-center text-slate-400 text-xs my-auto font-medium">No data available to display</div>}
                    </div>
                 </div>
                 
                 <div className="grid grid-cols-1 md:grid-cols-2 gap-6 h-[340px] overflow-hidden">
                   <MiniTable title="BY ZIMMEDARI LEVEL" data={groupCount(filteredData, "Zimmedari Level")} />
                   <MiniTable title="BY CHAIN" data={groupCount(filteredData, "Chain")} />
                 </div>
              </div>
            </main>
        </div> {/* END OF REPORT CARD VIEW */}

        <main className="max-w-[1400px] mx-auto px-4 sm:px-6 lg:px-8 pb-8 mt-2">
          {/* Table Section - Excluded from JPEG */}
          <div className="bg-white rounded-xl border border-slate-200 shadow-sm mt-6">
            <div className="p-4 border-b border-slate-200 flex flex-col md:flex-row justify-between items-center gap-4 bg-[#e0f2f1] rounded-t-xl">
              <div>
                <h2 className="text-sm font-bold text-teal-800 uppercase tracking-widest">Master Data List</h2>
                <p className="text-[10px] text-slate-600 mt-1 uppercase tracking-widest">Source: {rawData.length} &bull; Visible: {filteredData.length}</p>
              </div>
            </div>

            <div className="overflow-x-auto w-full custom-scrollbar">
              <table className="min-w-full text-left text-[11px] lg:text-xs">
                <thead className="bg-[#008b8b]">
                  <tr>
                    {["Name", "Contact Details", "Hierarchy Location", "Chain & Level", "Ajeer Status", "Zimmedari Tenure"].map(h => (
                      <th key={h} className="px-4 py-3 font-bold text-white uppercase tracking-wider">{h}</th>
                    ))}
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {pagedRows.length > 0 ? pagedRows.map((row, idx) => (
                    <tr key={idx} className="hover:bg-slate-50 transition-colors">
                      <td className="px-4 py-3 text-slate-800 whitespace-nowrap">
                          <div className="font-extrabold text-sm">{row["Name"] || "-"}</div>
                          <div className="text-[9px] text-slate-500 font-bold uppercase mt-0.5">{row["Current Location (Name)"] || "Unknown Location"}</div>
                      </td>
                      <td className="px-4 py-3 text-slate-600 whitespace-nowrap">
                          <div><span className="font-bold text-slate-700">M:</span> {row["Mob.No."] || "-"}</div>
                          {row["Wht.No."] && <div className="mt-0.5"><span className="font-bold text-green-700">W:</span> {row["Wht.No."]}</div>}
                          {row["Email Id"] && <div className="text-[10px] mt-0.5 text-blue-600">{row["Email Id"]}</div>}
                      </td>
                      <td className="px-4 py-3 text-slate-600 leading-tight">
                         <span className="font-bold text-teal-700">{row["Region"] || "-"}</span> <br/>
                         {row["State"]}, {row["Division"]}, {row["District"]}
                      </td>
                      <td className="px-4 py-3 whitespace-nowrap">
                          <span className="inline-block bg-slate-100 border border-slate-200 px-2 py-1 rounded text-[10px] uppercase font-bold text-slate-600 mr-2">{row["Chain"] || "-"}</span>
                          <span className="inline-block bg-teal-50 border border-teal-200 px-2 py-1 rounded text-[10px] uppercase font-bold text-teal-700">{row["Zimmedari Level"] || "-"}</span>
                      </td>
                      <td className="px-4 py-3">
                          <span className={`px-2 py-1 rounded text-[10px] font-bold uppercase border ${String(row["Ajeer (Yes/No)"]).toLowerCase().includes("yes") ? "bg-emerald-50 text-emerald-700 border-emerald-200" : "bg-slate-50 text-slate-500 border-slate-200"}`}>
                             {row["Ajeer (Yes/No)"] || "No"}
                          </span>
                      </td>
                      <td className="px-4 py-3 text-slate-600 leading-tight text-[10px]">
                         <div><span className="font-bold">Assigned:</span> {row["Zimmedari Assign Kab Se Huee Hai"] || "-"}</div>
                         <div><span className="font-bold">Until:</span> {row["Kab Tak Zimmedari Par The"] || "-"}</div>
                      </td>
                    </tr>
                  )) : (
                    <tr>
                      <td colSpan="6" className="px-6 py-12 text-center text-slate-500 text-xs uppercase tracking-widest">
                        <div className="flex flex-col items-center justify-center gap-3">
                          <Activity className="w-6 h-6 opacity-40 text-teal-600" />
                          <span>{loading ? "Fetching records..." : "No matching records found"}</span>
                        </div>
                      </td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>
            {filteredData.length > rowsPerPage && (
              <div className="p-4 border-t border-slate-200 bg-[#f8fafc] rounded-b-xl flex justify-center items-center gap-2">
                <button disabled={currentPage===1} onClick={()=>setCurrentPage(p=>p-1)} className="px-3 py-1.5 rounded-lg border border-teal-600 bg-white text-xs font-bold text-teal-700 disabled:opacity-30 shadow-sm transition-all hover:bg-teal-50">PREV</button>
                <span className="text-xs font-bold text-teal-700 px-4">PAGE {currentPage} OF {Math.ceil(filteredData.length / rowsPerPage)}</span>
                <button disabled={currentPage===Math.ceil(filteredData.length / rowsPerPage)} onClick={()=>setCurrentPage(p=>p+1)} className="px-3 py-1.5 rounded-lg border border-teal-600 bg-white text-xs font-bold text-teal-700 disabled:opacity-30 shadow-sm transition-all hover:bg-teal-50">NEXT</button>
              </div>
            )}
          </div>
        </main>
      </div>
    </div>
  );
}
