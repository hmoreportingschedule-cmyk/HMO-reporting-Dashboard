"use client";
import React, { useState, useEffect, useMemo } from "react";
import Papa from "papaparse";
import { ArrowLeft, BookOpenCheck, LogOut, RefreshCw, Download, Presentation, FileText, Clock, LayoutDashboard, Activity, Filter, Search, Image as ImageIcon } from "lucide-react";
import pptxgen from "pptxgenjs";
import * as XLSX from "xlsx";
import html2canvas from "html2canvas";

const SCRIPT_URL = "https://script.google.com/macros/s/AKfycbyF74lC0dNiWUalx0G7GEK3F802IMBMXfuCsqfuCi5-QYuGkOh-85R_BzK9U_O9mfkpUA/exec";
const DEFAULT_CSV_URL = "https://docs.google.com/spreadsheets/d/1GfMa7j1TIx17jG0g25tdEwU2YgHCm9_fbcrWIUTrSeI/gviz/tq?tqx=out:csv&sheet=Responses";

const sameClient = (a, b) => String(a || "").trim().toLowerCase() === String(b || "").trim().toLowerCase();

const groupCount = (data, key) => {
  const map = {};
  data.forEach(x => {
      const k = (x[key] || "Unknown").trim();
      const n = Number(String(x.report || "").replace(/,/g, ""));
      map[k] = (map[k] || 0) + (isNaN(n) ? 0 : n);
  });
  return Object.keys(map).sort((a,b) => map[b] - map[a]).map(k => ({ label: k, count: map[k] }));
};

const uniqValues = (data, key) => {
  return [...new Set(data.map(x => String(x[key] || "").trim()).filter(Boolean))].sort((a,b) => a.localeCompare(b));
};

const MiniTable = ({ title, data }) => (
  <div className="bg-white rounded-xl shadow-sm border border-slate-200 overflow-hidden flex flex-col h-[320px]">
    <div className="p-4 border-b border-slate-100">
      <h3 className="text-teal-700 font-bold uppercase tracking-widest text-sm">{title}</h3>
    </div>
    <div className="overflow-y-auto flex-1 custom-scrollbar">
      <table className="w-full text-left text-sm">
        <thead className="bg-teal-700 text-white sticky top-0 z-10">
          <tr>
            <th className="py-2.5 px-4 font-bold text-xs uppercase tracking-wider">Name</th>
            <th className="py-2.5 px-4 font-bold text-xs uppercase tracking-wider text-right">Report Qty</th>
          </tr>
        </thead>
        <tbody className="divide-y divide-slate-100">
          {data.length > 0 ? data.map((d, i) => (
            <tr key={i} className="hover:bg-slate-50 transition-colors">
              <td className="py-2.5 px-4 text-slate-700">{d.label}</td>
              <td className="py-2.5 px-4 text-teal-700 font-bold text-right">{d.count.toLocaleString("en-IN")}</td>
            </tr>
          )) : <tr><td colSpan="2" className="text-center py-4 text-slate-500 text-xs">No data available</td></tr>}
        </tbody>
      </table>
    </div>
  </div>
);

export default function WeeklyRisalaDashboard({ onBack, officeUser, onLogout }) {
  const [config, setConfig] = useState({ officeStatus: "ON", b1Value: "", b2Value: "", offMessage: "", risalaNo: "" });
  const [isConfigLoaded, setIsConfigLoaded] = useState(false);
  const [dashboardData, setDashboardData] = useState({ totalRows: 0 });
  const [rows, setRows] = useState([]);
  const [loading, setLoading] = useState(true);
  const [fetchError, setFetchError] = useState("");

  const [searchQuery, setSearchQuery] = useState("");
  const [region, setRegion] = useState(officeUser?.region && officeUser.region.toLowerCase() !== "all" ? officeUser.region : "");
  const [state, setState] = useState(officeUser?.state && officeUser.state.toLowerCase() !== "all" ? officeUser.state : "");
  const [division, setDivision] = useState(officeUser?.division && officeUser.division.toLowerCase() !== "all" ? officeUser.division : "");
  const [district, setDistrict] = useState(officeUser?.district && officeUser.district.toLowerCase() !== "all" ? officeUser.district : "");
  const [department, setDepartment] = useState(officeUser?.department && officeUser.department.toLowerCase() !== "all" ? officeUser.department : "");
  const [chain, setChain] = useState(officeUser?.chain && officeUser.chain.toLowerCase() !== "all" ? officeUser.chain : "");

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
    loadConfigAndData();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const loadConfigAndData = async () => {
    setLoading(true);
    setFetchError("");
    try {
      const res = await fetch(`${SCRIPT_URL}?action=getRisalaConfig`);
      const data = await res.json();
      if (data.status === "Success") {
        setConfig({
          officeStatus: String(data.officeStatus).toUpperCase().trim() || "OFF",
          b1Value: data.b1Value || "",
          b2Value: data.b2Value || "",
          offMessage: data.offMessage || "Dashboard is closed.",
          risalaNo: data.risalaNo || ""
        });
        setIsConfigLoaded(true);

        if (String(data.officeStatus).toUpperCase().trim() === "ON") {
          fetchData();
        } else {
          setLoading(false);
        }
      } else {
        setConfig({ officeStatus: "ON", b1Value: "", b2Value: "", offMessage: "", risalaNo: "" });
        setIsConfigLoaded(true);
        fetchData();
      }
    } catch(e) {
      setConfig({ officeStatus: "ON", b1Value: "", b2Value: "", offMessage: "", risalaNo: "" });
      setIsConfigLoaded(true);
      fetchData();
    }
  };

  const fetchData = async () => {
    setLoading(true);
    Papa.parse(DEFAULT_CSV_URL, {
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
              else if (nKey.includes('chain')) {
              let cVal = String(row[key] || "").trim();
              newRow.chain = cVal ? cVal.charAt(0).toUpperCase() + cVal.slice(1).toLowerCase() : "";
            }
              else if (nKey.includes('level') || nKey.includes('nigran') || nKey.includes('zimmedar')) newRow.level = row[key];
              else if (nKey.includes('department')) newRow.department = row[key];
              else if (nKey.includes('report') || nKey.includes('qty')) newRow.report = row[key];
              else if (nKey.includes('district')) newRow.district = row[key];
              else if (nKey.includes('division')) newRow.division = row[key];
              else if (nKey.includes('state')) newRow.state = row[key];
              else if (nKey.includes('region')) newRow.region = row[key];
              else if (nKey.includes('pincode') || nKey.includes('pin')) newRow.pincode = row[key];
              else if (nKey === 'country') newRow.country = row[key];
            }
            newRow.nigran = newRow.level;
            newRow.zimmedar = newRow.level;
            return newRow;
          }).filter(r => Object.keys(r).length > 1);
          
          setRows(mappedRows);
          setDashboardData({ totalRows: mappedRows.length });
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

  const filteredRows = useMemo(() => {
    let result = rows.filter(x =>
        (!region || sameClient(x.region, region)) &&
        (!state || sameClient(x.state, state)) &&
        (!division || sameClient(x.division, division)) &&
        (!district || sameClient(x.district, district)) &&
        (!department || sameClient(x.department, department)) &&
        (!chain || sameClient(x.chain, chain)) &&
        (!searchQuery || JSON.stringify(x).toLowerCase().includes(searchQuery.toLowerCase().trim()))
    );
    return result;
  }, [rows, searchQuery, region, state, division, district, department, chain]);

  useEffect(() => { setCurrentPage(1); }, [searchQuery, region, state, division, district, department, chain]);

  const pagedRows = useMemo(() => {
    const start = (currentPage - 1) * rowsPerPage;
    return filteredRows.slice(start, start + rowsPerPage);
  }, [filteredRows, currentPage]);

  const totalReportSum = useMemo(() => {
    return filteredRows.reduce((sum, x) => {
        const n = Number(String(x.report || "").replace(/,/g, ""));
        return sum + (isNaN(n) ? 0 : n);
    }, 0);
  }, [filteredRows]);

  const { dynamicGraphData, dynamicGraphTitle } = useMemo(() => {
    if (division) return { dynamicGraphData: groupCount(filteredRows, "district"), dynamicGraphTitle: "REPORTS BY DISTRICT" };
    if (state) return { dynamicGraphData: groupCount(filteredRows, "division"), dynamicGraphTitle: "REPORTS BY DIVISION" };
    if (region) return { dynamicGraphData: groupCount(filteredRows, "state"), dynamicGraphTitle: "REPORTS BY STATE" };
    return { dynamicGraphData: groupCount(filteredRows, "region"), dynamicGraphTitle: "REPORTS BY REGION" };
  }, [filteredRows, region, state, division]);

  const maxDynamicCount = dynamicGraphData.length ? Math.max(...dynamicGraphData.map(d => d.count)) : 0;

  const downloadPPT = () => {
    if (!filteredRows.length) return alert("No data to export");
    let pres = new pptxgen();
    let slide1 = pres.addSlide();
    slide1.background = { color: "ffffff" };
    slide1.addText("WEEKLY RISALA REPORT", { x: 1, y: 2, w: 8, fontSize: 36, bold: true, color: "008080", align: "center" });
    slide1.addText(`User: ${officeUser?.name || officeUser?.userId || "Admin"}`, { x: 1, y: 3, w: 8, fontSize: 16, color: "334155", align: "center" });
    slide1.addText(`Generated on: ${new Date().toLocaleDateString()}`, { x: 1, y: 3.5, w: 8, fontSize: 12, color: "64748b", align: "center" });

    let slide2 = pres.addSlide();
    slide2.background = { color: "ffffff" };
    slide2.addText("Key Performance Indicators", { x: 0.5, y: 0.5, w: 9, fontSize: 24, bold: true, color: "0f766e" });
    slide2.addShape(pres.ShapeType.rect, { x: 1, y: 1.5, w: 3.5, h: 2, fill: "f8fafc", line: {color: "008080", width: 1} });
    slide2.addText("TOTAL SUBMITTED", { x: 1, y: 1.8, w: 3.5, fontSize: 14, color: "64748b", align: "center" });
    slide2.addText(filteredRows.length.toLocaleString("en-IN"), { x: 1, y: 2.3, w: 3.5, fontSize: 32, bold: true, color: "0f766e", align: "center" });
    slide2.addShape(pres.ShapeType.rect, { x: 5.5, y: 1.5, w: 3.5, h: 2, fill: "f8fafc", line: {color: "008080", width: 1} });
    slide2.addText("REPORT QUANTITY", { x: 5.5, y: 1.8, w: 3.5, fontSize: 14, color: "64748b", align: "center" });
    slide2.addText(totalReportSum.toLocaleString("en-IN"), { x: 5.5, y: 2.3, w: 3.5, fontSize: 32, bold: true, color: "0f766e", align: "center" });

    if (dynamicGraphData.length > 0) {
        let slide3 = pres.addSlide();
        slide3.background = { color: "ffffff" };
        slide3.addText(dynamicGraphTitle, { x: 0.5, y: 0.5, w: 9, fontSize: 24, bold: true, color: "0f766e" });
        let chartData = [{ name: "Reports", labels: dynamicGraphData.slice(0, 8).map(d => d.label || "Unknown"), values: dynamicGraphData.slice(0, 8).map(d => d.count) }];
        slide3.addChart(pres.ChartType.bar, chartData, { x: 0.5, y: 1.2, w: 9, h: 3.5, barDir: "col", chartColors: ["008080"], valAxisLabelColor: "475569", catAxisLabelColor: "475569", showLegend: false });
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
        "District": x.district || "",
        "Division": x.division || "",
        "State": x.state || "",
        "Region": x.region || "",
        "Country": x.country || "",
        "Risala Report": x.report || ""
    }));
    const worksheet = XLSX.utils.json_to_sheet(exportData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Weekly Risala Data");
    XLSX.writeFile(workbook, `Weekly_Risala_RawData_${new Date().getTime()}.xlsx`);
  };

  const downloadJPEG = async () => {
    const element = document.getElementById("report-card-view");
    if (!element) return;
    try {
        const canvas = await html2canvas(element, { 
            scale: 2, 
            backgroundColor: "#e0f2f1", 
            useCORS: true,
            ignoreElements: (node) => node.hasAttribute("data-html2canvas-ignore")
        });
        const link = document.createElement('a');
        link.download = `Risala_Report_Card_${new Date().getTime()}.jpg`;
        link.href = canvas.toDataURL('image/jpeg', 0.9);
        link.click();
    } catch (err) {
        console.error("Failed to capture image", err);
        alert("Image download failed.");
    }
  };

  if (isConfigLoaded && config.officeStatus === "OFF") {
    return (
      <div className="min-h-screen flex items-center justify-center bg-[#e0f2f1] p-4 font-sans">
        <div className="bg-white p-8 rounded-xl shadow-lg border-t-[6px] border-red-500 max-w-lg w-full text-center">
          <h2 className="text-2xl font-bold text-red-600 mb-4">Dashboard Closed</h2>
          <p className="text-slate-700 font-medium mb-8 leading-relaxed">{config.offMessage}</p>
          <button onClick={onBack} className="bg-teal-600 hover:bg-teal-700 text-white font-bold py-2.5 px-8 rounded-lg transition-colors shadow-md">
            Return to Hub
          </button>
        </div>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-[#e0f2f1] text-slate-800 font-sans relative overflow-hidden">
      <div className="fixed inset-0 z-0 pointer-events-none">
        <div className="absolute top-0 right-0 w-[40rem] h-[40rem] bg-teal-100/40 rounded-full blur-[150px]"></div>
      </div>

      <div className="relative z-10">
        {/* REPORT CARD VIEW CONTAINER */}
        <div id="report-card-view" className="bg-[#e0f2f1] pb-6 pt-4">
            <header className="max-w-[1400px] mx-auto px-4 sm:px-6 lg:px-8 relative z-30">
              <div className="bg-white border border-slate-200 shadow-sm rounded-2xl px-6 py-5 flex flex-col md:flex-row items-center justify-between gap-6">
                <div className="flex items-center gap-4 w-full md:w-auto">
                  <button data-html2canvas-ignore="true" onClick={onBack} className="p-3 bg-slate-100 hover:bg-slate-200 border border-slate-200 text-slate-700 rounded-xl shadow-sm transition-colors">
                    <ArrowLeft className="w-5 h-5" />
                  </button>
                  <div className="p-3 bg-teal-700 text-white rounded-xl shadow-sm border border-teal-700 hidden sm:block">
                    <BookOpenCheck className="w-6 h-6" />
                  </div>
                  <div>
                    <h1 className="text-xl sm:text-2xl font-black text-slate-800 tracking-wide">WEEKLY RISALA <span className="font-light text-teal-700">REPORT</span></h1>
                    {(config.b1Value || config.b2Value) ? (
                        <p className="text-xs font-bold text-teal-700 uppercase tracking-widest mt-1">
                            {config.b1Value} {config.b1Value && config.b2Value && <span className="text-slate-400 mx-1">:</span>} {config.b2Value}
                        </p>
                    ) : (
                        <p className="text-[10px] font-bold text-teal-700 uppercase tracking-widest mt-0.5">Live Data Synchronization</p>
                    )}
                  </div>
                </div>
                
                <div data-html2canvas-ignore="true" className="flex flex-wrap items-center justify-center md:justify-end gap-3 w-full md:w-auto">
                  <div className="text-[11px] font-extrabold text-black uppercase tracking-widest bg-white px-4 py-2.5 rounded-lg border border-slate-200 hidden lg:flex items-center gap-2 whitespace-nowrap shadow-sm">
                    <Clock className="w-4 h-4 text-teal-600" />
                    <span>Date: {dateStr} &bull; Time: {timeStr}</span>
                  </div>
                  <div className="text-[10px] font-bold text-slate-600 uppercase tracking-widest bg-slate-50 px-3 py-2 rounded-lg border border-slate-200 hidden lg:block">
                    User: <span className="text-teal-700">{officeUser?.name || officeUser?.userId || "Admin"}</span>
                  </div>
                  <button onClick={() => fetchData()} disabled={loading} className="flex items-center gap-2 text-xs font-bold tracking-wider uppercase bg-teal-600 hover:bg-teal-700 text-white px-4 py-2.5 rounded-lg border border-teal-600 transition-all active:scale-95">
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
              <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                <div className="bg-white p-6 rounded-xl border border-slate-200 shadow-sm relative overflow-hidden group">
                  <div className="relative flex justify-between items-center z-10">
                    <div className="flex items-center gap-4">
                        <div className="p-4 rounded-xl bg-amber-50 text-amber-500 border border-amber-100">
                          <LayoutDashboard className="w-7 h-7" />
                        </div>
                        <div>
                            <p className="text-xs font-bold text-slate-500 uppercase tracking-widest mb-1.5">TOTAL SUBMITTED</p>
                            <h3 className="text-5xl font-black text-teal-800 tracking-tight">{filteredRows.length.toLocaleString("en-IN")}</h3>
                        </div>
                    </div>
                  </div>
                </div>
                <div className="bg-white p-6 rounded-xl border border-slate-200 shadow-sm relative overflow-hidden group">
                  <div className="relative flex justify-between items-center z-10">
                    <div className="flex items-center gap-4">
                        <div className="p-4 rounded-xl bg-teal-50 text-teal-500 border border-teal-100">
                          <Activity className="w-7 h-7" />
                        </div>
                        <div>
                            <p className="text-xs font-bold text-slate-500 uppercase tracking-widest mb-1.5">REPORT QUANTITY</p>
                            <h3 className="text-5xl font-black text-teal-800 tracking-tight">{totalReportSum.toLocaleString("en-IN")}</h3>
                        </div>
                    </div>
                  </div>
                </div>
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
                    { label: "Region", val: region, set: handleSetRegion, opts: uniqValues(rows, "region") },
                    { label: "State", val: state, set: handleSetState, opts: uniqValues(rows.filter(x=>!region||sameClient(x.region,region)), "state") },
                    { label: "Division", val: division, set: handleSetDivision, opts: uniqValues(rows.filter(x=>(!region||sameClient(x.region,region))&&(!state||sameClient(x.state,state))), "division") },
                    { label: "District", val: district, set: setDistrict, opts: uniqValues(rows.filter(x=>(!region||sameClient(x.region,region))&&(!state||sameClient(x.state,state))&&(!division||sameClient(x.division,division))), "district") },
                    { label: "Department", val: department, set: setDepartment, opts: uniqValues(rows, "department") },
                    { label: "Chain", val: chain, set: setChain, opts: uniqValues(rows, "chain") }
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
                 
                 <MiniTable title="REPORTS BY STATE" data={groupCount(filteredRows, "state")} />
                 <MiniTable title="REPORTS BY DEPARTMENT" data={groupCount(filteredRows, "department")} />
                 <MiniTable title="REPORTS BY DIVISION" data={groupCount(filteredRows, "division")} />
              </div>
            </main>
        </div> {/* END OF REPORT CARD VIEW */}

        <main className="max-w-[1400px] mx-auto px-4 sm:px-6 lg:px-8 pb-8 mt-2">
          {/* Optimized Table Section - Excluded from JPEG */}
          <div className="bg-white rounded-xl border border-slate-200 shadow-sm">
            <div className="p-4 border-b border-slate-200 flex flex-col md:flex-row justify-between items-center gap-4 bg-slate-50 rounded-t-xl">
              <div>
                <h2 className="text-sm font-bold text-teal-800 uppercase tracking-widest">Detailed Telemetry Output</h2>
                <p className="text-[10px] text-slate-600 mt-1 uppercase tracking-widest">Source: {dashboardData.totalRows} &bull; Visible: {filteredRows.length}</p>
              </div>
            </div>
            <div className="overflow-x-auto w-full custom-scrollbar">
              <table className="w-full text-left text-[11px] lg:text-xs">
                <thead className="bg-[#008b8b]">
                  <tr>
                    {["Date", "Name", "Contact", "Chain", "Level", "Department", "District", "Division", "State", "Region", "Report"].map(h => (
                      <th key={h} className={`px-2 py-3 font-bold text-white uppercase tracking-wider ${h==="Report"?"text-right":""}`}>{h}</th>
                    ))}
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {pagedRows.length > 0 ? pagedRows.map((row, idx) => (
                    <tr key={idx} className="hover:bg-slate-50 transition-colors">
                      <td className="px-2 py-2 text-slate-600 whitespace-nowrap">{row.date}</td>
                      <td className="px-2 py-2 font-bold text-slate-800 leading-tight">{row.name}</td>
                      <td className="px-2 py-2 text-slate-600 whitespace-nowrap">{row.contact}</td>
                      <td className="px-2 py-2 uppercase font-bold text-slate-500 whitespace-nowrap"><span className="bg-slate-100 border border-slate-200 px-1.5 py-0.5 rounded text-[9px]">{row.chain}</span></td>
                      <td className="px-2 py-2 text-slate-600 leading-tight">{row.nigran || row.zimmedar || row.level}</td>
                      <td className="px-2 py-2 text-slate-700 leading-tight">{row.department}</td>
                      <td className="px-2 py-2 text-slate-600 leading-tight">{row.district}</td>
                      <td className="px-2 py-2 text-slate-600 leading-tight">{row.division}</td>
                      <td className="px-2 py-2 text-slate-600 leading-tight">{row.state}</td>
                      <td className="px-2 py-2 text-slate-600 leading-tight">{row.region}</td>
                      <td className="px-2 py-2 text-right font-extrabold text-teal-700 text-sm whitespace-nowrap bg-teal-50/40">{row.report}</td>
                    </tr>
                  )) : (
                    <tr><td colSpan="11" className="px-6 py-12 text-center text-slate-500 text-xs uppercase tracking-widest">{loading ? "Fetching records..." : "No matching records found"}</td></tr>
                  )}
                </tbody>
              </table>
            </div>
            {filteredRows.length > rowsPerPage && (
              <div className="p-4 border-t border-slate-200 bg-[#f8fafc] rounded-b-xl flex justify-center items-center gap-2">
                <button disabled={currentPage===1} onClick={()=>setCurrentPage(p=>p-1)} className="px-3 py-1.5 rounded-lg border border-teal-600 bg-white text-xs font-bold text-teal-700 disabled:opacity-30 shadow-sm transition-all hover:bg-teal-50">PREV</button>
                <span className="text-xs font-bold text-teal-700 px-4">PAGE {currentPage} OF {Math.ceil(filteredRows.length / rowsPerPage)}</span>
                <button disabled={currentPage===Math.ceil(filteredRows.length / rowsPerPage)} onClick={()=>setCurrentPage(p=>p+1)} className="px-3 py-1.5 rounded-lg border border-teal-600 bg-white text-xs font-bold text-teal-700 disabled:opacity-30 shadow-sm transition-all hover:bg-teal-50">NEXT</button>
              </div>
            )}
          </div>
        </main>
      </div>
    </div>
  );
}
