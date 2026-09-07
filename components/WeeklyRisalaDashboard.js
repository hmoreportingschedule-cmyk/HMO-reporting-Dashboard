"use client";
import React, { useState, useEffect, useMemo } from "react";
import Papa from "papaparse";
import { ArrowLeft, BookOpenCheck, LogOut, RefreshCw, Download, Presentation, FileText, Clock, LayoutDashboard, Activity, Filter } from "lucide-react";
import pptxgen from "pptxgenjs";
import * as XLSX from "xlsx";

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
  const [config, setConfig] = useState({ officeStatus: "ON", risalaName: "Loading...", offMessage: "Loading...", risalaNo: "" });
  const [isConfigLoaded, setIsConfigLoaded] = useState(false);
  const [dashboardData, setDashboardData] = useState({ totalRows: 0 });
  const [rows, setRows] = useState([]);
  const [loading, setLoading] = useState(true);
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
          risalaName: data.risalaName || "N/A",
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
        setConfig({ officeStatus: "ON", risalaName: "Weekly Risala Report", offMessage: "", risalaNo: "" });
        setIsConfigLoaded(true);
        fetchData();
      }
    } catch(e) {
      setConfig({ officeStatus: "ON", risalaName: "Weekly Risala Report", offMessage: "", risalaNo: "" });
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
              else if (nKey.includes('chain')) newRow.chain = row[key];
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
        (!search || JSON.stringify(x).toLowerCase().includes(search.toLowerCase().trim()))
    );
    return result;
  }, [rows, search, region, state, division, district, department, chain]);

  useEffect(() => { setCurrentPage(1); }, [search, region, state, division, district, department, chain]);

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

    const regionDataPPT = groupCount(filteredRows, "region");
    if (regionDataPPT.length > 0) {
        let slide3 = pres.addSlide();
        slide3.background = { color: "ffffff" };
        slide3.addText("Reports by Region", { x: 0.5, y: 0.5, w: 9, fontSize: 24, bold: true, color: "0f766e" });
        let chartData = [{ name: "Reports", labels: regionDataPPT.slice(0, 8).map(d => d.label || "Unknown"), values: regionDataPPT.slice(0, 8).map(d => d.count) }];
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

  const regionData = groupCount(filteredRows, "region");
  const maxRegionCount = regionData.length ? Math.max(...regionData.map(d => d.count)) : 0;

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
        <header className="bg-white border-b border-slate-200 sticky top-0 z-30 shadow-sm">
          <div className="max-w-[1400px] mx-auto px-4 sm:px-6 lg:px-8 flex flex-col md:flex-row items-center justify-between py-4 gap-4">
            <div className="flex items-center gap-4 w-full md:w-auto">
              <button onClick={onBack} className="p-2.5 bg-slate-100 hover:bg-slate-200 border border-slate-200 text-slate-700 rounded-lg shadow-sm transition-colors">
                <ArrowLeft className="w-5 h-5" />
              </button>
              <div className="p-2.5 bg-teal-700 text-white rounded-lg shadow-sm border border-teal-700 hidden sm:block">
                <BookOpenCheck className="w-6 h-6" />
              </div>
              <div>
                <h1 className="text-lg sm:text-xl font-bold text-slate-800 tracking-wide">WEEKLY RISALA <span className="font-light text-teal-700">REPORT</span></h1>
                <p className="text-[10px] font-bold text-teal-700 uppercase tracking-widest mt-0.5">{dashboardData.b4Value}</p>
              </div>
            </div>
            <div className="flex flex-wrap items-center justify-center md:justify-end gap-3 w-full md:w-auto">
              <div className="text-[10px] font-bold text-slate-600 uppercase tracking-widest bg-slate-50 px-3 py-2 rounded-lg border border-slate-200 hidden lg:block">
                User: <span className="text-teal-700">{officeUser?.name || officeUser?.userId || "Admin"}</span>
              </div>
              <button onClick={() => fetchData()} disabled={loading} className="flex items-center gap-2 text-xs font-bold tracking-wider uppercase bg-teal-600 hover:bg-teal-700 text-white px-4 py-2.5 rounded-lg border border-teal-600 transition-all active:scale-95">
                <RefreshCw className={`w-4 h-4 ${loading ? "animate-spin" : ""}`} /> <span className="hidden sm:inline">Sync</span>
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

        <main className="max-w-[1400px] mx-auto px-4 sm:px-6 lg:px-8 py-8 space-y-6">
          
          {fetchError && (
            <div className="p-4 bg-red-50 border border-red-200 text-red-600 rounded-xl text-sm font-medium flex items-center gap-3">
              <div className="w-2 h-2 rounded-full bg-red-500 animate-ping"></div>
              {fetchError}
            </div>
          )}

          <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
            <div className="bg-white p-6 rounded-xl border border-slate-200 shadow-sm relative overflow-hidden group">
              <div className="relative flex justify-between items-center z-10">
                <div className="flex items-center gap-4">
                    <div className="p-3.5 rounded-lg bg-amber-50 text-amber-500 border border-amber-100">
                      <LayoutDashboard className="w-6 h-6" />
                    </div>
                    <div>
                        <p className="text-[11px] font-bold text-slate-500 uppercase tracking-widest mb-1">TOTAL SUBMITTED</p>
                        <h3 className="text-4xl font-extrabold text-teal-700 tracking-tight">{filteredRows.length.toLocaleString("en-IN")}</h3>
                    </div>
                </div>
              </div>
            </div>
            <div className="bg-white p-6 rounded-xl border border-slate-200 shadow-sm relative overflow-hidden group">
              <div className="relative flex justify-between items-center z-10">
                <div className="flex items-center gap-4">
                    <div className="p-3.5 rounded-lg bg-teal-50 text-teal-500 border border-teal-100">
                      <Activity className="w-6 h-6" />
                    </div>
                    <div>
                        <p className="text-[11px] font-bold text-slate-500 uppercase tracking-widest mb-1">REPORT QUANTITY</p>
                        <h3 className="text-4xl font-extrabold text-teal-700 tracking-tight">{totalReportSum.toLocaleString("en-IN")}</h3>
                    </div>
                </div>
              </div>
            </div>
          </div>

          <div className="bg-white p-5 rounded-xl border border-slate-200 shadow-sm">
            <div className="flex items-center gap-3 mb-6">
              <Filter className="w-5 h-5 text-teal-700" />
              <h2 className="text-sm font-bold text-slate-800 tracking-widest uppercase">Data Parameters</h2>
            </div>
            <div className="grid grid-cols-2 md:grid-cols-7 gap-4">
              <div className="col-span-2 md:col-span-1">
                <label className="block text-[10px] font-bold text-slate-500 uppercase tracking-widest mb-2">Search</label>
                <input type="text" placeholder="Search..." value={search} onChange={(e) => setSearch(e.target.value)} className="w-full p-3 rounded-lg border border-slate-300 bg-white text-sm font-medium text-slate-700 focus:outline-none focus:border-teal-500 focus:ring-1 focus:ring-teal-500" />
              </div>
              {[
                { label: "Region", val: region, set: handleSetRegion, opts: uniqValues(rows, "region") },
                { label: "State", val: state, set: handleSetState, opts: uniqValues(rows.filter(x=>!region||sameClient(x.region,region)), "state") },
                { label: "Division", val: division, set: handleSetDivision, opts: uniqValues(rows.filter(x=>(!region||sameClient(x.region,region))&&(!state||sameClient(x.state,state))), "division") },
                { label: "District", val: district, set: setDistrict, opts: uniqValues(rows.filter(x=>(!region||sameClient(x.region,region))&&(!state||sameClient(x.state,state))&&(!division||sameClient(x.division,division))), "district") },
                { label: "Department", val: department, set: setDepartment, opts: uniqValues(rows, "department") },
                { label: "Chain", val: chain, set: setChain, opts: uniqValues(rows, "chain") }
              ].map((f, i) => (
                <div key={i}>
                  <label className="block text-[10px] font-bold text-slate-500 uppercase tracking-widest mb-2">{f.label}</label>
                  <select value={f.val} onChange={(e) => f.set(e.target.value)} disabled={officeUser?.[f.label.toLowerCase()] && officeUser[f.label.toLowerCase()].toLowerCase() !== "all"} className="w-full p-3 rounded-lg border border-slate-300 bg-white text-sm font-medium text-slate-700 focus:outline-none focus:border-teal-500 focus:ring-1 focus:ring-teal-500 appearance-none">
                    <option value="">{f.label}</option>
                    {f.opts.map(o => <option key={o} value={o}>{o}</option>)}
                  </select>
                </div>
              ))}
            </div>
          </div>

          <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
            <div className="bg-white rounded-xl shadow-sm border border-slate-200 p-5 h-[320px] flex flex-col">
               <h3 className="text-xs font-bold text-teal-700 uppercase tracking-widest mb-4">REPORTS BY REGION</h3>
               <div className="flex-1 flex items-end gap-2 pb-2">
                 {regionData.length > 0 ? regionData.slice(0, 8).map((d, i) => {
                    const h = maxRegionCount ? (d.count / maxRegionCount) * 100 : 0;
                    return (
                      <div key={i} className="flex-1 flex flex-col justify-end items-center group relative h-full">
                         <div className="opacity-0 group-hover:opacity-100 absolute -top-8 bg-white border border-teal-200 text-teal-700 text-[10px] px-2 py-1 rounded shadow-lg transition-opacity whitespace-nowrap z-10 pointer-events-none">
                           {d.count.toLocaleString("en-IN")}
                         </div>
                         <div style={{height: `${Math.max(h, 2)}%`}} className="w-full bg-[#21496b] group-hover:bg-[#1e40af] rounded-t-sm transition-all" />
                         <span className="text-[9px] text-slate-600 mt-2 truncate w-full text-center px-1" title={d.label}>{d.label}</span>
                      </div>
                    )
                 }) : <div className="w-full text-center text-slate-500 text-xs my-auto">No data</div>}
               </div>
            </div>
            <MiniTable title="REPORTS BY STATE" data={groupCount(filteredRows, "state")} />
            <MiniTable title="REPORTS BY DEPARTMENT" data={groupCount(filteredRows, "department")} />
            <MiniTable title="REPORTS BY DIVISION" data={groupCount(filteredRows, "division")} />
          </div>

          <div className="bg-white rounded-xl border border-slate-200 overflow-hidden shadow-sm mt-6">
            <div className="p-4 border-b border-slate-200 flex flex-col md:flex-row justify-between items-center gap-4 bg-[#e0f2f1]">
              <div>
                <h2 className="text-sm font-bold text-teal-800 uppercase tracking-widest">Detailed Telemetry Output</h2>
                <p className="text-[11px] text-slate-600 mt-1 uppercase tracking-widest">Source: {dashboardData.totalRows} • Visible: {filteredRows.length}</p>
              </div>
            </div>
            <div className="overflow-x-auto">
              <table className="min-w-full text-left text-sm whitespace-nowrap">
                <thead className="bg-[#008b8b]">
                  <tr>
                    {["Date", "Name", "Contact", "Chain", "Level", "Department", "Report", "District", "Division", "State", "Region"].map(h => (
                      <th key={h} className={`px-4 py-4 text-[10px] font-bold text-white uppercase tracking-widest ${h==="Report"?"text-white text-right":""}`}>{h}</th>
                    ))}
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {pagedRows.length > 0 ? pagedRows.map((row, idx) => (
                    <tr key={idx} className="hover:bg-slate-50 transition-colors">
                      <td className="px-4 py-3 text-xs text-slate-600">{row.date}</td>
                      <td className="px-4 py-3 text-xs font-bold text-slate-800">{row.name}</td>
                      <td className="px-4 py-3 text-xs text-slate-600">{row.contact}</td>
                      <td className="px-4 py-3 text-[10px] uppercase tracking-wider font-bold text-slate-600"><span className="bg-slate-100 px-2 py-1 rounded">{row.chain}</span></td>
                      <td className="px-4 py-3 text-xs text-slate-600">{row.nigran || row.zimmedar || row.level}</td>
                      <td className="px-4 py-3 text-xs text-slate-700">{row.department}</td>
                      <td className="px-4 py-3 text-sm text-right font-extrabold text-teal-700">{row.report}</td>
                      <td className="px-4 py-3 text-xs text-slate-700">{row.district}</td>
                      <td className="px-4 py-3 text-xs text-slate-600">{row.division}</td>
                      <td className="px-4 py-3 text-xs text-slate-600">{row.state}</td>
                      <td className="px-4 py-3 text-xs text-slate-600">{row.region}</td>
                    </tr>
                  )) : (
                    <tr><td colSpan="11" className="px-6 py-12 text-center text-slate-500 text-xs uppercase tracking-widest">{loading ? "Fetching records..." : "No matching records found"}</td></tr>
                  )}
                </tbody>
              </table>
            </div>
            {filteredRows.length > rowsPerPage && (
              <div className="p-4 border-t border-slate-200 bg-[#f8fafc] flex justify-center items-center gap-2">
                <button disabled={currentPage===1} onClick={()=>setCurrentPage(p=>p-1)} className="px-3 py-1.5 rounded-lg border border-teal-600 bg-white text-xs font-bold text-teal-700 disabled:opacity-30">PREV</button>
                <span className="text-xs font-bold text-teal-700 px-4">PAGE {currentPage} OF {Math.ceil(filteredRows.length / rowsPerPage)}</span>
                <button disabled={currentPage===Math.ceil(filteredRows.length / rowsPerPage)} onClick={()=>setCurrentPage(p=>p+1)} className="px-3 py-1.5 rounded-lg border border-teal-600 bg-white text-xs font-bold text-teal-700 disabled:opacity-30">NEXT</button>
              </div>
            )}
          </div>
        </main>
      </div>
    </div>
  );
}
