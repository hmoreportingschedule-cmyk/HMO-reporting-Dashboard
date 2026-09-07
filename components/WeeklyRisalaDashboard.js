"use client";
import React, { useState, useEffect, useMemo } from "react";
import Papa from "papaparse";
import { ArrowLeft, BookOpenCheck, LogOut, RefreshCw, Filter, Search, Activity, Download, LayoutDashboard, Database, Presentation, FileText, Clock } from "lucide-react";
import pptxgen from "pptxgenjs";
import * as XLSX from "xlsx";

const SCRIPT_URL = "https://script.google.com/macros/s/AKfycbyF74lC0dNiWUalx0G7GEK3F802IMBMXfuCsqfuCi5-QYuGkOh-85R_BzK9U_O9mfkpUA/exec";
const DEFAULT_CSV_URL = "https://docs.google.com/spreadsheets/d/1GfMa7j1TIx17jG0g25tdEwU2YgHCm9_fbcrWIUTrSeI/gviz/tq?tqx=out:csv&sheet=Responses";

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
        // Fallback if script not updated yet
        setConfig({ officeStatus: "ON", risalaName: "Weekly Risala Report", offMessage: "", risalaNo: "" });
        setIsConfigLoaded(true);
        fetchData();
      }
    } catch(e) {
      // Fallback on network error for config
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
      error: (err) => {
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
    <div className="min-h-screen bg-[#e0f2f1] text-slate-800 font-sans p-4 sm:p-6 lg:p-8">
       <div className="max-w-[1400px] mx-auto space-y-6">
          
          {/* Top Header */}
          <div className="bg-white rounded-xl shadow-sm p-5 border-t-[6px] border-teal-600 flex flex-col gap-4">
             <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-4">
                <div>
                   <div className="bg-teal-700 text-white text-lg md:text-xl font-bold px-5 py-2 rounded-lg inline-flex items-center gap-2 mb-3 shadow-sm">
                      <BookOpenCheck className="w-5 h-5" />
                      Weekly Risala Report - Dashboard
                   </div>
                   <div className="text-teal-700 font-bold text-lg mb-1">Risala Name: {config.risalaName}</div>
                   <div className="text-slate-600 font-semibold text-sm">Access: {officeUser?.country || 'India'} / {officeUser?.chain || 'All'}, {officeUser?.department || 'All'}</div>
                </div>
                <div className="flex flex-col items-start md:items-end gap-3 w-full md:w-auto">
                   <div className="flex flex-wrap items-center gap-3">
                      {config.risalaNo && <span className="border border-teal-600 text-teal-700 font-bold px-3 py-1.5 rounded-lg text-sm bg-teal-50">{config.risalaNo}</span>}
                      <span className="text-teal-700 font-bold text-sm">Date: {new Date().toLocaleDateString('en-GB')} | Time: {new Date().toLocaleTimeString('en-GB')}</span>
                      <button onClick={onLogout} className="bg-[#dc3545] hover:bg-red-700 text-white font-bold py-1.5 px-4 rounded transition-colors text-sm shadow-sm flex items-center gap-1">
                         Logout
                      </button>
                   </div>
                   <div className="flex flex-wrap items-center gap-3 mt-1">
                      <span className="text-slate-800 font-bold text-sm bg-slate-100 px-3 py-1.5 rounded border border-slate-200">
                         User : {officeUser?.name || officeUser?.userId || "Admin"}
                      </span>
                      <button onClick={loadConfigAndData} disabled={loading} className="bg-teal-600 hover:bg-teal-700 text-white font-bold py-1.5 px-4 rounded transition-colors text-sm flex items-center gap-1 shadow-sm">
                         <RefreshCw className={`w-4 h-4 ${loading ? 'animate-spin' : ''}`} /> Refresh
                      </button>
                      <button onClick={downloadPPT} className="bg-teal-600 hover:bg-teal-700 text-white font-bold py-1.5 px-4 rounded transition-colors text-sm flex items-center gap-1 shadow-sm">
                         <Presentation className="w-4 h-4" /> PPT
                      </button>
                      <button onClick={downloadExcel} className="bg-[#198754] hover:bg-green-700 text-white font-bold py-1.5 px-4 rounded transition-colors text-sm flex items-center gap-1 shadow-sm">
                         <Download className="w-4 h-4" /> Download Excel
                      </button>
                   </div>
                </div>
             </div>
          </div>

          {fetchError && <div className="p-4 bg-red-50 border border-red-200 text-red-600 rounded-xl text-sm font-bold flex items-center gap-3 shadow-sm"><div className="w-2 h-2 rounded-full bg-red-500 animate-ping"></div>{fetchError}</div>}

          {/* KPIs */}
          <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
             <div className="bg-white rounded-xl shadow-sm p-6 border-l-[6px] border-teal-600 flex justify-between items-center transition-all hover:shadow-md">
                <div className="flex items-center gap-3">
                   <FileText className="text-amber-500 w-6 h-6" />
                   <span className="text-slate-700 font-bold text-xs uppercase tracking-wider">Total Submitted</span>
                </div>
                <span className="text-teal-700 font-extrabold text-4xl">{filteredRows.length.toLocaleString("en-IN")}</span>
             </div>
             <div className="bg-white rounded-xl shadow-sm p-6 border-l-[6px] border-teal-600 flex justify-between items-center transition-all hover:shadow-md">
                <div className="flex items-center gap-3">
                   <Clock className="text-teal-500 w-6 h-6" />
                   <span className="text-slate-700 font-bold text-xs uppercase tracking-wider">Report</span>
                </div>
                <span className="text-teal-700 font-extrabold text-4xl">{totalReportSum.toLocaleString("en-IN")}</span>
             </div>
          </div>

          {/* Filters */}
          <div className="bg-white rounded-xl shadow-sm p-5 flex flex-wrap items-end gap-4 border border-slate-200">
             <div className="flex flex-col flex-1 min-w-[150px]">
                <input type="text" placeholder="Search..." value={search} onChange={e=>setSearch(e.target.value)} className="border border-slate-300 rounded-lg px-4 py-2.5 text-sm outline-none focus:border-teal-500 focus:ring-1 focus:ring-teal-500 text-slate-700 placeholder-slate-400 w-full transition-all" />
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
                 <select value={f.val} onChange={(e) => f.set(e.target.value)} disabled={officeUser?.[f.label.toLowerCase()] && officeUser[f.label.toLowerCase()].toLowerCase() !== "all"} className="border border-slate-300 rounded-lg px-3 py-2.5 text-sm outline-none focus:border-teal-500 focus:ring-1 focus:ring-teal-500 text-slate-700 bg-white disabled:bg-slate-100 disabled:text-slate-400 transition-all cursor-pointer">
                   <option value="">{f.label}</option>
                   {f.opts.map(o => <option key={o} value={o}>{o}</option>)}
                 </select>
               </div>
             ))}
          </div>

          {/* Charts Grid */}
          <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
             <div className="bg-white rounded-xl shadow-sm border border-slate-200 p-5 flex flex-col h-[320px]">
                <h3 className="text-teal-700 font-bold uppercase tracking-widest text-sm mb-6 border-b border-slate-100 pb-3">REPORTS BY REGION</h3>
                <div className="flex-1 flex items-end gap-3 pb-2">
                  {regionData.length > 0 ? regionData.slice(0, 8).map((d, i) => {
                     const h = maxRegionCount ? (d.count / maxRegionCount) * 100 : 0;
                     const colors = ["bg-[#1e40af]", "bg-[#2563eb]", "bg-[#14b8a6]", "bg-[#f59e0b]", "bg-[#ef4444]", "bg-[#f97316]"];
                     return (
                       <div key={i} className="flex-1 flex flex-col justify-end items-center group relative h-full">
                          <div className="absolute -top-6 text-slate-800 text-[11px] font-bold">{d.count.toLocaleString("en-IN")}</div>
                          <div style={{height: `${Math.max(h, 2)}%`}} className={`w-full ${colors[i%colors.length] || 'bg-teal-600'} rounded-t-sm transition-all hover:opacity-80`} />
                          <span className="text-[10px] text-slate-600 mt-2 truncate w-full text-center px-1 font-medium">{d.label}</span>
                       </div>
                     )
                  }) : <div className="w-full text-center text-slate-400 text-xs my-auto font-medium">No data available to display</div>}
                </div>
             </div>
             <MiniTable title="REPORTS BY STATE" data={groupCount(filteredRows, "state")} />
             <MiniTable title="REPORTS BY DEPARTMENT" data={groupCount(filteredRows, "department")} />
             <MiniTable title="REPORTS BY DIVISION" data={groupCount(filteredRows, "division")} />
          </div>

          {/* Main Table */}
          <div className="bg-white rounded-xl shadow-sm border border-slate-200 overflow-hidden mt-6">
             <div className="bg-[#e0f2f1] border-b border-slate-200 p-4 flex flex-col md:flex-row justify-between items-center gap-4">
                <span className="text-teal-800 text-xs font-bold uppercase tracking-wider">
                   Responses rows: {dashboardData.totalRows} &bull; Visible rows: {filteredRows.length} &bull; Report quantity: {totalReportSum.toLocaleString('en-IN')}
                </span>
             </div>
             <div className="overflow-x-auto custom-scrollbar">
                <table className="min-w-full text-left text-sm whitespace-nowrap">
                   <thead className="bg-[#f8fafc] border-b border-slate-200 text-teal-700">
                      <tr>
                         {["Date", "Name", "Contact", "Chain", "Level", "Department", "Report", "District", "Division", "State", "Region"].map(h => (
                           <th key={h} className={`px-5 py-4 text-xs font-bold uppercase tracking-wider ${h==="Report"?"text-right":""}`}>{h}</th>
                         ))}
                      </tr>
                   </thead>
                   <tbody className="divide-y divide-slate-100">
                      {pagedRows.length > 0 ? pagedRows.map((row, idx) => (
                         <tr key={idx} className="hover:bg-slate-50 transition-colors">
                           <td className="px-5 py-3 text-slate-600">{row.date}</td>
                           <td className="px-5 py-3 font-semibold text-slate-800">{row.name}</td>
                           <td className="px-5 py-3 text-slate-600">{row.contact}</td>
                           <td className="px-5 py-3 text-xs uppercase tracking-wider font-bold text-slate-500">{row.chain}</td>
                           <td className="px-5 py-3 text-slate-600">{row.nigran || row.zimmedar || row.level}</td>
                           <td className="px-5 py-3 text-slate-700">{row.department}</td>
                           <td className="px-5 py-3 text-right font-extrabold text-teal-700 text-base">{row.report}</td>
                           <td className="px-5 py-3 text-slate-600">{row.district}</td>
                           <td className="px-5 py-3 text-slate-600">{row.division}</td>
                           <td className="px-5 py-3 text-slate-600">{row.state}</td>
                           <td className="px-5 py-3 text-slate-600">{row.region}</td>
                         </tr>
                      )) : (
                         <tr><td colSpan="11" className="px-6 py-12 text-center text-slate-500 text-sm font-medium">{loading ? "Fetching records..." : "No matching records found"}</td></tr>
                      )}
                   </tbody>
                </table>
             </div>
             {filteredRows.length > rowsPerPage && (
                <div className="p-4 border-t border-slate-200 bg-[#f8fafc] flex justify-center items-center gap-3">
                   <button disabled={currentPage===1} onClick={()=>setCurrentPage(p=>p-1)} className="px-4 py-2 rounded-lg border border-teal-600 text-teal-700 text-xs font-bold hover:bg-teal-50 transition-colors disabled:opacity-40 disabled:hover:bg-transparent">PREV</button>
                   <span className="text-xs font-bold text-slate-600 px-2">PAGE {currentPage} OF {Math.ceil(filteredRows.length / rowsPerPage)}</span>
                   <button disabled={currentPage===Math.ceil(filteredRows.length / rowsPerPage)} onClick={()=>setCurrentPage(p=>p+1)} className="px-4 py-2 rounded-lg border border-teal-600 text-teal-700 text-xs font-bold hover:bg-teal-50 transition-colors disabled:opacity-40 disabled:hover:bg-transparent">NEXT</button>
                </div>
             )}
          </div>

        </main>
      </div>
    </div>
  );
}
