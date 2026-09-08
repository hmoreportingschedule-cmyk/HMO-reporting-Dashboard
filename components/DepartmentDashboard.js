"use client";
import React, { useState, useEffect, useMemo } from "react";
import Papa from "papaparse";
import { ArrowLeft, Building2, LogOut, RefreshCw, Download, Presentation, Filter, Search, Image as ImageIcon, Target, TrendingUp, Percent, AlertCircle, Calendar } from "lucide-react";
import pptxgen from "pptxgenjs";
import * as XLSX from "xlsx";
import html2canvas from "html2canvas";

// ⚠️ YAHAN DEPARTMENT DASHBOARD KI GOOGLE SHEET KA CSV LINK PASTE KAREIN
const DEFAULT_CSV_URL = "PASTE_DEPARTMENT_CSV_LINK_HERE";

export default function DepartmentDashboard({ onBack, officeUser, onLogout }) {
  const [sheetUrl, setSheetUrl] = useState(DEFAULT_CSV_URL);
  const [rawData, setRawData] = useState([]);
  const [loading, setLoading] = useState(false);
  const [fetchError, setFetchError] = useState("");

  // Filters
  const [searchQuery, setSearchQuery] = useState("");
  const [selectedDept, setSelectedDept] = useState(officeUser?.department && officeUser.department.toLowerCase() !== "all" ? officeUser.department : "All");
  const [startDate, setStartDate] = useState("");
  const [endDate, setEndDate] = useState("");

  // Pagination
  const [currentPage, setCurrentPage] = useState(1);
  const rowsPerPage = 10;

  useEffect(() => {
    if (DEFAULT_CSV_URL !== "PASTE_DEPARTMENT_CSV_LINK_HERE") {
      fetchData();
    }
  }, []);

  const fetchData = async () => {
    if (!sheetUrl || sheetUrl === "PASTE_DEPARTMENT_CSV_LINK_HERE") {
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
          // Process numeric values for Target and Achievement
          const processed = results.data.map(row => ({
            ...row,
            Target: Number(String(row["Target"] || "0").replace(/,/g, "")),
            Achievement: Number(String(row["Achievement"] || "0").replace(/,/g, "")),
            // Assume format YYYY-MM-DD or parseable date string in 'Date' column
            ParsedDate: new Date(row["Date"]) 
          })).filter(r => !isNaN(r.Target) || !isNaN(r.Achievement)); // Keep rows with data
          setRawData(processed);
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

  // Setup YTD (Year To Date)
  const setYTD = () => {
    const today = new Date();
    const startOfYear = new Date(today.getFullYear(), 0, 1); // Jan 1st of current year
    
    // Format YYYY-MM-DD for input fields
    const formatDate = (date) => {
        const d = new Date(date);
        let month = '' + (d.getMonth() + 1);
        let day = '' + d.getDate();
        const year = d.getFullYear();

        if (month.length < 2) month = '0' + month;
        if (day.length < 2) day = '0' + day;

        return [year, month, day].join('-');
    };

    setStartDate(formatDate(startOfYear));
    setEndDate(formatDate(today));
  };

  // Unique Departments for Dropdown
  const departmentsList = useMemo(() => {
    const depts = new Set(rawData.map(r => (r["Department"] || "Unknown").trim()));
    return Array.from(depts).sort();
  }, [rawData]);

  // Filtering Logic
  const filteredData = useMemo(() => {
    return rawData.filter(row => {
      // Search
      const searchMatch = !searchQuery || 
        JSON.stringify(row).toLowerCase().includes(searchQuery.toLowerCase().trim());
      
      // Department
      const deptMatch = selectedDept === "All" || (row["Department"] || "").trim() === selectedDept;

      // Date Range
      let dateMatch = true;
      if (startDate || endDate) {
        const rowDate = row.ParsedDate;
        if (!isNaN(rowDate)) {
            const start = startDate ? new Date(startDate) : new Date("1900-01-01");
            const end = endDate ? new Date(endDate) : new Date("2100-01-01");
            // Set end date to end of the day for inclusive filtering
            end.setHours(23, 59, 59, 999);
            dateMatch = rowDate >= start && rowDate <= end;
        }
      }

      return searchMatch && deptMatch && dateMatch;
    });
  }, [rawData, searchQuery, selectedDept, startDate, endDate]);

  useEffect(() => { setCurrentPage(1); }, [searchQuery, selectedDept, startDate, endDate]);

  const pagedRows = useMemo(() => {
    const start = (currentPage - 1) * rowsPerPage;
    return filteredData.slice(start, start + rowsPerPage);
  }, [filteredData, currentPage]);

  // Calculations for Comparison (Target vs Achievement)
  const kpiStats = useMemo(() => {
    let target = 0;
    let achievement = 0;
    filteredData.forEach(r => {
        target += r.Target || 0;
        achievement += r.Achievement || 0;
    });
    const percentage = target > 0 ? ((achievement / target) * 100).toFixed(1) : 0;
    const shortfall = target > achievement ? target - achievement : 0;

    return { target, achievement, percentage, shortfall, totalRecords: filteredData.length };
  }, [filteredData]);

  // Downloads
  const downloadPPT = () => {
    if (!filteredData.length) return alert("No data to export");
    let pres = new pptxgen();
    
    let slide1 = pres.addSlide();
    slide1.background = { color: "ffffff" };
    slide1.addText("DEPARTMENT REPORT CARD", { x: 1, y: 2, w: 8, fontSize: 36, bold: true, color: "008080", align: "center" });
    slide1.addText(`User: ${officeUser?.name || officeUser?.userId || "Admin"}`, { x: 1, y: 3, w: 8, fontSize: 16, color: "334155", align: "center" });
    slide1.addText(`Generated on: ${new Date().toLocaleDateString()}`, { x: 1, y: 3.5, w: 8, fontSize: 12, color: "64748b", align: "center" });

    let slide2 = pres.addSlide();
    slide2.background = { color: "ffffff" };
    slide2.addText("Target vs Achievement", { x: 0.5, y: 0.5, w: 9, fontSize: 24, bold: true, color: "0f766e" });
    
    slide2.addShape(pres.ShapeType.rect, { x: 0.5, y: 1.5, w: 4, h: 1.5, fill: "f8fafc", line: {color: "008080", width: 1} });
    slide2.addText("TOTAL TARGET", { x: 0.5, y: 1.7, w: 4, fontSize: 12, color: "64748b", align: "center" });
    slide2.addText(kpiStats.target.toLocaleString("en-IN"), { x: 0.5, y: 2.2, w: 4, fontSize: 28, bold: true, color: "0f766e", align: "center" });

    slide2.addShape(pres.ShapeType.rect, { x: 5.5, y: 1.5, w: 4, h: 1.5, fill: "f8fafc", line: {color: "008080", width: 1} });
    slide2.addText("ACHIEVEMENT", { x: 5.5, y: 1.7, w: 4, fontSize: 12, color: "64748b", align: "center" });
    slide2.addText(kpiStats.achievement.toLocaleString("en-IN"), { x: 5.5, y: 2.2, w: 4, fontSize: 28, bold: true, color: "0f766e", align: "center" });

    slide2.addShape(pres.ShapeType.rect, { x: 0.5, y: 3.5, w: 4, h: 1.5, fill: "f8fafc", line: {color: "008080", width: 1} });
    slide2.addText("ACHIEVEMENT %", { x: 0.5, y: 3.7, w: 4, fontSize: 12, color: "64748b", align: "center" });
    slide2.addText(kpiStats.percentage + "%", { x: 0.5, y: 4.2, w: 4, fontSize: 28, bold: true, color: "0f766e", align: "center" });

    pres.writeFile({ fileName: `Department_Report_PPT_${new Date().getTime()}.pptx` });
  };

  const downloadExcel = () => {
    if (!filteredData.length) return alert("No data to export");
    const exportData = filteredData.map(x => {
        let rowCopy = { ...x };
        delete rowCopy.ParsedDate; // remove internal parsed object before export
        return rowCopy;
    });
    const worksheet = XLSX.utils.json_to_sheet(exportData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Department Data");
    XLSX.writeFile(workbook, `Department_RawData_${new Date().getTime()}.xlsx`);
  };

  const downloadJPEG = async () => {
    const element = document.getElementById("department-report-card-view");
    if (!element) return;
    try {
        const canvas = await html2canvas(element, { 
            scale: 2, 
            backgroundColor: "#e0f2f1", 
            useCORS: true,
            ignoreElements: (node) => node.hasAttribute("data-html2canvas-ignore")
        });
        const link = document.createElement('a');
        link.download = `Department_Report_Card_${new Date().getTime()}.jpg`;
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
        {/* REPORT CARD VIEW CONTAINER */}
        <div id="department-report-card-view" className="bg-[#e0f2f1] pb-6 pt-4">
            <header className="max-w-[1400px] mx-auto px-4 sm:px-6 lg:px-8 relative z-30">
              <div className="bg-white border border-slate-200 shadow-sm rounded-2xl px-6 py-5 flex flex-col md:flex-row items-center justify-between gap-6">
                <div className="flex items-center gap-4 w-full md:w-auto">
                  <button data-html2canvas-ignore="true" onClick={onBack} className="p-3 bg-slate-100 hover:bg-slate-200 border border-slate-200 text-slate-700 rounded-xl shadow-sm transition-colors">
                    <ArrowLeft className="w-5 h-5" />
                  </button>
                  <div className="p-3 bg-teal-700 text-white rounded-xl shadow-sm border border-teal-700 hidden sm:block">
                    <Building2 className="w-6 h-6" />
                  </div>
                  <div>
                    <h1 className="text-xl sm:text-2xl font-black text-slate-800 tracking-wide">DEPARTMENT <span className="font-light text-teal-700">REPORT</span></h1>
                    <p className="text-[10px] font-bold text-teal-700 uppercase tracking-widest mt-0.5">Target & Achievement Analytics</p>
                  </div>
                </div>
                
                <div data-html2canvas-ignore="true" className="flex flex-wrap items-center justify-center md:justify-end gap-3 w-full md:w-auto">
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

              {/* Filters & YTD Box */}
              <div className="bg-white p-5 rounded-xl border border-slate-200 shadow-sm">
                <div className="flex flex-col lg:flex-row justify-between lg:items-center gap-4 mb-5 border-b border-slate-100 pb-4">
                  <div className="flex items-center gap-3">
                    <Filter className="w-5 h-5 text-teal-700" />
                    <h2 className="text-sm font-bold text-slate-800 tracking-widest uppercase">Filter & Compare</h2>
                  </div>
                  {/* Date Range Selection */}
                  <div className="flex flex-wrap items-center gap-3">
                     <button onClick={setYTD} className="px-3 py-2 text-[10px] font-bold tracking-widest uppercase bg-teal-50 text-teal-700 rounded-lg border border-teal-200 hover:bg-teal-100 transition-colors">YTD (Year to Date)</button>
                     <div className="flex items-center gap-2 bg-slate-50 p-1.5 rounded-lg border border-slate-200">
                        <Calendar className="w-4 h-4 text-slate-400 ml-2" />
                        <input type="date" value={startDate} onChange={e=>setStartDate(e.target.value)} className="bg-transparent text-xs font-bold text-slate-700 outline-none" />
                        <span className="text-slate-400 text-xs font-bold">TO</span>
                        <input type="date" value={endDate} onChange={e=>setEndDate(e.target.value)} className="bg-transparent text-xs font-bold text-slate-700 outline-none pr-2" />
                     </div>
                  </div>
                </div>

                <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                  <div data-html2canvas-ignore="true">
                    <label className="block text-[10px] font-bold text-slate-500 uppercase tracking-widest mb-2">Search Records</label>
                    <div className="relative">
                      <span className="absolute inset-y-0 left-0 pl-3 flex items-center text-slate-400"><Search className="w-4 h-4" /></span>
                      <input type="text" placeholder="Search by any field..." value={searchQuery} onChange={(e) => setSearchQuery(e.target.value)} className="w-full pl-9 pr-3 py-2.5 rounded-lg border border-slate-300 bg-white text-sm font-medium text-slate-700 focus:outline-none focus:border-teal-500 focus:ring-1 focus:ring-teal-500" />
                    </div>
                  </div>
                  <div>
                      <label className="block text-[10px] font-bold text-slate-500 uppercase tracking-widest mb-2">Department</label>
                      <select value={selectedDept} onChange={(e) => setSelectedDept(e.target.value)} disabled={officeUser?.department && officeUser.department.toLowerCase() !== "all"} className="w-full p-2.5 rounded-lg border border-slate-300 bg-white text-sm font-medium text-slate-700 focus:outline-none focus:border-teal-500 focus:ring-1 focus:ring-teal-500 appearance-none cursor-pointer disabled:bg-slate-50 disabled:text-slate-400">
                        <option value="All">All Departments</option>
                        {departmentsList.map(o => <option key={o} value={o}>{o}</option>)}
                      </select>
                  </div>
                </div>
              </div>

              {/* KPIs: Target & Achievement */}
              <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-6">
                {[
                  { title: "Total Target", val: kpiStats.target, icon: Target, color: "text-blue-500", bg: "bg-blue-50", border: "border-blue-100" },
                  { title: "Achievement", val: kpiStats.achievement, icon: TrendingUp, color: "text-emerald-500", bg: "bg-emerald-50", border: "border-emerald-100" },
                  { title: "Achievement %", val: kpiStats.percentage + "%", icon: Percent, color: "text-purple-500", bg: "bg-purple-50", border: "border-purple-100" },
                  { title: "Shortfall", val: kpiStats.shortfall, icon: AlertCircle, color: "text-orange-500", bg: "bg-orange-50", border: "border-orange-100" }
                ].map((kpi, idx) => (
                  <div key={idx} className="bg-white p-6 rounded-xl border border-slate-200 shadow-sm relative overflow-hidden group">
                    <div className="relative flex justify-between items-center z-10">
                      <div className="flex items-center gap-4">
                          <div className={`p-3.5 rounded-xl ${kpi.bg} ${kpi.color} border ${kpi.border}`}>
                            <kpi.icon className="w-6 h-6" />
                          </div>
                          <div>
                              <p className="text-[11px] font-bold text-slate-500 uppercase tracking-widest mb-1.5">{kpi.title}</p>
                              <h3 className={`text-3xl font-black ${idx === 2 ? 'text-purple-700' : 'text-slate-800'} tracking-tight`}>
                                  {typeof kpi.val === 'number' ? kpi.val.toLocaleString("en-IN") : kpi.val}
                              </h3>
                          </div>
                      </div>
                    </div>
                  </div>
                ))}
              </div>

              {/* Comparison Visual */}
              <div className="bg-white p-6 rounded-xl border border-slate-200 shadow-sm">
                <h3 className="text-xs font-bold text-teal-700 uppercase tracking-widest mb-6">Target vs Achievement Progress</h3>
                <div className="relative w-full h-8 bg-slate-100 rounded-full overflow-hidden border border-slate-200">
                    <div 
                        className={`absolute top-0 left-0 h-full flex items-center justify-end pr-3 font-bold text-white text-[10px] transition-all duration-1000 ${kpiStats.percentage >= 100 ? 'bg-emerald-500' : 'bg-teal-600'}`}
                        style={{ width: `${Math.min(kpiStats.percentage, 100)}%` }}>
                        {kpiStats.percentage}%
                    </div>
                </div>
                <div className="flex justify-between items-center mt-3 text-[10px] font-bold text-slate-500 uppercase tracking-widest">
                    <span>0</span>
                    <span>Target: {kpiStats.target.toLocaleString('en-IN')}</span>
                </div>
              </div>

            </main>
        </div> {/* END OF REPORT CARD VIEW */}

        <main className="max-w-[1400px] mx-auto px-4 sm:px-6 lg:px-8 pb-8 mt-2">
          {/* Table Section - Excluded from JPEG */}
          <div className="bg-white rounded-xl border border-slate-200 shadow-sm mt-6">
            <div className="p-4 border-b border-slate-200 flex flex-col md:flex-row justify-between items-center gap-4 bg-[#e0f2f1] rounded-t-xl">
              <div>
                <h2 className="text-sm font-bold text-teal-800 uppercase tracking-widest">Detailed Department Records</h2>
                <p className="text-[10px] text-slate-600 mt-1 uppercase tracking-widest">Source: {rawData.length} &bull; Visible: {filteredData.length}</p>
              </div>
            </div>

            <div className="overflow-x-auto w-full custom-scrollbar">
              <table className="min-w-full text-left text-[11px] lg:text-xs">
                <thead className="bg-[#008b8b]">
                  <tr>
                    {["Date", "Department", "Location / Details", "Target", "Achievement"].map(h => (
                      <th key={h} className={`px-4 py-3 font-bold text-white uppercase tracking-wider ${h==="Target"||h==="Achievement"?"text-right":""}`}>{h}</th>
                    ))}
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {pagedRows.length > 0 ? pagedRows.map((row, idx) => (
                    <tr key={idx} className="hover:bg-slate-50 transition-colors">
                      <td className="px-4 py-3 text-slate-600 whitespace-nowrap font-medium">{row["Date"] || "-"}</td>
                      <td className="px-4 py-3 text-slate-800 font-bold leading-tight uppercase">{row["Department"] || "-"}</td>
                      <td className="px-4 py-3 text-slate-600 leading-tight">
                         {row["Region"] && <span className="font-bold text-teal-700 mr-1">{row["Region"]},</span>} 
                         {row["District"] || row["Details"] || "-"}
                      </td>
                      <td className="px-4 py-3 text-right font-bold text-slate-600 text-sm">{(row.Target || 0).toLocaleString("en-IN")}</td>
                      <td className="px-4 py-3 text-right font-extrabold text-teal-700 text-sm bg-teal-50/40">{(row.Achievement || 0).toLocaleString("en-IN")}</td>
                    </tr>
                  )) : (
                    <tr>
                      <td colSpan="5" className="px-6 py-12 text-center text-slate-500 text-xs uppercase tracking-widest">
                        {loading ? "Fetching records..." : "No matching records found"}
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
