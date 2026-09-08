"use client";
import React, { useState, useEffect, useMemo } from "react";
import Papa from "papaparse";
import { LogOut, RefreshCw, Filter, Layers, Calendar, MapPin, Search, Activity, ArrowLeft, User, Presentation, Download, BookOpenCheck, Image as ImageIcon, LayoutDashboard } from "lucide-react";
import * as XLSX from "xlsx";
import pptxgen from "pptxgenjs";
import html2canvas from "html2canvas";

// ⚠️ YAHAN DEENI KAAM KI GOOGLE SHEET KA CSV LINK PASTE KAREIN
const DEFAULT_SHEET_URL = "PASTE_DEENI_KAAM_CSV_LINK_HERE";

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

  useEffect(() => {
    if(DEFAULT_SHEET_URL !== "PASTE_DEENI_KAAM_CSV_LINK_HERE") {
      fetchData();
    }
  }, []);

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
    slide1.background = { color: "ffffff" };
    slide1.addText("12 DEENI KAAM REPORT", { x: 1, y: 2, w: 8, fontSize: 36, bold: true, color: "008080", align: "center" });
    slide1.addText(`User: ${officeUser?.name || officeUser?.userId || "Admin"}`, { x: 1, y: 3, w: 8, fontSize: 16, color: "334155", align: "center" });
    slide1.addText(`Generated on: ${new Date().toLocaleDateString()}`, { x: 1, y: 3.5, w: 8, fontSize: 12, color: "64748b", align: "center" });

    let slide2 = pres.addSlide();
    slide2.background = { color: "ffffff" };
    slide2.addText("Key Performance Indicators", { x: 0.5, y: 0.5, w: 9, fontSize: 24, bold: true, color: "0f766e" });
    
    slide2.addShape(pres.ShapeType.rect, { x: 0.5, y: 1.5, w: 4, h: 1.5, fill: "f8fafc", line: {color: "008080", width: 1} });
    slide2.addText("TOTAL MUBALLIGH", { x: 0.5, y: 1.7, w: 4, fontSize: 12, color: "64748b", align: "center" });
    slide2.addText(kpiStats.totalMuballigh.toLocaleString("en-IN"), { x: 0.5, y: 2.2, w: 4, fontSize: 28, bold: true, color: "0f766e", align: "center" });

    slide2.addShape(pres.ShapeType.rect, { x: 5.5, y: 1.5, w: 4, h: 1.5, fill: "f8fafc", line: {color: "008080", width: 1} });
    slide2.addText("TOTAL MASAJID", { x: 5.5, y: 1.7, w: 4, fontSize: 12, color: "64748b", align: "center" });
    slide2.addText(kpiStats.totalMasjid.toLocaleString("en-IN"), { x: 5.5, y: 2.2, w: 4, fontSize: 28, bold: true, color: "0f766e", align: "center" });

    slide2.addShape(pres.ShapeType.rect, { x: 0.5, y: 3.5, w: 4, h: 1.5, fill: "f8fafc", line: {color: "008080", width: 1} });
    slide2.addText("ZEILI HALQE", { x: 0.5, y: 3.7, w: 4, fontSize: 12, color: "64748b", align: "center" });
    slide2.addText(kpiStats.totalHalqe.toLocaleString("en-IN"), { x: 0.5, y: 4.2, w: 4, fontSize: 28, bold: true, color: "0f766e", align: "center" });

    slide2.addShape(pres.ShapeType.rect, { x: 5.5, y: 3.5, w: 4, h: 1.5, fill: "f8fafc", line: {color: "008080", width: 1} });
    slide2.addText("SYSTEM RECORDS", { x: 5.5, y: 3.7, w: 4, fontSize: 12, color: "64748b", align: "center" });
    slide2.addText(kpiStats.totalReports.toLocaleString("en-IN"), { x: 5.5, y: 4.2, w: 4, fontSize: 28, bold: true, color: "0f766e", align: "center" });

    pres.writeFile({ fileName: `12_Deeni_Kaam_PPT_${new Date().getTime()}.pptx` });
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
        link.download = `Deeni_Kaam_Report_Card_${new Date().getTime()}.jpg`;
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
                    <h1 className="text-xl sm:text-2xl font-black text-slate-800 tracking-wide">12 DEENI KAAM <span className="font-light text-teal-700">REPORT</span></h1>
                    <p className="text-[10px] font-bold text-teal-700 uppercase tracking-widest mt-0.5">Live Data Synchronization</p>
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
                  <div className="w-2 h-2 rounded-full bg-red-500 animate-ping"></div>{fetchError}
                </div>
              )}

              {/* KPIs */}
              <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-6">
                {[
                  { title: "Active Muballigh", val: kpiStats.totalMuballigh, icon: User },
                  { title: "Total Masajid", val: kpiStats.totalMasjid, icon: MapPin },
                  { title: "Zeili Halqe", val: kpiStats.totalHalqe, icon: Layers },
                  { title: "System Records", val: kpiStats.totalReports, icon: Activity }
                ].map((kpi, idx) => (
                  <div key={idx} className="bg-white p-6 rounded-xl border border-slate-200 shadow-sm relative overflow-hidden group">
                    <div className="relative flex justify-between items-center z-10">
                      <div className="flex items-center gap-4">
                          <div className="p-3.5 rounded-xl bg-teal-50 text-teal-600 border border-teal-100">
                            <kpi.icon className="w-6 h-6" />
                          </div>
                          <div>
                              <p className="text-[11px] font-bold text-slate-500 uppercase tracking-widest mb-1.5">{kpi.title}</p>
                              <h3 className="text-3xl font-black text-teal-800 tracking-tight">{kpi.val.toLocaleString("en-IN")}</h3>
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
                <div className="grid grid-cols-2 md:grid-cols-6 gap-4">
                  <div className="col-span-2 md:col-span-1" data-html2canvas-ignore="true">
                    <label className="block text-[10px] font-bold text-slate-500 uppercase tracking-widest mb-2">Search</label>
                    <div className="relative">
                      <span className="absolute inset-y-0 left-0 pl-3 flex items-center text-slate-400"><Search className="w-4 h-4" /></span>
                      <input type="text" placeholder="Search..." value={searchTerm} onChange={(e) => setSearchTerm(e.target.value)} className="w-full pl-9 pr-3 py-2.5 rounded-lg border border-slate-300 bg-white text-sm font-medium text-slate-700 focus:outline-none focus:border-teal-500 focus:ring-1 focus:ring-teal-500" />
                    </div>
                  </div>
                  {[{ label: "Year", value: selectedYear, setter: setSelectedYear, opts: filterOptions.years },
                    { label: "Month", value: selectedMonth, setter: setSelectedMonth, opts: filterOptions.months },
                    { label: "Region", value: selectedRegion, setter: setSelectedRegion, opts: filterOptions.regions },
                    { label: "District", value: selectedDistrict, setter: setSelectedDistrict, opts: filterOptions.districts },
                    { label: "Category", value: selectedCategory, setter: setSelectedCategory, opts: filterOptions.categories }
                  ].map((filter, idx) => (
                    <div key={idx} className="flex flex-col flex-1 min-w-[130px]">
                      <label className="block text-[10px] font-bold text-slate-500 uppercase tracking-widest mb-2">{filter.label}</label>
                      <select value={filter.value} onChange={(e) => filter.setter(e.target.value)} disabled={officeUser?.[filter.label === "Category" ? "department" : filter.label.toLowerCase()] && officeUser[filter.label === "Category" ? "department" : filter.label.toLowerCase()].toLowerCase() !== "all"} className="w-full p-2.5 rounded-lg border border-slate-300 bg-white text-sm font-medium text-slate-700 focus:outline-none focus:border-teal-500 focus:ring-1 focus:ring-teal-500 appearance-none cursor-pointer disabled:bg-slate-50 disabled:text-slate-400">
                        <option value="All">All {filter.label}s</option>
                        {filter.opts.map(o => <option key={o} value={o}>{o}</option>)}
                      </select>
                    </div>
                  ))}
                </div>
              </div>
            </main>
        </div> {/* END OF REPORT CARD VIEW */}

        <main className="max-w-[1400px] mx-auto px-4 sm:px-6 lg:px-8 pb-8 mt-2">
          {/* Optimized Table Section - Excluded from JPEG */}
          <div className="bg-white rounded-xl border border-slate-200 shadow-sm mt-6">
            <div className="p-4 border-b border-slate-200 flex flex-col md:flex-row justify-between items-center gap-4 bg-[#e0f2f1] rounded-t-xl">
              <div>
                <h2 className="text-sm font-bold text-teal-800 uppercase tracking-widest">Detailed Telemetry Output</h2>
                <p className="text-[10px] text-slate-600 mt-1 uppercase tracking-widest">Source: {rawData.length} &bull; Visible: {filteredData.length}</p>
              </div>
            </div>

            <div className="overflow-x-auto w-full custom-scrollbar">
              <table className="min-w-full text-left text-[11px] lg:text-xs">
                <thead className="bg-[#008b8b]">
                  <tr>
                    {["Location", "Category", "Parameter", "Value", "Timestamp"].map(h => (
                      <th key={h} className={`px-4 py-3 font-bold text-white uppercase tracking-wider ${h==="Value"?"text-right":""}`}>{h}</th>
                    ))}
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {filteredData.length > 0 ? (
                    filteredData.map((row, idx) => (
                      <tr key={idx} className="hover:bg-slate-50 transition-colors">
                        <td className="px-4 py-3 text-slate-600 whitespace-nowrap">
                          <div className="font-bold text-slate-800 leading-tight">{row["District"] || "-"}</div>
                          <div className="text-[10px] font-medium text-slate-500 mt-0.5">{row["Region"] || "-"}</div>
                        </td>
                        <td className="px-4 py-3 text-slate-600 whitespace-nowrap">
                          <span className="inline-flex bg-slate-100 border border-slate-200 px-2 py-1 rounded text-[10px] uppercase font-bold text-slate-500">
                            {row["Category"] || "-"}
                          </span>
                        </td>
                        <td className="px-4 py-3 text-slate-700 font-medium leading-tight">{row["Fields"] || "-"}</td>
                        <td className="px-4 py-3 text-right font-extrabold text-teal-700 text-sm whitespace-nowrap bg-teal-50/40">
                          {row["Report Value"] || "0"}
                        </td>
                        <td className="px-4 py-3 text-slate-600 leading-tight">
                          <div className="inline-flex items-center gap-1.5 px-2 py-1 bg-slate-50 rounded border border-slate-200 text-[10px] font-bold text-slate-600">
                            <Calendar className="w-3 h-3 text-slate-400" /> {row["Month"]} {row["Year"]}
                          </div>
                        </td>
                      </tr>
                    ))
                  ) : (
                    <tr>
                      <td colSpan="5" className="px-6 py-12 text-center text-slate-500 text-xs uppercase tracking-widest">
                        <div className="flex flex-col items-center justify-center gap-3">
                          <Activity className="w-6 h-6 opacity-40 text-teal-600" />
                          <span>{loading ? "Establishing connection..." : "No matching telemetry found"}</span>
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
