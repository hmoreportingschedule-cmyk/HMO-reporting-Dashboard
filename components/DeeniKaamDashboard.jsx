"use client";
import React, { useState, useEffect, useMemo, useTransition } from "react";
import Papa from "papaparse";
import { LogOut, RefreshCw, Filter, Search, Activity, ArrowLeft, Download, BookOpenCheck, Image as ImageIcon, Clock, LayoutDashboard, BarChart3, TrendingUp, Target, PieChart, Layers } from "lucide-react";
import * as XLSX from "xlsx";
import pptxgen from "pptxgenjs";
import html2canvas from "html2canvas";

const SHEET_URLS = [
  "https://docs.google.com/spreadsheets/d/1P1Ul-jXOfFfhuQLTKeQ-zOynKnCmywH-_gjZ9nJ8tO0/export?format=csv",
  "https://docs.google.com/spreadsheets/d/1S2tEIyaN8p-yu4Vd_GVumqBqwzgTzM3zlms4DwCJr00/export?format=csv",
  "https://docs.google.com/spreadsheets/d/1yWVgL9IVGrQFElLNeO8X_UGAIDSAGF7P8M31gGtoki8/export?format=csv"
];

const MASTER_CATEGORIES_MAP = [
  { category: "Basic", deeniKaam: "Tanzimi Malumat", field: "Total Active Muballigh" },
  { category: "Basic", deeniKaam: "Tanzimi Malumat", field: "Total Moallimin" },
  { category: "Basic", deeniKaam: "Tanzimi Malumat", field: "Total Masjid" },
  { category: "Basic", deeniKaam: "Tanzimi Malumat", field: "Apni Masjid" },
  { category: "Basic", deeniKaam: "Tanzimi Malumat", field: "Total Zeili Halqe" },
  { category: "Basic", deeniKaam: "Tanzimi Malumat", field: "Total Zaili Halqe Taqarrur" },
  { category: "Daily", deeniKaam: "Fajr Ke Liye Jagaen", field: "Fajr Ke Liye Jagaen" },
  { category: "Daily", deeniKaam: "Tafseer Sunna/Sunnana", field: "Tafseer Sunna/Sunnana" },
  { category: "Daily", deeniKaam: "Dars", field: "Masjid Dars" },
  { category: "Daily", deeniKaam: "Dars", field: "Area Dars" },
  { category: "Daily", deeniKaam: "Dars", field: "Ghar Dars" },
  { category: "Daily", deeniKaam: "Dars", field: "Total Dars" },
  { category: "Daily", deeniKaam: "Madrasatul Madina Baligan", field: "Madrasatul Madina Baligan Tadad (Masjid)" },
  { category: "Daily", deeniKaam: "Madrasatul Madina Baligan", field: "Madrasatul Madina Baligan Shurqa (Masjid)" },
  { category: "Daily", deeniKaam: "Madrasatul Madina Baligan", field: "Madrasatul Madina Baligan Tadad (Others)" },
  { category: "Daily", deeniKaam: "Madrasatul Madina Baligan", field: "Madrasatul Madina Baligan Shurqa (Others)" },
  { category: "Daily", deeniKaam: "Madrasatul Madina Baligan", field: "Total Madrasarul Madina Baligan Tadad" },
  { category: "Daily", deeniKaam: "Madrasatul Madina Baligan", field: "Total Madrasarul Madina Baligan Shurqa" },
  { category: "Weekly", deeniKaam: "Haftwar Ijtima", field: "Haftwar Ijtima Tadad" },
  { category: "Weekly", deeniKaam: "Haftwar Ijtima", field: "Haftwar Ijtima Shurqa" },
  { category: "Weekly", deeniKaam: "Haftwar Ijtima", field: "Raat Guzarne Walo Ki Tadad" },
  { category: "Weekly", deeniKaam: "Haftwar Ijtima", field: "Ijtima me Shurqa ki tadad 120 se zyada hai" },
  { category: "Weekly", deeniKaam: "Madani Muzakirah", field: "Madani Muzakirah Maqaamat" },
  { category: "Weekly", deeniKaam: "Madani Muzakirah", field: "Madani Muzakirah Shurqa" },
  { category: "Weekly", deeniKaam: "Ek Din Raahe Khuda Me", field: "Ek Din Raahe Khuda Me Tadad" },
  { category: "Weekly", deeniKaam: "Ek Din Raahe Khuda Me", field: "Ek Din Raahe Khuda Me Shurqa" },
  { category: "Weekly", deeniKaam: "Madani Halqa", field: "Madani Halqa Tadad" },
  { category: "Weekly", deeniKaam: "Madani Halqa", field: "Madani Halqa Shurqa" },
  { category: "Weekly", deeniKaam: "Haftwar Risala", field: "Haftwar Risala Padhne Wale / Sunne Wale" },
  { category: "Weekly", deeniKaam: "Alaqai Dora", field: "Alaqai Dora Kitni Baar" },
  { category: "Monthly", deeniKaam: "Qafila", field: "3 Din Qafila Tadad" },
  { category: "Monthly", deeniKaam: "Qafila", field: "3 Din Qafila Shurqa" },
  { category: "Monthly", deeniKaam: "Qafila", field: "12 Din Qafila Tadad" },
  { category: "Monthly", deeniKaam: "Qafila", field: "12 Din Qafila Shurqa" },
  { category: "Monthly", deeniKaam: "Qafila", field: "1 Maah Qafila Tadad" },
  { category: "Monthly", deeniKaam: "Qafila", field: "1 Maah Qafila Shurqa" },
  { category: "Monthly", deeniKaam: "Qafila", field: "12 Maah Qafila Tadad" },
  { category: "Monthly", deeniKaam: "Qafila", field: "12 Maah Qafila Shurqa" },
  { category: "Monthly", deeniKaam: "Courses", field: "Short Courses (Gair Riyaeshi) Tadad" },
  { category: "Monthly", deeniKaam: "Courses", field: "Short Courses (Gair Riyaeshi) Shurqa" },
  { category: "Monthly", deeniKaam: "Courses", field: "Long Courses (Riyaeshi) Tadad" },
  { category: "Monthly", deeniKaam: "Courses", field: "Long Courses (Riyaeshi) Shurqa" },
  { category: "Monthly", deeniKaam: "Courses", field: "Total Courses Tadad" },
  { category: "Monthly", deeniKaam: "Courses", field: "Total Courses Shurqa" },
  { category: "Monthly", deeniKaam: "Neak Aamal", field: "Neak Aamal Risala Wasool" }
];

const parseSheet = (url) => {
  return new Promise((resolve) => {
    Papa.parse(url, {
      download: true,
      header: true,
      skipEmptyLines: true,
      complete: (results) => resolve(results?.data || []),
      error: () => resolve([])
    });
  });
};

const sameClient = (a, b) => String(a || "").trim().toLowerCase() === String(b || "").trim().toLowerCase();

const formatMonthYearLabel = (val) => {
  if (!val) return "Month";
  if (/^\d{4}-\d{2}$/.test(val)) {
    const [year, month] = val.split("-");
    const date = new Date(parseInt(year, 10), parseInt(month, 10) - 1, 1);
    return isNaN(date.getTime()) ? val : date.toLocaleString('en-US', { month: 'short', year: 'numeric' });
  }
  return val;
};

const parseRowMonth = (row) => {
  if (!row) return "";
  let raw = String(row["Month"] || row["Date"] || row["Report Date"] || row["ReportMonth"] || "").trim();
  if (!raw) return "";

  if (/^\d{4}-\d{2}$/.test(raw)) return raw;
  if (/^\d{2}\/\d{4}$/.test(raw)) {
    const [mo, yr] = raw.split("/");
    return `${yr}-${mo}`;
  }

  let parsed = Date.parse(raw);
  if (!isNaN(parsed)) {
    let d = new Date(parsed);
    let yr = d.getFullYear();
    let mo = String(d.getMonth() + 1).padStart(2, '0');
    return `${yr}-${mo}`;
  }

  let lower = raw.toLowerCase();
  const months = {jan: '01', feb: '02', mar: '03', apr: '04', may: '05', jun: '06', jul: '07', aug: '08', sep: '09', oct: '10', nov: '11', dec: '12'};
  
  let matchYear = raw.match(/\b(20\d{2}|19\d{2})\b/);
  let yr = matchYear ? matchYear[0] : "2026";

  for (let m in months) {
    if (lower.includes(m)) {
      return `${yr}-${months[m]}`;
    }
  }
  return raw;
};

const MiniTable = ({ title, data = [] }) => (
  <div className="bg-white rounded-xl shadow-sm border border-slate-200 overflow-hidden flex flex-col h-[340px]">
    <div className="p-4 border-b border-slate-100 bg-teal-50">
      <h3 className="text-teal-800 font-bold uppercase tracking-widest text-xs">{title}</h3>
    </div>
    <div className="overflow-y-auto flex-1 custom-scrollbar">
      <table className="w-full text-left text-xs">
        <thead className="bg-teal-700 text-white sticky top-0 z-10">
          <tr>
            <th className="py-2.5 px-4 font-bold uppercase tracking-wider text-center">Name</th>
            <th className="py-2.5 px-4 font-bold uppercase tracking-wider text-right">Qty</th>
          </tr>
        </thead>
        <tbody className="divide-y divide-slate-100">
          {Array.isArray(data) && data.length > 0 ? data.slice(0, 15).map((d, i) => (
            <tr key={i} className="hover:bg-slate-50 transition-colors">
              <td className="py-2.5 px-4 text-slate-700 text-center font-medium">{d?.label || "-"}</td>
              <td className="py-2.5 px-4 text-teal-700 font-bold text-right">{(d?.count || 0).toLocaleString("en-IN")}</td>
            </tr>
          )) : <tr><td colSpan="2" className="text-center py-4 text-slate-500">No data available</td></tr>}
        </tbody>
      </table>
    </div>
  </div>
);

export default function DeeniKaamDashboard({ onBack, onLogout, officeUser }) {
  const [rawData, setRawData] = useState([]);
  const [loading, setLoading] = useState(false);
  const [fetchError, setFetchError] = useState("");
  const [, startTransition] = useTransition();

  const [activeViewMode, setActiveViewMode] = useState("table"); 
  const [activeTab, setActiveTab] = useState("Monthly Report");
  
  const [colMonthMain, setColMonthMain] = useState("2026-12"); 
  const [colMonth1, setColMonth1] = useState("2026-11"); 
  const [colMonth2, setColMonth2] = useState("2026-12"); 
  
  const [headerTargetMode, setHeaderTargetMode] = useState("");

  const [region, setRegion] = useState(officeUser?.region && officeUser.region.toLowerCase() !== "all" ? officeUser.region : "");
  const [state, setState] = useState(officeUser?.state && officeUser.state.toLowerCase() !== "all" ? officeUser.state : "");
  const [division, setDivision] = useState(officeUser?.division && officeUser.division.toLowerCase() !== "all" ? officeUser.division : "");
  const [district, setDistrict] = useState(officeUser?.district && officeUser.district.toLowerCase() !== "all" ? officeUser.district : "");
  
  const [selectedCategory, setSelectedCategory] = useState(""); 
  const [selectedDeeniKaam, setSelectedDeeniKaam] = useState(""); 
  const [selectedField, setSelectedField] = useState(""); 
  
  const [searchTerm, setSearchTerm] = useState("");
  const [currentPage, setCurrentPage] = useState(1);
  const rowsPerPage = 12;

  const [now, setNow] = useState(null);
  useEffect(() => {
    setNow(new Date());
    const timer = setInterval(() => setNow(new Date()), 1000);
    return () => clearInterval(timer);
  }, []);
  const dateStr = now ? now.toLocaleDateString('en-GB', { day: '2-digit', month: '2-digit', year: 'numeric' }).replace(/\//g, '-') : "";
  const timeStr = now ? now.toLocaleTimeString('en-US', { hour: '2-digit', minute: '2-digit', second: '2-digit', hour12: true }) : "";

  useEffect(() => {
    fetchData();
  }, []);

  const fetchData = async () => {
    setLoading(true);
    setFetchError("");
    try {
      const resultsArray = await Promise.all(SHEET_URLS.map(url => parseSheet(url)));
      const combinedData = resultsArray.flat().filter(Boolean);
      if (combinedData.length > 0) {
        const processed = combinedData.map(row => {
          if (!row) return null;
          let reportVal = Number(String(row["Report Value"] || row.report || "0").replace(/,/g, ""));
          let targetVal = Number(String(row["Target"] || "0").replace(/,/g, ""));
          let fld = String(row["Fileds"] || row["Fields"] || row["Deeni Activities"] || "").trim();
          
          let foundMap = MASTER_CATEGORIES_MAP.find(m => sameClient(m.field, fld));
          let cat = String(row["Category"] || foundMap?.category || "Basic").trim();
          let dk = String(row["Deeni Kaam"] || foundMap?.deeniKaam || "Tanzimi Malumat").trim();

          return {
            ...row,
            "Category": cat,
            "Deeni Kaam": dk,
            "Fileds": fld,
            Target: isNaN(targetVal) ? 0 : targetVal,
            Achievement: Number(String(row["Achievement"] || row["Achivement"] || "0").replace(/,/g, "")),
            NormalizedMonth: parseRowMonth(row),
            NumericReport: isNaN(reportVal) ? 0 : reportVal
          };
        }).filter(Boolean);
        setRawData(processed);
        
        const months = [...new Set(processed.map(r => r.NormalizedMonth).filter(Boolean))].sort();
        if (months.length > 0) {
          const latest = months[months.length - 1];
          setColMonthMain(latest);
          setColMonth2(latest);
          if (months.length > 1) {
            setColMonth1(months[months.length - 2]);
          } else {
            setColMonth1(latest);
          }
        }
      } else {
        setFetchError("Data stream empty.");
      }
      setLoading(false);
    } catch (err) {
      setFetchError("Network latency issue.");
      setLoading(false);
    }
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

  const availableCategories = useMemo(() => {
    return [...new Set(MASTER_CATEGORIES_MAP.map(m => m.category))].sort();
  }, []);

  const availableDeeniKaam = useMemo(() => {
    let subset = MASTER_CATEGORIES_MAP;
    if (selectedCategory) {
      subset = MASTER_CATEGORIES_MAP.filter(m => sameClient(m.category, selectedCategory));
    }
    return [...new Set(subset.map(m => m.deeniKaam))].sort();
  }, [selectedCategory]);

  const availableFields = useMemo(() => {
    let subset = MASTER_CATEGORIES_MAP;
    if (selectedDeeniKaam) {
      subset = MASTER_CATEGORIES_MAP.filter(m => sameClient(m.deeniKaam, selectedDeeniKaam));
    }
    return [...new Set(subset.map(m => m.field))].sort();
  }, [selectedDeeniKaam]);

  const availableMonths = useMemo(() => {
    const months = [...new Set(rawData.map(r => r.NormalizedMonth).filter(Boolean))];
    return months.sort().reverse();
  }, [rawData]);

  const leftHeaderTitle = useMemo(() => {
    if (district) return "";
    if (division) return "DISTRICT";
    if (state) return "DIVISION";
    if (region) return "STATE";
    return "COUNTRY";
  }, [region, state, division, district]);

  // Unified Dynamic Processing & Aggregation Engine
  const processedTableData = useMemo(() => {
    if (!Array.isArray(rawData) || rawData.length === 0) return [];

    const getLeftColName = (r) => {
      if (district) return "";
      if (division) return r["District"] || "-";
      if (state) return r["Division"] || "-";
      if (region) return r["State"] || "-";
      return "India";
    };

    const validFieldsMap = {};
    MASTER_CATEGORIES_MAP.forEach(m => {
      if (selectedCategory && !sameClient(m.category, selectedCategory)) return;
      if (selectedDeeniKaam && !sameClient(m.deeniKaam, selectedDeeniKaam)) return;
      if (selectedField && !sameClient(m.field, selectedField)) return;
      validFieldsMap[m.field] = { category: m.category, deeniKaam: m.deeniKaam };
    });

    rawData.forEach(r => {
      const fld = r["Fileds"];
      if (fld && !validFieldsMap[fld]) {
        if ((!selectedCategory || sameClient(r["Category"], selectedCategory)) &&
            (!selectedDeeniKaam || sameClient(r["Deeni Kaam"], selectedDeeniKaam)) &&
            (!selectedField || sameClient(fld, selectedField))) {
          validFieldsMap[fld] = { category: r["Category"], deeniKaam: r["Deeni Kaam"] };
        }
      }
    });

    const filteredRows = rawData.filter((row) => {
      if (!row) return false;
      const matchRegion = !region || sameClient(row["Region"], region);
      const matchState = !state || sameClient(row["State"], state);
      const matchDivision = !division || sameClient(row["Division"], division);
      const matchDistrict = !district || sameClient(row["District"], district);
      
      const catVal = row["Category"] || "";
      const matchCat = !selectedCategory || sameClient(catVal, selectedCategory);

      const deeniKaamVal = row["Deeni Kaam"] || "";
      const matchDeeniKaam = !selectedDeeniKaam || sameClient(deeniKaamVal, selectedDeeniKaam);

      const fieldVal = row["Fileds"] || "";
      const matchField = !selectedField || sameClient(fieldVal, selectedField);

      const matchSearch = !searchTerm || JSON.stringify(row).toLowerCase().includes(searchTerm.toLowerCase().trim());
      return matchRegion && matchState && matchDivision && matchDistrict && matchCat && matchDeeniKaam && matchField && matchSearch;
    });

    const subEntitiesSet = new Set();
    filteredRows.forEach(r => {
      subEntitiesSet.add(getLeftColName(r));
    });
    const subEntities = subEntitiesSet.size > 0 ? [...subEntitiesSet].filter(Boolean).sort() : ["India"];

    const resultList = [];

    subEntities.forEach(subEntity => {
      Object.keys(validFieldsMap).forEach(fldName => {
        const meta = validFieldsMap[fldName];

        let sumMain = 0; 
        let sumVal1 = 0; 
        let sumVal2 = 0; 
        let sumTarget = 0;

        filteredRows.forEach(r => {
          if (!r) return;
          if (String(r["Fileds"] || "").trim() !== fldName) return;

          const ent = getLeftColName(r);
          if (subEntity !== "India" && ent !== subEntity) return;

          const mo = r.NormalizedMonth || "";
          const repVal = r.NumericReport || 0;
          const targetVal = r.Target || 0;

          if (colMonthMain && mo === colMonthMain) {
            sumMain += repVal;
            let baseT = targetVal;
            if (headerTargetMode === "26%") baseT = baseT * 0.26;
            else if (headerTargetMode === "52%") baseT = baseT * 0.52;
            sumTarget += baseT;
          }

          if (colMonth1 && mo === colMonth1) {
            sumVal1 += repVal;
          }

          if (colMonth2 && mo === colMonth2) {
            sumVal2 += repVal;
          }
        });

        if (sumMain > 0 || sumVal1 > 0 || sumVal2 > 0 || (!selectedCategory && !selectedDeeniKaam && !selectedField)) {
          let v1 = sumVal1;
          let v2 = sumVal2;

          // Tab-specific calculation rules (Monthly, Average, Quarterly, Progress)
          if (activeTab === "Average Report" || activeViewMode === "average") {
            v1 = Math.round(v1 / 2);
            v2 = Math.round(v2 / 2);
          } else if (activeTab === "Quarterly Report") {
            v1 = v1 * 3;
            v2 = v2 * 3;
          }

          let diffPercent = 0;
          if (v1 > 0) {
            diffPercent = ((v2 - v1) / v1) * 100;
          } else if (v1 === 0 && v2 > 0) {
            diffPercent = 100;
          } else if (v1 > 0 && v2 === 0) {
            diffPercent = -100;
          }

          let compStr = `${diffPercent >= 0 ? "+" : ""}${diffPercent.toFixed(1)}%`;
          let targetRounded = Math.round(sumTarget);
          let achPct = targetRounded > 0 ? ((sumMain / targetRounded) * 100).toFixed(1) + "%" : "-";

          resultList.push({
            LeftColValue: subEntity,
            CategoryName: meta.category,
            DeeniKaamName: meta.deeniKaam,
            DeeniActivity: fldName,
            DynamicReportValue: sumMain,
            DynamicTarget: targetRounded,
            DynamicAchievement: achPct,
            Val1: v1,
            Val2: v2,
            CalculatedComparison: compStr,
            ColMonth1Label: formatMonthYearLabel(colMonth1),
            ColMonth2Label: formatMonthYearLabel(colMonth2),
            MainMonthLabel: formatMonthYearLabel(colMonthMain)
          });
        }
      });
    });

    return resultList;
  }, [rawData, region, state, division, district, selectedCategory, selectedDeeniKaam, selectedField, headerTargetMode, searchTerm, colMonthMain, colMonth1, colMonth2, activeTab, activeViewMode]);

  const pagedRows = useMemo(() => {
    if (!Array.isArray(processedTableData)) return [];
    const start = (currentPage - 1) * rowsPerPage;
    return processedTableData.slice(start, start + rowsPerPage);
  }, [processedTableData, currentPage]);

  const dynamicGraphData = useMemo(() => {
    const map = {};
    if (Array.isArray(processedTableData)) {
      processedTableData.forEach(x => {
        if (!x) return;
        const k = String(x.LeftColValue || "Unknown").trim();
        const n = Number(x.DynamicReportValue || 0);
        map[k] = (map[k] || 0) + (isNaN(n) ? 0 : n);
      });
    }
    return Object.keys(map).sort((a,b) => map[b] - map[a]).map(k => ({ label: k, count: map[k] }));
  }, [processedTableData]);

  const dynamicGraphTitle = division ? "REPORTS BY DISTRICT" : state ? "REPORTS BY DIVISION" : region ? "REPORTS BY STATE" : "REPORTS BY REGION";
  const maxDynamicCount = dynamicGraphData.length ? Math.max(...dynamicGraphData.map(d => d.count)) : 0;

  const downloadExcel = () => {
    if (!processedTableData.length) return alert("No data to export");
    const worksheet = XLSX.utils.json_to_sheet(processedTableData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Deeni Kaam Data");
    XLSX.writeFile(workbook, `12_Deeni_Kaam_RawData_${new Date().getTime()}.xlsx`);
  };

  const downloadPPT = () => {
    if (!processedTableData.length) return alert("No data to export");
    let pres = new pptxgen();
    let slide1 = pres.addSlide();
    slide1.background = { color: "ffffff" };
    slide1.addText("12 DEENI KAAM REPORT", { x: 1, y: 2, w: 8, fontSize: 36, bold: true, color: "008080", align: "center" });
    pres.writeFile({ fileName: `12_Deeni_Kaam_PPT_${new Date().getTime()}.pptx` });
  };

  const downloadJPEG = async () => {
    const element = document.getElementById("report-card-view");
    if (!element) return;
    try {
        const canvas = await html2canvas(element, { scale: 2, backgroundColor: "#e0f2f1", useCORS: true, ignoreElements: (node) => node.hasAttribute("data-html2canvas-ignore") });
        const link = document.createElement('a');
        link.download = `Deeni_Kaam_Report_Card_${new Date().getTime()}.jpg`;
        link.href = canvas.toDataURL('image/jpeg', 0.9);
        link.click();
    } catch (err) {
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
            <header className="max-w-[1400px] mx-auto px-4 sm:px-6 lg:px-8 relative z-35">
              <div className="bg-white border border-slate-200 shadow-sm rounded-2xl px-6 py-5 flex flex-col md:flex-row items-center justify-between gap-6">
                <div className="flex items-center gap-4 w-full md:w-auto">
                  <button data-html2canvas-ignore="true" onClick={onBack} className="p-3 bg-slate-100 hover:bg-slate-200 border border-slate-200 text-slate-700 rounded-xl shadow-sm transition-colors">
                    <ArrowLeft className="w-5 h-5" />
                  </button>
                  <div className="p-3 bg-teal-700 text-white rounded-xl shadow-sm border border-teal-700 hidden sm:block">
                    <BookOpenCheck className="w-6 h-6" />
                  </div>
                  <div>
                    <h1 className="text-xl sm:text-2xl font-black text-slate-800 tracking-wide">12 DEENI KAAM <span className="font-light text-teal-700">REPORT HUB</span></h1>
                    <p className="text-[10px] font-bold text-teal-700 uppercase tracking-widest mt-0.5">Professional Multi-View Analytics Dashboard</p>
                  </div>
                </div>
                
                <div className="flex flex-wrap items-center justify-center md:justify-end gap-3 w-full md:w-auto">
                  <div className="text-[10px] font-bold text-slate-600 uppercase tracking-widest bg-slate-50 px-3 py-2 rounded-lg border border-slate-200 hidden lg:block">
                    User: <span className="text-teal-700">{officeUser?.name || officeUser?.userId || "Admin"}</span>
                  </div>
                  <button data-html2canvas-ignore="true" onClick={fetchData} disabled={loading} className="flex items-center gap-2 text-xs font-bold tracking-wider uppercase bg-teal-600 hover:bg-teal-700 text-white px-4 py-2.5 rounded-lg border border-teal-600 transition-all active:scale-95">
                    <RefreshCw className={`w-4 h-4 ${loading ? "animate-spin" : ""}`} /> <span className="hidden sm:inline">Sync</span>
                  </button>
                  <button data-html2canvas-ignore="true" onClick={downloadJPEG} className="flex items-center gap-2 text-xs font-bold tracking-wider uppercase bg-[#0f766e] hover:bg-[#115e59] text-white px-4 py-2.5 rounded-lg border border-[#0f766e] transition-all active:scale-95 shadow-md">
                    <ImageIcon className="w-4 h-4" /> <span className="hidden sm:inline">JPEG</span>
                  </button>
                  <button data-html2canvas-ignore="true" onClick={downloadPPT} className="flex items-center gap-2 text-xs font-bold tracking-wider uppercase bg-teal-600 hover:bg-teal-700 text-white px-4 py-2.5 rounded-lg border border-teal-600 transition-all active:scale-95">
                    <span className="hidden sm:inline">PPT</span>
                  </button>
                  <button data-html2canvas-ignore="true" onClick={downloadExcel} className="flex items-center gap-2 text-xs font-bold tracking-wider uppercase bg-[#198754] hover:bg-green-700 text-white px-4 py-2.5 rounded-lg border border-[#198754] transition-all active:scale-95">
                    <Download className="w-4 h-4" /> <span className="hidden sm:inline">Excel</span>
                  </button>
                  
                  <div className="w-full flex justify-end items-center gap-3 mt-2">
                    <div className="text-[11px] font-extrabold text-black uppercase tracking-widest bg-white px-4 py-2.5 rounded-lg border border-slate-200 hidden lg:flex items-center gap-2 whitespace-nowrap shadow-sm">
                      <Clock className="w-4 h-4 text-teal-600" />
                      <span>Date: {dateStr} &bull; Time: {timeStr}</span>
                    </div>
                    <button data-html2canvas-ignore="true" onClick={onLogout} className="flex items-center gap-2 text-xs font-bold tracking-wider uppercase bg-[#dc3545] hover:bg-red-700 text-white px-4 py-2.5 rounded-lg border border-[#dc3545] transition-all active:scale-95">
                      <LogOut className="w-4 h-4" /> <span className="hidden sm:inline">Logout</span>
                    </button>
                  </div>
                </div>
              </div>
            </header>

            <main className="max-w-[1400px] mx-auto px-4 sm:px-6 lg:px-8 pt-6 space-y-6">
              {fetchError && (
                <div data-html2canvas-ignore="true" className="p-4 bg-red-50 border border-red-200 text-red-600 rounded-xl text-sm font-medium flex items-center gap-3">
                  <div className="w-2 h-2 rounded-full bg-red-500 animate-ping"></div>{fetchError}
                </div>
              )}

              {/* 4 PROFESSIONAL VIEW MODE BUTTONS */}
              <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4" data-html2canvas-ignore="true">
                {[
                  { id: "table", label: "DETAILED TABLE REPORT", icon: LayoutDashboard, desc: "Grid & Month-wise data" },
                  { id: "graphs", label: "VISUAL ANALYTICS & GRAPHS", icon: BarChart3, desc: "District/State charts" },
                  { id: "average", label: "AVERAGE & COMPARISON", icon: TrendingUp, desc: "Range growth analysis" },
                  { id: "targets", label: "TARGETS & ACHIEVEMENT", icon: Target, desc: "Goal performance matrix" }
                ].map((btn) => {
                  const IconComp = btn.icon;
                  const isActive = activeViewMode === btn.id;
                  return (
                    <button
                      key={btn.id}
                      onClick={() => startTransition(() => setActiveViewMode(btn.id))}
                      className={`p-5 rounded-2xl border text-left transition-all shadow-sm flex items-start gap-3.5 ${
                        isActive
                          ? "bg-[#0f4c47] text-white border-[#0f4c47] shadow-md transform -translate-y-0.5"
                          : "bg-white text-slate-700 hover:bg-slate-50 border-slate-200"
                      }`}
                    >
                      <div className={`p-3 rounded-xl ${isActive ? "bg-teal-700 text-white" : "bg-teal-50 text-teal-700"}`}>
                        <IconComp className="w-5 h-5" />
                      </div>
                      <div>
                        <h3 className="text-xs font-black uppercase tracking-wider">{btn.label}</h3>
                        <p className={`text-[11px] mt-1 ${isActive ? "text-teal-200" : "text-slate-400 font-medium"}`}>{btn.desc}</p>
                      </div>
                    </button>
                  );
                })}
              </div>

              {activeViewMode === "table" && (
                <div className="flex gap-3 border-b border-slate-200 pb-2" data-html2canvas-ignore="true">
                  {["Monthly Report", "Average Report", "Quarterly Report", "Progress Report"].map((tab) => (
                    <button
                      key={tab}
                      onClick={() => startTransition(() => setActiveTab(tab))}
                      className={`px-5 py-2.5 rounded-xl text-xs font-bold uppercase tracking-wider transition-all shadow-sm ${
                        activeTab === tab
                          ? "bg-teal-700 text-white shadow-md"
                          : "bg-white text-slate-700 hover:bg-slate-100 border border-slate-200"
                      }`}
                    >
                      {tab}
                    </button>
                  ))}
                </div>
              )}

              {/* DYNAMIC FILTERS BAR */}
              <div className="bg-white p-5 rounded-xl border border-slate-200 shadow-sm">
                <div className="flex flex-col lg:flex-row justify-between lg:items-center gap-4 mb-5 border-b border-slate-100 pb-4">
                  <div className="flex items-center gap-3">
                    <Filter className="w-5 h-5 text-teal-700" />
                    <h2 className="text-sm font-bold text-slate-800 tracking-widest uppercase">Dynamic Hierarchy & Field Filters</h2>
                  </div>
                </div>

                <div className="grid grid-cols-2 md:grid-cols-7 gap-3">
                  <div className="col-span-2 md:col-span-1" data-html2canvas-ignore="true">
                    <label className="block text-[10px] font-bold text-slate-500 uppercase tracking-widest mb-2">Search</label>
                    <div className="relative">
                      <span className="absolute inset-y-0 left-0 pl-3 flex items-center text-slate-400"><Search className="w-4 h-4" /></span>
                      <input type="text" placeholder="Search..." value={searchTerm} onChange={(e) => setSearchTerm(e.target.value)} className="w-full pl-9 pr-3 py-2.5 rounded-lg border border-slate-300 bg-white text-sm font-medium text-slate-700 focus:outline-none focus:border-teal-500 focus:ring-1 focus:ring-teal-500" />
                    </div>
                  </div>

                  {/* Category Filter */}
                  <div className="flex flex-col flex-1 min-w-[110px]">
                    <label className="block text-[10px] font-bold text-slate-500 uppercase tracking-widest mb-2">Category</label>
                    <select value={selectedCategory} onChange={(e) => { setSelectedCategory(e.target.value); setSelectedDeeniKaam(""); setSelectedField(""); }} className="w-full p-2.5 rounded-lg border border-slate-300 bg-white text-sm font-medium text-slate-700 focus:outline-none focus:border-teal-500 focus:ring-1 focus:ring-teal-500 appearance-none cursor-pointer">
                      <option value="">All Categories</option>
                      {availableCategories.map(cat => <option key={cat} value={cat}>{cat}</option>)}
                    </select>
                  </div>

                  {/* Deeni Kaam Filter */}
                  <div className="flex flex-col flex-1 min-w-[120px]">
                    <label className="block text-[10px] font-bold text-slate-500 uppercase tracking-widest mb-2">Deeni Kaam</label>
                    <select value={selectedDeeniKaam} onChange={(e) => { setSelectedDeeniKaam(e.target.value); setSelectedField(""); }} className="w-full p-2.5 rounded-lg border border-slate-300 bg-white text-sm font-medium text-slate-700 focus:outline-none focus:border-teal-500 focus:ring-1 focus:ring-teal-500 appearance-none cursor-pointer">
                      <option value="">All Deeni Kaam</option>
                      {availableDeeniKaam.map(dk => <option key={dk} value={dk}>{dk}</option>)}
                    </select>
                  </div>

                  {/* Fields Filter */}
                  <div className="flex flex-col flex-1 min-w-[120px]">
                    <label className="block text-[10px] font-bold text-slate-500 uppercase tracking-widest mb-2">Fields</label>
                    <select value={selectedField} onChange={(e) => setSelectedField(e.target.value)} className="w-full p-2.5 rounded-lg border border-slate-300 bg-white text-sm font-medium text-slate-700 focus:outline-none focus:border-teal-500 focus:ring-1 focus:ring-teal-500 appearance-none cursor-pointer">
                      <option value="">All Fields</option>
                      {availableFields.map(fld => <option key={fld} value={fld}>{fld}</option>)}
                    </select>
                  </div>

                  {[
                    { label: "Region", val: region, set: handleSetRegion, opts: uniqValues(rawData, "Region") },
                    { label: "State", val: state, set: handleSetState, opts: uniqValues(rawData.filter(x=>!region||sameClient(x["Region"],region)), "State") },
                    { label: "Division", val: division, set: handleSetDivision, opts: uniqValues(rawData.filter(x=>(!region||sameClient(x["Region"],region))&&(!state||sameClient(x["State"],state))), "Division") },
                    { label: "District", val: district, set: setDistrict, opts: uniqValues(rawData.filter(x=>(!region||sameClient(x["Region"],region))&&(!state||sameClient(x["State"],state))&&(!division||sameClient(x["Division"],division))), "District") }
                  ].map((f, i) => (
                    <div key={i} className="flex flex-col flex-1 min-w-[110px]">
                      <label className="block text-[10px] font-bold text-slate-500 uppercase tracking-widest mb-2">{f.label}</label>
                      <select value={f.val} onChange={(e) => f.set(e.target.value)} disabled={officeUser?.[f.label.toLowerCase()] && officeUser[f.label.toLowerCase()].toLowerCase() !== "all"} className="w-full p-2.5 rounded-lg border border-slate-300 bg-white text-sm font-medium text-slate-700 focus:outline-none focus:border-teal-500 focus:ring-1 focus:ring-teal-500 appearance-none cursor-pointer disabled:bg-slate-50 disabled:text-slate-400">
                        <option value="">{f.val || "All"}</option>
                        {f.opts.map(o => <option key={o} value={o}>{o}</option>)}
                      </select>
                    </div>
                  ))}
                </div>
              </div>

              {/* GRAPHS VIEW MODE */}
              {activeViewMode === "graphs" && (
                <div className="grid grid-cols-1 lg:grid-cols-2 gap-6 animate-fadeIn">
                  <div className="bg-white rounded-xl shadow-sm border border-slate-200 p-6 h-[400px] flex flex-col">
                    <h3 className="text-xs font-bold text-teal-700 uppercase tracking-widest mb-4">Visual Distribution Chart ({dynamicGraphTitle})</h3>
                    <div className="flex-1 flex w-full pt-4">
                      {dynamicGraphData.length > 0 ? (
                        <div className="flex-1 flex justify-start items-end gap-4 pl-4 pb-8 relative border-b border-slate-300 overflow-x-auto custom-scrollbar">
                          {dynamicGraphData.slice(0, 12).map((d, i) => {
                             const h = maxDynamicCount ? (d.count / maxDynamicCount) * 100 : 0;
                             return (
                               <div key={i} className="flex flex-col justify-end items-center relative h-full w-12 shrink-0">
                                  <span className="text-[10px] font-bold text-slate-700 mb-1">{d.count.toLocaleString("en-IN")}</span>
                                  <div style={{height: `${Math.max(h, 2)}%`}} className="w-10 bg-teal-700 rounded-t-md transition-all" />
                                  <span className="absolute -bottom-7 w-24 text-center text-[9px] text-slate-600 truncate">{d.label}</span>
                               </div>
                             )
                          })}
                        </div>
                      ) : <div className="mx-auto my-auto text-slate-400 text-xs">No graph data found</div>}
                    </div>
                  </div>
                  <div className="grid grid-cols-1 md:grid-cols-2 gap-6 h-[400px]">
                    <MiniTable title="CATEGORY WISE METRICS" data={groupCount(processedTableData, "CategoryName")} />
                    <MiniTable title="GEOGRAPHIC WISE METRICS" data={groupCount(processedTableData, "LeftColValue")} />
                  </div>
                </div>
              )}
            </main>
        </div>

        {/* MAIN DATA GRID / TABLE SECTION */}
        <main className="max-w-[1400px] mx-auto px-4 sm:px-6 lg:px-8 pb-8 mt-2">
          <div className="bg-white rounded-xl border border-slate-200 shadow-sm mt-2">
            <div className="p-4 border-b border-slate-200 flex flex-col md:flex-row justify-between items-center gap-4 bg-[#e0f2f1] rounded-t-xl">
              <div>
                <h2 className="text-sm font-bold text-teal-800 uppercase tracking-widest">
                  {activeTab === "Progress Report" ? "Progress & Growth Matrix" : activeTab === "Quarterly Report" ? "Quarterly Comparative Analysis" : activeTab === "Average Report" ? "Average Performance Matrix" : `Detailed Telemetry Output (${activeTab})`}
                </h2>
                <p className="text-[10px] font-bold text-teal-800 mt-1 uppercase tracking-widest">Source Records: {rawData.length} &bull; Visible Rows: {processedTableData.length}</p>
              </div>
            </div>

            <div className="overflow-x-auto w-full custom-scrollbar">
              <table className="min-w-full text-left text-[11px] lg:text-xs">
                <thead className="bg-[#008b8b]">
                  <tr>
                    {!district && (
                      <th className="px-2 py-3 font-bold text-white uppercase tracking-wider border-r border-white/25 align-middle text-center">
                        {leftHeaderTitle}
                      </th>
                    )}
                    <th className="px-2 py-3 font-bold text-white uppercase tracking-wider border-r border-white/25 align-middle text-center">CATEGORY</th>
                    <th className="px-2 py-3 font-bold text-white uppercase tracking-wider border-r border-white/25 align-middle text-center">DEENI KAAM</th>
                    <th className="px-2 py-3 font-bold text-white uppercase tracking-wider border-r border-white/25 align-middle text-center">FIELDS</th>
                    
                    {/* Main Report Column Month Selector */}
                    <th className="px-2 py-3 font-bold text-white uppercase tracking-wider border-r border-white/25 text-center bg-[#008b8b]">
                      <select 
                        value={colMonthMain} 
                        onChange={(e) => setColMonthMain(e.target.value)} 
                        className="bg-teal-900 text-white text-xs font-bold py-1.5 px-3 rounded-lg border border-white/70 outline-none cursor-pointer shadow-sm hover:bg-teal-950 transition-colors"
                      >
                        {availableMonths.map(m => (
                          <option key={m} value={m} className="bg-slate-900 text-white font-bold">{formatMonthYearLabel(m)}</option>
                        ))}
                      </select>
                    </th>

                    {/* Targets Column Selector */}
                    <th className="px-2 py-3 font-bold text-white uppercase tracking-wider border-r border-white/25 text-center bg-[#008b8b]">
                      <div className="flex items-center justify-center gap-2">
                        <span>TARGETS</span>
                        <select 
                          value={headerTargetMode} 
                          onChange={(e) => setHeaderTargetMode(e.target.value)}
                          className="bg-teal-900 text-white text-xs font-bold py-1 px-2 rounded-lg border border-white/70 outline-none cursor-pointer shadow-sm hover:bg-teal-950 transition-colors"
                        >
                          <option value="" className="bg-slate-900 text-white font-bold">Std</option>
                          <option value="26%" className="bg-slate-900 text-white font-bold">26%</option>
                          <option value="52%" className="bg-slate-900 text-white font-bold">52%</option>
                        </select>
                      </div>
                    </th>

                    <th className="px-2 py-3 font-bold text-white uppercase tracking-wider border-r border-white/25 text-center bg-[#008b8b]">ACHIEVEMENT (%)</th>
                    
                    {/* Comparison Column 1 Selector */}
                    <th className="px-2 py-3 font-bold text-white uppercase tracking-wider border-r border-white/25 text-center bg-[#008b8b]">
                      <select 
                        value={colMonth1} 
                        onChange={(e) => setColMonth1(e.target.value)} 
                        className="bg-teal-900 text-white text-xs font-bold py-1.5 px-3 rounded-lg border border-white/70 outline-none cursor-pointer shadow-sm hover:bg-teal-950 transition-colors"
                      >
                        {availableMonths.map(m => (
                          <option key={m} value={m} className="bg-slate-900 text-white font-bold">{formatMonthYearLabel(m)}</option>
                        ))}
                      </select>
                    </th>

                    {/* Comparison Column 2 Selector */}
                    <th className="px-2 py-3 font-bold text-white uppercase tracking-wider border-r border-white/25 text-center bg-[#008b8b]">
                      <select 
                        value={colMonth2} 
                        onChange={(e) => setColMonth2(e.target.value)} 
                        className="bg-teal-900 text-white text-xs font-bold py-1.5 px-3 rounded-lg border border-white/70 outline-none cursor-pointer shadow-sm hover:bg-teal-950 transition-colors"
                      >
                        {availableMonths.map(m => (
                          <option key={m} value={m} className="bg-slate-900 text-white font-bold">{formatMonthYearLabel(m)}</option>
                        ))}
                      </select>
                    </th>

                    <th className="px-2 py-3 font-bold text-white uppercase tracking-wider text-center bg-[#008b8b]">COMPARISON (%)</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {pagedRows.length > 0 ? pagedRows.map((row, idx) => (
                    <tr key={idx} className="hover:bg-slate-50 transition-colors">
                      {!district && (
                        <td className="px-2 py-2 font-bold text-teal-800 leading-tight border-r border-slate-100 text-center">{row?.LeftColValue || "-"}</td>
                      )}
                      
                      <td className="px-2 py-2 font-semibold text-slate-700 leading-tight border-r border-slate-100 text-center">{row?.CategoryName || "-"}</td>
                      <td className="px-2 py-2 font-bold text-slate-800 leading-tight border-r border-slate-100 text-center">{row?.DeeniKaamName || "-"}</td>
                      <td className="px-2 py-2 font-bold text-slate-700 leading-tight border-r border-slate-100 text-center">{row?.DeeniActivity || "-"}</td>
                      
                      <td className="px-2 py-2 text-slate-700 text-center border-r border-slate-100 font-semibold">{row?.DynamicReportValue ?? 0}</td>
                      <td className="px-2 py-2 text-slate-700 text-center border-r border-slate-100 font-semibold">{row?.DynamicTarget ?? "-"}</td>
                      <td className="px-2 py-2 text-blue-600 font-bold text-center border-r border-slate-100">{row?.DynamicAchievement ?? "-"}</td>
                      
                      <td className="px-2 py-2 text-slate-700 text-center border-r border-slate-100 font-semibold">{row?.Val1 ?? 0}</td>
                      <td className="px-2 py-2 text-slate-700 text-center border-r border-slate-100 font-semibold">{row?.Val2 ?? 0}</td>
                      
                      <td className={`px-2 py-2 font-bold text-center ${String(row?.CalculatedComparison || "").startsWith("+") ? "text-emerald-600" : "text-red-600"}`}>
                        {row?.CalculatedComparison || "0.0%"}
                      </td>
                    </tr>
                  )) : (
                    <tr>
                      <td colSpan={district ? 9 : 10} className="px-6 py-12 text-center text-slate-500 text-xs uppercase tracking-widest">
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
            {processedTableData.length > rowsPerPage && (
              <div className="p-4 border-t border-slate-200 bg-[#f8fafc] rounded-b-xl flex justify-center items-center gap-2">
                <button disabled={currentPage===1} onClick={()=>setCurrentPage(p=>p-1)} className="px-3 py-1.5 rounded-lg border border-teal-600 bg-white text-xs font-bold text-teal-700 disabled:opacity-30 shadow-sm transition-all hover:bg-teal-50">PREV</button>
                <span className="text-xs font-bold text-teal-700 px-4">PAGE {currentPage} OF {Math.ceil(processedTableData.length / rowsPerPage)}</span>
                <button disabled={currentPage===Math.ceil(processedTableData.length / rowsPerPage)} onClick={()=>setCurrentPage(p=>p+1)} className="px-3 py-1.5 rounded-lg border border-teal-600 bg-white text-xs font-bold text-teal-700 disabled:opacity-30 shadow-sm transition-all hover:bg-teal-50">NEXT</button>
              </div>
            )}
          </div>
        </main>
      </div>
    </div>
  );
}
