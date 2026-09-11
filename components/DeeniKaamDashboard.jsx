"use client";
import React, { useState, useEffect, useMemo, useTransition } from "react";
import Papa from "papaparse";
import { LogOut, RefreshCw, Filter, Calendar, Search, Activity, ArrowLeft, Download, BookOpenCheck, Image as ImageIcon, Clock, LayoutDashboard, BarChart3, TrendingUp, Target } from "lucide-react";
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
  { category: "Weekly", deeniKaam: "Haftawar Risala ", field: "Haftwar Risala Padhne Wale / Sunne Wale" },
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

const extractGoogleSheetId = (url) => {
  const m = String(url || "").match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  return m ? m[1] : "";
};

// Google Sheets CSV export can be blocked by browser CORS in some deployments.
// JSONP through the Visualization endpoint avoids that browser-side CORS problem
// for sheets that are shared/published for public access.
const parseGoogleSheetJSONP = (url) => new Promise((resolve, reject) => {
  const spreadsheetId = extractGoogleSheetId(url);
  if (!spreadsheetId) {
    reject(new Error("Invalid Google Sheet URL"));
    return;
  }

  const callbackName = `__gs_cb_${Date.now()}_${Math.random().toString(36).slice(2)}`;
  const script = document.createElement("script");
  const timeoutId = window.setTimeout(() => {
    cleanup();
    reject(new Error("Google Sheet request timed out"));
  }, 30000);

  const cleanup = () => {
    window.clearTimeout(timeoutId);
    try { delete window[callbackName]; } catch (_) { window[callbackName] = undefined; }
    if (script.parentNode) script.parentNode.removeChild(script);
  };

  window[callbackName] = (payload) => {
    try {
      if (!payload || payload.status !== "ok" || !payload.table) {
        throw new Error("Google Sheet is not publicly readable");
      }
      const cols = (payload.table.cols || []).map((c, i) => String(c.label || c.id || `Column_${i + 1}`).trim());
      const rows = (payload.table.rows || []).map((r) => {
        const obj = {};
        cols.forEach((key, i) => {
          const cell = r?.c?.[i];
          obj[key] = cell?.f ?? cell?.v ?? "";
        });
        return obj;
      });
      cleanup();
      resolve(rows);
    } catch (e) {
      cleanup();
      reject(e);
    }
  };

  const base = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/gviz/tq`;
  script.src = `${base}?tqx=out:json;responseHandler:${encodeURIComponent(callbackName)}&headers=1`;
  script.async = true;
  script.onerror = () => {
    cleanup();
    reject(new Error("Google Sheet could not be loaded"));
  };
  document.head.appendChild(script);
});

const parseSheetCSVFallback = (url) => new Promise((resolve, reject) => {
  Papa.parse(url, {
    download: true,
    header: true,
    skipEmptyLines: true,
    worker: true,
    complete: (results) => {
      if (results?.errors?.length && !results?.data?.length) {
        reject(new Error("CSV parse failed"));
        return;
      }
      resolve(results?.data || []);
    },
    error: (err) => reject(err)
  });
});

const parseSheet = async (url) => {
  try {
    const rows = await parseGoogleSheetJSONP(url);
    if (rows.length) return rows;
  } catch (jsonpError) {
    console.warn("Google Visualization sync failed; trying CSV fallback.", jsonpError);
  }
  return parseSheetCSVFallback(url);
};

const sameClient = (a, b) =>
  String(a ?? "").trim().toLowerCase() === String(b ?? "").trim().toLowerCase();

const cleanText = (v) => String(v ?? "").trim();

const toNumber = (v) => {
  if (typeof v === "number") return Number.isFinite(v) ? v : 0;
  const n = Number(String(v ?? "0").replace(/,/g, "").replace(/%/g, "").trim());
  return Number.isFinite(n) ? n : 0;
};

const getFirst = (row, keys) => {
  for (const key of keys) {
    if (row && row[key] !== undefined && row[key] !== null && cleanText(row[key]) !== "") {
      return row[key];
    }
  }
  return "";
};

const MONTH_NUMBERS = {
  january: "01", jan: "01", february: "02", feb: "02", march: "03", mar: "03",
  april: "04", apr: "04", may: "05", june: "06", jun: "06", july: "07", jul: "07",
  august: "08", aug: "08", september: "09", sep: "09", sept: "09", october: "10",
  oct: "10", november: "11", nov: "11", december: "12", dec: "12"
};

const normalizeMonth = (row) => {
  const monthRaw = cleanText(getFirst(row, ["Month", "month", "Report Month", "Month_Name"]));
  const yearRaw = cleanText(getFirst(row, ["Year", "year", "Report Year"]));

  if (/^\d{4}-\d{2}$/.test(monthRaw)) return monthRaw;

  const combined = `${monthRaw} ${yearRaw}`.trim();
  const yearMatch = combined.match(/(?:19|20)\d{2}/);
  const year = yearMatch ? yearMatch[0] : "";

  const numericMonth = combined.match(/\b(0?[1-9]|1[0-2])\b/);
  if (year && numericMonth && !/[a-zA-Z]/.test(monthRaw)) {
    return `${year}-${String(numericMonth[1]).padStart(2, "0")}`;
  }

  const lower = combined.toLowerCase();
  for (const name of Object.keys(MONTH_NUMBERS)) {
    if (lower.includes(name)) {
      const y = year || "2026";
      return `${y}-${MONTH_NUMBERS[name]}`;
    }
  }

  const parsed = Date.parse(combined);
  if (!Number.isNaN(parsed)) {
    const d = new Date(parsed);
    return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
  }
  return "";
};

const formatMonthYearLabel = (val) => {
  if (!val) return "Month";
  const [year, month] = String(val).split("-");
  if (!year || !month) return val;
  const date = new Date(Number(year), Number(month) - 1, 1);
  return Number.isNaN(date.getTime()) ? val :
    date.toLocaleString("en-US", { month: "short", year: "numeric" });
};

const getPreviousMonth = (ym) => {
  if (!ym || !/^\d{4}-\d{2}$/.test(ym)) return "";
  const [y, m] = ym.split("-").map(Number);
  const d = new Date(y, m - 2, 1);
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
};

const aggregateRows = (rows) => {
  const map = new Map();

  for (const row of rows || []) {
    if (!row) continue;

    const field = cleanText(getFirst(row, ["Fields", "Fileds", "Field", "Deeni Activities"]));
    if (!field) continue;

    const master = MASTER_CATEGORIES_MAP.find(m => sameClient(m.field, field));
    const category = cleanText(getFirst(row, ["Category"])) || master?.category || "Basic";
    const deeniKaam = cleanText(getFirst(row, ["Deeni Kaam", "DeeniKaam"])) || master?.deeniKaam || "";

    const normalized = {
      Month: normalizeMonth(row),
      Chain: cleanText(getFirst(row, ["Chain"])),
      Department: cleanText(getFirst(row, ["Department"])),
      Region: cleanText(getFirst(row, ["Region"])),
      State: cleanText(getFirst(row, ["State"])),
      Division: cleanText(getFirst(row, ["Division"])),
      District: cleanText(getFirst(row, ["District"])),
      Category: category,
      "Deeni Kaam": deeniKaam,
      Fields: field,
      "Multiple Field Name": cleanText(getFirst(row, ["Multiple Field Name", "Multiple_Field_Name"])),
      "Multiple Field Value": cleanText(getFirst(row, ["Multiple Field Value", "Multiple_Field_Value"])),
      Report: toNumber(getFirst(row, ["Report Value", "Report", "report"])),
      Target26: toNumber(getFirst(row, ["Target 26% (Value)", "Target 26%", "Target26", "Target_26Pct"])),
      Target52: toNumber(getFirst(row, ["Target 52% (Value)", "Target 52%", "Target52", "Target_52Pct"])),
      Source: cleanText(getFirst(row, ["Source Report", "Source_File"])),
      Notes: cleanText(getFirst(row, ["Notes"]))
    };

    if (!normalized.Month) continue;

    const key = [
      normalized.Month, normalized.Chain, normalized.Department,
      normalized.Region, normalized.State, normalized.Division, normalized.District,
      normalized.Category, normalized["Deeni Kaam"], normalized.Fields,
      normalized["Multiple Field Name"], normalized["Multiple Field Value"]
    ].map(cleanText).join("\u001F");

    const old = map.get(key);
    if (old) {
      old.Report += normalized.Report;
      old.Target26 += normalized.Target26;
      old.Target52 += normalized.Target52;
      if (!old.Source && normalized.Source) old.Source = normalized.Source;
    } else {
      map.set(key, { ...normalized });
    }
  }
  return Array.from(map.values());
};

const groupCount = (data = [], key) => {
  const map = new Map();
  for (const x of data) {
    const k = cleanText(x?.[key]) || "Unknown";
    map.set(k, (map.get(k) || 0) + toNumber(x?.DynamicReportValue ?? x?.Report));
  }
  return Array.from(map, ([label, count]) => ({ label, count }))
    .sort((a, b) => b.count - a.count);
};

const uniqValues = (data = [], key) =>
  [...new Set(data.map(x => cleanText(x?.[key])).filter(Boolean))]
    .sort((a, b) => a.localeCompare(b));

const MiniTable = ({ title, data = [] }) => (
  <div className="bg-white rounded-xl shadow-sm border border-slate-200 overflow-hidden flex flex-col h-[340px]">
    <div className="p-4 border-b border-slate-100">
      <h3 className="text-teal-700 font-bold uppercase tracking-widest text-xs">{title}</h3>
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
              <td className="py-2.5 px-4 text-slate-700 text-center">{d?.label || "-"}</td>
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
  const [startMonth, setStartMonth] = useState("");
  const [endMonth, setEndMonth] = useState("");

  const [region, setRegion] = useState(officeUser?.region && officeUser.region.toLowerCase() !== "all" ? officeUser.region : "");
  const [state, setState] = useState(officeUser?.state && officeUser.state.toLowerCase() !== "all" ? officeUser.state : "");
  const [division, setDivision] = useState(officeUser?.division && officeUser.division.toLowerCase() !== "all" ? officeUser.division : "");
  const [district, setDistrict] = useState(officeUser?.district && officeUser.district.toLowerCase() !== "all" ? officeUser.district : "");

  const [selectedCategory, setSelectedCategory] = useState("");
  const [selectedDeeniKaam, setSelectedDeeniKaam] = useState("");
  const [selectedField, setSelectedField] = useState("");
  const [selectedTargetPct, setSelectedTargetPct] = useState("");

  const [searchTerm, setSearchTerm] = useState("");
  const [currentPage, setCurrentPage] = useState(1);
  const rowsPerPage = 10;

  const [now, setNow] = useState(null);
  useEffect(() => {
    setNow(new Date());
    const timer = setInterval(() => setNow(new Date()), 1000);
    return () => clearInterval(timer);
  }, []);
  const dateStr = now ? now.toLocaleDateString("en-GB", { day: "2-digit", month: "2-digit", year: "numeric" }).replace(/\//g, "-") : "";
  const timeStr = now ? now.toLocaleTimeString("en-US", { hour: "2-digit", minute: "2-digit", second: "2-digit", hour12: true }) : "";

  useEffect(() => { fetchData(); }, []);

  const fetchData = async () => {
    setLoading(true);
    setFetchError("");
    try {
      const allRows = [];
      const failedSheets = [];

      // Keep downloads sequential: large sheets should not all sit in memory together.
      for (let i = 0; i < SHEET_URLS.length; i += 1) {
        const url = SHEET_URLS[i];
        try {
          const rows = await parseSheet(url);
          if (Array.isArray(rows) && rows.length) {
            allRows.push(...rows);
          } else {
            failedSheets.push(i + 1);
          }
        } catch (sheetError) {
          console.error(`Google Sheet ${i + 1} sync failed`, sheetError);
          failedSheets.push(i + 1);
        }
      }

      const compact = aggregateRows(allRows);
      if (compact.length) {
        setRawData(compact);
        setCurrentPage(1);
        if (failedSheets.length) {
          setFetchError(`Data synced from available Google Sheets. Sheet ${failedSheets.join(", ")} could not be read; please make those sheets "Anyone with the link" / Published and try Sync again.`);
        }
      } else {
        setRawData([]);
        setFetchError("Google Sheet se data read nahi ho raha. Google Sheet ko 'Anyone with the link → Viewer' ya 'Publish to web' karein, phir Sync dabayein.");
      }
    } catch (err) {
      console.error(err);
      setRawData([]);
      setFetchError("Google Sheet sync failed. Please make the sheet publicly readable and verify the Google Sheet URL.");
    } finally {
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

  const availableCategories = useMemo(() =>
    [...new Set(MASTER_CATEGORIES_MAP.map(m => m.category))], []);

  const availableDeeniKaam = useMemo(() => {
    let subset = MASTER_CATEGORIES_MAP;
    if (selectedCategory) subset = subset.filter(m => sameClient(m.category, selectedCategory));
    return [...new Set(subset.map(m => m.deeniKaam).filter(Boolean))]
      .sort((a,b) => a.localeCompare(b));
  }, [selectedCategory]);

  const availableFields = useMemo(() => {
    let subset = MASTER_CATEGORIES_MAP;
    if (selectedCategory) subset = subset.filter(m => sameClient(m.category, selectedCategory));
    if (selectedDeeniKaam) subset = subset.filter(m => sameClient(m.deeniKaam, selectedDeeniKaam));
    return [...new Set(subset.map(m => m.field).filter(Boolean))]
      .sort((a,b) => a.localeCompare(b));
  }, [selectedCategory, selectedDeeniKaam]);

  const leftHeaderTitle = useMemo(() => {
    if (district) return "";
    if (division) return "DISTRICT";
    if (state) return "DIVISION";
    if (region) return "STATE";
    return "COUNTRY";
  }, [region, state, division, district]);

  const selectedPeriod = endMonth || startMonth || "";
  const comparisonPeriod = startMonth && endMonth && startMonth !== endMonth
    ? startMonth
    : getPreviousMonth(selectedPeriod);

  const baseMatches = (row) => {
    if (region && !sameClient(row.Region, region)) return false;
    if (state && !sameClient(row.State, state)) return false;
    if (division && !sameClient(row.Division, division)) return false;
    if (district && !sameClient(row.District, district)) return false;
    if (selectedCategory && !sameClient(row.Category, selectedCategory)) return false;
    if (selectedDeeniKaam && !sameClient(row["Deeni Kaam"], selectedDeeniKaam)) return false;
    if (selectedField && !sameClient(row.Fields, selectedField)) return false;
    if (searchTerm) {
      const q = searchTerm.toLowerCase().trim();
      const hay = [row.Region,row.State,row.Division,row.District,row.Category,row["Deeni Kaam"],
        row.Fields,row["Multiple Field Name"],row["Multiple Field Value"]].join(" ").toLowerCase();
      if (!hay.includes(q)) return false;
    }
    return true;
  };

  const aggregatePeriod = (period, groupValue, groupLevel, field) => {
    if (!period) return { report: 0, target26: 0, target52: 0 };
    let report = 0, target26 = 0, target52 = 0;

    for (const r of rawData) {
      if (r.Month !== period || !baseMatches(r)) continue;
      if (field && !sameClient(r.Fields, field)) continue;

      if (groupLevel === "State" && !sameClient(r.State, groupValue)) continue;
      if (groupLevel === "Division" && !sameClient(r.Division, groupValue)) continue;
      if (groupLevel === "District" && !sameClient(r.District, groupValue)) continue;

      report += r.Report;
      target26 += r.Target26;
      target52 += r.Target52;
    }
    return { report, target26, target52 };
  };

  const processedTableData = useMemo(() => {
    if (!rawData.length) return [];

    let groupLevel = "India";
    let groups = [{ value: "India", level: "India" }];

    if (district) {
      groupLevel = "District";
      groups = [{ value: district, level: "District" }];
    } else if (division) {
      groupLevel = "District";
      groups = uniqValues(rawData.filter(r =>
        sameClient(r.Region, region) &&
        sameClient(r.State, state) &&
        sameClient(r.Division, division)
      ), "District").map(value => ({ value, level: "District" }));
    } else if (state) {
      groupLevel = "Division";
      groups = uniqValues(rawData.filter(r =>
        sameClient(r.Region, region) && sameClient(r.State, state)
      ), "Division").map(value => ({ value, level: "Division" }));
    } else if (region) {
      groupLevel = "State";
      groups = uniqValues(rawData.filter(r => sameClient(r.Region, region)), "State")
        .map(value => ({ value, level: "State" }));
    }

    const fields = selectedField
      ? [selectedField]
      : [...new Set(rawData.filter(baseMatches).map(r => r.Fields).filter(Boolean))]
          .sort((a,b) => a.localeCompare(b));

    let targetPeriod = selectedPeriod;
    if (!targetPeriod) {
      targetPeriod = rawData.reduce((latest, r) => r.Month > latest ? r.Month : latest, "");
    }
    const comparePeriod = comparisonPeriod || getPreviousMonth(targetPeriod);

    const output = [];

    for (const g of groups) {
      for (const field of fields) {
        const current = aggregatePeriod(targetPeriod, g.value, g.level, field);
        const previous = aggregatePeriod(comparePeriod, g.value, g.level, field);

        const target = selectedTargetPct === "26%" ? current.target26 :
          selectedTargetPct === "52%" ? current.target52 :
          0;

        const comparison = previous.report > 0
          ? ((current.report - previous.report) / previous.report) * 100
          : (current.report > 0 ? 100 : 0);

        output.push({
          LeftColValue: district ? "" : g.value,
          DeeniKaamName: MASTER_CATEGORIES_MAP.find(m => sameClient(m.field, field))?.deeniKaam || "",
          DeeniActivity: field,
          DynamicMonthDisplay: formatMonthYearLabel(targetPeriod),
          DynamicReportValue: current.report,
          DynamicTarget: target,
          DynamicAchievement: target > 0 ? `${((current.report / target) * 100).toFixed(1)}%` : "-",
          Val1: previous.report,
          Val2: current.report,
          CalculatedComparison: `${comparison >= 0 ? "+" : ""}${comparison.toFixed(1)}%`,
          State: groupLevel === "State" ? g.value : "",
          Division: groupLevel === "Division" ? g.value : "",
          District: groupLevel === "District" ? g.value : ""
        });
      }
    }
    return output;
  }, [
    rawData, region, state, division, district, selectedCategory, selectedDeeniKaam,
    selectedField, selectedTargetPct, searchTerm, startMonth, endMonth
  ]);

  const pagedRows = useMemo(() => {
    const start = (currentPage - 1) * rowsPerPage;
    return processedTableData.slice(start, start + rowsPerPage);
  }, [processedTableData, currentPage]);

  useEffect(() => { setCurrentPage(1); }, [
    region,state,division,district,selectedCategory,selectedDeeniKaam,selectedField,
    selectedTargetPct,startMonth,endMonth,searchTerm
  ]);

  const dynamicGraphData = useMemo(() => {
    const map = new Map();
    for (const x of processedTableData) {
      const k = district ? x.DeeniActivity : (x.LeftColValue || "India");
      map.set(k, (map.get(k) || 0) + toNumber(x.DynamicReportValue));
    }
    return Array.from(map, ([label,count]) => ({label,count}))
      .sort((a,b) => b.count-a.count);
  }, [processedTableData, district]);

  const dynamicGraphTitle = division ? "REPORTS BY DISTRICT" :
    state ? "REPORTS BY DIVISION" :
    region ? "REPORTS BY STATE" : "INDIA TOTAL REPORT";

  const maxDynamicCount = dynamicGraphData.length
    ? Math.max(...dynamicGraphData.map(d => d.count)) : 0;

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
                    <p className="text-[10px] font-bold text-teal-700 uppercase tracking-widest mt-0.5">Professional Multi-View Dashboard</p>
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
                  {["Monthly Report", "Average Report", "Quarterly Report"].map((tab) => (
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

              <div className="bg-white p-5 rounded-xl border border-slate-200 shadow-sm">
                <div className="flex flex-col lg:flex-row justify-between lg:items-center gap-4 mb-5 border-b border-slate-100 pb-4">
                  <div className="flex items-center gap-3">
                    <Filter className="w-5 h-5 text-teal-700" />
                    <h2 className="text-sm font-bold text-slate-800 tracking-widest uppercase">Global Filters & Range Picker</h2>
                  </div>
                  
                  <div className="flex flex-wrap items-center gap-3" data-html2canvas-ignore="true">
                     <div className="flex items-center gap-2 bg-slate-50 p-1.5 rounded-lg border border-slate-200">
                        <Calendar className="w-4 h-4 text-slate-400 ml-2" />
                        <input type="month" value={startMonth} onChange={e => setStartMonth(e.target.value)} className="bg-transparent text-xs font-bold text-slate-700 outline-none" />
                        <span className="text-slate-400 text-xs font-bold">TO</span>
                        <input type="month" value={endMonth} onChange={e => setEndMonth(e.target.value)} className="bg-transparent text-xs font-bold text-slate-700 outline-none pr-2" />
                     </div>
                     {(startMonth || endMonth) && (
                       <button onClick={() => { setStartMonth(""); setEndMonth(""); }} className="text-xs text-red-500 font-bold underline px-2">Reset</button>
                     )}
                  </div>
                </div>

                <div className="grid grid-cols-2 md:grid-cols-9 gap-3">
                  <div className="col-span-2 md:col-span-1" data-html2canvas-ignore="true">
                    <label className="block text-[10px] font-bold text-slate-500 uppercase tracking-widest mb-2">Search</label>
                    <div className="relative">
                      <span className="absolute inset-y-0 left-0 pl-3 flex items-center text-slate-400"><Search className="w-4 h-4" /></span>
                      <input type="text" placeholder="Search..." value={searchTerm} onChange={(e) => setSearchTerm(e.target.value)} className="w-full pl-9 pr-3 py-2.5 rounded-lg border border-slate-300 bg-white text-sm font-medium text-slate-700 focus:outline-none focus:border-teal-500 focus:ring-1 focus:ring-teal-500" />
                    </div>
                  </div>

                  {/* Category Filter */}
                  <div className="flex flex-col flex-1 min-w-[120px]">
                    <label className="block text-[10px] font-bold text-slate-500 uppercase tracking-widest mb-2">Category</label>
                    <select value={selectedCategory} onChange={(e) => { setSelectedCategory(e.target.value); setSelectedDeeniKaam(""); setSelectedField(""); }} className="w-full p-2.5 rounded-lg border border-slate-300 bg-white text-sm font-medium text-slate-700 focus:outline-none focus:border-teal-500 focus:ring-1 focus:ring-teal-500 appearance-none cursor-pointer">
                      <option value="">All Categories</option>
                      {availableCategories.map(cat => <option key={cat} value={cat}>{cat}</option>)}
                    </select>
                  </div>

                  {/* Deeni Kaam Filter */}
                  <div className="flex flex-col flex-1 min-w-[130px]">
                    <label className="block text-[10px] font-bold text-slate-500 uppercase tracking-widest mb-2">Deeni Kaam</label>
                    <select value={selectedDeeniKaam} onChange={(e) => { setSelectedDeeniKaam(e.target.value); setSelectedField(""); }} className="w-full p-2.5 rounded-lg border border-slate-300 bg-white text-sm font-medium text-slate-700 focus:outline-none focus:border-teal-500 focus:ring-1 focus:ring-teal-500 appearance-none cursor-pointer">
                      <option value="">All Deeni Kaam</option>
                      {availableDeeniKaam.map(dk => <option key={dk} value={dk}>{dk}</option>)}
                    </select>
                  </div>

                  {/* Fields Filter */}
                  <div className="flex flex-col flex-1 min-w-[130px]">
                    <label className="block text-[10px] font-bold text-slate-500 uppercase tracking-widest mb-2">Fields</label>
                    <select value={selectedField} onChange={(e) => setSelectedField(e.target.value)} className="w-full p-2.5 rounded-lg border border-slate-300 bg-white text-sm font-medium text-slate-700 focus:outline-none focus:border-teal-500 focus:ring-1 focus:ring-teal-500 appearance-none cursor-pointer">
                      <option value="">All Fields</option>
                      {availableFields.map(fld => <option key={fld} value={fld}>{fld}</option>)}
                    </select>
                  </div>

                  {/* Targets Filter (26% / 52%) */}
                  <div className="flex flex-col flex-1 min-w-[90px]">
                    <label className="block text-[10px] font-bold text-slate-500 uppercase tracking-widest mb-2">Targets</label>
                    <select value={selectedTargetPct} onChange={(e) => setSelectedTargetPct(e.target.value)} className="w-full p-2.5 rounded-lg border border-slate-300 bg-white text-sm font-medium text-slate-700 focus:outline-none focus:border-teal-500 focus:ring-1 focus:ring-teal-500 appearance-none cursor-pointer">
                      <option value="">Standard</option>
                      <option value="26%">26%</option>
                      <option value="52%">52%</option>
                    </select>
                  </div>

                  {[
                    { label: "Region", val: region, set: handleSetRegion, opts: uniqValues(rawData, "Region") },
                    { label: "State", val: state, set: handleSetState, opts: uniqValues(rawData.filter(x=>!region||sameClient(x["Region"],region)), "State") },
                    { label: "Division", val: division, set: handleSetDivision, opts: uniqValues(rawData.filter(x=>(!region||sameClient(x["Region"],region))&&(!state||sameClient(x["State"],state))), "Division") },
                    { label: "District", val: district, set: setDistrict, opts: uniqValues(rawData.filter(x=>(!region||sameClient(x["Region"],region))&&(!state||sameClient(x["State"],state))&&(!division||sameClient(x["Division"],division))), "District") }
                  ].map((f, i) => (
                    <div key={i} className="flex flex-col flex-1 min-w-[120px]">
                      <label className="block text-[10px] font-bold text-slate-500 uppercase tracking-widest mb-2">{f.label}</label>
                      <select value={f.val} onChange={(e) => f.set(e.target.value)} disabled={officeUser?.[f.label.toLowerCase()] && officeUser[f.label.toLowerCase()].toLowerCase() !== "all"} className="w-full p-2.5 rounded-lg border border-slate-300 bg-white text-sm font-medium text-slate-700 focus:outline-none focus:border-teal-500 focus:ring-1 focus:ring-teal-500 appearance-none cursor-pointer disabled:bg-slate-50 disabled:text-slate-400">
                        <option value="">{f.val || "All"}</option>
                        {f.opts.map(o => <option key={o} value={o}>{o}</option>)}
                      </select>
                    </div>
                  ))}
                </div>
              </div>

              {/* GRAPHS VIEW */}
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
                    <MiniTable title="STATE WISE METRICS" data={groupCount(processedTableData, "LeftColValue")} />
                    <MiniTable title="DISTRICT WISE METRICS" data={groupCount(processedTableData, "District")} />
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
                  {activeViewMode === "average" ? "Average & Range Comparison Matrix" : activeViewMode === "targets" ? "Targets & Achievement Matrix" : `Detailed Telemetry Output (${activeTab})`}
                </h2>
                <p className="text-[10px] font-bold text-teal-800 mt-1 uppercase tracking-widest">Source: {rawData.length} &bull; Visible: {processedTableData.length}</p>
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
                    <th className="px-2 py-3 font-bold text-white uppercase tracking-wider border-r border-white/25 align-middle text-center">DEENI KAAM</th>
                    <th className="px-2 py-3 font-bold text-white uppercase tracking-wider border-r border-white/25 align-middle text-center">FIELDS</th>
                    <th className="px-2 py-3 font-bold text-white uppercase tracking-wider border-r border-white/25 text-center bg-[#007a7a]">
                      {formatMonthYearLabel(endMonth || startMonth || "Month")}
                    </th>
                    <th className="px-2 py-3 font-bold text-white uppercase tracking-wider border-r border-white/25 text-center bg-[#007a7a]">TARGETS</th>
                    <th className="px-2 py-3 font-bold text-white uppercase tracking-wider border-r border-white/25 text-center bg-[#007a7a]">ACHIEVEMENT (%)</th>
                    <th className="px-2 py-3 font-bold text-white uppercase tracking-wider border-r border-white/25 text-center bg-[#007a7a]">
                      {formatMonthYearLabel(startMonth)}
                    </th>
                    <th className="px-2 py-3 font-bold text-white uppercase tracking-wider border-r border-white/25 text-center bg-[#007a7a]">
                      {formatMonthYearLabel(endMonth)}
                    </th>
                    <th className="px-2 py-3 font-bold text-white uppercase tracking-wider text-center bg-[#007a7a]">COMPARISON (%)</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {pagedRows.length > 0 ? pagedRows.map((row, idx) => (
                    <tr key={idx} className="hover:bg-slate-50 transition-colors">
                      {!district && (
                        <td className="px-2 py-2 font-bold text-teal-800 leading-tight border-r border-slate-100 text-center">{row?.LeftColValue || "-"}</td>
                      )}
                      
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
                      <td colSpan={district ? 8 : 9} className="px-6 py-12 text-center text-slate-500 text-xs uppercase tracking-widest">
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
