"use client";

import React, { useEffect, useMemo, useState, useTransition } from "react";
import Papa from "papaparse";
import {
  LogOut, RefreshCw, Filter, Calendar, Search, Activity, ArrowLeft,
  Download, BookOpenCheck, Image as ImageIcon, Clock, LayoutDashboard,
  BarChart3, TrendingUp, Target
} from "lucide-react";
import * as XLSX from "xlsx";
import pptxgen from "pptxgenjs";
import html2canvas from "html2canvas";

/*
  DATA SOURCE
  The dashboard reads the public Google Sheets URLs below.
  The sheets must be shared as "Anyone with the link -> Viewer"
  or published to the web.
*/
const SHEET_URLS = [
  "https://docs.google.com/spreadsheets/d/1P1Ul-jXOfFfhuQLTKeQ-zOynKnCmywH-_gjZ9nJ8tO0/export?format=csv",
  "https://docs.google.com/spreadsheets/d/1S2tEIyaN8p-yu4Vd_GVumqBqwzgTzM3zlms4DwCJr00/export?format=csv",
  "https://docs.google.com/spreadsheets/d/1yWVgL9IVGrQFElLNeO8X_UGAIDSAGF7P8M31gGtoki8/export?format=csv"
];

const FIELD_HIERARCHY = [
  ["Basic","Tanzimi Malumat","Total Active Muballigh"],
  ["Basic","Tanzimi Malumat","Total Moallimin"],
  ["Basic","Tanzimi Malumat","Total Masjid"],
  ["Basic","Tanzimi Malumat","Apni Masjid"],
  ["Basic","Tanzimi Malumat","Total Zeili Halqe"],
  ["Basic","Tanzimi Malumat","Total Zaili Halqe Taqarrur"],
  ["Daily","Fajr Ke Liye Jagaen","Fajr Ke Liye Jagaen"],
  ["Daily","Tafseer Sunna/Sunnana","Tafseer Sunna/Sunnana"],
  ["Daily","Dars","Masjid Dars"],
  ["Daily","Dars","Area Dars"],
  ["Daily","Dars","Ghar Dars"],
  ["Daily","Dars","Total Dars"],
  ["Daily","Madrasatul Madina Baligan","Madrasatul Madina Baligan Tadad (Masjid)"],
  ["Daily","Madrasatul Madina Baligan","Madrasatul Madina Baligan Shurqa (Masjid)"],
  ["Daily","Madrasatul Madina Baligan","Madrasatul Madina Baligan Tadad (Others)"],
  ["Daily","Madrasatul Madina Baligan","Madrasatul Madina Baligan Shurqa (Others)"],
  ["Daily","Madrasatul Madina Baligan","Total Madrasarul Madina Baligan Tadad"],
  ["Daily","Madrasatul Madina Baligan","Total Madrasarul Madina Baligan Shurqa"],
  ["Weekly","Haftwar Ijtima","Haftwar Ijtima Tadad"],
  ["Weekly","Haftwar Ijtima","Haftwar Ijtima Shurqa"],
  ["Weekly","Haftwar Ijtima","Raat Guzarne Walo Ki Tadad"],
  ["Weekly","Haftwar Ijtima","Ijtima me Shurqa ki tadad 120 se zyada hai"],
  ["Weekly","Madani Muzakirah","Madani Muzakirah Maqaamat"],
  ["Weekly","Madani Muzakirah","Madani Muzakirah Shurqa"],
  ["Weekly","Ek Din Raahe Khuda Me","Ek Din Raahe Khuda Me Tadad"],
  ["Weekly","Ek Din Raahe Khuda Me","Ek Din Raahe Khuda Me Shurqa"],
  ["Weekly","Madani Halqa","Madani Halqa Tadad"],
  ["Weekly","Madani Halqa","Madani Halqa Shurqa"],
  ["Weekly","Haftawar Risala ","Haftawar Risala Padhne Wale / Sunne Wale"],
  ["Weekly","Alaqai Dora","Alaqai Dora Kitni Baar"],
  ["Monthly","Qafila","3 Din Qafila Tadad"],
  ["Monthly","Qafila","3 Din Qafila Shurqa"],
  ["Monthly","Qafila","12 Din Qafila Tadad"],
  ["Monthly","Qafila","12 Din Qafila Shurqa"],
  ["Monthly","Qafila","1 Maah Qafila Tadad"],
  ["Monthly","Qafila","1 Maah Qafila Shurqa"],
  ["Monthly","Qafila","12 Maah Qafila Tadad"],
  ["Monthly","Qafila","12 Maah Qafila Shurqa"],
  ["Monthly","Courses","Short Courses (Gair Riyaeshi) Tadad"],
  ["Monthly","Courses","Short Courses (Gair Riyaeshi) Shurqa"],
  ["Monthly","Courses","Long Courses (Riyaeshi) Tadad"],
  ["Monthly","Courses","Long Courses (Riyaeshi) Shurqa"],
  ["Monthly","Courses","Total Courses Tadad"],
  ["Monthly","Courses","Total Courses Shurqa"],
  ["Monthly","Neak Aamal","Neak Aamal Risala Wasool"]
];

const FIELD_LOOKUP = new Map(FIELD_HIERARCHY.map(([category, deeni, field]) => [
  field.trim().toLowerCase(),
  { category, deeniKaam: deeni, field }
]));

const MONTH_NAMES = [
  "January","February","March","April","May","June",
  "July","August","September","October","November","December"
];
const MONTH_MAP = Object.fromEntries(MONTH_NAMES.map((m,i)=>[
  m.toLowerCase(), String(i+1).padStart(2,"0")
]));
const MONTH_SHORT_MAP = Object.fromEntries(MONTH_NAMES.map((m,i)=>[
  m.slice(0,3).toLowerCase(), String(i+1).padStart(2,"0")
]));

const norm = (v) => String(v ?? "").trim();
const same = (a,b) => norm(a).toLowerCase() === norm(b).toLowerCase();

const num = (v) => {
  if (typeof v === "number") return Number.isFinite(v) ? v : 0;
  const s = norm(v).replace(/,/g,"").replace(/%$/,"");
  if (!s || s === "-" || s.toLowerCase() === "na") return 0;
  const n = Number(s);
  return Number.isFinite(n) ? n : 0;
};

const uniq = (arr) => [...new Set(arr.filter(Boolean))];

const uniqSorted = (arr) =>
  uniq(arr.map(norm)).sort((a,b)=>a.localeCompare(b,undefined,{numeric:true,sensitivity:"base"}));

const fieldInfo = (fieldName, category) => {
  const hit = FIELD_LOOKUP.get(norm(fieldName).toLowerCase());
  return hit || {
    category: norm(category) || "Basic",
    deeniKaam: "",
    field: norm(fieldName)
  };
};

function normalizeMonth(row) {
  const rawMonth = norm(row?.Month || row?.["Report Month"] || row?.["Month Name"]);
  const rawYear = norm(row?.Year || row?.["Report Year"]);
  const dateRaw = norm(row?.["Report Date"] || row?.Date);

  if (rawMonth && /^\d{4}-\d{1,2}$/.test(rawMonth)) {
    const [y,m] = rawMonth.split("-");
    return `${y}-${String(m).padStart(2,"0")}`;
  }

  const monthNumber =
    /^\d{1,2}$/.test(rawMonth) ? Number(rawMonth) :
    MONTH_MAP[rawMonth.toLowerCase()] ? Number(MONTH_MAP[rawMonth.toLowerCase()]) :
    MONTH_SHORT_MAP[rawMonth.slice(0,3).toLowerCase()] ? Number(MONTH_SHORT_MAP[rawMonth.slice(0,3).toLowerCase()]) :
    null;

  if (monthNumber && /^\d{4}$/.test(rawYear)) {
    return `${rawYear}-${String(monthNumber).padStart(2,"0")}`;
  }

  const combined = rawMonth || dateRaw;
  const parsed = Date.parse(combined);
  if (!Number.isNaN(parsed)) {
    const d = new Date(parsed);
    const y = d.getFullYear();
    const m = String(d.getMonth()+1).padStart(2,"0");
    return `${y}-${m}`;
  }

  const yearMatch = combined.match(/\b(20\d{2})\b/);
  const year = yearMatch?.[1] || rawYear;
  if (year && monthNumber) return `${year}-${String(monthNumber).padStart(2,"0")}`;

  return "";
}

function formatMonth(val) {
  if (!val || !/^\d{4}-\d{2}$/.test(val)) return "Month";
  const [y,m] = val.split("-").map(Number);
  return new Date(y,m-1,1).toLocaleString("en-US",{month:"short",year:"numeric"});
}

function monthSort(a,b) { return String(a).localeCompare(String(b)); }

function quarterKeyForMonth(m) {
  return m ? m.slice(0,4)+"-Q"+Math.ceil(Number(m.slice(5,7))/3) : "";
}

function parseGviz(text) {
  const start = text.indexOf("{");
  const end = text.lastIndexOf("}");
  if (start < 0 || end <= start) return [];
  const json = JSON.parse(text.slice(start,end+1));
  const cols = (json.table?.cols || []).map(c => c.label || c.id || "");
  return (json.table?.rows || []).map(r => {
    const obj = {};
    (r.c || []).forEach((cell,i) => {
      obj[cols[i] || `Column${i+1}`] = cell?.v ?? "";
    });
    return obj;
  });
}

async function fetchText(url) {
  const res = await fetch(url, {
    method:"GET",
    mode:"cors",
    cache:"no-store",
    credentials:"omit"
  });
  if (!res.ok) throw new Error(`HTTP ${res.status}`);
  return res.text();
}

function makeGvizUrl(exportUrl) {
  const match = exportUrl.match(/\/spreadsheets\/d\/([^/]+)/);
  if (!match) return null;
  return `https://docs.google.com/spreadsheets/d/${match[1]}/gviz/tq?tqx=out:json`;
}

async function loadOneSheet(url) {
  const gvizUrl = makeGvizUrl(url);
  const attempts = [
    async () => {
      if (!gvizUrl) throw new Error("Invalid Google Sheet URL");
      return parseGviz(await fetchText(gvizUrl));
    },
    async () => new Promise((resolve,reject) => {
      Papa.parse(url, {
        download:true,
        header:true,
        skipEmptyLines:true,
        worker:true,
        complete:r=>resolve(r?.data || []),
        error:e=>reject(e)
      });
    })
  ];

  let lastError = null;
  for (const attempt of attempts) {
    try {
      const data = await attempt();
      if (Array.isArray(data) && data.length) return data;
    } catch (e) {
      lastError = e;
    }
  }
  throw lastError || new Error("No data returned");
}

/*
  IMPORTANT PERFORMANCE DESIGN:
  Raw Google Sheet rows are NOT kept in React state.
  Every source row is converted once into a compact aggregate record.
  This prevents hundreds of thousands of rows from being repeatedly scanned
  during every filter change.
*/
function aggregateSources(sourceArrays) {
  const map = new Map();

  for (const rows of sourceArrays) {
    for (const row of rows || []) {
      if (!row) continue;

      const month = normalizeMonth(row);
      if (!month) continue;

      const rawField = norm(row.Fields || row.Fileds || row["Deeni Activities"] || row["Field"]);
      if (!rawField) continue;

      const info = fieldInfo(rawField, row.Category);
      const category = norm(row.Category) || info.category || "Basic";
      const deeniKaam = norm(row["Deeni Kaam"]) || info.deeniKaam || "";

      const region = norm(row.Region);
      const state = norm(row.State);
      const division = norm(row.Division);
      const district = norm(row.District);
      const chain = norm(row.Chain);
      const department = norm(row.Department);
      const multiName = norm(row["Multiple Field Name"] || row["Multiple Fields Name"]);
      const multiValue = norm(row["Multiple Field Value"] || row["Multiple Fields Value"]);

      const report = num(row["Report Value"] ?? row.report);
      let target26 = num(row["Target 26% (Value)"] ?? row["Target 26%"] ?? row["Target 26"] ?? row["26% Target"]);
      let target52 = num(row["Target 52% (Value)"] ?? row["Target 52%"] ?? row["Target 52"] ?? row["52% Target"]);
      if (target26 === 0 && target52 === 0 && row["Target"] != null) target52 = num(row["Target"]);

      const key = [
        month, region, state, division, district, chain, department,
        category, deeniKaam, rawField, multiName, multiValue
      ].map(norm).join("\u001F");

      const old = map.get(key);
      if (old) {
        old.report += report;
        old.target26 += target26;
        old.target52 += target52;
      } else {
        map.set(key, {
          month, year: Number(month.slice(0,4)),
          quarter: `Q${Math.ceil(Number(month.slice(5,7))/3)}`,
          chain, department, region, state, division, district,
          category, deeniKaam, field: rawField,
          multiName, multiValue,
          report, target26, target52
        });
      }
    }
  }

  return [...map.values()];
}

function filterRows(rows, filters) {
  return rows.filter(r => {
    if (filters.region && !same(r.region, filters.region)) return false;
    if (filters.state && !same(r.state, filters.state)) return false;
    if (filters.division && !same(r.division, filters.division)) return false;
    if (filters.district && !same(r.district, filters.district)) return false;
    if (filters.category && !same(r.category, filters.category)) return false;
    if (filters.deeniKaam && !same(r.deeniKaam, filters.deeniKaam)) return false;
    if (filters.field && !same(r.field, filters.field)) return false;
    if (filters.chain && !same(r.chain, filters.chain)) return false;
    if (filters.department && !same(r.department, filters.department)) return false;
    if (filters.search) {
      const q = filters.search.toLowerCase();
      const hay = [
        r.region,r.state,r.division,r.district,r.category,
        r.deeniKaam,r.field,r.multiName,r.multiValue,r.chain,r.department
      ].join(" ").toLowerCase();
      if (!hay.includes(q)) return false;
    }
    return true;
  });
}

function aggregateForMonth(rows, month, level, filters, fieldOnly=false) {
  const map = new Map();

  for (const r of rows) {
    if (month && r.month !== month) continue;

    if (filters.region && !same(r.region,filters.region)) continue;
    if (filters.state && !same(r.state,filters.state)) continue;
    if (filters.division && !same(r.division,filters.division)) continue;
    if (filters.district && !same(r.district,filters.district)) continue;
    if (filters.category && !same(r.category,filters.category)) continue;
    if (filters.deeniKaam && !same(r.deeniKaam,filters.deeniKaam)) continue;
    if (filters.field && !same(r.field,filters.field)) continue;
    if (filters.chain && !same(r.chain,filters.chain)) continue;
    if (filters.department && !same(r.department,filters.department)) continue;

    let groupValue = "India";
    if (level === "region") groupValue = r.region || "Unknown Region";
    if (level === "state") groupValue = r.state || "Unknown State";
    if (level === "division") groupValue = r.division || "Unknown Division";
    if (level === "district") groupValue = r.district || "Unknown District";

    const multiKey = fieldOnly ? "" : `${r.multiName}\u001F${r.multiValue}`;
    const key = `${groupValue}\u001F${r.deeniKaam}\u001F${r.field}\u001F${multiKey}`;

    const old = map.get(key);
    if (old) {
      old.report += r.report;
      old.target26 += r.target26;
      old.target52 += r.target52;
    } else {
      map.set(key,{
        groupValue,
        deeniKaam:r.deeniKaam,
        field:r.field,
        multiName:r.multiName,
        multiValue:r.multiValue,
        report:r.report,
        target26:r.target26,
        target52:r.target52
      });
    }
  }
  return [...map.values()];
}

function pctChange(oldValue,newValue) {
  if (oldValue === 0 && newValue === 0) return 0;
  if (oldValue === 0 && newValue > 0) return 100;
  return ((newValue-oldValue)/Math.abs(oldValue))*100;
}

function averageValues(rowsByMonth, months, key) {
  const vals = months.map(m => rowsByMonth.get(`${m}\u001F${key}`) || 0);
  return vals.length ? vals.reduce((a,b)=>a+b,0)/vals.length : 0;
}

const MiniTable = ({title,data=[]}) => (
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
          {data.length ? data.slice(0,15).map((d,i)=>(
            <tr key={i} className="hover:bg-slate-50">
              <td className="py-2.5 px-4 text-slate-700 text-center">{d.label}</td>
              <td className="py-2.5 px-4 text-teal-700 font-bold text-right">{Math.round(d.count).toLocaleString("en-IN")}</td>
            </tr>
          )) : (
            <tr><td colSpan="2" className="text-center py-4 text-slate-500">No data available</td></tr>
          )}
        </tbody>
      </table>
    </div>
  </div>
);

export default function DeeniKaamDashboard({ onBack, onLogout, officeUser }) {
  const [data,setData] = useState([]);
  const [loading,setLoading] = useState(false);
  const [fetchError,setFetchError] = useState("");
  const [,startTransition] = useTransition();

  const [activeViewMode,setActiveViewMode] = useState("table");
  const [activeTab,setActiveTab] = useState("Monthly Report");

  const [startMonth,setStartMonth] = useState("");
  const [endMonth,setEndMonth] = useState("");

  const [region,setRegion] = useState(
    officeUser?.region && officeUser.region.toLowerCase()!=="all" ? officeUser.region : ""
  );
  const [state,setState] = useState(
    officeUser?.state && officeUser.state.toLowerCase()!=="all" ? officeUser.state : ""
  );
  const [division,setDivision] = useState(
    officeUser?.division && officeUser.division.toLowerCase()!=="all" ? officeUser.division : ""
  );
  const [district,setDistrict] = useState(
    officeUser?.district && officeUser.district.toLowerCase()!=="all" ? officeUser.district : ""
  );

  const [selectedCategory,setSelectedCategory] = useState("");
  const [selectedDeeniKaam,setSelectedDeeniKaam] = useState("");
  const [selectedField,setSelectedField] = useState("");
  const [selectedChain,setSelectedChain] = useState("");
  const [selectedDepartment,setSelectedDepartment] = useState("");
  const [selectedTarget,setSelectedTarget] = useState("52%");
  const [searchTerm,setSearchTerm] = useState("");
  const [currentPage,setCurrentPage] = useState(1);
  const rowsPerPage=10;

  const [now,setNow]=useState(null);
  useEffect(()=>{
    setNow(new Date());
    const timer=setInterval(()=>setNow(new Date()),1000);
    return ()=>clearInterval(timer);
  },[]);

  const dateStr=now ? now.toLocaleDateString("en-GB",{day:"2-digit",month:"2-digit",year:"numeric"}).replace(/\//g,"-") : "";
  const timeStr=now ? now.toLocaleTimeString("en-US",{hour:"2-digit",minute:"2-digit",second:"2-digit",hour12:true}) : "";

  const fetchData=async()=>{
    setLoading(true);
    setFetchError("");
    try {
      const sourceArrays=[];
      const errors=[];

      // Sequential fetch prevents three very large sheets from competing for memory.
      for (const url of SHEET_URLS) {
        try {
          const rows=await loadOneSheet(url);
          if (rows.length) sourceArrays.push(rows);
        } catch(e) {
          errors.push(e?.message || "Google Sheet request failed");
        }
      }

      const compact=aggregateSources(sourceArrays);

      if (!compact.length) {
        throw new Error(
          "Google Sheet data could not be loaded. Please set each Sheet to Anyone with the link → Viewer (or Publish to web) and check the URL."
        );
      }

      setData(compact);
      setCurrentPage(1);

      if (errors.length && sourceArrays.length) {
        setFetchError(`${sourceArrays.length} source(s) synced. ${errors.length} source(s) could not be read.`);
      }
    } catch(e) {
      setData([]);
      setFetchError(e?.message || "Google Sheet sync failed.");
    } finally {
      setLoading(false);
    }
  };

  useEffect(()=>{ fetchData(); },[]);

  // Reset dependent geography exactly according to Region → State → Division → District.
  const setRegionSafe=(v)=>{
    setRegion(v);
    if (!(officeUser?.state && officeUser.state.toLowerCase()!=="all")) setState("");
    if (!(officeUser?.division && officeUser.division.toLowerCase()!=="all")) setDivision("");
    if (!(officeUser?.district && officeUser.district.toLowerCase()!=="all")) setDistrict("");
    setCurrentPage(1);
  };
  const setStateSafe=(v)=>{
    setState(v);
    if (!(officeUser?.division && officeUser.division.toLowerCase()!=="all")) setDivision("");
    if (!(officeUser?.district && officeUser.district.toLowerCase()!=="all")) setDistrict("");
    setCurrentPage(1);
  };
  const setDivisionSafe=(v)=>{
    setDivision(v);
    if (!(officeUser?.district && officeUser.district.toLowerCase()!=="all")) setDistrict("");
    setCurrentPage(1);
  };

  const availableCategories=useMemo(
    ()=>uniqSorted(data.map(r=>r.category)),
    [data]
  );

  const availableDeeniKaam=useMemo(()=>{
    let x=data;
    if(selectedCategory) x=x.filter(r=>same(r.category,selectedCategory));
    const fromData=uniqSorted(x.map(r=>r.deeniKaam));
    // Always preserve the supplied hierarchy, even if a selected category currently has no rows.
    const fromMaster=uniqSorted(
      FIELD_HIERARCHY.filter(x=>!selectedCategory || same(x[0],selectedCategory)).map(x=>x[1])
    );
    return uniqSorted([...fromMaster,...fromData]);
  },[data,selectedCategory]);

  const availableFields=useMemo(()=>{
    let x=FIELD_HIERARCHY;
    if(selectedCategory) x=x.filter(r=>same(r[0],selectedCategory));
    if(selectedDeeniKaam) x=x.filter(r=>same(r[1],selectedDeeniKaam));
    const master=uniqSorted(x.map(r=>r[2]));
    const actual=uniqSorted(
      data.filter(r=>
        (!selectedCategory || same(r.category,selectedCategory)) &&
        (!selectedDeeniKaam || same(r.deeniKaam,selectedDeeniKaam))
      ).map(r=>r.field)
    );
    return uniqSorted([...master,...actual]);
  },[data,selectedCategory,selectedDeeniKaam]);

  const availableMonths=useMemo(
    ()=>data.map(r=>r.month).filter(Boolean).filter((v,i,a)=>a.indexOf(v)===i).sort(monthSort),
    [data]
  );

  // Latest month is selected automatically after first successful sync.
  useEffect(()=>{
    if(!data.length) return;
    const latest=availableMonths[availableMonths.length-1] || "";
    if(!startMonth && !endMonth && latest) {
      setEndMonth(latest);
      const idx=availableMonths.indexOf(latest);
      setStartMonth(idx>0 ? availableMonths[idx-1] : latest);
    }
  },[data,availableMonths,startMonth,endMonth]);

  const baseFilters=useMemo(()=>({
    region,state,division,district,
    category:selectedCategory,
    deeniKaam:selectedDeeniKaam,
    field:selectedField,
    chain:selectedChain,
    department:selectedDepartment,
    search:searchTerm
  }),[region,state,division,district,selectedCategory,selectedDeeniKaam,selectedField,searchTerm]);

  const selectedLevel= !region ? "country" : !state ? "state" : !division ? "division" : "district";
  const headerLabel= !region ? "COUNTRY" : !state ? "STATE" : !division ? "DIVISION" : "DISTRICT";

  const displayMonths=useMemo(()=>{
    const s=startMonth || endMonth;
    const e=endMonth || startMonth;
    if(!s && !e) return [];
    const lo=s && e ? Math.min(Number(s.replace("-","")),Number(e.replace("-",""))) : Number((s||e).replace("-",""));
    const hi=s && e ? Math.max(Number(s.replace("-","")),Number(e.replace("-",""))) : lo;
    return availableMonths.filter(m=>{
      const n=Number(m.replace("-",""));
      return n>=lo && n<=hi;
    });
  },[availableMonths,startMonth,endMonth]);

  const effectiveEnd=endMonth || startMonth || availableMonths[availableMonths.length-1] || "";
  const effectiveStart=startMonth || effectiveEnd;

  const processedTableData=useMemo(()=>{
    if(!data.length) return [];

    const source=filterRows(data,baseFilters);

    const quarterKey=(m)=>m ? m.slice(0,4)+"-Q"+Math.ceil(Number(m.slice(5,7))/3) : "";
    const quarterMonths=(q)=>{
      if(!q) return [];
      const [y,qn]=q.split("-Q");
      const start=(Number(qn)-1)*3+1;
      return [0,1,2].map(i=>`${y}-${String(start+i).padStart(2,"0")}`);
    };

    const yearMonths=(y)=>{
      if(!y) return [];
      return Array.from({length:12},(_,i)=>`${y}-${String(i+1).padStart(2,"0")}`);
    };

    let months=displayMonths.length ? displayMonths : [effectiveEnd];
    if(activeTab==="Quarterly Report") {
      months=uniqSorted([
        ...quarterMonths(quarterKey(effectiveStart)),
        ...quarterMonths(quarterKey(effectiveEnd))
      ]);
    }
    if(activeTab==="Yearly Report") {
      months=uniqSorted([
        ...yearMonths(effectiveStart.slice(0,4)),
        ...yearMonths(effectiveEnd.slice(0,4))
      ]);
    }

    // PERFORMANCE: build all required month maps in ONE pass over the
    // already compacted data instead of rescanning it once per month.
    const monthlyMaps=new Map(months.map(m=>[m,new Map()]));
    const wantedMonths=new Set(months);
    for (const r of source) {
      if (!wantedMonths.has(r.month)) continue;

      let groupValue = "India";
      if (selectedLevel === "region") groupValue = r.region || "Unknown Region";
      else if (selectedLevel === "state") groupValue = r.state || "Unknown State";
      else if (selectedLevel === "division") groupValue = r.division || "Unknown Division";
      else if (selectedLevel === "district") groupValue = r.district || "Unknown District";

      const key=`${groupValue}\u001F${r.deeniKaam}\u001F${r.field}\u001F${r.multiName}\u001F${r.multiValue}`;
      const map=monthlyMaps.get(r.month);
      const old=map.get(key);
      if(old) {
        old.report+=r.report;
        old.target26+=r.target26;
        old.target52+=r.target52;
      } else {
        map.set(key,{
          groupValue,
          deeniKaam:r.deeniKaam,
          field:r.field,
          multiName:r.multiName,
          multiValue:r.multiValue,
          report:r.report,
          target26:r.target26,
          target52:r.target52
        });
      }
    }

    let endMap=monthlyMaps.get(effectiveEnd) || new Map();
    let startMap=monthlyMaps.get(effectiveStart) || new Map();

    if(activeTab==="Quarterly Report") {
      const endQuarter=quarterMonths(quarterKey(effectiveEnd));
      const startQuarter=quarterMonths(quarterKey(effectiveStart));
      const sumMaps=(list)=>{
        const out=new Map();
        for(const m of list) {
          for(const [k,z] of (monthlyMaps.get(m)||new Map())) {
            const old=out.get(k);
            if(old) {
              old.report+=z.report; old.target26+=z.target26; old.target52+=z.target52;
            } else out.set(k,{...z});
          }
        }
        return out;
      };
      endMap=sumMaps(endQuarter);
      startMap=sumMaps(startQuarter);
    }

    if(activeTab==="Yearly Report") {
      const endYear=yearMonths(effectiveEnd.slice(0,4));
      const startYear=yearMonths(effectiveStart.slice(0,4));
      const sumMaps=(list)=>{
        const out=new Map();
        for(const m of list) {
          for(const [k,z] of (monthlyMaps.get(m)||new Map())) {
            const old=out.get(k);
            if(old) {
              old.report+=z.report; old.target26+=z.target26; old.target52+=z.target52;
            } else out.set(k,{...z});
          }
        }
        return out;
      };
      endMap=sumMaps(endYear);
      startMap=sumMaps(startYear);
    }

    const rows=[];
    for(const [key,r] of endMap) {
      const start=startMap.get(key)?.report || 0;
      const current=r.report;
      const target=selectedTarget==="26%" ? r.target26 : r.target52;
      const comparison=pctChange(start,current);

      // Average mode: selected end month against average of all months in selected range.
      let average=0;
      if(activeViewMode==="average" || activeTab==="Average Report") {
        let sum=0, count=0;
        for(const m of months) {
          const v=monthlyMaps.get(m)?.get(key)?.report;
          if(typeof v==="number") { sum+=v; count++; }
        }
        average=count ? sum/count : 0;
        rows.push({
          ...r, current, target,
          achievement:target>0 ? current/target*100 : null,
          startValue:start, endValue:current,
          average, comparison:pctChange(average,current)
        });
      } else if(activeTab==="Quarterly Report" || activeTab==="Yearly Report") {
        rows.push({
          ...r,current,target,
          achievement:target>0 ? current/target*100 : null,
          achievement26:r.target26>0 ? current/r.target26*100 : null,
          achievement52:r.target52>0 ? current/r.target52*100 : null,
          startValue:start,endValue:current,
          comparison
        });
      } else {
        rows.push({
          ...r,current,target,
          achievement:target>0 ? current/target*100 : null,
          achievement26:r.target26>0 ? current/r.target26*100 : null,
          achievement52:r.target52>0 ? current/r.target52*100 : null,
          startValue:start,endValue:current,
          comparison
        });
      }
    }

    return rows.sort((a,b)=>{
      const g=norm(a.groupValue).localeCompare(norm(b.groupValue),undefined,{numeric:true});
      if(g) return g;
      const d=norm(a.deeniKaam).localeCompare(norm(b.deeniKaam));
      if(d) return d;
      return norm(a.field).localeCompare(norm(b.field));
    });
  },[
    data,baseFilters,effectiveStart,effectiveEnd,displayMonths,
    selectedLevel,selectedTarget,activeViewMode,activeTab
  ]);

  useEffect(()=>{ setCurrentPage(1); },[
    region,state,division,district,selectedCategory,selectedDeeniKaam,
    selectedField,selectedTarget,selectedChain,selectedDepartment,startMonth,endMonth,searchTerm,activeViewMode,activeTab
  ]);

  const pagedRows=useMemo(()=>{
    const start=(currentPage-1)*rowsPerPage;
    return processedTableData.slice(start,start+rowsPerPage);
  },[processedTableData,currentPage]);

  const dynamicGraphData=useMemo(()=>{
    const map=new Map();
    for(const r of processedTableData) {
      const k=r.groupValue || "India";
      map.set(k,(map.get(k)||0)+r.current);
    }
    return [...map.entries()]
      .map(([label,count])=>({label,count}))
      .sort((a,b)=>b.count-a.count);
  },[processedTableData]);

  const dynamicGraphTitle=
    !region ? "INDIA TOTAL" :
    !state ? "REPORTS BY STATE" :
    !division ? "REPORTS BY DIVISION" :
    "REPORTS BY DISTRICT";

  const maxDynamicCount=dynamicGraphData.length
    ? Math.max(...dynamicGraphData.map(d=>d.count))
    : 0;

  const stateMetrics=useMemo(()=>{
    const m=new Map();
    for(const r of processedTableData) {
      const k=r.groupValue || "India";
      m.set(k,(m.get(k)||0)+r.current);
    }
    return [...m.entries()].map(([label,count])=>({label,count})).sort((a,b)=>b.count-a.count);
  },[processedTableData]);

  const downloadExcel=()=>{
    if(!processedTableData.length) return alert("No data to export");
    const rows=processedTableData.map(r=>({
      [headerLabel]:r.groupValue,
      "Category":r.category,
      "Deeni Kaam":r.deeniKaam,
      "Fields":r.field,
      "Multiple Field Name":r.multiName,
      "Multiple Field Value":r.multiValue,
      [formatMonth(effectiveEnd)]:r.current,
      "Target 26%":r.target26,
      "Target 52%":r.target52,
      "Selected Target":r.target,
      "Achievement %":r.achievement==null ? "" : Number(r.achievement.toFixed(2)),
      "Achievement 26%":r.achievement26==null ? "" : Number(r.achievement26.toFixed(2)),
      "Achievement 52%":r.achievement52==null ? "" : Number(r.achievement52.toFixed(2)),
      [formatMonth(effectiveStart)]:r.startValue,
      "Comparison %":Number(r.comparison.toFixed(2))
    }));
    const ws=XLSX.utils.json_to_sheet(rows);
    const wb=XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb,ws,"Dashboard Report");
    XLSX.writeFile(wb,`12_Deeni_Kaam_Report_${effectiveEnd || Date.now()}.xlsx`);
  };

  const downloadPPT=async()=>{
    if(!processedTableData.length) return alert("No data to export");
    const pres=new pptxgen();
    const slide=pres.addSlide();
    slide.background={color:"FFFFFF"};
    slide.addText("12 DEENI KAAM REPORT",{x:.5,y:.5,w:12,h:.5,fontSize:24,bold:true,color:"008B8B",align:"center"});
    slide.addText(`${headerLabel}: ${region||state||division||district||"India"} | ${activeTab==="Yearly Report" ? effectiveEnd.slice(0,4) : activeTab==="Quarterly Report" ? quarterKeyForMonth(effectiveEnd) : formatMonth(effectiveEnd)}`,{x:.6,y:1.1,w:11.8,h:.35,fontSize:11,color:"334155",align:"center"});
    const tableRows=[
      [headerLabel,"DEENI KAAM","FIELDS",formatMonth(effectiveEnd),"TARGET","ACHIEVEMENT","COMPARISON"],
      ...processedTableData.slice(0,15).map(r=>[
        r.groupValue,r.deeniKaam||"-",r.field,Math.round(r.current).toLocaleString("en-IN"),
        Math.round(r.target).toLocaleString("en-IN"),
        r.achievement==null ? "-" : `${r.achievement.toFixed(1)}%`,
        `${r.comparison>=0?"+":""}${r.comparison.toFixed(1)}%`
      ])
    ];
    slide.addTable(tableRows,{x:.4,y:1.7,w:12.5,h:5.2,fontSize:8,border:{type:"solid",color:"D9E1E8",pt:1},color:"1F2937",bold:false});
    await pres.writeFile({fileName:`12_Deeni_Kaam_Report_${Date.now()}.pptx`});
  };

  const downloadJPEG=async()=>{
    const element=document.getElementById("report-card-view");
    if(!element) return;
    try {
      const canvas=await html2canvas(element,{
        scale:2,backgroundColor:"#e0f2f1",useCORS:true,
        ignoreElements:n=>n.hasAttribute("data-html2canvas-ignore")
      });
      const link=document.createElement("a");
      link.download=`Deeni_Kaam_Report_Card_${Date.now()}.jpg`;
      link.href=canvas.toDataURL("image/jpeg",.9);
      link.click();
    } catch { alert("Image download failed."); }
  };

  const geoDisabled=(key)=>
    Boolean(officeUser?.[key] && String(officeUser[key]).toLowerCase()!=="all");

  return (
    <div className="min-h-screen bg-[#e0f2f1] text-slate-800 font-sans relative overflow-hidden">
      <div className="fixed inset-0 z-0 pointer-events-none">
        <div className="absolute top-0 right-0 w-[40rem] h-[40rem] bg-teal-100/40 rounded-full blur-[150px]" />
      </div>

      <div className="relative z-10" id="report-card-view">
        <header className="max-w-[1400px] mx-auto px-4 sm:px-6 lg:px-8 pt-4">
          <div className="bg-white border border-slate-200 shadow-sm rounded-2xl px-6 py-5 flex flex-col md:flex-row items-center justify-between gap-6">
            <div className="flex items-center gap-4 w-full md:w-auto">
              <button data-html2canvas-ignore="true" onClick={onBack} className="p-3 bg-slate-100 hover:bg-slate-200 border border-slate-200 text-slate-700 rounded-xl shadow-sm">
                <ArrowLeft className="w-5 h-5"/>
              </button>
              <div className="p-3 bg-teal-700 text-white rounded-xl hidden sm:block">
                <BookOpenCheck className="w-6 h-6"/>
              </div>
              <div>
                <h1 className="text-xl sm:text-2xl font-black text-slate-800 tracking-wide">12 DEENI KAAM <span className="font-light text-teal-700">REPORT HUB</span></h1>
                <p className="text-[10px] font-bold text-teal-700 uppercase tracking-widest mt-0.5">Professional Multi-View Dashboard</p>
              </div>
            </div>

            <div className="flex flex-wrap items-center justify-center md:justify-end gap-3 w-full md:w-auto">
              <div className="text-[10px] font-bold text-slate-600 uppercase tracking-widest bg-slate-50 px-3 py-2 rounded-lg border border-slate-200 hidden lg:block">
                USER: <span className="text-teal-700">{officeUser?.name || officeUser?.userId || "Admin"}</span>
              </div>
              <button data-html2canvas-ignore="true" onClick={fetchData} disabled={loading} className="flex items-center gap-2 text-xs font-bold uppercase bg-teal-600 hover:bg-teal-700 text-white px-4 py-2.5 rounded-lg">
                <RefreshCw className={`w-4 h-4 ${loading?"animate-spin":""}`}/> <span>SYNC</span>
              </button>
              <button data-html2canvas-ignore="true" onClick={downloadJPEG} className="flex items-center gap-2 text-xs font-bold uppercase bg-[#0f766e] text-white px-4 py-2.5 rounded-lg">
                <ImageIcon className="w-4 h-4"/> <span>JPEG</span>
              </button>
              <button data-html2canvas-ignore="true" onClick={downloadPPT} className="text-xs font-bold uppercase bg-teal-600 text-white px-4 py-2.5 rounded-lg">PPT</button>
              <button data-html2canvas-ignore="true" onClick={downloadExcel} className="flex items-center gap-2 text-xs font-bold uppercase bg-[#198754] text-white px-4 py-2.5 rounded-lg">
                <Download className="w-4 h-4"/> <span>EXCEL</span>
              </button>
              <div className="w-full flex justify-end items-center gap-3 mt-2">
                <div className="text-[11px] font-extrabold text-black uppercase bg-white px-4 py-2.5 rounded-lg border border-slate-200 hidden lg:flex items-center gap-2">
                  <Clock className="w-4 h-4 text-teal-600"/>
                  <span>DATE: {dateStr} • TIME: {timeStr}</span>
                </div>
                <button data-html2canvas-ignore="true" onClick={onLogout} className="flex items-center gap-2 text-xs font-bold uppercase bg-[#dc3545] text-white px-4 py-2.5 rounded-lg">
                  <LogOut className="w-4 h-4"/> <span>LOGOUT</span>
                </button>
              </div>
            </div>
          </div>
        </header>

        <main className="max-w-[1400px] mx-auto px-4 sm:px-6 lg:px-8 pt-6 space-y-6">
          {fetchError && (
            <div data-html2canvas-ignore="true" className={`p-4 rounded-xl text-sm font-medium flex items-center gap-3 ${data.length ? "bg-amber-50 border border-amber-200 text-amber-700" : "bg-red-50 border border-red-200 text-red-600"}`}>
              <div className={`w-2 h-2 rounded-full ${data.length?"bg-amber-500":"bg-red-500"}`}/>
              {fetchError}
            </div>
          )}

          <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4" data-html2canvas-ignore="true">
            {[
              {id:"table",label:"DETAILED TABLE REPORT",icon:LayoutDashboard,desc:"Grid & Month-wise data"},
              {id:"graphs",label:"VISUAL ANALYTICS & GRAPHS",icon:BarChart3,desc:"District/State charts"},
              {id:"average",label:"AVERAGE & COMPARISON",icon:TrendingUp,desc:"Average range analysis"},
              {id:"targets",label:"TARGETS & ACHIEVEMENT",icon:Target,desc:"26% / 52% performance"}
            ].map(btn=>{
              const I=btn.icon,isActive=activeViewMode===btn.id;
              return (
                <button key={btn.id} onClick={()=>startTransition(()=>setActiveViewMode(btn.id))}
                  className={`p-5 rounded-2xl border text-left shadow-sm flex items-start gap-3.5 ${isActive?"bg-[#0f4c47] text-white border-[#0f4c47]":"bg-white text-slate-700 border-slate-200"}`}>
                  <div className={`p-3 rounded-xl ${isActive?"bg-teal-700 text-white":"bg-teal-50 text-teal-700"}`}><I className="w-5 h-5"/></div>
                  <div><h3 className="text-xs font-black uppercase tracking-wider">{btn.label}</h3><p className={`text-[11px] mt-1 ${isActive?"text-teal-200":"text-slate-400 font-medium"}`}>{btn.desc}</p></div>
                </button>
              );
            })}
          </div>

          {activeViewMode==="table" && (
            <div className="flex gap-3 border-b border-slate-200 pb-2" data-html2canvas-ignore="true">
              {["Monthly Report","Average Report","Quarterly Report","Yearly Report"].map(tab=>(
                <button key={tab} onClick={()=>startTransition(()=>setActiveTab(tab))}
                  className={`px-5 py-2.5 rounded-xl text-xs font-bold uppercase tracking-wider shadow-sm ${activeTab===tab?"bg-teal-700 text-white":"bg-white text-slate-700 border border-slate-200"}`}>
                  {tab}
                </button>
              ))}
            </div>
          )}

          <div className="bg-white p-5 rounded-xl border border-slate-200 shadow-sm">
            <div className="flex flex-col lg:flex-row justify-between lg:items-center gap-4 mb-5 border-b border-slate-100 pb-4">
              <div className="flex items-center gap-3">
                <Filter className="w-5 h-5 text-teal-700"/>
                <h2 className="text-sm font-bold text-slate-800 tracking-widest uppercase">Global Filters & Range Picker</h2>
              </div>
              <div className="flex flex-wrap items-center gap-3" data-html2canvas-ignore="true">
                <div className="flex items-center gap-2 bg-slate-50 p-1.5 rounded-lg border border-slate-200">
                  <Calendar className="w-4 h-4 text-slate-400 ml-2"/>
                  <select value={startMonth} onChange={e=>setStartMonth(e.target.value)} className="bg-transparent text-xs font-bold outline-none max-w-[120px]">
                    <option value="">Start Month</option>
                    {availableMonths.map(m=><option key={m} value={m}>{formatMonth(m)}</option>)}
                  </select>
                  <span className="text-slate-400 text-xs font-bold">TO</span>
                  <select value={endMonth} onChange={e=>setEndMonth(e.target.value)} className="bg-transparent text-xs font-bold outline-none max-w-[120px]">
                    <option value="">End Month</option>
                    {availableMonths.map(m=><option key={m} value={m}>{formatMonth(m)}</option>)}
                  </select>
                </div>
                {(startMonth||endMonth) && <button onClick={()=>{setStartMonth("");setEndMonth("");}} className="text-xs text-red-500 font-bold underline px-2">Reset</button>}
              </div>
            </div>

            <div className="grid grid-cols-2 md:grid-cols-9 gap-3">
              <div className="col-span-2 md:col-span-1" data-html2canvas-ignore="true">
                <label className="block text-[10px] font-bold text-slate-500 uppercase tracking-widest mb-2">Search</label>
                <div className="relative">
                  <span className="absolute inset-y-0 left-0 pl-3 flex items-center text-slate-400"><Search className="w-4 h-4"/></span>
                  <input value={searchTerm} onChange={e=>setSearchTerm(e.target.value)} placeholder="Search..." className="w-full pl-9 pr-3 py-2.5 rounded-lg border border-slate-300 bg-white text-sm outline-none"/>
                </div>
              </div>

              <div className="flex flex-col">
                <label className="block text-[10px] font-bold text-slate-500 uppercase tracking-widest mb-2">Category</label>
                <select value={selectedCategory} onChange={e=>{setSelectedCategory(e.target.value);setSelectedDeeniKaam("");setSelectedField("");}} className="w-full p-2.5 rounded-lg border border-slate-300 bg-white text-sm font-medium">
                  <option value="">All Categories</option>
                  {availableCategories.map(v=><option key={v} value={v}>{v}</option>)}
                </select>
              </div>

              <div className="flex flex-col">
                <label className="block text-[10px] font-bold text-slate-500 uppercase tracking-widest mb-2">Deeni Kaam</label>
                <select value={selectedDeeniKaam} onChange={e=>{setSelectedDeeniKaam(e.target.value);setSelectedField("");}} className="w-full p-2.5 rounded-lg border border-slate-300 bg-white text-sm font-medium">
                  <option value="">All Deeni Kaam</option>
                  {availableDeeniKaam.map(v=><option key={v} value={v}>{v}</option>)}
                </select>
              </div>

              <div className="flex flex-col">
                <label className="block text-[10px] font-bold text-slate-500 uppercase tracking-widest mb-2">Fields</label>
                <select value={selectedField} onChange={e=>setSelectedField(e.target.value)} className="w-full p-2.5 rounded-lg border border-slate-300 bg-white text-sm font-medium">
                  <option value="">All Fields</option>
                  {availableFields.map(v=><option key={v} value={v}>{v}</option>)}
                </select>
              </div>

              <div className="flex flex-col">
                <label className="block text-[10px] font-bold text-slate-500 uppercase tracking-widest mb-2">Targets</label>
                <select value={selectedTarget} onChange={e=>setSelectedTarget(e.target.value)} className="w-full p-2.5 rounded-lg border border-slate-300 bg-white text-sm font-medium">
                  <option value="52%">52%</option>
                  <option value="26%">26%</option>
                  <option value="BOTH">Both</option>
                </select>
              </div>

              <div className="flex flex-col">
                <label className="block text-[10px] font-bold text-slate-500 uppercase tracking-widest mb-2">Chain</label>
                <select value={selectedChain} onChange={e=>setSelectedChain(e.target.value)} className="w-full p-2.5 rounded-lg border border-slate-300 bg-white text-sm font-medium">
                  <option value="">All</option>
                  {uniqSorted(data.map(r=>r.chain)).map(v=><option key={v} value={v}>{v}</option>)}
                </select>
              </div>

              <div className="flex flex-col">
                <label className="block text-[10px] font-bold text-slate-500 uppercase tracking-widest mb-2">Department</label>
                <select value={selectedDepartment} onChange={e=>setSelectedDepartment(e.target.value)} className="w-full p-2.5 rounded-lg border border-slate-300 bg-white text-sm font-medium">
                  <option value="">All</option>
                  {uniqSorted(data.map(r=>r.department)).map(v=><option key={v} value={v}>{v}</option>)}
                </select>
              </div>

              {[
                {label:"Region",val:region,set:setRegionSafe,opts:uniqSorted(data.map(r=>r.region)),key:"region"},
                {label:"State",val:state,set:setStateSafe,opts:uniqSorted(data.filter(r=>!region||same(r.region,region)).map(r=>r.state)),key:"state"},
                {label:"Division",val:division,set:setDivisionSafe,opts:uniqSorted(data.filter(r=>(!region||same(r.region,region))&&(!state||same(r.state,state))).map(r=>r.division)),key:"division"},
                {label:"District",val:district,set:setDistrict,opts:uniqSorted(data.filter(r=>(!region||same(r.region,region))&&(!state||same(r.state,state))&&(!division||same(r.division,division))).map(r=>r.district)),key:"district"}
              ].map(f=>(
                <div key={f.key} className="flex flex-col">
                  <label className="block text-[10px] font-bold text-slate-500 uppercase tracking-widest mb-2">{f.label}</label>
                  <select value={f.val} onChange={e=>f.set(e.target.value)} disabled={geoDisabled(f.key)} className="w-full p-2.5 rounded-lg border border-slate-300 bg-white text-sm font-medium disabled:bg-slate-50">
                    <option value="">All</option>
                    {f.opts.map(v=><option key={v} value={v}>{v}</option>)}
                  </select>
                </div>
              ))}
            </div>
          </div>

          {activeViewMode==="graphs" && (
            <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
              <div className="bg-white rounded-xl shadow-sm border border-slate-200 p-6 h-[400px] flex flex-col">
                <h3 className="text-xs font-bold text-teal-700 uppercase tracking-widest mb-4">{dynamicGraphTitle}</h3>
                <div className="flex-1 flex justify-start items-end gap-4 pl-4 pb-8 relative border-b border-slate-300 overflow-x-auto">
                  {dynamicGraphData.length ? dynamicGraphData.slice(0,12).map((d,i)=>{
                    const h=maxDynamicCount ? d.count/maxDynamicCount*100 : 0;
                    return <div key={i} className="flex flex-col justify-end items-center relative h-full w-12 shrink-0">
                      <span className="text-[10px] font-bold text-slate-700 mb-1">{Math.round(d.count).toLocaleString("en-IN")}</span>
                      <div style={{height:`${Math.max(h,2)}%`}} className="w-10 bg-teal-700 rounded-t-md"/>
                      <span className="absolute -bottom-7 w-24 text-center text-[9px] text-slate-600 truncate">{d.label}</span>
                    </div>
                  }) : <div className="m-auto text-slate-400 text-xs">No graph data found</div>}
                </div>
              </div>
              <div className="grid grid-cols-1 md:grid-cols-2 gap-6 h-[400px]">
                <MiniTable title={`${headerLabel} WISE METRICS`} data={stateMetrics}/>
                <MiniTable title="FIELD METRICS" data={processedTableData.map(r=>({label:r.field,count:r.current})).sort((a,b)=>b.count-a.count)}/>
              </div>
            </div>
          )}
        </main>

        <main className="max-w-[1400px] mx-auto px-4 sm:px-6 lg:px-8 pb-8 mt-2">
          <div className="bg-white rounded-xl border border-slate-200 shadow-sm">
            <div className="p-4 border-b border-slate-200 bg-[#e0f2f1] rounded-t-xl">
              <h2 className="text-sm font-bold text-teal-800 uppercase tracking-widest">
                {activeViewMode==="average" ? "Average & Range Comparison Matrix" : activeViewMode==="targets" ? "Targets & Achievement Matrix" : `Detailed Telemetry Output (${activeTab})`}
              </h2>
              <p className="text-[10px] font-bold text-teal-800 mt-1 uppercase tracking-widest">
                SOURCE: {data.length.toLocaleString("en-IN")} AGGREGATED RECORDS • VISIBLE: {processedTableData.length.toLocaleString("en-IN")}
              </p>
            </div>

            <div className="overflow-x-auto w-full">
              <table className="min-w-full text-left text-[11px] lg:text-xs">
                <thead className="bg-[#008b8b]">
                  <tr>
                    <th className="px-2 py-3 font-bold text-white uppercase border-r border-white/25 text-center">{headerLabel}</th>
                    <th className="px-2 py-3 font-bold text-white uppercase border-r border-white/25 text-center">DEENI KAAM</th>
                    <th className="px-2 py-3 font-bold text-white uppercase border-r border-white/25 text-center">FIELDS</th>
                    <th className="px-2 py-3 font-bold text-white uppercase border-r border-white/25 text-center">MULTIPLE FIELD</th>
                    <th className="px-2 py-3 font-bold text-white uppercase border-r border-white/25 text-center bg-[#007a7a]">{activeTab==="Yearly Report" ? effectiveEnd.slice(0,4) : activeTab==="Quarterly Report" ? quarterKeyForMonth(effectiveEnd) : formatMonth(effectiveEnd)}</th>
                    <th className="px-2 py-3 font-bold text-white uppercase border-r border-white/25 text-center bg-[#007a7a]">TARGET</th>
                    <th className="px-2 py-3 font-bold text-white uppercase border-r border-white/25 text-center bg-[#007a7a]">ACHIEVEMENT (%)</th>
                    <th className="px-2 py-3 font-bold text-white uppercase border-r border-white/25 text-center bg-[#007a7a]">{activeTab==="Yearly Report" ? effectiveStart.slice(0,4) : activeTab==="Quarterly Report" ? quarterKeyForMonth(effectiveStart) : formatMonth(effectiveStart)}</th>
                    <th className="px-2 py-3 font-bold text-white uppercase border-r border-white/25 text-center bg-[#007a7a]">COMPARISON (%)</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {pagedRows.length ? pagedRows.map((r,i)=>(
                    <tr key={`${r.groupValue}-${r.field}-${r.multiValue}-${i}`} className="hover:bg-slate-50">
                      <td className="px-2 py-2 font-bold text-teal-800 border-r border-slate-100 text-center">{r.groupValue||"India"}</td>
                      <td className="px-2 py-2 font-bold text-slate-800 border-r border-slate-100 text-center">{r.deeniKaam||"-"}</td>
                      <td className="px-2 py-2 font-bold text-slate-700 border-r border-slate-100 text-center">{r.field}</td>
                      <td className="px-2 py-2 text-slate-600 border-r border-slate-100 text-center">{r.multiName ? `${r.multiName}: ${r.multiValue||"-"}` : "-"}</td>
                      <td className="px-2 py-2 text-slate-700 border-r border-slate-100 text-center font-semibold">{Math.round(r.current).toLocaleString("en-IN")}</td>
                      <td className="px-2 py-2 text-slate-700 border-r border-slate-100 text-center font-semibold">
                        {selectedTarget==="BOTH" ? `26%: ${Math.round(r.target26).toLocaleString("en-IN")} | 52%: ${Math.round(r.target52).toLocaleString("en-IN")}` : Math.round(r.target).toLocaleString("en-IN")}
                      </td>
                      <td className="px-2 py-2 text-blue-600 font-bold border-r border-slate-100 text-center">
                        {selectedTarget==="BOTH"
                          ? `26%: ${r.achievement26==null?"-":r.achievement26.toFixed(1)+"%"} | 52%: ${r.achievement52==null?"-":r.achievement52.toFixed(1)+"%"}`
                          : r.achievement==null ? "-" : `${r.achievement.toFixed(1)}%`}
                      </td>
                      <td className="px-2 py-2 text-slate-700 border-r border-slate-100 text-center font-semibold">{Math.round(r.startValue).toLocaleString("en-IN")}</td>
                      <td className={`px-2 py-2 font-bold text-center ${r.comparison>=0?"text-emerald-600":"text-red-600"}`}>
                        {r.comparison>=0?"+":""}{r.comparison.toFixed(1)}%
                      </td>
                    </tr>
                  )) : (
                    <tr><td colSpan="9" className="px-6 py-12 text-center text-slate-500 text-xs uppercase tracking-widest">
                      <div className="flex flex-col items-center justify-center gap-3">
                        <Activity className="w-6 h-6 opacity-40 text-teal-600"/>
                        <span>{loading?"Loading Google Sheet data...":"No matching report found"}</span>
                      </div>
                    </td></tr>
                  )}
                </tbody>
              </table>
            </div>

            {processedTableData.length>rowsPerPage && (
              <div className="p-4 border-t border-slate-200 bg-[#f8fafc] rounded-b-xl flex justify-center items-center gap-2">
                <button disabled={currentPage===1} onClick={()=>setCurrentPage(p=>p-1)} className="px-3 py-1.5 rounded-lg border border-teal-600 bg-white text-xs font-bold text-teal-700 disabled:opacity-30">PREV</button>
                <span className="text-xs font-bold text-teal-700 px-4">PAGE {currentPage} OF {Math.ceil(processedTableData.length/rowsPerPage)}</span>
                <button disabled={currentPage===Math.ceil(processedTableData.length/rowsPerPage)} onClick={()=>setCurrentPage(p=>p+1)} className="px-3 py-1.5 rounded-lg border border-teal-600 bg-white text-xs font-bold text-teal-700 disabled:opacity-30">NEXT</button>
              </div>
            )}
          </div>
        </main>
      </div>
    </div>
  );
}
