import React, { useState, useEffect } from "react";
import * as XLSX from "xlsx";
import Papa from "papaparse";
import _ from "lodash";

/*───────────────────────────────────────────────────────────*
  Normalisation & helpers
 *───────────────────────────────────────────────────────────*/
const squash = (s = "") => s.toString().toLowerCase().replace(/\s+/g, " ").trim();
const stripLevel = (s = "") => squash(s).replace(/^\(l[0-9]+\)\s*/, "");
const cleanTitle = (s = "") =>
  stripLevel(s)
    .replace(/\((ms|dc) only\)/gi, "") // remove (MS only)/(DC only)
    .replace(/_x000d_\n/g, " ") // Excel linebreak artefact
    .replace(/[“”"']/g, "")
    .replace(/[^a-z0-9 ]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();

const cosine = (a, b) => {
  if (!a || !b) return 0;
  if (a === b) return 1;
  const words = [...new Set([...a.split(" "), ...b.split(" ")])];
  const vec = (t) => words.map((w) => t.split(" ").filter((x) => x === w).length);
  const v1 = vec(a);
  const v2 = vec(b);
  const dot = v1.reduce((s, v, i) => s + v * v2[i], 0);
  const mag = (v) => Math.sqrt(v.reduce((s, x) => s + x * x, 0));
  const m1 = mag(v1);
  const m2 = mag(v2);
  return !m1 || !m2 ? 0 : dot / (m1 * m2);
};

const readBuf = (f) => new Promise((res, rej) => { const r = new FileReader(); r.onload = (e) => res(e.target.result); r.onerror = rej; r.readAsArrayBuffer(f); });
const readTxt = (f) => new Promise((res, rej) => { const r = new FileReader(); r.onload = (e) => res(e.target.result); r.onerror = rej; r.readAsText(f); });

/*───────────────────────────────────────────────────────────*/
export default function BSSCISAnalyzer() {
  const [analysis, setAnalysis] = useState(null);
  const [loading, setLoading] = useState(false);

  /* inject css once */
  useEffect(() => {
    if (document.getElementById("bss-cis-css")) return;
    const s = document.createElement("style");
    s.id = "bss-cis-css";
    s.innerHTML = `:root{--gap:1rem}body{margin:0;font-family:system-ui,Segoe UI,Roboto,Helvetica,Arial,sans-serif;background:#f8fafc}h1{margin:0 0 1rem;font-size:1.7rem}.grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(170px,1fr));gap:var(--gap)}.card{background:#fff;border-radius:8px;border:2px solid var(--clr);padding:1rem;text-align:center;box-shadow:0 1px 4px rgba(0,0,0,.05)}.card h3{margin:.1rem 0 .4rem;color:#64748b;font-size:.8rem;font-weight:600}.card .v{font-size:1.8rem;font-weight:700;color:var(--clr)}table{width:100%;border-collapse:collapse;font-size:.8rem;margin-top:.5rem}th,td{padding:.5rem .65rem;border-bottom:1px solid #e2e8f0;text-align:left}th{background:#f1f5f9;font-weight:600}.pill{font-size:.65rem;color:#fff;padding:2px 6px;border-radius:4px}button.primary{padding:.55rem 1rem;background:#2563eb;color:#fff;border:none;border-radius:6px;cursor:pointer}button.export{background:#16a34a;margin-top:var(--gap)}`;
    document.head.appendChild(s);
  }, []);

  const compStatus = (c) => {
    if (!c) return "Not Scanned";
    if (c.failed_instances && c.failed_instances !== "None") return "Fail";
    if (c.passed_instances && c.passed_instances !== "None") return "Pass";
    return "Skipped";
  };

  const pickCol = (row, prefix) => {
    if (!row) return "";
    const k = Object.keys(row).find((x) => x.startsWith(prefix));
    return k ? row[k] : "";
  };

  const buildRec = (b, c) => ({
    BSS_ID: b ? b["cis #"] || "" : "",
    CIS_ID: c ? c.check_id || "" : "",
    BSS_Title: b ? b["synapxe setting title"] || b["cis setting title (for reference only)"] || "" : "",
    CIS_Title: c ? c.title || "" : "",
    Title_Match: cleanTitle(b ? b["synapxe setting title"] || "" : "") === cleanTitle(c ? c.title || "" : "") ? "Yes" : "No",
    BSS_Category: b ? b.category || b["cis section header"] || "Uncategorised" : "Uncategorised",
    "Synapxe Value": pickCol(b, "synapxe value"),
    "Synapxe Exceptions": pickCol(b, "synapxe exceptions"),
    "CIS Recommended Value": pickCol(b, "cis recommended value"),
    "Setting Applicability": pickCol(b, "setting applicability"),
    "Change Description / Remarks": pickCol(b, "change description"),
    Passed: c ? c.passed_instances : "",
    Failed: c ? c.failed_instances : "",
    Compliance: compStatus(c),
  });

  /* main */
  const processFiles = async (bssFile, cisFile) => {
    setLoading(true);
    try {
      /* BSS */
      const wb = XLSX.read(await readBuf(bssFile));
      const sheetName = wb.SheetNames.find((n) => /settings|windows/i.test(n)) || wb.SheetNames[0];
      const raw = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header: 1, defval: "" });
      const hdrIdx = raw.findIndex((r) => r.some((c) => squash(c).includes("cis #")));
      if (hdrIdx === -1) throw new Error("Header row not found");
      const hdrs = raw[hdrIdx].map((h) => squash(h));
      const bssRows = raw.slice(hdrIdx + 1).filter((r) => r.some(Boolean)).map((r) => { const o = {}; hdrs.forEach((h,i)=>o[h]=r[i]); return o; });

      /* CIS */
      const txt = await readTxt(cisFile);
      const arr = txt.split(/\r?\n/);
      const start = arr.findIndex((l)=>l.includes("check_id"));
      const cisRows = Papa.parse(arr.slice(start).join("\n"), { header:true,dynamicTyping:true,skipEmptyLines:true }).data;
      const cisMap = new Map(cisRows.map((c)=>[squash(c.check_id), c]));

      const merged=[];
      bssRows.forEach(b=>{
        const id=b["cis #"]; let cis=cisMap.get(squash(id));
        if(!cis){
          const t=cleanTitle(b["synapxe setting title"]||b["cis setting title (for reference only)"]);
          cis=cisRows.find(c=>cosine(t,cleanTitle(c.title))>0.93);
        }
        merged.push(buildRec(b,cis));
      });
      cisRows.forEach(c=>{ if(!merged.find(m=>squash(m.CIS_ID)===squash(c.check_id))) merged.push(buildRec(null,c)); });

      const titleGroups=_.groupBy(merged,m=>cleanTitle(m.BSS_Title||m.CIS_Title));
      const sameTitleDiffId=Object.values(titleGroups).filter(g=>_.uniqBy(g,x=>x.CIS_ID+"|"+x.BSS_ID).length>1).flat();

      const count=fn=>merged.filter(fn).length;
      const summary={
        total:merged.length,
        bssOnly:count(m=>m.BSS_ID&&!m.CIS_ID),
        cisOnly:count(m=>!m.BSS
