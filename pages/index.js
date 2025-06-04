import React, { useState, useEffect } from "react";
import * as XLSX from "xlsx";
import Papa from "papaparse";
import _ from "lodash";

/*───────────────────────────────────────────────────────────*
  Normalisation helpers
 *───────────────────────────────────────────────────────────*/
// – basic whitespace/ casing collapse
const squash = (s = "") => s.toString().toLowerCase().replace(/\s+/g, " ").trim();
// – remove CIS level prefixes e.g.  "(L1) " / "(L2) "
const stripLevel = (s = "") => squash(s).replace(/^\(l[0-9]+\)\s*/, "");
// – generic title clean (quotes, punctuation)
const cleanTitle = (s = "") => stripLevel(s).replace(/[“”"']/g, "").replace(/[^a-z0-9 ]+/g, " ").replace(/\s+/g, " ").trim();

/*───────────────────────────────────────────────────────────*
  Title similarity (fast) – exact after clean, else cosine
 *───────────────────────────────────────────────────────────*/
const cosine = (a, b) => {
  if (!a || !b) return 0;
  if (a === b) return 1;
  const w = [...new Set([...a.split(" "), ...b.split(" ")])];
  const v1 = w.map((t) => a.split(" ").filter((x) => x === t).length);
  const v2 = w.map((t) => b.split(" ").filter((x) => x === t).length);
  const dot = v1.reduce((p, c, i) => p + c * v2[i], 0);
  const mag = (v) => Math.sqrt(v.reduce((p, c) => p + c * c, 0));
  const m1 = mag(v1);
  const m2 = mag(v2);
  return !m1 || !m2 ? 0 : dot / (m1 * m2);
};

/*───────────────────────────────────────────────────────────*
  File‑reader tiny wrappers
 *───────────────────────────────────────────────────────────*/
const readBuf = (f) => new Promise((res, rej) => { const r = new FileReader(); r.onload = (e) => res(e.target.result); r.onerror = rej; r.readAsArrayBuffer(f); });
const readTxt = (f) => new Promise((res, rej) => { const r = new FileReader(); r.onload = (e) => res(e.target.result); r.onerror = rej; r.readAsText(f); });

/*───────────────────────────────────────────────────────────*
  Component
 *───────────────────────────────────────────────────────────*/
export default function BSSCISAnalyzer() {
  const [analysis, setAnalysis] = useState(null);
  const [loading, setLoading] = useState(false);

  /* inject mini css once */
  useEffect(() => {
    if (document.getElementById("bss-css")) return;
    const s = document.createElement("style");
    s.id = "bss-css";
    s.innerHTML = `
      :root{--g:1rem}
      body{margin:0;font-family:system-ui,Segoe UI,Roboto,Helvetica,Arial,sans-serif;background:#f8fafc}
      h1{margin:0 0 1rem 0;font-size:1.7rem}
      .grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(165px,1fr));gap:var(--g)}
      .card{background:#fff;border-radius:8px;border:2px solid var(--clr);padding:1rem;text-align:center;box-shadow:0 1px 4px rgba(0,0,0,.05)}
      .card h3{margin:.1rem 0 .4rem;color:#64748b;font-size:.8rem;font-weight:600}
      .card .v{font-size:1.8rem;font-weight:700;color:var(--clr)}
      table{width:100%;border-collapse:collapse;font-size:.8rem;margin-top:.5rem}
      th,td{padding:.5rem .65rem;border-bottom:1px solid #e2e8f0;text-align:left}
      th{background:#f1f5f9;font-weight:600}
      .pill{font-size:.65rem;color:#fff;padding:2px 6px;border-radius:4px}
      button.primary{padding:.55rem 1rem;background:#2563eb;color:#fff;border:none;border-radius:6px;cursor:pointer}
      button.export{background:#16a34a;margin-top:var(--g)}
    `;
    document.head.appendChild(s);
  }, []);

  /*────────────────── main processing ──────────────────*/
  const run = async (bssF, cisF) => {
    setLoading(true);
    try {
      /* BSS workbook */
      const wb = XLSX.read(await readBuf(bssF));
      const sName = wb.SheetNames.find((n) => /settings|windows/i.test(n)) || wb.SheetNames[0];
      const rows = XLSX.utils.sheet_to_json(wb.Sheets[sName], { header: 1, defval: "" });
      const hdrIdx = rows.findIndex((r) => r.some((c) => squash(c).includes("cis #")));
      if (hdrIdx === -1) throw new Error("Could not locate BSS header row");
      const hdrs = rows[hdrIdx].map((h) => squash(h));
      const bss = rows.slice(hdrIdx + 1).filter((r) => r.some(Boolean)).map((r) => {
        const o = {}; hdrs.forEach((h, i) => (o[h] = r[i])); return o;
      });

      /* CIS csv */
      const text = await readTxt(cisF);
      const arr = text.split(/\r?\n/);
      const start = arr.findIndex((l) => l.includes("check_id"));
      const cisArr = Papa.parse(arr.slice(start).join("\n"), { header: true, dynamicTyping: true, skipEmptyLines: true }).data;
      const cisMap = new Map(cisArr.map((c) => [squash(c.check_id), c]));

      /* merge */
      const merged = [];
      bss.forEach((b) => {
        const id = b["cis #"];
        let cis = cisMap.get(squash(id));
        if (!cis) {
          // fallback: compare cleaned titles sans (Lx)
          const bTitle = cleanTitle(b["synapxe setting title"] || b["cis setting title (for reference only)"]);
          cis = cisArr.find((c) => cosine(bTitle, cleanTitle(c.title)) > 0.93);
        }
        merged.push(makeRec(b, cis));
      });
      cisArr.forEach((c) => {
        if (!merged.find((m) => squash(m.CIS_ID) === squash(c.check_id))) merged.push(makeRec(null, c));
      });

      /* summary + cat stats */
      const calc = (cond) => merged.filter(cond).length;
      const summary = {
        total: merged.length,
        bssOnly: calc((m) => m.BSS_ID && !m.CIS_ID),
        cisOnly: calc((m) => !m.BSS_ID && m.CIS_ID),
        both: calc((m) => m.BSS_ID && m.CIS_ID),
        remarks: calc((m) => m["Change Description / Remarks"]),
        exceptions: calc((m) => m["Synapxe Exceptions"]),
        failed: calc((m) => m.Compliance === "Fail"),
        passed: calc((m) => m.Compliance === "Pass"),
        skipped: calc((m) => m.Compliance === "Skipped"),
      };
      const comp = _.mapValues(_.groupBy(merged, "BSS_Category"), (arr) => {
        const p = arr.filter((i) => i.Compliance === "Pass").length;
        return { total: arr.length, passed: p, failed: arr.length - p, rate: ((p / arr.length) * 100 || 0).toFixed(1) };
      });

      setAnalysis({ merged, summary, comp });
    } catch (e) { alert(e.message); console.error(e); }
    setLoading(false);
  };

  /*──────── build record ───────*/
  const makeRec = (b, c) => ({
    BSS_ID: b ? b["cis #"] || "" : "",
    CIS_ID: c ? c.check_id || "" : "",
    BSS_Title: b ? b["synapxe setting title"] || b["cis setting title (for reference only)"] || "" : "",
    CIS_Title: c ? c.title || "" : "",
    Title_Match: cleanTitle(b ? b["synapxe setting title"] || "" : "") === cleanTitle(c ? c.title || "" : "") ? "Yes" : "No",
    BSS_Category: b ? b.category || b["cis section header"] || "Uncategorised" : "Uncategorised",
    "Synapxe Value": b ? b["synapxe value"] || "" : "",
    "Synapxe Exceptions": b ? b["synapxe exceptions"] || "" : "",
    "CIS Recommended Value": b ? b["cis recommended value (for reference only)"] || "" : "",
    "Setting Applicability": b ? b["setting applicability"] || "" : "",
    "Change Description / Remarks": b ? b["change description / remarks"] || "" : "",
    Passed: c ? c.passed_instances : "",
    Failed: c ? c.failed_instances : "",
    Compliance: decideComp(c),
  });

  const decideComp = (c) => {
    if (!c) return "Not Scanned";
    if (c.failed_instances && c.failed_instances !== "None") return "Fail";
    if (c.passed_instances && c.passed_instances !== "None") return "Pass";
    return "Skipped";
  };

  /*──────── Excel export ───────*/
  const xport = () => {
    if (!analysis) return;
    const { merged, summary } = analysis;
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(merged), "Full Comparison");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(Object.entries(summary).map(([k,v])=>({Metric:k,Value:v}))), "Summary");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(merged.filter((m)=>m.Title_Match==="No")), "Title Mismatch");
    XLSX.writeFile(wb, `BSS_CIS_${new Date().toISOString().slice(0,10)}.xlsx`);
  };

  /*──────── small UI bits ──────*/
  const card = (t,v,c) => <div className="card" style={{"--clr":c}}><h3>{t}</h3><div className="v">{v}</div></div>;
  const clr = (r) => r>=90?"#16a34a":r>=70?"#fbbf24":r>=50?"#fb923c":"#ef4444";
  const pick = id => document.getElementById(id)?.files?.[0];

  return (
    <div style={{maxWidth:1200,margin:"0 auto",padding:"1.4rem"}}>
      <h1>Advanced BSS‑CIS Compliance Analyzer</h1>
      <div style={{display:"flex",flexWrap:"wrap",gap:"1rem",alignItems:"center",marginBottom:"var(--g)"}}>
        <label>BSS Excel:&nbsp;<input type="file" id="bss" accept=".xlsx,.xls"/></label>
        <label>CIS CSV:&nbsp;<input type="file" id="cis" accept=".csv"/></label>
        <button className="primary" disabled={loading} onClick={()=>{
          const b=pick("bss"); const c=pick("cis"); if(!b||!c) return alert("select both files"); run(b,c);
        }}>{loading?"Processing…":"Analyze Files"}</button>
      </div>
      {analysis && (<>
        <div className="grid" style={{marginBottom:"var(--g)"}}>
          {card("Total",analysis.summary.total,"#2563eb")}
          {card("BSS Only",analysis.summary.bssOnly,"#0ea5e9")}
          {card("CIS Only",analysis.summary.cisOnly,"#7e22ce")}
          {card("Title Mismatch",analysis.summary.total-analysis.summary.passed-analysis.summary.failed-analysis.summary.skipped,"#c026d3")}
          {card("Failed",analysis.summary.failed,"#ef4444")}
          {card("Passed",analysis.summary.passed,"#16a34a")}
        </div>
        <h2>Compliance by Category</h2>
        <div style={{overflowX:"auto"}}>
          <table>
            <thead><tr><th>Category</th><th>Total</th><th>Passed</th><th>Failed</th><th>Rate</th><th>Status</th></tr></thead>
            <tbody>{Object.entries(analysis.comp).map(([cat,s])=>(
              <tr key={cat}><td>{cat}</td><td>{s.total}</td><td>{s.passed}</td><td>{s.failed}</td><td>{s.rate}%</td><td><span className="pill" style={{background:clr(+s.rate)}}>{+s.rate>=90?"Excellent":+s.rate>=70?"Good":+s.rate>=50?"Fair":"Critical"}</span></td></tr>
            ))}</tbody>
          </table>
        </div>
        <button className="primary export" onClick={xport}>Export to Excel</button>
      </>)}
    </div>
  );
}
