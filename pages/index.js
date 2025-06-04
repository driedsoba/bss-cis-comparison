import React, { useState, useEffect } from "react";
import * as XLSX from "xlsx";
import Papa from "papaparse";
import _ from "lodash";

/*───────────────────────────────────────────────────────────*
  Utility helpers
 *───────────────────────────────────────────────────────────*/
const normalise = (s = "") => s.toString().toLowerCase().replace(/\s+/g, " ").trim();

// ── similarity helpers (only used as a fallback when IDs miss) ──
const getLev = (a, b) => {
  if (!a || !b) return 0;
  const m = Array.from({ length: b.length + 1 }, (_, i) =>
    Array(a.length + 1).fill(i)
  );
  for (let j = 0; j <= a.length; j++) m[0][j] = j;
  for (let i = 1; i <= b.length; i++) {
    for (let j = 1; j <= a.length; j++) {
      m[i][j] = Math.min(
        m[i - 1][j] + 1,
        m[i][j - 1] + 1,
        m[i - 1][j - 1] + (a[j - 1] === b[i - 1] ? 0 : 1)
      );
    }
  }
  const dist = m[b.length][a.length];
  return 1 - dist / Math.max(a.length, b.length);
};
const sim = (a, b) => {
  const s1 = normalise(a);
  const s2 = normalise(b);
  if (!s1 || !s2) return 0;
  if (s1 === s2) return 1;
  // quick bag‑of‑words cosine
  const words = [...new Set([...s1.split(" "), ...s2.split(" ")])];
  const v1 = words.map((w) => s1.split(" ").filter((x) => x === w).length);
  const v2 = words.map((w) => s2.split(" ").filter((x) => x === w).length);
  const dot = v1.reduce((p, c, i) => p + c * v2[i], 0);
  const mag = (v) => Math.sqrt(v.reduce((p, c) => p + c * c, 0));
  const cos = !mag(v1) || !mag(v2) ? 0 : dot / (mag(v1) * mag(v2));
  return 0.4 * getLev(s1, s2) + 0.6 * cos;
};

const readArrBuf = (f) =>
  new Promise((res, rej) => {
    const r = new FileReader();
    r.onload = (e) => res(e.target.result);
    r.onerror = rej;
    r.readAsArrayBuffer(f);
  });
const readText = (f) =>
  new Promise((res, rej) => {
    const r = new FileReader();
    r.onload = (e) => res(e.target.result);
    r.onerror = rej;
    r.readAsText(f);
  });

/*───────────────────────────────────────────────────────────*
  Main component
 *───────────────────────────────────────────────────────────*/
export default function BSSCISAnalyzer() {
  const [analysis, setAnalysis] = useState(null);
  const [loading, setLoading] = useState(false);

  // inject minimal css once
  useEffect(() => {
    if (document.getElementById("bss-cis-css")) return;
    const s = document.createElement("style");
    s.id = "bss-cis-css";
    s.innerHTML = `
      :root{--gap:1rem}
      body{font-family:system-ui,Segoe UI,Roboto,Helvetica,Arial,sans-serif;background:#f9fafb;margin:0}
      h1{margin:0 0 1rem 0;font-size:1.75rem}
      .grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(170px,1fr));gap:var(--gap)}
      .card{background:#fff;border-radius:8px;border:2px solid var(--clr);padding:1rem;text-align:center;box-shadow:0 2px 4px rgba(0,0,0,.05)}
      .card h3{margin:.25rem 0 .5rem;color:#666;font-size:.85rem;font-weight:600}
      .card .val{font-size:1.8rem;font-weight:700;color:var(--clr)}
      table{width:100%;border-collapse:collapse;font-size:.85rem}
      th,td{padding:.5rem .75rem;border-bottom:1px solid #ececec;text-align:left}
      th{background:#f1f5f9;font-weight:600}
      .status{font-size:.68rem;color:#fff;padding:2px 6px;border-radius:4px}
      button.primary{padding:.55rem 1rem;background:#1976d2;color:#fff;border:none;border-radius:6px;cursor:pointer}
      button.export{background:#4caf50;margin-top:var(--gap)}
    `;
    document.head.appendChild(s);
  }, []);

  /*────────────────────  processing  ────────────────────*/
  const process = async (bssFile, cisFile) => {
    setLoading(true);
    try {
      /* BSS Excel */
      const wb = XLSX.read(await readArrBuf(bssFile));
      const sheetName = wb.SheetNames.find((n) => /settings|windows/i.test(n)) || wb.SheetNames[0];
      const rows = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header: 1, defval: "" });
      const hdrRow = rows.findIndex((r) => r.some((c) => normalise(c).includes("cis #")));
      if (hdrRow === -1) throw new Error("Header row not found in BSS sheet");
      const headers = rows[hdrRow].map((h) => normalise(h));
      const bss = rows.slice(hdrRow + 1).filter((r) => r.some(Boolean)).map((r) => {
        const o = {};
        headers.forEach((h, i) => (o[h] = r[i]));
        return o;
      });

      /* CIS CSV */
      const txt = await readText(cisFile);
      const lines = txt.split(/\r?\n/);
      const start = lines.findIndex((l) => l.includes("check_id"));
      const cisArr = Papa.parse(lines.slice(start).join("\n"), { header: true, dynamicTyping: true, skipEmptyLines: true }).data;
      const cisMap = new Map(cisArr.map((c) => [normalise(c.check_id), c]));

      /* merge */
      const merged = [];
      bss.forEach((b) => {
        const id = b["cis #"];
        const cis = cisMap.get(normalise(id)) || cisArr.find((c) => sim(b["synapxe setting title"], c.title) > 0.92);
        merged.push(buildRecord(b, cis));
      });
      cisArr.forEach((c) => {
        if (!merged.find((m) => normalise(m.CIS_ID) === normalise(c.check_id))) merged.push(buildRecord(null, c));
      });

      /* metrics */
      const bssOnly = merged.filter((m) => m.BSS_ID && !m.CIS_ID).length;
      const cisOnly = merged.filter((m) => !m.BSS_ID && m.CIS_ID).length;
      const both = merged.length - bssOnly - cisOnly;
      const remarksCnt = merged.filter((m) => m["Change Description / Remarks"]).length;
      const excCnt = merged.filter((m) => m["Synapxe Exceptions"]).length;
      const failed = merged.filter((m) => m.Compliance === "Fail").length;
      const passed = merged.filter((m) => m.Compliance === "Pass").length;
      const skipped = merged.filter((m) => m.Compliance === "Skipped").length;

      /* summary object */
      const summary = { total: merged.length, bssOnly, cisOnly, both, remarksCnt, excCnt, failed, passed, skipped };

      /* compliance by cat */
      const byCat = _.groupBy(merged, "BSS_Category");
      const comp = {};
      Object.entries(byCat).forEach(([c, items]) => {
        const p = items.filter((i) => i.Compliance === "Pass").length;
        const f = items.filter((i) => i.Compliance === "Fail").length;
        comp[c] = { total: items.length, passed: p, failed: f, rate: ((p / items.length) * 100 || 0).toFixed(1) };
      });

      setAnalysis({ merged, summary, comp });
    } catch (e) {
      alert(e.message);
      console.error(e);
    }
    setLoading(false);
  };

  /* helper build */
  const buildRecord = (b, c) => {
    const obj = {
      BSS_ID: b ? b["cis #"] || "" : "",
      CIS_ID: c ? c.check_id || "" : "",
      BSS_Title: b ? b["synapxe setting title"] || b["cis setting title (for reference only)"] || "" : "",
      CIS_Title: c ? c.title || "" : "",
      BSS_Category: b ? b.category || b["cis section header"] || "Uncategorised" : "Uncategorised",
      "Synapxe Value": b ? b["synapxe value"] || "" : "",
      "Synapxe Exceptions": b ? b["synapxe exceptions"] || "" : "",
      "CIS Recommended Value": b ? b["cis recommended value (for reference only)"] || "" : "",
      "Setting Applicability": b ? b["setting applicability"] || "" : "",
      "Change Description / Remarks": b ? b["change description / remarks"] || "" : "",
      CIS_Level: c ? c.level || "" : "",
      Passed_Instances: c ? c.passed_instances || "" : "",
      Failed_Instances: c ? c.failed_instances || "" : "",
      Compliance: deriveCompliance(c),
    };
    return obj;
  };

  const deriveCompliance = (c) => {
    if (!c) return "Not Scanned";
    if (c.failed_instances && c.failed_instances !== "None" && c.failed_instances !== "") return "Fail";
    if (c.passed_instances && c.passed_instances !== "None" && c.passed_instances !== "") return "Pass";
    return "Skipped";
  };

  /*────────────── export ──────────────*/
  const exportExcel = () => {
    if (!analysis) return;
    const wb = XLSX.utils.book_new();
    const { merged, summary } = analysis;
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(merged), "Full Comparison");

    // summary sheet
    const sumArr = [
      { Metric: "Total Unique Controls", Value: summary.total },
      { Metric: "Controls in BSS Only", Value: summary.bssOnly },
      { Metric: "Controls in CIS Only", Value: summary.cisOnly },
      { Metric: "Controls in Both", Value: summary.both },
      { Metric: "Controls with Remarks", Value: summary.remarksCnt },
      { Metric: "Controls with Exceptions", Value: summary.excCnt },
      { Metric: "Failed Controls", Value: summary.failed },
      { Metric: "Passed Controls", Value: summary.passed },
      { Metric: "Skipped Controls", Value: summary.skipped },
    ];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(sumArr), "Summary");

    // remarks
    const remarks = merged.filter((m) => m["Change Description / Remarks"]);
    if (remarks.length)
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(remarks), "Controls with Remarks");

    // exceptions
    const exc = merged.filter((m) => m["Synapxe Exceptions"]);
    if (exc.length)
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(exc), "Controls with Exceptions");

    // non‑compliant
    const ncf = merged.filter((m) => m.Compliance === "Fail");
    if (ncf.length)
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(ncf), "Non‑Compliant");

    XLSX.writeFile(wb, `BSS_CIS_Comparison_${new Date().toISOString().slice(0, 10)}.xlsx`);
  };

  /*────────────── render ──────────────*/
  const statusClr = (rate) => (rate >= 90 ? "#4caf50" : rate >= 70 ? "#ffb74d" : rate >= 50 ? "#ff7043" : "#e53935");
  const card = (t, v, c) => (
    <div className="card" style={{ "--clr": c }}>
      <h3>{t}</h3>
      <div className="val">{v}</div>
    </div>
  );
  const pick = (id) => document.getElementById(id)?.files?.[0];

  return (
    <div style={{ maxWidth: 1200, margin: "0 auto", padding: "1.5rem" }}>
      <h1>Advanced BSS‑CIS Compliance Analyzer</h1>
      <div style={{ display: "flex", flexWrap: "wrap", gap: "1rem", alignItems: "center", marginBottom: "var(--gap)" }}>
        <label>
          BSS Excel:&nbsp;
          <input type="file" id="bss" accept=".xlsx,.xls" />
        </label>
        <label>
          CIS CSV:&nbsp;
          <input type="file" id="cis" accept=".csv" />
        </label>
        <button className="primary" disabled={loading} onClick={() => {
          const b = pick("bss");
          const c = pick("cis");
          if (!b || !c) return alert("Please select both files");
          process(b, c);
        }}>{loading ? "Processing…" : "Analyze Files"}</button>
      </div>

      {analysis && (
        <>
          <div className="grid" style={{ marginBottom: "var(--gap)" }}>
            {card("Total Controls", analysis.summary.total, "#1976d2")}
            {card("BSS Only", analysis.summary.bssOnly, "#0288d1")}
            {card("CIS Only", analysis.summary.cisOnly, "#7b1fa2")}
            {card("Remarks", analysis.summary.remarksCnt, "#6d4c41")}
            {card("Failed", analysis.summary.failed, "#e53935")}
            {card("Passed", analysis.summary.passed, "#43a047")}
          </div>

          <h2>Compliance by Category</h2>
          <div style={{ overflowX: "auto" }}>
            <table>
              <thead>
                <tr>
                  <th>Category</th>
                  <th>Total</th>
                  <th>Passed</th>
                  <th>Failed</th>
                  <th>Pass Rate</th>
                  <th>Status</th>
                </tr>
              </thead>
              <tbody>
                {Object.entries(analysis.comp).map(([cat, s]) => (
                  <tr key={cat}>
                    <td>{cat}</td>
                    <td>{s.total}</td>
                    <td>{s.passed}</td>
                    <td>{s.failed}</td>
                    <td>{s.rate}%</td>
                    <td><span className="status" style={{ background: statusClr(+s.rate) }}>{+s.rate >= 90 ? "Excellent" : +s.rate >= 70 ? "Good" : +s.rate >= 50 ? "Fair" : "Critical"}</span></td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>

          <button className="primary export" onClick={exportExcel}>Export to Excel</button>
        </>
      )}
    </div>
  );
}
