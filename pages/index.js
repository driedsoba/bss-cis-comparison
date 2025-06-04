import React, { useState, useEffect } from "react";
import * as XLSX from "xlsx";
import Papa from "papaparse";
import _ from "lodash";

/*───────────────────────────────────────────────────────────*
  Helper functions
 *───────────────────────────────────────────────────────────*/
const squash = (s = "") => s.toString().toLowerCase().replace(/\s+/g, " ").trim();
const stripLevel = (s = "") => squash(s).replace(/^\(l[0-9]+\)\s*/, "");
const cleanTitle = (s = "") =>
  stripLevel(s)
    .replace(/[“”"']/g, "")
    .replace(/[^a-z0-9 ]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();

const cosine = (a, b) => {
  if (!a || !b) return 0;
  if (a === b) return 1;
  const words = [...new Set([...a.split(" "), ...b.split(" ")])];
  const v1 = words.map((w) => a.split(" ").filter((x) => x === w).length);
  const v2 = words.map((w) => b.split(" ").filter((x) => x === w).length);
  const dot = v1.reduce((p, c, i) => p + c * v2[i], 0);
  const mag = (v) => Math.sqrt(v.reduce((p, c) => p + c * c, 0));
  const m1 = mag(v1);
  const m2 = mag(v2);
  return !m1 || !m2 ? 0 : dot / (m1 * m2);
};

const readBuf = (file) =>
  new Promise((res, rej) => {
    const r = new FileReader();
    r.onload = (e) => res(e.target.result);
    r.onerror = rej;
    r.readAsArrayBuffer(file);
  });

const readTxt = (file) =>
  new Promise((res, rej) => {
    const r = new FileReader();
    r.onload = (e) => res(e.target.result);
    r.onerror = rej;
    r.readAsText(file);
  });

/*───────────────────────────────────────────────────────────*
  React Component
 *───────────────────────────────────────────────────────────*/
export default function BSSCISAnalyzer() {
  const [analysis, setAnalysis] = useState(null);
  const [loading, setLoading] = useState(false);

  /* Inject minimal CSS once */
  useEffect(() => {
    if (document.getElementById("bss-cis-css")) return;
    const style = document.createElement("style");
    style.id = "bss-cis-css";
    style.innerHTML = `
      :root{--gap:1rem}
      body{margin:0;font-family:system-ui,Segoe UI,Roboto,Helvetica,Arial,sans-serif;background:#f8fafc}
      h1{margin:0 0 1rem;font-size:1.7rem}
      .grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(165px,1fr));gap:var(--gap)}
      .card{background:#fff;border-radius:8px;border:2px solid var(--clr);padding:1rem;text-align:center;box-shadow:0 1px 4px rgba(0,0,0,.05)}
      .card h3{margin:.1rem 0 .4rem;color:#64748b;font-size:.8rem;font-weight:600}
      .card .v{font-size:1.8rem;font-weight:700;color:var(--clr)}
      table{width:100%;border-collapse:collapse;font-size:.8rem;margin-top:.5rem}
      th,td{padding:.5rem .65rem;border-bottom:1px solid #e2e8f0;text-align:left}
      th{background:#f1f5f9;font-weight:600}
      .pill{font-size:.65rem;color:#fff;padding:2px 6px;border-radius:4px}
      button.primary{padding:.55rem 1rem;background:#2563eb;color:#fff;border:none;border-radius:6px;cursor:pointer}
      button.export{background:#16a34a;margin-top:var(--gap)}
    `;
    document.head.appendChild(style);
  }, []);

  /*──────────────────────────────────────────────────────*/
  /* Core Logic                                           */
  /*──────────────────────────────────────────────────────*/

  const buildCompliance = (cis) => {
    if (!cis) return "Not Scanned";
    if (cis.failed_instances && cis.failed_instances !== "None") return "Fail";
    if (cis.passed_instances && cis.passed_instances !== "None") return "Pass";
    return "Skipped";
  };

  const buildRecord = (bssRow, cisRow) => {
    const getCol = (row, prefix) => {
      if (!row) return "";
      const k = Object.keys(row).find((x) => x.startsWith(prefix));
      return k ? row[k] : "";
    };

    return {
      BSS_ID: bssRow ? bssRow["cis #"] || "" : "",
      CIS_ID: cisRow ? cisRow.check_id || "" : "",
      BSS_Title:
        bssRow ? bssRow["synapxe setting title"] || bssRow["cis setting title (for reference only)"] || "" : "",
      CIS_Title: cisRow ? cisRow.title || "" : "",
      Title_Match:
        cleanTitle(bssRow ? bssRow["synapxe setting title"] || "" : "") ===
        cleanTitle(cisRow ? cisRow.title || "" : "")
          ? "Yes"
          : "No",
      BSS_Category: bssRow ? bssRow.category || bssRow["cis section header"] || "Uncategorised" : "Uncategorised",
      "Synapxe Value": getCol(bssRow, "synapxe value"),
      "Synapxe Exceptions": getCol(bssRow, "synapxe exceptions"),
      "CIS Recommended Value": getCol(bssRow, "cis recommended value"),
      "Setting Applicability": getCol(bssRow, "setting applicability"),
      "Change Description / Remarks": getCol(bssRow, "change description"),
      Passed: cisRow ? cisRow.passed_instances : "",
      Failed: cisRow ? cisRow.failed_instances : "",
      Compliance: buildCompliance(cisRow),
    };
  };

  /* Main processing function */
  const processFiles = async (bssFile, cisFile) => {
    setLoading(true);
    try {
      /* ------- BSS Excel ------- */
      const wb = XLSX.read(await readBuf(bssFile));
      const sheetName = wb.SheetNames.find((n) => /settings|windows/i.test(n)) || wb.SheetNames[0];
      const rawRows = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header: 1, defval: "" });
      const hdrRowIdx = rawRows.findIndex((r) => r.some((c) => squash(c).includes("cis #")));
      if (hdrRowIdx === -1) throw new Error("Header row not found in BSS sheet");
      const headers = rawRows[hdrRowIdx].map((h) => squash(h));
      const bssRows = rawRows.slice(hdrRowIdx + 1).filter((r) => r.some(Boolean)).map((r) => {
        const obj = {};
        headers.forEach((h, i) => (obj[h] = r[i]));
        return obj;
      });

      /* ------- CIS CSV ------- */
      const txt = await readTxt(cisFile);
      const lines = txt.split(/\r?\n/);
      const csvStart = lines.findIndex((l) => l.includes("check_id"));
      const cisRows = Papa.parse(lines.slice(csvStart).join("\n"), {
        header: true,
        dynamicTyping: true,
        skipEmptyLines: true,
      }).data;
      const cisMap = new Map(cisRows.map((c) => [squash(c.check_id), c]));

      /* ------- Merge ------- */
      const merged = [];
      bssRows.forEach((b) => {
        const id = b["cis #"];
        let cis = cisMap.get(squash(id));
        if (!cis) {
          const bTitle = cleanTitle(b["synapxe setting title"] || b["cis setting title (for reference only)"]);
          cis = cisRows.find((c) => cosine(bTitle, cleanTitle(c.title)) > 0.93);
        }
        merged.push(buildRecord(b, cis));
      });
      cisRows.forEach((c) => {
        if (!merged.find((m) => squash(m.CIS_ID) === squash(c.check_id))) {
          merged.push(buildRecord(null, c));
        }
      });

      /* ------- Summary ------- */
      const total = merged.length;
      const bssOnly = merged.filter((m) => m.BSS_ID && !m.CIS_ID).length;
      const cisOnly = merged.filter((m) => !m.BSS_ID && m.CIS_ID).length;
      const both = total - bssOnly - cisOnly;
      const remarksCnt = merged.filter((m) => m["Change Description / Remarks"]).length;
      const excCnt = merged.filter((m) => m["Synapxe Exceptions"]).length;
      const failedCnt = merged.filter((m) => m.Compliance === "Fail").length;
      const passedCnt = merged.filter((m) => m.Compliance === "Pass").length;
      const skippedCnt = merged.filter((m) => m.Compliance === "Skipped").length;

      const summary = {
        total,
        bssOnly,
        cisOnly,
        both,
        remarksCnt,
        excCnt,
        failedCnt,
        passedCnt,
        skippedCnt,
      };

      const comp = _.mapValues(_.groupBy(merged, "BSS_Category"), (arr) => {
        const passed = arr.filter((x) => x.Compliance === "Pass").length;
        return {
          total: arr.length,
          passed,
          failed: arr.length - passed,
          rate: ((passed / arr.length) * 100 || 0).toFixed(1),
        };
      });

      setAnalysis({ merged, summary, comp });
    } catch (err) {
      alert(err.message);
      console.error(err);
    }
    setLoading(false);
  };

  /* ------- Excel Export ------- */
  const exportExcel = () => {
    if (!analysis) return;
    const { merged, summary } = analysis;
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(merged), "Full Comparison");

    const summarySheet = Object.entries(summary).map(([Metric,	Value]) => ({ Metric, Value }));
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(summarySheet), "Summary");

    XLSX.writeFile(wb, `BSS_CIS_${new Date().toISOString().slice(0, 10)}.xlsx`);
  };

  /* ------- UI Helpers ------- */
  const card = (t, v, c) => (
    <div className="card" style={{ "--clr": c }}>
      <h3>{t}</h3>
      <div className="v">{v}</div>
    </div>
  );

  const statusColor = (rate) => (rate >= 90 ? "#16a34a" : rate >= 70 ? "#fbbf24" : rate >= 50 ? "#fb923c" : "#ef4444");

  const fileById = (id) => document.getElementById(id)?.files?.[0];

  /* ------- JSX Return ------- */
  return (
    <div style={{ maxWidth: 1200, margin: "0 auto", padding: "1.5rem" }}>
      <h1>Advanced BSS‑CIS Compliance Analyzer</h1>

      {/* File pickers */}
      <div style={{ display: "flex", flexWrap: "wrap", gap: "1rem", alignItems: "center", marginBottom: "var(--gap)" }}>
        <label>
          BSS Excel:&nbsp;
          <input type="file" id="bssFile" accept=".xlsx,.xls" />
        </label>
        <label>
          CIS CSV:&nbsp;
          <input type="file" id="cisFile" accept=".csv" />
        </label>
        <button
          className="primary"
          disabled={loading}
          onClick={() => {
            const bss = fileById("bssFile");
            const cis = fileById("cisFile");
            if (!bss || !cis) return alert("Please select both files");
            processFiles(bss, cis);
          }}
        >
          {loading ? "Processing…" : "Analyze Files"}
        </button>
      </div>

      {analysis && (
        <>
          {/* Dashboard cards */}
          <div className="grid" style={{ marginBottom: "var(--gap)" }}>
            {card("Total", analysis.summary.total, "#2563eb")}
            {card("BSS Only", analysis.summary.bssOnly, "#0ea5e9")}
            {card("CIS Only", analysis.summary.cisOnly, "#7e22ce")}
            {card("Remarks", analysis.summary.remarksCnt, "#6d4c41")}
            {card("Failed", analysis.summary.failedCnt, "#ef4444")}
            {card("Passed", analysis.summary.passedCnt, "#16a34a")}
          </div>

          {/* Compliance table */}
          <h2>Compliance by Category</h2>
          <div style={{ overflowX: "auto" }}>
            <table>
              <thead>
                <tr>
                  <th>Category</th>
                  <th>Total</th>
                  <th>Passed</th>
                  <th>Failed</th>
                  <th>Rate</th>
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
                    <td>
                      <span className="pill" style={{ background: statusColor(+s.rate) }}>
                        {+s.rate >= 90 ? "Excellent" : +s.rate >= 70 ? "Good" : +s.rate >= 50 ? "Fair" : "Critical"}
                      </span>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>

          {/* Export button */}
          <button className="primary export" onClick={exportExcel}>
            Export to Excel
          </button>
        </>
      )}
    </div>
  );
}
