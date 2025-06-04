import React, { useState, useEffect } from "react";
import * as XLSX from "xlsx";
import Papa from "papaparse";
import _ from "lodash";

/*───────────────────────────────────────────────────────────*
  Utility helpers
 *───────────────────────────────────────────────────────────*/
const normalise = (s = "") =>
  s.toString().toLowerCase().replace(/\s+/g, " ").trim();

// ── similarity helpers (only used as a fuzzy‑fallback) ─────────
const getLevenshteinSimilarity = (s1, s2) => {
  const a = s1.length;
  const b = s2.length;
  if (!a || !b) return 0;
  const matrix = Array.from({ length: b + 1 }, () => Array(a + 1).fill(0));
  for (let i = 0; i <= a; i++) matrix[0][i] = i;
  for (let j = 0; j <= b; j++) matrix[j][0] = j;
  for (let j = 1; j <= b; j++) {
    for (let i = 1; i <= a; i++) {
      const cost = s1[i - 1] === s2[j - 1] ? 0 : 1;
      matrix[j][i] = Math.min(
        matrix[j - 1][i] + 1,
        matrix[j][i - 1] + 1,
        matrix[j - 1][i - 1] + cost
      );
    }
  }
  const distance = matrix[b][a];
  return 1 - distance / Math.max(a, b);
};
const getJaccardSimilarity = (s1, s2) => {
  const set1 = new Set(s1.split(/\s+/));
  const set2 = new Set(s2.split(/\s+/));
  const inter = [...set1].filter((x) => set2.has(x)).length;
  const union = new Set([...set1, ...set2]).size;
  return union === 0 ? 0 : inter / union;
};
const getCosineSimilarity = (s1, s2) => {
  const words = [...new Set([...s1.split(/\s+/), ...s2.split(/\s+/)])];
  const v1 = words.map((w) => s1.split(/\s+/).filter((x) => x === w).length);
  const v2 = words.map((w) => s2.split(/\s+/).filter((x) => x === w).length);
  const dot = v1.reduce((sum, v, i) => sum + v * v2[i], 0);
  const mag1 = Math.sqrt(v1.reduce((sum, v) => sum + v * v, 0));
  const mag2 = Math.sqrt(v2.reduce((sum, v) => sum + v * v, 0));
  return mag1 && mag2 ? dot / (mag1 * mag2) : 0;
};
const calculateSimilarity = (a, b) => {
  const s1 = normalise(a);
  const s2 = normalise(b);
  if (!s1 || !s2) return 0;
  if (s1 === s2) return 1;
  const lev = getLevenshteinSimilarity(s1, s2);
  const jac = getJaccardSimilarity(s1, s2);
  const cos = getCosineSimilarity(s1, s2);
  return 0.4 * lev + 0.3 * jac + 0.3 * cos;
};

const readFileAsArrayBuffer = (file) =>
  new Promise((res, rej) => {
    const r = new FileReader();
    r.onload = (e) => res(e.target.result);
    r.onerror = rej;
    r.readAsArrayBuffer(file);
  });
const readFileAsText = (file) =>
  new Promise((res, rej) => {
    const r = new FileReader();
    r.onload = (e) => res(e.target.result);
    r.onerror = rej;
    r.readAsText(file);
  });

/*───────────────────────────────────────────────────────────*
  Main component
 *───────────────────────────────────────────────────────────*/
export default function BSSCISAnalyzer() {
  const [analysis, setAnalysis] = useState(null);
  const [loading, setLoading] = useState(false);

  /* inject a tiny stylesheet once */
  useEffect(() => {
    const id = "bss-cis-style";
    if (document.getElementById(id)) return;
    const style = document.createElement("style");
    style.id = id;
    style.innerHTML = `
      body {font-family: system-ui,Segoe UI,Roboto,Helvetica,Arial,sans-serif;background:#fafafa;margin:0}
      h1,h2,h3{margin:0}
      .grid-cards{display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:1rem}
      .card{background:#fff;border-radius:8px;border:2px solid var(--clr);padding:1rem;text-align:center;box-shadow:0 2px 4px rgba(0,0,0,.05)}
      .card h3{font-size:.9rem;color:#666;font-weight:500;margin-bottom:.3rem}
      .card .val{font-size:2rem;font-weight:700;color:var(--clr)}
      table{width:100%;border-collapse:collapse;margin-top:.5rem}
      th,td{padding:.5rem .75rem;border-bottom:1px solid #eee;text-align:left;font-size:.85rem}
      th{background:#f2f2f2;font-weight:600}
      .status{color:#fff;font-size:.7rem;padding:2px 6px;border-radius:4px}
    `;
    document.head.appendChild(style);
  }, []);

  /*────────────────────  core logic  ────────────────────*/
  const processFiles = async (bssFile, cisFile) => {
    setLoading(true);
    try {
      /* 1. ── read BSS excel ─────────────────────────── */
      const bssBuffer = await readFileAsArrayBuffer(bssFile);
      const wb = XLSX.read(bssBuffer);
      const sheetName = wb.SheetNames.find((n) =>
        /settings|windows/i.test(n)
      ) || wb.SheetNames[0];
      const sheet = wb.Sheets[sheetName];
      const raw = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });

      // locate header row robustly
      const headerRowIdx = raw.findIndex((row) =>
        row.some((c) => normalise(c).includes("cis #"))
      );
      if (headerRowIdx === -1) throw new Error("Header row not found in BSS sheet");
      const headers = raw[headerRowIdx];
      const bssRows = raw.slice(headerRowIdx + 1).filter((r) => r.some(Boolean));
      const bssData = bssRows.map((row) => {
        const obj = {};
        headers.forEach((h, i) => {
          obj[normalise(h)] = row[i];
        });
        return obj;
      });

      /* 2. ── read CIS csv ───────────────────────────── */
      const cisText = await readFileAsText(cisFile);
      const lines = cisText.split(/\r?\n/);
      const hdrIdx = lines.findIndex((l) => l.includes("check_id"));
      const dataText = lines.slice(hdrIdx).join("\n");
      const cisData = Papa.parse(dataText, {
        header: true,
        dynamicTyping: true,
        skipEmptyLines: true,
      }).data;

      /* 3. ── merge ──────────────────────────────────── */
      const cisMap = new Map(
        cisData.map((c) => [normalise(c.check_id), c])
      );
      const merged = [];
      bssData.forEach((b) => {
        const bssId = b["cis #"];
        const cis = cisMap.get(normalise(bssId));
        let cisRecord = cis;

        // fallback fuzzy match if missing
        if (!cisRecord) {
          cisRecord = cisData.find(
            (c) => calculateSimilarity(b["synapxe setting title"], c.title) > 0.94
          );
        }

        merged.push({
          BSS_ID: bssId || "",
          CIS_ID: cisRecord?.check_id || "",
          BSS_Title: b["synapxe setting title"] || b["cis setting title (for reference only)"] || "",
          CIS_Title: cisRecord?.title || "",
          BSS_Category: b.category || b["cis section header"] || "Uncategorised",
          BSS_Expected_Value:
            b["synapxe value"] || b["cis recommended value (for reference only)"] || "",
          CIS_Value: extractCISValue(cisRecord),
          CIS_Status: deriveStatus(cisRecord),
          Compliance: deriveCompliance(cisRecord),
        });
      });

      // add CIS‑only rows
      cisData.forEach((c) => {
        if (!merged.find((m) => normalise(m.CIS_ID) === normalise(c.check_id))) {
          merged.push({
            BSS_ID: "",
            CIS_ID: c.check_id,
            BSS_Title: "",
            CIS_Title: c.title,
            BSS_Category: "Uncategorised",
            BSS_Expected_Value: "",
            CIS_Value: extractCISValue(c),
            CIS_Status: deriveStatus(c),
            Compliance: "No BSS Mapping",
          });
        }
      });

      /* 4. ── extra consistency checks ───────────────── */
      const titleGroups = _.groupBy(merged, (r) => normalise(r.CIS_Title || r.BSS_Title));
      const dupTitleDiffId = Object.values(titleGroups)
        .filter((g) => _.uniqBy(g, "CIS_ID").length > 1)
        .flat();
      const idGroups = _.groupBy(merged, (r) => normalise(r.CIS_ID || r.BSS_ID));
      const dupIdDiffTitle = Object.values(idGroups)
        .filter((g) =>
          _.uniqBy(g, (r) => normalise(r.CIS_Title || r.BSS_Title)).length > 1
        )
        .flat();

      /* 5. ── compliance summary & risk ─────────────── */
      const byCat = _.groupBy(merged, "BSS_Category");
      const complianceByCategory = {};
      Object.entries(byCat).forEach(([cat, items]) => {
        const passed = items.filter((i) => i.Compliance === "Pass").length;
        const failed = items.filter((i) => i.Compliance === "Fail").length;
        complianceByCategory[cat] = {
          total: items.length,
          passed,
          failed,
          passRate: ((passed / items.length) * 100).toFixed(1),
        };
      });

      const summary = {
        total: merged.length,
        passed: merged.filter((d) => d.Compliance === "Pass").length,
        failed: merged.filter((d) => d.Compliance === "Fail").length,
        skipped: merged.filter((d) => d.Compliance === "Skipped").length,
        dupTitle: dupTitleDiffId.length,
        dupId: dupIdDiffTitle.length,
      };

      const riskScore = (() => {
        const weights = { Fail: 1, Pass: 0, Skipped: 0.5 };
        const total = merged.length;
        const risk =
          merged.reduce((sum, r) => sum + (weights[r.Compliance] || 0), 0) / total;
        return Math.round(100 * (1 - risk));
      })();

      setAnalysis({ merged, complianceByCategory, summary, riskScore, dupTitleDiffId, dupIdDiffTitle });
    } catch (err) {
      alert(`Error: ${err.message}`);
      console.error(err);
    }
    setLoading(false);
  };

  /*─────────────  tiny render helpers  ─────────────*/
  const getStatusColour = (rate) => {
    if (rate >= 90) return "#4caf50";
    if (rate >= 70) return "#ffb300";
    if (rate >= 50) return "#ff7043";
    return "#e53935";
  };

  const quickFile = (id) => document.getElementById(id)?.files?.[0];

  /*───────────────────────── render ─────────────────────────*/
  return (
    <div style={{ padding: "1.5rem", maxWidth: 1200, margin: "0 auto" }}>
      <h1 style={{ marginBottom: "1rem" }}>Advanced BSS‑CIS Compliance Analyzer</h1>

      {/* file pickers */}
      <div style={{ display: "flex", gap: "1rem", flexWrap: "wrap", marginBottom: 16 }}>
        <div>
          <label style={{ fontWeight: 600 }}>BSS Excel: </label>
          <input id="bssFile" type="file" accept=".xls,.xlsx" />
        </div>
        <div>
          <label style={{ fontWeight: 600 }}>CIS CSV: </label>
          <input id="cisFile" type="file" accept=".csv" />
        </div>
        <button
          disabled={loading}
          onClick={() => {
            const bss = quickFile("bssFile");
            const cis = quickFile("cisFile");
            if (!bss || !cis) return alert("Please select both files.");
            processFiles(bss, cis);
          }}
          style={{
            padding: "0.6rem 1.2rem",
            background: "#1976d2",
            color: "#fff",
            border: "none",
            borderRadius: 6,
            cursor: loading ? "wait" : "pointer",
          }}
        >
          {loading ? "Processing…" : "Analyze Files"}
        </button>
      </div>

      {/* dashboard */}
      {analysis && (
        <>
          <div className="grid-cards" style={{ marginBottom: 24 }}>
            <Card title="Total Controls" value={analysis.summary.total} clr="#1976d2" />
            <Card title="Passed" value={analysis.summary.passed} clr="#4caf50" />
            <Card title="Failed" value={analysis.summary.failed} clr="#e53935" />
            <Card title="Same Title · Diff ID" value={analysis.summary.dupTitle} clr="#673ab7" />
            <Card title="Same ID · Diff Title" value={analysis.summary.dupId} clr="#009688" />
            <Card title="Risk Score" value={analysis.riskScore + "%"} clr="#ff9800" />
          </div>

          {/* compliance table */}
          <h2>Compliance by Category</h2>
          <div style={{ overflowX: "auto" }}>
            <table>
              <thead>
                <tr>
                  <th>Category</th>
                  <th>Total</th>
                  <th>Passed</th>
                  <th>Failed</th>
                  <th>Pass Rate</th>
                  <th>Status</th>
                </tr>
              </thead>
              <tbody>
                {Object.entries(analysis.complianceByCategory).map(([cat, s]) => (
                  <tr key={cat}>
                    <td>{cat}</td>
                    <td>{s.total}</td>
                    <td>{s.passed}</td>
                    <td>{s.failed}</td>
                    <td>{s.passRate}%</td>
                    <td>
                      <span
                        className="status"
                        style={{ background: getStatusColour(Number(s.passRate)) }}
                      >
                        {Number(s.passRate) >= 90
                          ? "Excellent"
                          : Number(s.passRate) >= 70
                          ? "Good"
                          : Number(s.passRate) >= 50
                          ? "Fair"
                          : "Critical"}
                      </span>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </>
      )}
    </div>
  );
}

/*───────────────────────── small sub‑component ───────────*/
function Card({ title, value, clr }) {
  return (
    <div className="card" style={{ "--clr": clr }}>
      <h3>{title}</h3>
      <div className="val">{value}</div>
    </div>
  );
}

/*───────────────────────── helper fns ────────────────────*/
const extractCISValue = (c) => {
  if (!c) return "";
  const has = (v) => v != null && v.toString().toLowerCase() !== "none" && v !== "";
  if (has(c.failed_instances)) return "Non‑Compliant";
  if (has(c.passed_instances)) return "Compliant";
  if (has(c.skipped_instances)) return "Skipped";
  return "Not Scanned";
};
const deriveStatus = (c) => {
  if (!c) return "Not Scanned";
  const has = (v) => v != null && v.toString().toLowerCase() !== "none" && v !== "";
  if (has(c.failed_instances)) return "FAILED";
  if (has(c.passed_instances)) return "PASSED";
  if (has(c.skipped_instances)) return "SKIPPED";
  return "Unknown";
};
const deriveCompliance = (c) => {
  if (!c) return "Not Scanned";
  if (c.skipped_instances) return "Skipped";
  if (c.failed_instances) return "Fail";
  if (c.passed_instances) return "Pass";
  return "Unknown";
};
