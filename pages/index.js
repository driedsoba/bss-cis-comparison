import { useState } from 'react';
import * as XLSX from 'xlsx';
import Papa from 'papaparse';

export default function Home() {
  const [bssFile, setBssFile] = useState(null);
  const [cisFile, setCisFile] = useState(null);
  const [results, setResults] = useState(null);
  const [loading, setLoading] = useState(false);

  // Utility: strip any leading "(L<number>) " from a string
  const stripLPrefix = (raw) => {
    return raw.replace(/^\(L\d+\)\s*/, '').trim();
  };

  const readBSSFile = async (file) => {
    const data = await file.arrayBuffer();
    const workbook = XLSX.read(data);

    // Find the sheet whose name contains "settings" or "server"
    const sheetName = Object.keys(workbook.Sheets).find((name) =>
      name.toLowerCase().includes('settings') || name.toLowerCase().includes('server')
    );
    if (!sheetName) throw new Error("Settings sheet not found");

    const worksheet = workbook.Sheets[sheetName];
    // Convert to a 2D array so we can locate "CIS #"
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

    // Find the row index where "CIS #" appears in any cell
    let headerRowIdx = jsonData.findIndex((row) =>
      row.some((cell) => cell && cell.toString().includes('CIS #'))
    );
    if (headerRowIdx === -1) throw new Error("Header row not found");

    const headers = jsonData[headerRowIdx];
    const data_rows = jsonData
      .slice(headerRowIdx + 1)
      .filter((row) => row[headers.indexOf('CIS #')]) // only keep rows where "CIS #" is nonempty
      .map((row) => {
        const obj = {};
        headers.forEach((header, idx) => {
          obj[header.toString().trim()] = row[idx];
        });
        return obj;
      });

    return data_rows;
  };

  const readCISFile = (file) => {
    return new Promise((resolve, reject) => {
      Papa.parse(file, {
        complete: (result) => {
          let data = result.data;
          // Locate the first row that starts with "check_id"
          let headerIdx = data.findIndex((row) =>
            row[0] === 'check_id' || (row.length > 1 && row[0].includes('check_id'))
          );
          if (headerIdx > 0) data = data.slice(headerIdx);

          // Re-parse with header: true
          const parsed = Papa.parse(Papa.unparse(data), {
            header: true,
            skipEmptyLines: true
          });
          resolve(parsed.data);
        },
        error: reject
      });
    });
  };

  const compareData = (bssData, cisData) => {
    const results = [];

    // Build fast lookup maps by ID
    const bssMap = new Map(bssData.map((row) => [row['CIS #']?.toString(), row]));
    const cisMap = new Map(cisData.map((row) => [row.check_id?.toString(), row]));

    const bssIds = new Set(bssData.map((row) => row['CIS #']?.toString()));
    const cisIds = new Set(cisData.map((row) => row.check_id?.toString()));
    const allIds = new Set([...bssIds, ...cisIds]);

    allIds.forEach((checkId) => {
      const bssRow = bssMap.get(checkId);
      const cisRow = cisMap.get(checkId);

      // Initialize our result object
      const result = {
        Check_ID: checkId,
        In_BSS: bssRow ? 'Yes' : 'No',
        In_CIS_Scan: cisRow ? 'Yes' : 'No',
        Non_Compliance_Reason: ''
      };

      // 1) Pull in all BSS fields (and strip "(L#)" from title)
      if (bssRow) {
        const titleCol = Object.keys(bssRow).find((key) => key.includes('Setting Title'));
        const appCol = Object.keys(bssRow).find((key) => key.includes('Setting Applicability'));
        const cisRecCol = Object.keys(bssRow).find((key) =>
          key.includes('CIS Recommended Value')
        );

        const rawBssTitle = bssRow[titleCol] || '';
        const cleanBssTitle = stripLPrefix(rawBssTitle);

        result.BSS_Title = cleanBssTitle;
        result.BSS_Category = bssRow.Category || '';
        result.Setting_Applicability = bssRow[appCol] || '';
        result.CIS_Recommended_Value = bssRow[cisRecCol] || '';
        result.Synapxe_Value = bssRow['Synapxe Value'] || '';
        result.Synapxe_Exceptions = bssRow['Synapxe Exceptions'] || '';
        result.Change_Description_Remarks = bssRow['Change Description / Remarks'] || '';
        result.BSS_ID = bssRow['BSS ID'] || bssRow['BSS #'] || checkId;
      } else {
        // If no BSS row, fill defaults
        result.BSS_Title = '';
        result.BSS_Category = '';
        result.Setting_Applicability = '';
        result.CIS_Recommended_Value = '';
        result.Synapxe_Value = '';
        result.Synapxe_Exceptions = '';
        result.Change_Description_Remarks = '';
        result.BSS_ID = checkId;
      }

      // 2) Pull CIS fields (and strip "(L#)" from CIS title)
      if (cisRow) {
        const rawCisTitle = cisRow.title || '';
        const cleanCisTitle = stripLPrefix(rawCisTitle);

        result.CIS_Title = cleanCisTitle;
        result.CIS_Level = cisRow.level || '';
        result.Failed_Instances = cisRow.failed_instances || '';
        result.Passed_Instances = cisRow.passed_instances || '';
      } else {
        result.CIS_Title = '';
        result.CIS_Level = '';
        result.Failed_Instances = '';
        result.Passed_Instances = '';
      }

      // 3) Determine CIS_Status (Failed / Passed / Skipped)
      if (cisRow) {
        if (
          cisRow.failed_instances &&
          cisRow.failed_instances !== 'None' &&
          cisRow.failed_instances.trim() !== ''
        ) {
          result.CIS_Status = 'Failed';
        } else if (
          cisRow.passed_instances &&
          cisRow.passed_instances !== 'None' &&
          cisRow.passed_instances.trim() !== ''
        ) {
          result.CIS_Status = 'Passed';
        } else {
          result.CIS_Status = 'Skipped';
        }
      } else {
        result.CIS_Status = '';
      }

      // 4) Check for title mismatch (using cleaned titles)
      let titleMismatch = false;
      if (bssRow && cisRow) {
        const bTitle = (result.BSS_Title || '').toString().trim();
        const cTitle = (result.CIS_Title || '').toString().trim();
        if (bTitle !== cTitle) {
          titleMismatch = true;
        }
      }
      result.Title_Mismatch = titleMismatch ? 'Yes' : 'No';

      // 5) Check for any nonempty remark
      const hasRemark =
        bssRow && (result.Change_Description_Remarks || '').toString().trim() !== '';
      result.Has_Remark = hasRemark ? 'Yes' : 'No';

      // 6) Derive final Compliance_Status & Non_Compliance_Reason
      if (bssRow && cisRow) {
        if (result.CIS_Status === 'Failed') {
          result.Compliance_Status = 'Non-Compliant';
          result.Non_Compliance_Reason = 'Scan Failed';
        } else if (result.CIS_Status === 'Passed') {
          if (titleMismatch) {
            result.Compliance_Status = 'Non-Compliant';
            result.Non_Compliance_Reason = 'Title Mismatch';
          } else if (hasRemark) {
            result.Compliance_Status = 'Non-Compliant';
            result.Non_Compliance_Reason = 'Has Remark';
          } else {
            result.Compliance_Status = 'Compliant';
            result.Non_Compliance_Reason = '';
          }
        } else {
          // CIS_Status === 'Skipped'
          result.Compliance_Status = 'Not Tested';
          result.Non_Compliance_Reason = '';
        }
      } else if (bssRow && !cisRow) {
        result.Compliance_Status = 'Not in CIS Scan';
        result.Non_Compliance_Reason = '';
      } else if (!bssRow && cisRow) {
        if (result.CIS_Status === 'Failed') {
          result.Compliance_Status = 'Non-Compliant';
          result.Non_Compliance_Reason = 'Scan Failed (No BSS Policy)';
        } else if (result.CIS_Status === 'Passed') {
          result.Compliance_Status = 'Compliant';
          result.Non_Compliance_Reason = 'No BSS Policy';
        } else {
          // CIS_Status === 'Skipped'
          result.Compliance_Status = 'Not Tested';
          result.Non_Compliance_Reason = 'No BSS Policy';
        }
      } else {
        result.Compliance_Status = '';
        result.Non_Compliance_Reason = '';
      }

      results.push(result);
    });

    return results.sort((a, b) => a.Check_ID.localeCompare(b.Check_ID));
  };

  const generateExcelReport = (data) => {
    const wb = XLSX.utils.book_new();

    // 1) Full Comparison sheet (with trimmed titles)
    const fullData = data.map((row) => ({
      'Check ID': row.Check_ID,
      'BSS ID': row.BSS_ID,
      'In BSS': row.In_BSS,
      'In CIS Scan': row.In_CIS_Scan,
      'BSS Category': row.BSS_Category,
      'BSS Title': row.BSS_Title,
      'CIS Title': row.CIS_Title,
      'Title Mismatch': row.Title_Mismatch,
      'Change Description / Remarks': row.Change_Description_Remarks,
      'Has Remark': row.Has_Remark,
      'CIS Status': row.CIS_Status,
      'CIS Level': row.CIS_Level,
      'CIS Recommended Value': row.CIS_Recommended_Value,
      'Synapxe Value': row.Synapxe_Value,
      'Synapxe Exceptions': row.Synapxe_Exceptions,
      'Failed Instances': row.Failed_Instances,
      'Passed Instances': row.Passed_Instances,
      'Compliance Status': row.Compliance_Status,
      'Non_Compliance_Reason': row.Non_Compliance_Reason,
      'Setting Applicability': row.Setting_Applicability
    }));
    const ws = XLSX.utils.json_to_sheet(fullData);
    XLSX.utils.book_append_sheet(wb, ws, 'Full Comparison');

    // 2) Summary sheet
    const summary = [
      { Metric: 'Total Unique Controls', Count: data.length },
      {
        Metric: 'Controls in BSS Only',
        Count: data.filter((r) => r.In_BSS === 'Yes' && r.In_CIS_Scan === 'No').length
      },
      {
        Metric: 'Controls in CIS Only',
        Count: data.filter((r) => r.In_BSS === 'No' && r.In_CIS_Scan === 'Yes').length
      },
      {
        Metric: 'Controls in Both',
        Count: data.filter((r) => r.In_BSS === 'Yes' && r.In_CIS_Scan === 'Yes').length
      },
      {
        Metric: 'Controls with Remarks',
        Count: data.filter((r) => r.Change_Description_Remarks?.trim()).length
      },
      {
        Metric: 'Controls with Exceptions',
        Count: data.filter((r) => r.Synapxe_Exceptions?.trim()).length
      },
      { Metric: 'Failed Controls', Count: data.filter((r) => r.CIS_Status === 'Failed').length },
      { Metric: 'Passed Controls', Count: data.filter((r) => r.CIS_Status === 'Passed').length },
      { Metric: 'Skipped Controls', Count: data.filter((r) => r.CIS_Status === 'Skipped').length }
    ];
    const summaryWs = XLSX.utils.json_to_sheet(summary);
    XLSX.utils.book_append_sheet(wb, summaryWs, 'Summary');

    // 3) Controls with Remarks
    const remarksData = data.filter((r) => r.Change_Description_Remarks?.trim());
    if (remarksData.length > 0) {
      const remarksWs = XLSX.utils.json_to_sheet(
        remarksData.map((row) => ({
          'Check ID': row.Check_ID,
          'BSS Category': row.BSS_Category,
          'BSS Title': row.BSS_Title,
          'Synapxe Value': row.Synapxe_Value,
          'Change Description / Remarks': row.Change_Description_Remarks,
          'Compliance Status': row.Compliance_Status,
          'Non_Compliance_Reason': row.Non_Compliance_Reason
        }))
      );
      XLSX.utils.book_append_sheet(wb, remarksWs, 'Controls with Remarks');
    }

    // 4) Controls with Exceptions
    const exceptionsData = data.filter((r) => r.Synapxe_Exceptions?.trim());
    if (exceptionsData.length > 0) {
      const exceptionsWs = XLSX.utils.json_to_sheet(
        exceptionsData.map((row) => ({
          'Check ID': row.Check_ID,
          'BSS Category': row.BSS_Category,
          'BSS Title': row.BSS_Title,
          'Synapxe Value': row.Synapxe_Value,
          'Synapxe Exceptions': row.Synapxe_Exceptions,
          'Compliance Status': row.Compliance_Status,
          'Non_Compliance_Reason': row.Non_Compliance_Reason
        }))
      );
      XLSX.utils.book_append_sheet(wb, exceptionsWs, 'Controls with Exceptions');
    }

    // 5) Non-Compliant Details
    const nonCompliantData = data.filter((r) => r.Compliance_Status === 'Non-Compliant');
    if (nonCompliantData.length > 0) {
      const nonCompliantWs = XLSX.utils.json_to_sheet(
        nonCompliantData.map((row) => ({
          'Check ID': row.Check_ID,
          'BSS Category': row.BSS_Category,
          'BSS Title': row.BSS_Title,
          'CIS Title': row.CIS_Title,
          'Title Mismatch': row.Title_Mismatch,
          'Has Remark': row.Has_Remark,
          'CIS Status': row.CIS_Status,
          'Non_Compliance_Reason': row.Non_Compliance_Reason,
          'Change Description / Remarks': row.Change_Description_Remarks
        }))
      );
      XLSX.utils.book_append_sheet(wb, nonCompliantWs, 'Non-Compliant Details');
    }

    // Finally, write the workbook to a file named with today's date
    XLSX.writeFile(wb, `BSS_CIS_Comparison_${new Date().toISOString().slice(0, 10)}.xlsx`);
  };

  const handleCompare = async () => {
    if (!bssFile || !cisFile) {
      alert('Please select both BSS and CIS files');
      return;
    }

    setLoading(true);
    try {
      const bssData = await readBSSFile(bssFile);
      const cisData = await readCISFile(cisFile);
      const comparisonResults = compareData(bssData, cisData);
      setResults(comparisonResults);
    } catch (error) {
      alert(`Error: ${error.message}`);
    } finally {
      setLoading(false);
    }
  };

  const styles = {
    container: {
      maxWidth: '1200px',
      margin: '0 auto',
      padding: '20px',
      fontFamily: 'Arial, sans-serif'
    },
    h1: {
      color: '#333',
      textAlign: 'center'
    },
    fileInputs: {
      background: '#f5f5f5',
      padding: '20px',
      borderRadius: '8px',
      marginBottom: '20px'
    },
    inputGroup: {
      marginBottom: '15px'
    },
    label: {
      display: 'block',
      marginBottom: '5px',
      fontWeight: 'bold'
    },
    fileInput: {
      width: '100%',
      padding: '5px'
    },
    button: {
      background: '#4CAF50',
      color: 'white',
      padding: '10px 20px',
      border: 'none',
      borderRadius: '4px',
      cursor: 'pointer',
      fontSize: '16px'
    },
    buttonDisabled: {
      background: '#cccccc',
      cursor: 'not-allowed'
    },
    results: {
      marginTop: '20px'
    },
    stats: {
      display: 'flex',
      gap: '20px',
      margin: '15px 0',
      fontSize: '14px'
    },
    statBox: {
      background: '#e3f2fd',
      padding: '10px',
      borderRadius: '4px'
    },
    tableContainer: {
      marginTop: '20px',
      overflowX: 'auto'
    },
    table: {
      width: '100%',
      borderCollapse: 'collapse'
    },
    th: {
      border: '1px solid #ddd',
      padding: '8px',
      textAlign: 'left',
      background: '#f2f2f2',
      fontWeight: 'bold'
    },
    td: {
      border: '1px solid #ddd',
      padding: '8px',
      textAlign: 'left'
    },
    trEven: {
      background: '#f9f9f9'
    },
    trFailed: {
      background: '#ffebee'
    }
  };

  // Compute insights once `results` is available
  let compliancePercent = 0;
  let testedControls = [];
  let numPassed = 0,
    numFailed = 0,
    numSkipped = 0;
  let reasonCounts = {};
  let failByCategory = {};
  let totalTested = 0,
    mismatchCount = 0,
    remarkCount = 0;

  if (results) {
    testedControls = results.filter((r) => r.In_CIS_Scan === 'Yes');
    totalTested = testedControls.length;
    numPassed = testedControls.filter((r) => r.Compliance_Status === 'Compliant').length;
    numFailed = testedControls.filter((r) => r.Compliance_Status === 'Non-Compliant').length;
    numSkipped = testedControls.filter((r) => r.Compliance_Status === 'Not Tested').length;
    compliancePercent = totalTested > 0 ? Math.round((numPassed / totalTested) * 100) : 0;

    // Reasons for non-compliance
    reasonCounts = {};
    testedControls.forEach((r) => {
      if (r.Compliance_Status === 'Non-Compliant') {
        const reason = r.Non_Compliance_Reason || 'Other';
        reasonCounts[reason] = (reasonCounts[reason] || 0) + 1;
      }
    });

    // Failures by category (BSS_Category)
    failByCategory = {};
    results.forEach((r) => {
      if (r.Compliance_Status === 'Non-Compliant') {
        const cat = r.BSS_Category || 'Unknown';
        failByCategory[cat] = (failByCategory[cat] || 0) + 1;
      }
    });

    // Title mismatches / remarks
    mismatchCount = results.filter((r) => r.Title_Mismatch === 'Yes').length;
    remarkCount = results.filter((r) => r.Has_Remark === 'Yes').length;
  }

  return (
    <div style={styles.container}>
      <h1 style={styles.h1}>BSS-CIS Comparison Tool</h1>

      <div style={styles.fileInputs}>
        <div style={styles.inputGroup}>
          <label style={styles.label}>BSS Excel File:</label>
          <input
            type="file"
            accept=".xlsx,.xls"
            onChange={(e) => setBssFile(e.target.files[0])}
            style={styles.fileInput}
          />
        </div>

        <div style={styles.inputGroup}>
          <label style={styles.label}>CIS CSV File:</label>
          <input
            type="file"
            accept=".csv"
            onChange={(e) => setCisFile(e.target.files[0])}
            style={styles.fileInput}
          />
        </div>

        <button
          onClick={handleCompare}
          disabled={loading}
          style={loading ? { ...styles.button, ...styles.buttonDisabled } : styles.button}
        >
          {loading ? 'Processing...' : 'Compare Files'}
        </button>
      </div>

      {results && (
        <div style={styles.results}>
          {/* Overall Compliance % */}
          <div style={{ margin: '20px 0', textAlign: 'center' }}>
            <h2>Overall Compliance</h2>
            <div
              style={{
                fontSize: '2rem',
                fontWeight: 'bold',
                color:
                  compliancePercent >= 90
                    ? 'green'
                    : compliancePercent >= 70
                    ? 'orange'
                    : 'red'
              }}
            >
              {compliancePercent}%
            </div>
            <div style={{ fontSize: '0.9rem', color: '#555' }}>
              ({numPassed} passed, {numFailed} failed, {numSkipped} skipped out of{' '}
              {totalTested} tested)
            </div>
          </div>

          {/* Badge for Title Mismatch & Remarks */}
          <div
            style={{
              margin: '20px 0',
              display: 'flex',
              gap: '12px',
              justifyContent: 'center'
            }}
          >
            <div
              style={{
                background: '#eef',
                padding: '6px 10px',
                borderRadius: '4px',
                fontSize: '0.9rem'
              }}
            >
              Title Mismatches: {mismatchCount} (
              {totalTested ? Math.round((mismatchCount / totalTested) * 100) : 0}%)
            </div>
            <div
              style={{
                background: '#efe',
                padding: '6px 10px',
                borderRadius: '4px',
                fontSize: '0.9rem'
              }}
            >
              Controls with Remarks: {remarkCount} (
              {totalTested ? Math.round((remarkCount / totalTested) * 100) : 0}%)
            </div>
          </div>

          {/* Reasons for Non-Compliance */}
          <div style={{ margin: '20px auto', maxWidth: '400px' }}>
            <h3 style={{ fontSize: '1.1rem', marginBottom: '8px', textAlign: 'center' }}>
              Reasons for Non-Compliance
            </h3>
            {Object.entries(reasonCounts).map(([reason, count]) => {
              const widthPercent = totalTested > 0 ? Math.round((count / totalTested) * 100) : 0;
              return (
                <div key={reason} style={{ marginBottom: '8px' }}>
                  <div style={{ fontSize: '0.85rem', marginBottom: '2px' }}>
                    {reason} ({count})
                  </div>
                  <div
                    style={{
                      background: '#fcb',
                      width: `${widthPercent}%`,
                      height: '8px',
                      border: '1px solid #faa',
                      borderRadius: '4px'
                    }}
                  />
                </div>
              );
            })}
          </div>

          {/* Failures by BSS Category */}
          <div style={{ margin: '20px 0' }}>
            <h3 style={{ fontSize: '1.1rem' }}>Non-Compliant by BSS Category</h3>
            <ul style={{ fontSize: '0.9rem', paddingLeft: '20px' }}>
              {Object.entries(failByCategory).map(([cat, cnt]) => (
                <li key={cat}>
                  {cat}: {cnt}
                </li>
              ))}
            </ul>
          </div>

          {/* Download Button */}
          <button onClick={() => generateExcelReport(results)} style={styles.button}>
            Download Excel Report
          </button>

          {/* Table of Results (first 50 rows) */}
          <div style={styles.tableContainer}>
            <table style={styles.table}>
              <thead>
                <tr>
                  <th style={styles.th}>Check ID</th>
                  <th style={styles.th}>Category</th>
                  <th style={styles.th}>Compliance Status</th>
                  <th style={styles.th}>CIS Rec Value</th>
                  <th style={styles.th}>Synapxe Value</th>
                  <th style={styles.th}>Exceptions</th>
                  <th style={styles.th}>Remarks</th>
                </tr>
              </thead>
              <tbody>
                {results.slice(0, 50).map((row, idx) => (
                  <tr
                    key={idx}
                    style={
                      row.CIS_Status === 'Failed'
                        ? styles.trFailed
                        : idx % 2 === 0
                        ? styles.trEven
                        : {}
                    }
                  >
                    <td style={styles.td}>{row.Check_ID}</td>
                    <td style={styles.td}>{row.BSS_Category || '-'}</td>
                    <td style={styles.td}>{row.Compliance_Status}</td>
                    <td style={styles.td}>{row.CIS_Recommended_Value || '-'}</td>
                    <td style={styles.td}>{row.Synapxe_Value || '-'}</td>
                    <td style={styles.td}>{row.Synapxe_Exceptions || '-'}</td>
                    <td style={styles.td}>{row.Change_Description_Remarks || '-'}</td>
                  </tr>
                ))}
              </tbody>
            </table>
            {results.length > 50 && <p>Showing first 50 of {results.length} results</p>}
          </div>
        </div>
      )}
    </div>
  );
}
