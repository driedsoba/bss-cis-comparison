// pages/index.js
import { useState } from 'react';
import * as XLSX from 'xlsx';
import Papa from 'papaparse';

export default function Home() {
  const [bssFile, setBssFile] = useState(null);
  const [cisFile, setCisFile] = useState(null);
  const [results, setResults] = useState(null);
  const [loading, setLoading] = useState(false);

  const readBSSFile = async (file) => {
    const data = await file.arrayBuffer();
    const workbook = XLSX.read(data);
    
    // Find settings sheet
    const sheetName = Object.keys(workbook.Sheets).find(name => 
      name.toLowerCase().includes('settings') || name.toLowerCase().includes('server')
    );
    
    if (!sheetName) throw new Error("Settings sheet not found");
    
    const worksheet = workbook.Sheets[sheetName];
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
    
    // Find header row
    let headerRowIdx = jsonData.findIndex(row => 
      row.some(cell => cell && cell.toString().includes('CIS #'))
    );
    
    if (headerRowIdx === -1) throw new Error("Header row not found");
    
    const headers = jsonData[headerRowIdx];
    const data_rows = jsonData.slice(headerRowIdx + 1)
      .filter(row => row[headers.indexOf('CIS #')])
      .map(row => {
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
          // Skip metadata rows if present
          let data = result.data;
          let headerIdx = data.findIndex(row => 
            row[0] === 'check_id' || (row.length > 1 && row[0].includes('check_id'))
          );
          
          if (headerIdx > 0) {
            data = data.slice(headerIdx);
          }
          
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
    const bssIds = new Set(bssData.map(row => row['CIS #']?.toString()));
    const cisIds = new Set(cisData.map(row => row.check_id?.toString()));
    const allIds = new Set([...bssIds, ...cisIds]);
    
    allIds.forEach(checkId => {
      const bssRow = bssData.find(row => row['CIS #']?.toString() === checkId);
      const cisRow = cisData.find(row => row.check_id?.toString() === checkId);
      
      const result = {
        Check_ID: checkId,
        In_BSS: bssRow ? 'Yes' : 'No',
        In_CIS_Scan: cisRow ? 'Yes' : 'No'
      };
      
      if (bssRow) {
        const titleCol = Object.keys(bssRow).find(key => key.includes('Setting Title'));
        const remarksCol = Object.keys(bssRow).find(key => key.toLowerCase().includes('remark'));
        
        result.BSS_Title = bssRow[titleCol] || '';
        result.BSS_Category = bssRow.Category || '';
        result.BSS_Remarks = bssRow[remarksCol] || '';
        result.Has_Remarks = result.BSS_Remarks ? 'Yes' : 'No';
      }
      
      if (cisRow) {
        result.CIS_Title = cisRow.title || '';
        result.CIS_Level = cisRow.level || '';
        
        if (cisRow.failed_instances && cisRow.failed_instances !== 'None') {
          result.CIS_Status = 'Failed';
          result.Compliance_Status = 'Non-Compliant';
        } else if (cisRow.passed_instances && cisRow.passed_instances !== 'None') {
          result.CIS_Status = 'Passed';
          result.Compliance_Status = 'Compliant';
        } else {
          result.CIS_Status = 'Skipped';
          result.Compliance_Status = 'Not Tested';
        }
        
        result.Failed_Instances = cisRow.failed_instances || '';
        result.Passed_Instances = cisRow.passed_instances || '';
      }
      
      if (result.In_BSS === 'Yes' && result.In_CIS_Scan === 'No') {
        result.Compliance_Status = 'Not in CIS Scan';
      } else if (result.In_BSS === 'No' && result.In_CIS_Scan === 'Yes') {
        result.Compliance_Status = 'Not in BSS Policy';
      }
      
      results.push(result);
    });
    
    return results.sort((a, b) => a.Check_ID.localeCompare(b.Check_ID));
  };

  const generateExcelReport = (data) => {
    const wb = XLSX.utils.book_new();
    
    // Full comparison sheet
    const ws = XLSX.utils.json_to_sheet(data);
    XLSX.utils.book_append_sheet(wb, ws, "Full Comparison");
    
    // Summary sheet
    const summary = [
      { Metric: 'Total Unique Controls', Count: data.length },
      { Metric: 'Controls in BSS Only', Count: data.filter(r => r.In_BSS === 'Yes' && r.In_CIS_Scan === 'No').length },
      { Metric: 'Controls in CIS Only', Count: data.filter(r => r.In_BSS === 'No' && r.In_CIS_Scan === 'Yes').length },
      { Metric: 'Controls in Both', Count: data.filter(r => r.In_BSS === 'Yes' && r.In_CIS_Scan === 'Yes').length },
      { Metric: 'Controls with Remarks', Count: data.filter(r => r.Has_Remarks === 'Yes').length },
      { Metric: 'Failed Controls', Count: data.filter(r => r.CIS_Status === 'Failed').length },
      { Metric: 'Passed Controls', Count: data.filter(r => r.CIS_Status === 'Passed').length },
      { Metric: 'Skipped Controls', Count: data.filter(r => r.CIS_Status === 'Skipped').length }
    ];
    const summaryWs = XLSX.utils.json_to_sheet(summary);
    XLSX.utils.book_append_sheet(wb, summaryWs, "Summary");
    
    // Controls with remarks
    const remarksData = data.filter(r => r.Has_Remarks === 'Yes');
    if (remarksData.length > 0) {
      const remarksWs = XLSX.utils.json_to_sheet(remarksData);
      XLSX.utils.book_append_sheet(wb, remarksWs, "Controls with Remarks");
    }
    
    // Non-compliant controls
    const failedData = data.filter(r => r.CIS_Status === 'Failed');
    if (failedData.length > 0) {
      const failedWs = XLSX.utils.json_to_sheet(failedData);
      XLSX.utils.book_append_sheet(wb, failedWs, "Non-Compliant Controls");
    }
    
    // Generate file
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

  return (
    <div className="container">
      <h1>BSS-CIS Comparison Tool</h1>
      
      <div className="file-inputs">
        <div className="input-group">
          <label>BSS Excel File:</label>
          <input 
            type="file" 
            accept=".xlsx,.xls" 
            onChange={(e) => setBssFile(e.target.files[0])}
          />
        </div>
        
        <div className="input-group">
          <label>CIS CSV File:</label>
          <input 
            type="file" 
            accept=".csv" 
            onChange={(e) => setCisFile(e.target.files[0])}
          />
        </div>
        
        <button onClick={handleCompare} disabled={loading}>
          {loading ? 'Processing...' : 'Compare Files'}
        </button>
      </div>
      
      {results && (
        <div className="results">
          <h2>Results Summary</h2>
          <div className="stats">
            <div>Total Controls: {results.length}</div>
            <div>Failed: {results.filter(r => r.CIS_Status === 'Failed').length}</div>
            <div>Passed: {results.filter(r => r.CIS_Status === 'Passed').length}</div>
            <div>With Remarks: {results.filter(r => r.Has_Remarks === 'Yes').length}</div>
          </div>
          
          <button onClick={() => generateExcelReport(results)}>
            Download Excel Report
          </button>
          
          <div className="table-container">
            <table>
              <thead>
                <tr>
                  <th>Check ID</th>
                  <th>In BSS</th>
                  <th>In CIS</th>
                  <th>Compliance Status</th>
                  <th>Has Remarks</th>
                </tr>
              </thead>
              <tbody>
                {results.slice(0, 50).map((row, idx) => (
                  <tr key={idx} className={row.CIS_Status === 'Failed' ? 'failed' : ''}>
                    <td>{row.Check_ID}</td>
                    <td>{row.In_BSS}</td>
                    <td>{row.In_CIS_Scan}</td>
                    <td>{row.Compliance_Status}</td>
                    <td>{row.Has_Remarks}</td>
                  </tr>
                ))}
              </tbody>
            </table>
            {results.length > 50 && <p>Showing first 50 of {results.length} results</p>}
          </div>
        </div>
      )}
      
      <style jsx>{`
        .container {
          max-width: 1200px;
          margin: 0 auto;
          padding: 20px;
          font-family: Arial, sans-serif;
        }
        
        h1 {
          color: #333;
          text-align: center;
        }
        
        .file-inputs {
          background: #f5f5f5;
          padding: 20px;
          border-radius: 8px;
          margin-bottom: 20px;
        }
        
        .input-group {
          margin-bottom: 15px;
        }
        
        label {
          display: block;
          margin-bottom: 5px;
          font-weight: bold;
        }
        
        input[type="file"] {
          width: 100%;
          padding: 5px;
        }
        
        button {
          background: #4CAF50;
          color: white;
          padding: 10px 20px;
          border: none;
          border-radius: 4px;
          cursor: pointer;
          font-size: 16px;
        }
        
        button:hover {
          background: #45a049;
        }
        
        button:disabled {
          background: #cccccc;
          cursor: not-allowed;
        }
        
        .results {
          margin-top: 20px;
        }
        
        .stats {
          display: flex;
          gap: 20px;
          margin: 15px 0;
          font-size: 14px;
        }
        
        .stats div {
          background: #e3f2fd;
          padding: 10px;
          border-radius: 4px;
        }
        
        .table-container {
          margin-top: 20px;
          overflow-x: auto;
        }
        
        table {
          width: 100%;
          border-collapse: collapse;
        }
        
        th, td {
          border: 1px solid #ddd;
          padding: 8px;
          text-align: left;
        }
        
        th {
          background: #f2f2f2;
          font-weight: bold;
        }
        
        tr:nth-child(even) {
          background: #f9f9f9;
        }
        
        tr.failed {
          background: #ffebee;
        }
        
        tr:hover {
          background: #e3f2fd;
        }
      `}</style>
    </div>
  );
}

// package.json
{
  "name": "bss-cis-comparison",
  "version": "1.0.0",
  "private": true,
  "scripts": {
    "dev": "next dev",
    "build": "next build",
    "start": "next start"
  },
  "dependencies": {
    "next": "latest",
    "react": "latest",
    "react-dom": "latest",
    "xlsx": "^0.18.5",
    "papaparse": "^5.4.1"
  }
}
