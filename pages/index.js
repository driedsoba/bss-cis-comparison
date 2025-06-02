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
        const appCol = Object.keys(bssRow).find(key => key.includes('Setting Applicability'));
        const cisRecCol = Object.keys(bssRow).find(key => key.includes('CIS Recommended Value'));
        
        result.BSS_Title = bssRow[titleCol] || '';
        result.BSS_Category = bssRow.Category || '';
        result.Setting_Applicability = bssRow[appCol] || '';
        result.CIS_Recommended_Value = bssRow[cisRecCol] || '';
        result.Synapxe_Value = bssRow['Synapxe Value'] || '';
        result.Synapxe_Exceptions = bssRow['Synapxe Exceptions'] || '';
        result.Change_Description_Remarks = bssRow['Change Description / Remarks'] || '';
        
        // Check if BSS has its own ID column
        result.BSS_ID = bssRow['BSS ID'] || bssRow['BSS #'] || checkId;
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
    
    // Full comparison sheet with all columns
    const fullData = data.map(row => ({
      'Check ID': row.Check_ID,
      'BSS ID': row.BSS_ID || row.Check_ID,
      'In BSS': row.In_BSS,
      'In CIS Scan': row.In_CIS_Scan,
      'Compliance Status': row.Compliance_Status,
      'CIS Status': row.CIS_Status || '',
      'BSS Category': row.BSS_Category || '',
      'Setting Applicability': row.Setting_Applicability || '',
      'BSS Title': row.BSS_Title || '',
      'CIS Title': row.CIS_Title || '',
      'CIS Level': row.CIS_Level || '',
      'CIS Recommended Value': row.CIS_Recommended_Value || '',
      'Synapxe Value': row.Synapxe_Value || '',
      'Synapxe Exceptions': row.Synapxe_Exceptions || '',
      'Change Description / Remarks': row.Change_Description_Remarks || '',
      'Failed Instances': row.Failed_Instances || '',
      'Passed Instances': row.Passed_Instances || ''
    }));
    
    const ws = XLSX.utils.json_to_sheet(fullData);
    XLSX.utils.book_append_sheet(wb, ws, "Full Comparison");
    
    // Summary sheet
    const summary = [
      { Metric: 'Total Unique Controls', Count: data.length },
      { Metric: 'Controls in BSS Only', Count: data.filter(r => r.In_BSS === 'Yes' && r.In_CIS_Scan === 'No').length },
      { Metric: 'Controls in CIS Only', Count: data.filter(r => r.In_BSS === 'No' && r.In_CIS_Scan === 'Yes').length },
      { Metric: 'Controls in Both', Count: data.filter(r => r.In_BSS === 'Yes' && r.In_CIS_Scan === 'Yes').length },
      { Metric: 'Controls with Remarks', Count: data.filter(r => r.Change_Description_Remarks && r.Change_Description_Remarks.trim()).length },
      { Metric: 'Controls with Exceptions', Count: data.filter(r => r.Synapxe_Exceptions && r.Synapxe_Exceptions.trim()).length },
      { Metric: 'Failed Controls', Count: data.filter(r => r.CIS_Status === 'Failed').length },
      { Metric: 'Passed Controls', Count: data.filter(r => r.CIS_Status === 'Passed').length },
      { Metric: 'Skipped Controls', Count: data.filter(r => r.CIS_Status === 'Skipped').length }
    ];
    const summaryWs = XLSX.utils.json_to_sheet(summary);
    XLSX.utils.book_append_sheet(wb, summaryWs, "Summary");
    
    // Controls with remarks
    const remarksData = data.filter(r => r.Change_Description_Remarks && r.Change_Description_Remarks.trim());
    if (remarksData.length > 0) {
      const remarksWs = XLSX.utils.json_to_sheet(remarksData.map(row => ({
        'Check ID': row.Check_ID,
        'BSS Category': row.BSS_Category,
        'BSS Title': row.BSS_Title,
        'Synapxe Value': row.Synapxe_Value,
        'Change Description / Remarks': row.Change_Description_Remarks,
        'Compliance Status': row.Compliance_Status
      })));
      XLSX.utils.book_append_sheet(wb, remarksWs, "Controls with Remarks");
    }
    
    // Controls with exceptions
    const exceptionsData = data.filter(r => r.Synapxe_Exceptions && r.Synapxe_Exceptions.trim());
    if (exceptionsData.length > 0) {
      const exceptionsWs = XLSX.utils.json_to_sheet(exceptionsData.map(row => ({
        'Check ID': row.Check_ID,
        'BSS Category': row.BSS_Category,
        'BSS Title': row.BSS_Title,
        'Synapxe Value': row.Synapxe_Value,
        'Synapxe Exceptions': row.Synapxe_Exceptions,
        'Compliance Status': row.Compliance_Status
      })));
      XLSX.utils.book_append_sheet(wb, exceptionsWs, "Controls with Exceptions");
    }
    
    // Non-compliant controls
    const failedData = data.filter(r => r.CIS_Status === 'Failed');
    if (failedData.length > 0) {
      const failedWs = XLSX.utils.json_to_sheet(failedData.map(row => ({
        'Check ID': row.Check_ID,
        'BSS Category': row.BSS_Category,
        'BSS Title': row.BSS_Title,
        'CIS Title': row.CIS_Title,
        'CIS Recommended Value': row.CIS_Recommended_Value,
        'Synapxe Value': row.Synapxe_Value,
        'Failed Instances': row.Failed_Instances,
        'Change Description / Remarks': row.Change_Description_Remarks
      })));
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
          style={loading ? {...styles.button, ...styles.buttonDisabled} : styles.button}
        >
          {loading ? 'Processing...' : 'Compare Files'}
        </button>
      </div>
      
      {results && (
        <div style={styles.results}>
          <h2>Results Summary</h2>
          <div style={styles.stats}>
            <div style={styles.statBox}>Total Controls: {results.length}</div>
            <div style={styles.statBox}>Failed: {results.filter(r => r.CIS_Status === 'Failed').length}</div>
            <div style={styles.statBox}>Passed: {results.filter(r => r.CIS_Status === 'Passed').length}</div>
            <div style={styles.statBox}>With Remarks: {results.filter(r => r.Change_Description_Remarks && r.Change_Description_Remarks.trim()).length}</div>
            <div style={styles.statBox}>With Exceptions: {results.filter(r => r.Synapxe_Exceptions && r.Synapxe_Exceptions.trim()).length}</div>
          </div>
          
          <button onClick={() => generateExcelReport(results)} style={styles.button}>
            Download Excel Report
          </button>
          
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
                      row.CIS_Status === 'Failed' ? styles.trFailed : 
                      idx % 2 === 0 ? styles.trEven : {}
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
