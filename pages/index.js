import React, { useState, useCallback } from 'react';
import * as XLSX from 'xlsx';
import _ from 'lodash';
import Papa from 'papaparse';

const BSSCISAnalyzer = () => {
  const [analysis, setAnalysis] = useState(null);
  const [loading, setLoading] = useState(false);

  // Enhanced similarity algorithm with multiple methods
  const calculateSimilarity = (str1, str2) => {
    if (!str1 || !str2) return 0;
    
    const s1 = str1.toString().toLowerCase().trim();
    const s2 = str2.toString().toLowerCase().trim();
    
    // Exact match
    if (s1 === s2) return 1.0;
    
    // Normalize quotes and special characters
    const normalize = (s) => s
      .replace(/['']/g, "'")
      .replace(/[""]/g, '"')
      .replace(/[–—]/g, '-')
      .replace(/\s+/g, ' ')
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '');
    
    const n1 = normalize(s1);
    const n2 = normalize(s2);
    
    if (n1 === n2) return 0.98;
    
    // Multiple similarity algorithms
    const levenshtein = getLevenshteinSimilarity(n1, n2);
    const jaccard = getJaccardSimilarity(n1, n2);
    const cosine = getCosineSimilarity(n1, n2);
    
    // Weighted average
    return (levenshtein * 0.4 + jaccard * 0.3 + cosine * 0.3);
  };

  // Levenshtein distance similarity
  const getLevenshteinSimilarity = (s1, s2) => {
    const matrix = [];
    for (let i = 0; i <= s2.length; i++) {
      matrix[i] = [i];
    }
    for (let j = 0; j <= s1.length; j++) {
      matrix[0][j] = j;
    }
    for (let i = 1; i <= s2.length; i++) {
      for (let j = 1; j <= s1.length; j++) {
        if (s2.charAt(i - 1) === s1.charAt(j - 1)) {
          matrix[i][j] = matrix[i - 1][j - 1];
        } else {
          matrix[i][j] = Math.min(
            matrix[i - 1][j - 1] + 1,
            matrix[i][j - 1] + 1,
            matrix[i - 1][j] + 1
          );
        }
      }
    }
    const distance = matrix[s2.length][s1.length];
    return 1 - distance / Math.max(s1.length, s2.length);
  };

  // Jaccard similarity
  const getJaccardSimilarity = (s1, s2) => {
    const set1 = new Set(s1.split(/\s+/));
    const set2 = new Set(s2.split(/\s+/));
    const intersection = new Set([...set1].filter(x => set2.has(x)));
    const union = new Set([...set1, ...set2]);
    return union.size === 0 ? 0 : intersection.size / union.size;
  };

  // Cosine similarity
  const getCosineSimilarity = (s1, s2) => {
    const words1 = s1.split(/\s+/);
    const words2 = s2.split(/\s+/);
    const allWords = [...new Set([...words1, ...words2])];
    
    const vector1 = allWords.map(word => words1.filter(w => w === word).length);
    const vector2 = allWords.map(word => words2.filter(w => w === word).length);
    
    const dotProduct = vector1.reduce((sum, val, i) => sum + val * vector2[i], 0);
    const magnitude1 = Math.sqrt(vector1.reduce((sum, val) => sum + val * val, 0));
    const magnitude2 = Math.sqrt(vector2.reduce((sum, val) => sum + val * val, 0));
    
    return magnitude1 === 0 || magnitude2 === 0 ? 0 : dotProduct / (magnitude1 * magnitude2);
  };

  // Enhanced pattern detection
  const detectPatterns = (data) => {
    const patterns = {
      idFormats: new Map(),
      valueMappings: new Map(),
      commonPrefixes: new Map(),
      implementationGaps: []
    };

    data.forEach(row => {
      // Detect ID format patterns
      if (row.BSS_ID) {
        const idStr = String(row.BSS_ID);
        const format = idStr.replace(/\d+/g, 'N');
        patterns.idFormats.set(format, (patterns.idFormats.get(format) || 0) + 1);
      }

      // Detect value mapping patterns
      if (row.BSS_Expected_Value && row.CIS_Value) {
        const mapping = `${row.BSS_Expected_Value} => ${row.CIS_Value}`;
        patterns.valueMappings.set(mapping, (patterns.valueMappings.get(mapping) || 0) + 1);
      }

      // Detect implementation gaps
      if (row.BSS_ID && !row.CIS_ID && row.Compliance === 'Fail') {
        patterns.implementationGaps.push({
          id: row.BSS_ID,
          title: row.BSS_Title,
          category: row.BSS_Category
        });
      }
    });

    return patterns;
  };

  // Advanced compliance analysis
  const analyzeCompliance = (data) => {
    const compliance = {
      byCategory: {},
      bySeverity: {},
      trends: [],
      riskScore: 0
    };

    // Group by category
    const byCategory = _.groupBy(data, 'BSS_Category');
    Object.entries(byCategory).forEach(([category, items]) => {
      const passed = items.filter(i => i.Compliance === 'Pass').length;
      const failed = items.filter(i => i.Compliance === 'Fail').length;
      const skipped = items.filter(i => i.Compliance === 'Skipped').length;
      
      compliance.byCategory[category] = {
        total: items.length,
        passed,
        failed,
        skipped,
        passRate: items.length > 0 ? (passed / items.length * 100).toFixed(1) : 0
      };
    });

    // Calculate risk score
    const weights = { critical: 10, high: 5, medium: 2, low: 1 };
    let totalWeight = 0;
    let failedWeight = 0;

    data.forEach(item => {
      const severity = item.BSS_Severity || 'medium';
      const weight = weights[severity.toLowerCase()] || 2;
      totalWeight += weight;
      if (item.Compliance === 'Fail') {
        failedWeight += weight;
      }
    });

    compliance.riskScore = totalWeight > 0 ? 
      Math.round((1 - failedWeight / totalWeight) * 100) : 100;

    return compliance;
  };

  // Generate remediation plan
  const generateRemediationPlan = (data) => {
    const failed = data.filter(d => d.Compliance === 'Fail');
    
    // Group by priority
    const byPriority = {
      critical: [],
      high: [],
      medium: [],
      low: []
    };

    failed.forEach(item => {
      const priority = determinePriority(item);
      const plan = {
        id: item.CIS_ID || item.BSS_ID,
        title: item.CIS_Title || item.BSS_Title,
        category: item.BSS_Category,
        currentValue: item.CIS_Value || 'Not Configured',
        expectedValue: item.BSS_Expected_Value,
        remediation: item.CIS_Remediation || generateDefaultRemediation(item),
        estimatedEffort: estimateEffort(item),
        dependencies: findDependencies(item, data)
      };
      
      byPriority[priority].push(plan);
    });

    return byPriority;
  };

  const determinePriority = (item) => {
    if (item.BSS_Category?.includes('Authentication') || 
        item.BSS_Category?.includes('Account')) return 'critical';
    if (item.BSS_Category?.includes('Audit') || 
        item.BSS_Category?.includes('Security')) return 'high';
    if (item.BSS_Category?.includes('Network')) return 'medium';
    return 'low';
  };

  const generateDefaultRemediation = (item) => {
    return `Configure ${item.BSS_Title || item.CIS_Title} to meet the expected value: ${item.BSS_Expected_Value}`;
  };

  const estimateEffort = (item) => {
    if (item.BSS_Expected_Value?.toLowerCase().includes('enabled')) return '5 minutes';
    if (item.BSS_Expected_Value?.toLowerCase().includes('configured')) return '15 minutes';
    return '30 minutes';
  };

  const findDependencies = (item, allData) => {
    const deps = [];
    const itemCategory = item.BSS_Category;
    const itemId = item.BSS_ID?.split('.').slice(0, 2).join('.');
    
    allData.forEach(other => {
      if (other.BSS_ID !== item.BSS_ID && 
          other.BSS_Category === itemCategory &&
          other.BSS_ID?.startsWith(itemId) &&
          other.Compliance === 'Fail') {
        deps.push(other.BSS_ID);
      }
    });
    
    return deps;
  };

  // Process files
  const processFiles = async (bssFile, cisFile) => {
    setLoading(true);
    try {
      // Read BSS Excel
      const bssBuffer = await readFileAsArrayBuffer(bssFile);
      const bssWorkbook = XLSX.read(bssBuffer);
      
      // Get the correct sheet - "Windows Servers 2019 settings"
      const sheetName = bssWorkbook.SheetNames.find(name => 
        name.includes('settings') || name.includes('Windows')
      ) || bssWorkbook.SheetNames[1];
      
      const bssSheet = bssWorkbook.Sheets[sheetName];
      
      // Read the sheet with header row at position 5 (0-indexed = 4)
      const rawData = XLSX.utils.sheet_to_json(bssSheet, {header: 1, defval: ''});
      
      // Find the header row
      const headerRowIndex = rawData.findIndex(row => 
        row.some(cell => cell && (
          cell.toString().includes('Category') || 
          cell.toString().includes('CIS #')
        ))
      );
      
      if (headerRowIndex === -1) {
        throw new Error('Could not find header row in BSS file');
      }
      
      // Extract headers and data
      const headers = rawData[headerRowIndex];
      const dataRows = rawData.slice(headerRowIndex + 1);
      
      // Convert to objects with proper field mapping
      const bssData = dataRows
        .filter(row => row.some(cell => cell)) // Skip empty rows
        .map(row => {
          const obj = {};
          headers.forEach((header, index) => {
            if (header && row[index] !== undefined) {
              // Clean header names
              const cleanHeader = header.toString().replace(/\r?\n/g, ' ').trim();
              obj[cleanHeader] = row[index];
            }
          });
          return obj;
        });

      // Read CIS CSV
      const cisText = await readFileAsText(cisFile);
      const cisLines = cisText.split('\n');
      const dataStartIndex = cisLines.findIndex(line => line.includes('check_id'));
      const cisDataOnly = cisLines.slice(dataStartIndex).join('\n');
      
      const cisParsed = Papa.parse(cisDataOnly, {
        header: true,
        dynamicTyping: true,
        skipEmptyLines: true
      });

      // Map and merge data
      const mergedData = mergeBSSAndCIS(bssData, cisParsed.data);
      
      // Perform analyses
      const patterns = detectPatterns(mergedData);
      const compliance = analyzeCompliance(mergedData);
      const remediation = generateRemediationPlan(mergedData);
      
      setAnalysis({
        data: mergedData,
        patterns,
        compliance,
        remediation,
        summary: generateSummary(mergedData)
      });

    } catch (error) {
      console.error('Error processing files:', error);
      alert('Error processing files: ' + error.message);
    }
    setLoading(false);
  };

  const mergeBSSAndCIS = (bssData, cisData) => {
    // Create lookup maps
    const cisMap = new Map();
    cisData.forEach(cis => {
      if (cis.check_id) {
        cisMap.set(cis.check_id, cis);
      }
    });

    // Merge data with enhanced matching
    const merged = [];
    
    bssData.forEach(bss => {
      // Map the BSS fields correctly based on actual column names
      const bssId = bss['CIS #'];
      const cisRecord = cisMap.get(bssId);
      
      const record = {
        BSS_ID: bssId,
        CIS_ID: cisRecord?.check_id || bssId,
        BSS_Title: bss['Synapxe Setting Title'] || bss['CIS Setting Title (for reference only)'],
        CIS_Title: cisRecord?.title,
        BSS_Category: bss['Category'] || bss['CIS Section Header'],
        BSS_Expected_Value: bss['Synapxe Value'] || bss['CIS Recommended Value (for reference only)'],
        CIS_Value: extractCISValue(cisRecord),
        CIS_Status: determineStatus(cisRecord),
        Compliance: determineCompliance(bss, cisRecord),
        Title_Similarity: calculateSimilarity(
          bss['Synapxe Setting Title'] || bss['CIS Setting Title (for reference only)'], 
          cisRecord?.title
        ),
        BSS_Exceptions: bss['Synapxe Exceptions (Default exceptions to refer to CIS Benchmark)'],
        CIS_Remediation: cisRecord?.remediation,
        Failed_Instances: cisRecord?.failed_instances,
        Passed_Instances: cisRecord?.passed_instances,
        Remarks: bss['Change Description / Remarks'],
        Applicability: bss['Setting Applicability (refer to compliance section under Cover tab)']
      };
      
      merged.push(record);
    });

    // Add CIS-only records
    cisData.forEach(cis => {
      if (!merged.find(m => m.CIS_ID === cis.check_id)) {
        merged.push({
          BSS_ID: null,
          CIS_ID: cis.check_id,
          BSS_Title: null,
          CIS_Title: cis.title,
          BSS_Category: 'Uncategorized',
          BSS_Expected_Value: null,
          CIS_Value: extractCISValue(cis),
          CIS_Status: determineStatus(cis),
          Compliance: 'No BSS Mapping',
          Title_Similarity: 0,
          CIS_Remediation: cis.remediation
        });
      }
    });

    return merged;
  };

  const extractCISValue = (cisRecord) => {
    if (!cisRecord) return null;
    // Extract actual value from description or use status
    if (cisRecord.failed_instances > 0) return 'Non-Compliant';
    if (cisRecord.passed_instances > 0) return 'Compliant';
    if (cisRecord.skipped_instances > 0) return 'Skipped';
    return 'Not Available';
  };

  const determineStatus = (cisRecord) => {
    if (!cisRecord) return 'Not Scanned';
    if (cisRecord.failed_instances > 0) return 'FAILED';
    if (cisRecord.passed_instances > 0) return 'PASSED';
    if (cisRecord.skipped_instances > 0) return 'SKIPPED';
    return 'Unknown';
  };

  const determineCompliance = (bss, cis) => {
    if (!cis) return 'Not Scanned';
    if (cis.skipped_instances > 0) return 'Skipped';
    if (cis.failed_instances > 0) return 'Fail';
    if (cis.passed_instances > 0) return 'Pass';
    return 'Unknown';
  };

  const generateSummary = (data) => {
    return {
      total: data.length,
      passed: data.filter(d => d.Compliance === 'Pass').length,
      failed: data.filter(d => d.Compliance === 'Fail').length,
      skipped: data.filter(d => d.Compliance === 'Skipped').length,
      notScanned: data.filter(d => d.Compliance === 'Not Scanned').length,
      titleMismatches: data.filter(d => d.Title_Similarity < 0.8 && d.Title_Similarity > 0).length,
      unmappedBSS: data.filter(d => d.BSS_ID && !d.CIS_ID).length,
      unmappedCIS: data.filter(d => !d.BSS_ID && d.CIS_ID).length
    };
  };

  // Export to Excel with enhanced formatting
  const exportToExcel = () => {
    if (!analysis) return;

    const wb = XLSX.utils.book_new();

    // Main comparison sheet
    const mainWs = XLSX.utils.json_to_sheet(analysis.data);
    XLSX.utils.book_append_sheet(wb, mainWs, 'Full Comparison');

    // Summary dashboard
    const summaryData = [
      { Metric: 'Total Controls', Value: analysis.summary.total },
      { Metric: 'Passed', Value: analysis.summary.passed },
      { Metric: 'Failed', Value: analysis.summary.failed },
      { Metric: 'Skipped', Value: analysis.summary.skipped },
      { Metric: 'Not Scanned', Value: analysis.summary.notScanned },
      { Metric: 'Title Mismatches', Value: analysis.summary.titleMismatches },
      { Metric: 'Risk Score', Value: analysis.compliance.riskScore + '%' }
    ];
    const summaryWs = XLSX.utils.json_to_sheet(summaryData);
    XLSX.utils.book_append_sheet(wb, summaryWs, 'Summary');

    // Compliance by category
    const categoryData = Object.entries(analysis.compliance.byCategory).map(([cat, stats]) => ({
      Category: cat,
      Total: stats.total,
      Passed: stats.passed,
      Failed: stats.failed,
      'Pass Rate': stats.passRate + '%'
    }));
    const categoryWs = XLSX.utils.json_to_sheet(categoryData);
    XLSX.utils.book_append_sheet(wb, categoryWs, 'By Category');

    // Remediation plan
    const remediationData = [];
    Object.entries(analysis.remediation).forEach(([priority, items]) => {
      items.forEach(item => {
        remediationData.push({
          Priority: priority.toUpperCase(),
          ID: item.id,
          Title: item.title,
          Category: item.category,
          'Current Value': item.currentValue,
          'Expected Value': item.expectedValue,
          Remediation: item.remediation,
          'Estimated Effort': item.estimatedEffort,
          Dependencies: item.dependencies.join(', ')
        });
      });
    });
    const remediationWs = XLSX.utils.json_to_sheet(remediationData);
    XLSX.utils.book_append_sheet(wb, remediationWs, 'Remediation Plan');

    // Pattern analysis
    const patternData = [
      { Pattern: 'Most Common ID Format', Value: [...analysis.patterns.idFormats.entries()].sort((a, b) => b[1] - a[1])[0]?.[0] || 'N/A' },
      { Pattern: 'Implementation Gaps', Value: analysis.patterns.implementationGaps.length },
      { Pattern: 'Unique Value Mappings', Value: analysis.patterns.valueMappings.size }
    ];
    const patternWs = XLSX.utils.json_to_sheet(patternData);
    XLSX.utils.book_append_sheet(wb, patternWs, 'Patterns');

    // Save file
    XLSX.writeFile(wb, `BSS_CIS_Advanced_Analysis_${new Date().toISOString().slice(0, 10)}.xlsx`);
  };

  // Helper functions
  const readFileAsArrayBuffer = (file) => {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e) => resolve(e.target.result);
      reader.onerror = reject;
      reader.readAsArrayBuffer(file);
    });
  };

  const readFileAsText = (file) => {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e) => resolve(e.target.result);
      reader.onerror = reject;
      reader.readAsText(file);
    });
  };

  // UI Components
  return (
    <div style={{ padding: '20px', fontFamily: 'Arial, sans-serif' }}>
      <h1>Advanced BSS-CIS Compliance Analyzer</h1>
      
      <div style={{ marginBottom: '20px' }}>
        <div style={{ marginBottom: '10px' }}>
          <label>BSS Excel File: </label>
          <input type="file" id="bssFile" accept=".xlsx,.xls" />
        </div>
        <div style={{ marginBottom: '10px' }}>
          <label>CIS CSV File: </label>
          <input type="file" id="cisFile" accept=".csv" />
        </div>
        <button 
          onClick={() => {
            const bssFile = document.getElementById('bssFile').files[0];
            const cisFile = document.getElementById('cisFile').files[0];
            if (bssFile && cisFile) {
              processFiles(bssFile, cisFile);
            } else {
              alert('Please select both files');
            }
          }}
          disabled={loading}
          style={{
            padding: '10px 20px',
            backgroundColor: '#2196F3',
            color: 'white',
            border: 'none',
            borderRadius: '4px',
            cursor: loading ? 'not-allowed' : 'pointer'
          }}
        >
          {loading ? 'Processing...' : 'Analyze Files'}
        </button>
      </div>

      {analysis && (
        <div>
          {/* Summary Dashboard */}
          <div style={{ 
            display: 'grid', 
            gridTemplateColumns: 'repeat(auto-fit, minmax(200px, 1fr))', 
            gap: '15px',
            marginBottom: '30px'
          }}>
            <DashboardCard 
              title="Total Controls" 
              value={analysis.summary.total}
              color="#2196F3"
            />
            <DashboardCard 
              title="Passed" 
              value={analysis.summary.passed}
              color="#4CAF50"
            />
            <DashboardCard 
              title="Failed" 
              value={analysis.summary.failed}
              color="#f44336"
            />
            <DashboardCard 
              title="Risk Score" 
              value={analysis.compliance.riskScore + '%'}
              color="#FF9800"
            />
          </div>

          {/* Compliance by Category Chart */}
          <div style={{ marginBottom: '30px' }}>
            <h2>Compliance by Category</h2>
            <div style={{ overflowX: 'auto' }}>
              <table style={{ width: '100%', borderCollapse: 'collapse' }}>
                <thead>
                  <tr style={{ backgroundColor: '#f5f5f5' }}>
                    <th style={tableHeaderStyle}>Category</th>
                    <th style={tableHeaderStyle}>Total</th>
                    <th style={tableHeaderStyle}>Passed</th>
                    <th style={tableHeaderStyle}>Failed</th>
                    <th style={tableHeaderStyle}>Pass Rate</th>
                    <th style={tableHeaderStyle}>Status</th>
                  </tr>
                </thead>
                <tbody>
                  {Object.entries(analysis.compliance.byCategory).map(([category, stats]) => (
                    <tr key={category}>
                      <td style={tableCellStyle}>{category}</td>
                      <td style={tableCellStyle}>{stats.total}</td>
                      <td style={tableCellStyle}>{stats.passed}</td>
                      <td style={tableCellStyle}>{stats.failed}</td>
                      <td style={tableCellStyle}>{stats.passRate}%</td>
                      <td style={tableCellStyle}>
                        <StatusIndicator passRate={parseFloat(stats.passRate)} />
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>

          {/* Remediation Priority */}
          <div style={{ marginBottom: '30px' }}>
            <h2>Remediation Priority</h2>
            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '20px' }}>
              {Object.entries(analysis.remediation).map(([priority, items]) => (
                <div key={priority} style={{
                  border: `2px solid ${getPriorityColor(priority)}`,
                  borderRadius: '8px',
                  padding: '15px'
                }}>
                  <h3 style={{ color: getPriorityColor(priority), marginTop: 0 }}>
                    {priority.toUpperCase()} ({items.length})
                  </h3>
                  <ul style={{ margin: 0, paddingLeft: '20px' }}>
                    {items.slice(0, 3).map((item, idx) => (
                      <li key={idx} style={{ marginBottom: '5px' }}>
                        {item.id}: {item.title.substring(0, 50)}...
                      </li>
                    ))}
                    {items.length > 3 && <li>...and {items.length - 3} more</li>}
                  </ul>
                </div>
              ))}
            </div>
          </div>

          {/* Export Button */}
          <button 
            onClick={exportToExcel}
            style={{
              padding: '10px 20px',
              backgroundColor: '#4CAF50',
              color: 'white',
              border: 'none',
              borderRadius: '4px',
              cursor: 'pointer',
              fontSize: '16px'
            }}
          >
            Export Detailed Analysis to Excel
          </button>
        </div>
      )}
    </div>
  );
};

// Helper Components
const DashboardCard = ({ title, value, color }) => (
  <div style={{
    backgroundColor: 'white',
    border: `2px solid ${color}`,
    borderRadius: '8px',
    padding: '20px',
    textAlign: 'center'
  }}>
    <h3 style={{ margin: '0 0 10px 0', color: '#666' }}>{title}</h3>
    <div style={{ fontSize: '2em', fontWeight: 'bold', color }}>{value}</div>
  </div>
);

const StatusIndicator = ({ passRate }) => {
  let color, text;
  if (passRate >= 90) {
    color = '#4CAF50';
    text = 'Excellent';
  } else if (passRate >= 70) {
    color = '#FF9800';
    text = 'Good';
  } else if (passRate >= 50) {
    color = '#ff5722';
    text = 'Needs Improvement';
  } else {
    color = '#f44336';
    text = 'Critical';
  }
  
  return (
    <span style={{
      backgroundColor: color,
      color: 'white',
      padding: '2px 8px',
      borderRadius: '4px',
      fontSize: '12px'
    }}>
      {text}
    </span>
  );
};

const getPriorityColor = (priority) => {
  const colors = {
    critical: '#f44336',
    high: '#ff5722',
    medium: '#FF9800',
    low: '#FFC107'
  };
  return colors[priority] || '#666';
};

// Styles
const tableHeaderStyle = {
  padding: '10px',
  textAlign: 'left',
  borderBottom: '2px solid #ddd',
  fontWeight: 'bold'
};

const tableCellStyle = {
  padding: '10px',
  borderBottom: '1px solid #eee'
};

export default BSSCISAnalyzer;
