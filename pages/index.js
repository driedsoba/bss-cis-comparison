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

  // Enhanced string similarity with normalization improvements
  const calculateSimilarity = (str1, str2) => {
    if (!str1 || !str2) return 0;
    
    const normalize = (s) => {
      return s
        .toLowerCase()
        .normalize('NFD') // Handle Unicode normalization
        .replace(/[\u0300-\u036f]/g, '') // Remove diacritics
        .replace(/[''""]/g, "'") // Normalize quotes/apostrophes
        .replace(/[^\w\s]/g, ' ') // Replace punctuation with spaces
        .replace(/\s+/g, ' ') // Collapse multiple spaces
        .trim();
    };
    
    const s1 = normalize(str1);
    const s2 = normalize(str2);
    
    if (s1 === s2) return 1.0;
    if (s1.length < 2 || s2.length < 2) return 0;
    
    // Use trigrams for better accuracy on longer strings
    const gramSize = Math.min(3, Math.max(2, Math.min(s1.length, s2.length) / 3));
    
    const getGrams = (str, size) => {
      const grams = new Set();
      for (let i = 0; i <= str.length - size; i++) {
        grams.add(str.substr(i, size));
      }
      return grams;
    };
    
    const grams1 = getGrams(s1, gramSize);
    const grams2 = getGrams(s2, gramSize);
    
    const intersection = new Set([...grams1].filter(x => grams2.has(x)));
    const union = new Set([...grams1, ...grams2]);
    
    // Jaccard similarity (more balanced than Dice for this use case)
    return intersection.size / union.size;
  };

  const readBSSFile = async (file) => {
    const data = await file.arrayBuffer();
    const workbook = XLSX.read(data);

    const sheetName = Object.keys(workbook.Sheets).find((name) =>
      name.toLowerCase().includes('settings') || name.toLowerCase().includes('server')
    );
    if (!sheetName) throw new Error("Settings sheet not found");

    const worksheet = workbook.Sheets[sheetName];
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

    let headerRowIdx = jsonData.findIndex((row) =>
      row.some((cell) => cell && cell.toString().includes('CIS #'))
    );
    if (headerRowIdx === -1) throw new Error("Header row not found");

    const headers = jsonData[headerRowIdx];
    const data_rows = jsonData
      .slice(headerRowIdx + 1)
      .filter((row) => row[headers.indexOf('CIS #')])
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
          let headerIdx = data.findIndex((row) =>
            row[0] === 'check_id' || (row.length > 1 && row[0].includes('check_id'))
          );
          if (headerIdx > 0) data = data.slice(headerIdx);

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
    const SIMILARITY_THRESHOLD = 0.75; // Raised from 0.7 based on analysis
    const EXACT_MATCH_THRESHOLD = 0.95; // For nearly identical strings

    // Build fast lookup maps by ID
    const bssMap = new Map(bssData.map((row) => [row['CIS #']?.toString(), row]));
    const cisMap = new Map(cisData.map((row) => [row.check_id?.toString(), row]));

    const bssIds = new Set(bssData.map((row) => row['CIS #']?.toString()));
    const cisIds = new Set(cisData.map((row) => row.check_id?.toString()));
    const allIds = new Set([...bssIds, ...cisIds]);

    allIds.forEach((checkId) => {
      const bssRow = bssMap.get(checkId);
      const cisRow = cisMap.get(checkId);

      const result = {
        Check_ID: checkId,
        In_BSS: bssRow ? 'Yes' : 'No',
        In_CIS_Scan: cisRow ? 'Yes' : 'No',
        Non_Compliance_Reason: ''
      };

      // 1) Pull in BSS fields (strip "(L#)" from title)
      if (bssRow) {
        const titleCol = Object.keys(bssRow).find((key) => key.includes('Setting Title'));
        const appCol = Object.keys(bssRow).find((key) => key.includes('Setting Applicability'));
        const cisRecCol = Object.keys(bssRow).find((key) =>
          key.includes('CIS Recommended Value')
        );

        const rawBssTitle = bssRow[titleCol] || '';
        result.BSS_Title = stripLPrefix(rawBssTitle);
        result.BSS_Category = bssRow.Category || '';
        result.Setting_Applicability = bssRow[appCol] || '';
        result.CIS_Recommended_Value = bssRow[cisRecCol] || '';
        result.Synapxe_Value = bssRow['Synapxe Value'] || '';
        result.Synapxe_Exceptions = bssRow['Synapxe Exceptions'] || '';
        result.Change_Description_Remarks = bssRow['Change Description / Remarks'] || '';
        result.BSS_ID = bssRow['BSS ID'] || bssRow['BSS #'] || checkId;
      } else {
        result.BSS_Title = '';
        result.BSS_Category = '';
        result.Setting_Applicability = '';
        result.CIS_Recommended_Value = '';
        result.Synapxe_Value = '';
        result.Synapxe_Exceptions = '';
        result.Change_Description_Remarks = '';
        result.BSS_ID = checkId;
      }

      // 2) Pull in CIS fields (strip "(L#)" from CIS title)
      if (cisRow) {
        const rawCisTitle = cisRow.title || '';
        result.CIS_Title = stripLPrefix(rawCisTitle);
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

      // 4) Enhanced title similarity check with debug info
      let titleMismatch = false;
      let similarityScore = 1.0;
      let mismatchReason = '';
      
      if (bssRow && cisRow && result.BSS_Title && result.CIS_Title) {
        similarityScore = calculateSimilarity(result.BSS_Title, result.CIS_Title);
        
        if (similarityScore >= EXACT_MATCH_THRESHOLD) {
          titleMismatch = false;
        } else if (similarityScore < SIMILARITY_THRESHOLD) {
          titleMismatch = true;
          mismatchReason = 'Low similarity';
        } else {
          // In the gray area - check for common false positives
          const normalizedBss = result.BSS_Title.toLowerCase().normalize('NFD').replace(/[''""]/g, "'");
          const normalizedCis = result.CIS_Title.toLowerCase().normalize('NFD').replace(/[''""]/g, "'");
          
          if (normalizedBss === normalizedCis) {
            titleMismatch = false; // Override - they're actually identical
          } else {
            titleMismatch = true;
            mismatchReason = 'Minor differences';
          }
        }
      }
      
      result.Title_Similarity = similarityScore;
      result.Title_Mismatch = titleMismatch ? 'Yes' : 'No';
      result.Mismatch_Reason = mismatchReason;

      // 5) Enhanced remark handling (treat "NIL", "None", "N/A" as no remark)
      let hasRemark = false;
      if (bssRow) {
        const rawRemark = (result.Change_Description_Remarks || '').toString().trim();
        const lowerRemark = rawRemark.toLowerCase();
        const nullValues = ['', 'nil', 'none', 'n/a', 'na', 'not applicable'];
        if (rawRemark && !nullValues.includes(lowerRemark)) {
          hasRemark = true;
        }
      }
      result.Has_Remark = hasRemark ? 'Yes' : 'No';

      // 6) Enhanced exception handling  
      let hasException = false;
      if (bssRow) {
        const rawException = (result.Synapxe_Exceptions || '').toString().trim();
        const lowerException = rawException.toLowerCase();
        const nullValues = ['', 'nil', 'none', 'n/a', 'na', 'not applicable'];
        if (rawException && !nullValues.includes(lowerException)) {
          hasException = true;
        }
      }
      result.Has_Exception = hasException ? 'Yes' : 'No';

      // 7) Determine final Compliance_Status & Non_Compliance_Reason
      if (bssRow && cisRow) {
        if (result.CIS_Status === 'Failed') {
          result.Compliance_Status = 'Non-Compliant';
          result.Non_Compliance_Reason = 'Scan Failed';
        } else if (result.CIS_Status === 'Passed') {
          if (titleMismatch) {
            result.Compliance_Status = 'Non-Compliant';
            result.Non_Compliance_Reason = `Title Mismatch (${mismatchReason})`;
          } else if (hasRemark) {
            result.Compliance_Status = 'Non-Compliant';
            result.Non_Compliance_Reason = 'Has Remark';
          } else if (hasException) {
            result.Compliance_Status = 'Non-Compliant';
            result.Non_Compliance_Reason = 'Has Exception';
          } else {
            result.Compliance_Status = 'Compliant';
            result.Non_Compliance_Reason = '';
          }
        } else {
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
          result.Compliance_Status = 'Not Tested';
          result.Non_Compliance_Reason = 'No BSS Policy';
        }
      } else {
        result.Compliance_Status = '';
        result.Non_Compliance_Reason = '';
      }

      results.push(result);
    });

    // Enhanced duplicate detection with fuzzy matching
    const titleGroups = new Map();
    results.forEach((row) => {
      if (row.BSS_Title) {
        const normalizedTitle = row.BSS_Title.toLowerCase()
          .normalize('NFD')
          .replace(/[''""]/g, "'")
          .replace(/[^\w\s]/g, ' ')
          .replace(/\s+/g, ' ')
          .trim();
        
        if (!titleGroups.has(normalizedTitle)) {
          titleGroups.set(normalizedTitle, []);
        }
        titleGroups.get(normalizedTitle).push(row.Check_ID);
      }
    });

    results.forEach((row) => {
      const normalizedTitle = row.BSS_Title?.toLowerCase()
        ?.normalize('NFD')
        ?.replace(/[''""]/g, "'")
        ?.replace(/[^\w\s]/g, ' ')
        ?.replace(/\s+/g, ' ')
        ?.trim();
      
      const group = titleGroups.get(normalizedTitle) || [];
      row.Duplicate_Title = group.length > 1 ? 'Yes' : 'No';
      if (group.length > 1) {
        row.Duplicate_IDs = group.filter(id => id !== row.Check_ID).join(', ');
      }
    });

    return results.sort((a, b) => a.Check_ID.localeCompare(b.Check_ID));
  };

  const generateExcelReport = (data) => {
    const wb = XLSX.utils.book_new();

    // 1) Full Comparison sheet
    const fullData = data.map((row) => ({
      'Check ID': row.Check_ID,
      'BSS ID': row.BSS_ID,
      'In BSS': row.In_BSS,
      'In CIS Scan': row.In_CIS_Scan,
      'BSS Category': row.BSS_Category,
      'BSS Title': row.BSS_Title,
      'CIS Title': row.CIS_Title,
      'Title Similarity': row.Title_Similarity?.toFixed(2) || '',
      'Title Mismatch': row.Title_Mismatch,
      'Duplicate Title': row.Duplicate_Title,
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
        Metric: 'Title Mismatches',
        Count: data.filter((r) => r.Title_Mismatch === 'Yes').length
      },
      {
        Metric: 'Duplicate Titles',
        Count: data.filter((r) => r.Duplicate_Title === 'Yes').length
      },
      {
        Metric: 'Controls with Remarks',
        Count: data.filter((r) => {
          const rawRemark = (r.Change_Description_Remarks || '').toString().trim();
          const lowerRemark = rawRemark.toLowerCase();
          return lowerRemark !== '' && lowerRemark !== 'nil' && lowerRemark !== 'none';
        }).length
      },
      {
        Metric: 'Controls with Exceptions',
        Count: data.filter((r) => {
          const rawExc = (r.Synapxe_Exceptions || '').toString().trim().toLowerCase();
          return rawExc !== '' && rawExc !== 'nil' && rawExc !== 'none';
        }).length
      },
      { Metric: 'Failed Controls', Count: data.filter((r) => r.CIS_Status === 'Failed').length },
      { Metric: 'Passed Controls', Count: data.filter((r) => r.CIS_Status === 'Passed').length },
      { Metric: 'Skipped Controls', Count: data.filter((r) => r.CIS_Status === 'Skipped').length }
    ];
    const summaryWs = XLSX.utils.json_to_sheet(summary);
    XLSX.utils.book_append_sheet(wb, summaryWs, 'Summary');

    // 3) Title Mismatches sheet
    const mismatchData = data.filter((r) => r.Title_Mismatch === 'Yes');
    if (mismatchData.length > 0) {
      const mismatchWs = XLSX.utils.json_to_sheet(
        mismatchData.map((row) => ({
          'Check ID': row.Check_ID,
          'BSS Title': row.BSS_Title,
          'CIS Title': row.CIS_Title,
          'Similarity Score': row.Title_Similarity?.toFixed(2) || '',
          'Compliance Status': row.Compliance_Status
        }))
      );
      XLSX.utils.book_append_sheet(wb, mismatchWs, 'Title Mismatches');
    }

    // 4) Controls with Remarks
    const remarksData = data.filter((r) => {
      const rawRemark = (r.Change_Description_Remarks || '').toString().trim();
      const lowerRemark = rawRemark.toLowerCase();
      return lowerRemark !== '' && lowerRemark !== 'nil' && lowerRemark !== 'none';
    });
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

    // 5) Controls with Exceptions
    const exceptionsData = data.filter((r) => {
      const rawExc = (r.Synapxe_Exceptions || '').toString().trim().toLowerCase();
      return rawExc !== '' && rawExc !== 'nil' && rawExc !== 'none';
    });
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

    // 6) Non-Compliant Details
    const nonCompliantData = data.filter((r) => r.Compliance_Status === 'Non-Compliant');
    if (nonCompliantData.length > 0) {
      const nonCompliantWs = XLSX.utils.json_to_sheet(
        nonCompliantData.map((row) => ({
          'Check ID': row.Check_ID,
          'BSS Category': row.BSS_Category,
          'BSS Title': row.BSS_Title,
          'CIS Title': row.CIS_Title,
          'Title Similarity': row.Title_Similarity?.toFixed(2) || '',
          'Title Mismatch': row.Title_Mismatch,
          'Has Remark': row.Has_Remark,
          'CIS Status': row.CIS_Status,
          'Non_Compliance_Reason': row.Non_Compliance_Reason,
          'Change Description / Remarks': row.Change_Description_Remarks
        }))
      );
      XLSX.utils.book_append_sheet(wb, nonCompliantWs, 'Non-Compliant Details');
    }

    // 7) FINDINGS Sheet - Categorized Analysis
    const findingsData = [];
    
    // Category 1: Steps done - Controls that are compliant
    const compliantControls = data.filter(r => r.Compliance_Status === 'Compliant');
    if (compliantControls.length > 0) {
      findingsData.push({
        'Category': 'Steps done',
        'Count': compliantControls.length,
        'Description': 'Controls that are compliant and properly configured',
        'Action Required': 'None - Continue monitoring',
        'Priority': 'Low'
      });
    }

    // Category 2: Check duplicates in BSS
    const duplicateControls = data.filter(r => r.Duplicate_Title === 'Yes');
    if (duplicateControls.length > 0) {
      findingsData.push({
        'Category': 'Check duplicates in BSS',
        'Count': duplicateControls.length,
        'Description': 'Controls with duplicate titles across different IDs',
        'Action Required': 'Review and consolidate duplicate controls',
        'Priority': 'Medium'
      });
    }

    // Category 3: Check for CIS ID (SCAN) in BSS but not in CIS (SCAN)
    const bssOnlyControls = data.filter(r => r.In_BSS === 'Yes' && r.In_CIS_Scan === 'No');
    if (bssOnlyControls.length > 0) {
      findingsData.push({
        'Category': 'Check for CIS ID (SCAN) in BSS but not in CIS (SCAN)',
        'Count': bssOnlyControls.length,
        'Description': 'Controls defined in baseline but not scanned by CIS',
        'Action Required': 'Verify scan coverage or update baseline scope',
        'Priority': 'Medium'
      });
    }

    // Category 4: Check for BSS ID (SCAN) in BSS but not in CIS (SCAN)
    const orphanedBssControls = data.filter(r => 
      r.In_BSS === 'Yes' && r.In_CIS_Scan === 'No' && 
      (r.Setting_Applicability || '').toLowerCase().includes('server')
    );
    if (orphanedBssControls.length > 0) {
      findingsData.push({
        'Category': 'Check for BSS ID (SCAN) in BSS but not in CIS (SCAN)',
        'Count': orphanedBssControls.length,
        'Description': 'Server-applicable controls not included in scan',
        'Action Required': 'Include in next scan or mark as manual check',
        'Priority': 'High'
      });
    }

    // Category 5: Check for Red Highlighted CIS ID (SCAN)
    const failedControls = data.filter(r => r.CIS_Status === 'Failed');
    if (failedControls.length > 0) {
      findingsData.push({
        'Category': 'Check for Red Highlighted CIS ID (SCAN)',
        'Count': failedControls.length,
        'Description': 'Controls that failed the CIS compliance scan',
        'Action Required': 'Immediate remediation required',
        'Priority': 'Critical'
      });
    }

    // Category 6: Check for Yellow Highlighted CIS ID (SCAN)
    const titleMismatchControls = data.filter(r => r.Title_Mismatch === 'Yes');
    if (titleMismatchControls.length > 0) {
      findingsData.push({
        'Category': 'Check for Yellow Highlighted CIS ID (SCAN)',
        'Count': titleMismatchControls.length,
        'Description': 'Controls with title mismatches between BSS and CIS',
        'Action Required': 'Review and align control definitions',
        'Priority': 'Medium'
      });
    }

    // Category 7: Check for Grey Highlighted CIS ID (SCAN)
    const skippedControls = data.filter(r => r.CIS_Status === 'Skipped');
    if (skippedControls.length > 0) {
      findingsData.push({
        'Category': 'Check for Grey Highlighted CIS ID (SCAN)',
        'Count': skippedControls.length,
        'Description': 'Controls that were skipped during the scan',
        'Action Required': 'Review scan configuration and applicability',
        'Priority': 'Low'
      });
    }

    // Add the FINDINGS sheet
    if (findingsData.length > 0) {
      const findingsWs = XLSX.utils.json_to_sheet(findingsData);
      
      // Add some styling information (Excel will need to be manually formatted)
      const range = XLSX.utils.decode_range(findingsWs['!ref']);
      
      // Set column widths
      findingsWs['!cols'] = [
        { width: 40 }, // Category
        { width: 10 }, // Count
        { width: 60 }, // Description
        { width: 50 }, // Action Required
        { width: 15 }  // Priority
      ];
      
      XLSX.utils.book_append_sheet(wb, findingsWs, 'FINDINGS');
    }

    // 8) Detailed Findings by Category
    const detailedFindingsWs = XLSX.utils.book_new();
    
    // Failed Controls Detail
    if (failedControls.length > 0) {
      const failedDetails = failedControls.map(row => ({
        'Category': 'CRITICAL - Failed Controls',
        'Check ID': row.Check_ID,
        'BSS Title': row.BSS_Title,
        'CIS Title': row.CIS_Title,
        'BSS Category': row.BSS_Category,
        'Synapxe Value': row.Synapxe_Value,
        'Failed Instances': row.Failed_Instances,
        'Remarks': row.Change_Description_Remarks,
        'Action': 'Immediate remediation required'
      }));
      
      const failedWs = XLSX.utils.json_to_sheet(failedDetails);
      XLSX.utils.book_append_sheet(wb, failedWs, 'Critical Issues');
    }

    // Title Mismatch Details
    if (titleMismatchControls.length > 0) {
      const mismatchDetails = titleMismatchControls.map(row => ({
        'Category': 'MEDIUM - Title Mismatches',
        'Check ID': row.Check_ID,
        'BSS Title': row.BSS_Title,
        'CIS Title': row.CIS_Title,
        'Similarity Score': row.Title_Similarity?.toFixed(2) || '',
        'BSS Category': row.BSS_Category,
        'Mismatch Reason': row.Mismatch_Reason || 'Title difference',
        'Action': 'Review and align control definitions'
      }));
      
      const mismatchWs = XLSX.utils.json_to_sheet(mismatchDetails);
      XLSX.utils.book_append_sheet(wb, mismatchWs, 'Title Mismatches');
    }

    // Coverage Gaps
    if (bssOnlyControls.length > 0) {
      const coverageDetails = bssOnlyControls.map(row => ({
        'Category': 'MEDIUM - Coverage Gaps',
        'Check ID': row.Check_ID,
        'BSS Title': row.BSS_Title,
        'BSS Category': row.BSS_Category,
        'Setting Applicability': row.Setting_Applicability,
        'Synapxe Value': row.Synapxe_Value,
        'Remarks': row.Change_Description_Remarks,
        'Action': 'Include in scan or mark as manual verification'
      }));
      
      const coverageWs = XLSX.utils.json_to_sheet(coverageDetails);
      XLSX.utils.book_append_sheet(wb, coverageWs, 'Coverage Gaps');
    }

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
    statBox: {
      background: '#e3f2fd',
      padding: '10px',
      borderRadius: '4px',
      fontSize: '0.9rem'
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
    remarkCount = 0,
    duplicateCount = 0;

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

    // Title mismatches / remarks / duplicates
    mismatchCount = results.filter((r) => r.Title_Mismatch === 'Yes').length;
    remarkCount = results.filter((r) => r.Has_Remark === 'Yes').length;
    duplicateCount = results.filter((r) => r.Duplicate_Title === 'Yes').length;
  }

  return (
    <div style={styles.container}>
      <h1 style={styles.h1}>BSS-CIS Comparison Tool with Fuzzy Matching</h1>

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

          {/* Enhanced badges for Title issues */}
          <div
            style={{
              margin: '20px 0',
              display: 'flex',
              gap: '12px',
              justifyContent: 'center',
              flexWrap: 'wrap'
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
                background: '#fef',
                padding: '6px 10px',
                borderRadius: '4px',
                fontSize: '0.9rem'
              }}
            >
              Duplicate Titles: {duplicateCount}
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

          {/* Enhanced Dashboard with Findings Categories */}
          <div style={{ margin: '20px 0' }}>
            <h3 style={{ fontSize: '1.1rem', textAlign: 'center' }}>Findings by Category</h3>
            <div style={{ 
              display: 'grid', 
              gridTemplateColumns: 'repeat(auto-fit, minmax(250px, 1fr))', 
              gap: '10px',
              margin: '15px 0'
            }}>
              {/* Critical Issues */}
              <div style={{
                background: '#ffebee',
                border: '2px solid #f44336',
                padding: '10px',
                borderRadius: '4px'
              }}>
                <h4 style={{ margin: '0 0 5px 0', color: '#d32f2f' }}>🔴 Critical Issues</h4>
                <div>Failed Controls: {results.filter(r => r.CIS_Status === 'Failed').length}</div>
                <div style={{ fontSize: '0.8rem', color: '#666' }}>Immediate remediation required</div>
              </div>

              {/* Medium Priority */}
              <div style={{
                background: '#fff3e0',
                border: '2px solid #ff9800',
                padding: '10px',
                borderRadius: '4px'
              }}>
                <h4 style={{ margin: '0 0 5px 0', color: '#f57c00' }}>🟡 Medium Priority</h4>
                <div>Title Mismatches: {mismatchCount}</div>
                <div>Coverage Gaps: {results.filter(r => r.In_BSS === 'Yes' && r.In_CIS_Scan === 'No').length}</div>
                <div style={{ fontSize: '0.8rem', color: '#666' }}>Review and alignment needed</div>
              </div>

              {/* Low Priority */}
              <div style={{
                background: '#f3e5f5',
                border: '2px solid #9c27b0',
                padding: '10px',
                borderRadius: '4px'
              }}>
                <h4 style={{ margin: '0 0 5px 0', color: '#7b1fa2' }}>🟣 Low Priority</h4>
                <div>Skipped Controls: {results.filter(r => r.CIS_Status === 'Skipped').length}</div>
                <div>Duplicates: {duplicateCount}</div>
                <div style={{ fontSize: '0.8rem', color: '#666' }}>Monitor and review periodically</div>
              </div>

              {/* Completed */}
              <div style={{
                background: '#e8f5e8',
                border: '2px solid '#4caf50',
                padding: '10px',
                borderRadius: '4px'
              }}>
                <h4 style={{ margin: '0 0 5px 0', color: '#388e3c' }}>✅ Completed</h4>
                <div>Compliant Controls: {numPassed}</div>
                <div style={{ fontSize: '0.8rem', color: '#666' }}>No action required</div>
              </div>
            </div>
          </div>

          {/* Action Items Summary */}
          <div style={{ margin: '20px 0' }}>
            <h3 style={{ fontSize: '1.1rem' }}>Action Items by Priority</h3>
            <div style={{ fontSize: '0.9rem' }}>
              {results.filter(r => r.CIS_Status === 'Failed').length > 0 && (
                <div style={{ color: '#d32f2f', marginBottom: '5px' }}>
                  🚨 <strong>CRITICAL:</strong> {results.filter(r => r.CIS_Status === 'Failed').length} failed controls need immediate remediation
                </div>
              )}
              {mismatchCount > 0 && (
                <div style={{ color: '#f57c00', marginBottom: '5px' }}>
                  ⚠️ <strong>MEDIUM:</strong> {mismatchCount} title mismatches need review
                </div>
              )}
              {results.filter(r => r.In_BSS === 'Yes' && r.In_CIS_Scan === 'No').length > 0 && (
                <div style={{ color: '#f57c00', marginBottom: '5px' }}>
                  📋 <strong>MEDIUM:</strong> {results.filter(r => r.In_BSS === 'Yes' && r.In_CIS_Scan === 'No').length} controls not scanned - verify coverage
                </div>
              )}
              {duplicateCount > 0 && (
                <div style={{ color: '#7b1fa2', marginBottom: '5px' }}>
                  📄 <strong>LOW:</strong> {duplicateCount} duplicate titles to consolidate
                </div>
              )}
            </div>
          </div>

          {/* Download Button */}
          <button onClick={() => generateExcelReport(results)} style={styles.button}>
            Download Excel Report
          </button>

          {/* Enhanced Table with Similarity */}
          <div style={styles.tableContainer}>
            <table style={styles.table}>
              <thead>
                <tr>
                  <th style={styles.th}>Check ID</th>
                  <th style={styles.th}>Category</th>
                  <th style={styles.th}>Compliance Status</th>
                  <th style={styles.th}>Title Similarity</th>
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
                    <td style={styles.td}>
                      <span style={{
                        color: row.Title_Similarity < 0.7 ? 'red' : 
                               row.Title_Similarity < 0.85 ? 'orange' : 'green'
                      }}>
                        {row.Title_Similarity ? row.Title_Similarity.toFixed(2) : '-'}
                      </span>
                    </td>
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
