import * as XLSX from 'xlsx';
import { normalizeName, normalizeLandmark, normalizeMobile, normalizeEmail } from '../utils/normalizer';

const ctx: any = self;

// Session-persistent caches for large spreadsheet datasets
let cachedMasterData: Record<string, unknown>[] = [];
let cachedMasterDataNormalized: Record<string, unknown>[] = [];
let cachedRouteData: Record<string, unknown>[] = [];

// Helper function to extract row values by mapped key or standard fallback columns
const getRowValue = (row: Record<string, unknown>, mappedKey: string, fallbacks: string[]): string => {
  if (mappedKey && row[mappedKey] !== undefined) {
    return String(row[mappedKey]);
  }
  for (const f of fallbacks) {
    if (row[f] !== undefined) return String(row[f]);
  }
  const normalizedMapped = mappedKey ? mappedKey.toLowerCase().replace(/[^a-z0-9]/g, '') : '';
  const keys = Object.keys(row);
  for (const k of keys) {
    const normKey = k.toLowerCase().replace(/[^a-z0-9]/g, '');
    if (normalizedMapped && normKey === normalizedMapped) return String(row[k]);
    if (fallbacks.some(f => normKey === f.toLowerCase().replace(/[^a-z0-9]/g, ''))) {
      return String(row[k]);
    }
  }
  return '';
};

const getFatherMobile = (row: Record<string, unknown>): string => {
  const keys = Object.keys(row);
  for (const k of keys) {
    const nk = k.toLowerCase().replace(/[^a-z0-9]/g, '');
    if (nk.includes('father') && (nk.includes('mobile') || nk.includes('phone') || nk.includes('contact') || nk.includes('no') || nk.includes('num'))) {
      return normalizeMobile(String(row[k]));
    }
  }
  return '';
};

const getMotherMobile = (row: Record<string, unknown>): string => {
  const keys = Object.keys(row);
  for (const k of keys) {
    const nk = k.toLowerCase().replace(/[^a-z0-9]/g, '');
    if (nk.includes('mother') && (nk.includes('mobile') || nk.includes('phone') || nk.includes('contact') || nk.includes('no') || nk.includes('num'))) {
      return normalizeMobile(String(row[k]));
    }
  }
  return '';
};

const getParentEmail = (row: Record<string, unknown>): string => {
  const keys = Object.keys(row);
  for (const k of keys) {
    const nk = k.toLowerCase().replace(/[^a-z0-9]/g, '');
    if (nk.includes('parent') && nk.includes('email')) {
      return normalizeEmail(String(row[k]));
    }
    if (nk.includes('father') && nk.includes('email')) {
      return normalizeEmail(String(row[k]));
    }
    if (nk.includes('mother') && nk.includes('email')) {
      return normalizeEmail(String(row[k]));
    }
    if (nk.includes('poc') && nk.includes('email')) {
      return normalizeEmail(String(row[k]));
    }
    if (nk === 'email' || nk === 'emailaddress') {
      return normalizeEmail(String(row[k]));
    }
  }
  return '';
};

const getRouteStopNo = (row: Record<string, unknown>): string => {
  const keys = Object.keys(row);
  for (const k of keys) {
    const nk = k.toLowerCase().replace(/[^a-z0-9]/g, '');
    if (nk === 'stopno' || nk === 'stopnumber' || nk === 'stop' || nk === 'stopnum') {
      return String(row[k]).trim();
    }
  }
  for (const k of keys) {
    const nk = k.toLowerCase();
    if (nk.includes('stop') && (nk.includes('no') || nk.includes('num') || nk.includes('number'))) {
      return String(row[k]).trim();
    }
  }
  return '';
};

const getRouteReportingTime = (row: Record<string, unknown>): string => {
  const keys = Object.keys(row);
  for (const k of keys) {
    const nk = k.toLowerCase().replace(/[^a-z0-9]/g, '');
    if (nk === 'reportingtime' || nk === 'reporttime' || nk === 'time' || nk === 'pickuptime' || nk === 'droptime' || nk === 'reporting') {
      return String(row[k]).trim();
    }
  }
  for (const k of keys) {
    const nk = k.toLowerCase();
    if (nk.includes('time') || nk.includes('reporting')) {
      return String(row[k]).trim();
    }
  }
  return '';
};

ctx.onmessage = (event: MessageEvent<any>) => {
  const { id, type, payload } = event.data;

  try {
    if (type === 'PARSE_FILE') {
      const { buffer, target } = payload;
      const data = new Uint8Array(buffer);
      const workbook = XLSX.read(data, { type: 'array', cellDates: true });
      
      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];
      const json = XLSX.utils.sheet_to_json<Record<string, unknown>>(sheet, { defval: '' });

      const rowCount = json.length;
      const headers = rowCount > 0 ? Object.keys(json[0]) : [];
      const previewRows = json.slice(0, 10);

      if (target === 'master') {
        cachedMasterData = json;
        cachedMasterDataNormalized = []; // Reset normalized cache
      } else {
        cachedRouteData = json;
      }

      ctx.postMessage({
        id,
        type: 'SUCCESS',
        payload: {
          sheetNames: workbook.SheetNames,
          rowCount,
          columnCount: headers.length,
          headers,
          previewRows,
        },
      });

    } else if (type === 'PROCESS_DATA') {
      const { mappings } = payload;
      const nameKey = mappings?.CommuterName;
      const refKey = mappings?.CommuterRefCode;
      const landmarkKey = mappings?.Landmark;
      const mobileKey = mappings?.POCMobile;
      const emailKey = mappings?.POCEmail;

      const cleanRows: Record<string, unknown>[] = [];
      const refCodes = new Set<string>();

      let duplicateRefCodes = 0;
      let blankRefCodes = 0;
      let blankNames = 0;
      let blankLandmarks = 0;

      cachedMasterData.forEach((row) => {
        const cleanRow = { ...row };

        const rawName = nameKey ? String(row[nameKey] || '') : '';
        const rawRef = refKey ? String(row[refKey] || '') : '';
        const rawLandmark = landmarkKey ? String(row[landmarkKey] || '') : '';
        const rawMobile = mobileKey ? String(row[mobileKey] || '') : '';
        const rawEmail = emailKey ? String(row[emailKey] || '') : '';

        const normName = normalizeName(rawName);
        const normRef = rawRef.trim().replace(/\s+/g, ' ');
        const normLandmark = normalizeLandmark(rawLandmark);
        const normMobile = normalizeMobile(rawMobile);
        const normEmail = normalizeEmail(rawEmail);

        if (nameKey) cleanRow[nameKey] = normName;
        if (refKey) cleanRow[refKey] = normRef;
        if (landmarkKey) cleanRow[landmarkKey] = normLandmark;
        if (mobileKey) cleanRow[mobileKey] = normMobile;
        if (emailKey) cleanRow[emailKey] = normEmail;

        cleanRows.push(cleanRow);

        if (!normName) blankNames++;

        if (!normRef) {
          blankRefCodes++;
        } else {
          const lowerRef = normRef.toLowerCase();
          if (refCodes.has(lowerRef)) {
            duplicateRefCodes++;
          } else {
            refCodes.add(lowerRef);
          }
        }

        if (!normLandmark) blankLandmarks++;
      });

      // Save processed rows in cache
      cachedMasterDataNormalized = cleanRows;

      const errorsCount = blankRefCodes + duplicateRefCodes + blankNames + blankLandmarks;

      ctx.postMessage({
        id,
        type: 'SUCCESS',
        payload: {
          auditSummary: {
            totalRecords: cachedMasterData.length,
            duplicateRefCodes,
            blankRefCodes,
            blankNames,
            blankLandmarks,
            errorsCount,
          },
        },
      });

    } else if (type === 'COMPARE_FILES') {
      const { masterMappings, routeMappings } = payload;

      // Mapped keys for Master File
      const mNameKey = masterMappings?.CommuterName || '';
      const mRefKey = masterMappings?.CommuterRefCode || '';
      const mLandmarkKey = masterMappings?.Landmark || '';

      // Mapped keys for Route File
      const rNameKey = routeMappings?.CommuterName || '';
      const rRefKey = routeMappings?.CommuterRefCode || '';
      const rLandmarkKey = routeMappings?.Landmark || '';
      const rRouteKey = routeMappings?.Route || '';

      // Metrics counters
      const totalRecords = cachedRouteData.length;
      let missingRefCodes = 0;
      let landmarkChanged = 0;
      let landmarkCorrect = 0;
      let landmarkMissing = 0;
      let duplicateNames = 0;
      let errorsFound = 0;

      const comparisonLogs: string[] = [];
      const routeSheetData: Record<string, unknown>[] = [];
      const errorReportData: Record<string, unknown>[] = [];

      const simplifyLandmark = (val: string): string => {
        if (!val) return '';
        return val.toLowerCase().replace(/[^a-z0-9]/g, '').trim();
      };

      interface ProcessedRow {
        originalIndex: number;
        name: string;
        refCode: string;
        routeNo: string;
        routeStopNo: string;
        routeReportingTime: string;
        routeLandmark: string;
        masterName: string;
        masterRefCode: string;
        masterLandmark: string;
        status: 'Valid' | 'Error';
        errorsList: string[];
        reasonsList: string[];
        confidence: number;
      }

      const processedRows: ProcessedRow[] = [];

      // Loop through route rows
      cachedRouteData.forEach((row, index) => {
        // Send progress updates every 5,000 records to keep UI responsive on massive spreadsheets
        if (index > 0 && index % 5000 === 0) {
          ctx.postMessage({
            type: 'COMPARE_PROGRESS',
            payload: Math.round((index / totalRecords) * 100),
          });
        }

        const name = getRowValue(row, rNameKey, ['Commuter Name', 'Name', 'Employee Name', 'Student Name', 'Passenger Name', 'Child Name']).trim();
        const refCode = getRowValue(row, rRefKey, ['Ref Code', 'Reference Code', 'Employee Code', 'Student Code', 'ID', 'Emp ID', 'Registration Number', 'Commuter Ref']).trim();
        const routeLandmark = getRowValue(row, rLandmarkKey, ['Landmark', 'Location', 'Stop', 'Pickup Stop', 'Pickup Location', 'Pickup Landmark']).trim();
        const routeNo = getRowValue(row, rRouteKey, ['Route Number', 'Vehicle Route', 'Route No', 'Route', 'Assigned Route', 'Route ID']).trim();
        const routeStopNo = getRouteStopNo(row);
        const routeReportingTime = getRouteReportingTime(row);

        // Normalizations for route items
        const normName = normalizeName(name);
        const normRef = refCode.trim().replace(/\s+/g, ' ');
        const normLandmark = normalizeLandmark(routeLandmark);

        let matchedMasterRow: Record<string, unknown> | null = null;
        let confidence = 0;
        let logMessage = '';
        let matchError = '';

        // Count missing ref codes in Route List
        if (!normRef) {
          missingRefCodes++;
        }

        // INTELLIGENT MATCHING SCHEMA (Step 2):
        if (normRef) {
          // Priority 1: Match by Ref Code
          const refMatches = cachedMasterDataNormalized.filter((mRow) => {
            const mRef = getRowValue(mRow, mRefKey, ['Ref Code', 'Reference Code', 'Employee Code', 'Student Code', 'ID', 'Emp ID', 'Registration Number', 'Commuter Ref']).trim().toLowerCase();
            return mRef && mRef === normRef.toLowerCase();
          });

          if (refMatches.length === 1) {
            matchedMasterRow = refMatches[0];
            confidence = 100;
            logMessage = `Row ${index + 1}: Matched '${normName}' using Ref Code (${refCode}). Confidence: 100%`;
          } else if (refMatches.length > 1) {
            // Ref Code matches multiple commuters. Disambiguate using Name.
            const nameMatches = refMatches.filter((mRow) => {
              const mName = getRowValue(mRow, mNameKey, ['Commuter Name', 'Name', 'Employee Name', 'Student Name', 'Passenger Name', 'Child Name']).trim().toLowerCase();
              return mName && mName === normName.toLowerCase();
            });

            if (nameMatches.length === 1) {
              matchedMasterRow = nameMatches[0];
              confidence = 95;
              logMessage = `Row ${index + 1}: Matched '${normName}' using Ref Code and Name duplicate resolution. Confidence: 95%`;
            } else {
              // Disambiguate using contact details
              let candidates = nameMatches.length > 0 ? nameMatches : refMatches;

              // Priority 3: Father's Mobile Number
              const rFatherMobile = getFatherMobile(row);
              if (rFatherMobile) {
                const filtered = candidates.filter(mRow => getFatherMobile(mRow) === rFatherMobile);
                if (filtered.length === 1) {
                  matchedMasterRow = filtered[0];
                  confidence = 75;
                } else if (filtered.length > 1) {
                  candidates = filtered;
                }
              }

              // Priority 4: Mother's Mobile Number
              if (!matchedMasterRow) {
                const rMotherMobile = getMotherMobile(row);
                if (rMotherMobile) {
                  const filtered = candidates.filter(mRow => getMotherMobile(mRow) === rMotherMobile);
                  if (filtered.length === 1) {
                    matchedMasterRow = filtered[0];
                    confidence = 65;
                  } else if (filtered.length > 1) {
                    candidates = filtered;
                  }
                }
              }

              // Priority 5: Parent Email
              if (!matchedMasterRow) {
                const rParentEmail = getParentEmail(row);
                if (rParentEmail) {
                  const filtered = candidates.filter(mRow => getParentEmail(mRow) === rParentEmail);
                  if (filtered.length === 1) {
                    matchedMasterRow = filtered[0];
                    confidence = 55;
                  } else if (filtered.length > 1) {
                    candidates = filtered;
                  }
                }
              }

              if (matchedMasterRow) {
                logMessage = `Row ${index + 1}: Matched '${normName}' using Ref Code with contacts fallback. Confidence: ${confidence}%`;
              } else {
                matchError = 'Duplicate Commuter';
                logMessage = `Row ${index + 1}: Ref Code Conflict. Multiple master entries match Ref Code (${refCode}) and cannot be disambiguated.`;
              }
            }
          } else {
            matchError = 'Commuter Not Found';
            logMessage = `Row ${index + 1}: Unmatched Ref Code. Ref Code (${refCode}) not found in Master Data.`;
          }
        } else {
          // Ref Code missing -> Priority 2: Match using Name
          if (normName) {
            const nameMatches = cachedMasterDataNormalized.filter((mRow) => {
              const mName = getRowValue(mRow, mNameKey, ['Commuter Name', 'Name', 'Employee Name', 'Student Name', 'Passenger Name', 'Child Name']).trim().toLowerCase();
              return mName && mName === normName.toLowerCase();
            });

            if (nameMatches.length === 1) {
              matchedMasterRow = nameMatches[0];
              confidence = 85;
              logMessage = `Row ${index + 1}: Matched '${normName}' using Commuter Name. Confidence: 85%`;
            } else if (nameMatches.length > 1) {
              let candidates = nameMatches;

              // Priority 3: Compare Father's Mobile Number
              const rFatherMobile = getFatherMobile(row);
              if (rFatherMobile) {
                const filtered = candidates.filter(mRow => getFatherMobile(mRow) === rFatherMobile);
                if (filtered.length === 1) {
                  matchedMasterRow = filtered[0];
                  confidence = 75;
                } else if (filtered.length > 1) {
                  candidates = filtered;
                }
              }

              // Priority 4: Compare Mother's Mobile Number
              if (!matchedMasterRow) {
                const rMotherMobile = getMotherMobile(row);
                if (rMotherMobile) {
                  const filtered = candidates.filter(mRow => getMotherMobile(mRow) === rMotherMobile);
                  if (filtered.length === 1) {
                    matchedMasterRow = filtered[0];
                    confidence = 65;
                  } else if (filtered.length > 1) {
                    candidates = filtered;
                  }
                }
              }

              // Priority 5: Compare Parent Email
              if (!matchedMasterRow) {
                const rParentEmail = getParentEmail(row);
                if (rParentEmail) {
                  const filtered = candidates.filter(mRow => getParentEmail(mRow) === rParentEmail);
                  if (filtered.length === 1) {
                    matchedMasterRow = filtered[0];
                    confidence = 55;
                  } else if (filtered.length > 1) {
                    candidates = filtered;
                  }
                }
              }

              if (matchedMasterRow) {
                logMessage = `Row ${index + 1}: Matched '${normName}' using Name with contacts fallback. Confidence: ${confidence}%`;
              } else {
                matchError = 'Duplicate Commuter';
                logMessage = `Row ${index + 1}: Duplicate Conflict. Multiple records match Name '${normName}' and cannot be disambiguated.`;
              }
            } else {
              matchError = 'Commuter Not Found';
              logMessage = `Row ${index + 1}: Unmatched Commuter. Name '${normName}' not found in Master Data.`;
            }
          } else {
            matchError = 'Commuter Not Found';
            logMessage = `Row ${index + 1}: Unmatched Commuter. Both Name and Ref Code are missing.`;
          }
        }

        let masterRefCodeVal = '';
        let masterNameVal = '';
        let masterLandmarkVal = '';
        let status: 'Valid' | 'Error' = 'Valid';
        const errorsList: string[] = [];
        const reasonsList: string[] = [];

        // Check for general Route list issues
        if (!routeNo) {
          errorsList.push('Route Missing');
          reasonsList.push('Route Missing');
          status = 'Error';
        }
        if (!routeStopNo) {
          errorsList.push('Invalid Stop Number');
          reasonsList.push('Invalid Stop Number');
          status = 'Error';
        }
        if (!routeReportingTime) {
          errorsList.push('Invalid Reporting Time');
          reasonsList.push('Invalid Reporting Time');
          status = 'Error';
        }

        // Check match results
        if (matchError) {
          errorsList.push(matchError);
          reasonsList.push(matchError);
          status = 'Error';
          if (matchError === 'Duplicate Commuter') {
            duplicateNames++;
          }
        }

        // If Ref Code is missing in Route list but successfully matched by name
        if (!normRef && matchedMasterRow) {
          errorsList.push('Missing Ref Code');
          reasonsList.push('Missing Ref Code');
          // Warning only, does not set status to 'Error' since commuter is successfully matched
        }

        if (matchedMasterRow) {
          // Extract fields from Master Data row
          const mRefKeyMatches = ['Ref Code', 'Reference Code', 'Employee Code', 'Student Code', 'ID', 'Emp ID', 'Registration Number', 'Commuter Ref'];
          const mNameKeyMatches = ['Commuter Name', 'Name', 'Employee Name', 'Student Name', 'Passenger Name', 'Child Name'];
          const mLandmarkMatches = ['Landmark', 'Location', 'Stop', 'Pickup Stop', 'Pickup Location', 'Pickup Landmark'];

          masterRefCodeVal = getRowValue(matchedMasterRow, mRefKey, mRefKeyMatches).trim();
          masterNameVal = getRowValue(matchedMasterRow, mNameKey, mNameKeyMatches).trim();
          masterLandmarkVal = getRowValue(matchedMasterRow, mLandmarkKey, mLandmarkMatches).trim();

          const normMasterLandmark = normalizeLandmark(masterLandmarkVal);

          // Check if parent contact details are completely missing
          const fMobile = getFatherMobile(matchedMasterRow);
          const mMobile = getMotherMobile(matchedMasterRow);
          const pEmail = getParentEmail(matchedMasterRow);
          if (!fMobile && !mMobile && !pEmail) {
            errorsList.push('Missing Parent Contact');
            reasonsList.push('Missing Parent Contact');
          }

          // Landmark Verification
          const simplifiedRoute = simplifyLandmark(normLandmark);
          const simplifiedMaster = simplifyLandmark(normMasterLandmark);

          if (simplifiedRoute === simplifiedMaster) {
            landmarkCorrect++;
          } else {
            landmarkChanged++;
            errorsList.push('Landmark Mismatch');
            reasonsList.push('Landmark Mismatch');
            status = 'Error';
          }
        } else {
          landmarkMissing++;
        }

        if (status === 'Error') {
          errorsFound++;
        }

        comparisonLogs.push(logMessage || `Row ${index + 1}: Match failed.`);

        processedRows.push({
          originalIndex: index,
          name: name || 'N/A',
          refCode: refCode || 'N/A',
          routeNo,
          routeStopNo,
          routeReportingTime,
          routeLandmark,
          masterName: masterNameVal,
          masterRefCode: masterRefCodeVal,
          masterLandmark: masterLandmarkVal,
          status,
          errorsList,
          reasonsList,
          confidence
        });
      });

      // Route Consistency Checks (Step 5)
      const routeGroups: Record<string, typeof processedRows> = {};
      processedRows.forEach((row) => {
        if (row.status === 'Valid' && row.routeNo && row.routeLandmark) {
          const key = `${row.routeNo.toLowerCase()}|${simplifyLandmark(row.routeLandmark)}`;
          if (!routeGroups[key]) {
            routeGroups[key] = [];
          }
          routeGroups[key].push(row);
        }
      });

      Object.keys(routeGroups).forEach((key) => {
        const group = routeGroups[key];
        if (group.length <= 1) return;

        // Modes counts
        const stopCounts = new Map<string, number>();
        const timeCounts = new Map<string, number>();

        group.forEach((r) => {
          stopCounts.set(r.routeStopNo, (stopCounts.get(r.routeStopNo) || 0) + 1);
          timeCounts.set(r.routeReportingTime, (timeCounts.get(r.routeReportingTime) || 0) + 1);
        });

        let maxStopCount = 0;
        let expectedStop = '';
        stopCounts.forEach((count, val) => {
          if (count > maxStopCount) {
            maxStopCount = count;
            expectedStop = val;
          }
        });

        let maxTimeCount = 0;
        let expectedTime = '';
        timeCounts.forEach((count, val) => {
          if (count > maxTimeCount) {
            maxTimeCount = count;
            expectedTime = val;
          }
        });

        group.forEach((r) => {
          const isStopInconsistent = (maxStopCount === 1 && group.length > 1) || (r.routeStopNo !== expectedStop);
          const isTimeInconsistent = (maxTimeCount === 1 && group.length > 1) || (r.routeReportingTime !== expectedTime);

          if (isStopInconsistent) {
            r.errorsList.push('Inconsistent Stop Number');
            r.reasonsList.push('Inconsistent Stop Number');
            if (r.status !== 'Error') {
              r.status = 'Error';
              errorsFound++;
            }
          }

          if (isTimeInconsistent) {
            r.errorsList.push('Inconsistent Reporting Time');
            r.reasonsList.push('Inconsistent Reporting Time');
            if (r.status !== 'Error') {
              r.status = 'Error';
              errorsFound++;
            }
          }
        });
      });

      // Filter and Sort Valid Rows for the Route Sheet
      const validRows = processedRows.filter(r => r.status === 'Valid');
      validRows.sort((a, b) => {
        const routeCompare = a.routeNo.localeCompare(b.routeNo, undefined, { numeric: true, sensitivity: 'base' });
        if (routeCompare !== 0) return routeCompare;

        const stopCompare = a.routeStopNo.localeCompare(b.routeStopNo, undefined, { numeric: true });
        if (stopCompare !== 0) return stopCompare;

        return a.masterName.localeCompare(b.masterName, undefined, { sensitivity: 'base' });
      });

      validRows.forEach((row) => {
        routeSheetData.push({
          'CommuterRefCode': row.masterRefCode,
          'Commuter Name': row.masterName,
          'Landmark Name': row.masterLandmark,
          'Route No': row.routeNo,
          'Stop No': row.routeStopNo,
          'Route Type': 'Bus',
          'Reporting Time': row.routeReportingTime
        });
      });

      // Build Error Report entries
      processedRows.forEach((row) => {
        row.reasonsList.forEach((reason) => {
          let suggestedFix = '';
          if (reason === 'Commuter Not Found') {
            suggestedFix = 'Verify commuter name and reference code match Master Data';
          } else if (reason === 'Duplicate Commuter') {
            suggestedFix = 'Provide unique Father/Mother Mobile or Parent Email in Master and Route List';
          } else if (reason === 'Landmark Mismatch') {
            suggestedFix = `Verify and align Landmark (Expected: "${row.masterLandmark}", Actual: "${row.routeLandmark}")`;
          } else if (reason === 'Missing Ref Code') {
            suggestedFix = 'Provide Commuter Reference Code in Route List';
          } else if (reason === 'Route Missing') {
            suggestedFix = 'Assign a valid Route Number in Route List';
          } else if (reason === 'Invalid Stop Number') {
            suggestedFix = 'Provide a valid Stop Number in Route List';
          } else if (reason === 'Invalid Reporting Time') {
            suggestedFix = 'Provide a valid Reporting Time (HH:MM) in Route List';
          } else if (reason === 'Missing Parent Contact') {
            suggestedFix = 'Provide Father/Mother Mobile or Parent Email in Master Data';
          } else if (reason === 'Inconsistent Stop Number') {
            suggestedFix = 'Align Stop Number for same Route and Landmark';
          } else if (reason === 'Inconsistent Reporting Time') {
            suggestedFix = 'Align Reporting Time for same Route and Landmark';
          } else {
            suggestedFix = 'Verify record data details';
          }

          errorReportData.push({
            'Row Number': row.originalIndex + 2,
            'Commuter Name': row.name,
            'Ref Code': row.refCode,
            'Reason': reason,
            'Suggested Fix': suggestedFix
          });
        });
      });

      // Calculate Average Confidence
      let sumConfidence = 0;
      validRows.forEach(r => sumConfidence += r.confidence);
      const averageConfidence = validRows.length > 0 ? Math.round(sumConfidence / validRows.length) : 0;

      // Generate Route Sheet workbook
      const routeWorkbook = XLSX.utils.book_new();
      const routeWorksheet = XLSX.utils.json_to_sheet(routeSheetData);
      XLSX.utils.book_append_sheet(routeWorkbook, routeWorksheet, 'Route Sheet');
      const routeBuffer = XLSX.write(routeWorkbook, { bookType: 'xlsx', type: 'array' }) as ArrayBuffer;

      // Generate Error Report workbook
      const errorWorkbook = XLSX.utils.book_new();
      const errorWorksheet = XLSX.utils.json_to_sheet(
        errorReportData.length > 0
          ? errorReportData
          : [{ 'Message': 'No validation errors or landmark mismatches found in route sheet list.' }]
      );
      XLSX.utils.book_append_sheet(errorWorkbook, errorWorksheet, 'Anomalies Report');
      const errorBuffer = XLSX.write(errorWorkbook, { bookType: 'xlsx', type: 'array' }) as ArrayBuffer;

      // Send buffers and results back to main thread using Transferables
      ctx.postMessage(
        {
          id,
          type: 'SUCCESS',
          payload: {
            metrics: {
              totalRecords,
              matchedRecords: validRows.length,
              missingRefCodes,
              landmarkChanged,
              landmarkCorrect,
              landmarkMissing,
              duplicateNames,
              errorsFound,
              averageConfidence,
            },
            routeSheetBuffer: routeBuffer,
            errorReportBuffer: errorBuffer,
            comparisonLogs,
          },
        },
        [routeBuffer, errorBuffer]
      );

    } else if (type === 'CLEAR_CACHE') {
      const { target } = payload;
      if (target === 'master' || target === 'both') {
        cachedMasterData = [];
        cachedMasterDataNormalized = [];
      }
      if (target === 'route' || target === 'both') {
        cachedRouteData = [];
      }
      
      ctx.postMessage({ id, type: 'SUCCESS', payload: null });
    }
  } catch (error: any) {
    const rawMessage = error instanceof Error ? error.message : String(error);
    const lowerMessage = rawMessage.toLowerCase();
    let friendlyMessage = rawMessage;

    if (lowerMessage.includes('password') || lowerMessage.includes('decrypt') || lowerMessage.includes('encrypt')) {
      friendlyMessage = 'This Excel file is password-protected or encrypted. Please remove password protection and try again.';
    } else if (lowerMessage.includes('corrupt') || lowerMessage.includes('zip') || lowerMessage.includes('invalid') || lowerMessage.includes('truncated')) {
      friendlyMessage = 'The Excel file is corrupted or uses an unsupported spreadsheet format.';
    } else if (lowerMessage.includes('empty') || lowerMessage.includes('no worksheets')) {
      friendlyMessage = 'The spreadsheet is empty or contains no valid records.';
    } else {
      friendlyMessage = `Failed to process workbook: ${rawMessage}`;
    }

    ctx.postMessage({
      id,
      type: 'ERROR',
      payload: friendlyMessage,
    });
  }
};
