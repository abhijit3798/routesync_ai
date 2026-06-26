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
      const { masterMappings, routeMappings, settings } = payload;
      const autoGenerateStops = settings?.autoGenerateStops !== false;
      const autoGenerateTimes = settings?.autoGenerateTimes !== false;
      const tripStartTime = settings?.tripStartTime || '08:00';
      const tripEndTime = settings?.tripEndTime || '09:00';

      const comparisonLogs: string[] = [];
      comparisonLogs.push(`Comparative validator started. Time window settings: ${tripStartTime} to ${tripEndTime}.`);

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

      const routeSheetData: Record<string, unknown>[] = [];
      const errorReportData: Record<string, unknown>[] = [];

      const simplifyLandmark = (val: string): string => {
        if (!val) return '';
        return val.toLowerCase().replace(/[^a-z0-9]/g, '').trim();
      };

      const getMode = (arr: string[]): string => {
        const counts: Record<string, number> = {};
        let maxCount = 0;
        let mode = '';
        arr.forEach((val) => {
          const v = (val || '').trim();
          if (!v) return;
          counts[v] = (counts[v] || 0) + 1;
          if (counts[v] > maxCount) {
            maxCount = counts[v];
            mode = v;
          }
        });
        return mode;
      };

      interface AuditLogEntry {
        operation: string;
        status: 'Corrected' | 'Generated' | 'Duplicate Commuter' | 'Commuter Not Found' | 'Unresolved';
        original: string;
        corrected: string;
        suggestedFix: string;
      }

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
        rowAuditLog: AuditLogEntry[];
        finalLandmark: string;
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
        let isDuplicateResolved = false;
        let isRefCodeRecovered = false;

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
                  isDuplicateResolved = true;
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
                    isDuplicateResolved = true;
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
                    isDuplicateResolved = true;
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
              isRefCodeRecovered = true;
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
                  isRefCodeRecovered = true;
                  isDuplicateResolved = true;
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
                    isRefCodeRecovered = true;
                    isDuplicateResolved = true;
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
                    isRefCodeRecovered = true;
                    isDuplicateResolved = true;
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
        const rowAuditLog: AuditLogEntry[] = [];

        // Check match results
        if (matchError) {
          errorsList.push(matchError);
          reasonsList.push(matchError);
          status = 'Error';
          errorsFound++;
          if (matchError === 'Duplicate Commuter') {
            duplicateNames++;
            rowAuditLog.push({
              operation: 'Requires Manual Review: Duplicate Commuter Conflict',
              status: 'Duplicate Commuter',
              original: name || '[Blank]',
              corrected: '',
              suggestedFix: 'Provide unique Father/Mother Mobile or Parent Email in Master and Route List'
            });
          } else {
            rowAuditLog.push({
              operation: 'Requires Manual Review: Commuter Not Found',
              status: 'Commuter Not Found',
              original: `Name: "${name || '[Blank]'}"; Ref Code: "${refCode || '[Blank]'}"`,
              corrected: '',
              suggestedFix: 'Verify commuter name and reference code match Master Data'
            });
          }
        }

        // Check for Route Number missing
        if (!routeNo) {
          errorsList.push('Route Missing');
          reasonsList.push('Route Missing');
          status = 'Error';
          errorsFound++;
          rowAuditLog.push({
            operation: 'Could Not Process: Route Number Missing',
            status: 'Unresolved',
            original: '',
            corrected: '',
            suggestedFix: 'Assign a valid Route Number in Route List'
          });
        }

        if (matchedMasterRow) {
          // Extract fields from Master Data row
          const mRefKeyMatches = ['Ref Code', 'Reference Code', 'Employee Code', 'Student Code', 'ID', 'Emp ID', 'Registration Number', 'Commuter Ref'];
          const mNameKeyMatches = ['Commuter Name', 'Name', 'Employee Name', 'Student Name', 'Passenger Name', 'Child Name'];
          const mLandmarkMatches = ['Landmark', 'Location', 'Stop', 'Pickup Stop', 'Pickup Location', 'Pickup Landmark'];

          masterRefCodeVal = getRowValue(matchedMasterRow, mRefKey, mRefKeyMatches).trim();
          masterNameVal = getRowValue(matchedMasterRow, mNameKey, mNameKeyMatches).trim();
          masterLandmarkVal = getRowValue(matchedMasterRow, mLandmarkKey, mLandmarkMatches).trim();

          // Recovered Ref Code from Master Data
          if (isRefCodeRecovered) {
            rowAuditLog.push({
              operation: 'Recovered Ref Code from Master Data',
              status: 'Corrected',
              original: refCode || '[Blank]',
              corrected: masterRefCodeVal,
              suggestedFix: ''
            });
          }

          // Resolved Duplicate Name
          if (isDuplicateResolved) {
            rowAuditLog.push({
              operation: 'Resolved Duplicate Commuter via Parent Contact',
              status: 'Corrected',
              original: name,
              corrected: `${masterNameVal} (${masterRefCodeVal})`,
              suggestedFix: ''
            });
          }

          // Check if parent contact details are completely missing
          const fMobile = getFatherMobile(matchedMasterRow);
          const mMobile = getMotherMobile(matchedMasterRow);
          const pEmail = getParentEmail(matchedMasterRow);
          if (!fMobile && !mMobile && !pEmail) {
            errorsList.push('Missing Parent Contact');
            reasonsList.push('Missing Parent Contact');
            rowAuditLog.push({
              operation: 'Missing Parent Contact',
              status: 'Corrected', // warning only, doesn't block route sheet inclusion
              original: '',
              corrected: '',
              suggestedFix: 'Provide Father/Mother Mobile or Parent Email in Master Data'
            });
          }
        } else {
          landmarkMissing++;
        }

        let finalLandmark = routeLandmark;
        if (matchedMasterRow) {
          const normMasterLandmark = normalizeLandmark(masterLandmarkVal);
          const simplifiedRoute = simplifyLandmark(normLandmark);
          const simplifiedMaster = simplifyLandmark(normMasterLandmark);

          if (simplifiedRoute === simplifiedMaster) {
            landmarkCorrect++;
            finalLandmark = masterLandmarkVal; // standardize to master casing
          } else {
            landmarkChanged++;
            finalLandmark = masterLandmarkVal; // overwrite with Master Data landmark!
            rowAuditLog.push({
              operation: 'Corrected Landmark Mismatch to Master Data',
              status: 'Corrected',
              original: routeLandmark || '[Blank]',
              corrected: masterLandmarkVal,
              suggestedFix: ''
            });
          }
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
          confidence,
          rowAuditLog,
          finalLandmark
        });
      });

      // Group Valid rows by Route for stop/time consistency and generation
      const routeGroups: Record<string, ProcessedRow[]> = {};
      processedRows.forEach((row) => {
        if (row.status === 'Valid' && row.routeNo) {
          const rNo = row.routeNo.toLowerCase().trim();
          if (!routeGroups[rNo]) {
            routeGroups[rNo] = [];
          }
          routeGroups[rNo].push(row);
        }
      });

      Object.keys(routeGroups).forEach((rNo) => {
        const group = routeGroups[rNo];
        if (group.length === 0) return;

        // Group rows in this route by finalLandmark
        const landmarkGroups: Record<string, ProcessedRow[]> = {};
        group.forEach((row) => {
          const key = simplifyLandmark(row.finalLandmark);
          if (!landmarkGroups[key]) {
            landmarkGroups[key] = [];
          }
          landmarkGroups[key].push(row);
        });

        // 1. Stop Number Generation or Standardization
        const stopNumbersMissing = group.every(row => !row.routeStopNo) || autoGenerateStops;

        if (stopNumbersMissing) {
          // Sort landmark groups by average original index to maintain supervisor sequence
          const sortedLandmarks = Object.keys(landmarkGroups).map(key => ({
            key,
            rows: landmarkGroups[key],
            avgIndex: landmarkGroups[key].reduce((sum, r) => sum + r.originalIndex, 0) / landmarkGroups[key].length
          }));
          sortedLandmarks.sort((a, b) => a.avgIndex - b.avgIndex);

          // Assign sequential stop numbers
          sortedLandmarks.forEach((landmarkItem, lIdx) => {
            const generatedStop = String(lIdx + 1);
            landmarkItem.rows.forEach((row) => {
              const origStop = row.routeStopNo;
              row.routeStopNo = generatedStop;
              row.rowAuditLog.push({
                operation: 'Generated Stop Number',
                status: 'Generated',
                original: origStop || '[Blank]',
                corrected: generatedStop,
                suggestedFix: ''
              });
            });
          });
        } else {
          // Standardize inconsistent stop numbers to the mode majority
          Object.keys(landmarkGroups).forEach((key) => {
            const landRows = landmarkGroups[key];
            const stopNums = landRows.map(r => r.routeStopNo).filter(Boolean);
            const modeStop = getMode(stopNums) || '1';

            landRows.forEach((row) => {
              const origStop = row.routeStopNo;
              if (origStop !== modeStop) {
                row.routeStopNo = modeStop;
                row.rowAuditLog.push({
                  operation: 'Standardized Inconsistent Stop Number to Mode',
                  status: 'Corrected',
                  original: origStop || '[Blank]',
                  corrected: modeStop,
                  suggestedFix: ''
                });
              }
            });
          });
        }

        // 2. Reporting Time Generation or Standardization
        const reportingTimesMissing = group.every(row => !row.routeReportingTime) || autoGenerateTimes;

        if (reportingTimesMissing) {
          // Sort landmark groups by stop number (numerical sort)
          const sortedLandmarks = Object.keys(landmarkGroups).map(key => ({
            key,
            rows: landmarkGroups[key],
            stopNum: parseInt(landmarkGroups[key][0].routeStopNo) || 999
          }));
          sortedLandmarks.sort((a, b) => a.stopNum - b.stopNum);

          const addMinutes = (timeStr: string, minutes: number): string => {
            const parts = timeStr.split(':');
            let hours = parseInt(parts[0]) || 8;
            let mins = parseInt(parts[1]) || 0;
            mins += minutes;
            hours += Math.floor(mins / 60);
            hours = hours % 24;
            mins = mins % 60;
            return `${String(hours).padStart(2, '0')}:${String(mins).padStart(2, '0')}`;
          };

          sortedLandmarks.forEach((landmarkItem, lIdx) => {
            const generatedTime = addMinutes(tripStartTime, lIdx * 2);
            landmarkItem.rows.forEach((row) => {
              const origTime = row.routeReportingTime;
              row.routeReportingTime = generatedTime;
              row.rowAuditLog.push({
                operation: 'Generated Reporting Time',
                status: 'Generated',
                original: origTime || '[Blank]',
                corrected: generatedTime,
                suggestedFix: ''
              });
            });
          });
        } else {
          // Standardize inconsistent reporting times to the mode majority
          Object.keys(landmarkGroups).forEach((key) => {
            const landRows = landmarkGroups[key];
            const times = landRows.map(r => r.routeReportingTime).filter(Boolean);
            const modeTime = getMode(times) || tripStartTime;

            landRows.forEach((row) => {
              const origTime = row.routeReportingTime;
              if (origTime !== modeTime) {
                row.routeReportingTime = modeTime;
                row.rowAuditLog.push({
                  operation: 'Standardized Inconsistent Reporting Time to Mode',
                  status: 'Corrected',
                  original: origTime || '[Blank]',
                  corrected: modeTime,
                  suggestedFix: ''
                });
              }
            });
          });
        }
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
          'Commuter Ref Code': row.masterRefCode,
          'Commuter Name': row.masterName,
          'Landmark Name': row.finalLandmark,
          'Route No': row.routeNo,
          'Stop No': row.routeStopNo,
          'Route Type': 'Bus',
          'Reporting Time': row.routeReportingTime
        });
      });

      // Build audit-centric Validation Report rows
      processedRows.forEach((row) => {
        row.rowAuditLog.forEach((logEntry) => {
          errorReportData.push({
            'Row Number': row.originalIndex + 2,
            'Commuter Name': row.name,
            'Ref Code': row.refCode,
            'Route No': row.routeNo,
            'Status': logEntry.status,
            'Operation': logEntry.operation,
            'Original Value': logEntry.original,
            'Corrected/Generated Value': logEntry.corrected,
            'Details/Suggested Fix': logEntry.suggestedFix
          });
        });
      });

      // Calculate Average Confidence
      let sumConfidence = 0;
      validRows.forEach(r => sumConfidence += r.confidence);
      const averageConfidence = validRows.length > 0 ? Math.round(sumConfidence / validRows.length) : 0;

      // Generate Route Sheet workbook
      const routeWorkbook = XLSX.utils.book_new();
      const routeWorksheet = XLSX.utils.json_to_sheet(routeSheetData, {
        header: [
          'Commuter Ref Code',
          'Commuter Name',
          'Landmark Name',
          'Route No',
          'Stop No',
          'Route Type',
          'Reporting Time'
        ]
      });
      XLSX.utils.book_append_sheet(routeWorkbook, routeWorksheet, 'Route Sheet');
      const routeOutput = XLSX.write(routeWorkbook, { bookType: 'xlsx', type: 'array' });
      const routeBuffer = new Uint8Array(routeOutput).buffer;

      // Generate Validation Report workbook
      const errorWorkbook = XLSX.utils.book_new();
      const errorWorksheet = XLSX.utils.json_to_sheet(
        errorReportData.length > 0
          ? errorReportData
          : [{ 'Message': 'No corrections, generations, or unresolved errors were recorded.' }],
        {
          header: [
            'Row Number',
            'Commuter Name',
            'Ref Code',
            'Route No',
            'Status',
            'Operation',
            'Original Value',
            'Corrected/Generated Value',
            'Details/Suggested Fix'
          ]
        }
      );
      XLSX.utils.book_append_sheet(errorWorkbook, errorWorksheet, 'Validation Report');
      const errorOutput = XLSX.write(errorWorkbook, { bookType: 'xlsx', type: 'array' });
      const errorBuffer = new Uint8Array(errorOutput).buffer;

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
