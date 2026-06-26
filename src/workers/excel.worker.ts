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
      const mPocNameKey = masterMappings?.POCName || '';
      const mPocMobileKey = masterMappings?.POCMobile || '';
      const mPocEmailKey = masterMappings?.POCEmail || '';

      // Mapped keys for Route File
      const rNameKey = routeMappings?.CommuterName || '';
      const rRefKey = routeMappings?.CommuterRefCode || '';
      const rLandmarkKey = routeMappings?.Landmark || '';
      const rAddressKey = routeMappings?.Address || '';
      const rPocNameKey = routeMappings?.POCName || '';
      const rPocMobileKey = routeMappings?.POCMobile || '';
      const rPocEmailKey = routeMappings?.POCEmail || '';
      const rRouteKey = routeMappings?.Route || '';
      const rRouteTypeKey = routeMappings?.RouteType || '';

      // Metrics counters
      const totalRecords = cachedRouteData.length;
      let matchedRecords = 0;
      let missingRefCodes = 0;
      let landmarkChanged = 0;
      let landmarkCorrect = 0;
      let landmarkMissing = 0;
      let duplicateNames = 0;
      let errorsFound = 0;
      let sumConfidence = 0;

      const comparisonLogs: string[] = [];
      const routeSheetData: Record<string, unknown>[] = [];
      const errorReportData: Record<string, unknown>[] = [];

      const simplifyLandmark = (val: string): string => {
        if (!val) return '';
        return val.toLowerCase().replace(/[^a-z0-9]/g, '');
      };

      // Build master ref code count map to check for duplicates in Master Data
      const masterRefCounts = new Map<string, number>();
      cachedMasterDataNormalized.forEach((mRow) => {
        const mRef = getRowValue(mRow, mRefKey, ['Ref Code', 'Reference Code', 'Employee Code', 'Student Code', 'ID', 'Emp ID', 'Registration Number', 'Commuter Ref']).trim().toLowerCase();
        if (mRef) {
          masterRefCounts.set(mRef, (masterRefCounts.get(mRef) || 0) + 1);
        }
      });
      const duplicateMasterRefs = new Set<string>();
      masterRefCounts.forEach((count, ref) => {
        if (count > 1) {
          duplicateMasterRefs.add(ref);
        }
      });

      interface ProcessedRow {
        refCode: string;
        name: string;
        landmark: string;
        routeNo: string;
        routeType: string;
        status: string;
        confidence: number;
        errorsList: string[];
        landmarkStatus: string;
        masterLandmarkVal: string;
        originalIndex: number;
        stopNo: number;
        reportingTime: string;
        address: string;
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
        const landmark = getRowValue(row, rLandmarkKey, ['Landmark', 'Location', 'Stop', 'Pickup Stop', 'Pickup Location', 'Pickup Landmark']).trim();
        const address = getRowValue(row, rAddressKey, ['Address', 'Commute Address', 'Street Address', 'Location Address']).trim();
        const pocName = getRowValue(row, rPocNameKey, ['Parent Name', 'Father Name', 'Guardian', 'Mother', 'POC Name', 'POC']).trim();
        const pocMobile = getRowValue(row, rPocMobileKey, ['Mobile', 'Contact Number', 'Phone', 'POC Mobile', 'POC Phone', 'Contact']).trim();
        const pocEmail = getRowValue(row, rPocEmailKey, ['Email', 'Email Address', 'POC Email', 'Parent Email']).trim();
        const routeNo = getRowValue(row, rRouteKey, ['Route Number', 'Vehicle Route', 'Route No', 'Route', 'Assigned Route', 'Route ID']).trim();
        const routeType = getRowValue(row, rRouteTypeKey, ['Pickup', 'Drop', 'Shift', 'Route Type', 'Type', 'Trip Type', 'Direction']).trim();

        // Normalizations for route items
        const normName = normalizeName(name);
        const normRef = refCode.trim().replace(/\s+/g, ' ');
        const normLandmark = normalizeLandmark(landmark);
        const normMobile = normalizeMobile(pocMobile);
        const normEmail = normalizeEmail(pocEmail);
        const normPocName = normalizeName(pocName);
        const normRouteNo = routeNo || 'Unassigned';
        const normRouteType = routeType || 'Pickup';

        let matchedMasterRow: Record<string, unknown> | null = null;
        let confidence = 0;
        let logMessage = '';
        let status = 'Valid';
        const errorsList: string[] = [];

        // Check if Ref Code was missing in the Uploaded Route List row
        if (!normRef) {
          missingRefCodes++;
        }

        // INTELLIGENT MATCHING ROUTINES:
        // Priority 1: Match by Ref Code (if it exists)
        if (normRef) {
          const matches = cachedMasterDataNormalized.filter((mRow) => {
            const mRef = getRowValue(mRow, mRefKey, ['Ref Code', 'Reference Code', 'Employee Code', 'Student Code', 'ID', 'Emp ID', 'Registration Number', 'Commuter Ref']).trim().toLowerCase();
            return mRef && mRef === normRef.toLowerCase();
          });

          if (matches.length === 1) {
            matchedMasterRow = matches[0];
            confidence = 100;
            logMessage = `Row ${index + 1}: Matched '${normName}' using Ref Code (${refCode}). Confidence: 100%`;
          } else if (matches.length > 1) {
            // Ref code matches multiple commuters in Master Data. Resolve using names.
            const nameMatches = matches.filter((mRow) => {
              const mName = getRowValue(mRow, mNameKey, ['Commuter Name', 'Name', 'Employee Name', 'Student Name', 'Passenger Name', 'Child Name']).trim().toLowerCase();
              return mName && mName === normName.toLowerCase();
            });

            if (nameMatches.length === 1) {
              matchedMasterRow = nameMatches[0];
              confidence = 95;
              logMessage = `Row ${index + 1}: Matched '${normName}' using Ref Code and Name duplicate resolution. Confidence: 95%`;
            } else {
              // Mark as duplicate conflict
              confidence = 0;
              status = 'Error';
              errorsList.push('Duplicate Ref Code Match Conflict');
              duplicateNames++;
              logMessage = `Row ${index + 1}: Ref Code Conflict. Multiple master entries match Ref Code (${refCode}).`;
            }
          } else {
            // matches.length === 0
            // If Ref Code exists but is unmatched, do NOT search by name.
            confidence = 0;
            status = 'Error';
            errorsList.push('Ref Code Not Found');
            logMessage = `Row ${index + 1}: Unmatched Ref Code. Ref Code (${refCode}) not found in Master Data.`;
          }
        } else {
          // Priority 2: Ref Code missing -> Search by Name
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
              // Multiple commuters found with the same name. Run fallbacks.
              // Priority 3: POC Name match
              const pocMatches = nameMatches.filter((mRow) => {
                const mPoc = getRowValue(mRow, mPocNameKey, ['Parent Name', 'Father Name', 'Guardian', 'Mother', 'POC Name', 'POC']).trim().toLowerCase();
                return mPoc && normPocName && mPoc === normPocName.toLowerCase();
              });

              if (pocMatches.length === 1) {
                matchedMasterRow = pocMatches[0];
                confidence = 75;
                logMessage = `Row ${index + 1}: Matched '${normName}' using Name & POC Name fallback. Confidence: 75%`;
              } else {
                // Priority 4: POC Mobile match
                const mobileMatches = nameMatches.filter((mRow) => {
                  const mMobile = getRowValue(mRow, mPocMobileKey, ['Mobile', 'Contact Number', 'Phone', 'POC Mobile', 'POC Phone', 'Contact']).trim();
                  const nMMobile = normalizeMobile(mMobile);
                  return nMMobile && normMobile && nMMobile === normMobile;
                });

                if (mobileMatches.length === 1) {
                  matchedMasterRow = mobileMatches[0];
                  confidence = 65;
                  logMessage = `Row ${index + 1}: Matched '${normName}' using Name & POC Mobile fallback. Confidence: 65%`;
                } else {
                  // Priority 5: POC Email match
                  const emailMatches = nameMatches.filter((mRow) => {
                    const mEmail = getRowValue(mRow, mPocEmailKey, ['Email', 'Email Address', 'POC Email', 'Parent Email']).trim().toLowerCase();
                    return mEmail && normEmail && mEmail === normEmail.toLowerCase();
                  });

                  if (emailMatches.length === 1) {
                    matchedMasterRow = emailMatches[0];
                    confidence = 55;
                    logMessage = `Row ${index + 1}: Matched '${normName}' using Name & POC Email fallback. Confidence: 55%`;
                  } else {
                    // Mark as duplicate conflict
                    confidence = 0;
                    status = 'Error';
                    errorsList.push('Duplicate Name Conflict');
                    duplicateNames++;
                    logMessage = `Row ${index + 1}: Duplicate Conflict. Multiple records match Name '${normName}' without distinct POC details.`;
                  }
                }
              }
            } else {
              // nameMatches.length === 0
              status = 'Error';
              errorsList.push('Commuter Unmatched');
              logMessage = `Row ${index + 1}: Unmatched Commuter. Name '${normName}' not found in Master Data.`;
            }
          } else {
            status = 'Error';
            errorsList.push('Commuter Unmatched (Missing Name and Ref Code)');
            logMessage = `Row ${index + 1}: Unmatched Commuter. Both Name and Ref Code are missing.`;
          }
        }

        // Check for anomalies: Blank Route, Blank Ref Code, Blank Landmark
        if (!routeNo.trim()) {
          errorsList.push('Blank Route');
          if (status !== 'Error') status = 'Review';
        }
        if (!normRef) {
          errorsList.push('Blank Ref Code');
          if (status !== 'Error') status = 'Review';
        }
        if (!normLandmark) {
          errorsList.push('Blank Landmark');
          if (status !== 'Error') status = 'Review';
        }

        // Perform Landmark Validation
        let landmarkStatus = 'Missing';
        let masterLandmarkVal = 'N/A';

        if (matchedMasterRow) {
          matchedRecords++;
          sumConfidence += confidence;

          const masterLandmark = getRowValue(matchedMasterRow, mLandmarkKey, ['Landmark', 'Location', 'Stop', 'Pickup Stop', 'Pickup Location', 'Pickup Landmark']).trim();
          const normMasterLandmark = normalizeLandmark(masterLandmark);
          masterLandmarkVal = normMasterLandmark;

          const mRef = getRowValue(matchedMasterRow, mRefKey, ['Ref Code', 'Reference Code', 'Employee Code', 'Student Code', 'ID', 'Emp ID', 'Registration Number', 'Commuter Ref']).trim();

          // Check for Duplicate Ref Code in Master
          if (mRef && duplicateMasterRefs.has(mRef.toLowerCase())) {
            errorsList.push('Duplicate Ref Code');
            if (status !== 'Error') status = 'Review';
          }

          // Check if matched Master Ref Code is blank
          if (!mRef) {
            errorsList.push('Blank Ref Code in Master');
            if (status !== 'Error') status = 'Review';
          }

          // Check if matched Master Landmark is blank
          if (!normMasterLandmark) {
            errorsList.push('Blank Landmark in Master');
            if (status !== 'Error') status = 'Review';
          }

          // Landmark Verification comparisons
          if (!normLandmark || !normMasterLandmark) {
            landmarkStatus = 'Missing';
            landmarkMissing++;
            if (!errorsList.includes('Blank Landmark') && !errorsList.includes('Landmark Missing') && !errorsList.includes('Blank Landmark in Master')) {
              errorsList.push('Landmark Missing');
            }
            if (status !== 'Error') status = 'Review';
          } else {
            const simplifiedRoute = simplifyLandmark(normLandmark);
            const simplifiedMaster = simplifyLandmark(normMasterLandmark);

            if (simplifiedRoute === simplifiedMaster) {
              landmarkStatus = 'Correct';
              landmarkCorrect++;
            } else {
              landmarkStatus = 'Changed';
              landmarkChanged++;
              errorsList.push(`Landmark mismatch (Master: ${normMasterLandmark})`);
              if (status !== 'Error') status = 'Review';
            }
          }
        } else {
          // If commuter is unmatched or has conflict
          landmarkStatus = 'Missing';
          landmarkMissing++;
        }

        // Increment errors count if status is Error
        if (status === 'Error') {
          errorsFound++;
        }

        // Compile row logs
        comparisonLogs.push(logMessage);

        processedRows.push({
          refCode: matchedMasterRow 
            ? (getRowValue(matchedMasterRow, mRefKey, ['Ref Code', 'Reference Code', 'Employee Code', 'Student Code', 'ID', 'Emp ID', 'Registration Number', 'Commuter Ref']).trim() || normRef)
            : normRef,
          name: normName || 'Unknown Commuter',
          landmark: normLandmark || 'Unassigned',
          routeNo: normRouteNo,
          routeType: normRouteType,
          status,
          confidence,
          errorsList,
          landmarkStatus,
          masterLandmarkVal,
          originalIndex: index,
          stopNo: 0,
          reportingTime: '',
          address: address || 'N/A',
        });
      });

      // Group by Route No
      const routeGroups: Record<string, ProcessedRow[]> = {};
      processedRows.forEach((row) => {
        if (!routeGroups[row.routeNo]) {
          routeGroups[row.routeNo] = [];
        }
        routeGroups[row.routeNo].push(row);
      });

      // Format time helper: starts at 08:00 AM, adding 5 mins per stop
      const formatTime = (minutesFromBase: number): string => {
        const baseHour = 8;
        const totalMinutes = baseHour * 60 + minutesFromBase;
        const hours = Math.floor(totalMinutes / 60) % 24;
        const mins = totalMinutes % 60;
        const hh = String(hours).padStart(2, '0');
        const mm = String(mins).padStart(2, '0');
        return `${hh}:${mm}`;
      };

      // Assign Stop Numbers and Reporting Times per Route maintaining appearance sequence of landmarks
      Object.keys(routeGroups).forEach((routeNo) => {
        const rowsInRoute = routeGroups[routeNo];
        rowsInRoute.sort((a, b) => a.originalIndex - b.originalIndex);

        // Identify unique landmarks in order of appearance
        const landmarkSequence: string[] = [];
        rowsInRoute.forEach((row) => {
          const lKey = row.landmark.toLowerCase().trim();
          if (!landmarkSequence.includes(lKey)) {
            landmarkSequence.push(lKey);
          }
        });

        // Map from landmark to stop details
        const landmarkStopMap = new Map<string, { stopNo: number; time: string }>();
        landmarkSequence.forEach((lKey, i) => {
          landmarkStopMap.set(lKey, {
            stopNo: i + 1,
            time: formatTime(i * 5),
          });
        });

        // Assign stop values
        rowsInRoute.forEach((row) => {
          const lKey = row.landmark.toLowerCase().trim();
          const stopDetails = landmarkStopMap.get(lKey);
          if (stopDetails) {
            row.stopNo = stopDetails.stopNo;
            row.reportingTime = stopDetails.time;
          }
        });
      });

      // Flatten and sort: Route Number, Stop Number, Commuter Name
      const flatProcessedRows: ProcessedRow[] = [];
      Object.keys(routeGroups).forEach((routeNo) => {
        flatProcessedRows.push(...routeGroups[routeNo]);
      });

      flatProcessedRows.sort((a, b) => {
        const routeCompare = a.routeNo.localeCompare(b.routeNo, undefined, { numeric: true, sensitivity: 'base' });
        if (routeCompare !== 0) return routeCompare;

        if (a.stopNo !== b.stopNo) {
          return a.stopNo - b.stopNo;
        }

        return a.name.localeCompare(b.name, undefined, { sensitivity: 'base' });
      });

      // Populate sheet data arrays
      flatProcessedRows.forEach((row) => {
        routeSheetData.push({
          'CommuterRefCode': row.refCode || 'N/A',
          'Commuter Name': row.name || 'Unknown Commuter',
          'Landmark Name': row.landmark || 'N/A',
          'Route No': row.routeNo || 'N/A',
          'Stop No': row.stopNo,
          'Route Type': row.routeType || 'N/A',
          'Reporting Time': row.reportingTime
        });

        if (row.status === 'Error' || row.status === 'Review') {
          const corrections: string[] = [];
          row.errorsList.forEach((err) => {
            if (err.includes('Landmark mismatch')) {
              corrections.push('Update Route Stop to match Master landmark');
            } else if (err === 'Landmark Missing' || err === 'Blank Landmark' || err === 'Blank Landmark in Master') {
              corrections.push('Provide missing landmark value');
            } else if (err === 'Blank Route') {
              corrections.push('Assign a valid Route Number');
            } else if (err === 'Blank Ref Code' || err === 'Blank Ref Code in Master') {
              corrections.push('Assign a unique Reference Code');
            } else if (err === 'Duplicate Ref Code') {
              corrections.push('Resolve duplicated Reference Code in Master');
            } else if (err === 'Duplicate Name Conflict') {
              corrections.push('Disambiguate commuter using unique POC details');
            } else if (err === 'Duplicate Ref Code Match Conflict') {
              corrections.push('Resolve duplicated Reference Code conflict');
            } else if (err === 'Commuter Unmatched') {
              corrections.push('Verify commuter exists in Master database');
            }
          });
          const recommendedCorrection = corrections.join('; ') || 'Verify record data details';

          errorReportData.push({
            'Row Index': row.originalIndex + 2,
            'Commuter Name': row.name,
            'Ref Code': row.refCode,
            'Landmark': row.landmark,
            'Status': row.status,
            'Anomalies Detected': row.errorsList.join('; '),
            'Recommended Correction': recommendedCorrection,
          });
        }
      });

      // Calculate Average Confidence
      const matchedCount = matchedRecords || 1;
      const averageConfidence = Math.round(sumConfidence / matchedCount);

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
              matchedRecords,
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
