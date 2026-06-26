import { useState, useEffect } from 'react';
import {
  Sun,
  Moon,
  Play,
  Download,
  RefreshCw,
  Database,
  Grid,
  FileSpreadsheet,
  AlertCircle,
  Terminal,
  Check,
  Settings,
} from 'lucide-react';

import { useTheme } from './context/ThemeContext';
import { useToast } from './context/ToastContext';
import { UploadCard } from './components/UploadCard';
import { ExcelHelper } from './utils/excelHelper';
import { logger } from './utils/logger';

interface ValidationMetrics {
  totalRecords: number;
  matchedRecords: number;
  missingRefCodes: number;
  landmarkChanged: number;
  landmarkCorrect: number;
  landmarkMissing: number;
  duplicateNames: number;
  errorsFound: number;
  averageConfidence: number;
}

interface MasterAudit {
  totalRecords: number;
  duplicateRefCodes: number;
  blankRefCodes: number;
  blankNames: number;
  blankLandmarks: number;
  errorsCount: number;
}

const initialMasterAudit: MasterAudit = {
  totalRecords: 0,
  duplicateRefCodes: 0,
  blankRefCodes: 0,
  blankNames: 0,
  blankLandmarks: 0,
  errorsCount: 0,
};

const initialMetrics: ValidationMetrics = {
  totalRecords: 0,
  matchedRecords: 0,
  missingRefCodes: 0,
  landmarkChanged: 0,
  landmarkCorrect: 0,
  landmarkMissing: 0,
  duplicateNames: 0,
  errorsFound: 0,
  averageConfidence: 0,
};


function App() {
  const { theme, toggleTheme } = useTheme();
  const { addToast } = useToast();

  // File Upload states
  const [masterFileName, setMasterFileName] = useState<string | null>(null);
  const [masterRows, setMasterRows] = useState(0);
  const [masterCols, setMasterCols] = useState(0);
  const [masterMappings, setMasterMappings] = useState<Record<string, string>>({});
  const [masterAudit, setMasterAudit] = useState<MasterAudit>(initialMasterAudit);

  const [routeFileName, setRouteFileName] = useState<string | null>(null);
  const [routeRows, setRouteRows] = useState(0);
  const [routeCols, setRouteCols] = useState(0);
  const [routeMappings, setRouteMappings] = useState<Record<string, string>>({});

  // Processing states
  const [isProcessing, setIsProcessing] = useState(false);
  const [processingStatus, setProcessingStatus] = useState<string>('Idle');
  const [matchingProgress, setMatchingProgress] = useState(0);
  const [elapsedTime, setElapsedTime] = useState(0);
  const [showSuccessBadge, setShowSuccessBadge] = useState(false);
  
  // Output and Metrics
  const [metrics, setMetrics] = useState<ValidationMetrics>(initialMetrics);
  const [routeSheetBuffer, setRouteSheetBuffer] = useState<ArrayBuffer | null>(null);
  const [errorReportBuffer, setErrorReportBuffer] = useState<ArrayBuffer | null>(null);
  const [logs, setLogs] = useState<string[]>([]);

  // Auto-generation configurations
  const [autoGenerateStops, setAutoGenerateStops] = useState<boolean>(true);
  const [autoGenerateTimes, setAutoGenerateTimes] = useState<boolean>(true);
  const [tripStartTime, setTripStartTime] = useState<string>('08:00');
  const [tripEndTime, setTripEndTime] = useState<string>('09:00');

  // Cache to track the last run configuration and prevent duplicate processing
  const [lastRunConfig, setLastRunConfig] = useState<{
    masterFileName: string | null;
    masterRows: number;
    masterCols: number;
    masterMappings: string;
    routeFileName: string | null;
    routeRows: number;
    routeCols: number;
    routeMappings: string;
  } | null>(null);

  const isBothFilesUploaded = masterFileName !== null && routeFileName !== null;



  // Master file loaded callback
  const handleMasterLoaded = async (
    _data: Record<string, unknown>[],
    name: string,
    rows: number,
    cols: number,
    mappings: Record<string, string>
  ) => {
    setMasterFileName(name);
    setMasterRows(rows);
    setMasterCols(cols);
    setMasterMappings(mappings);
    clearPreviousRuns();
    setProcessingStatus('Processing Master Data...');

    // FUTURE API INTEGRATION PLACEHOLDER:
    // To switch from client-side worker sorting/auditing to a backend REST API, uncomment the following block:
    /*
    try {
      const response = await fetch('/api/master/process', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ mappings }),
      });
      if (!response.ok) throw new Error('API server returned error status');
      const result = await response.json();
      setMasterAudit(result.auditSummary);
      setProcessingStatus(routeFileName ? 'Files Ready' : 'Idle');
      addToast('Master Data parsed and normalized successfully via REST API.', 'success');
      return;
    } catch (apiErr) {
      logger.error('Backend API error', apiErr);
    }
    */

    try {
      const result = await ExcelHelper.processData(mappings);
      setMasterAudit(result.auditSummary);
      setProcessingStatus(routeFileName ? 'Files Ready' : 'Idle');
      addToast('Master Data parsed and normalized successfully.', 'success');
      logger.info('Master Data audit counts generated: ', result.auditSummary);
    } catch (e) {
      logger.error('Failed to process master data on upload', e);
      addToast('Master file parsed but cleanup processing failed.', 'warning');
      setProcessingStatus(routeFileName ? 'Files Ready' : 'Idle');
    }
  };

  // Master file cleared callback
  const handleMasterCleared = () => {
    setMasterFileName(null);
    setMasterRows(0);
    setMasterCols(0);
    setMasterMappings({});
    setMasterAudit(initialMasterAudit);
    clearPreviousRuns();
    setProcessingStatus(routeFileName ? 'Master Data Missing' : 'Idle');
  };

  // Handle manual column mapping modifications for Master Data
  const handleMasterMappingChange = async (newMappings: Record<string, string>) => {
    setMasterMappings(newMappings);
    if (masterFileName) {
      try {
        setProcessingStatus('Processing Master Data...');
        // FUTURE API INTEGRATION PLACEHOLDER:
        // To swap to backend API mapping change processing:
        /*
        const response = await fetch('/api/master/process', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ mappings: newMappings }),
        });
        const result = await response.json();
        setMasterAudit(result.auditSummary);
        setProcessingStatus(routeFileName ? 'Files Ready' : 'Idle');
        addToast('Master Data re-normalized via REST API.', 'info');
        return;
        */
        const result = await ExcelHelper.processData(newMappings);
        setMasterAudit(result.auditSummary);
        setProcessingStatus(routeFileName ? 'Files Ready' : 'Idle');
        addToast('Master Data re-normalized with new column mappings.', 'info');
      } catch (e) {
        logger.error('Failed to re-process master data mapping change', e);
        addToast('Failed to apply new column mappings to Master data.', 'error');
        setProcessingStatus(routeFileName ? 'Files Ready' : 'Idle');
      }
    }
  };

  // Route file loaded callback
  const handleRouteLoaded = (
    _data: Record<string, unknown>[],
    name: string,
    rows: number,
    cols: number,
    mappings: Record<string, string>
  ) => {
    setRouteFileName(name);
    setRouteRows(rows);
    setRouteCols(cols);
    setRouteMappings(mappings);
    clearPreviousRuns();
    setProcessingStatus(masterFileName ? 'Files Ready' : 'Idle');
    addToast('Route List parsed successfully.', 'success');
  };

  // Route file cleared callback
  const handleRouteCleared = () => {
    setRouteFileName(null);
    setRouteRows(0);
    setRouteCols(0);
    setRouteMappings({});
    clearPreviousRuns();
    setProcessingStatus(masterFileName ? 'Route List Missing' : 'Idle');
  };

  // Reset generated metrics and download states
  const clearPreviousRuns = () => {
    setMetrics(initialMetrics);
    setRouteSheetBuffer(null);
    setErrorReportBuffer(null);
    setLogs([]);
    setMatchingProgress(0);
    setElapsedTime(0);
    setShowSuccessBadge(false);
  };

  // Keyboard Shortcuts Registration
  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      // Ctrl + Enter -> Generate Route Sheet
      if (e.ctrlKey && e.key === 'Enter') {
        e.preventDefault();
        const generateBtn = document.getElementById('generate-sheet-btn') as HTMLButtonElement | null;
        if (generateBtn && !generateBtn.disabled) {
          generateBtn.click();
        }
      }
      // Alt + M -> Browse/Replace Master Data
      if (e.altKey && e.key.toLowerCase() === 'm') {
        e.preventDefault();
        const masterBrowseBtn = document.getElementById('master-upload-browse-btn') as HTMLButtonElement | null;
        if (masterBrowseBtn) {
          masterBrowseBtn.click();
        } else {
          const replaceBtn = document.getElementById('master-upload-replace-btn') as HTMLButtonElement | null;
          if (replaceBtn) replaceBtn.click();
        }
      }
      // Alt + R -> Browse/Replace Route List
      if (e.altKey && e.key.toLowerCase() === 'r') {
        e.preventDefault();
        const routeBrowseBtn = document.getElementById('route-upload-browse-btn') as HTMLButtonElement | null;
        if (routeBrowseBtn) {
          routeBrowseBtn.click();
        } else {
          const replaceBtn = document.getElementById('route-upload-replace-btn') as HTMLButtonElement | null;
          if (replaceBtn) replaceBtn.click();
        }
      }
      // Esc -> Clear results, modal, or files
      if (e.key === 'Escape') {
        if (showSuccessBadge) {
          setShowSuccessBadge(false);
        } else if (routeSheetBuffer || errorReportBuffer || logs.length > 0) {
          clearPreviousRuns();
          addToast('Cleared validation results.', 'info');
        } else {
          const masterRemoveBtn = document.getElementById('master-upload-remove-btn') as HTMLButtonElement | null;
          const routeRemoveBtn = document.getElementById('route-upload-remove-btn') as HTMLButtonElement | null;
          if (masterRemoveBtn || routeRemoveBtn) {
            if (masterRemoveBtn) masterRemoveBtn.click();
            if (routeRemoveBtn) routeRemoveBtn.click();
            addToast('Removed uploaded spreadsheet files.', 'info');
          }
        }
      }
    };

    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [
    masterFileName,
    routeFileName,
    isProcessing,
    routeSheetBuffer,
    errorReportBuffer,
    logs,
    masterMappings,
    routeMappings,
    masterRows,
    masterCols,
    routeRows,
    routeCols,
    showSuccessBadge,
    addToast
  ]);

  // Trigger worker comparison and route sheet generation
  const handleGenerateRouteSheet = async () => {
    if (!isBothFilesUploaded) return;

    // Prevent duplicate processing triggers if parameters are identical to the previous run
    const currentConfig = {
      masterFileName,
      masterRows,
      masterCols,
      masterMappings: JSON.stringify(masterMappings),
      routeFileName,
      routeRows,
      routeCols,
      routeMappings: JSON.stringify(routeMappings)
    };

    if (
      lastRunConfig &&
      JSON.stringify(lastRunConfig) === JSON.stringify(currentConfig) &&
      routeSheetBuffer
    ) {
      addToast('Using cached sheets. Re-downloading...', 'info');
      triggerDownload(routeSheetBuffer, 'Standardized_Route_Sheet.xlsx');
      if (errorReportBuffer) {
        setTimeout(() => {
          triggerDownload(errorReportBuffer, 'Route_Validation_Report.xlsx');
        }, 300);
      }
      setShowSuccessBadge(true);
      return;
    }

    setIsProcessing(true);
    setMatchingProgress(0);
    setElapsedTime(0);
    setShowSuccessBadge(false);
    setProcessingStatus('Comparing commuter reference codes...');
    addToast('Comparing files and matching commuter details in background...', 'info', 2000);
    logger.info(`Starting comparison routine. Master: ${masterRows} rows, Route: ${routeRows} rows.`);

    const startTime = performance.now();
    const timerInterval = setInterval(() => {
      setElapsedTime(Math.round(performance.now() - startTime));
    }, 100);

    // FUTURE API INTEGRATION PLACEHOLDER:
    // Replace this local client Web Worker promise chain with a POST request to your backend REST API.
    // E.g.:
    /*
    try {
      const response = await fetch('/api/routes/compare', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ masterMappings, routeMappings }),
      });
      if (!response.ok) throw new Error('Comparative API call failed.');
      const result = await response.json();
      
      clearInterval(timerInterval);
      const endTime = performance.now();
      setElapsedTime(Math.round(endTime - startTime));
      
      setMetrics(result.metrics);
      // Decode Base64 binary downloads if returned by backend REST API:
      // const routeBuffer = Uint8Array.from(atob(result.routeSheetBase64), c => c.charCodeAt(0)).buffer;
      // const errorBuffer = Uint8Array.from(atob(result.errorReportBase64), c => c.charCodeAt(0)).buffer;
      
      setRouteSheetBuffer(routeBuffer);
      setErrorReportBuffer(errorBuffer);
      setLogs(result.comparisonLogs || []);
      setProcessingStatus('Success');
      setShowSuccessBadge(true);
      setLastRunConfig(currentConfig);
      addToast('Route Sheet generated successfully via REST API!', 'success');
      triggerDownload(routeBuffer, 'Standardized_Route_Sheet.xlsx');
      return;
    } catch (apiErr) {
      clearInterval(timerInterval);
      logger.error('Comparative API validation failed', apiErr);
    }
    */

    try {
      // Execute comparative validation in worker thread
      const result = await ExcelHelper.compareFiles(
        masterMappings,
        routeMappings,
        {
          autoGenerateStops,
          autoGenerateTimes,
          tripStartTime,
          tripEndTime,
        },
        (progress) => {
          setMatchingProgress(progress);
        }
      );

      clearInterval(timerInterval);
      const endTime = performance.now();
      setElapsedTime(Math.round(endTime - startTime));

      setMetrics(result.metrics);
      setRouteSheetBuffer(result.routeSheetBuffer);
      setErrorReportBuffer(result.errorReportBuffer);
      setLogs(result.comparisonLogs || []);
      setProcessingStatus('Success');
      setShowSuccessBadge(true);
      setLastRunConfig(currentConfig);
      addToast('Route Sheet generated successfully!', 'success');
      logger.info('Route comparison metrics compiled and returned from background.');

      // Automatically trigger download of standardized route sheet
      triggerDownload(result.routeSheetBuffer, 'Standardized_Route_Sheet.xlsx');

      // If there are errors or reviews, automatically trigger download of validation report too
      const hasAnomalies = result.metrics.errorsFound > 0 || 
                           result.metrics.landmarkChanged > 0 || 
                           result.metrics.landmarkMissing > 0 ||
                           result.metrics.missingRefCodes > 0;
      if (hasAnomalies) {
        setTimeout(() => {
          triggerDownload(result.errorReportBuffer, 'Route_Validation_Report.xlsx');
        }, 300);
      }

    } catch (err: any) {
      clearInterval(timerInterval);
      logger.error('Failed file processing comparison', err);
      addToast(err?.message || 'Error validating route list details.', 'error');
      setProcessingStatus('Failed');
      clearPreviousRuns();
    } finally {
      setIsProcessing(false);
    }
  };

  // Trigger Excel file download from state ArrayBuffer
  const triggerDownload = (buffer: ArrayBuffer | null, filename: string) => {
    if (!buffer) return;

    try {
      const blob = new Blob([buffer], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      });
      const url = window.URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.setAttribute('download', filename);
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      window.URL.revokeObjectURL(url);
      addToast(`Downloaded ${filename} successfully!`, 'success');
    } catch (e) {
      logger.error(`Download fail for ${filename}`, e);
      addToast('Failed to trigger file download.', 'error');
    }
  };

  return (
    <div style={appWrapperStyle}>
      {/* Top Navbar Header */}
      <header style={headerStyle} className="glass-panel">
        <div style={headerLeftStyle}>
          <div style={logoWrapperStyle}>
            <svg viewBox="0 0 32 32" fill="none" xmlns="http://www.w3.org/2000/svg" style={{ width: '22px', height: '22px' }}>
              <defs>
                <linearGradient id="headerRLogoGradient" x1="0%" y1="0%" x2="100%" y2="100%">
                  <stop offset="0%" stopColor="#2563EB"/>
                  <stop offset="100%" stopColor="#3B82F6"/>
                </linearGradient>
              </defs>
              <circle cx="8.75" cy="24" r="3" fill="url(#headerRLogoGradient)" stroke="#FFFFFF" strokeWidth="1"/>
              <path d="M8.75 8H17C20.866 8 24 11.134 24 15C24 18.866 20.866 22 17 22H8.75V18.5H17C18.933 18.5 20.5 16.933 20.5 15C20.5 13.067 18.933 11.5 17 11.5H8.75V8Z" fill="url(#headerRLogoGradient)"/>
              <rect x="7" y="8" width="3.5" height="16" rx="1.75" fill="url(#headerRLogoGradient)"/>
              <path d="M15.5 20.5 L20.5 27" stroke="url(#headerRLogoGradient)" strokeWidth="3.5" strokeLinecap="round"/>
              <path d="M21.5 21C19.5 21 18 22.5 18 24.5C18 27.5 21.5 31 21.5 31C21.5 31 25 27.5 25 24.5C25 22.5 23.5 21 21.5 21Z" fill="url(#headerRLogoGradient)"/>
              <circle cx="21.5" cy="24.5" r="1.5" fill="#FFFFFF"/>
            </svg>
          </div>
          <div>
            <div style={logoTextStyle}>RouteSync <span style={{ color: 'var(--primary)' }}>AI</span></div>
            <span style={logoSubtitleStyle}>Validate. Match. Generate.</span>
          </div>
        </div>

        <button id="theme-toggle-btn" onClick={toggleTheme} style={themeToggleButtonStyle} aria-label="Toggle visual theme">
          {theme === 'dark' ? <Sun size={18} /> : <Moon size={18} />}
        </button>
      </header>

      {/* Main Core Dashboard Content */}
      <main className="container animate-fade-in" style={mainStyle}>
        {/* Banner Area */}
        <section style={bannerSectionStyle}>
          <h1 style={bannerHeadingStyle}>RouteSync AI</h1>
          <p style={taglineStyle}>
            <span style={{ color: '#3b82f6' }}>Validate.</span>{' '}
            <span style={{ color: '#10b981' }}>Match.</span>{' '}
            <span style={{ color: '#f97316' }}>Generate.</span>
          </p>
          <p style={bannerSubheadingStyle}>
            Upload your Master Data and Route List to intelligently validate commuter records, detect reference code and landmark mismatches, resolve duplicate commuters, and generate a standardized Route Sheet in a single click.
          </p>
        </section>

        {/* Upload Cards Grid */}
        <section style={uploadGridStyle}>
          <UploadCard
            idPrefix="master-upload"
            title="Upload Master Data"
            onFileLoaded={handleMasterLoaded}
            onFileCleared={handleMasterCleared}
            onMappingChange={handleMasterMappingChange}
            isLoading={false}
            progress={0}
          />
          <UploadCard
            idPrefix="route-upload"
            title="Upload Route List"
            onFileLoaded={handleRouteLoaded}
            onFileCleared={handleRouteCleared}
            onMappingChange={setRouteMappings}
            isLoading={false}
            progress={0}
          />

        </section>

        {/* Trip & Generation Settings panel */}
        {isBothFilesUploaded && (
          <section className="glass-panel animate-fade-in" style={settingsPanelStyle}>
            <h3 style={settingsTitleStyle}>
              <Settings size={18} color="var(--primary)" style={{ marginRight: '8px' }} />
              Trip & Generation Settings
            </h3>
            
            <div style={settingsGridContainerStyle}>
              <div style={settingsGroupStyle}>
                <span style={settingsLabelStyle}>Stop Number Options</span>
                <div style={checkboxWrapperStyle}>
                  <input
                    type="checkbox"
                    id="auto-generate-stops-check"
                    checked={autoGenerateStops}
                    onChange={(e) => setAutoGenerateStops(e.target.checked)}
                    style={checkboxStyle}
                  />
                  <label htmlFor="auto-generate-stops-check" style={checkboxLabelStyle}>
                    Auto-generate Stop Numbers sequentially if missing
                  </label>
                </div>
              </div>

              <div style={settingsGroupStyle}>
                <span style={settingsLabelStyle}>Reporting Time Options</span>
                <div style={checkboxWrapperStyle}>
                  <input
                    type="checkbox"
                    id="auto-generate-times-check"
                    checked={autoGenerateTimes}
                    onChange={(e) => setAutoGenerateTimes(e.target.checked)}
                    style={checkboxStyle}
                  />
                  <label htmlFor="auto-generate-times-check" style={checkboxLabelStyle}>
                    Auto-generate Reporting Times if missing
                  </label>
                </div>
              </div>

              <div style={settingsGroupStyle}>
                <span style={settingsLabelStyle}>Trip Window</span>
                <div style={timeInputsContainerStyle}>
                  <div style={timeInputWrapperStyle}>
                    <span style={timeInputLabelStyle}>Start:</span>
                    <input
                      type="time"
                      value={tripStartTime}
                      onChange={(e) => setTripStartTime(e.target.value)}
                      style={timeInputStyle}
                      className="glass-panel"
                    />
                  </div>
                  <div style={timeInputWrapperStyle}>
                    <span style={timeInputLabelStyle}>End:</span>
                    <input
                      type="time"
                      value={tripEndTime}
                      onChange={(e) => setTripEndTime(e.target.value)}
                      style={timeInputStyle}
                      className="glass-panel"
                    />
                  </div>
                </div>
              </div>
            </div>
          </section>
        )}

        {/* Trigger Button Section */}
        <section style={triggerButtonStyleContainer}>
          <button
            id="generate-sheet-btn"
            onClick={handleGenerateRouteSheet}
            disabled={!isBothFilesUploaded || isProcessing}
            style={
              !isBothFilesUploaded || isProcessing ? disabledTriggerButtonStyle : triggerButtonStyle
            }
          >
            {isProcessing ? (
              <>
                <RefreshCw className="spin-animation" size={18} style={{ marginRight: '10px' }} />
                Validating Commuters...
              </>
            ) : (
              <>
                <Play size={18} style={{ marginRight: '10px' }} />
                Generate Route Sheet <kbd style={kbdStyle}>Ctrl + Enter</kbd>
              </>
            )}
          </button>

          {isProcessing && (
            <div style={progressContainerStyle} className="glass-panel animate-slide-up">
              <div style={progressHeaderStyle}>
                <span style={progressStatusTextStyle}>
                  {processingStatus} ({matchingProgress}%)
                </span>
                <span style={progressTimerStyle}>
                  Elapsed: {(elapsedTime / 1000).toFixed(1)}s
                </span>
              </div>
              <div style={progressBarTrackStyle}>
                <div
                  style={{
                    ...progressBarFillStyle,
                    width: `${matchingProgress}%`,
                  }}
                />
              </div>
            </div>
          )}
        </section>

        {/* Summary Details Section */}
        <section style={sectionBlockStyle}>
          <h3 style={sectionTitleStyle}>
            <Database size={18} color="var(--primary)" style={{ marginRight: '8px' }} />
            Processing Details
          </h3>
          <div style={detailsGridStyle}>
            {/* Card 1: Master File Status */}
            <div className="glass-panel" style={detailCardStyle}>
              <span style={detailLabelStyle}>Master File</span>
              <span
                style={{
                  ...detailValueStyle,
                  color: masterFileName ? 'var(--success)' : 'var(--text-muted)',
                }}
              >
                {masterFileName ? 'Uploaded' : 'Not Uploaded'}
              </span>
              <span style={detailSubtextStyle}>
                {masterFileName ? masterFileName : 'Awaiting spreadsheet file'}
              </span>
            </div>

            {/* Card 2: Route File Status */}
            <div className="glass-panel" style={detailCardStyle}>
              <span style={detailLabelStyle}>Route File</span>
              <span
                style={{
                  ...detailValueStyle,
                  color: routeFileName ? 'var(--success)' : 'var(--text-muted)',
                }}
              >
                {routeFileName ? 'Uploaded' : 'Not Uploaded'}
              </span>
              <span style={detailSubtextStyle}>
                {routeFileName ? routeFileName : 'Awaiting spreadsheet file'}
              </span>
            </div>

            {/* Card 3: Rows Detected */}
            <div className="glass-panel" style={detailCardStyle}>
              <span style={detailLabelStyle}>Rows Detected</span>
              <span style={detailValueStyle}>
                {masterRows || routeRows ? `${masterRows} (M) / ${routeRows} (R)` : '0'}
              </span>
              <span style={detailSubtextStyle}>Total records parsed across uploads</span>
            </div>

            {/* Card 4: Columns Detected */}
            <div className="glass-panel" style={detailCardStyle}>
              <span style={detailLabelStyle}>Columns Detected</span>
              <span style={detailValueStyle}>
                {masterCols || routeCols ? `${masterCols} (M) / ${routeCols} (R)` : '0'}
              </span>
              <span style={detailSubtextStyle}>Number of headers detected in tables</span>
            </div>

            {/* Card 5: Processing Status */}
            <div className="glass-panel" style={detailCardStyle}>
              <span style={detailLabelStyle}>Processing Status</span>
              <span
                style={{
                  ...detailValueStyle,
                  color:
                    processingStatus === 'Success'
                      ? 'var(--success)'
                      : processingStatus === 'Failed'
                      ? 'var(--error)'
                      : 'var(--primary)',
                }}
              >
                {processingStatus}
              </span>
              <span style={detailSubtextStyle}>Current engine validator state</span>
            </div>
          </div>
        </section>

        {/* Master Data Quality Summary */}
        <section style={sectionBlockStyle}>
          <h3 style={sectionTitleStyle}>
            <Grid size={18} color="var(--primary)" style={{ marginRight: '8px' }} />
            Master Data Quality Report
          </h3>
          <div style={metricsGridStyle}>
            {/* Metric 1: Total Records */}
            <div className="glass-panel" style={metricCardStyle}>
              <span style={metricLabelStyle}>Total Records</span>
              <span style={metricValueStyle}>{masterAudit.totalRecords}</span>
            </div>

            {/* Metric 2: Duplicate Ref Codes */}
            <div className="glass-panel" style={metricCardStyle}>
              <span style={metricLabelStyle}>Duplicate Ref Codes</span>
              <span style={{ ...metricValueStyle, color: masterAudit.duplicateRefCodes > 0 ? 'var(--error)' : 'var(--text-primary)' }}>
                {masterAudit.duplicateRefCodes}
              </span>
            </div>

            {/* Metric 3: Blank Ref Codes */}
            <div className="glass-panel" style={metricCardStyle}>
              <span style={metricLabelStyle}>Blank Ref Codes</span>
              <span style={{ ...metricValueStyle, color: masterAudit.blankRefCodes > 0 ? 'var(--error)' : 'var(--text-primary)' }}>
                {masterAudit.blankRefCodes}
              </span>
            </div>

            {/* Metric 4: Blank Commuter Names */}
            <div className="glass-panel" style={metricCardStyle}>
              <span style={metricLabelStyle}>Blank Names</span>
              <span style={{ ...metricValueStyle, color: masterAudit.blankNames > 0 ? 'var(--error)' : 'var(--text-primary)' }}>
                {masterAudit.blankNames}
              </span>
            </div>

            {/* Metric 5: Blank Landmarks */}
            <div className="glass-panel" style={metricCardStyle}>
              <span style={metricLabelStyle}>Blank Landmarks</span>
              <span style={{ ...metricValueStyle, color: masterAudit.blankLandmarks > 0 ? 'var(--error)' : 'var(--text-primary)' }}>
                {masterAudit.blankLandmarks}
              </span>
            </div>

            {/* Metric 6: Errors Found */}
            <div className="glass-panel" style={metricCardStyle}>
              <span style={metricLabelStyle}>Errors Found</span>
              <span style={{ ...metricValueStyle, color: masterAudit.errorsCount > 0 ? 'var(--error)' : 'var(--text-primary)' }}>
                {masterAudit.errorsCount}
              </span>
            </div>
          </div>
        </section>

        {/* Validation Summary Cards */}
        <section style={sectionBlockStyle}>

          <h3 style={sectionTitleStyle}>
            <Grid size={18} color="var(--primary)" style={{ marginRight: '8px' }} />
            Validation Summary
          </h3>
          <div style={metricsGridStyle}>
            {/* Metric 1: Total Records */}
            <div className="glass-panel" style={metricCardStyle}>
              <span style={metricLabelStyle}>Total Records</span>
              <span style={metricValueStyle}>{metrics.totalRecords}</span>
            </div>

            {/* Metric 2: Matched Records */}
            <div className="glass-panel" style={metricCardStyle}>
              <span style={metricLabelStyle}>Matched Records</span>
              <span style={{ ...metricValueStyle, color: metrics.matchedRecords > 0 ? 'var(--success)' : 'var(--text-primary)' }}>
                {metrics.matchedRecords}
              </span>
            </div>

            {/* Metric 3: Missing Ref Codes */}
            <div className="glass-panel" style={metricCardStyle}>
              <span style={metricLabelStyle}>Missing Ref Codes</span>
              <span style={{ ...metricValueStyle, color: metrics.missingRefCodes > 0 ? 'var(--error)' : 'var(--text-primary)' }}>
                {metrics.missingRefCodes}
              </span>
            </div>

            {/* Metric 4: Landmark Changed */}
            <div className="glass-panel" style={metricCardStyle}>
              <span style={metricLabelStyle}>Landmark Changed</span>
              <span style={{ ...metricValueStyle, color: metrics.landmarkChanged > 0 ? 'var(--warning)' : 'var(--text-primary)' }}>
                {metrics.landmarkChanged}
              </span>
            </div>

            {/* Metric 4.1: Landmark Correct */}
            <div className="glass-panel" style={metricCardStyle}>
              <span style={metricLabelStyle}>Landmark Correct</span>
              <span style={{ ...metricValueStyle, color: metrics.landmarkCorrect > 0 ? 'var(--success)' : 'var(--text-primary)' }}>
                {metrics.landmarkCorrect}
              </span>
            </div>

            {/* Metric 4.2: Landmark Missing */}
            <div className="glass-panel" style={metricCardStyle}>
              <span style={metricLabelStyle}>Landmark Missing</span>
              <span style={{ ...metricValueStyle, color: metrics.landmarkMissing > 0 ? 'var(--error)' : 'var(--text-primary)' }}>
                {metrics.landmarkMissing}
              </span>
            </div>

            {/* Metric 5: Duplicate Names */}
            <div className="glass-panel" style={metricCardStyle}>
              <span style={metricLabelStyle}>Duplicate Names</span>
              <span style={{ ...metricValueStyle, color: metrics.duplicateNames > 0 ? 'var(--error)' : 'var(--text-primary)' }}>
                {metrics.duplicateNames}
              </span>
            </div>

            {/* Metric 6: Errors Found */}
            <div className="glass-panel" style={metricCardStyle}>
              <span style={metricLabelStyle}>Errors Found</span>
              <span style={{ ...metricValueStyle, color: metrics.errorsFound > 0 ? 'var(--error)' : 'var(--text-primary)' }}>
                {metrics.errorsFound}
              </span>
            </div>

            {/* Metric 7: Match Confidence */}
            <div className="glass-panel" style={metricCardStyle}>
              <span style={metricLabelStyle}>Match Confidence</span>
              <span
                style={{
                  ...metricValueStyle,
                  color:
                    metrics.averageConfidence >= 80
                      ? 'var(--success)'
                      : metrics.averageConfidence >= 55
                      ? 'var(--warning)'
                      : 'var(--text-muted)',
                }}
              >
                {metrics.averageConfidence > 0 ? `${metrics.averageConfidence}%` : '0%'}
              </span>
            </div>
          </div>
        </section>

        {/* Matching Execution Logs Console */}
        <section style={sectionBlockStyle}>
          <h3 style={sectionTitleStyle}>
            <Terminal size={18} color="var(--primary)" style={{ marginRight: '8px' }} />
            Matching Execution Logs
          </h3>
          <div id="matching-logs-console" className="glass-panel" style={consoleContainerStyle}>
            {logs.length === 0 ? (
              <div style={consolePlaceholderStyle}>
                Awaiting matching engine execution... Upload files and click 'Generate Route Sheet' to validate commuter paths.
              </div>
            ) : (
              <div style={consoleLogListStyle}>
                {logs.map((log, i) => (
                  <div key={i} style={consoleLogLineStyle}>
                    <span style={consoleTimestampStyle}>[{new Date().toLocaleTimeString()}]</span> {log}
                  </div>
                ))}
              </div>
            )}
          </div>
        </section>


        {/* Generated Output Downloads */}
        <section style={sectionBlockStyle}>
          <h3 style={sectionTitleStyle}>
            <FileSpreadsheet size={18} color="var(--primary)" style={{ marginRight: '8px' }} />
            Generated Output
          </h3>
          <div style={downloadsContainerStyle}>
            <button
              id="download-route-sheet-btn"
              onClick={() => triggerDownload(routeSheetBuffer, 'Standardized_Route_Sheet.xlsx')}
              disabled={!routeSheetBuffer}
              style={routeSheetBuffer ? downloadButtonStyle : disabledDownloadButtonStyle}
              className="glass-panel"
            >
              <Download size={16} style={{ marginRight: '8px' }} />
              Download Route Sheet
            </button>
            <button
              id="download-error-report-btn"
              onClick={() => triggerDownload(errorReportBuffer, 'Route_Validation_Report.xlsx')}
              disabled={!errorReportBuffer}
              style={errorReportBuffer ? downloadButtonStyle_error : disabledDownloadButtonStyle}
              className="glass-panel"
            >
              <AlertCircle size={16} style={{ marginRight: '8px' }} />
              Download Validation Report
            </button>
          </div>
        </section>
      </main>

      {/* Footer Branding */}
      <footer style={footerStyle}>
        <div className="container" style={footerContainerStyle}>
          <span>Developed by Stark Labs AI</span>
          <span>Version 1.0</span>
        </div>
      </footer>

      {/* Success Animation Glassmorphic Badge Modal */}
      {showSuccessBadge && (
        <div
          className="success-overlay success-overlay-fade"
          style={successOverlayStyle}
          onClick={() => setShowSuccessBadge(false)}
        >
          <div
            className="success-card success-card-slide"
            style={successCardStyle}
            onClick={(e) => e.stopPropagation()}
          >
            <div className="success-checkmark-wrapper checkmark-bounce" style={successCheckmarkWrapperStyle}>
              <Check size={40} color="var(--success)" />
            </div>
            
            <h2 style={successTitleStyle}>Route Sync Complete</h2>
            <p style={successSubtitleStyle}>
              The route sheets and validation report have been generated and downloaded automatically.
            </p>

            <div style={successStatsContainerStyle}>
              <div style={successStatRowStyle}>
                <span style={successStatLabelStyle}>Total Records:</span>
                <span style={successStatValueStyle}>{metrics.totalRecords}</span>
              </div>
              <div style={successStatRowStyle}>
                <span style={successStatLabelStyle}>Matched Records:</span>
                <span style={successStatValueStyle}>
                  {metrics.matchedRecords} ({metrics.averageConfidence}% Confidence)
                </span>
              </div>
              <div style={successStatRowStyle}>
                <span style={successStatLabelStyle}>Mismatched Landmarks:</span>
                <span style={{ ...successStatValueStyle, color: metrics.landmarkChanged > 0 ? 'var(--warning)' : 'var(--success)' }}>
                  {metrics.landmarkChanged}
                </span>
              </div>
              <div style={successStatRowStyle}>
                <span style={successStatLabelStyle}>Errors Detected:</span>
                <span style={{ ...successStatValueStyle, color: metrics.errorsFound > 0 ? 'var(--error)' : 'var(--success)' }}>
                  {metrics.errorsFound}
                </span>
              </div>
              <div style={successStatRowStyle}>
                <span style={successStatLabelStyle}>Processing Time:</span>
                <span style={successStatValueStyle}>{(elapsedTime / 1000).toFixed(2)}s</span>
              </div>
            </div>

            <button
              onClick={() => setShowSuccessBadge(false)}
              style={successDoneButtonStyle}
              className="glass-panel"
            >
              Close Details [Esc]
            </button>
          </div>
        </div>
      )}
    </div>
  );
}

export default App;

// Inline CSS Styles
const appWrapperStyle: React.CSSProperties = {
  display: 'flex',
  flexDirection: 'column',
  minHeight: '100vh',
  width: '100%',
};

const headerStyle: React.CSSProperties = {
  position: 'sticky',
  top: 0,
  zIndex: 100,
  display: 'flex',
  justifyContent: 'space-between',
  alignItems: 'center',
  padding: '16px 32px',
  borderRadius: '0',
  borderBottom: '1px solid var(--border-color)',
  background: 'var(--bg-surface)',
};

const headerLeftStyle: React.CSSProperties = {
  display: 'flex',
  alignItems: 'center',
  gap: '12px',
};

const logoWrapperStyle: React.CSSProperties = {
  backgroundColor: 'var(--primary-glow)',
  width: '40px',
  height: '40px',
  borderRadius: '10px',
  display: 'flex',
  alignItems: 'center',
  justifyContent: 'center',
  boxShadow: 'var(--shadow-sm)',
};

const logoTextStyle: React.CSSProperties = {
  fontSize: '1.2rem',
  fontWeight: 800,
  lineHeight: 1.1,
  fontFamily: 'var(--font-heading)',
  letterSpacing: '-0.02em',
};

const logoSubtitleStyle: React.CSSProperties = {
  fontSize: '0.65rem',
  color: 'var(--text-secondary)',
  fontWeight: 600,
  textTransform: 'uppercase',
  letterSpacing: '0.05em',
};

const themeToggleButtonStyle: React.CSSProperties = {
  background: 'none',
  border: 'none',
  color: 'var(--text-primary)',
  cursor: 'pointer',
  display: 'flex',
  alignItems: 'center',
  justifyContent: 'center',
  padding: '8px',
  borderRadius: '8px',
  backgroundColor: 'var(--border-color)',
  transition: 'background-color 0.2s',
};

const mainStyle: React.CSSProperties = {
  flexGrow: 1,
  paddingTop: '32px',
  paddingBottom: '60px',
  display: 'flex',
  flexDirection: 'column',
};

const bannerSectionStyle: React.CSSProperties = {
  textAlign: 'center',
  maxWidth: '800px',
  margin: '0 auto 36px auto',
};

const bannerHeadingStyle: React.CSSProperties = {
  fontSize: '3rem',
  lineHeight: 1.1,
  fontWeight: 800,
  marginBottom: '8px',
  fontFamily: 'var(--font-heading)',
  letterSpacing: '-0.03em',
};

const taglineStyle: React.CSSProperties = {
  fontSize: '1.3rem',
  fontWeight: 700,
  marginBottom: '20px',
  fontFamily: 'var(--font-heading)',
  letterSpacing: '0.02em',
  display: 'flex',
  justifyContent: 'center',
  flexWrap: 'wrap',
  gap: '8px',
};

const bannerSubheadingStyle: React.CSSProperties = {
  fontSize: '1rem',
  color: 'var(--text-secondary)',
  lineHeight: 1.5,
};

const uploadGridStyle: React.CSSProperties = {
  display: 'flex',
  flexWrap: 'wrap',
  gap: '24px',
  maxWidth: '960px',
  width: '100%',
  margin: '0 auto 32px auto',
};

const triggerButtonStyleContainer: React.CSSProperties = {
  display: 'flex',
  flexDirection: 'column',
  alignItems: 'center',
  justifyContent: 'center',
  marginBottom: '48px',
  width: '100%',
  gap: '16px',
};

const triggerButtonStyle: React.CSSProperties = {
  backgroundColor: 'var(--primary)',
  color: 'white',
  border: 'none',
  padding: '14px 32px',
  borderRadius: 'var(--radius-md)',
  fontSize: '1rem',
  fontWeight: 700,
  cursor: 'pointer',
  display: 'flex',
  alignItems: 'center',
  boxShadow: 'var(--shadow-md)',
  transition: 'background-color 0.2s, transform 0.1s, opacity 0.2s',
};

const disabledTriggerButtonStyle: React.CSSProperties = {
  ...triggerButtonStyle,
  backgroundColor: 'var(--text-muted)',
  opacity: 0.6,
  cursor: 'not-allowed',
  boxShadow: 'none',
};

const kbdStyle: React.CSSProperties = {
  backgroundColor: 'rgba(255, 255, 255, 0.2)',
  border: '1px solid rgba(255, 255, 255, 0.3)',
  borderRadius: '4px',
  padding: '2px 6px',
  fontSize: '0.75rem',
  marginLeft: '10px',
  fontFamily: 'var(--font-sans)',
  color: 'white',
  boxShadow: '0 1px 0 rgba(0,0,0,0.2)',
};

const progressContainerStyle: React.CSSProperties = {
  maxWidth: '480px',
  width: '100%',
  padding: '16px 20px',
  borderRadius: 'var(--radius-md)',
  backgroundColor: 'var(--bg-surface)',
  boxShadow: 'var(--shadow-sm)',
};

const progressHeaderStyle: React.CSSProperties = {
  display: 'flex',
  justifyContent: 'space-between',
  alignItems: 'center',
  marginBottom: '8px',
  fontSize: '0.85rem',
  fontWeight: 600,
};

const progressStatusTextStyle: React.CSSProperties = {
  color: 'var(--text-primary)',
};

const progressTimerStyle: React.CSSProperties = {
  color: 'var(--primary)',
  fontFamily: 'monospace',
};

const progressBarTrackStyle: React.CSSProperties = {
  width: '100%',
  height: '8px',
  backgroundColor: 'var(--border-color)',
  borderRadius: 'var(--radius-full)',
  overflow: 'hidden',
};

const progressBarFillStyle: React.CSSProperties = {
  height: '100%',
  background: 'linear-gradient(90deg, var(--primary) 0%, hsl(260, 91%, 60%) 100%)',
  borderRadius: 'var(--radius-full)',
  transition: 'width 0.2s cubic-bezier(0.1, 0.8, 0.2, 1)',
  boxShadow: '0 0 8px var(--primary-glow)',
};

const successOverlayStyle: React.CSSProperties = {
  position: 'fixed',
  top: 0,
  left: 0,
  width: '100vw',
  height: '100vh',
  backgroundColor: 'rgba(15, 23, 42, 0.4)',
  backdropFilter: 'blur(12px)',
  WebkitBackdropFilter: 'blur(12px)',
  display: 'flex',
  justifyContent: 'center',
  alignItems: 'center',
  zIndex: 1000,
};

const successCardStyle: React.CSSProperties = {
  background: 'var(--bg-surface-elevated)',
  border: '1px solid var(--border-color)',
  borderRadius: 'var(--radius-lg)',
  padding: '32px',
  maxWidth: '480px',
  width: '90%',
  textAlign: 'center',
  boxShadow: 'var(--shadow-lg)',
  display: 'flex',
  flexDirection: 'column',
  alignItems: 'center',
};

const successCheckmarkWrapperStyle: React.CSSProperties = {
  width: '80px',
  height: '80px',
  borderRadius: '50%',
  backgroundColor: 'var(--success-glow)',
  border: '2px solid var(--success)',
  display: 'flex',
  alignItems: 'center',
  justifyContent: 'center',
  marginBottom: '20px',
};

const successTitleStyle: React.CSSProperties = {
  fontSize: '1.6rem',
  fontWeight: 800,
  marginBottom: '10px',
  color: 'var(--text-primary)',
};

const successSubtitleStyle: React.CSSProperties = {
  fontSize: '0.9rem',
  color: 'var(--text-secondary)',
  marginBottom: '24px',
  lineHeight: 1.5,
};

const successStatsContainerStyle: React.CSSProperties = {
  width: '100%',
  backgroundColor: 'rgba(0, 0, 0, 0.02)',
  border: '1px solid var(--border-color)',
  borderRadius: 'var(--radius-md)',
  padding: '16px',
  marginBottom: '24px',
  display: 'flex',
  flexDirection: 'column',
  gap: '10px',
};

const successStatRowStyle: React.CSSProperties = {
  display: 'flex',
  justifyContent: 'space-between',
  fontSize: '0.85rem',
};

const successStatLabelStyle: React.CSSProperties = {
  color: 'var(--text-secondary)',
  fontWeight: 500,
  textAlign: 'left',
};

const successStatValueStyle: React.CSSProperties = {
  color: 'var(--text-primary)',
  fontWeight: 700,
  textAlign: 'right',
};

const successDoneButtonStyle: React.CSSProperties = {
  backgroundColor: 'var(--success)',
  color: 'white',
  border: 'none',
  padding: '12px 32px',
  borderRadius: 'var(--radius-md)',
  fontSize: '0.95rem',
  fontWeight: 700,
  cursor: 'pointer',
  boxShadow: '0 4px 12px var(--success-glow)',
  transition: 'background-color 0.2s',
  width: '100%',
};

const settingsPanelStyle: React.CSSProperties = {
  maxWidth: '960px',
  width: '100%',
  margin: '0 auto 32px auto',
  padding: '24px',
  borderRadius: 'var(--radius-lg)',
  background: 'var(--bg-surface)',
  border: '1px solid var(--border-color)',
};

const settingsTitleStyle: React.CSSProperties = {
  fontSize: '1.1rem',
  fontWeight: 700,
  marginBottom: '20px',
  color: 'var(--text-primary)',
  display: 'flex',
  alignItems: 'center',
  borderBottom: '1px solid var(--border-color)',
  paddingBottom: '8px',
};

const settingsGridContainerStyle: React.CSSProperties = {
  display: 'grid',
  gridTemplateColumns: 'repeat(auto-fit, minmax(250px, 1fr))',
  gap: '24px',
};

const settingsGroupStyle: React.CSSProperties = {
  display: 'flex',
  flexDirection: 'column',
  textAlign: 'left',
};

const settingsLabelStyle: React.CSSProperties = {
  fontSize: '0.85rem',
  fontWeight: 600,
  color: 'var(--text-primary)',
  marginBottom: '10px',
};

const checkboxWrapperStyle: React.CSSProperties = {
  display: 'flex',
  alignItems: 'center',
  gap: '10px',
  cursor: 'pointer',
  minHeight: '36px',
};

const checkboxStyle: React.CSSProperties = {
  width: '18px',
  height: '18px',
  borderRadius: '4px',
  border: '1px solid var(--border-color)',
  cursor: 'pointer',
};

const checkboxLabelStyle: React.CSSProperties = {
  fontSize: '0.85rem',
  color: 'var(--text-secondary)',
  cursor: 'pointer',
  userSelect: 'none',
};

const timeInputsContainerStyle: React.CSSProperties = {
  display: 'flex',
  gap: '16px',
};

const timeInputWrapperStyle: React.CSSProperties = {
  display: 'flex',
  alignItems: 'center',
  gap: '8px',
};

const timeInputLabelStyle: React.CSSProperties = {
  fontSize: '0.8rem',
  color: 'var(--text-secondary)',
};

const timeInputStyle: React.CSSProperties = {
  background: 'var(--bg-surface-elevated)',
  border: '1px solid var(--border-color)',
  borderRadius: '6px',
  color: 'var(--text-primary)',
  padding: '6px 12px',
  fontSize: '0.9rem',
  outline: 'none',
  width: '100px',
};

const sectionBlockStyle: React.CSSProperties = {
  marginBottom: '40px',
};

const sectionTitleStyle: React.CSSProperties = {
  fontSize: '1.2rem',
  fontWeight: 700,
  marginBottom: '16px',
  color: 'var(--text-primary)',
  display: 'flex',
  alignItems: 'center',
  borderBottom: '1px solid var(--border-color)',
  paddingBottom: '8px',
};

const detailsGridStyle: React.CSSProperties = {
  display: 'grid',
  gridTemplateColumns: 'repeat(auto-fit, minmax(180px, 1fr))',
  gap: '16px',
};

const detailCardStyle: React.CSSProperties = {
  padding: '20px',
  borderRadius: 'var(--radius-md)',
  display: 'flex',
  flexDirection: 'column',
  textAlign: 'left',
  background: 'var(--bg-surface)',
};

const detailLabelStyle: React.CSSProperties = {
  fontSize: '0.75rem',
  fontWeight: 600,
  color: 'var(--text-secondary)',
  textTransform: 'uppercase',
  letterSpacing: '0.05em',
  marginBottom: '8px',
};

const detailValueStyle: React.CSSProperties = {
  fontSize: '1.15rem',
  fontWeight: 700,
  color: 'var(--text-primary)',
  marginBottom: '4px',
};

const detailSubtextStyle: React.CSSProperties = {
  fontSize: '0.7rem',
  color: 'var(--text-muted)',
  whiteSpace: 'nowrap',
  overflow: 'hidden',
  textOverflow: 'ellipsis',
};

const metricsGridStyle: React.CSSProperties = {
  display: 'grid',
  gridTemplateColumns: 'repeat(auto-fit, minmax(140px, 1fr))',
  gap: '16px',
};

const metricCardStyle: React.CSSProperties = {
  padding: '20px 16px',
  borderRadius: 'var(--radius-md)',
  display: 'flex',
  flexDirection: 'column',
  alignItems: 'center',
  justifyContent: 'center',
  textAlign: 'center',
  background: 'var(--bg-surface)',
};

const metricLabelStyle: React.CSSProperties = {
  fontSize: '0.8rem',
  fontWeight: 600,
  color: 'var(--text-secondary)',
  marginBottom: '10px',
};

const metricValueStyle: React.CSSProperties = {
  fontSize: '1.8rem',
  fontWeight: 800,
  color: 'var(--text-primary)',
};

const downloadsContainerStyle: React.CSSProperties = {
  display: 'flex',
  flexWrap: 'wrap',
  gap: '16px',
};

const downloadButtonStyle: React.CSSProperties = {
  backgroundColor: 'var(--bg-surface-elevated)',
  border: '1px solid var(--primary)',
  color: 'var(--primary)',
  padding: '12px 24px',
  borderRadius: 'var(--radius-md)',
  fontSize: '0.9rem',
  fontWeight: 700,
  cursor: 'pointer',
  display: 'flex',
  alignItems: 'center',
  transition: 'background-color 0.2s, box-shadow 0.2s',
  boxShadow: 'var(--shadow-sm)',
};

const downloadButtonStyle_error: React.CSSProperties = {
  ...downloadButtonStyle,
  border: '1px solid var(--error)',
  color: 'var(--error)',
};

const disabledDownloadButtonStyle: React.CSSProperties = {
  ...downloadButtonStyle,
  border: '1px solid var(--border-color)',
  color: 'var(--text-muted)',
  background: 'var(--bg-surface)',
  opacity: 0.5,
  cursor: 'not-allowed',
  boxShadow: 'none',
};

const footerStyle: React.CSSProperties = {
  borderTop: '1px solid var(--border-color)',
  padding: '20px 0',
  marginTop: 'auto',
  background: 'var(--bg-surface)',
};

const footerContainerStyle: React.CSSProperties = {
  display: 'flex',
  justifyContent: 'space-between',
  alignItems: 'center',
  fontSize: '0.8rem',
  color: 'var(--text-muted)',
  fontWeight: 500,
};

const consoleContainerStyle: React.CSSProperties = {
  backgroundColor: 'var(--bg-surface-elevated, #1e1e1e)',
  border: '1px solid var(--border-color)',
  borderRadius: 'var(--radius-md)',
  padding: '16px',
  fontFamily: 'monospace, Courier New, Courier',
  fontSize: '0.85rem',
  color: 'var(--text-primary)',
  maxHeight: '240px',
  overflowY: 'auto',
  boxShadow: 'inset 0 2px 4px rgba(0,0,0,0.06)',
  display: 'flex',
  flexDirection: 'column',
};

const consolePlaceholderStyle: React.CSSProperties = {
  color: 'var(--text-muted)',
  fontStyle: 'italic',
  textAlign: 'center',
  padding: '24px 0',
  width: '100%',
};

const consoleLogListStyle: React.CSSProperties = {
  display: 'flex',
  flexDirection: 'column',
  gap: '6px',
  textAlign: 'left',
};

const consoleLogLineStyle: React.CSSProperties = {
  lineHeight: 1.4,
  wordBreak: 'break-all',
  borderLeft: '2px solid var(--primary)',
  paddingLeft: '8px',
};

const consoleTimestampStyle: React.CSSProperties = {
  color: 'var(--primary)',
  marginRight: '6px',
  fontWeight: 600,
};

// Global Styles Injection
if (typeof document !== 'undefined') {
  const style = document.createElement('style');
  style.innerHTML = `
    .spin-animation {
      animation: spin 1s linear infinite;
    }
    @keyframes spin {
      from { transform: rotate(0deg); }
      to { transform: rotate(360deg); }
    }
    
    .success-overlay-fade {
      animation: successFadeIn 0.3s ease-out forwards;
    }
    .success-card-slide {
      animation: successSlideIn 0.4s cubic-bezier(0.34, 1.56, 0.64, 1) forwards;
    }
    .checkmark-bounce {
      animation: checkmarkZoom 0.5s cubic-bezier(0.34, 1.56, 0.64, 1) forwards, checkmarkPulse 2s infinite;
    }
    
    @keyframes successFadeIn {
      from { opacity: 0; }
      to { opacity: 1; }
    }
    @keyframes successSlideIn {
      from {
        opacity: 0;
        transform: scale(0.9) translateY(20px);
      }
      to {
        opacity: 1;
        transform: scale(1) translateY(0);
      }
    }
    @keyframes checkmarkZoom {
      0% {
        transform: scale(0);
        opacity: 0;
      }
      50% {
        transform: scale(1.2);
      }
      100% {
        transform: scale(1);
        opacity: 1;
      }
    }
    @keyframes checkmarkPulse {
      0% {
        box-shadow: 0 0 0 0 rgba(16, 185, 129, 0.4);
      }
      70% {
        box-shadow: 0 0 0 12px rgba(16, 185, 129, 0);
      }
      100% {
        box-shadow: 0 0 0 0 rgba(16, 185, 129, 0);
      }
    }
  `;
  document.head.appendChild(style);
}
