import React, { useState, useRef } from 'react';
import { UploadCloud, CheckCircle, Trash2, RefreshCw, AlertCircle, ChevronDown, ChevronUp, Link } from 'lucide-react';
import { logger } from '../utils/logger';
import { SYSTEM_COLUMNS, detectColumnMappings } from '../utils/columnDetector';

interface UploadCardProps {
  idPrefix: string;
  title: string;
  onFileLoaded: (
    data: Record<string, unknown>[],
    fileName: string,
    rows: number,
    cols: number,
    mappings: Record<string, string>
  ) => void;
  onFileCleared: () => void;
  onMappingChange: (mappings: Record<string, string>) => void;
  isLoading: boolean;
  progress: number;
}


export const UploadCard: React.FC<UploadCardProps> = ({
  idPrefix,
  title,
  onFileLoaded,
  onFileCleared,
  onMappingChange,
  isLoading,
  progress,
}) => {
  const [isDragActive, setIsDragActive] = useState(false);
  const [fileName, setFileName] = useState<string | null>(null);
  const [errorMessage, setErrorMessage] = useState<string | null>(null);
  const [fileDetails, setFileDetails] = useState<{
    size: string;
    rows: number;
    cols: number;
    time: string;
  } | null>(null);
  const [sheetColumns, setSheetColumns] = useState<string[]>([]);
  const [mappings, setMappings] = useState<Record<string, string>>({});
  const [isMappingExpanded, setIsMappingExpanded] = useState(true);
  const fileInputRef = useRef<HTMLInputElement>(null);


  const handleDrag = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    if (e.type === 'dragenter' || e.type === 'dragover') {
      setIsDragActive(true);
    } else if (e.type === 'dragleave') {
      setIsDragActive(false);
    }
  };

  const processFile = (file: File) => {
    setErrorMessage(null);
    const validExtensions = ['.xlsx', '.xls'];
    const fileExtension = file.name.substring(file.name.lastIndexOf('.')).toLowerCase();

    if (!validExtensions.includes(fileExtension)) {
      const err = `Invalid file format. Only .xls and .xlsx Excel files are accepted.`;
      logger.warn(`[${title}] File validation failed: ${file.name}`);
      setErrorMessage(err);
      return;
    }

    if (file.size === 0) {
      const err = `Empty file. The selected spreadsheet contains no data (0 bytes).`;
      logger.warn(`[${title}] Empty file check failed: ${file.name}`);
      setErrorMessage(err);
      return;
    }

    setFileName(file.name);
    logger.info(`[${title}] File accepted: ${file.name} (${file.size} bytes)`);

    // Parse file using FileReader and helper
    const reader = new FileReader();
    reader.onload = async (e) => {
      try {
        const buffer = e.target?.result as ArrayBuffer;
        if (!buffer) throw new Error('File buffer is empty');
        
        const { ExcelHelper } = await import('../utils/excelHelper');
        const target = idPrefix === 'master-upload' ? 'master' : 'route';
        const parseResult = await ExcelHelper.parseFile(file, target);
        
        const rowsCount = parseResult.rowCount;
        const cols = parseResult.headers;
        const previewRows = parseResult.previewRows;

        if (rowsCount === 0) {
          throw new Error('Worksheet contains no data rows.');
        }

        // Run column detection using preview rows
        const autoMappings = detectColumnMappings(cols, previewRows);
        setSheetColumns(cols);
        setMappings(autoMappings);

        const now = new Date();
        const uploadTime = now.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit', second: '2-digit' });

        setFileDetails({
          size: `${(file.size / 1024).toFixed(1)} KB`,
          rows: rowsCount,
          cols: cols.length,
          time: uploadTime,
        });

        onFileLoaded(previewRows, file.name, rowsCount, cols.length, autoMappings);

      } catch (err: any) {
        logger.error(`[${title}] Parse error`, err);
        setErrorMessage(err?.message || 'Failed to read or parse Excel worksheet.');
        setFileName(null);
        setFileDetails(null);
        onFileCleared();
      }
    };
    reader.readAsArrayBuffer(file);
  };


  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragActive(false);
    if (isLoading || fileName) return;

    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      processFile(e.dataTransfer.files[0]);
    }
  };

  const handleChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    e.preventDefault();
    if (isLoading) return;

    if (e.target.files && e.target.files[0]) {
      processFile(e.target.files[0]);
    }
  };

  const onButtonClick = (e: React.MouseEvent) => {
    e.stopPropagation();
    if (isLoading) return;
    fileInputRef.current?.click();
  };

  const handleRemove = async (e: React.MouseEvent) => {
    e.stopPropagation();
    setFileName(null);
    setErrorMessage(null);
    setFileDetails(null);
    setSheetColumns([]);
    setMappings({});
    if (fileInputRef.current) fileInputRef.current.value = '';
    onFileCleared();

    try {
      const { ExcelHelper } = await import('../utils/excelHelper');
      await ExcelHelper.clearCache(idPrefix === 'master-upload' ? 'master' : 'route');
    } catch (err) {
      logger.error('Failed to clear worker cache', err);
    }

    logger.info(`[${title}] File removed by user.`);
  };



  return (
    <div
      className="glass-panel"
      style={{
        ...cardWrapperStyle,
        borderColor: isDragActive ? 'var(--primary)' : 'var(--border-color)',
        backgroundColor: isDragActive ? 'var(--primary-glow)' : 'var(--bg-surface)',
      }}
    >
      <h3 style={cardTitleStyle}>{title}</h3>

      <div
        style={{
          ...uploadAreaStyle,
          borderColor: isDragActive ? 'var(--primary)' : 'var(--border-color)',
        }}
        onDragEnter={handleDrag}
        onDragOver={handleDrag}
        onDragLeave={handleDrag}
        onDrop={handleDrop}
      >
        <input
          id={`${idPrefix}-file-input`}
          ref={fileInputRef}
          type="file"
          accept=".xlsx, .xls"
          style={{ display: 'none' }}
          onChange={handleChange}
          disabled={isLoading}
        />

        {isLoading ? (
          <div style={innerContentStyle}>
            <div style={spinnerStyle}></div>
            <p style={statusTextStyle}>Parsing worksheet ({progress}%)</p>
          </div>
        ) : fileName ? (
          /* File Loaded View */
          <div style={innerContentStyle}>
            <CheckCircle size={36} color="var(--success)" style={{ marginBottom: '12px' }} />
            <h4 style={fileNameTextStyle} title={fileName}>
              {fileName}
            </h4>
            <div style={successBadgeStyle}>Upload Success</div>
            
            {fileDetails && (
              <div style={metadataGridStyle}>
                <div style={metadataGridItemStyle}>
                  <span style={metadataGridLabelStyle}>Size:</span>
                  <span style={metadataGridValueStyle}>{fileDetails.size}</span>
                </div>
                <div style={metadataGridItemStyle}>
                  <span style={metadataGridLabelStyle}>Rows:</span>
                  <span style={metadataGridValueStyle}>{fileDetails.rows}</span>
                </div>
                <div style={metadataGridItemStyle}>
                  <span style={metadataGridLabelStyle}>Cols:</span>
                  <span style={metadataGridValueStyle}>{fileDetails.cols}</span>
                </div>
                <div style={metadataGridItemStyle}>
                  <span style={metadataGridLabelStyle}>Time:</span>
                  <span style={metadataGridValueStyle}>{fileDetails.time}</span>
                </div>
              </div>
            )}

            {/* Column Mappings section */}
            <div style={mappingSectionStyle}>
              <button
                type="button"
                onClick={() => setIsMappingExpanded(!isMappingExpanded)}
                style={mappingHeaderStyle}
                className="glass-panel"
              >
                <div style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                  <Link size={13} color="var(--primary)" />
                  <span style={mappingTitleStyle}>Column Mappings</span>
                </div>
                {isMappingExpanded ? <ChevronUp size={14} /> : <ChevronDown size={14} />}
              </button>

              {isMappingExpanded && (
                <div style={mappingGridStyle} className="animate-fade-in">
                  {SYSTEM_COLUMNS.map((sysCol) => {
                    const mappedValue = mappings[sysCol.key] || '';
                    
                    return (
                      <div key={sysCol.key} style={mappingRowStyle}>
                        <label style={mappingLabelStyle} title={sysCol.label}>
                          {sysCol.label.split(' / ')[0]}
                        </label>
                        <select
                          id={`${idPrefix}-map-${sysCol.key}`}
                          value={mappedValue}
                          onChange={(e) => {
                            const newMapping = {
                              ...mappings,
                              [sysCol.key]: e.target.value,
                            };
                            setMappings(newMapping);
                            onMappingChange(newMapping);
                            logger.info(`[${title}] Manually changed mapping for ${sysCol.key} -> ${e.target.value}`);
                          }}
                          style={mappingSelectStyle}
                        >
                          <option value="">-- Select Column --</option>
                          {sheetColumns.map((colName) => (
                            <option key={colName} value={colName}>
                              {colName}
                            </option>
                          ))}
                        </select>
                      </div>
                    );
                  })}
                </div>
              )}
            </div>

            <div style={actionsContainerStyle}>


              <button
                id={`${idPrefix}-replace-btn`}
                onClick={onButtonClick}
                style={actionButtonStyle}
                className="glass-panel"
              >
                <RefreshCw size={12} style={{ marginRight: '4px' }} />
                Replace <span style={shortcutBadgeStyle_small}>{idPrefix === 'master-upload' ? 'Alt+M' : 'Alt+R'}</span>
              </button>
              <button
                id={`${idPrefix}-remove-btn`}
                onClick={handleRemove}
                style={{ ...actionButtonStyle, color: 'var(--error)' }}
                className="glass-panel"
              >
                <Trash2 size={12} style={{ marginRight: '4px' }} />
                Remove
              </button>
            </div>
          </div>
        ) : (
          /* Drag and Drop Prompt */
          <div style={innerContentStyle} onClick={onButtonClick}>
            <UploadCloud size={40} color="var(--text-secondary)" style={{ marginBottom: '12px' }} />
            <p style={dragPromptStyle}>Drag & drop file here or click to browse</p>
            <button
              id={`${idPrefix}-browse-btn`}
              onClick={onButtonClick}
              style={browseButtonStyle}
              className="glass-panel"
            >
              Browse File <span style={shortcutBadgeStyle}>{idPrefix === 'master-upload' ? 'Alt + M' : 'Alt + R'}</span>
            </button>
            <span style={fileHintStyle}>Accepts .xls and .xlsx</span>
          </div>
        )}
      </div>

      {errorMessage && (
        <div style={errorContainerStyle}>
          <AlertCircle size={14} color="var(--error)" style={{ marginRight: '6px', flexShrink: 0 }} />
          <span>{errorMessage}</span>
        </div>
      )}
    </div>
  );
};

// Styles
const cardWrapperStyle: React.CSSProperties = {
  flex: '1 1 300px',
  display: 'flex',
  flexDirection: 'column',
  padding: '24px',
  borderRadius: 'var(--radius-lg)',
  minWidth: '280px',
};

const cardTitleStyle: React.CSSProperties = {
  fontSize: '1.1rem',
  fontWeight: 700,
  marginBottom: '16px',
  color: 'var(--text-primary)',
  textAlign: 'left',
};

const uploadAreaStyle: React.CSSProperties = {
  border: '2px dashed var(--border-color)',
  borderRadius: 'var(--radius-md)',
  padding: '24px 16px',
  display: 'flex',
  flexDirection: 'column',
  alignItems: 'center',
  justifyContent: 'center',
  cursor: 'pointer',
  minHeight: '180px',
  transition: 'border-color var(--transition-fast), background-color var(--transition-fast)',
  position: 'relative',
};

const innerContentStyle: React.CSSProperties = {
  display: 'flex',
  flexDirection: 'column',
  alignItems: 'center',
  textAlign: 'center',
  width: '100%',
};

const dragPromptStyle: React.CSSProperties = {
  fontSize: '0.85rem',
  color: 'var(--text-secondary)',
  marginBottom: '12px',
};

const browseButtonStyle: React.CSSProperties = {
  background: 'var(--bg-surface-elevated)',
  color: 'var(--text-primary)',
  border: '1px solid var(--border-color)',
  fontSize: '0.8rem',
  fontWeight: 600,
  padding: '6px 16px',
  borderRadius: '6px',
  cursor: 'pointer',
  marginBottom: '8px',
};

const fileHintStyle: React.CSSProperties = {
  fontSize: '0.7rem',
  color: 'var(--text-muted)',
};

const fileNameTextStyle: React.CSSProperties = {
  fontSize: '0.9rem',
  fontWeight: 600,
  color: 'var(--text-primary)',
  marginBottom: '4px',
  maxWidth: '220px',
  whiteSpace: 'nowrap',
  overflow: 'hidden',
  textOverflow: 'ellipsis',
};

const successBadgeStyle: React.CSSProperties = {
  fontSize: '0.7rem',
  fontWeight: 700,
  color: 'var(--success)',
  backgroundColor: 'var(--success-glow)',
  padding: '2px 8px',
  borderRadius: 'var(--radius-full)',
  marginBottom: '16px',
};

const actionsContainerStyle: React.CSSProperties = {
  display: 'flex',
  gap: '8px',
};

const actionButtonStyle: React.CSSProperties = {
  background: 'var(--bg-surface-elevated)',
  border: '1px solid var(--border-color)',
  color: 'var(--text-secondary)',
  fontSize: '0.75rem',
  fontWeight: 600,
  padding: '6px 12px',
  borderRadius: '4px',
  cursor: 'pointer',
  display: 'flex',
  alignItems: 'center',
};

const errorContainerStyle: React.CSSProperties = {
  display: 'flex',
  alignItems: 'center',
  marginTop: '12px',
  padding: '8px 12px',
  backgroundColor: 'var(--error-glow)',
  border: '1px solid var(--error)',
  borderRadius: '6px',
  color: 'var(--error)',
  fontSize: '0.75rem',
  fontWeight: 500,
  textAlign: 'left',
};

const spinnerStyle: React.CSSProperties = {
  width: '32px',
  height: '32px',
  border: '3px solid var(--border-color)',
  borderTopColor: 'var(--primary)',
  borderRadius: '50%',
  animation: 'spin 1s linear infinite',
  marginBottom: '12px',
};

const statusTextStyle: React.CSSProperties = {
  fontSize: '0.8rem',
  color: 'var(--text-secondary)',
  fontWeight: 500,
};

const metadataGridStyle: React.CSSProperties = {
  display: 'grid',
  gridTemplateColumns: '1fr 1fr',
  gap: '8px 16px',
  width: '100%',
  maxWidth: '240px',
  margin: '8px 0 16px 0',
  padding: '10px',
  background: 'rgba(0, 0, 0, 0.02)',
  borderRadius: '6px',
  border: '1px solid var(--border-color)',
};

const metadataGridItemStyle: React.CSSProperties = {
  display: 'flex',
  justifyContent: 'space-between',
  fontSize: '0.75rem',
  gap: '4px',
};

const metadataGridLabelStyle: React.CSSProperties = {
  color: 'var(--text-secondary)',
  fontWeight: 500,
};

const metadataGridValueStyle: React.CSSProperties = {
  color: 'var(--text-primary)',
  fontWeight: 600,
};

const mappingSectionStyle: React.CSSProperties = {
  width: '100%',
  maxWidth: '280px',
  marginBottom: '16px',
  textAlign: 'left',
};

const mappingHeaderStyle: React.CSSProperties = {
  width: '100%',
  display: 'flex',
  justifyContent: 'space-between',
  alignItems: 'center',
  padding: '8px 12px',
  background: 'var(--bg-surface-elevated)',
  border: '1px solid var(--border-color)',
  borderRadius: '6px',
  cursor: 'pointer',
  fontSize: '0.8rem',
  fontWeight: 600,
  color: 'var(--text-primary)',
  outline: 'none',
};

const mappingTitleStyle: React.CSSProperties = {
  fontSize: '0.75rem',
  fontWeight: 700,
};

const mappingGridStyle: React.CSSProperties = {
  display: 'flex',
  flexDirection: 'column',
  gap: '8px',
  marginTop: '8px',
  padding: '12px',
  background: 'rgba(0, 0, 0, 0.01)',
  borderRadius: '6px',
  border: '1px solid var(--border-color)',
  maxHeight: '180px',
  overflowY: 'auto',
};

const mappingRowStyle: React.CSSProperties = {
  display: 'flex',
  justifyContent: 'space-between',
  alignItems: 'center',
  gap: '12px',
};

const mappingLabelStyle: React.CSSProperties = {
  fontSize: '0.75rem',
  fontWeight: 500,
  color: 'var(--text-secondary)',
  whiteSpace: 'nowrap',
  overflow: 'hidden',
  textOverflow: 'ellipsis',
  maxWidth: '120px',
};

const mappingSelectStyle: React.CSSProperties = {
  padding: '4px 6px',
  fontSize: '0.7rem',
  borderRadius: '4px',
  background: 'var(--bg-surface)',
  color: 'var(--text-primary)',
  border: '1px solid var(--border-color)',
  maxWidth: '110px',
  flexGrow: 1,
};

const shortcutBadgeStyle: React.CSSProperties = {
  fontSize: '0.65rem',
  backgroundColor: 'var(--border-color)',
  border: '1px solid rgba(0,0,0,0.05)',
  borderRadius: '3px',
  padding: '2px 5px',
  marginLeft: '8px',
  color: 'var(--text-secondary)',
  fontWeight: 'normal',
  display: 'inline-block',
};

const shortcutBadgeStyle_small: React.CSSProperties = {
  fontSize: '0.6rem',
  backgroundColor: 'var(--border-color)',
  borderRadius: '2px',
  padding: '1px 3px',
  marginLeft: '4px',
  color: 'var(--text-muted)',
  fontWeight: 'normal',
};


