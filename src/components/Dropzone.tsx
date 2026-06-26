import React, { useState, useRef } from 'react';
import { UploadCloud, FileSpreadsheet, AlertCircle } from 'lucide-react';
import { logger } from '../utils/logger';

interface DropzoneProps {
  onFileSelect: (file: File) => void;
  isLoading: boolean;
  progress: number;
}

export const Dropzone: React.FC<DropzoneProps> = ({ onFileSelect, isLoading, progress }) => {
  const [isDragActive, setIsDragActive] = useState(false);
  const [errorMessage, setErrorMessage] = useState<string | null>(null);
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

  const validateAndSelectFile = (file: File) => {
    setErrorMessage(null);
    const maxSizeBytes = 50 * 1024 * 1024; // 50MB
    const validExtensions = ['.xlsx', '.xls', '.csv'];
    const fileExtension = file.name.substring(file.name.lastIndexOf('.')).toLowerCase();

    if (!validExtensions.includes(fileExtension)) {
      const err = `Invalid file type. Please upload an Excel (.xlsx, .xls) or CSV file.`;
      logger.warn(err);
      setErrorMessage(err);
      return;
    }

    if (file.size > maxSizeBytes) {
      const err = `File size exceeds the 50MB upload limit.`;
      logger.warn(err);
      setErrorMessage(err);
      return;
    }

    logger.info(`File validation passed: ${file.name} (${(file.size / (1024 * 1024)).toFixed(2)} MB)`);
    onFileSelect(file);
  };

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragActive(false);

    if (isLoading) return;

    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      validateAndSelectFile(e.dataTransfer.files[0]);
    }
  };

  const handleChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    e.preventDefault();
    if (isLoading) return;

    if (e.target.files && e.target.files[0]) {
      validateAndSelectFile(e.target.files[0]);
    }
  };

  const onButtonClick = () => {
    if (isLoading) return;
    fileInputRef.current?.click();
  };

  return (
    <div
      className="glass-panel"
      style={{
        ...containerStyle,
        borderColor: isDragActive ? 'var(--primary)' : 'var(--border-color)',
        backgroundColor: isDragActive ? 'var(--primary-glow)' : 'var(--bg-surface)',
        transform: isDragActive ? 'scale(1.01)' : 'scale(1)',
      }}
      onDragEnter={handleDrag}
      onDragOver={handleDrag}
      onDragLeave={handleDrag}
      onDrop={handleDrop}
    >
      <input
        id="excel-dropzone-input"
        ref={fileInputRef}
        type="file"
        style={{ display: 'none' }}
        accept=".xlsx, .xls, .csv"
        onChange={handleChange}
        disabled={isLoading}
      />

      {isLoading ? (
        <div style={contentStyle}>
          <div style={spinnerContainerStyle}>
            <div style={spinnerStyle}></div>
            <FileSpreadsheet size={28} color="var(--primary)" style={iconCenteredStyle} />
          </div>
          <h3 style={titleStyle}>Parsing Document</h3>
          <p style={subStyle}>Please wait while the Web Worker processes file structure...</p>
          <div style={progressWrapperStyle}>
            <div style={progressBarContainerStyle}>
              <div style={{ ...progressBarFillStyle, width: `${progress}%` }} />
            </div>
            <span style={progressLabelStyle}>{progress}%</span>
          </div>
        </div>
      ) : (
        <div style={contentStyle}>
          <UploadCloud
            size={48}
            color={isDragActive ? 'var(--primary)' : 'var(--text-secondary)'}
            style={{ marginBottom: '16px', transition: 'color var(--transition-fast)' }}
          />
          <h3 style={titleStyle}>Upload your Excel files</h3>
          <p style={subStyle}>Drag & drop sheets here or click to browse files</p>
          <button id="excel-browse-btn" onClick={onButtonClick} style={buttonStyle} className="glass-panel">
            Browse Files
          </button>
          <span style={hintStyle}>Supports .xlsx, .xls, and .csv (Max 50MB)</span>

          {errorMessage && (
            <div style={errorContainerStyle}>
              <AlertCircle size={16} color="var(--error)" style={{ marginRight: '6px' }} />
              <span>{errorMessage}</span>
            </div>
          )}
        </div>
      )}
    </div>
  );
};

// Styles
const containerStyle: React.CSSProperties = {
  border: '2px dashed var(--border-color)',
  borderRadius: 'var(--radius-lg)',
  padding: '48px 24px',
  textAlign: 'center',
  cursor: 'pointer',
  transition: 'border-color var(--transition-normal), background-color var(--transition-normal), transform var(--transition-normal)',
  outline: 'none',
  position: 'relative',
  overflow: 'hidden',
};

const contentStyle: React.CSSProperties = {
  display: 'flex',
  flexDirection: 'column',
  alignItems: 'center',
  justifyContent: 'center',
};

const titleStyle: React.CSSProperties = {
  fontSize: '1.25rem',
  fontWeight: 600,
  marginBottom: '8px',
  color: 'var(--text-primary)',
};

const subStyle: React.CSSProperties = {
  fontSize: '0.875rem',
  color: 'var(--text-secondary)',
  marginBottom: '24px',
};

const buttonStyle: React.CSSProperties = {
  background: 'var(--bg-surface-elevated)',
  color: 'var(--text-primary)',
  fontWeight: 600,
  fontSize: '0.875rem',
  padding: '10px 24px',
  borderRadius: 'var(--radius-sm)',
  cursor: 'pointer',
  outline: 'none',
  marginBottom: '16px',
  transition: 'background-color 0.2s',
};

const hintStyle: React.CSSProperties = {
  fontSize: '0.75rem',
  color: 'var(--text-muted)',
};

const errorContainerStyle: React.CSSProperties = {
  display: 'flex',
  alignItems: 'center',
  justifyContent: 'center',
  marginTop: '16px',
  padding: '8px 16px',
  backgroundColor: 'var(--error-glow)',
  border: '1px solid var(--error)',
  borderRadius: 'var(--radius-sm)',
  color: 'var(--error)',
  fontSize: '0.8rem',
  fontWeight: 500,
};

const spinnerContainerStyle: React.CSSProperties = {
  position: 'relative',
  width: '64px',
  height: '64px',
  marginBottom: '20px',
  display: 'flex',
  alignItems: 'center',
  justifyContent: 'center',
};

const spinnerStyle: React.CSSProperties = {
  position: 'absolute',
  top: 0,
  left: 0,
  width: '100%',
  height: '100%',
  border: '4px solid var(--border-color)',
  borderTopColor: 'var(--primary)',
  borderRadius: '50%',
  animation: 'spin 1s linear infinite',
};

const iconCenteredStyle: React.CSSProperties = {
  position: 'absolute',
  zIndex: 1,
};

const progressWrapperStyle: React.CSSProperties = {
  width: '100%',
  maxWidth: '280px',
  display: 'flex',
  alignItems: 'center',
  gap: '12px',
};

const progressBarContainerStyle: React.CSSProperties = {
  flexGrow: 1,
  height: '6px',
  backgroundColor: 'var(--border-color)',
  borderRadius: 'var(--radius-full)',
  overflow: 'hidden',
};

const progressBarFillStyle: React.CSSProperties = {
  height: '100%',
  backgroundColor: 'var(--primary)',
  borderRadius: 'var(--radius-full)',
  transition: 'width 0.2s cubic-bezier(0.16, 1, 0.3, 1)',
};

const progressLabelStyle: React.CSSProperties = {
  fontSize: '0.75rem',
  fontWeight: 600,
  color: 'var(--text-secondary)',
  minWidth: '32px',
};

// Add standard spinning animation keyframe to head
if (typeof window !== 'undefined' && typeof document !== 'undefined') {
  const style = document.createElement('style');
  style.innerHTML = `
    @keyframes spin {
      0% { transform: rotate(0deg); }
      100% { transform: rotate(360deg); }
    }
  `;
  document.head.appendChild(style);
}
