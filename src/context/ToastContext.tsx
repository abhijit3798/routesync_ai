import React, { createContext, useContext, useState, useCallback } from 'react';
import { AlertCircle, CheckCircle, Info, X, AlertTriangle } from 'lucide-react';
import { logger } from '../utils/logger';

export type ToastType = 'success' | 'error' | 'warning' | 'info';

export interface Toast {
  id: string;
  message: string;
  type: ToastType;
  duration?: number;
}

interface ToastContextType {
  toasts: Toast[];
  addToast: (message: string, type: ToastType, duration?: number) => void;
  removeToast: (id: string) => void;
}

const ToastContext = createContext<ToastContextType | undefined>(undefined);

export const ToastProvider: React.FC<{ children: React.ReactNode }> = ({ children }) => {
  const [toasts, setToasts] = useState<Toast[]>([]);

  const removeToast = useCallback((id: string) => {
    setToasts((prev) => prev.filter((t) => t.id !== id));
  }, []);

  const addToast = useCallback((message: string, type: ToastType, duration = 4000) => {
    const id = Math.random().toString(36).substring(2, 9);
    setToasts((prev) => [...prev, { id, message, type, duration }]);
    logger.info(`Toast added [${type}]: ${message}`);

    setTimeout(() => {
      removeToast(id);
    }, duration);
  }, [removeToast]);

  return (
    <ToastContext.Provider value={{ toasts, addToast, removeToast }}>
      {children}
      <ToastContainer toasts={toasts} removeToast={removeToast} />
    </ToastContext.Provider>
  );
};

export const useToast = (): ToastContextType => {
  const context = useContext(ToastContext);
  if (context === undefined) {
    throw new Error('useToast must be used within a ToastProvider');
  }
  return context;
};

// Toast Container & Individual Toast Component styles
const ToastContainer: React.FC<{ toasts: Toast[]; removeToast: (id: string) => void }> = ({
  toasts,
  removeToast,
}) => {
  return (
    <div style={containerStyle}>
      {toasts.map((toast) => (
        <ToastItem key={toast.id} toast={toast} onClose={() => removeToast(toast.id)} />
      ))}
    </div>
  );
};

const ToastItem: React.FC<{ toast: Toast; onClose: () => void }> = ({ toast, onClose }) => {
  const getIcon = () => {
    switch (toast.type) {
      case 'success':
        return <CheckCircle size={18} color="var(--success)" />;
      case 'error':
        return <AlertCircle size={18} color="var(--error)" />;
      case 'warning':
        return <AlertTriangle size={18} color="var(--warning)" />;
      case 'info':
      default:
        return <Info size={18} color="var(--info)" />;
    }
  };

  const getBorderColor = () => {
    switch (toast.type) {
      case 'success':
        return 'var(--success)';
      case 'error':
        return 'var(--error)';
      case 'warning':
        return 'var(--warning)';
      case 'info':
      default:
        return 'var(--info)';
    }
  };

  return (
    <div
      className="glass-panel"
      style={{
        ...toastStyle,
        borderLeft: `4px solid ${getBorderColor()}`,
      }}
    >
      <div style={iconStyle}>{getIcon()}</div>
      <div style={messageStyle}>{toast.message}</div>
      <button onClick={onClose} style={closeButtonStyle} aria-label="Close notification">
        <X size={16} />
      </button>
    </div>
  );
};

// Inline Styles (complements global variables)
const containerStyle: React.CSSProperties = {
  position: 'fixed',
  bottom: '24px',
  right: '24px',
  zIndex: 9999,
  display: 'flex',
  flexDirection: 'column',
  gap: '12px',
  maxWidth: '400px',
  width: 'calc(100% - 48px)',
  pointerEvents: 'none',
};

const toastStyle: React.CSSProperties = {
  display: 'flex',
  alignItems: 'center',
  padding: '12px 16px',
  borderRadius: '8px',
  background: 'var(--bg-surface-elevated)',
  boxShadow: 'var(--shadow-lg)',
  border: '1px solid var(--border-color)',
  pointerEvents: 'auto',
  animation: 'toastSlideIn 0.3s cubic-bezier(0.16, 1, 0.3, 1) forwards',
};

const iconStyle: React.CSSProperties = {
  display: 'flex',
  alignItems: 'center',
  justifyContent: 'center',
  marginRight: '12px',
  flexShrink: 0,
};

const messageStyle: React.CSSProperties = {
  fontSize: '0.875rem',
  fontWeight: 500,
  color: 'var(--text-primary)',
  flexGrow: 1,
  lineHeight: 1.4,
  wordBreak: 'break-word',
};

const closeButtonStyle: React.CSSProperties = {
  background: 'none',
  border: 'none',
  color: 'var(--text-muted)',
  cursor: 'pointer',
  display: 'flex',
  alignItems: 'center',
  justifyContent: 'center',
  marginLeft: '12px',
  padding: '4px',
  borderRadius: '4px',
  transition: 'color 0.2s, background-color 0.2s',
  flexShrink: 0,
};

// Add slide-in animation styles directly to head if not present
if (typeof document !== 'undefined') {
  const style = document.createElement('style');
  style.innerHTML = `
    @keyframes toastSlideIn {
      from {
        opacity: 0;
        transform: translateY(20px) scale(0.95);
      }
      to {
        opacity: 1;
        transform: translateY(0) scale(1);
      }
    }
  `;
  document.head.appendChild(style);
}
