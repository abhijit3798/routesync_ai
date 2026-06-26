import React, { Component } from 'react';
import type { ErrorInfo, ReactNode } from 'react';
import { logger } from '../utils/logger';
import { AlertOctagon } from 'lucide-react';

interface Props {
  children?: ReactNode;
}

interface State {
  hasError: boolean;
  error: Error | null;
}

export class ErrorBoundary extends Component<Props, State> {
  public state: State = {
    hasError: false,
    error: null,
  };

  public static getDerivedStateFromError(error: Error): State {
    return { hasError: true, error };
  }

  public componentDidCatch(error: Error, errorInfo: ErrorInfo) {
    logger.error('Uncaught component boundary error:', error, errorInfo);
  }

  private handleReset = () => {
    this.setState({ hasError: false, error: null });
    window.location.reload();
  };

  public render() {
    if (this.state.hasError) {
      return (
        <div style={containerStyle}>
          <div className="glass-panel animate-slide-up" style={cardStyle}>
            <AlertOctagon size={48} color="var(--error)" style={{ marginBottom: '16px' }} />
            <h2 style={{ marginBottom: '8px', fontFamily: 'var(--font-heading)' }}>
              Something went wrong
            </h2>
            <p style={{ color: 'var(--text-secondary)', marginBottom: '20px', fontSize: '0.95rem' }}>
              An unexpected error occurred in the RouteSync AI dashboard.
            </p>
            {this.state.error && (
              <pre style={debugStyle}>
                <code>{this.state.error.toString()}</code>
              </pre>
            )}
            <button onClick={this.handleReset} style={buttonStyle}>
              Reload Application
            </button>
          </div>
        </div>
      );
    }

    return this.props.children;
  }
}

const containerStyle: React.CSSProperties = {
  display: 'flex',
  alignItems: 'center',
  justifyContent: 'center',
  minHeight: '100vh',
  padding: '24px',
  backgroundColor: 'var(--bg-app)',
};

const cardStyle: React.CSSProperties = {
  maxWidth: '480px',
  width: '100%',
  padding: '32px',
  textAlign: 'center',
  display: 'flex',
  flexDirection: 'column',
  alignItems: 'center',
  background: 'var(--bg-surface)',
};

const debugStyle: React.CSSProperties = {
  width: '100%',
  background: 'rgba(0, 0, 0, 0.05)',
  padding: '12px',
  borderRadius: '6px',
  textAlign: 'left',
  overflowX: 'auto',
  fontSize: '0.8rem',
  marginBottom: '24px',
  maxHeight: '150px',
};

const buttonStyle: React.CSSProperties = {
  backgroundColor: 'var(--primary)',
  color: 'white',
  border: 'none',
  padding: '10px 20px',
  borderRadius: '8px',
  cursor: 'pointer',
  fontSize: '0.95rem',
  fontWeight: 600,
  transition: 'background-color 0.2s',
};
