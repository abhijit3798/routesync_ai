type LogLevel = 'DEBUG' | 'INFO' | 'WARN' | 'ERROR';

class Logger {
  private isDevelopment: boolean;

  constructor() {
    // Check vite environment mode
    this.isDevelopment = import.meta.env?.DEV ?? true;
  }

  private formatMessage(level: LogLevel, message: string): string {
    const timestamp = new Date().toISOString();
    return `[RouteSync AI] [${level}] [${timestamp}] ${message}`;
  }

  debug(message: string, ...args: unknown[]): void {
    if (this.isDevelopment) {
      console.debug(this.formatMessage('DEBUG', message), ...args);
    }
  }

  info(message: string, ...args: unknown[]): void {
    if (this.isDevelopment) {
      console.info(this.formatMessage('INFO', message), ...args);
    }
  }

  warn(message: string, ...args: unknown[]): void {
    console.warn(this.formatMessage('WARN', message), ...args);
  }

  error(message: string, error?: unknown, ...args: unknown[]): void {
    const formatted = this.formatMessage('ERROR', message);
    if (error instanceof Error) {
      console.error(formatted, error.message, error.stack, ...args);
    } else {
      console.error(formatted, error, ...args);
    }
  }
}

export const logger = new Logger();
