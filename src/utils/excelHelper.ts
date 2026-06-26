import { logger } from './logger';

export interface ParseMetadata {
  sheetNames: string[];
  rowCount: number;
  columnCount: number;
  headers: string[];
  previewRows: Record<string, unknown>[];
}

export class ExcelHelper {
  private static workerInstance: Worker | null = null;
  private static pendingPromises = new Map<string, { resolve: Function; reject: Function }>();
  private static progressCallback: ((percentage: number) => void) | null = null;
  private static messageIdCounter = 0;

  private static getWorker(): Worker {
    if (!this.workerInstance) {
      logger.info('Initializing persistent Web Worker instance.');
      this.workerInstance = new Worker(
        new URL('../workers/excel.worker.ts', import.meta.url),
        { type: 'module' }
      );

      this.workerInstance.onmessage = (event) => {
        const { id, type, payload } = event.data;

        if (type === 'COMPARE_PROGRESS') {
          if (this.progressCallback) {
            this.progressCallback(payload);
          }
          return;
        }

        const promise = this.pendingPromises.get(id);
        if (!promise) return;

        this.pendingPromises.delete(id);

        if (type === 'SUCCESS') {
          promise.resolve(payload);
        } else if (type === 'ERROR') {
          promise.reject(payload instanceof Error ? payload : new Error(payload));
        }
      };

      this.workerInstance.onerror = (err) => {
        logger.error('Persistent worker global runtime error', err);
      };
    }
    return this.workerInstance;
  }

  private static sendRequest(type: string, payload: any, transferables?: Transferable[]): Promise<any> {
    return new Promise((resolve, reject) => {
      const id = `msg_${++this.messageIdCounter}_${Date.now()}`;
      this.pendingPromises.set(id, { resolve, reject });
      this.getWorker().postMessage({ id, type, payload }, transferables || []);
    });
  }

  /**
   * Parse Excel spreadsheet in background and cache the dataset in worker memory.
   */
  static parseFile(file: File, target: 'master' | 'route'): Promise<ParseMetadata> {
    return new Promise((resolve, reject) => {
      logger.info(`Requesting parse for ${target} file: ${file.name} (${file.size} bytes)`);
      const reader = new FileReader();

      reader.onload = async (e) => {
        const buffer = e.target?.result as ArrayBuffer;
        if (!buffer) {
          reject(new Error('Failed to read file into ArrayBuffer'));
          return;
        }

        try {
          const result = await this.sendRequest('PARSE_FILE', { buffer, target }, [buffer]);
          resolve(result);
        } catch (err) {
          reject(err);
        }
      };

      reader.onerror = (err) => {
        logger.error('FileReader error during parsing', err);
        reject(err);
      };

      reader.readAsArrayBuffer(file);
    });
  }

  /**
   * Run normalization audits on cached master dataset.
   */
  static processData(mappings: Record<string, string>): Promise<{
    auditSummary: {
      totalRecords: number;
      duplicateRefCodes: number;
      blankRefCodes: number;
      blankNames: number;
      blankLandmarks: number;
      errorsCount: number;
    };
  }> {
    return this.sendRequest('PROCESS_DATA', { mappings });
  }

  /**
   * Run background matching engine and landmark validation.
   */
  static compareFiles(
    masterMappings: Record<string, string>,
    routeMappings: Record<string, string>,
    settings: {
      autoGenerateStops: boolean;
      autoGenerateTimes: boolean;
      tripStartTime: string;
      tripEndTime: string;
    },
    onProgress?: (percentage: number) => void
  ): Promise<{
    metrics: {
      totalRecords: number;
      matchedRecords: number;
      missingRefCodes: number;
      landmarkChanged: number;
      landmarkCorrect: number;
      landmarkMissing: number;
      duplicateNames: number;
      errorsFound: number;
      averageConfidence: number;
    };
    routeSheetBuffer: ArrayBuffer;
    errorReportBuffer: ArrayBuffer;
    comparisonLogs: string[];
  }> {
    this.progressCallback = onProgress || null;
    return this.sendRequest('COMPARE_FILES', { masterMappings, routeMappings, settings });
  }

  /**
   * Clear cached datasets from background worker memory.
   */
  static clearCache(target: 'master' | 'route' | 'both'): Promise<void> {
    return this.sendRequest('CLEAR_CACHE', { target });
  }
}
