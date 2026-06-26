import { logger } from './logger';

export interface SystemColumn {
  key: string;
  label: string;
  aliases: string[];
  pattern?: RegExp;
}

export const SYSTEM_COLUMNS: SystemColumn[] = [
  {
    key: 'CommuterRefCode',
    label: 'Commuter Reference Code',
    aliases: ['ref code', 'reference code', 'commuter ref', 'employee code', 'student code', 'registration number', 'id', 'emp id', 'reg no', 'commuter id'],
    pattern: /^[a-z0-9-]{3,15}$/i,
  },
  {
    key: 'CommuterName',
    label: 'Commuter Name',
    aliases: ['commuter name', 'student name', 'employee name', 'child name', 'passenger name', 'name', 'commuter', 'employee', 'student'],
  },
  {
    key: 'Landmark',
    label: 'Landmark / Stop Location',
    aliases: ['pickup landmark', 'stop', 'pickup stop', 'pickup location', 'landmark', 'location', 'assigned stop', 'drop stop'],
  },
  {
    key: 'Route',
    label: 'Route Identifier',
    aliases: ['route number', 'vehicle route', 'route no', 'route', 'assigned route', 'route id'],
  },
  {
    key: 'RouteType',
    label: 'Route Type / Shift',
    aliases: ['pickup', 'drop', 'shift', 'route type', 'type', 'trip type', 'direction'],
  },
  {
    key: 'POCName',
    label: 'POC Name / Parent',
    aliases: ['parent name', 'father name', 'guardian', 'mother', 'poc name', 'poc', 'emergency contact name'],
  },
  {
    key: 'POCMobile',
    label: 'POC Contact Mobile',
    aliases: ['mobile', 'contact number', 'phone', 'poc mobile', 'poc phone', 'contact', 'telephone', 'mobile no', 'cell phone'],
    pattern: /^\+?[0-9\s-]{7,15}$/,
  },
  {
    key: 'POCEmail',
    label: 'POC Email',
    aliases: ['email', 'email address', 'poc email', 'parent email', 'e-mail'],
    pattern: /^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$/,
  },
];

/**
 * Normalizes a string by stripping spaces and symbols, making it case-insensitive
 */
export const normalizeHeader = (header: string): string => {
  return header.toLowerCase().replace(/[^a-z0-9]/g, '');
};

/**
 * Autodetect system column mappings for a given list of sheet columns and sample data
 */
export const detectColumnMappings = (
  sheetColumns: string[],
  sampleRows: Record<string, unknown>[]
): Record<string, string> => {
  const mappings: Record<string, string> = {};
  logger.info(`Starting intelligent column detection across headers: ${sheetColumns.join(', ')}`);

  SYSTEM_COLUMNS.forEach((sysCol) => {
    // 1. Try matching aliases
    let matchedColumn = sheetColumns.find((sheetCol) => {
      const normalizedSheetCol = normalizeHeader(sheetCol);
      return sysCol.aliases.some((alias) => {
        const normalizedAlias = normalizeHeader(alias);
        // Direct match or contains
        return normalizedSheetCol === normalizedAlias || normalizedSheetCol.includes(normalizedAlias);
      });
    });

    // 2. If alias match fails and pattern exists, test sample data
    if (!matchedColumn && sysCol.pattern && sampleRows.length > 0) {
      for (const sheetCol of sheetColumns) {
        // Test first few rows of data
        const samples = sampleRows.slice(0, 10).map(r => String(r[sheetCol] || '').trim()).filter(Boolean);
        if (samples.length > 0) {
          const matchCount = samples.filter(val => sysCol.pattern!.test(val)).length;
          // If > 70% match data pattern, assign mapping
          if (matchCount / samples.length >= 0.7) {
            matchedColumn = sheetCol;
            logger.info(`Pattern match succeeded for system column ${sysCol.key} -> sheet column ${sheetCol}`);
            break;
          }
        }
      }
    }

    if (matchedColumn) {
      mappings[sysCol.key] = matchedColumn;
    } else {
      mappings[sysCol.key] = ''; // unmapped initially
    }
  });

  logger.info('Detected Column Mappings result:', mappings);
  return mappings;
};
