/**
 * Normalizes commuter or parent name: trims spaces, replaces duplicate spaces, and capitalizes words.
 */
export const normalizeName = (name: string): string => {
  if (!name) return '';
  return name
    .trim()
    .replace(/\s+/g, ' ')
    .toLowerCase()
    .split(' ')
    .map((word) => word.charAt(0).toUpperCase() + word.slice(1))
    .join(' ');
};

/**
 * Normalizes landmark/stop locations: trims spaces, replaces duplicate spaces, and capitalizes words.
 */
export const normalizeLandmark = (landmark: string): string => {
  if (!landmark) return '';
  return landmark
    .trim()
    .replace(/\s+/g, ' ')
    .toLowerCase()
    .split(' ')
    .map((word) => word.charAt(0).toUpperCase() + word.slice(1))
    .join(' ');
};

/**
 * Normalizes mobile number strings: strips non-numeric symbols but preserves leading plus indicators.
 */
export const normalizeMobile = (mobile: string): string => {
  if (!mobile) return '';
  return mobile.trim().replace(/[^\d+]/g, '');
};

/**
 * Normalizes emails: trims spaces and converts characters to lowercase.
 */
export const normalizeEmail = (email: string): string => {
  if (!email) return '';
  return email.trim().toLowerCase();
};
