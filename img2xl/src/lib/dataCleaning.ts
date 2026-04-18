/**
 * Advanced OCR Data Cleaning Logic
 */

interface CleanerStats {
  corrections: number;
}

/**
 * Attempts to correct common OCR character confusion errors.
 * e.g., 'O' -> '0' when surrounded by numbers.
 */
function fixCharacterConfusion(val: string): string {
  // If the string is mostly numeric but has common OCR errors
  // Rule: If it looks like a number but has O, I, l, |, etc.
  
  // Replace O/o with 0 if it's in a numeric context
  // Regex to find O/o surrounded by digits or start/end of string
  let cleaned = val;
  
  // We apply these selectively if the string "should" be numeric
  const looksNumeric = /^[0-9OolI|.,$\s-]+$/.test(val) && /[0-9]/.test(val);
  
  if (looksNumeric) {
    cleaned = cleaned
      .replace(/O/g, '0')
      .replace(/o/g, '0')
      .replace(/l/g, '1')
      .replace(/I/g, '1')
      .replace(/\|/g, '1')
      .replace(/\s/g, ''); // Remove spaces in numbers
  }
  
  return cleaned;
}

/**
 * Normalizes inconsistent spacing.
 */
function normalizeSpacing(val: string): string {
  if (typeof val !== 'string') return val;
  return val.trim().replace(/\s+/g, ' ');
}

/**
 * Standardizes date formats to YYYY-MM-DD where possible.
 */
function standardizeDate(val: string): string {
  if (typeof val !== 'string') return val;
  
  // Try to match DD/MM/YYYY, MM/DD/YYYY, or similar
  // This is tricky without locale, but we can handle obvious patterns
  const datePattern = /^(\d{1,2})[./-](\d{1,2})[./-](\d{2,4})$/;
  const match = val.match(datePattern);
  
  if (match) {
    let [_, d1, d2, year] = match;
    if (year.length === 2) {
      year = parseInt(year) > 50 ? '19' + year : '20' + year;
    }
    
    // Heuristic: If d1 > 12, it's likely Day. If d2 > 12, it's likely Day.
    // Default to ISO-like if ambiguous
    return `${year}-${d1.padStart(2, '0')}-${d2.padStart(2, '0')}`;
  }
  
  return val;
}

/**
 * Main cleaning function for a single value.
 */
function cleanValue(val: any): any {
  if (val === null || val === undefined) return val;
  if (typeof val !== 'string') return val;
  
  let cleaned = val;
  
  // 1. Normalize Spacing
  cleaned = normalizeSpacing(cleaned);
  
  // 2. Fix OCR Char Confusion
  cleaned = fixCharacterConfusion(cleaned);
  
  // 3. Standardize Dates (if it looks like a date)
  if (cleaned.length >= 6 && /[\.\-\/]/.test(cleaned)) {
    cleaned = standardizeDate(cleaned);
  }
  
  return cleaned;
}

/**
 * Cleans an entire array of extracted rows.
 */
export function cleanExtractedData(data: any[]): any[] {
  if (!Array.isArray(data)) return data;
  
  return data.map(row => {
    const cleanedRow: any = {};
    for (const [key, value] of Object.entries(row)) {
      // Clean both keys (column names) and values
      const cleanedKey = normalizeSpacing(key);
      cleanedRow[cleanedKey] = cleanValue(value);
    }
    return cleanedRow;
  });
}
