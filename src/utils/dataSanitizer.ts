import * as XLSX from 'xlsx';

// Robust data sanitization, date parsing, and accounting currency formatting for CA tools

/**
 * Parses financial amount from any string, number, accounting parenthesis, or dirty format
 * Example: "(45,000.00)" -> -45000
 * Example: "₹ 1,50,000.00" -> 150000
 * Example: "  2,300.50 DR " -> 2300.50
 * Example: "  2,300.50 CR " -> -2300.50
 */
export function parseAmount(val: any): number {
  if (val === null || val === undefined) return 0;
  if (typeof val === 'number') {
    return isNaN(val) ? 0 : val;
  }

  const str = val.toString().trim();
  if (!str || str === '-' || str.toLowerCase() === 'nil' || str.toLowerCase() === 'na' || str.toLowerCase() === 'n/a') {
    return 0;
  }

  // Check accounting parenthesis: (1,234.56) -> negative
  let isNegative = false;
  if (str.startsWith('(') && str.endsWith(')')) {
    isNegative = true;
  } else if (str.startsWith('-')) {
    isNegative = true;
  } else if (/cr\b/i.test(str)) {
    // Credit in some ledgers can be negative or positive depending on context
    isNegative = false;
  }

  // Strip non-numeric chars except dot
  const cleaned = str.replace(/[^\d.]/g, '');
  if (!cleaned) return 0;

  const num = parseFloat(cleaned);
  if (isNaN(num)) return 0;

  return isNegative ? -Math.abs(num) : Math.abs(num);
}

/**
/**
 * Robust date parser for accounting spreadsheets and forms.
 * Accurately parses Excel serial numbers, DD/MM/YYYY, DD-MM-YYYY, DD.MM.YYYY,
 * DD-MMM-YYYY, YYYY-MM-DD, ISO timestamps, and JS Date objects into a valid Date.
 * Returns null if input cannot be parsed.
 */
export function parseDateObj(val: any): Date | null {
  if (val === null || val === undefined || val === '') return null;

  if (val instanceof Date) {
    return isNaN(val.getTime()) ? null : val;
  }

  // 1. Check if Excel numeric date serial (e.g. 44820 = Sep 2022)
  if (typeof val === 'number' || (typeof val === 'string' && /^\d+(\.\d+)?$/.test(val.trim()) && Number(val) > 10000 && Number(val) < 80000)) {
    const serial = Number(val);
    // Excel base date: Dec 30, 1899 (accounting for 1900 leap year bug)
    const msSince1970 = Math.round((serial - 25569) * 86400 * 1000);
    const tempUtc = new Date(msSince1970);
    if (!isNaN(tempUtc.getTime())) {
      // Create local date from UTC parts to eliminate local timezone offset shift
      return new Date(tempUtc.getUTCFullYear(), tempUtc.getUTCMonth(), tempUtc.getUTCDate());
    }
    return null;
  }

  if (typeof val === 'string') {
    const s = val.trim();
    if (!s || s === '-' || s.toLowerCase() === 'n/a' || s.toLowerCase() === 'null') return null;

    // Strip trailing time portion if present (e.g., "15-09-2026 14:32:00" -> "15-09-2026")
    const dateOnlyMatch = s.match(/^(\d{1,4}[\/\-\.]\S+?)(?:\s+\d{1,2}:\d{2}(?::\d{2})?(?:\s*[ap]m)?)?$/i);
    const dateStr = dateOnlyMatch ? dateOnlyMatch[1] : s;

    // Match DD/MM/YYYY or DD-MM-YYYY or DD.MM.YYYY
    const dmyMatch = dateStr.match(/^(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{2,4})$/);
    if (dmyMatch) {
      let day = parseInt(dmyMatch[1], 10);
      let month = parseInt(dmyMatch[2], 10) - 1;
      let year = parseInt(dmyMatch[3], 10);
      if (year < 100) year += year > 50 ? 1900 : 2000;
      if (month >= 0 && month <= 11 && day >= 1 && day <= 31) {
        const d = new Date(year, month, day);
        if (!isNaN(d.getTime())) return d;
      }
    }

    // Match DD-MMM-YYYY or DD-MMM-YY (e.g. 15-Sep-2026, 15-Sep-26, 04-OCTOBER-2026)
    const dMonthMatch = dateStr.match(/^(\d{1,2})[\/\-\.\s]([A-Za-z]{3,9})[\/\-\.\s](\d{2,4})$/);
    const monthsMap: Record<string, number> = {
      jan: 0, january: 0,
      feb: 1, february: 1,
      mar: 2, march: 2,
      apr: 3, april: 3,
      may: 4,
      jun: 5, june: 5,
      jul: 6, july: 6,
      aug: 7, august: 7,
      sep: 8, sept: 8, september: 8,
      oct: 9, october: 9,
      nov: 10, november: 10,
      dec: 11, december: 11,
    };

    if (dMonthMatch) {
      const day = parseInt(dMonthMatch[1], 10);
      const mKey = dMonthMatch[2].toLowerCase();
      const mIdx = monthsMap[mKey] !== undefined ? monthsMap[mKey] : -1;
      let year = parseInt(dMonthMatch[3], 10);
      if (year < 100) year += year > 50 ? 1900 : 2000;

      if (mIdx !== -1 && day >= 1 && day <= 31) {
        const d = new Date(year, mIdx, day);
        if (!isNaN(d.getTime())) return d;
      }
    }

    // Match YYYY-MM-DD or YYYY/MM/DD or YYYY.MM.DD
    const ymdMatch = dateStr.match(/^(\d{4})[\/\-\.](\d{1,2})[\/\-\.](\d{1,2})/);
    if (ymdMatch) {
      const year = parseInt(ymdMatch[1], 10);
      const month = parseInt(ymdMatch[2], 10) - 1;
      const day = parseInt(ymdMatch[3], 10);
      if (month >= 0 && month <= 11 && day >= 1 && day <= 31) {
        const d = new Date(year, month, day);
        if (!isNaN(d.getTime())) return d;
      }
    }

    // Native Date.parse fallback
    const parsed = Date.parse(s);
    if (!isNaN(parsed)) {
      const d = new Date(parsed);
      if (!isNaN(d.getTime())) return d;
    }
  }

  return null;
}

/**
 * Parses diverse date inputs into a standardized string format.
 */
export function parseDate(val: any, outputFormat: 'DD-MM-YYYY' | 'YYYY-MM-DD' | 'DD/MM/YYYY' = 'DD-MM-YYYY'): string {
  const d = parseDateObj(val);
  if (!d) {
    return val !== null && val !== undefined ? String(val).trim() : '';
  }

  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');

  if (outputFormat === 'YYYY-MM-DD') {
    return `${y}-${m}-${day}`;
  }
  if (outputFormat === 'DD/MM/YYYY') {
    return `${day}/${m}/${y}`;
  }
  return `${day}-${m}-${y}`;
}

/**
 * Standard Indian currency formatter (Lakhs and Crores)
 * Example: 1500000 -> ₹ 15,00,000.00
 */
export function formatINR(val: number, includeDecimals = true): string {
  if (isNaN(val) || val === null || val === undefined) return '₹ 0.00';
  const isNegative = val < 0;
  const absVal = Math.abs(val);

  const parts = absVal.toFixed(includeDecimals ? 2 : 0).split('.');
  let lastThree = parts[0].substring(parts[0].length - 3);
  const otherNumbers = parts[0].substring(0, parts[0].length - 3);
  if (otherNumbers !== '') {
    lastThree = ',' + lastThree;
  }
  const res = otherNumbers.replace(/\B(?=(\d{2})+(?!\d))/g, ',') + lastThree;
  const decimalStr = includeDecimals && parts[1] ? `.${parts[1]}` : '';

  return `${isNegative ? '-' : ''}₹ ${res}${decimalStr}`;
}

/**
 * Validates Indian 15-character GSTIN
 */
export function isValidGstin(gstin: string): boolean {
  if (!gstin) return false;
  const clean = gstin.trim().toUpperCase();
  const regex = /^[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z]{1}[1-9A-Z]{1}Z[0-9A-Z]{1}$/;
  return regex.test(clean);
}

/**
 * Normalizes an invoice or reference number by removing non-alphanumeric chars
 * Useful for matching "INV/2026/001" with "INV-2026-001" or "INV2026001"
 */
export function normalizeRefNo(ref: string): string {
  if (!ref) return '';
  return ref.toString().trim().toUpperCase().replace(/[^A-Z0-9]/g, '');
}

/**
 * Validates Indian 10-character Permanent Account Number (PAN)
 * Format: 5 letters, 4 digits, 1 letter (e.g. ABCDE1234F)
 */
export function isValidPan(pan: string): boolean {
  if (!pan) return false;
  const clean = pan.trim().toUpperCase();
  const regex = /^[A-Z]{5}[0-9]{4}[A-Z]{1}$/;
  return regex.test(clean);
}

export interface SmartTableResult {
  headers: string[];
  rows: Record<string, any>[];
  headerRowIndex: number;
}

/**
 * Intelligently locates the table header row and extracts data rows from an Excel worksheet.
 * Solves the common problem where accounting software (Tally, SAP, Marg, Busy) prints
 * 1 to 5 company title, period, or metadata rows before the actual column headers.
 */
export function extractSmartTableFromSheet(ws: XLSX.WorkSheet): SmartTableResult {
  if (!ws || !ws['!ref']) {
    return { headers: [], rows: [], headerRowIndex: 0 };
  }

  // 1. Convert entire worksheet to raw 2D array
  const rawAOA: any[][] = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
  if (!rawAOA || rawAOA.length === 0) {
    return { headers: [], rows: [], headerRowIndex: 0 };
  }

  // Keywords that strongly indicate an accounting column header
  const headerKeywordsRegex =
    /\b(date|dt|period|inv|bill|vch|voucher|doc|party|vendor|supplier|customer|client|name|particular|description|narration|gstin|gst|pan|code|account|acc|head|amt|amount|taxable|cgst|sgst|igst|tax|total|gross|net|debit|credit|dr|cr|balance|bal|rate|qty|quantity|units|price|serial|s\.no|sl\.no|sr\.no|status|ref|reference|chq|cheque|utr|rem|remark|remarks|hsn|sac)\b/i;

  let bestHeaderRowIndex = 0;
  let maxKeywordScore = -1;

  // Scan the first 25 rows to identify the genuine column header row
  const scanLimit = Math.min(rawAOA.length, 25);
  for (let r = 0; r < scanLimit; r++) {
    const row = rawAOA[r];
    if (!Array.isArray(row) || row.length === 0) continue;

    // Filter to non-empty string cells
    const textCells = row
      .map((c) => (c !== null && c !== undefined ? String(c).trim() : ''))
      .filter((s) => s.length > 0);

    if (textCells.length < 2) continue; // Single title cell like "Apex Pvt Ltd" is not a table header

    let score = 0;
    for (const cell of textCells) {
      if (headerKeywordsRegex.test(cell)) {
        score += 2;
      } else if (cell.length > 0 && cell.length <= 35 && !/^[\d\.,\-\/]+$/.test(cell)) {
        // Short alphanumeric label
        score += 0.5;
      }
    }

    // Must have at least 2 distinct recognizable column headers
    if (score > maxKeywordScore && score >= 2) {
      maxKeywordScore = score;
      bestHeaderRowIndex = r;
    }
  }

  // Extract headers from the best header row
  const headerRow = rawAOA[bestHeaderRowIndex] || [];
  const headers: string[] = [];
  const seenHeaderCounts: Record<string, number> = {};

  for (let colIdx = 0; colIdx < headerRow.length; colIdx++) {
    let rawColName = String(headerRow[colIdx] ?? '').trim();
    if (!rawColName) {
      rawColName = `Column_${colIdx + 1}`;
    }
    // Disambiguate duplicate column names (e.g. two "Amount" or "Date" columns)
    if (seenHeaderCounts[rawColName]) {
      seenHeaderCounts[rawColName]++;
      headers.push(`${rawColName}_${seenHeaderCounts[rawColName]}`);
    } else {
      seenHeaderCounts[rawColName] = 1;
      headers.push(rawColName);
    }
  }

  // Extract structured rows following the header row
  const structuredRows: Record<string, any>[] = [];
  for (let r = bestHeaderRowIndex + 1; r < rawAOA.length; r++) {
    const row = rawAOA[r];
    if (!Array.isArray(row)) continue;

    // Check if row is completely empty or just whitespace
    const hasData = row.some((c) => c !== null && c !== undefined && String(c).trim() !== '');
    if (!hasData) continue;

    const rowObj: Record<string, any> = {};
    let nonBlankInRow = 0;

    for (let c = 0; c < headers.length; c++) {
      const val = row[c] !== undefined ? row[c] : '';
      rowObj[headers[c]] = val;
      if (val !== null && val !== undefined && String(val).trim() !== '') {
        nonBlankInRow++;
      }
    }

    // Skip trailing totals / footer summary rows if they only have 1 single label and no numbers
    if (nonBlankInRow > 0) {
      structuredRows.push(rowObj);
    }
  }

  return {
    headers,
    rows: structuredRows,
    headerRowIndex: bestHeaderRowIndex,
  };
}
