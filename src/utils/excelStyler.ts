import * as XLSX from 'xlsx';

export interface SheetConfig {
  name: string;
  data: Record<string, any>[];
  headers?: string[];
  currencyColumns?: string[];
  dateColumns?: string[];
  columnWidths?: Record<string, number>;
}

/**
 * Parses string representations of numbers, accounting parentheses, or currency symbols
 */
function parseNumericAmount(val: any): number | null {
  if (typeof val === 'number') {
    return isNaN(val) || !isFinite(val) ? 0 : val;
  }
  if (typeof val !== 'string') return null;

  const s = val.trim();
  if (!s || s === '-' || s.toLowerCase() === 'nil' || s.toLowerCase() === 'na' || s.toLowerCase() === 'n/a') {
    return 0;
  }

  const isParenNeg = s.startsWith('(') && s.endsWith(')');
  const cleanStr = s.replace(/[₹$,\s()]/g, '');
  if (/^-?\d+(\.\d+)?$/.test(cleanStr)) {
    const num = parseFloat(cleanStr);
    if (!isNaN(num)) {
      return isParenNeg ? -Math.abs(num) : num;
    }
  }
  return null;
}

/**
 * Formats dates into DD-MM-YYYY
 */
function formatDateValue(val: any): string {
  if (val instanceof Date && !isNaN(val.getTime())) {
    const dd = String(val.getDate()).padStart(2, '0');
    const mm = String(val.getMonth() + 1).padStart(2, '0');
    const yyyy = val.getFullYear();
    return `${dd}-${mm}-${yyyy}`;
  }
  if (typeof val === 'string') {
    const isoMatch = val.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (isoMatch) {
      return `${isoMatch[3]}-${isoMatch[2]}-${isoMatch[1]}`;
    }
    return val.trim();
  }
  return String(val || '');
}

/**
 * Automatically calculates optimal column widths based on maximum string lengths
 * to prevent text clipping and the infamous '###' display in Excel.
 */
function calculateColumnWidths(
  data: Record<string, any>[],
  headers: string[],
  customWidths?: Record<string, number>
): XLSX.ColInfo[] {
  if (!headers || headers.length === 0) return [];

  const widths: XLSX.ColInfo[] = [];
  const sampleSize = Math.min(data.length, 500);

  for (const header of headers) {
    let maxLen = header.length;

    for (let i = 0; i < sampleSize; i++) {
      const val = data[i]?.[header];
      if (val !== undefined && val !== null) {
        let strVal = String(val);
        if (typeof val === 'number') {
          strVal = val.toLocaleString('en-IN', { minimumFractionDigits: 2 });
        }
        // Handle multiline strings by taking the longest line
        const lines = strVal.split(/\r?\n/);
        for (const line of lines) {
          if (line.trim().length > maxLen) {
            maxLen = line.trim().length;
          }
        }
      }
    }

    const custom = customWidths?.[header];
    // Safety margin (+5) prevents Excel autofilter dropdown arrow from obscuring cell text or causing '###'
    const computedWidth = custom || Math.min(Math.max(maxLen + 5, 14), 70);
    widths.push({ wch: computedWidth });
  }

  return widths;
}

/**
 * Formats a SheetJS worksheet with optimal widths, number formats, freeze pane, and auto-filters.
 * Enforces deterministic column ordering and robust cell value sanitization.
 */
export function createStyledWorksheet(
  data: Record<string, any>[],
  config: {
    headers?: string[];
    currencyColumns?: string[];
    dateColumns?: string[];
    columnWidths?: Record<string, number>;
    freezeHeader?: boolean;
  } = {}
): XLSX.WorkSheet {
  if (!data || data.length === 0) {
    return XLSX.utils.json_to_sheet([{ Note: 'No data records found' }]);
  }

  // 1. Determine unified, deterministic headers
  let headers: string[] = [];
  if (config.headers && config.headers.length > 0) {
    headers = [...config.headers];
  } else {
    const headerSet = new Set<string>();
    // First row keys
    for (const k of Object.keys(data[0] || {})) {
      headerSet.add(k);
    }
    // Collect subsequent keys
    for (let i = 1; i < data.length; i++) {
      for (const k of Object.keys(data[i])) {
        headerSet.add(k);
      }
    }
    headers = Array.from(headerSet);
  }

  if (headers.length === 0) {
    headers = ['Note'];
  }

  // 2. Prepare Currency & Date identification
  const currencyHeaders = new Set<string>(
    (config.currencyColumns || []).map((c) => c.toLowerCase())
  );
  const dateHeaders = new Set<string>(
    (config.dateColumns || []).map((d) => d.toLowerCase())
  );

  // Auto-detect financial & date columns
  headers.forEach((h) => {
    const lower = h.toLowerCase();
    if (
      lower.includes('amount') ||
      lower.includes('amt') ||
      lower.includes('debit') ||
      lower.includes('credit') ||
      lower.includes('balance') ||
      lower.includes('bal') ||
      lower.includes('taxable') ||
      lower.includes('cgst') ||
      lower.includes('sgst') ||
      lower.includes('igst') ||
      lower.includes('cess') ||
      lower.includes('total') ||
      lower.includes('turnover') ||
      lower.includes('shortfall') ||
      lower.includes('variance') ||
      lower.includes('diff') ||
      lower.includes('interest') ||
      lower.includes('disallowance') ||
      lower.includes('addition') ||
      lower.includes('withdrawal') ||
      lower.includes('deposit') ||
      lower.includes('fee') ||
      lower.includes('price') ||
      lower.includes('cost') ||
      lower.includes('value') ||
      lower.includes('val') ||
      lower.includes('salary') ||
      lower.includes('basic') ||
      lower.includes('hra') ||
      lower.includes('gross') ||
      lower.includes('net') ||
      lower.includes('deduction') ||
      lower.includes('wdv') ||
      lower.includes('depreciation') ||
      lower.includes('salvage') ||
      lower.includes('exposure') ||
      lower.includes('fine') ||
      lower.includes('penalty') ||
      lower.includes('rate') ||
      lower.includes('₹') ||
      lower.includes('$')
    ) {
      currencyHeaders.add(lower);
    }

    if (
      lower.includes('date') ||
      lower.includes('dated') ||
      lower.includes('period') ||
      lower.includes('timestamp') ||
      lower.includes('due')
    ) {
      dateHeaders.add(lower);
    }
  });

  // 3. Build sanitized records with strict column alignment
  const sanitizedData: Record<string, any>[] = [];
  for (const row of data) {
    const cleanRow: Record<string, any> = {};
    for (const h of headers) {
      const val = row[h];
      if (val === null || val === undefined) {
        cleanRow[h] = '';
      } else if (typeof val === 'number') {
        cleanRow[h] = isNaN(val) || !isFinite(val) ? 0 : val;
      } else if (typeof val === 'boolean') {
        cleanRow[h] = val ? 'YES' : 'NO';
      } else if (val instanceof Date) {
        cleanRow[h] = formatDateValue(val);
      } else if (Array.isArray(val)) {
        cleanRow[h] = val.join(', ');
      } else if (typeof val === 'object') {
        cleanRow[h] = val.name || val.title || val.message || JSON.stringify(val);
      } else if (typeof val === 'string') {
        const lowerH = h.toLowerCase();
        // Dynamically auto-convert any accounting strings (e.g. ₹ 1,50,000.00 or (45,000.00)) to true Excel numbers
        const isLikelyCurrency = currencyHeaders.has(lowerH) || /[₹$,]|^\s*\([0-9,.]+\)\s*$/.test(val);
        const parsed = isLikelyCurrency ? parseNumericAmount(val) : null;

        if (parsed !== null) {
          cleanRow[h] = parsed;
          currencyHeaders.add(lowerH);
        } else if (dateHeaders.has(lowerH)) {
          cleanRow[h] = formatDateValue(val);
        } else {
          cleanRow[h] = val.trim();
        }
      } else {
        cleanRow[h] = String(val);
      }
    }
    sanitizedData.push(cleanRow);
  }

  // 4. Create worksheet with explicit header order
  const ws = XLSX.utils.json_to_sheet(sanitizedData, { header: headers });

  // 5. Column Widths (Prevents ### in Excel)
  ws['!cols'] = calculateColumnWidths(sanitizedData, headers, config.columnWidths);

  // 6. Header Row Height
  ws['!rows'] = [{ hpt: 26 }];

  // 7. Freeze Header Pane (keeps header row locked while scrolling)
  if (config.freezeHeader !== false) {
    ws['!views'] = [{ state: 'frozen', ySplit: 1, activeCell: 'A2' }];
  }

  // 8. Auto-Filter on Header Row
  if (ws['!ref']) {
    ws['!autofilter'] = { ref: ws['!ref'] };
  }

  // 9. Apply number and date formatting to individual cells
  if (ws['!ref']) {
    const range = XLSX.utils.decode_range(ws['!ref']);
    for (let colIdx = range.s.c; colIdx <= range.e.c; colIdx++) {
      const colHeader = headers[colIdx];
      if (!colHeader) continue;

      const lowerH = colHeader.toLowerCase();
      const isCurrency = currencyHeaders.has(lowerH);
      const isDate = dateHeaders.has(lowerH);
      const isPercentage = lowerH.includes('percent') || lowerH.includes('ratio') || lowerH.includes('%');
      const isCount =
        lowerH.includes('s.no') ||
        lowerH.includes('sno') ||
        lowerH.includes('sl no') ||
        lowerH.includes('count') ||
        lowerH.includes('age') ||
        lowerH.includes('days') ||
        lowerH.includes('items');

      for (let rowIdx = range.s.r + 1; rowIdx <= range.e.r; rowIdx++) {
        const cellRef = XLSX.utils.encode_cell({ r: rowIdx, c: colIdx });
        const cell = ws[cellRef];
        if (!cell) continue;

        if (isCurrency && typeof cell.v === 'number') {
          cell.t = 'n';
          cell.z = '#,##0.00;[Red](#,##0.00);"-"';
        } else if (isCount && typeof cell.v === 'number') {
          cell.t = 'n';
          cell.z = '#,##0';
        } else if (isPercentage && typeof cell.v === 'number') {
          cell.t = 'n';
          cell.z = cell.v <= 1 ? '0.00%' : '0.00"%"';
        } else if (isDate && typeof cell.v === 'string') {
          cell.z = 'dd-mm-yyyy';
        }
      }
    }
  }

  return ws;
}

/**
 * Creates a complete styled workbook with multiple sheets and triggers Excel download
 */
export function exportStyledWorkbook(sheets: SheetConfig[], fileName: string): void {
  const wb = XLSX.utils.book_new();

  for (const sheet of sheets) {
    const ws = createStyledWorksheet(sheet.data, {
      headers: sheet.headers,
      currencyColumns: sheet.currencyColumns,
      dateColumns: sheet.dateColumns,
      columnWidths: sheet.columnWidths,
    });
    // SheetJS sheet name limit is 31 characters, remove forbidden characters
    const cleanSheetName = sheet.name.replace(/[*?:/\\\[\]]/g, '').trim().slice(0, 31) || 'Sheet';
    XLSX.utils.book_append_sheet(wb, ws, cleanSheetName);
  }

  const outFileName = fileName.endsWith('.xlsx') ? fileName : `${fileName}.xlsx`;
  XLSX.writeFile(wb, outFileName);
}
