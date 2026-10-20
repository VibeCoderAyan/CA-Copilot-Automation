import * as XLSX from 'xlsx';
import { parseAmount, parseDate, parseDateObj, isValidGstin } from './dataSanitizer';
import { createStyledWorksheet, exportStyledWorkbook } from './excelStyler';

export interface InvoiceRecord {
  id: string;
  rowIndex: number;
  vendorName: string;
  gstin: string;
  invoiceNumber: string;
  normalizedInvoiceNo: string;
  invoiceDate: string;
  dateTimestamp: number;
  taxableAmount: number;
  taxAmount: number;
  totalAmount: number;
  raw: Record<string, any>;
}

export type DuplicateMatchType =
  | 'exact'
  | 'normalized_no'
  | 'close_date_same_amt'
  | 'same_no_diff_amt'
  | 'same_no_amt_diff_vendor';

export type ReviewStatus = 'unresolved' | 'duplicate' | 'valid';

export interface DuplicateFlag {
  id: string;
  pairKey: string;
  matchType: DuplicateMatchType;
  severity: 'critical' | 'high' | 'medium';
  confidence: number;
  reasons: string[];
  recordA: InvoiceRecord;
  recordB: InvoiceRecord;
  status: ReviewStatus;
  userNotes?: string;
}

export interface DuplicateDetectionConfig {
  checkExact: boolean;
  checkNormalizedNo: boolean;
  checkCloseDateSameAmount: boolean;
  checkSameNoDiffAmount: boolean;
  checkSameNoDiffVendor: boolean;
  dateToleranceDays: number;
  amountTolerance: number;
}

export interface DuplicateDetectionResult {
  totalRecords: number;
  duplicatePairsCount: number;
  exactCount: number;
  probableCount: number;
  totalDuplicateExposure: number;
  unresolvedCount: number;
  confirmedDuplicateCount: number;
  markedValidCount: number;
  flags: DuplicateFlag[];
  records: InvoiceRecord[];
}

export interface InvoiceColumnMapping {
  vendorCol: string;
  gstinCol: string;
  invoiceNoCol: string;
  dateCol: string;
  taxableCol: string;
  taxCol: string;
  totalCol: string;
}

/**
 * Normalizes invoice number by stripping punctuation, spaces, and leading zeros
 * e.g., "INV/2026/0045" -> "inv202645"
 * e.g., "BIL-00089-A" -> "bil89a"
 */
export function normalizeInvoiceNumber(invNo: string): string {
  if (!invNo) return '';
  return invNo
    .toLowerCase()
    .replace(/[^a-z0-9]/g, '')
    .replace(/^0+/, '');
}

/**
 * Normalizes vendor name for fuzzy comparison
 */
export function normalizeVendorName(name: string): string {
  if (!name) return '';
  return name
    .toLowerCase()
    .replace(/\b(pvt|ltd|limited|private|llp|inc|corp|enterprises|solutions|services|industries)\b/g, '')
    .replace(/[^a-z0-9]/g, '')
    .trim();
}

import { detectColumnCategory } from './columnDetector';

/**
 * Auto-detects invoice headers from Excel/CSV using fuzzy keyword intelligence
 */
export function detectInvoiceColumns(headers: string[]): InvoiceColumnMapping {
  const mapping: InvoiceColumnMapping = {
    vendorCol: '',
    gstinCol: '',
    invoiceNoCol: '',
    dateCol: '',
    taxableCol: '',
    taxCol: '',
    totalCol: '',
  };

  // Phase 1: Use intelligent column category detector
  for (const h of headers) {
    const cat = detectColumnCategory(h);
    if (!mapping.vendorCol && cat === 'partyName') mapping.vendorCol = h;
    else if (!mapping.gstinCol && cat === 'gstin') mapping.gstinCol = h;
    else if (!mapping.invoiceNoCol && cat === 'voucherNo') mapping.invoiceNoCol = h;
    else if (!mapping.dateCol && cat === 'date') mapping.dateCol = h;
    else if (!mapping.taxableCol && cat === 'taxableAmount') mapping.taxableCol = h;
    else if (!mapping.taxCol && (cat === 'cgst' || cat === 'sgst' || cat === 'igst')) mapping.taxCol = h;
    else if (!mapping.totalCol && cat === 'totalAmount') mapping.totalCol = h;
  }

  // Phase 2: Targeted invoice keyword matching for unmapped columns
  for (const h of headers) {
    const clean = h.toLowerCase().trim();
    if (!mapping.vendorCol && /vendor|party|supplier|seller|trade.*name|biller/i.test(clean)) {
      mapping.vendorCol = h;
    } else if (!mapping.gstinCol && /gstin|gst.*no|supplier.*gst|tax.*id/i.test(clean)) {
      mapping.gstinCol = h;
    } else if (!mapping.invoiceNoCol && /inv.*no|invoice.*num|bill.*no|doc.*no|voucher.*no/i.test(clean)) {
      mapping.invoiceNoCol = h;
    } else if (!mapping.dateCol && /inv.*date|invoice.*dt|bill.*date|doc.*date|date/i.test(clean)) {
      mapping.dateCol = h;
    } else if (!mapping.taxableCol && /taxable|base.*amt|assessable|net.*amt/i.test(clean)) {
      mapping.taxableCol = h;
    } else if (!mapping.taxCol && /tax.*amt|total.*tax|gst.*amt|tax.*val/i.test(clean)) {
      mapping.taxCol = h;
    } else if (!mapping.totalCol && /total|invoice.*val|gross.*amt|bill.*amt|net.*payable/i.test(clean)) {
      mapping.totalCol = h;
    }
  }

  return mapping;
}

/**
 * Parses raw JSON rows into structured InvoiceRecord items
 */
export function parseInvoiceRows(
  rows: Record<string, any>[],
  mapping: InvoiceColumnMapping
): InvoiceRecord[] {
  return rows.map((r, idx) => {
    const vendorName = String(r[mapping.vendorCol] || '').trim();
    const gstin = String(r[mapping.gstinCol] || '').trim().toUpperCase();
    const invoiceNumber = String(r[mapping.invoiceNoCol] || '').trim();
    const rawDate = r[mapping.dateCol];
    const dObj = parseDateObj(rawDate);
    const dateTimestamp = dObj && !isNaN(dObj.getTime()) ? dObj.getTime() : 0;
    const invoiceDate = dObj ? parseDate(dObj, 'DD-MM-YYYY') : (rawDate ? String(rawDate).trim() : '');

    const taxableAmount = parseAmount(r[mapping.taxableCol]);
    const taxAmount = parseAmount(r[mapping.taxCol]);
    let totalAmount = parseAmount(r[mapping.totalCol]);

    // If total amount is missing but taxable + tax exist
    if (totalAmount === 0 && (taxableAmount > 0 || taxAmount > 0)) {
      totalAmount = taxableAmount + taxAmount;
    }

    return {
      id: `inv-${idx + 1}`,
      rowIndex: idx + 2, // Excel 1-based + 1 header row
      vendorName,
      gstin,
      invoiceNumber,
      normalizedInvoiceNo: normalizeInvoiceNumber(invoiceNumber),
      invoiceDate,
      dateTimestamp,
      taxableAmount,
      taxAmount,
      totalAmount,
      raw: r,
    };
  });
}

/**
 * Main Detection Algorithm: Compares records and identifies exact and possible duplicates
 */
export function detectDuplicateInvoices(
  records: InvoiceRecord[],
  config: DuplicateDetectionConfig,
  existingFlags: DuplicateFlag[] = []
): DuplicateDetectionResult {
  const flags: DuplicateFlag[] = [];
  const existingStatusMap = new Map<string, { status: ReviewStatus; notes?: string }>();
  existingFlags.forEach((f) => {
    existingStatusMap.set(f.pairKey, { status: f.status, notes: f.userNotes });
  });

  const n = records.length;
  const processedPairs = new Set<string>();

  for (let i = 0; i < n; i++) {
    const a = records[i];
    if (!a.invoiceNumber && a.totalAmount === 0) continue;

    for (let j = i + 1; j < n; j++) {
      const b = records[j];
      if (!b.invoiceNumber && b.totalAmount === 0) continue;

      const pairKey = a.rowIndex < b.rowIndex ? `${a.id}_${b.id}` : `${b.id}_${a.id}`;
      if (processedPairs.has(pairKey)) continue;

      const sameGstin = a.gstin && b.gstin && a.gstin === b.gstin;
      const sameVendorNormalized =
        normalizeVendorName(a.vendorName) === normalizeVendorName(b.vendorName) &&
        normalizeVendorName(a.vendorName).length > 2;
      const vendorMatch = sameGstin || sameVendorNormalized;

      const exactInvNo =
        a.invoiceNumber.toLowerCase() === b.invoiceNumber.toLowerCase() &&
        a.invoiceNumber.trim() !== '';
      const normInvNoMatch =
        !exactInvNo &&
        a.normalizedInvoiceNo &&
        b.normalizedInvoiceNo &&
        a.normalizedInvoiceNo === b.normalizedInvoiceNo;

      const amountDiff = Math.abs(a.totalAmount - b.totalAmount);
      const isAmountMatch = a.totalAmount > 0 && amountDiff <= config.amountTolerance;

      const dayDiff =
        a.dateTimestamp && b.dateTimestamp
          ? Math.abs(a.dateTimestamp - b.dateTimestamp) / (1000 * 60 * 60 * 24)
          : 9999;

      let matchType: DuplicateMatchType | null = null;
      let severity: 'critical' | 'high' | 'medium' = 'medium';
      let confidence = 0;
      const reasons: string[] = [];

      // 1. Exact Duplicate
      if (config.checkExact && vendorMatch && exactInvNo && isAmountMatch) {
        matchType = 'exact';
        severity = 'critical';
        confidence = 100;
        reasons.push('Identical Vendor/GSTIN, Invoice Number, and Total Amount.');
        if (dayDiff === 0) {
          reasons.push('Identical Invoice Date.');
        } else {
          reasons.push(`Invoices dated ${dayDiff} days apart.`);
        }
      }
      // 2. Normalized Invoice Number Match
      else if (config.checkNormalizedNo && vendorMatch && normInvNoMatch && isAmountMatch) {
        matchType = 'normalized_no';
        severity = 'high';
        confidence = 92;
        reasons.push(
          `Same Vendor & Amount with normalized invoice match ("${a.invoiceNumber}" vs "${b.invoiceNumber}").`
        );
        if (dayDiff <= config.dateToleranceDays) {
          reasons.push(`Dated within ${Math.round(dayDiff)} days.`);
        }
      }
      // 3. Same Vendor & Same Amount within Close Date
      else if (
        config.checkCloseDateSameAmount &&
        vendorMatch &&
        isAmountMatch &&
        dayDiff <= config.dateToleranceDays &&
        !exactInvNo &&
        !normInvNoMatch
      ) {
        matchType = 'close_date_same_amt';
        severity = 'high';
        confidence = 82;
        reasons.push(
          `Identical Vendor & Amount (₹${a.totalAmount.toLocaleString('en-IN')}) billed within ${Math.round(dayDiff)} days.`
        );
        reasons.push(`Different Invoice Numbers: "${a.invoiceNumber}" vs "${b.invoiceNumber}".`);
      }
      // 4. Same Invoice Number + Same Vendor but Different Amount
      else if (
        config.checkSameNoDiffAmount &&
        vendorMatch &&
        (exactInvNo || normInvNoMatch) &&
        !isAmountMatch
      ) {
        matchType = 'same_no_diff_amt';
        severity = 'medium';
        confidence = 74;
        reasons.push(
          `Same Vendor and Invoice Number ("${a.invoiceNumber}"), but differing amounts: ₹${a.totalAmount.toLocaleString('en-IN')} vs ₹${b.totalAmount.toLocaleString('en-IN')}.`
        );
        reasons.push('Potential revised bill, partial booking, or clerical amount error.');
      }
      // 5. Same Invoice Number + Same Amount but Different Vendor
      else if (
        config.checkSameNoDiffVendor &&
        !vendorMatch &&
        (exactInvNo || normInvNoMatch) &&
        isAmountMatch
      ) {
        matchType = 'same_no_amt_diff_vendor';
        severity = 'medium';
        confidence = 68;
        reasons.push(
          `Matching Invoice Number ("${a.invoiceNumber}") and Amount (₹${a.totalAmount.toLocaleString('en-IN')}) entered under different vendor names ("${a.vendorName}" vs "${b.vendorName}").`
        );
      }

      if (matchType) {
        processedPairs.add(pairKey);
        const saved = existingStatusMap.get(pairKey);

        flags.push({
          id: `dup-${flags.length + 1}`,
          pairKey,
          matchType,
          severity,
          confidence,
          reasons,
          recordA: a,
          recordB: b,
          status: saved ? saved.status : 'unresolved',
          userNotes: saved?.notes,
        });
      }
    }
  }

  // Calculate summary metrics
  const totalDuplicateExposure = flags.reduce((acc, f) => {
    if (f.status === 'valid') return acc;
    return acc + Math.min(f.recordA.totalAmount, f.recordB.totalAmount);
  }, 0);

  const exactCount = flags.filter((f) => f.matchType === 'exact').length;
  const probableCount = flags.length - exactCount;
  const unresolvedCount = flags.filter((f) => f.status === 'unresolved').length;
  const confirmedDuplicateCount = flags.filter((f) => f.status === 'duplicate').length;
  const markedValidCount = flags.filter((f) => f.status === 'valid').length;

  return {
    totalRecords: records.length,
    duplicatePairsCount: flags.length,
    exactCount,
    probableCount,
    totalDuplicateExposure,
    unresolvedCount,
    confirmedDuplicateCount,
    markedValidCount,
    flags,
    records,
  };
}

/**
 * Generates structured 3-sheet Excel workbook export for duplicate invoice analysis
 */
export function exportDuplicateAuditReport(
  result: DuplicateDetectionResult,
  clientName = 'Audit Client',
  period = 'FY 2026-27'
): XLSX.WorkBook {
  const wb = XLSX.utils.book_new();

  // Sheet 1: Executive Audit Summary
  const summaryData = [
    { Parameter: 'Client Entity', Value: clientName },
    { Parameter: 'Financial Period', Value: period },
    { Parameter: 'Total Invoices Scanned', Value: result.totalRecords },
    { Parameter: 'Potential Duplicate Pairs Detected', Value: result.duplicatePairsCount },
    { Parameter: 'Exact Duplicates (100% Match)', Value: result.exactCount },
    { Parameter: 'Probable / Heuristic Duplicates', Value: result.probableCount },
    { Parameter: 'Confirmed Duplicates Marked by Auditor', Value: result.confirmedDuplicateCount },
    { Parameter: 'Cleared / Marked as Valid', Value: result.markedValidCount },
    { Parameter: 'Pending Unresolved Review', Value: result.unresolvedCount },
    { Parameter: 'Estimated Duplicate Financial Exposure (₹)', Value: result.totalDuplicateExposure },
    { Parameter: 'Report Generation Timestamp', Value: new Date().toLocaleString('en-IN') },
    { Parameter: 'Rule Compliance Note', Value: 'Source register preserved. No entries deleted.' },
  ];
  const wsSummary = createStyledWorksheet(summaryData, {
    currencyColumns: ['Value'],
    columnWidths: { Parameter: 36, Value: 34 },
  });
  XLSX.utils.book_append_sheet(wb, wsSummary, 'Audit_Summary');

  // Sheet 2: Flagged Duplicate Invoices (Side-by-Side Comparison)
  const duplicateRows = result.flags.map((f, idx) => ({
    'Flag #': idx + 1,
    'Match Type': f.matchType.replace(/_/g, ' ').toUpperCase(),
    'Confidence %': `${f.confidence}%`,
    'Review Status': f.status.toUpperCase(),
    'Primary Row #': f.recordA.rowIndex,
    'Primary Vendor': f.recordA.vendorName,
    'Primary GSTIN': f.recordA.gstin,
    'Primary Inv #': f.recordA.invoiceNumber,
    'Primary Date': f.recordA.invoiceDate,
    'Primary Amount (₹)': f.recordA.totalAmount,
    'Comparison Row #': f.recordB.rowIndex,
    'Comparison Vendor': f.recordB.vendorName,
    'Comparison GSTIN': f.recordB.gstin,
    'Comparison Inv #': f.recordB.invoiceNumber,
    'Comparison Date': f.recordB.invoiceDate,
    'Comparison Amount (₹)': f.recordB.totalAmount,
    'Variance (₹)': Math.abs(f.recordA.totalAmount - f.recordB.totalAmount),
    'Audit Reasons': f.reasons.join(' | '),
    'Auditor Notes': f.userNotes || '',
  }));
  const wsDuplicates = createStyledWorksheet(duplicateRows, {
    currencyColumns: ['Primary Amount (₹)', 'Comparison Amount (₹)', 'Variance (₹)'],
  });
  XLSX.utils.book_append_sheet(wb, wsDuplicates, 'Flagged_Duplicates');

  // Sheet 3: Full Invoices Register with Audit Status
  const flagMap = new Map<string, string>();
  result.flags.forEach((f) => {
    const desc = `${f.matchType.toUpperCase()} (${f.status})`;
    flagMap.set(f.recordA.id, desc);
    flagMap.set(f.recordB.id, desc);
  });

  const fullRegisterRows = result.records.map((r) => ({
    'Row #': r.rowIndex,
    'Vendor Name': r.vendorName,
    'GSTIN': r.gstin,
    'Invoice Number': r.invoiceNumber,
    'Invoice Date': r.invoiceDate,
    'Taxable Value (₹)': r.taxableAmount,
    'Tax Amount (₹)': r.taxAmount,
    'Total Amount (₹)': r.totalAmount,
    'Duplicate Status': flagMap.get(r.id) || 'Clean / Single Entry',
  }));
  const wsRegister = createStyledWorksheet(fullRegisterRows, {
    currencyColumns: ['Taxable Value (₹)', 'Tax Amount (₹)', 'Total Amount (₹)'],
  });
  XLSX.utils.book_append_sheet(wb, wsRegister, 'All_Invoices_Register');

  return wb;
}

/**
 * High-quality realistic sample dataset with intentional duplicate test cases
 */
export function getSamplePurchaseRegisterRows(): Record<string, any>[] {
  return [
    {
      'Row ID': 1,
      'Vendor Name': 'Larsen & Toubro Ltd',
      'Vendor GSTIN': '27AAACL0149G1ZN',
      'Invoice Number': 'LT/PUN/2026/0412',
      'Invoice Date': '04-09-2026',
      'Taxable Amount': 450000.0,
      'GST Amount': 81000.0,
      'Total Amount': 531000.0,
      'Remarks': 'Fabrication work - Phase 2',
    },
    {
      'Row ID': 2,
      'Vendor Name': 'Larsen & Toubro Ltd',
      'Vendor GSTIN': '27AAACL0149G1ZN',
      'Invoice Number': 'LT/PUN/2026/0412', // EXACT DUPLICATE
      'Invoice Date': '04-09-2026',
      'Taxable Amount': 450000.0,
      'GST Amount': 81000.0,
      'Total Amount': 531000.0,
      'Remarks': 'Duplicate voucher booked by junior accountant',
    },
    {
      'Row ID': 3,
      'Vendor Name': 'Tata Consultancy Services Ltd',
      'Vendor GSTIN': '27AAACT2727Q1ZW',
      'Invoice Number': 'TCS-INV-00891',
      'Invoice Date': '08-09-2026',
      'Taxable Amount': 185000.0,
      'GST Amount': 33300.0,
      'Total Amount': 218300.0,
      'Remarks': 'Cloud hosting charges',
    },
    {
      'Row ID': 4,
      'Vendor Name': 'Tata Consultancy Services Ltd',
      'Vendor GSTIN': '27AAACT2727Q1ZW',
      'Invoice Number': 'TCS/INV/891', // NORMALIZED NUMBER MATCH
      'Invoice Date': '10-09-2026',
      'Taxable Amount': 185000.0,
      'GST Amount': 33300.0,
      'Total Amount': 218300.0,
      'Remarks': 'Entered from scanned copy with slash format',
    },
    {
      'Row ID': 5,
      'Vendor Name': 'Reliance Jio Infocomm Ltd',
      'Vendor GSTIN': '24AAACR5055K1Z1',
      'Invoice Number': 'JIO/CORP/98214',
      'Invoice Date': '12-09-2026',
      'Taxable Amount': 75000.0,
      'GST Amount': 13500.0,
      'Total Amount': 88500.0,
      'Remarks': 'Monthly enterprise connectivity',
    },
    {
      'Row ID': 6,
      'Vendor Name': 'Reliance Jio Infocomm Ltd',
      'Vendor GSTIN': '24AAACR5055K1Z1',
      'Invoice Number': 'JIO/CORP/98230', // SAME AMOUNT & VENDOR WITHIN 2 DAYS
      'Invoice Date': '14-09-2026',
      'Taxable Amount': 75000.0,
      'GST Amount': 13500.0,
      'Total Amount': 88500.0,
      'Remarks': 'Duplicate subscription billed across circles',
    },
    {
      'Row ID': 7,
      'Vendor Name': 'Godrej & Boyce Mfg Co Ltd',
      'Vendor GSTIN': '27AAACG0549H1ZT',
      'Invoice Number': 'GB/FUR/26-27/102',
      'Invoice Date': '15-09-2026',
      'Taxable Amount': 240000.0,
      'GST Amount': 43200.0,
      'Total Amount': 283200.0,
      'Remarks': 'Executive office workstation setup',
    },
    {
      'Row ID': 8,
      'Vendor Name': 'Godrej & Boyce Mfg Co Ltd',
      'Vendor GSTIN': '27AAACG0549H1ZT',
      'Invoice Number': 'GB/FUR/26-27/102', // SAME INVOICE NO, DIFFERENT AMOUNT
      'Invoice Date': '18-09-2026',
      'Taxable Amount': 260000.0,
      'GST Amount': 46800.0,
      'Total Amount': 306800.0,
      'Remarks': 'Revised debit note / amendment not adjusted',
    },
    {
      'Row ID': 9,
      'Vendor Name': 'Infosys Limited',
      'Vendor GSTIN': '29AAACI4390G1ZT',
      'Invoice Number': 'INF/BLR/2026/901',
      'Invoice Date': '20-09-2026',
      'Taxable Amount': 320000.0,
      'GST Amount': 57600.0,
      'Total Amount': 377600.0,
      'Remarks': 'ERP maintenance support Q2',
    },
    {
      'Row ID': 10,
      'Vendor Name': 'Bharat Petroleum Corp Ltd',
      'Vendor GSTIN': '27AAACB2902M1ZT',
      'Invoice Number': 'BPCL/FLEET/4412',
      'Invoice Date': '22-09-2026',
      'Taxable Amount': 98000.0,
      'GST Amount': 17640.0,
      'Total Amount': 115640.0,
      'Remarks': 'Fleet diesel bulk purchase',
    },
    {
      'Row ID': 11,
      'Vendor Name': 'Blue Dart Express Ltd',
      'Vendor GSTIN': '27AAACB3137H1ZU',
      'Invoice Number': 'BD/COURIER/88210',
      'Invoice Date': '24-09-2026',
      'Taxable Amount': 14200.0,
      'GST Amount': 2556.0,
      'Total Amount': 16756.0,
      'Remarks': 'Audit file courier dispatches',
    },
    {
      'Row ID': 12,
      'Vendor Name': 'Kirloskar Brothers Ltd',
      'Vendor GSTIN': '27AAACK3192G1ZU',
      'Invoice Number': 'KBL/PUMP/2026/089',
      'Invoice Date': '28-09-2026',
      'Taxable Amount': 510000.0,
      'GST Amount': 91800.0,
      'Total Amount': 601800.0,
      'Remarks': 'Submersible pump motor installation',
    },
  ];
}
