import * as XLSX from 'xlsx';
import { parseAmount, parseDate, parseDateObj, normalizeRefNo } from './dataSanitizer';
import { detectColumnCategory } from './columnDetector';
import { createStyledWorksheet } from './excelStyler';

export interface GstVoucher {
  id: string;
  source: 'books' | 'gstr2b';
  supplierGstin: string;
  supplierName: string;
  invoiceNumber: string;
  invoiceDate: string;
  rawDate?: any;
  taxableAmount: number;
  cgst: number;
  sgst: number;
  igst: number;
  cess?: number;
  totalTax: number;
  totalInvoiceAmount: number;
  isDuplicate?: boolean;
}

export interface GstComparisonRow {
  id: string;
  supplierGstin: string;
  supplierName: string;
  invoiceNumber: string;
  invoiceDate: string;
  booksTaxable: number;
  booksTotalTax: number;
  gstr2bTaxable: number;
  gstr2bTotalTax: number;
  diffTaxable: number;
  diffAmount: number;
  matchScore: number;
  status: 'Exact Match' | 'Probable Match' | 'Tax Mismatch' | 'Missing in 2B' | 'Missing in Books' | 'Duplicate Record';
  itcEligibility: 'Eligible ITC' | 'Ineligible / Blocked' | 'Vendor Follow-up Required' | 'Unclaimed Entry in Books' | 'Duplicate ITC Risk';
  actionNote: string;
  resolutionStatus?: 'Pending Review' | 'Approved for 3B' | 'Query Sent to Vendor' | 'Debit Note Required' | 'Entry Booked';
  userNotes?: string;
  sourceReference?: string;
}

export interface GstReconciliationSummary {
  totalBooksTax: number;
  total2bTax: number;
  eligibleItc: number;
  atRiskItc: number;
  unclaimedItc: number;
  matchedCount: number;
  probableCount: number;
  mismatchCount: number;
  missingIn2bCount: number;
  missingInBooksCount: number;
  duplicateCount: number;
}

export interface GstReconciliationConfig {
  taxTolerance: number; // e.g. ₹2.0
  taxableTolerance: number; // e.g. ₹10.0
  dateToleranceDays: number; // e.g. 30
  fuzzyMatching: boolean;
}

export const DEFAULT_GST_CONFIG: GstReconciliationConfig = {
  taxTolerance: 2.0,
  taxableTolerance: 10.0,
  dateToleranceDays: 30,
  fuzzyMatching: true,
};

/**
 * Normalizes invoice serial numbers for matching by stripping punctuation, spaces, and leading zeros
 */
export function normalizeInvoiceNumber(invNo: string): string {
  if (!invNo) return '';
  const cleaned = invNo.toString().trim().toUpperCase();
  const parts = cleaned.split(/[^A-Z0-9]+/).filter(Boolean);
  if (parts.length === 0) {
    return cleaned.replace(/[^A-Z0-9]/g, '').replace(/^0+(?=\d)/, '');
  }
  return parts
    .map(p => {
      // Strip leading zeros from numeric segment or alphanumeric transition
      return p.replace(/^0+(?=\d)/, '').replace(/([A-Z]+)0+(?=\d)/g, '$1');
    })
    .join('');
}

/**
 * Parses raw Purchase Register rows into standardized GstVouchers
 */
export function parsePurchaseRegisterRows(rawRows: Record<string, any>[]): GstVoucher[] {
  if (!rawRows || rawRows.length === 0) return [];
  const sample = rawRows[0];
  const colMap: Record<string, string> = {};

  for (const key of Object.keys(sample)) {
    const cat = detectColumnCategory(key);
    if (cat) colMap[cat] = key;
  }

  const seenInvoices = new Map<string, number>();

  return rawRows.map((row, idx) => {
    const gstin = (colMap.gstin ? row[colMap.gstin] : row.gstin || row.GSTIN || '').toString().trim().toUpperCase();
    const name = (colMap.partyName ? row[colMap.partyName] : row.supplierName || row.partyName || row['Party Name'] || '').toString().trim();
    const invNo = (colMap.voucherNo ? row[colMap.voucherNo] : row.invoiceNumber || row.invoiceNo || row['Invoice No'] || '').toString().trim();
    const rawDate = colMap.date ? row[colMap.date] : row.invoiceDate || row.Date || '';
    const invDate = parseDate(rawDate, 'DD-MM-YYYY');

    const taxable = parseAmount(colMap.taxableAmount ? row[colMap.taxableAmount] : row.taxableAmount || row.Taxable || 0);
    const cgst = parseAmount(colMap.cgst ? row[colMap.cgst] : row.cgst || row.CGST || 0);
    const sgst = parseAmount(colMap.sgst ? row[colMap.sgst] : row.sgst || row.SGST || 0);
    const igst = parseAmount(colMap.igst ? row[colMap.igst] : row.igst || row.IGST || 0);
    const cess = parseAmount(row.cess || row.Cess || 0);
    const totalTax = cgst + sgst + igst + cess > 0 ? cgst + sgst + igst + cess : parseAmount(colMap.totalAmount ? row[colMap.totalAmount] : 0) - taxable;
    const totalAmt = parseAmount(colMap.totalAmount ? row[colMap.totalAmount] : row.totalAmount || row.Total || 0) || (taxable + totalTax);

    // Duplicate check
    const dedupKey = `${gstin}|${normalizeInvoiceNumber(invNo)}|${taxable.toFixed(2)}`;
    const prevCount = seenInvoices.get(dedupKey) || 0;
    seenInvoices.set(dedupKey, prevCount + 1);

    return {
      id: `pr-${idx + 1}`,
      source: 'books',
      supplierGstin: gstin,
      supplierName: name || 'Supplier in Books',
      invoiceNumber: invNo,
      invoiceDate: invDate,
      rawDate,
      taxableAmount: taxable,
      cgst,
      sgst,
      igst,
      cess,
      totalTax: Math.max(0, totalTax),
      totalInvoiceAmount: totalAmt,
      isDuplicate: prevCount > 0,
    };
  });
}

/**
 * Parses raw GSTR-2B rows (or GSTN portal export JSON/Excel) into standardized GstVouchers
 */
export function parseGstr2bRows(rawInput: any): GstVoucher[] {
  // Check if GSTN JSON format with b2b array
  if (rawInput && typeof rawInput === 'object' && !Array.isArray(rawInput) && rawInput.b2b) {
    const vouchers: GstVoucher[] = [];
    let count = 1;
    for (const b2bItem of rawInput.b2b) {
      const ctin = (b2bItem.ctin || '').toUpperCase();
      const cname = b2bItem.cname || b2bItem.trade_name || 'Vendor';
      for (const inv of b2bItem.inv || []) {
        const inum = inv.inum || '';
        const idt = parseDate(inv.idt || '', 'DD-MM-YYYY');
        const val = parseAmount(inv.val || 0);
        let igst = 0;
        let cgst = 0;
        let sgst = 0;
        let cess = 0;
        let txval = 0;

        for (const itm of inv.items || []) {
          const det = itm.itm_det || itm;
          txval += parseAmount(det.txval || 0);
          igst += parseAmount(det.iamt || 0);
          cgst += parseAmount(det.camt || 0);
          sgst += parseAmount(det.samt || 0);
          cess += parseAmount(det.csamt || 0);
        }

        vouchers.push({
          id: `2b-${count++}`,
          source: 'gstr2b',
          supplierGstin: ctin,
          supplierName: cname,
          invoiceNumber: inum,
          invoiceDate: idt,
          rawDate: inv.idt,
          taxableAmount: txval,
          cgst,
          sgst,
          igst,
          cess,
          totalTax: igst + cgst + sgst + cess,
          totalInvoiceAmount: val || (txval + igst + cgst + sgst + cess),
        });
      }
    }
    return vouchers;
  }

  // Otherwise standard Excel / CSV rows array
  const rawRows = Array.isArray(rawInput) ? rawInput : [];
  if (rawRows.length === 0) return [];
  const sample = rawRows[0];
  const colMap: Record<string, string> = {};

  for (const key of Object.keys(sample)) {
    const cat = detectColumnCategory(key);
    if (cat) colMap[cat] = key;
  }

  return rawRows.map((row, idx) => {
    const gstin = (colMap.gstin ? row[colMap.gstin] : row.gstin || row['GSTIN of supplier'] || row.ctin || '').toString().trim().toUpperCase();
    const name = (colMap.partyName ? row[colMap.partyName] : row.supplierName || row['Trade/Legal name'] || row.cname || '').toString().trim();
    const invNo = (colMap.voucherNo ? row[colMap.voucherNo] : row.invoiceNumber || row['Invoice number'] || row.inum || '').toString().trim();
    const rawDate = colMap.date ? row[colMap.date] : row.invoiceDate || row['Invoice date'] || row.idt || '';
    const invDate = parseDate(rawDate, 'DD-MM-YYYY');

    const taxable = parseAmount(colMap.taxableAmount ? row[colMap.taxableAmount] : row.taxableAmount || row['Taxable Value (₹)'] || row.txval || 0);
    const cgst = parseAmount(colMap.cgst ? row[colMap.cgst] : row.cgst || row['Central Tax (₹)'] || row.camt || 0);
    const sgst = parseAmount(colMap.sgst ? row[colMap.sgst] : row.sgst || row['State/UT Tax (₹)'] || row.samt || 0);
    const igst = parseAmount(colMap.igst ? row[colMap.igst] : row.igst || row['Integrated Tax (₹)'] || row.iamt || 0);
    const cess = parseAmount(row['Cess (₹)'] || row.cess || 0);
    const totalTax = cgst + sgst + igst + cess > 0 ? cgst + sgst + igst + cess : parseAmount(colMap.totalAmount ? row[colMap.totalAmount] : 0) - taxable;
    const totalAmt = parseAmount(colMap.totalAmount ? row[colMap.totalAmount] : row.totalAmount || row['Invoice Value (₹)'] || row.val || 0) || (taxable + totalTax);

    return {
      id: `2b-${idx + 1}`,
      source: 'gstr2b',
      supplierGstin: gstin,
      supplierName: name || 'Supplier in 2B',
      invoiceNumber: invNo,
      invoiceDate: invDate,
      rawDate,
      taxableAmount: taxable,
      cgst,
      sgst,
      igst,
      cess,
      totalTax: Math.max(0, totalTax),
      totalInvoiceAmount: totalAmt,
    };
  });
}

/**
 * Reconciles Purchase Register against GSTR-2B with configurable tolerances
 */
export function runGstReconciliation(
  booksVouchers: GstVoucher[],
  gstr2bVouchers: GstVoucher[],
  config: GstReconciliationConfig = DEFAULT_GST_CONFIG
): {
  comparisonRecords: GstComparisonRow[];
  summary: GstReconciliationSummary;
} {
  const comparisonRecords: GstComparisonRow[] = [];
  const matched2bIds = new Set<string>();

  // Pass 1: Exact Match (Exact GSTIN, Normalized Invoice Number, Tax & Taxable within tolerance)
  for (const b of booksVouchers) {
    if (b.isDuplicate) {
      comparisonRecords.push({
        id: `comp-dup-${b.id}`,
        supplierGstin: b.supplierGstin,
        supplierName: b.supplierName,
        invoiceNumber: b.invoiceNumber,
        invoiceDate: b.invoiceDate,
        booksTaxable: b.taxableAmount,
        booksTotalTax: b.totalTax,
        gstr2bTaxable: 0,
        gstr2bTotalTax: 0,
        diffTaxable: b.taxableAmount,
        diffAmount: b.totalTax,
        matchScore: 0,
        status: 'Duplicate Record',
        itcEligibility: 'Duplicate ITC Risk',
        actionNote: 'Duplicate purchase entry detected in books. High risk of double ITC claim.',
        resolutionStatus: 'Debit Note Required',
      });
      continue;
    }

    const normBInv = normalizeInvoiceNumber(b.invoiceNumber);

    const exactMatch = gstr2bVouchers.find((g) => {
      if (matched2bIds.has(g.id)) return false;
      const gstinMatch = b.supplierGstin && g.supplierGstin ? b.supplierGstin === g.supplierGstin : true;
      const invMatch = normalizeInvoiceNumber(g.invoiceNumber) === normBInv;
      return gstinMatch && invMatch;
    });

    if (exactMatch) {
      matched2bIds.add(exactMatch.id);
      const taxDiff = Math.abs(b.totalTax - exactMatch.totalTax);
      const taxableDiff = Math.abs(b.taxableAmount - exactMatch.taxableAmount);

      if (taxDiff <= config.taxTolerance && taxableDiff <= config.taxableTolerance) {
        comparisonRecords.push({
          id: `comp-${b.id}-${exactMatch.id}`,
          supplierGstin: b.supplierGstin || exactMatch.supplierGstin,
          supplierName: b.supplierName || exactMatch.supplierName,
          invoiceNumber: b.invoiceNumber,
          invoiceDate: b.invoiceDate || exactMatch.invoiceDate,
          booksTaxable: b.taxableAmount,
          booksTotalTax: b.totalTax,
          gstr2bTaxable: exactMatch.taxableAmount,
          gstr2bTotalTax: exactMatch.totalTax,
          diffTaxable: taxableDiff,
          diffAmount: taxDiff,
          matchScore: 100,
          status: 'Exact Match',
          itcEligibility: 'Eligible ITC',
          actionNote: '100% Validated ITC under Section 16(2)(aa). Safe to claim in GSTR-3B.',
          resolutionStatus: 'Approved for 3B',
        });
      } else {
        comparisonRecords.push({
          id: `comp-${b.id}-${exactMatch.id}`,
          supplierGstin: b.supplierGstin || exactMatch.supplierGstin,
          supplierName: b.supplierName || exactMatch.supplierName,
          invoiceNumber: b.invoiceNumber,
          invoiceDate: b.invoiceDate || exactMatch.invoiceDate,
          booksTaxable: b.taxableAmount,
          booksTotalTax: b.totalTax,
          gstr2bTaxable: exactMatch.taxableAmount,
          gstr2bTotalTax: exactMatch.totalTax,
          diffTaxable: taxableDiff,
          diffAmount: taxDiff,
          matchScore: 85,
          status: 'Tax Mismatch',
          itcEligibility: 'Ineligible / Blocked',
          actionNote: `Tax variance of ₹${taxDiff.toFixed(2)}. Claim lower portal amount or demand credit/debit note.`,
          resolutionStatus: 'Debit Note Required',
        });
      }
      continue;
    }

    // Pass 2: Fuzzy / Probable Match (Same GSTIN + same Taxable Amount within ₹1, slight invoice number shift)
    if (config.fuzzyMatching) {
      const probableMatch = gstr2bVouchers.find((g) => {
        if (matched2bIds.has(g.id)) return false;
        const gstinMatch = b.supplierGstin && g.supplierGstin && b.supplierGstin === g.supplierGstin;
        const taxMatch = Math.abs(b.totalTax - g.totalTax) <= config.taxTolerance;
        const taxableMatch = Math.abs(b.taxableAmount - g.taxableAmount) <= config.taxableTolerance;
        return gstinMatch && taxMatch && taxableMatch;
      });

      if (probableMatch) {
        matched2bIds.add(probableMatch.id);
        const taxDiff = Math.abs(b.totalTax - probableMatch.totalTax);
        const taxableDiff = Math.abs(b.taxableAmount - probableMatch.taxableAmount);

        comparisonRecords.push({
          id: `comp-prob-${b.id}-${probableMatch.id}`,
          supplierGstin: b.supplierGstin,
          supplierName: b.supplierName || probableMatch.supplierName,
          invoiceNumber: `${b.invoiceNumber} (2B: ${probableMatch.invoiceNumber})`,
          invoiceDate: b.invoiceDate || probableMatch.invoiceDate,
          booksTaxable: b.taxableAmount,
          booksTotalTax: b.totalTax,
          gstr2bTaxable: probableMatch.taxableAmount,
          gstr2bTotalTax: probableMatch.totalTax,
          diffTaxable: taxableDiff,
          diffAmount: taxDiff,
          matchScore: 90,
          status: 'Probable Match',
          itcEligibility: 'Eligible ITC',
          actionNote: `Invoice numbers differ slightly (${b.invoiceNumber} vs ${probableMatch.invoiceNumber}). Amounts and GSTIN align.`,
          resolutionStatus: 'Pending Review',
        });
        continue;
      }
    }

    // Pass 3: Invoices missing in GSTR-2B
    comparisonRecords.push({
      id: `comp-missing2b-${b.id}`,
      supplierGstin: b.supplierGstin,
      supplierName: b.supplierName,
      invoiceNumber: b.invoiceNumber,
      invoiceDate: b.invoiceDate,
      booksTaxable: b.taxableAmount,
      booksTotalTax: b.totalTax,
      gstr2bTaxable: 0,
      gstr2bTotalTax: 0,
      diffTaxable: b.taxableAmount,
      diffAmount: b.totalTax,
      matchScore: 0,
      status: 'Missing in 2B',
      itcEligibility: 'Vendor Follow-up Required',
      actionNote: 'Invoice missing in GSTR-2B. Vendor has not filed GSTR-1. Ineligible until filed under Sec 16(2)(aa).',
      resolutionStatus: 'Query Sent to Vendor',
    });
  }

  // Pass 4: Invoices present in GSTR-2B but missing in Books
  for (const g of gstr2bVouchers) {
    if (!matched2bIds.has(g.id)) {
      comparisonRecords.push({
        id: `comp-missingBooks-${g.id}`,
        supplierGstin: g.supplierGstin,
        supplierName: g.supplierName,
        invoiceNumber: g.invoiceNumber,
        invoiceDate: g.invoiceDate,
        booksTaxable: 0,
        booksTotalTax: 0,
        gstr2bTaxable: g.taxableAmount,
        gstr2bTotalTax: g.totalTax,
        diffTaxable: g.taxableAmount,
        diffAmount: g.totalTax,
        matchScore: 0,
        status: 'Missing in Books',
        itcEligibility: 'Unclaimed Entry in Books',
        actionNote: 'Appears on GST portal but missing in accounting books. Obtain invoice and book entry.',
        resolutionStatus: 'Pending Review',
      });
    }
  }

  const totalBooksTax = comparisonRecords.reduce((acc, r) => acc + r.booksTotalTax, 0);
  const total2bTax = comparisonRecords.reduce((acc, r) => acc + r.gstr2bTotalTax, 0);
  const eligibleItc = comparisonRecords
    .filter((r) => r.status === 'Exact Match' || r.status === 'Probable Match')
    .reduce((acc, r) => acc + r.booksTotalTax, 0);
  const atRiskItc = comparisonRecords
    .filter((r) => r.status === 'Missing in 2B' || r.status === 'Tax Mismatch')
    .reduce((acc, r) => acc + r.diffAmount, 0);
  const unclaimedItc = comparisonRecords
    .filter((r) => r.status === 'Missing in Books')
    .reduce((acc, r) => acc + r.gstr2bTotalTax, 0);

  const matchedCount = comparisonRecords.filter((r) => r.status === 'Exact Match').length;
  const probableCount = comparisonRecords.filter((r) => r.status === 'Probable Match').length;
  const mismatchCount = comparisonRecords.filter((r) => r.status === 'Tax Mismatch').length;
  const missingIn2bCount = comparisonRecords.filter((r) => r.status === 'Missing in 2B').length;
  const missingInBooksCount = comparisonRecords.filter((r) => r.status === 'Missing in Books').length;
  const duplicateCount = comparisonRecords.filter((r) => r.status === 'Duplicate Record').length;

  return {
    comparisonRecords,
    summary: {
      totalBooksTax,
      total2bTax,
      eligibleItc,
      atRiskItc,
      unclaimedItc,
      matchedCount,
      probableCount,
      mismatchCount,
      missingIn2bCount,
      missingInBooksCount,
      duplicateCount,
    },
  };
}

/**
 * Generates CSV export of GST Reconciliation
 */
export function exportGstReconciliationCsv(records: GstComparisonRow[]): string {
  const headers = [
    'S.No',
    'Supplier GSTIN',
    'Supplier Name',
    'Invoice Number',
    'Invoice Date',
    'Books Taxable (₹)',
    'Books Tax (₹)',
    'GSTR-2B Taxable (₹)',
    'GSTR-2B Tax (₹)',
    'Tax Variance (₹)',
    'Reconciliation Status',
    'ITC Eligibility Tag',
    'Audit Action Note',
    'Resolution Status',
  ];

  const escapeCsv = (val: any) => `"${String(val ?? '').replace(/"/g, '""')}"`;

  const rows = records.map((r, idx) => [
    idx + 1,
    escapeCsv(r.supplierGstin),
    escapeCsv(r.supplierName),
    escapeCsv(r.invoiceNumber),
    escapeCsv(r.invoiceDate),
    r.booksTaxable.toFixed(2),
    r.booksTotalTax.toFixed(2),
    r.gstr2bTaxable.toFixed(2),
    r.gstr2bTotalTax.toFixed(2),
    r.diffAmount.toFixed(2),
    escapeCsv(r.status),
    escapeCsv(r.itcEligibility),
    escapeCsv(r.actionNote),
    escapeCsv(r.resolutionStatus || 'Pending Review'),
  ]);

  return [headers.join(','), ...rows.map((r) => r.join(','))].join('\r\n');
}

/**
 * Generates structured 6-sheet GST Reconciliation Excel workbook
 */
export function createGstReconciliationWorkbook(
  records: GstComparisonRow[],
  summary: GstReconciliationSummary,
  sourcePrRaw?: Record<string, any>[],
  source2bRaw?: Record<string, any>[]
): XLSX.WorkBook {
  const wb = XLSX.utils.book_new();

  // Sheet 1: Executive Summary & Legal Disclaimer
  const summarySheet = [
    { Parameter: 'Total Books Tax (Purchase Register)', 'Amount (₹)': summary.totalBooksTax },
    { Parameter: 'Total GSTR-2B Tax (GST Portal)', 'Amount (₹)': summary.total2bTax },
    { Parameter: '100% Eligible Input Tax Credit (ITC)', 'Amount (₹)': summary.eligibleItc },
    { Parameter: 'At-Risk / Ineligible ITC (Follow-up Required)', 'Amount (₹)': summary.atRiskItc },
    { Parameter: 'Unclaimed Available ITC (Missing in Books)', 'Amount (₹)': summary.unclaimedItc },
    { Parameter: 'Exact Matched Invoices Count', 'Amount (₹)': summary.matchedCount },
    { Parameter: 'Probable Matched Invoices Count', 'Amount (₹)': summary.probableCount },
    { Parameter: 'Tax Mismatched Invoices Count', 'Amount (₹)': summary.mismatchCount },
    { Parameter: 'Missing on Portal Count (Not in GSTR-2B)', 'Amount (₹)': summary.missingIn2bCount },
    { Parameter: 'Missing in Accounting Books Count', 'Amount (₹)': summary.missingInBooksCount },
    { Parameter: 'Duplicate Invoices Identified', 'Amount (₹)': summary.duplicateCount },
    { Parameter: 'Section 16(2)(aa) Legal Notice', 'Amount (₹)': 'Verification of GSTR-2B does not legally guarantee claim eligibility. Assessee must ensure physical possession of tax invoice, receipt of goods, and actual payment of tax.' },
  ];
  const wsSummary = createStyledWorksheet(summarySheet, {
    currencyColumns: ['Amount (₹)'],
    columnWidths: { Parameter: 45, 'Amount (₹)': 45 },
  });
  XLSX.utils.book_append_sheet(wb, wsSummary, 'Executive_Summary');

  // Sheet 2: Matched Invoices
  const matched = records.filter((r) => r.status === 'Exact Match' || r.status === 'Probable Match');
  const matchedRows = matched.map((r, idx) => ({
    'S.No': idx + 1,
    'Supplier GSTIN': r.supplierGstin,
    'Supplier Name': r.supplierName,
    'Invoice No': r.invoiceNumber,
    'Invoice Date': r.invoiceDate,
    'Books Taxable (₹)': r.booksTaxable,
    'Books Tax (₹)': r.booksTotalTax,
    'GSTR-2B Taxable (₹)': r.gstr2bTaxable,
    'GSTR-2B Tax (₹)': r.gstr2bTotalTax,
    'Variance (₹)': r.diffAmount,
    'Match Type': r.status,
    'ITC Claim Status': 'Eligible for GSTR-3B Claim',
  }));
  const wsMatched = createStyledWorksheet(matchedRows, {
    currencyColumns: ['Books Taxable (₹)', 'Books Tax (₹)', 'GSTR-2B Taxable (₹)', 'GSTR-2B Tax (₹)', 'Variance (₹)'],
    dateColumns: ['Invoice Date'],
    columnWidths: { 'S.No': 8, 'Supplier GSTIN': 18, 'Supplier Name': 30, 'Invoice No': 18, 'Invoice Date': 14, 'Books Taxable (₹)': 18, 'Books Tax (₹)': 16, 'GSTR-2B Taxable (₹)': 18, 'GSTR-2B Tax (₹)': 16, 'Variance (₹)': 14, 'Match Type': 18, 'ITC Claim Status': 25 },
  });
  XLSX.utils.book_append_sheet(wb, wsMatched, 'Matched_Invoices');

  // Sheet 3: Missing in GSTR-2B (Vendor Follow-up)
  const missing2b = records.filter((r) => r.status === 'Missing in 2B');
  const missing2bRows = missing2b.map((r, idx) => ({
    'S.No': idx + 1,
    'Supplier GSTIN': r.supplierGstin,
    'Supplier Name': r.supplierName,
    'Invoice No': r.invoiceNumber,
    'Invoice Date': r.invoiceDate,
    'Taxable (₹)': r.booksTaxable,
    'Blocked ITC (₹)': r.booksTotalTax,
    'Notice Follow-up Template': `Dear ${r.supplierName}, invoice ${r.invoiceNumber} dated ${r.invoiceDate} for ₹${r.booksTotalTax} GST is not reflected in our GSTR-2B. Please file your GSTR-1 return immediately to prevent Section 16(2)(aa) ITC reversal.`,
  }));
  const wsMissing2b = createStyledWorksheet(missing2bRows.length > 0 ? missing2bRows : [{ Note: 'Zero missing invoices in GSTR-2B' }], {
    currencyColumns: ['Taxable (₹)', 'Blocked ITC (₹)'],
    columnWidths: { 'S.No': 8, 'Supplier GSTIN': 18, 'Supplier Name': 30, 'Invoice No': 18, 'Invoice Date': 14, 'Taxable (₹)': 18, 'Blocked ITC (₹)': 18, 'Notice Follow-up Template': 65 },
  });
  XLSX.utils.book_append_sheet(wb, wsMissing2b, 'Missing_in_GSTR2B');

  // Sheet 4: Missing in Books (Unclaimed ITC)
  const missingBooks = records.filter((r) => r.status === 'Missing in Books');
  const missingBooksRows = missingBooks.map((r, idx) => ({
    'S.No': idx + 1,
    'Supplier GSTIN': r.supplierGstin,
    'Supplier Name': r.supplierName,
    'Invoice No': r.invoiceNumber,
    'Invoice Date': r.invoiceDate,
    'Portal Taxable (₹)': r.gstr2bTaxable,
    'Available Unclaimed ITC (₹)': r.gstr2bTotalTax,
    'Action Required': 'Obtain physical tax invoice copy from vendor and record entry in purchase daybook.',
  }));
  const wsMissingBooks = createStyledWorksheet(missingBooksRows.length > 0 ? missingBooksRows : [{ Note: 'Zero unclaimed items' }], {
    currencyColumns: ['Portal Taxable (₹)', 'Available Unclaimed ITC (₹)'],
    columnWidths: { 'S.No': 8, 'Supplier GSTIN': 18, 'Supplier Name': 30, 'Invoice No': 18, 'Invoice Date': 14, 'Portal Taxable (₹)': 18, 'Available Unclaimed ITC (₹)': 24, 'Action Required': 45 },
  });
  XLSX.utils.book_append_sheet(wb, wsMissingBooks, 'Missing_in_Books');

  // Sheet 5: Tax Differences & Mismatches
  const diffs = records.filter((r) => r.status === 'Tax Mismatch');
  const diffRows = diffs.map((r, idx) => ({
    'S.No': idx + 1,
    'Supplier GSTIN': r.supplierGstin,
    'Supplier Name': r.supplierName,
    'Invoice No': r.invoiceNumber,
    'Books Taxable (₹)': r.booksTaxable,
    'GSTR-2B Taxable (₹)': r.gstr2bTaxable,
    'Taxable Variance (₹)': r.diffTaxable,
    'Books Tax (₹)': r.booksTotalTax,
    'GSTR-2B Tax (₹)': r.gstr2bTotalTax,
    'Tax Discrepancy (₹)': r.diffAmount,
    'Audit Action': r.actionNote,
  }));
  const wsDiffs = createStyledWorksheet(diffRows.length > 0 ? diffRows : [{ Note: 'Zero tax mismatches' }], {
    currencyColumns: ['Books Taxable (₹)', 'GSTR-2B Taxable (₹)', 'Taxable Variance (₹)', 'Books Tax (₹)', 'GSTR-2B Tax (₹)', 'Tax Discrepancy (₹)'],
    columnWidths: { 'S.No': 8, 'Supplier GSTIN': 18, 'Supplier Name': 30, 'Invoice No': 18, 'Books Taxable (₹)': 18, 'GSTR-2B Taxable (₹)': 18, 'Taxable Variance (₹)': 18, 'Books Tax (₹)': 16, 'GSTR-2B Tax (₹)': 16, 'Tax Discrepancy (₹)': 18, 'Audit Action': 45 },
  });
  XLSX.utils.book_append_sheet(wb, wsDiffs, 'Tax_Differences');

  // Sheet 6: Duplicates Report
  const dups = records.filter((r) => r.status === 'Duplicate Record');
  const dupRows = dups.map((r, idx) => ({
    'Duplicate #': idx + 1,
    'Supplier GSTIN': r.supplierGstin,
    'Supplier Name': r.supplierName,
    'Invoice No': r.invoiceNumber,
    'Invoice Date': r.invoiceDate,
    'Books Value (₹)': r.booksTaxable,
    'Books Tax (₹)': r.booksTotalTax,
    'Audit Warning': 'Double entry in Purchase Register. Potential double claim in GSTR-3B.',
  }));
  const wsDups = createStyledWorksheet(dupRows.length > 0 ? dupRows : [{ Note: 'Zero duplicate invoices detected' }], {
    currencyColumns: ['Books Value (₹)', 'Books Tax (₹)'],
    columnWidths: { 'Duplicate #': 14, 'Supplier GSTIN': 18, 'Supplier Name': 30, 'Invoice No': 18, 'Invoice Date': 14, 'Books Value (₹)': 18, 'Books Tax (₹)': 16, 'Audit Warning': 45 },
  });
  XLSX.utils.book_append_sheet(wb, wsDups, 'Duplicates_Report');

  return wb;
}
