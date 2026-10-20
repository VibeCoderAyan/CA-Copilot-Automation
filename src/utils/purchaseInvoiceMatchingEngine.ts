import * as XLSX from 'xlsx';
import { parseDate, parseDateObj, parseAmount, normalizeRefNo } from './dataSanitizer';
import { createStyledWorksheet } from './excelStyler';

export type PurchaseMatchStatus =
  | 'matched'
  | 'amount_mismatch'
  | 'tax_mismatch'
  | 'date_mismatch'
  | 'gstin_mismatch'
  | 'missing_in_invoices'
  | 'missing_in_pr'
  | 'duplicate_invoice'
  | 'review_needed';

export interface PurchaseRegisterRow {
  id: string;
  sourceFile: string;
  sheetName: string;
  rowIndex: number;
  supplierName: string;
  gstin: string;
  invoiceNo: string;
  normalizedInvoiceNo: string;
  invoiceDate: string;
  taxableValue: number;
  cgst: number;
  sgst: number;
  igst: number;
  totalAmount: number;
  raw: Record<string, any>;
}

export interface SupplierInvoiceRow {
  id: string;
  sourceFile: string;
  sheetName: string;
  rowIndex: number;
  supplierName: string;
  gstin: string;
  invoiceNo: string;
  normalizedInvoiceNo: string;
  invoiceDate: string;
  taxableValue: number;
  cgst: number;
  sgst: number;
  igst: number;
  totalAmount: number;
  raw: Record<string, any>;
}

export interface PurchaseMatchItem {
  id: string;
  invoiceNo: string;
  status: PurchaseMatchStatus;
  severity: 'clean' | 'warning' | 'critical';
  prRow?: PurchaseRegisterRow;
  invRow?: SupplierInvoiceRow;
  supplierName: string;
  gstin: string;
  prTotal: number;
  invTotal: number;
  totalVariance: number;
  prTaxable: number;
  invTaxable: number;
  taxableVariance: number;
  prTax: number;
  invTax: number;
  taxVariance: number;
  prDate: string;
  invDate: string;
  dateDifferenceDays: number;
  discrepancyDetails: string;
  reviewStatus: 'unresolved' | 'approved_match' | 'rejected_discrepancy' | 'query_sent';
  auditorNotes?: string;
}

export interface PurchaseMatchSummary {
  totalPrRecords: number;
  totalRegisterRows?: number;
  totalInvRecords: number;
  matchedCount: number;
  mismatchCount: number;
  missingInInvCount: number;
  missingInPrCount: number;
  duplicateCount: number;
  reviewNeededCount: number;
  totalPrValue: number;
  totalRegisterAmount?: number;
  totalInvValue: number;
  netVariance: number;
  matchPercentage: number;
  items: PurchaseMatchItem[];
  summary?: PurchaseMatchSummary;
}

export interface PurchaseMatchingConfig {
  amountTolerance: number; // default ₹1.00
  taxTolerance: number; // default ₹0.50
  dateToleranceDays: number; // default 3 days
  requireGstinMatch: boolean;
  fuzzySupplierMatch: boolean;
}

export interface PurchaseColumnMapping {
  supplierCol: string;
  gstinCol: string;
  invoiceNoCol: string;
  dateCol: string;
  taxableCol: string;
  cgstCol: string;
  sgstCol: string;
  igstCol: string;
  totalCol: string;
}

import { detectColumnCategory } from './columnDetector';

export function detectPurchaseColumns(headers: string[]): PurchaseColumnMapping {
  const m: PurchaseColumnMapping = {
    supplierCol: '',
    gstinCol: '',
    invoiceNoCol: '',
    dateCol: '',
    taxableCol: '',
    cgstCol: '',
    sgstCol: '',
    igstCol: '',
    totalCol: '',
  };

  for (const h of headers) {
    const cat = detectColumnCategory(h);
    if (!m.supplierCol && cat === 'partyName') m.supplierCol = h;
    else if (!m.gstinCol && cat === 'gstin') m.gstinCol = h;
    else if (!m.invoiceNoCol && cat === 'voucherNo') m.invoiceNoCol = h;
    else if (!m.dateCol && cat === 'date') m.dateCol = h;
    else if (!m.taxableCol && cat === 'taxableAmount') m.taxableCol = h;
    else if (!m.cgstCol && cat === 'cgst') m.cgstCol = h;
    else if (!m.sgstCol && cat === 'sgst') m.sgstCol = h;
    else if (!m.igstCol && cat === 'igst') m.igstCol = h;
    else if (!m.totalCol && cat === 'totalAmount') m.totalCol = h;
  }

  const isUsed = (headerName: string) => Object.values(m).includes(headerName);

  // Secondary regex pass
  for (const h of headers) {
    if (isUsed(h)) continue;
    const clean = h.toLowerCase().trim();
    if (!m.supplierCol && /supplier|vendor|party|seller|trade.*name|biller/i.test(clean)) m.supplierCol = h;
    else if (!m.gstinCol && /gstin|gst.*no|tax.*id/i.test(clean)) m.gstinCol = h;
    else if (!m.invoiceNoCol && /inv.*no|bill.*no|invoice.*num|doc.*no|vch.*no/i.test(clean)) m.invoiceNoCol = h;
    else if (!m.dateCol && /date|inv.*dt|bill.*dt/i.test(clean)) m.dateCol = h;
    else if (!m.taxableCol && /taxable|basic|assessable/i.test(clean)) m.taxableCol = h;
    else if (!m.cgstCol && /cgst/i.test(clean)) m.cgstCol = h;
    else if (!m.sgstCol && /sgst|utgst/i.test(clean)) m.sgstCol = h;
    else if (!m.igstCol && /igst/i.test(clean)) m.igstCol = h;
    else if (!m.totalCol && /total|inv.*amt|gross|net.*amt|net.*payable/i.test(clean)) m.totalCol = h;
  }

  // Safety collision check: if taxableCol and totalCol collided on the same header
  if (m.taxableCol && m.totalCol && m.taxableCol === m.totalCol) {
    const altTotal = headers.find((h) => h !== m.taxableCol && /gross|total|net.*amt/i.test(h));
    if (altTotal) {
      m.totalCol = altTotal;
    } else {
      const altTaxable = headers.find((h) => h !== m.totalCol && /assessable|taxable|basic/i.test(h));
      if (altTaxable) {
        m.taxableCol = altTaxable;
      }
    }
  }

  return m;
}

/**
 * Main matching engine comparing Purchase Register against Supplier Invoices
 */
export function matchPurchaseRegisterWithInvoices(
  prRawRows: Record<string, any>[],
  invRawRows: Record<string, any>[],
  prMapping: PurchaseColumnMapping,
  invMapping: PurchaseColumnMapping,
  config: PurchaseMatchingConfig = {
    amountTolerance: 1.0,
    taxTolerance: 0.5,
    dateToleranceDays: 3,
    requireGstinMatch: false,
    fuzzySupplierMatch: true,
  },
  prFileName = 'Purchase_Register.xlsx',
  invFileName = 'Supplier_Invoices.xlsx'
): PurchaseMatchSummary {
  // Parse Purchase Register
  const prRows: PurchaseRegisterRow[] = prRawRows.map((r, idx) => {
    const invNo = String(r[prMapping.invoiceNoCol] || '').trim();
    const cgst = parseAmount(r[prMapping.cgstCol]);
    const sgst = parseAmount(r[prMapping.sgstCol]);
    const igst = parseAmount(r[prMapping.igstCol]);
    const taxableValue = parseAmount(r[prMapping.taxableCol]);
    let totalAmount = parseAmount(r[prMapping.totalCol]);
    if (totalAmount === 0 && taxableValue > 0) {
      totalAmount = taxableValue + cgst + sgst + igst;
    }

    return {
      id: `pr-${idx + 1}`,
      sourceFile: prFileName,
      sheetName: 'Purchase_Register',
      rowIndex: idx + 2,
      supplierName: String(r[prMapping.supplierCol] || '').trim(),
      gstin: String(r[prMapping.gstinCol] || '').trim().toUpperCase(),
      invoiceNo: invNo,
      normalizedInvoiceNo: normalizeRefNo(invNo),
      invoiceDate: parseDate(r[prMapping.dateCol], 'DD-MM-YYYY'),
      taxableValue,
      cgst,
      sgst,
      igst,
      totalAmount,
      raw: r,
    };
  });

  // Parse Supplier Invoices
  const invRows: SupplierInvoiceRow[] = invRawRows.map((r, idx) => {
    const invNo = String(r[invMapping.invoiceNoCol] || '').trim();
    const cgst = parseAmount(r[invMapping.cgstCol]);
    const sgst = parseAmount(r[invMapping.sgstCol]);
    const igst = parseAmount(r[invMapping.igstCol]);
    const taxableValue = parseAmount(r[invMapping.taxableCol]);
    let totalAmount = parseAmount(r[invMapping.totalCol]);
    if (totalAmount === 0 && taxableValue > 0) {
      totalAmount = taxableValue + cgst + sgst + igst;
    }

    return {
      id: `inv-${idx + 1}`,
      sourceFile: invFileName,
      sheetName: 'Supplier_Invoices',
      rowIndex: idx + 2,
      supplierName: String(r[invMapping.supplierCol] || '').trim(),
      gstin: String(r[invMapping.gstinCol] || '').trim().toUpperCase(),
      invoiceNo: invNo,
      normalizedInvoiceNo: normalizeRefNo(invNo),
      invoiceDate: parseDate(r[invMapping.dateCol], 'DD-MM-YYYY'),
      taxableValue,
      cgst,
      sgst,
      igst,
      totalAmount,
      raw: r,
    };
  });

  const matchedInvIds = new Set<string>();
  const items: PurchaseMatchItem[] = [];

  // Track duplicate invoice numbers in PR
  const prDupTracker = new Map<string, number>();
  prRows.forEach((pr) => {
    if (pr.normalizedInvoiceNo) {
      prDupTracker.set(pr.normalizedInvoiceNo, (prDupTracker.get(pr.normalizedInvoiceNo) || 0) + 1);
    }
  });

  let matchedCount = 0;
  let mismatchCount = 0;
  let missingInInvCount = 0;
  let duplicateCount = 0;
  let reviewNeededCount = 0;

  // Process Purchase Register Rows
  prRows.forEach((pr) => {
    // Duplicate check
    const isPrDuplicate = (prDupTracker.get(pr.normalizedInvoiceNo) || 0) > 1;

    // Search for match in Supplier Invoices (ensure normalized invoice number is non-empty)
    const candidate = invRows.find(
      (inv) =>
        !matchedInvIds.has(inv.id) &&
        pr.normalizedInvoiceNo !== '' &&
        inv.normalizedInvoiceNo === pr.normalizedInvoiceNo
    );

    if (candidate) {
      matchedInvIds.add(candidate.id);

      const totalDiff = Math.abs(pr.totalAmount - candidate.totalAmount);
      const taxableDiff = Math.abs(pr.taxableValue - candidate.taxableValue);
      const prTotalTax = pr.cgst + pr.sgst + pr.igst;
      const invTotalTax = candidate.cgst + candidate.sgst + candidate.igst;
      const taxDiff = Math.abs(prTotalTax - invTotalTax);

      // Date difference using robust date parsing
      let dateDiffDays = 0;
      const pDateObj = parseDateObj(pr.invoiceDate);
      const iDateObj = parseDateObj(candidate.invoiceDate);
      if (pDateObj && iDateObj && !isNaN(pDateObj.getTime()) && !isNaN(iDateObj.getTime())) {
        const diffMs = Math.abs(pDateObj.getTime() - iDateObj.getTime());
        dateDiffDays = Math.round(diffMs / (1000 * 60 * 60 * 24));
      }
      if (isNaN(dateDiffDays)) dateDiffDays = 0;

      // GSTIN check
      const isGstinMatch =
        !config.requireGstinMatch ||
        !pr.gstin ||
        !candidate.gstin ||
        pr.gstin === candidate.gstin;

      let status: PurchaseMatchStatus = 'matched';
      let severity: 'clean' | 'warning' | 'critical' = 'clean';
      const reasons: string[] = [];

      if (isPrDuplicate) {
        status = 'duplicate_invoice';
        severity = 'critical';
        reasons.push('Duplicate invoice number in Purchase Register');
        duplicateCount++;
      } else if (totalDiff > config.amountTolerance) {
        status = 'amount_mismatch';
        severity = 'critical';
        reasons.push(`Total amount variance of ₹${totalDiff.toFixed(2)} (PR: ₹${pr.totalAmount.toLocaleString('en-IN')} vs Inv: ₹${candidate.totalAmount.toLocaleString('en-IN')})`);
        mismatchCount++;
      } else if (taxDiff > config.taxTolerance) {
        status = 'tax_mismatch';
        severity = 'warning';
        reasons.push(`Tax component variance of ₹${taxDiff.toFixed(2)} (PR: ₹${prTotalTax.toFixed(2)} vs Inv: ₹${invTotalTax.toFixed(2)})`);
        mismatchCount++;
      } else if (dateDiffDays > config.dateToleranceDays) {
        status = 'date_mismatch';
        severity = 'warning';
        reasons.push(`Invoice date difference of ${dateDiffDays} days (PR: ${pr.invoiceDate} vs Inv: ${candidate.invoiceDate})`);
        mismatchCount++;
      } else if (!isGstinMatch) {
        status = 'gstin_mismatch';
        severity = 'warning';
        reasons.push(`GSTIN mismatch: PR has ${pr.gstin}, Invoice has ${candidate.gstin}`);
        mismatchCount++;
      } else {
        matchedCount++;
      }

      items.push({
        id: `match-${pr.id}-${candidate.id}`,
        invoiceNo: pr.invoiceNo,
        status,
        severity,
        prRow: pr,
        invRow: candidate,
        supplierName: pr.supplierName || candidate.supplierName,
        gstin: pr.gstin || candidate.gstin,
        prTotal: pr.totalAmount,
        invTotal: candidate.totalAmount,
        totalVariance: totalDiff,
        prTaxable: pr.taxableValue,
        invTaxable: candidate.taxableValue,
        taxableVariance: taxableDiff,
        prTax: prTotalTax,
        invTax: invTotalTax,
        taxVariance: taxDiff,
        prDate: pr.invoiceDate,
        invDate: candidate.invoiceDate,
        dateDifferenceDays: dateDiffDays,
        discrepancyDetails: reasons.length > 0 ? reasons.join(' | ') : '100% Exact Match',
        reviewStatus: status === 'matched' ? 'approved_match' : 'unresolved',
      });
    } else {
      // Missing in Supplier Invoices
      missingInInvCount++;
      items.push({
        id: `missing-inv-${pr.id}`,
        invoiceNo: pr.invoiceNo,
        status: 'missing_in_invoices',
        severity: 'critical',
        prRow: pr,
        supplierName: pr.supplierName,
        gstin: pr.gstin,
        prTotal: pr.totalAmount,
        invTotal: 0,
        totalVariance: pr.totalAmount,
        prTaxable: pr.taxableValue,
        invTaxable: 0,
        taxableVariance: pr.taxableValue,
        prTax: pr.cgst + pr.sgst + pr.igst,
        invTax: 0,
        taxVariance: pr.cgst + pr.sgst + pr.igst,
        prDate: pr.invoiceDate,
        invDate: '-',
        dateDifferenceDays: 0,
        discrepancyDetails: `Booked in Purchase Register (Row ${pr.rowIndex}) but not found in Supplier Invoices`,
        reviewStatus: 'unresolved',
      });
    }
  });

  // Check remaining unmatched supplier invoices (Missing in Purchase Register)
  let missingInPrCount = 0;
  invRows.forEach((inv) => {
    if (!matchedInvIds.has(inv.id)) {
      missingInPrCount++;
      items.push({
        id: `missing-pr-${inv.id}`,
        invoiceNo: inv.invoiceNo,
        status: 'missing_in_pr',
        severity: 'warning',
        invRow: inv,
        supplierName: inv.supplierName,
        gstin: inv.gstin,
        prTotal: 0,
        invTotal: inv.totalAmount,
        totalVariance: inv.totalAmount,
        prTaxable: 0,
        invTaxable: inv.taxableValue,
        taxableVariance: inv.taxableValue,
        prTax: 0,
        invTax: inv.cgst + inv.sgst + inv.igst,
        taxVariance: inv.cgst + inv.sgst + inv.igst,
        prDate: '-',
        invDate: inv.invoiceDate,
        dateDifferenceDays: 0,
        discrepancyDetails: `Billed by Supplier (Row ${inv.rowIndex}) but omitted from Purchase Register`,
        reviewStatus: 'unresolved',
      });
    }
  });

  const totalPrAmt = prRows.reduce((sum, r) => sum + r.totalAmount, 0);
  const totalInvAmt = invRows.reduce((sum, r) => sum + r.totalAmount, 0);
  const netVariance = Math.abs(totalPrAmt - totalInvAmt);
  const totalRecords = prRows.length + missingInPrCount;
  const matchPercentage = totalRecords > 0 ? (matchedCount / prRows.length) * 100 : 0;

  const result: PurchaseMatchSummary = {
    totalPrRecords: prRows.length,
    totalRegisterRows: prRows.length,
    totalInvRecords: invRows.length,
    matchedCount,
    mismatchCount,
    missingInInvCount,
    missingInPrCount,
    duplicateCount,
    reviewNeededCount,
    totalPrValue: totalPrAmt,
    totalRegisterAmount: totalPrAmt,
    totalInvValue: totalInvAmt,
    netVariance,
    matchPercentage,
    items,
  };
  result.summary = result;
  return result;
}

/**
 * Exports complete 6-sheet reconciliation workbook
 */
export function exportPurchaseMatchingWorkbook(
  summary: PurchaseMatchSummary,
  clientName = 'Client Enterprise',
  period = 'FY 2026-27'
): XLSX.WorkBook {
  const wb = XLSX.utils.book_new();

  // Sheet 1: Executive Summary
  const summaryRows = [
    { Parameter: 'Client Entity', Value: clientName },
    { Parameter: 'Reconciliation Period', Value: period },
    { Parameter: 'Purchase Register Records', Value: summary.totalPrRecords },
    { Parameter: 'Supplier Invoices Scanned', Value: summary.totalInvRecords },
    { Parameter: 'Fully Matched Invoices', Value: summary.matchedCount },
    { Parameter: 'Matching Success Rate', Value: `${summary.matchPercentage.toFixed(1)}%` },
    { Parameter: 'Discrepant Records (Amount / Date / Tax)', Value: summary.mismatchCount },
    { Parameter: 'Missing in Supplier Invoices', Value: summary.missingInInvCount },
    { Parameter: 'Missing in Purchase Register', Value: summary.missingInPrCount },
    { Parameter: 'Duplicate Invoice Numbers', Value: summary.duplicateCount },
    { Parameter: 'Total Purchase Register Value (₹)', Value: summary.totalPrValue },
    { Parameter: 'Total Supplier Invoices Value (₹)', Value: summary.totalInvValue },
    { Parameter: 'Net Reconciliation Variance (₹)', Value: summary.netVariance },
  ];
  const wsSummary = createStyledWorksheet(summaryRows, {
    currencyColumns: ['Value'],
    columnWidths: { Parameter: 38, Value: 30 },
  });
  XLSX.utils.book_append_sheet(wb, wsSummary, 'Summary');

  // Sheet 2: Matched Invoices
  const matchedRows = summary.items
    .filter((i) => i.status === 'matched')
    .map((m, idx) => ({
      '#': idx + 1,
      'Invoice #': m.invoiceNo,
      'Supplier Name': m.supplierName,
      'GSTIN': m.gstin,
      'Invoice Date': m.prDate,
      'Taxable Value (₹)': m.prTaxable,
      'Total Amount (₹)': m.prTotal,
      'PR Row': m.prRow?.rowIndex || '-',
      'Invoice Row': m.invRow?.rowIndex || '-',
      'Status': 'MATCHED',
    }));
  const wsMatched = createStyledWorksheet(matchedRows, {
    currencyColumns: ['Taxable Value (₹)', 'Total Amount (₹)'],
  });
  XLSX.utils.book_append_sheet(wb, wsMatched, 'Matched_Invoices');

  // Sheet 3: PR Exceptions (Missing in Invoices / Variances)
  const prExceptions = summary.items
    .filter((i) => i.prRow && i.status !== 'matched')
    .map((e, idx) => ({
      '#': idx + 1,
      'Invoice #': e.invoiceNo,
      'Supplier Name': e.supplierName,
      'PR Amount (₹)': e.prTotal,
      'Supplier Amount (₹)': e.invTotal,
      'Variance (₹)': e.totalVariance,
      'PR Date': e.prDate,
      'Supplier Date': e.invDate,
      'PR Row': e.prRow?.rowIndex || '-',
      'Classification': e.status.replace(/_/g, ' ').toUpperCase(),
      'Discrepancy Details': e.discrepancyDetails,
      'Review Status': e.reviewStatus.toUpperCase(),
    }));
  const wsPrExceptions = createStyledWorksheet(prExceptions, {
    currencyColumns: ['PR Amount (₹)', 'Supplier Amount (₹)', 'Variance (₹)'],
  });
  XLSX.utils.book_append_sheet(wb, wsPrExceptions, 'PR_Exceptions');

  // Sheet 4: Invoice Exceptions (Missing in PR)
  const invExceptions = summary.items
    .filter((i) => i.status === 'missing_in_pr')
    .map((e, idx) => ({
      '#': idx + 1,
      'Invoice #': e.invoiceNo,
      'Supplier Name': e.supplierName,
      'GSTIN': e.gstin,
      'Invoice Date': e.invDate,
      'Taxable Value (₹)': e.invTaxable,
      'Invoice Amount (₹)': e.invTotal,
      'Invoice Row': e.invRow?.rowIndex || '-',
      'Auditor Note': 'Billed by vendor but omitted in internal Purchase Register (ITC Risk)',
    }));
  const wsInvExceptions = createStyledWorksheet(invExceptions, {
    currencyColumns: ['Taxable Value (₹)', 'Invoice Amount (₹)'],
  });
  XLSX.utils.book_append_sheet(wb, wsInvExceptions, 'Invoice_Exceptions');

  // Sheet 5: Duplicates
  const duplicates = summary.items
    .filter((i) => i.status === 'duplicate_invoice')
    .map((d, idx) => ({
      '#': idx + 1,
      'Duplicate Invoice #': d.invoiceNo,
      'Supplier Name': d.supplierName,
      'Amount (₹)': d.prTotal,
      'PR Row': d.prRow?.rowIndex || '-',
      'Details': d.discrepancyDetails,
    }));
  const wsDuplicates = createStyledWorksheet(duplicates, {
    currencyColumns: ['Amount (₹)'],
  });
  XLSX.utils.book_append_sheet(wb, wsDuplicates, 'Duplicates');

  // Sheet 6: Manual Review Items
  const reviewItems = summary.items
    .filter((i) => i.status !== 'matched')
    .map((r, idx) => ({
      '#': idx + 1,
      'Invoice #': r.invoiceNo,
      'Supplier': r.supplierName,
      'Status': r.status.replace(/_/g, ' ').toUpperCase(),
      'Variance (₹)': r.totalVariance,
      'Audit Observation': r.discrepancyDetails,
      'Review Resolution': r.reviewStatus.toUpperCase(),
    }));
  const wsReview = createStyledWorksheet(reviewItems, {
    currencyColumns: ['Variance (₹)'],
  });
  XLSX.utils.book_append_sheet(wb, wsReview, 'Manual_Review');

  return wb;
}

/**
 * Realistic Dual Sample Dataset (Purchase Register vs Supplier Invoices)
 */
export function getSamplePurchaseMatchingData(): {
  purchaseRegister: Record<string, any>[];
  supplierInvoices: Record<string, any>[];
  scannedInvoices: Record<string, any>[];
} {
  const purchaseRegister = [
    { 'Invoice No': 'INV/2026/041', 'Date': '05-08-2026', 'Supplier Name': 'Tata Steel Downstream Ltd', 'GSTIN': '27AAACT2727Q1ZG', 'Taxable Value': 450000, 'CGST': 40500, 'SGST': 40500, 'IGST': 0, 'Total Amount': 531000 },
    { 'Invoice No': 'INV/2026/042', 'Date': '08-08-2026', 'Supplier Name': 'Godrej Security Systems', 'GSTIN': '27AAACG0549H1ZT', 'Taxable Value': 120000, 'CGST': 0, 'SGST': 0, 'IGST': 21600, 'Total Amount': 141600 },
    { 'Invoice No': 'INV/2026/043', 'Date': '12-08-2026', 'Supplier Name': 'Apex Industrial Packaging', 'GSTIN': '27AAACA9921L1ZM', 'Taxable Value': 85000, 'CGST': 7650, 'SGST': 7650, 'IGST': 0, 'Total Amount': 100300 }, // AMOUNT MISMATCH! Supplier billed 105,000
    { 'Invoice No': 'INV/2026/044', 'Date': '15-08-2026', 'Supplier Name': 'Schneider Electric India', 'GSTIN': '27AAACS1234D1ZF', 'Taxable Value': 240000, 'CGST': 0, 'SGST': 0, 'IGST': 43200, 'Total Amount': 283200 }, // DATE MISMATCH! Vendor billed on 28-08-2026
    { 'Invoice No': 'INV/2026/045', 'Date': '20-08-2026', 'Supplier Name': 'Siemens Healthcare Diagnostics', 'GSTIN': '27AAACS8899K1ZV', 'Taxable Value': 350000, 'CGST': 31500, 'SGST': 31500, 'IGST': 0, 'Total Amount': 413000 }, // MISSING IN SUPPLIER INVOICES!
    { 'Invoice No': 'INV/2026/046', 'Date': '24-08-2026', 'Supplier Name': 'L&T Technology Services', 'GSTIN': '27AAACL1441R1ZP', 'Taxable Value': 180000, 'CGST': 16200, 'SGST': 16200, 'IGST': 0, 'Total Amount': 212400 },
    { 'Invoice No': 'INV/2026/041', 'Date': '29-08-2026', 'Supplier Name': 'Tata Steel Downstream Ltd', 'GSTIN': '27AAACT2727Q1ZG', 'Taxable Value': 450000, 'CGST': 40500, 'SGST': 40500, 'IGST': 0, 'Total Amount': 531000 }, // DUPLICATE INVOICE NO
  ];

  const supplierInvoices = [
    { 'Bill No': 'INV-2026-041', 'Bill Date': '05-08-2026', 'Vendor Name': 'Tata Steel Downstream Ltd', 'Tax ID': '27AAACT2727Q1ZG', 'Assessable Amount': 450000, 'CGST': 40500, 'SGST': 40500, 'IGST': 0, 'Gross Total': 531000 },
    { 'Bill No': 'INV-2026-042', 'Bill Date': '08-08-2026', 'Vendor Name': 'Godrej Security Systems', 'Tax ID': '27AAACG0549H1ZT', 'Assessable Amount': 120000, 'CGST': 0, 'SGST': 0, 'IGST': 21600, 'Gross Total': 141600 },
    { 'Bill No': 'INV-2026-043', 'Bill Date': '12-08-2026', 'Vendor Name': 'Apex Industrial Packaging', 'Tax ID': '27AAACA9921L1ZM', 'Assessable Amount': 89000, 'CGST': 8010, 'SGST': 8010, 'IGST': 0, 'Gross Total': 105020 }, // Variance of ₹4,720
    { 'Bill No': 'INV-2026-044', 'Bill Date': '28-08-2026', 'Vendor Name': 'Schneider Electric India', 'Tax ID': '27AAACS1234D1ZF', 'Assessable Amount': 240000, 'CGST': 0, 'SGST': 0, 'IGST': 43200, 'Gross Total': 283200 }, // 13 days lag
    { 'Bill No': 'INV-2026-046', 'Bill Date': '24-08-2026', 'Vendor Name': 'L&T Technology Services', 'Tax ID': '27AAACL1441R1ZP', 'Assessable Amount': 180000, 'CGST': 16200, 'SGST': 16200, 'IGST': 0, 'Gross Total': 212400 },
    { 'Bill No': 'INV-2026-999', 'Bill Date': '30-08-2026', 'Vendor Name': 'Bosch Rexroth Hydraulics', 'Tax ID': '27AAACB5521N1ZK', 'Assessable Amount': 95000, 'CGST': 0, 'SGST': 0, 'IGST': 17100, 'Gross Total': 112100 }, // MISSING IN PURCHASE REGISTER!
  ];

  return {
    purchaseRegister,
    supplierInvoices,
    scannedInvoices: supplierInvoices,
  };
}
