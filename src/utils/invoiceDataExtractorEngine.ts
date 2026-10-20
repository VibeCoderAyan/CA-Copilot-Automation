import * as XLSX from 'xlsx';
import { parseAmount, parseDate, isValidGstin } from './dataSanitizer';
import { createStyledWorksheet } from './excelStyler';
import { parseInvoicePdf } from './pdfExtractor';

export interface ExtractedInvoiceRecord {
  id: string;
  fileName: string;
  fileSize?: string;
  pageNumber?: number;
  invoiceNumber: string;
  invoiceDate: string;
  supplierName: string;
  supplierGstin: string;
  customerGstin?: string;
  placeOfSupply?: string;
  taxableAmount: number;
  cgst: number;
  sgst: number;
  igst: number;
  cess?: number;
  totalTax: number;
  totalAmount: number;
  currency?: string;
  status: 'Verified' | 'Needs Review' | 'Duplicate';
  isDuplicate?: boolean;
  duplicateReason?: string;
  validationIssues: string[];
  confidenceScore: number;
  raw?: Record<string, any>;
}

export interface InvoiceBatchResult {
  totalInvoices: number;
  verifiedCount: number;
  needsReviewCount: number;
  duplicateCount: number;
  totalTaxable: number;
  totalTax: number;
  totalInvoiceAmount: number;
  invoices: ExtractedInvoiceRecord[];
}

/**
 * Validates a single extracted invoice record against statutory tax math and GSTIN formats
 */
export function validateInvoiceRecord(
  record: Omit<ExtractedInvoiceRecord, 'status' | 'validationIssues' | 'totalTax'>
): ExtractedInvoiceRecord {
  const issues: string[] = [];
  const cess = record.cess || 0;
  const totalTax = record.cgst + record.sgst + record.igst + cess;

  // 1. Mandatory Invoice Number Check
  if (!record.invoiceNumber || record.invoiceNumber.trim() === '' || record.invoiceNumber === '-') {
    issues.push('Missing Invoice Number');
  }

  // 2. Mandatory Date Check
  if (!record.invoiceDate || record.invoiceDate.trim() === '' || record.invoiceDate === '-') {
    issues.push('Missing / Invalid Date');
  }

  // 3. Supplier GSTIN Check
  const gstinClean = record.supplierGstin ? record.supplierGstin.trim().toUpperCase() : '';
  if (!gstinClean) {
    issues.push('Missing Supplier GSTIN');
  } else if (!isValidGstin(gstinClean)) {
    issues.push(`Invalid Supplier GSTIN Format: ${gstinClean}`);
  }

  // 4. Customer GSTIN Check if present
  if (record.customerGstin && !isValidGstin(record.customerGstin.trim().toUpperCase())) {
    issues.push(`Invalid Customer GSTIN Format: ${record.customerGstin}`);
  }

  // 5. Tax Math Validation (Taxable + Tax == Total with ±₹1.50 rounding tolerance)
  const calculatedTotal = record.taxableAmount + totalTax;
  if (record.totalAmount > 0 && calculatedTotal > 0) {
    const diff = Math.abs(calculatedTotal - record.totalAmount);
    if (diff > 1.5) {
      issues.push(
        `Tax Math Mismatch (Taxable ₹${record.taxableAmount.toFixed(2)} + Tax ₹${totalTax.toFixed(2)} != Total ₹${record.totalAmount.toFixed(2)})`
      );
    }
  } else if (record.totalAmount === 0 && record.taxableAmount === 0) {
    issues.push('Zero / Unparseable Financial Values');
  }

  let status: 'Verified' | 'Needs Review' | 'Duplicate' = issues.length === 0 ? 'Verified' : 'Needs Review';
  if (record.isDuplicate) {
    status = 'Duplicate';
  }

  return {
    ...record,
    supplierGstin: gstinClean,
    cess,
    currency: record.currency || 'INR',
    totalTax,
    status,
    validationIssues: issues,
  };
}

/**
 * Detects duplicate invoices within a batch based on configurable invoice # and supplier rules
 */
export function detectDuplicatesInBatch(records: ExtractedInvoiceRecord[]): ExtractedInvoiceRecord[] {
  const seenMap = new Map<string, number>();

  // Count occurrences
  records.forEach((r) => {
    const normInv = r.invoiceNumber.toLowerCase().replace(/[^a-z0-9]/g, '');
    const normSupplier = (r.supplierGstin || r.supplierName).toLowerCase().replace(/[^a-z0-9]/g, '');
    if (normInv && normSupplier) {
      const key = `${normSupplier}_${normInv}_${Math.round(r.totalAmount)}`;
      seenMap.set(key, (seenMap.get(key) || 0) + 1);
    }
  });

  return records.map((r) => {
    const normInv = r.invoiceNumber.toLowerCase().replace(/[^a-z0-9]/g, '');
    const normSupplier = (r.supplierGstin || r.supplierName).toLowerCase().replace(/[^a-z0-9]/g, '');
    if (normInv && normSupplier) {
      const key = `${normSupplier}_${normInv}_${Math.round(r.totalAmount)}`;
      if ((seenMap.get(key) || 0) > 1) {
        const issues = [...r.validationIssues];
        if (!issues.some((i) => i.includes('Duplicate invoice'))) {
          issues.push(`Duplicate invoice detected: Same supplier, invoice #${r.invoiceNumber}, and total amount.`);
        }
        return {
          ...r,
          isDuplicate: true,
          duplicateReason: 'Identical supplier, invoice number, and total value in batch',
          status: 'Duplicate',
          validationIssues: issues,
        };
      }
    }
    return r;
  });
}

/**
 * Extracts and validates invoice data from an uploaded digital PDF file
 */
export async function extractFromInvoicePdf(file: File): Promise<ExtractedInvoiceRecord> {
  const pdfRes = await parseInvoicePdf(file);

  if (!pdfRes.success || !pdfRes.invoice) {
    throw new Error(
      pdfRes.errorMessage ||
        `Unsupported PDF '${file.name}': Not a valid digital tax invoice. Missing required invoice fields (GSTIN, Invoice Number, or Financial Amounts). Process halted.`
    );
  }

  const inv = pdfRes.invoice;
  const initialRecord = {
    id: `inv-${Date.now()}-${Math.random().toString(36).substr(2, 4)}`,
    fileName: file.name,
    fileSize: `${Math.round(file.size / 1024)} KB`,
    invoiceNumber: inv.invoiceNumber,
    invoiceDate: inv.invoiceDate ? parseDate(inv.invoiceDate, 'DD-MM-YYYY') : '',
    supplierName: inv.supplierName,
    supplierGstin: inv.supplierGstin,
    customerGstin: (inv as any).customerGstin || '',
    placeOfSupply: (inv as any).placeOfSupply || '',
    taxableAmount: inv.taxableAmount,
    cgst: inv.cgst,
    sgst: inv.sgst,
    igst: inv.igst,
    cess: (inv as any).cess || 0,
    totalAmount: inv.totalAmount,
    currency: (inv as any).currency || 'INR',
    confidenceScore: 95,
  };

  return validateInvoiceRecord(initialRecord);
}

import { detectColumnCategory } from './columnDetector';

/**
 * Extracts and validates invoice records from a spreadsheet row or image metadata
 */
export function extractFromSpreadsheetRows(rows: Record<string, any>[], sourceFileName: string): InvoiceBatchResult {
  const invoices: ExtractedInvoiceRecord[] = [];
  if (!rows || rows.length === 0) {
    return {
      totalInvoices: 0,
      verifiedCount: 0,
      needsReviewCount: 0,
      duplicateCount: 0,
      totalTaxable: 0,
      totalTax: 0,
      totalInvoiceAmount: 0,
      invoices: [],
    };
  }

  // Detect column mapping once from first non-empty row
  const sampleKeys = Object.keys(rows[0] || {});
  let invNoKey = '';
  let dateKey = '';
  let partyKey = '';
  let gstinKey = '';
  let taxableKey = '';
  let cgstKey = '';
  let sgstKey = '';
  let igstKey = '';
  let totalKey = '';

  for (const k of sampleKeys) {
    const cat = detectColumnCategory(k);
    const clean = k.toLowerCase().trim();

    if (!invNoKey && (cat === 'voucherNo' || /\b(invoice|inv|bill|vch|doc|inum)\b/i.test(clean))) invNoKey = k;
    else if (!dateKey && (cat === 'date' || /\b(date|dt)\b/i.test(clean))) dateKey = k;
    else if (!partyKey && (cat === 'partyName' || /\b(party|vendor|supplier|customer|trade_name)\b/i.test(clean))) partyKey = k;
    else if (!gstinKey && (cat === 'gstin' || /\b(gstin|gst|tax id)\b/i.test(clean))) gstinKey = k;
    else if (!taxableKey && (cat === 'taxableAmount' || /\b(taxable|basic|net amt)\b/i.test(clean))) taxableKey = k;
    else if (!cgstKey && (cat === 'cgst' || /\bcgst\b/i.test(clean))) cgstKey = k;
    else if (!sgstKey && (cat === 'sgst' || /\bsgst\b/i.test(clean))) sgstKey = k;
    else if (!igstKey && (cat === 'igst' || /\bigst\b/i.test(clean))) igstKey = k;
    else if (!totalKey && (cat === 'totalAmount' || /\b(total|gross|inv amt|amount)\b/i.test(clean))) totalKey = k;
  }

  // Secondary search for customer GSTIN, POS, Cess, Currency
  const custGstinKey = sampleKeys.find((k) => /\b(customer|buyer|recipient).*gst/i.test(k.toLowerCase().trim())) || '';
  const posKey = sampleKeys.find((k) => /\b(place.*supply|pos|supply.*state)\b/i.test(k.toLowerCase().trim())) || '';
  const cessKey = sampleKeys.find((k) => /\bcess\b/i.test(k.toLowerCase().trim())) || '';
  const currencyKey = sampleKeys.find((k) => /\b(currency|curr)\b/i.test(k.toLowerCase().trim())) || '';

  for (let idx = 0; idx < rows.length; idx++) {
    const r = rows[idx];

    const invNo = invNoKey ? String(r[invNoKey] ?? '').trim() : '';
    const rawDate = dateKey ? String(r[dateKey] ?? '').trim() : '';
    const invDate = rawDate ? parseDate(rawDate, 'DD-MM-YYYY') : '';
    const party = partyKey ? String(r[partyKey] ?? '').trim() : '';
    const gstin = gstinKey ? String(r[gstinKey] ?? '').trim() : '';
    const custGstin = custGstinKey ? String(r[custGstinKey] ?? '').trim() : '';
    const pos = posKey ? String(r[posKey] ?? '').trim() : '';
    const currency = currencyKey ? String(r[currencyKey] ?? '').trim().toUpperCase() : 'INR';

    const taxable = taxableKey ? Math.abs(parseAmount(r[taxableKey])) : 0;
    const cgst = cgstKey ? Math.abs(parseAmount(r[cgstKey])) : 0;
    const sgst = sgstKey ? Math.abs(parseAmount(r[sgstKey])) : 0;
    const igst = igstKey ? Math.abs(parseAmount(r[igstKey])) : 0;
    const cess = cessKey ? Math.abs(parseAmount(r[cessKey])) : 0;
    let total = totalKey ? Math.abs(parseAmount(r[totalKey])) : 0;

    if (total === 0 && taxable > 0) {
      total = taxable + cgst + sgst + igst + cess;
    }

    const validated = validateInvoiceRecord({
      id: `inv-${idx + 1}`,
      fileName: sourceFileName,
      invoiceNumber: invNo,
      invoiceDate: invDate || rawDate,
      supplierName: party,
      supplierGstin: gstin,
      customerGstin: custGstin,
      placeOfSupply: pos,
      taxableAmount: taxable,
      cgst,
      sgst,
      igst,
      cess,
      totalAmount: total,
      currency,
      confidenceScore: invNo && party && total > 0 ? 95 : 60,
      raw: r,
    });

    invoices.push(validated);
  }

  // Detect duplicates across the whole batch
  const dedupedInvoices = detectDuplicatesInBatch(invoices);

  const totalTaxable = dedupedInvoices.reduce((sum, i) => sum + i.taxableAmount, 0);
  const totalTax = dedupedInvoices.reduce((sum, i) => sum + i.totalTax, 0);
  const totalInvoiceAmount = dedupedInvoices.reduce((sum, i) => sum + i.totalAmount, 0);
  const verifiedCount = dedupedInvoices.filter((i) => i.status === 'Verified').length;
  const duplicateCount = dedupedInvoices.filter((i) => i.status === 'Duplicate').length;
  const needsReviewCount = dedupedInvoices.filter((i) => i.status === 'Needs Review').length;

  return {
    totalInvoices: dedupedInvoices.length,
    verifiedCount,
    needsReviewCount,
    duplicateCount,
    totalTaxable,
    totalTax,
    totalInvoiceAmount,
    invoices: dedupedInvoices,
  };
}

/**
 * Generates standard CSV export string for accounting import
 */
export function exportInvoiceCsv(invoices: ExtractedInvoiceRecord[]): string {
  const headers = [
    'Source File',
    'Invoice Number',
    'Invoice Date',
    'Supplier Name',
    'Supplier GSTIN',
    'Customer GSTIN',
    'Place of Supply',
    'Taxable Value',
    'CGST',
    'SGST',
    'IGST',
    'Cess',
    'Total Tax',
    'Total Amount',
    'Currency',
    'Status',
    'Validation Issues',
  ];

  const escapeCsv = (val: any) => {
    const s = String(val ?? '').replace(/"/g, '""');
    return `"${s}"`;
  };

  const rows = invoices.map((inv) => [
    escapeCsv(inv.fileName),
    escapeCsv(inv.invoiceNumber),
    escapeCsv(inv.invoiceDate),
    escapeCsv(inv.supplierName),
    escapeCsv(inv.supplierGstin),
    escapeCsv(inv.customerGstin || ''),
    escapeCsv(inv.placeOfSupply || ''),
    inv.taxableAmount.toFixed(2),
    inv.cgst.toFixed(2),
    inv.sgst.toFixed(2),
    inv.igst.toFixed(2),
    (inv.cess || 0).toFixed(2),
    inv.totalTax.toFixed(2),
    inv.totalAmount.toFixed(2),
    escapeCsv(inv.currency || 'INR'),
    escapeCsv(inv.status),
    escapeCsv(inv.validationIssues.join('; ')),
  ]);

  return [headers.join(','), ...rows.map((r) => r.join(','))].join('\r\n');
}

/**
 * Creates professional 4-sheet Excel workbook of extracted, verified invoices
 */
export function createInvoiceExportWorkbook(
  invoices: ExtractedInvoiceRecord[],
  batchTitle: string = 'Invoice_Extraction_Register'
): XLSX.WorkBook {
  const wb = XLSX.utils.book_new();

  // Sheet 1: Consolidated Invoices Register
  const exportRows = invoices.map((inv, idx) => ({
    'S.No': idx + 1,
    'Source File': inv.fileName,
    'Invoice Number': inv.invoiceNumber,
    'Invoice Date': inv.invoiceDate,
    'Supplier / Vendor Name': inv.supplierName,
    'Supplier GSTIN': inv.supplierGstin,
    'Customer GSTIN': inv.customerGstin || '(Not Provided)',
    'Place of Supply': inv.placeOfSupply || '(State)',
    'Taxable Value (₹)': inv.taxableAmount,
    'CGST (₹)': inv.cgst,
    'SGST (₹)': inv.sgst,
    'IGST (₹)': inv.igst,
    'Cess (₹)': inv.cess || 0,
    'Total Tax (₹)': inv.totalTax,
    'Invoice Total (₹)': inv.totalAmount,
    'Currency': inv.currency || 'INR',
    'Audit Status': inv.status,
    'Validation Notes': inv.validationIssues.join('; ') || 'All statutory checks passed',
  }));

  const wsRegister = createStyledWorksheet(exportRows, {
    currencyColumns: ['Taxable Value (₹)', 'CGST (₹)', 'SGST (₹)', 'IGST (₹)', 'Cess (₹)', 'Total Tax (₹)', 'Invoice Total (₹)'],
    dateColumns: ['Invoice Date'],
    columnWidths: {
      'S.No': 8,
      'Source File': 24,
      'Invoice Number': 18,
      'Invoice Date': 14,
      'Supplier / Vendor Name': 32,
      'Supplier GSTIN': 18,
      'Customer GSTIN': 18,
      'Place of Supply': 16,
      'Taxable Value (₹)': 18,
      'CGST (₹)': 14,
      'SGST (₹)': 14,
      'IGST (₹)': 14,
      'Cess (₹)': 12,
      'Total Tax (₹)': 16,
      'Invoice Total (₹)': 18,
      'Currency': 10,
      'Audit Status': 14,
      'Validation Notes': 45,
    },
  });
  XLSX.utils.book_append_sheet(wb, wsRegister, 'Consolidated_Invoices');

  // Sheet 2: Exception Review (Invoices requiring review)
  const exceptions = invoices.filter((i) => i.status === 'Needs Review');
  const exceptionRows = exceptions.map((f, idx) => ({
    'Flag #': idx + 1,
    'Source File': f.fileName,
    'Invoice No': f.invoiceNumber,
    'Vendor Name': f.supplierName,
    'Supplier GSTIN': f.supplierGstin || '(Missing)',
    'Invoice Total (₹)': f.totalAmount,
    'Detected Discrepancies': f.validationIssues.join('; '),
    'Recommended Action': 'Verify physical tax invoice copy or request debit note from supplier',
  }));
  const wsExceptions = createStyledWorksheet(
    exceptionRows.length > 0 ? exceptionRows : [{ Note: 'Zero invoice discrepancies detected. All invoices verified.' }],
    {
      currencyColumns: ['Invoice Total (₹)'],
      columnWidths: {
        'Flag #': 10,
        'Source File': 24,
        'Invoice No': 18,
        'Vendor Name': 30,
        'Supplier GSTIN': 18,
        'Invoice Total (₹)': 18,
        'Detected Discrepancies': 50,
        'Recommended Action': 45,
      },
    }
  );
  XLSX.utils.book_append_sheet(wb, wsExceptions, 'Exception_Review');

  // Sheet 3: Duplicate Invoices Report
  const duplicates = invoices.filter((i) => i.status === 'Duplicate');
  const dupRows = duplicates.map((d, idx) => ({
    'Duplicate #': idx + 1,
    'Source File': d.fileName,
    'Invoice No': d.invoiceNumber,
    'Vendor Name': d.supplierName,
    'Supplier GSTIN': d.supplierGstin,
    'Invoice Total (₹)': d.totalAmount,
    'Duplicate Flag Reason': d.duplicateReason || 'Duplicate bill found in batch',
    'Audit Risk': 'Risk of double payment or duplicate Input Tax Credit (ITC) claim',
  }));
  const wsDuplicates = createStyledWorksheet(
    dupRows.length > 0 ? dupRows : [{ Note: 'No duplicate invoices identified in this batch.' }],
    {
      currencyColumns: ['Invoice Total (₹)'],
      columnWidths: {
        'Duplicate #': 12,
        'Source File': 24,
        'Invoice No': 18,
        'Vendor Name': 30,
        'Supplier GSTIN': 18,
        'Invoice Total (₹)': 18,
        'Duplicate Flag Reason': 40,
        'Audit Risk': 45,
      },
    }
  );
  XLSX.utils.book_append_sheet(wb, wsDuplicates, 'Duplicates_Report');

  // Sheet 4: Executive Summary & Reconciliation
  const totalTaxable = invoices.reduce((sum, i) => sum + i.taxableAmount, 0);
  const totalCGST = invoices.reduce((sum, i) => sum + i.cgst, 0);
  const totalSGST = invoices.reduce((sum, i) => sum + i.sgst, 0);
  const totalIGST = invoices.reduce((sum, i) => sum + i.igst, 0);
  const totalCess = invoices.reduce((sum, i) => sum + (i.cess || 0), 0);
  const totalTax = invoices.reduce((sum, i) => sum + i.totalTax, 0);
  const totalAmount = invoices.reduce((sum, i) => sum + i.totalAmount, 0);

  const summaryRows = [
    { Parameter: 'Batch Identifier', Value: batchTitle },
    { Parameter: 'Total Invoices Processed', Value: invoices.length },
    { Parameter: '100% Verified Compliant Invoices', Value: invoices.filter((i) => i.status === 'Verified').length },
    { Parameter: 'Invoices Flagged for Review', Value: exceptions.length },
    { Parameter: 'Duplicate Invoices Detected', Value: duplicates.length },
    { Parameter: 'Total Cumulative Taxable Value (₹)', Value: totalTaxable },
    { Parameter: 'Total Cumulative CGST (₹)', Value: totalCGST },
    { Parameter: 'Total Cumulative SGST (₹)', Value: totalSGST },
    { Parameter: 'Total Cumulative IGST (₹)', Value: totalIGST },
    { Parameter: 'Total Cumulative Cess (₹)', Value: totalCess },
    { Parameter: 'Total Cumulative GST Tax (₹)', Value: totalTax },
    { Parameter: 'Total Cumulative Invoice Value (₹)', Value: totalAmount },
    { Parameter: 'Extraction & Verification Timestamp', Value: new Date().toLocaleString('en-IN') },
  ];
  const wsSummary = createStyledWorksheet(summaryRows, {
    currencyColumns: ['Value'],
    columnWidths: { Parameter: 40, Value: 30 },
  });
  XLSX.utils.book_append_sheet(wb, wsSummary, 'Batch_Summary');

  return wb;
}
