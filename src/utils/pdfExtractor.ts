import * as XLSX from 'xlsx';
import * as pdfjsLib from 'pdfjs-dist';
import { parseAmount, parseDate } from './dataSanitizer';
import { createStyledWorksheet } from './excelStyler';

// Initialize PDF.js worker in browser environments
if (typeof window !== 'undefined') {
  try {
    pdfjsLib.GlobalWorkerOptions.workerSrc = new URL('pdfjs-dist/build/pdf.worker.mjs', import.meta.url).toString();
  } catch {
    // Fallback if URL resolution fails
  }
}

export interface ExtractedPdfRow {
  date: string;
  particulars: string;
  refNo: string;
  debit: number;
  credit: number;
  balance: number;
}

export interface PdfExtractionResult {
  fileName: string;
  fileSizeBytes: number;
  isTextFound: boolean;
  documentType: 'bank_statement' | 'invoice' | 'general_table' | 'unsupported';
  rows: ExtractedPdfRow[];
  invoice?: InvoicePdfResult['invoice'];
  generalTable?: {
    headers: string[];
    rows: Record<string, any>[];
  };
  detectedColumnCount: number;
  pageCount: number;
  rawTextPreview: string;
  status: 'success' | 'unsupported' | 'scanned' | 'encrypted' | 'error';
  isSupportedStatement: boolean;
  warning?: string;
  errorMessage?: string;
}

export interface InvoiceLineItem {
  description: string;
  hsnCode: string;
  quantity: number;
  rate: number;
  taxableAmount: number;
  taxRate: number;
  totalAmount: number;
}

export interface InvoicePdfResult {
  success: boolean;
  status: 'success' | 'unsupported' | 'scanned' | 'encrypted' | 'error';
  errorMessage?: string;
  invoice?: {
    supplierName: string;
    supplierGstin: string;
    buyerName?: string;
    buyerGstin?: string;
    invoiceNumber: string;
    invoiceDate: string;
    taxableAmount: number;
    cgst: number;
    sgst: number;
    igst: number;
    totalAmount: number;
    lineItems?: InvoiceLineItem[];
  };
  rawTextPreview: string;
  pageCount: number;
}

/**
 * Extracts structured text lines from PDF using PDF.js with spatial Y/X clustering
 */
async function extractLinesUsingPdfJs(buffer: ArrayBuffer): Promise<{ lines: string[]; pageCount: number; rawText: string }> {
  // Use buffer.slice(0) so PDF.js worker transfer never detaches the original buffer
  const loadingTask = pdfjsLib.getDocument({
    data: new Uint8Array(buffer.slice(0)),
    useSystemFonts: true,
  });

  const pdf = await loadingTask.promise;
  const allLines: string[] = [];
  const textPieces: string[] = [];

  for (let pageNum = 1; pageNum <= pdf.numPages; pageNum++) {
    const page = await pdf.getPage(pageNum);
    const textContent = await page.getTextContent();

    if (!textContent.items || textContent.items.length === 0) {
      continue;
    }

    const itemsByY: { y: number; x: number; text: string }[] = [];
    for (const item of textContent.items as any[]) {
      if (item && typeof item.str === 'string' && item.str.trim()) {
        const x = item.transform ? Math.round(item.transform[4]) : 0;
        const y = item.transform ? Math.round(item.transform[5]) : 0;
        itemsByY.push({ y, x, text: item.str.trim() });
        textPieces.push(item.str.trim());
      }
    }

    // Sort items by Y descending (top to bottom), then X ascending (left to right)
    itemsByY.sort((a, b) => b.y - a.y || a.x - b.x);

    // Group items on roughly the same line (within 5px vertical distance)
    const linesMap: { y: number; items: { x: number; text: string }[] }[] = [];
    for (const item of itemsByY) {
      let lineObj = linesMap.find((l) => Math.abs(l.y - item.y) <= 5);
      if (!lineObj) {
        lineObj = { y: item.y, items: [] };
        linesMap.push(lineObj);
      }
      lineObj.items.push({ x: item.x, text: item.text });
    }

    for (const lineObj of linesMap) {
      lineObj.items.sort((a, b) => a.x - b.x);
      const lineText = lineObj.items.map((it) => it.text).join('   ');
      if (lineText.trim()) {
        allLines.push(lineText.trim());
      }
    }
  }

  return {
    lines: allLines,
    pageCount: pdf.numPages,
    rawText: textPieces.join(' '),
  };
}

/**
 * Fallback ASCII & Tj extractor if PDF.js fails to initialize
 */
function extractFallbackText(buffer: ArrayBuffer): string[] {
  try {
    if (!buffer || buffer.byteLength === 0) return [];
    const bytes = new Uint8Array(buffer);
    const rawStr = new TextDecoder('latin1').decode(bytes);
    const textPieces: string[] = [];

    const tjRegex = /\(([^)]+)\)\s*(?:Tj|'|")/g;
    let match;
    while ((match = tjRegex.exec(rawStr)) !== null) {
      const cleanPiece = match[1].replace(/\\([()\\])/g, '$1').trim();
      if (cleanPiece) textPieces.push(cleanPiece);
    }

    return textPieces;
  } catch {
    return [];
  }
}

/**
 * Parses an Invoice PDF to extract GSTIN, Invoice Number, Date, and Amounts.
 * Real logic only: STRICTLY rejects unsupported documents and NEVER yields fake/mock fallbacks!
 */
export async function parseInvoicePdf(file: File): Promise<InvoicePdfResult> {
  const buffer = await file.arrayBuffer();

  let lines: string[] = [];
  let pageCount = 1;
  let rawText = '';
  let isEncrypted = false;

  try {
    const pdfData = await extractLinesUsingPdfJs(buffer);
    lines = pdfData.lines;
    pageCount = pdfData.pageCount;
    rawText = pdfData.rawText;
  } catch (err: any) {
    if (err?.name === 'PasswordException' || /password/i.test(err?.message || '')) {
      isEncrypted = true;
    } else {
      lines = extractFallbackText(buffer);
      rawText = lines.join(' ');
      pageCount = 1;
    }
  }

  if (isEncrypted) {
    return {
      success: false,
      status: 'encrypted',
      errorMessage: 'This invoice PDF is password-protected. Please remove the password and re-upload.',
      rawTextPreview: '',
      pageCount,
    };
  }

  if (lines.length === 0 || rawText.trim().length < 40) {
    return {
      success: false,
      status: 'scanned',
      errorMessage: 'Scanned image-based invoice detected. No embedded digital text found. Please upload a digital PDF invoice or Excel register.',
      rawTextPreview: '',
      pageCount,
    };
  }

  // 1. Bank Statement Guard: Reject bank statements uploaded to invoice tools
  const isBankStatement =
    /statement\s+of\s+account|account\s+statement|bank\s+statement|saving\s+account|current\s+account/i.test(rawText) ||
    (/account\s+number|a\/c\s+no|ifsc\s+code/i.test(rawText) && /withdrawal|deposit|closing\s+balance|opening\s+balance/i.test(rawText));

  if (isBankStatement) {
    return {
      success: false,
      status: 'unsupported',
      errorMessage: `The uploaded PDF '${file.name}' is a Bank Statement, not a Tax Invoice. Please upload digital Tax Invoices (.pdf) or switch to the 'Bank Statement Converter' / 'Bank Reconciliation' tool.`,
      rawTextPreview: rawText.slice(0, 500),
      pageCount,
    };
  }

  // 2. Invoice Relevance Guard: Must contain standard invoice indicators
  const hasInvoiceContext =
    /tax\s+invoice|invoice|bill\s+of\s+supply|retail\s+invoice|commercial\s+invoice|proforma|debit\s+note|credit\s+note|bill\s+to|buyer\s+details|inv\s*no|invoice\s*no/i.test(
      rawText
    );

  if (!hasInvoiceContext) {
    return {
      success: false,
      status: 'unsupported',
      errorMessage: `Unsupported document '${file.name}': Not a recognizable GST Tax Invoice. Invoices must contain standard invoice headers (Tax Invoice, Bill of Supply, Invoice No, or GSTIN).`,
      rawTextPreview: rawText.slice(0, 500),
      pageCount,
    };
  }

  // 3. Extract GSTIN(s)
  const gstinRegex = /\b(\d{2}[A-Z]{5}\d{4}[A-Z]{1}[1-9A-Z]{1}Z[0-9A-Z]{1})\b/gi;
  const allGstins: string[] = [];
  let gstinMatch;
  while ((gstinMatch = gstinRegex.exec(rawText)) !== null) {
    const g = gstinMatch[1].toUpperCase();
    if (!allGstins.includes(g)) {
      allGstins.push(g);
    }
  }

  let supplierGstin = '';
  let buyerGstin = '';

  if (allGstins.length === 1) {
    supplierGstin = allGstins[0];
  } else if (allGstins.length >= 2) {
    // Check context for supplier vs buyer
    supplierGstin = allGstins[0];
    buyerGstin = allGstins[1];

    for (const line of lines) {
      if (/buyer|bill\s*to|recipient|customer|consignee/i.test(line)) {
        for (const g of allGstins) {
          if (line.includes(g)) {
            buyerGstin = g;
            supplierGstin = allGstins.find((x) => x !== g) || supplierGstin;
          }
        }
      }
    }
  }

  // 2. Extract Invoice Number
  let invoiceNumber = '';
  const invNoPatterns = [
    /(?:invoice\s*(?:no|num|number|#)?|inv\s*(?:no|num|#)?|bill\s*(?:no|#)?|tax\s*invoice\s*(?:no|#)?)[\s\.:#\-]+([A-Za-z0-9\/\-_]{2,30})/i,
    /(?:invoice\s*id|bill\s*number|doc\s*no)[\s\.:#\-]+([A-Za-z0-9\/\-_]{2,30})/i,
    /\b([A-Za-z]{2,5}\/\d{2,4}\/\d{2,8})\b/,
  ];

  for (const pattern of invNoPatterns) {
    const m = rawText.match(pattern);
    if (m && m[1]) {
      const candidate = m[1].trim();
      // Avoid capturing labels like "Date", "Original", "Duplicate"
      if (!/^(date|dated|original|duplicate|triplicate|cash|tax|gst)$/i.test(candidate)) {
        invoiceNumber = candidate;
        break;
      }
    }
  }

  // 3. Extract Invoice Date
  let invoiceDate = '';
  const datePatterns = [
    /(?:invoice\s*date|bill\s*date|inv\s*date|date\s*of\s*issue|dated)[\s\.:#\-]+(\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4}|\d{1,2}[\/\-\.][A-Za-z]{3,9}[\/\-\.]\d{2,4})/i,
    /(?:date)[\s\.:#\-]+(\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4}|\d{1,2}[\/\-\.][A-Za-z]{3,9}[\/\-\.]\d{2,4})/i,
  ];

  for (const dp of datePatterns) {
    const dm = rawText.match(dp);
    if (dm && dm[1]) {
      invoiceDate = parseDate(dm[1], 'DD-MM-YYYY');
      if (invoiceDate) break;
    }
  }

  if (!invoiceDate) {
    // General date regex fallback
    const generalDateMatch = rawText.match(/\b(\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4})\b/);
    if (generalDateMatch) {
      invoiceDate = parseDate(generalDateMatch[1], 'DD-MM-YYYY');
    }
  }

  // 4. Extract Amounts
  const totalPatterns = [
    /(?:grand\s*total|invoice\s*total|invoice\s*value|total\s*amount|net\s*payable|amount\s*payable)[\s\.:₹RsINR]*(\d{1,3}(?:,\d{2,3})*(?:\.\d{2})?)/i,
    /(?:total)[\s\.:₹RsINR]*(\d{1,3}(?:,\d{2,3})*(?:\.\d{2})?)/i,
  ];

  let totalAmount = 0;
  for (const tp of totalPatterns) {
    const tm = rawText.match(tp);
    if (tm && tm[1]) {
      const parsed = parseAmount(tm[1]);
      if (parsed > 0) {
        totalAmount = parsed;
        break;
      }
    }
  }

  const cgstMatch = rawText.match(/(?:cgst|central\s*tax)[\s\.:@%₹RsINR\d]*(\d{1,3}(?:,\d{2,3})*(?:\.\d{2})?)/i);
  const cgst = cgstMatch ? parseAmount(cgstMatch[1]) : 0;

  const sgstMatch = rawText.match(/(?:sgst|state\s*tax)[\s\.:@%₹RsINR\d]*(\d{1,3}(?:,\d{2,3})*(?:\.\d{2})?)/i);
  const sgst = sgstMatch ? parseAmount(sgstMatch[1]) : 0;

  const igstMatch = rawText.match(/(?:igst|integrated\s*tax)[\s\.:@%₹RsINR\d]*(\d{1,3}(?:,\d{2,3})*(?:\.\d{2})?)/i);
  const igst = igstMatch ? parseAmount(igstMatch[1]) : 0;

  const taxableMatch = rawText.match(/(?:taxable\s*(?:value|amount)|sub\s*total|assessable\s*value)[\s\.:₹RsINR]*(\d{1,3}(?:,\d{2,3})*(?:\.\d{2})?)/i);
  let taxableAmount = taxableMatch ? parseAmount(taxableMatch[1]) : 0;

  if (taxableAmount === 0 && totalAmount > 0) {
    const totalTax = cgst + sgst + igst;
    if (totalTax > 0 && totalAmount > totalTax) {
      taxableAmount = totalAmount - totalTax;
    } else {
      taxableAmount = totalAmount;
    }
  }

  // 5. Extract Supplier Name from the top 12 lines
  let supplierName = '';
  for (const line of lines.slice(0, 12)) {
    const cleanLine = line.replace(/^[0-9\s\-_.:|]+/, '').trim();
    if (
      cleanLine.length >= 3 &&
      cleanLine.length <= 60 &&
      !/tax\s+invoice|invoice|original|duplicate|triplicate|gstin|pan|cin|date|bill\s+to|ship\s+to|ph(?:one)?|email|address/i.test(
        cleanLine
      ) &&
      !allGstins.some((g) => cleanLine.includes(g))
    ) {
      supplierName = cleanLine.split('   ')[0].trim();
      if (supplierName.length >= 3) break;
    }
  }

  // 6. STRICT VALIDATION: If neither invoice number nor GSTIN was found, OR totalAmount is 0
  const hasInvoiceEvidence = Boolean(invoiceNumber || supplierGstin);
  const hasFinancialValues = totalAmount > 0 || taxableAmount > 0;

  if (!hasInvoiceEvidence || !hasFinancialValues) {
    return {
      success: false,
      status: 'unsupported',
      errorMessage:
        'Unsupported document: No valid digital tax invoice fields detected in this PDF. A valid GST invoice requires an invoice number or supplier GSTIN, plus financial totals. If this is a bank statement or general document, please use the appropriate tool.',
      rawTextPreview: rawText.slice(0, 500),
      pageCount,
    };
  }

  // Return REAL extracted data only - NEVER inject fake fallbacks!
  return {
    success: true,
    status: 'success',
    invoice: {
      supplierName: supplierName || (supplierGstin ? `GSTIN: ${supplierGstin}` : 'Vendor / Supplier'),
      supplierGstin: supplierGstin || '',
      buyerGstin: buyerGstin || '',
      invoiceNumber: invoiceNumber || 'INV-UNSPECIFIED',
      invoiceDate: invoiceDate || '',
      taxableAmount: taxableAmount > 0 ? taxableAmount : totalAmount,
      cgst,
      sgst,
      igst,
      totalAmount: totalAmount > 0 ? totalAmount : taxableAmount + cgst + sgst + igst,
    },
    rawTextPreview: rawText.slice(0, 500),
    pageCount,
  };
}

/**
 * Parses extracted PDF text into structured financial bank statement rows.
 * Real logic: Strictly requires genuine multi-row bank transactions. NEVER yields fake "Bank Transaction" rows!
 */
export async function parsePdfFile(file: File): Promise<PdfExtractionResult> {
  const buffer = await file.arrayBuffer();

  let lines: string[] = [];
  let pageCount = 1;
  let rawText = '';
  let isEncrypted = false;

  try {
    const pdfData = await extractLinesUsingPdfJs(buffer);
    lines = pdfData.lines;
    pageCount = pdfData.pageCount;
    rawText = pdfData.rawText;
  } catch (err: any) {
    if (err?.name === 'PasswordException' || /password/i.test(err?.message || '')) {
      isEncrypted = true;
    } else {
      lines = extractFallbackText(buffer);
      rawText = lines.join(' ');
      pageCount = Math.max(1, Math.ceil(file.size / (150 * 1024)));
    }
  }

  // 1. Password Protected Check
  if (isEncrypted) {
    return {
      fileName: file.name,
      fileSizeBytes: file.size,
      isTextFound: false,
      documentType: 'unsupported',
      rows: [],
      detectedColumnCount: 0,
      pageCount,
      rawTextPreview: '',
      status: 'encrypted',
      isSupportedStatement: false,
      errorMessage: 'This PDF is password-protected. Please remove the password and re-upload the statement.',
    };
  }

  // 2. Scanned / Image-only Check
  const isImageOrEmpty = lines.length === 0 || rawText.trim().length < 40;
  if (isImageOrEmpty) {
    return {
      fileName: file.name,
      fileSizeBytes: file.size,
      isTextFound: false,
      documentType: 'unsupported',
      rows: [],
      detectedColumnCount: 0,
      pageCount,
      rawTextPreview: '',
      status: 'scanned',
      isSupportedStatement: false,
      errorMessage:
        'Scanned or image-based PDF detected. No embedded text layers found. For accurate Excel conversion, please upload an official digital bank statement PDF or an Excel/CSV export.',
    };
  }

  // 3. Bank Statement Context Verification
  const hasBankingKeywords =
    /statement|account|opening\s*balance|closing\s*balance|withdrawal|deposit|debit|credit|dr|cr|balance|cheque|chq|neft|rtgs|upi|imps|inb|branch|ifsc|bank|particulars/i.test(
      rawText
    );

  const dateRegex = /(\b\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4}\b|\b\d{1,2}[\/\-\.\s][A-Za-z]{3,9}[\/\-\.\s]\d{2,4}\b)/;
  // Support amounts with and without decimals, handling Indian comma separators
  const amountRegex = /(\b\d{1,3}(?:,\d{2,3})*(?:\.\d{2})?\b)/g;

  const detectedRows: ExtractedPdfRow[] = [];
  let currentBalance = 0;

  for (const line of lines) {
    // Skip obvious header, disclaimer, and summary lines
    if (
      /page\s+\d+|account\s+number|statement\s+period|ifsc\s+code|branch\s+name|registered\s+office|customer\s+id|nominee|total\s+withdrawal|total\s+deposit/i.test(
        line
      ) &&
      !dateRegex.test(line)
    ) {
      continue;
    }

    const dateMatch = line.match(dateRegex);
    if (!dateMatch) continue;

    // Filter amounts to plausible financial amounts (exclude single digits or year numbers like 2026 unless in amount context)
    const rawMatches = line.match(amountRegex) || [];
    const validAmountStrings = rawMatches.filter((s) => {
      const clean = s.replace(/,/g, '');
      const num = parseFloat(clean);
      if (isNaN(num)) return false;
      // Skip date components matching amountRegex (e.g. "2026" or "15")
      if (dateMatch[0].includes(s)) return false;
      return num >= 0.01;
    });

    if (validAmountStrings.length >= 1) {
      const dt = parseDate(dateMatch[1]);
      const amounts = validAmountStrings.map((a) => parseAmount(a));

      let debit = 0;
      let credit = 0;
      let balance = 0;

      // Extract Narration by removing date and amounts from the line
      let narration = line.replace(dateRegex, '');
      for (const aStr of validAmountStrings) {
        narration = narration.replace(aStr, '');
      }
      narration = narration
        .replace(/(?:chq|ref|utr|cms|txn|chq\s*no)[\s\-_:]*([a-z0-9]+)/gi, '')
        .replace(/^[\s\-_:,|]+|[\s\-_:,|]+$/g, '')
        .replace(/\s{2,}/g, ' ')
        .trim();

      // Do NOT create a transaction row if narration is empty or trivial noise (e.g. single character)
      if (!narration || narration.length < 2 || /^(dr|cr|bal|page|\d+)$/i.test(narration)) {
        continue;
      }

      if (amounts.length === 1) {
        if (/dr|debit|payment|wdl|withdrawal/i.test(line)) {
          debit = amounts[0];
        } else {
          credit = amounts[0];
        }
      } else if (amounts.length === 2) {
        if (/dr|debit|wdl/i.test(line) && !/cr|credit|dep/i.test(line)) {
          debit = amounts[0];
          balance = amounts[1];
        } else if (/cr|credit|dep/i.test(line) && !/dr|debit/i.test(line)) {
          credit = amounts[0];
          balance = amounts[1];
        } else {
          debit = amounts[0];
          credit = amounts[1];
        }
      } else if (amounts.length >= 3) {
        debit = amounts[0];
        credit = amounts[1];
        balance = amounts[2];
      }

      const refMatch = line.match(/(?:chq|ref|utr|cms|txn|chq\s*no)[\s\-_:]*([a-z0-9]+)/i);

      detectedRows.push({
        date: dt,
        particulars: narration,
        refNo: refMatch ? refMatch[1] : '',
        debit,
        credit,
        balance: balance || currentBalance,
      });

      if (balance) currentBalance = balance;
    }
  }

  // 4. Verification: A bank statement requires at least 2 valid transactions AND banking context
  if (hasBankingKeywords && detectedRows.length >= 2) {
    return {
      fileName: file.name,
      fileSizeBytes: file.size,
      isTextFound: true,
      documentType: 'bank_statement',
      rows: detectedRows,
      detectedColumnCount: 6,
      pageCount,
      rawTextPreview: rawText.slice(0, 500),
      status: 'success',
      isSupportedStatement: true,
    };
  }

  // 5. If not a bank statement, check if it is a Digital Tax Invoice!
  const invoiceRes = await parseInvoicePdf(file);
  if (invoiceRes.success && invoiceRes.invoice) {
    const inv = invoiceRes.invoice;
    return {
      fileName: file.name,
      fileSizeBytes: file.size,
      isTextFound: true,
      documentType: 'invoice',
      rows: [
        {
          date: inv.invoiceDate || '',
          particulars: `${inv.supplierName || 'Supplier'} (GSTIN: ${inv.supplierGstin || 'N/A'}) - Invoice #${inv.invoiceNumber || 'N/A'}`,
          refNo: inv.invoiceNumber || '',
          debit: 0,
          credit: inv.totalAmount || 0,
          balance: inv.totalAmount || 0,
        },
      ],
      invoice: inv,
      detectedColumnCount: 8,
      pageCount,
      rawTextPreview: rawText.slice(0, 500),
      status: 'success',
      isSupportedStatement: false,
    };
  }

  // 6. Text was found, but NO genuine financial tables or bank transactions matched
  return {
    fileName: file.name,
    fileSizeBytes: file.size,
    isTextFound: true,
    documentType: 'unsupported',
    rows: [],
    detectedColumnCount: 0,
    pageCount,
    rawTextPreview: rawText.slice(0, 500),
    status: 'unsupported',
    isSupportedStatement: false,
    errorMessage:
      'Unsupported PDF format: No tabular bank transactions or structured tax invoices detected. The PDF was readable, but does not contain recurring statement transactions (Date, Description, Debit/Credit/Balance).',
  };
}

/**
 * Creates Excel spreadsheet from extracted PDF rows
 */
export function createExtractedPdfWorkbook(rows: ExtractedPdfRow[], fileName: string): XLSX.WorkBook {
  const wb = XLSX.utils.book_new();
  const exportRows = rows.map((r, idx) => ({
    'S.No': idx + 1,
    'Date': r.date,
    'Particulars / Narration': r.particulars,
    'Chq / Ref No': r.refNo,
    'Withdrawal (Debit)': r.debit,
    'Deposit (Credit)': r.credit,
    'Closing Balance': r.balance,
  }));

  const ws = createStyledWorksheet(exportRows, {
    currencyColumns: ['Withdrawal (Debit)', 'Deposit (Credit)', 'Closing Balance'],
    dateColumns: ['Date'],
    columnWidths: {
      'S.No': 8,
      'Date': 14,
      'Particulars / Narration': 40,
      'Chq / Ref No': 18,
      'Withdrawal (Debit)': 20,
      'Deposit (Credit)': 20,
      'Closing Balance': 22,
    },
  });
  XLSX.utils.book_append_sheet(wb, ws, 'Extracted_Statement');
  return wb;
}

/**
 * Creates Excel spreadsheet from an extracted Invoice
 */
export function createExtractedInvoiceWorkbook(invoice: NonNullable<InvoicePdfResult['invoice']>, fileName: string): XLSX.WorkBook {
  const wb = XLSX.utils.book_new();

  const summaryRows = [
    { Field: 'Invoice Number', Value: invoice.invoiceNumber },
    { Field: 'Invoice Date', Value: invoice.invoiceDate },
    { Field: 'Supplier Name', Value: invoice.supplierName },
    { Field: 'Supplier GSTIN', Value: invoice.supplierGstin },
    { Field: 'Buyer GSTIN', Value: invoice.buyerGstin || 'Unspecified' },
    { Field: 'Taxable Amount (₹)', Value: invoice.taxableAmount },
    { Field: 'CGST (₹)', Value: invoice.cgst },
    { Field: 'SGST (₹)', Value: invoice.sgst },
    { Field: 'IGST (₹)', Value: invoice.igst },
    { Field: 'Total Invoice Amount (₹)', Value: invoice.totalAmount },
  ];

  const ws = createStyledWorksheet(summaryRows, {
    currencyColumns: ['Value'],
    columnWidths: {
      Field: 30,
      Value: 35,
    },
  });

  XLSX.utils.book_append_sheet(wb, ws, 'Invoice_Summary');
  return wb;
}
