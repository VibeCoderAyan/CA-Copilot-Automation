import JSZip from 'jszip';
import * as XLSX from 'xlsx';
import { createStyledWorksheet } from './excelStyler';

export type AuditDocCategory =
  | 'Bank Statements'
  | 'Sales Invoices'
  | 'Purchase Invoices'
  | 'Expense Vouchers'
  | 'Ledger Accounts'
  | 'Trial Balances'
  | 'Tax Records'
  | 'Supporting Schedules';

export interface AuditChecklistItem {
  id: string;
  category: AuditDocCategory;
  title: string;
  description: string;
  requiredFor: string;
  isFound: boolean;
  matchedFilesCount: number;
}

export interface FileItemToRename {
  id: string;
  file: File;
  originalName: string;
  originalExtension: string;
  sizeBytes: number;
  detectedDate: string;
  detectedVendor: string;
  detectedInvoiceNo: string;
  detectedCategory: AuditDocCategory;
  categoryCode: string;
  evidenceIndex: string; // e.g. "WP-BNK-001"
  auditArea: string;
  reviewerNotes: string;
  fileHash?: string;
  isDuplicate?: boolean;
  duplicateOf?: string;
  newName: string;
  proposedName?: string;
  customOverride?: string;
  hasCollision?: boolean;
  status: 'Verified & Indexed' | 'Needs Review' | 'Duplicate Excluded';
  confidenceScore: number;
}

export type CaseOption = 'original' | 'uppercase' | 'lowercase' | 'titlecase' | 'camelcase';
export type SpaceOption = 'keep' | 'underscore' | 'hyphen' | 'remove';

export interface RenamerConfig {
  pattern: string; // e.g. "{evidenceIndex}_{client}_{category}_{date}"
  prefix: string;
  suffix: string;
  clientName: string;
  clientCode?: string;
  dateFormat: 'YYYY-MM-DD' | 'DD-MM-YYYY' | 'YYYYMMDD' | 'DD_MMM_YYYY';
  caseOption: CaseOption;
  spaceOption: SpaceOption;
  seqStart: number;
  seqPadding: number;
  autoResolveDuplicates: boolean;
  organizeIntoFolders: boolean;
}

export const DEFAULT_RENAMER_CONFIG: RenamerConfig = {
  pattern: '{evidenceIndex}_{client}_{category}_{date}',
  prefix: '',
  suffix: '',
  clientName: 'Apex_Enterprises',
  dateFormat: 'YYYY-MM-DD',
  caseOption: 'original',
  spaceOption: 'underscore',
  seqStart: 1,
  seqPadding: 3,
  autoResolveDuplicates: true,
  organizeIntoFolders: true,
};

export const STANDARD_AUDIT_CHECKLIST: AuditChecklistItem[] = [
  { id: 'chk-1', category: 'Bank Statements', title: 'Bank Statements (12 Months)', description: 'Complete chronological bank statements for the entire financial year', requiredFor: 'Bank Balance & BRS Verification', isFound: false, matchedFilesCount: 0 },
  { id: 'chk-2', category: 'Tax Records', title: 'GSTR-3B Filed Return Acknowledgements', description: 'Monthly/quarterly filed GSTR-3B challans and return summaries', requiredFor: 'Statutory GST Liability Verification', isFound: false, matchedFilesCount: 0 },
  { id: 'chk-3', category: 'Tax Records', title: 'GSTR-1 Filed Return Extracts', description: 'Outward supplies outward declaration filed under Section 37', requiredFor: 'Revenue & Turnover Reconciliation', isFound: false, matchedFilesCount: 0 },
  { id: 'chk-4', category: 'Tax Records', title: 'Form 26AS & AIS / TIS Statement', description: 'Annual Information Statement & 26AS TDS tax credit statement', requiredFor: 'Income Tax Credit & Advance Tax Reconciliation', isFound: false, matchedFilesCount: 0 },
  { id: 'chk-5', category: 'Trial Balances', title: 'Final Signed Trial Balance', description: 'Closing pre-audit trial balance approved by management', requiredFor: 'Schedule III Statement Compilation', isFound: false, matchedFilesCount: 0 },
  { id: 'chk-6', category: 'Supporting Schedules', title: 'Fixed Asset Register (FAR)', description: 'Schedule of capital additions, deletions, and Section 32 depreciation', requiredFor: 'PPE & Depreciation Schedule Testing', isFound: false, matchedFilesCount: 0 },
  { id: 'chk-7', category: 'Ledger Accounts', title: 'Director & Related Party Ledgers', description: 'Section 188 related party ledger extracts and drawings accounts', requiredFor: 'CARO 2020 & Related Party Disclosures', isFound: false, matchedFilesCount: 0 },
  { id: 'chk-8', category: 'Supporting Schedules', title: 'Statutory Audit Engagement Letter', description: 'Signed engagement letter defining audit scope under SA 210', requiredFor: 'ICAI Audit Standards Compliance', isFound: false, matchedFilesCount: 0 },
];

/**
 * Maps category to 3-letter ICAI Working Paper Code
 */
export function getCategoryCode(category: AuditDocCategory): string {
  switch (category) {
    case 'Bank Statements': return 'BNK';
    case 'Sales Invoices': return 'SAL';
    case 'Purchase Invoices': return 'PUR';
    case 'Expense Vouchers': return 'EXP';
    case 'Ledger Accounts': return 'LDG';
    case 'Trial Balances': return 'TBS';
    case 'Tax Records': return 'TAX';
    case 'Supporting Schedules': return 'SCH';
    default: return 'DOC';
  }
}

/**
 * Intelligent ICAI Document Classifier based on filename & metadata
 */
export function classifyAuditDocument(filename: string): {
  category: AuditDocCategory;
  confidence: number;
  auditArea: string;
} {
  const f = filename.toLowerCase();

  // 1. Bank Statements
  if (f.includes('statement') || f.includes('bank') || f.includes('hdfc') || f.includes('sbi') || f.includes('icici') || f.includes('axis') || f.includes('kotak') || f.includes('passbook')) {
    return { category: 'Bank Statements', confidence: 95, auditArea: 'Cash & Bank Balances' };
  }

  // 2. Tax Records
  if (f.includes('gstr') || f.includes('26as') || f.includes('ais') || f.includes('tis') || f.includes('challan') || f.includes('tds') || f.includes('advance tax') || f.includes('gst')) {
    return { category: 'Tax Records', confidence: 94, auditArea: 'Statutory Dues & Taxation' };
  }

  // 3. Trial Balance
  if (f.includes('trial') || f.includes(' tb') || f.includes('_tb') || f.includes('-tb') || f.includes('trial_balance')) {
    return { category: 'Trial Balances', confidence: 98, auditArea: 'General Ledger Integrity' };
  }

  // 4. Ledger
  if (f.includes('ledger') || f.includes('daybook') || f.includes('cashbook') || f.includes('account statement')) {
    return { category: 'Ledger Accounts', confidence: 92, auditArea: 'Trade & Financial Ledgers' };
  }

  // 5. Sales Invoices
  if (f.includes('sale') || f.includes('outward') || f.includes('customer bill') || f.includes('client invoice')) {
    return { category: 'Sales Invoices', confidence: 90, auditArea: 'Revenue & Debtors Testing' };
  }

  // 6. Purchase Invoices
  if (f.includes('purchase') || f.includes('supplier bill') || f.includes('vendor bill') || f.includes('inward')) {
    return { category: 'Purchase Invoices', confidence: 90, auditArea: 'Purchases & Creditors Testing' };
  }

  // 7. Expense Vouchers
  if (f.includes('voucher') || f.includes('expense') || f.includes('petty') || f.includes('reimbursement') || f.includes('receipt')) {
    return { category: 'Expense Vouchers', confidence: 88, auditArea: 'Administrative & Operating Costs' };
  }

  // 8. Supporting Schedules
  if (f.includes('far') || f.includes('asset') || f.includes('schedule') || f.includes('confirmation') || f.includes('resolution') || f.includes('minutes') || f.includes('agreement')) {
    return { category: 'Supporting Schedules', confidence: 85, auditArea: 'Statutory Disclosures & Schedules' };
  }

  // Default to Purchase Invoices if invoice keyword found
  if (f.includes('inv') || f.includes('bill')) {
    return { category: 'Purchase Invoices', confidence: 75, auditArea: 'Purchases & Payables' };
  }

  return { category: 'Supporting Schedules', confidence: 60, auditArea: 'Audit Working Papers' };
}

/**
 * Extracts metadata hints from filename
 */
export function extractFileMetadata(filename: string): {
  date: string;
  vendor: string;
  invoiceNo: string;
  category: AuditDocCategory;
} {
  const base = filename.replace(/\.[^/.]+$/, '');

  let date = new Date().toISOString().split('T')[0];
  const dateMatch1 = base.match(/\b(20\d{2}[-_/]\d{1,2}[-_/]\d{1,2})\b/);
  const dateMatch2 = base.match(/\b(\d{1,2}[-_/]\d{1,2}[-_/]20\d{2})\b/);
  const dateMatch3 = base.match(/\b(20\d{6})\b/);
  const monthMatch = base.match(/(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[-_]?20\d{2}/i);

  if (dateMatch1) {
    date = dateMatch1[1].replace(/[/_]/g, '-');
  } else if (dateMatch2) {
    const parts = dateMatch2[1].split(/[-_/]/);
    date = `${parts[2]}-${parts[1].padStart(2, '0')}-${parts[0].padStart(2, '0')}`;
  } else if (dateMatch3) {
    const s = dateMatch3[1];
    date = `${s.substring(0, 4)}-${s.substring(4, 6)}-${s.substring(6, 8)}`;
  } else if (monthMatch) {
    const monthsMap: Record<string, string> = {
      jan: '01', feb: '02', mar: '03', apr: '04', may: '05', jun: '06',
      jul: '07', aug: '08', sep: '09', oct: '10', nov: '11', dec: '12'
    };
    const monStr = monthMatch[1].toLowerCase().substring(0, 3);
    const mm = monthsMap[monStr] || '01';
    const yrMatch = monthMatch[0].match(/20\d{2}/);
    const yyyy = yrMatch ? yrMatch[0] : new Date().getFullYear().toString();
    date = `${yyyy}-${mm}-01`;
  }

  let invoiceNo = '';
  const invMatch = base.match(/(?:INV|BILL|INVOICE|REC|VCH)[-_#]?\s*([A-Za-z0-9-]+)/i);
  const numHashMatch = base.match(/#([0-9]{3,8})/);
  if (invMatch) {
    invoiceNo = invMatch[0].replace(/\s+/g, '_').toUpperCase();
  } else if (numHashMatch) {
    invoiceNo = `INV-${numHashMatch[1]}`;
  } else {
    const digits = base.match(/\b\d{3,6}\b/);
    invoiceNo = digits ? `INV-${digits[0]}` : '';
  }

  let vendor = '';
  const commonVendors = ['Apollo', 'TotalLubes', 'AWS', 'Google', 'ShreeKrishna', 'Reliance', 'TCS', 'Infosys', 'Tata', 'Airtel', 'Vodafone', 'Adobe'];
  for (const v of commonVendors) {
    if (new RegExp(v, 'i').test(base)) {
      vendor = v;
      break;
    }
  }
  if (!vendor) {
    const cleanWords = base.split(/[-_\s]+/).filter((w) => w.length > 2 && !/^\d+$/.test(w) && !/^(inv|bill|doc|file|scan|pdf|image|img)/i.test(w));
    if (cleanWords.length > 0) {
      vendor = cleanWords[0].charAt(0).toUpperCase() + cleanWords[0].slice(1).toLowerCase();
    } else {
      vendor = 'Client_Entity';
    }
  }

  const classified = classifyAuditDocument(filename);

  return { date, vendor, invoiceNo, category: classified.category };
}

/**
 * Evaluates audit checklist completeness against uploaded files
 */
export function evaluateAuditChecklist(items: FileItemToRename[]): AuditChecklistItem[] {
  return STANDARD_AUDIT_CHECKLIST.map((chk) => {
    let matchedCount = 0;

    for (const item of items) {
      if (item.status === 'Duplicate Excluded') continue;
      const f = item.originalName.toLowerCase();

      if (chk.id === 'chk-1' && item.detectedCategory === 'Bank Statements') matchedCount++;
      else if (chk.id === 'chk-2' && (f.includes('3b') || f.includes('gstr-3b') || f.includes('gstr3b'))) matchedCount++;
      else if (chk.id === 'chk-3' && (f.includes('gstr-1') || f.includes('gstr1'))) matchedCount++;
      else if (chk.id === 'chk-4' && (f.includes('26as') || f.includes('ais') || f.includes('tis'))) matchedCount++;
      else if (chk.id === 'chk-5' && item.detectedCategory === 'Trial Balances') matchedCount++;
      else if (chk.id === 'chk-6' && (f.includes('far') || f.includes('fixed asset') || f.includes('asset register'))) matchedCount++;
      else if (chk.id === 'chk-7' && (f.includes('director') || f.includes('related') || f.includes('drawing'))) matchedCount++;
      else if (chk.id === 'chk-8' && (f.includes('engagement') || f.includes('sa210') || f.includes('appointment'))) matchedCount++;
    }

    return {
      ...chk,
      isFound: matchedCount > 0,
      matchedFilesCount: matchedCount,
    };
  });
}

function formatDate(isoDate: string, format: RenamerConfig['dateFormat']): string {
  if (!isoDate) return '2026-10-17';
  const parts = isoDate.split('-');
  if (parts.length < 3) return isoDate;
  const year = parts[0];
  const month = parts[1];
  const day = parts[2];

  const monthNames = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  const monthIndex = parseInt(month, 10) - 1;
  const monthStr = monthNames[monthIndex] || month;

  switch (format) {
    case 'DD-MM-YYYY': return `${day}-${month}-${year}`;
    case 'YYYYMMDD': return `${year}${month}${day}`;
    case 'DD_MMM_YYYY': return `${day}_${monthStr}_${year}`;
    case 'YYYY-MM-DD':
    default: return `${year}-${month}-${day}`;
  }
}

function applyTransforms(text: string, caseOpt: CaseOption, spaceOpt: SpaceOption): string {
  let result = text;

  switch (spaceOpt) {
    case 'underscore': result = result.replace(/[\s-]+/g, '_'); break;
    case 'hyphen': result = result.replace(/[\s_]+/g, '-'); break;
    case 'remove': result = result.replace(/[\s_-]+/g, ''); break;
    case 'keep': default: break;
  }

  switch (caseOpt) {
    case 'uppercase': return result.toUpperCase();
    case 'lowercase': return result.toLowerCase();
    case 'titlecase':
      return result.replace(/\b\w/g, (c) => c.toUpperCase());
    case 'camelcase':
      return result
        .replace(/[-_\s]+(.)?/g, (_, c) => (c ? c.toUpperCase() : ''))
        .replace(/^[A-Z]/, (c) => c.toLowerCase());
    case 'original': default: return result;
  }
}

/**
 * Computes preview of new standardized names and assign Working Paper indexing
 */
export function computeRenamedFiles(
  items: FileItemToRename[],
  config?: Partial<RenamerConfig>
): FileItemToRename[] {
  const activeConfig: RenamerConfig = {
    ...DEFAULT_RENAMER_CONFIG,
    ...(config || {}),
    clientName: config?.clientName || config?.clientCode || DEFAULT_RENAMER_CONFIG.clientName,
  };
  const categoryIndices = new Map<string, number>();
  const seenSizes = new Map<number, string>();
  const nameCounts = new Map<string, number>();
  const safeItems = Array.isArray(items) ? items : [];

  return safeItems.map((item) => {
    // 1. Assign Evidence Index e.g. WP-BNK-001
    const catCode = item.categoryCode || getCategoryCode(item.detectedCategory);
    const currIdx = (categoryIndices.get(catCode) || 0) + 1;
    categoryIndices.set(catCode, currIdx);
    const evidenceIndex = `WP-${catCode}-${String(currIdx).padStart(3, '0')}`;

    // 2. Duplicate detection by size and name
    let isDuplicate = false;
    let duplicateOf = undefined;
    if (seenSizes.has(item.sizeBytes) && seenSizes.get(item.sizeBytes) !== item.originalName) {
      isDuplicate = true;
      duplicateOf = seenSizes.get(item.sizeBytes);
    } else {
      seenSizes.set(item.sizeBytes, item.originalName);
    }

    if (item.customOverride && item.customOverride.trim() !== '') {
      return {
        ...item,
        evidenceIndex,
        isDuplicate,
        duplicateOf,
        newName: item.customOverride.trim(),
        proposedName: item.customOverride.trim(),
        status: isDuplicate ? 'Duplicate Excluded' : 'Verified & Indexed',
      };
    }

    const formattedDate = formatDate(item.detectedDate, activeConfig.dateFormat);
    const catClean = (item.detectedCategory || 'General').replace(/\s+/g, '_');
    const pattern = activeConfig.pattern || DEFAULT_RENAMER_CONFIG.pattern;

    let computed = pattern
      .replace(/{evidenceIndex}/g, evidenceIndex)
      .replace(/{date}/g, formattedDate)
      .replace(/{vendor}/g, item.detectedVendor || 'Vendor')
      .replace(/{invNo}/g, item.detectedInvoiceNo || 'Bill')
      .replace(/{client}/g, activeConfig.clientName || 'Client')
      .replace(/{category}/g, catClean)
      .replace(/{original}/g, (item.originalName || '').replace(/\.[^/.]+$/, ''));

    if (activeConfig.prefix && !pattern.includes('{prefix}')) {
      computed = `${activeConfig.prefix}_${computed}`;
    }
    if (activeConfig.suffix && !pattern.includes('{suffix}')) {
      computed = `${computed}_${activeConfig.suffix}`;
    }

    computed = applyTransforms(computed, activeConfig.caseOption, activeConfig.spaceOption);
    let targetFilename = `${computed}.${item.originalExtension}`;

    // Collision Resolution
    const lowerKey = targetFilename.toLowerCase();
    const existingCount = nameCounts.get(lowerKey) || 0;
    nameCounts.set(lowerKey, existingCount + 1);

    let status: FileItemToRename['status'] = isDuplicate ? 'Duplicate Excluded' : (item.confidenceScore < 70 ? 'Needs Review' : 'Verified & Indexed');
    let hasCollision = false;

    if (existingCount > 0) {
      if (activeConfig.autoResolveDuplicates) {
        targetFilename = `${computed}_${existingCount + 1}.${item.originalExtension}`;
      } else {
        hasCollision = true;
      }
    }

    return {
      ...item,
      categoryCode: catCode,
      evidenceIndex,
      isDuplicate,
      duplicateOf,
      newName: targetFilename,
      proposedName: targetFilename,
      hasCollision,
      status,
    };
  });
}

/**
 * Creates structured ZIP archive organizing files into categorized audit folders
 */
export async function generateRenamedZip(
  items: FileItemToRename[],
  config: RenamerConfig = DEFAULT_RENAMER_CONFIG
): Promise<Blob> {
  const zip = new JSZip();

  const folderCategoryMap: Record<AuditDocCategory, string> = {
    'Bank Statements': '01_Bank_Statements',
    'Sales Invoices': '02_Sales_Invoices',
    'Purchase Invoices': '03_Purchase_Invoices',
    'Expense Vouchers': '04_Expense_Vouchers',
    'Ledger Accounts': '05_Ledgers',
    'Trial Balances': '06_Trial_Balances',
    'Tax Records': '07_Tax_Records',
    'Supporting Schedules': '08_Supporting_Schedules',
  };

  for (const item of items) {
    const arrayBuffer = await item.file.arrayBuffer();
    if (config.organizeIntoFolders) {
      const folderName = folderCategoryMap[item.detectedCategory] || '08_Supporting_Schedules';
      zip.folder(folderName)?.file(item.newName, arrayBuffer);
    } else {
      zip.file(item.newName, arrayBuffer);
    }
  }

  // Include detailed audit manifest
  let manifestText = `================================================================================\n`;
  manifestText += `ICAI AUDIT FILE AUTO-PROCESSING & WORKING PAPER EVIDENCE INDEX\n`;
  manifestText += `Client: ${config.clientName} | Timestamp: ${new Date().toLocaleString('en-IN')}\n`;
  manifestText += `Total Files Indexed: ${items.length}\n`;
  manifestText += `================================================================================\n\n`;
  manifestText += `INDEX        | CATEGORY             | ORIGINAL FILENAME                 | NEW STANDARDIZED NAME\n`;
  manifestText += `--------------------------------------------------------------------------------\n`;

  for (const item of items) {
    manifestText += `${item.evidenceIndex.padEnd(12)} | ${item.detectedCategory.padEnd(20)} | ${item.originalName.padEnd(33)} -> ${item.newName}\n`;
  }

  zip.file('00_AUDIT_EVIDENCE_INDEX_MANIFEST.txt', manifestText);

  return await zip.generateAsync({
    type: 'blob',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 },
  });
}

/**
 * Generates an Excel audit evidence register and missing document checklist
 */
export function createRenamingManifestWorkbook(
  items: FileItemToRename[],
  checklist: AuditChecklistItem[]
): XLSX.WorkBook {
  const wb = XLSX.utils.book_new();

  // Sheet 1: Audit Evidence Register
  const manifestData = items.map((item, index) => ({
    'S.No': index + 1,
    'WP Evidence Index': item.evidenceIndex,
    'Document Category': item.detectedCategory,
    'Original File Name': item.originalName,
    'New Standardized Name': item.newName,
    'Audit Area': item.auditArea,
    'Reviewer Working Paper Notes': item.reviewerNotes || 'Verified against underlying vouchers',
    'Size (KB)': Math.round(item.sizeBytes / 1024),
    'Detected Date': item.detectedDate,
    'Classification Confidence': `${item.confidenceScore}%`,
    'Status': item.status,
  }));

  const ws = createStyledWorksheet(manifestData, {
    columnWidths: {
      'S.No': 8,
      'WP Evidence Index': 20,
      'Document Category': 22,
      'Original File Name': 35,
      'New Standardized Name': 45,
      'Audit Area': 26,
      'Reviewer Working Paper Notes': 36,
      'Size (KB)': 12,
      'Detected Date': 16,
      'Classification Confidence': 24,
      'Status': 20,
    },
  });
  XLSX.utils.book_append_sheet(wb, ws, 'Audit_Evidence_Index');

  // Sheet 2: Missing Documents Checklist
  const chkData = checklist.map((chk, idx) => ({
    'Checklist #': idx + 1,
    'Required Audit Document': chk.title,
    'Category': chk.category,
    'Mandatory Audit Purpose': chk.requiredFor,
    'Document Description': chk.description,
    'Status': chk.isFound ? `Provided (${chk.matchedFilesCount} files)` : 'MISSING / PENDING FROM CLIENT',
    'Audit Impact': chk.isFound ? 'Satisfied' : 'Qualification / Audit Query to Management',
  }));
  const wsChk = createStyledWorksheet(chkData, {
    columnWidths: {
      'Checklist #': 12,
      'Required Audit Document': 32,
      'Category': 22,
      'Mandatory Audit Purpose': 32,
      'Document Description': 45,
      'Status': 30,
      'Audit Impact': 36,
    },
  });
  XLSX.utils.book_append_sheet(wb, wsChk, 'Missing_Documents_Checklist');

  // Sheet 3: Duplicates Report
  const dups = items.filter((i) => i.isDuplicate || i.status === 'Duplicate Excluded');
  const dupData = dups.map((d, idx) => ({
    'Duplicate #': idx + 1,
    'Evidence Index': d.evidenceIndex,
    'Filename': d.originalName,
    'File Size (Bytes)': d.sizeBytes,
    'Duplicate of Original': d.duplicateOf || 'Identical file content detected',
    'Action Taken': 'Excluded from audit workpapers to avoid double counting',
  }));
  const wsDups = createStyledWorksheet(
    dupData.length > 0 ? dupData : [{ Note: 'Zero duplicate files identified.' }],
    {
      columnWidths: { 'Duplicate #': 12, 'Evidence Index': 18, Filename: 35, 'File Size (Bytes)': 18, 'Duplicate of Original': 35, 'Action Taken': 45 },
    }
  );
  XLSX.utils.book_append_sheet(wb, wsDups, 'Duplicate_Files_Report');

  return wb;
}

/**
 * Generates realistic sample financial documents for testing Tool 5
 */
export function generateSampleFilesForRenamer(): FileItemToRename[] {
  const samples = [
    { name: 'hdfc_bank_statement_q2_fy26.xlsx', size: 58210, date: '2026-09-30', vendor: 'HDFC_Bank', inv: 'STMT-SEP26', cat: 'Bank Statements' as AuditDocCategory, conf: 98 },
    { name: 'gstr3b_september_2026_filed_ack.pdf', size: 84200, date: '2026-10-18', vendor: 'GST_Portal', inv: 'ACK-3B-09', cat: 'Tax Records' as AuditDocCategory, conf: 96 },
    { name: 'form_26as_ay2027_q2.pdf', size: 142100, date: '2026-09-30', vendor: 'IT_Department', inv: '26AS-Q2', cat: 'Tax Records' as AuditDocCategory, conf: 95 },
    { name: 'trial_balance_30_sept_2026.xlsx', size: 76500, date: '2026-09-30', vendor: 'Accounts_Dept', inv: 'TB-SEP26', cat: 'Trial Balances' as AuditDocCategory, conf: 99 },
    { name: 'bill_apollo_tyres_inv_4412.pdf', size: 142850, date: '2026-09-15', vendor: 'Apollo_Tyres', inv: 'INV-4412', cat: 'Purchase Invoices' as AuditDocCategory, conf: 92 },
    { name: 'sales_register_summary_q2.xlsx', size: 94100, date: '2026-09-30', vendor: 'Sales_Dept', inv: 'SAL-Q2', cat: 'Sales Invoices' as AuditDocCategory, conf: 94 },
    { name: 'petty_cash_vouchers_sept_batch.pdf', size: 68100, date: '2026-09-25', vendor: 'PettyCash', inv: 'VCH-092', cat: 'Expense Vouchers' as AuditDocCategory, conf: 90 },
    { name: 'fixed_asset_register_far_additions.xlsx', size: 54100, date: '2026-09-30', vendor: 'Internal_Audit', inv: 'FAR-2026', cat: 'Supporting Schedules' as AuditDocCategory, conf: 92 },
    // Intentional duplicate file by size
    { name: 'hdfc_bank_statement_q2_fy26_copy.xlsx', size: 58210, date: '2026-09-30', vendor: 'HDFC_Bank', inv: 'STMT-SEP26', cat: 'Bank Statements' as AuditDocCategory, conf: 98 },
  ];

  return samples.map((s, idx) => {
    const ext = s.name.split('.').pop() || 'pdf';
    const blob = new Blob([`Authentic financial document audit stream for ${s.name}`], { type: 'application/octet-stream' });
    const file = new File([blob], s.name, { type: 'application/octet-stream' });
    const catCode = getCategoryCode(s.cat);

    return {
      id: `sample-${idx}-${Date.now()}`,
      file,
      originalName: s.name,
      originalExtension: ext,
      sizeBytes: s.size,
      detectedDate: s.date,
      detectedVendor: s.vendor,
      detectedInvoiceNo: s.inv,
      detectedCategory: s.cat,
      categoryCode: catCode,
      evidenceIndex: `WP-${catCode}-${String(idx + 1).padStart(3, '0')}`,
      auditArea: classifyAuditDocument(s.name).auditArea,
      reviewerNotes: '',
      newName: '',
      status: 'Verified & Indexed',
      confidenceScore: s.conf,
    };
  });
}
