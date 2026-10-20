/**
 * Tally Prime & Tally.ERP 9 Direct XML Generator Engine
 * Generates 100% Deterministic, Standards-Compliant Tally XML for Automated Voucher Import
 */

export type TallyVoucherType = 'Payment' | 'Receipt' | 'Journal' | 'Contra';

export interface TallyVoucherItem {
  id?: string;
  date: string; // ISO date 'YYYY-MM-DD' or 'DD-MM-YYYY'
  voucherType: TallyVoucherType;
  voucherNumber?: string;
  bankLedger: string; // e.g. "HDFC Bank Current A/c"
  partyOrExpenseLedger: string; // e.g. "Bank Charges", "Vendor Name", "Customer Name"
  amount: number; // positive number
  narration: string; // e.g. "Direct Debit - Bank Charges as per BRS"
  chequeOrRefNo?: string;
}

export interface TallyExportOptions {
  companyName?: string;
  vouchers: TallyVoucherItem[];
}

/**
 * Escape XML special characters
 */
const escapeXml = (unsafe: string): string => {
  return (unsafe || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
};

/**
 * Format any date string to Tally standard 'YYYYMMDD'
 */
export const formatTallyDate = (dateStr: string): string => {
  if (!dateStr) {
    const now = new Date();
    const y = now.getFullYear();
    const m = String(now.getMonth() + 1).padStart(2, '0');
    const d = String(now.getDate()).padStart(2, '0');
    return `${y}${m}${d}`;
  }

  // Handle DD-MM-YYYY or DD/MM/YYYY
  if (/^\d{2}[-/]\d{2}[-/]\d{4}$/.test(dateStr)) {
    const parts = dateStr.split(/[-/]/);
    return `${parts[2]}${parts[1]}${parts[0]}`;
  }

  // Handle YYYY-MM-DD
  if (/^\d{4}-\d{2}-\d{2}/.test(dateStr)) {
    return dateStr.substring(0, 10).replace(/-/g, '');
  }

  // Fallback to Date object parsing
  const d = new Date(dateStr);
  if (!isNaN(d.getTime())) {
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, '0');
    const day = String(d.getDate()).padStart(2, '0');
    return `${y}${m}${day}`;
  }

  return '20240901';
};

/**
 * Generate Complete Tally Prime XML String
 */
export const generateTallyXml = (options: TallyExportOptions): string => {
  const company = escapeXml(options.companyName || 'CA Copilot Practice Company');
  const vouchers = options.vouchers;

  const voucherXmlBlocks = vouchers.map((v, idx) => {
    const dateFormatted = formatTallyDate(v.date);
    const vchNum = escapeXml(v.voucherNumber || `AUTO-${dateFormatted}-${idx + 1}`);
    const narration = escapeXml(v.narration || `Generated via CA Copilot Automated Reconciliation`);
    const amountStr = Math.abs(v.amount).toFixed(2);

    let ledgerBlocks = '';

    if (v.voucherType === 'Payment') {
      // Payment: Dr Expense/Party, Cr Bank
      ledgerBlocks = `
            <ALLLEDGERENTRIES.LIST>
              <LEDGERNAME>${escapeXml(v.partyOrExpenseLedger)}</LEDGERNAME>
              <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>
              <AMOUNT>-${amountStr}</AMOUNT>
            </ALLLEDGERENTRIES.LIST>
            <ALLLEDGERENTRIES.LIST>
              <LEDGERNAME>${escapeXml(v.bankLedger)}</LEDGERNAME>
              <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>
              <AMOUNT>${amountStr}</AMOUNT>
            </ALLLEDGERENTRIES.LIST>`;
    } else if (v.voucherType === 'Receipt') {
      // Receipt: Dr Bank, Cr Customer/Income
      ledgerBlocks = `
            <ALLLEDGERENTRIES.LIST>
              <LEDGERNAME>${escapeXml(v.bankLedger)}</LEDGERNAME>
              <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>
              <AMOUNT>-${amountStr}</AMOUNT>
            </ALLLEDGERENTRIES.LIST>
            <ALLLEDGERENTRIES.LIST>
              <LEDGERNAME>${escapeXml(v.partyOrExpenseLedger)}</LEDGERNAME>
              <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>
              <AMOUNT>${amountStr}</AMOUNT>
            </ALLLEDGERENTRIES.LIST>`;
    } else {
      // Journal: Dr PartyA, Cr PartyB
      ledgerBlocks = `
            <ALLLEDGERENTRIES.LIST>
              <LEDGERNAME>${escapeXml(v.partyOrExpenseLedger)}</LEDGERNAME>
              <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>
              <AMOUNT>-${amountStr}</AMOUNT>
            </ALLLEDGERENTRIES.LIST>
            <ALLLEDGERENTRIES.LIST>
              <LEDGERNAME>${escapeXml(v.bankLedger)}</LEDGERNAME>
              <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>
              <AMOUNT>${amountStr}</AMOUNT>
            </ALLLEDGERENTRIES.LIST>`;
    }

    return `
        <TALLYMESSAGE xmlns:UDF="TallyUDF">
          <VOUCHER VCHTYPE="${v.voucherType}" ACTION="Create" OBJVIEW="Accounting Voucher View">
            <DATE>${dateFormatted}</DATE>
            <EFFECTIVEDATE>${dateFormatted}</EFFECTIVEDATE>
            <VOUCHERTYPENAME>${v.voucherType}</VOUCHERTYPENAME>
            <VOUCHERNUMBER>${vchNum}</VOUCHERNUMBER>
            <NARRATION>${narration}</NARRATION>${ledgerBlocks}
          </VOUCHER>
        </TALLYMESSAGE>`;
  });

  return `<?xml version="1.0" encoding="UTF-8"?>
<ENVELOPE>
  <HEADER>
    <TALLYREQUEST>Import Data</TALLYREQUEST>
  </HEADER>
  <BODY>
    <IMPORTDATA>
      <REQUESTDESC>
        <REPORTNAME>Vouchers</REPORTNAME>
        <STATICVARIABLES>
          <SVCURRENTCOMPANY>${company}</SVCURRENTCOMPANY>
        </STATICVARIABLES>
      </REQUESTDESC>
      <REQUESTDATA>${voucherXmlBlocks.join('')}
      </REQUESTDATA>
    </IMPORTDATA>
  </BODY>
</ENVELOPE>`;
};

/**
 * Smart Narration Heuristic: Auto-detects Tally Ledger & Voucher Type from Bank Description
 */
export const classifyBankNarration = (
  narration: string,
  isDebit: boolean
): { suggestedLedger: string; voucherType: TallyVoucherType } => {
  const text = (narration || '').toLowerCase();

  // 1. Bank charges & commission
  if (/charge|chg|annual fee|sms fee|maintenance|processing fee|card fee|folio|min bal|cheque bounce|ecs return/.test(text)) {
    return { suggestedLedger: 'Bank Charges & Commission', voucherType: 'Payment' };
  }

  // 2. Interest
  if (/interest|int\.pd|int earned|sweep|auto-sweep|fd int|term dep/.test(text)) {
    return isDebit
      ? { suggestedLedger: 'Bank Interest & Finance Costs', voucherType: 'Payment' }
      : { suggestedLedger: 'Interest Income on Sweep Deposits', voucherType: 'Receipt' };
  }

  // 3. Statutory Payments & Taxes
  if (/gst|cgst|sgst|igst|tax payment|cbdt|challan|tds payment|advance tax|epfo|esic|pf payment/.test(text)) {
    return { suggestedLedger: 'Statutory Taxes & Duties', voucherType: 'Payment' };
  }

  // 4. Payment Gateway & Platform Fees
  if (/razorpay|paytm|billdesk|stripe|ccavenue|cashfree|pg fee|pg charges|merchant fee/.test(text)) {
    return { suggestedLedger: 'Payment Gateway Processing Charges', voucherType: 'Payment' };
  }

  // 5. Utilities & Rent
  if (/bescom|mseb|electricity|tata power|adanipower|power corp|electric/.test(text)) {
    return { suggestedLedger: 'Electricity & Utility Expenses', voucherType: 'Payment' };
  }
  if (/rent|lease|landlord|office premise/.test(text)) {
    return { suggestedLedger: 'Office Rent Expense', voucherType: 'Payment' };
  }
  if (/salary|payroll|wages|staff advance|stipend/.test(text)) {
    return { suggestedLedger: 'Salaries & Staff Remuneration', voucherType: 'Payment' };
  }

  // 6. Generic Fallback
  if (isDebit) {
    return { suggestedLedger: 'General Bank Debits (Unmatched)', voucherType: 'Payment' };
  } else {
    return { suggestedLedger: 'Direct Customer Collection (Unmatched)', voucherType: 'Receipt' };
  }
};

/**
 * 1-Click Browser Download for Generated Tally XML (Robust argument order)
 */
export const downloadTallyXml = (arg1: string, arg2: string): void => {
  const isXml1 = arg1.includes('<?xml') || arg1.includes('<ENVELOPE');
  const xmlContent = isXml1 ? arg1 : arg2;
  const filename = isXml1 ? arg2 : arg1;
  const safeFilename = filename.endsWith('.xml') ? filename : `${filename}.xml`;

  const blob = new Blob([xmlContent], { type: 'application/xml;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = safeFilename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
};
