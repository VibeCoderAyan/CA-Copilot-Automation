// Intelligent column detection and fuzzy alias mapping for CA automation spreadsheets

export type ColumnCategory =
  | 'date'
  | 'dueDate'
  | 'voucherNo'
  | 'partyName'
  | 'gstin'
  | 'taxableAmount'
  | 'cgst'
  | 'sgst'
  | 'igst'
  | 'totalAmount'
  | 'debit'
  | 'credit'
  | 'balance'
  | 'narration'
  | 'refNo'
  | 'accountCode'
  | 'head';

export interface ColumnMappingResult {
  detected: Record<string, ColumnCategory>;
  unmapped: string[];
  missingMandatory: ColumnCategory[];
  allFound: boolean;
}

const ALIASES: Record<ColumnCategory, string[]> = {
  dueDate: [
    'due_date',
    'duedate',
    'payment_due_date',
    'paymentduedate',
    'pay_due_date',
    'payduedate',
    'pay_due',
    'paydue',
    'bill_due_date',
    'billduedate',
    'invoice_due_date',
    'invoiceduedate',
    'inv_due_date',
    'invduedate',
    'due_on',
    'dueon',
    'maturity_date',
    'maturitydate',
  ],
  date: [
    'date',
    'inv_date',
    'invoice_date',
    'invoicedate',
    'bill_date',
    'billdate',
    'voucher_date',
    'voucherdate',
    'vch_date',
    'vchdate',
    'txn_date',
    'txndate',
    'transaction_date',
    'transactiondate',
    'posting_date',
    'postingdate',
    'value_date',
    'valuedate',
    'entry_date',
    'doc_date',
    'docdate',
    'bill_dt',
    'inv_dt',
    'vch_dt',
    'txn_dt',
  ],
  voucherNo: [
    'voucher_no',
    'voucherno',
    'vch_no',
    'vchno',
    'vouchernum',
    'voucher_num',
    'voucher_number',
    'vouchernumber',
    'invoice_no',
    'invoiceno',
    'invoice_number',
    'invoicenumber',
    'inv_no',
    'invno',
    'inv_num',
    'bill_no',
    'billno',
    'bill_num',
    'bill_number',
    'billnumber',
    'document_no',
    'doc_no',
    'docno',
    'doc_num',
    'document_number',
    'reference_no',
    'ref_no',
    'refno',
    'chq_no',
    'cheque_no',
    'chequeno',
    'inum',
    'serial_no',
    'entry_no',
    'entryno',
  ],
  partyName: [
    'party_name',
    'partyname',
    'party',
    'customer_name',
    'customername',
    'customer',
    'vendor_name',
    'vendorname',
    'vendor',
    'supplier_name',
    'suppliername',
    'supplier',
    'trade_name',
    'tradename',
    'legal_name',
    'legalname',
    'account_name',
    'accountname',
    'account',
    'ledger_name',
    'cname',
    'debtor',
    'debtors',
    'debtor_name',
    'debtorname',
    'debtors_name',
    'debtorsname',
    'sundry_debtors',
    'sundry_debtor',
    'sundrydebtors',
    'sundrydebtor',
    'creditor',
    'creditors',
    'creditor_name',
    'creditorname',
    'creditors_name',
    'creditorsname',
    'sundry_creditors',
    'sundry_creditor',
    'sundrycreditors',
    'sundrycreditor',
    'client',
    'client_name',
    'clientname',
    'buyer',
    'buyer_name',
    'buyername',
    'consignee',
    'consignee_name',
    'billed_to',
    'billedto',
    'entity_name',
    'entityname',
  ],
  gstin: [
    'gstin',
    'gst_no',
    'gstno',
    'gstin_uin',
    'gstinuin',
    'gstin_of_supplier',
    'supplier_gstin',
    'party_gstin',
    'customer_gstin',
    'buyer_gstin',
    'recipient_gstin',
    'gstin_of_customer',
    'gstin_of_recipient',
    'gst_number',
    'gstnumber',
    'pan_gstin',
    'tax_id',
    'taxid',
    'ctin',
  ],
  taxableAmount: [
    'taxable_amount',
    'taxableamount',
    'taxable_amt',
    'taxableamt',
    'taxable_value',
    'taxablevalue',
    'taxable',
    'assessable_amount',
    'assessableamount',
    'assessable_value',
    'assessablevalue',
    'assessable_amt',
    'assessableamt',
    'assessable',
    'basic_amount',
    'basic_amt',
    'basicamount',
    'txval',
  ],
  cgst: [
    'cgst',
    'central_tax',
    'centraltax',
    'cgst_amount',
    'cgst_amt',
    'cgst_rate',
    'camt',
  ],
  sgst: [
    'sgst',
    'state_tax',
    'statetax',
    'utgst',
    'sgst_amount',
    'sgst_amt',
    'sgst_rate',
    'samt',
  ],
  igst: [
    'igst',
    'integrated_tax',
    'integratedtax',
    'igst_amount',
    'igst_amt',
    'igst_rate',
    'iamt',
  ],
  totalAmount: [
    'total_amount',
    'totalamount',
    'total_amt',
    'totalamt',
    'gross_total',
    'grosstotal',
    'total_invoice',
    'totalinvoice',
    'invoice_value',
    'invoicevalue',
    'invoice_amount',
    'invoiceamount',
    'invoice_total',
    'invoicetotal',
    'gross_amount',
    'grossamount',
    'net_amount',
    'netamount',
    'total_bill',
    'totalbill',
    'bill_amount',
    'billamount',
    'grand_total',
    'grandtotal',
    'voucher_amount',
    'voucheramount',
    'total',
    'amount',
    'amt',
    'val',
  ],
  debit: [
    'debit',
    'dr',
    'debit_amount',
    'debitamount',
    'debit_amt',
    'debitamt',
    'dr_amount',
    'dramount',
    'dr_amt',
    'dramt',
    'withdrawal',
    'withdrawals',
    'payment',
    'payments',
    'paid_out',
    'paidout',
  ],
  credit: [
    'credit',
    'cr',
    'credit_amount',
    'creditamount',
    'credit_amt',
    'creditamt',
    'cr_amount',
    'cramount',
    'cr_amt',
    'cramt',
    'deposit',
    'deposits',
    'receipt',
    'receipts',
    'received',
  ],
  balance: [
    'balance',
    'closing_balance',
    'closingbalance',
    'closing_bal',
    'closingbal',
    'available_balance',
    'availablebalance',
    'outstanding_balance',
    'outstandingbalance',
    'outstanding_amount',
    'outstandingamount',
    'outstanding',
    'net_balance',
    'netbalance',
    'due_amount',
    'dueamount',
    'bal',
    'book_balance',
  ],
  narration: [
    'narration',
    'description',
    'item_description',
    'itemdescription',
    'item_desc',
    'product_description',
    'goods_description',
    'service_description',
    'remarks',
    'particulars',
    'details',
    'transaction_details',
    'memo',
    'notes',
    'comments',
  ],
  refNo: [
    'ref_no',
    'refno',
    'reference',
    'reference_no',
    'chq_no',
    'chqno',
    'cheque_no',
    'chequeno',
    'utr',
    'utr_no',
    'txn_id',
    'transaction_id',
  ],
  accountCode: [
    'account_code',
    'accountcode',
    'ac_code',
    'accode',
    'ledger_code',
    'code',
    'gl_code',
  ],
  head: [
    'head',
    'account_head',
    'group',
    'account_group',
    'category',
    'classification',
    'type',
  ],
};

/**
 * Normalizes a column name by stripping non-alphanumeric chars and lowercasing
 */
export function normalizeHeader(header: string): string {
  if (!header) return '';
  return header
    .toString()
    .trim()
    .toLowerCase()
    .replace(/[₹$€,"'()[\]{}]/g, '') // remove currency and brackets/quotes
    .replace(/[\s\-_./\\]+/g, '_')    // normalize spaces, dashes, dots, slashes, underscores to single underscore
    .replace(/^_+|_+$/g, '');        // trim leading/trailing underscores
}

// Priority order for semantic fallback matching:
// Technical specific fields (GSTIN, Tax, DueDate, Date, Balance, Debit, Credit) are checked
// before generic fields (PartyName, Narration).
const CATEGORY_PRIORITY: ColumnCategory[] = [
  'gstin',
  'cgst',
  'sgst',
  'igst',
  'taxableAmount',
  'totalAmount',
  'balance',
  'dueDate',
  'date',
  'voucherNo',
  'partyName',
  'debit',
  'credit',
  'refNo',
  'accountCode',
  'head',
  'narration',
];

/**
 * Identifies the best-matching category for a given column header
 */
export function detectColumnCategory(rawHeader: string): ColumnCategory | null {
  const norm = normalizeHeader(rawHeader);
  if (!norm) return null;
  const strippedNorm = norm.replace(/_/g, '');

  // Phase 1: Exact matches across ALL categories
  for (const cat of CATEGORY_PRIORITY) {
    const aliases = ALIASES[cat] || [];
    if (aliases.includes(norm)) {
      return cat;
    }
  }

  // Phase 2: Stripped underscore exact matches across all categories
  for (const cat of CATEGORY_PRIORITY) {
    const aliases = ALIASES[cat] || [];
    if (aliases.some((alias) => alias.replace(/_/g, '') === strippedNorm)) {
      return cat;
    }
  }

  // Phase 3: Tokenized & compound keyword matching using strict category priority
  for (const cat of CATEGORY_PRIORITY) {
    const aliases = ALIASES[cat] || [];
    for (const alias of aliases) {
      const strippedAlias = alias.replace(/_/g, '');

      // Strict boundary check: full token boundary separated by underscores
      const isTokenBoundary =
        norm === alias ||
        norm.startsWith(`${alias}_`) ||
        norm.endsWith(`_${alias}`) ||
        norm.includes(`_${alias}_`);

      // Substring check: only allowed for longer identifiers (>= 4 characters)
      // Prevents 2-3 char abbreviations ('dr', 'cr', 'val', 'amt', 'bal') from matching inside unrelated words ('address', 'description', 'sundry')
      const isSafeSubstring = strippedAlias.length >= 4 && strippedNorm.includes(strippedAlias);

      if (isTokenBoundary || isSafeSubstring) {
        return cat;
      }
    }
  }

  return null;
}

/**
 * Maps all columns of an uploaded spreadsheet and identifies missing mandatory fields
 */
export function mapSpreadsheetColumns(
  headers: string[],
  mandatoryCategories: ColumnCategory[] = []
): ColumnMappingResult {
  const detected: Record<string, ColumnCategory> = {};
  const unmapped: string[] = [];
  const assignedCategories = new Set<ColumnCategory>();

  for (const header of headers) {
    const category = detectColumnCategory(header);
    if (category && !assignedCategories.has(category)) {
      detected[header] = category;
      assignedCategories.add(category);
    } else {
      unmapped.push(header);
    }
  }

  const missingMandatory = mandatoryCategories.filter((cat) => !assignedCategories.has(cat));

  return {
    detected,
    unmapped,
    missingMandatory,
    allFound: missingMandatory.length === 0,
  };
}
