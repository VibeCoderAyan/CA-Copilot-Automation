import * as XLSX from 'xlsx';
import { parseAmount, parseDate } from './dataSanitizer';
import { createStyledWorksheet } from './excelStyler';

export type SamplingMethod =
  | 'simple_random'
  | 'systematic'
  | 'stratified'
  | 'high_value_targeted';

export interface SamplingConfig {
  method: SamplingMethod;
  sampleSize: number;
  randomSeed: number; // for reproducibility
  monetaryThreshold: number; // for targeted high value
  systematicInterval?: number;
  strataBoundaries?: number[]; // e.g. [50000, 200000]
  strataSampleSizes?: number[]; // e.g. [3, 4, 5]
}

export interface PopulationRecord {
  id: string;
  sourceRowIndex: number;
  transactionId: string;
  date: string;
  partyName: string;
  amount: number;
  accountHead: string;
  narration: string;
  raw: Record<string, any>;
}

export interface SampledRecord extends PopulationRecord {
  sampleCategory: 'Random Sample' | 'High-Value Targeted' | 'Stratum 1' | 'Stratum 2' | 'Stratum 3';
  samplingRationale: string;
  auditTestingResult?: 'Satisfactory' | 'Exception Found' | 'Pending Review';
  varianceAmount?: number;
  supportingVerified?: boolean;
  auditorNotes?: string;
}

export interface PopulationScrutiny {
  totalCount: number;
  totalValue: number;
  averageValue: number;
  maxValue: number;
  minValue: number;
  missingIdCount: number;
  duplicateIdCount: number;
  invalidAmountCount: number;
  negativeAmountCount: number;
  records: PopulationRecord[];
}

export interface SamplingResult {
  clientName: string;
  auditPeriod: string;
  populationName: string;
  config: SamplingConfig;
  executionTimestamp: string;
  populationScrutiny: PopulationScrutiny;
  sampledRecords: SampledRecord[];
  totalSampledValue: number;
  coveragePercentage: number;
  isReproducible: boolean;
}

export interface SamplingColumnMapping {
  idCol: string;
  dateCol: string;
  partyCol: string;
  amountCol: string;
  headCol: string;
  narrationCol: string;
}

import { detectColumnCategory } from './columnDetector';

export function detectSamplingColumns(headers: string[]): SamplingColumnMapping {
  const m: SamplingColumnMapping = {
    idCol: '',
    dateCol: '',
    partyCol: '',
    amountCol: '',
    headCol: '',
    narrationCol: '',
  };

  for (const h of headers) {
    const cat = detectColumnCategory(h);
    if (!m.idCol && cat === 'voucherNo') m.idCol = h;
    else if (!m.dateCol && cat === 'date') m.dateCol = h;
    else if (!m.partyCol && cat === 'partyName') m.partyCol = h;
    else if (!m.amountCol && (cat === 'totalAmount' || cat === 'debit' || cat === 'credit')) m.amountCol = h;
    else if (!m.headCol && (cat === 'head' || cat === 'accountCode')) m.headCol = h;
    else if (!m.narrationCol && cat === 'narration') m.narrationCol = h;
  }

  for (const h of headers) {
    const clean = h.toLowerCase().trim();
    if (!m.idCol && /vch.*no|voucher|doc.*no|txn.*id|trans.*no|ref/i.test(clean)) m.idCol = h;
    else if (!m.dateCol && /date|dt|trans.*date/i.test(clean)) m.dateCol = h;
    else if (!m.partyCol && /party|vendor|customer|supplier|payee/i.test(clean)) m.partyCol = h;
    else if (!m.amountCol && /amount|amt|debit|value|total/i.test(clean)) m.amountCol = h;
    else if (!m.headCol && /account|head|ledger|category|particulars/i.test(clean)) m.headCol = h;
    else if (!m.narrationCol && /narration|remark|memo|description/i.test(clean)) m.narrationCol = h;
  }

  return m;
}

/**
 * Seeded pseudo-random number generator (Mulberry32) for reproducible audit sampling
 */
function createSeededRandom(seed: number) {
  return function () {
    let t = (seed += 0x6d2b79f5);
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

/**
 * Scrutinizes and parses the raw transaction population
 */
export function scrutinizePopulation(
  rawRows: Record<string, any>[],
  mapping: SamplingColumnMapping
): PopulationScrutiny {
  const records: PopulationRecord[] = [];
  const idTracker = new Map<string, number>();

  let missingIdCount = 0;
  let duplicateIdCount = 0;
  let invalidAmountCount = 0;
  let negativeAmountCount = 0;
  let totalValue = 0;
  let maxValue = 0;
  let minValue = Infinity;

  rawRows.forEach((r, idx) => {
    const rawId = String(r[mapping.idCol] || '').trim();
    const date = parseDate(r[mapping.dateCol], 'DD-MM-YYYY');
    const partyName = String(r[mapping.partyCol] || '').trim();
    const amount = parseAmount(r[mapping.amountCol]);
    const accountHead = String(r[mapping.headCol] || '').trim();
    const narration = String(r[mapping.narrationCol] || '').trim();

    if (!rawId) missingIdCount++;
    else {
      if (idTracker.has(rawId.toLowerCase())) {
        duplicateIdCount++;
      } else {
        idTracker.set(rawId.toLowerCase(), idx + 2);
      }
    }

    if (amount === 0 && r[mapping.amountCol] !== 0 && r[mapping.amountCol] !== '0') {
      invalidAmountCount++;
    }
    if (amount < 0) negativeAmountCount++;

    totalValue += amount;
    if (amount > maxValue) maxValue = amount;
    if (amount < minValue) minValue = amount;

    records.push({
      id: `pop-${idx + 1}`,
      sourceRowIndex: idx + 2,
      transactionId: rawId || `ROW-${idx + 2}`,
      date,
      partyName,
      amount,
      accountHead,
      narration,
      raw: r,
    });
  });

  const totalCount = records.length;
  const averageValue = totalCount > 0 ? totalValue / totalCount : 0;
  if (minValue === Infinity) minValue = 0;

  return {
    totalCount,
    totalValue,
    averageValue,
    maxValue,
    minValue,
    missingIdCount,
    duplicateIdCount,
    invalidAmountCount,
    negativeAmountCount,
    records,
  };
}

/**
 * Executes statistical or judgmental audit sampling on population
 */
export function executeAuditSampling(
  scrutiny: PopulationScrutiny,
  config: SamplingConfig,
  clientName = 'Corporate Client',
  auditPeriod = 'FY 2026-27',
  populationName = 'General Ledger Expenses & Payments'
): SamplingResult {
  const rng = createSeededRandom(config.randomSeed || 12345);
  const pool = [...scrutiny.records];
  const sampledRecords: SampledRecord[] = [];

  if (config.method === 'high_value_targeted') {
    // 100% Testing of items exceeding monetary threshold
    pool.forEach((r) => {
      if (r.amount >= config.monetaryThreshold) {
        sampledRecords.push({
          ...r,
          sampleCategory: 'High-Value Targeted',
          samplingRationale: `100% targeted review: Amount ₹${r.amount.toLocaleString('en-IN')} exceeds threshold of ₹${config.monetaryThreshold.toLocaleString('en-IN')}`,
          supportingVerified: false,
        });
      }
    });
  } else if (config.method === 'simple_random') {
    // Simple Random Sampling without replacement
    const targetSize = Math.min(config.sampleSize, pool.length);
    const available = [...pool];
    for (let i = 0; i < targetSize; i++) {
      const randIdx = Math.floor(rng() * available.length);
      const chosen = available.splice(randIdx, 1)[0];
      sampledRecords.push({
        ...chosen,
        sampleCategory: 'Random Sample',
        samplingRationale: `Statistical Random Selection (Seed ${config.randomSeed}, Rank ${i + 1} of ${targetSize})`,
        supportingVerified: false,
      });
    }
  } else if (config.method === 'systematic') {
    // Systematic Sampling with random start
    const N = pool.length;
    const n = Math.min(config.sampleSize, N);
    const k = Math.max(1, Math.floor(N / n));
    const startIdx = Math.floor(rng() * k);

    for (let i = startIdx; i < N && sampledRecords.length < n; i += k) {
      sampledRecords.push({
        ...pool[i],
        sampleCategory: 'Random Sample',
        samplingRationale: `Systematic selection every ${k}th item starting at offset ${startIdx + 1}`,
        supportingVerified: false,
      });
    }
  } else if (config.method === 'stratified') {
    // Stratified Sampling by amount
    const b1 = config.strataBoundaries?.[0] || 50000;
    const b2 = config.strataBoundaries?.[1] || 200000;
    const s1Size = config.strataSampleSizes?.[0] || 2;
    const s2Size = config.strataSampleSizes?.[1] || 3;
    const s3Size = config.strataSampleSizes?.[2] || 4;

    const stratum1 = pool.filter((r) => r.amount < b1);
    const stratum2 = pool.filter((r) => r.amount >= b1 && r.amount < b2);
    const stratum3 = pool.filter((r) => r.amount >= b2);

    const pickStratum = (arr: PopulationRecord[], count: number, label: any) => {
      const avail = [...arr];
      const target = Math.min(count, avail.length);
      for (let i = 0; i < target; i++) {
        const randIdx = Math.floor(rng() * avail.length);
        const chosen = avail.splice(randIdx, 1)[0];
        sampledRecords.push({
          ...chosen,
          sampleCategory: label,
          samplingRationale: `Stratified sample from ${label} group`,
          supportingVerified: false,
        });
      }
    };

    pickStratum(stratum1, s1Size, 'Stratum 1');
    pickStratum(stratum2, s2Size, 'Stratum 2');
    pickStratum(stratum3, s3Size, 'Stratum 3');
  }

  const totalSampledVal = sampledRecords.reduce((sum, r) => sum + r.amount, 0);
  const coveragePercent = scrutiny.totalValue > 0 ? (totalSampledVal / scrutiny.totalValue) * 100 : 0;

  return {
    clientName,
    auditPeriod,
    populationName,
    config,
    executionTimestamp: new Date().toISOString(),
    populationScrutiny: scrutiny,
    sampledRecords,
    totalSampledValue: totalSampledVal,
    coveragePercentage: coveragePercent,
    isReproducible: config.method === 'simple_random' || config.method === 'systematic',
  };
}

/**
 * Exports complete Excel sampling register workbook with blank audit testing columns
 */
export function exportSamplingWorkbook(result: SamplingResult): XLSX.WorkBook {
  const wb = XLSX.utils.book_new();

  // Sheet 1: Methodology & Parameters
  const summaryRows = [
    { Parameter: 'Client Entity Name', Value: result.clientName },
    { Parameter: 'Audit Financial Period', Value: result.auditPeriod },
    { Parameter: 'Population Scope', Value: result.populationName },
    { Parameter: 'Sampling Standard Applied', Value: 'ICAI SA 530 / ISA 530 (Audit Sampling)' },
    { Parameter: 'Selection Methodology', Value: result.config.method.replace(/_/g, ' ').toUpperCase() },
    { Parameter: 'Random Seed (Reproducibility)', Value: result.config.randomSeed },
    { Parameter: 'Monetary Threshold (₹)', Value: result.config.monetaryThreshold },
    { Parameter: 'Total Population Transactions (N)', Value: result.populationScrutiny.totalCount },
    { Parameter: 'Total Population Monetary Value (₹)', Value: result.populationScrutiny.totalValue },
    { Parameter: 'Selected Sample Size (n)', Value: result.sampledRecords.length },
    { Parameter: 'Total Sampled Value (₹)', Value: result.totalSampledValue },
    { Parameter: 'Audit Coverage by Monetary Value', Value: `${result.coveragePercentage.toFixed(1)}%` },
    { Parameter: 'Execution Timestamp', Value: result.executionTimestamp },
    { Parameter: 'Mandatory Auditor Disclaimer', Value: 'Sample testing does not guarantee detection of all errors or fraud in the population.' },
  ];
  const wsSummary = createStyledWorksheet(summaryRows, {
    currencyColumns: ['Value'],
    columnWidths: { Parameter: 38, Value: 34 },
  });
  XLSX.utils.book_append_sheet(wb, wsSummary, 'Sampling_Methodology');

  // Sheet 2: Selected Audit Sample Register
  const registerRows = result.sampledRecords.map((s, idx) => ({
    'Sample #': idx + 1,
    'Source Row': s.sourceRowIndex,
    'Voucher / Ref #': s.transactionId,
    'Date': s.date,
    'Party / Payee': s.partyName,
    'Account Head': s.accountHead,
    'Amount (₹)': s.amount,
    'Selection Classification': s.sampleCategory,
    'Sampling Rationale': s.samplingRationale,
    // Blank audit testing columns for the user to complete
    'Supporting Bill Attached? (Y/N)': '',
    'Vouching Test Result (Pass/Fail)': '',
    'Variance / Disallowance (₹)': '',
    'Auditor Initials & Remarks': '',
  }));

  const wsRegister = createStyledWorksheet(registerRows, {
    currencyColumns: ['Amount (₹)', 'Variance / Disallowance (₹)'],
    columnWidths: {
      'Sample #': 8,
      'Source Row': 10,
      'Voucher / Ref #': 18,
      'Date': 14,
      'Party / Payee': 26,
      'Account Head': 24,
      'Amount (₹)': 18,
      'Selection Classification': 22,
      'Sampling Rationale': 40,
      'Supporting Bill Attached? (Y/N)': 28,
      'Vouching Test Result (Pass/Fail)': 28,
      'Variance / Disallowance (₹)': 26,
      'Auditor Initials & Remarks': 34,
    },
  });
  XLSX.utils.book_append_sheet(wb, wsRegister, 'Audit_Sample_Register');

  return wb;
}

/**
 * Realistic accounting transaction population sample
 */
export function getSampleTransactionPopulation(): Record<string, any>[] {
  return [
    { 'Voucher No': 'VR-2026-101', 'Date': '02-08-2026', 'Party Name': 'Godrej Consumer Products', 'Account Head': 'Factory Raw Materials', 'Amount': 485000, 'Narration': 'Soap noodles procurement' },
    { 'Voucher No': 'VR-2026-102', 'Date': '04-08-2026', 'Party Name': 'Maharashtra State Electricity', 'Account Head': 'Power & Fuel', 'Amount': 95000, 'Narration': 'Factory HT power bill' },
    { 'Voucher No': 'VR-2026-103', 'Date': '06-08-2026', 'Party Name': 'Blue Dart Express Ltd', 'Account Head': 'Courier & Freight', 'Amount': 18200, 'Narration': 'Sample dispatches to clients' },
    { 'Voucher No': 'VR-2026-104', 'Date': '08-08-2026', 'Party Name': 'Kirloskar Brothers Ltd', 'Account Head': 'Plant Machinery Repairs', 'Amount': 620000, 'Narration': 'Annual turbine overhaul (HIGH VALUE)' },
    { 'Voucher No': 'VR-2026-105', 'Date': '11-08-2026', 'Party Name': 'HDFC Standard Life Insurance', 'Account Head': 'Staff Welfare Insurance', 'Amount': 125000, 'Narration': 'Group mediclaim policy' },
    { 'Voucher No': 'VR-2026-106', 'Date': '14-08-2026', 'Party Name': 'Apex Industrial Packaging', 'Account Head': 'Packing Material', 'Amount': 74000, 'Narration': 'Corrugated carton boxes' },
    { 'Voucher No': 'VR-2026-107', 'Date': '16-08-2026', 'Party Name': 'Vodafone Idea Limited', 'Account Head': 'Telephone & Internet', 'Amount': 12400, 'Narration': 'Leased line charges' },
    { 'Voucher No': 'VR-2026-108', 'Date': '19-08-2026', 'Party Name': 'L&T Heavy Engineering', 'Account Head': 'Boiler Spare Parts', 'Amount': 890000, 'Narration': 'Pressure vessel fabrication (HIGH VALUE)' },
    { 'Voucher No': 'VR-2026-109', 'Date': '22-08-2026', 'Party Name': 'Dr. Sharma Diagnostics', 'Account Head': 'Employee Health Checkup', 'Amount': 28000, 'Narration': 'Annual health checkup camp' },
    { 'Voucher No': 'VR-2026-110', 'Date': '24-08-2026', 'Party Name': 'Tata Consultancy Services', 'Account Head': 'ERP AMC & Support', 'Amount': 340000, 'Narration': 'SAP ERP quarterly maintenance' },
    { 'Voucher No': 'VR-2026-111', 'Date': '26-08-2026', 'Party Name': 'Office Stationery Mart', 'Account Head': 'Printing & Stationery', 'Amount': 8900, 'Narration': 'Printer cartridges and registers' },
    { 'Voucher No': 'VR-2026-112', 'Date': '28-08-2026', 'Party Name': 'Siemens Financial Services', 'Account Head': 'Machinery Lease Rental', 'Amount': 540000, 'Narration': 'CNC machine monthly lease (HIGH VALUE)' },
    { 'Voucher No': 'VR-2026-113', 'Date': '29-08-2026', 'Party Name': 'Shree Security Guards', 'Account Head': 'Security Agency Charges', 'Amount': 64000, 'Narration': 'Factory security deployment' },
    { 'Voucher No': 'VR-2026-114', 'Date': '30-08-2026', 'Party Name': 'Advocate R. K. Singhal', 'Account Head': 'Legal Retainer Fees', 'Amount': 50000, 'Narration': 'High court arbitration retainer' },
    { 'Voucher No': 'VR-2026-115', 'Date': '31-08-2026', 'Party Name': 'Indian Oil Corporation', 'Account Head': 'Generator Diesel', 'Amount': 145000, 'Narration': 'DG set fuel procurement' },
  ];
}
