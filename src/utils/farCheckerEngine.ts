import * as XLSX from 'xlsx';
import { parseAmount, parseDate } from './dataSanitizer';
import { createStyledWorksheet } from './excelStyler';

export type DeprMethod = 'companies_act_slm' | 'companies_act_wdv' | 'it_act_block';

export interface ScheduleIILifeGuide {
  category: string;
  defaultUsefulLifeYears: number;
  itBlockRatePercent: number;
}

export const SCHEDULE_II_GUIDE: Record<string, ScheduleIILifeGuide> = {
  plant_machinery: { category: 'Plant & Machinery', defaultUsefulLifeYears: 15, itBlockRatePercent: 15 },
  computers_laptops: { category: 'Computers & Laptops', defaultUsefulLifeYears: 3, itBlockRatePercent: 40 },
  servers_networks: { category: 'Servers & Network Equipment', defaultUsefulLifeYears: 6, itBlockRatePercent: 40 },
  office_equipment: { category: 'Office Equipment', defaultUsefulLifeYears: 5, itBlockRatePercent: 15 },
  furniture_fixtures: { category: 'Furniture & Fixtures', defaultUsefulLifeYears: 10, itBlockRatePercent: 10 },
  motor_vehicles: { category: 'Motor Vehicles (Personal/Office)', defaultUsefulLifeYears: 8, itBlockRatePercent: 15 },
  commercial_vehicles: { category: 'Commercial Vehicles (Lorries/Taxis)', defaultUsefulLifeYears: 6, itBlockRatePercent: 30 },
  buildings_factory: { category: 'Buildings (Factory)', defaultUsefulLifeYears: 30, itBlockRatePercent: 10 },
  buildings_office: { category: 'Buildings (Other / Office)', defaultUsefulLifeYears: 60, itBlockRatePercent: 10 },
};

export interface FarAssetRecord {
  id: string;
  rowIndex: number;
  assetId: string;
  assetDescription: string;
  category: string;
  matchedCategoryKey: string;
  capitalizationDate: string;
  capitalizationDateObj: Date | null;
  grossCost: number;
  openingAccumDepr: number;
  openingWdv: number;
  bookDepreciation: number;
  closingBookWdv: number;
  usefulLifeYears: number;
  deprMethod: string;

  // Recomputed metrics
  recomputedDepreciation: number;
  depreciationVariance: number;
  recomputedClosingWdv: number;
  isAdditionDuringYear: boolean;
  daysUsedInYear: number;
  isPutToUseLessThan180Days: boolean;

  issues: FarIssue[];
  raw: Record<string, any>;
}

export interface FarIssue {
  id: string;
  assetId: string;
  rowIndex: number;
  category:
    | 'duplicate_asset_id'
    | 'missing_date'
    | 'missing_cost'
    | 'inconsistent_useful_life'
    | 'depreciation_mismatch'
    | 'exceeded_95_percent_limit'
    | 'full_year_on_new_addition'
    | 'arithmetic_inconsistency';
  severity: 'critical' | 'warning' | 'info';
  title: string;
  description: string;
  suggestedAction: string;
  bookValue: number;
  recomputedValue: number;
  variance: number;
}

export interface FarAuditSummary {
  totalAssets: number;
  totalAssetsCount?: number;
  totalGrossCost: number;
  totalGrossBlock?: number;
  totalOpeningAccumDepr: number;
  totalBookDepreciation: number;
  totalDepreciation?: number;
  totalRecomputedDepreciation: number;
  netDepreciationVariance: number;
  totalClosingBookWdv: number;
  criticalIssuesCount: number;
  warningIssuesCount: number;
  duplicateIdsCount: number;
  depreciationDiscrepanciesCount: number;
  issues: FarIssue[];
}

export interface FarColumnMapping {
  assetIdCol: string;
  descCol: string;
  categoryCol: string;
  dateCol: string;
  costCol: string;
  accumDeprCol: string;
  openingWdvCol: string;
  bookDeprCol: string;
  closingWdvCol: string;
  usefulLifeCol: string;
  methodCol: string;
}

export interface FarCheckerConfig {
  deprMethod: DeprMethod;
  residualValuePercent: number; // default 5.0
  fyStartYear: number; // e.g. 2026 for FY 2026-27
  enforce180DayRule: boolean;
}

export const DEFAULT_FAR_CONFIG: FarCheckerConfig = {
  deprMethod: 'companies_act_slm',
  residualValuePercent: 5.0,
  fyStartYear: 2026,
  enforce180DayRule: true,
};

import { detectColumnCategory } from './columnDetector';

/**
 * Auto-detects columns for Fixed Asset Register
 */
export function detectFarColumns(headers: string[]): FarColumnMapping {
  const mapping: FarColumnMapping = {
    assetIdCol: '',
    descCol: '',
    categoryCol: '',
    dateCol: '',
    costCol: '',
    accumDeprCol: '',
    openingWdvCol: '',
    bookDeprCol: '',
    closingWdvCol: '',
    usefulLifeCol: '',
    methodCol: '',
  };

  for (const h of headers) {
    const clean = h.toLowerCase().trim();
    const cat = detectColumnCategory(h);

    if (!mapping.assetIdCol && (cat === 'voucherNo' || /asset.*id|tag.*no|asset.*code|asset.*no|item.*code|ref.*no/i.test(clean))) {
      mapping.assetIdCol = h;
    } else if (!mapping.descCol && /description|asset.*name|particulars|item.*name|details/i.test(clean)) {
      mapping.descCol = h;
    } else if (!mapping.categoryCol && /category|block|asset.*type|group|class/i.test(clean)) {
      mapping.categoryCol = h;
    } else if (!mapping.dateCol && (cat === 'date' || /date.*capital|put.*to.*use|acq.*date|date.*purchase|cap.*date|date/i.test(clean))) {
      mapping.dateCol = h;
    } else if (!mapping.costCol && (cat === 'totalAmount' || /original.*cost|gross.*block|cost|purchase.*cost|hist.*cost/i.test(clean))) {
      mapping.costCol = h;
    } else if (!mapping.accumDeprCol && /accum.*dep|depr.*opening|op.*accum/i.test(clean)) {
      mapping.accumDeprCol = h;
    } else if (!mapping.openingWdvCol && /opening.*wdv|op.*wdv|opening.*book/i.test(clean)) {
      mapping.openingWdvCol = h;
    } else if (!mapping.bookDeprCol && /depr.*year|current.*depr|depr.*claimed|depreciation|depr.*amt/i.test(clean)) {
      mapping.bookDeprCol = h;
    } else if (!mapping.closingWdvCol && (cat === 'balance' || /closing.*wdv|net.*block|closing.*book|cl.*wdv/i.test(clean))) {
      mapping.closingWdvCol = h;
    } else if (!mapping.usefulLifeCol && /useful.*life|life.*years|life|years/i.test(clean)) {
      mapping.usefulLifeCol = h;
    } else if (!mapping.methodCol && /method|depr.*type|basis/i.test(clean)) {
      mapping.methodCol = h;
    }
  }

  return mapping;
}

/**
 * Matches asset text to standard Schedule II asset category
 */
export function matchScheduleIICategory(rawCategory: string, description: string): string {
  const combined = `${rawCategory} ${description}`.toLowerCase();
  if (/laptop|computer|desktop|pc|monitor|macbook/i.test(combined)) return 'computers_laptops';
  if (/server|switch|router|rack|firewall|lan/i.test(combined)) return 'servers_networks';
  if (/car|vehicle|motor|jeep|sedan|suv|van/i.test(combined)) return 'motor_vehicles';
  if (/truck|lorry|tempo|trailer|commercial/i.test(combined)) return 'commercial_vehicles';
  if (/chair|table|desk|fixture|workstation|sofa|furniture|cupboard/i.test(combined)) return 'furniture_fixtures';
  if (/printer|scanner|projector|copier|xerox|shredder|office.*eq/i.test(combined)) return 'office_equipment';
  if (/building|factory|shed|structure/i.test(combined)) return 'buildings_factory';
  return 'plant_machinery';
}

/**
 * Main FAR Audit & Verification Engine
 */
export function auditFixedAssetRegister(
  rows: Record<string, any>[],
  mapping: FarColumnMapping,
  config?: Partial<FarCheckerConfig>
): { summary: FarAuditSummary; assets: FarAssetRecord[] } {
  const activeConfig: FarCheckerConfig = {
    ...DEFAULT_FAR_CONFIG,
    ...(config || {}),
  };
  const safeRows = Array.isArray(rows) ? rows : [];
  const safeMapping: FarColumnMapping = mapping || {
    assetIdCol: '',
    descCol: '',
    categoryCol: '',
    dateCol: '',
    costCol: '',
    accumDeprCol: '',
    openingWdvCol: '',
    bookDeprCol: '',
    closingWdvCol: '',
    usefulLifeCol: '',
    methodCol: '',
  };

  const assets: FarAssetRecord[] = [];
  const issues: FarIssue[] = [];

  const fyStart = new Date(activeConfig.fyStartYear, 3, 1); // 1st April
  const fyEnd = new Date(activeConfig.fyStartYear + 1, 2, 31); // 31st March
  const octFirst = new Date(activeConfig.fyStartYear, 9, 1); // 1st October (180 days marker)

  // Track duplicate asset IDs
  const assetIdCounts = new Map<string, number>();
  safeRows.forEach((r) => {
    const id = String(r[safeMapping.assetIdCol] || '').trim();
    if (id) {
      assetIdCounts.set(id, (assetIdCounts.get(id) || 0) + 1);
    }
  });

  let totalGrossCost = 0;
  let totalOpeningAccumDepr = 0;
  let totalBookDepreciation = 0;
  let totalRecomputedDepreciation = 0;
  let totalClosingBookWdv = 0;
  let duplicateIdsCount = 0;
  let depreciationDiscrepanciesCount = 0;

  safeRows.forEach((r, idx) => {
    const rowIndex = idx + 2;
    const assetId = String(r[safeMapping.assetIdCol] || '').trim() || `ASSET-${idx + 1}`;
    const assetDescription = String(r[safeMapping.descCol] || '').trim();
    const rawCategory = String(r[safeMapping.categoryCol] || '').trim();
    const matchedCategoryKey = matchScheduleIICategory(rawCategory, assetDescription);
    const standardGuide = SCHEDULE_II_GUIDE[matchedCategoryKey];

    const rawDate = r[safeMapping.dateCol];
    const capitalizationDate = parseDate(rawDate, 'DD-MM-YYYY');

    let capDateObj: Date | null = null;
    if (capitalizationDate) {
      const parts = capitalizationDate.split('-');
      if (parts.length === 3) {
        capDateObj = new Date(Number(parts[2]), Number(parts[1]) - 1, Number(parts[0]));
      }
    }

    const grossCost = parseAmount(r[safeMapping.costCol]);
    let openingAccumDepr = parseAmount(r[safeMapping.accumDeprCol]);
    let openingWdv = parseAmount(r[safeMapping.openingWdvCol]);
    const bookDepreciation = parseAmount(r[safeMapping.bookDeprCol]);
    let closingBookWdv = parseAmount(r[safeMapping.closingWdvCol]);

    let usefulLifeYears = parseAmount(r[safeMapping.usefulLifeCol]);
    if (usefulLifeYears <= 0) {
      usefulLifeYears = standardGuide.defaultUsefulLifeYears;
    }

    // Auto-derive opening WDV if blank
    if (openingWdv === 0 && grossCost > 0) {
      openingWdv = Math.max(0, grossCost - openingAccumDepr);
    }
    // Auto-derive opening accum depr if blank
    if (openingAccumDepr === 0 && grossCost > 0 && openingWdv > 0) {
      openingAccumDepr = Math.max(0, grossCost - openingWdv);
    }
    // Auto-derive closing WDV if blank
    if (closingBookWdv === 0 && openingWdv > 0) {
      closingBookWdv = Math.max(0, openingWdv - bookDepreciation);
    }

    const assetIssues: FarIssue[] = [];

    // 1. Duplicate Asset ID check
    if (assetIdCounts.get(assetId)! > 1) {
      duplicateIdsCount++;
      const iss: FarIssue = {
        id: `iss-far-${issues.length + 1}`,
        assetId,
        rowIndex,
        category: 'duplicate_asset_id',
        severity: 'critical',
        title: `Duplicate Asset ID: ${assetId}`,
        description: `Asset ID "${assetId}" is duplicated across ${assetIdCounts.get(assetId)} rows in the register.`,
        suggestedAction: 'Assign unique asset tag identifiers to ensure physical tagging and verification integrity.',
        bookValue: grossCost,
        recomputedValue: grossCost,
        variance: 0,
      };
      issues.push(iss);
      assetIssues.push(iss);
    }

    // 2. Missing Capitalization Date
    if (!capitalizationDate || !capDateObj) {
      const iss: FarIssue = {
        id: `iss-far-${issues.length + 1}`,
        assetId,
        rowIndex,
        category: 'missing_date',
        severity: 'critical',
        title: 'Missing Put-to-Use / Capitalization Date',
        description: `Asset "${assetDescription}" lacks a valid capitalization date. Pro-rata depreciation cannot be certified.`,
        suggestedAction: 'Retrieve installation certificate, purchase invoice, or transport GR to establish capitalization date.',
        bookValue: 0,
        recomputedValue: 0,
        variance: 0,
      };
      issues.push(iss);
      assetIssues.push(iss);
    }

    // 3. Missing or Zero Historical Cost
    if (grossCost <= 0) {
      const iss: FarIssue = {
        id: `iss-far-${issues.length + 1}`,
        assetId,
        rowIndex,
        category: 'missing_cost',
        severity: 'critical',
        title: 'Missing or Non-Positive Historical Cost',
        description: `Gross cost recorded is ₹${grossCost}. Fixed assets must possess an audited historical cost.`,
        suggestedAction: 'Check original procurement invoice and capitalize eligible freight and installation charges.',
        bookValue: grossCost,
        recomputedValue: 0,
        variance: grossCost,
      };
      issues.push(iss);
      assetIssues.push(iss);
    }

    // 4. Inconsistent Useful Life check under Schedule II
    const maxAllowedLife = standardGuide.defaultUsefulLifeYears;
    if (usefulLifeYears > maxAllowedLife * 1.5 || usefulLifeYears < maxAllowedLife * 0.5) {
      const iss: FarIssue = {
        id: `iss-far-${issues.length + 1}`,
        assetId,
        rowIndex,
        category: 'inconsistent_useful_life',
        severity: 'warning',
        title: `Inconsistent Useful Life (${usefulLifeYears} yrs vs Std ${maxAllowedLife} yrs)`,
        description: `Asset categorized as "${standardGuide.category}" specifies useful life of ${usefulLifeYears} years. Standard Schedule II life is ${maxAllowedLife} years.`,
        suggestedAction: 'Require formal chartered engineer technical justification if departing from Schedule II useful life.',
        bookValue: usefulLifeYears,
        recomputedValue: maxAllowedLife,
        variance: usefulLifeYears - maxAllowedLife,
      };
      issues.push(iss);
      assetIssues.push(iss);
    }

    // 5. Depreciation Recomputation
    let isAdditionDuringYear = false;
    let daysUsedInYear = 365;
    let isPutToUseLessThan180Days = false;

    if (capDateObj) {
      if (capDateObj >= fyStart && capDateObj <= fyEnd) {
        isAdditionDuringYear = true;
        const diffMs = fyEnd.getTime() - capDateObj.getTime();
        daysUsedInYear = Math.max(1, Math.min(365, Math.floor(diffMs / (1000 * 60 * 60 * 24)) + 1));
        if (capDateObj >= octFirst) {
          isPutToUseLessThan180Days = true;
        }
      } else if (capDateObj > fyEnd) {
        // Asset acquired in future period!
        daysUsedInYear = 0;
      }
    }

    // Compute residual value (default 5% under Schedule II)
    const residualVal = grossCost * (activeConfig.residualValuePercent / 100);
    const depreciableBase = Math.max(0, grossCost - residualVal);

    let recomputedDepreciation = 0;

    if (activeConfig.deprMethod === 'companies_act_slm') {
      const annualSlm = usefulLifeYears > 0 ? depreciableBase / usefulLifeYears : 0;
      if (isAdditionDuringYear) {
        recomputedDepreciation = Math.round((annualSlm * daysUsedInYear) / 365);
      } else {
        recomputedDepreciation = Math.round(annualSlm);
      }
    } else if (activeConfig.deprMethod === 'companies_act_wdv') {
      // Schedule II WDV Rate: R = 1 - (s/c)^(1/n)
      let wdvRate = 0;
      if (grossCost > 0 && usefulLifeYears > 0) {
        wdvRate = 1 - Math.pow(activeConfig.residualValuePercent / 100, 1 / usefulLifeYears);
      }
      const baseForWdv = openingWdv > 0 ? openingWdv : grossCost;
      if (isAdditionDuringYear) {
        recomputedDepreciation = Math.round(((baseForWdv * wdvRate) * daysUsedInYear) / 365);
      } else {
        recomputedDepreciation = Math.round(baseForWdv * wdvRate);
      }
    } else {
      // Income Tax Act Sec 32 Block Rate
      const itRate = standardGuide.itBlockRatePercent / 100;
      const baseVal = openingWdv > 0 ? openingWdv : grossCost;
      if (isAdditionDuringYear && isPutToUseLessThan180Days && activeConfig.enforce180DayRule) {
        // Half-rate rule
        recomputedDepreciation = Math.round(baseVal * (itRate / 2));
      } else {
        recomputedDepreciation = Math.round(baseVal * itRate);
      }
    }

    // Cap depreciation so asset is not depreciated below residual value (95% rule)
    const maxDepreciableAllowed = Math.max(0, (openingWdv || grossCost) - residualVal);
    if (recomputedDepreciation > maxDepreciableAllowed) {
      recomputedDepreciation = maxDepreciableAllowed;
    }

    const depreciationVariance = Math.abs(bookDepreciation - recomputedDepreciation);
    const recomputedClosingWdv = Math.max(residualVal, (openingWdv || grossCost) - recomputedDepreciation);

    // 6. Check Depreciation Variance (> ₹50 tolerance)
    if (depreciationVariance > 50 && grossCost > 0) {
      depreciationDiscrepanciesCount++;
      const iss: FarIssue = {
        id: `iss-far-${issues.length + 1}`,
        assetId,
        rowIndex,
        category: 'depreciation_mismatch',
        severity: 'critical',
        title: `Depreciation Mismatch (Variance: ₹${depreciationVariance.toLocaleString('en-IN')})`,
        description: `Book depreciation claimed is ₹${bookDepreciation.toLocaleString('en-IN')}, whereas audited recomputed depreciation is ₹${recomputedDepreciation.toLocaleString('en-IN')}.`,
        suggestedAction: 'Review whether pro-rata days or residual value 5% was omitted in client register calculation.',
        bookValue: bookDepreciation,
        recomputedValue: recomputedDepreciation,
        variance: bookDepreciation - recomputedDepreciation,
      };
      issues.push(iss);
      assetIssues.push(iss);
    }

    // 7. Check 95% Limit Violation (Asset depreciated below 5% residual value)
    if (closingBookWdv < residualVal && grossCost > 0) {
      const iss: FarIssue = {
        id: `iss-far-${issues.length + 1}`,
        assetId,
        rowIndex,
        category: 'exceeded_95_percent_limit',
        severity: 'warning',
        title: 'Closing WDV Below 5% Residual Value Limit',
        description: `Closing WDV of ₹${closingBookWdv.toLocaleString('en-IN')} is less than 5% residual value (₹${residualVal.toLocaleString('en-IN')}).`,
        suggestedAction: 'Under Schedule II, an asset shall retain 5% of gross cost as residual value unless technical justification exists.',
        bookValue: closingBookWdv,
        recomputedValue: residualVal,
        variance: residualVal - closingBookWdv,
      };
      issues.push(iss);
      assetIssues.push(iss);
    }

    // 8. Addition during year claimed for full year instead of half year or pro-rata
    if (isAdditionDuringYear && isPutToUseLessThan180Days && bookDepreciation > recomputedDepreciation * 1.5) {
      const iss: FarIssue = {
        id: `iss-far-${issues.length + 1}`,
        assetId,
        rowIndex,
        category: 'full_year_on_new_addition',
        severity: 'critical',
        title: 'Full Year Depreciation Claimed on Mid-Year Addition',
        description: `Asset put to use on ${capitalizationDate} (< 180 days in FY), but full-year depreciation appears to have been claimed.`,
        suggestedAction: 'Apply pro-rata days calculation (Companies Act) or 50% half-rate rule (Income Tax Sec 32).',
        bookValue: bookDepreciation,
        recomputedValue: recomputedDepreciation,
        variance: bookDepreciation - recomputedDepreciation,
      };
      issues.push(iss);
      assetIssues.push(iss);
    }

    totalGrossCost += grossCost;
    totalOpeningAccumDepr += openingAccumDepr;
    totalBookDepreciation += bookDepreciation;
    totalRecomputedDepreciation += recomputedDepreciation;
    totalClosingBookWdv += closingBookWdv;

    assets.push({
      id: `far-${idx + 1}`,
      rowIndex,
      assetId,
      assetDescription,
      category: rawCategory || standardGuide.category,
      matchedCategoryKey,
      capitalizationDate,
      capitalizationDateObj: capDateObj,
      grossCost,
      openingAccumDepr,
      openingWdv,
      bookDepreciation,
      closingBookWdv,
      usefulLifeYears,
      deprMethod: activeConfig.deprMethod,
      recomputedDepreciation,
      depreciationVariance,
      recomputedClosingWdv,
      isAdditionDuringYear,
      daysUsedInYear,
      isPutToUseLessThan180Days,
      issues: assetIssues,
      raw: r,
    });
  });

  const criticalIssuesCount = issues.filter((i) => i.severity === 'critical').length;
  const warningIssuesCount = issues.filter((i) => i.severity === 'warning').length;
  const netDepreciationVariance = Math.abs(totalBookDepreciation - totalRecomputedDepreciation);

  return {
    summary: {
      totalAssets: assets.length,
      totalAssetsCount: assets.length,
      totalGrossCost,
      totalGrossBlock: totalGrossCost,
      totalOpeningAccumDepr,
      totalBookDepreciation,
      totalDepreciation: totalBookDepreciation,
      totalRecomputedDepreciation,
      netDepreciationVariance,
      totalClosingBookWdv,
      criticalIssuesCount,
      warningIssuesCount,
      duplicateIdsCount,
      depreciationDiscrepanciesCount,
      issues,
    },
    assets,
  };
}

/**
 * Exports Audit Review Report and Corrected Fixed Asset Register
 */
export function exportFarAuditWorkbook(
  summary: FarAuditSummary,
  assets: FarAssetRecord[],
  clientName = 'Audit Client',
  period = 'FY 2026-27'
): XLSX.WorkBook {
  const wb = XLSX.utils.book_new();

  // Sheet 1: Executive FAR Audit Summary
  const summaryData = [
    { Parameter: 'Client Entity', Value: clientName },
    { Parameter: 'Financial Period', Value: period },
    { Parameter: 'Total Fixed Assets Audited', Value: summary.totalAssets },
    { Parameter: 'Total Gross Block / Historical Cost (₹)', Value: summary.totalGrossCost },
    { Parameter: 'Opening Accumulated Depreciation (₹)', Value: summary.totalOpeningAccumDepr },
    { Parameter: 'Current Year Book Depreciation Claimed (₹)', Value: summary.totalBookDepreciation },
    { Parameter: 'Audited Recomputed Depreciation (₹)', Value: summary.totalRecomputedDepreciation },
    { Parameter: 'Net Depreciation Variance (₹)', Value: summary.netDepreciationVariance },
    { Parameter: 'Closing Net Book Value / WDV (₹)', Value: summary.totalClosingBookWdv },
    { Parameter: 'Duplicate Asset ID Violations', Value: summary.duplicateIdsCount },
    { Parameter: 'Depreciation Discrepancy Flags', Value: summary.depreciationDiscrepanciesCount },
    { Parameter: 'Critical Audit Exceptions', Value: summary.criticalIssuesCount },
    { Parameter: 'Audit Standard Compliance', Value: 'Schedule II, Companies Act 2013 & Sec 32 Income Tax Act' },
    { Parameter: 'Integrity Rule Note', Value: 'Original source register preserved intact. No modifications applied without sign-off.' },
  ];
  const wsSummary = createStyledWorksheet(summaryData, {
    currencyColumns: ['Value'],
    columnWidths: { Parameter: 40, Value: 34 },
  });
  XLSX.utils.book_append_sheet(wb, wsSummary, 'FAR_Audit_Summary');

  // Sheet 2: Audit Findings & Issues
  const issueRows = summary.issues.map((iss, idx) => ({
    'Issue #': idx + 1,
    'Severity': iss.severity.toUpperCase(),
    'Category': iss.category.replace(/_/g, ' ').toUpperCase(),
    'Asset Tag #': iss.assetId,
    'Row #': iss.rowIndex,
    'Audit Finding': iss.title,
    'Detailed Observation': iss.description,
    'Book Value (₹)': iss.bookValue,
    'Recomputed Value (₹)': iss.recomputedValue,
    'Variance (₹)': iss.variance,
    'Suggested Corrective Action': iss.suggestedAction,
  }));
  const wsIssues = createStyledWorksheet(issueRows, {
    currencyColumns: ['Book Value (₹)', 'Recomputed Value (₹)', 'Variance (₹)'],
  });
  XLSX.utils.book_append_sheet(wb, wsIssues, 'Audit_Findings');

  // Sheet 3: Depreciation Recomputation Schedule
  const scheduleRows = assets.map((a) => ({
    'Row #': a.rowIndex,
    'Asset ID': a.assetId,
    'Description': a.assetDescription,
    'Asset Class': a.category,
    'Put-to-Use Date': a.capitalizationDate,
    'Days Used in FY': a.daysUsedInYear,
    'Gross Cost (₹)': a.grossCost,
    'Opening WDV (₹)': a.openingWdv,
    'Book Depr Claimed (₹)': a.bookDepreciation,
    'Audited Depr (₹)': a.recomputedDepreciation,
    'Variance (₹)': a.bookDepreciation - a.recomputedDepreciation,
    'Book Closing WDV (₹)': a.closingBookWdv,
    'Audited Closing WDV (₹)': a.recomputedClosingWdv,
    'Status': a.issues.length > 0 ? `${a.issues.length} FLAGS DETECTED` : 'VERIFIED CLEAN',
  }));
  const wsSchedule = createStyledWorksheet(scheduleRows, {
    currencyColumns: [
      'Gross Cost (₹)',
      'Opening WDV (₹)',
      'Book Depr Claimed (₹)',
      'Audited Depr (₹)',
      'Variance (₹)',
      'Book Closing WDV (₹)',
      'Audited Closing WDV (₹)',
    ],
  });
  XLSX.utils.book_append_sheet(wb, wsSchedule, 'Depr_Recomputation');

  // Sheet 4: Corrected Register Copy (Approved Schedule II Asset Register)
  const correctedRows = assets.map((a) => ({
    'Asset Tag': a.assetId,
    'Asset Description': a.assetDescription,
    'Block Category': a.category,
    'Capitalization Date': a.capitalizationDate,
    'Useful Life (Yrs)': a.usefulLifeYears,
    'Historical Cost (₹)': a.grossCost,
    'Opening WDV (₹)': a.openingWdv,
    'Audited Depreciation (₹)': a.recomputedDepreciation,
    'Corrected Closing WDV (₹)': a.recomputedClosingWdv,
    'Residual Value (5%) (₹)': Math.round(a.grossCost * 0.05),
    'Audit Certification': 'Verified under Schedule II Companies Act 2013',
  }));
  const wsCorrected = createStyledWorksheet(correctedRows, {
    currencyColumns: [
      'Historical Cost (₹)',
      'Opening WDV (₹)',
      'Audited Depreciation (₹)',
      'Corrected Closing WDV (₹)',
      'Residual Value (5%) (₹)',
    ],
  });
  XLSX.utils.book_append_sheet(wb, wsCorrected, 'Approved_Corrected_FAR');

  return wb;
}

/**
 * Realistic sample fixed asset register dataset with deliberate audit issues
 */
export function getSampleFixedAssetRows(): Record<string, any>[] {
  return [
    {
      'Asset ID': 'FA-PLANT-001',
      'Asset Description': 'CNC Milling Machine - 5 Axis',
      'Category': 'Plant & Machinery',
      'Capitalization Date': '15-05-2022',
      'Gross Cost': 2800000,
      'Opening WDV': 2185000,
      'Book Depreciation': 177333,
      'Closing WDV': 2007667,
      'Useful Life': 15,
      'Method': 'SLM',
    },
    {
      'Asset ID': 'FA-PLANT-002',
      'Asset Description': 'Hydraulic Press 100 Ton',
      'Category': 'Plant & Machinery',
      'Capitalization Date': '10-08-2023',
      'Gross Cost': 1450000,
      'Opening WDV': 1285000,
      'Book Depreciation': 91833,
      'Closing WDV': 1193167,
      'Useful Life': 15,
      'Method': 'SLM',
    },
    {
      'Asset ID': 'FA-IT-010',
      'Asset Description': 'Dell Latitude Executive Laptops (Batch of 5)',
      'Category': 'Computers & Laptops',
      'Capitalization Date': '02-06-2024',
      'Gross Cost': 450000,
      'Opening WDV': 307500,
      'Book Depreciation': 45000, // ERROR: Depreciated over 10 years instead of 3 years!
      'Closing WDV': 262500,
      'Useful Life': 10, // WRONG USEFUL LIFE (Std is 3 yrs)
      'Method': 'SLM',
    },
    {
      'Asset ID': 'FA-IT-010', // DUPLICATE ASSET TAG NUMBER
      'Asset Description': 'Apple MacBook Pro M3 for Design Team',
      'Category': 'Computers & Laptops',
      'Capitalization Date': '14-07-2025',
      'Gross Cost': 210000,
      'Opening WDV': 143500,
      'Book Depreciation': 66500,
      'Closing WDV': 77000,
      'Useful Life': 3,
      'Method': 'SLM',
    },
    {
      'Asset ID': 'FA-VEH-004',
      'Asset Description': 'Toyota Innova Crysta Staff Van',
      'Category': 'Motor Vehicles',
      'Capitalization Date': '20-11-2026', // ADDITION DURING YEAR (< 180 DAYS)
      'Gross Cost': 2400000,
      'Opening WDV': 2400000,
      'Book Depreciation': 285000, // ERROR: Claimed FULL YEAR depreciation instead of pro-rata days (~131 days)!
      'Closing WDV': 2115000,
      'Useful Life': 8,
      'Method': 'SLM',
    },
    {
      'Asset ID': 'FA-FUR-008',
      'Asset Description': 'Conference Room Modular Table & Chairs',
      'Category': 'Furniture & Fixtures',
      'Capitalization Date': '18-04-2021',
      'Gross Cost': 320000,
      'Opening WDV': 224000,
      'Book Depreciation': 30400,
      'Closing WDV': 193600,
      'Useful Life': 10,
      'Method': 'SLM',
    },
    {
      'Asset ID': 'FA-OFF-012',
      'Asset Description': 'Heavy Duty Kyocera Multi-function Copier',
      'Category': 'Office Equipment',
      'Capitalization Date': '', // MISSING CAPITALIZATION DATE
      'Gross Cost': 180000,
      'Opening WDV': 114000,
      'Book Depreciation': 34200,
      'Closing WDV': 79800,
      'Useful Life': 5,
      'Method': 'SLM',
    },
    {
      'Asset ID': 'FA-PLANT-009',
      'Asset Description': 'Submersible Water Cooling Plant',
      'Category': 'Plant & Machinery',
      'Capitalization Date': '10-01-2015',
      'Gross Cost': 500000,
      'Opening WDV': 35000,
      'Book Depreciation': 31666, // ERROR: Closing WDV drops to 3,334 which is below 5% residual value (₹25,000)!
      'Closing WDV': 3334,
      'Useful Life': 15,
      'Method': 'SLM',
    },
    {
      'Asset ID': 'FA-IT-022',
      'Asset Description': 'HPE ProLiant Rack Server DL380',
      'Category': 'Servers & Networks',
      'Capitalization Date': '05-09-2025',
      'Gross Cost': 620000,
      'Opening WDV': 521917,
      'Book Depreciation': 98167,
      'Closing WDV': 423750,
      'Useful Life': 6,
      'Method': 'SLM',
    },
    {
      'Asset ID': 'FA-BLD-001',
      'Asset Description': 'Industrial Factory Shed - Unit 2',
      'Category': 'Buildings (Factory)',
      'Capitalization Date': '01-04-2018',
      'Gross Cost': 12500000,
      'Opening WDV': 9114583,
      'Book Depreciation': 395833,
      'Closing WDV': 8718750,
      'Useful Life': 30,
      'Method': 'SLM',
    },
  ];
}
