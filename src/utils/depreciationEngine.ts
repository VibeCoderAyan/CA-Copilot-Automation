import * as XLSX from 'xlsx';
import { parseDateObj } from './dataSanitizer';
import { createStyledWorksheet } from './excelStyler';

export type AssetCategory =
  | 'computers'
  | 'servers'
  | 'plant_machinery'
  | 'furniture'
  | 'office_equipment'
  | 'motor_cars'
  | 'commercial_vehicles'
  | 'factory_building'
  | 'office_building'
  | 'software_intangibles';

export interface AssetCategorySpec {
  name: string;
  usefulLifeYears: number; // Companies Act 2013 Schedule II
  companiesActWdvRate: number; // %
  companiesActSlmRate: number; // %
  itActBlockRate: number; // % Section 32 Income Tax Act
}

export const ASSET_CATEGORIES: Record<AssetCategory, AssetCategorySpec> = {
  computers: {
    name: 'Computers & Laptops',
    usefulLifeYears: 3,
    companiesActWdvRate: 63.16,
    companiesActSlmRate: 31.67,
    itActBlockRate: 40.0,
  },
  servers: {
    name: 'Servers & Network Equipment',
    usefulLifeYears: 6,
    companiesActWdvRate: 39.30,
    companiesActSlmRate: 15.83,
    itActBlockRate: 40.0,
  },
  plant_machinery: {
    name: 'Plant & Machinery (General)',
    usefulLifeYears: 15,
    companiesActWdvRate: 18.10,
    companiesActSlmRate: 6.33,
    itActBlockRate: 15.0,
  },
  furniture: {
    name: 'Furniture & Fixtures',
    usefulLifeYears: 10,
    companiesActWdvRate: 25.89,
    companiesActSlmRate: 9.50,
    itActBlockRate: 10.0,
  },
  office_equipment: {
    name: 'Office Equipment',
    usefulLifeYears: 5,
    companiesActWdvRate: 45.07,
    companiesActSlmRate: 19.0,
    itActBlockRate: 15.0,
  },
  motor_cars: {
    name: 'Motor Cars (Non-Commercial)',
    usefulLifeYears: 8,
    companiesActWdvRate: 31.23,
    companiesActSlmRate: 11.88,
    itActBlockRate: 15.0,
  },
  commercial_vehicles: {
    name: 'Commercial Vehicles (Lorries/Buses)',
    usefulLifeYears: 10,
    companiesActWdvRate: 25.89,
    companiesActSlmRate: 9.50,
    itActBlockRate: 30.0,
  },
  factory_building: {
    name: 'Factory Buildings (RCC)',
    usefulLifeYears: 30,
    companiesActWdvRate: 9.50,
    companiesActSlmRate: 3.17,
    itActBlockRate: 10.0,
  },
  office_building: {
    name: 'Office Buildings & Residential',
    usefulLifeYears: 60,
    companiesActWdvRate: 4.87,
    companiesActSlmRate: 1.58,
    itActBlockRate: 5.0,
  },
  software_intangibles: {
    name: 'Software & Intangible Assets',
    usefulLifeYears: 5,
    companiesActWdvRate: 45.07,
    companiesActSlmRate: 19.0,
    itActBlockRate: 25.0,
  },
};

export interface AssetRecord {
  id: string;
  code: string;
  name: string;
  category: AssetCategory;
  purchaseDate: string;
  putToUseDate: string;
  originalCost: number;
  residualValue: number; // default 5%
  openingWdv: number;
  additionsMoreThan180Days: number;
  additionsLessThan180Days: number;
  deletionsDeductions: number;
  saleRealization: number;
  method: 'WDV' | 'SLM';
}

export interface DepreciationCalculationResult {
  assetId: string;
  assetName: string;
  categoryName: string;
  // Companies Act 2013 Schedule II
  coActMethod: 'WDV' | 'SLM';
  coActUsefulLife: number;
  coActOpeningWdv: number;
  coActTotalAdditions: number;
  coActDeletions: number;
  coActDepreciation: number;
  coActClosingWdv: number;
  // Income Tax Act 1961 Section 32
  itActBlockRate: number;
  itActOpeningWdv: number;
  itActAdditionsMore180: number;
  itActAdditionsLess180: number;
  itActDeductions: number;
  itActDepreciation: number;
  itActClosingWdv: number;
  // Timing Difference (AS 22 / Ind AS 12)
  timingDifference: number; // coActDepreciation - itActDepreciation
  taxEffect: number; // timingDifference * 25.17% (corporate tax rate)
  validationWarnings: string[];
}

/**
 * Calculates both Companies Act and Income Tax Act depreciation for an asset
 */
export function calculateAssetDepreciation(
  asset: AssetRecord,
  corporateTaxRate: number = 0.2517 // 25.17% standard base tax rate
): DepreciationCalculationResult {
  const spec = ASSET_CATEGORIES[asset.category] || ASSET_CATEGORIES.plant_machinery;
  const warnings: string[] = [];

  if (asset.originalCost <= 0) {
    warnings.push('Asset original cost must be greater than zero.');
  }
  if (asset.residualValue > asset.originalCost) {
    warnings.push('Residual value cannot exceed original asset cost.');
  }

  // Determine addition split if dates are provided (Indian FY: April 1 to March 31)
  const putToUse = parseDateObj(asset.putToUseDate || asset.purchaseDate);
  let isLessThan180 = false;
  if (putToUse && !isNaN(putToUse.getTime())) {
    const month = putToUse.getMonth(); // 0 = Jan, 2 = Mar, 3 = Apr, 9 = Oct
    const year = putToUse.getFullYear();
    // If put to use in Jan-Mar (months 0, 1, 2), FY began on April 1 of previous year (year - 1).
    // Cutoff (180 days before March 31) is 4th October of that FY.
    const fyStartYear = month < 3 ? year - 1 : year;
    const cutoffDate = new Date(fyStartYear, 9, 4); // 4th October
    isLessThan180 = putToUse >= cutoffDate;
  }

  // Companies Act Calculation (Schedule II)
  const coActOpening = asset.openingWdv > 0 ? asset.openingWdv : asset.originalCost;
  const coActAdditions = asset.additionsMoreThan180Days + asset.additionsLessThan180Days;
  const coActDeletions = asset.deletionsDeductions;
  const coActBase = Math.max(0, coActOpening + coActAdditions - coActDeletions);

  let coActDepreciation = 0;
  if (asset.method === 'SLM') {
    const depreciableAmount = Math.max(0, asset.originalCost - asset.residualValue);
    const annualSlm = depreciableAmount / spec.usefulLifeYears;
    const additionMoreDepr = asset.additionsMoreThan180Days * (spec.companiesActSlmRate / 100);
    const additionLessDepr = asset.additionsLessThan180Days * (spec.companiesActSlmRate / 100) * 0.5;
    coActDepreciation = annualSlm + additionMoreDepr + additionLessDepr;
  } else {
    // WDV method
    const rate = spec.companiesActWdvRate / 100;
    const deprOpening = coActOpening * rate;
    const deprMore180 = asset.additionsMoreThan180Days * rate;
    const deprLess180 = asset.additionsLessThan180Days * rate * 0.5;
    coActDepreciation = deprOpening + deprMore180 + deprLess180;
  }
  // Schedule II Rule: Net book value cannot be depreciated below residual value (5% salvage)
  const maxPermissibleDepreciation = Math.max(0, coActBase - asset.residualValue);
  coActDepreciation = Math.min(coActDepreciation, maxPermissibleDepreciation);
  coActDepreciation = Math.round(coActDepreciation * 100) / 100;
  const coActClosing = Math.round((coActBase - coActDepreciation) * 100) / 100;

  // Income Tax Act Section 32 Calculation (Block of Assets)
  const itActOpening = asset.openingWdv > 0 ? asset.openingWdv : asset.originalCost;
  const itRate = spec.itActBlockRate / 100;
  const itActMore180 = asset.additionsMoreThan180Days;
  const itActLess180 = asset.additionsLessThan180Days;
  const itActDeductions = asset.saleRealization > 0 ? asset.saleRealization : asset.deletionsDeductions;

  // Sec 32 rule: Net block before depreciation
  const netBlock = itActOpening + itActMore180 + itActLess180 - itActDeductions;
  let itActDepreciation = 0;

  if (netBlock > 0) {
    // If additions < 180 days exist, they get 50% of the prescribed rate
    // But if netBlock < additions < 180 days, depreciation is capped at 50% rate on netBlock
    if (itActLess180 > 0) {
      const eligibleForHalf = Math.min(netBlock, itActLess180);
      const eligibleForFull = Math.max(0, netBlock - eligibleForHalf);
      itActDepreciation = eligibleForFull * itRate + eligibleForHalf * (itRate * 0.5);
    } else {
      itActDepreciation = netBlock * itRate;
    }
  }
  itActDepreciation = Math.round(itActDepreciation * 100) / 100;
  const itActClosing = Math.max(0, netBlock - itActDepreciation);

  // Timing difference (Accounting Depreciation vs Tax Depreciation)
  // Positive means Accounting depr > Tax depr => Deferred Tax Asset (DTA)
  // Negative means Tax depr > Accounting depr => Deferred Tax Liability (DTL)
  const timingDifference = Math.round((coActDepreciation - itActDepreciation) * 100) / 100;
  const taxEffect = Math.round(timingDifference * corporateTaxRate * 100) / 100;

  return {
    assetId: asset.id,
    assetName: asset.name,
    categoryName: spec.name,
    coActMethod: asset.method,
    coActUsefulLife: spec.usefulLifeYears,
    coActOpeningWdv: Math.round(coActOpening),
    coActTotalAdditions: Math.round(coActAdditions),
    coActDeletions: Math.round(coActDeletions),
    coActDepreciation,
    coActClosingWdv: Math.round(coActClosing),
    itActBlockRate: spec.itActBlockRate,
    itActOpeningWdv: Math.round(itActOpening),
    itActAdditionsMore180: Math.round(itActMore180),
    itActAdditionsLess180: Math.round(itActLess180),
    itActDeductions: Math.round(itActDeductions),
    itActDepreciation,
    itActClosingWdv: Math.round(itActClosing),
    timingDifference,
    taxEffect,
    validationWarnings: warnings,
  };
}

/**
 * Returns realistic default sample assets for testing
 */
export function getSampleDepreciationAssets(): AssetRecord[] {
  return [
    {
      id: 'ast-1',
      code: 'FA-CMP-01',
      name: 'MacBook Pro & Dell Developer Laptops',
      category: 'computers',
      purchaseDate: '2025-06-15',
      putToUseDate: '2025-06-20',
      originalCost: 650000,
      residualValue: 32500,
      openingWdv: 420000,
      additionsMoreThan180Days: 120000,
      additionsLessThan180Days: 85000,
      deletionsDeductions: 0,
      saleRealization: 0,
      method: 'WDV',
    },
    {
      id: 'ast-2',
      code: 'FA-SRV-01',
      name: 'HPE ProLiant Rack Server 32-Core',
      category: 'servers',
      purchaseDate: '2024-04-10',
      putToUseDate: '2024-04-15',
      originalCost: 1200000,
      residualValue: 60000,
      openingWdv: 728000,
      additionsMoreThan180Days: 0,
      additionsLessThan180Days: 0,
      deletionsDeductions: 0,
      saleRealization: 0,
      method: 'WDV',
    },
    {
      id: 'ast-3',
      code: 'FA-PLM-01',
      name: 'CNC Milling & Lathe Machine Unit A',
      category: 'plant_machinery',
      purchaseDate: '2023-11-20',
      putToUseDate: '2023-11-25',
      originalCost: 4500000,
      residualValue: 225000,
      openingWdv: 3680000,
      additionsMoreThan180Days: 450000,
      additionsLessThan180Days: 200000,
      deletionsDeductions: 150000,
      saleRealization: 120000,
      method: 'SLM',
    },
    {
      id: 'ast-4',
      code: 'FA-FNT-01',
      name: 'Executive Ergonomic Workstations & Conference Tables',
      category: 'furniture',
      purchaseDate: '2024-08-01',
      putToUseDate: '2024-08-05',
      originalCost: 850000,
      residualValue: 42500,
      openingWdv: 630000,
      additionsMoreThan180Days: 60000,
      additionsLessThan180Days: 30000,
      deletionsDeductions: 0,
      saleRealization: 0,
      method: 'WDV',
    },
    {
      id: 'ast-5',
      code: 'FA-VEH-01',
      name: 'Toyota Innova Crysta (Company Car)',
      category: 'motor_cars',
      purchaseDate: '2025-01-10',
      putToUseDate: '2025-01-15',
      originalCost: 2400000,
      residualValue: 120000,
      openingWdv: 1980000,
      additionsMoreThan180Days: 0,
      additionsLessThan180Days: 0,
      deletionsDeductions: 0,
      saleRealization: 0,
      method: 'WDV',
    },
  ];
}

/**
 * Exports formatted 4-sheet Excel depreciation report
 */
export function exportDepreciationWorkbook(
  assets: AssetRecord[],
  results: DepreciationCalculationResult[],
  entityName: string = 'Apex Industries Ltd',
  financialYear: string = 'FY 2026-27'
): XLSX.WorkBook {
  const wb = XLSX.utils.book_new();

  // Sheet 1: Schedule II (Companies Act 2013)
  const coActData = results.map((r, idx) => ({
    'Sl No': idx + 1,
    'Asset Code': assets[idx]?.code || `AST-${idx + 1}`,
    'Asset Description': r.assetName,
    'Category': r.categoryName,
    'Method': r.coActMethod,
    'Useful Life (Yrs)': r.coActUsefulLife,
    'Opening Gross Block': assets[idx]?.originalCost || 0,
    'Opening WDV (₹)': r.coActOpeningWdv,
    'Additions (₹)': r.coActTotalAdditions,
    'Deletions (₹)': r.coActDeletions,
    'Depreciation for Year (₹)': r.coActDepreciation,
    'Closing WDV (₹)': r.coActClosingWdv,
  }));

  const wsCoAct = createStyledWorksheet(coActData, {
    currencyColumns: ['Opening Gross Block', 'Opening WDV (₹)', 'Additions (₹)', 'Deletions (₹)', 'Depreciation for Year (₹)', 'Closing WDV (₹)'],
    columnWidths: {
      'Sl No': 8,
      'Asset Code': 14,
      'Asset Description': 35,
      'Category': 24,
      'Method': 12,
      'Useful Life (Yrs)': 16,
      'Opening Gross Block': 20,
      'Opening WDV (₹)': 20,
      'Additions (₹)': 18,
      'Deletions (₹)': 16,
      'Depreciation for Year (₹)': 22,
      'Closing WDV (₹)': 20,
    },
  });
  XLSX.utils.book_append_sheet(wb, wsCoAct, 'Companies Act Schedule II');

  // Sheet 2: Tax Block of Assets (Section 32 Income Tax Act)
  const itActData = results.map((r, idx) => ({
    'Sl No': idx + 1,
    'Asset / Block Name': r.assetName,
    'Prescribed IT Rate (%)': `${r.itActBlockRate}%`,
    'Opening WDV (₹)': r.itActOpeningWdv,
    'Additions >= 180 Days (₹)': r.itActAdditionsMore180,
    'Additions < 180 Days (₹)': r.itActAdditionsLess180,
    'Moneys Payable / Sale (₹)': r.itActDeductions,
    'Tax Depreciation (₹)': r.itActDepreciation,
    'Closing WDV (₹)': r.itActClosingWdv,
  }));

  const wsItAct = createStyledWorksheet(itActData, {
    currencyColumns: ['Opening WDV (₹)', 'Additions >= 180 Days (₹)', 'Additions < 180 Days (₹)', 'Moneys Payable / Sale (₹)', 'Tax Depreciation (₹)', 'Closing WDV (₹)'],
    columnWidths: {
      'Sl No': 8,
      'Asset / Block Name': 35,
      'Prescribed IT Rate (%)': 22,
      'Opening WDV (₹)': 20,
      'Additions >= 180 Days (₹)': 22,
      'Additions < 180 Days (₹)': 22,
      'Moneys Payable / Sale (₹)': 22,
      'Tax Depreciation (₹)': 22,
      'Closing WDV (₹)': 20,
    },
  });
  XLSX.utils.book_append_sheet(wb, wsItAct, 'IT Act Sec 32 Block');

  // Sheet 3: Deferred Tax Variance (AS 22 / Ind AS 12)
  const varianceData = results.map((r, idx) => ({
    'Sl No': idx + 1,
    'Asset Name': r.assetName,
    'Companies Act Depr (₹)': r.coActDepreciation,
    'Income Tax Depr (₹)': r.itActDepreciation,
    'Timing Difference (₹)': r.timingDifference,
    'Nature of Timing Diff': r.timingDifference > 0 ? 'Deferred Tax Asset (DTA)' : 'Deferred Tax Liability (DTL)',
    'Tax Effect @ 25.17% (₹)': Math.abs(r.taxEffect),
  }));

  const wsVariance = createStyledWorksheet(varianceData, {
    currencyColumns: ['Companies Act Depr (₹)', 'Income Tax Depr (₹)', 'Timing Difference (₹)', 'Tax Effect @ 25.17% (₹)'],
    columnWidths: {
      'Sl No': 8,
      'Asset Name': 35,
      'Companies Act Depr (₹)': 22,
      'Income Tax Depr (₹)': 22,
      'Timing Difference (₹)': 22,
      'Nature of Timing Diff': 26,
      'Tax Effect @ 25.17% (₹)': 22,
    },
  });
  XLSX.utils.book_append_sheet(wb, wsVariance, 'Deferred Tax Timing Diff');

  // Sheet 4: Executive Summary
  const totalCoActDepr = results.reduce((acc, r) => acc + r.coActDepreciation, 0);
  const totalItActDepr = results.reduce((acc, r) => acc + r.itActDepreciation, 0);
  const totalCoActClosing = results.reduce((acc, r) => acc + r.coActClosingWdv, 0);
  const totalItActClosing = results.reduce((acc, r) => acc + r.itActClosingWdv, 0);

  const summaryData = [
    { Parameter: 'Entity Name', Value: entityName },
    { Parameter: 'Financial Year', Value: financialYear },
    { Parameter: 'Total Fixed Assets Tracked', Value: results.length },
    { Parameter: 'Total Companies Act Depreciation', Value: totalCoActDepr.toLocaleString('en-IN', { style: 'currency', currency: 'INR' }) },
    { Parameter: 'Total Income Tax Act Depreciation', Value: totalItActDepr.toLocaleString('en-IN', { style: 'currency', currency: 'INR' }) },
    { Parameter: 'Net Timing Difference (AS 22)', Value: (totalCoActDepr - totalItActDepr).toLocaleString('en-IN', { style: 'currency', currency: 'INR' }) },
    { Parameter: 'Closing Net Block (Books)', Value: totalCoActClosing.toLocaleString('en-IN', { style: 'currency', currency: 'INR' }) },
    { Parameter: 'Closing WDV (Income Tax)', Value: totalItActClosing.toLocaleString('en-IN', { style: 'currency', currency: 'INR' }) },
    { Parameter: 'Generation Date', Value: new Date().toLocaleDateString('en-IN') },
  ];

  const wsSummary = createStyledWorksheet(summaryData, {
    columnWidths: { Parameter: 35, Value: 40 },
  });
  XLSX.utils.book_append_sheet(wb, wsSummary, 'Depreciation Summary');

  return wb;
}
