/**
 * Statutory Income Tax Section 206AB & Section 206CCA Non-Filer Verification Engine
 * Enforces mandatory higher TDS/TCS rates for specified non-filers under the Indian Income Tax Act 1961
 */

import * as XLSX from 'xlsx';
import { BackendStore } from '../db/store';
import { ExportService } from './exportService';
import type { ExportArtifact } from '../types';

export interface Section206abInputRecord {
  id: string;
  vendorName: string;
  pan?: string;
  gstin?: string;
  tdsSection: string; // e.g. '194C', '194J', '194Q', '194H', '194I', '192'
  paymentAmount: number;
  normalTdsRate: number; // e.g. 2 for 2%, 10 for 10%
  aggregatePriorYearTds?: number; // Aggregate TDS/TCS in relevant preceding year
  hasFiledItr?: boolean; // Verified ITR filing status from Income Tax Compliance Portal
  itrAckNumber?: string;
  isNonResidentWithoutPe?: boolean; // Proviso exemption for foreign vendors without Indian PE
}

export interface Section206abAuditRecord {
  id: string;
  vendorName: string;
  pan: string;
  gstin: string;
  panEntityType: string;
  isPanValid: boolean;
  tdsSection: string;
  paymentAmount: number;
  normalTdsRate: number;
  applicableTdsRate: number;
  normalTdsAmount: number;
  applicableTdsAmount: number;
  differentialTdsExposure: number;
  complianceStatus: 'COMPLIANT_FILER' | 'HIGH_RISK_SPECIFIED_PERSON' | 'EXEMPT_BELOW_THRESHOLD' | 'STATUTORILY_EXCLUDED_SECTION' | 'INVALID_PAN_SECTION_206AA';
  statutoryReason: string;
}

export interface Section206abSummary {
  clientName: string;
  financialYear: string;
  totalVendorsScrutinized: number;
  totalPaymentValue: number;
  highRiskSpecifiedPersonsCount: number;
  totalDifferentialTdsExposure: number;
  compliantFilersCount: number;
  exemptBelowThresholdCount: number;
  statutorilyExcludedCount: number;
  records: Section206abAuditRecord[];
}

// Sections specifically excluded from Section 206AB under subsection (1) proviso
const EXCLUDED_SECTIONS = new Set(['192', '192A', '194B', '194BA', '194BB', '194-IA', '194IA', '194-IB', '194IB', '194M', '194N']);

export class Section206abService {
  private store: BackendStore;
  private exportService: ExportService;

  constructor(store?: BackendStore) {
    this.store = store || BackendStore.getInstance();
    this.exportService = new ExportService(this.store);
  }

  /**
   * Extracts clean 10-character PAN from direct PAN field or 15-character GSTIN
   */
  public static extractPan(pan?: string, gstin?: string): { pan: string; isValid: boolean; entityType: string } {
    let clean = (pan || '').trim().toUpperCase();

    // Derive from GSTIN if PAN is missing (chars 3 to 12 in 15-digit GSTIN)
    if (!clean && gstin) {
      const cleanGstin = gstin.trim().toUpperCase();
      if (cleanGstin.length === 15) {
        clean = cleanGstin.substring(2, 12);
      }
    }

    const panRegex = /^[A-Z]{5}[0-9]{4}[A-Z]{1}$/;
    const isValid = panRegex.test(clean);

    let entityType = 'Unknown';
    if (isValid) {
      const typeChar = clean[3];
      switch (typeChar) {
        case 'C': entityType = 'Company'; break;
        case 'P': entityType = 'Individual'; break;
        case 'F': entityType = 'Partnership Firm / LLP'; break;
        case 'H': entityType = 'HUF'; break;
        case 'A': entityType = 'AOP'; break;
        case 'T': entityType = 'Trust'; break;
        case 'B': entityType = 'BOI'; break;
        case 'L': entityType = 'Local Authority'; break;
        case 'J': entityType = 'Artificial Juridical Person'; break;
        case 'G': entityType = 'Government'; break;
        default: entityType = 'Other';
      }
    }

    return { pan: clean, isValid, entityType };
  }

  /**
   * Evaluates Section 206AB & Section 206CCA statutory compliance for an array of vendor records
   */
  public evaluateRecords(
    rawRecords: Section206abInputRecord[],
    clientName: string = 'Apex Industries Ltd',
    financialYear: string = 'FY 2026-27'
  ): Section206abSummary {
    let totalPaymentValue = 0;
    let highRiskCount = 0;
    let totalDiffExposure = 0;
    let compliantCount = 0;
    let exemptThresholdCount = 0;
    let excludedCount = 0;

    const auditedRecords: Section206abAuditRecord[] = rawRecords.map((r) => {
      const payment = r.paymentAmount || 0;
      totalPaymentValue += payment;

      const { pan, isValid: isPanValid, entityType } = Section206abService.extractPan(r.pan, r.gstin);
      const cleanGstin = (r.gstin || '').trim().toUpperCase();
      const cleanSec = (r.tdsSection || '194C').trim().toUpperCase().replace(/^SEC(?:TION)?\s*/i, '');
      const normalRate = r.normalTdsRate !== undefined && r.normalTdsRate >= 0 ? r.normalTdsRate : 2.0;

      let complianceStatus: Section206abAuditRecord['complianceStatus'] = 'COMPLIANT_FILER';
      let statutoryReason = 'Vendor has filed return of income or is compliant';
      let applicableRate = normalRate;

      // 1. Invalid or Missing PAN -> Section 206AA overrides (Higher of 20% or 206AB rate)
      if (!isPanValid) {
        complianceStatus = 'INVALID_PAN_SECTION_206AA';
        applicableRate = Math.max(normalRate * 2, 5, 20);
        statutoryReason = 'Invalid/Missing PAN. Section 206AA mandates minimum 20% TDS deduction rate.';
      }
      // 2. Statutorily Excluded Section under Section 206AB(1) Proviso
      else if (EXCLUDED_SECTIONS.has(cleanSec)) {
        complianceStatus = 'STATUTORILY_EXCLUDED_SECTION';
        applicableRate = normalRate;
        statutoryReason = `Section ${cleanSec} is expressly excluded from the provisions of Section 206AB under subsection (1) proviso.`;
        excludedCount++;
      }
      // 3. Foreign Vendor without Indian Permanent Establishment
      else if (r.isNonResidentWithoutPe) {
        complianceStatus = 'STATUTORILY_EXCLUDED_SECTION';
        applicableRate = normalRate;
        statutoryReason = 'Non-resident vendor without a permanent establishment (PE) in India is exempt from Section 206AB.';
        excludedCount++;
      }
      // 4. Threshold Check: Prior Year TDS/TCS aggregate must be >= ₹50,000
      else if (r.aggregatePriorYearTds !== undefined && r.aggregatePriorYearTds < 50000) {
        complianceStatus = 'EXEMPT_BELOW_THRESHOLD';
        applicableRate = normalRate;
        statutoryReason = `Aggregate TDS/TCS in relevant preceding year (₹${r.aggregatePriorYearTds.toFixed(2)}) is below the ₹50,000 threshold. Section 206AB does not apply.`;
        exemptThresholdCount++;
      }
      // 5. Check if vendor has filed ITR
      else if (r.hasFiledItr === false) {
        // High Risk Specified Person: Higher of (2 * normalRate) or 5%
        complianceStatus = 'HIGH_RISK_SPECIFIED_PERSON';
        applicableRate = Math.max(normalRate * 2, 5.0);
        statutoryReason = `Specified Person u/s 206AB: Prior year ITR not filed and TDS threshold >= ₹50k. Higher rate applies: Max(2 * ${normalRate}%, 5%) = ${applicableRate}%.`;
        highRiskCount++;
      } else {
        complianceStatus = 'COMPLIANT_FILER';
        applicableRate = normalRate;
        statutoryReason = r.itrAckNumber
          ? `Verified compliant. ITR Acknowledgement: ${r.itrAckNumber}.`
          : 'Compliant vendor. Normal TDS rate applies.';
        compliantCount++;
      }

      const normalTds = Math.round((payment * normalRate) / 100 * 100) / 100;
      const applicableTds = Math.round((payment * applicableRate) / 100 * 100) / 100;
      const diffExposure = Math.max(0, Math.round((applicableTds - normalTds) * 100) / 100);

      totalDiffExposure += diffExposure;

      return {
        id: r.id,
        vendorName: r.vendorName,
        pan,
        gstin: cleanGstin,
        panEntityType: entityType,
        isPanValid,
        tdsSection: cleanSec,
        paymentAmount: payment,
        normalTdsRate: normalRate,
        applicableTdsRate: applicableRate,
        normalTdsAmount: normalTds,
        applicableTdsAmount: applicableTds,
        differentialTdsExposure: diffExposure,
        complianceStatus,
        statutoryReason,
      };
    });

    return {
      clientName,
      financialYear,
      totalVendorsScrutinized: auditedRecords.length,
      totalPaymentValue: Math.round(totalPaymentValue * 100) / 100,
      highRiskSpecifiedPersonsCount: highRiskCount,
      totalDifferentialTdsExposure: Math.round(totalDiffExposure * 100) / 100,
      compliantFilersCount: compliantCount,
      exemptBelowThresholdCount: exemptThresholdCount,
      statutorilyExcludedCount: excludedCount,
      records: auditedRecords,
    };
  }

  /**
   * Generates a 4-sheet Section 206AB statutory compliance audit workbook
   */
  public generateAuditWorkbook(summary: Section206abSummary): XLSX.WorkBook {
    const wb = XLSX.utils.book_new();

    // Sheet 1: Executive Summary
    const summaryRows = [
      { Parameter: 'Client Legal Entity', Value: summary.clientName },
      { Parameter: 'Statutory Financial Year', Value: summary.financialYear },
      { Parameter: 'Audit Standard', Value: 'Income Tax Act 1961 - Section 206AB & Section 206CCA' },
      { Parameter: 'Total Vendors Scrutinized', Value: summary.totalVendorsScrutinized },
      { Parameter: 'Total Payments Scrutinized (₹)', Value: summary.totalPaymentValue },
      { Parameter: 'High-Risk Specified Persons (Non-Filers)', Value: summary.highRiskSpecifiedPersonsCount },
      { Parameter: 'Potential TDS Under-Deduction Exposure (₹)', Value: summary.totalDifferentialTdsExposure },
      { Parameter: 'Verified Compliant Filers', Value: summary.compliantFilersCount },
      { Parameter: 'Exempt Below ₹50,000 Threshold', Value: summary.exemptBelowThresholdCount },
      { Parameter: 'Statutorily Excluded Sections (192/194M/etc.)', Value: summary.statutorilyExcludedCount },
      { Parameter: 'Audit Opinion / Compliance Status', Value: summary.highRiskSpecifiedPersonsCount === 0 ? 'CLEAN_COMPLIANT' : 'ACTION_REQUIRED_HIGHER_RATE_APPLICABLE' },
    ];
    const wsSummary = XLSX.utils.json_to_sheet(summaryRows);
    wsSummary['!cols'] = [{ wch: 45 }, { wch: 35 }];
    XLSX.utils.book_append_sheet(wb, wsSummary, 'Executive_Summary');

    // Sheet 2: Non-Filer Risk Schedule
    const scheduleRows = summary.records.map((r, i) => ({
      'S.No': i + 1,
      'Vendor / Payee Name': r.vendorName,
      'PAN': r.pan || 'MISSING',
      'Entity Type': r.panEntityType,
      'GSTIN': r.gstin || '-',
      'TDS Section': `Sec ${r.tdsSection}`,
      'Payment Amount (₹)': r.paymentAmount,
      'Normal Rate (%)': `${r.normalTdsRate.toFixed(2)}%`,
      'Sec 206AB Rate (%)': `${r.applicableTdsRate.toFixed(2)}%`,
      'Normal TDS (₹)': r.normalTdsAmount,
      'Mandatory TDS (₹)': r.applicableTdsAmount,
      'Differential Exposure (₹)': r.differentialTdsExposure,
      'Compliance Status': r.complianceStatus.replace(/_/g, ' '),
      'Statutory Finding & Notes': r.statutoryReason,
    }));
    const wsSchedule = XLSX.utils.json_to_sheet(scheduleRows);
    wsSchedule['!cols'] = [
      { wch: 8 },
      { wch: 32 },
      { wch: 14 },
      { wch: 20 },
      { wch: 18 },
      { wch: 14 },
      { wch: 20 },
      { wch: 15 },
      { wch: 18 },
      { wch: 16 },
      { wch: 18 },
      { wch: 22 },
      { wch: 26 },
      { wch: 45 },
    ];
    XLSX.utils.book_append_sheet(wb, wsSchedule, 'Section_206AB_Audit');

    // Sheet 3: Excluded Sections Reference
    const refRows = [
      { 'Section': '192', 'Description': 'Payment of Salary', 'Exclusion Rationale': 'Salaries subject to slab rates; excluded by Sec 206AB(1) proviso.' },
      { 'Section': '192A', 'Description': 'Payment of accumulated balance due to an employee in EPF', 'Exclusion Rationale': 'Retirement benefits excluded by Sec 206AB(1) proviso.' },
      { 'Section': '194B', 'Description': 'Winnings from lottery or crossword puzzle', 'Exclusion Rationale': 'Flat 30% already applies; excluded by Sec 206AB(1) proviso.' },
      { 'Section': '194BA', 'Description': 'Winnings from online games', 'Exclusion Rationale': 'Net winnings provisions apply; excluded by Sec 206AB(1) proviso.' },
      { 'Section': '194BB', 'Description': 'Winnings from horse race', 'Exclusion Rationale': 'Excluded by Sec 206AB(1) proviso.' },
      { 'Section': '194-IA', 'Description': 'Payment on transfer of certain immovable property other than agricultural land', 'Exclusion Rationale': 'Individual purchaser compliance threshold; excluded by proviso.' },
      { 'Section': '194-IB', 'Description': 'Payment of rent by certain individuals or HUF', 'Exclusion Rationale': 'Individual tenant relief; excluded by Sec 206AB(1) proviso.' },
      { 'Section': '194M', 'Description': 'Payment of certain sums by certain individuals or HUF', 'Exclusion Rationale': 'Excluded by Sec 206AB(1) proviso.' },
      { 'Section': '194N', 'Description': 'Payment of certain amounts in cash by banking companies', 'Exclusion Rationale': 'Cash withdrawal deterrence; separate tiered structure.' },
    ];
    const wsRef = XLSX.utils.json_to_sheet(refRows);
    wsRef['!cols'] = [{ wch: 14 }, { wch: 45 }, { wch: 45 }];
    XLSX.utils.book_append_sheet(wb, wsRef, 'Statutory_Exclusions');

    // Sheet 4: Statutory Vendor Declaration Template
    const templateRows = [
      { Field: 'DECLARATION UNDER SECTION 206AB & 206CCA OF THE INCOME TAX ACT, 1961', Details: '' },
      { Field: 'To', Details: summary.clientName },
      { Field: 'Subject', Details: `Confirmation of ITR filing status for relevant Assessment Years preceding ${summary.financialYear}` },
      { Field: 'Vendor Legal Name', Details: '[Insert Vendor Legal Name]' },
      { Field: 'Permanent Account Number (PAN)', Details: '[Insert 10-Digit Valid PAN]' },
      { Field: 'GSTIN', Details: '[Insert 15-Digit GSTIN]' },
      { Field: 'ITR Filing Status for Preceding Year', Details: 'Filed / Not Applicable' },
      { Field: 'ITR Acknowledgement Number', Details: '[Insert 15-digit e-filing Acknowledgement Number]' },
      { Field: 'Aggregate TDS/TCS in Preceding FY Exceeded ₹50,000?', Details: 'Yes / No' },
      { Field: 'Permanent Establishment in India (for Foreign Vendor)', Details: 'Yes / No' },
      { Field: 'Authorized Signatory Name & Designation', Details: '[Name, Designation, Company Stamp]' },
      { Field: 'Date & Place', Details: '[DD/MM/YYYY, City]' },
    ];
    const wsTemplate = XLSX.utils.json_to_sheet(templateRows);
    wsTemplate['!cols'] = [{ wch: 45 }, { wch: 45 }];
    XLSX.utils.book_append_sheet(wb, wsTemplate, 'Vendor_Declaration_Template');

    return wb;
  }

  /**
   * Generates and stores ExportArtifact in store
   */
  public exportAuditArtifact(
    summary: Section206abSummary,
    firmId: string = 'firm_apex_ca',
    clientId: string = 'client_apex_technologies'
  ): ExportArtifact {
    const wb = this.generateAuditWorkbook(summary);
    const buffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
    const base64 = buffer.toString('base64');
    const artifactId = `art-sec206ab-${Date.now()}`;
    const fileName = `Section_206AB_NonFiler_Audit_${summary.financialYear.replace(/\s+/g, '_')}_${Date.now()}.xlsx`;

    const artifact: ExportArtifact = {
      id: artifactId,
      firmId,
      clientId,
      workflowRunId: 'sec206ab-audit',
      fileName,
      fileType: 'xlsx',
      fileSizeBytes: buffer.length,
      downloadUrl: `/api/artifacts/${artifactId}/download`,
      generatedAt: new Date().toISOString(),
      contentBase64: base64,
      base64Payload: base64,
      sheetNames: wb.SheetNames,
      totalRecords: summary.totalVendorsScrutinized,
      integrityVerified: true,
    };

    this.store.saveArtifact(artifact);
    return artifact;
  }
}
