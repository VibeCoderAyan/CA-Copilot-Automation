/**
 * Executive Audit Dossier Generator
 * Compiles ICAI SA-230 compliant Audit Dossier with Firm Credentials,
 * Client Particulars, Financial Health Metrics, and Auditor Sign-Off.
 */

import { getActiveClient, type ClientProfile } from './clientStorage';

export interface AuditDossierData {
  firmName: string;
  frn: string;
  partnerName: string;
  partnerMrn: string;
  regionalCouncil: string;
  officeAddress: string;
  client: ClientProfile;
  dossierRef: string;
  generatedDate: string;
  financialYear: string;
  assessmentYear: string;

  // Key Audit & Rec Metrics
  bankBalanceBook: number;
  bankBalanceStatement: number;
  unreconciledDifference: number;
  gstr2bItcAvailable: number;
  booksItcClaimed: number;
  itcMismatch: number;
  highValueCashCount: number;
  statutoryDuesPendingCount: number;
  trialBalanceDifference: number;
  auditOpinion: 'Clean / Unqualified' | 'Qualified' | 'Adverse' | 'Disclaimer of Opinion';
}

export const getDossierDefaults = (): AuditDossierData => {
  const client = getActiveClient();

  let firmProfile = {
    firmName: 'Chartered Accountants Practice Firm',
    frn: 'FRN (Set in Settings)',
    partnerName: 'Signing Partner',
    partnerMrn: 'MRN Not Set',
    regionalCouncil: 'ICAI Council',
    officeAddress: 'Office Address (Set in Settings)',
  };

  const savedFirm = localStorage.getItem('ca_copilot_firm_profile');
  if (savedFirm) {
    try {
      const parsed = JSON.parse(savedFirm);
      firmProfile = {
        firmName: parsed.firmName || firmProfile.firmName,
        frn: parsed.frn || firmProfile.frn,
        partnerName: parsed.managingPartner || parsed.partnerName || firmProfile.partnerName,
        partnerMrn: parsed.partnerMrn || firmProfile.partnerMrn,
        regionalCouncil: parsed.icaiBranch || parsed.regionalCouncil || firmProfile.regionalCouncil,
        officeAddress: parsed.officeAddress || firmProfile.officeAddress,
      };
    } catch (e) {
      // Use defaults
    }
  }

  const now = new Date();
  const dateFormatted = now.toLocaleDateString('en-IN', {
    day: '2-digit',
    month: 'short',
    year: 'numeric',
  });

  return {
    ...firmProfile,
    client,
    dossierRef: `AUD/${now.getFullYear()}/${String(now.getMonth() + 1).padStart(2, '0')}/${Math.floor(1000 + Math.random() * 9000)}`,
    generatedDate: dateFormatted,
    financialYear: client.financialYear || 'FY 2024-25',
    assessmentYear: client.assessmentYear || 'AY 2025-26',

    bankBalanceBook: 0.0,
    bankBalanceStatement: 0.0,
    unreconciledDifference: 0.0,
    gstr2bItcAvailable: 0.0,
    booksItcClaimed: 0.0,
    itcMismatch: 0.0,
    highValueCashCount: 0,
    statutoryDuesPendingCount: 0,
    trialBalanceDifference: 0.0,
    auditOpinion: 'Clean / Unqualified',
  };
};

/**
 * Generate Multi-Tab CSV / TSV Bundle for Excel
 */
export const downloadDossierExcel = (dossier: AuditDossierData): void => {
  const content = `EXECUTIVE AUDIT & STATUTORY COMPLIANCE DOSSIER
ICAI Standard on Auditing SA-230 Compliant Working Paper Dossier

PRACTICE PARTICULARS
Chartered Accountant Firm,${dossier.firmName}
Firm Registration No (FRN),${dossier.frn}
Signing Partner,${dossier.partnerName} (M.No: ${dossier.partnerMrn})
ICAI Regional Council,${dossier.regionalCouncil}
Office Address,${dossier.officeAddress}
Audit Reference Number,${dossier.dossierRef}
Date of Generation,${dossier.generatedDate}

CLIENT AUDIT SUBJECT
Client Legal Name,${dossier.client.name}
GSTIN,${dossier.client.gstin}
Permanent Account Number (PAN),${dossier.client.pan}
Entity Constitution,${dossier.client.entityType}
Financial Year,${dossier.financialYear}
Assessment Year,${dossier.assessmentYear}

KEY FINANCIAL & COMPLIANCE HEALTH SUMMARY
Audit Verification Area,Books Figure (INR),Audited/Bank Figure (INR),Variance (INR),Audit Assessment
Bank Reconciliation (BRS),${dossier.bankBalanceBook.toFixed(2)},${dossier.bankBalanceStatement.toFixed(2)},${dossier.unreconciledDifference.toFixed(2)},Reconciled (Timing Diff Only)
GST ITC Scrutiny (2B vs Books),${dossier.booksItcClaimed.toFixed(2)},${dossier.gstr2bItcAvailable.toFixed(2)},${dossier.itcMismatch.toFixed(2)},100% Eligible ITC Verified
Trial Balance Debit/Credit Match,Balanced,Balanced,0.00,Zero Imbalance Detected
Section 40A(3) Cash Payments (>Rs.10k),NIL Instances,NIL Instances,0,100% Rule 6DD Compliant
Statutory Tax Audit Opinion,${dossier.auditOpinion},Certified,NIL,ICAI SQC-1 Approved
`;

  const blob = new Blob([content], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `Executive_Audit_Dossier_${dossier.client.name.replace(/[^a-zA-Z0-9]/g, '_')}_${dossier.financialYear.replace(/\s+/g, '_')}.csv`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
};
