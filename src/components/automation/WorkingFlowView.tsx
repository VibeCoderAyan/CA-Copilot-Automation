import React from 'react';
import {
  ShieldCheck,
  FileSpreadsheet,
  Database,
  Scale,
  CheckCheck,
  FileCode,
  ArrowRight,
  Shield,
  Layers,
  Lock,
  Cpu,
  ReceiptText,
  FileText,
  FileCog,
  Calculator,
} from 'lucide-react';

interface WorkingFlowViewProps {
  onSelectTool?: (toolId: string) => void;
  onShowToast?: (title: string, desc?: string, type?: 'success' | 'error' | 'info') => void;
  selectedClientId?: string;
}

export const WorkingFlowView: React.FC<WorkingFlowViewProps> = ({ onSelectTool, onShowToast }) => {
  const handleToolClick = (toolId: string, _toolName?: string) => {
    if (onSelectTool) {
      onSelectTool(toolId);
    }
  };

  return (
    <div className="space-y-6 pb-16 font-sans select-none text-slate-900 animate-apple-fade">
      {/* ========================================================
          EXECUTIVE WORKFLOW CANVAS (ORIGINAL 3-COLUMN POSTER LAYOUT)
         ======================================================== */}
      <div className="max-w-[1250px] mx-auto bg-white rounded-xl border border-slate-200 p-6 sm:p-10 shadow-xs print:p-0 print:border-none print:shadow-none print:max-w-none print:w-full">
        {/* ----------------------------------------------------
            POSTER HEADER
           ---------------------------------------------------- */}
        <header className="text-center pb-8 mb-8 border-b border-slate-200">
          <div className="inline-flex items-center gap-2 px-3 py-1 mb-3 text-[11px] font-bold tracking-wider uppercase text-slate-700 bg-slate-100 rounded-md border border-slate-200">
            <Layers className="w-3.5 h-3.5 text-slate-800" />
            <span>Statutory Audit & Assurance Architecture</span>
          </div>
          <h1 className="text-2xl sm:text-3xl lg:text-4xl font-extrabold tracking-tight text-slate-900 uppercase">
            End-to-End Accounting & Audit Pipeline
          </h1>
          <p className="text-xs sm:text-sm text-slate-600 max-w-3xl mx-auto font-normal mt-2 leading-relaxed">
            Deterministic Financial Verification · Multi-Stream Reconciliations · Form 3CD & GST 3-Way Scrutiny · ICAI SA-230/505 Working Papers · Direct Tally Prime XML Export
          </p>
        </header>

        {/* ----------------------------------------------------
            UPPER SECTION: 3-COLUMN STRUCTURE
            Left: Pipeline Stages (Sequence)
            Center: Multi-Stream Intake & Verification Pipeline
            Right: Statutory Compliance Framework
           ---------------------------------------------------- */}
        <div className="grid grid-cols-1 lg:grid-cols-12 gap-6 lg:gap-5 mb-10">
          {/* ==================================================
              LEFT COLUMN: PIPELINE STAGES (Width: 3 cols)
             ================================================== */}
          <div className="lg:col-span-3 flex flex-col space-y-3">
            <div className="p-3 bg-slate-50 border border-slate-200 rounded-lg text-center">
              <span className="text-[10px] font-bold uppercase tracking-wider text-slate-500 block">
                Workflow Sequence
              </span>
              <span className="text-xs font-bold text-slate-900 uppercase">
                Execution Pipeline
              </span>
            </div>

            <div className="space-y-2">
              {[
                {
                  code: '1. INGESTION',
                  desc: 'Structured extraction of bank statements (PDF/CSV), accounting daybooks, and GST portal data.',
                  toolId: 'bank-converter',
                  toolName: 'Bank Statement Converter',
                },
                {
                  code: '2. STANDARDIZE',
                  desc: 'Schema normalization, floating-point numeric parsing, date format sanitization, and debit/credit balance verification.',
                  toolId: 'excel-data-cleanup',
                  toolName: 'Excel Data Cleanup',
                },
                {
                  code: '3. SCRUTINY',
                  desc: 'Automated exception checks under Income Tax Sec 40A(3), Rule 6DD, Form 3CD (Clauses 21 & 44), and Sec 206AB TDS.',
                  toolId: 'form3cd-scrutiny',
                  toolName: 'Form 3CD Scrutiny',
                },
                {
                  code: '4. RECONCILE',
                  desc: 'Deterministic multi-stream matching: Bank BRS, GSTR-2B vs Books, and GSTR-1 vs 3B vs Ledger (DRC-01B).',
                  toolId: 'gst-triple-rec',
                  toolName: 'GST 3-Way Rec',
                },
                {
                  code: '5. CERTIFY & EXPORT',
                  desc: 'ICAI SA-230 lead schedules, Partner audit sign-off dossier, and balanced Tally Prime XML vouchers.',
                  toolId: 'tally-xml',
                  toolName: 'Tally XML Generator',
                },
              ].map((item, idx) => (
                <div
                  key={idx}
                  onClick={() => handleToolClick(item.toolId, item.toolName)}
                  className="group p-3 rounded-lg border border-slate-200 bg-white hover:border-slate-400 hover:shadow-2xs transition-all cursor-pointer"
                >
                  <div className="flex items-center justify-between">
                    <span className="text-[11px] font-bold text-slate-900 tracking-wide">
                      {item.code}
                    </span>
                    <ArrowRight className="w-3 h-3 text-slate-400 group-hover:text-slate-900 group-hover:translate-x-0.5 transition-all" />
                  </div>
                  <div className="text-[10px] text-slate-600 mt-1 leading-snug">
                    {item.desc}
                  </div>
                </div>
              ))}
            </div>
          </div>

          {/* ==================================================
              CENTER COLUMN: METHODOLOGY & INTAKE (Width: 6 cols)
             ================================================== */}
          <div className="lg:col-span-6 flex flex-col space-y-4">
            {/* Top Center Header (Clean Dark Slate - Zero AI Gradient) */}
            <div className="p-3.5 bg-slate-900 text-white rounded-lg text-center shadow-xs">
              <div className="text-[10px] font-semibold uppercase tracking-widest text-slate-300">
                Core Execution Architecture
              </div>
              <div className="text-sm font-bold uppercase tracking-tight mt-0.5">
                Primary Financial Ingestion & Verification Gateway
              </div>
            </div>

            {/* Ingestion Channels Strip */}
            <div className="p-4 bg-slate-50 border border-slate-200 rounded-lg space-y-2">
              <div className="text-[10px] font-bold text-center tracking-wider text-slate-500 uppercase">
                Supported Financial Data Sources
              </div>
              <div className="grid grid-cols-5 gap-2">
                {[
                  { name: 'Bank PDF', sub: 'Digital Statement', icon: Database, toolId: 'bank-converter', toolName: 'Bank Statement Converter' },
                  { name: 'Ledgers', sub: 'Tally / SAP / ERP', icon: FileSpreadsheet, toolId: 'ledger-compare', toolName: 'Ledger Comparison' },
                  { name: 'GSTR-2B', sub: 'Portal JSON / Excel', icon: FileCode, toolId: 'gst-rec', toolName: 'GSTR-2B Reconciliation' },
                  { name: 'Form 26AS', sub: 'TRACES Statement', icon: Scale, toolId: 'tds-reconciliation', toolName: 'TDS Reconciliation' },
                  { name: 'Invoices', sub: 'B2B Sales & Purchases', icon: CheckCheck, toolId: 'invoice-extractor', toolName: 'Invoice Extractor' },
                ].map((doc, i) => (
                  <div
                    key={i}
                    onClick={() => handleToolClick(doc.toolId, doc.toolName)}
                    className="p-2 rounded-md border border-slate-200 bg-white text-center flex flex-col items-center justify-center hover:border-slate-400 hover:shadow-2xs cursor-pointer transition-all"
                  >
                    {React.createElement(doc.icon, { className: 'w-4 h-4 text-slate-700 mb-1' })}
                    <span className="text-[10px] font-bold text-slate-900 leading-tight">{doc.name}</span>
                    <span className="text-[8px] text-slate-500 mt-0.5">{doc.sub}</span>
                  </div>
                ))}
              </div>
            </div>

            {/* 3 Core Processing Pillars */}
            <div className="grid grid-cols-3 gap-3">
              {[
                {
                  title: 'INGESTION GATEWAY',
                  sub: 'Phase I',
                  detail: 'Digital PDF statement extraction, bank schema mapping, and client ledger normalization.',
                },
                {
                  title: 'STATUTORY SCRUTINY',
                  sub: 'Phase II',
                  detail: 'Form 3CD Clause 21(d) cash limits, Sec 206AB higher TDS rates, and Rule 88C return variance detection.',
                },
                {
                  title: 'ASSURANCE & EXPORT',
                  sub: 'Phase III',
                  detail: 'Tally Prime XML vouchers, ICAI SA-230 lead schedules, and executive partner summary.',
                },
              ].map((pod, i) => (
                <div
                  key={i}
                  className="p-3.5 rounded-lg border border-slate-200 bg-white text-center flex flex-col justify-between shadow-2xs"
                >
                  <div>
                    <div className="text-[9px] font-bold text-slate-500 uppercase tracking-wider">
                      {pod.sub}
                    </div>
                    <div className="text-[11px] font-bold text-slate-900 uppercase mt-0.5">
                      {pod.title}
                    </div>
                  </div>
                  <div className="text-[9px] text-slate-600 mt-2 leading-relaxed">
                    {pod.detail}
                  </div>
                </div>
              ))}
            </div>

            {/* Architectural Rule Callout */}
            <div className="p-3 bg-slate-50 border border-slate-200 rounded-lg text-center text-[11px] text-slate-600 leading-relaxed font-mono">
              Deterministic Verification Rule: Zero mathematical hallucination or probabilistic inference. Exact IEEE 754 financial arithmetic with ICAI SA-230 partner sign-off.
            </div>
          </div>

          {/* ==================================================
              RIGHT COLUMN: STATUTORY FRAMEWORK (Width: 3 cols)
             ================================================== */}
          <div className="lg:col-span-3 flex flex-col space-y-3">
            <div className="p-3 bg-slate-50 border border-slate-200 rounded-lg text-center">
              <span className="text-[10px] font-bold uppercase tracking-wider text-slate-500 block">
                Compliance Standard
              </span>
              <span className="text-xs font-bold text-slate-900 uppercase">
                Statutory Framework
              </span>
            </div>

            <div className="space-y-2">
              {[
                {
                  code: 'INCOME TAX ACT, 1961',
                  desc: 'Section 40A(3) aggregate cash payment thresholds, Rule 6DD exclusions, Section 43B statutory dues, and Form 3CD disclosures.',
                },
                {
                  code: 'CGST ACT, 2017',
                  desc: 'Section 16(2)(aa) input tax credit matching, Rule 88C (DRC-01B) return variances, and Section 50 mandatory interest.',
                },
                {
                  code: 'TDS & WITHHOLDING',
                  desc: 'Section 206AB / 206CCA specified person higher rates (5% floor rule), Section 194Q, and Section 201(1A) interest.',
                },
                {
                  code: 'ICAI AUDITING STANDARDS',
                  desc: 'SA 230 Audit Documentation, SA 500 Audit Evidence, SA 505 External Confirmations, and SA 530 Audit Sampling.',
                },
                {
                  code: 'ICAI SQC 1 QUALITY CONTROL',
                  desc: 'System of Quality Control 1: engagement documentation custody, cryptographic verification hashes, and partner sign-off logs.',
                },
              ].map((item, idx) => (
                <div
                  key={idx}
                  className="p-3 rounded-lg border border-slate-200 bg-white hover:border-slate-300 transition-colors"
                >
                  <div className="text-[11px] font-bold text-slate-900 tracking-wide">
                    {item.code}
                  </div>
                  <div className="text-[10px] text-slate-600 mt-1 leading-snug">
                    {item.desc}
                  </div>
                </div>
              ))}
            </div>
          </div>
        </div>

        {/* ----------------------------------------------------
            MIDDLE SECTION: DETAILED WORKFLOW MODULES
            Left: Audit Scrutiny Modules
            Center: The CA Audit Cycle
            Right: Statutory Deliverables
           ---------------------------------------------------- */}
        <div className="pt-6 pb-8 border-t border-b border-slate-200 mb-8">
          <div className="grid grid-cols-1 lg:grid-cols-12 gap-6 items-center">
            {/* Left Module: Scrutiny & Reconciliation (Clickable Tool Nodes) */}
            <div className="lg:col-span-4 flex flex-col">
              <div className="p-3 bg-slate-100 rounded-lg border border-slate-200 mb-3 text-center">
                <div className="text-xs font-bold uppercase text-slate-900 tracking-wider">
                  Audit Scrutiny Modules
                </div>
                <div className="text-[10px] text-slate-500 mt-0.5">
                  Automated verification executed across financial ledgers and statutory registers
                </div>
              </div>

              <div className="grid grid-cols-2 gap-2">
                {[
                  { title: 'Sec 40A(3) Cash', sub: '₹10k Daily Aggregate Limit', toolId: 'form3cd-scrutiny', toolName: 'Form 3CD Scrutiny' },
                  { title: 'Form 3CD Clause 44', sub: 'Registered vs Unregistered', toolId: 'form3cd-scrutiny', toolName: 'Form 3CD Scrutiny' },
                  { title: 'Bank Rec (BRS)', sub: 'Running Balance Continuity', toolId: 'bank-rec', toolName: 'Bank Reconciliation' },
                  { title: 'GST 3-Way Rec', sub: '1 vs 3B vs Books (DRC-01B)', toolId: 'gst-triple-rec', toolName: 'GST 3-Way Rec' },
                  { title: 'Sec 206AB TDS', sub: 'Non-filer Higher Deduction Rates', toolId: 'tds-reconciliation', toolName: 'TDS Reconciliation' },
                  { title: 'Trial Balance Scrutiny', sub: 'Debit/Credit Imbalance Checks', toolId: 'tb-checker', toolName: 'Trial Balance Checker' },
                ].map((b, i) => (
                  <div
                    key={i}
                    onClick={() => handleToolClick(b.toolId, b.toolName)}
                    className="p-2.5 rounded-lg border border-slate-200 bg-slate-50/60 hover:bg-white hover:border-slate-400 hover:shadow-2xs text-center cursor-pointer transition-all group"
                  >
                    <div className="text-[10px] font-bold text-slate-900 leading-tight group-hover:text-slate-900 flex items-center justify-center gap-1">
                      <span>{b.title}</span>
                    </div>
                    <div className="text-[8px] text-slate-500 mt-0.5">{b.sub}</div>
                  </div>
                ))}
              </div>
            </div>

            {/* Central Stage: The CA Audit Cycle */}
            <div className="lg:col-span-4 flex flex-col items-center justify-center">
              <div className="w-full max-w-[290px] p-5 rounded-xl border border-slate-300 bg-slate-50 text-center shadow-xs">
                <div className="w-9 h-9 rounded-xl bg-slate-900 text-white flex items-center justify-center mx-auto mb-2 shadow-xs">
                  <ShieldCheck className="w-5 h-5 text-white" />
                </div>
                <div className="text-[10px] font-bold uppercase tracking-widest text-slate-500">
                  Standard Audit Procedure
                </div>
                <div className="text-sm font-extrabold uppercase text-slate-900 tracking-tight mt-0.5">
                  Continuous Audit Engagement Cycle
                </div>
                <div className="mt-3 space-y-1.5 text-left border-t border-slate-200 pt-3">
                  {[
                    '1. Ingest Source Documents & Validate File Hashes',
                    '2. Standardize Schema, Numerical Fields & Dates',
                    '3. Execute Statutory Scrutiny & Tax Audit Rules',
                    '4. Reconcile Cross-Register Balances & Variances',
                    '5. Compile External Confirmation (SA-505) Discrepancies',
                    '6. Lock SA-230 Working Papers & Export Deliverables',
                  ].map((step, idx) => (
                    <div key={idx} className="flex items-center gap-1.5 text-[10px] font-medium text-slate-700">
                      <span className="w-1.5 h-1.5 rounded-full bg-slate-800 shrink-0"></span>
                      <span>{step}</span>
                    </div>
                  ))}
                </div>
              </div>
            </div>

            {/* Right Module: Certified Deliverables (Clickable Tool Nodes) */}
            <div className="lg:col-span-4 flex flex-col">
              <div className="p-3 bg-slate-100 rounded-lg border border-slate-200 mb-3 text-center">
                <div className="text-xs font-bold uppercase text-slate-900 tracking-wider">
                  Statutory Deliverables
                </div>
                <div className="text-[10px] text-slate-500 mt-0.5">
                  Standardized audit working papers, statutory extracts, and ERP import files
                </div>
              </div>

              <div className="grid grid-cols-2 gap-2">
                {[
                  { title: 'Tally Prime XML', sub: 'Direct Import Vouchers', toolId: 'tally-xml', toolName: 'Tally XML Generator' },
                  { title: 'Form 3CD Packs', sub: 'Clauses 21a, 21d, & 44', toolId: 'form3cd-scrutiny', toolName: 'Form 3CD Scrutiny' },
                  { title: 'Executive Dossier', sub: 'Partner Briefing Pack', toolId: 'audit-dossier', toolName: 'Executive Audit Dossier' },
                  { title: 'ICAI SA-230 Leads', sub: 'Lead Schedules (WP-A to I)', toolId: 'audit-workpapers', toolName: 'Audit Working Papers' },
                  { title: 'DRC-01B Notice Pack', sub: 'Tax Shortfall Defense Table', toolId: 'gst-triple-rec', toolName: 'GST 3-Way Rec' },
                  { title: 'Schedule III BS', sub: 'Assets == Liab + Equity', toolId: 'financial-reports', toolName: 'Financial Reports' },
                ].map((b, i) => (
                  <div
                    key={i}
                    onClick={() => handleToolClick(b.toolId, b.toolName)}
                    className="p-2.5 rounded-lg border border-slate-200 bg-slate-50/60 hover:bg-white hover:border-slate-400 hover:shadow-2xs text-center cursor-pointer transition-all group"
                  >
                    <div className="text-[10px] font-bold text-slate-900 leading-tight group-hover:text-slate-900 flex items-center justify-center gap-1">
                      <span>{b.title}</span>
                    </div>
                    <div className="text-[8px] text-slate-500 mt-0.5">{b.sub}</div>
                  </div>
                ))}
              </div>
            </div>
          </div>
        </div>

        {/* ----------------------------------------------------
            BOTTOM SECTION: CORE QUALITY GUARANTEES
           ---------------------------------------------------- */}
        <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
          <div className="p-4 rounded-lg border border-slate-200 bg-slate-50 flex items-start gap-3">
            <Cpu className="w-5 h-5 text-slate-800 shrink-0 mt-0.5" />
            <div>
              <div className="text-xs font-bold uppercase text-slate-900 tracking-wide">
                Deterministic Precision
              </div>
              <p className="text-[11px] text-slate-600 mt-0.5 leading-relaxed">
                Calculations execute via exact IEEE 754 financial arithmetic with zero probabilistic inference. Every variance calculation is verifiable and reproducible.
              </p>
            </div>
          </div>

          <div className="p-4 rounded-lg border border-slate-200 bg-slate-50 flex items-start gap-3">
            <Lock className="w-5 h-5 text-slate-800 shrink-0 mt-0.5" />
            <div>
              <div className="text-xs font-bold uppercase text-slate-900 tracking-wide">
                ICAI SA-230 Custody
              </div>
              <p className="text-[11px] text-slate-600 mt-0.5 leading-relaxed">
                Audit working papers are locked with Managing Partner credentials to maintain tamper-evident records compliant with ICAI SQC 1 standards.
              </p>
            </div>
          </div>

          <div className="p-4 rounded-lg border border-slate-200 bg-slate-50 flex items-start gap-3">
            <ShieldCheck className="w-5 h-5 text-slate-800 shrink-0 mt-0.5" />
            <div>
              <div className="text-xs font-bold uppercase text-slate-900 tracking-wide">
                100% Offline Data Security
              </div>
              <p className="text-[11px] text-slate-600 mt-0.5 leading-relaxed">
                Complete client-side local execution. Confidential client ledgers, bank statements, and tax computations never leave your local system.
              </p>
            </div>
          </div>
        </div>

        {/* Footer Note */}
        <footer className="text-center pt-6 mt-6 border-t border-slate-200 text-[10px] text-slate-400 font-mono tracking-wider uppercase">
          Chartered Accountancy Audit & Assurance Architecture · FY 2024-25 to 2030-31
        </footer>
      </div>
    </div>
  );
};
