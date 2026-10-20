import React from 'react';

interface ToolIconProps {
  toolId: string;
  size?: 'sm' | 'md' | 'lg';
  className?: string;
}

/**
 * Premium Dribbble/Figma-grade App Icons for CA Automation Tools.
 * Each tool features a bespoke category-aligned gradient container,
 * zero dark/black muddy stops, and high-precision crisp white vector iconography.
 */

const TOOL_GRADIENTS: Record<string, string> = {
  // Reconciliations (Ocean Blues, Teals & Cyans)
  'bank-rec': 'from-sky-500 via-blue-600 to-indigo-600',
  'bank-converter': 'from-cyan-500 via-blue-600 to-blue-700',
  'ledger-compare': 'from-blue-600 via-indigo-600 to-violet-600',
  'ledger-rec': 'from-blue-600 via-indigo-600 to-violet-600',
  'tb-checker': 'from-teal-500 via-cyan-600 to-blue-600',
  'trial-balance': 'from-teal-500 via-cyan-600 to-blue-600',
  'multi-client-batch': 'from-indigo-500 via-blue-600 to-purple-600',
  'receivables-aging': 'from-sky-500 via-blue-600 to-indigo-700',
  'purchase-invoice-matching': 'from-blue-500 via-indigo-600 to-purple-600',
  'purchase-matching': 'from-blue-500 via-indigo-600 to-purple-600',
  'voucher-verification': 'from-blue-500 via-cyan-600 to-teal-600',
  'workflow-orchestrator': 'from-indigo-500 via-purple-600 to-blue-600',

  // Tax & GST (Violets, Purples, Emeralds & Ambers)
  'gst-rec': 'from-purple-500 via-violet-600 to-indigo-600',
  'gst-triple-rec': 'from-violet-600 via-purple-600 to-indigo-700',
  'triple-rec': 'from-violet-600 via-purple-600 to-indigo-700',
  'tds-reconciliation': 'from-amber-500 via-orange-500 to-amber-600',
  'tds-rec': 'from-amber-500 via-orange-500 to-amber-600',
  'form3cd-scrutiny': 'from-amber-500 via-orange-600 to-rose-600',
  'form3cd': 'from-amber-500 via-orange-600 to-rose-600',
  'invoice-extractor': 'from-violet-500 via-purple-600 to-indigo-600',
  'invoice-data': 'from-violet-500 via-purple-600 to-indigo-600',
  'duplicate-invoice': 'from-rose-500 via-red-500 to-orange-500',
  reports: 'from-emerald-600 via-teal-600 to-cyan-700',
  'financial-reports': 'from-emerald-600 via-teal-600 to-cyan-700',
  'recurring-compliance': 'from-rose-500 via-pink-600 to-purple-600',
  invoices: 'from-amber-500 via-orange-600 to-amber-700',

  // Audit & Scrutiny (Deep Royals, Indigos, Emeralds & Warm Ambers)
  'month-end-closing': 'from-indigo-500 via-blue-600 to-purple-600',
  'audit-checklist': 'from-blue-600 via-indigo-600 to-blue-700',
  audit: 'from-indigo-600 via-violet-600 to-purple-700',
  'audit-tools': 'from-indigo-600 via-violet-600 to-purple-700',
  'audit-sampling': 'from-blue-600 via-indigo-600 to-purple-700',
  depreciation: 'from-teal-500 via-emerald-600 to-green-600',
  'far-checker': 'from-amber-500 via-yellow-500 to-orange-600',
  'balance-confirmation': 'from-cyan-500 via-blue-600 to-indigo-600',
  'journal-validator': 'from-emerald-500 via-teal-600 to-cyan-700',
  'audit-workpapers': 'from-blue-600 via-indigo-600 to-violet-700',
  'audit-workpaper': 'from-blue-600 via-indigo-600 to-violet-700',
  'ratio-analyzer': 'from-fuchsia-500 via-purple-600 to-indigo-700',
  'cashflow-generator': 'from-emerald-400 via-teal-500 to-cyan-600',
  'cash-flow': 'from-emerald-400 via-teal-500 to-cyan-600',
  'payroll-calculator': 'from-indigo-500 via-purple-600 to-blue-600',
  payroll: 'from-indigo-500 via-purple-600 to-blue-600',
  'client-documents': 'from-sky-500 via-blue-600 to-indigo-600',

  // Excel Utilities (Rich Emeralds, Cyans, Rose & Purples)
  excel: 'from-emerald-500 via-teal-600 to-green-700',
  'file-compare': 'from-blue-500 via-indigo-600 to-purple-600',
  'pdf-excel': 'from-rose-500 via-pink-600 to-purple-600',
  pdf: 'from-rose-500 via-pink-600 to-purple-600',
  'pdf-to-excel': 'from-rose-500 via-pink-600 to-purple-600',
  'formula-checker': 'from-rose-500 via-pink-600 to-red-600',
  'formula-auditor': 'from-rose-500 via-pink-600 to-red-600',
  'excel-data-cleanup': 'from-cyan-500 via-teal-500 to-emerald-600',
  'excel-validation': 'from-cyan-500 via-teal-500 to-emerald-600',
  'excel-validator': 'from-cyan-500 via-teal-500 to-emerald-600',
  'excel-templates': 'from-blue-500 via-cyan-600 to-indigo-600',
  'bulk-renamer': 'from-indigo-500 via-purple-600 to-blue-600',
  'tally-xml': 'from-slate-700 via-slate-800 to-slate-900',
  'audit-dossier': 'from-gray-800 via-slate-900 to-black',
  'client-vault': 'from-slate-700 via-gray-800 to-slate-900',
  'offline-engines': 'from-slate-700 via-gray-800 to-slate-900',
};

const DEFAULT_GRADIENT = 'from-blue-500 via-indigo-600 to-purple-600';

export const ToolIcon: React.FC<ToolIconProps> = ({
  toolId,
  size = 'md',
  className = '',
}) => {
  const containerSize =
    size === 'sm'
      ? 'w-8 h-8 rounded-lg'
      : size === 'lg'
      ? 'w-14 h-14 rounded-2xl'
      : 'w-11 h-11 rounded-xl';

  const gradient = TOOL_GRADIENTS[toolId] || DEFAULT_GRADIENT;
  const iconSize =
    size === 'sm' ? 'w-4 h-4' : size === 'lg' ? 'w-7 h-7' : 'w-5 h-5';

  switch (toolId) {
    case 'excel':
      // Excel Cleaner & Styler (Spreadsheet Emerald)
      return (
        <div
          className={`${containerSize} bg-gradient-to-br ${gradient} p-0.5 shadow-sm ring-1 ring-black/5 flex items-center justify-center text-white shrink-0 ${className}`}
        >
          <svg className={iconSize} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
            <rect width="18" height="18" x="3" y="3" rx="2" />
            <path d="M3 9h18" />
            <path d="M9 21V9" />
            <path d="m13.5 13 5 5" strokeWidth="2.2" />
            <path d="m18.5 13-5 5" strokeWidth="2.2" />
          </svg>
        </div>
      );

    case 'file-compare':
      // File Comparison & Diff (Corporate Blue / Indigo Diff)
      return (
        <div
          className={`${containerSize} bg-gradient-to-br ${gradient} p-0.5 shadow-sm ring-1 ring-black/5 flex items-center justify-center text-white shrink-0 ${className}`}
        >
          <svg className={iconSize} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
            <circle cx="18" cy="18" r="3" fill="currentColor" fillOpacity="0.2" />
            <circle cx="6" cy="6" r="3" fill="currentColor" fillOpacity="0.2" />
            <path d="M6 9v12" />
            <path d="M13 6h3a2 2 0 0 1 2 2v7" />
            <line x1="10" y1="12" x2="16" y2="12" strokeWidth="2" />
          </svg>
        </div>
      );

    case 'bank-converter':
      // Bank Statement Converter (Ocean Blue Vault)
      return (
        <div
          className={`${containerSize} bg-gradient-to-br ${gradient} p-0.5 shadow-sm ring-1 ring-black/5 flex items-center justify-center text-white shrink-0 ${className}`}
        >
          <svg className={iconSize} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
            <path d="M3 21h18" />
            <path d="M5 21V10" />
            <path d="M19 21V10" />
            <path d="M10 21V10" />
            <path d="M14 21V10" />
            <polygon points="12 3 2 8 22 8 12 3" fill="currentColor" fillOpacity="0.25" />
          </svg>
        </div>
      );

    case 'formula-checker':
    case 'formula-auditor':
      // Formula & Error Auditor (Rose / Crimson Precision)
      return (
        <div
          className={`${containerSize} bg-gradient-to-br ${gradient} p-0.5 shadow-sm ring-1 ring-black/5 flex items-center justify-center text-white shrink-0 ${className}`}
        >
          <svg className={iconSize} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
            <rect width="16" height="20" x="4" y="2" rx="2" fill="currentColor" fillOpacity="0.15" />
            <path d="M8 8h3l2 8h3" strokeWidth="2" />
            <path d="M9 12h6" strokeWidth="2" />
            <circle cx="16" cy="17" r="1.5" fill="currentColor" />
          </svg>
        </div>
      );

    case 'ledger-compare':
    case 'ledger-rec':
      // Ledger Comparison & Rec (Royal Blue / Violet Scale)
      return (
        <div
          className={`${containerSize} bg-gradient-to-br ${gradient} p-0.5 shadow-sm ring-1 ring-black/5 flex items-center justify-center text-white shrink-0 ${className}`}
        >
          <svg className={iconSize} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
            <path d="m16 16 3-8 3 8c-.87.65-1.92 1-3 1s-2.13-.35-3-1Z" fill="currentColor" fillOpacity="0.25" />
            <path d="m2 16 3-8 3 8c-.87.65-1.92 1-3 1s-2.13-.35-3-1Z" fill="currentColor" fillOpacity="0.25" />
            <path d="M7 21h10" />
            <path d="M12 3v18" />
            <path d="M3 7h2c2 0 5-1 7-2 2 1 5 2 7 2h2" />
          </svg>
        </div>
      );

    case 'invoice-extractor':
    case 'invoice-data':
      // Invoice Data Extractor (Purple / Fuchsia Scanner)
      return (
        <div
          className={`${containerSize} bg-gradient-to-br ${gradient} p-0.5 shadow-sm ring-1 ring-black/5 flex items-center justify-center text-white shrink-0 ${className}`}
        >
          <svg className={iconSize} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
            <path d="M4 2v20l2-1 2 1 2-1 2 1 2-1 2 1 2-1 2 1 2-1 2 1V2l-2 1-2-1-2 1-2-1-2 1-2-1-2 1-2-1-2 1Z" fill="currentColor" fillOpacity="0.15" />
            <path d="M16 8h-6a2 2 0 1 0 0 4h4a2 2 0 1 1 0 4H8" />
            <path d="M12 17V7" />
          </svg>
        </div>
      );

    case 'bulk-renamer':
      // Bulk Document Renamer (Indigo / Purple Files)
      return (
        <div
          className={`${containerSize} bg-gradient-to-br ${gradient} p-0.5 shadow-sm ring-1 ring-black/5 flex items-center justify-center text-white shrink-0 ${className}`}
        >
          <svg className={iconSize} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
            <path d="M20 20a2 2 0 0 0 2-2V8a2 2 0 0 0-2-2h-7.9a2 2 0 0 1-1.69-.9L9.6 3.9A2 2 0 0 0 7.93 3H4a2 2 0 0 0-2 2v13a2 2 0 0 0 2 2Z" fill="currentColor" fillOpacity="0.2" />
            <path d="m13 13 3 3" strokeWidth="2.2" />
            <path d="m16 13-3 3" strokeWidth="2.2" />
            <path d="M8 14h.01" strokeWidth="2.5" />
          </svg>
        </div>
      );

    case 'audit-checklist':
      // Audit Checklist Manager (Royal Blue Seal)
      return (
        <div
          className={`${containerSize} bg-gradient-to-br ${gradient} p-0.5 shadow-sm ring-1 ring-black/5 flex items-center justify-center text-white shrink-0 ${className}`}
        >
          <svg className={iconSize} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
            <rect width="8" height="4" x="8" y="2" rx="1" ry="1" fill="currentColor" fillOpacity="0.25" />
            <path d="M16 4h2a2 2 0 0 1 2 2v14a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2V6a2 2 0 0 1 2-2h2" />
            <path d="m9 14 2 2 4-4" strokeWidth="2.2" />
          </svg>
        </div>
      );

    case 'depreciation':
      // Depreciation Schedule Generator (Teal / Emerald Curve)
      return (
        <div
          className={`${containerSize} bg-gradient-to-br ${gradient} p-0.5 shadow-sm ring-1 ring-black/5 flex items-center justify-center text-white shrink-0 ${className}`}
        >
          <svg className={iconSize} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
            <path d="M3 3v18h18" />
            <path d="M7 8c4 1 7 5 13 10" strokeWidth="2.2" />
            <polyline points="15 18 20 18 20 13" strokeWidth="2.2" />
          </svg>
        </div>
      );

    case 'tb-checker':
    case 'trial-balance':
      // Trial Balance & Ledger Checker (Teal / Emerald Scrutiny Check)
      return (
        <div
          className={`${containerSize} bg-gradient-to-br ${gradient} p-0.5 shadow-sm ring-1 ring-black/5 flex items-center justify-center text-white shrink-0 ${className}`}
        >
          <svg className={iconSize} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
            <path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10" fill="currentColor" fillOpacity="0.2" />
            <path d="m9 12 2 2 4-4" strokeWidth="2.2" />
          </svg>
        </div>
      );

    case 'excel-templates':
      // Excel Template Generator (Cyan / Sky Layout)
      return (
        <div
          className={`${containerSize} bg-gradient-to-br ${gradient} p-0.5 shadow-sm ring-1 ring-black/5 flex items-center justify-center text-white shrink-0 ${className}`}
        >
          <svg className={iconSize} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
            <rect width="18" height="7" x="3" y="3" rx="1" fill="currentColor" fillOpacity="0.25" />
            <rect width="8" height="8" x="3" y="13" rx="1" fill="currentColor" fillOpacity="0.15" />
            <rect width="7" height="8" x="14" y="13" rx="1" fill="currentColor" fillOpacity="0.15" />
          </svg>
        </div>
      );

    case 'bank-rec':
      // Bank Reconciliation (Sky / Cobalt Rec)
      return (
        <div
          className={`${containerSize} bg-gradient-to-br ${gradient} p-0.5 shadow-sm ring-1 ring-black/5 flex items-center justify-center text-white shrink-0 ${className}`}
        >
          <svg className={iconSize} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
            <path d="M4 7h13a3 3 0 0 1 3 3v0a3 3 0 0 1-3 3H4" />
            <path d="m8 3-4 4 4 4" strokeWidth="2.2" />
            <path d="M20 17H7a3 3 0 0 1-3-3v0a3 3 0 0 1 3-3h13" />
            <path d="m16 21 4-4-4-4" strokeWidth="2.2" />
          </svg>
        </div>
      );

    case 'gst-rec':
      // GST Reconciliation (Violet / Indigo Portal Rec)
      return (
        <div
          className={`${containerSize} bg-gradient-to-br ${gradient} p-0.5 shadow-sm ring-1 ring-black/5 flex items-center justify-center text-white shrink-0 ${className}`}
        >
          <svg className={iconSize} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
            <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" fill="currentColor" fillOpacity="0.2" />
            <polyline points="14 2 14 8 20 8" />
            <path d="m9 15 2 2 4-4" strokeWidth="2.2" />
          </svg>
        </div>
      );

    case 'pdf-excel':
    case 'pdf':
    case 'pdf-to-excel':
      // PDF to Excel (Rose / Coral Converter)
      return (
        <div
          className={`${containerSize} bg-gradient-to-br ${gradient} p-0.5 shadow-sm ring-1 ring-black/5 flex items-center justify-center text-white shrink-0 ${className}`}
        >
          <svg className={iconSize} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
            <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8Z" fill="currentColor" fillOpacity="0.15" />
            <path d="M14 2v6h6" />
            <path d="M8 13h3" strokeWidth="2" />
            <path d="M8 17h8" strokeWidth="2" />
            <path d="M16 11l2 2-2 2" strokeWidth="2" />
          </svg>
        </div>
      );

    case 'invoices':
      // Invoice Processing OCR (Amber / Gold OCR)
      return (
        <div
          className={`${containerSize} bg-gradient-to-br ${gradient} p-0.5 shadow-sm ring-1 ring-black/5 flex items-center justify-center text-white shrink-0 ${className}`}
        >
          <svg className={iconSize} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
            <path d="M3 7V5a2 2 0 0 1 2-2h2" />
            <path d="M17 3h2a2 2 0 0 1 2 2v2" />
            <path d="M21 17v2a2 2 0 0 1-2 2h-2" />
            <path d="M7 21H5a2 2 0 0 1-2-2v-2" />
            <rect width="10" height="8" x="7" y="8" rx="1" fill="currentColor" fillOpacity="0.2" />
            <line x1="10" y1="12" x2="14" y2="12" strokeWidth="2" />
          </svg>
        </div>
      );

    case 'audit':
    case 'audit-tools':
      // Statutory Audit Tools (Deep Violet / Indigo Shield)
      return (
        <div
          className={`${containerSize} bg-gradient-to-br ${gradient} p-0.5 shadow-sm ring-1 ring-black/5 flex items-center justify-center text-white shrink-0 ${className}`}
        >
          <svg className={iconSize} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
            <path d="M20 13c0 5-3.5 7.5-7.66 8.95a1 1 0 0 1-.67-.01C7.5 20.5 4 18 4 13V6a1 1 0 0 1 1-1c2 0 4.5-1.2 6.24-2.72a1.17 1.17 0 0 1 1.52 0C14.51 3.81 17 5 19 5a1 1 0 0 1 1 1z" fill="currentColor" fillOpacity="0.2" />
            <path d="m9 12 2 2 4-4" strokeWidth="2.2" />
          </svg>
        </div>
      );

    case 'reports':
    case 'financial-reports':
      // Financial Reports Schedule III (Emerald Green Statements)
      return (
        <div
          className={`${containerSize} bg-gradient-to-br ${gradient} p-0.5 shadow-sm ring-1 ring-black/5 flex items-center justify-center text-white shrink-0 ${className}`}
        >
          <svg className={iconSize} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
            <line x1="18" y1="20" x2="18" y2="4" strokeWidth="2.5" />
            <line x1="12" y1="20" x2="12" y2="9" strokeWidth="2.5" />
            <line x1="6" y1="20" x2="6" y2="14" strokeWidth="2.5" />
            <path d="M3 20h18" strokeWidth="2" />
          </svg>
        </div>
      );

    case 'duplicate-invoice':
      // Duplicate Invoice Detector (Amber / Rose Double Invoices)
      return (
        <div
          className={`${containerSize} bg-gradient-to-br ${gradient} p-0.5 shadow-sm ring-1 ring-black/5 flex items-center justify-center text-white shrink-0 ${className}`}
        >
          <svg className={iconSize} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
            <rect x="7" y="3" width="13" height="15" rx="2" fill="currentColor" fillOpacity="0.2" />
            <path d="M4 8H3a1 1 0 0 0-1 1v11a2 2 0 0 0 2 2h11a1 1 0 0 0 1-1v-1" />
            <path d="m11 9 2 2 4-4" strokeWidth="2.2" />
          </svg>
        </div>
      );

    case 'journal-validator':
      // Journal Entry Validator (Emerald / Teal Balancing Scales)
      return (
        <div
          className={`${containerSize} bg-gradient-to-br ${gradient} p-0.5 shadow-sm ring-1 ring-black/5 flex items-center justify-center text-white shrink-0 ${className}`}
        >
          <svg className={iconSize} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
            <path d="M3 6h18" />
            <path d="M12 3v18" />
            <path d="M7 21h10" />
            <path d="m6 10-3 5h6z" fill="currentColor" fillOpacity="0.25" />
            <path d="m18 10-3 5h6z" fill="currentColor" fillOpacity="0.25" />
          </svg>
        </div>
      );

    case 'receivables-aging':
      // Receivables & Payables Aging (Sky / Indigo Timeline Clock)
      return (
        <div
          className={`${containerSize} bg-gradient-to-br ${gradient} p-0.5 shadow-sm ring-1 ring-black/5 flex items-center justify-center text-white shrink-0 ${className}`}
        >
          <svg className={iconSize} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
            <rect width="18" height="18" x="3" y="4" rx="2" fill="currentColor" fillOpacity="0.15" />
            <line x1="16" x2="16" y1="2" y2="6" />
            <line x1="8" x2="8" y1="2" y2="6" />
            <line x1="3" x2="21" y1="10" y2="10" />
            <circle cx="12" cy="15" r="3" fill="currentColor" fillOpacity="0.25" />
            <polyline points="12 13.5 12 15 13.5 15" strokeWidth="1.5" />
          </svg>
        </div>
      );

    case 'far-checker':
      // Fixed Asset Register Checker (Amber / Gold Machinery & Plant)
      return (
        <div
          className={`${containerSize} bg-gradient-to-br ${gradient} p-0.5 shadow-sm ring-1 ring-black/5 flex items-center justify-center text-white shrink-0 ${className}`}
        >
          <svg className={iconSize} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
            <path d="M2 20h20" />
            <path d="M5 20V8l4-4h6l4 4v12" fill="currentColor" fillOpacity="0.2" />
            <rect x="9" y="11" width="6" height="5" rx="1" />
            <path d="m9 7 3-3 3 3" />
          </svg>
        </div>
      );

    case 'audit-workpapers':
    case 'audit-workpaper':
      // Audit Working Paper Generator (Royal Blue / Navy Binder Folder)
      return (
        <div
          className={`${containerSize} bg-gradient-to-br ${gradient} p-0.5 shadow-sm ring-1 ring-black/5 flex items-center justify-center text-white shrink-0 ${className}`}
        >
          <svg className={iconSize} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
            <path d="M4 19.5v-15A2.5 2.5 0 0 1 6.5 2H20v20H6.5a2.5 2.5 0 0 1-2.5-2.5Z" fill="currentColor" fillOpacity="0.2" />
            <path d="M9 7h6" />
            <path d="M9 11h6" />
            <path d="M9 15h4" />
            <circle cx="6.5" cy="19.5" r="1.5" fill="currentColor" />
          </svg>
        </div>
      );

    case 'payroll-calculator':
    case 'payroll':
      // Payroll & Salary Calculator (Indigo / Purple Payslip)
      return (
        <div
          className={`${containerSize} bg-gradient-to-br ${gradient} p-0.5 shadow-sm ring-1 ring-black/5 flex items-center justify-center text-white shrink-0 ${className}`}
        >
          <svg className={iconSize} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
            <path d="M16 21v-2a4 4 0 0 0-4-4H6a4 4 0 0 0-4 4v2" fill="currentColor" fillOpacity="0.15" />
            <circle cx="9" cy="7" r="4" fill="currentColor" fillOpacity="0.25" />
            <path d="M17 11h5" strokeWidth="2" />
            <path d="M17 15h3" strokeWidth="2" />
            <path d="M17 7h4" strokeWidth="2" />
          </svg>
        </div>
      );

    case 'ratio-analyzer':
      // Financial Ratio Analyzer (Emerald / Cyan Analytics)
      return (
        <div
          className={`${containerSize} bg-gradient-to-br ${gradient} p-0.5 shadow-sm ring-1 ring-black/5 flex items-center justify-center text-white shrink-0 ${className}`}
        >
          <svg className={iconSize} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
            <line x1="18" y1="20" x2="18" y2="10" strokeWidth="2.2" />
            <line x1="12" y1="20" x2="12" y2="4" strokeWidth="2.2" />
            <line x1="6" y1="20" x2="6" y2="14" strokeWidth="2.2" />
            <path d="m4 12 5-5 4 4 7-7" strokeWidth="2" />
            <circle cx="12" cy="4" r="1.5" fill="currentColor" />
          </svg>
        </div>
      );

    case 'voucher-verification':
      // Voucher & Ledger Verification (Teal / Cyan Verification Seal)
      return (
        <div
          className={`${containerSize} bg-gradient-to-br ${gradient} p-0.5 shadow-sm ring-1 ring-black/5 flex items-center justify-center text-white shrink-0 ${className}`}
        >
          <svg className={iconSize} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
            <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" fill="currentColor" fillOpacity="0.15" />
            <polyline points="14 2 14 8 20 8" />
            <circle cx="12" cy="14" r="4" fill="currentColor" fillOpacity="0.25" />
            <path d="m10.5 14 1 1 2-2" strokeWidth="2" />
          </svg>
        </div>
      );

    case 'tds-reconciliation':
    case 'tds-rec':
      // TDS Reconciliation Tool (Warm Amber / Gold Statutory Shield)
      return (
        <div
          className={`${containerSize} bg-gradient-to-br ${gradient} p-0.5 shadow-sm ring-1 ring-black/5 flex items-center justify-center text-white shrink-0 ${className}`}
        >
          <svg className={iconSize} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
            <path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z" fill="currentColor" fillOpacity="0.2" />
            <path d="M8 8h6a2 2 0 0 1 0 4H8" strokeWidth="2" />
            <path d="M8 12h5l3 5" strokeWidth="2" />
            <path d="M7 10h8" strokeWidth="2" />
          </svg>
        </div>
      );

    case 'cashflow-generator':
    case 'cash-flow':
      // Cash Flow Statement Generator (Emerald / Mint Liquidity Wave)
      return (
        <div
          className={`${containerSize} bg-gradient-to-br ${gradient} p-0.5 shadow-sm ring-1 ring-black/5 flex items-center justify-center text-white shrink-0 ${className}`}
        >
          <svg className={iconSize} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
            <path d="M12 2v20" strokeWidth="2.2" />
            <path d="M17 5H9.5a3.5 3.5 0 0 0 0 7h5a3.5 3.5 0 0 1 0 7H6" strokeWidth="2.2" />
            <path d="M3 12c3-2 6-2 9 0s6 2 9 0" strokeOpacity="0.3" strokeWidth="2" />
          </svg>
        </div>
      );

    case 'excel-data-cleanup':
    case 'excel-validation':
    case 'excel-validator':
      // Excel Data Validation & Cleanup (Cyan / Emerald Data Cleaner)
      return (
        <div
          className={`${containerSize} bg-gradient-to-br ${gradient} p-0.5 shadow-sm ring-1 ring-black/5 flex items-center justify-center text-white shrink-0 ${className}`}
        >
          <svg className={iconSize} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
            <path d="M4 4h16v16H4z" fill="currentColor" fillOpacity="0.1" />
            <path d="M4 10h16" />
            <path d="M10 4v16" />
            <path d="m14 14 2 2 4-4" strokeWidth="2.2" />
          </svg>
        </div>
      );

    case 'purchase-invoice-matching':
    case 'purchase-matching':
      // Purchase Register vs Invoice Matching (Cobalt / Indigo Match)
      return (
        <div
          className={`${containerSize} bg-gradient-to-br ${gradient} p-0.5 shadow-sm ring-1 ring-black/5 flex items-center justify-center text-white shrink-0 ${className}`}
        >
          <svg className={iconSize} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
            <path d="M6 2 3 6v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2V6l-3-4Z" fill="currentColor" fillOpacity="0.15" />
            <path d="M3 6h18" />
            <path d="m9 14 2 2 4-4" strokeWidth="2.2" />
          </svg>
        </div>
      );

    case 'month-end-closing':
      // Month-End Closing Assistant (Royal Indigo / Purple Calendar)
      return (
        <div
          className={`${containerSize} bg-gradient-to-br ${gradient} p-0.5 shadow-sm ring-1 ring-black/5 flex items-center justify-center text-white shrink-0 ${className}`}
        >
          <svg className={iconSize} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
            <rect width="18" height="18" x="3" y="4" rx="2" fill="currentColor" fillOpacity="0.2" />
            <line x1="16" x2="16" y1="2" y2="6" />
            <line x1="8" x2="8" y1="2" y2="6" />
            <line x1="3" x2="21" y1="10" y2="10" />
            <path d="m9 16 2 2 4-4" strokeWidth="2.2" />
          </svg>
        </div>
      );

    case 'audit-sampling':
      // Audit Sampling Tool (Indigo / Blue Crosshair SA 530)
      return (
        <div
          className={`${containerSize} bg-gradient-to-br ${gradient} p-0.5 shadow-sm ring-1 ring-black/5 flex items-center justify-center text-white shrink-0 ${className}`}
        >
          <svg className={iconSize} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
            <circle cx="12" cy="12" r="9" fill="currentColor" fillOpacity="0.15" />
            <line x1="22" x2="18" y1="12" y2="12" strokeWidth="2" />
            <line x1="6" x2="2" y1="12" y2="12" strokeWidth="2" />
            <line x1="12" x2="12" y1="6" y2="2" strokeWidth="2" />
            <line x1="12" x2="12" y1="22" y2="18" strokeWidth="2" />
            <circle cx="12" cy="12" r="3" fill="currentColor" fillOpacity="0.3" />
          </svg>
        </div>
      );

    case 'balance-confirmation':
      // Balance Confirmation Tracker (Cyan / Azure Confirmation Letter)
      return (
        <div
          className={`${containerSize} bg-gradient-to-br ${gradient} p-0.5 shadow-sm ring-1 ring-black/5 flex items-center justify-center text-white shrink-0 ${className}`}
        >
          <svg className={iconSize} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
            <rect width="20" height="16" x="2" y="4" rx="2" fill="currentColor" fillOpacity="0.2" />
            <path d="m22 7-8.97 5.7a1.94 1.94 0 0 1-2.06 0L2 7" />
            <path d="m15 15 2 2 4-4" strokeWidth="2" />
          </svg>
        </div>
      );

    case 'client-documents':
      // Client Document Tracker & Requisitions (Sky / Royal Document Shield)
      return (
        <div
          className={`${containerSize} bg-gradient-to-br ${gradient} p-0.5 shadow-sm ring-1 ring-black/5 flex items-center justify-center text-white shrink-0 ${className}`}
        >
          <svg className={iconSize} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
            <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" fill="currentColor" fillOpacity="0.15" />
            <polyline points="14 2 14 8 20 8" />
            <rect x="8" y="12" width="8" height="6" rx="1" fill="currentColor" fillOpacity="0.25" />
            <path d="M10 12v-1a2 2 0 1 1 4 0v1" strokeWidth="1.8" />
          </svg>
        </div>
      );

    case 'recurring-compliance':
      // Recurring Compliance & Statutory Calendar (Teal / Cyan Sync Cycle)
      return (
        <div
          className={`${containerSize} bg-gradient-to-br ${gradient} p-0.5 shadow-sm ring-1 ring-black/5 flex items-center justify-center text-white shrink-0 ${className}`}
        >
          <svg className={iconSize} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
            <rect width="18" height="18" x="3" y="4" rx="2" fill="currentColor" fillOpacity="0.15" />
            <line x1="16" x2="16" y1="2" y2="6" />
            <line x1="8" x2="8" y1="2" y2="6" />
            <line x1="3" x2="21" y1="10" y2="10" />
            <path d="M8 15a4 4 0 0 1 7-2" strokeWidth="2" />
            <path d="M16 13v-2h-2" strokeWidth="2" />
            <path d="M16 16a4 4 0 0 1-7 2" strokeWidth="2" />
            <path d="M8 18v2h2" strokeWidth="2" />
          </svg>
        </div>
      );

    case 'multi-client-batch':
      // Multi-Client Batch Automation (Royal Indigo Multi-Stack Layers)
      return (
        <div
          className={`${containerSize} bg-gradient-to-br ${gradient} p-0.5 shadow-sm ring-1 ring-black/5 flex items-center justify-center text-white shrink-0 ${className}`}
        >
          <svg className={iconSize} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
            <polygon points="12 2 22 7.5 12 13 2 7.5 12 2" fill="currentColor" fillOpacity="0.25" />
            <path d="m2 12 10 5.5 10-5.5" strokeWidth="2.2" />
            <path d="m2 16.5 10 5.5 10-5.5" strokeWidth="2.2" />
          </svg>
        </div>
      );

    case 'workflow-orchestrator':
      return (
        <div
          className={`${containerSize} bg-gradient-to-br ${gradient} p-0.5 shadow-sm ring-1 ring-black/5 flex items-center justify-center text-white shrink-0 ${className}`}
        >
          <svg className={iconSize} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
            <circle cx="18" cy="18" r="3" />
            <circle cx="6" cy="6" r="3" />
            <path d="M6 9v12" />
            <path d="M18 9a9 9 0 0 0-9 9" />
          </svg>
        </div>
      );

    case 'form3cd-scrutiny':
    case 'form3cd':
      // Form 3CD Tax Audit Scrutiny (Authoritative Document with Balance Scale & Check)
      return (
        <div
          className={`${containerSize} bg-gradient-to-br ${gradient} p-0.5 shadow-sm ring-1 ring-black/5 flex items-center justify-center text-white shrink-0 ${className}`}
        >
          <svg className={iconSize} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
            <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" fill="currentColor" fillOpacity="0.15" />
            <polyline points="14 2 14 8 20 8" />
            <path d="m9 15 2 2 4-4" strokeWidth="2" />
            <line x1="8" x2="16" y1="11" y2="11" strokeWidth="1.8" />
          </svg>
        </div>
      );

    case 'gst-triple-rec':
    case 'triple-rec':
      // GSTR-1 vs 3B vs Books 3-Way Rec (Tripartite Tri-Directional Sync Cycle)
      return (
        <div
          className={`${containerSize} bg-gradient-to-br ${gradient} p-0.5 shadow-sm ring-1 ring-black/5 flex items-center justify-center text-white shrink-0 ${className}`}
        >
          <svg className={iconSize} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
            <path d="M21 12a9 9 0 0 0-9-9 9.75 9.75 0 0 0-6.74 2.74L3 8" />
            <polyline points="3 3 3 8 8 8" />
            <path d="M3 12a9 9 0 0 0 9 9 9.75 9.75 0 0 0 6.74-2.74L21 16" />
            <polyline points="16 16 21 16 21 21" />
            <circle cx="12" cy="12" r="2" fill="currentColor" />
          </svg>
        </div>
      );


    case 'tally-xml':
      // Tally Prime XML Direct Generator (XML Tag Code glyph)
      return (
        <div
          className={`${containerSize} bg-gradient-to-br ${gradient} p-0.5 shadow-sm ring-1 ring-black/5 flex items-center justify-center text-white shrink-0 ${className}`}
        >
          <svg className={iconSize} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
            <polyline points="16 18 22 12 16 6" />
            <polyline points="8 6 2 12 8 18" />
            <line x1="14" x2="10" y1="4" y2="20" strokeWidth="2" />
          </svg>
        </div>
      );

    case 'audit-dossier':
      // Executive Audit Dossier (ICAI SA-230 Certified File)
      return (
        <div
          className={`${containerSize} bg-gradient-to-br ${gradient} p-0.5 shadow-sm ring-1 ring-black/5 flex items-center justify-center text-white shrink-0 ${className}`}
        >
          <svg className={iconSize} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
            <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" />
            <polyline points="14 2 14 8 20 8" />
            <line x1="16" x2="8" y1="13" y2="13" />
            <line x1="16" x2="8" y1="17" y2="17" />
            <polyline points="10 9 9 9 8 9" />
          </svg>
        </div>
      );

    case 'client-vault':
      // Multi-Client Practice Vault
      return (
        <div
          className={`${containerSize} bg-gradient-to-br ${gradient} p-0.5 shadow-sm ring-1 ring-black/5 flex items-center justify-center text-white shrink-0 ${className}`}
        >
          <svg className={iconSize} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
            <path d="M3 21h18" />
            <path d="M5 21V7l8-4v18" />
            <path d="M19 21V11l-6-4" />
            <path d="M9 9h1" />
            <path d="M9 13h1" />
            <path d="M9 17h1" />
          </svg>
        </div>
      );

    case 'offline-engines':
      // Offline Practice Engine (Hard drive + server nodes)
      return (
        <div
          className={`${containerSize} bg-gradient-to-br ${gradient} p-0.5 shadow-sm ring-1 ring-black/5 flex items-center justify-center text-white shrink-0 ${className}`}
        >
          <svg className={iconSize} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
            <rect x="2" y="2" width="20" height="8" rx="2" ry="2" />
            <rect x="2" y="14" width="20" height="8" rx="2" ry="2" />
            <line x1="6" y1="6" x2="6.01" y2="6" strokeWidth="2" />
            <line x1="6" y1="18" x2="6.01" y2="18" strokeWidth="2" />
          </svg>
        </div>
      );

    default:
      // High-precision Fallback App Icon
      return (
        <div
          className={`${containerSize} bg-gradient-to-br ${gradient} p-0.5 shadow-sm ring-1 ring-black/5 flex items-center justify-center text-white shrink-0 ${className}`}
        >
          <svg className={iconSize} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
            <rect width="18" height="18" x="3" y="3" rx="2" fill="currentColor" fillOpacity="0.2" />
            <path d="M3 9h18" />
            <path d="M9 21V9" />
          </svg>
        </div>
      );
  }
};
