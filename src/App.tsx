import React, { useState, useEffect, useRef } from 'react';
import {
  FileSpreadsheet,
  ArrowLeftRight,
  ReceiptText,
  FileCode,
  FileText,
  ScanLine,
  ShieldCheck,
  FileBarChart,
  ArrowRight,
  ArrowUpRight,
  GitCompare,
  FileCog,
  Calculator,
  Scale,
  Receipt,
  FolderArchive,
  ClipboardCheck,
  TrendingDown,
  CheckCheck,
  LayoutTemplate,
  Search,
  X,
  Clock,
  Building,
  Mail,
  CalendarCheck,
  Layers,
  ArrowLeft,
  GitMerge,
  Workflow,
} from 'lucide-react';
import type { ToastMessage } from './types';
import { ToastContainer } from './components/common/Toast';
import { Sidebar, type NavItem, type SidebarCategory } from './components/common/Sidebar';
import { ToolIcon } from './components/common/ToolIcon';
import { ToolHeader } from './components/common/ToolHeader';
import { AuditorSignOffFooter } from './components/common/AuditorSignOffFooter';
import { SegmentedControl, type SegmentedTab } from './components/common/SegmentedControl';

import { ExcelCleanerTool } from './components/tools/ExcelCleanerTool';
import { BankReconciliationTool } from './components/tools/BankReconciliationTool';
import { GstReconciliationTool } from './components/tools/GstReconciliationTool';
import { PdfToExcelTool } from './components/tools/PdfToExcelTool';
import { InvoiceProcessingTool } from './components/tools/InvoiceProcessingTool';
import { AuditToolsTool } from './components/tools/AuditToolsTool';
import { FinancialReportsTool } from './components/tools/FinancialReportsTool';
import { FileComparisonTool } from './components/tools/FileComparisonTool';
import { BankStatementConverterTool } from './components/tools/BankStatementConverterTool';
import { FormulaCheckerTool } from './components/tools/FormulaCheckerTool';
import { LedgerComparisonTool } from './components/tools/LedgerComparisonTool';
import { InvoiceDataExtractorTool } from './components/tools/InvoiceDataExtractorTool';
import { BulkRenamerTool } from './components/tools/BulkRenamerTool';
import { AuditChecklistTool } from './components/tools/AuditChecklistTool';
import { DepreciationTool } from './components/tools/DepreciationTool';
import { TrialBalanceCheckerTool } from './components/tools/TrialBalanceCheckerTool';
import { ExcelTemplateGeneratorTool } from './components/tools/ExcelTemplateGeneratorTool';
import { DuplicateInvoiceTool } from './components/tools/DuplicateInvoiceTool';
import { JournalValidatorTool } from './components/tools/JournalValidatorTool';
import { ReceivablesAgingTool } from './components/tools/ReceivablesAgingTool';
import { FarCheckerTool } from './components/tools/FarCheckerTool';
import { AuditWorkpaperTool } from './components/tools/AuditWorkpaperTool';
import { PayrollTool } from './components/tools/PayrollTool';
import { RatioAnalyzerTool } from './components/tools/RatioAnalyzerTool';
import { VoucherVerificationTool } from './components/tools/VoucherVerificationTool';
import { TdsReconciliationTool } from './components/tools/TdsReconciliationTool';
import { CashFlowTool } from './components/tools/CashFlowTool';
import { ExcelDataValidationTool } from './components/tools/ExcelDataValidationTool';
import { PurchaseInvoiceMatchingTool } from './components/tools/PurchaseInvoiceMatchingTool';
import { MonthEndClosingTool } from './components/tools/MonthEndClosingTool';
import { AuditSamplingTool } from './components/tools/AuditSamplingTool';
import { BalanceConfirmationTool } from './components/tools/BalanceConfirmationTool';
import { ClientDocumentTool } from './components/tools/ClientDocumentTool';
import { RecurringComplianceTool } from './components/tools/RecurringComplianceTool';
import { MultiClientBatchTool } from './components/tools/MultiClientBatchTool';
import { Form3cdScrutinyTool } from './components/tools/Form3cdScrutinyTool';
import { GstTripleRecTool } from './components/tools/GstTripleRecTool';
import { WorkingFlowView } from './components/automation/WorkingFlowView';
import { SettingsView } from './components/settings/SettingsView';
import { PinLockScreen } from './components/common/PinLockScreen';
import { isPinEnabled } from './utils/security';
import { getActiveClient, type ClientProfile } from './utils/clientStorage';
import { ClientVaultModal } from './components/clients/ClientVaultModal';
import { TallyXmlExportModal } from './components/tally/TallyXmlExportModal';
import { ExecutiveAuditDossierModal } from './components/dossier/ExecutiveAuditDossierModal';

import { TallyXmlGeneratorTool } from './components/tools/TallyXmlGeneratorTool';
import { ExecutiveAuditDossierTool } from './components/tools/ExecutiveAuditDossierTool';
import { AppStartupAnimation } from './components/common/AppStartupAnimation';

type ToolId =
  | 'excel'
  | 'bank-rec'
  | 'gst-rec'
  | 'gst-triple-rec'
  | 'pdf-excel'
  | 'invoices'
  | 'audit'
  | 'reports'
  | 'file-compare'
  | 'bank-converter'
  | 'formula-checker'
  | 'ledger-compare'
  | 'invoice-extractor'
  | 'bulk-renamer'
  | 'audit-checklist'
  | 'depreciation'
  | 'tb-checker'
  | 'excel-templates'
  | 'duplicate-invoice'
  | 'journal-validator'
  | 'receivables-aging'
  | 'far-checker'
  | 'audit-workpapers'
  | 'payroll-calculator'
  | 'ratio-analyzer'
  | 'voucher-verification'
  | 'tds-reconciliation'
  | 'form3cd-scrutiny'
  | 'cashflow-generator'
  | 'excel-data-cleanup'
  | 'purchase-invoice-matching'
  | 'month-end-closing'
  | 'audit-sampling'
  | 'balance-confirmation'
  | 'client-documents'
  | 'recurring-compliance'
  | 'multi-client-batch'
  | 'tally-xml'
  | 'audit-dossier'
  | 'client-vault';



interface ToolCard {
  id: ToolId;
  name: string;
  description: string;
  category: SidebarCategory;
  categoryLabel: string;
  featureTag: string;
  badge?: string;
  icon: React.ComponentType<{ className?: string; strokeWidth?: number }>;
}

const TOOLS: ToolCard[] = [
  // ==========================================
  // TIER 1: CORE RECONCILIATIONS (TOP PRIORITY)
  // ==========================================
  {
    id: 'bank-rec',
    name: 'Bank Reconciliation (BRS)',
    description: 'Match bank transactions with ledger entries automatically using amount and reference heuristics.',
    category: 'reconciliation',
    categoryLabel: 'Reconciliation',
    featureTag: 'Bank & Ledger Rec',
    badge: 'Top Priority',
    icon: ArrowLeftRight,
  },
  {
    id: 'gst-rec',
    name: 'GSTR-2B vs Books Reconciliation',
    description: 'Compare purchase registers with GSTR-2B portal data to maximize input tax credit (ITC) claims.',
    category: 'tax',
    categoryLabel: 'Taxation & GST',
    featureTag: 'GSTR-2B vs Books',
    badge: 'Sec 16(2)(aa)',
    icon: ReceiptText,
  },
  {
    id: 'gst-triple-rec',
    name: 'GSTR-1 vs 3B vs Books',
    description: 'Tripartite monthly reconciliation of turnover and tax heads (IGST/CGST/SGST) with automated Rule 88C / DRC-01B notice risk detection.',
    category: 'tax',
    categoryLabel: 'Taxation & GST',
    featureTag: 'Rule 88C / DRC-01B',
    badge: 'Notice Risk',
    icon: ReceiptText,
  },
  {
    id: 'bank-converter',
    name: 'Bank Statement Converter',
    description: 'Convert PDF, Excel, and CSV statements with smart multi-bank header and column mapping.',
    category: 'reconciliation',
    categoryLabel: 'Reconciliation',
    featureTag: 'Multi-Bank Mapper',
    badge: 'High Value',
    icon: FileCog,
  },
  {
    id: 'tally-xml',
    name: 'Tally Prime Direct XML Generator',
    description: 'Generate 100% compliant Tally XML vouchers (Bank Charges, Interest, Journal, Payment) for direct 1-click import into Tally Prime & Tally.ERP 9.',
    category: 'reconciliation',
    categoryLabel: 'Reconciliation',
    featureTag: 'Tally Prime XML',
    badge: 'Tally Prime',
    icon: FileCode,
  },
  {
    id: 'tds-reconciliation',
    name: 'TDS Reconciliation',
    description: 'Reconcile TDS deductions with challan payments, detect Section 206AA invalid PANs, and compute Sec 201(1A) late interest.',
    category: 'tax',
    categoryLabel: 'Taxation & GST',
    featureTag: 'Sec 206AA & 201(1A)',
    badge: 'Statutory',
    icon: ShieldCheck,
  },
  {
    id: 'invoice-extractor',
    name: 'Invoice Data Extractor',
    description: 'Batch extract invoice PDFs & spreadsheets with statutory tax math checks and inline editing.',
    category: 'tax',
    categoryLabel: 'Taxation & GST',
    featureTag: 'GSTIN & Tax Check',
    icon: Receipt,
  },

  // ==========================================
  // TIER 2: AUDIT SCRUTINY & WORKFLOWS
  // ==========================================
  {
    id: 'tb-checker',
    name: 'Trial Balance Checker',
    description: 'Audit scrutiny for debit-credit imbalances, negative cash, inverted revenue/capital, and duplicate codes.',
    category: 'reconciliation',
    categoryLabel: 'Reconciliation',
    featureTag: 'Red Flag Detector',
    badge: 'Scrutiny',
    icon: CheckCheck,
  },
  {
    id: 'ledger-compare',
    name: 'Inter-Company Ledger Rec',
    description: 'Reconcile two ledgers with inter-company Dr/Cr inversion, voucher matching, and variance tracking.',
    category: 'reconciliation',
    categoryLabel: 'Reconciliation',
    featureTag: 'Dr/Cr Inversion',
    badge: 'Audit Grade',
    icon: Scale,
  },
  {
    id: 'multi-client-batch',
    name: 'Multi-Client Batch Audit',
    description: 'Run automated reconciliations across multiple clients in parallel with strict container isolation.',
    category: 'reconciliation',
    categoryLabel: 'Reconciliation',
    featureTag: 'Isolated Multi-Client',
    badge: 'Multi-Client',
    icon: Layers,
  },
  {
    id: 'receivables-aging',
    name: 'Receivables & Payables Aging',
    description: 'Classify open balances into configurable aging buckets with due dates, MSMED 45d alerts, and printable summaries.',
    category: 'reconciliation',
    categoryLabel: 'Reconciliation',
    featureTag: 'Schedule III & MSMED',
    badge: 'Schedule III',
    icon: Clock,
  },
  {
    id: 'duplicate-invoice',
    name: 'Duplicate Invoice Detector',
    description: 'Detect exact and probable duplicate vendor invoices using GSTIN, invoice #, close dates, and amounts.',
    category: 'tax',
    categoryLabel: 'Taxation & GST',
    featureTag: 'Side-by-Side Review',
    badge: 'Popular',
    icon: ReceiptText,
  },
  {
    id: 'reports',
    name: 'Financial Reports (Schedule III)',
    description: 'Prepare Schedule III compliant balance sheets, P&L statements, and working papers.',
    category: 'tax',
    categoryLabel: 'Compliance',
    featureTag: 'Schedule III Pack',
    icon: FileBarChart,
  },

  // ==========================================
  // TIER 3: STATUTORY AUDIT & CLOSING
  // ==========================================
  {
    id: 'audit-dossier',
    name: 'Executive Audit Dossier (ICAI SA-230)',
    description: 'Consolidated audit working papers summary with CA firm letterhead, client particulars, BRS/GST/TB metrics, and printable partner sign-off certification.',
    category: 'audit',
    categoryLabel: 'Statutory Audit',
    featureTag: 'ICAI SA-230',
    badge: 'Working Paper',
    icon: FileText,
  },
  {
    id: 'client-vault',
    name: 'Client Practice Vault',
    description: 'Manage, switch, and isolate client master profiles with GSTIN/PAN auto-derivation and zero cloud data leaks.',
    category: 'audit',
    categoryLabel: 'Practice Management',
    featureTag: 'Offline Multi-Client',
    badge: 'Vault',
    icon: Building,
  },
  {
    id: 'month-end-closing',
    name: 'Period-End Closing Checklist',
    description: 'Track recurring closing activities across BRS, cut-offs, provisions, and statutory reviews with audit sign-off memo.',
    category: 'audit',
    categoryLabel: 'Statutory Audit',
    featureTag: 'Checklist & Sign-Off',
    badge: 'Workflow',
    icon: ClipboardCheck,
  },
  {
    id: 'audit-checklist',
    name: 'Audit Checklist & Engagement',
    description: 'Statutory (Companies Act 2013), Tax (Form 3CD), and GST audit checklists with client requisition exports.',
    category: 'audit',
    categoryLabel: 'Statutory Audit',
    featureTag: 'Companies Act & 3CD',
    badge: 'Statutory',
    icon: ClipboardCheck,
  },
  {
    id: 'form3cd-scrutiny',
    name: 'Form 3CD Tax Audit',
    description: 'Scrutinize Clause 21(d) cash limits (> ₹10k/35k), Clause 44 GST spend break-up, Clause 34 Chapter XVII-B TDS, and Clause 21(a) disallowances.',
    category: 'audit',
    categoryLabel: 'Statutory Audit',
    featureTag: 'Sec 44AB & Clause 44',
    badge: 'Sec 44AB',
    icon: ShieldCheck,
  },
  {
    id: 'audit',
    name: 'CARO Audit Scrutiny',
    description: 'Run audit sampling, materiality thresholds, CARO 2020 verification, and compliance checks.',
    category: 'audit',
    categoryLabel: 'Statutory Audit',
    featureTag: 'Sampling & CARO',
    icon: ShieldCheck,
  },
  {
    id: 'audit-sampling',
    name: 'Audit Sampling',
    description: 'Select statistically valid or high-value targeted samples under SA 530 with monetary coverage and vouching registers.',
    category: 'audit',
    categoryLabel: 'Statutory Audit',
    featureTag: 'SA 530 & High-Value',
    badge: 'SA Compliant',
    icon: ShieldCheck,
  },
  {
    id: 'depreciation',
    name: 'Depreciation Calculator',
    description: 'Automate Companies Act Schedule II (WDV/SLM), Income Tax Sec 32 block depr, and AS 22 deferred tax.',
    category: 'audit',
    categoryLabel: 'Statutory Audit',
    featureTag: 'Sec 32 & Half-Rate',
    badge: 'AS 22 Ready',
    icon: TrendingDown,
  },
  {
    id: 'far-checker',
    name: 'Fixed Asset Register (FAR)',
    description: 'Audit asset IDs, Schedule II useful lives, 5% residual values, additions < 180 days, and recompute depreciation.',
    category: 'audit',
    categoryLabel: 'Statutory Audit',
    featureTag: 'Schedule II & CARO',
    badge: 'Statutory',
    icon: Building,
  },
  {
    id: 'recurring-compliance',
    name: 'Statutory Compliance Calendar',
    description: 'Track GSTR-1, GSTR-3B, TDS, and Advance Tax statutory due dates with filing proof verification.',
    category: 'tax',
    categoryLabel: 'Taxation & GST',
    featureTag: 'Statutory Calendar',
    badge: 'Statutory',
    icon: CalendarCheck,
  },
  {
    id: 'balance-confirmation',
    name: 'Balance Confirmation (SA 505)',
    description: 'Manage external confirmation requests under SA 505 for debtors, creditors, and banks with draft letters and variance reports.',
    category: 'audit',
    categoryLabel: 'Statutory Audit',
    featureTag: 'SA 505 Confirmations',
    badge: 'Direct Response',
    icon: Receipt,
  },

  // ==========================================
  // TIER 4: OPERATIONAL & VERIFICATION TOOLS
  // ==========================================
  {
    id: 'purchase-invoice-matching',
    name: 'Purchase vs Invoice Matching',
    description: 'Reconcile Purchase Register with Supplier Invoices to detect amount, date, and tax variances, and unbooked invoices.',
    category: 'reconciliation',
    categoryLabel: 'Reconciliation',
    featureTag: 'PR vs Invoices (6 Sheets)',
    badge: 'Audit Grade',
    icon: Scale,
  },
  {
    id: 'voucher-verification',
    name: 'Voucher & Ledger Verification',
    description: 'Cross-verify Daybook vouchers against General Ledger postings to detect unposted items, variances, and duplicate numbers.',
    category: 'reconciliation',
    categoryLabel: 'Reconciliation',
    featureTag: 'Daybook vs General Ledger',
    badge: 'Audit Grade',
    icon: Scale,
  },
  {
    id: 'journal-validator',
    name: 'Journal Entry Validator',
    description: 'Audit voucher debit-credit imbalances, Sec 40A(3) cash limits, round sum anomalies, and missing dates.',
    category: 'audit',
    categoryLabel: 'Statutory Audit',
    featureTag: 'Voucher Imbalance & 40A(3)',
    badge: 'Audit Scrutiny',
    icon: Scale,
  },
  {
    id: 'audit-workpapers',
    name: 'Audit Working Papers',
    description: 'Generate standardized lead schedules (WP-A to I), materiality analytical reviews, testing plans, and confirmations.',
    category: 'audit',
    categoryLabel: 'Statutory Audit',
    featureTag: 'Lead Schedules (WP-A to I)',
    badge: 'SA Compliant',
    icon: ShieldCheck,
  },
  {
    id: 'ratio-analyzer',
    name: 'Financial Ratio Analyzer',
    description: 'Analyze Balance Sheet & P&L ratios across comparative periods with formula transparency, ICAI benchmarks, and reports.',
    category: 'audit',
    categoryLabel: 'Financial Analysis',
    featureTag: 'Comparative Ratios & ICAI',
    badge: 'Popular',
    icon: Calculator,
  },
  {
    id: 'cashflow-generator',
    name: 'Cash Flow Statement',
    description: 'Prepare AS-3 / Ind AS 7 Indirect Method Cash Flow Statements with working capital movements and cash tie-out validation.',
    category: 'audit',
    categoryLabel: 'Financial Statements',
    featureTag: 'AS-3 / Ind AS 7 Indirect',
    badge: 'AS-3 Compliant',
    icon: FileSpreadsheet,
  },
  {
    id: 'payroll-calculator',
    name: 'Payroll & Salary Calculator',
    description: 'Import employee salary data, calculate gross-to-net, verify statutory PF/ESI/PT/TDS rules, and export payslips.',
    category: 'audit',
    categoryLabel: 'Payroll & HR Audit',
    featureTag: 'PF, ESI & Payslip PDF',
    badge: 'Statutory',
    icon: Receipt,
  },
  {
    id: 'client-documents',
    name: 'Client Document Tracker',
    description: 'Track client document lists, verify SHA-256 upload integrity hashes, and draft follow-up reminder notices.',
    category: 'audit',
    categoryLabel: 'Statutory Audit',
    featureTag: 'SHA-256 & Requisitions',
    badge: 'Automation',
    icon: Mail,
  },
  {
    id: 'invoices',
    name: 'Invoice Processing OCR',
    description: 'Extract and validate vendor invoice information with line-item breakdowns and statutory taxes.',
    category: 'tax',
    categoryLabel: 'Taxation & GST',
    featureTag: 'Vendor OCR Parse',
    icon: ScanLine,
  },

  // ==========================================
  // TIER 5: EXCEL ENGINES & FILE UTILITIES
  // ==========================================
  {
    id: 'excel',
    name: 'Excel Sheet Cleaner',
    description: 'Clean whitespace, standardize dates, fix column widths, and eliminate ### cell errors.',
    category: 'excel',
    categoryLabel: 'Excel Engine',
    featureTag: 'Format Preserver',
    icon: FileSpreadsheet,
  },
  {
    id: 'file-compare',
    name: 'File Comparison & Diff',
    description: 'Compare two Excel or CSV registers with cell-level change tracking and key column matching.',
    category: 'excel',
    categoryLabel: 'Excel Engine',
    featureTag: '5-Sheet Diff Report',
    badge: 'Popular',
    icon: GitCompare,
  },
  {
    id: 'pdf-excel',
    name: 'PDF to Excel Converter',
    description: 'Extract structured tabular financial data from digital PDF bank statements and accounting ledgers.',
    category: 'excel',
    categoryLabel: 'Excel Engine',
    featureTag: 'Tabular PDF Engine',
    icon: FileCode,
  },
  {
    id: 'formula-checker',
    name: 'Formula & Integrity Checker',
    description: 'Scan workbooks for #REF!, #DIV/0!, broken references, and hardcoded numbers in formula ranges.',
    category: 'excel',
    categoryLabel: 'Excel Engine',
    featureTag: 'Error Scanner & Fix',
    icon: Calculator,
  },
  {
    id: 'excel-data-cleanup',
    name: 'Excel Data Cleanup',
    description: 'Diagnose and clean whitespace, date formatting, duplicate rows, empty cells, and casing inconsistencies with itemized approvals.',
    category: 'excel',
    categoryLabel: 'Excel Engine',
    featureTag: 'Data Quality & Cleanse',
    badge: 'Popular',
    icon: FileSpreadsheet,
  },
  {
    id: 'excel-templates',
    name: 'Excel Accounting Templates',
    description: 'Ready-to-use professional accounting templates for BRS, GSTR-2B, petty cash, ledger, and working papers.',
    category: 'excel',
    categoryLabel: 'Excel Engine',
    featureTag: '7 Ready Schedules',
    badge: 'New Tool',
    icon: LayoutTemplate,
  },
  {
    id: 'bulk-renamer',
    name: 'Bulk File Renamer',
    description: 'Rename multiple invoice PDFs, vouchers, and statements with collision protection & ZIP archive.',
    category: 'excel',
    categoryLabel: 'Utilities',
    featureTag: 'ZIP & Manifest',
    badge: 'New Tool',
    icon: FolderArchive,
  },
];

export function App() {
  const [showStartup, setShowStartup] = useState(true);
  const [activeNav, setActiveNav] = useState<NavItem>('home');
  const [selectedCategory, setSelectedCategory] = useState<SidebarCategory>('all');
  const [selectedTool, setSelectedTool] = useState<ToolId | null>(null);
  const [searchQuery, setSearchQuery] = useState('');
  const [toasts, setToasts] = useState<ToastMessage[]>([]);
  const [isLocked, setIsLocked] = useState(() => isPinEnabled());
  const [activeClient, setActiveClientState] = useState<ClientProfile>(() => getActiveClient());
  const [isClientVaultOpen, setIsClientVaultOpen] = useState(false);
  const [isTallyXmlOpen, setIsTallyXmlOpen] = useState(false);
  const [isAuditDossierOpen, setIsAuditDossierOpen] = useState(false);
  const mainContentRef = useRef<HTMLElement>(null);

  // Sync client profile whenever client is switched across the app
  useEffect(() => {
    const handleClientSwitch = () => {
      setActiveClientState(getActiveClient());
    };
    window.addEventListener('ca_copilot_client_switched', handleClientSwitch);
    return () => window.removeEventListener('ca_copilot_client_switched', handleClientSwitch);
  }, []);

  // Auto-scroll to top immediately whenever a tool is opened, closed, or switched
  useEffect(() => {
    if (mainContentRef.current) {
      mainContentRef.current.scrollTop = 0;
      mainContentRef.current.scrollTo({ top: 0, left: 0, behavior: 'instant' });
    }
    window.scrollTo({ top: 0, left: 0, behavior: 'instant' });
    document.documentElement.scrollTop = 0;
    document.body.scrollTop = 0;
  }, [selectedTool, selectedCategory, activeNav]);

  const showToast = (_title: string, _description?: string, _type: 'success' | 'error' | 'info' = 'success') => {
    // All notifications and popups disabled per user request
  };

  const dismissToast = (_id: string) => {
    // No-op
  };

  const handleNavSelect = (item: NavItem) => {
    setActiveNav(item);
    setSelectedTool(null);
    if (item === 'all-tools') {
      setSelectedCategory('all');
    }
  };

  const handleCategorySelect = (cat: SidebarCategory) => {
    setSelectedCategory(cat);
    setActiveNav('home');
    setSelectedTool(null);
  };

  const filteredTools = TOOLS.filter((tool) => {
    const matchesCategory =
      selectedCategory === 'all' || tool.category === selectedCategory;
    const matchesSearch =
      tool.name.toLowerCase().includes(searchQuery.toLowerCase()) ||
      tool.description.toLowerCase().includes(searchQuery.toLowerCase()) ||
      tool.featureTag.toLowerCase().includes(searchQuery.toLowerCase()) ||
      tool.categoryLabel.toLowerCase().includes(searchQuery.toLowerCase());
    return matchesCategory && matchesSearch;
  });

  const activeToolObj = selectedTool ? TOOLS.find((t) => t.id === selectedTool) : null;

  const CATEGORY_TABS: SegmentedTab<SidebarCategory>[] = [
    { id: 'all', label: 'All Tools' },
    { id: 'reconciliation', label: 'Reconciliations' },
    { id: 'audit', label: 'Audit & Compliance' },
    { id: 'tax', label: 'Tax & Invoices' },
    { id: 'excel', label: 'Excel Utilities' },
  ];

  const categoryLabels: Record<SidebarCategory, string> = {
    all: 'All Tools',
    reconciliation: 'Reconciliations',
    audit: 'Audit & Compliance',
    tax: 'Tax & Invoices',
    excel: 'Excel Utilities',
  };

  return (
    <>
      {showStartup && (
        <AppStartupAnimation onComplete={() => setShowStartup(false)} />
      )}
      <div className="h-screen w-full bg-[#F8F9FA] text-gray-900 font-sans antialiased flex overflow-hidden print:h-auto print:overflow-visible print:bg-white">
      {/* 1. Uniform Sidebar with Brand at Top & Settings Pinned at Bottom */}
      <Sidebar
        activeNav={activeNav}
        selectedCategory={selectedCategory}
        selectedTool={selectedTool}
        isPinActive={isPinEnabled()}
        onLockApp={() => setIsLocked(true)}
        onSelectNav={handleNavSelect}
        onSelectCategory={handleCategorySelect}
        onSelectTool={(toolId) => {
          setSelectedTool(toolId as ToolId);
          setActiveNav('home');
        }}
      />

        {/* Right side container with top window drag region */}
        <div className="flex-1 min-w-0 h-full flex flex-col overflow-hidden">
          {/* Top Window Drag Bar (Window Controls sit at the top right of this bar) */}
          <div className="h-9 w-full shrink-0 select-none app-drag-region flex items-center justify-end pr-36 print:hidden">
            {/* Draggable empty region */}
          </div>

          {/* Main Content Canvas: Laser-aligned with Header at px-8 */}
          <main
            ref={mainContentRef}
            className="flex-1 min-w-0 overflow-y-auto bg-[#F8F9FA] px-8 pb-7 pt-1 print:p-0 print:m-0 print:h-auto print:overflow-visible print:bg-white"
          >
            <div className="w-full max-w-[1400px] space-y-6 print:max-w-none print:w-full print:space-y-0">
            {/* View 1: Practice Settings */}
            {activeNav === 'settings' ? (
              <SettingsView
                onShowToast={showToast}
                onLockApp={() => setIsLocked(true)}
                onOpenClientVault={() => setIsClientVaultOpen(true)}
              />
            ) : activeNav === 'working-flow' && selectedTool === null ? (
              <WorkingFlowView
                onSelectTool={(toolId) => {
                  setSelectedTool(toolId as ToolId);
                  setActiveNav('home');
                }}
                onShowToast={showToast}
                selectedClientId=""
              />
            ) : selectedTool !== null ? (


              /* View 2: Active Tool Workspace */
              <div key={selectedTool} className="animate-apple-in">
                {selectedTool === 'excel' && (
                  <ExcelCleanerTool
                    onBack={() => setSelectedTool(null)}
                    onShowToast={showToast}
                  />
                )}

                {selectedTool === 'bank-rec' && (
                  <BankReconciliationTool
                    onBack={() => setSelectedTool(null)}
                    onShowToast={showToast}
                  />
                )}

                {selectedTool === 'gst-rec' && (
                  <GstReconciliationTool
                    onBack={() => setSelectedTool(null)}
                    onShowToast={showToast}
                  />
                )}

                {selectedTool === 'pdf-excel' && (
                  <PdfToExcelTool
                    onBack={() => setSelectedTool(null)}
                    onShowToast={showToast}
                  />
                )}

                {selectedTool === 'invoices' && (
                  <InvoiceProcessingTool
                    onBack={() => setSelectedTool(null)}
                    onShowToast={showToast}
                  />
                )}

                {selectedTool === 'audit' && (
                  <AuditToolsTool
                    onBack={() => setSelectedTool(null)}
                    onShowToast={showToast}
                  />
                )}

                {selectedTool === 'reports' && (
                  <FinancialReportsTool
                    onBack={() => setSelectedTool(null)}
                    onShowToast={showToast}
                  />
                )}

                {selectedTool === 'file-compare' && (
                  <FileComparisonTool
                    onBack={() => setSelectedTool(null)}
                    onShowToast={showToast}
                  />
                )}

                {selectedTool === 'bank-converter' && (
                  <BankStatementConverterTool
                    onBack={() => setSelectedTool(null)}
                    onShowToast={showToast}
                  />
                )}

                {selectedTool === 'formula-checker' && (
                  <FormulaCheckerTool
                    onBack={() => setSelectedTool(null)}
                    onShowToast={showToast}
                  />
                )}

                {selectedTool === 'ledger-compare' && (
                  <LedgerComparisonTool
                    onBack={() => setSelectedTool(null)}
                    onShowToast={showToast}
                  />
                )}

                {selectedTool === 'invoice-extractor' && (
                  <InvoiceDataExtractorTool
                    onBack={() => setSelectedTool(null)}
                    onShowToast={showToast}
                  />
                )}

                {selectedTool === 'bulk-renamer' && (
                  <BulkRenamerTool
                    onBack={() => setSelectedTool(null)}
                    onShowToast={showToast}
                  />
                )}

                {selectedTool === 'audit-checklist' && (
                  <AuditChecklistTool
                    onBack={() => setSelectedTool(null)}
                    onShowToast={showToast}
                  />
                )}

                {selectedTool === 'depreciation' && (
                  <DepreciationTool
                    onBack={() => setSelectedTool(null)}
                    onShowToast={showToast}
                  />
                )}

                {selectedTool === 'tb-checker' && (
                  <TrialBalanceCheckerTool
                    onBack={() => setSelectedTool(null)}
                    onShowToast={showToast}
                  />
                )}

                {selectedTool === 'excel-templates' && (
                  <ExcelTemplateGeneratorTool
                    onBack={() => setSelectedTool(null)}
                    onShowToast={showToast}
                  />
                )}

                {selectedTool === 'duplicate-invoice' && (
                  <DuplicateInvoiceTool
                    onBack={() => setSelectedTool(null)}
                    onShowToast={showToast}
                  />
                )}

                {selectedTool === 'journal-validator' && (
                  <JournalValidatorTool
                    onBack={() => setSelectedTool(null)}
                    onShowToast={showToast}
                  />
                )}

                {selectedTool === 'receivables-aging' && (
                  <ReceivablesAgingTool
                    onBack={() => setSelectedTool(null)}
                    onShowToast={showToast}
                  />
                )}

                {selectedTool === 'far-checker' && (
                  <FarCheckerTool
                    onBack={() => setSelectedTool(null)}
                    onShowToast={showToast}
                  />
                )}

                {selectedTool === 'audit-workpapers' && (
                  <AuditWorkpaperTool
                    onBack={() => setSelectedTool(null)}
                    onShowToast={showToast}
                  />
                )}

                {selectedTool === 'payroll-calculator' && (
                  <PayrollTool
                    onBack={() => setSelectedTool(null)}
                    onShowToast={showToast}
                  />
                )}

                {selectedTool === 'ratio-analyzer' && (
                  <RatioAnalyzerTool
                    onBack={() => setSelectedTool(null)}
                    onShowToast={showToast}
                  />
                )}

                {selectedTool === 'voucher-verification' && (
                  <VoucherVerificationTool
                    onBack={() => setSelectedTool(null)}
                    onShowToast={showToast}
                  />
                )}

                {selectedTool === 'tds-reconciliation' && (
                  <TdsReconciliationTool
                    onBack={() => setSelectedTool(null)}
                    onShowToast={showToast}
                  />
                )}

                {selectedTool === 'cashflow-generator' && (
                  <CashFlowTool
                    onBack={() => setSelectedTool(null)}
                    onShowToast={showToast}
                  />
                )}

                {selectedTool === 'excel-data-cleanup' && (
                  <ExcelDataValidationTool
                    onBack={() => setSelectedTool(null)}
                    onShowToast={showToast}
                  />
                )}

                {selectedTool === 'purchase-invoice-matching' && (
                  <PurchaseInvoiceMatchingTool
                    onBack={() => setSelectedTool(null)}
                    onShowToast={showToast}
                  />
                )}

                {selectedTool === 'month-end-closing' && (
                  <MonthEndClosingTool
                    onBack={() => setSelectedTool(null)}
                    onShowToast={showToast}
                  />
                )}

                {selectedTool === 'audit-sampling' && (
                  <AuditSamplingTool
                    onBack={() => setSelectedTool(null)}
                    onShowToast={showToast}
                  />
                )}

                {selectedTool === 'balance-confirmation' && (
                  <BalanceConfirmationTool
                    onBack={() => setSelectedTool(null)}
                    onShowToast={showToast}
                  />
                )}

                {selectedTool === 'client-documents' && (
                  <ClientDocumentTool
                    onBack={() => setSelectedTool(null)}
                    onShowToast={showToast}
                  />
                )}

                {selectedTool === 'recurring-compliance' && (
                  <RecurringComplianceTool
                    onBack={() => setSelectedTool(null)}
                    onShowToast={showToast}
                  />
                )}

                {selectedTool === 'multi-client-batch' && (
                  <MultiClientBatchTool
                    onBack={() => setSelectedTool(null)}
                    onShowToast={showToast}
                  />
                )}

                {selectedTool === 'form3cd-scrutiny' && (
                  <Form3cdScrutinyTool
                    onBack={() => setSelectedTool(null)}
                    onShowToast={showToast}
                  />
                )}

                {selectedTool === 'gst-triple-rec' && (
                  <GstTripleRecTool
                    onBack={() => setSelectedTool(null)}
                    onShowToast={showToast}
                  />
                )}

                {selectedTool === 'tally-xml' && (
                  <TallyXmlGeneratorTool
                    onBack={() => setSelectedTool(null)}
                    onShowToast={showToast}
                    onOpenClientVault={() => setIsClientVaultOpen(true)}
                  />
                )}

                {selectedTool === 'audit-dossier' && (
                  <ExecutiveAuditDossierTool
                    onBack={() => setSelectedTool(null)}
                    onShowToast={showToast}
                  />
                )}

                {/* Official ICAI SA-230 Auditor Sign-off Certification (rendered only on print) */}
                <AuditorSignOffFooter />
              </div>
            ) : (
              /* View 3: Clean, Professional CA Tools Grid (Hidden on print) */
              <div className="space-y-5 print:hidden animate-apple-fade">
                {/* Clean Bordered Header Card */}
                <div className="bg-white border border-gray-200 rounded-xl p-5 shadow-2xs space-y-4">
                  {/* Row 1: Title & Description Left, Search Input Right */}
                  <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4">
                    <div>
                      <h1 className="text-2xl sm:text-3xl font-extrabold text-gray-900 tracking-tight leading-tight">
                        CA Copilot Workspace
                      </h1>
                      <p className="text-xs sm:text-sm text-gray-600 mt-1.5 font-normal leading-relaxed max-w-2xl">
                        Smart accounting and audit automation tools built for Chartered Accountants.
                      </p>
                    </div>

                    {/* Symmetrical Search Input */}
                    <div className="relative w-full sm:w-72 shrink-0">
                      <Search className="w-4 h-4 text-gray-400 absolute left-3 top-1/2 -translate-y-1/2 pointer-events-none" />
                      <input
                        type="text"
                        value={searchQuery}
                        onChange={(e) => setSearchQuery(e.target.value)}
                        placeholder="Filter automations..."
                        className="w-full pl-9 pr-8 py-2 bg-white border border-gray-200 rounded-lg text-xs text-gray-900 placeholder-gray-400 focus:outline-none focus:border-gray-900 shadow-2xs transition-colors"
                      />
                      {searchQuery && (
                        <button
                          onClick={() => setSearchQuery('')}
                          className="absolute right-2.5 top-1/2 -translate-y-1/2 text-gray-400 hover:text-gray-600 p-0.5 cursor-pointer"
                        >
                          <X className="w-3.5 h-3.5" />
                        </button>
                      )}
                    </div>
                  </div>

                  {/* Row 2: Category Tabs (No Numbers!) & Counter */}
                  <div className="pt-3.5 border-t border-gray-100 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                    {/* Clean Enclosed Boxed Segmented Control with Apple Sliding Animation */}
                    <SegmentedControl
                      tabs={CATEGORY_TABS}
                      activeTab={selectedCategory}
                      onChange={(catId) => setSelectedCategory(catId)}
                    />

                    {/* Symmetrical Count Indicator */}
                    <div className="text-xs text-gray-500 font-medium flex items-center gap-2 self-end sm:self-auto">
                      <span className="font-semibold text-gray-900 font-mono transition-all duration-200">{filteredTools.length}</span>
                      <span>{filteredTools.length === 1 ? 'tool available' : 'tools available'}</span>
                      {selectedCategory !== 'all' && (
                        <button
                          onClick={() => setSelectedCategory('all')}
                          className="text-gray-400 hover:text-gray-700 underline text-[11px] ml-1 cursor-pointer"
                        >
                          Reset
                        </button>
                      )}
                    </div>
                  </div>
                </div>

                {/* Real Figma/Dribbble Style Tool Cards Grid with Smooth Transition */}
                <div
                  key={selectedCategory}
                  className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-4"
                >
                  {filteredTools.map((tool, index) => {
                    return (
                      <div
                        key={tool.id}
                        onClick={() => {
                          if (tool.id === 'client-vault') {
                            setIsClientVaultOpen(true);
                          } else {
                            setSelectedTool(tool.id);
                          }
                        }}
                        style={{
                          animationDelay: `${Math.min(index * 25, 200)}ms`,
                        }}
                        className="group bg-white border border-gray-200/90 rounded-xl p-5 cursor-pointer select-none apple-card flex flex-col justify-between shadow-2xs animate-card-in"
                      >
                        <div>
                          <div className="flex items-center justify-between gap-3 mb-3.5">
                            <ToolIcon
                              toolId={tool.id}
                              size="md"
                              className="group-hover:scale-105 transition-transform duration-200"
                            />
                            <div className="w-7 h-7 rounded-lg bg-gray-50 border border-gray-200/80 group-hover:bg-gray-100 group-hover:border-gray-300 flex items-center justify-center text-gray-400 group-hover:text-gray-900 transition-all duration-200 shrink-0 shadow-2xs">
                              <ArrowRight
                                className="w-3.5 h-3.5 group-hover:translate-x-0.5 transition-transform duration-200"
                                strokeWidth={2}
                              />
                            </div>
                          </div>

                          <h2 className="text-sm font-semibold text-gray-900 group-hover:text-gray-900 transition-colors">
                            {tool.name}
                          </h2>
                          <p className="text-xs text-gray-500 leading-relaxed mt-1.5 line-clamp-2">
                            {tool.description}
                          </p>
                        </div>
                      </div>
                    );
                  })}
                </div>

                {/* Zero State if Filter Doesn't Match */}
                {filteredTools.length === 0 && (
                  <div className="text-center py-12 bg-white rounded-xl border border-gray-200 shadow-xs">
                    <p className="text-xs text-gray-500">
                      No tools matched your search query "{searchQuery}".
                    </p>
                    <button
                      onClick={() => {
                        setSearchQuery('');
                        setSelectedCategory('all');
                      }}
                      className="mt-2 text-xs text-gray-900 font-semibold hover:underline"
                    >
                      Clear filter
                    </button>
                  </div>
                )}
              </div>
            )}
          </div>
        </main>
      </div>

      {/* 4-Digit Security PIN App Lock Screen */}
      {isLocked && <PinLockScreen onUnlock={() => setIsLocked(false)} />}

      {/* Client Practice Vault Modal */}
      <ClientVaultModal
        isOpen={isClientVaultOpen}
        onClose={() => {
          setIsClientVaultOpen(false);
          setActiveClientState(getActiveClient());
        }}
        onShowToast={showToast}
      />

      {/* Tally Prime Direct XML Generator Modal */}
      <TallyXmlExportModal
        isOpen={isTallyXmlOpen}
        onClose={() => setIsTallyXmlOpen(false)}
        clientName={activeClient.name}
        onShowToast={showToast}
      />

      {/* Executive Audit Dossier Modal */}
      <ExecutiveAuditDossierModal
        isOpen={isAuditDossierOpen}
        onClose={() => setIsAuditDossierOpen(false)}
        onShowToast={showToast}
      />

      {/* Real-time Toast Notifications */}
      <ToastContainer toasts={toasts} onDismiss={dismissToast} />
    </div>
    </>
  );
}

export default App;
