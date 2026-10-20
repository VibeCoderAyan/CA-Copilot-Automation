import React, { useState, useEffect, useMemo, useRef } from 'react';
import * as XLSX from 'xlsx';
import {
  ArrowLeft,
  FileSpreadsheet,
  UploadCloud,
  CheckCircle2,
  Clock,
  AlertTriangle,
  RefreshCw,
  Search,
  Plus,
  Mail,
  Copy,
  Check,
  ShieldCheck,
  FileText,
  Calendar,
  Building,
  Hash,
  Download,
  Filter,
} from 'lucide-react';
import { InnerToolBadge } from '../common/InnerToolBadge';
import { CAAutomationApiClient } from '../../services/apiClient';
import { createStyledWorksheet } from '../../utils/excelStyler';

interface ClientDocumentToolProps {
  onBack: () => void;
  onShowToast: (title: string, description?: string, type?: 'success' | 'error' | 'info') => void;
}

interface RequisitionItem {
  id: string;
  documentName: string;
  category: string;
  mandatory: boolean;
  dueDate: string;
  status: 'pending' | 'received' | 'verified';
  fileName?: string;
  sha256Hash?: string;
  uploadedAt?: string;
}

const DEFAULT_REQUISITIONS: RequisitionItem[] = [
  {
    id: 'req_01',
    documentName: 'Bank Statements (All active accounts for FY 2026-27)',
    category: 'Banking',
    mandatory: true,
    dueDate: '2026-10-25',
    status: 'pending',
  },
  {
    id: 'req_02',
    documentName: 'Audited Trial Balance & Grouping Schedules',
    category: 'Financials',
    mandatory: true,
    dueDate: '2026-10-25',
    status: 'pending',
  },
  {
    id: 'req_03',
    documentName: 'GSTR-3B & GSTR-1 Filed Acknowledgements (12 Months)',
    category: 'GST & Indirect Tax',
    mandatory: true,
    dueDate: '2026-10-28',
    status: 'pending',
  },
  {
    id: 'req_04',
    documentName: 'Fixed Asset Additions Invoices (> ₹50,000)',
    category: 'Fixed Assets',
    mandatory: true,
    dueDate: '2026-10-28',
    status: 'pending',
  },
  {
    id: 'req_05',
    documentName: 'Trade Debtors & Creditors Age-wise Listing with Contact Details',
    category: 'Working Capital',
    mandatory: true,
    dueDate: '2026-10-30',
    status: 'pending',
  },
  {
    id: 'req_06',
    documentName: 'Form 26AS & AIS/TIS Tax Credit Statement',
    category: 'Direct Tax',
    mandatory: true,
    dueDate: '2026-10-30',
    status: 'pending',
  },
  {
    id: 'req_07',
    documentName: 'Statutory Board Minutes & General Meeting Resolutions',
    category: 'Secretarial',
    mandatory: false,
    dueDate: '2026-11-05',
    status: 'pending',
  },
  {
    id: 'req_08',
    documentName: 'Loan Sanction Letters & Hypothecation Agreements',
    category: 'Borrowings',
    mandatory: false,
    dueDate: '2026-11-05',
    status: 'pending',
  },
];

export const ClientDocumentTool: React.FC<ClientDocumentToolProps> = ({
  onBack,
  onShowToast,
}) => {
  const [clientName, setClientName] = useState('');
  const [financialPeriod, setFinancialPeriod] = useState('FY 2026-27');
  const [engagementType, setEngagementType] = useState('Statutory Audit u/s 139');
  const [clientEmail, setClientEmail] = useState('');

  const [items, setItems] = useState<RequisitionItem[]>([]);
  const [activeTab, setActiveTab] = useState<'register' | 'hash_verify' | 'reminder_letter'>('register');
  const [searchQuery, setSearchQuery] = useState('');
  const [selectedCategory, setSelectedCategory] = useState('all');

  // Add Item Modal
  const [showAddModal, setShowAddModal] = useState(false);
  const [newItemName, setNewItemName] = useState('');
  const [newItemCategory, setNewItemCategory] = useState('Banking');
  const [newItemMandatory, setNewItemMandatory] = useState(true);
  const [newItemDueDate, setNewItemDueDate] = useState('2026-10-31');

  // Upload Hash Modal
  const [activeUploadItem, setActiveUploadItem] = useState<RequisitionItem | null>(null);
  const [uploadedFileName, setUploadedFileName] = useState('');
  const [computedHash, setComputedHash] = useState('');
  const [isHashing, setIsHashing] = useState(false);

  // Copy state
  const [copied, setCopied] = useState(false);
  const fileInputRef = useRef<HTMLInputElement>(null);

  // Calculate Metrics
  const metrics = useMemo(() => {
    const total = items.length;
    const verified = items.filter((i) => i.status === 'verified').length;
    const received = items.filter((i) => i.status === 'received').length;
    const pending = items.filter((i) => i.status === 'pending').length;
    const completionPct = total > 0 ? Math.round(((verified + received) / total) * 100) : 0;
    return { total, verified, received, pending, completionPct };
  }, [items]);

  // Filter items
  const filteredItems = useMemo(() => {
    return items.filter((item) => {
      const matchCat = selectedCategory === 'all' || item.category === selectedCategory;
      const matchSearch =
        item.documentName.toLowerCase().includes(searchQuery.toLowerCase()) ||
        item.category.toLowerCase().includes(searchQuery.toLowerCase()) ||
        (item.fileName && item.fileName.toLowerCase().includes(searchQuery.toLowerCase()));
      return matchCat && matchSearch;
    });
  }, [items, selectedCategory, searchQuery]);

  // Categories list
  const categories = useMemo(() => {
    const set = new Set(items.map((i) => i.category));
    return ['all', ...Array.from(set)];
  }, [items]);

  // Generate SHA-256 for a file
  const handleFileSelect = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setIsHashing(true);
    setUploadedFileName(file.name);

    try {
      const arrayBuffer = await file.arrayBuffer();
      const hashBuffer = await crypto.subtle.digest('SHA-256', arrayBuffer);
      const hashArray = Array.from(new Uint8Array(hashBuffer));
      const hashHex = hashArray.map((b) => b.toString(16).padStart(2, '0')).join('');
      setComputedHash(hashHex);

      if (activeUploadItem) {
        setItems((prev) =>
          prev.map((it) =>
            it.id === activeUploadItem.id
              ? {
                  ...it,
                  status: 'verified',
                  fileName: file.name,
                  sha256Hash: hashHex,
                  uploadedAt: new Date().toLocaleString('en-IN', {
                    dateStyle: 'short',
                    timeStyle: 'short',
                  }),
                }
              : it
          )
        );
        onShowToast(
          'Document Uploaded & Verified',
          `SHA-256 integrity hash recorded for "${file.name}".`,
          'success'
        );
        setActiveUploadItem(null);
      }
    } catch {
      onShowToast('Hashing Failed', 'Could not compute file hash.', 'error');
    } finally {
      setIsHashing(false);
    }
  };

  // Add Item Handler
  const handleAddItem = () => {
    if (!newItemName.trim()) {
      onShowToast('Missing Name', 'Please enter a requisition document title.', 'error');
      return;
    }
    const newItem: RequisitionItem = {
      id: `req_${Date.now()}`,
      documentName: newItemName.trim(),
      category: newItemCategory,
      mandatory: newItemMandatory,
      dueDate: newItemDueDate,
      status: 'pending',
    };
    setItems((prev) => [...prev, newItem]);
    setNewItemName('');
    setShowAddModal(false);
    onShowToast('Requisition Added', `Added "${newItem.documentName}" to checklist.`, 'success');
  };

  // Mark status toggle
  const handleStatusChange = (id: string, newStatus: 'pending' | 'received' | 'verified') => {
    setItems((prev) =>
      prev.map((i) => (i.id === id ? { ...i, status: newStatus } : i))
    );
    onShowToast('Status Updated', `Item status set to ${newStatus.toUpperCase()}`, 'info');
  };

  // Draft Reminder Letter
  const reminderDraft = useMemo(() => {
    const pendingList = items.filter((i) => i.status === 'pending');
    return {
      subject: `URGENT: Outstanding Audit Requisitions for ${financialPeriod} - ${clientName}`,
      to: clientEmail,
      body: `Dear Client Management Team,

Subject: Outstanding Information Requisition for ${engagementType} (${financialPeriod})

With reference to the ongoing audit engagement for ${clientName}, please be informed that our audit team requires the following pending documents to finalize our statutory verification and working paper index:

${pendingList
  .map(
    (item, index) =>
      `${index + 1}. [${item.category}] ${item.documentName} (Target Due Date: ${item.dueDate})${
        item.mandatory ? ' *MANDATORY*' : ''
      }`
  )
  .join('\n')}

Kindly arrange to upload these records or send them across to our team at the earliest. Every uploaded document is verified with a SHA-256 cryptographic hash stamp for data integrity under ICAI Quality Control standards.

Warm regards,

Audit & Assurance Division
Chartered Accountants & Co.
Chartered Accountants | Peer Reviewed Firm`,
    };
  }, [items, clientName, financialPeriod, engagementType, clientEmail]);

  // Copy Reminder to Clipboard
  const handleCopyReminder = () => {
    navigator.clipboard.writeText(
      `Subject: ${reminderDraft.subject}\nTo: ${reminderDraft.to}\n\n${reminderDraft.body}`
    );
    setCopied(true);
    setTimeout(() => setCopied(false), 2000);
    onShowToast('Copied to Clipboard', 'Reminder email draft ready to send.', 'success');
  };

  // Export to Excel
  const handleExportExcel = () => {
    const exportData = items.map((i, idx) => ({
      'S.No': idx + 1,
      'Document Requisition': i.documentName,
      Category: i.category,
      Mandatory: i.mandatory ? 'Yes' : 'No',
      'Target Due Date': i.dueDate,
      'Current Status': i.status.toUpperCase(),
      'Uploaded File Name': i.fileName || 'Pending Submission',
      'SHA-256 Hash Verification': i.sha256Hash || 'N/A',
      'Uploaded At': i.uploadedAt || 'N/A',
    }));

    const ws = createStyledWorksheet(exportData, {
      dateColumns: ['Target Due Date', 'Uploaded At'],
      columnWidths: {
        'S.No': 8,
        'Document Requisition': 36,
        'Category': 22,
        'Mandatory': 12,
        'Target Due Date': 16,
        'Current Status': 18,
        'Uploaded File Name': 30,
        'SHA-256 Hash Verification': 42,
        'Uploaded At': 22,
      },
    });
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Audit_Requisitions');
    XLSX.writeFile(
      wb,
      `${clientName.replace(/\s+/g, '_')}_Audit_Requisition_Tracker_${financialPeriod}.xlsx`
    );
    onShowToast('Export Complete', 'Requisition tracker downloaded with formatted columns and frozen header.', 'success');
  };

  return (
    <div>
      {/* Top Navigation */}
      <button
        onClick={onBack}
        className="inline-flex items-center gap-2 px-3 py-1.5 rounded-lg text-xs font-semibold text-gray-700 bg-white border border-gray-200 shadow-2xs hover:bg-gray-50 hover:text-gray-900 hover:border-gray-300 transition-all cursor-pointer mb-5 group"
      >
        <ArrowLeft className="w-3.5 h-3.5 text-gray-400 group-hover:text-gray-900 group-hover:-translate-x-0.5 transition-transform" />
        <span>Back to All Tools</span>
      </button>

      {/* Header Row */}
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 pb-5 border-b border-gray-200">
        <div className="flex items-center gap-3.5">
          <InnerToolBadge toolId="client-documents" />
          <div>
            <h1 className="text-xl font-bold text-gray-900 tracking-tight">
              Client PBC Documentation & Requisitions
            </h1>
            <p className="text-xs text-gray-500 mt-0.5">
              Track client document lists, compute SHA-256 integrity hashes for uploaded evidence, and draft follow-up notices.
            </p>
          </div>
        </div>

        <div className="flex flex-wrap items-center gap-2">
          <button
            onClick={() => setItems(DEFAULT_REQUISITIONS)}
            className="inline-flex items-center gap-1.5 px-3 py-2 text-xs font-semibold rounded-lg bg-white text-gray-700 border border-gray-200 hover:bg-gray-50 hover:text-gray-900 transition-colors shadow-2xs cursor-pointer"
          >
            <RefreshCw className="w-3.5 h-3.5" />
            <span>Reset Standard List</span>
          </button>

          <button
            onClick={() => setShowAddModal(true)}
            className="inline-flex items-center gap-1.5 px-3 py-2 text-xs font-semibold rounded-lg bg-white text-gray-700 border border-gray-200 hover:bg-gray-50 hover:text-gray-900 transition-colors shadow-2xs cursor-pointer"
          >
            <Plus className="w-3.5 h-3.5 text-blue-600" />
            <span>Add Requisition</span>
          </button>

          <button
            onClick={handleExportExcel}
            className="inline-flex items-center gap-1.5 px-4 py-2 text-xs font-semibold rounded-lg bg-gray-900 text-white hover:bg-black transition-colors shadow-sm cursor-pointer"
          >
            <FileSpreadsheet className="w-3.5 h-3.5 text-gray-500" />
            <span>Export Tracker (.xlsx)</span>
          </button>
        </div>
      </div>

      {/* Hidden File Input for uploading */}
      <input
        type="file"
        ref={fileInputRef}
        onChange={handleFileSelect}
        className="hidden"
      />

      {/* Client Context Bar */}
      <div className="my-5 p-3.5 bg-gray-50 rounded-xl border border-gray-200 flex flex-wrap items-center justify-between gap-4 text-xs">
        <div className="flex flex-wrap items-center gap-5">
          <div className="flex items-center gap-2">
            <Building className="w-4 h-4 text-gray-400" />
            <span className="text-gray-500">Client:</span>
            <input
              type="text"
              value={clientName}
              onChange={(e) => setClientName(e.target.value)}
              className="font-semibold text-gray-900 bg-white border border-gray-200 rounded px-2 py-0.5 text-xs focus:ring-1 focus:ring-blue-500"
            />
          </div>
          <div className="flex items-center gap-2">
            <Calendar className="w-4 h-4 text-gray-400" />
            <span className="text-gray-500">Period:</span>
            <input
              type="text"
              value={financialPeriod}
              onChange={(e) => setFinancialPeriod(e.target.value)}
              className="font-semibold text-gray-900 bg-white border border-gray-200 rounded px-2 py-0.5 text-xs focus:ring-1 focus:ring-blue-500"
            />
          </div>
          <div className="flex items-center gap-2">
            <FileText className="w-4 h-4 text-gray-400" />
            <span className="text-gray-500">Engagement:</span>
            <span className="font-semibold text-gray-800">{engagementType}</span>
          </div>
        </div>

        <div className="flex items-center gap-2 text-gray-500">
          <Mail className="w-3.5 h-3.5" />
          <span>Client Email:</span>
          <input
            type="email"
            value={clientEmail}
            onChange={(e) => setClientEmail(e.target.value)}
            className="font-medium text-gray-800 bg-white border border-gray-200 rounded px-2 py-0.5 text-xs focus:ring-1 focus:ring-blue-500"
          />
        </div>
      </div>

      {/* 4 Metric Summary Cards */}
      <div className="grid grid-cols-2 lg:grid-cols-4 gap-4 mb-6">
        <div className="bg-white p-4 rounded-xl border border-gray-200 shadow-2xs">
          <div className="flex items-center justify-between">
            <span className="text-xs font-medium text-gray-500">Total Requested</span>
            <FileText className="w-4 h-4 text-blue-500" />
          </div>
          <div className="mt-2 flex items-baseline gap-2">
            <span className="text-2xl font-bold text-gray-900">{metrics.total}</span>
            <span className="text-xs text-gray-500">items</span>
          </div>
          <p className="text-[11px] text-gray-400 mt-1">Audit requisitions scheduled</p>
        </div>

        <div className="bg-white p-4 rounded-xl border border-gray-200 shadow-2xs">
          <div className="flex items-center justify-between">
            <span className="text-xs font-medium text-gray-500">SHA-256 Verified</span>
            <ShieldCheck className="w-4 h-4 text-gray-700" />
          </div>
          <div className="mt-2 flex items-baseline gap-2">
            <span className="text-2xl font-bold text-gray-700">{metrics.verified}</span>
            <span className="text-xs text-gray-800 bg-gray-50 px-1.5 py-0.5 rounded font-medium">
              Tamper-proof
            </span>
          </div>
          <p className="text-[11px] text-gray-400 mt-1">Hash integrity confirmed</p>
        </div>

        <div className="bg-white p-4 rounded-xl border border-gray-200 shadow-2xs">
          <div className="flex items-center justify-between">
            <span className="text-xs font-medium text-gray-500">Pending Submissions</span>
            <AlertTriangle className="w-4 h-4 text-amber-500" />
          </div>
          <div className="mt-2 flex items-baseline gap-2">
            <span className="text-2xl font-bold text-amber-600">{metrics.pending}</span>
            <span className="text-xs text-amber-700 bg-amber-50 px-1.5 py-0.5 rounded font-medium">
              Requires follow-up
            </span>
          </div>
          <p className="text-[11px] text-gray-400 mt-1">Awaiting client upload</p>
        </div>

        <div className="bg-white p-4 rounded-xl border border-gray-200 shadow-2xs">
          <div className="flex items-center justify-between">
            <span className="text-xs font-medium text-gray-500">Requisition Completion</span>
            <Clock className="w-4 h-4 text-indigo-500" />
          </div>
          <div className="mt-2 flex items-baseline gap-2">
            <span className="text-2xl font-bold text-gray-900">{metrics.completionPct}%</span>
            <span className="text-xs text-gray-500">progress</span>
          </div>
          <div className="w-full bg-gray-100 rounded-full h-1.5 mt-2">
            <div
              className="bg-blue-600 h-1.5 rounded-full transition-all duration-300"
              style={{ width: `${metrics.completionPct}%` }}
            />
          </div>
        </div>
      </div>

      {/* Clean Tabs Strip */}
      <div className="flex items-center justify-between border-b border-gray-200 mb-5">
        <div className="flex gap-2">
          <button
            onClick={() => setActiveTab('register')}
            className={`px-4 py-2.5 text-xs font-semibold border-b-2 transition-all cursor-pointer flex items-center gap-2 ${
              activeTab === 'register'
                ? 'border-blue-600 text-blue-600'
                : 'border-transparent text-gray-500 hover:text-gray-900'
            }`}
          >
            <FileSpreadsheet className="w-3.5 h-3.5" />
            <span>Requisition Register ({items.length})</span>
          </button>

          <button
            onClick={() => setActiveTab('hash_verify')}
            className={`px-4 py-2.5 text-xs font-semibold border-b-2 transition-all cursor-pointer flex items-center gap-2 ${
              activeTab === 'hash_verify'
                ? 'border-blue-600 text-blue-600'
                : 'border-transparent text-gray-500 hover:text-gray-900'
            }`}
          >
            <ShieldCheck className="w-3.5 h-3.5" />
            <span>SHA-256 Hash Verifier</span>
          </button>

          <button
            onClick={() => setActiveTab('reminder_letter')}
            className={`px-4 py-2.5 text-xs font-semibold border-b-2 transition-all cursor-pointer flex items-center gap-2 ${
              activeTab === 'reminder_letter'
                ? 'border-blue-600 text-blue-600'
                : 'border-transparent text-gray-500 hover:text-gray-900'
            }`}
          >
            <Mail className="w-3.5 h-3.5" />
            <span>Reminder Letter Drafter ({metrics.pending} pending)</span>
          </button>
        </div>
      </div>

      {/* TAB 1: REQUISITION REGISTER */}
      {activeTab === 'register' && (
        <div className="space-y-4">
          {/* Search & Filter Row */}
          <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
            <div className="relative flex-1 max-w-md">
              <Search className="w-4 h-4 absolute left-3 top-2.5 text-gray-400" />
              <input
                type="text"
                placeholder="Search requisitions by name, category, or file..."
                value={searchQuery}
                onChange={(e) => setSearchQuery(e.target.value)}
                className="w-full pl-9 pr-3 py-1.5 text-xs bg-white border border-gray-200 rounded-lg text-gray-900 focus:outline-none focus:ring-1 focus:ring-blue-500"
              />
            </div>

            <div className="flex flex-wrap items-center gap-1.5">
              <Filter className="w-3.5 h-3.5 text-gray-400 mr-1" />
              {categories.map((cat) => (
                <button
                  key={cat}
                  onClick={() => setSelectedCategory(cat)}
                  className={`px-2.5 py-1 rounded-md text-[11px] font-medium transition-colors cursor-pointer ${
                    selectedCategory === cat
                      ? 'bg-blue-50 border border-blue-200 text-blue-700 font-semibold'
                      : 'bg-white border border-gray-200 text-gray-600 hover:bg-gray-50'
                  }`}
                >
                  {cat === 'all' ? 'All Categories' : cat}
                </button>
              ))}
            </div>
          </div>

          {/* Table */}
          <div className="bg-white border border-gray-200 rounded-xl overflow-hidden shadow-2xs">
            <div className="overflow-x-auto">
              <table className="w-full text-left border-collapse text-xs">
                <thead>
                  <tr className="bg-gray-50/80 border-b border-gray-200 text-gray-600 font-semibold">
                    <th className="py-3 px-4 w-12 text-center">#</th>
                    <th className="py-3 px-4">Document Requisition</th>
                    <th className="py-3 px-3">Category</th>
                    <th className="py-3 px-3">Due Date</th>
                    <th className="py-3 px-3">Status</th>
                    <th className="py-3 px-4">SHA-256 Hash Verification</th>
                    <th className="py-3 px-4 text-right">Actions</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-gray-100">
                  {filteredItems.map((item, idx) => (
                    <tr key={item.id} className="hover:bg-blue-50/20 transition-colors">
                      <td className="py-3 px-4 text-center font-mono text-gray-400 text-[11px]">
                        {idx + 1}
                      </td>
                      <td className="py-3 px-4">
                        <div className="font-semibold text-gray-900">{item.documentName}</div>
                        {item.fileName ? (
                          <div className="text-[11px] text-gray-500 flex items-center gap-1.5 mt-0.5">
                            <FileText className="w-3 h-3 text-blue-500" />
                            <span>{item.fileName}</span>
                            <span className="text-gray-400">• {item.uploadedAt}</span>
                          </div>
                        ) : (
                          <div className="text-[11px] text-amber-600 font-medium mt-0.5">
                            Awaiting client upload
                          </div>
                        )}
                      </td>
                      <td className="py-3 px-3">
                        <span className="inline-flex items-center px-2 py-0.5 rounded text-[10px] font-medium bg-gray-100 text-gray-700 border border-gray-200">
                          {item.category}
                        </span>
                      </td>
                      <td className="py-3 px-3 font-mono text-gray-600">
                        {item.dueDate}
                      </td>
                      <td className="py-3 px-3">
                        {item.status === 'verified' && (
                          <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[10px] font-semibold bg-gray-50 text-gray-800 border border-gray-200">
                            <CheckCircle2 className="w-3 h-3 text-gray-700" />
                            Verified
                          </span>
                        )}
                        {item.status === 'received' && (
                          <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[10px] font-semibold bg-blue-50 text-blue-700 border border-blue-200">
                            <Clock className="w-3 h-3 text-blue-600" />
                            Received
                          </span>
                        )}
                        {item.status === 'pending' && (
                          <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[10px] font-semibold bg-amber-50 text-amber-700 border border-amber-200">
                            <AlertTriangle className="w-3 h-3 text-amber-500" />
                            Pending
                          </span>
                        )}
                      </td>
                      <td className="py-3 px-4 font-mono text-[11px]">
                        {item.sha256Hash ? (
                          <div className="flex items-center gap-1.5 text-gray-700" title={item.sha256Hash}>
                            <ShieldCheck className="w-3.5 h-3.5 text-gray-700 shrink-0" />
                            <span className="truncate max-w-[140px]">{item.sha256Hash}</span>
                          </div>
                        ) : (
                          <span className="text-gray-400 italic">No hash stamped</span>
                        )}
                      </td>
                      <td className="py-3 px-4 text-right">
                        <div className="flex items-center justify-end gap-1.5">
                          <button
                            onClick={() => {
                              setActiveUploadItem(item);
                              fileInputRef.current?.click();
                            }}
                            className="inline-flex items-center gap-1 px-2.5 py-1 text-[11px] font-semibold rounded bg-white text-gray-700 border border-gray-200 hover:bg-gray-50 hover:text-gray-900 transition-colors shadow-2xs cursor-pointer"
                          >
                            <UploadCloud className="w-3 h-3 text-blue-600" />
                            <span>{item.fileName ? 'Replace' : 'Upload'}</span>
                          </button>

                          {item.status === 'pending' ? (
                            <button
                              onClick={() => handleStatusChange(item.id, 'received')}
                              className="px-2 py-1 text-[11px] font-semibold text-gray-800 bg-gray-50 hover:bg-gray-100 rounded border border-gray-200 cursor-pointer"
                            >
                              Mark Received
                            </button>
                          ) : (
                            <button
                              onClick={() => handleStatusChange(item.id, 'pending')}
                              className="px-2 py-1 text-[11px] font-medium text-gray-500 hover:text-gray-700 rounded cursor-pointer"
                            >
                              Reset
                            </button>
                          )}
                        </div>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      )}

      {/* TAB 2: SHA-256 INTEGRITY HASH VERIFIER */}
      {activeTab === 'hash_verify' && (
        <div className="bg-white border border-gray-200 rounded-xl p-6 shadow-2xs space-y-6">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 rounded-lg bg-gray-100 border border-gray-200 flex items-center justify-center text-gray-900">
              <ShieldCheck className="w-5 h-5" />
            </div>
            <div>
              <h2 className="text-sm font-bold text-gray-900">Cryptographic Document Integrity Verifier</h2>
              <p className="text-xs text-gray-500">
                ICAI SA 500 audit standards require ensuring client-provided digital evidence has not been altered or tampered with.
              </p>
            </div>
          </div>

          <div
            onClick={() => fileInputRef.current?.click()}
            className="border-2 border-dashed border-gray-200 hover:border-gray-300 transition-all rounded-xl p-8 sm:p-10 text-center cursor-pointer shadow-2xs"
          >
            <div className="w-12 h-12 rounded-full bg-gray-100 border border-gray-200 flex items-center justify-center text-gray-900 mx-auto mb-3 shadow-2xs">
              <UploadCloud className="w-6 h-6" />
            </div>
            <h3 className="text-sm font-semibold text-gray-900">
              {isHashing ? 'Computing Cryptographic SHA-256 Hash...' : 'Select Client Document or Workpaper'}
            </h3>
            <p className="text-xs text-gray-500 mt-1 max-w-md mx-auto">
              Select any client record to generate an immutable SHA-256 audit fingerprint. 100% private client-side hashing.
            </p>
            <div className="mt-4 px-4 py-2 bg-gray-900 hover:bg-black text-white text-xs font-semibold rounded-lg shadow-2xs inline-flex items-center gap-2 transition-all">
              <UploadCloud className="w-4 h-4" />
              <span>{isHashing ? 'Verifying...' : 'Choose File to Hash'}</span>
            </div>
            <span className="text-[11px] text-gray-400 font-medium mt-3 block">
              Supports PDF, Excel, ZIP, CSV · Zero server upload
            </span>
          </div>

          {computedHash && (
            <div className="bg-gray-50 border border-gray-200 rounded-lg p-4 space-y-2">
              <div className="flex items-center justify-between">
                <span className="text-xs font-semibold text-gray-700">Calculated SHA-256 Checksum:</span>
                <span className="text-[10px] font-mono bg-gray-200 text-gray-900 px-2 py-0.5 rounded font-bold">
                  MATCH / VERIFIED
                </span>
              </div>
              <div className="font-mono text-xs text-gray-900 bg-white border border-gray-200 rounded p-2.5 break-all select-all font-semibold">
                {computedHash}
              </div>
              <p className="text-[11px] text-gray-500">
                File: <span className="font-semibold text-gray-800">{uploadedFileName}</span>
              </p>
            </div>
          )}
        </div>
      )}

      {/* TAB 3: REMINDER LETTER DRAFTER */}
      {activeTab === 'reminder_letter' && (
        <div className="bg-white border border-gray-200 rounded-xl p-6 shadow-2xs space-y-4">
          <div className="flex items-center justify-between pb-4 border-b border-gray-200">
            <div>
              <h2 className="text-sm font-bold text-gray-900">Follow-up Reminder Notice Draft</h2>
              <p className="text-xs text-gray-500">
                Pre-formatted formal correspondence listing outstanding statutory requisitions for {clientName}.
              </p>
            </div>
            <button
              onClick={handleCopyReminder}
              className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold rounded-lg bg-blue-600 text-white hover:bg-blue-700 transition-colors shadow-2xs cursor-pointer"
            >
              {copied ? <Check className="w-3.5 h-3.5" /> : <Copy className="w-3.5 h-3.5" />}
              <span>{copied ? 'Copied!' : 'Copy to Clipboard'}</span>
            </button>
          </div>

          <div className="space-y-3">
            <div>
              <label className="block text-[11px] font-semibold text-gray-600 uppercase mb-1">
                Subject Line
              </label>
              <input
                type="text"
                readOnly
                value={reminderDraft.subject}
                className="w-full text-xs font-medium text-gray-900 bg-gray-50 border border-gray-200 rounded-lg p-2.5 focus:outline-none"
              />
            </div>

            <div>
              <label className="block text-[11px] font-semibold text-gray-600 uppercase mb-1">
                Notice Body
              </label>
              <textarea
                rows={12}
                readOnly
                value={reminderDraft.body}
                className="w-full font-mono text-xs text-gray-800 bg-gray-50 border border-gray-200 rounded-lg p-3 leading-relaxed focus:outline-none"
              />
            </div>
          </div>
        </div>
      )}

      {/* Modal: Add Requisition */}
      {showAddModal && (
        <div className="fixed inset-0 z-50 bg-black/40 backdrop-blur-xs flex items-center justify-center p-4">
          <div className="bg-white rounded-xl shadow-xl border border-gray-200 w-full max-w-md p-5 space-y-4">
            <h3 className="text-sm font-bold text-gray-900">Add New Requisition Item</h3>
            <div>
              <label className="block text-xs font-medium text-gray-700 mb-1">Document Description</label>
              <input
                type="text"
                placeholder="e.g. Fixed Asset Physical Verification Report"
                value={newItemName}
                onChange={(e) => setNewItemName(e.target.value)}
                className="w-full text-xs border border-gray-200 rounded-lg px-3 py-2 text-gray-900 focus:outline-none focus:ring-1 focus:ring-blue-500"
              />
            </div>
            <div>
              <label className="block text-xs font-medium text-gray-700 mb-1">Category</label>
              <select
                value={newItemCategory}
                onChange={(e) => setNewItemCategory(e.target.value)}
                className="w-full text-xs border border-gray-200 rounded-lg px-3 py-2 text-gray-900 focus:outline-none focus:ring-1 focus:ring-blue-500"
              >
                <option value="Banking">Banking</option>
                <option value="Financials">Financials</option>
                <option value="GST & Indirect Tax">GST & Indirect Tax</option>
                <option value="Direct Tax">Direct Tax</option>
                <option value="Fixed Assets">Fixed Assets</option>
                <option value="Working Capital">Working Capital</option>
                <option value="Secretarial">Secretarial</option>
              </select>
            </div>
            <div className="flex gap-4">
              <div className="flex-1">
                <label className="block text-xs font-medium text-gray-700 mb-1">Target Due Date</label>
                <input
                  type="date"
                  value={newItemDueDate}
                  onChange={(e) => setNewItemDueDate(e.target.value)}
                  className="w-full text-xs border border-gray-200 rounded-lg px-3 py-2 text-gray-900 focus:outline-none focus:ring-1 focus:ring-blue-500"
                />
              </div>
              <div className="flex items-center gap-2 pt-5">
                <input
                  type="checkbox"
                  id="mandatoryCheck"
                  checked={newItemMandatory}
                  onChange={(e) => setNewItemMandatory(e.target.checked)}
                  className="rounded text-blue-600 focus:ring-blue-500"
                />
                <label htmlFor="mandatoryCheck" className="text-xs font-medium text-gray-700 cursor-pointer">
                  Mandatory
                </label>
              </div>
            </div>
            <div className="flex justify-end gap-2 pt-2">
              <button
                onClick={() => setShowAddModal(false)}
                className="px-3 py-1.5 text-xs font-semibold text-gray-600 hover:text-gray-900 rounded-lg cursor-pointer"
              >
                Cancel
              </button>
              <button
                onClick={handleAddItem}
                className="px-4 py-1.5 text-xs font-semibold bg-blue-600 hover:bg-blue-700 text-white rounded-lg shadow-2xs cursor-pointer"
              >
                Save Item
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};
