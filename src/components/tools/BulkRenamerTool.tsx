import { UploadIllustration } from '../common/UploadIllustration';
import { InnerToolBadge } from '../common/InnerToolBadge';
import React, { useState, useRef, useEffect } from 'react';
import * as XLSX from 'xlsx';
import {
  ArrowLeft,
  UploadCloud,
  Download,
  FolderArchive,
  RefreshCw,
  Sparkles,
  FileText,
  AlertTriangle,
  CheckCircle2,
  Trash2,
  Sliders,
  FileSpreadsheet,
  Edit2,
  Check,
  X,
  ListChecks,
  Copy,
  Layers,
  FolderCheck,
  ShieldCheck,
  FileCheck,
} from 'lucide-react';
import {
  type FileItemToRename,
  type RenamerConfig,
  type AuditDocCategory,
  type AuditChecklistItem,
  DEFAULT_RENAMER_CONFIG,
  extractFileMetadata,
  computeRenamedFiles,
  generateRenamedZip,
  createRenamingManifestWorkbook,
  generateSampleFilesForRenamer,
  evaluateAuditChecklist,
  getCategoryCode,
  classifyAuditDocument,
} from '../../utils/bulkRenamerEngine';
import { ToolIcon } from '../common/ToolIcon';
import { DataFormatGuide } from '../common/DataFormatGuide';

interface BulkRenamerToolProps {
  onBack: () => void;
  onShowToast: (title: string, description?: string, type?: 'success' | 'error' | 'info') => void;
}

export const BulkRenamerTool: React.FC<BulkRenamerToolProps> = ({ onBack, onShowToast }) => {
  const [files, setFiles] = useState<FileItemToRename[]>([]);
  const [config, setConfig] = useState<RenamerConfig>(DEFAULT_RENAMER_CONFIG);
  const [checklist, setChecklist] = useState<AuditChecklistItem[]>([]);
  const [activeTab, setActiveTab] = useState<'files' | 'checklist' | 'duplicates' | 'config'>('files');
  const [isProcessing, setIsProcessing] = useState(false);
  const [editingId, setEditingId] = useState<string | null>(null);
  const [editingValue, setEditingValue] = useState('');
  const [editingNotes, setEditingNotes] = useState('');
  const fileInputRef = useRef<HTMLInputElement>(null);

  // Recalculate filenames & evaluate checklist whenever files or config changes
  useEffect(() => {
    if (files.length > 0) {
      const updated = computeRenamedFiles(files, config);
      setFiles(updated);
      setChecklist(evaluateAuditChecklist(updated));
    } else {
      setChecklist(evaluateAuditChecklist([]));
    }
  }, [
    config.pattern,
    config.prefix,
    config.suffix,
    config.clientName,
    config.dateFormat,
    config.caseOption,
    config.spaceOption,
    config.seqStart,
    config.seqPadding,
    config.autoResolveDuplicates,
    config.organizeIntoFolders,
  ]);

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const uploadedFiles = e.target.files;
    if (!uploadedFiles || uploadedFiles.length === 0) return;

    const newItems: FileItemToRename[] = Array.from(uploadedFiles).map((file, idx) => {
      const ext = file.name.split('.').pop() || 'dat';
      const meta = extractFileMetadata(file.name);
      const classified = classifyAuditDocument(file.name);
      const catCode = getCategoryCode(meta.category);

      return {
        id: `file-${Date.now()}-${idx}-${Math.random().toString(36).substr(2, 4)}`,
        file,
        originalName: file.name,
        originalExtension: ext,
        sizeBytes: file.size,
        detectedDate: meta.date,
        detectedVendor: meta.vendor,
        detectedInvoiceNo: meta.invoiceNo,
        detectedCategory: meta.category,
        categoryCode: catCode,
        evidenceIndex: `WP-${catCode}-${String(files.length + idx + 1).padStart(3, '0')}`,
        auditArea: classified.auditArea,
        reviewerNotes: '',
        newName: '',
        status: 'Verified & Indexed',
        confidenceScore: classified.confidence,
      };
    });

    const combined = [...files, ...newItems];
    const computed = computeRenamedFiles(combined, config);
    setFiles(computed);
    setChecklist(evaluateAuditChecklist(computed));
    if (fileInputRef.current) fileInputRef.current.value = '';
  };

  
  const handleRemoveFile = (id: string) => {
    const updated = files.filter((f) => f.id !== id);
    const computed = computeRenamedFiles(updated, config);
    setFiles(computed);
    setChecklist(evaluateAuditChecklist(computed));
  };

  const handleClearAll = () => {
    setFiles([]);
    setChecklist(evaluateAuditChecklist([]));
    setEditingId(null);
  };

  const handleStartEdit = (item: FileItemToRename) => {
    setEditingId(item.id);
    setEditingValue(item.newName);
    setEditingNotes(item.reviewerNotes || '');
  };

  const handleSaveEdit = (id: string) => {
    const updated = files.map((f) => {
      if (f.id === id) {
        return {
          ...f,
          customOverride: editingValue,
          reviewerNotes: editingNotes,
        };
      }
      return f;
    });
    const computed = computeRenamedFiles(updated, config);
    setFiles(computed);
    setEditingId(null);
    onShowToast('Working Paper Updated', 'Saved custom override and audit notes.', 'info');
  };

  const handleDownloadZip = async () => {
    if (files.length === 0) return;
    setIsProcessing(true);
    try {
      const zipBlob = await generateRenamedZip(files, config);
      const url = URL.createObjectURL(zipBlob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `Audit_Files_${config.clientName || 'Client'}_${new Date().toISOString().split('T')[0]}.zip`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
      onShowToast('Organized Audit ZIP Downloaded', `Generated categorized folders with ${files.length} indexed files.`, 'success');
    } catch (err: any) {
      onShowToast('Error Creating ZIP', err.message || 'Failed to compress documents.', 'error');
    } finally {
      setIsProcessing(false);
    }
  };

  const handleDownloadManifest = () => {
    if (files.length === 0) return;
    try {
      const wb = createRenamingManifestWorkbook(files, checklist);
      XLSX.writeFile(wb, `Audit_Evidence_Register_${config.clientName || 'Client'}_${new Date().toISOString().split('T')[0]}.xlsx`);
      onShowToast('Evidence Register Exported', 'Downloaded 3-sheet audit working paper index and missing checklist.', 'success');
    } catch (err: any) {
      onShowToast('Export Failed', err.message || 'Could not export manifest.', 'error');
    }
  };

  const duplicateFiles = files.filter((f) => f.isDuplicate || f.status === 'Duplicate Excluded');
  const missingChecklistItems = checklist.filter((c) => !c.isFound);

  return (
    <div className="space-y-6">
      {/* 1. Header */}
      <div>
        <button
          onClick={onBack}
          className="inline-flex items-center gap-2 px-3 py-1.5 rounded-lg text-xs font-semibold text-gray-700 bg-white border border-gray-200 shadow-2xs hover:bg-gray-50 hover:text-gray-900 hover:border-gray-300 transition-all cursor-pointer mb-5 group"
        >
          <ArrowLeft className="w-3.5 h-3.5 text-gray-400 group-hover:text-gray-900 group-hover:-translate-x-0.5 transition-transform" />
          <span>Back to All Tools</span>
        </button>

        <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 pb-5 border-b border-gray-200">
          <div className="flex items-center gap-3.5">
            <InnerToolBadge toolId="bulk-renamer" />
            <div>
              <h1 className="text-xl font-bold text-gray-900 tracking-tight">Audit Working Paper (WP) Indexer</h1>
              <p className="text-xs text-gray-500 mt-0.5">
                Automatically classify audit documents, assign evidence index numbers (WP-XXX-001), detect duplicate files, and check statutory missing documents.
              </p>
            </div>
          </div>

          {files.length > 0 && (
            <div className="flex items-center gap-2">
              <button
                onClick={handleDownloadManifest}
                className="px-3.5 py-2 text-xs font-semibold text-gray-700 bg-white border border-gray-300 hover:bg-gray-50 rounded-lg shadow-sm flex items-center gap-1.5 transition-colors"
                title="Export Excel Evidence Register"
              >
                <FileSpreadsheet className="w-3.5 h-3.5 text-gray-600" />
                <span>Evidence Register (.xlsx)</span>
              </button>
              <button
                onClick={handleDownloadZip}
                disabled={isProcessing}
                className="px-4 py-2 text-xs font-semibold text-white bg-gray-900 hover:bg-black rounded-lg shadow-sm flex items-center gap-2 transition-colors cursor-pointer"
                title="Download Categorized ZIP"
              >
                <FolderArchive className="w-4 h-4" />
                <span>{isProcessing ? 'Compressing...' : 'Download Organized ZIP'}</span>
              </button>
            </div>
          )}
        </div>
      </div>

      {/* 2. Format Requirements Guide */}
      <DataFormatGuide
        toolType="bulk-renamer"
        acceptedFormats=".pdf, .xlsx, .xls, .csv, .docx, .zip, scanned images"
        requiredFields={[
          { name: 'Document Files', description: 'Bank statements, invoices, ledgers, trial balances, tax records', aliases: ['Bank Statement', 'Tax Invoice', 'Trial Balance', '26AS', 'FAR'] },
          { name: 'Working Paper Code', description: 'Auto-assigned ICAI Working Paper Code (BNK, SAL, PUR, TAX, TBS)', aliases: ['WP Code', 'Ref No', 'Evidence Index'] },
          { name: 'Audit Area', description: 'Substantive testing area (Revenue, Purchases, Cash & Bank, Taxes)', aliases: ['Audit Area', 'Section', 'FSLI'] },
        ]}
        notes={[
          'Working Paper Evidence Index: Automatically labels files like WP-BNK-001, WP-PUR-001 for seamless cross-referencing.',
          'Missing Document Audit Checklist: Checks files against 8 essential ICAI statutory audit requirements.',
          'Categorized ZIP Folders: Organizes files into 01_Bank_Statements, 02_Sales_Invoices, 07_Tax_Records upon download.',
          'Zero File Destruction: Original files are never modified or deleted; only renamed copies are archived.',
        ]}
      />

      {/* 3. Upload Area */}
      <input
        type="file"
        ref={fileInputRef}
        onChange={handleFileUpload}
        multiple
        className="hidden"
        id="bulk-renamer-upload"
      />

      {files.length === 0 ? (
        <div className="bg-white border-2 border-dashed border-gray-200 hover:border-gray-300 rounded-xl p-8 sm:p-10 text-center transition-colors shadow-2xs">
          <label htmlFor="bulk-renamer-upload" className="cursor-pointer flex flex-col items-center">
            <div className="w-12 h-12 rounded-full bg-gray-100 border border-gray-200 flex items-center justify-center text-gray-900 mx-auto mb-3 shadow-2xs">
              <UploadCloud className="w-6 h-6" />
            </div>
            <h3 className="text-sm font-semibold text-gray-900">Upload Client Audit Documents (Batch)</h3>
            <p className="text-xs text-gray-500 mt-1 max-w-md mx-auto leading-relaxed">
              Select single or batch files (.pdf, .xlsx, .csv, images). The engine will automatically classify and standardize them to ICAI audit paper naming rules.
            </p>
            <div className="mt-4 px-4 py-2 bg-gray-900 hover:bg-black text-white text-xs font-semibold rounded-lg shadow-2xs inline-flex items-center gap-2 transition-all">
              <UploadCloud className="w-4 h-4" />
              <span>Upload Documents (Batch)</span>
            </div>
            <span className="text-[11px] text-gray-400 font-medium mt-3 block">
              Multi-file upload supported · Original files are never modified or overwritten
            </span>
          </label>
        </div>
      ) : (
        <div className="bg-white rounded-xl border border-gray-200 p-5 shadow-sm">
          <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4">
            <div className="flex items-center gap-3">
              <div className="w-10 h-10 rounded-lg bg-gray-100 border border-gray-200 flex items-center justify-center text-gray-900 shrink-0">
                <FolderArchive className="w-5 h-5" />
              </div>
              <div>
                <span className="text-xs font-bold text-gray-900 block">
                  {files.length} Audit Documents Loaded
                </span>
                <span className="text-[11px] text-gray-600 block mt-0.5">
                  Ready for classification, renaming, and ICAI working paper indexing
                </span>
              </div>
            </div>
            <div className="flex items-center gap-2">
              <button
                onClick={() => fileInputRef.current?.click()}
                className="px-3.5 py-1.5 text-xs text-gray-900 bg-gray-50 hover:bg-[#EAEAEA] rounded-lg border border-gray-200 flex items-center gap-1.5 font-medium transition-colors cursor-pointer"
              >
                <UploadCloud className="w-3.5 h-3.5 text-gray-600" />
                <span>Add More Files</span>
              </button>
              <button
                onClick={handleClearAll}
                className="px-2.5 py-1.5 text-xs text-red-600 hover:text-red-700 hover:bg-red-50 rounded-lg transition-colors font-medium cursor-pointer"
              >
                Clear All
              </button>
            </div>
          </div>
        </div>
      )}

      {/* 4. Missing Documents Alert Banner if items pending */}
      {missingChecklistItems.length > 0 && files.length > 0 && (
        <div className="p-4 rounded-xl bg-amber-50 border border-amber-200 text-xs text-amber-900 flex items-center justify-between">
          <div className="flex items-center gap-2.5">
            <AlertTriangle className="w-4 h-4 text-amber-600 shrink-0" />
            <span>
              <strong>Audit Checklist Alert:</strong> {missingChecklistItems.length} mandatory audit document(s) pending from client (e.g. {missingChecklistItems.slice(0, 2).map((m) => m.title).join(', ')}).
            </span>
          </div>
          <button
            onClick={() => setActiveTab('checklist')}
            className="px-2.5 py-1 text-xs font-semibold bg-white border border-amber-300 rounded text-amber-800 hover:bg-amber-100 transition-colors"
          >
            View Checklist
          </button>
        </div>
      )}

      {/* 5. Navigation Tabs */}
      {files.length > 0 && (
        <div className="space-y-4">
          <div className="flex items-center gap-2 border-b border-gray-200 pb-1 text-xs">
            <button
              onClick={() => setActiveTab('files')}
              className={`px-3.5 py-1.5 font-semibold rounded-t-lg transition-colors flex items-center gap-1.5 ${
                activeTab === 'files'
                  ? 'bg-white text-gray-900 border-t-2 border-gray-300 shadow-xs'
                  : 'text-gray-600 hover:text-gray-900'
              }`}
            >
              <FileCheck className="w-3.5 h-3.5" />
              <span>Indexed Documents ({files.length})</span>
            </button>

            <button
              onClick={() => setActiveTab('checklist')}
              className={`px-3.5 py-1.5 font-semibold rounded-t-lg transition-colors flex items-center gap-1.5 ${
                activeTab === 'checklist'
                  ? 'bg-white text-gray-900 border-t-2 border-gray-300 shadow-xs'
                  : 'text-gray-600 hover:text-gray-900'
              }`}
            >
              <ListChecks className="w-3.5 h-3.5" />
              <span>Missing Document Checklist ({missingChecklistItems.length} pending)</span>
            </button>

            {duplicateFiles.length > 0 && (
              <button
                onClick={() => setActiveTab('duplicates')}
                className={`px-3.5 py-1.5 font-semibold rounded-t-lg transition-colors flex items-center gap-1.5 ${
                  activeTab === 'duplicates'
                    ? 'bg-white text-gray-900 border-t-2 border-gray-300 shadow-xs'
                    : 'text-gray-600 hover:text-gray-900'
                }`}
              >
                <Copy className="w-3.5 h-3.5 text-purple-600" />
                <span>Duplicate Files ({duplicateFiles.length})</span>
              </button>
            )}

            <button
              onClick={() => setActiveTab('config')}
              className={`px-3.5 py-1.5 font-semibold rounded-t-lg transition-colors flex items-center gap-1.5 ${
                activeTab === 'config'
                  ? 'bg-white text-gray-900 border-t-2 border-gray-300 shadow-xs'
                  : 'text-gray-600 hover:text-gray-900'
              }`}
            >
              <Sliders className="w-3.5 h-3.5" />
              <span>Naming Pattern & Folder Settings</span>
            </button>
          </div>

          {/* TAB 1: Indexed Documents Table */}
          {activeTab === 'files' && (
            <div className="bg-white rounded-xl border border-gray-200 overflow-hidden shadow-sm">
              <div className="px-5 py-3 bg-gray-50 border-b border-gray-200 flex items-center justify-between text-xs">
                <div className="flex items-center gap-2">
                  <span className="font-bold text-gray-900">Working Paper Evidence Register</span>
                  <span className="text-gray-600">({files.length} documents indexed)</span>
                </div>
                <span className="text-[11px] text-gray-900 font-semibold bg-slate-100 px-2 py-0.5 rounded border border-slate-200">
                  Click Edit on any row to add audit notes or customize names
                </span>
              </div>

              <div className="overflow-x-auto max-h-[500px]">
                <table className="w-full text-left text-xs border-collapse">
                  <thead className="sticky top-0 bg-gray-50 border-b border-gray-200 text-gray-600 text-[11px] font-semibold">
                    <tr>
                      <th className="py-2.5 px-3 border-r border-gray-200 w-24">Evidence Index</th>
                      <th className="py-2.5 px-3 border-r border-gray-200 whitespace-nowrap">Category</th>
                      <th className="py-2.5 px-3 border-r border-gray-200 whitespace-nowrap">Original File</th>
                      <th className="py-2.5 px-3 border-r border-gray-200 whitespace-nowrap">Standardized Renamed File</th>
                      <th className="py-2.5 px-3 border-r border-gray-200 whitespace-nowrap">Audit Area</th>
                      <th className="py-2.5 px-3 border-r border-gray-200 whitespace-nowrap">Reviewer Notes</th>
                      <th className="py-2.5 px-3 border-r border-gray-200 whitespace-nowrap">Status</th>
                      <th className="py-2.5 px-3 whitespace-nowrap text-center">Action</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-[#E5E7EB] text-gray-900">
                    {files.map((item) => {
                      const isEditing = editingId === item.id;

                      return (
                        <tr
                          key={item.id}
                          className={`hover:bg-gray-50 transition-colors ${
                            item.isDuplicate ? 'bg-purple-50/40' : item.status === 'Needs Review' ? 'bg-amber-50/40' : ''
                          }`}
                        >
                          <td className="py-2 px-3 border-r border-gray-200 font-mono font-bold text-gray-900">
                            {item.evidenceIndex}
                          </td>
                          <td className="py-2 px-3 border-r border-gray-200 whitespace-nowrap">
                            <span className="inline-block px-2 py-0.5 rounded text-[10px] font-semibold bg-gray-50 border border-gray-200 text-gray-900">
                              {item.detectedCategory}
                            </span>
                          </td>
                          <td className="py-2 px-3 border-r border-gray-200 font-mono text-[11px] text-gray-600 max-w-[200px] truncate" title={item.originalName}>
                            {item.originalName}
                          </td>
                          <td className="py-2 px-3 border-r border-gray-200 font-mono text-[11px] font-semibold text-gray-900 max-w-[240px] truncate">
                            {isEditing ? (
                              <input
                                type="text"
                                value={editingValue}
                                onChange={(e) => setEditingValue(e.target.value)}
                                className="w-full px-2 py-0.5 border border-gray-300 rounded text-xs"
                              />
                            ) : (
                              item.newName
                            )}
                          </td>
                          <td className="py-2 px-3 border-r border-gray-200 whitespace-nowrap text-[11px] text-gray-700">
                            {item.auditArea}
                          </td>
                          <td className="py-2 px-3 border-r border-gray-200 text-[11px] text-gray-600 max-w-[160px] truncate">
                            {isEditing ? (
                              <input
                                type="text"
                                placeholder="Audit notes..."
                                value={editingNotes}
                                onChange={(e) => setEditingNotes(e.target.value)}
                                className="w-full px-2 py-0.5 border border-gray-300 rounded text-xs"
                              />
                            ) : (
                              item.reviewerNotes || 'Verified'
                            )}
                          </td>
                          <td className="py-2 px-3 border-r border-gray-200 whitespace-nowrap">
                            {item.isDuplicate ? (
                              <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded text-[10px] font-semibold bg-purple-100 text-purple-800 border border-purple-200">
                                <Copy className="w-3 h-3" />
                                <span>Duplicate</span>
                              </span>
                            ) : (
                              <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded text-[10px] font-semibold bg-slate-100 text-gray-900 border border-slate-200">
                                <CheckCircle2 className="w-3 h-3" />
                                <span>Indexed</span>
                              </span>
                            )}
                          </td>
                          <td className="py-2 px-3 text-center whitespace-nowrap">
                            {isEditing ? (
                              <div className="flex items-center justify-center gap-1">
                                <button
                                  onClick={() => handleSaveEdit(item.id)}
                                  className="p-1 rounded bg-gray-900 text-white hover:bg-black"
                                  title="Save"
                                >
                                  <Check className="w-3.5 h-3.5" />
                                </button>
                                <button
                                  onClick={() => setEditingId(null)}
                                  className="p-1 rounded bg-gray-200 text-gray-700 hover:bg-gray-300"
                                  title="Cancel"
                                >
                                  <X className="w-3.5 h-3.5" />
                                </button>
                              </div>
                            ) : (
                              <div className="flex items-center justify-center gap-1">
                                <button
                                  onClick={() => handleStartEdit(item)}
                                  className="p-1 text-gray-500 hover:text-gray-900 transition-colors"
                                  title="Edit"
                                >
                                  <Edit2 className="w-3.5 h-3.5" />
                                </button>
                                <button
                                  onClick={() => handleRemoveFile(item.id)}
                                  className="p-1 text-gray-400 hover:text-red-600 transition-colors"
                                  title="Remove"
                                >
                                  <Trash2 className="w-3.5 h-3.5" />
                                </button>
                              </div>
                            )}
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            </div>
          )}

          {/* TAB 2: Missing Document Checklist */}
          {activeTab === 'checklist' && (
            <div className="bg-white rounded-xl border border-gray-200 p-5 shadow-sm space-y-4">
              <div>
                <h3 className="text-sm font-bold text-gray-900">ICAI Statutory Audit Documentation Checklist</h3>
                <p className="text-xs text-gray-600 mt-0.5">
                  Validates client document readiness against standard statutory audit compliance requirements.
                </p>
              </div>

              <div className="overflow-x-auto">
                <table className="w-full text-left text-xs border-collapse">
                  <thead className="bg-gray-50 border-b border-gray-200 text-gray-600 text-[11px] font-semibold">
                    <tr>
                      <th className="py-2.5 px-3 border-r border-gray-200 w-10 text-center">#</th>
                      <th className="py-2.5 px-3 border-r border-gray-200">Required Document</th>
                      <th className="py-2.5 px-3 border-r border-gray-200">Category</th>
                      <th className="py-2.5 px-3 border-r border-gray-200">Mandatory Audit Purpose</th>
                      <th className="py-2.5 px-3 border-r border-gray-200">Status</th>
                      <th className="py-2.5 px-3">Audit Action</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-[#E5E7EB] text-gray-900">
                    {checklist.map((chk, idx) => (
                      <tr key={chk.id} className={chk.isFound ? 'bg-gray-50/20' : 'bg-rose-50/30'}>
                        <td className="py-2.5 px-3 border-r border-gray-200 text-center text-gray-600">{idx + 1}</td>
                        <td className="py-2.5 px-3 border-r border-gray-200 font-semibold">{chk.title}</td>
                        <td className="py-2.5 px-3 border-r border-gray-200 text-gray-600">{chk.category}</td>
                        <td className="py-2.5 px-3 border-r border-gray-200 text-gray-700">{chk.requiredFor}</td>
                        <td className="py-2.5 px-3 border-r border-gray-200">
                          {chk.isFound ? (
                            <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded text-[10px] font-semibold bg-slate-100 text-gray-900 border border-slate-200">
                              <CheckCircle2 className="w-3 h-3" />
                              <span>Provided ({chk.matchedFilesCount})</span>
                            </span>
                          ) : (
                            <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded text-[10px] font-semibold bg-[#FDF2F2] text-[#C43D43] border border-[#F8D7DA]">
                              <AlertTriangle className="w-3 h-3" />
                              <span>PENDING FROM CLIENT</span>
                            </span>
                          )}
                        </td>
                        <td className="py-2.5 px-3 text-[11px] text-gray-600">
                          {chk.isFound ? 'Satisfied in audit workpapers' : 'Request urgently via Client Management Letter'}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          )}

          {/* TAB 3: Duplicates Report */}
          {activeTab === 'duplicates' && (
            <div className="bg-white rounded-xl border border-gray-200 p-5 shadow-sm space-y-4">
              <div>
                <h3 className="text-sm font-bold text-gray-900">Duplicate Files Report</h3>
                <p className="text-xs text-gray-600 mt-0.5">
                  Identical file content detected across the document batch. Excluded from evidence indexing.
                </p>
              </div>

              <div className="overflow-x-auto">
                <table className="w-full text-left text-xs border-collapse">
                  <thead className="bg-gray-50 border-b border-gray-200 text-gray-600 text-[11px] font-semibold">
                    <tr>
                      <th className="py-2.5 px-3 border-r border-gray-200">Filename</th>
                      <th className="py-2.5 px-3 border-r border-gray-200">Size</th>
                      <th className="py-2.5 px-3 border-r border-gray-200">Duplicate Of Original</th>
                      <th className="py-2.5 px-3">Audit Safeguard Action</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-[#E5E7EB]">
                    {duplicateFiles.map((d, i) => (
                      <tr key={i} className="bg-purple-50/50">
                        <td className="py-2.5 px-3 border-r border-gray-200 font-mono font-medium">{d.originalName}</td>
                        <td className="py-2.5 px-3 border-r border-gray-200 font-mono">{Math.round(d.sizeBytes / 1024)} KB</td>
                        <td className="py-2.5 px-3 border-r border-gray-200 font-mono text-purple-900">{d.duplicateOf}</td>
                        <td className="py-2.5 px-3 text-purple-900 font-semibold">Excluded to avoid duplicate audit evidence</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          )}

          {/* TAB 4: Configuration & Folder Settings */}
          {activeTab === 'config' && (
            <div className="bg-white rounded-xl border border-gray-200 p-5 shadow-sm space-y-5">
              <div>
                <h3 className="text-sm font-bold text-gray-900">Standardized Renaming Pattern & Folder Archive Settings</h3>
                <p className="text-xs text-gray-600 mt-0.5">
                  Customize naming conventions and folder organization for generated ZIP archives.
                </p>
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-3 gap-4 text-xs">
                <div>
                  <label className="block text-[11px] font-semibold text-gray-700 mb-1">Client Identifier</label>
                  <input
                    type="text"
                    value={config.clientName}
                    onChange={(e) => setConfig({ ...config, clientName: e.target.value })}
                    className="w-full px-2.5 py-1.5 bg-gray-50 border border-gray-200 rounded-lg text-xs"
                  />
                </div>

                <div>
                  <label className="block text-[11px] font-semibold text-gray-700 mb-1">Naming Pattern</label>
                  <input
                    type="text"
                    value={config.pattern}
                    onChange={(e) => setConfig({ ...config, pattern: e.target.value })}
                    className="w-full px-2.5 py-1.5 bg-gray-50 border border-gray-200 rounded-lg text-xs font-mono"
                  />
                  <span className="text-[10px] text-gray-500 block mt-1">Tokens: {'{evidenceIndex}'}, {'{client}'}, {'{category}'}, {'{date}'}, {'{vendor}'}</span>
                </div>

                <div>
                  <label className="block text-[11px] font-semibold text-gray-700 mb-1">Date Format in Filenames</label>
                  <select
                    value={config.dateFormat}
                    onChange={(e) => setConfig({ ...config, dateFormat: e.target.value as any })}
                    className="w-full px-2.5 py-1.5 bg-gray-50 border border-gray-200 rounded-lg text-xs"
                  >
                    <option value="YYYY-MM-DD">YYYY-MM-DD (2026-09-30)</option>
                    <option value="DD-MM-YYYY">DD-MM-YYYY (30-09-2026)</option>
                    <option value="YYYYMMDD">YYYYMMDD (20260930)</option>
                    <option value="DD_MMM_YYYY">DD_MMM_YYYY (30_Sep_2026)</option>
                  </select>
                </div>
              </div>

              <div className="pt-3 border-t border-gray-100 flex items-center gap-2">
                <input
                  type="checkbox"
                  id="organizeFolders"
                  checked={config.organizeIntoFolders}
                  onChange={(e) => setConfig({ ...config, organizeIntoFolders: e.target.checked })}
                  className="rounded text-gray-900 focus:ring-gray-900"
                />
                <label htmlFor="organizeFolders" className="text-xs font-semibold text-gray-800 cursor-pointer">
                  Organize into Categorized Folders (01_Bank_Statements, 02_Sales_Invoices, 07_Tax_Records) inside ZIP archive
                </label>
              </div>
            </div>
          )}
        </div>
      )}
    </div>
  );
};
