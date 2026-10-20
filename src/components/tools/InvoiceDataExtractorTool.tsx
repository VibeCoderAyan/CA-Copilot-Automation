import { ToolHeader } from '../common/ToolHeader';
import React, { useState, useRef } from 'react';
import * as XLSX from 'xlsx';
import {
  ArrowLeft,
  UploadCloud,
  Download,
  CheckCircle2,
  AlertTriangle,
  Receipt,
  FileSpreadsheet,
  Edit2,
  Check,
  X,
  FileText,
  Layers,
  Sparkles,
  Info
} from 'lucide-react';
import {
  extractFromInvoicePdf,
  extractFromSpreadsheetRows,
  validateInvoiceRecord,
  createInvoiceExportWorkbook,
  type ExtractedInvoiceRecord,
} from '../../utils/invoiceDataExtractorEngine';
import { DataFormatGuide } from '../common/DataFormatGuide';
import { formatINR } from '../../utils/dataSanitizer';

interface InvoiceDataExtractorToolProps {
  onBack: () => void;
  onShowToast: (title: string, desc?: string, type?: 'success' | 'error' | 'info') => void;
}

export const InvoiceDataExtractorTool: React.FC<InvoiceDataExtractorToolProps> = ({ onBack, onShowToast }) => {
  const [invoices, setInvoices] = useState<ExtractedInvoiceRecord[]>([]);
  const [isExtracted, setIsExtracted] = useState(false);
  const [activeFilter, setActiveFilter] = useState<'all' | 'verified' | 'review'>('all');
  const [editingId, setEditingId] = useState<string | null>(null);
  const [editRow, setEditRow] = useState<Partial<ExtractedInvoiceRecord>>({});
  const [validationError, setValidationError] = useState<string | null>(null);
  const [isProcessing, setIsProcessing] = useState(false);

  const fileInputRef = useRef<HTMLInputElement>(null);

  const handleBatchUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files;
    if (!files || files.length === 0) return;
    setValidationError(null);
    setIsProcessing(true);

    try {
      const extractedList: ExtractedInvoiceRecord[] = [];

      for (let i = 0; i < files.length; i++) {
        const file = files[i];
        const lower = file.name.toLowerCase();

        if (lower.endsWith('.pdf')) {
          const rec = await extractFromInvoicePdf(file);
          extractedList.push(rec);
        } else if (lower.endsWith('.xlsx') || lower.endsWith('.xls') || lower.endsWith('.csv')) {
          const buffer = await file.arrayBuffer();
          const wb = XLSX.read(buffer, { type: 'array', cellDates: true });
          const json = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]) as Record<string, any>[];
          if (json && json.length > 0) {
            const batch = extractFromSpreadsheetRows(json, file.name);
            extractedList.push(...batch.invoices);
          }
        }
      }

      if (extractedList.length === 0) {
        throw new Error('No readable invoice documents or data rows found.');
      }

      setInvoices(extractedList);
      setIsExtracted(true);
    } catch (err: any) {
      setValidationError(`Error processing invoices: ${err.message}`);
      onShowToast('Extraction Error', err.message, 'error');
    } finally {
      setIsProcessing(false);
    }
  };

  
  const handleStartEdit = (inv: ExtractedInvoiceRecord) => {
    setEditingId(inv.id);
    setEditRow({ ...inv });
  };

  const handleSaveEdit = (id: string) => {
    setInvoices((prev) =>
      prev.map((item) => {
        if (item.id !== id) return item;

        const updatedRaw = {
          ...item,
          ...editRow,
          taxableAmount: typeof editRow.taxableAmount === 'number' ? editRow.taxableAmount : item.taxableAmount,
          cgst: typeof editRow.cgst === 'number' ? editRow.cgst : item.cgst,
          sgst: typeof editRow.sgst === 'number' ? editRow.sgst : item.sgst,
          igst: typeof editRow.igst === 'number' ? editRow.igst : item.igst,
          totalAmount: typeof editRow.totalAmount === 'number' ? editRow.totalAmount : item.totalAmount,
        };

        return validateInvoiceRecord(updatedRaw);
      })
    );
    setEditingId(null);
    setEditRow({});
    onShowToast('Invoice Updated', 'Re-calculated statutory tax math and updated status.', 'success');
  };

  const handleDownload = () => {
    if (invoices.length === 0) return;
    const wb = createInvoiceExportWorkbook(invoices, 'Batch_Invoices_FY26');
    XLSX.writeFile(wb, 'Verified_Invoice_Register_FY26.xlsx');
    onShowToast('Register Downloaded', 'Exported 3-sheet verified invoice register to Excel.', 'success');
  };

  const filteredInvoices = invoices.filter((i) => {
    if (activeFilter === 'verified') return i.status === 'Verified';
    if (activeFilter === 'review') return i.status === 'Needs Review';
    return true;
  });

  const totalTaxable = invoices.reduce((sum, i) => sum + i.taxableAmount, 0);
  const totalTax = invoices.reduce((sum, i) => sum + i.totalTax, 0);
  const totalAmount = invoices.reduce((sum, i) => sum + i.totalAmount, 0);
  const verifiedCount = invoices.filter((i) => i.status === 'Verified').length;
  const needsReviewCount = invoices.filter((i) => i.status === 'Needs Review').length;

  return (
    <div className="space-y-6">
      <ToolHeader
        toolId="invoice-extractor"
        title="Invoice Data Extractor"
        description="Batch extract invoice fields, validate statutory tax math, edit inline, and export verified registers."
        onBack={onBack}
        actionButtons={
          invoices.length > 0 && (
            <button
              onClick={handleDownload}
              className="px-4 py-2 text-xs font-semibold text-white bg-gray-900 hover:bg-black rounded-lg shadow-sm flex items-center gap-2 transition-colors cursor-pointer"
            >
              <Download className="w-4 h-4" />
              <span>Export Verified Register (.xlsx)</span>
            </button>
          )
        }
      />

      {/* 2. Format Requirements Guide */}
      <DataFormatGuide
        toolType="invoices"
        acceptedFormats=".pdf, .xlsx, .xls, .csv"
        requiredFields={[
          { name: 'Invoice Number', description: 'Tax invoice serial number', aliases: ['Invoice No', 'Bill No', 'Doc No'] },
          { name: 'Invoice Date', description: 'Date of invoice issue', aliases: ['Date', 'Invoice Date', 'Bill Date'] },
          { name: 'Supplier / Vendor', description: 'Legal name of seller', aliases: ['Supplier Name', 'Vendor', 'Party Name'] },
          { name: 'Supplier GSTIN', description: '15-digit GST identification number', aliases: ['GSTIN', 'Supplier GSTIN', 'GST No'] },
          { name: 'Taxable & Tax Breakup', description: 'Taxable value, CGST, SGST, IGST, and Total amount', aliases: ['Taxable', 'CGST', 'SGST', 'IGST', 'Total'] },
        ]}
        notes={[
          'Batch PDF Processing: Upload multiple invoice PDFs simultaneously for automated field extraction.',
          'Statutory Tax Math Check: Flags when Taxable + CGST + SGST + IGST does not equal Total Amount.',
          'Editable In-Place Table: Click on any cell in the review table to correct extracted values before exporting.',
        ]}
      />

      {/* 3. Validation Error Banner */}
      {validationError && (
        <div className="p-4 rounded-xl bg-[#FDF2F2] border border-[#F8D7DA] text-xs text-[#9B2C2C] flex items-center justify-between">
          <div className="flex items-center gap-2">
            <AlertTriangle className="w-4 h-4 shrink-0" />
            <span>{validationError}</span>
          </div>
          <button onClick={() => setValidationError(null)} className="text-xs font-semibold underline">
            Dismiss
          </button>
        </div>
      )}

      {/* 4. Batch Upload Area */}
      <input
        type="file"
        ref={fileInputRef}
        onChange={handleBatchUpload}
        accept=".pdf,.xlsx,.xls,.csv"
        multiple
        className="hidden"
        id="invoice-batch-upload"
      />

      {invoices.length === 0 ? (
        <div className="bg-white border-2 border-dashed border-gray-200 hover:border-gray-300 rounded-xl p-8 sm:p-10 text-center transition-colors shadow-2xs">
          <label htmlFor="invoice-batch-upload" className="cursor-pointer flex flex-col items-center">
            <div className="w-12 h-12 rounded-full bg-gray-100 border border-gray-200 flex items-center justify-center text-gray-900 mx-auto mb-3 shadow-2xs">
              <Receipt className="w-6 h-6" />
            </div>
            <h3 className="text-sm font-semibold text-gray-900">Upload Digital Invoices (PDF, Excel, or CSV)</h3>
            <p className="text-xs text-gray-500 mt-1 max-w-md mx-auto leading-relaxed">
              Select single or batch invoice files to automatically extract GSTINs, invoice numbers, line items, and statutory tax calculations.
            </p>
            <div className="mt-4 px-4 py-2 bg-gray-900 hover:bg-black text-white text-xs font-semibold rounded-lg shadow-2xs inline-flex items-center gap-2 transition-all">
              <UploadCloud className="w-4 h-4" />
              <span>{isProcessing ? 'Extracting Invoices...' : 'Upload Invoices (Batch)'}</span>
            </div>
            <span className="text-[11px] text-gray-400 font-medium mt-3 block">
              Supports PDF, Excel (.xlsx, .xls), and CSV (.csv) · Batch multi-file extraction supported
            </span>
          </label>
        </div>
      ) : (
        <div className="bg-white rounded-xl border border-gray-200 p-5 shadow-sm">
          <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4">
            <div className="flex items-center gap-3">
              <div className="w-10 h-10 rounded-lg bg-gray-100 border border-gray-200 flex items-center justify-center text-gray-900 shrink-0">
                <Receipt className="w-5 h-5" />
              </div>
              <div>
                <span className="text-xs font-bold text-gray-900 block">
                  {invoices.length} Invoices Loaded for Review
                </span>
                <span className="text-[11px] text-gray-600 block mt-0.5">
                  All invoice lines digitized and verified against GST statutory rules
                </span>
              </div>
            </div>
            <div className="flex items-center gap-2">
              <button
                onClick={() => fileInputRef.current?.click()}
                disabled={isProcessing}
                className="px-3.5 py-1.5 text-xs text-gray-900 bg-gray-50 hover:bg-[#EAEAEA] rounded-lg border border-gray-200 flex items-center gap-1.5 font-medium transition-colors cursor-pointer"
              >
                <UploadCloud className="w-3.5 h-3.5 text-gray-600" />
                <span>{isProcessing ? 'Extracting...' : 'Upload More Invoices'}</span>
              </button>
            </div>
          </div>
        </div>
      )}

      {/* 5. Extraction Summary & Interactive Review Workspace */}
      {isExtracted && invoices.length > 0 && (
        <div className="space-y-4">
          {/* Executive Metrics Overview */}
          <div className="grid grid-cols-2 sm:grid-cols-4 gap-4">
            <div className="bg-white rounded-xl border border-gray-200 p-4 shadow-sm">
              <span className="text-[11px] font-semibold text-gray-600 block">Total Invoices</span>
              <span className="text-xl font-bold text-gray-900 mt-1 block">
                {invoices.length}
              </span>
              <span className="text-[10px] text-gray-900 block mt-0.5">
                {verifiedCount} verified • {needsReviewCount} flagged
              </span>
            </div>

            <div className="bg-white rounded-xl border border-gray-200 p-4 shadow-sm">
              <span className="text-[11px] font-semibold text-gray-600 block">Total Taxable Value</span>
              <span className="text-xl font-bold text-gray-900 mt-1 block">
                {formatINR(totalTaxable)}
              </span>
              <span className="text-[10px] text-gray-600 block mt-0.5">Base commercial amount</span>
            </div>

            <div className="bg-white rounded-xl border border-gray-200 p-4 shadow-sm">
              <span className="text-[11px] font-semibold text-gray-600 block">Total GST Input Credit</span>
              <span className="text-xl font-bold text-gray-900 mt-1 block">
                {formatINR(totalTax)}
              </span>
              <span className="text-[10px] text-gray-600 block mt-0.5">CGST + SGST + IGST</span>
            </div>

            <div className="bg-white rounded-xl border border-gray-200 p-4 shadow-sm">
              <span className="text-[11px] font-semibold text-gray-600 block">Total Invoice Value</span>
              <span className="text-xl font-bold text-gray-900 mt-1 block">
                {formatINR(totalAmount)}
              </span>
              <span className="text-[10px] text-gray-600 block mt-0.5">Gross payable</span>
            </div>
          </div>

          {/* Editable Review Table */}
          <div className="bg-white rounded-xl border border-gray-200 overflow-hidden shadow-sm">
            <div className="px-5 py-3 bg-gray-50 border-b border-gray-200 flex flex-wrap items-center justify-between gap-3 text-xs">
              <div className="flex items-center gap-2">
                <span className="font-bold text-gray-900">Extracted Invoice Register</span>
                <span className="text-gray-600">({filteredInvoices.length} invoices shown)</span>
                <span className="text-[11px] text-gray-900 bg-gray-100 px-2 py-0.5 rounded border border-gray-200 font-medium ml-2">
                  Click Edit to modify any field
                </span>
              </div>

              {/* Filters */}
              <div className="flex items-center gap-1.5">
                <button
                  onClick={() => setActiveFilter('all')}
                  className={`px-2.5 py-1 rounded-md text-[11px] font-medium transition-colors ${
                    activeFilter === 'all'
                      ? 'bg-white text-gray-900 border border-gray-200 shadow-sm font-semibold'
                      : 'text-gray-600 hover:text-gray-900'
                  }`}
                >
                  All ({invoices.length})
                </button>
                <button
                  onClick={() => setActiveFilter('verified')}
                  className={`px-2.5 py-1 rounded-md text-[11px] font-medium transition-colors ${
                    activeFilter === 'verified'
                      ? 'bg-white text-gray-900 border border-gray-200 shadow-sm font-semibold'
                      : 'text-gray-600 hover:text-gray-900'
                  }`}
                >
                  Verified Only ({verifiedCount})
                </button>
                <button
                  onClick={() => setActiveFilter('review')}
                  className={`px-2.5 py-1 rounded-md text-[11px] font-medium transition-colors ${
                    activeFilter === 'review'
                      ? 'bg-white text-[#C43D43] border border-gray-200 shadow-sm font-semibold'
                      : 'text-gray-600 hover:text-gray-900'
                  }`}
                >
                  Needs Review ({needsReviewCount})
                </button>
              </div>
            </div>

            <div className="overflow-x-auto max-h-[500px]">
              <table className="w-full text-left text-xs border-collapse">
                <thead className="sticky top-0 bg-gray-50 border-b border-gray-200 text-gray-600 text-[11px] font-semibold">
                  <tr>
                    <th className="py-2.5 px-3 border-r border-gray-200 w-12 text-center">#</th>
                    <th className="py-2.5 px-3 border-r border-gray-200 whitespace-nowrap">File Source</th>
                    <th className="py-2.5 px-3 border-r border-gray-200 whitespace-nowrap">Invoice No</th>
                    <th className="py-2.5 px-3 border-r border-gray-200 whitespace-nowrap">Date</th>
                    <th className="py-2.5 px-3 border-r border-gray-200 whitespace-nowrap">Vendor / Supplier</th>
                    <th className="py-2.5 px-3 border-r border-gray-200 whitespace-nowrap">GSTIN</th>
                    <th className="py-2.5 px-3 border-r border-gray-200 whitespace-nowrap text-right">Taxable (₹)</th>
                    <th className="py-2.5 px-3 border-r border-gray-200 whitespace-nowrap text-right">CGST (₹)</th>
                    <th className="py-2.5 px-3 border-r border-gray-200 whitespace-nowrap text-right">SGST (₹)</th>
                    <th className="py-2.5 px-3 border-r border-gray-200 whitespace-nowrap text-right">IGST (₹)</th>
                    <th className="py-2.5 px-3 border-r border-gray-200 whitespace-nowrap text-right font-bold">Total (₹)</th>
                    <th className="py-2.5 px-3 border-r border-gray-200 whitespace-nowrap">Status</th>
                    <th className="py-2.5 px-3 whitespace-nowrap w-20 text-center">Action</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-[#E5E7EB] text-gray-900">
                  {filteredInvoices.map((inv, idx) => {
                    const isEditing = editingId === inv.id;

                    return (
                      <tr
                        key={inv.id}
                        className={`transition-colors ${
                          inv.status === 'Needs Review' ? 'bg-[#FEF9EE]' : 'hover:bg-gray-50'
                        }`}
                      >
                        <td className="py-2 px-3 border-r border-gray-200 text-center text-gray-600 text-[11px]">
                          {idx + 1}
                        </td>

                        <td className="py-2 px-3 border-r border-gray-200 font-mono text-[11px] text-gray-600 max-w-[150px] truncate" title={inv.fileName}>
                          {inv.fileName}
                        </td>

                        {/* Invoice Number */}
                        <td className="py-2 px-3 border-r border-gray-200 font-mono font-semibold">
                          {isEditing ? (
                            <input
                              type="text"
                              value={editRow.invoiceNumber ?? inv.invoiceNumber}
                              onChange={(e) => setEditRow({ ...editRow, invoiceNumber: e.target.value })}
                              className="w-24 px-1.5 py-0.5 border border-gray-300 rounded text-xs"
                            />
                          ) : (
                            inv.invoiceNumber
                          )}
                        </td>

                        {/* Date */}
                        <td className="py-2 px-3 border-r border-gray-200 whitespace-nowrap">
                          {isEditing ? (
                            <input
                              type="text"
                              value={editRow.invoiceDate ?? inv.invoiceDate}
                              onChange={(e) => setEditRow({ ...editRow, invoiceDate: e.target.value })}
                              className="w-24 px-1.5 py-0.5 border border-gray-300 rounded text-xs"
                            />
                          ) : (
                            inv.invoiceDate
                          )}
                        </td>

                        {/* Vendor Name */}
                        <td className="py-2 px-3 border-r border-gray-200 max-w-[180px] truncate" title={inv.supplierName}>
                          {isEditing ? (
                            <input
                              type="text"
                              value={editRow.supplierName ?? inv.supplierName}
                              onChange={(e) => setEditRow({ ...editRow, supplierName: e.target.value })}
                              className="w-36 px-1.5 py-0.5 border border-gray-300 rounded text-xs"
                            />
                          ) : (
                            inv.supplierName
                          )}
                        </td>

                        {/* GSTIN */}
                        <td className="py-2 px-3 border-r border-gray-200 font-mono text-[11px]">
                          {isEditing ? (
                            <input
                              type="text"
                              value={editRow.supplierGstin ?? inv.supplierGstin}
                              onChange={(e) => setEditRow({ ...editRow, supplierGstin: e.target.value })}
                              className="w-32 px-1.5 py-0.5 border border-gray-300 rounded text-xs uppercase"
                            />
                          ) : (
                            inv.supplierGstin || '-'
                          )}
                        </td>

                        {/* Taxable Amount */}
                        <td className="py-2 px-3 border-r border-gray-200 text-right font-mono">
                          {isEditing ? (
                            <input
                              type="number"
                              value={editRow.taxableAmount ?? inv.taxableAmount}
                              onChange={(e) => setEditRow({ ...editRow, taxableAmount: parseFloat(e.target.value) || 0 })}
                              className="w-20 px-1.5 py-0.5 border border-gray-300 rounded text-xs text-right"
                            />
                          ) : (
                            inv.taxableAmount.toLocaleString('en-IN', { minimumFractionDigits: 2 })
                          )}
                        </td>

                        {/* CGST */}
                        <td className="py-2 px-3 border-r border-gray-200 text-right font-mono">
                          {isEditing ? (
                            <input
                              type="number"
                              value={editRow.cgst ?? inv.cgst}
                              onChange={(e) => setEditRow({ ...editRow, cgst: parseFloat(e.target.value) || 0 })}
                              className="w-16 px-1.5 py-0.5 border border-gray-300 rounded text-xs text-right"
                            />
                          ) : (
                            inv.cgst > 0 ? inv.cgst.toLocaleString('en-IN', { minimumFractionDigits: 2 }) : '-'
                          )}
                        </td>

                        {/* SGST */}
                        <td className="py-2 px-3 border-r border-gray-200 text-right font-mono">
                          {isEditing ? (
                            <input
                              type="number"
                              value={editRow.sgst ?? inv.sgst}
                              onChange={(e) => setEditRow({ ...editRow, sgst: parseFloat(e.target.value) || 0 })}
                              className="w-16 px-1.5 py-0.5 border border-gray-300 rounded text-xs text-right"
                            />
                          ) : (
                            inv.sgst > 0 ? inv.sgst.toLocaleString('en-IN', { minimumFractionDigits: 2 }) : '-'
                          )}
                        </td>

                        {/* IGST */}
                        <td className="py-2 px-3 border-r border-gray-200 text-right font-mono">
                          {isEditing ? (
                            <input
                              type="number"
                              value={editRow.igst ?? inv.igst}
                              onChange={(e) => setEditRow({ ...editRow, igst: parseFloat(e.target.value) || 0 })}
                              className="w-16 px-1.5 py-0.5 border border-gray-300 rounded text-xs text-right"
                            />
                          ) : (
                            inv.igst > 0 ? inv.igst.toLocaleString('en-IN', { minimumFractionDigits: 2 }) : '-'
                          )}
                        </td>

                        {/* Total Amount */}
                        <td className="py-2 px-3 border-r border-gray-200 text-right font-mono font-bold text-gray-900">
                          {isEditing ? (
                            <input
                              type="number"
                              value={editRow.totalAmount ?? inv.totalAmount}
                              onChange={(e) => setEditRow({ ...editRow, totalAmount: parseFloat(e.target.value) || 0 })}
                              className="w-24 px-1.5 py-0.5 border border-gray-300 rounded text-xs text-right"
                            />
                          ) : (
                            inv.totalAmount.toLocaleString('en-IN', { minimumFractionDigits: 2 })
                          )}
                        </td>

                        {/* Status & Issues */}
                        <td className="py-2 px-3 border-r border-gray-200 whitespace-nowrap">
                          {inv.status === 'Verified' ? (
                            <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded text-[10px] font-semibold bg-gray-100 text-gray-900 border border-gray-200">
                              <CheckCircle2 className="w-3 h-3" />
                              <span>Verified</span>
                            </span>
                          ) : (
                            <div className="space-y-0.5">
                              <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded text-[10px] font-semibold bg-[#FDF2F2] text-[#C43D43] border border-[#F8D7DA]">
                                <AlertTriangle className="w-3 h-3" />
                                <span>Needs Review</span>
                              </span>
                              <span className="block text-[9px] text-[#991B1B] max-w-[140px] truncate" title={inv.validationIssues.join('; ')}>
                                {inv.validationIssues.join('; ')}
                              </span>
                            </div>
                          )}
                        </td>

                        {/* Action Edit / Save */}
                        <td className="py-2 px-3 text-center whitespace-nowrap">
                          {isEditing ? (
                            <div className="flex items-center justify-center gap-1">
                              <button
                                onClick={() => handleSaveEdit(inv.id)}
                                className="p-1 rounded bg-gray-900 text-white hover:bg-black transition-colors"
                                title="Save & Recalculate"
                              >
                                <Check className="w-3 h-3" />
                              </button>
                              <button
                                onClick={() => setEditingId(null)}
                                className="p-1 rounded bg-[#E5E7EB] text-gray-900 hover:bg-gray-200 transition-colors"
                                title="Cancel"
                              >
                                <X className="w-3 h-3" />
                              </button>
                            </div>
                          ) : (
                            <button
                              onClick={() => handleStartEdit(inv)}
                              className="p-1 text-gray-600 hover:text-gray-900 hover:bg-gray-100 rounded transition-colors"
                              title="Edit Row"
                            >
                              <Edit2 className="w-3.5 h-3.5" />
                            </button>
                          )}
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};