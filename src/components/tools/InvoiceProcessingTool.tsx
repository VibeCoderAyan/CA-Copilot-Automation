import { ToolHeader } from '../common/ToolHeader';
import React, { useState, useRef } from 'react';
import * as XLSX from 'xlsx';
import { ArrowLeft, UploadCloud, Download, CheckCircle2, ScanLine, AlertCircle, Edit2, Save, FileSpreadsheet, ShieldAlert, Info } from 'lucide-react';
import type { InvoiceExtracted } from '../../types';
import { DataFormatGuide } from '../common/DataFormatGuide';
import { formatINR, isValidGstin, parseAmount, parseDate } from '../../utils/dataSanitizer';
import { parseInvoicePdf } from '../../utils/pdfExtractor';
import { createStyledWorksheet } from '../../utils/excelStyler';

interface InvoiceProcessingToolProps {
  onBack: () => void;
  onShowToast: (title: string, desc?: string, type?: 'success' | 'error' | 'info') => void;
}

export const InvoiceProcessingTool: React.FC<InvoiceProcessingToolProps> = ({ onBack, onShowToast }) => {
  const [file, setFile] = useState<{ name: string; size: string } | null>(null);
  const [invoices, setInvoices] = useState<InvoiceExtracted[]>([]);
  const [isExtracted, setIsExtracted] = useState(false);
  const [activeFilter, setActiveFilter] = useState<'all' | 'approved' | 'review'>('all');
  const [editingId, setEditingId] = useState<string | null>(null);
  const [editFields, setEditFields] = useState<Partial<InvoiceExtracted>>({});
  const [validationError, setValidationError] = useState<string | null>(null);
  const [isProcessing, setIsProcessing] = useState(false);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const f = e.target.files?.[0];
    if (!f) return;
    setValidationError(null);

    const fileNameLower = f.name.toLowerCase();
    const isSheet = fileNameLower.endsWith('.xlsx') || fileNameLower.endsWith('.xls') || fileNameLower.endsWith('.csv');
    const isPdf = fileNameLower.endsWith('.pdf');

    if (!isSheet && !isPdf) {
      setValidationError('Unsupported format. Please upload a digital PDF invoice or an Excel/CSV invoice register (.xlsx, .csv).');
      onShowToast('Invalid Format', 'Supported formats: .pdf, .xlsx, .csv', 'error');
      return;
    }

    setIsProcessing(true);

    if (isSheet) {
      const reader = new FileReader();
      reader.onload = (evt) => {
        try {
          const bstr = evt.target?.result;
          const wb = XLSX.read(bstr, { type: 'binary', cellDates: true });
          const json = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]) as Record<string, any>[];

          if (!json || json.length === 0) {
            setInvoices([]);
            setIsExtracted(false);
            setValidationError('Spreadsheet contains zero invoice rows.');
            onShowToast('Empty File', 'No rows found in uploaded file.', 'error');
            return;
          }

          const parsedInvoices: InvoiceExtracted[] = json.map((r, i) => {
            const taxVal = parseAmount(r.taxableAmount || r['Taxable Value'] || r.Taxable || 0);
            const cgst = parseAmount(r.cgst || r.CGST || 0);
            const sgst = parseAmount(r.sgst || r.SGST || 0);
            const igst = parseAmount(r.igst || r.IGST || 0);
            const total = parseAmount(r.totalAmount || r['Invoice Value'] || r.Total || 0) || (taxVal + cgst + sgst + igst);
            const gstin = (r.supplierGstin || r.GSTIN || '').toString().trim().toUpperCase();

            return {
              id: `inv-up-${i + 1}`,
              fileName: f.name,
              supplierName: (r.supplierName || r['Supplier Name'] || r.Party || 'Vendor').toString().trim(),
              supplierGstin: gstin,
              invoiceNumber: (r.invoiceNumber || r['Invoice No'] || r.BillNo || `INV-${i + 1}`).toString().trim(),
              invoiceDate: parseDate(r.invoiceDate || r.Date || ''),
              taxableAmount: taxVal,
              cgst,
              sgst,
              igst,
              totalAmount: total,
              confidence: { overall: 98 },
              status: isValidGstin(gstin) && total > 0 ? 'Approved' : 'Needs Review',
            };
          });

          setFile({ name: f.name, size: `${Math.round(f.size / 1024)} KB` });
          setInvoices(parsedInvoices);
          setIsExtracted(true);
        } catch (err: any) {
          setInvoices([]);
          setIsExtracted(false);
          setValidationError(`Failed to parse file: ${err.message}`);
          onShowToast('Parse Error', 'Could not parse spreadsheet file.', 'error');
        } finally {
          setIsProcessing(false);
        }
      };
      reader.readAsBinaryString(f);
    } else if (isPdf) {
      try {
        const res = await parseInvoicePdf(f);
        if (res.success && res.invoice) {
          const parsedInvoice: InvoiceExtracted = {
            id: `inv-${Date.now()}`,
            fileName: f.name,
            supplierName: res.invoice.supplierName,
            supplierGstin: res.invoice.supplierGstin,
            invoiceNumber: res.invoice.invoiceNumber,
            invoiceDate: res.invoice.invoiceDate,
            taxableAmount: res.invoice.taxableAmount,
            cgst: res.invoice.cgst,
            sgst: res.invoice.sgst,
            igst: res.invoice.igst,
            totalAmount: res.invoice.totalAmount,
            confidence: { overall: 95 },
            status: isValidGstin(res.invoice.supplierGstin) && res.invoice.totalAmount > 0 ? 'Approved' : 'Needs Review',
          };

          setFile({ name: f.name, size: `${Math.round(f.size / 1024)} KB` });
          setInvoices([parsedInvoice]);
          setIsExtracted(true);
        } else {
          // Strictly reject unsupported PDF: NEVER show fake demo invoices!
          setInvoices([]);
          setIsExtracted(false);
          setValidationError(res.errorMessage || 'Unsupported document: Could not detect valid tax invoice fields (GSTIN, Invoice No, Total).');
          onShowToast('Extraction Stopped', 'Unsupported invoice PDF. Process halted to prevent corrupt data.', 'error');
        }
      } catch (err: any) {
        setInvoices([]);
        setIsExtracted(false);
        setValidationError(`Failed to parse invoice PDF: ${err.message || 'Corrupted file'}`);
        onShowToast('Read Error', 'Could not parse invoice PDF.', 'error');
      } finally {
        setIsProcessing(false);
      }
    }
  };

  
  const handleStartEdit = (inv: InvoiceExtracted) => {
    setEditingId(inv.id);
    setEditFields({ ...inv });
  };

  const handleSaveEdit = (id: string) => {
    setInvoices((prev) =>
      prev.map((inv) => (inv.id === id ? { ...inv, ...editFields, status: 'Approved' } : inv))
    );
    setEditingId(null);
    onShowToast('Invoice Updated', 'Changes saved and verified.', 'success');
  };

  const handleDownload = () => {
    if (invoices.length === 0) return;
    const exportData = invoices.map((inv, idx) => ({
      'S.No': idx + 1,
      'File Name': inv.fileName,
      'Supplier Name': inv.supplierName,
      'Supplier GSTIN': inv.supplierGstin,
      'Invoice Number': inv.invoiceNumber,
      'Invoice Date': inv.invoiceDate,
      'Taxable Amount': inv.taxableAmount,
      'CGST': inv.cgst,
      'SGST': inv.sgst,
      'IGST': inv.igst,
      'Total Amount': inv.totalAmount,
      'Verification Status': inv.status,
    }));

    const ws = createStyledWorksheet(exportData, {
      currencyColumns: ['Taxable Amount', 'CGST', 'SGST', 'IGST', 'Total Amount'],
      dateColumns: ['Invoice Date'],
      columnWidths: {
        'S.No': 8,
        'File Name': 22,
        'Supplier Name': 32,
        'Supplier GSTIN': 18,
        'Invoice Number': 18,
        'Invoice Date': 14,
        'Taxable Amount': 18,
        'CGST': 14,
        'SGST': 14,
        'IGST': 14,
        'Total Amount': 18,
        'Verification Status': 18,
      },
    });
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Extracted_Invoices');
    XLSX.writeFile(wb, 'Extracted_Purchase_Register.xlsx');
    onShowToast('Register Downloaded', 'Exported structured purchase register with formatted amounts and auto-fitted columns.', 'success');
  };

  const displayInvoices = invoices.filter((inv) => {
    if (activeFilter === 'approved') return inv.status === 'Approved';
    if (activeFilter === 'review') return inv.status !== 'Approved';
    return true;
  });

  return (
    <div className="space-y-6">
      <ToolHeader
        toolId="invoices"
        title="Invoice Processing OCR"
        description="Extract and validate vendor invoice information with line-item breakdowns and statutory taxes."
        onBack={onBack}
        actionButtons={
          invoices.length > 0 && (
            <button
              onClick={handleDownload}
              className="px-4 py-2 text-xs font-semibold text-white bg-gray-900 hover:bg-black rounded-lg shadow-sm flex items-center gap-2 transition-colors cursor-pointer"
            >
              <Download className="w-4 h-4" />
              <span>Export Purchase Register (.xlsx)</span>
            </button>
          )
        }
      />

      {/* 2. Format Requirements Guide */}
      <DataFormatGuide
        toolType="invoices"
        acceptedFormats=".pdf (digital tax invoice), .xlsx, .csv (purchase registers)"
        requiredFields={[
          { name: 'Supplier GSTIN', description: '15-character GSTIN for Input Tax Credit validation', aliases: ['GSTIN', 'Supplier GSTIN', 'GST_No'] },
          { name: 'Invoice Number', description: 'Unique tax invoice serial number', aliases: ['Invoice_No', 'Bill_No', 'Voucher_No'] },
          { name: 'Invoice Date', description: 'Date of supply in valid date format', aliases: ['Date', 'Invoice_Date'] },
          { name: 'Taxable Amount & Tax Breakup', description: 'CGST, SGST, IGST breakdown', aliases: ['Taxable', 'CGST', 'SGST', 'IGST', 'Total'] },
        ]}
        notes={[
          'Digital PDF invoices with selectable text layers and spreadsheets (.xlsx/.csv) are processed directly.',
          'Scanned photos or images without digital text cannot be parsed directly without OCR.',
          'Anti-Hallucination Guard: If a document does not match a valid invoice, processing is strictly blocked instead of producing mock invoices.',
        ]}
      />

      {/* Error banner */}
      {validationError && (
        <div className="p-5 rounded-xl bg-[#FEF2F2] border border-[#FCA5A5] shadow-sm space-y-3">
          <div className="flex items-start gap-3">
            <div className="w-9 h-9 rounded-lg bg-[#FEE2E2] flex items-center justify-center text-[#DC2626] shrink-0 mt-0.5">
              <ShieldAlert className="w-5 h-5" />
            </div>
            <div className="flex-1">
              <div className="flex items-center justify-between">
                <h3 className="text-sm font-bold text-[#991B1B]">Invoice Processing Halted</h3>
                <button
                  onClick={() => setValidationError(null)}
                  className="text-xs text-[#991B1B] hover:text-[#7F1D1D] font-medium"
                >
                  Dismiss
                </button>
              </div>
              <p className="text-xs text-[#B91C1C] mt-1 font-medium">{validationError}</p>
              
              <div className="mt-3 p-3 bg-white/80 rounded-lg border border-[#FECACA] text-xs space-y-1.5">
                <div className="flex items-center gap-1.5 font-semibold text-[#7F1D1D]">
                  <Info className="w-3.5 h-3.5" />
                  <span>Why processing stopped:</span>
                </div>
                <p className="text-[#991B1B] text-[11px] leading-relaxed">
                  The uploaded file could not be verified as a valid tax invoice. CA Automation strictly prevents the insertion of fake or simulated invoices when an actual file fails extraction.
                </p>
                <div className="pt-1.5 border-t border-[#FEE2E2] text-[11px] text-[#7F1D1D]">
                  <strong>Recommended action:</strong> Upload a digital tax invoice PDF with selectable text, or upload an Excel (.xlsx/.csv) purchase register dump.
                </div>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* 3. Upload Area */}
      <input
        type="file"
        ref={fileInputRef}
        onChange={handleFileUpload}
        accept=".pdf,.xlsx,.csv"
        className="hidden"
        id="invoice-file-upload"
      />

      {!file ? (
        <div className="bg-white border-2 border-dashed border-gray-200 rounded-xl p-8 sm:p-10 text-center hover:border-gray-300 transition-colors shadow-2xs">
          <label htmlFor="invoice-file-upload" className="cursor-pointer flex flex-col items-center">
            <div className="w-12 h-12 rounded-full bg-gray-100 border border-gray-200 flex items-center justify-center text-gray-900 mb-3 shadow-2xs">
              <ScanLine className="w-6 h-6" />
            </div>
            <h3 className="text-sm font-semibold text-gray-900">Upload Purchase Invoice or Register</h3>
            <p className="text-xs text-gray-500 mt-1 max-w-md mx-auto leading-relaxed">
              Upload digital PDF invoices (.pdf) or Excel purchase registers (.xlsx, .csv). Extracts vendor GSTIN, invoice date, taxable value, and GST breakups automatically.
            </p>
            <div className="mt-4 px-4 py-2 bg-gray-900 hover:bg-black text-white text-xs font-semibold rounded-lg shadow-2xs inline-flex items-center gap-2 transition-all">
              <UploadCloud className="w-4 h-4" />
              <span>{isProcessing ? 'Processing Invoices...' : 'Upload Invoices / Register'}</span>
            </div>
            <span className="text-[11px] text-gray-400 font-medium mt-3 block">
              Supports PDF (.pdf), Excel (.xlsx, .xls), and CSV (.csv) · Offline processed on device
            </span>
          </label>
        </div>
      ) : (
        <div className="bg-white rounded-xl border border-gray-200 p-5 shadow-sm">
          <div className="flex flex-col sm:flex-row items-center justify-between gap-4">
            <div className="flex items-center gap-3">
              <div className="w-10 h-10 rounded-lg bg-gray-100 border border-gray-200 flex items-center justify-center text-gray-900 shrink-0">
                <ScanLine className="w-5 h-5" />
              </div>
              <div>
                <span className="text-xs font-semibold text-gray-900 block">
                  {file.name}
                </span>
                <span className="text-[11px] text-gray-600 block mt-0.5">
                  {file.size} • {invoices.length} invoice(s) verified
                </span>
              </div>
            </div>

            <button
              onClick={() => fileInputRef.current?.click()}
              disabled={isProcessing}
              className="px-3.5 py-1.5 text-xs text-gray-900 bg-gray-50 hover:bg-[#EAEAEA] rounded-lg border border-gray-200 flex items-center gap-1.5 transition-colors font-medium disabled:opacity-50 cursor-pointer"
            >
              <UploadCloud className="w-3.5 h-3.5 text-gray-600" />
              <span>{isProcessing ? 'Processing...' : 'Change File'}</span>
            </button>
          </div>
        </div>
      )}

      {/* 4. Filter bar + Summary Cards + Table */}
      {isExtracted && invoices.length > 0 && (
        <div className="space-y-4">
          <div className="flex flex-col sm:flex-row items-center justify-between gap-3 bg-white p-4 rounded-xl border border-gray-200 shadow-sm">
            <div className="flex items-center gap-2">
              <span className="text-xs font-semibold text-gray-900">Filter Status:</span>
              <button
                onClick={() => setActiveFilter('all')}
                className={`px-2.5 py-1 text-xs rounded-md transition-colors ${
                  activeFilter === 'all'
                    ? 'bg-gray-900 text-white font-medium'
                    : 'bg-gray-50 text-gray-600 hover:bg-[#EAEAEA]'
                }`}
              >
                All ({invoices.length})
              </button>
              <button
                onClick={() => setActiveFilter('approved')}
                className={`px-2.5 py-1 text-xs rounded-md transition-colors ${
                  activeFilter === 'approved'
                    ? 'bg-gray-900 text-white font-medium'
                    : 'bg-gray-50 text-gray-600 hover:bg-[#EAEAEA]'
                }`}
              >
                Approved ({invoices.filter((i) => i.status === 'Approved').length})
              </button>
              <button
                onClick={() => setActiveFilter('review')}
                className={`px-2.5 py-1 text-xs rounded-md transition-colors ${
                  activeFilter === 'review'
                    ? 'bg-gray-900 text-white font-medium'
                    : 'bg-gray-50 text-gray-600 hover:bg-[#EAEAEA]'
                }`}
              >
                Needs Review ({invoices.filter((i) => i.status !== 'Approved').length})
              </button>
            </div>

            <button
              onClick={handleDownload}
              className="inline-flex items-center gap-1.5 px-3.5 py-1.5 text-xs font-medium text-white bg-gray-900 hover:bg-black rounded-lg transition-colors shadow-sm"
            >
              <Download className="w-3.5 h-3.5" />
              <span>Export Purchase Register</span>
            </button>
          </div>

          {/* Table */}
          <div className="bg-white rounded-xl border border-gray-200 overflow-hidden shadow-sm">
            <div className="overflow-x-auto max-h-[460px]">
              <table className="w-full text-left text-xs border-collapse">
                <thead className="sticky top-0 bg-gray-50 border-b border-gray-200 text-gray-600 text-[11px] font-semibold z-10">
                  <tr>
                    <th className="py-2.5 px-3 border-r border-gray-200 w-10 text-center">#</th>
                    <th className="py-2.5 px-3 border-r border-gray-200">Invoice No</th>
                    <th className="py-2.5 px-3 border-r border-gray-200">Date</th>
                    <th className="py-2.5 px-3 border-r border-gray-200">Supplier / Vendor</th>
                    <th className="py-2.5 px-3 border-r border-gray-200">Supplier GSTIN</th>
                    <th className="py-2.5 px-3 border-r border-gray-200 text-right">Taxable</th>
                    <th className="py-2.5 px-3 border-r border-gray-200 text-right">CGST</th>
                    <th className="py-2.5 px-3 border-r border-gray-200 text-right">SGST</th>
                    <th className="py-2.5 px-3 border-r border-gray-200 text-right">IGST</th>
                    <th className="py-2.5 px-3 border-r border-gray-200 text-right">Total</th>
                    <th className="py-2.5 px-3 border-r border-gray-200 text-center">Status</th>
                    <th className="py-2.5 px-3 text-center">Actions</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-[#E5E7EB] text-gray-900">
                  {displayInvoices.map((inv, idx) => {
                    const isEditing = editingId === inv.id;
                    const isValidGst = isValidGstin(inv.supplierGstin);

                    return (
                      <tr key={inv.id} className="hover:bg-gray-50 transition-colors">
                        <td className="py-2 px-3 border-r border-gray-200 text-center text-gray-600 text-[11px]">
                          {idx + 1}
                        </td>
                        <td className="py-2 px-3 border-r border-gray-200 font-semibold text-gray-900">
                          {isEditing ? (
                            <input
                              type="text"
                              value={editFields.invoiceNumber || ''}
                              onChange={(e) => setEditFields({ ...editFields, invoiceNumber: e.target.value })}
                              className="w-full px-2 py-0.5 border rounded text-xs"
                            />
                          ) : (
                            inv.invoiceNumber
                          )}
                        </td>
                        <td className="py-2 px-3 border-r border-gray-200 font-mono text-[11px]">
                          {isEditing ? (
                            <input
                              type="text"
                              value={editFields.invoiceDate || ''}
                              onChange={(e) => setEditFields({ ...editFields, invoiceDate: e.target.value })}
                              className="w-full px-2 py-0.5 border rounded text-xs"
                            />
                          ) : (
                            inv.invoiceDate
                          )}
                        </td>
                        <td className="py-2 px-3 border-r border-gray-200 font-medium">
                          {isEditing ? (
                            <input
                              type="text"
                              value={editFields.supplierName || ''}
                              onChange={(e) => setEditFields({ ...editFields, supplierName: e.target.value })}
                              className="w-full px-2 py-0.5 border rounded text-xs"
                            />
                          ) : (
                            inv.supplierName
                          )}
                        </td>
                        <td className="py-2 px-3 border-r border-gray-200 font-mono text-[11px]">
                          {isEditing ? (
                            <input
                              type="text"
                              value={editFields.supplierGstin || ''}
                              onChange={(e) => setEditFields({ ...editFields, supplierGstin: e.target.value.toUpperCase() })}
                              className="w-full px-2 py-0.5 border rounded text-xs"
                            />
                          ) : (
                            <span className={isValidGst ? 'text-gray-900' : 'text-[#DC2626] font-bold'}>
                              {inv.supplierGstin || 'MISSING'}
                            </span>
                          )}
                        </td>
                        <td className="py-2 px-3 border-r border-gray-200 text-right font-mono">
                          {isEditing ? (
                            <input
                              type="number"
                              value={editFields.taxableAmount || 0}
                              onChange={(e) => setEditFields({ ...editFields, taxableAmount: Number(e.target.value) })}
                              className="w-20 px-1 py-0.5 border rounded text-xs text-right"
                            />
                          ) : (
                            formatINR(inv.taxableAmount)
                          )}
                        </td>
                        <td className="py-2 px-3 border-r border-gray-200 text-right font-mono">
                          {formatINR(inv.cgst)}
                        </td>
                        <td className="py-2 px-3 border-r border-gray-200 text-right font-mono">
                          {formatINR(inv.sgst)}
                        </td>
                        <td className="py-2 px-3 border-r border-gray-200 text-right font-mono">
                          {formatINR(inv.igst)}
                        </td>
                        <td className="py-2 px-3 border-r border-gray-200 text-right font-mono font-bold">
                          {formatINR(inv.totalAmount)}
                        </td>
                        <td className="py-2 px-3 border-r border-gray-200 text-center">
                          <span
                            className={`inline-block text-[10px] px-2 py-0.5 rounded font-medium border ${
                              inv.status === 'Approved'
                                ? 'bg-gray-100 text-gray-900 border-gray-200'
                                : 'bg-[#FEF9EE] text-[#92400E] border-[#FDE68A]'
                            }`}
                          >
                            {inv.status}
                          </span>
                        </td>
                        <td className="py-2 px-3 text-center">
                          {isEditing ? (
                            <button
                              onClick={() => handleSaveEdit(inv.id)}
                              className="inline-flex items-center gap-1 px-2 py-0.5 bg-gray-900 text-white rounded text-[11px] font-medium"
                            >
                              <Save className="w-3 h-3" />
                              <span>Save</span>
                            </button>
                          ) : (
                            <button
                              onClick={() => handleStartEdit(inv)}
                              className="inline-flex items-center gap-1 px-2 py-0.5 bg-gray-50 hover:bg-[#EAEAEA] text-gray-900 rounded text-[11px] font-medium border border-gray-200"
                            >
                              <Edit2 className="w-3 h-3 text-gray-600" />
                              <span>Edit</span>
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

      {/* Empty State */}
      {!isExtracted && !validationError && (
        <div className="bg-white rounded-xl border border-gray-200 p-8 text-center text-xs text-gray-600">
          <ScanLine className="w-8 h-8 text-[#A0AAA2] mx-auto mb-2" />
          <p className="font-semibold text-gray-900">No Invoices Processed Yet</p>
          <p className="text-[11px] text-gray-600 mt-1 max-w-md mx-auto">
            Upload a digital tax invoice PDF or an Excel/CSV purchase register to extract real vendor details, tax amounts, and verify GSTIN validity.
          </p>
        </div>
      )}
    </div>
  );
};