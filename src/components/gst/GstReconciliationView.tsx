import React, { useState } from 'react';
import * as XLSX from 'xlsx';
import { exportStyledWorkbook } from '../../utils/excelStyler';
import type { GstRecord } from '../../types';
import {
  Download,
  ReceiptText,
  UploadCloud,
  Search,
  CheckCircle2,
  AlertTriangle,
  Play
} from 'lucide-react';

interface GstReconciliationViewProps {
  gstRecords: GstRecord[];
  onShowToast: (title: string, desc?: string, type?: 'success' | 'error' | 'info') => void;
  selectedClientId: string;
}

export const GstReconciliationView: React.FC<GstReconciliationViewProps> = ({
  gstRecords: initialRecords,
  onShowToast,
}) => {
  const [records] = useState<GstRecord[]>(initialRecords);
  const [filterStatus, setFilterStatus] = useState<string>('all');
  const [searchQuery, setSearchQuery] = useState('');

  // Filter records
  const filteredRecords = records.filter((rec) => {
    const matchesStatus =
      filterStatus === 'all'
        ? true
        : filterStatus === 'matched'
        ? rec.status === 'Matched'
        : filterStatus === 'mismatch'
        ? rec.status === 'Tax Mismatch'
        : filterStatus === 'missing_2b'
        ? rec.status === 'Missing in 2B'
        : filterStatus === 'missing_books'
        ? rec.status === 'Missing in Books'
        : true;

    const matchesSearch =
      rec.supplierName.toLowerCase().includes(searchQuery.toLowerCase()) ||
      rec.supplierGstin.toLowerCase().includes(searchQuery.toLowerCase()) ||
      rec.invoiceNumber.toLowerCase().includes(searchQuery.toLowerCase());

    return matchesStatus && matchesSearch;
  });

  // Calculate ITC totals
  const totalBooksTax = records.reduce((acc, r) => acc + r.booksTotalTax, 0);
  const total2bTax = records.reduce((acc, r) => acc + r.gstr2bTotalTax, 0);
  const eligibleItc = records
    .filter((r) => r.itcEligibility === 'Eligible')
    .reduce((acc, r) => acc + r.booksTotalTax, 0);
  const atRiskItc = records
    .filter((r) => r.status === 'Missing in 2B' || r.status === 'Tax Mismatch')
    .reduce((acc, r) => acc + r.diffAmount, 0);

  // Export GST Reconciliation Report to XLSX
  const handleExportGstReport = () => {
    const exportData = records.map((r, idx) => ({
      'S.No': idx + 1,
      'Supplier Name': r.supplierName,
      'Supplier GSTIN': r.supplierGstin,
      'Invoice Number': r.invoiceNumber,
      'Invoice Date': r.invoiceDate,
      'Books Taxable (₹)': r.booksTaxable,
      'Books Total Tax (₹)': r.booksTotalTax,
      'GSTR-2B Taxable (₹)': r.gstr2bTaxable,
      'GSTR-2B Total Tax (₹)': r.gstr2bTotalTax,
      'Tax Variance (₹)': r.diffAmount,
      'Audit Status': r.status,
      'ITC Claimable': r.status === 'Matched' ? 'Eligible' : 'Blocked / Pending',
    }));

    exportStyledWorkbook(
      [{ name: 'GSTR2B_vs_Books', data: exportData }],
      'GSTR2B_Reconciliation_Sept2026.xlsx'
    );

    onShowToast('GST Working Paper Exported', 'Clean GSTR-2B vs Purchase Register spreadsheet downloaded with auto-fitted columns.', 'success');
  };

  return (
    <div className="space-y-4">
      {/* 1. Page Header (Section 5) */}
      <div className="bg-white rounded-lg border border-[#E5E7EB] p-5 flex flex-col md:flex-row md:items-center justify-between gap-4">
        <div>
          <h2 className="text-lg font-semibold text-[#111827] tracking-tight">GST Reconciliation</h2>
          <p className="text-xs text-[#6B7280] mt-0.5">
            Compare purchase register vouchers against government GSTR-2B auto-drafted ITC data.
          </p>
        </div>

        <div className="flex items-center gap-2.5 shrink-0">
          <button
            onClick={() => onShowToast('Reconciliation Started', 'Comparing Purchase Register with GSTR-2B portal extract.', 'info')}
            className="px-3.5 py-2 text-xs font-medium bg-gray-900 hover:bg-black text-white rounded-md flex items-center gap-1.5 transition-colors"
          >
            <Play className="w-3.5 h-3.5" />
            <span>Run Comparison</span>
          </button>
        </div>
      </div>

      {/* 2. File Upload & Comparison Header */}
      <div className="bg-white rounded-lg border border-[#E5E7EB] p-4 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
        <div className="flex items-center gap-3">
          <div className="w-8 h-8 rounded bg-gray-50 border border-[#E5E7EB] text-gray-900 flex items-center justify-center shrink-0">
            <ReceiptText className="w-4 h-4" />
          </div>
          <div>
            <div className="flex items-center gap-2">
              <span className="text-xs font-semibold text-[#111827]">
                Purchase Register vs GSTR-2B
              </span>
              <span className="text-[10px] bg-gray-100 text-gray-900 border border-gray-200 px-2 py-0.5 rounded font-medium">
                Sept 2026
              </span>
            </div>
            <p className="text-[11px] text-[#6B7280] mt-0.5">
              Source: GSTR-2B JSON / Excel dump vs Books Purchase Daybook
            </p>
          </div>
        </div>

        <div className="flex items-center gap-2 shrink-0">
          <button
            className="px-3 py-1.5 text-xs text-[#111827] bg-gray-50 hover:bg-[#EAEAEA] rounded-md border border-[#E5E7EB] flex items-center gap-1.5 font-medium transition-colors cursor-pointer"
          >
            <UploadCloud className="w-3.5 h-3.5 text-[#6B7280]" />
            <span>Upload Books</span>
          </button>
          <button
            onClick={handleExportGstReport}
            className="px-3.5 py-1.5 text-xs font-medium bg-gray-900 hover:bg-black text-white rounded-md flex items-center gap-1.5 transition-colors"
          >
            <Download className="w-3.5 h-3.5" />
            <span>Download Working Paper (.xlsx)</span>
          </button>
        </div>
      </div>

      {/* 3. Summary Cards: ITC & Discrepancies */}
      <div className="bg-white rounded-lg border border-[#E5E7EB] p-3.5 grid grid-cols-2 md:grid-cols-4 gap-2.5 text-xs">
        <div className="p-3 bg-gray-50 rounded-md border border-[#E5E7EB]">
          <span className="text-[10px] uppercase font-semibold text-[#6B7280] block">Books ITC (Total Tax)</span>
          <span className="text-sm font-semibold font-mono text-[#111827] tabular-nums block mt-0.5">
            ₹ {totalBooksTax.toLocaleString('en-IN', { minimumFractionDigits: 2 })}
          </span>
          <span className="text-[10px] text-[#6B7280] block mt-0.5">Entered in accounting ledgers</span>
        </div>

        <div className="p-3 bg-gray-50 rounded-md border border-[#E5E7EB]">
          <span className="text-[10px] uppercase font-semibold text-[#6B7280] block">GSTR-2B ITC (Portal)</span>
          <span className="text-sm font-semibold font-mono text-gray-900 tabular-nums block mt-0.5">
            ₹ {total2bTax.toLocaleString('en-IN', { minimumFractionDigits: 2 })}
          </span>
          <span className="text-[10px] text-[#6B7280] block mt-0.5">Auto-drafted by suppliers</span>
        </div>

        <div className="p-3 bg-gray-100 rounded-md border border-gray-200">
          <span className="text-[10px] uppercase font-semibold text-gray-900 block">Eligible ITC (Claimable)</span>
          <span className="text-sm font-semibold font-mono text-gray-900 tabular-nums block mt-0.5">
            ₹ {eligibleItc.toLocaleString('en-IN', { minimumFractionDigits: 2 })}
          </span>
          <span className="text-[10px] text-gray-900 font-medium block mt-0.5">Matched with zero dispute</span>
        </div>

        <div className="p-3 bg-[#FDF2F2] rounded-md border border-[#FDE8E8]">
          <span className="text-[10px] uppercase font-semibold text-[#C43D43] block">ITC Variance (At Risk)</span>
          <span className="text-sm font-semibold font-mono text-[#C43D43] tabular-nums block mt-0.5">
            ₹ {atRiskItc.toLocaleString('en-IN', { minimumFractionDigits: 2 })}
          </span>
          <span className="text-[10px] text-[#C43D43] font-medium block mt-0.5">Actionable discrepancies</span>
        </div>
      </div>

      {/* 4. Search and Filter Bar */}
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2.5">
        <div className="inline-flex p-1 bg-gray-50 border border-[#E5E7EB] rounded-md">
          <button
            onClick={() => setFilterStatus('all')}
            className={`px-3 py-1 text-xs font-medium rounded transition-all ${
              filterStatus === 'all'
                ? 'bg-white text-gray-900 font-semibold border border-[#E5E7EB] shadow-xs'
                : 'text-[#6B7280] hover:text-[#111827]'
            }`}
          >
            All Invoices ({records.length})
          </button>
          <button
            onClick={() => setFilterStatus('matched')}
            className={`px-3 py-1 text-xs font-medium rounded transition-all ${
              filterStatus === 'matched'
                ? 'bg-white text-gray-900 font-semibold border border-[#E5E7EB] shadow-xs'
                : 'text-[#6B7280] hover:text-[#111827]'
            }`}
          >
            Matched ({records.filter((r) => r.status === 'Matched').length})
          </button>
          <button
            onClick={() => setFilterStatus('mismatch')}
            className={`px-3 py-1 text-xs font-medium rounded transition-all ${
              filterStatus === 'mismatch'
                ? 'bg-white text-gray-900 font-semibold border border-[#E5E7EB] shadow-xs'
                : 'text-[#6B7280] hover:text-[#111827]'
            }`}
          >
            Tax Mismatch ({records.filter((r) => r.status === 'Tax Mismatch').length})
          </button>
          <button
            onClick={() => setFilterStatus('missing_2b')}
            className={`px-3 py-1 text-xs font-medium rounded transition-all ${
              filterStatus === 'missing_2b'
                ? 'bg-white text-gray-900 font-semibold border border-[#E5E7EB] shadow-xs'
                : 'text-[#6B7280] hover:text-[#111827]'
            }`}
          >
            Missing in 2B ({records.filter((r) => r.status === 'Missing in 2B').length})
          </button>
        </div>

        <div className="relative">
          <Search className="w-3.5 h-3.5 text-[#6B7280] absolute left-3 top-1/2 -translate-y-1/2" />
          <input
            type="text"
            value={searchQuery}
            onChange={(e) => setSearchQuery(e.target.value)}
            placeholder="Search supplier, GSTIN, invoice..."
            className="pl-8 pr-3 py-1.5 bg-white border border-[#E5E7EB] rounded-md text-xs text-[#111827] placeholder-[#6B7280] focus:outline-none focus:border-gray-900 w-full sm:w-64"
          />
        </div>
      </div>

      {/* 5. Clean GST Reconciliation Table */}
      <div className="bg-white rounded-lg border border-[#E5E7EB] overflow-hidden">
        <div className="overflow-x-auto max-h-[480px]">
          <table className="w-full text-left text-xs border-collapse">
            <thead className="sticky top-0 bg-gray-50 border-b border-[#E5E7EB] text-[#6B7280] text-[11px] font-medium z-10">
              <tr>
                <th className="py-2.5 px-3 border-r border-[#E5E7EB] w-10 text-center">#</th>
                <th className="py-2.5 px-3 border-r border-[#E5E7EB]">Supplier Name</th>
                <th className="py-2.5 px-3 border-r border-[#E5E7EB]">Supplier GSTIN</th>
                <th className="py-2.5 px-3 border-r border-[#E5E7EB]">Invoice No & Date</th>
                <th className="py-2.5 px-3 border-r border-[#E5E7EB] text-right">Books Tax</th>
                <th className="py-2.5 px-3 border-r border-[#E5E7EB] text-right">GSTR-2B Tax</th>
                <th className="py-2.5 px-3 border-r border-[#E5E7EB] text-right">Variance</th>
                <th className="py-2.5 px-3">Status</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-[#E5E7EB] text-[#111827]">
              {filteredRecords.map((r, idx) => {
                const isMatched = r.status === 'Matched';
                const isDiff = r.diffAmount > 0;

                return (
                  <tr
                    key={r.id}
                    className={`transition-colors ${
                      isMatched
                        ? 'hover:bg-gray-50'
                        : 'bg-[#FDF2F2]/40 hover:bg-[#FDF2F2]'
                    }`}
                  >
                    <td className="py-2 px-3 border-r border-[#E5E7EB] text-center text-[#6B7280] text-[11px]">
                      {idx + 1}
                    </td>
                    <td className="py-2 px-3 border-r border-[#E5E7EB] font-medium">
                      {r.supplierName}
                    </td>
                    <td className="py-2 px-3 border-r border-[#E5E7EB] font-mono text-[11px]">
                      {r.supplierGstin}
                    </td>
                    <td className="py-2 px-3 border-r border-[#E5E7EB] font-mono text-[11px]">
                      <div>{r.invoiceNumber}</div>
                      <div className="text-[10px] text-[#6B7280]">{r.invoiceDate}</div>
                    </td>
                    <td className="py-2 px-3 border-r border-[#E5E7EB] text-right font-mono tabular-nums">
                      ₹ {r.booksTotalTax.toLocaleString('en-IN', { minimumFractionDigits: 2 })}
                    </td>
                    <td className="py-2 px-3 border-r border-[#E5E7EB] text-right font-mono tabular-nums">
                      {r.gstr2bTotalTax > 0 ? (
                        <span>₹ {r.gstr2bTotalTax.toLocaleString('en-IN', { minimumFractionDigits: 2 })}</span>
                      ) : (
                        <span className="text-[#C43D43] italic font-normal">Not in 2B</span>
                      )}
                    </td>
                    <td className="py-2 px-3 border-r border-[#E5E7EB] text-right font-mono tabular-nums font-semibold">
                      {isDiff ? (
                        <span className="text-[#C43D43]">
                          ₹ {r.diffAmount.toLocaleString('en-IN', { minimumFractionDigits: 2 })}
                        </span>
                      ) : (
                        <span className="text-gray-900">₹ 0.00</span>
                      )}
                    </td>
                    <td className="py-2 px-3">
                      <span className={`inline-flex items-center px-2 py-0.5 rounded text-[11px] font-medium ${
                        isMatched
                          ? 'bg-gray-100 text-gray-900 border border-gray-200'
                          : 'bg-[#FDF2F2] text-[#C43D43] border border-[#FDE8E8]'
                      }`}>
                        {r.status}
                      </span>
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
};
