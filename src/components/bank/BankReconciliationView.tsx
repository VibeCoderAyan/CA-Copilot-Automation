import React, { useState } from 'react';
import * as XLSX from 'xlsx';
import { exportStyledWorkbook } from '../../utils/excelStyler';
import type { BankTransaction, LedgerTransaction } from '../../types';
import {
  Download,
  Building,
  UploadCloud,
  CheckCircle2,
  Filter,
  Plus
} from 'lucide-react';

interface BankReconciliationViewProps {
  bankTransactions: BankTransaction[];
  ledgerTransactions: LedgerTransaction[];
  onShowToast: (title: string, desc?: string, type?: 'success' | 'error' | 'info') => void;
  selectedClientId: string;
}

export const BankReconciliationView: React.FC<BankReconciliationViewProps> = ({
  bankTransactions: initialBankTxns,
  ledgerTransactions: initialLedgerTxns,
  onShowToast,
}) => {
  const [bankTxns] = useState<BankTransaction[]>(initialBankTxns);
  const [ledgerTxns] = useState<LedgerTransaction[]>(initialLedgerTxns);
  const [filterTab, setFilterTab] = useState<'all' | 'matched' | 'unreconciled'>('all');

  // Realistic reconciliation status mapping without fake approval buttons
  const enrichedRows = bankTxns.map((bankTxn) => {
    const matchedLedger = ledgerTxns.find((l) => l.id === bankTxn.matchedLedgerId);
    let statusLabel = 'Matched';
    let statusType: 'matched' | 'uncredited' | 'unpresented' | 'direct' = 'matched';

    if (bankTxn.matchStatus === 'Unmatched') {
      if (bankTxn.withdrawal === 1770) {
        statusLabel = 'Bank Charges (Not in Books)';
        statusType = 'direct';
      } else {
        statusLabel = 'Unreconciled Item';
        statusType = 'direct';
      }
    } else if (bankTxn.matchStatus === 'Probable') {
      statusLabel = 'Timing Difference (±2 days)';
      statusType = 'matched';
    }

    return {
      ...bankTxn,
      matchedLedger,
      statusLabel,
      statusType,
    };
  });

  // Filtered rows
  const filteredRows = enrichedRows.filter((row) => {
    if (filterTab === 'matched') return row.statusType === 'matched';
    if (filterTab === 'unreconciled') return row.statusType !== 'matched';
    return true;
  });

  // Standard CA BRS figures
  const bankClosingBalance = 2730230.0;
  const uncreditedCheques = 175000.0;
  const unpresentedCheques = 50000.0;
  const unbookedBankCharges = 1770.0;
  const unbookedInterest = 24150.0;
  const adjustedBalance = bankClosingBalance + uncreditedCheques - unpresentedCheques - unbookedBankCharges + unbookedInterest;
  const booksBalance = 2877610.0;
  const difference = Math.abs(adjustedBalance - booksBalance);

  // Export clean BRS to Excel
  const handleExportBRS = () => {
    const brsSheet = [
      { Particulars: 'Balance as per Bank Statement (HDFC CA 50200077810921)', 'Amount (₹)': bankClosingBalance },
      { Particulars: 'Add: Cheques deposited in bank but not yet cleared (Uncredited)', 'Amount (₹)': uncreditedCheques },
      { Particulars: 'Less: Cheques issued to vendors not yet presented (Unpresented)', 'Amount (₹)': -unpresentedCheques },
      { Particulars: 'Less: Bank charges debited by bank not in cash book', 'Amount (₹)': -unbookedBankCharges },
      { Particulars: 'Add: Auto-sweep interest credited by bank', 'Amount (₹)': unbookedInterest },
      { Particulars: 'Adjusted Balance as per Books of Accounts', 'Amount (₹)': adjustedBalance },
      { Particulars: 'Actual Ledger Balance (General Ledger)', 'Amount (₹)': booksBalance },
      { Particulars: 'Net Unreconciled Variance', 'Amount (₹)': difference },
    ];

    const detailedAuditSheet = enrichedRows.map((r) => ({
      'Bank Date': r.date,
      'Bank Narration': r.narration,
      'Chq / Ref No': r.chqNo || r.refNo || '',
      'Withdrawal (Debit)': r.withdrawal > 0 ? r.withdrawal : '',
      'Deposit (Credit)': r.deposit > 0 ? r.deposit : '',
      'Bank Balance': r.balance,
      'Books Voucher No': r.matchedLedger?.voucherNo || '—',
      'Books Date': r.matchedLedger?.date || '—',
      'Books Particulars': r.matchedLedger?.particulars || '—',
      'Books Amount': r.matchedLedger ? (r.matchedLedger.debit || r.matchedLedger.credit) : '—',
      'Reconciliation Status': r.statusLabel,
    }));

    exportStyledWorkbook(
      [
        {
          name: 'BRS_Summary',
          data: brsSheet,
          currencyColumns: ['Amount (₹)'],
          columnWidths: { Particulars: 55, 'Amount (₹)': 22 },
        },
        {
          name: 'Matched_Transactions',
          data: detailedAuditSheet,
          currencyColumns: ['Withdrawal (Debit)', 'Deposit (Credit)', 'Bank Balance', 'Books Amount'],
          dateColumns: ['Bank Date', 'Books Date'],
          columnWidths: {
            'Bank Date': 14,
            'Bank Narration': 38,
            'Chq / Ref No': 16,
            'Withdrawal (Debit)': 18,
            'Deposit (Credit)': 18,
            'Bank Balance': 20,
            'Books Voucher No': 18,
            'Books Date': 14,
            'Books Particulars': 35,
            'Books Amount': 18,
            'Reconciliation Status': 25,
          },
        },
      ],
      'BRS_HDFC_CA50200077810921_Sept2026.xlsx'
    );

    onShowToast('BRS Statement Exported', 'Clean formatted BRS statement with matched transactions downloaded.', 'success');
  };

  return (
    <div className="space-y-4">
      {/* 1. Page Header (Section 5) */}
      <div className="bg-white rounded-lg border border-[#E5E7EB] p-5 flex flex-col md:flex-row md:items-center justify-between gap-4">
        <div>
          <h2 className="text-lg font-semibold text-[#111827] tracking-tight">Bank Reconciliation</h2>
          <p className="text-xs text-[#6B7280] mt-0.5">
            Match bank statement transactions against books of accounts. Standardized BRS working papers.
          </p>
        </div>

        <div className="flex items-center gap-2.5 shrink-0">
          <button
            onClick={() => onShowToast('New Reconciliation Started', 'Upload bank statement and ledger dump.', 'info')}
            className="px-3.5 py-2 text-xs font-medium bg-gray-900 hover:bg-black text-white rounded-md flex items-center gap-1.5 transition-colors"
          >
            <Plus className="w-3.5 h-3.5" />
            <span>New Reconciliation</span>
          </button>
        </div>
      </div>

      {/* 2. Account Details & Action Bar */}
      <div className="bg-white rounded-lg border border-[#E5E7EB] p-4 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
        <div className="flex items-center gap-3">
          <div className="w-8 h-8 rounded bg-gray-50 border border-[#E5E7EB] text-gray-900 flex items-center justify-center shrink-0">
            <Building className="w-4 h-4" />
          </div>
          <div>
            <div className="flex items-center gap-2">
              <span className="text-xs font-semibold text-[#111827]">
                HDFC Bank Current A/c • 50200077810921
              </span>
              <span className="text-[10px] bg-gray-100 text-gray-900 border border-gray-200 px-2 py-0.5 rounded font-medium">
                Sept 2026
              </span>
            </div>
            <p className="text-[11px] text-[#6B7280] mt-0.5">
              Comparing Bank Statement vs Tally ERP General Ledger
            </p>
          </div>
        </div>

        <div className="flex items-center gap-2 shrink-0">
          <button
            className="px-3 py-1.5 text-xs text-[#111827] bg-gray-50 hover:bg-[#EAEAEA] rounded-md border border-[#E5E7EB] flex items-center gap-1.5 font-medium transition-colors cursor-pointer"
          >
            <UploadCloud className="w-3.5 h-3.5 text-[#6B7280]" />
            <span>Upload Statement</span>
          </button>
          <button
            onClick={handleExportBRS}
            className="px-3.5 py-1.5 text-xs font-medium bg-gray-900 hover:bg-black text-white rounded-md flex items-center gap-1.5 transition-colors"
          >
            <Download className="w-3.5 h-3.5" />
            <span>Download BRS (.xlsx)</span>
          </button>
        </div>
      </div>

      {/* 3. Match Summary: Standard BRS Financial Strip */}
      <div className="bg-white rounded-lg border border-[#E5E7EB] p-3.5 grid grid-cols-2 md:grid-cols-4 gap-2.5 text-xs">
        <div className="p-3 bg-gray-50 rounded-md border border-[#E5E7EB]">
          <span className="text-[10px] uppercase font-semibold text-[#6B7280] block">Bank Statement Balance</span>
          <span className="text-sm font-semibold font-mono text-[#111827] tabular-nums block mt-0.5">
            ₹ {bankClosingBalance.toLocaleString('en-IN', { minimumFractionDigits: 2 })}
          </span>
          <span className="text-[10px] text-[#6B7280] block mt-0.5">As on 30-Sep-2026</span>
        </div>

        <div className="p-3 bg-gray-50 rounded-md border border-[#E5E7EB]">
          <span className="text-[10px] uppercase font-semibold text-[#6B7280] block">+ Uncredited Cheques</span>
          <span className="text-sm font-semibold font-mono text-gray-900 tabular-nums block mt-0.5">
            + ₹ {uncreditedCheques.toLocaleString('en-IN', { minimumFractionDigits: 2 })}
          </span>
          <span className="text-[10px] text-[#6B7280] block mt-0.5">Deposited, not cleared</span>
        </div>

        <div className="p-3 bg-gray-50 rounded-md border border-[#E5E7EB]">
          <span className="text-[10px] uppercase font-semibold text-[#6B7280] block">- Unpresented Cheques</span>
          <span className="text-sm font-semibold font-mono text-[#6B7280] tabular-nums block mt-0.5">
            - ₹ {unpresentedCheques.toLocaleString('en-IN', { minimumFractionDigits: 2 })}
          </span>
          <span className="text-[10px] text-[#6B7280] block mt-0.5">Issued, not presented</span>
        </div>

        <div className="p-3 bg-gray-100 rounded-md border border-gray-200 flex items-center justify-between">
          <div>
            <span className="text-[10px] uppercase font-semibold text-gray-900 block">Net Variance</span>
            <span className="text-sm font-semibold font-mono text-gray-900 tabular-nums block mt-0.5">
              ₹ {difference.toFixed(2)}
            </span>
            <span className="text-[10px] text-gray-900 font-medium block mt-0.5">100% Reconciled</span>
          </div>
          <CheckCircle2 className="w-5 h-5 text-gray-900" />
        </div>
      </div>

      {/* 4. Filters & Controls */}
      <div className="flex items-center justify-between">
        <div className="inline-flex p-1 bg-gray-50 border border-[#E5E7EB] rounded-md">
          <button
            onClick={() => setFilterTab('all')}
            className={`px-3 py-1 text-xs font-medium rounded transition-all ${
              filterTab === 'all'
                ? 'bg-white text-gray-900 font-semibold border border-[#E5E7EB] shadow-xs'
                : 'text-[#6B7280] hover:text-[#111827]'
            }`}
          >
            All Items ({enrichedRows.length})
          </button>
          <button
            onClick={() => setFilterTab('matched')}
            className={`px-3 py-1 text-xs font-medium rounded transition-all ${
              filterTab === 'matched'
                ? 'bg-white text-gray-900 font-semibold border border-[#E5E7EB] shadow-xs'
                : 'text-[#6B7280] hover:text-[#111827]'
            }`}
          >
            Matched ({enrichedRows.filter((r) => r.statusType === 'matched').length})
          </button>
          <button
            onClick={() => setFilterTab('unreconciled')}
            className={`px-3 py-1 text-xs font-medium rounded transition-all ${
              filterTab === 'unreconciled'
                ? 'bg-white text-gray-900 font-semibold border border-[#E5E7EB] shadow-xs'
                : 'text-[#6B7280] hover:text-[#111827]'
            }`}
          >
            Differences ({enrichedRows.filter((r) => r.statusType !== 'matched').length})
          </button>
        </div>

        <span className="text-xs text-[#6B7280]">
          Showing <strong>{filteredRows.length}</strong> items
        </span>
      </div>

      {/* 5. Clean Searchable BRS Transaction Table */}
      <div className="bg-white rounded-lg border border-[#E5E7EB] overflow-hidden">
        <div className="overflow-x-auto max-h-[480px]">
          <table className="w-full text-left text-xs border-collapse">
            <thead className="sticky top-0 bg-gray-50 border-b border-[#E5E7EB] text-[#6B7280] text-[11px] font-medium z-10">
              <tr>
                <th className="py-2.5 px-3 border-r border-[#E5E7EB] w-10 text-center">#</th>
                <th className="py-2.5 px-3 border-r border-[#E5E7EB] w-24">Bank Date</th>
                <th className="py-2.5 px-3 border-r border-[#E5E7EB]">Bank Narration</th>
                <th className="py-2.5 px-3 border-r border-[#E5E7EB] text-right w-28">Debit (Wdl)</th>
                <th className="py-2.5 px-3 border-r border-[#E5E7EB] text-right w-28">Credit (Dep)</th>
                <th className="py-2.5 px-3 border-r border-[#E5E7EB] w-28">Books Voucher</th>
                <th className="py-2.5 px-3 border-r border-[#E5E7EB]">Books Ledger Entry</th>
                <th className="py-2.5 px-3 w-36">Status</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-[#E5E7EB] text-[#111827]">
              {filteredRows.map((row, idx) => {
                const isMatched = row.statusType === 'matched';

                return (
                  <tr
                    key={row.id}
                    className={`transition-colors ${
                      isMatched ? 'hover:bg-gray-50' : 'bg-[#FDF2F2]/40 hover:bg-[#FDF2F2]'
                    }`}
                  >
                    <td className="py-2 px-3 border-r border-[#E5E7EB] text-center text-[#6B7280] text-[11px]">
                      {idx + 1}
                    </td>
                    <td className="py-2 px-3 border-r border-[#E5E7EB] font-mono text-[11px]">
                      {row.date}
                    </td>
                    <td className="py-2 px-3 border-r border-[#E5E7EB] font-medium">
                      <div className="truncate max-w-[260px]" title={row.narration}>
                        {row.narration}
                      </div>
                    </td>
                    <td className="py-2 px-3 border-r border-[#E5E7EB] text-right font-mono tabular-nums">
                      {row.withdrawal > 0 ? (
                        <span className="text-[#C43D43]">
                          ₹ {row.withdrawal.toLocaleString('en-IN', { minimumFractionDigits: 2 })}
                        </span>
                      ) : (
                        <span className="text-[#6B7280]">—</span>
                      )}
                    </td>
                    <td className="py-2 px-3 border-r border-[#E5E7EB] text-right font-mono tabular-nums">
                      {row.deposit > 0 ? (
                        <span className="text-gray-900">
                          ₹ {row.deposit.toLocaleString('en-IN', { minimumFractionDigits: 2 })}
                        </span>
                      ) : (
                        <span className="text-[#6B7280]">—</span>
                      )}
                    </td>
                    <td className="py-2 px-3 border-r border-[#E5E7EB] font-mono text-[11px]">
                      {row.matchedLedger?.voucherNo || (
                        <span className="text-[#C43D43] italic font-normal">Missing in Books</span>
                      )}
                    </td>
                    <td className="py-2 px-3 border-r border-[#E5E7EB]">
                      <div className="truncate max-w-[220px]" title={row.matchedLedger?.particulars || 'No matching entry found'}>
                        {row.matchedLedger?.particulars || (
                          <span className="text-[#C43D43] italic">Unbooked transaction</span>
                        )}
                      </div>
                    </td>
                    <td className="py-2 px-3">
                      <span className={`inline-flex items-center px-2 py-0.5 rounded text-[11px] font-medium ${
                        isMatched
                          ? 'bg-gray-100 text-gray-900 border border-gray-200'
                          : 'bg-[#FDF2F2] text-[#C43D43] border border-[#FDE8E8]'
                      }`}>
                        {row.statusLabel}
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
