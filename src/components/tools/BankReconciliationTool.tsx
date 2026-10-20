import React, { useState } from 'react';
import * as XLSX from 'xlsx';
import { ArrowLeft, Download, CheckCircle2, AlertCircle, RefreshCw, UploadCloud, BookOpen, FileCode } from 'lucide-react';
import {
  parseBankStatementRows,
  parseLedgerRows,
  runBankReconciliation,
  createBrsWorkbook,
  type ParsedBankItem,
  type ParsedLedgerItem,
  type BrsCalculation,
} from '../../utils/bankReconciliationEngine';
import { DataFormatGuide } from '../common/DataFormatGuide';
import { formatINR } from '../../utils/dataSanitizer';
import { ToolHeader } from '../common/ToolHeader';
import {
  type TallyVoucherItem,
  generateTallyXml,
  downloadTallyXml,
  classifyBankNarration,
} from '../../utils/tallyXmlGenerator';
import { getActiveClient } from '../../utils/clientStorage';

interface BankReconciliationToolProps {
  onBack: () => void;
  onShowToast: (title: string, desc?: string, type?: 'success' | 'error' | 'info') => void;
}




export const BankReconciliationTool: React.FC<BankReconciliationToolProps> = ({ onBack, onShowToast }) => {
  const [bankFile, setBankFile] = useState<{ name: string; size: string } | null>(null);
  const [ledgerFile, setLedgerFile] = useState<{ name: string; size: string } | null>(null);

  const [bankItems, setBankItems] = useState<ParsedBankItem[]>([]);
  const [ledgerItems, setLedgerItems] = useState<ParsedLedgerItem[]>([]);

  const [reconciledBank, setReconciledBank] = useState<ParsedBankItem[]>([]);
  const [reconciledLedger, setReconciledLedger] = useState<ParsedLedgerItem[]>([]);
  const [calculation, setCalculation] = useState<BrsCalculation | null>(null);

  const [isReconciled, setIsReconciled] = useState(false);
  const [activeTab, setActiveTab] = useState<'all' | 'matched' | 'unrecBank' | 'unrecLedger'>('all');
  const [validationError, setValidationError] = useState<string | null>(null);

  const handleBankUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const f = e.target.files?.[0];
    if (!f) return;
    setValidationError(null);

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const bstr = evt.target?.result;
        const wb = XLSX.read(bstr, { type: 'binary', cellDates: true });
        const json = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]) as Record<string, any>[];

        if (!json || json.length === 0) {
          setValidationError(`Bank statement file "${f.name}" contains zero transaction rows.`);
          onShowToast('Empty Statement', 'No rows found in uploaded bank file.', 'error');
          return;
        }

        const parsed = parseBankStatementRows(json);
        setBankFile({ name: f.name, size: `${Math.round(f.size / 1024)} KB` });
        setBankItems(parsed);
        setIsReconciled(false);
      } catch (err: any) {
        setValidationError(`Failed to parse bank statement: ${err.message || 'Corrupted file'}`);
        onShowToast('Parse Error', 'Could not read bank statement.', 'error');
      }
    };
    reader.readAsBinaryString(f);
  };

  const handleLedgerUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const f = e.target.files?.[0];
    if (!f) return;
    setValidationError(null);

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const bstr = evt.target?.result;
        const wb = XLSX.read(bstr, { type: 'binary', cellDates: true });
        const json = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]) as Record<string, any>[];

        if (!json || json.length === 0) {
          setValidationError(`Ledger file "${f.name}" contains zero transaction rows.`);
          onShowToast('Empty Ledger', 'No rows found in uploaded ledger dump.', 'error');
          return;
        }

        const parsed = parseLedgerRows(json);
        setLedgerFile({ name: f.name, size: `${Math.round(f.size / 1024)} KB` });
        setLedgerItems(parsed);
        setIsReconciled(false);
      } catch (err: any) {
        setValidationError(`Failed to parse ledger file: ${err.message || 'Corrupted file'}`);
        onShowToast('Parse Error', 'Could not read ledger dump.', 'error');
      }
    };
    reader.readAsBinaryString(f);
  };

  const handleReconcile = () => {
    const bData = bankItems;
    const lData = ledgerItems;

    if (bData.length === 0 || lData.length === 0) {
      return;
    }

    const result = runBankReconciliation(bData, lData);
    setReconciledBank(result.reconciledBank);
    setReconciledLedger(result.reconciledLedger);
    setCalculation(result.calculation);
    setIsReconciled(true);

    // Synchronize BRS figures to active session storage for 1-click Dossier pulling
    try {
      const activeClient = getActiveClient();
      sessionStorage.setItem(
        'ca_copilot_session_brs',
        JSON.stringify({
          clientName: activeClient?.name || 'Active Client',
          bankClosingBalance: result.calculation.bankClosingBalance,
          booksClosingBalance: result.calculation.booksClosingBalance,
          uncreditedDeposits: result.calculation.uncreditedDeposits,
          unpresentedCheques: result.calculation.unpresentedCheques,
          unbookedBankCharges: result.calculation.unbookedBankCharges,
          unbookedBankInterest: result.calculation.unbookedBankInterest,
          adjustedBankBalance: result.calculation.adjustedBankBalance,
          netVariance: result.calculation.netVariance,
          matchedCount: result.calculation.matchedCount,
          unreconciledBankCount: result.reconciledBank.filter((b) => b.matchStatus === 'Unmatched').length,
          unreconciledLedgerCount: result.reconciledLedger.filter((l) => !l.isMatched).length,
          timestamp: new Date().toISOString(),
        })
      );
    } catch {
      // Storage unavailable or disabled
    }

    onShowToast(
      'Reconciliation Completed',
      `Matched ${result.calculation.matchedCount} transactions. Uncredited: ${formatINR(result.calculation.uncreditedDeposits)}, Unpresented: ${formatINR(result.calculation.unpresentedCheques)}.`,
      'success'
    );
  };

  const handleExportToTallyXml = () => {
    if (!calculation || reconciledBank.length === 0) return;
    const unrec = reconciledBank.filter((b) => b.matchStatus === 'Unmatched');
    if (unrec.length === 0) {
      onShowToast(
        'Fully Reconciled',
        'All bank transactions match ledger entries. No unbooked bank adjustments required!',
        'info'
      );
      return;
    }

    const activeClient = getActiveClient();
    const companyName = activeClient?.name || 'Client Accounts';
    const bankLedgerName = 'Bank Account';

    const vouchers: TallyVoucherItem[] = unrec.map((item, idx) => {
      const isDebit = item.withdrawal > 0;
      const classification = classifyBankNarration(item.narration, isDebit);
      return {
        id: `BRS-ADJ-${idx + 1}`,
        date: item.date || new Date().toISOString().substring(0, 10),
        voucherType: classification.voucherType,
        voucherNumber: `BRS/${item.refNo || idx + 1}`,
        bankLedger: bankLedgerName,
        partyOrExpenseLedger: classification.suggestedLedger,
        amount: isDebit ? item.withdrawal : item.deposit,
        narration: `BRS Auto-Adjustment: ${item.narration}${item.refNo ? ` [Ref: ${item.refNo}]` : ''}`,
        chequeOrRefNo: item.refNo || undefined,
      };
    });

    // Save batch to session storage so Tally Generator tool can also inspect and edit it!
    try {
      sessionStorage.setItem('ca_copilot_tally_batch', JSON.stringify(vouchers));
    } catch {
      // Storage unavailable
    }

    // Generate XML and trigger direct download
    const xml = generateTallyXml({ companyName, vouchers });
    downloadTallyXml(
      xml,
      `Tally_BRS_Adjustments_${new Date().toISOString().substring(0, 10)}.xml`
    );

    onShowToast(
      'Tally XML Generated',
      `Exported ${vouchers.length} unbooked bank adjustments to Tally Prime XML. Also synced with Tally Direct Generator.`,
      'success'
    );
  };

  const handleDownload = () => {
    if (!calculation) return;
    const wb = createBrsWorkbook(calculation, reconciledBank, reconciledLedger);
    XLSX.writeFile(wb, 'Bank_Reconciliation_Statement.xlsx');
    onShowToast('Downloaded', 'BRS statement & transaction working papers downloaded in Excel.', 'success');
  };

  return (
    <div className="space-y-6">
      <ToolHeader
        toolId="bank-rec"
        title="Bank Reconciliation (BRS)"
        description="Match bank transactions with ledger entries automatically using amount and reference heuristics."
        badge="Top Priority"
        onBack={onBack}
      />

      {/* 2. Format Requirements Guide */}
      <DataFormatGuide
        toolType="bank-rec"
        acceptedFormats=".xlsx, .xls, .csv"
        requiredFields={[
          { name: 'Transaction Date', description: 'Date of debit/credit', aliases: ['Date', 'Txn_Date', 'Value_Date', 'Posting_Date'] },
          { name: 'Narration / Particulars', description: 'Transaction memo or party name', aliases: ['Narration', 'Particulars', 'Description', 'Party_Name'] },
          { name: 'Cheque / Ref No', description: 'Cheque or UTR reference', aliases: ['Chq_No', 'Cheque_No', 'Ref_No', 'UTR', 'Voucher_No'] },
          { name: 'Withdrawal / Debit', description: 'Payments out of bank', aliases: ['Withdrawal', 'Debit', 'Dr', 'Paid_Out'] },
          { name: 'Deposit / Credit', description: 'Receipts into bank', aliases: ['Deposit', 'Credit', 'Cr', 'Received'] },
          { name: 'Running Balance', description: 'Closing balance after transaction', aliases: ['Balance', 'Closing_Balance', 'Available_Balance'] },
        ]}
        notes={[
          'Upload both the official Bank Statement (.xlsx/.csv) and the Tally/SAP Bank Ledger dump.',
          'The engine performs dual matching: exact Cheque/UTR numbers, exact amount with date proximity, and approximate party name narration matching.',
          'Automatically distinguishes unpresented cheques, uncredited deposits, unbooked bank charges, and auto-sweep interest.',
        ]}
      />

      {/* Validation Error Banner */}
      {validationError && (
        <div className="p-4 rounded-xl bg-[#FDF2F2] border border-[#F8D7DA] text-xs text-[#9B2C2C] flex items-center justify-between">
          <div className="flex items-center gap-2">
            <AlertCircle className="w-4 h-4 shrink-0" />
            <span>{validationError}</span>
          </div>
          <button onClick={() => setValidationError(null)} className="underline hover:no-underline font-semibold">
            Dismiss
          </button>
        </div>
      )}

      {/* 3. Upload Area: Statement + Ledger */}
      <div className="bg-white rounded-xl border border-gray-200 p-6 shadow-sm">
        <div className="grid grid-cols-1 md:grid-cols-2 gap-5">
          {/* Box 1: Bank Statement */}
          <input
            type="file"
            accept=".xlsx,.xls,.csv"
            className="hidden"
            id="brs-bank-upload"
            onChange={handleBankUpload}
          />
          {!bankFile ? (
            <div className="bg-white border-2 border-dashed border-gray-200 hover:border-gray-300 rounded-xl p-6 sm:p-8 text-center transition-colors shadow-2xs">
              <label htmlFor="brs-bank-upload" className="cursor-pointer flex flex-col items-center">
                <div className="w-12 h-12 rounded-full bg-gray-100 border border-gray-200 flex items-center justify-center text-gray-900 mb-3 shadow-2xs">
                  <UploadCloud className="w-6 h-6" />
                </div>
                <h3 className="text-sm font-semibold text-gray-900">1. Client Bank Statement</h3>
                <p className="text-xs text-gray-500 mt-1 max-w-xs mx-auto">
                  Official bank statement export in Excel (.xlsx, .xls) or CSV (.csv)
                </p>
                <div className="mt-4 px-4 py-2 bg-gray-900 hover:bg-black text-white text-xs font-semibold rounded-lg shadow-2xs inline-flex items-center gap-2 transition-all">
                  <UploadCloud className="w-4 h-4" />
                  <span>Upload Bank Statement</span>
                </div>
              </label>
            </div>
          ) : (
            <div className="bg-white rounded-xl border border-gray-200 p-5 shadow-sm flex flex-col justify-between">
              <div className="flex items-center gap-3">
                <div className="w-10 h-10 rounded-lg bg-gray-100 border border-gray-200 flex items-center justify-center text-gray-900 shrink-0">
                  <UploadCloud className="w-5 h-5" />
                </div>
                <div className="min-w-0 flex-1">
                  <div className="flex items-center gap-2">
                    <span className="text-xs font-bold text-gray-900 block truncate">{bankFile.name}</span>
                    <span className="text-[10px] font-semibold text-gray-900 bg-gray-100 border border-gray-200 px-2 py-0.5 rounded">
                      Ready
                    </span>
                  </div>
                  <span className="text-[11px] text-gray-600 block mt-0.5">
                    Bank Statement • {bankItems.length} transactions • {bankFile.size}
                  </span>
                </div>
              </div>
              <div className="mt-4 pt-3 border-t border-gray-100 flex items-center justify-between">
                <span className="text-[11px] text-gray-500 font-medium">Ready for reconciliation</span>
                <label htmlFor="brs-bank-upload" className="cursor-pointer px-3 py-1.5 text-xs text-gray-900 bg-gray-50 hover:bg-[#EAEAEA] rounded-lg border border-gray-200 flex items-center gap-1.5 font-medium transition-colors">
                  <UploadCloud className="w-3.5 h-3.5 text-gray-600" />
                  <span>Change Statement</span>
                </label>
              </div>
            </div>
          )}

          {/* Box 2: Accounting Ledger */}
          <input
            type="file"
            accept=".xlsx,.xls,.csv"
            className="hidden"
            id="brs-ledger-upload"
            onChange={handleLedgerUpload}
          />
          {!ledgerFile ? (
            <div className="bg-white border-2 border-dashed border-gray-200 hover:border-gray-300 rounded-xl p-6 sm:p-8 text-center transition-colors shadow-2xs">
              <label htmlFor="brs-ledger-upload" className="cursor-pointer flex flex-col items-center">
                <div className="w-12 h-12 rounded-full bg-gray-100 border border-gray-200 flex items-center justify-center text-gray-900 mb-3 shadow-2xs">
                  <BookOpen className="w-6 h-6" />
                </div>
                <h3 className="text-sm font-semibold text-gray-900">2. Accounting Ledger (Bank Book)</h3>
                <p className="text-xs text-gray-500 mt-1 max-w-xs mx-auto">
                  Bank ledger export from Tally, SAP, Zoho Books, or ERP (.xlsx, .csv)
                </p>
                <div className="mt-4 px-4 py-2 bg-gray-900 hover:bg-black text-white text-xs font-semibold rounded-lg shadow-2xs inline-flex items-center gap-2 transition-all">
                  <UploadCloud className="w-4 h-4" />
                  <span>Upload Accounting Ledger</span>
                </div>
              </label>
            </div>
          ) : (
            <div className="bg-white rounded-xl border border-gray-200 p-5 shadow-sm flex flex-col justify-between">
              <div className="flex items-center gap-3">
                <div className="w-10 h-10 rounded-lg bg-gray-100 border border-gray-200 flex items-center justify-center text-gray-900 shrink-0">
                  <BookOpen className="w-5 h-5" />
                </div>
                <div className="min-w-0 flex-1">
                  <div className="flex items-center gap-2">
                    <span className="text-xs font-bold text-gray-900 block truncate">{ledgerFile.name}</span>
                    <span className="text-[10px] font-semibold text-gray-900 bg-gray-100 border border-gray-200 px-2 py-0.5 rounded">
                      Ready
                    </span>
                  </div>
                  <span className="text-[11px] text-gray-600 block mt-0.5">
                    Accounting Ledger • {ledgerItems.length} vouchers • {ledgerFile.size}
                  </span>
                </div>
              </div>
              <div className="mt-4 pt-3 border-t border-gray-100 flex items-center justify-between">
                <span className="text-[11px] text-gray-500 font-medium">Ready for reconciliation</span>
                <label htmlFor="brs-ledger-upload" className="cursor-pointer px-3 py-1.5 text-xs text-gray-900 bg-gray-50 hover:bg-[#EAEAEA] rounded-lg border border-gray-200 flex items-center gap-1.5 font-medium transition-colors">
                  <UploadCloud className="w-3.5 h-3.5 text-gray-600" />
                  <span>Change Ledger</span>
                </label>
              </div>
            </div>
          )}
        </div>

        <div className="mt-6 pt-5 border-t border-gray-200 flex flex-col items-center justify-center gap-2">
          <button
            onClick={handleReconcile}
            disabled={bankItems.length === 0 || ledgerItems.length === 0}
            className="px-6 py-2.5 text-xs font-semibold text-white bg-gray-900 hover:bg-black disabled:opacity-40 disabled:cursor-not-allowed rounded-xl transition-all shadow-sm flex items-center justify-center gap-2 cursor-pointer"
          >
            <RefreshCw className="w-4 h-4" />
            <span>Reconcile Statements</span>
          </button>
          <span className="text-[11px] text-gray-500 text-center">
            {bankItems.length > 0 && ledgerItems.length > 0
              ? `Ready: ${bankItems.length} bank transactions & ${ledgerItems.length} ledger vouchers loaded`
              : 'Both Bank Statement and Accounting Ledger will be cross-matched by amount and reference'}
          </span>
        </div>
      </div>

      {/* 4. Reconciliation Results & BRS Statement */}
      {isReconciled && calculation && (
        <div className="space-y-5">
          {/* Formal BRS Calculation Box */}
          <div className="bg-white rounded-xl border border-gray-200 p-6 shadow-sm space-y-4">
            <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 pb-4 border-b border-gray-200">
              <div>
                <h3 className="text-sm font-bold text-gray-900 tracking-tight">
                  Bank Reconciliation Statement (BRS)
                </h3>
                <p className="text-xs text-gray-600 mt-0.5">
                  Reconciled against books of accounts with all timing and ledger variances accounted for.
                </p>
              </div>
              <div className="flex flex-wrap items-center gap-2">
                <button
                  onClick={handleExportToTallyXml}
                  title="Generate standard Tally Prime XML for all unbooked bank items"
                  className="inline-flex items-center gap-1.5 px-3 py-2 bg-white hover:bg-gray-50 text-gray-800 border border-gray-300 rounded-lg text-xs font-semibold transition-colors shadow-2xs shrink-0 cursor-pointer"
                >
                  <FileCode className="w-3.5 h-3.5 text-gray-600" />
                  <span>Export Adjustments to Tally XML</span>
                </button>
                <button
                  onClick={handleDownload}
                  className="inline-flex items-center gap-1.5 px-4 py-2 bg-gray-900 hover:bg-black text-white rounded-lg text-xs font-semibold transition-colors shadow-sm shrink-0 cursor-pointer"
                >
                  <Download className="w-3.5 h-3.5" />
                  <span>Export BRS Workbook</span>
                </button>
              </div>
            </div>

            {/* BRS Table */}
            <div className="border border-gray-200 rounded-lg overflow-hidden bg-white text-xs">
              <table className="w-full text-left">
                <thead className="bg-gray-50 border-b border-gray-200 text-gray-600 text-[11px] font-semibold">
                  <tr>
                    <th className="py-2.5 px-4 border-r border-gray-200">Particulars</th>
                    <th className="py-2.5 px-4 text-right w-48">Amount (₹)</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-[#E5E7EB] text-gray-900">
                  <tr>
                    <td className="py-2 px-4 border-r border-gray-200 font-semibold">
                      Balance as per Bank Statement
                    </td>
                    <td className="py-2 px-4 text-right font-mono font-semibold">
                      {formatINR(calculation.bankClosingBalance)}
                    </td>
                  </tr>
                  <tr className="bg-white">
                    <td className="py-2 px-4 border-r border-gray-200 text-gray-900">
                      Add: Cheques deposited but not yet cleared (Uncredited)
                    </td>
                    <td className="py-2 px-4 text-right font-mono text-gray-900">
                      +{formatINR(calculation.uncreditedDeposits)}
                    </td>
                  </tr>
                  <tr>
                    <td className="py-2 px-4 border-r border-gray-200 text-[#C43D43]">
                      Less: Cheques issued but not yet presented for payment (Unpresented)
                    </td>
                    <td className="py-2 px-4 text-right font-mono text-[#C43D43]">
                      -{formatINR(calculation.unpresentedCheques)}
                    </td>
                  </tr>
                  <tr className="bg-white">
                    <td className="py-2 px-4 border-r border-gray-200 text-[#C43D43]">
                      Less: Bank charges debited by bank not in cash book
                    </td>
                    <td className="py-2 px-4 text-right font-mono text-[#C43D43]">
                      -{formatINR(calculation.unbookedBankCharges)}
                    </td>
                  </tr>
                  <tr>
                    <td className="py-2 px-4 border-r border-gray-200 text-gray-900">
                      Add: Interest credited by bank not recorded in books
                    </td>
                    <td className="py-2 px-4 text-right font-mono text-gray-900">
                      +{formatINR(calculation.unbookedBankInterest)}
                    </td>
                  </tr>
                  <tr className="bg-gray-50 font-bold text-gray-900">
                    <td className="py-2.5 px-4 border-r border-gray-200">
                      Adjusted Balance as per Books of Accounts
                    </td>
                    <td className="py-2.5 px-4 text-right font-mono">
                      {formatINR(calculation.adjustedBankBalance)}
                    </td>
                  </tr>
                  <tr>
                    <td className="py-2 px-4 border-r border-gray-200 font-semibold text-gray-900">
                      Actual Balance as per Bank Ledger (Cash Book)
                    </td>
                    <td className="py-2 px-4 text-right font-mono font-semibold">
                      {formatINR(calculation.booksClosingBalance)}
                    </td>
                  </tr>
                  <tr className={`font-bold ${calculation.netVariance < 1 ? 'bg-gray-100 text-gray-900' : 'bg-[#FDF2F2] text-[#C43D43]'}`}>
                    <td className="py-2.5 px-4 border-r border-gray-200">
                      Net Reconciliation Variance
                    </td>
                    <td className="py-2.5 px-4 text-right font-mono">
                      {calculation.netVariance < 1 ? '₹ 0.00 (Fully Reconciled)' : formatINR(calculation.netVariance)}
                    </td>
                  </tr>
                </tbody>
              </table>
            </div>
          </div>

          {/* Tabular Transaction Preview */}
          <div className="bg-white rounded-xl border border-gray-200 overflow-hidden shadow-sm">
            <div className="px-5 py-3 bg-gray-50 border-b border-gray-200 flex flex-col sm:flex-row sm:items-center justify-between gap-2 text-xs">
              <span className="font-bold text-gray-900">
                Itemized Reconciliation Register
              </span>

              {/* Tabs */}
              <div className="flex items-center gap-1.5">
                <button
                  onClick={() => setActiveTab('all')}
                  className={`px-2.5 py-1 rounded-md text-[11px] font-medium transition-colors ${
                    activeTab === 'all'
                      ? 'bg-white text-gray-900 border border-gray-200 shadow-sm font-semibold'
                      : 'text-gray-600 hover:text-gray-900'
                  }`}
                >
                  All Bank Items ({reconciledBank.length})
                </button>
                <button
                  onClick={() => setActiveTab('matched')}
                  className={`px-2.5 py-1 rounded-md text-[11px] font-medium transition-colors ${
                    activeTab === 'matched'
                      ? 'bg-white text-gray-900 border border-gray-200 shadow-sm font-semibold'
                      : 'text-gray-600 hover:text-gray-900'
                  }`}
                >
                  Matched ({calculation.matchedCount})
                </button>
                <button
                  onClick={() => setActiveTab('unrecBank')}
                  className={`px-2.5 py-1 rounded-md text-[11px] font-medium transition-colors ${
                    activeTab === 'unrecBank'
                      ? 'bg-white text-[#C43D43] border border-gray-200 shadow-sm font-semibold'
                      : 'text-gray-600 hover:text-gray-900'
                  }`}
                >
                  Bank Unbooked ({calculation.unreconciledBankCount})
                </button>
                <button
                  onClick={() => setActiveTab('unrecLedger')}
                  className={`px-2.5 py-1 rounded-md text-[11px] font-medium transition-colors ${
                    activeTab === 'unrecLedger'
                      ? 'bg-white text-[#B45309] border border-gray-200 shadow-sm font-semibold'
                      : 'text-gray-600 hover:text-gray-900'
                  }`}
                >
                  Ledger In-Transit ({calculation.unreconciledLedgerCount})
                </button>
              </div>
            </div>

            <div className="overflow-x-auto max-h-[460px]">
              <table className="w-full text-left text-xs border-collapse">
                <thead className="sticky top-0 bg-gray-50 border-b border-gray-200 text-gray-600 text-[11px] font-semibold z-10">
                  <tr>
                    <th className="py-2.5 px-3 border-r border-gray-200 w-10 text-center">#</th>
                    <th className="py-2.5 px-3 border-r border-gray-200">Date</th>
                    <th className="py-2.5 px-3 border-r border-gray-200">Particulars / Narration</th>
                    <th className="py-2.5 px-3 border-r border-gray-200">Ref / Chq</th>
                    <th className="py-2.5 px-3 border-r border-gray-200 text-right">Debit (Wdl)</th>
                    <th className="py-2.5 px-3 border-r border-gray-200 text-right">Credit (Dep)</th>
                    <th className="py-2.5 px-3 whitespace-nowrap">Reconciliation Status</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-[#E5E7EB] text-gray-900">
                  {activeTab !== 'unrecLedger'
                    ? reconciledBank
                        .filter((b) => {
                          if (activeTab === 'matched') return b.matchStatus !== 'Unmatched';
                          if (activeTab === 'unrecBank') return b.matchStatus === 'Unmatched';
                          return true;
                        })
                        .map((row, idx) => (
                          <tr key={row.id} className="hover:bg-gray-50 transition-colors">
                            <td className="py-2 px-3 border-r border-gray-200 text-center text-gray-600 text-[11px]">
                              {idx + 1}
                            </td>
                            <td className="py-2 px-3 border-r border-gray-200 whitespace-nowrap font-mono text-[11px]">
                              {row.date}
                            </td>
                            <td className="py-2 px-3 border-r border-gray-200">
                              <span className="font-medium text-gray-900 block">{row.narration}</span>
                              {row.matchReason && (
                                <span className="text-[10px] text-gray-600 block mt-0.5">
                                  {row.matchReason}
                                </span>
                              )}
                            </td>
                            <td className="py-2 px-3 border-r border-gray-200 font-mono text-[11px]">
                              {row.refNo || '-'}
                            </td>
                            <td className="py-2 px-3 border-r border-gray-200 text-right font-mono">
                              {row.withdrawal > 0 ? formatINR(row.withdrawal) : '-'}
                            </td>
                            <td className="py-2 px-3 border-r border-gray-200 text-right font-mono">
                              {row.deposit > 0 ? formatINR(row.deposit) : '-'}
                            </td>
                            <td className="py-2 px-3 whitespace-nowrap">
                              {row.matchStatus === 'Matched' ? (
                                <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded text-[10px] font-semibold bg-gray-100 text-gray-900">
                                  Matched
                                </span>
                              ) : row.matchStatus === 'Probable' ? (
                                <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded text-[10px] font-semibold bg-[#FEF9EE] text-[#B45309]">
                                  Probable Match
                                </span>
                              ) : (
                                <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded text-[10px] font-semibold bg-[#FDF2F2] text-[#C43D43]">
                                  Unbooked Bank Item
                                </span>
                              )}
                            </td>
                          </tr>
                        ))
                    : reconciledLedger
                        .filter((l) => !l.isMatched)
                        .map((l, idx) => (
                          <tr key={l.id} className="hover:bg-gray-50 transition-colors">
                            <td className="py-2 px-3 border-r border-gray-200 text-center text-gray-600 text-[11px]">
                              {idx + 1}
                            </td>
                            <td className="py-2 px-3 border-r border-gray-200 whitespace-nowrap font-mono text-[11px]">
                              {l.date}
                            </td>
                            <td className="py-2 px-3 border-r border-gray-200 font-medium text-gray-900">
                              {l.particulars}
                            </td>
                            <td className="py-2 px-3 border-r border-gray-200 font-mono text-[11px]">
                              {l.voucherNo}
                            </td>
                            <td className="py-2 px-3 border-r border-gray-200 text-right font-mono">
                              {l.credit > 0 ? formatINR(l.credit) : '-'}
                            </td>
                            <td className="py-2 px-3 border-r border-gray-200 text-right font-mono">
                              {l.debit > 0 ? formatINR(l.debit) : '-'}
                            </td>
                            <td className="py-2 px-3 whitespace-nowrap">
                              <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded text-[10px] font-semibold bg-[#FEF3C7] text-[#92400E]">
                                {l.credit > 0 ? 'Cheque Issued (Unpresented)' : 'Cheque Deposited (Uncredited)'}
                              </span>
                            </td>
                          </tr>
                        ))}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};