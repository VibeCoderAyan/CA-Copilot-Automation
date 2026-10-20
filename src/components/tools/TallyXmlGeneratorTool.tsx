import React, { useState } from 'react';
import {
  FileCode,
  Download,
  Copy,
  Check,
  Plus,
  Trash2,
  Info,
  RotateCcw,
  Sparkles,
} from 'lucide-react';
import { ToolHeader } from '../common/ToolHeader';
import {
  type TallyVoucherItem,
  generateTallyXml,
  downloadTallyXml,
} from '../../utils/tallyXmlGenerator';
import { getActiveClient, type ClientProfile } from '../../utils/clientStorage';

interface TallyXmlGeneratorToolProps {
  onBack: () => void;
  onShowToast: (title: string, desc?: string, type?: 'success' | 'error' | 'info') => void;
  onOpenClientVault?: () => void;
}

const SAMPLE_DEMO_VOUCHERS: TallyVoucherItem[] = [
  {
    id: 'demo_1',
    date: '2026-03-31',
    voucherType: 'Payment',
    voucherNumber: 'BRS/001',
    bankLedger: 'HDFC Bank Current A/c',
    partyOrExpenseLedger: 'Bank Charges & Commission',
    amount: 1450.0,
    narration: 'Being quarterly ledger charges debited by bank per BRS.',
  },
  {
    id: 'demo_2',
    date: '2026-03-31',
    voucherType: 'Receipt',
    voucherNumber: 'BRS/002',
    bankLedger: 'HDFC Bank Current A/c',
    partyOrExpenseLedger: 'Interest on Fixed Deposit',
    amount: 8340.0,
    narration: 'Being interest on auto-sweep deposit credited directly by bank.',
  },
];

export const TallyXmlGeneratorTool: React.FC<TallyXmlGeneratorToolProps> = ({
  onBack,
  onShowToast,
  onOpenClientVault,
}) => {
  const [activeClient] = useState<ClientProfile>(getActiveClient());
  const [bankLedger, setBankLedger] = useState('');
  const [vouchers, setVouchers] = useState<TallyVoucherItem[]>([]);
  const [activeTab, setActiveTab] = useState<'register' | 'xml'>('register');
  const [copied, setCopied] = useState(false);

  const [sessionBatchCount, setSessionBatchCount] = useState<number>(() => {
    try {
      const raw = sessionStorage.getItem('ca_copilot_tally_batch');
      if (raw) {
        const parsed = JSON.parse(raw);
        return Array.isArray(parsed) ? parsed.length : 0;
      }
    } catch {}
    return 0;
  });

  const handleLoadSessionBatch = () => {
    try {
      const raw = sessionStorage.getItem('ca_copilot_tally_batch');
      if (!raw) {
        onShowToast(
          'No Session Adjustments',
          'Run Bank Reconciliation first to generate adjustment entries.',
          'info'
        );
        return;
      }
      const parsed: TallyVoucherItem[] = JSON.parse(raw);
      if (!Array.isArray(parsed) || parsed.length === 0) {
        onShowToast('Empty Batch', 'No vouchers found in session storage.', 'info');
        return;
      }

      setVouchers(parsed);
      if (parsed[0]?.bankLedger && !bankLedger) {
        setBankLedger(parsed[0].bankLedger);
      }
    } catch {
      onShowToast('Load Failed', 'Could not parse vouchers from session.', 'error');
    }
  };

  // New Voucher Entry State
  const [newDate, setNewDate] = useState(() => new Date().toISOString().slice(0, 10));
  const [newType, setNewType] = useState<'Payment' | 'Receipt' | 'Journal'>('Payment');
  const [newLedger, setNewLedger] = useState('');
  const [newAmount, setNewAmount] = useState('');
  const [newNarration, setNewNarration] = useState('');

  const effectiveBankLedger = bankLedger.trim() || 'Bank A/c';

  const xmlContent = generateTallyXml({
    companyName: activeClient.name || 'Practice Auditee',
    vouchers: vouchers.map((v) => ({
      ...v,
      bankLedger: v.bankLedger || effectiveBankLedger,
    })),
  });

  const totalPayments = vouchers
    .filter((v) => v.voucherType === 'Payment')
    .reduce((s, v) => s + v.amount, 0);

  const totalReceipts = vouchers
    .filter((v) => v.voucherType === 'Receipt')
    .reduce((s, v) => s + v.amount, 0);

  const handleAddVoucher = (e: React.FormEvent) => {
    e.preventDefault();
    const parsedAmount = parseFloat(newAmount);
    if (!newLedger.trim() || isNaN(parsedAmount) || parsedAmount <= 0) {
      onShowToast(
        'Invalid Input',
        'Please enter an account ledger name and valid positive amount.',
        'error'
      );
      return;
    }

    const newVoucher: TallyVoucherItem = {
      id: `v_${Date.now()}_${Math.random().toString(36).substring(2, 5)}`,
      date: newDate,
      voucherType: newType,
      voucherNumber: `ADJ/${vouchers.length + 1}`,
      bankLedger: effectiveBankLedger,
      partyOrExpenseLedger: newLedger.trim(),
      amount: parsedAmount,
      narration:
        newNarration.trim() ||
        `Being ${newType.toLowerCase()} adjustment posted per reconciliation.`,
    };

    setVouchers([...vouchers, newVoucher]);
    setNewLedger('');
    setNewAmount('');
    setNewNarration('');
    onShowToast('Voucher Added', `Added ${newVoucher.partyOrExpenseLedger} to batch.`, 'success');
  };

  const handleDeleteVoucher = (id?: string) => {
    if (!id) return;
    setVouchers(vouchers.filter((v) => v.id !== id));
    onShowToast('Removed', 'Voucher entry removed from batch.', 'info');
  };

  const handleClearAll = () => {
    setVouchers([]);
    onShowToast('Cleared', 'All batch vouchers cleared.', 'info');
  };

  const handleLoadDemo = () => {
    setBankLedger('HDFC Bank Current A/c');
    setVouchers(SAMPLE_DEMO_VOUCHERS);
  };

  const handleCopyXml = async () => {
    if (vouchers.length === 0) {
      onShowToast('Batch Empty', 'Add at least one voucher to generate Tally XML.', 'error');
      return;
    }
    try {
      await navigator.clipboard.writeText(xmlContent);
      setCopied(true);
      onShowToast(
        'XML Copied',
        'Tally Prime XML copied to clipboard. Ready for paste / import.',
        'success'
      );
      setTimeout(() => setCopied(false), 2500);
    } catch {
      onShowToast('Copy Failed', 'Could not copy XML to clipboard.', 'error');
    }
  };

  const handleDownloadXml = () => {
    if (vouchers.length === 0) {
      onShowToast('Batch Empty', 'Add at least one voucher to export Tally XML.', 'error');
      return;
    }
    const safeName = (activeClient.name || 'Client')
      .replace(/[^a-zA-Z0-9]/g, '_')
      .substring(0, 30);
    const filename = `Tally_Import_${safeName}_${new Date().toISOString().slice(0, 10)}.xml`;
    downloadTallyXml(xmlContent, filename);
    onShowToast('XML Downloaded', `Saved ${filename}. Ready for Tally Prime Alt + O import.`, 'success');
  };

  return (
    <div className="space-y-6 animate-apple-fade">
      {/* 1. Header */}
      <ToolHeader
        toolId="tally-xml"
        title="Tally Prime Direct XML Generator"
        description="Deterministic Tally XML voucher generator for automated adjustment import (Bank charges, Interest, Direct collections) without manual entry."
        onBack={onBack}
        actionButtons={
          <div className="flex items-center gap-2">
            <button
              onClick={handleCopyXml}
              disabled={vouchers.length === 0}
              className={`inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-semibold border transition-all cursor-pointer shadow-2xs ${
                vouchers.length === 0
                  ? 'opacity-40 cursor-not-allowed bg-gray-50 text-gray-400 border-gray-200'
                  : 'bg-white text-gray-800 border-gray-200 hover:bg-gray-50 hover:text-gray-900'
              }`}
            >
              {copied ? (
                <>
                  <Check className="w-3.5 h-3.5 text-gray-900" />
                  <span className="text-gray-900">Copied!</span>
                </>
              ) : (
                <>
                  <Copy className="w-3.5 h-3.5 text-gray-400" />
                  <span>Copy XML</span>
                </>
              )}
            </button>
            <button
              onClick={handleDownloadXml}
              disabled={vouchers.length === 0}
              className={`inline-flex items-center gap-1.5 px-3.5 py-1.5 rounded-lg text-xs font-semibold text-white transition-all shadow-2xs ${
                vouchers.length === 0
                  ? 'opacity-40 cursor-not-allowed bg-gray-400'
                  : 'bg-gray-900 hover:bg-black cursor-pointer'
              }`}
            >
              <Download className="w-3.5 h-3.5" />
              <span>Download .xml File</span>
            </button>
          </div>
        }
      />

      {/* 2. Top Client & Bank Ledger Row */}
      <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
        {/* Client Entity */}
        <div className="p-4 rounded-xl bg-white border border-gray-200 shadow-2xs">
          <div className="flex items-center justify-between mb-1.5">
            <span className="text-[11px] font-bold text-gray-500 uppercase tracking-wider">
              Auditee Client
            </span>
            {onOpenClientVault && (
              <button
                onClick={onOpenClientVault}
                className="text-xs text-gray-900 hover:underline font-semibold cursor-pointer"
              >
                Change Client
              </button>
            )}
          </div>
          <div className="text-sm font-bold text-gray-900 truncate">
            {activeClient.name || 'No Client Selected'}
          </div>
          <div className="text-xs font-mono text-gray-500 mt-0.5">
            {activeClient.gstin ? `GSTIN: ${activeClient.gstin}` : 'GSTIN: Not Set'}
          </div>
        </div>

        {/* Bank Ledger Setting */}
        <div className="p-4 rounded-xl bg-white border border-gray-200 shadow-2xs md:col-span-2">
          <label className="block text-[11px] font-bold text-gray-500 uppercase tracking-wider mb-1.5">
            Tally Bank Ledger Name (Exact match in Tally Chart of Accounts)
          </label>
          <input
            type="text"
            value={bankLedger}
            onChange={(e) => setBankLedger(e.target.value)}
            className="w-full px-3 py-2 bg-gray-50 border border-gray-200 rounded-lg text-xs font-mono text-gray-900 focus:bg-white focus:outline-none focus:border-gray-900 transition-colors"
            placeholder="e.g. HDFC Bank Current A/c - 50200012345678"
          />
        </div>
      </div>

      {/* 3. Summary Metrics Bar */}
      <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
        <div className="p-3.5 rounded-xl bg-white border border-gray-200 shadow-2xs">
          <span className="text-[11px] font-semibold text-gray-500 uppercase tracking-wider block">
            Adjustment Vouchers
          </span>
          <span className="text-lg font-bold font-mono text-gray-900 mt-1 block">
            {vouchers.length} Entries
          </span>
        </div>
        <div className="p-3.5 rounded-xl bg-white border border-gray-200 shadow-2xs">
          <span className="text-[11px] font-semibold text-gray-500 uppercase tracking-wider block">
            Total Payments (Debits)
          </span>
          <span className="text-lg font-bold font-mono text-gray-900 mt-1 block">
            ₹{totalPayments.toLocaleString('en-IN', { minimumFractionDigits: 2 })}
          </span>
        </div>
        <div className="p-3.5 rounded-xl bg-white border border-gray-200 shadow-2xs">
          <span className="text-[11px] font-semibold text-gray-500 uppercase tracking-wider block">
            Total Receipts (Credits)
          </span>
          <span className="text-lg font-bold font-mono text-gray-900 mt-1 block">
            ₹{totalReceipts.toLocaleString('en-IN', { minimumFractionDigits: 2 })}
          </span>
        </div>
      </div>

      {/* 4. Tab Navigation & Controls */}
      <div className="flex items-center justify-between border-b border-gray-200 pb-2">
        <div className="flex items-center gap-2">
          <button
            onClick={() => setActiveTab('register')}
            className={`px-3.5 py-1.5 rounded-lg text-xs font-semibold transition-all cursor-pointer ${
              activeTab === 'register'
                ? 'bg-gray-900 text-white shadow-2xs'
                : 'text-gray-600 hover:text-gray-900 hover:bg-gray-100'
            }`}
          >
            Voucher Register ({vouchers.length})
          </button>
          <button
            onClick={() => setActiveTab('xml')}
            className={`px-3.5 py-1.5 rounded-lg text-xs font-semibold transition-all cursor-pointer ${
              activeTab === 'xml'
                ? 'bg-gray-900 text-white shadow-2xs'
                : 'text-gray-600 hover:text-gray-900 hover:bg-gray-100'
            }`}
          >
            Raw Tally XML Preview
          </button>
        </div>

        <div className="flex items-center gap-2 text-xs">
          {sessionBatchCount > 0 && (
            <button
              onClick={handleLoadSessionBatch}
              className="inline-flex items-center gap-1.5 px-2.5 py-1 text-[11px] font-semibold text-gray-900 bg-gray-100 hover:bg-gray-200 border border-gray-300 rounded-md cursor-pointer transition-colors shadow-2xs"
              title="Load unreconciled bank adjustments generated in current BRS session"
            >
              <Sparkles className="w-3 h-3 text-gray-600" />
              <span>Load BRS Batch ({sessionBatchCount})</span>
            </button>
          )}

          {vouchers.length === 0 ? (
            <button
              onClick={handleLoadDemo}
              className="text-gray-500 hover:text-gray-900 underline text-[11px] cursor-pointer"
            >
              Load Sample Demo Vouchers
            </button>
          ) : (
            <button
              onClick={handleClearAll}
              className="text-gray-500 hover:text-gray-900 text-[11px] flex items-center gap-1 cursor-pointer"
            >
              <RotateCcw className="w-3 h-3" />
              <span>Clear Batch</span>
            </button>
          )}
        </div>
      </div>

      {/* 5. Main Content */}
      {activeTab === 'register' ? (
        <div className="space-y-4">
          {/* Add Entry Form */}
          <form
            onSubmit={handleAddVoucher}
            className="p-4 rounded-xl bg-white border border-gray-200 shadow-2xs space-y-3"
          >
            <div className="flex items-center justify-between">
              <span className="text-xs font-bold text-gray-900 flex items-center gap-2">
                <Plus className="w-4 h-4 text-gray-500" />
                Add Voucher to Batch
              </span>
              <span className="text-[11px] text-gray-400">
                Tally.ERP 9 &amp; Tally Prime Standard
              </span>
            </div>

            <div className="grid grid-cols-1 sm:grid-cols-12 gap-2.5">
              <div className="sm:col-span-2">
                <label className="block text-[10px] font-bold text-gray-500 uppercase mb-1">
                  Voucher Date
                </label>
                <input
                  type="date"
                  value={newDate}
                  onChange={(e) => setNewDate(e.target.value)}
                  className="w-full px-2.5 py-1.5 text-xs bg-gray-50 border border-gray-200 rounded-lg text-gray-900 focus:bg-white focus:outline-none focus:border-gray-900 font-mono"
                  required
                />
              </div>

              <div className="sm:col-span-2">
                <label className="block text-[10px] font-bold text-gray-500 uppercase mb-1">
                  Voucher Type
                </label>
                <select
                  value={newType}
                  onChange={(e) =>
                    setNewType(e.target.value as 'Payment' | 'Receipt' | 'Journal')
                  }
                  className="w-full px-2.5 py-1.5 text-xs bg-gray-50 border border-gray-200 rounded-lg text-gray-900 focus:bg-white focus:outline-none focus:border-gray-900 font-medium"
                >
                  <option value="Payment">Payment</option>
                  <option value="Receipt">Receipt</option>
                  <option value="Journal">Journal</option>
                </select>
              </div>

              <div className="sm:col-span-4">
                <label className="block text-[10px] font-bold text-gray-500 uppercase mb-1">
                  Opposite Ledger Name
                </label>
                <input
                  type="text"
                  placeholder="e.g. Bank Charges / Customer Name"
                  value={newLedger}
                  onChange={(e) => setNewLedger(e.target.value)}
                  className="w-full px-2.5 py-1.5 text-xs bg-gray-50 border border-gray-200 rounded-lg text-gray-900 focus:bg-white focus:outline-none focus:border-gray-900"
                  required
                />
              </div>

              <div className="sm:col-span-2">
                <label className="block text-[10px] font-bold text-gray-500 uppercase mb-1">
                  Amount (₹)
                </label>
                <input
                  type="number"
                  step="0.01"
                  placeholder="0.00"
                  value={newAmount}
                  onChange={(e) => setNewAmount(e.target.value)}
                  className="w-full px-2.5 py-1.5 text-xs bg-gray-50 border border-gray-200 rounded-lg text-gray-900 focus:bg-white focus:outline-none focus:border-gray-900 font-mono"
                  required
                />
              </div>

              <div className="sm:col-span-2 flex items-end">
                <button
                  type="submit"
                  className="w-full py-1.5 px-3 bg-gray-900 hover:bg-black text-white text-xs font-semibold rounded-lg transition-all cursor-pointer shadow-2xs"
                >
                  + Add Entry
                </button>
              </div>
            </div>

            <div>
              <input
                type="text"
                placeholder="Narration (e.g. Being bank charges for March 2026 per BRS)..."
                value={newNarration}
                onChange={(e) => setNewNarration(e.target.value)}
                className="w-full px-2.5 py-1.5 text-xs bg-gray-50 border border-gray-200 rounded-lg text-gray-900 focus:bg-white focus:outline-none focus:border-gray-900"
              />
            </div>
          </form>

          {/* Table / Empty State */}
          <div className="bg-white border border-gray-200 rounded-xl overflow-hidden shadow-2xs">
            {vouchers.length === 0 ? (
              <div className="py-12 px-4 text-center space-y-2">
                <div className="w-10 h-10 rounded-xl bg-gray-100 text-gray-500 mx-auto flex items-center justify-center">
                  <FileCode className="w-5 h-5" />
                </div>
                <div className="text-sm font-semibold text-gray-800">
                  No Vouchers in Current Batch
                </div>
                <p className="text-xs text-gray-500 max-w-md mx-auto">
                  Add adjustment vouchers using the form above. Once added, you can preview the raw XML or download the file for Tally Prime.
                </p>
                {sessionBatchCount > 0 ? (
                  <div className="pt-2 flex items-center justify-center gap-3">
                    <button
                      onClick={handleLoadSessionBatch}
                      className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-semibold text-white bg-gray-900 hover:bg-black transition-all cursor-pointer shadow-2xs"
                    >
                      <Sparkles className="w-3.5 h-3.5" />
                      <span>Load {sessionBatchCount} Adjustments from BRS Session</span>
                    </button>
                    <button
                      onClick={handleLoadDemo}
                      className="text-xs text-gray-500 hover:text-gray-900 underline cursor-pointer"
                    >
                      or load sample demo
                    </button>
                  </div>
                ) : (
                  <button
                    onClick={handleLoadDemo}
                    className="mt-2 text-xs font-semibold text-gray-900 hover:underline cursor-pointer"
                  >
                    Load 2 Sample Entries for Preview
                  </button>
                )}
              </div>
            ) : (
              <div className="overflow-x-auto">
                <table className="w-full text-left text-xs">
                  <thead>
                    <tr className="bg-gray-50 border-b border-gray-200 text-[11px] font-bold text-gray-600 uppercase tracking-wider">
                      <th className="py-2.5 px-3.5">Date</th>
                      <th className="py-2.5 px-3.5">Type</th>
                      <th className="py-2.5 px-3.5">Vch Ref</th>
                      <th className="py-2.5 px-3.5">Opposite Ledger</th>
                      <th className="py-2.5 px-3.5 text-right">Amount (₹)</th>
                      <th className="py-2.5 px-3.5">Narration</th>
                      <th className="py-2.5 px-3.5 text-center">Action</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-gray-100">
                    {vouchers.map((v) => (
                      <tr key={v.id} className="hover:bg-gray-50/70 transition-colors">
                        <td className="py-2.5 px-3.5 font-mono text-gray-700 whitespace-nowrap">
                          {v.date}
                        </td>
                        <td className="py-2.5 px-3.5">
                          <span className="inline-block px-2 py-0.5 rounded text-[10px] font-bold bg-gray-100 text-gray-800 border border-gray-200">
                            {v.voucherType}
                          </span>
                        </td>
                        <td className="py-2.5 px-3.5 font-mono text-gray-500 whitespace-nowrap">
                          {v.voucherNumber}
                        </td>
                        <td className="py-2.5 px-3.5 font-medium text-gray-900">
                          {v.partyOrExpenseLedger}
                        </td>
                        <td className="py-2.5 px-3.5 font-mono font-semibold text-right text-gray-900 whitespace-nowrap">
                          ₹{v.amount.toLocaleString('en-IN', { minimumFractionDigits: 2 })}
                        </td>
                        <td className="py-2.5 px-3.5 text-gray-500 max-w-xs truncate">
                          {v.narration}
                        </td>
                        <td className="py-2.5 px-3.5 text-center">
                          <button
                            onClick={() => handleDeleteVoucher(v.id)}
                            className="p-1 rounded text-gray-400 hover:text-gray-900 hover:bg-gray-100 transition-colors cursor-pointer"
                            title="Delete entry"
                          >
                            <Trash2 className="w-3.5 h-3.5" />
                          </button>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}
          </div>
        </div>
      ) : (
        /* Tab 2: Raw XML Preview */
        <div className="bg-white border border-gray-200 rounded-xl overflow-hidden shadow-2xs">
          <div className="p-3 bg-gray-50 border-b border-gray-200 flex items-center justify-between">
            <span className="text-xs font-mono text-gray-600">
              Tally Schema Envelope (v11.0 / Prime Compatible)
            </span>
            <span className="text-[11px] font-mono text-gray-500">
              {(xmlContent.length / 1024).toFixed(1)} KB
            </span>
          </div>
          <pre className="p-4 text-[11px] font-mono text-gray-800 bg-gray-50/40 overflow-x-auto max-h-[460px] whitespace-pre">
            {xmlContent}
          </pre>
        </div>
      )}

      {/* 6. Step-by-Step Tally Prime Import Guide */}
      <div className="p-4 rounded-xl bg-gray-50 border border-gray-200 text-xs text-gray-600 space-y-2">
        <div className="font-bold text-gray-900 flex items-center gap-1.5">
          <Info className="w-4 h-4 text-gray-700 shrink-0" />
          <span>How to Import into Tally Prime:</span>
        </div>
        <ol className="list-decimal list-inside space-y-1 pl-1 text-[11px] text-gray-700">
          <li>
            Click <strong>"Download .xml File"</strong> above to save the file locally.
          </li>
          <li>
            Open your company in <strong>Tally Prime</strong> and press{' '}
            <kbd className="px-1.5 py-0.5 bg-white border border-gray-300 rounded text-[10px] font-mono">
              Alt + O
            </kbd>{' '}
            (or click <em>Import</em> at the top menu).
          </li>
          <li>
            Choose <strong>Transactions</strong>, select the downloaded XML file, and press{' '}
            <kbd className="px-1.5 py-0.5 bg-white border border-gray-300 rounded text-[10px] font-mono">
              Enter
            </kbd>
            .
          </li>
          <li>
            All adjustments will be posted directly to your Daybook with complete narrations.
          </li>
        </ol>
      </div>
    </div>
  );
};
