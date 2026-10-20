import React, { useState } from 'react';
import {
  FileCode,
  Download,
  Copy,
  Check,
  X,
  ExternalLink,
  Layers,
  ArrowRight,
  Info,
} from 'lucide-react';
import {
  type TallyVoucherItem,
  generateTallyXml,
  downloadTallyXml,
} from '../../utils/tallyXmlGenerator';

interface TallyXmlExportModalProps {
  isOpen: boolean;
  onClose: () => void;
  clientName?: string;
  defaultBankLedger?: string;
  initialVouchers?: TallyVoucherItem[];
  onShowToast: (title: string, desc?: string, type?: 'success' | 'error' | 'info') => void;
}

export const TallyXmlExportModal: React.FC<TallyXmlExportModalProps> = ({
  isOpen,
  onClose,
  clientName = 'Client Company',
  defaultBankLedger = 'HDFC Bank Current Account',
  initialVouchers,
  onShowToast,
}) => {
  const [bankLedger, setBankLedger] = useState(defaultBankLedger);
  const [copied, setCopied] = useState(false);
  const [activeTab, setActiveTab] = useState<'preview' | 'instructions'>('preview');

  // Sample or provided reconciliation vouchers
  const sampleVouchers: TallyVoucherItem[] = initialVouchers && initialVouchers.length > 0
    ? initialVouchers
    : [
        {
          date: '2024-09-15',
          voucherType: 'Payment',
          voucherNumber: 'BRS-CHG-001',
          bankLedger: defaultBankLedger,
          partyOrExpenseLedger: 'Bank Charges & Commissions',
          amount: 590.0,
          narration: 'Bank Charges debited as per Bank Statement (Reconciled via CA Copilot)',
        },
        {
          date: '2024-09-20',
          voucherType: 'Receipt',
          voucherNumber: 'BRS-DIR-002',
          bankLedger: defaultBankLedger,
          partyOrExpenseLedger: 'Interest on Bank Savings/FD',
          amount: 14250.0,
          narration: 'Direct interest credited by bank not in cash book',
        },
        {
          date: '2024-09-28',
          voucherType: 'Payment',
          voucherNumber: 'BRS-TAX-003',
          bankLedger: defaultBankLedger,
          partyOrExpenseLedger: 'Advance Income Tax A/c',
          amount: 50000.0,
          narration: 'Direct challan debit for Advance Tax Q2',
        },
      ];

  const xmlOutput = generateTallyXml({
    companyName: clientName,
    vouchers: sampleVouchers.map((v) => ({ ...v, bankLedger })),
  });

  if (!isOpen) return null;

  const handleDownload = () => {
    const filename = `Tally_Import_${clientName.replace(/[^a-zA-Z0-9]/g, '_')}_${Date.now()}.xml`;
    downloadTallyXml(filename, xmlOutput);
    onShowToast('Tally XML Downloaded', `Generated ${sampleVouchers.length} vouchers ready for Tally Prime.`, 'success');
  };

  const handleCopy = () => {
    navigator.clipboard.writeText(xmlOutput);
    setCopied(true);
    setTimeout(() => setCopied(false), 2000);
    onShowToast('Copied to Clipboard', 'Tally XML payload copied.', 'info');
  };

  return (
    <div className="fixed inset-0 z-50 bg-black/40 backdrop-blur-xs flex items-center justify-center p-4 animate-apple-fade select-none">
      <div className="bg-white rounded-2xl border border-gray-200 w-full max-w-3xl shadow-xl overflow-hidden flex flex-col max-h-[88vh] animate-apple-scale">
        {/* Header */}
        <div className="h-14 px-6 border-b border-gray-100 flex items-center justify-between shrink-0 bg-white">
          <div className="flex items-center gap-2.5">
            <div className="w-8 h-8 rounded-lg bg-emerald-50 text-emerald-600 flex items-center justify-center font-bold text-xs">
              TALLY
            </div>
            <div>
              <h3 className="text-sm font-bold text-gray-900">Tally Prime Direct XML Generator</h3>
              <p className="text-[11px] text-gray-500">
                Export auto-reconciled adjustments directly into Tally Prime without manual data entry
              </p>
            </div>
          </div>
          <button
            onClick={onClose}
            className="w-8 h-8 rounded-lg flex items-center justify-center text-gray-400 hover:text-gray-700 hover:bg-gray-100 transition-colors cursor-pointer"
          >
            <X className="w-4 h-4" />
          </button>
        </div>

        {/* Content Body */}
        <div className="flex-1 overflow-y-auto p-6 space-y-4">
          {/* Top Options Bar */}
          <div className="grid grid-cols-1 sm:grid-cols-2 gap-3.5 bg-gray-50 p-4 rounded-xl border border-gray-200 text-xs">
            <div>
              <label className="font-semibold text-gray-700 block mb-1">Target Client Company</label>
              <div className="font-medium text-gray-900 bg-white px-3 py-1.5 rounded-lg border border-gray-200 truncate">
                {clientName}
              </div>
            </div>
            <div>
              <label className="font-semibold text-gray-700 block mb-1">Tally Bank Ledger Name</label>
              <input
                type="text"
                value={bankLedger}
                onChange={(e) => setBankLedger(e.target.value)}
                placeholder="e.g. HDFC Bank Current A/c"
                className="w-full bg-white px-3 py-1.5 rounded-lg border border-gray-200 text-gray-900 focus:outline-none focus:border-gray-900"
              />
            </div>
          </div>

          {/* Segmented View: XML Preview vs Instructions */}
          <div className="flex items-center justify-between pt-1">
            <div className="inline-flex p-1 bg-gray-100 border border-gray-200 rounded-lg text-xs">
              <button
                onClick={() => setActiveTab('preview')}
                className={`px-3 py-1 rounded-md font-semibold transition-all cursor-pointer ${
                  activeTab === 'preview' ? 'bg-white text-gray-900 shadow-2xs' : 'text-gray-500 hover:text-gray-900'
                }`}
              >
                XML Preview ({sampleVouchers.length} Vouchers)
              </button>
              <button
                onClick={() => setActiveTab('instructions')}
                className={`px-3 py-1 rounded-md font-semibold transition-all cursor-pointer ${
                  activeTab === 'instructions' ? 'bg-white text-gray-900 shadow-2xs' : 'text-gray-500 hover:text-gray-900'
                }`}
              >
                How to Import in Tally Prime
              </button>
            </div>

            <div className="flex items-center gap-2">
              <button
                onClick={handleCopy}
                className="inline-flex items-center gap-1.5 px-3 py-1.5 bg-gray-100 hover:bg-gray-200 text-gray-700 rounded-lg text-xs font-semibold transition-all active:scale-95 cursor-pointer"
              >
                {copied ? <Check className="w-3.5 h-3.5 text-emerald-600" /> : <Copy className="w-3.5 h-3.5" />}
                <span>{copied ? 'Copied' : 'Copy XML'}</span>
              </button>
            </div>
          </div>

          {/* Tab 1: XML Preview */}
          {activeTab === 'preview' ? (
            <div className="relative">
              <pre className="font-mono text-[11px] bg-slate-900 text-slate-100 p-4 rounded-xl overflow-x-auto max-h-72 border border-slate-800 leading-relaxed">
                {xmlOutput}
              </pre>
            </div>
          ) : (
            /* Tab 2: 3-Step Tally Prime Import Guide */
            <div className="bg-slate-50 border border-slate-200 rounded-xl p-5 space-y-4">
              <h4 className="text-xs font-bold text-slate-900 uppercase tracking-wider">
                Official Tally Prime Import Workflow (100% Offline)
              </h4>
              <div className="space-y-3 text-xs text-slate-700">
                <div className="flex items-start gap-3">
                  <div className="w-5 h-5 rounded-full bg-slate-200 text-slate-800 font-bold text-[10px] flex items-center justify-center shrink-0 mt-0.5">
                    1
                  </div>
                  <div>
                    <strong className="text-slate-900 block">Download the Generated XML File</strong>
                    <span className="text-slate-500 text-[11px]">Click the &ldquo;Download Tally XML&rdquo; button below to save the file on your PC.</span>
                  </div>
                </div>

                <div className="flex items-start gap-3">
                  <div className="w-5 h-5 rounded-full bg-slate-200 text-slate-800 font-bold text-[10px] flex items-center justify-center shrink-0 mt-0.5">
                    2
                  </div>
                  <div>
                    <strong className="text-slate-900 block">Open Tally Prime &gt; Import Menu</strong>
                    <span className="text-slate-500 text-[11px]">On the Gateway of Tally, press <kbd className="px-1 py-0.5 bg-white border border-slate-300 rounded text-[10px]">Alt + O</kbd> (or click <strong>Import</strong> in the top bar) and select <strong>Transactions</strong>.</span>
                  </div>
                </div>

                <div className="flex items-start gap-3">
                  <div className="w-5 h-5 rounded-full bg-slate-200 text-slate-800 font-bold text-[10px] flex items-center justify-center shrink-0 mt-0.5">
                    3
                  </div>
                  <div>
                    <strong className="text-slate-900 block">Select the XML File &amp; Import</strong>
                    <span className="text-slate-500 text-[11px]">Select the downloaded XML file. Tally Prime will instantly parse and create all payment and receipt vouchers with proper ledgers and narration!</span>
                  </div>
                </div>
              </div>
            </div>
          )}
        </div>

        {/* Modal Footer */}
        <div className="h-14 px-6 border-t border-gray-100 flex items-center justify-between shrink-0 bg-white">
          <div className="text-[11px] text-gray-500">
            Compliant with Tally.ERP 9 and Tally Prime 1.0 through 5.0
          </div>
          <div className="flex items-center gap-2.5">
            <button
              onClick={onClose}
              className="px-4 py-2 text-xs font-semibold text-gray-600 hover:text-gray-900 bg-gray-100 hover:bg-gray-200 rounded-xl cursor-pointer"
            >
              Close
            </button>
            <button
              onClick={handleDownload}
              className="inline-flex items-center gap-1.5 px-4 py-2 bg-emerald-600 hover:bg-emerald-700 text-white rounded-xl text-xs font-semibold shadow-2xs transition-all active:scale-95 cursor-pointer"
            >
              <Download className="w-4 h-4" />
              <span>Download Tally XML (.xml)</span>
            </button>
          </div>
        </div>
      </div>
    </div>
  );
};
