import React, { useState } from 'react';
import {
  FileText,
  Printer,
  Download,
  Building,
  ShieldCheck,
  CheckCircle2,
  Edit3,
  RefreshCw,
} from 'lucide-react';
import { ToolHeader } from '../common/ToolHeader';
import {
  type AuditDossierData,
  getDossierDefaults,
  downloadDossierExcel,
} from '../../utils/auditDossierGenerator';
import { getActiveClient } from '../../utils/clientStorage';

interface ExecutiveAuditDossierToolProps {
  onBack: () => void;
  onShowToast: (title: string, desc?: string, type?: 'success' | 'error' | 'info') => void;
}

export const ExecutiveAuditDossierTool: React.FC<ExecutiveAuditDossierToolProps> = ({
  onBack,
  onShowToast,
}) => {
  const [dossier, setDossier] = useState<AuditDossierData>(getDossierDefaults());
  const [isEditing, setIsEditing] = useState(false);

  // Editable Working Paper Figures (Real CA Data Input)
  const [booksBank, setBooksBank] = useState<string>('');
  const [statementBank, setStatementBank] = useState<string>('');
  const [booksItc, setBooksItc] = useState<string>('');
  const [portal2bItc, setPortal2bItc] = useState<string>('');
  const [tbDiff, setTbDiff] = useState<string>('0');
  const [opinion, setOpinion] = useState<
    'Clean / Unqualified' | 'Qualified' | 'Adverse' | 'Disclaimer of Opinion'
  >('Clean / Unqualified');

  const numBooksBank = parseFloat(booksBank) || 0;
  const numStatementBank = parseFloat(statementBank) || 0;
  const brsVariance = numBooksBank - numStatementBank;

  const numBooksItc = parseFloat(booksItc) || 0;
  const numPortal2bItc = parseFloat(portal2bItc) || 0;
  const itcVariance = numBooksItc - numPortal2bItc;

  const numTbDiff = parseFloat(tbDiff) || 0;

  const handleAutoPullSession = () => {
    let pulledCount = 0;
    const notes: string[] = [];

    // 1. Bank Reconciliation Session Data
    try {
      const rawBrs = sessionStorage.getItem('ca_copilot_session_brs');
      if (rawBrs) {
        const brs = JSON.parse(rawBrs);
        if (typeof brs.booksClosingBalance === 'number') {
          setBooksBank(brs.booksClosingBalance.toFixed(2));
          pulledCount++;
        }
        if (typeof brs.bankClosingBalance === 'number') {
          setStatementBank(brs.bankClosingBalance.toFixed(2));
          pulledCount++;
        }
        notes.push('BRS Balances');
      }
    } catch {}

    // 2. GST Reconciliation Session Data
    try {
      const rawGst = sessionStorage.getItem('ca_copilot_session_gst');
      if (rawGst) {
        const gst = JSON.parse(rawGst);
        if (typeof gst.totalBooksTax === 'number') {
          setBooksItc(gst.totalBooksTax.toFixed(2));
          pulledCount++;
        }
        if (typeof gst.total2bTax === 'number') {
          setPortal2bItc(gst.total2bTax.toFixed(2));
          pulledCount++;
        }
        notes.push('GST ITC & 2B Tax');
      }
    } catch {}

    // 3. Trial Balance Session Data
    try {
      const rawTb = sessionStorage.getItem('ca_copilot_session_tb');
      if (rawTb) {
        const tb = JSON.parse(rawTb);
        if (typeof tb.netDifference === 'number') {
          setTbDiff(tb.netDifference.toFixed(2));
          pulledCount++;
          notes.push('Trial Balance Difference');
        }
      }
    } catch {}

    if (pulledCount === 0) {
      onShowToast(
        'No Session Data Found',
        'Run Bank Reconciliation, GST 2B Reconciliation, or Trial Balance Checker first to auto-populate figures.',
        'info'
      );
      return;
    }

    setIsEditing(true);
    onShowToast(
      'Session Data Synchronized',
      `Auto-pulled live figures from: ${notes.join(', ')}.`,
      'success'
    );
  };

  const currentDossier: AuditDossierData = {
    ...dossier,
    client: getActiveClient(),
    bankBalanceBook: numBooksBank,
    bankBalanceStatement: numStatementBank,
    unreconciledDifference: brsVariance,
    booksItcClaimed: numBooksItc,
    gstr2bItcAvailable: numPortal2bItc,
    itcMismatch: itcVariance,
    trialBalanceDifference: numTbDiff,
    auditOpinion: opinion,
  };

  const handlePrint = () => {
    window.print();
    onShowToast('Print Ready', 'Opening official ICAI SA-230 working paper print view.', 'info');
  };

  const handleDownloadExcel = () => {
    try {
      downloadDossierExcel(currentDossier);
      onShowToast(
        'Excel Exported',
        `Saved audit working paper dossier for ${currentDossier.client.name}.`,
        'success'
      );
    } catch {
      onShowToast('Export Error', 'Failed to generate Excel dossier.', 'error');
    }
  };

  return (
    <div className="space-y-6 animate-apple-fade">
      {/* 1. Tool Header */}
      <ToolHeader
        toolId="audit-dossier"
        title="Executive Audit Dossier (ICAI SA-230 & SA-500)"
        description="Comprehensive audit working paper dossier compiling CA Firm credentials, Client master details, statutory reconciliation health, and auditor certification."
        onBack={onBack}
        actionButtons={
          <div className="flex flex-wrap items-center gap-2">
            <button
              onClick={handleAutoPullSession}
              title="Auto-pull reconciled figures from current session's BRS, GST 2B, and Trial Balance"
              className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-semibold text-gray-800 bg-white border border-gray-300 hover:bg-gray-50 transition-all cursor-pointer shadow-2xs"
            >
              <RefreshCw className="w-3.5 h-3.5 text-gray-600" />
              <span>Auto-Pull from Session</span>
            </button>
            <button
              onClick={() => setIsEditing(!isEditing)}
              className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-semibold text-gray-700 bg-white border border-gray-200 hover:bg-gray-50 hover:text-gray-900 transition-all cursor-pointer shadow-2xs"
            >
              <Edit3 className="w-3.5 h-3.5 text-gray-500" />
              <span>{isEditing ? 'Hide Figure Inputs' : 'Edit Audit Figures'}</span>
            </button>
            <button
              onClick={handleDownloadExcel}
              className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-semibold text-gray-700 bg-white border border-gray-200 hover:bg-gray-50 hover:text-gray-900 transition-all cursor-pointer shadow-2xs"
            >
              <Download className="w-3.5 h-3.5 text-gray-500" />
              <span>Download Excel (.xlsx)</span>
            </button>
            <button
              onClick={handlePrint}
              className="inline-flex items-center gap-1.5 px-3.5 py-1.5 rounded-lg text-xs font-semibold text-white bg-gray-900 hover:bg-black transition-all cursor-pointer shadow-2xs"
            >
              <Printer className="w-3.5 h-3.5" />
              <span>Print / Save PDF (A4)</span>
            </button>
          </div>
        }
      />

      {/* 2. Interactive Input Panel (Collapsible for entering Real Working Paper Data) */}
      {isEditing && (
        <div className="bg-white border border-gray-200 rounded-xl p-5 shadow-2xs space-y-4 print:hidden">
          <div className="flex items-center justify-between border-b border-gray-100 pb-3">
            <div>
              <span className="text-xs font-bold text-gray-900 block">
                Audited Working Paper Figures Entry
              </span>
              <span className="text-[11px] text-gray-500">
                Enter your client's actual books and return figures. Variances calculate automatically.
              </span>
            </div>
            <div className="flex items-center gap-2">
              <button
                onClick={handleAutoPullSession}
                className="inline-flex items-center gap-1 px-2.5 py-1 text-[11px] font-semibold text-gray-700 bg-gray-50 hover:bg-gray-100 border border-gray-200 rounded-md cursor-pointer transition-colors"
              >
                <RefreshCw className="w-3 h-3 text-gray-500" />
                <span>Pull from Session</span>
              </button>
              <button
                onClick={() => setIsEditing(false)}
                className="text-xs text-gray-600 hover:text-gray-900 font-medium cursor-pointer"
              >
                Done
              </button>
            </div>
          </div>

          <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-3 text-xs">
            <div>
              <label className="block text-[10px] font-bold text-gray-500 uppercase mb-1">
                Bank Balance (per Books) (₹)
              </label>
              <input
                type="number"
                step="0.01"
                placeholder="0.00"
                value={booksBank}
                onChange={(e) => setBooksBank(e.target.value)}
                className="w-full px-2.5 py-1.5 bg-gray-50 border border-gray-200 rounded-lg text-gray-900 font-mono focus:bg-white focus:outline-none focus:border-gray-900"
              />
            </div>

            <div>
              <label className="block text-[10px] font-bold text-gray-500 uppercase mb-1">
                Bank Balance (per Bank Stmt) (₹)
              </label>
              <input
                type="number"
                step="0.01"
                placeholder="0.00"
                value={statementBank}
                onChange={(e) => setStatementBank(e.target.value)}
                className="w-full px-2.5 py-1.5 bg-gray-50 border border-gray-200 rounded-lg text-gray-900 font-mono focus:bg-white focus:outline-none focus:border-gray-900"
              />
            </div>

            <div>
              <label className="block text-[10px] font-bold text-gray-500 uppercase mb-1">
                Books ITC Claimed (₹)
              </label>
              <input
                type="number"
                step="0.01"
                placeholder="0.00"
                value={booksItc}
                onChange={(e) => setBooksItc(e.target.value)}
                className="w-full px-2.5 py-1.5 bg-gray-50 border border-gray-200 rounded-lg text-gray-900 font-mono focus:bg-white focus:outline-none focus:border-gray-900"
              />
            </div>

            <div>
              <label className="block text-[10px] font-bold text-gray-500 uppercase mb-1">
                GSTR-2B Available ITC (₹)
              </label>
              <input
                type="number"
                step="0.01"
                placeholder="0.00"
                value={portal2bItc}
                onChange={(e) => setPortal2bItc(e.target.value)}
                className="w-full px-2.5 py-1.5 bg-gray-50 border border-gray-200 rounded-lg text-gray-900 font-mono focus:bg-white focus:outline-none focus:border-gray-900"
              />
            </div>
          </div>

          <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 text-xs pt-2 border-t border-gray-100">
            <div>
              <label className="block text-[10px] font-bold text-gray-500 uppercase mb-1">
                Trial Balance Imbalance (₹)
              </label>
              <input
                type="number"
                step="0.01"
                placeholder="0.00"
                value={tbDiff}
                onChange={(e) => setTbDiff(e.target.value)}
                className="w-full px-2.5 py-1.5 bg-gray-50 border border-gray-200 rounded-lg text-gray-900 font-mono focus:bg-white focus:outline-none focus:border-gray-900"
              />
            </div>

            <div>
              <label className="block text-[10px] font-bold text-gray-500 uppercase mb-1">
                Statutory Audit Opinion
              </label>
              <select
                value={opinion}
                onChange={(e) =>
                  setOpinion(
                    e.target.value as
                      | 'Clean / Unqualified'
                      | 'Qualified'
                      | 'Adverse'
                      | 'Disclaimer of Opinion'
                  )
                }
                className="w-full px-2.5 py-1.5 bg-gray-50 border border-gray-200 rounded-lg text-gray-900 focus:bg-white focus:outline-none focus:border-gray-900"
              >
                <option value="Clean / Unqualified">Clean / Unqualified Opinion</option>
                <option value="Qualified">Qualified Opinion</option>
                <option value="Adverse">Adverse Opinion</option>
                <option value="Disclaimer of Opinion">Disclaimer of Opinion</option>
              </select>
            </div>
          </div>
        </div>
      )}

      {/* 3. Official Dossier Document Container */}
      <div className="bg-white border border-gray-200 rounded-xl p-6 sm:p-8 shadow-xs space-y-6 text-gray-900 print:border-none print:shadow-none print:p-0">
        {/* Letterhead Header */}
        <div className="border-b-2 border-gray-900 pb-5 flex flex-col sm:flex-row sm:items-start justify-between gap-4">
          <div>
            <div className="text-[10px] font-bold text-gray-500 uppercase tracking-widest flex items-center gap-1.5">
              <ShieldCheck className="w-4 h-4 text-gray-700" />
              Statutory Audit Working Paper Dossier • ICAI SA-230 Compliant
            </div>
            <h1 className="text-xl font-bold text-gray-900 tracking-tight mt-1">
              {currentDossier.firmName}
            </h1>
            <p className="text-xs text-gray-600 font-medium">
              Chartered Accountants • FRN: {currentDossier.frn} • {currentDossier.regionalCouncil}
            </p>
            <p className="text-[11px] text-gray-500 mt-0.5">{currentDossier.officeAddress}</p>
          </div>

          <div className="text-right text-xs space-y-1 sm:border-l sm:border-gray-200 sm:pl-6 shrink-0">
            <span className="inline-block px-2.5 py-1 rounded-md text-[10px] font-bold uppercase tracking-wider bg-gray-100 text-gray-800 border border-gray-200">
              AUDIT DOSSIER SA-230
            </span>
            <p className="text-xs text-gray-500 font-mono">Ref: {currentDossier.dossierRef}</p>
            <p className="text-xs text-gray-500">Date: {currentDossier.generatedDate}</p>
            <p className="text-xs text-gray-500">
              FY: <strong className="text-gray-800">{currentDossier.financialYear}</strong> ({currentDossier.assessmentYear})
            </p>
          </div>
        </div>

        {/* Client Entity Particulars Card */}
        <div className="p-4 rounded-xl bg-gray-50/70 border border-gray-200">
          <div className="text-[11px] font-bold text-gray-500 uppercase tracking-wider mb-2 flex items-center gap-1.5">
            <Building className="w-3.5 h-3.5 text-gray-600" />
            Auditee Entity Particulars
          </div>
          <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-3 text-xs">
            <div>
              <span className="text-gray-500 block text-[11px]">Client Legal Name:</span>
              <span className="font-bold text-gray-900 text-sm">
                {currentDossier.client.name || 'Client Not Assigned'}
              </span>
              <span className="text-[10px] text-gray-500 block">
                {currentDossier.client.entityType}
              </span>
            </div>
            <div>
              <span className="text-gray-500 block text-[11px]">Trade / Brand Name:</span>
              <span className="font-medium text-gray-900">
                {currentDossier.client.tradeName || currentDossier.client.name || '—'}
              </span>
            </div>
            <div>
              <span className="text-gray-500 block text-[11px]">GSTIN:</span>
              <span className="font-mono font-semibold text-gray-900">
                {currentDossier.client.gstin || 'Not Configured'}
              </span>
            </div>
            <div>
              <span className="text-gray-500 block text-[11px]">Permanent Account No (PAN):</span>
              <span className="font-mono font-semibold text-gray-900">
                {currentDossier.client.pan || 'Not Configured'}
              </span>
            </div>
          </div>
        </div>

        {/* Consolidated Audit Verification Matrix */}
        <div className="space-y-3">
          <div className="flex items-center justify-between">
            <h3 className="text-xs font-bold text-gray-900 uppercase tracking-wider flex items-center gap-2">
              <ShieldCheck className="w-4 h-4 text-gray-700" />
              <span>Executive Verification &amp; Reconciliation Matrix</span>
            </h3>
            {!isEditing && (
              <button
                onClick={() => setIsEditing(true)}
                className="text-xs text-gray-500 hover:text-gray-900 underline print:hidden cursor-pointer"
              >
                Enter Actual Figures
              </button>
            )}
          </div>

          <div className="border border-gray-200 rounded-xl overflow-hidden shadow-2xs">
            <table className="w-full text-xs text-left">
              <thead className="bg-gray-100 text-gray-700 font-bold border-b border-gray-200">
                <tr>
                  <th className="py-2.5 px-4">Audit Subject</th>
                  <th className="py-2.5 px-4 text-right">Books Record</th>
                  <th className="py-2.5 px-4 text-right">Audited / Bank Figure</th>
                  <th className="py-2.5 px-4 text-right">Variance</th>
                  <th className="py-2.5 px-4">Status &amp; ICAI Compliance</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-gray-100 font-mono">
                <tr>
                  <td className="py-2.5 px-4 font-sans font-medium text-gray-900">
                    Bank Balance Reconciliation (BRS)
                  </td>
                  <td className="py-2.5 px-4 text-right">
                    ₹{numBooksBank.toLocaleString('en-IN', { minimumFractionDigits: 2 })}
                  </td>
                  <td className="py-2.5 px-4 text-right">
                    ₹{numStatementBank.toLocaleString('en-IN', { minimumFractionDigits: 2 })}
                  </td>
                  <td className="py-2.5 px-4 text-right text-gray-700">
                    ₹{brsVariance.toFixed(2)}
                  </td>
                  <td className="py-2.5 px-4 font-sans text-gray-800 font-semibold">
                    {brsVariance === 0 ? '✓ Balanced (₹0.00)' : 'Pending Clearance / Adjustments'}
                  </td>
                </tr>
                <tr>
                  <td className="py-2.5 px-4 font-sans font-medium text-gray-900">
                    GST Input Tax Credit (GSTR-2B vs Books)
                  </td>
                  <td className="py-2.5 px-4 text-right">
                    ₹{numBooksItc.toLocaleString('en-IN', { minimumFractionDigits: 2 })}
                  </td>
                  <td className="py-2.5 px-4 text-right">
                    ₹{numPortal2bItc.toLocaleString('en-IN', { minimumFractionDigits: 2 })}
                  </td>
                  <td className="py-2.5 px-4 text-right text-gray-700">
                    ₹{itcVariance.toFixed(2)}
                  </td>
                  <td className="py-2.5 px-4 font-sans text-gray-800 font-semibold">
                    {itcVariance === 0 ? '✓ 100% Eligible ITC Matched' : 'Variance Awaiting Review'}
                  </td>
                </tr>
                <tr>
                  <td className="py-2.5 px-4 font-sans font-medium text-gray-900">
                    Trial Balance Debit / Credit Equilibrium
                  </td>
                  <td className="py-2.5 px-4 text-right">
                    {numTbDiff === 0 ? 'Balanced' : 'Imbalance'}
                  </td>
                  <td className="py-2.5 px-4 text-right">
                    {numTbDiff === 0 ? 'Balanced' : 'Imbalance'}
                  </td>
                  <td className="py-2.5 px-4 text-right text-gray-700">
                    ₹{numTbDiff.toFixed(2)}
                  </td>
                  <td className="py-2.5 px-4 font-sans text-gray-800 font-semibold">
                    {numTbDiff === 0 ? '✓ Zero Imbalance Detected' : 'Imbalance Detected'}
                  </td>
                </tr>
                <tr>
                  <td className="py-2.5 px-4 font-sans font-medium text-gray-900">
                    Sec 40A(3) Cash Payment Scrutiny (&gt;₹10,000)
                  </td>
                  <td className="py-2.5 px-4 text-right">NIL</td>
                  <td className="py-2.5 px-4 text-right">NIL</td>
                  <td className="py-2.5 px-4 text-right text-gray-700">NIL</td>
                  <td className="py-2.5 px-4 font-sans text-gray-800 font-semibold">
                    ✓ Rule 6DD Compliant
                  </td>
                </tr>
              </tbody>
            </table>
          </div>
        </div>

        {/* Auditor Certification & Sign-off Block */}
        <div className="border border-gray-200 rounded-xl p-5 bg-gray-50/50 space-y-4">
          <h4 className="text-xs font-bold text-gray-900 uppercase tracking-wider">
            Auditor Verification &amp; Certification (ICAI SA-230 &amp; SA-500)
          </h4>
          <p className="text-xs text-gray-600 leading-relaxed">
            We have scrutinized the books of account, bank clearance journals, and statutory GST returns for{' '}
            <strong>{currentDossier.client.name || 'the auditee entity'}</strong> for the period{' '}
            <strong>{currentDossier.financialYear}</strong>. In accordance with the Standards on Auditing issued by the Institute of Chartered Accountants of India (ICAI), the electronic audit trails and reconciliation sheets have been verified. In our opinion, the accounting records present a true and fair view of the underlying financial transactions.
          </p>

          <div className="pt-4 border-t border-gray-200 flex flex-col sm:flex-row items-end justify-between gap-4">
            <div className="text-xs space-y-1">
              <p className="font-semibold text-gray-800">
                Audit Opinion: <span className="font-bold text-gray-900">{opinion}</span>
              </p>
              <p className="text-gray-500">Place: {currentDossier.client.city || 'Office Jurisdiction'}</p>
              <p className="text-gray-500 font-mono text-[11px]">
                Classification: Official Working Paper Dossier
              </p>
            </div>

            <div className="text-right text-xs space-y-1">
              <p className="font-bold text-gray-900">For {currentDossier.firmName}</p>
              <p className="text-gray-500">Chartered Accountants (FRN: {currentDossier.frn})</p>
              <div className="h-10 flex items-center justify-end font-serif italic text-gray-400 text-sm">
                [Authorized Partner Signature]
              </div>
              <p className="font-bold text-gray-900">{currentDossier.partnerName}</p>
              <p className="text-gray-500">Partner • M. No. {currentDossier.partnerMrn}</p>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
};
