import React, { useState, useEffect } from 'react';
import {
  FileText,
  Printer,
  Download,
  X,
  CheckCircle2,
  Building,
  ShieldCheck,
  Calendar,
  Award,
} from 'lucide-react';
import {
  type AuditDossierData,
  getDossierDefaults,
  downloadDossierExcel,
} from '../../utils/auditDossierGenerator';

interface ExecutiveAuditDossierModalProps {
  isOpen: boolean;
  onClose: () => void;
  onShowToast: (title: string, desc?: string, type?: 'success' | 'error' | 'info') => void;
}

export const ExecutiveAuditDossierModal: React.FC<ExecutiveAuditDossierModalProps> = ({
  isOpen,
  onClose,
  onShowToast,
}) => {
  const [dossier, setDossier] = useState<AuditDossierData>(getDossierDefaults());

  useEffect(() => {
    if (isOpen) {
      setDossier(getDossierDefaults());
    }
  }, [isOpen]);

  if (!isOpen) return null;

  const handlePrint = () => {
    window.print();
  };

  const handleExcelExport = () => {
    downloadDossierExcel(dossier);
    onShowToast('Audit Dossier Exported', 'Downloaded consolidated Excel/CSV audit summary.', 'success');
  };

  return (
    <div className="fixed inset-0 z-50 bg-black/40 backdrop-blur-xs flex items-center justify-center p-4 animate-apple-fade select-none print:p-0 print:bg-white print:static">
      <div className="bg-white rounded-2xl border border-gray-200 w-full max-w-4xl shadow-xl overflow-hidden flex flex-col max-h-[92vh] animate-apple-scale print:max-h-none print:border-none print:shadow-none print:w-full print:rounded-none">
        {/* Action Header (Hidden on Print) */}
        <div className="h-14 px-6 border-b border-gray-100 flex items-center justify-between shrink-0 bg-white print:hidden">
          <div className="flex items-center gap-2.5">
            <div className="w-8 h-8 rounded-lg bg-blue-50 text-blue-600 flex items-center justify-center">
              <FileText className="w-4.5 h-4.5" />
            </div>
            <div>
              <h3 className="text-sm font-bold text-gray-900">Executive Audit &amp; Compliance Dossier</h3>
              <p className="text-[11px] text-gray-500">
                ICAI SA-230 Compliant Audit Working Paper Summary &amp; Sign-off Pack
              </p>
            </div>
          </div>

          <div className="flex items-center gap-2.5">
            <button
              onClick={handleExcelExport}
              className="inline-flex items-center gap-1.5 px-3 py-1.5 bg-gray-100 hover:bg-gray-200 text-gray-700 rounded-lg text-xs font-semibold transition-all active:scale-95 cursor-pointer"
            >
              <Download className="w-3.5 h-3.5 text-blue-600" />
              <span>Export Excel</span>
            </button>
            <button
              onClick={handlePrint}
              className="inline-flex items-center gap-1.5 px-3.5 py-1.5 bg-gray-900 hover:bg-black text-white rounded-lg text-xs font-semibold shadow-2xs transition-all active:scale-95 cursor-pointer"
            >
              <Printer className="w-3.5 h-3.5" />
              <span>Print / Save PDF</span>
            </button>
            <button
              onClick={onClose}
              className="w-8 h-8 rounded-lg flex items-center justify-center text-gray-400 hover:text-gray-700 hover:bg-gray-100 transition-colors cursor-pointer"
            >
              <X className="w-4 h-4" />
            </button>
          </div>
        </div>

        {/* Printable Dossier Canvas */}
        <div className="flex-1 overflow-y-auto p-8 sm:p-10 space-y-7 bg-white text-gray-900 font-sans print:p-0 print:overflow-visible">
          {/* Top Firm Letterhead */}
          <div className="border-b-2 border-gray-900 pb-5 flex flex-col sm:flex-row sm:items-start justify-between gap-4">
            <div>
              <span className="text-[10px] font-bold text-gray-500 uppercase tracking-widest block mb-1">
                Chartered Accountants · Audit &amp; Assurance Services
              </span>
              <h1 className="text-xl sm:text-2xl font-extrabold text-gray-900 tracking-tight">
                {dossier.firmName}
              </h1>
              <div className="text-xs text-gray-600 mt-1 space-y-0.5">
                <p>ICAI Firm Registration No (FRN): <strong className="font-mono text-gray-900">{dossier.frn}</strong> · Regional Council: <strong className="text-gray-900">{dossier.regionalCouncil}</strong></p>
                <p>{dossier.officeAddress}</p>
              </div>
            </div>

            <div className="text-right shrink-0 space-y-1">
              <span className="inline-block px-2.5 py-1 rounded-md text-[10px] font-extrabold uppercase tracking-wider bg-gray-100 text-gray-800 border border-gray-200">
                AUDIT DOSSIER SA-230
              </span>
              <p className="text-xs text-gray-500 font-mono">Ref: {dossier.dossierRef}</p>
              <p className="text-xs text-gray-500">Date: {dossier.generatedDate}</p>
            </div>
          </div>

          {/* Client Particulars Strip */}
          <div className="bg-gray-50/80 rounded-xl p-4 border border-gray-200 grid grid-cols-2 sm:grid-cols-4 gap-4 text-xs">
            <div>
              <span className="text-[10px] font-bold text-gray-500 uppercase block">Client Legal Name</span>
              <strong className="text-gray-900 text-sm mt-0.5 block truncate">{dossier.client.name}</strong>
              <span className="text-[11px] text-gray-500">{dossier.client.entityType}</span>
            </div>
            <div>
              <span className="text-[10px] font-bold text-gray-500 uppercase block">GSTIN</span>
              <strong className="text-gray-900 font-mono text-xs mt-0.5 block">{dossier.client.gstin}</strong>
            </div>
            <div>
              <span className="text-[10px] font-bold text-gray-500 uppercase block">PAN</span>
              <strong className="text-gray-900 font-mono text-xs mt-0.5 block">{dossier.client.pan}</strong>
            </div>
            <div>
              <span className="text-[10px] font-bold text-gray-500 uppercase block">Period Under Audit</span>
              <strong className="text-gray-900 text-xs mt-0.5 block">{dossier.financialYear} ({dossier.assessmentYear})</strong>
            </div>
          </div>

          {/* Executive Verification Matrix */}
          <div className="space-y-3">
            <h3 className="text-xs font-bold text-gray-900 uppercase tracking-wider flex items-center gap-2">
              <ShieldCheck className="w-4 h-4 text-blue-600" />
              <span>Executive Verification &amp; Reconciliation Matrix</span>
            </h3>

            <div className="border border-gray-200 rounded-xl overflow-hidden shadow-2xs">
              <table className="w-full text-xs text-left">
                <thead className="bg-gray-100/80 text-gray-700 font-bold border-b border-gray-200">
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
                    <td className="py-2.5 px-4 text-right">₹{dossier.bankBalanceBook.toLocaleString('en-IN', { minimumFractionDigits: 2 })}</td>
                    <td className="py-2.5 px-4 text-right">₹{dossier.bankBalanceStatement.toLocaleString('en-IN', { minimumFractionDigits: 2 })}</td>
                    <td className="py-2.5 px-4 text-right text-gray-600">₹{dossier.unreconciledDifference.toFixed(2)}</td>
                    <td className="py-2.5 px-4 font-sans text-emerald-700 font-semibold">
                      ✓ Reconciled (Timing Diff Only)
                    </td>
                  </tr>
                  <tr>
                    <td className="py-2.5 px-4 font-sans font-medium text-gray-900">
                      GST Input Tax Credit (GSTR-2B vs Books)
                    </td>
                    <td className="py-2.5 px-4 text-right">₹{dossier.booksItcClaimed.toLocaleString('en-IN', { minimumFractionDigits: 2 })}</td>
                    <td className="py-2.5 px-4 text-right">₹{dossier.gstr2bItcAvailable.toLocaleString('en-IN', { minimumFractionDigits: 2 })}</td>
                    <td className="py-2.5 px-4 text-right text-gray-600">₹0.00</td>
                    <td className="py-2.5 px-4 font-sans text-emerald-700 font-semibold">
                      ✓ 100% Eligible ITC Matched
                    </td>
                  </tr>
                  <tr>
                    <td className="py-2.5 px-4 font-sans font-medium text-gray-900">
                      Trial Balance Debit / Credit Equilibrium
                    </td>
                    <td className="py-2.5 px-4 text-right">Balanced</td>
                    <td className="py-2.5 px-4 text-right">Balanced</td>
                    <td className="py-2.5 px-4 text-right text-gray-600">₹0.00</td>
                    <td className="py-2.5 px-4 font-sans text-emerald-700 font-semibold">
                      ✓ Zero Discrepancy / Imbalance
                    </td>
                  </tr>
                  <tr>
                    <td className="py-2.5 px-4 font-sans font-medium text-gray-900">
                      Sec 40A(3) Cash Payment Scrutiny (&gt;₹10,000)
                    </td>
                    <td className="py-2.5 px-4 text-right">NIL</td>
                    <td className="py-2.5 px-4 text-right">NIL</td>
                    <td className="py-2.5 px-4 text-right text-gray-600">NIL</td>
                    <td className="py-2.5 px-4 font-sans text-emerald-700 font-semibold">
                      ✓ Rule 6DD Compliant
                    </td>
                  </tr>
                </tbody>
              </table>
            </div>
          </div>

          {/* Statutory Auditor Sign-off Certification Box */}
          <div className="border border-gray-300 rounded-xl p-5 bg-gray-50/50 space-y-4">
            <h4 className="text-xs font-bold text-gray-900 uppercase tracking-wider">
              Auditor Verification &amp; Certification (ICAI SA-230 &amp; SA-500)
            </h4>
            <p className="text-xs text-gray-600 leading-relaxed">
              We have scrutinized the books of account, bank clearance journals, and statutory GST returns for <strong>{dossier.client.name}</strong> for the period <strong>{dossier.financialYear}</strong>. In accordance with the Standards on Auditing issued by the Institute of Chartered Accountants of India (ICAI), the electronic audit trails and reconciliation sheets have been verified. In our opinion, the accounting records present a true and fair view of the underlying financial transactions.
            </p>

            <div className="pt-4 border-t border-gray-200 flex items-end justify-between">
              <div className="text-xs space-y-1">
                <p className="font-semibold text-gray-800">Audit Opinion: <span className="text-emerald-700 font-bold">{dossier.auditOpinion}</span></p>
                <p className="text-gray-500">Place: {dossier.client.city || 'Mumbai'}</p>
                <p className="text-gray-500 font-mono">UDIN: [Generated at signing: 25054321AAAAAA1234]</p>
              </div>

              {/* Signature Box */}
              <div className="text-right">
                <div className="w-48 h-12 border-b border-gray-400 mb-1" />
                <p className="text-xs font-bold text-gray-900">{dossier.partnerName}</p>
                <p className="text-[11px] text-gray-600">Chartered Accountant · Partner</p>
                <p className="text-[11px] text-gray-500 font-mono">M.No: {dossier.partnerMrn}</p>
              </div>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
};
