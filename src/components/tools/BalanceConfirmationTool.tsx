import { ToolHeader } from '../common/ToolHeader';
import React, { useState, useRef, useMemo } from 'react';
import * as XLSX from 'xlsx';
import {
  ArrowLeft,
  UploadCloud,
  FileSpreadsheet,
  AlertTriangle,
  AlertCircle,
  CheckCircle2,
  Check,
  X,
  Search,
  Sliders,
  ShieldCheck,
  Eye,
  RefreshCw,
  Building,
  Mail,
  Send,
  Printer,
  Calendar,
  FileText,
  UserCheck,
  Phone,
} from 'lucide-react';
import { InnerToolBadge } from '../common/InnerToolBadge';
import {
  type ConfirmationParty,
  type ConfirmationSummary,
  type ConfirmationStatus,
  type ConfirmationType,
  type ConfirmationColumnMapping,
  detectConfirmationColumns,
  evaluateConfirmationSummary,
  generateConfirmationLetterText,
  exportConfirmationWorkbook,
} from '../../utils/balanceConfirmationEngine';

interface BalanceConfirmationToolProps {
  onBack: () => void;
  onShowToast: (title: string, description?: string, type?: 'success' | 'error' | 'info') => void;
}

export const BalanceConfirmationTool: React.FC<BalanceConfirmationToolProps> = ({
  onBack,
  onShowToast,
}) => {
  const [parties, setParties] = useState<ConfirmationParty[]>([]);
  const [headers, setHeaders] = useState<string[]>([]);
  const [fileName, setFileName] = useState('');
  const [clientName, setClientName] = useState('');
  const [balanceAsOnDate, setBalanceAsOnDate] = useState('30-09-2026');
  const [auditorFirm, setAuditorFirm] = useState('M/s S. K. Agarwal & Associates, Chartered Accountants');
  const [auditorEmail, setAuditorEmail] = useState('audit.confirmations@skagarwal.in');

  const [activeTab, setActiveTab] = useState<'register' | 'differences' | 'letter_preview'>('register');
  const [selectedPartyId, setSelectedPartyId] = useState<string>('');
  const [searchQuery, setSearchQuery] = useState('');
  const [typeFilter, setTypeFilter] = useState<string>('all');
  const [statusFilter, setStatusFilter] = useState<string>('all');

  const fileInputRef = useRef<HTMLInputElement>(null);

  // Evaluate summary
  const summary: ConfirmationSummary = useMemo(() => {
    return evaluateConfirmationSummary(parties, clientName, balanceAsOnDate);
  }, [parties, clientName, balanceAsOnDate]);

  // Selected party for letter
  const selectedParty = useMemo(() => {
    return parties.find((p) => p.id === selectedPartyId) || parties[0];
  }, [parties, selectedPartyId]);

  // Handle file upload
  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setFileName(file.name);
    const reader = new FileReader();

    reader.onload = (event) => {
      try {
        const data = new Uint8Array(event.target?.result as ArrayBuffer);
        const wb = XLSX.read(data, { type: 'array' });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json<Record<string, any>>(ws, { defval: '' });

        if (json.length === 0) {
          onShowToast('Empty File', 'The uploaded file has no data rows.', 'error');
          return;
        }

        const detectedHeaders = Object.keys(json[0]);
        setHeaders(detectedHeaders);
        const map = detectConfirmationColumns(detectedHeaders);

        const loadedParties: ConfirmationParty[] = json.map((r, idx) => {
          const expectedBalance = typeof r[map.expectedBalCol] === 'number' ? r[map.expectedBalCol] : parseFloat(String(r[map.expectedBalCol]).replace(/[^0-9.-]/g, '')) || 0;
          return {
            id: `party-${idx + 1}`,
            partyName: String(r[map.partyCol] || `Party ${idx + 1}`).trim(),
            accountCode: String(r[map.codeCol] || '').trim(),
            confirmationType: 'Receivable',
            contactPerson: String(r[map.contactCol] || '').trim(),
            emailOrPhone: '',
            balanceDate: balanceAsOnDate,
            expectedBalance,
            confirmedBalance: null,
            status: 'ready_to_send',
            isDirectResponse: false,
            notes: '',
            raw: r,
          };
        });

        setParties(loadedParties);
        if (loadedParties.length > 0) setSelectedPartyId(loadedParties[0].id);
      } catch (err: any) {
        onShowToast('Upload Error', err.message || 'Failed to parse file', 'error');
      }
    };

    reader.readAsArrayBuffer(file);
    e.target.value = '';
  };

  // Update party field
  const handleUpdateParty = (id: string, updates: Partial<ConfirmationParty>) => {
    setParties((prev) =>
      prev.map((p) => (p.id === id ? { ...p, ...updates } : p))
    );
  };

  // Export Excel workbook
  const handleExportExcel = () => {
    if (parties.length === 0) {
      onShowToast('No Data', 'Please load accounts before exporting.', 'error');
      return;
    }

    try {
      const wb = exportConfirmationWorkbook(summary);
      XLSX.writeFile(wb, `Balance_Confirmation_Register_${clientName.replace(/\s+/g, '_')}_${balanceAsOnDate}.xlsx`);
      onShowToast('Workbook Exported', 'Exported 3-sheet balance confirmation register (Summary, Register, Differences).', 'success');
    } catch (err: any) {
      onShowToast('Export Error', err.message || 'Failed to generate Excel file', 'error');
    }
  };

  // Print Letter / Summary
  const handlePrint = () => {
    if (!selectedParty) {
      onShowToast('No Party Selected', 'Select a party to print confirmation letter.', 'error');
      return;
    }
    setActiveTab('letter_preview');
    setTimeout(() => {
      window.print();
    }, 300);
  };

  // Filtered parties
  const filteredParties = useMemo(() => {
    return parties.filter((p) => {
      if (typeFilter !== 'all' && p.confirmationType !== typeFilter) return false;
      if (statusFilter !== 'all' && p.status !== statusFilter) return false;

      if (searchQuery.trim()) {
        const q = searchQuery.toLowerCase();
        const name = p.partyName.toLowerCase();
        const code = p.accountCode.toLowerCase();
        const person = p.contactPerson.toLowerCase();
        return name.includes(q) || code.includes(q) || person.includes(q);
      }
      return true;
    });
  }, [parties, typeFilter, statusFilter, searchQuery]);

  return (
    <div className="space-y-6">
      <ToolHeader
        toolId="balance-confirmation"
        title="Balance Confirmation Tracker"
        description="Track debtors, creditors, and bank balance confirmation requests, compute variance tie-outs, and draft formal confirmation letters."
        badge="Direct Response"
        onBack={onBack}
        hidePrintHeader={activeTab === 'letter_preview'}
        actionButtons={
          <div className="flex flex-wrap items-center gap-2">
            

            <button
              onClick={() => fileInputRef.current?.click()}
              className="px-3 py-1.5 text-xs font-semibold rounded-md bg-white text-gray-900 border border-gray-200 hover:bg-gray-50 transition-colors flex items-center gap-1.5 shadow-sm"
            >
              <UploadCloud className="w-3.5 h-3.5 text-gray-900" />
              {fileName ? 'Re-upload Register' : 'Import Register'}
            </button>
            <input
              type="file"
              ref={fileInputRef}
              onChange={handleFileUpload}
              accept=".xlsx,.xls,.csv"
              className="hidden"
              id="balance-confirmation-upload"
            />

            <button
              onClick={handlePrint}
              disabled={parties.length === 0}
              className="px-3 py-1.5 text-xs font-semibold rounded-md bg-white text-gray-900 border border-gray-200 hover:bg-gray-50 disabled:opacity-50 transition-colors flex items-center gap-1.5"
            >
              <Printer className="w-3.5 h-3.5 text-gray-600" />
              Print Letter / PDF
            </button>

            <button
              onClick={handleExportExcel}
              disabled={parties.length === 0}
              className="px-3.5 py-1.5 text-xs font-semibold rounded-md bg-gray-900 text-white hover:bg-black disabled:opacity-50 transition-colors flex items-center gap-1.5 shadow-sm"
            >
              <FileSpreadsheet className="w-3.5 h-3.5" />
              Export Register (.xlsx)
            </button>
          </div>
        }
      />

      {/* Client & Date Row */}
      <div className={`bg-white border border-gray-200 rounded-xl p-4 shadow-2xs flex flex-wrap items-center gap-4 text-xs ${activeTab === 'letter_preview' ? 'print:hidden' : ''}`}>
          <div className="flex items-center gap-2">
            <Building className="w-4 h-4 text-gray-600" />
            <span className="font-semibold text-gray-600">Entity:</span>
            <input
              type="text"
              value={clientName}
              onChange={(e) => setClientName(e.target.value)}
              className="px-2 py-1 border border-gray-200 rounded text-xs font-medium focus:border-gray-300 outline-none w-52"
            />
          </div>

          <div className="flex items-center gap-2">
            <Calendar className="w-4 h-4 text-gray-600" />
            <span className="font-semibold text-gray-600">Balance As On:</span>
            <input
              type="text"
              value={balanceAsOnDate}
              onChange={(e) => setBalanceAsOnDate(e.target.value)}
              className="px-2 py-1 border border-gray-200 rounded text-xs font-medium focus:border-gray-300 outline-none w-28"
            />
          </div>

          <div className="flex items-center gap-2">
            <span className="font-semibold text-gray-600">Auditor:</span>
            <input
              type="text"
              value={auditorFirm}
              onChange={(e) => setAuditorFirm(e.target.value)}
              className="px-2 py-1 border border-gray-200 rounded text-xs font-medium focus:border-gray-300 outline-none w-64"
            />
          </div>
      </div>

      {/* Upload Dropzone when no file selected */}
      {!fileName && (
        <div className="bg-white border-2 border-dashed border-gray-200 hover:border-gray-300 rounded-xl p-8 sm:p-10 text-center transition-colors shadow-2xs">
          <label htmlFor="balance-confirmation-upload" className="cursor-pointer flex flex-col items-center">
            <div className="w-12 h-12 rounded-full bg-gray-100 border border-gray-200 flex items-center justify-center text-gray-900 mx-auto mb-3 shadow-2xs">
              <UploadCloud className="w-6 h-6" />
            </div>
            <h3 className="text-sm font-semibold text-gray-900">Upload Debtor / Creditor Register (.xlsx, .xls, .csv)</h3>
            <p className="text-xs text-gray-500 mt-1 max-w-md mx-auto leading-relaxed">
              Upload customer sundry debtors or vendor sundry creditors ledger. Automatically formats formal SA 505 audit balance confirmation letters for each party.
            </p>
            <div className="mt-4 px-4 py-2 bg-gray-900 hover:bg-black text-white text-xs font-semibold rounded-lg shadow-2xs inline-flex items-center gap-2 transition-all">
              <UploadCloud className="w-4 h-4" />
              <span>Import Ledger Register</span>
            </div>
            <span className="text-[11px] text-gray-400 font-medium mt-3 block">
              Supports .xlsx, .xls, .csv · SA 505 External Confirmations Compliant · Direct print & PDF export
            </span>
          </label>
        </div>
      )}

      {/* Main Container */}
      <div className="space-y-6">
        {/* KPI Cards */}
        {parties.length > 0 && (
          <div className={`grid grid-cols-2 md:grid-cols-5 gap-4 ${activeTab === 'letter_preview' ? 'print:hidden' : ''}`}>
            <div className="bg-white border border-gray-200 rounded-xl p-4 shadow-sm">
              <span className="text-xs font-semibold text-gray-600 uppercase">Parties In Scope</span>
              <div className="text-xl font-bold text-gray-900 mt-1">
                {summary.totalParties} accounts
              </div>
              <div className="text-xs text-gray-600 mt-0.5">
                Books: ₹{summary.totalExpectedValue.toLocaleString('en-IN')}
              </div>
            </div>

            <div className="bg-white border border-gray-200 rounded-xl p-4 shadow-sm">
              <span className="text-xs font-semibold text-gray-600 uppercase">Responses Received</span>
              <div className="text-xl font-bold text-gray-900 mt-1">
                {summary.responseReceivedCount + summary.closedCount}
              </div>
              <div className="text-xs text-gray-600 mt-0.5">Direct external confirmations</div>
            </div>

            <div className="bg-white border border-gray-200 rounded-xl p-4 shadow-sm">
              <span className="text-xs font-semibold text-gray-600 uppercase">Differences Reported</span>
              <div className="text-xl font-bold text-[#DC2626] mt-1">
                {summary.differenceReportedCount}
              </div>
              <div className="text-xs text-gray-600 mt-0.5">
                Variance: ₹{summary.netDifferenceValue.toLocaleString('en-IN')}
              </div>
            </div>

            <div className="bg-white border border-gray-200 rounded-xl p-4 shadow-sm">
              <span className="text-xs font-semibold text-gray-600 uppercase">Follow-ups Required</span>
              <div className="text-xl font-bold text-amber-600 mt-1">
                {summary.followupRequiredCount}
              </div>
              <div className="text-xs text-gray-600 mt-0.5">Awaiting party reply</div>
            </div>

            <div className="bg-white border border-gray-200 rounded-xl p-4 shadow-sm">
              <span className="text-xs font-semibold text-gray-600 uppercase">Ready / Sent</span>
              <div className="text-xl font-bold text-gray-600 mt-1">
                {summary.readyToSendCount + summary.sentCount}
              </div>
              <div className="text-xs text-gray-600 mt-0.5">Letters in flight</div>
            </div>
          </div>
        )}

        {/* Empty State */}
        {parties.length === 0 ? (
          <div className="bg-white border border-dashed border-gray-200 rounded-2xl p-12 text-center shadow-sm">
            <div className="w-16 h-16 bg-gray-50 rounded-full flex items-center justify-center mx-auto mb-4 border border-gray-300/20">
              <Mail className="w-8 h-8 text-gray-900" />
            </div>
            <h3 className="text-base font-bold text-gray-900">No Balance Confirmation Register Loaded</h3>
            <p className="text-xs text-gray-600 max-w-md mx-auto mt-1 mb-6">
              Upload your client's accounts receivable, accounts payable, or bank ledger export (Excel or CSV) to track external balance confirmations under SA 505.
            </p>
            <div className="flex items-center justify-center gap-3">
              
              <button
                onClick={() => fileInputRef.current?.click()}
                className="px-4 py-2 text-xs font-semibold rounded-lg bg-gray-900 text-white hover:bg-black transition-colors shadow-sm inline-flex items-center gap-2"
              >
                <UploadCloud className="w-4 h-4" />
                <span>Upload Confirmation Register</span>
              </button>
            </div>
          </div>
        ) : (
          <div className="space-y-6">
            {/* Tabs Bar */}
            <div className="flex flex-col sm:flex-row sm:items-center sm:justify-between gap-4 border-b border-gray-200 pb-1 print:hidden">
              <div className="flex items-center gap-2">
                <button
                  onClick={() => setActiveTab('register')}
                  className={`px-4 py-2 text-xs font-bold border-b-2 transition-colors ${
                    activeTab === 'register'
                      ? 'border-gray-300 text-gray-900'
                      : 'border-transparent text-gray-600 hover:text-gray-900'
                  }`}
                >
                  Confirmation Master Register ({parties.length})
                </button>
                <button
                  onClick={() => setActiveTab('differences')}
                  className={`px-4 py-2 text-xs font-bold border-b-2 transition-colors ${
                    activeTab === 'differences'
                      ? 'border-[#DC2626] text-[#DC2626]'
                      : 'border-transparent text-gray-600 hover:text-gray-900'
                  }`}
                >
                  Differences & Discrepancies ({summary.parties.filter((p) => p.differenceAmount !== null && p.differenceAmount !== 0).length})
                </button>
                <button
                  onClick={() => setActiveTab('letter_preview')}
                  className={`px-4 py-2 text-xs font-bold border-b-2 transition-colors ${
                    activeTab === 'letter_preview'
                      ? 'border-gray-300 text-gray-900'
                      : 'border-transparent text-gray-600 hover:text-gray-900'
                  }`}
                >
                  Draft Confirmation Letter Preview
                </button>
              </div>

              {activeTab === 'register' && (
                <div className="flex items-center gap-2">
                  <div className="relative">
                    <Search className="w-3.5 h-3.5 absolute left-2.5 top-1/2 -translate-y-1/2 text-gray-600" />
                    <input
                      type="text"
                      placeholder="Search party or code..."
                      value={searchQuery}
                      onChange={(e) => setSearchQuery(e.target.value)}
                      className="pl-8 pr-3 py-1.5 text-xs border border-gray-200 rounded-lg w-48 outline-none focus:border-gray-300"
                    />
                  </div>

                  <select
                    value={typeFilter}
                    onChange={(e) => setTypeFilter(e.target.value)}
                    className="text-xs border border-gray-200 rounded-lg px-2 py-1.5 outline-none"
                  >
                    <option value="all">All Types</option>
                    <option value="Receivable">Receivable</option>
                    <option value="Payable">Payable</option>
                    <option value="Bank Balance">Bank Balance</option>
                    <option value="Loan / Borrowing">Loan / Borrowing</option>
                  </select>
                </div>
              )}
            </div>

            {/* TAB 1: Master Register Table */}
            {activeTab === 'register' && (
              <div className="bg-white border border-gray-200 rounded-xl overflow-hidden shadow-sm">
                <div className="overflow-x-auto">
                  <table className="w-full text-xs text-left">
                    <thead className="bg-gray-900 text-white font-semibold">
                      <tr>
                        <th className="p-3 w-10 text-center">#</th>
                        <th className="p-3">Party Name</th>
                        <th className="p-3">Type</th>
                        <th className="p-3 text-right">Books Balance (₹)</th>
                        <th className="p-3 text-right">Confirmed Bal (₹)</th>
                        <th className="p-3 text-right">Variance (₹)</th>
                        <th className="p-3 text-center">Status</th>
                        <th className="p-3 text-center">Direct Response?</th>
                        <th className="p-3">Auditor Notes & Actions</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-[#E5E7EB]">
                      {filteredParties.map((p, idx) => (
                        <tr
                          key={p.id}
                          className={`hover:bg-gray-50 transition-colors ${
                            p.status === 'difference_reported'
                              ? 'bg-red-50/30'
                              : p.isOverdueFollowup
                              ? 'bg-amber-50/20'
                              : ''
                          }`}
                        >
                          <td className="p-3 text-center text-gray-600 font-medium">{idx + 1}</td>
                          <td className="p-3">
                            <div className="font-bold text-gray-900">{p.partyName}</div>
                            <div className="text-[11px] text-gray-600">
                              {p.contactPerson || p.emailOrPhone || p.accountCode || '-'}
                            </div>
                          </td>
                          <td className="p-3 font-medium text-gray-900">{p.confirmationType}</td>
                          <td className="p-3 text-right font-mono font-bold">
                            ₹{p.expectedBalance.toLocaleString('en-IN')}
                          </td>
                          <td className="p-3 text-right font-mono">
                            <input
                              type="number"
                              value={p.confirmedBalance ?? ''}
                              onChange={(e) => {
                                const val = e.target.value === '' ? null : parseFloat(e.target.value);
                                handleUpdateParty(p.id, {
                                  confirmedBalance: val,
                                  status: val !== null && val !== p.expectedBalance ? 'difference_reported' : val !== null ? 'response_received' : p.status,
                                });
                              }}
                              placeholder="Awaiting"
                              className="w-24 text-right border border-gray-200 rounded px-1.5 py-0.5 text-xs outline-none focus:border-gray-300"
                            />
                          </td>
                          <td className="p-3 text-right font-mono font-bold">
                            {p.differenceAmount !== null && p.differenceAmount !== undefined ? (
                              p.differenceAmount !== 0 ? (
                                <span className="text-[#DC2626]">₹{p.differenceAmount.toLocaleString('en-IN')}</span>
                              ) : (
                                <span className="text-gray-900">₹0.00</span>
                              )
                            ) : (
                              <span className="text-gray-600">-</span>
                            )}
                          </td>
                          <td className="p-3 text-center">
                            <select
                              value={p.status}
                              onChange={(e) => handleUpdateParty(p.id, { status: e.target.value as ConfirmationStatus })}
                              className="text-[11px] border rounded px-1.5 py-0.5 outline-none font-semibold"
                            >
                              <option value="not_prepared">Not Prepared</option>
                              <option value="ready_to_send">Ready to Send</option>
                              <option value="sent">Sent</option>
                              <option value="response_received">Response Received</option>
                              <option value="difference_reported">Difference Reported</option>
                              <option value="followup_required">Follow-up Required</option>
                              <option value="closed">Closed</option>
                            </select>
                          </td>
                          <td className="p-3 text-center">
                            <input
                              type="checkbox"
                              checked={p.isDirectResponse}
                              onChange={(e) => handleUpdateParty(p.id, { isDirectResponse: e.target.checked })}
                              className="rounded text-gray-900 focus:ring-gray-900"
                            />
                          </td>
                          <td className="p-3">
                            <div className="flex items-center gap-2">
                              <input
                                type="text"
                                value={p.notes}
                                onChange={(e) => handleUpdateParty(p.id, { notes: e.target.value })}
                                placeholder="Add discrepancy note..."
                                className="w-full text-xs border border-transparent hover:border-gray-200 rounded px-1.5 py-0.5 outline-none bg-transparent"
                              />
                              <button
                                onClick={() => {
                                  setSelectedPartyId(p.id);
                                  setActiveTab('letter_preview');
                                }}
                                className="text-[10px] text-gray-900 font-bold hover:underline shrink-0"
                              >
                                View Letter
                              </button>
                            </div>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            )}

            {/* TAB 2: Differences & Action Items */}
            {activeTab === 'differences' && (
              <div className="bg-white border border-gray-200 rounded-xl p-6 shadow-sm space-y-4">
                <div>
                  <h3 className="text-sm font-bold text-gray-900">Reported Balance Discrepancies (SA 505 Action Schedule)</h3>
                  <p className="text-xs text-gray-600">
                    Itemized list of parties where confirmed balances diverge from book balances
                  </p>
                </div>

                <div className="overflow-x-auto">
                  <table className="w-full text-xs text-left">
                    <thead className="bg-gray-900 text-white font-semibold">
                      <tr>
                        <th className="p-3">Party Name</th>
                        <th className="p-3">Account Type</th>
                        <th className="p-3 text-right">Books Balance (₹)</th>
                        <th className="p-3 text-right">Confirmed Balance (₹)</th>
                        <th className="p-3 text-right">Discrepancy (₹)</th>
                        <th className="p-3">Audit Observation & Action Plan</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-[#E5E7EB]">
                      {summary.parties
                        .filter((p) => p.differenceAmount !== null && p.differenceAmount !== 0)
                        .map((d) => (
                          <tr key={d.id} className="hover:bg-gray-50">
                            <td className="p-3 font-bold text-gray-900">{d.partyName}</td>
                            <td className="p-3 text-gray-900 font-semibold">{d.confirmationType}</td>
                            <td className="p-3 text-right font-mono">₹{d.expectedBalance.toLocaleString('en-IN')}</td>
                            <td className="p-3 text-right font-mono">₹{(d.confirmedBalance || 0).toLocaleString('en-IN')}</td>
                            <td className="p-3 text-right font-mono font-bold text-[#DC2626]">
                              ₹{(d.differenceAmount || 0).toLocaleString('en-IN')}
                            </td>
                            <td className="p-3 text-[11px] text-gray-900">
                              {d.notes || 'In-transit payments / credit notes under reconciliation.'}
                            </td>
                          </tr>
                        ))}
                    </tbody>
                  </table>
                </div>
              </div>
            )}

            {/* TAB 3: Draft Confirmation Letter Preview */}
            {activeTab === 'letter_preview' && selectedParty && (
              <div className="space-y-4">
                {/* Party selector bar */}
                <div className="flex items-center justify-between bg-white border border-gray-200 rounded-xl p-3 print:hidden">
                  <div className="flex items-center gap-2 text-xs">
                    <span className="font-semibold text-gray-600">Previewing Confirmation Letter for:</span>
                    <select
                      value={selectedPartyId}
                      onChange={(e) => setSelectedPartyId(e.target.value)}
                      className="border border-gray-200 rounded px-2.5 py-1 text-xs font-bold text-gray-900 outline-none"
                    >
                      {parties.map((p) => (
                        <option key={p.id} value={p.id}>
                          {p.partyName} ({p.confirmationType}) - ₹{p.expectedBalance.toLocaleString('en-IN')}
                        </option>
                      ))}
                    </select>
                  </div>

                  <button
                    onClick={() => window.print()}
                    className="px-3 py-1.5 bg-gray-900 text-white rounded text-xs font-semibold hover:bg-black flex items-center gap-1.5"
                  >
                    <Printer className="w-3.5 h-3.5" />
                    Print Letter
                  </button>
                </div>

                {/* Formal Letter Template */}
                <div className="bg-white border border-gray-200 rounded-xl p-10 shadow-sm print:border-none print:shadow-none print:p-0 font-serif leading-relaxed text-xs text-gray-900 whitespace-pre-wrap">
                  {generateConfirmationLetterText(selectedParty, clientName, auditorFirm, auditorEmail)}
                </div>
              </div>
            )}
          </div>
        )}
      </div>
    </div>
  );
};