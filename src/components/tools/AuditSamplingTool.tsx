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
  Filter,
  Printer,
  Calendar,
  Layers,
  HelpCircle,
  Hash,
} from 'lucide-react';
import { InnerToolBadge } from '../common/InnerToolBadge';
import {
  type SamplingMethod,
  type SamplingConfig,
  type SamplingResult,
  type SamplingColumnMapping,
  detectSamplingColumns,
  scrutinizePopulation,
  executeAuditSampling,
  exportSamplingWorkbook,
} from '../../utils/auditSamplingEngine';

interface AuditSamplingToolProps {
  onBack: () => void;
  onShowToast: (title: string, description?: string, type?: 'success' | 'error' | 'info') => void;
}

export const AuditSamplingTool: React.FC<AuditSamplingToolProps> = ({
  onBack,
  onShowToast,
}) => {
  const [rawRows, setRawRows] = useState<Record<string, any>[]>([]);
  const [headers, setHeaders] = useState<string[]>([]);
  const [fileName, setFileName] = useState('');
  const [clientName, setClientName] = useState('');
  const [auditPeriod, setAuditPeriod] = useState('FY 2026-27');
  const [populationName, setPopulationName] = useState('Voucher Expenses Population');

  const [mapping, setMapping] = useState<SamplingColumnMapping>({
    idCol: '',
    dateCol: '',
    partyCol: '',
    amountCol: '',
    headCol: '',
    narrationCol: '',
  });

  const [config, setConfig] = useState<SamplingConfig>({
    method: 'high_value_targeted',
    sampleSize: 5,
    randomSeed: 4289,
    monetaryThreshold: 500000,
    strataBoundaries: [50000, 200000],
    strataSampleSizes: [2, 3, 4],
  });

  const [activeTab, setActiveTab] = useState<'sample_register' | 'population' | 'methodology' | 'printable_summary'>('sample_register');
  const [searchQuery, setSearchQuery] = useState('');
  const [showConfigPanel, setShowConfigPanel] = useState(false);
  const [showMappingPanel, setShowMappingPanel] = useState(false);

  const fileInputRef = useRef<HTMLInputElement>(null);

  // Scrutinize population & execute sampling
  const samplingResult: SamplingResult = useMemo(() => {
    if (rawRows.length === 0) {
      const emptyScrutiny = {
        totalCount: 0,
        totalValue: 0,
        averageValue: 0,
        maxValue: 0,
        minValue: 0,
        missingIdCount: 0,
        duplicateIdCount: 0,
        invalidAmountCount: 0,
        negativeAmountCount: 0,
        records: [],
      };
      return {
        clientName,
        auditPeriod,
        populationName,
        config,
        executionTimestamp: new Date().toISOString(),
        populationScrutiny: emptyScrutiny,
        sampledRecords: [],
        totalSampledValue: 0,
        coveragePercentage: 0,
        isReproducible: true,
      };
    }

    const scrutiny = scrutinizePopulation(rawRows, mapping);
    return executeAuditSampling(scrutiny, config, clientName, auditPeriod, populationName);
  }, [rawRows, mapping, config, clientName, auditPeriod, populationName]);

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
        setRawRows(json);
        setMapping(detectSamplingColumns(detectedHeaders));
      } catch (err: any) {
        onShowToast('Upload Error', err.message || 'Failed to parse file', 'error');
      }
    };

    reader.readAsArrayBuffer(file);
    e.target.value = '';
  };

  // Export Excel sampling register
  const handleExportExcel = () => {
    if (samplingResult.sampledRecords.length === 0) {
      onShowToast('No Sample Selected', 'Please load population and generate sample first.', 'error');
      return;
    }

    try {
      const wb = exportSamplingWorkbook(samplingResult);
      XLSX.writeFile(wb, `Audit_Sampling_Register_${clientName.replace(/\s+/g, '_')}_${auditPeriod}.xlsx`);
      onShowToast('Workbook Exported', 'Exported 2-sheet audit sampling register with blank testing columns.', 'success');
    } catch (err: any) {
      onShowToast('Export Error', err.message || 'Failed to generate Excel file', 'error');
    }
  };

  // Print PDF summary
  const handlePrint = () => {
    if (samplingResult.sampledRecords.length === 0) {
      onShowToast('No Sample Selected', 'Please generate sample first.', 'error');
      return;
    }
    setActiveTab('printable_summary');
    setTimeout(() => {
      window.print();
    }, 300);
  };

  // Filtered sampled records
  const filteredSampledRecords = useMemo(() => {
    if (!searchQuery.trim()) return samplingResult.sampledRecords;
    const q = searchQuery.toLowerCase();
    return samplingResult.sampledRecords.filter((r) => {
      const id = r.transactionId.toLowerCase();
      const party = r.partyName.toLowerCase();
      const head = r.accountHead.toLowerCase();
      return id.includes(q) || party.includes(q) || head.includes(q);
    });
  }, [samplingResult.sampledRecords, searchQuery]);

  return (
    <div className="space-y-6">
      <ToolHeader
        toolId="audit-sampling"
        title="Audit Sampling Tool"
        description="Select reproducible transaction samples, calculate monetary coverage, and export blank audit vouching registers."
        badge="SA Compliant"
        onBack={onBack}
        actionButtons={
          <div className="flex flex-wrap items-center gap-2">
            

            <button
              onClick={() => fileInputRef.current?.click()}
              className="px-3 py-1.5 text-xs font-semibold rounded-md bg-white text-gray-900 border border-gray-200 hover:bg-gray-50 transition-colors flex items-center gap-1.5 shadow-sm"
            >
              <UploadCloud className="w-3.5 h-3.5 text-gray-900" />
              {fileName ? 'Re-upload File' : 'Import Population'}
            </button>
            <input
              type="file"
              ref={fileInputRef}
              onChange={handleFileUpload}
              accept=".xlsx,.xls,.csv"
              className="hidden"
              id="audit-sampling-upload"
            />

            <button
              onClick={() => setShowConfigPanel(!showConfigPanel)}
              className={`px-3 py-1.5 text-xs font-semibold rounded-md border transition-colors flex items-center gap-1.5 ${
                showConfigPanel
                  ? 'bg-gray-900 text-white border-gray-300'
                  : 'bg-white text-gray-900 border-gray-200 hover:bg-gray-50'
              }`}
            >
              <Sliders className="w-3.5 h-3.5" />
              Sampling Parameters
            </button>

            <button
              onClick={() => setShowMappingPanel(!showMappingPanel)}
              className={`px-3 py-1.5 text-xs font-semibold rounded-md border transition-colors flex items-center gap-1.5 ${
                showMappingPanel
                  ? 'bg-gray-900 text-white border-gray-300'
                  : 'bg-white text-gray-900 border-gray-200 hover:bg-gray-50'
              }`}
            >
              Map Fields
            </button>

            <button
              onClick={handlePrint}
              disabled={samplingResult.sampledRecords.length === 0}
              className="px-3 py-1.5 text-xs font-semibold rounded-md bg-white text-gray-900 border border-gray-200 hover:bg-gray-50 disabled:opacity-50 transition-colors flex items-center gap-1.5"
            >
              <Printer className="w-3.5 h-3.5 text-gray-600" />
              PDF Summary
            </button>

            <button
              onClick={handleExportExcel}
              disabled={samplingResult.sampledRecords.length === 0}
              className="px-3.5 py-1.5 text-xs font-semibold rounded-md bg-gray-900 text-white hover:bg-black disabled:opacity-50 transition-colors flex items-center gap-1.5 shadow-sm"
            >
              <FileSpreadsheet className="w-3.5 h-3.5" />
              Export Sampling Register (.xlsx)
            </button>
          </div>
        }
      />

      {/* Client & Period Header */}
      <div className="bg-white border border-gray-200 rounded-xl p-4 shadow-2xs flex flex-wrap items-center gap-4 text-xs">
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
            <span className="font-semibold text-gray-600">Audit Period:</span>
            <input
              type="text"
              value={auditPeriod}
              onChange={(e) => setAuditPeriod(e.target.value)}
              className="px-2 py-1 border border-gray-200 rounded text-xs font-medium focus:border-gray-300 outline-none w-36"
            />
          </div>

          {fileName && (
            <div className="ml-auto text-gray-600">
              Source: <strong className="text-gray-900">{fileName}</strong> ({rawRows.length} items)
            </div>
          )}
      </div>

      {/* Upload Dropzone when no file selected */}
      {!fileName && (
        <div className="bg-white border-2 border-dashed border-gray-200 hover:border-gray-300 rounded-xl p-8 sm:p-10 text-center transition-colors shadow-2xs">
          <label htmlFor="audit-sampling-upload" className="cursor-pointer flex flex-col items-center">
            <div className="w-12 h-12 rounded-full bg-gray-100 border border-gray-200 flex items-center justify-center text-gray-900 mx-auto mb-3 shadow-2xs">
              <UploadCloud className="w-6 h-6" />
            </div>
            <h3 className="text-sm font-semibold text-gray-900">Upload Audit Population Spreadsheet (.xlsx, .xls, .csv)</h3>
            <p className="text-xs text-gray-500 mt-1 max-w-md mx-auto leading-relaxed">
              Upload client voucher listings, sales ledgers, or expense ledgers. Automatically draws statistically valid or MUS audit samples adhering to SA 530 standards.
            </p>
            <div className="mt-4 px-4 py-2 bg-gray-900 hover:bg-black text-white text-xs font-semibold rounded-lg shadow-2xs inline-flex items-center gap-2 transition-all">
              <UploadCloud className="w-4 h-4" />
              <span>Import Population File</span>
            </div>
            <span className="text-[11px] text-gray-400 font-medium mt-3 block">
              Supports .xlsx, .xls, .csv · SA 530 Compliant · Stratified, Systematic, MUS & Random sampling
            </span>
          </label>
        </div>
      )}

      {/* Main Container */}
      <div className="space-y-6">
        {/* Sampling Configuration Drawer */}
        {showConfigPanel && (
          <div className="bg-white border border-gray-200 rounded-xl p-5 shadow-sm space-y-4">
            <div className="flex items-center justify-between pb-3 border-b border-gray-200">
              <div>
                <h3 className="text-sm font-bold text-gray-900">Sampling Methodology & Parameters (SA 530)</h3>
                <p className="text-xs text-gray-600">Select statistical or targeted sampling method and configure sample constraints</p>
              </div>
              <button onClick={() => setShowConfigPanel(false)} className="text-xs text-gray-600 hover:text-gray-900">Close</button>
            </div>

            <div className="grid grid-cols-1 md:grid-cols-4 gap-4 text-xs">
              <div>
                <label className="block font-semibold text-gray-600 mb-1">Sampling Method</label>
                <select
                  value={config.method}
                  onChange={(e) => setConfig({ ...config, method: e.target.value as SamplingMethod })}
                  className="w-full border border-gray-200 rounded p-2 focus:border-gray-300 outline-none font-medium"
                >
                  <option value="high_value_targeted">Targeted High-Value (&gt; Threshold)</option>
                  <option value="simple_random">Simple Random Sampling (SRS)</option>
                  <option value="systematic">Systematic Sampling (Every kth)</option>
                  <option value="stratified">Stratified Sampling (By Amount)</option>
                </select>
              </div>

              <div>
                <label className="block font-semibold text-gray-600 mb-1">
                  {config.method === 'high_value_targeted' ? 'Monetary Threshold (₹)' : 'Sample Size (n)'}
                </label>
                {config.method === 'high_value_targeted' ? (
                  <input
                    type="number"
                    value={config.monetaryThreshold}
                    onChange={(e) => setConfig({ ...config, monetaryThreshold: parseFloat(e.target.value) || 0 })}
                    className="w-full border border-gray-200 rounded p-2 focus:border-gray-300 outline-none font-mono"
                  />
                ) : (
                  <input
                    type="number"
                    value={config.sampleSize}
                    onChange={(e) => setConfig({ ...config, sampleSize: parseInt(e.target.value) || 1 })}
                    className="w-full border border-gray-200 rounded p-2 focus:border-gray-300 outline-none font-mono"
                  />
                )}
              </div>

              <div>
                <label className="block font-semibold text-gray-600 mb-1">Random Seed (For Reproducibility)</label>
                <input
                  type="number"
                  value={config.randomSeed}
                  onChange={(e) => setConfig({ ...config, randomSeed: parseInt(e.target.value) || 123 })}
                  className="w-full border border-gray-200 rounded p-2 focus:border-gray-300 outline-none font-mono"
                />
                <span className="text-[11px] text-gray-600 mt-1 block">Fixed seed enables identical sample re-selection</span>
              </div>

              <div className="flex items-center pt-5">
                <span className="text-[11px] text-gray-600 leading-relaxed">
                  ICAI Standard on Auditing (SA 530): Requires sampling risk evaluation and representative stratification.
                </span>
              </div>
            </div>
          </div>
        )}

        {/* Mapping Drawer */}
        {showMappingPanel && (
          <div className="bg-white border border-gray-200 rounded-xl p-5 shadow-sm space-y-4">
            <div className="flex items-center justify-between pb-3 border-b border-gray-200">
              <div>
                <h3 className="text-sm font-bold text-gray-900">Field Mapping: Transaction Population Data</h3>
                <p className="text-xs text-gray-600">Map voucher numbers, amounts, dates, and accounts</p>
              </div>
              <button onClick={() => setShowMappingPanel(false)} className="text-xs text-gray-600 hover:text-gray-900">Close</button>
            </div>

            <div className="grid grid-cols-2 md:grid-cols-4 gap-4 text-xs">
              <div>
                <label className="block font-semibold text-gray-600 mb-1">Voucher / ID Column *</label>
                <select
                  value={mapping.idCol}
                  onChange={(e) => setMapping({ ...mapping, idCol: e.target.value })}
                  className="w-full border rounded p-1.5 outline-none"
                >
                  <option value="">-- Select --</option>
                  {headers.map((h) => <option key={h} value={h}>{h}</option>)}
                </select>
              </div>

              <div>
                <label className="block font-semibold text-gray-600 mb-1">Amount Column *</label>
                <select
                  value={mapping.amountCol}
                  onChange={(e) => setMapping({ ...mapping, amountCol: e.target.value })}
                  className="w-full border rounded p-1.5 outline-none"
                >
                  <option value="">-- Select --</option>
                  {headers.map((h) => <option key={h} value={h}>{h}</option>)}
                </select>
              </div>

              <div>
                <label className="block font-semibold text-gray-600 mb-1">Party / Payee Column</label>
                <select
                  value={mapping.partyCol}
                  onChange={(e) => setMapping({ ...mapping, partyCol: e.target.value })}
                  className="w-full border rounded p-1.5 outline-none"
                >
                  <option value="">-- Select --</option>
                  {headers.map((h) => <option key={h} value={h}>{h}</option>)}
                </select>
              </div>

              <div>
                <label className="block font-semibold text-gray-600 mb-1">Date Column</label>
                <select
                  value={mapping.dateCol}
                  onChange={(e) => setMapping({ ...mapping, dateCol: e.target.value })}
                  className="w-full border rounded p-1.5 outline-none"
                >
                  <option value="">-- Select --</option>
                  {headers.map((h) => <option key={h} value={h}>{h}</option>)}
                </select>
              </div>
            </div>
          </div>
        )}

        {/* Empty State */}
        {rawRows.length === 0 ? (
          <div className="bg-white border border-dashed border-gray-200 rounded-2xl p-12 text-center shadow-sm">
            <div className="w-16 h-16 bg-gray-50 rounded-full flex items-center justify-center mx-auto mb-4 border border-gray-300/20">
              <Layers className="w-8 h-8 text-gray-900" />
            </div>
            <h3 className="text-base font-bold text-gray-900">No Audit Population Loaded</h3>
            <p className="text-xs text-gray-600 max-w-md mx-auto mt-1 mb-6">
              Upload your client's transaction register (Excel or CSV) to scrutinize population integrity and draw statistical or high-value targeted samples.
            </p>
            <div className="flex items-center justify-center gap-3">
              
              <button
                onClick={() => fileInputRef.current?.click()}
                className="px-4 py-2 text-xs font-semibold rounded-lg bg-gray-900 text-white hover:bg-black transition-colors shadow-sm inline-flex items-center gap-2"
              >
                <UploadCloud className="w-4 h-4" />
                <span>Upload Transaction Population</span>
              </button>
            </div>
          </div>
        ) : (
          <div className="space-y-6">
            {/* KPI Summary Cards */}
            <div className="grid grid-cols-2 md:grid-cols-5 gap-4">
              <div className="bg-white border border-gray-200 rounded-xl p-4 shadow-sm">
                <span className="text-xs font-semibold text-gray-600 uppercase">Population Size (N)</span>
                <div className="text-xl font-bold text-gray-900 mt-1">
                  {samplingResult.populationScrutiny.totalCount} items
                </div>
                <div className="text-xs text-gray-600 mt-0.5">
                  Total ₹{samplingResult.populationScrutiny.totalValue.toLocaleString('en-IN')}
                </div>
              </div>

              <div className="bg-white border border-gray-200 rounded-xl p-4 shadow-sm">
                <span className="text-xs font-semibold text-gray-600 uppercase">Sample Size (n)</span>
                <div className="text-xl font-bold text-gray-900 mt-1">
                  {samplingResult.sampledRecords.length} items
                </div>
                <div className="text-xs text-gray-600 mt-0.5">
                  Sampled ₹{samplingResult.totalSampledValue.toLocaleString('en-IN')}
                </div>
              </div>

              <div className="bg-white border border-gray-200 rounded-xl p-4 shadow-sm">
                <span className="text-xs font-semibold text-gray-600 uppercase">Monetary Coverage</span>
                <div className="text-xl font-bold text-gray-900 mt-1">
                  {samplingResult.coveragePercentage.toFixed(1)}%
                </div>
                <div className="text-xs text-gray-600 mt-0.5">Tested value proportion</div>
              </div>

              <div className="bg-white border border-gray-200 rounded-xl p-4 shadow-sm">
                <span className="text-xs font-semibold text-gray-600 uppercase">Sampling Method</span>
                <div className="text-sm font-bold text-gray-900 mt-1 truncate">
                  {config.method.replace(/_/g, ' ').toUpperCase()}
                </div>
                <div className="text-xs text-gray-600 mt-0.5">Seed: {config.randomSeed}</div>
              </div>

              <div className="bg-white border border-gray-200 rounded-xl p-4 shadow-sm">
                <span className="text-xs font-semibold text-gray-600 uppercase">Population Scrutiny</span>
                <div className="flex items-center gap-1.5 mt-1">
                  {samplingResult.populationScrutiny.duplicateIdCount === 0 &&
                  samplingResult.populationScrutiny.missingIdCount === 0 ? (
                    <span className="px-2 py-0.5 text-xs font-bold rounded bg-[#F3F4F6] text-gray-900">
                      ✓ Clean Data
                    </span>
                  ) : (
                    <span className="px-2 py-0.5 text-xs font-bold rounded bg-[#FEE2E2] text-[#DC2626]">
                      {samplingResult.populationScrutiny.duplicateIdCount} Duplicates
                    </span>
                  )}
                </div>
                <div className="text-xs text-gray-600 mt-0.5">
                  Max: ₹{samplingResult.populationScrutiny.maxValue.toLocaleString('en-IN')}
                </div>
              </div>
            </div>

            {/* Tabs Bar */}
            <div className="flex flex-col sm:flex-row sm:items-center sm:justify-between gap-4 border-b border-gray-200 pb-1 print:hidden">
              <div className="flex items-center gap-2">
                <button
                  onClick={() => setActiveTab('sample_register')}
                  className={`px-4 py-2 text-xs font-bold border-b-2 transition-colors ${
                    activeTab === 'sample_register'
                      ? 'border-gray-300 text-gray-900'
                      : 'border-transparent text-gray-600 hover:text-gray-900'
                  }`}
                >
                  Selected Sample Register ({samplingResult.sampledRecords.length})
                </button>
                <button
                  onClick={() => setActiveTab('population')}
                  className={`px-4 py-2 text-xs font-bold border-b-2 transition-colors ${
                    activeTab === 'population'
                      ? 'border-gray-300 text-gray-900'
                      : 'border-transparent text-gray-600 hover:text-gray-900'
                  }`}
                >
                  Full Population ({samplingResult.populationScrutiny.totalCount})
                </button>
                <button
                  onClick={() => setActiveTab('methodology')}
                  className={`px-4 py-2 text-xs font-bold border-b-2 transition-colors ${
                    activeTab === 'methodology'
                      ? 'border-gray-300 text-gray-900'
                      : 'border-transparent text-gray-600 hover:text-gray-900'
                  }`}
                >
                  Methodology & SA 530
                </button>
                <button
                  onClick={() => setActiveTab('printable_summary')}
                  className={`px-4 py-2 text-xs font-bold border-b-2 transition-colors ${
                    activeTab === 'printable_summary'
                      ? 'border-gray-300 text-gray-900'
                      : 'border-transparent text-gray-600 hover:text-gray-900'
                  }`}
                >
                  Printable Audit Memorandum
                </button>
              </div>

              {activeTab === 'sample_register' && (
                <div className="relative">
                  <Search className="w-3.5 h-3.5 absolute left-2.5 top-1/2 -translate-y-1/2 text-gray-600" />
                  <input
                    type="text"
                    placeholder="Search voucher or payee..."
                    value={searchQuery}
                    onChange={(e) => setSearchQuery(e.target.value)}
                    className="pl-8 pr-3 py-1.5 text-xs border border-gray-200 rounded-lg w-52 outline-none focus:border-gray-300"
                  />
                </div>
              )}
            </div>

            {/* TAB 1: Sample Register Table */}
            {activeTab === 'sample_register' && (
              <div className="bg-white border border-gray-200 rounded-xl overflow-hidden shadow-sm">
                <div className="overflow-x-auto">
                  <table className="w-full text-xs text-left">
                    <thead className="bg-gray-900 text-white font-semibold">
                      <tr>
                        <th className="p-3 w-10 text-center">Sample #</th>
                        <th className="p-3">Source Row</th>
                        <th className="p-3">Voucher #</th>
                        <th className="p-3">Date</th>
                        <th className="p-3">Party / Payee</th>
                        <th className="p-3">Account Head</th>
                        <th className="p-3 text-right">Amount (₹)</th>
                        <th className="p-3">Classification</th>
                        <th className="p-3">Sampling Rationale</th>
                        <th className="p-3 text-center">Testing Result</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-[#E5E7EB]">
                      {filteredSampledRecords.map((s, idx) => (
                        <tr key={s.id} className="hover:bg-gray-50 transition-colors">
                          <td className="p-3 text-center font-bold text-gray-900">{idx + 1}</td>
                          <td className="p-3 text-gray-600 font-medium text-center">Row {s.sourceRowIndex}</td>
                          <td className="p-3 font-mono font-bold text-gray-900">{s.transactionId}</td>
                          <td className="p-3">{s.date}</td>
                          <td className="p-3 font-semibold text-gray-900">{s.partyName}</td>
                          <td className="p-3 text-gray-600">{s.accountHead}</td>
                          <td className="p-3 text-right font-mono font-bold text-sm">
                            ₹{s.amount.toLocaleString('en-IN')}
                          </td>
                          <td className="p-3">
                            <span
                              className={`px-2 py-0.5 rounded text-[10px] font-bold ${
                                s.sampleCategory === 'High-Value Targeted'
                                  ? 'bg-purple-100 text-purple-800'
                                  : 'bg-gray-100 text-gray-900'
                              }`}
                            >
                              {s.sampleCategory.toUpperCase()}
                            </span>
                          </td>
                          <td className="p-3 text-[11px] text-gray-600 max-w-xs">{s.samplingRationale}</td>
                          <td className="p-3 text-center">
                            <span className="px-2 py-0.5 rounded text-[10px] bg-gray-100 text-gray-700 font-medium">
                              Pending Vouching
                            </span>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            )}

            {/* TAB 2: Full Population */}
            {activeTab === 'population' && (
              <div className="bg-white border border-gray-200 rounded-xl overflow-hidden shadow-sm">
                <div className="p-3 bg-gray-50 border-b border-gray-200 flex justify-between text-xs text-gray-600">
                  <span>Complete population records ({samplingResult.populationScrutiny.totalCount} items)</span>
                  <span>Average Amount: <strong>₹{samplingResult.populationScrutiny.averageValue.toFixed(2)}</strong></span>
                </div>
                <div className="overflow-x-auto max-h-[500px]">
                  <table className="w-full text-xs text-left">
                    <thead className="bg-gray-100 text-gray-900 sticky top-0 font-semibold">
                      <tr>
                        <th className="p-3 w-10 text-center">Row</th>
                        <th className="p-3">Voucher #</th>
                        <th className="p-3">Date</th>
                        <th className="p-3">Party Name</th>
                        <th className="p-3">Account Head</th>
                        <th className="p-3 text-right">Amount (₹)</th>
                        <th className="p-3">Narration</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-[#E5E7EB]">
                      {samplingResult.populationScrutiny.records.map((r) => (
                        <tr key={r.id} className="hover:bg-gray-50">
                          <td className="p-3 text-center text-gray-600">{r.sourceRowIndex}</td>
                          <td className="p-3 font-mono font-bold text-gray-900">{r.transactionId}</td>
                          <td className="p-3">{r.date}</td>
                          <td className="p-3">{r.partyName}</td>
                          <td className="p-3 text-gray-600">{r.accountHead}</td>
                          <td className="p-3 text-right font-mono font-semibold">₹{r.amount.toLocaleString('en-IN')}</td>
                          <td className="p-3 text-[11px] text-gray-600 max-w-xs truncate">{r.narration}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            )}

            {/* TAB 3: Methodology */}
            {activeTab === 'methodology' && (
              <div className="bg-white border border-gray-200 rounded-xl p-6 shadow-sm space-y-4 text-xs">
                <div>
                  <h3 className="font-bold text-sm text-gray-900">Standard on Auditing (SA) 530 Compliance Notes</h3>
                  <p className="text-xs text-gray-600">Audit documentation of sampling risk, sample design, and sample selection</p>
                </div>

                <div className="p-4 bg-gray-50 border border-gray-300/30 rounded-lg text-gray-900 leading-relaxed">
                  <strong>Auditor Sampling Principle:</strong> When designing an audit sample, the auditor shall consider the purpose of the audit procedure and the characteristics of the population from which the sample will be drawn. Selection of high-value items does not constitute statistical sampling, as items below the threshold have a zero probability of selection; it represents 100% examination of a specific stratum.
                </div>

                <div className="border border-gray-200 rounded-lg p-4 space-y-2">
                  <div className="flex justify-between py-1 border-b">
                    <span className="text-gray-600">Selection Method:</span>
                    <span className="font-bold">{config.method.replace(/_/g, ' ').toUpperCase()}</span>
                  </div>
                  <div className="flex justify-between py-1 border-b">
                    <span className="text-gray-600">Reproducibility Seed:</span>
                    <span className="font-mono font-bold">{config.randomSeed}</span>
                  </div>
                  <div className="flex justify-between py-1 border-b">
                    <span className="text-gray-600">Monetary Coverage:</span>
                    <span className="font-bold text-gray-900">{samplingResult.coveragePercentage.toFixed(1)}%</span>
                  </div>
                  <div className="flex justify-between py-1">
                    <span className="text-gray-600">Execution Timestamp:</span>
                    <span className="font-mono">{samplingResult.executionTimestamp}</span>
                  </div>
                </div>
              </div>
            )}

            {/* TAB 4: Printable Summary */}
            {activeTab === 'printable_summary' && (
              <div className="bg-white border border-gray-200 rounded-xl p-8 shadow-sm print:border-none print:shadow-none print:p-0">
                <div className="border-b-2 border-gray-300 pb-4 mb-6">
                  <div className="flex justify-between items-start">
                    <div>
                      <h2 className="text-xl font-bold text-gray-900">{clientName}</h2>
                      <h3 className="text-base font-semibold text-gray-900">AUDIT SAMPLING WORKING PAPER (SA 530)</h3>
                      <p className="text-xs text-gray-600">Scope: {populationName} • Audit Period: {auditPeriod}</p>
                    </div>
                    <div className="text-right text-xs text-gray-600">
                      <div>Method: {config.method.toUpperCase()}</div>
                      <div>Coverage: {samplingResult.coveragePercentage.toFixed(1)}%</div>
                      <div>Date: {new Date().toLocaleDateString('en-IN')}</div>
                    </div>
                  </div>
                </div>

                <table className="w-full text-xs text-left mb-6 border-collapse">
                  <thead>
                    <tr className="bg-gray-900 text-white">
                      <th className="p-2 border border-gray-300">#</th>
                      <th className="p-2 border border-gray-300">Voucher #</th>
                      <th className="p-2 border border-gray-300">Date</th>
                      <th className="p-2 border border-gray-300">Party</th>
                      <th className="p-2 border border-gray-300 text-right">Amount (₹)</th>
                      <th className="p-2 border border-gray-300">Selection Basis</th>
                      <th className="p-2 border border-gray-300">Vouching Done?</th>
                    </tr>
                  </thead>
                  <tbody>
                    {samplingResult.sampledRecords.map((s, idx) => (
                      <tr key={s.id}>
                        <td className="p-2 border border-gray-300 text-center">{idx + 1}</td>
                        <td className="p-2 border border-gray-300 font-mono font-bold">{s.transactionId}</td>
                        <td className="p-2 border border-gray-300">{s.date}</td>
                        <td className="p-2 border border-gray-300">{s.partyName}</td>
                        <td className="p-2 border border-gray-300 text-right font-mono">₹{s.amount.toLocaleString('en-IN')}</td>
                        <td className="p-2 border border-gray-300 text-[10px]">{s.samplingRationale}</td>
                        <td className="p-2 border border-gray-300 text-center font-mono">[   ] Pass / [   ] Fail</td>
                      </tr>
                    ))}
                  </tbody>
                </table>

                <div className="mt-8 pt-4 border-t border-gray-200 text-[11px] text-gray-600 leading-relaxed">
                  <strong>Statutory Disclaimer:</strong> Sampling risk is the risk that the auditor's conclusion based on a sample may be different from the conclusion if the entire population were subjected to the same audit procedure. Selection of this sample does not guarantee detection of all irregularities.
                </div>
              </div>
            )}
          </div>
        )}
      </div>
    </div>
  );
};