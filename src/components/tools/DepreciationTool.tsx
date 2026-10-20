import React, { useState, useMemo } from 'react';
import * as XLSX from 'xlsx';
import {
  ArrowLeft,
  Download,
  TrendingDown,
  Plus,
  Printer,
  Sparkles,
  Calculator,
  Trash2,
  AlertTriangle,
  Scale
} from 'lucide-react';
import {
  type AssetRecord,
  type AssetCategory,
  ASSET_CATEGORIES,
  calculateAssetDepreciation,
  exportDepreciationWorkbook,
} from '../../utils/depreciationEngine';
import { DataFormatGuide } from '../common/DataFormatGuide';
import { ToolHeader } from '../common/ToolHeader';

interface DepreciationToolProps {
  onBack: () => void;
  onShowToast: (title: string, description?: string, type?: 'success' | 'error' | 'info') => void;
}

export const DepreciationTool: React.FC<DepreciationToolProps> = ({ onBack, onShowToast }) => {
  const [assets, setAssets] = useState<AssetRecord[]>([]);
  const [activeTab, setActiveTab] = useState<'co_act' | 'it_act' | 'deferred_tax'>('co_act');
  const [entityName, setEntityName] = useState('');
  const [financialYear, setFinancialYear] = useState('FY 2026-27');
  const [showAddModal, setShowAddModal] = useState(false);

  // New Asset Form State
  const [newCode, setNewCode] = useState('');
  const [newName, setNewName] = useState('');
  const [newCategory, setNewCategory] = useState<AssetCategory>('plant_machinery');
  const [newCost, setNewCost] = useState('500000');
  const [newOpeningWdv, setNewOpeningWdv] = useState('400000');
  const [newAddMore180, setNewAddMore180] = useState('0');
  const [newAddLess180, setNewAddLess180] = useState('0');
  const [newDeletions, setNewDeletions] = useState('0');
  const [newMethod, setNewMethod] = useState<'WDV' | 'SLM'>('WDV');

  // Calculate results for all assets
  const calculationResults = useMemo(() => {
    return assets.map((a) => calculateAssetDepreciation(a));
  }, [assets]);

  const totalCoActDepr = calculationResults.reduce((acc, r) => acc + r.coActDepreciation, 0);
  const totalItActDepr = calculationResults.reduce((acc, r) => acc + r.itActDepreciation, 0);
  const totalTimingDiff = totalCoActDepr - totalItActDepr;
  const totalCoActClosing = calculationResults.reduce((acc, r) => acc + r.coActClosingWdv, 0);

  const handleAddAsset = (e: React.FormEvent) => {
    e.preventDefault();
    if (!newName.trim()) return;

    const costNum = parseFloat(newCost) || 0;
    const residualVal = Math.round(costNum * 0.05); // default 5% under Companies Act

    const newAsset: AssetRecord = {
      id: `ast-${Date.now()}`,
      code: newCode.trim() || `AST-${assets.length + 1}`,
      name: newName.trim(),
      category: newCategory,
      purchaseDate: '2026-04-01',
      putToUseDate: '2026-04-05',
      originalCost: costNum,
      residualValue: residualVal,
      openingWdv: parseFloat(newOpeningWdv) || costNum,
      additionsMoreThan180Days: parseFloat(newAddMore180) || 0,
      additionsLessThan180Days: parseFloat(newAddLess180) || 0,
      deletionsDeductions: parseFloat(newDeletions) || 0,
      saleRealization: parseFloat(newDeletions) || 0,
      method: newMethod,
    };

    setAssets([...assets, newAsset]);
    setShowAddModal(false);
    setNewName('');
    setNewCode('');
    onShowToast('Asset Added', `Appended ${newAsset.name} to the register.`, 'success');
  };

  const handleDeleteAsset = (id: string) => {
    setAssets(assets.filter((a) => a.id !== id));
    onShowToast('Asset Removed', 'Removed asset from the schedule.', 'info');
  };


  const handleResetAssets = () => {
    setAssets([]);
    onShowToast('Register Cleared', 'All assets removed from schedule.', 'info');
  };

  const handleExportExcel = () => {
    if (assets.length === 0) {
      onShowToast('No Assets Found', 'Please add assets before exporting.', 'error');
      return;
    }
    try {
      const safeEntity = (entityName.trim() || 'Client_Entity').replace(/\s+/g, '_');
      const wb = exportDepreciationWorkbook(assets, calculationResults, safeEntity, financialYear);
      XLSX.writeFile(wb, `Depreciation_Schedule_${safeEntity}_${financialYear}.xlsx`);
      onShowToast('Schedule Exported', 'Downloaded 4-sheet Companies Act and IT Act depreciation workbook.', 'success');
    } catch (err: any) {
      onShowToast('Export Failed', err.message || 'Could not export schedule.', 'error');
    }
  };

  const handlePrint = () => {
    window.print();
  };

  const formatInr = (val: number) => {
    return val.toLocaleString('en-IN', { style: 'currency', currency: 'INR', maximumFractionDigits: 0 });
  };

  return (
    <div className="space-y-6">
      <ToolHeader
        toolId="depreciation"
        title="Depreciation Schedule Generator"
        description="Year-wise fixed asset register with automated Companies Act (WDV/SLM), Income Tax Act (Block of Assets), and AS 22 deferred tax variance."
        badge="Schedule II & Sec 32"
        onBack={onBack}
        actionButtons={
          <div className="flex items-center gap-2">
            {assets.length > 0 && (
              <button
                onClick={handleResetAssets}
                className="px-3 py-2 bg-white border border-gray-200 hover:border-gray-300 text-gray-600 hover:text-gray-900 rounded-lg text-xs font-medium transition-colors cursor-pointer"
              >
                Clear
              </button>
            )}
            <button
              onClick={handlePrint}
              className="flex items-center gap-1.5 px-3 py-2 bg-white border border-gray-200 hover:border-gray-300 text-gray-900 rounded-lg text-xs font-semibold shadow-sm transition-colors cursor-pointer"
              title="Print clean schedule report"
            >
              <Printer className="w-3.5 h-3.5 text-gray-900" />
              Print Schedule
            </button>
            <button
              onClick={handleExportExcel}
              className="flex items-center gap-1.5 px-4 py-2 bg-gray-900 hover:bg-black text-white rounded-lg text-xs font-semibold shadow-sm transition-colors cursor-pointer"
            >
              <Download className="w-3.5 h-3.5" />
              Export Schedule (Excel)
            </button>
          </div>
        }
      />

      <DataFormatGuide
        toolType="audit"
        acceptedFormats="Excel (.xlsx, .xls), CSV (.csv)"
        requiredFields={[
          { name: 'Asset Name', description: 'Name of the fixed asset', aliases: ['assetName', 'name', 'description'] },
          { name: 'Cost', description: 'Original cost of the asset', aliases: ['cost', 'grossBlock'] },
        ]}
      />

      {/* Entity Details & Metrics Card */}
      <div className="bg-white border border-gray-200 rounded-xl p-5 shadow-sm space-y-4">
        <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 pb-3 border-b border-gray-200">
          <div className="flex flex-wrap items-center gap-3">
            <div>
              <label className="text-[10px] font-semibold text-gray-600 uppercase block mb-0.5">Entity Name</label>
              <input
                type="text"
                value={entityName}
                onChange={(e) => setEntityName(e.target.value)}
                placeholder="e.g. Apex Precision Pvt Ltd"
                className="px-2.5 py-1 bg-white border border-gray-200 rounded text-xs font-semibold text-gray-900 focus:outline-none focus:border-gray-300"
              />
            </div>
            <div>
              <label className="text-[10px] font-semibold text-gray-600 uppercase block mb-0.5">Financial Year</label>
              <select
                value={financialYear}
                onChange={(e) => setFinancialYear(e.target.value)}
                className="px-2.5 py-1 bg-white border border-gray-200 rounded text-xs font-semibold text-gray-900 focus:outline-none focus:border-gray-300"
              >
                <option value="FY 2030-31">FY 2030-31</option>
                <option value="FY 2029-30">FY 2029-30</option>
                <option value="FY 2028-29">FY 2028-29</option>
                <option value="FY 2027-28">FY 2027-28</option>
                <option value="FY 2026-27">FY 2026-27</option>
                <option value="FY 2025-26">FY 2025-26</option>
                <option value="FY 2024-25">FY 2024-25</option>
                <option value="FY 2023-24">FY 2023-24</option>
                <option value="FY 2022-23">FY 2022-23</option>
              </select>
            </div>
          </div>

          <div className="flex items-center gap-2 print:hidden">
            <button
              onClick={() => setShowAddModal(true)}
              className="flex items-center gap-1.5 px-3 py-1.5 bg-gray-900 hover:bg-black text-white text-xs font-semibold rounded-lg shadow-sm transition-colors cursor-pointer"
            >
              <Plus className="w-3.5 h-3.5" />
              Add Fixed Asset
            </button>
          </div>
        </div>

        {/* 4 Summary Cards */}
        <div className="grid grid-cols-2 sm:grid-cols-4 gap-3">
          <div className="p-3 bg-gray-50 border border-gray-200 rounded-lg">
            <span className="text-[11px] font-semibold text-gray-600 block">Companies Act Depr (Books)</span>
            <span className="text-base font-bold text-gray-900 mt-0.5 block">{formatInr(totalCoActDepr)}</span>
            <span className="text-[10px] text-gray-600">Schedule II Useful Life</span>
          </div>
          <div className="p-3 bg-gray-100 border border-gray-200 rounded-lg">
            <span className="text-[11px] font-semibold text-gray-900 block">Income Tax Depr (Sec 32)</span>
            <span className="text-base font-bold text-gray-900 mt-0.5 block">{formatInr(totalItActDepr)}</span>
            <span className="text-[10px] text-gray-900">Block of Assets Rate</span>
          </div>
          <div className="p-3 bg-blue-50 border border-blue-200 rounded-lg">
            <span className="text-[11px] font-semibold text-blue-900 block">AS 22 Timing Difference</span>
            <span className="text-base font-bold text-blue-900 mt-0.5 block">{formatInr(totalTimingDiff)}</span>
            <span className="text-[10px] text-blue-700">
              {totalTimingDiff >= 0 ? 'DTA Impact' : 'DTL Impact'}
            </span>
          </div>
          <div className="p-3 bg-slate-50 border border-slate-200 rounded-lg">
            <span className="text-[11px] font-semibold text-slate-700 block">Closing Net Block (Books)</span>
            <span className="text-base font-bold text-slate-800 mt-0.5 block">{formatInr(totalCoActClosing)}</span>
            <span className="text-[10px] text-slate-600">Balance Sheet Net PPE</span>
          </div>
        </div>
      </div>

      {/* Tabs */}
      <div className="flex border-b border-gray-200 print:hidden">
        <button
          onClick={() => setActiveTab('co_act')}
          className={`px-4 py-2 text-xs font-semibold border-b-2 transition-colors ${
            activeTab === 'co_act'
              ? 'border-gray-300 text-gray-900 bg-white'
              : 'border-transparent text-gray-600 hover:text-gray-900'
          }`}
        >
          Companies Act 2013 (Schedule II)
        </button>
        <button
          onClick={() => setActiveTab('it_act')}
          className={`px-4 py-2 text-xs font-semibold border-b-2 transition-colors ${
            activeTab === 'it_act'
              ? 'border-gray-300 text-gray-900 bg-white'
              : 'border-transparent text-gray-600 hover:text-gray-900'
          }`}
        >
          Income Tax Act 1961 (Section 32 Block)
        </button>
        <button
          onClick={() => setActiveTab('deferred_tax')}
          className={`px-4 py-2 text-xs font-semibold border-b-2 transition-colors ${
            activeTab === 'deferred_tax'
              ? 'border-gray-300 text-gray-900 bg-white'
              : 'border-transparent text-gray-600 hover:text-gray-900'
          }`}
        >
          Deferred Tax Difference (AS 22 / Ind AS 12)
        </button>
      </div>

      {/* Tab 1: Companies Act 2013 */}
      {activeTab === 'co_act' && (
        <div className="bg-white border border-gray-200 rounded-xl overflow-hidden shadow-sm">
          <div className="overflow-x-auto">
            <table className="w-full text-left text-xs border-collapse">
              <thead className="bg-gray-50 border-b border-gray-200 text-gray-600 uppercase font-bold text-[10px] tracking-wider">
                <tr>
                  <th className="py-2.5 px-3 w-12 text-center">#</th>
                  <th className="py-2.5 px-4">Asset Code & Name</th>
                  <th className="py-2.5 px-3">Category</th>
                  <th className="py-2.5 px-3">Method</th>
                  <th className="py-2.5 px-3 text-right">Opening WDV</th>
                  <th className="py-2.5 px-3 text-right">Additions</th>
                  <th className="py-2.5 px-3 text-right">Deletions</th>
                  <th className="py-2.5 px-4 text-right text-gray-900">Depreciation</th>
                  <th className="py-2.5 px-4 text-right font-bold">Closing WDV</th>
                  <th className="py-2.5 px-2 text-center w-10 print:hidden"></th>
                </tr>
              </thead>
              <tbody className="divide-y divide-[#E5E7EB]">
                {calculationResults.length === 0 ? (
                  <tr>
                    <td colSpan={10} className="py-12 text-center text-gray-500">
                      <div className="flex flex-col items-center justify-center gap-2">
                        <TrendingDown className="w-7 h-7 text-gray-400" />
                        <span className="text-xs font-semibold text-gray-700">No fixed assets registered in schedule</span>
                        <p className="text-[11px] text-gray-500">Add an asset manually using the button above.</p>
                        <button
                          onClick={() => setShowAddModal(true)}
                          className="mt-2 px-3.5 py-1.5 bg-gray-900 text-white text-xs font-semibold rounded-lg hover:bg-black transition-colors cursor-pointer flex items-center gap-1.5"
                        >
                          <Plus className="w-3.5 h-3.5" />
                          <span>Add Fixed Asset</span>
                        </button>
                      </div>
                    </td>
                  </tr>
                ) : (
                  calculationResults.map((res, idx) => {
                    const asset = assets[idx];
                    return (
                      <tr key={res.assetId} className="hover:bg-gray-50 transition-colors">
                        <td className="py-3 px-3 text-center text-gray-600">{idx + 1}</td>
                        <td className="py-3 px-4">
                          <div className="font-semibold text-xs text-gray-900">{res.assetName}</div>
                          <div className="text-[10px] font-mono text-gray-600">{asset?.code}</div>
                        </td>
                        <td className="py-3 px-3 text-gray-600">
                          <span className="text-[11px] bg-slate-100 text-slate-700 px-1.5 py-0.5 rounded">
                            {res.categoryName} ({res.coActUsefulLife}y)
                          </span>
                        </td>
                        <td className="py-3 px-3">
                          <span className="text-[10px] font-semibold bg-gray-100 text-gray-900 px-1.5 py-0.5 rounded border border-gray-200">
                            {res.coActMethod}
                          </span>
                        </td>
                        <td className="py-3 px-3 text-right font-mono text-xs">{formatInr(res.coActOpeningWdv)}</td>
                        <td className="py-3 px-3 text-right font-mono text-xs text-gray-900">
                          {formatInr(res.coActTotalAdditions)}
                        </td>
                        <td className="py-3 px-3 text-right font-mono text-xs text-red-600">
                          {res.coActDeletions > 0 ? formatInr(res.coActDeletions) : '-'}
                        </td>
                        <td className="py-3 px-4 text-right font-mono text-xs font-semibold text-gray-900">
                          {formatInr(res.coActDepreciation)}
                        </td>
                        <td className="py-3 px-4 text-right font-mono text-xs font-bold text-gray-900">
                          {formatInr(res.coActClosingWdv)}
                        </td>
                        <td className="py-3 px-2 text-center print:hidden">
                          <button
                            onClick={() => handleDeleteAsset(res.assetId)}
                            className="p-1 text-gray-600 hover:text-[#DC2626] transition-colors"
                            title="Remove asset"
                          >
                            <Trash2 className="w-3.5 h-3.5" />
                          </button>
                        </td>
                      </tr>
                    );
                  })
                )}
              </tbody>
            </table>
          </div>
        </div>
      )}

      {/* Tab 2: Income Tax Act Section 32 */}
      {activeTab === 'it_act' && (
        <div className="bg-white border border-gray-200 rounded-xl overflow-hidden shadow-sm">
          <div className="overflow-x-auto">
            <table className="w-full text-left text-xs border-collapse">
              <thead className="bg-gray-50 border-b border-gray-200 text-gray-600 uppercase font-bold text-[10px] tracking-wider">
                <tr>
                  <th className="py-2.5 px-3 w-12 text-center">#</th>
                  <th className="py-2.5 px-4">Block / Asset Name</th>
                  <th className="py-2.5 px-3 text-center">IT Block Rate</th>
                  <th className="py-2.5 px-3 text-right">Opening WDV</th>
                  <th className="py-2.5 px-3 text-right">Additions &ge; 180d</th>
                  <th className="py-2.5 px-3 text-right">Additions &lt; 180d (Half)</th>
                  <th className="py-2.5 px-3 text-right">Deductions</th>
                  <th className="py-2.5 px-4 text-right text-gray-900">Tax Depr</th>
                  <th className="py-2.5 px-4 text-right font-bold">Closing Tax WDV</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-[#E5E7EB]">
                {calculationResults.length === 0 ? (
                  <tr>
                    <td colSpan={9} className="py-12 text-center text-gray-500">
                      <div className="flex flex-col items-center justify-center gap-2">
                        <Scale className="w-7 h-7 text-gray-400" />
                        <span className="text-xs font-semibold text-gray-700">No Block of Assets Registered</span>
                        <p className="text-[11px] text-gray-500">Section 32 rates (15%, 40%, 10%) are automatically mapped to asset blocks.</p>
                        <button
                          onClick={() => setShowAddModal(true)}
                          className="mt-2 px-3.5 py-1.5 bg-gray-900 text-white text-xs font-semibold rounded-lg hover:bg-black transition-colors cursor-pointer flex items-center gap-1.5"
                        >
                          <Plus className="w-3.5 h-3.5" />
                          <span>Add Fixed Asset</span>
                        </button>
                      </div>
                    </td>
                  </tr>
                ) : (
                  calculationResults.map((res, idx) => (
                    <tr key={res.assetId} className="hover:bg-gray-50 transition-colors">
                      <td className="py-3 px-3 text-center text-gray-600">{idx + 1}</td>
                      <td className="py-3 px-4 font-semibold text-xs text-gray-900">{res.assetName}</td>
                      <td className="py-3 px-3 text-center">
                        <span className="text-[11px] font-bold bg-amber-50 text-amber-900 border border-amber-200 px-2 py-0.5 rounded">
                          {res.itActBlockRate}%
                        </span>
                      </td>
                      <td className="py-3 px-3 text-right font-mono text-xs">{formatInr(res.itActOpeningWdv)}</td>
                      <td className="py-3 px-3 text-right font-mono text-xs">{formatInr(res.itActAdditionsMore180)}</td>
                      <td className="py-3 px-3 text-right font-mono text-xs text-amber-800">
                        {res.itActAdditionsLess180 > 0 ? (
                          <span>{formatInr(res.itActAdditionsLess180)} (50%)</span>
                        ) : (
                          '-'
                        )}
                      </td>
                      <td className="py-3 px-3 text-right font-mono text-xs text-red-600">
                        {res.itActDeductions > 0 ? formatInr(res.itActDeductions) : '-'}
                      </td>
                      <td className="py-3 px-4 text-right font-mono text-xs font-semibold text-gray-900">
                        {formatInr(res.itActDepreciation)}
                      </td>
                      <td className="py-3 px-4 text-right font-mono text-xs font-bold text-gray-900">
                        {formatInr(res.itActClosingWdv)}
                      </td>
                    </tr>
                  ))
                )}
              </tbody>
            </table>
          </div>
        </div>
      )}

      {/* Tab 3: Deferred Tax Variance */}
      {activeTab === 'deferred_tax' && (
        <div className="bg-white border border-gray-200 rounded-xl overflow-hidden shadow-sm">
          <div className="overflow-x-auto">
            <table className="w-full text-left text-xs border-collapse">
              <thead className="bg-gray-50 border-b border-gray-200 text-gray-600 uppercase font-bold text-[10px] tracking-wider">
                <tr>
                  <th className="py-2.5 px-3 w-12 text-center">#</th>
                  <th className="py-2.5 px-4">Asset Name</th>
                  <th className="py-2.5 px-4 text-right">Books Depr (Companies Act)</th>
                  <th className="py-2.5 px-4 text-right">Tax Depr (IT Act Sec 32)</th>
                  <th className="py-2.5 px-4 text-right font-bold">Timing Difference</th>
                  <th className="py-2.5 px-4">Deferred Tax Nature</th>
                  <th className="py-2.5 px-4 text-right">Tax Impact @ 25.17%</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-[#E5E7EB]">
                {calculationResults.length === 0 ? (
                  <tr>
                    <td colSpan={7} className="py-12 text-center text-gray-500">
                      <div className="flex flex-col items-center justify-center gap-2">
                        <Scale className="w-7 h-7 text-gray-400" />
                        <span className="text-xs font-semibold text-gray-700">No Timing Differences to Compute</span>
                        <p className="text-[11px] text-gray-500">AS 22 / Ind AS 12 Deferred Tax will calculate once assets are populated.</p>
                        <button
                          onClick={() => setShowAddModal(true)}
                          className="mt-2 px-3.5 py-1.5 bg-gray-900 text-white text-xs font-semibold rounded-lg hover:bg-black transition-colors cursor-pointer flex items-center gap-1.5"
                        >
                          <Plus className="w-3.5 h-3.5" />
                          <span>Add Fixed Asset</span>
                        </button>
                      </div>
                    </td>
                  </tr>
                ) : (
                  calculationResults.map((res, idx) => (
                  <tr key={res.assetId} className="hover:bg-gray-50 transition-colors">
                    <td className="py-3 px-3 text-center text-gray-600">{idx + 1}</td>
                    <td className="py-3 px-4 font-semibold text-xs text-gray-900">{res.assetName}</td>
                    <td className="py-3 px-4 text-right font-mono text-xs">{formatInr(res.coActDepreciation)}</td>
                    <td className="py-3 px-4 text-right font-mono text-xs">{formatInr(res.itActDepreciation)}</td>
                    <td className="py-3 px-4 text-right font-mono text-xs font-bold">
                      <span className={res.timingDifference >= 0 ? 'text-gray-900' : 'text-blue-700'}>
                        {formatInr(res.timingDifference)}
                      </span>
                    </td>
                    <td className="py-3 px-4">
                      {res.timingDifference >= 0 ? (
                        <span className="text-[10px] font-semibold bg-gray-100 text-gray-900 px-2 py-0.5 rounded border border-gray-200">
                          Deferred Tax Asset (DTA)
                        </span>
                      ) : (
                        <span className="text-[10px] font-semibold bg-blue-50 text-blue-800 px-2 py-0.5 rounded border border-blue-200">
                          Deferred Tax Liability (DTL)
                        </span>
                      )}
                    </td>
                    <td className="py-3 px-4 text-right font-mono text-xs font-semibold text-gray-900">
                      {formatInr(Math.abs(res.taxEffect))}
                    </td>
                  </tr>
                )))}
              </tbody>
            </table>
          </div>
        </div>
      )}

      {/* Modal: Add Fixed Asset */}
      {showAddModal && (
        <div className="fixed inset-0 bg-black/40 backdrop-blur-xs flex items-center justify-center z-50 p-4">
          <div className="bg-white border border-gray-200 rounded-xl shadow-xl w-full max-w-lg p-6 space-y-4">
            <h3 className="text-sm font-bold text-gray-900 uppercase tracking-wider">
              Add Fixed Asset to Register
            </h3>
            <form onSubmit={handleAddAsset} className="space-y-3 text-xs">
              <div className="grid grid-cols-3 gap-3">
                <div>
                  <label className="font-semibold text-gray-900 block mb-1">Asset Code</label>
                  <input
                    type="text"
                    placeholder="e.g. FA-01"
                    value={newCode}
                    onChange={(e) => setNewCode(e.target.value)}
                    className="w-full px-3 py-2 bg-white border border-gray-200 rounded-lg text-xs text-gray-900 focus:outline-none focus:border-gray-300"
                  />
                </div>
                <div className="col-span-2">
                  <label className="font-semibold text-gray-900 block mb-1">Asset Name</label>
                  <input
                    type="text"
                    required
                    placeholder="e.g. Injection Moulding Machine"
                    value={newName}
                    onChange={(e) => setNewName(e.target.value)}
                    className="w-full px-3 py-2 bg-white border border-gray-200 rounded-lg text-xs text-gray-900 focus:outline-none focus:border-gray-300"
                  />
                </div>
              </div>

              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="font-semibold text-gray-900 block mb-1">Category & Useful Life</label>
                  <select
                    value={newCategory}
                    onChange={(e) => setNewCategory(e.target.value as AssetCategory)}
                    className="w-full px-3 py-2 bg-white border border-gray-200 rounded-lg text-xs text-gray-900 focus:outline-none focus:border-gray-300"
                  >
                    {Object.entries(ASSET_CATEGORIES).map(([catKey, spec]) => (
                      <option key={catKey} value={catKey}>
                        {spec.name} ({spec.usefulLifeYears}y / IT: {spec.itActBlockRate}%)
                      </option>
                    ))}
                  </select>
                </div>
                <div>
                  <label className="font-semibold text-gray-900 block mb-1">Companies Act Method</label>
                  <select
                    value={newMethod}
                    onChange={(e) => setNewMethod(e.target.value as any)}
                    className="w-full px-3 py-2 bg-white border border-gray-200 rounded-lg text-xs text-gray-900 focus:outline-none focus:border-gray-300"
                  >
                    <option value="WDV">Written Down Value (WDV)</option>
                    <option value="SLM">Straight Line Method (SLM)</option>
                  </select>
                </div>
              </div>

              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="font-semibold text-gray-900 block mb-1">Original Cost (₹)</label>
                  <input
                    type="number"
                    required
                    value={newCost}
                    onChange={(e) => setNewCost(e.target.value)}
                    className="w-full px-3 py-2 bg-white border border-gray-200 rounded-lg text-xs text-gray-900 focus:outline-none focus:border-gray-300"
                  />
                </div>
                <div>
                  <label className="font-semibold text-gray-900 block mb-1">Opening WDV (₹)</label>
                  <input
                    type="number"
                    value={newOpeningWdv}
                    onChange={(e) => setNewOpeningWdv(e.target.value)}
                    className="w-full px-3 py-2 bg-white border border-gray-200 rounded-lg text-xs text-gray-900 focus:outline-none focus:border-gray-300"
                  />
                </div>
              </div>

              <div className="grid grid-cols-3 gap-3">
                <div>
                  <label className="font-semibold text-gray-900 block mb-1">Additions &ge; 180d</label>
                  <input
                    type="number"
                    value={newAddMore180}
                    onChange={(e) => setNewAddMore180(e.target.value)}
                    className="w-full px-3 py-2 bg-white border border-gray-200 rounded-lg text-xs text-gray-900 focus:outline-none focus:border-gray-300"
                  />
                </div>
                <div>
                  <label className="font-semibold text-gray-900 block mb-1">Additions &lt; 180d (Half)</label>
                  <input
                    type="number"
                    value={newAddLess180}
                    onChange={(e) => setNewAddLess180(e.target.value)}
                    className="w-full px-3 py-2 bg-white border border-gray-200 rounded-lg text-xs text-gray-900 focus:outline-none focus:border-gray-300"
                  />
                </div>
                <div>
                  <label className="font-semibold text-gray-900 block mb-1">Deletions / Sales</label>
                  <input
                    type="number"
                    value={newDeletions}
                    onChange={(e) => setNewDeletions(e.target.value)}
                    className="w-full px-3 py-2 bg-white border border-gray-200 rounded-lg text-xs text-gray-900 focus:outline-none focus:border-gray-300"
                  />
                </div>
              </div>

              <div className="flex justify-end gap-2 pt-3 border-t border-gray-200">
                <button
                  type="button"
                  onClick={() => setShowAddModal(false)}
                  className="px-3 py-1.5 border border-gray-200 rounded-lg text-gray-600 hover:bg-gray-50"
                >
                  Cancel
                </button>
                <button
                  type="submit"
                  className="px-4 py-1.5 bg-gray-900 hover:bg-black text-white font-semibold rounded-lg shadow-sm"
                >
                  Add Asset
                </button>
              </div>
            </form>
          </div>
        </div>
      )}
    </div>
  );
};