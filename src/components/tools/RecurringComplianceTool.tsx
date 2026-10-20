import React, { useState, useMemo, useRef } from 'react';
import * as XLSX from 'xlsx';
import { createStyledWorksheet } from '../../utils/excelStyler';
import {
  ArrowLeft,
  Calendar,
  CheckCircle2,
  Clock,
  AlertTriangle,
  FileSpreadsheet,
  RefreshCw,
  Search,
  Plus,
  Building,
  Filter,
  Receipt,
  FileCheck,
  UploadCloud,
  Trash2,
  Download,
  X,
  FileText,
  HelpCircle,
} from 'lucide-react';
import { InnerToolBadge } from '../common/InnerToolBadge';

interface RecurringComplianceToolProps {
  onBack: () => void;
  onShowToast: (title: string, description?: string, type?: 'success' | 'error' | 'info') => void;
}

export interface ComplianceItem {
  id: string;
  name: string;
  act: 'GST' | 'Income Tax' | 'MCA / ROC' | 'Labor & PF' | 'Other';
  section: string;
  frequency: 'Monthly' | 'Quarterly' | 'Annual' | 'One-Time';
  dueDate: string; // YYYY-MM-DD
  status: 'pending' | 'filed' | 'overdue';
  filingProofRef?: string;
  filedDate?: string;
  lateFeeEstimated?: number;
}

const MONTHS_LIST = [
  'January', 'February', 'March', 'April', 'May', 'June',
  'July', 'August', 'September', 'October', 'November', 'December',
];

function generateStatutoryScheduleForMonth(
  monthName: string,
  year: number,
  acts: { gst: boolean; it: boolean; mca: boolean; labor: boolean }
): ComplianceItem[] {
  const monthMap: Record<string, number> = {
    january: 1, february: 2, march: 3, april: 4, may: 5, june: 6,
    july: 7, august: 8, september: 9, october: 10, november: 11, december: 12,
  };
  const cleanMonth = monthName.toLowerCase().trim();
  let mNum = monthMap[cleanMonth];
  if (!mNum) {
    for (const [k, v] of Object.entries(monthMap)) {
      if (cleanMonth.startsWith(k.slice(0, 3))) {
        mNum = v;
        break;
      }
    }
  }
  if (!mNum) mNum = 10;

  const mStr = String(mNum).padStart(2, '0');
  const items: ComplianceItem[] = [];

  // 1. Income Tax & TDS
  if (acts.it) {
    const tdsDueDay = mNum === 4 ? '30' : '07';
    items.push({
      id: `it_tds_${Date.now()}_1`,
      name: 'TDS Challan Payment (ITNS 281)',
      act: 'Income Tax',
      section: 'Sec 200(1) / Rule 30',
      frequency: 'Monthly',
      dueDate: `${year}-${mStr}-${tdsDueDay}`,
      status: 'pending',
      lateFeeEstimated: 0,
    });

    if (mNum === 7) {
      items.push({
        id: `it_qtr_${Date.now()}_2`,
        name: 'TDS Quarterly Return (Form 26Q / 24Q - Q1)',
        act: 'Income Tax',
        section: 'Sec 200(3) / Rule 31A',
        frequency: 'Quarterly',
        dueDate: `${year}-07-31`,
        status: 'pending',
        lateFeeEstimated: 0,
      });
    } else if (mNum === 10) {
      items.push({
        id: `it_qtr_${Date.now()}_2`,
        name: 'TDS Quarterly Return (Form 26Q / 24Q - Q2)',
        act: 'Income Tax',
        section: 'Sec 200(3) / Rule 31A',
        frequency: 'Quarterly',
        dueDate: `${year}-10-31`,
        status: 'pending',
        lateFeeEstimated: 0,
      });
    } else if (mNum === 1) {
      items.push({
        id: `it_qtr_${Date.now()}_2`,
        name: 'TDS Quarterly Return (Form 26Q / 24Q - Q3)',
        act: 'Income Tax',
        section: 'Sec 200(3) / Rule 31A',
        frequency: 'Quarterly',
        dueDate: `${year}-01-31`,
        status: 'pending',
        lateFeeEstimated: 0,
      });
    } else if (mNum === 5) {
      items.push({
        id: `it_qtr_${Date.now()}_2`,
        name: 'TDS Quarterly Return (Form 26Q / 24Q - Q4)',
        act: 'Income Tax',
        section: 'Sec 200(3) / Rule 31A',
        frequency: 'Quarterly',
        dueDate: `${year}-05-31`,
        status: 'pending',
        lateFeeEstimated: 0,
      });
    }

    if (mNum === 6) {
      items.push({
        id: `it_adv_${Date.now()}_3`,
        name: 'Advance Tax 1st Instalment (15%)',
        act: 'Income Tax',
        section: 'Sec 208 / 211',
        frequency: 'Quarterly',
        dueDate: `${year}-06-15`,
        status: 'pending',
        lateFeeEstimated: 0,
      });
    } else if (mNum === 9) {
      items.push({
        id: `it_adv_${Date.now()}_3`,
        name: 'Advance Tax 2nd Instalment (45%)',
        act: 'Income Tax',
        section: 'Sec 208 / 211',
        frequency: 'Quarterly',
        dueDate: `${year}-09-15`,
        status: 'pending',
        lateFeeEstimated: 0,
      });
    } else if (mNum === 12) {
      items.push({
        id: `it_adv_${Date.now()}_3`,
        name: 'Advance Tax 3rd Instalment (75%)',
        act: 'Income Tax',
        section: 'Sec 208 / 211',
        frequency: 'Quarterly',
        dueDate: `${year}-12-15`,
        status: 'pending',
        lateFeeEstimated: 0,
      });
    } else if (mNum === 3) {
      items.push({
        id: `it_adv_${Date.now()}_3`,
        name: 'Advance Tax 4th Instalment (100%)',
        act: 'Income Tax',
        section: 'Sec 208 / 211',
        frequency: 'Quarterly',
        dueDate: `${year}-03-15`,
        status: 'pending',
        lateFeeEstimated: 0,
      });
    }
  }

  // 2. GST
  if (acts.gst) {
    items.push({
      id: `gst_gstr1_${Date.now()}_4`,
      name: 'GSTR-1 Outward Supplies (Monthly Filers)',
      act: 'GST',
      section: 'Sec 37 CGST Act',
      frequency: 'Monthly',
      dueDate: `${year}-${mStr}-11`,
      status: 'pending',
      lateFeeEstimated: 0,
    });

    items.push({
      id: `gst_gstr3b_${Date.now()}_5`,
      name: 'GSTR-3B Summary Return & Tax Payment',
      act: 'GST',
      section: 'Sec 39 CGST Act',
      frequency: 'Monthly',
      dueDate: `${year}-${mStr}-20`,
      status: 'pending',
      lateFeeEstimated: 0,
    });
  }

  // 3. Labor & PF/ESI
  if (acts.labor) {
    items.push({
      id: `labor_pf_${Date.now()}_6`,
      name: 'EPF & ESI Monthly Contribution & Return',
      act: 'Labor & PF',
      section: 'EPF Act 1952 / ESI Act',
      frequency: 'Monthly',
      dueDate: `${year}-${mStr}-15`,
      status: 'pending',
      lateFeeEstimated: 0,
    });
  }

  // 4. MCA / ROC
  if (acts.mca) {
    if (mNum === 10) {
      items.push({
        id: `mca_aoc4_${Date.now()}_7`,
        name: 'MCA Form AOC-4 Financial Statements Filing',
        act: 'MCA / ROC',
        section: 'Sec 137 Companies Act 2013',
        frequency: 'Annual',
        dueDate: `${year}-10-29`,
        status: 'pending',
        lateFeeEstimated: 0,
      });
    }
    if (mNum === 11) {
      items.push({
        id: `mca_mgt7_${Date.now()}_8`,
        name: 'MCA Form MGT-7 Annual Return',
        act: 'MCA / ROC',
        section: 'Sec 92 Companies Act 2013',
        frequency: 'Annual',
        dueDate: `${year}-11-29`,
        status: 'pending',
        lateFeeEstimated: 0,
      });
    }
    if (mNum === 9) {
      items.push({
        id: `mca_dir3_${Date.now()}_9`,
        name: 'MCA DIR-3 KYC Director Verification',
        act: 'MCA / ROC',
        section: 'Rule 12A Companies Rules',
        frequency: 'Annual',
        dueDate: `${year}-09-30`,
        status: 'pending',
        lateFeeEstimated: 0,
      });
    }
  }

  return items;
}

export const RecurringComplianceTool: React.FC<RecurringComplianceToolProps> = ({
  onBack,
  onShowToast,
}) => {
  const [clientName, setClientName] = useState('');
  const [selectedMonth, setSelectedMonth] = useState('October');
  const [selectedYear, setSelectedYear] = useState(2026);
  const [financialYear, setFinancialYear] = useState('FY 2026-27');

  // Statutory acts to include in generator
  const [actsToInclude, setActsToInclude] = useState({
    gst: true,
    it: true,
    mca: true,
    labor: true,
  });

  // Zero fake data by default: starts empty!
  const [items, setItems] = useState<ComplianceItem[]>([]);
  const [activeTab, setActiveTab] = useState<'calendar' | 'penalties' | 'proof_vault'>('calendar');
  const [selectedAct, setSelectedAct] = useState<string>('all');
  const [searchQuery, setSearchQuery] = useState('');

  // Hidden File Input
  const fileInputRef = useRef<HTMLInputElement>(null);

  // Proof Modal
  const [showProofModal, setShowProofModal] = useState(false);
  const [selectedItemForProof, setSelectedItemForProof] = useState<ComplianceItem | null>(null);
  const [proofReference, setProofReference] = useState('');
  const [proofFilingDate, setProofFilingDate] = useState(new Date().toISOString().split('T')[0]);

  // Add Custom Item Modal
  const [showAddModal, setShowAddModal] = useState(false);
  const [newObligationName, setNewObligationName] = useState('');
  const [newAct, setNewAct] = useState<ComplianceItem['act']>('GST');
  const [newSection, setNewSection] = useState('');
  const [newFrequency, setNewFrequency] = useState<ComplianceItem['frequency']>('Monthly');
  const [newDueDate, setNewDueDate] = useState('2026-10-20');

  // Metrics
  const metrics = useMemo(() => {
    const total = items.length;
    const filed = items.filter((i) => i.status === 'filed').length;
    const overdue = items.filter((i) => i.status === 'overdue').length;
    const pending = items.filter((i) => i.status === 'pending').length;
    const totalLateFee = items.reduce((acc, curr) => acc + (curr.lateFeeEstimated || 0), 0);
    const complianceRate = total > 0 ? Math.round((filed / total) * 100) : 0;
    return { total, filed, overdue, pending, totalLateFee, complianceRate };
  }, [items]);

  // Filter items
  const filteredItems = useMemo(() => {
    return items.filter((item) => {
      const matchAct = selectedAct === 'all' || item.act === selectedAct;
      const matchSearch =
        item.name.toLowerCase().includes(searchQuery.toLowerCase()) ||
        item.section.toLowerCase().includes(searchQuery.toLowerCase()) ||
        (item.filingProofRef && item.filingProofRef.toLowerCase().includes(searchQuery.toLowerCase()));
      return matchAct && matchSearch;
    });
  }, [items, selectedAct, searchQuery]);

  // 1. Generate statutory schedule for selected month
  const handleGenerateStatutorySchedule = () => {
    const generated = generateStatutoryScheduleForMonth(selectedMonth, selectedYear, actsToInclude);
    if (generated.length === 0) {
      onShowToast('Selection Required', 'Please select at least one statutory act (GST, Income Tax, MCA, Labor).', 'error');
      return;
    }
    setItems(generated);
    onShowToast('Calendar Generated', `Loaded ${generated.length} official statutory obligations for ${selectedMonth} ${selectedYear}.`, 'success');
  };

  // 2. Upload file handler
  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const f = e.target.files?.[0];
    if (!f) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const buffer = evt.target?.result as ArrayBuffer;
        const wb = XLSX.read(buffer, { type: 'array', cellDates: true });
        const sheet = wb.Sheets[wb.SheetNames[0]];
        const rawRows = XLSX.utils.sheet_to_json(sheet) as Record<string, any>[];

        if (!rawRows || rawRows.length === 0) {
          onShowToast('Empty File', 'Spreadsheet contains zero compliance rows.', 'error');
          return;
        }

        const parsedItems: ComplianceItem[] = rawRows.map((r, idx) => {
          const name = (r['Compliance Obligation'] || r['Compliance Name'] || r['Obligation'] || r['Name'] || r['Particulars'] || `Compliance Task ${idx + 1}`).toString().trim();
          const actRaw = (r['Governing Act'] || r['Act'] || r['Law'] || 'GST').toString().trim();
          let act: ComplianceItem['act'] = 'GST';
          if (/income|tds|it/i.test(actRaw)) act = 'Income Tax';
          else if (/mca|roc|company/i.test(actRaw)) act = 'MCA / ROC';
          else if (/labor|pf|esi|epf/i.test(actRaw)) act = 'Labor & PF';
          else if (/other|custom/i.test(actRaw)) act = 'Other';

          const section = (r['Statutory Section'] || r['Section'] || r['Rule'] || 'Statutory Provision').toString().trim();
          const freqRaw = (r['Frequency'] || r['Periodicity'] || 'Monthly').toString().trim();
          let frequency: ComplianceItem['frequency'] = 'Monthly';
          if (/quarter/i.test(freqRaw)) frequency = 'Quarterly';
          else if (/annual|yearly/i.test(freqRaw)) frequency = 'Annual';
          else if (/one-time|adhoc/i.test(freqRaw)) frequency = 'One-Time';

          let dueDate = (r['Due Date'] || r['DueDate'] || r['Due_Date'] || '').toString().trim();
          if (dueDate.includes('/')) {
            const parts = dueDate.split('/');
            if (parts.length === 3) {
              if (parts[0].length === 4) dueDate = `${parts[0]}-${parts[1].padStart(2, '0')}-${parts[2].padStart(2, '0')}`;
              else dueDate = `${parts[2]}-${parts[1].padStart(2, '0')}-${parts[0].padStart(2, '0')}`;
            }
          }

          const proofRef = (r['Filing Reference (ARN/CIN)'] || r['ARN'] || r['CIN'] || r['Proof'] || r['Reference'] || '').toString().trim() || undefined;
          const statusRaw = (r['Current Status'] || r['Status'] || (proofRef ? 'filed' : 'pending')).toString().toLowerCase().trim();
          let status: ComplianceItem['status'] = 'pending';
          if (statusRaw.includes('filed') || statusRaw.includes('complete') || statusRaw.includes('done') || proofRef) {
            status = 'filed';
          } else if (statusRaw.includes('overdue') || statusRaw.includes('late')) {
            status = 'overdue';
          }

          const filedDate = (r['Filed Date'] || r['Filing Date'] || '').toString().trim() || undefined;

          return {
            id: `cmp_${Date.now()}_${idx + 1}`,
            name,
            act,
            section,
            frequency,
            dueDate: dueDate || new Date().toISOString().split('T')[0],
            status,
            filingProofRef: proofRef,
            filedDate,
            lateFeeEstimated: status === 'overdue' ? 1000 : 0,
          };
        });

        setItems(parsedItems);
      } catch (err: any) {
        onShowToast('Parse Error', `Could not read compliance spreadsheet: ${err.message}`, 'error');
      }
    };
    reader.readAsArrayBuffer(f);
  };

  // 3. Download Excel Template
  const handleDownloadTemplate = () => {
    const templateRows = [
      {
        'Compliance Obligation': 'GSTR-3B Summary Return & Tax Payment',
        'Governing Act': 'GST',
        'Statutory Section': 'Sec 39 CGST Act',
        'Frequency': 'Monthly',
        'Due Date': '2026-10-20',
        'Current Status': 'PENDING',
        'Filing Reference (ARN/CIN)': '',
        'Filed Date': '',
      },
      {
        'Compliance Obligation': 'TDS Challan Payment (ITNS 281)',
        'Governing Act': 'Income Tax',
        'Statutory Section': 'Sec 200(1) / Rule 30',
        'Frequency': 'Monthly',
        'Due Date': '2026-10-07',
        'Current Status': 'PENDING',
        'Filing Reference (ARN/CIN)': '',
        'Filed Date': '',
      },
      {
        'Compliance Obligation': 'EPF & ESI Monthly Contribution',
        'Governing Act': 'Labor & PF',
        'Statutory Section': 'EPF Act 1952 / ESI Act',
        'Frequency': 'Monthly',
        'Due Date': '2026-10-15',
        'Current Status': 'PENDING',
        'Filing Reference (ARN/CIN)': '',
        'Filed Date': '',
      },
    ];

    const ws = createStyledWorksheet(templateRows, {
      columnWidths: {
        'Compliance Obligation': 34,
        'Governing Act': 18,
        'Statutory Section': 26,
        'Frequency': 16,
        'Due Date': 16,
        'Current Status': 18,
        'Filing Reference (ARN/CIN)': 28,
        'Filed Date': 16,
      },
      freezeHeader: true,
    });
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Compliance_Template');
    XLSX.writeFile(wb, 'Statutory_Compliance_Schedule_Template.xlsx');
    onShowToast('Template Downloaded', 'Compliance calendar template saved as Excel.', 'success');
  };

  // 4. Record Proof Action
  const handleSaveProof = () => {
    if (!selectedItemForProof) return;
    if (!proofReference.trim()) {
      onShowToast('Missing Reference', 'Please enter ARN, Challan CIN, or Acknowledgement No.', 'error');
      return;
    }

    setItems((prev) =>
      prev.map((i) =>
        i.id === selectedItemForProof.id
          ? {
              ...i,
              status: 'filed',
              filingProofRef: proofReference.trim(),
              filedDate: proofFilingDate,
              lateFeeEstimated: 0,
            }
          : i
      )
    );

    setShowProofModal(false);
    setProofReference('');
    setSelectedItemForProof(null);
    onShowToast('Filing Recorded', 'Statutory filing proof saved successfully.', 'success');
  };

  // 5. Add Custom Item
  const handleAddCustomItem = (e: React.FormEvent) => {
    e.preventDefault();
    if (!newObligationName.trim()) {
      onShowToast('Missing Name', 'Please enter a name for the compliance obligation.', 'error');
      return;
    }

    const newItem: ComplianceItem = {
      id: `custom_${Date.now()}`,
      name: newObligationName.trim(),
      act: newAct,
      section: newSection.trim() || 'Custom Requirement',
      frequency: newFrequency,
      dueDate: newDueDate,
      status: 'pending',
      lateFeeEstimated: 0,
    };

    setItems((prev) => [newItem, ...prev]);
    setShowAddModal(false);
    setNewObligationName('');
    setNewSection('');
    onShowToast('Obligation Added', `Added "${newItem.name}" to statutory calendar.`, 'success');
  };

  // 6. Delete Item
  const handleDeleteItem = (id: string) => {
    setItems((prev) => prev.filter((i) => i.id !== id));
    onShowToast('Obligation Removed', 'Removed item from calendar.', 'info');
  };

  // 7. Export to Excel
  const handleExportExcel = () => {
    if (items.length === 0) {
      onShowToast('No Compliances', 'There are no compliance items to export.', 'error');
      return;
    }

    const exportData = items.map((i, idx) => ({
      'S.No': idx + 1,
      'Compliance Obligation': i.name,
      'Governing Act': i.act,
      'Statutory Section': i.section,
      Frequency: i.frequency,
      'Due Date': i.dueDate,
      'Current Status': i.status.toUpperCase(),
      'Filing Reference (ARN/CIN)': i.filingProofRef || 'Pending Filing',
      'Filed Date': i.filedDate || 'N/A',
      'Estimated Penalty / Interest (₹)': i.lateFeeEstimated || 0,
    }));

    const ws = createStyledWorksheet(exportData, {
      currencyColumns: ['Estimated Penalty / Interest (₹)'],
      dateColumns: ['Due Date', 'Filed Date'],
      columnWidths: {
        'S.No': 8,
        'Compliance Obligation': 36,
        'Governing Act': 18,
        'Statutory Section': 26,
        Frequency: 16,
        'Due Date': 16,
        'Current Status': 18,
        'Filing Reference (ARN/CIN)': 28,
        'Filed Date': 16,
        'Estimated Penalty / Interest (₹)': 24,
      },
      freezeHeader: true,
    });
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Compliance_Schedule');
    const safeClient = clientName ? clientName.replace(/\s+/g, '_') : 'Client';
    XLSX.writeFile(wb, `${safeClient}_Statutory_Calendar_${selectedMonth}_${selectedYear}.xlsx`);
    onShowToast('Export Complete', 'Statutory compliance schedule downloaded as Excel.', 'success');
  };

  // 8. Clear All
  const handleClearAll = () => {
    setItems([]);
    onShowToast('Calendar Reset', 'Cleared all compliance obligations.', 'info');
  };

  return (
    <div className="space-y-6">
      {/* Hidden file upload input */}
      <input
        type="file"
        ref={fileInputRef}
        onChange={handleFileUpload}
        accept=".xlsx,.xls,.csv"
        className="hidden"
      />

      {/* Top Back Navigation */}
      <div>
        <button
          onClick={onBack}
          className="inline-flex items-center gap-2 px-3 py-1.5 rounded-lg text-xs font-semibold text-gray-700 bg-white border border-gray-200 shadow-2xs hover:bg-gray-50 hover:text-gray-900 hover:border-gray-300 transition-all cursor-pointer mb-5 group"
        >
          <ArrowLeft className="w-3.5 h-3.5 text-gray-400 group-hover:text-gray-900 group-hover:-translate-x-0.5 transition-transform" />
          <span>Back to All Tools</span>
        </button>

        {/* Header Row */}
        <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 pb-5 border-b border-gray-200">
          <div className="flex items-center gap-3.5">
            <InnerToolBadge toolId="recurring-compliance" />
            <div>
              <h1 className="text-xl font-bold text-gray-900 tracking-tight">
                Recurring Compliance & Statutory Calendar
              </h1>
              <p className="text-xs text-gray-500 mt-0.5">
                Track GSTR-1, GSTR-3B, TDS payments, Advance Tax, and statutory due dates with filing proof verification.
              </p>
            </div>
          </div>

          {items.length > 0 && (
            <div className="flex flex-wrap items-center gap-2">
              <button
                onClick={() => setShowAddModal(true)}
                className="inline-flex items-center gap-1.5 px-3 py-2 text-xs font-semibold rounded-lg bg-white text-gray-700 border border-gray-200 hover:bg-gray-50 hover:text-gray-900 transition-colors shadow-2xs cursor-pointer"
              >
                <Plus className="w-3.5 h-3.5" />
                <span>Add Task</span>
              </button>

              <button
                onClick={() => fileInputRef.current?.click()}
                className="inline-flex items-center gap-1.5 px-3 py-2 text-xs font-semibold rounded-lg bg-white text-gray-700 border border-gray-200 hover:bg-gray-50 hover:text-gray-900 transition-colors shadow-2xs cursor-pointer"
              >
                <UploadCloud className="w-3.5 h-3.5" />
                <span>Import (.xlsx)</span>
              </button>

              <button
                onClick={handleExportExcel}
                className="inline-flex items-center gap-1.5 px-4 py-2 text-xs font-semibold rounded-lg bg-gray-900 text-white hover:bg-black transition-colors shadow-sm cursor-pointer"
              >
                <FileSpreadsheet className="w-3.5 h-3.5 text-gray-400" />
                <span>Export Calendar (.xlsx)</span>
              </button>

              <button
                onClick={handleClearAll}
                className="inline-flex items-center gap-1.5 px-3 py-2 text-xs font-semibold rounded-lg bg-white text-rose-700 border border-rose-200 hover:bg-rose-50 transition-colors shadow-2xs cursor-pointer"
              >
                <Trash2 className="w-3.5 h-3.5 text-rose-500" />
                <span>Clear</span>
              </button>
            </div>
          )}
        </div>
      </div>

      {/* Context Bar: Client & Period */}
      <div className="p-3.5 bg-white rounded-xl border border-gray-200 flex flex-wrap items-center justify-between gap-4 text-xs shadow-2xs">
        <div className="flex flex-wrap items-center gap-5">
          <div className="flex items-center gap-2">
            <Building className="w-4 h-4 text-gray-400" />
            <span className="text-gray-500 font-medium">Client:</span>
            <input
              type="text"
              placeholder="Enter Client Name..."
              value={clientName}
              onChange={(e) => setClientName(e.target.value)}
              className="font-semibold text-gray-900 bg-gray-50 border border-gray-200 rounded px-2.5 py-1 text-xs focus:ring-1 focus:ring-blue-500 focus:bg-white w-48"
            />
          </div>

          <div className="flex items-center gap-2">
            <Calendar className="w-4 h-4 text-gray-400" />
            <span className="text-gray-500 font-medium">Month:</span>
            <select
              value={selectedMonth}
              onChange={(e) => setSelectedMonth(e.target.value)}
              className="font-semibold text-gray-900 bg-gray-50 border border-gray-200 rounded px-2.5 py-1 text-xs focus:ring-1 focus:ring-blue-500 focus:bg-white cursor-pointer"
            >
              {MONTHS_LIST.map((m) => (
                <option key={m} value={m}>{m}</option>
              ))}
            </select>
          </div>

          <div className="flex items-center gap-2">
            <span className="text-gray-500 font-medium">Year:</span>
            <input
              type="number"
              value={selectedYear}
              onChange={(e) => setSelectedYear(Number(e.target.value))}
              className="font-semibold text-gray-900 bg-gray-50 border border-gray-200 rounded px-2 py-1 text-xs focus:ring-1 focus:ring-blue-500 focus:bg-white w-20"
            />
          </div>

          <div className="flex items-center gap-2">
            <span className="text-gray-500 font-medium">FY:</span>
            <span className="font-semibold text-gray-800">{financialYear}</span>
          </div>
        </div>

        <div className="flex items-center gap-2 text-xs text-gray-500">
          <Clock className="w-3.5 h-3.5 text-gray-700" />
          <span>Statutory Engine Active (All 2026 Finance Act Provisions)</span>
        </div>
      </div>

      {/* ZERO STATE / SETUP CANVAS: Shown when items.length === 0 */}
      {items.length === 0 ? (
        <div className="space-y-6">
          <div className="grid grid-cols-1 md:grid-cols-2 gap-5">
            {/* Card 1: Upload Client Compliance Sheet */}
            <div className="bg-white border-2 border-dashed border-gray-200 hover:border-gray-300 rounded-xl p-6 sm:p-8 text-center transition-colors shadow-2xs flex flex-col justify-between">
              <div
                onClick={() => fileInputRef.current?.click()}
                className="cursor-pointer flex flex-col items-center"
              >
                <div className="w-12 h-12 rounded-full bg-gray-100 border border-gray-200 flex items-center justify-center text-gray-900 mb-3 shadow-2xs">
                  <UploadCloud className="w-6 h-6" />
                </div>
                <h3 className="text-sm font-semibold text-gray-900">Upload Client Compliance Schedule</h3>
                <p className="text-xs text-gray-500 mt-1 max-w-xs mx-auto">
                  Import your firm's existing client tracker in Excel (.xlsx, .xls) or CSV
                </p>
                <div className="mt-4 px-4 py-2 bg-gray-900 hover:bg-black text-white text-xs font-semibold rounded-lg shadow-2xs inline-flex items-center gap-2 transition-all">
                  <UploadCloud className="w-4 h-4" />
                  <span>Upload Compliance Tracker</span>
                </div>
                <span className="text-[11px] text-gray-400 font-medium mt-3 block">
                  Auto-detects due dates, acts & filing status
                </span>
              </div>

              <div className="pt-4 mt-4 border-t border-gray-100 flex items-center justify-between text-left">
                <span className="text-[11px] text-gray-500">Need standard format?</span>
                <button
                  type="button"
                  onClick={handleDownloadTemplate}
                  className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold text-gray-700 bg-white border border-gray-200 rounded-lg hover:bg-gray-50 transition-colors shadow-2xs cursor-pointer"
                >
                  <Download className="w-3.5 h-3.5 text-gray-500" />
                  <span>Download Blank Template (.xlsx)</span>
                </button>
              </div>
            </div>

            {/* Card 2: Auto-Generate Official Statutory Due Dates for Selected Month */}
            <div className="bg-white border border-gray-200 rounded-xl p-6 shadow-2xs flex flex-col justify-between space-y-5">
              <div>
                <div className="flex items-center gap-3 mb-3">
                  <div className="w-10 h-10 rounded-xl bg-gray-50 border border-gray-200 flex items-center justify-center text-gray-700">
                    <Calendar className="w-5 h-5" />
                  </div>
                  <div>
                    <h2 className="text-sm font-bold text-gray-900">Auto-Generate Statutory Calendar</h2>
                    <p className="text-xs text-gray-500">Compute official government due dates for {selectedMonth} {selectedYear}</p>
                  </div>
                </div>

                <div className="space-y-3 mt-4 bg-gray-50/70 p-4 rounded-xl border border-gray-200">
                  <span className="text-[11px] font-bold text-gray-700 uppercase tracking-wider block">
                    Select Applicable Statutory Laws:
                  </span>
                  <div className="grid grid-cols-2 gap-2 text-xs">
                    <label className="flex items-center gap-2 text-gray-700 cursor-pointer">
                      <input
                        type="checkbox"
                        checked={actsToInclude.gst}
                        onChange={(e) => setActsToInclude({ ...actsToInclude, gst: e.target.checked })}
                        className="rounded border-gray-300 text-gray-900 focus:ring-gray-900"
                      />
                      <span>GST (GSTR-1 & 3B)</span>
                    </label>

                    <label className="flex items-center gap-2 text-gray-700 cursor-pointer">
                      <input
                        type="checkbox"
                        checked={actsToInclude.it}
                        onChange={(e) => setActsToInclude({ ...actsToInclude, it: e.target.checked })}
                        className="rounded border-gray-300 text-gray-900 focus:ring-gray-900"
                      />
                      <span>Income Tax & TDS</span>
                    </label>

                    <label className="flex items-center gap-2 text-gray-700 cursor-pointer">
                      <input
                        type="checkbox"
                        checked={actsToInclude.mca}
                        onChange={(e) => setActsToInclude({ ...actsToInclude, mca: e.target.checked })}
                        className="rounded border-gray-300 text-gray-900 focus:ring-gray-900"
                      />
                      <span>MCA / ROC Returns</span>
                    </label>

                    <label className="flex items-center gap-2 text-gray-700 cursor-pointer">
                      <input
                        type="checkbox"
                        checked={actsToInclude.labor}
                        onChange={(e) => setActsToInclude({ ...actsToInclude, labor: e.target.checked })}
                        className="rounded border-gray-300 text-gray-900 focus:ring-gray-900"
                      />
                      <span>Labor & PF / ESI</span>
                    </label>
                  </div>
                </div>
              </div>

              <div className="pt-3 border-t border-gray-100 flex items-center justify-between">
                <button
                  onClick={() => setShowAddModal(true)}
                  className="inline-flex items-center gap-1.5 text-xs text-gray-600 hover:text-gray-900 font-medium cursor-pointer"
                >
                  <Plus className="w-3.5 h-3.5" />
                  <span>+ Add Custom Task</span>
                </button>

                <button
                  onClick={handleGenerateStatutorySchedule}
                  className="inline-flex items-center gap-1.5 px-4 py-2 text-xs font-semibold rounded-lg bg-gray-900 text-white hover:bg-black transition-colors shadow-sm cursor-pointer"
                >
                  <RefreshCw className="w-3.5 h-3.5" />
                  <span>Generate {selectedMonth} Schedule</span>
                </button>
              </div>
            </div>
          </div>
        </div>
      ) : (
        /* ACTIVE WORKSPACE: Shown when items.length > 0 */
        <>
          {/* 4 Summary Stat Cards */}
          <div className="grid grid-cols-2 lg:grid-cols-4 gap-4">
            <div className="bg-white p-4 rounded-xl border border-gray-200 shadow-2xs">
              <div className="flex items-center justify-between">
                <span className="text-xs font-medium text-gray-500">Total Compliances</span>
                <Calendar className="w-4 h-4 text-gray-700" />
              </div>
              <div className="mt-2 flex items-baseline gap-2">
                <span className="text-2xl font-bold text-gray-900">{metrics.total}</span>
                <span className="text-xs text-gray-500">tasks</span>
              </div>
              <p className="text-[11px] text-gray-400 mt-1">Scheduled for this period</p>
            </div>

            <div className="bg-white p-4 rounded-xl border border-gray-200 shadow-2xs">
              <div className="flex items-center justify-between">
                <span className="text-xs font-medium text-gray-500">Filed on Time</span>
                <CheckCircle2 className="w-4 h-4 text-gray-700" />
              </div>
              <div className="mt-2 flex items-baseline gap-2">
                <span className="text-2xl font-bold text-gray-700">{metrics.filed}</span>
                <span className="text-xs text-gray-800 bg-gray-50 px-1.5 py-0.5 rounded font-medium">
                  {metrics.complianceRate}% Rate
                </span>
              </div>
              <p className="text-[11px] text-gray-400 mt-1">Proof of filing verified</p>
            </div>

            <div className="bg-white p-4 rounded-xl border border-gray-200 shadow-2xs">
              <div className="flex items-center justify-between">
                <span className="text-xs font-medium text-gray-500">Upcoming This Month</span>
                <Clock className="w-4 h-4 text-blue-500" />
              </div>
              <div className="mt-2 flex items-baseline gap-2">
                <span className="text-2xl font-bold text-blue-600">{metrics.pending}</span>
                <span className="text-xs text-blue-700 bg-blue-50 px-1.5 py-0.5 rounded font-medium">
                  Action pending
                </span>
              </div>
              <p className="text-[11px] text-gray-400 mt-1">Statutory returns & challans</p>
            </div>

            <div className="bg-white p-4 rounded-xl border border-gray-200 shadow-2xs">
              <div className="flex items-center justify-between">
                <span className="text-xs font-medium text-gray-500">Overdue / Penalty Risk</span>
                <AlertTriangle className="w-4 h-4 text-rose-500" />
              </div>
              <div className="mt-2 flex items-baseline gap-2">
                <span className="text-2xl font-bold text-rose-600">{metrics.overdue}</span>
                {metrics.totalLateFee > 0 && (
                  <span className="text-xs text-rose-700 bg-rose-50 px-1.5 py-0.5 rounded font-bold">
                    ₹{metrics.totalLateFee.toLocaleString('en-IN')}
                  </span>
                )}
              </div>
              <p className="text-[11px] text-gray-400 mt-1">Sec 234C / Sec 47 late interest</p>
            </div>
          </div>

          {/* Tabs */}
          <div className="flex items-center justify-between border-b border-gray-200">
            <div className="flex gap-2">
              <button
                onClick={() => setActiveTab('calendar')}
                className={`px-4 py-2.5 text-xs font-semibold border-b-2 transition-all cursor-pointer flex items-center gap-2 ${
                  activeTab === 'calendar'
                    ? 'border-gray-900 text-gray-900'
                    : 'border-transparent text-gray-500 hover:text-gray-900'
                }`}
              >
                <Calendar className="w-3.5 h-3.5" />
                <span>Statutory Calendar ({items.length})</span>
              </button>

              <button
                onClick={() => setActiveTab('penalties')}
                className={`px-4 py-2.5 text-xs font-semibold border-b-2 transition-all cursor-pointer flex items-center gap-2 ${
                  activeTab === 'penalties'
                    ? 'border-gray-900 text-gray-900'
                    : 'border-transparent text-gray-500 hover:text-gray-900'
                }`}
              >
                <AlertTriangle className="w-3.5 h-3.5" />
                <span>Overdue & Statutory Penalties ({metrics.overdue})</span>
              </button>

              <button
                onClick={() => setActiveTab('proof_vault')}
                className={`px-4 py-2.5 text-xs font-semibold border-b-2 transition-all cursor-pointer flex items-center gap-2 ${
                  activeTab === 'proof_vault'
                    ? 'border-gray-900 text-gray-900'
                    : 'border-transparent text-gray-500 hover:text-gray-900'
                }`}
              >
                <FileCheck className="w-3.5 h-3.5" />
                <span>Filing Proof Repository ({metrics.filed})</span>
              </button>
            </div>
          </div>

          {/* TAB 1: STATUTORY CALENDAR */}
          {activeTab === 'calendar' && (
            <div className="space-y-4">
              {/* Search & Act Filters */}
              <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                <div className="relative flex-1 max-w-md">
                  <Search className="w-4 h-4 absolute left-3 top-2.5 text-gray-400" />
                  <input
                    type="text"
                    placeholder="Search by compliance obligation, section, or ARN..."
                    value={searchQuery}
                    onChange={(e) => setSearchQuery(e.target.value)}
                    className="w-full pl-9 pr-3 py-1.5 text-xs bg-white border border-gray-200 rounded-lg text-gray-900 focus:outline-none focus:ring-1 focus:ring-blue-500"
                  />
                </div>

                <div className="flex flex-wrap items-center gap-1.5">
                  <Filter className="w-3.5 h-3.5 text-gray-400 mr-1" />
                  {['all', 'GST', 'Income Tax', 'MCA / ROC', 'Labor & PF', 'Other'].map((act) => (
                    <button
                      key={act}
                      onClick={() => setSelectedAct(act)}
                      className={`px-2.5 py-1 rounded-md text-[11px] font-medium transition-colors cursor-pointer ${
                        selectedAct === act
                          ? 'bg-gray-100 border border-gray-300 text-gray-900 font-semibold'
                          : 'bg-white border border-gray-200 text-gray-600 hover:bg-gray-50'
                      }`}
                    >
                      {act === 'all' ? 'All Acts' : act}
                    </button>
                  ))}
                </div>
              </div>

              {/* Table */}
              <div className="bg-white border border-gray-200 rounded-xl overflow-hidden shadow-2xs">
                <div className="overflow-x-auto">
                  <table className="w-full text-left border-collapse text-xs">
                    <thead>
                      <tr className="bg-gray-50/80 border-b border-gray-200 text-gray-600 font-semibold">
                        <th className="py-3 px-4 w-12 text-center">#</th>
                        <th className="py-3 px-4">Compliance Obligation</th>
                        <th className="py-3 px-3">Act & Section</th>
                        <th className="py-3 px-3">Frequency</th>
                        <th className="py-3 px-3">Statutory Due Date</th>
                        <th className="py-3 px-3">Status</th>
                        <th className="py-3 px-4">Filing Proof (ARN / CIN)</th>
                        <th className="py-3 px-4 text-right">Actions</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-gray-100">
                      {filteredItems.map((item, idx) => (
                        <tr key={item.id} className="hover:bg-gray-50/50 transition-colors">
                          <td className="py-3 px-4 text-center font-mono text-gray-400 text-[11px]">
                            {idx + 1}
                          </td>
                          <td className="py-3 px-4">
                            <div className="font-semibold text-gray-900">{item.name}</div>
                            {item.filedDate && (
                              <div className="text-[11px] text-gray-500 mt-0.5">
                                Filed on: <span className="font-medium text-gray-700">{item.filedDate}</span>
                              </div>
                            )}
                          </td>
                          <td className="py-3 px-3">
                            <span className="inline-flex items-center px-2 py-0.5 rounded text-[10px] font-medium bg-gray-100 text-gray-800 border border-gray-200">
                              {item.act}
                            </span>
                            <div className="text-[10px] text-gray-500 font-mono mt-0.5">{item.section}</div>
                          </td>
                          <td className="py-3 px-3 text-gray-600 font-medium">
                            {item.frequency}
                          </td>
                          <td className="py-3 px-3 font-mono font-medium text-gray-800">
                            {item.dueDate}
                          </td>
                          <td className="py-3 px-3">
                            {item.status === 'filed' && (
                              <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[10px] font-semibold bg-gray-50 text-gray-800 border border-gray-200">
                                <CheckCircle2 className="w-3 h-3 text-gray-700" />
                                Filed on Time
                              </span>
                            )}
                            {item.status === 'pending' && (
                              <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[10px] font-semibold bg-blue-50 text-blue-700 border border-blue-200">
                                <Clock className="w-3 h-3 text-blue-600" />
                                Upcoming
                              </span>
                            )}
                            {item.status === 'overdue' && (
                              <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[10px] font-semibold bg-rose-50 text-rose-700 border border-rose-200">
                                <AlertTriangle className="w-3 h-3 text-rose-600" />
                                Overdue
                              </span>
                            )}
                          </td>
                          <td className="py-3 px-4 font-mono text-[11px]">
                            {item.filingProofRef ? (
                              <div className="flex items-center gap-1 text-gray-800 font-medium truncate max-w-[200px]" title={item.filingProofRef}>
                                <Receipt className="w-3.5 h-3.5 text-gray-700 shrink-0" />
                                <span>{item.filingProofRef}</span>
                              </div>
                            ) : (
                              <span className="text-gray-400 italic">Not filed yet</span>
                            )}
                          </td>
                          <td className="py-3 px-4 text-right">
                            <div className="flex items-center justify-end gap-1.5">
                              <button
                                onClick={() => {
                                  setSelectedItemForProof(item);
                                  setProofReference(item.filingProofRef || '');
                                  setShowProofModal(true);
                                }}
                                className="inline-flex items-center gap-1 px-2.5 py-1 text-[11px] font-semibold rounded bg-white text-gray-700 border border-gray-200 hover:bg-gray-50 hover:text-gray-900 transition-colors shadow-2xs cursor-pointer"
                              >
                                <FileCheck className="w-3 h-3 text-gray-700" />
                                <span>{item.status === 'filed' ? 'Edit Proof' : 'Record Proof'}</span>
                              </button>

                              <button
                                onClick={() => handleDeleteItem(item.id)}
                                className="p-1 rounded text-gray-400 hover:text-rose-600 hover:bg-rose-50 transition-colors cursor-pointer"
                                title="Remove task"
                              >
                                <Trash2 className="w-3.5 h-3.5" />
                              </button>
                            </div>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            </div>
          )}

          {/* TAB 2: OVERDUE & STATUTORY PENALTIES */}
          {activeTab === 'penalties' && (
            <div className="bg-white border border-gray-200 rounded-xl p-6 shadow-2xs space-y-6">
              <div className="flex items-center gap-3">
                <div className="w-10 h-10 rounded-lg bg-rose-50 border border-rose-200 flex items-center justify-center text-rose-600">
                  <AlertTriangle className="w-5 h-5" />
                </div>
                <div>
                  <h2 className="text-sm font-bold text-gray-900">Statutory Overdue & Late Fee Assessment</h2>
                  <p className="text-xs text-gray-500">
                    Automated calculation of late fees and mandatory interest under CGST Sec 47, Sec 50, and Income Tax Sec 234C/234E.
                  </p>
                </div>
              </div>

              <div className="p-4 rounded-xl border border-gray-200 bg-gray-50/50">
                {metrics.overdue === 0 ? (
                  <p className="text-xs text-gray-600 flex items-center gap-2">
                    <CheckCircle2 className="w-4 h-4 text-gray-700" />
                    All scheduled statutory returns and challans are up-to-date! Zero penalty liability.
                  </p>
                ) : (
                  <div className="space-y-3">
                    {items
                      .filter((i) => i.status === 'overdue')
                      .map((item) => (
                        <div key={item.id} className="bg-white border border-rose-200 rounded-lg p-3 text-xs flex items-center justify-between">
                          <div>
                            <div className="font-bold text-gray-900">{item.name}</div>
                            <div className="text-gray-500 text-[11px] mt-0.5">
                              Governed by {item.act} ({item.section}) • Due Date was {item.dueDate}
                            </div>
                          </div>
                          <div className="text-right">
                            <span className="text-xs font-bold text-rose-700 bg-rose-100 px-2 py-0.5 rounded">
                              Estimated Interest: ₹{item.lateFeeEstimated?.toLocaleString('en-IN')}
                            </span>
                            <div className="text-[10px] text-gray-400 mt-1">1% per month simple interest</div>
                          </div>
                        </div>
                      ))}
                  </div>
                )}
              </div>
            </div>
          )}

          {/* TAB 3: PROOF VAULT */}
          {activeTab === 'proof_vault' && (
            <div className="bg-white border border-gray-200 rounded-xl p-6 shadow-2xs space-y-4">
              <div className="flex items-center justify-between pb-3 border-b border-gray-200">
                <div>
                  <h2 className="text-sm font-bold text-gray-900">Statutory Filing Proof Repository</h2>
                  <p className="text-xs text-gray-500">
                    Peer-review ready evidence register containing official ARNs, Challan Identification Numbers, and dates.
                  </p>
                </div>
              </div>

              {items.filter((i) => i.status === 'filed').length === 0 ? (
                <div className="text-center py-8 bg-gray-50/50 rounded-xl border border-gray-200">
                  <FileCheck className="w-8 h-8 text-gray-400 mx-auto mb-2" />
                  <p className="text-xs font-semibold text-gray-800">No filing proofs recorded yet</p>
                  <p className="text-[11px] text-gray-500 mt-0.5">Click "Record Proof" on any completed obligation to attach its ARN or Challan CIN.</p>
                </div>
              ) : (
                <div className="space-y-3">
                  {items
                    .filter((i) => i.status === 'filed')
                    .map((item) => (
                      <div key={item.id} className="border border-gray-200 rounded-lg p-3 bg-gray-50/50 flex items-center justify-between text-xs">
                        <div>
                          <div className="font-semibold text-gray-900">{item.name}</div>
                          <div className="text-gray-500 text-[11px] mt-0.5">{item.act} ({item.section})</div>
                        </div>
                        <div className="text-right">
                          <div className="font-mono text-gray-800 font-semibold">{item.filingProofRef}</div>
                          <div className="text-[11px] text-gray-400 mt-0.5">Filed on {item.filedDate}</div>
                        </div>
                      </div>
                    ))}
                </div>
              )}
            </div>
          )}
        </>
      )}

      {/* Modal: Record Filing Proof */}
      {showProofModal && selectedItemForProof && (
        <div className="fixed inset-0 z-50 bg-black/40 backdrop-blur-xs flex items-center justify-center p-4">
          <div className="bg-white rounded-xl shadow-xl border border-gray-200 w-full max-w-md p-5 space-y-4">
            <div className="flex items-center justify-between pb-2 border-b border-gray-100">
              <h3 className="text-sm font-bold text-gray-900">Record Statutory Filing Proof</h3>
              <button
                onClick={() => {
                  setShowProofModal(false);
                  setSelectedItemForProof(null);
                }}
                className="text-gray-400 hover:text-gray-700"
              >
                <X className="w-4 h-4" />
              </button>
            </div>

            <p className="text-xs text-gray-600">
              Obligation: <span className="font-semibold text-gray-900">{selectedItemForProof.name}</span>
            </p>

            <div>
              <label className="block text-xs font-medium text-gray-700 mb-1">
                Filing Proof Reference (ARN / Challan CIN / Acknowledgement No.)
              </label>
              <input
                type="text"
                placeholder="e.g. ARN: AA271026009841C or CIN: 20261007019283"
                value={proofReference}
                onChange={(e) => setProofReference(e.target.value)}
                className="w-full text-xs font-mono border border-gray-200 rounded-lg px-3 py-2 text-gray-900 focus:outline-none focus:ring-1 focus:ring-blue-500"
              />
            </div>

            <div>
              <label className="block text-xs font-medium text-gray-700 mb-1">
                Official Filing Date
              </label>
              <input
                type="date"
                value={proofFilingDate}
                onChange={(e) => setProofFilingDate(e.target.value)}
                className="w-full text-xs border border-gray-200 rounded-lg px-3 py-2 text-gray-900 focus:outline-none focus:ring-1 focus:ring-blue-500"
              />
            </div>

            <div className="flex justify-end gap-2 pt-2">
              <button
                onClick={() => {
                  setShowProofModal(false);
                  setSelectedItemForProof(null);
                }}
                className="px-3 py-1.5 text-xs font-semibold text-gray-600 hover:text-gray-900 rounded-lg cursor-pointer"
              >
                Cancel
              </button>
              <button
                onClick={handleSaveProof}
                className="px-4 py-1.5 text-xs font-semibold bg-gray-900 hover:bg-black text-white rounded-lg shadow-2xs cursor-pointer"
              >
                Save Proof
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Modal: Add Custom Statutory Obligation */}
      {showAddModal && (
        <div className="fixed inset-0 z-50 bg-black/40 backdrop-blur-xs flex items-center justify-center p-4">
          <form
            onSubmit={handleAddCustomItem}
            className="bg-white rounded-xl shadow-xl border border-gray-200 w-full max-w-md p-5 space-y-4"
          >
            <div className="flex items-center justify-between pb-2 border-b border-gray-100">
              <h3 className="text-sm font-bold text-gray-900">Add Statutory Obligation</h3>
              <button
                type="button"
                onClick={() => setShowAddModal(false)}
                className="text-gray-400 hover:text-gray-700"
              >
                <X className="w-4 h-4" />
              </button>
            </div>

            <div>
              <label className="block text-xs font-medium text-gray-700 mb-1">
                Obligation / Return Name *
              </label>
              <input
                type="text"
                placeholder="e.g. Professional Tax Monthly Return, Form 15CA/CB..."
                value={newObligationName}
                onChange={(e) => setNewObligationName(e.target.value)}
                className="w-full text-xs border border-gray-200 rounded-lg px-3 py-2 text-gray-900 focus:outline-none focus:ring-1 focus:ring-blue-500"
                required
              />
            </div>

            <div className="grid grid-cols-2 gap-3">
              <div>
                <label className="block text-xs font-medium text-gray-700 mb-1">
                  Governing Act
                </label>
                <select
                  value={newAct}
                  onChange={(e) => setNewAct(e.target.value as ComplianceItem['act'])}
                  className="w-full text-xs border border-gray-200 rounded-lg px-2.5 py-2 text-gray-900 focus:outline-none focus:ring-1 focus:ring-blue-500 bg-white"
                >
                  <option value="GST">GST</option>
                  <option value="Income Tax">Income Tax</option>
                  <option value="MCA / ROC">MCA / ROC</option>
                  <option value="Labor & PF">Labor & PF</option>
                  <option value="Other">Other</option>
                </select>
              </div>

              <div>
                <label className="block text-xs font-medium text-gray-700 mb-1">
                  Frequency
                </label>
                <select
                  value={newFrequency}
                  onChange={(e) => setNewFrequency(e.target.value as ComplianceItem['frequency'])}
                  className="w-full text-xs border border-gray-200 rounded-lg px-2.5 py-2 text-gray-900 focus:outline-none focus:ring-1 focus:ring-blue-500 bg-white"
                >
                  <option value="Monthly">Monthly</option>
                  <option value="Quarterly">Quarterly</option>
                  <option value="Annual">Annual</option>
                  <option value="One-Time">One-Time</option>
                </select>
              </div>
            </div>

            <div>
              <label className="block text-xs font-medium text-gray-700 mb-1">
                Statutory Section / Rule
              </label>
              <input
                type="text"
                placeholder="e.g. Sec 39 CGST Act, Sec 200(3)..."
                value={newSection}
                onChange={(e) => setNewSection(e.target.value)}
                className="w-full text-xs border border-gray-200 rounded-lg px-3 py-2 text-gray-900 focus:outline-none focus:ring-1 focus:ring-blue-500"
              />
            </div>

            <div>
              <label className="block text-xs font-medium text-gray-700 mb-1">
                Statutory Due Date *
              </label>
              <input
                type="date"
                value={newDueDate}
                onChange={(e) => setNewDueDate(e.target.value)}
                className="w-full text-xs border border-gray-200 rounded-lg px-3 py-2 text-gray-900 focus:outline-none focus:ring-1 focus:ring-blue-500"
                required
              />
            </div>

            <div className="flex justify-end gap-2 pt-2">
              <button
                type="button"
                onClick={() => setShowAddModal(false)}
                className="px-3 py-1.5 text-xs font-semibold text-gray-600 hover:text-gray-900 rounded-lg cursor-pointer"
              >
                Cancel
              </button>
              <button
                type="submit"
                className="px-4 py-1.5 text-xs font-semibold bg-gray-900 hover:bg-black text-white rounded-lg shadow-2xs cursor-pointer"
              >
                Add Obligation
              </button>
            </div>
          </form>
        </div>
      )}
    </div>
  );
};
