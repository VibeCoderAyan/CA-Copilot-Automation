import React, { useState, useMemo } from 'react';
import * as XLSX from 'xlsx';
import {
  ArrowLeft,
  Calendar,
  CheckCircle2,
  Clock,
  AlertTriangle,
  AlertCircle,
  FileSpreadsheet,
  Printer,
  Copy,
  Plus,
  Building,
  Sliders,
  Filter,
  RefreshCw,
  Search,
  Check,
  ChevronRight,
  User,
  Paperclip,
  X,
} from 'lucide-react';
import { InnerToolBadge } from '../common/InnerToolBadge';
import { ToolHeader } from '../common/ToolHeader';
import {
  type ClosingTask,
  type ClosingTaskStatus,
  type ClosingCategory,
  type MonthEndClosingState,
  type ClosingMetrics,
  getInitialClosingTasks,
  evaluateClosingMetrics,
  duplicateChecklistForNewPeriod,
  exportMonthEndClosingWorkbook,
} from '../../utils/monthEndClosingEngine';

interface MonthEndClosingToolProps {
  onBack: () => void;
  onShowToast: (title: string, description?: string, type?: 'success' | 'error' | 'info') => void;
}

export const MonthEndClosingTool: React.FC<MonthEndClosingToolProps> = ({
  onBack,
  onShowToast,
}) => {
  const [companyName, setCompanyName] = useState('');
  const [accountingPeriod, setAccountingPeriod] = useState('September 2026');
  const [financialYear, setFinancialYear] = useState('FY 2026-27');
  const [targetClosingDate, setTargetClosingDate] = useState('10-10-2026');

  const [tasks, setTasks] = useState<ClosingTask[]>([]);
  const [activeTab, setActiveTab] = useState<'checklist' | 'outstanding' | 'printable_summary'>('checklist');
  const [selectedCategory, setSelectedCategory] = useState<string>('all');
  const [selectedStatus, setSelectedStatus] = useState<string>('all');
  const [searchQuery, setSearchQuery] = useState('');

  // Add custom task modal
  const [showAddModal, setShowAddModal] = useState(false);
  const [newTaskTitle, setNewTaskTitle] = useState('');
  const [newTaskCategory, setNewTaskCategory] = useState<ClosingCategory>('Cash & Bank');
  const [newTaskAssignee, setNewTaskAssignee] = useState('');
  const [newTaskDueDate, setNewTaskDueDate] = useState('10-10-2026');
  const [newTaskDesc, setNewTaskDesc] = useState('');

  // Evaluate metrics
  const metrics: ClosingMetrics = useMemo(() => {
    return evaluateClosingMetrics(tasks);
  }, [tasks]);

  const state: MonthEndClosingState = {
    companyName,
    accountingPeriod,
    financialYear,
    targetClosingDate,
    tasks,
  };

  // Update task status
  const handleStatusChange = (taskId: string, newStatus: ClosingTaskStatus) => {
    setTasks((prev) =>
      prev.map((t) => {
        if (t.id === taskId) {
          const completedDate =
            newStatus === 'completed' ? new Date().toLocaleDateString('en-IN') : undefined;
          return {
            ...t,
            status: newStatus,
            completedDate,
            history: [
              ...t.history,
              {
                timestamp: new Date().toISOString(),
                action: `Status changed to ${newStatus.toUpperCase()}`,
                changedBy: 'Audit Senior / User',
              },
            ],
          };
        }
        return t;
      })
    );
    onShowToast('Status Updated', `Task marked as ${newStatus.replace(/_/g, ' ').toUpperCase()}`, 'info');
  };

  // Update notes
  const handleNotesChange = (taskId: string, notes: string) => {
    setTasks((prev) =>
      prev.map((t) => (t.id === taskId ? { ...t, notes } : t))
    );
  };

  // Duplicate for Next Month
  const handleDuplicateNextMonth = () => {
    const nextPeriod = 'October 2026';
    const duplicated = duplicateChecklistForNewPeriod(tasks, nextPeriod, 30);
    setTasks(duplicated);
    setAccountingPeriod(nextPeriod);
    setTargetClosingDate('10-11-2026');
    onShowToast(
      'New Month Checklist Created',
      `Duplicated checklist for ${nextPeriod}. All tasks reset to Not Started with rolled-forward due dates.`,
      'success'
    );
  };

  // Add custom task
  const handleAddTask = () => {
    if (!newTaskTitle.trim()) {
      onShowToast('Missing Title', 'Please enter a title for the activity.', 'error');
      return;
    }

    const newTask: ClosingTask = {
      id: `custom-task-${Date.now()}`,
      category: newTaskCategory,
      title: newTaskTitle.trim(),
      description: newTaskDesc.trim() || 'Custom closing review activity',
      assignee: newTaskAssignee.trim() || 'Unassigned',
      dueDate: newTaskDueDate,
      status: 'not_started',
      priority: 'medium',
      notes: '',
      history: [
        {
          timestamp: new Date().toISOString(),
          action: 'Created custom task',
          changedBy: 'User',
        },
      ],
    };

    setTasks((prev) => [newTask, ...prev]);
    setShowAddModal(false);
    setNewTaskTitle('');
    setNewTaskDesc('');
    onShowToast('Task Added', `Added "${newTask.title}" to closing checklist`, 'success');
  };

  // Export Excel workbook
  const handleExportExcel = () => {
    try {
      const wb = exportMonthEndClosingWorkbook(state, metrics);
      XLSX.writeFile(wb, `MonthEnd_Closing_${companyName.replace(/\s+/g, '_')}_${accountingPeriod.replace(/\s+/g, '_')}.xlsx`);
      onShowToast('Workbook Exported', 'Exported complete closing checklist and outstanding items workbook.', 'success');
    } catch (err: any) {
      onShowToast('Export Error', err.message || 'Failed to generate Excel file', 'error');
    }
  };

  // Print PDF summary
  const handlePrint = () => {
    setActiveTab('printable_summary');
    setTimeout(() => {
      window.print();
    }, 300);
  };

  // Filter tasks
  const filteredTasks = useMemo(() => {
    return tasks.filter((t) => {
      if (selectedCategory !== 'all' && t.category !== selectedCategory) return false;
      if (selectedStatus !== 'all' && t.status !== selectedStatus) return false;

      if (searchQuery.trim()) {
        const q = searchQuery.toLowerCase();
        const title = t.title.toLowerCase();
        const desc = t.description.toLowerCase();
        const ass = t.assignee.toLowerCase();
        return title.includes(q) || desc.includes(q) || ass.includes(q);
      }
      return true;
    });
  }, [tasks, selectedCategory, selectedStatus, searchQuery]);

  return (
    <div className="space-y-6">
      <ToolHeader
        toolId="month-end-closing"
        title="Financial Period Close Assistant"
        description="Track recurring closing activities across BRS, cut-offs, provisions, and statutory reviews with audit sign-off memo."
        badge="Workflow"
        onBack={onBack}
        actionButtons={
          <div className="flex flex-wrap items-center gap-2">
            {tasks.length === 0 ? (
              <button
                onClick={() => {
                  setTasks(getInitialClosingTasks());
                  onShowToast('Standard Closing Checklist Initialized', '15 standard activities loaded.', 'success');
                }}
                className="px-3.5 py-1.5 text-xs font-semibold rounded-md bg-gray-900 text-white hover:bg-black transition-colors flex items-center gap-1.5 shadow-sm cursor-pointer"
              >
                <CheckCircle2 className="w-3.5 h-3.5 text-emerald-400" />
                Initialize Standard Checklist
              </button>
            ) : (
              <>
                <button
                  onClick={handleDuplicateNextMonth}
                  className="px-3 py-1.5 text-xs font-semibold rounded-md bg-gray-50 text-gray-900 border border-gray-300/30 hover:bg-gray-100 transition-colors flex items-center gap-1.5 cursor-pointer"
                >
                  <Copy className="w-3.5 h-3.5" />
                  Duplicate for Next Month
                </button>

                <button
                  onClick={handlePrint}
                  className="px-3 py-1.5 text-xs font-semibold rounded-md bg-white text-gray-900 border border-gray-200 hover:bg-gray-50 transition-colors flex items-center gap-1.5 cursor-pointer"
                >
                  <Printer className="w-3.5 h-3.5 text-gray-600" />
                  PDF Summary
                </button>

                <button
                  onClick={handleExportExcel}
                  className="px-3.5 py-1.5 text-xs font-semibold rounded-md bg-gray-900 text-white hover:bg-black transition-colors flex items-center gap-1.5 shadow-sm cursor-pointer"
                >
                  <FileSpreadsheet className="w-3.5 h-3.5" />
                  Export Excel Checklist (.xlsx)
                </button>
              </>
            )}

            <button
              onClick={() => setShowAddModal(true)}
              className="px-3 py-1.5 text-xs font-semibold rounded-md bg-white text-gray-900 border border-gray-200 hover:bg-gray-50 transition-colors flex items-center gap-1.5 shadow-sm cursor-pointer"
            >
              <Plus className="w-3.5 h-3.5 text-gray-900" />
              Add Task
            </button>
          </div>
        }
      />

      {/* Company & Period Row */}
      <div className="bg-white border border-gray-200 rounded-xl p-4 shadow-2xs flex flex-wrap items-center gap-4 text-xs">
        <div className="flex items-center gap-2">
          <Building className="w-4 h-4 text-gray-600" />
          <span className="font-semibold text-gray-600">Entity:</span>
          <input
            type="text"
            value={companyName}
            onChange={(e) => setCompanyName(e.target.value)}
            className="px-2 py-1 border border-gray-200 rounded text-xs font-medium focus:border-gray-300 outline-none w-52"
          />
        </div>

        <div className="flex items-center gap-2">
          <Calendar className="w-4 h-4 text-gray-600" />
          <span className="font-semibold text-gray-600">Closing Period:</span>
          <input
            type="text"
            value={accountingPeriod}
            onChange={(e) => setAccountingPeriod(e.target.value)}
            className="px-2 py-1 border border-gray-200 rounded text-xs font-medium focus:border-gray-300 outline-none w-36"
          />
        </div>

        <div className="flex items-center gap-2">
          <span className="font-semibold text-gray-600">Target Date:</span>
          <input
            type="text"
            value={targetClosingDate}
            onChange={(e) => setTargetClosingDate(e.target.value)}
            className="px-2 py-1 border border-gray-200 rounded text-xs font-medium focus:border-gray-300 outline-none w-28"
          />
        </div>
      </div>

      {/* Main Container */}
      <div className="space-y-6">
        {tasks.length === 0 ? (
          <div className="bg-white border border-dashed border-gray-300 rounded-xl p-12 text-center max-w-2xl mx-auto shadow-sm my-6">
            <div className="w-12 h-12 bg-gray-100 rounded-xl flex items-center justify-center mx-auto mb-4 text-gray-600">
              <Calendar className="w-6 h-6" />
            </div>
            <h3 className="text-sm font-bold text-gray-900 mb-1">No Closing Activities Scheduled Yet</h3>
            <p className="text-xs text-gray-500 mb-6 max-w-md mx-auto">
              Start your month-end closing by loading the ICAI-recommended 15-step checklist (covering BRS, GST vs Books, TDS, cut-off reviews, fixed assets, and provisions) or add custom tasks.
            </p>
            <div className="flex items-center justify-center gap-3">
              <button
                onClick={() => {
                  setTasks(getInitialClosingTasks());
                  onShowToast('Standard Closing Checklist Initialized', '15 standard activities loaded.', 'success');
                }}
                className="px-4 py-2 bg-gray-900 text-white rounded-lg text-xs font-semibold hover:bg-black transition-colors flex items-center gap-2 shadow-sm cursor-pointer"
              >
                <CheckCircle2 className="w-4 h-4 text-emerald-400" />
                Initialize Standard Closing Checklist
              </button>
              <button
                onClick={() => setShowAddModal(true)}
                className="px-4 py-2 bg-white text-gray-700 border border-gray-200 rounded-lg text-xs font-semibold hover:bg-gray-50 transition-colors flex items-center gap-2 cursor-pointer"
              >
                <Plus className="w-4 h-4" />
                Add Custom Activity
              </button>
            </div>
          </div>
        ) : (
          <>
            {/* KPI & Progress Cards */}
            <div className="grid grid-cols-2 md:grid-cols-5 gap-4">
          <div className="bg-white border border-gray-200 rounded-xl p-4 shadow-sm">
            <span className="text-xs font-semibold text-gray-600 uppercase">Closing Progress</span>
            <div className="flex items-baseline gap-2 mt-1">
              <span className="text-2xl font-bold text-gray-900">{metrics.completionPercentage.toFixed(0)}%</span>
              <span className="text-xs text-gray-600">({metrics.completedCount}/{metrics.totalTasks})</span>
            </div>
            <div className="w-full bg-gray-100 rounded-full h-1.5 mt-2 overflow-hidden">
              <div
                className="bg-gray-900 h-full rounded-full transition-all duration-300"
                style={{ width: `${metrics.completionPercentage}%` }}
              />
            </div>
          </div>

          <div className="bg-white border border-gray-200 rounded-xl p-4 shadow-sm">
            <span className="text-xs font-semibold text-gray-600 uppercase">In Progress</span>
            <div className="text-2xl font-bold text-amber-600 mt-1">
              {metrics.inProgressCount}
            </div>
            <div className="text-xs text-gray-600 mt-0.5">Active activities under review</div>
          </div>

          <div className="bg-white border border-gray-200 rounded-xl p-4 shadow-sm">
            <span className="text-xs font-semibold text-gray-600 uppercase">Blocked / Exceptions</span>
            <div className="text-2xl font-bold text-[#DC2626] mt-1">
              {metrics.blockedCount}
            </div>
            <div className="text-xs text-gray-600 mt-0.5">Awaiting information / partner</div>
          </div>

          <div className="bg-white border border-gray-200 rounded-xl p-4 shadow-sm">
            <span className="text-xs font-semibold text-gray-600 uppercase">Overdue Tasks</span>
            <div className="text-2xl font-bold text-[#DC2626] mt-1">
              {metrics.overdueCount}
            </div>
            <div className="text-xs text-gray-600 mt-0.5">Exceeded target closing date</div>
          </div>

          <div className="bg-white border border-gray-200 rounded-xl p-4 shadow-sm">
            <span className="text-xs font-semibold text-gray-600 uppercase">Not Started</span>
            <div className="text-2xl font-bold text-gray-600 mt-1">
              {metrics.notStartedCount}
            </div>
            <div className="text-xs text-gray-600 mt-0.5">Pending prior tasks completion</div>
          </div>
        </div>

        {/* Tab Selection */}
        <div className="flex flex-col sm:flex-row sm:items-center sm:justify-between gap-4 border-b border-gray-200 pb-1 print:hidden">
          <div className="flex items-center gap-2">
            <button
              onClick={() => setActiveTab('checklist')}
              className={`px-4 py-2 text-xs font-bold border-b-2 transition-colors ${
                activeTab === 'checklist'
                  ? 'border-gray-300 text-gray-900'
                  : 'border-transparent text-gray-600 hover:text-gray-900'
              }`}
            >
              Master Closing Checklist ({tasks.length})
            </button>
            <button
              onClick={() => setActiveTab('outstanding')}
              className={`px-4 py-2 text-xs font-bold border-b-2 transition-colors ${
                activeTab === 'outstanding'
                  ? 'border-[#DC2626] text-[#DC2626]'
                  : 'border-transparent text-gray-600 hover:text-gray-900'
              }`}
            >
              Outstanding & Blocked Items ({metrics.outstandingTasks.length})
            </button>
            <button
              onClick={() => setActiveTab('printable_summary')}
              className={`px-4 py-2 text-xs font-bold border-b-2 transition-colors ${
                activeTab === 'printable_summary'
                  ? 'border-gray-300 text-gray-900'
                  : 'border-transparent text-gray-600 hover:text-gray-900'
              }`}
            >
              Closing Summary Memo
            </button>
          </div>

          {activeTab === 'checklist' && (
            <div className="flex items-center gap-2">
              <div className="relative">
                <Search className="w-3.5 h-3.5 absolute left-2.5 top-1/2 -translate-y-1/2 text-gray-600" />
                <input
                  type="text"
                  placeholder="Search activity or assignee..."
                  value={searchQuery}
                  onChange={(e) => setSearchQuery(e.target.value)}
                  className="pl-8 pr-3 py-1.5 text-xs border border-gray-200 rounded-lg w-48 focus:border-gray-300 outline-none"
                />
              </div>

              <select
                value={selectedCategory}
                onChange={(e) => setSelectedCategory(e.target.value)}
                className="text-xs border border-gray-200 rounded-lg px-2 py-1.5 outline-none"
              >
                <option value="all">All Categories</option>
                <option value="Cash & Bank">Cash & Bank</option>
                <option value="Receivables & Payables">Receivables & Payables</option>
                <option value="General Ledger & Journals">General Ledger</option>
                <option value="Fixed Assets">Fixed Assets</option>
                <option value="Accruals & Cut-off">Accruals & Cut-off</option>
                <option value="Taxation & Statutory">Taxation</option>
                <option value="Financial Reporting">Financial Reporting</option>
              </select>

              <select
                value={selectedStatus}
                onChange={(e) => setSelectedStatus(e.target.value)}
                className="text-xs border border-gray-200 rounded-lg px-2 py-1.5 outline-none"
              >
                <option value="all">All Statuses</option>
                <option value="not_started">Not Started</option>
                <option value="in_progress">In Progress</option>
                <option value="blocked">Blocked</option>
                <option value="completed">Completed</option>
              </select>
            </div>
          )}
        </div>

        {/* TAB 1: Master Checklist Table */}
        {activeTab === 'checklist' && (
          <div className="bg-white border border-gray-200 rounded-xl overflow-hidden shadow-sm">
            <div className="overflow-x-auto">
              <table className="w-full text-xs text-left">
                <thead className="bg-gray-900 text-white font-semibold">
                  <tr>
                    <th className="p-3 w-10 text-center">#</th>
                    <th className="p-3">Category</th>
                    <th className="p-3">Activity / Task Details</th>
                    <th className="p-3">Priority</th>
                    <th className="p-3">Assignee</th>
                    <th className="p-3">Due Date</th>
                    <th className="p-3 text-center">Current Status</th>
                    <th className="p-3">Notes & Supporting File</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-[#E5E7EB]">
                  {filteredTasks.map((t, idx) => (
                    <tr
                      key={t.id}
                      className={`hover:bg-gray-50 transition-colors ${
                        t.status === 'blocked'
                          ? 'bg-red-50/30'
                          : t.isOverdue
                          ? 'bg-amber-50/20'
                          : ''
                      }`}
                    >
                      <td className="p-3 text-center text-gray-600 font-medium">{idx + 1}</td>
                      <td className="p-3 font-semibold text-gray-900">{t.category}</td>
                      <td className="p-3">
                        <div className="font-bold text-gray-900 text-xs">{t.title}</div>
                        <div className="text-[11px] text-gray-600 mt-0.5 max-w-md">{t.description}</div>
                      </td>
                      <td className="p-3">
                        <span
                          className={`px-2 py-0.5 rounded text-[10px] font-bold ${
                            t.priority === 'critical'
                              ? 'bg-red-100 text-red-800'
                              : t.priority === 'high'
                              ? 'bg-amber-100 text-amber-800'
                              : 'bg-gray-100 text-gray-700'
                          }`}
                        >
                          {t.priority.toUpperCase()}
                        </span>
                      </td>
                      <td className="p-3 text-gray-900 font-medium flex items-center gap-1.5 mt-2">
                        <User className="w-3.5 h-3.5 text-gray-600" />
                        {t.assignee}
                      </td>
                      <td className="p-3">
                        <div className="font-mono text-xs">{t.dueDate}</div>
                        {t.isOverdue && (
                          <span className="text-[10px] font-bold text-[#DC2626] block">OVERDUE</span>
                        )}
                        {t.completedDate && (
                          <span className="text-[10px] text-gray-900 block">Done: {t.completedDate}</span>
                        )}
                      </td>
                      <td className="p-3 text-center">
                        <select
                          value={t.status}
                          onChange={(e) => handleStatusChange(t.id, e.target.value as any)}
                          className={`text-xs font-bold border rounded px-2.5 py-1 outline-none ${
                            t.status === 'completed'
                              ? 'bg-gray-50 text-gray-900 border-gray-200'
                              : t.status === 'in_progress'
                              ? 'bg-amber-50 text-amber-800 border-amber-300'
                              : t.status === 'blocked'
                              ? 'bg-red-50 text-red-800 border-red-300'
                              : 'bg-gray-50 text-gray-700 border-gray-300'
                          }`}
                        >
                          <option value="not_started">Not Started</option>
                          <option value="in_progress">In Progress</option>
                          <option value="blocked">Blocked</option>
                          <option value="completed">Completed</option>
                        </select>
                      </td>
                      <td className="p-3">
                        <input
                          type="text"
                          value={t.notes}
                          onChange={(e) => handleNotesChange(t.id, e.target.value)}
                          placeholder="Add findings / notes..."
                          className="w-full text-xs border border-transparent hover:border-gray-200 focus:border-gray-300 rounded px-1.5 py-0.5 outline-none bg-transparent"
                        />
                        {t.supportingDocRef && (
                          <div className="flex items-center gap-1 text-[11px] text-gray-900 mt-1 font-mono">
                            <Paperclip className="w-3 h-3" />
                            {t.supportingDocRef}
                          </div>
                        )}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        )}

        {/* TAB 2: Outstanding & Blocked Items */}
        {activeTab === 'outstanding' && (
          <div className="space-y-4">
            <div className="p-4 bg-white border border-gray-200 rounded-xl shadow-sm">
              <h3 className="font-bold text-sm text-gray-900">Unresolved Closing Exceptions & Outstanding Tasks</h3>
              <p className="text-xs text-gray-600 mt-0.5">
                Review incomplete activities before financial books can be formally sealed for {accountingPeriod}
              </p>
            </div>

            <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
              {metrics.outstandingTasks.map((t) => (
                <div
                  key={t.id}
                  className={`border rounded-xl p-4 bg-white shadow-sm space-y-2 ${
                    t.status === 'blocked' ? 'border-red-300 bg-red-50/20' : 'border-gray-200'
                  }`}
                >
                  <div className="flex justify-between items-start">
                    <span className="font-bold text-xs text-gray-900">{t.category}</span>
                    <span
                      className={`px-2 py-0.5 rounded text-[10px] font-bold ${
                        t.status === 'blocked'
                          ? 'bg-red-100 text-red-800'
                          : 'bg-amber-100 text-amber-800'
                      }`}
                    >
                      {t.status.toUpperCase()}
                    </span>
                  </div>
                  <h4 className="font-bold text-sm text-gray-900">{t.title}</h4>
                  <p className="text-xs text-gray-600">{t.description}</p>
                  <div className="pt-2 border-t border-gray-100 flex justify-between items-center text-xs text-gray-600">
                    <span>Assignee: <strong className="text-gray-900">{t.assignee}</strong></span>
                    <span>Due: <strong className="text-gray-900">{t.dueDate}</strong></span>
                  </div>
                  {t.notes && (
                    <div className="p-2 rounded bg-gray-50 text-[11px] text-gray-900">
                      <strong>Blocker Reason:</strong> {t.notes}
                    </div>
                  )}
                </div>
              ))}
            </div>
          </div>
        )}

        {/* TAB 3: Printable Summary */}
        {activeTab === 'printable_summary' && (
          <div className="bg-white border border-gray-200 rounded-xl p-8 shadow-sm print:border-none print:shadow-none print:p-0">
            <div className="border-b-2 border-gray-300 pb-4 mb-6">
              <div className="flex justify-between items-start">
                <div>
                  <h2 className="text-xl font-bold text-gray-900">{companyName}</h2>
                  <h3 className="text-base font-semibold text-gray-900">MONTH-END CLOSING SIGN-OFF MEMO</h3>
                  <p className="text-xs text-gray-600">Accounting Period: {accountingPeriod} ({financialYear})</p>
                </div>
                <div className="text-right text-xs text-gray-600">
                  <div>Date: {new Date().toLocaleDateString('en-IN')}</div>
                  <div>Status: {metrics.completionPercentage === 100 ? 'CLOSED' : 'PENDING'}</div>
                </div>
              </div>
            </div>

            <div className="grid grid-cols-3 gap-4 mb-6 text-xs">
              <div className="border border-gray-200 rounded p-3">
                <span className="text-gray-600">Total Checklist Items:</span>
                <span className="font-bold ml-2">{metrics.totalTasks}</span>
              </div>
              <div className="border border-gray-200 rounded p-3">
                <span className="text-gray-600">Completed Activities:</span>
                <span className="font-bold ml-2 text-gray-900">{metrics.completedCount}</span>
              </div>
              <div className="border border-gray-200 rounded p-3">
                <span className="text-gray-600">Overall Completion:</span>
                <span className="font-bold ml-2 text-gray-900">{metrics.completionPercentage.toFixed(1)}%</span>
              </div>
            </div>

            <table className="w-full text-xs text-left mb-6 border-collapse">
              <thead>
                <tr className="bg-gray-900 text-white">
                  <th className="p-2 border border-gray-300">Category</th>
                  <th className="p-2 border border-gray-300">Activity</th>
                  <th className="p-2 border border-gray-300">Assignee</th>
                  <th className="p-2 border border-gray-300">Status</th>
                  <th className="p-2 border border-gray-300">Sign-Off Notes</th>
                </tr>
              </thead>
              <tbody>
                {tasks.map((t) => (
                  <tr key={t.id}>
                    <td className="p-2 border border-gray-300 font-semibold">{t.category}</td>
                    <td className="p-2 border border-gray-300 font-bold">{t.title}</td>
                    <td className="p-2 border border-gray-300">{t.assignee}</td>
                    <td className="p-2 border border-gray-300 font-semibold text-[10px]">
                      {t.status.toUpperCase()}
                    </td>
                    <td className="p-2 border border-gray-300 text-[10px] text-gray-600">{t.notes || '-'}</td>
                  </tr>
                ))}
              </tbody>
            </table>

            <div className="mt-12 pt-6 border-t border-gray-200 flex justify-between text-xs text-gray-600">
              <div>Prepared By: Accounts Team Leader</div>
              <div>Reviewed By: Senior Statutory Auditor</div>
              <div>Approved By: Partner / CFO</div>
            </div>
          </div>
        )}
          </>
        )}
      </div>

      {/* Add Custom Task Modal */}
      {showAddModal && (
        <div className="fixed inset-0 bg-black/40 flex items-center justify-center z-50 p-4">
          <div className="bg-white rounded-xl max-w-md w-full p-6 shadow-xl space-y-4">
            <div className="flex justify-between items-center border-b pb-3">
              <h3 className="font-bold text-sm text-gray-900">Add Month-End Closing Activity</h3>
              <button onClick={() => setShowAddModal(false)} className="text-gray-400 hover:text-gray-600">
                <X className="w-5 h-5" />
              </button>
            </div>

            <div className="space-y-3 text-xs">
              <div>
                <label className="block font-semibold text-gray-600 mb-1">Activity Title *</label>
                <input
                  type="text"
                  value={newTaskTitle}
                  onChange={(e) => setNewTaskTitle(e.target.value)}
                  placeholder="e.g. Foreign Exchange Revaluation"
                  className="w-full border rounded p-2 outline-none focus:border-gray-300"
                />
              </div>

              <div>
                <label className="block font-semibold text-gray-600 mb-1">Category</label>
                <select
                  value={newTaskCategory}
                  onChange={(e) => setNewTaskCategory(e.target.value as any)}
                  className="w-full border rounded p-2 outline-none focus:border-gray-300"
                >
                  <option value="Cash & Bank">Cash & Bank</option>
                  <option value="Receivables & Payables">Receivables & Payables</option>
                  <option value="General Ledger & Journals">General Ledger & Journals</option>
                  <option value="Fixed Assets">Fixed Assets</option>
                  <option value="Accruals & Cut-off">Accruals & Cut-off</option>
                  <option value="Taxation & Statutory">Taxation & Statutory</option>
                  <option value="Financial Reporting">Financial Reporting</option>
                </select>
              </div>

              <div>
                <label className="block font-semibold text-gray-600 mb-1">Assignee</label>
                <input
                  type="text"
                  value={newTaskAssignee}
                  onChange={(e) => setNewTaskAssignee(e.target.value)}
                  placeholder="e.g. Audit Senior"
                  className="w-full border rounded p-2 outline-none focus:border-gray-300"
                />
              </div>

              <div>
                <label className="block font-semibold text-gray-600 mb-1">Due Date</label>
                <input
                  type="text"
                  value={newTaskDueDate}
                  onChange={(e) => setNewTaskDueDate(e.target.value)}
                  className="w-full border rounded p-2 outline-none focus:border-gray-300"
                />
              </div>

              <div>
                <label className="block font-semibold text-gray-600 mb-1">Description / Standard Procedure</label>
                <textarea
                  value={newTaskDesc}
                  onChange={(e) => setNewTaskDesc(e.target.value)}
                  rows={2}
                  placeholder="Detail the audit steps to perform..."
                  className="w-full border rounded p-2 outline-none focus:border-gray-300"
                />
              </div>
            </div>

            <div className="flex justify-end gap-2 pt-2 border-t">
              <button
                onClick={() => setShowAddModal(false)}
                className="px-3 py-1.5 text-xs rounded border border-gray-300 text-gray-700 hover:bg-gray-50"
              >
                Cancel
              </button>
              <button
                onClick={handleAddTask}
                className="px-4 py-1.5 text-xs rounded bg-gray-900 text-white font-bold hover:bg-black"
              >
                Add Activity
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};
