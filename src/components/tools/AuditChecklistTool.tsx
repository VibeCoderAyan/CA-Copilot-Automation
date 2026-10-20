import React, { useState } from 'react';
import * as XLSX from 'xlsx';
import {
  ArrowLeft,
  Download,
  ClipboardCheck,
  Plus,
  Printer,
  CheckCircle2,
  Clock,
  MinusCircle,
  FileSpreadsheet,
  Building2,
  Calendar,
  User,
  Filter,
  Trash2,
  Edit3
} from 'lucide-react';
import {
  type ClientAuditProject,
  type AuditChecklistItem,
  type AuditType,
  AUDIT_TEMPLATES,
  getInitialClientProjects,
  calculateChecklistMetrics,
  exportAuditChecklistWorkbook,
  createDefaultChecklistItems,
} from '../../utils/auditChecklistEngine';
import { DataFormatGuide } from '../common/DataFormatGuide';
import { ToolHeader } from '../common/ToolHeader';

interface AuditChecklistToolProps {
  onBack: () => void;
  onShowToast: (title: string, description?: string, type?: 'success' | 'error' | 'info') => void;
}

export const AuditChecklistTool: React.FC<AuditChecklistToolProps> = ({ onBack, onShowToast }) => {
  const [projects, setProjects] = useState<ClientAuditProject[]>(getInitialClientProjects());
  const [activeProjectId, setActiveProjectId] = useState<string>(projects[0]?.id || '');
  const [statusFilter, setStatusFilter] = useState<'all' | 'pending' | 'received' | 'na'>('all');
  const [searchTerm, setSearchTerm] = useState('');
  const [showNewProjectModal, setShowNewProjectModal] = useState(false);
  const [showAddItemModal, setShowAddItemModal] = useState(false);

  // New Project Form State
  const [newClientName, setNewClientName] = useState('');
  const [newFY, setNewFY] = useState('FY 2026-27');
  const [newAuditType, setNewAuditType] = useState<AuditType>('statutory_audit');
  const [newPartner, setNewPartner] = useState('');
  const [newSenior, setNewSenior] = useState('');
  const [newTargetDate, setNewTargetDate] = useState('2026-10-31');

  // New Item Form State
  const [newItemParticulars, setNewItemParticulars] = useState('');
  const [newItemCategory, setNewItemCategory] = useState('General Compliances');
  const [newItemRequirement, setNewItemRequirement] = useState('');
  const [newItemPriority, setNewItemPriority] = useState<'mandatory' | 'recommended' | 'optional'>('mandatory');

  const currentProject = projects.find((p) => p.id === activeProjectId) || projects[0];

  const metrics = currentProject ? calculateChecklistMetrics(currentProject.items) : { total: 0, received: 0, pending: 0, na: 0, percentage: 0 };

  const handleUpdateItemStatus = (itemId: string, newStatus: 'received' | 'pending' | 'na') => {
    setProjects((prev) =>
      prev.map((proj) => {
        if (proj.id !== activeProjectId) return proj;
        const updatedItems = proj.items.map((item) => {
          if (item.id === itemId) {
            return {
              ...item,
              status: newStatus,
              receivedDate: newStatus === 'received' ? new Date().toISOString().split('T')[0] : undefined,
            };
          }
          return item;
        });
        return { ...proj, items: updatedItems, updatedAt: new Date().toISOString().split('T')[0] };
      })
    );
  };

  const handleUpdateItemNotes = (itemId: string, notes: string) => {
    setProjects((prev) =>
      prev.map((proj) => {
        if (proj.id !== activeProjectId) return proj;
        const updatedItems = proj.items.map((item) => (item.id === itemId ? { ...item, notes } : item));
        return { ...proj, items: updatedItems };
      })
    );
  };

  const handleUpdateItemRef = (itemId: string, documentRef: string) => {
    setProjects((prev) =>
      prev.map((proj) => {
        if (proj.id !== activeProjectId) return proj;
        const updatedItems = proj.items.map((item) => (item.id === itemId ? { ...item, documentRef } : item));
        return { ...proj, items: updatedItems };
      })
    );
  };

  const handleDeleteItem = (itemId: string) => {
    setProjects((prev) =>
      prev.map((proj) => {
        if (proj.id !== activeProjectId) return proj;
        return { ...proj, items: proj.items.filter((i) => i.id !== itemId) };
      })
    );
    onShowToast('Item Removed', 'Checklist item deleted.', 'info');
  };

  const handleCreateProject = (e: React.FormEvent) => {
    e.preventDefault();
    if (!newClientName.trim()) {
      onShowToast('Validation Error', 'Please enter a client legal entity name.', 'error');
      return;
    }

    const newProject: ClientAuditProject = {
      id: `proj-${Date.now()}-${Math.random().toString(36).substr(2, 4)}`,
      clientName: newClientName.trim(),
      financialYear: newFY,
      auditType: newAuditType,
      engagementPartner: newPartner,
      leadSenior: newSenior,
      targetCompletionDate: newTargetDate,
      items: createDefaultChecklistItems(newAuditType),
      updatedAt: new Date().toISOString().split('T')[0],
    };

    setProjects([newProject, ...projects]);
    setActiveProjectId(newProject.id);
    setShowNewProjectModal(false);
    setNewClientName('');
    onShowToast('Audit Engagement Created', `Initialized checklist for ${newProject.clientName}.`, 'success');
  };

  const handleAddItem = (e: React.FormEvent) => {
    e.preventDefault();
    if (!newItemParticulars.trim()) return;

    const newItem: AuditChecklistItem = {
      id: `item-custom-${Date.now()}`,
      code: `REQ-${currentProject.items.length + 1}`,
      category: newItemCategory,
      particulars: newItemParticulars.trim(),
      requirement: newItemRequirement.trim() || 'Provide supporting documentation and reconciliation.',
      priority: newItemPriority,
      status: 'pending',
    };

    setProjects((prev) =>
      prev.map((proj) => {
        if (proj.id !== activeProjectId) return proj;
        return { ...proj, items: [...proj.items, newItem] };
      })
    );

    setShowAddItemModal(false);
    setNewItemParticulars('');
    setNewItemRequirement('');
    onShowToast('Item Added', 'New checklist item appended.', 'success');
  };

  const handleExportExcel = () => {
    if (!currentProject) return;
    try {
      const wb = exportAuditChecklistWorkbook(currentProject);
      XLSX.writeFile(wb, `Audit_Checklist_${currentProject.clientName.replace(/\s+/g, '_')}_${currentProject.financialYear}.xlsx`);
      onShowToast('Workbook Exported', 'Downloaded 3-sheet audit checklist and client requisition report.', 'success');
    } catch (err: any) {
      onShowToast('Export Failed', err.message || 'Could not export Excel file.', 'error');
    }
  };

  const handlePrintRequisition = () => {
    window.print();
  };

  const filteredItems = currentProject
    ? currentProject.items.filter((item) => {
        const matchesFilter =
          statusFilter === 'all' || item.status === statusFilter;
        const matchesSearch =
          item.particulars.toLowerCase().includes(searchTerm.toLowerCase()) ||
          item.category.toLowerCase().includes(searchTerm.toLowerCase()) ||
          item.code.toLowerCase().includes(searchTerm.toLowerCase()) ||
          item.requirement.toLowerCase().includes(searchTerm.toLowerCase());
        return matchesFilter && matchesSearch;
      })
    : [];

  return (
    <div className="space-y-6">
      <ToolHeader
        toolId="audit-checklist"
        title="Statutory Audit Checklist & Engagement"
        description="Standardized statutory, tax, and GST audit checklists with client progress tracking and pending requisition exports."
        badge="Multi-Client & Statutory"
        onBack={onBack}
        actionButtons={
          projects.length > 0 ? (
            <>
              <button
                onClick={handlePrintRequisition}
                className="flex items-center gap-1.5 px-3 py-2 bg-white border border-gray-200 hover:border-gray-300 text-gray-900 rounded-lg text-xs font-semibold shadow-sm transition-colors cursor-pointer"
                title="Print formal client requisition letter"
              >
                <Printer className="w-3.5 h-3.5 text-gray-900" />
                Print Requisition
              </button>
              <button
                onClick={handleExportExcel}
                className="flex items-center gap-1.5 px-4 py-2 bg-gray-900 hover:bg-black text-white rounded-lg text-xs font-semibold shadow-sm transition-colors cursor-pointer"
              >
                <Download className="w-3.5 h-3.5" />
                Export Audit Papers (Excel)
              </button>
            </>
          ) : (
            <button
              onClick={() => setShowNewProjectModal(true)}
              className="flex items-center gap-1.5 px-3.5 py-1.5 bg-gray-900 hover:bg-black text-white rounded-lg text-xs font-semibold shadow-sm transition-colors cursor-pointer"
            >
              <Plus className="w-3.5 h-3.5" />
              <span>New Client Engagement</span>
            </button>
          )
        }
      />

      <DataFormatGuide
        toolType="audit"
        acceptedFormats="Excel (.xlsx, .xls), CSV (.csv)"
        requiredFields={[
          { name: 'Particulars', description: 'Audit checklist item or verification requirement', aliases: ['particulars', 'item', 'description'] },
          { name: 'Category', description: 'Audit area or section', aliases: ['category', 'area'] },
        ]}
      />

      {projects.length === 0 ? (
        <div className="bg-white rounded-xl border border-dashed border-gray-300 p-12 text-center space-y-4 shadow-xs">
          <div className="w-12 h-12 rounded-full bg-gray-50 text-gray-700 flex items-center justify-center mx-auto">
            <ClipboardCheck className="w-6 h-6" />
          </div>
          <div className="max-w-md mx-auto">
            <h3 className="text-sm font-semibold text-gray-900">
              No Audit Engagements Created Yet
            </h3>
            <p className="text-xs text-gray-500 mt-1">
              Start an engagement for Statutory Audit (Companies Act), Tax Audit (Form 3CD), GST Audit (9C), Internal Financial Controls (IFC), or Bank Concurrent Audit.
            </p>
          </div>

          <div className="flex items-center justify-center gap-3 pt-2">
            <button
              onClick={() => setShowNewProjectModal(true)}
              className="flex items-center gap-1.5 px-4 py-2 bg-gray-900 hover:bg-black text-white text-xs font-semibold rounded-lg shadow-sm transition-colors cursor-pointer"
            >
              <Plus className="w-3.5 h-3.5" />
              <span>Create New Audit Engagement</span>
            </button>
          </div>
        </div>
      ) : (
        <>

      {/* Client Engagement Header Card */}
      <div className="bg-white border border-gray-200 rounded-xl p-5 shadow-sm space-y-4">
        <div className="flex flex-col md:flex-row md:items-center justify-between gap-4 pb-4 border-b border-gray-200">
          {/* Client Selector */}
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 rounded-lg bg-gray-100 border border-gray-200 flex items-center justify-center text-gray-900 shrink-0">
              <Building2 className="w-5 h-5" />
            </div>
            <div>
              <div className="flex items-center gap-2">
                <select
                  value={activeProjectId}
                  onChange={(e) => setActiveProjectId(e.target.value)}
                  className="text-sm font-bold text-gray-900 bg-white border border-gray-200 rounded-lg px-2.5 py-1 focus:outline-none focus:border-gray-300"
                >
                  {projects.map((p) => (
                    <option key={p.id} value={p.id}>
                      {p.clientName} ({p.financialYear})
                    </option>
                  ))}
                </select>
                <button
                  onClick={() => setShowNewProjectModal(true)}
                  className="flex items-center gap-1 px-2.5 py-1 bg-gray-100 hover:bg-gray-200 border border-gray-200 text-gray-900 text-xs font-semibold rounded-lg transition-colors"
                >
                  <Plus className="w-3 h-3" />
                  New Client
                </button>
              </div>
              <p className="text-xs text-gray-600 mt-0.5">
                {AUDIT_TEMPLATES[currentProject?.auditType]?.name} | Partner: {currentProject?.engagementPartner}
              </p>
            </div>
          </div>

          {/* Quick Metrics */}
          <div className="flex items-center gap-4">
            <div className="text-right">
              <div className="text-xs font-semibold text-gray-600">Audit Completion</div>
              <div className="text-xl font-bold text-gray-900">{metrics.percentage}%</div>
            </div>
            {/* Visual Progress Meter */}
            <div className="w-32 bg-[#E5E7EB] h-2.5 rounded-full overflow-hidden">
              <div
                className="bg-gray-900 h-full transition-all duration-300"
                style={{ width: `${metrics.percentage}%` }}
              />
            </div>
          </div>
        </div>

        {/* 4 Metrics Cards */}
        <div className="grid grid-cols-2 sm:grid-cols-4 gap-3">
          <div className="p-3 bg-gray-50 border border-gray-200 rounded-lg">
            <span className="text-[11px] font-semibold text-gray-600 block">Total Documents</span>
            <span className="text-lg font-bold text-gray-900 mt-0.5 block">{metrics.total}</span>
          </div>
          <div className="p-3 bg-gray-100 border border-gray-200 rounded-lg">
            <span className="text-[11px] font-semibold text-gray-900 flex items-center gap-1">
              <CheckCircle2 className="w-3 h-3" /> Received & Verified
            </span>
            <span className="text-lg font-bold text-gray-900 mt-0.5 block">{metrics.received}</span>
          </div>
          <div className="p-3 bg-amber-50 border border-amber-200 rounded-lg">
            <span className="text-[11px] font-semibold text-amber-800 flex items-center gap-1">
              <Clock className="w-3 h-3" /> Pending Client Action
            </span>
            <span className="text-lg font-bold text-amber-800 mt-0.5 block">{metrics.pending}</span>
          </div>
          <div className="p-3 bg-slate-50 border border-slate-200 rounded-lg">
            <span className="text-[11px] font-semibold text-slate-600 flex items-center gap-1">
              <MinusCircle className="w-3 h-3" /> Not Applicable
            </span>
            <span className="text-lg font-bold text-slate-700 mt-0.5 block">{metrics.na}</span>
          </div>
        </div>
      </div>

      {/* Filter Bar & Search */}
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 bg-white border border-gray-200 rounded-xl p-3 shadow-sm">
        <div className="flex items-center gap-1.5">
          <button
            onClick={() => setStatusFilter('all')}
            className={`px-3 py-1.5 rounded-lg text-xs font-semibold transition-colors ${
              statusFilter === 'all'
                ? 'bg-gray-900 text-white shadow-sm'
                : 'bg-transparent text-gray-600 hover:bg-gray-50'
            }`}
          >
            All Items ({metrics.total})
          </button>
          <button
            onClick={() => setStatusFilter('pending')}
            className={`px-3 py-1.5 rounded-lg text-xs font-semibold transition-colors ${
              statusFilter === 'pending'
                ? 'bg-amber-600 text-white shadow-sm'
                : 'bg-transparent text-amber-800 hover:bg-amber-50'
            }`}
          >
            Pending ({metrics.pending})
          </button>
          <button
            onClick={() => setStatusFilter('received')}
            className={`px-3 py-1.5 rounded-lg text-xs font-semibold transition-colors ${
              statusFilter === 'received'
                ? 'bg-gray-900 text-white shadow-sm'
                : 'bg-transparent text-gray-900 hover:bg-gray-100'
            }`}
          >
            Received ({metrics.received})
          </button>
          <button
            onClick={() => setStatusFilter('na')}
            className={`px-3 py-1.5 rounded-lg text-xs font-semibold transition-colors ${
              statusFilter === 'na'
                ? 'bg-slate-600 text-white shadow-sm'
                : 'bg-transparent text-slate-600 hover:bg-slate-100'
            }`}
          >
            N/A ({metrics.na})
          </button>
        </div>

        <div className="flex items-center gap-2">
          <input
            type="text"
            placeholder="Search documents or requirements..."
            value={searchTerm}
            onChange={(e) => setSearchTerm(e.target.value)}
            className="w-full sm:w-64 px-3 py-1.5 bg-white border border-gray-200 rounded-lg text-xs text-gray-900 focus:outline-none focus:border-gray-300"
          />
          <button
            onClick={() => setShowAddItemModal(true)}
            className="flex items-center gap-1 px-3 py-1.5 bg-gray-900 hover:bg-black text-white text-xs font-semibold rounded-lg shadow-sm shrink-0 transition-colors"
          >
            <Plus className="w-3.5 h-3.5" />
            Add Item
          </button>
        </div>
      </div>

      {/* Checklist Table */}
      <div className="bg-white border border-gray-200 rounded-xl overflow-hidden shadow-sm">
        <div className="overflow-x-auto">
          <table className="w-full text-left text-xs border-collapse">
            <thead className="bg-gray-50 border-b border-gray-200 text-gray-600 uppercase font-bold text-[10px] tracking-wider">
              <tr>
                <th className="py-2.5 px-3 w-12 text-center">Ref</th>
                <th className="py-2.5 px-4">Audit Area & Description</th>
                <th className="py-2.5 px-3 w-28">Priority</th>
                <th className="py-2.5 px-3 w-48 text-center">Audit Status</th>
                <th className="py-2.5 px-3 w-32">Working Paper</th>
                <th className="py-2.5 px-4">Notes & Observations</th>
                <th className="py-2.5 px-2 w-10 text-center"></th>
              </tr>
            </thead>
            <tbody className="divide-y divide-[#E5E7EB]">
              {filteredItems.map((item) => (
                <tr key={item.id} className="hover:bg-gray-50 transition-colors">
                  <td className="py-3 px-3 text-center font-mono text-[11px] text-gray-600">
                    {item.code}
                  </td>
                  <td className="py-3 px-4">
                    <div className="font-semibold text-xs text-gray-900">{item.particulars}</div>
                    <div className="text-[11px] text-gray-600 mt-0.5">{item.requirement}</div>
                    <span className="inline-block mt-1 text-[10px] bg-slate-100 text-slate-700 px-1.5 py-0.5 rounded">
                      {item.category}
                    </span>
                  </td>
                  <td className="py-3 px-3">
                    {item.priority === 'mandatory' ? (
                      <span className="text-[10px] font-bold uppercase tracking-wider bg-red-50 text-red-700 px-2 py-0.5 rounded border border-red-200">
                        Mandatory
                      </span>
                    ) : item.priority === 'recommended' ? (
                      <span className="text-[10px] font-semibold uppercase tracking-wider bg-amber-50 text-amber-800 px-2 py-0.5 rounded border border-amber-200">
                        Recommended
                      </span>
                    ) : (
                      <span className="text-[10px] font-medium uppercase tracking-wider bg-slate-100 text-slate-600 px-2 py-0.5 rounded">
                        Optional
                      </span>
                    )}
                  </td>
                  <td className="py-3 px-3 text-center">
                    <div className="inline-flex rounded-lg border border-gray-200 p-0.5 bg-gray-50">
                      <button
                        onClick={() => handleUpdateItemStatus(item.id, 'received')}
                        className={`px-2 py-1 text-[11px] font-semibold rounded transition-colors ${
                          item.status === 'received'
                            ? 'bg-gray-900 text-white shadow-xs'
                            : 'text-gray-600 hover:text-gray-900'
                        }`}
                        title="Mark as Received"
                      >
                        Received
                      </button>
                      <button
                        onClick={() => handleUpdateItemStatus(item.id, 'pending')}
                        className={`px-2 py-1 text-[11px] font-semibold rounded transition-colors ${
                          item.status === 'pending'
                            ? 'bg-amber-600 text-white shadow-xs'
                            : 'text-gray-600 hover:text-gray-900'
                        }`}
                        title="Mark as Pending"
                      >
                        Pending
                      </button>
                      <button
                        onClick={() => handleUpdateItemStatus(item.id, 'na')}
                        className={`px-2 py-1 text-[11px] font-semibold rounded transition-colors ${
                          item.status === 'na'
                            ? 'bg-slate-600 text-white shadow-xs'
                            : 'text-gray-600 hover:text-gray-900'
                        }`}
                        title="Mark as Not Applicable"
                      >
                        N/A
                      </button>
                    </div>
                  </td>
                  <td className="py-3 px-3">
                    <input
                      type="text"
                      placeholder="e.g. WP-A1"
                      value={item.documentRef || ''}
                      onChange={(e) => handleUpdateItemRef(item.id, e.target.value)}
                      className="w-full px-2 py-1 bg-white border border-gray-200 rounded text-xs text-gray-900 focus:outline-none focus:border-gray-300"
                    />
                  </td>
                  <td className="py-3 px-4">
                    <input
                      type="text"
                      placeholder="Add auditor remarks or client note..."
                      value={item.notes || ''}
                      onChange={(e) => handleUpdateItemNotes(item.id, e.target.value)}
                      className="w-full px-2 py-1 bg-white border border-gray-200 rounded text-xs text-gray-900 focus:outline-none focus:border-gray-300"
                    />
                  </td>
                  <td className="py-3 px-2 text-center">
                    <button
                      onClick={() => handleDeleteItem(item.id)}
                      className="p-1 text-gray-600 hover:text-[#DC2626] transition-colors"
                      title="Delete item"
                    >
                      <Trash2 className="w-3.5 h-3.5" />
                    </button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
      </>
      )}

      {/* Modal: Create Client Audit Engagement */}
      {showNewProjectModal && (
        <div className="fixed inset-0 bg-black/40 backdrop-blur-xs flex items-center justify-center z-50 p-4">
          <div className="bg-white border border-gray-200 rounded-xl shadow-xl w-full max-w-lg p-6 space-y-4">
            <h3 className="text-sm font-bold text-gray-900 uppercase tracking-wider">
              Initialize New Client Engagement
            </h3>
            <form onSubmit={handleCreateProject} className="space-y-3 text-xs">
              <div>
                <label className="font-semibold text-gray-900 block mb-1">Client Legal Name</label>
                <input
                  type="text"
                  required
                  placeholder="Enter Client Name..."
                  value={newClientName}
                  onChange={(e) => setNewClientName(e.target.value)}
                  className="w-full px-3 py-2 bg-white border border-gray-200 rounded-lg text-xs text-gray-900 focus:outline-none focus:border-gray-300"
                />
              </div>

              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="font-semibold text-gray-900 block mb-1">Financial Year</label>
                  <select
                    value={newFY}
                    onChange={(e) => setNewFY(e.target.value)}
                    className="w-full px-3 py-2 bg-white border border-gray-200 rounded-lg text-xs text-gray-900 focus:outline-none focus:border-gray-300"
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
                <div>
                  <label className="font-semibold text-gray-900 block mb-1">Engagement Type</label>
                  <select
                    value={newAuditType}
                    onChange={(e) => setNewAuditType(e.target.value as AuditType)}
                    className="w-full px-3 py-2 bg-white border border-gray-200 rounded-lg text-xs text-gray-900 focus:outline-none focus:border-gray-300"
                  >
                    <option value="statutory_audit">Statutory Audit (Companies Act)</option>
                    <option value="tax_audit_3cd">Tax Audit (Form 3CD)</option>
                    <option value="gst_audit_9c">GST Annual Audit (GSTR-9C)</option>
                    <option value="ifc_review">Internal Financial Controls (IFC)</option>
                    <option value="bank_concurrent">Bank Concurrent Audit</option>
                  </select>
                </div>
              </div>

              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="font-semibold text-gray-900 block mb-1">Engagement Partner</label>
                  <input
                    type="text"
                    value={newPartner}
                    onChange={(e) => setNewPartner(e.target.value)}
                    className="w-full px-3 py-2 bg-white border border-gray-200 rounded-lg text-xs text-gray-900 focus:outline-none focus:border-gray-300"
                  />
                </div>
                <div>
                  <label className="font-semibold text-gray-900 block mb-1">Lead Audit Senior</label>
                  <input
                    type="text"
                    value={newSenior}
                    onChange={(e) => setNewSenior(e.target.value)}
                    className="w-full px-3 py-2 bg-white border border-gray-200 rounded-lg text-xs text-gray-900 focus:outline-none focus:border-gray-300"
                  />
                </div>
              </div>

              <div>
                <label className="font-semibold text-gray-900 block mb-1">Target Sign-Off Date</label>
                <input
                  type="date"
                  value={newTargetDate}
                  onChange={(e) => setNewTargetDate(e.target.value)}
                  className="w-full px-3 py-2 bg-white border border-gray-200 rounded-lg text-xs text-gray-900 focus:outline-none focus:border-gray-300"
                />
              </div>

              <div className="flex justify-end gap-2 pt-3 border-t border-gray-200">
                <button
                  type="button"
                  onClick={() => setShowNewProjectModal(false)}
                  className="px-3 py-1.5 border border-gray-200 rounded-lg text-gray-600 hover:bg-gray-50"
                >
                  Cancel
                </button>
                <button
                  type="submit"
                  className="px-4 py-1.5 bg-gray-900 hover:bg-black text-white font-semibold rounded-lg shadow-sm"
                >
                  Create Engagement
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* Modal: Add Checklist Item */}
      {showAddItemModal && (
        <div className="fixed inset-0 bg-black/40 backdrop-blur-xs flex items-center justify-center z-50 p-4">
          <div className="bg-white border border-gray-200 rounded-xl shadow-xl w-full max-w-md p-6 space-y-4">
            <h3 className="text-sm font-bold text-gray-900 uppercase tracking-wider">
              Add Audit Checklist Requirement
            </h3>
            <form onSubmit={handleAddItem} className="space-y-3 text-xs">
              <div>
                <label className="font-semibold text-gray-900 block mb-1">Document / Requirement Name</label>
                <input
                  type="text"
                  required
                  placeholder="e.g. Foreign Inward Remittance Certificates (FIRC)"
                  value={newItemParticulars}
                  onChange={(e) => setNewItemParticulars(e.target.value)}
                  className="w-full px-3 py-2 bg-white border border-gray-200 rounded-lg text-xs text-gray-900 focus:outline-none focus:border-gray-300"
                />
              </div>
              <div>
                <label className="font-semibold text-gray-900 block mb-1">Category / Audit Head</label>
                <input
                  type="text"
                  value={newItemCategory}
                  onChange={(e) => setNewItemCategory(e.target.value)}
                  className="w-full px-3 py-2 bg-white border border-gray-200 rounded-lg text-xs text-gray-900 focus:outline-none focus:border-gray-300"
                />
              </div>
              <div>
                <label className="font-semibold text-gray-900 block mb-1">Audit Requirement Details</label>
                <textarea
                  rows={2}
                  placeholder="Specific instructions or sampling parameters..."
                  value={newItemRequirement}
                  onChange={(e) => setNewItemRequirement(e.target.value)}
                  className="w-full px-3 py-2 bg-white border border-gray-200 rounded-lg text-xs text-gray-900 focus:outline-none focus:border-gray-300"
                />
              </div>
              <div>
                <label className="font-semibold text-gray-900 block mb-1">Priority</label>
                <select
                  value={newItemPriority}
                  onChange={(e) => setNewItemPriority(e.target.value as any)}
                  className="w-full px-3 py-2 bg-white border border-gray-200 rounded-lg text-xs text-gray-900 focus:outline-none focus:border-gray-300"
                >
                  <option value="mandatory">Mandatory</option>
                  <option value="recommended">Recommended</option>
                  <option value="optional">Optional</option>
                </select>
              </div>

              <div className="flex justify-end gap-2 pt-3 border-t border-gray-200">
                <button
                  type="button"
                  onClick={() => setShowAddItemModal(false)}
                  className="px-3 py-1.5 border border-gray-200 rounded-lg text-gray-600 hover:bg-gray-50"
                >
                  Cancel
                </button>
                <button
                  type="submit"
                  className="px-4 py-1.5 bg-gray-900 hover:bg-black text-white font-semibold rounded-lg shadow-sm"
                >
                  Add to Checklist
                </button>
              </div>
            </form>
          </div>
        </div>
      )}
    </div>
  );
};