import React, { useState, useEffect } from 'react';
import {
  Building2,
  Plus,
  Search,
  Check,
  X,
  Edit2,
  Trash2,
  ShieldCheck,
  AlertCircle,
  Briefcase,
  Calendar,
  CreditCard,
  FileText,
} from 'lucide-react';
import {
  type ClientProfile,
  type EntityType,
  getClients,
  getActiveClient,
  setActiveClient,
  addClient,
  updateClient,
  deleteClient,
  derivePanFromGstin,
  isValidGstin,
  isValidPan,
} from '../../utils/clientStorage';

interface ClientVaultModalProps {
  isOpen: boolean;
  onClose: () => void;
  onShowToast: (title: string, desc?: string, type?: 'success' | 'error' | 'info') => void;
}

const ENTITY_TYPES: EntityType[] = [
  'Private Limited',
  'Public Limited',
  'LLP',
  'Partnership Firm',
  'Sole Proprietorship',
  'Individual / HUF',
  'Trust / Society',
];

export const ClientVaultModal: React.FC<ClientVaultModalProps> = ({
  isOpen,
  onClose,
  onShowToast,
}) => {
  const [clients, setClients] = useState<ClientProfile[]>([]);
  const [activeClient, setActiveClientState] = useState<ClientProfile | null>(null);
  const [searchQuery, setSearchQuery] = useState('');
  const [viewMode, setViewMode] = useState<'list' | 'add' | 'edit'>('list');
  const [editingClientId, setEditingClientId] = useState<string | null>(null);

  // Form State
  const [formData, setFormData] = useState({
    name: '',
    tradeName: '',
    gstin: '',
    pan: '',
    entityType: 'Private Limited' as EntityType,
    financialYear: 'FY 2024-25',
    assessmentYear: 'AY 2025-26',
    contactPerson: '',
    email: '',
    phone: '',
    city: '',
  });

  const [formError, setFormError] = useState('');

  const reloadClients = () => {
    const list = getClients();
    setClients(list);
    setActiveClientState(getActiveClient());
  };

  useEffect(() => {
    if (isOpen) {
      reloadClients();
      setViewMode('list');
      setSearchQuery('');
    }
  }, [isOpen]);

  if (!isOpen) return null;

  // Filtered clients list
  const filteredClients = clients.filter(
    (c) =>
      c.name.toLowerCase().includes(searchQuery.toLowerCase()) ||
      c.gstin.toLowerCase().includes(searchQuery.toLowerCase()) ||
      c.pan.toLowerCase().includes(searchQuery.toLowerCase()) ||
      (c.tradeName && c.tradeName.toLowerCase().includes(searchQuery.toLowerCase()))
  );

  // Auto-fill PAN when GSTIN is typed
  const handleGstinChange = (val: string) => {
    const cleanGstin = val.toUpperCase().trim();
    const derivedPan = derivePanFromGstin(cleanGstin);
    setFormData((prev) => ({
      ...prev,
      gstin: cleanGstin,
      pan: derivedPan || prev.pan,
    }));
  };

  const handleSelectClient = (client: ClientProfile) => {
    setActiveClient(client.id);
    setActiveClientState(client);
    onShowToast('Active Client Switched', `Now managing: ${client.name}`, 'success');
    onClose();
  };

  const handleOpenAdd = () => {
    setFormData({
      name: '',
      tradeName: '',
      gstin: '',
      pan: '',
      entityType: 'Private Limited',
      financialYear: 'FY 2024-25',
      assessmentYear: 'AY 2025-26',
      contactPerson: '',
      email: '',
      phone: '',
      city: '',
    });
    setFormError('');
    setViewMode('add');
  };

  const handleOpenEdit = (client: ClientProfile) => {
    setEditingClientId(client.id);
    setFormData({
      name: client.name,
      tradeName: client.tradeName || '',
      gstin: client.gstin,
      pan: client.pan,
      entityType: client.entityType,
      financialYear: client.financialYear,
      assessmentYear: client.assessmentYear,
      contactPerson: client.contactPerson || '',
      email: client.email || '',
      phone: client.phone || '',
      city: client.city || '',
    });
    setFormError('');
    setViewMode('edit');
  };

  const handleDelete = (client: ClientProfile) => {
    if (clients.length <= 1) {
      onShowToast('Action Blocked', 'At least one client must remain in your practice vault.', 'error');
      return;
    }
    if (window.confirm(`Are you sure you want to delete client "${client.name}"?`)) {
      deleteClient(client.id);
      reloadClients();
      onShowToast('Client Deleted', `${client.name} removed from vault.`, 'info');
    }
  };

  const handleSaveForm = (e: React.FormEvent) => {
    e.preventDefault();
    setFormError('');

    if (!formData.name.trim()) {
      setFormError('Client legal name is required.');
      return;
    }

    if (formData.gstin && !isValidGstin(formData.gstin)) {
      setFormError('Invalid 15-character GSTIN format (e.g. 27AABCS1429B1Z1).');
      return;
    }

    if (formData.pan && !isValidPan(formData.pan)) {
      setFormError('Invalid 10-character PAN format (e.g. AABCS1429B).');
      return;
    }

    if (viewMode === 'add') {
      const created = addClient(formData);
      reloadClients();
      onShowToast('Client Created', `${created.name} added to practice vault.`, 'success');
      setViewMode('list');
    } else if (viewMode === 'edit' && editingClientId) {
      updateClient(editingClientId, formData);
      reloadClients();
      onShowToast('Client Updated', `${formData.name} details saved.`, 'success');
      setViewMode('list');
    }
  };

  return (
    <div className="fixed inset-0 z-50 bg-black/40 backdrop-blur-xs flex items-center justify-center p-4 animate-apple-fade select-none">
      <div className="bg-white rounded-2xl border border-gray-200 w-full max-w-2xl shadow-xl overflow-hidden flex flex-col max-h-[85vh] animate-apple-scale">
        {/* Modal Header */}
        <div className="h-14 px-6 border-b border-gray-100 flex items-center justify-between shrink-0 bg-white">
          <div className="flex items-center gap-2.5">
            <div className="w-8 h-8 rounded-lg bg-blue-50 text-blue-600 flex items-center justify-center">
              <Briefcase className="w-4.5 h-4.5" />
            </div>
            <div>
              <h3 className="text-sm font-bold text-gray-900">Practice Client Vault</h3>
              <p className="text-[11px] text-gray-500">Switch active client or register new practice entity</p>
            </div>
          </div>
          <button
            onClick={onClose}
            className="w-8 h-8 rounded-lg flex items-center justify-center text-gray-400 hover:text-gray-700 hover:bg-gray-100 transition-colors cursor-pointer"
          >
            <X className="w-4 h-4" />
          </button>
        </div>

        {/* Modal Body */}
        {viewMode === 'list' ? (
          <div className="flex-1 overflow-y-auto p-6 space-y-4">
            {/* Search and Add Toolbar */}
            <div className="flex items-center gap-3">
              <div className="relative flex-1">
                <Search className="w-4 h-4 text-gray-400 absolute left-3 top-1/2 -translate-y-1/2 pointer-events-none" />
                <input
                  type="text"
                  placeholder="Search clients by name, GSTIN or PAN..."
                  value={searchQuery}
                  onChange={(e) => setSearchQuery(e.target.value)}
                  className="w-full pl-9 pr-3.5 py-2 bg-gray-50 hover:bg-gray-100/80 focus:bg-white border border-gray-200 rounded-xl text-xs text-gray-900 placeholder-gray-400 focus:outline-none focus:border-gray-900 transition-all shadow-2xs"
                />
              </div>
              <button
                onClick={handleOpenAdd}
                className="inline-flex items-center gap-1.5 px-3.5 py-2 bg-gray-900 hover:bg-black text-white rounded-xl text-xs font-semibold shadow-2xs transition-all active:scale-95 cursor-pointer shrink-0"
              >
                <Plus className="w-4 h-4" />
                <span>Add Client</span>
              </button>
            </div>

            {/* Client Cards Grid */}
            <div className="space-y-2.5">
              {filteredClients.length === 0 ? (
                <div className="text-center py-12 bg-gray-50/70 rounded-xl border border-dashed border-gray-200 space-y-2">
                  <Building2 className="w-8 h-8 text-gray-400 mx-auto mb-1" />
                  <p className="text-xs font-semibold text-gray-700">No client profiles found</p>
                  <p className="text-[11px] text-gray-400 max-w-sm mx-auto">
                    Click "+ Add Client" to create your first auditee master record.
                  </p>
                  <button
                    onClick={handleOpenAdd}
                    className="mt-2 inline-flex items-center gap-1.5 px-3 py-1.5 bg-gray-900 hover:bg-black text-white text-xs font-semibold rounded-lg transition-all cursor-pointer shadow-2xs"
                  >
                    <Plus className="w-3.5 h-3.5" />
                    <span>Add First Client</span>
                  </button>
                </div>
              ) : (
                filteredClients.map((client) => {
                  const isActive = activeClient?.id === client.id;
                  return (
                    <div
                      key={client.id}
                      onClick={() => handleSelectClient(client)}
                      className={`p-4 rounded-xl border transition-all cursor-pointer group flex items-start justify-between gap-3 ${
                        isActive
                          ? 'bg-gray-50 border-gray-900 shadow-2xs'
                          : 'bg-white border-gray-200 hover:border-gray-300 hover:bg-gray-50/60'
                      }`}
                    >
                      <div className="space-y-1.5 min-w-0">
                        <div className="flex items-center gap-2">
                          <h4 className="text-xs font-bold text-gray-900 truncate">
                            {client.name}
                          </h4>
                          {isActive && (
                            <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded text-[10px] font-bold bg-gray-900 text-white shadow-2xs">
                              <Check className="w-3 h-3 stroke-[3]" />
                              Active Client
                            </span>
                          )}
                          <span className="px-2 py-0.5 rounded text-[10px] font-medium bg-gray-100 text-gray-600 border border-gray-200">
                            {client.entityType}
                          </span>
                        </div>

                        {client.tradeName && (
                          <div className="text-[11px] text-gray-500 font-medium truncate">
                            Trade Name: {client.tradeName}
                          </div>
                        )}

                        <div className="flex flex-wrap items-center gap-x-4 gap-y-1 text-[11px] text-gray-500 font-mono">
                          {client.gstin && (
                            <span>
                              GSTIN: <strong className="text-gray-700">{client.gstin}</strong>
                            </span>
                          )}
                          {client.pan && (
                            <span>
                              PAN: <strong className="text-gray-700">{client.pan}</strong>
                            </span>
                          )}
                          <span className="font-sans text-gray-400">·</span>
                          <span className="font-sans text-gray-600">{client.financialYear}</span>
                        </div>
                      </div>

                      {/* Client Action Buttons */}
                      <div
                        className="flex items-center gap-1 shrink-0"
                        onClick={(e) => e.stopPropagation()}
                      >
                        <button
                          onClick={() => handleOpenEdit(client)}
                          className="p-1.5 rounded-lg text-gray-400 hover:text-gray-700 hover:bg-white border border-transparent hover:border-gray-200 transition-colors cursor-pointer"
                          title="Edit Client Profile"
                        >
                          <Edit2 className="w-3.5 h-3.5" />
                        </button>
                        <button
                          onClick={() => handleDelete(client)}
                          className="p-1.5 rounded-lg text-gray-400 hover:text-rose-600 hover:bg-rose-50 border border-transparent hover:border-rose-200 transition-colors cursor-pointer"
                          title="Delete Client"
                        >
                          <Trash2 className="w-3.5 h-3.5" />
                        </button>
                      </div>
                    </div>
                  );
                })
              )}
            </div>
          </div>
        ) : (
          /* Add / Edit Form */
          <form onSubmit={handleSaveForm} className="flex-1 overflow-y-auto p-6 space-y-4">
            <div className="flex items-center justify-between pb-3 border-b border-gray-100">
              <h4 className="text-xs font-bold text-gray-900">
                {viewMode === 'add' ? 'Register New Practice Client' : 'Edit Client Profile'}
              </h4>
              <button
                type="button"
                onClick={() => setViewMode('list')}
                className="text-xs text-gray-500 hover:text-gray-900 cursor-pointer font-medium"
              >
                ← Back to Client List
              </button>
            </div>

            {formError && (
              <div className="flex items-center gap-2 p-2.5 rounded-xl bg-rose-50 border border-rose-200 text-xs text-rose-700">
                <AlertCircle className="w-4 h-4 shrink-0" />
                <span>{formError}</span>
              </div>
            )}

            <div className="grid grid-cols-1 sm:grid-cols-2 gap-3.5 text-xs">
              <div className="sm:col-span-2">
                <label className="font-semibold text-gray-700 block mb-1">Client Legal Name *</label>
                <input
                  type="text"
                  required
                  value={formData.name}
                  onChange={(e) => setFormData({ ...formData, name: e.target.value })}
                  placeholder="e.g. Reliance Infrastructure Limited"
                  className="w-full px-3 py-1.5 border border-gray-200 rounded-xl text-gray-900 placeholder-gray-400 focus:outline-none focus:border-gray-900 shadow-2xs"
                />
              </div>

              <div>
                <label className="font-semibold text-gray-700 block mb-1">Trade Name (Optional)</label>
                <input
                  type="text"
                  value={formData.tradeName}
                  onChange={(e) => setFormData({ ...formData, tradeName: e.target.value })}
                  placeholder="e.g. Reliance Infra"
                  className="w-full px-3 py-1.5 border border-gray-200 rounded-xl text-gray-900 placeholder-gray-400 focus:outline-none focus:border-gray-900 shadow-2xs"
                />
              </div>

              <div>
                <label className="font-semibold text-gray-700 block mb-1">Entity Constitution</label>
                <select
                  value={formData.entityType}
                  onChange={(e) => setFormData({ ...formData, entityType: e.target.value as EntityType })}
                  className="w-full px-3 py-1.5 border border-gray-200 rounded-xl text-gray-900 focus:outline-none focus:border-gray-900 shadow-2xs"
                >
                  {ENTITY_TYPES.map((t) => (
                    <option key={t} value={t}>
                      {t}
                    </option>
                  ))}
                </select>
              </div>

              <div>
                <label className="font-semibold text-gray-700 block mb-1">GSTIN (15 Digits)</label>
                <input
                  type="text"
                  maxLength={15}
                  value={formData.gstin}
                  onChange={(e) => handleGstinChange(e.target.value)}
                  placeholder="27AABCS1429B1Z1"
                  className="w-full px-3 py-1.5 font-mono border border-gray-200 rounded-xl text-gray-900 placeholder-gray-400 focus:outline-none focus:border-gray-900 shadow-2xs"
                />
              </div>

              <div>
                <label className="font-semibold text-gray-700 block mb-1">PAN (Auto-Derived from GSTIN)</label>
                <input
                  type="text"
                  maxLength={10}
                  value={formData.pan}
                  onChange={(e) => setFormData({ ...formData, pan: e.target.value.toUpperCase().trim() })}
                  placeholder="AABCS1429B"
                  className="w-full px-3 py-1.5 font-mono border border-gray-200 rounded-xl text-gray-900 placeholder-gray-400 focus:outline-none focus:border-gray-900 shadow-2xs"
                />
              </div>

              <div>
                <label className="font-semibold text-gray-700 block mb-1">Active Financial Year</label>
                <select
                  value={formData.financialYear}
                  onChange={(e) => {
                    const fy = e.target.value;
                    const match = fy.match(/FY\s*(\d{4})-(\d{2})/i);
                    const ay = match
                      ? `AY ${parseInt(match[1]) + 1}-${String((parseInt(match[2]) + 1) % 100).padStart(2, '0')}`
                      : 'AY 2026-27';
                    setFormData({
                      ...formData,
                      financialYear: fy,
                      assessmentYear: ay,
                    });
                  }}
                  className="w-full px-3 py-1.5 border border-gray-200 rounded-xl text-gray-900 focus:outline-none focus:border-gray-900 shadow-2xs"
                >
                  <option value="FY 2030-31">FY 2030-31 (AY 2031-32)</option>
                  <option value="FY 2029-30">FY 2029-30 (AY 2030-31)</option>
                  <option value="FY 2028-29">FY 2028-29 (AY 2029-30)</option>
                  <option value="FY 2027-28">FY 2027-28 (AY 2028-29)</option>
                  <option value="FY 2026-27">FY 2026-27 (AY 2027-28)</option>
                  <option value="FY 2025-26">FY 2025-26 (AY 2026-27)</option>
                  <option value="FY 2024-25">FY 2024-25 (AY 2025-26)</option>
                  <option value="FY 2023-24">FY 2023-24 (AY 2024-25)</option>
                  <option value="FY 2022-23">FY 2022-23 (AY 2023-24)</option>
                </select>
              </div>

              <div>
                <label className="font-semibold text-gray-700 block mb-1">City / State</label>
                <input
                  type="text"
                  value={formData.city}
                  onChange={(e) => setFormData({ ...formData, city: e.target.value })}
                  placeholder="e.g. Mumbai, Maharashtra"
                  className="w-full px-3 py-1.5 border border-gray-200 rounded-xl text-gray-900 placeholder-gray-400 focus:outline-none focus:border-gray-900 shadow-2xs"
                />
              </div>

              <div>
                <label className="font-semibold text-gray-700 block mb-1">Accounts Contact Email</label>
                <input
                  type="email"
                  value={formData.email}
                  onChange={(e) => setFormData({ ...formData, email: e.target.value })}
                  placeholder="accounts@client.com"
                  className="w-full px-3 py-1.5 border border-gray-200 rounded-xl text-gray-900 placeholder-gray-400 focus:outline-none focus:border-gray-900 shadow-2xs"
                />
              </div>

              <div>
                <label className="font-semibold text-gray-700 block mb-1">Contact Phone</label>
                <input
                  type="text"
                  value={formData.phone}
                  onChange={(e) => setFormData({ ...formData, phone: e.target.value })}
                  placeholder="+91 98200 12345"
                  className="w-full px-3 py-1.5 border border-gray-200 rounded-xl text-gray-900 placeholder-gray-400 focus:outline-none focus:border-gray-900 shadow-2xs"
                />
              </div>
            </div>

            <div className="pt-3 border-t border-gray-100 flex items-center justify-end gap-2.5">
              <button
                type="button"
                onClick={() => setViewMode('list')}
                className="px-3.5 py-1.5 text-xs font-medium text-gray-600 hover:text-gray-900 bg-gray-100 hover:bg-gray-200 rounded-xl cursor-pointer"
              >
                Cancel
              </button>
              <button
                type="submit"
                className="px-4 py-1.5 text-xs font-semibold text-white bg-blue-600 hover:bg-blue-700 rounded-xl shadow-2xs transition-all active:scale-95 cursor-pointer"
              >
                {viewMode === 'add' ? 'Save Client' : 'Update Client'}
              </button>
            </div>
          </form>
        )}
      </div>
    </div>
  );
};
