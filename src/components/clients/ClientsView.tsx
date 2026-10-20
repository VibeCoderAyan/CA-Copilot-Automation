import React, { useState } from 'react';
import type { Client } from '../../types';
import { Badge } from '../common/Badge';
import {
  Building2,
  Plus,
  Search,
  Filter,
  CheckCircle2,
  AlertCircle,
  ExternalLink,
  Mail,
  Phone,
  UserCheck,
  X
} from 'lucide-react';

interface ClientsViewProps {
  clients: Client[];
  onAddClient: (newClient: Client) => void;
  onSelectClient: (clientId: string) => void;
  onShowToast: (title: string, desc?: string, type?: 'success' | 'error' | 'info') => void;
}

export const ClientsView: React.FC<ClientsViewProps> = ({
  clients,
  onAddClient,
  onSelectClient,
  onShowToast,
}) => {
  const [searchQuery, setSearchQuery] = useState('');
  const [filterType, setFilterType] = useState('All');
  const [isModalOpen, setIsModalOpen] = useState(false);

  // New Client Form State
  const [formData, setFormData] = useState({
    name: '',
    tradeName: '',
    pan: '',
    gstin: '',
    entityType: 'Private Limited' as Client['entityType'],
    category: '',
    turnover: '',
    contactPerson: '',
    email: '',
    phone: '',
    assignedTo: 'Audit Partner',
  });

  const filteredClients = clients.filter((c) => {
    const matchesSearch =
      c.name.toLowerCase().includes(searchQuery.toLowerCase()) ||
      c.gstin.toLowerCase().includes(searchQuery.toLowerCase()) ||
      c.pan.toLowerCase().includes(searchQuery.toLowerCase());

    const matchesType = filterType === 'All' || c.entityType === filterType;
    return matchesSearch && matchesType;
  });

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    if (!formData.name || !formData.pan || !formData.gstin) {
      onShowToast('Validation Error', 'Entity Name, PAN, and GSTIN are required.', 'error');
      return;
    }

    const newClient: Client = {
      id: formData.name.toLowerCase().replace(/[^a-z0-9]/g, '-').slice(0, 16),
      ...formData,
      status: 'Active',
      pendingReviews: 0,
    };

    onAddClient(newClient);
    setIsModalOpen(false);
    setFormData({
      name: '',
      tradeName: '',
      pan: '',
      gstin: '',
      entityType: 'Private Limited',
      category: '',
      turnover: '',
      contactPerson: '',
      email: '',
      phone: '',
      assignedTo: 'Audit Partner',
    });
    onShowToast('Client Enrolled', `${newClient.name} added to practice master.`, 'success');
  };

  return (
    <div className="space-y-6">
      {/* Header Card */}
      <div className="bg-white rounded border border-gray-200 p-5">
        <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4">
          <div>
            <div className="flex items-center gap-2">
              <Building2 className="w-4 h-4 text-gray-900" />
              <h3 className="text-sm font-semibold text-[#111827]">Client Master Directory</h3>
              <Badge variant="navy">{clients.length} Entities Enrolled</Badge>
            </div>
            <p className="text-xs text-gray-600 mt-1">
              Centralized corporate master with GSTIN profiles, entity structures, and audit team allocations
            </p>
          </div>

          <button
            onClick={() => setIsModalOpen(true)}
            className="flex items-center gap-1.5 px-3.5 py-1.5 bg-gray-900 hover:bg-black text-white rounded text-xs font-medium transition-colors shadow-sm self-start sm:self-auto"
          >
            <Plus className="w-3.5 h-3.5" />
            <span>Add New Client</span>
          </button>
        </div>

        {/* Filter and Search Bar */}
        <div className="mt-4 pt-4 border-t border-gray-200 flex flex-col md:flex-row md:items-center justify-between gap-3">
          <div className="flex items-center gap-2 overflow-x-auto pb-1 md:pb-0">
            {['All', 'Private Limited', 'Public Limited', 'LLP', 'Partnership'].map((type) => (
              <button
                key={type}
                onClick={() => setFilterType(type)}
                className={`px-3 py-1.5 text-xs font-medium rounded border transition-colors shrink-0 ${
                  filterType === type
                    ? 'bg-gray-100 text-gray-900 border-gray-200'
                    : 'bg-white text-gray-600 border-gray-200 hover:bg-gray-50'
                }`}
              >
                {type}
              </button>
            ))}
          </div>

          <div className="relative w-full md:w-64">
            <Search className="w-3.5 h-3.5 text-gray-600 absolute left-3 top-1/2 -translate-y-1/2" />
            <input
              type="text"
              value={searchQuery}
              onChange={(e) => setSearchQuery(e.target.value)}
              placeholder="Search by name, GSTIN, PAN..."
              className="w-full pl-8 pr-3 py-1.5 bg-gray-50 border border-gray-200 rounded text-xs text-[#111827] placeholder-[#4B5563] focus:outline-none focus:bg-white focus:border-gray-300"
            />
          </div>
        </div>
      </div>

      {/* Clients Cards Grid */}
      <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
        {filteredClients.map((client) => {
          const statusVariants: Record<string, 'success' | 'warning' | 'error' | 'navy'> = {
            Active: 'success',
            'Audit in Progress': 'navy',
            'Return Due': 'error',
            Completed: 'success',
          };

          return (
            <div
              key={client.id}
              className="bg-white rounded border border-gray-200 p-4 flex flex-col justify-between hover:border-gray-300 transition-colors"
            >
              <div>
                <div className="flex items-start justify-between gap-2 mb-2">
                  <div>
                    <h4 className="text-xs font-bold text-[#111827] leading-snug">{client.name}</h4>
                    {client.tradeName && (
                      <p className="text-[11px] text-gray-600">t/a {client.tradeName}</p>
                    )}
                  </div>
                  <Badge variant={statusVariants[client.status] || 'neutral'}>
                    {client.status}
                  </Badge>
                </div>

                <div className="space-y-1.5 text-xs mt-3 pt-3 border-t border-gray-200">
                  <div className="flex justify-between">
                    <span className="text-gray-600">Entity Type:</span>
                    <span className="font-medium text-[#111827]">{client.entityType}</span>
                  </div>
                  <div className="flex justify-between">
                    <span className="text-gray-600">PAN:</span>
                    <span className="font-mono text-[#111827] font-medium">{client.pan}</span>
                  </div>
                  <div className="flex justify-between">
                    <span className="text-gray-600">GSTIN:</span>
                    <span className="font-mono text-[#111827] font-medium">{client.gstin}</span>
                  </div>
                  <div className="flex justify-between">
                    <span className="text-gray-600">Annual Turnover:</span>
                    <span className="font-semibold text-gray-900">{client.turnover}</span>
                  </div>
                  <div className="flex justify-between">
                    <span className="text-gray-600">Assigned CA:</span>
                    <span className="text-[#111827] truncate max-w-[150px]">{client.assignedTo}</span>
                  </div>
                </div>
              </div>

              <div className="pt-3 mt-4 border-t border-gray-200 flex items-center justify-between">
                <span className="text-[11px] text-gray-600">
                  Pending Tasks:{' '}
                  <strong className={client.pendingReviews > 0 ? 'text-[#C43D43]' : 'text-gray-900'}>
                    {client.pendingReviews}
                  </strong>
                </span>

                <button
                  onClick={() => {
                    onSelectClient(client.id);
                    onShowToast('Client Workspace Activated', `Switched active scope to ${client.name}`, 'info');
                  }}
                  className="px-2.5 py-1 text-xs font-medium text-gray-900 hover:text-gray-900 hover:bg-gray-100 rounded border border-gray-200 transition-colors"
                >
                  Open Client Workspace
                </button>
              </div>
            </div>
          );
        })}
      </div>

      {/* Add Client Modal */}
      {isModalOpen && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-[#111827]/40 p-4">
          <div className="bg-white rounded border border-gray-200 shadow-modal w-full max-w-lg overflow-hidden animate-in fade-in duration-150">
            <div className="px-5 py-4 border-b border-gray-200 flex items-center justify-between">
              <h3 className="text-sm font-semibold text-[#111827]">Enroll New Client Entity</h3>
              <button onClick={() => setIsModalOpen(false)} className="text-gray-600 hover:text-[#111827]">
                <X className="w-4 h-4" />
              </button>
            </div>

            <form onSubmit={handleSubmit} className="p-5 space-y-3.5 text-xs">
              <div>
                <label className="font-medium text-[#111827] block mb-1">Company / Entity Legal Name *</label>
                <input
                  type="text"
                  required
                  placeholder="e.g. Zenith Tech Solutions Pvt Ltd"
                  value={formData.name}
                  onChange={(e) => setFormData({ ...formData, name: e.target.value })}
                  className="w-full px-3 py-1.5 border border-gray-200 rounded focus:border-gray-300 focus:outline-none"
                />
              </div>

              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="font-medium text-[#111827] block mb-1">Entity Type</label>
                  <select
                    value={formData.entityType}
                    onChange={(e) => setFormData({ ...formData, entityType: e.target.value as any })}
                    className="w-full px-3 py-1.5 border border-gray-200 rounded focus:border-gray-300 focus:outline-none"
                  >
                    <option value="Private Limited">Private Limited</option>
                    <option value="Public Limited">Public Limited</option>
                    <option value="LLP">LLP</option>
                    <option value="Partnership">Partnership</option>
                    <option value="Proprietorship">Proprietorship</option>
                  </select>
                </div>

                <div>
                  <label className="font-medium text-[#111827] block mb-1">Industry Category</label>
                  <input
                    type="text"
                    placeholder="e.g. IT & Software Services"
                    value={formData.category}
                    onChange={(e) => setFormData({ ...formData, category: e.target.value })}
                    className="w-full px-3 py-1.5 border border-gray-200 rounded focus:border-gray-300 focus:outline-none"
                  />
                </div>
              </div>

              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="font-medium text-[#111827] block mb-1">Permanent Account No (PAN) *</label>
                  <input
                    type="text"
                    required
                    maxLength={10}
                    placeholder="e.g. AABCA1234F"
                    value={formData.pan}
                    onChange={(e) => setFormData({ ...formData, pan: e.target.value.toUpperCase() })}
                    className="w-full px-3 py-1.5 font-mono border border-gray-200 rounded focus:border-gray-300 focus:outline-none"
                  />
                </div>

                <div>
                  <label className="font-medium text-[#111827] block mb-1">GSTIN *</label>
                  <input
                    type="text"
                    required
                    maxLength={15}
                    placeholder="e.g. 27AABCA1234F1Z5"
                    value={formData.gstin}
                    onChange={(e) => setFormData({ ...formData, gstin: e.target.value.toUpperCase() })}
                    className="w-full px-3 py-1.5 font-mono border border-gray-200 rounded focus:border-gray-300 focus:outline-none"
                  />
                </div>
              </div>

              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="font-medium text-[#111827] block mb-1">Turnover Estimate</label>
                  <input
                    type="text"
                    placeholder="e.g. ₹25.00 Cr"
                    value={formData.turnover}
                    onChange={(e) => setFormData({ ...formData, turnover: e.target.value })}
                    className="w-full px-3 py-1.5 border border-gray-200 rounded focus:border-gray-300 focus:outline-none"
                  />
                </div>

                <div>
                  <label className="font-medium text-[#111827] block mb-1">Lead Audit Partner</label>
                  <select
                    value={formData.assignedTo}
                    onChange={(e) => setFormData({ ...formData, assignedTo: e.target.value })}
                    className="w-full px-3 py-1.5 border border-gray-200 rounded focus:border-gray-300 focus:outline-none"
                  >
                    <option value="Audit Partner">Audit Partner</option>
                    <option value="Senior Auditor">Senior Auditor</option>
                    <option value="Article Assistant">Article Assistant</option>
                  </select>
                </div>
              </div>

              <div className="pt-3 border-t border-gray-200 flex justify-end gap-2">
                <button
                  type="button"
                  onClick={() => setIsModalOpen(false)}
                  className="px-3 py-1.5 text-xs text-gray-600 border border-gray-200 rounded hover:bg-gray-50"
                >
                  Cancel
                </button>
                <button
                  type="submit"
                  className="px-4 py-1.5 text-xs font-medium bg-gray-900 hover:bg-black text-white rounded shadow-sm"
                >
                  Enroll Client
                </button>
              </div>
            </form>
          </div>
        </div>
      )}
    </div>
  );
};
