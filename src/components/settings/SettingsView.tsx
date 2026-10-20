import React, { useState, useEffect } from 'react';
import {
  Lock,
  Building2,
  Sliders,
  Database,
  Save,
  CheckCircle2,
  KeyRound,
  ShieldCheck,
  Download,
  Upload,
  Trash2,
  AlertCircle,
  Eye,
  EyeOff,
  RefreshCw,
  FileText,
  Server,
} from 'lucide-react';
import {
  isPinSet,
  isPinEnabled,
  savePin,
  verifyPin,
  removePin,
  togglePinEnabled,
} from '../../utils/security';
import { SegmentedControl, type SegmentedTab } from '../common/SegmentedControl';
import { OfflineToolsView } from './OfflineToolsView';

interface SettingsViewProps {
  onShowToast: (title: string, desc?: string, type?: 'success' | 'error' | 'info') => void;
  onLockApp?: () => void;
  onOpenClientVault?: () => void;
}

type SettingsTab = 'security' | 'firm' | 'rules' | 'data' | 'offline';

export const SettingsView: React.FC<SettingsViewProps> = ({ onShowToast, onLockApp, onOpenClientVault }) => {
  const [activeTab, setActiveTab] = useState<SettingsTab>('security');

  // ==========================================
  // 1. PIN & SECURITY STATE
  // ==========================================
  const [pinEnabled, setPinEnabled] = useState(isPinEnabled());
  const [hasPin, setHasPin] = useState(isPinSet());

  // Set / Change PIN modal / form state
  const [showPinModal, setShowPinModal] = useState(false);
  const [modalMode, setModalMode] = useState<'create' | 'change' | 'remove'>('create');
  const [oldPin, setOldPin] = useState('');
  const [newPin, setNewPin] = useState('');
  const [confirmPin, setConfirmPin] = useState('');
  const [pinError, setPinError] = useState('');

  // ==========================================
  // 2. FIRM PROFILE STATE (Saved in localStorage)
  // ==========================================
  const [firmConfig, setFirmConfig] = useState(() => {
    const saved = localStorage.getItem('ca_copilot_firm_profile');
    if (saved) {
      try {
        return JSON.parse(saved);
      } catch (e) {
        // Fallback
      }
    }
    return {
      firmName: '',
      frn: '',
      icaiBranch: 'WIRC',
      managingPartner: '',
      partnerMrn: '',
      udinPrefix: '',
      officeEmail: '',
      officePhone: '',
      officeAddress: '',
    };
  });

  // ==========================================
  // 3. STATUTORY RULES STATE (Saved in localStorage)
  // ==========================================
  const [rulesConfig, setRulesConfig] = useState(() => {
    const saved = localStorage.getItem('ca_copilot_rules_config');
    if (saved) {
      try {
        return JSON.parse(saved);
      } catch (e) {
        // Fallback
      }
    }
    return {
      assessmentYear: 'AY 2025-26',
      financialYear: 'FY 2024-25',
      accountingFramework: 'AS',
      dateToleranceDays: 3,
      amountToleranceRs: 0,
      gstr2bRoundingTolerance: 1.0,
      materialityPercentage: 0.5,
      flagSec40A3: true,
      defaultDateFormat: 'DD-MM-YYYY',
    };
  });

  // Storage diagnosis
  const [storageUsedKb, setStorageUsedKb] = useState(0);

  useEffect(() => {
    let totalBytes = 0;
    for (let i = 0; i < localStorage.length; i++) {
      const key = localStorage.key(i);
      if (key) {
        totalBytes += (localStorage.getItem(key) || '').length * 2;
      }
    }
    setStorageUsedKb(Math.round(totalBytes / 1024));
  }, []);

  // Save Firm Profile
  const handleSaveFirm = (e: React.FormEvent) => {
    e.preventDefault();
    localStorage.setItem('ca_copilot_firm_profile', JSON.stringify(firmConfig));
    onShowToast('Firm Profile Saved', 'Credentials updated for reports & audit certificates.', 'success');
  };

  // Save Rules
  const handleSaveRules = (e: React.FormEvent) => {
    e.preventDefault();
    localStorage.setItem('ca_copilot_rules_config', JSON.stringify(rulesConfig));
    onShowToast('Statutory Rules Saved', 'Audit materiality & BRS tolerances updated.', 'success');
  };

  // Handle PIN Setup / Change
  const handlePinSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    setPinError('');

    if (modalMode === 'change' || modalMode === 'remove') {
      if (!verifyPin(oldPin)) {
        setPinError('Current PIN is incorrect.');
        return;
      }
    }

    if (modalMode === 'remove') {
      removePin(oldPin);
      setHasPin(false);
      setPinEnabled(false);
      setShowPinModal(false);
      setOldPin('');
      onShowToast('PIN Removed', 'App lock is now disabled.', 'info');
      return;
    }

    // Validation for create or change
    if (!/^\d{4}$/.test(newPin)) {
      setPinError('PIN must be exactly 4 numeric digits (0-9).');
      return;
    }

    if (newPin !== confirmPin) {
      setPinError('New PIN and confirmation PIN do not match.');
      return;
    }

    savePin(newPin);
    setHasPin(true);
    setPinEnabled(true);
    setShowPinModal(false);
    setOldPin('');
    setNewPin('');
    setConfirmPin('');
    onShowToast(
      modalMode === 'create' ? 'PIN Created Successfully' : 'PIN Changed Successfully',
      'Your practice workspace is now protected with 4-digit security.',
      'success'
    );
  };

  // Export Full Practice Backup
  const handleExportBackup = () => {
    const backupData = {
      version: '2.4',
      exportDate: new Date().toISOString(),
      firmConfig,
      rulesConfig,
      pinProtected: hasPin,
    };

    const blob = new Blob([JSON.stringify(backupData, null, 2)], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `ca_copilot_backup_${new Date().toISOString().split('T')[0]}.json`;
    a.click();
    URL.revokeObjectURL(url);
    onShowToast('Backup Downloaded', 'Firm settings & rules saved to local file.', 'success');
  };

  // Import Backup
  const handleImportBackup = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (event) => {
      try {
        const data = JSON.parse(event.target?.result as string);
        if (data.firmConfig) {
          setFirmConfig(data.firmConfig);
          localStorage.setItem('ca_copilot_firm_profile', JSON.stringify(data.firmConfig));
        }
        if (data.rulesConfig) {
          setRulesConfig(data.rulesConfig);
          localStorage.setItem('ca_copilot_rules_config', JSON.stringify(data.rulesConfig));
        }
        onShowToast('Backup Restored', 'Settings and firm profile loaded successfully.', 'success');
      } catch (err) {
        onShowToast('Import Failed', 'Invalid JSON backup file format.', 'error');
      }
    };
    reader.readAsText(file);
    e.target.value = '';
  };

  // Safe Cache Clear
  const handleClearCache = () => {
    if (window.confirm('Clear temporary conversion data & cached tables? Your firm profile, PIN, and rules will NOT be deleted.')) {
      // Retain critical keys
      const firm = localStorage.getItem('ca_copilot_firm_profile');
      const rules = localStorage.getItem('ca_copilot_rules_config');
      const pinHash = localStorage.getItem('ca_copilot_pin_hash');
      const pinEn = localStorage.getItem('ca_copilot_pin_enabled');

      localStorage.clear();

      if (firm) localStorage.setItem('ca_copilot_firm_profile', firm);
      if (rules) localStorage.setItem('ca_copilot_rules_config', rules);
      if (pinHash) localStorage.setItem('ca_copilot_pin_hash', pinHash);
      if (pinEn) localStorage.setItem('ca_copilot_pin_enabled', pinEn);

      onShowToast('Cache Cleared', 'Temporary tables purged. Settings preserved.', 'success');
    }
  };

  return (
    <div className="max-w-4xl space-y-6 pb-12 animate-apple-fade">
      {/* 1. Header Title & Description */}
      <div className="pb-4 border-b border-gray-200">
        <h1 className="text-2xl font-bold text-gray-900 tracking-tight">Practice Settings</h1>
        <p className="text-xs text-gray-500 mt-1">
          Firm profile, PIN lock, and audit configurations.
        </p>
      </div>

      {/* 2. Top Segmented Navigation Tabs */}
      <SegmentedControl
        tabs={[
          { id: 'security', label: 'PIN Lock', icon: Lock },
          { id: 'firm', label: 'Firm Profile', icon: Building2 },
          { id: 'rules', label: 'Audit Rules', icon: Sliders },
          { id: 'data', label: 'Backups', icon: Database },
          { id: 'offline', label: 'Offline Tools', icon: Server },
        ]}
        activeTab={activeTab}
        onChange={(tabId) => setActiveTab(tabId as SettingsTab)}
      />

      {/* ========================================================
          TAB 1: SECURITY & 4-DIGIT PIN
         ======================================================== */}
      {activeTab === 'security' && (
        <div className="space-y-6">
          {/* Main PIN Card */}
          <div className="bg-white rounded-2xl border border-gray-200 p-6 shadow-2xs space-y-5">
            <div className="flex items-start gap-3.5">
              <div className="w-10 h-10 rounded-xl bg-gray-100 border border-gray-200 flex items-center justify-center shrink-0">
                <Lock className="w-5 h-5 text-gray-700" />
              </div>
              <div>
                <h3 className="text-sm font-bold text-gray-900">4-Digit PIN Lock</h3>
                <p className="text-xs text-gray-500 mt-0.5 max-w-lg leading-relaxed">
                  Require a 4-digit PIN to open the application.
                </p>
              </div>
            </div>

            {/* Action Bar */}
            <div className="pt-4 border-t border-gray-100 flex flex-wrap items-center justify-between gap-3">
              <div className="flex items-center gap-2.5">
                {!hasPin ? (
                  <button
                    onClick={() => {
                      setModalMode('create');
                      setPinError('');
                      setOldPin('');
                      setNewPin('');
                      setConfirmPin('');
                      setShowPinModal(true);
                    }}
                    className="inline-flex items-center gap-1.5 px-4 py-2 bg-gray-900 hover:bg-black text-white rounded-xl text-xs font-semibold transition-all active:scale-95 shadow-2xs cursor-pointer"
                  >
                    <Lock className="w-3.5 h-3.5" />
                    <span>Create 4-Digit PIN</span>
                  </button>
                ) : (
                  <>
                    <button
                      onClick={() => {
                        setModalMode('change');
                        setPinError('');
                        setOldPin('');
                        setNewPin('');
                        setConfirmPin('');
                        setShowPinModal(true);
                      }}
                      className="inline-flex items-center gap-1.5 px-3.5 py-2 bg-gray-100 hover:bg-gray-200 text-gray-800 rounded-xl text-xs font-semibold transition-all active:scale-95 cursor-pointer"
                    >
                      <RefreshCw className="w-3.5 h-3.5 text-gray-600" />
                      <span>Change PIN</span>
                    </button>
                    <button
                      onClick={() => {
                        setModalMode('remove');
                        setPinError('');
                        setOldPin('');
                        setShowPinModal(true);
                      }}
                      className="inline-flex items-center gap-1.5 px-3.5 py-2 bg-rose-50 hover:bg-rose-100 text-rose-700 border border-rose-200 rounded-xl text-xs font-semibold transition-all active:scale-95 cursor-pointer"
                    >
                      <Trash2 className="w-3.5 h-3.5 text-rose-600" />
                      <span>Remove PIN</span>
                    </button>
                  </>
                )}
              </div>

              {/* Instant Lock Button */}
              {hasPin && (
                <button
                  onClick={() => {
                    if (onLockApp) {
                      onLockApp();
                    } else {
                      window.location.reload();
                    }
                  }}
                  className="inline-flex items-center gap-1.5 px-3.5 py-2 bg-gray-900 hover:bg-black text-white rounded-xl text-xs font-semibold transition-all active:scale-95 shadow-2xs cursor-pointer"
                  title="Lock workspace immediately"
                >
                  <Lock className="w-3.5 h-3.5 text-blue-400" />
                  <span>Lock App Now</span>
                </button>
              )}
            </div>
          </div>
        </div>
      )}


      {/* ========================================================
          TAB 2: FIRM & ICAI PROFILE
         ======================================================== */}
      {activeTab === 'firm' && (
        <form onSubmit={handleSaveFirm} className="bg-white rounded-2xl border border-gray-200 p-6 space-y-5 shadow-2xs">
          <div>
            <h3 className="text-sm font-bold text-gray-900">Chartered Accountant Practice Profile</h3>
            <p className="text-xs text-gray-500 mt-0.5">
              These credentials automatically populate official SA-230 Audit Workpapers, BRS Sign-off Sheets, and PDF Exports.
            </p>
          </div>

          <div className="grid grid-cols-1 sm:grid-cols-2 gap-4 text-xs">
            <div className="sm:col-span-2">
              <label className="font-semibold text-gray-700 block mb-1.5">CA Firm Legal Name</label>
              <input
                type="text"
                value={firmConfig.firmName}
                onChange={(e) => setFirmConfig({ ...firmConfig, firmName: e.target.value })}
                placeholder="e.g. K. Sharma & Associates, Chartered Accountants"
                className="w-full px-3.5 py-2 bg-white border border-gray-200 rounded-xl text-gray-900 placeholder-gray-400 focus:outline-none focus:border-gray-900 shadow-2xs transition-colors"
              />
            </div>

            <div>
              <label className="font-semibold text-gray-700 block mb-1.5">ICAI Firm Registration No (FRN)</label>
              <input
                type="text"
                value={firmConfig.frn}
                onChange={(e) => setFirmConfig({ ...firmConfig, frn: e.target.value })}
                placeholder="e.g. 120488N"
                className="w-full px-3.5 py-2 font-mono bg-white border border-gray-200 rounded-xl text-gray-900 placeholder-gray-400 focus:outline-none focus:border-gray-900 shadow-2xs transition-colors"
              />
            </div>

            <div>
              <label className="font-semibold text-gray-700 block mb-1.5">ICAI Regional Council</label>
              <select
                value={firmConfig.icaiBranch}
                onChange={(e) => setFirmConfig({ ...firmConfig, icaiBranch: e.target.value })}
                className="w-full px-3.5 py-2 bg-white border border-gray-200 rounded-xl text-gray-900 focus:outline-none focus:border-gray-900 shadow-2xs transition-colors"
              >
                <option value="WIRC">WIRC (Western India Regional Council)</option>
                <option value="NIRC">NIRC (Northern India Regional Council)</option>
                <option value="SIRC">SIRC (Southern India Regional Council)</option>
                <option value="EIRC">EIRC (Eastern India Regional Council)</option>
                <option value="CIRC">CIRC (Central India Regional Council)</option>
              </select>
            </div>

            <div>
              <label className="font-semibold text-gray-700 block mb-1.5">Signing Partner / Proprietor</label>
              <input
                type="text"
                value={firmConfig.managingPartner}
                onChange={(e) => setFirmConfig({ ...firmConfig, managingPartner: e.target.value })}
                placeholder="e.g. CA Rajesh Sharma, FCA"
                className="w-full px-3.5 py-2 bg-white border border-gray-200 rounded-xl text-gray-900 placeholder-gray-400 focus:outline-none focus:border-gray-900 shadow-2xs transition-colors"
              />
            </div>

            <div>
              <label className="font-semibold text-gray-700 block mb-1.5">Partner Membership No (MRN)</label>
              <input
                type="text"
                value={firmConfig.partnerMrn}
                onChange={(e) => setFirmConfig({ ...firmConfig, partnerMrn: e.target.value })}
                placeholder="e.g. 054321"
                className="w-full px-3.5 py-2 font-mono bg-white border border-gray-200 rounded-xl text-gray-900 placeholder-gray-400 focus:outline-none focus:border-gray-900 shadow-2xs transition-colors"
              />
            </div>

            <div>
              <label className="font-semibold text-gray-700 block mb-1.5">Official Office Email</label>
              <input
                type="email"
                value={firmConfig.officeEmail}
                onChange={(e) => setFirmConfig({ ...firmConfig, officeEmail: e.target.value })}
                placeholder="audit@cafirm.com"
                className="w-full px-3.5 py-2 bg-white border border-gray-200 rounded-xl text-gray-900 placeholder-gray-400 focus:outline-none focus:border-gray-900 shadow-2xs transition-colors"
              />
            </div>

            <div>
              <label className="font-semibold text-gray-700 block mb-1.5">Office Phone / Mobile</label>
              <input
                type="text"
                value={firmConfig.officePhone}
                onChange={(e) => setFirmConfig({ ...firmConfig, officePhone: e.target.value })}
                placeholder="+91 98765 43210"
                className="w-full px-3.5 py-2 bg-white border border-gray-200 rounded-xl text-gray-900 placeholder-gray-400 focus:outline-none focus:border-gray-900 shadow-2xs transition-colors"
              />
            </div>

            <div className="sm:col-span-2">
              <label className="font-semibold text-gray-700 block mb-1.5">Practice Office Address</label>
              <input
                type="text"
                value={firmConfig.officeAddress}
                onChange={(e) => setFirmConfig({ ...firmConfig, officeAddress: e.target.value })}
                placeholder="e.g. Suite 402, Commerce House, Nariman Point, Mumbai"
                className="w-full px-3.5 py-2 bg-white border border-gray-200 rounded-xl text-gray-900 placeholder-gray-400 focus:outline-none focus:border-gray-900 shadow-2xs transition-colors"
              />
            </div>
          </div>

          <div className="pt-4 border-t border-gray-100 flex items-center justify-between">
            {onOpenClientVault ? (
              <button
                type="button"
                onClick={onOpenClientVault}
                className="inline-flex items-center gap-1.5 px-3.5 py-2 bg-white border border-gray-300 hover:border-gray-900 text-gray-800 text-xs font-semibold rounded-xl transition-all cursor-pointer shadow-2xs"
              >
                <Building2 className="w-3.5 h-3.5 text-blue-600" />
                <span>Open Client Practice Vault</span>
              </button>
            ) : <div />}

            <button
              type="submit"
              className="inline-flex items-center gap-1.5 px-4 py-2 bg-gray-900 hover:bg-black text-white text-xs font-semibold rounded-xl transition-all active:scale-95 shadow-2xs cursor-pointer"
            >
              <Save className="w-3.5 h-3.5" />
              <span>Save Firm Credentials</span>
            </button>
          </div>
        </form>
      )}

      {/* ========================================================
          TAB 3: AUDIT & STATUTORY RULES
         ======================================================== */}
      {activeTab === 'rules' && (
        <form onSubmit={handleSaveRules} className="bg-white rounded-2xl border border-gray-200 p-6 space-y-5 shadow-2xs">
          <div>
            <h3 className="text-sm font-bold text-gray-900">Audit Scrutiny & Tolerance Settings</h3>
            <p className="text-xs text-gray-500 mt-0.5">
              Set standard tolerances for automatic reconciliation matching, materiality thresholds, and statutory tax flags.
            </p>
          </div>

          <div className="grid grid-cols-1 sm:grid-cols-2 gap-4 text-xs">
            <div>
              <label className="font-semibold text-gray-700 block mb-1.5">Assessment Year (AY)</label>
              <select
                value={rulesConfig.assessmentYear}
                onChange={(e) => setRulesConfig({ ...rulesConfig, assessmentYear: e.target.value })}
                className="w-full px-3.5 py-2 bg-white border border-gray-200 rounded-xl text-gray-900 focus:outline-none focus:border-gray-900 shadow-2xs transition-colors"
              >
                <option value="AY 2031-32">AY 2031-32 (FY 2030-31)</option>
                <option value="AY 2030-31">AY 2030-31 (FY 2029-30)</option>
                <option value="AY 2029-30">AY 2029-30 (FY 2028-29)</option>
                <option value="AY 2028-29">AY 2028-29 (FY 2027-28)</option>
                <option value="AY 2027-28">AY 2027-28 (FY 2026-27)</option>
                <option value="AY 2026-27">AY 2026-27 (FY 2025-26)</option>
                <option value="AY 2025-26">AY 2025-26 (FY 2024-25)</option>
                <option value="AY 2024-25">AY 2024-25 (FY 2023-24)</option>
                <option value="AY 2023-24">AY 2023-24 (FY 2022-23)</option>
                <option value="AY 2022-23">AY 2022-23 (FY 2021-22)</option>
              </select>
            </div>

            <div>
              <label className="font-semibold text-gray-700 block mb-1.5">Accounting Framework</label>
              <select
                value={rulesConfig.accountingFramework}
                onChange={(e) => setRulesConfig({ ...rulesConfig, accountingFramework: e.target.value })}
                className="w-full px-3.5 py-2 bg-white border border-gray-200 rounded-xl text-gray-900 focus:outline-none focus:border-gray-900 shadow-2xs transition-colors"
              >
                <option value="AS">Accounting Standards (Indian GAAP / Non-Ind AS)</option>
                <option value="Ind AS">Indian Accounting Standards (Ind AS / IFRS Aligned)</option>
              </select>
            </div>

            <div>
              <label className="font-semibold text-gray-700 block mb-1.5">
                BRS Cheque Clearance Tolerance (± Days)
              </label>
              <input
                type="number"
                min="0"
                max="30"
                value={rulesConfig.dateToleranceDays}
                onChange={(e) => setRulesConfig({ ...rulesConfig, dateToleranceDays: parseInt(e.target.value) || 0 })}
                className="w-full px-3.5 py-2 font-mono bg-white border border-gray-200 rounded-xl text-gray-900 focus:outline-none focus:border-gray-900 shadow-2xs transition-colors"
              />
              <span className="text-[11px] text-gray-400 mt-1 block">Allowable bank clearing lag days.</span>
            </div>

            <div>
              <label className="font-semibold text-gray-700 block mb-1.5">
                GSTR-2B vs Books Rounding Tolerance (₹)
              </label>
              <input
                type="number"
                step="0.01"
                min="0"
                max="10"
                value={rulesConfig.gstr2bRoundingTolerance}
                onChange={(e) => setRulesConfig({ ...rulesConfig, gstr2bRoundingTolerance: parseFloat(e.target.value) || 0 })}
                className="w-full px-3.5 py-2 font-mono bg-white border border-gray-200 rounded-xl text-gray-900 focus:outline-none focus:border-gray-900 shadow-2xs transition-colors"
              />
              <span className="text-[11px] text-gray-400 mt-1 block">Maximum allowable paisa rounding discrepancy.</span>
            </div>

            <div>
              <label className="font-semibold text-gray-700 block mb-1.5">
                Audit Materiality Threshold (%)
              </label>
              <input
                type="number"
                step="0.1"
                min="0.1"
                max="5"
                value={rulesConfig.materialityPercentage}
                onChange={(e) => setRulesConfig({ ...rulesConfig, materialityPercentage: parseFloat(e.target.value) || 0.5 })}
                className="w-full px-3.5 py-2 font-mono bg-white border border-gray-200 rounded-xl text-gray-900 focus:outline-none focus:border-gray-900 shadow-2xs transition-colors"
              />
              <span className="text-[11px] text-gray-400 mt-1 block">Benchmark percentage for SA-320 materiality checks.</span>
            </div>

            <div>
              <label className="font-semibold text-gray-700 block mb-1.5">Default Date Standard</label>
              <select
                value={rulesConfig.defaultDateFormat}
                onChange={(e) => setRulesConfig({ ...rulesConfig, defaultDateFormat: e.target.value })}
                className="w-full px-3.5 py-2 bg-white border border-gray-200 rounded-xl text-gray-900 focus:outline-none focus:border-gray-900 shadow-2xs transition-colors"
              >
                <option value="DD-MM-YYYY">DD-MM-YYYY (Indian Standard)</option>
                <option value="YYYY-MM-DD">YYYY-MM-DD (ISO 8601)</option>
              </select>
            </div>
          </div>

          <div className="pt-4 border-t border-gray-100 flex justify-end">
            <button
              type="submit"
              className="inline-flex items-center gap-1.5 px-4 py-2 bg-gray-900 hover:bg-black text-white text-xs font-semibold rounded-xl transition-all active:scale-95 shadow-2xs cursor-pointer"
            >
              <Save className="w-3.5 h-3.5" />
              <span>Save Statutory Defaults</span>
            </button>
          </div>
        </form>
      )}

      {/* ========================================================
          TAB 4: DATA & BACKUPS
         ======================================================== */}
      {activeTab === 'data' && (
        <div className="space-y-5">
          {/* Backup & Restore Card */}
          <div className="bg-white rounded-2xl border border-gray-200 p-6 space-y-5 shadow-2xs">
            <div>
              <h3 className="text-sm font-bold text-gray-900">Practice Backup & Restore</h3>
              <p className="text-xs text-gray-500 mt-0.5">
                Export a full JSON backup of your practice credentials, preferences, and offline rules.
              </p>
            </div>

            <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
              {/* Export Backup */}
              <div className="p-4 rounded-xl border border-gray-200 bg-gray-50/60 flex flex-col justify-between space-y-3">
                <div>
                  <h4 className="text-xs font-bold text-gray-900">Export Practice Configuration</h4>
                  <p className="text-[11px] text-gray-500 mt-1 leading-snug">
                    Save all firm details, tolerances, and configurations as a single portable JSON file.
                  </p>
                </div>
                <button
                  onClick={handleExportBackup}
                  className="inline-flex items-center justify-center gap-1.5 px-3.5 py-2 bg-white border border-gray-200 hover:border-gray-300 text-gray-800 rounded-xl text-xs font-semibold shadow-2xs transition-all active:scale-95 cursor-pointer"
                >
                  <Download className="w-3.5 h-3.5 text-blue-600" />
                  <span>Download Backup (.json)</span>
                </button>
              </div>

              {/* Import Backup */}
              <div className="p-4 rounded-xl border border-gray-200 bg-gray-50/60 flex flex-col justify-between space-y-3">
                <div>
                  <h4 className="text-xs font-bold text-gray-900">Restore Configuration</h4>
                  <p className="text-[11px] text-gray-500 mt-1 leading-snug">
                    Restore settings from a previously saved JSON backup file.
                  </p>
                </div>
                <label className="inline-flex items-center justify-center gap-1.5 px-3.5 py-2 bg-white border border-gray-200 hover:border-gray-300 text-gray-800 rounded-xl text-xs font-semibold shadow-2xs transition-all active:scale-95 cursor-pointer">
                  <Upload className="w-3.5 h-3.5 text-emerald-600" />
                  <span>Select Backup File</span>
                  <input
                    type="file"
                    accept=".json"
                    onChange={handleImportBackup}
                    className="hidden"
                  />
                </label>
              </div>
            </div>
          </div>

          {/* Local Storage Diagnostic & Purge */}
          <div className="bg-white rounded-2xl border border-gray-200 p-6 space-y-4 shadow-2xs">
            <div className="flex items-center justify-between">
              <div>
                <h3 className="text-sm font-bold text-gray-900">Local Cache Management</h3>
                <p className="text-xs text-gray-500 mt-0.5">
                  Approx. <span className="font-mono font-bold text-gray-800">{storageUsedKb} KB</span> occupied in browser local storage.
                </p>
              </div>

              <button
                onClick={handleClearCache}
                className="inline-flex items-center gap-1.5 px-3 py-1.5 bg-rose-50 hover:bg-rose-100 text-rose-700 border border-rose-200 rounded-xl text-xs font-semibold transition-all active:scale-95 cursor-pointer"
              >
                <Trash2 className="w-3.5 h-3.5" />
                <span>Clear Temp Cache</span>
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ========================================================
          TAB 5: OFFLINE PRACTICE TOOLS
         ======================================================== */}
      {activeTab === 'offline' && (
        <div className="space-y-6">
          <OfflineToolsView
            onShowToast={onShowToast}
            onOpenVault={onOpenClientVault}
          />
        </div>
      )}

      {/* ========================================================
          5. PIN SETUP / CHANGE / REMOVE MODAL
         ======================================================== */}
      {showPinModal && (
        <div className="fixed inset-0 z-50 bg-black/40 backdrop-blur-xs flex items-center justify-center p-4 animate-apple-fade">
          <div className="bg-white rounded-2xl border border-gray-200 p-6 w-full max-w-sm shadow-xl space-y-4 animate-apple-scale">
            <div className="flex items-center justify-between">
              <div className="flex items-center gap-2">
                <div className="w-7 h-7 rounded-lg bg-blue-50 text-blue-600 flex items-center justify-center">
                  <KeyRound className="w-4 h-4" />
                </div>
                <h3 className="text-sm font-bold text-gray-900">
                  {modalMode === 'create' && 'Create 4-Digit Security PIN'}
                  {modalMode === 'change' && 'Change Security PIN'}
                  {modalMode === 'remove' && 'Remove PIN Protection'}
                </h3>
              </div>
              <button
                onClick={() => setShowPinModal(false)}
                className="text-gray-400 hover:text-gray-700 text-sm font-bold cursor-pointer"
              >
                ✕
              </button>
            </div>

            <form onSubmit={handlePinSubmit} className="space-y-3.5">
              {/* Old PIN if changing or removing */}
              {(modalMode === 'change' || modalMode === 'remove') && (
                <div>
                  <label className="text-xs font-semibold text-gray-700 block mb-1">
                    Enter Current 4-Digit PIN
                  </label>
                  <input
                    type="password"
                    maxLength={4}
                    autoFocus
                    value={oldPin}
                    onChange={(e) => setOldPin(e.target.value.replace(/\D/g, ''))}
                    placeholder="••••"
                    className="w-full text-center text-lg tracking-widest px-3 py-2 border border-gray-200 rounded-xl font-mono focus:border-blue-600 focus:outline-none"
                  />
                </div>
              )}

              {/* New PIN & Confirm PIN if creating or changing */}
              {modalMode !== 'remove' && (
                <>
                  <div>
                    <label className="text-xs font-semibold text-gray-700 block mb-1">
                      New 4-Digit PIN (0-9)
                    </label>
                    <input
                      type="password"
                      maxLength={4}
                      autoFocus={modalMode === 'create'}
                      value={newPin}
                      onChange={(e) => setNewPin(e.target.value.replace(/\D/g, ''))}
                      placeholder="••••"
                      className="w-full text-center text-lg tracking-widest px-3 py-2 border border-gray-200 rounded-xl font-mono focus:border-blue-600 focus:outline-none"
                    />
                  </div>

                  <div>
                    <label className="text-xs font-semibold text-gray-700 block mb-1">
                      Confirm New 4-Digit PIN
                    </label>
                    <input
                      type="password"
                      maxLength={4}
                      value={confirmPin}
                      onChange={(e) => setConfirmPin(e.target.value.replace(/\D/g, ''))}
                      placeholder="••••"
                      className="w-full text-center text-lg tracking-widest px-3 py-2 border border-gray-200 rounded-xl font-mono focus:border-blue-600 focus:outline-none"
                    />
                  </div>
                </>
              )}

              {/* Error Alert */}
              {pinError && (
                <div className="flex items-center gap-1.5 p-2 rounded-lg bg-rose-50 border border-rose-200 text-xs text-rose-700">
                  <AlertCircle className="w-3.5 h-3.5 shrink-0" />
                  <span>{pinError}</span>
                </div>
              )}

              {/* Action Buttons */}
              <div className="pt-2 flex items-center justify-end gap-2">
                <button
                  type="button"
                  onClick={() => setShowPinModal(false)}
                  className="px-3.5 py-1.5 text-xs font-medium text-gray-600 hover:text-gray-900 bg-gray-100 hover:bg-gray-200 rounded-lg cursor-pointer"
                >
                  Cancel
                </button>
                <button
                  type="submit"
                  className="px-4 py-1.5 text-xs font-semibold text-white bg-blue-600 hover:bg-blue-700 rounded-lg shadow-2xs transition-all active:scale-95 cursor-pointer"
                >
                  {modalMode === 'create' && 'Save & Enable PIN'}
                  {modalMode === 'change' && 'Update PIN'}
                  {modalMode === 'remove' && 'Confirm Removal'}
                </button>
              </div>
            </form>
          </div>
        </div>
      )}
    </div>
  );
};
