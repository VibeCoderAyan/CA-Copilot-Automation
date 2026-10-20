/**
 * Multi-Client Practice Vault & Storage Utility
 * 100% Local-First & Offline Storage for CA Firms
 */

export type EntityType =
  | 'Private Limited'
  | 'Public Limited'
  | 'LLP'
  | 'Partnership Firm'
  | 'Sole Proprietorship'
  | 'Individual / HUF'
  | 'Trust / Society';

export interface ClientProfile {
  id: string;
  name: string;
  tradeName?: string;
  gstin: string;
  pan: string;
  entityType: EntityType;
  financialYear: string;
  assessmentYear: string;
  contactPerson?: string;
  email?: string;
  phone?: string;
  city?: string;
  createdAt: string;
  updatedAt: string;
}

const CLIENTS_STORAGE_KEY = 'ca_copilot_clients_vault';
const ACTIVE_CLIENT_KEY = 'ca_copilot_active_client_id';

export const BLANK_CLIENT: ClientProfile = {
  id: 'client_default',
  name: 'Primary Auditee (Set in Vault)',
  tradeName: '',
  gstin: '',
  pan: '',
  entityType: 'Private Limited',
  financialYear: 'FY 2024-25',
  assessmentYear: 'AY 2025-26',
  createdAt: new Date().toISOString(),
  updatedAt: new Date().toISOString(),
};

/**
 * Extract 10-digit PAN from 15-digit GSTIN (Characters 3 to 12)
 */
export const derivePanFromGstin = (gstin: string): string => {
  const clean = gstin.trim().toUpperCase();
  if (clean.length >= 12) {
    const candidate = clean.substring(2, 12);
    if (/^[A-Z]{5}[0-9]{4}[A-Z]{1}$/.test(candidate)) {
      return candidate;
    }
  }
  return '';
};

/**
 * Validate GSTIN format
 */
export const isValidGstin = (gstin: string): boolean => {
  return /^[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z]{1}[1-9A-Z]{1}Z[0-9A-Z]{1}$/.test(gstin.trim().toUpperCase());
};

/**
 * Validate PAN format
 */
export const isValidPan = (pan: string): boolean => {
  return /^[A-Z]{5}[0-9]{4}[A-Z]{1}$/.test(pan.trim().toUpperCase());
};

/**
 * Retrieve all practice clients from local storage (zero fake mock data)
 */
export const getClients = (): ClientProfile[] => {
  const saved = localStorage.getItem(CLIENTS_STORAGE_KEY);
  if (!saved) {
    return [];
  }
  try {
    const parsed = JSON.parse(saved);
    if (!Array.isArray(parsed)) return [];
    // Purge any old fake demo entries
    const cleaned = parsed.filter(
      (c) => c.id !== 'client_sharma_ent' && c.id !== 'client_apex_eng'
    );
    if (cleaned.length !== parsed.length) {
      localStorage.setItem(CLIENTS_STORAGE_KEY, JSON.stringify(cleaned));
    }
    return cleaned;
  } catch (e) {
    return [];
  }
};

/**
 * Retrieve currently active client
 */
export const getActiveClient = (): ClientProfile => {
  const clients = getClients();
  const activeId = localStorage.getItem(ACTIVE_CLIENT_KEY);
  const found = clients.find((c) => c.id === activeId);
  return found || clients[0] || BLANK_CLIENT;
};

/**
 * Set active client ID
 */
export const setActiveClient = (clientId: string): void => {
  localStorage.setItem(ACTIVE_CLIENT_KEY, clientId);
  window.dispatchEvent(new CustomEvent('ca_copilot_client_switched', { detail: { clientId } }));
};

/**
 * Add a new client profile
 */
export const addClient = (
  data: Omit<ClientProfile, 'id' | 'createdAt' | 'updatedAt'>
): ClientProfile => {
  const clients = getClients();
  const id = `client_${Date.now()}_${Math.random().toString(36).substring(2, 6)}`;
  const now = new Date().toISOString();

  const pan = data.pan || derivePanFromGstin(data.gstin) || '';

  const newClient: ClientProfile = {
    ...data,
    id,
    pan,
    createdAt: now,
    updatedAt: now,
  };

  const updatedList = [newClient, ...clients];
  localStorage.setItem(CLIENTS_STORAGE_KEY, JSON.stringify(updatedList));
  setActiveClient(id);
  return newClient;
};

/**
 * Update existing client profile
 */
export const updateClient = (id: string, updates: Partial<ClientProfile>): void => {
  const clients = getClients();
  const updatedList = clients.map((c) => {
    if (c.id === id) {
      return {
        ...c,
        ...updates,
        updatedAt: new Date().toISOString(),
      };
    }
    return c;
  });
  localStorage.setItem(CLIENTS_STORAGE_KEY, JSON.stringify(updatedList));
};

/**
 * Delete client profile
 */
export const deleteClient = (id: string): void => {
  const clients = getClients();
  const filtered = clients.filter((c) => c.id !== id);
  localStorage.setItem(CLIENTS_STORAGE_KEY, JSON.stringify(filtered));

  const activeId = localStorage.getItem(ACTIVE_CLIENT_KEY);
  if (activeId === id) {
    const nextClient = filtered[0];
    if (nextClient) {
      setActiveClient(nextClient.id);
    }
  }
};
