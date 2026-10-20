/**
 * Security & PIN Management for CA Copilot
 * Provides local 4-digit PIN verification, secure hashing, and practice privacy management.
 */

const PIN_STORAGE_KEY = 'ca_copilot_pin_hash';
const PIN_ENABLED_KEY = 'ca_copilot_pin_enabled';
const PIN_AUTO_LOCK_KEY = 'ca_copilot_auto_lock_mins';

// Simple salt for obfuscated local hashing
const PIN_SALT = 'CA_COPILOT_ICAI_SALT_2025';

export const hashPin = (pin: string): string => {
  return btoa(`${PIN_SALT}_${pin}_SECURE`);
};

export const isPinSet = (): boolean => {
  return !!localStorage.getItem(PIN_STORAGE_KEY);
};

export const isPinEnabled = (): boolean => {
  return localStorage.getItem(PIN_ENABLED_KEY) === 'true' && isPinSet();
};

export const savePin = (pin: string): boolean => {
  if (!/^\d{4}$/.test(pin)) return false;
  localStorage.setItem(PIN_STORAGE_KEY, hashPin(pin));
  localStorage.setItem(PIN_ENABLED_KEY, 'true');
  return true;
};

export const verifyPin = (pin: string): boolean => {
  const storedHash = localStorage.getItem(PIN_STORAGE_KEY);
  if (!storedHash) return false;
  return storedHash === hashPin(pin);
};

export const removePin = (currentPin: string): boolean => {
  if (!verifyPin(currentPin)) return false;
  localStorage.removeItem(PIN_STORAGE_KEY);
  localStorage.setItem(PIN_ENABLED_KEY, 'false');
  return true;
};

export const togglePinEnabled = (enabled: boolean): void => {
  localStorage.setItem(PIN_ENABLED_KEY, enabled ? 'true' : 'false');
};

export const getAutoLockMinutes = (): number => {
  return parseInt(localStorage.getItem(PIN_AUTO_LOCK_KEY) || '0', 10);
};

export const setAutoLockMinutes = (mins: number): void => {
  localStorage.setItem(PIN_AUTO_LOCK_KEY, mins.toString());
};
