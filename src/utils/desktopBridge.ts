/**
 * Desktop Application Bridge & Native System Abstraction Layer
 * Provides seamless bridge between Web Browser and Electron / Tauri Desktop Run-times
 */

export interface DesktopEnvironmentInfo {
  isDesktop: boolean;
  platform: 'windows' | 'mac' | 'linux' | 'browser';
  hasNativeFileSystem: boolean;
  hasNativeNotifications: boolean;
}

/**
 * Detects if the application is executing inside a native desktop container (Electron / Tauri / WebView2)
 */
export const isDesktopApp = (): boolean => {
  if (typeof window === 'undefined') return false;
  return Boolean(
    (window as any).electron ||
    (window as any).electronAPI ||
    (window as any).__TAURI__ ||
    navigator.userAgent.toLowerCase().includes('electron')
  );
};

/**
 * Returns comprehensive OS and runtime information
 */
export const getDesktopEnvironmentInfo = (): DesktopEnvironmentInfo => {
  const isDesktop = isDesktopApp();
  const ua = typeof navigator !== 'undefined' ? navigator.userAgent.toLowerCase() : '';
  
  let platform: DesktopEnvironmentInfo['platform'] = 'browser';
  if (ua.includes('win')) platform = 'windows';
  else if (ua.includes('mac')) platform = 'mac';
  else if (ua.includes('linux')) platform = 'linux';

  const hasNativeFileSystem = typeof window !== 'undefined' && 'showSaveFilePicker' in window;
  const hasNativeNotifications = typeof window !== 'undefined' && 'Notification' in window;

  return {
    isDesktop,
    platform,
    hasNativeFileSystem,
    hasNativeNotifications,
  };
};

/**
 * Unified native or browser file saving
 * In desktop mode: uses native OS Save As dialog if available
 * In web mode: triggers standard browser Blob download
 */
export const saveFileNative = async (
  content: Blob | string | ArrayBuffer,
  fileName: string,
  mimeType: string = 'application/octet-stream'
): Promise<boolean> => {
  const blob = content instanceof Blob ? content : new Blob([content], { type: mimeType });

  // 1. If native Windows File System Access API is supported (Chromium / Electron)
  if (typeof window !== 'undefined' && 'showSaveFilePicker' in window) {
    try {
      const ext = fileName.split('.').pop() || '';
      const handle = await (window as any).showSaveFilePicker({
        suggestedName: fileName,
        types: [
          {
            description: `${ext.toUpperCase()} Document`,
            accept: { [mimeType]: [`.${ext}`] },
          },
        ],
      });
      const writable = await handle.createWritable();
      await writable.write(blob);
      await writable.close();
      return true;
    } catch (err: any) {
      if (err.name === 'AbortError') {
        // User cancelled the save dialog
        return false;
      }
      // Fallback to standard download if picker errors
    }
  }

  // 2. Standard Blob fallback
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = fileName;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
  return true;
};

/**
 * Displays OS-level desktop notification
 */
export const sendDesktopNotification = (_title: string, _body: string): void => {
  // All desktop notifications disabled per user request
  return;
};
