/**
 * Multi-User Office LAN Collaboration & Workpaper Sign-Off Service
 * 
 * 100% Offline: Operates strictly within the local office network (WiFi/Switch router).
 * Zero Internet Required: Uses local private IP addressing (192.168.X.X / 10.X.X.X).
 * 
 * Provides:
 * - Local LAN IP discovery and connection link generation for article assistants
 * - Role-Based Access Controls (PrincipalPartner, AuditManager, ArticleAssistant)
 * - Workpaper Lock Management (ICAI SA-230 Auditor Sign-Off lock)
 */

import os from 'os';

export type UserRole = 'PrincipalPartner' | 'AuditManager' | 'ArticleAssistant';

export interface LanNetworkInfo {
  hostname: string;
  isLanActive: boolean;
  lanAddresses: {
    interfaceName: string;
    ip: string;
    url: string;
    type: 'IPv4' | 'IPv6';
  }[];
  primaryLanUrl: string;
  recommendedShareUrl: string;
}

export interface WorkpaperLockState {
  workpaperId: string;
  clientId: string;
  isLocked: boolean;
  lockedBy?: string;
  partnerMrn?: string;
  lockedAt?: string;
  status: 'Draft' | 'UnderReview' | 'AuditorApproved';
  reviewNotes?: string[];
}

export class OfficeLanService {
  private workpaperLocks: Map<string, WorkpaperLockState> = new Map();

  /**
   * 1. Discover all active local network (LAN) IP addresses on this computer
   */
  public getNetworkInfo(port: number = 5000): LanNetworkInfo {
    const interfaces = os.networkInterfaces();
    const lanAddresses: LanNetworkInfo['lanAddresses'] = [];
    const hostname = os.hostname();

    for (const [name, addrs] of Object.entries(interfaces)) {
      if (!addrs) continue;
      for (const addr of addrs) {
        // Skip internal loopback (127.0.0.1) to find actual LAN Wi-Fi/Ethernet IP
        if (!addr.internal && addr.family === 'IPv4') {
          lanAddresses.push({
            interfaceName: name,
            ip: addr.address,
            url: `http://${addr.address}:${port}`,
            type: 'IPv4',
          });
        }
      }
    }

    const primaryIp = lanAddresses.length > 0 ? lanAddresses[0].ip : '127.0.0.1';
    const primaryUrl = `http://${primaryIp}:${port}`;

    return {
      hostname,
      isLanActive: lanAddresses.length > 0,
      lanAddresses,
      primaryLanUrl: primaryUrl,
      recommendedShareUrl: primaryUrl,
    };
  }

  /**
   * 2. Role Permissions Check
   */
  public checkPermission(role: UserRole, action: 'upload' | 'reconcile' | 'sign_off' | 'lock_workpaper' | 'delete_data'): boolean {
    switch (action) {
      case 'upload':
      case 'reconcile':
        return true; // Articles, Managers, Partners can all reconcile
      case 'sign_off':
      case 'lock_workpaper':
        return role === 'PrincipalPartner' || role === 'AuditManager';
      case 'delete_data':
        return role === 'PrincipalPartner';
      default:
        return false;
    }
  }

  /**
   * 3. Lock Workpaper upon ICAI SA-230 Auditor Sign-off
   */
  public lockWorkpaper(
    workpaperId: string,
    clientId: string,
    partnerName: string,
    partnerMrn: string,
    notes?: string
  ): { success: boolean; state: WorkpaperLockState; message: string } {
    const lock: WorkpaperLockState = {
      workpaperId,
      clientId,
      isLocked: true,
      lockedBy: partnerName,
      partnerMrn,
      lockedAt: new Date().toISOString(),
      status: 'AuditorApproved',
      reviewNotes: notes ? [notes] : ['Official SA-230 Sign-off recorded.'],
    };

    this.workpaperLocks.set(workpaperId, lock);
    return {
      success: true,
      state: lock,
      message: `Workpaper ${workpaperId} successfully signed off & locked by ${partnerName} (MRN: ${partnerMrn}). Modifications by article staff are now prevented.`,
    };
  }

  /**
   * 4. Unlock Workpaper (Partner Override)
   */
  public unlockWorkpaper(workpaperId: string, partnerName: string): { success: boolean; message: string } {
    const existing = this.workpaperLocks.get(workpaperId);
    if (!existing) {
      return { success: false, message: 'Workpaper is not locked.' };
    }

    existing.isLocked = false;
    existing.status = 'UnderReview';
    existing.reviewNotes = existing.reviewNotes || [];
    existing.reviewNotes.push(`Unlocked by ${partnerName} at ${new Date().toISOString()}`);

    return {
      success: true,
      message: `Workpaper ${workpaperId} unlocked by Partner for further rectification.`,
    };
  }

  /**
   * 5. Get Workpaper Lock Status
   */
  public getWorkpaperStatus(workpaperId: string): WorkpaperLockState {
    return (
      this.workpaperLocks.get(workpaperId) || {
        workpaperId,
        clientId: '',
        isLocked: false,
        status: 'Draft',
        reviewNotes: [],
      }
    );
  }
}
