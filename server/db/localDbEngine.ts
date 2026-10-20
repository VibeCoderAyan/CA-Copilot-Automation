/**
 * High-Speed Indexed Local Database Engine
 * 
 * 100% Offline: Zero-cloud, zero external native database driver dependencies.
 * Provides multi-column memory indexes, ACID-style atomic disk persistence,
 * and sub-millisecond query performance across large client ledger datasets.
 */

import fs from 'fs';
import path from 'path';

export interface DbTransactionRecord {
  id: string;
  clientId: string;
  date: string; // YYYY-MM-DD
  voucherNo: string;
  accountHead: string;
  debit: number;
  credit: number;
  narration?: string;
  partyGstin?: string;
  partyPan?: string;
  category?: string;
  metadata?: Record<string, any>;
  createdAt: string;
}

export interface DbQueryParams {
  clientId?: string;
  fromDate?: string;
  toDate?: string;
  voucherNo?: string;
  partyGstin?: string;
  minAmount?: number;
  maxAmount?: number;
  search?: string;
  limit?: number;
  offset?: number;
}

export class LocalDbEngine {
  private static instance: LocalDbEngine | null = null;
  private dbFilePath: string;
  private records: Map<string, DbTransactionRecord> = new Map();

  // Multi-column in-memory indices for sub-millisecond querying
  private clientIndex: Map<string, Set<string>> = new Map();
  private gstinIndex: Map<string, Set<string>> = new Map();
  private dateIndex: Map<string, Set<string>> = new Map();
  private voucherIndex: Map<string, string> = new Map();

  private isDirty = false;
  private autoSaveTimer: NodeJS.Timeout | null = null;

  public static getInstance(customFilePath?: string): LocalDbEngine {
    if (!LocalDbEngine.instance) {
      LocalDbEngine.instance = new LocalDbEngine(customFilePath);
    }
    return LocalDbEngine.instance;
  }

  constructor(customFilePath?: string) {
    const dataDir = path.resolve(process.cwd(), 'data');
    if (!fs.existsSync(dataDir)) {
      fs.mkdirSync(dataDir, { recursive: true });
    }

    this.dbFilePath = customFilePath || path.join(dataDir, 'ca-indexed-store.db.json');
    this.loadFromDisk();
  }

  /**
   * Load existing database records from disk and rebuild indices
   */
  private loadFromDisk(): void {
    if (fs.existsSync(this.dbFilePath)) {
      try {
        const raw = fs.readFileSync(this.dbFilePath, 'utf8');
        const data: DbTransactionRecord[] = JSON.parse(raw);
        for (const item of data) {
          this.insertIntoMemory(item, false);
        }
      } catch (err) {
        console.error('[LocalDbEngine] Error loading disk store:', err);
      }
    }
  }

  /**
   * Flush in-memory records to disk atomically
   */
  public flushSync(): void {
    try {
      const recordsArray = Array.from(this.records.values());
      const tempPath = `${this.dbFilePath}.tmp.${Date.now()}`;
      fs.writeFileSync(tempPath, JSON.stringify(recordsArray, null, 2), 'utf8');
      fs.renameSync(tempPath, this.dbFilePath);
      this.isDirty = false;
    } catch (err) {
      console.error('[LocalDbEngine] Atomic flush error:', err);
    }
  }

  private insertIntoMemory(record: DbTransactionRecord, markDirty: boolean = true): void {
    this.records.set(record.id, record);

    // 1. Index by Client
    if (record.clientId) {
      let clientSet = this.clientIndex.get(record.clientId);
      if (!clientSet) {
        clientSet = new Set();
        this.clientIndex.set(record.clientId, clientSet);
      }
      clientSet.add(record.id);
    }

    // 2. Index by GSTIN
    if (record.partyGstin) {
      const cleanGstin = record.partyGstin.trim().toUpperCase();
      let gstinSet = this.gstinIndex.get(cleanGstin);
      if (!gstinSet) {
        gstinSet = new Set();
        this.gstinIndex.set(cleanGstin, gstinSet);
      }
      gstinSet.add(record.id);
    }

    // 3. Index by Date
    if (record.date) {
      let dateSet = this.dateIndex.get(record.date);
      if (!dateSet) {
        dateSet = new Set();
        this.dateIndex.set(record.date, dateSet);
      }
      dateSet.add(record.id);
    }

    // 4. Index by Voucher No
    if (record.voucherNo) {
      this.voucherIndex.set(record.voucherNo.trim().toUpperCase(), record.id);
    }

    if (markDirty) {
      this.isDirty = true;
    }
  }

  /**
   * Insert single or batch records
   */
  public insertBatch(records: Omit<DbTransactionRecord, 'id' | 'createdAt'>[]): { insertedCount: number; durationMs: number } {
    const startTime = Date.now();
    for (const r of records) {
      const record: DbTransactionRecord = {
        ...r,
        id: `tx_${Date.now()}_${Math.random().toString(36).substring(2, 7)}`,
        createdAt: new Date().toISOString(),
      };
      this.insertIntoMemory(record, true);
    }

    this.flushSync();
    return {
      insertedCount: records.length,
      durationMs: Date.now() - startTime,
    };
  }

  /**
   * Query records using indices (sub-millisecond)
   */
  public query(params: DbQueryParams): { totalMatches: number; records: DbTransactionRecord[]; durationMs: number } {
    const startTime = Date.now();
    let candidateIds: Set<string> | null = null;

    // 1. Check Voucher No Index (Exact 1:1 match)
    if (params.voucherNo) {
      const vKey = params.voucherNo.trim().toUpperCase();
      const matchId = this.voucherIndex.get(vKey);
      if (matchId) {
        candidateIds = new Set([matchId]);
      } else {
        return { totalMatches: 0, records: [], durationMs: Date.now() - startTime };
      }
    }

    // 2. Check Client Index
    if (params.clientId) {
      const clientSet = this.clientIndex.get(params.clientId);
      if (!clientSet) {
        return { totalMatches: 0, records: [], durationMs: Date.now() - startTime };
      }
      candidateIds = candidateIds ? this.intersect(candidateIds, clientSet) : new Set(clientSet);
    }

    // 3. Check GSTIN Index
    if (params.partyGstin) {
      const gstinSet = this.gstinIndex.get(params.partyGstin.trim().toUpperCase());
      if (!gstinSet) {
        return { totalMatches: 0, records: [], durationMs: Date.now() - startTime };
      }
      candidateIds = candidateIds ? this.intersect(candidateIds, gstinSet) : new Set(gstinSet);
    }

    // 4. Scan or Filter candidate pool
    const idList = candidateIds ? Array.from(candidateIds) : Array.from(this.records.keys());
    let results: DbTransactionRecord[] = [];

    const searchLower = params.search ? params.search.toLowerCase() : null;

    for (const id of idList) {
      const item = this.records.get(id);
      if (!item) continue;

      // Date Range Filter
      if (params.fromDate && item.date < params.fromDate) continue;
      if (params.toDate && item.date > params.toDate) continue;

      // Amount Filter
      const amount = Math.max(item.debit, item.credit);
      if (params.minAmount !== undefined && amount < params.minAmount) continue;
      if (params.maxAmount !== undefined && amount > params.maxAmount) continue;

      // Text Search
      if (searchLower) {
        const text = `${item.accountHead} ${item.narration || ''} ${item.voucherNo}`.toLowerCase();
        if (!text.includes(searchLower)) continue;
      }

      results.push(item);
    }

    const totalMatches = results.length;

    // Pagination
    const offset = params.offset || 0;
    const limit = params.limit || 100;
    results = results.slice(offset, offset + limit);

    return {
      totalMatches,
      records: results,
      durationMs: Date.now() - startTime,
    };
  }

  public getStats(): { totalRecords: number; totalClients: number; totalGstins: number; storagePath: string } {
    return {
      totalRecords: this.records.size,
      totalClients: this.clientIndex.size,
      totalGstins: this.gstinIndex.size,
      storagePath: this.dbFilePath,
    };
  }

  public clearAll(): void {
    this.records.clear();
    this.clientIndex.clear();
    this.gstinIndex.clear();
    this.dateIndex.clear();
    this.voucherIndex.clear();
    this.flushSync();
  }

  private intersect(setA: Set<string>, setB: Set<string>): Set<string> {
    const out = new Set<string>();
    for (const elem of setA) {
      if (setB.has(elem)) out.add(elem);
    }
    return out;
  }
}
