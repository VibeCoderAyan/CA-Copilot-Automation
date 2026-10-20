/**
 * Universal Reusable Matching Engine
 * Reconciles Bank vs Ledger, GST Register vs GSTR-2B, Invoices vs Purchase Register
 */

import type { ExtractedRecord, RecordMatch, MatchingConfig, MatchType } from '../types';
import { normalizeInvoiceNumber } from '../../src/utils/gstReconciliationEngine';

export const DEFAULT_MATCHING_CONFIG: MatchingConfig = {
  dateToleranceDays: 3,
  amountTolerance: 1.0,
  checkDirection: true,
  fuzzyNarrationThreshold: 0.6,
  allowOneToMany: true,
  preventReusingMatches: true,
};

export class UniversalMatchingEngine {
  /**
   * Universal record matching pipeline
   */
  public static executeMatching(
    primaryRecords: ExtractedRecord[],
    secondaryRecords: ExtractedRecord[],
    workflowRunId: string,
    userConfig: Partial<MatchingConfig> = {}
  ): {
    matches: RecordMatch[];
    unmatchedPrimary: ExtractedRecord[];
    unmatchedSecondary: ExtractedRecord[];
    summary: {
      totalPrimary: number;
      totalSecondary: number;
      exactMatches: number;
      probableMatches: number;
      partialMatches: number;
      unmatchedCount: number;
    };
  } {
    const config: MatchingConfig = { ...DEFAULT_MATCHING_CONFIG, ...userConfig };
    const matches: RecordMatch[] = [];
    const usedSecondaryIds = new Set<string>();

    const parseDateToTime = (dStr?: string): number => {
      if (!dStr) return 0;
      // standard DD-MM-YYYY or YYYY-MM-DD
      const parts = dStr.split(/[-/.]/);
      if (parts.length === 3) {
        if (parts[0].length === 4) return new Date(`${parts[0]}-${parts[1]}-${parts[2]}`).getTime();
        return new Date(`${parts[2]}-${parts[1]}-${parts[0]}`).getTime();
      }
      return new Date(dStr).getTime() || 0;
    };

    // 1. Pass 1: Exact Match (Amount + Exact RefNo/InvoiceNo OR Amount + Exact Date)
    for (const pri of primaryRecords) {
      const priAmt = pri.normalizedFields.amount || Math.max(pri.normalizedFields.debit || 0, pri.normalizedFields.credit || 0);
      const priRef = (pri.normalizedFields.refNo || pri.normalizedFields.invoiceNo || '').trim();
      const priNormRef = normalizeInvoiceNumber(priRef);
      const priTime = parseDateToTime(pri.normalizedFields.date);

      let matchedSec: ExtractedRecord | undefined;
      let matchType: MatchType = 'exact';
      let matchScore = 100;
      let explanation = '';

      // Check candidate secondaries
      for (const sec of secondaryRecords) {
        if (config.preventReusingMatches && usedSecondaryIds.has(sec.id)) continue;

        const secAmt = sec.normalizedFields.amount || Math.max(sec.normalizedFields.debit || 0, sec.normalizedFields.credit || 0);
        const secRef = (sec.normalizedFields.refNo || sec.normalizedFields.invoiceNo || '').trim();
        const secNormRef = normalizeInvoiceNumber(secRef);
        const secTime = parseDateToTime(sec.normalizedFields.date);

        // Direction check: if comparing bank statement vs ledger
        if (config.checkDirection) {
          const priWdr = pri.normalizedFields.withdrawal || pri.normalizedFields.debit || 0;
          const priDep = pri.normalizedFields.deposit || pri.normalizedFields.credit || 0;
          const secCr = sec.normalizedFields.credit || 0;
          const secDr = sec.normalizedFields.debit || 0;

          const isBankToLedger = priWdr > 0 && secCr > 0;
          const isDepositToLedger = priDep > 0 && secDr > 0;
          if ((priWdr > 0 || priDep > 0) && !isBankToLedger && !isDepositToLedger) {
            // Not matching directions
            continue;
          }
        }

        const amtDiff = Math.abs(priAmt - secAmt);
        const daysDiff = priTime && secTime ? Math.abs(priTime - secTime) / (1000 * 60 * 60 * 24) : 999;

        // Exact Ref and Amount Match
        if (priNormRef && secNormRef && priNormRef === secNormRef && amtDiff <= config.amountTolerance) {
          matchedSec = sec;
          matchType = 'exact';
          matchScore = 100;
          explanation = `Exact Match on Identifier "${priRef}" and Amount (₹${priAmt.toFixed(2)} vs ₹${secAmt.toFixed(2)})`;
          break;
        }

        // Exact Date and Exact Amount Match
        if (daysDiff === 0 && amtDiff <= 0.01) {
          matchedSec = sec;
          matchType = 'exact';
          matchScore = 98;
          explanation = `Exact Match on Transaction Date (${pri.normalizedFields.date}) and Amount (₹${priAmt.toFixed(2)})`;
          break;
        }
      }

      if (matchedSec) {
        usedSecondaryIds.add(matchedSec.id);
        matches.push({
          id: `match-${Date.now()}-${matches.length + 1}`,
          workflowRunId,
          matchType,
          matchScore,
          explanation,
          primaryRecord: pri,
          matchedRecords: [matchedSec],
          isConfirmed: matchType === 'exact',
          varianceAmount: Math.abs(priAmt - (matchedSec.normalizedFields.amount || 0)),
          resolutionStatus: 'auto_matched',
        });
        continue;
      }

      // 2. Pass 2: Probable Match (Amount within tolerance + Date within tolerance)
      for (const sec of secondaryRecords) {
        if (config.preventReusingMatches && usedSecondaryIds.has(sec.id)) continue;

        const secAmt = sec.normalizedFields.amount || Math.max(sec.normalizedFields.debit || 0, sec.normalizedFields.credit || 0);
        const secTime = parseDateToTime(sec.normalizedFields.date);
        const amtDiff = Math.abs(priAmt - secAmt);
        const daysDiff = priTime && secTime ? Math.abs(priTime - secTime) / (1000 * 60 * 60 * 24) : 999;

        if (amtDiff <= config.amountTolerance && daysDiff <= config.dateToleranceDays) {
          matchedSec = sec;
          matchType = 'probable';
          matchScore = Math.max(70, Math.round(95 - daysDiff * 5 - amtDiff * 5));
          explanation = `Probable Match: Amount difference ₹${amtDiff.toFixed(2)} (within ₹${config.amountTolerance}) and ${daysDiff.toFixed(0)} days variance`;
          break;
        }
      }

      if (matchedSec) {
        usedSecondaryIds.add(matchedSec.id);
        matches.push({
          id: `match-${Date.now()}-${matches.length + 1}`,
          workflowRunId,
          matchType,
          matchScore,
          explanation,
          primaryRecord: pri,
          matchedRecords: [matchedSec],
          isConfirmed: false,
          varianceAmount: Math.abs(priAmt - (matchedSec.normalizedFields.amount || 0)),
          varianceDays: Math.round(Math.abs(priTime - parseDateToTime(matchedSec.normalizedFields.date)) / (1000 * 60 * 60 * 24)),
          resolutionStatus: 'auto_matched',
        });
        continue;
      }

      // 3. Pass 3: One-to-Many Match (if allowed)
      if (config.allowOneToMany) {
        const potentialGroup: ExtractedRecord[] = [];
        let accumulatedSecAmt = 0;

        for (const sec of secondaryRecords) {
          if (config.preventReusingMatches && usedSecondaryIds.has(sec.id)) continue;
          const secAmt = sec.normalizedFields.amount || Math.max(sec.normalizedFields.debit || 0, sec.normalizedFields.credit || 0);
          if (accumulatedSecAmt + secAmt <= priAmt + config.amountTolerance) {
            potentialGroup.push(sec);
            accumulatedSecAmt += secAmt;
            if (Math.abs(priAmt - accumulatedSecAmt) <= config.amountTolerance && potentialGroup.length > 1) {
              break;
            }
          }
        }

        if (Math.abs(priAmt - accumulatedSecAmt) <= config.amountTolerance && potentialGroup.length > 1) {
          for (const g of potentialGroup) usedSecondaryIds.add(g.id);
          matches.push({
            id: `match-${Date.now()}-${matches.length + 1}`,
            workflowRunId,
            matchType: 'partial',
            matchScore: 85,
            explanation: `1-to-Many Group Match: Primary ₹${priAmt.toFixed(2)} matched to ${potentialGroup.length} secondary records totalling ₹${accumulatedSecAmt.toFixed(2)}`,
            primaryRecord: pri,
            matchedRecords: potentialGroup,
            isConfirmed: false,
            varianceAmount: Math.abs(priAmt - accumulatedSecAmt),
            resolutionStatus: 'auto_matched',
          });
          continue;
        }
      }

      // 4. Record remains Unmatched
      matches.push({
        id: `unmatch-${Date.now()}-${matches.length + 1}`,
        workflowRunId,
        matchType: 'unmatched',
        matchScore: 0,
        explanation: 'No matching transaction found within configured date and amount tolerances',
        primaryRecord: pri,
        matchedRecords: [],
        isConfirmed: false,
        resolutionStatus: 'auto_matched',
      });
    }

    const matchedPrimaryIds = new Set(matches.filter((m) => m.matchType !== 'unmatched').map((m) => m.primaryRecord.id));
    const unmatchedPrimary = primaryRecords.filter((p) => !matchedPrimaryIds.has(p.id));
    const unmatchedSecondary = secondaryRecords.filter((s) => !usedSecondaryIds.has(s.id));

    return {
      matches,
      unmatchedPrimary,
      unmatchedSecondary,
      summary: {
        totalPrimary: primaryRecords.length,
        totalSecondary: secondaryRecords.length,
        exactMatches: matches.filter((m) => m.matchType === 'exact').length,
        probableMatches: matches.filter((m) => m.matchType === 'probable').length,
        partialMatches: matches.filter((m) => m.matchType === 'partial').length,
        unmatchedCount: matches.filter((m) => m.matchType === 'unmatched').length + unmatchedSecondary.length,
      },
    };
  }
}
