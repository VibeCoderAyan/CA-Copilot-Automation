/**
 * Tally Prime & Busy Direct Live Sync Service
 * 
 * Communicates with Tally Prime's local XML server over HTTP loopback (default 127.0.0.1:9000).
 * 100% Offline: Loopback networking never leaves local machine RAM/CPU.
 * 
 * Supports:
 * - Ping / Health Check on Tally local port
 * - Fetching active company profile
 * - Pulling Daybook vouchers directly into internal transaction objects
 * - Pulling Ledger accounts and closing balances
 * - Pulling Trial Balance
 * - Pushing reconciled journal vouchers directly into Tally Prime
 */

import http from 'http';

export interface TallyConnectionConfig {
  host: string;
  port: number;
  timeoutMs?: number;
}

export interface TallyCompany {
  name: string;
  guid?: string;
  booksFrom?: string;
  financialYearFrom?: string;
}

export interface TallyVoucher {
  date: string; // YYYYMMDD or DD-MM-YYYY
  voucherType: string;
  voucherNumber: string;
  ledgerName: string;
  amount: number;
  isDebit: boolean;
  narration?: string;
  partyName?: string;
}

export interface TallySyncResult {
  success: boolean;
  message: string;
  statusCode?: number;
  data?: any;
  rawXml?: string;
}

export class TallyLiveSyncService {
  private config: TallyConnectionConfig;

  constructor(config?: Partial<TallyConnectionConfig>) {
    this.config = {
      host: config?.host || '127.0.0.1',
      port: config?.port || 9000,
      timeoutMs: config?.timeoutMs || 5000,
    };
  }

  /**
   * Send XML payload to local Tally Prime XML server
   */
  public async executeTallyRequest(xmlPayload: string): Promise<string> {
    return new Promise((resolve, reject) => {
      const options = {
        hostname: this.config.host,
        port: this.config.port,
        path: '/',
        method: 'POST',
        headers: {
          'Content-Type': 'text/xml;charset=utf-8',
          'Content-Length': Buffer.byteLength(xmlPayload),
        },
        timeout: this.config.timeoutMs,
      };

      const req = http.request(options, (res) => {
        let responseData = '';
        res.setEncoding('utf8');
        res.on('data', (chunk) => {
          responseData += chunk;
        });
        res.on('end', () => {
          resolve(responseData);
        });
      });

      req.on('timeout', () => {
        req.destroy();
        reject(new Error(`Connection to Tally Prime at ${this.config.host}:${this.config.port} timed out after ${this.config.timeoutMs}ms.`));
      });

      req.on('error', (err: any) => {
        if (err.code === 'ECONNREFUSED') {
          reject(new Error(`Tally Prime is not running or ODBC/XML Server is disabled on ${this.config.host}:${this.config.port}. Enable in Tally: F1 -> Settings -> Connectivity -> Port 9000.`));
        } else {
          reject(err);
        }
      });

      req.write(xmlPayload);
      req.end();
    });
  }

  /**
   * 1. Check if Tally Prime is running and responding
   */
  public async checkHealth(): Promise<{ isConnected: boolean; host: string; port: number; message: string; companies?: string[] }> {
    const pingXml = `
      <ENVELOPE>
        <HEADER>
          <VERSION>1</VERSION>
          <TALLYREQUEST>Export</TALLYREQUEST>
          <TYPE>Data</TYPE>
          <ID>List of Companies</ID>
        </HEADER>
        <BODY>
          <DESC>
            <STATICVARIABLES>
              <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
            </STATICVARIABLES>
          </DESC>
        </BODY>
      </ENVELOPE>
    `.trim();

    try {
      const responseXml = await this.executeTallyRequest(pingXml);
      const companies = this.parseCompanyNamesFromXml(responseXml);
      return {
        isConnected: true,
        host: this.config.host,
        port: this.config.port,
        message: 'Tally Prime local XML server is active and responding.',
        companies,
      };
    } catch (err: any) {
      return {
        isConnected: false,
        host: this.config.host,
        port: this.config.port,
        message: err.message || 'Cannot connect to Tally Prime.',
      };
    }
  }

  /**
   * 2. Fetch Daybook Vouchers from Tally
   */
  public async fetchDaybook(companyName: string, fromDate: string, toDate: string): Promise<TallySyncResult> {
    const fromTallyDate = this.formatDateForTally(fromDate);
    const toTallyDate = this.formatDateForTally(toDate);

    const xml = `
      <ENVELOPE>
        <HEADER>
          <VERSION>1</VERSION>
          <TALLYREQUEST>Export</TALLYREQUEST>
          <TYPE>Data</TYPE>
          <ID>Day Book</ID>
        </HEADER>
        <BODY>
          <DESC>
            <STATICVARIABLES>
              <SVCURRENTCOMPANY>${this.escapeXml(companyName)}</SVCURRENTCOMPANY>
              <SVFROMDATE>${fromTallyDate}</SVFROMDATE>
              <SVTODATE>${toTallyDate}</SVTODATE>
              <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
            </STATICVARIABLES>
          </DESC>
        </BODY>
      </ENVELOPE>
    `.trim();

    try {
      const responseXml = await this.executeTallyRequest(xml);
      const vouchers = this.parseVouchersFromXml(responseXml);
      return {
        success: true,
        message: `Successfully fetched ${vouchers.length} vouchers from Tally Prime for ${companyName}.`,
        data: {
          companyName,
          period: { fromDate, toDate },
          count: vouchers.length,
          vouchers,
        },
        rawXml: responseXml.length > 5000 ? responseXml.substring(0, 5000) + '...[truncated]' : responseXml,
      };
    } catch (err: any) {
      return {
        success: false,
        message: err.message,
      };
    }
  }

  /**
   * 3. Fetch Trial Balance from Tally
   */
  public async fetchTrialBalance(companyName: string, fromDate: string, toDate: string): Promise<TallySyncResult> {
    const fromTallyDate = this.formatDateForTally(fromDate);
    const toTallyDate = this.formatDateForTally(toDate);

    const xml = `
      <ENVELOPE>
        <HEADER>
          <VERSION>1</VERSION>
          <TALLYREQUEST>Export</TALLYREQUEST>
          <TYPE>Data</TYPE>
          <ID>Trial Balance</ID>
        </HEADER>
        <BODY>
          <DESC>
            <STATICVARIABLES>
              <SVCURRENTCOMPANY>${this.escapeXml(companyName)}</SVCURRENTCOMPANY>
              <SVFROMDATE>${fromTallyDate}</SVFROMDATE>
              <SVTODATE>${toTallyDate}</SVTODATE>
              <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
            </STATICVARIABLES>
          </DESC>
        </BODY>
      </ENVELOPE>
    `.trim();

    try {
      const responseXml = await this.executeTallyRequest(xml);
      const accounts = this.parseTrialBalanceFromXml(responseXml);
      return {
        success: true,
        message: `Successfully extracted Trial Balance from Tally Prime (${accounts.length} ledgers).`,
        data: {
          companyName,
          accounts,
        },
      };
    } catch (err: any) {
      return {
        success: false,
        message: err.message,
      };
    }
  }

  /**
   * 4. Push Adjustment Vouchers (e.g. Bank Charges, Reconciled Adjustments) into Tally
   */
  public async pushVouchers(companyName: string, vouchers: TallyVoucher[]): Promise<TallySyncResult> {
    if (!vouchers || vouchers.length === 0) {
      return { success: false, message: 'No vouchers provided to push to Tally.' };
    }

    const voucherXmlItems = vouchers.map((v) => {
      const tallyDate = this.formatDateForTally(v.date);
      const isNegative = v.amount < 0;
      const amountVal = Math.abs(v.amount);

      return `
        <TALLYMESSAGE xmlns:UDF="TallyUDF">
          <VOUCHER VCHTYPE="${this.escapeXml(v.voucherType || 'Journal')}" ACTION="Create">
            <DATE>${tallyDate}</DATE>
            <VOUCHERTYPENAME>${this.escapeXml(v.voucherType || 'Journal')}</VOUCHERTYPENAME>
            <VOUCHERNUMBER>${this.escapeXml(v.voucherNumber || 'AUTO')}</VOUCHERNUMBER>
            <NARRATION>${this.escapeXml(v.narration || 'Reconciled entry posted via CA Copilot')}</NARRATION>
            <ALLLEDGERENTRIES.LIST>
              <LEDGERNAME>${this.escapeXml(v.ledgerName)}</LEDGERNAME>
              <ISDEEMEDPOSITIVE>${v.isDebit ? 'Yes' : 'No'}</ISDEEMEDPOSITIVE>
              <AMOUNT>${v.isDebit ? -amountVal : amountVal}</AMOUNT>
            </ALLLEDGERENTRIES.LIST>
          </VOUCHER>
        </TALLYMESSAGE>
      `.trim();
    }).join('\n');

    const envelope = `
      <ENVELOPE>
        <HEADER>
          <VERSION>1</VERSION>
          <TALLYREQUEST>Import</TALLYREQUEST>
          <TYPE>Data</TYPE>
          <ID>All Masters and Vouchers</ID>
        </HEADER>
        <BODY>
          <DESC>
            <STATICVARIABLES>
              <SVCURRENTCOMPANY>${this.escapeXml(companyName)}</SVCURRENTCOMPANY>
            </STATICVARIABLES>
          </DESC>
          <DATA>
            ${voucherXmlItems}
          </DATA>
        </BODY>
      </ENVELOPE>
    `.trim();

    try {
      const responseXml = await this.executeTallyRequest(envelope);
      const isSuccess = responseXml.includes('<CREATED>1</CREATED>') || responseXml.includes('<STATUS>1</STATUS>') || !responseXml.includes('<ERROR>');
      return {
        success: isSuccess,
        message: isSuccess
          ? `Successfully pushed ${vouchers.length} voucher(s) into Tally Prime.`
          : 'Tally returned an error response upon voucher import.',
        rawXml: responseXml,
      };
    } catch (err: any) {
      return {
        success: false,
        message: err.message,
      };
    }
  }

  // Helper parsers
  private parseCompanyNamesFromXml(xml: string): string[] {
    const companies: string[] = [];
    const regex = /<COMPANYNAME>([^<]+)<\/COMPANYNAME>|<NAME[^>]*>([^<]+)<\/NAME>/gi;
    let match;
    while ((match = regex.exec(xml)) !== null) {
      const name = (match[1] || match[2] || '').trim();
      if (name && !companies.includes(name) && !name.startsWith('$$')) {
        companies.push(name);
      }
    }
    return companies;
  }

  private parseVouchersFromXml(xml: string): any[] {
    const vouchers: any[] = [];
    const voucherBlocks = xml.split(/<VOUCHER\b/i).slice(1);

    for (const block of voucherBlocks) {
      const dateMatch = block.match(/<DATE>([^<]+)<\/DATE>/i);
      const typeMatch = block.match(/<VOUCHERTYPENAME>([^<]+)<\/VOUCHERTYPENAME>/i);
      const numberMatch = block.match(/<VOUCHERNUMBER>([^<]+)<\/VOUCHERNUMBER>/i);
      const narrationMatch = block.match(/<NARRATION>([^<]+)<\/NARRATION>/i);
      const amountMatch = block.match(/<AMOUNT>([^<]+)<\/AMOUNT>/i);
      const partyMatch = block.match(/<PARTYLEDGERNAME>([^<]+)<\/PARTYLEDGERNAME>/i);

      vouchers.push({
        date: dateMatch ? dateMatch[1] : '',
        voucherType: typeMatch ? typeMatch[1] : 'Journal',
        voucherNumber: numberMatch ? numberMatch[1] : '',
        narration: narrationMatch ? narrationMatch[1] : '',
        amount: amountMatch ? parseFloat(amountMatch[1]) || 0 : 0,
        partyName: partyMatch ? partyMatch[1] : '',
      });
    }

    return vouchers;
  }

  private parseTrialBalanceFromXml(xml: string): any[] {
    const accounts: any[] = [];
    const ledgerBlocks = xml.split(/<LEDGER\b/i).slice(1);

    for (const block of ledgerBlocks) {
      const nameMatch = block.match(/<NAME[^>]*>([^<]+)<\/NAME>/i);
      const parentMatch = block.match(/<PARENT>([^<]+)<\/PARENT>/i);
      const closingMatch = block.match(/<CLOSINGBALANCE>([^<]+)<\/CLOSINGBALANCE>/i);

      if (nameMatch) {
        accounts.push({
          accountName: nameMatch[1].trim(),
          group: parentMatch ? parentMatch[1].trim() : 'General',
          closingBalance: closingMatch ? parseFloat(closingMatch[1]) || 0 : 0,
        });
      }
    }

    return accounts;
  }

  private formatDateForTally(dateStr: string): string {
    if (!dateStr) return '20250401';
    // Clean and normalize
    const clean = dateStr.replace(/[^0-9]/g, '');
    if (clean.length === 8) {
      // YYYYMMDD
      if (parseInt(clean.substring(0, 4)) > 1900) return clean;
      // DDMMYYYY -> YYYYMMDD
      return `${clean.substring(4, 8)}${clean.substring(2, 4)}${clean.substring(0, 2)}`;
    }
    return '20250401';
  }

  private escapeXml(unsafe: string): string {
    if (!unsafe) return '';
    return unsafe
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&apos;');
  }
}
