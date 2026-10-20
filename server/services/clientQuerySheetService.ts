/**
 * Automated Client Query Sheet & Formal Audit Observation Generator
 * 
 * 100% Offline: Generates formal auditor discrepancy sheets, Excel response registers,
 * and standard .eml draft files conforming to ICAI SA-230 / SA-505 professional standards.
 */

import * as XLSX from 'xlsx';

export interface AuditQueryItem {
  srNo: number;
  date: string;
  referenceNo: string;
  ledgerHead: string;
  partyName?: string;
  amount: number;
  discrepancyType: 'Missing_GSTR2B' | 'Bank_Unreconciled' | 'Sec_40A3_Cash_Limit' | 'TDS_Short_Deduction' | 'Negative_Balance' | 'Other';
  auditObservation: string;
  requiredClientAction: string;
}

export interface ClientQueryPackage {
  clientName: string;
  financialYear: string;
  assessmentYear: string;
  firmName: string;
  firmFrn: string;
  partnerName: string;
  partnerMrn: string;
  clientEmail?: string;
  queryDate: string;
  queryItems: AuditQueryItem[];
}

export interface GeneratedQueryArtifacts {
  htmlLetter: string;
  emlRawContent: string;
  excelBuffer: Buffer;
  totalQueries: number;
  totalDisputedAmount: number;
}

export class ClientQuerySheetService {
  /**
   * Generate formal audit query letter, Excel response template, and .eml email draft
   */
  public generateQueryPackage(pkg: ClientQueryPackage): GeneratedQueryArtifacts {
    const totalDisputedAmount = pkg.queryItems.reduce((acc, q) => acc + (q.amount || 0), 0);
    const htmlLetter = this.buildHtmlLetter(pkg, totalDisputedAmount);
    const emlRawContent = this.buildEmlDraft(pkg, totalDisputedAmount);
    const excelBuffer = this.buildExcelQueryRegister(pkg);

    return {
      htmlLetter,
      emlRawContent,
      excelBuffer,
      totalQueries: pkg.queryItems.length,
      totalDisputedAmount,
    };
  }

  private buildHtmlLetter(pkg: ClientQueryPackage, totalAmount: number): string {
    const queryRows = pkg.queryItems.map((q) => `
      <tr>
        <td style="padding: 8px 12px; border: 1px solid #E5E7EB; text-align: center; font-weight: 600;">${q.srNo}</td>
        <td style="padding: 8px 12px; border: 1px solid #E5E7EB; white-space: nowrap;">${q.date}</td>
        <td style="padding: 8px 12px; border: 1px solid #E5E7EB; font-family: monospace;">${q.referenceNo}</td>
        <td style="padding: 8px 12px; border: 1px solid #E5E7EB;"><strong>${q.ledgerHead}</strong>${q.partyName ? `<br/><span style="color:#6B7280; font-size:11px;">${q.partyName}</span>` : ''}</td>
        <td style="padding: 8px 12px; border: 1px solid #E5E7EB; text-align: right; font-family: monospace; font-weight: 600;">₹${q.amount.toLocaleString('en-IN', { minimumFractionDigits: 2 })}</td>
        <td style="padding: 8px 12px; border: 1px solid #E5E7EB; color: #111827;">${q.auditObservation}</td>
        <td style="padding: 8px 12px; border: 1px solid #E5E7EB; color: #1F2937; background: #F9FAFB;">${q.requiredClientAction}</td>
      </tr>
    `).join('\n');

    return `
      <!DOCTYPE html>
      <html>
      <head>
        <meta charset="utf-8" />
        <title>Audit Query Sheet - ${pkg.clientName}</title>
        <style>
          body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; color: #111827; background: #FFFFFF; padding: 32px; font-size: 13px; line-height: 1.5; }
          .header { border-bottom: 2px solid #111827; padding-bottom: 16px; margin-bottom: 24px; }
          .title { font-size: 18px; font-weight: 800; margin: 0; color: #111827; }
          .sub { font-size: 11px; color: #6B7280; margin-top: 4px; }
          .meta-box { background: #F8F9FA; border: 1px solid #E5E7EB; border-radius: 8px; padding: 12px 16px; margin-bottom: 20px; }
          table { width: 100%; border-collapse: collapse; margin-top: 16px; font-size: 12px; }
          th { background: #F3F4F6; color: #374151; font-weight: 700; text-align: left; padding: 8px 12px; border: 1px solid #D1D5DB; }
          .footer { margin-top: 32px; padding-top: 16px; border-top: 1px solid #E5E7EB; font-size: 12px; color: #4B5563; }
        </style>
      </head>
      <body>
        <div class="header">
          <h1 class="title">${pkg.firmName}</h1>
          <div class="sub">Chartered Accountants • ICAI Firm Reg. No: ${pkg.firmFrn}</div>
        </div>

        <div class="meta-box">
          <div style="font-weight: 700; font-size: 14px; margin-bottom: 4px;">AUDIT OBSERVATION & REQUISITION QUERY SHEET</div>
          <div><strong>Client Entity:</strong> ${pkg.clientName} | <strong>Financial Year:</strong> ${pkg.financialYear} (${pkg.assessmentYear})</div>
          <div><strong>Date:</strong> ${pkg.queryDate} | <strong>Total Discrepancies Flagged:</strong> ${pkg.queryItems.length} (Total Value: ₹${totalAmount.toLocaleString('en-IN', { minimumFractionDigits: 2 })})</div>
        </div>

        <p>Dear Sir / Madam,</p>
        <p>During our statutory audit and reconciliation scrutiny for the period ${pkg.financialYear}, the following discrepancies and unvouched balances were identified. Kindly provide necessary clarification, missing invoices, or supporting documentary evidence against each item:</p>

        <table>
          <thead>
            <tr>
              <th style="width: 40px; text-align: center;">#</th>
              <th>Date</th>
              <th>Ref / Voucher</th>
              <th>Account Head / Party</th>
              <th style="text-align: right;">Amount (₹)</th>
              <th>Audit Discrepancy Observation</th>
              <th>Required Action / Documents</th>
            </tr>
          </thead>
          <tbody>
            ${queryRows}
          </tbody>
        </table>

        <div class="footer">
          <p>Please return the attached Excel response sheet with your remarks at the earliest.</p>
          <br/>
          <strong>For ${pkg.firmName}</strong><br/>
          Chartered Accountants<br/>
          <strong>${pkg.partnerName}</strong>, Partner (M.No: ${pkg.partnerMrn})
        </div>
      </body>
      </html>
    `.trim();
  }

  private buildEmlDraft(pkg: ClientQueryPackage, totalAmount: number): string {
    const boundary = `----=_Part_${Date.now()}`;
    const dateFormatted = new Date().toUTCString();
    const recipient = pkg.clientEmail || 'accounts@client.com';
    const subject = `Audit Queries & Requisition for ${pkg.financialYear} - ${pkg.clientName}`;

    return [
      `Date: ${dateFormatted}`,
      `From: "${pkg.partnerName}" <audit@cafirm.com>`,
      `To: "${pkg.clientName}" <${recipient}>`,
      `Subject: ${subject}`,
      `MIME-Version: 1.0`,
      `Content-Type: multipart/alternative; boundary="${boundary}"`,
      ``,
      `--${boundary}`,
      `Content-Type: text/plain; charset=UTF-8`,
      `Content-Transfer-Encoding: 7bit`,
      ``,
      `Dear Sir/Madam,`,
      ``,
      `In the course of our audit for ${pkg.clientName} for ${pkg.financialYear}, we have identified ${pkg.queryItems.length} reconciliation and voucher discrepancies totaling ₹${totalAmount.toLocaleString('en-IN', { minimumFractionDigits: 2 })}.`,
      ``,
      `Please review the items and provide missing supporting bills or clarifications:`,
      ``,
      ...pkg.queryItems.map(
        (q) => `[Item ${q.srNo}] ${q.date} | ${q.referenceNo} | ${q.ledgerHead} | ₹${q.amount} => ${q.auditObservation} (Action: ${q.requiredClientAction})`
      ),
      ``,
      `Regards,`,
      `${pkg.partnerName}`,
      `${pkg.firmName}, Chartered Accountants (FRN: ${pkg.firmFrn})`,
      ``,
      `--${boundary}--`,
    ].join('\r\n');
  }

  private buildExcelQueryRegister(pkg: ClientQueryPackage): Buffer {
    const wb = XLSX.utils.book_new();

    const excelData = pkg.queryItems.map((q) => ({
      'Sr No': q.srNo,
      'Transaction Date': q.date,
      'Voucher / Ref No': q.referenceNo,
      'Account / Ledger Head': q.ledgerHead,
      'Party Name': q.partyName || '',
      'Amount (INR)': q.amount,
      'Discrepancy Category': q.discrepancyType.replace(/_/g, ' '),
      'Auditor Observation': q.auditObservation,
      'Requested Document / Evidence': q.requiredClientAction,
      'Client Management Response (To be filled by Client)': '',
      'Supporting Document Attached (Y/N)': '',
    }));

    const ws = XLSX.utils.json_to_sheet(excelData);
    XLSX.utils.book_append_sheet(wb, ws, 'Audit_Queries_Client_Action');

    return XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
  }
}
