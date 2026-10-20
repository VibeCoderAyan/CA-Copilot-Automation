import React from 'react';

export interface AuditorSignOffFooterProps {
  preparedBy?: string;
  reviewedBy?: string;
  partnerName?: string;
  udin?: string;
  classification?: string;
}

export const AuditorSignOffFooter: React.FC<AuditorSignOffFooterProps> = ({
  preparedBy = 'Audit Associate / CA Assistant',
  reviewedBy = 'Audit Manager / Senior Reviewer',
  partnerName = 'Engagement Partner / Statutory Auditor',
  udin,
  classification = 'STRICTLY CONFIDENTIAL — STATUTORY AUDIT WORKING PAPER PRESERVED UNDER ICAI SA-230',
}) => {
  return (
    <div className="hidden print:block mt-8 pt-4 border-t-2 border-gray-900 break-inside-avoid page-break-inside-avoid">
      <div className="grid grid-cols-3 gap-4 text-[10px] text-gray-800">
        <div className="border border-gray-300 p-2.5 rounded bg-white">
          <div className="font-bold text-gray-900 mb-6 uppercase tracking-wider text-[9px]">
            1. Prepared By ({preparedBy}):
          </div>
          <div className="border-b border-gray-400 mb-1.5"></div>
          <div className="flex justify-between text-[8.5px] text-gray-500 font-medium">
            <span>Signature: ________________</span>
            <span>Date: ___/___/2026</span>
          </div>
        </div>

        <div className="border border-gray-300 p-2.5 rounded bg-white">
          <div className="font-bold text-gray-900 mb-6 uppercase tracking-wider text-[9px]">
            2. Reviewed By ({reviewedBy}):
          </div>
          <div className="border-b border-gray-400 mb-1.5"></div>
          <div className="flex justify-between text-[8.5px] text-gray-500 font-medium">
            <span>Signature: ________________</span>
            <span>Date: ___/___/2026</span>
          </div>
        </div>

        <div className="border border-gray-300 p-2.5 rounded bg-white">
          <div className="font-bold text-gray-900 mb-6 uppercase tracking-wider text-[9px]">
            3. Final Audit Sign-Off ({partnerName}):
          </div>
          <div className="border-b border-gray-400 mb-1.5"></div>
          <div className="flex justify-between text-[8.5px] text-gray-500 font-medium">
            <span>M. No: ____________</span>
            <span>UDIN: {udin || '_________________'}</span>
          </div>
        </div>
      </div>
      <div className="text-center text-[8.5px] text-gray-500 mt-2.5 font-medium tracking-wide">
        {classification}
      </div>
    </div>
  );
};
