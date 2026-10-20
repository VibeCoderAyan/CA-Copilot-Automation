import React, { useState } from 'react';
import { HelpCircle, ChevronDown, ChevronUp, Download, CheckCircle2, FileSpreadsheet } from 'lucide-react';
import { downloadTemplate } from '../../utils/sampleTemplates';

interface DataFormatGuideProps {
  toolType: 'excel-cleaner' | 'bank-rec' | 'gst-rec' | 'trial-balance' | 'audit' | 'pdf' | 'invoices' | 'bulk-renamer';
  acceptedFormats: string;
  requiredFields: { name: string; description: string; aliases: string[] }[];
  notes?: string[];
}

export const DataFormatGuide: React.FC<DataFormatGuideProps> = ({
  toolType,
  acceptedFormats,
  requiredFields,
  notes = [],
}) => {
  const [isOpen, setIsOpen] = useState(false);

  const canDownloadTemplate = ['excel-cleaner', 'bank-rec', 'gst-rec', 'trial-balance', 'audit'].includes(toolType);

  const handleDownloadTemplate = (e: React.MouseEvent) => {
    e.stopPropagation();
    if (canDownloadTemplate) {
      downloadTemplate(toolType as any);
    }
  };

  return (
    <div className="bg-white border border-[#E5E7EB] rounded-xl overflow-hidden shadow-sm transition-all no-print print:hidden">
      <div
        onClick={() => setIsOpen(!isOpen)}
        className="px-4 py-3 flex items-center justify-between cursor-pointer hover:bg-gray-50/70 transition-colors select-none"
      >
        <div className="flex items-center gap-2.5">
          <div className="w-7 h-7 rounded-lg bg-gray-100 flex items-center justify-center text-gray-700">
            <HelpCircle className="w-4 h-4" />
          </div>
          <div>
            <div className="flex items-center gap-2 flex-wrap">
              <span className="text-xs font-bold text-gray-900">File Format & Column Requirements</span>
              <span className="text-[10px] font-semibold bg-gray-100 text-gray-700 px-2 py-0.5 rounded border border-gray-200">
                {acceptedFormats}
              </span>
            </div>
            <p className="text-[11px] text-gray-500 mt-0.5">
              Click to view recognized column aliases and template requirements
            </p>
          </div>
        </div>

        <div className="flex items-center gap-2">
          {canDownloadTemplate && (
            <button
              type="button"
              onClick={handleDownloadTemplate}
              className="flex items-center gap-1.5 px-2.5 py-1 text-xs font-medium text-gray-700 bg-white hover:bg-gray-50 border border-gray-300 rounded-md shadow-2xs transition-colors cursor-pointer"
              title="Download standard Excel sample template"
            >
              <Download className="w-3.5 h-3.5 text-gray-600" />
              <span>Sample Template</span>
            </button>
          )}
          <button
            type="button"
            className="p-1 text-gray-500 hover:text-gray-700 rounded transition-colors cursor-pointer"
            aria-label="Toggle Guide"
          >
            {isOpen ? <ChevronUp className="w-4 h-4" /> : <ChevronDown className="w-4 h-4" />}
          </button>
        </div>
      </div>

      {isOpen && (
        <div className="px-5 py-4 border-t border-gray-200 bg-gray-50/50 space-y-4 text-xs">
          {/* Workflow Steps Indicator */}
          <div className="grid grid-cols-2 sm:grid-cols-4 gap-2 pt-1 pb-2">
            <div className="p-2.5 rounded-lg bg-white border border-gray-200 flex items-center gap-2 shadow-2xs">
              <span className="w-5 h-5 rounded-full bg-gray-900 text-white text-[10px] font-bold flex items-center justify-center">1</span>
              <span className="text-[11px] font-medium text-gray-800">Upload & Validate</span>
            </div>
            <div className="p-2.5 rounded-lg bg-white border border-gray-200 flex items-center gap-2 shadow-2xs">
              <span className="w-5 h-5 rounded-full bg-gray-900 text-white text-[10px] font-bold flex items-center justify-center">2</span>
              <span className="text-[11px] font-medium text-gray-800">Auto-Map Columns</span>
            </div>
            <div className="p-2.5 rounded-lg bg-white border border-gray-200 flex items-center gap-2 shadow-2xs">
              <span className="w-5 h-5 rounded-full bg-gray-900 text-white text-[10px] font-bold flex items-center justify-center">3</span>
              <span className="text-[11px] font-medium text-gray-800">Preview Results</span>
            </div>
            <div className="p-2.5 rounded-lg bg-white border border-gray-200 flex items-center gap-2 shadow-2xs">
              <span className="w-5 h-5 rounded-full bg-gray-900 text-white text-[10px] font-bold flex items-center justify-center">4</span>
              <span className="text-[11px] font-medium text-gray-800">Export Excel Report</span>
            </div>
          </div>

          {/* Required Fields Table */}
          <div>
            <span className="text-[11px] font-bold text-gray-900 uppercase tracking-wider block mb-2">
              Recognized Columns & Common Variations
            </span>
            <div className="border border-gray-200 rounded-lg overflow-hidden bg-white shadow-2xs">
              <table className="w-full text-left text-xs">
                <thead className="bg-gray-50 border-b border-gray-200 text-gray-600 text-[11px] font-semibold">
                  <tr>
                    <th className="py-2 px-3 border-r border-gray-200 w-36">Field Name</th>
                    <th className="py-2 px-3 border-r border-gray-200 w-52">Description</th>
                    <th className="py-2 px-3">Supported Column Aliases (Auto-Detected)</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-gray-200 text-gray-800">
                  {requiredFields.map((f, i) => (
                    <tr key={i} className="hover:bg-gray-50/80">
                      <td className="py-2 px-3 border-r border-gray-200 font-semibold text-gray-900 font-mono text-[11px]">
                        {f.name}
                      </td>
                      <td className="py-2 px-3 border-r border-gray-200 text-gray-600 text-[11px]">
                        {f.description}
                      </td>
                      <td className="py-2 px-3 text-[11px] font-mono text-gray-600">
                        {f.aliases.join(', ')}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>

          {/* Specific Notes */}
          {notes.length > 0 && (
            <div className="space-y-1.5 pt-1">
              <span className="text-[11px] font-bold text-[#111827] block">Important Processing Notes:</span>
              <ul className="list-disc list-inside space-y-1 text-[11px] text-[#6B7280]">
                {notes.map((n, i) => (
                  <li key={i}>{n}</li>
                ))}
              </ul>
            </div>
          )}
        </div>
      )}
    </div>
  );
};
