import React from 'react';
import { FileWarning, ArrowRight, X, CheckCircle2, FileSpreadsheet } from 'lucide-react';

export interface UnsupportedFileWarningProps {
  toolName: string;
  fileName?: string;
  errorMessage: string;
  reason?: string;
  requiredFormat: string;
  guidelines: string[];
  onDismiss: () => void;
  onRetryUpload?: () => void;
}

export const UnsupportedFileWarning: React.FC<UnsupportedFileWarningProps> = ({
  toolName,
  fileName,
  errorMessage,
  reason,
  requiredFormat,
  guidelines,
  onDismiss,
  onRetryUpload,
}) => {
  return (
    <div
      className="bg-white/95 backdrop-blur-sm border border-rose-200/90 rounded-2xl p-4 sm:p-5 shadow-[0_12px_28px_-6px_rgba(225,29,72,0.10),0_2px_8px_-2px_rgba(15,23,42,0.04)] relative overflow-hidden animate-in fade-in slide-in-from-top-2 duration-200"
      role="alert"
    >
      {/* Top Refined Accent Line */}
      <div className="absolute top-0 left-0 right-0 h-1 bg-gradient-to-r from-rose-500 via-rose-400 to-amber-400" />

      <div className="flex items-start gap-3.5">
        {/* Soft Crimson Icon Container */}
        <div className="w-9 h-9 rounded-xl bg-rose-50 border border-rose-200/80 flex items-center justify-center shrink-0 shadow-2xs mt-0.5">
          <FileWarning className="w-4 h-4 text-rose-600 stroke-[2.2]" />
        </div>

        <div className="flex-1 min-w-0">
          {/* Header & Status */}
          <div className="flex items-center justify-between gap-2">
            <div className="flex items-center gap-2 flex-wrap">
              <span className="text-[9.5px] font-bold uppercase tracking-wider px-2 py-0.5 rounded-full bg-rose-50 text-rose-700 border border-rose-200/80">
                FORMAT MISMATCH
              </span>
              <h3 className="text-xs sm:text-sm font-bold text-slate-900 tracking-tight">
                File Ingestion Halted for {toolName}
              </h3>
            </div>
            <button
              onClick={onDismiss}
              className="w-6 h-6 rounded-md hover:bg-slate-100 flex items-center justify-center text-slate-400 hover:text-slate-700 transition-colors shrink-0"
              aria-label="Dismiss warning"
              title="Dismiss warning"
            >
              <X className="w-3.5 h-3.5" />
            </button>
          </div>

          {/* Rejection Details Box */}
          <div className="mt-2.5 p-3 rounded-xl bg-rose-50/50 border border-rose-100 text-xs text-rose-950">
            {fileName && (
              <p className="font-semibold text-slate-900 flex items-center gap-1.5 flex-wrap">
                <FileSpreadsheet className="w-3.5 h-3.5 text-rose-600 shrink-0" />
                <span>Uploaded:</span>
                <span className="font-mono text-[11px] font-semibold bg-white border border-rose-200/80 px-1.5 py-0.5 rounded text-rose-950 shadow-2xs">
                  {fileName}
                </span>
              </p>
            )}
            <p className="mt-1 text-slate-700 leading-relaxed font-normal">
              {errorMessage}
            </p>
            {reason && (
              <p className="mt-1.5 text-rose-900 text-[11px] font-medium bg-rose-100/60 rounded px-2 py-1">
                <strong>Detection note:</strong> {reason}
              </p>
            )}
          </div>

          {/* Guidelines Box */}
          <div className="mt-3 pt-3 border-t border-slate-100">
            <h4 className="text-[11px] font-bold text-slate-800 uppercase tracking-wide">
              Required Format & Expected Structure
            </h4>
            <p className="text-[11.5px] text-slate-500 mt-0.5">
              Accepted format: <strong className="text-slate-800 font-semibold">{requiredFormat}</strong>
            </p>

            {guidelines.length > 0 && (
              <ul className="mt-2 space-y-1">
                {guidelines.map((instruction, idx) => (
                  <li key={idx} className="flex items-start gap-1.5 text-[11.5px] text-slate-600">
                    <CheckCircle2 className="w-3.5 h-3.5 text-gray-700 mt-0.5 shrink-0" />
                    <span className="leading-snug">{instruction}</span>
                  </li>
                ))}
              </ul>
            )}
          </div>

          {/* Action CTAs */}
          <div className="mt-3.5 pt-2.5 border-t border-slate-100 flex items-center justify-between gap-3 flex-wrap">
            <p className="text-[10.5px] text-slate-400 italic">
              Execution paused to protect your general ledger & reconciliation books from corrupt entries.
            </p>

            <div className="flex items-center gap-2">
              {onRetryUpload && (
                <button
                  onClick={onRetryUpload}
                  className="px-3 py-1.5 text-xs font-semibold rounded-lg border border-slate-200 bg-white hover:bg-slate-50 text-slate-700 hover:text-slate-900 transition-colors shadow-2xs cursor-pointer"
                >
                  Upload Different File
                </button>
              )}
            </div>
          </div>
        </div>
      </div>
    </div>
  );
};
