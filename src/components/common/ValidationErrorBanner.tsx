import React from 'react';
import { AlertCircle, X, FileSpreadsheet } from 'lucide-react';

export interface ValidationErrorBannerProps {
  message: string;
  onDismiss: () => void;
  title?: string;
  className?: string;
}

/**
 * Highlights single or double quoted strings as monospace file chips
 */
function renderMessageWithChips(msg: string) {
  const quoteRegex = /(['"])(.*?)\1/g;
  const parts: React.ReactNode[] = [];
  let lastIndex = 0;
  let match: RegExpExecArray | null;

  while ((match = quoteRegex.exec(msg)) !== null) {
    if (match.index > lastIndex) {
      parts.push(msg.substring(lastIndex, match.index));
    }
    const filename = match[2];
    parts.push(
      <span
        key={`banner-chip-${match.index}`}
        className="inline-flex items-center gap-1 font-mono text-[11px] font-semibold text-rose-950 bg-white border border-rose-200/90 rounded px-1.5 py-0.5 mx-1 shadow-2xs align-middle"
      >
        <FileSpreadsheet className="w-3 h-3 text-rose-500 shrink-0" />
        <span className="truncate max-w-[220px]">{filename}</span>
      </span>
    );
    lastIndex = match.index + match[0].length;
  }

  if (lastIndex < msg.length) {
    parts.push(msg.substring(lastIndex));
  }

  return parts.length > 0 ? parts : msg;
}

export const ValidationErrorBanner: React.FC<ValidationErrorBannerProps> = ({
  message,
  onDismiss,
  title = 'VALIDATION NOTICE',
  className = '',
}) => {
  return (
    <div
      className={`p-3 sm:p-3.5 rounded-xl bg-rose-50/70 border border-rose-200/90 shadow-2xs flex items-center justify-between gap-3 text-xs text-rose-950 animate-in fade-in slide-in-from-top-1 duration-150 print:hidden ${className}`}
      role="alert"
    >
      <div className="flex items-center gap-2.5 min-w-0">
        <div className="w-7 h-7 rounded-lg bg-rose-100 border border-rose-200/80 text-rose-600 flex items-center justify-center shrink-0">
          <AlertCircle className="w-4 h-4 stroke-[2.2]" />
        </div>
        <div className="min-w-0 flex-1 leading-snug">
          <span className="text-[9.5px] font-bold uppercase tracking-wider text-rose-700 bg-white/90 border border-rose-200/80 px-1.5 py-0.5 rounded-full mr-2 inline-block">
            {title}
          </span>
          <span className="text-slate-800 text-[11.5px] font-normal align-middle">
            {renderMessageWithChips(message)}
          </span>
        </div>
      </div>

      <button
        onClick={onDismiss}
        className="inline-flex items-center gap-1 px-2.5 py-1 text-[11px] font-semibold text-rose-700 hover:text-rose-900 bg-white/90 hover:bg-white border border-rose-200 rounded-md transition-all shadow-2xs shrink-0 cursor-pointer"
        aria-label="Dismiss error"
      >
        <X className="w-3 h-3" />
        <span>Dismiss</span>
      </button>
    </div>
  );
};
