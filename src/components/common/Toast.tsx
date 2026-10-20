import React, { useEffect, useState, useRef } from 'react';
import { CheckCircle2, AlertCircle, Info, X, FileSpreadsheet } from 'lucide-react';

export interface ToastMessage {
  id: string;
  type: 'success' | 'error' | 'info';
  title: string;
  description?: string;
}

interface ToastContainerProps {
  toasts: ToastMessage[];
  onDismiss: (id: string) => void;
}

export const ToastContainer: React.FC<ToastContainerProps> = ({ toasts, onDismiss }) => {
  if (toasts.length === 0) return null;

  return (
    <div
      className="fixed top-5 right-5 z-[9999] flex flex-col items-end gap-2 max-w-sm w-full px-4 pointer-events-none no-print print:hidden"
      aria-live="polite"
    >
      {toasts.map((toast) => (
        <ToastItem key={toast.id} toast={toast} onDismiss={onDismiss} />
      ))}
    </div>
  );
};

const ToastItem: React.FC<{ toast: ToastMessage; onDismiss: (id: string) => void }> = ({
  toast,
  onDismiss,
}) => {
  const isError =
    toast.type === 'error' ||
    /invalid|halted|failed|error|unsupported|corrupted/i.test(toast.title);

  // Error stays a bit longer so user can read
  const duration = isError ? 4500 : 3000;

  useEffect(() => {
    const timer = setTimeout(() => {
      onDismiss(toast.id);
    }, duration);
    return () => clearTimeout(timer);
  }, [toast.id, duration, onDismiss]);

  return (
    <div
      className="pointer-events-auto w-full bg-white border border-gray-200 rounded-lg shadow-lg shadow-black/5 p-3 transition-all duration-150 animate-in fade-in slide-in-from-top-2 flex items-start gap-2.5"
      role="alert"
    >
      {/* Clean, standard non-AI icon */}
      <div className="shrink-0 mt-0.5">
        {isError ? (
          <AlertCircle className="w-4 h-4 text-red-600 stroke-[2]" />
        ) : toast.type === 'info' ? (
          <Info className="w-4 h-4 text-gray-500 stroke-[2]" />
        ) : (
          <CheckCircle2 className="w-4 h-4 text-gray-900 stroke-[2]" />
        )}
      </div>

      {/* Message Text */}
      <div className="flex-1 min-w-0 pr-1">
        <h4 className="text-xs font-semibold text-gray-900 leading-snug">
          {toast.title}
        </h4>
        {toast.description && (
          <p className="text-[11px] text-gray-500 font-normal leading-relaxed mt-0.5 break-words">
            {toast.description}
          </p>
        )}
      </div>

      {/* Minimal dismiss button */}
      <button
        onClick={() => onDismiss(toast.id)}
        className="w-5 h-5 rounded hover:bg-gray-100 text-gray-400 hover:text-gray-700 flex items-center justify-center transition-colors shrink-0 -mr-1 -mt-0.5 cursor-pointer"
        aria-label="Dismiss notification"
      >
        <X className="w-3.5 h-3.5" />
      </button>
    </div>
  );
};
