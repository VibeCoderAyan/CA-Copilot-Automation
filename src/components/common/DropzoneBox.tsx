import React, { useState, useRef } from 'react';
import { UploadCloud } from 'lucide-react';

interface DropzoneBoxProps {
  title: string;
  description: string;
  buttonLabel?: string;
  accept?: string;
  acceptedFormatsHint?: string;
  icon?: React.ComponentType<{ className?: string }>;
  onFileSelect: (file: File) => void;
  id?: string;
  className?: string;
}

export const DropzoneBox: React.FC<DropzoneBoxProps> = ({
  title,
  description,
  buttonLabel = 'Upload Spreadsheet File',
  accept = '.xlsx,.xls,.csv',
  acceptedFormatsHint = 'Supports Excel (.xlsx, .xls) and CSV (.csv)',
  icon: Icon = UploadCloud,
  onFileSelect,
  id = 'file-dropzone-input',
  className = '',
}) => {
  const [isDragging, setIsDragging] = useState(false);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const handleDragOver = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragging(true);
  };

  const handleDragLeave = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragging(false);
  };

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragging(false);

    const file = e.dataTransfer.files?.[0];
    if (file) {
      onFileSelect(file);
    }
  };

  const handleChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) {
      onFileSelect(file);
    }
    e.target.value = '';
  };

  return (
    <div
      onDragOver={handleDragOver}
      onDragLeave={handleDragLeave}
      onDrop={handleDrop}
      className={`bg-white border-2 border-dashed rounded-xl p-8 sm:p-10 text-center transition-all shadow-2xs ${
        isDragging
          ? 'border-gray-900 bg-gray-50/80 ring-2 ring-gray-900/10'
          : 'border-gray-200 hover:border-gray-300'
      } ${className}`}
    >
      <input
        type="file"
        ref={fileInputRef}
        onChange={handleChange}
        accept={accept}
        className="hidden"
        id={id}
      />

      <label htmlFor={id} className="cursor-pointer flex flex-col items-center">
        {/* Clean Monochrome Rounded Icon Container */}
        <div className="w-12 h-12 rounded-full bg-gray-100 border border-gray-200 flex items-center justify-center text-gray-900 mb-3 shadow-2xs transition-transform group-hover:scale-105">
          <Icon className="w-6 h-6 text-gray-800" />
        </div>

        {/* Title */}
        <h3 className="text-sm font-semibold text-gray-900">{title}</h3>

        {/* Description */}
        <p className="text-xs text-gray-500 mt-1 max-w-md mx-auto leading-relaxed">
          {description}
        </p>

        {/* Dark Action Pill Button */}
        <div className="mt-4 px-4 py-2 bg-gray-900 hover:bg-black text-white text-xs font-semibold rounded-lg shadow-2xs inline-flex items-center gap-2 transition-all active:scale-95">
          <UploadCloud className="w-4 h-4" />
          <span>{buttonLabel}</span>
        </div>

        {/* Formats Hint */}
        {acceptedFormatsHint && (
          <span className="text-[11px] text-gray-400 font-medium mt-3 block">
            {acceptedFormatsHint}
          </span>
        )}
      </label>
    </div>
  );
};
