import React from 'react';

interface UploadIllustrationProps {
  size?: 'sm' | 'md' | 'lg';
  className?: string;
  hasFile?: boolean;
}

/**
 * Universal "Real Type" File Upload SVG Component.
 * Implements the industry-standard Cloud Upload & Ingestion Vector
 * recognized across top SaaS platforms (Stripe, Dropbox, GitHub, Tailwind UI).
 */
export const UploadIllustration: React.FC<UploadIllustrationProps> = ({
  size = 'md',
  className = '',
  hasFile = false,
}) => {
  if (hasFile) {
    // When a file is already selected/active
    return (
      <div
        className={`inline-flex items-center justify-center shrink-0 rounded-xl bg-gray-100 border border-gray-200 text-gray-900 shadow-2xs ${
          size === 'sm' ? 'w-10 h-10' : size === 'lg' ? 'w-14 h-14' : 'w-12 h-12'
        } ${className}`}
      >
        <svg
          className={size === 'sm' ? 'w-5 h-5' : size === 'lg' ? 'w-7 h-7' : 'w-6 h-6'}
          viewBox="0 0 24 24"
          fill="none"
          stroke="currentColor"
          strokeWidth="2"
          strokeLinecap="round"
          strokeLinejoin="round"
        >
          <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" />
          <polyline points="14 2 14 8 20 8" />
          <path d="m9 15 2 2 4-4" />
        </svg>
      </div>
    );
  }

  if (size === 'sm') {
    return (
      <div
        className={`inline-flex items-center justify-center shrink-0 w-10 h-10 rounded-xl bg-gray-50 border border-gray-200 text-gray-700 shadow-2xs group-hover:bg-gray-100 group-hover:border-gray-300 transition-colors ${className}`}
      >
        <svg
          className="w-5 h-5 text-gray-600"
          viewBox="0 0 24 24"
          fill="none"
          stroke="currentColor"
          strokeWidth="1.8"
          strokeLinecap="round"
          strokeLinejoin="round"
        >
          <path d="M4 14.899A7 7 0 0 1 15.71 8h1.79a4.5 4.5 0 0 1 2.5 8.242" />
          <path d="M12 12v9" />
          <path d="m16 16-4-4-4 4" />
        </svg>
      </div>
    );
  }

  if (size === 'lg') {
    return (
      <div
        className={`relative inline-flex items-center justify-center shrink-0 w-16 h-16 rounded-2xl bg-gradient-to-b from-gray-50 to-gray-100 border border-gray-200 text-gray-800 shadow-xs ring-4 ring-gray-100/80 group-hover:scale-105 group-hover:border-gray-300 transition-all ${className}`}
      >
        <svg
          className="w-8 h-8 text-gray-700"
          viewBox="0 0 24 24"
          fill="none"
          stroke="currentColor"
          strokeWidth="1.8"
          strokeLinecap="round"
          strokeLinejoin="round"
        >
          {/* Cloud body */}
          <path
            d="M4 14.899A7 7 0 0 1 15.71 8h1.79a4.5 4.5 0 0 1 2.5 8.242"
            stroke="currentColor"
            fill="#F8FAFC"
          />
          {/* Stem & upward arrow */}
          <path d="M12 12v9" stroke="currentColor" strokeWidth="2" />
          <path d="m16 16-4-4-4 4" stroke="currentColor" strokeWidth="2" />
        </svg>
      </div>
    );
  }

  // Default 'md'
  return (
    <div
      className={`inline-flex items-center justify-center shrink-0 w-12 h-12 rounded-xl bg-gray-50 border border-gray-200 text-gray-700 shadow-2xs ring-2 ring-gray-100 group-hover:bg-gray-100 group-hover:border-gray-300 transition-all ${className}`}
    >
      <svg
        className="w-6 h-6 text-gray-700"
        viewBox="0 0 24 24"
        fill="none"
        stroke="currentColor"
        strokeWidth="1.8"
        strokeLinecap="round"
        strokeLinejoin="round"
      >
        <path
          d="M4 14.899A7 7 0 0 1 15.71 8h1.79a4.5 4.5 0 0 1 2.5 8.242"
          fill="#F8FAFC"
        />
        <path d="M12 12v9" strokeWidth="2" />
        <path d="m16 16-4-4-4 4" strokeWidth="2" />
      </svg>
    </div>
  );
};

