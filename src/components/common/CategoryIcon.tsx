import React from 'react';
import type { SidebarCategory } from './Sidebar';

export type CategoryIconType = SidebarCategory | 'workflow';

interface CategoryIconProps {
  type: CategoryIconType;
  className?: string;
  size?: number;
}

/**
 * Hand-Crafted Multi-Colored Vector SVGs for CA Copilot Navigation
 * - Custom illustrated vector graphics (NOT generic line icons inside boxes)
 * - Rich gradients and dimensional layers
 * - Scalable, sharp, and lightweight
 */
export const CategoryIcon: React.FC<CategoryIconProps> = ({
  type,
  className = 'w-5 h-5',
  size = 20,
}) => {
  const style = size ? { width: size, height: size } : undefined;

  // 1. Workflow Map: Connected Automated Pipeline
  if (type === 'workflow') {
    return (
      <svg
        viewBox="0 0 24 24"
        fill="none"
        xmlns="http://www.w3.org/2000/svg"
        className={className}
        style={style}
      >
        <defs>
          <linearGradient id="wfPath" x1="4" y1="4" x2="20" y2="20" gradientUnits="userSpaceOnUse">
            <stop offset="0%" stopColor="#6366F1" />
            <stop offset="50%" stopColor="#8B5CF6" />
            <stop offset="100%" stopColor="#06B6D4" />
          </linearGradient>
        </defs>
        {/* Connecting Flow Paths */}
        <path
          d="M6 7C10 7 10 17 14 17H18"
          stroke="url(#wfPath)"
          strokeWidth="2.2"
          strokeLinecap="round"
        />
        <path
          d="M6 7C6 12 14 12 18 12"
          stroke="url(#wfPath)"
          strokeWidth="1.6"
          strokeLinecap="round"
          strokeDasharray="2 3"
          opacity="0.7"
        />
        {/* Interactive Nodes */}
        <circle cx="5" cy="7" r="3" fill="#6366F1" />
        <circle cx="5" cy="7" r="1.3" fill="#FFFFFF" />

        <circle cx="18" cy="12" r="2.5" fill="#8B5CF6" />
        <circle cx="18" cy="12" r="1" fill="#FFFFFF" />

        <circle cx="18" cy="17" r="3" fill="#06B6D4" />
        <circle cx="18" cy="17" r="1.3" fill="#FFFFFF" />
      </svg>
    );
  }

  // 2. All Automations: 4-Quadrant Modern App Matrix
  if (type === 'all') {
    return (
      <svg
        viewBox="0 0 24 24"
        fill="none"
        xmlns="http://www.w3.org/2000/svg"
        className={className}
        style={style}
      >
        <rect x="3" y="3" width="7.5" height="7.5" rx="2.5" fill="#2563EB" />
        <rect x="13.5" y="3" width="7.5" height="7.5" rx="2.5" fill="#06B6D4" />
        <rect x="3" y="13.5" width="7.5" height="7.5" rx="2.5" fill="#6366F1" />
        <rect x="13.5" y="13.5" width="7.5" height="7.5" rx="2.5" fill="#10B981" />
        {/* Inner subtle core accents */}
        <circle cx="6.75" cy="6.75" r="1" fill="#FFFFFF" opacity="0.8" />
        <circle cx="17.25" cy="6.75" r="1" fill="#FFFFFF" opacity="0.8" />
        <circle cx="6.75" cy="17.25" r="1" fill="#FFFFFF" opacity="0.8" />
        <circle cx="17.25" cy="17.25" r="1" fill="#FFFFFF" opacity="0.8" />
      </svg>
    );
  }

  // 3. Reconciliations: Dynamic Inter-matching Balance Arrows
  if (type === 'reconciliation') {
    return (
      <svg
        viewBox="0 0 24 24"
        fill="none"
        xmlns="http://www.w3.org/2000/svg"
        className={className}
        style={style}
      >
        <defs>
          <linearGradient id="recTop" x1="3" y1="8" x2="21" y2="8" gradientUnits="userSpaceOnUse">
            <stop offset="0%" stopColor="#2563EB" />
            <stop offset="100%" stopColor="#06B6D4" />
          </linearGradient>
          <linearGradient id="recBottom" x1="21" y1="16" x2="3" y2="16" gradientUnits="userSpaceOnUse">
            <stop offset="0%" stopColor="#4F46E5" />
            <stop offset="100%" stopColor="#3B82F6" />
          </linearGradient>
        </defs>
        {/* Top Forward Arrow */}
        <path
          d="M4 8.5H19M19 8.5L14.5 4M19 8.5L14.5 13"
          stroke="url(#recTop)"
          strokeWidth="2.3"
          strokeLinecap="round"
          strokeLinejoin="round"
        />
        {/* Bottom Reverse Return Arrow */}
        <path
          d="M20 15.5H5M5 15.5L9.5 11M5 15.5L9.5 20"
          stroke="url(#recBottom)"
          strokeWidth="2.3"
          strokeLinecap="round"
          strokeLinejoin="round"
        />
      </svg>
    );
  }

  // 4. Audit & Scrutiny: Dimensional Verification Shield
  if (type === 'audit') {
    return (
      <svg
        viewBox="0 0 24 24"
        fill="none"
        xmlns="http://www.w3.org/2000/svg"
        className={className}
        style={style}
      >
        <defs>
          <linearGradient id="auditShield" x1="12" y1="2" x2="12" y2="22" gradientUnits="userSpaceOnUse">
            <stop offset="0%" stopColor="#10B981" />
            <stop offset="100%" stopColor="#047857" />
          </linearGradient>
        </defs>
        {/* Shield Body */}
        <path
          d="M12 2L4 5.5V11C4 16.5 7.5 20.5 12 22C16.5 20.5 20 16.5 20 11V5.5L12 2Z"
          fill="url(#auditShield)"
        />
        {/* Crisp White Checkmark */}
        <path
          d="M8.5 11.5L11 14L15.5 9.5"
          stroke="#FFFFFF"
          strokeWidth="2.2"
          strokeLinecap="round"
          strokeLinejoin="round"
        />
      </svg>
    );
  }

  // 5. Tax & Invoices: Statutory Invoice Document
  if (type === 'tax') {
    return (
      <svg
        viewBox="0 0 24 24"
        fill="none"
        xmlns="http://www.w3.org/2000/svg"
        className={className}
        style={style}
      >
        <defs>
          <linearGradient id="taxGrad" x1="4" y1="3" x2="20" y2="21" gradientUnits="userSpaceOnUse">
            <stop offset="0%" stopColor="#F59E0B" />
            <stop offset="100%" stopColor="#EA580C" />
          </linearGradient>
        </defs>
        {/* Invoice Page Base with Folded Corner */}
        <path
          d="M5 4C5 2.89543 5.89543 2 7 2H14.5L19 6.5V20C19 21.1046 18.1046 22 17 22H7C5.89543 22 5 21.1046 5 20V4Z"
          fill="url(#taxGrad)"
        />
        {/* Corner Fold */}
        <path d="M14 2V7H19" fill="#FDE68A" opacity="0.6" />
        {/* Tax Rows */}
        <line x1="8" y1="10" x2="16" y2="10" stroke="#FFFFFF" strokeWidth="1.8" strokeLinecap="round" />
        <line x1="8" y1="13.5" x2="16" y2="13.5" stroke="#FFFFFF" strokeWidth="1.8" strokeLinecap="round" />
        <line x1="8" y1="17" x2="12" y2="17" stroke="#FFFFFF" strokeWidth="1.8" strokeLinecap="round" />
      </svg>
    );
  }

  // 6. Excel Utilities: Authentic Spreadsheet Grid
  return (
    <svg
      viewBox="0 0 24 24"
      fill="none"
      xmlns="http://www.w3.org/2000/svg"
      className={className}
      style={style}
    >
      <defs>
        <linearGradient id="excelGrad" x1="3" y1="3" x2="21" y2="21" gradientUnits="userSpaceOnUse">
          <stop offset="0%" stopColor="#10B981" />
          <stop offset="100%" stopColor="#059669" />
        </linearGradient>
      </defs>
      {/* Workbook Body */}
      <rect x="3" y="3" width="18" height="18" rx="3.5" fill="url(#excelGrad)" />
      {/* Grid Lines */}
      <line x1="3" y1="9" x2="21" y2="9" stroke="#FFFFFF" strokeWidth="1.3" opacity="0.75" />
      <line x1="3" y1="15" x2="21" y2="15" stroke="#FFFFFF" strokeWidth="1.3" opacity="0.75" />
      <line x1="9" y1="3" x2="9" y2="21" stroke="#FFFFFF" strokeWidth="1.3" opacity="0.75" />
      {/* Top Header highlight cell */}
      <rect x="4" y="4" width="4" height="4" rx="1" fill="#FFFFFF" opacity="0.9" />
    </svg>
  );
};
