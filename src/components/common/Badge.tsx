import React from 'react';

export type BadgeVariant = 'success' | 'error' | 'warning' | 'navy' | 'neutral';

interface BadgeProps {
  variant?: BadgeVariant;
  children: React.ReactNode;
  size?: 'sm' | 'md';
  icon?: React.ReactNode;
  className?: string;
}

export const Badge: React.FC<BadgeProps> = ({
  variant = 'neutral',
  children,
  size = 'sm',
  icon,
  className = '',
}) => {
  const variantStyles: Record<BadgeVariant, string> = {
    // Clean subtle success:
    success: 'bg-gray-100 text-gray-900 border-gray-200 font-semibold',
    // Red: exclusively for errors and failed checks
    error: 'bg-rose-50 text-rose-700 border-rose-200 font-semibold',
    // Amber for warnings
    warning: 'bg-amber-50 text-amber-800 border-amber-200 font-semibold',
    // Navy / Slate
    navy: 'bg-slate-50 text-slate-700 border-slate-200 font-semibold',
    // Neutral
    neutral: 'bg-gray-100 text-gray-700 border-gray-200 font-medium',
  };

  const sizeStyles = {
    sm: 'text-[11px] px-2 py-0.5 font-medium',
    md: 'text-xs px-2.5 py-1 font-medium',
  };

  return (
    <span
      className={`inline-flex items-center gap-1.5 rounded-md border ${variantStyles[variant]} ${sizeStyles[size]} ${className}`}
    >
      {icon && <span className="inline-block">{icon}</span>}
      {children}
    </span>
  );
};
