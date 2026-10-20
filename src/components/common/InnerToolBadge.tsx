import React from 'react';
import { ToolIcon } from './ToolIcon';

interface InnerToolBadgeProps {
  toolId?: string;
  label?: string;
  className?: string;
}

/**
 * Dedicated Inner Workspace Badge for CA Automation Tools.
 * Seamlessly matches the outer catalog cards with vibrant, harmonic Figma-grade iconography.
 */
export const InnerToolBadge: React.FC<InnerToolBadgeProps> = ({ toolId, label, className = '' }) => {
  if (label && !toolId) {
    return (
      <span className={`px-2.5 py-0.5 text-[11px] font-semibold rounded-full bg-gray-100 text-gray-700 border border-gray-200 inline-flex items-center ${className}`}>
        {label}
      </span>
    );
  }

  if (toolId) {
    return <ToolIcon toolId={toolId} size="md" className={className} />;
  }

  return null;
};
