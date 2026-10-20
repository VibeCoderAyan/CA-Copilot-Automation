import React from 'react';
import { Settings, ArrowLeft } from 'lucide-react';
import { BrandLogo } from './BrandLogo';

export interface HeaderProps {
  onOpenSettings: () => void;
  onGoHome?: () => void;
  activeNav?: string;
  selectedToolId?: string | null;
  selectedToolName?: string | null;
  // Kept for backward compatibility
  selectedCategory?: string;
  onSelectCategory?: (cat: any) => void;
  activeToolCategory?: string;
  categoryLabel?: string;
  searchQuery?: string;
  onSearchChange?: (query: string) => void;
  onOpenWorkflowMap?: () => void;
}

/**
 * Clean & Normal Header for CA Copilot
 */
export const Header: React.FC<HeaderProps> = ({
  onOpenSettings,
  onGoHome,
  activeNav = 'home',
  selectedToolId = null,
  selectedToolName = null,
}) => {
  return (
    <header className="h-12 bg-white border-b border-gray-200 flex items-center justify-between px-6 sticky top-0 z-40 shrink-0 select-none print:hidden">
      {/* 1. Header Left: Brand Mark or Active Tool Breadcrumb */}
      <div className="flex items-center gap-3">
        {selectedToolId ? (
          <div className="flex items-center gap-2">
            <button
              onClick={onGoHome}
              className="inline-flex items-center gap-1.5 px-2.5 py-1 rounded-md text-xs font-medium text-gray-700 hover:text-gray-900 bg-gray-100 hover:bg-gray-200 transition-colors cursor-pointer"
              title="Return to Workspace"
            >
              <ArrowLeft className="w-3.5 h-3.5" />
              <span>Back</span>
            </button>
            <div className="w-[1px] h-3.5 bg-gray-300" />
            <span className="font-semibold text-xs text-gray-900 truncate max-w-[240px]">
              {selectedToolName || selectedToolId}
            </span>
          </div>
        ) : (
          <div
            onClick={onGoHome}
            className="flex items-center gap-2.5 cursor-pointer hover:opacity-85 transition-opacity"
            title="CA Copilot"
          >
            <div className="w-6 h-6 flex items-center justify-center shrink-0">
              <BrandLogo className="w-5.5 h-5.5" />
            </div>
            <span className="font-semibold text-sm text-gray-900 tracking-tight">CA Copilot</span>
          </div>
        )}
      </div>

      {/* 2. Header Right: Settings */}
      <div className="flex items-center gap-2">
        <button
          onClick={onOpenSettings}
          className={`w-8 h-8 rounded-full flex items-center justify-center transition-colors cursor-pointer ${
            activeNav === 'settings'
              ? 'bg-gray-200 text-gray-900'
              : 'text-gray-500 hover:text-gray-900 hover:bg-gray-100'
          }`}
          title="Settings"
        >
          <Settings className="w-4 h-4" />
        </button>
      </div>
    </header>
  );
};
