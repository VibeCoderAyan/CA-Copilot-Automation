import React from 'react';
import { Settings, Lock } from 'lucide-react';
import { BrandLogo } from './BrandLogo';
import { CategoryIcon } from './CategoryIcon';

export type NavItem = 'home' | 'all-tools' | 'settings' | 'working-flow';
export type SidebarCategory = 'all' | 'reconciliation' | 'audit' | 'tax' | 'excel';

export interface SidebarProps {
  activeNav: NavItem;
  selectedCategory?: SidebarCategory;
  selectedTool?: string | null;
  onSelectNav: (item: NavItem) => void;
  onSelectCategory?: (cat: SidebarCategory) => void;
  onSelectTool?: (toolId: string) => void;
  isPinActive?: boolean;
  onLockApp?: () => void;
}

interface NavSection {
  id: SidebarCategory;
  label: string;
}

const SECTIONS: NavSection[] = [
  { id: 'all', label: 'All Automations' },
  { id: 'reconciliation', label: 'Reconciliations' },
  { id: 'audit', label: 'Audit & Scrutiny' },
  { id: 'tax', label: 'Tax & Invoices' },
  { id: 'excel', label: 'Excel Utilities' },
];

/**
 * Pixel-Perfect Symmetrical Apple Sidebar with:
 * - Brand Logo laser-aligned at 20px
 * - Practice Tool Categories
 * - Pinned Settings & App Lock
 */
export const Sidebar: React.FC<SidebarProps> = ({
  activeNav,
  selectedCategory = 'all',
  selectedTool = null,
  onSelectNav,
  onSelectCategory,
  isPinActive,
  onLockApp,
}) => {
  const isWorkflowActive = activeNav === 'working-flow' && selectedTool === null;
  const isSettingsActive = activeNav === 'settings' && selectedTool === null;

  return (
    <aside className="w-52 bg-white border-r border-gray-200/80 flex flex-col h-full select-none shrink-0 z-20 print:hidden">
      {/* 1. Brand Header with Laser-Aligned Symmetry & Drag Region */}
      <div
        className="h-14 px-5 border-b border-gray-100 flex items-center shrink-0 select-none app-drag-region"
      >
        <button
          onClick={() => {
            onSelectNav('home');
            if (onSelectCategory) onSelectCategory('all');
          }}
          className="flex items-center gap-2.5 cursor-pointer hover:opacity-90 transition-opacity group select-none app-no-drag focus:outline-none"
          title="CA Copilot — Home"
        >
          <BrandLogo className="w-7 h-7 shrink-0 group-hover:scale-105 transition-transform duration-200" />
          <span className="font-bold text-[15px] text-gray-900 tracking-tight leading-none select-none">
            CA Copilot
          </span>
        </button>
      </div>

      {/* 2. Middle Navigation Sections */}
      <nav className="flex-1 px-2.5 py-3 space-y-1 overflow-y-auto">
        <div className="px-2.5 pb-1.5 text-[10px] font-semibold tracking-wider text-gray-400 uppercase">
          Workflow
        </div>

        {/* Workflow Map Button */}
        <button
          onClick={() => onSelectNav('working-flow')}
          className={`w-full flex items-center gap-2.5 px-2.5 py-2 rounded-lg text-xs font-medium transition-all active:scale-[0.98] cursor-pointer group ${
            isWorkflowActive
              ? 'bg-gray-100/90 text-gray-900 font-semibold'
              : 'text-gray-700 hover:bg-gray-50 hover:text-gray-900'
          }`}
        >
          <CategoryIcon type="workflow" size={18} className="shrink-0 group-hover:scale-105 transition-transform" />
          <span className="truncate">Workflow Map</span>
        </button>

        <div className="pt-3 pb-1 px-2.5 text-[10px] font-semibold tracking-wider text-gray-400 uppercase">
          Tool Categories
        </div>

        {SECTIONS.map((section) => {
          const isActive =
            activeNav === 'home' &&
            selectedCategory === section.id &&
            selectedTool === null;

          return (
            <button
              key={section.id}
              onClick={() => {
                onSelectNav('home');
                if (onSelectCategory) onSelectCategory(section.id);
              }}
              className={`w-full flex items-center gap-2.5 px-2.5 py-2 rounded-lg text-xs font-medium transition-all active:scale-[0.98] group cursor-pointer ${
                isActive
                  ? 'bg-gray-100/90 text-gray-900 font-semibold'
                  : 'text-gray-700 hover:bg-gray-50 hover:text-gray-900'
              }`}
            >
              <CategoryIcon type={section.id} size={18} className="shrink-0 group-hover:scale-105 transition-transform" />
              <span className="truncate">{section.label}</span>
            </button>
          );
        })}
      </nav>

      {/* 4. Settings Pinned to the Bottom */}
      <div className="p-2.5 border-t border-gray-100 shrink-0 flex items-center gap-1.5">
        <button
          onClick={() => onSelectNav('settings')}
          className={`flex-1 flex items-center gap-2.5 px-2.5 py-2 rounded-lg text-xs font-medium transition-all active:scale-[0.98] group cursor-pointer ${
            isSettingsActive
              ? 'bg-gray-100 text-gray-900 font-semibold'
              : 'text-gray-700 hover:bg-gray-50 hover:text-gray-900'
          }`}
          title="Practice Settings"
        >
          <Settings className="w-4.5 h-4.5 text-gray-500 group-hover:text-gray-900 transition-colors shrink-0" />
          <span className="truncate">Settings</span>
        </button>

        {isPinActive && onLockApp && (
          <button
            onClick={onLockApp}
            className="p-2 rounded-lg text-gray-400 hover:text-gray-900 hover:bg-gray-100 transition-all cursor-pointer"
            title="Lock App (4-Digit PIN)"
          >
            <Lock className="w-4 h-4" />
          </button>
        )}
      </div>
    </aside>
  );
};
