import React, { useRef, useState, useEffect } from 'react';

export interface SegmentedTab<T extends string> {
  id: T;
  label: string;
  icon?: React.ComponentType<{ className?: string }>;
}

interface SegmentedControlProps<T extends string> {
  tabs: SegmentedTab<T>[];
  activeTab: T;
  onChange: (tabId: T) => void;
  className?: string;
  size?: 'sm' | 'md';
}

export function SegmentedControl<T extends string>({
  tabs,
  activeTab,
  onChange,
  className = '',
  size = 'sm',
}: SegmentedControlProps<T>) {
  const containerRef = useRef<HTMLDivElement>(null);
  const tabRefs = useRef<Record<string, HTMLButtonElement | null>>({});
  const [indicatorStyle, setIndicatorStyle] = useState({
    left: 4,
    top: 4,
    width: 0,
    height: 0,
  });
  const [isReady, setIsReady] = useState(false);

  useEffect(() => {
    const updateIndicator = () => {
      const activeBtn = tabRefs.current[activeTab];
      if (activeBtn) {
        setIndicatorStyle({
          left: activeBtn.offsetLeft,
          top: activeBtn.offsetTop,
          width: activeBtn.offsetWidth,
          height: activeBtn.offsetHeight,
        });
        if (!isReady) {
          // Allow first render to position pill without animating from (0,0)
          requestAnimationFrame(() => {
            setIsReady(true);
          });
        }
      }
    };

    updateIndicator();

    // Re-calculate on resize or font load
    window.addEventListener('resize', updateIndicator);

    const container = containerRef.current;
    let observer: ResizeObserver | null = null;
    if (container && typeof ResizeObserver !== 'undefined') {
      observer = new ResizeObserver(() => updateIndicator());
      observer.observe(container);
    }

    return () => {
      window.removeEventListener('resize', updateIndicator);
      if (observer) observer.disconnect();
    };
  }, [activeTab, tabs]);

  const paddingY = size === 'sm' ? 'py-1.5' : 'py-2';
  const paddingX = size === 'sm' ? 'px-3' : 'px-3.5';
  const fontSize = size === 'sm' ? 'text-xs' : 'text-sm';

  return (
    <div
      ref={containerRef}
      role="tablist"
      className={`relative inline-flex p-1 bg-gray-100 border border-gray-200 rounded-lg shrink-0 overflow-x-auto no-scrollbar items-center select-none ${className}`}
    >
      {/* Apple Fluid Sliding Pill Indicator */}
      <div
        className={`absolute rounded-md bg-white border border-gray-200/90 shadow-2xs pointer-events-none will-change-[left,width] ${
          isReady
            ? 'transition-all duration-300 ease-[cubic-bezier(0.16,1,0.3,1)]'
            : 'transition-none opacity-0'
        }`}
        style={{
          left: `${indicatorStyle.left}px`,
          top: `${indicatorStyle.top}px`,
          width: `${indicatorStyle.width}px`,
          height: `${indicatorStyle.height}px`,
          opacity: indicatorStyle.width > 0 ? 1 : 0,
        }}
      />

      {/* Tab Buttons */}
      {tabs.map((tab) => {
        const isActive = activeTab === tab.id;
        const Icon = tab.icon;

        return (
          <button
            key={tab.id}
            ref={(el) => {
              tabRefs.current[tab.id] = el;
            }}
            role="tab"
            aria-selected={isActive}
            onClick={() => onChange(tab.id)}
            className={`relative z-10 inline-flex items-center gap-1.5 ${paddingX} ${paddingY} rounded-md ${fontSize} cursor-pointer whitespace-nowrap transition-colors duration-200 ${
              isActive
                ? 'text-gray-900 font-bold'
                : 'text-gray-600 hover:text-gray-900 font-medium'
            }`}
          >
            {Icon && (
              <Icon
                className={`w-3.5 h-3.5 transition-colors duration-200 ${
                  isActive ? 'text-gray-900' : 'text-gray-400'
                }`}
              />
            )}
            <span>{tab.label}</span>
          </button>
        );
      })}
    </div>
  );
}
