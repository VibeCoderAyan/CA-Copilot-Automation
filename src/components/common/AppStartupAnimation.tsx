import React, { useState, useEffect } from 'react';
import { BrandLogo } from './BrandLogo';

interface AppStartupAnimationProps {
  onComplete?: () => void;
  durationMs?: number;
}

/**
 * Normal, simple & clean auto-open animation:
 * - Shows only the official Brand Logo and Name ("CA Copilot")
 * - Subtle, smooth fade & scale entrance
 * - Fades away smoothly to reveal the main application
 * - Pure minimalism without cluttered tickers or progress bars
 */
export const AppStartupAnimation: React.FC<AppStartupAnimationProps> = ({
  onComplete,
  durationMs = 1100,
}) => {
  const [stage, setStage] = useState<'enter' | 'visible' | 'leave'>('enter');
  const [removed, setRemoved] = useState(false);

  useEffect(() => {
    // 1. Trigger smooth fade-in entrance immediately
    const enterTimer = requestAnimationFrame(() => {
      setStage('visible');
    });

    // 2. Start gentle fade-out transition after duration
    const leaveTimer = setTimeout(() => {
      setStage('leave');
    }, durationMs);

    // 3. Unmount completely from DOM
    const removeTimer = setTimeout(() => {
      setRemoved(true);
      if (onComplete) onComplete();
    }, durationMs + 450);

    return () => {
      cancelAnimationFrame(enterTimer);
      clearTimeout(leaveTimer);
      clearTimeout(removeTimer);
    };
  }, [durationMs, onComplete]);

  if (removed) return null;

  return (
    <div
      aria-hidden="true"
      className={`fixed inset-0 z-[9999] flex items-center justify-center bg-[#F8F9FA] select-none transition-opacity duration-400 ease-out pointer-events-none ${
        stage === 'leave' ? 'opacity-0' : 'opacity-100'
      }`}
    >
      <div
        className={`flex flex-col items-center gap-4 transition-all duration-500 ease-out transform ${
          stage === 'visible'
            ? 'opacity-100 scale-100 translate-y-0'
            : stage === 'leave'
            ? 'opacity-0 scale-102 -translate-y-1'
            : 'opacity-0 scale-95 translate-y-2'
        }`}
      >
        {/* Simple Brand Logo */}
        <BrandLogo className="w-16 h-16 object-contain drop-shadow-sm" size={64} />

        {/* Simple Brand Name */}
        <h1 className="text-2xl font-bold tracking-tight text-gray-900 font-sans">
          CA Copilot
        </h1>
      </div>
    </div>
  );
};
