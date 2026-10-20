import React, { useState, useEffect } from 'react';
import { Delete } from 'lucide-react';
import { BrandLogo } from './BrandLogo';
import { verifyPin } from '../../utils/security';

interface PinLockScreenProps {
  onUnlock: () => void;
}

export const PinLockScreen: React.FC<PinLockScreenProps> = ({ onUnlock }) => {
  const [pin, setPin] = useState('');
  const [error, setError] = useState(false);
  const [errorMsg, setErrorMsg] = useState('');

  const handleDigit = (digit: string) => {
    if (pin.length < 4) {
      const nextPin = pin + digit;
      setPin(nextPin);
      setError(false);
      setErrorMsg('');

      if (nextPin.length === 4) {
        // Automatically verify upon 4 digits entered
        setTimeout(() => {
          if (verifyPin(nextPin)) {
            onUnlock();
          } else {
            setError(true);
            setErrorMsg('Incorrect PIN. Please try again.');
            setPin('');
          }
        }, 150);
      }
    }
  };

  const handleBackspace = () => {
    setPin((prev) => prev.slice(0, -1));
    setError(false);
    setErrorMsg('');
  };

  // Keyboard support: listen for 0-9, Backspace, Escape
  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      if (/^[0-9]$/.test(e.key)) {
        handleDigit(e.key);
      } else if (e.key === 'Backspace') {
        handleBackspace();
      }
    };

    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [pin]);

  return (
    <div className="fixed inset-0 z-50 bg-[#F8F9FA]/95 backdrop-blur-xl flex flex-col items-center justify-center p-6 select-none animate-apple-fade">
      <div className="w-full max-w-sm flex flex-col items-center text-center">
        {/* Brand Logo */}
        <div className="w-16 h-16 rounded-2xl bg-white border border-gray-200/80 shadow-md flex items-center justify-center mb-4">
          <BrandLogo className="w-10 h-10" />
        </div>

        {/* Brand Name (Without 'Protected') */}
        <h2 className="text-xl font-bold text-gray-900 tracking-tight">CA Copilot</h2>
        <p className="text-xs text-gray-500 mt-1 max-w-xs">
          Enter your 4-digit PIN to continue
        </p>

        {/* 4-Digit Indicator Dots */}
        <div className={`flex items-center justify-center gap-3.5 my-6 ${error ? 'animate-shake' : ''}`}>
          {[0, 1, 2, 3].map((idx) => {
            const isFilled = pin.length > idx;
            return (
              <div
                key={idx}
                className={`w-3.5 h-3.5 rounded-full transition-all duration-150 ${
                  isFilled
                    ? 'bg-gray-900 scale-105'
                    : 'bg-gray-200 border border-gray-300'
                }`}
              />
            );
          })}
        </div>

        {/* Error message */}
        <div className="h-5 mb-2">
          {errorMsg && (
            <span className="text-xs font-semibold text-rose-600 animate-apple-fade">
              {errorMsg}
            </span>
          )}
        </div>

        {/* Keypad */}
        <div className="grid grid-cols-3 gap-3 w-64">
          {['1', '2', '3', '4', '5', '6', '7', '8', '9'].map((digit) => (
            <button
              key={digit}
              onClick={() => handleDigit(digit)}
              className="h-14 rounded-2xl bg-white border border-gray-200 text-lg font-semibold text-gray-900 hover:bg-gray-50 hover:border-gray-300 active:scale-95 transition-all shadow-2xs flex items-center justify-center cursor-pointer"
            >
              {digit}
            </button>
          ))}
          <div />
          <button
            onClick={() => handleDigit('0')}
            className="h-14 rounded-2xl bg-white border border-gray-200 text-lg font-semibold text-gray-900 hover:bg-gray-50 hover:border-gray-300 active:scale-95 transition-all shadow-2xs flex items-center justify-center cursor-pointer"
          >
            0
          </button>
          <button
            onClick={handleBackspace}
            disabled={pin.length === 0}
            className="h-14 rounded-2xl bg-white border border-gray-200 text-gray-500 hover:bg-gray-50 hover:text-gray-900 active:scale-95 disabled:opacity-30 disabled:pointer-events-none transition-all shadow-2xs flex items-center justify-center cursor-pointer"
            title="Backspace"
          >
            <Delete className="w-5 h-5" />
          </button>
        </div>
      </div>
    </div>
  );
};
