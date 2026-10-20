import React from 'react';
import logoUrl from '../../assets/logo.png';

interface BrandLogoProps {
  className?: string;
  size?: number;
}

/**
 * Microsoft Copilot Style — Multi-Color Fluid Folded Gradient Ribbon Loop
 * - Distinct vibrant ribbons: Royal Blue/Cyan, Magenta/Violet, and Mint/Emerald
 * - Continuous harmonious Möbius loop
 * - Strictly NO alphabets / NO letters
 * - Crisp on clean white background (no dark theme box)
 */
export const BrandLogo: React.FC<BrandLogoProps> = ({
  className = 'w-7 h-7',
  size,
}) => {
  const style = size ? { width: size, height: size } : undefined;

  return (
    <img
      src={logoUrl}
      alt="CA Copilot"
      className={`${className} object-contain select-none`}
      style={style}
      loading="eager"
    />
  );
};
