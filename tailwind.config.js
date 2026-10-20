/** @type {import('tailwindcss').Config} */
export default {
  content: [
    "./index.html",
    "./src/**/*.{js,ts,jsx,tsx}",
  ],
  theme: {
    extend: {
      colors: {
        brand: {
          50: '#F8FAFC',
          100: '#F1F5F9',
          200: '#E2E8F0',
          500: '#0F172A', // Primary Green
          600: '#1E293B',
          700: '#0F172A', // Dark Green
          800: '#020617',
        },
        navy: {
          50: '#F4F7FB',
          100: '#E8EDF5',
          500: '#243B5A', // Navy Blue
          600: '#1F324D',
          700: '#18273D',
        },
        danger: {
          50: '#FDF2F2',
          100: '#FDE8E8',
          500: '#C43D43', // Error Red
          600: '#B03338',
        },
        neutral: {
          bg: '#F8F9FA',       // Main Background
          surface: '#FFFFFF',  // Card Background
          border: '#E5E7EB',   // Borders
          textPrimary: '#111827', // Primary Text
          textSecondary: '#4B5563', // Secondary Text
          muted: '#9CA3AF',
        }
      },
      fontFamily: {
        sans: ['Inter', 'system-ui', '-apple-system', 'BlinkMacSystemFont', 'Segoe UI', 'Roboto', 'sans-serif'],
      },
      boxShadow: {
        'subtle': '0 1px 2px 0 rgba(0, 0, 0, 0.04)',
        'card': '0 1px 3px 0 rgba(32, 41, 35, 0.05), 0 1px 2px -1px rgba(32, 41, 35, 0.05)',
        'modal': '0 10px 25px -5px rgba(32, 41, 35, 0.1), 0 8px 10px -6px rgba(32, 41, 35, 0.1)',
      }
    },
  },
  plugins: [],
}
