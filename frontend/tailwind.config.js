/** @type {import('tailwindcss').Config} */
export default {
  content: ['./index.html', './src/**/*.{js,jsx}'],
  theme: {
    extend: {
      colors: {
        // Artilegenz brand palette
        navy:   { DEFAULT: '#263248', 50: '#3a4a68', 900: '#1a2233', 950: '#10151f' },
        slate:  { DEFAULT: '#7E8AA2', 50: '#9eaabe', 200: '#7E8AA2', 700: '#4a566c' },
        amber:  { DEFAULT: '#FF9800', 50: '#FFB347', 600: '#cc7a00' },
        silver: '#D9D9D9',

        // Operator console additions
        ink:     '#0B0F18',     // page background
        carbon:  '#13192A',     // panel background
        graphite:'#1C2238',     // elevated surface
        cable:   '#2A3349',     // borders
        signal:  '#FF9800',     // primary accent (amber)
        critical:'#FF4444',
        warn:    '#FFAA33',
        ok:      '#3DDC84',
        muted:   '#6E7891',
      },
      fontFamily: {
        // Operator console: monospace data, geometric display
        mono:    ['JetBrains Mono', 'Fira Code', 'Consolas', 'Menlo', 'monospace'],
        display: ['"Space Grotesk"', 'system-ui', 'sans-serif'],
        body:    ['"Inter Tight"', 'system-ui', 'sans-serif'],
      },
      letterSpacing: {
        wider2: '0.18em',
      },
      boxShadow: {
        panel: '0 0 0 1px #2A3349, 0 8px 32px -8px rgba(0,0,0,.6)',
        glow:  '0 0 24px -6px rgba(255,152,0,.5)',
      },
      animation: {
        pulse_slow: 'pulse 3s ease-in-out infinite',
        scan:       'scan 2.4s linear infinite',
      },
      keyframes: {
        scan: {
          '0%':   { transform: 'translateX(-100%)' },
          '100%': { transform: 'translateX(100%)'  },
        },
      },
    },
  },
  plugins: [],
}
