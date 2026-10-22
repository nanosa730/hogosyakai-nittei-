/** @type {import('tailwindcss').Config} */
module.exports = {
  content: ['./src/index.html', './src/app.js'],
  theme: {
    extend: {
      colors: {
        page: '#f6f7f9',
        ink: '#1f2933',
        muted: '#5f6b7a',
        line: '#d9dee5',
        accent: { DEFAULT: '#1a6fd1', soft: '#f0f6ff', over: '#d9e9ff' },
        ok: { DEFAULT: '#2f8f4e', bg: '#e9f7ee', line: '#b6e0c4' },
        warn: { DEFAULT: '#b7791f', bg: '#fff6e0', line: '#f0d58c' },
        bad: { DEFAULT: '#c53030', bg: '#fde8e8', line: '#f3b4b4' },
        sel: { DEFAULT: '#ffe58f', line: '#d4a106', ring: '#f5cd4a' },
        sib: { DEFAULT: '#7b4dc4', bg: '#f3edfc', soft: '#f6f2fd', line: '#d9c9f3' },
        faint: '#b0b7c3',
      },
      fontFamily: {
        ui: ['"Yu Gothic UI"', 'Meiryo', '"Hiragino Sans"', 'sans-serif'],
        letter: ['"UD デジタル 教科書体 NP-R"', '"UD Digi Kyokasho NP-R"', '"UD デジタル 教科書体 NP"', '"BIZ UD明朝"', '"ＭＳ 明朝"', '"MS Mincho"', 'serif'],
        sched: ['"UD デジタル 教科書体 NP-R"', '"UD Digi Kyokasho NP-R"', '"BIZ UDゴシック"', 'Meiryo', 'sans-serif'],
      },
    },
  },
  plugins: [],
};
