import { readFileSync, writeFileSync } from 'fs';
import { resolve } from 'path';

const file = resolve('src/App.tsx');
let content = readFileSync(file, 'utf-8');

const replacements = [
  // Emojis and special markers
  ['ðŸ“¡', '📡'],
  ['ðŸ”’', '🔒'],
  ['ðŸ”„', '🔄'],
  ['ðŸ’³', '💳'],
  ['ðŸ“²', '📲'],
  ['ðŸ›¡ï¸ ', '🛡️'],
  ['ðŸ’¸', '💸'],
  ['ðŸŽ‰', '🎉'],
  ['ðŸ”Ž', '🔎'],
  ['ðŸ› ï¸ ', '🛠️'],
  ['ðŸš€', '🚀'],
  ['ðŸ‘‹', '👋'],
  ['ðŸ’¡', '💡'],
  ['ðŸšš', '🚚'],
  ['ðŸ“¦', '📦'],
  ['ðŸ‘¥', '👥'],
  ['ðŸ’°', '💰'],
  ['ðŸ§¬', '🧪'],
  ['ðŸ§¾', '📊'],
  ['ðŸ“ˆ', '📈'],
  ['ðŸ“Š', '📊'],
  ['ðŸ“„', '📄'],
  ['ðŸ””', '🔔'],
  ['ðŸ—‘ï¸ ', '🗑️'],
  ['ðŸ’¾', '💾'],
  ['ðŸ†“', '🆓'],
  ['ðŸš©', '🚩'],
  ['ðŸ”´', '🔴'],
  ['ðŸ“ ', '📌'],
  ['ðŸ”Œ', '🔌'],
  ['ðŸ’¬', '💬'],
  ['ðŸŽ›ï¸ ', '🎛️'],
  ['ðŸ“ ', '📌'],
  ['â€¢', '•'],
  ['Ã—', '×'],
  ['Â±', '±'],
  ['âœ“', '✓'],
  ['âš ï¸ ', '⚠️'],
  ['âœ…', '✅'],
  ['â€”', '—'],
  ['â€“', '–'],
  ['âš™ï¸ ', '⚙️'],
  ['âœ—', '✗'],
  ['â–▲', '▲'],
  ['â–▼', '▼'],
  ['â–²', '▲'],
  ['â–▼', '▼'],
  ['âœ•', '✕'],
  ['âœ ï¸ ', '✍️'],
  ['â Œ', '❌'],
  ['â†©ï¸ ', '↩️'],
  ['âž•', '➕'],
  ['â ³', '⏳'],
  ['â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€', '───────────'],
];

let totalFixed = 0;
for (const [corrupt, clean] of replacements) {
  const count = content.split(corrupt).length - 1;
  if (count > 0) {
    content = content.replaceAll(corrupt, clean);
    console.log(`✓ Restored ${count} occurrences of "${corrupt}" -> "${clean}"`);
    totalFixed += count;
  }
}

// Special case for 'ðŸ” ' which might have different trailing control chars
const specialSearches = [
  { target: 'ðŸ” \x91', replacement: '🔑' },
  { target: 'ðŸ” \x8d', replacement: '🔍' },
  { target: 'ðŸ” ', replacement: '🔍' } // general fallback
];

for (const { target, replacement } of specialSearches) {
  if (content.includes(target)) {
    const count = content.split(target).length - 1;
    content = content.replaceAll(target, replacement);
    console.log(`✓ Restored ${count} special occurrences of "${target}" -> "${replacement}"`);
    totalFixed += count;
  }
}

writeFileSync(file, content, 'utf-8');
console.log(`✓ Successfully fixed all ${totalFixed} corrupt emojis and UI elements in App.tsx!`);
