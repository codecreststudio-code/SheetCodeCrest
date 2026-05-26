import { readFileSync, writeFileSync } from 'fs';
import { resolve } from 'path';

const file = resolve('src/App.tsx');
let content = readFileSync(file, 'utf-8');

console.log(`Original file length: ${content.length} characters.`);

// 1. Restore RGB colors that were mangled by replacing ",1" with "?"
const colorsMap = {
  'rgba(245?58?1,': 'rgba(245,158,11,',
  'rgba(100?16?39,': 'rgba(100,116,139,',
  'rgba(16?85?29,': 'rgba(16,185,129,',
  'rgba(59?30,': 'rgba(59,130,',
  'rgba(16?85?29,': 'rgba(16,185,129,'
};

let colorChanges = 0;
for (const [mangled, clean] of Object.entries(colorsMap)) {
  const count = content.split(mangled).length - 1;
  if (count > 0) {
    content = content.replaceAll(mangled, clean);
    console.log(`✓ Fixed ${count} occurrences of ${mangled} -> ${clean}`);
    colorChanges += count;
  }
}

// 2. Restore SVG path polyline chevron that got corrupted
const polylineCorrupt = 'polyline points="22,6 12?3 2,6"';
const polylineFixed = 'polyline points="22,6 12,13 2,6"';
if (content.includes(polylineCorrupt)) {
  content = content.replaceAll(polylineCorrupt, polylineFixed);
  console.log(`✓ Fixed corrupt SVG polyline points`);
}

// 3. Restore all mangled Rupee symbols (where U+20B9 '₹' became '?')
// We target '?' when followed by a digit, '$', or '{' (used inside template strings or formulas)
const rupeeRegex = /\?([0-9]|\$|\{)/g;
const matches = content.match(rupeeRegex);
const rupeeCount = matches ? matches.length : 0;

if (rupeeCount > 0) {
  content = content.replace(rupeeRegex, '₹$1');
  console.log(`✓ Fixed ${rupeeCount} occurrences of Rupee symbols`);
}

// Write the fixed file back
writeFileSync(file, content, 'utf-8');
console.log('✓ Successfully restored App.tsx file encoding and styles!');
