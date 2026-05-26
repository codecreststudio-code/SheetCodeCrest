/**
 * SheetCodeCrest - Emoji Recovery and UTF-8 Repair Tool
 * 
 * This script scans source files (like src/App.tsx) for double-encoded UTF-8 characters
 * (corrupted emojis like "ðŸ›¡ï¸" or "âš™ï¸") and restores them to their clean, native Unicode
 * equivalents. It works using both a precise replacement dictionary and a fallback
 * byte-reconstruction engine (Windows-1252 / ISO-8859-1 to UTF-8 translation).
 */

const fs = require('fs');
const path = require('path');

// Dictionary of known emoji corruptions for high-fidelity correction
const EMOJI_DICTIONARY = {
  // Common UI Emojis
  'ðŸ›¡ï¸': '🛡️',  // Shield / Security
  'âš™ï¸': '⚙️',  // Gear / Settings
  'ðŸ”': '🔎',   // Magnifying glass / Search
  'ðŸ—‘ï¸': '🗑️',  // Trash / Delete
  'ðŸ’¬': '💬',   // Chat bubble
  'ðŸŽ›ï¸': '🎛️',  // Control knobs / Dashboard
  'âœï¸': '✏️',   // Pencil / Edit
  'ðŸ’°': '💰',   // Money bag
  'ðŸ’³': '💳',   // Credit card
  'â ³': '⏳',    // Hourglass
  'âš ï¸': '⚠️',  // Warning
  'ðŸ”¥': '🔥',   // Fire / Dynamic
  'ðŸš€': '🚀',   // Rocket / Launch
  'âœ…': '✅',    // Green checkmark
  'ðŸ“ ': '📂',   // Folder / Saved sheets
  '🆓': '🆓',     // Free tier badge
  'â†©ï¸': '↩️',  // Return / Refund
  'â Œ': '❌',    // Crossmark
  'ðŸ’¥': '💥',   // Explosion
  'ðŸ“Š': '📊',   // Bar chart
  'ðŸ“ˆ': '📈',   // Line chart
  'ðŸ’💡': '💡',  // Lightbulb
  'ðŸ”🔒': '🔒',  // Lock
  'ðŸ”🔓': '🔓',  // Unlock
  'ðŸŒ': '🌐',    // Globe
  'ðŸ§': '🧠',    // Brain / AI analyst
  'ðŸ“‹': '📋',   // Clipboard
  'ðŸ“Œ': '📌',   // Pin
};

// Character map for decoding Windows-1252 / CP1252 bytes that get mangled
// by typical JS string buffer reading when double-encoded.
const CP1252_MAP = {
  0x80: 0x20AC, 0x82: 0x201A, 0x83: 0x0192, 0x84: 0x201E, 0x85: 0x2026,
  0x86: 0x2020, 0x87: 0x2021, 0x88: 0x02C6, 0x89: 0x2030, 0x8A: 0x0160,
  0x8B: 0x2039, 0x8C: 0x0152, 0x8E: 0x017D, 0x91: 0x2018, 0x92: 0x2019,
  0x93: 0x201C, 0x94: 0x201D, 0x95: 0x2022, 0x96: 0x2013, 0x97: 0x2014,
  0x98: 0x02DC, 0x99: 0x2122, 0x9A: 0x0161, 0x9B: 0x203A, 0x9C: 0x0153,
  0x9E: 0x017E, 0x9F: 0x0178
};

/**
 * Reconstructs a clean UTF-8 string from a double-encoded string
 * by mapping JavaScript char codes back to their original byte values
 * and interpreting those bytes as UTF-8.
 */
function autoRepairDoubleEncoding(str) {
  const bytes = [];
  
  for (let i = 0; i < str.length; i++) {
    const code = str.charCodeAt(i);
    
    if (code <= 0xFF) {
      bytes.push(code);
    } else {
      // Check if it's a character from CP-1252
      let found = false;
      for (const [byteVal, uniVal] of Object.entries(CP1252_MAP)) {
        if (uniVal === code) {
          bytes.push(parseInt(byteVal, 10));
          found = true;
          break;
        }
      }
      if (!found) {
        // Fallback: split wide characters into high/low bytes if applicable
        // or keep as is if it can't be decoded
        bytes.push(code & 0xFF);
      }
    }
  }
  
  try {
    const buffer = Buffer.from(bytes);
    return buffer.toString('utf8');
  } catch (e) {
    return str; // Return original if reconstruction fails
  }
}

/**
 * Main function to scan and repair a file
 */
function repairFile(filePath) {
  console.log(`🔍 Scanning file: ${filePath}`);
  
  if (!fs.existsSync(filePath)) {
    console.error(`❌ Error: File not found at ${filePath}`);
    return;
  }
  
  const content = fs.readFileSync(filePath, 'utf8');
  let repairedContent = content;
  let dictionaryFixCount = 0;
  
  // 1. Apply high-fidelity dictionary replacements
  for (const [corrupted, clean] of Object.entries(EMOJI_DICTIONARY)) {
    if (repairedContent.includes(corrupted)) {
      // Use split/join to replace all occurrences
      const occurrences = repairedContent.split(corrupted).length - 1;
      repairedContent = repairedContent.split(corrupted).join(clean);
      console.log(`✨ Dictionary Match: Restored "${corrupted}" ➔ "${clean}" (${occurrences} occurrences)`);
      dictionaryFixCount += occurrences;
    }
  }
  
  // 2. Identify potential remaining corrupted sequences starting with typical signature prefixes
  // (e.g. sequences containing mangled characters like â, ð, ï, etc.)
  const corruptionRegex = /(?:ðŸ[^\s"'>]*|âš[^\s"'>]*|âœ[^\s"'>]*|ï¸[^\s"'>]*)/g;
  let autoFixCount = 0;
  
  repairedContent = repairedContent.replace(corruptionRegex, (match) => {
    // Avoid double-fixing if already repaired or normal text
    if (match.length < 2) return match;
    
    const repaired = autoRepairDoubleEncoding(match);
    if (repaired !== match && repaired.length > 0 && repaired !== '\uFFFD') {
      console.log(`🤖 Auto-Reconstructed: "${match}" ➔ "${repaired}"`);
      autoFixCount++;
      return repaired;
    }
    return match;
  });

  const totalFixes = dictionaryFixCount + autoFixCount;
  
  if (totalFixes > 0) {
    // Create backup file before writing
    const backupPath = `${filePath}.bak`;
    fs.writeFileSync(backupPath, content, 'utf8');
    console.log(`💾 Saved secure backup of original file to: ${path.basename(backupPath)}`);
    
    // Write repaired content
    fs.writeFileSync(filePath, repairedContent, 'utf8');
    console.log(`✅ Success! Successfully repaired ${totalFixes} corrupted character strings in ${path.basename(filePath)}.\n`);
  } else {
    console.log(`🛡️  Clean! No corrupted emoji strings or double-encoded characters found in ${path.basename(filePath)}.\n`);
  }
}

// Automatically target App.tsx when run
const targetFile = path.resolve(__dirname, '../src/App.tsx');
repairFile(targetFile);
