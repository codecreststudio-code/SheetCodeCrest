import { readFileSync, writeFileSync } from 'fs';
import { resolve } from 'path';

const file = resolve('src/App.tsx');
let content = readFileSync(file, 'utf-8');

const target = '    const key = getClaudeKey();\r\n    if (!key) throw new Error("No Claude API key configured.");';
const targetLF = '    const key = getClaudeKey();\n    if (!key) throw new Error("No Claude API key configured.");';

const replacement = `    const key = getClaudeKey();
    if (!key) {
      console.log("No browser API key found, trying secure backend proxy...");
      try {
        const res = await fetch(API_PROXY_URL, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ messages, system, model: "claude-sonnet-4-20250514" }),
        });
        if (res.ok) {
          const data = await res.json();
          if (data.error) throw new Error(data.error);
          return data.content?.[0]?.text || data.result || "";
        }
        const txt = await res.text();
        throw new Error(\`Proxy status \${res.status}: \${txt}\`);
      } catch (proxyErr: any) {
        console.warn("Backend proxy failed/unavailable:", proxyErr);
        throw new Error(\`No Claude API key configured. Set VITE_CLAUDE_API_KEY in .env or provide it via browser settings. (Proxy error: \${proxyErr.message || proxyErr})\`);
      }
    }`;

if (content.includes(target)) {
  content = content.replace(target, replacement);
  console.log('✓ Patched with CRLF target');
} else if (content.includes(targetLF)) {
  content = content.replace(targetLF, replacement);
  console.log('✓ Patched with LF target');
} else {
  // Try normal replace of just the key check line
  const fallbackTarget = 'if (!key) throw new Error("No Claude API key configured.");';
  if (content.includes(fallbackTarget)) {
    content = content.replace(fallbackTarget, `if (!key) {
      console.log("No browser API key found, trying secure backend proxy...");
      try {
        const res = await fetch(API_PROXY_URL, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ messages, system, model: "claude-sonnet-4-20250514" }),
        });
        if (res.ok) {
          const data = await res.json();
          if (data.error) throw new Error(data.error);
          return data.content?.[0]?.text || data.result || "";
        }
        const txt = await res.text();
        throw new Error(\`Proxy status \${res.status}: \${txt}\`);
      } catch (proxyErr: any) {
        console.warn("Backend proxy failed/unavailable:", proxyErr);
        throw new Error(\`No Claude API key configured. Set VITE_CLAUDE_API_KEY in .env or provide it via browser settings. (Proxy error: \${proxyErr.message || proxyErr})\`);
      }
    }`);
    console.log('✓ Patched with fallback target');
  } else {
    console.log('✗ Target string not found in App.tsx!');
  }
}

writeFileSync(file, content, 'utf-8');
