# Fix App.tsx encoding corruption:
# 1. Replace all U+20B9 (₹ rupee sign) back to ? (ternary operator)
# 2. Remove the duplicate proxy-based api function (lines 75-98)

$file = "src\App.tsx"
$content = Get-Content $file -Raw -Encoding UTF8

Write-Host "Original file size: $($content.Length) chars"
Write-Host "Rupee sign (U+20B9) count before: $(([regex]::Matches($content, [char]0x20B9)).Count)"

# Step 1: Replace ₹ (U+20B9) with ? (ternary operator)
$rupeeChar = [char]0x20B9
$content = $content.Replace($rupeeChar, '?')

Write-Host "Rupee sign count after: $(([regex]::Matches($content, [char]0x20B9)).Count)"

# Step 2: Remove the FIRST (proxy-based) duplicate api function (lines 75-98)
# It starts with:  `const api = async (messages: any[], system: string) => {` 
# and uses `API_PROXY_URL` internally.
# We'll remove just the first occurrence up to the closing `};` followed by a newline before the second `const api`

$firstApiStart = @"
const api = async (messages: any[], system: string) => {
  try {
    const key = getClaudeKey();
    if (!key) throw new Error("No Claude API key configured.");

    const res = await fetch(API_PROXY_URL, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ messages, system, model: "claude-sonnet-4-20250514" }),
    });

    if (!res.ok) {
      const txt = await res.text();
      throw new Error(`AI API error ${res.status}: ${txt}`);
    }

    const data = await res.json();
    if (data.error) throw new Error(data.error);
    return data.content?.[0]?.text || data.result || "";
  } catch (err: any) {
    console.error("api error", err);
    return `Error: ${err.message || err}`;
  }
};
"@

if ($content.Contains($firstApiStart)) {
    $content = $content.Replace($firstApiStart, "")
    Write-Host "Removed duplicate proxy-based api function."
} else {
    Write-Host "WARNING: Could not find exact duplicate api function text. Manual check needed."
    # Try a regex-based approach for the first api block
    $pattern = '(?s)const api = async \(messages: any\[\], system: string\) => \{\r?\n  try \{\r?\n    const key = getClaudeKey\(\);\r?\n    if \(!key\) throw new Error\("No Claude API key configured\."\);\r?\n\r?\n    const res = await fetch\(API_PROXY_URL,'
    if ($content -match $pattern) {
        Write-Host "Found via regex, attempting removal..."
        # Find end of this block
        $match = [regex]::Match($content, $pattern)
        $startIdx = $match.Index
        # Find the closing }; after this block
        $searchFrom = $startIdx + $match.Length
        $braceDepth = 1
        $i = $searchFrom
        while ($i -lt $content.Length -and $braceDepth -gt 0) {
            if ($content[$i] -eq '{') { $braceDepth++ }
            elseif ($content[$i] -eq '}') { $braceDepth-- }
            $i++
        }
        # i is now past the closing }
        # skip the semicolon and newlines
        while ($i -lt $content.Length -and ($content[$i] -eq ';' -or $content[$i] -eq "`r" -or $content[$i] -eq "`n")) { $i++ }
        $content = $content.Substring(0, $startIdx) + $content.Substring($i)
        Write-Host "Removed via regex."
    }
}

# Write the fixed content back
[System.IO.File]::WriteAllText((Resolve-Path $file), $content, [System.Text.Encoding]::UTF8)
Write-Host "Fixed file written to $file"
Write-Host "New file size: $($content.Length) chars"
