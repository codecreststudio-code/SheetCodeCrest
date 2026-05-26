import http from "http";
import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// 1. Parse .env files manually to avoid dependency requirement
function loadEnv() {
  const envPaths = [
    path.join(__dirname, ".env.local"),
    path.join(__dirname, ".env")
  ];
  for (const envPath of envPaths) {
    if (fs.existsSync(envPath)) {
      try {
        const content = fs.readFileSync(envPath, "utf-8");
        content.split("\n").forEach((line) => {
          const match = line.match(/^\s*([\w.\-_]+)\s*=\s*(.*)?\s*$/);
          if (match) {
            const key = match[1];
            let value = (match[2] || "").trim();
            // remove quotes if present
            if (value.startsWith('"') && value.endsWith('"')) value = value.slice(1, -1);
            if (value.startsWith("'") && value.endsWith("'")) value = value.slice(1, -1);
            process.env[key] = value;
          }
        });
        console.log(`☁️ Loaded environmental keys from: ${path.basename(envPath)}`);
      } catch (err) {
        console.error(`❌ Failed to read environment file at ${envPath}:`, err);
      }
    }
  }
}

loadEnv();

const PORT = process.env.VITE_BACKEND_PORT || 5001;
const CLAUDE_API_KEY = process.env.VITE_CLAUDE_API_KEY || process.env.CLAUDE_API_KEY;

if (!CLAUDE_API_KEY) {
  console.warn("⚠️ WARNING: No VITE_CLAUDE_API_KEY or CLAUDE_API_KEY found in your environment variables. The proxy server will fail to route AI queries.");
} else {
  console.log("🔒 Secure Claude API Key verified. Backend proxy is armed and ready.");
}

// 2. Build the HTTP server
const server = http.createServer(async (req, res) => {
  // CORS Headers
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type, x-api-key, Authorization");

  // Handle Preflight Request
  if (req.method === "OPTIONS") {
    res.writeHead(204);
    res.end();
    return;
  }

  // Handle Proxy API Chat
  if (req.method === "POST" && req.url === "/api/chat") {
    let bodyStr = "";
    req.on("data", (chunk) => {
      bodyStr += chunk;
    });

    req.on("end", async () => {
      try {
        if (!CLAUDE_API_KEY) {
          throw new Error("Claude API key not configured on the backend server.");
        }

        const payload = JSON.parse(bodyStr);
        const { messages, system, model } = payload;

        if (!messages || !Array.isArray(messages)) {
          res.writeHead(400, { "Content-Type": "application/json" });
          res.end(JSON.stringify({ error: "Missing required parameter 'messages' or it is not an array." }));
          return;
        }

        console.log(`💬 AI proxy request received. Model: ${model || "claude-sonnet-4-20250514"} | Messages: ${messages.length}`);

        // Direct server-to-server request to Anthropic Claude
        const anthropicRes = await fetch("https://api.anthropic.com/v1/messages", {
          method: "POST",
          headers: {
            "Content-Type": "application/json",
            "x-api-key": CLAUDE_API_KEY,
            "anthropic-version": "2023-06-01"
          },
          body: JSON.stringify({
            model: model || "claude-3-5-sonnet-20241022", // Default to upgraded premium Sonnet
            max_tokens: 1500,
            system: system || "",
            messages: messages
          })
        });

        const anthropicData = await anthropicRes.json();

        if (!anthropicRes.ok) {
          console.error("❌ Anthropic API call failed:", anthropicData);
          res.writeHead(anthropicRes.status, { "Content-Type": "application/json" });
          res.end(JSON.stringify(anthropicData));
          return;
        }

        res.writeHead(200, { "Content-Type": "application/json" });
        res.end(JSON.stringify(anthropicData));
      } catch (err) {
        console.error("❌ Proxy Server Error:", err.message);
        res.writeHead(500, { "Content-Type": "application/json" });
        res.end(JSON.stringify({ error: err.message || "Internal server error" }));
      }
    });
  } else {
    // Route not found
    res.writeHead(404, { "Content-Type": "application/json" });
    res.end(JSON.stringify({ error: "Endpoint not found." }));
  }
});

server.listen(PORT, () => {
  console.log(`⚡ SheetCodeCrest Secure AI Proxy running on http://localhost:${PORT}`);
  console.log(`👉 Connect frontend queries to: http://localhost:${PORT}/api/chat`);
});
