/**
 * 🔐 securityUtils.ts
 * Security-First Vibe Coding — SheetCodeCrest
 * 
 * Implements:
 *  1. Web Crypto API password hashing (SHA-256 + salt)
 *  2. Input sanitization & validation
 *  3. Rate limiting / login lockout (brute force protection)
 */

// ──────────────────────────────────────────────────────────────────────────────
// 1. PASSWORD HASHING (Web Crypto API — SHA-256 + random salt)
// ──────────────────────────────────────────────────────────────────────────────

/**
 * Converts an ArrayBuffer to a hex string
 */
function bufToHex(buf: ArrayBuffer): string {
  return Array.from(new Uint8Array(buf))
    .map((b) => b.toString(16).padStart(2, "0"))
    .join("");
}

/**
 * Generates a cryptographically-random 16-byte salt encoded as hex
 */
export function generateSalt(): string {
  const saltBytes = new Uint8Array(16);
  crypto.getRandomValues(saltBytes);
  return bufToHex(saltBytes.buffer);
}

/**
 * Hashes a plain-text password using SHA-256 with the given salt.
 * Returns a string in the format: `sha256$<salt>$<hash>`
 * Safe for browser — uses the Web Crypto API.
 */
export async function hashPassword(password: string, salt?: string): Promise<string> {
  const usedSalt = salt || generateSalt();
  const encoded = new TextEncoder().encode(usedSalt + password);
  const hashBuffer = await crypto.subtle.digest("SHA-256", encoded);
  const hashHex = bufToHex(hashBuffer);
  return `sha256$${usedSalt}$${hashHex}`;
}

/**
 * Verifies a plain-text password against a stored hash string.
 * Handles:
 *  - New `sha256$<salt>$<hash>` format
 *  - Google OAuth sentinel strings (prefixed with `google_oauth_`)
 *  - Legacy plain-text passwords (migrates on next successful login)
 */
export async function verifyPassword(
  plaintext: string,
  stored: string
): Promise<boolean> {
  if (!stored || !plaintext) return false;

  // Google OAuth accounts — cannot login with password
  if (stored.startsWith("google_oauth_")) return false;

  // New hashed format
  if (stored.startsWith("sha256$")) {
    const parts = stored.split("$");
    if (parts.length !== 3) return false;
    const [, salt] = parts;
    const recomputed = await hashPassword(plaintext, salt);
    return recomputed === stored;
  }

  // Legacy plain-text comparison (will be migrated to hash on next successful login)
  return stored === plaintext;
}

/**
 * Returns true if the stored hash is in the legacy plain-text format
 * and needs to be migrated.
 */
export function isLegacyPassword(stored: string): boolean {
  return (
    !!stored &&
    !stored.startsWith("sha256$") &&
    !stored.startsWith("google_oauth_")
  );
}


// ──────────────────────────────────────────────────────────────────────────────
// 2. INPUT SANITIZATION & VALIDATION
// ──────────────────────────────────────────────────────────────────────────────

/** Strips all leading/trailing whitespace and normalizes internal spaces */
export function sanitizeText(input: string): string {
  return input.trim().replace(/\s+/g, " ");
}

/** Validates a username — alphanumeric, underscores, hyphens, 3-32 chars */
export function isValidUsername(username: string): boolean {
  return /^[a-zA-Z0-9_.-]{3,32}$/.test(username);
}

/** Validates a password — minimum 6 chars */
export function isValidPassword(password: string): boolean {
  return password.length >= 6;
}

/** Validates a mobile number — 10 digits (Indian format) */
export function isValidMobile(mobile: string): boolean {
  return /^[6-9]\d{9}$/.test(mobile.replace(/\s+/g, ""));
}

/** Validates email format */
export function isValidEmail(email: string): boolean {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}


// ──────────────────────────────────────────────────────────────────────────────
// 3. RATE LIMITING / BRUTE FORCE LOCKOUT
// ──────────────────────────────────────────────────────────────────────────────

const RATE_LIMIT_KEY_PREFIX = "sheetcc_ratelimit_";
const MAX_ATTEMPTS = 5;
const LOCKOUT_MS = 5 * 60 * 1000; // 5 minutes

interface RateLimitState {
  attempts: number;
  lockedUntil: number; // epoch ms, 0 means not locked
}

function getRateLimitState(username: string): RateLimitState {
  try {
    const raw = sessionStorage.getItem(RATE_LIMIT_KEY_PREFIX + username);
    if (raw) return JSON.parse(raw);
  } catch (_) {}
  return { attempts: 0, lockedUntil: 0 };
}

function setRateLimitState(username: string, state: RateLimitState): void {
  try {
    sessionStorage.setItem(RATE_LIMIT_KEY_PREFIX + username, JSON.stringify(state));
  } catch (_) {}
}

/**
 * Checks if the username is currently locked out.
 * Returns the remaining lockout seconds (0 if not locked).
 */
export function getLockoutSecondsRemaining(username: string): number {
  const state = getRateLimitState(username);
  if (state.lockedUntil && Date.now() < state.lockedUntil) {
    return Math.ceil((state.lockedUntil - Date.now()) / 1000);
  }
  return 0;
}

/**
 * Records a failed login attempt for the given username.
 * Locks the account for LOCKOUT_MS if MAX_ATTEMPTS is exceeded.
 * Returns true if the account is now locked.
 */
export function recordFailedAttempt(username: string): boolean {
  const state = getRateLimitState(username);

  // Reset if previous lockout has expired
  if (state.lockedUntil && Date.now() >= state.lockedUntil) {
    setRateLimitState(username, { attempts: 1, lockedUntil: 0 });
    return false;
  }

  state.attempts += 1;
  if (state.attempts >= MAX_ATTEMPTS) {
    state.lockedUntil = Date.now() + LOCKOUT_MS;
  }

  setRateLimitState(username, state);
  return state.attempts >= MAX_ATTEMPTS;
}

/**
 * Resets the failed attempt counter for the given username (after successful login).
 */
export function clearFailedAttempts(username: string): void {
  try {
    sessionStorage.removeItem(RATE_LIMIT_KEY_PREFIX + username);
  } catch (_) {}
}
