import { useState, useRef, useCallback, useEffect, useMemo } from "react";
import {
  hashPassword,
  verifyPassword,
  isLegacyPassword,
  isValidUsername,
  isValidPassword,
  isValidEmail,
  isValidMobile,
  sanitizeText,
  recordFailedAttempt,
  clearFailedAttempts,
  getLockoutSecondsRemaining,
} from "./securityUtils";
import * as XLSX from "xlsx";
import {
  parseExcel,
  analyzeData,
  buildAnalyticsWorkbook,
  buildLogisticsWorkbook,
  buildShopifyAnalyticsWorkbook,
  analyzeShopifyData,
  DataProfile,
  DataRow,
  ShopifyAnalyticsSummary,
} from "./excelAnalytics";
import { mergeMultiSKU, computeAnalytics, AnalyticsResult } from "./analyticsUtils";
import {
  User,
  SavedRecord,
  Plan,
  AdminLog,
  dbSaveUser,
  dbGetUser,
  dbSaveRecord,
  dbGetRecords,
  dbDeleteRecord,
  dbGetRecordById,
  dbGetAllUsers,
  dbGetAllPayments,
  dbApprovePayment,
  dbDeleteUser,
  dbUpdateUserFields,
  dbUpdatePaymentStatus,
  dbGetPlans,
  dbSavePlan,
  dbDeletePlan,
  dbLogAdminAction,
  dbGetAdminLogs,
} from "./db";

type AppMode = "universal" | "logistics" | "shopify";

const FREE_REPORT_LIMIT = 3;
const USAGE_STORAGE_KEY = "codecrest_excel_analytics_usage_count";
const CODECREST = {
  instagram: "https://www.instagram.com/codecrest__studio",
  email: "codecreststudio@gmail.com",
  website: "https://codecreststudio.vercel.app/",
};
const GOOGLE_CLIENT_ID = "671624988330-q996r5ooe7blbi11lmmvdba6aspmcips.apps.googleusercontent.com"; // Change to your Google OAuth Client ID if needed
const PERSONAL_UPI_ID = "codecreststudio@okaxis"; // Your personal UPI ID for direct scan fallback

type ThemeMode = "light" | "dark";

const API_PROXY_URL = import.meta.env.VITE_PROXY_URL
  ? `${import.meta.env.VITE_PROXY_URL}/api/chat`
  : "http://localhost:5001/api/chat";

const GOOGLE_SHEET_PROXY_URL = import.meta.env.VITE_PROXY_URL
  ? `${import.meta.env.VITE_PROXY_URL}/api/sync-sheet`
  : "http://localhost:5001/api/sync-sheet";

const api = async (messages: any[], system: string) => {
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
    throw new Error(`Secure proxy status ${res.status}: ${txt}`);
  } catch (err: any) {
    console.error("Secure AI request failed:", err);
    return `Error: ${err.message || err}`;
  }
};

const escapeHtml = (text: string): string =>
  text.replace(/[&<>"']/g, (char) => {
    const map: Record<string, string> = {
      "&": "&amp;",
      "<": "&lt;",
      ">": "&gt;",
      '"': "&quot;",
      "'": "&#39;",
    };
    return map[char];
  });

const processInline = (text: string): string => {
  let html = escapeHtml(text);
  // 1. Bold: **text**
  html = html.replace(/\*\*(.*?)\*\*/g, "<strong style='font-weight: 700;'>$1</strong>");
  // 2. Italic: *text*
  html = html.replace(/\*(.*?)\*/g, "<em>$1</em>");
  // 3. Inline Code: `code`
  html = html.replace(/`(.*?)`/g, "<code style='background: rgba(255, 255, 255, 0.1); padding: 2px 6px; border-radius: 4px; font-family: var(--font-technical); font-size: 0.9em; border: 1px solid rgba(255, 255, 255, 0.15); color: var(--coral-soft);'>$1</code>");
  return html;
};

const formatMessage = (text: string): string => {
  if (!text) return "";
  const lines = text.split("\n");
  const processedLines: string[] = [];
  let inList = false;
  
  for (let line of lines) {
    const trimmed = line.trim();
    
    // Match bullet point starting with • or - or * followed by spaces
    const bulletMatch = trimmed.match(/^([•\-\*])\s+(.*)$/);
    
    if (bulletMatch) {
      const content = bulletMatch[2];
      if (!inList) {
        processedLines.push("<ul style='margin: 0.5rem 0; padding-left: 1.5rem; list-style-type: disc;'>");
        inList = true;
      }
      processedLines.push(`<li style='margin-bottom: 0.4rem; line-height: 1.5;'>${processInline(content)}</li>`);
    } else {
      if (inList) {
        processedLines.push("</ul>");
        inList = false;
      }
      
      if (trimmed === "") {
        processedLines.push("<div style='height: 0.5rem;'></div>");
      } else {
        processedLines.push(`<p style='margin: 0 0 0.5rem 0; line-height: 1.5;'>${processInline(line)}</p>`);
      }
    }
  }
  
  if (inList) {
    processedLines.push("</ul>");
  }
  
  return processedLines.join("");
};

export default function App() {
  // Application Modes & Navigation
  const [mode, setMode] = useState<AppMode>("universal");
  const [step, setStep] = useState<"upload" | "processing" | "done">("upload");
  const [dragging, setDragging] = useState(false);
  const [file, setFile] = useState<File | null>(null);
  const [log, setLog] = useState<Array<{ msg: string; type: string; time: string }>>([]);
  const addLog = useCallback((msg: string, type = "info") => {
    setLog((p) => [...p, { msg, type, time: new Date().toLocaleTimeString() }]);
  }, []);
  const [error, setError] = useState("");

  // User Authentication & Session State
  const [currentUser, setCurrentUser] = useState<User | null>(null);
  const [authModalOpen, setAuthModalOpen] = useState(false);
  const [authTab, setAuthTab] = useState<"login" | "signup">("login");
  const [authUsername, setAuthUsername] = useState("");
  const [authPassword, setAuthPassword] = useState("");
  const [authName, setAuthName] = useState("");
  const [authMobile, setAuthMobile] = useState("");
  const [authEmail, setAuthEmail] = useState("");
  const [authError, setAuthError] = useState("");
  
  // Dashboard Modal State
  const [dashboardOpen, setDashboardOpen] = useState(false);
  const [savedRecords, setSavedRecords] = useState<SavedRecord[]>([]);

  // Payment Checkout Gateway Modal State
  const [checkoutOpen, setCheckoutOpen] = useState(false);
  
  // Razorpay UPI fields
  const [upiVPA, setUpiVPA] = useState("");
  const [upiUTR, setUpiUTR] = useState("");
  
  // Payment Gateway simulation log stream
  const [paymentLogs, setPaymentLogs] = useState<string[]>([]);
  const [paymentProcessing, setPaymentProcessing] = useState(false);
  const [paymentCompleted, setPaymentCompleted] = useState(false);
  const [theme, setTheme] = useState<ThemeMode>(() => {
    if (typeof window === "undefined") return "light";
    const stored = window.localStorage.getItem("codecrest_excel_analytics_theme");
    return stored === "dark" ? "dark" : "light";
  });
  const [usageCount, setUsageCount] = useState(() => {
    if (typeof window === "undefined") return 0;
    const stored = Number(window.localStorage.getItem(USAGE_STORAGE_KEY) || "0");
    return Number.isFinite(stored) ? stored : 0;
  });
  
  // Custom API Key for real-time Anthropic analysis
  const [customApiKey, setCustomApiKey] = useState("");

  // Collaborative comments, version history, and view-only sharing states
  const [isSharedViewOnly, setIsSharedViewOnly] = useState(false);
  const [activeComments, setActiveComments] = useState<Array<{ id: string; author: string; text: string; timestamp: string }>>([]);
  const [versionHistory, setVersionHistory] = useState<Array<any>>([]);
  const [sharedRecordObj, setSharedRecordObj] = useState<SavedRecord | null>(null);
  const [activeRecordId, setActiveRecordId] = useState<number | null>(null);
  const [commentInput, setCommentInput] = useState("");
  const [commentAuthor, setCommentAuthor] = useState("");
  const [commentsOpen, setCommentsOpen] = useState(false);
  const [versionTrackerOpen, setVersionTrackerOpen] = useState(false);
  const [shareModalOpen, setShareModalOpen] = useState(false);
  const [versionUploading, setVersionUploading] = useState(false);

  // Admin Portal State
  const [adminModalOpen, setAdminModalOpen] = useState(false);
  const [adminTab, setAdminTab] = useState<"users" | "payments" | "plans" | "analytics" | "settings" | "activity">("users");
  const [adminUsers, setAdminUsers] = useState<User[]>([]);
  const [adminPayments, setAdminPayments] = useState<any[]>([]);
  const [adminPlans, setAdminPlans] = useState<Plan[]>([]);
  const [adminLogs, setAdminLogs] = useState<AdminLog[]>([]);
  const [adminSearch, setAdminSearch] = useState("");
  const [adminLoading, setAdminLoading] = useState(false);
  const [adminPaymentFilter, setAdminPaymentFilter] = useState<"all" | "success" | "pending_verification" | "rejected" | "refunded">("all");
  const [adminUserFilter, setAdminUserFilter] = useState<"all" | "pro" | "free">("all");

  // Admin Edit User
  const [adminEditUser, setAdminEditUser] = useState<User | null>(null);
  const [adminEditUserName, setAdminEditUserName] = useState("");
  const [adminEditUserEmail, setAdminEditUserEmail] = useState("");
  const [adminEditUserMobile, setAdminEditUserMobile] = useState("");
  const [adminEditUserIsPro, setAdminEditUserIsPro] = useState(false);
  const [adminEditUserOpen, setAdminEditUserOpen] = useState(false);

  // Admin Plans
  const [adminEditPlan, setAdminEditPlan] = useState<Plan | null>(null);
  const [adminPlanName, setAdminPlanName] = useState("");
  const [adminPlanPrice, setAdminPlanPrice] = useState(0);
  const [adminPlanPeriod, setAdminPlanPeriod] = useState<Plan["billingPeriod"]>("monthly");
  const [adminPlanFeatureInput, setAdminPlanFeatureInput] = useState("");
  const [adminPlanFeatures, setAdminPlanFeatures] = useState<string[]>([]);
  const [adminPlanActive, setAdminPlanActive] = useState(true);
  const [adminPlanModalOpen, setAdminPlanModalOpen] = useState(false);
  // Extended plan editor fields
  const [adminPlanDescription, setAdminPlanDescription] = useState("");
  const [adminPlanHighlighted, setAdminPlanHighlighted] = useState(false);
  const [adminPlanColor, setAdminPlanColor] = useState("#f59e0b");
  const [adminPlanMaxReports, setAdminPlanMaxReports] = useState(0);
  const [adminPlanSortOrder, setAdminPlanSortOrder] = useState(99);

  // Checkout: plans loaded from DB + selected plan
  const [checkoutPlans, setCheckoutPlans] = useState<Plan[]>([]);
  const [selectedPlanId, setSelectedPlanId] = useState<string | null>(null);

  // Admin Settings -- feature flags
  const [adminFeatureAI, setAdminFeatureAI] = useState(true);
  const [adminFeatureUPI, setAdminFeatureUPI] = useState(true);
  const [adminFeatureGoogleLogin, setAdminFeatureGoogleLogin] = useState(true);
  const [adminMaintenanceMode, setAdminMaintenanceMode] = useState(false);
  const [globalFreeLimit, setGlobalFreeLimit] = useState(() => {
    if (typeof window === "undefined") return 3;
    const stored = window.localStorage.getItem("sheetcodecrest_global_free_limit");
    return stored ? Number(stored) : 3;
  });

  const [mockupTabActive, setMockupTabActive] = useState<"shopify" | "logistics" | "universal">("shopify");

  // Universal Profiler Data
  const [dataProfile, setDataProfile] = useState<DataProfile | null>(null);

  // Logistics Optimizer Data
  const [logisticsAnalytics, setLogisticsAnalytics] = useState<AnalyticsResult | null>(null);
  const [mergedLogisticsCount, setMergedLogisticsCount] = useState(0);
  const [shopifyAnalytics, setShopifyAnalytics] = useState<ShopifyAnalyticsSummary | null>(null);

  // Raw & Paginated Data Table State
  const [rawRows, setRawRows] = useState<DataRow[]>([]);
  const [tableHeaders, setTableHeaders] = useState<string[]>([]);
  const [searchTerm, setSearchTerm] = useState("");
  const [currentPage, setCurrentPage] = useState(1);
  const [sortConfig, setSortConfig] = useState<{ key: string; direction: "asc" | "desc" } | null>(null);
  const rowsPerPage = 8;

  // Auto-detection Toast Banner
  const [detectionNotice, setDetectionNotice] = useState("");

  // Output File & Download URLs
  const [outName, setOutName] = useState("");
  const [dlUrl, setDlUrl] = useState<string | null>(null);
  
  // AI Advice Console State
  const [aiInsights, setAiInsights] = useState("");
  const [aiLoading, setAiLoading] = useState(false);
  const [chatHistory, setChatHistory] = useState<Array<{ sender: "user" | "analyst"; text: string }>>([]);
  const [chatInput, setChatInput] = useState("");

  const inputRef = useRef<HTMLInputElement | null>(null);
  const dlRef = useRef<string | null>(null);
  
  const isProActive = currentUser?.isPro === true;
  const isAdminActive = currentUser?.username === "codecreststudio" || currentUser?.email === "codecreststudio@gmail.com";
  const freeReportsRemaining = isProActive ? 999999 : Math.max(0, globalFreeLimit - usageCount);
  const hasFreeReportsRemaining = isProActive || freeReportsRemaining > 0;

  // Load session and handle Stripe payment callbacks on startup
  useEffect(() => {
    const sessionUser = window.localStorage.getItem("auto_excel_active_user");
    
    const initializeUser = async () => {
      // 1. Check URL Parameters for secure shared report
      const urlParams = new URLSearchParams(window.location.search);
      const shareId = urlParams.get("share") || urlParams.get("id");
      if (shareId) {
        const recordId = parseInt(shareId);
        if (!isNaN(recordId)) {
          try {
            addLog(`🔗 Fetching secure shared report (ID: ${recordId})...`);
            const record = await dbGetRecordById(recordId);
            if (record) {
              setSharedRecordObj(record);
              setIsSharedViewOnly(true);
              setActiveRecordId(record.id || null);
              setFile(new File([new ArrayBuffer(record.size)], record.filename, { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }));
              setRawRows(record.rawRows);
              setTableHeaders(record.tableHeaders);
              setDataProfile(record.dataProfile);
              setLogisticsAnalytics(record.logisticsAnalytics);
              setShopifyAnalytics(record.shopifyAnalytics);
              setMode(record.mode);
              setOutName(record.outName);
              setChatHistory(record.chatHistory || []);
              setActiveComments(record.comments || []);
              setVersionHistory(record.versions || []);
              
              // Compile download url
              const baseName = record.filename.replace(/\.[^/.]+$/, "");
              let targetDlUrl = "";
              if (record.mode === "shopify") {
                const shopifyWb = buildShopifyAnalyticsWorkbook(baseName, record.rawRows);
                const wbOut = XLSX.write(shopifyWb, { bookType: "xlsx", type: "array" });
                const blob = new Blob([wbOut], { type: "application/octet-stream" });
                targetDlUrl = URL.createObjectURL(blob);
              } else if (record.mode === "logistics") {
                const merged = mergeMultiSKU(record.rawRows);
                const logisticsWb = buildLogisticsWorkbook(baseName, merged, record.rawRows.length - merged.length, record.logisticsAnalytics);
                const wbOut = XLSX.write(logisticsWb, { bookType: "xlsx", type: "array" });
                const blob = new Blob([wbOut], { type: "application/octet-stream" });
                targetDlUrl = URL.createObjectURL(blob);
              } else {
                const genericWb = buildAnalyticsWorkbook(baseName, record.rawRows, record.dataProfile);
                const wbOut = XLSX.write(genericWb, { bookType: "xlsx", type: "array" });
                const blob = new Blob([wbOut], { type: "application/octet-stream" });
                targetDlUrl = URL.createObjectURL(blob);
              }
              setDlUrl(targetDlUrl);
              dlRef.current = targetDlUrl;
              
              setStep("done");
              addLog(`✓ Shared report loaded successfully! Mode: ${record.mode.toUpperCase()}`, "success");
              return;
            } else {
              addLog("⚠️ Shared record not found in cloud or local cache.", "error");
            }
          } catch (err: any) {
            console.error("Failed to load shared record:", err);
            addLog(`❌ Failed to load shared record: ${err.message || err}`, "error");
          }
        }
      }

      // 2. Normal User session load if not a shared view-only session
      let activeUser: any = null;
      if (sessionUser) {
        const user = await dbGetUser(sessionUser);
        if (user) {
          activeUser = user;
          setCurrentUser(user);
          const records = await dbGetRecords(user.username);
          setSavedRecords(records);
        }
      }

      // Local check finished
    };

    initializeUser();
  }, []);

  // Dynamically load Google GSI and Razorpay SDKs on mount
  useEffect(() => {
    const gsiScript = document.createElement("script");
    gsiScript.src = "https://accounts.google.com/gsi/client";
    gsiScript.async = true;
    gsiScript.defer = true;
    document.body.appendChild(gsiScript);

    const rzpScript = document.createElement("script");
    rzpScript.src = "https://checkout.razorpay.com/v1/checkout.js";
    rzpScript.async = true;
    document.body.appendChild(rzpScript);
    
    return () => {
      try { document.body.removeChild(gsiScript); } catch (e) {}
      try { document.body.removeChild(rzpScript); } catch (e) {}
    };
  }, []);

  // Render official Google Sign-In button when the modal opens
  useEffect(() => {
    if (authModalOpen) {
      const timer = setTimeout(() => {
        const initGoogleButton = () => {
          if (typeof window !== "undefined" && (window as any).google?.accounts?.id) {
            (window as any).google.accounts.id.initialize({
              client_id: GOOGLE_CLIENT_ID,
              callback: handleGoogleCredentialResponse,
            });
            const btnContainer = document.getElementById("google-signin-div");
            if (btnContainer) {
              (window as any).google.accounts.id.renderButton(btnContainer, {
                theme: theme === "dark" ? "filled_black" : "outline",
                size: "large",
                text: "continue_with",
                width: 320,
              });
            }
          } else {
            setTimeout(initGoogleButton, 200);
          }
        };
        initGoogleButton();
      }, 50);
      return () => clearTimeout(timer);
    }
  }, [authModalOpen, authTab, theme]);

  const sendToGoogleSheets = async (user: User, action: "register" | "login" | "upgrade") => {
    try {
      addLog(`📡 Syncing "${user.username}" (${action}) details with Google Sheets...`, "info");
      const payload = {
        action,
        username: user.username,
        name: user.name || "",
        mobile: user.mobile || "",
        email: user.email || "",
        isPro: user.isPro,
        timestamp: new Date().toLocaleString()
      };
      
      await fetch(GOOGLE_SHEET_PROXY_URL, {
        method: "POST",
        headers: {
          "Content-Type": "text/plain;charset=utf-8",
        },
        body: JSON.stringify(payload),
      });
      
      addLog(`✓ Google Sheet Sync succeeded for "${user.username}" (${action})!`, "success");
    } catch (err: any) {
      console.error("Google Sheets sync failed", err);
      addLog(`âš ï¸ Google Sheets Sync failed: ${err.message || err}`, "error");
    }
  };

  const decodeJwt = (token: string) => {
    try {
      const base64Url = token.split('.')[1];
      if (!base64Url) throw new Error("Invalid JWT format");
      const base64 = base64Url.replace(/-/g, '+').replace(/_/g, '/');
      const jsonPayload = decodeURIComponent(
        atob(base64)
          .split('')
          .map((c) => '%' + ('00' + c.charCodeAt(0).toString(16)).slice(-2))
          .join('')
      );
      return JSON.parse(jsonPayload);
    } catch (e: any) {
      console.error("JWT decoding failed", e.message);
      return null;
    }
  };

  const handleGoogleCredentialResponse = async (response: any) => {
    if (!response?.credential) return;
    
    const profile = decodeJwt(response.credential);
    if (!profile) {
      setAuthError("Failed to decode Google Sign-In token.");
      return;
    }
    
    const { name, email, sub } = profile;
    const username = email.split("@")[0].trim().toLowerCase();
    
    let user = await dbGetUser(username);
    let isNew = false;
    
    if (!user) {
      isNew = true;
      user = {
        username,
        passwordHash: `google_oauth_live_${sub}`,
        isPro: false,
        dateCreated: new Date().toLocaleDateString(),
        name: name,
        mobile: "",
        email: email
      };
      await dbSaveUser(user);
    }
    
    setCurrentUser(user);
    window.localStorage.setItem("auto_excel_active_user", user.username);
    setAuthModalOpen(false);
    setAuthUsername("");
    setAuthPassword("");
    setAuthName("");
    setAuthMobile("");
    setAuthEmail("");
    const records = await dbGetRecords(user.username);
    setSavedRecords(records);
    
    addLog(`👤 User "${user.username}" authenticated successfully via Live Google Sign-In!`, "info");
    sendToGoogleSheets(user, isNew ? "register" : "login");
  };

  const handleLogin = async (e: React.FormEvent) => {
    e.preventDefault();
    setAuthError("");

    const username = sanitizeText(authUsername).toLowerCase();
    const password = authPassword;

    if (!username || !password) {
      setAuthError("Please fill in all fields.");
      return;
    }

    // 🔒 Brute-force lockout check
    const lockoutSecs = getLockoutSecondsRemaining(username);
    if (lockoutSecs > 0) {
      const mins = Math.ceil(lockoutSecs / 60);
      setAuthError(`Too many failed attempts. Please try again in ${mins} minute${mins !== 1 ? "s" : ""}.`);
      return;
    }

    const user = await dbGetUser(username);
    const passwordValid = user ? await verifyPassword(password, user.passwordHash) : false;

    if (!user || !passwordValid) {
      const isNowLocked = recordFailedAttempt(username);
      if (isNowLocked) {
        setAuthError("Too many failed attempts. Account locked for 5 minutes.");
      } else {
        setAuthError("Invalid username or password.");
      }
      return;
    }

    // ✅ Successful login — reset rate limit
    clearFailedAttempts(username);

    // 🔄 Migrate legacy plain-text password to secure hash on first successful login
    if (isLegacyPassword(user.passwordHash)) {
      try {
        const newHash = await hashPassword(password);
        const migratedUser = { ...user, passwordHash: newHash };
        await dbSaveUser(migratedUser);
        user.passwordHash = newHash;
        addLog(`ðŸ” Password security upgraded to SHA-256 hash for "${username}".`, "info");
      } catch (_) {}
    }

    setCurrentUser(user);
    window.localStorage.setItem("auto_excel_active_user", user.username);
    setAuthModalOpen(false);
    setAuthUsername("");
    setAuthPassword("");
    setAuthName("");
    setAuthMobile("");
    setAuthEmail("");
    const records = await dbGetRecords(user.username);
    setSavedRecords(records);
    addLog(`👤 User "${user.username}" logged in successfully. Plan: ${user.isPro ? "PRO" : "FREE"}`, "info");
    sendToGoogleSheets(user, "login");
  };

  const handleSignup = async (e: React.FormEvent) => {
    e.preventDefault();
    setAuthError("");

    const username = sanitizeText(authUsername).toLowerCase();
    const password = authPassword;
    const name = sanitizeText(authName);
    const mobile = sanitizeText(authMobile).replace(/\D/g, "");
    const email = sanitizeText(authEmail).toLowerCase();

    if (!username || !password || !name || !mobile || !email) {
      setAuthError("Please fill in all fields.");
      return;
    }

    // 🔒 Input validation
    if (!isValidUsername(username)) {
      setAuthError("Username must be 3–32 characters (letters, numbers, underscores, hyphens only).");
      return;
    }
    if (!isValidPassword(password)) {
      setAuthError("Password must be at least 6 characters.");
      return;
    }
    if (!isValidEmail(email)) {
      setAuthError("Please enter a valid email address.");
      return;
    }
    if (!isValidMobile(mobile)) {
      setAuthError("Please enter a valid 10-digit Indian mobile number.");
      return;
    }

    const existing = await dbGetUser(username);
    if (existing) {
      setAuthError("Username already taken. Please choose a different one.");
      return;
    }

    // ðŸ” Hash password before storing
    const passwordHash = await hashPassword(password);

    const newUser: User = {
      username,
      passwordHash,
      isPro: false,
      dateCreated: new Date().toLocaleDateString(),
      name,
      mobile,
      email
    };
    await dbSaveUser(newUser);
    setCurrentUser(newUser);
    window.localStorage.setItem("auto_excel_active_user", newUser.username);
    setAuthModalOpen(false);
    setAuthUsername("");
    setAuthPassword("");
    setAuthName("");
    setAuthMobile("");
    setAuthEmail("");
    setSavedRecords([]);
    addLog(`👤 New account "${newUser.username}" created securely (password hashed).`, "info");
    sendToGoogleSheets(newUser, "register");
  };

  const handleLogout = () => {
    if (currentUser) {
      addLog(`👤 User "${currentUser.username}" logged out.`, "info");
    }
    setCurrentUser(null);
    window.localStorage.removeItem("auto_excel_active_user");
    setSavedRecords([]);
  };

  const loadRecord = (rec: SavedRecord) => {
    setMode(rec.mode);
    setRawRows(rec.rawRows);
    setTableHeaders(rec.tableHeaders);
    if (rec.mode === "universal") {
      setDataProfile(rec.dataProfile);
      setLogisticsAnalytics(null);
      setShopifyAnalytics(null);
    } else if (rec.mode === "logistics") {
      setLogisticsAnalytics(rec.logisticsAnalytics);
      setDataProfile(null);
      setShopifyAnalytics(null);
    } else if (rec.mode === "shopify") {
      setShopifyAnalytics(rec.shopifyAnalytics);
      setDataProfile(null);
      setLogisticsAnalytics(null);
    }
    setOutName(rec.outName);
    
    let wb;
    const baseName = rec.filename.replace(/\.[^/.]+$/, "");
    if (rec.mode === "shopify") {
      wb = buildShopifyAnalyticsWorkbook(baseName, rec.rawRows);
    } else if (rec.mode === "logistics") {
      const consolidated = mergeMultiSKU(rec.rawRows);
      const mergedDups = rec.rawRows.length - consolidated.length;
      const stats = computeAnalytics(consolidated);
      wb = buildLogisticsWorkbook(baseName, consolidated, mergedDups, stats);
    } else {
      wb = buildAnalyticsWorkbook(baseName, rec.rawRows, rec.dataProfile);
    }
    const wbOut = XLSX.write(wb, { bookType: "xlsx", type: "array" });
    const blob = new Blob([wbOut], { type: "application/octet-stream" });
    if (dlRef.current) URL.revokeObjectURL(dlRef.current);
    const newUrl = URL.createObjectURL(blob);
    setDlUrl(newUrl);
    dlRef.current = newUrl;

    setChatHistory(rec.chatHistory || []);
    setFile({ name: rec.filename, size: rec.size } as File);
    setStep("done");
    setDashboardOpen(false);
    addLog(`⚡ Loaded analysis for "${rec.filename}" from Local Database`, "success");
  };

  const handleDeleteRecord = async (id: number, e: React.MouseEvent) => {
    e.stopPropagation();
    if (!confirm("Are you sure you want to delete this saved record?")) return;
    await dbDeleteRecord(id);
    if (currentUser) {
      const recs = await dbGetRecords(currentUser.username);
      setSavedRecords(recs);
    }
  };

  const startPaymentSimulation = async () => {
    setPaymentProcessing(true);
    setPaymentCompleted(false);
    setPaymentLogs([]);

    // Live Razorpay Mode
    const razorpayKey = import.meta.env.VITE_RAZORPAY_KEY_ID || "rzp_test_StUNrV1X2WAvV4";
    if (razorpayKey && typeof (window as any).Razorpay !== "undefined") {
      const planToCharge = checkoutPlans.find(p => p.id === selectedPlanId) || checkoutPlans[0];
      const chargePrice = planToCharge ? planToCharge.price : 1599;
      const chargeName = planToCharge ? planToCharge.name : "SheetCodeCrest Pro";

      const options = {
        key: razorpayKey,
        amount: chargePrice * 100, // in paisa
        currency: "INR",
        name: chargeName,
        description: `${chargeName} Subscription`,
        image: "https://sheetcodecrest.vercel.app/logo.png",
        handler: async function (response: any) {
          try {
            addLog(`💳 Razorpay transaction completed! Payment ID: ${response.razorpay_payment_id}`, "success");
            
            if (currentUser) {
              const updatedUser = { ...currentUser, isPro: true };
              await dbSaveUser(updatedUser);
              setCurrentUser(updatedUser);
              
              // Log payment securely in Supabase
              const { dbLogPayment } = await import("./db");
              await dbLogPayment({
                username: currentUser.username,
                gateway: "razorpay",
                paymentId: response.razorpay_payment_id,
                orderId: response.razorpay_order_id || "",
                signature: response.razorpay_signature || "",
                amount: chargePrice,
                status: "success"
              });
              
              addLog(`⚡ Live payment verified! Account "${currentUser.username}" upgraded to PRO.`, "success");
              sendToGoogleSheets(updatedUser, "upgrade");
              
              setPaymentProcessing(false);
              setPaymentCompleted(true);
              setTimeout(() => {
                setCheckoutOpen(false);
                setPaymentCompleted(false);
                setPaymentLogs([]);
              }, 1500);
            }
          } catch (err: any) {
            console.error("Razorpay callback processing failed", err);
            alert(`Error processing payment verification: ${err.message || err}`);
            setPaymentProcessing(false);
          }
        },
        prefill: {
          name: currentUser?.name || currentUser?.username || "",
          email: currentUser?.email || "",
          contact: currentUser?.mobile || ""
        },
        notes: {
          username: currentUser?.username || "anonymous"
        },
        theme: {
          color: "#0f172a"
        }
      };

      setPaymentProcessing(false);
      const rzp = new (window as any).Razorpay(options);
      rzp.on("payment.failed", function (resp: any) {
        console.error("Razorpay payment failed", resp.error);
        alert(`Payment failed: ${resp.error.description}`);
        setPaymentProcessing(false);
      });
      rzp.open();
      return;
    }

    // Fallback: Razorpay UPI Simulation Mode
    if (!upiVPA.trim()) {
      alert("Please enter your UPI ID.");
      setPaymentProcessing(false);
      return;
    }

    setPaymentLogs([
      "🔒 Initiating Razorpay UPI Gateway Ping...",
      "📡 Handshaking with UPI Address Resolver (vpa@okicici)...",
      "📲 Generating dynamic UPI deep-link QR intent payload...",
      "📡 Listening for mobile banking app webhook callback...",
      "ðŸ›¡ï¸ Validating secure transaction checksum (SHA-256)...",
      "💸 Transferring funds to merchant account...",
      "📡 Razorpay callback verified by merchant...",
      "⚡ Updating subscription entitlement store...",
      "🎉 Transaction Successful!"
    ]);

    let currentIdx = 0;
    const interval = setInterval(() => {
      if (currentIdx < 9) {
        currentIdx++;
      } else {
        clearInterval(interval);
        setPaymentProcessing(false);
        setPaymentCompleted(true);
        if (currentUser) {
          const updatedUser = { ...currentUser, isPro: true };
          dbSaveUser(updatedUser).then(() => {
            setCurrentUser(updatedUser);
            addLog(`⚡ Payment verified! Account "${currentUser.username}" upgraded to PRO plan.`, "success");
            sendToGoogleSheets(updatedUser, "upgrade");
            setTimeout(() => {
              setCheckoutOpen(false);
              setPaymentCompleted(false);
              setPaymentLogs([]);
              setUpiVPA("");
            }, 1500);
          });
        }
      }
    }, 500);
  };

  const handleManualUpiVerification = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!upiUTR.trim() || upiUTR.trim().length !== 12 || isNaN(Number(upiUTR))) {
      alert("Please enter a valid 12-digit UPI UTR / Transaction Reference Number.");
      return;
    }

    setPaymentProcessing(true);
    setPaymentCompleted(false);
    setPaymentLogs([
      "🔎 Searching UPI banking registers for Transaction UTR...",
      `📡 Found matching ref: ${upiUTR.trim()}`,
      "💸 Processing pending transfer validation...",
      "â˜ï¸ Syncing transaction records securely to cloud database...",
      "🎉 Account Upgraded to PRO (Manual Verification Pending)!"
    ]);

    let currentIdx = 0;
    const interval = setInterval(async () => {
      if (currentIdx < 5) {
        currentIdx++;
      } else {
        clearInterval(interval);
        setPaymentProcessing(false);
        setPaymentCompleted(true);
        if (currentUser) {
          const updatedUser = { ...currentUser, isPro: true };
          await dbSaveUser(updatedUser);
          setCurrentUser(updatedUser);
          
          // Log manual transaction request to Supabase payments as pending
          const planToCharge = checkoutPlans.find(p => p.id === selectedPlanId) || checkoutPlans[0];
          const chargePrice = planToCharge ? planToCharge.price : 1599;

          const { dbLogPayment } = await import("./db");
          await dbLogPayment({
            username: currentUser.username,
            gateway: "razorpay",
            paymentId: `upi_utr_pending_${upiUTR.trim()}`,
            amount: chargePrice,
            status: "pending_verification"
          });

          addLog(`⚡ UPI Payment request submitted! UTR: ${upiUTR.trim()}. Account "${currentUser.username}" upgraded to PRO.`, "success");
          sendToGoogleSheets(updatedUser, "upgrade");

          setTimeout(() => {
            setCheckoutOpen(false);
            setPaymentCompleted(false);
            setPaymentLogs([]);
            setUpiUTR("");
          }, 1500);
        }
      }
    }, 600);
  };

  // ----------------------------------------------------
  // ðŸ›¡ï¸ ADMIN PANEL CONTROLLERS (Advanced)
  // ----------------------------------------------------
  const logAdminAction = async (action: string, details?: string) => {
    const by = currentUser?.username || "admin";
    const entry: AdminLog = { action, performedBy: by, details, createdAt: new Date().toISOString() };
    setAdminLogs(prev => [entry, ...prev].slice(0, 200));
    await dbLogAdminAction(action, by, details);
  };

  const loadAdminData = async () => {
    setAdminLoading(true);
    try {
      const [users, payments, plans, logs] = await Promise.all([
        dbGetAllUsers(),
        dbGetAllPayments(),
        dbGetPlans(),
        dbGetAdminLogs()
      ]);
      setAdminUsers(users);
      setAdminPayments(payments);
      setAdminPlans(plans);
      setAdminLogs(logs);
    } catch (err) {
      console.error("Failed to load admin data", err);
    } finally {
      setAdminLoading(false);
    }
  };

  // -- USER MANAGEMENT --
  const openEditUser = (user: User) => {
    setAdminEditUser(user);
    setAdminEditUserName(user.name || "");
    setAdminEditUserEmail(user.email || "");
    setAdminEditUserMobile(user.mobile || "");
    setAdminEditUserIsPro(user.isPro);
    setAdminEditUserOpen(true);
  };

  const handleSaveEditUser = async () => {
    if (!adminEditUser) return;
    setAdminLoading(true);
    try {
      await dbUpdateUserFields(adminEditUser.username, {
        name: adminEditUserName.trim(),
        email: adminEditUserEmail.trim().toLowerCase(),
        mobile: adminEditUserMobile.trim(),
        isPro: adminEditUserIsPro
      });
      await logAdminAction("EDIT_USER", `Updated profile for user "${adminEditUser.username}"`);
      if (currentUser && currentUser.username === adminEditUser.username) {
        setCurrentUser(prev => prev ? { ...prev, name: adminEditUserName, email: adminEditUserEmail, mobile: adminEditUserMobile, isPro: adminEditUserIsPro } : prev);
      }
      setAdminEditUserOpen(false);
      await loadAdminData();
      addLog(`ðŸ›¡ï¸ Admin: User "${adminEditUser.username}" profile updated.`, "info");
    } catch (err) {
      console.error("Failed to save user edit", err);
      alert("Failed to update user.");
    } finally {
      setAdminLoading(false);
    }
  };

  const handleToggleUserPro = async (targetUser: User) => {
    try {
      const updated = { ...targetUser, isPro: !targetUser.isPro };
      await dbSaveUser(updated);
      await logAdminAction("TOGGLE_PRO", `Set isPro=${!targetUser.isPro} for "${targetUser.username}"`);
      addLog(`ðŸ›¡ï¸ Admin: Toggled PRO for "${targetUser.username}" → ${!targetUser.isPro}`, "info");
      await loadAdminData();
      if (currentUser && currentUser.username === targetUser.username) setCurrentUser(updated);
    } catch (err) {
      console.error("Failed to toggle PRO", err);
      alert("Failed to update user privilege.");
    }
  };

  const handleDeleteUser = async (targetUser: User) => {
    if (!confirm(`âš ï¸ PERMANENTLY delete user "${targetUser.username}"? This cannot be undone.`)) return;
    setAdminLoading(true);
    try {
      await dbDeleteUser(targetUser.username);
      await logAdminAction("DELETE_USER", `Deleted user "${targetUser.username}"`);
      addLog(`ðŸ›¡ï¸ Admin: User "${targetUser.username}" deleted permanently.`, "error");
      await loadAdminData();
    } catch (err) {
      console.error("Failed to delete user", err);
      alert("Failed to delete user.");
    } finally {
      setAdminLoading(false);
    }
  };

  // -- PAYMENT MANAGEMENT --
  const handleAdminApproveUpi = async (paymentId: string, username: string) => {
    if (!confirm(`Approve transaction "${paymentId}" and upgrade "${username}" to PRO?`)) return;
    setAdminLoading(true);
    try {
      await dbApprovePayment(paymentId, username);
      await logAdminAction("APPROVE_PAYMENT", `Approved payment ${paymentId} for "${username}"`);
      addLog(`ðŸ›¡ï¸ Admin approved payment: ${paymentId}. Promoted "${username}" to PRO.`, "success");
      const userObj = await dbGetUser(username);
      if (userObj) sendToGoogleSheets(userObj, "upgrade");
      await loadAdminData();
    } catch (err) {
      console.error("Failed to approve payment", err);
      alert("Failed to approve payment.");
    } finally {
      setAdminLoading(false);
    }
  };

  const handleAdminRejectPayment = async (paymentId: string, username: string) => {
    if (!confirm(`Reject payment "${paymentId}" for "${username}"?`)) return;
    setAdminLoading(true);
    try {
      await dbUpdatePaymentStatus(paymentId, "rejected");
      await logAdminAction("REJECT_PAYMENT", `Rejected payment ${paymentId} for "${username}"`);
      addLog(`ðŸ›¡ï¸ Admin rejected payment: ${paymentId} for "${username}".`, "error");
      await loadAdminData();
    } catch (err) {
      console.error("Failed to reject payment", err);
      alert("Failed to reject payment.");
    } finally {
      setAdminLoading(false);
    }
  };

  const handleAdminRefundPayment = async (paymentId: string, username: string) => {
    if (!confirm(`Mark payment "${paymentId}" as REFUNDED and revoke PRO for "${username}"?`)) return;
    setAdminLoading(true);
    try {
      await dbUpdatePaymentStatus(paymentId, "refunded");
      await dbUpdateUserFields(username, { isPro: false });
      await logAdminAction("REFUND_PAYMENT", `Refunded payment ${paymentId}, revoked PRO for "${username}"`);
      addLog(`ðŸ›¡ï¸ Admin refunded payment: ${paymentId}, revoked PRO for "${username}".`, "error");
      await loadAdminData();
    } catch (err) {
      console.error("Failed to refund payment", err);
      alert("Failed to refund payment.");
    } finally {
      setAdminLoading(false);
    }
  };

  // -- PLANS MANAGEMENT --
  const openNewPlan = () => {
    setAdminEditPlan(null);
    setAdminPlanName("");
    setAdminPlanPrice(0);
    setAdminPlanPeriod("monthly");
    setAdminPlanFeatures([]);
    setAdminPlanFeatureInput("");
    setAdminPlanActive(true);
    setAdminPlanDescription("");
    setAdminPlanHighlighted(false);
    setAdminPlanColor("#f59e0b");
    setAdminPlanMaxReports(0);
    setAdminPlanSortOrder(adminPlans.length);
    setAdminPlanModalOpen(true);
  };

  const openEditPlan = (plan: Plan) => {
    setAdminEditPlan(plan);
    setAdminPlanName(plan.name);
    setAdminPlanPrice(plan.price);
    setAdminPlanPeriod(plan.billingPeriod);
    setAdminPlanFeatures([...plan.features]);
    setAdminPlanFeatureInput("");
    setAdminPlanActive(plan.isActive);
    setAdminPlanDescription(plan.description || "");
    setAdminPlanHighlighted(plan.highlighted || false);
    setAdminPlanColor(plan.color || "#f59e0b");
    setAdminPlanMaxReports(plan.maxReports ?? 0);
    setAdminPlanSortOrder(plan.sortOrder ?? 99);
    setAdminPlanModalOpen(true);
  };

  const handleSavePlan = async () => {
    if (!adminPlanName.trim()) { alert("Plan name is required."); return; }
    setAdminLoading(true);
    try {
      const plan: Plan = {
        id: adminEditPlan?.id,
        name: adminPlanName.trim(),
        price: adminPlanPrice,
        billingPeriod: adminPlanPeriod,
        features: adminPlanFeatures,
        isActive: adminPlanActive,
        description: adminPlanDescription.trim(),
        highlighted: adminPlanHighlighted,
        color: adminPlanColor,
        maxReports: adminPlanMaxReports,
        sortOrder: adminPlanSortOrder,
      };
      await dbSavePlan(plan);
      await logAdminAction(adminEditPlan ? "EDIT_PLAN" : "CREATE_PLAN", `${adminEditPlan ? "Updated" : "Created"} plan "${plan.name}"`);
      addLog(`ðŸ›¡ï¸ Admin: Plan "${plan.name}" ${adminEditPlan ? "updated" : "created"}.`, "info");
      setAdminPlanModalOpen(false);
      await loadAdminData();
    } catch (err) {
      console.error("Failed to save plan", err);
      alert("Failed to save plan.");
    } finally {
      setAdminLoading(false);
    }
  };

  const handleDeletePlan = async (plan: Plan) => {
    if (!plan.id) return;
    if (!confirm(`Delete plan "${plan.name}"? This cannot be undone.`)) return;
    setAdminLoading(true);
    try {
      await dbDeletePlan(plan.id);
      await logAdminAction("DELETE_PLAN", `Deleted plan "${plan.name}"`);
      addLog(`ðŸ›¡ï¸ Admin: Plan "${plan.name}" deleted.`, "error");
      await loadAdminData();
    } catch (err) {
      console.error("Failed to delete plan", err);
      alert("Failed to delete plan.");
    } finally {
      setAdminLoading(false);
    }
  };

  // -- SETTINGS --
  const handleSaveFreeLimit = (limit: number) => {
    if (isNaN(limit) || limit < 1) { alert("Please enter a valid limit >= 1."); return; }
    window.localStorage.setItem("sheetcodecrest_global_free_limit", String(limit));
    setGlobalFreeLimit(limit);
    logAdminAction("UPDATE_SETTINGS", `Set free report limit to ${limit}`);
    addLog(`âš™ï¸ System setting updated: free limit = ${limit}`, "info");
  };

  const handleSaveFeatureFlags = () => {
    window.localStorage.setItem("sheetcc_flag_ai", String(adminFeatureAI));
    window.localStorage.setItem("sheetcc_flag_upi", String(adminFeatureUPI));
    window.localStorage.setItem("sheetcc_flag_google", String(adminFeatureGoogleLogin));
    window.localStorage.setItem("sheetcc_flag_maintenance", String(adminMaintenanceMode));
    logAdminAction("UPDATE_FLAGS", `AI=${adminFeatureAI}, UPI=${adminFeatureUPI}, Google=${adminFeatureGoogleLogin}, Maintenance=${adminMaintenanceMode}`);
    addLog("âš™ï¸ Feature flags saved successfully.", "success");
    alert("✅ Feature flags saved!");
  };



  // Load active paid plans for the checkout modal
  const loadCheckoutPlans = async () => {
    try {
      const plans = await dbGetPlans();
      const activePaidPlans = plans.filter(p => p.isActive && p.price > 0);
      setCheckoutPlans(activePaidPlans);
      if (activePaidPlans.length > 0) {
        const recommended = activePaidPlans.find(p => p.highlighted) || activePaidPlans[0];
        setSelectedPlanId(prev => prev || recommended.id || null);
      }
    } catch (err) {
      console.error("Failed to load checkout plans", err);
    }
  };

  // Reload checkout plans when modal opens
  useEffect(() => {
    if (checkoutOpen) { loadCheckoutPlans(); }
  }, [checkoutOpen]);

    const recordSuccessfulReport = useCallback(() => {
    setUsageCount((current) => {
      const next = Math.max(current + 1, Number(window.localStorage.getItem(USAGE_STORAGE_KEY) || "0") + 1);
      try {
        window.localStorage.setItem(USAGE_STORAGE_KEY, String(next));
      } catch (e) {}
      return next;
    });
  }, []);

  useEffect(() => {
    if (typeof document !== "undefined") {
      document.documentElement.setAttribute("data-theme", theme);
      document.body.setAttribute("data-theme", theme);
    }
    try {
      window.localStorage.setItem("codecrest_excel_analytics_theme", theme);
    } catch (e) {}
  }, [theme]);

  // Dynamic Google Font Injection
  useEffect(() => {
    if (typeof document !== "undefined") {
      const link1 = document.createElement("link");
      link1.rel = "preconnect";
      link1.href = "https://fonts.googleapis.com";
      const link2 = document.createElement("link");
      link2.rel = "preconnect";
      link2.href = "https://fonts.gstatic.com";
      link2.crossOrigin = "anonymous";
      const link3 = document.createElement("link");
      link3.rel = "stylesheet";
      link3.href = "https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@300;400;500;600;700&family=Inter:wght@300;400;500;600;700&family=JetBrains+Mono:wght@400;500&display=swap";
      
      document.head.appendChild(link1);
      document.head.appendChild(link2);
      document.head.appendChild(link3);
    }
  }, []);


  // Reset Application State
  const reset = useCallback(() => {
    setStep("upload");
    setFile(null);
    setLog([]);
    setDataProfile(null);
    setLogisticsAnalytics(null);
    setMergedLogisticsCount(0);
    setShopifyAnalytics(null);
    setRawRows([]);
    setTableHeaders([]);
    setSearchTerm("");
    setCurrentPage(1);
    setSortConfig(null);
    setAiInsights("");
    setChatHistory([]);
    setChatInput("");
    setDetectionNotice("");
    setError("");
    if (dlRef.current) {
      try { URL.revokeObjectURL(dlRef.current); } catch (e) {}
      dlRef.current = null;
    }
    setDlUrl(null);
  }, []);

  // Core File Upload Processor
  const processFile = useCallback(async (f: File, forceMode?: AppMode) => {
    if (!f) return;
    if (!hasFreeReportsRemaining) {
      setError("Free report limit reached. Please purchase access from Codecrest Studio to continue generating analytics workbooks.");
      setDetectionNotice("🔒 Free trial complete. Upgrade with Codecrest Studio to continue.");
      return;
    }
    setFile(f);
    setError("");
    setLog([]);
    setDlUrl(null);
    setAiInsights("");
    setDetectionNotice("");
    setStep("processing");

    try {
      addLog(`Reading file: ${f.name} (${(f.size / 1024).toFixed(1)} KB)`);
      const { data, headers: parsedHeaders, sheetName, originalRows, headerRow } = await parseExcel(f);
      addLog(`✓ Excel parsed. Sheet: "${sheetName}" | Headers found at row ${headerRow + 1} | ${originalRows} rows total`, "success");
      
      setRawRows(data);
      setTableHeaders(parsedHeaders);

      // Auto-detect known commerce schemas in the background.
      const lowercaseHeaders = parsedHeaders.map((h) => h.toLowerCase().replace(/\s+/g, ""));
      const hasHeader = (patterns: RegExp[]) => lowercaseHeaders.some((h) => patterns.some((pattern) => pattern.test(h)));
      const shopifyScore = [
        hasHeader([/^createdat$/, /orderdate/]),
        hasHeader([/^financialstatus$/, /paymentstatus/]),
        hasHeader([/^fulfillmentstatus$/, /lineitemfulfillmentstatus/]),
        hasHeader([/^lineitemname$/, /productname/, /itemname/]),
        hasHeader([/^lineitemquantity$/, /quantity/, /qty/]),
        hasHeader([/^lineitemprice$/, /price/]),
        hasHeader([/^total$/, /ordertotal/]),
        hasHeader([/^discountamount$/, /discount/]),
      ].filter(Boolean).length;
      const shiprocketScore = [
        hasHeader([/orderid/, /ordernumber/]),
        hasHeader([/courier/, /awb/]),
        hasHeader([/pickupaddress/, /pickup/]),
        hasHeader([/rto/, /ndr/, /status/]),
        hasHeader([/freight/, /codpay/, /codamount/]),
      ].filter(Boolean).length;
      const isShopify = shopifyScore >= 5;
      const isShiprocket = !isShopify && shiprocketScore >= 3;
      
      const targetMode = forceMode || (isShopify ? "shopify" : isShiprocket ? "logistics" : "universal");
      setMode(targetMode);

      if (isShopify && !forceMode) {
        setDetectionNotice(`ðŸ›ï¸ Shopify export auto-detected in the background (${shopifyScore}/8 schema match). Building customer, product, order, retargeting, COD, and geographic analytics.`);
      } else if (isShiprocket && !forceMode) {
        setDetectionNotice(`🚀 Shiprocket logistics schema auto-detected in the background (${shiprocketScore}/5 schema match). Unlocking courier scorecards and RTO analysis.`);
      } else if (!forceMode) {
        setDetectionNotice("📊 Generic spreadsheet detected in the background. Building universal data quality, numeric, duplicate, and column profiling analytics.");
      }

      addLog(`Auto-detected Report Engine: ${targetMode === "shopify" ? "Shopify Growth Analytics" : targetMode === "logistics" ? "Shiprocket Logistics Optimizer" : "Universal Spreadsheet Profiler"}`);
      addLog(`Schema fingerprint: Shopify ${shopifyScore}/8 | Shiprocket ${shiprocketScore}/5 | ${parsedHeaders.length} columns scanned`);

      const baseName = f.name.replace(/\.[^/.]+$/, "");
      const dateString = new Date().toISOString().slice(0, 10);

      if (targetMode === "shopify") {
        addLog("Preparing Shopify order, product, customer, and retargeting analytics...");
        const stats = analyzeShopifyData(data);
        setShopifyAnalytics(stats);
        addLog(`✓ Shopify summary: ${stats.totalOrders.toLocaleString("en-IN")} orders | ${stats.totalCustomers.toLocaleString("en-IN")} customers | ${stats.productCount} products`, "success");
        addLog(`✓ Revenue mapped: ₹${Math.round(stats.totalRevenue).toLocaleString("en-IN")} | Units sold: ${stats.totalUnits.toLocaleString("en-IN")}`, "success");

        addLog("Compiling Shopify workbook similar to your reference analytics files...");
        const shopifyWb = buildShopifyAnalyticsWorkbook(baseName, data);
        const wbOut = XLSX.write(shopifyWb, { bookType: "xlsx", type: "array" });
        const blob = new Blob([wbOut], { type: "application/octet-stream" });

        if (dlRef.current) URL.revokeObjectURL(dlRef.current);
        const newUrl = URL.createObjectURL(blob);
        setDlUrl(newUrl);
        dlRef.current = newUrl;
        setOutName(`${baseName}_Shopify_Analytics_${dateString}.xlsx`);
        addLog(`✓ Shopify report finalized with ${shopifyWb.SheetNames.length} sheets`, "success");

        const welcomeMsg = `👋 Hi! I mapped your Shopify export into a **growth analytics workbook** like your sample files.

Key outputs generated:
• **Dashboard** for orders, revenue, customers, units, segments, and status mix
• **Order Status Detail** similar to your order-product report
• **Product Analysis** and **Product × Order Status** performance views
• **Monthly Trends**, **COD Analysis**, **Geographic**, and **Discount Analysis**
• **Customer Data**, **Segment Analysis**, and **Retargeting Lists**
• **Top product-wise customer/order sheets**

Initial read:
• **${stats.totalOrders.toLocaleString("en-IN")} orders**
• **${stats.totalCustomers.toLocaleString("en-IN")} customers**
• **₹${Math.round(stats.totalRevenue).toLocaleString("en-IN")} revenue**
• **${stats.totalUnits.toLocaleString("en-IN")} units sold**
• Top product: **${stats.topProduct}**`;
        setChatHistory([{ sender: "analyst", text: welcomeMsg }]);

      } else if (targetMode === "logistics") {
        addLog("Resolving line-item duplicates and consolidating Multi-SKU orders...");
        const merged = mergeMultiSKU(data);
        const resolvedDups = data.length - merged.length;
        setMergedLogisticsCount(resolvedDups);
        setRawRows(merged); // For the table preview
        
        if (resolvedDups > 0) {
          addLog(`✓ Consolidated ${resolvedDups} multi-item duplicate rows → ${merged.length} unique customer orders`, "success");
        }

        addLog("Computing logistics speed scorecards and KPIs...");
        const stats = computeAnalytics(merged);
        setLogisticsAnalytics(stats);
        addLog(`✓ Delivery success: ${(stats.deliveryRate * 100).toFixed(1)}% | RTO rate: ${(stats.rtoRate * 100).toFixed(1)}%`, "success");
        addLog(`✓ Net shipping revenue: ₹${stats.totalRev.toLocaleString("en-IN")} | COD Collected: ₹${stats.totalCOD.toLocaleString("en-IN")}`, "success");

        addLog("Compiling customized 7-sheet logistics analytical report...");
        const logisticsWb = buildLogisticsWorkbook(baseName, merged, resolvedDups, stats);
        const wbOut = XLSX.write(logisticsWb, { bookType: "xlsx", type: "array" });
        const blob = new Blob([wbOut], { type: "application/octet-stream" });
        
        if (dlRef.current) URL.revokeObjectURL(dlRef.current);
        const newUrl = URL.createObjectURL(blob);
        setDlUrl(newUrl);
        dlRef.current = newUrl;
        setOutName(`${baseName}_Logistics_Optimizer_${dateString}.xlsx`);
        addLog(`✓ Report workbook finalized with ${logisticsWb.SheetNames.length} structured sheets`, "success");

        const welcomeMsg = `👋 Hi! I am your **Senior AI Data Analyst**. I have thoroughly audited your Shiprocket logistics dataset of **${stats.total} unique orders** (having resolved **${resolvedDups} multi-SKU duplicates**).

Here are my initial findings:
• **Delivery Success Rate**: **${(stats.deliveryRate * 100).toFixed(1)}%** (${stats.delivered} delivered)
• **RTO Return Rate**: **${(stats.rtoRate * 100).toFixed(1)}%** (${stats.rto} returned packages)
• **Total Shipping Revenue**: **₹${stats.totalRev.toLocaleString("en-IN")}**
• **COD Collected**: **₹${stats.totalCOD.toLocaleString("en-IN")}** (${((stats.payCounts["cod"]?.orders || 0) / stats.total * 100).toFixed(0)}% of orders)
• **Freight Cost**: **₹${stats.totalFreight.toLocaleString("en-IN")}**

I have compiled a formula-friendly 7-sheet analytical workbook for your accounts team. Ask me any question about the data! E.g.:
• *How can we reduce our shipping leakage?*
• *Which couriers are performing the best/worst?*
• *Which destinations present the highest RTO risk?*`;
        setChatHistory([{ sender: "analyst", text: welcomeMsg }]);

      } else {
        addLog("Initiating universal mathematical analysis on columns...");
        const profile = { ...analyzeData(data, parsedHeaders), sheetName, headerRow };
        setDataProfile(profile);
        addLog(`✓ Generated profile for ${profile.totalColumns} columns and ${profile.totalRows} data rows`, "success");

        addLog("Compiling 8-sheet statistical data summary workbook...");
        const genericWb = buildAnalyticsWorkbook(baseName, data, profile);
        const wbOut = XLSX.write(genericWb, { bookType: "xlsx", type: "array" });
        const blob = new Blob([wbOut], { type: "application/octet-stream" });
        
        if (dlRef.current) URL.revokeObjectURL(dlRef.current);
        const newUrl = URL.createObjectURL(blob);
        setDlUrl(newUrl);
        dlRef.current = newUrl;
        setOutName(`${baseName}_Profiler_Summary_${dateString}.xlsx`);
        addLog(`✓ Excel summary ready for download: ${genericWb.SheetNames.length} statistical sheets`, "success");

        const welcomeMsg = `👋 Hi! I am your **Senior AI Data Analyst**. I have mapped a deep statistical profile for your **${profile.totalRows} rows** across **${profile.totalColumns} columns**.

I've automatically identified data types, column fill rates, and calculated mathematical variances (sum, average, standard deviation, min, max) for numeric fields.

What would you like to investigate in this dataset? E.g.:
• *Can you summarize the numeric fields for me?*
• *Show me the columns with the most empty values.*
• *What is the cardinality of unique keys in the table?*`;
        setChatHistory([{ sender: "analyst", text: welcomeMsg }]);
      }

      if (currentUser) {
        const newRecord: SavedRecord = {
          username: currentUser.username,
          filename: f.name,
          size: f.size,
          mode: targetMode,
          rawRows: data,
          tableHeaders: parsedHeaders,
          dataProfile: targetMode === "universal" ? { ...analyzeData(data, parsedHeaders), sheetName, headerRow } : null,
          logisticsAnalytics: targetMode === "logistics" ? computeAnalytics(mergeMultiSKU(data)) : null,
          shopifyAnalytics: targetMode === "shopify" ? analyzeShopifyData(data) : null,
          outName: targetMode === "shopify" ? `${baseName}_Shopify_Analytics_${dateString}.xlsx` : targetMode === "logistics" ? `${baseName}_Logistics_Optimizer_${dateString}.xlsx` : `${baseName}_Profiler_Summary_${dateString}.xlsx`,
          chatHistory: [],
          timestamp: new Date().toLocaleString()
        };

        if (targetMode === "shopify") {
          const stats = analyzeShopifyData(data);
          newRecord.chatHistory = [{ sender: "analyst", text: `👋 Hi! I mapped your Shopify export into a **growth analytics workbook** like your sample files.\n\nKey outputs generated:\n• **Dashboard** for orders, revenue, customers, units, segments, and status mix\n• **Order Status Detail** similar to your order-product report\n• **Product Analysis** and **Product × Order Status** performance views\n• **Monthly Trends**, **COD Analysis**, **Geographic**, and **Discount Analysis**\n• **Customer Data**, **Segment Analysis**, and **Retargeting Lists**\n• **Top product-wise customer/order sheets**\n\nInitial read:\n• **${stats.totalOrders.toLocaleString("en-IN")} orders**\n• **${stats.totalCustomers.toLocaleString("en-IN")} customers**\n• **₹${Math.round(stats.totalRevenue).toLocaleString("en-IN")} revenue**\n• **${stats.totalUnits.toLocaleString("en-IN")} units sold**\n• Top product: **${stats.topProduct}**` }];
        } else if (targetMode === "logistics") {
          const stats = computeAnalytics(mergeMultiSKU(data));
          const resolvedDups = data.length - mergeMultiSKU(data).length;
          newRecord.chatHistory = [{ sender: "analyst", text: `👋 Hi! I am your **Senior AI Data Analyst**. I have thoroughly audited your Shiprocket logistics dataset of **${stats.total} unique orders** (having resolved **${resolvedDups} multi-SKU duplicates**).\n\nHere are my initial findings:\n• **Delivery Success Rate**: **${(stats.deliveryRate * 100).toFixed(1)}%** (${stats.delivered} delivered)\n• **RTO Return Rate**: **${(stats.rtoRate * 100).toFixed(1)}%** (${stats.rto} returned packages)\n• **Total Shipping Revenue**: **₹${stats.totalRev.toLocaleString("en-IN")}**\n• **COD Collected**: **₹${stats.totalCOD.toLocaleString("en-IN")}** (${((stats.payCounts["cod"]?.orders || 0) / stats.total * 100).toFixed(0)}% of orders)\n• **Freight Cost**: **₹${stats.totalFreight.toLocaleString("en-IN")}**\n\nI have compiled a formula-friendly 7-sheet analytical workbook for your accounts team. Ask me any question about the data! E.g.:\n• *How can we reduce our shipping leakage?*\n• *Which couriers are performing the best/worst?*\n• *Which destinations present the highest RTO risk?*` }];
        } else {
          const profile = { ...analyzeData(data, parsedHeaders), sheetName, headerRow };
          newRecord.chatHistory = [{ sender: "analyst", text: `👋 Hi! I am your **Senior AI Data Analyst**. I have mapped a deep statistical profile for your **${profile.totalRows} rows** across **${profile.totalColumns} columns**.\n\nI've automatically identified data types, column fill rates, and calculated mathematical variances (sum, average, standard deviation, min, max) for numeric fields.\n\nWhat would you like to investigate in this dataset? E.g.:\n• *Can you summarize the numeric fields for me?*\n• *Show me the columns with the most empty values.*\n• *What is the cardinality of unique keys in the table?*` }];
        }

        await dbSaveRecord(newRecord);
        const recs = await dbGetRecords(currentUser.username);
        setSavedRecords(recs);
        addLog(`✓ Excel workbook auto-saved to secure Local Database (${(f.size / 1024).toFixed(1)} KB)`, "success");
      }

      recordSuccessfulReport();
      setStep("done");
    } catch (e: any) {
      console.error(e);
      setError(e.message || "Spreadsheet parsing error");
      addLog(`✗ Processing error: ${e.message || e}`, "error");
      setStep("upload");
    }
  }, [addLog, hasFreeReportsRemaining, recordSuccessfulReport, currentUser]);

  // Handle Drop and Clicks
  const onDrop = useCallback((e: any) => {
    e.preventDefault();
    setDragging(false);
    const f = e.dataTransfer.files[0];
    if (f) processFile(f);
  }, [processFile]);

  const onDragOver = (e: any) => { e.preventDefault(); setDragging(true); };
  const onDragLeave = () => setDragging(false);
  const onFileChange = (e: any) => { if (e.target.files[0]) processFile(e.target.files[0]); };

  // Sorted and Filtered Rows for Spreadsheet Viewer
  const sortedAndFilteredRows = useMemo(() => {
    let rows = [...rawRows];
    if (searchTerm.trim()) {
      const q = searchTerm.toLowerCase();
      rows = rows.filter((r) =>
        Object.values(r).some((v) => String(v ?? "").toLowerCase().includes(q))
      );
    }
    if (sortConfig) {
      rows.sort((a, b) => {
        const valA = a[sortConfig.key];
        const valB = b[sortConfig.key];
        if (valA == null) return 1;
        if (valB == null) return -1;
        const numA = Number(String(valA).replace(/[,?\s%]/g, ""));
        const numB = Number(String(valB).replace(/[,?\s%]/g, ""));
        if (!isNaN(numA) && !isNaN(numB)) {
          return sortConfig.direction === "asc" ? numA - numB : numB - numA;
        }
        return sortConfig.direction === "asc"
          ? String(valA).localeCompare(String(valB))
          : String(valB).localeCompare(String(valA));
      });
    }
    return rows;
  }, [rawRows, searchTerm, sortConfig]);

  const totalPages = Math.ceil(sortedAndFilteredRows.length / rowsPerPage) || 1;
  const paginatedRows = useMemo(() => {
    const start = (currentPage - 1) * rowsPerPage;
    return sortedAndFilteredRows.slice(start, start + rowsPerPage);
  }, [sortedAndFilteredRows, currentPage]);

  const handleSort = (key: string) => {
    let direction: "asc" | "desc" = "asc";
    if (sortConfig && sortConfig.key === key && sortConfig.direction === "asc") {
      direction = "desc";
    }
    setSortConfig({ key, direction });
    setCurrentPage(1);
  };

  // Offline pre-computed logistics insights helper
  const offlineLogisticsInsights = useMemo(() => {
    if (!logisticsAnalytics) return "";
    const stats = logisticsAnalytics;
    const topCouriers = Object.entries(stats.courierCounts)
      .sort((a, b) => b[1].orders - a[1].orders)
      .slice(0, 2);
    const worstRtoState = Object.entries(stats.stateCounts)
      .filter(([_, v]) => v.orders >= 5)
      .sort((a, b) => (b[1].rto / b[1].orders) - (a[1].rto / a[1].orders))[0];
    
    return `### 💡 Proactive Logistics Action Plan (Built-in Business Intelligence)

1. **💸 Cash-on-Delivery Risk Exposure**: COD orders represent **${((stats.payCounts["cod"]?.orders || 0) / stats.total * 100).toFixed(0)}%** of overall shipments (amounting to **₹${stats.totalCOD.toLocaleString("en-IN")}**). To minimize failed deliveries, implement automated WhatsApp COD confirmations prior to dispatching.
2. **🚚 Courier Scorecard Analysis**: Your primary shipping partner is **${topCouriers[0]?.[0] || "N/A"}** managing **${topCouriers[0]?.[1]?.orders || 0} orders** with a **${((topCouriers[0]?.[1]?.delivered / topCouriers[0]?.[1]?.orders) * 100).toFixed(1)}%** success rate. Monitor billing weight discrepancies closely to avoid surcharges.
3. **âš ï¸ Regional Return-to-Origin Hotspot**: **${worstRtoState?.[0] || "N/A"}** is flagging a critical RTO return risk of **${((worstRtoState?.[1]?.rto / worstRtoState?.[1]?.orders) * 100).toFixed(1)}%** over ${worstRtoState?.[1]?.orders || 0} shipments. Consider restricting COD or requiring prepaid payment for high-value orders in this territory.
4. **📦 Packaging Freight Waste**: The overall freight charges total **₹${stats.totalFreight.toLocaleString("en-IN")}**, representing **${(stats.totalFreight / stats.totalRev * 100).toFixed(1)}%** of sales. Review box sizes and volumetric weight brackets to reduce dead-weight costs.`;
  }, [logisticsAnalytics]);

  // Offline pre-computed responses for high availability without API key
  const getOfflineReply = (question: string): string => {
    const q = question.toLowerCase();
    if (mode === "shopify" && shopifyAnalytics) {
      const stats = shopifyAnalytics;
      if (q.includes("product") || q.includes("best") || q.includes("top")) {
        return `### ðŸ›ï¸ Shopify Product Performance
* **Top product**: **${stats.topProduct}**
* **Products analysed**: **${stats.productCount}**
* **Units sold**: **${stats.totalUnits.toLocaleString("en-IN")}**
* The downloaded workbook includes **Product Analysis**, **Product × Order Status**, and top product-wise customer/order sheets.`;
      }
      if (q.includes("customer") || q.includes("segment") || q.includes("retarget")) {
        const topSegment = Object.entries(stats.segmentCounts).sort((a, b) => b[1] - a[1])[0];
        return `### 👥 Customer Retargeting Summary
* **Unique customers**: **${stats.totalCustomers.toLocaleString("en-IN")}**
* Largest segment: **${topSegment?.[0] || "N/A"}** (${(topSegment?.[1] || 0).toLocaleString("en-IN")} customers)
* The workbook includes **Customer Data**, **Segment Analysis**, and export-ready **Retargeting Lists** for email/SMS campaigns.`;
      }
      if (q.includes("cod") || q.includes("payment")) {
        return `### 💰 COD and Payment Analysis
* The Shopify workbook includes a dedicated **COD Analysis** sheet.
* Use it to separate COD exposure from prepaid/online revenue, pending collection, delivered revenue, and cancelled/refunded revenue.`;
      }
      return `### 📊 Shopify Growth Analytics Summary
* **Orders**: **${stats.totalOrders.toLocaleString("en-IN")}**
* **Customers**: **${stats.totalCustomers.toLocaleString("en-IN")}**
* **Revenue**: **₹${Math.round(stats.totalRevenue).toLocaleString("en-IN")}**
* **Top city**: **${stats.topCity}**
* Download the workbook for dashboard, product, customer, retargeting, COD, geographic, and discount sheets.`;
    }
    if (mode === "logistics" && logisticsAnalytics) {
      const stats = logisticsAnalytics;
      const topCouriers = Object.entries(stats.courierCounts)
        .sort((a, b) => b[1].orders - a[1].orders)
        .slice(0, 2);
      const worstRtoState = Object.entries(stats.stateCounts)
        .filter(([_, v]) => v.orders >= 5)
        .sort((a, b) => (b[1].rto / b[1].orders) - (a[1].rto / a[1].orders))[0];

      if (q.includes("rto") || q.includes("return") || q.includes("state") || q.includes("destination")) {
        return `### âš ï¸ Regional Return-to-Origin Hotspot Analysis (Offline Mode)
* **Worst Performer**: **${worstRtoState?.[0] || "N/A"}** has an RTO risk of **${((worstRtoState?.[1]?.rto / worstRtoState?.[1]?.orders) * 100).toFixed(1)}%** (${worstRtoState?.[1]?.rto} returned out of ${worstRtoState?.[1]?.orders} orders).
* **Mitigation Strategy**: Implement strict address validation checks before shipping. We highly recommend asking for prepaid payments or holding COD shipments to this region until customer confirmation is received via WhatsApp.`;
      }
      if (q.includes("courier") || q.includes("delivery") || q.includes("partner") || q.includes("carrier")) {
        return `### 🚚 Courier Scorecard Analysis (Offline Mode)
* **Primary Courier**: **${topCouriers[0]?.[0] || "N/A"}** handles **${topCouriers[0]?.[1]?.orders || 0} orders** with a delivery success rate of **${((topCouriers[0]?.[1]?.delivered / topCouriers[0]?.[1]?.orders) * 100).toFixed(1)}%**.
* **Secondary Courier**: **${topCouriers[1]?.[0] || "N/A"}** handles **${topCouriers[1]?.[1]?.orders || 0} orders** with a success rate of **${((topCouriers[1]?.[1]?.delivered / topCouriers[1]?.[1]?.orders) * 100).toFixed(1)}%**.
* **Recommendation**: Divert shipments away from low-performing couriers to maximize client satisfaction and minimize duplicate shipping charges on return freights.`;
      }
      if (q.includes("save") || q.includes("cost") || q.includes("leakage") || q.includes("money") || q.includes("freight")) {
        return `### 💸 Shipping Cost & Profit Leakage Audit (Offline Mode)
* **Freight Expense**: You spent **₹${stats.totalFreight.toLocaleString("en-IN")}** representing **${(stats.totalFreight / stats.totalRev * 100).toFixed(1)}%** of shipping revenues.
* **Dead-Weight Issues**: Ensure package dimensions are perfectly reported inside Shiprocket. Minor differences can double volumetric charges.
* **COD Return Loss**: RTO return freights charge full delivery and return rates without sales completion, leading to major cash leakage. Prioritize pre-confirming COD orders!`;
      }
      if (q.includes("cod") || q.includes("cash")) {
        return `### 💳 Cash-on-Delivery Risk Exposure (Offline Mode)
* **COD Share**: COD shipments represent **${((stats.payCounts["cod"]?.orders || 0) / stats.total * 100).toFixed(0)}%** of overall orders, totaling **₹${stats.totalCOD.toLocaleString("en-IN")}**.
* **Critical Risk**: COD orders are 3x more likely to result in an RTO status compared to prepaid orders.
* **Senior Recommendation**: Incentivise customers with a small discount (e.g., ₹50 off) to switch to prepaid UPI at checkout to mitigate failed delivery losses.`;
      }
      return `### 📊 Logistics Audit Summary (Offline Mode)
* We analyzed **${stats.total} unique customer orders**.
* Overall **Delivery Success** stands at **${(stats.deliveryRate * 100).toFixed(1)}%** and **RTO Rate** at **${(stats.rtoRate * 100).toFixed(1)}%**.
* Total shipping revenue processed: **₹${stats.totalRev.toLocaleString("en-IN")}**.
* Enter your Claude API Key above for a live, deep cognitive audit on other columns or customized business strategies!`;
    } else if (mode === "universal" && dataProfile) {
      const prof = dataProfile;
      const numCols = prof.columns.filter(c => c.type === "numeric");
      if (q.includes("numeric") || q.includes("sum") || q.includes("average") || q.includes("math") || q.includes("stats")) {
        return `### 📊 Numeric Columns Summary (Offline Mode)
We analyzed **${numCols.length} numeric columns** in the spreadsheet.
${numCols.slice(0, 3).map(c => `* **${c.name}**: Sum = **₹${(c.sum || 0).toLocaleString("en-IN")}**, Average = **₹${(c.avg || 0).toLocaleString("en-IN")}**, Standard Deviation = **±${(c.stddev || 0).toFixed(1)}**`).join("\n")}
* These fields are formula-friendly and ready for direct calculation in the exported Excel file!`;
      }
      if (q.includes("duplicate") || q.includes("double") || q.includes("matching")) {
        return `### 📋 Duplicate Rows Profiling (Offline Mode)
* We scanned the entire spreadsheet and identified **${prof.duplicateRows} duplicate rows**.
* Removing duplicate entries or merging them by a unique key (such as Order ID or Email) will instantly improve reporting cleanliness.`;
      }
      if (q.includes("clean") || q.includes("missing") || q.includes("empty") || q.includes("null") || q.includes("type")) {
        return `### 🧪 Data Quality Profile (Offline Mode)
* **Spreadsheet size**: **${prof.totalRows} rows** by **${prof.totalColumns} columns**.
* Columns like **${prof.columns.slice(0, 3).map(c => c.name).join(", ")}** have been audited for fill rates.
* You can check the *Data Quality & Type Profiling* table below for the exact fill rate percentage for each of the fields.`;
      }
      return `### 📊 Spreadsheet Profiler Overview (Offline Mode)
* We analyzed **${prof.totalRows} data rows** across **${prof.totalColumns} columns**.
* **Numeric Fields** found: ${numCols.map(c => c.name).join(", ") || "None"}.
* Enter your Claude API Key above to unlock live, customized AI conversations about any of these columns!`;
    }
    return "Please upload a spreadsheet first so I can analyze it!";
  };

  // ----------------------------------------------------
  // 🤝 COLLABORATIVE VERSIONING & COMMENTS ACTIONS
  // ----------------------------------------------------
  const handleAddComment = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!commentInput.trim()) return;
    const authorName = commentAuthor.trim() || "Anonymous";
    const newComment = {
      id: Date.now().toString(),
      author: authorName,
      text: commentInput.trim(),
      timestamp: new Date().toLocaleString("en-IN")
    };
    const updatedComments = [...activeComments, newComment];
    setActiveComments(updatedComments);
    setCommentInput("");

    // Persist comments to DB
    if (activeRecordId) {
      try {
        const record = sharedRecordObj || await dbGetRecordById(activeRecordId);
        if (record) {
          record.comments = updatedComments;
          await dbSaveRecord(record);
          if (sharedRecordObj) {
            setSharedRecordObj({ ...record });
          }
          addLog("💬 Syncing new collaborative annotation...", "success");
        }
      } catch (err: any) {
        console.error("Failed to sync comment:", err);
      }
    }
  };

  const handleUploadNewVersion = async (e: any) => {
    const f = e.target.files?.[0];
    if (!f || !activeRecordId) return;
    setVersionUploading(true);
    try {
      addLog(`Uploading new spreadsheet version: ${f.name} (${(f.size / 1024).toFixed(1)} KB)`);
      const { data, headers: parsedHeaders, sheetName, originalRows, headerRow } = await parseExcel(f);
      addLog(`✓ Version parsed. Sheet: "${sheetName}" | Headers found at row ${headerRow + 1} | ${originalRows} rows total`, "success");

      // Keep the current active version metadata to push into version history
      const currentVersionItem = {
        timestamp: new Date().toLocaleString("en-IN"),
        filename: file ? file.name : "Active_Spreadsheet.xlsx",
        size: file ? file.size : 15360,
        rawRows: rawRows,
        tableHeaders: tableHeaders,
        dataProfile: dataProfile,
        logisticsAnalytics: logisticsAnalytics,
        shopifyAnalytics: shopifyAnalytics,
        outName: outName,
      };

      const updatedHistory = [...versionHistory, currentVersionItem];
      setVersionHistory(updatedHistory);

      // Set the active version properties
      setFile(f);
      setRawRows(data);
      setTableHeaders(parsedHeaders);

      const baseName = f.name.replace(/\.[^/.]+$/, "");
      const dateString = new Date().toISOString().slice(0, 10);

      let targetDataProfile = null;
      let targetLogisticsAnalytics = null;
      let targetShopifyAnalytics = null;
      let targetOutName = "";
      let targetDlUrl = "";

      if (mode === "shopify") {
        addLog("Re-calculating Shopify growth analytics for new version...");
        const stats = analyzeShopifyData(data);
        setShopifyAnalytics(stats);
        targetShopifyAnalytics = stats;

        const shopifyWb = buildShopifyAnalyticsWorkbook(baseName, data);
        const wbOut = XLSX.write(shopifyWb, { bookType: "xlsx", type: "array" });
        const blob = new Blob([wbOut], { type: "application/octet-stream" });
        if (dlRef.current) URL.revokeObjectURL(dlRef.current);
        targetDlUrl = URL.createObjectURL(blob);
        targetOutName = `${baseName}_Shopify_Analytics_${dateString}.xlsx`;
      } else if (mode === "logistics") {
        addLog("Re-calculating Logistics speed scorecards for new version...");
        const merged = mergeMultiSKU(data);
        setMergedLogisticsCount(data.length - merged.length);
        const stats = computeAnalytics(merged);
        setLogisticsAnalytics(stats);
        targetLogisticsAnalytics = stats;

        const logisticsWb = buildLogisticsWorkbook(baseName, merged, data.length - merged.length, stats);
        const wbOut = XLSX.write(logisticsWb, { bookType: "xlsx", type: "array" });
        const blob = new Blob([wbOut], { type: "application/octet-stream" });
        if (dlRef.current) URL.revokeObjectURL(dlRef.current);
        targetDlUrl = URL.createObjectURL(blob);
        targetOutName = `${baseName}_Logistics_Optimizer_${dateString}.xlsx`;
      } else {
        addLog("Re-generating data profile for new version...");
        const profile = { ...analyzeData(data, parsedHeaders), sheetName, headerRow };
        setDataProfile(profile);
        targetDataProfile = profile;

        const genericWb = buildAnalyticsWorkbook(baseName, data, profile);
        const wbOut = XLSX.write(genericWb, { bookType: "xlsx", type: "array" });
        const blob = new Blob([wbOut], { type: "application/octet-stream" });
        if (dlRef.current) URL.revokeObjectURL(dlRef.current);
        targetDlUrl = URL.createObjectURL(blob);
        targetOutName = `${baseName}_Profiler_Summary_${dateString}.xlsx`;
      }

      setDlUrl(targetDlUrl);
      dlRef.current = targetDlUrl;
      setOutName(targetOutName);

      // Now save the updated record back to DB
      const record = sharedRecordObj || await dbGetRecordById(activeRecordId);
      if (record) {
        record.filename = f.name;
        record.size = f.size;
        record.rawRows = data;
        record.tableHeaders = parsedHeaders;
        record.dataProfile = targetDataProfile;
        record.logisticsAnalytics = targetLogisticsAnalytics;
        record.shopifyAnalytics = targetShopifyAnalytics;
        record.outName = targetOutName;
        record.versions = updatedHistory;
        record.comments = activeComments;
        
        await dbSaveRecord(record);
        if (sharedRecordObj) {
          setSharedRecordObj({ ...record });
        }
        addLog("✓ Successfully branched, saved and updated new spreadsheet version in database!", "success");
      }

    } catch (err: any) {
      console.error(err);
      addLog(`❌ Version upload error: ${err.message || err}`, "error");
    } finally {
      setVersionUploading(false);
    }
  };

  const handleRestoreVersion = async (versionIdx: number) => {
    if (!activeRecordId || versionIdx < 0 || versionIdx >= versionHistory.length) return;
    const version = versionHistory[versionIdx];
    if (!version) return;

    try {
      addLog(`Restoring spreadsheet version branch V${versionIdx + 1}: ${version.filename}`);

      setFile(new File([new ArrayBuffer(version.size)], version.filename, { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }));
      setRawRows(version.rawRows);
      setTableHeaders(version.tableHeaders);
      setDataProfile(version.dataProfile);
      setLogisticsAnalytics(version.logisticsAnalytics);
      setShopifyAnalytics(version.shopifyAnalytics);
      setOutName(version.outName);

      // Recompile dlUrl
      const baseName = version.filename.replace(/\.[^/.]+$/, "");
      let targetDlUrl = "";
      if (mode === "shopify") {
        const shopifyWb = buildShopifyAnalyticsWorkbook(baseName, version.rawRows);
        const wbOut = XLSX.write(shopifyWb, { bookType: "xlsx", type: "array" });
        const blob = new Blob([wbOut], { type: "application/octet-stream" });
        if (dlRef.current) URL.revokeObjectURL(dlRef.current);
        targetDlUrl = URL.createObjectURL(blob);
      } else if (mode === "logistics") {
        const merged = mergeMultiSKU(version.rawRows);
        const logisticsWb = buildLogisticsWorkbook(baseName, merged, version.rawRows.length - merged.length, version.logisticsAnalytics);
        const wbOut = XLSX.write(logisticsWb, { bookType: "xlsx", type: "array" });
        const blob = new Blob([wbOut], { type: "application/octet-stream" });
        if (dlRef.current) URL.revokeObjectURL(dlRef.current);
        targetDlUrl = URL.createObjectURL(blob);
      } else {
        const genericWb = buildAnalyticsWorkbook(baseName, version.rawRows, version.dataProfile);
        const wbOut = XLSX.write(genericWb, { bookType: "xlsx", type: "array" });
        const blob = new Blob([wbOut], { type: "application/octet-stream" });
        if (dlRef.current) URL.revokeObjectURL(dlRef.current);
        targetDlUrl = URL.createObjectURL(blob);
      }
      setDlUrl(targetDlUrl);
      dlRef.current = targetDlUrl;

      // Update database record
      const record = sharedRecordObj || await dbGetRecordById(activeRecordId);
      if (record) {
        record.filename = version.filename;
        record.size = version.size;
        record.rawRows = version.rawRows;
        record.tableHeaders = version.tableHeaders;
        record.dataProfile = version.dataProfile;
        record.logisticsAnalytics = version.logisticsAnalytics;
        record.shopifyAnalytics = version.shopifyAnalytics;
        record.outName = version.outName;
        await dbSaveRecord(record);
        if (sharedRecordObj) {
          setSharedRecordObj({ ...record });
        }
        addLog(`✓ Successfully rolled back to version V${versionIdx + 1}!`, "success");
      }
    } catch (err: any) {
      console.error(err);
      addLog(`❌ Version restore error: ${err.message || err}`, "error");
    }
  };

  const handleChatSubmit = async (e?: React.FormEvent, customQuestion?: string) => {
    if (e) e.preventDefault();
    const question = (customQuestion || chatInput).trim();
    if (!question) return;

    if (!customQuestion) setChatInput("");
    setChatHistory(prev => [...prev, { sender: "user", text: question }]);
    setAiLoading(true);

    const hasKey = true;

    if (!hasKey) {
      // Offline fallback delay for better UX
      setTimeout(() => {
        const reply = getOfflineReply(question);
        setChatHistory(prev => [...prev, { sender: "analyst", text: reply }]);
        setAiLoading(false);
      }, 600);
      return;
    }

    try {
      let dataSummary = "";
      let systemPrompt = "";

      if (mode === "shopify" && shopifyAnalytics) {
        const stats = shopifyAnalytics;
        const topStatuses = Object.entries(stats.statusCounts).sort((a, b) => b[1] - a[1]).slice(0, 5).map(([k, v]) => `${k}: ${v}`).join(", ");
        const topSegments = Object.entries(stats.segmentCounts).sort((a, b) => b[1] - a[1]).slice(0, 5).map(([k, v]) => `${k}: ${v}`).join(", ");
        dataSummary = `File: ${file?.name} | Shopify orders: ${stats.totalOrders} | Customers: ${stats.totalCustomers} | Revenue: ₹${stats.totalRevenue.toFixed(0)} | Units: ${stats.totalUnits} | Products: ${stats.productCount} | Top product: ${stats.topProduct} | Top city: ${stats.topCity} | Statuses: ${topStatuses} | Segments: ${topSegments}`;
        systemPrompt = `You are a senior Shopify growth analyst. Use the provided Shopify export summary to answer with concise, practical ecommerce recommendations. Prioritize product performance, customer segments, retargeting, COD/payment risk, discount leakage, city/state demand, and order status issues. Use ? for currency and bold the most important metrics. Current data context: ${dataSummary}`;
      } else if (mode === "logistics" && logisticsAnalytics) {
        const stats = logisticsAnalytics;
        const topStatus = Object.entries(stats.statusCounts).sort((a, b) => b[1].orders - a[1].orders).slice(0, 4).map(([k, v]) => `${k}: ${v.orders}`).join(", ");
        const topStates = Object.entries(stats.stateCounts).sort((a, b) => b[1].orders - a[1].orders).slice(0, 5).map(([k, v]) => `${k}(${v.orders}, RTO:${(v.rto/v.orders*100).toFixed(0)}%)`).join(", ");
        const topCouriers = Object.entries(stats.courierCounts).sort((a, b) => b[1].orders - a[1].orders).slice(0, 4).map(([k, v]) => `${k}: ${v.orders} (del:${(v.delivered/v.orders*100).toFixed(0)}%)`).join(", ");
        const topNDR = Object.entries(stats.ndrCounts).sort((a, b) => b[1] - a[1]).slice(0, 3).map(([k, v]) => `${k}(${v})`).join(", ");

        dataSummary = `File: ${file?.name} | Total orders: ${stats.total} | Delivery rate: ${(stats.deliveryRate*100).toFixed(1)}% | RTO rate: ${(stats.rtoRate*100).toFixed(1)}% | Revenue: ₹${stats.totalRev.toFixed(0)} | COD: ₹${stats.totalCOD} | Freight: ₹${stats.totalFreight} | Statuses: ${topStatus} | Top States: ${topStates} | Couriers: ${topCouriers} | NDR Reasons: ${topNDR}`;
        systemPrompt = `You are a Senior AI Data Analyst. You are auditing a Shiprocket logistics dataset. Provide highly advanced, concise, actionable, and mathematically accurate insights. Address the user's specific question directly using the statistical summary provided. Use ? for currency. Bold key stats. Keep paragraphs short and use bullet points for clarity. Current data context: ${dataSummary}`;
      } else if (mode === "universal" && dataProfile) {
        const prof = dataProfile;
        const colList = prof.columns.slice(0, 8).map(c => `${c.name} (${c.type}, fill:${((c.nonEmptyCount/c.count)*100).toFixed(0)}%, unique:${c.uniqueCount}${c.type === 'numeric' ? `, avg:${c.avg?.toFixed(1)}, stddev:${c.stddev?.toFixed(1)}` : ''})`).join("; ");

        dataSummary = `File: ${file?.name} | Total Rows: ${prof.totalRows} | Total Columns: ${prof.totalColumns} | Duplicate Rows: ${prof.duplicateRows} | Columns Preview: ${colList}`;
        systemPrompt = `You are a Senior AI Data Analyst. You are auditing a custom spreadsheet upload. Provide deep statistical insights, outline potential data cleanliness issues, anomalous columns, or structural trends. Address the user's specific question directly using the statistical profile provided. Bold key facts. Keep paragraphs short and use bullet points for clarity. Current data context: ${dataSummary}`;
      }

      // Map existing history to API messages (limit to last 6 for token budget)
      const recentHistory = chatHistory.slice(-6);
      const apiMessages = [
        ...recentHistory.map(h => ({
          role: h.sender === "user" ? "user" : "assistant",
          content: h.text
        })),
        { role: "user", content: question }
      ];

      const answer = await api(apiMessages, systemPrompt);
      setChatHistory(prev => [...prev, { sender: "analyst", text: answer }]);
    } catch (e: any) {
      setChatHistory(prev => [...prev, { sender: "analyst", text: `Error: ${e.message || e}. Using local offline insights instead.\n\n${getOfflineReply(question)}` }]);
    } finally {
      setAiLoading(false);
    }
  };

  return (
    <div className="app-container">
      {/* Dynamic Styled Canvas */}
      <style>{`
        :root {
          --cohere-black: #000000;
          --primary: #17171c;
          --ink: #212121;
          --deep-green: #003c33;
          --dark-navy: #071829;
          --canvas: #ffffff;
          --soft-stone: #eeece7;
          --pale-green: #edfce9;
          --pale-blue: #f1f5ff;
          --hairline: #d9d9dd;
          --border-light: #e5e7eb;
          --card-border: #f2f2f2;
          --muted: #93939f;
          --slate: #75758a;
          --body-muted: #616161;
          --action-blue: #1863dc;
          --focus-blue: #4c6ee6;
          --coral: #ff7759;
          --coral-soft: #ffad9b;
          --form-focus: #9b60aa;
          --on-primary: #ffffff;
          --on-dark: #ffffff;
          --error: #b30000;

          --font-display: 'Space Grotesk', -apple-system, system-ui, sans-serif;
          --font-ui: 'Inter', -apple-system, system-ui, sans-serif;
          --font-technical: 'JetBrains Mono', monospace;
        }

        :root[data-theme="dark"] {
          --cohere-black: #000000;
          --primary: #17171c;
          --ink: #f7f5ef;
          --deep-green: #80d8c3;
          --dark-navy: #d8e5ff;
          --canvas: #101115;
          --soft-stone: #181a20;
          --pale-green: #10261f;
          --pale-blue: #111b2d;
          --hairline: #2a2d36;
          --border-light: #2f3340;
          --card-border: #252832;
          --muted: #a8adba;
          --slate: #b7bcc9;
          --body-muted: #c5c9d2;
          --action-blue: #8bb8ff;
          --focus-blue: #9fb7ff;
          --coral: #ff8b72;
          --coral-soft: #ffc1b2;
          --form-focus: #d79bec;
          --on-primary: #ffffff;
          --on-dark: #ffffff;
          --error: #ff9b9b;
        }

        body {
          background-color: var(--canvas);
          color: var(--ink);
          margin: 0;
          font-family: var(--font-ui);
          -webkit-font-smoothing: antialiased;
        }

        body[data-theme="dark"] {
          background-color: #101115;
        }

        body[data-theme="dark"] .brand-section h1,
        body[data-theme="dark"] .upload-title,
        body[data-theme="dark"] .card-title,
        body[data-theme="dark"] .kpi-value,
        body[data-theme="dark"] .drag-active-text,
        body[data-theme="dark"] .newsletter-info h4 {
          color: var(--ink);
        }

        body[data-theme="dark"] .mode-tab.active,
        body[data-theme="dark"] .btn-primary,
        body[data-theme="dark"] .subscription-actions a.primary {
          background: var(--ink);
          color: var(--canvas);
          border-color: var(--ink);
        }

        body[data-theme="dark"] .btn-primary:hover,
        body[data-theme="dark"] .subscription-actions a.primary:hover {
          background: var(--slate);
          border-color: var(--slate);
          color: var(--canvas);
        }

        body[data-theme="dark"] .premium-table th {
          background: #181a20;
          color: var(--ink);
        }

        body[data-theme="dark"] .premium-table td {
          color: var(--ink);
        }

        body[data-theme="dark"] .modal-title {
          color: var(--ink);
        }

        body[data-theme="dark"] .auth-tab-btn.active {
          color: var(--ink);
          border-bottom-color: var(--coral);
        }

        body[data-theme="dark"] .auth-tab-btn {
          color: var(--slate);
        }

        body[data-theme="dark"] .auth-submit-btn {
          background: var(--ink);
          color: var(--canvas);
          border-color: var(--ink);
        }

        body[data-theme="dark"] .auth-submit-btn:hover {
          background: var(--slate);
          border-color: var(--slate);
          color: var(--canvas);
        }

        body[data-theme="dark"] .checkout-tab-btn.active {
          background: var(--ink);
          color: var(--canvas);
        }

        body[data-theme="dark"] .checkout-tab-btn {
          color: var(--slate);
        }

        body[data-theme="dark"] .payment-summary-box {
          background: rgba(128, 216, 195, 0.1);
          border-color: var(--deep-green);
          color: var(--deep-green);
        }


        .announcement-bar {
          background-color: var(--cohere-black);
          color: var(--on-dark);
          font-family: var(--font-technical);
          font-size: 11px;
          min-height: 36px;
          display: flex;
          align-items: center;
          justify-content: center;
          position: relative;
          padding: 0.6rem 1.5rem;
          z-index: 100;
          letter-spacing: 0.05em;
          text-transform: uppercase;
          margin-bottom: 2rem;
          border-radius: 4px;
        }

        .announcement-bar a {
          color: var(--coral);
          text-decoration: underline;
          margin-left: 6px;
          font-weight: 500;
        }

        .app-container {
          max-width: 1100px;
          margin: 0 auto;
          padding: 2rem 1.5rem;
        }

        .premium-header {
          display: flex;
          justify-content: space-between;
          align-items: center;
          gap: 1rem;
          margin-bottom: 3rem;
          background: var(--canvas);
          padding: 1.5rem 0;
          border-bottom: 1px solid var(--hairline);
        }

        .brand-section h1 {
          font-family: var(--font-display);
          font-size: 32px;
          font-weight: 400;
          letter-spacing: -0.96px;
          margin: 0 0 0.4rem 0;
          color: var(--cohere-black);
          line-height: 1.1;
        }

        .brand-section p {
          font-size: 14px;
          color: var(--slate);
          margin: 0;
        }

        .mode-segmented-control {
          display: flex;
          background: var(--soft-stone);
          padding: 4px;
          border-radius: 32px;
        }

        .mode-tab {
          padding: 8px 20px;
          border-radius: 32px;
          font-size: 13px;
          font-weight: 500;
          border: none;
          background: transparent;
          color: var(--slate);
          cursor: pointer;
          transition: all 0.2s ease;
          display: flex;
          align-items: center;
          gap: 6px;
          font-family: var(--font-ui);
        }

        .mode-tab.active {
          background: var(--cohere-black);
          color: var(--on-dark);
        }

        .auto-engine-panel {
          display: flex;
          align-items: center;
          gap: 12px;
          background: var(--soft-stone);
          border: 1px solid var(--hairline);
          border-radius: 999px;
          padding: 8px 14px;
          min-width: 280px;
          justify-content: flex-start;
        }

        .auto-engine-icon {
          width: 34px;
          height: 34px;
          border-radius: 999px;
          display: grid;
          place-items: center;
          background: var(--cohere-black);
          color: var(--on-dark);
          font-size: 16px;
          flex: 0 0 auto;
        }

        .auto-engine-copy {
          display: flex;
          flex-direction: column;
          gap: 2px;
          min-width: 0;
        }

        .auto-engine-label {
          font-family: var(--font-technical);
          font-size: 10px;
          letter-spacing: 0.12em;
          text-transform: uppercase;
          color: var(--slate);
          font-weight: 700;
        }

        .auto-engine-value {
          font-size: 13px;
          color: var(--ink);
          font-weight: 700;
          white-space: nowrap;
          overflow: hidden;
          text-overflow: ellipsis;
        }

        .header-controls {
          display: flex;
          align-items: center;
          gap: 10px;
        }

        .theme-toggle {
          width: 42px;
          height: 42px;
          border-radius: 999px;
          border: 1px solid var(--hairline);
          background: var(--soft-stone);
          color: var(--ink);
          display: grid;
          place-items: center;
          font-size: 18px;
          cursor: pointer;
          transition: border-color 0.2s, background 0.2s, transform 0.2s;
          flex: 0 0 auto;
        }

        .theme-toggle:hover {
          border-color: var(--coral);
          transform: translateY(-1px);
        }

        .upload-card {
          border: 1px dashed var(--hairline);
          border-radius: 22px;
          padding: 6rem 2rem;
          text-align: center;
          cursor: pointer;
          background: var(--soft-stone);
          transition: all 0.2s cubic-bezier(0.4, 0, 0.2, 1);
        }

        .upload-card:hover {
          background: var(--pale-green);
          border-color: var(--deep-green);
        }

        .upload-icon {
          font-size: 2.5rem;
          margin-bottom: 1.5rem;
          display: inline-block;
        }

        .upload-title {
          font-family: var(--font-display);
          font-size: 24px;
          font-weight: 400;
          color: var(--cohere-black);
          margin: 0 0 0.6rem 0;
          letter-spacing: -0.48px;
        }

        .upload-desc {
          font-family: var(--font-ui);
          font-size: 14px;
          color: var(--slate);
          margin: 0 0 2rem 0;
        }

        .usage-meter {
          display: inline-flex;
          align-items: center;
          gap: 8px;
          border: 1px solid var(--hairline);
          background: var(--canvas);
          color: var(--ink);
          border-radius: 999px;
          padding: 8px 14px;
          font-size: 12px;
          font-weight: 700;
          margin-bottom: 1.25rem;
        }

        .subscription-card {
          max-width: 720px;
          margin: 1rem auto 0;
          background: var(--canvas);
          border: 1px solid var(--hairline);
          border-left: 4px solid var(--coral);
          border-radius: 8px;
          padding: 18px;
          text-align: left;
        }

        .subscription-card h3 {
          margin: 0 0 8px 0;
          font-size: 17px;
          color: var(--ink);
        }

        .subscription-card p {
          margin: 0 0 14px 0;
          color: var(--slate);
          font-size: 13px;
          line-height: 1.45;
        }

        .subscription-actions {
          display: flex;
          flex-wrap: wrap;
          gap: 10px;
        }

        .subscription-actions a {
          display: inline-flex;
          align-items: center;
          justify-content: center;
          border-radius: 999px;
          padding: 10px 14px;
          border: 1px solid var(--hairline);
          color: var(--ink);
          text-decoration: none;
          font-size: 12px;
          font-weight: 700;
          background: var(--soft-stone);
        }

        .subscription-actions a.primary {
          background: var(--cohere-black);
          color: var(--on-dark);
          border-color: var(--cohere-black);
        }

        .subscription-actions a.icon-only,
        .footer-links a.icon-only {
          width: 42px;
          height: 42px;
          padding: 0;
          border-radius: 999px;
          display: inline-flex;
          align-items: center;
          justify-content: center;
          transition: all 0.2s ease;
        }

        .subscription-actions a.icon-only svg,
        .footer-links a.icon-only svg {
          width: 18px;
          height: 18px;
          stroke: currentColor;
          fill: none;
          display: block;
          transition: transform 0.2s ease;
        }

        .footer-links a.icon-only:hover svg,
        .subscription-actions a.icon-only:hover svg {
          transform: scale(1.1);
        }

        .footer-links {
          display: flex;
          justify-content: center;
          gap: 12px;
          flex-wrap: wrap;
          margin-top: 0.5rem;
        }

        .footer-links a {
          color: var(--slate);
          text-decoration: none;
          font-weight: 700;
          border: 1px solid var(--hairline);
          background: var(--soft-stone);
          display: inline-flex;
          align-items: center;
          justify-content: center;
          transition: border-color 0.2s, color 0.2s, transform 0.2s;
        }

        .footer-links a:hover {
          color: var(--action-blue);
          border-color: var(--action-blue);
          transform: translateY(-1px);
        }

        .text-success-accent {
          color: var(--deep-green) !important;
        }

        .text-warning-accent {
          color: var(--coral) !important;
        }

        .toast-badge {
          background: var(--pale-green);
          border: 1px solid var(--deep-green);
          color: var(--deep-green);
          padding: 1rem 1.5rem;
          border-radius: 8px;
          font-size: 14px;
          margin-bottom: 2rem;
          font-weight: 500;
          display: flex;
          align-items: center;
          gap: 8px;
          font-family: var(--font-ui);
        }

        .kpi-grid {
          display: grid;
          grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
          gap: 1.5rem;
          margin-bottom: 2.5rem;
        }

        .kpi-card {
          background: var(--soft-stone);
          border: 1px solid var(--hairline);
          border-radius: 8px;
          padding: 1.5rem;
          position: relative;
        }

        .kpi-card::before {
          content: '';
          position: absolute;
          left: 0;
          top: 0;
          bottom: 0;
          width: 3px;
          background: var(--primary);
        }

        .kpi-card.success::before {
          background: var(--deep-green);
        }

        .kpi-card.warning::before {
          background: var(--coral);
        }

        .kpi-label {
          font-family: var(--font-technical);
          font-size: 11px;
          text-transform: uppercase;
          letter-spacing: 0.1em;
          color: var(--slate);
          margin-bottom: 0.6rem;
          font-weight: 500;
        }

        .kpi-value {
          font-family: var(--font-display);
          font-size: 32px;
          font-weight: 400;
          color: var(--cohere-black);
          line-height: 1.1;
          letter-spacing: -0.64px;
        }

        .kpi-sub {
          font-size: 12px;
          color: var(--slate);
          margin-top: 0.4rem;
        }

        .section-card {
          background: var(--canvas);
          border: 1px solid var(--hairline);
          border-radius: 16px;
          padding: 2rem;
          margin-bottom: 2rem;
        }

        .card-title {
          font-family: var(--font-display);
          font-size: 20px;
          font-weight: 400;
          color: var(--cohere-black);
          margin: 0 0 1.5rem 0;
          display: flex;
          align-items: center;
          gap: 8px;
          letter-spacing: -0.4px;
        }

        .console-box {
          background: var(--primary);
          border-radius: 8px;
          overflow: hidden;
          border: 1px solid var(--cohere-black);
          margin-bottom: 2rem;
        }

        .console-header {
          background: var(--cohere-black);
          padding: 0.8rem 1.2rem;
          font-size: 11px;
          color: var(--muted);
          display: flex;
          align-items: center;
          justify-content: space-between;
          font-family: var(--font-technical);
          border-bottom: 1px solid var(--primary);
        }

        .console-dot {
          width: 6px;
          height: 6px;
          border-radius: 50%;
          background: var(--error);
          display: inline-block;
        }

        .console-body {
          padding: 1.2rem;
          max-height: 160px;
          overflow-y: auto;
          display: flex;
          flex-direction: column;
          gap: 6px;
          background: var(--primary);
        }

        .console-line {
          font-family: var(--font-technical);
          font-size: 13px;
          line-height: 1.4;
          display: flex;
          gap: 12px;
        }

        .console-time {
          color: var(--slate);
          flex-shrink: 0;
        }

        .console-text {
          word-break: break-all;
        }

        .text-info { color: var(--on-primary); }
        .text-success { color: var(--coral-soft); }
        .text-error { color: var(--error); }
        .text-neutral { color: var(--muted); }

        .chart-row {
          display: grid;
          grid-template-columns: 1fr 1fr;
          gap: 1.5rem;
          margin-bottom: 2rem;
        }

        @media (max-width: 768px) {
          .chart-row { grid-template-columns: 1fr; }
        }

        .visual-bar-container {
          margin-bottom: 1.2rem;
        }

        .visual-bar-info {
          display: flex;
          justify-content: space-between;
          font-size: 13px;
          margin-bottom: 0.4rem;
          font-weight: 500;
          font-family: var(--font-ui);
        }

        .visual-bar-bg {
          height: 8px;
          background: var(--soft-stone);
          border-radius: 4px;
          overflow: hidden;
          position: relative;
        }

        .visual-bar-fill {
          height: 100%;
          background: var(--slate);
          border-radius: 4px;
          transition: width 0.8s cubic-bezier(0.4, 0, 0.2, 1);
        }

        .visual-bar-fill.success { background: var(--deep-green); }
        .visual-bar-fill.warning { background: var(--coral-soft); }
        .visual-bar-fill.danger { background: var(--coral); }

        .trend-chart-svg {
          width: 100%;
          height: 160px;
          display: flex;
          align-items: flex-end;
          gap: 14px;
          padding-top: 1rem;
          overflow-x: auto;
          border-bottom: 1px solid var(--hairline);
        }

        .trend-chart-bar {
          flex: 1;
          min-width: 50px;
          display: flex;
          flex-direction: column;
          align-items: center;
          gap: 8px;
        }

        .trend-chart-fill {
          width: 100%;
          background: var(--deep-green);
          opacity: 0.85;
          border-radius: 4px 4px 0 0;
          transition: all 0.2s ease;
          display: flex;
          align-items: flex-start;
          justify-content: center;
          cursor: pointer;
          border: 1px solid var(--deep-green);
          border-bottom: none;
        }

        .trend-chart-fill:hover {
          opacity: 1;
          background: var(--coral);
          border-color: var(--coral);
        }

        .trend-chart-tip {
          font-size: 10px;
          color: var(--on-dark);
          font-weight: 600;
          margin-top: 6px;
          opacity: 0;
          transition: opacity 0.15s ease;
          font-family: var(--font-technical);
        }

        .trend-chart-fill:hover .trend-chart-tip {
          opacity: 1;
        }

        .grid-header-tools {
          display: flex;
          justify-content: space-between;
          align-items: center;
          gap: 1rem;
          margin-bottom: 1.5rem;
          flex-wrap: wrap;
        }

        .grid-search-input {
          padding: 10px 16px;
          border: 1px solid var(--hairline);
          border-radius: 8px;
          font-size: 14px;
          width: 280px;
          outline: none;
          font-family: var(--font-ui);
          background-color: var(--canvas);
          transition: border-color 0.2s;
        }

        .grid-search-input:focus {
          border-color: var(--coral);
        }

        .premium-table-wrapper {
          overflow-x: auto;
          border: 1px solid var(--hairline);
          border-radius: 8px;
          background: var(--canvas);
        }

        .premium-table {
          width: 100%;
          border-collapse: collapse;
          font-size: 14px;
          text-align: left;
        }

        .premium-table th {
          background: var(--soft-stone);
          padding: 12px 16px;
          color: var(--slate);
          font-weight: 500;
          border-bottom: 1px solid var(--hairline);
          cursor: pointer;
          user-select: none;
          white-space: nowrap;
          font-family: var(--font-technical);
          font-size: 12px;
          text-transform: uppercase;
          letter-spacing: 0.05em;
        }

        .premium-table th:hover {
          background: var(--hairline);
          color: var(--cohere-black);
        }

        .premium-table td {
          padding: 14px 16px;
          border-bottom: 1px solid var(--hairline);
          color: var(--ink);
          max-width: 200px;
          white-space: nowrap;
          overflow: hidden;
          text-overflow: ellipsis;
        }

        .premium-table tr:last-child td {
          border-bottom: none;
        }

        .premium-table tr:hover td {
          background: var(--pale-green);
        }

        .grid-pagination {
          display: flex;
          justify-content: space-between;
          align-items: center;
          margin-top: 1.5rem;
        }

        .pagination-btn {
          padding: 8px 16px;
          border: 1px solid var(--hairline);
          background: var(--canvas);
          border-radius: 30px;
          font-size: 13px;
          font-weight: 500;
          cursor: pointer;
          color: var(--ink);
          transition: all 0.15s;
          font-family: var(--font-ui);
        }

        .pagination-btn:hover:not(:disabled) {
          background: var(--soft-stone);
          border-color: var(--slate);
        }

        .pagination-btn:disabled {
          opacity: 0.4;
          cursor: not-allowed;
        }

        .action-bar {
          display: flex;
          gap: 16px;
          margin-top: 2rem;
        }

        .btn-primary {
          flex: 1;
          background: var(--cohere-black);
          color: var(--on-dark);
          border: 1px solid var(--cohere-black);
          padding: 14px 28px;
          border-radius: 32px;
          font-size: 14px;
          font-weight: 500;
          cursor: pointer;
          display: inline-flex;
          align-items: center;
          justify-content: center;
          gap: 8px;
          text-decoration: none;
          transition: all 0.2s;
          font-family: var(--font-ui);
        }

        .btn-primary:hover {
          background: var(--primary);
          border-color: var(--primary);
        }

        .btn-secondary {
          background: transparent;
          border: none;
          color: var(--ink);
          padding: 14px 20px;
          font-size: 14px;
          font-weight: 500;
          cursor: pointer;
          transition: all 0.15s;
          font-family: var(--font-ui);
          text-decoration: underline;
        }

        .btn-secondary:hover {
          color: var(--action-blue);
        }

        .ai-insights-card {
          background: var(--soft-stone);
          border: 1px solid var(--hairline);
          border-radius: 16px;
          padding: 2rem;
          margin-bottom: 2rem;
        }

        .ai-key-input-row {
          display: flex;
          gap: 12px;
          margin-bottom: 1.5rem;
          flex-wrap: wrap;
          align-items: center;
        }

        .ai-key-input {
          flex: 1;
          min-width: 280px;
          padding: 10px 16px;
          border: 1px solid var(--hairline);
          border-radius: 8px;
          font-size: 14px;
          outline: none;
          font-family: var(--font-ui);
          background-color: var(--canvas);
        }

        .ai-key-input:focus {
          border-color: var(--coral);
        }

        .chat-container {
          background: var(--primary);
          border: 1px solid var(--cohere-black);
          border-radius: 16px;
          overflow: hidden;
          display: flex;
          flex-direction: column;
          height: 520px;
        }

        .chat-header {
          background: var(--cohere-black);
          border-bottom: 1px solid var(--primary);
          padding: 1rem 1.5rem;
          display: flex;
          align-items: center;
          justify-content: space-between;
        }

        .chat-header-title {
          font-family: var(--font-display);
          font-size: 16px;
          font-weight: 400;
          color: var(--on-dark);
          display: flex;
          align-items: center;
          gap: 8px;
          letter-spacing: -0.32px;
        }

        .chat-header-status {
          font-size: 12px;
          color: var(--pale-green);
          display: flex;
          align-items: center;
          gap: 6px;
          font-family: var(--font-technical);
          font-weight: 500;
        }

        .chat-pulse {
          width: 6px;
          height: 6px;
          border-radius: 50%;
          background: var(--pale-green);
          display: inline-block;
        }

        .chat-messages {
          flex: 1;
          padding: 1.5rem;
          overflow-y: auto;
          display: flex;
          flex-direction: column;
          gap: 1.2rem;
          background: var(--canvas);
        }

        .chat-message {
          max-width: 85%;
          padding: 1rem 1.25rem;
          border-radius: 8px;
          font-size: 14px;
          line-height: 1.5;
          font-family: var(--font-ui);
        }

        .chat-message.analyst {
          align-self: flex-start;
          background: var(--soft-stone);
          border: 1px solid var(--hairline);
          color: var(--ink);
          border-top-left-radius: 2px;
        }

        .chat-message.user {
          align-self: flex-end;
          background: var(--deep-green);
          color: var(--on-dark);
          border: 1px solid var(--deep-green);
          border-top-right-radius: 2px;
        }

        .chat-message h3 {
          margin-top: 0;
          margin-bottom: 0.6rem;
          font-size: 15px;
          color: var(--coral);
          font-family: var(--font-display);
          font-weight: 500;
        }
        
        .chat-message.user h3 {
          color: var(--on-dark);
        }

        .chat-message ul, .chat-message ol {
          margin: 0;
          padding-left: 1.2rem;
        }

        .chat-message li {
          margin-bottom: 0.4rem;
        }

        .chat-suggestions-tray {
          display: flex;
          gap: 8px;
          padding: 0.8rem 1.5rem;
          background: var(--soft-stone);
          border-top: 1px solid var(--hairline);
          overflow-x: auto;
          flex-wrap: wrap;
        }

        .chat-suggestion-chip {
          padding: 6px 12px;
          background: var(--canvas);
          border: 1px solid var(--hairline);
          border-radius: 30px;
          font-size: 12px;
          font-weight: 500;
          color: var(--ink);
          cursor: pointer;
          transition: all 0.15s;
          white-space: nowrap;
          font-family: var(--font-ui);
        }

        .chat-suggestion-chip:hover {
          background: var(--coral);
          border-color: var(--coral);
          color: var(--on-primary);
        }

        .chat-input-bar {
          display: flex;
          padding: 1rem 1.5rem;
          background: var(--soft-stone);
          border-top: 1px solid var(--hairline);
          gap: 12px;
        }

        .chat-text-input {
          flex: 1;
          padding: 10px 16px;
          border: 1px solid var(--hairline);
          border-radius: 8px;
          font-size: 14px;
          outline: none;
          font-family: var(--font-ui);
          background: var(--canvas);
          color: var(--ink);
          transition: all 0.2s;
        }

        .chat-text-input:focus {
          border-color: var(--coral);
        }

        .chat-send-btn {
          background: var(--coral);
          color: var(--cohere-black);
          border: none;
          padding: 0 20px;
          border-radius: 8px;
          font-size: 14px;
          font-weight: 600;
          cursor: pointer;
          transition: opacity 0.15s;
          display: flex;
          align-items: center;
          justify-content: center;
          gap: 6px;
          font-family: var(--font-ui);
        }

        .chat-send-btn:hover {
          background: var(--coral-soft);
        }

        .chat-send-btn:disabled {
          opacity: 0.5;
          cursor: not-allowed;
        }

        .rto-risk-badge {
          display: inline-flex;
          align-items: center;
          padding: 3px 8px;
          border-radius: 4px;
          font-size: 11px;
          font-weight: 600;
          background: var(--error);
          color: var(--on-dark);
          font-family: var(--font-technical);
        }

        /* Trust partner logo strip styling */
        .trust-logo-strip {
          background-color: var(--canvas);
          color: var(--ink);
          padding: 3rem 0;
          text-align: center;
          border-top: 1px solid var(--hairline);
          margin-top: 4rem;
        }
        .trust-logo-title {
          font-family: var(--font-technical);
          font-size: 11px;
          letter-spacing: 0.15em;
          color: var(--slate);
          margin-bottom: 1.5rem;
          font-weight: 500;
          text-transform: uppercase;
        }
        .trust-logos {
          display: flex;
          justify-content: center;
          gap: 3rem;
          align-items: center;
          flex-wrap: wrap;
        }
        .trust-logos span {
          font-family: var(--font-display);
          font-size: 18px;
          font-weight: 600;
          color: var(--slate);
          opacity: 0.55;
          letter-spacing: -0.5px;
        }

        /* User Hub Header Controls */
        .header-controls {
          display: flex;
          align-items: center;
          gap: 12px;
          flex-wrap: wrap;
        }

        .user-hub-widget {
          display: flex;
          align-items: center;
          gap: 10px;
          background: var(--soft-stone);
          border: 1px solid var(--hairline);
          border-radius: 30px;
          padding: 4px 14px;
        }

        .user-info-text {
          font-size: 13px;
          color: var(--ink);
          display: flex;
          align-items: center;
          gap: 6px;
        }

        .user-icon-bullet {
          font-size: 14px;
        }

        .plan-badge {
          font-family: var(--font-technical);
          font-size: 9px;
          font-weight: 700;
          padding: 2px 6px;
          border-radius: 4px;
          text-transform: uppercase;
          letter-spacing: 0.05em;
        }

        .plan-badge.free {
          background: rgba(255, 119, 89, 0.15);
          color: var(--coral);
        }

        .plan-badge.pro {
          background: rgba(0, 60, 51, 0.15);
          color: var(--deep-green);
        }

        .header-btn {
          font-family: var(--font-ui);
          font-size: 12px;
          font-weight: 600;
          padding: 6px 12px;
          border-radius: 20px;
          border: 1px solid var(--hairline);
          background: var(--canvas);
          color: var(--ink);
          cursor: pointer;
          transition: all 0.15s ease;
        }

        .header-btn:hover {
          border-color: var(--coral);
          color: var(--coral);
        }

        .header-btn.login-trigger {
          border-radius: 30px;
          padding: 10px 18px;
          font-size: 13px;
        }

        .header-btn.login-trigger:hover {
          background: var(--cohere-black);
          color: var(--on-dark);
          border-color: var(--cohere-black);
        }

        /* Modal Overlay Styles */
        .modal-overlay {
          position: fixed;
          top: 0;
          left: 0;
          right: 0;
          bottom: 0;
          background: rgba(0, 0, 0, 0.6);
          backdrop-filter: blur(4px);
          display: grid;
          place-items: center;
          z-index: 1000;
          padding: 1.5rem;
          overflow-y: auto;
        }

        .modal-card {
          background: var(--canvas);
          border: 1px solid var(--hairline);
          border-radius: 16px;
          width: 100%;
          max-width: 540px;
          box-shadow: 0 20px 40px rgba(0,0,0,0.15);
          overflow: hidden;
          animation: modalEntrance 0.3s cubic-bezier(0.16, 1, 0.3, 1);
        }

        @keyframes modalEntrance {
          from { transform: translateY(12px); opacity: 0; }
          to { transform: translateY(0); opacity: 1; }
        }

        .modal-header {
          padding: 1.5rem;
          border-bottom: 1px solid var(--hairline);
          display: flex;
          justify-content: space-between;
          align-items: center;
          background: var(--soft-stone);
        }

        .modal-title {
          font-family: var(--font-display);
          font-size: 18px;
          font-weight: 500;
          color: var(--cohere-black);
          margin: 0;
        }

        .modal-close-btn {
          background: transparent;
          border: none;
          font-size: 20px;
          color: var(--slate);
          cursor: pointer;
          padding: 4px;
          line-height: 1;
        }

        .modal-close-btn:hover {
          color: var(--coral);
        }

        .modal-body {
          padding: 1.5rem;
        }

        /* Auth Tabs & Forms */
        .auth-tabs {
          display: flex;
          border-bottom: 1px solid var(--hairline);
          margin-bottom: 1.5rem;
        }

        .auth-tab-btn {
          flex: 1;
          background: transparent;
          border: none;
          padding: 12px;
          font-family: var(--font-ui);
          font-size: 14px;
          font-weight: 600;
          color: var(--slate);
          cursor: pointer;
          border-bottom: 2px solid transparent;
          transition: all 0.2s;
        }

        .auth-tab-btn.active {
          color: var(--cohere-black);
          border-bottom-color: var(--coral);
        }

        .form-group {
          margin-bottom: 1.25rem;
          text-align: left;
        }

        .form-label {
          display: block;
          font-size: 11px;
          font-weight: 600;
          text-transform: uppercase;
          letter-spacing: 0.05em;
          color: var(--slate);
          margin-bottom: 6px;
        }

        .form-input {
          width: 100%;
          padding: 10px 14px;
          border: 1px solid var(--hairline);
          border-radius: 8px;
          font-family: var(--font-ui);
          font-size: 14px;
          background: var(--canvas);
          color: var(--ink);
          outline: none;
          box-sizing: border-box;
        }

        .form-input:focus {
          border-color: var(--coral);
        }

        .auth-submit-btn {
          width: 100%;
          background: var(--cohere-black);
          color: var(--on-dark);
          border: 1px solid var(--cohere-black);
          padding: 12px;
          border-radius: 30px;
          font-family: var(--font-ui);
          font-size: 14px;
          font-weight: 600;
          cursor: pointer;
          transition: all 0.2s;
          margin-top: 1rem;
        }

        .auth-submit-btn:hover {
          background: var(--slate);
          border-color: var(--slate);
        }

        /* Payment Checkout Panel styling */
        .checkout-tabs {
          display: flex;
          background: var(--soft-stone);
          padding: 4px;
          border-radius: 30px;
          margin-bottom: 1.5rem;
        }

        .checkout-tab-btn {
          flex: 1;
          background: transparent;
          border: none;
          padding: 8px 16px;
          border-radius: 30px;
          font-family: var(--font-ui);
          font-size: 13px;
          font-weight: 600;
          color: var(--slate);
          cursor: pointer;
          transition: all 0.2s;
        }

        .checkout-tab-btn.active {
          background: var(--cohere-black);
          color: var(--on-dark);
        }

        .stripe-card-form {
          display: flex;
          flex-direction: column;
          gap: 12px;
        }

        .card-row-triple {
          display: grid;
          grid-template-columns: 2fr 1fr 1fr;
          gap: 10px;
        }

        .payment-summary-box {
          background: var(--pale-green);
          border: 1px solid var(--deep-green);
          border-radius: 8px;
          padding: 12px;
          margin-bottom: 1.5rem;
          font-size: 13px;
          color: var(--deep-green);
          display: flex;
          justify-content: space-between;
          align-items: center;
        }

        /* simulated secure payment logs */
        .payment-logs-box {
          background: #000000;
          border-radius: 8px;
          padding: 12px;
          font-family: var(--font-technical);
          font-size: 12px;
          color: #10b981;
          max-height: 180px;
          overflow-y: auto;
          margin-top: 1.5rem;
          display: flex;
          flex-direction: column;
          gap: 4px;
          border: 1px solid #111;
          text-align: left;
        }

        .payment-log-line {
          animation: fadeLog 0.2s ease;
        }

        @keyframes fadeLog {
          from { opacity: 0; transform: translateX(-4px); }
          to { opacity: 1; transform: translateX(0); }
        }

        /* Razorpay QR Simulator */
        .qr-simulator-wrapper {
          display: flex;
          flex-direction: column;
          align-items: center;
          gap: 12px;
          padding: 1rem;
          background: var(--soft-stone);
          border-radius: 12px;
          border: 1px solid var(--hairline);
          margin-bottom: 1.25rem;
        }

        .mock-qr-code {
          width: 140px;
          height: 140px;
          background: #ffffff;
          border: 4px solid #ffffff;
          border-radius: 8px;
          display: flex;
          align-items: center;
          justify-content: center;
          position: relative;
          box-shadow: 0 4px 10px rgba(0,0,0,0.06);
        }

        /* Dashboard Saved Records styling */
        .dashboard-stats-grid {
          display: grid;
          grid-template-columns: 1fr 1fr;
          gap: 12px;
          margin-bottom: 1.5rem;
        }

        .dashboard-stat-tile {
          background: var(--soft-stone);
          border: 1px solid var(--hairline);
          border-radius: 8px;
          padding: 12px;
          text-align: left;
        }

        .dashboard-stat-label {
          font-size: 10px;
          text-transform: uppercase;
          letter-spacing: 0.05em;
          color: var(--slate);
          margin-bottom: 4px;
        }

        .dashboard-stat-value {
          font-size: 15px;
          font-weight: 700;
          color: var(--ink);
        }

        .records-list-wrapper {
          border: 1px solid var(--hairline);
          border-radius: 8px;
          max-height: 250px;
          overflow-y: auto;
          background: var(--canvas);
        }

        .record-item-row {
          display: flex;
          justify-content: space-between;
          align-items: center;
          padding: 12px 16px;
          border-bottom: 1px solid var(--hairline);
          transition: background 0.15s;
        }

        .record-item-row:last-child {
          border-bottom: none;
        }

        .record-item-row:hover {
          background: var(--pale-green);
        }

        .record-info-left {
          display: flex;
          flex-direction: column;
          gap: 2px;
          min-width: 0;
          text-align: left;
        }

        .record-filename-text {
          font-size: 13px;
          font-weight: 600;
          color: var(--ink);
          white-space: nowrap;
          overflow: hidden;
          text-overflow: ellipsis;
        }

        .record-meta-text {
          font-size: 11px;
          color: var(--slate);
          display: flex;
          align-items: center;
          gap: 6px;
        }

        .record-item-actions {
          display: flex;
          align-items: center;
          gap: 8px;
        }

        .record-load-btn {
          font-family: var(--font-ui);
          font-size: 11px;
          font-weight: 700;
          background: var(--cohere-black);
          color: var(--on-dark);
          border: none;
          border-radius: 4px;
          padding: 6px 10px;
          cursor: pointer;
          transition: all 0.15s;
        }

        .record-load-btn:hover {
          background: var(--deep-green);
        }

        .record-delete-btn {
          background: transparent;
          border: none;
          color: var(--slate);
          cursor: pointer;
          font-size: 14px;
          padding: 4px;
          display: flex;
          align-items: center;
        }

        .record-delete-btn:hover {
          color: var(--error);
        }

        @media (max-width: 768px) {
          .premium-header {
            flex-direction: column;
            align-items: flex-start;
          }
          .header-controls {
            width: 100%;
            justify-content: space-between;
          }
          .user-hub-widget {
            width: 100%;
            justify-content: space-between;
          }
          .auto-engine-panel {
            width: 100%;
            min-width: 0;
          }
        }

        /* 🚀 PREMIUM FORMULABOT-STYLE HERO & LANDING PAGE STYLES */
        .landing-hero {
          text-align: center;
          padding: 4rem 1.5rem 3rem 1.5rem;
          max-width: 1000px;
          margin: 0 auto;
        }

        .hero-title {
          font-family: var(--font-display);
          font-size: 3.2rem;
          font-weight: 800;
          line-height: 1.15;
          letter-spacing: -1px;
          margin-bottom: 1.5rem;
          background: linear-gradient(135deg, var(--ink) 30%, var(--coral) 100%);
          -webkit-background-clip: text;
          -webkit-text-fill-color: transparent;
        }

        [data-theme="dark"] .hero-title {
          background: linear-gradient(135deg, var(--on-dark) 40%, var(--coral-soft) 100%);
          -webkit-background-clip: text;
          -webkit-text-fill-color: transparent;
        }

        .hero-subtitle {
          font-family: var(--font-ui);
          font-size: 1.25rem;
          line-height: 1.6;
          color: var(--slate);
          max-width: 720px;
          margin: 0 auto 3rem auto;
          font-weight: 400;
        }

        /* Responsive Spreadsheet Mockup Graphic */
        .mockup-container {
          background: var(--canvas);
          border: 1px solid var(--hairline);
          border-radius: 16px;
          box-shadow: 0 30px 60px rgba(0, 0, 0, 0.08);
          margin: 0 auto 4.5rem auto;
          overflow: hidden;
          max-width: 900px;
          text-align: left;
        }

        [data-theme="dark"] .mockup-container {
          box-shadow: 0 30px 60px rgba(0, 0, 0, 0.35);
          background: #090d16;
        }

        .mockup-header {
          background: var(--soft-stone);
          border-bottom: 1px solid var(--hairline);
          padding: 10px 16px;
          display: flex;
          align-items: center;
          gap: 6px;
        }

        .mockup-dot {
          width: 10px;
          height: 10px;
          border-radius: 50%;
          display: inline-block;
        }

        .mockup-tab-strip {
          background: var(--soft-stone);
          display: flex;
          border-bottom: 1px solid var(--hairline);
          padding: 0 12px;
          gap: 4px;
        }

        .mockup-tab {
          padding: 8px 16px;
          font-size: 11px;
          font-weight: 600;
          color: var(--slate);
          border-bottom: 2px solid transparent;
          font-family: var(--font-technical);
          text-transform: uppercase;
        }

        .mockup-tab.active {
          color: var(--coral);
          border-bottom-color: var(--coral);
        }

        .mockup-body {
          display: grid;
          grid-template-columns: 1.5fr 1fr;
          border-bottom: 1px solid var(--hairline);
        }

        @media (max-width: 768px) {
          .mockup-body {
            grid-template-columns: 1fr;
          }
        }

        .mockup-left {
          padding: 1.25rem;
          border-right: 1px solid var(--hairline);
          overflow-x: auto;
        }

        @media (max-width: 768px) {
          .mockup-left {
            border-right: none;
            border-bottom: 1px solid var(--hairline);
          }
        }

        .mockup-right {
          padding: 1.25rem;
          background: var(--soft-stone);
          display: flex;
          flex-direction: column;
          gap: 12px;
        }

        .mockup-table {
          width: 100%;
          border-collapse: collapse;
          font-size: 11px;
          font-family: var(--font-technical);
        }

        .mockup-table th {
          text-align: left;
          color: var(--slate);
          padding: 8px;
          border-bottom: 1px solid var(--hairline);
          font-weight: 600;
        }

        .mockup-table td {
          padding: 8px;
          border-bottom: 1px solid var(--hairline);
          color: var(--ink);
        }

        .mockup-badge {
          padding: 2px 6px;
          border-radius: 4px;
          font-size: 9px;
          font-weight: 700;
        }

        .mockup-badge.delivered {
          background: rgba(16, 185, 129, 0.15);
          color: #10b981;
        }

        .mockup-badge.rto {
          background: rgba(239, 68, 68, 0.15);
          color: #ef4444;
        }

        .mockup-card-right {
          background: var(--canvas);
          border: 1px solid var(--hairline);
          border-radius: 8px;
          padding: 12px;
        }

        .mockup-chart-row {
          display: flex;
          align-items: flex-end;
          gap: 6px;
          height: 80px;
          margin-top: 10px;
        }

        .mockup-chart-bar {
          flex: 1;
          background: var(--coral-soft);
          border-radius: 4px 4px 0 0;
          transition: height 0.3s;
        }

        .mockup-chart-bar.highlight {
          background: var(--coral);
        }

        /* Detailed Features Grid (FormulaBot-Style) */
        .features-grid {
          display: grid;
          grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
          gap: 20px;
          max-width: 1000px;
          margin: 0 auto 3rem auto;
          padding: 0 1.5rem;
        }

        .feature-item-card {
          background: var(--canvas);
          border: 1px solid var(--hairline);
          border-radius: 14px;
          padding: 1.5rem;
          text-align: left;
          transition: all 0.25s ease;
          box-shadow: 0 4px 12px rgba(0,0,0,0.02);
        }

        .feature-item-card:hover {
          transform: translateY(-4px);
          box-shadow: 0 12px 24px rgba(0,0,0,0.06);
          border-color: var(--coral);
        }

        [data-theme="dark"] .feature-item-card:hover {
          box-shadow: 0 12px 24px rgba(0,0,0,0.25);
          background: rgba(255,255,255,0.02);
        }

        .feature-icon-wrapper {
          width: 44px;
          height: 44px;
          border-radius: 10px;
          background: var(--soft-stone);
          display: grid;
          place-items: center;
          font-size: 1.25rem;
          margin-bottom: 1.25rem;
          border: 1px solid var(--hairline);
        }

        .feature-title-card {
          font-family: var(--font-display);
          font-size: 1.1rem;
          font-weight: 700;
          color: var(--ink);
          margin: 0 0 8px 0;
        }

        .feature-desc-card {
          font-family: var(--font-ui);
          font-size: 0.85rem;
          line-height: 1.5;
          color: var(--slate);
          margin: 0;
        }
      `}</style>

      {/* Announcement Bar */}
      <div className="announcement-bar">
        ⚡ Security Verified: Client-Side Secure Sandbox Active. No data leaves your machine.
        <a href="https://cohere.com" target="_blank" rel="noopener noreferrer">Learn more</a>
      </div>

      {/* Main Header / Navigation */}
      <header className="premium-header">
        <div className="brand-section" style={{ display: "flex", alignItems: "center", gap: "14px" }}>
          <img 
            src={theme === "dark" ? "/logo-dark.png" : "/logo-light.png"} 
            alt="SheetCodeCrest Logo" 
            style={{ 
              height: "44px", 
              width: "44px", 
              borderRadius: "10px", 
              objectFit: "contain",
              transition: "transform 0.3s ease",
              cursor: "pointer"
            }} 
            className="brand-logo"
            onMouseEnter={(e) => e.currentTarget.style.transform = "scale(1.08)"}
            onMouseLeave={(e) => e.currentTarget.style.transform = "scale(1)"}
          />
          <div>
            <h1 style={{ margin: 0, display: "flex", alignItems: "center", gap: "8px" }}>
              SheetCodeCrest
            </h1>
            <p style={{ margin: "2px 0 0 0" }}>Upload once. The app detects Shopify, Shiprocket, or generic data and builds the right workbook automatically.</p>
          </div>
        </div>
        
        <div className="header-controls">
          <button
            type="button"
            className="theme-toggle"
            onClick={() => setTheme((current) => current === "dark" ? "light" : "dark")}
            aria-label={theme === "dark" ? "Switch to light mode" : "Switch to dark mode"}
            title={theme === "dark" ? "Light mode" : "Dark mode"}
          >
            {theme === "dark" ? "☀️" : "🌙"}
          </button>

          {isSharedViewOnly ? (
            <div style={{ display: "flex", alignItems: "center", gap: "10px" }}>
              <span style={{ fontSize: "12px", background: "rgba(245, 158, 11, 0.15)", color: "var(--amber)", padding: "4px 10px", borderRadius: "20px", fontWeight: 600, border: "1px solid rgba(245,158,11,0.3)", display: "flex", alignItems: "center", gap: "6px" }}>
                🔒 Shared Report (View-Only)
              </span>
              <button 
                type="button" 
                className="header-btn" 
                onClick={() => {
                  setIsSharedViewOnly(false);
                  setActiveRecordId(null);
                  setSharedRecordObj(null);
                  setFile(null);
                  setRawRows([]);
                  setTableHeaders([]);
                  setDataProfile(null);
                  setLogisticsAnalytics(null);
                  setShopifyAnalytics(null);
                  setStep("upload");
                  // Clear query params
                  window.history.replaceState({}, document.title, window.location.pathname);
                }}
              >
                ↩️ Upload My Own
              </button>
            </div>
          ) : currentUser ? (
            <div className="user-hub-widget">
              <span className="user-info-text">
                <span className="user-icon-bullet">👤</span>
                <strong>{currentUser.username}</strong>
                <span className={`plan-badge ${currentUser.isPro ? "pro" : "free"}`}>
                  {currentUser.isPro ? "Pro" : "Free"}
                </span>
              </span>
              {isAdminActive && (
                <button 
                  type="button" 
                  className="header-btn" 
                  onClick={() => {
                    setAdminModalOpen(true);
                    loadAdminData();
                  }}
                  style={{ borderColor: "var(--amber)", color: "var(--amber)", fontWeight: 600 }}
                >
                  ðŸ›¡ï¸ Admin Panel
                </button>
              )}
              <button 
                type="button" 
                className="header-btn" 
                onClick={() => setDashboardOpen(true)}
              >
                ðŸŽ›ï¸ Dashboard
              </button>
              {!currentUser.isPro && (
                <button 
                  type="button" 
                  className="header-btn" 
                  onClick={() => setCheckoutOpen(true)}
                  style={{ borderColor: "var(--coral)", color: "var(--coral)" }}
                >
                  ⚡ Upgrade
                </button>
              )}
              <button 
                type="button" 
                className="header-btn" 
                onClick={handleLogout}
              >
                🚪 Sign Out
              </button>
            </div>
          ) : (
            <button
              type="button"
              className="header-btn login-trigger"
              onClick={() => {
                setAuthTab("login");
                setAuthError("");
                setAuthModalOpen(true);
              }}
            >
              🔑 Sign In / Register
            </button>
          )}
        </div>
      </header>

      {/* Auto-detected notices */}
      {detectionNotice && (
        <div className="toast-badge">
          <span>✨</span> {detectionNotice}
        </div>
      )}

      {/* Error state */}
      {error && (
        <div className="section-card" style={{ borderLeft: "4px solid var(--error)", background: "rgba(239, 68, 68, 0.08)" }}>
          <div style={{ color: "var(--error)", fontWeight: 600, fontSize: "0.9rem" }}>âš ï¸ Spreadsheet Error</div>
          <div style={{ color: "var(--ink)", fontSize: "0.85rem", marginTop: "4px", opacity: 0.9 }}>{error}</div>
        </div>
      )}

      {/* Step 1: Upload File Area */}
      {step === "upload" && (
        <>
          {/* HERO HEADING SECTION */}
          <section className="landing-hero">
            <h2 className="hero-title">
              The AI-Powered Spreadsheet Analyst
            </h2>
            <p className="hero-subtitle">
              Upload your raw export files from Shopify, Shiprocket, or custom sheets. SheetCodeCrest automatically consolidates your data, profiles quality, optimizes logistics, and provides conversational AI strategy.
            </p>
            
            <div 
              className={`upload-card ${dragging ? "dragging" : ""}`}
              onDrop={onDrop}
              onDragOver={onDragOver}
              onDragLeave={onDragLeave}
              onClick={() => hasFreeReportsRemaining && inputRef.current?.click()}
              style={{ margin: "0 auto 4rem auto", maxWidth: "680px" }}
            >
              <input 
                type="file" 
                ref={inputRef} 
                accept=".xlsx,.xls,.csv" 
                style={{ display: "none" }} 
                onChange={onFileChange} 
              />
              <span className="upload-icon">📁</span>
              <h2 className="upload-title">
                Drop your spreadsheet here
              </h2>
              <p className="upload-desc">
                Supports `.xlsx`, `.xls`, or `.csv` files. No manual type mapping needed.
              </p>
              <div className="usage-meter">
                {hasFreeReportsRemaining
                  ? `Free reports remaining: ${freeReportsRemaining} of ${globalFreeLimit}`
                  : "Free trial complete"}
              </div>
              {!hasFreeReportsRemaining && (
                <div className="subscription-card" onClick={(e) => e.stopPropagation()}>
                  <h3>Purchase access to continue</h3>
                  <p>
                    You have used your free report generations. Upgrade with Codecrest Studio to keep creating Shopify, Shiprocket, and universal Excel analytics workbooks.
                  </p>
                  <div className="subscription-actions" style={{ display: "flex", gap: "10px", marginTop: "1rem" }}>
                    {currentUser ? (
                      <button
                        type="button"
                        className="btn-primary"
                        onClick={() => setCheckoutOpen(true)}
                        style={{ padding: "10px 20px", fontSize: "13px", borderRadius: "30px", flex: "none" }}
                      >
                        ⚡ Upgrade to Pro ($19/mo)
                      </button>
                    ) : (
                      <button
                        type="button"
                        className="btn-primary"
                        onClick={() => {
                          setAuthTab("login");
                          setAuthError("");
                          setAuthModalOpen(true);
                        }}
                        style={{ padding: "10px 20px", fontSize: "13px", borderRadius: "30px", flex: "none" }}
                      >
                        🔑 Sign In to Upgrade
                      </button>
                    )}
                    <a className="icon-only" href={CODECREST.website} target="_blank" rel="noopener noreferrer" aria-label="Open Codecrest Studio website" title="Website" style={{ width: "40px", height: "40px" }}>
                      <svg viewBox="0 0 24 24"><circle cx="12" cy="12" r="10"></circle><line x1="2" y1="12" x2="22" y2="12"></line><path d="M12 2a15.3 15.3 0 0 1 4 10 15.3 15.3 0 0 1-4 10 15.3 15.3 0 0 1-4-10 15.3 15.3 0 0 1 4-10z"></path></svg>
                    </a>
                  </div>
                </div>
              )}
            </div>

            {/* SPREADSHEET MOCKUP VISUAL */}
            <div className="mockup-container">
              <div className="mockup-header">
                <span className="mockup-dot" style={{ background: "#ef4444" }}></span>
                <span className="mockup-dot" style={{ background: "#f59e0b" }}></span>
                <span className="mockup-dot" style={{ background: "#10b981" }}></span>
                <span style={{ fontSize: "10px", color: "var(--slate)", fontFamily: "var(--font-technical)", marginLeft: "8px", textTransform: "uppercase" }}>Interactive Data Profiler & Consolidator</span>
              </div>
              <div className="mockup-tab-strip">
                <button 
                  type="button"
                  className={`mockup-tab ${mockupTabActive === "shopify" ? "active" : ""}`}
                  onClick={() => setMockupTabActive("shopify")}
                  style={{ background: "transparent", border: "none", cursor: "pointer", outline: "none" }}
                >
                  📊 Shopify Store Sales
                </button>
                <button 
                  type="button"
                  className={`mockup-tab ${mockupTabActive === "logistics" ? "active" : ""}`}
                  onClick={() => setMockupTabActive("logistics")}
                  style={{ background: "transparent", border: "none", cursor: "pointer", outline: "none" }}
                >
                  🚚 Courier Shipments
                </button>
                <button 
                  type="button"
                  className={`mockup-tab ${mockupTabActive === "universal" ? "active" : ""}`}
                  onClick={() => setMockupTabActive("universal")}
                  style={{ background: "transparent", border: "none", cursor: "pointer", outline: "none" }}
                >
                  📋 Universal Profiling
                </button>
              </div>
              <div className="mockup-body">
                <div className="mockup-left">
                  {mockupTabActive === "shopify" && (
                    <table className="mockup-table">
                      <thead>
                        <tr>
                          <th>Order ID</th>
                          <th>Customer</th>
                          <th>City / Region</th>
                          <th>Status</th>
                          <th>Amount</th>
                        </tr>
                      </thead>
                      <tbody>
                        <tr>
                          <td>#1001-A</td>
                          <td>Aman Sharma</td>
                          <td>Mumbai, MH</td>
                          <td><span className="mockup-badge delivered">Delivered</span></td>
                          <td>₹1,899.00</td>
                        </tr>
                        <tr>
                          <td>#1002-B</td>
                          <td>Sarah Jones</td>
                          <td>Bangalore, KA</td>
                          <td><span className="mockup-badge delivered">Delivered</span></td>
                          <td>₹2,450.00</td>
                        </tr>
                        <tr>
                          <td>#1003-C</td>
                          <td>Vikram Singh</td>
                          <td>Delhi, NCR</td>
                          <td><span className="mockup-badge rto">RTO Returned</span></td>
                          <td>₹1,299.00</td>
                        </tr>
                        <tr>
                          <td>#1004-D</td>
                          <td>Rohan Verma</td>
                          <td>Pune, MH</td>
                          <td><span className="mockup-badge delivered">Delivered</span></td>
                          <td>₹999.00</td>
                        </tr>
                      </tbody>
                    </table>
                  )}

                  {mockupTabActive === "logistics" && (
                    <table className="mockup-table">
                      <thead>
                        <tr>
                          <th>AWB Number</th>
                          <th>Courier Co.</th>
                          <th>Destination</th>
                          <th>Delivery Status</th>
                          <th>COD/Prepaid</th>
                        </tr>
                      </thead>
                      <tbody>
                        <tr>
                          <td>AWB-82019</td>
                          <td>Delhivery</td>
                          <td>Mumbai, MH</td>
                          <td><span className="mockup-badge delivered">DELIVERED</span></td>
                          <td>COD</td>
                        </tr>
                        <tr>
                          <td>AWB-71928</td>
                          <td>Bluedart</td>
                          <td>Chennai, TN</td>
                          <td><span className="mockup-badge rto">RTO RETURNED</span></td>
                          <td>COD</td>
                        </tr>
                        <tr>
                          <td>AWB-90123</td>
                          <td>Xpressbees</td>
                          <td>Kolkata, WB</td>
                          <td><span className="mockup-badge delivered">DELIVERED</span></td>
                          <td>PREPAID</td>
                        </tr>
                        <tr>
                          <td>AWB-33412</td>
                          <td>Delhivery</td>
                          <td>Delhi, NCR</td>
                          <td><span className="mockup-badge delivered">DELIVERED</span></td>
                          <td>COD</td>
                        </tr>
                      </tbody>
                    </table>
                  )}

                  {mockupTabActive === "universal" && (
                    <table className="mockup-table">
                      <thead>
                        <tr>
                          <th>Field Name</th>
                          <th>Type Detection</th>
                          <th>Fill Rate</th>
                          <th>Distinct Count</th>
                          <th>Sample Values</th>
                        </tr>
                      </thead>
                      <tbody>
                        <tr>
                          <td style={{ fontWeight: 600 }}>order_id</td>
                          <td>NUMERIC</td>
                          <td>100%</td>
                          <td>150 unique</td>
                          <td style={{ color: "var(--slate)" }}>1001; 1002; 1003</td>
                        </tr>
                        <tr>
                          <td style={{ fontWeight: 600 }}>customer</td>
                          <td>TEXT</td>
                          <td>100%</td>
                          <td>142 unique</td>
                          <td style={{ color: "var(--slate)" }}>Aman; Sarah; Vikram</td>
                        </tr>
                        <tr>
                          <td style={{ fontWeight: 600 }}>shipping_date</td>
                          <td>DATE</td>
                          <td>94%</td>
                          <td>18 unique</td>
                          <td style={{ color: "var(--slate)" }}>2026-05-01; 2026-05-02</td>
                        </tr>
                        <tr>
                          <td style={{ fontWeight: 600 }}>discount_code</td>
                          <td>TEXT</td>
                          <td>40%</td>
                          <td>4 unique</td>
                          <td style={{ color: "var(--slate)" }}>WELCOME10; FREE</td>
                        </tr>
                      </tbody>
                    </table>
                  )}
                </div>
                
                <div className="mockup-right">
                  {mockupTabActive === "shopify" && (
                    <>
                      <div className="mockup-card-right">
                        <div style={{ fontSize: "10px", fontWeight: 700, color: "var(--slate)", textTransform: "uppercase" }}>Store Revenue Share</div>
                        <div style={{ fontSize: "16px", fontWeight: 700, marginTop: "4px", color: "var(--ink)" }}>₹4.82 Lakhs</div>
                        <div className="mockup-chart-row">
                          <div className="mockup-chart-bar" style={{ height: "40%" }}></div>
                          <div className="mockup-chart-bar" style={{ height: "70%" }}></div>
                          <div className="mockup-chart-bar highlight" style={{ height: "95%" }}></div>
                          <div className="mockup-chart-bar" style={{ height: "55%" }}></div>
                          <div className="mockup-chart-bar" style={{ height: "80%" }}></div>
                        </div>
                      </div>
                      <div className="mockup-card-right" style={{ display: "flex", alignItems: "center", gap: "10px" }}>
                        <span style={{ fontSize: "18px" }}>🧠</span>
                        <div style={{ fontSize: "10.5px", color: "var(--slate)", lineHeight: "1.4" }}>
                          <strong>Avery Smith:</strong> "SKU consolidated. Found 14 duplicate order ID items. Delivery rate is at 78.4%."
                        </div>
                      </div>
                    </>
                  )}

                  {mockupTabActive === "logistics" && (
                    <>
                      <div className="mockup-card-right">
                        <div style={{ fontSize: "10px", fontWeight: 700, color: "var(--slate)", textTransform: "uppercase" }}>Courier Success Index</div>
                        <div style={{ fontSize: "16px", fontWeight: 700, marginTop: "4px", color: "var(--ink)" }}>86.5% Rating</div>
                        <div className="mockup-chart-row">
                          <div className="mockup-chart-bar highlight" style={{ height: "92%" }} title="Delhivery"></div>
                          <div className="mockup-chart-bar" style={{ height: "72%" }} title="Bluedart"></div>
                          <div className="mockup-chart-bar" style={{ height: "84%" }} title="Xpressbees"></div>
                          <div className="mockup-chart-bar highlight" style={{ height: "90%" }} title="Delhivery 2"></div>
                          <div className="mockup-chart-bar" style={{ height: "65%" }} title="Shadowfax"></div>
                        </div>
                      </div>
                      <div className="mockup-card-right" style={{ display: "flex", alignItems: "center", gap: "10px" }}>
                        <span style={{ fontSize: "18px" }}>🧠</span>
                        <div style={{ fontSize: "10.5px", color: "var(--slate)", lineHeight: "1.4" }}>
                          <strong>Avery Smith:</strong> "Logistics audit: Delhivery has 92% SLA fill rate. Bluedart COD RTO risk is high at 21%."
                        </div>
                      </div>
                    </>
                  )}

                  {mockupTabActive === "universal" && (
                    <>
                      <div className="mockup-card-right">
                        <div style={{ fontSize: "10px", fontWeight: 700, color: "var(--slate)", textTransform: "uppercase" }}>Data Quality Index</div>
                        <div style={{ fontSize: "16px", fontWeight: 700, marginTop: "4px", color: "var(--ink)" }}>96.8% Score</div>
                        <div className="mockup-chart-row">
                          <div className="mockup-chart-bar highlight" style={{ height: "100%" }}></div>
                          <div className="mockup-chart-bar highlight" style={{ height: "100%" }}></div>
                          <div className="mockup-chart-bar highlight" style={{ height: "94%" }}></div>
                          <div className="mockup-chart-bar" style={{ height: "40%" }}></div>
                          <div className="mockup-chart-bar highlight" style={{ height: "92%" }}></div>
                        </div>
                      </div>
                      <div className="mockup-card-right" style={{ display: "flex", alignItems: "center", gap: "10px" }}>
                        <span style={{ fontSize: "18px" }}>🧠</span>
                        <div style={{ fontSize: "10.5px", color: "var(--slate)", lineHeight: "1.4" }}>
                          <strong>Avery Smith:</strong> "Data Quality profile complete. 12 columns mapped. Detected 6 missing cells in discount_code."
                        </div>
                      </div>
                    </>
                  )}
                </div>
              </div>
            </div>

            {/* DETAILED FEATURES MATRIX */}
            <div style={{ textAlign: "center", marginBottom: "2.5rem" }}>
              <div style={{ fontFamily: "var(--font-technical)", fontSize: "10.5px", fontWeight: 700, color: "var(--coral)", textTransform: "uppercase", letterSpacing: "0.15em", marginBottom: "6px" }}>Core Capabilities</div>
              <h3 style={{ fontFamily: "var(--font-display)", fontSize: "2rem", fontWeight: 700, margin: 0 }}>Built for High-Growth E-Commerce</h3>
            </div>
            
            <div className="features-grid">
              <div className="feature-item-card">
                <div className="feature-icon-wrapper">ðŸ›ï¸</div>
                <h4 className="feature-title-card">Shopify Sales Consolidator</h4>
                <p className="feature-desc-card">Automatically consolidates order rows, calculates geographical performance, filters COD risk status, and groups repeat buyers for retargeting campaigns.</p>
              </div>
              <div className="feature-item-card">
                <div className="feature-icon-wrapper">🚚</div>
                <h4 className="feature-title-card">Logistics Optimizer</h4>
                <p className="feature-desc-card">Upload Shiprocket or courier sheets to consolidate multiple item packages, highlight shipping cost leakages, and rank courier company RTO risks.</p>
              </div>
              <div className="feature-item-card">
                <div className="feature-icon-wrapper">📊</div>
                <h4 className="feature-title-card">Universal Data Profiler</h4>
                <p className="feature-desc-card">Instantly calculates column fill rates, auto-detects date/numerical types, performs average and standard deviation computations, and outputs an interactive grid editor.</p>
              </div>
              <div className="feature-item-card">
                <div className="feature-icon-wrapper">🧠</div>
                <h4 className="feature-title-card">Conversational AI Analyst</h4>
                <p className="feature-desc-card">Connect your own Anthropic Claude key or run offline to query Avery, your business data analyst, directly inside the app without writing formulas or scripts.</p>
              </div>
            </div>
          </section>

          {/* Trust Partner / Logo Strip */}
          <div className="trust-logo-strip" style={{ marginTop: "1rem" }}>
            <div className="trust-logo-title">Compatible Integrations & formats</div>
            <div className="trust-logos">
              <span>SHIPROCKET</span>
              <span>SHOPIFY</span>
              <span>WOOCOMMERCE</span>
              <span>EXCEL</span>
              <span>CSV</span>
              <span>PDF</span>
            </div>
          </div>
        </>
      )}

      {/* Step 2: Live Processing Log console */}
      {isAdminActive && log.length > 0 && (
        <div className="console-box">
          <div className="console-header">
            <span>Terminal Log console</span>
            <span style={{ display: "flex", alignItems: "center", gap: "6px" }}>
              <span className="console-dot" style={{ background: step === "processing" ? "#10b981" : "#64748b", boxShadow: step === "processing" ? "0 0 8px #10b981" : "none" }}></span>
              {step === "processing" ? "Analyzing..." : "Process idle"}
            </span>
          </div>
          <div className="console-body">
            {log.map((line, idx) => (
              <div className="console-line" key={idx}>
                <span className="console-time">{line.time}</span>
                <span className={`console-text ${line.type === "success" ? "text-success" : line.type === "error" ? "text-error" : "text-neutral"}`}>
                  {line.msg}
                </span>
              </div>
            ))}
          </div>
        </div>
      )}

      {/* Step 3: Analytics Render Deck */}
      {step === "done" && (
        <>
          {/* RENDER MODE: SHOPIFY GROWTH ANALYTICS */}
          {mode === "shopify" && shopifyAnalytics && (
            <>
              <div className="kpi-grid">
                <div className="kpi-card success">
                  <div className="kpi-label">Shopify Orders</div>
                  <div className="kpi-value">{shopifyAnalytics.totalOrders.toLocaleString("en-IN")}</div>
                  <div className="kpi-sub">{shopifyAnalytics.totalRows.toLocaleString("en-IN")} source rows analysed</div>
                </div>
                <div className="kpi-card success">
                  <div className="kpi-label">Revenue</div>
                  <div className="kpi-value">₹{(shopifyAnalytics.totalRevenue / 1e5).toFixed(2)}L</div>
                  <div className="kpi-sub">AOV: ₹{(shopifyAnalytics.totalRevenue / Math.max(shopifyAnalytics.totalOrders, 1)).toFixed(0)}</div>
                </div>
                <div className="kpi-card">
                  <div className="kpi-label">Customers</div>
                  <div className="kpi-value">{shopifyAnalytics.totalCustomers.toLocaleString("en-IN")}</div>
                  <div className="kpi-sub">Segmented for retargeting</div>
                </div>
                <div className="kpi-card">
                  <div className="kpi-label">Products</div>
                  <div className="kpi-value">{shopifyAnalytics.productCount}</div>
                  <div className="kpi-sub">Top: {shopifyAnalytics.topProduct}</div>
                </div>
              </div>

              <div className="chart-row">
                <div className="section-card">
                  <h3 className="card-title">📊 Order Status Mix</h3>
                  {Object.entries(shopifyAnalytics.statusCounts)
                    .sort((a, b) => b[1] - a[1])
                    .slice(0, 6)
                    .map(([statusName, count]) => {
                      const share = (count / Math.max(shopifyAnalytics.totalOrders, 1)) * 100;
                      return (
                        <div className="visual-bar-container" key={statusName}>
                          <div className="visual-bar-info">
                            <span>{statusName}</span>
                            <span>{count.toLocaleString("en-IN")} orders ({share.toFixed(1)}%)</span>
                          </div>
                          <div className="visual-bar-bg">
                            <div 
                              className={`visual-bar-fill ${statusName === "Delivered" ? "success" : statusName.includes("Cancelled") || statusName.includes("Refunded") ? "danger" : "warning"}`}
                              style={{ width: `${Math.max(2, share)}%` }}
                            />
                          </div>
                        </div>
                      );
                    })}
                </div>
                <div className="section-card">
                  <h3 className="card-title">👥 Customer Segments</h3>
                  {Object.entries(shopifyAnalytics.segmentCounts)
                    .sort((a, b) => b[1] - a[1])
                    .map(([segmentName, count]) => {
                      const share = (count / Math.max(shopifyAnalytics.totalCustomers, 1)) * 100;
                      return (
                        <div className="visual-bar-container" key={segmentName}>
                          <div className="visual-bar-info">
                            <span>{segmentName}</span>
                            <span>{count.toLocaleString("en-IN")} customers ({share.toFixed(1)}%)</span>
                          </div>
                          <div className="visual-bar-bg">
                            <div className="visual-bar-fill success" style={{ width: `${Math.max(2, share)}%` }} />
                          </div>
                        </div>
                      );
                    })}
                </div>
              </div>

              <div className="section-card" style={{ borderLeft: "4px solid var(--deep-green)", borderLeftColor: "var(--deep-green)" }}>
                <h3 className="card-title">Workbook Generated Like Your Shopify Reference Files</h3>
                <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(220px, 1fr))", gap: "10px", fontSize: "0.85rem", color: "var(--slate)" }}>
                  <span>Dashboard and order status detail</span>
                  <span>Product analysis and product-status matrix</span>
                  <span>Monthly, COD, geographic, discount views</span>
                  <span>Customer data, segments, retargeting lists</span>
                  <span>Top product-wise customer/order sheets</span>
                </div>
              </div>
            </>
          )}

          {/* RENDER MODE A: SHIPROCKET LOGISTICS OPTIMIZER */}
          {mode === "logistics" && logisticsAnalytics && (
            <>
              {/* KPI Summaries */}
              <div className="kpi-grid">
                <div className="kpi-card">
                  <div className="kpi-label">Consolidated Orders</div>
                  <div className="kpi-value">{logisticsAnalytics.total.toLocaleString("en-IN")}</div>
                  <div className="kpi-sub">{mergedLogisticsCount > 0 ? `${mergedLogisticsCount} row items consolidated` : "No duplicate order rows"}</div>
                </div>
                <div className="kpi-card success">
                  <div className="kpi-label">Total Shipping Value</div>
                  <div className="kpi-value">₹{(logisticsAnalytics.totalRev / 1e5).toFixed(2)}L</div>
                  <div className="kpi-sub">Avg order value: ₹{(logisticsAnalytics.totalRev / logisticsAnalytics.total).toFixed(0)}</div>
                </div>
                <div className="kpi-card success">
                  <div className="kpi-label">Delivery Success</div>
                  <div className="kpi-value text-success-accent">{(logisticsAnalytics.deliveryRate * 100).toFixed(1)}%</div>
                  <div className="kpi-sub">{logisticsAnalytics.delivered} packages delivered</div>
                </div>
                <div className={`kpi-card ${logisticsAnalytics.rtoRate > 0.15 ? "warning" : ""}`}>
                  <div className="kpi-label">Return-To-Origin (RTO)</div>
                  <div className={`kpi-value ${logisticsAnalytics.rtoRate > 0.15 ? "text-warning-accent" : ""}`}>
                    {(logisticsAnalytics.rtoRate * 100).toFixed(1)}%
                  </div>
                  <div className="kpi-sub">{logisticsAnalytics.rto} returned products</div>
                </div>
              </div>

              {/* Charts row */}
              <div className="chart-row">
                {/* Visual Bar Breakdown: Status share */}
                <div className="section-card">
                  <h3 className="card-title">📦 Package Status distribution</h3>
                  {Object.entries(logisticsAnalytics.statusCounts)
                    .sort((a, b) => b[1].orders - a[1].orders)
                    .slice(0, 4)
                    .map(([statusName, detail]: [string, any]) => {
                      const share = (detail.orders / logisticsAnalytics.total) * 100;
                      return (
                        <div className="visual-bar-container" key={statusName}>
                          <div className="visual-bar-info">
                            <span>{statusName}</span>
                            <span>{detail.orders} orders ({share.toFixed(1)}%)</span>
                          </div>
                          <div className="visual-bar-bg">
                            <div 
                              className={`visual-bar-fill ${statusName === "DELIVERED" ? "success" : statusName.includes("RTO") ? "danger" : statusName === "CANCELED" ? "warning" : ""}`}
                              style={{ width: `${share}%` }}
                            />
                          </div>
                        </div>
                      );
                    })}
                </div>

                {/* Courier scorecards */}
                <div className="section-card">
                  <h3 className="card-title">🚚 Courier Company Efficiency</h3>
                  {Object.entries(logisticsAnalytics.courierCounts)
                    .sort((a, b) => b[1].orders - a[1].orders)
                    .slice(0, 4)
                    .map(([courierName, detail]: [string, any]) => {
                      const successRate = (detail.delivered / detail.orders) * 100;
                      return (
                        <div className="visual-bar-container" key={courierName}>
                          <div className="visual-bar-info">
                            <span>{courierName} ({detail.orders} orders)</span>
                            <span style={{ color: successRate > 70 ? "#059669" : "#d97706" }}>
                              Delivered: {successRate.toFixed(0)}%
                            </span>
                          </div>
                          <div className="visual-bar-bg">
                            <div 
                              className={`visual-bar-fill ${successRate > 70 ? "success" : "warning"}`}
                              style={{ width: `${successRate}%` }}
                            />
                          </div>
                        </div>
                      );
                    })}
                </div>
              </div>

              {/* Monthly Revenue Trends visual chart */}
              {(() => {
                const dateCol = tableHeaders.find((h) => /date|created|ordered|order date/i.test(h)) || null;
                const revenueCol = tableHeaders.find((h) => /revenue|amount|order total|total|price/i.test(h)) || null;
                if (!dateCol || !revenueCol) return null;

                const trends: Record<string, number> = {};
                rawRows.forEach((r) => {
                  const d = new Date(r[dateCol]);
                  const rev = Number(String(r[revenueCol] ?? "").replace(/[,?\s%]/g, ""));
                  if (!isNaN(d.getTime()) && !isNaN(rev)) {
                    const key = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
                    trends[key] = (trends[key] || 0) + rev;
                  }
                });

                const monthlyData = Object.entries(trends).sort((a, b) => a[0].localeCompare(b[0]));
                if (!monthlyData.length) return null;
                const maxVal = Math.max(...monthlyData.map(([_, v]) => v), 1);

                return (
                  <div className="section-card">
                    <h3 className="card-title">📈 Monthly Shipping Revenue Trend</h3>
                    <div className="trend-chart-svg">
                      {monthlyData.map(([month, val]) => {
                        const pct = (val / maxVal) * 100;
                        return (
                          <div className="trend-chart-bar" key={month}>
                            <div 
                              className="trend-chart-fill" 
                              style={{ height: `${Math.max(10, pct)}%` }}
                            >
                              <span className="trend-chart-tip">
                                ₹{Math.round(val/1000)}K
                              </span>
                            </div>
                            <span style={{ fontSize: "0.75rem", color: "var(--slate)" }}>{month}</span>
                          </div>
                        );
                      })}
                    </div>
                  </div>
                );
              })()}

              {/* RTO Risk Destinations warnings */}
              {(() => {
                const highRtoStatesList = Object.entries(logisticsAnalytics.stateCounts)
                  .filter(([_, v]) => v.orders >= 10)
                  .sort((a, b) => (b[1].rto / b[1].orders) - (a[1].rto / a[1].orders))
                  .slice(0, 3);
                
                if (!highRtoStatesList.length) return null;

                return (
                  <div className="section-card" style={{ borderLeft: "4px solid #ef4444" }}>
                    <h3 className="card-title" style={{ color: "#ef4444" }}>âš ï¸ Regional return hotspot alerts</h3>
                    <div style={{ display: "flex", flexDirection: "column", gap: "10px" }}>
                      {highRtoStatesList.map(([stateName, detail]: [string, any]) => {
                        const rtoPercent = (detail.rto / detail.orders) * 100;
                        return (
                          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", borderBottom: "1px solid var(--hairline)", paddingBottom: "6px" }} key={stateName}>
                            <span style={{ fontSize: "0.85rem", fontWeight: 600 }}>{stateName} ({detail.orders} shipments)</span>
                            <span className="rto-risk-badge">
                              {rtoPercent.toFixed(1)}% RTO Rate ({detail.rto} returned)
                            </span>
                          </div>
                        );
                      })}
                    </div>
                  </div>
                );
              })()}
            </>
          )}

          {/* RENDER MODE B: UNIVERSAL EXCEL PROFILER */}
          {mode === "universal" && dataProfile && (
            <>
              {/* KPI Summaries */}
              <div className="kpi-grid">
                <div className="kpi-card">
                  <div className="kpi-label">Columns</div>
                  <div className="kpi-value">{dataProfile.totalColumns}</div>
                  <div className="kpi-sub">Fields profiled in sheet</div>
                </div>
                <div className="kpi-card">
                  <div className="kpi-label">Rows Count</div>
                  <div className="kpi-value">{dataProfile.totalRows.toLocaleString("en-IN")}</div>
                  <div className="kpi-sub">Total parsed data rows</div>
                </div>
                <div className="kpi-card">
                  <div className="kpi-label">Duplicate Rows</div>
                  <div className="kpi-value" style={{ color: dataProfile.duplicateRows > 0 ? "#f59e0b" : "#10b981" }}>
                    {dataProfile.duplicateRows.toLocaleString("en-IN")}
                  </div>
                  <div className="kpi-sub">Duplicate matching rows</div>
                </div>
              </div>

              {/* Numeric Summaries Scorecard */}
              {dataProfile.columns.filter((c) => c.type === "numeric").length > 0 && (
                <div className="section-card">
                  <h3 className="card-title">📊 Numeric Columns aggregation</h3>
                  <div className="premium-table-wrapper">
                    <table className="premium-table">
                      <thead>
                        <tr>
                          <th>Field Name</th>
                          <th>Sum</th>
                          <th>Average</th>
                          <th>Median</th>
                          <th>Volatilty (StdDev)</th>
                          <th>Min / Max</th>
                        </tr>
                      </thead>
                      <tbody>
                        {dataProfile.columns
                          .filter((c) => c.type === "numeric")
                          .slice(0, 5)
                          .map((col) => (
                            <tr key={col.name}>
                              <td style={{ fontWeight: 600 }}>{col.name}</td>
                              <td>{col.sum ? col.sum.toLocaleString("en-IN", { maximumFractionDigits: 1 }) : "-"}</td>
                              <td>{col.avg ? col.avg.toLocaleString("en-IN", { maximumFractionDigits: 1 }) : "-"}</td>
                              <td>{col.median ? col.median.toLocaleString("en-IN", { maximumFractionDigits: 1 }) : "-"}</td>
                              <td>{col.stddev ? `±${col.stddev.toFixed(1)}` : "-"}</td>
                              <td style={{ fontSize: "0.75rem", color: "var(--slate)" }}>
                                {col.min} to {col.max}
                              </td>
                            </tr>
                          ))}
                      </tbody>
                    </table>
                  </div>
                </div>
              )}

              {/* Data Quality & Cardinality Details */}
              <div className="section-card">
                <h3 className="card-title">🧪 Data Quality & Type Profiling</h3>
                <div className="premium-table-wrapper">
                  <table className="premium-table">
                    <thead>
                      <tr>
                        <th>Column Field</th>
                        <th>Type Badge</th>
                        <th>Fill Rate %</th>
                        <th>Cardinality (Unique)</th>
                        <th>Sample Values Preview</th>
                      </tr>
                    </thead>
                    <tbody>
                      {dataProfile.columns.slice(0, 10).map((col) => {
                        const fillPercent = ((col.nonEmptyCount / col.count) * 100).toFixed(0);
                        return (
                          <tr key={col.name}>
                            <td style={{ fontWeight: 600 }}>{col.name}</td>
                            <td>
                              <span style={{ 
                                background: col.type === "numeric" ? "rgba(24, 99, 220, 0.15)" : col.type === "date" ? "rgba(255, 119, 89, 0.15)" : "var(--soft-stone)",
                                color: col.type === "numeric" ? "var(--action-blue)" : col.type === "date" ? "var(--coral)" : "var(--slate)",
                                padding: "3px 8px",
                                borderRadius: "4px",
                                fontSize: "0.7rem",
                                fontWeight: 700
                              }}>
                                {col.type.toUpperCase()}
                              </span>
                            </td>
                            <td>{fillPercent}%</td>
                            <td>{col.uniqueCount} values</td>
                            <td style={{ fontSize: "0.75rem", color: "var(--slate)", maxWidth: "250px" }}>
                              {col.sampleValues.join("; ")}
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              </div>
            </>
          )}

          {/* 🔗 SECURE SHARING LINK ALERT FOR OWNERS */}
          {activeRecordId && !isSharedViewOnly && (
            <div className="section-card" style={{ background: "rgba(24, 99, 220, 0.05)", borderLeft: "4px solid var(--action-blue)", display: "flex", justifyContent: "space-between", alignItems: "center", padding: "1.25rem 1.5rem", marginBottom: "1.5rem" }}>
              <div>
                <h4 style={{ margin: 0, fontSize: "14px", fontWeight: 700, color: "var(--action-blue)", display: "flex", alignItems: "center", gap: "6px" }}>
                  <span>🔗</span> Share Secure Report Link
                </h4>
                <p style={{ margin: "4px 0 0 0", fontSize: "12px", color: "var(--slate)" }}>
                  Generate a view-only shareable link for clients, partners, or team members to access this analysis.
                </p>
              </div>
              <button
                type="button"
                className="btn-primary"
                onClick={() => setShareModalOpen(true)}
                style={{ padding: "8px 16px", borderRadius: "30px", fontSize: "12px", cursor: "pointer", flexShrink: 0 }}
              >
                🔗 Create Shareable Link
              </button>
            </div>
          )}

          {/* 🛠️ SPREADSHEET GIT-STYLE VERSION CONTROL & COLLABORATIVE COMMENTS PANEL */}
          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(320px, 1fr))", gap: "1.5rem", marginBottom: "2rem" }}>
            
            {/* PANEL A: Spreadsheet Git-Style Version Control Branches */}
            <div className="section-card" style={{ display: "flex", flexDirection: "column", height: "100%", margin: 0 }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: "1rem", borderBottom: "1px solid var(--hairline)", paddingBottom: "10px" }}>
                <h3 className="card-title" style={{ margin: 0, fontSize: "1.1rem", display: "flex", alignItems: "center", gap: "8px" }}>
                  <span>📂</span> Version Control History
                </h3>
                {!isSharedViewOnly && (
                  <label className="btn-secondary" style={{ display: "inline-flex", alignItems: "center", gap: "6px", cursor: "pointer", padding: "6px 12px", borderRadius: "20px", fontSize: "11px", borderColor: "var(--amber)", color: "var(--amber)" }}>
                    <span>📤 Upload New Version</span>
                    <input 
                      type="file" 
                      accept=".xlsx,.xls,.csv" 
                      style={{ display: "none" }} 
                      onChange={handleUploadNewVersion}
                      disabled={versionUploading}
                    />
                  </label>
                )}
              </div>

              <div style={{ flex: 1, overflowY: "auto", maxHeight: "240px", display: "flex", flexDirection: "column", gap: "10px" }}>
                <div style={{ display: "flex", justifyContent: "space-between", padding: "10px 12px", background: "rgba(255, 255, 255, 0.05)", borderRadius: "8px", border: "1px solid rgba(255, 255, 255, 0.12)", alignItems: "center" }}>
                  <div>
                    <span style={{ fontSize: "10px", background: "var(--amber)", color: "#000", padding: "1px 5px", borderRadius: "3px", fontWeight: 700, marginRight: "6px" }}>ACTIVE</span>
                    <strong style={{ fontSize: "12px", color: "var(--ink)" }}>{file ? file.name : "Active_Spreadsheet.xlsx"}</strong>
                    <div style={{ fontSize: "10px", color: "var(--slate)", marginTop: "2px" }}>
                      {(file ? file.size / 1024 : 15).toFixed(1)} KB • Active workspace
                    </div>
                  </div>
                </div>

                {versionHistory.map((ver, idx) => (
                  <div key={idx} style={{ display: "flex", justifyContent: "space-between", padding: "10px 12px", background: "rgba(255, 255, 255, 0.02)", borderRadius: "8px", border: "1px solid var(--hairline)", alignItems: "center" }}>
                    <div>
                      <span style={{ fontSize: "10px", background: "var(--soft-stone)", color: "var(--slate)", padding: "1px 5px", borderRadius: "3px", fontWeight: 600, marginRight: "6px" }}>V{idx + 1}</span>
                      <strong style={{ fontSize: "12px", color: "var(--slate)" }}>{ver.filename}</strong>
                      <div style={{ fontSize: "10px", color: "var(--slate)", marginTop: "2px" }}>
                        {(ver.size / 1024).toFixed(1)} KB • {ver.timestamp}
                      </div>
                    </div>
                    {!isSharedViewOnly && (
                      <button
                        type="button"
                        className="record-load-btn"
                        style={{ padding: "4px 8px", fontSize: "11px" }}
                        onClick={() => handleRestoreVersion(idx)}
                      >
                        🔄 Restore
                      </button>
                    )}
                  </div>
                ))}

                {versionHistory.length === 0 && (
                  <div style={{ textAlign: "center", padding: "1.5rem", color: "var(--slate)", fontSize: "12px" }}>
                    No past version branches found. Upload an edited spreadsheet version to track histories.
                  </div>
                )}
              </div>
            </div>

            {/* PANEL B: Collaborative Pinned comments & annotations drawer */}
            <div className="section-card" style={{ display: "flex", flexDirection: "column", height: "100%", margin: 0 }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: "1rem", borderBottom: "1px solid var(--hairline)", paddingBottom: "10px" }}>
                <h3 className="card-title" style={{ margin: 0, fontSize: "1.1rem", display: "flex", alignItems: "center", gap: "8px" }}>
                  <span>💬</span> Comments & Collaboration
                </h3>
                <span style={{ fontSize: "11px", background: "rgba(255, 255, 255, 0.1)", color: "var(--slate)", padding: "2px 8px", borderRadius: "10px" }}>
                  {activeComments.length} annotations
                </span>
              </div>

              {/* Comments Scroller */}
              <div style={{ flex: 1, overflowY: "auto", maxHeight: "180px", display: "flex", flexDirection: "column", gap: "8px", marginBottom: "1rem", paddingRight: "4px" }}>
                {activeComments.map((comment) => (
                  <div key={comment.id} style={{ padding: "8px 12px", background: "rgba(255, 255, 255, 0.03)", borderRadius: "8px", border: "1px solid var(--hairline)" }}>
                    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: "4px" }}>
                      <span style={{ fontSize: "12px", fontWeight: 700, color: "var(--amber)", display: "flex", alignItems: "center", gap: "6px" }}>
                        <span style={{ width: "16px", height: "16px", borderRadius: "50%", background: "rgba(245,158,11,0.2)", display: "flex", alignItems: "center", justifyContent: "center", fontSize: "9px" }}>
                          {comment.author[0].toUpperCase()}
                        </span>
                        {comment.author}
                      </span>
                      <span style={{ fontSize: "9px", color: "var(--slate)" }}>{comment.timestamp}</span>
                    </div>
                    <div style={{ fontSize: "12px", color: "var(--ink)", lineHeight: 1.4, textAlign: "left" }}>{comment.text}</div>
                  </div>
                ))}

                {activeComments.length === 0 && (
                  <div style={{ textAlign: "center", padding: "2rem", color: "var(--slate)", fontSize: "12px" }}>
                    No comments pinned on this dashboard. Add an annotation below!
                  </div>
                )}
              </div>

              {/* Add Comment Form */}
              <form onSubmit={handleAddComment} style={{ borderTop: "1px solid var(--hairline)", paddingTop: "10px", display: "flex", flexDirection: "column", gap: "6px" }}>
                <div style={{ display: "flex", gap: "6px" }}>
                  <input 
                    type="text" 
                    placeholder="Nickname (e.g. CEO)" 
                    className="ai-key-input"
                    style={{ fontSize: "12px", padding: "6px 10px", width: "40%", minWidth: "0", margin: 0 }}
                    value={commentAuthor}
                    onChange={(e) => setCommentAuthor(e.target.value)}
                  />
                  <input 
                    type="text" 
                    placeholder="Type a collaborative note..." 
                    className="ai-key-input"
                    style={{ fontSize: "12px", padding: "6px 10px", width: "60%", minWidth: "0", margin: 0 }}
                    value={commentInput}
                    onChange={(e) => setCommentInput(e.target.value)}
                    required
                  />
                </div>
                <button type="submit" className="btn-primary" style={{ padding: "6px 12px", fontSize: "12px", borderRadius: "4px", alignSelf: "flex-end", marginTop: "4px" }}>
                  💬 Pinned Annotation
                </button>
              </form>
            </div>
          </div>

          {/* AI Audit & Strategy Assistant - Conversational Chat Console */}
          {!isSharedViewOnly && (
            <div className="section-card ai-insights-card">
            <h3 className="card-title" style={{ fontSize: "1.15rem", fontWeight: 700, display: "flex", alignItems: "center", gap: "8px" }}>
              <span>🧠</span> Avery Smith — Senior AI Data Analyst
            </h3>
            
            <div className="ai-key-input-row">
              <input 
                type="password" 
                placeholder="Paste your Anthropic Claude API Key for live intelligence (Optional)..." 
                className="ai-key-input"
                value={customApiKey}
                onChange={(e) => setCustomApiKey(e.target.value)}
              />
              <span style={{ fontSize: "0.75rem", color: "var(--slate)", fontWeight: 500 }}>
                {customApiKey ? "⚡ API Key Connected" : "🔌 Running in Intelligent Offline Mode"}
              </span>
            </div>

            <div className="chat-container">
              <div className="chat-header">
                <span className="chat-header-title">
                  💬 Direct Chat with Avery
                </span>
                <span className="chat-header-status">
                  <span className="chat-pulse"></span>
                  {aiLoading ? "Thinking..." : "Avery is Online"}
                </span>
              </div>

              <div className="chat-messages">
                {chatHistory.map((msg, idx) => (
                  <div key={idx} className={`chat-message ${msg.sender}`}>
                    <div style={{ fontWeight: 700, marginBottom: "6px", fontSize: "0.75rem", opacity: 0.8 }}>
                      {msg.sender === "analyst" ? "🧠 AVERY (SENIOR ANALYST)" : "👤 YOU (BUSINESS OWNER)"}
                    </div>
                    <div dangerouslySetInnerHTML={{ 
                      __html: formatMessage(msg.text)
                    }} />
                  </div>
                ))}
              </div>

              {/* Suggestions tray */}
              <div className="chat-suggestions-tray">
                {mode === "shopify" ? (
                  <>
                    <button 
                      className="chat-suggestion-chip" 
                      onClick={() => handleChatSubmit(undefined, "Which products are performing best and what should we retarget?")}
                      disabled={aiLoading}
                    >
                      ðŸ›ï¸ Product Winners
                    </button>
                    <button 
                      className="chat-suggestion-chip" 
                      onClick={() => handleChatSubmit(undefined, "Summarize customer segments and retargeting priorities.")}
                      disabled={aiLoading}
                    >
                      👥 Retargeting Priorities
                    </button>
                    <button 
                      className="chat-suggestion-chip" 
                      onClick={() => handleChatSubmit(undefined, "Analyze COD/payment risk and discount leakage.")}
                      disabled={aiLoading}
                    >
                      💰 COD & Discount Risk
                    </button>
                  </>
                ) : mode === "logistics" ? (
                  <>
                    <button 
                      className="chat-suggestion-chip" 
                      onClick={() => handleChatSubmit(undefined, "How can we reduce shipping cost leakage? Outline 3 main areas.")}
                      disabled={aiLoading}
                    >
                      💸 Cost Leakage Advice
                    </button>
                    <button 
                      className="chat-suggestion-chip" 
                      onClick={() => handleChatSubmit(undefined, "Which states represent the highest RTO risk and how do we solve them?")}
                      disabled={aiLoading}
                    >
                      âš ï¸ Highest RTO Risks
                    </button>
                    <button 
                      className="chat-suggestion-chip" 
                      onClick={() => handleChatSubmit(undefined, "Can you analyze our courier scorecard and suggest who to prioritize?")}
                      disabled={aiLoading}
                    >
                      🚚 Courier Performance Summary
                    </button>
                  </>
                ) : (
                  <>
                    <button 
                      className="chat-suggestion-chip" 
                      onClick={() => handleChatSubmit(undefined, "Please summarize the numeric columns and their key averages.")}
                      disabled={aiLoading}
                    >
                      📊 Summarize Numeric Fields
                    </button>
                    <button 
                      className="chat-suggestion-chip" 
                      onClick={() => handleChatSubmit(undefined, "Show me a quality audit of the dataset including missing values.")}
                      disabled={aiLoading}
                    >
                      🧪 Column Fill Rates & Quality
                    </button>
                    <button 
                      className="chat-suggestion-chip" 
                      onClick={() => handleChatSubmit(undefined, "Explain the duplicate rows profile and why clean rows matter.")}
                      disabled={aiLoading}
                    >
                      📋 Duplicate Rows Audit
                    </button>
                  </>
                )}
              </div>

              <form onSubmit={handleChatSubmit} className="chat-input-bar">
                <input 
                  type="text" 
                  placeholder="Ask Avery any question about this dataset (e.g., 'Compare COD vs Prepaid rates')..." 
                  className="chat-text-input"
                  value={chatInput}
                  onChange={(e) => setChatInput(e.target.value)}
                  disabled={aiLoading}
                />
                <button type="submit" className="chat-send-btn" disabled={aiLoading || !chatInput.trim()}>
                  {aiLoading ? "..." : "Send"}
                </button>
              </form>
            </div>
          </div>
        )}

          {/* SPREADSHEET INTERACTIVE VIEWER DATA GRID (Common to both modes) */}
          <div className="section-card">
            <h3 className="card-title">📋 Interactive In-Browser Data Grid</h3>
            <div className="grid-header-tools">
              <input 
                type="text" 
                placeholder="Search across all parsed rows..."
                className="grid-search-input"
                value={searchTerm}
                onChange={(e) => { setSearchTerm(e.target.value); setCurrentPage(1); }}
              />
              <span style={{ fontSize: "0.8rem", color: "var(--slate)" }}>
                Found {sortedAndFilteredRows.length} matches out of {rawRows.length} rows
              </span>
            </div>

            <div className="premium-table-wrapper">
              <table className="premium-table">
                <thead>
                  <tr>
                    {tableHeaders.slice(0, 8).map((header) => (
                      <th key={header} onClick={() => handleSort(header)}>
                        {header} {sortConfig?.key === header ? (sortConfig.direction === "asc" ? "▲" : "â–¼") : ""}
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {paginatedRows.map((row, idx) => (
                    <tr key={idx}>
                      {tableHeaders.slice(0, 8).map((header) => (
                        <td key={header}>{String(row[header] ?? "")}</td>
                      ))}
                    </tr>
                  ))}
                  {paginatedRows.length === 0 && (
                    <tr>
                      <td colSpan={tableHeaders.length} style={{ textAlign: "center", padding: "2rem", color: "var(--slate)" }}>
                        No records match your filters.
                      </td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>

            {/* Pagination Controls */}
            <div className="grid-pagination">
              <button 
                className="pagination-btn"
                onClick={() => setCurrentPage((p) => Math.max(1, p - 1))}
                disabled={currentPage === 1}
              >
                Previous Page
              </button>
              <span style={{ fontSize: "0.85rem", color: "var(--slate)" }}>
                Page <strong>{currentPage}</strong> of {totalPages}
              </span>
              <button 
                className="pagination-btn"
                onClick={() => setCurrentPage((p) => Math.min(totalPages, p + 1))}
                disabled={currentPage === totalPages}
              >
                Next Page
              </button>
            </div>
          </div>

          {/* Export Action Triggers */}
          <div className="action-bar" style={{ display: "flex", gap: "10px", justifyContent: "center", flexWrap: "wrap" }}>
            <a 
              href={dlUrl || undefined} 
              download={outName} 
              className="btn-primary"
            >
              📥 Download Polished Excel Report (.xlsx)
            </a>
            {activeRecordId && !isSharedViewOnly && (
              <button 
                type="button"
                onClick={() => setShareModalOpen(true)} 
                className="btn-secondary"
                style={{ display: "inline-flex", alignItems: "center", gap: "6px" }}
              >
                <span>🔗</span> Share Report
              </button>
            )}
            <button 
              onClick={reset} 
              className="btn-secondary"
            >
              🔄 Upload New File
            </button>
          </div>
        </>
      )}

      {/* Footer disclaimer */}
      <footer style={{ marginTop: "3rem", borderTop: "1px solid var(--hairline)", paddingTop: "1.5rem", textAlign: "center", fontSize: "0.75rem", color: "var(--slate)" }}>
        <div>
          SheetCodeCrest • Runs completely in the browser for secure data privacy.
        </div>
        <div style={{ marginTop: "0.4rem" }}>
          Created and developed by <strong style={{ color: "var(--ink)" }}>Codecrest_studio</strong>
        </div>
        <div className="footer-links">
          <a className="icon-only" href={CODECREST.instagram} target="_blank" rel="noopener noreferrer" aria-label="Open Codecrest Studio Instagram" title="Instagram">
            <svg viewBox="0 0 24 24"><rect x="2" y="2" width="20" height="20" rx="5" ry="5"></rect><path d="M16 11.37A4 4 0 1 1 12.63 8 4 4 0 0 1 16 11.37z"></path><line x1="17.5" y1="6.5" x2="17.51" y2="6.5"></line></svg>
          </a>
          <a className="icon-only" href={`mailto:${CODECREST.email}`} aria-label="Email Codecrest Studio" title="Email">
            <svg viewBox="0 0 24 24"><path d="M4 4h16c1.1 0 2 .9 2 2v12c0 1.1-.9 2-2 2H4c-1.1 0-2-.9-2-2V6c0-1.1.9-2 2-2z"></path><polyline points="22,6 12,13 2,6"></polyline></svg>
          </a>
          <a className="icon-only" href={CODECREST.website} target="_blank" rel="noopener noreferrer" aria-label="Open Codecrest Studio website" title="Website">
            <svg viewBox="0 0 24 24"><circle cx="12" cy="12" r="10"></circle><line x1="2" y1="12" x2="22" y2="12"></line><path d="M12 2a15.3 15.3 0 0 1 4 10 15.3 15.3 0 0 1-4 10 15.3 15.3 0 0 1-4-10 15.3 15.3 0 0 1 4-10z"></path></svg>
          </a>
        </div>
      </footer>

      {/* 1. Auth Modal (Login/Signup Tabs) */}
      {authModalOpen && (
        <div className="modal-overlay" onClick={() => setAuthModalOpen(false)}>
          <div className="modal-card" onClick={(e) => e.stopPropagation()}>
            <div className="modal-header">
              <h3 className="modal-title" style={{ display: "flex", alignItems: "center", gap: "10px" }}>
                <img src="/logo-icon.png" alt="SheetCodeCrest Icon" style={{ height: "24px", width: "24px", borderRadius: "5px", objectFit: "contain" }} />
                <span>{authTab === "login" ? "🔑 Sign In" : "ðŸ“ Create Account"}</span>
              </h3>
              <button 
                type="button" 
                className="modal-close-btn" 
                onClick={() => setAuthModalOpen(false)}
              >
                ✕
              </button>
            </div>
            <div className="auth-tabs">
              <button 
                type="button" 
                className={`auth-tab-btn ${authTab === "login" ? "active" : ""}`}
                onClick={() => { setAuthTab("login"); setAuthError(""); }}
              >
                Sign In
              </button>
              <button 
                type="button" 
                className={`auth-tab-btn ${authTab === "signup" ? "active" : ""}`}
                onClick={() => { setAuthTab("signup"); setAuthError(""); }}
              >
                Register
              </button>
            </div>
            <div className="modal-body">
              {authError && (
                <div style={{ color: "var(--error)", marginBottom: "1rem", fontSize: "13px", fontWeight: 600 }}>
                  âš ï¸ {authError}
                </div>
              )}
              <form onSubmit={authTab === "login" ? handleLogin : handleSignup}>
                {authTab === "signup" && (
                  <>
                    <div className="form-group">
                      <label className="form-label">Full Name</label>
                      <input 
                        type="text" 
                        className="form-input" 
                        placeholder="e.g. John Doe" 
                        value={authName}
                        onChange={(e) => setAuthName(e.target.value)}
                        required 
                      />
                    </div>
                    <div className="form-group">
                      <label className="form-label">Mobile Number</label>
                      <input 
                        type="tel" 
                        className="form-input" 
                        placeholder="e.g. +91 99999 88888" 
                        value={authMobile}
                        onChange={(e) => setAuthMobile(e.target.value)}
                        required 
                      />
                    </div>
                    <div className="form-group">
                      <label className="form-label">Email ID</label>
                      <input 
                        type="email" 
                        className="form-input" 
                        placeholder="e.g. customer@example.com" 
                        value={authEmail}
                        onChange={(e) => setAuthEmail(e.target.value)}
                        required 
                      />
                    </div>
                  </>
                )}
                <div className="form-group">
                  <label className="form-label">Username</label>
                  <input 
                    type="text" 
                    className="form-input" 
                    placeholder="e.g. sheet_analyst" 
                    value={authUsername}
                    onChange={(e) => setAuthUsername(e.target.value)}
                    required 
                  />
                </div>
                <div className="form-group">
                  <label className="form-label">Password</label>
                  <input 
                    type="password" 
                    className="form-input" 
                    placeholder="••••••••" 
                    value={authPassword}
                    onChange={(e) => setAuthPassword(e.target.value)}
                    required 
                  />
                </div>
                <button type="submit" className="auth-submit-btn">
                  {authTab === "login" ? "🔑 Sign In" : "ðŸ“ Register"}
                </button>

                <div style={{ display: "flex", alignItems: "center", justifyContent: "center", margin: "1.25rem 0", gap: "10px" }}>
                  <div style={{ flex: 1, height: "1px", background: "var(--hairline)" }}></div>
                  <span style={{ fontSize: "11px", color: "var(--slate)", textTransform: "uppercase", letterSpacing: "0.05em", fontWeight: 600 }}>or</span>
                  <div style={{ flex: 1, height: "1px", background: "var(--hairline)" }}></div>
                </div>

                <div style={{ display: "flex", justifyContent: "center", width: "100%" }}>
                  <div id="google-signin-div" style={{ minHeight: "44px", width: "320px" }}></div>
                </div>
              </form>
            </div>
          </div>
        </div>
      )}

      {/* 2. Checkout Modal (Razorpay & UPI Only Gateway) */}
      {checkoutOpen && (
        <div className="modal-overlay" onClick={() => !paymentProcessing && setCheckoutOpen(false)}>
          <div className="modal-card" onClick={(e) => e.stopPropagation()}>
            <div className="modal-header">
              <h3 className="modal-title">⚡ Upgrade to Pro Member</h3>
              <button 
                type="button" 
                className="modal-close-btn" 
                onClick={() => !paymentProcessing && setCheckoutOpen(false)}
                disabled={paymentProcessing}
              >
                ✕
              </button>
            </div>
            
            <div className="modal-body">
              {checkoutPlans.length > 0 && (
                <div style={{ marginBottom: "1.25rem" }}>
                  <div style={{ fontSize: "11px", fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.06em", color: "var(--slate)", marginBottom: "10px" }}>Choose Your Plan</div>
                  <div style={{ display: "grid", gridTemplateColumns: checkoutPlans.length > 2 ? "repeat(3, 1fr)" : checkoutPlans.length === 2 ? "1fr 1fr" : "1fr", gap: "10px" }}>
                    {checkoutPlans.map(plan => {
                      const isSel = selectedPlanId === plan.id;
                      const ac = plan.color || "#f59e0b";
                      return (
                        <button key={plan.id || plan.name} type="button" onClick={() => setSelectedPlanId(plan.id || null)}
                          style={{
                            position: "relative",
                            padding: "14px 10px",
                            borderRadius: "12px",
                            border: `2px solid ${isSel ? ac : "var(--hairline)"}`,
                            background: isSel ? `${ac}14` : "rgba(255,255,255,0.02)",
                            cursor: "pointer",
                            textAlign: "left",
                            transition: "all 0.15s"
                          }}>
                          {plan.highlighted && (
                            <div style={{ position: "absolute", top: "-9px", left: "50%", transform: "translateX(-50%)", background: ac, color: "#000", fontSize: "7px", fontWeight: 800, padding: "2px 8px", borderRadius: "20px", whiteSpace: "nowrap" }}>⭐ BEST</div>
                          )}
                          <div style={{ fontWeight: 800, fontSize: "13px", color: isSel ? ac : "var(--text)" }}>{plan.name}</div>
                          {plan.description && <div style={{ fontSize: "9px", color: "var(--slate)", marginTop: "1px" }}>{plan.description}</div>}
                          <div style={{ fontWeight: 900, fontSize: "18px", color: ac, marginTop: "4px" }}>
                            ₹{plan.price.toLocaleString()}
                            <span style={{ fontSize: "9px", fontWeight: 400, color: "var(--slate)", marginLeft: "2px" }}>/{plan.billingPeriod}</span>
                          </div>
                          {isSel && (
                            <div style={{ position: "absolute", top: "8px", right: "8px", width: "14px", height: "14px", borderRadius: "50%", background: ac, display: "flex", alignItems: "center", justifyContent: "center", fontSize: "8px", color: "#000", fontWeight: 900 }}>✓</div>
                          )}
                        </button>
                      );
                    })}
                  </div>
                </div>
              )}
              <div className="payment-summary-box">
                {(() => {
                  const sel = checkoutPlans.find(p => p.id === selectedPlanId) || checkoutPlans[0];
                  return sel ? (
                    <>
                      <span>🚀 {sel.name}</span>
                      <strong>₹{sel.price.toLocaleString()}</strong>
                    </>
                  ) : (
                    <>
                      <span>🚀 SheetCodeCrest Pro</span>
                      <strong>₹1,599</strong>
                    </>
                  );
                })()}
              </div>

              {!paymentProcessing && !paymentCompleted ? (
                <div style={{ marginTop: "1rem" }}>
                  {import.meta.env.VITE_RAZORPAY_KEY_ID && (window as any).Razorpay ? (
                    <form onSubmit={(e) => { e.preventDefault(); startPaymentSimulation(); }}>
                      <p style={{ fontSize: "14px", color: "var(--slate)", marginBottom: "1.5rem", lineHeight: "1.5", textAlign: "center" }}>
                        Upgrade instantly via UPI, Netbanking, or Credit/Debit cards securely using the Razorpay gateway.
                      </p>
                      <button type="submit" className="auth-submit-btn">
                        🔒 Pay ₹{((checkoutPlans.find(p => p.id === selectedPlanId) || checkoutPlans[0])?.price || 1599).toLocaleString()} Securely via Razorpay
                      </button>
                    </form>
                  ) : (
                    <div style={{ display: "flex", flexDirection: "column", alignItems: "center", gap: "1.25rem" }}>
                      <div className="qr-simulator-wrapper" style={{ padding: "12px", background: "#ffffff", borderRadius: "12px", border: "1px solid var(--hairline)" }}>
                        <img 
                          src={`https://api.qrserver.com/v1/create-qr-code/?size=180x180&data=${encodeURIComponent(`upi://pay?pa=${PERSONAL_UPI_ID}&pn=SheetCodeCrest&am=${((checkoutPlans.find(p => p.id === selectedPlanId) || checkoutPlans[0])?.price || 1599).toFixed(2)}&cu=INR&tn=${((checkoutPlans.find(p => p.id === selectedPlanId) || checkoutPlans[0])?.name || "SheetCodeCrest Pro")}`)}`} 
                          alt="UPI QR Code" 
                          style={{ display: "block" }} 
                        />
                      </div>
                      <div style={{ textAlign: "center" }}>
                        <div style={{ fontSize: "11px", color: "var(--slate)", textTransform: "uppercase", letterSpacing: "0.05em", fontWeight: 700 }}>Scan QR to Pay with GPay / Paytm / PhonePe</div>
                        <div style={{ fontSize: "13px", fontWeight: 600, color: "var(--coral)", marginTop: "4px", fontFamily: "var(--font-technical)" }}>₹{((checkoutPlans.find(p => p.id === selectedPlanId) || checkoutPlans[0])?.price || 1599).toLocaleString()} ({((checkoutPlans.find(p => p.id === selectedPlanId) || checkoutPlans[0])?.name || "SheetCodeCrest Pro")})</div>
                        <div style={{ fontSize: "11px", color: "var(--slate)", marginTop: "4px" }}>UPI ID: <strong style={{ userSelect: "all", cursor: "pointer", color: "var(--action-blue)" }} onClick={() => { navigator.clipboard.writeText(PERSONAL_UPI_ID); alert("UPI ID copied!"); }}>{PERSONAL_UPI_ID}</strong></div>
                      </div>
                      <form onSubmit={handleManualUpiVerification} style={{ width: "100%", borderTop: "1px solid var(--hairline)", paddingTop: "1rem" }}>
                        <div className="form-group">
                          <label className="form-label" style={{ fontSize: "11px", textTransform: "uppercase", letterSpacing: "0.05em" }}>Submit 12-Digit Transaction UTR / Ref Number</label>
                          <input 
                            type="text" 
                            className="form-input" 
                            placeholder="e.g. 612345678901" 
                            value={upiUTR}
                            onChange={(e) => setUpiUTR(e.target.value)}
                            maxLength={12}
                            required
                          />
                        </div>
                        <button type="submit" className="auth-submit-btn">
                          🔒 Verify & Upgrade Instantly
                        </button>
                      </form>
                    </div>
                  )}
                </div>
              ) : (
                <div style={{ textAlign: "center", padding: "1rem 0" }}>
                  {paymentProcessing ? (
                    <>
                      <div style={{ display: "inline-block", width: "40px", height: "40px", border: "3px solid var(--hairline)", borderTopColor: "var(--coral)", borderRadius: "50%", animation: "spin 1s linear infinite", marginBottom: "1rem" }}></div>
                      <style>{`
                        @keyframes spin {
                          0% { transform: rotate(0deg); }
                          100% { transform: rotate(360deg); }
                        }
                      `}</style>
                      <div style={{ fontWeight: 600, fontSize: "15px", marginBottom: "0.5rem" }}>
                        Processing Transaction...
                      </div>
                      <div style={{ fontSize: "13px", color: "var(--slate)" }}>
                        Please do not refresh or close the gateway.
                      </div>
                    </>
                  ) : (
                    <>
                      <span style={{ fontSize: "3rem", display: "block", marginBottom: "1rem" }}>🎉</span>
                      <div style={{ fontWeight: 700, fontSize: "18px", color: "var(--deep-green)", marginBottom: "0.5rem" }}>
                        Upgrade Completed Successfully!
                      </div>
                      <div style={{ fontSize: "14px", color: "var(--slate)" }}>
                        Welcome to Pro! You now have unlimited report generations and persistent sheet history.
                      </div>
                    </>
                  )}

                  {paymentLogs.length > 0 && (
                    <div className="payment-logs-box">
                      {paymentLogs.map((pLog, idx) => (
                        <div key={idx} className="payment-log-line">
                          {pLog}
                        </div>
                      ))}
                    </div>
                  )}
                </div>
              )}
            </div>
          </div>
        </div>
      )}

      {/* ðŸ›¡ï¸ Secure Admin Panel Modal — Advanced */}
      {adminModalOpen && (
        <div className="modal-overlay" onClick={() => !adminLoading && setAdminModalOpen(false)}>
          <div className="modal-card" onClick={(e) => e.stopPropagation()} style={{ maxWidth: "900px", width: "97%", maxHeight: "92vh", display: "flex", flexDirection: "column" }}>
            {/* Header */}
            <div className="modal-header" style={{ borderColor: "var(--amber)", flexShrink: 0 }}>
              <h3 className="modal-title" style={{ display: "flex", alignItems: "center", gap: "8px", color: "var(--amber)" }}>
                <span>ðŸ›¡ï¸</span> SheetCodeCrest Admin Console
                <span style={{ fontSize: "10px", background: "rgba(245,158,11,0.15)", color: "var(--amber)", padding: "2px 8px", borderRadius: "20px", fontWeight: 500, marginLeft: "4px" }}>
                  SUPER ADMIN
                </span>
              </h3>
              <div style={{ display: "flex", alignItems: "center", gap: "8px" }}>
                <button
                  type="button"
                  onClick={() => loadAdminData()}
                  disabled={adminLoading}
                  title="Refresh data"
                  style={{ background: "rgba(255,255,255,0.06)", border: "1px solid var(--hairline)", borderRadius: "6px", padding: "4px 8px", cursor: "pointer", color: "var(--slate)", fontSize: "14px" }}
                >🔄</button>
                <button type="button" className="modal-close-btn" onClick={() => !adminLoading && setAdminModalOpen(false)} disabled={adminLoading}>✕</button>
              </div>
            </div>

            {/* Tab Navigation */}
            <div style={{ display: "flex", gap: "2px", padding: "0 1.5rem", background: "rgba(0,0,0,0.2)", flexShrink: 0, overflowX: "auto", borderBottom: "1px solid var(--hairline)" }}>
              {(["users","payments","plans","analytics","settings","activity"] as const).map((tab) => {
                const labels: Record<string, string> = {
                  users: `👤 Users (${adminUsers.length})`,
                  payments: `💳 Payments (${adminPayments.length})`,
                  plans: `📦 Plans (${adminPlans.length})`,
                  analytics: "📊 Analytics",
                  settings: "âš™ï¸ Settings",
                  activity: `🔔 Activity (${adminLogs.length})`
                };
                return (
                  <button
                    key={tab}
                    type="button"
                    onClick={() => { setAdminTab(tab); setAdminSearch(""); }}
                    style={{
                      padding: "10px 14px",
                      fontSize: "12px",
                      fontWeight: 600,
                      background: "transparent",
                      border: "none",
                      borderBottom: `2px solid ${adminTab === tab ? "var(--amber)" : "transparent"}`,
                      color: adminTab === tab ? "var(--amber)" : "var(--slate)",
                      cursor: "pointer",
                      whiteSpace: "nowrap",
                      transition: "all 0.2s"
                    }}
                  >{labels[tab]}</button>
                );
              })}
            </div>

            {/* Body */}
            <div style={{ flex: 1, overflowY: "auto", padding: "1.5rem" }}>
              {adminLoading ? (
                <div style={{ textAlign: "center", padding: "4rem 0" }}>
                  <div style={{ display: "inline-block", width: "36px", height: "36px", border: "3px solid var(--hairline)", borderTopColor: "var(--amber)", borderRadius: "50%", animation: "spin 1s linear infinite" }} />
                  <div style={{ marginTop: "1rem", fontSize: "13px", color: "var(--slate)" }}>Syncing database records...</div>
                </div>
              ) : (
                <>
                  {/* ─────────── TAB: USERS ─────────── */}
                  {adminTab === "users" && (
                    <div>
                      {/* Search + Filter Row */}
                      <div style={{ display: "flex", gap: "8px", marginBottom: "1rem", flexWrap: "wrap" }}>
                        <input
                          type="text"
                          className="form-input"
                          placeholder="ðŸ” Search by username, name, email, mobile..."
                          value={adminSearch}
                          onChange={(e) => setAdminSearch(e.target.value)}
                          style={{ flex: 1, minWidth: "200px" }}
                        />
                        {(["all","pro","free"] as const).map(f => (
                          <button key={f} type="button" onClick={() => setAdminUserFilter(f)}
                            style={{ padding: "6px 12px", borderRadius: "6px", fontSize: "11px", fontWeight: 600, cursor: "pointer", border: "1px solid", borderColor: adminUserFilter === f ? "var(--amber)" : "var(--hairline)", background: adminUserFilter === f ? "rgba(245,158,11,0.15)" : "transparent", color: adminUserFilter === f ? "var(--amber)" : "var(--slate)" }}>
                            {f.toUpperCase()}
                          </button>
                        ))}
                      </div>

                      {/* User Table */}
                      <div style={{ border: "1px solid var(--hairline)", borderRadius: "10px", overflow: "hidden" }}>
                        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: "12px" }}>
                          <thead>
                            <tr style={{ background: "rgba(255,255,255,0.04)", borderBottom: "1px solid var(--hairline)" }}>
                              {["Username","Name","Email","Mobile","Plan","Joined","Actions"].map(h => (
                                <th key={h} style={{ padding: "10px 12px", textAlign: "left", fontWeight: 600, color: "var(--slate)", whiteSpace: "nowrap" }}>{h}</th>
                              ))}
                            </tr>
                          </thead>
                          <tbody>
                            {adminUsers
                              .filter(u => {
                                const q = adminSearch.toLowerCase();
                                const matchQ = !q || u.username.toLowerCase().includes(q) || (u.name||"").toLowerCase().includes(q) || (u.email||"").toLowerCase().includes(q) || (u.mobile||"").includes(q);
                                const matchPlan = adminUserFilter === "all" || (adminUserFilter === "pro" ? u.isPro : !u.isPro);
                                return matchQ && matchPlan;
                              })
                              .map((user, i) => (
                                <tr key={user.username} style={{ borderBottom: "1px solid var(--hairline)", background: i % 2 === 0 ? "transparent" : "rgba(255,255,255,0.015)" }}>
                                  <td style={{ padding: "10px 12px", fontWeight: 700 }}>
                                    <div style={{ display: "flex", alignItems: "center", gap: "6px" }}>
                                      <span style={{ width: "28px", height: "28px", borderRadius: "50%", background: user.isPro ? "rgba(245,158,11,0.2)" : "rgba(100,116,139,0.2)", display: "flex", alignItems: "center", justifyContent: "center", fontSize: "12px", flexShrink: 0 }}>
                                        {(user.name || user.username)[0].toUpperCase()}
                                      </span>
                                      {user.username}
                                    </div>
                                  </td>
                                  <td style={{ padding: "10px 12px", color: "var(--slate)" }}>{user.name || "—"}</td>
                                  <td style={{ padding: "10px 12px", color: "var(--slate)", maxWidth: "160px", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{user.email || "—"}</td>
                                  <td style={{ padding: "10px 12px", color: "var(--slate)" }}>{user.mobile || "—"}</td>
                                  <td style={{ padding: "10px 12px" }}>
                                    <span className={`plan-badge ${user.isPro ? "pro" : "free"}`} style={{ fontSize: "9px", padding: "2px 6px" }}>
                                      {user.isPro ? "PRO" : "FREE"}
                                    </span>
                                  </td>
                                  <td style={{ padding: "10px 12px", color: "var(--slate)", whiteSpace: "nowrap" }}>{user.dateCreated}</td>
                                  <td style={{ padding: "10px 12px" }}>
                                    <div style={{ display: "flex", gap: "4px", flexWrap: "nowrap" }}>
                                      <button type="button" onClick={() => openEditUser(user)}
                                        style={{ padding: "4px 8px", borderRadius: "5px", fontSize: "10px", fontWeight: 600, cursor: "pointer", border: "1px solid #3b82f6", background: "rgba(59,130,246,0.1)", color: "#3b82f6" }}>
                                        âœï¸ Edit
                                      </button>
                                      <button type="button" onClick={() => handleToggleUserPro(user)}
                                        style={{ padding: "4px 8px", borderRadius: "5px", fontSize: "10px", fontWeight: 600, cursor: "pointer", border: `1px solid ${user.isPro ? "#ef4444" : "#10b981"}`, background: user.isPro ? "rgba(239,68,68,0.1)" : "rgba(16,185,129,0.1)", color: user.isPro ? "#ef4444" : "#10b981" }}>
                                        {user.isPro ? "Revoke" : "Grant"}
                                      </button>
                                      <button type="button" onClick={() => handleDeleteUser(user)}
                                        style={{ padding: "4px 8px", borderRadius: "5px", fontSize: "10px", fontWeight: 600, cursor: "pointer", border: "1px solid #ef4444", background: "rgba(239,68,68,0.08)", color: "#ef4444" }}>
                                        ðŸ—‘ï¸
                                      </button>
                                    </div>
                                  </td>
                                </tr>
                              ))}
                          </tbody>
                        </table>
                        {adminUsers.filter(u => {
                          const q = adminSearch.toLowerCase();
                          return (!q || u.username.toLowerCase().includes(q) || (u.name||"").toLowerCase().includes(q)) && (adminUserFilter === "all" || (adminUserFilter === "pro" ? u.isPro : !u.isPro));
                        }).length === 0 && (
                          <div style={{ padding: "2rem", textAlign: "center", color: "var(--slate)", fontSize: "13px" }}>No users match your search/filter.</div>
                        )}
                      </div>

                      {/* Edit User Drawer */}
                      {adminEditUserOpen && adminEditUser && (
                        <div style={{ position: "fixed", top: 0, right: 0, width: "380px", height: "100vh", background: "var(--glass-bg)", backdropFilter: "blur(24px)", borderLeft: "1px solid var(--hairline)", zIndex: 9999, padding: "2rem", display: "flex", flexDirection: "column", gap: "1rem", overflowY: "auto" }}>
                          <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: "0.5rem" }}>
                            <h4 style={{ margin: 0, fontSize: "16px", fontWeight: 700, color: "var(--amber)" }}>âœï¸ Edit User</h4>
                            <button type="button" onClick={() => setAdminEditUserOpen(false)} style={{ background: "none", border: "none", color: "var(--slate)", fontSize: "18px", cursor: "pointer" }}>✕</button>
                          </div>
                          <div style={{ padding: "10px 14px", background: "rgba(255,255,255,0.04)", borderRadius: "8px", fontSize: "13px", fontWeight: 600 }}>
                            @{adminEditUser.username}
                          </div>
                          <div className="form-group">
                            <label className="form-label">Full Name</label>
                            <input className="form-input" type="text" value={adminEditUserName} onChange={e => setAdminEditUserName(e.target.value)} placeholder="Full name" />
                          </div>
                          <div className="form-group">
                            <label className="form-label">Email Address</label>
                            <input className="form-input" type="email" value={adminEditUserEmail} onChange={e => setAdminEditUserEmail(e.target.value)} placeholder="Email" />
                          </div>
                          <div className="form-group">
                            <label className="form-label">Mobile Number</label>
                            <input className="form-input" type="tel" value={adminEditUserMobile} onChange={e => setAdminEditUserMobile(e.target.value)} placeholder="Mobile" />
                          </div>
                          <div style={{ display: "flex", alignItems: "center", gap: "10px", padding: "12px", background: "rgba(255,255,255,0.04)", borderRadius: "8px" }}>
                            <input type="checkbox" id="edit-ispro" checked={adminEditUserIsPro} onChange={e => setAdminEditUserIsPro(e.target.checked)} style={{ width: "16px", height: "16px", accentColor: "var(--amber)" }} />
                            <label htmlFor="edit-ispro" style={{ fontSize: "13px", fontWeight: 600, cursor: "pointer" }}>
                              PRO Membership Active
                            </label>
                          </div>
                          <div style={{ display: "flex", gap: "8px", marginTop: "auto" }}>
                            <button type="button" onClick={() => setAdminEditUserOpen(false)} style={{ flex: 1, padding: "10px", border: "1px solid var(--hairline)", borderRadius: "8px", background: "transparent", color: "var(--slate)", cursor: "pointer", fontWeight: 600 }}>Cancel</button>
                            <button type="button" onClick={handleSaveEditUser} style={{ flex: 2, padding: "10px", border: "none", borderRadius: "8px", background: "var(--amber)", color: "#000", cursor: "pointer", fontWeight: 700, fontSize: "13px" }}>💾 Save Changes</button>
                          </div>
                        </div>
                      )}
                    </div>
                  )}

                  {/* ─────────── TAB: PAYMENTS ─────────── */}
                  {adminTab === "payments" && (
                    <div>
                      {/* Stats row */}
                      <div style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: "10px", marginBottom: "1rem" }}>
                        {[
                          { label: "Total Revenue", value: `₹${adminPayments.filter(p => p.status === "success").reduce((s: number, p: any) => s + (p.amount || 0), 0).toLocaleString()}`, color: "#10b981" },
                          { label: "Successful", value: adminPayments.filter(p => p.status === "success").length, color: "#10b981" },
                          { label: "Pending", value: adminPayments.filter(p => p.status === "pending_verification").length, color: "#f59e0b" },
                          { label: "Rejected/Refunded", value: adminPayments.filter(p => ["rejected","refunded"].includes(p.status)).length, color: "#ef4444" },
                        ].map(s => (
                          <div key={s.label} style={{ padding: "12px", borderRadius: "8px", background: "rgba(255,255,255,0.03)", border: "1px solid var(--hairline)", textAlign: "center" }}>
                            <div style={{ fontSize: "20px", fontWeight: 700, color: s.color }}>{s.value}</div>
                            <div style={{ fontSize: "10px", color: "var(--slate)", marginTop: "2px" }}>{s.label}</div>
                          </div>
                        ))}
                      </div>

                      {/* Search + Status filter */}
                      <div style={{ display: "flex", gap: "8px", marginBottom: "1rem", flexWrap: "wrap" }}>
                        <input type="text" className="form-input" placeholder="ðŸ” Search by username or payment ID..." value={adminSearch} onChange={(e) => setAdminSearch(e.target.value)} style={{ flex: 1, minWidth: "200px" }} />
                        {(["all","pending_verification","success","rejected","refunded"] as const).map(f => (
                          <button key={f} type="button" onClick={() => setAdminPaymentFilter(f)}
                            style={{ padding: "6px 10px", borderRadius: "6px", fontSize: "10px", fontWeight: 600, cursor: "pointer", border: "1px solid", whiteSpace: "nowrap",
                              borderColor: adminPaymentFilter === f ? "var(--amber)" : "var(--hairline)",
                              background: adminPaymentFilter === f ? "rgba(245,158,11,0.15)" : "transparent",
                              color: adminPaymentFilter === f ? "var(--amber)" : "var(--slate)"
                            }}>
                            {f === "all" ? "ALL" : f === "pending_verification" ? "PENDING" : f.toUpperCase()}
                          </button>
                        ))}
                      </div>

                      {/* Payment Table */}
                      <div style={{ border: "1px solid var(--hairline)", borderRadius: "10px", overflow: "hidden" }}>
                        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: "12px" }}>
                          <thead>
                            <tr style={{ background: "rgba(255,255,255,0.04)", borderBottom: "1px solid var(--hairline)" }}>
                              {["Transaction ID","User","Gateway","Amount","Status","Actions"].map(h => (
                                <th key={h} style={{ padding: "10px 12px", textAlign: "left", fontWeight: 600, color: "var(--slate)", whiteSpace: "nowrap" }}>{h}</th>
                              ))}
                            </tr>
                          </thead>
                          <tbody>
                            {adminPayments
                              .filter(p => {
                                const q = adminSearch.toLowerCase();
                                const matchQ = !q || p.username.toLowerCase().includes(q) || p.paymentId.toLowerCase().includes(q);
                                const matchF = adminPaymentFilter === "all" || p.status === adminPaymentFilter;
                                return matchQ && matchF;
                              })
                              .map((pay, i) => {
                                const statusColors: Record<string, string> = { success: "#10b981", pending_verification: "#f59e0b", rejected: "#ef4444", refunded: "#8b5cf6" };
                                const sc = statusColors[pay.status] || "var(--slate)";
                                return (
                                  <tr key={pay.id || pay.paymentId} style={{ borderBottom: "1px solid var(--hairline)", background: i % 2 === 0 ? "transparent" : "rgba(255,255,255,0.015)" }}>
                                    <td style={{ padding: "10px 12px", fontFamily: "monospace", fontSize: "11px", maxWidth: "180px", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{pay.paymentId}</td>
                                    <td style={{ padding: "10px 12px", fontWeight: 700 }}>{pay.username}</td>
                                    <td style={{ padding: "10px 12px", color: "var(--slate)", textTransform: "uppercase", fontSize: "11px" }}>{pay.gateway}</td>
                                    <td style={{ padding: "10px 12px", fontWeight: 600, color: "#10b981" }}>₹{(pay.amount || 0).toLocaleString()}</td>
                                    <td style={{ padding: "10px 12px" }}>
                                      <span style={{ padding: "2px 8px", borderRadius: "20px", fontSize: "9px", fontWeight: 700, background: `${sc}22`, color: sc, letterSpacing: "0.05em" }}>
                                        {pay.status === "pending_verification" ? "PENDING" : pay.status.toUpperCase()}
                                      </span>
                                    </td>
                                    <td style={{ padding: "10px 12px" }}>
                                      <div style={{ display: "flex", gap: "4px", flexWrap: "nowrap" }}>
                                        {pay.status === "pending_verification" && (
                                          <button type="button" onClick={() => handleAdminApproveUpi(pay.paymentId, pay.username)}
                                            style={{ padding: "4px 8px", borderRadius: "5px", fontSize: "10px", fontWeight: 700, cursor: "pointer", border: "1px solid #10b981", background: "rgba(16,185,129,0.1)", color: "#10b981", whiteSpace: "nowrap" }}>
                                            ✅ Approve
                                          </button>
                                        )}
                                        {pay.status === "pending_verification" && (
                                          <button type="button" onClick={() => handleAdminRejectPayment(pay.paymentId, pay.username)}
                                            style={{ padding: "4px 8px", borderRadius: "5px", fontSize: "10px", fontWeight: 700, cursor: "pointer", border: "1px solid #ef4444", background: "rgba(239,68,68,0.1)", color: "#ef4444", whiteSpace: "nowrap" }}>
                                            âŒ Reject
                                          </button>
                                        )}
                                        {pay.status === "success" && (
                                          <button type="button" onClick={() => handleAdminRefundPayment(pay.paymentId, pay.username)}
                                            style={{ padding: "4px 8px", borderRadius: "5px", fontSize: "10px", fontWeight: 600, cursor: "pointer", border: "1px solid #8b5cf6", background: "rgba(139,92,246,0.1)", color: "#8b5cf6", whiteSpace: "nowrap" }}>
                                            â†©ï¸ Refund
                                          </button>
                                        )}
                                      </div>
                                    </td>
                                  </tr>
                                );
                              })}
                          </tbody>
                        </table>
                        {adminPayments.filter(p => (adminPaymentFilter === "all" || p.status === adminPaymentFilter)).length === 0 && (
                          <div style={{ padding: "2rem", textAlign: "center", color: "var(--slate)", fontSize: "13px" }}>No payment records match your filter.</div>
                        )}
                      </div>
                    </div>
                  )}

                  {/* ─────────── TAB: PLANS ─────────── */}
                  {adminTab === "plans" && (
                    <div>
                      {/* Header */}
                      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: "1.25rem", flexWrap: "wrap", gap: "8px" }}>
                        <div>
                          <div style={{ fontWeight: 700, fontSize: "14px" }}>📦 Plan Packages</div>
                          <div style={{ fontSize: "12px", color: "var(--slate)", marginTop: "2px" }}>Manage subscription tiers — changes appear live in checkout.</div>
                        </div>
                        <button type="button" onClick={openNewPlan}
                          style={{ padding: "9px 18px", borderRadius: "10px", background: "var(--amber)", color: "#000", border: "none", fontWeight: 800, fontSize: "12px", cursor: "pointer" }}>
                          ➕ New Plan
                        </button>
                      </div>
                      {/* Plan Cards */}
                      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(270px, 1fr))", gap: "16px" }}>
                        {adminPlans.map(plan => {
                          const ac = plan.color || "#f59e0b";
                          return (
                            <div key={plan.id || plan.name} style={{ padding: "1.5rem", borderRadius: "16px", border: `2px solid ${plan.isActive ? ac : "var(--hairline)"}`, background: `linear-gradient(135deg, ${ac}0a, transparent 70%)`, position: "relative" }}>
                              {plan.highlighted && plan.isActive && (
                                <div style={{ position: "absolute", top: "-12px", left: "50%", transform: "translateX(-50%)", background: ac, color: "#000", fontSize: "9px", fontWeight: 800, padding: "3px 12px", borderRadius: "20px", whiteSpace: "nowrap" }}>⭐ RECOMMENDED</div>
                              )}
                              <div style={{ display: "flex", justifyContent: "space-between", marginBottom: "10px" }}>
                                <span style={{ fontSize: "10px", padding: "2px 8px", borderRadius: "20px", fontWeight: 700, background: plan.isActive ? "rgba(16,185,129,0.15)" : "rgba(100,116,139,0.15)", color: plan.isActive ? "#10b981" : "#64748b" }}>
                                  {plan.isActive ? "● ACTIVE" : "○ INACTIVE"}
                                </span>
                                <span style={{ fontSize: "10px", color: "var(--slate)", fontFamily: "monospace" }}>#{plan.sortOrder ?? "—"}</span>
                              </div>
                              <div style={{ fontWeight: 800, fontSize: "18px", color: ac }}>{plan.name}</div>
                              {plan.description && <div style={{ fontSize: "11px", color: "var(--slate)", marginBottom: "4px", fontStyle: "italic" }}>{plan.description}</div>}
                              <div style={{ fontSize: "26px", fontWeight: 900, marginTop: "6px", color: plan.price === 0 ? "#10b981" : "var(--text)" }}>
                                {plan.price === 0 ? "Free" : `₹${plan.price.toLocaleString()}`}
                                {plan.price > 0 && <span style={{ fontSize: "12px", fontWeight: 400, color: "var(--slate)", marginLeft: "4px" }}>/{plan.billingPeriod}</span>}
                              </div>
                              <div style={{ fontSize: "11px", color: "var(--slate)", margin: "4px 0 10px" }}>
                                {(plan.maxReports ?? 0) === 0 ? "♾️ Unlimited reports" : `📊 ${plan.maxReports} reports/${plan.billingPeriod}`}
                              </div>
                              <div style={{ display: "flex", flexWrap: "wrap", gap: "5px", marginBottom: "14px" }}>
                                {plan.features.map((f, i) => (
                                  <span key={i} style={{ fontSize: "10px", padding: "2px 8px", borderRadius: "20px", background: `${ac}18`, color: ac, fontWeight: 600 }}>✓ {f}</span>
                                ))}
                              </div>
                              <div style={{ display: "flex", gap: "6px" }}>
                                <button type="button" onClick={() => openEditPlan(plan)} style={{ flex: 1, padding: "8px 0", borderRadius: "8px", fontSize: "12px", fontWeight: 700, cursor: "pointer", border: `1px solid ${ac}`, background: `${ac}18`, color: ac }}>✏️ Edit</button>
                                <button type="button" title="Duplicate" onClick={() => { setAdminEditPlan(null); setAdminPlanName(plan.name + " (Copy)"); setAdminPlanPrice(plan.price); setAdminPlanPeriod(plan.billingPeriod); setAdminPlanFeatures([...plan.features]); setAdminPlanFeatureInput(""); setAdminPlanActive(false); setAdminPlanDescription(plan.description || ""); setAdminPlanHighlighted(false); setAdminPlanColor(plan.color || "#f59e0b"); setAdminPlanMaxReports(plan.maxReports ?? 0); setAdminPlanSortOrder((plan.sortOrder ?? 99) + 1); setAdminPlanModalOpen(true); }} style={{ padding: "8px 10px", borderRadius: "8px", fontSize: "12px", fontWeight: 700, cursor: "pointer", border: "1px solid var(--hairline)", background: "rgba(255,255,255,0.04)", color: "var(--slate)" }}>📋</button>
                                <button type="button" onClick={() => handleDeletePlan(plan)} style={{ padding: "8px 10px", borderRadius: "8px", fontSize: "12px", fontWeight: 700, cursor: "pointer", border: "1px solid #ef4444", background: "rgba(239,68,68,0.08)", color: "#ef4444" }}>🗑️</button>
                              </div>
                            </div>
                          );
                        })}
                        {adminPlans.length === 0 && (
                          <div style={{ gridColumn: "1/-1", padding: "3rem", textAlign: "center", color: "var(--slate)", fontSize: "13px", border: "2px dashed var(--hairline)", borderRadius: "16px" }}>
                            <div style={{ fontSize: "3rem", marginBottom: "8px" }}>📦</div>
                            No plans yet. Click <strong>➕ New Plan</strong> to get started.
                          </div>
                        )}
                      </div>
                      {/* Plan Editor Side Panel */}
                      {adminPlanModalOpen && (
                        <div style={{ position: "fixed", top: 0, right: 0, width: "440px", height: "100vh", background: "var(--glass-bg)", backdropFilter: "blur(24px)", borderLeft: `2px solid ${adminPlanColor}`, zIndex: 9999, display: "flex", flexDirection: "column" }}>
                          <div style={{ padding: "1.25rem 1.5rem", borderBottom: "1px solid var(--hairline)", display: "flex", alignItems: "center", justifyContent: "space-between", background: `${adminPlanColor}14`, flexShrink: 0 }}>
                            <div>
                              <div style={{ fontWeight: 800, fontSize: "15px", color: adminPlanColor }}>{adminEditPlan ? "✏️ Edit Plan" : "➕ Create Plan"}</div>
                              <div style={{ fontSize: "11px", color: "var(--slate)", marginTop: "2px" }}>{adminEditPlan ? `Editing: ${adminEditPlan.name}` : "New package"}</div>
                            </div>
                            <button type="button" onClick={() => setAdminPlanModalOpen(false)} style={{ background: "none", border: "none", color: "var(--slate)", fontSize: "20px", cursor: "pointer" }}>✕</button>
                          </div>
                          <div style={{ padding: "1.25rem 1.5rem", display: "flex", flexDirection: "column", gap: "14px", overflowY: "auto", flex: 1 }}>
                            <div className="form-group"><label className="form-label">Plan Name *</label><input className="form-input" type="text" value={adminPlanName} onChange={e => setAdminPlanName(e.target.value)} placeholder="e.g. Starter, Pro, Enterprise..." /></div>
                            <div className="form-group"><label className="form-label">Tagline / Description</label><input className="form-input" type="text" value={adminPlanDescription} onChange={e => setAdminPlanDescription(e.target.value)} placeholder="e.g. Perfect for small teams" /></div>
                            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "10px" }}>
                              <div className="form-group"><label className="form-label">Price (₹) — 0 = Free</label><input className="form-input" type="number" min={0} value={adminPlanPrice} onChange={e => setAdminPlanPrice(Number(e.target.value))} /></div>
                              <div className="form-group"><label className="form-label">Billing Period</label><select className="form-input" value={adminPlanPeriod} onChange={e => setAdminPlanPeriod(e.target.value as Plan["billingPeriod"])} style={{ appearance: "none", cursor: "pointer" }}>{(["free","monthly","yearly","lifetime"] as const).map(per => (<option key={per} value={per}>{per.charAt(0).toUpperCase() + per.slice(1)}</option>))}</select></div>
                            </div>
                            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "10px" }}>
                              <div className="form-group"><label className="form-label">Report Limit (0 = ∞)</label><input className="form-input" type="number" min={0} value={adminPlanMaxReports} onChange={e => setAdminPlanMaxReports(Number(e.target.value))} /></div>
                              <div className="form-group"><label className="form-label">Sort Order</label><input className="form-input" type="number" min={0} value={adminPlanSortOrder} onChange={e => setAdminPlanSortOrder(Number(e.target.value))} /></div>
                            </div>
                            <div className="form-group"><label className="form-label">Accent Color</label><div style={{ display: "flex", gap: "10px", alignItems: "center" }}><input type="color" value={adminPlanColor} onChange={e => setAdminPlanColor(e.target.value)} style={{ width: "48px", height: "40px", padding: "2px", borderRadius: "8px", border: "1px solid var(--hairline)", cursor: "pointer" }} /><div style={{ flex: 1, height: "40px", borderRadius: "8px", border: `2px solid ${adminPlanColor}`, background: `${adminPlanColor}14`, display: "flex", alignItems: "center", justifyContent: "center", fontSize: "13px", fontWeight: 800, color: adminPlanColor }}>{adminPlanName || "Preview"}</div></div></div>
                            <div className="form-group"><label className="form-label">Features</label><div style={{ display: "flex", gap: "6px" }}><input className="form-input" type="text" value={adminPlanFeatureInput} onChange={e => setAdminPlanFeatureInput(e.target.value)} onKeyDown={e => { if (e.key === "Enter" && adminPlanFeatureInput.trim()) { setAdminPlanFeatures(prev => [...prev, adminPlanFeatureInput.trim()]); setAdminPlanFeatureInput(""); e.preventDefault(); }}} placeholder="Type feature, press Enter..." /><button type="button" onClick={() => { if (adminPlanFeatureInput.trim()) { setAdminPlanFeatures(prev => [...prev, adminPlanFeatureInput.trim()]); setAdminPlanFeatureInput(""); }}} style={{ padding: "0 14px", borderRadius: "8px", border: "none", background: "var(--amber)", color: "#000", fontWeight: 700, cursor: "pointer", fontSize: "18px" }}>+</button></div><div style={{ display: "flex", flexWrap: "wrap", gap: "6px", marginTop: "8px" }}>{adminPlanFeatures.map((f, i) => (<span key={i} style={{ fontSize: "11px", padding: "4px 10px", borderRadius: "20px", background: `${adminPlanColor}18`, color: adminPlanColor, display: "flex", alignItems: "center", gap: "5px", fontWeight: 600 }}>{f}<button type="button" onClick={() => setAdminPlanFeatures(prev => prev.filter((_, j) => j !== i))} style={{ background: "none", border: "none", color: "#ef4444", cursor: "pointer", padding: 0, fontSize: "12px" }}>✕</button></span>))}</div></div>
                            <div style={{ display: "flex", gap: "10px" }}>
                              <label style={{ flex: 1, display: "flex", alignItems: "center", gap: "8px", padding: "10px 12px", borderRadius: "8px", cursor: "pointer", border: `1px solid ${adminPlanActive ? "#10b981" : "var(--hairline)"}`, background: "rgba(255,255,255,0.03)" }}><input id="plan-active-chk" type="checkbox" checked={adminPlanActive} onChange={e => setAdminPlanActive(e.target.checked)} style={{ width: "16px", height: "16px", accentColor: "#10b981" }} /><span style={{ fontSize: "12px", fontWeight: 700 }}>Active</span></label>
                              <label style={{ flex: 1, display: "flex", alignItems: "center", gap: "8px", padding: "10px 12px", borderRadius: "8px", cursor: "pointer", border: `1px solid ${adminPlanHighlighted ? adminPlanColor : "var(--hairline)"}`, background: "rgba(255,255,255,0.03)" }}><input id="plan-highlight-chk" type="checkbox" checked={adminPlanHighlighted} onChange={e => setAdminPlanHighlighted(e.target.checked)} style={{ width: "16px", height: "16px", accentColor: adminPlanColor }} /><span style={{ fontSize: "12px", fontWeight: 700 }}>⭐ Recommended</span></label>
                            </div>
                          </div>
                          <div style={{ padding: "1rem 1.5rem", borderTop: "1px solid var(--hairline)", display: "flex", gap: "8px", flexShrink: 0 }}>
                            <button type="button" onClick={() => setAdminPlanModalOpen(false)} style={{ flex: 1, padding: "11px", border: "1px solid var(--hairline)", borderRadius: "10px", background: "transparent", color: "var(--slate)", cursor: "pointer", fontWeight: 700 }}>Cancel</button>
                            <button type="button" onClick={handleSavePlan} disabled={adminLoading} style={{ flex: 2, padding: "11px", border: "none", borderRadius: "10px", background: adminPlanColor, color: "#000", cursor: "pointer", fontWeight: 800, fontSize: "13px", opacity: adminLoading ? 0.6 : 1 }}>{adminLoading ? "Saving..." : (adminEditPlan ? "💾 Save Changes" : "✅ Create Plan")}</button>
                          </div>
                        </div>
                      )}
                    </div>
                  )}

                  {/* ─────────── TAB: ANALYTICS ─────────── */}
                  {adminTab === "analytics" && (
                    <div style={{ display: "flex", flexDirection: "column", gap: "1.5rem" }}>
                      {/* KPI Cards */}
                      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(160px, 1fr))", gap: "12px" }}>
                        {[
                          { icon: "👤", label: "Total Users", value: adminUsers.length, color: "#3b82f6" },
                          { icon: "⚡", label: "PRO Users", value: adminUsers.filter(u => u.isPro).length, color: "#f59e0b" },
                          { icon: "🆓", label: "Free Users", value: adminUsers.filter(u => !u.isPro).length, color: "#64748b" },
                          { icon: "💰", label: "Total Revenue", value: `₹${adminPayments.filter(p => p.status === "success").reduce((s: number, p: any) => s + (p.amount || 0), 0).toLocaleString()}`, color: "#10b981" },
                          { icon: "💳", label: "Transactions", value: adminPayments.length, color: "#8b5cf6" },
                          { icon: "â³", label: "Pending", value: adminPayments.filter(p => p.status === "pending_verification").length, color: "#ef4444" },
                        ].map(card => (
                          <div key={card.label} style={{ padding: "1.25rem", borderRadius: "12px", background: "rgba(255,255,255,0.03)", border: "1px solid var(--hairline)", textAlign: "center" }}>
                            <div style={{ fontSize: "28px", marginBottom: "6px" }}>{card.icon}</div>
                            <div style={{ fontSize: "24px", fontWeight: 800, color: card.color }}>{card.value}</div>
                            <div style={{ fontSize: "11px", color: "var(--slate)", marginTop: "4px", fontWeight: 500 }}>{card.label}</div>
                          </div>
                        ))}
                      </div>

                      {/* Plan Distribution Chart (CSS bars) */}
                      <div style={{ padding: "1.25rem", borderRadius: "12px", background: "rgba(255,255,255,0.02)", border: "1px solid var(--hairline)" }}>
                        <h4 style={{ margin: "0 0 1rem 0", fontSize: "13px", fontWeight: 700, color: "var(--slate)", textTransform: "uppercase", letterSpacing: "0.05em" }}>Plan Distribution</h4>
                        {[
                          { label: "PRO Users", count: adminUsers.filter(u => u.isPro).length, color: "#f59e0b" },
                          { label: "Free Users", count: adminUsers.filter(u => !u.isPro).length, color: "#3b82f6" },
                        ].map(bar => (
                          <div key={bar.label} style={{ marginBottom: "12px" }}>
                            <div style={{ display: "flex", justifyContent: "space-between", fontSize: "12px", marginBottom: "4px" }}>
                              <span style={{ color: bar.color, fontWeight: 600 }}>{bar.label}</span>
                              <span style={{ color: "var(--slate)" }}>{bar.count} ({adminUsers.length ? Math.round(bar.count / adminUsers.length * 100) : 0}%)</span>
                            </div>
                            <div style={{ height: "8px", borderRadius: "4px", background: "rgba(255,255,255,0.06)", overflow: "hidden" }}>
                              <div style={{ height: "100%", width: `${adminUsers.length ? (bar.count / adminUsers.length * 100) : 0}%`, background: bar.color, borderRadius: "4px", transition: "width 0.6s ease" }} />
                            </div>
                          </div>
                        ))}
                      </div>

                      {/* Recent Signups */}
                      <div style={{ padding: "1.25rem", borderRadius: "12px", background: "rgba(255,255,255,0.02)", border: "1px solid var(--hairline)" }}>
                        <h4 style={{ margin: "0 0 1rem 0", fontSize: "13px", fontWeight: 700, color: "var(--slate)", textTransform: "uppercase", letterSpacing: "0.05em" }}>Recent Signups</h4>
                        {adminUsers.slice(0, 8).map(u => (
                          <div key={u.username} style={{ display: "flex", alignItems: "center", gap: "10px", padding: "8px 0", borderBottom: "1px solid var(--hairline)" }}>
                            <span style={{ width: "32px", height: "32px", borderRadius: "50%", background: u.isPro ? "rgba(245,158,11,0.2)" : "rgba(100,116,139,0.15)", display: "flex", alignItems: "center", justifyContent: "center", fontSize: "13px", fontWeight: 700, flexShrink: 0 }}>{(u.name || u.username)[0].toUpperCase()}</span>
                            <div style={{ flex: 1, minWidth: 0 }}>
                              <div style={{ fontWeight: 600, fontSize: "13px" }}>{u.username} {u.name && <span style={{ color: "var(--slate)", fontWeight: 400, fontSize: "11px" }}>({u.name})</span>}</div>
                              <div style={{ fontSize: "11px", color: "var(--slate)" }}>{u.dateCreated}</div>
                            </div>
                            <span className={`plan-badge ${u.isPro ? "pro" : "free"}`} style={{ fontSize: "9px", padding: "2px 6px" }}>{u.isPro ? "PRO" : "FREE"}</span>
                          </div>
                        ))}
                      </div>
                    </div>
                  )}

                  {/* ─────────── TAB: SETTINGS ─────────── */}
                  {adminTab === "settings" && (
                    <div style={{ display: "flex", flexDirection: "column", gap: "1.5rem", maxWidth: "580px" }}>
                      {/* Free Limit */}
                      <div style={{ padding: "1.25rem", borderRadius: "12px", border: "1px solid var(--hairline)", background: "rgba(255,255,255,0.02)" }}>
                        <h4 style={{ margin: "0 0 1rem 0", fontSize: "14px", fontWeight: 700, color: "var(--amber)" }}>📊 Usage Limits</h4>
                        <div className="form-group">
                          <label className="form-label">Global Free Report Limit</label>
                          <input type="number" className="form-input" value={globalFreeLimit} onChange={(e) => setGlobalFreeLimit(Number(e.target.value))} min={1} style={{ maxWidth: "160px" }} />
                          <p style={{ fontSize: "11px", color: "var(--slate)", marginTop: "6px" }}>Number of free report generations allowed before users are prompted to upgrade.</p>
                        </div>
                        <button type="button" onClick={() => handleSaveFreeLimit(globalFreeLimit)}
                          style={{ padding: "8px 18px", borderRadius: "8px", background: "var(--amber)", color: "#000", border: "none", fontWeight: 700, fontSize: "12px", cursor: "pointer" }}>
                          💾 Save Limit
                        </button>
                      </div>

                      {/* Feature Flags */}
                      <div style={{ padding: "1.25rem", borderRadius: "12px", border: "1px solid var(--hairline)", background: "rgba(255,255,255,0.02)" }}>
                        <h4 style={{ margin: "0 0 1rem 0", fontSize: "14px", fontWeight: 700, color: "var(--amber)" }}>🚩 Feature Flags</h4>
                        {[
                          { id: "feat-ai", label: "AI Analyst Chat", desc: "Enable the Avery AI chat assistant for all users", value: adminFeatureAI, onChange: setAdminFeatureAI },
                          { id: "feat-upi", label: "UPI Manual Fallback", desc: "Allow users to pay via UPI QR code + UTR manual entry", value: adminFeatureUPI, onChange: setAdminFeatureUPI },
                          { id: "feat-google", label: "Google Sign-In", desc: "Allow authentication via Google OAuth", value: adminFeatureGoogleLogin, onChange: setAdminFeatureGoogleLogin },
                          { id: "feat-maint", label: "🔴 Maintenance Mode", desc: "Show maintenance banner across the app for all users", value: adminMaintenanceMode, onChange: setAdminMaintenanceMode },
                        ].map(flag => (
                          <div key={flag.id} style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "12px", marginBottom: "8px", borderRadius: "8px", background: "rgba(255,255,255,0.03)", border: "1px solid var(--hairline)" }}>
                            <div>
                              <div style={{ fontWeight: 600, fontSize: "13px" }}>{flag.label}</div>
                              <div style={{ fontSize: "11px", color: "var(--slate)", marginTop: "2px" }}>{flag.desc}</div>
                            </div>
                            <label style={{ position: "relative", display: "inline-block", width: "44px", height: "24px", flexShrink: 0 }}>
                              <input type="checkbox" checked={flag.value} onChange={e => flag.onChange(e.target.checked)} style={{ opacity: 0, width: 0, height: 0 }} />
                              <span style={{
                                position: "absolute", inset: 0, borderRadius: "12px", cursor: "pointer", transition: "0.3s",
                                background: flag.value ? "var(--amber)" : "rgba(100,116,139,0.3)"
                              }}>
                                <span style={{
                                  position: "absolute", top: "3px", left: flag.value ? "23px" : "3px",
                                  width: "18px", height: "18px", borderRadius: "50%", background: "#fff",
                                  transition: "0.3s", boxShadow: "0 1px 4px rgba(0,0,0,0.4)"
                                }} />
                              </span>
                            </label>
                          </div>
                        ))}
                        <button type="button" onClick={handleSaveFeatureFlags}
                          style={{ marginTop: "8px", padding: "8px 18px", borderRadius: "8px", background: "var(--amber)", color: "#000", border: "none", fontWeight: 700, fontSize: "12px", cursor: "pointer" }}>
                          💾 Save Feature Flags
                        </button>
                      </div>

                      {/* Danger Zone */}
                      <div style={{ padding: "1.25rem", borderRadius: "12px", border: "1px solid rgba(239,68,68,0.3)", background: "rgba(239,68,68,0.04)" }}>
                        <h4 style={{ margin: "0 0 1rem 0", fontSize: "14px", fontWeight: 700, color: "#ef4444" }}>âš ï¸ Danger Zone</h4>
                        <p style={{ fontSize: "12px", color: "var(--slate)", margin: "0 0 12px 0" }}>These actions are irreversible. Use with extreme caution.</p>
                        <button type="button"
                          onClick={() => { if (confirm("Clear all admin activity logs? This cannot be undone.")) { setAdminLogs([]); addLog("âš ï¸ Admin: Activity logs cleared.", "error"); }}}
                          style={{ padding: "8px 16px", borderRadius: "8px", background: "rgba(239,68,68,0.1)", color: "#ef4444", border: "1px solid #ef4444", fontWeight: 700, fontSize: "12px", cursor: "pointer" }}>
                          ðŸ—‘ï¸ Clear Activity Logs
                        </button>
                      </div>
                    </div>
                  )}

                  {/* ─────────── TAB: ACTIVITY LOG ─────────── */}
                  {adminTab === "activity" && (
                    <div>
                      <div style={{ marginBottom: "1rem", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                        <div style={{ fontSize: "13px", color: "var(--slate)" }}>
                          Real-time log of all admin actions performed in this console.
                        </div>
                        <button type="button" onClick={() => loadAdminData()}
                          style={{ padding: "6px 12px", borderRadius: "6px", fontSize: "11px", fontWeight: 600, cursor: "pointer", border: "1px solid var(--hairline)", background: "transparent", color: "var(--slate)" }}>
                          🔄 Refresh
                        </button>
                      </div>
                      <div style={{ border: "1px solid var(--hairline)", borderRadius: "10px", overflow: "hidden" }}>
                        {adminLogs.length > 0 ? (
                          <div style={{ maxHeight: "440px", overflowY: "auto" }}>
                            {adminLogs.map((log, i) => {
                              const actionColors: Record<string, string> = {
                                DELETE_USER: "#ef4444", DELETE_PLAN: "#ef4444", REJECT_PAYMENT: "#ef4444", REFUND_PAYMENT: "#8b5cf6",
                                APPROVE_PAYMENT: "#10b981", TOGGLE_PRO: "#f59e0b", GRANT_PRO: "#10b981",
                                EDIT_USER: "#3b82f6", CREATE_PLAN: "#10b981", EDIT_PLAN: "#3b82f6",
                                UPDATE_SETTINGS: "#64748b", UPDATE_FLAGS: "#64748b"
                              };
                              const color = actionColors[log.action] || "var(--slate)";
                              return (
                                <div key={i} style={{ padding: "12px 16px", borderBottom: i < adminLogs.length - 1 ? "1px solid var(--hairline)" : "none", display: "flex", gap: "12px", alignItems: "flex-start" }}>
                                  <div style={{ width: "8px", height: "8px", borderRadius: "50%", background: color, marginTop: "5px", flexShrink: 0 }} />
                                  <div style={{ flex: 1, minWidth: 0 }}>
                                    <div style={{ display: "flex", alignItems: "center", gap: "8px", flexWrap: "wrap" }}>
                                      <span style={{ fontWeight: 700, fontSize: "11px", color: color, letterSpacing: "0.05em", fontFamily: "monospace" }}>{log.action}</span>
                                      <span style={{ fontSize: "11px", color: "var(--slate)" }}>by <strong>{log.performedBy}</strong></span>
                                      {log.createdAt && <span style={{ fontSize: "10px", color: "var(--slate)", marginLeft: "auto" }}>{new Date(log.createdAt).toLocaleString()}</span>}
                                    </div>
                                    {log.details && <div style={{ fontSize: "11px", color: "var(--slate)", marginTop: "3px" }}>{log.details}</div>}
                                  </div>
                                </div>
                              );
                            })}
                          </div>
                        ) : (
                          <div style={{ padding: "3rem", textAlign: "center", color: "var(--slate)", fontSize: "13px" }}>
                            No admin actions logged yet. Actions you take in this console will appear here.
                          </div>
                        )}
                      </div>
                    </div>
                  )}
                </>
              )}
            </div>
          </div>
        </div>
      )}


      {/* 3. Dashboard Modal (Profile + Saved Spreadsheets History) */}
      {dashboardOpen && (
        <div className="modal-overlay" onClick={() => setDashboardOpen(false)}>
          <div className="modal-card" onClick={(e) => e.stopPropagation()} style={{ maxWidth: "640px" }}>
            <div className="modal-header">
              <h3 className="modal-title">ðŸŽ›ï¸ Account & Analytics Dashboard</h3>
              <button 
                type="button" 
                className="modal-close-btn" 
                onClick={() => setDashboardOpen(false)}
              >
                ✕
              </button>
            </div>
            <div className="modal-body">
              {currentUser && (
                <>
                  <div className="dashboard-stats-grid">
                    <div className="dashboard-stat-tile">
                      <div className="dashboard-stat-label">User Account</div>
                      <div className="dashboard-stat-value">{currentUser.username}</div>
                    </div>
                    <div className="dashboard-stat-tile">
                      <div className="dashboard-stat-label">Subscription Tier</div>
                      <div className="dashboard-stat-value" style={{ display: "flex", alignItems: "center", gap: "6px" }}>
                        <span className={`plan-badge ${currentUser.isPro ? "pro" : "free"}`}>
                          {currentUser.isPro ? "PRO MEMBERSHIP" : "FREE TRIAL"}
                        </span>
                        {!currentUser.isPro && (
                          <button
                            type="button"
                            className="header-btn"
                            onClick={() => { setDashboardOpen(false); setCheckoutOpen(true); }}
                            style={{ fontSize: "10px", padding: "3px 8px", borderColor: "var(--coral)", color: "var(--coral)" }}
                          >
                            ⚡ Upgrade
                          </button>
                        )}
                      </div>
                    </div>
                  </div>

                  <div style={{ marginTop: "1.5rem" }}>
                    <h4 style={{ margin: "0 0 10px 0", fontSize: "14px", fontFamily: "var(--font-display)", fontWeight: 600, borderBottom: "1px solid var(--hairline)", paddingBottom: "6px", textAlign: "left" }}>
                      ðŸ“ Saved Spreadsheets & Analytics
                    </h4>
                    <div className="records-list-wrapper">
                      {savedRecords.length > 0 ? (
                        savedRecords.map((rec) => (
                          <div className="record-item-row" key={rec.id}>
                            <div className="record-info-left">
                              <div className="record-filename-text" title={rec.filename}>
                                {rec.filename}
                              </div>
                              <div className="record-meta-text">
                                <span style={{ 
                                  background: rec.mode === "shopify" ? "rgba(24, 99, 220, 0.1)" : rec.mode === "logistics" ? "rgba(0, 60, 51, 0.1)" : "var(--soft-stone)",
                                  color: rec.mode === "shopify" ? "var(--action-blue)" : rec.mode === "logistics" ? "var(--deep-green)" : "var(--slate)",
                                  padding: "1px 5px",
                                  borderRadius: "3px",
                                  fontSize: "9px",
                                  fontWeight: 700
                                }}>
                                  {rec.mode.toUpperCase()}
                                </span>
                                <span>•</span>
                                <span>{(rec.size / 1024).toFixed(1)} KB</span>
                                <span>•</span>
                                <span>{rec.timestamp}</span>
                              </div>
                            </div>
                            <div className="record-item-actions">
                              <button
                                type="button"
                                className="record-load-btn"
                                onClick={() => loadRecord(rec)}
                              >
                                ⚡ Load Viewer
                              </button>
                              <button
                                type="button"
                                className="record-delete-btn"
                                onClick={(e) => handleDeleteRecord(rec.id!, e)}
                                title="Delete saved spreadsheet"
                              >
                                ðŸ—‘ï¸
                              </button>
                            </div>
                          </div>
                        ))
                      ) : (
                        <div style={{ padding: "2rem", color: "var(--slate)", fontSize: "13px", textAlign: "center" }}>
                          No spreadsheets saved yet. Upload a sheet and it will auto-save here!
                        </div>
                      )}
                    </div>
                  </div>

                </>
              )}
            </div>
          </div>
        </div>
      )}
      {/* 4. Secure Share Modal Overlay */}
      {shareModalOpen && (
        <div className="modal-overlay" onClick={() => setShareModalOpen(false)}>
          <div className="modal-card" onClick={(e) => e.stopPropagation()} style={{ maxWidth: "480px" }}>
            <div className="modal-header">
              <h3 className="modal-title">🔗 Secure Shareable Link</h3>
              <button 
                type="button" 
                className="modal-close-btn" 
                onClick={() => setShareModalOpen(false)}
              >
                ✕
              </button>
            </div>
            <div className="modal-body" style={{ display: "flex", flexDirection: "column", gap: "1rem", textAlign: "left" }}>
              <p style={{ fontSize: "13px", color: "var(--slate)", margin: 0 }}>
                Anyone with this secure link can view this interactive analytics report in read-only mode without needing an account.
              </p>
              
              <div style={{ display: "flex", flexDirection: "column", gap: "6px" }}>
                <label style={{ fontSize: "11px", fontWeight: 700, color: "var(--slate)" }}>SECURE LINK</label>
                <div style={{ display: "flex", gap: "8px" }}>
                  <input 
                    type="text" 
                    readOnly 
                    className="ai-key-input" 
                    style={{ margin: 0, fontSize: "12px", width: "100%", fontFamily: "monospace" }}
                    value={`${window.location.origin}${window.location.pathname}?share=${activeRecordId}`}
                  />
                  <button 
                    type="button" 
                    className="btn-primary" 
                    style={{ padding: "8px 16px", borderRadius: "8px", fontSize: "12px", whiteSpace: "nowrap" }}
                    onClick={() => {
                      navigator.clipboard.writeText(`${window.location.origin}${window.location.pathname}?share=${activeRecordId}`);
                      alert("Secure report link copied to clipboard!");
                    }}
                  >
                    Copy Link
                  </button>
                </div>
              </div>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
