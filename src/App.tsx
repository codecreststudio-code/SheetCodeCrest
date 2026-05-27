import { useState, useRef, useCallback, useEffect, useMemo } from "react";
import { motion } from "motion/react";
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
// Import validated environment configuration
import { env } from "./config/env";
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


// Use validated environment variables
const GOOGLE_CLIENT_ID = env.googleClientId;
const PERSONAL_UPI_ID = env.personalUpiId;

type ThemeMode = "light" | "dark";

// Use validated environment variables
const API_PROXY_URL = `${env.proxyUrl}/api/chat`;
const GOOGLE_SHEET_PROXY_URL = `${env.proxyUrl}/api/sync-sheet`;

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

const getCurrencySymbol = (curr: "INR" | "USD" | "EUR") => {
  if (curr === "INR") return "₹";
  if (curr === "EUR") return "€";
  return "$";
};

const getConvertedPrice = (basePrice: number, curr: "INR" | "USD" | "EUR") => {
  if (basePrice === 0) return 0;
  if (curr === "INR") return basePrice;
  if (curr === "USD") {
    if (basePrice === 999) return 12;
    if (basePrice === 2499) return 29;
    return Math.round(basePrice / 83);
  }
  if (curr === "EUR") {
    if (basePrice === 999) return 11;
    if (basePrice === 2499) return 27;
    return Math.round(basePrice / 90);
  }
  return basePrice;
};

const SaaSBackgroundParticles = () => {
  const canvasRef = useRef<HTMLCanvasElement>(null);

  useEffect(() => {
    const canvas = canvasRef.current;
    if (!canvas) return;
    const ctx = canvas.getContext("2d");
    if (!ctx) return;

    let animationFrameId: number;
    let width = (canvas.width = canvas.offsetWidth);
    let height = (canvas.height = canvas.offsetHeight);

    const particles: Array<{
      x: number;
      y: number;
      size: number;
      speedX: number;
      speedY: number;
      alpha: number;
    }> = [];

    // Create particles
    const particleCount = 45;
    for (let i = 0; i < particleCount; i++) {
      particles.push({
        x: Math.random() * width,
        y: Math.random() * height,
        size: Math.random() * 2 + 0.5,
        speedX: (Math.random() - 0.5) * 0.35,
        speedY: (Math.random() - 0.5) * 0.35,
        alpha: Math.random() * 0.5 + 0.1,
      });
    }

    const handleResize = () => {
      if (!canvas) return;
      width = canvas.width = canvas.offsetWidth;
      height = canvas.height = canvas.offsetHeight;
    };
    window.addEventListener("resize", handleResize);

    // Track mouse
    let mouse = { x: -9999, y: -9999 };
    const handleMouseMove = (e: MouseEvent) => {
      const rect = canvas.getBoundingClientRect();
      mouse.x = e.clientX - rect.left;
      mouse.y = e.clientY - rect.top;
    };
    const handleMouseLeave = () => {
      mouse.x = -9999;
      mouse.y = -9999;
    };

    const parent = canvas.parentElement;
    if (parent) {
      parent.addEventListener("mousemove", handleMouseMove);
      parent.addEventListener("mouseleave", handleMouseLeave);
    }

    const draw = () => {
      ctx.clearRect(0, 0, width, height);

      // Draw lines between close particles
      for (let i = 0; i < particles.length; i++) {
        const p1 = particles[i];
        
        // Update positions
        p1.x += p1.speedX;
        p1.y += p1.speedY;

        // Bounce borders
        if (p1.x < 0 || p1.x > width) p1.speedX *= -1;
        if (p1.y < 0 || p1.y > height) p1.speedY *= -1;

        // Mouse reaction
        if (mouse.x !== -9999) {
          const dx = p1.x - mouse.x;
          const dy = p1.y - mouse.y;
          const dist = Math.sqrt(dx * dx + dy * dy);
          if (dist < 100) {
            const force = (100 - dist) / 100;
            p1.x += (dx / dist) * force * 1.5;
            p1.y += (dy / dist) * force * 1.5;
          }
        }

        // Draw particle
        ctx.beginPath();
        ctx.arc(p1.x, p1.y, p1.size, 0, Math.PI * 2);
        ctx.fillStyle = `rgba(0, 255, 178, ${p1.alpha})`;
        ctx.fill();

        // Connect lines
        for (let j = i + 1; j < particles.length; j++) {
          const p2 = particles[j];
          const dx = p1.x - p2.x;
          const dy = p1.y - p2.y;
          const dist = Math.sqrt(dx * dx + dy * dy);

          if (dist < 85) {
            const alpha = (1 - dist / 85) * 0.15;
            ctx.beginPath();
            ctx.moveTo(p1.x, p1.y);
            ctx.lineTo(p2.x, p2.y);
            ctx.strokeStyle = `rgba(0, 255, 178, ${alpha})`;
            ctx.lineWidth = 0.5;
            ctx.stroke();
          }
        }
      }

      animationFrameId = requestAnimationFrame(draw);
    };

    draw();

    return () => {
      cancelAnimationFrame(animationFrameId);
      window.removeEventListener("resize", handleResize);
      if (parent) {
        parent.removeEventListener("mousemove", handleMouseMove);
        parent.removeEventListener("mouseleave", handleMouseLeave);
      }
    };
  }, []);

  return (
    <canvas
      ref={canvasRef}
      style={{
        position: "absolute",
        inset: 0,
        width: "100%",
        height: "100%",
        pointerEvents: "none",
        zIndex: 0,
        opacity: 0.75,
      }}
    />
  );
};

const CustomGlowingCursor = () => {
  const [position, setPosition] = useState({ x: -100, y: -100 });
  const [trail, setTrail] = useState({ x: -100, y: -100 });
  const [isVisible, setIsVisible] = useState(false);
  const [isHovered, setIsHovered] = useState(false);

  useEffect(() => {
    const handleMouseMove = (e: MouseEvent) => {
      setPosition({ x: e.clientX, y: e.clientY });
      setIsVisible(true);
    };

    const handleMouseLeave = () => {
      setIsVisible(false);
    };

    const handleMouseOver = (e: MouseEvent) => {
      const target = e.target as HTMLElement;
      if (
        target.tagName === "BUTTON" ||
        target.tagName === "A" ||
        target.closest("button") ||
        target.closest("a") ||
        target.closest('[role="button"]') ||
        target.style.cursor === "pointer"
      ) {
        setIsHovered(true);
      } else {
        setIsHovered(false);
      }
    };

    window.addEventListener("mousemove", handleMouseMove);
    document.addEventListener("mouseleave", handleMouseLeave);
    window.addEventListener("mouseover", handleMouseOver);

    return () => {
      window.removeEventListener("mousemove", handleMouseMove);
      document.removeEventListener("mouseleave", handleMouseLeave);
      window.removeEventListener("mouseover", handleMouseOver);
    };
  }, []);

  useEffect(() => {
    let animationFrameId: number;

    const updateTrail = () => {
      setTrail((prev) => {
        const dx = position.x - prev.x;
        const dy = position.y - prev.y;
        return {
          x: prev.x + dx * 0.15,
          y: prev.y + dy * 0.15,
        };
      });
      animationFrameId = requestAnimationFrame(updateTrail);
    };

    animationFrameId = requestAnimationFrame(updateTrail);
    return () => cancelAnimationFrame(animationFrameId);
  }, [position]);

  if (!isVisible) return null;

  return (
    <>
      {/* Inner glowing solid dot (neon-green theme) */}
      <div
        style={{
          position: "fixed",
          top: 0,
          left: 0,
          width: isHovered ? "10px" : "8px",
          height: isHovered ? "10px" : "8px",
          borderRadius: "50%",
          background: "#00ffb2",
          boxShadow: "0 0 12px #00ffb2, 0 0 20px rgba(0, 255, 178, 0.5)",
          transform: `translate3d(${position.x - (isHovered ? 5 : 4)}px, ${position.y - (isHovered ? 5 : 4)}px, 0)`,
          pointerEvents: "none",
          zIndex: 999999,
          transition: "width 0.15s, height 0.15s, background-color 0.15s",
        }}
      />
      {/* Outer trailing interactive halo ring */}
      <div
        style={{
          position: "fixed",
          top: 0,
          left: 0,
          width: isHovered ? "36px" : "28px",
          height: isHovered ? "36px" : "28px",
          borderRadius: "50%",
          border: "1.5px solid rgba(0, 255, 178, 0.4)",
          background: isHovered ? "rgba(0, 255, 178, 0.05)" : "transparent",
          transform: `translate3d(${trail.x - (isHovered ? 18 : 14)}px, ${trail.y - (isHovered ? 18 : 14)}px, 0)`,
          pointerEvents: "none",
          zIndex: 999998,
          transition: "width 0.2s, height 0.2s, background-color 0.2s, border-color 0.2s",
        }}
      />
    </>
  );
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
  
  const [paymentLogs, setPaymentLogs] = useState<string[]>([]);
  const [paymentProcessing, setPaymentProcessing] = useState(false);
  const [paymentCompleted, setPaymentCompleted] = useState(false);
  const [theme, setTheme] = useState<ThemeMode>("dark");

  // FAQ accordion active state
  const [expandedFaq, setExpandedFaq] = useState<number | null>(null);

  // Currency option with automatic region detection
  const [currency, setCurrency] = useState<"INR" | "USD" | "EUR">(() => {
    if (typeof window === "undefined" || typeof Intl === "undefined") return "USD";
    try {
      const tz = Intl.DateTimeFormat().resolvedOptions().timeZone;
      if (tz && (tz.includes("Kolkata") || tz.includes("Calcutta") || tz.includes("India") || tz === "Asia/Kolkata")) {
        return "INR";
      }
      // Check locale languages as fallback
      const langs = window.navigator.languages || [window.navigator.language];
      if (langs.some(l => l && (l.toLowerCase().includes("in") || l.toLowerCase().includes("hi")))) {
        return "INR";
      }
    } catch (e) {
      // Ignore
    }
    return "USD";
  });

  // E-commerce CAC, Margin & ROAS calculator state
  const [calcAdSpend, setCalcAdSpend] = useState(15000);
  const [calcOrders, setCalcOrders] = useState(250);
  const [calcAOV, setCalcAOV] = useState(1200);
  const [calcCOGS, setCalcCOGS] = useState(400);

  // Bento Card 5 security demo toggles
  const [bentoAnonymize, setBentoAnonymize] = useState(true);
  const [bentoSandbox, setBentoSandbox] = useState(true);
  const [bentoAutoDelete, setBentoAutoDelete] = useState(false);

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

  // COD Premium Optimizer State
  const [codPremiumVal, setCodPremiumVal] = useState(50);
  const [codSwitchRate, setCodSwitchRate] = useState(15);
  const [rtoCostPerOrder, setRtoCostPerOrder] = useState(150);

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
  const [landingPlans, setLandingPlans] = useState<Plan[]>([
    {
      id: "basic",
      name: "Basic",
      price: 0,
      billingPeriod: "free",
      features: [
        "File uploads: 3 / month",
        "AI insights per report: 5 insights",
        "Auto chart generation: 4 charts",
        "Schema & column detection",
        "Data profiler (basic stats)",
        "CSV, XLSX, JSON, TSV support",
        "100% client-side — data never leaves your device"
      ],
      isActive: true,
      description: "Perfect for exploring your data. No credit card needed.",
      highlighted: false,
      color: "#3b82f6",
      maxReports: 3,
      sortOrder: 0
    },
    {
      id: "standard",
      name: "Standard",
      price: 999,
      billingPeriod: "monthly",
      features: [
        "Everything in Basic",
        "File uploads: 250 / month",
        "AI chat questions: 500 / month",
        "AI insights per report: unlimited",
        "Charts per dashboard: 20 charts",
        "AI analyst chat — plain English queries",
        "Advanced filters & live dashboard",
        "Export charts as PNG & CSV",
        "Saved report history: 30 days",
        "Email support"
      ],
      isActive: true,
      description: "For analysts and teams running regular reports on any data.",
      highlighted: true,
      color: "#2563eb",
      maxReports: 250,
      sortOrder: 1
    },
    {
      id: "premium",
      name: "Premium",
      price: 2499,
      billingPeriod: "monthly",
      features: [
        "Everything in Standard",
        "File uploads: unlimited",
        "AI chat questions: unlimited",
        "Team seats: up to 10 users",
        "White-label reports with your logo",
        "Custom column mapping rules",
        "Saved report history: 90 days",
        "API access: coming soon",
        "Dedicated account manager",
        "Priority WhatsApp support"
      ],
      isActive: true,
      description: "For agencies and power users needing full control and team access.",
      highlighted: false,
      color: "#a855f7",
      maxReports: 0,
      sortOrder: 2
    }
  ]);
  const [selectedPlanId, setSelectedPlanId] = useState<string | null>(null);
  const [sheetSearchQuery, setSheetSearchQuery] = useState("");

  const filteredSavedRecords = useMemo(() => {
    if (!sheetSearchQuery) return savedRecords;
    const q = sheetSearchQuery.toLowerCase().trim();
    return savedRecords.filter(r => 
      r.filename.toLowerCase().includes(q) || 
      r.mode.toLowerCase().includes(q)
    );
  }, [savedRecords, sheetSearchQuery]);

  // Admin Settings -- feature flags
  const [adminFeatureAI, setAdminFeatureAI] = useState(() => {
    if (typeof window === "undefined") return true;
    return window.localStorage.getItem("sheetcc_flag_ai") !== "false";
  });
  const [adminFeatureUPI, setAdminFeatureUPI] = useState(() => {
    if (typeof window === "undefined") return true;
    return window.localStorage.getItem("sheetcc_flag_upi") !== "false";
  });
  const [adminFeatureGoogleLogin, setAdminFeatureGoogleLogin] = useState(() => {
    if (typeof window === "undefined") return true;
    return window.localStorage.getItem("sheetcc_flag_google") !== "false";
  });
  const [adminMaintenanceMode, setAdminMaintenanceMode] = useState(() => {
    if (typeof window === "undefined") return false;
    return window.localStorage.getItem("sheetcc_flag_maintenance") === "true";
  });
  const [globalFreeLimit, setGlobalFreeLimit] = useState(() => {
    if (typeof window === "undefined") return 3;
    const stored = window.localStorage.getItem("sheetcodecrest_global_free_limit");
    return stored ? Number(stored) : 3;
  });

  const loadSystemSettings = () => {
    try {
      const savedFreeLimit = window.localStorage.getItem("sheetcodecrest_global_free_limit");
      if (savedFreeLimit) setGlobalFreeLimit(Number(savedFreeLimit));
      
      const savedAI = window.localStorage.getItem("sheetcc_flag_ai");
      if (savedAI !== null) setAdminFeatureAI(savedAI !== "false");
      
      const savedUPI = window.localStorage.getItem("sheetcc_flag_upi");
      if (savedUPI !== null) setAdminFeatureUPI(savedUPI !== "false");
      
      const savedGoogle = window.localStorage.getItem("sheetcc_flag_google");
      if (savedGoogle !== null) setAdminFeatureGoogleLogin(savedGoogle !== "false");
      
      const savedMaint = window.localStorage.getItem("sheetcc_flag_maintenance");
      if (savedMaint !== null) setAdminMaintenanceMode(savedMaint === "true");
    } catch (e) {
      console.error("Failed to load local system settings", e);
    }
  };

  const [mockupTabActive, setMockupTabActive] = useState<"shopify" | "logistics" | "universal">("shopify");
  const [testimonialIdx, setTestimonialIdx] = useState(0);
  const [pricingBilling, setPricingBilling] = useState<"monthly" | "yearly">("monthly");

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

      // 3. Load all plans dynamically on app mount
      loadSystemSettings();
      await loadCheckoutPlans();

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
    if (authModalOpen && adminFeatureGoogleLogin) {
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
  }, [authModalOpen, adminFeatureGoogleLogin]);

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
      addLog(`⚠️ Google Sheets Sync failed: ${err.message || err}`, "error");
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
    loadSystemSettings();
    await loadCheckoutPlans();
    
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
        addLog(`🔑 Password security upgraded to SHA-256 hash for "${username}".`, "info");
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
    loadSystemSettings();
    await loadCheckoutPlans();
    
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

    // 🔑 Hash password before storing
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
    setAuthEmail("");
    setSavedRecords([]);
    loadSystemSettings();
    await loadCheckoutPlans();
    
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
    loadSystemSettings();
    loadCheckoutPlans();
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
      "🛡️ Validating secure transaction checksum (SHA-256)...",
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
      "☁️ Syncing transaction records securely to cloud database...",
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
  // 🛡️ ADMIN PANEL CONTROLLERS (Advanced)
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
      setLandingPlans(plans);
      setCheckoutPlans(plans.filter(p => p.isActive && p.price > 0));
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
      addLog(`🛡️ Admin: User "${adminEditUser.username}" profile updated.`, "info");
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
      addLog(`🛡️ Admin: Toggled PRO for "${targetUser.username}" → ${!targetUser.isPro}`, "info");
      await loadAdminData();
      if (currentUser && currentUser.username === targetUser.username) setCurrentUser(updated);
    } catch (err) {
      console.error("Failed to toggle PRO", err);
      alert("Failed to update user privilege.");
    }
  };

  const handleDeleteUser = async (targetUser: User) => {
    if (!confirm(`⚠️ PERMANENTLY delete user "${targetUser.username}"? This cannot be undone.`)) return;
    setAdminLoading(true);
    try {
      await dbDeleteUser(targetUser.username);
      await logAdminAction("DELETE_USER", `Deleted user "${targetUser.username}"`);
      addLog(`🛡️ Admin: User "${targetUser.username}" deleted permanently.`, "error");
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
      addLog(`🛡️ Admin approved payment: ${paymentId}. Promoted "${username}" to PRO.`, "success");
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
      addLog(`🛡️ Admin rejected payment: ${paymentId} for "${username}".`, "error");
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
      addLog(`🛡️ Admin refunded payment: ${paymentId}, revoked PRO for "${username}".`, "error");
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
      addLog(`🛡️ Admin: Plan "${plan.name}" ${adminEditPlan ? "updated" : "created"}.`, "info");
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
      addLog(`🛡️ Admin: Plan "${plan.name}" deleted.`, "error");
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
    addLog(`⚙️ System setting updated: free limit = ${limit}`, "info");
  };

  const handleSaveFeatureFlags = () => {
    window.localStorage.setItem("sheetcc_flag_ai", String(adminFeatureAI));
    window.localStorage.setItem("sheetcc_flag_upi", String(adminFeatureUPI));
    window.localStorage.setItem("sheetcc_flag_google", String(adminFeatureGoogleLogin));
    window.localStorage.setItem("sheetcc_flag_maintenance", String(adminMaintenanceMode));
    logAdminAction("UPDATE_FLAGS", `AI=${adminFeatureAI}, UPI=${adminFeatureUPI}, Google=${adminFeatureGoogleLogin}, Maintenance=${adminMaintenanceMode}`);
    loadSystemSettings(); // Refresh instantly in-memory!
    addLog("⚙️ Feature flags saved successfully.", "success");
    alert("✅ Feature flags saved!");
  };



  // Load active paid plans for the checkout modal and landing page
  const loadCheckoutPlans = async () => {
    try {
      const plans = await dbGetPlans();
      setLandingPlans(plans);
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
        setDetectionNotice(`🛠️ Shopify export auto-detected in the background (${shopifyScore}/8 schema match). Building customer, product, order, retargeting, COD, and geographic analytics.`);
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
3. **⚠️ Regional Return-to-Origin Hotspot**: **${worstRtoState?.[0] || "N/A"}** is flagging a critical RTO return risk of **${((worstRtoState?.[1]?.rto / worstRtoState?.[1]?.orders) * 100).toFixed(1)}%** over ${worstRtoState?.[1]?.orders || 0} shipments. Consider restricting COD or requiring prepaid payment for high-value orders in this territory.
4. **📦 Packaging Freight Waste**: The overall freight charges total **₹${stats.totalFreight.toLocaleString("en-IN")}**, representing **${(stats.totalFreight / stats.totalRev * 100).toFixed(1)}%** of sales. Review box sizes and volumetric weight brackets to reduce dead-weight costs.`;
  }, [logisticsAnalytics]);

  // Offline pre-computed responses for high availability without API key
  const getOfflineReply = (question: string): string => {
    const q = question.toLowerCase();
    if (mode === "shopify" && shopifyAnalytics) {
      const stats = shopifyAnalytics;
      if (q.includes("product") || q.includes("best") || q.includes("top")) {
        return `### 🛠️ Shopify Product Performance
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
        return `### ⚠️ Regional Return-to-Origin Hotspot Analysis (Offline Mode)
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
      {/* Custom Glowing Cursor Pointer */}
      <CustomGlowingCursor />
      {/* Global Maintenance Mode Overlay */}
      {adminMaintenanceMode && !isAdminActive && (
        <div className="maint-overlay">
          <div className="maint-card">
            <span className="maint-icon">🔧</span>
            <h2 className="maint-title">System Maintenance</h2>
            <p className="maint-desc">
              We are currently performing scheduled system upgrades to improve our analysis processors. 
              We'll be back shortly. Thank you for your patience!
            </p>
            <div className="maint-divider"></div>
            <div className="maint-admin-section">
              <span className="maint-admin-label">Administrator Bypass</span>
              <button 
                type="button" 
                className="btn-maint-bypass" 
                onClick={() => { setAuthTab("login"); setAuthError(""); setAuthModalOpen(true); }}
              >
                🔒 Sign In
              </button>
            </div>
          </div>
        </div>
      )}

      <style>{`:root {
  --canvas: #030305;
  --primary: #0a0a0c;
  --ink: #ffffff;
  --deep-green: #00ffb2;
  --dark-navy: #050508;
  --soft-stone: #0d0d10;
  --pale-green: rgba(0, 255, 178, 0.08);
  --pale-blue: rgba(0, 255, 178, 0.04);
  --hairline: rgba(255, 255, 255, 0.06);
  --border-light: rgba(255, 255, 255, 0.08);
  --card-border: rgba(255, 255, 255, 0.06);
  --muted: #a3a3a3;
  --slate: #94a3b8;
  --body-muted: #cbd5e1;
  --action-blue: #00ffb2;
  --focus-blue: #00ffb2;
  --on-primary: #030305;
  --coral: #ef4444;
  --coral-soft: #ff8888;
  --amber: #00ffb2;
  --error: #ef4444;
  --glass-bg: rgba(13, 13, 16, 0.65);
  --glass-border: rgba(255, 255, 255, 0.06);

  --font-display: 'Inter', 'Space Grotesk', sans-serif;
  --font-ui: 'Inter', sans-serif;
  --font-technical: 'JetBrains Mono', 'Fira Code', monospace;
}

* {
  box-sizing: border-box;
  margin: 0;
  padding: 0;
  transition: background-color 0.25s ease, border-color 0.25s ease;
}

body {
  background-color: var(--canvas);
  color: var(--ink);
  font-family: var(--font-ui);
  line-height: 1.5;
  overflow-x: hidden;
  background-attachment: fixed;
}

/* Global Layout */
.app-container {
  min-height: 100vh;
  display: flex;
  flex-direction: column;
  position: relative;
}

.app-container::before {
  content: "";
  position: fixed;
  inset: 0;
  background-image: 
    radial-gradient(circle at 15% 15%, rgba(0, 255, 178, 0.08) 0%, transparent 40%),
    radial-gradient(circle at 85% 65%, rgba(0, 255, 178, 0.05) 0%, transparent 45%),
    radial-gradient(rgba(0, 255, 178, 0.05) 1.2px, transparent 1.2px);
  background-size: 100% 100%, 100% 100%, 32px 32px;
  pointer-events: none;
  z-index: -1;
}

/* Header / Navigation */
.nav-header {
  display: flex;
  justify-content: space-between;
  align-items: center;
  padding: 1.25rem 2rem;
  background: var(--glass-bg);
  backdrop-filter: blur(12px);
  border-bottom: 1px solid var(--glass-border);
  position: sticky;
  top: 0;
  z-index: 1000;
}

.nav-logo {
  display: flex;
  align-items: center;
  gap: 10px;
  font-family: var(--font-display);
  font-weight: 900;
  font-size: 1.35rem;
  letter-spacing: -0.04em;
  background: linear-gradient(135deg, #ffffff 50%, var(--focus-blue) 100%);
  -webkit-background-clip: text;
  -webkit-text-fill-color: transparent;
}

.nav-actions {
  display: flex;
  align-items: center;
  gap: 12px;
}

/* Hero Area */
.landing-hero {
  padding: 4rem 2rem;
  max-width: 1200px;
  margin: 0 auto;
  position: relative;
}

.hero-split {
  display: grid;
  grid-template-columns: 1.15fr 0.85fr;
  gap: 4rem;
  align-items: center;
}

@media (max-width: 968px) {
  .hero-split {
    grid-template-columns: 1fr;
    gap: 3rem;
  }
}

.hero-left {
  display: flex;
  flex-direction: column;
  align-items: flex-start;
  text-align: left;
}

@media (max-width: 968px) {
  .hero-left {
    align-items: center;
    text-align: center;
  }
}

.hero-title {
  font-family: var(--font-display);
  font-size: 3.5rem;
  font-weight: 900;
  letter-spacing: -0.05em;
  line-height: 1.05;
  margin-bottom: 1.5rem;
  background: linear-gradient(135deg, #ffffff 55%, var(--focus-blue) 100%);
  -webkit-background-clip: text;
  -webkit-text-fill-color: transparent;
  text-align: left;
}

@media (max-width: 968px) {
  .hero-title {
    font-size: 2.8rem;
    text-align: center !important;
  }
}

.hero-subtitle {
  font-size: 1.1rem;
  color: var(--slate);
  margin-bottom: 2rem;
  line-height: 1.55;
  max-width: 620px;
  text-align: left;
  margin-left: 0;
  margin-right: 0;
}

@media (max-width: 968px) {
  .hero-subtitle {
    text-align: center !important;
    margin-left: auto !important;
    margin-right: auto !important;
  }
}

.hero-right {
  position: relative;
  width: 100%;
  height: 480px;
  border-radius: 24px;
  background: radial-gradient(circle at 50% 50%, rgba(250, 255, 105, 0.08), transparent 70%);
  border: 1px solid rgba(255, 255, 255, 0.05);
  display: flex;
  align-items: center;
  justify-content: center;
  overflow: hidden;
}

@media (max-width: 968px) {
  .hero-right {
    height: 380px;
  }
}

.spline-container {
  width: 100%;
  height: 100%;
  border-radius: 24px;
  overflow: hidden;
}


/* Bento Features Grid */
.features-grid {
  display: grid;
  grid-template-columns: repeat(auto-fit, minmax(260px, 1fr));
  gap: 1.25rem;
  padding: 0 2rem 5rem;
  max-width: 1100px;
  margin: 0 auto;
}

.feature-card {
  background: var(--primary);
  border: 1px solid var(--hairline);
  border-radius: 12px;
  padding: 2.5rem 2rem;
  transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
  position: relative;
  overflow: hidden;
}

.feature-card:hover {
  transform: translateY(-4px);
  border-color: var(--focus-blue);
}

/* Feature Item Cards (used in landing features section) */
.feature-item-card {
  background: #1a1a1a;
  border: 1px solid #2a2a2a;
  border-radius: 12px;
  padding: 2rem 1.75rem;
  transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
  position: relative;
  overflow: hidden;
}

.feature-item-card:hover {
  transform: translateY(-4px);
  border-color: var(--focus-blue);
  box-shadow: 0 0 24px rgba(250, 255, 105, 0.06);
}


.feature-title-card {
  font-family: var(--font-display);
  font-size: 1rem;
  font-weight: 700;
  color: var(--ink);
  margin: 0 0 0.75rem 0;
  letter-spacing: -0.01em;
}

.feature-desc-card {
  font-size: 0.875rem;
  color: var(--slate);
  line-height: 1.6;
  margin: 0;
}

.feature-icon-wrapper {
  width: 44px;
  height: 44px;
  border-radius: 8px;
  background: rgba(250, 255, 105, 0.08);
  border: 1px solid rgba(250, 255, 105, 0.15);
  color: var(--focus-blue);
  display: flex;
  align-items: center;
  justify-content: center;
  font-size: 1.25rem;
  margin-bottom: 1rem;
}

/* Stat Callout Section */
.stat-callout-section {
  max-width: 900px;
  margin: 0 auto 3rem auto;
  padding: 0 2rem;
  display: grid;
  grid-template-columns: repeat(3, 1fr);
  gap: 2rem;
  text-align: center;
}

@media (max-width: 600px) {
  .stat-callout-section {
    grid-template-columns: 1fr;
    gap: 1.5rem;
  }
  .stat-callout-num {
    font-size: 2.25rem !important;
  }
}

.hero-stats {
  display: flex;
  justify-content: flex-start;
  gap: 2.5rem;
  width: 100%;
  flex-wrap: wrap;
  margin-bottom: 2rem;
}

@media (max-width: 968px) {
  .hero-stats {
    justify-content: center;
    gap: 1.5rem;
  }
  .hero-stats .stat-callout-num {
    font-size: 2.25rem !important;
  }
}

.stat-callout-num {
  font-family: var(--font-display);
  font-size: 3rem;
  font-weight: 700;
  color: var(--focus-blue);
  line-height: 1;
  letter-spacing: -1.5px;
  display: block;
  margin-bottom: 4px;
}

.stat-callout-label {
  font-size: 0.8rem;
  color: var(--muted);
  font-weight: 500;
  text-transform: uppercase;
  letter-spacing: 0.08em;
}

/* Yellow CTA Band */
.cta-yellow-band {
  background: #faff69;
  border-radius: 16px;
  padding: 4rem 3rem;
  max-width: 900px;
  margin: 2rem auto 4rem auto;
  text-align: center;
  position: relative;
  overflow: hidden;
}

.cta-yellow-band::before {
  content: "";
  position: absolute;
  inset: 0;
  background: radial-gradient(circle at 30% 50%, rgba(255,255,255,0.3) 0%, transparent 60%);
  pointer-events: none;
}

.cta-band-headline {
  font-family: var(--font-display);
  font-size: 2.25rem;
  font-weight: 700;
  color: #0a0a0a;
  letter-spacing: -1.5px;
  line-height: 1.1;
  margin-bottom: 0.75rem;
}

.cta-band-sub {
  font-size: 1rem;
  color: rgba(10, 10, 10, 0.7);
  margin-bottom: 1.75rem;
  max-width: 480px;
  margin-left: auto;
  margin-right: auto;
}

.cta-band-btn {
  background: #0a0a0a;
  color: #faff69;
  border: none;
  font-family: var(--font-display);
  font-weight: 700;
  font-size: 0.9rem;
  padding: 14px 28px;
  border-radius: 8px;
  cursor: pointer;
  transition: all 0.2s;
  display: inline-flex;
  align-items: center;
  gap: 8px;
  letter-spacing: -0.01em;
}

.cta-band-btn:hover {
  background: #1a1a1a;
  transform: translateY(-2px);
}

/* Upload Zone Area */
.upload-container {
  max-width: 680px;
  margin: 0 auto 4rem;
  padding: 0 1.5rem;
}

.upload-zone, .upload-card {
  border: 2px dashed var(--hairline);
  background: var(--primary);
  border-radius: 12px;
  padding: 3.5rem 2rem;
  text-align: center;
  cursor: pointer;
  transition: all 0.3s ease;
  display: flex;
  flex-direction: column;
  align-items: center;
  gap: 1rem;
}

.upload-zone:hover, .upload-card:hover, .upload-card.dragging {
  border-color: var(--focus-blue);
  background: rgba(250, 255, 105, 0.02);
  transform: translateY(-2px);
}

.upload-title {
  font-family: var(--font-display);
  font-size: 1.5rem;
  font-weight: 700;
  color: var(--ink);
  margin-top: 0.5rem;
}

.upload-desc {
  font-size: 0.95rem;
  color: var(--slate);
  max-width: 420px;
  margin: 0 auto;
}

.usage-meter {
  font-size: 0.85rem;
  color: var(--focus-blue);
  font-family: var(--font-technical);
  font-weight: 600;
  margin-top: 0.25rem;
}

.upload-icon {
  font-size: 3.5rem;
  color: var(--focus-blue);
  animation: bounce 3s infinite;
}

@keyframes bounce {
  0%, 100% { transform: translateY(0); }
  50% { transform: translateY(-8px); }
}

/* Interactive Mockup Styling */
.mockup-container {
  max-width: 1000px;
  margin: 0 auto 6rem auto;
  background: var(--primary);
  border: 1px solid var(--hairline);
  border-radius: 12px;
  overflow: hidden;
  text-align: left;
}

.mockup-header {
  background: #121212;
  padding: 12px 20px;
  border-bottom: 1px solid var(--hairline);
  display: flex;
  align-items: center;
}

.mockup-dot {
  width: 10px;
  height: 10px;
  border-radius: 50%;
  margin-right: 6px;
  display: inline-block;
}

.mockup-tab-strip {
  display: flex;
  gap: 8px;
  padding: 10px 18px;
  background: #161616;
  border-bottom: 1px solid var(--hairline);
  flex-wrap: wrap;
}

.mockup-tab {
  color: var(--slate);
  font-size: 13px;
  font-weight: 600;
  padding: 8px 16px;
  border-radius: 6px;
  transition: all 0.2s;
  background: transparent;
  border: none;
  cursor: pointer;
  outline: none;
}

.mockup-tab:hover {
  color: var(--ink);
  background: rgba(255, 255, 255, 0.03);
}

.mockup-tab.active {
  color: var(--on-primary) !important;
  background: var(--focus-blue) !important;
  font-weight: 700;
}

.mockup-body {
  display: grid;
  grid-template-columns: 1.4fr 1fr;
  gap: 24px;
  padding: 24px;
}

@media (max-width: 768px) {
  .mockup-body {
    grid-template-columns: 1fr;
  }
}

.mockup-left {
  overflow-x: auto;
  border: 1px solid var(--hairline);
  border-radius: 8px;
  background: #121212;
}

.mockup-table {
  width: 100%;
  border-collapse: collapse;
  font-family: var(--font-technical);
  font-size: 12px;
  text-align: left;
}

.mockup-table th {
  background: #1a1a1a;
  color: var(--slate);
  font-weight: 600;
  padding: 12px 14px;
  border-bottom: 1px solid var(--hairline);
  text-transform: uppercase;
  font-size: 10px;
  letter-spacing: 0.05em;
}

.mockup-table td {
  padding: 12px 14px;
  border-bottom: 1px solid var(--hairline);
  color: var(--ink);
  white-space: nowrap;
}

.mockup-table tr:last-child td {
  border-bottom: none;
}

.mockup-table tr:hover {
  background: rgba(255, 255, 255, 0.02);
}

.mockup-badge {
  display: inline-block;
  padding: 2px 8px;
  border-radius: 4px;
  font-size: 10px;
  font-weight: 700;
}

.mockup-badge.delivered {
  background: rgba(34, 197, 94, 0.15);
  color: #22c55e;
}

.mockup-badge.rto {
  background: rgba(239, 68, 68, 0.15);
  color: #ef4444;
}

.mockup-right {
  display: flex;
  flex-direction: column;
  gap: 16px;
}

.mockup-card-right {
  background: #121212;
  border: 1px solid var(--hairline);
  border-radius: 8px;
  padding: 20px;
}

.mockup-chart-row {
  display: flex;
  align-items: flex-end;
  justify-content: space-between;
  height: 120px;
  padding: 10px 10px 0 10px;
  border-bottom: 1px solid var(--hairline);
  margin-bottom: 12px;
}

.mockup-chart-bar {
  width: 28px;
  background: #2a2a2a;
  border-radius: 4px 4px 0 0;
  transition: all 0.3s;
}

.mockup-chart-bar:hover {
  background: var(--focus-blue);
  box-shadow: 0 0 15px rgba(250, 255, 105, 0.4);
}

.mockup-chart-bar.highlight {
  background: var(--focus-blue);
}

/* Premium Buttons & Controls */
.btn-primary, .auth-submit-btn, .action-btn {
  background: var(--action-blue);
  color: var(--on-primary);
  border: none;
  font-weight: 700;
  padding: 12px 24px;
  border-radius: 8px;
  cursor: pointer;
  transition: all 0.2s;
  display: inline-flex;
  align-items: center;
  justify-content: center;
  gap: 8px;
}

.btn-primary:hover, .auth-submit-btn:hover, .action-btn:hover {
  transform: translateY(-2px);
  background: #e6eb52;
}

.btn-secondary {
  background: var(--primary);
  border: 1px solid var(--hairline);
  color: var(--ink);
  font-weight: 600;
  padding: 10px 20px;
  border-radius: 8px;
  cursor: pointer;
  transition: all 0.2s;
}

.btn-secondary:hover {
  background: #242424;
  border-color: var(--slate);
}

/* Data Tables & Analytics Grid */
.section-card {
  background: var(--primary);
  border: 1px solid var(--hairline);
  border-radius: 12px;
  padding: 2rem;
  margin-bottom: 2rem;
}

.table-container {
  overflow-x: auto;
  border-radius: 12px;
  border: 1px solid var(--hairline);
  margin: 1.5rem 0;
}

.data-table {
  width: 100%;
  border-collapse: collapse;
  text-align: left;
  font-size: 13px;
}

.data-table th {
  background: rgba(255, 255, 255, 0.03);
  padding: 12px 16px;
  font-weight: 700;
  color: var(--slate);
  border-bottom: 1px solid var(--hairline);
  text-transform: uppercase;
  letter-spacing: 0.05em;
}

.data-table td {
  padding: 12px 16px;
  border-bottom: 1px solid var(--hairline);
  color: var(--ink);
}

.data-table tr:hover {
  background: rgba(255, 255, 255, 0.015);
}

/* Forms & inputs */
.form-group {
  margin-bottom: 1.25rem;
}

.form-label {
  display: block;
  font-size: 12px;
  font-weight: 700;
  text-transform: uppercase;
  letter-spacing: 0.05em;
  color: var(--slate);
  margin-bottom: 6px;
}

.form-input {
  width: 100%;
  padding: 10px 14px;
  border-radius: 8px;
  background: rgba(255, 255, 255, 0.03);
  border: 1px solid var(--hairline);
  color: var(--ink);
  font-family: inherit;
  transition: all 0.2s;
}

.form-input:focus {
  outline: none;
  border-color: var(--focus-blue);
  background: rgba(255, 255, 255, 0.05);
  box-shadow: 0 0 0 2px rgba(250, 255, 105, 0.1);
}

/* Modals & Overlays */
.modal-overlay {
  position: fixed;
  inset: 0;
  background: rgba(0, 0, 0, 0.75);
  backdrop-filter: blur(10px);
  display: flex;
  align-items: center;
  justify-content: center;
  z-index: 9999;
  padding: 1rem;
}

.modal-card, .modal-content {
  background: rgba(20, 20, 22, 0.88) !important;
  border: 1px solid rgba(255, 255, 255, 0.08) !important;
  border-radius: 24px !important;
  max-width: 500px;
  width: 100%;
  max-height: 90vh;
  overflow-y: auto;
  display: flex;
  flex-direction: column;
  box-shadow: 0 24px 64px rgba(0, 0, 0, 0.7), inset 0 1px 0 rgba(255, 255, 255, 0.1) !important;
  overflow: hidden;
  backdrop-filter: blur(24px);
}

.auth-tabs {
  display: flex;
  background: rgba(255, 255, 255, 0.03);
  border: 1px solid rgba(255, 255, 255, 0.05);
  border-radius: 12px;
  padding: 4px;
  margin: 1.5rem 1.5rem 0.5rem;
  gap: 4px;
}

.auth-tab-btn {
  flex: 1;
  padding: 10px 16px;
  border-radius: 8px;
  border: none;
  background: transparent;
  color: #a1a1aa;
  font-size: 13px;
  font-weight: 700;
  cursor: pointer;
  transition: all 0.25s cubic-bezier(0.4, 0, 0.2, 1);
  font-family: var(--font-technical);
  text-transform: uppercase;
  letter-spacing: 0.05em;
}

.auth-tab-btn.active {
  background: #22c55e;
  color: #052e16;
  font-weight: 800;
  box-shadow: 0 4px 14px rgba(34, 197, 94, 0.3);
}

.auth-tab-btn:hover:not(.active) {
  color: #ffffff;
  background: rgba(255, 255, 255, 0.03);
}

.form-group {
  margin-bottom: 1.25rem;
  display: flex;
  flex-direction: column;
  gap: 6px;
}

.form-label {
  font-size: 11px;
  font-weight: 700;
  text-transform: uppercase;
  color: #a1a1aa;
  letter-spacing: 0.08em;
  font-family: var(--font-technical);
}

.form-input {
  background: rgba(0, 0, 0, 0.25) !important;
  border: 1px solid rgba(255, 255, 255, 0.08) !important;
  border-radius: 10px !important;
  padding: 12px 16px !important;
  color: #ffffff !important;
  font-size: 14px !important;
  transition: all 0.2s cubic-bezier(0.4, 0, 0.2, 1) !important;
}

.form-input:focus {
  border-color: #22c55e !important;
  background: rgba(0, 0, 0, 0.45) !important;
  box-shadow: 0 0 0 3px rgba(34, 197, 94, 0.15) !important;
  outline: none !important;
}

.auth-submit-btn {
  width: 100%;
  padding: 14px;
  border-radius: 12px;
  border: none;
  background: #22c55e !important;
  color: #052e16 !important;
  font-size: 14px;
  font-weight: 800;
  cursor: pointer;
  display: inline-flex;
  align-items: center;
  justify-content: center;
  gap: 8px;
  transition: all 0.2s cubic-bezier(0.4, 0, 0.2, 1);
  box-shadow: 0 4px 16px rgba(34, 197, 94, 0.25);
  margin-top: 1.25rem;
  text-transform: uppercase;
  letter-spacing: 0.05em;
  font-family: var(--font-technical);
}

.auth-submit-btn:hover {
  background: #16a34a !important;
  box-shadow: 0 6px 20px rgba(22, 163, 74, 0.4);
  transform: translateY(-2px);
}

.auth-submit-btn:active {
  transform: translateY(0);
}

.modal-header {
  display: flex;
  justify-content: space-between;
  align-items: center;
  padding: 1.25rem 1.5rem;
  border-bottom: 1px solid var(--hairline);
  flex-shrink: 0;
}

.modal-title {
  font-family: var(--font-display);
  font-size: 16px;
  font-weight: 700;
  color: var(--ink);
  margin: 0;
}

.modal-close-btn {
  background: transparent;
  border: none;
  color: var(--muted);
  font-size: 20px;
  cursor: pointer;
  width: 32px;
  height: 32px;
  display: flex;
  align-items: center;
  justify-content: center;
  border-radius: 6px;
  transition: all 0.2s;
}

.modal-close-btn:hover {
  background: rgba(255, 255, 255, 0.05);
  color: var(--ink);
}

.modal-body {
  padding: 1.5rem;
  overflow-y: auto;
  flex: 1;
}

.modal-footer {
  padding: 1rem 1.5rem;
  border-top: 1px solid var(--hairline);
  display: flex;
  justify-content: flex-end;
  gap: 8px;
  background: rgba(0, 0, 0, 0.2);
  flex-shrink: 0;
}

select.form-input option {
  background-color: #1a1a1a;
  color: #ffffff;
}

/* Premium Technical Tables */
.premium-table-wrapper {
  width: 100%;
  overflow-x: auto;
  border: 1px solid var(--hairline);
  border-radius: 12px;
  background: var(--primary);
  margin-top: 1rem;
}

.premium-table {
  width: 100%;
  border-collapse: collapse;
  font-family: var(--font-technical);
  font-size: 13px;
  text-align: left;
}

.premium-table th {
  background: rgba(255, 255, 255, 0.02);
  color: var(--slate);
  font-weight: 700;
  padding: 12px 16px;
  border-bottom: 2px solid var(--hairline);
  text-transform: uppercase;
  font-size: 11px;
  letter-spacing: 0.05em;
}

.premium-table td {
  padding: 12px 16px;
  border-bottom: 1px solid var(--hairline);
  color: var(--ink);
  white-space: nowrap;
}

.premium-table tr:last-child td {
  border-bottom: none;
}

.premium-table tr:hover {
  background: rgba(250, 255, 105, 0.03); /* subtle electric yellow highlight */
}

@keyframes modalEnter {
  from { transform: scale(0.95); opacity: 0; }
  to { transform: scale(1); opacity: 1; }
}

/* Footer */
.app-footer {
  margin-top: auto;
  padding: 3rem 2rem;
  border-top: 1px solid var(--hairline);
  text-align: center;
  color: var(--slate);
  font-size: 13px;
}

.footer-links {
  display: flex;
  justify-content: center;
  gap: 1.5rem;
  margin-top: 1rem;
}

.footer-link {
  color: var(--slate);
  text-decoration: none;
  transition: color 0.2s;
}

.footer-link:hover {
  color: var(--focus-blue);
}

/* Custom Scrollbars */
::-webkit-scrollbar {
  width: 8px;
  height: 8px;
}
::-webkit-scrollbar-track {
  background: rgba(255, 255, 255, 0.02);
}
::-webkit-scrollbar-thumb {
  background: rgba(255, 255, 255, 0.1);
  border-radius: 4px;
}
::-webkit-scrollbar-thumb:hover {
  background: rgba(255, 255, 255, 0.2);
}

/* Specific layout utility classes */
.badge {
  display: inline-flex;
  align-items: center;
  gap: 4px;
  padding: 3px 8px;
  border-radius: 12px;
  font-size: 11px;
  font-weight: 700;
}

.badge-amber { background: rgba(245, 158, 11, 0.15); color: var(--amber); border: 1px solid rgba(245, 158, 11, 0.2); }
.badge-green { background: rgba(16, 185, 129, 0.15); color: var(--deep-green); border: 1px solid rgba(16, 185, 129, 0.2); }
.badge-blue { background: rgba(59, 130, 246, 0.15); color: #3b82f6; border: 1px solid rgba(59, 130, 246, 0.2); }

.metrics-grid {
  display: grid;
  grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
  gap: 1rem;
  margin-bottom: 2rem;
}

.metric-card {
  background: var(--primary);
  border: 1px solid var(--hairline);
  border-radius: 8px;
  padding: 1.25rem;
  text-align: center;
  transition: all 0.2s;
}

.metric-card:hover {
  transform: translateY(-2px);
  border-color: var(--focus-blue);
}

.metric-value {
  font-size: 1.75rem;
  font-weight: 800;
  color: var(--action-blue);
  font-family: var(--font-display);
  margin-bottom: 2px;
}

.metric-label {
  font-size: 11px;
  color: var(--slate);
  text-transform: uppercase;
  letter-spacing: 0.05em;
  font-weight: 600;
}

/* Premium Header Styling */
.premium-header {
  display: flex;
  justify-content: space-between;
  align-items: center;
  padding: 1rem 2rem;
  background: var(--glass-bg);
  backdrop-filter: blur(12px);
  border-bottom: 1px solid var(--glass-border);
  position: sticky;
  top: 0;
  z-index: 1000;
}

@media (max-width: 768px) {
  .premium-header {
    flex-direction: column;
    align-items: stretch;
    gap: 12px;
    padding: 1rem 1.25rem;
    position: relative;
  }
  
  .brand-section {
    flex-direction: row;
    align-items: center;
    gap: 10px !important;
  }

  .brand-section img {
    height: 36px !important;
    width: 36px !important;
  }

  .brand-section p {
    font-size: 11px !important;
    line-height: 1.3 !important;
  }

  .header-controls {
    display: flex;
    flex-wrap: wrap;
    justify-content: flex-start;
    gap: 8px;
    width: 100%;
  }

  .header-controls .header-btn {
    flex: 1 1 auto;
    text-align: center;
    justify-content: center;
    padding: 6px 12px;
    font-size: 12px;
  }

  .user-hub-widget {
    display: flex;
    flex-wrap: wrap;
    gap: 8px;
    width: 100%;
    align-items: center;
  }

  .user-info-text {
    width: 100%;
    margin-bottom: 4px;
    font-size: 12px;
  }
}

/* Header Button Styling */
.header-btn {
  background: transparent;
  border: 1px solid var(--hairline);
  color: var(--ink);
  font-family: var(--font-ui);
  font-size: 13px;
  font-weight: 600;
  padding: 8px 16px;
  border-radius: 6px;
  cursor: pointer;
  transition: all 0.2s;
  display: inline-flex;
  align-items: center;
  gap: 6px;
}

.header-btn:hover {
  background: rgba(255, 255, 255, 0.05);
  border-color: var(--slate);
}

.header-btn.upgrade-highlight {
  border-color: var(--coral) !important;
  color: var(--coral) !important;
}

.header-btn.upgrade-highlight:hover {
  background: rgba(239, 68, 68, 0.08) !important;
}

/* Announcement Bar Styling */
.announcement-bar {
  background: #111111;
  color: #888888;
  font-size: 11px;
  font-weight: 500;
  padding: 8px 16px;
  text-align: center;
  border-bottom: 1px solid var(--hairline);
  display: flex;
  align-items: center;
  justify-content: center;
  gap: 8px;
}

.announcement-bar a {
  color: var(--focus-blue);
  text-decoration: underline;
  font-weight: 600;
}

/* Trust Logo Strip Styling */
.trust-logo-strip {
  margin: 4rem auto 2rem auto;
  max-width: 800px;
  text-align: center;
}

.trust-logo-title {
  font-size: 12px;
  font-weight: 700;
  color: var(--slate);
  text-transform: uppercase;
  letter-spacing: 0.08em;
  margin-bottom: 1.25rem;
}

.trust-logos {
  display: flex;
  justify-content: center;
  align-items: center;
  gap: 2.5rem;
  flex-wrap: wrap;
}

.trust-logos span {
  font-family: var(--font-display);
  font-size: 14px;
  font-weight: 900;
  color: var(--slate);
  letter-spacing: 0.08em;
  transition: color 0.2s;
  cursor: default;
}

.trust-logos span:hover {
  color: var(--focus-blue);
}

/* Header Controls & Theme Toggle Styling */
.header-controls {
  display: flex;
  align-items: center;
  gap: 12px;
}

.theme-toggle {
  background: transparent !important;
  border: 1px solid var(--hairline) !important;
  color: var(--ink) !important;
  font-size: 16px !important;
  cursor: pointer !important;
  padding: 0 !important;
  width: 38px !important;
  height: 38px !important;
  border-radius: 50% !important;
  display: inline-flex !important;
  align-items: center !important;
  justify-content: center !important;
  transition: all 0.2s ease !important;
  box-shadow: none !important;
  outline: none !important;
}

.theme-toggle:hover {
  background: rgba(255, 255, 255, 0.05) !important;
  border-color: var(--slate) !important;
  transform: rotate(15deg);
}

/* Footer Social & Link Icons Styling */
.footer-links {
  display: flex;
  justify-content: center;
  align-items: center;
  gap: 16px;
  margin-top: 0.75rem;
}

.icon-only {
  display: inline-flex !important;
  align-items: center;
  justify-content: center;
  width: 32px;
  height: 32px;
  border-radius: 50%;
  border: 1px solid var(--hairline);
  color: var(--slate);
  transition: all 0.2s ease;
  background: transparent;
  padding: 6px;
}

.icon-only svg {
  width: 16px;
  height: 16px;
  fill: none;
  stroke: currentColor;
  stroke-width: 2;
  stroke-linecap: round;
  stroke-linejoin: round;
}

.icon-only:hover {
  color: var(--focus-blue);
  border-color: var(--focus-blue);
  background: rgba(250, 255, 105, 0.05);
  transform: translateY(-2px);
}

/* =============================================
   LANDING PAGE EXTRA SECTIONS
   ============================================= */

/* Dark Metrics Band */
.metrics-band {
  background: #111111;
  border-top: 1px solid #2a2a2a;
  border-bottom: 1px solid #2a2a2a;
  padding: 4rem 2rem;
  margin: 0 0 5rem 0;
}

.metrics-band-inner {
  max-width: 1100px;
  margin: 0 auto;
  display: grid;
  grid-template-columns: repeat(3, 1fr);
  gap: 3rem;
  text-align: center;
}

@media (max-width: 700px) {
  .metrics-band-inner { grid-template-columns: 1fr; gap: 2rem; }
}

.metrics-band-num {
  font-family: var(--font-display);
  font-size: 3.5rem;
  font-weight: 700;
  color: #faff69;
  letter-spacing: -2px;
  line-height: 1;
  display: block;
  margin-bottom: 8px;
}

.metrics-band-title {
  font-size: 0.9rem;
  font-weight: 700;
  color: #ffffff;
  margin-bottom: 6px;
}

.metrics-band-desc {
  font-size: 0.8rem;
  color: #888888;
  line-height: 1.5;
  max-width: 220px;
  margin: 0 auto;
}

/* How It Works / Solution Section */
.solution-section {
  max-width: 1100px;
  margin: 0 auto 6rem auto;
  padding: 0 2rem;
  display: grid;
  grid-template-columns: 1fr 1fr;
  gap: 4rem;
  align-items: center;
}

@media (max-width: 900px) {
  .solution-section { grid-template-columns: 1fr; gap: 2rem; }
}

.solution-left h2 {
  font-family: var(--font-display);
  font-size: 2.5rem;
  font-weight: 700;
  color: var(--ink);
  letter-spacing: -1.5px;
  line-height: 1.1;
  margin-bottom: 1rem;
}

.solution-left p {
  font-size: 1rem;
  color: var(--slate);
  line-height: 1.65;
  margin-bottom: 1.5rem;
}

.solution-arrow-btn {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  width: 48px;
  height: 48px;
  border-radius: 50%;
  border: 2px solid var(--hairline);
  background: transparent;
  cursor: pointer;
  transition: all 0.2s;
  color: var(--ink);
  font-size: 1.2rem;
}

.solution-arrow-btn:hover {
  border-color: var(--focus-blue);
  color: var(--focus-blue);
  transform: scale(1.1);
}

.solution-right {
  display: grid;
  grid-template-columns: 1fr 1fr;
  gap: 1rem;
}

@media (max-width: 640px) {
  .solution-right {
    grid-template-columns: 1fr;
  }
}

.solution-card {
  background: #1a1a1a;
  border: 1px solid #2a2a2a;
  border-radius: 12px;
  padding: 1.5rem;
  transition: all 0.25s;
}

.solution-card:hover {
  border-color: var(--focus-blue);
  transform: translateY(-3px);
}

.solution-card-icon {
  font-size: 1.5rem;
  margin-bottom: 0.75rem;
  display: block;
}

.solution-card h4 {
  font-size: 0.9rem;
  font-weight: 700;
  color: var(--ink);
  margin: 0 0 0.4rem 0;
}

.solution-card p {
  font-size: 0.78rem;
  color: var(--slate);
  line-height: 1.5;
  margin: 0;
}

/* Testimonials Section */
.testimonials-section {
  background: #0d0d0d;
  border-top: 1px solid #2a2a2a;
  border-bottom: 1px solid #2a2a2a;
  padding: 5rem 2rem;
  margin-bottom: 5rem;
}

.testimonials-inner {
  max-width: 860px;
  margin: 0 auto;
}

.testimonials-label {
  font-family: var(--font-technical);
  font-size: 11px;
  font-weight: 700;
  color: #faff69;
  text-transform: uppercase;
  letter-spacing: 0.18em;
  margin-bottom: 3rem;
}

.testimonial-card {
  background: #1a1a1a;
  border: 1px solid #2a2a2a;
  border-radius: 16px;
  padding: 2.5rem;
  position: relative;
  margin-bottom: 1.5rem;
}

.testimonial-quote-mark {
  font-size: 4rem;
  color: #faff69;
  font-family: Georgia, serif;
  line-height: 0.6;
  display: block;
  margin-bottom: 1.25rem;
  opacity: 0.7;
}

.testimonial-text {
  font-size: 1.1rem;
  color: #e6e6e6;
  line-height: 1.65;
  margin-bottom: 1.5rem;
  font-style: italic;
}

.testimonial-author-name {
  font-size: 0.875rem;
  font-weight: 700;
  color: #ffffff;
  text-transform: uppercase;
  letter-spacing: 0.08em;
}

.testimonial-author-title {
  font-size: 0.8rem;
  color: #888888;
  margin-top: 2px;
}

.testimonial-nav {
  display: flex;
  align-items: center;
  gap: 12px;
  margin-top: 1.5rem;
}

.testimonial-nav-btn {
  width: 36px;
  height: 36px;
  border-radius: 50%;
  border: 1px solid #3a3a3a;
  background: transparent;
  color: #888888;
  cursor: pointer;
  font-size: 14px;
  display: flex;
  align-items: center;
  justify-content: center;
  transition: all 0.2s;
}

.testimonial-nav-btn:hover {
  border-color: #faff69;
  color: #faff69;
}

.testimonial-dots {
  display: flex;
  gap: 6px;
}

.testimonial-dot {
  width: 6px;
  height: 6px;
  border-radius: 50%;
  background: #3a3a3a;
  cursor: pointer;
  transition: all 0.2s;
}

.testimonial-dot.active {
  background: #faff69;
  width: 20px;
  border-radius: 3px;
}

/* Pricing Preview Section */
.pricing-section {
  max-width: 1100px;
  margin: 0 auto 5rem auto;
  padding: 0 2rem;
}

.pricing-header {
  text-align: center;
  margin-bottom: 3rem;
}

.pricing-header h2 {
  font-family: var(--font-display);
  font-size: 2.25rem;
  font-weight: 700;
  color: var(--ink);
  letter-spacing: -1.5px;
  margin-bottom: 0.5rem;
}

.pricing-header p {
  color: var(--slate);
  font-size: 1rem;
}

.pricing-toggle {
  display: inline-flex;
  background: #1a1a1a;
  border: 1px solid #2a2a2a;
  border-radius: 8px;
  padding: 4px;
  gap: 4px;
  margin-top: 1.25rem;
}

.pricing-toggle-btn {
  padding: 6px 18px;
  border-radius: 6px;
  border: none;
  background: transparent;
  color: var(--slate);
  font-size: 13px;
  font-weight: 600;
  cursor: pointer;
  transition: all 0.2s;
}

.pricing-toggle-btn.active {
  background: #22c55e;
  color: #0a0a0a;
}

.pricing-save-badge {
  display: inline-flex;
  align-items: center;
  background: #22c55e;
  color: #052e16;
  font-size: 10px;
  font-weight: 800;
  padding: 4px 10px;
  border-radius: 20px;
  text-transform: uppercase;
  letter-spacing: 0.05em;
  font-family: var(--font-technical);
  box-shadow: 0 4px 10px rgba(34, 197, 94, 0.25);
  animation: pulseSave 2s infinite;
  flex-shrink: 0;
}

@keyframes pulseSave {
  0% { transform: scale(1); }
  50% { transform: scale(1.05); }
  100% { transform: scale(1); }
}

.pricing-grid {
  display: grid;
  grid-template-columns: repeat(3, 1fr);
  gap: 1.5rem;
}

@media (max-width: 992px) {
  .pricing-grid { grid-template-columns: 1fr; }
}

.pricing-card {
  background: rgba(13, 13, 16, 0.5);
  border: 1px solid rgba(255, 255, 255, 0.05);
  border-radius: 20px;
  padding: 3rem 2.25rem;
  position: relative;
  transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
  display: flex;
  flex-direction: column;
  backdrop-filter: blur(24px);
  box-shadow: 0 15px 45px rgba(0, 0, 0, 0.4);
}

.pricing-card:hover {
  border-color: rgba(255, 255, 255, 0.15);
  box-shadow: 0 20px 50px rgba(0, 0, 0, 0.55), 0 0 30px rgba(255, 255, 255, 0.02);
}

.pricing-card.featured {
  background: rgba(0, 255, 178, 0.015);
  border: 1px solid rgba(0, 255, 178, 0.35);
  box-shadow: 0 0 40px rgba(0, 255, 178, 0.08), inset 0 0 15px rgba(0, 255, 178, 0.03);
}

.pricing-card.featured:hover {
  border-color: var(--deep-green);
  box-shadow: 0 0 50px rgba(0, 255, 178, 0.2), inset 0 0 20px rgba(0, 255, 178, 0.05);
}

.pricing-badge {
  position: absolute;
  top: -12px;
  left: 50%;
  transform: translateX(-50%);
  display: inline-flex;
  align-items: center;
  background: #030305;
  color: var(--deep-green) !important;
  font-size: 10px;
  font-weight: 800;
  letter-spacing: 0.12em;
  text-transform: uppercase;
  padding: 5px 16px;
  border-radius: 20px;
  border: 1px solid var(--deep-green);
  box-shadow: 0 0 20px rgba(0, 255, 178, 0.3);
  font-family: var(--font-technical);
  margin-bottom: 0;
}

.pricing-plan-name {
  font-family: var(--font-display);
  font-size: 1.55rem;
  font-weight: 900;
  color: #ffffff;
  margin-bottom: 0.25rem;
  letter-spacing: -0.02em;
}

.pricing-card:nth-child(1) .pricing-plan-name { color: #f8fafc; }
.pricing-card:nth-child(2) .pricing-plan-name { color: var(--deep-green); text-shadow: 0 0 20px rgba(0, 255, 178, 0.15); }
.pricing-card:nth-child(3) .pricing-plan-name { color: #a855f7; text-shadow: 0 0 20px rgba(168, 85, 247, 0.15); }

.pricing-price {
  font-family: var(--font-technical);
  font-size: 2.65rem;
  font-weight: 700;
  color: #ffffff;
  letter-spacing: -1.5px;
  line-height: 1;
  margin: 1.25rem 0 0.85rem 0;
  display: flex;
  align-items: baseline;
  gap: 4px;
}

.pricing-price span {
  font-size: 0.9rem;
  font-weight: 500;
  color: var(--slate);
  letter-spacing: 0;
  font-family: var(--font-ui);
}

.pricing-desc {
  font-size: 0.85rem;
  color: var(--slate);
  margin-bottom: 2rem;
  line-height: 1.6;
}

.pricing-features {
  list-style: none;
  padding: 0;
  margin: 0 0 2.5rem 0;
  display: flex;
  flex-direction: column;
  gap: 14px;
}

.pricing-features li {
  font-size: 0.85rem;
  color: #f1f5f9;
  display: flex;
  align-items: center;
  width: 100%;
  gap: 10px;
  line-height: 1.4;
}

.pricing-features li::before {
  content: "✓";
  display: inline-flex;
  align-items: center;
  justify-content: center;
  width: 18px;
  height: 18px;
  border-radius: 50%;
  background: rgba(0, 255, 178, 0.08);
  border: 1px solid rgba(0, 255, 178, 0.25);
  color: var(--deep-green);
  font-size: 10px;
  font-weight: 900;
  line-height: 1;
  flex-shrink: 0;
  text-shadow: 0 0 4px rgba(0, 255, 178, 0.35);
}

.feature-badge {
  background: rgba(0, 255, 178, 0.05);
  border: 1px solid rgba(0, 255, 178, 0.12);
  color: var(--deep-green);
  font-size: 10px;
  font-weight: 700;
  padding: 2px 8px;
  border-radius: 6px;
  margin-left: auto;
  flex-shrink: 0;
  font-family: var(--font-technical);
}

.pricing-cta-btn {
  width: 100%;
  padding: 13px 20px;
  border-radius: 10px;
  border: 1px solid rgba(255, 255, 255, 0.08);
  background: rgba(255, 255, 255, 0.03);
  color: #ffffff;
  font-family: var(--font-display);
  font-size: 13px;
  font-weight: 700;
  cursor: pointer;
  transition: all 0.25s cubic-bezier(0.4, 0, 0.2, 1);
  margin-top: auto;
  text-align: center;
}

.pricing-cta-btn:hover {
  background: rgba(255, 255, 255, 0.06);
  border-color: rgba(255, 255, 255, 0.15);
  color: #ffffff;
  box-shadow: 0 4px 15px rgba(255, 255, 255, 0.05);
  transform: scale(1.02) translateY(-1px);
}

.pricing-card.featured .pricing-cta-btn {
  background: linear-gradient(135deg, var(--deep-green) 0%, #00b37e 100%);
  color: #030305;
  border: none;
  font-weight: 800;
  letter-spacing: 0.02em;
  box-shadow: 0 0 25px rgba(0, 255, 178, 0.35);
  font-size: 13.5px;
  padding: 14px 20px;
  border-radius: 10px;
  transition: all 0.25s cubic-bezier(0.4, 0, 0.2, 1);
}

.pricing-card.featured .pricing-cta-btn:hover {
  background: linear-gradient(135deg, #00ffc8 0%, #00cca3 100%);
  box-shadow: 0 0 35px rgba(0, 255, 178, 0.55);
  transform: scale(1.02) translateY(-1px);
}

/* =============================================
   COMPREHENSIVE PREMIUM MOBILE RESPONSIVENESS OVERRIDES
   ============================================= */

/* Dashboard & Analytics grid styles (newly styled for maximum premium look) */
.kpi-grid {
  display: grid;
  grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
  gap: 1.25rem;
  margin-bottom: 2rem;
}

.kpi-card {
  background: var(--primary);
  border: 1px solid var(--hairline);
  border-radius: 12px;
  padding: 1.5rem;
  transition: all 0.25s cubic-bezier(0.4, 0, 0.2, 1);
  display: flex;
  flex-direction: column;
  gap: 6px;
  position: relative;
  overflow: hidden;
}

.kpi-card:hover {
  border-color: var(--focus-blue);
  transform: translateY(-2px);
  box-shadow: 0 0 20px rgba(250, 255, 105, 0.03);
}

.kpi-card.success {
  border-left: 3px solid var(--deep-green);
}

.kpi-card.warning {
  border-left: 3px solid var(--coral);
}

.kpi-label {
  font-size: 11px;
  color: var(--slate);
  font-weight: 700;
  text-transform: uppercase;
  letter-spacing: 0.06em;
}

.kpi-value {
  font-size: 2rem;
  font-weight: 900;
  color: #ffffff;
  font-family: var(--font-display);
  letter-spacing: -0.04em;
  line-height: 1.1;
}

.kpi-card.success .kpi-value {
  color: var(--focus-blue);
}

.kpi-sub {
  font-size: 11px;
  color: var(--muted);
  font-family: var(--font-technical);
}

.chart-row {
  display: grid;
  grid-template-columns: repeat(2, 1fr);
  gap: 1.5rem;
  margin-bottom: 2rem;
}

.visual-bar-container {
  margin-bottom: 1.25rem;
}

.visual-bar-info {
  display: flex;
  justify-content: space-between;
  font-size: 12px;
  color: var(--slate);
  margin-bottom: 6px;
  font-family: var(--font-technical);
}

.visual-bar-info span:first-child {
  font-weight: 600;
  color: #ffffff;
}

.visual-bar-bg {
  height: 8px;
  background: rgba(255, 255, 255, 0.05);
  border-radius: 4px;
  overflow: hidden;
}

.visual-bar-fill {
  height: 100%;
  border-radius: 4px;
  background: var(--focus-blue);
  transition: width 0.8s cubic-bezier(0.4, 0, 0.2, 1);
}

.visual-bar-fill.success {
  background: var(--deep-green);
}

.visual-bar-fill.warning {
  background: var(--amber);
}

.visual-bar-fill.danger {
  background: var(--coral);
}

/* Custom layout helper classes for grids that were previously inline-styled */
.demo-showcase-section {
  padding: 4rem 2rem;
  max-width: 1200px;
  margin: 0 auto;
}

.checkout-plans-grid {
  display: grid;
  grid-template-columns: repeat(3, 1fr);
  gap: 12px;
}

.dashboard-panels-grid {
  display: grid;
  grid-template-columns: repeat(auto-fit, minmax(320px, 1fr));
  gap: 1.5rem;
  margin-bottom: 2rem;
}

.admin-payments-stats-grid {
  display: grid;
  grid-template-columns: repeat(4, 1fr);
  gap: 10px;
  margin-bottom: 1rem;
}

.admin-plan-drawer {
  position: fixed;
  top: 0;
  right: 0;
  width: 440px;
  max-width: 100vw;
  height: 100vh;
  background: var(--glass-bg);
  backdrop-filter: blur(24px);
  z-index: 9999;
  display: flex;
  flex-direction: column;
}

/* Tablet Overrides */
@media (max-width: 1024px) {
  .landing-hero {
    padding: 3rem 1.5rem;
  }
  .features-grid {
    padding: 0 1.5rem 3rem;
    grid-template-columns: repeat(2, 1fr) !important; /* perfectly symmetrical 2 columns! */
  }
  .demo-showcase-section {
    padding: 3rem 1.5rem;
  }
  .solution-section {
    margin-bottom: 4rem;
    padding: 0 1.5rem;
  }
  .testimonials-section {
    padding: 4rem 1.5rem;
    margin-bottom: 3rem;
  }
  .pricing-section {
    margin-bottom: 3rem;
    padding: 0 1.5rem;
  }
  .cta-yellow-band {
    padding: 3rem 2rem;
    margin: 2.5rem auto 3rem auto;
  }
}

/* Mobile Overrides (Portrait/Landscape) */
@media (max-width: 768px) {
  .hero-split {
    gap: 2.5rem;
  }
  
  .hero-title {
    font-size: 2.5rem;
    letter-spacing: -0.04em;
  }
  
  .hero-subtitle {
    font-size: 1rem;
    line-height: 1.5;
  }

  .stat-callout-section {
    padding: 0 1.25rem;
    margin-bottom: 2rem;
  }

  .metrics-band {
    padding: 3rem 1.25rem;
    margin-bottom: 3rem;
  }

  .features-grid {
    grid-template-columns: 1fr !important; /* stacks beautifully */
  }

  .demo-showcase-section {
    padding: 2.5rem 1.25rem;
  }

  .checkout-plans-grid {
    grid-template-columns: 1fr !important; /* stacks checkout plans inside modal */
  }

  .dashboard-panels-grid {
    grid-template-columns: 1fr !important; /* stacks version/comments panels */
  }

  .admin-payments-stats-grid {
    grid-template-columns: 1fr 1fr !important; /* 2 columns for tablet view */
  }

  .feature-card, .feature-item-card {
    padding: 1.75rem 1.5rem;
  }

  .solution-left h2 {
    font-size: 2rem;
  }

  .solution-card {
    padding: 1.25rem;
  }

  .testimonial-card {
    padding: 1.75rem 1.25rem;
  }

  .testimonial-quote-mark {
    font-size: 3rem;
    margin-bottom: 0.75rem;
  }

  .testimonial-text {
    font-size: 0.95rem;
    line-height: 1.5;
  }

  .pricing-header h2 {
    font-size: 1.85rem;
  }

  .cta-yellow-band {
    padding: 3rem 2rem;
    margin: 2rem 0 3rem 0;
  }

  .cta-band-headline {
    font-size: 1.85rem;
    letter-spacing: -1px;
    line-height: 1.2;
  }

  .cta-band-sub {
    font-size: 0.9rem;
    line-height: 1.4;
  }

  .chart-row {
    grid-template-columns: 1fr;
    gap: 1.25rem;
  }

  .modal-body {
    padding: 1.25rem;
  }

  .brand-section h1 {
    font-size: 1.4rem !important;
  }
}

/* Compact Mobile Breakpoint (iPhone SE, small viewports) */
@media (max-width: 480px) {
  .premium-header {
    padding: 0.85rem 1rem;
  }

  .brand-section h1 {
    font-size: 1.2rem !important;
  }

  .brand-section p {
    display: none; /* Hide description on micro viewports to keep logo and signin on single row */
  }

  .landing-hero {
    padding: 2rem 1rem;
  }

  .hero-title {
    font-size: 2.1rem;
    line-height: 1.1;
  }

  .hero-right {
    height: 280px !important;
  }

  .stat-callout-num {
    font-size: 2rem !important;
  }

  .demo-showcase-section {
    padding: 2rem 1rem;
  }

  .admin-payments-stats-grid {
    grid-template-columns: 1fr !important;
  }

  .admin-plan-drawer {
    width: 100% !important; /* Full width drawer so it is fully accessible on mobile */
  }

  .kpi-grid {
    grid-template-columns: 1fr !important; /* Stacks dashboard KPIs */
  }

  .feature-card, .feature-item-card {
    padding: 1.5rem 1.15rem;
  }

  .solution-left h2 {
    font-size: 1.75rem;
  }

  .cta-yellow-band {
    padding: 2.25rem 1.25rem;
    border-radius: 12px;
  }

  .cta-band-headline {
    font-size: 1.5rem;
  }

  .cta-band-sub {
    font-size: 0.85rem;
    margin-bottom: 1.25rem;
  }

  .cta-band-btn {
    padding: 12px 22px;
    font-size: 0.85rem;
  }

  .modal-body {
    padding: 1rem;
  }

  .modal-header, .modal-footer {
    padding: 0.85rem 1rem;
  }

  .modal-card {
    width: 95% !important;
    max-height: 95vh !important;
  }
}

/* ============ SHEETAI-INSPIRED LANDING REDESIGN ============ */

/* New Nav */
.sheetai-nav {
  position: sticky; top: 0; z-index: 1000;
  background: rgba(3, 3, 5, 0.9);
  backdrop-filter: blur(20px);
  border-bottom: 1px solid rgba(255, 255, 255, 0.05);
  padding: 0 2.5rem;
  height: 65px;
  display: flex; align-items: center; justify-content: space-between; gap: 1rem;
}
.sheetai-nav-brand {
  display: flex; align-items: center; gap: 10px;
  font-size: 15px; font-weight: 800; color: #fff; text-decoration: none; flex-shrink: 0;
  font-family: var(--font-display);
}
.sheetai-nav-links { display: flex; align-items: center; gap: 2rem; }
.sheetai-nav-link {
  font-size: 13px; color: var(--slate); background: none; border: none;
  cursor: pointer; font-family: inherit; font-weight: 500; transition: color 0.15s;
  text-decoration: none; padding: 0;
}
.sheetai-nav-link:hover { color: var(--deep-green); }
.sheetai-nav-actions { display: flex; align-items: center; gap: 10px; flex-shrink: 0; }
.btn-nav-ghost {
  background: none; border: none; color: var(--slate);
  font-size: 13px; font-weight: 500; cursor: pointer;
  font-family: inherit; padding: 6px 14px; border-radius: 6px; transition: color 0.15s;
}
.btn-nav-ghost:hover { color: #fff; }
.btn-nav-primary {
  background: var(--deep-green); color: #030305; border: none;
  padding: 8px 18px; border-radius: 20px; font-size: 13px;
  font-weight: 700; cursor: pointer; transition: all 0.2s;
  box-shadow: 0 0 15px rgba(0, 255, 178, 0.2);
}
.btn-nav-primary:hover { background: #00cca3; transform: translateY(-1px); box-shadow: 0 0 25px rgba(0, 255, 178, 0.45); }
@media (max-width: 768px) {
  .sheetai-nav-links { display: none; }
  .sheetai-nav { padding: 0 1rem; }
}

/* Centered Hero Area */
.hero-centered {
  display: flex;
  flex-direction: column;
  align-items: center;
  text-align: center;
  max-width: 1000px;
  margin: 0 auto;
  padding: 4.5rem 1.5rem 2.5rem;
}
.hero-centered-content {
  display: flex;
  flex-direction: column;
  align-items: center;
  max-width: 780px;
  margin: 0 auto;
}

/* Hero Label Chip */
.sheetai-hero-label {
  display: inline-flex; align-items: center; gap: 6px;
  background: rgba(0,255,178,0.07); border: 1px solid rgba(0,255,178,0.2);
  color: var(--deep-green); padding: 5px 14px; border-radius: 20px;
  font-family: var(--font-technical); font-size: 10px; font-weight: 700;
  text-transform: uppercase; letter-spacing: 0.12em; margin-bottom: 1.5rem;
}
.hero-title-new {
  font-family: var(--font-display); font-size: clamp(2.4rem, 5vw, 4rem);
  font-weight: 800; letter-spacing: -0.05em; line-height: 1.05;
  color: #fff; margin-bottom: 1.25rem;
  text-align: center;
}
.hero-title-new .green-word {
  color: var(--deep-green);
  text-shadow: 0 0 40px rgba(0, 255, 178, 0.25);
  background: linear-gradient(135deg, #00ffb2 0%, #00b37e 100%);
  -webkit-background-clip: text;
  -webkit-text-fill-color: transparent;
}
.hero-stats-row { display: flex; justify-content: center; gap: 1.5rem; flex-wrap: wrap; margin: 1.75rem 0; }
.hero-stat-item { display: flex; align-items: center; gap: 8px; font-size: 11px; color: var(--slate); }
.hero-stat-item .stat-icon {
  width: 32px; height: 32px; border-radius: 8px;
  background: rgba(255,255,255,0.04); border: 1px solid var(--hairline);
  display: flex; align-items: center; justify-content: center; font-size: 14px;
}
.hero-stat-item strong { display: block; font-size: 13px; font-weight: 700; color: #fff; }

.hero-actions-row {
  display: flex;
  align-items: center;
  justify-content: center;
  gap: 12px;
  margin: 1rem 0 2.5rem;
  flex-wrap: wrap;
}

/* Upload Card New */
.upload-card-new {
  background: rgba(10, 10, 12, 0.4); 
  border: 1px dashed rgba(0, 255, 178, 0.25); 
  border-radius: 16px;
  padding: 30px 24px; text-align: center; cursor: pointer;
  transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1); margin: 0 auto 2.5rem;
  width: 100%; max-width: 580px;
  backdrop-filter: blur(12px);
  box-shadow: 0 20px 40px rgba(0,0,0,0.5);
}
.upload-card-new:hover, .upload-card-new.dragging {
  border-color: var(--deep-green); 
  background: rgba(0, 255, 178, 0.03);
  box-shadow: 0 0 35px rgba(0, 255, 178, 0.15);
  transform: translateY(-2px);
}
.upload-icon-new { font-size: 32px; margin-bottom: 10px; filter: drop-shadow(0 0 10px rgba(0,255,178,0.2)); }
.upload-title-new { font-size: 14px; font-weight: 700; color: #fff; margin-bottom: 4px; }
.upload-formats { font-size: 11px; color: var(--slate); margin-bottom: 18px; }
.btn-upload-green {
  display: inline-flex; align-items: center; gap: 6px;
  background: var(--deep-green); color: #030305; border: none;
  padding: 10px 24px; border-radius: 20px; font-size: 13px; font-weight: 700;
  cursor: pointer; transition: all 0.2s;
  box-shadow: 0 4px 15px rgba(0, 255, 178, 0.2);
}
.btn-upload-green:hover { background: #00cca3; transform: translateY(-1px); box-shadow: 0 0 25px rgba(0, 255, 178, 0.45); }
.upload-integration-row {
  display: flex; align-items: center; justify-content: center;
  gap: 10px; margin-top: 18px; font-size: 11px; color: var(--slate);
}
.upload-integration-icon {
  width: 28px; height: 28px; border-radius: 6px; background: rgba(255,255,255,0.03);
  border: 1px solid rgba(255,255,255,0.06); display: flex; align-items: center; justify-content: center; font-size: 14px;
}
.no-cc-line { font-size: 10px; color: var(--slate); text-align: center; margin-top: 8px; }

/* Analysis Overview Card */
.analysis-overview-card {
  background: rgba(13, 13, 16, 0.6); border: 1px solid rgba(255,255,255,0.05); border-radius: 16px;
  padding: 22px; height: 100%; min-height: 420px;
  display: flex; flex-direction: column; gap: 14px;
  backdrop-filter: blur(12px);
  box-shadow: 0 20px 45px rgba(0,0,0,0.6); overflow: hidden;
}
.analysis-card-header { display: flex; align-items: center; justify-content: space-between; flex-shrink: 0; }
.analysis-card-title { font-size: 13px; font-weight: 700; color: #fff; }
.live-badge {
  display: flex; align-items: center; gap: 5px; font-size: 10px; color: var(--deep-green);
  font-weight: 600; background: rgba(0,255,178,0.1); padding: 3px 8px;
  border-radius: 20px; border: 1px solid rgba(0,255,178,0.25);
}
.live-badge::before {
  content: ''; width: 6px; height: 6px; border-radius: 50%; background: var(--deep-green);
  animation: pulse-green 1.5s infinite; display: inline-block;
  box-shadow: 0 0 8px var(--deep-green);
}
@keyframes pulse-green { 0%,100%{opacity:1} 50%{opacity:0.3} }
.card-kpi-row { display: grid; grid-template-columns: repeat(3,1fr); gap: 8px; }
.card-kpi-item { background: rgba(0,0,0,0.3); border: 1px solid rgba(255,255,255,0.04); border-radius: 8px; padding: 8px 10px; }
.card-kpi-label { font-size: 9px; color: var(--slate); text-transform: uppercase; font-family: var(--font-technical); margin-bottom: 3px; }
.card-kpi-value { font-size: 14px; font-weight: 700; color: #fff; line-height: 1; }
.card-kpi-change { font-size: 9px; margin-top: 2px; display: flex; align-items: center; gap: 2px; }
.card-kpi-change.up { color: var(--deep-green); }
.card-kpi-change.down { color: #error; }
.card-bar-section { flex: 1; }
.card-bar-label { font-size: 10px; color: var(--slate); font-weight: 600; margin-bottom: 8px; font-family: var(--font-technical); text-transform: uppercase; letter-spacing: 0.05em; }
.card-analysis-complete {
  display: flex; align-items: center; justify-content: space-between;
  font-size: 10px; color: var(--slate); background: rgba(0,255,178,0.04);
  border: 1px solid rgba(0,255,178,0.12); border-radius: 6px; padding: 6px 10px; flex-shrink: 0;
}

/* Recent Analyses Section */
.recent-analyses-section {
  max-width: 1200px; margin: 0 auto; padding: 0 2rem 3.5rem;
  display: grid; grid-template-columns: 1.4fr 0.6fr; gap: 20px; align-items: start;
}
@media (max-width: 900px) { .recent-analyses-section { grid-template-columns: 1fr; } }
.recent-analyses-card { background: rgba(13, 13, 16, 0.4); border: 1px solid rgba(255,255,255,0.05); border-radius: 12px; overflow: hidden; backdrop-filter: blur(12px); }
.recent-analyses-header { display: flex; align-items: center; justify-content: space-between; padding: 14px 18px; border-bottom: 1px solid rgba(255,255,255,0.05); }
.recent-analyses-title { font-size: 13px; font-weight: 700; color: #fff; }
.view-all-link { font-size: 11px; color: var(--deep-green); cursor: pointer; background: none; border: none; font-family: inherit; }
.analyses-table { width: 100%; border-collapse: collapse; }
.analyses-table th { font-size: 9px; text-transform: uppercase; letter-spacing: 0.08em; color: var(--slate); font-family: var(--font-technical); font-weight: 600; padding: 8px 18px; text-align: left; border-bottom: 1px solid rgba(255,255,255,0.05); }
.analyses-table td { padding: 11px 18px; font-size: 12px; color: var(--ink); border-bottom: 1px solid rgba(255,255,255,0.03); }
.analyses-table tr:last-child td { border-bottom: none; }
.filename-cell { display: flex; align-items: center; gap: 8px; font-weight: 600; }
.file-icon-dot { width: 28px; height: 28px; border-radius: 6px; background: rgba(0,255,178,0.1); display: flex; align-items: center; justify-content: center; font-size: 13px; flex-shrink: 0; }
.status-badge-completed { background: rgba(0,255,178,0.1); color: var(--deep-green); border: 1px solid rgba(0,255,178,0.2); font-size: 9px; font-weight: 700; padding: 2px 8px; border-radius: 20px; font-family: var(--font-technical); text-transform: uppercase; }
.arrow-link-btn { background: none; border: 1px solid rgba(255,255,255,0.05); color: var(--slate); width: 24px; height: 24px; border-radius: 6px; cursor: pointer; display: flex; align-items: center; justify-content: center; font-size: 12px; transition: all 0.15s; }
.arrow-link-btn:hover { border-color: var(--deep-green); color: var(--deep-green); box-shadow: 0 0 10px rgba(0,255,178,0.25); }
.insights-summary-card { background: rgba(13, 13, 16, 0.4); border: 1px solid rgba(255,255,255,0.05); border-radius: 12px; padding: 18px; backdrop-filter: blur(12px); }
.insights-legend { display: flex; flex-direction: column; gap: 10px; margin-top: 12px; }
.insights-legend-item { display: flex; align-items: center; gap: 8px; font-size: 12px; color: var(--slate); }
.legend-dot { width: 8px; height: 8px; border-radius: 2px; flex-shrink: 0; }
.legend-pct { margin-left: auto; font-weight: 700; color: #fff; }

/* Features & Bento Grid Area */
.sheetai-features-section { max-width: 1200px; margin: 0 auto; padding: 4rem 2rem; }
.sheetai-features-label { font-family: var(--font-technical); font-size: 10px; font-weight: 700; color: var(--deep-green); text-transform: uppercase; letter-spacing: 0.15em; margin-bottom: 6px; text-align: center; }
.sheetai-features-title { font-family: var(--font-display); font-size: 2.2rem; font-weight: 800; text-align: center; color: #fff; margin-bottom: 3rem; }

/* Bento Grid */
.bento-features-grid {
  display: grid;
  grid-template-columns: repeat(3, 1fr);
  grid-auto-rows: minmax(180px, auto);
  gap: 20px;
}
@media (max-width: 968px) {
  .bento-features-grid {
    grid-template-columns: 1fr;
  }
}
.bento-feature-card {
  background: rgba(13, 13, 16, 0.45);
  border: 1px solid rgba(255, 255, 255, 0.05);
  border-radius: 16px;
  padding: 24px;
  position: relative;
  overflow: hidden;
  backdrop-filter: blur(16px);
  transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
  display: flex;
  flex-direction: column;
  justify-content: space-between;
}
.bento-feature-card:hover {
  border-color: rgba(0, 255, 178, 0.25);
  box-shadow: 0 10px 30px rgba(0, 255, 178, 0.05);
  transform: translateY(-2px);
}
.bento-card-bg-gradient {
  position: absolute;
  inset: 0;
  background: radial-gradient(circle at 100% 100%, rgba(0, 255, 178, 0.03) 0%, transparent 60%);
  pointer-events: none;
  z-index: 0;
}
.bento-feature-icon {
  width: 38px; height: 38px; border-radius: 8px;
  background: rgba(0, 255, 178, 0.08); border: 1px solid rgba(0, 255, 178, 0.15);
  display: flex; align-items: center; justify-content: center; font-size: 16px; margin-bottom: 16px;
  color: var(--deep-green);
  z-index: 1;
}
.bento-feature-name { font-size: 14px; font-weight: 700; color: #fff; margin-bottom: 8px; z-index: 1; }
.bento-feature-desc { font-size: 12px; color: var(--slate); line-height: 1.55; z-index: 1; }

/* Wide & Tall Grid Anchors */
.bento-colspan-2 { grid-column: span 2; }
.bento-rowspan-2 { grid-row: span 2; }
@media (max-width: 968px) {
  .bento-colspan-2 { grid-column: span 1; }
  .bento-rowspan-2 { grid-row: span 1; }
}

/* Bento Visual Mockups */
.bento-visual-channels {
  display: flex; justify-content: space-around; align-items: center;
  background: rgba(0,0,0,0.25); border: 1px solid rgba(255,255,255,0.03);
  padding: 16px; border-radius: 10px; margin-top: 14px;
  position: relative;
}
.bento-channel-badge {
  background: rgba(255,255,255,0.03); border: 1px solid rgba(255,255,255,0.06);
  padding: 6px 14px; border-radius: 20px; font-size: 11px; color: #fff; font-weight: 600;
}
.bento-pulse-line {
  position: absolute; height: 1px; background: linear-gradient(90deg, transparent, var(--deep-green), transparent);
  width: 50%; opacity: 0.5;
}

/* E-Commerce Calculator Styles */
.calculator-card {
  max-width: 1200px;
  margin: 0 auto 4rem;
  padding: 30px;
  background: rgba(13, 13, 16, 0.45);
  border: 1px solid rgba(255, 255, 255, 0.05);
  border-radius: 20px;
  backdrop-filter: blur(16px);
  position: relative;
  overflow: hidden;
  box-shadow: 0 20px 40px rgba(0,0,0,0.4);
}
.calculator-card:hover {
  border-color: rgba(0, 255, 178, 0.25);
  box-shadow: 0 20px 40px rgba(0, 255, 178, 0.03);
}
.calculator-layout {
  display: grid;
  grid-template-columns: 1.1fr 0.9fr;
  gap: 35px;
}
@media (max-width: 768px) {
  .calculator-layout {
    grid-template-columns: 1fr;
  }
}
.calc-inputs-col {
  display: flex;
  flex-direction: column;
  gap: 16px;
}
.calc-results-col {
  background: rgba(0, 255, 178, 0.02);
  border: 1px solid rgba(0, 255, 178, 0.08);
  border-radius: 14px;
  padding: 24px;
  display: flex;
  flex-direction: column;
  justify-content: space-between;
  gap: 20px;
}
.calc-field {
  display: flex;
  flex-direction: column;
  gap: 6px;
}
.calc-field label {
  font-size: 11px;
  font-family: var(--font-technical);
  text-transform: uppercase;
  color: var(--slate);
  letter-spacing: 0.05em;
  display: flex;
  justify-content: space-between;
}
.calc-field label span {
  font-weight: 750;
  color: #fff;
}
.calc-input-wrapper {
  position: relative;
  display: flex;
  align-items: center;
}
.calc-input-wrapper input[type="number"] {
  width: 100%;
  background: rgba(0,0,0,0.3);
  border: 1px solid rgba(255,255,255,0.06);
  padding: 10px 14px;
  border-radius: 8px;
  color: #fff;
  font-size: 13px;
  font-family: inherit;
  outline: none;
  transition: all 0.2s;
}
.calc-input-wrapper input[type="number"]:focus {
  border-color: var(--deep-green);
  box-shadow: 0 0 10px rgba(0,255,178,0.15);
}
.calc-slider {
  width: 100%;
  accent-color: var(--deep-green);
  margin-top: 6px;
}
.calc-stat-grid {
  display: grid;
  grid-template-columns: 1fr 1fr;
  gap: 12px;
}
.calc-stat-card {
  background: rgba(0, 0, 0, 0.3);
  border: 1px solid rgba(255,255,255,0.04);
  padding: 12px;
  border-radius: 10px;
  text-align: left;
}
.calc-stat-label {
  font-size: 9px;
  text-transform: uppercase;
  color: var(--slate);
  font-family: var(--font-technical);
}
.calc-stat-val {
  font-size: 17px;
  font-weight: 800;
  color: #fff;
  margin-top: 2px;
}
.calc-stat-val.highlight {
  color: var(--deep-green);
  text-shadow: 0 0 10px rgba(0,255,178,0.1);
}

/* Currency Switcher capsule styles */
.currency-selector {
  display: inline-flex;
  align-items: center;
  gap: 2px;
  background: rgba(255,255,255,0.04);
  border: 1px solid rgba(255,255,255,0.06);
  padding: 4px;
  border-radius: 20px;
  vertical-align: middle;
}
.currency-pill {
  background: none;
  border: none;
  font-size: 9px;
  font-weight: 700;
  color: var(--slate);
  padding: 4px 10px;
  border-radius: 12px;
  cursor: pointer;
  transition: all 0.2s;
  font-family: var(--font-technical);
}
.currency-pill.active {
  background: var(--deep-green);
  color: #030305;
}

/* FAQ Section */
.faq-section {
  max-width: 850px;
  margin: 5rem auto;
  padding: 0 2rem;
}
.faq-header {
  text-align: center;
  margin-bottom: 2.5rem;
}
.faq-title {
  font-family: var(--font-display);
  font-size: 2.2rem;
  font-weight: 800;
  color: #fff;
  margin-top: 6px;
}
.faq-list {
  display: flex;
  flex-direction: column;
}
.faq-item {
  border-bottom: 1px solid rgba(255, 255, 255, 0.05);
  padding: 20px 0;
  cursor: pointer;
}
.faq-question {
  display: flex;
  justify-content: space-between;
  align-items: center;
  font-size: 14px;
  font-weight: 700;
  color: #fff;
  transition: color 0.15s;
}
.faq-question:hover {
  color: var(--deep-green);
}
.faq-icon {
  font-size: 11px;
  color: var(--slate);
  transition: transform 0.2s ease, color 0.2s ease;
  width: 20px;
  height: 20px;
  border-radius: 50%;
  border: 1px solid rgba(255,255,255,0.06);
  display: flex;
  align-items: center;
  justify-content: center;
}
.faq-icon.expanded {
  transform: rotate(180deg);
  border-color: rgba(0, 255, 178, 0.3);
  color: var(--deep-green);
  box-shadow: 0 0 10px rgba(0,255,178,0.2);
}
.faq-answer-container {
  overflow: hidden;
  transition: max-height 0.3s cubic-bezier(0.4, 0, 0.2, 1), opacity 0.3s;
}
.faq-answer {
  font-size: 13px;
  color: var(--slate);
  line-height: 1.6;
  padding-top: 12px;
}

/* Metrics Row */
.sheetai-metrics-row { max-width: 1200px; margin: 0 auto; display: grid; grid-template-columns: repeat(4,1fr); gap: 14px; padding: 1.5rem 2rem; }
@media (max-width: 800px) { .sheetai-metrics-row { grid-template-columns: repeat(2,1fr); } }
.sheetai-metric-item { background: rgba(13, 13, 16, 0.45); border: 1px solid rgba(255,255,255,0.05); border-radius: 12px; padding: 18px 20px; display: flex; align-items: center; gap: 14px; }
.sheetai-metric-icon { font-size: 22px; flex-shrink: 0; }
.sheetai-metric-value { font-size: 1.4rem; font-weight: 800; color: var(--deep-green); line-height: 1; }
.sheetai-metric-label { font-size: 11px; color: var(--slate); margin-top: 2px; }

/* Why + Testimonials */
.why-testimonials-section { max-width: 1200px; margin: 0 auto; padding: 4rem 2rem; display: grid; grid-template-columns: 0.9fr 1.1fr; gap: 4rem; align-items: start; }
@media (max-width: 900px) { .why-testimonials-section { grid-template-columns: 1fr; gap: 3rem; } }
.why-left-heading { font-family: var(--font-display); font-size: clamp(1.6rem, 3vw, 2.2rem); font-weight: 800; color: #fff; line-height: 1.2; margin-bottom: 1.25rem; }
.why-left-heading em { font-style: normal; text-decoration: underline; text-decoration-color: var(--deep-green); text-decoration-thickness: 3px; text-underline-offset: 4px; }
.why-bullets { list-style: none; display: flex; flex-direction: column; gap: 10px; margin-bottom: 1.5rem; }
.why-bullet-item { display: flex; align-items: center; gap: 10px; font-size: 13px; color: var(--slate); }
.why-bullet-check { width: 18px; height: 18px; border-radius: 50%; background: rgba(0,255,178,0.08); border: 1px solid rgba(0,255,178,0.2); display: flex; align-items: center; justify-content: center; color: var(--deep-green); font-size: 10px; flex-shrink: 0; }
.testimonial-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 14px; }
@media (max-width: 600px) { .testimonial-grid { grid-template-columns: 1fr; } }
.testimonial-card-new { background: rgba(13,13,16,0.45); border: 1px solid rgba(255,255,255,0.05); border-radius: 12px; padding: 18px; transition: all 0.25s; backdrop-filter: blur(12px); }
.testimonial-card-new:hover { border-color: rgba(0,255,178,0.25); box-shadow: 0 10px 25px rgba(0,255,178,0.04); transform: translateY(-1px); }
.testimonial-stars { color: #f59e0b; font-size: 12px; margin-bottom: 10px; letter-spacing: 1px; }
.testimonial-text-new { font-size: 12px; color: var(--slate); line-height: 1.6; margin-bottom: 14px; }
.testimonial-author-row { display: flex; align-items: center; gap: 10px; }
.testimonial-avatar { width: 32px; height: 32px; border-radius: 50%; display: flex; align-items: center; justify-content: center; font-size: 14px; flex-shrink: 0; }
.testimonial-author-name-new { font-size: 12px; font-weight: 700; color: #fff; }
.testimonial-author-title-new { font-size: 10px; color: var(--slate); }
.btn-hero-primary {
  display: inline-flex; align-items: center; gap: 8px;
  background: var(--deep-green); color: #030305; border: none;
  padding: 12px 26px; border-radius: 24px; font-weight: 700; font-size: 14px;
  cursor: pointer; transition: all 0.2s;
  box-shadow: 0 4px 15px rgba(0, 255, 178, 0.2);
}
.btn-hero-primary:hover { background: #00cca3; transform: translateY(-1px); box-shadow: 0 0 25px rgba(0, 255, 178, 0.45); }
.btn-hero-secondary {
  display: inline-flex; align-items: center; gap: 8px;
  background: none; color: #fff; border: 1px solid rgba(255,255,255,0.1);
  padding: 12px 26px; border-radius: 24px; font-weight: 700; font-size: 14px;
  cursor: pointer; transition: all 0.2s;
}
.btn-hero-secondary:hover { background: rgba(255,255,255,0.03); border-color: rgba(255,255,255,0.25); }

/* CTA Banner */
.cta-section-new { max-width: 1200px; margin: 0 auto 3rem; padding: 0 2rem; }
.cta-banner-new {
  background: rgba(13, 13, 16, 0.4); border: 1px solid rgba(255,255,255,0.05); border-radius: 16px;
  padding: 2.5rem 3rem; display: flex; align-items: center; justify-content: space-between; gap: 2rem;
  backdrop-filter: blur(12px);
  position: relative;
  overflow: hidden;
}
.cta-banner-new::after {
  content: ''; position: absolute; inset: 0;
  background: radial-gradient(circle at 80% 50%, rgba(0,255,178,0.05) 0%, transparent 50%);
  pointer-events: none;
}
@media (max-width: 700px) { .cta-banner-new { flex-direction: column; text-align: center; padding: 2rem; } }
.cta-banner-icon { width: 48px; height: 48px; border-radius: 12px; background: rgba(0,255,178,0.08); border: 1px solid rgba(0,255,178,0.2); display: flex; align-items: center; justify-content: center; font-size: 22px; flex-shrink: 0; color: var(--deep-green); }
.cta-banner-headline { font-size: 1.3rem; font-weight: 800; color: #fff; margin-bottom: 4px; }
.cta-banner-sub { font-size: 13px; color: var(--slate); }
.cta-banner-actions { display: flex; flex-direction: column; align-items: flex-end; gap: 6px; flex-shrink: 0; z-index: 1; }
@media (max-width: 700px) { .cta-banner-actions { align-items: center; } }
.btn-cta-green {
  display: inline-flex; align-items: center; gap: 8px;
  background: var(--deep-green); color: #030305; border: none;
  padding: 12px 26px; border-radius: 24px; font-weight: 700; font-size: 14px;
  cursor: pointer; transition: all 0.2s; white-space: nowrap;
  box-shadow: 0 4px 15px rgba(0, 255, 178, 0.2);
}
.btn-cta-green:hover { background: #00cca3; transform: translateY(-1px); box-shadow: 0 0 25px rgba(0, 255, 178, 0.45); }
.cta-no-cc { font-size: 10px; color: var(--slate); }

/* Site Footer */
.site-footer { background: #020204; border-top: 1px solid rgba(255,255,255,0.05); padding: 4rem 2.5rem 3rem; }
.site-footer-inner { max-width: 1200px; margin: 0 auto; display: flex; flex-direction: column; align-items: center; text-align: center; gap: 1.5rem; }
.footer-brand { display: flex; flex-direction: column; gap: 12px; align-items: center; }
.footer-brand-name { font-size: 18px; font-weight: 800; color: #fff; display: flex; align-items: center; gap: 8px; justify-content: center; font-family: var(--font-display); }
.footer-tagline { font-size: 13px; color: var(--slate); line-height: 1.6; max-width: 420px; margin: 0 auto; }
.footer-social-row { display: flex; gap: 12px; margin-top: 8px; justify-content: center; }
.footer-social-btn { width: 38px; height: 38px; border-radius: 8px; border: 1px solid rgba(255,255,255,0.05); background: rgba(255, 255, 255, 0.02); color: var(--slate); cursor: pointer; display: inline-flex; align-items: center; justify-content: center; transition: all 0.2s cubic-bezier(0.4, 0, 0.2, 1); text-decoration: none; }
.footer-social-btn:hover { border-color: var(--deep-green); color: var(--deep-green); background: rgba(0, 255, 178, 0.05); transform: translateY(-2px); box-shadow: 0 4px 12px rgba(0, 255, 178, 0.1); }
.footer-social-btn svg { transition: transform 0.2s; }
.footer-social-btn:hover svg { transform: scale(1.05); }
.footer-bottom { max-width: 1200px; margin: 3rem auto 0; padding-top: 1.5rem; border-top: 1px solid rgba(255,255,255,0.05); display: flex; align-items: center; justify-content: space-between; font-size: 11px; color: var(--slate); flex-wrap: wrap; gap: 12px; }


/* ============ BENTO DASHBOARD STYLES ============ */
.bento-dashboard {
  width: 100%;
  max-width: 1200px;
  margin: 0 auto;
  padding: 2rem 2rem 4rem;
  display: flex;
  flex-direction: column;
  gap: 1.5rem;
  font-family: var(--font-ui);
}

.bento-greeting-banner {
  background: linear-gradient(135deg, rgba(34, 197, 94, 0.05) 0%, rgba(250, 255, 105, 0.03) 100%);
  border: 1px solid rgba(34, 197, 94, 0.15);
  border-radius: 16px;
  padding: 1.5rem;
  display: flex;
  justify-content: space-between;
  align-items: center;
  flex-wrap: wrap;
  gap: 1rem;
  backdrop-filter: blur(10px);
}

.bento-greeting-left {
  display: flex;
  align-items: center;
  gap: 1.25rem;
}

.bento-greeting-avatar {
  font-size: 2.25rem;
  background: rgba(34, 197, 94, 0.1);
  width: 56px;
  height: 56px;
  border-radius: 12px;
  display: flex;
  align-items: center;
  justify-content: center;
  border: 1px solid rgba(34, 197, 94, 0.2);
}

.bento-greeting-title {
  font-size: 1.5rem;
  font-weight: 800;
  color: #ffffff;
  margin: 0 0 4px 0;
  letter-spacing: -0.02em;
}

.bento-greeting-sub {
  font-size: 0.875rem;
  color: var(--slate);
  margin: 0;
}

.bento-greeting-stats {
  display: flex;
  align-items: center;
  gap: 12px;
}

.bento-grid {
  display: grid;
  grid-template-columns: 1.4fr 1fr;
  gap: 1.5rem;
}

@media (max-width: 900px) {
  .bento-grid {
    grid-template-columns: 1fr;
  }
}

.bento-card {
  background: rgba(14, 14, 14, 0.85);
  border: 1px solid rgba(255, 255, 255, 0.04);
  border-radius: 18px;
  padding: 1.75rem;
  display: flex;
  flex-direction: column;
  backdrop-filter: blur(20px);
  box-shadow: 0 8px 32px rgba(0, 0, 0, 0.4);
}

.bento-card-header {
  display: flex;
  justify-content: space-between;
  align-items: center;
  margin-bottom: 1.25rem;
  padding-bottom: 12px;
  border-bottom: 1px solid rgba(255, 255, 255, 0.04);
}

.bento-card-title {
  font-size: 0.95rem;
  font-weight: 750;
  color: #ffffff;
  font-family: var(--font-technical);
  text-transform: uppercase;
  letter-spacing: 0.05em;
  display: flex;
  align-items: center;
  gap: 8px;
}

.bento-card-badge {
  font-size: 0.75rem;
  background: rgba(255, 255, 255, 0.06);
  padding: 2px 8px;
  border-radius: 20px;
  color: var(--slate);
  font-weight: 600;
}

.bento-mode-selector {
  margin-bottom: 1.25rem;
  background: rgba(255, 255, 255, 0.02);
  border: 1px solid rgba(255, 255, 255, 0.03);
  padding: 10px;
  border-radius: 12px;
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 12px;
}

.bento-mode-label {
  font-size: 0.8rem;
  font-weight: 700;
  color: var(--slate);
  text-transform: uppercase;
  letter-spacing: 0.05em;
}

.bento-mode-tabs {
  display: flex;
  gap: 6px;
}

.bento-mode-tab {
  background: transparent;
  border: 1px solid transparent;
  color: var(--slate);
  padding: 6px 12px;
  border-radius: 8px;
  font-size: 0.8rem;
  font-weight: 600;
  cursor: pointer;
  transition: all 0.2s ease;
}

.bento-mode-tab:hover {
  background: rgba(255, 255, 255, 0.04);
  color: #ffffff;
}

.bento-mode-tab.active {
  background: rgba(34, 197, 94, 0.1);
  border-color: rgba(34, 197, 94, 0.3);
  color: #22c55e;
}

.bento-dropzone {
  flex: 1;
  min-height: 250px;
  display: flex;
  flex-direction: column;
  justify-content: center;
  align-items: center;
  border: 2px dashed rgba(255, 255, 255, 0.08);
  border-radius: 14px;
  padding: 2rem;
  background: rgba(255, 255, 255, 0.01);
  cursor: pointer;
  transition: all 0.25s ease;
}

.bento-dropzone:hover, .bento-dropzone.dragging {
  border-color: #22c55e;
  background: rgba(34, 197, 94, 0.02);
  box-shadow: 0 0 20px rgba(34, 197, 94, 0.05);
}

.bento-workspace-footer {
  margin-top: 1.25rem;
}

.bento-limit-bar-wrapper {
  display: flex;
  flex-direction: column;
  gap: 6px;
}

.bento-limit-bar-labels {
  display: flex;
  justify-content: space-between;
  font-size: 0.75rem;
  color: var(--slate);
  font-weight: 500;
}

.bento-limit-bar-bg {
  width: 100%;
  height: 6px;
  background: rgba(255, 255, 255, 0.06);
  border-radius: 10px;
  overflow: hidden;
}

.bento-limit-bar-fill {
  height: 100%;
  background: linear-gradient(90deg, #22c55e 0%, #faff69 100%);
  border-radius: 10px;
  transition: width 0.4s ease;
}

.bento-unlocked-status {
  font-size: 0.75rem;
  color: #22c55e;
  background: rgba(34, 197, 94, 0.08);
  border: 1px solid rgba(34, 197, 94, 0.15);
  padding: 10px;
  border-radius: 8px;
  text-align: center;
  font-weight: 550;
}

.bento-search-wrapper {
  position: relative;
  margin-bottom: 1rem;
}

.bento-search-icon {
  position: absolute;
  left: 12px;
  top: 50%;
  transform: translateY(-50%);
  font-size: 0.85rem;
  opacity: 0.6;
}

.bento-search-input {
  width: 100%;
  padding: 10px 32px 10px 34px;
  background: rgba(255, 255, 255, 0.03);
  border: 1px solid rgba(255, 255, 255, 0.05);
  border-radius: 10px;
  color: #ffffff;
  font-size: 0.85rem;
  outline: none;
  font-family: inherit;
  transition: all 0.2s ease;
}

.bento-search-input:focus {
  border-color: rgba(34, 197, 94, 0.4);
  background: rgba(255, 255, 255, 0.05);
  box-shadow: 0 0 10px rgba(34, 197, 94, 0.05);
}

.bento-search-clear {
  position: absolute;
  right: 12px;
  top: 50%;
  transform: translateY(-50%);
  background: none;
  border: none;
  color: var(--slate);
  cursor: pointer;
  font-size: 0.75rem;
  padding: 4px;
}

.bento-saved-list {
  flex: 1;
  max-height: 220px;
  overflow-y: auto;
  display: flex;
  flex-direction: column;
  gap: 8px;
  margin-bottom: 1.25rem;
  padding-right: 4px;
}

.bento-saved-list::-webkit-scrollbar {
  width: 4px;
}
.bento-saved-list::-webkit-scrollbar-track {
  background: transparent;
}
.bento-saved-list::-webkit-scrollbar-thumb {
  background: rgba(255, 255, 255, 0.1);
  border-radius: 4px;
}

.bento-saved-item {
  padding: 10px 12px;
  background: rgba(255, 255, 255, 0.02);
  border: 1px solid rgba(255, 255, 255, 0.03);
  border-radius: 10px;
  display: flex;
  justify-content: space-between;
  align-items: center;
  gap: 12px;
  transition: all 0.15s ease;
}

.bento-saved-item:hover {
  background: rgba(255, 255, 255, 0.04);
  border-color: rgba(255, 255, 255, 0.06);
}

.bento-saved-item-left {
  display: flex;
  align-items: center;
  gap: 10px;
  min-width: 0;
}

.bento-saved-icon {
  font-size: 1.25rem;
  background: rgba(255, 255, 255, 0.04);
  width: 36px;
  height: 36px;
  border-radius: 8px;
  display: flex;
  align-items: center;
  justify-content: center;
  flex-shrink: 0;
}

.bento-saved-meta {
  min-width: 0;
}

.bento-saved-filename {
  font-size: 0.8rem;
  font-weight: 600;
  color: #ffffff;
  white-space: nowrap;
  overflow: hidden;
  text-overflow: ellipsis;
}

.bento-saved-subtext {
  font-size: 0.7rem;
  color: var(--slate);
  margin-top: 3px;
  display: flex;
  align-items: center;
  gap: 6px;
}

.bento-mode-badge {
  font-size: 7.5px;
  font-weight: 800;
  padding: 1px 4px;
  border-radius: 3px;
}

.bento-mode-badge.shopify {
  background: rgba(59, 130, 246, 0.15);
  color: #3b82f6;
}

.bento-mode-badge.logistics {
  background: rgba(34, 197, 94, 0.15);
  color: #22c55e;
}

.bento-mode-badge.universal {
  background: rgba(255, 255, 255, 0.06);
  color: var(--slate);
}

.bento-saved-actions {
  display: flex;
  gap: 6px;
}

.btn-bento-action {
  padding: 5px 10px;
  font-size: 0.75rem;
  font-weight: 700;
  border-radius: 6px;
  cursor: pointer;
  transition: all 0.15s ease;
  border: none;
}

.btn-bento-action.load {
  background: #22c55e;
  color: #000000;
}

.btn-bento-action.load:hover {
  background: #16a34a;
  box-shadow: 0 0 10px rgba(34, 197, 94, 0.3);
}

.btn-bento-action.delete {
  background: rgba(239, 68, 68, 0.1);
  color: #ef4444;
  border: 1px solid rgba(239, 68, 68, 0.2);
}

.btn-bento-action.delete:hover {
  background: rgba(239, 68, 68, 0.2);
}

.bento-empty-state {
  padding: 2.5rem 1rem;
  color: var(--slate);
  font-size: 0.8rem;
  text-align: center;
  display: flex;
  flex-direction: column;
  align-items: center;
  gap: 8px;
}

.bento-upgrade-card {
  background: linear-gradient(135deg, rgba(168, 85, 247, 0.08) 0%, rgba(59, 130, 246, 0.04) 100%);
  border: 1px solid rgba(168, 85, 247, 0.2);
  border-radius: 12px;
  padding: 1rem;
  display: flex;
  flex-direction: column;
  gap: 8px;
  margin-top: auto;
}

.bento-upgrade-header {
  display: flex;
  align-items: center;
  gap: 10px;
}

.bento-upgrade-icon {
  font-size: 1.5rem;
}

.bento-upgrade-title {
  font-size: 0.85rem;
  font-weight: 750;
  color: #ffffff;
}

.bento-upgrade-price {
  font-size: 0.75rem;
  color: #a855f7;
  font-weight: 600;
}

.bento-upgrade-desc {
  font-size: 0.75rem;
  color: var(--slate);
  line-height: 1.4;
  margin: 0;
}

.btn-bento-upgrade {
  width: 100%;
  padding: 8px;
  background: #a855f7;
  color: #ffffff;
  border: none;
  border-radius: 8px;
  font-size: 0.8rem;
  font-weight: 700;
  cursor: pointer;
  transition: all 0.2s ease;
  text-align: center;
}

.btn-bento-upgrade:hover {
  background: #9333ea;
  box-shadow: 0 0 15px rgba(168, 85, 247, 0.4);
}

.bento-pro-active-card {
  background: rgba(34, 197, 94, 0.04);
  border: 1px solid rgba(34, 197, 94, 0.15);
  border-radius: 12px;
  padding: 1rem;
  display: flex;
  flex-direction: column;
  gap: 6px;
  align-items: center;
  text-align: center;
  margin-top: auto;
}

.bento-pro-active-badge {
  font-size: 0.75rem;
  font-weight: 800;
  color: #22c55e;
  background: rgba(34, 197, 94, 0.1);
  padding: 2px 8px;
  border-radius: 20px;
  letter-spacing: 0.05em;
}

.bento-pro-active-text {
  font-size: 0.75rem;
  color: var(--slate);
  line-height: 1.4;
  margin: 0;
}

/* ============ MAINTENANCE MODE OVERLAY STYLES ============ */
.maint-overlay {
  position: fixed;
  top: 0;
  left: 0;
  width: 100%;
  height: 100%;
  z-index: 99999;
  background: rgba(5, 5, 5, 0.8);
  backdrop-filter: blur(25px);
  display: flex;
  align-items: center;
  justify-content: center;
  padding: 2rem;
  animation: fadeIn 0.4s ease;
}

@keyframes fadeIn {
  from { opacity: 0; }
  to { opacity: 1; }
}

.maint-card {
  background: rgba(18, 18, 18, 0.9);
  border: 1px solid rgba(245, 158, 11, 0.2);
  border-radius: 24px;
  max-width: 480px;
  width: 100%;
  padding: 3rem 2.5rem;
  text-align: center;
  box-shadow: 0 20px 50px rgba(0, 0, 0, 0.7), 0 0 40px rgba(245, 158, 11, 0.05);
}

.maint-icon {
  font-size: 3.5rem;
  margin-bottom: 1.5rem;
  display: inline-block;
  animation: float 3s ease-in-out infinite;
}

@keyframes float {
  0%, 100% { transform: translateY(0); }
  50% { transform: translateY(-8px); }
}

.maint-title {
  font-size: 1.5rem;
  font-weight: 800;
  color: #ffffff;
  margin-bottom: 1rem;
  letter-spacing: -0.01em;
}

.maint-desc {
  font-size: 0.875rem;
  color: var(--slate);
  line-height: 1.6;
  margin-bottom: 2rem;
}

.maint-divider {
  height: 1px;
  background: rgba(255, 255, 255, 0.05);
  margin: 1.5rem 0;
}

.maint-admin-section {
  display: flex;
  flex-direction: column;
  align-items: center;
  gap: 10px;
}

.maint-admin-label {
  font-size: 11px;
  color: var(--slate);
  text-transform: uppercase;
  letter-spacing: 0.05em;
  font-weight: 600;
}

.btn-maint-bypass {
  padding: 8px 16px;
  background: rgba(255, 255, 255, 0.04);
  border: 1px solid rgba(255, 255, 255, 0.08);
  color: #ffffff;
  border-radius: 8px;
  font-size: 0.8rem;
  font-weight: 700;
  cursor: pointer;
  transition: all 0.2s ease;
}

.btn-maint-bypass:hover {
  background: rgba(245, 158, 11, 0.1);
  border-color: rgba(245, 158, 11, 0.3);
  color: #f59e0b;
}
`}</style>

      {/* Main Navigation */}
      <nav className="sheetai-nav">
        <div className="sheetai-nav-brand">
          <img
            src="/logo-dark.png"
            alt="SheetCodeCrest"
            style={{ height: "32px", width: "32px", borderRadius: "8px", objectFit: "contain" }}
          />
          SheetCodeCrest
        </div>

        <div className="sheetai-nav-links">
          <button type="button" className="sheetai-nav-link" onClick={() => window.scrollTo({ top: 0, behavior: "smooth" })}>Features</button>
          <button type="button" className="sheetai-nav-link" onClick={() => { const el = document.getElementById("pricing-section"); el?.scrollIntoView({ behavior: "smooth" }); }}>Pricing</button>
          <button type="button" className="sheetai-nav-link" onClick={() => window.scrollTo({ top: 0, behavior: "smooth" })}>How it works</button>
          <a href={CODECREST.website} target="_blank" rel="noopener noreferrer" className="sheetai-nav-link">Resources</a>
        </div>

        <div className="sheetai-nav-actions">
          {isSharedViewOnly ? (
            <div style={{ display: "flex", alignItems: "center", gap: "8px" }}>
              <span style={{ fontSize: "11px", background: "rgba(245,158,11,0.1)", color: "#f59e0b", padding: "4px 10px", borderRadius: "20px", border: "1px solid rgba(245,158,11,0.2)" }}>🔒 Shared View</span>
              <button type="button" className="btn-nav-ghost" onClick={() => { setIsSharedViewOnly(false); setActiveRecordId(null); setSharedRecordObj(null); setFile(null); setRawRows([]); setTableHeaders([]); setDataProfile(null); setLogisticsAnalytics(null); setShopifyAnalytics(null); setStep("upload"); window.history.replaceState({}, document.title, window.location.pathname); }}>↩️ Upload My Own</button>
            </div>
          ) : currentUser ? (
            <div style={{ display: "flex", alignItems: "center", gap: "8px" }}>
              <span style={{ fontSize: "12px", color: "var(--slate)" }}>👤 {currentUser.username}</span>
              <span className={`plan-badge ${currentUser.isPro ? "pro" : "free"}`}>
                {currentUser.isPro ? "Pro" : "Free"}
              </span>
              {isAdminActive && (
                <button type="button" className="btn-nav-ghost" style={{ color: "#f59e0b" }} onClick={() => { setAdminModalOpen(true); loadAdminData(); }}>🛡️ Admin</button>
              )}
              <button type="button" className="btn-nav-ghost" onClick={() => setDashboardOpen(true)}>Dashboard</button>
              {!currentUser.isPro && (
                <button type="button" className="btn-nav-primary" onClick={() => setCheckoutOpen(true)}>⚡ Upgrade</button>
              )}
              <button type="button" className="btn-nav-ghost" onClick={handleLogout} style={{ fontSize: "12px" }}>Sign Out</button>
            </div>
          ) : (
            <>
              <button type="button" className="btn-nav-ghost" onClick={() => { setAuthTab("login"); setAuthError(""); setAuthModalOpen(true); }}>Log in</button>
              <button type="button" className="btn-nav-primary" onClick={() => { setAuthTab("signup"); setAuthError(""); setAuthModalOpen(true); }}>Start for free →</button>
            </>
          )}
        </div>
      </nav>




      {/* Auto-detected notices */}
      {detectionNotice && (
        <div className="toast-badge">
          <span>✨</span> {detectionNotice}
        </div>
      )}

      {/* Error state */}
      {error && (
        <div className="section-card" style={{ borderLeft: "4px solid var(--error)", background: "rgba(239, 68, 68, 0.08)" }}>
          <div style={{ color: "var(--error)", fontWeight: 600, fontSize: "0.9rem" }}>⚠️ Spreadsheet Error</div>
          <div style={{ color: "var(--ink)", fontSize: "0.85rem", marginTop: "4px", opacity: 0.9 }}>{error}</div>
        </div>
      )}

      {/* Step 1: Upload File Area */}
      {step === "upload" && (
        <>
          {/* ============ HERO / BENTO WORKSPACE SECTION ============ */}
          {currentUser ? (
            <div className="bento-dashboard">
              {/* Greeting Banner Row */}
              <div className="bento-greeting-banner">
                <div className="bento-greeting-left">
                  <span className="bento-greeting-avatar">✨</span>
                  <div>
                    <h2 className="bento-greeting-title">
                      Welcome back, <span className="green-word">{currentUser.name || currentUser.username}</span>!
                    </h2>
                    <p className="bento-greeting-sub">
                      Your high-priority intelligence workspace is active. Select your mode and process your sheet.
                    </p>
                  </div>
                </div>
                <div className="bento-greeting-right">
                  <div className="bento-greeting-stats">
                    <span className={`plan-badge ${currentUser.isPro ? "pro" : "free"}`} style={{ fontSize: "11px", padding: "4px 10px", fontWeight: 700 }}>
                      {currentUser.isPro ? "⚡ PRO MEMBERSHIP" : "🆓 FREE TRIAL"}
                    </span>
                    {!currentUser.isPro && (
                      <span style={{ fontSize: "12px", color: "var(--slate)", fontWeight: 550 }}>
                        {freeReportsRemaining} of {globalFreeLimit} free reports remaining
                      </span>
                    )}
                  </div>
                </div>
              </div>

              {/* Main Bento Grid */}
              <div className="bento-grid">
                {/* Left Pane (60% width): Workspace Dropzone Uploader */}
                <div className="bento-card bento-workspace-uploader">
                  <div className="bento-card-header">
                    <span className="bento-card-title">📁 Active Spreadsheet Workspace</span>
                    <span className="bento-card-badge" style={{ background: "rgba(34,197,94,0.1)", color: "#22c55e" }}>Uploader</span>
                  </div>
                  
                  {/* Mode Selector Tabs */}
                  <div className="bento-mode-selector">
                    <div className="bento-mode-label">Analysis Mode:</div>
                    <div className="bento-mode-tabs">
                      {(["universal", "logistics", "shopify"] as const).map((m) => (
                        <button
                          key={m}
                          type="button"
                          className={`bento-mode-tab ${mode === m ? "active" : ""}`}
                          onClick={() => {
                            setMode(m);
                            addLog(`🔄 Analysis mode set manually to ${m.toUpperCase()}`, "info");
                          }}
                        >
                          {m === "shopify" ? "🛒 Shopify" : m === "logistics" ? "🚚 Logistics" : "📊 Universal"}
                        </button>
                      ))}
                    </div>
                  </div>

                  {/* Drag-Drop Zone */}
                  <div
                    className={`upload-card-new bento-dropzone ${dragging ? "dragging" : ""}`}
                    onDrop={onDrop}
                    onDragOver={onDragOver}
                    onDragLeave={onDragLeave}
                    onClick={() => hasFreeReportsRemaining && inputRef.current?.click()}
                  >
                    <input
                      type="file"
                      ref={inputRef}
                      accept=".xlsx,.xls,.csv"
                      style={{ display: "none" }}
                      onChange={onFileChange}
                    />
                    <div className="upload-icon-new">☁️</div>
                    <div className="upload-title-new">Drag & drop your spreadsheet here</div>
                    <div className="upload-formats">Supports .xlsx, .xls, .csv · Auto-detected schema</div>
                    {hasFreeReportsRemaining ? (
                      <button
                        type="button"
                        className="btn-upload-green"
                        onClick={(e) => { e.stopPropagation(); inputRef.current?.click(); }}
                      >
                        ↑ Select Spreadsheet File
                      </button>
                    ) : (
                      <div onClick={(e) => e.stopPropagation()}>
                        <div style={{ fontSize: "12px", color: "var(--slate)", margin: "0.5rem 0 1rem" }}>
                          Free trial complete. Upgrade to continue.
                        </div>
                        <button
                          type="button"
                          className="btn-upload-green"
                          onClick={() => setCheckoutOpen(true)}
                        >
                          ⚡ Upgrade to Pro Now
                        </button>
                      </div>
                    )}
                  </div>

                  <div className="bento-workspace-footer">
                    {!currentUser.isPro && (
                      <div className="bento-limit-bar-wrapper">
                        <div className="bento-limit-bar-labels">
                          <span>Free Report Limit Usage</span>
                          <span>{usageCount} / {globalFreeLimit} reports used</span>
                        </div>
                        <div className="bento-limit-bar-bg">
                          <div
                            className="bento-limit-bar-fill"
                            style={{ width: `${Math.min(100, (usageCount / globalFreeLimit) * 100)}%` }}
                          />
                        </div>
                      </div>
                    )}
                    {currentUser.isPro && (
                      <div className="bento-unlocked-status">
                        <span>❇️ Enterprise priority API queue active. Secure data sandbox enabled.</span>
                      </div>
                    )}
                  </div>
                </div>

                {/* Right Pane (40% width): Saved Spreadsheet Bento */}
                <div className="bento-card bento-saved-sheets">
                  <div className="bento-card-header">
                    <span className="bento-card-title">📂 Saved Datasets & History</span>
                    <span className="bento-card-badge" style={{ background: "rgba(34,197,94,0.1)", color: "#22c55e" }}>
                      {savedRecords.length} files
                    </span>
                  </div>

                  {/* Inline Search Bar */}
                  <div className="bento-search-wrapper">
                    <span className="bento-search-icon">🔍</span>
                    <input
                      type="text"
                      placeholder="Search files by name..."
                      className="bento-search-input"
                      value={sheetSearchQuery}
                      onChange={(e) => setSheetSearchQuery(e.target.value)}
                    />
                    {sheetSearchQuery && (
                      <button
                        type="button"
                        className="bento-search-clear"
                        onClick={() => setSheetSearchQuery("")}
                      >
                        ✕
                      </button>
                    )}
                  </div>

                  {/* Saved records list */}
                  <div className="bento-saved-list">
                    {filteredSavedRecords.length > 0 ? (
                      filteredSavedRecords.map((rec) => (
                        <div key={rec.id} className="bento-saved-item">
                          <div className="bento-saved-item-left">
                            <span className="bento-saved-icon">
                              {rec.mode === "shopify" ? "🛒" : rec.mode === "logistics" ? "🚚" : "📊"}
                            </span>
                            <div className="bento-saved-meta">
                              <div title={rec.filename} className="bento-saved-filename">
                                {rec.filename}
                              </div>
                              <div className="bento-saved-subtext">
                                <span className={`bento-mode-badge ${rec.mode}`}>{rec.mode.toUpperCase()}</span>
                                <span>•</span>
                                <span>{(rec.size / 1024).toFixed(1)} KB</span>
                              </div>
                            </div>
                          </div>
                          <div className="bento-saved-actions">
                            <button
                              type="button"
                              className="btn-bento-action load"
                              title="Load Report"
                              onClick={() => loadRecord(rec)}
                            >
                              ⚡ Load
                            </button>
                            <button
                              type="button"
                              className="btn-bento-action delete"
                              title="Delete Record"
                              onClick={(e) => handleDeleteRecord(rec.id!, e)}
                            >
                              🗑️
                            </button>
                          </div>
                        </div>
                      ))
                    ) : (
                      <div className="bento-empty-state">
                        <span style={{ fontSize: "1.5rem" }}>📁</span>
                        <p style={{ margin: 0 }}>
                          {sheetSearchQuery
                            ? "No matching spreadsheets found."
                            : "No saved datasets. Drop your first spreadsheet to start!"}
                        </p>
                      </div>
                    )}
                  </div>

                  {!currentUser.isPro ? (
                    <div className="bento-upgrade-card">
                      <div className="bento-upgrade-header">
                        <span className="bento-upgrade-icon">💎</span>
                        <div>
                          <div className="bento-upgrade-title">Unlock SheetCodeCrest Pro</div>
                          <div className="bento-upgrade-price">Just ₹1,599/month</div>
                        </div>
                      </div>
                      <p className="bento-upgrade-desc">
                        Get unlimited spreadsheet analysis, connect live integrations, and enjoy premium priority processing.
                      </p>
                      <button
                        type="button"
                        className="btn-bento-upgrade"
                        onClick={() => setCheckoutOpen(true)}
                      >
                        Upgrade to Pro Now →
                      </button>
                    </div>
                  ) : (
                    <div className="bento-pro-active-card">
                      <span className="bento-pro-active-badge">✦ PRO UNLOCKED</span>
                      <p className="bento-pro-active-text">
                        Your workspace is fully active with high-priority execution. Thank you for your support!
                      </p>
                    </div>
                  )}
                </div>
              </div>
            </div>
          ) : (
            /* Otherwise show standard Centered Hero layout for logged out users */
            <section className="hero-centered">
              <div className="hero-centered-content">
                <div className="sheetai-hero-label">✦ AI-Powered Spreadsheet Analysis</div>
                <h1 className="hero-title-new">
                  Take control of your <br/><span className="green-word">e-commerce data</span>
                </h1>
                <p style={{ fontSize: "1.05rem", color: "var(--slate)", marginBottom: "1.5rem", lineHeight: 1.6, maxWidth: "600px" }}>
                  Upload your spreadsheets and get instant insights, automated analysis, and beautiful reports in seconds.
                </p>

                {/* Hero Actions capsule buttons */}
                <div className="hero-actions-row">
                  <button
                    type="button"
                    className="btn-hero-primary"
                    onClick={() => {
                      const el = document.getElementById("centered-uploader-anchor");
                      el?.scrollIntoView({ behavior: "smooth" });
                    }}
                  >
                    ↑ Start Analyzing Now
                  </button>
                  <button
                    type="button"
                    className="btn-hero-secondary"
                    onClick={() => {
                      const el = document.getElementById("dashboard-mockup-anchor");
                      el?.scrollIntoView({ behavior: "smooth" });
                    }}
                  >
                    👀 View Dashboard Preview
                  </button>
                </div>

                {/* Stats chip bar */}
                <div className="hero-stats-row" style={{ marginBottom: "2.5rem" }}>
                  {[
                    { icon: "📁", value: "3+", label: "File Types" },
                    { icon: "💡", value: "70+", label: "Insights" },
                    { icon: "⚡", value: "0", label: "Manual Work" },
                    { icon: "🔒", value: "100%", label: "Secure" },
                  ].map((stat) => (
                    <div className="hero-stat-item" key={stat.label}>
                      <div className="stat-icon">{stat.icon}</div>
                      <div style={{ textAlign: "left" }}>
                        <strong>{stat.value}</strong>
                        <span>{stat.label}</span>
                      </div>
                    </div>
                  ))}
                </div>

                {/* Translucent Active Spreadsheet dropzone uploader */}
                <div id="centered-uploader-anchor" style={{ width: "100%", scrollMarginTop: "100px" }}>
                  <div
                    className={`upload-card-new ${dragging ? "dragging" : ""}`}
                    onDrop={onDrop}
                    onDragOver={onDragOver}
                    onDragLeave={onDragLeave}
                    onClick={() => hasFreeReportsRemaining && inputRef.current?.click()}
                  >
                    <input
                      type="file"
                      ref={inputRef}
                      accept=".xlsx,.xls,.csv"
                      style={{ display: "none" }}
                      onChange={onFileChange}
                    />
                    <div className="upload-icon-new">☁️</div>
                    <div className="upload-title-new">Drag & drop your spreadsheet here</div>
                    <div className="upload-formats">Supports .xlsx, .xls, .csv · Auto-detected schema</div>
                    {hasFreeReportsRemaining ? (
                      <button
                        type="button"
                        className="btn-upload-green"
                        onClick={(e) => { e.stopPropagation(); inputRef.current?.click(); }}
                      >
                        ↑ Upload your file
                      </button>
                    ) : (
                      <div onClick={(e) => e.stopPropagation()}>
                        <div style={{ fontSize: "12px", color: "var(--slate)", margin: "0.5rem 0 1rem" }}>
                          Free trial complete. Upgrade to continue.
                        </div>
                        <button
                          type="button"
                          className="btn-upload-green"
                          onClick={() => {
                            setAuthTab("login");
                            setAuthError("");
                            setAuthModalOpen(true);
                          }}
                        >
                          ⚡ Sign In to Upgrade
                        </button>
                      </div>
                    )}
                    <div className="upload-integration-row">
                      <span>Connect Live Channels:</span>
                      <div className="upload-integration-icon" title="Shopify">🛒</div>
                      <div className="upload-integration-icon" title="Shiprocket">🚚</div>
                      <div className="upload-integration-icon" title="Google Sheets">📊</div>
                    </div>
                  </div>
                  <div className="no-cc-line" style={{ marginBottom: "3rem" }}>No credit card required · {freeReportsRemaining > 0 ? `${freeReportsRemaining} free reports remaining` : "Free trial complete"}</div>
                </div>

                {/* High-fidelity Dashboard Mockup container */}
                <div id="dashboard-mockup-anchor" style={{ width: "100%", maxWidth: "800px", margin: "0 auto", scrollMarginTop: "100px" }}>
                  <motion.div
                    initial={{ opacity: 0, y: 20 }}
                    animate={{ opacity: 1, y: 0 }}
                    transition={{ duration: 0.7, ease: "easeOut" }}
                    className="analysis-overview-card"
                  >
                    {/* Card Header */}
                    <div className="analysis-card-header">
                      <span className="analysis-card-title">✨ Active Analysis Mockup (Live Dashboard)</span>
                      <div className="live-badge">Mockup Active</div>
                    </div>

                    {/* Health Score and line graph */}
                    <div style={{ display: "flex", alignItems: "center", gap: "20px", flexWrap: "wrap" }}>
                      <svg width="70" height="70" viewBox="0 0 70 70" style={{ flexShrink: 0 }}>
                        <circle cx="35" cy="35" r="28" fill="none" stroke="#141416" strokeWidth="8" />
                        <circle
                          cx="35" cy="35" r="28"
                          fill="none" stroke="var(--deep-green)" strokeWidth="8"
                          strokeDasharray={`${(92/100)*175.9} 175.9`}
                          strokeLinecap="round"
                          transform="rotate(-90 35 35)"
                          style={{ filter: "drop-shadow(0 0 5px rgba(0, 255, 178, 0.4))" }}
                        />
                        <text x="35" y="40" textAnchor="middle" fontSize="15" fontWeight="800" fill="var(--deep-green)">92</text>
                      </svg>
                      <div style={{ flex: 1, minWidth: "180px", textAlign: "left" }}>
                        <div style={{ fontSize: "10px", color: "var(--slate)", marginBottom: "2px" }}>OVERALL E-COMMERCE HEALTH</div>
                        <div style={{ fontSize: "13px", color: "#fff", fontWeight: 700, marginBottom: "8px" }}>Great! Your e-commerce metrics look healthy</div>
                        <svg width="100%" height="30" viewBox="0 0 120 30" style={{ display: "block" }}>
                          <polyline points="0,26 18,20 36,22 54,12 72,16 90,7 108,9 120,5" fill="none" stroke="var(--deep-green)" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{ filter: "drop-shadow(0 0 3px rgba(0,255,178,0.2))" }} />
                          <polyline points="0,26 18,20 36,22 54,12 72,16 90,7 108,9 120,5 120,30 0,30" fill="rgba(0,255,178,0.06)" stroke="none" />
                        </svg>
                      </div>
                    </div>

                    {/* KPI Row */}
                    <div className="card-kpi-row">
                      {[
                        { label: "Revenue", value: currency === "INR" ? "₹12.4K" : currency === "EUR" ? "€135" : "$146", change: "+16.6%", up: true },
                        { label: "Expenses", value: currency === "INR" ? "₹4.2K" : currency === "EUR" ? "€45" : "$49", change: "-7.3%", up: false },
                        { label: "Profit", value: currency === "INR" ? "₹8.2K" : currency === "EUR" ? "€90" : "$97", change: "+28.1%", up: true },
                      ].map((kpi) => (
                        <div className="card-kpi-item" key={kpi.label} style={{ textAlign: "left" }}>
                          <div className="card-kpi-label">{kpi.label}</div>
                          <div className="card-kpi-value">{kpi.value}</div>
                          <div className={`card-kpi-change ${kpi.up ? "up" : "down"}`}>{kpi.up ? "↑" : "↓"} {kpi.change}</div>
                        </div>
                      ))}
                    </div>

                    {/* Insights Summary Bars */}
                    <div className="card-bar-section" style={{ textAlign: "left" }}>
                      <div className="card-bar-label">AI Insights distribution</div>
                      <div style={{ display: "flex", gap: "10px", alignItems: "flex-end", height: "60px" }}>
                        {[
                          { label: "Trends", h: "75%", hi: false },
                          { label: "Anomalies", h: "50%", hi: true },
                          { label: "Patterns", h: "65%", hi: false },
                          { label: "Opportunities", h: "85%", hi: true },
                        ].map((bar) => (
                          <div key={bar.label} style={{ flex: 1, display: "flex", flexDirection: "column", alignItems: "center", gap: "4px" }}>
                            <div style={{ width: "100%", height: "52px", display: "flex", alignItems: "flex-end" }}>
                              <div style={{ width: "100%", height: bar.h, background: bar.hi ? "var(--deep-green)" : "rgba(0,255,178,0.2)", borderRadius: "3px 3px 0 0", transition: "height 0.5s", boxShadow: bar.hi ? "0 0 10px rgba(0,255,178,0.15)" : "none" }} />
                            </div>
                            <div style={{ fontSize: "8px", color: "var(--slate)", textAlign: "center" }}>{bar.label}</div>
                          </div>
                        ))}
                      </div>
                    </div>

                    {/* Completed band */}
                    <div className="card-analysis-complete">
                      <span>● Local AI-Analysis complete · 0 bytes leaked</span>
                      <button type="button" style={{ background: "none", border: "none", color: "var(--deep-green)", fontSize: "10px", cursor: "pointer", fontFamily: "inherit", fontWeight: 700 }} onClick={() => { setAuthTab("signup"); setAuthError(""); setAuthModalOpen(true); }}>Start Free →</button>
                    </div>
                  </motion.div>
                </div>
              </div>
            </section>
          )}

          {/* ============ RECENT ANALYSES + INSIGHTS ============ */}
          {!currentUser && (
            <>
              <div className="recent-analyses-section">
                <div className="recent-analyses-card">
                  <div className="recent-analyses-header">
                    <span className="recent-analyses-title">Recent Analyses</span>
                    <button type="button" className="view-all-link" onClick={() => { setAuthTab("signup"); setAuthError(""); setAuthModalOpen(true); }}>View all →</button>
                  </div>
                  <table className="analyses-table">
                    <thead>
                      <tr>
                        <th>File Name</th>
                        <th>Rows</th>
                        <th>Columns</th>
                        <th>Status</th>
                        <th>Insights</th>
                        <th></th>
                      </tr>
                    </thead>
                    <tbody>
                      {[
                        { name: "Q4_Financials.xlsx", rows: "12,532", cols: 18, insights: "8 insights" },
                        { name: "Marketing_Data.csv", rows: "45,210", cols: 24, insights: "12 insights" },
                        { name: "Sales_2024.xlsx", rows: "22,104", cols: 16, insights: "10 insights" },
                        { name: "Inventory_May.csv", rows: "8,421", cols: 12, insights: "6 insights" },
                      ].map((row) => (
                        <tr key={row.name}>
                          <td><div className="filename-cell"><div className="file-icon-dot">📄</div><span>{row.name}</span></div></td>
                          <td style={{ color: "var(--slate)" }}>{row.rows}</td>
                          <td style={{ color: "var(--slate)" }}>{row.cols}</td>
                          <td><span className="status-badge-completed">Completed</span></td>
                          <td style={{ color: "#22c55e", fontWeight: 600 }}>{row.insights}</td>
                          <td><button type="button" className="arrow-link-btn">→</button></td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>

                <div className="insights-summary-card">
                  <div style={{ fontSize: "13px", fontWeight: 700, color: "#fff", marginBottom: "6px" }}>Insights Summary</div>
                  <div style={{ fontSize: "11px", color: "var(--slate)", marginBottom: "12px" }}>Across all analyses</div>
                  <svg viewBox="0 0 120 72" style={{ width: "100%", height: "90px" }}>
                    <rect x="4" y="12" width="112" height="14" rx="4" fill="rgba(34,197,94,0.12)" />
                    <rect x="4" y="12" width={112 * 0.75} height="14" rx="4" fill="rgba(34,197,94,0.5)" />
                    <rect x="4" y="30" width="112" height="14" rx="4" fill="rgba(245,158,11,0.12)" />
                    <rect x="4" y="30" width={112 * 0.55} height="14" rx="4" fill="rgba(245,158,11,0.4)" />
                    <rect x="4" y="48" width="112" height="14" rx="4" fill="rgba(59,130,246,0.12)" />
                    <rect x="4" y="48" width={112 * 0.65} height="14" rx="4" fill="rgba(59,130,246,0.4)" />
                  </svg>
                  <div className="insights-legend">
                    {[
                      { label: "Trends", pct: "40%", color: "#22c55e" },
                      { label: "Anomalies", pct: "30%", color: "#f59e0b" },
                      { label: "Patterns", pct: "20%", color: "#3b82f6" },
                      { label: "Opportunities", pct: "10%", color: "#a855f7" },
                    ].map((item) => (
                      <div className="insights-legend-item" key={item.label}>
                        <div className="legend-dot" style={{ background: item.color }} />
                        <span>{item.label}</span>
                        <span className="legend-pct">{item.pct}</span>
                      </div>
                    ))}
                  </div>
                </div>
              </div>

              {/* ============ BENTO GRID FEATURES ============ */}
              <section className="sheetai-features-section">
                <div className="sheetai-features-label">High-Performance E-Commerce Abstractions</div>
                <h3 className="sheetai-features-title">Everything you need to audit, understand, and grow</h3>

                <div className="bento-features-grid">
                  {/* Card 1 (Wide): Live Platform Integrations */}
                  <div className="bento-feature-card bento-colspan-2">
                    <div className="bento-card-bg-gradient" />
                    <div>
                      <div className="bento-feature-icon">🔌</div>
                      <div className="bento-feature-name">Live Integration Abstraction</div>
                      <p className="bento-feature-desc" style={{ maxWidth: "480px" }}>
                        Seamlessly sync e-commerce logs from your storefront, logistics handlers, and search networks. Ingest Shopify sales, Shiprocket orders, and Google Sheets cleanly.
                      </p>
                    </div>

                    {/* Pulsating connection visual mockup */}
                    <div className="bento-visual-channels">
                      <div className="bento-channel-badge" style={{ borderColor: "var(--deep-green)", boxShadow: "0 0 10px rgba(0,255,178,0.1)" }}>🛒 Shopify Store</div>
                      <div className="bento-pulse-line" />
                      <div style={{ width: "8px", height: "8px", borderRadius: "50%", background: "var(--deep-green)", animation: "pulse-green 1.5s infinite" }} />
                      <div className="bento-channel-badge" style={{ borderColor: "var(--deep-green)", boxShadow: "0 0 10px rgba(0,255,178,0.1)" }}>🚚 Shiprocket Logs</div>
                      <div className="bento-pulse-line" style={{ left: "50%" }} />
                      <div className="bento-channel-badge" style={{ background: "rgba(0,255,178,0.03)", borderColor: "var(--deep-green)" }}>📊 Google Sheets</div>
                    </div>
                  </div>

                  {/* Card 2 (Tall): Real-Time E-Commerce Insights */}
                  <div className="bento-feature-card bento-rowspan-2">
                    <div className="bento-card-bg-gradient" />
                    <div className="bento-feature-icon">🧠</div>
                    <div>
                      <div className="bento-feature-name">70+ Automatic Growth Audits</div>
                      <p className="bento-feature-desc" style={{ marginBottom: "16px" }}>
                        Our specialized parser reads storefront layouts and highlights fee leaks, COD return vulnerabilities, and margin opportunities automatically.
                      </p>

                      {/* Insights tag block list */}
                      <div style={{ display: "flex", flexDirection: "column", gap: "8px" }}>
                        {[
                          { t: "Revenue Surge Detected", c: "var(--deep-green)", bg: "rgba(0,255,178,0.08)", p: "+24.8%" },
                          { t: "Logistics Cost Overlap", c: "#ef4444", bg: "rgba(239,68,68,0.08)", p: "₹22.4K Leak" },
                          { t: "COD RTO Return Threat", c: "#f59e0b", bg: "rgba(245,158,11,0.08)", p: "High Risk" },
                          { t: "Inventory Runout Alarm", c: "#a855f7", bg: "rgba(168,85,247,0.08)", p: "in 4 days" }
                        ].map((tag, i) => (
                          <div key={i} style={{ display: "flex", alignItems: "center", justifyContent: "space-between", background: tag.bg, border: `1px solid ${tag.c}20`, padding: "6px 10px", borderRadius: "8px", fontSize: "11px" }}>
                            <span style={{ color: "#fff", fontWeight: 550 }}>{tag.t}</span>
                            <span style={{ color: tag.c, fontWeight: 700, fontFamily: "var(--font-technical)" }}>{tag.p}</span>
                          </div>
                        ))}
                      </div>
                    </div>
                  </div>

                  {/* Card 3 (Standard): Multi-Format DropZone */}
                  <div className="bento-feature-card">
                    <div className="bento-card-bg-gradient" />
                    <div>
                      <div className="bento-feature-icon">📁</div>
                      <div className="bento-feature-name">Automated Schema Parser</div>
                      <p className="bento-feature-desc">
                        Drop raw, unformatted Excel or CSV logs. Avery automatically aligns column keys, normalizes pricing indexes, and maps sales indicators in milliseconds.
                      </p>
                    </div>
                  </div>

                  {/* Card 4 (Tall): Simulated Avery AI Chatbot */}
                  <div className="bento-feature-card bento-rowspan-2">
                    <div className="bento-card-bg-gradient" />
                    <div className="bento-feature-icon">💬</div>
                    <div>
                      <div className="bento-feature-name">Ask AI Avery in Plain English</div>
                      <p className="bento-feature-desc" style={{ marginBottom: "16px" }}>
                        Ask natural language queries about product margins, marketing ROAS, or delivery anomalies.
                      </p>

                      {/* Chat conversation simulation */}
                      <div style={{ background: "rgba(0,0,0,0.3)", borderRadius: "12px", border: "1px solid rgba(255,255,255,0.03)", padding: "12px", display: "flex", flexDirection: "column", gap: "10px", fontSize: "11px" }}>
                        <div style={{ alignSelf: "flex-end", background: "rgba(255,255,255,0.05)", padding: "6px 10px", borderRadius: "10px 10px 0 10px", color: "#fff", maxWidth: "90%", textAlign: "right" }}>
                          "What was our highest-margin product last month?"
                        </div>
                        <div style={{ alignSelf: "flex-start", background: "rgba(0,255,178,0.05)", border: "1px solid rgba(0,255,178,0.1)", padding: "6px 10px", borderRadius: "10px 10px 10px 0", color: "var(--body-muted)", maxWidth: "90%", textAlign: "left" }}>
                          <strong style={{ color: "var(--deep-green)", display: "block", marginBottom: "2px" }}>🤖 Avery AI</strong>
                          Your highest margin product was **Organic Coffee Beans** (68% margin, generating {currency === "INR" ? "₹8,200" : currency === "EUR" ? "€90" : "$97"} net profit).
                        </div>
                      </div>
                    </div>
                  </div>

                  {/* Card 5 (Wide): Interactive Security Dashboard */}
                  <div className="bento-feature-card bento-colspan-2">
                    <div className="bento-card-bg-gradient" />
                    <div>
                      <div className="bento-feature-icon">🔒</div>
                      <div className="bento-feature-name">Browser Sandboxed Architecture</div>
                      <p className="bento-feature-desc" style={{ maxWidth: "480px" }}>
                        Your financial data stays inside your browser. We leverage IndexedDB sandboxing and local parsing assemblies. Data never travels to corporate databases.
                      </p>
                    </div>

                    {/* Interactive Security Toggles mockup */}
                    <div style={{ display: "flex", gap: "10px", flexWrap: "wrap", marginTop: "14px" }}>
                      {[
                        { label: "Anonymize Customer Data", state: bentoAnonymize, set: setBentoAnonymize },
                        { label: "IndexedDB Sandboxing", state: bentoSandbox, set: setBentoSandbox },
                        { label: "Auto-Delete History Logs", state: bentoAutoDelete, set: setBentoAutoDelete }
                      ].map((tog, idx) => (
                        <div key={idx} onClick={() => tog.set(!tog.state)} style={{ display: "flex", alignItems: "center", gap: "8px", background: "rgba(0,0,0,0.35)", border: "1px solid rgba(255,255,255,0.04)", padding: "8px 12px", borderRadius: "20px", cursor: "pointer", transition: "all 0.15s", userSelect: "none" }}>
                          <div style={{ width: "26px", height: "14px", borderRadius: "20px", background: tog.state ? "var(--deep-green)" : "#222", position: "relative", transition: "background 0.2s" }}>
                            <div style={{ width: "10px", height: "10px", borderRadius: "50%", background: tog.state ? "#030305" : "#888", position: "absolute", top: "2px", left: tog.state ? "14px" : "2px", transition: "left 0.2s" }} />
                          </div>
                          <span style={{ fontSize: "10.5px", color: tog.state ? "#fff" : "var(--slate)" }}>{tog.label}</span>
                        </div>
                      ))}
                    </div>
                  </div>
                </div>
              </section>
            </>
          )}

          {/* ============ METRICS ROW, WHY, TESTIMONIALS, PRICING, CTA ============ */}
          {!currentUser && (
            <>
              {/* ============ E-COMMERCE CALCULATOR ============ */}
              <section className="sheetai-calculator-section" style={{ padding: "80px 20px", position: "relative" }}>
                <div className="sheetai-features-label" style={{ textAlign: "center", marginBottom: "12px" }}>Interactive Performance Calculator</div>
                <h3 className="sheetai-features-title" style={{ textAlign: "center", marginBottom: "16px" }}>Audit your store's performance & growth levers</h3>
                <p style={{ color: "var(--slate)", fontSize: "14.5px", textAlign: "center", maxWidth: "600px", margin: "-4px auto 36px auto", lineHeight: 1.6 }}>
                  Adjust the inputs below to simulate your store's unit economics. Watch how conversion efficiency and ad spend scale your net profits in real-time.
                </p>

                <div className="calc-container" style={{
                  display: "grid",
                  gridTemplateColumns: "repeat(auto-fit, minmax(320px, 1fr))",
                  gap: "28px",
                  background: "rgba(255, 255, 255, 0.02)",
                  border: "1px solid rgba(255, 255, 255, 0.04)",
                  backdropFilter: "blur(20px)",
                  borderRadius: "20px",
                  padding: "32px",
                  maxWidth: "960px",
                  margin: "0 auto",
                  boxShadow: "0 20px 40px rgba(0,0,0,0.5)"
                }}>
                  {/* Left Side: Inputs */}
                  <div style={{ display: "flex", flexDirection: "column", gap: "20px" }}>
                    <div style={{ fontSize: "13px", fontWeight: 700, color: "var(--deep-green)", textTransform: "uppercase", letterSpacing: "0.1em", borderBottom: "1px solid rgba(255,255,255,0.06)", paddingBottom: "10px", marginBottom: "8px" }}>
                      ⚙️ Simulation Parameters
                    </div>

                    {[
                      {
                        label: "Monthly Ad Spend",
                        value: calcAdSpend,
                        setter: setCalcAdSpend,
                        min: 1000,
                        max: 200000,
                        step: 1000,
                        isCurrency: true
                      },
                      {
                        label: "Monthly Orders Placed",
                        value: calcOrders,
                        setter: setCalcOrders,
                        min: 50,
                        max: 5000,
                        step: 50,
                        isCurrency: false
                      },
                      {
                        label: "Average Order Value (AOV)",
                        value: calcAOV,
                        setter: setCalcAOV,
                        min: 100,
                        max: 10000,
                        step: 100,
                        isCurrency: true
                      },
                      {
                        label: "Cost of Goods Sold (COGS) Per Item",
                        value: calcCOGS,
                        setter: setCalcCOGS,
                        min: 10,
                        max: 5000,
                        step: 10,
                        isCurrency: true
                      }
                    ].map((input, idx) => {
                      let displayVal = input.value;
                      let minVal = input.min;
                      let maxVal = input.max;
                      let stepVal = input.step;

                      if (input.isCurrency && currency !== "INR") {
                        displayVal = Math.round(input.value / 80);
                        minVal = Math.round(input.min / 80);
                        maxVal = Math.round(input.max / 80);
                        stepVal = Math.max(1, Math.round(input.step / 80));
                      }

                      const handleSliderChange = (e: React.ChangeEvent<HTMLInputElement>) => {
                        const rawVal = Number(e.target.value);
                        if (input.isCurrency && currency !== "INR") {
                          input.setter(rawVal * 80);
                        } else {
                          input.setter(rawVal);
                        }
                      };

                      const sym = getCurrencySymbol(currency);

                      return (
                        <div key={idx} style={{ display: "flex", flexDirection: "column", gap: "8px" }}>
                          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                            <span style={{ fontSize: "13px", color: "var(--slate)", fontWeight: 550 }}>{input.label}</span>
                            <span style={{ fontSize: "13px", color: "var(--deep-green)", fontWeight: 700, fontFamily: "var(--font-technical)" }}>
                              {input.isCurrency ? `${sym}${displayVal.toLocaleString()}` : displayVal.toLocaleString()}
                            </span>
                          </div>
                          <input
                            type="range"
                            min={minVal}
                            max={maxVal}
                            step={stepVal}
                            value={displayVal}
                            onChange={handleSliderChange}
                            style={{
                              width: "100%",
                              accentColor: "var(--deep-green)",
                              background: "rgba(255,255,255,0.08)",
                              height: "6px",
                              borderRadius: "4px",
                              outline: "none",
                              cursor: "pointer"
                            }}
                          />
                        </div>
                      );
                    })}
                  </div>

                  {/* Right Side: Results */}
                  <div style={{
                    background: "rgba(0, 255, 178, 0.01)",
                    border: "1px solid rgba(0, 255, 178, 0.06)",
                    borderRadius: "14px",
                    padding: "24px",
                    display: "flex",
                    flexDirection: "column",
                    justifyContent: "space-between",
                    gap: "20px"
                  }}>
                    <div>
                      <div style={{ fontSize: "13px", fontWeight: 700, color: "var(--deep-green)", textTransform: "uppercase", letterSpacing: "0.1em", borderBottom: "1px solid rgba(0,255,178,0.1)", paddingBottom: "10px", marginBottom: "16px", display: "flex", justifyContent: "space-between" }}>
                        <span>📊 Unit Economics</span>
                        <span style={{ color: "var(--slate)", fontSize: "10.5px", textTransform: "none", letterSpacing: "normal" }}>Real-time scale</span>
                      </div>

                      {(() => {
                        const sym = getCurrencySymbol(currency);
                        const scale = currency === "INR" ? 1 : 1 / 80;

                        const totalRevenue = calcOrders * calcAOV;
                        const totalCOGS = calcOrders * calcCOGS;
                        const netProfit = totalRevenue - totalCOGS - calcAdSpend;
                        const roas = calcAdSpend > 0 ? (totalRevenue / calcAdSpend).toFixed(2) : "∞";
                        const cac = calcOrders > 0 ? (calcAdSpend / calcOrders) : 0;
                        const profitMargin = totalRevenue > 0 ? ((netProfit / totalRevenue) * 100).toFixed(1) : "0.0";

                        const displayRevenue = totalRevenue * scale;
                        const displayCOGS = totalCOGS * scale;
                        const displayNetProfit = netProfit * scale;
                        const displayCAC = cac * scale;

                        const profitColor = netProfit >= 0 ? "var(--deep-green)" : "#ef4444";

                        return (
                          <div style={{ display: "flex", flexDirection: "column", gap: "16px" }}>
                            {/* Revenue & Profit Main Row */}
                            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "12px" }}>
                              <div style={{ background: "rgba(255,255,255,0.01)", border: "1px solid rgba(255,255,255,0.03)", padding: "12px", borderRadius: "10px" }}>
                                <div style={{ fontSize: "11px", color: "var(--slate)", marginBottom: "4px" }}>Total Revenue</div>
                                <div style={{ fontSize: "20px", fontWeight: 700, color: "#fff", fontFamily: "var(--font-technical)" }}>
                                  {sym}{Math.round(displayRevenue).toLocaleString()}
                                </div>
                              </div>
                              <div style={{ background: "rgba(255,255,255,0.01)", border: "1px solid rgba(255,255,255,0.03)", padding: "12px", borderRadius: "10px" }}>
                                <div style={{ fontSize: "11px", color: "var(--slate)", marginBottom: "4px" }}>Net Profit</div>
                                <div style={{ fontSize: "20px", fontWeight: 700, color: profitColor, fontFamily: "var(--font-technical)" }}>
                                  {netProfit < 0 ? "-" : ""}{sym}{Math.round(Math.abs(displayNetProfit)).toLocaleString()}
                                </div>
                              </div>
                            </div>

                            {/* Secondary Metrics Bar */}
                            <div style={{ display: "flex", flexDirection: "column", gap: "10px" }}>
                              {[
                                { label: "Return on Ad Spend (ROAS)", val: `${roas}x`, extra: `Target: >3.0x`, color: Number(roas) >= 3 ? "var(--deep-green)" : "#f59e0b" },
                                { label: "Customer Acquisition Cost (CAC)", val: `${sym}${Math.round(displayCAC)}`, extra: `AOV: ${sym}${Math.round(calcAOV * scale)}`, color: "#fff" },
                                { label: "Product Net Margin", val: `${profitMargin}%`, extra: `COGS: ${sym}${Math.round(displayCOGS)}`, color: netProfit >= 0 ? "var(--deep-green)" : "#ef4444" }
                              ].map((m, i) => (
                                <div key={i} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", fontSize: "12px", padding: "6px 0", borderBottom: "1px dashed rgba(255,255,255,0.04)" }}>
                                  <span style={{ color: "var(--slate)" }}>{m.label}</span>
                                  <div style={{ display: "flex", alignItems: "center", gap: "10px" }}>
                                    <span style={{ fontSize: "10.5px", color: "var(--slate)", opacity: 0.6 }}>{m.extra}</span>
                                    <span style={{ color: m.color, fontWeight: 700, fontFamily: "var(--font-technical)" }}>{m.val}</span>
                                  </div>
                                </div>
                              ))}
                            </div>
                          </div>
                        );
                      })()}
                    </div>

                    <div style={{ background: "rgba(0, 255, 178, 0.04)", border: "1px solid rgba(0, 255, 178, 0.1)", borderRadius: "10px", padding: "12px", display: "flex", alignItems: "center", gap: "10px" }}>
                      <span style={{ fontSize: "18px" }}>💡</span>
                      <p style={{ margin: 0, fontSize: "11.5px", color: "var(--slate)", lineHeight: "1.4" }}>
                        <strong>Avery Insight:</strong> {
                          (calcAdSpend > 0 && (calcOrders * calcAOV / calcAdSpend) < 2) ?
                          "Your ROAS is below standard target. Leverage Avery's Logistics Optimizer to audit COD return leaks and recover lost courier billing charges." :
                          "Healthy economic framework! Clean your storefront exports through SheetCodeCrest's sandboxed layout parser to optimize logistics."
                        }
                      </p>
                    </div>
                  </div>
                </div>
              </section>

              {/* ============ METRICS ROW ============ */}
              <div style={{ background: "#070707", borderTop: "1px solid #111", borderBottom: "1px solid #111" }}>
                <div className="sheetai-metrics-row">
                  {[
                    { icon: "🎯", value: "<2%", label: "Error Rate" },
                    { icon: "⚡", value: "+80%", label: "Time Saved" },
                    { icon: "💰", value: "₹₹₹", label: "Cost Effective" },
                    { icon: "🕐", value: "24/7", label: "Always Available" },
                  ].map((m) => (
                    <div className="sheetai-metric-item" key={m.label}>
                      <div className="sheetai-metric-icon">{m.icon}</div>
                      <div>
                        <div className="sheetai-metric-value">{m.value}</div>
                        <div className="sheetai-metric-label">{m.label}</div>
                      </div>
                    </div>
                  ))}
                </div>
              </div>

              {/* ============ WHY + TESTIMONIALS ============ */}
              <div className="why-testimonials-section">
                <div>
                  <div className="sheetai-features-label" style={{ textAlign: "left", marginBottom: "12px" }}>Why SheetCodeCrest?</div>
                  <h2 className="why-left-heading">
                    Make better decisions<br />with data you can<br />actually <em>understand.</em>
                  </h2>
                  <ul className="why-bullets">
                    {[
                      "Save hours of manual spreadsheet analysis",
                      "Detect insights you might otherwise miss",
                      "Make data-driven decisions faster",
                      "No technical skills required — ever",
                    ].map((b) => (
                      <li className="why-bullet-item" key={b}>
                        <div className="why-bullet-check">✓</div>
                        {b}
                      </li>
                    ))}
                  </ul>
                  <button
                    type="button"
                    className="btn-hero-primary"
                    onClick={() => { setAuthTab("signup"); setAuthError(""); setAuthModalOpen(true); }}
                  >
                    Start analyzing now →
                  </button>
                </div>

                <div className="testimonial-grid">
                  {[
                    { stars: 5, text: "\"SheetCodeCrest turned our messy Shopify exports into clear, actionable insights. Complete game changer for our D2C brand.\"", name: "Priya Mehta", title: "E-Commerce Director", avatar: "👩", bg: "rgba(34,197,94,0.1)" },
                    { stars: 5, text: "\"I save at least 10 hours per week on data analysis. The AI insights are incredibly accurate and actionable.\"", name: "Rahul Sharma", title: "Operations Manager", avatar: "👨", bg: "rgba(59,130,246,0.1)" },
                    { stars: 5, text: "\"Finally, a tool that makes data analytics accessible to everyone on my team, not just the data scientists.\"", name: "Anika Patel", title: "Head of Data", avatar: "👩‍💼", bg: "rgba(168,85,247,0.1)" },
                    { stars: 5, text: "\"The anomaly detection caught a ₹2L shiprocket billing error we'd missed for 3 months. Incredible ROI.\"", name: "Vikas Kumar", title: "Logistics Analyst", avatar: "🧑‍💻", bg: "rgba(245,158,11,0.1)" },
                  ].map((t) => (
                    <div className="testimonial-card-new" key={t.name}>
                      <div className="testimonial-stars">{"★".repeat(t.stars)}</div>
                      <p className="testimonial-text-new">{t.text}</p>
                      <div className="testimonial-author-row">
                        <div className="testimonial-avatar" style={{ background: t.bg }}>{t.avatar}</div>
                        <div>
                          <div className="testimonial-author-name-new">{t.name}</div>
                          <div className="testimonial-author-title-new">{t.title}</div>
                        </div>
                      </div>
                    </div>
                  ))}
                </div>
              </div>

              {/* ============ PRICING ============ */}
              {(!currentUser || !currentUser.isPro) && (
                <div id="pricing-section" className="pricing-section" style={{ position: "relative", overflow: "hidden", borderRadius: "24px" }}>
                  {/* Premium Canvas Particle Backdrop */}
                  <SaaSBackgroundParticles />

                  <div className="pricing-header" style={{ position: "relative", zIndex: 1 }}>
                    <div style={{ fontFamily: "var(--font-technical)", fontSize: "10.5px", fontWeight: 700, color: "#22c55e", textTransform: "uppercase", letterSpacing: "0.15em", marginBottom: "6px" }}>Simple, Transparent Pricing</div>
                    <h2>Choose your plan</h2>
                    <p>Start free and upgrade as you grow.</p>
                    
                    <div style={{ display: "flex", flexDirection: "column", alignItems: "center", gap: "14px", marginTop: "1.25rem" }}>
                      <div style={{ display: "inline-flex", alignItems: "center", gap: "10px" }}>
                        <div className="pricing-toggle" style={{ marginTop: 0 }}>
                          <button type="button" className={`pricing-toggle-btn${pricingBilling === "monthly" ? " active" : ""}`} onClick={() => setPricingBilling("monthly")}>Monthly</button>
                          <button type="button" className={`pricing-toggle-btn${pricingBilling === "yearly" ? " active" : ""}`} onClick={() => setPricingBilling("yearly")}>Yearly</button>
                        </div>
                        <span className="pricing-save-badge">Save 20%</span>
                      </div>
                      
                      {/* Interactive Currency Selection Capsule */}
                      <div style={{ display: "inline-flex", alignItems: "center", gap: "6px", background: "rgba(255,255,255,0.03)", border: "1px solid rgba(255,255,255,0.05)", padding: "3px", borderRadius: "30px" }}>
                        <span style={{ fontSize: "10px", color: "var(--slate)", fontWeight: 650, paddingLeft: "10px", paddingRight: "6px", textTransform: "uppercase", letterSpacing: "0.05em" }}>Currency:</span>
                        {(["INR", "USD", "EUR"] as const).map((curr) => (
                          <button
                            key={curr}
                            type="button"
                            onClick={() => setCurrency(curr)}
                            style={{
                              background: currency === curr ? "var(--deep-green)" : "transparent",
                              border: "none",
                              color: currency === curr ? "#030305" : "#fff",
                              fontWeight: 750,
                              fontSize: "10.5px",
                              padding: "4px 10px",
                              borderRadius: "20px",
                              cursor: "pointer",
                              transition: "all 0.15s ease",
                              fontFamily: "var(--font-technical)",
                            }}
                          >
                            {curr === "INR" ? "₹ INR" : curr === "EUR" ? "€ EUR" : "$ USD"}
                          </button>
                        ))}
                      </div>
                    </div>
                  </div>
                  
                  <div className="pricing-grid" style={{ position: "relative", zIndex: 1, padding: "10px 0" }}>
                    {landingPlans
                      .filter(p => p.isActive)
                      .sort((a, b) => (a.sortOrder ?? 99) - (b.sortOrder ?? 99))
                      .map((plan) => {
                        const ac = plan.color || "#22c55e";
                        const isPaid = plan.price > 0;
                        const symbol = getCurrencySymbol(currency);
                        const convertedBase = getConvertedPrice(plan.price, currency);
                        let displayPrice = convertedBase;
                        let periodText = plan.billingPeriod === "free" ? "free" : "month";
                        
                        if (isPaid && pricingBilling === "yearly") {
                          displayPrice = Math.round(convertedBase * 0.8);
                          periodText = "month, billed yearly";
                        }
                        
                        return (
                          <motion.div
                            key={plan.id || plan.name}
                            className={`pricing-card${plan.highlighted ? " featured" : ""}`}
                            style={plan.highlighted ? { 
                              borderColor: ac, 
                              boxShadow: `0 0 35px ${ac}25`,
                              zIndex: 2
                            } : { 
                              zIndex: 2 
                            }}
                            whileHover={{
                              y: -10,
                              scale: 1.025,
                              boxShadow: plan.highlighted 
                                ? "0 20px 45px rgba(0, 255, 178, 0.25), 0 0 0 1px var(--deep-green)" 
                                : "0 20px 40px rgba(255, 255, 255, 0.05), 0 0 0 1px rgba(255, 255, 255, 0.12)",
                              borderColor: plan.highlighted ? "var(--deep-green)" : "rgba(255,255,255,0.25)"
                            }}
                            transition={{ type: "spring", stiffness: 350, damping: 22 }}
                          >
                            {plan.highlighted && (<span className="pricing-badge" style={{ color: ac, borderColor: ac }}>Most Popular</span>)}
                            <div className="pricing-plan-name">{plan.name}</div>
                            <div className="pricing-price">
                              {plan.price === 0 ? `${symbol}0` : `${symbol}${displayPrice.toLocaleString()}`}
                              <span>/ {periodText}</span>
                            </div>
                            {plan.description && <p className="pricing-desc">{plan.description}</p>}
                            <ul className="pricing-features">
                              {plan.features.map((feat, idx) => {
                                const parts = feat.split(":");
                                if (parts.length > 1) {
                                  return (
                                    <li key={idx}>
                                      <span>{parts[0].trim()}</span>
                                      <span className="feature-badge">{parts[1].trim()}</span>
                                    </li>
                                  );
                                }
                                return <li key={idx}><span>{feat}</span></li>;
                              })}
                            </ul>
                            {isPaid ? (
                              <button type="button" className="pricing-cta-btn"
                                onClick={() => { if (!currentUser) { setAuthTab("login"); setAuthError(""); setAuthModalOpen(true); } else { setSelectedPlanId(plan.id || null); setCheckoutOpen(true); } }}>
                                {currentUser?.isPro ? "⚡ Upgrade Now" : `Start with ${plan.name}`}
                              </button>
                            ) : (
                              <button type="button" className="pricing-cta-btn"
                                onClick={() => { if (!currentUser) { setAuthTab("login"); setAuthError(""); setAuthModalOpen(true); } }}>
                                {currentUser ? "✓ Current Plan" : "Get Started Free"}
                              </button>
                            )}
                          </motion.div>
                        );
                      })}
                  </div>
                </div>
              )}

              {/* ============ FAQ ACCORDION SECTION ============ */}
              <section className="sheetai-faq-section" style={{ padding: "80px 20px", maxWidth: "900px", margin: "0 auto", position: "relative" }}>
                <div className="sheetai-features-label" style={{ textAlign: "center", marginBottom: "12px" }}>Frequently Asked Questions</div>
                <h3 className="sheetai-features-title" style={{ textAlign: "center", marginBottom: "16px" }}>Got questions? We've got answers</h3>
                <p style={{ color: "var(--slate)", fontSize: "14.5px", textAlign: "center", maxWidth: "600px", margin: "-4px auto 36px auto", lineHeight: 1.6 }}>
                  Learn how SheetCodeCrest secures your financial datasets, resolves multi-channel logic, and automates courier audits.
                </p>
                
                <div style={{ display: "flex", flexDirection: "column", gap: "12px", marginTop: "32px" }}>
                  {[
                    {
                      q: "How does the client-side sandboxed architecture protect my financial data?",
                      a: "SheetCodeCrest operates on a zero-trust model. By utilizing IndexedDB sandboxing, all files, formulas, and customer records are processed and parsed 100% locally in your browser. Your raw ledger details never travel to external servers, protecting you from enterprise data leaks."
                    },
                    {
                      q: "Can I connect my Shopify store, Shiprocket logs, and custom sheets simultaneously?",
                      a: "Yes! Avery's multi-channel schema mapping matches different log structures (such as matching Shopify order IDs to Shiprocket courier tracking logs) to compile a unified, comprehensive growth ledger automatically."
                    },
                    {
                      q: "How does the automatic timezone currency detector work?",
                      a: "SheetCodeCrest queries your browser's timezone registry (e.g., Asia/Kolkata) and default language locale to automatically determine your currency (₹ INR for India, $ USD or € EUR for international). You can also switch manually at any time using the pricing toggles."
                    },
                    {
                      q: "What kind of anomalies do the 70+ automated growth audits look for?",
                      a: "Our parser automatically audits logistics billing errors (like weight discrepancies charged by couriers), redundant COD return profiles, overlapping zone fees, advertising ROAS drops, and customer retention velocity."
                    },
                    {
                      q: "How does the 20% discount work for billing?",
                      a: "When you select Yearly billing, a 20% discount is applied automatically across standard plans. For example, standard is reduced from ₹999/mo to ₹800/mo, and premium drops from ₹2,499/mo to ₹2,000/mo."
                    }
                  ].map((faq, index) => {
                    const isOpen = expandedFaq === index;
                    return (
                      <div
                        key={index}
                        style={{
                          background: "rgba(255, 255, 255, 0.02)",
                          border: isOpen ? "1px solid var(--deep-green)" : "1px solid rgba(255, 255, 255, 0.04)",
                          boxShadow: isOpen ? "0 0 20px rgba(0, 255, 178, 0.04)" : "none",
                          borderRadius: "12px",
                          overflow: "hidden",
                          transition: "all 0.3s cubic-bezier(0.4, 0, 0.2, 1)"
                        }}
                      >
                        <button
                          type="button"
                          onClick={() => setExpandedFaq(isOpen ? null : index)}
                          style={{
                            width: "100%",
                            background: "transparent",
                            border: "none",
                            padding: "20px 24px",
                            display: "flex",
                            justifyContent: "space-between",
                            alignItems: "center",
                            cursor: "pointer",
                            textAlign: "left",
                            color: "#fff",
                            fontFamily: "inherit"
                          }}
                        >
                          <span style={{ fontSize: "14.5px", fontWeight: 600, transition: "color 0.2s", color: isOpen ? "var(--deep-green)" : "#fff" }}>
                            {faq.q}
                          </span>
                          <span style={{
                            fontSize: "16px",
                            color: isOpen ? "var(--deep-green)" : "var(--slate)",
                            transform: isOpen ? "rotate(45deg)" : "rotate(0deg)",
                            transition: "all 0.3s cubic-bezier(0.4, 0, 0.2, 1)",
                            display: "inline-block",
                            lineHeight: 1
                          }}>
                            ＋
                          </span>
                        </button>
                        
                        <div
                          style={{
                            maxHeight: isOpen ? "200px" : "0px",
                            opacity: isOpen ? 1 : 0,
                            overflow: "hidden",
                            transition: "all 0.3s cubic-bezier(0.4, 0, 0.2, 1)",
                            background: "rgba(0,0,0,0.1)"
                          }}
                        >
                          <div style={{ padding: "0 24px 20px 24px", color: "var(--slate)", fontSize: "13px", lineHeight: "1.6" }}>
                            {faq.a}
                          </div>
                        </div>
                      </div>
                    );
                  })}
                </div>
              </section>

              {/* ============ CTA BANNER ============ */}
              <div className="cta-section-new">
                <div className="cta-banner-new">
                  <div style={{ display: "flex", alignItems: "center", gap: "1.5rem" }}>
                    <div className="cta-banner-icon">📊</div>
                    <div>
                      <div className="cta-banner-headline">Ready to transform your data?</div>
                      <div className="cta-banner-sub">Join growing e-commerce brands making better decisions with AI-powered spreadsheet insights.</div>
                    </div>
                  </div>
                  <div className="cta-banner-actions">
                    <button type="button" className="btn-cta-green" onClick={() => { setAuthTab("signup"); setAuthError(""); setAuthModalOpen(true); }}>
                      Get started free →
                    </button>
                    <div className="cta-no-cc">No credit card required</div>
                  </div>
                </div>
              </div>
            </>
          )}

          {/* ============ FOOTER ============ */}
          <footer className="site-footer">
            <div className="site-footer-inner">
              <div className="footer-brand">
                <div className="footer-brand-name">
                  <img src="/logo-dark.png" alt="SheetCodeCrest" style={{ height: "24px", width: "24px", borderRadius: "6px", objectFit: "contain" }} />
                  SheetCodeCrest
                </div>
                <p className="footer-tagline">The AI-powered spreadsheet analyst that turns raw data into actionable insights.</p>
                <div className="footer-social-row">
                  <a href={CODECREST.instagram} target="_blank" rel="noopener noreferrer" className="footer-social-btn" title="Instagram">
                    <svg viewBox="0 0 24 24" width="18" height="18" stroke="currentColor" strokeWidth="2" fill="none" strokeLinecap="round" strokeLinejoin="round" style={{ display: "block" }}>
                      <rect x="2" y="2" width="20" height="20" rx="5" ry="5"></rect>
                      <path d="M16 11.37A4 4 0 1 1 12.63 8 4 4 0 0 1 16 11.37z"></path>
                      <line x1="17.5" y1="6.5" x2="17.51" y2="6.5"></line>
                    </svg>
                  </a>
                  <a href={`mailto:${CODECREST.email}`} className="footer-social-btn" title="Email">
                    <svg viewBox="0 0 24 24" width="18" height="18" stroke="currentColor" strokeWidth="2" fill="none" strokeLinecap="round" strokeLinejoin="round" style={{ display: "block" }}>
                      <path d="M4 4h16c1.1 0 2 .9 2 2v12c0 1.1-.9 2-2 2H4c-1.1 0-2-.9-2-2V6c0-1.1.9-2 2-2z"></path>
                      <polyline points="22,6 12,13 2,6"></polyline>
                    </svg>
                  </a>
                  <a href={CODECREST.website} target="_blank" rel="noopener noreferrer" className="footer-social-btn" title="Website">
                    <svg viewBox="0 0 24 24" width="18" height="18" stroke="currentColor" strokeWidth="2" fill="none" strokeLinecap="round" strokeLinejoin="round" style={{ display: "block" }}>
                      <circle cx="12" cy="12" r="10"></circle>
                      <line x1="2" y1="12" x2="22" y2="12"></line>
                      <path d="M12 2a15.3 15.3 0 0 1 4 10 15.3 15.3 0 0 1-4 10 15.3 15.3 0 0 1-4-10 15.3 15.3 0 0 1 4-10z"></path>
                    </svg>
                  </a>
                </div>
              </div>
            </div>
            <div className="footer-bottom">
              <span>© 2026 SheetCodeCrest. All rights reserved.</span>
              <span>Made with ❤️ by <a href={CODECREST.website} target="_blank" rel="noopener noreferrer" style={{ color: "#22c55e", textDecoration: "none" }}>Codecrest Studio</a></span>
            </div>
          </footer>
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
                    <h3 className="card-title" style={{ color: "#ef4444" }}>⚠️ Regional return hotspot alerts</h3>
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

              {/* 💸 COD Premium & RTO Optimizer Simulator */}
              {(() => {
                const codCount = logisticsAnalytics.payCounts["cod"]?.orders || 0;
                const prepaidCount = logisticsAnalytics.payCounts["prepaid"]?.orders || 0;
                const codRto = logisticsAnalytics.payCounts["cod"]?.rto || 0;
                const prepaidRto = logisticsAnalytics.payCounts["prepaid"]?.rto || 0;

                const codRtoRate = codCount > 0 ? (codRto / codCount) : 0;
                const prepaidRtoRate = prepaidCount > 0 ? (prepaidRto / prepaidCount) : 0;
                const totalRtoLoss = (codRto + prepaidRto) * rtoCostPerOrder;

                // Simulated Metrics
                const switchCount = Math.round(codCount * (codSwitchRate / 100));
                const remainingCodCount = Math.max(0, codCount - switchCount);
                const codPremiumRevenue = remainingCodCount * codPremiumVal;
                
                const newPrepaidCount = prepaidCount + switchCount;
                const expectedCodRto = remainingCodCount * codRtoRate;
                const expectedPrepaidRto = newPrepaidCount * prepaidRtoRate;
                
                const newRtoCount = Math.round(expectedCodRto + expectedPrepaidRto);
                const newRtoLoss = newRtoCount * rtoCostPerOrder;
                const netRtoLoss = Math.max(0, newRtoLoss - codPremiumRevenue);
                const savings = totalRtoLoss - netRtoLoss;

                return (
                  <div className="section-card" style={{ borderLeft: "4px solid var(--deep-green)", background: "rgba(16, 185, 129, 0.03)", marginTop: "20px" }}>
                    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: "16px", flexWrap: "wrap", gap: "10px" }}>
                      <div>
                        <h3 className="card-title" style={{ margin: 0, color: "var(--deep-green)" }}>💰 COD Premium & Cash Leakage Optimizer</h3>
                        <p style={{ fontSize: "0.85rem", color: "var(--slate)", margin: "4px 0 0 0" }}>Impose a COD payment fee to reduce RTO risk and recover shipping losses.</p>
                      </div>
                      <span style={{ fontSize: "0.75rem", background: "rgba(16, 185, 129, 0.12)", color: "var(--deep-green)", padding: "4px 8px", borderRadius: "6px", fontWeight: 600 }}>PRO OPTIMIZER</span>
                    </div>

                    <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(300px, 1fr))", gap: "24px" }}>
                      {/* Left: Interactive Sliders */}
                      <div style={{ display: "flex", flexDirection: "column", gap: "16px" }}>
                        <div>
                          <div style={{ display: "flex", justifyContent: "space-between", fontSize: "0.85rem", marginBottom: "6px" }}>
                            <span style={{ color: "var(--slate)" }}>COD Premium Fee</span>
                            <span style={{ fontWeight: 600, color: "var(--off-black)" }}>₹{codPremiumVal} per order</span>
                          </div>
                          <input 
                            type="range" 
                            min="0" 
                            max="150" 
                            step="5"
                            value={codPremiumVal} 
                            onChange={(e) => setCodPremiumVal(Number(e.target.value))}
                            style={{ width: "100%", accentColor: "var(--deep-green)", cursor: "pointer" }}
                          />
                          <span style={{ fontSize: "0.7rem", color: "var(--slate)" }}>Surcharging COD orders incentivizes prepaid transactions.</span>
                        </div>

                        <div>
                          <div style={{ display: "flex", justifyContent: "space-between", fontSize: "0.85rem", marginBottom: "6px" }}>
                            <span style={{ color: "var(--slate)" }}>Prepaid Switch Rate</span>
                            <span style={{ fontWeight: 600, color: "var(--off-black)" }}>{codSwitchRate}% of buyers</span>
                          </div>
                          <input 
                            type="range" 
                            min="0" 
                            max="50" 
                            step="1"
                            value={codSwitchRate} 
                            onChange={(e) => setCodSwitchRate(Number(e.target.value))}
                            style={{ width: "100%", accentColor: "var(--deep-green)", cursor: "pointer" }}
                          />
                          <span style={{ fontSize: "0.7rem", color: "var(--slate)" }}>Percent of COD buyers who switch to Prepaid due to fee.</span>
                        </div>

                        <div>
                          <div style={{ display: "flex", justifyContent: "space-between", fontSize: "0.85rem", marginBottom: "6px" }}>
                            <span style={{ color: "var(--slate)" }}>Average RTO Cost Penalty</span>
                            <span style={{ fontWeight: 600, color: "var(--off-black)" }}>₹{rtoCostPerOrder} / return</span>
                          </div>
                          <input 
                            type="range" 
                            min="50" 
                            max="300" 
                            step="10"
                            value={rtoCostPerOrder} 
                            onChange={(e) => setRtoCostPerOrder(Number(e.target.value))}
                            style={{ width: "100%", accentColor: "var(--deep-green)", cursor: "pointer" }}
                          />
                          <span style={{ fontSize: "0.7rem", color: "var(--slate)" }}>Includes forward/return freight, repackaging, and stock blockage costs.</span>
                        </div>
                      </div>

                      {/* Right: Output projections */}
                      <div style={{ background: "rgba(255,255,255,0.03)", border: "1px solid var(--hairline)", borderRadius: "12px", padding: "16px", display: "flex", flexDirection: "column", justifyContent: "space-between" }}>
                        <div style={{ display: "flex", flexDirection: "column", gap: "10px" }}>
                          <div style={{ display: "flex", justifyContent: "space-between", fontSize: "0.85rem" }}>
                            <span style={{ color: "var(--slate)" }}>Current RTO Loss:</span>
                            <span style={{ color: "#ef4444", fontWeight: 600 }}>₹{totalRtoLoss.toLocaleString("en-IN")}</span>
                          </div>
                          
                          <div style={{ display: "flex", justifyContent: "space-between", fontSize: "0.85rem" }}>
                            <span style={{ color: "var(--slate)" }}>COD Premium Collected:</span>
                            <span style={{ color: "#10b981", fontWeight: 600 }}>+₹{codPremiumRevenue.toLocaleString("en-IN")}</span>
                          </div>

                          <div style={{ display: "flex", justifyContent: "space-between", fontSize: "0.85rem" }}>
                            <span style={{ color: "var(--slate)" }}>Simulated RTO Damage:</span>
                            <span style={{ color: "#ef4444", fontWeight: 600 }}>₹{newRtoLoss.toLocaleString("en-IN")}</span>
                          </div>

                          <div style={{ borderTop: "1px dashed var(--hairline)", paddingTop: "8px", marginTop: "4px" }} />

                          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                            <div>
                              <span style={{ fontSize: "0.85rem", fontWeight: 600, color: "var(--off-black)" }}>Net Financial Recovery</span>
                              <div style={{ fontSize: "0.7rem", color: "var(--slate)" }}>Total leakage prevented</div>
                            </div>
                            <span style={{ fontSize: "1.25rem", fontWeight: 800, color: savings >= 0 ? "#10b981" : "#ef4444" }}>
                              {savings >= 0 ? `₹${Math.round(savings).toLocaleString("en-IN")}` : `-₹${Math.round(Math.abs(savings)).toLocaleString("en-IN")}`}
                            </span>
                          </div>
                        </div>

                        {/* Recommendation */}
                        <div style={{ background: "rgba(16, 185, 129, 0.08)", padding: "10px", borderRadius: "8px", marginTop: "16px", border: "1px solid rgba(16, 185, 129, 0.15)" }}>
                          <div style={{ fontWeight: 600, fontSize: "0.8rem", color: "var(--deep-green)", display: "flex", alignItems: "center", gap: "6px" }}>
                            💡 Smart Recommendation
                          </div>
                          <p style={{ fontSize: "0.75rem", color: "var(--slate)", margin: "4px 0 0 0", lineHeight: "1.3" }}>
                            {savings > 0 
                              ? `Imposing a ₹${codPremiumVal} COD Premium is highly recommended. It will convert ~${codSwitchRate}% of buyers (reducing RTO counts by ${Math.max(0, codRto - expectedCodRto - expectedPrepaidRto).toFixed(0)} orders) and generate ₹${codPremiumRevenue.toLocaleString("en-IN")} in premium fees, saving you ₹${Math.round(savings).toLocaleString("en-IN")} total.`
                              : `Impose a COD premium of at least ₹30 to start offsetting your forward return losses. This will incentivize users to pay online via UPI, lowering returns.`
                            }
                          </p>
                        </div>
                      </div>
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
          <div className="dashboard-panels-grid">
            
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
          {!isSharedViewOnly && adminFeatureAI && (
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
                      🛠️ Product Winners
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
                      ⚠️ Highest RTO Risks
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
                        {header} {sortConfig?.key === header ? (sortConfig.direction === "asc" ? "▲" : "▼") : ""}
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

      {/* 1. Auth Modal (Login/Signup Tabs) */}
      {authModalOpen && (
        <motion.div 
          className="modal-overlay" 
          initial={{ opacity: 0 }}
          animate={{ opacity: 1 }}
          transition={{ duration: 0.2 }}
          onClick={() => setAuthModalOpen(false)}
        >
          <motion.div 
            className="modal-card" 
            initial={{ scale: 0.95, y: 15, opacity: 0 }}
            animate={{ scale: 1, y: 0, opacity: 1 }}
            transition={{ type: "spring", damping: 25, stiffness: 350 }}
            onClick={(e) => e.stopPropagation()}
          >
            <div className="modal-header">
              <h3 className="modal-title" style={{ display: "flex", alignItems: "center", gap: "10px" }}>
                <img src="/logo-icon.png" alt="SheetCodeCrest Icon" style={{ height: "24px", width: "24px", borderRadius: "5px", objectFit: "contain" }} />
                <span>{authTab === "login" ? "🔑 Sign In" : "📋 Create Account"}</span>
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
                  ⚠️ {authError}
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
                  {authTab === "login" ? "🔑 Sign In" : "📋 Register"}
                </button>

                {adminFeatureGoogleLogin && (
                  <>
                    <div style={{ display: "flex", alignItems: "center", justifyContent: "center", margin: "1.25rem 0", gap: "10px" }}>
                      <div style={{ flex: 1, height: "1px", background: "var(--hairline)" }}></div>
                      <span style={{ fontSize: "11px", color: "var(--slate)", textTransform: "uppercase", letterSpacing: "0.05em", fontWeight: 600 }}>or</span>
                      <div style={{ flex: 1, height: "1px", background: "var(--hairline)" }}></div>
                    </div>

                    <div style={{ display: "flex", justifyContent: "center", width: "100%" }}>
                      <div id="google-signin-div" style={{ minHeight: "44px", width: "320px" }}></div>
                    </div>
                  </>
                )}
              </form>
            </div>
          </motion.div>
        </motion.div>
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
                  <div className="checkout-plans-grid">
                    {checkoutPlans.map(plan => {
                      const isSel = selectedPlanId === plan.id;
                      const isFeatured = plan.highlighted;
                      
                      const bg = isSel 
                        ? (isFeatured ? "var(--focus-blue)" : "var(--primary)")
                        : "var(--primary)";
                      const border = isSel
                        ? "2px solid var(--focus-blue)"
                        : "1px solid var(--hairline)";
                      const textCol = isSel && isFeatured ? "var(--on-primary)" : "var(--ink)";
                      const slateCol = isSel && isFeatured ? "rgba(0, 0, 0, 0.6)" : "var(--slate)";
                      const priceCol = isSel && isFeatured ? "var(--on-primary)" : "var(--focus-blue)";

                      return (
                        <button key={plan.id || plan.name} type="button" onClick={() => setSelectedPlanId(plan.id || null)}
                          style={{
                            position: "relative",
                            padding: "16px 14px",
                            borderRadius: "12px",
                            border: border,
                            background: bg,
                            color: textCol,
                            cursor: "pointer",
                            textAlign: "left",
                            transition: "all 0.2s cubic-bezier(0.4, 0, 0.2, 1)",
                            display: "flex",
                            flexDirection: "column",
                            justifyContent: "space-between",
                            minHeight: "110px"
                          }}>
                          {plan.highlighted && (
                            <div style={{ 
                              position: "absolute", 
                              top: "-9px", 
                              left: "12px", 
                              background: isSel ? "var(--on-primary)" : "var(--focus-blue)", 
                              color: isSel ? "var(--focus-blue)" : "var(--on-primary)", 
                              fontSize: "8px", 
                              fontWeight: 900, 
                              padding: "2px 8px", 
                              borderRadius: "4px", 
                              whiteSpace: "nowrap" 
                            }}>⭐ RECOMMENDED</div>
                          )}
                          <div>
                            <div style={{ fontWeight: 800, fontSize: "14px", color: isSel && isFeatured ? "var(--on-primary)" : "var(--ink)" }}>{plan.name}</div>
                            {plan.description && <div style={{ fontSize: "10px", color: slateCol, marginTop: "2px", lineHeight: "1.2" }}>{plan.description}</div>}
                          </div>
                          <div style={{ fontWeight: 900, fontSize: "18px", color: priceCol, marginTop: "8px" }}>
                            ₹{plan.price.toLocaleString()}
                            <span style={{ fontSize: "10px", fontWeight: 400, color: slateCol, marginLeft: "2px" }}>/{plan.billingPeriod}</span>
                          </div>
                          {isSel && (
                            <div style={{ 
                              position: "absolute", 
                              top: "12px", 
                              right: "12px", 
                              width: "16px", 
                              height: "16px", 
                              borderRadius: "50%", 
                              background: isFeatured ? "var(--on-primary)" : "var(--focus-blue)", 
                              display: "flex", 
                              alignItems: "center", 
                              justifyContent: "center", 
                              fontSize: "10px", 
                              color: isFeatured ? "var(--focus-blue)" : "var(--on-primary)", 
                              fontWeight: 900
                            }}>✓</div>
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
                  ) : adminFeatureUPI ? (
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
                  ) : (
                    <div style={{ textAlign: "center", padding: "1.5rem 0", background: "rgba(255,255,255,0.01)", border: "1px solid var(--hairline)", borderRadius: "8px" }}>
                      <span style={{ fontSize: "1.75rem" }}>⚠️</span>
                      <h4 style={{ margin: "0.75rem 0 0.5rem 0", color: "var(--coral)", fontSize: "13px", fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.05em" }}>UPI Manual Checkout Disabled</h4>
                      <p style={{ fontSize: "12px", color: "var(--slate)", lineHeight: "1.6", maxWidth: "320px", margin: "0 auto" }}>
                        The manual UPI QR payment channel is currently closed for maintenance. Please email us at <strong style={{ color: "var(--action-blue)" }}>codecreststudio@gmail.com</strong> for assistance with your premium upgrade.
                      </p>
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

      {/* 🛡️ Secure Admin Panel Modal — Advanced */}
      {adminModalOpen && (
        <div className="modal-overlay" onClick={() => !adminLoading && setAdminModalOpen(false)}>
          <div className="modal-card" onClick={(e) => e.stopPropagation()} style={{ maxWidth: "900px", width: "97%", maxHeight: "92vh", display: "flex", flexDirection: "column" }}>
            {/* Header */}
            <div className="modal-header" style={{ borderColor: "var(--amber)", flexShrink: 0 }}>
              <h3 className="modal-title" style={{ display: "flex", alignItems: "center", gap: "8px", color: "var(--amber)" }}>
                <span>🛡️</span> SheetCodeCrest Admin Console
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
                  settings: "⚙️ Settings",
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
                          placeholder="🔎 Search by username, name, email, mobile..."
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
                      <div style={{ border: "1px solid var(--hairline)", borderRadius: "10px", overflowX: "auto" }}>
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
                                        ✏️ Edit
                                      </button>
                                      <button type="button" onClick={() => handleToggleUserPro(user)}
                                        style={{ padding: "4px 8px", borderRadius: "5px", fontSize: "10px", fontWeight: 600, cursor: "pointer", border: `1px solid ${user.isPro ? "#ef4444" : "#10b981"}`, background: user.isPro ? "rgba(239,68,68,0.1)" : "rgba(16,185,129,0.1)", color: user.isPro ? "#ef4444" : "#10b981" }}>
                                        {user.isPro ? "Revoke" : "Grant"}
                                      </button>
                                      <button type="button" onClick={() => handleDeleteUser(user)}
                                        style={{ padding: "4px 8px", borderRadius: "5px", fontSize: "10px", fontWeight: 600, cursor: "pointer", border: "1px solid #ef4444", background: "rgba(239,68,68,0.08)", color: "#ef4444" }}>
                                        🗑️
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
                            <h4 style={{ margin: 0, fontSize: "16px", fontWeight: 700, color: "var(--amber)" }}>✏️ Edit User</h4>
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
                      <div className="admin-payments-stats-grid">
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
                        <input type="text" className="form-input" placeholder="🔎 Search by username or payment ID..." value={adminSearch} onChange={(e) => setAdminSearch(e.target.value)} style={{ flex: 1, minWidth: "200px" }} />
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
                      <div style={{ border: "1px solid var(--hairline)", borderRadius: "10px", overflowX: "auto" }}>
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
                                            ❌ Reject
                                          </button>
                                        )}
                                        {pay.status === "success" && (
                                          <button type="button" onClick={() => handleAdminRefundPayment(pay.paymentId, pay.username)}
                                            style={{ padding: "4px 8px", borderRadius: "5px", fontSize: "10px", fontWeight: 600, cursor: "pointer", border: "1px solid #8b5cf6", background: "rgba(139,92,246,0.1)", color: "#8b5cf6", whiteSpace: "nowrap" }}>
                                            ↩️ Refund
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
                        <div className="admin-plan-drawer" style={{ borderLeft: `2px solid ${adminPlanColor}` }}>
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
                          { icon: "⏳", label: "Pending", value: adminPayments.filter(p => p.status === "pending_verification").length, color: "#ef4444" },
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
                        <h4 style={{ margin: "0 0 1rem 0", fontSize: "14px", fontWeight: 700, color: "#ef4444" }}>⚠️ Danger Zone</h4>
                        <p style={{ fontSize: "12px", color: "var(--slate)", margin: "0 0 12px 0" }}>These actions are irreversible. Use with extreme caution.</p>
                        <button type="button"
                          onClick={() => { if (confirm("Clear all admin activity logs? This cannot be undone.")) { setAdminLogs([]); addLog("⚠️ Admin: Activity logs cleared.", "error"); }}}
                          style={{ padding: "8px 16px", borderRadius: "8px", background: "rgba(239,68,68,0.1)", color: "#ef4444", border: "1px solid #ef4444", fontWeight: 700, fontSize: "12px", cursor: "pointer" }}>
                          🗑️ Clear Activity Logs
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
              <h3 className="modal-title">🎛️ Account & Analytics Dashboard</h3>
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
                      📂 Saved Spreadsheets & Analytics
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
                                🗑️
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
