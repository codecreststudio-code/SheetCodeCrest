import { supabase, isSupabaseConfigured } from "./supabaseClient";

export interface User {
  username: string;
  passwordHash: string;
  isPro: boolean;
  dateCreated: string;
  name?: string;
  mobile?: string;
  email?: string;
}

export interface SavedRecord {
  id?: number;
  username: string;
  filename: string;
  size: number;
  mode: "universal" | "logistics" | "shopify";
  rawRows: any[];
  tableHeaders: string[];
  dataProfile: any | null;
  logisticsAnalytics: any | null;
  shopifyAnalytics: any | null;
  outName: string;
  chatHistory: any[];
  timestamp: string;
  comments?: Array<{ id: string; author: string; text: string; timestamp: string }>;
  versions?: Array<{
    timestamp: string;
    filename: string;
    size: number;
    rawRows: any[];
    tableHeaders: string[];
    dataProfile: any | null;
    logisticsAnalytics: any | null;
    shopifyAnalytics: any | null;
    outName: string;
  }>;
}

export interface PaymentLog {
  id?: number;
  username: string;
  gateway: "stripe" | "razorpay";
  paymentId: string;
  orderId?: string;
  signature?: string;
  amount: number;
  status: string;
}

export interface Plan {
  id?: string;
  name: string;
  price: number;
  billingPeriod: "monthly" | "yearly" | "lifetime" | "free";
  features: string[];
  isActive: boolean;
  description?: string;       // Short tagline shown under the plan name
  highlighted?: boolean;      // Show a "Recommended" badge on this plan
  color?: string;             // Accent hex color for the plan card
  maxReports?: number;        // 0 = unlimited; >0 = capped per period
  sortOrder?: number;         // Display ordering (lower = first)
  createdAt?: string;
}

export interface AdminLog {
  id?: number;
  action: string;
  performedBy: string;
  details?: string;
  createdAt?: string;
}

const DB_NAME = "auto_excel_intel_db";
const DB_VERSION = 1;

// ----------------------------------------------------
// 💾 LOCAL INDEXEDDB UTILITIES (Cache & Fallback)
// ----------------------------------------------------
export function openDB(): Promise<IDBDatabase> {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(DB_NAME, DB_VERSION);
    request.onupgradeneeded = (event: any) => {
      const db = event.target.result;
      if (!db.objectStoreNames.contains("users")) {
        db.createObjectStore("users", { keyPath: "username" });
      }
      if (!db.objectStoreNames.contains("records")) {
        const recordsStore = db.createObjectStore("records", { keyPath: "id", autoIncrement: true });
        recordsStore.createIndex("username", "username", { unique: false });
      }
    };
    request.onsuccess = (event: any) => resolve(event.target.result);
    request.onerror = (event: any) => reject(event.target.error);
  });
}

// ----------------------------------------------------
// 👤 USER DB OPERATIONS
// ----------------------------------------------------
export async function dbSaveUser(user: User): Promise<void> {
  // 1. Replicate to Local Cache (IndexedDB)
  const localDb = await openDB();
  await new Promise<void>((resolve, reject) => {
    const transaction = localDb.transaction("users", "readwrite");
    const store = transaction.objectStore("users");
    const request = store.put(user);
    request.onsuccess = () => resolve();
    request.onerror = () => reject(request.error);
  });

  // 2. Sync to Supabase Cloud if configured
  if (isSupabaseConfigured) {
    try {
      const { error } = await supabase
        .from("users")
        .upsert({
          username: user.username,
          password_hash: user.passwordHash,
          is_pro: user.isPro,
          date_created: user.dateCreated,
          name: user.name || "",
          mobile: user.mobile || "",
          email: user.email || ""
        });
      if (error) throw error;
      console.log(`☁️ Supabase: User "${user.username}" synced successfully.`);
    } catch (err) {
      console.error("❌ Supabase user sync error:", err);
    }
  }
}

export async function dbGetUser(username: string): Promise<User | null> {
  // 1. Try to fetch from Supabase Cloud if configured
  if (isSupabaseConfigured) {
    try {
      const { data, error } = await supabase
        .from("users")
        .select("*")
        .eq("username", username)
        .maybeSingle();
      
      if (error) throw error;
      
      if (data) {
        const mappedUser: User = {
          username: data.username,
          passwordHash: data.password_hash,
          isPro: data.is_pro,
          dateCreated: data.date_created,
          name: data.name,
          mobile: data.mobile,
          email: data.email
        };
        // Update local cache
        const localDb = await openDB();
        const transaction = localDb.transaction("users", "readwrite");
        transaction.objectStore("users").put(mappedUser);
        return mappedUser;
      }
    } catch (err) {
      console.error("❌ Supabase user fetch error, falling back to local storage:", err);
    }
  }

  // 2. Fallback to Local Cache (IndexedDB)
  const localDb = await openDB();
  return new Promise((resolve, reject) => {
    const transaction = localDb.transaction("users", "readonly");
    const store = transaction.objectStore("users");
    const request = store.get(username);
    request.onsuccess = () => resolve(request.result || null);
    request.onerror = () => reject(request.error);
  });
}

// ----------------------------------------------------
// 📊 EXCEL RECORDS DB OPERATIONS
// ----------------------------------------------------
export async function dbSaveRecord(record: SavedRecord): Promise<number> {
  let generatedId = record.id || 0;

  // 1. Sync to Supabase Cloud if configured
  if (isSupabaseConfigured) {
    try {
      const payload: any = {
        username: record.username,
        filename: record.filename,
        size: record.size,
        mode: record.mode,
        raw_rows: record.rawRows,
        table_headers: record.tableHeaders,
        data_profile: record.dataProfile,
        logistics_analytics: record.logisticsAnalytics,
        shopify_analytics: record.shopifyAnalytics,
        out_name: record.outName,
        chat_history: record.chatHistory,
        timestamp: record.timestamp,
        comments: record.comments || [],
        versions: record.versions || []
      };

      if (record.id) {
        payload.id = record.id;
      }

      const { data, error } = await supabase
        .from("records")
        .upsert(payload)
        .select("id")
        .single();

      if (error) throw error;
      if (data) {
        generatedId = data.id;
        record.id = generatedId;
        console.log(`☁️ Supabase: Record auto-saved and synced (ID: ${generatedId}).`);
      }
    } catch (err) {
      console.error("❌ Supabase record sync error:", err);
    }
  }

  // 2. Persist to Local Cache (IndexedDB)
  const localDb = await openDB();
  return new Promise((resolve, reject) => {
    const transaction = localDb.transaction("records", "readwrite");
    const store = transaction.objectStore("records");
    const request = store.put(record);
    request.onsuccess = (event: any) => resolve(record.id || event.target.result);
    request.onerror = () => reject(request.error);
  });
}

export async function dbGetRecordById(id: number): Promise<SavedRecord | null> {
  // 1. Try to fetch from Supabase Cloud if configured
  if (isSupabaseConfigured) {
    try {
      const { data, error } = await supabase
        .from("records")
        .select("*")
        .eq("id", id)
        .maybeSingle();

      if (error) throw error;

      if (data) {
        return {
          id: data.id,
          username: data.username,
          filename: data.filename,
          size: data.size,
          mode: data.mode as any,
          rawRows: data.raw_rows,
          tableHeaders: data.table_headers,
          dataProfile: data.data_profile,
          logisticsAnalytics: data.logistics_analytics,
          shopifyAnalytics: data.shopify_analytics,
          outName: data.out_name,
          chatHistory: data.chat_history || [],
          timestamp: data.timestamp,
          comments: data.comments || [],
          versions: data.versions || []
        };
      }
    } catch (err) {
      console.error("❌ Supabase record by ID fetch error, falling back to local cache:", err);
    }
  }

  // 2. Fallback to Local Cache (IndexedDB)
  const localDb = await openDB();
  return new Promise((resolve, reject) => {
    const transaction = localDb.transaction("records", "readonly");
    const store = transaction.objectStore("records");
    const request = store.get(id);
    request.onsuccess = () => resolve(request.result || null);
    request.onerror = () => reject(request.error);
  });
}

export async function dbGetRecords(username: string): Promise<SavedRecord[]> {
  // 1. Try to fetch from Supabase Cloud if configured
  if (isSupabaseConfigured) {
    try {
      const { data, error } = await supabase
        .from("records")
        .select("*")
        .eq("username", username)
        .order("id", { ascending: false });

      if (error) throw error;

      if (data) {
        const mappedRecords: SavedRecord[] = data.map((d: any) => ({
          id: d.id,
          username: d.username,
          filename: d.filename,
          size: d.size,
          mode: d.mode as any,
          rawRows: d.raw_rows,
          tableHeaders: d.table_headers,
          dataProfile: d.data_profile,
          logisticsAnalytics: d.logistics_analytics,
          shopifyAnalytics: d.shopify_analytics,
          outName: d.out_name,
          chatHistory: d.chat_history,
          timestamp: d.timestamp,
          comments: d.comments || [],
          versions: d.versions || []
        }));

        // Replicate to local Cache for offline viewing speed
        const localDb = await openDB();
        const transaction = localDb.transaction("records", "readwrite");
        const store = transaction.objectStore("records");
        for (const rec of mappedRecords) {
          store.put(rec);
        }
        return mappedRecords;
      }
    } catch (err) {
      console.error("❌ Supabase records fetch error, falling back to local cache:", err);
    }
  }

  // 2. Fallback to Local Cache (IndexedDB)
  const localDb = await openDB();
  return new Promise((resolve, reject) => {
    const transaction = localDb.transaction("records", "readonly");
    const store = transaction.objectStore("records");
    const index = store.index("username");
    const request = index.getAll(username);
    request.onsuccess = () => {
      const sorted = (request.result || []).sort((a: any, b: any) => (b.id || 0) - (a.id || 0));
      resolve(sorted);
    };
    request.onerror = () => reject(request.error);
  });
}

export async function dbDeleteRecord(id: number): Promise<void> {
  // 1. Sync deletion to Supabase if configured
  if (isSupabaseConfigured) {
    try {
      const { error } = await supabase
        .from("records")
        .delete()
        .eq("id", id);
      if (error) throw error;
      console.log(`☁️ Supabase: Record (ID: ${id}) deleted successfully.`);
    } catch (err) {
      console.error("❌ Supabase record delete error:", err);
    }
  }

  // 2. Delete from Local Cache (IndexedDB)
  const localDb = await openDB();
  await new Promise<void>((resolve, reject) => {
    const transaction = localDb.transaction("records", "readwrite");
    const store = transaction.objectStore("records");
    const request = store.delete(id);
    request.onsuccess = () => resolve();
    request.onerror = () => reject(request.error);
  });
}

// ----------------------------------------------------
// 💳 TRANSACTIONS LOGGING
// ----------------------------------------------------
export async function dbLogPayment(log: PaymentLog): Promise<void> {
  if (isSupabaseConfigured) {
    try {
      const { error } = await supabase
        .from("payments")
        .insert({
          username: log.username,
          gateway: log.gateway,
          payment_id: log.paymentId,
          order_id: log.orderId || null,
          signature: log.signature || null,
          amount: log.amount,
          status: log.status
        });
      if (error) throw error;
      console.log(`☁️ Supabase: Payment Log successfully saved in cloud.`);
    } catch (err) {
      console.error("❌ Supabase payment log error:", err);
    }
  }
}

// ----------------------------------------------------
// 🛡️ ADMIN PORTAL DATABASE OPERATIONS
// ----------------------------------------------------
export async function dbGetAllUsers(): Promise<User[]> {
  if (isSupabaseConfigured) {
    try {
      const { data, error } = await supabase
        .from("users")
        .select("*")
        .order("username", { ascending: true });
      if (error) throw error;
      if (data) {
        return data.map((d: any) => ({
          username: d.username,
          passwordHash: d.password_hash,
          isPro: d.is_pro,
          dateCreated: d.date_created,
          name: d.name,
          mobile: d.mobile,
          email: d.email
        }));
      }
    } catch (err) {
      console.error("❌ Supabase get all users error, falling back to local database:", err);
    }
  }
  
  // Local cache fallback
  const localDb = await openDB();
  return new Promise((resolve, reject) => {
    const transaction = localDb.transaction("users", "readonly");
    const store = transaction.objectStore("users");
    const request = store.getAll();
    request.onsuccess = () => resolve(request.result || []);
    request.onerror = () => reject(request.error);
  });
}

export async function dbGetAllPayments(): Promise<PaymentLog[]> {
  if (isSupabaseConfigured) {
    try {
      const { data, error } = await supabase
        .from("payments")
        .select("*")
        .order("id", { ascending: false });
      if (error) throw error;
      if (data) {
        return data.map((d: any) => ({
          id: d.id,
          username: d.username,
          gateway: d.gateway as any,
          paymentId: d.payment_id,
          orderId: d.order_id,
          signature: d.signature,
          amount: d.amount,
          status: d.status
        }));
      }
    } catch (err) {
      console.error("❌ Supabase get all payments error:", err);
    }
  }
  return [];
}

export async function dbApprovePayment(paymentId: string, username: string): Promise<void> {
  // 1. Sync to Supabase Cloud
  if (isSupabaseConfigured) {
    try {
      // Update payment status
      const { error: pError } = await supabase
        .from("payments")
        .update({ status: "success" })
        .eq("payment_id", paymentId);
      if (pError) throw pError;
      
      // Update user role
      const { error: uError } = await supabase
        .from("users")
        .update({ is_pro: true })
        .eq("username", username);
      if (uError) throw uError;
      
      console.log(`☁️ Supabase: Payment reference "${paymentId}" approved and upgraded user "${username}".`);
    } catch (err) {
      console.error("❌ Supabase payment approval error:", err);
    }
  }
  
  // 2. Replicate locally to cache
  const userObj = await dbGetUser(username);
  if (userObj) {
    userObj.isPro = true;
    await dbSaveUser(userObj);
  }
}

// ----------------------------------------------------
// 🗑️ USER MANAGEMENT (Delete, Update)
// ----------------------------------------------------
export async function dbDeleteUser(username: string): Promise<void> {
  if (isSupabaseConfigured) {
    try {
      const { error } = await supabase
        .from("users")
        .delete()
        .eq("username", username);
      if (error) throw error;
      console.log(`☁️ Supabase: User "${username}" deleted.`);
    } catch (err) {
      console.error("❌ Supabase delete user error:", err);
    }
  }
  // Also delete from local cache
  const localDb = await openDB();
  await new Promise<void>((resolve, reject) => {
    const transaction = localDb.transaction("users", "readwrite");
    const store = transaction.objectStore("users");
    const request = store.delete(username);
    request.onsuccess = () => resolve();
    request.onerror = () => reject(request.error);
  });
}

export async function dbUpdateUserFields(
  username: string,
  updates: Partial<Pick<User, "name" | "mobile" | "email" | "isPro">>
): Promise<void> {
  const user = await dbGetUser(username);
  if (!user) throw new Error(`User "${username}" not found.`);
  const updated = { ...user, ...updates };
  await dbSaveUser(updated);
}

// ----------------------------------------------------
// 💳 PAYMENT MANAGEMENT (Reject / Update Status)
// ----------------------------------------------------
export async function dbUpdatePaymentStatus(paymentId: string, status: string): Promise<void> {
  if (isSupabaseConfigured) {
    try {
      const { error } = await supabase
        .from("payments")
        .update({ status })
        .eq("payment_id", paymentId);
      if (error) throw error;
      console.log(`☁️ Supabase: Payment "${paymentId}" status updated to "${status}".`);
    } catch (err) {
      console.error("❌ Supabase update payment status error:", err);
    }
  }
}

// ----------------------------------------------------
// 📦 PLANS CRUD
// ----------------------------------------------------
export async function dbGetPlans(): Promise<Plan[]> {
  if (isSupabaseConfigured) {
    try {
      const { data, error } = await supabase
        .from("plans")
        .select("*")
        .order("sort_order", { ascending: true });
      if (error) throw error;
      if (data) {
        return data.map((d: any) => ({
          id: d.id,
          name: d.name,
          price: d.price,
          billingPeriod: d.billing_period,
          features: d.features || [],
          isActive: d.is_active,
          description: d.description || "",
          highlighted: d.highlighted || false,
          color: d.color || "#f59e0b",
          maxReports: d.max_reports ?? 0,
          sortOrder: d.sort_order ?? 99,
          createdAt: d.created_at
        }));
      }
    } catch (err) {
      console.error("❌ Supabase get plans error:", err);
    }
  }
  // LocalStorage fallback
  try {
    const stored = localStorage.getItem("sheetcc_plans_local");
    if (stored) {
      const plans = JSON.parse(stored) as Plan[];
      if (plans.length > 0) return plans.sort((a, b) => (a.sortOrder ?? 99) - (b.sortOrder ?? 99));
    }
  } catch (e) {}
  // Return default plans as fallback if no Supabase
  return [
    {
      id: "basic",
      name: "Basic",
      price: 0,
      billingPeriod: "free" as const,
      features: [
        "3 free report generations",
        "Shopify & Shiprocket modes",
        "Interactive data mockup viewer",
        "100% client-side — no data stored"
      ],
      isActive: true,
      description: "Perfect for trying SheetCodeCrest on your first few exports",
      highlighted: false,
      color: "#3b82f6",
      maxReports: 3,
      sortOrder: 0
    },
    {
      id: "standard",
      name: "Standard",
      price: 1599,
      billingPeriod: "monthly" as const,
      features: [
        "Unlimited report generations",
        "All Starter features included",
        "AI Analyst (Avery) — conversational mode",
        "Saved report history & cloud sync",
        "Standard support"
      ],
      isActive: true,
      description: "For growing e-commerce brands running weekly reports",
      highlighted: true,
      color: "#faff69", // Electric Yellow ClickHouse style
      maxReports: 0,
      sortOrder: 1
    },
    {
      id: "premium",
      name: "Premium",
      price: 3999,
      billingPeriod: "monthly" as const,
      features: [
        "Everything in Standard",
        "Multi-user team access",
        "Custom column mapping rules",
        "Dedicated account manager",
        "API access (coming soon)"
      ],
      isActive: true,
      description: "For agencies, D2C brands, and teams needing multi-user and custom integrations",
      highlighted: false,
      color: "#a855f7",
      maxReports: 0,
      sortOrder: 2
    }
  ];
}

export async function dbSavePlan(plan: Plan): Promise<string> {
  if (isSupabaseConfigured) {
    try {
      const payload: any = {
        name: plan.name,
        price: plan.price,
        billing_period: plan.billingPeriod,
        features: plan.features,
        is_active: plan.isActive,
        description: plan.description || "",
        highlighted: plan.highlighted || false,
        color: plan.color || "#f59e0b",
        max_reports: plan.maxReports ?? 0,
        sort_order: plan.sortOrder ?? 99
      };
      if (plan.id) {
        payload.id = plan.id;
      }
      const { data, error } = await supabase
        .from("plans")
        .upsert(payload)
        .select("id")
        .single();
      if (error) throw error;
      console.log(`☁️ Supabase: Plan "${plan.name}" saved.`);
      return data?.id || plan.id || "";
    } catch (err) {
      console.error("❌ Supabase save plan error:", err);
    }
  }
  // LocalStorage fallback — persist plans locally when Supabase is not configured
  try {
    const stored = JSON.parse(localStorage.getItem("sheetcc_plans_local") || "[]") as Plan[];
    const idx = stored.findIndex(p => p.id === plan.id);
    const savedPlan = { ...plan, id: plan.id || Date.now().toString() };
    if (idx >= 0) stored[idx] = savedPlan;
    else stored.push(savedPlan);
    localStorage.setItem("sheetcc_plans_local", JSON.stringify(stored));
    return savedPlan.id!;
  } catch (e) {}
  return plan.id || Date.now().toString();
}

export async function dbDeletePlan(planId: string): Promise<void> {
  if (isSupabaseConfigured) {
    try {
      const { error } = await supabase
        .from("plans")
        .delete()
        .eq("id", planId);
      if (error) throw error;
      console.log(`☁️ Supabase: Plan "${planId}" deleted.`);
    } catch (err) {
      console.error("❌ Supabase delete plan error:", err);
    }
  }
}

// ----------------------------------------------------
// 📋 ADMIN ACTIVITY LOG
// ----------------------------------------------------
export async function dbLogAdminAction(action: string, performedBy: string, details?: string): Promise<void> {
  if (isSupabaseConfigured) {
    try {
      const { error } = await supabase
        .from("admin_logs")
        .insert({
          action,
          performed_by: performedBy,
          details: details || null
        });
      if (error) throw error;
    } catch (err) {
      console.error("❌ Supabase admin log error:", err);
    }
  }
}

export async function dbGetAdminLogs(): Promise<AdminLog[]> {
  if (isSupabaseConfigured) {
    try {
      const { data, error } = await supabase
        .from("admin_logs")
        .select("*")
        .order("created_at", { ascending: false })
        .limit(100);
      if (error) throw error;
      if (data) {
        return data.map((d: any) => ({
          id: d.id,
          action: d.action,
          performedBy: d.performed_by,
          details: d.details,
          createdAt: d.created_at
        }));
      }
    } catch (err) {
      console.error("❌ Supabase get admin logs error:", err);
    }
  }
  return [];
}
