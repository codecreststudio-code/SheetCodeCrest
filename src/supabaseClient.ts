import { createClient } from "@supabase/supabase-js";

// Load configuration from Vite or Next.js environment variables, or fallback to your live credentials
const supabaseUrl = 
  import.meta.env.VITE_SUPABASE_URL || 
  import.meta.env.NEXT_PUBLIC_SUPABASE_URL || 
  "https://lodmofzpstnvkmnxaove.supabase.co";

const supabaseAnonKey = 
  import.meta.env.VITE_SUPABASE_ANON_KEY || 
  import.meta.env.NEXT_PUBLIC_SUPABASE_PUBLISHABLE_KEY || 
  "sb_publishable_aUM6KI_e8zo9Bduv7H9JoA_xGIR5kNU";

// Detect if keys are configured
export const isSupabaseConfigured = 
  supabaseUrl && 
  supabaseAnonKey && 
  !supabaseUrl.includes("your-project-id") && 
  !supabaseAnonKey.includes("your-supabase-anon-key");

if (!isSupabaseConfigured) {
  console.warn(
    "⚠️ Supabase is not fully configured yet. Running SheetCodeCrest in local offline mode (using IndexedDB fallback)."
  );
} else {
  console.log("☁️ Supabase Cloud Database successfully initialized for SheetCodeCrest!");
}

export const supabase = isSupabaseConfigured 
  ? createClient(supabaseUrl, supabaseAnonKey) 
  : (null as any);
