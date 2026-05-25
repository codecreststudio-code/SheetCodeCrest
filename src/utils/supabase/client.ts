import { createBrowserClient } from "@supabase/ssr";

// Load credentials from environment variables, falling back to your live keys
const supabaseUrl = 
  import.meta.env.VITE_SUPABASE_URL || 
  import.meta.env.NEXT_PUBLIC_SUPABASE_URL || 
  "https://lodmofzpstnvkmnxaove.supabase.co";

const supabaseKey = 
  import.meta.env.VITE_SUPABASE_ANON_KEY || 
  import.meta.env.NEXT_PUBLIC_SUPABASE_PUBLISHABLE_KEY || 
  "sb_publishable_aUM6KI_e8zo9Bduv7H9JoA_xGIR5kNU";

export const createClient = () =>
  createBrowserClient(
    supabaseUrl!,
    supabaseKey!,
  );
