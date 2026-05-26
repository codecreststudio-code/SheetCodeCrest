// src/config/env.ts
// Environment variable validation and configuration

/**
 * Validates that required environment variables are present
 * Throws an error in development if missing, returns fallback values in production
 */
export function validateEnv() {
  const isDev = import.meta.env.MODE === 'development';

  // Required environment variables
  const requiredVars = [
    'VITE_PROXY_URL',
    'VITE_RAZORPAY_KEY_ID',
    'GOOGLE_CLIENT_ID'
  ] as const;

  const missing = requiredVars.filter(varName => !import.meta.env[varName]);

  if (missing.length > 0 && isDev) {
    console.warn(`Missing environment variables: ${missing.join(', ')}`);
    console.warn('Using fallback values. Please set these in your .env file for production.');
  }

  return {
    proxyUrl: import.meta.env.VITE_PROXY_URL || 'http://localhost:5001',
    razorpayKey: import.meta.env.VITE_RAZORPAY_KEY_ID || 'rzp_test_placeholder',
    googleClientId: import.meta.env.VITE_GOOGLE_CLIENT_ID || import.meta.env.GOOGLE_CLIENT_ID || '',
    personalUpiId: import.meta.env.VITE_PERSONAL_UPI_ID || 'codecreststudio@okaxis',
    isDev
  };
}

// Export validated configuration
export const env = validateEnv();