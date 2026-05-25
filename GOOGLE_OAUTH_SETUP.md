# 🔐 Setting Up Your Live Google OAuth Client ID

To enable live Google Sign-In on **SheetCodeCrest**, you must configure your own Google OAuth Client ID. Because Google's authentication system strictly validates JavaScript origins and project ownership, a custom Client ID mapped to your specific domain/local port is required.

Follow this simple, step-by-step guide to get your live Client ID in under 2 minutes:

---

## 🛠️ Step-by-Step Setup Guide

### 1. Open Google Cloud Console
- Go to the [Google Cloud Console Credentials Page](https://console.cloud.google.com/apis/credentials).
- Sign in with your developer account (`codecreststudio@gmail.com` or your preferred workspace administrator account).
- If you don't have a project yet, click **Select a Project** at the top right, then click **New Project** and name it **SheetCodeCrest**.

---

### 2. Configure OAuth Consent Screen *(Required only once per project)*
If you haven't configured the consent screen for this project yet:
1. Click **OAuth Consent Screen** in the left sidebar.
2. Select **External** (available to any Google account) and click **Create**.
3. Fill in the required fields:
   - **App Name**: `SheetCodeCrest`
   - **User Support Email**: `codecreststudio@gmail.com`
   - **Developer Contact Email**: `codecreststudio@gmail.com`
4. Click **Save and Continue** through the remaining screens (Scopes, Test Users, Summary) without adding extra settings.
5. Go back to the OAuth Consent Screen dashboard and click **Publish App** under the publishing status to make it active.

---

### 3. Create OAuth Credentials
1. Click **Credentials** in the left sidebar.
2. Click the **+ Create Credentials** button at the top and select **OAuth client ID**.
3. Under **Application type**, select **Web application**.
4. Set the **Name** to `SheetCodeCrest Web Client`.

---

### 4. Authorize JavaScript Origins (CRITICAL)
Under the **Authorized JavaScript origins** section, click **+ Add URI** and add the exact addresses where your application runs:
- **Local Dev Server:**
  - `http://localhost:5173`
- **Optional Additional Local Ports:**
  - `http://localhost:3000`
- **Production Server (Vercel/Netlify):**
  - Add your live deployment URL (e.g., `https://sheetcodecrest.vercel.app`) if you have deployed it.

> [!WARNING]  
> Google Identity Services will reject authentication requests if the request originates from a domain or port not explicitly listed here. Do not add a trailing slash (e.g. `/`) at the end of the URIs.

---

### 5. Copy Your Client ID
1. Click **Create**.
2. A modal will pop up with your **Client ID** (it looks like a long string ending in `.apps.googleusercontent.com`).
3. Copy this **Client ID**.

---

## 🚀 How to Apply the ID to SheetCodeCrest

1. Open your code editor and go to [src/App.tsx](file:///c:/Users/new/Documents/Omer_DLS_Files/file_decode/src/App.tsx).
2. Find the constant `GOOGLE_CLIENT_ID` at line 35:
   ```typescript
   const GOOGLE_CLIENT_ID = "YOUR_NEW_CLIENT_ID.apps.googleusercontent.com";
   ```
3. Replace the placeholder string with your copied **Client ID** and save the file.
4. **Alternative:** Simply paste your new Client ID into our conversation chat, and I will instantly update the file, commit the change, and push it to your GitHub repository for you!
