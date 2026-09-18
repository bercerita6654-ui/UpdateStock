import { initializeApp, getApps, getApp } from 'firebase/app';
import {
  getAuth,
  signInWithPopup,
  GoogleAuthProvider,
  onAuthStateChanged,
  signOut,
  User,
} from 'firebase/auth';
import firebaseConfig from '../../firebase-applet-config.json';

const app = getApps().length === 0 ? initializeApp(firebaseConfig) : getApp();
export const auth = getAuth(app);

export const SCOPES = [
  'https://www.googleapis.com/auth/spreadsheets',
  'https://www.googleapis.com/auth/drive.file',
  'https://www.googleapis.com/auth/drive.readonly',
  'https://www.googleapis.com/auth/drive',
];

const TOKEN_STORAGE_KEY = 'shopee_google_access_token';
const TOKEN_EXPIRY_KEY = 'shopee_google_token_expiry';
const LAST_EMAIL_KEY = 'shopee_google_last_email';

// Helper to get stored token if not expired (tokens usually valid for 1 hour, check with 5-min margin)
function getStoredToken(): string | null {
  try {
    const token = localStorage.getItem(TOKEN_STORAGE_KEY);
    const expiryStr = localStorage.getItem(TOKEN_EXPIRY_KEY);
    if (token && expiryStr) {
      const expiry = parseInt(expiryStr, 10);
      if (Date.now() < expiry) {
        return token;
      }
    }
  } catch (e) {
    console.warn('Could not read stored auth token:', e);
  }
  return null;
}

function saveStoredToken(token: string, email?: string | null) {
  try {
    // Save token with 55-minute expiration
    const expiry = Date.now() + 55 * 60 * 1000;
    localStorage.setItem(TOKEN_STORAGE_KEY, token);
    localStorage.setItem(TOKEN_EXPIRY_KEY, String(expiry));
    if (email) {
      localStorage.setItem(LAST_EMAIL_KEY, email);
    }
  } catch (e) {
    console.warn('Could not save auth token:', e);
  }
}

function clearStoredToken() {
  try {
    localStorage.removeItem(TOKEN_STORAGE_KEY);
    localStorage.removeItem(TOKEN_EXPIRY_KEY);
  } catch (e) {
    console.warn('Could not clear auth token:', e);
  }
}

const provider = new GoogleAuthProvider();
SCOPES.forEach((scope) => {
  provider.addScope(scope);
});

// Configure provider with automatic account selection if available
function configureProvider(emailHint?: string | null) {
  const customParams: Record<string, string> = {
    include_granted_scopes: 'true',
  };
  const targetEmail = emailHint || localStorage.getItem(LAST_EMAIL_KEY);
  if (targetEmail) {
    customParams.login_hint = targetEmail;
  }
  provider.setCustomParameters(customParams);
}

configureProvider();

let isSigningIn = false;
let cachedAccessToken: string | null = getStoredToken();
let ongoingSignInPromise: Promise<{ user: User; accessToken: string } | null> | null = null;

export const initAuth = (
  onAuthSuccess?: (user: User, token: string) => void,
  onAuthFailure?: () => void
) => {
  return onAuthStateChanged(auth, async (user: User | null) => {
    if (user) {
      const activeToken = cachedAccessToken || getStoredToken();
      if (activeToken) {
        cachedAccessToken = activeToken;
        if (user.email) {
          localStorage.setItem(LAST_EMAIL_KEY, user.email);
        }
        if (onAuthSuccess) onAuthSuccess(user, activeToken);
      } else if (!isSigningIn) {
        // Auth state is active in Firebase, but OAuth token needs refreshment
        if (onAuthFailure) onAuthFailure();
      }
    } else {
      cachedAccessToken = null;
      clearStoredToken();
      if (onAuthFailure) onAuthFailure();
    }
  });
};

export const googleSignIn = async (): Promise<{ user: User; accessToken: string } | null> => {
  // If a sign-in popup is already open/in-flight, reuse the pending promise to avoid auth/cancelled-popup-request
  if (ongoingSignInPromise) {
    return ongoingSignInPromise;
  }

  ongoingSignInPromise = (async () => {
    try {
      isSigningIn = true;
      // Configure with login hint if user previously logged in
      const lastEmail = localStorage.getItem(LAST_EMAIL_KEY);
      configureProvider(lastEmail);

      const result = await signInWithPopup(auth, provider);
      const credential = GoogleAuthProvider.credentialFromResult(result);
      if (!credential?.accessToken) {
        throw new Error('Gagal mendapatkan token akses dari Google. Pastikan izin akses telah disetujui.');
      }
      cachedAccessToken = credential.accessToken;
      saveStoredToken(cachedAccessToken, result.user.email);

      return { user: result.user, accessToken: cachedAccessToken };
    } catch (error: any) {
      // Handle benign popup cancellations gracefully without throwing unhandled exceptions
      if (
        error?.code === 'auth/cancelled-popup-request' ||
        error?.code === 'auth/popup-closed-by-user' ||
        error?.message?.includes('cancelled-popup-request') ||
        error?.message?.includes('popup-closed-by-user')
      ) {
        console.warn('Google sign-in popup was cancelled or closed by user.');
        return null;
      }

      if (error?.code === 'auth/popup-blocked') {
        throw new Error('Popup login diblokir oleh browser. Harap izinkan popup di browser Anda.');
      }

      console.error('Sign in error:', error);
      throw error;
    } finally {
      isSigningIn = false;
      ongoingSignInPromise = null;
    }
  })();

  return ongoingSignInPromise;
};

export const getAccessToken = async (): Promise<string | null> => {
  return cachedAccessToken || getStoredToken();
};

export const setCachedToken = (token: string | null) => {
  cachedAccessToken = token;
  if (token) {
    saveStoredToken(token);
  } else {
    clearStoredToken();
  }
};

export const logout = async () => {
  await signOut(auth);
  cachedAccessToken = null;
  clearStoredToken();
};

