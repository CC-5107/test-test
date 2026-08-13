// Server-only authentication for direct Google Sheets API access.
//
// This version is independent of Lovable's Google connector. It authenticates
// directly with Google using a service account and an OAuth 2.0 JWT assertion.
//
// Required server/Vercel environment variables:
//   GOOGLE_SERVICE_ACCOUNT_EMAIL   — service-account client_email
//   GOOGLE_PRIVATE_KEY             — service-account private_key
//
// Optional:
//   GOOGLE_SPREADSHEET_ID           — overrides the built-in project sheet ID
//
// IMPORTANT: share the target Google Sheet with GOOGLE_SERVICE_ACCOUNT_EMAIL
// as an Editor if the website needs to write to it.

export const SHEETS_API_BASE = "https://sheets.googleapis.com/v4";

const GOOGLE_TOKEN_URL = "https://oauth2.googleapis.com/token";
const SHEETS_SCOPE = "https://www.googleapis.com/auth/spreadsheets";
const TOKEN_SKEW_MS = 60_000;

type CachedAccessToken = {
  token: string;
  expiresAt: number;
};

let cachedAccessToken: CachedAccessToken | null = null;
let accessTokenRequest: Promise<CachedAccessToken> | null = null;
let cachedSigningKey: CryptoKey | null = null;
let cachedPrivateKeyText = "";

export function getServerEnv(name: string): string {
  const value = (globalThis as any)?.process?.env?.[name];
  return typeof value === "string" ? value.trim() : "";
}

function bytesToBase64(bytes: Uint8Array): string {
  let binary = "";
  const chunkSize = 0x8000;
  for (let i = 0; i < bytes.length; i += chunkSize) {
    binary += String.fromCharCode(...bytes.subarray(i, i + chunkSize));
  }
  return btoa(binary);
}

function base64UrlBytes(bytes: Uint8Array): string {
  return bytesToBase64(bytes)
    .replace(/=/g, "")
    .replace(/\+/g, "-")
    .replace(/\//g, "_");
}

function base64UrlText(value: string): string {
  return base64UrlBytes(new TextEncoder().encode(value));
}

function base64ToBytes(value: string): Uint8Array {
  const binary = atob(value.replace(/\s+/g, ""));
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
  return bytes;
}

function getServiceAccountEmail(): string {
  const value = getServerEnv("GOOGLE_SERVICE_ACCOUNT_EMAIL");
  if (!value) {
    throw new Error(
      "Google Sheets service account is not configured: missing GOOGLE_SERVICE_ACCOUNT_EMAIL",
    );
  }
  return value;
}

function getPrivateKey(): string {
  let value = getServerEnv("GOOGLE_PRIVATE_KEY");
  if (!value) {
    throw new Error(
      "Google Sheets service account is not configured: missing GOOGLE_PRIVATE_KEY",
    );
  }

  // Vercel accepts multiline values, but keys are also commonly pasted with
  // literal "\\n" characters. Support both forms.
  value = value.replace(/\\n/g, "\n");

  // Be forgiving if the whole key was pasted with surrounding quotes.
  if (
    (value.startsWith('"') && value.endsWith('"')) ||
    (value.startsWith("'") && value.endsWith("'"))
  ) {
    value = value.slice(1, -1).replace(/\\n/g, "\n");
  }

  if (!value.includes("BEGIN PRIVATE KEY")) {
    throw new Error(
      "GOOGLE_PRIVATE_KEY does not look like a Google service-account private key",
    );
  }
  return value;
}

async function getSigningKey(): Promise<CryptoKey> {
  const privateKeyText = getPrivateKey();
  if (cachedSigningKey && cachedPrivateKeyText === privateKeyText) {
    return cachedSigningKey;
  }

  const derBase64 = privateKeyText
    .replace(/-----BEGIN PRIVATE KEY-----/g, "")
    .replace(/-----END PRIVATE KEY-----/g, "")
    .replace(/\s+/g, "");

  const key = await crypto.subtle.importKey(
    "pkcs8",
    base64ToBytes(derBase64),
    { name: "RSASSA-PKCS1-v1_5", hash: "SHA-256" },
    false,
    ["sign"],
  );

  cachedSigningKey = key;
  cachedPrivateKeyText = privateKeyText;
  return key;
}

async function buildSignedJwt(): Promise<string> {
  const now = Math.floor(Date.now() / 1000);
  const header = base64UrlText(JSON.stringify({ alg: "RS256", typ: "JWT" }));
  const claims = base64UrlText(
    JSON.stringify({
      iss: getServiceAccountEmail(),
      scope: SHEETS_SCOPE,
      aud: GOOGLE_TOKEN_URL,
      iat: now,
      exp: now + 3600,
    }),
  );
  const unsignedToken = `${header}.${claims}`;
  const signature = await crypto.subtle.sign(
    "RSASSA-PKCS1-v1_5",
    await getSigningKey(),
    new TextEncoder().encode(unsignedToken),
  );
  return `${unsignedToken}.${base64UrlBytes(new Uint8Array(signature))}`;
}

async function requestAccessToken(): Promise<CachedAccessToken> {
  const assertion = await buildSignedJwt();
  const body = new URLSearchParams({
    grant_type: "urn:ietf:params:oauth:grant-type:jwt-bearer",
    assertion,
  });

  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), 15_000);
  try {
    const response = await fetch(GOOGLE_TOKEN_URL, {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: body.toString(),
      signal: controller.signal,
    });

    const raw = await response.text();
    let payload: any = {};
    try {
      payload = raw ? JSON.parse(raw) : {};
    } catch {
      // Keep raw response for the error below.
    }

    if (!response.ok || !payload?.access_token) {
      throw new Error(
        `Google OAuth token request failed [${response.status}]: ${payload?.error_description || payload?.error || raw || "Unknown error"}`,
      );
    }

    const expiresInSeconds = Number(payload.expires_in || 3600);
    return {
      token: String(payload.access_token),
      expiresAt: Date.now() + Math.max(60, expiresInSeconds) * 1000,
    };
  } finally {
    clearTimeout(timeout);
  }
}

export async function getSheetsAccessToken(): Promise<string> {
  if (
    cachedAccessToken &&
    cachedAccessToken.expiresAt - TOKEN_SKEW_MS > Date.now()
  ) {
    return cachedAccessToken.token;
  }

  if (!accessTokenRequest) {
    accessTokenRequest = requestAccessToken()
      .then((token) => {
        cachedAccessToken = token;
        return token;
      })
      .finally(() => {
        accessTokenRequest = null;
      });
  }

  return (await accessTokenRequest).token;
}

export async function sheetsAuthHeaders(): Promise<Record<string, string>> {
  const accessToken = await getSheetsAccessToken();
  return {
    Authorization: `Bearer ${accessToken}`,
    "Content-Type": "application/json",
  };
}

export function invalidateSheetsAccessToken() {
  cachedAccessToken = null;
}
