/* global Office */
import {
  createNestablePublicClientApplication,
  InteractionRequiredAuthError,
} from "@azure/msal-browser";

// Single owner of all identity concerns for the add-in. Everything that needs
// to know *who the user is* or *how to get a token* goes through this module.
//
// Auth model: Entra ID SSO via Nested App Authentication (NAA). The Outlook
// host brokers the token, so most requests are fully silent. There is no
// API-key fallback — non-NAA clients are blocked by the caller (task pane).

// Read lazily so the values resolve at call time (window.__ENV is injected at
// runtime by /env-config.js, after module evaluation).
function env(name) {
  const ENV = (typeof window !== "undefined" && window.__ENV) || {};
  return ENV[name] || "";
}

let msalInstance = null;
let msalInitPromise = null;

function requireConfig(value, name) {
  if (!value) {
    throw new Error(
      `Missing ${name}. Set it in your environment / .env before building.`,
    );
  }
  return value;
}

/**
 * True when the current Office host supports Nested App Authentication.
 * Callers should gate all auth on this — non-NAA clients get a blocking
 * message, not a silent failure.
 */
function isNaaSupported() {
  try {
    return Boolean(
      Office.context.requirements.isSetSupported("NestedAppAuth", "1.1"),
    );
  } catch {
    return false;
  }
}

/**
 * Lazily builds (and memoizes) the nestable public client. Concurrent callers
 * share the same in-flight creation promise.
 */
async function initMsal() {
  if (msalInstance) return msalInstance;
  if (!msalInitPromise) {
    const clientId = requireConfig(env("ENTRA_CLIENT_ID"), "ENTRA_CLIENT_ID");
    const tenantId = requireConfig(env("ENTRA_TENANT_ID"), "ENTRA_TENANT_ID");
    const config = {
      auth: {
        clientId,
        // Single-tenant: the tenant's own directory, not "common".
        authority: `https://login.microsoftonline.com/${tenantId}`,
      },
      cache: {
        // Load-bearing: taskpane.html and commands.html are same-origin, so a
        // localStorage cache is what lets ribbon commands reuse the token the
        // task pane acquired. Do NOT change to sessionStorage.
        cacheLocation: "localStorage",
      },
    };
    msalInitPromise = createNestablePublicClientApplication(config);
  }
  msalInstance = await msalInitPromise;
  return msalInstance;
}

/**
 * Acquire an access token for LibreChat.
 *
 * Tries the silent path first (MSAL serves from cache and refreshes
 * underneath). When interaction is required, it escalates to a popup ONLY if
 * `allowInteractive` is true — the ribbon commands pass false because they
 * have no UI surface to host a prompt.
 *
 * @param {{allowInteractive: boolean}} options
 * @returns {Promise<string>} the access token
 */
async function getAccessToken({allowInteractive} = {allowInteractive: false}) {
  const scope = requireConfig(env("ENTRA_API_SCOPE"), "ENTRA_API_SCOPE");
  const client = await initMsal();
  const activeAccount = client.getActiveAccount?.();
  const request = {scopes: [scope]};
  if (activeAccount) request.account = activeAccount;

  try {
    const result = await client.acquireTokenSilent(request);
    if (result.account) client.setActiveAccount(result.account);
    return result.accessToken;
  } catch (err) {
    if (err instanceof InteractionRequiredAuthError && allowInteractive) {
      const result = await client.acquireTokenPopup(request);
      if (result.account) client.setActiveAccount(result.account);
      return result.accessToken;
    }
    throw err;
  }
}

/**
 * The signed-in account's username (for display in Settings), or null when
 * no account is cached yet.
 *
 * @returns {Promise<string|null>}
 */
async function getSignedInUser() {
  const client = await initMsal();
  const account =
    client.getActiveAccount?.() || (client.getAllAccounts() || [])[0];
  return account ? account.username || null : null;
}

async function signIn() {
  const scope = requireConfig(env("ENTRA_API_SCOPE"), "ENTRA_API_SCOPE");
  const client = await initMsal();
  const result = await client.acquireTokenPopup({
    scopes: [scope],
    prompt: "login",
  });
  if (result.account) client.setActiveAccount(result.account);
  return result.accessToken;
}

async function signOut() {
  const client = await initMsal();
  // Nested App Auth does not support logoutPopup, logoutRedirect, or
  // clearCache. Clear the active account so this task pane stops using it.
  client.setActiveAccount(null);
}

export {
  isNaaSupported,
  initMsal,
  getAccessToken,
  getSignedInUser,
  signIn,
  signOut,
};
