import {describe, it, expect, beforeEach, vi, afterEach} from "vitest";
import {InteractionRequiredAuthError} from "@azure/msal-browser";

// We mock the @azure/msal-browser module so unit tests do not touch a real
// Entra endpoint. Only the surface src/auth.js uses is stubbed.
const mockAcquireTokenSilent = vi.fn();
const mockAcquireTokenPopup = vi.fn();
const mockGetAllAccounts = vi.fn();
const mockCreateNestable = vi.fn();

vi.mock("@azure/msal-browser", async (importOriginal) => {
  const actual = await importOriginal();
  return {
    ...actual,
    createNestablePublicClientApplication: (...args) =>
      mockCreateNestable(...args),
  };
});

const CLIENT_ID = "00000000-0000-0000-0000-000000000001";
const TENANT_ID = "00000000-0000-0000-0000-000000000002";
const SCOPE = "api://librechat-app-id/access_as_user";

function setEnv(overrides = {}) {
  window.__ENV = {
    ENTRA_CLIENT_ID: CLIENT_ID,
    ENTRA_TENANT_ID: TENANT_ID,
    ENTRA_API_SCOPE: SCOPE,
    ...overrides,
  };
}

function makeClient({accounts = []} = {}) {
  return {
    acquireTokenSilent: mockAcquireTokenSilent,
    acquireTokenPopup: mockAcquireTokenPopup,
    getAllAccounts: mockGetAllAccounts.mockReturnValue(accounts),
  };
}

describe("auth.js", () => {
  let auth;

  beforeEach(async () => {
    vi.clearAllMocks();
    mockCreateNestable.mockResolvedValue(makeClient());
    window.__ENV = {};
    globalThis.Office = {
      context: {
        requirements: {isSetSupported: vi.fn(() => true)},
      },
    };
    // Re-import a fresh copy of the module for each test so the lazily
    // created MSAL instance does not leak between tests.
    vi.resetModules();
    auth = await import("../src/auth.js");
  });

  afterEach(() => {
    delete window.__ENV;
    delete globalThis.Office;
  });

  describe("isNaaSupported", () => {
    it("returns true when the NestedAppAuth 1.1 requirement set is supported", () => {
      globalThis.Office.context.requirements.isSetSupported.mockReturnValue(true);
      expect(auth.isNaaSupported()).toBe(true);
      expect(
        globalThis.Office.context.requirements.isSetSupported,
      ).toHaveBeenCalledWith("NestedAppAuth", "1.1");
    });

    it("returns false when the NestedAppAuth 1.1 requirement set is not supported", () => {
      globalThis.Office.context.requirements.isSetSupported.mockReturnValue(false);
      expect(auth.isNaaSupported()).toBe(false);
    });

    it("returns false instead of throwing when the requirements API is unavailable", () => {
      globalThis.Office = {};
      expect(auth.isNaaSupported()).toBe(false);
    });
  });

  describe("initMsal", () => {
    it("builds the nestable client with single-tenant authority and localStorage cache", async () => {
      setEnv();
      await auth.initMsal();
      expect(mockCreateNestable).toHaveBeenCalledTimes(1);
      const config = mockCreateNestable.mock.calls[0][0];
      expect(config.auth.clientId).toBe(CLIENT_ID);
      expect(config.auth.authority).toBe(
        `https://login.microsoftonline.com/${TENANT_ID}`,
      );
      expect(config.auth.authority).not.toContain("common");
      expect(config.cache.cacheLocation).toBe("localStorage");
    });

    it("is lazy and memoized — only creates the client once", async () => {
      setEnv();
      await auth.initMsal();
      await auth.initMsal();
      await auth.initMsal();
      expect(mockCreateNestable).toHaveBeenCalledTimes(1);
    });

    it("throws a descriptive error when ENTRA_CLIENT_ID is missing", async () => {
      setEnv({ENTRA_CLIENT_ID: ""});
      await expect(auth.initMsal()).rejects.toThrow(/ENTRA_CLIENT_ID/);
    });

    it("throws a descriptive error when ENTRA_TENANT_ID is missing", async () => {
      setEnv({ENTRA_TENANT_ID: ""});
      await expect(auth.initMsal()).rejects.toThrow(/ENTRA_TENANT_ID/);
    });
  });

  describe("getAccessToken", () => {
    it("returns the silent token when acquisition succeeds", async () => {
      setEnv();
      mockAcquireTokenSilent.mockResolvedValue({accessToken: "silent-token"});
      const token = await auth.getAccessToken({allowInteractive: true});
      expect(token).toBe("silent-token");
      expect(mockAcquireTokenPopup).not.toHaveBeenCalled();
    });

    it("requests the configured Entra scope", async () => {
      setEnv();
      mockAcquireTokenSilent.mockResolvedValue({accessToken: "tok"});
      await auth.getAccessToken({allowInteractive: true});
      const request = mockAcquireTokenSilent.mock.calls[0][0];
      expect(request.scopes).toEqual([SCOPE]);
    });

    it("escalates to a popup on InteractionRequiredAuthError when allowInteractive is true", async () => {
      setEnv();
      mockAcquireTokenSilent.mockRejectedValue(
        new InteractionRequiredAuthError("interaction_required"),
      );
      mockAcquireTokenPopup.mockResolvedValue({accessToken: "popup-token"});
      const token = await auth.getAccessToken({allowInteractive: true});
      expect(token).toBe("popup-token");
      expect(mockAcquireTokenPopup).toHaveBeenCalledTimes(1);
    });

    it("rethrows InteractionRequiredAuthError without a popup when allowInteractive is false", async () => {
      setEnv();
      const err = new InteractionRequiredAuthError("interaction_required");
      mockAcquireTokenSilent.mockRejectedValue(err);
      await expect(auth.getAccessToken({allowInteractive: false})).rejects.toBe(
        err,
      );
      expect(mockAcquireTokenPopup).not.toHaveBeenCalled();
    });

    it("rethrows non-interaction errors without attempting a popup", async () => {
      setEnv();
      const err = new Error("network boom");
      mockAcquireTokenSilent.mockRejectedValue(err);
      await expect(auth.getAccessToken({allowInteractive: true})).rejects.toBe(
        err,
      );
      expect(mockAcquireTokenPopup).not.toHaveBeenCalled();
    });

    it("throws a descriptive error when ENTRA_API_SCOPE is missing", async () => {
      setEnv({ENTRA_API_SCOPE: ""});
      await expect(auth.getAccessToken({allowInteractive: true})).rejects.toThrow(
        /ENTRA_API_SCOPE/,
      );
    });
  });

  describe("getSignedInUser", () => {
    it("returns the first cached account's username", async () => {
      setEnv();
      mockCreateNestable.mockResolvedValue(
        makeClient({accounts: [{username: "user@nos.pt"}]}),
      );
      const user = await auth.getSignedInUser();
      expect(user).toBe("user@nos.pt");
    });

    it("returns null when no account is cached", async () => {
      setEnv();
      mockCreateNestable.mockResolvedValue(makeClient({accounts: []}));
      const user = await auth.getSignedInUser();
      expect(user).toBeNull();
    });
  });
});
