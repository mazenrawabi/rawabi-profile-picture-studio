import { PublicClientApplication } from "@azure/msal-browser";

export const tenantId = import.meta.env.VITE_ENTRA_TENANT_ID;
export const clientId = import.meta.env.VITE_ENTRA_CLIENT_ID;
export const authDisabled = import.meta.env.VITE_AUTH_DISABLED === "true";

export const isAuthConfigured = Boolean(
  tenantId &&
    clientId &&
    clientId !== "replace-with-application-client-id",
);

export const msal = isAuthConfigured
  ? new PublicClientApplication({
      auth: {
        clientId,
        authority: `https://login.microsoftonline.com/${tenantId}`,
        redirectUri: window.location.origin,
        postLogoutRedirectUri: window.location.origin,
        navigateToLoginRequestUrl: true,
      },
      cache: {
        cacheLocation: "sessionStorage",
        storeAuthStateInCookie: false,
      },
    })
  : null;

export const loginRequest = {
  scopes: ["openid", "profile", "email"],
  prompt: "select_account",
};
