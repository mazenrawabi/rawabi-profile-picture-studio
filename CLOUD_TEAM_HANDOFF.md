# Cloud Team handoff

## Requested tenant setup

Please create the following Microsoft Entra application registration:

- **Name:** Rawabi Profile Picture Studio
- **Supported account type:** Accounts in this organizational directory only (Rawabi Holding — single tenant)
- **Platform:** Single-page application (SPA)
- **Initial redirect URI:** `http://localhost:5173`
- **Production redirect URI:** Add the final HTTPS hosting origin after the static web application is created
- **Client secret:** None
- **API permissions:** No Microsoft Graph application permissions are required for the current editor

Please provide the **Application (client) ID** to Mazen so it can be added to the deployment configuration.

## Requested Azure access

Mazen currently sees zero Azure subscriptions and receives HTTP 401 when opening the Entra new-registration page. Please do one of the following:

1. Create an Azure resource group for this application and grant Mazen **Contributor** on that resource group; or
2. Have Cloud Team create and operate the hosting resource, then provide a deployment method.

For the pilot, Azure Static Web Apps Free is sufficient because authentication is performed by the single-tenant SPA registration. If policy requires the hosting layer itself to enforce tenant membership, use Azure Static Web Apps Standard with a custom Entra provider or an Azure App Service with built-in authentication.

## Repository and deployment settings

Cloud Team remains the owner of the Azure resources, Entra registration, and deployment credential. Mazen can maintain the application source in the connected GitHub repository.

Use these Azure Static Web Apps build settings:

- **App location:** `/`
- **API location:** Leave blank
- **Build command:** `npm run build`
- **Output location:** `dist`
- **Node version:** 20 LTS or later

The production build requires:

- `VITE_ENTRA_TENANT_ID=c904431e-5720-4d00-8188-8504a3c2facb`
- `VITE_ENTRA_CLIENT_ID=<Application (client) ID supplied by IT>`
- `VITE_AUTH_DISABLED=false`

The Azure Static Web Apps deployment token should be stored as a GitHub Actions secret and administered by IT. Do not commit it to this repository.

## Pilot deployment checklist

1. Cloud Team creates the single-tenant Entra SPA registration with no client secret.
2. Cloud Team creates the Azure Static Web App and connects this repository and its `main` branch.
3. The workflow is configured with the build settings and environment variables above.
4. Cloud Team adds the final Azure Static Web Apps origin to the SPA redirect URIs.
5. Deploy and confirm that an authorized Rawabi user can sign in.
6. Upload a representative headshot, test background removal and manual restore around white clothing, and export a 605 × 688 PNG.
7. Confirm that no photo is sent to a backend or retained after the browser session.
8. Cloud Team records operational ownership and the rollback procedure before wider release.

## Security model

- The SPA authority is locked to tenant `c904431e-5720-4d00-8188-8504a3c2facb`.
- Employee photos and background removal stay in the browser.
- No photo, token, or employee record is stored by the current release.
- The future Submit → Approval → Directory Update capability must use a protected backend that validates tokens and isolates privileged Microsoft Graph permissions from the browser.
