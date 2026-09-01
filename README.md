# Rawabi Profile Picture Studio

An Entra sign-in-gated React application for creating Rawabi employee profile pictures at exactly 605 × 688 pixels.

## What stays on the employee device

- The uploaded headshot
- Background removal processing
- Restore/erase mask edits
- The composed PNG until the employee downloads it

The current version has no submission backend and does not upload employee photos.

## Configuration

Copy `.env.example` to `.env.local` and set:

```text
VITE_ENTRA_TENANT_ID=c904431e-5720-4d00-8188-8504a3c2facb
VITE_ENTRA_CLIENT_ID=<single-tenant SPA application client ID>
VITE_AUTH_DISABLED=false
```

`VITE_AUTH_DISABLED=true` is only for local UI testing. Never use it for a deployment.

## Commands

```text
npm install
npm run dev
npm run build
```

The production build is written to `dist/`.

## Azure Static Web Apps deployment

Use these build settings when connecting this repository:

| Setting | Value |
| --- | --- |
| App location | `/` |
| API location | *(leave blank)* |
| Build command | `npm run build` |
| Output location | `dist` |

Configure these build-time environment variables in the deployment workflow:

```text
VITE_ENTRA_TENANT_ID=c904431e-5720-4d00-8188-8504a3c2facb
VITE_ENTRA_CLIENT_ID=<Application (client) ID supplied by IT>
VITE_AUTH_DISABLED=false
```

After Azure assigns the production hostname, IT must add its origin as a **Single-page application** redirect URI in the Entra registration, for example `https://<app-name>.azurestaticapps.net`.

See [CLOUD_TEAM_HANDOFF.md](./CLOUD_TEAM_HANDOFF.md) for the ownership model and deployment checklist.

## Entra registration

Use a single-tenant SPA registration and add each deployed origin as a redirect URI. The application requests only OpenID identity scopes (`openid`, `profile`, and `email`) and does not use a client secret.

For the future approval and directory-update workflow, add a protected backend. That backend must validate Microsoft identity tokens and hold any Microsoft Graph application permissions; Graph credentials must never be placed in this browser application.
