# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Commands

```bash
# Development (with Teams Toolkit local config)
npm run dev:teamsfx

# Development (standalone, rebuilds frontend then watches backend)
npm run dev

# Production build (backend via tsup + frontend via vite)
npm run build

# Frontend only
npm run build:frontend

# Clean build artifacts
npm run clean
```

There are no lint or test commands configured in this project.

## Architecture

**DDS Coffee Talk** is a Microsoft Teams Tab application that creates random coffee meeting pairings between Teams channel members and schedules the meetings via Microsoft Graph API.

### Stack

- **Frontend:** React 19 + TypeScript, bundled by Vite into `dist/client/`
- **Backend:** Node.js/Express via `@microsoft/teams.apps`, compiled by tsup into `dist/`
- **Auth:** Azure AD `ClientSecretCredential` (`AAD_APP_TENANT_ID`, `AAD_APP_CLIENT_ID`, `AAD_APP_CLIENT_SECRET`)
- **Graph SDK:** `@microsoft/microsoft-graph-client` for calendar and channel operations

### Source layout

```
src/
  index.ts          Backend entry — Express server + REST API endpoints
  Tab/
    App.tsx         Main React component (all UI state lives here)
    client.tsx      React app entry point
    App.css         DEHN design system (CSS custom properties)
```

### Backend API (`src/index.ts`)

Four POST endpoints:

| Endpoint | Purpose |
|---|---|
| `/api/status` | Health check + Graph permission validation |
| `/api/channel-members` | Fetch Teams channel members via Graph |
| `/api/random-meetings/now` | Find next free slot and create meetings immediately |
| `/api/random-meetings/at-time` | Schedule meetings at a specific datetime |

Key functions:
- `createRandomGroups()` — shuffles members and chunks them into groups
- `areParticipantsFree()` — calls Graph `getSchedule` to check calendar availability
- `findAvailableSlot()` — iterates 30-min slots within business hours (9–17, weekdays) over a 7-day window
- `createCalendarEvent()` — creates a Teams Online Meeting; first member is organizer, rest are required attendees

### Frontend (`src/Tab/App.tsx`)

Pure React hooks (no external state lib). On load it calls `app.getContext()` from the Teams JS SDK to extract `teamId`/`channelId`/`userObjectId`, then populates member lists and meeting results via the backend API. User-configurable parameters: number of meetings, group size, duration, timezone.

### Configuration

- `vite.config.js` — base path is `/tabs/home`
- `nodemon.json` — watches `src/**/*.{ts,tsx}`, runs `npm run build` on change
- `tsup.config.js` — builds CJS + ESM with DTS and source maps
- `m365agents.yml` — Microsoft 365 Toolkit provisioning (Azure App Service + Teams manifest)
- `.localConfigs` — local dev overrides (port 3978, SSL certs, Azure AD creds); not committed

### Environment variables

| Variable | Purpose |
|---|---|
| `AAD_APP_TENANT_ID` | Azure AD tenant |
| `AAD_APP_CLIENT_ID` | App registration client ID |
| `AAD_APP_CLIENT_SECRET` | App registration secret |
| `DEFAULT_TIME_ZONE` | Defaults to `Europe/Madrid` |
