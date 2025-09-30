# Lumina Sheets – Rebuilt Core

This repository now ships with a **ground-up rewrite** of the Lumina Sheets
workspace. The previous collection of disjoint Apps Script files has been
replaced by a modular runtime that treats Google Sheets as a database, delivers
fast cached reads, and exposes a clean API to the HTML client.

## Highlights

- **Deterministic bootstrap** – every required sheet is ensured with the
  expected headers before the first request executes. The schema mirrors the
  original `MainUtilities` constants so existing data continues to work.
- **Token-based authentication** – sign-in issues short-lived or remember-me
  tokens that are stored inside the `Sessions` sheet with user agent and IP
  metadata.
- **Cache-aware repositories** – heavy lookups against `Users`, `Campaigns`,
  `Pages`, and `Notifications` are memoised through `CacheService` to keep the
  UI snappy even for large datasets.
- **Compact HTML client** – the new `V2_Login.html` and `V2_AppShell.html`
  provide a modern authenticated experience that talks to the Apps Script API
  (`google.script.run`) without relying on legacy global functions.

## Project layout

```
Lumina-Sheets/
├── Code.js             # V2 runtime (authentication, routing, sheet access)
├── V2_Login.html       # Standalone login screen
├── V2_AppShell.html    # Authenticated workspace shell + dashboard view
├── appsscript.json     # Manifest (V8 runtime + web-app settings)
└── ...                 # Legacy HTML/GS files kept for reference
```

The old services remain in the repository for historical reference but are no
longer required. The runtime entry points (`doGet`, `doPost`, `appLogin`, etc.)
all delegate to the new `LuminaAppV2` implementation inside `Code.js`.

## Getting started

1. Deploy the Apps Script project (`clasp push` or manual upload) so the new
   files replace the existing script project.
2. Make sure your Google Sheet already contains the historical data. During the
   first run the runtime will automatically ensure all core sheets exist with the
   correct headers.
3. Open the web app URL. You should see the refreshed login page. Sign in using
   an existing account (passwords must match the original SHA-256 + Base64 hash
   stored in the `Users` sheet).
4. Upon successful login the token is cached in `localStorage` and the workspace
   renders the dashboard view with campaign memberships and unread notifications.

## Client API

The HTML client communicates with Apps Script through the following public
functions:

| Function        | Description                                                   |
| --------------- | ------------------------------------------------------------- |
| `appLogin`      | Validates credentials and issues a session token.             |
| `appLogout`     | Revokes the provided token and clears it from storage.        |
| `appBootstrap`  | Returns the authenticated user profile, campaign list, and    |
|                 | navigation metadata.                                          |
| `appDashboard`  | Provides dashboard aggregates including unread notifications. |
| `appNotifications` | Retrieves the unread notifications list.                   |

Each method expects an object containing the `token` (except `appLogin`) and
returns a JSON-serialisable payload that the front-end renders.

## Extending the workspace

The rebuilt runtime separates concerns so additional features are easy to add:

- **New tables** – add the sheet name + header array to the `SHEETS`/`HEADERS`
  objects, call `ensureSheetWithHeaders`, and build a repository helper using the
  existing patterns (`readTable`, `appendRow`, etc.).
- **Additional APIs** – create a function that receives a payload, calls
  `requireSession(token)` to resolve the user, performs the necessary data work,
  and expose it through a new global wrapper (`function appSomething()`).
- **Front-end modules** – extend `V2_AppShell.html` with additional navigation
  buttons and call your new Apps Script functions via `google.script.run`.

## Migration notes

- All legacy Apps Script files are still available in the repository should you
  need to port individual workflows. They no longer participate in the runtime
  and can be deleted once the migration is complete.
- Password validation now expects the Base64-encoded SHA-256 digest of the raw
  password, aligning with the historical storage format in the `Users` sheet. If
  your dataset used a different hashing approach, update `hashPassword` inside
  `Code.js` accordingly.
- Sessions are stored exclusively in the `Sessions` sheet. The cache layer keeps
  active session lookups fast, but expired tokens are pruned when a user attempts
  to reuse them.

## Support

If you run into issues while extending the rebuilt runtime, review the inline
comments in `Code.js` – every helper has been documented with the intended
responsibilities to make future maintenance straightforward.
