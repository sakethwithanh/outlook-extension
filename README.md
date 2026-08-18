# Export Thread — Outlook add-in

Adds an **Export Thread** button to the Outlook ribbon. Select any message,
click it, and every message in that conversation is written to one formatted
Markdown file.

Works in Classic Outlook (Windows), New Outlook, Outlook on the web, and Mac.
One codebase, no VBA, no per-user install.

## Output

```markdown
# Mail Thread: Q3 Vendor Pricing

Exported: 2026-08-18 14:22
Messages: 6

---

## [1] Q3 Vendor Pricing

**From:** A. Kumar <a.kumar@example.com>
**To:** M. Rao <m.rao@example.com>
**Date:** 2026-08-11 09:14
**Attachments:** pricing-v3.xlsx

Sharing the revised sheet ahead of Thursday.

---
```

Each message shows only its own new text. Quoted reply history is removed by
Microsoft Graph's `uniqueBody`, not by pattern matching, so nothing is
duplicated and nothing is wrongly truncated.

## How it works

| File | Role |
|---|---|
| `manifest.template.xml` | Ribbon button, task pane location, permissions |
| `src/config.js` | Your Azure client ID and options |
| `src/auth.js` | Token via Nested App Authentication (MSAL) |
| `src/graph.js` | Lists the conversation, fetches bodies and attachment names |
| `src/format.js` | Builds the Markdown |
| `src/taskpane.js` | Task pane UI and orchestration |

The add-in reads only `conversationId` from the open item. The messages
themselves come from Microsoft Graph.

> **Note on EWS.** An earlier approach used `makeEwsRequestAsync`, which avoids
> Azure app registration. Microsoft begins phased EWS disablement on
> 1 October 2026 and retires it fully on 1 April 2027, so this uses Graph.

## Setup

### 1. Azure app registration

In the [Azure portal](https://portal.azure.com) → **App registrations** → **New registration**:

- Name: `Outlook Export Thread`
- Supported account types: *Accounts in any organizational directory and personal Microsoft accounts* (for testing), or *single tenant* (for the org rollout)
- Platform: **Single-page application**
- Redirect URIs — add all of these:
  - `https://<you>.github.io/<repo>/src/taskpane.html`
  - `brk-multihub://<you>.github.io` *(required for Nested App Authentication)*

Then **API permissions** → Microsoft Graph → Delegated → **Mail.Read**.

`Mail.Read` is user-consentable, so no admin consent is needed unless your
tenant has disabled user consent.

Copy the **Application (client) ID** into `src/config.js`.

### 2. Publish to GitHub Pages

```bash
git init
git add .
git commit -m "Add Outlook thread export add-in"
git branch -M main
git remote add origin https://github.com/<you>/<repo>.git
git push -u origin main
```

Then in the repo: **Settings → Pages → Source: `main` / root**.

### 3. Stamp the URL into the manifest

```powershell
.\scripts\set-base-url.ps1 -BaseUrl "https://<you>.github.io/<repo>"
```

### 4. Install

**Test it on yourself** (no admin needed):

Outlook → **Get Add-ins** → **My add-ins** → **Custom Add-ins** →
**Add from file** → pick `manifest.xml`.

**Roll out to everyone** (needs a Microsoft 365 admin):

[Microsoft 365 admin center](https://admin.microsoft.com) → **Settings** →
**Integrated apps** → **Upload custom apps** → upload `manifest.xml` →
assign to **Everyone**. The button appears for all users within about 24 hours.

## Mailbox requirements

| Mailbox | Supported |
|---|---|
| Microsoft 365 / Exchange Online | Yes |
| outlook.com / hotmail.com / live.com | Yes |
| Gmail or other IMAP/POP in Outlook | **No** — add-ins do not load for these accounts |

A Microsoft account that merely *signs in* with a Gmail address is not enough.
The account needs an actual Outlook mailbox.

## Limits

- Only messages still in your mailbox appear. Deleted or permanently-archived
  replies cannot be recovered.
- Message bodies export as plain text. Inline images and rich formatting drop.
- On Classic Outlook the task pane runs in WebView2, which can block file
  downloads. The pane detects this and offers **Copy to clipboard** instead.
- One Graph call per message, capped at 4 concurrent. A 50-message thread takes
  a few seconds.

## Changing the output

Everything about the file lives in `src/format.js` — `buildMarkdown()` for the
layout, `buildFileName()` for the name. Set `NEWEST_FIRST` in `src/config.js`
to flip message order.

## Why MSAL is vendored

`vendor/msal-browser.min.js` is a copy of `@azure/msal-browser` (v5.18.0),
taken from the npm package's `lib/` folder.

Microsoft stopped publishing MSAL.js to a CDN at v3. The only supported way to
use it without a bundler is to serve the lib file yourself, so it is committed
here rather than fetched at runtime. It is a UMD build and exposes `window.msal`.

To update it:

```bash
npm pack @azure/msal-browser
tar -xzf azure-msal-browser-*.tgz
cp package/lib/msal-browser.min.js vendor/
```
