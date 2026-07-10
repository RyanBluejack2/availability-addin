# Calendar Availability — Outlook Web Add-in

A sideloadable Outlook add-in that generates a bulleted list of your open
calendar time. Works in **classic Outlook, New Outlook, and Outlook on the web**.
No Visual Studio required.

Output format:
```
• Monday 7/13: 10:00am-1:00pm CT, 2:30pm-4:00pm CT
• Tuesday 7/14: 10:00am-2:00pm CT
...
```

---

## How it works (the moving parts)

Unlike a desktop (VSTO) add-in, a web add-in runs in a sandbox and can't read
your mailbox directly. It reads your calendar through the **Microsoft Graph API**,
authenticated with **MSAL** (Microsoft's browser sign-in library). That's why
setup has an Azure step — you're registering this add-in as an app allowed to
read *your own* calendar.

Three things have to be true for it to run:
1. The files (`index.html`, `src/`, `assets/`) are **hosted at an https URL**.
2. An **Azure app registration** exists, with your hosting URL as a redirect URI
   and `Calendars.Read` permission.
3. The **manifest** is **sideloaded** into Outlook.

You have Azure admin, so all three are self-serve. Budget ~30 minutes the first time.

---

## STEP 1 — Host the files (https is required)

Office add-ins must load over https. Pick whichever is easiest for you:

### Option A — GitHub Pages (free, recommended)
1. Create a new GitHub repo (e.g. `availability-addin`).
2. Upload the contents of this folder (`index.html`, `manifest.xml`, `src/`, `assets/`).
3. Repo **Settings → Pages** → Source: `main` branch, root → **Save**.
4. After a minute you'll get a URL like:
   `https://YOURNAME.github.io/availability-addin/`
   That is **YOUR_HOST**. Note it down.

### Option B — Azure Static Web Apps / any static host
Any https static host works (Netlify, Cloudflare Pages, an internal server).
Whatever the public https base URL is, that's **YOUR_HOST**.

### Option C — local testing only
`npx http-server` gives you http://localhost — but Outlook requires https, so
for anything beyond a quick look, use A or B.

---

## STEP 2 — Register the app in Azure (Entra ID)

1. Go to **https://entra.microsoft.com** → **Applications → App registrations → New registration**.
2. Name: `Calendar Availability Add-in`.
3. Supported account types: **Accounts in this organizational directory only** (Bluejack single tenant).
4. **Redirect URI**: platform dropdown = **Single-page application (SPA)**, value =
   `https://YOUR_HOST/index.html` (exact, including `index.html`).
5. Click **Register**.
6. On the Overview page, copy the **Application (client) ID** and the **Directory (tenant) ID**.

### Add the calendar permission
7. Left menu → **API permissions → Add a permission → Microsoft Graph → Delegated permissions**.
8. Search `Calendars.Read`, check it, **Add permissions**.
9. Click **Grant admin consent for Bluejack** → **Yes**. (You can do this because you're admin.)
   The status should flip to a green "Granted".

### Confirm SPA redirect (important)
10. Left menu → **Authentication**. Confirm your redirect URI is listed under
    **Single-page application** (NOT "Web"). If it's under Web, remove it and re-add
    under SPA — MSAL browser auth requires SPA type.

---

## STEP 3 — Fill in config and the manifest

### 3a. Edit `src/config.js`
Replace the three placeholders:
```js
CLIENT_ID:   "the Application (client) ID from step 2.6",
TENANT_ID:   "the Directory (tenant) ID from step 2.6",
REDIRECT_URI:"https://YOUR_HOST/index.html"
```

### 3b. Edit `manifest.xml`
Find-and-replace every `https://REPLACE_WITH_YOUR_HOST` with your actual
`https://YOUR_HOST` (no trailing slash). There are several — replace them all
(icons, SourceLocation, Urls).

Also replace the placeholder **Id** GUID near the top with a fresh one. Any GUID
generator works, or in PowerShell: `[guid]::NewGuid()`.

### 3c. Re-upload
Push the edited files back to your host (re-upload to GitHub Pages, etc.).

---

## STEP 4 — Sideload the manifest into Outlook

### New Outlook / Outlook on the web
1. Open Outlook on the web (outlook.office.com).
2. **Settings (gear) → General → Manage add-ins** (or **Get add-ins**).
3. Choose **My add-ins → Add a custom add-in → Add from file**.
4. Select your edited `manifest.xml`. Confirm.

### Classic Outlook (desktop)
Same "Get add-ins → My add-ins → Add a custom add-in → Add from file" flow —
the dialog is reachable from the Home ribbon's **Get Add-ins** button.

> Org note: since you're admin, you could instead deploy this centrally via the
> **Microsoft 365 admin center → Settings → Integrated apps → Upload custom apps**,
> which pushes it to users without each person sideloading. For just you, the
> per-user sideload above is fastest.

---

## STEP 5 — Use it

1. Open any email or appointment in Outlook.
2. On the ribbon (or the `...` / "Apps" menu in New Outlook) you'll see
   **Availability**. Click it — the task pane opens on the right.
3. First time: click **Sign in**, complete the Microsoft prompt, approve access.
4. Set Start Date, Stop Date, Start Time, End Time, Time Zone → **Generate**.
5. **Copy as bullets** puts a real bulleted list on the clipboard for pasting
   into an email body.

---

## Troubleshooting

- **Button doesn't appear**: Give it a few minutes after sideloading, then fully
  restart Outlook. Confirm the manifest uploaded without validation errors.
- **Sign-in popup blocked / loops**: Usually the redirect URI mismatch. It must
  match `src/config.js` REDIRECT_URI *and* the Azure SPA redirect *exactly*,
  including `index.html`.
- **"AADSTS65001" or consent error**: admin consent (step 2.9) wasn't granted, or
  was granted for a different permission. Re-check API permissions shows
  `Calendars.Read` = Granted.
- **Times look off by an hour**: the zone list uses IANA names and Graph converts
  via the `Prefer` header; make sure you picked the right zone in the dropdown.
- **Blank calendar / no events**: confirm you're signed into the same account
  whose calendar you want, and the date range actually contains meetings.

---

## Files

- `manifest.xml` — the add-in definition Outlook loads
- `index.html` — the task pane UI
- `src/config.js` — **your** Azure IDs (edit this)
- `src/auth.js` — MSAL sign-in
- `src/graph.js` — Microsoft Graph calendar fetch
- `src/engine.js` — free-gap computation + formatting
- `src/app.js` — UI wiring
- `src/styles.css` — styling
- `assets/` — icons (placeholder blue calendar; replace anytime)
