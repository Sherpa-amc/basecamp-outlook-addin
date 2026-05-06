# Base Camp — Outlook Add-in

This Office Add-in replaces the legacy VBA macro for saving email PDF
attachments to fund folders. It works on:

- Outlook for Microsoft 365 (Windows desktop) — primary target
- Outlook on the Web
- Outlook for Mac

It does **not** require any VBA — there's no AV friction.

---

## How it works

```
[Outlook] User opens email with PDF → clicks "Save to Fund" ribbon button
   ↓
[Add-in task pane] Reads attachments via Office.js
   ↓ HTTPS
[Local Flask server] localhost:8765 — runs on each user's machine
   ↓
[O:\] Saves PDF to the matching fund folder
```

The local Flask server (started via `Open Fund Manager.bat`) does the file
writes. The add-in is just a thin UI client that talks to it over HTTPS.

---

## Files in this folder

| File | Purpose |
|------|---------|
| `manifest.xml` | Office Add-in manifest. Edit `HOST_URL` before deploying. |
| `taskpane.html` | The UI shown in the Outlook side pane. |
| `taskpane.js` | Logic: match attachment, prompt user, save via Flask. |
| `taskpane.css` | Styling. |
| `commands.html` + `commands.js` | Required helper page for ribbon commands. |
| `assets/icon-*.png` | Ribbon icons (16, 32, 64, 80, 128 px). |

---

## Deployment (admin one-time)

### Step 1 — Host the static files

The manifest URL and all referenced files (HTML/JS/CSS/icons) must be served
over **HTTPS**. Pick one of these hosting options:

#### Option A: GitHub Pages (recommended)

1. Create a new private repository, e.g. `basecamp-outlook-addin`.
2. Push the entire `outlook_addin/` folder to the repo root.
3. In GitHub: Settings → Pages → Source = `main` branch, root.
4. After ~1 minute you'll get an HTTPS URL like
   `https://your-org.github.io/basecamp-outlook-addin/`.

#### Option B: Cloudflare Pages / Azure Static Web Apps

Both have free tiers. Connect to the same repo. Pick whichever your team
already uses.

#### Option C: Internal IIS (if firewall-restricted)

Drop the folder onto an internal HTTPS web server and note the URL.

### Step 2 — Edit the manifest

Open `manifest.xml` and replace **every occurrence** of `HOST_URL` with your
hosting URL (no trailing slash). For example:

```
HOST_URL/taskpane.html  →  https://your-org.github.io/basecamp-outlook-addin/taskpane.html
```

There are ~7 places to update. Commit + push the edited manifest.

### Step 3 — Deploy via Microsoft 365 Admin Center

(Requires Global Admin or Exchange Admin rights.)

1. Sign in to <https://admin.microsoft.com>.
2. Settings → Integrated apps → Upload custom apps.
3. Choose **Office Add-in** type.
4. Provide the URL of your hosted `manifest.xml` (e.g.
   `https://your-org.github.io/basecamp-outlook-addin/manifest.xml`).
5. Assign to a security group (e.g. "Liquid Strategy") or the entire tenant.
6. Submit. Within 12-24h Outlook auto-installs the add-in for everyone in scope.

If you don't have admin rights, use the per-user sideload steps below
instead — each colleague does it once, takes ~30 seconds.

---

## Installation (per user, fallback)

If centralized deployment isn't available:

### Outlook Desktop (Windows / Mac)

1. Open Outlook.
2. On the Home ribbon, click **Get Add-ins**.
3. In the dialog, click **My add-ins** in the left menu.
4. Scroll down to "Custom Addins" → **Add a custom add-in** → **Add from URL...**
5. Paste the manifest URL (e.g.
   `https://your-org.github.io/basecamp-outlook-addin/manifest.xml`) and click
   **OK**.
6. Confirm installation.
7. Restart Outlook.

### Outlook on the Web

1. Settings (gear icon) → View all Outlook settings → **General → Manage add-ins**.
2. Click **My add-ins** → **Add a custom add-in** → **Add from URL...**
3. Same URL as above.

---

## First-time use

1. **Start the local server**: double-click `Open Fund Manager.bat` (this is
   the same .bat you already use). It now launches Flask on
   `https://localhost:8765`. On first launch it auto-generates a self-signed
   certificate (one-time, takes 2 seconds).
2. **Open an email with a PDF attachment** in Outlook.
3. **Click "Save to Fund"** on the ribbon (you'll see a "BC" icon).
4. The task pane opens on the right. The header shows a green dot if the local
   server is running.
5. Click **Save attachments**.
6. For each PDF: pick a doc type → file is saved to the right folder. If the
   sender/subject/filename doesn't match a known fund, you'll be prompted to
   create one.
7. For "Letter" docs: you'll be asked whether to copy the file to the weekly
   reading pack folder (Manager / Research).

---

## Troubleshooting

### "Server not running" red dot in the task pane
- Start the server: double-click `Open Fund Manager.bat`.
- If it's already running, check that the URL is `https://localhost:8765` (not
  `http://`). The legacy HTTP server is incompatible with Office add-ins.

### First-time HTTPS warning in browser
- Browsers may warn about the self-signed cert. The add-in itself works fine
  inside Outlook (Edge WebView2 trusts `localhost`). If you want to silence the
  warning when opening the web UI in your browser, install the cert from
  `server/certs/localhost.pem` into the Trusted Root store:
  - Right-click `localhost.pem` → Install Certificate → Local Machine →
    Place in "Trusted Root Certification Authorities".

### "Save to Fund" button doesn't appear in Outlook
- Wait 24h after admin deploy. Or restart Outlook.
- Check Outlook → Get Add-ins → My add-ins to confirm it's installed.

### Attachment fails to save with "Could not read attachment"
- This can happen if the email is in a local PST that's not synced to Exchange.
  Workaround: forward the email to yourself, then save from the new copy.

### Multi-select (selecting multiple emails) doesn't work
- v1 supports one email at a time. Multi-select is on the roadmap. For now,
  open each email and click the button.

---

## Updating the add-in

Push changes to the hosted repo. Users see the new version on the next
Outlook session (Office caches but invalidates per its rules; bump
`<Version>` in `manifest.xml` to force a refresh).

---

## Decommissioning the VBA macro

After 2 weeks of parallel use:

1. Confirm no regressions on real workflows.
2. Update `Base Camp - Setup Guide.docx` to reference the add-in install
   process instead of the macro.
3. Each user can remove the FactsheetSaver module from Outlook VBA (Alt+F11
   → right-click module → Remove).
4. Keep `migration/write_macro.py` in the repo as historical reference.
