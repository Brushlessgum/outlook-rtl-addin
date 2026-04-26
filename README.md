# RTL by Default — Outlook Add-in

A small Outlook event-based add-in that **automatically applies RTL (right-to-left) direction and right alignment to every new email** you compose. Replicates the "default text direction" option that Windows Outlook has, but Outlook for Mac is missing.

Works in: **New Outlook for Mac**, **Outlook on the Web**, **Outlook for Windows** (new and classic).

---

## What you'll do (high level)

1. **Host these files on the public internet over HTTPS** (free with GitHub Pages).
2. **Update the manifest** to point at your hosting URL.
3. **Sideload the manifest** into Outlook.

Total time: ~15 minutes.

---

## Step 1 — Host the files on GitHub Pages (free, HTTPS)

1. Go to <https://github.com> and sign in (or create a free account).
2. Click **+ → New repository**.
3. Name it `outlook-rtl-addin`. Keep it **Public**. Click **Create repository**.
4. On the new repo's page, click **uploading an existing file** (or drag-and-drop).
5. Drag the **entire contents of this folder** (manifest.xml, commands.html, commands.js, icons/) into the upload area, then **Commit changes**.
6. Go to repo **Settings → Pages** (left sidebar).
7. Under **Branch**, pick `main` and `/ (root)`. Click **Save**.
8. Wait ~1 minute. The page will show: **"Your site is live at https://YOUR-USERNAME.github.io/outlook-rtl-addin/"**. Copy that URL.

Test it works by visiting `https://YOUR-USERNAME.github.io/outlook-rtl-addin/commands.js` in your browser — you should see the JavaScript code.

---

## Step 2 — Update the manifest with your URL

Open `manifest.xml` in any text editor (TextEdit, VS Code, etc).

**Find and replace** every occurrence of:

```
YOUR-USERNAME.github.io
```

with your actual GitHub username, e.g.:

```
gilrosenkrantz.github.io
```

There are 11 occurrences. Save the file.

**Re-upload the updated `manifest.xml`** to your GitHub repo (overwrite the old one).

---

## Step 3 — Sideload into Outlook for Mac

### Option A — Through New Outlook UI (easiest)

1. Open **New Outlook for Mac**.
2. In the toolbar, click **⚙️ → View all Outlook settings** (or click the **Apps** / **Get Add-ins** icon — the icon looks like four squares).
3. Choose **My add-ins** → **Custom Addins** → **Add a custom add-in** → **Add from file…**
4. Select the `manifest.xml` file from this folder.
5. Click **Install**. Accept the warning about custom add-ins.

### Option B — Drop manifest in the Outlook folder (fallback)

If the UI option above is missing, you can install via filesystem:

1. In Finder, press **Cmd + Shift + G** and paste:
   ```
   ~/Library/Containers/com.microsoft.Outlook/Data/Documents/wef
   ```
   (If the `wef` folder doesn't exist, create it.)
2. Copy `manifest.xml` into that folder.
3. Restart Outlook.

---

## Step 4 — Test it

1. Click **New Mail**.
2. As soon as the compose window opens, your cursor should be on the right side of the body, and any signature should be right-aligned.
3. Start typing in Hebrew or Arabic — text flows RTL automatically.

If it didn't work, see **Troubleshooting** below.

---

## Customization

### Want to keep using LTR sometimes?
Change `commands.js` so it only runs when an `Option` key was held, or add a check for the recipient's domain. Or simpler: **uninstall** when you're going to write English-heavy emails for a while.

### Want only the body wrapped (not the signature)?
In `commands.js`, change the `rtlBody` line so it inserts an empty RTL `<div>` *before* the existing body instead of wrapping it:

```js
const rtlBody = '<div dir="rtl" style="direction:rtl;text-align:right;"><br></div>' + existingBody;
```

### Want a button you can click instead of automatic?
Replace the `<ExtensionPoint xsi:type="LaunchEvent">` block in the manifest with a `MessageComposeCommandSurface` extension point that adds a ribbon button. Microsoft has a working sample here: <https://learn.microsoft.com/office/dev/add-ins/outlook/add-in-commands-for-outlook>

---

## Troubleshooting

**"My add-ins doesn't show Custom Addins"**
You may be on the *legacy* Outlook for Mac. Toggle **New Outlook** on (top-right of the main window).

**"It installed but nothing happens on new emails"**
- Quit and reopen Outlook fully (Cmd + Q, then relaunch).
- Check that your hosting URL works in the browser.
- Open the developer console: **Help → Add-in Diagnostics → Show Logging Output** and look for errors mentioning `onNewMessageComposeHandler`.

**"It works in new mail but not on Reply"**
The `OnNewMessageCompose` event covers new mail, reply, reply-all, and forward in current Outlook builds. If reply isn't firing, your Outlook is likely on an older Mailbox requirement set — update Outlook to the latest version.

**"My signature is now in Hebrew alignment too"**
That's expected — the wrapper covers the whole body. Use the "Want only the body wrapped" tweak above.

**"I want to remove the add-in"**
Outlook → ⚙️ → Add-ins → My add-ins → click the `...` next to "RTL by Default" → Remove.

---

## Files in this folder

| File | Purpose |
|------|---------|
| `manifest.xml` | The add-in's identity card. Tells Outlook what URLs to load and what events to handle. |
| `commands.html` | Headless page that loads Office.js + commands.js. Never visually shown. |
| `commands.js` | The actual logic — wraps the body in an RTL div on every new compose. |
| `icons/icon-*.png` | Required icons in 5 sizes (16, 32, 64, 80, 128). |
| `README.md` | This file. |

---

## How it works (technical)

When you open a new compose window, Outlook fires the `OnNewMessageCompose` event. The manifest registers `onNewMessageComposeHandler` (in `commands.js`) as the handler for that event. The handler:

1. Reads the current HTML body (which may already include your signature).
2. Wraps it in `<div dir="rtl" style="direction:rtl;text-align:right;unicode-bidi:embed;">…</div>`.
3. Writes it back via `item.body.setAsync`.
4. Calls `event.completed()` so Outlook knows it's done.

The whole thing runs headlessly — no UI is shown, no user action needed.
