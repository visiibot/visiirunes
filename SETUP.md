# The Runic System — multi-device setup

This gives every student's device (and yours) a shared, synced rune database, backed by a Google Sheet you own.

## New: per-adventurer PINs

Every adventurer name now has its own 4-digit PIN, separate from the teacher's passphrase and separate from the 6-digit one-time milestone codes:

- **The first time** anyone types a given name and tries to roll/transform/craft, the app asks them to **set** a 4-digit PIN for that name. Whatever they enter becomes that adventurer's permanent PIN.
- **After that**, the same name will ask for that PIN every time (on a device that hasn't already unlocked it this session) before letting anyone roll, transform, or craft under that name.
- **Viewing** the My Runes / Codex tabs doesn't require a PIN — only the three actions that change someone's runes do.
- Wrong PIN attempts are tracked server-side; **5 wrong attempts locks that adventurer's account for 60 seconds**, even across devices, to slow down guessing.
- If a student forgets their PIN, you (the teacher) can reset it from the **Teacher > Students** table — a "Reset PIN" button clears it so they're prompted to set a new one next time.
- The 🔒/🔓 icon next to the Adventurer field shows whether the currently-typed name is unlocked *on this device* this session — click it to manually lock/unlock.

If you already deployed the app before this feature existed, see "Upgrading" below — it's non-destructive to existing rune data.

## 1. Create the backend (~5 minutes, one time)

1. Go to **sheets.google.com** and create a new blank spreadsheet. Name it something like `Runic System Data`.
2. In the menu bar: **Extensions > Apps Script**. A code editor opens in a new tab.
3. Delete the placeholder `function myFunction() {}` code, and paste in the entire contents of **`Code.gs`**.
4. Save (Ctrl/Cmd+S). Name the project `Runic System API` if asked.
5. At the top of the editor, in the function dropdown, choose **`setup`**, then click **Run**.
   - The first time, Google will ask you to authorize the script. Click through **Review permissions > (your account) > Advanced > Go to Runic System API (unsafe) > Allow**. This warning is normal for scripts you write yourself — it's only unverified because it isn't published on the Marketplace.
   - Check the Sheet — you should now see three tabs: `Students`, `Pins`, `Config`.
6. Click **Deploy > New deployment**.
   - Click the gear icon next to "Select type" and choose **Web app**.
   - Execute as: **Me**
   - Who has access: **Anyone**
   - Click **Deploy**, authorize again if asked.
7. Copy the **Web app URL** — it looks like `https://script.google.com/macros/s/AKfycb.../exec`. Keep it somewhere safe; you'll paste it into the app next.

**If you ever edit `Code.gs` later:** go to **Deploy > Manage deployments > (pencil icon) > Version: New version > Deploy**. Editing the code alone does *not* update the live URL's behavior until you do this.

## Upgrading an existing deployment

If you already set up the backend before per-adventurer PINs existed:

1. Open your Apps Script project (from the Sheet: **Extensions > Apps Script**).
2. Replace all the code with the new `Code.gs`.
3. Run **`setup`** again from the function dropdown. This only *adds* the new PIN columns to your existing `Students` sheet — it does not touch anyone's rune counts, crafted items, or milestone code history.
4. **Deploy > Manage deployments > pencil (edit) icon > Version: New version > Deploy.** This step is required — just saving the code does not update the live URL.
5. The Web App URL stays exactly the same, so you don't need to update it in the frontend. Just replace `runic-system-cloud.html` on illswap.net with the new version.
6. Every existing student will simply be asked to set a PIN the next time they act — their rune inventory is untouched.

## 2. Connect the app

1. Open `runic-system-cloud.html` in Chrome (double-click it, or host it at illswap.net and visit the page).
2. A "Backend connection" box appears automatically the first time. Paste the Web App URL from step 1.7, click **Save & Connect**.
3. The little dot under the header should turn green: "Connected to shared class database."
4. Do this once per device (teacher's computer, each student's Chromebook, etc.) — it's just remembering the URL, not creating separate data.

## 3. Set your teacher passphrase

- Go to the **Teacher** tab, type any passphrase (3+ characters), click **Unlock**. Whatever you type the *first time, on any device,* becomes the shared class passphrase from then on — because it's stored on the server (the Sheet), not the browser.
- Use that same passphrase from any device to unlock Teacher tools there too.

## 4. Hosting on illswap.net

`runic-system-cloud.html` is a single static file — no server needed on your end. Upload it to illswap.net exactly as you would any HTML page (e.g. `illswap.net/runes/index.html`). It will call out to your Apps Script URL from any visitor's browser, so it works the same whether opened locally or from your domain.

## Notes & limits

- **Concurrency:** the backend locks briefly during each write, so two students redeeming PINs at the same moment won't corrupt data — one just waits a beat.
- **Google Sheets isn't built for heavy traffic.** Fine for a class of dozens rolling occasionally; if you ever scale to hundreds of simultaneous users, you'd want a proper database instead.
- **Security is classroom-grade, not bank-grade.** The teacher passphrase and student names are sent in plain requests over HTTPS — fine for a fun in-class tool, not for anything sensitive.
- **Do not use the Claude "Publish" artifact link for this version** — Claude's hosted preview sandbox blocks outbound calls to script.google.com, so the cloud version won't sync there. Always open it as a real file (local double-click or hosted on illswap.net).
- The old fully-offline `runic-system.html` (browser-only, no backend) is still available if you ever want a no-setup fallback for a single kiosk computer.
