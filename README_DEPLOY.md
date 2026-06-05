# Deploying the Furniture Tracker — Supabase + GitHub + Vercel

This turns the artifact into a real web app with **designer logins**, a
**cloud database** (data syncs across devices), and a **secure AI proxy**
(your API key stays on the server).

You'll do it in this order:
**1) Local project → 2) Supabase → 3) Wire it in → 4) GitHub → 5) Vercel.**

> Good news: the app uses **no Tailwind** (all styling is built in), so
> there's nothing extra to configure for styles.

---

## 1. Create the local project (Vite + React)

You need Node 18+ installed (`node -v` to check).

```bash
npm create vite@latest furniture-tracker -- --template react
cd furniture-tracker
npm install
npm install @supabase/supabase-js lucide-react xlsx
```

Now copy these files into the project (paths shown relative to the project root):

| From this bundle            | Put it here                     |
|-----------------------------|---------------------------------|
| `furniture_tracker.jsx`     | `src/furniture_tracker.jsx`     |
| `src/main.jsx`              | `src/main.jsx`  (overwrite)     |
| `src/AuthGate.jsx`          | `src/AuthGate.jsx`              |
| `src/lib/supabaseClient.js` | `src/lib/supabaseClient.js`     |
| `src/lib/store.js`          | `src/lib/store.js`             |
| `api/anthropic.js`          | `api/anthropic.js`             |
| `.env.example`              | `.env.example`                 |

---

## 2. Set up Supabase

1. Go to **supabase.com** → create a free account → **New project**.
   Pick a name, a strong database password, and a region near Phoenix.
2. When it finishes provisioning, open **SQL Editor** → **New query**,
   paste the contents of **`schema.sql`**, and click **Run**.
   (This creates the `kv` table and locks every row to its owner.)
3. Go to **Project Settings → API** and copy two values:
   - **Project URL**  →  `VITE_SUPABASE_URL`
   - **anon public key**  →  `VITE_SUPABASE_ANON_KEY`
4. (Optional, recommended for first tests) **Authentication → Providers →
   Email**: turn **OFF "Confirm email"** so you can sign in immediately
   without clicking a confirmation link. Turn it back on later.

Create a local `.env` file (copy from `.env.example`) and fill in the two
Supabase values. You'll add `ANTHROPIC_API_KEY` in Vercel later (the AI
features won't run on `localhost` unless you also run `vercel dev`).

---

## 3. Wire the app to Supabase (two small edits to `furniture_tracker.jsx`)

**Edit A — use the cloud store instead of the sandbox storage.**
Find this block near the top of `furniture_tracker.jsx` (~line 392):

```js
const store = {
  async load(key) { ... window.storage ... },
  async save(key, val) { ... window.storage ... },
};
```

**Delete that whole `const store = { ... };` block** and instead add this
import at the very top of the file, with the other imports:

```js
import { store } from "./lib/store";
```

Everything else in the app already calls `store.load` / `store.save`, so
that's the only change needed for persistence.

**Edit B — point the AI calls at your proxy.**
There are **3** places that call the Anthropic API directly. Find each:

```js
const res = await fetch("https://api.anthropic.com/v1/messages", {
```

and change the URL to your serverless proxy:

```js
const res = await fetch("/api/anthropic", {
```

(Leave the `method`, headers, and `body` exactly as they are.)

That's it — two edits.

---

## 4. Run it locally

```bash
npm run dev
```

Open the local URL. You'll see the **sign-in screen** (AuthGate) → create
an account → you're in. The app will seed a sample "Mercer" project on
first load and save it to your Supabase account. Try editing something,
refresh, and confirm it persists. (The link auto-fill / AI features need
the proxy + key, which work once deployed — or run `npx vercel dev`
locally with the key set.)

---

## 5. Push to GitHub

```bash
git init
git add -A
git commit -m "Furniture tracker — initial"
```

Create an empty repo at **github.com** (no README), then:

```bash
git remote add origin https://github.com/<you>/furniture-tracker.git
git branch -M main
git push -u origin main
```

> Make sure `.env` is **not** committed — Vite's default `.gitignore`
> already ignores it. Never commit your keys.

---

## 6. Deploy on Vercel

1. Go to **vercel.com** → **Add New → Project** → **Import** your GitHub
   repo. Vercel auto-detects Vite (build: `npm run build`, output: `dist`).
2. Before deploying, expand **Environment Variables** and add all three:
   - `VITE_SUPABASE_URL`
   - `VITE_SUPABASE_ANON_KEY`
   - `ANTHROPIC_API_KEY`  ← secret; used only by `/api/anthropic`
3. Click **Deploy**. You'll get a live `https://...vercel.app` URL.
4. Back in **Supabase → Authentication → URL Configuration**, add your
   Vercel URL to **Site URL** / **Redirect URLs** so auth works in prod.

Every `git push` to `main` now auto-redeploys.

---

## What works after this

- A real **designer login** (email + password), data in **Postgres**,
  syncing across her laptop, desktop, and phone.
- The **AI link auto-fill, Gmail scan, and tracking** calls run through
  your server with the key hidden.
- Everything else in the app (schedule, finance, presentations, Excel
  export, etc.) works as-is.

## What's still a prototype (Phase 2)

- **External client logins.** Right now the client portal is reachable
  only inside the designer's own session (the "Open client view" preview).
  Letting a client open their own link securely needs a small
  `/api/portal` serverless function that uses the Supabase **service role**
  key to serve just the client-visible slice (selections, approvals,
  comments, shared files) validated by the access code — without exposing
  the database. The schema notes at the bottom of `schema.sql` are the
  starting point. Tell me when you want this and I'll build those endpoints
  and wire the client side to them.
- **File uploads** (receipts, mood images, presentation photos) are stored
  inline as data in the database. That's fine for now; for lots of large
  images, moving them to **Supabase Storage** buckets is the clean upgrade.
- **Scheduled price/stock re-checks** ("re-check weekly") need a Vercel
  Cron job hitting an endpoint — also a quick Phase 2 add.

## Rough costs

- **Supabase** free tier: plenty for one designer + several projects.
- **Vercel** Hobby tier: free for this.
- **Anthropic API**: pay-per-use; only the auto-fill/scan/tracking calls
  cost anything, typically fractions of a cent each.
