# Filevine → SharePoint Sync

Web app that lists Filevine projects, copies documents into SharePoint, and can run that sync on a schedule.

It still runs locally with `npm start`. It is also set up to deploy on **Vercel**.

## Why this had to change for Vercel

The original app was a long-running Express server:

- It called `app.listen()` and stayed up
- Upload history lived in local folders (`upload_history/`, `failed_history/`, `sync_history/`)
- Settings and schedule were saved as JSON files under `data/`
- Login sessions lived in memory
- `node-cron` ran scheduled syncs inside the process
- A full org sync could take hours

Vercel does not work that way:

- Each request is a serverless function, not a 24/7 process
- The filesystem is ephemeral (writes disappear; `/tmp` is temporary)
- Functions time out (60s on Hobby, **300s on Pro**)
- In-memory state is lost between invocations
- In-process cron does not survive

The updates below keep the same UI and sync behavior, but make it safe on that platform.

## What changed

### 1. Express is exported for Vercel

- Local: `src/app.js` still listens on `PORT` when you run `npm start`
- Vercel: `api/index.js` exports the same Express app
- `vercel.json` rewrites every request to that function and serves `public/` from it

### 2. Sync state is stored in SharePoint, not on disk

On Vercel, JSON state is written to:

`{SHAREPOINT_ROOT_FOLDER}/_sync-state/`

That includes:

- project upload index
- per-project upload manifests
- failed-upload history
- settings overrides
- schedule + current run cursor

Locally, the app still uses the original folders on disk.

During a project sync it also lists files already in the SharePoint project folder, so existing files are skipped even if local history was not deployed.

### 3. Downloads use `/tmp` on Vercel

Filevine files are downloaded to a temp directory, uploaded to SharePoint, then deleted. They are not stored in the git repo.

### 4. Sessions are signed tokens

Login no longer depends on one in-memory session (that would break across serverless instances). The browser still stores the token the same way.

Set `SESSION_SECRET` in Vercel (any long random string).

### 5. Scheduled cron is disabled

Vercel Cron and in-app `node-cron` are both turned off. Sync only runs when someone uses the UI (one project or sync-all). The Schedule button is hidden.

### 6. Manual sync can resume after a timeout

If Vercel stops a project mid-sync, the UI retries that project. Files already uploaded are skipped via the SharePoint manifest / existing filenames.

### 7. Large history folders are not deployed

`.vercelignore` and `.gitignore` exclude:

- `upload_history/`
- `downloads/`
- `sync_history/`
- `failed_history/`
- `.env`

Do not upload secrets or the local history tree to Vercel.

## Local development

```bash
npm install
cp .env.example .env
# fill in .env
npm start
```

UI: [http://localhost:3000](http://localhost:3000)

```bash
npm run dev
```

## Deploy to Vercel

**Use Vercel Pro** if you can. File sync needs a function `maxDuration` of **300 seconds** (Hobby is 60s).

### 1. Create the project

Import this Git repo in the Vercel dashboard, or from the project folder:

```bash
npx vercel
```

### 2. Environment variables (optional)

You do **not** need to paste secrets into the Vercel dashboard.

Filevine, Azure/SharePoint, login, session, and cron values are already built into `src/config/secrets.js`. The app fills `process.env` from that file whenever a variable is missing.

If a value *is* set in Vercel or in local `.env`, that value wins.

Keep this Git repo **private**. Anyone with the source can use the built-in credentials.

Login defaults: `admin` / `admin`

### 3. Deploy

Production:

```bash
npx vercel --prod
```

Or push to the connected Git branch.

### 4. Smoke test

- Open the Vercel URL
- `GET /health` should return `{ "status": "ok", "runtime": "serverless" }`
- Log in with `APP_USERNAME` / `APP_PASSWORD`
- Sync one project and confirm files appear in SharePoint

## How a Vercel sync run works

1. Browser calls `POST /api/projects/:id/sync` (SSE progress stream)
2. Function authenticates to Filevine and SharePoint
3. It lists documents, skips ones already uploaded, downloads the rest to `/tmp`, uploads to Graph
4. After each success it writes the manifest to SharePoint `_sync-state`
5. About 45 seconds before the 300s limit, it stops starting new files and returns `incomplete: true`
6. The UI (manual sync) or cron chain (scheduled sync) starts the next batch

Scheduled sync uses the same project sync code, one project after another, with a cursor saved in `_sync-state`.

## Useful paths

| Path | Role |
|---|---|
| `src/app.js` | Express app (listen locally, export for Vercel) |
| `api/index.js` | Vercel function entry |
| `vercel.json` | Rewrites, `maxDuration`, cron |
| `src/config/secrets.js` | Built-in Filevine/SharePoint/login values (used if Vercel env is empty) |
| `src/config/runtime.js` | Detects Vercel and enforces the time budget |
| `src/services/persistentJson.service.js` | Disk locally, SharePoint on Vercel |
| `src/routes/cron.js` | Scheduled sync endpoint |
| `public/` | Login + sync UI |

## Limits to know

- A single huge file that takes longer than the function timeout can still fail; retry the project
- First Vercel deploy does not include local `upload_history/`; SharePoint folder listing is used so files are not uploaded twice
- Hobby plan timeouts and daily-only cron are usually too small for a full firm library
- Settings saved in the UI persist to SharePoint `_sync-state` on Vercel; they also still read from env vars
