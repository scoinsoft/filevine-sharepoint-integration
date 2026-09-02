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

### 3. Large files stream Filevine → SharePoint

Vercel `/tmp` cannot hold a multi-GB legal video. The sync no longer downloads the whole file to disk.

Instead it:

1. Gets a Filevine presigned URL
2. Pulls a small HTTP Range chunk (10 MiB on Vercel)
3. PUTs that chunk into a Microsoft Graph upload session
4. Repeats until SharePoint has the file

If the source does not support Range requests, it streams the body with backpressure so only one chunk is buffered at a time.

Locally, leftover disk-download helpers still exist, but the sync path does not write the full file under `downloads/` or `/tmp`.

### 4. Sessions are signed tokens

Login no longer depends on one in-memory session (that would break across serverless instances). The browser still stores the token the same way.

Set `SESSION_SECRET` in Vercel (any long random string).

### 5. Scheduled sync at 2:00 AM Mountain Time

Vercel Hobby allows **one cron per day**, and cron times are UTC. This app uses:

`0 8 * * *` → **08:00 UTC** = **2:00 AM Mountain Daylight Time** (most of the year). In winter (MST) that same tick is 1:00 AM.

That one call starts the nightly run. It then **chains to itself** until every project is done — that is not extra cron jobs.

Local `npm start` uses `node-cron` at 2:00 AM in `America/Denver` (true 2am MT including DST).

The Schedule button can turn this off.

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
3. It lists documents, skips ones already uploaded, then **streams** the rest Filevine → Graph (chunked upload session; no full file on `/tmp`)
4. After each success it writes the manifest to SharePoint `_sync-state`
5. About 20 seconds before the 300s limit, it **pauses** the current large file, saves the Graph upload session, and returns `incomplete: true`
6. The UI retries that project. The next run **resumes** the same SharePoint upload session from the last committed byte instead of starting over.

A 6 GB video may take several 5-minute function runs. Large files (32 MB+) upload one at a time so they do not compete for the time budget.

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

- Streaming removes the `/tmp` size limit. A 6 GB file still needs several Vercel function runs; the Graph upload session is saved and resumed
- First Vercel deploy does not include local `upload_history/`; SharePoint folder listing is used so files are not uploaded twice
- Hobby plan timeouts are usually too small for a full firm library
- Settings saved in the UI persist to SharePoint `_sync-state` on Vercel; they also still read from env vars
