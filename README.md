# IPQC Tracker — No React, No Database

Plain HTML/CSS/JS frontend + a small Node/Express backend that reads and
writes a real `.xlsx` file instead of a database. Historical data (imported
once from your existing tracker) and new submissions from the web form live
in the same Excel file, so Power BI always sees both when it refreshes.

```
ipqc-app/
  frontend/          static site — deploy as a Render "Static Site"
    index.html
    style.css
    config.js         <- set your backend URL here
    app.js
  backend/            deploy as a Render "Web Service"
    server.js
    import-existing-data.js
    package.json
    data/
      ipqc-tracker.xlsx   <- created automatically on first run
```

## Why this shape

- **No Node needed on your work laptop.** You only edit plain text files
  (HTML/CSS/JS). Node is only required on Render's build servers, which you
  never touch directly.
- **No database.** The Excel file at `backend/data/ipqc-tracker.xlsx` *is*
  the data store. The frontend never touches it directly — it always goes
  through the backend's API, which keeps reads/writes safe and consistent.
- **Power BI compatibility.** Power BI already knows how to refresh from an
  Excel file over the web. Point it at `GET /api/download` on your deployed
  backend and turn on scheduled refresh — no extra connector needed.

## 1. Import your real historical data

Before deploying, bring in your actual IPQC tracker so old records show up
in the app (and Power BI) from day one:

```bash
cd backend
npm install
node import-existing-data.js "/path/to/Your Existing IPQC Tracker.xlsx"
```

This reads the first worksheet of your file, matches columns by header text
(e.g. "Audit Date", "IPQC Auditor", "PIC Finding", "ICAR No.", etc. — see
`HEADER_ALIASES` at the top of `import-existing-data.js` if your headers are
worded differently), and writes everything into `backend/data/ipqc-tracker.xlsx`
in the format the app expects. Re-run it any time to start over from your
source file.

If your headers don't match any alias, the script tells you which columns
it recognized so you can add more aliases before re-running.

## 2. Run it locally to check everything works

```bash
cd backend
npm install
npm start          # starts the API on http://localhost:4000
```

In another terminal/tab, just open `frontend/index.html` directly in a
browser (double-click it, or use a simple static server). `config.js`
already points at `http://localhost:4000` for local testing.

## 3. Deploy the backend to Render

1. Push this whole `ipqc-app` folder to a GitHub repo.
2. On Render: **New → Web Service**, connect the repo, set **Root Directory**
   to `backend`.
3. Build command: `npm install`. Start command: `npm start`.
4. **Add a persistent disk** (Render dashboard → your service → Disks) and
   mount it at, say, `/var/data`. Without this, the Excel file gets wiped
   every time Render restarts or redeploys your service.
5. Add an environment variable `DATA_DIR` = `/var/data` so the server writes
   the Excel file (and uploaded images) to the persistent disk instead of
   the app folder.
6. Deploy. Note the resulting URL, e.g. `https://ipqc-backend.onrender.com`.

## 4. Deploy the frontend to Render

1. Edit `frontend/config.js` and set `API_BASE` to your backend's URL from
   step 3.
2. On Render: **New → Static Site**, same repo, **Root Directory** set to
   `frontend`, no build command needed (publish directory: `frontend`
   itself, or `.` if Root Directory is already `frontend`).
3. Deploy. This is the link you share with IPQC auditors.

## 5. Connect Power BI

In Power BI Desktop: **Get Data → Web**, paste:

```
https://ipqc-backend.onrender.com/api/download
```

Load it, build your visuals, then set up a **scheduled refresh** (Power BI
Service, if you publish there) or just hit Refresh before each review. Every
refresh pulls the current state of the Excel file — your imported history
plus every audit submitted through the web app since.

## Notes & limitations

- **Concurrent edits:** the backend reads and rewrites the whole Excel file
  per request. Fine for a small team submitting audits throughout a shift;
  not designed for many people saving at the exact same second.
- **Images:** uploaded photos are stored on the same persistent disk under
  `backend/data/uploads/` and served as a public URL saved into the
  "Picture URL" column — Power BI can't render the image itself but you'll
  have the link.
- **Free tier sleep:** Render's free web services spin down after
  inactivity and take a few seconds to wake on the next request. Fine for
  an internship pilot; consider a paid instance if this goes to full
  production.
- **Migrating to OneDrive/SharePoint later:** if you do get Microsoft Graph
  API access down the line, only `backend/server.js`'s Excel read/write
  functions need to change — the frontend and Power BI connection stay the
  same.
