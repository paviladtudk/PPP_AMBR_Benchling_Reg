# AMBR Plate Mapping (web version)

A browser-based version of the AMBR plate-mapping tool: upload your AMBR
timepoints CSV and Benchling export, answer a few questions on-screen, and
download the generated `plate_mapping_output.xlsx` (and, optionally, an
HPLC/SOA repack file) &mdash; no Python, Anaconda, or Spyder install required
for anyone using it. This is a from-scratch rebuild of
`AMBRScryptBenchlingReg2.py` as a Flask web app, since Render.com (and web
hosting generally) has no display for the original script's desktop popups
and no access to a visitor's local files.

- `core.py` &mdash; the parsing/merging/repacking logic, unchanged from the
  desktop script (same regexes, same rules: Plate/Well always come from the
  AMBR file's literal text, samples not in the scheme are flagged but never
  dropped, plate-format confirmation with contradiction override, pooled
  transposition across all source plates).
- `app.py` &mdash; the Flask app: a 4-step wizard (upload -> exclude volumes ->
  confirm plate format -> repack choice -> download).
- `templates/`, `static/style.css` &mdash; the web pages.
- `requirements.txt`, `Procfile`, `render.yaml` &mdash; deployment config for Render.

I ran this end-to-end against synthetic AMBR/Benchling data before handing
it over (upload -> volume exclusion -> plate-format confirm -> 24-to-96
repack -> both file downloads), and confirmed the generated workbook's rows
matched what the desktop script would have produced. I could not test the
exact `gunicorn app:app` production launch command inside this sandbox
(its network policy blocks installing `gunicorn` specifically here), so I
can't personally confirm that exact command succeeds end-to-end on Render
until you deploy it &mdash; but `gunicorn` is a very standard, actively
maintained Flask deployment server, and Render's own quick-start guides use
this exact `Procfile` pattern.

## 1. Put this on GitHub (no local git needed)

1. Go to [github.com/new](https://github.com/new) and create a new repository
   (e.g. `ambr-plate-mapping`). Keep it empty (don't add a README there).
2. On the new repo's page, click **uploading an existing file**.
3. Drag in every file from this folder, *preserving the folder structure*:
   `app.py`, `core.py`, `requirements.txt`, `Procfile`, `render.yaml`,
   `.gitignore`, `README.md`, and the `templates/` and `static/` folders with
   their contents. (GitHub's drag-and-drop upload preserves subfolders if you
   drag the folders themselves, not just files out of them.)
4. Commit directly to `main`.

If you'd rather use git from a terminal instead, the usual sequence is:
```bash
git init
git add .
git commit -m "AMBR plate mapping web app"
git branch -M main
git remote add origin https://github.com/<your-username>/ambr-plate-mapping.git
git push -u origin main
```

## 2. Deploy to Render.com

1. Go to [dashboard.render.com](https://dashboard.render.com) and sign in
   (GitHub login is the easiest option, since it can read your new repo
   directly).
2. Click **New +** -> **Web Service**.
3. Connect the `ambr-plate-mapping` GitHub repo you just created.
4. Render should auto-detect the included `render.yaml` and pre-fill the
   settings. If it asks manually instead, use:
   - **Environment**: Python 3
   - **Build Command**: `pip install -r requirements.txt`
   - **Start Command**: `gunicorn app:app`
   - **Instance Type**: Free is fine to start
5. Click **Create Web Service**. Render will build and deploy automatically;
   the first build takes a couple of minutes.
6. Once it says "Live", Render gives you a URL like
   `https://ambr-plate-mapping.onrender.com` &mdash; that's the tool. Anyone
   with the link can open it in a browser and use it with zero install.

Note on the free tier: Render's free web services spin down after periods of
inactivity and take ~30-60 seconds to wake back up on the next visit. If
that's a problem for daily lab use, Render's paid "Starter" instance tier
keeps it always-on.

## 3. Using it day to day

Open the Render URL, and:

1. Upload the AMBR timepoints `.csv` and Benchling Timepoint-Sample `.csv`
   (a sampling scheme `.xlsx` is optional &mdash; skip it to get chronological
   S00/S01/S02... numbering and an auto-generated scheme file).
2. Tick any sample volumes you want dropped (e.g. end-of-run 35 mL).
3. Confirm whether the AMBR run actually used 24-well or 96-well destination
   plates.
4. Choose whether to also generate an HPLC/SOA repack file, and to which
   plate format.
5. Download the resulting file(s) directly from the browser.

## Limitations carried over from the desktop version

- Plate format (24- vs 96-well) still can't be proven purely from well usage
  &mdash; an under-filled 96-well plate looks identical to a 24-well plate from
  well addresses alone, so the tool always asks you to confirm, while still
  overriding your answer if the data proves a contradiction (e.g. a well like
  F8 can only exist on a 96-well plate).
- Nothing from the AMBR file is ever silently dropped: samples not
  anticipated by a supplied sampling scheme are included and flagged rather
  than discarded.

## A note on privacy

Uploaded files are processed in server memory/temp storage for your session
only and are not shared between different visitors. That said, Render's free
tier is a shared public cloud service, not validated for regulated or
sensitive data &mdash; if this experiment data is subject to institutional data
handling policies, that's worth checking with your IT/compliance contact
before relying on this for real lab data, since I can't make that
determination on your behalf.
