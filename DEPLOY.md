# Deploying to Streamlit Community Cloud

This gets you one URL you can send to anyone at riivo - no install, no
login required to use it (though Streamlit Cloud itself needs you, the
deployer, to sign in once to set it up).

These steps need to happen from your own machine/accounts - they involve
logging into GitHub and Streamlit Cloud as yourself, which I can't do on
your behalf.

## 1. Get the code onto GitHub

The unzipped folder is already a git repo with one commit. You just need
to point it at a real GitHub repository:

1. On github.com, create a **new, empty** repository (don't add a
   README/license/gitignore - the folder already has one). Either under
   your own account or a riivo org, and either public or private - both
   work with Streamlit Community Cloud's free tier.
2. In a terminal, `cd` into the unzipped folder, then:
   ```bash
   git remote add origin https://github.com/<your-account-or-org>/<repo-name>.git
   git branch -M main
   git push -u origin main
   ```

## 2. Deploy on Streamlit Community Cloud

1. Go to [share.streamlit.io](https://share.streamlit.io) and sign in with
   GitHub (free - this is what creates your Streamlit Cloud account).
2. Click **New app** → **Deploy a public app from GitHub** (or the private
   variant if your repo is private).
3. Pick the repository you just pushed, branch `main`, main file path
   `app.py`.
4. Click **Deploy**. It installs `requirements.txt` and builds - takes a
   couple of minutes the first time.
5. You'll land on a URL like `https://<something>.streamlit.app` - that's
   the link to send around. Anyone who opens it sees the upload screen
   from a fresh browser session; nothing is shared or visible between
   different people using the app at the same time.

## 3. Updating it later

Any time the tool changes: commit and `git push` to `main` on the same
repo, and Streamlit Cloud redeploys automatically within a minute or two.
No redeploy step to remember.

## Before you send the link around

Worth a beat of thought, not a blocker: uploaded files are processed in
memory on Streamlit's infrastructure for the duration of that browser
session, then discarded - nothing is written to disk or kept after the
tab closes. That's a step outside riivo's own network, though, which is
different from everyone running this locally. If riivo has a policy on
where client invoice data is allowed to transit, it's worth a quick check
with whoever owns that before this goes out company-wide. If that's a
blocker, the packaged-desktop-app option from earlier keeps everything on
each person's own machine instead.

## Restricting who can see it, later

The free tier doesn't support login-gating a specific app - anyone with
the URL can open it. Since that's fine for now, nothing to do here. If
that changes later, Streamlit Community Cloud has no paid tier for this
specifically, but a small app-level password check (a text input compared
against a shared secret, gating the rest of the page) could be added to
`app.py` in about 10 lines if it's ever needed - just say so.
