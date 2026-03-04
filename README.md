# WV BBL Tax Reporter

Automated monthly BBL tax report generator for Mountain State Brewing Co.

Upload your QuickBooks sales export → get a completed WV Upload Template instantly.
QB Report is a saved custom report:  WV Monthly BBL Tax

---

## Deploy to Streamlit Cloud (15 min setup)

### Step 1 — Create a GitHub account
If you don't have one: https://github.com/signup (free)

### Step 2 — Create a new repository
1. Go to https://github.com/new
2. Name it `bbl-reporter` (or anything you like)
3. Set it to **Private**
4. Click **Create repository**

### Step 3 — Upload the app files
1. On your new repo page, click **Add file → Upload files**
2. Upload both files from this folder:
   - `app.py`
   - `requirements.txt`
3. Click **Commit changes**

### Step 4 — Deploy on Streamlit Cloud
1. Go to https://share.streamlit.io and sign in with your GitHub account
2. Click **New app**
3. Select your `bbl-reporter` repository
4. Set **Main file path** to `app.py`
5. Click **Deploy**

Your app will be live at a URL like:
`https://your-username-bbl-reporter-app-xxxx.streamlit.app`

Bookmark it — that's your tool from now on!

---

## Every month — how to use it

1. Open your Streamlit URL
2. Upload `__WV_Upload_Template.xlsx` (keep a copy saved somewhere handy)
3. Upload your monthly QuickBooks export (e.g. `Nov_25.xlsx`)
4. Click **Process Files**
5. Download the `_Final.xlsx` — ready to submit

You can upload multiple months at once.

---

## Adding a new distributor

If a new distributor shows up that isn't recognized:
1. The app will warn you which distributor is unmapped
2. Expand the **⚙️ Manage Distributor ABCA Numbers** section
3. Enter the customer name exactly as it appears in QuickBooks, their ABCA license number, and display name
4. Click **Add Distributor** and re-run

> Note: Added distributors only persist for your current browser session.
> To save them permanently, ask Claude to add them to the `DEFAULT_ABCA` dictionary in `app.py` and re-upload the file to GitHub.

---

## BBL Multipliers used

| Container | Multiplier |
|-----------|-----------|
| 1/2 BBL keg | 0.5 |
| 1/6 BBL keg | 0.166667 |
| Cans (24/12oz case) | 0.07258 |
