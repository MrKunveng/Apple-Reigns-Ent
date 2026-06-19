# Apple Reigns Enterprise — Billing Suite

A branded Streamlit app with three modes:
- **Invoice** — editable line items, VAT, discount, payment details, A4 PDF.
- **Payment receipt** — payer, amount, mode of payment, amount-in-words, PAID stamp, A4 PDF.
- **Revenue analytics** — last 7-day / 30-day / all-time inflows, daily + by-mode
  charts, recent-payments table, CSV export, and manual payment logging.

Every recorded payment (from a receipt or the manual logger) feeds the analytics.

## Run locally
```bash
pip install -r requirements.txt
streamlit run streamlit_app.py
```

## Revenue ledger storage
The app stores payments in **Google Sheets** when configured (survives
redeploys), otherwise in a local `transactions.csv`. The active backend is shown
on the Revenue analytics tab.

### Set up Google Sheets (recommended for cloud)
1. Create a Google Sheet; copy its URL.
2. In **Google Cloud Console**: create a project → enable **Google Sheets API**
   and **Google Drive API**.
3. Create a **Service Account** → **Keys** → **Add key → JSON** (downloads a key file).
4. **Share the Google Sheet** with the service account's `client_email`
   (give it **Editor** access).
5. Add credentials to Streamlit secrets:
   - Local: copy `.streamlit/secrets.toml.example` → `.streamlit/secrets.toml`
     and fill in the sheet URL + JSON key fields.
   - Cloud: app → **Settings → Secrets**, paste the same TOML.

> Without secrets the app silently falls back to the local CSV — nothing breaks,
> but on Streamlit Cloud that CSV resets on every redeploy/reboot.

## Deploy on Streamlit Community Cloud
1. Push this folder to a **GitHub repo** (include `streamlit_app.py`,
   `requirements.txt`, `assets_logo.png`, `Logo1.png`, `.streamlit/config.toml`).
   Do **not** commit `secrets.toml` (it is gitignored).
2. https://share.streamlit.io → **Create app** → sign in with GitHub.
3. Pick repo/branch, set **Main file path** = `streamlit_app.py`, add Secrets, **Deploy**.

> Netlify cannot host this — it only serves static sites + short serverless
> functions, not a long-running Python server. Use Streamlit Community Cloud
> (or Render / Railway / Hugging Face Spaces).

## Files
| File | Purpose |
|------|---------|
| `streamlit_app.py` | The app (invoice · receipt · analytics) |
| `requirements.txt` | Python deps |
| `.streamlit/config.toml` | Brand theme |
| `.streamlit/secrets.toml.example` | Google Sheets credentials template |
| `assets_logo.png` / `Logo1.png` | Logo (optimised / original) |
