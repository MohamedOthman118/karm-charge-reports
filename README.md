# ⚡ KARM CHARGE — EV Charger Operations Report Generator

A web app that turns AI-generated JSON into a professionally formatted Word (.docx) operations report.

## How it works

1. The user copies an AI prompt from the website.
2. They paste it (plus their raw field notes) into ChatGPT / Claude / Gemini.
3. The AI returns structured JSON.
4. They paste that JSON back into the site → download a polished `.docx`.

No API keys needed. No external services called.

---

## 🚀 Deploy to the internet (free, ~5 minutes)

The easiest way to give this app a public URL is **Streamlit Community Cloud**. It's free, has no credit card requirement, and is built specifically for Streamlit apps.

### Step 1 — Put the code on GitHub

If you don't already have a GitHub account, create one at https://github.com.

1. Click the green **New** button to create a new repository.
2. Name it something like `karm-charge-reports`.
3. Choose **Public** (Community Cloud only deploys from public repos on the free tier — if you need private, see paid options below).
4. Click **Create repository**.
5. Upload these three files using the **Add file → Upload files** button:
   - `app.py`
   - `requirements.txt`
   - `.streamlit/config.toml`  *(create the `.streamlit` folder first; on GitHub web UI, type `.streamlit/config.toml` in the filename field when uploading)*

### Step 2 — Deploy on Streamlit Community Cloud

1. Go to **https://share.streamlit.io** and click **Sign in with GitHub**.
2. Authorise Streamlit to access your repos.
3. Click **Create app → Deploy a public app from GitHub**.
4. Fill in:
   - **Repository:** `your-username/karm-charge-reports`
   - **Branch:** `main`
   - **Main file path:** `app.py`
   - **App URL** *(optional):* pick a custom subdomain, e.g. `karm-charge-reports`
5. Click **Deploy**.

⏱️ First deploy takes ~2 minutes (installing Python dependencies). After that, you'll get a public URL like:

```
https://karm-charge-reports.streamlit.app
```

Share that URL with anyone. They can use the tool from any browser, no install required.

### Step 3 — Updating the app later

Any time you push a change to the GitHub repo (edit `app.py` directly on GitHub, or push from your computer), Streamlit Community Cloud automatically rebuilds and redeploys within a minute. No manual steps needed.

---

## 🔒 Alternative: private deployments

If you don't want the code public, here are options:

| Platform | Free tier | Notes |
|----------|-----------|-------|
| **Hugging Face Spaces** | Yes | Supports private Spaces. Upload files via web UI. Same `requirements.txt`. |
| **Render** | Yes (cold starts after 15 min idle) | Connect a private GitHub repo. Use start command: `streamlit run app.py --server.port $PORT --server.address 0.0.0.0` |
| **Railway** | $5 free credit/month | Same approach as Render. |
| **Fly.io** | Yes (small allowance) | Needs a `Dockerfile`. |
| **Streamlit for Teams** | Paid | Private apps on `share.streamlit.io`. |
| **Self-hosted (VPS)** | ~$5/mo (DigitalOcean, Hetzner) | Run `streamlit run app.py` behind nginx + a domain. |

---

## 💻 Run locally (for testing or offline use)

```bash
pip install -r requirements.txt
streamlit run app.py
```

The app opens automatically in your browser at `http://localhost:8501`. To make it accessible to other people on your local network:

```bash
streamlit run app.py --server.address 0.0.0.0
```

Then they can visit `http://YOUR-COMPUTER-IP:8501`.

---

## 📁 File structure

```
.
├── app.py                  # The full Streamlit app
├── requirements.txt        # Python dependencies
├── .streamlit/
│   └── config.toml         # Theme + server settings
└── README.md               # This file
```

---

## 🐛 Troubleshooting

**"Module not found: docx"** — Streamlit Cloud didn't pick up `requirements.txt`. Make sure it's in the **repo root**, not inside a subfolder, and that it's named exactly `requirements.txt` (lowercase).

**App boots but shows "Please wait..."** — First deploy is installing dependencies; wait 60–90 seconds.

**Deploy succeeded but the URL shows an error page** — Click **Manage app** in the bottom right of the deployed app → **Logs** to see the Python traceback.

**JSON parsing fails for some users** — The app strips ` ```json ... ``` ` fences automatically, but if the AI added commentary before/after the JSON, the user needs to delete that. Consider tightening the prompt with: *"Return ONLY the JSON. No text before or after."* (already included in the prompt template.)
