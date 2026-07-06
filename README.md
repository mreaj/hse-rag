# HSE Assistant — shared team app

A public-URL chatbot over your **public** HSE SharePoint documents. An admin syncs the
library once into a shared index; everyone else just opens the URL and asks — no login.

**Stack (free/cheap, no GPU, no Azure infra approvals):**
Streamlit Community Cloud (hosting + public URL) · Qdrant Cloud (shared index) ·
Mistral API open models (embeddings + chat) · Microsoft Graph app-only (read SharePoint).

> Use this only for documents approved as public / for external cloud processing. Chat
> content is sent to Mistral and vectors are stored in Qdrant Cloud.

---

## One-time setup (≈20 min)

### 1. Qdrant Cloud (shared vector store)
- Create a free cluster at cloud.qdrant.io → copy the **URL** and an **API key**.

### 2. Mistral key
- Get an API key at console.mistral.ai.

### 3. SharePoint app access (you already have Sites.Selected)
- You need `TENANT_ID`, `CLIENT_ID`, `CLIENT_SECRET` for your app registration.
- The app must be **granted on the site** (one-time, by an admin) — Application permission,
  read is enough. `SITE_URL` is the full folder URL you want indexed.

### 4. Deploy to Streamlit Community Cloud (gives the public URL)
1. Push this folder to a GitHub repo.
2. Go to share.streamlit.io → **New app** → pick the repo → main file `app.py`.
3. In **Advanced settings → Secrets**, paste the contents of
   `.streamlit/secrets.toml.example` filled in with your real values.
4. Deploy. You get a public URL like `https://your-app.streamlit.app` — share it with your team.

### 5. Load the documents
- Open the app → **Admin** tab → enter the admin password → **Sync SharePoint library → index**.
- Wait for "Done…". Now the **Ask** tab works for everyone with the URL.

---

## Run locally (optional)
```bash
pip install -r requirements.txt
cp .streamlit/secrets.toml.example .streamlit/secrets.toml   # fill in values
streamlit run app.py
```

## Notes / limits
- **Index is shared and persistent** (Qdrant Cloud) — sync once, everyone benefits. Re-run
  Sync to pick up new/changed docs (re-embeds each run; fine for moderate libraries).
- **No OCR / xlsx** in this version (keeps it light for Streamlit Cloud). Scanned PDFs won't
  extract text. Ask me to add OCR if you need it.
- **Chat has no login** by design (docs are public). The **Admin** tab is password-gated so
  only you can Sync/Clear.
- **Open models:** `open-mistral-7b` (default) or `open-mixtral-8x7b` (set `CHAT_MODEL`).
- If your data ever stops being public, this stack is no longer appropriate — that needs the
  in-tenant / self-hosted design instead.
