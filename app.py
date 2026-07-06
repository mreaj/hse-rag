"""
HSE Assistant — shared team app
────────────────────────────────────────────────────────────────────────────
Public docs → one shared index → give the URL to your team, they just ask.

Stack (all free / cheap, no GPU, no Azure approvals):
  • Streamlit Community Cloud   → hosting + public URL
  • Qdrant Cloud                → shared, persistent vector store
  • Mistral API (open models)   → embeddings (mistral-embed) + chat (open-mistral / open-mixtral)
  • Microsoft Graph (app-only)  → read the public SharePoint library (Sites.Selected)

Two tabs:
  • Chat  — public, no login. Anyone with the URL asks questions.
  • Admin — password-gated. Sync the SharePoint library into the shared index.
"""
import os, io, re, json, time, uuid, hashlib
from pathlib import Path

import requests
import streamlit as st
from pypdf import PdfReader
import docx as _docx
from qdrant_client import QdrantClient
from qdrant_client.models import Distance, VectorParams, PointStruct

st.set_page_config(page_title="HSE Assistant", page_icon="🦺", layout="wide")

# ─────────────────────────────────────────────────────────── config (st.secrets)
def cfg(key, default=None):
    if key in st.secrets:
        return st.secrets[key]
    return os.getenv(key, default)

MISTRAL_KEY   = cfg("MISTRAL_API_KEY", "")
QDRANT_URL    = cfg("QDRANT_URL", "")
QDRANT_KEY    = cfg("QDRANT_API_KEY", "")
ADMIN_PW      = cfg("ADMIN_PASSWORD", "")

TENANT_ID     = cfg("TENANT_ID", "")
CLIENT_ID     = cfg("CLIENT_ID", "")
CLIENT_SECRET = cfg("CLIENT_SECRET", "")
SITE_URL      = cfg("SITE_URL", "")          # https://tenant.sharepoint.com/sites/HSE[/Library/Folder]

CHAT_MODEL    = cfg("CHAT_MODEL", "open-mistral-7b")   # open-mixtral-8x7b for stronger answers
EMBED_MODEL   = "mistral-embed"                          # 1024-dim
COLLECTION    = "hse_docs"
TOP_K         = int(cfg("TOP_K", 6))
CHUNK_SIZE    = 1000
CHUNK_OVERLAP = 200
SUPPORTED     = {".pdf", ".docx", ".txt", ".md"}
GRAPH         = "https://graph.microsoft.com/v1.0"

SUGGESTED = [
    "What PPE is required for confined space entry?",
    "Summarise the permit-to-work procedure.",
    "What are the steps in incident reporting?",
    "What does the standard say about working at height?",
]

# ─────────────────────────────────────────────────────────── Mistral helpers
def mistral_embed(texts):
    """Return list[list[float]] for a list of strings (batched)."""
    if isinstance(texts, str):
        texts = [texts]
    out = []
    for i in range(0, len(texts), 32):
        batch = texts[i:i + 32]
        r = requests.post(
            "https://api.mistral.ai/v1/embeddings",
            headers={"Authorization": f"Bearer {MISTRAL_KEY}"},
            json={"model": EMBED_MODEL, "input": batch}, timeout=120)
        r.raise_for_status()
        out.extend([d["embedding"] for d in r.json()["data"]])
    return out

def mistral_chat(system, user, temperature=0.1):
    r = requests.post(
        "https://api.mistral.ai/v1/chat/completions",
        headers={"Authorization": f"Bearer {MISTRAL_KEY}"},
        json={"model": CHAT_MODEL, "temperature": temperature, "max_tokens": 1200,
              "messages": [{"role": "system", "content": system},
                           {"role": "user", "content": user}]}, timeout=180)
    r.raise_for_status()
    return r.json()["choices"][0]["message"]["content"]

# ─────────────────────────────────────────────────────────── Qdrant
@st.cache_resource(show_spinner=False)
def qdrant():
    return QdrantClient(url=QDRANT_URL, api_key=QDRANT_KEY, timeout=60)

def ensure_collection():
    c = qdrant()
    names = [x.name for x in c.get_collections().collections]
    if COLLECTION not in names:
        c.create_collection(COLLECTION,
                            vectors_config=VectorParams(size=1024, distance=Distance.COSINE))

def index_count():
    try:
        return qdrant().count(COLLECTION).count
    except Exception:
        return 0

# ─────────────────────────────────────────────────────────── SharePoint (app-only, read-only)
def graph_app_token():
    r = requests.post(
        f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token",
        data={"client_id": CLIENT_ID, "client_secret": CLIENT_SECRET,
              "grant_type": "client_credentials",
              "scope": "https://graph.microsoft.com/.default"}, timeout=30)
    r.raise_for_status()
    return r.json()["access_token"]

def gget(url, token):
    for a in range(5):
        r = requests.get(url, headers={"Authorization": f"Bearer {token}"}, timeout=60)
        if r.status_code in (429, 503, 504):
            time.sleep(int(r.headers.get("Retry-After", 2 ** a)))
            continue
        return r
    return r

def parse_site_url(url):
    m = re.match(r"https?://([^/]+)((?:/sites/[^/?#]+)?)", url, re.I)
    host, site = (m.group(1), m.group(2) or "/") if m else (None, None)
    from urllib.parse import urlparse, unquote
    path = unquote(urlparse(url).path)
    path = re.sub(r"/Forms/[^/]*$", "", path)
    path = re.sub(r"/[^/]+\.aspx$", "", path).rstrip("/")
    return host, site, path

def iter_files(site_id, drive_id, token, folder="root"):
    stack = [folder]
    while stack:
        fid = stack.pop()
        url = (f"{GRAPH}/sites/{site_id}/drives/{drive_id}/items/{fid}/children"
               "?$top=200&$select=id,name,file,folder,webUrl,size")
        while url:
            r = gget(url, token); r.raise_for_status(); data = r.json()
            for it in data.get("value", []):
                if it.get("folder"):
                    stack.append(it["id"])
                elif it.get("file"):
                    yield it
            url = data.get("@odata.nextLink")

def download(drive_id, item_id, token):
    r = gget(f"{GRAPH}/drives/{drive_id}/items/{item_id}/content", token)
    r.raise_for_status()
    return r.content

# ─────────────────────────────────────────────────────────── parsing + chunking
def extract_text(name, data):
    ext = Path(name).suffix.lower()
    try:
        if ext == ".pdf":
            reader = PdfReader(io.BytesIO(data))
            return "\n".join((p.extract_text() or "") for p in reader.pages)
        if ext == ".docx":
            d = _docx.Document(io.BytesIO(data))
            return "\n".join(p.text for p in d.paragraphs)
        if ext in (".txt", ".md"):
            return data.decode("utf-8", errors="ignore")
    except Exception as e:
        print("parse error", name, e)
    return ""

def chunk(text):
    text = re.sub(r"\s+\n", "\n", text).strip()
    out, i = [], 0
    while i < len(text):
        piece = text[i:i + CHUNK_SIZE]
        if len(piece.strip()) > 50:
            out.append(piece)
        i += CHUNK_SIZE - CHUNK_OVERLAP
    return out

def point_id(item_id, ordinal):
    return str(uuid.uuid5(uuid.NAMESPACE_URL, f"{item_id}:{ordinal}"))

# ─────────────────────────────────────────────────────────── ingestion
def sync_library(log):
    ensure_collection()
    token = graph_app_token()
    host, site, folder_path = parse_site_url(SITE_URL)
    site_resp = gget(f"{GRAPH}/sites/{host}:{site}", token).json()
    site_id = site_resp["id"]
    drives = gget(f"{GRAPH}/sites/{site_id}/drives", token).json().get("value", [])
    # pick the drive matching the folder path, else the default
    target, subfolder = None, ""
    for d in drives:
        from urllib.parse import urlparse, unquote
        dp = unquote(urlparse(d.get("webUrl", "")).path).rstrip("/")
        if folder_path.startswith(dp):
            target = d; subfolder = folder_path[len(dp):].lstrip("/"); break
    if not target:
        target = drives[0]
    drive_id = target["id"]

    start_folder = "root"
    if subfolder:
        r = gget(f"{GRAPH}/sites/{site_id}/drives/{drive_id}/root:/{subfolder}", token)
        if r.status_code == 200:
            start_folder = r.json()["id"]

    scanned = indexed = chunks_total = 0
    for it in iter_files(site_id, drive_id, token, start_folder):
        scanned += 1
        if Path(it["name"]).suffix.lower() not in SUPPORTED:
            continue
        try:
            text = extract_text(it["name"], download(drive_id, it["id"], token))
            pieces = chunk(text)
            if not pieces:
                log(f"skip (no text): {it['name']}"); continue
            vecs = mistral_embed(pieces)
            pts = [PointStruct(id=point_id(it["id"], j), vector=v,
                               payload={"item_id": it["id"], "name": it["name"],
                                        "web_url": it.get("webUrl", ""), "text": ch})
                   for j, (ch, v) in enumerate(zip(pieces, vecs))]
            qdrant().upsert(COLLECTION, points=pts)
            indexed += 1; chunks_total += len(pts)
            log(f"indexed: {it['name']} ({len(pts)} chunks)")
        except Exception as e:
            log(f"FAILED {it['name']}: {e}")
    log(f"Done. Scanned {scanned}, indexed {indexed} docs, {chunks_total} chunks.")

# ─────────────────────────────────────────────────────────── RAG
def answer(question):
    if any(k in question.lower() for k in ("how many document", "number of document")):
        return f"There are **{index_count()} chunks** indexed in the knowledge base.", []
    qv = mistral_embed(question)[0]
    hits = qdrant().query_points(COLLECTION, query=qv, limit=TOP_K, with_payload=True).points
    if not hits:
        return "I don't have any indexed documents that cover this yet.", []
    ctx, refs, seen = [], [], set()
    for i, h in enumerate(hits):
        p = h.payload
        ctx.append(f"[{i+1}] (from: {p['name']})\n{p['text']}")
        key = (p["name"], p.get("web_url", ""))
        if key not in seen:
            seen.add(key); refs.append({"name": p["name"], "web_url": p.get("web_url", "")})
    system = ("You are an HSE (Health, Safety & Environment) assistant. Answer ONLY from the "
              "context. Cite the source filename for each fact. If the answer isn't in the "
              "context, say so — never invent information.")
    user = "Context:\n" + "\n\n---\n\n".join(ctx) + f"\n\nQuestion: {question}\n\nAnswer:"
    return mistral_chat(system, user), refs

# ─────────────────────────────────────────────────────────── UI
st.title("🦺 HSE Assistant")

missing = [k for k, v in {"MISTRAL_API_KEY": MISTRAL_KEY, "QDRANT_URL": QDRANT_URL,
                          "QDRANT_API_KEY": QDRANT_KEY}.items() if not v]
if missing:
    st.error("Missing configuration: " + ", ".join(missing) +
             ". Set them in the app's Secrets (see README).")
    st.stop()

tab_chat, tab_admin = st.tabs(["💬 Ask", "⚙️ Admin"])

with tab_chat:
    n = index_count()
    if n == 0:
        st.info("The knowledge base is empty. An admin needs to run **Sync** in the Admin tab first.")
    else:
        st.caption(f"{n} chunks indexed · ask anything about the HSE documents")

    st.session_state.setdefault("messages", [])
    for m in st.session_state.messages:
        with st.chat_message(m["role"]):
            st.markdown(m["content"])

    if not st.session_state.messages and n:
        cols = st.columns(2)
        for i, s in enumerate(SUGGESTED):
            if cols[i % 2].button(s, key=f"sug{i}", use_container_width=True):
                st.session_state._pending = s
                st.rerun()

    q = st.chat_input("Ask about HSE procedures, standards, PPE, permits…")
    if not q and st.session_state.get("_pending"):
        q = st.session_state.pop("_pending")
    if q:
        st.session_state.messages.append({"role": "user", "content": q})
        with st.chat_message("user"):
            st.markdown(q)
        with st.chat_message("assistant"):
            with st.spinner("Searching documents…"):
                text, refs = answer(q)
            if refs:
                text += "\n\n**Sources**\n" + "\n".join(
                    f"- [{r['name']}]({r['web_url']})" if r["web_url"] else f"- {r['name']}"
                    for r in refs)
            st.markdown(text)
            st.session_state.messages.append({"role": "assistant", "content": text})

with tab_admin:
    if ADMIN_PW and not st.session_state.get("is_admin"):
        pw = st.text_input("Admin password", type="password")
        if st.button("Unlock"):
            st.session_state.is_admin = (pw == ADMIN_PW)
            st.rerun()
        st.stop()

    st.write(f"**Site:** {SITE_URL or '(SITE_URL not set)'}")
    st.write(f"**Index:** {index_count()} chunks")
    if not (TENANT_ID and CLIENT_ID and CLIENT_SECRET and SITE_URL):
        st.warning("Set TENANT_ID, CLIENT_ID, CLIENT_SECRET and SITE_URL in Secrets to enable Sync.")
    else:
        if st.button("🔄 Sync SharePoint library → index"):
            area = st.empty(); buf = []
            def log(m): buf.append(str(m)); area.code("\n".join(buf[-300:]))
            with st.spinner("Crawling and indexing… (this can take a while)"):
                try:
                    sync_library(log)
                except Exception as e:
                    log(f"ERROR: {e}")
    st.divider()
    if st.button("🗑️ Clear entire index"):
        try:
            qdrant().delete_collection(COLLECTION); ensure_collection()
            st.success("Index cleared.")
        except Exception as e:
            st.error(str(e))
