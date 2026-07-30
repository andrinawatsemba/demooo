"""
Persistence layer for Streamlit Community Cloud's free tier, which
wipes local file writes on every redeploy/reboot.

Setup (one-time, done by me and never touches end users):
  1. Create a GitHub personal access token with 'repo' write scope.
  2. In Streamlit Cloud -> App settings -> Secrets, add:

     [github]
     token  = "ghp_..."
     repo   = "your-username/your-repo"
     branch = "main"

If these secrets aren't set, every function here quietly no-ops and
prints an info line - the app still works locally, it just won't
persist across cloud restarts. Nothing breaks if this isn't configured
yet.
"""

import base64
import streamlit as st
import requests

API_BASE = "https://api.github.com"


def _get_config():
    try:
        gh = st.secrets["github"]
        return gh["token"], gh["repo"], gh.get("branch", "main")
    except Exception:
        return None, None, None


def push_file_to_github(local_path, repo_path, message=None):
    """Commit a local file's current contents to the configured repo.
    Safe to call even if GitHub isn't configured - just no-ops."""
    token, repo, branch = _get_config()
    if not token:
        print("[INFO] GitHub persistence not configured - skipping push (local-only mode)")
        return False

    try:
        with open(local_path, "rb") as f:
            content_b64 = base64.b64encode(f.read()).decode()
    except FileNotFoundError:
        print(f"[WARNING] Cannot push - local file not found: {local_path}")
        return False

    url = f"{API_BASE}/repos/{repo}/contents/{repo_path}"
    headers = {"Authorization": f"token {token}", "Accept": "application/vnd.github+json"}

    existing = requests.get(url, headers=headers, params={"ref": branch})
    sha = existing.json().get("sha") if existing.status_code == 200 else None

    payload = {
        "message": message or f"Update {repo_path} via NBS Ad Tracker",
        "content": content_b64,
        "branch": branch,
    }
    if sha:
        payload["sha"] = sha

    resp = requests.put(url, headers=headers, json=payload)
    if resp.status_code in (200, 201):
        print(f"[OK] Pushed {repo_path} to GitHub")
        return True

    print(f"[WARNING] GitHub push failed ({resp.status_code}): {resp.text[:200]}")
    return False


def pull_file_from_github(repo_path, local_path):
    """Fetch the latest version of a file from the repo down to the
    local (ephemeral) filesystem. Call this at app startup, before
    checking whether local files exist - a fresh container has
    nothing locally until this runs.
    Returns True if a file was pulled, False if not configured or
    the file doesn't exist in the repo yet (e.g. very first run)."""
    token, repo, branch = _get_config()
    if not token:
        return False

    url = f"{API_BASE}/repos/{repo}/contents/{repo_path}"
    headers = {"Authorization": f"token {token}", "Accept": "application/vnd.github+json"}
    resp = requests.get(url, headers=headers, params={"ref": branch})

    if resp.status_code != 200:
        print(f"[INFO] No existing {repo_path} in GitHub yet (status {resp.status_code}) - starting fresh")
        return False

    content_b64 = resp.json().get("content", "")
    local_path.parent.mkdir(parents=True, exist_ok=True)
    with open(local_path, "wb") as f:
        f.write(base64.b64decode(content_b64))

    print(f"[OK] Pulled latest {repo_path} from GitHub")
    return True
