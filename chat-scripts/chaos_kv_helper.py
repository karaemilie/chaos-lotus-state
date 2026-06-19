"""Minimal chaos_kv_helper: load/save state.json in chaos-lotus-state via GitHub API."""
import json, base64, os, urllib.request

REPO = "karaemilie/chaos-lotus-state"
PATH = "state.json"
def _token():
    t = os.environ.get("CHAOS_GH_TOKEN")
    if not t:
        t = open(os.path.expanduser("/home/claude/.gh_token")).read().strip()
    return t

def _api(url, method="GET", data=None):
    req = urllib.request.Request(url, method=method,
        headers={"Authorization": f"token {_token()}", "Accept": "application/vnd.github+json"})
    if data is not None:
        req.data = json.dumps(data).encode()
        req.add_header("Content-Type", "application/json")
    return json.load(urllib.request.urlopen(req))

def load_state():
    d = _api(f"https://api.github.com/repos/{REPO}/contents/{PATH}")
    return json.loads(base64.b64decode(d["content"]))

def _get_sha():
    d = _api(f"https://api.github.com/repos/{REPO}/contents/{PATH}")
    return d["sha"]

def save_state(state, commit_msg="Update state.json"):
    sha = _get_sha()
    content = base64.b64encode(json.dumps(state, indent=1).encode()).decode()
    body = {"message": commit_msg, "content": content, "sha": sha}
    return _api(f"https://api.github.com/repos/{REPO}/contents/{PATH}", method="PUT", data=body)
