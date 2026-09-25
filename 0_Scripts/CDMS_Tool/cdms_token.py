"""CDMS login token. token.txt = access token (24 h), refresh.txt = refresh token (~4 days).

get_token() returns the token from token.txt; if it has expired, it gets a new one from
refresh.txt by itself. Only when both have expired does it stop with a clear message -
then log in once and paste both (HOW_TO_RUN.txt). No password is stored.
"""
import base64, json, os, sys, time, urllib.request
from datetime import datetime

HERE = os.path.dirname(os.path.abspath(__file__))
TOKEN_FILE = os.path.join(HERE, "token.txt")
REFRESH_FILE = os.path.join(HERE, "refresh.txt")
REFRESH_API = "https://api-cdms.ripplr.in/api/champ/refresh"
HEADERS = {"content-type": "application/json", "accept": "application/json, text/plain, */*",
           "origin": "https://cdms.ripplr.in", "referer": "https://cdms.ripplr.in/",
           "user-agent": "Mozilla/5.0"}
REFRESH = "log in to CDMS and paste 'access' into CDMS_Tool\\token.txt and 'refresh' into refresh.txt (HOW_TO_RUN.txt)"


def expires_at(tok):
    """Unix expiry from the JWT payload; 0 if it cannot be read."""
    try:
        p = tok.split(".")[1]
        return int(json.loads(base64.urlsafe_b64decode(p + "=" * (-len(p) % 4)))["exp"])
    except Exception:
        return 0


def _read(path):
    if not os.path.exists(path):
        return ""
    with open(path) as f:
        return f.read().strip().replace("Bearer ", "").strip()


def _write(path, tok):
    with open(path, "w") as f:
        f.write(tok)


def refresh_access():
    """Use refresh.txt (valid ~4 days) to get a new access token; saves it. '' if not possible."""
    ref = _read(REFRESH_FILE)
    if not ref or expires_at(ref) < time.time() + 60:
        return ""
    body = json.dumps({"refresh": ref, "refresh_token": ref}).encode()
    req = urllib.request.Request(REFRESH_API, data=body, method="POST",
                                 headers={**HEADERS, "authorization": "Bearer " + ref})
    try:
        with urllib.request.urlopen(req, timeout=60) as r:
            tok = json.loads(r.read().decode())["data"]["token"]
    except Exception as e:
        print(f"CDMS auto-refresh failed ({type(e).__name__})")
        return ""
    if isinstance(tok, dict):                 # login-style {access, refresh}
        if tok.get("refresh"):
            _write(REFRESH_FILE, tok["refresh"])
        tok = tok.get("access", "")
    if tok:
        _write(TOKEN_FILE, tok)
        print("CDMS token auto-refreshed")
    return tok


def get_token():
    """-> (token, 'valid until ...' text). Auto-refreshes when expired; exits if that fails too."""
    tok = _read(TOKEN_FILE)
    exp = expires_at(tok)
    if not tok or (exp and exp < time.time() + 60):
        tok = refresh_access() or tok
        exp = expires_at(tok)
    if not tok:
        sys.exit(f"no CDMS token - {REFRESH}")
    if exp and exp < time.time() + 60:
        sys.exit(f"CDMS token expired at {datetime.fromtimestamp(exp):%d-%b %H:%M} - {REFRESH}")
    return tok, (f"valid until {datetime.fromtimestamp(exp):%d-%b %H:%M}" if exp else "expiry unknown")


def _selftest():
    pay = base64.urlsafe_b64encode(json.dumps({"exp": 2000000000}).encode()).decode().rstrip("=")
    assert expires_at(f"h.{pay}.s") == 2000000000 and expires_at("not-a-jwt") == 0
    print("selftest OK")


if __name__ == "__main__":
    _selftest()
