"""CDMS login token - you refresh it by hand (HOW_TO_RUN.txt): paste it into token.txt.

get_token() returns the token from token.txt, or stops with a clear message when it is
missing or expired. The expiry is written inside the token, so we know before calling CDMS.
"""
import base64, json, os, sys, time
from datetime import datetime

HERE = os.path.dirname(os.path.abspath(__file__))
TOKEN_FILE = os.path.join(HERE, "token.txt")
HEADERS = {"content-type": "application/json", "accept": "application/json, text/plain, */*",
           "origin": "https://cdms.ripplr.in", "referer": "https://cdms.ripplr.in/",
           "user-agent": "Mozilla/5.0"}
REFRESH = "paste a fresh token into CDMS_Tool\\token.txt (HOW_TO_RUN.txt)"


def expires_at(tok):
    """Unix expiry from the JWT payload; 0 if it cannot be read."""
    try:
        p = tok.split(".")[1]
        return int(json.loads(base64.urlsafe_b64decode(p + "=" * (-len(p) % 4)))["exp"])
    except Exception:
        return 0


def get_token():
    """-> (token, 'valid until ...' text). Exits if missing or expired."""
    tok = ""
    if os.path.exists(TOKEN_FILE):
        with open(TOKEN_FILE) as f:
            tok = f.read().strip().replace("Bearer ", "").strip()
    if not tok:
        sys.exit(f"no CDMS token - {REFRESH}")
    exp = expires_at(tok)
    if exp and exp < time.time() + 60:
        sys.exit(f"CDMS token expired at {datetime.fromtimestamp(exp):%d-%b %H:%M} - {REFRESH}")
    return tok, (f"valid until {datetime.fromtimestamp(exp):%d-%b %H:%M}" if exp else "expiry unknown")


def _selftest():
    pay = base64.urlsafe_b64encode(json.dumps({"exp": 2000000000}).encode()).decode().rstrip("=")
    assert expires_at(f"h.{pay}.s") == 2000000000 and expires_at("not-a-jwt") == 0
    print("selftest OK")


if __name__ == "__main__":
    _selftest()
