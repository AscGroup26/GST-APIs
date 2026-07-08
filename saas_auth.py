"""
SaaS Authentication and User Management
GSTR-1 Dashboard — ASC Consulting Pvt. Ltd.
"""

import hashlib
import os
import base64
import json
import secrets
import smtplib
import yaml
from datetime import datetime, date
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path

import streamlit as st

# ─────────────────────────────────────────────────────────────────
# FILE PATHS
# ─────────────────────────────────────────────────────────────────
_DIR              = Path(__file__).parent
_CONFIG           = _DIR / "auth_config.yaml"
_DATA_DIR         = _DIR / ".saas_data"
_LOGIN_HISTORY    = _DATA_DIR / "login_history.json"
_DOWNLOAD_HISTORY = _DATA_DIR / "download_history.json"
_ANNOUNCEMENTS    = _DATA_DIR / "announcements.json"
_SESSIONS_FILE    = _DATA_DIR / "sessions.json"

_DATA_DIR.mkdir(exist_ok=True)

# ─────────────────────────────────────────────────────────────────
# .env loader (no external dependency)
# ─────────────────────────────────────────────────────────────────
def _load_dotenv():
    env_file = _DIR / ".env"
    if not env_file.exists():
        return
    with open(env_file, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, _, val = line.partition("=")
            key = key.strip()
            # Strip surrounding quotes; do NOT strip inline # (part of password)
            val = val.strip()
            if (val.startswith('"') and val.endswith('"')) or \
               (val.startswith("'") and val.endswith("'")):
                val = val[1:-1]
            if key:
                os.environ[key] = val   # always write — allows .env edits without full restart

_load_dotenv()

# ─────────────────────────────────────────────────────────────────
# SMTP — email verification
# ─────────────────────────────────────────────────────────────────
def _smtp_cfg() -> dict:
    """Read SMTP settings from st.secrets or environment variables."""
    try:
        sec = st.secrets.get("smtp", {})
    except Exception:
        sec = {}
    return {
        "host"      : sec.get("host",      os.getenv("SMTP_HOST", "")),
        "port"      : int(sec.get("port",  os.getenv("SMTP_PORT", 587))),
        "user"      : sec.get("user",      os.getenv("SMTP_USER", "")),
        "password"  : sec.get("password",  os.getenv("SMTP_PASS", "")),
        "from_name" : sec.get("from_name", os.getenv("SMTP_FROM_NAME", "ASC Consulting")),
        "app_url"   : sec.get("app_url",   os.getenv("APP_URL", "http://localhost:8501")),
    }

def _smtp_configured() -> bool:
    cfg = _smtp_cfg()
    return bool(cfg["host"] and cfg["user"] and cfg["password"])

def send_verification_email(to_email: str, to_name: str, token: str) -> bool:
    """Send a verification e-mail. Returns True on success."""
    _load_dotenv()          # re-read .env in case it changed since startup
    cfg = _smtp_cfg()
    if not _smtp_configured():
        print("[MAIL ERROR] SMTP not configured — check .env SMTP_HOST/SMTP_USER/SMTP_PASS")
        return False
    verify_url = cfg["app_url"].rstrip("/") + f"/?verify={token}"
    html = f"""
<html><body style="font-family:Arial,sans-serif;background:#f4f6fb;padding:40px 20px;">
<div style="max-width:520px;margin:0 auto;background:#fff;border-radius:12px;
            overflow:hidden;box-shadow:0 2px 12px rgba(0,0,0,.08);">
  <div style="background:#1f2f60;padding:28px 32px;text-align:center;">
    <h1 style="color:#fff;margin:0;font-size:22px;">ASC Consulting</h1>
    <p style="color:rgba(255,255,255,.65);margin:5px 0 0;font-size:13px;">GSTR-1 Dashboard</p>
  </div>
  <div style="padding:32px;">
    <h2 style="color:#1f2f60;margin:0 0 14px;font-size:20px;">Verify your email address</h2>
    <p style="color:#4a5568;line-height:1.7;">Hi {to_name},</p>
    <p style="color:#4a5568;line-height:1.7;">
      Thank you for registering. Click the button below to verify your
      email address and activate your account.
    </p>
    <div style="text-align:center;margin:28px 0;">
      <a href="{verify_url}"
         style="background:#1f2f60;color:#fff;text-decoration:none;
                padding:14px 36px;border-radius:8px;font-size:15px;
                font-weight:700;display:inline-block;">
        ✓ &nbsp;Verify Email Address
      </a>
    </div>
    <p style="color:#94adc8;font-size:12px;line-height:1.6;">
      If you did not create an account, please ignore this email.
    </p>
  </div>
  <div style="background:#f8fafc;padding:14px 32px;text-align:center;border-top:1px solid #e2e8f0;">
    <p style="color:#94adc8;font-size:11px;margin:0;">
      © 2026 ASC Consulting Pvt. Ltd. · All rights reserved
    </p>
  </div>
</div>
</body></html>"""
    msg = MIMEMultipart("alternative")
    msg["Subject"] = "Verify your email — ASC GSTR-1 Dashboard"
    msg["From"]    = f"{cfg['from_name']} <{cfg['user']}>"
    msg["To"]      = to_email
    msg.attach(MIMEText(html, "html"))
    try:
        with smtplib.SMTP(cfg["host"], cfg["port"], timeout=10) as srv:
            srv.ehlo()
            srv.starttls()
            srv.ehlo()
            srv.login(cfg["user"], cfg["password"])
            srv.sendmail(cfg["user"], to_email, msg.as_string())
        print(f"[MAIL OK] Verification email sent to {to_email}")
        return True
    except Exception as _e:
        print(f"[MAIL ERROR] Failed to send to {to_email}: {_e}")
        return False

def verify_email_token(token: str) -> tuple:
    """Mark email as verified for the user matching this token."""
    users = load_users()
    for username, data in users.items():
        if data.get("verification_token") == token:
            users[username]["email_verified"]     = True
            users[username]["verification_token"] = None
            save_users(users)
            return True, data.get("name", "User")
    return False, "Invalid or expired verification link."

# ─────────────────────────────────────────────────────────────────
# PASSWORD — PBKDF2-SHA256 (built-in, no extra deps)
# ─────────────────────────────────────────────────────────────────
def _hash_password(password: str) -> str:
    salt = os.urandom(16)
    dk   = hashlib.pbkdf2_hmac("sha256", password.encode(), salt, 200_000)
    return base64.b64encode(salt + dk).decode()

def _verify_password(password: str, stored: str) -> bool:
    try:
        raw  = base64.b64decode(stored.encode())
        salt, dk = raw[:16], raw[16:]
        return hashlib.pbkdf2_hmac("sha256", password.encode(), salt, 200_000) == dk
    except Exception:
        return False

# ─────────────────────────────────────────────────────────────────
# USER CRUD
# ─────────────────────────────────────────────────────────────────
def load_users() -> dict:
    if not _CONFIG.exists():
        return {}
    with open(_CONFIG, encoding="utf-8") as f:
        data = yaml.safe_load(f) or {}
    return data.get("users", {})

def save_users(users: dict):
    with open(_CONFIG, "w", encoding="utf-8") as f:
        yaml.dump({"users": users}, f, default_flow_style=False, allow_unicode=True)

def _find_username_by_email(email: str):
    """Return the username key for a given email address, or None."""
    users = load_users()
    email_lower = email.strip().lower()
    for uname, data in users.items():
        if data.get("email", "").strip().lower() == email_lower:
            return uname
    return None

def add_user(username: str, name: str, email: str, password: str,
             role: str = "client", expiry=None) -> tuple:
    users = load_users()
    if username in users:
        return False, "Username already exists"
    if _find_username_by_email(email):
        return False, "An account with this email address already exists"
    is_admin   = (role == "admin")
    token      = None if is_admin else secrets.token_urlsafe(32)
    users[username] = {
        "name"               : name,
        "email"              : email,
        "password"           : _hash_password(password),
        "role"               : role,
        "active"             : True,
        "expiry"             : str(expiry) if expiry else None,
        "created"            : str(date.today()),
        "email_verified"     : is_admin,      # admins pre-verified
        "verification_token" : token,
    }
    if not is_admin:
        # Send email BEFORE saving — roll back naturally if it fails
        mail_ok = send_verification_email(email, name, token)
        if not mail_ok:
            return False, "Account could not be created: verification email failed to send. Please check your email address or contact the administrator."
    save_users(users)
    return True, "User created successfully"

def remove_user(username: str) -> tuple:
    users = load_users()
    if username not in users:
        return False, "User not found"
    del users[username]
    save_users(users)
    return True, "User removed"

def update_password(username: str, new_password: str) -> bool:
    users = load_users()
    if username not in users:
        return False
    users[username]["password"] = _hash_password(new_password)
    save_users(users)
    return True

def update_expiry(username: str, expiry) -> bool:
    users = load_users()
    if username not in users:
        return False
    users[username]["expiry"] = str(expiry) if expiry else None
    save_users(users)
    return True

def toggle_active(username: str, active: bool) -> bool:
    users = load_users()
    if username not in users:
        return False
    users[username]["active"] = active
    save_users(users)
    return True

def update_role(username: str, role: str) -> tuple:
    users = load_users()
    if username not in users:
        return False, "User not found"
    if users[username].get("role") == role:
        return False, f"User is already a {role}"
    # Never leave the system without an active admin
    if users[username].get("role") == "admin" and role != "admin":
        other_admins = [u for u, d in users.items()
                        if u != username and d.get("role") == "admin" and d.get("active", True)]
        if not other_admins:
            return False, "Cannot demote the only active admin"
    users[username]["role"] = role
    if role == "admin":
        # same rule as add_user: admins are pre-verified
        users[username]["email_verified"]     = True
        users[username]["verification_token"] = None
    save_users(users)
    return True, f"Role changed to {role}"

def set_email_verified(username: str) -> tuple:
    users = load_users()
    if username not in users:
        return False, "User not found"
    users[username]["email_verified"]     = True
    users[username]["verification_token"] = None
    save_users(users)
    return True, "Email marked as verified"

def resend_verification(username: str) -> tuple:
    users = load_users()
    if username not in users:
        return False, "User not found"
    u = users[username]
    if u.get("email_verified", True):
        return False, "Email is already verified"
    email = u.get("email", "").strip()
    if not email:
        return False, "User has no email address on file"
    token = u.get("verification_token") or secrets.token_urlsafe(32)
    users[username]["verification_token"] = token
    save_users(users)
    if send_verification_email(email, u.get("name", username), token):
        return True, f"Verification email sent to {email}"
    return False, "Failed to send verification email — check SMTP settings (journalctl: grep MAIL)"

def _is_expired(user_data: dict) -> bool:
    exp = user_data.get("expiry")
    if not exp or str(exp).lower() == "none":
        return False
    try:
        return date.today() > date.fromisoformat(str(exp))
    except Exception:
        return False

def _ensure_default_admin():
    """Create default admin account on first run."""
    if not _CONFIG.exists() or not load_users():
        add_user("admin", "Admin", "admin@ascgroup.in", "Admin@123", "admin")

_ensure_default_admin()

# ─────────────────────────────────────────────────────────────────
# PERSISTENT SESSION — survives browser refreshes via ?_sid=TOKEN
# ─────────────────────────────────────────────────────────────────
def _load_sessions() -> dict:
    if not _SESSIONS_FILE.exists():
        return {}
    try:
        with open(_SESSIONS_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}

def _save_sessions(sessions: dict):
    with open(_SESSIONS_FILE, "w", encoding="utf-8") as f:
        json.dump(sessions, f, indent=2)

def create_persistent_session(username: str) -> str:
    """Generate a server-side session token and persist it."""
    token = secrets.token_urlsafe(32)
    sessions = _load_sessions()
    sessions[token] = {
        "username": username,
        "created": datetime.now().isoformat(timespec="seconds"),
    }
    # Keep only the 50 most recent sessions to avoid unbounded growth
    if len(sessions) > 50:
        sorted_items = sorted(sessions.items(), key=lambda x: x[1].get("created", ""))
        sessions = dict(sorted_items[-50:])
    _save_sessions(sessions)
    return token

def validate_persistent_session(token: str):
    """Return user data dict if the token is valid, else None."""
    sessions = _load_sessions()
    entry = sessions.get(token)
    if not entry:
        return None
    username = entry.get("username")
    if not username:
        return None
    users = load_users()
    if username not in users:
        del sessions[token]
        _save_sessions(sessions)
        return None
    u = users[username]
    if not u.get("active", True) or _is_expired(u):
        return None
    if not u.get("email_verified", True):
        return None
    return {**u, "username": username}

def revoke_persistent_session(token: str):
    sessions = _load_sessions()
    if token in sessions:
        del sessions[token]
        _save_sessions(sessions)

# ─────────────────────────────────────────────────────────────────
# AUTHENTICATION
# ─────────────────────────────────────────────────────────────────
_SESSION_KEY      = "_saas_auth_user"
_PERSIST_TOKEN_KEY = "_saas_persist_token"

def authenticate(username: str, password: str) -> tuple:
    """Returns (success, user_dict | error_message)"""
    users = load_users()
    if username not in users:
        return False, "Invalid username or password"
    u = users[username]
    if not u.get("active", True):
        return False, "Account is deactivated. Contact admin."
    if _is_expired(u):
        return False, "Account has expired. Contact admin."
    if not _verify_password(password, u.get("password", "")):
        return False, "Invalid username or password"
    if not u.get("email_verified", True):   # default True for old accounts
        return False, "__EMAIL_NOT_VERIFIED__"
    return True, {**u, "username": username}

def get_current_user() -> dict:
    return st.session_state.get(_SESSION_KEY, {})

def logout():
    sid = st.session_state.get(_PERSIST_TOKEN_KEY)
    if sid:
        revoke_persistent_session(sid)
    st.session_state.pop(_SESSION_KEY, None)
    st.session_state.pop("_saas_page", None)
    st.session_state.pop(_PERSIST_TOKEN_KEY, None)
    st.rerun()

def saas_auth_gate():
    """Show login page and halt if user is not authenticated."""
    params = st.query_params

    # ── Email verification link (?verify=TOKEN) ──────────────────
    if "verify" in params:
        token = params["verify"]
        ok, name_or_msg = verify_email_token(token)
        st.query_params.clear()
        if ok:
            st.session_state["_saas_verify_ok"] = (
                f"Email verified! Welcome, {name_or_msg}. You can now sign in."
            )
        else:
            st.session_state["_saas_verify_err"] = name_or_msg

    # ── Restore session from URL on refresh (?_sid=TOKEN) ────────
    # ?_sid is kept in the URL while the user is logged in.
    # On refresh Streamlit re-reads query params from the browser URL,
    # so Python can validate the token directly — no JavaScript needed.
    if "_sid" in params and not st.session_state.get(_SESSION_KEY):
        sid = params["_sid"]
        user_data = validate_persistent_session(sid)
        if user_data:
            st.session_state[_SESSION_KEY]       = user_data
            st.session_state[_PERSIST_TOKEN_KEY] = sid
            # keep ?_sid in URL so subsequent refreshes also work
        else:
            # Token invalid/expired — clear it and show login
            st.query_params.clear()

    # ── Already authenticated — keep ?_sid in URL and continue ───
    if st.session_state.get(_SESSION_KEY):
        sid = st.session_state.get(_PERSIST_TOKEN_KEY, "")
        if sid and params.get("_sid") != sid:
            st.query_params["_sid"] = sid   # restore to URL if missing
        return

    _show_login_page()
    st.stop()

# ─────────────────────────────────────────────────────────────────
# HISTORY — JSON files
# ─────────────────────────────────────────────────────────────────
def _load_json(path: Path) -> list:
    if not path.exists():
        return []
    try:
        with open(path, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return []

def _save_json(path: Path, data: list):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)

def log_login(username: str, success: bool, reason: str = ""):
    history = _load_json(_LOGIN_HISTORY)
    history.append({
        "username"  : username,
        "timestamp" : datetime.now().isoformat(timespec="seconds"),
        "success"   : success,
        "reason"    : reason,
    })
    _save_json(_LOGIN_HISTORY, history[-5000:])  # keep last 5000 entries

def log_download(username: str, filename: str):
    history = _load_json(_DOWNLOAD_HISTORY)
    history.append({
        "username"  : username,
        "filename"  : filename,
        "timestamp" : datetime.now().isoformat(timespec="seconds"),
    })
    _save_json(_DOWNLOAD_HISTORY, history[-5000:])

def get_login_history() -> list:
    return _load_json(_LOGIN_HISTORY)

def get_download_history(username: str = None) -> list:
    h = _load_json(_DOWNLOAD_HISTORY)
    return [x for x in h if x["username"] == username] if username else h

# ─────────────────────────────────────────────────────────────────
# ANNOUNCEMENTS
# ─────────────────────────────────────────────────────────────────
def add_announcement(title: str, message: str, author: str):
    anns = _load_json(_ANNOUNCEMENTS)
    anns.append({
        "id"      : (max((a["id"] for a in anns), default=0) + 1),
        "title"   : title,
        "message" : message,
        "author"  : author,
        "created" : datetime.now().isoformat(timespec="seconds"),
    })
    _save_json(_ANNOUNCEMENTS, anns)

def get_announcements() -> list:
    return _load_json(_ANNOUNCEMENTS)

def delete_announcement(ann_id: int):
    anns = [a for a in _load_json(_ANNOUNCEMENTS) if a.get("id") != ann_id]
    _save_json(_ANNOUNCEMENTS, anns)

# ─────────────────────────────────────────────────────────────────
# STREAMLIT — LOGIN PAGE
# ─────────────────────────────────────────────────────────────────
_LOGO_WHITE = "https://i.ibb.co/b5JwJC5Y/ASC-New-Logo-White.png"
_LOGO_BLUE  = "https://i.ibb.co/LXmWddkt/ASC-New-Logo-Blue.png"

_LOGIN_CSS = """
<style>
* { font-family: Arial, 'Arial Narrow', Helvetica, sans-serif !important; }

/* display:none removes layout space; visibility:hidden does not */
header[data-testid="stHeader"],
div[data-testid="stToolbar"],
div[data-testid="stDecoration"],
div[data-testid="stStatusWidget"],
section[data-testid="stSidebar"],
#MainMenu, footer { display: none !important; }

/* Zero out Streamlit's --header-height CSS variable that drives padding-top */
:root { --header-height: 0px !important; }

/* No scrollbar anywhere */
html, body { overflow: hidden !important; height: 100vh !important; margin: 0 !important; padding: 0 !important; }
.stApp { overflow: hidden !important; height: 100vh !important; margin: 0 !important; padding: 0 !important; }

/* Remove ALL padding/margin from every Streamlit container level */
.stApp,
section[data-testid="stAppViewContainer"],
section[data-testid="stAppViewContainer"] > div,
div[data-testid="stAppViewBlockContainer"],
div[data-testid="stMainBlockContainer"],
section.main,
section.main > div,
div.block-container {
    padding: 0 !important;
    margin: 0 !important;
    max-width: 100% !important;
}

/* Zero the top-level stVerticalBlock (Streamlit's main content wrapper) gap */
div.block-container > div[data-testid="stVerticalBlock"],
div[data-testid="stAppViewBlockContainer"] > div[data-testid="stVerticalBlock"],
div[data-testid="stMainBlockContainer"] > div[data-testid="stVerticalBlock"] {
    padding: 0 !important;
    margin: 0 !important;
    gap: 0 !important;
}

/* Two-panel split fills full viewport */
div[data-testid="stHorizontalBlock"] {
    gap: 0 !important;
    align-items: stretch !important;
    height: 100vh !important;
    overflow: hidden !important;
    margin: 0 !important;
}

/* Strip padding/margin/gap from every wrapper level inside both columns */
div[data-testid="stHorizontalBlock"] div[data-testid="stColumn"],
div[data-testid="stHorizontalBlock"] div[data-testid="stColumn"] > div,
div[data-testid="stHorizontalBlock"] div[data-testid="stColumn"] > div > div,
div[data-testid="stHorizontalBlock"] div[data-testid="stVerticalBlockBorderWrapper"],
div[data-testid="stHorizontalBlock"] div[data-testid="stVerticalBlock"] {
    padding: 0 !important;
    margin: 0 !important;
    gap: 0 !important;
}

/* LEFT column — blue background on column and all its structural wrappers */
div[data-testid="stHorizontalBlock"] > div[data-testid="stColumn"]:first-child,
div[data-testid="stHorizontalBlock"] > div[data-testid="stColumn"]:first-child > div,
div[data-testid="stHorizontalBlock"] > div[data-testid="stColumn"]:first-child > div > div,
div[data-testid="stHorizontalBlock"] > div[data-testid="stColumn"]:first-child div[data-testid="stVerticalBlockBorderWrapper"],
div[data-testid="stHorizontalBlock"] > div[data-testid="stColumn"]:first-child div[data-testid="stVerticalBlock"] {
    background: #1f2f60 !important;
    height: 100vh !important;
    overflow: hidden !important;
}

/* RIGHT column — background covers all structural levels */
div[data-testid="stHorizontalBlock"] > div[data-testid="stColumn"]:last-child,
div[data-testid="stHorizontalBlock"] > div[data-testid="stColumn"]:last-child > div,
div[data-testid="stHorizontalBlock"] > div[data-testid="stColumn"]:last-child > div > div,
div[data-testid="stHorizontalBlock"] > div[data-testid="stColumn"]:last-child div[data-testid="stVerticalBlockBorderWrapper"],
div[data-testid="stHorizontalBlock"] > div[data-testid="stColumn"]:last-child div[data-testid="stVerticalBlock"] {
    background: #f4f6fb !important;
}

/* height/overflow only on outer wrappers — NOT on stVerticalBlock, to avoid clipping content */
div[data-testid="stHorizontalBlock"] > div[data-testid="stColumn"]:last-child,
div[data-testid="stHorizontalBlock"] > div[data-testid="stColumn"]:last-child > div,
div[data-testid="stHorizontalBlock"] > div[data-testid="stColumn"]:last-child > div > div,
div[data-testid="stHorizontalBlock"] > div[data-testid="stColumn"]:last-child div[data-testid="stVerticalBlockBorderWrapper"] {
    height: 100vh !important;
    overflow: hidden !important;
}

/* Form inputs */
div[data-testid="stColumn"]:last-child input[type="text"],
div[data-testid="stColumn"]:last-child input[type="password"] {
    border: 1.5px solid #d8e2ef !important; border-radius: 8px !important;
    background: #ffffff !important; font-size: 14px !important; color: #1f2f60 !important;
}
div[data-testid="stColumn"]:last-child input:focus {
    border-color: #1f2f60 !important;
    box-shadow: 0 0 0 3px rgba(31,47,96,.08) !important;
}

/* Sign In button */
div[data-testid="stColumn"]:last-child div[data-testid="stFormSubmitButton"] button {
    background: #1f2f60 !important; color: #ffffff !important;
    border: none !important; border-radius: 8px !important;
    height: 52px !important; font-size: 14.5px !important;
    font-weight: 700 !important; width: 100% !important;
}
div[data-testid="stColumn"]:last-child div[data-testid="stFormSubmitButton"] button:hover {
    background: #2a3d72 !important;
}

/* Labels */
div[data-testid="stColumn"]:last-child label {
    color: #2d4a6a !important; font-weight: 700 !important; font-size: 13px !important;
}

/* Hide "Press Enter to submit form" hint on all form inputs */
div[data-testid="InputInstructions"] { display: none !important; }

/* Add vertical gap between form children (inputs, button etc.) */
div[data-testid="stColumn"]:last-child div[data-testid="stVerticalBlock"] {
    gap: 6px !important;
}

/* Create New Account button */
div[data-testid="stColumn"]:last-child div[data-testid="stBaseButton-secondary"] button {
    background: #ffffff !important; color: #1f2f60 !important;
    border: 1.5px solid #d8e2ef !important; border-radius: 8px !important;
    height: 52px !important; font-size: 14px !important;
    font-weight: 600 !important; width: 100% !important;
    margin-top: 4px !important;
}
div[data-testid="stColumn"]:last-child div[data-testid="stBaseButton-secondary"] button:hover {
    background: #f0f2f8 !important; border-color: #1f2f60 !important;
    color: #1f2f60 !important;
}

/* Registration form primary button (matches Sign In style) */
div[data-testid="stColumn"]:last-child div[data-testid="stBaseButton-primary"] button {
    background: #1f2f60 !important; color: #ffffff !important;
    border: none !important; border-radius: 8px !important;
    height: 52px !important; font-size: 14.5px !important;
    font-weight: 700 !important; width: 100% !important;
}
div[data-testid="stColumn"]:last-child div[data-testid="stBaseButton-primary"] button:hover {
    background: #2a3d72 !important;
}
</style>
"""

def _show_login_page():
    st.markdown(_LOGIN_CSS, unsafe_allow_html=True)

    col_left, col_right = st.columns([44, 56])

    # ── LEFT: Product pitch ────────────────────────────────────────
    with col_left:
        st.markdown(f"""
<div style="font-family:Arial,sans-serif;padding:32px 44px 28px;display:flex;flex-direction:column;height:100vh;box-sizing:border-box;position:relative;overflow:hidden;">
<div style="position:absolute;right:-120px;top:-120px;width:400px;height:400px;
     border-radius:50%;border:60px solid rgba(255,255,255,.03);pointer-events:none;"></div>
<div style="position:absolute;right:-80px;bottom:50px;width:300px;height:300px;
     border-radius:50%;border:44px solid rgba(255,255,255,.025);pointer-events:none;"></div>

<!-- Logo -->
<div style="margin-bottom:32px;position:relative;z-index:1;">
    <img src="{_LOGO_WHITE}" alt="ASC Consulting" style="height:52px;width:auto;display:block;"
         onerror="this.style.display='none';document.getElementById('asc-fallback-left').style.display='block'">
    <span id="asc-fallback-left" style="display:none;font-size:26px;font-weight:700;color:#fff;
          letter-spacing:3px;border-bottom:3px solid #f9be3e;padding-bottom:3px;">ASC</span>
</div>

<!-- Badge -->
<div style="display:inline-flex;align-items:center;gap:8px;background:rgba(255,255,255,.08);
     border:1px solid rgba(255,255,255,.14);border-radius:100px;padding:7px 16px 7px 12px;
     width:fit-content;margin-bottom:30px;position:relative;z-index:1;">
    <span style="width:7px;height:7px;border-radius:50%;background:#f9be3e;flex-shrink:0;display:inline-block;"></span>
    <span style="font-size:10.5px;font-weight:700;letter-spacing:1.6px;color:rgba(255,255,255,.78);text-transform:uppercase;font-family:Arial,sans-serif;">GSTR-1 FILING PLATFORM</span>
</div>

<h1 style="font-size:36px;font-weight:700;line-height:1.2;color:#fff;margin-bottom:16px;font-family:Arial,sans-serif;position:relative;z-index:1;">
    Automate your<br>
    <span style="color:#f9be3e;">GSTR-1 filing</span><br>
    with precision
</h1>

<p style="font-size:14px;line-height:1.7;color:rgba(255,255,255,.5);max-width:370px;margin-bottom:34px;font-family:Arial,sans-serif;position:relative;z-index:1;">
    Process multi-state sales data, generate B2B&nbsp;/&nbsp;B2CS&nbsp;/&nbsp;HSN sections,
    and export complete GSTR-1 returns — all in one powerful tool.
</p>

<!-- Stats -->
<div style="display:grid;grid-template-columns:repeat(3,1fr);gap:10px;margin-bottom:36px;position:relative;z-index:1;">
    <div style="background:rgba(255,255,255,.07);border:1px solid rgba(255,255,255,.1);border-radius:11px;padding:16px 10px 14px;text-align:center;">
        <div style="font-size:25px;font-weight:700;color:#f9be3e;font-variant-numeric:tabular-nums;margin-bottom:4px;font-family:Arial,sans-serif;">28</div>
        <div style="font-size:10px;font-weight:700;letter-spacing:1.3px;color:rgba(255,255,255,.38);text-transform:uppercase;font-family:Arial,sans-serif;">States</div>
    </div>
    <div style="background:rgba(255,255,255,.07);border:1px solid rgba(255,255,255,.1);border-radius:11px;padding:16px 10px 14px;text-align:center;">
        <div style="font-size:25px;font-weight:700;color:#f9be3e;font-variant-numeric:tabular-nums;margin-bottom:4px;font-family:Arial,sans-serif;">13+</div>
        <div style="font-size:10px;font-weight:700;letter-spacing:1.3px;color:rgba(255,255,255,.38);text-transform:uppercase;font-family:Arial,sans-serif;">Sections</div>
    </div>
    <div style="background:rgba(255,255,255,.07);border:1px solid rgba(255,255,255,.1);border-radius:11px;padding:16px 10px 14px;text-align:center;">
        <div style="font-size:25px;font-weight:700;color:#f9be3e;font-variant-numeric:tabular-nums;margin-bottom:4px;font-family:Arial,sans-serif;">24/7</div>
        <div style="font-size:10px;font-weight:700;letter-spacing:1.3px;color:rgba(255,255,255,.38);text-transform:uppercase;font-family:Arial,sans-serif;">Available</div>
    </div>
</div>

<!-- Features -->
<div style="display:flex;flex-direction:column;gap:16px;flex:1;position:relative;z-index:1;">
    <div style="display:flex;align-items:flex-start;gap:13px;">
        <div style="width:38px;height:38px;border-radius:9px;background:rgba(249,190,62,.13);border:1px solid rgba(249,190,62,.30);display:flex;align-items:center;justify-content:center;font-size:16px;flex-shrink:0;">📊</div>
        <div><div style="font-size:13px;font-weight:700;color:rgba(255,255,255,.9);margin-bottom:3px;font-family:Arial,sans-serif;">B2B / B2CL / B2CS processing</div>
        <div style="font-size:11.5px;color:rgba(255,255,255,.42);line-height:1.55;font-family:Arial,sans-serif;">Auto-classify invoices by GSTIN, state code and taxable value</div></div>
    </div>
    <div style="display:flex;align-items:flex-start;gap:13px;">
        <div style="width:38px;height:38px;border-radius:9px;background:rgba(249,190,62,.13);border:1px solid rgba(249,190,62,.30);display:flex;align-items:center;justify-content:center;font-size:16px;flex-shrink:0;">📋</div>
        <div><div style="font-size:13px;font-weight:700;color:rgba(255,255,255,.9);margin-bottom:3px;font-family:Arial,sans-serif;">HSN Summary &amp; Document Series</div>
        <div style="font-size:11.5px;color:rgba(255,255,255,.42);line-height:1.55;font-family:Arial,sans-serif;">HSN-wise tax breakup with invoice series gap detection</div></div>
    </div>
    <div style="display:flex;align-items:flex-start;gap:13px;">
        <div style="width:38px;height:38px;border-radius:9px;background:rgba(249,190,62,.13);border:1px solid rgba(249,190,62,.30);display:flex;align-items:center;justify-content:center;font-size:16px;flex-shrink:0;">↩️</div>
        <div><div style="font-size:13px;font-weight:700;color:rgba(255,255,255,.9);margin-bottom:3px;font-family:Arial,sans-serif;">Sales Return &amp; CDNR integration</div>
        <div style="font-size:11.5px;color:rgba(255,255,255,.42);line-height:1.55;font-family:Arial,sans-serif;">Filter taken returns, compute credit notes automatically</div></div>
    </div>
    <div style="display:flex;align-items:flex-start;gap:13px;">
        <div style="width:38px;height:38px;border-radius:9px;background:rgba(249,190,62,.13);border:1px solid rgba(249,190,62,.30);display:flex;align-items:center;justify-content:center;font-size:16px;flex-shrink:0;">📥</div>
        <div><div style="font-size:13px;font-weight:700;color:rgba(255,255,255,.9);margin-bottom:3px;font-family:Arial,sans-serif;">Excel export — 13+ sections</div>
        <div style="font-size:11.5px;color:rgba(255,255,255,.42);line-height:1.55;font-family:Arial,sans-serif;">One-click download with all GSTR-1 sheets formatted for filing</div></div>
    </div>
</div>

<p style="margin-top:32px;font-size:11.5px;color:rgba(255,255,255,.24);position:relative;z-index:1;font-family:Arial,sans-serif;">
    © 2026 ASC Consulting Pvt. Ltd. · All rights reserved
</p>
</div>
""", unsafe_allow_html=True)

    # ── RIGHT: Login / Register form ───────────────────────────────
    with col_right:
        _, form_mid, _ = st.columns([1, 8, 1])
        with form_mid:
            _show_reg = st.session_state.get("_saas_show_register", False)

            # ── Shared logo header ─────────────────────────────────
            st.markdown(f"""
<div style="text-align:center;margin-bottom:20px;padding-top:24px;">
    <img src="{_LOGO_BLUE}" alt="ASC Consulting" style="height:60px;width:auto;"
         onerror="this.style.display='none';document.getElementById('asc-fb-r2').style.display='inline-block'">
    <span id="asc-fb-r2" style="display:none;font-size:28px;font-weight:700;color:#1f2f60;
          letter-spacing:3px;border-bottom:3px solid #f9be3e;padding-bottom:3px;font-family:Arial,sans-serif;">ASC</span>
    <div style="font-size:10.5px;letter-spacing:2.5px;font-weight:700;color:#94ADC8;
                margin-top:10px;text-transform:uppercase;font-family:Arial,sans-serif;">GSTR-1 Dashboard</div>
</div>
""", unsafe_allow_html=True)

            if _show_reg:
                # ── REGISTRATION VIEW ──────────────────────────────
                st.markdown("""
<h2 style="font-size:26px;font-weight:700;color:#1f2f60;text-align:center;
           margin-bottom:6px;font-family:Arial,sans-serif;">Create Account</h2>
<p style="font-size:13.5px;color:#6B88A8;text-align:center;margin-bottom:24px;font-family:Arial,sans-serif;">
    Fill in the details below to get started
</p>
""", unsafe_allow_html=True)

                with st.form("_saas_register_form", clear_on_submit=False):
                    reg_name     = st.text_input("Full Name",        placeholder="Enter your full name")
                    reg_email    = st.text_input("Email Address",    placeholder="Enter your email address")
                    reg_username = st.text_input("Username",         placeholder="Choose a username")
                    reg_password = st.text_input("Password",         type="password",
                                                  placeholder="Create a password (min 6 chars)")
                    reg_confirm  = st.text_input("Confirm Password", type="password",
                                                  placeholder="Re-enter your password")
                    reg_submitted = st.form_submit_button(
                        "👤  Create Account", use_container_width=True, type="primary"
                    )

                if reg_submitted:
                    if not all([reg_name.strip(), reg_email.strip(),
                                reg_username.strip(), reg_password, reg_confirm]):
                        st.error("All fields are required.")
                    elif reg_password != reg_confirm:
                        st.error("Passwords do not match. Please try again.")
                    elif len(reg_password) < 6:
                        st.error("Password must be at least 6 characters long.")
                    else:
                        ok, msg = add_user(
                            reg_username.strip(), reg_name.strip(),
                            reg_email.strip(), reg_password, role="client"
                        )
                        if ok:
                            st.session_state["_saas_show_register"]       = False
                            st.session_state["_saas_pending_verify_email"] = reg_email.strip()
                            st.rerun()
                        else:
                            st.error(msg)

                st.markdown("""
<div style="display:flex;align-items:center;gap:14px;margin:24px 0 16px;font-family:Arial,sans-serif;">
    <div style="flex:1;height:1px;background:#d8e2ef;"></div>
    <span style="font-size:12px;color:#94adc8;letter-spacing:0.5px;">or</span>
    <div style="flex:1;height:1px;background:#d8e2ef;"></div>
</div>
""", unsafe_allow_html=True)

                if st.button("← Back to Sign In", use_container_width=True):
                    st.session_state["_saas_show_register"] = False
                    st.rerun()

            else:
                # ── LOGIN VIEW ─────────────────────────────────────
                st.markdown("""
<h2 style="font-size:26px;font-weight:700;color:#1f2f60;text-align:center;
           margin-bottom:6px;font-family:Arial,sans-serif;">Welcome back</h2>
<p style="font-size:13.5px;color:#6B88A8;text-align:center;margin-bottom:28px;font-family:Arial,sans-serif;">
    Sign in to your account to continue
</p>
""", unsafe_allow_html=True)

                # Verification success / error banners from URL param
                if st.session_state.get("_saas_verify_ok"):
                    st.success(st.session_state.pop("_saas_verify_ok"))
                if st.session_state.get("_saas_verify_err"):
                    st.error(st.session_state.pop("_saas_verify_err"))
                # Post-registration banner
                if st.session_state.get("_saas_pending_verify_email"):
                    _pve = st.session_state.pop("_saas_pending_verify_email")
                    _smtp_ok = _smtp_configured()
                    if _smtp_ok:
                        st.info(
                            f"📧 A verification link has been sent to **{_pve}**. "
                            "Please check your inbox and verify your email to access your account."
                        )
                    else:
                        st.success(
                            "Account created! (Email verification is not configured — "
                            "you can sign in directly.)"
                        )

                with st.form("_saas_login_form", clear_on_submit=False):
                    login_email = st.text_input("Email Address", placeholder="Enter your email address")
                    password    = st.text_input("Password", type="password",
                                                placeholder="Enter your password")
                    submitted = st.form_submit_button(
                        "→  Sign In", use_container_width=True, type="primary"
                    )

                if submitted:
                    if not login_email.strip() or not password:
                        st.error("Please enter your email address and password.")
                    else:
                        resolved = _find_username_by_email(login_email.strip())
                        if not resolved:
                            st.error("No account found with this email address.")
                        else:
                            ok, result = authenticate(resolved, password)
                            if ok:
                                st.session_state[_SESSION_KEY] = result
                                log_login(resolved, True)
                                sid = create_persistent_session(resolved)
                                st.session_state[_PERSIST_TOKEN_KEY] = sid
                                st.query_params["_sid"] = sid   # persist across refresh
                                st.rerun()
                            elif result == "__EMAIL_NOT_VERIFIED__":
                                st.warning(
                                    "📧 Your email address is not verified. "
                                    "Please check your inbox and click the verification link."
                                )
                            else:
                                log_login(resolved, False, result)
                                st.error(result)

                st.markdown("""
<div style="display:flex;align-items:center;gap:14px;margin:28px 0 20px;font-family:Arial,sans-serif;">
    <div style="flex:1;height:1px;background:#d8e2ef;"></div>
    <span style="font-size:12px;color:#94adc8;letter-spacing:0.5px;">or</span>
    <div style="flex:1;height:1px;background:#d8e2ef;"></div>
</div>
""", unsafe_allow_html=True)

                if st.button("👤+  Create New Account", use_container_width=True):
                    st.session_state["_saas_show_register"] = True
                    st.rerun()

            st.markdown("""
<p style="text-align:center;font-size:11.5px;color:#94ADC8;margin-top:32px;padding-bottom:24px;font-family:Arial,sans-serif;">
    © 2026 ASC Consulting Pvt. Ltd. · All rights reserved
</p>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────
# STREAMLIT — SIDEBAR USER INFO + NAVIGATION
# ─────────────────────────────────────────────────────────────────
def saas_sidebar_nav() -> str:
    """
    Render user info + navigation in the sidebar.
    Returns the selected page: 'dashboard', 'admin', or 'profile'.
    """
    user = get_current_user()
    role = user.get("role", "client")

    # ASC logo at the top of the sidebar
    st.sidebar.markdown(f"""
<div style="text-align:center;padding:20px 0 16px;
            border-bottom:1px solid rgba(255,255,255,0.12);margin-bottom:12px;">
    <img src="{_LOGO_WHITE}" alt="ASC Consulting"
         style="height:100px;width:100px;object-fit:contain;display:inline-block;"
         onerror="this.style.display='none';document.getElementById('sb-asc-fb').style.display='inline-block'">
    <span id="sb-asc-fb" style="display:none;font-size:22px;font-weight:700;color:#fff;
          letter-spacing:3px;border-bottom:3px solid #f9be3e;padding-bottom:2px;">ASC</span>
</div>
""", unsafe_allow_html=True)

    st.sidebar.markdown(
        f"<div style='padding:4px 0 4px'>"
        f"<b style='color:#fff;'>{'🔑' if role=='admin' else '👤'} {user.get('name','User')}</b><br>"
        f"<span style='font-size:12px;color:rgba(255,255,255,0.55)'>{role.capitalize()}</span>"
        f"</div>",
        unsafe_allow_html=True,
    )

    nav_options = ["📊 Dashboard", "👤 Profile"]
    if role == "admin":
        nav_options.insert(1, "⚙️ Admin Panel")

    sel = st.sidebar.radio(
        "Navigate",
        nav_options,
        key="_saas_nav",
        label_visibility="collapsed",
    )

    if st.sidebar.button("🚪 Logout", use_container_width=True, key="_saas_logout_btn"):
        log_login(user.get("username", ""), True, "logged out")
        logout()

    st.sidebar.divider()

    if "Admin Panel" in sel:
        return "admin"
    if "Profile" in sel:
        return "profile"
    return "dashboard"

# ─────────────────────────────────────────────────────────────────
# STREAMLIT — ANNOUNCEMENTS BANNER
# ─────────────────────────────────────────────────────────────────
def show_announcements_banner():
    anns = get_announcements()
    if anns:
        latest = anns[-1]
        st.info(f"📢 **{latest['title']}** — {latest['message']}")

# ─────────────────────────────────────────────────────────────────
# STREAMLIT — ADMIN PANEL
# ─────────────────────────────────────────────────────────────────
def show_admin_panel(user: dict):
    import pandas as pd

    st.markdown("""
    <div style='background:linear-gradient(90deg,#1e3a5f,#2d6a9f);
    padding:18px 28px;border-radius:10px;color:white;margin-bottom:20px'>
    <h2 style='margin:0'>⚙️ Admin Panel</h2>
    <p style='margin:4px 0 0;opacity:.85;font-size:13px'>
    User Management · Login History · Announcements · Downloads</p>
    </div>
    """, unsafe_allow_html=True)

    tab_users, tab_login, tab_ann, tab_dl = st.tabs(
        ["👥 Users", "📋 Login History", "📢 Announcements", "📥 Downloads"]
    )

    # ── Tab 1: Users ─────────────────────────────────────────────
    with tab_users:
        st.markdown("### All Users")
        users = load_users()

        if users:
            # Table header
            h = st.columns([2, 2, 3, 1.2, 1, 1.5, 1.5, 0.8])
            for col, label in zip(h, ["Username","Name","Email","Role","Active","Expiry","Created",""]):
                col.markdown(f"<span style='font-size:12px;color:#6b88a8;font-weight:700;'>{label}</span>",
                             unsafe_allow_html=True)
            st.markdown("<hr style='margin:4px 0 6px;border-color:#e2e8f0;'>", unsafe_allow_html=True)

            cur_username = user.get("username", "")
            for uname, u in users.items():
                c = st.columns([2, 2, 3, 1.2, 1, 1.5, 1.5, 0.8])
                c[0].markdown(f"`{uname}`")
                c[1].write(u.get("name", ""))
                c[2].write(u.get("email", ""))
                c[3].write(u.get("role", "client").capitalize())
                c[4].write("✅" if u.get("active", True) else "❌")
                c[5].write(u.get("expiry") or "Never")
                c[6].write(u.get("created", ""))
                if uname == cur_username:
                    c[7].write("—")
                else:
                    if c[7].button("🗑️", key=f"_del_{uname}", help=f"Delete {uname}"):
                        st.session_state[f"_confirm_del_{uname}"] = True

                # Confirmation row
                if st.session_state.get(f"_confirm_del_{uname}"):
                    with st.container():
                        st.warning(f"Delete **{uname}** ({u.get('name','')})? This cannot be undone.")
                        yes_col, no_col, _ = st.columns([1, 1, 5])
                        if yes_col.button("Yes, delete", key=f"_yes_{uname}", type="primary"):
                            remove_user(uname)
                            st.session_state.pop(f"_confirm_del_{uname}", None)
                            st.success(f"User **{uname}** deleted.")
                            st.rerun()
                        if no_col.button("Cancel", key=f"_no_{uname}"):
                            st.session_state.pop(f"_confirm_del_{uname}", None)
                            st.rerun()
        else:
            st.info("No users found.")

        st.markdown("---")
        col_add, col_manage = st.columns(2)

        with col_add:
            st.markdown("#### ➕ Add New User")
            with st.form("_saas_add_user"):
                nu_username = st.text_input("Username *")
                nu_name     = st.text_input("Full Name *")
                nu_email    = st.text_input("Email")
                nu_password = st.text_input("Password *", type="password")
                nu_role     = st.selectbox("Role", ["client", "admin"])
                nu_expiry   = st.date_input("Expiry Date (leave blank = never)", value=None)
                nu_submit   = st.form_submit_button("➕ Add User", type="primary")

            if nu_submit:
                if not nu_username.strip() or not nu_name.strip() or not nu_password:
                    st.error("Username, Name and Password are required.")
                else:
                    ok, msg = add_user(
                        nu_username.strip(), nu_name.strip(), nu_email.strip(),
                        nu_password, nu_role,
                        nu_expiry if nu_expiry else None,
                    )
                    if ok:
                        st.success(f"✅ {msg}")
                        st.rerun()
                    else:
                        st.error(msg)

        with col_manage:
            st.markdown("#### ✏️ Manage User")
            users_now = load_users()
            if users_now:
                sel_u = st.selectbox("Select User", list(users_now.keys()), key="_saas_sel_user")
                u_data = users_now[sel_u]
                is_active = u_data.get("active", True)

                ca, cb = st.columns(2)
                with ca:
                    if st.button(
                        "❌ Deactivate" if is_active else "✅ Activate",
                        use_container_width=True, key="_saas_toggle_active"
                    ):
                        toggle_active(sel_u, not is_active)
                        st.success("Updated.")
                        st.rerun()
                with cb:
                    if sel_u != user.get("username"):  # can't remove self
                        if st.button("🗑️ Remove User", use_container_width=True,
                                     key="_saas_remove_user", type="secondary"):
                            ok, msg = remove_user(sel_u)
                            if ok:
                                st.success(msg)
                                st.rerun()
                    else:
                        st.caption("(Can't remove yourself)")

                new_exp = st.date_input("Update Expiry Date", value=None, key="_saas_upd_exp")
                if st.button("📅 Update Expiry", use_container_width=True, key="_saas_btn_exp"):
                    update_expiry(sel_u, new_exp if new_exp else None)
                    st.success("Expiry updated.")
                    st.rerun()

                new_pwd = st.text_input("New Password", type="password", key="_saas_upd_pwd")
                if st.button("🔑 Reset Password", use_container_width=True, key="_saas_btn_pwd"):
                    if new_pwd:
                        update_password(sel_u, new_pwd)
                        st.success("Password reset successfully.")
                    else:
                        st.warning("Enter a new password first.")

                # ── Change role ───────────────────────────────────
                if sel_u != user.get("username"):
                    cur_role = u_data.get("role", "client")
                    new_role = st.selectbox("Role", ["client", "admin"],
                                            index=(1 if cur_role == "admin" else 0),
                                            key="_saas_upd_role")
                    if st.button("🔄 Change Role", use_container_width=True, key="_saas_btn_role"):
                        ok, msg = update_role(sel_u, new_role)
                        if ok:
                            st.success(msg)
                            st.rerun()
                        else:
                            st.error(msg)
                else:
                    st.caption("(Can't change your own role)")

                # ── Email verification status / actions ───────────
                if not u_data.get("email_verified", True):
                    st.warning("📧 Email not verified — user cannot log in yet.")
                    cv, cr = st.columns(2)
                    with cv:
                        if st.button("✅ Mark Verified", use_container_width=True, key="_saas_btn_verify"):
                            ok, msg = set_email_verified(sel_u)
                            if ok:
                                st.success(msg)
                                st.rerun()
                            else:
                                st.error(msg)
                    with cr:
                        if st.button("📧 Resend Email", use_container_width=True, key="_saas_btn_resend"):
                            ok, msg = resend_verification(sel_u)
                            if ok:
                                st.success(msg)
                            else:
                                st.error(msg)
                else:
                    st.caption("📧 Email verified ✓")
            else:
                st.info("No users to manage.")

    # ── Tab 2: Login History ──────────────────────────────────────
    with tab_login:
        st.markdown("### Login History")
        history = get_login_history()[::-1]  # newest first
        if history:
            df_h = pd.DataFrame(history)
            df_h["timestamp"] = pd.to_datetime(df_h["timestamp"]).dt.strftime("%Y-%m-%d %H:%M")
            df_h["status"]    = df_h["success"].map({True: "✅ Success", False: "❌ Failed"})
            df_h = df_h[["timestamp", "username", "status", "reason"]].rename(
                columns={"timestamp": "Time", "username": "User",
                         "status": "Status", "reason": "Reason"}
            )
            st.dataframe(df_h, use_container_width=True, hide_index=True, height=400)
            st.caption(f"Total records: {len(history)}")
        else:
            st.info("No login history yet.")

    # ── Tab 3: Announcements ──────────────────────────────────────
    with tab_ann:
        st.markdown("### Post Announcement")
        with st.form("_saas_ann_form"):
            ann_title = st.text_input("Title")
            ann_msg   = st.text_area("Message", height=100)
            ann_sub   = st.form_submit_button("📢 Post", type="primary")

        if ann_sub:
            if ann_title.strip() and ann_msg.strip():
                add_announcement(ann_title.strip(), ann_msg.strip(), user.get("name", "Admin"))
                st.success("Announcement posted.")
                st.rerun()
            else:
                st.error("Title and message are required.")

        st.markdown("---")
        st.markdown("### Active Announcements")
        anns = get_announcements()[::-1]
        if anns:
            for a in anns:
                with st.expander(f"📢 {a['title']}  ·  {a['created'][:10]}"):
                    st.write(a["message"])
                    st.caption(f"Posted by: {a['author']}")
                    if st.button("🗑️ Delete", key=f"_saas_del_ann_{a['id']}"):
                        delete_announcement(a["id"])
                        st.rerun()
        else:
            st.info("No announcements yet.")

    # ── Tab 4: Downloads ──────────────────────────────────────────
    with tab_dl:
        st.markdown("### Download History (All Users)")
        dl = get_download_history()[::-1]
        if dl:
            df_d = pd.DataFrame(dl)
            df_d["timestamp"] = pd.to_datetime(df_d["timestamp"]).dt.strftime("%Y-%m-%d %H:%M")
            df_d = df_d.rename(columns={
                "timestamp": "Time", "username": "User", "filename": "File"
            })
            st.dataframe(df_d, use_container_width=True, hide_index=True, height=400)
            st.caption(f"Total downloads: {len(dl)}")
        else:
            st.info("No downloads recorded yet.")

# ─────────────────────────────────────────────────────────────────
# STREAMLIT — PROFILE PAGE
# ─────────────────────────────────────────────────────────────────
def show_profile_page(user: dict):
    import pandas as pd

    st.markdown("""
    <div style='background:linear-gradient(90deg,#1e3a5f,#2d6a9f);
    padding:18px 28px;border-radius:10px;color:white;margin-bottom:20px'>
    <h2 style='margin:0'>👤 My Profile</h2>
    <p style='margin:4px 0 0;opacity:.85;font-size:13px'>
    Account details · Change password · Download history</p>
    </div>
    """, unsafe_allow_html=True)

    col1, col2 = st.columns([1, 1])

    with col1:
        st.markdown("### Account Details")
        st.markdown(f"""
| Field | Value |
|---|---|
| **Name** | {user.get('name', '')} |
| **Username** | {user.get('username', '')} |
| **Email** | {user.get('email', '') or '—'} |
| **Role** | {user.get('role', 'client').capitalize()} |
| **Account Expiry** | {user.get('expiry') or 'Never'} |
| **Account Status** | {'✅ Active' if user.get('active', True) else '❌ Inactive'} |
""")

    with col2:
        st.markdown("### 🔑 Change Password")
        with st.form("_saas_chg_pwd"):
            old_pwd  = st.text_input("Current Password", type="password")
            new_pwd1 = st.text_input("New Password", type="password")
            new_pwd2 = st.text_input("Confirm New Password", type="password")
            pwd_sub  = st.form_submit_button("Update Password", type="primary")

        if pwd_sub:
            if not old_pwd or not new_pwd1 or not new_pwd2:
                st.error("All fields are required.")
            elif new_pwd1 != new_pwd2:
                st.error("New passwords do not match.")
            elif len(new_pwd1) < 6:
                st.error("Password must be at least 6 characters.")
            else:
                users = load_users()
                u_stored = users.get(user["username"], {})
                if _verify_password(old_pwd, u_stored.get("password", "")):
                    update_password(user["username"], new_pwd1)
                    # Update session user data too
                    st.session_state[_SESSION_KEY]["password"] = _hash_password(new_pwd1)
                    st.success("✅ Password updated successfully.")
                else:
                    st.error("Current password is incorrect.")

    st.markdown("---")
    st.markdown("### 📥 My Download History")
    dl = get_download_history(user.get("username", ""))[::-1]
    if dl:
        df_d = pd.DataFrame(dl)
        df_d["timestamp"] = pd.to_datetime(df_d["timestamp"]).dt.strftime("%Y-%m-%d %H:%M")
        df_d = df_d[["timestamp", "filename"]].rename(
            columns={"timestamp": "Time", "filename": "File Downloaded"}
        )
        st.dataframe(df_d, use_container_width=True, hide_index=True, height=300)
        st.caption(f"Total: {len(dl)} download(s)")
    else:
        st.info("No downloads yet.")
