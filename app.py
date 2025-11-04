# ======================================== 
# Gmail Mail Merge Tool - Modern UI Edition (Duplicate-Safe Edition)
# ========================================
import streamlit as st
import pandas as pd
import base64
import time
import re
import json
import random
import os
from datetime import datetime, timedelta
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from google_auth_oauthlib.flow import Flow
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

# ========================================
# Streamlit Page Setup
# ========================================
st.set_page_config(page_title="Gmail Mail Merge", layout="wide")

# Sidebar
with st.sidebar:
    st.image("logo.png", width=180)
    st.markdown("---")
    st.markdown("### 📧 Gmail Mail Merge Tool")
    st.markdown("A Gmail-based mail merge app with resume and follow-up protection.")
    st.markdown("---")
    st.caption("Developed by Ranjith")

# Main Header
st.markdown("<h1 style='text-align:center;'>📧 Gmail Mail Merge Tool</h1>", unsafe_allow_html=True)
st.markdown("<p style='text-align:center;color:gray;'>Duplicate-Safe Edition with Resume Support</p>", unsafe_allow_html=True)
st.markdown("---")

# ========================================
# Gmail API Setup
# ========================================
SCOPES = [
    "https://www.googleapis.com/auth/gmail.send",
    "https://www.googleapis.com/auth/gmail.modify",
    "https://www.googleapis.com/auth/gmail.labels",
    "https://www.googleapis.com/auth/gmail.compose",
]

CLIENT_CONFIG = {
    "web": {
        "client_id": st.secrets["gmail"]["client_id"],
        "client_secret": st.secrets["gmail"]["client_secret"],
        "auth_uri": "https://accounts.google.com/o/oauth2/auth",
        "token_uri": "https://oauth2.googleapis.com/token",
        "redirect_uris": [st.secrets["gmail"]["redirect_uri"]],
    }
}

# ========================================
# Constants & Files
# ========================================
DONE_FILE = "/tmp/mailmerge_done.json"
LOCK_FILE = "/tmp/mailmerge_lock.json"
WORKING_CSV = "/tmp/mailmerge_working.csv"
BATCH_SIZE_DEFAULT = 50
DRAFT_BATCH_SIZE_DEFAULT = 110

# ========================================
# Safe Lock Helpers (atomic)
# ========================================
def _read_lock():
    try:
        with open(LOCK_FILE, "r") as f:
            return json.load(f)
    except Exception:
        return None

def is_lock_active():
    info = _read_lock()
    if not info:
        return False
    try:
        ts = datetime.fromisoformat(info.get("start_time"))
        if datetime.now() - ts > timedelta(hours=24):
            os.remove(LOCK_FILE)
            return False
        return True
    except Exception:
        try:
            os.remove(LOCK_FILE)
        except Exception:
            pass
        return False

def create_lock(info: dict):
    info["start_time"] = datetime.now().isoformat()
    try:
        fd = os.open(LOCK_FILE, os.O_WRONLY | os.O_CREAT | os.O_EXCL)
        with os.fdopen(fd, "w") as f:
            json.dump(info, f)
        return True
    except FileExistsError:
        return False
    except Exception:
        try:
            with open(LOCK_FILE, "w") as f:
                json.dump(info, f)
            return True
        except Exception:
            return False

def remove_lock():
    try:
        if os.path.exists(LOCK_FILE):
            os.remove(LOCK_FILE)
    except Exception:
        pass

# ========================================
# Progress Helpers
# ========================================
def save_progress_df(df):
    try:
        df.to_csv(WORKING_CSV, index=False)
    except Exception:
        pass

def load_progress_df_if_exists():
    if os.path.exists(WORKING_CSV):
        try:
            return pd.read_csv(WORKING_CSV, dtype=str).fillna("")
        except Exception:
            return None
    return None

# ========================================
# Helpers
# ========================================
EMAIL_REGEX = re.compile(r"[\w\.-]+@[\w\.-]+\.\w+")

def extract_email(value: str):
    if not value:
        return None
    match = EMAIL_REGEX.search(str(value))
    return match.group(0) if match else None

def convert_bold(text):
    if not text:
        return ""
    text = re.sub(r"\*\*(.*?)\*\*", r"<b>\1</b>", text)
    text = re.sub(
        r"\[(.*?)\]\((https?://[^\s)]+)\)",
        r'<a href="\2" style="color:#1a73e8;text-decoration:underline;" target="_blank">\1</a>',
        text,
    )
    text = text.replace("\n", "<br>").replace("  ", "&nbsp;&nbsp;")
    return f"<html><body style='font-family:Arial;font-size:14px;line-height:1.6'>{text}</body></html>"

def get_or_create_label(service, label_name="Mail Merge Sent"):
    try:
        labels = service.users().labels().list(userId="me").execute().get("labels", [])
        for label in labels:
            if label["name"].lower() == label_name.lower():
                return label["id"]
        created = service.users().labels().create(
            userId="me",
            body={"name": label_name, "labelListVisibility": "labelShow", "messageListVisibility": "show"},
        ).execute()
        return created["id"]
    except Exception:
        return None

def send_email_backup(service, csv_path):
    try:
        user_email = service.users().getProfile(userId="me").execute()["emailAddress"]
        msg = MIMEMultipart()
        msg["To"] = user_email
        msg["From"] = user_email
        msg["Subject"] = f"📁 Mail Merge Backup CSV - {datetime.now().strftime('%Y-%m-%d %H:%M')}"
        msg.attach(MIMEText("Attached is the backup CSV for your mail merge run.", "plain"))
        with open(csv_path, "rb") as f:
            part = MIMEApplication(f.read(), Name=os.path.basename(csv_path))
        part["Content-Disposition"] = f'attachment; filename="{os.path.basename(csv_path)}"'
        msg.attach(part)
        raw = base64.urlsafe_b64encode(msg.as_bytes()).decode()
        service.users().messages().send(userId="me", body={"raw": raw}).execute()
        st.info(f"📧 Backup CSV emailed to {user_email}")
    except Exception:
        pass

def fetch_message_id_header(service, message_id):
    for _ in range(6):
        try:
            msg_detail = service.users().messages().get(
                userId="me", id=message_id, format="metadata", metadataHeaders=["Message-ID"]
            ).execute()
            for h in msg_detail.get("payload", {}).get("headers", []):
                if h.get("name", "").lower() == "message-id":
                    return h.get("value")
        except Exception:
            pass
        time.sleep(random.uniform(1, 2))
    return ""

# ========================================
# OAuth Flow
# ========================================
if "creds" not in st.session_state:
    st.session_state["creds"] = None

if st.session_state["creds"]:
    creds = Credentials.from_authorized_user_info(json.loads(st.session_state["creds"]), SCOPES)
else:
    code = st.experimental_get_query_params().get("code", None)
    if code:
        flow = Flow.from_client_config(CLIENT_CONFIG, scopes=SCOPES)
        flow.redirect_uri = st.secrets["gmail"]["redirect_uri"]
        flow.fetch_token(code=code[0])
        creds = flow.credentials
        st.session_state["creds"] = creds.to_json()
        st.rerun()
    else:
        flow = Flow.from_client_config(CLIENT_CONFIG, scopes=SCOPES)
        flow.redirect_uri = st.secrets["gmail"]["redirect_uri"]
        auth_url, _ = flow.authorization_url(prompt="consent", access_type="offline", include_granted_scopes="true")
        st.markdown(f"### 🔑 Please [authorize the app]({auth_url}) to send emails using your Gmail account.")
        st.stop()

creds = Credentials.from_authorized_user_info(json.loads(st.session_state["creds"]), SCOPES)
service = build("gmail", "v1", credentials=creds)

# ========================================
# Session Setup
# ========================================
if "sending" not in st.session_state:
    st.session_state["sending"] = False
if "done" not in st.session_state:
    st.session_state["done"] = False
if "batch_completed" not in st.session_state:
    st.session_state["batch_completed"] = False

# ========================================
# UI Logic
# ========================================
if not st.session_state["sending"]:
    st.subheader("📤 Step 1: Upload Recipient List")
    uploaded_file = st.file_uploader("Upload CSV or Excel file", type=["csv", "xlsx"])

    if uploaded_file:
        if uploaded_file.name.lower().endswith("csv"):
            try:
                df = pd.read_csv(uploaded_file, encoding="utf-8")
            except UnicodeDecodeError:
                uploaded_file.seek(0)
                df = pd.read_csv(uploaded_file, encoding="latin1")
        else:
            df = pd.read_excel(uploaded_file)

        # Resume logic
        working_df = load_progress_df_if_exists()
        if working_df is not None:
            st.info("🔁 Found an in-progress run — resuming from last saved progress.")
            df = working_df

        for col in ["ThreadId", "RfcMessageId", "Status"]:
            if col not in df.columns:
                df[col] = ""

        st.markdown("### ✏️ Edit Your Contact List")
        df = st.data_editor(df, num_rows="dynamic", use_container_width=True)

        st.markdown("---")
        subject_template = st.text_input("✉️ Subject", "{Company Name}")
        body_template = st.text_area("📝 Body (Markdown)", "Dear {First Name},\n\nWelcome to Mail Merge Demo.\n\nThanks,", height=200)
        label_name = st.text_input("🏷️ Gmail label", "Mail Merge Sent")
        delay = st.slider("⏱️ Delay (seconds)", 20, 75, 25)
        send_mode = st.radio("📬 Mode", ["🆕 New Email", "↩️ Follow-up (Reply)", "💾 Save as Draft"])

        # ===============================
        # Preview First Email Feature
        # ===============================
        preview_first = st.checkbox("👀 Preview First Email", value=False)
        if preview_first and not df.empty:
            first_row = df.iloc[0]
            try:
                subject_preview = subject_template.format(**first_row)
                body_preview = convert_bold(body_template.format(**first_row))
                st.markdown("### ✉️ Preview of First Email")
                st.markdown(f"**Subject:** {subject_preview}")
                st.markdown("**Body:**")
                st.markdown(body_preview, unsafe_allow_html=True)
            except Exception as e:
                st.error(f"⚠️ Could not generate preview: {e}")

        if st.button("🚀 Start Mail Merge"):
            if is_lock_active():
                st.error("⚠️ Run lock exists. Reset before starting new run.")
                st.stop()

            # compute pending indices (safe)
            now = datetime.now()
            pending_indices = []
            for idx in df.index:
                st_status = str(df.at[idx, "Status"]).strip()
                if st_status in ["Sent", "Draft", "Skipped"]:
                    continue
                if st_status.startswith("InProgress"):
                    try:
                        parts = st_status.split("|", 1)
                        if len(parts) == 2:
                            ts = datetime.fromisoformat(parts[1])
                            if now - ts > timedelta(hours=24):
                                pending_indices.append(idx)
                            else:
                                continue
                    except Exception:
                        continue
                else:
                    pending_indices.append(idx)

            lock_info = {"uploader": "user", "file_name": uploaded_file.name, "pending_count": len(pending_indices)}
            if not create_lock(lock_info):
                st.error("⚠️ Could not create run lock. Aborting.")
                st.stop()

            st.session_state.update({
                "sending": True,
                "df": df.fillna(""),
                "pending_indices": pending_indices,
                "subject_template": subject_template,
                "body_template": body_template,
                "label_name": label_name,
                "delay": delay,
                "send_mode": send_mode
            })
            st.rerun()

# ========================================
# Sending Mode
# ========================================
if st.session_state["sending"]:
    df = st.session_state["df"]
    pending_indices = st.session_state["pending_indices"]
    subject_template = st.session_state["subject_template"]
    body_template = st.session_state["body_template"]
    label_name = st.session_state["label_name"]
    delay = st.session_state["delay"]
    send_mode = st.session_state["send_mode"]

    st.subheader("📨 Sending Emails...")
    progress = st.progress(0)
    status_box = st.empty()
    label_id = None
    if send_mode == "🆕 New Email":
        label_id = get_or_create_label(service, label_name)

    total = len(pending_indices)
    sent_message_ids, errors, skipped = [], [], []
    sent_count, batch_count = 0, 0

    try:
        for i, idx in enumerate(pending_indices):
            batch_limit = DRAFT_BATCH_SIZE_DEFAULT if send_mode == "💾 Save as Draft" else BATCH_SIZE_DEFAULT
            if batch_count >= batch_limit:
                break

            row = df.loc[idx]
            if str(row.get("Status", "")).strip() in ["Sent", "Draft"]:
                continue

            df.loc[idx, "Status"] = f"InProgress|{datetime.now().isoformat()}"
            save_progress_df(df)

            pct = int(((i + 1) / total) * 100)
            progress.progress(min(max(pct, 0), 100))
            status_box.info(f"📩 Processing {i + 1}/{total}")

            to_addr = extract_email(str(row.get("Email", "")).strip())
            if not to_addr:
                df.loc[idx, "Status"] = "Skipped"
                save_progress_df(df)
                skipped.append(row.get("Email"))
                continue

            try:
                subject = subject_template.format(**row)
                body_html = convert_bold(body_template.format(**row))
                message = MIMEText(body_html, "html")
                message["To"] = to_addr
                message["Subject"] = subject

                # --- Gmail API call ---
                raw = base64.urlsafe_b64encode(message.as_bytes()).decode()
                msg_body = {"raw": raw}

                if send_mode == "💾 Save as Draft":
                    service.users().drafts().create(userId="me", body={"message": msg_body}).execute()
                    df.loc[idx, "Status"] = "Draft"
                else:
                    sent_msg = service.users().messages().send(userId="me", body=msg_body).execute()
                    msg_id = sent_msg.get("id", "")
                    df.loc[idx, "ThreadId"] = sent_msg.get("threadId", "")
                    df.loc[idx, "RfcMessageId"] = fetch_message_id_header(service, msg_id) or msg_id
                    df.loc[idx, "Status"] = "Sent"
                    if label_id:
                        sent_message_ids.append(msg_id)

                save_progress_df(df)
                sent_count += 1
                batch_count += 1
                time.sleep(random.uniform(delay * 0.9, delay * 1.1))
            except Exception as e:
                df.loc[idx, "Status"] = f"Error|{e}"
                save_progress_df(df)
                errors.append((to_addr, str(e)))

        # Label + backup
        if send_mode != "💾 Save as Draft" and sent_message_ids and label_id:
            try:
                service.users().messages().batchModify(
                    userId="me", body={"ids": sent_message_ids, "addLabelIds": [label_id]}
                ).execute()
            except Exception:
                pass

        # Save CSV
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        file_name = f"Updated_{label_name}_{timestamp}.csv"
        file_path = os.path.join("/tmp", file_name)
        df.to_csv(file_path, index=False)
        send_email_backup(service, file_path)

        with open(DONE_FILE, "w") as f:
            json.dump({"done_time": str(datetime.now()), "file": file_path}, f)

        st.session_state["sending"] = False
        st.session_state["done"] = True
        st.session_state["batch_completed"] = True
        st.session_state["summary"] = {"sent": sent_count, "errors": errors, "skipped": skipped}

    finally:
        remove_lock()
        st.rerun()

# ========================================
# Completion
# ========================================
if st.session_state["done"]:
    s = st.session_state.get("summary", {})
    st.subheader("✅ Mail Merge Completed")
    st.success(f"Sent: {s.get('sent', 0)}")
    if s.get("errors"):
        st.error(f"❌ {len(s['errors'])} errors")
    if s.get("skipped"):
        st.warning(f"⚠️ Skipped: {s['skipped']}")
    if st.button("🔁 Reset for New Run"):
        [os.remove(f) for f in [DONE_FILE, LOCK_FILE, WORKING_CSV] if os.path.exists(f)]
        st.session_state.clear()
        st.experimental_rerun()
