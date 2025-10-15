import streamlit as st
import pandas as pd
import base64
import time
import re
import json
import random
import os
from datetime import datetime, timedelta
import pytz
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
st.title("📧 Gmail Mail Merge Tool (with Follow-up Replies + Draft Save)")

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
# Recovery / Completion Lock
# ========================================
DONE_FILE = "/tmp/mailmerge_done.json"

if os.path.exists(DONE_FILE):
    # ✅ Strong completion lock — stop automatic restarts
    st.session_state["sending"] = False
    st.session_state["done"] = True
    try:
        with open(DONE_FILE, "r") as f:
            done_info = json.load(f)
        file_path = done_info.get("file")
        if file_path and os.path.exists(file_path):
            st.success("✅ Previous mail merge completed successfully.")
            st.download_button(
                "⬇️ Download Updated CSV",
                data=open(file_path, "rb"),
                file_name=os.path.basename(file_path),
                mime="text/csv",
            )
        if st.button("🔁 Reset for New Run"):
            os.remove(DONE_FILE)
            st.session_state.clear()
            st.experimental_rerun()
        st.stop()
    except Exception as e:
        st.warning(f"⚠️ Could not load completion file: {e}")
        st.stop()

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
        r'<a href="\2" style="color:#1a73e8; text-decoration:underline;" target="_blank">\1</a>',
        text,
    )
    text = text.replace("\n", "<br>").replace("  ", "&nbsp;&nbsp;")
    return f"""
    <html>
        <body style="font-family: Verdana, sans-serif; font-size: 14px; line-height: 1.6;">
            {text}
        </body>
    </html>
    """

def get_or_create_label(service, label_name="Mail Merge Sent"):
    try:
        labels = service.users().labels().list(userId="me").execute().get("labels", [])
        for label in labels:
            if label["name"].lower() == label_name.lower():
                return label["id"]
        label_obj = {"name": label_name, "labelListVisibility": "labelShow", "messageListVisibility": "show"}
        created_label = service.users().labels().create(userId="me", body=label_obj).execute()
        return created_label["id"]
    except Exception as e:
        st.warning(f"Could not get/create label: {e}")
        return None

def send_email_backup(service, csv_path):
    try:
        user_profile = service.users().getProfile(userId="me").execute()
        user_email = user_profile.get("emailAddress")

        msg = MIMEMultipart()
        msg["To"] = user_email
        msg["From"] = user_email
        msg["Subject"] = f"📁 Mail Merge Backup CSV - {datetime.now().strftime('%Y-%m-%d %H:%M')}"
        body = MIMEText("Attached is the backup CSV for your mail merge run.", "plain")
        msg.attach(body)

        with open(csv_path, "rb") as f:
            part = MIMEApplication(f.read(), Name=os.path.basename(csv_path))
        part["Content-Disposition"] = f'attachment; filename="{os.path.basename(csv_path)}"'
        msg.attach(part)

        raw = base64.urlsafe_b64encode(msg.as_bytes()).decode()
        service.users().messages().send(userId="me", body={"raw": raw}).execute()
        st.info(f"📧 Backup CSV emailed to {user_email}")
    except Exception as e:
        st.warning(f"⚠️ Could not send backup email: {e}")

def fetch_message_id_header(service, message_id):
    for _ in range(6):
        try:
            msg_detail = service.users().messages().get(
                userId="me", id=message_id, format="metadata", metadataHeaders=["Message-ID"]
            ).execute()
            headers = msg_detail.get("payload", {}).get("headers", [])
            for h in headers:
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
# Session State Setup
# ========================================
if "sending" not in st.session_state:
    st.session_state["sending"] = False
if "done" not in st.session_state:
    st.session_state["done"] = False

# ========================================
# MAIN UI
# ========================================
if not st.session_state["sending"]:
    st.header("📤 Upload Recipient List")
    st.info("⚠️ Upload up to **70–80 contacts** for smooth performance.")

    uploaded_file = st.file_uploader("Upload CSV or Excel", type=["csv", "xlsx"])

    if uploaded_file:
        if uploaded_file.name.lower().endswith("csv"):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)

        st.dataframe(df.head())
        st.info("📌 Include 'ThreadId' and 'RfcMessageId' columns for follow-ups if needed.")
        df = st.data_editor(df, num_rows="dynamic", use_container_width=True)

        subject_template = st.text_input("Subject", "Hello {Name}")
        body_template = st.text_area(
            "Body",
            """Dear {Name},

Welcome to **Mail Merge App** demo.

Thanks,  
**Your Company**""",
            height=250,
        )

        # Template Preview
        if uploaded_file is not None and not df.empty:
            st.markdown("### 👁️ Template Preview")
            try:
                sample_row = df.iloc[0].to_dict()
                preview_subject = subject_template.format(**sample_row)
                preview_body_html = convert_bold(body_template.format(**sample_row))
                st.write(f"**📨 Preview Subject:** {preview_subject}")
                st.markdown(
                    f"<div style='border:1px solid #ddd; padding:10px; border-radius:8px;'>{preview_body_html}</div>",
                    unsafe_allow_html=True,
                )
            except Exception as e:
                st.warning(f"⚠️ Preview unavailable: {e}")

        label_name = st.text_input("Gmail label", "Mail Merge Sent")
        delay = st.slider("Delay (seconds)", 20, 75, 20)

        send_mode = st.radio("Choose mode", ["🆕 New Email", "↩️ Follow-up (Reply)", "💾 Save as Draft"])

        if st.button("🚀 Send Emails / Save Drafts"):
            st.session_state.update({
                "sending": True,
                "df": df,
                "subject_template": subject_template,
                "body_template": body_template,
                "label_name": label_name,
                "delay": delay,
                "send_mode": send_mode
            })
            st.rerun()

# ========================================
# SENDING MODE
# ========================================
if st.session_state["sending"]:
    if st.session_state.get("done", False):
        st.warning("⚠️ This session is already complete. Please reset before starting again.")
        st.stop()

    df = st.session_state["df"]
    subject_template = st.session_state["subject_template"]
    body_template = st.session_state["body_template"]
    label_name = st.session_state["label_name"]
    delay = st.session_state["delay"]
    send_mode = st.session_state["send_mode"]

    st.markdown("<h3>📨 Sending emails... please wait.</h3>", unsafe_allow_html=True)
    progress = st.progress(0)
    status_box = st.empty()

    with open("/tmp/mailmerge_running.json", "w") as f:
        json.dump({"start": str(datetime.now()), "total": len(df)}, f)

    label_id = None
    if send_mode == "🆕 New Email":
        label_id = get_or_create_label(service, label_name)

    if "ThreadId" not in df.columns:
        df["ThreadId"] = None
    if "RfcMessageId" not in df.columns:
        df["RfcMessageId"] = None

    total = len(df)
    if total == 0:
        st.error("❌ No rows found in uploaded file.")
        st.stop()

    sent_count, skipped, errors = 0, [], []

    for idx, row in df.iterrows():
        pct = int(((idx + 1) / total) * 100)
        progress.progress(min(max(pct, 0), 100))
        status_box.info(f"Processing {idx + 1}/{total}")

        to_addr = extract_email(str(row.get("Email", "")).strip())
        if not to_addr:
            skipped.append(row.get("Email"))
            continue

        try:
            subject = subject_template.format(**row)
            body_html = convert_bold(body_template.format(**row))
            message = MIMEText(body_html, "html")
            message["To"] = to_addr
            message["Subject"] = subject

            msg_body = {}
            if send_mode == "↩️ Follow-up (Reply)":
                thread_id = str(row.get("ThreadId", "")).strip()
                rfc_id = str(row.get("RfcMessageId", "")).strip()
                if thread_id and rfc_id:
                    message["In-Reply-To"] = rfc_id
                    message["References"] = rfc_id
                    raw = base64.urlsafe_b64encode(message.as_bytes()).decode()
                    msg_body = {"raw": raw, "threadId": thread_id}
                else:
                    raw = base64.urlsafe_b64encode(message.as_bytes()).decode()
                    msg_body = {"raw": raw}
            else:
                raw = base64.urlsafe_b64encode(message.as_bytes()).decode()
                msg_body = {"raw": raw}

            if send_mode == "💾 Save as Draft":
                draft = service.users().drafts().create(userId="me", body={"message": msg_body}).execute()
                df.loc[idx, "ThreadId"] = draft.get("message", {}).get("threadId", "")
                df.loc[idx, "RfcMessageId"] = draft.get("message", {}).get("id", "")
                st.info(f"📝 Draft saved for {to_addr}")
            else:
                sent_msg = service.users().messages().send(userId="me", body=msg_body).execute()
                msg_id = sent_msg.get("id", "")
                df.loc[idx, "ThreadId"] = sent_msg.get("threadId", "")
                message_id_header = fetch_message_id_header(service, msg_id)
                df.loc[idx, "RfcMessageId"] = message_id_header or msg_id
                if send_mode == "🆕 New Email" and label_id:
                    try:
                        service.users().messages().modify(
                            userId="me", id=msg_id, body={"addLabelIds": [label_id]}
                        ).execute()
                    except Exception:
                        pass
                st.info(f"✅ Sent to {to_addr}")

            sent_count += 1
            if send_mode != "💾 Save as Draft":
                time.sleep(random.uniform(delay * 0.9, delay * 1.1))
        except Exception as e:
            errors.append((to_addr, str(e)))
            st.error(f"Error for {to_addr}: {e}")

    progress.progress(100)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    safe_label = re.sub(r'[^A-Za-z0-9_-]', '_', label_name)
    file_name = f"Updated_{safe_label}_{timestamp}.csv"
    file_path = os.path.join("/tmp", file_name)
    df.to_csv(file_path, index=False)

    try:
        send_email_backup(service, file_path)
    except Exception as e:
        st.warning(f"Backup email failed: {e}")

    with open(DONE_FILE, "w") as f:
        json.dump({"done_time": str(datetime.now()), "file": file_path}, f)

    if os.path.exists("/tmp/mailmerge_running.json"):
        os.remove("/tmp/mailmerge_running.json")

    st.session_state["sending"] = False
    st.session_state["done"] = True
    st.session_state["summary"] = {"sent": sent_count, "errors": errors, "skipped": skipped}

    st.experimental_rerun()

# ========================================
# COMPLETION STATE
# ========================================
if st.session_state["done"]:
    summary = st.session_state.get("summary", {})
    st.success(f"✅ Process completed. Sent: {summary.get('sent', 0)}")
    if summary.get("errors"):
        st.error(f"❌ {len(summary['errors'])} errors occurred.")
    if summary.get("skipped"):
        st.warning(f"⚠️ Skipped: {summary['skipped']}")

    if os.path.exists(DONE_FILE):
        with open(DONE_FILE, "r") as f:
            done_info = json.load(f)
        file_path = done_info.get("file")
        if file_path and os.path.exists(file_path):
            st.download_button(
                "⬇️ Download Updated CSV
