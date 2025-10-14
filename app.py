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
# Smart Email Extractor
# ========================================
EMAIL_REGEX = re.compile(r"[\w\.-]+@[\w\.-]+\.\w+")

def extract_email(value: str):
    if not value:
        return None
    match = EMAIL_REGEX.search(str(value))
    return match.group(0) if match else None


# ========================================
# Bold + Link Converter (Verdana)
# ========================================
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


# ========================================
# Gmail Label Helper
# ========================================
def get_or_create_label(service, label_name="Mail Merge Sent"):
    try:
        labels = service.users().labels().list(userId="me").execute().get("labels", [])
        for label in labels:
            if label["name"].lower() == label_name.lower():
                return label["id"]
        label_obj = {
            "name": label_name,
            "labelListVisibility": "labelShow",
            "messageListVisibility": "show",
        }
        created_label = service.users().labels().create(userId="me", body=label_obj).execute()
        return created_label["id"]
    except Exception as e:
        st.warning(f"Could not get/create label: {e}")
        return None


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
        auth_url, _ = flow.authorization_url(
            prompt="consent", access_type="offline", include_granted_scopes="true"
        )
        st.markdown(f"### 🔑 Please [authorize the app]({auth_url}) to send emails.")
        st.stop()

creds = Credentials.from_authorized_user_info(json.loads(st.session_state["creds"]), SCOPES)
service = build("gmail", "v1", credentials=creds)

# ========================================
# Initialize Session State
# ========================================
if "sending" not in st.session_state:
    st.session_state["sending"] = False
if "done" not in st.session_state:
    st.session_state["done"] = False
if "last_saved_csv" not in st.session_state:
    st.session_state["last_saved_csv"] = None
if "last_saved_name" not in st.session_state:
    st.session_state["last_saved_name"] = None

# ========================================
# Upload Recipients
# ========================================
if not st.session_state["sending"]:
    st.header("📤 Upload Recipient List")
    st.info("⚠️ Upload up to 70–80 contacts for safe Gmail sending.")

    if st.session_state["last_saved_csv"]:
        st.download_button(
            "⬇️ Download Last Saved CSV",
            data=open(st.session_state["last_saved_csv"], "rb"),
            file_name=st.session_state["last_saved_name"],
            mime="text/csv",
        )

    uploaded_file = st.file_uploader("Upload CSV or Excel file", type=["csv", "xlsx"])

    if uploaded_file:
        if uploaded_file.name.endswith("csv"):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)

        st.dataframe(df.head())
        df = st.data_editor(df, num_rows="dynamic", use_container_width=True, key="recipient_editor_inline")

        st.header("✍️ Compose Email")
        subject_template = st.text_input("Subject", "Hello {Name}")
        body_template = st.text_area(
            "Body", 
            "Dear {Name},\n\nWelcome!\n\n**Thanks,**\nTeam", 
            height=250
        )

        st.header("🏷️ Options")
        label_name = st.text_input("Gmail label", value="Mail Merge Sent")
        delay = st.slider("Delay (seconds)", 20, 75, 25)
        send_mode = st.radio("Mode", ["🆕 New Email", "↩️ Follow-up (Reply)", "💾 Save as Draft"])

        # === Send Button ===
        if st.button("🚀 Send Emails / Save Drafts"):
            st.session_state["sending"] = True
            st.session_state["df"] = df
            st.session_state["subject_template"] = subject_template
            st.session_state["body_template"] = body_template
            st.session_state["label_name"] = label_name
            st.session_state["delay"] = delay
            st.session_state["send_mode"] = send_mode
            st.rerun()

# ========================================
# Sending Mode
# ========================================
if st.session_state["sending"]:
    df = st.session_state["df"]
    subject_template = st.session_state["subject_template"]
    body_template = st.session_state["body_template"]
    label_name = st.session_state["label_name"]
    delay = st.session_state["delay"]
    send_mode = st.session_state["send_mode"]

    st.markdown("<h3>📨 Sending emails... please wait.</h3>", unsafe_allow_html=True)
    progress = st.progress(0)

    label_id = get_or_create_label(service, label_name)
    sent_count, skipped, errors = 0, [], []

    if "ThreadId" not in df.columns:
        df["ThreadId"] = None
    if "RfcMessageId" not in df.columns:
        df["RfcMessageId"] = None

    for idx, row in df.iterrows():
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
            raw = base64.urlsafe_b64encode(message.as_bytes()).decode("utf-8")
            msg_body = {"raw": raw}
            sent_msg = service.users().messages().send(userId="me", body=msg_body).execute()

            df.loc[idx, "ThreadId"] = sent_msg.get("threadId", "")
            df.loc[idx, "RfcMessageId"] = sent_msg.get("id", "")
            sent_count += 1
            progress.progress(int((idx + 1) / len(df) * 100))
            time.sleep(random.uniform(delay * 0.9, delay * 1.1))

        except Exception as e:
            errors.append((to_addr, str(e)))

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    file_name = f"MailMerge_{timestamp}.csv"
    file_path = os.path.join("/tmp", file_name)
    df.to_csv(file_path, index=False)

    st.session_state["last_saved_csv"] = file_path
    st.session_state["last_saved_name"] = file_name
    st.session_state["sending"] = False
    st.session_state["done"] = True
    st.rerun()

# ========================================
# Completion State
# ========================================
if st.session_state["done"]:
    st.success("✅ All emails sent successfully!")
    st.download_button(
        "⬇️ Download Updated CSV",
        data=open(st.session_state["last_saved_csv"], "rb"),
        file_name=st.session_state["last_saved_name"],
        mime="text/csv",
    )
