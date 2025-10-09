import streamlit as st
import pandas as pd
import base64
import time
import re
import json
from email.mime.text import MIMEText
from google_auth_oauthlib.flow import Flow
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

# ========================================
# Streamlit Page Setup
# ========================================
st.set_page_config(page_title="Gmail Mail Merge", layout="wide")
st.title("📧 Gmail Mail Merge Tool (with Follow-up Replies)")

# ========================================
# Gmail API Setup
# ========================================
SCOPES = [
    "https://www.googleapis.com/auth/gmail.send",
    "https://www.googleapis.com/auth/gmail.modify",
    "https://www.googleapis.com/auth/gmail.labels",
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
# Bold + Link Converter
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
        <body style="font-family: Arial, sans-serif; font-size: 14px; line-height: 1.6;">
            {text}
        </body>
    </html>
    """

# ========================================
# OAuth Flow
# ========================================
if "creds" not in st.session_state:
    st.session_state["creds"] = None

if st.session_state["creds"]:
    creds = Credentials.from_authorized_user_info(
        json.loads(st.session_state["creds"]), SCOPES
    )
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
        st.markdown(
            f"### 🔑 Please [authorize the app]({auth_url}) to send emails using your Gmail account."
        )
        st.stop()

# Build Gmail API client
creds = Credentials.from_authorized_user_info(json.loads(st.session_state["creds"]), SCOPES)
service = build("gmail", "v1", credentials=creds)

# ========================================
# Upload Recipients
# ========================================
st.header("📤 Upload Recipient List")
uploaded_file = st.file_uploader("Upload CSV or Excel file", type=["csv", "xlsx"])

if uploaded_file:
    if uploaded_file.name.endswith("csv"):
        df = pd.read_csv(uploaded_file)
    else:
        df = pd.read_excel(uploaded_file)

    st.write("✅ Preview of uploaded data:")
    st.dataframe(df.head())
    st.info("📌 Include 'ThreadId' and 'RfcMessageId' columns for follow-ups if needed.")

    # ========================================
    # Email Template
    # ========================================
    st.header("✍️ Compose Your Email")
    subject_template = st.text_input("Subject", "Hello {Name}")
    body_template = st.text_area(
        "Body (supports **bold**, [link](https://example.com), and line breaks)",
        """Dear {Name},

Welcome to our **Mail Merge App** demo.

You can add links like [Visit Google](https://google.com)
and preserve formatting exactly.

Thanks,  
**Your Company**""",
        height=250,
    )

    # ========================================
    # Preview Section
    # ========================================
    st.subheader("👁️ Preview Email")
    if not df.empty:
        recipient_options = df["Email"].astype(str).tolist()
        selected_email = st.selectbox("Select recipient to preview", recipient_options)
        try:
            preview_row = df[df["Email"] == selected_email].iloc[0]
            preview_subject = subject_template.format(**preview_row)
            preview_body = body_template.format(**preview_row)
            preview_html = convert_bold(preview_body)
            st.markdown(f"**Subject:** {preview_subject}")
            st.markdown("---")
            st.markdown(preview_html, unsafe_allow_html=True)
        except KeyError as e:
            st.error(f"⚠️ Missing column in data: {e}")

    # ========================================
    # Label & Delay Options
    # ========================================
    st.header("🏷️ Label & Timing Options")
    label_name = st.text_input("Gmail label to apply (new emails only)", value="Mail Merge Sent")
    delay = st.number_input("Delay between emails (seconds)", min_value=0, max_value=60, value=2, step=1)

    # ========================================
    # Send Mode
    # ========================================
    send_mode = st.radio("Choose sending mode", ["🆕 New Email", "↩️ Follow-up (Reply)"])

    if st.button("🚀 Send Emails"):
        label_id = get_or_create_label(service, label_name)
        sent_count = 0
        skipped, errors = [], []

        with st.spinner("📨 Sending emails... please wait."):
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

                    msg_body = {}

                    # ===== Follow-up (Reply) mode =====
                    if send_mode == "↩️ Follow-up (Reply)" and "ThreadId" in row and "RfcMessageId" in row:
                        thread_id = str(row["ThreadId"]).strip()
                        rfc_id = str(row["RfcMessageId"]).strip()

                        if thread_id and thread_id.lower() != "nan" and rfc_id:
                            message["In-Reply-To"] = rfc_id
                            message["References"] = rfc_id
                            raw = base64.urlsafe_b64encode(message.as_bytes()).decode("utf-8")
                            msg_body = {"raw": raw, "threadId": thread_id}
                        else:
                            raw = base64.urlsafe_b64encode(message.as_bytes()).decode("utf-8")
                            msg_body = {"raw": raw}

                    # ===== New Email Mode =====
                    else:
                        raw = base64.urlsafe_b64encode(message.as_bytes()).decode("utf-8")
                        msg_body = {"raw": raw}

                    # Send the email
                    sent_msg = service.users().messages().send(userId="me", body=msg_body).execute()

                    # ✅ FIX: Wait & Retry to Fetch RFC Message-ID
                    time.sleep(1.5)
                    message_id_header = None
                    for attempt in range(3):
                        msg_detail = service.users().messages().get(
                            userId="me",
                            id=sent_msg["id"],
                            format="metadata",
                            metadataHeaders=["Message-ID"],
                        ).execute()

                        message_id_header = next(
                            (h["value"] for h in msg_detail.get("payload", {}).get("headers", [])
                             if h["name"] == "Message-ID"),
                            None,
                        )
                        if message_id_header:
                            break
                        time.sleep(1.5)

                    # Apply label for new emails only
                    if send_mode == "🆕 New Email" and label_id:
                        time.sleep(1)
                        service.users().messages().modify(
                            userId="me",
                            id=sent_msg["id"],
                            body={"addLabelIds": [label_id]},
                        ).execute()

                    # ===== Save ThreadId + Real Message-ID =====
                    if "ThreadId" not in df.columns:
                        df["ThreadId"] = None
                    if "RfcMessageId" not in df.columns:
                        df["RfcMessageId"] = None
                    df.loc[idx, "ThreadId"] = sent_msg.get("threadId", "")
                    df.loc[idx, "RfcMessageId"] = message_id_header or ""

                    sent_count += 1
                    time.sleep(delay)

                except Exception as e:
                    errors.append((to_addr, str(e)))

        # ========================================
        # Summary + CSV Export
        # ========================================
        st.success(f"✅ Successfully sent {sent_count} emails.")
        if skipped:
            st.warning(f"⚠️ Skipped {len(skipped)} invalid emails: {skipped}")
        if errors:
            st.error(f"❌ Failed to send {len(errors)} emails: {errors}")

        csv = df.to_csv(index=False).encode("utf-8")
        st.download_button(
            "⬇️ Download Updated CSV (with ThreadId + RfcMessageId)",
            csv,
            "updated_mailmerge.csv",
            "text/csv",
        )
