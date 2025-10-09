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
st.set_page_config(page_title="📧 Gmail Mail Merge + Follow-ups", layout="wide")
st.title("📧 Gmail Mail Merge Tool (with Follow-up Replies)")

# ========================================
# Utility Functions
# ========================================
def extract_email(text):
    """Extracts a valid email address from a string."""
    match = re.search(r"[\w\.-]+@[\w\.-]+\.\w+", text)
    return match.group(0) if match else None


def get_header_value(headers, key):
    """Extract a specific header value from Gmail message."""
    for h in headers:
        if h.get("name", "").lower() == key.lower():
            return h.get("value")
    return None


def convert_bold(text):
    """Convert *bold* syntax in plain text to <b> tags for HTML."""
    return re.sub(r"\*(.*?)\*", r"<b>\1</b>", text)


def get_or_create_label(service, label_name="MailMergeSent"):
    """Create or get Gmail label for sent messages."""
    labels = service.users().labels().list(userId="me").execute()
    existing = [l for l in labels.get("labels", []) if l["name"] == label_name]
    if existing:
        return existing[0]["id"]
    new_label = service.users().labels().create(
        userId="me", body={"name": label_name}
    ).execute()
    return new_label["id"]


# ========================================
# Gmail Authentication
# ========================================
st.sidebar.header("🔐 Gmail Authentication")

if "credentials" not in st.session_state:
    st.session_state["credentials"] = None

uploaded_creds = st.sidebar.file_uploader("Upload your Gmail client_secret.json", type=["json"])

if uploaded_creds:
    creds_data = json.load(uploaded_creds)
    st.session_state["creds_data"] = creds_data

redirect_uri = "http://localhost:8501"

if "credentials" not in st.session_state or not st.session_state["credentials"]:
    if st.sidebar.button("🔗 Authenticate with Gmail"):
        flow = Flow.from_client_config(
            st.session_state["creds_data"],
            scopes=[
                "https://www.googleapis.com/auth/gmail.send",
                "https://www.googleapis.com/auth/gmail.modify",
                "https://www.googleapis.com/auth/gmail.readonly",
            ],
            redirect_uri=redirect_uri,
        )
        auth_url, _ = flow.authorization_url(prompt="consent")
        st.sidebar.markdown(f"[Click to authorize Gmail access]({auth_url})")
else:
    st.sidebar.success("✅ Gmail authenticated")

# ========================================
# File Upload and Preview
# ========================================
st.header("📤 Upload Recipient List (with optional Thread Info)")

uploaded_file = st.file_uploader("Upload CSV or Excel file", type=["csv", "xlsx"])
if uploaded_file:
    if uploaded_file.name.endswith("csv"):
        df = pd.read_csv(uploaded_file)
    else:
        df = pd.read_excel(uploaded_file)

    # Ensure thread-tracking columns exist
    for col in ["ThreadId", "RfcMessageId", "LastSentDate", "Status"]:
        if col not in df.columns:
            df[col] = ""

    st.dataframe(df.head())

    # ========================================
    # Email Composition
    # ========================================
    st.subheader("✉️ Compose Your Email")

    subject_template = st.text_input("Subject (use {Name} etc. for placeholders):")
    body_template = st.text_area(
        "Email Body (supports *bold* and {placeholders})",
        height=200,
    )

    delay = st.number_input("Delay between emails (seconds)", 0.5, 10.0, 2.0)

    if st.button("🚀 Send Emails"):
        try:
            st.info("Authenticating Gmail...")
            creds = Credentials.from_authorized_user_file("token.json")
            service = build("gmail", "v1", credentials=creds)
            label_id = get_or_create_label(service, "MailMergeSent")

            sent_count, errors = 0, []

            progress = st.progress(0)
            for idx, row in df.iterrows():
                to_addr = extract_email(str(row.get("Email", "")).strip())
                if not to_addr:
                    continue

                try:
                    subject = subject_template.format(**row)
                    body_text = body_template.format(**row)
                    html_body = convert_bold(body_text)

                    message = MIMEText(html_body, "html")
                    message["To"] = to_addr
                    message["Subject"] = subject

                    # Base64 encode
                    raw = base64.urlsafe_b64encode(message.as_bytes()).decode("utf-8")
                    msg_body = {"raw": raw}

                    # If ThreadId exists → treat as follow-up (reply)
                    thread_id = str(row.get("ThreadId", "")).strip()
                    rfc_message_id = str(row.get("RfcMessageId", "")).strip()

                    if thread_id:
                        if rfc_message_id:
                            message["In-Reply-To"] = rfc_message_id
                            message["References"] = rfc_message_id
                        msg_body["threadId"] = thread_id
                        raw = base64.urlsafe_b64encode(message.as_bytes()).decode("utf-8")
                        msg_body["raw"] = raw

                    # Send email
                    sent_msg = service.users().messages().send(userId="me", body=msg_body).execute()

                    # Fetch message headers
                    full = service.users().messages().get(
                        userId="me", id=sent_msg["id"], format="full"
                    ).execute()
                    headers = full.get("payload", {}).get("headers", [])
                    new_rfc_id = get_header_value(headers, "Message-ID")

                    # Update DataFrame
                    df.at[idx, "ThreadId"] = sent_msg.get("threadId")
                    df.at[idx, "RfcMessageId"] = new_rfc_id
                    df.at[idx, "LastSentDate"] = pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S")
                    df.at[idx, "Status"] = "Sent"

                    sent_count += 1
                    progress.progress(sent_count / len(df))
                    time.sleep(delay)

                except Exception as e:
                    df.at[idx, "Status"] = "Failed"
                    errors.append((to_addr, str(e)))

            st.success(f"✅ Sent {sent_count} emails successfully.")
            if errors:
                st.error(f"❌ {len(errors)} failed. Check logs or retry later.")

            csv_data = df.to_csv(index=False).encode("utf-8")
            st.download_button(
                label="📥 Download Updated CSV (with Thread IDs)",
                data=csv_data,
                file_name="recipients_updated.csv",
                mime="text/csv",
            )

        except Exception as e:
            st.error(f"⚠️ Error: {e}")
