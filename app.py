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
st.title("📧 Gmail Mail Merge Tool")

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
    """Extracts the first valid email from a string, or None if not found."""
    if not value:
        return None
    match = EMAIL_REGEX.search(str(value))
    return match.group(0) if match else None

# ========================================
# Gmail Label Helper
# ========================================
def get_or_create_label(service, label_name="Mail Merge Sent"):
    """Returns the label ID for the given label name, creates it if missing."""
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
    """
    Converts **bold** syntax and [text](url) to working HTML.
    Preserves spacing, line breaks, and Gmail-style link formatting.
    """
    if not text:
        return ""

    # Bold conversion
    text = re.sub(r"\*\*(.*?)\*\*", r"<b>\1</b>", text)

    # Link conversion [text](https://example.com)
    text = re.sub(
        r"\[(.*?)\]\((https?://[^\s)]+)\)",
        r'<a href="\2" style="color:#1a73e8; text-decoration:underline;" target="_blank">\1</a>',
        text,
    )

    # Preserve newlines & spaces
    text = text.replace("\n", "<br>").replace("  ", "&nbsp;&nbsp;")

    # Wrap with full HTML for Gmail rendering
    html_body = f"""
    <html>
        <body style="font-family: Arial, sans-serif; font-size: 14px; line-height: 1.6;">
            {text}
        </body>
    </html>
    """
    return html_body

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
and preserve spacing or formatting exactly.

Thanks,  
**Your Company**
""",
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
        except Exception as e:
            st.error(f"Error rendering preview: {e}")

    # ========================================
    # Label & Delay Options
    # ========================================
    st.header("🏷️ Label & Timing Options")
    label_name = st.text_input("Gmail label to apply", value="Mail Merge Sent")
    delay = st.number_input("Delay between emails (seconds)", min_value=0, max_value=60, value=2, step=1)

    # ========================================
    # Send Emails (Option 1 — Label After Send)
    # ========================================
    if st.button("🚀 Send Emails"):
        label_id = get_or_create_label(service, label_name)
        sent_count = 0
        skipped = []
        errors = []

        with st.spinner("📨 Sending emails... please wait."):
            for idx, row in df.iterrows():
                to_addr_raw = str(row.get("Email", "")).strip()
                to_addr = extract_email(to_addr_raw)
                if not to_addr:
                    skipped.append(to_addr_raw)
                    continue

                try:
                    subject = subject_template.format(**row)
                    body_text = body_template.format(**row)
                    html_body = convert_bold(body_text)

                    # Build and send HTML email
                    message = MIMEText(html_body, "html")
                    message["to"] = to_addr
                    message["subject"] = subject
                    raw = base64.urlsafe_b64encode(message.as_bytes()).decode("utf-8")
                    msg_body = {"raw": raw}

                    # ✅ Step 1: Send email
                    sent_msg = service.users().messages().send(userId="me", body=msg_body).execute()

                    # ✅ Step 2: Apply label after sending
                    if label_id:
                        service.users().messages().modify(
                            userId="me",
                            id=sent_msg["id"],
                            body={"addLabelIds": [label_id]},
                        ).execute()

                    sent_count += 1
                    time.sleep(delay)

                except Exception as e:
                    errors.append((to_addr, str(e)))

        # ========================================
        # Summary
        # ========================================
        st.success(f"✅ Successfully sent {sent_count} emails.")
        if skipped:
            st.warning(f"⚠️ Skipped {len(skipped)} invalid emails: {skipped}")
        if errors:
            st.error(f"❌ Failed to send {len(errors)} emails: {errors}")
