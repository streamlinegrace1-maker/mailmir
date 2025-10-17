import streamlit as st
import pandas as pd
import base64
import time
import re
import json
import random
from datetime import datetime, timedelta
import pytz
from email.mime.text import MIMEText
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
st.info("⚠️ Upload maximum of **70–80 contacts** for smooth operation and to protect your Gmail account.")

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

            # Subject line preview in Verdana
            st.markdown(
                f'<span style="font-family: Verdana, sans-serif; font-size:16px;"><b>Subject:</b> {preview_subject}</span>',
                unsafe_allow_html=True
            )
            st.markdown("---")
            st.markdown(preview_html, unsafe_allow_html=True)
        except KeyError as e:
            st.error(f"⚠️ Missing column in data: {e}")

    # ========================================
    # Label & Timing Options
    # ========================================
    st.header("🏷️ Label & Timing Options")
    label_name = st.text_input("Gmail label to apply (new emails only)", value="Mail Merge Sent")

    delay = st.slider(
        "Delay between emails (seconds)",
        min_value=20,
        max_value=75,
        value=20,
        step=1,
        help="Minimum 20 seconds delay required for safe Gmail sending. Applies to New, Follow-up, and Draft modes."
    )

    # ========================================
    # ✅ "Ready to Send" Button + ETA (All Modes)
    # ========================================
    eta_ready = st.button("🕒 Ready to Send / Calculate ETA")

    if eta_ready:
        try:
            total_contacts = len(df)
            avg_delay = delay
            total_seconds = total_contacts * avg_delay
            total_minutes = total_seconds / 60

            # Local timezone
            local_tz = pytz.timezone("Asia/Kolkata")  # change if needed
            now_local = datetime.now(local_tz)
            eta_start = now_local
            eta_end = now_local + timedelta(seconds=total_seconds)

            eta_start_str = eta_start.strftime("%I:%M %p")
            eta_end_str = eta_end.strftime("%I:%M %p")

            st.success(
                f"📋 Total Recipients: {total_contacts}\n\n"
                f"⏳ Estimated Duration: {total_minutes:.1f} min (±10%)\n\n"
                f"🕒 ETA Window: **{eta_start_str} – {eta_end_str}** (Local Time)\n\n"
                f"✅ Applies to all send modes: New, Follow-up, Draft"
            )
        except Exception as e:
            st.warning(f"ETA calculation failed: {e}")

    # ========================================
    # Send Mode (with Save Draft)
    # ========================================
    send_mode = st.radio(
        "Choose sending mode",
        ["🆕 New Email", "↩️ Follow-up (Reply)", "💾 Save as Draft"]
    )

    # ========================================
    # Main Send/Draft Button
    # ========================================
    if st.button("🚀 Send Emails / Save Drafts"):
        label_id = get_or_create_label(service, label_name)
        sent_count = 0
        skipped, errors = [], []

        with st.spinner("📨 Processing emails... please wait."):
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
                    else:
                        raw = base64.urlsafe_b64encode(message.as_bytes()).decode("utf-8")
                        msg_body = {"raw": raw}

                    # ===============================
                    # ✉️ Send or Save as Draft
                    # ===============================
                    if send_mode == "💾 Save as Draft":
                        draft = service.users().drafts().create(userId="me", body={"message": msg_body}).execute()
                        sent_msg = draft.get("message", {})
                        st.info(f"📝 Draft saved for {to_addr}")
                    else:
                        sent_msg = service.users().messages().send(userId="me", body=msg_body).execute()

                    # 🕒 Delay between operations
                    if delay > 0:
                        time.sleep(random.uniform(delay * 0.9, delay * 1.1))

                    # ✅ RFC Message-ID Fetch
                    message_id_header = None
                    for attempt in range(5):
                        time.sleep(random.uniform(2, 4))
                        try:
                            msg_detail = service.users().messages().get(
                                userId="me",
                                id=sent_msg.get("id", ""),
                                format="metadata",
                                metadataHeaders=["Message-ID"],
                            ).execute()

                            headers = msg_detail.get("payload", {}).get("headers", [])
                            for h in headers:
                                if h.get("name", "").lower() == "message-id":
                                    message_id_header = h.get("value")
                                    break
                            if message_id_header:
                                break
                        except Exception:
                            continue

                    # 🏷️ Apply label to new emails
                    if send_mode == "🆕 New Email" and label_id and sent_msg.get("id"):
                        success = False
                        for attempt in range(3):
                            try:
                                service.users().messages().modify(
                                    userId="me",
                                    id=sent_msg["id"],
                                    body={"addLabelIds": [label_id]},
                                ).execute()
                                success = True
                                break
                            except Exception:
                                time.sleep(1)
                        if not success:
                            st.warning(f"⚠️ Could not apply label to {to_addr}")

                    df.loc[idx, "ThreadId"] = sent_msg.get("threadId", "")
                    df.loc[idx, "RfcMessageId"] = message_id_header or ""

                    sent_count += 1

                except Exception as e:
                    errors.append((to_addr, str(e)))

        # ========================================
        # Summary
        # ========================================
        if send_mode == "💾 Save as Draft":
            st.success(f"📝 Saved {sent_count} draft(s) to your Gmail Drafts folder.")
        else:
            st.success(f"✅ Successfully processed {sent_count} emails.")

        if skipped:
            st.warning(f"⚠️ Skipped {len(skipped)} invalid emails: {skipped}")
        if errors:
            st.error(f"❌ Failed to process {len(errors)}: {errors}")

        # ========================================
        # CSV Download only for New Email mode
        # ========================================
        if send_mode == "🆕 New Email":
            csv = df.to_csv(index=False).encode("utf-8")
            safe_label = re.sub(r'[^A-Za-z0-9_-]', '_', label_name)
            file_name = f"{safe_label}.csv"

            # Visible download button
            st.download_button(
                "⬇️ Download Updated CSV (Click if not auto-downloaded)",
                csv,
                file_name,
                "text/csv",
                key="manual_download"
            )

            # Auto-download via hidden link
            b64 = base64.b64encode(csv).decode()
            st.markdown(
                f'''
                <a id="auto-download-link" href="data:file/csv;base64,{b64}" download="{file_name}"></a>
                <script>
                    document.getElementById("auto-download-link").click();
                </script>
                ''',
                unsafe_allow_html=True
            )
