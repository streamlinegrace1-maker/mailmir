# app.py
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

def send_email_backup(service, csv_path):
    """Send the backup CSV to the user's own email inbox (attachment)."""
    try:
        user_profile = service.users().getProfile(userId="me").execute()
        user_email = user_profile.get("emailAddress")

        msg = MIMEMultipart()
        msg["To"] = user_email
        msg["From"] = user_email
        msg["Subject"] = f"📁 Mail Merge Backup CSV - {datetime.now().strftime('%Y-%m-%d %H:%M')}"

        body = MIMEText(
            "Attached is the backup CSV file for your recent mail merge run.\n\nYou can re-upload this file anytime for follow-ups.",
            "plain",
        )
        msg.attach(body)

        with open(csv_path, "rb") as f:
            part = MIMEApplication(f.read(), Name=os.path.basename(csv_path))
        part["Content-Disposition"] = f'attachment; filename="{os.path.basename(csv_path)}"'
        msg.attach(part)

        raw = base64.urlsafe_b64encode(msg.as_bytes()).decode()
        service.users().messages().send(userId="me", body={"raw": raw}).execute()

        st.info(f"📧 Backup CSV emailed to your Gmail inbox ({user_email}).")
    except Exception as e:
        st.warning(f"⚠️ Could not send backup email: {e}")

def fetch_message_id_header(service, message_id):
    """Try to fetch the Message-ID header for a sent message (attempts a few times)."""
    message_id_header = None
    for _ in range(6):
        try:
            msg_detail = service.users().messages().get(
                userId="me",
                id=message_id,
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
            pass
        time.sleep(random.uniform(1, 2))
    return message_id_header or ""

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
        st.markdown(f"### 🔑 Please [authorize the app]({auth_url}) to send emails using your Gmail account.")
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
if "send_progress" not in st.session_state:
    st.session_state["send_progress"] = {"total": 0, "sent": 0}

# ========================================
# MAIN UI (Shown when not sending)
# ========================================
if not st.session_state["sending"]:
    st.header("📤 Upload Recipient List")
    st.info("⚠️ Upload maximum of **70–80 contacts** for smooth operation and to protect your Gmail account.")

    if st.session_state["last_saved_csv"]:
        st.info("📁 Backup from previous session available:")
        st.download_button(
            "⬇️ Download Last Saved CSV",
            data=open(st.session_state["last_saved_csv"], "rb"),
            file_name=st.session_state["last_saved_name"],
            mime="text/csv",
        )

    uploaded_file = st.file_uploader("Upload CSV or Excel file", type=["csv", "xlsx"])

    if uploaded_file:
        if uploaded_file.name.lower().endswith("csv"):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)

        st.write("✅ Preview of uploaded data:")
        st.dataframe(df.head())
        st.info("📌 Include 'ThreadId' and 'RfcMessageId' columns for follow-ups if needed.")

        df = st.data_editor(
            df,
            num_rows="dynamic",
            use_container_width=True,
            key="recipient_editor_inline",
            disabled=False
        )

        # Email Template
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

        # Preview
        st.subheader("👁️ Preview Email")
        if not df.empty:
            recipient_options = df["Email"].astype(str).tolist()
            selected_email = st.selectbox("Select recipient to preview", recipient_options)
            try:
                preview_row = df[df["Email"] == selected_email].iloc[0]
                preview_subject = subject_template.format(**preview_row)
                preview_body = body_template.format(**preview_row)
                preview_html = convert_bold(preview_body)

                st.markdown(
                    f'<span style="font-family: Verdana, sans-serif; font-size:16px;"><b>Subject:</b> {preview_subject}</span>',
                    unsafe_allow_html=True
                )
                st.markdown("---")
                st.markdown(preview_html, unsafe_allow_html=True)
            except KeyError as e:
                st.error(f"⚠️ Missing column in data: {e}")

        # Label & Timing Options
        st.header("🏷️ Label & Timing Options")
        label_name = st.text_input("Gmail label to apply (new emails only)", value="Mail Merge Sent")

        delay = st.slider(
            "Delay between emails (seconds)",
            min_value=20,
            max_value=75,
            value=20,
            step=1,
            help="Minimum 20 seconds delay required for safe Gmail sending."
        )

        eta_ready = st.button("🕒 Ready to Send / Calculate ETA")

        if eta_ready:
            try:
                total_contacts = len(df)
                total_seconds = total_contacts * delay
                total_minutes = total_seconds / 60
                local_tz = pytz.timezone("Asia/Kolkata")
                now_local = datetime.now(local_tz)
                eta_end = now_local + timedelta(seconds=total_seconds)
                st.success(
                    f"📋 Total Recipients: {total_contacts}\n\n"
                    f"⏳ Estimated Duration: {total_minutes:.1f} min\n\n"
                    f"🕒 ETA End: **{eta_end.strftime('%I:%M %p')}**"
                )
            except Exception as e:
                st.warning(f"ETA calculation failed: {e}")

        send_mode = st.radio(
            "Choose sending mode",
            ["🆕 New Email", "↩️ Follow-up (Reply)", "💾 Save as Draft"]
        )

        # When Send button is pressed, stash everything in session_state and flip sending flag
        if st.button("🚀 Send Emails / Save Drafts"):
            # store all required state
            st.session_state["sending"] = True
            st.session_state["df"] = df.copy()
            st.session_state["subject_template"] = subject_template
            st.session_state["body_template"] = body_template
            st.session_state["label_name"] = label_name
            st.session_state["delay"] = delay
            st.session_state["send_mode"] = send_mode
            st.session_state["send_progress"] = {"total": len(df), "sent": 0}
            st.rerun()

# ========================================
# SENDING MODE (UI Hidden) - run the heavy work here
# ========================================
if st.session_state["sending"]:
    # Load everything from session_state
    df = st.session_state.get("df", pd.DataFrame())
    subject_template = st.session_state.get("subject_template", "Hello {Name}")
    body_template = st.session_state.get("body_template", "Hello {Name}")
    label_name = st.session_state.get("label_name", "Mail Merge Sent")
    delay = st.session_state.get("delay", 20)
    send_mode = st.session_state.get("send_mode", "🆕 New Email")

    st.markdown("<h3>📨 Sending emails... please wait.</h3>", unsafe_allow_html=True)
    progress = st.progress(0)
    status_box = st.empty()

    label_id = None
    if send_mode == "🆕 New Email":
        label_id = get_or_create_label(service, label_name)

    sent_count = 0
    skipped = []
    errors = []

    # Ensure columns exist
    if "ThreadId" not in df.columns:
        df["ThreadId"] = None
    if "RfcMessageId" not in df.columns:
        df["RfcMessageId"] = None

    total = len(df)
    if total == 0:
        st.warning("No recipients found in uploaded file.")
    else:
        for idx, row in df.iterrows():
            # Update progress display
            try:
                pct = int((idx / total) * 100)
                progress.progress(pct)
                status_box.info(f"Processing {idx+1}/{total}")
            except Exception:
                pass

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

                # Reply / follow-up handling
                msg_body = {}
                if send_mode == "↩️ Follow-up (Reply)":
                    thread_id = str(row.get("ThreadId", "") or "").strip()
                    rfc_id = str(row.get("RfcMessageId", "") or "").strip()
                    if thread_id and thread_id.lower() != "nan" and rfc_id and rfc_id.lower() != "nan":
                        # set reply headers
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

                # Save as Draft
                if send_mode == "💾 Save as Draft":
                    # create draft
                    body_for_draft = {"message": msg_body}
                    draft = service.users().drafts().create(userId="me", body=body_for_draft).execute()
                    # The draft response may include message id inside draft['message']
                    draft_msg = draft.get("message", {})
                    draft_msg_id = draft_msg.get("id", "")
                    df.loc[idx, "ThreadId"] = draft_msg.get("threadId", "")
                    df.loc[idx, "RfcMessageId"] = draft_msg_id or ""
                    st.info(f"📝 Draft saved for {to_addr}")

                else:
                    # Send message
                    sent_msg = service.users().messages().send(userId="me", body=msg_body).execute()
                    # For messages.send -> response contains 'id' and 'threadId'
                    sent_msg_id = sent_msg.get("id", "")
                    thread_id_resp = sent_msg.get("threadId", "")
                    df.loc[idx, "ThreadId"] = thread_id_resp or df.loc[idx, "ThreadId"]
                    # Try to fetch the Message-ID header (RFC message id)
                    message_id_header = ""
                    if sent_msg_id:
                        message_id_header = fetch_message_id_header(service, sent_msg_id)
                    # fallback to ID if header not found
                    df.loc[idx, "RfcMessageId"] = message_id_header or sent_msg_id or ""
                    st.info(f"✅ Sent to {to_addr}")

                    # Apply label for new emails
                    if send_mode == "🆕 New Email" and label_id and sent_msg.get("id"):
                        try:
                            service.users().messages().modify(
                                userId="me",
                                id=sent_msg["id"],
                                body={"addLabelIds": [label_id]},
                            ).execute()
                        except Exception:
                            st.warning(f"⚠️ Could not apply label to {to_addr}")

                sent_count += 1
                st.session_state["send_progress"]["sent"] = sent_count

                # delay between sends
                if delay > 0 and send_mode != "💾 Save as Draft":  # keep delays for sends; drafts are fast
                    time.sleep(random.uniform(delay * 0.9, delay * 1.1))

            except Exception as e:
                errors.append((to_addr, str(e)))
                st.error(f"Error for {to_addr}: {e}")

    # Finish: save updated CSV and email backup
    try:
        progress.progress(100)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        safe_label = re.sub(r'[^A-Za-z0-9_-]', '_', label_name)
        file_name = f"Updated_{safe_label}_{timestamp}.csv"
        file_path = os.path.join("/tmp", file_name)
        df.to_csv(file_path, index=False)

        st.session_state["last_saved_csv"] = file_path
        st.session_state["last_saved_name"] = file_name

        # Send backup email with CSV attached (works in both modes)
        try:
            send_email_backup(service, file_path)
        except Exception as e:
            st.warning(f"Backup email failed: {e}")

    except Exception as e:
        st.error(f"⚠️ CSV save or backup email failed: {e}")

    # mark done and rerun to show completion UI
    st.session_state["sending"] = False
    st.session_state["done"] = True
    # store summary
    st.session_state["summary"] = {
        "sent_count": sent_count,
        "skipped": skipped,
        "errors": errors,
        "file_path": st.session_state.get("last_saved_csv"),
        "file_name": st.session_state.get("last_saved_name"),
    }
    st.rerun()

# ========================================
# COMPLETION STATE
# ========================================
if st.session_state["done"]:
    summary = st.session_state.get("summary", {})
    sent_count = summary.get("sent_count", 0)
    skipped = summary.get("skipped", [])
    errors = summary.get("errors", [])
    file_path = summary.get("file_path")
    file_name = summary.get("file_name")

    st.success(f"✅ Process completed. Sent: {sent_count}")
    if skipped:
        st.warning(f"⚠️ Skipped invalid emails: {skipped}")
    if errors:
        st.error(f"❌ Errors occurred for {len(errors)} recipients. See logs above.")
        for e in errors:
            st.write(e)

    if file_path and os.path.exists(file_path):
        st.download_button(
            "⬇️ Download Updated CSV",
            data=open(file_path, "rb"),
            file_name=file_name,
            mime="text/csv",
        )

    # allow user to reset for a new run
    if st.button("🔁 New Run / Reset"):
        # clear keys used for previous run
        keys_to_clear = ["df", "subject_template", "body_template", "label_name", "delay", "send_mode", "summary", "last_saved_csv", "last_saved_name"]
        for k in keys_to_clear:
            if k in st.session_state:
                del st.session_state[k]
        st.session_state["sending"] = False
        st.session_state["done"] = False
        st.session_state["send_progress"] = {"total": 0, "sent": 0}
        st.experimental_rerun()
