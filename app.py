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
import hashlib

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
    # Duplicate-run prevention + internal stop + progress (keeps UI the same)
    # ========================================
    def compute_run_signature(df, subject_template, body_template, send_mode, label_name):
        try:
            sample = df.head(50).to_csv(index=False).encode("utf-8")
        except Exception:
            sample = b""
        key = (
            sample
            + str(subject_template).encode()
            + str(body_template).encode()
            + str(send_mode).encode()
            + str(label_name).encode()
        )
        return hashlib.sha1(key).hexdigest()[:12]

    # session state keys
    if "sending" not in st.session_state:
        st.session_state["sending"] = False
    if "stop_flag" not in st.session_state:
        st.session_state["stop_flag"] = False
    if "completed_run" not in st.session_state:
        st.session_state["completed_run"] = None
    if "progress_bar" not in st.session_state:
        st.session_state["progress_bar"] = None

    current_sig = compute_run_signature(df, subject_template, body_template, send_mode, label_name)

    # Render the same single-row Send button area as original.
    # When sending starts, we'll temporarily show a Stop button in the same place (so UI layout is unchanged)
    if not st.session_state["sending"]:
        send_pressed = st.button("🚀 Send Emails / Save Drafts")
    else:
        # while sending: show a Stop button in same spot so the UI occupies same layout
        stop_pressed = st.button("🛑 Stop Sending")

    # If the user clicked Stop while sending, set stop flag
    # (We handle stop inside the loop.)
    if st.session_state["sending"] and 'stop_pressed' in locals() and stop_pressed:
        st.session_state["stop_flag"] = True
        st.warning("🛑 Stop requested. Current email will finish, then process will halt.")

    # If the user just clicked the Send button
    if 'send_pressed' in locals() and send_pressed:
        # Prevent re-send of exact same job
        if st.session_state["completed_run"] == current_sig:
            st.info("✅ This exact job has already completed. No re-send will occur.")
            st.stop()

        # Prevent concurrent runs
        if st.session_state["sending"]:
            st.warning("⚠️ A send operation is already running. Please wait for it to complete.")
            st.stop()

        # initialize run state
        st.session_state["sending"] = True
        st.session_state["stop_flag"] = False
        sent_count = 0
        skipped, errors = [], []

        # progress UI placeholders (kept inside run so nothing shows when idle)
        progress_bar = st.progress(0)
        progress_text = st.empty()

        try:
            with st.spinner("📨 Processing emails... please wait."):
                if "ThreadId" not in df.columns:
                    df["ThreadId"] = None
                if "RfcMessageId" not in df.columns:
                    df["RfcMessageId"] = None

                label_id = get_or_create_label(service, label_name)

                total_rows = len(df)
                start_time = time.time()

                for idx, row in df.iterrows():
                    # check stop flag each iteration
                    if st.session_state["stop_flag"]:
                        st.warning("🛑 Sending stopped by user.")
                        break

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
                        # ===== Follow-up (Reply) mode handling (if ThreadId & RfcMessageId present)
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

                        # Send or Save as Draft
                        if send_mode == "💾 Save as Draft":
                            draft = service.users().drafts().create(userId="me", body={"message": msg_body}).execute()
                            sent_msg = draft.get("message", {})
                            st.info(f"📝 Draft saved for {to_addr}")
                        else:
                            sent_msg = service.users().messages().send(userId="me", body=msg_body).execute()

                        # Delay between operations (jitter)
                        if delay > 0:
                            time.sleep(random.uniform(delay * 0.9, delay * 1.1))

                        # Try to fetch Message-ID header (best-effort)
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

                        # Apply label to new emails
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

                        # update df identifiers
                        df.loc[idx, "ThreadId"] = sent_msg.get("threadId", "")
                        df.loc[idx, "RfcMessageId"] = message_id_header or ""

                        sent_count += 1

                        # progress update
                        elapsed = time.time() - start_time
                        avg_time = elapsed / sent_count if sent_count > 0 else 0
                        remaining = (total_rows - sent_count) * avg_time
                        progress = sent_count / total_rows if total_rows > 0 else 1.0
                        progress_bar.progress(progress)
                        progress_text.text(
                            f"📤 Sent {sent_count}/{total_rows} ({progress*100:.1f}%) — ETA {remaining/60:.1f} min"
                        )

                    except Exception as e:
                        errors.append((to_addr, str(e)))

            # Summary
            if st.session_state["stop_flag"]:
                st.warning(f"🛑 Stopped manually after {sent_count} emails.")
            else:
                if send_mode == "💾 Save as Draft":
                    st.success(f"📝 Saved {sent_count} draft(s) to your Gmail Drafts folder.")
                else:
                    st.success(f"✅ Successfully processed {sent_count} emails.")

            if skipped:
                st.warning(f"⚠️ Skipped {len(skipped)} invalid emails: {skipped}")
            if errors:
                st.error(f"❌ Failed to process {len(errors)}: {errors}")

            # CSV Download only for New Email mode (manual only)
            if send_mode == "🆕 New Email":
                csv = df.to_csv(index=False).encode("utf-8")
                safe_label = re.sub(r'[^A-Za-z0-9_-]', '_', label_name)
                file_name = f"{safe_label}.csv"

                # Visible download button (manual)
                st.download_button(
                    "⬇️ Download Updated CSV",
                    csv,
                    file_name,
                    "text/csv",
                    key="manual_download"
                )
                st.info("✅ Sending completed. Click above to download your updated CSV.")

            # Mark run as completed to prevent duplicate reruns
            if not st.session_state["stop_flag"]:
                st.session_state["completed_run"] = current_sig

        except Exception as e:
            st.error(f"Send loop failed: {e}")

        finally:
            # always clear the sending flag so UI returns to idle state
            st.session_state["sending"] = False
            st.session_state["stop_flag"] = False
            # ensure progress shows complete
            try:
                progress_bar.progress(1.0)
                progress_text.text("✅ Process complete.")
            except Exception:
                pass
