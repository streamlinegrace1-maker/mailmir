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
    <html><body style="font-family: Verdana, sans-serif; font-size: 14px; line-height: 1.6;">
    {text}
    </body></html>
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
# Sending Mode
# ========================================
if "sending_mode" not in st.session_state:
    st.session_state["sending_mode"] = False

# ========================================
# Upload Recipients
# ========================================
if not st.session_state["sending_mode"]:
    st.header("📤 Upload Recipient List")
    st.info("⚠️ Upload up to 70–80 contacts for safe Gmail sending.")

    if "last_saved_csv" in st.session_state:
        st.info("📁 Backup from previous session available:")
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

        st.write("✅ Preview of uploaded data:")
        st.dataframe(df.head())
        df = st.data_editor(df, num_rows="dynamic", use_container_width=True, key="edit")

        st.header("✍️ Compose Your Email")
        subject_template = st.text_input("Subject", "Hello {Name}")
        body_template = st.text_area(
            "Body", """Dear {Name},\n\nWelcome to our **Mail Merge App** demo.\n\nThanks,\n**Your Company**""", height=250
        )

        st.subheader("👁️ Preview Email")
        if not df.empty:
            recipient_options = df["Email"].astype(str).tolist()
            selected_email = st.selectbox("Select recipient to preview", recipient_options)
            try:
                preview_row = df[df["Email"] == selected_email].iloc[0]
                preview_subject = subject_template.format(**preview_row)
                preview_body = body_template.format(**preview_row)
                st.markdown(f"**Subject:** {preview_subject}")
                st.markdown(convert_bold(preview_body), unsafe_allow_html=True)
            except KeyError as e:
                st.error(f"Missing column: {e}")

        st.header("🏷️ Label & Timing Options")
        label_name = st.text_input("Gmail label", "Mail Merge Sent")

        delay = st.slider("Delay between emails (seconds)", 20, 75, 20, 1)

        eta_ready = st.button("🕒 Ready to Send / Calculate ETA")
        if eta_ready:
            try:
                total_contacts = len(df)
                min_delay = delay * 0.9
                max_delay = delay * 1.1
                min_total = total_contacts * min_delay
                max_total = total_contacts * max_delay
                local_tz = pytz.timezone("Asia/Kolkata")
                now = datetime.now(local_tz)
                eta_min = now + timedelta(seconds=min_total)
                eta_max = now + timedelta(seconds=max_total)
                st.success(
                    f"📋 Total: {total_contacts}\n\n"
                    f"⏳ Duration: {min_total/60:.1f}–{max_total/60:.1f} min\n\n"
                    f"🕒 ETA: {eta_min.strftime('%I:%M %p')}–{eta_max.strftime('%I:%M %p')}"
                )
            except Exception as e:
                st.warning(f"ETA calc failed: {e}")

        send_mode = st.radio("Choose mode", ["🆕 New Email", "↩️ Follow-up (Reply)", "💾 Save as Draft"])

        if st.button("🚀 Send Emails / Save Drafts"):
            st.session_state["sending_mode"] = True
            st.session_state["send_payload"] = {
                "df": df.to_dict(),
                "subject_template": subject_template,
                "body_template": body_template,
                "label_name": label_name,
                "delay": delay,
                "send_mode": send_mode,
            }
            st.rerun()

else:
    st.header("📬 Sending in Progress...")
    data = st.session_state["send_payload"]
    df = pd.DataFrame(data["df"])
    delay = data["delay"]
    label_name = data["label_name"]
    send_mode = data["send_mode"]
    subject_template = data["subject_template"]
    body_template = data["body_template"]

    progress = st.progress(0)
    status = st.empty()

    # ========================================
    # Fixed Backup Email Function
    # ========================================
    def send_email_backup(service, csv_path):
        try:
            user_profile = service.users().getProfile(userId="me").execute()
            user_email = user_profile.get("emailAddress")

            msg = MIMEMultipart()
            msg["To"] = user_email
            msg["From"] = user_email
            msg["Subject"] = f"📁 Mail Merge Backup CSV - {datetime.now().strftime('%Y-%m-%d %H:%M')}"

            # Email body
            body_text = "Hello,\n\nYour Gmail Mail Merge backup CSV is attached.\n\nBest,\nMail Merge Tool"
            msg.attach(MIMEText(body_text, "plain"))

            # Attach CSV
            with open(csv_path, "rb") as f:
                part = MIMEApplication(f.read(), _subtype="csv")
                part.add_header("Content-Disposition", "attachment", filename=os.path.basename(csv_path))
                msg.attach(part)

            raw = base64.urlsafe_b64encode(msg.as_bytes()).decode()
            service.users().messages().send(userId="me", body={"raw": raw}).execute()
            st.success("✅ Backup CSV sent to your Gmail inbox.")

        except Exception as e:
            st.warning(f"⚠️ Backup email failed to send: {e}")

    label_id = get_or_create_label(service, label_name)
    sent_count, errors = 0, []

    for idx, row in df.iterrows():
        try:
            to_addr = extract_email(str(row.get("Email", "")).strip())
            if not to_addr:
                continue
            subject = subject_template.format(**row)
            body_html = convert_bold(body_template.format(**row))
            message = MIMEText(body_html, "html")
            message["To"] = to_addr
            message["Subject"] = subject
            raw = base64.urlsafe_b64encode(message.as_bytes()).decode("utf-8")
            msg_body = {"raw": raw}

            if send_mode == "💾 Save as Draft":
                sent_msg = service.users().drafts().create(userId="me", body={"message": msg_body}).execute()
            else:
                sent_msg = service.users().messages().send(userId="me", body=msg_body).execute()

            if send_mode == "🆕 New Email" and label_id:
                service.users().messages().modify(
                    userId="me", id=sent_msg["id"], body={"addLabelIds": [label_id]}
                ).execute()

            sent_count += 1
            progress.progress(sent_count / len(df))
            status.text(f"Sent {sent_count}/{len(df)} to {to_addr}")
            time.sleep(random.uniform(delay * 0.9, delay * 1.1))

        except Exception as e:
            errors.append((to_addr, str(e)))

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    file_name = f"MailMerge_Backup_{timestamp}.csv"
    file_path = os.path.join("/tmp", file_name)
    df.to_csv(file_path, index=False)
    st.success(f"✅ Completed! Sent {sent_count}/{len(df)} emails.")
    st.download_button("⬇️ Download Updated CSV", open(file_path, "rb"), file_name=file_name)

    # Send backup CSV to yourself
    send_email_backup(service, file_path)

    if st.button("⬅ Back to Mail Merge"):
        st.session_state["sending_mode"] = False
        st.rerun()
