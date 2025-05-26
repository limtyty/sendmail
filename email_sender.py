import streamlit as st
import pandas as pd
import smtplib
import imaplib
import time
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from string import Template

st.title("📧 Bulk HTML Email Sender (with Batching + IMAP Sent Logging)")

# Step 1: Upload HTML Template
html_file = st.file_uploader("Upload HTML Template (use {{name}} for placeholder)", type="html")
html_template = None
if html_file is not None:
    html_template = html_file.read().decode("utf-8")
    st.success("HTML template uploaded successfully.")

# Step 2: Upload CSV file with email list
csv_file = st.file_uploader("Upload CSV File (with 'name' and 'email' columns)", type="csv")
df = None
if csv_file is not None:
    df = pd.read_csv(csv_file)
    if 'name' in df.columns and 'email' in df.columns:
        df['Send'] = True
        selected_rows = st.data_editor(df, use_container_width=True, num_rows="dynamic")
    else:
        st.error("CSV must contain 'name' and 'email' columns.")

# Step 3: Email credentials input
st.subheader("Email Sender Credentials")
# CC Email Input
cc_email = st.text_input("CC Email Address (optional)")
sender_email = st.text_input("Your Email Address",value="registration@canadacambodiamission.com")
sender_password = st.text_input("Your App Password", type="password",)
smtp_server = st.text_input("SMTP Server", value="smtp.hostinger.com")
smtp_port = st.number_input("SMTP Port", value=465)

# Optional: IMAP Sent Folder
imap_folder = st.text_input("IMAP 'TeamCanada' Folder Name", value="TeamCanada")

# Step 4: Batch Controls
st.subheader("Batch Sending Controls")
batch_size = st.number_input("Batch Size", min_value=1, value=200)
delay_between_emails = st.number_input("Delay Between Emails (seconds)", min_value=0, value=2)
delay_between_batches = st.number_input("Delay Between Batches (seconds)", min_value=0, value=60)

# Step 5: Subject Line
subject = st.text_input("Email Subject", value="Invitation to Team Canada Trade Mission to Cambodia")

# Step 6: Send Emails
if st.button("Send Emails"):
    if not html_template or df is None:
        st.error("Upload both HTML template and CSV file.")
    else:
        try:
            # Login to SMTP
            server = smtplib.SMTP_SSL(smtp_server, smtp_port)
            server.login(sender_email, sender_password)

            # Login to IMAP
            imap_server_host = smtp_server.replace("smtp", "imap")
            imap = imaplib.IMAP4_SSL(imap_server_host)
            imap.login(sender_email, sender_password)

            to_send = selected_rows[selected_rows['Send']].reset_index(drop=True)
            total = len(to_send)

            progress_bar = st.progress(0)

            for i in range(0, total, batch_size):
                batch = to_send.iloc[i:i+batch_size]
                st.info(f"Sending batch {i//batch_size + 1}...")

                for idx, (_, row) in enumerate(batch.iterrows()):
                    try:
                        # Prepare email
                        msg = MIMEMultipart("alternative")
                        msg["Subject"] = subject
                        msg["From"] = f"Registration Trade Mission Events <{sender_email}>"
                        msg["To"] = row["email"]
                        if cc_email.strip():
                            msg["Cc"] = cc_email

                        # Substitute placeholders
                        html_content = html_template.replace("{{name}}", row["name"])
                        # html_content = Template(html_template).safe_substitute(name=row["name"])
                        part = MIMEText(html_content, "html")
                        msg.attach(part)

                        # Send via SMTP
                        recipients = [row["email"]]
                        if cc_email.strip():
                            recipients.append(cc_email)
                        server.sendmail(sender_email, recipients, msg.as_string())
                        st.write(f"✅ Sent to {row['email']}")

                        # Save to IMAP "Sent" folder
                        imap.append('"INBOX.Sent"', '\\Seen', imaplib.Time2Internaldate(time.time()), msg.as_bytes())
                        # imap.append(f'"{imap_folder}"', '', imaplib.Time2Internaldate(time.time()), msg.as_bytes())

                        # Delay
                        time.sleep(delay_between_emails)
                        progress_bar.progress(min(1.0, (i + idx + 1) / total))

                    except Exception as e:
                        st.error(f"❌ Failed to send to {row['email']}: {e}")

                st.success(f"✅ Batch {i//batch_size + 1} sent.")
                if i + batch_size < total:
                    time.sleep(delay_between_batches)

            server.quit()
            imap.logout()
            st.success("🎉 All emails sent and saved to Sent folder successfully!")

        except Exception as e:
            st.error(f"❌ Error during sending process: {e}")