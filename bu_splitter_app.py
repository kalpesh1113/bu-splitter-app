import streamlit as st
import pandas as pd
import os
import io
import smtplib
from email.message import EmailMessage
from email.utils import formataddr
from datetime import datetime

# --- Load Secure Credentials from Streamlit Secrets ---
SENDER_EMAIL = st.secrets["SENDER_EMAIL"]
SENDER_NAME = st.secrets["SENDER_NAME"]
APP_PASSWORD = st.secrets["APP_PASSWORD"]
DEFAULT_RECIPIENTS = st.secrets["DEFAULT_RECIPIENTS"]

st.set_page_config(page_title="BU Splitter & Mailer", layout="centered")
st.title("📂 BU Splitter & Email Tool")

# --- Upload Excel File ---
uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file, header=1, dtype=str)
        if "BU" not in df.columns:
            st.error("'BU' column not found in file.")
        else:
            st.success("File loaded successfully!")

            # --- BU Selection ---
            bu_list = sorted(df["BU"].dropna().unique())
            default_bu = [bu for bu in bu_list if bu in ["4158", "4341", "4359", "4360"]]

            selected_bu = st.multiselect("Select Billing Units (BU)", bu_list, default=default_bu)

            # --- Email Recipients ---
            st.markdown("---")
            st.subheader("📧 Email Settings")
            recipients_input = st.text_input("Recipient Emails (comma-separated)", value=", ".join(DEFAULT_RECIPIENTS))

            # --- Export & Email Buttons ---
            if st.button("🚀 Export BU Files"):
                if not selected_bu:
                    st.warning("Please select at least one BU.")
                else:
                    # Generate and download files (no zip)
                    for bu in selected_bu:
                        bu_df = df[df["BU"] == bu]
                        if not bu_df.empty:
                            towrite = io.BytesIO()
                            bu_df.to_excel(towrite, index=False, engine='openpyxl')
                            towrite.seek(0)
                            st.download_button(
                                label=f"Download BU {bu}",
                                data=towrite,
                                file_name=f"BU_{bu}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )

            if st.button("📤 Send Email with BU Files"):
                if not selected_bu:
                    st.warning("Please select at least one BU.")
                else:
                    try:
                        # Compose Email
                        msg = EmailMessage()
                        filename = uploaded_file.name
                        subject = os.path.splitext(filename)[0]
                        msg['Subject'] = subject
                        msg['From'] = formataddr((SENDER_NAME, SENDER_EMAIL))
                        msg['To'] = recipients_input
                        msg.set_content("Please find attached BU-wise Excel files.")

                        # Attach each BU file individually
                        for bu in selected_bu:
                            bu_df = df[df["BU"] == bu]
                            if not bu_df.empty:
                                excel_bytes = io.BytesIO()
                                bu_df.to_excel(excel_bytes, index=False, engine='openpyxl')
                                excel_bytes.seek(0)
                                msg.add_attachment(
                                    excel_bytes.read(),
                                    maintype="application",
                                    subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    filename=f"BU_{bu}.xlsx"
                                )

                        # Send Email
                        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
                            smtp.login(SENDER_EMAIL, APP_PASSWORD)
                            smtp.send_message(msg)

                        st.success("✅ Email sent successfully!")

                    except Exception as e:
                        st.error(f"Email failed: {e}")

    except Exception as e:
        st.error(f"Failed to read file: {e}")
else:
    st.info("Please upload an Excel file with BU column.")
