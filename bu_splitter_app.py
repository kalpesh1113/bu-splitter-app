import streamlit as st
import pandas as pd
import os
import io
import smtplib
from email.message import EmailMessage
from email.utils import formataddr

# --- Load Secure Credentials from Streamlit Secrets ---
SENDER_EMAIL = st.secrets["SENDER_EMAIL"]
SENDER_NAME = st.secrets["SENDER_NAME"]
APP_PASSWORD = st.secrets["APP_PASSWORD"]
DEFAULT_RECIPIENTS = st.secrets["DEFAULT_RECIPIENTS"]

# --- Page Config ---
st.set_page_config(
    page_title="BU Splitter & Mailer",
    layout="centered"
)

st.title("📂 BU Splitter & Email Tool")

# --- Upload Excel File ---
uploaded_file = st.file_uploader(
    "Upload Excel File",
    type=["xlsx"]
)

if uploaded_file:

    try:
        # --- Read Excel ---
        df = pd.read_excel(
            uploaded_file,
            header=1,
            dtype=str
        )

        # --- Check BU Column ---
        if "BU" not in df.columns:
            st.error("'BU' column not found in file.")

        else:
            st.success("✅ File loaded successfully!")

            # =========================================================
            # BU GROUPING
            # =========================================================
            grouped_bu = {
                "BU_4158.xlsx": ["4158"],
                "BU_4341.xlsx": ["4341"],
                "BU_4697.xlsx": ["4697"],
                "BU_4359_4360.xlsx": ["4359", "4360"]
            }

            # =========================================================
            # REMOVE SMART METER ROWS FUNCTION
            # =========================================================
            def remove_smart_meter_rows(dataframe):

                required_cols = [
                    "Smart Metetered Cons (Y/N)",
                    "Meter_type_code"
                ]

                for col in required_cols:
                    if col not in dataframe.columns:
                        st.error(f"'{col}' column not found in file.")
                        return None

                filtered_df = dataframe[
                    ~(
                        (
                            dataframe["Smart Metetered Cons (Y/N)"]
                            .astype(str)
                            .str.upper()
                            == "Y"
                        )
                        &
                        (
                            dataframe["Meter_type_code"]
                            .astype(str)
                            .str.strip()
                            != "61"
                        )
                    )
                ]

                return filtered_df

            # =========================================================
            # EMAIL SETTINGS
            # =========================================================
            st.markdown("---")
            st.subheader("📧 Email Settings")

            recipients_input = st.text_input(
                "Recipient Emails (comma-separated)",
                value=", ".join(DEFAULT_RECIPIENTS)
            )

            st.markdown("---")

            # =========================================================
            # EXPORT NORMAL BU FILES
            # =========================================================
            if st.button("🚀 Export BU Files"):

                st.info("Generating BU files...")

                for filename, bu_codes in grouped_bu.items():

                    bu_df = df[df["BU"].isin(bu_codes)]

                    if not bu_df.empty:

                        towrite = io.BytesIO()

                        bu_df.to_excel(
                            towrite,
                            index=False,
                            engine='openpyxl'
                        )

                        towrite.seek(0)

                        st.download_button(
                            label=f"Download {filename}",
                            data=towrite,
                            file_name=filename,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )

                st.success("✅ BU Files Ready!")

            # =========================================================
            # EXPORT WITHOUT SMART METER
            # =========================================================
            if st.button("🚀 Export BU Files Without SMART Meter"):

                filtered_df = remove_smart_meter_rows(df)

                if filtered_df is not None:

                    st.success(
                        f"✅ Smart Meter Consumers Removed! Remaining Rows: {len(filtered_df)}"
                    )

                    for filename, bu_codes in grouped_bu.items():

                        bu_df = filtered_df[
                            filtered_df["BU"].isin(bu_codes)
                        ]

                        if not bu_df.empty:

                            towrite = io.BytesIO()

                            bu_df.to_excel(
                                towrite,
                                index=False,
                                engine='openpyxl'
                            )

                            towrite.seek(0)

                            st.download_button(
                                label=f"Download WITHOUT_SMART_{filename}",
                                data=towrite,
                                file_name=f"WITHOUT_SMART_{filename}",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )

            # =========================================================
            # SEND NORMAL EMAIL
            # =========================================================
            if st.button("📤 Send Email with BU Files"):

                try:

                    msg = EmailMessage()

                    filename = uploaded_file.name
                    subject = os.path.splitext(filename)[0]

                    msg['Subject'] = subject
                    msg['From'] = formataddr(
                        (SENDER_NAME, SENDER_EMAIL)
                    )
                    msg['To'] = recipients_input

                    msg.set_content(
                        "Please find attached BU-wise Excel files."
                    )

                    for filename, bu_codes in grouped_bu.items():

                        bu_df = df[df["BU"].isin(bu_codes)]

                        if not bu_df.empty:

                            excel_bytes = io.BytesIO()

                            bu_df.to_excel(
                                excel_bytes,
                                index=False,
                                engine='openpyxl'
                            )

                            excel_bytes.seek(0)

                            msg.add_attachment(
                                excel_bytes.read(),
                                maintype="application",
                                subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                filename=filename
                            )

                    with smtplib.SMTP_SSL(
                        "smtp.gmail.com",
                        465
                    ) as smtp:

                        smtp.login(
                            SENDER_EMAIL,
                            APP_PASSWORD
                        )

                        smtp.send_message(msg)

                    st.success("✅ Email sent successfully!")

                except Exception as e:
                    st.error(f"Email failed: {e}")

            # =========================================================
            # SEND EMAIL WITHOUT SMART METER
            # =========================================================
            if st.button("📤 Send Email with BU Files Without SMART Meter"):

                try:

                    filtered_df = remove_smart_meter_rows(df)

                    if filtered_df is not None:

                        msg = EmailMessage()

                        filename = uploaded_file.name

                        subject = (
                            os.path.splitext(filename)[0]
                            + " WITHOUT SMART METER"
                        )

                        msg['Subject'] = subject

                        msg['From'] = formataddr(
                            (SENDER_NAME, SENDER_EMAIL)
                        )

                        msg['To'] = recipients_input

                        msg.set_content(
                            "Please find attached BU-wise Excel files without SMART meter consumers."
                        )

                        for filename, bu_codes in grouped_bu.items():

                            bu_df = filtered_df[
                                filtered_df["BU"].isin(bu_codes)
                            ]

                            if not bu_df.empty:

                                excel_bytes = io.BytesIO()

                                bu_df.to_excel(
                                    excel_bytes,
                                    index=False,
                                    engine='openpyxl'
                                )

                                excel_bytes.seek(0)

                                msg.add_attachment(
                                    excel_bytes.read(),
                                    maintype="application",
                                    subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    filename=f"WITHOUT_SMART_{filename}"
                                )

                        with smtplib.SMTP_SSL(
                            "smtp.gmail.com",
                            465
                        ) as smtp:

                            smtp.login(
                                SENDER_EMAIL,
                                APP_PASSWORD
                            )

                            smtp.send_message(msg)

                        st.success(
                            "✅ Email sent successfully!"
                        )

                except Exception as e:
                    st.error(f"Email failed: {e}")

    except Exception as e:
        st.error(f"Failed to read file: {e}")

else:
    st.info(
        "Please upload an Excel file with BU column."
    )
