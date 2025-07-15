import streamlit as st
import pandas as pd
import tempfile
import os
from main import process_all_files  # You already wrote this!

st.set_page_config(page_title="Name Parser", layout="centered")

st.title("📬 Name and ZIP Parser")

st.markdown("Upload your Excel files from **Mail** and **Email** communications:")

mail_files = st.file_uploader("Upload Mail Files (.xlsx)", type="xlsx", accept_multiple_files=True)
email_files = st.file_uploader("Upload Email Files (.xlsx)", type="xlsx", accept_multiple_files=True)

if st.button("Run Parser"):
    if not mail_files and not email_files:
        st.error("Please upload at least one file.")
    else:
        with st.spinner("Processing files..."):

            # Temp folder to save uploaded files
            with tempfile.TemporaryDirectory() as temp_dir:
                mail_dir = os.path.join(temp_dir, "mail_files")
                email_dir = os.path.join(temp_dir, "email_files")
                os.makedirs(mail_dir, exist_ok=True)
                os.makedirs(email_dir, exist_ok=True)

                for file in mail_files:
                    with open(os.path.join(mail_dir, file.name), "wb") as f:
                        f.write(file.read())

                for file in email_files:
                    with open(os.path.join(email_dir, file.name), "wb") as f:
                        f.write(file.read())

                # Run your parser logic
                all_data = []
                all_data += process_all_files(mail_dir, extract_plan_ssn=False, communication_type="Print")
                all_data += process_all_files(email_dir, extract_plan_ssn=True, communication_type="Email")

                result_df = pd.DataFrame(all_data)

                st.success("✅ Done parsing!")
                st.dataframe(result_df.head(10))

                # Offer download
                output = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
                result_df.to_excel(output.name, index=False)

                with open(output.name, "rb") as f:
                    st.download_button(
                        label="📥 Download parsed Excel",
                        data=f,
                        file_name="parsed_names.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
# Local URL: http://localhost:8501
# Network URL: http://172.20.2.51:8501