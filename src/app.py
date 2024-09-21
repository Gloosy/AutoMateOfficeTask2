import os
import streamlit as st
from .excel_processing import process_data_and_generate_chart
from .email_util import send_email
from .mail_merge import mail_merge
from .powerpoint import create_powerpoint

st.title("Automate Office Tasks")

menu = ["Process Excel Data", "Send Emails", "Mail Merge", "Generate PowerPoint"]
choice = st.sidebar.selectbox("Select Task", menu)

if choice == "Process Excel Data":
    st.header("Process Excel Data and Generate Charts")
    uploaded_file = st.file_uploader("Upload Financial Data Excel", type=["xlsx"])
    if uploaded_file:
        file_path = 'data/Financial_Data.xlsx'
        with open(file_path, 'wb') as f:
            f.write(uploaded_file.getbuffer())
        try:
            chart_path, summary = process_data_and_generate_chart(file_path)
            st.success(f"Chart generated and saved at {chart_path}")
            st.bar_chart(summary.set_index('Category'))
        except Exception as e:
            st.error(f"Error processing data: {str(e)}")

elif choice == "Send Emails":
    st.header("Send Emails with Attachments")
    recipient = st.text_input("Recipient Email")
    subject = st.text_input("Subject")
    cc = st.text_input("CC Emails (separated by commas)")
    attachment = st.file_uploader("Upload Attachment", type=["pdf", "docx", "xlsx", "pptx"])
    if st.button("Send Email"):
        attachment_path = ''
        if attachment:
            attachment_path = os.path.join('data/Attachments', attachment.name)
            with open(attachment_path, 'wb') as f:
                f.write(attachment.getbuffer())
        try:
            send_email(recipient, subject, cc, attachment_path)
            st.success("Email sent successfully!")
        except Exception as e:
            st.error(f"Error sending email: {str(e)}")

elif choice == "Mail Merge":
    st.header("Perform Mail Merge")
    uploaded_template = st.file_uploader("Upload Mail Merge Template (Word)", type=["docx"])
    uploaded_data = st.file_uploader("Upload Email List Excel", type=["xlsx"])
    if st.button("Run Mail Merge"):
        if uploaded_template and uploaded_data:
            try:
                with open('src/templates/mail_merge_template.docx', 'wb') as f:
                    f.write(uploaded_template.getbuffer())
                with open('data/Email_List.xlsx', 'wb') as f:
                    f.write(uploaded_data.getbuffer())
                mail_merge()
                st.success("Mail merge completed successfully!")
            except Exception as e:
                st.error(f"Error during mail merge: {str(e)}")
        else:
            st.error("Please upload both template and data files.")

elif choice == "Generate PowerPoint":
    st.header("Generate PowerPoint Presentation")
    if st.button("Create PowerPoint"):
        try:
            ppt_path = create_powerpoint()
            st.success(f"PowerPoint presentation created at {ppt_path}")
            with open(ppt_path, "rb") as file:
                btn = st.download_button(
                    label="Download PowerPoint",
                    data=file,
                    file_name=os.path.basename(ppt_path),
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                )
        except Exception as e:
            st.error(f"Error creating PowerPoint: {str(e)}")
