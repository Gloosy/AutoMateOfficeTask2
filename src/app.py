import streamlit as st
from config import CHARTS_FOLDER, REPORTS_FOLDER, EMAIL_LIST_PATH, MAIL_MERGE_TEMPLATE_PATH
from functions.pdf import convert_pdf
from functions.email import send_email
from functions.excel import interactive_excel, create_project_structure_template
from functions.powerpoint import create_powerpoint
from functions.mail_merge import mail_merge

st.title("StreamLab Options")

option = st.selectbox(
    "Automate Office Task",
    ("Convert PDF", "Send Mail", "Interactive Excel", "Download Project Structure Template", "Create PowerPoint from Charts", "Mail Merge")
)

if option == "Convert PDF":
    st.subheader("Convert PDF")
    uploaded_file = st.file_uploader("Upload a PDF file", type="pdf")
    if uploaded_file is not None:
        converted_pdf = convert_pdf(uploaded_file)
        st.download_button("Download Converted PDF", converted_pdf, file_name="converted.pdf")

elif option == "Send Mail":
    st.subheader("Send Mail")
    sender_email = st.text_input("Sender Email")
    receiver_email = st.text_input("Receiver Email")
    subject = st.text_input("Subject")
    message = st.text_area("Message")
    
    if st.button("Send Email"):
        send_email(sender_email, receiver_email, subject, message)

elif option == "Interactive Excel":
    st.subheader("Interactive Excel")
    excel_file = st.file_uploader("Upload an Excel file", type="xlsx")
    if excel_file is not None:
        interactive_excel(excel_file)

elif option == "Download Project Structure Template":
    st.subheader("Download Project Structure Excel Template")
    template = create_project_structure_template()
    st.download_button("Download Project Structure Template", template, file_name="project_structure_template.xlsx")

elif option == "Create PowerPoint from Charts":
    st.subheader("Create PowerPoint Presentation from Charts")
    if st.button("Generate PowerPoint"):
        ppt_path = create_powerpoint()
        st.success(f"PowerPoint presentation created: {ppt_path}")
        st.download_button("Download PowerPoint Presentation", open(ppt_path, 'rb'), file_name=os.path.basename(ppt_path))

elif option == "Mail Merge":
    st.subheader("Mail Merge Documents")
    if st.button("Perform Mail Merge"):
        mail_merge()
        st.success("Mail merge completed. Check the output folder for documents.")
