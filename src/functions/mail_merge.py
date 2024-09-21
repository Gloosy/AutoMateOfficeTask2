import os
import pandas as pd
from docx import Document
from config import MAIL_MERGE_TEMPLATE_PATH, EMAIL_LIST_PATH, MERGED_FOLDER

def mail_merge(template_path=MAIL_MERGE_TEMPLATE_PATH, excel_data_path=EMAIL_LIST_PATH, output_folder=MERGED_FOLDER):
    os.makedirs(output_folder, exist_ok=True)
    df = pd.read_excel(excel_data_path, sheet_name='Email_List')

    for _, row in df.iterrows():
        recipient_name = row['Recipient Name']
        document = Document(template_path)

        for paragraph in document.paragraphs:
            if '[[Recipient]]' in paragraph.text:
                paragraph.text = paragraph.text.replace('[[Recipient]]', recipient_name)

        output_path = os.path.join(output_folder, f"{recipient_name}_letter.docx")
        document.save(output_path)
