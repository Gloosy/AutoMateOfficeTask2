import os

# Constants for folder paths (update with your actual paths)
CHARTS_FOLDER = 'path/to/charts'
REPORTS_FOLDER = 'path/to/reports'
EMAIL_LIST_PATH = 'path/to/email_list.xlsx'
MAIL_MERGE_TEMPLATE_PATH = 'path/to/mail_merge_template.docx'
MERGED_FOLDER = 'output/merged_documents'

# Ensure output directories exist
os.makedirs(MERGED_FOLDER, exist_ok=True)
