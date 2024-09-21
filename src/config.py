import os
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

EMAIL_ADDRESS = os.getenv('EMAIL_ADDRESS')
EMAIL_PASSWORD = os.getenv('EMAIL_PASSWORD')
REDIS_URL = os.getenv('REDIS_URL', 'redis://localhost:6379/0')

# File paths
EXCEL_FILE_PATH = os.path.abspath('../data/Financial_Data.xlsx')
EMAIL_LIST_PATH = os.path.abspath('../data/Email_List.xlsx')
EMAIL_TEMPLATE_PATH = os.path.abspath('./templates/email_template.html')
MAIL_MERGE_TEMPLATE_PATH = os.path.abspath('./templates/mail_merge_template.docx')
CHARTS_FOLDER = os.path.abspath('../charts/')
MERGED_FOLDER = os.path.abspath('../merged/')
REPORTS_FOLDER = os.path.abspath('../reports/')
ATTACHMENTS_FOLDER = os.path.abspath('../data/Attachments/')
