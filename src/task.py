from celery import Celery
from celery.schedules import crontab
from .config import REDIS_URL
from .functions.excel import process_data_and_generate_chart
from .functions.powerpoint import create_powerpoint
from .email_utils import send_email
from .functions.mail_merge import mail_merge
import os

# Initialize Celery
app = Celery('tasks', broker=REDIS_URL, backend=REDIS_URL)

# Configure Celery Beat schedule
app.conf.beat_schedule = {
    'generate-and-send-weekly-report': {
        'task': 'src.tasks.generate_weekly_report',
        'schedule': crontab(hour=8, minute=0, day_of_week='mon'),  # Every Monday at 8:00 AM
    },
}
app.conf.timezone = 'UTC'

@app.task
def generate_weekly_report():
    """
    Celery task to generate weekly financial report and send it via email.
    """
    # Process Excel and generate chart
    chart_path, summary = process_data_and_generate_chart()

    # Create PowerPoint presentation
    ppt_path = create_powerpoint()

    # Perform mail merge
    mail_merge()

    # Send email with PowerPoint report
    subject = f"Weekly Financial Report - {pd.Timestamp.today().strftime('%Y-%m-%d')}"
    recipient_email = "team@example.com"  # Replace with actual recipient
    cc_emails = "manager@example.com"      # Replace with actual CC
    attachment_path = ppt_path

    send_email(recipient_email, subject, cc_emails, attachment_path)
