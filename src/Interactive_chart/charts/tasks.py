from celery import shared_task
from .data_processing import process_excel_data

@shared_task
def generate_weekly_chart():
    # Generate the chart and store it
    process_excel_data()
