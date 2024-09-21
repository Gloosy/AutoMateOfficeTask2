import pandas as pd
import plotly.express as px
import os
import requests
from celery import Celery
from celery.schedules import crontab

# Function to process data from Excel and generate a chart
def process_data_and_generate_chart(file_path):
    # Load the data from the Excel file
    data = pd.read_excel(file_path)

    # Check for required columns
    if 'Category' not in data.columns or 'Sales' not in data.columns:
        raise ValueError("Columns 'Category' and 'Sales' are required.")
    
    # Group data by category and sum sales
    summary = data.groupby('Category')['Sales'].sum().reset_index()

    # Create the interactive chart
    fig = px.bar(summary, x='Category', y='Sales', title='Sales by Category', 
                 labels={'Category': 'Category', 'Sales': 'Total Sales'},
                 color_discrete_sequence=["#1f77b4"])

    # Save the chart as an HTML file
    output_path = os.path.join(os.path.dirname(file_path), 'sales_by_category.html')
    fig.write_html(output_path)

    return output_path  # Return the path of the generated chart

# Function to send the chart to the Gemini API
def send_chart_to_gemini_api(chart_path):
    url = "https://api.gemini.com/v1/chart/upload"  # Replace with actual Gemini API endpoint
    with open(chart_path, 'rb') as f:
        response = requests.post(url, files={'file': f})
    return response.json()

# Initialize Celery
app = Celery('tasks', broker='redis://localhost:6379/0')

# Define the Celery task to generate the weekly chart
@app.task
def generate_weekly_chart():
    file_path = r'D:\AI_TUX\Python\AutoMateOfficeTask2\interactive_chart_project\Financial_Data.xlsx'  # Use raw string for Windows paths
    chart_path = process_data_and_generate_chart(file_path)
    response = send_chart_to_gemini_api(chart_path)
    return response

@app.on_after_configure.connect
def setup_periodic_tasks(sender, **kwargs):
    # Calls generate_weekly_chart every Monday at 8:00 AM
    sender.add_periodic_task(crontab(hour=8, minute=0, day_of_week='mon'), generate_weekly_chart.s())

# Start the Celery worker and beat in separate terminals
