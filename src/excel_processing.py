import os
import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt
from pptx import Presentation
from pptx.util import Inches
from .config import EXCEL_FILE_PATH, CHARTS_FOLDER, REPORTS_FOLDER

def process_data_and_generate_chart(file_path=EXCEL_FILE_PATH, charts_folder=CHARTS_FOLDER):
    """
    Reads Excel data, processes it, and generates a chart.
    Saves the chart as an image in the charts folder.
    """
    # Load the data from the Excel file
    data = pd.read_excel(file_path)

    # Check if required columns exist
    if 'Category' not in data.columns or 'Sales' not in data.columns:
        raise ValueError("The Excel file must contain 'Category' and 'Sales' columns.")

    # Summarize the sales data by category
    summary = data.groupby('Category')['Sales'].sum().reset_index()

    # Create a bar chart using Seaborn
    plt.figure(figsize=(10, 6))
    sns.barplot(x='Category', y='Sales', data=summary, palette="Blues_d")
    plt.title('Sales by Category')
    plt.xlabel('Category')
    plt.ylabel('Total Sales')
    plt.tight_layout()

    # Save the chart
    chart_filename = 'sales_by_category.png'
    chart_path = os.path.join(charts_folder, chart_filename)
    plt.savefig(chart_path)
    plt.close()

    return chart_path, summary
