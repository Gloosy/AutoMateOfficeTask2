import pandas as pd
import plotly.express as px
import os

def process_excel_data():
    file_path = os.path.join(os.path.dirname(__file__), '..', 'data', 'Financial_Data.xlsx')
    
    try:
        df = pd.read_excel(file_path, sheet_name="Data")
        
        sales_by_country = df.groupby("Country")["Sales"].sum().reset_index()

        fig = px.bar(sales_by_country, x="Country", y="Sales", 
                     title="Financial Data By Country", 
                     labels={"Country": "Country", "Sales": "Total Sales"},
                     color_discrete_sequence=["#00008B"])

        chart_html = fig.to_html(full_html=False)
        return chart_html
    
    except FileNotFoundError:
        return "Error: Financial Data file not found"
