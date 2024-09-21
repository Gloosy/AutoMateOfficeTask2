from django.shortcuts import render
from .data_processing import process_excel_data

def display_chart(request):
    chart_html = process_excel_data()
    return render(request,'charts.display_chart.html',{'chart' : chart_html})
    