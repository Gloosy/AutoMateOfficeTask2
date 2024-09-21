from django.urls import path
from charts.views import display_chart

urlpatterns = [
    path('chart/', display_chart, name='display_chart'),
]
