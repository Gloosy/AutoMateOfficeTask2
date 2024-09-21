import os
from celery import Celery

os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'interactive_chart_project.settings')

app = Celery('interactive_chart_project')
app.config_from_object('django.conf:settings', namespace='CELERY')
app.autodiscover_tasks()
