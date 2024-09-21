from celery.schedules import crontab

app.conf.beat_schedule = {
    'generate-chart-every-week': {
        'task':git  'tasks.generate_weekly_chart',
        'schedule': crontab(0, 0, day_of_week='mon'),  # Every Monday at midnight
    },
}
