FROM python:3.9-slim

WORKDIR /code

COPY requirements.txt /code/
RUN pip install --upgrade pip && pip install -r requirements.txt

COPY . /code/

# Run celery worker
CMD ["celery", "-A", "interactive_chart_project", "worker", "--loglevel=info"]
