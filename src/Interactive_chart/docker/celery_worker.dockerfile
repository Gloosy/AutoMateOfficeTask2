FROM python:3.9-slim

#Set the workin directory
WORKDIR /code

#Copy the requiremnt file
COPY requiremnt.txt /code/

# Install dependencies
RUN pip install --upgrade pip && pip install -r requiremnt.txt

# Copy the rest of the project files
COPY . /code/

# Command to run the worker
CMD ["celery","-A","interative_chart_project","worker","--loglevel=info"]