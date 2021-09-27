FROM python:3.9-slim

RUN useradd -ms /app/bash  myuser

ENV APP_HOME /app
WORKDIR $APP_HOME
RUN pip install --upgrade pip
COPY --chown=myuser:myuser requirements.txt requirements.txt
RUN pip install --no-cache-dir -r requirements.txt

ENV PATH="/home/myuser/.local/bin:${PATH}"

COPY --chown=myuser:myuser . .

CMD exec gunicorn --bind :$PORT --workers 1 --threads 8 --timeout 0 main:app