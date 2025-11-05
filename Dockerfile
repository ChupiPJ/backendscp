FROM python:3.11-slim

env debian_frontend=noninteractive

RUN apt-get update && apt-get install -y libreoffice &&  apt-get clean

WORKDIR /app

COPY requirements.txt ./requirements.txt

RUN pip install --no-cache-dir -r requirements.txt

COPY app ./app

COPY Procfile ./Procfile

COPY runtime.txt ./runtime.txt

EXPOSE 10000

CMD ["uvicorn", "app.main:app", "--host", "0.0.0.0", "--port", "10000"]