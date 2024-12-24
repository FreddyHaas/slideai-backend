FROM python:3.11-slim

# Set environment variables
ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONUNBUFFERED=1


WORKDIR /app


COPY requirements.txt .


RUN pip install --no-cache-dir --upgrade -r requirements.txt


COPY ./app .

CMD ["fastapi", "run", "main.py", "--port", "8080"]