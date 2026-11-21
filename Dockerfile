FROM python:3.11-slim

WORKDIR /app

COPY app/requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Glob rather than naming each module: main.py imports its siblings as top-level
# modules, so one left out of this list builds a fine image that dies on the first
# request with ImportError.
COPY app/*.py ./
COPY app/templates ./templates
COPY app/static ./static

ENV STATIC_DIR=/app/static
ENV PORT=8080
EXPOSE 8080

CMD ["sh", "-c", "uvicorn main:app --host 0.0.0.0 --port ${PORT}"]
