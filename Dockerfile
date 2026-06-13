# EULER · Claude Partner Network — Deal Desk
# Zero-dependency: Python standard library only.
FROM python:3.12-slim

WORKDIR /app
COPY . /app

# Render / Railway / Fly inject $PORT; default to 8080 locally.
ENV PORT=8080
# Persist SQLite on a mounted disk in production: set DB_PATH=/data/dealdesk.db
ENV DB_PATH=/app/dealdesk.db
EXPOSE 8080

CMD ["python", "server.py"]
