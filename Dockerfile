# For hosts with a persistent disk (a VPS, Render, Railway, Fly.io ...). Mount the disk at /data.
FROM python:3.12-slim
WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY . .
ENV INVOICE_DB_PATH=/data/invoice.sqlite3
VOLUME /data
EXPOSE 8501
# APP_PASSWORD must be supplied as an environment variable (or a mounted .streamlit/secrets.toml)
CMD ["streamlit", "run", "streamlit_app.py", "--server.port=8501", "--server.address=0.0.0.0"]
