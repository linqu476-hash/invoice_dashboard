# For hosts with a persistent disk (a VPS, Render, Railway, Fly.io ...). Mount the disk at /data.
FROM python:3.14-slim
ENV PYTHONDONTWRITEBYTECODE=1 PYTHONUNBUFFERED=1 PIP_NO_CACHE_DIR=1
WORKDIR /app
COPY requirements.txt .
RUN pip install -r requirements.txt
# secrets, databases and logs are excluded by .dockerignore: pass them at run time, never bake them into the image
COPY . .
# run as a normal user, not root: a bug in the app cannot touch the rest of the container
RUN useradd --system --uid 10001 --no-create-home app && mkdir /data && chown app:app /data
USER app
ENV INVOICE_DB_PATH=/data/invoice.sqlite3
VOLUME /data
EXPOSE 8501
HEALTHCHECK --interval=30s --timeout=5s --retries=3 CMD python -c "import urllib.request,sys; sys.exit(0 if urllib.request.urlopen('http://127.0.0.1:8501/_stcore/health', timeout=4).status == 200 else 1)"
# APP_PASSWORD_HASH (or APP_PASSWORD) must be supplied as an environment variable (or a mounted .streamlit/secrets.toml)
CMD ["streamlit", "run", "streamlit_app.py", "--server.port=8501", "--server.address=0.0.0.0"]
