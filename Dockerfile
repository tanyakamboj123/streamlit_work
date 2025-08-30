# ---- Base image ----
FROM python:3.11-slim AS base

# Prevent Python from writing .pyc files and enable buffered stdout
ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1

# Set a working directory
WORKDIR /app

# System deps (curl for healthcheck; clean up apt cache)
RUN apt-get update \
 && apt-get install -y --no-install-recommends curl \
 && rm -rf /var/lib/apt/lists/*

# Copy app code first (adjust if your file is named differently)
# app.py is your Streamlit app from our previous step
COPY app.py /app/app.py

# If you have a requirements.txt, copy it, else we’ll install inline
# COPY requirements.txt /app/requirements.txt

# Install Python deps (pin or expand as needed)
RUN pip install --no-cache-dir --upgrade pip \
    && pip install --no-cache-dir \
        streamlit \
        python-pptx \
        requests \
        pydantic

# Create a non-root user to run the app
RUN useradd -m -u 1000 streamlit
USER streamlit

# Expose Streamlit default port
EXPOSE 8501

# Healthcheck: Streamlit serves /healthz when running
HEALTHCHECK --interval=30s --timeout=5s --retries=3 \
  CMD curl -f http://127.0.0.1:8501/healthz || exit 1

# Default command. You can override port/address with env vars below.
# STREAMLIT_SERVER_PORT (default 8501)
# STREAMLIT_SERVER_ADDRESS (default 0.0.0.0)
ENV STREAMLIT_SERVER_PORT=8501 \
    STREAMLIT_SERVER_ADDRESS=0.0.0.0

CMD ["bash", "-lc", "streamlit run /app/app.py --server.address=${STREAMLIT_SERVER_ADDRESS} --server.port=${STREAMLIT_SERVER_PORT}"]
