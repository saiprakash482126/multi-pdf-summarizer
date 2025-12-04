# Dockerfile
FROM python:3.11-slim

WORKDIR /app

# Install system dependencies
RUN apt-get update && apt-get install -y --no-install-recommends \
    gcc \
    python3-dev \
    && rm -rf /var/lib/apt/lists/*

# Install Python dependencies
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt uvicorn

# Copy application code
COPY . .

# Set Python path
ENV PYTHONPATH=/app

# Expose the port the app runs on
EXPOSE 8080

# Command to run the application
CMD ["python", "api.py"]