FROM python:3.11-slim

# Install system dependencies (LibreOffice for PDF conversion)
RUN apt-get update && apt-get install -y \
    libreoffice \
    libreoffice-writer \
    libreoffice-common \
    libreoffice-java-common \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Copy requirements first (for better caching)
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy application code
COPY . .

# Expose port
EXPOSE 5001

# Start command
CMD ["gunicorn", "app:app", "--bind", "0.0.0.0:5001"]
