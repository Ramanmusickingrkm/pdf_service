FROM python:3.11-slim

WORKDIR /app

# Copy requirements first (for better caching)
COPY requirements.txt .

# Install dependencies
RUN pip install --no-cache-dir -r requirements.txt

# Copy application code
COPY app.py .

# Expose port
EXPOSE 5001

# Run the application
CMD ["gunicorn", "app:app", "--bind", "0.0.0.0:5001", "--timeout", "120"]
