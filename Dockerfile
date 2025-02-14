FROM python:3.9-slim

# Update and install LibreOffice (runs as root)
RUN apt-get update && \
    apt-get install -y libreoffice && \
    rm -rf /var/lib/apt/lists/*

# Set working directory
WORKDIR /app

# Copy only the requirements first
COPY requirements.txt /app/

# Install Python dependencies
RUN pip install --no-cache-dir -r requirements.txt

# Copy the rest of the application code
COPY . /app

# Use the default command to start your app
CMD ["python", "app.py"]
