FROM python:3.9-slim

# Update and install LibreOffice (runs as root)
# Install system dependencies: LibreOffice and packages needed by lxml
RUN apt-get update && apt-get install -y \
    libreoffice \
    libxml2-dev \
    libxslt1-dev \
    gcc \
    && rm -rf /var/lib/apt/lists/*

# Set working directory
WORKDIR /app

# Copy only the requirements first
COPY requirements.txt /app/

# Install Python dependencies
RUN pip install --no-cache-dir -r requirements.txt

# Copy the rest of the application code
COPY . /app
# For Linux deployment, set the LibreOffice path to the Linux executable.
ENV LIBREOFFICE_PATH=/usr/bin/libreoffice
CMD ["python", "app.py"]

