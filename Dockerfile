FROM python:3.9-slim

# Update and install LibreOffice without sudo (the container runs as root)
RUN apt-get update && apt-get install -y libreoffice && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY . /app

# Install Python dependencies
RUN pip install --no-cache-dir -r requirements.txt

# Use the default command to start your app
CMD ["python", "app.py"]
