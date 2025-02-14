FROM python:3.9-slim

# Install system dependencies: LibreOffice and libraries needed by lxml
RUN apt-get update && apt-get install -y \
    libreoffice \
    libxml2-dev \
    libxslt-dev \
    zlib1g-dev \
    && rm -rf /var/lib/apt/lists/*

# Set working directory
WORKDIR /app

# Copy the requirements file first for caching
COPY requirements.txt /app/
RUN pip install --no-cache-dir -r requirements.txt

# Copy the rest of the application code into the container
COPY . /app
# Use the default command to start your Streamlit app
CMD ["streamlit", "run", "app.py", "--server.port=8501", "--server.enableCORS=false"]
