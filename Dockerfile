# Use official Python image with slim variant
FROM python:3.11-slim

# Set working directory
WORKDIR /app

# Install minimal system dependencies (including font support)
RUN apt-get update && apt-get install -y \
    gcc \
    g++ \
    libgomp1 \
    libglib2.0-0 \
    wget \
    chromium \
    chromium-driver \
    curl \
    fontconfig \
    fonts-liberation \
    fonts-dejavu \
    && rm -rf /var/lib/apt/lists/*

# Create fonts directory in system font location
RUN mkdir -p /usr/share/fonts/custom

# Copy custom fonts from your project's fonts directory
# The trailing slash on destination indicates it's a directory
COPY fonts/ /usr/share/fonts/custom/

# Update font cache so system recognizes the new fonts
RUN fc-cache -fv

# Copy requirements first (for better caching)
COPY requirements.txt .

# Install Python dependencies
RUN pip install --no-cache-dir -r requirements.txt

# Install kaleido for chart export
RUN pip install --no-cache-dir kaleido

# Copy the rest of the application
COPY . .

# Create necessary directories for data persistence
RUN mkdir -p /app/data/raw /app/data/processed /app/reports /app/logs

# Expose Streamlit port
EXPOSE 8501

# Run the application
CMD ["streamlit", "run", "src/app.py", "--server.port=8501", "--server.address=0.0.0.0"]