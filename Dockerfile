# Base image
FROM python:3.10-windowsservercore-1809

# Set the working directory
WORKDIR /app

# Copy the requirements file
COPY requirements.txt .

# Install the Python dependencies
RUN pip install --no-cache-dir -r requirements.txt

# Copy the application code
COPY . .

# Expose the required port
EXPOSE 8000

# Run the web app
CMD ["python", "app.py"]
