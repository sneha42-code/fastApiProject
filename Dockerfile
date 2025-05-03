# Dockerfile for building the FastAPI application image

# Use the official Python image as the base image
FROM python:3.9-slim

# Set the working directory in the container
WORKDIR /app

# Copy the requirements file into the container
COPY requirements.txt ./


RUN mkdir -p /app/file_uploads
RUN mkdir -p /app/attrition_reports
RUN chmod 777 /app/file_uploads /app/attrition_reports  # Set permissions if needed


# Install the required Python packages
RUN pip install --no-cache-dir -r requirements.txt

# Copy the application code into the container
COPY . .

# Expose port 8000 for the FastAPI application
EXPOSE 8000

# Command to run the FastAPI application
CMD ["uvicorn", "src.main:app", "--host", "0.0.0.0", "--port", "8000"]
