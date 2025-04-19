#!/bin/bash

set -e

echo "Starting deployment process for AutomaterOperating Backend..."

if ! command -v docker &> /dev/null; then
    echo "Docker is not installed. Please install Docker first."
    exit 1
fi

if ! command -v docker-compose &> /dev/null; then
    echo "Docker Compose is not installed. Please install Docker Compose first."
    exit 1
fi

if [ ! -f .env ]; then
    echo "Creating .env file..."
    echo "SECRET_KEY=$(openssl rand -hex 32)" > .env
    echo "ENVIRONMENT=production" >> .env
    echo "DB_PASSWORD=$(openssl rand -hex 8)" >> .env
    echo "DB_USER=postgres" >> .env
    echo "DB_NAME=automater" >> .env
    echo ".env file created with random secret key and database password."
else
    echo ".env file already exists, using existing configuration."
fi

echo "Building and starting Docker containers..."
docker-compose down
docker-compose build --no-cache
docker-compose up -d

sleep 5

if curl -s http://localhost:8000/api/health | grep -q "healthy"; then
    echo "Deployment successful! API is up and running."
    echo "API is available at: http://localhost:8000"
    echo "API documentation is available at: http://localhost:8000/docs"
else
    echo "Something went wrong. API health check failed."
    echo "Check the logs with: docker-compose logs api"
    exit 1
fi

echo "Deployment complete!"