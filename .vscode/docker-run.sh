#!/bin/bash

# Default to development if no environment specified
ENV=${1:-development}

if [ "$ENV" = "production" ]; then
    echo "Starting production environment..."
    docker-compose -f docker-compose.yml up -d
elif [ "$ENV" = "development" ]; then
    echo "Starting development environment..."
    docker-compose -f docker-compose.dev.yml up -d
else
    echo "Invalid environment. Use 'development' or 'production'"
    exit 1
fi 