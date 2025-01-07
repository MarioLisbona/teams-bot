#!/bin/bash

# Default to development if no environment specified
ENV=${1:-dev}

if [ "$ENV" = "prod" ]; then
    echo "Starting production environment..."
    docker-compose -f docker-compose.yml up -d
elif [ "$ENV" = "dev" ]; then
    echo "Starting development environment..."
    docker-compose -f docker-compose.dev.yml up -d
else
    echo "Invalid environment. Use 'dev' or 'prod'"
    exit 1
fi 