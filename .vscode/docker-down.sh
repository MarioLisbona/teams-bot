#!/bin/bash

# Default to development if no environment specified
ENV=${1:-dev}

if [ "$ENV" = "prod" ]; then
    echo "Stopping production environment..."
    docker-compose -f docker-compose.yml down
elif [ "$ENV" = "dev" ]; then
    echo "Stopping development environment..."
    docker-compose -f docker-compose.dev.yml down
else
    echo "Invalid environment. Use 'dev' or 'prod'"
    exit 1
fi 