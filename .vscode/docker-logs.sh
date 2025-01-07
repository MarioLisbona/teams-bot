#!/bin/bash

# Default to development if no environment specified
ENV=${1:-dev}

if [ "$ENV" = "prod" ]; then
    echo "Viewing production logs..."
    docker-compose -f docker-compose.yml logs -f teams-bot-prod
elif [ "$ENV" = "dev" ]; then
    echo "Viewing development logs..."
    docker-compose -f docker-compose.dev.yml logs -f teams-bot-dev
else
    echo "Invalid environment. Use 'dev' or 'prod'"
    exit 1
fi