#!/bin/bash

# Get the container ID for dev environment specifically
CONTAINER_ID=$(docker ps | grep teams-bot-dev | head -n 1 | cut -d' ' -f1)

if [ -z "$CONTAINER_ID" ]; then
    echo "Development container not found!"
    exit 1
fi

echo "Found container: $CONTAINER_ID"
echo "Sending restart signal to nodemon..."

# Send rs\n directly to the container's stdin
docker exec $CONTAINER_ID sh -c 'echo "rs" > /proc/1/fd/0'

# Alternative method using touch
docker exec $CONTAINER_ID touch /usr/src/app/app.js

echo "Restart signal sent!" 