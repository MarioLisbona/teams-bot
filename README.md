# Teams Bot Application

This repository contains a Teams Bot application that can be run in both development and production environments using Docker containers.

## Prerequisites

- Docker and Docker Compose installed
- Node.js 20.x (for local development without Docker)
- VS Code with Terminals Manager extension installed (for enhanced development environment)

## Environment Configuration

The application uses different environment files for development and production:

- `.env.local` - Development environment variables
- `.env.production` - Production environment variables

### Environment Variables

Both environment files contain configuration for:

- Microsoft Tenant ID
- Microsoft Teams Bot credentials
- Microsoft Excel Access App credentials
- SharePoint integration settings
- Sharepoint Root Directory Name and Template ID
- Azure OpenAI configuration
- Server configuration

`.env.sample` contains a description for each environment variable

## Docker Configuration

### Dockerfile

The project uses a multi-stage Dockerfile:

1. `base` stage - Common Node.js setup and dependencies
2. `development` stage - Includes nodemon for hot-reloading
3. `production` stage - Optimized for production deployment

### Docker Compose Files

#### Development (`docker-compose.dev.yml`)

yaml services:

teams-bot-dev:

- Runs in development mode
- Mounts source code for hot-reloading
- Uses .env.local
- Exposes port 3978

#### Production (`docker-compose.yml`)

yaml
services:

teams-bot-prod:

- Runs in production mode
- Uses .env.production
- Exposes port 3978
- Optimized for production deployment

## Running the Application

### Development Mode

- Start development container - `./.vscode/docker-run.sh dev`
- View server logs - `./.vscode/docker-logs.sh dev`
- Stop containers - `./.vscode/docker-down.sh dev`

### Production Mode

- Start production container - `./.vscode/docker-run.sh prod`
- View server logs - `./.vscode/docker-logs.sh prod`
- Stop containers - `./.vscode/docker-down.sh prod`

## VS Code Terminals Manager Integration

The project includes VS Code configuration for an enhanced development experience.

### Install the Terminals Manager extension in VS Code

Extension ID: `fabiospampinato.vscode-terminals`

### Run the development environment with Terminals Manager

Open the VS Code Command Palette

- On Windows/Linux: Ctrl + Shift + P
- On macOS: Cmd + Shift + P
- `Terminals: Run` to open the terminal manager.

### Integrated Terminals (`terminals.json`)

The `.vscode/terminals.json` configures automatic terminal setup:

1. **Development Server Terminal**

   - Automatically runs the development container
   - Uses `./.vscode/docker-run.sh dev`

2. **Server Logs Terminal**

   - Shows real-time container logs
   - Uses `./.vscode/docker-logs.sh dev`
   - Starts after a 5-second delay to ensure container is running

3. **Additional Terminals**
   - Two additional zsh terminals for general use

### Development Tools

#### Hot Reloading

- Development environment uses `nodemon` for automatic reloading
- Manual reload available using `./.vscode/docker-rs.sh`

#### Logging

- View logs in real-time using the docker-logs.sh script

## Scripts Reference

### `docker-rs.sh`

- Restarts the development container's nodemon process

### `docker-run.sh`

- Starts containers based on environment argument
- Defaults to development if no argument provided
- Uses appropriate docker-compose file
- Usage: `./docker-run.sh dev` or `./docker-run.sh prod`

### `docker-logs.sh`

- Displays container logs in real-time
- Supports both development and production environments
- Defaults to development if no argument provided
- Usage: `./docker-logs.sh dev` or `./docker-logs.sh prod`

### `docker-down.sh`

- Stops and removes containers
- Supports both development and production environments
- Defaults to development if no argument provided
- Usage: `./docker-down.sh dev` or `./docker-down.sh prod`
