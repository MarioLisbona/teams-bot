# Base stage for shared dependencies
FROM node:20-slim as base
WORKDIR /usr/src/app

# Install nodemon globally in the base image
RUN npm install -g nodemon

# Copy package files first for better caching
COPY package*.json ./
RUN npm install

# Copy the rest of the application
COPY . .

# Development stage
FROM base as development
ENV NODE_ENV=development
EXPOSE 3978
# Use global nodemon
CMD ["nodemon", "--inspect=0.0.0.0:9229", "app.js"]

# Production stage
FROM base as production
ENV NODE_ENV=production
EXPOSE 3978
CMD ["npm", "start"] 