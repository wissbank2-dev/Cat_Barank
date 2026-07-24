# Use the official Playwright image containing Node.js and all browser dependencies
FROM mcr.microsoft.com/playwright:v1.40.0-jammy

# Set the working directory
WORKDIR /app

# Copy package files
COPY package*.json ./

# Install dependencies
RUN npm install

# Copy application source code
COPY . .

# Expose the application port
EXPOSE 3000

# Start the application
CMD ["node", "index.js"]
