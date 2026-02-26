# Use Node.js 18
FROM node:18-alpine

# Install cifs-utils (needed if you want to mount inside container, 
# but usually better to mount on Proxmox host)
RUN apk add --no-cache tzdata

# Set working directory
WORKDIR /usr/src/app

# Copy package files and install dependencies
COPY package*.json ./
RUN npm install --production

# Copy the rest of your code
COPY . .

# Create the folder for reports if it doesn't exist
RUN mkdir -p ChicagoReport

# Start the bot
CMD [ "node", "index.js" ]