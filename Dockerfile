FROM node:20-bookworm

# Install LibreOffice and Poppler
RUN apt-get update && \
    apt-get install -y --no-install-recommends \
    libreoffice \
    libreoffice-impress \
    libreoffice-writer \
    poppler-utils \
    fonts-dejavu \
    fonts-liberation \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Copy package files first for better Docker caching
COPY package*.json ./

RUN npm install

# Copy the complete bot
COPY . .

# Start ST-BOT
CMD ["npm", "start"]
