FROM node:20-slim

WORKDIR /app

# Install dependencies first for efficient docker caching
COPY package*.json ./
RUN npm ci --omit=dev || npm install --omit=dev

# Copy app source
COPY . .

ENV PORT=3000
EXPOSE 3000

CMD ["node", "server.js"]
