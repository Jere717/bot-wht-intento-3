FROM node:18-slim

# Instalar dependencias necesarias para Puppeteer/Chromium
RUN apt-get update && apt-get install -y \
  wget \
  ca-certificates \
  fonts-liberation \
  libappindicator3-1 \
  libasound2 \
  libatk-bridge2.0-0 \
  libatk1.0-0 \
  libcups2 \
  libdbus-1-3 \
  libgdk-pixbuf2.0-0 \
  libnspr4 \
  libnss3 \
  libx11-xcb1 \
  libxcomposite1 \
  libxdamage1 \
  libxrandr2 \
  xdg-utils \
  libdrm2 \
  libgbm1 \
  libxshmfence1 \
  libgl1 \
  libpango-1.0-0 \
  libpangocairo-1.0-0 \
  libatspi2.0-0 \
  libwayland-client0 \
  libwayland-cursor0 \
  libwayland-egl1 \
  libxkbcommon0 \
  --no-install-recommends \
  && apt-get clean && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY package*.json ./
RUN npm install

COPY . .

RUN mkdir -p /app/sessions

EXPOSE 3000

CMD ["node", "server.js"]