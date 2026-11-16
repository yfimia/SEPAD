# SEPAD (1998) sobre el runtime moderno de ASP clásico en Node.js
FROM node:20-bookworm-slim
# xsltproc: transforma los prólogos XML con la hoja XSL original (WD-xsl de IE5)
RUN apt-get update && apt-get install -y --no-install-recommends xsltproc && rm -rf /var/lib/apt/lists/*
WORKDIR /app
COPY package.json package-lock.json ./
RUN npm ci --omit=dev
COPY . .
RUN npm run setup
ENV PORT=8080 \
    SEPAD_RESET_ON_START=1 \
    SEPAD_RESET_HOURS=6 \
    SEPAD_LOG=0
EXPOSE 8080
CMD ["node", "server/index.js"]
