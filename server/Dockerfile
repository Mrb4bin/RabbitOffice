# Rabbit Office Server in Docker. Build from the project root:
#   docker build -f server/Dockerfile -t rabbit-office .
FROM node:20-alpine
RUN apk add --no-cache tar openssl
WORKDIR /opt/rabbit-office
COPY server/package.json server/package.json
RUN cd server && npm install --omit=dev --no-audit --no-fund
COPY server server
COPY shared shared
COPY web web
ENV RABBIT_DATA=/data
VOLUME /data
EXPOSE 8080 80 443
HEALTHCHECK --interval=30s --timeout=5s CMD wget -qO- http://127.0.0.1:${PORT:-8080}/api/health || exit 1
CMD ["node", "server/src/server.js"]
