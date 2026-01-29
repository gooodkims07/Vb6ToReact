# Backend Dockerfile
FROM node:18-slim

# Oracle Instant Client가 필요할 수 있으나, modern node-oracledb(v6+)는 Thin mode를 지원하므로 기본 라이브러리만 설치
RUN apt-get update && apt-get install -y libaio1 && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY package*.json ./
RUN npm install --production

COPY . .

EXPOSE 3000

CMD ["node", "server.js"]
