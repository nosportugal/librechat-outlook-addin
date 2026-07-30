# -- Build stage --
FROM node:24-alpine AS build

WORKDIR /app

COPY package.json package-lock.json* ./
RUN npm ci --ignore-scripts

COPY . .

RUN npm run build

# -- Runtime stage --
FROM node:24-alpine

WORKDIR /app

COPY --from=build /app/dist ./dist
COPY server.js ./

RUN chown -R node:node /app

USER node

EXPOSE 3000

CMD ["node", "server.js"]
