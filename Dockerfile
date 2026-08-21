FROM node:22-alpine AS build

WORKDIR /app
RUN corepack enable

COPY package.json pnpm-lock.yaml pnpm-workspace.yaml ./
RUN pnpm install --frozen-lockfile

COPY src ./src
COPY tsconfig.json tsdown.config.ts ts-builds.config.json ./
RUN pnpm build

FROM node:22-alpine AS runtime

ENV NODE_ENV=production \
    MCP_HOST=0.0.0.0 \
    MCP_PORT=3001

WORKDIR /app
RUN corepack enable

COPY package.json pnpm-lock.yaml pnpm-workspace.yaml ./
RUN pnpm install --prod --frozen-lockfile
COPY --from=build /app/dist ./dist

USER node
EXPOSE 3001

CMD ["node", "dist/http-server.js"]
