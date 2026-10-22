# Build context: repo root. Single-container image: nginx (static tab export +
# /api reverse proxy, port 8080) and the Azure Functions host (port 80) in one
# container, so the Container App only needs the minimum 0.25 vCPU / 0.5Gi
# instead of one allocation per container. web/Dockerfile and api/Dockerfile
# (the two-container sidecar layout, built by the GitHub workflow) still
# build the same pieces.
FROM node:22-alpine AS web
WORKDIR /src
COPY web/package*.json ./
RUN npm ci
COPY web/ .
# Same build-time constant as web/Dockerfile.
ENV NEXT_PUBLIC_API_BASE_URL=/api
RUN npm run build

FROM node:22-alpine AS api
WORKDIR /src
COPY api/package*.json ./
RUN npm ci
COPY api/tsconfig.json api/host.json ./
COPY api/src ./src
COPY api/assets ./assets
RUN npm run build
RUN npm prune --omit=dev

FROM mcr.microsoft.com/azure-functions/node:4-node22
# Debian's nginx package ships a default site listening on port 80, which is
# the Functions host's port here — remove it so only web/nginx.conf (8080) loads.
RUN apt-get update \
    && apt-get install -y --no-install-recommends nginx \
    && rm -rf /var/lib/apt/lists/* /etc/nginx/sites-enabled/default
ENV AzureWebJobsScriptRoot=/home/site/wwwroot \
    AzureFunctionsJobHost__Logging__Console__IsEnabled=true
COPY web/nginx.conf /etc/nginx/conf.d/default.conf
COPY --from=web /src/out /usr/share/nginx/html
COPY --from=api /src/host.json /home/site/wwwroot/host.json
COPY --from=api /src/package.json /home/site/wwwroot/package.json
COPY --from=api /src/dist /home/site/wwwroot/dist
COPY --from=api /src/assets /home/site/wwwroot/assets
COPY --from=api /src/node_modules /home/site/wwwroot/node_modules
EXPOSE 8080
# nginx daemonizes itself; the Functions host stays the foreground process, so
# the container's lifetime follows the Functions host (the base image's own CMD).
CMD ["sh", "-c", "nginx && exec /opt/startup/start_nonappservice.sh"]
