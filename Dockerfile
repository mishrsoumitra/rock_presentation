# syntax=docker/dockerfile:1
# The editor and the server in one image. Node runs the TypeScript sources as they are; the build stage only bundles
# the editor (apps/web/dist) and the page Chromium draws slides with (apps/server/dist/render-page).
#
#   docker build -t rock-presentation .
#   docker run -p 3000:3000 -v pptx-data:/data -e API_TOKENS="$(openssl rand -hex 32)" rock-presentation
#
# The editor is then at http://localhost:3000/ and the API under /api (README.md, Deploying).

ARG NODE_IMAGE=node:24.14.0-bookworm-slim@sha256:d8e448a56fc63242f70026718378bd4b00f8c82e78d20eefb199224a4d8e33d8

FROM ${NODE_IMAGE} AS build
WORKDIR /app
ENV COREPACK_ENABLE_DOWNLOAD_PROMPT=0
RUN corepack enable
COPY . .
RUN pnpm install --frozen-lockfile \
 && pnpm --dir apps/web build \
 && pnpm --filter @pptx/server build:render-page

FROM ${NODE_IMAGE}
LABEL org.opencontainers.image.source="https://github.com/mishrsoumitra/rock_presentation" \
      org.opencontainers.image.description="A PowerPoint editor for the browser, with a server for slide images and PDFs"
WORKDIR /app
ENV COREPACK_ENABLE_DOWNLOAD_PROMPT=0 PLAYWRIGHT_BROWSERS_PATH=/ms-playwright
# tini reaps Chromium's child processes and passes signals on, so the server shuts down cleanly
RUN corepack enable \
 && apt-get update && apt-get install -y --no-install-recommends tini \
 && rm -rf /var/lib/apt/lists/*
COPY . .
# the server and the workspace packages it runs, without development dependencies; then Chromium's headless shell
RUN pnpm install --frozen-lockfile --prod --filter "@pptx/server..." \
 && pnpm --filter @pptx/server exec playwright install --with-deps --only-shell chromium \
 && rm -rf /var/lib/apt/lists/* /root/.cache /root/.local/share/pnpm
COPY --from=build /app/apps/web/dist apps/web/dist
COPY --from=build /app/apps/server/dist apps/server/dist
RUN mkdir -p /data && chown node:node /data
USER node
# Chromium's own sandbox needs user namespaces, which containers usually withhold: the container is the sandbox
# (run it read-only, without capabilities; see README.md). The API refuses to start without API_TOKENS unless
# ALLOW_ANONYMOUS_API=1 says something in front of it authenticates.
ENV NODE_ENV=production HOST=0.0.0.0 PORT=3000 DATA_DIR=/data RENDER_SANDBOX=0
EXPOSE 3000
VOLUME ["/data"]
HEALTHCHECK --interval=30s --timeout=5s --start-period=30s --retries=3 \
  CMD ["node", "-e", "fetch('http://127.0.0.1:' + (process.env.PORT || 3000) + '/api/health').then(r => process.exit(r.ok ? 0 : 1), () => process.exit(1))"]
ENTRYPOINT ["tini", "--"]
CMD ["node", "apps/server/src/main.ts"]
