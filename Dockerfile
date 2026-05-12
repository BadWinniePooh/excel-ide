# ── Stage 1: Build ──────────────────────────────────────────────────────────
# Compiles JSX → plain JS and vendors React production UMD bundles so the
# runtime image needs no CDN and no client-side transpiler (removes the
# need for unsafe-eval in CSP).
FROM node:22-alpine AS builder
WORKDIR /build

COPY package.json babel.config.json ./
RUN npm install

COPY project/ ./project/

# Compile JSX to plain JS
RUN npx babel project/app.jsx         -o project/app.js && \
    npx babel project/tweaks-panel.jsx -o project/tweaks-panel.js

# Vendor React 18 production UMD builds (served locally, no CDN dependency)
RUN mkdir -p project/vendor && \
    cp node_modules/react/umd/react.production.min.js \
       project/vendor/react.production.min.js && \
    cp node_modules/react-dom/umd/react-dom.production.min.js \
       project/vendor/react-dom.production.min.js

# ── Stage 2: Runtime ────────────────────────────────────────────────────────
FROM nginx:1.27-alpine AS runtime

# Create a dedicated non-root user; nginx master still starts as root to bind
# the socket, then drops privileges to this user for worker processes.
RUN addgroup -g 1001 -S appgroup && \
    adduser  -u 1001 -S appuser -G appgroup

# Custom nginx configuration (replaces both main conf and default server block)
COPY nginx/nginx.conf          /etc/nginx/nginx.conf
COPY nginx/default.conf        /etc/nginx/conf.d/default.conf

# Copy only the compiled/vendored production assets — .jsx source files
# are intentionally excluded from the runtime image.
COPY --from=builder /build/project/index.html        /usr/share/nginx/html/
COPY --from=builder /build/project/formula.js        /usr/share/nginx/html/
COPY --from=builder /build/project/app.js            /usr/share/nginx/html/
COPY --from=builder /build/project/vendor/           /usr/share/nginx/html/vendor/

# Grant the non-root user ownership of the static asset directory.
# /tmp is already world-writable (1777) so the PID file and all nginx temp
# paths (client_body, proxy, …) can be written there without extra chowns.
# Logs go to /dev/stdout and /dev/stderr — no /var/log/nginx writes needed.
RUN chown -R appuser:appgroup /usr/share/nginx/html

USER appuser
EXPOSE 8080
CMD ["nginx", "-g", "daemon off;"]
