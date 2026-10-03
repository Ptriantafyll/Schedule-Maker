# ==========================================
# Stage 1: Build the Flutter Web Application
# ==========================================
FROM debian:bookworm-slim AS builder

# Install required build tools for Flutter
RUN apt-get update && apt-get install -y --no-install-recommends \
    curl \
    git \
    unzip \
    ca-certificates \
    && rm -rf /var/lib/apt/lists/*

# Clone the official Flutter SDK pinned to exact version 3.47.5
RUN git clone --depth 1 --branch 3.47.5 https://github.com/flutter/flutter.git /usr/local/flutter
ENV PATH="/usr/local/flutter/bin:/usr/local/flutter/bin/cache/dart-sdk/bin:${PATH}"

# Disable analytics and precache web engine artifacts
RUN flutter config --no-analytics && \
    flutter precache --web

WORKDIR /app

# Step 1: Copy dependency lockfiles first for optimal caching
COPY pubspec.yaml pubspec.lock ./
RUN flutter pub get

# Step 2: Copy remaining source code
COPY . .

# Step 3: Build the web application in release mode (dynamic origin detection)
RUN flutter build web --release

# ==========================================
# Stage 2: Production Nginx Runtime
# ==========================================
FROM nginx:1.27.4-alpine

# Copy compiled static assets from builder stage
COPY --from=builder /app/build/web /usr/share/nginx/html

# Copy internal Nginx configuration for SPA routing
COPY nginx.conf /etc/nginx/conf.d/default.conf

# Expose internal port
EXPOSE 80

# Start Nginx in foreground
CMD ["nginx", "-g", "daemon off;"]
