# Docker Implementation Plan & Containerization Guide

This document outlines the containerization architecture, configuration, and step-by-step rollout plan for the **Schedule-Maker** application using **Docker** and **Docker Compose** with a dedicated **API Gateway** pattern.

---

## 1. Architecture Overview (3-Service Gateway Pattern)

```
                            [ Host / Browser ]
                                    |
                                    | Port 3000
                                    v
+-------------------------------------------------------------------------------+
| Docker Network (`schedule-net`)                                               |
|                                                                               |
|   +-----------------------------------------------------------------------+   |
|   | `gateway` container (nginx:alpine)                                    |   |
|   | - Sole public port exposed: 3000 (host: 3000)                         |   |
|   | - Routes `/` -> `frontend:80` (Flutter Web SPA)                       |   |
|   | - Routes `/api/` -> `backend:8000/api/` (FastAPI REST endpoints)      |   |
|   | - Routes `/docs` & `/openapi.json` -> `backend:8000/` (API docs)      |   |
|   +--------------------+----------------------------------+---------------+   |
|                        |                                  |                   |
|       Internal HTTP    |                 Internal HTTP    |                   |
|       (Private)        v                 (Private)        v                   |
|   +------------------------+             +--------------------------------+   |
|   | `frontend` container   |             | `backend` container            |   |
|   | - flutter web build    |             | - python:3.11-slim + uv        |   |
|   | - nginx internal (:80) |             | - uvicorn src.main:app (:8000) |   |
|   +------------------------+             +----------------+---------------+   |
|                                                           |                   |
+-----------------------------------------------------------|-------------------+
                                                            |
                                           Volume Mount     v
                                       +----------------------------------------+
                                       | `./backend/data` (Host Directory)      |
                                       | - `hospital_schedule.db` (Persistent)  |
                                       +----------------------------------------+
```

---

## 2. Key Technical Decisions

### 2.1 Gateway Service (`gateway/`)
1. **Single Point of Ingress**:
   - Only the `gateway` container binds to host port `3000`.
   - `frontend` and `backend` containers do not expose any ports to the host machine, keeping the application topology secure and decoupled.
2. **Reverse Proxy Configuration (`gateway/nginx.conf`)**:
   - Port: `3000`.
   - `location /`: Proxies requests to `http://frontend:80`.
   - `location /api/`: Proxies requests to `http://backend:8000/api/` with forwarded headers (`X-Real-IP`, `X-Forwarded-For`, `X-Forwarded-Proto`).
   - `location /docs`: Proxies OpenAPI documentation to `http://backend:8000/docs`.
   - `location /openapi.json`: Proxies OpenAPI specification schema to `http://backend:8000/openapi.json`.

### 2.2 Backend Service (`backend/`)
1. **Base Image**: `python:3.11-slim`.
2. **Package Manager (`uv`)**:
   - Copy the official `uv` binary from `ghcr.io/astral-sh/uv:latest`.
   - Exploit Docker layer caching: copy `pyproject.toml` and `uv.lock` first, run `uv sync --frozen --no-dev`.
   - Copy application code (`src/`).
3. **Database & Storage**:
   - SQLite initially with `DATABASE_URL=sqlite:////app/data/hospital_schedule.db`.
   - Host directory `./backend/data` is mounted to `/app/data` to ensure schedules and users persist across container rebuilds.
   - Future PostgreSQL support is recorded in [`backend/docs/backlog.md`](../backend/docs/backlog.md#bl-016-postgresql-service--driver-integration-for-docker-compose) under `BL-016`.
4. **Server & Entrypoint**:
   - Command: `uv run uvicorn src.main:app --host 0.0.0.0 --port 8000`.

### 2.3 Frontend Service (`frontend/`)
1. **Multi-Stage Build**:
   - **Stage 1 (`builder`)**: Uses `ghcr.io/cirruslabs/flutter:stable` to run `flutter pub get` and `flutter build web --release --dart-define=API_BASE_URL=""` (relative API base URL so calls go seamlessly through the gateway).
   - **Stage 2 (`runner`)**: Uses `nginx:alpine` to serve static assets on internal port `80`.
2. **SPA Routing**:
   - Internal `nginx.conf` handles client-side routing (`try_files $uri $uri/ /index.html`).

### 2.4 Docker Compose (`docker-compose.yml`)
- Located at the repository root.
- Coordinates service startup dependencies (`gateway` depends on `frontend` and `backend`).
- Connects all three containers via a shared bridge network (`schedule-net`).

---

## 3. Step-by-Step Implementation Roadmap

| Step | Action | Files Touched / Created | Status |
| :--- | :--- | :--- | :--- |
| **Step 1** | Plan & Architecture Documentation | [`docs/docker_implementation_plan.md`](docker_implementation_plan.md), [`backend/docs/backlog.md`](../backend/docs/backlog.md) | `[x] Completed` |
| **Step 2** | Backend Containerization | `backend/.dockerignore`, `backend/Dockerfile` | `[x] Completed` |
| **Step 3** | Frontend Containerization | `frontend/.dockerignore`, `frontend/nginx.conf`, `frontend/Dockerfile` | `[x] Completed` |
| **Step 4** | Gateway Service Setup | `gateway/nginx.conf`, `gateway/Dockerfile` | `[x] Completed` |
| **Step 5** | Docker Compose Orchestration | `docker-compose.yml` | `[x] Completed` |
| **Step 6** | Server Deployment & Verification | Docker run on remote server | `[x] Ready for Server Deployment` |

---

## 4. Verification & Testing Procedure

Once all Docker files are created:
1. **Build images**:
   ```bash
   docker compose build
   ```
2. **Start containers in detached mode**:
   ```bash
   docker compose up -d
   ```
3. **Verify running containers**:
   ```bash
   docker compose ps
   ```
4. **Health & Endpoint Check**:
   - Gateway Web Access: `http://localhost:3000` (renders Flutter app).
   - Gateway API Docs: `http://localhost:3000/docs` (renders FastAPI Swagger UI).
   - Gateway API Check: `http://localhost:3000/api/v1/auth/me` (proxied directly to FastAPI).
   - End-to-End: Log in through `http://localhost:3000` and generate a test schedule.
