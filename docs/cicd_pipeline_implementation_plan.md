# CI/CD Pipeline Implementation Plan: GHCR.io, Zero-Downtime Deployment & Hetzner ARM64 Guide

This document defines the architecture, workflow specifications, and step-by-step rollout plan for the **Schedule-Maker** automated CI/CD pipeline using **GitHub Actions**, **GitHub Container Registry (`ghcr.io`)**, and a **Self-Hosted Runner** on a **Hetzner Cloud ARM64 Virtual Machine**.

---

## 1. Environment & Target Infrastructure

Based on system inspection of the target server:
- **Cloud Provider**: Hetzner Cloud (vServer)
- **Hostname**: `board-game-ubuntu`
- **Operating System**: Ubuntu 24.04.3 LTS (Linux Kernel 6.8)
- **Architecture**: **`arm64` (aarch64)** *(Critical: runner package and base images must target Linux ARM64)*
- **Ingress Strategy**:
  - Current: `ngrok http 3000`
  - Future: Cloudflare custom domain with HTTPS reverse proxy
- **Container Registry**: **GitHub Container Registry (`ghcr.io`)**
  - Private, zero-maintenance OCI registry hosted directly on GitHub.
  - Future-proof: standard image repository for Kubernetes transitions (`imagePullSecrets`).
- **Container Architecture**: 3-Service Gateway Pattern (`gateway:3000`, `backend:8000`, `frontend:80`) orchestrated via `docker-compose.yml`.

---

## 2. Architecture & Pipeline Topology

```
+-----------------------------------------------------------------------------------------------+
| GitHub Cloud                                                                                  |
|                                                                                               |
|  Developer Push / PR to main                                                                  |
|           |                                                                                   |
|           +-----------------------------------+                                               |
|           |                                   |                                               |
|           v (Parallel)                        v (Parallel)                                    |
|  +-----------------------------+     +-----------------------------+                          |
|  | Job: Backend CI             |     | Job: Frontend CI            |                          |
|  | - ubuntu-latest (GitHub)    |     | - ubuntu-latest (GitHub)    |                          |
|  | - Python 3.11.9 + uv        |     | - Flutter 3.47.5            |                          |
|  | - uv run pytest (785 tests) |     | - flutter analyze & test    |                          |
|  +--------------+--------------+     +--------------+--------------+                          |
|                 |                                   |                                         |
|                 +-----------------+-----------------+                                         |
|                                   | All Tests Pass?                                           |
|                                   v                                                           |
|                    Quality Gate Approved (on push to main)                                    |
|                                   |                                                           |
+-----------------------------------|-----------------------------------------------------------+
                                    v (GitHub Actions Runner Webhook)
+-----------------------------------------------------------------------------------------------+
| Hetzner Cloud VM (`board-game-ubuntu` - Ubuntu 24.04 ARM64)                                   |
|                                                                                               |
|  +-----------------------------------------------------------------------------------------+  |
|  | GitHub Actions Self-Hosted Runner (systemd service)                                     |  |
|  |                                                                                         |  |
|  |  Step 1: actions/checkout (Fetch latest main branch)                                    |  |
|  |  Step 2: Log in to GHCR (echo $GITHUB_TOKEN | docker login ghcr.io)                     |  |
|  |  Step 3: Build & Tag Native ARM64 Images                                                |  |
|  |          - ghcr.io/<owner>/schedule-maker-gateway:latest & :sha                         |  |
|  |          - ghcr.io/<owner>/schedule-maker-backend:latest & :sha                         |  |
|  |          - ghcr.io/<owner>/schedule-maker-frontend:latest & :sha                        |  |
|  |  Step 4: Push Images to GHCR (Private Registry for VM & Future Kubernetes)             |  |
|  |  Step 5: Zero-Downtime Rollout                                                          |  |
|  |          docker compose up -d --remove-orphans                                          |  |
|  |  Step 6: Verify Local Health (curl http://localhost:3000/health)                        |  |
|  |  Step 7: Disk Hygiene (docker image prune -f)                                           |  |
|  +-----------------------------------------------------------------------------------------+  |
|                                                                                               |
|  +-----------------------------------------------------------------------------------------+  |
|  | Docker Host Runtime (schedule-net)                                                      |  |
|  |  - schedule_maker_gateway  (:3000 -> ngrok / Cloudflare)                                |  |
|  |  - schedule_maker_backend  (:8000 internal) -> data/hospital_schedule.db               |  |
|  |  - schedule_maker_frontend (:80 internal)                                               |  |
|  +-----------------------------------------------------------------------------------------+  |
+-----------------------------------------------------------------------------------------------+
                                    |
                                    v (Images Pushed & Stored)
+-----------------------------------------------------------------------------------------------+
| GitHub Container Registry (ghcr.io)                                                           |
|  - Private OCI images stored for version control, rollbacks, and future Kubernetes deployment |
+-----------------------------------------------------------------------------------------------+
```

---

## 3. Zero-Downtime Deployment Strategy

### Will `docker compose down` and then `up` cause downtime?
**Yes, absolutely.** 
Running `docker compose down`:
1. Immediately kills all active containers (`gateway`, `backend`, `frontend`).
2. Drops the internal bridge network.
3. Leaves a downtime window where every user request or browser connection receives `Connection Refused` or `502 Bad Gateway` until the new containers finish building and starting (often 30 seconds to several minutes).

### How We Achieve Zero-Downtime:
**Never execute `docker compose down` in production.**

Instead, we use in-place atomic recreation:
```bash
# 1. Build and tag the new images in the background while the old containers are STILL serving traffic
docker compose build

# 2. Recreate containers in-place with zero downtime
docker compose up -d --remove-orphans
```

**Why this achieves zero downtime:**
1. **Old containers keep running** and serving web traffic while the new Docker images build.
2. When `docker compose up -d` executes, Docker Compose starts the new containers and instantly cuts over the network endpoints in a fraction of a second without tearing down the network or volume mounts.
3. Persistent database files (`./backend/data/hospital_schedule.db`) remain mounted and uninterrupted.
4. When moving to **Kubernetes** later, this naturally evolves into native K8s `RollingUpdate` with zero-downtime readiness probes.

---

## 4. Key Technical Decisions & ARM64 Considerations

### 4.1 Compute Separation (CI on GitHub Cloud, CD on VM)
- **CI Jobs (`test-backend` & `test-frontend`)**: Run on GitHub-hosted `ubuntu-latest`.
  - Fast, parallel execution without consuming VM CPU/RAM.
  - Prevents unverified/broken code from ever touching the production VM.
- **CD Job (`deploy`)**: Runs on `runs-on: [self-hosted, linux, ARM64]`.
  - Triggers only when CI tests pass and the commit lands on `main`.

### 4.2 GitHub Container Registry (`ghcr.io`) Integration
- Images will be tagged with both `:latest` and the commit SHA `:${{ github.sha }}`.
- Pushed to `ghcr.io/<owner>/schedule-maker-gateway`, `schedule-maker-backend`, and `schedule-maker-frontend`.
- Authenticated automatically using `${{ secrets.GITHUB_TOKEN }}` (no external credentials required).
- Ready for immediate use by Kubernetes via standard `imagePullSecrets`.

### 4.3 Dedicated Non-Root Service User on VM
- GitHub Actions runner refuses to run as `root` by default.
- We will configure a dedicated user (`github-runner`) belonging to the `docker` group.
- The runner runs as a background `systemd` daemon (`svc.sh`), surviving VM restarts.

### 4.4 Gateway Health Route
- We add `location /health` to [`gateway/nginx.conf`](../gateway/nginx.conf) proxying to `http://backend:8000/health`.
- Allows the CD deployment step to verify that the entire gateway $\rightarrow$ backend stack is healthy via `curl http://localhost:3000/health`.

### 4.5 Disk Hygiene on the VM
- Multi-stage Flutter builds and Docker layer caches produce dangling intermediate images.
- An automated post-deployment step `docker image prune -f` reclaims disk space after every rollout.

---

## 5. Step-by-Step Implementation Roadmap

| Phase | Milestone | Actions | Target Files |
| :--- | :--- | :--- | :--- |
| **Phase 1** | **Gateway Ingress Enhancement** | Expose `/health` route through Nginx gateway for automated smoke tests | [`gateway/nginx.conf`](../gateway/nginx.conf) |
| **Phase 2** | **GitHub Actions Workflow** | Create CI/CD workflow defining backend tests, frontend tests, GHCR image push, and zero-downtime deployment | [`.github/workflows/ci-cd.yml`](../.github/workflows/ci-cd.yml) |
| **Phase 3** | **VM Setup Guide** | Comprehensive documentation for provisioning the ARM64 self-hosted runner and systemd daemon on Ubuntu 24.04 | [`docs/vm_self_hosted_runner_guide.md`](vm_self_hosted_runner_guide.md) |
| **Phase 4** | **Deployment & End-to-End Verification** | Test CI jobs on GitHub, connect self-hosted runner on Hetzner VM, and execute live deployment | Live Hetzner VM & GitHub Actions |

---

## 6. Verification Plan

### 6.1 Automated Quality Gates
1. **Backend Tests**: `uv run pytest` runs 785 unit and integration tests on `ubuntu-latest`.
2. **Frontend Tests**: `flutter analyze` + `flutter test` executes 255 widget tests on `ubuntu-latest`.

### 6.2 Deployment Smoke Test (on VM)
1. **Zero-Downtime Rollout**:
   - `docker compose up -d --remove-orphans` updates containers without dropping existing connections.
2. **Healthcheck Probe**:
   - Script polls `http://localhost:3000/health` up to 12 times (60s timeout) expecting `HTTP 200` with `{"status":"healthy"}`.
3. **Registry Verification**:
   - Verify that images appear under your GitHub profile/organization: `https://github.com/<owner>?tab=packages`.
4. **Data Persistence**:
   - Verify that SQLite database `./backend/data/hospital_schedule.db` retains user accounts and schedules across container restarts.
