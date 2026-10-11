# Hetzner ARM64 Ubuntu VM: Self-Hosted Runner Setup & Operations Guide

This guide provides step-by-step instructions for provisioning, configuring, and maintaining the **GitHub Actions Self-Hosted Runner** on your **Hetzner Cloud ARM64 Ubuntu 24.04 VM** (`board-game-ubuntu`).

---

## 1. System Requirements & Architecture

- **Operating System**: Ubuntu 24.04.3 LTS (Kernel 6.8+)
- **Architecture**: **`arm64` (aarch64)** *(Critical: must install the ARM64 runner build)*
- **Privilege Separation**: Dedicated service account `github-runner` with `docker` group membership.

---

## 2. Step-by-Step Installation

### Step 2.1: Verify Docker and Docker Compose
SSH into your Hetzner VM as `root`:

```bash
ssh root@<YOUR_VM_IP>
```

Verify that Docker and the Docker Compose plugin are installed:

```bash
docker --version
docker compose version
```

If Docker is not yet installed, install Docker Engine and the Compose plugin via the official Docker repository:

```bash
apt-get update
apt-get install -y ca-certificates curl gnupg
install -m 0755 -d /etc/apt/keyrings
curl -fsSL https://download.docker.com/linux/ubuntu/gpg | gpg --dearmor -o /etc/apt/keyrings/docker.gpg
chmod a+r /etc/apt/keyrings/docker.gpg

echo \
  "deb [arch=arm64 signed-by=/etc/apt/keyrings/docker.gpg] https://download.docker.com/linux/ubuntu \
  $(. /etc/os-release && echo "$VERSION_CODENAME") stable" | \
  tee /etc/apt/sources.list.d/docker.list > /dev/null

apt-get update
apt-get install -y docker-ce docker-ce-cli containerd.io docker-buildx-plugin docker-compose-plugin
systemctl enable --now docker
```

---

### Step 2.2: Create Dedicated Runner User
GitHub Actions runners should **never run directly as `root`** for security and stability.

Create a dedicated non-root user `github-runner` and grant it Docker access:

```bash
# 1. Create user with a home directory and bash shell
useradd -m -s /bin/bash github-runner

# 2. Add github-runner to the docker group
usermod -aG docker github-runner

# 3. Allow github-runner passwordless sudo for managing the systemd service (optional but recommended)
echo "github-runner ALL=(ALL) NOPASSWD: /home/github-runner/actions-runner/bin/installdependencies.sh" > /etc/sudoers.d/github-runner
```

Now switch to the `github-runner` account:

```bash
su - github-runner
```

Verify that Docker works without `sudo`:

```bash
docker ps
```
*(Should return table headers without permission errors).*

---

### Step 2.3: Generate Registration Token on GitHub

1. Open your repository on GitHub: `https://github.com/Ptriantafyll/Schedule-Maker`.
2. Navigate to: **Settings** $\rightarrow$ **Actions** $\rightarrow$ **Runners**.
3. Click the green button: **New self-hosted runner**.
4. Select:
   - **Runner image**: `Linux`
   - **Architecture**: `ARM64` *(Make sure to select ARM64!)*
5. GitHub will display a download URL and a registration token (`--token XXXXXXXX`). Keep this tab open.

---

### Step 2.4: Download and Configure the Runner
On the VM (still logged in as `github-runner`):

```bash
# 1. Create runner directory
mkdir -p ~/actions-runner && cd ~/actions-runner

# 2. Download the latest Linux ARM64 runner package
# (Check GitHub UI tab for the exact current version number, e.g. 2.322.0)
curl -o actions-runner-linux-arm64.tar.gz -L https://github.com/actions/runner/releases/download/v2.322.0/actions-runner-linux-arm64-2.322.0.tar.gz

# 3. Extract the archive
tar xzf ./actions-runner-linux-arm64.tar.gz

# 4. Configure the runner using the token from your GitHub UI
./config.sh --url https://github.com/Ptriantafyll/Schedule-Maker --token <YOUR_COPIED_TOKEN>
```

When prompted:
- **Runner group**: Press `Enter` (Default).
- **Runner name**: Press `Enter` (or enter `hetzner-arm64-vm`).
- **Additional labels**: Enter `self-hosted,linux,ARM64` (or press `Enter` for defaults).
- **Work folder**: Press `Enter` (Default `_work`).

---

### Step 2.5: Install and Start as a Systemd Service
To ensure the runner stays running permanently in the background and starts automatically on server boot:

```bash
# Return to root terminal (or run with sudo from root)
exit

# From root in the runner folder:
cd /home/github-runner/actions-runner
./svc.sh install github-runner
./svc.sh start
./svc.sh status
```

You should see:
```text
● actions.runner.Ptriantafyll-Schedule-Maker.hetzner-arm64-vm.service - GitHub Actions Runner
     Loaded: loaded (/etc/systemd/system/actions.runner...; enabled)
     Active: active (running)
```

---

## 3. GitHub Repository Permissions Setup

To allow your workflows to push container images to **GitHub Container Registry (`ghcr.io`)**:

1. In your GitHub repository:
   - Go to **Settings** $\rightarrow$ **Actions** $\rightarrow$ **General**.
2. Scroll down to **Workflow permissions**:
   - Select **Read and write permissions**.
   - Check **Allow GitHub Actions to create and approve pull requests** (optional).
   - Click **Save**.

---

## 4. Verification

### Step 4.1: Check Runner Status on GitHub
1. In your repository: **Settings** $\rightarrow$ **Actions** $\rightarrow$ **Runners**.
2. You should now see your runner `hetzner-arm64-vm` with a green dot and status **Idle**.

### Step 4.2: Trigger Your First Automated Deployment
1. Go to the **Actions** tab in your GitHub repository.
2. Under All workflows, select **CD**.
3. Click **Run workflow** $\rightarrow$ **Branch: main** $\rightarrow$ **Run workflow**.
4. Watch the pipeline:
   - It will pick up on your Hetzner VM runner.
   - Build native ARM64 Docker images.
   - Push them to `ghcr.io/ptriantafyll/schedule-maker-...`.
   - Swap containers in place with zero downtime.
   - Verify `http://localhost:3000/health`.
5. Check your GitHub Packages tab:
   - Go to `https://github.com/Ptriantafyll?tab=packages`.
   - You will see your private container images listed there!

---

## 5. Maintenance & Useful Commands

### Viewing Runner Service Logs on VM
```bash
journalctl -u actions.runner.* -f
```

### Stopping / Restarting Runner Service
```bash
cd /home/github-runner/actions-runner
sudo ./svc.sh stop
sudo ./svc.sh start
sudo ./svc.sh status
```

### Checking Running Containers on VM
```bash
docker ps
```

### Checking Container Logs
```bash
cd /home/github-runner/actions-runner/_work/Schedule-Maker/Schedule-Maker
docker compose logs -f gateway
docker compose logs -f backend
```
