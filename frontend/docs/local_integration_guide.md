# Local Frontend & Backend Integration Guide

This guide documents the step-by-step procedure to run and verify the **Schedule-Maker** FastAPI backend alongside the Flutter frontend for local development and testing.

---

## 1. Prerequisites

- **Python & uv**: Installed with project dependencies synced in `backend/`.
- **Flutter SDK**: 3.x with Windows, Web, or Android toolchains configured.

---

## 2. Backend Setup & Startup

### Step 2.1: Bootstrap a Super-Admin Account
From the `backend/` directory:

```powershell
uv run --no-sync python -m scripts.bootstrap_super_admin `
    --email "admin@hospital.org" `
    --full-name "System Administrator"
```

The script will prompt for a secure password twice and insert the record into `hospital_schedule.db`.

### Step 2.2: Start the FastAPI API Server
From `backend/`:

```powershell
uv run --no-sync uvicorn src.main:app --reload
```

The API will start listening on `http://127.0.0.1:8000`. You can confirm health at:
- `http://127.0.0.1:8000/docs` (Swagger UI)
- `http://127.0.0.1:8000/health` (Liveness probe)

---

## 3. Running the Flutter Frontend

All authentication routes are mounted under the `/api/v1` prefix (`/api/v1/auth/login`, `/api/v1/auth/signup`, etc.).

### Option A: Windows Desktop (Recommended for Local Dev)
Run directly from `frontend/`:

```powershell
flutter run -d windows
```
*(Uses the default `http://127.0.0.1:8000` configured in `lib/core/config/app_config.dart`)*.

### Option B: Chrome / Web
```powershell
flutter run -d chrome --web-port=3000
```
> [!NOTE]
> FastAPI is configured with CORS for `http://localhost:3000` and `http://127.0.0.1:3000`. Running on port 3000 ensures browser CORS requests succeed.

### Option C: Android Emulator
Because the Android emulator runs inside a virtual network, `127.0.0.1` refers to the emulator itself. To connect to the host PC's FastAPI server, pass `10.0.2.2`:

```powershell
flutter run -d emulator-5554 --dart-define=API_BASE_URL=http://10.0.2.2:8000
```

---

## 4. End-to-End Verification Scenarios

| Scenario | Action | Expected Result |
| :--- | :--- | :--- |
| **1. Cold Launch** | App starts without stored tokens | `AuthGate` briefly checks storage and displays `LoginScreen`. |
| **2. Super Admin Login** | Enter `admin@hospital.org` and password | Button shows inline spinner, receives JWT tokens, and `AuthGate` displays dashboard (`Authenticated: System Administrator`). |
| **3. Session Persistence** | Hot restart (`R`) or close and relaunch app | App automatically reads saved token from secure storage and restores session directly to the dashboard without prompting to log in. |
| **4. Logout** | Tap **Log Out** button or icon | Tokens are wiped from secure storage and app dynamically returns to `LoginScreen`. |
| **5. Navigation Toggle** | Tap "Complete Onboarding Registration" | Smoothly switches between `LoginScreen` and `SignupScreen`. |
