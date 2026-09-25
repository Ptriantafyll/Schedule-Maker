# Frontend & Client Authentication Integration Guide

Comprehensive integration reference for frontend and mobile engineers connecting client applications (Flutter Mobile, Flutter Web, SPAs) to the **Schedule-Maker** backend API. See the [frontend roadmap](frontend_implementation_roadmap.md) and [auth implementation plan](frontend_auth_implementation_plan.md) for the current build sequence.

---

## 1. Architectural Overview & Dual Transport Modes

The Schedule-Maker backend implements a **hybrid, dual-mode authentication transport architecture**:

1. **Native / Mobile Mode (Flutter Android, iOS, Desktop)**:
   - Uses standard OAuth2 Bearer access tokens in the `Authorization` header.
   - Refresh tokens are stored in platform-supported secure storage (Keystore/Keychain on mobile; suitable secure storage on Windows) and exchanged in JSON request/response bodies.
2. **Web / Browser Mode (Flutter Web, SPAs)**:
   - Uses browser-managed **`HttpOnly`, `SameSite=Lax` cookies** for refresh tokens, defending against token theft via Cross-Site Scripting (XSS).
   - Enforces **Double-Submit CSRF** defense using a readable `csrf_token` cookie and required `X-CSRF-Token` HTTP header on refresh and logout operations.

All API routes are prefixed under `/api/v1`.

---

## 2. Endpoint Security Map

| Endpoint | Method | Security | Transport Mode | Description |
| :--- | :--- | :--- | :--- | :--- |
| `/health` | `GET` | Public | Any | Server liveness & health check |
| `/api/v1/auth/login` | `POST` | Public (Rate Limited) | Any | Authenticate user credentials & issue tokens |
| `/api/v1/auth/signup` | `POST` | Public (Rate Limited) | Any | Consume an invitation token & activate account |
| `/api/v1/auth/refresh` | `POST` | Public (Rate Limited) | Body or Cookie+CSRF | Rotate refresh token & mint fresh access token |
| `/api/v1/auth/logout` | `POST` | Public | Body or Cookie+CSRF | Terminate active refresh session & clear cookies |
| `/api/v1/auth/me` | `GET` | Bearer Token Required | `Authorization: Bearer <token>` | Retrieve authenticated user profile |
| `/api/v1/*` (all domain routes) | Any | Bearer Token Required | `Authorization: Bearer <token>` | Departments, doctors, shifts, teams, etc. |

---

## 3. Mobile Client Flow (Flutter iOS / Android / Desktop)

### 3.1 Token Storage
Use hardware-backed secure storage (e.g. `flutter_secure_storage`):
- `access_token`: Short-lived (backend default 60 minutes; configurable). Can be kept in-memory or secure storage.
- `refresh_token`: Long-lived (14 days). **Must** be stored in hardware-backed secure storage.

### 3.2 Login
Submit credentials using standard `application/x-www-form-urlencoded` format:

```http
POST /api/v1/auth/login HTTP/1.1
Host: api.hospital.org
Content-Type: application/x-www-form-urlencoded

username=doctor@hospital.org&password=SecurePassword123!
```

**Successful Response (`200 OK`)**:
```json
{
  "access_token": "eyJhbGciOi...",
  "token_type": "bearer",
  "refresh_token": "uR9xK...",
  "csrf_token": "eF2bA..."
}
```

Store `access_token` and `refresh_token`.

### 3.3 Authenticated Requests
Attach the access token in the `Authorization` header for all domain requests:

```http
GET /api/v1/doctors/ HTTP/1.1
Host: api.hospital.org
Authorization: Bearer eyJhbGciOi...
```

### 3.4 Automatic Token Refresh (Dio / HTTP Interceptor)
When an access token expires, the server returns `401 Unauthorized` with `WWW-Authenticate: Bearer`.
Configure your HTTP client interceptor to:
1. Intercept the `401 Unauthorized` response.
2. Send `POST /api/v1/auth/refresh` with the stored `refresh_token`:

```http
POST /api/v1/auth/refresh HTTP/1.1
Host: api.hospital.org
Content-Type: application/json

{
  "refresh_token": "uR9xK..."
}
```

**Successful Refresh Response (`200 OK`)**:
```json
{
  "access_token": "eyJhbGciOi...NEW...",
  "token_type": "bearer",
  "refresh_token": "vS0yL...NEW...",
  "csrf_token": "fG3cB...NEW..."
}
```

3. **Important (RFC 6819 Rotation)**: Save the **new** `refresh_token`. The previous refresh token has been invalidated.
4. Retry the original failed request with the new `access_token`.
5. If the refresh request itself returns `401 Unauthorized`, the session has expired or was revoked. Purge stored tokens and redirect the user to the Login screen.

### 3.5 Logout
```http
POST /api/v1/auth/logout HTTP/1.1
Host: api.hospital.org
Content-Type: application/json

{
  "refresh_token": "vS0yL..."
}
```
Delete local tokens from secure storage upon completion.

---

## 4. Web Client Flow (Flutter Web / Browser SPA)

### 4.1 Browser Cookie Handling
When operating in a browser, refresh tokens should be stored in **`HttpOnly` cookies** rather than `localStorage` or `sessionStorage` to mitigate token exfiltration via XSS.

- Set `withCredentials = true` (or `credentials: 'include'`) on all HTTP requests.
- The browser automatically handles storing, sending, and clearing the `refresh_token` cookie.

### 4.2 CSRF Protection on Web Refresh & Logout
The backend **enforces** the double-submit CSRF pattern for cookie-based refresh. Send the CSRF header for cookie logout as well, but note that the current logout controller does **not** validate it; backend hardening is needed before treating logout as CSRF-protected.

1. When logging in or refreshing, the server sets:
   - `refresh_token`: `HttpOnly`, `SameSite=Lax` (cannot be read by JavaScript).
   - `csrf_token`: `SameSite=Lax` (readable by JavaScript).
   - Set `SECURE_COOKIE=true` and use HTTPS in production; the current backend defaults this setting to false for local development.
2. When calling `/api/v1/auth/refresh` or `/api/v1/auth/logout`:
   - Extract the `csrf_token` cookie value using standard JavaScript cookie parsing:
     ```javascript
     const csrfToken = document.cookie
       .split('; ')
       .find(row => row.startsWith('csrf_token='))
       ?.split('=')[1];
     ```
   - Send it in the `X-CSRF-Token` header:
     ```http
     POST /api/v1/auth/refresh HTTP/1.1
     Host: api.hospital.org
     X-CSRF-Token: <extracted_csrf_token>
     ```

### 4.3 CORS Configuration
- In development: Ensure the frontend origin (e.g. `http://localhost:3000`) is in the backend's allowed CORS origins list.
- Web requests must include `Authorization`, `Content-Type`, and `X-CSRF-Token` in allowed headers.

---

## 5. User Onboarding & Invitation Signup

New hospital staff and administrators are invited via secure, one-time invitation tokens.

### 5.1 Invitation Link Structure
Invited users receive an onboarding URL:
```text
https://app.hospital.org/signup?token=<raw_invitation_token>
```

### 5.2 Signup Request
```http
POST /api/v1/auth/signup HTTP/1.1
Host: api.hospital.org
Content-Type: application/json

{
  "invitation_token": "abc123rawTokenHere...",
  "first_name": "Gregory",
  "last_name": "House",
  "email": "ghouse@hospital.org",
  "password": "SecurePassword123!"
}
```

> [!NOTE]
> The signup schema **strictly forbids** passing `role`, `department_id`, or `doctor_id`. These attributes are derived securely on the backend from the cryptographically verified invitation token.

**Successful signup returns `201 Created` with a `UserRead` profile (`id`, `email`, `full_name`, `role`, nullable `department_id`/`doctor_id`), not an access or refresh token.** Show confirmation and direct the user to the Login screen; there is no authenticated session to restore yet.

### 5.3 Password Complexity Policy
The backend enforces strong password rules:
- **Length**: At least **10 characters** and at most **72 UTF-8 bytes** (the strict bcrypt truncation boundary).
- **Character Diversity**: Must contain at least:
  - 1 uppercase letter (`A-Z`)
  - 1 lowercase letter (`a-z`)
  - 1 digit (`0-9`)
  - 1 special symbol (`!@#$%^&*...`)
- **Blocklist**: Common passwords (e.g. `Password123!`, `Admin12345!`) are rejected even if they meet diversity criteria.

Frontend clients can mirror these user-facing rules (including character count, UTF-8 byte length, and blocklist) for feedback; the backend remains authoritative.

---

## 6. Error Handling & Rate Limiting UX

### 6.1 Status Codes Reference

| HTTP Status | Meaning | Frontend Action |
| :--- | :--- | :--- |
| `200 OK` / `201 Created` | Success | Proceed with response data |
| `400 Bad Request` | Invalid input or revoked/expired invitation | Display error `detail` from response |
| `401 Unauthorized` | Invalid/expired access token or credentials | Refresh only after a protected request fails; on a login 401 show a generic credentials error. Redirect to login if refresh is revoked/expired. |
| `403 Forbidden` | Insufficient permission or rejected cookie-refresh CSRF token | Show access denied for role failures; treat refresh CSRF failures as session/auth errors rather than permission changes. |
| `404 Not Found` | Resource not found or in different department | Display not found screen |
| `409 Conflict` | Email already registered or identity collision | Prompt user to log in or use a different email |
| `422 Unprocessable` | Request validation error | Highlight offending form field |
| `429 Too Many Requests` | Rate limit exceeded | Display countdown timer using `Retry-After` header |

### 6.2 Anti-Enumeration Design
To prevent user email harvesting:
- Login failures return generic HTTP 401: `{"detail": "Username or password is incorrect"}` regardless of whether the email exists.
- The UI should never display messages such as *"User with this email does not exist"*.

### 6.3 Handling Rate Limiting (`HTTP 429`)
Sensitive endpoints (`/login`, `/signup`, `/refresh`) enforce sliding-window rate limits.

When a client is throttled, the server returns:
```http
HTTP/1.1 429 Too Many Requests
Retry-After: 45
Content-Type: application/json

{
  "detail": "Too many requests. Please try again later."
}
```

**Recommended UX Implementation**:
1. Inspect the `Retry-After` response header (value is in seconds).
2. Disable the submit button.
3. Display a countdown timer: *"Too many attempts. Please try again in 45 seconds."*
4. Re-enable the button when the timer expires.
