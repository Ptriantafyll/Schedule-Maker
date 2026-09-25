# Implementation Plan: Flutter Frontend Authentication Feature

This document outlines the step-by-step engineering plan for implementing the **Authentication feature** (`features/auth`) in the Flutter frontend, applying the **Model-View-ViewModel (MVVM)** pattern with **Flutter Riverpod 2.x**, designed for high scalability, maintainability, and test-driven development (TDD). It is step 1 of the [frontend roadmap](frontend_implementation_roadmap.md); these screens and session flows are planned, not yet implemented.

---

## 1. Architectural Overview (MVVM + Riverpod 2.x)

```mermaid
flowchart TD
    subgraph View ["Presentation Layer (View)"]
        LoginScreen["LoginScreen / SignupScreen"]
        AuthGate["AuthGate (Root Router)"]
    end

    subgraph ViewModel ["Presentation Layer (ViewModel)"]
        AuthController["AuthController (AsyncNotifier<AuthState>)"]
        Providers["Derived Providers (currentUser, isAuthenticated)"]
    end

    subgraph Model ["Data & Domain Layer (Model)"]
        AuthRepository["AuthRepository"]
        AuthRemoteDataSource["AuthRemoteDataSource (Dio)"]
        TokenStorage["Platform session transport (native secure storage / web cookies)"]
        DomainModels["Domain Entities (User, UserRole, AuthTokens)"]
    end

    LoginScreen -->|ref.read(authController.notifier).login()| AuthController
    AuthGate -->|ref.watch(authControllerProvider)| AuthController
    AuthController -->|Mutates State (AsyncValue<AuthState>)| View
    AuthController -->|Calls repository| AuthRepository
    AuthRepository -->|Network calls| AuthRemoteDataSource
    AuthRepository -->|Persist tokens| TokenStorage
    AuthRemoteDataSource --> DomainModels
```

### Core Design Decisions
1. **Unidirectional Data Flow**: The View observes state emitted by the ViewModel and notifies the ViewModel of user actions. The ViewModel communicates with the Model (Repository) to execute business logic. The View never makes network calls or stores raw tokens.
2. **Form Data vs JSON**: Backend `/api/v1/auth/login` requires `application/x-www-form-urlencoded` format (`username` and `password`), while `/api/v1/auth/signup` and others require `application/json`. `AuthRemoteDataSource` handles these formats explicitly.
3. **Dual Transport Architecture**:
   - **Native (Android, iOS, Windows)**: Persist rotating refresh tokens in platform-supported secure storage; keep short-lived access tokens in memory where practical.
   - **Web (Browser)**: Keep the access token in memory; use the browser-managed `HttpOnly` refresh cookie and send `X-CSRF-Token` for cookie refresh. Never store the JSON refresh token in browser storage.
4. **Invitation signup is not login**: `/api/v1/auth/signup` accepts `invitation_token`, `first_name`, `last_name`, `email`, and `password`, returns `201 UserRead` **without tokens**, and sends the user to the login screen.
5. **Roles are server-defined**: Derive `role`, `department_id`, and `doctor_id` from `/auth/me`. The super admin has no department scope and must not be routed to departmental schedule endpoints; UI guards do not replace backend authorization.

---

## 2. Target Folder Structure

```text
frontend/lib/
├── features/auth/
│   ├── domain/                                 # MODEL: Entities & Domain Logic
│   │   ├── models/
│   │   │   ├── user.dart                       # User entity (id, email, fullName, role, departmentId, doctorId)
│   │   │   ├── user_role.dart                  # UserRole enum (superAdmin, departmentAdmin, doctor, viewer)
│   │   │   └── auth_tokens.dart                # Pure in-memory credentials model; no JSON parsing
│   │   └── state/
│   │       └── auth_state.dart                 # AuthState (Unauthenticated, Authenticated); boot via AsyncValue.loading
│   │
│   ├── data/                                   # MODEL: Data Sources & Repositories
│   │   ├── dtos/
│   │   │   ├── user_dto.dart                   # Parse UserRead and map to User
│   │   │   └── auth_tokens_dto.dart            # Parse login/refresh; never log tokens
│   │   ├── datasources/
│   │   │   └── auth_remote_data_source.dart    # Dio HTTP calls to FastAPI auth endpoints
│   │   └── repositories/
│   │       └── auth_repository.dart            # Coordinates remote API & platform session transport
│   │
│   └── presentation/                           # VIEW & VIEWMODEL
│       ├── controllers/
│       │   └── auth_controller.dart            # VIEWMODEL: Riverpod AsyncNotifier<AuthState>
│       ├── screens/
│       │   ├── login_screen.dart               # VIEW: Material 3 email/password login
│       │   └── signup_screen.dart              # VIEW: Invitation-token registration screen
│       └── widgets/
│           ├── auth_gate.dart                  # VIEW: Session-based root widget switcher
│           └── auth_text_field.dart            # Reusable styled input field with validation
│
└── test/features/auth/                         # TDD Unit & Widget Tests
    ├── domain/user_test.dart
    ├── data/auth_repository_test.dart
    ├── presentation/auth_controller_test.dart
    └── presentation/login_screen_test.dart
```

---

## 3. Step-by-Step Implementation Sequence

### Step 1: Domain Layer (Entities & State Models)
**Goal:** Define pure Dart entities representing users, credentials, and application authentication state with zero Flutter UI dependencies.

1. **`lib/features/auth/domain/models/user_role.dart`**:
   - Enum matching backend roles: `superAdmin` (`"super_admin"`), `departmentAdmin` (`"department_admin"`), `doctor` (`"doctor"`), `viewer` (`"viewer"`).
   - Helper methods: `requiresDepartment`, `requiresDoctor`, `displayName`.
2. **`lib/features/auth/domain/models/user.dart`**:
   - Immutable entity: `id` (String), `email` (String), `fullName` (String), `role` (`UserRole`), `departmentId` (String?), `doctorId` (String?).
   - Parse `/auth/me` JSON in `data/dtos/user_dto.dart`, then map to this pure domain entity; reject unknown roles rather than guessing.
3. **`lib/features/auth/domain/models/auth_tokens.dart`**:
   - Holds `accessToken`, `tokenType` (bearer), `refreshToken` (nullable for web), `csrfToken` (nullable for mobile).
   - Parse token JSON in `data/dtos/auth_tokens_dto.dart`; never include token values in debug output or logs.
4. **`lib/features/auth/domain/state/auth_state.dart`**:
   - Sealed class representing the discrete states of authentication:
     - `AuthStateUnauthenticated`: No active session; render login.
     - `AuthStateAuthenticated(User user)`: Active session with authenticated user profile.
   - Use Riverpod `AsyncValue.loading` during startup instead of duplicating it with `AuthStateInitial`.

---

### Step 2: Data Layer (Remote Data Source & Repository)
**Goal:** Implement low-level HTTP calls and coordinate with platform-specific session transport behind a repository interface. Complete the existing `lib/core/network/api_client.dart`, `core/storage/token_storage.dart`, and `core/config/app_config.dart` rather than adding parallel infrastructure.

1. **`lib/features/auth/data/datasources/auth_remote_data_source.dart`**:
   - Direct network interactions using `ApiClient` / `Dio`:
     - `login(String username, String password)`: Sends `application/x-www-form-urlencoded` to `POST /api/v1/auth/login`. Returns `AuthTokens`.
     - `signup({required String invitationToken, required String firstName, required String lastName, required String email, required String password})`: Sends JSON to `POST /api/v1/auth/signup`. Returns a mapped `UserRead`, **not** tokens; never send role or department from the client.
     - `refresh({String? refreshToken, String? csrfToken})`: Calls `POST /api/v1/auth/refresh`. Returns `AuthTokens`.
     - `logout({String? refreshToken, String? csrfToken})`: Calls `POST /api/v1/auth/logout`.
     - `getCurrentUser()`: Calls `GET /api/v1/auth/me` with the access token attached by the client. Map the response through `UserDto`.
2. **`lib/features/auth/data/repositories/auth_repository.dart`**:
   - Orchestrates `AuthRemoteDataSource` and native secure storage or browser cookie transport:
     - On login: save credentials using the correct platform policy and fetch `/auth/me` before returning an authenticated user.
     - On signup: return signup success without saving tokens, fetching `/auth/me`, or setting authenticated state; navigate to login.
     - On logout: revoke the server session where possible, clear local credentials, and surface errors instead of falsely claiming server logout.
     - On app startup: validate or refresh the session before fetching `/auth/me`; distinguish invalid credentials from a temporary network outage.
     - Exposes `authRepositoryProvider` via Riverpod.

---

### Step 3: Network Hardening (Platform Transport & Single-Flight Refresh)
**Goal:** Prevent unexpected logouts when short-lived access tokens expire.

- In `lib/core/network/api_client.dart`:
  - Validate `API_BASE_URL`, permit a form-encoded login request, and map 409/422/429 (`Retry-After`) and connectivity errors without leaking credentials. Intercept only protected-request `401`s; public login failures must remain visible as login errors.
  - When an access token expires:
    1. Share **one in-flight refresh operation** among concurrent failures; a queued interceptor by itself does not guarantee a single rotation.
    2. Native: read the securely stored refresh token. Web: send browser cookie credentials and `X-CSRF-Token` without sending a JSON refresh token.
    3. Call `/api/v1/auth/refresh` once; never recursively refresh the refresh call.
    4. Update the access token and persist the rotated native refresh token or accept the new browser cookie.
    5. Re-dispatch the original queued request with the updated `Authorization: Bearer <new_token>` header.
    6. Ensure step 5 happens at most once per original request, then release waiting requests.
    7. On expired/revoked refresh, clear the session and notify the auth controller; surface temporary network errors instead of silently logging out.

---

### Step 4: Presentation Layer - ViewModel (Riverpod `AuthController`)
**Goal:** Expose reactive state to the UI using Riverpod 2.x `AsyncNotifier`.

- **`lib/features/auth/presentation/controllers/auth_controller.dart`**:
  - Extends `AsyncNotifier<AuthState>`.
  - Methods:
    - `Future<AuthState> build()`: Restores a native or browser session, fetches the user if valid, and explicitly reports loading or errors.
    - `Future<void> login({required String email, required String password})`: Sets state to `AsyncValue.loading()`, awaits `repository.login()`, updates state to `AsyncValue.data(AuthStateAuthenticated(user))`.
    - `Future<void> signup(...)`: Consumes invitation, remains unauthenticated, and sends the user to login after a success message.
    - `Future<void> logout()`: Calls `repository.logout()`, resets state to `AsyncValue.data(AuthStateUnauthenticated())`.
  - Derived convenience providers:
    - `currentUserProvider`: returns a user only when the current state is `AuthStateAuthenticated`.
    - `isAuthenticatedProvider`: true only after `/auth/me` has validated the session, never just after signup.
    - `userRoleProvider`: `ref.watch(currentUserProvider)?.role`

---

### Step 5: Presentation Layer - Views (Screens & AuthGate)
**Goal:** Material 3 UI widgets bound cleanly to the Riverpod ViewModel.

> [!IMPORTANT]
> **Presentation Layer Rule**: For the presentation layer (UI screens, dialogs, forms, layout widgets), **always ask the user for the design first and ask clarifying questions** before writing any code or proposing UI designs.

1. **`lib/features/auth/presentation/widgets/auth_gate.dart`**:
   - Root widget listening to `authControllerProvider`:
     - When `loading`: Renders a loading splash screen.
     - When `AuthStateUnauthenticated`: Renders `LoginScreen`.
     - When `AuthStateAuthenticated`: Renders the role-aware app shell; super admins go to provisioning, while department members see only their allowed destinations.
2. **`lib/features/auth/presentation/widgets/auth_text_field.dart`**:
   - Start with Material 3 `TextFormField` for both forms. Keep password visibility auth-owned; extract shared field/action styling into common widgets only if later features need the same behavior.
3. **`lib/features/auth/presentation/screens/login_screen.dart`**:
   - Form with email and password fields.
   - "Sign In" button with loading spinner when `AsyncValue.isLoading`.
   - Listens to `authControllerProvider` using `ref.listen` to show Material 3 `SnackBar` on failure.
   - Link to "Have an invitation? Sign up here".
4. **`lib/features/auth/presentation/screens/signup_screen.dart`**:
   - Form taking an invitation token (possibly prefilled from a link), first name, last name, email, and password.
   - On `201 UserRead`, shows success and navigates to login; do not persist the one-time invitation token in logs.

---

## 4. Verification & Testing Plan

### Automated Unit & Widget Tests
```bash
# Run tests specifically for the auth feature
flutter test test/features/auth/

# Run complete test suite
flutter test

# Run static analysis
flutter analyze
```

#### Test Cases:
- **Domain Tests (`test/features/auth/domain/user_test.dart`)**:
  - Validate role helpers and pure entities; parse `UserRead` and token JSON in data DTO tests for all four roles, nullable IDs, and unknown/malformed roles.
- **Repository Tests (`test/features/auth/data/auth_repository_test.dart`)**:
  - Verify native tokens rotate and clear on logout, web never stores JSON refresh tokens, and signup stores none. Distinguish invalid credentials, revoked invitation, rate limit, and transient network errors.
- **Controller Tests (`test/features/auth/presentation/auth_controller_test.dart`)**:
  - Test login `Unauthenticated -> Loading -> Authenticated`, signup success staying unauthenticated, restoration/expiration, and error propagation with Riverpod overrides.
- **Widget Tests (`test/features/auth/presentation/login_screen_test.dart`)**:
  - Test form input validation (empty email/password).
  - Test button tap dispatches `login()` to `AuthController`.
  - Test loading while authenticating, signup fields/success -> login, and role-aware navigation.
- **Core Network and Platform Tests (`test/core/network/`)**:
  - Test one refresh for concurrent protected 401s, retry only once, no refresh for login 401, browser cookie/CSRF, native secure storage rotation, and explicit refresh failures. Integrate with Web, Windows, and Android/iOS before marking auth complete.
