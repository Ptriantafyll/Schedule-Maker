# Implementation Plan: Flutter Frontend Authentication Feature

This document outlines the step-by-step engineering plan for implementing the **Authentication feature** (`features/auth`) in the Flutter frontend, applying the **Model-View-ViewModel (MVVM)** pattern with **Flutter Riverpod 2.x**, designed for high scalability, maintainability, and test-driven development (TDD).

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
        TokenStorage["TokenStorage (SecureStorage)"]
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
   - **Native (Android, iOS, Desktop)**: Tokens are stored in hardware-backed secure storage (`flutter_secure_storage`).
   - **Web (Browser)**: Refresh tokens use `HttpOnly` cookies, with automatic `X-CSRF-Token` header attachment.

---

## 2. Target Folder Structure

```text
frontend/lib/
├── features/auth/
│   ├── domain/                                 # MODEL: Entities & Domain Logic
│   │   ├── models/
│   │   │   ├── user.dart                       # User entity (id, email, fullName, role, departmentId, doctorId)
│   │   │   ├── user_role.dart                  # UserRole enum (superAdmin, departmentAdmin, doctor, viewer)
│   │   │   └── auth_tokens.dart                # AuthTokens model (accessToken, refreshToken, csrfToken)
│   │   └── state/
│   │       └── auth_state.dart                 # AuthState sealed class (Initial, Unauthenticated, Authenticated)
│   │
│   ├── data/                                   # MODEL: Data Sources & Repositories
│   │   ├── datasources/
│   │   │   └── auth_remote_data_source.dart    # Dio HTTP calls to FastAPI auth endpoints
│   │   └── repositories/
│   │       └── auth_repository.dart            # Coordinates remote API & TokenStorage
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
   - Factory constructor: `User.fromJson(Map<String, dynamic> json)`.
3. **`lib/features/auth/domain/models/auth_tokens.dart`**:
   - Holds `accessToken`, `tokenType` (bearer), `refreshToken` (nullable for web), `csrfToken` (nullable for mobile).
   - Factory constructor: `AuthTokens.fromJson(Map<String, dynamic> json)`.
4. **`lib/features/auth/domain/state/auth_state.dart`**:
   - Sealed class representing the discrete states of authentication:
     - `AuthStateInitial`: App booting, inspecting secure storage.
     - `AuthStateUnauthenticated`: No active session; render login.
     - `AuthStateAuthenticated(User user)`: Active session with authenticated user profile.

---

### Step 2: Data Layer (Remote Data Source & Repository)
**Goal:** Implement low-level HTTP calls and coordinate with `TokenStorage` behind an abstracted repository interface.

1. **`lib/features/auth/data/datasources/auth_remote_data_source.dart`**:
   - Direct network interactions using `ApiClient` / `Dio`:
     - `login(String username, String password)`: Sends `application/x-www-form-urlencoded` to `POST /api/v1/auth/login`. Returns `AuthTokens`.
     - `signup({required String invitationToken, required String email, required String password, required String fullName})`: Sends JSON to `POST /api/v1/auth/signup`. Returns `AuthTokens`.
     - `refresh({String? refreshToken, String? csrfToken})`: Calls `POST /api/v1/auth/refresh`. Returns `AuthTokens`.
     - `logout({String? refreshToken, String? csrfToken})`: Calls `POST /api/v1/auth/logout`.
     - `getCurrentUser(String accessToken)`: Calls `GET /api/v1/auth/me`. Returns `User`.
2. **`lib/features/auth/data/repositories/auth_repository.dart`**:
   - Orchestrates `AuthRemoteDataSource` and `TokenStorage`:
     - On login/signup: saves access token and refresh token via `TokenStorage.saveTokens()`. Fetches user profile via `getCurrentUser()`.
     - On logout: clears stored tokens via `TokenStorage.clearTokens()`.
     - Exposes `authRepositoryProvider` via Riverpod.

---

### Step 3: Network Hardening (Silent Refresh & Concurrency Queue)
**Goal:** Prevent unexpected logouts when short-lived access tokens expire.

- In `lib/core/network/api_client.dart`:
  - Attach a `QueuedInterceptorsWrapper` to intercept `401 Unauthorized`.
  - When a 401 occurs:
    1. Lock incoming requests to prevent duplicate refresh calls.
    2. Read stored `refreshToken`.
    3. Call `/api/v1/auth/refresh`.
    4. Save the new `accessToken` and `refreshToken`.
    5. Re-dispatch the original queued request with the updated `Authorization: Bearer <new_token>` header.
    6. Unlock the interceptor queue.
    7. If refresh fails (expired session), clear tokens and notify the auth controller.

---

### Step 4: Presentation Layer - ViewModel (Riverpod `AuthController`)
**Goal:** Expose reactive state to the UI using Riverpod 2.x `AsyncNotifier`.

- **`lib/features/auth/presentation/controllers/auth_controller.dart`**:
  - Extends `AsyncNotifier<AuthState>`.
  - Methods:
    - `Future<AuthState> build()`: Initializes auth state on startup by checking `TokenStorage`. If a valid token exists, loads current user; otherwise returns `AuthStateUnauthenticated`.
    - `Future<void> login({required String email, required String password})`: Sets state to `AsyncValue.loading()`, awaits `repository.login()`, updates state to `AsyncValue.data(AuthStateAuthenticated(user))`.
    - `Future<void> signup(...)`: Consumes invitation and authenticates user.
    - `Future<void> logout()`: Calls `repository.logout()`, resets state to `AsyncValue.data(AuthStateUnauthenticated())`.
  - Derived convenience providers:
    - `currentUserProvider`: `ref.watch(authControllerProvider).valueOrNull?.user`
    - `isAuthenticatedProvider`: `ref.watch(authControllerProvider).valueOrNull is AuthStateAuthenticated`
    - `userRoleProvider`: `ref.watch(currentUserProvider)?.role`

---

### Step 5: Presentation Layer - Views (Screens & AuthGate)
**Goal:** Material 3 UI widgets bound cleanly to the Riverpod ViewModel.

1. **`lib/features/auth/presentation/widgets/auth_gate.dart`**:
   - Root widget listening to `authControllerProvider`:
     - When `AuthStateInitial` or `loading`: Renders a loading splash screen.
     - When `AuthStateUnauthenticated`: Renders `LoginScreen`.
     - When `AuthStateAuthenticated`: Renders `HomeScreen` (or dashboard).
2. **`lib/features/auth/presentation/widgets/auth_text_field.dart`**:
   - Material 3 styled text field with label, validation, password toggle visibility, and error states.
3. **`lib/features/auth/presentation/screens/login_screen.dart`**:
   - Form with email and password fields.
   - "Sign In" button with loading spinner when `AsyncValue.isLoading`.
   - Listens to `authControllerProvider` using `ref.listen` to show Material 3 `SnackBar` on failure.
   - Link to "Have an invitation? Sign up here".
4. **`lib/features/auth/presentation/screens/signup_screen.dart`**:
   - Form taking `invitation_token`, `full_name`, `email`, and `password`.
   - Consumes invitation and transitions immediately to authenticated state.

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
  - Parse JSON responses for each backend role (`super_admin`, `department_admin`, `doctor`, `viewer`).
  - Validate role permission helper getters.
- **Repository Tests (`test/features/auth/data/auth_repository_test.dart`)**:
  - Verify tokens are saved to secure storage on successful login, and cleared on logout.
  - Verify `ApiException` is mapped properly when credentials are invalid.
- **Controller Tests (`test/features/auth/presentation/auth_controller_test.dart`)**:
  - Test state transitions with Riverpod `ProviderContainer`: `Unauthenticated` ➔ `Loading` ➔ `Authenticated`.
  - Test failure propagation to `AsyncValue.error`.
- **Widget Tests (`test/features/auth/presentation/login_screen_test.dart`)**:
  - Test form input validation (empty email/password).
  - Test button tap dispatches `login()` to `AuthController`.
  - Test loading indicator appears while authenticating.
