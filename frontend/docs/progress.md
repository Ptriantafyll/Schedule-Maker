# Authentication Feature Progress & Step-by-Step Implementation

This document tracks the end-to-end implementation of the **Authentication feature** (`features/auth`) in the Flutter frontend, following **MVVM** and **Riverpod 2.x**.

For each step, it outlines:
1. **Business Reason (The Problem We Are Solving)**: Why a hospital scheduling app requires this capability.
2. **Technical Implementation Needed**: The exact classes, methods, and tests required.
3. **Current Status**: `[ ] Pending`, `[ ] In Progress`, or `[x] Completed`.

---

## Progress Overview

| Phase | Description | Status |
| :--- | :--- | :--- |
| **Phase 1** | Domain Layer (Pure Business Entities & State Models) | `[ ] In Progress` |
| **Phase 2** | Data Layer: Remote Data Source (`AuthRemoteDataSource`) | `[ ] Pending` |
| **Phase 3** | Data Layer: Repository (`AuthRepository`) | `[ ] Pending` |
| **Phase 4** | Core Network Polish: Silent 401 Refresh Queue | `[ ] Pending` |
| **Phase 5** | ViewModel: State Management (`AuthController` via Riverpod) | `[ ] Pending` |
| **Phase 6** | View Layer: Presentation (`AuthGate`, `LoginScreen`, `SignupScreen`) | `[ ] Pending` |
| **Phase 7** | End-to-End Verification & Backend Integration | `[ ] Pending` |

---

## Phase 1: Domain Layer (Pure Business Entities & State Models)

The Domain Layer defines the core data contracts and state definitions of the application. It has zero dependencies on Flutter UI (`BuildContext`, widgets) and zero dependencies on HTTP libraries (`Dio`).

### Step 1.1: `UserRole` Enum & Permission Rules
- **Status:** `[ ] Pending`
- **Business Reason:** A hospital operates on strict organizational hierarchy. A Super Admin manages hospital-wide tenants, a Department Admin manages shifts and doctor assignments for their department, a Doctor can only view schedules and submit time-off requests, and a Viewer has read-only access. The frontend must enforce these roles so doctors never see administrative controls, and department admins never accidentally mutate other departments.
- **Technical Implementation Needed:**
  - Create `lib/features/auth/domain/models/user_role.dart`.
  - Enum values: `superAdmin` (`"super_admin"`), `departmentAdmin` (`"department_admin"`), `doctor` (`"doctor"`), `viewer` (`"viewer"`).
  - Factory `UserRole.fromString(String value)` to parse backend strings, throwing a descriptive `FormatException` on invalid values.
  - Helper getters:
    - `bool get isAdmin => this == superAdmin || this == departmentAdmin;`
    - `bool get requiresDepartment => this != superAdmin;`
    - `bool get requiresDoctor => this == doctor;`
    - `String get displayName` for human-readable labels in the UI.
  - **TDD Test:** Create `test/features/auth/domain/models/user_role_test.dart` verifying parsing of all 4 roles, invalid string rejection, and permission helpers.

---

### Step 1.2: `User` Entity
- **Status:** `[ ] Pending`
- **Business Reason:** The app must display the doctor's or administrator's identity (e.g. "Welcome, Dr. Smith", display email, active role) and know which department they belong to so schedule queries are automatically scoped.
- **Technical Implementation Needed:**
  - Create `lib/features/auth/domain/models/user.dart`.
  - Immutable class with fields: `id` (String), `email` (String), `fullName` (String), `role` (`UserRole`), `departmentId` (String?), `doctorId` (String?).
  - Factory constructor `User.fromJson(Map<String, dynamic> json)` parsing the backend `/api/v1/auth/me` response.
  - Value equality (`==` and `hashCode`) and `toString()` for debugging and test comparisons.
  - **TDD Test:** Create `test/features/auth/domain/models/user_test.dart` testing deserialization with and without nullable `departmentId` / `doctorId`.

---

### Step 1.3: `AuthTokens` Model
- **Status:** `[ ] Pending`
- **Business Reason:** The backend returns credentials consisting of a short-lived access token, a long-lived refresh token, and an optional CSRF token. The frontend needs a structured in-memory model to hold this payload after authentication.
- **Technical Implementation Needed:**
  - Create `lib/features/auth/domain/models/auth_tokens.dart`.
  - Fields: `accessToken` (String), `tokenType` (String), `refreshToken` (String?), `csrfToken` (String?).
  - Factory `AuthTokens.fromJson(Map<String, dynamic> json)`.
  - **TDD Test:** Create `test/features/auth/domain/models/auth_tokens_test.dart` verifying parsing of login responses.

---

### Step 1.4: `AuthState` Sealed Class
- **Status:** `[ ] Pending`
- **Business Reason:** The application UI needs a clean, compile-time safe way to know what screen to show at any given second: is the app still reading saved keys from the phone? Is the user logged out? Or are they logged in as Dr. Smith? Using a sealed class guarantees that the UI handles every possible auth state without missing any edge case.
- **Technical Implementation Needed:**
  - Create `lib/features/auth/domain/state/auth_state.dart`.
  - Dart 3 sealed class hierarchy:
    - `AuthStateInitial`: App is booting and checking secure storage.
    - `AuthStateUnauthenticated`: No active session; show Login / Signup.
    - `AuthStateAuthenticated(User user)`: Active session with user identity; show Dashboard.
  - **TDD Test:** Create `test/features/auth/domain/state/auth_state_test.dart` testing pattern matching and equality.

---

## Phase 2: Data Layer: Remote Data Source (`AuthRemoteDataSource`)

### Step 2.1: Login Endpoint Call (`POST /api/v1/auth/login`)
- **Status:** `[ ] Pending`
- **Business Reason:** Doctors and administrators need to enter their hospital credentials to securely authenticate into the system.
- **Technical Implementation Needed:**
  - Create `lib/features/auth/data/datasources/auth_remote_data_source.dart`.
  - Implement `login({required String username, required String password})`:
    - Sends an `application/x-www-form-urlencoded` POST request to `/api/v1/auth/login`.
    - Returns `AuthTokens`.

---

### Step 2.2: Invitation Signup Endpoint Call (`POST /api/v1/auth/signup`)
- **Status:** `[ ] Pending`
- **Business Reason:** Hospital security policy prohibits open public signups. Doctors and staff must receive an invitation link containing a one-time cryptographic token from an administrator before creating an account.
- **Technical Implementation Needed:**
  - In `AuthRemoteDataSource`, implement `signup({required String invitationToken, required String email, required String password, required String fullName})`:
    - Sends a JSON POST request to `/api/v1/auth/signup`.
    - Returns `AuthTokens`.

---

### Step 2.3: Refresh, Logout, and Current User Calls
- **Status:** `[ ] Pending`
- **Business Reason:** Support extending sessions, cleanly terminating sessions (revoking the refresh token in the backend database), and fetching user profiles on startup.
- **Technical Implementation Needed:**
  - `refresh({String? refreshToken, String? csrfToken})`: Calls `POST /api/v1/auth/refresh`.
  - `logout({String? refreshToken, String? csrfToken})`: Calls `POST /api/v1/auth/logout`.
  - `getCurrentUser(String accessToken)`: Calls `GET /api/v1/auth/me` with Bearer token, returns `User`.
  - **TDD Test:** Create `test/features/auth/data/auth_remote_data_source_test.dart` using a mocked HTTP client.

---

## Phase 3: Data Layer: Repository (`AuthRepository`)

### Step 3.1: Session & Storage Orchestration
- **Status:** `[ ] Pending`
- **Business Reason:** The presentation layer and ViewModels should never know about raw HTTP details or secure storage key names. The repository acts as the single source of truth for authentication: when login succeeds, it automatically stores tokens securely and fetches the user profile before returning.
- **Technical Implementation Needed:**
  - Create `lib/features/auth/data/repositories/auth_repository.dart`.
  - Coordinates `AuthRemoteDataSource` and `TokenStorage`:
    - `login(...)`: Calls remote login, calls `tokenStorage.saveTokens()`, fetches `getCurrentUser()`, returns `User`.
    - `signup(...)`: Calls remote signup, saves tokens, fetches `getCurrentUser()`, returns `User`.
    - `logout(...)`: Calls remote logout, calls `tokenStorage.clearTokens()`.
    - `restoreSession()`: Reads token from storage; if present, verifies by calling `getCurrentUser()`; if invalid/expired, clears storage.
  - Expose `authRepositoryProvider = Provider<AuthRepository>((ref) => ...)`.
  - **TDD Test:** Create `test/features/auth/data/auth_repository_test.dart` mocking the data source and storage.

---

## Phase 4: Core Network Polish: Silent 401 Refresh Queue

### Step 4.1: Concurrency-Safe Refresh Interceptor
- **Status:** `[ ] Pending`
- **Business Reason:** JWT access tokens expire every 15–30 minutes for security. If a doctor is working in the app when the token expires, their next actions should not fail or kick them out to the login screen. Furthermore, if three widgets fetch data concurrently and all get a 401, the app must not fire three simultaneous refresh requests (which would invalidate each other); it must pause, refresh once, and replay all three requests seamlessly.
- **Technical Implementation Needed:**
  - Update `lib/core/network/api_client.dart` with a `QueuedInterceptorsWrapper`.
  - On `onError` with `statusCode == 401`:
    - Lock request queue.
    - Read refresh token from `TokenStorage`.
    - Call `/api/v1/auth/refresh`.
    - Store new tokens.
    - Retry original request with new access token.
    - Unlock queue.
    - If refresh fails (e.g. refresh token revoked/expired), purge storage and notify auth state.
  - **TDD Test:** Create `test/core/network/api_client_refresh_test.dart`.

---

## Phase 5: ViewModel: State Management (`AuthController`)

### Step 5.1: Riverpod `AsyncNotifier` ViewModel
- **Status:** `[ ] Pending`
- **Business Reason:** Screens need a reactive state manager that handles loading indicators, captures error messages, and triggers navigation when authentication succeeds or fails.
- **Technical Implementation Needed:**
  - Create `lib/features/auth/presentation/controllers/auth_controller.dart`.
  - Class `AuthController extends AsyncNotifier<AuthState>`:
    - `build()`: Calls `repository.restoreSession()`; returns `AuthStateAuthenticated(user)` or `AuthStateUnauthenticated()`.
    - `login({required String email, required String password})`: Sets state to `AsyncValue.loading()`, awaits repository call, sets `AsyncValue.data(AuthStateAuthenticated(user))`. On failure, sets `AsyncValue.error(exception)`.
    - `signup(...)`: Similar flow for invitation acceptance.
    - `logout()`: Calls `repository.logout()`, sets `AsyncValue.data(const AuthStateUnauthenticated())`.
  - Derived convenience providers:
    - `currentUserProvider`: `ref.watch(authControllerProvider).valueOrNull?.user`
    - `isAuthenticatedProvider`: `ref.watch(authControllerProvider).valueOrNull is AuthStateAuthenticated`
    - `userRoleProvider`: `ref.watch(currentUserProvider)?.role`
  - **TDD Test:** Create `test/features/auth/presentation/auth_controller_test.dart` using Riverpod `ProviderContainer`.

---

## Phase 6: View Layer: Presentation (`AuthGate`, `LoginScreen`, `SignupScreen`)

### Step 6.1: `AuthGate` Navigation Controller Widget
- **Status:** `[ ] Pending`
- **Business Reason:** When the user opens the application, they should automatically see the Dashboard if already logged in, or the Login screen if logged out, without flickering or race conditions.
- **Technical Implementation Needed:**
  - Create `lib/features/auth/presentation/widgets/auth_gate.dart`.
  - Listens to `authControllerProvider`:
    - `AuthStateInitial` / loading ➔ Displays a centered Material 3 loading indicator or splash screen.
    - `AuthStateUnauthenticated` ➔ Renders `LoginScreen`.
    - `AuthStateAuthenticated` ➔ Renders `HomeScreen` (or dashboard placeholder).

---

### Step 6.2: `LoginScreen` & Reusable Auth Text Field
- **Status:** `[ ] Pending`
- **Business Reason:** A clean, accessible Material 3 interface allowing doctors and staff to enter email and password, see clear validation errors (e.g. empty fields, invalid email format), and view a progress spinner during submission.
- **Technical Implementation Needed:**
  - Create `lib/features/auth/presentation/widgets/auth_text_field.dart`:
    - Reusable styled input with validation, obscuring toggle for password, and helper text.
  - Create `lib/features/auth/presentation/screens/login_screen.dart`:
    - Form with email and password fields.
    - Button calling `ref.read(authControllerProvider.notifier).login(...)`.
    - Uses `ref.listen` on `authControllerProvider` to display errors via `SnackBar`.
    - Link to "Have an invitation? Sign up here".
  - **Widget Test:** Create `test/features/auth/presentation/login_screen_test.dart` testing validation, button disabled/loading state, and error display.

---

### Step 6.3: `SignupScreen` (Invitation Acceptance)
- **Status:** `[ ] Pending`
- **Business Reason:** A dedicated screen where a doctor or administrator invited to the hospital can enter their invitation token, name, email, and choose their password to activate their account.
- **Technical Implementation Needed:**
  - Create `lib/features/auth/presentation/screens/signup_screen.dart`:
    - Form with `invitationToken`, `fullName`, `email`, and `password`.
    - Submits via `ref.read(authControllerProvider.notifier).signup(...)`.
  - **Widget Test:** Create `test/features/auth/presentation/signup_screen_test.dart`.

---

## Phase 7: End-to-End Verification & Backend Integration

### Step 7.1: Live Integration Verification
- **Status:** `[ ] Pending`
- **Business Reason:** Verify that the frontend seamlessly connects with the live FastAPI backend, successfully authenticates a bootstrapped super admin, rotates tokens, and accesses protected endpoints.
- **Verification Routine:**
  1. Run FastAPI backend: `uv run uvicorn src.main:app --reload`
  2. Run Flutter app: `flutter run`
  3. Log in with Super Admin credentials.
  4. Verify transition to Home screen.
  5. Verify secure token storage on Android/iOS/Desktop.
  6. Trigger logout and verify return to Login screen.
