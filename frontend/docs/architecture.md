# Frontend Architecture

## 1. Overview

The **Schedule-Maker Frontend** (**MedShift**) is a multi-platform Flutter application targeting **Web, Windows, Android, and iOS**. It provides hospital administrators, doctors, and viewers with a unified interface to authenticate, onboard via invitations, generate optimized monthly duty schedules (from the in-app department roster or uploaded Excel workbooks), preview shift assignments in interactive calendar and table views, publish schedules, and export formatted `.xlsx` workbooks.

The application is built on a **Feature-First Clean Architecture** that implements the **MVVM (Model-View-ViewModel)** pattern using **Flutter Riverpod** and **unidirectional data flow**.

---

## 2. Technology Stack

| Layer / Concern | Package / Technology | Purpose |
| :--- | :--- | :--- |
| **UI Framework** | Flutter (Dart `^3.12.0`) | Cross-platform declarative UI with **Material 3** (`AppTheme` light & dark schemes) |
| **State Management & DI** | `flutter_riverpod` (`^3.4.3`) | Reactive ViewModels (`Notifier`, `AsyncNotifier`), dependency injection, and state binding |
| **Networking** | `dio` (`^5.11.0`) | HTTP client (`ApiClient`) with Bearer token injection, single-flight 401 refresh, and error mapping |
| **Secure Storage** | `flutter_secure_storage` (`^11.0.0`) | Hardware-backed persistence (`TokenStorage`) for JWT access and refresh tokens |
| **File I/O** | `file_picker` (`^13.1.0`) | Cross-platform `.xlsx` file selection for upload and saving exported schedules |
| **Loading UX** | `skeletonizer` (`^3.0.0`) | Shimmer skeleton placeholders rendered over real widget layouts during async loads |
| **Localization / Dates** | `intl` (`^0.20.3`) | Date and calendar formatting |

---

## 3. How the Project Uses MVVM (Model-View-ViewModel)

Each feature module (`lib/features/<feature>/`) is structured around the **MVVM** architectural pattern, combined with Clean Architecture's separation of `domain`, `data`, and `presentation` layers.

```mermaid
flowchart TB
    subgraph View ["View Layer (presentation/screens & widgets)"]
        UI["Flutter Screens & Widgets\n(AdminScreen, LoginScreen, AuthGate, DraftPreviewCanvas)"]
    end

    subgraph ViewModel ["ViewModel Layer (presentation/controllers & state)"]
        VM["Riverpod Notifier / AsyncNotifier\n(ScheduleGenerationController, AuthController)"]
        State["Immutable UI State\n(ScheduleGenerationState, AsyncValue&lt;AuthState&gt;)"]
        VM -->|Emits new immutable state| State
    end

    subgraph Model ["Model Layer (domain/ & data/)"]
        Repo["Repositories\n(ScheduleDraftRepository, AuthRepository)"]
        Domain["Immutable Domain Models\n(ScheduleDraft, ScheduleSummary, TargetMonthInfo, User, AuthTokens)"]
        DS["Remote Data Sources & DTOs\n(ScheduleDraftRemoteDataSource, AuthRemoteDataSource)"]
        Core["Core Infrastructure\n(ApiClient / Dio, TokenStorage)"]
        Repo -->|Returns| Domain
        Repo -->|Delegates to| DS
        DS -->|Maps JSON via DTO .toDomain()| Domain
        DS -->|Calls| Core
    end

    UI -- "1. Dispatches user intents\nref.read(provider.notifier).action()" --> VM
    State -- "2. Reactive data binding\nref.watch(provider) / ref.listen(provider)" --> UI
    VM -- "3. Invokes domain operations\nref.read(repositoryProvider)" --> Repo
```

### 3.1. Model (`domain/` and `data/`)

In this project, the **Model** represents both the business domain entities and the data access layer responsible for retrieving, persisting, and transforming them. It has zero dependency on Flutter widgets or UI state.

1. **Domain Models (`domain/models/`)**:
   - Pure, immutable Dart classes with value equality (`operator ==`, `hashCode`, `copyWith`) and domain-specific computed properties.
   - **Examples**:
     - `User` and `UserRole` (`lib/features/auth/domain/models/`): Encapsulate user identity and role rules (`isAdmin`, `requiresDepartment`, `requiresDoctor`, `displayName`).
     - `ScheduleDraft` and `ScheduleAssignment` (`lib/features/admin/domain/models/schedule_draft.dart`): Represent a generated or published monthly schedule and its individual shift assignments, exposing domain getters like `isPublished` (`status == 'published'`) and `isOneTime` (`status == 'one_time'`).
     - `ScheduleSummary`, `TargetMonthInfo`, and `ExcelRosterDiff` (`lib/features/admin/domain/models/`): Represent schedule history items, next target month resolution (`hasPublishedSchedules`), and Excel-vs-database roster diffs (`requiresReconciliation`).

2. **Data Transfer Objects (`data/dtos/`)**:
   - Isolate raw JSON serialization (`fromJson`, `toJson`) from domain models and validate required backend fields (throwing `FormatException` on malformed contracts).
   - Convert API payloads into clean domain entities via `.toDomain()` (e.g., `UserDto.toDomain()`, `ScheduleDraftDto.toDomain()`, `ExcelRosterDiffDto.toDomain()`).

3. **Data Sources (`data/datasources/`)**:
   - Perform low-level HTTP calls through the shared `ApiClient` (`Dio`) and return mapped domain objects.
   - **Examples**: `AuthRemoteDataSource` (`/api/v1/auth/*` endpoints) and `ScheduleDraftRemoteDataSource` (`/api/v1/schedules/*` endpoints, multipart `.xlsx` uploads, and binary byte exports).

4. **Repositories (`data/repositories/`)**:
   - Define abstract contracts (`AuthRepository`, `ScheduleDraftRepository`) and concrete implementations (`AuthRepositoryImpl`, `ScheduleDraftRepositoryImpl`) provided via Riverpod (`authRepositoryProvider`, `scheduleDraftRepositoryProvider`).
   - Coordinate across data sources and local storage. For example, `AuthRepositoryImpl.login()` authenticates via `AuthRemoteDataSource`, persists tokens in `TokenStorage`, and fetches the authenticated `User` profile before returning it to the ViewModel.

---

### 3.2. ViewModel (`presentation/controllers/` and `presentation/state/`)

The **ViewModel** layer is implemented using **Riverpod `Notifier` and `AsyncNotifier` controllers** paired with immutable **State** classes. ViewModels expose observable UI state to the View and process user commands by calling Repositories.

1. **Immutable UI State (`presentation/state/` & `domain/state/`)**:
   - Holds all presentation state required to render a screen.
   - **`ScheduleGenerationState`** (`lib/features/admin/presentation/state/schedule_generation_state.dart`):
     - Tracks `status` (`GenerationStatus.idle`, `solving`, `success`, `error`), `draft` (`ScheduleDraft?`), `selectedMonth`, `targetMonthInfo`, `scheduleHistory`, `isLoadingDashboard`, `isExporting`, `isPublishing`, and `errorMessage`.
     - Provides derived UI getters (`isGenerated`, `isSolving`, `isPublished`) and an immutable `copyWith(...)` method.
   - **`AuthState`** (`lib/features/auth/domain/state/auth_state.dart`):
     - A sealed class hierarchy (`AuthStateUnauthenticated`, `AuthStateAuthenticated(User user)`) wrapped in Riverpod's `AsyncValue<AuthState>` to represent loading, authenticated, unauthenticated, and error states.

2. **ViewModel Controllers (`presentation/controllers/`)**:
   - Never import `package:flutter/material.dart` or hold `BuildContext` references, making them 100% unit-testable in pure Dart.
   - **`AuthController extends AsyncNotifier<AuthState>`** (`lib/features/auth/presentation/controllers/auth_controller.dart`):
     - Restores the user session automatically in `build()` via `authRepository.restoreSession()`.
     - Exposes command methods (`login()`, `signup()`, `logout()`) that transition `state` through `AsyncValue.loading()`, `AsyncValue.data(...)`, and `AsyncValue.error(...)`.
     - Provides derived selector providers (`currentUserProvider`, `isAuthenticatedProvider`, `userRoleProvider`) so Views can watch fine-grained slices of state.
   - **`ScheduleGenerationController extends Notifier<ScheduleGenerationState>`** (`lib/features/admin/presentation/controllers/schedule_generation_controller.dart`):
     - Orchestrates the Admin dashboard lifecycle: `initializeDashboard()`, `selectMonth()`, `loadActiveDraft()`, `generateFromCurrentRoster()`, `generateFromExcel()`, `publishCurrentDraft()`, and `exportCurrentDraft()`.
     - Catches `ApiException` from the repository layer and maps it into `state.copyWith(status: GenerationStatus.error, errorMessage: ...)`.

---

### 3.3. View (`presentation/screens/`, `presentation/widgets/`, and `shared/widgets/`)

The **View** layer consists of declarative Flutter widgets (`ConsumerWidget`, `ConsumerStatefulWidget`, `StatelessWidget`, `StatefulWidget`) whose sole responsibility is rendering UI from state and forwarding user gestures to the ViewModel.

1. **Two-Way MVVM Binding with Riverpod**:
   - **State Binding (ViewModel ➔ View)**:
     - Views call `ref.watch(scheduleGenerationControllerProvider)` or `ref.watch(authControllerProvider)` inside `build()` to declaratively rebuild whenever the ViewModel emits a new state.
     - Views use `ref.listen(...)` inside `build()` for one-off UI side effects such as displaying error `SnackBar`s when `next.status == GenerationStatus.error` or `next.hasError`.
   - **Command Binding (View ➔ ViewModel)**:
     - Event callbacks (`onPressed`, `onMonthSelected`, `initState` post-frame callbacks) invoke ViewModel methods using `ref.read(provider.notifier).<command>()`.
2. **Separation of Ephemeral View State vs. ViewModel State**:
   - **ViewModel State**: Business and screen-level state (`ScheduleGenerationState`, `AuthState`) lives in Riverpod Notifiers.
   - **Ephemeral View State**: Purely visual, widget-local state stays inside `StatefulWidget` / `ConsumerStatefulWidget`:
     - `LoginScreen` and `SignupScreen` manage their own `TextEditingController`s, `GlobalKey<FormState>`, and `_obscurePassword` toggle.
     - `DraftPreviewCanvas` manages its own `_viewMode` (`CanvasViewMode.calendar` vs. `CanvasViewMode.table`) and `_selectedDay` calendar highlight while receiving the `ScheduleDraft` data and callbacks from `AdminScreen`.
     - `ExcelUploadDialog` and `GenerationModeDialog` manage modal selection state and return the user's choice (`PlatformFile` or `GenerationMode`) to `AdminScreen`, which then invokes `ScheduleGenerationController`.

---

## 4. Folder Structure

```text
frontend/
├── docs/                                               # Frontend architecture, implementation plans & guides
│   ├── architecture.md                                 # This document
│   └── ...
├── lib/
│   ├── main.dart                                       # Entry point wrapping App in ProviderScope & AuthGate
│   │
│   ├── core/                                           # Shared cross-feature infrastructure
│   │   ├── config/
│   │   │   └── app_config.dart                         # Environment-aware API_BASE_URL resolution (Web vs. Native)
│   │   ├── network/
│   │   │   ├── api_client.dart                         # Dio wrapper, Bearer injection, single-flight 401 refresh
│   │   │   └── api_exception.dart                      # Standardized ApiException and ApiErrorType enum
│   │   └── storage/
│   │       └── token_storage.dart                      # FlutterSecureStorage wrapper for access & refresh tokens
│   │
│   ├── theme/
│   │   └── app_theme.dart                              # Material 3 Light and Dark ColorSchemes & ThemeData
│   │
│   ├── shared/
│   │   └── widgets/                                    # Cross-feature UI components ("Rule of Two")
│   │       ├── bottom_nav_bar.dart                     # Role-aware Material 3 NavigationBar
│   │       └── profile_drawer.dart                     # User profile drawer with role badge & logout action
│   │
│   └── features/                                       # Self-contained MVVM feature modules
│       ├── auth/                                       # Authentication & Staff Onboarding feature
│       │   ├── data/
│       │   │   ├── datasources/
│       │   │   │   └── auth_remote_data_source.dart    # /api/v1/auth/* HTTP calls
│       │   │   ├── dtos/
│       │   │   │   ├── auth_tokens_dto.dart            # Token JSON parsing & redaction
│       │   │   │   └── user_dto.dart                   # User JSON parsing & validation
│       │   │   └── repositories/
│       │   │       └── auth_repository.dart            # AuthRepository interface & implementation
│       │   ├── domain/
│       │   │   ├── models/
│       │   │   │   ├── auth_tokens.dart                # Immutable AuthTokens entity
│       │   │   │   ├── user.dart                       # Immutable User entity
│       │   │   │   └── user_role.dart                  # UserRole enum & role capability helpers
│       │   │   └── state/
│       │   │       └── auth_state.dart                 # Sealed AuthState (Authenticated / Unauthenticated)
│       │   └── presentation/
│       │       ├── controllers/
│       │       │   └── auth_controller.dart            # AuthController (AsyncNotifier ViewModel) & selectors
│       │       ├── screens/
│       │       │   ├── login_screen.dart               # Staff login form View
│       │       │   └── signup_screen.dart              # Invitation-based registration form View
│       │       └── widgets/
│       │           └── auth_gate.dart                  # Top-level reactive auth router & Skeletonizer wrapper
│       │
│       └── admin/                                      # Admin Schedule Generation, Preview & Publishing feature
│           ├── data/
│           │   ├── datasources/
│           │   │   └── schedule_draft_remote_data_source.dart # /api/v1/schedules/* HTTP & multipart calls
│           │   ├── dtos/
│           │   │   ├── excel_roster_diff_dto.dart      # Roster diff JSON parsing
│           │   │   ├── schedule_draft_dto.dart         # ScheduleDraft & assignment JSON parsing
│           │   │   ├── schedule_summary_dto.dart       # ScheduleSummary JSON parsing
│           │   │   └── target_month_info_dto.dart      # TargetMonthInfo JSON parsing
│           │   └── repositories/
│           │       └── schedule_draft_repository.dart  # ScheduleDraftRepository interface & implementation
│           ├── domain/
│           │   └── models/
│           │       ├── excel_roster_diff.dart          # ExcelRosterDiff, MissingDoctorDiffItem, NewDoctorDiffItem
│           │       ├── schedule_draft.dart             # ScheduleDraft & ScheduleAssignment domain models
│           │       ├── schedule_summary.dart           # ScheduleSummary domain model
│           │       └── target_month_info.dart          # TargetMonthInfo domain model
│           └── presentation/
│               ├── controllers/
│               │   └── schedule_generation_controller.dart # ScheduleGenerationController (Notifier ViewModel)
│               ├── state/
│               │   └── schedule_generation_state.dart  # Immutable ScheduleGenerationState & GenerationStatus
│               ├── screens/
│               │   └── admin_screen.dart               # Main Admin workspace screen View
│               └── widgets/
│                   ├── admin_hero_card.dart            # Target month selector, status badge & generate CTA
│                   ├── admin_metric_cards.dart         # Responsive (Desktop/Tablet/Mobile) KPI metric cards
│                   ├── draft_preview_canvas.dart       # Calendar/Table toggle, shift dot grid, export & publish bar
│                   ├── excel_upload_dialog.dart        # Modal dialog for picking .xlsx workbooks
│                   └── generation_mode_dialog.dart     # Modal dialog choosing Current Roster vs. Excel Upload
│
└── test/                                               # Mirrored unit & widget test suite
    ├── core/                                           # Tests for AppConfig, ApiClient (401 refresh), TokenStorage
    ├── shared/                                         # Widget tests for BottomNavBar and ProfileDrawer
    └── features/
        ├── auth/                                       # Unit & widget tests for Auth DTOs, repo, ViewModel, screens
        └── admin/                                      # Unit & widget tests for Admin DTOs, repo, ViewModel, widgets
```

---

## 5. Core Cross-Cutting Mechanisms

### 5.1. Networking & Single-Flight Silent Token Refresh (`lib/core/network/api_client.dart`)
- **Automatic Bearer Injection**: Protected requests read the access token from `TokenStorage` and attach `Authorization: Bearer <token>`.
- **Single-Flight 401 Recovery**: If one or more concurrent requests fail with `401 Unauthorized`, `ApiClient` shares a single in-flight `_refreshFuture` calling `POST /api/v1/auth/refresh`. Once the token pair is rotated and persisted in `TokenStorage`, the original failed requests are transparently retried once. If the refresh token itself is rejected (`401`), tokens are cleared from `TokenStorage`.
- **Domain Exception Mapping**: Raw `DioException` instances never leak into ViewModels or Views; `_mapDioException()` translates HTTP status codes and socket errors into structured `ApiException` objects (`ApiErrorType.validation`, `unauthorized`, `forbidden`, `notFound`, `server`, `network`, `unknown`).

### 5.2. Declarative Auth Gate & Skeleton Loading (`lib/features/auth/presentation/widgets/auth_gate.dart`)
- `AuthGate` wraps the authenticated home screen (`AdminScreen`) at the root of the widget tree.
- While `AuthController` is asynchronously restoring the session on startup (`authState.isLoading && !authState.hasValue`), `AuthGate` renders the target screen inside a `Skeletonizer` shimmer wrapper rather than a blank spinner.
- Similarly, `AdminHeroCard`, `AdminMetricCards`, and `DraftPreviewCanvas` (using a static `_skeletonDraft` placeholder) wrap their layouts in `Skeletonizer(enabled: isLoading)` so layout geometry remains stable during data fetches.

### 5.3. Testability (`test/`)
- The `test/` directory mirrors `lib/` 1-to-1.
- Because repositories and data sources are injected via Riverpod providers, ViewModel tests (`auth_controller_test.dart`, `schedule_generation_controller_test.dart`) and Widget tests (`admin_screen_test.dart`, `login_screen_test.dart`) override `authRepositoryProvider` or `scheduleDraftRepositoryProvider` inside a `ProviderContainer` or `ProviderScope(overrides: [...])` without making real network or secure-storage calls.
