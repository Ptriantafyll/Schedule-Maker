# GEMINI.md

## Project Overview

Schedule-Maker Frontend is a multi-platform Flutter application (Mobile, Web, Desktop) providing an intuitive interface for hospital administrators, doctors, and viewers. It enables users to view monthly on-duty schedules, request leaves and preferences, manage department resources, and trigger solver schedule generation via the FastAPI backend API.

## Architecture

The frontend follows a **Feature-First Clean Architecture** with **unidirectional data flow**:

- **Framework**: **Flutter** (Dart `^3.12.0`) with **Material 3** design system.
- **State Management**: **Flutter Riverpod** (`flutter_riverpod: ^2.6.1`) for compile-safe, declarative, and testable dependency injection and state management.
- **HTTP Client**: **Dio** (`dio: ^5.11.0`) configured with custom interceptors for:
  - Automatic `Authorization: Bearer <access_token>` injection.
  - Silent token refresh interceptor on `401 Unauthorized` with request queueing.
  - Dual transport mode handling (Secure Storage on native platforms; HttpOnly cookies + CSRF tokens on Web).
- **Secure Storage**: **flutter_secure_storage** for hardware-backed encryption (Android Keystore / iOS Keychain) of long-lived refresh tokens.
- **Date & Calendar Formatting**: **intl** (`intl: ^0.20.3`) for localized calendar and shift dates.

---

## Folder Structure (Feature-First Clean Architecture)

The codebase is organized by **feature** rather than by technical layer. Each feature encapsulates its own presentation, domain, and data layers:

```text
frontend/
├── lib/
│   ├── core/                        # Shared infrastructure across all features
│   │   ├── config/                  # App environment, base URLs (app_config.dart)
│   │   ├── network/                 # Dio client, interceptors, error mappings (api_client.dart)
│   │   ├── storage/                 # Secure hardware-backed storage (token_storage.dart)
│   │   ├── exceptions/              # Standardized failure models (AppFailure, NetworkException)
│   │   └── utils/                   # Date formatters, validators, extensions
│   │
│   ├── common_widgets/              # Shared UI components (buttons, text fields, loaders, dialogs)
│   │
│   ├── theme/                       # Material 3 color palettes, typography, theme data (app_theme.dart)
│   │
│   ├── features/                    # Self-contained feature modules
│   │   ├── auth/                    # Authentication & session management
│   │   │   ├── data/                # AuthRepository, remote API data sources, DTOs
│   │   │   ├── domain/              # User entity, auth token models
│   │   │   └── presentation/        # LoginScreen, SignupScreen, AuthController (Riverpod Notifier)
│   │   │
│   │   ├── schedule/                # Monthly calendar views & shift assignments
│   │   │   ├── data/                # ScheduleRepository, Solver API calls
│   │   │   ├── domain/              # Schedule, ShiftAssignment entities
│   │   │   └── presentation/        # CalendarView, ShiftCard, ScheduleController
│   │   │
│   │   ├── doctor/                  # Doctor profiles, unavailability, preferences
│   │   │   ├── data/                # DoctorRepository
│   │   │   ├── domain/              # Doctor entity, Unavailability model
│   │   │   └── presentation/        # DoctorListScreen, AvailabilityPicker
│   │   │
│   │   └── department/              # Department settings, positions, shifts
│   │       ├── data/                # DepartmentRepository
│   │       ├── domain/              # Department, Position, Shift models
│   │       └── presentation/        # DepartmentSettingsScreen, ShiftListWidget
│   │
│   └── main.dart                    # App entry point wrapped with ProviderScope
│
├── test/                            # Mirrored test hierarchy
│   ├── core/                        # Tests for network interceptors, storage, formatters
│   └── features/                    # Feature unit, widget, and provider tests
│       ├── auth/
│       └── schedule/
└── pubspec.yaml
```

---

## Best Practices

### 1. Layers within Each Feature (Unidirectional Data Flow)
Inside each feature, dependencies flow strictly inward: `Data` ➔ `Domain` ➔ `Presentation`.

1. **Domain Layer (`domain/`)**:
   - Pure Dart only: zero Flutter UI dependencies.
   - Contains immutable business entities and value objects.
2. **Data Layer (`data/`)**:
   - **Data Sources**: Perform low-level network calls using `Dio`.
   - **Repositories**: Abstract data access and return clean domain entities to controllers.
   - **DTOs**: Handle `fromJson` and `toJson` serialization.
3. **Presentation Layer (`presentation/`)**:
   - **Widgets / Screens**: Stateless or `ConsumerWidget`s that render UI and forward user gestures to controllers.
   - **Controllers**: Implement `AsyncNotifier` or `Notifier` from Riverpod, managing UI state as `AsyncValue<T>` (automatically handling Data, Loading, and Error states).

### 2. Riverpod 2.x Best Practices
- **Use `Notifier` / `AsyncNotifier`**: Avoid legacy `ChangeNotifier` and `StateNotifier`.
- **Co-locate Providers**: Place provider definitions inside their respective feature folders (e.g. `auth_controller.dart` defines `authControllerProvider`), rather than accumulating a massive global providers file.
- **`ref.watch` vs `ref.read`**:
  - Use `ref.watch` inside `build()` to rebuild widgets reactively when state changes.
  - Use `ref.read` only inside event callbacks (e.g. `onPressed`) to invoke methods without creating rebuild subscriptions.
- **Dependency Injection**: Provide repositories and HTTP clients via Riverpod providers to make mocking straightforward in tests (`ProviderScope(overrides: [...])`).

### 3. Networking & Security
- **Silent Refresh Queue**: When multiple concurrent requests fail with `401 Unauthorized`, queue subsequent requests and trigger only a single `/refresh` call.
- **Dual Transport Security**: Support hardware-backed Keystore/Keychain storage for mobile/desktop, and `HttpOnly` cookies + CSRF tokens for web.
- **Standardized Error Handling**: Never expose raw `DioException` directly to the UI. Map HTTP errors into structured, user-friendly domain failures (`AppFailure`).

---

## Preferences

- When making changes don't paste all the code at once. Instead go step by step in small chunks of code explaining the process each time.
- When the user is learning and doing most of the coding themselves, explain: (1) why we are making a decision, (2) what the best practices are and (3) if there is a new feature that we haven't touched explain how it works.
- When the user stops and corrects a suggestion or says they don't like something, add that preference to this GEMINI.md file.
- After every change, check if anything can be made cleaner and if there is repeated code that can be extracted into a reusable widget or helper function.
- Don't do everything by yourself, the user wants to write most of the code by themselves. Only write code if you were specifically asked to.
- Do not write code suggestions first. Instead, explain the high-level logic, requirements, or design first, let the user think and write the code themselves, and then review it.
- Always try to do the simplest and most minimal solution.
- Always document any new activity or procedure in a new document in the `docs/` folder.
- Always ask the user for approval when adding or changing a file.
- Always build with future scalability, responsive layout adaptation (Mobile vs Tablet/Desktop/Web), and code readability in mind.
- Maintain a TDD mindset: write widget and unit tests for providers and repositories alongside feature implementations.
- Always work feature-by-feature: create a clear, step-by-step implementation plan for each feature before writing code, and execute it incrementally step by step.
