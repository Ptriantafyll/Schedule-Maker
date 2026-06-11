# Feature-based API Project Structure

This repository uses a feature-based layout: each feature (domain area) owns its HTTP routes, handlers, persistence, models, and feature-specific tests. That keeps related code co-located, improves discoverability, and makes large projects easier to reason about.

## Principles

- Co-locate everything for a single feature under `src/<feature>/` (or `features/<feature>/`).
- Prefer small files with single responsibility (e.g., `routes.py` only declares routes).
- Keep clear layer boundaries inside each feature: routes -> controllers -> services -> repository -> models/schemas.
- Share truly cross-cutting code (DB connection, auth, shared models) under `src/shared/` or `src/common/`.

## Top-level layout (opinionated)

- `src/` — application code
  - `src/<feature>/` — feature package (e.g. `department`, `schedules`, `users`)
    - `routes.py` — router / HTTP endpoint wiring
    - `controllers.py` — thin handlers / adapters to services
    - `services.py` (optional) — business logic for the feature
    - `repository.py` — persistence access for the feature
    - `models.py` — domain / ORM models local to the feature
    - `schemas.py` — Pydantic / input/output schemas for HTTP/API
    - `tests/` — unit + feature tests (mocks, fixtures)
  - `src/shared/` or `src/common/` — DB connection, migration helpers, shared types
- `docs/` — architecture notes and API guidance
- `tests/` — higher-level integration tests and test utilities
- `scripts/` — automation and developer helpers

## Example: `department` feature

This repository contains a `department` feature at `src/department/` with the following files:

- `src/department/routes.py` — HTTP routes for departments (create, list, get)
- `src/department/controllers.py` — request adapters and response mapping
- `src/department/repository.py` — DB access for departments
- `src/department/models.py` — department ORM model(s)
- `src/department/schemas.py` — request/response schemas for department endpoints

See the actual routes implementation here: [src/department/routes.py](src/department/routes.py#L1-L200)

## Models & Schemas: one file vs per-feature

- Recommendation: prefer per-feature `models.py` and `schemas.py` when possible. Benefits:
  - Encapsulation: feature code is self-contained and easier to refactor or extract.
  - Reduced merge conflicts in large teams working on different features.
  - Easier tests and smaller refactors when a feature's domain changes.
- When to keep a single shared file:
  - You have many small features that share the same core types and splitting adds noise.
  - There are genuinely global domain types used across many features (put those in `src/shared/models.py`).
- Hybrid approach (recommended for medium projects): per-feature models/schemas plus a `src/shared/` package for cross-cutting types.

## File & layering conventions (feature-aware)

- Routes: wire HTTP -> controller only. No business logic.
- Controllers: validate and normalize HTTP input, call `services` and convert errors to HTTP responses.
- Services: contain the feature's business rules and depend on repository interfaces.
- Repository: implement DB access; keep SQL/ORM code here.

## Testing guidance

- Unit tests inside `src/<feature>/tests/` mock repository interfaces or DB sessions.
- Integration tests that span multiple features live in the top-level `tests/` folder.

## Migration guidance

- Keep DB schema and migration scripts centralized (e.g. `src/shared/db/` or `migrations/`) and have repositories reference the shared models where appropriate.

## Short note on transitioning

- When moving from a layer-based to feature-based layout, migrate one feature at a time and keep shared utilities stable. Start by creating `src/<feature>/` folders and moving files for a single feature, update imports, and run tests.
