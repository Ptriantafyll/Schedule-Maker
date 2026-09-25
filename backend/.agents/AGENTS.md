# GEMINI.md

## Project Overview

Schedule-Maker is a Python tool that generates optimized monthly on-duty schedules for hospital doctors. It uses Google OR-Tools CP-SAT constraint programming solver to produce fair, balanced duty assignments while respecting doctor unavailability and quality-of-life preferences.

## Architecture

The project follows a hybrid, decoupled architecture designed to support both local offline scheduling and central cloud-based access:

- **Web Framework**: **FastAPI** serves as the API backend framework, providing fast, asynchronous endpoints with automatic OpenAPI documentation.
- **Server**: **Uvicorn** is used as the ASGI server to run and serve the FastAPI application.
- **ORM / Data Validation**: **SQLModel** (combining SQLAlchemy and Pydantic) maps database tables to Python objects and enforces API request/response validation.
- **Databases**:
  - **SQLite**: Used for the local desktop administration environment. A single-file database (`hospital_schedule.db`) supports standalone, zero-installation operations.
  - **PostgreSQL**: Used for the central cloud/server environment to handle concurrent requests from web and mobile users.
- **Solver Engine**: A dedicated Python component utilizing **Google OR-Tools CP-SAT** to model and solve the shift scheduling constraints.
- **Synchronization**: A sync mechanism utilizing UUIDv4 identifiers, update timestamps (`updated_at`), soft deletes (`is_deleted`), and sync tracking flags (`sync_status`) keeps the local SQLite database aligned with the central PostgreSQL instance.

## Folder Structure

The folder structure is in docs/structure.md

## Preferences

- When making changes don't paste all the code at once. Instead go step by step in small chunks of code explaining the process each time
- When the user is learning and doing most of the coding themselves, explain: (1) why we are making a decision, (2) what the best practices are and (3) if there is a new feature that we haven't touched explain how it works
- When the user stops and corrects a suggestion or says they don't like something, add that preference to this AGENTS.md file
- After every change, check if anything can be made cleaner and if there is repeated code that can be extracted into a helper function
- Use TDD (Test-Driven Development) approach: write tests first, then implement the code to make them pass
- Don't do everything by yourself, the user wants to write most of the code by themselves. Only write code if you were specifically asked to.
- Do not write code suggestions first. Instead, explain the high-level logic, requirements, or design first, let the user think and write the code themselves, and then review it.
- Always try to do the simplest and most minimal solution.
- Always document any new activity or anything that will be revisited in a new document in the docs/ folder. E.g. things that need to be documented are how to add a new library with uv, what the architecture is, how to deploy locally etc. The user wants to have a clear step by step document for every procedure. Do this whenever there is a new procedure or if you are not sure ask the user whether to include something in a new doc.
- Always ask the user for approval when adding or changing a file
- Always build with future scalability and code readability in mind
- Always record things that need to be done for future scalability (such as distributed rate limiting with Redis for multi-server or Kubernetes deployments) in docs/backlog.md
- Always work feature-by-feature: create a clear, step-by-step implementation plan for each feature before writing code, and execute it incrementally step by step
- Verify actual backend code before designing client contracts: Never assume generic REST conventions or rely on older design docs. Always inspect the live backend routes.py and Pydantic schemas.py to confirm exact request field names, HTTP status codes, and response bodies
- Enforce Backend Readiness Gates: Distinguish between mounted backend endpoints (registered in src/main.py) and planned/unmounted routes. If a backend capability does not exist yet (e.g. schedule publishing or pending-request approvals), build against explicit test fixtures and mark the live integration blocked at the gate. Never connect a UI action to an unrelated or premature endpoint
- Avoid premature UI abstractions ("Rule of Two"): Keep form fields, buttons, and layout widgets owned by their feature using standard Material 3 widgets first. Extract a component into common widgets only after at least two real features demonstrate the need for the exact same behavior
- Account for platform transport differences and concurrency: Always design networking with dual-transport awareness (Native uses secure hardware storage; Web uses HttpOnly cookies + CSRF). Ensure sensitive asynchronous operations (like token refresh) use single-flight execution (sharing one in-flight Future) to prevent race conditions during concurrent 401 failures
- Enforce tenant and role boundaries in navigation: Account for role shape differences (e.g. a Super Admin has department_id = null and cannot access departmental rosters or schedules) so routing directs users to their authorized workspace
- For the presentation layer (UI screens, dialogs, forms, layout widgets), always ask the user for the design first and ask clarifying questions before writing code or proposing UI designs.

