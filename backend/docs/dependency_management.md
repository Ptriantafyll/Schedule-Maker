# Dependency Management & Web Architecture

This document describes the libraries chosen for the **Schedule-Maker** backend API, the rationale behind their selection, and step-by-step instructions on how to manage dependencies using the `uv` tool.

---

## 1. Core Libraries Overview

For building our backend API and database layer, we use three core libraries:

| Library | Rationale | Best Practice |
| :--- | :--- | :--- |
| **FastAPI** | Extremely fast web framework for building APIs with automatic OpenAPI docs (`/docs`). | Enforces strict input validation using Pydantic, preventing dirty data from hitting the database. |
| **Uvicorn** | An ASGI web server designed to serve asynchronous Python frameworks like FastAPI. | Runs FastAPI application instances locally with optional auto-reload for development. |
| **SQLModel** | A unified library combining SQLAlchemy and Pydantic. | Eliminates duplicate code by allowing one class definition to serve as both the database schema and the API data-transfer object (DTO). |

---

## 2. Managing Dependencies with `uv`

The project uses `uv` for environment and package management. It is a extremely fast Rust-based replacement for `pip`, `pip-tools`, and `virtualenv`.

### Procedure: Adding a New Dependency

Whenever you need to add a library to the project, follow these steps:

#### 1. Add a Production Dependency

To add a library required for the application to run (e.g., `fastapi`, `sqlmodel`):

```bash
uv add <package-name>
# Example:
uv add fastapi uvicorn sqlmodel
```

* **What this does:**
  * Adds the package and its version constraints to the `dependencies` array in `pyproject.toml`.
  * Resolves all sub-dependencies and updates `uv.lock` to lock specific versions.
  * Installs the package into the virtual environment (`.venv`).

#### 2. Add a Development Dependency

To add a library required only for testing, linting, or building (e.g., `pytest`, `ruff`):

```bash
uv add --dev <package-name>
# Example:
uv add --dev pytest
```

* **What this does:**
  * Adds the package to the `dev` dependency group under `[dependency-groups]` in `pyproject.toml`.
  * Ensures development tools are not bundled in a production distribution.

#### 3. Sync the Environment

If you pull down code with changes in `pyproject.toml`, synchronize your local `.venv` by running:

```bash
uv sync
```

---

## 3. Key Concepts

### Pydantic & SQLAlchemy Combined (SQLModel)

Traditionally, Python SQL web applications require duplicate model code:

1. **SQLAlchemy Table Model:** Defines columns, keys, and database-level constraints.
2. **Pydantic Schema Model:** Defines request payloads, response payloads, types, and validators.

With **SQLModel**, you write a single class:

```python
from sqlmodel import SQLModel, Field
import uuid

class Doctor(SQLModel, table=True):
    id: uuid.UUID = Field(default_factory=uuid.uuid4, primary_key=True)
    name: str
    email: str
```

By adding `table=True`, SQLModel registers this class as a database table internally using SQLAlchemy. At the same time, because it inherits from `SQLModel` (which inherits from Pydantic's `BaseModel`), it acts as a schema for FastAPI request/response validation.

### Deterministic Builds (pyproject.toml vs uv.lock)

* **`pyproject.toml`:** Declares *abstract* dependencies and user preferences (e.g., `sqlmodel>=0.0.14`). It is human-readable and intended to be edited directly.
* **`uv.lock`:** Declares the *exact* resolved dependency tree and cryptographic hashes of every installed package. It ensures that every machine (local dev, CI pipeline, remote cloud server) installs the exact same dependencies byte-for-byte. **Never edit this file manually.**
