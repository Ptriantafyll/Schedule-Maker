# UUIDv7 Primary Key Architecture & Transition Guide

## 1. Overview & Purpose

Schedule-Maker uses **UUIDv7** ([RFC 9562](https://www.rfc-editor.org/rfc/rfc9562)) as the default primary key generation algorithm for all database models.

This document serves as the permanent reference for:
1. Why UUIDv7 was chosen over UUIDv4.
2. How the dependency was added using `uv`.
3. How `SyncBase` centralizes primary key generation across all tables.
4. Backward compatibility with existing records.
5. The future migration path to native Python 3.14+.

---

## 2. Rationale: UUIDv4 vs. UUIDv7

| Metric / Behavior | UUIDv4 (Legacy) | UUIDv7 (Current) |
| :--- | :--- | :--- |
| **Structure** | 122 bits pseudo-random | 48-bit millisecond Unix timestamp + 74 bits randomness |
| **Index Locality** | **Fragmented**: Random inserts cause frequent B-Tree page splits and random I/O in SQLite & PostgreSQL | **Sequential**: Appends monotonically to the right leaf of the B-Tree index, maximizing buffer cache hits |
| **Sorting** | Cannot be sorted by creation time | Naturally sortable (`ORDER BY id ASC/DESC` matches creation order) |
| **Sync Conflict Resolution** | Relies purely on `updated_at` / `created_at` timestamp comparison | Timestamp embedded directly in the 128-bit key |
| **Collision Resistance** | Cryptographically strong ($2^{122}$) | Extremely strong ($2^{74}$ per millisecond) |

---

## 3. Dependency Management with `uv`

Because the runtime is **Python 3.11.9** (where the standard library `uuid` module does not yet include `uuid.uuid7`), we use the zero-dependency, pure-Python library **`uuid6`** from PyPI.

### Procedure: Adding `uuid6`
To add the package to the project:
```bash
uv add uuid6
```

### Resulting Files:
1. **`pyproject.toml`**:
   ```toml
   dependencies = [
       ...,
       "uuid6>=2025.0.1",
   ]
   ```
2. **`uv.lock`**: Pinned exact package version and cryptographic wheel hash.
3. **`.venv/`**: Installed automatically into the local virtual environment.

---

## 4. Single Source of Truth: `SyncBase`

All 14 database models in Schedule-Maker inherit from `SyncBase` in [`src/db/schemas.py`](file:///C:/Users/ptria/source/repos/Schedule-Maker/backend/src/db/schemas.py):

```python
import uuid
import datetime
from sqlmodel import SQLModel, Field
from uuid6 import uuid7


class SyncBase(SQLModel):
    """Abstract base class providing synchronization metadata for all tables."""
    id: uuid.UUID = Field(default_factory=uuid7, primary_key=True)
    created_at: datetime.datetime = Field(
        default_factory=lambda: datetime.datetime.now(datetime.timezone.utc)
    )
    updated_at: datetime.datetime = Field(
        default_factory=lambda: datetime.datetime.now(datetime.timezone.utc)
    )
    is_deleted: bool = Field(default=False)
    sync_status: bool = Field(default=False)
```

### Affected Models:
- **Core Entities**: `User`, `Department`, `Doctor`, `Position`, `Shift`, `Team`
- **Junction & Association Tables**: `DoctorPosition`, `DoctorUnavailability`, `DoctorPreAssignment`, `ShiftAssignment`
- **Application State**: `ScheduleDraft`, `ScheduleConfig`, `Invitation`, `RefreshSession`

---

## 5. Backward Compatibility & Database Migrations

- **Zero Alter-Table Migrations**: Both UUIDv4 and UUIDv7 share the exact same 128-bit binary representation. In SQLite, both store as `CHAR(36)`; in PostgreSQL, both store as native `UUID`.
- **Existing Records**: Existing rows generated with UUIDv4 remain 100% valid and queryable.
- **Client Transport (Flutter / Web / Native)**: The frontend receives UUID strings in standard 8-4-4-4-12 hex format (e.g. `01924b10-6789-7000-8000-0123456789ab`). Dart `Uuid` parsers and string decoders handle UUIDv7 transparently without modification.

---

## 6. Future Path: Python 3.14+ Native Deprecation

Python 3.14 adds native `uuid.uuid7` to the standard library ([RFC 9562](https://www.rfc-editor.org/rfc/rfc9562)).

When the project eventually upgrades to Python 3.14 (once Google OR-Tools and compiled C/Rust packages have mature pre-built wheels for 3.14), replacing `uuid6` is a single line change in [`src/db/schemas.py`](file:///C:/Users/ptria/source/repos/Schedule-Maker/backend/src/db/schemas.py):

```python
# Change from:
from uuid6 import uuid7

# To:
from uuid import uuid7
```

Followed by removing the dependency:
```bash
uv remove uuid6
```
