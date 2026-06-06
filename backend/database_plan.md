# Hospital Shift Scheduler: Database Architecture & Sync Plan

1. **Executive Summary**
   This document outlines the hybrid, offline-first database strategy for the Hospital Shift Scheduling application. To ensure maximum uptime inside the hospital (even during outages) while still allowing doctors to view schedules via mobile/web apps, the system employs a dual-database layout linked by a secure REST API synchronization layer.

2. **Dual-Database Architecture**
   
   **Local Desktop Environment (Admin App)**
   * **Engine:** SQLite
   * **Storage:** Embedded single-file (`hospital_schedule.db`) inside the local application data directory.
   * **Concurrency:** Single-user write access (optimized for the scheduling administrator).
   * **Dependency:** Zero-installation required; completely standalone.

   **Central Cloud Environment (Web/Mobile Apps)**
   * **Engine:** PostgreSQL
   * **Storage:** Managed cloud instance or dedicated on-premise hospital server database.
   * **Concurrency:** Multi-user read/write access (optimized for hundreds of concurrent doctor requests).
   * **Dependency:** Requires internet connectivity to update or read.

3. **Core Synchronization Principles**

   * To prevent data collision and corruption between offline clients and the cloud server, all database models must adhere to the following three design constraints:
   * Global Unique Identifiers (UUIDv4): Classic integer primary keys (1, 2, 3...) are strictly banned. If two local machines create a new doctor or shift while offline, integer IDs will collide during synchronization. All tables must use UUID strings generated on the client-side.
   * State and Sync Auditing: Every table must contain flags tracking synchronization state and modification windows.
   * sync_status (Boolean): Set to False on local creation/modification; switched to True only after successful server acknowledgment.
   * updated_at (DateTime): Tracks exactly when a record was last modified to resolve version conflicts.
   * Soft Deletes (is_deleted): Dropping a row from SQLite while offline leaves the cloud server unaware of the deletion. Records are given an is_deleted flag. The sync engine sends this flag to the server, and the server handles the actual purge or filters it from active view models.

4. **Unified Data Schema (SQLModel Framework)**
   By utilizing Python's SQLModel, the exact same source code models are used to construct the local SQLite tables and the remote PostgreSQL tables without structural modifications.

   ```python
   import uuid
   from datetime import datetime
   from typing import Optional
   from sqlmodel import Field, SQLModel

   class SyncBase(SQLModel):
      """Abstract base class tracking synchronization states."""
      id: uuid.UUID = Field(default_factory=uuid.uuid4, primary_key=True, index=True)
      created_at: datetime = Field(default_factory=datetime.utcnow)
      updated_at: datetime = Field(default_factory=datetime.utcnow)
      is_deleted: bool = Field(default=False)
      sync_status: bool = Field(default=False)

   class Department(SyncBase, table=True):
      name: str = Field(index=True, unique=True)
      code: str

   class Position(SyncBase, table=True):
      title: str
      required_certifications: Optional[str] = None

   class Doctor(SyncBase, table=True):
      name: str
      email: str
      max_consecutive_shifts: int = 5
      department_id: uuid.UUID = Field(foreign_key="department.id")
      position_id: uuid.UUID = Field(foreign_key="position.id")

   class Shift(SyncBase, table=True):
      date: str  # Format ISO: YYYY-MM-DD
      shift_type: str  # e.g., "Morning", "Night", "On-Call"
      department_id: uuid.UUID = Field(foreign_key="department.id")
      doctor_id: Optional[uuid.UUID] = Field(default=None, foreign_key="doctor.id")
   ```

5. **The Sync Workflow**
   [Offline Activity]          [Network Connection Restored]            [Cloud Persistence]
   ┌─────────────────┐         ┌───────────────────────────┐         ┌──────────────────────┐
   │ Admin modifies  │ ──────► │ Background task loops     │ ──────► │ Central API validates│
   │ Shift in local  │         │ local records where       │         │ incoming payloads via│
   │ SQLite DB.      │         │ sync_status == False      │         │ unique UUID keys.    │
   └─────────────────┘         └───────────────────────────┘         └──────────────────────┘
             │                                                                   │
             ▼                                                                   ▼
   ┌─────────────────┐                                               ┌──────────────────────┐
   │ Record saved;   │                                               │ PostgreSQL commits   │
   │ sync_status=False│                                              │ changes; returns HTTP│
   └─────────────────┘                                               │ 200 OK success state.│
                                                                     └──────────────────────┘
                                                                                 │
                                                                                 ▼
                                                                     ┌──────────────────────┐
                                                                     │ Local loop catches   │
                                                                     │ 200 OK and updates   │
                                                                     │ local sync_status=True│
                                                                     └──────────────────────┘
   Sync Strategy Details:
   Upstream Synchronization: The local background scheduler aggregates modified rows, batches them into single payloads, and performs a payload swap with the server endpoint (POST /api/v1/sync).

   Downstream Synchronization: On launch or manually triggered intervals, the client queries GET /api/v1/sync?since={last_successful_sync_timestamp} to catch up with changes made via the web or other administrative clients.

6. **Phased Implementation Roadmap**
   * Phase 1: Local Setup. Migrate current plain-text or script constraints into local SQLite models using SQLModel. Update the current schedule-generation algorithm script to pull dynamically from database sessions.
   * Phase 2: Local UI Integration. Construct the Flutter desktop shell and wire up basic UI pages interacting with the local SQLite data engine via local FastAPI routes.
   * Phase 3: Centralization. Spin up a target cloud platform instance running PostgreSQL. Deploy a mirror copy of the FastAPI layer to handle central authenticated traffic.
   * Phase 4: Synchronization Engine. Build the automatic background sync loop linking client-side SQLite targets to server-side Postgres containers.
