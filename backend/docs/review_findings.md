# Repository Review Findings

## Scope and limitations

This report consolidates the security, code-quality, and test-coverage
reviews performed on the current backend source tree.

- Test execution and Python compilation could not be run in the review
  environment because command execution was denied.
- Findings about coverage are based on static inspection of source and tests.
- The team day-off loop defect reported during the first pass is fixed in the
  current `src/scheduler.py`, but its regression test remains ineffective.

## Security findings

### Critical: unauthenticated, unscoped API access

**Locations:** `src/main.py`, all feature `routes.py` files, and feature
repositories.

Every API endpoint depends only on `get_session`. There is no authentication,
role check, or department scope. Any reachable caller can read doctor email
addresses and availability, enumerate departments, teams, positions, shifts,
and assignments, and alter scheduling data for any department.

**Recommendation:** Require an authenticated principal on every API route,
derive its department and role on the server, and scope every query and
mutation to that department. Verify ownership for every path and body ID.

### High: permissive credentialed CORS

**Location:** `src/main.py`

The API allows every origin, method, and header while allowing credentials.
An untrusted website can issue browser requests to the reachable API and read
responses.

**Recommendation:** Configure an explicit allowlist of trusted frontend
origins, disable credentials unless needed, and use CSRF protection if
cookie-based authentication is introduced.

### Medium: shift-assignment invariants can be bypassed

**Locations:** `src/shift/schemas.py`, `src/shift/controllers.py`, and
`src/shift/models.py`

`doctors_per_shift` accepts zero or negative values. The controller uses a
non-atomic equality check for capacity, so a negative capacity admits
unlimited assignments and concurrent requests can oversubscribe a shift. It
also does not verify that an assigned doctor belongs to the shift's
department.

**Recommendation:** Validate a positive bounded capacity, verify doctor and
shift compatibility, and enforce critical invariants transactionally and in
the database.

### Security hardening

- Set `echo=False` in `src/db/connection.py` outside development. Current log
  configuration suppresses engine output, so this is hardening rather than a
  confirmed data leak.
- When CI is introduced, use `uv sync --locked` and pin external actions and
  container images to immutable SHAs or digests.

No confirmed SQL injection, SSRF, XSS, committed-secret, unsafe
deserialization, HTTP-exposed path traversal, or dependency-lock integrity
issue was found.

## Code-quality and correctness findings

### High: invalid relation handling causes server errors and invalid data

**Locations:** `src/shift/controllers.py` and `src/doctor/controllers.py`

Shift-assignment creation dereferences a missing shift, and doctor-position
creation dereferences a missing position. Both paths can return a 500 instead
of a client error. Assignment creation also does not reject missing or
deleted doctors, cross-department doctors, or a doctor already assigned to a
different shift on the same date.

**Recommendation:** Add active-resource guards before dereferencing objects,
return explicit 4xx errors, and centralize relation validation.

### High: database invariants exist only in application code

**Locations:** `src/doctor/models.py`, `src/shift/models.py`, and controller
duplicate checks.

The association and scheduling tables have no database uniqueness constraints
for records that controllers treat as unique. Examples include
doctor/date unavailability, doctor/date pre-assignments, doctor/position
associations, and shift/date/doctor assignments. Application-only checks are
race-prone and SQLite foreign keys are not enabled by the default connection.

**Recommendation:** Add appropriate database constraints, enable SQLite
foreign-key enforcement for local use, and translate integrity errors into
clear API responses.

### High: inconsistent import roots

**Locations:** `src/scheduler.py`, `src/utils/excel.py`, and
`src/tests/test_scheduler.py`

The scheduler and spreadsheet utility use bare `models` and `scheduler`
imports, while the API uses `src.*` imports. The scheduler tests also use
bare imports despite `pyproject.toml` setting `pythonpath = "."`. A clean
package invocation or test collection can raise `ModuleNotFoundError`.

**Recommendation:** Standardize all application and test imports on `src.*`.

### Medium: configured solver timeout is ignored

**Locations:** `src/models.py`, `src/db/schemas.py`, and `src/scheduler.py`

`ScheduleConfig.solver_time_limit` is defined but never applied to the
CP-SAT solver in `ShiftScheduler.create_schedule`.

**Recommendation:** Set the solver time-limit parameter before solving and
test that configuration propagation.

### Medium: two divergent scheduling implementations

**Locations:** `src/scheduler.py` and `src/schedule_maker.py`

Both files implement scheduling dates, coverage rules, balancing, and weekend
logic with different APIs and configuration models. Maintaining both risks
behavioral drift.

**Recommendation:** Retire the legacy implementation or make it a thin CLI
wrapper around `ShiftScheduler`.

### Medium: dependency and quality-tooling drift

**Locations:** `requirements.txt`, `pyproject.toml`, `uv.lock`,
`quality_check.md`, and `excel_to_ics.py`

`requirements.txt` omits most runtime dependencies. `excel_to_ics.py` imports
`ics`, which is not declared in the project dependency metadata. The listed
quality tools are not installed by the development dependency group.

**Recommendation:** Make `pyproject.toml` and `uv.lock` the supported source
of truth, either remove or generate `requirements.txt`, add the ICS
dependency, and declare the documented quality tools as development
dependencies.

### Lower-priority cleanup

- `ShiftUpdate` uses `Optional[str]` for a numeric field and has the typo
  `grants_day_odd`; several update DTOs do not accurately model partial
  updates.
- `src/utils/excel.py` iterates through the same doctor/date grid twice and
  prints each doctor name. Combine the loops and use structured logging when
  output is needed.
- `excel_to_ics.py` writes its output, reads it back, and writes it again only
  to remove blank lines. Generate the final serialized content once.

## Test findings and missing coverage

### The team day-off test did not detect the original bug

**Location:** `src/tests/test_scheduler.py:test_max_one_day_off_team`

The test creates only one day-off-granting shift with
`doctors_per_shift=1`. That globally limits any day to one day-off assignment,
so every team necessarily has at most one even if the team-specific constraint
does not exist. It also accepts one solver-produced feasible schedule instead
of trying to prove that an invalid schedule is forbidden.

**Required regression test:** Use two teams and a day-off-granting shift with
capacity two. Force two doctors from the first team onto the same prior-day
shift, then assert that the model is infeasible. This would fail against the
previous dedented loop and pass against the current implementation.

### High-priority missing tests

1. End-to-end `create_schedule` coverage for the complete production
   constraint set and solver-time-limit propagation.
2. Scheduler invariants for one shift per doctor per day, all doctors
   pre-assigned, infeasible inputs, out-of-month or ineligible
   pre-assignments, and reuse of a scheduler instance.
3. Soft-constraint behavior tests that compare objective outcomes, rather than
   only checking that penalty variables were created.
4. Shift-assignment tests for a missing or deleted shift, a missing or deleted
   doctor, cross-department assignment, a second shift on the same date, and
   zero or negative capacity.
5. Doctor-position tests for a missing or deleted position.
6. Parent-integrity tests for a nonexistent or deleted department, team, or
   position when creating dependent resources.
7. Route coverage for the unfinished department and team list tests, plus the
   health endpoint.
8. Unit tests for Excel export, ICS generation, and structured logging.
9. After authentication is implemented, 401/403, cross-department isolation,
   and CORS allowlist tests.

### Test-suite maintainability

The in-memory SQLModel session fixture and FastAPI dependency-override client
fixture are repeated in department, team, doctor, position, and shift tests.
Doctor, shift, and unavailability builders are also duplicated.

**Recommendation:** Move common fixtures to `backend/conftest.py`, move model
builders to shared test factories, and replace broad
`pytest.raises(Exception)` assertions with exact exception types. Complete the
two TODO list-route tests before adding further feature tests.

## Recommended remediation order

1. Add authentication, authorization, department scoping, and restrictive
   CORS.
2. Fix assignment validation and enforce database invariants.
3. Standardize imports and establish a runnable test baseline.
4. Add the scheduler regression and end-to-end tests.
5. Honor solver configuration and consolidate scheduling implementations.
6. Modularize fixtures and align dependency and quality-tool metadata.
