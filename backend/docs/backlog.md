# Backend Backlog

Deferred issues that should be addressed outside the current implementation step.

## Open

### BL-001: Enable SQLite foreign-key enforcement

**Area:** Database integrity  
**Priority:** High

#### Problem

The SQLModel models declare foreign keys, but SQLite does not enforce them
unless `PRAGMA foreign_keys=ON` is enabled for every database connection.
The current database engine does not enable this setting.

This can allow orphaned records, such as a user whose `department_id` does not
reference an existing department. It can also make development and tests behave
differently from production databases that enforce foreign keys by default.

#### Required work

1. Enable `PRAGMA foreign_keys=ON` for every SQLAlchemy SQLite connection.
2. Keep the configuration conditional so it is not executed for other database
   engines.
3. Add a regression test proving that a nonexistent foreign-key value raises an
   integrity error.
4. Add a test proving that nullable foreign keys still accept `NULL`, including
   the tenantless super-admin case.

#### Completion criteria

- Every application and test SQLite connection reports foreign-key enforcement
  as enabled.
- SQLite rejects records that reference nonexistent parent records.
- Nullable foreign-key columns continue to accept `NULL`.
- The existing test suite remains green.
