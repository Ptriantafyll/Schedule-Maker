# Frontend Backlog

Deferred issues that should be addressed outside the current implementation step.

## Open

### BL-001: Self-Service Password Reset Flow (Full-Stack)

**Area:** Authentication / Self-Service Account Recovery  
**Priority:** Medium

#### Problem

Currently, hospital doctors and staff who forget their login credentials must contact an administrator or the IT Help Desk (Ext. 4420) to have their credentials reset manually. While the Login screen displays a Help Desk modal upon tapping *"Forgot Password?"*, an automated, self-service recovery mechanism is necessary for production scaling and operational convenience across hospital shifts.

Because password reset requires cryptographically secure, time-limited single-use reset tokens and automated email delivery, implementing this capability requires coordinated full-stack changes across both the FastAPI backend and the Flutter frontend.

#### Required work

1. **Backend Implementation**:
   - Create reset token model or signed cryptographic token generator (e.g. `itsdangerous` or hashed database tokens with strict expiration, e.g. 15–30 minutes).
   - Create request endpoint: `POST /api/v1/auth/forgot-password` accepting user email. Always return a generic 200/202 success message to prevent user enumeration attacks.
   - Integrate with email dispatch service (see backend backlog `BL-002`) to email the reset link containing the token.
   - Create confirmation endpoint: `POST /api/v1/auth/reset-password` accepting token and new password. Validates token validity/expiry, updates password hash in database, invalidates existing user refresh tokens/sessions, and revokes the used token.
   - Pytest unit and integration test suite covering token generation, valid reset, expired token, reused token, and invalid email.

2. **Frontend Implementation**:
   - **Data Layer**:
     - Add `forgotPassword(String email)` and `resetPassword({required String token, required String newPassword})` to `AuthRemoteDataSource` and `AuthRepository`.
   - **Presentation Layer**:
     - Connect *"Forgot Password?"* on `LoginScreen` to open a "Reset Password" dialog or dedicated screen.
     - Email request step: input email with validation, submitting to `forgotPassword`.
     - Token & New Password step: form input with password visibility toggle and confirmation password validation.
     - Provide clear feedback to the user on completion and route back to `LoginScreen`.
   - **Widget and Controller Tests**:
     - Comprehensive unit and widget tests covering input validation, submission states, and error handling.

#### Completion criteria

- Staff members can independently request a password reset via email.
- Password reset tokens are single-use, cryptographically secure, and expire after 15–30 minutes.
- User enumeration is prevented by returning generic responses on reset requests.
- Passwords can be successfully reset, allowing login with the new credentials.
- All backend and frontend test suites pass with 100% verification.

---

### BL-002: Dynamic Admin Dashboard Metrics Calculation (Days Left, Live Badges & Capacities)

**Area:** Admin / Dashboard State Management  
**Priority:** Medium (Scheduled for Phase 4 Integration)

#### Problem

The `AdminMetricCards` presentation widget accepts constructor parameters (`dueDate`, `daysLeft`, `submissionStatus`, `pendingApprovalsCount`, `availableStaffCount`, `bottleneckDaysCount`, etc.). While this allows clean UI isolation and testing, in a production hospital environment these metrics must be dynamically calculated in real time:
1. **Countdown & Deadlines**: Calculate `daysLeft = targetDeadline.difference(DateTime.now()).inDays` and update status pills dynamically (e.g. `Overdue`, `Closing Soon`, `Requests Closed`).
2. **Pending Approvals**: Query database/API for live unapproved time-off requests and show action-required warnings if count > 0.
3. **Staff Pool & Workload**: Query total active doctors and required shift slots to compute `~X.X shifts/doctor` ratio.
4. **Coverage Bottlenecks**: Run pre-flight constraint checks against doctor unavailabilities to flag dates where available doctors < required shift slots.

#### Required work

1. **Domain Model (`AdminDashboardMetrics`)**:
   - Model containing computed metric values and status enums.
2. **ViewModel / State Management (`adminDashboardMetricsProvider`)**:
   - Riverpod `AsyncNotifier` or `FutureProvider` that reads `DateTime.now()` and combines data from doctor roster and requests repositories.
   - Computes deadline differences, urgency badge colors, and bottleneck dates.
3. **AdminScreen Integration**:
   - Connect `AdminScreen` to watch `adminDashboardMetricsProvider` and pass computed properties into `AdminMetricCards`.

---

### BL-003: Comprehensive Manual Administration Flow (Doctors, Positions, Shifts, Unavailabilities & Pre-Assignments)

**Area:** Admin / Manual Roster & Operational Configuration  
**Priority:** Medium (Deferred to Backlog; Excel Workbook Generation is current priority)

#### Problem

The backend solver engine (`backend/src/scheduler.py`) requires a fully structured operational model to construct the Google OR-Tools CP-SAT constraint problem:
1. **Positions**: Clinical roles (e.g. ER, ICU, Attending, Resident) with specific operational `duty_days` (days of the week they operate).
2. **Shifts**: Shift definitions per position (e.g. Morning, Night, 24h On-Call), including `doctors_per_shift` coverage counts and `grants_day_off` post-shift rest rules.
3. **Doctors**: Staff profiles with full name, unique email, assigned/eligible positions, and optional team affiliations (`Team`).
4. **Unavailabilities**: Specific calendar dates where a doctor cannot be scheduled for duty.
5. **Pre-Assignments**: Hard-locked `(date, shift)` assignments confirmed before running the solver.

Currently, the in-app dashboard does not have interactive management interfaces to manually author, edit, or configure these entities. Constructing this full management suite requires several nested CRUD screens/dialogs. To deliver rapid scheduling capability to hospital administrators, this manual flow is captured here in the backlog while development focuses on **Phase 4: Excel Workbook Generation**.

#### Required Work

1. **Positions & Shifts Management**:
   - UI views to create/edit positions and define duty days.
   - Child forms to configure shifts per position (`doctors_per_shift`, `grants_day_off`).
2. **Doctor Roster Management**:
   - Form dialog to add/edit doctors and map them to positions and optional teams.
3. **Unavailability & Pre-Assignment Matrix**:
   - Visual date picker / matrix to input doctor unavailable dates.
   - Quick-assignment tool to lock in pre-assigned shifts.
4. **Backend Solver Routing for In-App Roster**:
   - Mount `POST /api/v1/schedules/generate-roster` in FastAPI backend consuming the relational database entities.

#### Completion Criteria

- Administrators can manually configure their entire departmental roster, positions, shifts, unavailabilities, and pre-assignments through the app.
- In-app roster solver produces valid CP-SAT assignments respecting all configured constraints.

