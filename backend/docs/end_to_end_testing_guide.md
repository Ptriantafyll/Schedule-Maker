# End-to-End Testing Guide: Excel Schedule Generation

This document outlines the step-by-step procedure to test the schedule generation flow end-to-end between the FastAPI backend and Flutter admin frontend.

---

## 1. Prerequisites & Environment

- **Backend**: Python 3.12+ managed with `uv`.
- **Database**: Local SQLite database (`hospital_schedule.db`).
- **Frontend**: Flutter SDK with Riverpod state management.
- **Port Configuration**: Backend running at `http://localhost:8000`, Flutter app configured to point to it (default in `AppConfig.apiBaseUrl`).

---

## 2. Test Credentials & Seed Data

### Department
- **Name**: `Emergency Medicine`
- **Code**: `EMERG`
- **Department ID**: `6c05dfac-9fd1-496f-a9dc-8ec636e02f6b`

### Department Admin Account
- **Email**: `admin.emergency@hospital.org`
- **Password**: `EmergencyAdmin2026!`
- **Role**: `department_admin`

> [!NOTE]
> A `SUPER_ADMIN` account cannot generate or view departmental schedules because `department_id` is `null`. Schedule endpoints require `department_admin` or `department_member` scope.

---

## 3. Sample Dataset (`sample_hospital_schedule_10_doctors.xlsx`)

Located at: `backend/data/sample_hospital_schedule_10_doctors.xlsx`

### Sheet 1: `Doctors` (10 Doctors)
| Name | Email | Position | Team | Unavailability | Pre-Assignments |
| :--- | :--- | :--- | :--- | :--- | :--- |
| Dr. Gregory House | `house@hospital.org` | Emergency | Team Alpha | `4, 18` | |
| Dr. Allison Cameron | `cameron@hospital.org` | Emergency | Team Beta | `11, 25` | |
| Dr. Eric Foreman | `foreman@hospital.org` | Emergency | Team Gamma | | |
| Dr. Robert Chase | `chase@hospital.org` | Emergency | Team Delta | | `10:Night, 24:Night` |
| Dr. Chris Taub | `taub@hospital.org` | Emergency | Team Epsilon | | |
| Dr. James Wilson | `wilson@hospital.org` | ICU | Team Alpha | `7, 8` | |
| Dr. Lisa Cuddy | `cuddy@hospital.org` | ICU | Team Beta | `14, 15` | |
| Dr. Remy Hadley | `hadley@hospital.org` | ICU | Team Gamma | | `12:On-Call, 26:On-Call` |
| Dr. Lawrence Kutner | `kutner@hospital.org` | ICU | Team Delta | | |
| Dr. Martha Masters | `masters@hospital.org` | ICU | Team Epsilon | | |

### Sheet 2: `Shifts`
| Position | Shift Name | Doctors Per Shift | Duty Days | Grants Day Off |
| :--- | :--- | :--- | :--- | :--- |
| Emergency | `Night` | 1 | `Mon-Sun` | `Yes` |
| ICU | `On-Call` | 1 | `Mon-Sun` | `Yes` |

---

## 4. Step-by-Step Execution

### Step 1: Start Backend API Server
```powershell
cd C:\Users\ptria\source\repos\Schedule-Maker\backend
uv run uvicorn src.main:app --reload --port 8000
```
Verify the server starts without errors at `http://127.0.0.1:8000`.

### Step 2: Start Frontend Application
In a separate terminal:
```powershell
cd C:\Users\ptria\source\repos\Schedule-Maker\frontend
flutter run -d chrome
# or: flutter run -d windows
```

### Step 3: Log In as Department Admin
1. On the MedShift login screen, enter:
   - **Email**: `admin.emergency@hospital.org`
   - **Password**: `EmergencyAdmin2026!`
2. Submit the form to transition to the Admin Dashboard.

### Step 4: Generate Schedule via Excel
1. Click the **"Generate Schedule"** button on the `AdminHeroCard`.
2. Select **"Upload Excel Roster"**.
3. Choose the sample file `backend/data/sample_hospital_schedule_10_doctors.xlsx`.
4. Ensure the target month is set to **`2026-11`**.
5. Click **Upload & Generate**.

### Step 5: Verify Generated Schedule
1. **Summary Banner**: Solver status displays `OPTIMAL` with 60 total duties.
2. **Table View**:
   - Verify sorted rows by date, position, and shift.
   - Confirm Dr. Robert Chase is assigned to `10:Night` and `24:Night`.
   - Confirm Dr. Remy Hadley is assigned to `12:On-Call` and `26:On-Call`.
3. **Calendar View**:
   - Switch toggle to Calendar view.
   - Inspect green shift dots on calendar days.
   - Click days (e.g. Day 10, Day 12, Day 24, Day 26) to view the detail inspection panel.
4. **Export Excel**:
   - Click **"Export to Excel"** in the top bar.
   - Verify `schedule_2026-11.xlsx` downloads and contains the formatted schedule.

---

## 5. Troubleshooting & Known Behaviors

1. **Super Admin 401/403 Errors**:
   - `POST /api/v1/schedules/generate-from-excel` requires role `department_admin`.
   - If a Super Admin attempts schedule generation, the backend rejects it with 403 Forbidden because a Super Admin has no departmental roster scope. Always log in as a Department Admin for scheduling workflows.
2. **Missing UI Error Feedback**:
   - If schedule generation fails (e.g. invalid file format or network issue), the UI should display an error SnackBar and not remain silently on the empty placeholder. Ensure `AdminScreen` listens to `scheduleGenerationControllerProvider` errors.
