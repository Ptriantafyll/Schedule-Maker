# Schedule Maker — Project Plan

## Overview

A hospital on-duty scheduling tool that generates optimized monthly schedules per department using constraint programming. Accessible via a Windows and mobile app with per-department authentication.

---

## Architecture

### Backend

- Either **Go** or **Python (FastAPI)**
- Scheduling logic is always in **Python** using Google OR-Tools CP-SAT
- If Go is used as the main backend, Python runs as a separate microservice responsible only for schedule generation
- Handles authentication, department data, doctor management, and schedule storage

### Frontend

- **Flutter** — single codebase targeting Windows and mobile (Android/iOS)
- Each department head logs in and has access only to their department's data and schedules

### Database

- **PostgreSQL** or **SQLite**

---

## Domain Model

### Doctor

- Belongs to one department
- Has a name, email, and unavailability dates per month
- Max 7 on-duty days per month (across all departments they cover)

### Department

- Has a name and a list of positions, each with a required number of doctors per night (e.g. Internal Medicine: "ER" needs 2 doctors/night + "Clinic" needs 2 doctors/night; Hematology: "Clinic" needs 1 doctor/night)
- Has its own pool of doctors
- Can have a pre-defined backup department that provides extra doctors when the primary pool is insufficient
- Has its own scheduling preferences expressed as constraint weights (e.g. a department may prefer clustered duties with longer rest periods vs. duties spread evenly across the month). These override the global defaults in `ScheduleConfig`.

### Position

- Belongs to a department
- Has a name (e.g. "ER", "Clinic")
- Has one or more named shifts per night (e.g. ER has "1st shift" and "2nd shift"; Clinic may have only one)
- A doctor can only cover one shift per night

### Shift

- Belongs to a position
- Has a name (e.g. "1st shift", "2nd shift")
- Requires exactly 1 doctor per night
- The assignment is explicit — the schedule specifies which doctor covers which shift, not just which position

### Schedule

- Generated per department per month
- If a department lacks enough doctors to cover the month, doctors from the backup department are automatically assigned to fill the gap
- Cross-department assignments respect the 7-duty cap and no-consecutive-days rule across both schedules

---

## Scheduling Logic (Python / OR-Tools)

- One CP-SAT model per run, covering all departments simultaneously to enforce cross-department constraints
- Hard constraints:
  - Exactly N doctors on duty per night per department (N = slots_per_night)
  - No consecutive duty nights per doctor
  - Max 7 duties per doctor per month
  - At least one full weekend off (Fri + Sat + Sun) per doctor
  - Balanced total duties across doctors (difference of at most 1)
  - Balanced weekend duties across doctors
- Soft constraints (weighted objective):
  - Penalize every-other-day patterns
  - Penalize short gaps between duties
  - Spread duties evenly across the month (block balancing)
  - Reward full weekends off
  - Balance full weekends off across doctors
  - Balance Saturday vs Sunday duty distribution

---

## Authentication

- Email-based login
- Each user is linked to a department
- A user can only view and manage their own department's schedule and doctor list

---

## OOP Refactor Plan (current script → backend service)

Classes to introduce:

| Class            | Responsibility                                                                                    |
| ---------------- | ------------------------------------------------------------------------------------------------- |
| `ScheduleConfig` | Weights, solver time limit, max duties — defined per department, with global defaults as fallback |
| `Doctor`         | Name, department, unavailability                                                                  |
| `Position`       | Name, list of shifts, parent department                                                           |
| `Shift`          | Name (e.g. "1st shift"), parent position — always requires exactly 1 doctor per night             |
| `Department`     | Name, list of positions, doctor list, backup department reference                                 |
| `ScheduleApp`    | Orchestrates the full pipeline: load data, build model, solve, export                             |

The `x` assignment variable becomes keyed on `(day_index, shift, doctor)`.

---

## Steps

1. Refactor current script into OOP (`ScheduleConfig`, `Doctor`, `Department`, `ScheduleApp`)
2. Extend model to support multiple departments and slots per night
3. Add cross-department doctor borrowing logic
4. Build backend API (Go or FastAPI) with auth and database
5. Build Flutter frontend (department login, unavailability form, schedule view)
