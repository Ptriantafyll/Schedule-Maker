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

- Has a name and a number of duty slots per night (e.g. ER = 3, Cardiology = 1)
- Has its own pool of doctors
- Can have a pre-defined backup department that provides extra doctors when the primary pool is insufficient
- Has its own scheduling preferences expressed as constraint weights (e.g. a department may prefer clustered duties with longer rest periods vs. duties spread evenly across the month). These override the global defaults in `ScheduleConfig`.

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
| `Department`     | Name, slots per night, doctor list, backup department reference                                   |
| `ScheduleApp`    | Orchestrates the full pipeline: load data, build model, solve, export                             |

The `x` assignment variable becomes keyed on `(day_index, department, slot_index, doctor)`.

---

## Steps

1. Refactor current script into OOP (`ScheduleConfig`, `Doctor`, `Department`, `ScheduleApp`)
2. Extend model to support multiple departments and slots per night
3. Add cross-department doctor borrowing logic
4. Build backend API (Go or FastAPI) with auth and database
5. Build Flutter frontend (department login, unavailability form, schedule view)
