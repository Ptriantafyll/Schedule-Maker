# pylint: disable=protected-access
# pylint: disable=line-too-long
"""
Unit tests for ShiftScheduler constraints and date calculations.
"""

import datetime
from ortools.sat.python import cp_model
from scheduler import ShiftScheduler
from models import Department, Team, Doctor, Position, Shift, ScheduleConfig


def _build_and_solve(department, month=4, year=2026, constraint_names=None):
    """Helper that builds model, adds constraints by method name, solves, and returns (scheduler, solver, dates)."""
    scheduler = ShiftScheduler(department=department)
    scheduler._calculate_days_for_schedule(month=month, year=year)
    scheduler._build_model()

    has_soft_constraints = any(name.startswith(
        "_add_soft_constraint") for name in (constraint_names or []))
    
    if has_soft_constraints:
        scheduler._combine_objectives()

    for name in (constraint_names or []):
        getattr(scheduler, name)()
    solver = cp_model.CpSolver()
    status = solver.Solve(scheduler.model)
    assert status in (cp_model.OPTIMAL, cp_model.FEASIBLE), \
        "Solver failed to find a solution"
    return scheduler, solver


def _make_test_department():
    """Helper that builds a small department for testing."""
    doctors = [
        Doctor(name="Dr. A", email="a@test.com"),
        Doctor(name="Dr. B", email="b@test.com"),
        Doctor(name="Dr. C", email="C@test.com"),
        Doctor(name="Dr. D", email="D@test.com"),
        Doctor(name="Dr. E", email="E@test.com")
    ]
    team = Team(name="Team 1", doctors=doctors)
    shift = Shift(name="Night", doctors_per_shift=1)
    position = Position(name="ER", shifts=[shift])
    return Department(name="Test", teams=[team], positions=[position])


def test_april_has_30_days():
    """Verifies correct number of days are generated for April 2026."""
    scheduler = ShiftScheduler(department=Department(name="test"))
    scheduler._calculate_days_for_schedule(month=4, year=2026)
    assert len(scheduler.dates) == 30


def test_first_and_last_date():
    """Verifies first and last dates are correct for April 2026."""
    scheduler = ShiftScheduler(department=Department(name="test"))
    scheduler._calculate_days_for_schedule(month=4, year=2026)
    assert scheduler.dates[0] == datetime.date(2026, 4, 1)
    assert scheduler.dates[-1] == datetime.date(2026, 4, 30)


def test_weekends_identified_correctly():
    """Verifies weekends are correctly identified for April 2026."""
    scheduler = ShiftScheduler(department=Department(name="test"))
    scheduler._calculate_days_for_schedule(month=4, year=2026)
    assert scheduler.is_weekend[datetime.date(2026, 4, 4)] is True
    assert scheduler.is_weekend[datetime.date(2026, 4, 6)] is False


def test_leap_year_identified_correctly():
    """Verifies February 2024 has 29 days due to leap year."""
    scheduler = ShiftScheduler(department=Department(name="test"))
    scheduler._calculate_days_for_schedule(month=2, year=2024)
    assert len(scheduler.dates) == 29


def test_hard_constraint_no_consecutive_duties():
    """Verifies no doctor is assigned on two consecutive days."""
    department = _make_test_department()
    scheduler, solver = _build_and_solve(
        department,
        constraint_names=[
            "_add_hard_constraint_no_consecutive_shifts",
            "_add_hard_constraint_doctors_per_shift",
        ]
    )

    for doctor in department.doctors:
        for day_index in range(len(scheduler.dates) - 1):
            today = any(
                solver.Value(var)
                for var in scheduler._get_assignment_vars_for(day_index, doctor=doctor)
            )
            tomorrow = any(
                solver.Value(var)
                for var in scheduler._get_assignment_vars_for(day_index + 1, doctor=doctor)
            )
            assert not (today and tomorrow), \
                f"{doctor.name} has consecutive duties on days {day_index} and {day_index + 1}"


def test_pre_assignments():
    """Verifies pre-assigned doctors work only their chosen days and no others."""
    shift = Shift(name="Night", doctors_per_shift=1)
    doctors = [
        Doctor(name="Dr. A", email="a@test.com"),
        Doctor(name="Dr. B", email="b@test.com"),
        Doctor(name="Dr. C", email="C@test.com"),
        Doctor(name="Dr. D", email="D@test.com"),
        Doctor(name="Dr. E", email="E@test.com"),
        Doctor(name="Dr. Pre-assigned", email="preassigned@test.com",
                    pre_assignments=[
                        (datetime.date(2026, 4, 6), shift),
                        (datetime.date(2026, 4, 8), shift),
                        (datetime.date(2026, 4, 15), shift),
                        (datetime.date(2026, 4, 21), shift)
                    ])
    ]
    team = Team(name="Team 1", doctors=doctors)
    position = Position(name="ER", shifts=[shift])
    department = Department(name="test", teams=[team], positions=[position])

    scheduler, solver = _build_and_solve(
        department,
        constraint_names=[
            "_add_hard_constraint_no_consecutive_shifts",
            "_add_hard_constraint_doctors_per_shift",
            "_add_hard_constraint_max_duties_per_doc_per_month"
        ]
    )

    for (day_index, _, sh, doc), var in scheduler.shift_assignments.items():
        if doc == doctors[5]:
            if (scheduler.dates[day_index], sh) in doctors[5].pre_assignments:
                assert solver.Value(
                    var) == 1, f"Should be assigned on day {day_index}"
            else:
                assert solver.Value(
                    var) == 0, f"Should not be assigned on day {day_index}"


def test_hard_constraint_doctors_per_shift():
    """Verifies each shift has exactly the required number of doctors per night."""
    department = _make_test_department()
    scheduler, solver = _build_and_solve(
        department,
        constraint_names=["_add_hard_constraint_doctors_per_shift"]
    )

    for day_index in range(len(scheduler.dates)):
        for position in department.positions:
            for shift in position.shifts:
                assigned_count = sum(
                    solver.Value(var)
                    for var in scheduler._get_assignment_vars_for(day_index, position, shift)
                )
                assert assigned_count == shift.doctors_per_shift, f"Day {day_index}, {position.name}/{shift.name}: expected {shift.doctors_per_shift} doctor(s), got {assigned_count}"


def test_hard_constraint_doctors_per_shift_multiple():
    """Tests with a shift that requires 2 doctors."""
    doctors = [
        Doctor(name=f"Dr. {c}", email=f"{c}@test.com")
        for c in "ABCDEFGH"
    ]
    team = Team(name="Team 1", doctors=doctors)
    shift = Shift(name="Night", doctors_per_shift=2)
    position = Position(name="ER", shifts=[shift])
    department = Department(name="Test", teams=[team], positions=[position])

    scheduler, solver = _build_and_solve(
        department,
        constraint_names=["_add_hard_constraint_doctors_per_shift"]
    )

    for day_index in range(len(scheduler.dates)):
        assigned_count = sum(
            solver.Value(var)
            for var in scheduler._get_assignment_vars_for(day_index, position, shift)
        )
        assert assigned_count == 2, f"Day {day_index}: expected 2 doctors, got {assigned_count}"


def test_hard_constraint_max_duties_per_doc_per_month():
    """Tests with max 5 duties per month - needs enough doctors to cover 30 days"""
    doctors = [
        Doctor(name=f"Dr. {c}", email=f"{c}@test.com")
        for c in "ABCDEFGH"
    ]
    team = Team(name="Team 1", doctors=doctors)
    shift = Shift(name="Night", doctors_per_shift=1)
    position = Position(name="ER", shifts=[shift])
    config = ScheduleConfig(max_duties_per_month=5)
    department = Department(name="Test", teams=[team], positions=[
        position], config=config)

    scheduler, solver = _build_and_solve(
        department=department,
        constraint_names=[
            "_add_hard_constraint_no_consecutive_shifts",
            "_add_hard_constraint_doctors_per_shift",
            "_add_hard_constraint_max_duties_per_doc_per_month"
        ]
    )

    for doctor in department.doctors:
        total_duties = sum(
            solver.Value(var)
            for var in scheduler._get_assignment_vars_for(doctor=doctor)
        )
        assert total_duties <= 5, f"{doctor.name} has {total_duties} duties, expected at most 5"


def test_get_weekends_april_2026():
    """April 2026: first Friday is Apr 3, last full weekend ends Apr 26."""
    scheduler = ShiftScheduler(department=Department(name="test"))
    scheduler._calculate_days_for_schedule(month=4, year=2026)
    weekends = scheduler._get_weekends()

    # April 2026 has 4 full Fri-Sat-Sun weekends
    assert len(weekends) == 4

    # First weekend: Apr 3 (Fri) = day_index 2, Apr 4, Apr 5
    assert weekends[0] == [2, 3, 4]

    # Verify all weekends start on a Friday
    for fri, sat, sun in weekends:
        assert scheduler.dates[fri].weekday(
        ) == 4, f"Day {fri} is not a Friday"
        assert scheduler.dates[sat].weekday(
        ) == 5, f"Day {sat} is not a Saturday"
        assert scheduler.dates[sun].weekday(
        ) == 6, f"Day {sun} is not a Sunday"


def test_get_weekends_month_starting_saturday():
    """August 2025 starts on a Friday — first weekend should be complete."""
    scheduler = ShiftScheduler(department=Department(name="test"))
    scheduler._calculate_days_for_schedule(month=8, year=2025)
    weekends = scheduler._get_weekends()

    # Aug 1 2025 is a Friday, so first weekend is day 0, 1, 2
    assert weekends[0] == [0, 1, 2]
    assert scheduler.dates[0].weekday() == 4


def test_hard_constraint_one_full_weekend_off_per_doctor():
    """Verifies every doctor has at least one full weekend (Fri+Sat+Sun) off."""
    test_department = _make_test_department()
    scheduler, solver = _build_and_solve(
        department=test_department,
        constraint_names=[
            "_add_hard_constraint_no_consecutive_shifts",
            "_add_hard_constraint_doctors_per_shift",
            "_add_hard_constraint_one_full_weekend_off_per_doctor"
        ]
    )

    weekends = scheduler._get_weekends()
    for doctor in test_department.doctors:
        has_full_weekend_off = False
        for (fri, sat, sun) in weekends:
            fri_off = not any(solver.Value(var) for var in scheduler._get_assignment_vars_for(
                day_index=fri, doctor=doctor))
            sat_off = not any(solver.Value(var) for var in scheduler._get_assignment_vars_for(
                day_index=sat, doctor=doctor))
            sun_off = not any(solver.Value(var) for var in scheduler._get_assignment_vars_for(
                day_index=sun, doctor=doctor))

            if fri_off and sat_off and sun_off:
                has_full_weekend_off = True
                break

        assert has_full_weekend_off, f"{doctor.name} has no full weekend off"


def test_balanced_total_duties_across_doctors():
    """Verifies total duties per doctor differ by at most 1 across all doctors."""
    test_department = _make_test_department()

    scheduler, solver = _build_and_solve(
        department=test_department,
        constraint_names=[
            "_add_hard_constraint_no_consecutive_shifts",
            "_add_hard_constraint_doctors_per_shift",
            "_add_hard_constraint_one_full_weekend_off_per_doctor",
            "_add_hard_constraint_balanced_total_duties_across_doctors"
        ]
    )

    number_of_doctors = len(test_department.doctors)
    total_duties = sum(
        shift.doctors_per_shift *
        sum(1 for d in scheduler.dates if d.weekday() in position.duty_days)
        for position in test_department.positions
        for shift in position.shifts
    )

    min_duties = total_duties // number_of_doctors
    max_duties = min_duties + 1

    for doctor in test_department.doctors:
        doctor_assignments = sum(
            solver.Value(var) for var in scheduler._get_assignment_vars_for(doctor=doctor)
        )
        assert doctor_assignments <= max_duties, f"{doctor.name} has more than the max allowed duties"
        assert doctor_assignments >= min_duties, f"{doctor.name} has more than the min allowed duties"


def test_balanced_weekend_duties_across_doctors():
    """Verifies each doctor's weekend duty count differs by at most 1 from any other."""
    test_department = _make_test_department()

    scheduler, solver = _build_and_solve(
        department=test_department,
        constraint_names=[
            "_add_hard_constraint_no_consecutive_shifts",
            "_add_hard_constraint_doctors_per_shift",
            "_add_hard_constraint_one_full_weekend_off_per_doctor",
            "_add_hard_constraint_balanced_total_duties_across_doctors",
            "_add_hard_constraint_balanced_weekend_duties_across_doctors"
        ]
    )

    total_weekend_duties = sum(
        shift.doctors_per_shift *
        sum(1 for d in scheduler.dates if
            scheduler.is_weekend[d] and d.weekday() in position.duty_days)
        for position in scheduler.department.positions
        for shift in position.shifts
    )

    doctors_available = sum(
        1 for doctor in scheduler.department.doctors if not doctor.pre_assignments)

    min_weekend_duties = total_weekend_duties // doctors_available
    max_weekend_duties = min_weekend_duties + 1

    for doctor in test_department.doctors:
        total_weekend_assignments = 0
        for day_index, date in enumerate(scheduler.dates):
            if not scheduler.is_weekend[date]:
                continue

            total_weekend_assignments += sum(
                solver.Value(var) for var in scheduler._get_assignment_vars_for(doctor=doctor, day_index=day_index)
            )
        assert total_weekend_assignments <= max_weekend_duties, f"{doctor.name} has more than the max allowed duties"
        assert total_weekend_assignments >= min_weekend_duties, f"{doctor.name} has more than the min allowed duties"


def calculate_yesterday_day_off_count(position: Position, day_index: int, doctor: Doctor, scheduler: ShiftScheduler, solver: cp_model.CpSolver) -> int:
    """Helper to calculate how many doctors in the team had a day off yesterday due to a shift that grants day off."""
    yesterday_day_off_count = 0
    for shift in position.shifts:
        if not shift.grants_day_off:
            continue

        key = (day_index - 1, position, shift, doctor)
        if key in scheduler.shift_assignments:
            if solver.Value(scheduler.shift_assignments[key]):
                yesterday_day_off_count += 1

    return yesterday_day_off_count


def test_max_one_day_off_team():
    """Verifies at most 1 doctor per team has a post-shift day off on the same day."""
    doctors = [
        Doctor(name="Dr. A", email="a@test.com"),
        Doctor(name="Dr. B", email="b@test.com"),
        Doctor(name="Dr. C", email="C@test.com"),
        Doctor(name="Dr. D", email="D@test.com"),
        Doctor(name="Dr. E", email="E@test.com"),
        Doctor(name="Dr. F", email="D@test.com"),
        Doctor(name="Dr. G", email="D@test.com"),
        Doctor(name="Dr. H", email="D@test.com"),
        Doctor(name="Dr. I", email="D@test.com")
    ]
    teams = [
        Team(name="Team 1", doctors=doctors[:3]),
        Team(name="Team 2", doctors=doctors[3:])
    ]
    shifts = [
        Shift(name="ER", doctors_per_shift=1, grants_day_off=True),
        Shift(name="Night", doctors_per_shift=1)
    ]
    position = Position(name="ER", shifts=shifts)
    test_department = Department(
        name="Test", teams=teams, positions=[position])

    scheduler, solver = _build_and_solve(
        department=test_department,
        constraint_names=[
            "_add_hard_constraint_no_consecutive_shifts",
            "_add_hard_constraint_doctors_per_shift",
            "_add_hard_constraint_one_full_weekend_off_per_doctor",
            "_add_hard_constraint_balanced_total_duties_across_doctors",
            "_add_hard_constraint_balanced_weekend_duties_across_doctors",
            "_add_hard_constraint_max_one_team_day_off"
        ]
    )

    for day_index in range(1, len(scheduler.dates)):
        for team in teams:
            yesterday_day_off_count = 0
            for doctor in team.doctors:
                yesterday_day_off_count += calculate_yesterday_day_off_count(
                    position, day_index, doctor, scheduler, solver)

            assert yesterday_day_off_count <= 1, f"{team.name} has more than 1 doctor with the day off"


def get_full_weekend_off_count_for_doctor(doctor: Doctor, weekends: list, scheduler: ShiftScheduler, solver: cp_model.CpSolver) -> int:
    """Helper to calculate how many full weekends off a doctor has."""
    full_weekend_off_count = 0
    for (fri, sat, sun) in weekends:
        fri_off = not any(solver.Value(var) for var in scheduler._get_assignment_vars_for(
            day_index=fri, doctor=doctor))
        sat_off = not any(solver.Value(var) for var in scheduler._get_assignment_vars_for(
            day_index=sat, doctor=doctor))
        sun_off = not any(solver.Value(var) for var in scheduler._get_assignment_vars_for(
            day_index=sun, doctor=doctor))

        if fri_off and sat_off and sun_off:
            full_weekend_off_count += 1

    return full_weekend_off_count


def test_reward_full_weekends_off():
    """Verifies that the soft constraint for balancing full weekends off is working"""
    test_department = _make_test_department()
    scheduler, solver = _build_and_solve(
        department=test_department,
        constraint_names=[
            "_add_hard_constraint_no_consecutive_shifts",
            "_add_hard_constraint_doctors_per_shift",
            "_add_hard_constraint_one_full_weekend_off_per_doctor",
            "_add_hard_constraint_balanced_total_duties_across_doctors",
            "_add_hard_constraint_balanced_weekend_duties_across_doctors",
            "_add_hard_constraint_max_one_team_day_off",
        ]
    )

    weekends = scheduler._get_weekends()

    full_weekend_off_counts = []
    for doctor in test_department.doctors:
        full_weekend_off_count = get_full_weekend_off_count_for_doctor(
            doctor, weekends, scheduler, solver)
        full_weekend_off_counts.append(full_weekend_off_count)

    min_full_weekends_off = min(full_weekend_off_counts)
    max_full_weekends_off = max(full_weekend_off_counts)

    scheduler, solver = _build_and_solve(
        department=test_department,
        constraint_names=[
            "_add_hard_constraint_no_consecutive_shifts",
            "_add_hard_constraint_doctors_per_shift",
            "_add_hard_constraint_one_full_weekend_off_per_doctor",
            "_add_hard_constraint_balanced_total_duties_across_doctors",
            "_add_hard_constraint_balanced_weekend_duties_across_doctors",
            "_add_hard_constraint_max_one_team_day_off",
            "_add_soft_constraint_reward_full_weekends_off"
        ]
    )

    full_weekend_off_counts_with_soft = []
    for doctor in test_department.doctors:
        full_weekend_off_count = get_full_weekend_off_count_for_doctor(
            doctor, weekends, scheduler, solver)
        full_weekend_off_counts_with_soft.append(full_weekend_off_count)

    min_full_weekends_off_with_soft = min(full_weekend_off_counts_with_soft)
    max_full_weekends_off_with_soft = max(full_weekend_off_counts_with_soft)

    assert max_full_weekends_off_with_soft - min_full_weekends_off_with_soft <= max_full_weekends_off - min_full_weekends_off, \
        "Soft constraint did not improve balance of full weekends off"
