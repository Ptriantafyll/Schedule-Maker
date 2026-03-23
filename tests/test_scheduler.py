from models import Department, Team, Doctor, Position, Shift, ScheduleConfig
import datetime
from ortools.sat.python import cp_model
from scheduler import ShiftScheduler


def _build_and_solve(department, month=4, year=2026, constraint_names=None):
    """Jelper that builds model, adds constraints by method name, solves, and returns (scheduler, solver, dates)."""
    scheduler = ShiftScheduler(department=department)
    dates = scheduler._calculate_days_for_schedule(month=month, year=year)
    scheduler._build_model(dates)
    for name in (constraint_names or []):
        getattr(scheduler, name)(dates)
    solver = cp_model.CpSolver()
    status = solver.Solve(scheduler.model)
    assert status in (cp_model.OPTIMAL, cp_model.FEASIBLE), \
        "Solver failed to find a solution"
    return scheduler, solver, dates


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
    scheduler = ShiftScheduler(department=Department(name="test"))
    dates = scheduler._calculate_days_for_schedule(month=4, year=2026)
    assert len(dates) == 30


def test_first_and_last_date():
    scheduler = ShiftScheduler(department=Department(name="test"))
    dates = scheduler._calculate_days_for_schedule(month=4, year=2026)
    assert dates[0] == datetime.date(2026, 4, 1)
    assert dates[-1] == datetime.date(2026, 4, 30)


def test_weekends_identified_correctly():
    scheduler = ShiftScheduler(department=Department(name="test"))
    scheduler._calculate_days_for_schedule(month=4, year=2026)
    assert scheduler.is_weekend[datetime.date(2026, 4, 4)] is True
    assert scheduler.is_weekend[datetime.date(2026, 4, 6)] is False


def test_leap_year_identified_correctly():
    scheduler = ShiftScheduler(department=Department(name="test"))
    dates = scheduler._calculate_days_for_schedule(month=2, year=2024)
    assert len(dates) == 29


def test_hard_constraint_no_consecutive_duties():
    department = _make_test_department()
    scheduler, solver, dates = _build_and_solve(
        department,
        constraint_names=[
            "_add_hard_constraint_no_consecutive_shifts",
            "_add_hard_constraint_doctors_per_shift",
        ]
    )

    for doctor in department.doctors:
        for day_index in range(len(dates) - 1):
            today = any(
                solver.Value(var)
                for var in scheduler._get_assignments_for(day_index, doctor=doctor)
            )
            tomorrow = any(
                solver.Value(var)
                for var in scheduler._get_assignments_for(day_index + 1, doctor=doctor)
            )
            assert not (today and tomorrow), \
                f"{doctor.name} has consecutive duties on days {day_index} and {day_index + 1}"


def test_hard_constraint_doctors_per_shift():
    department = _make_test_department()
    scheduler, solver, dates = _build_and_solve(
        department,
        constraint_names=["_add_hard_constraint_doctors_per_shift"]
    )

    for day_index in range(len(dates)):
        for position in department.positions:
            for shift in position.shifts:
                assigned_count = sum(
                    solver.Value(var)
                    for var in scheduler._get_assignments_for(day_index, position, shift)
                )
                assert assigned_count == shift.doctors_per_shift, \
                    f"Day {day_index}, {position.name}/{shift.name}: expected {shift.doctors_per_shift} doctor(s), got {assigned_count}"


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

    scheduler, solver, dates = _build_and_solve(
        department,
        constraint_names=["_add_hard_constraint_doctors_per_shift"]
    )

    for day_index in range(len(dates)):
        assigned_count = sum(
            solver.Value(var)
            for var in scheduler._get_assignments_for(day_index, position, shift)
        )
        assert assigned_count == 2, \
            f"Day {day_index}: expected 2 doctors, got {assigned_count}"


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

    scheduler, solver, dates = _build_and_solve(
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
            for var in scheduler._get_assignments_for(doctor=doctor)
        )
        assert total_duties <= 5, \
            f"{doctor.name} has {total_duties} duties, expected at most 5"


def test_get_weekends_april_2026():
    """April 2026: first Friday is Apr 3, last full weekend ends Apr 26."""
    scheduler = ShiftScheduler(department=Department(name="test"))
    dates = scheduler._calculate_days_for_schedule(month=4, year=2026)
    weekends = scheduler._get_weekends(dates)

    # April 2026 has 4 full Fri-Sat-Sun weekends
    assert len(weekends) == 4

    # First weekend: Apr 3 (Fri) = day_index 2, Apr 4, Apr 5
    assert weekends[0] == [2, 3, 4]

    # Verify all weekends start on a Friday
    for fri, sat, sun in weekends:
        assert dates[fri].weekday() == 4, f"Day {fri} is not a Friday"
        assert dates[sat].weekday() == 5, f"Day {sat} is not a Saturday"
        assert dates[sun].weekday() == 6, f"Day {sun} is not a Sunday"


def test_get_weekends_month_starting_saturday():
    """August 2025 starts on a Friday — first weekend should be complete."""
    scheduler = ShiftScheduler(department=Department(name="test"))
    dates = scheduler._calculate_days_for_schedule(month=8, year=2025)
    weekends = scheduler._get_weekends(dates)

    # Aug 1 2025 is a Friday, so first weekend is day 0, 1, 2
    assert weekends[0] == [0, 1, 2]
    assert dates[0].weekday() == 4
