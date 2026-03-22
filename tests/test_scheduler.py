from models import Department, Team, Doctor, Position, Shift
import datetime
from scheduler import ShiftScheduler


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
    dates, _ = scheduler._calculate_days_for_schedule(month=4, year=2026)
    assert len(dates) == 30


def test_first_and_last_date():
    scheduler = ShiftScheduler(department=Department(name="test"))
    dates, _ = scheduler._calculate_days_for_schedule(month=4, year=2026)
    assert dates[0] == datetime.date(2026, 4, 1)
    assert dates[-1] == datetime.date(2026, 4, 30)


def test_weekends_identified_correctly():
    scheduler = ShiftScheduler(department=Department(name="test"))
    _, is_weekend = scheduler._calculate_days_for_schedule(month=4, year=2026)
    assert is_weekend[datetime.date(2026, 4, 4)] is True
    assert is_weekend[datetime.date(2026, 4, 6)] is False


def test_leap_year_identified_correctly():
    scheduler = ShiftScheduler(department=Department(name="test"))
    dates, _ = scheduler._calculate_days_for_schedule(month=2, year=2024)
    assert len(dates) == 29


def test_hard_constraint_no_consecutive_duties():
    department = _make_test_department()
    scheduler = ShiftScheduler(department=department)

    dates, _ = scheduler._calculate_days_for_schedule(month=4, year=2026)
    scheduler._build_model(dates)
    scheduler._add_hard_constraint_no_consecutive_shifts(dates)

    # Force exactly 1 doctor per night so the solver actually assigns duties
    # Without this, the solver could assign nobody and trivially satisfy the constraint
    for day_index in range(len(dates)):
        daily_vars = [
            scheduler.shift_assignments[(day_index, position, shift, doctor)]
            for position in department.positions
            for shift in position.shifts
            for doctor in department.doctors
            if (day_index, position, shift, doctor) in scheduler.shift_assignments
        ]
        scheduler.model.Add(sum(daily_vars) == 1)

    # Step 5: Solve and verify no doctor has consecutive days
    from ortools.sat.python import cp_model
    solver = cp_model.CpSolver()
    status = solver.Solve(scheduler.model)

    assert status in (cp_model.OPTIMAL, cp_model.FEASIBLE), \
        "Solver failed to find a solution"

    for doctor in department.doctors:
        for day_index in range(len(dates) - 1):
            today = any(
                solver.Value(scheduler.shift_assignments[(day_index, pos, sh, doctor)])
                for pos in department.positions
                for sh in pos.shifts
                if (day_index, pos, sh, doctor) in scheduler.shift_assignments
            )
            tomorrow = any(
                solver.Value(scheduler.shift_assignments[(day_index + 1, pos, sh, doctor)])
                for pos in department.positions
                for sh in pos.shifts
                if (day_index + 1, pos, sh, doctor) in scheduler.shift_assignments
            )
            assert not (today and tomorrow), \
                f"{doctor.name} has consecutive duties on days {day_index} and {day_index + 1}"
