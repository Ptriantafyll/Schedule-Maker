import datetime
from scheduler import ScheduleApp


def test_april_has_30_days():
    app = ScheduleApp(departments=[])
    dates, _ = app._calculate_days_for_schedule(month=4, year=2026)
    assert len(dates) == 30


def test_first_and_last_date():
    app = ScheduleApp(departments=[])
    dates, _ = app._calculate_days_for_schedule(month=4, year=2026)
    assert dates[0] == datetime.date(2026, 4, 1)
    assert dates[-1] == datetime.date(2026, 4, 30)


def test_weekends_identified_correctly():
    app = ScheduleApp(departments=[])
    _, is_weekend = app._calculate_days_for_schedule(month=4, year=2026)
    assert is_weekend[datetime.date(2026, 4, 4)] is True
    assert is_weekend[datetime.date(2026, 4, 6)] is False


def test_leap_year_identified_correctly():
    app = ScheduleApp(departments=[])
    dates, _ = app._calculate_days_for_schedule(month=2, year=2024)
    assert len(dates) == 29
