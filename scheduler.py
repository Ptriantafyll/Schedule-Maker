from models import Department
import datetime
import calendar
from ortools.sat.python import cp_model


class ShiftScheduler:
    def __init__(self, departments: list[Department]):
        self.departments = departments

    def _calculate_days_for_schedule(self, month: int, year: int) -> tuple[list[datetime.date], dict[datetime.date, bool]]:
        """Returns the list of dates and a weekend lookup dict for the given month."""

        first_day = datetime.date(year, month, 1)
        days_in_month = calendar.monthrange(
            year, month)[1]
        dates = [first_day + datetime.timedelta(days=i)
                 for i in range(days_in_month)]

        # Identify weekends
        is_weekend = {d: (d.weekday() >= 5) for d in dates}
        return dates, is_weekend

    def _build_model(self, dates):
        """Creates the CP-SAT model and shift assignment variables."""
        self.model = cp_model.CpModel()

        self.shift_assignments = {}

        for department in self.departments:
            for position in department.positions:
                for shift in position.shifts:
                    for team in department.teams:
                        for doctor in team.doctors:
                            for day_index, date in enumerate(dates):
                                if date in doctor.unavailability:
                                    continue

                                self.shift_assignments[(day_index, shift, doctor)] = self.model.NewBoolVar(
                                    f"shift_assignment_{day_index}_{position}_{shift.name}_{doctor}")

    def create_schedule(self, month: int, year: int):
        pass


if __name__ == "__main__":
    app = ShiftScheduler(departments=[])
    dates, is_weekend = app._calculate_days_for_schedule(month=4, year=2026)
    print(dates[0], "→", dates[-1])
    print("Days in month:", len(dates))
    print("Weekend days:", sum(1 for d in dates if is_weekend[d]))
