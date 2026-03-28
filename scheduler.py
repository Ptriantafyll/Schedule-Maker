from models import Department, Doctor
import datetime
import calendar
from ortools.sat.python import cp_model


class ShiftScheduler:
    def __init__(self, department: Department):
        self.department = department

    def _calculate_days_for_schedule(self, month: int, year: int) -> tuple[list[datetime.date]]:
        """Returns the list of dates and a weekend lookup dict for the given month."""

        first_day = datetime.date(year, month, 1)
        days_in_month = calendar.monthrange(
            year, month)[1]
        dates = [first_day + datetime.timedelta(days=i)
                 for i in range(days_in_month)]

        # Identify weekends
        self.is_weekend = {d: (d.weekday() >= 5) for d in dates}
        return dates

    def _get_assignment_vars_for(self, day_index=None, position=None, shift=None, doctor=None):
        """Returns all shift assignment variables matching the given filters."""
        return [
            var for (d, p, s, doc), var in self.shift_assignments.items()
            if (day_index is None or d == day_index)
            and (position is None or p == position)
            and (shift is None or s == shift)
            and (doctor is None or doc == doctor)
        ]

    def _get_weekends(self, dates: list[datetime.date]) -> list[list[int]]:
        """Returns list of (fri, sat, sun) day index tuples for complete weekends."""
        weekends = []

        for day_index, date in enumerate(dates):
            if date.weekday() == 4 and day_index + 2 < len(dates):
                weekends.append([day_index, day_index+1, day_index+2])

        return weekends

    def _build_model(self, dates: list[datetime.date]):
        """Creates the CP-SAT model and shift assignment variables."""
        self.model = cp_model.CpModel()

        self.shift_assignments = {}
        self.penalties = []
        self.rewards = []

        for position in self.department.positions:
            position_doctors = position.eligible_doctors if position.eligible_doctors else self.department.doctors

            for doctor in position_doctors:
                pre_assignments = set(doctor.pre_assignments)

                for shift in position.shifts:
                    for day_index, date in enumerate(dates):
                        if date in doctor.unavailability:
                            continue

                        if date.weekday() not in position.duty_days:
                            continue

                        self.shift_assignments[(day_index, position, shift, doctor)] = self.model.new_bool_var(
                            f"shift_assignment_{day_index}_{position}_{shift}_{doctor}")

                        if (date, shift) in pre_assignments:
                            self.model.add(self.shift_assignments[(
                                day_index, position, shift, doctor)] == 1)
                        else:
                            if pre_assignments:
                                self.model.add(self.shift_assignments[(
                                    day_index, position, shift, doctor)] == 0)

    def _combine_objectives(self):
        self.model.minimize(sum(self.penalties) - sum(self.rewards))

    def _add_hard_constraint_doctors_per_shift(self, dates: list[datetime.date]):
        """Adds a hard constraint that a shift must have the exact number of doctor as specified"""

        for day_index, date in enumerate(dates):
            for position in self.department.positions:
                if date.weekday() not in position.duty_days:
                    continue

                for shift in position.shifts:

                    self.model.add(
                        sum(self._get_assignment_vars_for(
                            day_index, position, shift))
                        == shift.doctors_per_shift
                    )

    def _add_hard_constraint_no_consecutive_shifts(self, dates: list[datetime.date]):
        """Adds a hard constraint that a doctor cannot be on duty for 2 consecutive days"""

        for doctor in self.department.doctors:
            for day_index in range(len(dates) - 1):
                today_shifts = self._get_assignment_vars_for(
                    day_index=day_index, doctor=doctor)

                tomorrow_shifts = self._get_assignment_vars_for(
                    day_index=day_index + 1, doctor=doctor)

                self.model.add(sum(today_shifts + tomorrow_shifts) <= 1)

    def _add_hard_constraint_max_duties_per_doc_per_month(self, dates: list[datetime.date]):
        """Adds a hard constraint that sets the max duties per month a doctor can do"""

        for doctor in self.department.doctors:
            if doctor.pre_assignments:
                continue

            shifts = self._get_assignment_vars_for(doctor=doctor)
            self.model.add(
                sum(shifts) <= self.department.config.max_duties_per_month
            )

    def _add_hard_constraint_one_full_weekend_off_per_doctor(self, dates: list[datetime.date]):
        """Ensures every doctor has at least one full weekend (Fri+Sat+Sun) off."""
        weekends = self._get_weekends(dates=dates)

        for doctor in self.department.doctors:
            weekend_off_vars = []

            for weekend_index, (fri, sat, sun) in enumerate(weekends):
                weekend_off = self.model.new_bool_var(
                    f"full_weekend_off_{weekend_index}_{doctor}"
                )

                psk_shifts = (
                    self._get_assignment_vars_for(
                        day_index=fri, doctor=doctor
                    ) + self._get_assignment_vars_for(
                        day_index=sat, doctor=doctor
                    ) + self._get_assignment_vars_for(
                        day_index=sun, doctor=doctor
                    )
                )

                weekend_off_vars.append(weekend_off)
                self.model.add(sum(psk_shifts) + len(psk_shifts)
                               * weekend_off <= len(psk_shifts))

            self.model.add(sum(weekend_off_vars) >= 1)

    def _add_hard_constraint_balanced_total_duties_across_doctors(self, dates: list[datetime.date]):
        """Ensures all doctors have the same number of duties (+-1)"""
        pre_assigned_duties = sum(len(doctor.pre_assignments)
                                  for doctor in self.department.doctors if doctor.pre_assignments)

        total_duties = sum(
            shift.doctors_per_shift *
            sum(1 for d in dates if d.weekday() in position.duty_days)
            for position in self.department.positions
            for shift in position.shifts
        )

        duties_needed = total_duties - pre_assigned_duties
        doctors_available = sum(
            1 for doctor in self.department.doctors if not doctor.pre_assignments)

        min_duties = duties_needed // doctors_available
        max_duties = min_duties + 1

        for doctor in self.department.doctors:
            if doctor.pre_assignments:
                continue

            doctor_assignments = self._get_assignment_vars_for(doctor=doctor)
            self.model.add(sum(doctor_assignments) <= max_duties)
            self.model.add(sum(doctor_assignments) >= min_duties)

    def _add_hard_constraint_balanced_weekend_duties_across_doctors(self, dates: list[datetime.date]):
        """Ensures all doctors have the same number of weekend duties (+-1)"""
        pre_assigned_weekends = 0

        for doctor in self.department.doctors:
            if not doctor.pre_assignments:
                continue

            for (pre_assigned_date, _) in doctor.pre_assignments:
                if self.is_weekend[pre_assigned_date]:
                    pre_assigned_weekends += 1

        total_weekend_duties = sum(
            shift.doctors_per_shift *
            sum(1 for d in dates if
                self.is_weekend[d] and d.weekday() in position.duty_days)
            for position in self.department.positions
            for shift in position.shifts
        )

        weekend_duties_needed = total_weekend_duties - pre_assigned_weekends
        doctors_available = sum(
            1 for doctor in self.department.doctors if not doctor.pre_assignments)

        min_weekend_duties = weekend_duties_needed // doctors_available
        max_weekend_duties = min_weekend_duties + 1

        for doctor in self.department.doctors:
            if doctor.pre_assignments:
                continue
            assigned_weekends = []

            for day_index, date in enumerate(dates):
                if not self.is_weekend[date]:
                    continue

                assigned_weekends.extend(self._get_assignment_vars_for(
                    doctor=doctor, day_index=day_index))

            self.model.add(sum(assigned_weekends) <= max_weekend_duties)
            self.model.add(sum(assigned_weekends) >= min_weekend_duties)

    def _add_hard_constraint_max_one_team_day_off(self, dates: list[datetime.date]):
        """Ensures at most 1 doctor per team has a post-shift day off on the same day."""
        for day_index in range(1, len(dates)):
            for team in self.department.teams:
                yesterday_day_off_shifts = []

                for doctor in team.doctors:
                    for position in self.department.positions:
                        for shift in position.shifts:
                            if not shift.grants_day_off:
                                continue

                            key = (day_index - 1, position, shift, doctor)
                            if key in self.shift_assignments:
                                yesterday_day_off_shifts.append(
                                    self.shift_assignments[key])

            self.model.add(sum(yesterday_day_off_shifts) <= 1)

    def _add_soft_constraint_penalize_every_other_day_on_duty(self, dates: list[datetime.date]):
        """Penalizes on-off-on patterns where a doctor works day N and day N+2."""
        for doctor in self.department.doctors:
            if doctor.pre_assignments:
                continue

            for day_index in range(len(dates) - 2):
                day_a_var = self._get_assignment_vars_for(
                    day_index=day_index, doctor=doctor)
                day_c_var = self._get_assignment_vars_for(
                    day_index=day_index+2, doctor=doctor)

                if (not day_a_var) or (not day_c_var):
                    continue

                is_every_other = self.model.new_bool_var(
                    f"every_other_penalty_{day_index}_{doctor}_"
                )

                self.model.add(sum(day_a_var) + sum(day_c_var) ==
                               2).only_enforce_if(is_every_other)
                self.model.add(sum(day_a_var) + sum(day_c_var) !=
                               2).only_enforce_if(is_every_other.Not())

                self.penalties.append(
                    self.department.config.w_every_other_penalty * is_every_other)

    def create_schedule(self, month: int, year: int):
        pass


if __name__ == "__main__":
    app = ShiftScheduler(department=Department(name="test"))
    dates = app._calculate_days_for_schedule(month=4, year=2026)
    print(dates[0], "→", dates[-1])
    print("Days in month:", len(dates))
    print("Weekend days:", sum(
        1 for d in dates if ShiftScheduler.is_weekend[d]))
