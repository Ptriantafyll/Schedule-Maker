"""
Defines data models for the scheduling application. 
Doctors, shifts, positions, teams, and departments.
"""

from dataclasses import dataclass, field
import datetime
from typing import Optional


@dataclass(eq=False)
class ScheduleConfig:  # pylint: disable=too-many-instance-attributes
    """ Configuration for scheduling constraints and solver settings."""

    # Soft constraint weights
    w_every_other_penalty: int = 4
    w_gap_penalty: int = 2
    w_block_dev_penalty: int = 2
    w_full_wkend_off_bonus: int = 5
    w_balance_full_wkends_off: int = 20
    w_diff_wkend_duty_day: int = 2

    # Solver settings
    solver_time_limit: int = 120
    max_duties_per_month: int = 8

    # Schedule variables
    month_blocks: int = 3


@dataclass(eq=False)
class Shift:
    """ Represents a work shift, such as "Morning", "Afternoon", or "Night". """
    name: str
    doctors_per_shift: int = 1
    grants_day_off: bool = False


@dataclass(eq=False)
class Doctor:
    """ Represents a doctor working in the department. """
    name: str
    email: str
    unavailability: set[datetime.date] = field(default_factory=set)
    pre_assignments: list[tuple[datetime.date, Shift]
                          ] = field(default_factory=list)


@dataclass(eq=False)
class Position:
    """ Represents a position that needs to be staffed, such as "ER" or "ICU". """
    name: str
    shifts: list[Shift] = field(default_factory=list)
    duty_days: set[int] = field(default_factory=lambda: {0, 1, 2, 3, 4, 5, 6})
    eligible_doctors: list[Doctor] = field(default_factory=list)


@dataclass(eq=False)
class Team:
    """ Represents a team of doctors working together. """
    name: str
    doctors: list[Doctor] = field(default_factory=list)


@dataclass(eq=False)
class Department:
    """ Represents a hospital department with its scheduling needs. """
    name: str
    positions: list[Position] = field(default_factory=list)
    teams: list[Team] = field(default_factory=list)
    config: ScheduleConfig = field(default_factory=ScheduleConfig)
    backup_department: Optional['Department'] = None
    teamless_doctors: list[Doctor] = field(default_factory=list)
    # to remove later
    doctor_order: list[Doctor] = field(default_factory=list)

    @property
    def doctors(self) -> list[Doctor]:
        """ 
        Returns a list of all doctors in the department
        Including those in teams and teamless doctors.
        """
        team_doctors = [
            doctor for team in self.teams for doctor in team.doctors]
        return team_doctors + self.teamless_doctors
