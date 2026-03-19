from dataclasses import dataclass, field
import datetime
from typing import Optional


@dataclass
class ScheduleConfig:
    # Soft constraint weights
    w_every_other_penalty: int = 4
    w_gap_penalty: int = 2
    w_block_dev_penalty: int = 2
    w_full_wkend_off_bonus: int = 5
    w_balance_full_wkends_off: int = 20
    w_diff_wkend_duty_day: int = 2

    # Solver settings
    solver_time_limit: int = 120
    max_duties_per_month: int = 7


@dataclass
class Doctor:
    name: str
    email: str
    unavailability: set[datetime.date] = field(default_factory=set)


@dataclass
class Shift:
    name: str
    doctors_per_shift: int = 1


@dataclass
class Position:
    name: str
    shifts: list[Shift] = field(default_factory=list)


@dataclass
class Department:
    name: str
    positions: list[Position] = field(default_factory=list)
    doctors: list[Doctor] = field(default_factory=list)
    config: ScheduleConfig = field(default_factory=ScheduleConfig)
    backup_department: Optional['Department'] = None
