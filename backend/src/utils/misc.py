"""
Module: misc.py. Contains functions that are used with miscellaneous things
"""

from time import perf_counter


def elapsed_ms(started: float) -> float:
    """Return elapsed miliseconds since a 'perf_counter()' reading."""
    return round((perf_counter()-started) * 1000, 2)
