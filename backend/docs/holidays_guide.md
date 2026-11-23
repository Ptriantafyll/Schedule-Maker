# Public Holidays Management Guide

## 1. Overview

In hospital scheduling, public holidays represent high-strain duty periods comparable to weekends. To prevent unfair workloads, the scheduler treats public holidays as weekend days for the **Weekend/Holiday Duty Balance** constraint, ensuring duties on statutory holidays and weekends are evenly distributed among doctors ($\pm 1$).

We use the [`python-holidays`](https://github.com/vacanza/python-holidays) library (`holidays`), an open-source package providing fast, offline statutory holiday calculations for over 100 countries.

---

## 2. Why `python-holidays`?

1. **Zero Runtime Latency**: Holidays are calculated algorithmically in memory (computing fixed dates, Easter-relative movable feasts, and bank holiday substitute rules) with **no external API calls or internet connection needed**.
2. **Comprehensive Regional Support**: Covers national and regional holiday rules for countries such as Greece (`'GR'`), Cyprus (`'CY'`), United Kingdom (`'GB'`), United States (`'US'`), and Germany (`'DE'`).
3. **Movable Feasts**: Automatically calculates Orthodox and Catholic movable holidays (e.g., Clean Monday, Orthodox Good Friday, Easter Monday, Pentecost, Holy Spirit Day) for any target year.

---

## 3. How It Works

### Basic Usage

```python
import datetime
import holidays

# Instantiate statutory holidays for a specific country and year
gr_holidays = holidays.country_holidays("GR", years=2026)

target_date = datetime.date(2026, 10, 28)

# Check if a date is a holiday:
if target_date in gr_holidays:
    # Get the official holiday name (e.g. "Ημέρα του Όχι")
    print(f"{target_date} is a holiday: {gr_holidays.get(target_date)}")
```

### Supported Country Codes (ISO 3166-1 alpha-2)

The library accepts standard two-letter ISO country codes:
- `'GR'` — Greece
- `'CY'` — Cyprus
- `'GB'` — United Kingdom
- `'US'` — United States
- `'DE'` — Germany
- `'FR'` — France
- `'IT'` — Italy

---

## 4. Scheduler Integration Architecture

In [`src/scheduler.py`](file:///C:/Users/ptria/source/repos/Schedule-Maker/backend/src/scheduler.py):
1. When generating a schedule for month $M$ and year $Y$, the scheduler queries the holiday list for the target country:
   ```python
   country_holidays = set(holidays.country_holidays(country_code, years=year).keys()) if country_code else set()
   ```
2. The `is_weekend` lookup mapping is populated as:
   ```python
   self.is_weekend = {
       d: (d.weekday() >= 5 or d in country_holidays)
       for d in self.dates
   }
   ```
3. **Duty Balancing**: Any shift worked on a day where `self.is_weekend[date] == True` increments the doctor's weekend/holiday duty total. The CP-SAT solver enforces that all doctors' totals differ by at most 1.
4. **Shift Operations**: Shifts continue to operate based strictly on their configured `duty_days`. If a shift operates Monday–Sunday or Monday–Friday, it operates on a holiday falling on that day.
5. **Full Weekends Off**: Full weekends off remain strictly tied to calendar Friday–Saturday–Sunday sequences; isolated midweek holidays do not alter full weekend definitions.

---

## 5. Adding or Upgrading with `uv`

To install or update the package:
```bash
# Add to backend dependencies:
uv add holidays

# Update to latest version:
uv lock --upgrade-package holidays
uv sync
```

---

## 6. Client Locale Detection & API Integration

### Automatic Country Detection
The Flutter client automatically resolves the user's country code without manual configuration or external IP geolocators:
```dart
String _resolveCountryCode(String? countryCode) {
  if (countryCode != null && countryCode.isNotEmpty) {
    return countryCode;
  }
  try {
    final detectedCode =
        WidgetsBinding.instance.platformDispatcher.locale.countryCode;
    if (detectedCode != null && detectedCode.isNotEmpty) {
      return detectedCode;
    }
  } catch (_) {
    // In isolated environments or before Flutter binding initializes, safely default
  }
  return 'GR';
}
```

### Transport
- **Multipart Form Upload** (`/api/v1/schedules/generate-from-excel`): sends `country_code: "GR"` in `FormData`.
- **Excel Export** (`/api/v1/schedules/export-excel`): passes `?country_code=GR` in URL query parameters so the generated spreadsheet styles holiday column headers with weekend highlight shading and populates the `"Total Holiday Duties"` summary column immediately to the right of `"Total Duties"`.
