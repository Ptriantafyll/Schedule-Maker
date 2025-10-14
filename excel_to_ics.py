# calendar_event_creator.py
import argparse
import pandas as pd
from ics import Calendar, Event


def create_calendar_from_excel(excel_path, output_ics="duty_schedule.ics", arguments=None,):
    '''
    Creates an ICS calendar file from an Excel schedule.

    Args:    
        excel_path: Path to the Excel file.
        output_ics: Path to save the generated ICS file.
        timezone: Timezone for the events.
        doctor_name: If provided, only events for this doctor are included.

    Returns:
        None. Writes the ICS file to output_ics.
    '''

    # timezone="Europe/Athens"
    doctor = arguments.doctor_name
    df = pd.read_excel(excel_path)
    cal = Calendar()

    for _, row in df.iterrows():
        if (row['Assigned Doctor']) != doctor and doctor is not None:
            continue

        e = Event()
        e.name = "Εφημερία"
        e.begin = pd.to_datetime(row['Date']).date()
        e.make_all_day()
        e.description = ''
        e.location = "ΠΓΝΠ"

        cal.events.add(e)

    with open(output_ics, "w", encoding="utf-8") as f:
        f.writelines(cal)

    # Remove blank lines
    with open(output_ics, "r", encoding="utf-8") as f:
        ilnes = f.readlines()

    with open(output_ics, "w", encoding="utf-8") as f:
        for line in ilnes:
            if line.strip():
                f.write(line)

    print(f"✅ Created {output_ics} with all-day events.")


# Example usage
if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Create ICS calendar from Excel schedule.")
    parser.add_argument("-d", "--doctor-name", type=str,
                        help="Name of the doctor to filter events for.")
    args = parser.parse_args()

    create_calendar_from_excel(
        "monthly_schedule.xlsx", "duty_schedule.ics", args)
