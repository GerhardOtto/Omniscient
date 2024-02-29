import pandas as pd
from icalendar import Calendar, Event


def read_exam_schedule(file_path):
    df = pd.read_excel(file_path, header=2)
    module_code_col = 'Module'
    opportunity_col = 'Paper'
    date_col = 'Date'
    start_time_col = 'Start'
    finish_time_col = 'End'
    return df[[module_code_col, opportunity_col, date_col, start_time_col, finish_time_col]]


def create_calendar_event(module_code, date, start_time, finish_time):
    event = Event()
    event.add('summary', f'Exam: {module_code}')
    event.add('dtstart', pd.to_datetime(f'{date} {start_time}'))
    event.add('dtend', pd.to_datetime(f'{date} {finish_time}'))
    return event


def main():
    file_path = input("Enter the path to the Excel file: ")
    modules_to_track = input("Enter module codes (comma-separated): ").split(',')

    exam_schedule = read_exam_schedule(file_path)
    processed_modules = set()

    cal = Calendar()

    for module_code in modules_to_track:
        if module_code not in exam_schedule['Module'].unique() and module_code not in processed_modules:
            print(f"Warning: Module '{module_code}' not found in the exam schedule.")
            processed_modules.add(module_code)
            continue

        module_data = exam_schedule[exam_schedule['Module'] == module_code]

        for index, row in module_data.iterrows():
            date = row['Date']
            start_time = row['Start']
            finish_time = row['End']

            # Calculate duration by subtracting start time from end time
            try:
                start_datetime = pd.to_datetime(f'{date} {start_time}')
                end_datetime = pd.to_datetime(f'{date} {finish_time}')
                duration = end_datetime - start_datetime
            except ValueError:
                print(f"Warning: Unable to calculate duration for module '{module_code}'. Skipping.")
                print("Row with problematic duration calculation:")
                print(row)
                continue

            # Print details about the event before adding it
            print(f"Adding event for module '{module_code}':")
            print(f"Date: {date}, Start Time: {start_time}, Finish Time: {finish_time}, Duration: {duration}")

            event = create_calendar_event(module_code, date, start_time, finish_time)
            cal.add_component(event)

    # Write all calendar events to a single .ics file in the directory of the Python file
    output_file_path = 'exam_schedule.ics'
    with open(output_file_path, 'wb') as f:
        f.write(cal.to_ical())

    print(f'Calendar events for all modules added to {output_file_path}')


if __name__ == "__main__":
    main()
