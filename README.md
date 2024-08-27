# Theatre School Schedule Planner

Well, most of the code was AI generated, please modify whatever you want, I only had an hour to do it.
The Readme is also generated, if it is stupid just skip it.


This project is a Python-based scheduling tool designed for a theatre school. It generates an annual course schedule considering various constraints like holidays, vacations, and start dates for classes. The tool outputs the schedule as an Excel file, with classes organized by location and color-coded for easy reference.


## Features

- **Automated Schedule Generation**: Automatically schedules courses based on input configurations.
- **Configurable**: Supports configuration files for classes, holidays, and vacations.
- **Holiday & Vacation Handling**: Courses are not scheduled on holidays or during vacation periods.
- **Custom Start Dates**: Allows setting a start date for each course to ensure classes do not begin before a specified date.
- **Color-Coded Output**: Classes are color-coded in the output Excel file based on configurations.
- **Merged Cells for Time Periods**: Merges cells for months, weeks, and years to enhance readability in the Excel output.

## Installation

1. **Clone the repository**:

   ```bash
   git clone git@github.com:LNGGlago/planning_generator.git
   cd planning_generator
   ```

2. **Create a virtual environment** (optional but recommended):

   ```bash
   python3 -m venv venv
   source venv/bin/activate  # On Windows use `venv\Scripts\activate`
   ```

3. **Install the required dependencies**:

   ```bash
   pip install -r requirements.txt
   ```

## Usage

1. **Prepare Configuration Files**:

   - `config/classes.yaml`: Define your classes, including the name, location, time, day of the week, number of sessions, color, and optional start date.
   - `config/holidays.yaml`: List the holidays when classes should not be scheduled.
   - `config/vacations.yaml`: Specify vacation periods to be avoided in the schedule.

2. **Run the Script**:

   ```bash
   python generate_planning.py
   ```

3. **Output**:
   - The generated schedule will be saved as an Excel file in the `output` directory, with the name `planning_formatted_with_years.xlsx`.

## Configuration File Examples

### `classes.yaml`

```yaml
classes:
  - name: "Class A"
    location: "Theatre 1"
    time: "10:00 AM - 12:00 PM"
    day_of_week: "Monday"
    num_classes: 31
    color: "FF5733"  # Hex color code
    start_date: "2024-09-12"
  - name: "Class B"
    location: "Theatre 2"
    time: "2:00 PM - 4:00 PM"
    day_of_week: "Wednesday"
    num_classes: 31
    color: "33C1FF"
```

### `holidays.yaml`

```yaml
holidays:
  - "2024-12-25"
  - "2025-01-01"
```

### `vacations.yaml`

```yaml
vacations:
  - start: "2024-12-21"
    end: "2025-01-05"
  - start: "2025-02-15"
    end: "2025-03-03"
```

