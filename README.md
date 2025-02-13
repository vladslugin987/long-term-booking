# Long Term Booking Generator

A Python application for generating long-term desk bookings in an office environment. The application supports both English and German interfaces.

## Features

- Generate Excel files for long-term desk bookings
- Select date ranges via calendar widget
- Choose specific weekdays for bookings
- First/Second half day booking options
- Private booking option
- Email validation (@devoteam.com domain)
- Bilingual interface (English/German)

## Installation

1. Clone the repository:
```bash
git clone https://github.com/vladslugin987/long-term-booking.git
cd long-term-booking
```

2. Create a virtual environment and activate it:
```bash
python -m venv venv
# Windows:
venv\Scripts\activate
# Linux/Mac:
source venv/bin/activate
```

3. Install required packages:
```bash
pip install -r requirements.txt
```

## Usage

Run the application:
```bash
python src/long_term_booking.py
```

## Building Executable

The repository includes GitHub Actions workflow that automatically builds the executable when pushing to the main branch.

To build manually:
```bash
pip install pyinstaller
pyinstaller --onefile src/long_term_booking.py
```

The executable will be created in the `dist` directory.

## Author

Created by Vladislav Slugin / vsdev.top
