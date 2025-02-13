"""
Long Term Booking
-----------------
This script generates an Excel file for long term booking reservations.
Each row in the Excel file includes:
    - id: a random UUID
    - seatID: a table seat with a prefix "P" (e.g. P24)
    - date: in dd.mm.yyyy format
    - timestamp: corresponding to the start of the day in milliseconds
    - firstHalf: booking for first half of the day (TRUE/FALSE)
    - secondHalf: booking for second half of the day (TRUE/FALSE)
    - email: must end with @devoteam.com
    - private: booking marked as private (TRUE/FALSE)

Additional features:
    - Choose a date range via a calendar widget (tkcalendar)
    - Select weekdays (Monday to Friday) to include in the booking
    - Select first/second half-day and mark booking as private
    - Switch the UI language between English and German
    - Developer info is shown (Vladislav Slugin / vsdev.top)

To build an executable from this script, you can create a GitHub repository (e.g. long-term-booking)
and then use PyInstaller:
    pyinstaller --onefile long_term_booking.py
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from tkcalendar import DateEntry
import uuid
from datetime import datetime, timedelta
import openpyxl

# Translation dictionary for English and German
translations = {
    "en": {
         "title": "Long Term Booking",
         "start_date": "Start Date",
         "end_date": "End Date",
         "weekdays": "Select Weekdays",
         "monday": "Monday",
         "tuesday": "Tuesday",
         "wednesday": "Wednesday",
         "thursday": "Thursday",
         "friday": "Friday",
         "first_half": "First Half",
         "second_half": "Second Half",
         "private": "Private",
         "email": "Email",
         "seat": "Seat Number (e.g. 24, not P24)",
         "generate": "Generate Excel",
         "invalid_email": "Please enter a valid email ending with @devoteam.com",
         "success": "Excel file generated successfully!",
         "developer": "Script by Vladislav Slugin / vsdev.top",
         "select_language": "Select Language",
         "booking_details": "Booking Details",
         "start_date_must_be_before_end_date": "Start date must be before end date.",
         "select_at_least_one_weekday": "Please select at least one weekday."
    },
    "de": {
         "title": "Langzeitbuchung",
         "start_date": "Startdatum",
         "end_date": "Enddatum",
         "weekdays": "Wochentage auswählen",
         "monday": "Montag",
         "tuesday": "Dienstag",
         "wednesday": "Mittwoch",
         "thursday": "Donnerstag",
         "friday": "Freitag",
         "first_half": "Erste Hälfte",
         "second_half": "Zweite Hälfte",
         "private": "Privat",
         "email": "E-Mail",
         "seat": "Platznummer (z. B. 24, nicht P24)",
         "generate": "Excel generieren",
         "invalid_email": "Bitte geben Sie eine gültige E-Mail mit @devoteam.com an",
         "success": "Excel-Datei erfolgreich generiert!",
         "developer": "Script von Vladislav Slugin / vsdev.top",
         "select_language": "Sprache wählen",
         "booking_details": "Buchungsdetails",
         "start_date_must_be_before_end_date": "Startdatum muss vor Enddatum liegen.",
         "select_at_least_one_weekday": "Bitte wählen Sie mindestens einen Wochentag."
    }
}

# Global variable for current language (default is English)
current_lang = "en"

def t(key):
    """Return the translated text for the given key based on the current language."""
    return translations[current_lang].get(key, key)

class BookingApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(t("title"))
        self.geometry("500x600")
        self.resizable(False, False)

        # Variables for UI elements
        self.language_var = tk.StringVar(value=current_lang)
        self.start_date_var = tk.StringVar()
        self.end_date_var = tk.StringVar()
        self.email_var = tk.StringVar()
        self.seat_var = tk.StringVar()
        self.first_half_var = tk.BooleanVar(value=True)
        self.second_half_var = tk.BooleanVar(value=True)
        self.private_var = tk.BooleanVar(value=False)

        # Variables for weekdays (0 = Monday, ... 4 = Friday)
        self.weekday_vars = {
            0: tk.BooleanVar(value=True),  # Monday
            1: tk.BooleanVar(value=True),  # Tuesday
            2: tk.BooleanVar(value=True),  # Wednesday
            3: tk.BooleanVar(value=True),  # Thursday
            4: tk.BooleanVar(value=True),  # Friday
        }

        self.create_widgets()

    def create_widgets(self):
        # --- Language Selection Frame ---
        self.lang_frame = ttk.Frame(self)
        self.lang_frame.pack(pady=10)

        self.lang_label = ttk.Label(self.lang_frame, text=t("select_language"))
        self.lang_label.pack(side=tk.LEFT, padx=(0,5))

        self.lang_combo = ttk.Combobox(
            self.lang_frame, textvariable=self.language_var, state="readonly", values=["en", "de"]
        )
        self.lang_combo.pack(side=tk.LEFT)
        self.lang_combo.bind("<<ComboboxSelected>>", self.change_language)

        # --- Header Label ---
        self.header_label = ttk.Label(self, text=t("title"), font=("Arial", 16, "bold"))
        self.header_label.pack(pady=10)

        # --- Date Range Selection Frame ---
        self.period_frame = ttk.LabelFrame(self, text=f"{t('start_date')} / {t('end_date')}")
        self.period_frame.pack(padx=10, pady=10, fill="x")

        # Start Date Label and Calendar
        self.start_label = ttk.Label(self.period_frame, text=t("start_date"))
        self.start_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.start_cal = DateEntry(self.period_frame, textvariable=self.start_date_var, date_pattern='dd.mm.yyyy')
        self.start_cal.grid(row=0, column=1, padx=5, pady=5)

        # End Date Label and Calendar
        self.end_label = ttk.Label(self.period_frame, text=t("end_date"))
        self.end_label.grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.end_cal = DateEntry(self.period_frame, textvariable=self.end_date_var, date_pattern='dd.mm.yyyy')
        self.end_cal.grid(row=1, column=1, padx=5, pady=5)

        # --- Weekdays Selection Frame ---
        self.weekdays_frame = ttk.LabelFrame(self, text=t("weekdays"))
        self.weekdays_frame.pack(padx=10, pady=10, fill="x")

        self.weekday_checks = {}
        weekday_keys = ["monday", "tuesday", "wednesday", "thursday", "friday"]
        for i, key in enumerate(weekday_keys):
            cb = ttk.Checkbutton(self.weekdays_frame, text=t(key), variable=self.weekday_vars[i])
            cb.grid(row=0, column=i, padx=5, pady=5)
            self.weekday_checks[i] = cb

        # --- Booking Details Frame ---
        self.booking_frame = ttk.LabelFrame(self, text=t("booking_details"))
        self.booking_frame.pack(padx=10, pady=10, fill="x")

        # First Half Checkbutton
        self.first_half_check = ttk.Checkbutton(
            self.booking_frame, text=t("first_half"), variable=self.first_half_var
        )
        self.first_half_check.grid(row=0, column=0, padx=5, pady=5, sticky="w")

        # Second Half Checkbutton
        self.second_half_check = ttk.Checkbutton(
            self.booking_frame, text=t("second_half"), variable=self.second_half_var
        )
        self.second_half_check.grid(row=0, column=1, padx=5, pady=5, sticky="w")

        # Private Booking Checkbutton
        self.private_check = ttk.Checkbutton(
            self.booking_frame, text=t("private"), variable=self.private_var
        )
        self.private_check.grid(row=0, column=2, padx=5, pady=5, sticky="w")

        # Email Label and Entry
        self.email_label = ttk.Label(self.booking_frame, text=t("email"))
        self.email_label.grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.email_entry = ttk.Entry(self.booking_frame, textvariable=self.email_var, width=30)
        self.email_entry.grid(row=1, column=1, padx=5, pady=5, columnspan=2, sticky="w")

        # Seat Label and Entry
        self.seat_label = ttk.Label(self.booking_frame, text=t("seat"))
        self.seat_label.grid(row=2, column=0, padx=5, pady=5, sticky="w")
        self.seat_entry = ttk.Entry(self.booking_frame, textvariable=self.seat_var, width=10)
        self.seat_entry.grid(row=2, column=1, padx=5, pady=5, sticky="w")

        # --- Generate Excel Button ---
        self.generate_button = ttk.Button(self, text=t("generate"), command=self.generate_excel)
        self.generate_button.pack(pady=20)

        # --- Developer Information ---
        self.developer_label = ttk.Label(self, text=t("developer"), font=("Arial", 8))
        self.developer_label.pack(side="bottom", pady=5)

    def change_language(self, event=None):
        """Change the current language and update the UI texts."""
        global current_lang
        current_lang = self.language_var.get()
        self.update_language()

    def update_language(self):
        """Update the text of all UI elements based on the selected language."""
        self.title(t("title"))
        self.header_label.config(text=t("title"))
        self.lang_label.config(text=t("select_language"))
        self.period_frame.config(text=f"{t('start_date')} / {t('end_date')}")
        self.start_label.config(text=t("start_date"))
        self.end_label.config(text=t("end_date"))

        # Update weekdays checkbuttons
        weekday_keys = ["monday", "tuesday", "wednesday", "thursday", "friday"]
        for i, key in enumerate(weekday_keys):
            self.weekday_checks[i].config(text=t(key))

        self.booking_frame.config(text=t("booking_details"))
        self.first_half_check.config(text=t("first_half"))
        self.second_half_check.config(text=t("second_half"))
        self.private_check.config(text=t("private"))
        self.email_label.config(text=t("email"))
        self.seat_label.config(text=t("seat"))
        self.generate_button.config(text=t("generate"))
        self.developer_label.config(text=t("developer"))

    def generate_excel(self):
        """Generate the Excel file based on the user input."""
        start_date = self.start_cal.get_date()
        end_date = self.end_cal.get_date()
        email = self.email_var.get().strip()
        seat = self.seat_var.get().strip()

        # Validate email domain
        if not email.lower().endswith("@devoteam.com"):
            messagebox.showerror(t("title"), t("invalid_email"))
            return

        # Validate date range
        if start_date > end_date:
            messagebox.showerror(t("title"), t("start_date_must_be_before_end_date"))
            return

        # Get list of selected weekdays (0 = Monday, ... 4 = Friday)
        selected_weekdays = [day for day, var in self.weekday_vars.items() if var.get()]
        if not selected_weekdays:
            messagebox.showerror(t("title"), t("select_at_least_one_weekday"))
            return

        # Create a new Excel workbook and worksheet
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Bookings"
        headers = ["id", "seatID", "date", "timestamp", "firstHalf", "secondHalf", "email", "private"]
        ws.append(headers)

        current_day = start_date
        while current_day <= end_date:
            if current_day.weekday() in selected_weekdays:
                # Format the date as dd.mm.yyyy (e.g., 07.03.2024)
                date_str = current_day.strftime("%d.%m.%Y")

                # Get timestamp in milliseconds for the start of the day
                dt_obj = datetime.combine(current_day, datetime.min.time())
                timestamp_ms = int(dt_obj.timestamp() * 1000)

                # Generate a random UUID
                random_id = str(uuid.uuid4())

                # Create seatID with prefix "P"
                seatID = "P" + seat

                # Get boolean values for first and second half and private booking
                first_half = self.first_half_var.get()
                second_half = self.second_half_var.get()
                private = self.private_var.get()

                # Append the row to the worksheet
                row = [
                    random_id,
                    seatID,
                    date_str,
                    timestamp_ms,
                    str(first_half).upper(),
                    str(second_half).upper(),
                    email,
                    str(private).upper()
                ]
                ws.append(row)
            current_day += timedelta(days=1)

        # Ask user for the file path to save the Excel file
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            title="Save Excel File"
        )
        if file_path:
            try:
                wb.save(file_path)
                messagebox.showinfo(t("title"), t("success"))
            except Exception as e:
                messagebox.showerror(t("title"), str(e))

if __name__ == "__main__":
    app = BookingApp()
    app.mainloop()
