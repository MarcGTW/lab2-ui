# Laboratory Work 2 — Examination Ticket Generator

A desktop application built with Python and Tkinter/ttk. It reads student lists from Excel, examination tickets from Word, and persists a complete assignment history into an Excel journal.

## Features

- **Modern Dashboard UI**: Sleek card-based interface with dynamic theme switching (Dark 🌙 / Light ☀️ modes).
- **Group Progress Tracking**: Visual progress bar showing the percentage and count of students who have already received their tickets.
- **Smart Candidate Profile & Live Search**: Fast interactive filtering for student selection, plus automatic status badges distinguishing first-time exams (`✨ First Issue`) from repeat attempts (`⚠️ Reissue`).
- **Live Assignment History**: An interactive real-time table displaying past ticket assignments directly from `results.xlsx` with filtering and detailed inspection on double-click.
- **Official Examination Ticket Window**: Formatted as an academic exam slip with question cards, quick copy-to-clipboard, export to `.txt`, and immediate focus restoration.
- **Atomic Persistence**: Thread-safe, atomic file replacement using temporary workbooks to prevent journal corruption.
- **Keyboard Hotkeys**: `Enter` to generate, `Esc` to close ticket modal, `F5` / `Ctrl+R` to reload sources, `Ctrl+T` to toggle theme, `Ctrl+O` to open the journal in Excel.

## Getting Started

### Prerequisites

Python 3.10+ with Tkinter installed.

### Installation

Clone the repository and install required dependencies:

```powershell
py -m pip install -r requirements.txt
```

### Running the Application

To run the application from the project root:

```powershell
py main.py
```

*(If using an activated virtual environment, replace `py` with `python`).*

By default, files are loaded from the application's directory. To specify a custom working folder:

```powershell
py main.py --data-dir "C:\ExamData"
```

Demo files (`students.xlsx` and `tickets.docx`) are included: two groups with students, an empty group, and sample tickets. To generate fresh demo files in a separate folder without overwriting existing ones:

```powershell
py create_demo_data.py --data-dir demo
```

You can also click the **`🧪 Demo Files`** button directly in the application header.

## Usage Guide

1. **Select Academic Group**: Choose a group from the dropdown. The student list and progress metrics will update automatically.
2. **Select Student**: Use the search field to filter names or select directly from the dropdown. The candidate profile card displays their previous exam status.
3. **Generate Ticket**: Click **«Сгенерировать экзаменационный билет»** (or press `Enter`). The assignment is safely persisted to `results.xlsx` before the ticket window appears.
4. **View Ticket**: The ticket window displays the student's identity, ticket number, and three examination questions. You can copy the ticket to the clipboard or export it to `.txt`.
5. **Next Student**: Press `Esc` or close the window. The application resets the selection and immediately returns focus to the student selector.

### Business Rules & Repeat Policy

- A student is uniquely identified by the tuple `(Group, Surname, Name)`.
- If a student has already received a ticket, the application reuses the exact ticket number from the **first** entry in `results.xlsx` and appends a new record marked with `да`.
- For new students, a ticket is chosen uniformly at random from the set of valid tickets and appended with `нет`.
- Previous records are immutable and append-only.

## File Specifications

- `students.xlsx`: Each sheet represents an academic group (sheet title = group name). Row 1 contains column headers; row 2 onward contains Column A (`Фамилия` / Surname) and Column B (`Имя` / Name). Blank rows are skipped; incomplete rows are logged and ignored. Empty groups are supported.
- `tickets.docx`: Each ticket starts with a standalone paragraph `Билет N` (where N is a positive integer), followed by three numbered paragraphs `1. Question`, `2. Question`, `3. Question`. Word native numbered lists (`List Number`) are supported. Blocks with fewer than 3 questions or malformed headings are skipped and logged to `app.log`.
- `results.xlsx`: Created automatically on the first run. Columns A–F: `Группа`, `Фамилия`, `Имя`, `Номер билета`, `Дата и время`, `Повтор (да/нет)`. Timestamps are saved as native Excel dates.
- **File Locking**: If `results.xlsx` is open and locked in Excel, the app shows a `Retry / Cancel` dialog. Retrying saves the exact pending record without re-randomizing or losing data.

## Running Tests

Run the comprehensive unit and integration test suite:

```powershell
py -m unittest discover -s tests -v
```

The tests verify Word parsing, group handling, empty groups, idempotent repeat assignments, file lock recovery, atomic file replacements, and Tkinter GUI events in isolated temporary directories.

## Project Structure

- `main.py` — Application entry point and CLI argument parser.
- `ticket_generator/core.py` — Data models (`Student`, `Ticket`), XLSX/DOCX parsers, and atomic `ResultJournal`.
- `ticket_generator/gui.py` — Modern Tkinter/ttk user interface, themes, and interactive dashboard.
- `create_demo_data.py` — Safe generator of demonstration datasets.
- `tests/` — Automated test suite (`test_core.py`, `test_gui.py`).
- `COPILOT_PROMPT.md` — Detailed technical specification.

## Version Control

Working branch: `lab2-ui-tickets`.

```powershell
git add .
git commit -m "Lab 2: UI-генератор билетов"
git push -u origin lab2-ui-tickets
```
