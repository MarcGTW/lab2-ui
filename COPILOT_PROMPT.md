# Lab 2 — implementation specification for Copilot Chat

Implement a Python 3.10+ desktop examination ticket generator using Tkinter/ttk, openpyxl, and python-docx. Use a Russian UI. Separate GUI code from source parsing and assignment/journal logic. Use pathlib and a configurable data directory, defaulting to the application's directory.

## Input contracts

- `students.xlsx`: each worksheet is a group; preserve worksheet order and names. Row 1 contains headers; A is surname, B is given name; student rows start at row 2. Trim surrounding whitespace, skip blank rows, log and skip incomplete names. Preserve empty groups. The student identity is the tuple `(group, surname, given_name)`.
- `tickets.docx`: each ticket begins with a standalone paragraph matching `Билет N` with a positive integer N. Questions are paragraphs of the form `N. text`; also handle Word native numbered lists. Show exactly three questions per ticket. Skip tickets with fewer than three questions, log malformed headers and discarded blocks, keep parsing other tickets. Do not merge questions from a malformed block into a preceding ticket. Log duplicate IDs and retain the first valid ticket. Keep original ticket IDs after filtering. If more than three questions exist, use the first three and log it.
- Missing, unreadable or corrupt sources must produce clear Russian messages without terminating the application. With no valid tickets, disable generation. Provide a reload button so users can fix files without restarting.

## GUI behavior

1. On startup, populate a read-only group combobox from the workbook; initially select neither group nor student.
2. Selecting a group resets the student selection and populates the student combobox with `Surname GivenName`. For an empty group, keep it empty and disable generation.
3. Enable `Сгенерировать билет` only when a valid group and student are selected and valid tickets are available.
4. On generation, determine the ticket and persist the journal immediately. Display a modal child window only after saving succeeds. It contains `Ваш билет № N`, group, student, and three numbered questions. Use a scrollable area for long questions.
5. Escape and the child window's close button close only that window, reset the student selection, and return to the main window. The app remains ready for another student.
6. If Excel locks the output, show Retry/Cancel; Retry tries saving again, Cancel returns to selection. Failed writes must leave the prior journal intact and must not display an unsaved result.

## Assignment and journal

Create `results.xlsx` on first launch if absent, using columns A–F:
`Группа`, `Фамилия`, `Имя`, `Номер билета`, `Дата и время`, `Повтор (да/нет)`.

Before every generation, reopen the journal and scan rows in physical order. For a matching identity use the ticket ID from the FIRST matching row, regardless of later rows. Append a NEW row with that same ID, current local date/time and `да`. Never modify previous rows.

For a new student choose uniformly among valid ticket IDs and append a new row with `нет`. When all source tickets have consecutive IDs, this is equivalent to a random integer from 1 to N. After skipping malformed tickets, do not accidentally issue a missing ID or renumber the remaining tickets.

If a previously assigned ticket has been removed or is now invalid, show a clear error and require restoration of the source ticket; do not silently assign a different ticket. Reject a corrupt journal or unexpected journal headers without overwriting it. Save to a temporary sibling workbook, close the file, and atomically replace the destination; clean up temporary files on failure. Scope: one running instance at a time; no database or concurrent multi-process coordination is required.

## Tests and deliverables

Use standard unittest and temporary generated XLSX/DOCX fixtures. Cover valid Word parsing, malformed/incomplete tickets, native numbering, grouping and empty groups, missing/corrupt input, repeat assignment across service restarts, first-record precedence, identity separation by all three fields, immediate persistence and append-only history, locked-write failure/retry, missing previously assigned tickets, and journal corruption. Include a Tk integration test for selection gating and Escape, skipping it explicitly when Tk/display is unavailable.

Deliver runnable sources, requirements.txt, Russian README, sample students.xlsx/tickets.docx, a safe demo data generator which never overwrites existing inputs, and tests. Do not track real student data, runtime logs or the result journal in source control. No network, account, or installer functionality is needed.
