import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from docx import Document
from openpyxl import Workbook, load_workbook

from ticket_generator.core import DataError, HEADERS, ResultJournal, Student, Ticket, read_students, read_tickets


class TemporaryFileTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.folder = Path(self.temp.name)

    def document(self, paragraphs):
        path = self.folder / "tickets.docx"
        doc = Document()
        for text in paragraphs:
            doc.add_paragraph(text)
        doc.save(path)
        return path


class FileTests(TemporaryFileTests):
    def test_parse_valid_tickets(self):
        path = self.document(["Билет 1", "1. Первый", "2. Второй", "3. Третий",
                              "Билет 2", "1. A", "2. B", "3. C"])
        tickets = read_tickets(path)
        self.assertEqual(tuple(tickets), (1, 2))
        self.assertEqual(tickets[1].questions, ("Первый", "Второй", "Третий"))

    def test_bad_blocks_skipped_without_mixing_questions(self):
        path = self.document(["Билет 1", "1. A", "2. B", "Билет №2", "1. C", "2. D", "3. E",
                              "Билет 7", "1. X", "2. Y", "3. Z"])
        with self.assertLogs("ticket_generator.core", level="WARNING"):
            tickets = read_tickets(path)
        self.assertEqual(tuple(tickets), (7,))
        self.assertEqual(tickets[7].questions, ("X", "Y", "Z"))

    def test_native_word_numbered_list(self):
        path = self.folder / "tickets.docx"
        doc = Document()
        doc.add_paragraph("Билет 3")
        for question in ("Один", "Два", "Три"):
            doc.add_paragraph(question, style="List Number")
        doc.save(path)
        self.assertEqual(read_tickets(path)[3].questions, ("Один", "Два", "Три"))

    def test_unrecognized_heading_cannot_complete_previous_ticket(self):
        path = self.document(["Билет 1", "1. A", "2. B", "Ticket two", "1. C", "2. D", "3. E",
                              "Билет 3", "1. X", "2. Y", "3. Z"])
        with self.assertLogs("ticket_generator.core"):
            self.assertEqual(tuple(read_tickets(path)), (3,))

    def test_no_valid_tickets(self):
        path = self.document(["Билет 1", "1. Только один"])
        with self.assertLogs("ticket_generator.core"), self.assertRaises(DataError):
            read_tickets(path)

    def test_duplicate_ticket_keeps_first(self):
        path = self.document(["Билет 1", "1. A", "2. B", "3. C", "Билет 1", "1. D", "2. E", "3. F"])
        with self.assertLogs("ticket_generator.core"):
            tickets = read_tickets(path)
        self.assertEqual(tickets[1].questions, ("A", "B", "C"))

    def test_students_by_group_and_empty_sheet(self):
        path = self.folder / "students.xlsx"
        book = Workbook()
        book.active.title = "ИВТ-21"
        book.active.append(("Фамилия", "Имя"))
        book.active.append((" Иванов ", "Иван"))
        other = book.create_sheet("ИВТ-22")
        other.append(("Фамилия", "Имя"))
        other.append(("Петров", "Пётр"))
        book.create_sheet("Пустая").append(("Фамилия", "Имя"))
        book.save(path)
        book.close()
        groups = read_students(path)
        self.assertEqual(groups["ИВТ-21"], [Student("ИВТ-21", "Иванов", "Иван")])
        self.assertEqual(groups["ИВТ-22"][0].display_name, "Петров Пётр")
        self.assertEqual(groups["Пустая"], [])

    def test_missing_and_corrupt_sources(self):
        for reader, filename in ((read_students, "students.xlsx"), (read_tickets, "tickets.docx")):
            with self.subTest(filename=filename):
                path = self.folder / filename
                with self.assertRaises(DataError):
                    reader(path)
                path.write_text("corrupt", encoding="utf-8")
                with self.assertRaises(DataError):
                    reader(path)


class JournalTests(TemporaryFileTests):
    def setUp(self):
        super().setUp()
        self.journal = ResultJournal(self.folder / "results.xlsx")
        self.student = Student("Группа 1", "Иванов", "Иван")
        self.tickets = {1: Ticket(1, ("A", "B", "C")), 7: Ticket(7, ("D", "E", "F"))}

    def rows(self):
        book = load_workbook(self.journal.path)
        try:
            return list(book.active.values)
        finally:
            book.close()

    def test_initialization_is_idempotent(self):
        self.journal.initialize()
        original = self.journal.path.read_bytes()
        self.journal.initialize()
        self.assertEqual(self.journal.path.read_bytes(), original)
        self.assertEqual(self.rows(), [HEADERS])

    def test_repeat_after_restart_appends_and_preserves_history(self):
        ticket, repeated = self.journal.assign(self.student, self.tickets)
        self.assertFalse(repeated)
        first_rows = self.rows()
        restarted = ResultJournal(self.journal.path)
        with patch("ticket_generator.core.random.choice", side_effect=AssertionError("Must not draw again")):
            second, repeated = restarted.assign(self.student, self.tickets)
        self.assertTrue(repeated)
        self.assertEqual(ticket, second)
        rows = self.rows()
        self.assertEqual(rows[:2], first_rows)
        self.assertEqual(len(rows), 3)
        self.assertEqual(rows[1][3], rows[2][3])
        self.assertEqual((rows[1][5], rows[2][5]), ("нет", "да"))
        self.assertLessEqual(rows[1][4], rows[2][4])

    def test_first_record_wins_if_later_record_differs(self):
        self.journal.initialize()
        book = load_workbook(self.journal.path)
        for number in (7, 1):
            book.active.append(("Группа 1", "Иванов", "Иван", number, "2026-01-01", "нет"))
        book.save(self.journal.path)
        book.close()
        ticket, repeated = self.journal.assign(self.student, self.tickets)
        self.assertTrue(repeated)
        self.assertEqual(ticket.number, 7)
        self.assertEqual(self.rows()[-1][3], 7)

    def test_identity_includes_group_and_both_names(self):
        students = [self.student, Student("Группа 2", "Иванов", "Иван"),
                    Student("Группа 1", "Иванов", "Пётр"), Student("Группа 1", "Петров", "Иван")]
        for person in students:
            self.assertFalse(self.journal.assign(person, self.tickets)[1])
        self.assertEqual(len(self.rows()), 5)

    def test_locked_file_preserves_old_bytes_and_retry_succeeds(self):
        self.journal.initialize()
        original = self.journal.path.read_bytes()
        with patch("ticket_generator.core.os.replace", side_effect=PermissionError("Excel lock")):
            with self.assertRaises(PermissionError):
                self.journal.assign(self.student, self.tickets)
        self.assertEqual(self.journal.path.read_bytes(), original)
        self.assertEqual(list(self.folder.glob("*.xlsx")), [self.journal.path])
        self.journal.assign(self.student, self.tickets)
        self.assertEqual(len(self.rows()), 2)

    def test_missing_previous_ticket_does_not_reassign(self):
        self.journal.assign(self.student, {7: self.tickets[7]})
        original = self.journal.path.read_bytes()
        with self.assertRaises(DataError):
            self.journal.assign(self.student, {1: self.tickets[1]})
        self.assertEqual(self.journal.path.read_bytes(), original)

    def test_invalid_journal_not_overwritten(self):
        self.journal.path.write_bytes(b"not an xlsx")
        with self.assertRaises(DataError):
            self.journal.assign(self.student, self.tickets)
        self.assertEqual(self.journal.path.read_bytes(), b"not an xlsx")

    def test_wrong_headers_not_overwritten(self):
        book = Workbook()
        book.active.append(("Другие данные",))
        book.save(self.journal.path)
        book.close()
        original = self.journal.path.read_bytes()
        with self.assertRaises(DataError):
            self.journal.initialize()
        self.assertEqual(self.journal.path.read_bytes(), original)


if __name__ == "__main__":
    unittest.main()
