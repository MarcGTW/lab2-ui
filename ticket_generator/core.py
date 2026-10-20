from __future__ import annotations

import logging
import os
import random
import re
import tempfile
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

from docx import Document
from openpyxl import Workbook, load_workbook

logger = logging.getLogger(__name__)
HEADERS = ("Группа", "Фамилия", "Имя", "Номер билета", "Дата и время", "Повтор (да/нет)")


class DataError(Exception):
    """A source or journal cannot be used safely."""


@dataclass(frozen=True)
class Student:
    group: str
    surname: str
    name: str

    @property
    def display_name(self) -> str:
        return f"{self.surname} {self.name}"


@dataclass(frozen=True)
class Ticket:
    number: int
    questions: tuple[str, ...]


def read_students(path: Path) -> dict[str, list[Student]]:
    if not path.is_file():
        raise DataError(f"Не найден файл {path.name}. Поместите его в папку данных.")
    try:
        workbook = load_workbook(path, read_only=True, data_only=True)
        try:
            groups = {}
            for sheet in workbook:
                students = []
                for row_number, row in enumerate(sheet.iter_rows(min_row=2, max_col=2, values_only=True), 2):
                    surname, name = (str(value).strip() if value is not None else "" for value in row)
                    if not surname and not name:
                        continue
                    if not surname or not name:
                        logger.warning("%s, строка %s: неполное имя, строка пропущена", sheet.title, row_number)
                        continue
                    student = Student(sheet.title, surname, name)
                    if student not in students:
                        students.append(student)
                groups[sheet.title] = students
            return groups
        finally:
            workbook.close()
    except Exception as exc:
        raise DataError(f"Не удалось прочитать {path.name}: {exc}") from exc


def read_tickets(path: Path) -> dict[int, Ticket]:
    if not path.is_file():
        raise DataError(f"Не найден файл {path.name}. Поместите его в папку данных.")
    try:
        document = Document(path)
    except Exception as exc:
        raise DataError(f"Не удалось прочитать {path.name}: {exc}") from exc
    tickets: dict[int, Ticket] = {}
    current: int | None = None
    questions: list[str] = []

    def finish() -> None:
        if current is None:
            return
        if len(questions) < 3:
            logger.warning("Билет %s: меньше трёх вопросов, билет пропущен", current)
        elif current in tickets:
            logger.warning("Повторный заголовок билета %s пропущен", current)
        else:
            tickets[current] = Ticket(current, tuple(questions[:3]))
            if len(questions) > 3:
                logger.warning("Билет %s: используются первые три вопроса", current)

    for paragraph in document.paragraphs:
        text = paragraph.text.strip()
        if not text:
            continue
        heading = re.fullmatch(r"Билет\s+([1-9]\d*)", text, re.IGNORECASE)
        if heading:
            finish()
            current, questions = int(heading.group(1)), []
        elif re.match(r"^билет\b", text, re.IGNORECASE):
            finish()
            logger.warning("Нераспознанный заголовок %r: блок пропущен", text)
            current, questions = None, []
        elif current is not None:
            question = re.fullmatch(r"\d+\.\s*(\S.*)", text)
            # Word can store list numbers as numbering properties, outside paragraph.text.
            numbered = (paragraph._p.pPr is not None and paragraph._p.pPr.numPr is not None)
            numbered = numbered or bool(paragraph.style and paragraph.style.name.startswith("List Number"))
            if question:
                questions.append(question.group(1).strip())
            elif numbered:
                questions.append(text)
            else:
                finish()
                logger.warning("Нераспознанный абзац %r: блок до следующего заголовка пропущен", text)
                current, questions = None, []
        else:
            logger.warning("Абзац вне распознанного билета пропущен: %r", text)
    finish()
    if not tickets:
        raise DataError("В tickets.docx нет корректных билетов с тремя вопросами. Подробности в app.log.")
    return tickets


class ResultJournal:
    def __init__(self, path: Path):
        self.path = path

    def _open(self):
        if not self.path.exists():
            book = Workbook()
            book.active.title = "Результаты"
            book.active.append(HEADERS)
            book.active.freeze_panes = "A2"
            for column, width in zip("ABCDEF", (20, 26, 24, 20, 24, 23)):
                book.active.column_dimensions[column].width = width
            return book
        try:
            book = load_workbook(self.path)
        except PermissionError:
            raise
        except Exception as exc:
            raise DataError(f"Не удалось прочитать results.xlsx: {exc}. Журнал не изменён.") from exc
        if tuple(cell.value for cell in book.active[1]) != HEADERS:
            book.close()
            raise DataError("Неверные заголовки results.xlsx. Журнал не изменён; проверьте формат файла.")
        return book

    def _save(self, book) -> None:
        # Replace only after a complete workbook has been written, preserving the old file on failure.
        temporary = None
        try:
            with tempfile.NamedTemporaryFile(dir=self.path.parent, suffix=".xlsx", delete=False) as stream:
                temporary = Path(stream.name)
            book.save(temporary)
            os.replace(temporary, self.path)
        finally:
            if temporary is not None:
                temporary.unlink(missing_ok=True)

    def initialize(self) -> None:
        book = self._open()
        try:
            if not self.path.exists():
                self._save(book)
        finally:
            book.close()

    def assign(self, student: Student, tickets: dict[int, Ticket]) -> tuple[Ticket, bool]:
        book = self._open()
        try:
            sheet = book.active
            number = None
            for row in sheet.iter_rows(min_row=2, values_only=True):
                if tuple(row[:3]) == (student.group, student.surname, student.name):
                    number = row[3]
                    if isinstance(number, bool) or not isinstance(number, (int, float)) or int(number) != number:
                        raise DataError("В первой записи студента указан некорректный номер билета.")
                    break
            repeated = number is not None
            if repeated:
                if number not in tickets:
                    raise DataError(f"Ранее назначенный билет № {number} отсутствует среди корректных билетов. Восстановите его в tickets.docx.")
                ticket = tickets[number]
            else:
                if not tickets:
                    raise DataError("Нет доступных билетов.")
                ticket = random.choice(list(tickets.values()))
            sheet.append((student.group, student.surname, student.name, ticket.number,
                          datetime.now(), "да" if repeated else "нет"))
            sheet.cell(sheet.max_row, 5).number_format = "DD.MM.YYYY HH:MM:SS"
            self._save(book)
            return ticket, repeated
        finally:
            book.close()
