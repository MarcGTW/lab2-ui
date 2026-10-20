"""Create demonstration input files; never overwrite existing input."""
import argparse
from pathlib import Path

from docx import Document
from openpyxl import Workbook


def create_demo(folder: Path):
    folder.mkdir(parents=True, exist_ok=True)
    students = folder / "students.xlsx"
    tickets = folder / "tickets.docx"
    if not students.exists():
        book = Workbook()
        book.active.title = "ИВТ-21"
        for group, people in (("ИВТ-21", [("Иванов", "Иван"), ("Петрова", "Анна"), ("Сидоров", "Михаил")]),
                              ("ИВТ-22", [("Смирнова", "Мария"), ("Кузнецов", "Алексей")]),
                              ("Пустая группа", [])):
            sheet = book[group] if group in book.sheetnames else book.create_sheet(group)
            sheet.append(("Фамилия", "Имя"))
            for person in people:
                sheet.append(person)
            sheet.freeze_panes = "A2"
            sheet.column_dimensions["A"].width = 26
            sheet.column_dimensions["B"].width = 24
        book.save(students)
        book.close()
        print(f"Created {students.name}")
    if not tickets.exists():
        doc = Document()
        examples = [
            ("Что такое графический интерфейс пользователя?", "Чем событие отличается от обработчика события?", "Как работает главный цикл Tkinter?"),
            ("Как прочитать лист Excel с помощью openpyxl?", "Как организовать проверку входных данных?", "Для чего нужны unit-тесты?"),
            ("Как прочитать абзацы документа Word?", "Как обработать исключение при записи файла?", "Что такое идемпотентность назначения билета?"),
            ("Как связать два выпадающих списка?", "Как сохранить историю выдачи билетов?", "Чем коммит отличается от ветки Git?"),
            ("Как разделить интерфейс и бизнес-логику?", "Почему журнал сохраняется сразу после генерации?", "Как проверить повторную выдачу билета?"),
        ]
        for number, questions in enumerate(examples, 1):
            doc.add_paragraph(f"Билет {number}", style="Heading 1")
            for index, question in enumerate(questions, 1):
                doc.add_paragraph(f"{index}. {question}")
        doc.save(tickets)
        print(f"Created {tickets.name}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--data-dir", type=Path, default=Path(__file__).resolve().parent)
    create_demo(parser.parse_args().data_dir)
