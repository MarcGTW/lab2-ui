import argparse
import logging
from pathlib import Path
import tkinter as tk

from ticket_generator.gui import TicketApp


def main():
    parser = argparse.ArgumentParser(description="Генератор экзаменационных билетов")
    parser.add_argument("--data-dir", type=Path, default=Path(__file__).resolve().parent,
                        help="Папка с students.xlsx, tickets.docx и results.xlsx")
    args = parser.parse_args()
    try:
        logging.basicConfig(filename=args.data_dir / "app.log", level=logging.INFO,
                            format="%(asctime)s %(levelname)s %(message)s", encoding="utf-8")
    except OSError:
        logging.basicConfig(level=logging.INFO)
    root = tk.Tk()
    TicketApp(root, args.data_dir)
    root.mainloop()


if __name__ == "__main__":
    main()
