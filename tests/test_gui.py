"""Integration test requires a working graphical Tk installation."""
import tempfile
import tkinter as tk
import unittest
from pathlib import Path

from create_demo_data import create_demo
from ticket_generator.gui import TicketApp


class GuiTests(unittest.TestCase):
    def test_selection_save_and_escape(self):
        try:
            root = tk.Tk()
        except tk.TclError as exc:
            self.skipTest(f"Tk недоступен: {exc}")
        self.addCleanup(root.destroy)
        with tempfile.TemporaryDirectory() as folder:
            create_demo(Path(folder))
            app = TicketApp(root, Path(folder))
            for timer in root.tk.splitlist(root.tk.call("after", "info")):
                root.after_cancel(timer)
            app.load_sources()
            self.assertTrue(app.generate_button.instate(["disabled"]))
            app.group.set("Пустая группа")
            app.select_group()
            self.assertEqual(app.student["values"], "")
            self.assertTrue(app.generate_button.instate(["disabled"]))
            app.group.set("ИВТ-21")
            app.select_group()
            self.assertEqual(len(app.student["values"]), 3)
            app.student.current(0)
            app.select_student()
            self.assertTrue(app.generate_button.instate(["!disabled"]))
            app.generate()
            root.update()
            result = next(child for child in root.winfo_children() if isinstance(child, tk.Toplevel))
            self.assertTrue((Path(folder) / "results.xlsx").exists())
            result.event_generate("<Escape>")
            root.update()
            self.assertFalse(result.winfo_exists())
            self.assertTrue(root.winfo_exists())
            self.assertTrue(app.generate_button.instate(["disabled"]))
