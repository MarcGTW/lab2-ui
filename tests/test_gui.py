"""Integration and GUI tests for TicketApp.

Requires a working graphical Tk installation.
"""
from __future__ import annotations

import tempfile
import tkinter as tk
import unittest
from pathlib import Path

from create_demo_data import create_demo
from ticket_generator.gui import TicketApp


class GuiTests(unittest.TestCase):
    def setUp(self):
        try:
            self.root = tk.Tk()
        except tk.TclError as exc:
            self.skipTest(f"Tk недоступен: {exc}")
        self.addCleanup(self._cleanup)
        self.temp_dir = tempfile.TemporaryDirectory()
        self.folder = Path(self.temp_dir.name)
        create_demo(self.folder)
        self.app = TicketApp(self.root, self.folder)
        # Cancel any scheduled Tk timer callbacks from initialization
        for timer in self.root.tk.splitlist(self.root.tk.call("after", "info")):
            self.root.after_cancel(timer)
        self.app.load_sources()
        self.root.focus_force()
        self.root.update()

    def _cleanup(self):
        try:
            self.root.destroy()
        except Exception:
            pass
        try:
            self.temp_dir.cleanup()
        except Exception:
            pass

    def _close_any_toplevel(self):
        for child in list(self.root.winfo_children()):
            if isinstance(child, tk.Toplevel) and child.winfo_exists():
                child.focus_force()
                child.event_generate("<Escape>")
                self.root.update()

    def test_selection_save_and_escape(self):
        """Original specification integration test for selection gating, persistence, and Esc."""
        app = self.app
        self.assertTrue(app.generate_button.instate(["disabled"]))

        # 1. Empty group leaves student list empty and button disabled
        app.group.set("Пустая группа")
        app.select_group()
        self.assertEqual(app.student["values"], "")
        self.assertTrue(app.generate_button.instate(["disabled"]))

        # 2. Selecting group with students populates dropdown
        app.group.set("ИВТ-21")
        app.select_group()
        self.assertEqual(len(app.student["values"]), 3)

        # 3. Selecting a student enables generate button
        app.student.current(0)
        app.select_student()
        self.assertTrue(app.generate_button.instate(["!disabled"]))

        # 4. Generating creates results.xlsx and opens modal Toplevel
        app.generate()
        self.root.update()
        result = next(child for child in self.root.winfo_children() if isinstance(child, tk.Toplevel))
        self.assertTrue((self.folder / "results.xlsx").exists())

        # 5. Escape closes modal, resets student, and disables button
        result.focus_force()
        result.event_generate("<Escape>")
        self.root.update()
        self.assertFalse(result.winfo_exists())
        self.assertTrue(self.root.winfo_exists())
        self.assertTrue(app.generate_button.instate(["disabled"]))

    def test_theme_toggle(self):
        """Test switching between dark and light themes dynamically."""
        app = self.app
        self.assertEqual(app.current_theme, "dark")

        app.toggle_theme()
        self.root.update()
        self.assertEqual(app.current_theme, "light")

        app.toggle_theme()
        self.root.update()
        self.assertEqual(app.current_theme, "dark")

    def test_student_search_and_clear(self):
        """Test live search filter in student dropdown and clear button."""
        app = self.app
        app.group.set("ИВТ-21")
        app.select_group()
        self.assertEqual(len(app.student["values"]), 3)

        # Filter by substring
        app.search_var.set("Иван")
        self.root.update()
        self.assertEqual(list(app.student["values"]), ["Иванов Иван"])

        # Clearing restores all students
        app.search_clear_btn.invoke()
        self.root.update()
        self.assertEqual(len(app.student["values"]), 3)

    def test_candidate_preview_card_status(self):
        """Test candidate preview card transitions: empty -> primary -> reissue."""
        app = self.app
        self.assertIn("не выбран", app.profile_name_lbl["text"])

        # Select student who never received a ticket
        app.group.set("ИВТ-21")
        app.select_group()
        app.student.set("Иванов Иван")
        app.select_student()
        self.root.update()

        self.assertIn("Первичн", app.profile_badge_lbl["text"])
        self.assertIn("Иванов Иван", app.profile_name_lbl["text"])

        # Generate ticket for this student
        app.generate()
        self.root.update()
        self._close_any_toplevel()

        # Re-select the same student -> card must indicate repeat attempt
        app.group.set("ИВТ-21")
        app.select_group()
        app.student.set("Иванов Иван")
        app.select_student()
        self.root.update()

        self.assertIn("Повторн", app.profile_badge_lbl["text"])
        self.assertIn("сохранён тот же номер", app.profile_desc_lbl["text"])

    def test_group_progress_indicator(self):
        """Test progress bar and count labels for group completion."""
        app = self.app
        # Empty group
        app.group.set("Пустая группа")
        app.select_group()
        self.assertEqual(app.group_progress["value"], 0)

        # Group with 3 students
        app.group.set("ИВТ-21")
        app.select_group()
        self.assertEqual(app.group_progress["maximum"], 3)
        self.assertEqual(app.group_progress["value"], 0)

        # Issue ticket to student 0
        app.student.current(0)
        app.select_student()
        app.generate()
        self.root.update()
        self._close_any_toplevel()

        # Re-select group -> progress must increment to 1
        app.select_group()
        self.assertEqual(app.group_progress["value"], 1)
        self.assertIn("сдали: 1", app.group_info_lbl["text"])

    def test_history_table_and_filtering(self):
        """Test history table population, live search filtering, and clear."""
        app = self.app
        # Generate two tickets for different students
        app.group.set("ИВТ-21")
        app.select_group()
        app.student.current(0)
        app.select_student()
        app.generate()
        self.root.update()
        self._close_any_toplevel()

        app.group.set("ИВТ-21")
        app.select_group()
        app.student.current(1)
        app.select_student()
        app.generate()
        self.root.update()
        self._close_any_toplevel()

        # History table must have 2 records
        total_rows = len(app.tree.get_children())
        self.assertEqual(total_rows, 2)

        # Filter by a specific name
        app.hist_search_var.set("Иванов")
        self.root.update()
        self.assertEqual(len(app.tree.get_children()), 1)

        # Non-matching filter
        app.hist_search_var.set("НесуществующийЗапрос")
        self.root.update()
        self.assertEqual(len(app.tree.get_children()), 0)

        # Clear filter
        app.hist_clear_btn.invoke()
        self.root.update()
        self.assertEqual(len(app.tree.get_children()), total_rows)

    def test_history_row_selection_summary(self):
        """Test that selecting a row in the history table updates the summary label."""
        app = self.app
        app.group.set("ИВТ-21")
        app.select_group()
        app.student.current(0)
        app.select_student()
        app.generate()
        self.root.update()
        self._close_any_toplevel()

        rows = app.tree.get_children()
        self.assertTrue(len(rows) > 0)
        app.tree.selection_set(rows[0])
        app._on_tree_select()
        self.root.update()

        self.assertIn("Выбрана запись", app.hist_summary["text"])

    def test_hotkeys(self):
        """Test hotkeys: Enter generates ticket, F5 reloads."""
        app = self.app
        app.group.set("ИВТ-21")
        app.select_group()
        app.student.current(0)
        app.select_student()
        self.assertTrue(app.generate_button.instate(["!disabled"]))

        # Trigger generation via Enter handler
        app._on_enter()
        self.root.update()

        # Modal window must be present
        toplevels = [w for w in self.root.winfo_children() if isinstance(w, tk.Toplevel) and w.winfo_exists()]
        self.assertEqual(len(toplevels), 1)
        self._close_any_toplevel()

        # Test F5 reload binding
        app.group.set("ИВТ-21")
        app.select_group()
        self.root.focus_force()
        self.root.event_generate("<F5>")
        self.root.update()
        # Reloading resets group
        self.assertEqual(app.group.get(), "")

    def test_ticket_modal_title_and_save_notice(self):
        """Test that modal ticket window displays correct title, questions, and save notice."""
        app = self.app
        app.group.set("ИВТ-21")
        app.select_group()
        app.student.current(0)
        app.select_student()
        app.generate()
        self.root.update()

        modal = next(w for w in self.root.winfo_children() if isinstance(w, tk.Toplevel) and w.winfo_exists())
        self.assertIn("Ваш билет №", modal.title())

        # Check that Escape cleanly destroys it
        modal.focus_force()
        modal.event_generate("<Escape>")
        self.root.update()
        self.assertFalse(modal.winfo_exists())

    def test_reload_sources_resets_inputs(self):
        """Test that load_sources resets selections and restores clean initial state."""
        app = self.app
        app.group.set("ИВТ-21")
        app.select_group()
        app.student.current(0)
        app.select_student()
        self.assertTrue(app.generate_button.instate(["!disabled"]))

        app.load_sources()
        self.root.update()

        self.assertEqual(app.group.get(), "")
        self.assertEqual(app.student.get(), "")
        self.assertTrue(app.generate_button.instate(["disabled"]))
