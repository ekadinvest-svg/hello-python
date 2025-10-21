import sys
import json
import os
from pathlib import Path
from typing import Any
from datetime import datetime
os.environ.setdefault('QT_API', 'pyside6')
from matplotlib.figure import Figure
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
import matplotlib.dates as mdates

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QLabel, QLineEdit, QPushButton,
    QVBoxLayout, QHBoxLayout, QGridLayout, QWidget, QToolBar, QStatusBar,
    QTableWidget, QTableWidgetItem, QSizePolicy, QComboBox, QMenu,
    QInputDialog, QMessageBox
)
from PySide6.QtGui import QAction, QDoubleValidator, QIntValidator, QValidator
from PySide6.QtCore import Qt, QSize, QEvent


# טבלה שמאזנת עמודות לרוחב שווה בכל שינוי גודל
class EqualWidthTable(QTableWidget):
    def resizeEvent(self, event):
        super().resizeEvent(event)
        self._equalize_columns()

    def showEvent(self, event):
        super().showEvent(event)
        self._equalize_columns()

    def _equalize_columns(self):
        cols = self.columnCount()
        if cols <= 0:
            return
        # רוחב פנימי נטו
        total = self.viewport().width()
        w = max(50, total // cols)
        for i in range(cols):
            self.setColumnWidth(i, w)


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("אני מתאמן")
        self.resize(950, 560)

        # נתונים
        self.exercise_names: list[str] = []
        self.weights: list[str] = []
        self.sets_list: list[int] = []
        self.reps_list: list[int] = []
        self.last_reps_list: list[int] = []

        container = QWidget()
        container.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
        layout = QVBoxLayout(container)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(10)

    # Form: label and input on same row; inputs aligned in a single column
        form_width = 150
        form_height = 28

        form_layout = QGridLayout()
        form_layout.setHorizontalSpacing(12)
        form_layout.setVerticalSpacing(6)

        # --- שם תרגיל ---
        name_label = QLabel("שם תרגיל:")
        name_label.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        # use editable combo so user can type or pick from previous names
        self.input_name = QComboBox()
        self.input_name.setEditable(True)
        self.input_name.setInsertPolicy(QComboBox.NoInsert)
        # set combo size so the arrow button doesn't overlap the text
        self.input_name.setFixedSize(form_width, form_height)
        # placeholder on the internal line edit; make it slightly narrower than combo
        name_line = self.input_name.lineEdit()
        name_line.setPlaceholderText("לדוגמה: לחיצת חזה")
        name_line.setAlignment(Qt.AlignmentFlag.AlignLeft)
        name_line.setMinimumWidth(max(10, form_width - 28))
        name_line.setFixedHeight(form_height)
        name_line.textChanged.connect(self._update_add_enabled)
        form_layout.addWidget(name_label, 0, 0)
        form_layout.addWidget(self.input_name, 0, 1)

        # --- משקל ---
        weight_label = QLabel("משקל:")
        weight_label.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        self.input_weight = QLineEdit()
        self.input_weight.setPlaceholderText("לדוגמה: 60.5")
        self.input_weight.setAlignment(Qt.AlignmentFlag.AlignLeft)
        self.input_weight.setFixedSize(form_width, form_height)
        dbl_validator = QDoubleValidator(0.0, 10000.0, 3, self)
        dbl_validator.setNotation(QDoubleValidator.Notation.StandardNotation)
        self.input_weight.setValidator(dbl_validator)
        self.input_weight.textChanged.connect(self._update_add_enabled)
        form_layout.addWidget(weight_label, 1, 0)
        form_layout.addWidget(self.input_weight, 1, 1)

        # --- סטים ---
        sets_label = QLabel("סטים:")
        sets_label.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        self.input_sets = QLineEdit()
        self.input_sets.setPlaceholderText("לדוגמה: 4")
        self.input_sets.setAlignment(Qt.AlignmentFlag.AlignLeft)
        self.input_sets.setFixedSize(form_width, form_height)
        self.input_sets.setValidator(QIntValidator(1, 1000, self))
        self.input_sets.textChanged.connect(self._update_add_enabled)
        form_layout.addWidget(sets_label, 2, 0)
        form_layout.addWidget(self.input_sets, 2, 1)

        # --- חזרות ---
        reps_label = QLabel("חזרות:")
        reps_label.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        self.input_reps = QLineEdit()
        self.input_reps.setPlaceholderText("לדוגמה: 12")
        self.input_reps.setAlignment(Qt.AlignmentFlag.AlignLeft)
        self.input_reps.setFixedSize(form_width, form_height)
        self.input_reps.setValidator(QIntValidator(1, 1000, self))
        self.input_reps.textChanged.connect(self._update_add_enabled)
        form_layout.addWidget(reps_label, 3, 0)
        form_layout.addWidget(self.input_reps, 3, 1)

        # --- חזרות בסט האחרון ---
        last_reps_label = QLabel("סט אחרון:")
        last_reps_label.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        self.input_last_reps = QLineEdit()
        self.input_last_reps.setPlaceholderText("לדוגמה: 15")
        self.input_last_reps.setAlignment(Qt.AlignmentFlag.AlignLeft)
        self.input_last_reps.setFixedSize(form_width, form_height)
        self.input_last_reps.setValidator(QIntValidator(1, 1000, self))
        self.input_last_reps.textChanged.connect(self._update_add_enabled)
        form_layout.addWidget(last_reps_label, 4, 0)
        form_layout.addWidget(self.input_last_reps, 4, 1)

        # הוספת הטופס + אזור גרף בפריסה עליונה
        top_row = QHBoxLayout()
        # form widget
        form_widget = QWidget()
        form_widget.setLayout(form_layout)
        # graph widget (canvas)
        self.figure = Figure(figsize=(4, 3))
        self.canvas = FigureCanvas(self.figure)
        graph_container = QWidget()
        graph_layout = QVBoxLayout(graph_container)
        graph_layout.setContentsMargins(0, 0, 0, 0)
        graph_layout.addWidget(self.canvas)
        # place graph opposite the form
        top_row.addWidget(graph_container)
        top_row.addWidget(form_widget)
        layout.addLayout(top_row)

        # --- כפתורים ---
        self.btn_add = QPushButton("הוסף")
        self.btn_add.setEnabled(False)
        self.btn_add.clicked.connect(self.add_entry)

        self.btn_pop = QPushButton("נקה אחרון")
        self.btn_pop.setEnabled(False)
        self.btn_pop.clicked.connect(self.pop_last)

        buttons_row = QHBoxLayout()
        buttons_row.addWidget(self.btn_add)
        # save button
        self.btn_save = QPushButton("שמור")
        self.btn_save.clicked.connect(self.save_state)
        # plot button
        self.btn_plot = QPushButton("צור גרף")
        self.btn_plot.clicked.connect(self.plot_selected_exercise)
        buttons_row.addWidget(self.btn_plot)
        buttons_row.addWidget(self.btn_save)
        buttons_row.addWidget(self.btn_pop)
        buttons_row.addStretch(1)

        # --- טבלה (תאריך, שם, משקל, סטים, חזרות, סט אחרון) ---
        self.table = EqualWidthTable(0, 6)
        # enable custom context menu for deleting rows
        self.table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self._show_table_context_menu)
        self.table.setHorizontalHeaderLabels(["תאריך", "שם תרגיל", "משקל (Kg)", "סטים", "חזרות", "סט אחרון"])
        self.table.verticalHeader().setVisible(False)
        self.table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.table.setAlternatingRowColors(True)
        self.table.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        self.table.setStyleSheet(
            "QTableWidget {"
            "  background: #f8f8f8;"
            "  border: 1px solid #d9d9d9;"
            "  border-radius: 6px;"
            "  gridline-color: #cccccc;"
            "  alternate-background-color: #f1f1f1;"
            "}"
        )

        # double-click to edit date cell
        self.table.cellDoubleClicked.connect(self._edit_date_cell)

        # פריסה
        # form_layout already added above; now add buttons and table
        layout.addLayout(buttons_row)
        layout.addWidget(self.table)

        self.setCentralWidget(container)

        # טולבר + סטוס
        self._build_toolbar()
        self.status = QStatusBar()
        self.setStatusBar(self.status)
        self.status.showMessage("מוכן להזנה")

        # Enter מוסיף אם אפשר
        # for combo use its lineEdit
        name_line = self.input_name.lineEdit()
        for w in [name_line, self.input_weight, self.input_sets, self.input_reps, self.input_last_reps]:
            # QLineEdit has returnPressed
            try:
                w.returnPressed.connect(self._try_add_on_enter)
            except Exception:
                pass

        # ניווט עם ↑/↓ בין השדות
        self._inputs = [name_line, self.input_weight, self.input_sets, self.input_reps, self.input_last_reps]
        for w in self._inputs:
            w.installEventFilter(self)

    def plot_selected_exercise(self):
        # אם יש בחירה בטבלה - נשתמש בשם התרגיל מהשורה הנבחרת קודם
        name = self.input_name.currentText().strip()
        selected_rows = sorted({idx.row() for idx in self.table.selectedIndexes()})
        if selected_rows:
            # קח את השם מהשורה הראשונה שנבחרה
            sel_item = self.table.item(selected_rows[0], 1)
            if sel_item and sel_item.text().strip():
                name = sel_item.text().strip()

        if not name:
            self.status.showMessage("בחר שם תרגיל לקביעת גרף", 2000)
            return

        # אסוף נתונים מהטבלה עבור אותו שם; המטרה: להשתמש בתאריכים בעמודה 0 כציר X
        points: list[tuple[datetime, float]] = []
        for r in range(self.table.rowCount()):
            item = self.table.item(r, 1)
            if item and item.text() == name:
                date_item = self.table.item(r, 0)
                weight_item = self.table.item(r, 2)
                # parse weight
                try:
                    wtxt = weight_item.text().split()[0] if weight_item is not None else "0"
                    wval = float(wtxt.replace(",", "."))
                except Exception:
                    wval = 0.0
                # parse date (expecting YYYY-MM-DD)
                try:
                    dstr = date_item.text().strip() if date_item is not None else ""
                    dval = datetime.strptime(dstr, "%Y-%m-%d")
                except Exception:
                    # fallback to row index date (use today with offset)
                    dval = datetime.now()
                points.append((dval, wval))

        if not points:
            self.status.showMessage("אין רשומות עבור התרגיל הנבחר", 2000)
            return

        # מיין לפי תאריך
        points.sort(key=lambda x: x[0])
        xs = [p[0] for p in points]
        ys = [p[1] for p in points]

        # צייר גרף קווי עם ציר תאריכים
        self.figure.clear()
        ax = self.figure.add_subplot(111)
        ax.plot_date(xs, ys, '-o')
        # פורמט תאריכים על הציר
        ax.xaxis.set_major_locator(mdates.AutoDateLocator())
        ax.xaxis.set_major_formatter(mdates.DateFormatter('%Y-%m-%d'))
        self.figure.autofmt_xdate(rotation=30)
        # Hebrew (RTL) טקסטים - עטוף בסימון כיוון כדי שהמילה לא תופיע הפוכה
        RTL = '\u202B'
        POP = '\u202C'
        ax.set_title(RTL + f"גרף משקל: {name}" + POP)
        ax.set_xlabel(RTL + 'תאריך' + POP)
        ax.set_ylabel(RTL + 'משקל' + POP)
        ax.grid(True)
        self.canvas.draw()

        # טען מצב שמור אם קיים
        try:
            self.load_state()
        except Exception:
            pass

    # ---------- ניווט עם חצים ----------

    def eventFilter(self, obj, event):
        if event.type() == QEvent.Type.KeyPress and obj in self._inputs:
            key = event.key()
            idx = self._inputs.index(obj)
            if key == Qt.Key_Down:
                self._inputs[(idx + 1) % len(self._inputs)].setFocus()
                return True
            if key == Qt.Key_Up:
                self._inputs[(idx - 1) % len(self._inputs)].setFocus()
                return True
        return super().eventFilter(obj, event)

    # ---------- עזרים ----------

    def _build_toolbar(self):
        toolbar = QToolBar("כלים")
        toolbar.setIconSize(QSize(16, 16))
        self.addToolBar(toolbar)
        act_clear = QAction("נקה הכול", self)
        act_clear.triggered.connect(self.clear_all)
        toolbar.addAction(act_clear)

    def _weight_state(self, text: str) -> QValidator.State:
        v = self.input_weight.validator()
        if isinstance(v, QDoubleValidator):
            res: Any = v.validate(text, 0)
            if isinstance(res, tuple) and len(res) > 0 and isinstance(res[0], QValidator.State):
                return res[0]
        return QValidator.State.Invalid

    def _int_state(self, widget: QLineEdit) -> QValidator.State:
        v = widget.validator()
        if isinstance(v, QIntValidator):
            res: Any = v.validate(widget.text(), 0)
            if isinstance(res, tuple) and len(res) > 0 and isinstance(res[0], QValidator.State):
                return res[0]
        return QValidator.State.Invalid

    # ---------- לוגיקה ----------

    def _update_add_enabled(self):
        name_ok = self.input_name.currentText().strip() != ""
        weight_ok = self._weight_state(self.input_weight.text().strip().replace(",", ".")) == QValidator.State.Acceptable
        sets_ok = self._int_state(self.input_sets) == QValidator.State.Acceptable
        reps_ok = self._int_state(self.input_reps) == QValidator.State.Acceptable
        last_reps_ok = self._int_state(self.input_last_reps) == QValidator.State.Acceptable
        self.btn_add.setEnabled(name_ok and weight_ok and sets_ok and reps_ok and last_reps_ok)

    def _try_add_on_enter(self):
        if self.btn_add.isEnabled():
            self.add_entry()

    def add_entry(self):
        name = self.input_name.currentText().strip()
        weight_raw = self.input_weight.text().strip().replace(",", ".")
        sets_raw = self.input_sets.text().strip()
        reps_raw = self.input_reps.text().strip()
        last_reps_raw = self.input_last_reps.text().strip()

        if not (name and weight_raw and sets_raw and reps_raw and last_reps_raw):
            self.status.showMessage("מלא את כל השדות.", 2000)
            return

        try:
            weight_val = float(weight_raw)
            weight_str = f"{int(weight_val)}" if weight_val.is_integer() else f"{weight_val:.3f}".rstrip("0").rstrip(".")
            sets_val = int(sets_raw)
            reps_val = int(reps_raw)
            last_reps_val = int(last_reps_raw)
        except ValueError:
            self.status.showMessage("קלט לא תקין.", 2000)
            return

        # תאריך
        date_str = datetime.now().strftime("%Y-%m-%d")

        # הוספה לטבלה
        row = self.table.rowCount()
        self.table.insertRow(row)

        data = [date_str, name, f"{weight_str} Kg", sets_val, reps_val, last_reps_val]
        aligns = [
            Qt.AlignmentFlag.AlignHCenter,  # תאריך
            Qt.AlignmentFlag.AlignLeft,     # שם תרגיל
            Qt.AlignmentFlag.AlignHCenter,  # משקל
            Qt.AlignmentFlag.AlignHCenter,  # סטים
            Qt.AlignmentFlag.AlignHCenter,  # חזרות
            Qt.AlignmentFlag.AlignHCenter,  # סט אחרון
        ]
        for col, value in enumerate(data):
            item = QTableWidgetItem(str(value))
            item.setTextAlignment(aligns[col] | Qt.AlignmentFlag.AlignVCenter)
            self.table.setItem(row, col, item)

        # הוספת השם ל־combo אם חדש
        if name:
            existing = [self.input_name.itemText(i) for i in range(self.input_name.count())]
            if name not in existing:
                self.input_name.addItem(name)

        # ניקוי שדות (לשדה combo ננקה את ה-lineEdit בלבד)
        self.input_name.lineEdit().clear()
        for field in [self.input_weight, self.input_sets, self.input_reps, self.input_last_reps]:
            field.clear()
        self.input_name.setFocus()
        self._update_add_enabled()
        self.btn_pop.setEnabled(True)
        self.status.showMessage(f"התווסף: {name} ({weight_str} Kg, {sets_val}x{reps_val})", 2000)

    def pop_last(self):
        rows = self.table.rowCount()
        if rows > 0:
            # get name from last row before removing
            item = self.table.item(rows - 1, 1)
            name = item.text() if item is not None else None
            self.table.removeRow(rows - 1)
            # remove name from combo if no longer present in table
            if name:
                self._remove_name_if_unused(name)
            self.btn_pop.setEnabled(self.table.rowCount() > 0)
            self.status.showMessage("נמחק האחרון.", 2000)

    def clear_all(self):
        self.table.setRowCount(0)
        # clear combo's line edit only (don't remove saved names)
        try:
            self.input_name.lineEdit().clear()
        except Exception:
            # fallback if input_name is a QLineEdit
            try:
                self.input_name.clear()
            except Exception:
                pass
        for field in [self.input_weight, self.input_sets, self.input_reps, self.input_last_reps]:
            field.clear()
        self._update_add_enabled()
        self.btn_pop.setEnabled(False)
        self.status.showMessage("נוקה הכול", 2000)

    # ---------- שמירה/טעינה של מצב ----------
    def save_state(self):
        state = {
            "names": [self.input_name.itemText(i) for i in range(self.input_name.count())],
            "rows": []
        }
        for r in range(self.table.rowCount()):
            row_data = [self.table.item(r, c).text() if self.table.item(r, c) is not None else "" for c in range(self.table.columnCount())]
            state["rows"].append(row_data)
        path = Path.cwd() / "exercise_state.json"
        try:
            with open(path, "w", encoding="utf-8") as f:
                json.dump(state, f, ensure_ascii=False, indent=2)
            self.status.showMessage(f"נשמר ל־{path}", 2000)
        except Exception as e:
            self.status.showMessage(f"שגיאה בשמירה: {e}", 2000)

    def load_state(self):
        path = Path.cwd() / "exercise_state.json"
        if not path.exists():
            return
        try:
            with open(path, "r", encoding="utf-8") as f:
                state = json.load(f)
            # load names
            self.input_name.clear()
            for name in state.get("names", []):
                self.input_name.addItem(name)
            # load rows
            self.table.setRowCount(0)
            for row_data in state.get("rows", []):
                r = self.table.rowCount()
                self.table.insertRow(r)
                for c, val in enumerate(row_data):
                    item = QTableWidgetItem(str(val))
                    # center or left align as before
                    aligns = [Qt.AlignmentFlag.AlignHCenter, Qt.AlignmentFlag.AlignLeft, Qt.AlignmentFlag.AlignHCenter, Qt.AlignmentFlag.AlignHCenter, Qt.AlignmentFlag.AlignHCenter, Qt.AlignmentFlag.AlignHCenter]
                    item.setTextAlignment(aligns[c] | Qt.AlignmentFlag.AlignVCenter)
                    self.table.setItem(r, c, item)
            self.btn_pop.setEnabled(self.table.rowCount() > 0)
            self.status.showMessage(f"טען מצב מ־{path}", 2000)
        except Exception as e:
            self.status.showMessage(f"שגיאה בטעינה: {e}", 2000)

    def _remove_name_if_unused(self, name: str):
        # אם אין מופעים של השם בטבלה, הסר מהקומבו
        for r in range(self.table.rowCount()):
            it = self.table.item(r, 1)
            if it and it.text() == name:
                return
        # לא מצאנו מופעים - נסיר מהקומבו
        for i in range(self.input_name.count() - 1, -1, -1):
            if self.input_name.itemText(i) == name:
                self.input_name.removeItem(i)

    def _show_table_context_menu(self, pos):
        # הצג תפריט שמאפשר למחוק שורות נבחרות
        menu = QMenu()
        act_delete = menu.addAction("מחק שורות נבחרות")
        action = menu.exec(self.table.viewport().mapToGlobal(pos))
        if action == act_delete:
            self.delete_selected_rows()

    def delete_selected_rows(self):
        # מחק את כל השורות שנבחרו, והסר שמות לא בשימוש
        selected = sorted({idx.row() for idx in self.table.selectedIndexes()}, reverse=True)
        for r in selected:
            item = self.table.item(r, 1)
            name = item.text() if item is not None else None
            self.table.removeRow(r)
            if name:
                self._remove_name_if_unused(name)
        self.btn_pop.setEnabled(self.table.rowCount() > 0)

    def _edit_date_cell(self, row: int, column: int):
        # Allow editing only the date column (0)
        if column != 0:
            return
        item = self.table.item(row, column)
        current = item.text() if item is not None else datetime.now().strftime("%Y-%m-%d")
        # ask user for new date (YYYY-MM-DD)
        text, ok = QInputDialog.getText(self, "ערוך תאריך", "תאריך (YYYY-MM-DD):", text=current)
        if not ok:
            return
        new = text.strip()
        # validate
        try:
            _ = datetime.strptime(new, "%Y-%m-%d")
        except Exception:
            QMessageBox.warning(self, "תאריך לא תקין", "פורמט התאריך צריך להיות YYYY-MM-DD")
            return
        # update cell
        new_item = QTableWidgetItem(new)
        new_item.setTextAlignment(Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignVCenter)
        self.table.setItem(row, column, new_item)
        # persist change
        try:
            self.save_state()
        except Exception:
            pass


def main():
    app = QApplication(sys.argv)
    w = MainWindow()
    w.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
