import sys
from typing import Any
from datetime import datetime

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QLabel, QLineEdit, QPushButton,
    QVBoxLayout, QHBoxLayout, QWidget, QToolBar, QStatusBar,
    QTableWidget, QTableWidgetItem, QSizePolicy
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

        # Form: stack labels above inputs and make all inputs identical size
        form_width = 300
        form_height = 28

        # --- שם תרגיל ---
        name_label = QLabel("שם תרגיל:")
        self.input_name = QLineEdit()
        self.input_name.setPlaceholderText("לדוגמה: לחיצת חזה")
        self.input_name.setAlignment(Qt.AlignmentFlag.AlignLeft)
        self.input_name.setFixedSize(form_width, form_height)
        self.input_name.textChanged.connect(self._update_add_enabled)

        name_layout = QVBoxLayout()
        name_layout.addWidget(name_label)
        name_layout.addWidget(self.input_name)

        # --- משקל ---
        weight_label = QLabel("משקל:")
        self.input_weight = QLineEdit()
        self.input_weight.setPlaceholderText("לדוגמה: 60.5")
        self.input_weight.setAlignment(Qt.AlignmentFlag.AlignLeft)
        self.input_weight.setFixedSize(form_width, form_height)
        dbl_validator = QDoubleValidator(0.0, 10000.0, 3, self)
        dbl_validator.setNotation(QDoubleValidator.Notation.StandardNotation)
        self.input_weight.setValidator(dbl_validator)
        self.input_weight.textChanged.connect(self._update_add_enabled)

        weight_layout = QVBoxLayout()
        weight_layout.addWidget(weight_label)
        weight_layout.addWidget(self.input_weight)

        # --- סטים ---
        sets_label = QLabel("סטים:")
        self.input_sets = QLineEdit()
        self.input_sets.setPlaceholderText("לדוגמה: 4")
        self.input_sets.setAlignment(Qt.AlignmentFlag.AlignLeft)
        self.input_sets.setFixedSize(form_width, form_height)
        self.input_sets.setValidator(QIntValidator(1, 1000, self))
        self.input_sets.textChanged.connect(self._update_add_enabled)

        sets_layout = QVBoxLayout()
        sets_layout.addWidget(sets_label)
        sets_layout.addWidget(self.input_sets)

        # --- חזרות ---
        reps_label = QLabel("חזרות:")
        self.input_reps = QLineEdit()
        self.input_reps.setPlaceholderText("לדוגמה: 12")
        self.input_reps.setAlignment(Qt.AlignmentFlag.AlignLeft)
        self.input_reps.setFixedSize(form_width, form_height)
        self.input_reps.setValidator(QIntValidator(1, 1000, self))
        self.input_reps.textChanged.connect(self._update_add_enabled)

        reps_layout = QVBoxLayout()
        reps_layout.addWidget(reps_label)
        reps_layout.addWidget(self.input_reps)

        # --- חזרות בסט האחרון ---
        last_reps_label = QLabel("סט אחרון:")
        self.input_last_reps = QLineEdit()
        self.input_last_reps.setPlaceholderText("לדוגמה: 15")
        self.input_last_reps.setAlignment(Qt.AlignmentFlag.AlignLeft)
        self.input_last_reps.setFixedSize(form_width, form_height)
        self.input_last_reps.setValidator(QIntValidator(1, 1000, self))
        self.input_last_reps.textChanged.connect(self._update_add_enabled)

        last_reps_layout = QVBoxLayout()
        last_reps_layout.addWidget(last_reps_label)
        last_reps_layout.addWidget(self.input_last_reps)

        # --- כפתורים ---
        self.btn_add = QPushButton("הוסף")
        self.btn_add.setEnabled(False)
        self.btn_add.clicked.connect(self.add_entry)

        self.btn_pop = QPushButton("נקה אחרון")
        self.btn_pop.setEnabled(False)
        self.btn_pop.clicked.connect(self.pop_last)

        buttons_row = QHBoxLayout()
        buttons_row.addWidget(self.btn_add)
        buttons_row.addWidget(self.btn_pop)
        buttons_row.addStretch(1)

        # --- טבלה (תאריך, שם, משקל, סטים, חזרות, סט אחרון) ---
        self.table = EqualWidthTable(0, 6)
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

        # פריסה
        layout.addLayout(name_layout)
        layout.addLayout(weight_layout)
        layout.addLayout(sets_layout)
        layout.addLayout(reps_layout)
        layout.addLayout(last_reps_layout)
        layout.addLayout(buttons_row)
        layout.addWidget(self.table)

        self.setCentralWidget(container)

        # טולבר + סטוס
        self._build_toolbar()
        self.status = QStatusBar()
        self.setStatusBar(self.status)
        self.status.showMessage("מוכן להזנה")

        # Enter מוסיף אם אפשר
        for w in [self.input_name, self.input_weight, self.input_sets, self.input_reps, self.input_last_reps]:
            w.returnPressed.connect(self._try_add_on_enter)

        # ניווט עם ↑/↓ בין השדות
        self._inputs = [self.input_name, self.input_weight, self.input_sets, self.input_reps, self.input_last_reps]
        for w in self._inputs:
            w.installEventFilter(self)

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
        name_ok = self.input_name.text().strip() != ""
        weight_ok = self._weight_state(self.input_weight.text().strip().replace(",", ".")) == QValidator.State.Acceptable
        sets_ok = self._int_state(self.input_sets) == QValidator.State.Acceptable
        reps_ok = self._int_state(self.input_reps) == QValidator.State.Acceptable
        last_reps_ok = self._int_state(self.input_last_reps) == QValidator.State.Acceptable
        self.btn_add.setEnabled(name_ok and weight_ok and sets_ok and reps_ok and last_reps_ok)

    def _try_add_on_enter(self):
        if self.btn_add.isEnabled():
            self.add_entry()

    def add_entry(self):
        name = self.input_name.text().strip()
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

        # ניקוי שדות
        for field in [self.input_name, self.input_weight, self.input_sets, self.input_reps, self.input_last_reps]:
            field.clear()
        self.input_name.setFocus()
        self._update_add_enabled()
        self.btn_pop.setEnabled(True)
        self.status.showMessage(f"התווסף: {name} ({weight_str} Kg, {sets_val}x{reps_val})", 2000)

    def pop_last(self):
        rows = self.table.rowCount()
        if rows > 0:
            self.table.removeRow(rows - 1)
            self.btn_pop.setEnabled(self.table.rowCount() > 0)
            self.status.showMessage("נמחק האחרון.", 2000)

    def clear_all(self):
        self.table.setRowCount(0)
        for field in [self.input_name, self.input_weight, self.input_sets, self.input_reps, self.input_last_reps]:
            field.clear()
        self._update_add_enabled()
        self.btn_pop.setEnabled(False)
        self.status.showMessage("נוקה הכול", 2000)


def main():
    app = QApplication(sys.argv)
    w = MainWindow()
    w.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
