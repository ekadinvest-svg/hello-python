import sys
from typing import Any

from PySide6.QtCore import QSize, Qt
from PySide6.QtGui import QAction, QDoubleValidator, QValidator
from PySide6.QtWidgets import (
    QApplication,
    QFrame,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QPushButton,
    QSizePolicy,
    QStatusBar,
    QToolBar,
    QVBoxLayout,
    QWidget,
)


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("אני מתאמן")
        self.resize(760, 500)

        # נתונים מצטברים
        self.exercise_names: list[str] = []
        self.weights: list[str] = []

        # --- מכולה כללית + RTL + מרווחים עקביים כדי ששתי השורות יהיו באותו רוחב ---
        container = QWidget()
        container.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
        layout = QVBoxLayout(container)
        layout.setContentsMargins(16, 16, 16, 16)   # שוליים אחידים
        layout.setSpacing(10)

        # --- שורה: שם תרגיל ---
        name_label = QLabel("שם תרגיל:")
        self.input_name = QLineEdit()
        self.input_name.setPlaceholderText("לדוגמה: לחיצת חזה")
        self.input_name.setAlignment(Qt.AlignmentFlag.AlignRight)
        self.input_name.textChanged.connect(self._update_add_enabled)

        row_name = QHBoxLayout()
        row_name.setContentsMargins(0, 0, 0, 0)
        row_name.addWidget(name_label)
        row_name.addWidget(self.input_name)

        # --- שורה: משקל (מספרים לא שלמים) ---
        weight_label = QLabel("משקל:")
        self.input_weight = QLineEdit()
        self.input_weight.setPlaceholderText("לדוגמה: 60.5")
        self.input_weight.setAlignment(Qt.AlignmentFlag.AlignRight)

        validator = QDoubleValidator(0.0, 10000.0, 3, self)
        validator.setNotation(QDoubleValidator.Notation.StandardNotation)
        self.input_weight.setValidator(validator)
        self.input_weight.textChanged.connect(self._update_add_enabled)

        row_weight = QHBoxLayout()
        row_weight.setContentsMargins(0, 0, 0, 0)
        row_weight.addWidget(weight_label)
        row_weight.addWidget(self.input_weight)

        # --- כפתורים: הוסף + נקה אחרון ---
        self.btn_add = QPushButton("הוסף")
        self.btn_add.setEnabled(False)
        self.btn_add.clicked.connect(self.add_entry)

        self.btn_pop = QPushButton("נקה אחרון")
        self.btn_pop.setEnabled(False)
        self.btn_pop.clicked.connect(self.pop_last)

        buttons_row = QHBoxLayout()
        buttons_row.setContentsMargins(0, 0, 0, 0)
        buttons_row.setSpacing(8)
        buttons_row.addWidget(self.btn_add)
        buttons_row.addWidget(self.btn_pop)
        buttons_row.addStretch(1)  # דוחף את הכפתורים לקצה ימין בצורה יפה

        # --- אזור תצוגה תחתון: שתי שורות באותו רוחב, יישור לימין, רקע "אפור" ---
        self.output_names = QLabel("")
        self._style_output_label(self.output_names)

        self.output_weights = QLabel("")
        self._style_output_label(self.output_weights)

        # פריסה
        layout.addLayout(row_name)
        layout.addLayout(row_weight)
        layout.addLayout(buttons_row)
        layout.addSpacing(10)
        layout.addWidget(self.output_names)
        layout.addWidget(self.output_weights)

        self.setCentralWidget(container)

        # טולבר + סטטוס
        self._build_toolbar()
        self.status = QStatusBar()
        self.setStatusBar(self.status)
        self.status.showMessage("מוכן להזנה")

        # Enter מוסיף אם תקין
        self.input_name.returnPressed.connect(self._try_add_on_enter)
        self.input_weight.returnPressed.connect(self._try_add_on_enter)

    # ---------- עזרים ותצוגה ----------

    def _style_output_label(self, lbl: QLabel):
        """עיצוב שתי שורות התצוגה: אותו רוחב, יישור לימין, רקע נעים ומסגרת קלה."""
        lbl.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        lbl.setWordWrap(True)
        lbl.setFrameShape(QFrame.StyledPanel)
        lbl.setLineWidth(1)
        lbl.setAutoFillBackground(True)
        lbl.setStyleSheet(
            "QLabel {"
            "  background: #f3f3f3;"
            "  border: 1px solid #d9d9d9;"
            "  border-radius: 6px;"
            "  padding: 8px 10px;"
            "}"
        )
        # מבטיח שהשורה "תתפוס" את כל הרוחב האפשרי (כדי ששתי השורות יתחילו ויסתיימו באותו מקום)
        lbl.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        lbl.setMinimumHeight(36)

    def _build_toolbar(self):
        toolbar = QToolBar("כלים")
        toolbar.setIconSize(QSize(16, 16))
        self.addToolBar(toolbar)

        act_clear = QAction("נקה הכול", self)
        act_clear.triggered.connect(self.clear_all)
        toolbar.addAction(act_clear)

    def _weight_state(self, text: str) -> QValidator.State:
        """מאמת טקסט משקל ומחזיר QValidator.State (בטוח טיפוסית)."""
        v = self.input_weight.validator()
        if isinstance(v, QDoubleValidator):
            res: Any = v.validate(text, 0)
            if isinstance(res, tuple) and len(res) > 0:
                state = res[0]
                if isinstance(state, QValidator.State):
                    return state
        return QValidator.State.Invalid

    # ---------- לוגיקה ----------

    def _update_add_enabled(self):
        name_ok = self.input_name.text().strip() != ""
        weight_text = self.input_weight.text().strip().replace(",", ".")
        state = self._weight_state(weight_text)
        weight_ok = (weight_text != "") and (state == QValidator.State.Acceptable)
        self.btn_add.setEnabled(name_ok and weight_ok)

    def _try_add_on_enter(self):
        if self.btn_add.isEnabled():
            self.add_entry()

    def add_entry(self):
        name = self.input_name.text().strip()
        weight_raw = self.input_weight.text().strip().replace(",", ".")
        if not name:
            self.status.showMessage("מלא שם תרגיל.", 2000)
            return

        if self._weight_state(weight_raw) != QValidator.State.Acceptable:
            self.status.showMessage("משקל חייב להיות מספר (אפשר עשרוני).", 2500)
            return

        # נורמליזציה להצגה יפה
        try:
            weight_value = float(weight_raw)
            weight_str = (
                f"{int(weight_value)}"
                if weight_value.is_integer()
                else f"{weight_value:.3f}".rstrip("0").rstrip(".")
            )
        except ValueError:
            self.status.showMessage("קלט משקל לא חוקי.", 2000)
            return

        # עדכון רשימות ותצוגה
        self.exercise_names.append(name)
        self.weights.append(f"{weight_str} Kg")

        self._refresh_outputs()

        # ניקוי ושחזור פוקוס
        self.input_name.clear()
        self.input_weight.clear()
        self.input_name.setFocus()
        self._update_add_enabled()
        self.status.showMessage(f"התווסף: {name} ({weight_str} Kg)", 2000)
        self.btn_pop.setEnabled(len(self.exercise_names) > 0)

    def pop_last(self):
        if self.exercise_names:
            last_name = self.exercise_names.pop()
            last_weight = self.weights.pop() if self.weights else ""
            self._refresh_outputs()
            self.status.showMessage(f"נמחק האחרון: {last_name} {last_weight}", 2000)
            self.btn_pop.setEnabled(len(self.exercise_names) > 0)

    def _refresh_outputs(self):
        # מציג עם פסיק ורווח, מיושר לימין, שתי שורות באותו רוחב
        self.output_names.setText(", ".join(self.exercise_names))
        self.output_weights.setText(", ".join(self.weights))

    def clear_all(self):
        self.exercise_names.clear()
        self.weights.clear()
        self._refresh_outputs()
        self.input_name.clear()
        self.input_weight.clear()
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
