import json
import os
import sys
from datetime import datetime
from pathlib import Path
from typing import Any

from openpyxl import Workbook
from openpyxl.chart import LineChart, Reference
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

os.environ.setdefault('QT_API', 'pyside6')
from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure

os.environ["QT_API"] = "pyside6"
import matplotlib.dates as mdates
from PySide6.QtCore import QDate, QEvent, QSize, Qt
from PySide6.QtGui import (
    QAction,
    QDoubleValidator,
    QIntValidator,
    QKeySequence,
    QShortcut,
    QValidator,
)
from PySide6.QtWidgets import (
    QApplication,
    QCalendarWidget,
    QDialog,
    QDialogButtonBox,
    QGridLayout,
    QHBoxLayout,
    QInputDialog,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMenu,
    QMessageBox,
    QPushButton,
    QSizePolicy,
    QStatusBar,
    QTableWidget,
    QTableWidgetItem,
    QTabWidget,
    QToolBar,
    QVBoxLayout,
    QWidget,
)


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
        width = self.viewport().width()
        col_width = width // cols
        for c in range(cols):
            self.setColumnWidth(c, col_width)


class ExerciseTab(QWidget):
    def __init__(self, exercise_name: str):
        super().__init__()
        self.exercise_name = exercise_name
        self.setContentsMargins(5, 5, 5, 5)
        self._has_unsaved_changes = False
        # מערכת Undo/Redo
        self._undo_stack = []  # מחסנית של מצבי טבלה קודמים
        self._redo_stack = []  # מחסנית של מצבים לשחזור
        self._max_undo = 5  # מקסימום 5 פעולות
        self._is_restoring = False  # דגל למניעת שמירה בזמן שחזור
        self._init_ui()
        try:
            self.load_state()
        except Exception:
            pass
        # אחרי טעינת המצב, נאפס את דגל השינויים
        self._has_unsaved_changes = False
        # שמירת מצב ראשוני
        self._save_state_to_undo()

    def _init_ui(self):
        layout = QVBoxLayout()

        # טופס הכנסת נתונים
        form = QGridLayout()
        form.setContentsMargins(0, 0, 0, 0)

        # הגדרת שורות טופס ההכנסה
        # תאריך ומשקל
        self.input_weight = QLineEdit()
        self.input_weight.setPlaceholderText("משקל")
        self.input_weight.setValidator(QDoubleValidator(0, 1000, 3))

        # סטים וחזרות
        self.input_sets = QLineEdit()
        self.input_sets.setPlaceholderText("סטים")
        self.input_sets.setValidator(QIntValidator(0, 1000))

        self.input_reps = QLineEdit()
        self.input_reps.setPlaceholderText("חזרות")
        self.input_reps.setValidator(QIntValidator(0, 1000))

        self.input_last_reps = QLineEdit()
        self.input_last_reps.setPlaceholderText("סט אחרון")
        self.input_last_reps.setValidator(QIntValidator(0, 1000))

        # כפתורים: הוסף ומחק
        self.btn_add = QPushButton("הוסף")
        self.btn_pop = QPushButton("מחק אחרון")
        self.btn_delete_row = QPushButton("מחק שורה")
        self.btn_duplicate_row = QPushButton("שכפל שורה")
        self.btn_plot = QPushButton("הצג גרף")
        self.btn_back = QPushButton("חזור לטבלה")
        self.btn_back.hide()
        
        # סגנון מיוחד לכפתורים
        self.btn_plot.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;
            }
            QPushButton:hover {
                background-color: #388E3C;
            }
        """)
        
        delete_buttons_style = """
            QPushButton {
                background-color: #f44336;
            }
            QPushButton:hover {
                background-color: #d32f2f;
            }
            QPushButton:disabled {
                background-color: #ffcdd2;
            }
        """
        self.btn_pop.setStyleSheet(delete_buttons_style)
        self.btn_delete_row.setStyleSheet(delete_buttons_style)
        
        # עיצוב כפתור שכפול
        duplicate_button_style = """
            QPushButton {
                background-color: #FF9800;
            }
            QPushButton:hover {
                background-color: #F57C00;
            }
            QPushButton:disabled {
                background-color: #FFE0B2;
            }
        """
        self.btn_duplicate_row.setStyleSheet(duplicate_button_style)
        
        # התחלתי מצב כפתורים - מבוטלים
        self.btn_pop.setEnabled(False)
        self.btn_delete_row.setEnabled(False)
        self.btn_duplicate_row.setEnabled(False)

        self.btn_add.setEnabled(False)
        self.btn_pop.setEnabled(False)
        self.btn_delete_row.setEnabled(False)
        self.btn_duplicate_row.setEnabled(False)

        # יצירת תצוגת סיכום
        summary_layout = QHBoxLayout()
        summary_layout.setSpacing(15)
        
        # עיצוב תוויות הסיכום בקופסאות
        # קופסה כחולה לאימונים
        exercises_style = """
            QLabel {
                font-size: 16pt;
                font-weight: bold;
                color: white;
                padding: 15px 25px;
                border-radius: 8px;
                background: qlineargradient(x1:0, y1:0, x2:1, y2:1,
                    stop:0 #2196F3, stop:1 #1976D2);
                border: 2px solid #1565C0;
            }
        """
        
        # קופסה ירוקה למשקל שהרמתי
        weight_style = """
            QLabel {
                font-size: 16pt;
                font-weight: bold;
                color: white;
                padding: 15px 25px;
                border-radius: 8px;
                background: qlineargradient(x1:0, y1:0, x2:1, y2:1,
                    stop:0 #4CAF50, stop:1 #388E3C);
                border: 2px solid #2E7D32;
            }
        """
        
        # קופסה כתומה לממוצע
        avg_style = """
            QLabel {
                font-size: 16pt;
                font-weight: bold;
                color: white;
                padding: 15px 25px;
                border-radius: 8px;
                background: qlineargradient(x1:0, y1:0, x2:1, y2:1,
                    stop:0 #FF9800, stop:1 #F57C00);
                border: 2px solid #E65100;
            }
        """
        
        self.total_exercises_label = QLabel('<div style="text-align: center;">אימונים<br><span style="font-size: 24pt;">0</span><br><span style="font-size: 32pt;">💪</span></div>')
        self.total_exercises_label.setStyleSheet(exercises_style)
        self.total_exercises_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.total_exercises_label.setMinimumWidth(300)
        self.total_exercises_label.setMaximumWidth(300)
        self.total_exercises_label.setTextFormat(Qt.TextFormat.RichText)
        
        self.total_weight_label = QLabel('<div style="text-align: center;">משקל שהרמתי<br><span style="font-size: 24pt;">0 ק"ג</span><br><span style="font-size: 32pt;">🏋️</span></div>')
        self.total_weight_label.setStyleSheet(weight_style)
        self.total_weight_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.total_weight_label.setMinimumWidth(300)
        self.total_weight_label.setMaximumWidth(300)
        self.total_weight_label.setTextFormat(Qt.TextFormat.RichText)
        
        self.avg_weight_label = QLabel('<div style="text-align: center;">משקל לסט<br><span style="font-size: 24pt;">0 ק"ג</span><br><span style="font-size: 32pt;">📊</span></div>')
        self.avg_weight_label.setStyleSheet(avg_style)
        self.avg_weight_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.avg_weight_label.setMinimumWidth(300)
        self.avg_weight_label.setMaximumWidth(300)
        self.avg_weight_label.setTextFormat(Qt.TextFormat.RichText)
        
        summary_layout.addWidget(self.total_exercises_label)
        summary_layout.addWidget(self.total_weight_label)
        summary_layout.addWidget(self.avg_weight_label)
        
        # הוספת שדות לטופס ללא תוויות
        input_layout = QVBoxLayout()
        fields = [
            self.input_weight,
            self.input_sets,
            self.input_reps,
            self.input_last_reps,
        ]
        
         # הגדרת רוחב מקסימלי לשדות הקלט
        for field in fields:
            field.setMaximumWidth(150)
            input_layout.addWidget(field)
        
        # סידור השדות והסיכום בשורה אחת
        inputs_and_summary = QHBoxLayout()
        inputs_and_summary.addLayout(summary_layout)
        inputs_and_summary.addStretch()
        inputs_and_summary.addLayout(input_layout)
        
        form.addLayout(inputs_and_summary, 0, 0)

        # טבלת נתונים עם עמודות שוות רוחב
        self._inputs = [
            self.input_weight,
            self.input_sets,
            self.input_reps,
            self.input_last_reps,
        ]
        for inp in self._inputs:
            inp.textChanged.connect(self._update_add_enabled)
            inp.returnPressed.connect(self._try_add_on_enter)
            inp.installEventFilter(self)

        # טבלת נתונים עם עמודות שוות רוחב
        self.table = EqualWidthTable()
        self.table.setColumnCount(5)
        self.table.setHorizontalHeaderLabels(["סט אחרון", "חזרות", "סטים", "משקל", "תאריך"])
        self.table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)  # ביטול עריכה ישירה
        self.table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self._show_table_context_menu)
        self.table.cellDoubleClicked.connect(self._edit_date_cell)  # חיבור לאירוע לחיצה כפולה

        # אירועי כפתורים
        self.btn_add.clicked.connect(self.add_entry)
        self.btn_pop.clicked.connect(self.pop_last)
        self.btn_delete_row.clicked.connect(self.delete_selected_row)
        self.btn_duplicate_row.clicked.connect(self.duplicate_selected_row)
        self.btn_plot.clicked.connect(self.plot_selected_exercise)
        self.btn_back.clicked.connect(self.restore_normal_view)
        
        # חיבור לאירוע בחירת שורה בטבלה
        self.table.itemSelectionChanged.connect(self._update_delete_button)
        
        # קיצורי מקלדת למחיקה ושכפול שורה
        delete_shortcut = QShortcut(QKeySequence("Ctrl+E"), self)
        delete_shortcut.activated.connect(self.delete_selected_row)
        
        delete_shortcut_he = QShortcut(QKeySequence("Ctrl+ק"), self)
        delete_shortcut_he.activated.connect(self.delete_selected_row)
        
        duplicate_shortcut = QShortcut(QKeySequence("Ctrl+D"), self)
        duplicate_shortcut.activated.connect(self.duplicate_selected_row)
        
        duplicate_shortcut_he = QShortcut(QKeySequence("Ctrl+ג"), self)
        duplicate_shortcut_he.activated.connect(self.duplicate_selected_row)

        # מסגרת גרף
        self.figure = Figure(figsize=(6, 4))
        self.canvas = FigureCanvas(self.figure)

        # הוספת רכיבים לממשק
        bottom_buttons = QHBoxLayout()
        bottom_buttons.addWidget(self.btn_add)
        bottom_buttons.addWidget(self.btn_pop)
        bottom_buttons.addWidget(self.btn_delete_row)
        bottom_buttons.addWidget(self.btn_duplicate_row)
        bottom_buttons.addWidget(self.btn_plot)
        bottom_buttons.addWidget(self.btn_back)

        self.input_container = QWidget()
        self.input_container.setLayout(form)
        self.input_container.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Maximum)

        layout.addWidget(self.input_container)
        layout.addLayout(bottom_buttons)
        layout.addWidget(self.table)
        layout.addWidget(self.canvas)

        self.setLayout(layout)

    def _update_add_enabled(self):
        weight_ok = self._weight_state(self.input_weight.text().strip().replace(",", ".")) == QValidator.State.Acceptable
        sets_ok = self._int_state(self.input_sets) == QValidator.State.Acceptable
        reps_ok = self._int_state(self.input_reps) == QValidator.State.Acceptable
        last_reps_ok = self._int_state(self.input_last_reps) == QValidator.State.Acceptable
        self.btn_add.setEnabled(weight_ok and sets_ok and reps_ok and last_reps_ok)

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

    def eventFilter(self, obj, event):
        if event.type() == QEvent.Type.KeyPress and obj in self._inputs:
            key = event.key()
            idx = self._inputs.index(obj)
            if key == Qt.Key.Key_Down:
                self._inputs[(idx + 1) % len(self._inputs)].setFocus()
                return True
            if key == Qt.Key.Key_Up:
                self._inputs[(idx - 1) % len(self._inputs)].setFocus()
                return True
        return super().eventFilter(obj, event)

    def _try_add_on_enter(self):
        if self.btn_add.isEnabled():
            self.add_entry()

    def _calculate_total_weight(self):
        """חישוב סך המשקל המצטבר מכל האימונים"""
        total = 0
        for row in range(self.table.rowCount()):
            try:
                # קבלת הערכים מהטבלה
                try:
                    weight_item = self.table.item(row, 3)
                    sets_item = self.table.item(row, 2)
                    reps_item = self.table.item(row, 1)
                    last_reps_item = self.table.item(row, 0)
                    
                    # בדיקה מקיפה של תקינות הנתונים
                    if not all([
                        isinstance(item, QTableWidgetItem) and item.text()
                        for item in [weight_item, sets_item, reps_item, last_reps_item]
                    ]):
                        continue

                    # המרת הערכים למספרים
                    weight_text = weight_item.text().split()[0].replace(",", ".")  # type: ignore
                    weight = float(weight_text)
                    sets = int(sets_item.text())  # type: ignore
                    reps = int(reps_item.text())  # type: ignore
                    last_reps = int(last_reps_item.text())  # type: ignore
                    
                    # חישוב: (סטים-1 * חזרות * משקל) + (סט אחרון * משקל)
                    total += ((sets - 1) * reps * weight) + (last_reps * weight)
                except (ValueError, AttributeError, IndexError):
                    continue
            except (ValueError, AttributeError, IndexError):
                continue
        return total

    def _update_summary(self):
        """עדכון תוויות הסיכום"""
        # עדכון מספר האימונים
        exercises_count = self.table.rowCount()
        self.total_exercises_label.setText(f'<div style="text-align: center;">אימונים<br><span style="font-size: 24pt;">{exercises_count}</span><br><span style="font-size: 32pt;">💪</span></div>')
        
        # עדכון סך המשקל
        total_weight = self._calculate_total_weight()
        self.total_weight_label.setText(f'<div style="text-align: center;">משקל שהרמתי<br><span style="font-size: 24pt;">{total_weight:,.0f} ק"ג</span><br><span style="font-size: 32pt;">🏋️</span></div>')
        
        # עדכון משקל ממוצע לסט
        if exercises_count > 0:
            avg_weight = total_weight / exercises_count
            self.avg_weight_label.setText(f'<div style="text-align: center;">משקל לסט<br><span style="font-size: 24pt;">{avg_weight:,.0f} ק"ג</span><br><span style="font-size: 32pt;">📊</span></div>')
        else:
            self.avg_weight_label.setText('<div style="text-align: center;">משקל לסט<br><span style="font-size: 24pt;">0 ק"ג</span><br><span style="font-size: 32pt;">📊</span></div>')

    def add_entry(self):
        weight_raw = self.input_weight.text().strip().replace(",", ".")
        sets_raw = self.input_sets.text().strip()
        reps_raw = self.input_reps.text().strip()
        last_reps_raw = self.input_last_reps.text().strip()

        if not (weight_raw and sets_raw and reps_raw and last_reps_raw):
            window = self.window()
            if isinstance(window, QMainWindow) and window.statusBar():
                window.statusBar().showMessage("מלא את כל השדות.", 2000)
            return
            
        self._has_unsaved_changes = True

        try:
            weight_val = float(weight_raw)
            weight_str = f"{int(weight_val)}" if weight_val.is_integer() else f"{weight_val:.3f}".rstrip("0").rstrip(".")
            sets_val = int(sets_raw)
            reps_val = int(reps_raw)
            last_reps_val = int(last_reps_raw)
        except ValueError:
            window = self.window()
            if isinstance(window, QMainWindow) and window.statusBar():
                window.statusBar().showMessage("קלט לא תקין.", 2000)
            return

        # תאריך
        date_str = datetime.now().strftime("%d/%m/%Y")

        # שמירה למחסנית Undo לפני השינוי
        self._save_state_to_undo()

        # הוספה לטבלה
        row = self.table.rowCount()
        self.table.insertRow(row)

        data = [last_reps_val, reps_val, sets_val, f"{weight_str} Kg", date_str]
        aligns = [Qt.AlignmentFlag.AlignHCenter] * 5  # כל העמודות ממורכזות
        
        for col, value in enumerate(data):
            item = QTableWidgetItem(str(value))
            item.setTextAlignment(aligns[col] | Qt.AlignmentFlag.AlignVCenter)
            self.table.setItem(row, col, item)
        
        # עדכון הסיכום
        self._update_summary()

        # ניקוי שדות
        for field in [self.input_weight, self.input_sets, self.input_reps, self.input_last_reps]:
            field.clear()
        self.input_weight.setFocus()
        self._update_add_enabled()
        self.btn_pop.setEnabled(True)
        window = self.window()
        if isinstance(window, QMainWindow) and window.statusBar():
            window.statusBar().showMessage(f"התווסף: {weight_str} Kg, {sets_val}x{reps_val}", 2000)

    def pop_last(self):
        rows = self.table.rowCount()
        if rows > 0:
            # שמירה למחסנית Undo לפני השינוי
            self._save_state_to_undo()
            self.table.removeRow(rows - 1)
            self.btn_pop.setEnabled(self.table.rowCount() > 0)
            self._has_unsaved_changes = True
            self._update_summary()  # עדכון הסיכום
            window = self.window()
            if isinstance(window, QMainWindow) and window.statusBar():
                window.statusBar().showMessage("נמחק האחרון.", 2000)

    def plot_selected_exercise(self):
        # הסתר את האזורים שלא נחוצים בתצוגת גרף
        self.input_container.hide()
        self.table.hide()
        self.btn_add.hide()
        self.btn_pop.hide()
        self.btn_delete_row.hide()
        self.btn_duplicate_row.hide()
        self.btn_plot.hide()
        self.btn_back.show()

        # אסוף את כל הנתונים מהטבלה
        points: list[tuple[datetime, float]] = []
        for r in range(self.table.rowCount()):
            date_item = self.table.item(r, 4)  # תאריך עכשיו בעמודה האחרונה
            weight_item = self.table.item(r, 3)  # משקל עכשיו בעמודה הרביעית
            try:
                wtxt = weight_item.text().split()[0] if weight_item is not None else "0"
                wval = float(wtxt.replace(",", "."))
            except Exception:
                wval = 0.0
            try:
                dstr = date_item.text().strip() if date_item is not None else ""
                dval = datetime.strptime(dstr, "%d/%m/%Y")
            except Exception:
                dval = datetime.now()
            points.append((dval, wval))

        if not points:
            window = self.window()
            if isinstance(window, QMainWindow) and window.statusBar():
                window.statusBar().showMessage("אין רשומות להצגה", 2000)
            return

        # מיין לפי תאריך
        points.sort(key=lambda x: x[0])
        xs = [p[0] for p in points]
        ys = [p[1] for p in points]

        # צייר גרף קווי עם ציר תאריכים
        self.figure.clear()
        # הגדר סגנון גרף
        self.figure.patch.set_facecolor('#ffffff')
        ax = self.figure.add_subplot(111)
        ax.set_facecolor('#f8f9fa')
        
        dates = mdates.date2num(xs)
        ax.plot(dates, ys, '-o', color='#2196F3', linewidth=2, markersize=8, 
                markerfacecolor='white', markeredgecolor='#2196F3', markeredgewidth=2)
        
        ax.xaxis.set_major_locator(mdates.AutoDateLocator())
        ax.xaxis.set_major_formatter(mdates.DateFormatter('%d/%m/%Y'))
        self.figure.autofmt_xdate(rotation=30)
        
        # שימוש בסימן LRM (Left-to-Right Mark) לסידור הטקסט
        LRM = '\u200E'
        title = f"גרף משקלים - {self.exercise_name}"
        ax.set_title(f"{LRM}{title[::-1]}", fontsize=12, pad=15)  # הופך את סדר האותיות
        
        # הוספת kg למספרים על ציר Y
        from matplotlib.ticker import FuncFormatter
        def kg_formatter(x, pos):
            return f'{int(x)} kg'
        ax.yaxis.set_major_formatter(FuncFormatter(kg_formatter))
        
        # הגדרת רשת עדינה
        ax.grid(True, linestyle='--', alpha=0.3)
        ax.set_axisbelow(True)  # הרשת מאחורי הנתונים
        
        # עיצוב שולי הגרף
        for spine in ax.spines.values():
            spine.set_color('#cccccc')
            
        # התאמת צבע וגודל תוויות הצירים
        ax.tick_params(axis='both', colors='#666666', labelsize=9)
        
        self.canvas.draw()
        self.canvas.show()

    def save_state(self):
        state = {
            "rows": []
        }
        for r in range(self.table.rowCount()):
            row_data = []
            for c in range(self.table.columnCount()):
                item = self.table.item(r, c)
                row_data.append(item.text() if item is not None else "")
            state["rows"].append(row_data)

        path = Path.cwd() / f"exercise_state_{self.exercise_name}.json"
        try:
            with open(path, "w", encoding="utf-8") as f:
                json.dump(state, f, ensure_ascii=False, indent=2)
            self._has_unsaved_changes = False  # מאפס את דגל השינויים אחרי שמירה
            window = self.window()
            if isinstance(window, QMainWindow) and window.statusBar():
                window.statusBar().showMessage(f"נשמר ל־{path}", 2000)
        except Exception as e:
            window = self.window()
            if isinstance(window, QMainWindow) and window.statusBar():
                window.statusBar().showMessage(f"שגיאה בשמירה: {e}", 2000)

    def load_state(self):
        path = Path.cwd() / f"exercise_state_{self.exercise_name}.json"
        if not path.exists():
            return
        try:
            with open(path, "r", encoding="utf-8") as f:
                state = json.load(f)
            self.table.setRowCount(0)
            for row_data in state.get("rows", []):
                r = self.table.rowCount()
                self.table.insertRow(r)
                for c, val in enumerate(row_data):
                    item = QTableWidgetItem(str(val))
                    item.setTextAlignment(Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignVCenter)
                    self.table.setItem(r, c, item)
            self.btn_pop.setEnabled(self.table.rowCount() > 0)
            # עדכון הסיכום (הקופסאות) אחרי טעינת הנתונים
            self._update_summary()
            window = self.window()
            if isinstance(window, QMainWindow) and window.statusBar():
                window.statusBar().showMessage(f"טען מצב מ־{path}", 2000)
        except Exception as e:
            window = self.window()
            if isinstance(window, QMainWindow) and window.statusBar():
                window.statusBar().showMessage(f"שגיאה בטעינה: {e}", 2000)

    def _show_table_context_menu(self, pos):
        menu = QMenu()
        act_delete = menu.addAction("מחק שורות נבחרות")
        action = menu.exec(self.table.viewport().mapToGlobal(pos))
        if action == act_delete:
            self.delete_selected_rows()

    def delete_selected_rows(self):
        selected = sorted({idx.row() for idx in self.table.selectedIndexes()}, reverse=True)
        if selected:  # רק אם יש שורות נבחרות
            # שמירה למחסנית Undo לפני השינוי
            self._save_state_to_undo()
            self._has_unsaved_changes = True
            for r in selected:
                self.table.removeRow(r)
            self.btn_pop.setEnabled(self.table.rowCount() > 0)
            self._update_summary()

    def restore_normal_view(self):
        """החזרת התצוגה למצב רגיל"""
        self.input_container.show()
        self.table.show()
        self.btn_add.show()
        self.btn_pop.show()
        self.btn_delete_row.show()
        self.btn_duplicate_row.show()
        self.btn_plot.show()
        self.btn_back.hide()
        self.canvas.hide()

    def _update_delete_button(self):
        """עדכון מצב כפתור מחיקת שורה בהתאם לבחירה"""
        selected_rows = len({idx.row() for idx in self.table.selectedIndexes()})
        self.btn_delete_row.setEnabled(selected_rows == 1)
        self.btn_duplicate_row.setEnabled(selected_rows == 1)
    
    def delete_selected_row(self):
        """מחיקת השורה הנבחרת"""
        selected_rows = {idx.row() for idx in self.table.selectedIndexes()}
        if len(selected_rows) == 1:
            # שמירה למחסנית Undo לפני השינוי
            self._save_state_to_undo()
            row = selected_rows.pop()
            self.table.removeRow(row)
            self._has_unsaved_changes = True
            self.btn_pop.setEnabled(self.table.rowCount() > 0)
            self._update_summary()
            window = self.window()
            if isinstance(window, QMainWindow) and window.statusBar():
                window.statusBar().showMessage("השורה נמחקה.", 2000)
    
    def duplicate_selected_row(self):
        """שכפול השורה הנבחרת"""
        selected_rows = {idx.row() for idx in self.table.selectedIndexes()}
        if len(selected_rows) == 1:
            row = selected_rows.pop()
            
            # שמירה למחסנית Undo לפני השינוי
            self._save_state_to_undo()
            
            # שכפול הנתונים מהשורה הנבחרת
            row_data = []
            for col in range(self.table.columnCount()):
                item = self.table.item(row, col)
                if item:
                    row_data.append(item.text())
                else:
                    row_data.append("")
            
            # הוספת שורה חדשה עם הנתונים המשוכפלים
            new_row = self.table.rowCount()
            self.table.insertRow(new_row)
            
            for col, value in enumerate(row_data):
                new_item = QTableWidgetItem(value)
                new_item.setTextAlignment(Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignVCenter)
                self.table.setItem(new_row, col, new_item)
            
            self._has_unsaved_changes = True
            self.btn_pop.setEnabled(True)
            self._update_summary()
            window = self.window()
            if isinstance(window, QMainWindow) and window.statusBar():
                window.statusBar().showMessage("השורה שוכפלה.", 2000)
    
    def _save_state_to_undo(self):
        """שמירת המצב הנוכחי למחסנית ה-Undo לפני ביצוע פעולה"""
        # אם אנחנו בתהליך שחזור, לא נשמור
        if self._is_restoring:
            return
        
        # שומר את המצב הנוכחי לפני השינוי
        state = self._get_current_table_state()
        # אם זה המצב הראשון, או שהמצב שונה מהמצב האחרון במחסנית
        if not self._undo_stack or state != self._undo_stack[-1]:
            self._undo_stack.append(state)
            # שמירה של מקסימום 5+1 מצבים (כולל המצב הנוכחי)
            if len(self._undo_stack) > self._max_undo + 1:
                self._undo_stack.pop(0)
        # כאשר נעשית פעולה חדשה, מנקים את מחסנית ה-Redo
        self._redo_stack.clear()
    
    def _get_current_table_state(self):
        """קבלת המצב הנוכחי של הטבלה"""
        state = []
        for r in range(self.table.rowCount()):
            row_data = []
            for c in range(self.table.columnCount()):
                item = self.table.item(r, c)
                row_data.append(item.text() if item is not None else "")
            state.append(row_data)
        return state
    
    def _restore_table_state(self, state):
        """שחזור מצב הטבלה"""
        self._is_restoring = True  # מסמן שאנחנו בתהליך שחזור
        try:
            self.table.setRowCount(0)
            for row_data in state:
                r = self.table.rowCount()
                self.table.insertRow(r)
                for c, val in enumerate(row_data):
                    item = QTableWidgetItem(str(val))
                    item.setTextAlignment(Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignVCenter)
                    self.table.setItem(r, c, item)
            self.btn_pop.setEnabled(self.table.rowCount() > 0)
            self._update_summary()
        finally:
            self._is_restoring = False  # מסיים את תהליך השחזור
    
    def undo(self):
        """ביטול הפעולה האחרונה"""
        if len(self._undo_stack) < 1:
            window = self.window()
            if isinstance(window, QMainWindow) and window.statusBar():
                window.statusBar().showMessage("אין מה לבטל", 2000)
            return
        
        # שמירת המצב הנוכחי ל-Redo (רק אם עדיין לא שמרנו אותו)
        current_state = self._get_current_table_state()
        if not self._redo_stack or current_state != self._redo_stack[-1]:
            self._redo_stack.append(current_state)
            if len(self._redo_stack) > self._max_undo:
                self._redo_stack.pop(0)
        
        # שחזור המצב הקודם
        previous_state = self._undo_stack.pop()
        self._restore_table_state(previous_state)
        self._has_unsaved_changes = True
        
        window = self.window()
        if isinstance(window, QMainWindow) and window.statusBar():
            window.statusBar().showMessage("בוטל", 1000)
    
    def redo(self):
        """שחזור הפעולה שבוטלה"""
        if not self._redo_stack:
            window = self.window()
            if isinstance(window, QMainWindow) and window.statusBar():
                window.statusBar().showMessage("אין מה לשחזר", 2000)
            return
        
        # שמירת המצב הנוכחי ל-Undo
        current_state = self._get_current_table_state()
        if not self._undo_stack or current_state != self._undo_stack[-1]:
            self._undo_stack.append(current_state)
            if len(self._undo_stack) > self._max_undo + 1:
                self._undo_stack.pop(0)
        
        # שחזור המצב מ-Redo
        state = self._redo_stack.pop()
        self._restore_table_state(state)
        self._has_unsaved_changes = True
        
        window = self.window()
        if isinstance(window, QMainWindow) and window.statusBar():
            window.statusBar().showMessage("שוחזר", 1000)
        
    def _edit_date_cell(self, row: int, column: int):
        if column != 4:  # עמודת תאריך היא 4
            self.table.clearSelection()
            return
        item = self.table.item(row, column)
        if item is None:
            return
        
        # קריאת התאריך הנוכחי
        current = item.text() if item is not None else datetime.now().strftime("%d/%m/%Y")
        try:
            current_date = datetime.strptime(current, "%d/%m/%Y")
        except Exception:
            current_date = datetime.now()
        
        # יצירת דיאלוג עם לוח שנה
        dialog = QDialog(self)
        dialog.setWindowTitle("בחר תאריך")
        dialog.setModal(True)
        
        layout = QVBoxLayout()
        
        # יצירת לוח שנה
        calendar = QCalendarWidget()
        calendar.setGridVisible(True)
        
        # הגבלה: לא ניתן לבחור תאריך עתידי
        today = QDate.currentDate()
        calendar.setMaximumDate(today)
        
        calendar.setSelectedDate(QDate(current_date.year, current_date.month, current_date.day))
        
        # תווית להצגת התאריך הנבחר
        date_label = QLabel()
        date_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        date_label.setStyleSheet("font-size: 12pt; padding: 10px; background-color: #E3F2FD; border-radius: 4px;")
        
        def update_label():
            selected = calendar.selectedDate()
            date_label.setText(f"תאריך נבחר: {selected.toString('dd/MM/yyyy')}")
        
        update_label()
        calendar.selectionChanged.connect(update_label)
        
        # כפתורי אישור וביטול
        button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        button_box.accepted.connect(dialog.accept)
        button_box.rejected.connect(dialog.reject)
        
        layout.addWidget(calendar)
        layout.addWidget(date_label)
        layout.addWidget(button_box)
        dialog.setLayout(layout)
        
        # הצגת הדיאלוג
        if dialog.exec() == QDialog.DialogCode.Accepted:
            # שמירה למחסנית Undo לפני השינוי
            self._save_state_to_undo()
            
            selected = calendar.selectedDate()
            new_date = selected.toString("dd/MM/yyyy")
            item.setText(new_date)
            item.setTextAlignment(Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignVCenter)
            self._has_unsaved_changes = True
            # עדכון רוחב העמודה כדי שיתאים לתוכן
            self.table._equalize_columns()
            # נקה בחירה ופוקוס
            self.table.clearSelection()
            self.table.clearFocus()
            self.table.setCurrentCell(-1, -1)
            try:
                self.save_state()
            except Exception:
                pass
        else:
            self.table.clearSelection()
            self.table.clearFocus()


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        # הגדרות חלון ראשי
        self.setWindowTitle("מעקב משקלים")
        self.setMinimumSize(QSize(800, 600))
        self.showMaximized()  # פתיחה במסך מלא

        # יצירת סטטוס בר
        self.setStatusBar(QStatusBar())

        # יצירת מיכל מרכזי
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        layout = QVBoxLayout()
        main_widget.setLayout(layout)

        # הגדרת טאבים
        self.tab_widget = QTabWidget()
        layout.addWidget(self.tab_widget)

        # יצירת סרגל כלים
        toolbar = QToolBar()
        self.addToolBar(toolbar)
        
        # כפתור שמירה בסרגל כלים עם קיצור מקלדת
        save_action = QAction("שמור", self)
        save_action.setShortcuts([QKeySequence("Ctrl+S"), QKeySequence("Ctrl+ד")])  # תמיכה באנגלית ועברית
        save_action.triggered.connect(self._save_current_tab)
        toolbar.addAction(save_action)

        # תפריט קובץ
        file_menu = self.menuBar().addMenu("קובץ")
        
        # פעולת Undo
        undo_action = QAction("אחורה", self)
        undo_action.setShortcut(QKeySequence("Ctrl+Z"))
        undo_action.triggered.connect(self._undo_current_tab)
        file_menu.addAction(undo_action)
        self.addAction(undo_action)  # הוספה לחלון עצמו כדי שקיצור המקלדת יעבוד
        
        # קיצור נוסף בעברית ל-Undo
        undo_shortcut_he = QShortcut(QKeySequence("Ctrl+ז"), self)
        undo_shortcut_he.activated.connect(self._undo_current_tab)
        
        # פעולת Redo
        redo_action = QAction("קדימה", self)
        redo_action.setShortcut(QKeySequence("Ctrl+Y"))
        redo_action.triggered.connect(self._redo_current_tab)
        file_menu.addAction(redo_action)
        self.addAction(redo_action)  # הוספה לחלון עצמו כדי שקיצור המקלדת יעבוד
        
        # קיצורים נוספים בעברית ל-Redo
        redo_shortcut_he = QShortcut(QKeySequence("Ctrl+ט"), self)
        redo_shortcut_he.activated.connect(self._redo_current_tab)
        
        file_menu.addSeparator()
        
        # פעולת שמירה בתפריט (משתמש באותו Action כמו הסרגל)
        file_menu.addAction(save_action)
        
        # פעולת שחזור
        restore_action = QAction("שחזר", self)
        restore_action.setShortcuts([QKeySequence("Ctrl+R"), QKeySequence("Ctrl+ר")])  # תמיכה באנגלית ועברית
        restore_action.triggered.connect(self._restore_current_tab)
        file_menu.addAction(restore_action)
        
        # פעולת ייצוא לאקסל
        export_action = QAction("ייצא לאקסל", self)
        export_action.setShortcuts([QKeySequence("Ctrl+E"), QKeySequence("Ctrl+ק")])  # תמיכה באנגלית ועברית
        export_action.triggered.connect(self._export_to_excel)
        file_menu.addAction(export_action)
        
        file_menu.addSeparator()
        
        # פעולת עזרה
        help_action = QAction("עזרה", self)
        help_action.triggered.connect(self._show_help)
        file_menu.addAction(help_action)

        # תפריט עריכה
        edit_menu = self.menuBar().addMenu("עריכה")
        
        # פעולת הוספת תרגיל
        add_exercise_action = QAction("הוסף תרגיל", self)
        add_exercise_action.setShortcuts([QKeySequence("Ctrl+N"), QKeySequence("Ctrl+מ")])  # תמיכה באנגלית ועברית
        add_exercise_action.triggered.connect(self._add_exercise)
        edit_menu.addAction(add_exercise_action)

        # פעולת ניקוי עמוד נוכחי
        clear_current_action = QAction("מחק עמוד", self)
        clear_current_action.triggered.connect(self._clear_current_tab)
        edit_menu.addAction(clear_current_action)

        # פעולת ניקוי נתונים בעמוד הנוכחי
        clear_data_action = QAction("נקה עמוד", self)
        clear_data_action.triggered.connect(self._clear_current_tab_data)
        edit_menu.addAction(clear_data_action)

        # פעולת ניקוי כל העמודים
        clear_all_action = QAction("מחק הכל", self)
        clear_all_action.triggered.connect(self._clear_all_tabs)
        edit_menu.addAction(clear_all_action)

        # שמירה בסגירה
        self._closing = False

    def _add_exercise(self):
        title, ok = QInputDialog.getText(self, "הוספת תרגיל", "שם התרגיל:")
        if ok and title.strip():
            # בדוק אם תרגיל עם שם זהה כבר קיים
            existing = set()
            for i in range(self.tab_widget.count()):
                tab = self.tab_widget.widget(i)
                if isinstance(tab, ExerciseTab):
                    existing.add(tab.exercise_name)
            if title not in existing:
                tab = ExerciseTab(title)
                self.tab_widget.addTab(tab, title)
                self.tab_widget.setCurrentWidget(tab)

    def _save_current_tab(self):
        current = self.tab_widget.currentWidget()
        if isinstance(current, ExerciseTab):
            try:
                current.save_state()
            except Exception as e:
                QMessageBox.warning(self, "שגיאה בשמירה", str(e))

    def _restore_current_tab(self):
        current = self.tab_widget.currentWidget()
        if isinstance(current, ExerciseTab):
            try:
                current.load_state()
                self.statusBar().showMessage("שוחזר בהצלחה מקובץ", 2000)
            except Exception as e:
                QMessageBox.warning(self, "שגיאה בשחזור", str(e))
    
    def _export_to_excel(self):
        """ייצוא כל העמודים לקובץ אקסל, כל עמוד לגיליון נפרד"""
        # בדוק אם יש עמודים לייצא
        if self.tab_widget.count() == 0:
            QMessageBox.warning(self, "שגיאה", "אין עמודים לייצוא")
            return
        
        # צור שם קובץ ברירת מחדל
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        default_filename = f"תרגילים_{timestamp}.xlsx"
        
        # בקש מהמשתמש שם קובץ
        from PySide6.QtWidgets import QFileDialog
        filename, _ = QFileDialog.getSaveFileName(
            self,
            "שמור קובץ אקסל",
            default_filename,
            "Excel Files (*.xlsx)"
        )
        
        if not filename:
            return  # המשתמש ביטל
        
        try:
            # צור חוברת עבודה חדשה
            wb = Workbook()
            # הסר את הגיליון הראשון שנוצר אוטומטית
            if wb.active:
                wb.remove(wb.active)
            
            # עבור על כל העמודים באפליקציה
            for tab_index in range(self.tab_widget.count()):
                tab = self.tab_widget.widget(tab_index)
                if not isinstance(tab, ExerciseTab):
                    continue
                
                exercise_name = tab.exercise_name
                
                # צור גיליון חדש לתרגיל הזה
                ws = wb.create_sheet(title=exercise_name[:31])  # שם גיליון מוגבל ל-31 תווים
                
                # הגדר את הגיליון להיות מימין לשמאל (RTL)
                ws.sheet_view.rightToLeft = True
                
                # הוסף כותרות
                headers = ["תאריך", "משקל", "סטים", "חזרות", "סט אחרון"]
                ws.append(headers)
                
                # עצב את שורת הכותרת
                header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
                header_font = Font(bold=True, color="FFFFFF", size=12)
                header_alignment = Alignment(horizontal="center", vertical="center")
                
                for col_num, header in enumerate(headers, 1):
                    cell = ws.cell(row=1, column=col_num)
                    cell.fill = header_fill
                    cell.font = header_font
                    cell.alignment = header_alignment
                
                # הוסף נתונים מהטבלה
                table = tab.table
                for row in range(table.rowCount()):
                    row_data = []
                    
                    # קרא את כל הערכים מהשורה
                    values = []
                    for col in range(table.columnCount()):
                        item = table.item(row, col)
                        values.append(item.text() if item else "")
                    
                    # הוסף בסדר הפוך: תאריך (4), משקל (3), סטים (2), חזרות (1), סט אחרון (0)
                    for col_index in [4, 3, 2, 1, 0]:
                        text = values[col_index]
                        if not text:
                            row_data.append("")
                            continue
                        
                        # עמודה 4 מהטבלה היא תאריך
                        if col_index == 4:
                            try:
                                date_obj = datetime.strptime(text, "%d/%m/%Y")
                                row_data.append(date_obj)
                            except ValueError:
                                row_data.append(text)
                        # עמודה 3 היא משקל - נסיר את "kg"
                        elif col_index == 3:
                            clean_text = text.replace("kg", "").replace("KG", "").strip()
                            parts = text.split()
                            if parts:
                                clean_text = parts[0].replace(",", ".")
                            
                            try:
                                if '.' in clean_text or ',' in clean_text:
                                    row_data.append(float(clean_text.replace(",", ".")))
                                else:
                                    row_data.append(int(clean_text))
                            except (ValueError, AttributeError):
                                try:
                                    import re
                                    number_str = re.search(r'[\d.,]+', text)
                                    if number_str:
                                        num_text = number_str.group().replace(",", ".")
                                        if '.' in num_text:
                                            row_data.append(float(num_text))
                                        else:
                                            row_data.append(int(num_text))
                                    else:
                                        row_data.append(0)
                                except Exception:
                                    row_data.append(0)
                        # עמודות 0-2 הן מספרים אחרים
                        else:
                            try:
                                if '.' in text:
                                    row_data.append(float(text))
                                else:
                                    row_data.append(int(text))
                            except ValueError:
                                row_data.append(text)
                    
                    ws.append(row_data)
                
                # הפוך את הטבלה לטבלה חכמה של Excel
                from openpyxl.worksheet.table import Table, TableStyleInfo
                
                max_row = ws.max_row
                max_col = ws.max_column
                if max_row > 1:
                    table_range = f"A1:{get_column_letter(max_col)}{max_row}"
                    excel_table = Table(displayName=f"DataTable{tab_index}", ref=table_range)
                    
                    style = TableStyleInfo(
                        name="TableStyleMedium2",
                        showFirstColumn=False,
                        showLastColumn=False,
                        showRowStripes=True,
                        showColumnStripes=False
                    )
                    excel_table.tableStyleInfo = style
                    ws.add_table(excel_table)
                
                # התאם רוחב עמודות
                for col in range(1, max_col + 1):
                    ws.column_dimensions[get_column_letter(col)].width = 15
                
                # עצב את עמודת התאריך
                for row in range(2, max_row + 1):
                    date_cell = ws.cell(row=row, column=1)
                    if date_cell.value and isinstance(date_cell.value, datetime):
                        date_cell.number_format = 'DD/MM/YYYY'
                
                # צור גרף קווי
                if max_row > 1:
                    chart = LineChart()
                    chart.title = f"גרף משקלים - {exercise_name}"
                    chart.style = 10
                    chart.y_axis.title = None
                    chart.x_axis.title = None
                    chart.legend = None
                    
                    data = Reference(ws, min_col=2, min_row=2, max_col=2, max_row=max_row)
                    dates = Reference(ws, min_col=1, min_row=2, max_row=max_row)
                    
                    chart.add_data(data, titles_from_data=False)
                    chart.set_categories(dates)
                    
                    if len(chart.series) > 0:
                        series = chart.series[0]
                        series.smooth = False
                        
                        from openpyxl.chart.marker import Marker
                        from openpyxl.drawing.line import LineProperties
                        
                        line = LineProperties()
                        line.solidFill = "2196F3"
                        line.width = 25000
                        series.graphicalProperties.line = line
                        
                        marker = Marker(symbol='circle', size=5)
                        series.marker = marker
                    
                    chart.width = 20
                    chart.height = 12
                    
                    ws.add_chart(chart, f"A{max_row + 3}")
            
            # שמור את הקובץ
            wb.save(filename)
            
            self.statusBar().showMessage(f"נשמר בהצלחה: {filename}", 3000)
            QMessageBox.information(self, "הצלחה", f"הקובץ נשמר בהצלחה:\n{filename}\n\nיוצאו {self.tab_widget.count()} תרגילים")
            
        except Exception as e:
            QMessageBox.critical(self, "שגיאה", f"שגיאה בשמירת הקובץ:\n{str(e)}")
    
    def _undo_current_tab(self):
        """ביטול הפעולה האחרונה בעמוד הנוכחי"""
        current = self.tab_widget.currentWidget()
        if isinstance(current, ExerciseTab):
            current.undo()
    
    def _redo_current_tab(self):
        """שחזור הפעולה שבוטלה בעמוד הנוכחי"""
        current = self.tab_widget.currentWidget()
        if isinstance(current, ExerciseTab):
            current.redo()

    def _show_help(self):
        """הצגת חלון עזרה עם מידע על האפליקציה"""
        help_text = """
        <div dir="rtl" style="text-align: left; font-size: 11pt; direction: rtl;">
        <h2 style="text-align: left;">אפליקציית מעקב משקלים</h2>
        <p style="text-align: left;">אפליקציה לניהול ומעקב אחר התקדמות באימוני כוח.</p>
        
        <h3 style="text-align: left;">תכונות עיקריות:</h3>
        <ul style="text-align: left;">
            <li style="text-align: left;"><b>ניהול תרגילים מרובים</b> - ניתן ליצור עמודים נפרדים לכל תרגיל</li>
            <li style="text-align: left;"><b>מעקב מפורט</b> - רישום משקל, מספר סטים, חזרות וסט אחרון</li>
            <li style="text-align: left;"><b>גרפים ויזואליים</b> - הצגת התקדמות לאורך זמן</li>
            <li style="text-align: left;"><b>חישוב סטטיסטיקות</b> - סיכום אימונים וסך משקל מצטבר</li>
            <li style="text-align: left;"><b>שמירה אוטומטית</b> - כל הנתונים נשמרים למחשב</li>
        </ul>
        
        <h3 style="text-align: left;">קיצורי מקלדת:</h3>
        <ul style="text-align: left;">
            <li style="text-align: left;"><b>Ctrl+Z</b> - אחורה (ביטול פעולה)</li>
            <li style="text-align: left;"><b>Ctrl+Y</b> - קדימה (שחזור פעולה)</li>
            <li style="text-align: left;"><b>Ctrl+S</b> - שמור</li>
            <li style="text-align: left;"><b>Ctrl+R</b> - שחזר מקובץ</li>
            <li style="text-align: left;"><b>Ctrl+N</b> - הוסף תרגיל חדש</li>
            <li style="text-align: left;"><b>Enter</b> - הוסף רשומה (כשכל השדות מלאים)</li>
            <li style="text-align: left;"><b>חיצים ↑↓</b> - מעבר בין שדות קלט</li>
        </ul>
        
        <h3 style="text-align: left;">טיפים:</h3>
        <ul style="text-align: left;">
            <li style="text-align: left;">לחץ פעמיים על תאריך לעריכה</li>
            <li style="text-align: left;">בחר שורה ולחץ "מחק שורה" למחיקה</li>
            <li style="text-align: left;">השתמש ב"הצג גרף" לראות התקדמות ויזואלית</li>
        </ul>
        
        <p style="margin-top: 20px; color: #666; text-align: left;">
        גרסה 1.0 | 2025
        </p>
        </div>
        """
        
        msg = QMessageBox(self)
        msg.setWindowTitle("עזרה - מעקב משקלים")
        msg.setTextFormat(Qt.TextFormat.RichText)
        msg.setText(help_text)
        msg.setIcon(QMessageBox.Icon.Information)
        msg.setStandardButtons(QMessageBox.StandardButton.Ok)
        msg.exec()

    def _clear_current_tab_data(self):
        """ניקוי כל הנתונים מהעמוד הנוכחי אבל שמירת העמוד עצמו"""
        current = self.tab_widget.currentWidget()
        if not isinstance(current, ExerciseTab):
            return
            
        reply = QMessageBox.question(
            self,
            "אישור ניקוי נתונים",
            f"האם אתה בטוח שברצונך למחוק את כל הנתונים מהעמוד '{current.exercise_name}'?\n\nהעמוד יישאר קיים אך ללא נתונים.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        
        if reply == QMessageBox.StandardButton.Yes:
            try:
                # מחיקת כל השורות מהטבלה
                current.table.setRowCount(0)
                
                # איפוס כפתורי המחיקה
                current.btn_pop.setEnabled(False)
                current.btn_delete_row.setEnabled(False)
                
                # עדכון הסיכום
                current._update_summary()
                
                # סימון שיש שינויים לא שמורים
                current._has_unsaved_changes = True
                
                # מחיקת קובץ השמירה
                path = Path.cwd() / f"exercise_state_{current.exercise_name}.json"
                if path.exists():
                    os.remove(path)
                
                self.statusBar().showMessage(f"נמחקו כל הנתונים מהעמוד '{current.exercise_name}'", 2000)
            except Exception as e:
                QMessageBox.warning(self, "שגיאה בניקוי", str(e))

    def _clear_current_tab(self):
        current = self.tab_widget.currentWidget()
        if not isinstance(current, ExerciseTab):
            return
            
        reply = QMessageBox.question(
            self,
            "אישור ניקוי",
            f"האם אתה בטוח שברצונך למחוק את כל הנתונים מהעמוד '{current.exercise_name}'?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        
        if reply == QMessageBox.StandardButton.Yes:
            try:
                path = Path.cwd() / f"exercise_state_{current.exercise_name}.json"
                if path.exists():
                    os.remove(path)
                
                # מחק את הטאב הנוכחי
                idx = self.tab_widget.currentIndex()
                self.tab_widget.removeTab(idx)
                current.deleteLater()

                # אם זה היה הטאב האחרון, הצג דיאלוג ליצירת תרגיל חדש
                if self.tab_widget.count() == 0:
                    title, ok = QInputDialog.getText(self, "תרגיל ראשון", "שם התרגיל:")
                    if ok and title.strip():
                        tab = ExerciseTab(title)
                        self.tab_widget.addTab(tab, title)

                self.statusBar().showMessage(f"נמחקו כל הנתונים מהעמוד '{current.exercise_name}'", 2000)
            except Exception as e:
                QMessageBox.warning(self, "שגיאה בניקוי", str(e))

    def _clear_all_tabs(self):
        reply = QMessageBox.question(
            self,
            "אישור ניקוי",
            "האם אתה בטוח שברצונך למחוק את כל הנתונים מכל העמודים?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        
        if reply == QMessageBox.StandardButton.Yes:
            try:
                # מחק את כל הקבצים
                for file in Path.cwd().glob("exercise_state_*.json"):
                    try:
                        os.remove(file)
                    except Exception:
                        pass

                # סגור את כל הטאבים
                while self.tab_widget.count() > 0:
                    tab = self.tab_widget.widget(0)
                    self.tab_widget.removeTab(0)
                    if isinstance(tab, ExerciseTab):
                        tab.deleteLater()

                self.statusBar().showMessage("נמחקו כל הנתונים וכל העמודים", 2000)

                # הצג דיאלוג ליצירת תרגיל חדש
                title, ok = QInputDialog.getText(self, "תרגיל ראשון", "שם התרגיל:")
                if ok and title.strip():
                    tab = ExerciseTab(title)
                    self.tab_widget.addTab(tab, title)

            except Exception as e:
                QMessageBox.warning(self, "שגיאה בניקוי", str(e))

    def closeEvent(self, event):
        if self._closing:
            event.accept()
            return
            
        # בדיקה אם יש שינויים שלא נשמרו
        unsaved_tabs = []
        for i in range(self.tab_widget.count()):
            tab = self.tab_widget.widget(i)
            if isinstance(tab, ExerciseTab) and tab._has_unsaved_changes:
                unsaved_tabs.append(tab)
        
        if unsaved_tabs:
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Icon.Question)
            msg.setWindowTitle("שינויים לא שמורים")
            if len(unsaved_tabs) == 1:
                msg.setText(f"יש שינויים שלא נשמרו בעמוד '{unsaved_tabs[0].exercise_name}'.\nהאם ברצונך לשמור לפני היציאה?")
            else:
                msg.setText("יש שינויים שלא נשמרו במספר עמודים.\nהאם ברצונך לשמור לפני היציאה?")
            msg.setStandardButtons(
                QMessageBox.StandardButton.Save | 
                QMessageBox.StandardButton.Discard | 
                QMessageBox.StandardButton.Cancel
            )
            msg.setDefaultButton(QMessageBox.StandardButton.Save)
            ret = msg.exec()
            
            if ret == QMessageBox.StandardButton.Save:
                # שמירת כל העמודים עם שינויים
                for tab in unsaved_tabs:
                    try:
                        tab.save_state()
                    except Exception:
                        pass
                self._closing = True
                event.accept()
            elif ret == QMessageBox.StandardButton.Discard:
                self._closing = True
                event.accept()
            else:  # Cancel
                event.ignore()
                return
        else:
            self._closing = True
            event.accept()


def apply_stylesheet(app: QApplication):
    # הגדרת סגנון כללי לאפליקציה
    app.setStyleSheet("""
        QMainWindow {
            background-color: #f0f0f0;
        }
        QTabWidget::pane {
            border: 1px solid #cccccc;
            background: white;
            border-radius: 5px;
        }
        QTabBar::tab {
            background: #e1e1e1;
            border: 1px solid #cccccc;
            padding: 8px 15px;
            margin-right: 2px;
            border-top-left-radius: 4px;
            border-top-right-radius: 4px;
            font-size: 11pt;
        }
        QTabBar::tab:selected {
            background: white;
            border-bottom-color: white;
        }
        QLineEdit {
            padding: 6px;
            border: 1px solid #cccccc;
            border-radius: 4px;
            background-color: white;
            font-size: 11pt;
        }
        QLineEdit:focus {
            border: 1px solid #2196F3;
        }
        QPushButton {
            padding: 6px 12px;
            background-color: #2196F3;
            color: white;
            border: none;
            border-radius: 4px;
            font-size: 11pt;
        }
        QPushButton:hover {
            background-color: #1976D2;
        }
        QPushButton:disabled {
            background-color: #BDBDBD;
        }
        QTableWidget {
            border: 1px solid #cccccc;
            border-radius: 4px;
            font-size: 11pt;
        }
        QTableWidget::item {
            padding: 4px;
            min-height: 24px;
        }
        QTableWidget {
            padding: 4px;
            min-height: 400px;
        }
        QTableWidget {
            gridline-color: #cccccc;
        }
        QTableWidget::item:selected {
            background-color: #E3F2FD;
            color: black;
        }
        QHeaderView::section {
            background-color: #f5f5f5;
            padding: 6px;
            border: 1px solid #cccccc;
            font-size: 11pt;
            font-weight: bold;
        }
        QLabel {
            font-size: 11pt;
        }
        QMenu {
            background-color: white;
            border: 1px solid #cccccc;
        }
        QMenu::item {
            padding: 6px 20px;
        }
        QMenu::item:selected {
            background-color: #E3F2FD;
        }
        QStatusBar {
            background-color: #f5f5f5;
            color: #333333;
            font-size: 10pt;
        }
        QToolBar {
            background-color: #f5f5f5;
            border-bottom: 1px solid #cccccc;
            spacing: 5px;
            padding: 5px;
        }
        QToolBar QToolButton {
            background-color: #2196F3;
            color: white;
            border-radius: 4px;
            padding: 5px 10px;
            font-size: 11pt;
        }
        QToolBar QToolButton:hover {
            background-color: #1976D2;
        }
        QMessageBox {
            font-size: 11pt;
        }
        QMessageBox QPushButton {
            min-width: 80px;
        }
    """)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    apply_stylesheet(app)
    window = MainWindow()

    # חפש קבצי שמירה קיימים
    exercise_files = list(Path.cwd().glob("exercise_state_*.json"))
    
    if exercise_files:
        # אם יש קבצים קיימים, טען אותם
        for file in exercise_files:
            # חלץ את שם התרגיל מהקובץ
            exercise_name = file.stem.replace("exercise_state_", "")
            tab = ExerciseTab(exercise_name)
            window.tab_widget.addTab(tab, exercise_name)
    else:
        # אם אין קבצים קיימים, בקש שם תרגיל חדש
        title, ok = QInputDialog.getText(window, "תרגיל ראשון", "שם התרגיל:")
        if ok and title.strip():
            tab = ExerciseTab(title)
            window.tab_widget.addTab(tab, title)

    window.show()
    app.exec()