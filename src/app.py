import json
import os
import sys
from datetime import datetime
from pathlib import Path
from typing import Any

os.environ.setdefault('QT_API', 'pyside6')
from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure

os.environ["QT_API"] = "pyside6"
import matplotlib.dates as mdates
from PySide6.QtCore import QEvent, QSize, Qt
from PySide6.QtGui import (
    QAction,
    QDoubleValidator,
    QIntValidator,
    QKeySequence,
    QValidator,
)
from PySide6.QtWidgets import (
    QApplication,
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
        self._init_ui()
        try:
            self.load_state()
        except Exception:
            pass
        # אחרי טעינת המצב, נאפס את דגל השינויים
        self._has_unsaved_changes = False

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
        
        # התחלתי מצב כפתורי מחיקה - מבוטלים
        self.btn_pop.setEnabled(False)
        self.btn_delete_row.setEnabled(False)

        self.btn_add.setEnabled(False)
        self.btn_pop.setEnabled(False)
        self.btn_delete_row.setEnabled(False)

        # יצירת תצוגת סיכום
        summary_layout = QVBoxLayout()
        
        # עיצוב תוויות הסיכום
        summary_style = """
            QLabel {
                font-size: 14pt;
                font-weight: bold;
                color: #1976D2;
                padding: 5px;
            }
        """
        
        self.total_exercises_label = QLabel("סה\"כ אימונים: 0")
        self.total_exercises_label.setStyleSheet(summary_style)
        self.total_exercises_label.setAlignment(Qt.AlignmentFlag.AlignRight)
        
        self.total_weight_label = QLabel("סה\"כ משקל שהרמת: 0 ק\"ג")
        self.total_weight_label.setStyleSheet(summary_style)
        self.total_weight_label.setAlignment(Qt.AlignmentFlag.AlignRight)
        
        summary_layout.addWidget(self.total_exercises_label)
        summary_layout.addWidget(self.total_weight_label)
        summary_layout.addStretch()
        
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
        inputs_and_summary.addLayout(input_layout)
        inputs_and_summary.addLayout(summary_layout)
        
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
        self.btn_plot.clicked.connect(self.plot_selected_exercise)
        self.btn_back.clicked.connect(self.restore_normal_view)
        
        # חיבור לאירוע בחירת שורה בטבלה
        self.table.itemSelectionChanged.connect(self._update_delete_button)

        # מסגרת גרף
        self.figure = Figure(figsize=(6, 4))
        self.canvas = FigureCanvas(self.figure)

        # הוספת רכיבים לממשק
        bottom_buttons = QHBoxLayout()
        bottom_buttons.addWidget(self.btn_add)
        bottom_buttons.addWidget(self.btn_pop)
        bottom_buttons.addWidget(self.btn_delete_row)
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
        self.total_exercises_label.setText(f"סה\"כ אימונים: {exercises_count}")
        
        # עדכון סך המשקל
        total_weight = self._calculate_total_weight()
        self.total_weight_label.setText(f"סה\"כ משקל שהרמת: {total_weight:,.0f} ק\"ג")

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
        date_str = datetime.now().strftime("%Y-%m-%d")

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
        self.btn_delete_row.hide()  # הסתרת כפתור מחק שורה
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
                dval = datetime.strptime(dstr, "%Y-%m-%d")
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
        ax.xaxis.set_major_formatter(mdates.DateFormatter('%Y-%m-%d'))
        self.figure.autofmt_xdate(rotation=30)
        
        # שימוש בסימן LRM (Left-to-Right Mark) לסידור הטקסט
        LRM = '\u200E'
        title = f"גרף משקלים - {self.exercise_name}"
        ax.set_title(f"{LRM}{title[::-1]}", fontsize=12, pad=15)  # הופך את סדר האותיות
        ax.set_xlabel(f"{LRM}{'תאריך'[::-1]}", fontsize=10, labelpad=10)  # הופך את סדר האותיות
        ax.set_ylabel(f"{LRM}{'משקל )ק\"ג('[::-1]}", fontsize=10, labelpad=10)  # הופך את סדר האותיות עם סוגריים הפוכים
        
        # הגדרת רשת עדינה
        ax.grid(True, linestyle='--', alpha=0.3)
        ax.set_axisbelow(True)  # הרשת מאחורי הנתונים
        
        # עיצוב שולי הגרף
        for spine in ax.spines.values():
            spine.set_color('#cccccc')
            
        # התאמת צבע וגודל תוויות הצירים
        ax.tick_params(axis='both', colors='#666666', labelsize=9)
        
        self.canvas.draw()

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
        self.btn_delete_row.show()  # החזרת כפתור מחק שורה
        self.btn_plot.show()
        self.btn_back.hide()

    def _update_delete_button(self):
        """עדכון מצב כפתור מחיקת שורה בהתאם לבחירה"""
        selected_rows = len({idx.row() for idx in self.table.selectedIndexes()})
        self.btn_delete_row.setEnabled(selected_rows == 1)
    
    def delete_selected_row(self):
        """מחיקת השורה הנבחרת"""
        selected_rows = {idx.row() for idx in self.table.selectedIndexes()}
        if len(selected_rows) == 1:
            row = selected_rows.pop()
            self.table.removeRow(row)
            self._has_unsaved_changes = True
            self.btn_pop.setEnabled(self.table.rowCount() > 0)
            window = self.window()
            if isinstance(window, QMainWindow) and window.statusBar():
                window.statusBar().showMessage("השורה נמחקה.", 2000)
        
    def _edit_date_cell(self, row: int, column: int):
        if column != 4:  # עמודת תאריך היא 4
            self.table.clearSelection()
            return
        item = self.table.item(row, column)
        if item is None:
            return
        current = item.text() if item is not None else datetime.now().strftime("%Y-%m-%d")
        text, ok = QInputDialog.getText(self, "ערוך תאריך", "תאריך (YYYY-MM-DD):", text=current)
        if not ok:
            return
        new = text.strip()
        try:
            _ = datetime.strptime(new, "%Y-%m-%d")
        except Exception:
            QMessageBox.warning(self, "תאריך לא תקין", "פורמט התאריך צריך להיות YYYY-MM-DD")
            return
        item.setText(new)
        item.setTextAlignment(Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignVCenter)
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
        
        # פעולת שמירה בתפריט (משתמש באותו Action כמו הסרגל)
        file_menu.addAction(save_action)
        
        # פעולת שחזור
        restore_action = QAction("שחזר", self)
        restore_action.setShortcuts([QKeySequence("Ctrl+R"), QKeySequence("Ctrl+ר")])  # תמיכה באנגלית ועברית
        restore_action.triggered.connect(self._restore_current_tab)
        file_menu.addAction(restore_action)

        # תפריט עריכה
        edit_menu = self.menuBar().addMenu("עריכה")
        
        # פעולת הוספת תרגיל
        add_exercise_action = QAction("הוסף תרגיל", self)
        add_exercise_action.setShortcuts([QKeySequence("Ctrl+N"), QKeySequence("Ctrl+מ")])  # תמיכה באנגלית ועברית
        add_exercise_action.triggered.connect(self._add_exercise)
        edit_menu.addAction(add_exercise_action)

        # פעולת ניקוי עמוד נוכחי
        clear_current_action = QAction("נקה עמוד נוכחי", self)
        clear_current_action.triggered.connect(self._clear_current_tab)
        edit_menu.addAction(clear_current_action)

        # פעולת ניקוי כל העמודים
        clear_all_action = QAction("נקה את כל העמודים", self)
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
            min-height: 200px;
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