"""University Tracker desktop application, built with PySide6."""
from __future__ import annotations

import csv
import io
import json
import os
import re
import shutil
import sys
import tempfile
import webbrowser
from datetime import date, datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from PySide6.QtCore import QDate, QEvent, QObject, QSize, Qt, QTimer
from PySide6.QtGui import QAction, QColor, QIcon, QKeySequence, QPainter, QPalette, QPen, QPixmap
from PySide6.QtWidgets import (
    QApplication, QCalendarWidget, QComboBox, QDateEdit, QDialog, QDialogButtonBox, QFileDialog,
    QAbstractItemView, QFormLayout, QFrame, QGridLayout, QHBoxLayout, QHeaderView, QLabel, QLineEdit, QListView, QMainWindow,
    QMenu, QMessageBox, QPushButton, QScrollArea, QSizePolicy, QStackedWidget,
    QTabWidget, QTableWidget, QTableWidgetItem, QListWidget, QListWidgetItem,
    QTextEdit, QToolButton, QVBoxLayout, QWidget,
)

APP_NAME = "University Tracker"
APP_VERSION = "1.0.1"
FILE_NAME = "universities.xlsx"
HEADERS = [
    "Country", "University", "City", "University Type", "University Website",
    "International Students Accepted?", "International Student Notes", "Degree Level",
    "Program Name", "Department / Faculty", "AI Specialization?", "English-Taught?",
    "Program Duration", "Intake", "Application Deadline", "Minimum GPA", "GPA Scale",
    "IELTS Requirement", "TOEFL Requirement", "Other English Requirements", "GRE Requirement",
    "Other Standardized Tests", "Bachelor's Prerequisites", "Required Documents",
    "Statement of Purpose (SOP)", "Recommendation Letters", "CV / Resume", "Portfolio Required?",
    "Interview Required?", "Tuition Fee", "Tuition Currency", "Application Fee",
    "Application Fee Currency", "Scholarship Available?", "Scholarship Name", "Funding Type",
    "Funding Amount", "Funding Currency", "Full / Partial Funding", "Tuition Waiver?",
    "Research Assistantship (RA)?", "Teaching Assistantship (TA)?", "Assistantship Details",
    "Monthly / Annual Stipend", "Stipend Currency", "Estimated Living Cost",
    "Living Cost Currency", "Other Funding Conditions", "Application Portal", "Research Areas",
    "Potential Supervisors", "Research Fit", "University Ranking", "Priority",
    "Application Status", "Deadline Reminder", "Notes", "Last Checked",
]
GROUPS = {
    "University": ["Country", "University", "City", "University Type", "University Website",
                   "International Students Accepted?", "International Student Notes"],
    "Program": ["Degree Level", "Program Name", "Department / Faculty", "AI Specialization?",
                "English-Taught?", "Program Duration", "Intake", "Application Deadline"],
    "Admission": ["Minimum GPA", "GPA Scale", "IELTS Requirement", "TOEFL Requirement",
                  "Other English Requirements", "GRE Requirement", "Other Standardized Tests",
                  "Bachelor's Prerequisites", "Required Documents", "Statement of Purpose (SOP)",
                  "Recommendation Letters", "CV / Resume", "Portfolio Required?", "Interview Required?"],
    "Funding": ["Tuition Fee", "Tuition Currency", "Application Fee", "Application Fee Currency",
                "Scholarship Available?", "Scholarship Name", "Funding Type", "Funding Amount",
                "Funding Currency", "Full / Partial Funding", "Tuition Waiver?",
                "Research Assistantship (RA)?", "Teaching Assistantship (TA)?", "Assistantship Details",
                "Monthly / Annual Stipend", "Stipend Currency", "Estimated Living Cost",
                "Living Cost Currency", "Other Funding Conditions"],
    "Research & Application": ["Application Portal", "Research Areas", "Potential Supervisors",
                "Research Fit", "University Ranking", "Priority", "Application Status",
                "Deadline Reminder", "Notes", "Last Checked"],
}
YES_NO = {"International Students Accepted?", "AI Specialization?", "English-Taught?",
          "Statement of Purpose (SOP)", "CV / Resume", "Portfolio Required?",
          "Interview Required?", "Scholarship Available?", "Tuition Waiver?",
          "Research Assistantship (RA)?", "Teaching Assistantship (TA)?"}
OPTIONS = {
    "University Type": ["Public", "Private", "Research University", "Other"],
    "Degree Level": ["Bachelor's", "Master's", "PhD", "Other"],
    "Funding Type": ["Full Scholarship", "Partial Scholarship", "Tuition Waiver", "RA", "TA",
                     "RA + TA", "Fellowship", "Self-Funded", "Multiple Options", "Other"],
    "Full / Partial Funding": ["Full", "Partial", "Tuition Only", "Living Expenses Only", "Unknown"],
    "Research Fit": ["Excellent", "Very Good", "Good", "Moderate", "Poor", "Unknown"],
    "Priority": ["High", "Medium", "Low", "Backup"],
    "Application Status": ["Not Started", "Researching", "Preparing Documents", "Ready to Apply",
                           "Applied", "Interview", "Accepted", "Rejected", "Waitlisted", "Withdrawn"],
}
STAGES = OPTIONS["Application Status"]
TERMINAL = {"Applied", "Accepted", "Rejected", "Withdrawn"}
TEXT_FIELDS = {"International Student Notes", "Other English Requirements", "Bachelor's Prerequisites",
               "Required Documents", "Other Funding Conditions", "Assistantship Details",
               "Research Areas", "Potential Supervisors", "Deadline Reminder", "Notes"}
COMPARE_FIELDS = ["Degree Level", "Application Deadline", "Application Status", "Minimum GPA",
                  "GPA Scale", "IELTS Requirement", "TOEFL Requirement", "Tuition Fee",
                  "Tuition Currency", "Application Fee", "Application Fee Currency",
                  "Scholarship Available?", "Scholarship Name", "Funding Type", "Funding Amount",
                  "Funding Currency", "Full / Partial Funding", "Research Fit", "Priority"]


def chevron_icon(color: str = "#60778F") -> QIcon:
    pixmap = QPixmap(16, 16)
    pixmap.fill(Qt.GlobalColor.transparent)
    painter = QPainter(pixmap)
    painter.setRenderHint(QPainter.RenderHint.Antialiasing)
    painter.setPen(QPen(QColor(color), 1.8, Qt.PenStyle.SolidLine,
                        Qt.PenCapStyle.RoundCap, Qt.PenJoinStyle.RoundJoin))
    painter.drawLine(3, 6, 8, 11)
    painter.drawLine(8, 11, 13, 6)
    painter.end()
    return QIcon(pixmap)


class ElegantComboBox(QComboBox):
    """Consistent, comfortably sized combo box used throughout the app."""
    def __init__(self, parent=None):
        super().__init__(parent)
        self._mouse_action = None
        self.setMinimumHeight(38)
        self.setMaxVisibleItems(10)
        self.setInsertPolicy(QComboBox.InsertPolicy.NoInsert)
        self.setFocusPolicy(Qt.FocusPolicy.StrongFocus)
        popup = self.view()
        popup.setMouseTracking(True)
        popup.setVerticalScrollMode(QAbstractItemView.ScrollMode.ScrollPerPixel)
        if isinstance(popup, QListView):
            popup.setUniformItemSizes(True)
            popup.setSpacing(3)
        self.arrow_button = QToolButton(self)
        self.arrow_button.setObjectName("comboArrow")
        self.arrow_button.setIcon(chevron_icon())
        self.arrow_button.setIconSize(QSize(13, 13))
        self.arrow_button.setCursor(Qt.CursorShape.PointingHandCursor)
        self.arrow_button.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        self.arrow_button.setToolTip("Show options")
        self.arrow_button.clicked.connect(self.showPopup)
        self.set_popup_theme()
        self.arrow_button.raise_()

    def set_popup_theme(self, accent="#2563EB", hover="#F0F6FE", selected="#DCEBFF",
                        text="#24425F", arrow="#60778F"):
        self.arrow_button.setIcon(chevron_icon(arrow))
        self.view().setStyleSheet(f"""
            QListView {{
                color: {text}; background: #FFFFFF;
                border: 1px solid #C9D9EA; border-radius: 10px;
                padding: 6px; outline: 0;
            }}
            QListView::item {{
                min-height: 30px; padding: 8px 12px; margin: 2px 3px;
                border-radius: 6px;
            }}
            QListView::item:hover {{ color: {text}; background: {hover}; }}
            QListView::item:selected {{ color: {text}; background: {selected}; border-left: 3px solid {accent}; }}
            QScrollBar:vertical {{ background: #F1F5FA; width: 9px; margin: 3px 2px; border-radius: 4px; }}
            QScrollBar::handle:vertical {{ background: #B9CBE0; min-height: 28px; border-radius: 4px; margin: 1px; }}
            QScrollBar::handle:vertical:hover {{ background: {accent}; }}
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{ height: 0; border: 0; }}
            QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {{ background: transparent; }}
        """)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        width = 31
        self.arrow_button.setGeometry(self.width() - width - 1, 1, width, max(0, self.height() - 2))
        self.arrow_button.raise_()

    def mousePressEvent(self, event):
        # Open on release instead of press. On some Windows/Qt styles the popup
        # steals the still-held mouse and treats the opening click as a choice.
        if event.button() == Qt.MouseButton.LeftButton:
            self._mouse_action = "close" if self.view().isVisible() else "open"
            self.setFocus(Qt.FocusReason.MouseFocusReason)
            event.accept()
            return
        super().mousePressEvent(event)

    def mouseReleaseEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton and self._mouse_action:
            action, self._mouse_action = self._mouse_action, None
            event.accept()
            if action == "open" and self.rect().contains(event.position().toPoint()):
                QTimer.singleShot(0, self.showPopup)
            elif action == "close":
                self.hidePopup()
            return
        super().mouseReleaseEvent(event)


def app_root() -> Path:
    return Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent


def data_root() -> Path:
    """Installed copies use per-user storage; source and portable copies stay local."""
    root = app_root()
    if getattr(sys, "frozen", False) and (root / "installed-mode.txt").is_file():
        local_app_data = os.environ.get("LOCALAPPDATA")
        return (Path(local_app_data) if local_app_data else Path.home() / "AppData" / "Local") / APP_NAME
    return root


def safe_sheet_name(country: str) -> str:
    return re.sub(r'[\\/*?:\[\]]', '', str(country or "Unknown").strip())[:31] or "Unknown"


def text_value(value) -> str:
    if value is None:
        return ""
    if isinstance(value, (datetime, date)):
        return value.strftime("%Y-%m-%d")
    return str(value).strip()


def parse_deadline(value) -> date | None:
    raw = text_value(value)
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y", "%d-%m-%Y", "%b %d, %Y", "%B %d, %Y"):
        try:
            return datetime.strptime(raw, fmt).date()
        except ValueError:
            pass
    return None


def record_key(record: dict) -> tuple[str, str, str]:
    return tuple(record.get(field, "").strip().casefold()
                 for field in ("University", "Program Name", "Country"))


class UniversityStore:
    """Excel-backed storage compatible with the original tracker workbook."""
    def __init__(self):
        self.folder = data_root() / "Excel_Sheets"
        self.folder.mkdir(parents=True, exist_ok=True)
        self.path = self.folder / FILE_NAME
        self.records: list[dict] = []
        self.load()

    def load(self):
        self.records = []
        if not self.path.exists():
            return
        try:
            workbook = load_workbook(self.path, data_only=True, read_only=True)
            for sheet in workbook.worksheets:
                if sheet.max_row < 2:
                    continue
                headers = [cell.value for cell in sheet[1]]
                index = {header: i for i, header in enumerate(headers) if header}
                for row in sheet.iter_rows(min_row=2, values_only=True):
                    if not any(text_value(value) for value in row):
                        continue
                    item = {header: "" for header in HEADERS}
                    for header in HEADERS:
                        if header in index and index[header] < len(row):
                            item[header] = text_value(row[index[header]])
                    self.records.append(item)
            workbook.close()
        except Exception as exc:
            raise RuntimeError(f"Could not read {self.path}\n\n{exc}") from exc

    def save(self):
        workbook = Workbook()
        workbook.remove(workbook.active)
        grouped: dict[str, tuple[str, list[dict]]] = {}
        for record in self.records:
            base = safe_sheet_name(record.get("Country", "Unknown"))
            candidate, suffix = base, 2
            while candidate.casefold() in grouped and grouped[candidate.casefold()][0] != candidate:
                candidate = f"{base[:27]}_{suffix}"
                suffix += 1
            grouped.setdefault(candidate.casefold(), (candidate, []))[1].append(record)
        if not grouped:
            grouped["universities"] = ("Universities", [])
        for name, items in grouped.values():
            sheet = workbook.create_sheet(name)
            sheet.append(HEADERS)
            for record in items:
                sheet.append([record.get(header, "") for header in HEADERS])
                for cell in sheet[sheet.max_row]:
                    cell.alignment = Alignment(vertical="top", wrap_text=True)
            for cell in sheet[1]:
                cell.fill = PatternFill("solid", fgColor="173F6B")
                cell.font = Font(color="FFFFFF", bold=True)
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            sheet.freeze_panes = "A2"
            sheet.auto_filter.ref = f"A1:{get_column_letter(len(HEADERS))}{max(sheet.max_row, 1)}"
            sheet.row_dimensions[1].height = 34
            for col, header in enumerate(HEADERS, 1):
                sheet.column_dimensions[get_column_letter(col)].width = min(max(len(header) + 3, 14), 30)
        temp_path = None
        try:
            with tempfile.NamedTemporaryFile(prefix="universities_", suffix=".xlsx", dir=self.folder, delete=False) as temp:
                temp_path = Path(temp.name)
            workbook.save(temp_path)
            os.replace(temp_path, self.path)
        except PermissionError as exc:
            raise RuntimeError("The workbook is open or locked. Close universities.xlsx and try again.") from exc
        finally:
            if temp_path and temp_path.exists():
                temp_path.unlink(missing_ok=True)


class DeadlineCalendarFilter(QObject):
    """Open an empty deadline picker on the current month rather than year 1900."""
    def __init__(self, picker: QDateEdit):
        super().__init__(picker)
        self.picker = picker

    def eventFilter(self, watched, event):
        if (event.type() == QEvent.Type.Show
                and self.picker.date() == self.picker.minimumDate()):
            QTimer.singleShot(0, self._show_current_month)
        return False

    def _show_current_month(self):
        if self.picker.date() != self.picker.minimumDate():
            return
        today = QDate.currentDate()
        calendar = self.picker.calendarWidget()
        calendar.setCurrentPage(today.year(), today.month())


class DeadlinePickerWidget(QWidget):
    """Optional date field with a styled popup calendar and a one-click clear action."""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("deadlineField")
        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(7)
        self.picker = QDateEdit()
        self.picker.setObjectName("deadlinePicker")
        self.picker.setCalendarPopup(True)
        self.picker.setDisplayFormat("ddd, dd MMM yyyy")
        self.picker.setMinimumDate(QDate(1900, 1, 1))
        self.picker.setDate(self.picker.minimumDate())
        self.picker.setSpecialValueText("No deadline selected")
        self.picker.setToolTip("Choose an application deadline from the calendar")
        self.picker.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        calendar = self.picker.calendarWidget()
        calendar.setObjectName("deadlineCalendar")
        calendar.setGridVisible(False)
        calendar.setNavigationBarVisible(True)
        calendar.setVerticalHeaderFormat(QCalendarWidget.VerticalHeaderFormat.NoVerticalHeader)
        calendar.setHorizontalHeaderFormat(QCalendarWidget.HorizontalHeaderFormat.ShortDayNames)
        calendar.setFirstDayOfWeek(Qt.DayOfWeek.Monday)
        calendar.setMinimumSize(330, 260)
        self.calendar_filter = DeadlineCalendarFilter(self.picker)
        calendar.installEventFilter(self.calendar_filter)
        self.picker.dateChanged.connect(self._date_changed)
        self.clear_button = QPushButton("Clear")
        self.clear_button.setObjectName("clearDeadlineButton")
        self.clear_button.setToolTip("Remove the deadline")
        self.clear_button.clicked.connect(self.clear)
        layout.addWidget(self.picker, 1)
        layout.addWidget(self.clear_button)

    def _date_changed(self, _date):
        self.picker.setProperty("unparsedText", "")

    def set_value(self, value):
        raw = text_value(value)
        parsed = parse_deadline(raw)
        minimum = self.picker.minimumDate()
        qdate = QDate(parsed.year, parsed.month, parsed.day) if parsed else QDate()
        self.picker.blockSignals(True)
        if qdate.isValid() and qdate > minimum:
            self.picker.setDate(qdate)
            self.picker.setProperty("unparsedText", "")
        else:
            self.picker.setDate(minimum)
            self.picker.setProperty("unparsedText", raw if raw else "")
        self.picker.blockSignals(False)

    def value(self) -> str:
        if self.picker.date() == self.picker.minimumDate():
            return self.picker.property("unparsedText") or ""
        return self.picker.date().toString("yyyy-MM-dd")

    def clear(self):
        self.picker.blockSignals(True)
        self.picker.setDate(self.picker.minimumDate())
        self.picker.setProperty("unparsedText", "")
        self.picker.blockSignals(False)


class RecordDialog(QDialog):
    def __init__(self, parent, record: dict | None = None):
        super().__init__(parent)
        self.setWindowTitle("Edit application" if record else "Add a program")
        self.setMinimumSize(760, 700)
        self.fields: dict[str, QWidget] = {}
        outer = QVBoxLayout(self)
        intro = QLabel("Keep university, admissions, funding, and research details together.")
        intro.setObjectName("dialogIntro")
        outer.addWidget(intro)
        tabs = QTabWidget()
        outer.addWidget(tabs, 1)
        for group, names in GROUPS.items():
            page = QWidget()
            form = QFormLayout(page)
            form.setContentsMargins(18, 18, 18, 18)
            form.setHorizontalSpacing(20)
            form.setVerticalSpacing(10)
            for name in names:
                if name in YES_NO or name in OPTIONS:
                    widget = ElegantComboBox()
                    widget.addItem("")
                    widget.addItems(["Yes", "No"] if name in YES_NO else OPTIONS[name])
                    widget.setEditable(False)
                elif name in TEXT_FIELDS:
                    widget = QTextEdit()
                    widget.setFixedHeight(78)
                    widget.setPlaceholderText("Add details…")
                elif name == "Application Deadline":
                    widget = DeadlinePickerWidget()
                else:
                    widget = QLineEdit()
                widget.setObjectName("field")
                if name == "Application Deadline":
                    widget.setObjectName("deadlineField")
                self.fields[name] = widget
                form.addRow(name, widget)
            scroll = QScrollArea()
            scroll.setWidgetResizable(True)
            scroll.setWidget(page)
            tabs.addTab(scroll, group)
        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Save | QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(self._validate_and_accept)
        buttons.rejected.connect(self.reject)
        outer.addWidget(buttons)
        for header, value in (record or {}).items():
            if header in self.fields:
                self._set_field(header, value)

    def _set_field(self, name: str, value):
        widget = self.fields[name]
        value = text_value(value)
        if name == "Application Deadline":
            widget.set_value(value)
        elif isinstance(widget, QComboBox):
            if value and widget.findText(value) < 0:
                widget.addItem(value)
            widget.setCurrentText(value)
        elif isinstance(widget, QTextEdit):
            widget.setPlainText(value)
        else:
            widget.setText(value)

    def values(self) -> dict:
        result = {}
        for header in HEADERS:
            widget = self.fields[header]
            if header == "Application Deadline":
                result[header] = widget.value()
            elif isinstance(widget, QComboBox):
                result[header] = widget.currentText().strip()
            elif isinstance(widget, QTextEdit):
                result[header] = widget.toPlainText().strip()
            else:
                result[header] = widget.text().strip()
        return result

    def _validate_and_accept(self):
        values = self.values()
        missing = [field for field in ("Country", "University", "Program Name") if not values[field]]
        if missing:
            QMessageBox.warning(self, "Required information", "Please fill in: " + ", ".join(missing))
            self.fields[missing[0]].setFocus()
            return
        self.accept()


class TrackerWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.store = UniversityStore()
        self.filtered: list[tuple[int, dict]] = []
        self.current_page = "Dashboard"
        self.setWindowTitle(APP_NAME)
        self.setMinimumSize(1120, 740)
        self.resize(1510, 960)
        self.setStyleSheet(STYLESHEET)
        self._build_shell()
        self._build_pages()
        self._install_shortcuts()
        self.refresh_all()
        self.statusBar().showMessage(f"Ready · workbook: {self.store.path}")

    def _build_shell(self):
        shell = QWidget()
        shell_layout = QHBoxLayout(shell)
        shell_layout.setContentsMargins(0, 0, 0, 0)
        shell_layout.setSpacing(0)
        sidebar = QFrame()
        sidebar.setObjectName("sidebar")
        sidebar.setFixedWidth(238)
        side = QVBoxLayout(sidebar)
        side.setContentsMargins(18, 24, 18, 18)
        side.setSpacing(8)
        brand = QLabel("◈  UNI TRACKER")
        brand.setObjectName("brand")
        side.addWidget(brand)
        tagline = QLabel("APPLICATION PLANNER")
        tagline.setObjectName("sideTagline")
        side.addWidget(tagline)
        side.addSpacing(25)
        self.nav_buttons = {}
        for name, icon in (("Dashboard", "▦"), ("Applications", "▤"), ("Deadlines", "◷"),
                           ("Compare", "⇄"), ("Data tools", "⚙")):
            button = QPushButton(f"{icon}    {name}")
            button.setObjectName("navButton")
            button.setCheckable(True)
            button.clicked.connect(lambda checked=False, page=name: self.navigate(page))
            self.nav_buttons[name] = button
            side.addWidget(button)
        side.addStretch(1)
        self.sidebar_count = QLabel("0 PROGRAMS TRACKED")
        self.sidebar_count.setObjectName("sidebarCount")
        side.addWidget(self.sidebar_count)
        sidebar_note = QLabel("Your data stays on this device.")
        sidebar_note.setObjectName("sidebarFootnote")
        side.addWidget(sidebar_note)
        shell_layout.addWidget(sidebar)
        content = QWidget()
        content_layout = QVBoxLayout(content)
        content_layout.setContentsMargins(30, 23, 30, 14)
        content_layout.setSpacing(18)
        header = QHBoxLayout()
        title_col = QVBoxLayout()
        self.page_title = QLabel("Dashboard")
        self.page_title.setObjectName("pageTitle")
        self.page_subtitle = QLabel("Your university application command center")
        self.page_subtitle.setObjectName("pageSubtitle")
        title_col.addWidget(self.page_title)
        title_col.addWidget(self.page_subtitle)
        header.addLayout(title_col)
        header.addStretch(1)
        self.quick_add = QPushButton("＋  Add program")
        self.quick_add.setObjectName("primaryButton")
        self.quick_add.clicked.connect(self.add_record)
        header.addWidget(self.quick_add)
        content_layout.addLayout(header)
        self.stack = QStackedWidget()
        content_layout.addWidget(self.stack, 1)
        shell_layout.addWidget(content, 1)
        self.setCentralWidget(shell)

    def _build_pages(self):
        self.pages = {}
        for name in ("Dashboard", "Applications", "Deadlines", "Compare", "Data tools"):
            page = QWidget()
            layout = QVBoxLayout(page)
            layout.setContentsMargins(0, 0, 0, 0)
            layout.setSpacing(16)
            self.stack.addWidget(page)
            self.pages[name] = page
        self._build_dashboard(self.pages["Dashboard"].layout())
        self._build_applications(self.pages["Applications"].layout())
        self._build_deadlines(self.pages["Deadlines"].layout())
        self._build_compare(self.pages["Compare"].layout())
        self._build_tools(self.pages["Data tools"].layout())

    def _build_dashboard(self, layout):
        cards = QGridLayout()
        cards.setSpacing(12)
        self.stat_values = {}
        for column, (key, label, color) in enumerate((
            ("total", "PROGRAMS TRACKED", "blue"), ("active", "IN PROGRESS", "purple"),
            ("accepted", "OFFERS RECEIVED", "green"), ("upcoming", "DUE IN 14 DAYS", "orange"),
        )):
            card = QFrame()
            card.setObjectName("statCard")
            box = QVBoxLayout(card)
            box.setContentsMargins(20, 17, 20, 17)
            small = QLabel(label)
            small.setObjectName("statLabel")
            value = QLabel("0")
            value.setObjectName(f"statValue_{color}")
            box.addWidget(small)
            box.addWidget(value)
            cards.addWidget(card, 0, column)
            self.stat_values[key] = value
        layout.addLayout(cards)
        row = QHBoxLayout()
        row.setSpacing(14)
        pipeline_card = QFrame()
        pipeline_card.setObjectName("panel")
        pipeline_layout = QVBoxLayout(pipeline_card)
        pipeline_layout.setContentsMargins(20, 18, 20, 18)
        pipeline_layout.addWidget(self._section_title("Application pipeline", "Select a stage to see those programs"))
        self.pipeline_buttons = {}
        for stage in STAGES:
            button = QPushButton(stage)
            button.setObjectName("stageButton")
            button.clicked.connect(lambda checked=False, value=stage: self.open_stage(value))
            pipeline_layout.addWidget(button)
            self.pipeline_buttons[stage] = button
        pipeline_layout.addStretch(1)
        row.addWidget(pipeline_card, 1)
        deadline_card = QFrame()
        deadline_card.setObjectName("panel")
        deadline_layout = QVBoxLayout(deadline_card)
        deadline_layout.setContentsMargins(20, 18, 20, 18)
        deadline_layout.addWidget(self._section_title("Deadline center", "Overdue and upcoming applications"))
        self.dashboard_table = self._table(["DUE", "UNIVERSITY", "PROGRAM", "STATUS", "DAYS"])
        self.dashboard_table.setMaximumHeight(390)
        self.dashboard_table.cellDoubleClicked.connect(self._open_table_record)
        deadline_layout.addWidget(self.dashboard_table, 1)
        row.addWidget(deadline_card, 2)
        layout.addLayout(row, 1)
        tip = QLabel("Tip: Add deadlines to your programs to keep the next steps visible here.")
        tip.setObjectName("hint")
        layout.addWidget(tip)

    def _build_applications(self, layout):
        filters = QFrame()
        filters.setObjectName("panel")
        filter_layout = QGridLayout(filters)
        filter_layout.setContentsMargins(14, 12, 14, 12)
        filter_layout.setHorizontalSpacing(9)
        self.search = QLineEdit()
        self.search.setPlaceholderText("Search universities, programs, notes…")
        self.search.textChanged.connect(self._debounce_refresh)
        filter_layout.addWidget(self.search, 0, 0, 1, 2)
        self.filter_status = self._combo(["All statuses"] + STAGES)
        self.filter_priority = self._combo(["All priorities"] + OPTIONS["Priority"])
        self.filter_country = self._combo(["All countries"])
        self.filter_degree = self._combo(["All degrees", *OPTIONS["Degree Level"]])
        self.filter_deadline = self._combo(["Any deadline", "Overdue", "Next 14 days", "Next 30 days", "Next 60 days", "No deadline"])
        self.sort_combo = self._combo(["Recently checked", "University (A–Z)", "Deadline (earliest)", "Priority first"])
        for column, widget in enumerate((self.filter_status, self.filter_priority, self.filter_country,
                                         self.filter_degree, self.filter_deadline, self.sort_combo)):
            filter_layout.addWidget(widget, 1, column)
            widget.currentTextChanged.connect(self._debounce_refresh)
        layout.addWidget(filters)
        split = QHBoxLayout()
        split.setSpacing(14)
        table_panel = QFrame()
        table_panel.setObjectName("panel")
        table_layout = QVBoxLayout(table_panel)
        table_layout.setContentsMargins(12, 12, 12, 12)
        self.application_table = self._table(["UNIVERSITY", "PROGRAM", "COUNTRY", "DEGREE", "STATUS", "PRIORITY", "DEADLINE"])
        self.application_table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.application_table.setSelectionMode(QTableWidget.SelectionMode.ExtendedSelection)
        self.application_table.itemSelectionChanged.connect(self._selection_changed)
        self.application_table.cellDoubleClicked.connect(lambda *_: self.edit_record())
        self.application_table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.application_table.customContextMenuRequested.connect(self._application_context_menu)
        table_layout.addWidget(self.application_table, 1)
        actions = QHBoxLayout()
        for label, callback, kind in (("＋ New", self.add_record, "primaryButton"),
                                      ("Edit", self.edit_record, "secondaryButton"),
                                      ("Duplicate", self.duplicate_record, "secondaryButton"),
                                      ("Delete", self.delete_records, "dangerButton")):
            button = QPushButton(label)
            button.setObjectName(kind)
            button.clicked.connect(callback)
            actions.addWidget(button)
        actions.addStretch(1)
        table_layout.addLayout(actions)
        split.addWidget(table_panel, 3)
        detail = QFrame()
        detail.setObjectName("panel")
        detail.setMinimumWidth(245)
        detail.setMaximumWidth(330)
        detail_layout = QVBoxLayout(detail)
        detail_layout.setContentsMargins(20, 18, 20, 18)
        detail_layout.addWidget(self._section_title("Program details", "Select an application"))
        self.detail_title = QLabel("Nothing selected")
        self.detail_title.setObjectName("detailTitle")
        self.detail_subtitle = QLabel("Choose a row to preview its key information.")
        self.detail_subtitle.setWordWrap(True)
        self.detail_body = QLabel("")
        self.detail_body.setObjectName("detailBody")
        self.detail_body.setWordWrap(True)
        self.detail_body.setAlignment(Qt.AlignmentFlag.AlignTop)
        detail_layout.addWidget(self.detail_title)
        detail_layout.addWidget(self.detail_subtitle)
        detail_layout.addWidget(self.detail_body, 1)
        open_link = QPushButton("Open university / portal ↗")
        open_link.setObjectName("secondaryButton")
        open_link.clicked.connect(self.open_website)
        detail_layout.addWidget(open_link)
        split.addWidget(detail)
        layout.addLayout(split, 1)

    def _build_deadlines(self, layout):
        intro = QLabel("Deadlines that need attention, sorted from soonest to latest.")
        intro.setObjectName("hint")
        layout.addWidget(intro)
        panel = QFrame()
        panel.setObjectName("panel")
        box = QVBoxLayout(panel)
        box.setContentsMargins(12, 12, 12, 12)
        self.deadline_table = self._table(["DUE DATE", "DAYS", "UNIVERSITY", "PROGRAM", "COUNTRY", "STATUS", "REMINDER"])
        self.deadline_table.cellDoubleClicked.connect(self._open_table_record)
        box.addWidget(self.deadline_table)
        layout.addWidget(panel, 1)

    def _build_compare(self, layout):
        intro = QLabel("Choose up to four programs to compare admissions, cost, funding, and research fit.")
        intro.setObjectName("hint")
        layout.addWidget(intro)
        split = QHBoxLayout()
        left = QFrame()
        left.setObjectName("panel")
        left_layout = QVBoxLayout(left)
        left_layout.setContentsMargins(16, 16, 16, 16)
        left_layout.addWidget(self._section_title("Your shortlist", "Click rows to select or deselect programs (up to four)"))
        self.compare_list = QListWidget()
        self.compare_list.setObjectName("compareList")
        self.compare_list.setAlternatingRowColors(True)
        palette = self.compare_list.palette()
        for group in (QPalette.ColorGroup.Active, QPalette.ColorGroup.Inactive, QPalette.ColorGroup.Disabled):
            palette.setColor(group, QPalette.ColorRole.Base, QColor("#FFFFFF"))
            palette.setColor(group, QPalette.ColorRole.AlternateBase, QColor("#F3F7FC"))
            palette.setColor(group, QPalette.ColorRole.Text, QColor("#24425F"))
            palette.setColor(group, QPalette.ColorRole.Highlight, QColor("#DCEBFF"))
            palette.setColor(group, QPalette.ColorRole.HighlightedText, QColor("#123A63"))
        self.compare_list.setPalette(palette)
        self.compare_list.setSelectionMode(QAbstractItemView.SelectionMode.MultiSelection)
        self.compare_list.setWordWrap(True)
        self.compare_list.setTextElideMode(Qt.TextElideMode.ElideNone)
        self.compare_list.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.compare_list.itemSelectionChanged.connect(self._compare_selection_changed)
        left_layout.addWidget(self.compare_list, 1)
        self.compare_selection_hint = QLabel("0 of 4 selected")
        self.compare_selection_hint.setObjectName("hint")
        left_layout.addWidget(self.compare_selection_hint)
        compare_button = QPushButton("Compare selected")
        compare_button.setObjectName("primaryButton")
        compare_button.clicked.connect(self.update_comparison)
        left_layout.addWidget(compare_button)
        split.addWidget(left, 1)
        right = QFrame()
        right.setObjectName("panel")
        right_layout = QVBoxLayout(right)
        right_layout.setContentsMargins(16, 16, 16, 16)
        right_layout.addWidget(self._section_title("Side-by-side", "Key decision details"))
        self.comparison_table = self._table(["DETAIL", "SELECT PROGRAMS"])
        self.comparison_table.horizontalHeader().setMinimumHeight(62)
        self.comparison_table.setRowCount(len(COMPARE_FIELDS))
        for row, field in enumerate(COMPARE_FIELDS):
            self.comparison_table.setItem(row, 0, QTableWidgetItem(field))
        prompt = QTableWidgetItem("Choose programs in the list, then click Compare selected.")
        prompt.setToolTip(prompt.text())
        self.comparison_table.setItem(0, 1, prompt)
        self._format_comparison_table()
        right_layout.addWidget(self.comparison_table, 1)
        split.addWidget(right, 2)
        layout.addLayout(split, 1)

    def _build_tools(self, layout):
        intro = QLabel("Move your shortlist in and out safely. Your working workbook is stored locally.")
        intro.setObjectName("hint")
        layout.addWidget(intro)
        grid = QGridLayout()
        grid.setSpacing(14)
        tools = [
            ("Import programs", "Import a CSV with Country, University, and Program Name columns.", "Choose CSV…", self.import_csv),
            ("Export shortlist", "Export the currently filtered applications as UTF-8 CSV.", "Export CSV…", self.export_csv),
            ("Back up workbook", "Save a dated copy of your Excel workbook before major edits.", "Create backup…", self.backup_workbook),
            ("Open data folder", "Find the workbook that this app reads and updates.", "Open folder", self.open_data_folder),
            ("Reload workbook", "Discard unsaved external changes and reload from disk.", "Reload data", self.reload_workbook),
        ]
        for index, (title, description, action, callback) in enumerate(tools):
            card = QFrame()
            card.setObjectName("toolCard")
            box = QVBoxLayout(card)
            box.setContentsMargins(20, 18, 20, 18)
            heading = QLabel(title)
            heading.setObjectName("toolTitle")
            desc = QLabel(description)
            desc.setWordWrap(True)
            desc.setObjectName("hint")
            button = QPushButton(action)
            button.setObjectName("secondaryButton")
            button.clicked.connect(callback)
            box.addWidget(heading)
            box.addWidget(desc, 1)
            box.addWidget(button)
            grid.addWidget(card, index // 2, index % 2)
        layout.addLayout(grid)
        layout.addStretch(1)
        workbook_note = QLabel(f"Workbook location\n{self.store.path}")
        workbook_note.setObjectName("pathNote")
        layout.addWidget(workbook_note)

    def _section_title(self, title: str, subtitle: str) -> QWidget:
        wrapper = QWidget()
        box = QVBoxLayout(wrapper)
        box.setContentsMargins(0, 0, 0, 8)
        box.setSpacing(3)
        label = QLabel(title)
        label.setObjectName("sectionTitle")
        note = QLabel(subtitle)
        note.setObjectName("hint")
        box.addWidget(label)
        box.addWidget(note)
        return wrapper

    def _table(self, headers: list[str]) -> QTableWidget:
        table = QTableWidget(0, len(headers))
        table.setHorizontalHeaderLabels(headers)
        table.verticalHeader().setVisible(False)
        table.setAlternatingRowColors(True)
        table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        table.setShowGrid(False)
        table.setWordWrap(True)
        table.setTextElideMode(Qt.TextElideMode.ElideNone)
        table.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        header = table.horizontalHeader()
        header.setDefaultAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        header.setStretchLastSection(False)
        header.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        header.setDefaultSectionSize(132)
        header.setMinimumSectionSize(76)
        header.setMaximumSectionSize(360)
        header.setResizeContentsPrecision(100)
        table.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        return table

    def _combo(self, values: list[str]) -> QComboBox:
        widget = ElegantComboBox()
        widget.addItems(values)
        return widget

    def _install_shortcuts(self):
        for sequence, callback in (("Ctrl+N", self.add_record), ("Ctrl+E", self.export_csv),
                                   ("Ctrl+I", self.import_csv), ("Ctrl+D", self.duplicate_record),
                                   ("Delete", self.delete_records), ("Ctrl+F", lambda: (self.navigate("Applications"), self.search.setFocus()))):
            action = QAction(self)
            action.setShortcut(QKeySequence(sequence))
            action.triggered.connect(callback)
            self.addAction(action)

    def navigate(self, page: str):
        self.current_page = page
        self.stack.setCurrentWidget(self.pages[page])
        self.page_title.setText(page)
        subtitles = {"Dashboard": "Your university application command center",
                     "Applications": "Search, sort, and manage your program shortlist",
                     "Deadlines": "Stay ahead of important application dates",
                     "Compare": "Make informed choices side by side",
                     "Data tools": "Import, export, back up, and manage your workbook"}
        self.page_subtitle.setText(subtitles[page])
        for name, button in self.nav_buttons.items():
            button.setChecked(name == page)
        if page == "Applications":
            self.refresh_applications()
        elif page == "Compare":
            self.refresh_compare_choices()

    def open_stage(self, stage: str):
        self.navigate("Applications")
        self.filter_status.setCurrentText(stage)

    def _debounce_refresh(self):
        if not hasattr(self, "search_timer"):
            self.search_timer = QTimer(self)
            self.search_timer.setSingleShot(True)
            self.search_timer.timeout.connect(self.refresh_applications)
        self.search_timer.start(170)

    def _selected_pairs(self) -> list[tuple[int, dict]]:
        rows = sorted({item.row() for item in self.application_table.selectedItems()})
        result = []
        for row in rows:
            cell = self.application_table.item(row, 0)
            if cell:
                index = cell.data(Qt.ItemDataRole.UserRole)
                if isinstance(index, int) and 0 <= index < len(self.store.records):
                    result.append((index, self.store.records[index]))
        return result

    def _fill_table(self, table: QTableWidget, rows: list[tuple[list[str], int | None]]):
        table.setUpdatesEnabled(False)
        table.clearContents()
        table.setRowCount(len(rows))
        for row_index, (values, record_index) in enumerate(rows):
            for column in range(table.columnCount()):
                value = values[column] if column < len(values) else ""
                cell = QTableWidgetItem(value or "—")
                cell.setTextAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
                cell.setToolTip(value or "—")
                if record_index is not None:
                    cell.setData(Qt.ItemDataRole.UserRole, record_index)
                table.setItem(row_index, column, cell)
        table.resizeColumnsToContents()
        table.resizeRowsToContents()
        table.setUpdatesEnabled(True)

    def refresh_applications(self, *_):
        if not hasattr(self, "application_table"):
            return
        query = self.search.text().strip().casefold()
        today = date.today()
        selected_status = self.filter_status.currentText()
        selected_priority = self.filter_priority.currentText()
        selected_country = self.filter_country.currentText()
        selected_degree = self.filter_degree.currentText()
        deadline_mode = self.filter_deadline.currentText()
        results = []
        for index, record in enumerate(self.store.records):
            if query and query not in " ".join(text_value(record.get(header)) for header in HEADERS).casefold():
                continue
            if selected_status != "All statuses" and record.get("Application Status") != selected_status:
                continue
            if selected_priority != "All priorities" and record.get("Priority") != selected_priority:
                continue
            if selected_country != "All countries" and record.get("Country") != selected_country:
                continue
            if selected_degree != "All degrees" and record.get("Degree Level") != selected_degree:
                continue
            due = parse_deadline(record.get("Application Deadline"))
            status = record.get("Application Status", "")
            days = (due - today).days if due else None
            if deadline_mode == "No deadline" and due:
                continue
            if deadline_mode == "Overdue" and (days is None or days >= 0 or status in TERMINAL):
                continue
            if deadline_mode.startswith("Next "):
                limit = int(deadline_mode.split()[1])
                if days is None or not 0 <= days <= limit or status in TERMINAL:
                    continue
            results.append((index, record))
        sort_mode = self.sort_combo.currentText()
        if sort_mode == "University (A–Z)":
            results.sort(key=lambda pair: (pair[1].get("University", "").casefold(), pair[1].get("Program Name", "").casefold()))
        elif sort_mode == "Deadline (earliest)":
            results.sort(key=lambda pair: parse_deadline(pair[1].get("Application Deadline")) or date.max)
        elif sort_mode == "Priority first":
            order = {"High": 0, "Medium": 1, "Low": 2, "Backup": 3}
            results.sort(key=lambda pair: (order.get(pair[1].get("Priority"), 4), pair[1].get("University", "").casefold()))
        else:
            results.sort(key=lambda pair: pair[1].get("Last Checked", ""), reverse=True)
        self.filtered = results
        table_rows = []
        for index, record in results:
            table_rows.append(([record.get("University", ""), record.get("Program Name", ""), record.get("Country", ""),
                                record.get("Degree Level", ""), record.get("Application Status", "Not Started"),
                                record.get("Priority", ""), record.get("Application Deadline", "")], index))
        self._fill_table(self.application_table, table_rows)
        self._selection_changed()
        self.statusBar().showMessage(f"Showing {len(results)} of {len(self.store.records)} programs")

    def _refresh_country_choices(self):
        """Update country options after data changes, never while a filter popup is active."""
        countries = sorted({record.get("Country", "").strip() for record in self.store.records
                            if record.get("Country", "").strip()}, key=str.casefold)
        values = ["All countries"] + countries
        if [self.filter_country.itemText(i) for i in range(self.filter_country.count())] == values:
            return
        current = self.filter_country.currentText()
        self.filter_country.blockSignals(True)
        self.filter_country.clear()
        self.filter_country.addItems(values)
        self.filter_country.setCurrentText(current if current in values else "All countries")
        self.filter_country.blockSignals(False)

    def _selection_changed(self):
        pairs = self._selected_pairs()
        if len(pairs) != 1:
            self.detail_title.setText(f"{len(pairs)} programs selected" if pairs else "Nothing selected")
            self.detail_subtitle.setText("Select one row to preview its key information." if len(pairs) != 1 else "Bulk actions are available below the list.")
            self.detail_body.setText("")
            return
        record = pairs[0][1]
        self.detail_title.setText(record.get("University") or "University")
        self.detail_subtitle.setText(record.get("Program Name") or "Program")
        fields = ["Country", "City", "Degree Level", "Application Status", "Priority", "Application Deadline",
                  "Minimum GPA", "IELTS Requirement", "Tuition Fee", "Scholarship Name", "Research Fit"]
        self.detail_body.setText("\n\n".join(f"{field.upper()}\n{record.get(field) or '—'}" for field in fields))

    def refresh_dashboard(self):
        records = self.store.records
        today = date.today()
        self.stat_values["total"].setText(str(len(records)))
        self.stat_values["active"].setText(str(sum(1 for r in records if r.get("Application Status") in
                                                     ("Researching", "Preparing Documents", "Ready to Apply", "Interview", "Waitlisted"))))
        self.stat_values["accepted"].setText(str(sum(1 for r in records if r.get("Application Status") == "Accepted")))
        self.stat_values["upcoming"].setText(str(sum(1 for r in records if (d := parse_deadline(r.get("Application Deadline")))
                                                       and 0 <= (d - today).days <= 14 and r.get("Application Status") not in TERMINAL)))
        for stage, button in self.pipeline_buttons.items():
            count = sum(1 for record in records if record.get("Application Status") == stage)
            button.setText(f"{stage}  ·  {count}")
        deadlines = []
        for index, record in enumerate(records):
            due = parse_deadline(record.get("Application Deadline"))
            if due and record.get("Application Status") not in TERMINAL:
                days = (due - today).days
                if days <= 60:
                    deadlines.append((due, index, record, days))
        deadlines.sort(key=lambda item: item[0])
        dashboard_rows = []
        for due, index, record, days in deadlines[:12]:
            dashboard_rows.append(([due.strftime("%b %d, %Y"), record.get("University", ""),
                                    record.get("Program Name", ""), record.get("Application Status", "Not Started"),
                                    "Overdue" if days < 0 else f"{days}d"], index))
        self._fill_table(self.dashboard_table, dashboard_rows)
        self.sidebar_count.setText(f"{len(records)} PROGRAM{'S' if len(records) != 1 else ''} TRACKED")

    def refresh_deadlines(self):
        today = date.today()
        items = []
        for index, record in enumerate(self.store.records):
            due = parse_deadline(record.get("Application Deadline"))
            if due and record.get("Application Status") not in TERMINAL:
                items.append((due, index, record))
        items.sort(key=lambda item: item[0])
        rows = []
        for due, index, record in items:
            days = (due - today).days
            rows.append(([due.strftime("%Y-%m-%d"), "Overdue" if days < 0 else f"{days} days",
                          record.get("University", ""), record.get("Program Name", ""), record.get("Country", ""),
                          record.get("Application Status", "Not Started"), record.get("Deadline Reminder", "")], index))
        self._fill_table(self.deadline_table, rows)

    def refresh_compare_choices(self):
        selected_keys = {
            item.data(Qt.ItemDataRole.UserRole)
            for item in self.compare_list.selectedItems()
        }
        self.compare_list.clear()
        occurrences: dict[tuple[str, str, str], int] = {}
        for record in self.store.records:
            label = f"{record.get('University', 'University')}  ·  {record.get('Program Name', 'Program')} ({record.get('Country', '')})"
            identity = record_key(record)
            occurrence = occurrences.get(identity, 0)
            occurrences[identity] = occurrence + 1
            selection_key = json.dumps([*identity, occurrence], ensure_ascii=False, separators=(",", ":"))
            item = QListWidgetItem(label)
            item.setData(Qt.ItemDataRole.UserRole, selection_key)
            item.setToolTip(label)
            self.compare_list.addItem(item)
            item.setSelected(selection_key in selected_keys)
        self._compare_selection_changed()

    def _compare_selection_changed(self):
        if not hasattr(self, "compare_list"):
            return
        count = len(self.compare_list.selectedItems())
        self.compare_selection_hint.setText(f"{count} of 4 selected")

    def update_comparison(self):
        selected_keys = {
            item.data(Qt.ItemDataRole.UserRole)
            for item in self.compare_list.selectedItems()
        }
        selected = []
        occurrences: dict[tuple[str, str, str], int] = {}
        for record in self.store.records:
            identity = record_key(record)
            occurrence = occurrences.get(identity, 0)
            occurrences[identity] = occurrence + 1
            selection_key = json.dumps([*identity, occurrence], ensure_ascii=False, separators=(",", ":"))
            if selection_key in selected_keys:
                selected.append(record)
        if len(selected) > 4:
            QMessageBox.information(self, "Choose up to four", "Deselect programs until four or fewer remain, then compare again.")
            return
        self.comparison_table.setColumnCount(max(1, len(selected)) + 1)
        self.comparison_table.horizontalHeader().setMinimumHeight(62)
        if selected:
            labels = []
            for record in selected:
                university = record.get("University", "University").strip() or "University"
                program = record.get("Program Name", "Program").strip() or "Program"
                country = record.get("Country", "").strip()
                labels.append(f"{university}\n{program}" + (f"\n{country}" if country else ""))
            self.comparison_table.setHorizontalHeaderLabels(["DETAIL"] + labels)
        else:
            self.comparison_table.setHorizontalHeaderLabels(["DETAIL", "SELECT PROGRAMS"])
        self.comparison_table.setRowCount(len(COMPARE_FIELDS))
        for row, field in enumerate(COMPARE_FIELDS):
            self.comparison_table.setItem(row, 0, QTableWidgetItem(field))
            for col, record in enumerate(selected, 1):
                value = record.get(field, "") or "—"
                cell = QTableWidgetItem(value)
                cell.setToolTip(value)
                self.comparison_table.setItem(row, col, cell)
        if not selected:
            prompt = QTableWidgetItem("Choose programs in the list, then click Compare selected.")
            prompt.setToolTip(prompt.text())
            self.comparison_table.setItem(0, 1, prompt)
        for column in range(1, self.comparison_table.columnCount()):
            header_item = self.comparison_table.horizontalHeaderItem(column)
            if selected:
                record = selected[column - 1]
                header_item.setToolTip(" · ".join(filter(None, (
                    record.get("University", "").strip(), record.get("Program Name", "").strip(),
                    record.get("Country", "").strip(),
                ))))
        self._format_comparison_table()

    def _format_comparison_table(self):
        table = self.comparison_table
        table.setSelectionMode(QTableWidget.SelectionMode.NoSelection)
        table.setAlternatingRowColors(False)
        for column in range(table.columnCount()):
            table.setColumnWidth(column, 175 if column == 0 else 230)
            header = table.horizontalHeaderItem(column)
            if header:
                alignment = Qt.AlignmentFlag.AlignLeft if column == 0 else Qt.AlignmentFlag.AlignHCenter
                header.setTextAlignment(alignment | Qt.AlignmentFlag.AlignVCenter)
        foreground = QColor("#173F5F")
        for row in range(table.rowCount()):
            background = QColor("#FFFFFF" if row % 2 == 0 else "#EEF4FB")
            for column in range(table.columnCount()):
                cell = table.item(row, column)
                if not cell:
                    continue
                cell.setForeground(foreground)
                cell.setBackground(background)
                alignment = Qt.AlignmentFlag.AlignLeft if column == 0 else Qt.AlignmentFlag.AlignHCenter
                cell.setTextAlignment(alignment | Qt.AlignmentFlag.AlignVCenter)
        table.resizeRowsToContents()

    def refresh_all(self):
        self.refresh_dashboard()
        self._refresh_country_choices()
        self.refresh_applications()
        self.refresh_deadlines()
        if hasattr(self, "compare_list"):
            self.refresh_compare_choices()

    def _open_table_record(self, row: int, _column: int):
        table = self.sender()
        cell = table.item(row, 0)
        if not cell:
            return
        index = cell.data(Qt.ItemDataRole.UserRole)
        if isinstance(index, int) and index < len(self.store.records):
            self._edit_index(index)

    def _application_context_menu(self, position):
        menu = QMenu(self)
        edit_action = menu.addAction("Edit selected")
        duplicate_action = menu.addAction("Duplicate selected")
        menu.addSeparator()
        status_menu = menu.addMenu("Set application status")
        status_actions = {status_menu.addAction(stage): stage for stage in STAGES}
        menu.addSeparator()
        delete_action = menu.addAction("Delete selected")
        chosen = menu.exec(self.application_table.viewport().mapToGlobal(position))
        if chosen == edit_action:
            self.edit_record()
        elif chosen == duplicate_action:
            self.duplicate_record()
        elif chosen == delete_action:
            self.delete_records()
        elif chosen in status_actions:
            self.set_selected_status(status_actions[chosen])

    def add_record(self):
        dialog = RecordDialog(self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            item = dialog.values()
            duplicate = self._find_duplicate(item)
            if duplicate is not None and QMessageBox.question(
                    self, "Possible duplicate", "This university and program already exist. Update that record?") == QMessageBox.StandardButton.Yes:
                self._replace_record(duplicate, item)
            elif duplicate is None:
                self._append_record(item)

    def _find_duplicate(self, candidate: dict, excluding: int | None = None) -> int | None:
        key = (candidate.get("University", "").strip().casefold(), candidate.get("Program Name", "").strip().casefold())
        return next((index for index, record in enumerate(self.store.records)
                     if index != excluding and (record.get("University", "").strip().casefold(),
                     record.get("Program Name", "").strip().casefold()) == key), None)

    def _append_record(self, item: dict):
        item["Last Checked"] = item.get("Last Checked") or date.today().isoformat()
        self.store.records.append(item)
        self._save_refresh(f"Added {item.get('University')} · {item.get('Program Name')}")

    def _replace_record(self, index: int, item: dict):
        item["Last Checked"] = item.get("Last Checked") or date.today().isoformat()
        self.store.records[index] = item
        self._save_refresh("Program updated")

    def _edit_index(self, index: int):
        if not 0 <= index < len(self.store.records):
            return
        dialog = RecordDialog(self, self.store.records[index])
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return
        item = dialog.values()
        duplicate = self._find_duplicate(item, excluding=index)
        if duplicate is not None:
            if QMessageBox.question(self, "Possible duplicate", "Another matching university/program exists. Update that record instead?") != QMessageBox.StandardButton.Yes:
                return
            index = duplicate
        item["Last Checked"] = date.today().isoformat()
        self.store.records[index] = item
        self._save_refresh("Program updated")

    def edit_record(self):
        selected = self._selected_pairs()
        if len(selected) != 1:
            QMessageBox.information(self, "Select one program", "Choose a single row to edit.")
            return
        self._edit_index(selected[0][0])

    def duplicate_record(self):
        selected = self._selected_pairs()
        if len(selected) != 1:
            QMessageBox.information(self, "Select one program", "Choose a single row to duplicate.")
            return
        copy = dict(selected[0][1])
        copy["Last Checked"] = ""
        dialog = RecordDialog(self, copy)
        dialog.setWindowTitle("Duplicate program")
        if dialog.exec() == QDialog.DialogCode.Accepted:
            item = dialog.values()
            duplicate = self._find_duplicate(item)
            if duplicate is None:
                self._append_record(item)
            elif QMessageBox.question(self, "Possible duplicate", "This university and program already exist. Update that record?") == QMessageBox.StandardButton.Yes:
                self._replace_record(duplicate, item)

    def delete_records(self):
        selected = self._selected_pairs()
        if not selected:
            QMessageBox.information(self, "Select programs", "Select one or more applications to delete.")
            return
        count = len(selected)
        prompt = f"Delete {count} selected program{'s' if count != 1 else ''}? This cannot be undone."
        if QMessageBox.question(self, "Delete applications", prompt) != QMessageBox.StandardButton.Yes:
            return
        indexes = {index for index, _ in selected}
        self.store.records = [record for index, record in enumerate(self.store.records) if index not in indexes]
        self._save_refresh(f"Deleted {count} program(s)")

    def set_selected_status(self, status: str):
        selected = self._selected_pairs()
        if not selected:
            return
        for _index, record in selected:
            record["Application Status"] = status
            record["Last Checked"] = date.today().isoformat()
        self._save_refresh(f"Updated {len(selected)} program(s) to {status}")

    def _save_refresh(self, message: str):
        try:
            self.store.save()
            self.refresh_all()
            self.statusBar().showMessage(f"{message} · saved", 6000)
        except Exception as exc:
            QMessageBox.critical(self, "Could not save", str(exc))
            try:
                self.store.load()
                self.refresh_all()
            except Exception:
                pass

    def open_website(self):
        selected = self._selected_pairs()
        if len(selected) != 1:
            QMessageBox.information(self, "Select one program", "Choose one application to open its website or portal.")
            return
        record = selected[0][1]
        url = record.get("University Website") or record.get("Application Portal")
        if not url:
            QMessageBox.information(self, "No website", "Add a university website or application portal to this record first.")
            return
        if not url.startswith(("https://", "http://")):
            url = "https://" + url
        webbrowser.open(url)

    def import_csv(self):
        path, _ = QFileDialog.getOpenFileName(self, "Import programs", "", "CSV files (*.csv);;All files (*.*)")
        if not path:
            return
        try:
            raw = Path(path).read_bytes()
            try:
                content = raw.decode("utf-8-sig")
            except UnicodeDecodeError:
                content = raw.decode("cp1252")
            reader = csv.DictReader(io.StringIO(content, newline=""))
            if not reader.fieldnames or not {"Country", "University", "Program Name"}.issubset(reader.fieldnames):
                raise ValueError("CSV must contain Country, University, and Program Name columns.")
            imported = []
            for row in reader:
                item = {header: text_value(row.get(header)) for header in HEADERS}
                if item["Country"] and item["University"] and item["Program Name"]:
                    imported.append(item)
            if not imported:
                QMessageBox.information(self, "Nothing to import", "No complete program rows were found.")
                return
            update = QMessageBox.question(self, "Matching records", "Update existing records with the same university/program?\nChoose No to skip duplicates.") == QMessageBox.StandardButton.Yes
            by_key = {(r.get("University", "").strip().casefold(), r.get("Program Name", "").strip().casefold()): i
                      for i, r in enumerate(self.store.records)}
            added = updated = skipped = 0
            for item in imported:
                key = (item["University"].casefold(), item["Program Name"].casefold())
                if key in by_key:
                    if update:
                        self.store.records[by_key[key]].update(item)
                        updated += 1
                    else:
                        skipped += 1
                else:
                    by_key[key] = len(self.store.records)
                    self.store.records.append(item)
                    added += 1
            if added or updated:
                self._save_refresh(f"Imported {added} new, updated {updated}, skipped {skipped}")
            QMessageBox.information(self, "Import complete", f"Added: {added}\nUpdated: {updated}\nSkipped: {skipped}")
        except Exception as exc:
            QMessageBox.critical(self, "Import failed", str(exc))

    def export_csv(self):
        path, _ = QFileDialog.getSaveFileName(self, "Export filtered programs", "university_tracker_export.csv", "CSV files (*.csv)")
        if not path:
            return
        try:
            with open(path, "w", newline="", encoding="utf-8-sig") as stream:
                writer = csv.DictWriter(stream, fieldnames=HEADERS, extrasaction="ignore")
                writer.writeheader()
                writer.writerows({header: item.get(header, "") for header in HEADERS} for _, item in self.filtered)
            self.statusBar().showMessage(f"Exported {len(self.filtered)} filtered programs", 6000)
        except Exception as exc:
            QMessageBox.critical(self, "Export failed", str(exc))

    def backup_workbook(self):
        if not self.store.path.exists():
            QMessageBox.information(self, "No workbook yet", "Save a program first to create a workbook backup.")
            return
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        path, _ = QFileDialog.getSaveFileName(self, "Back up workbook", f"universities_backup_{stamp}.xlsx", "Excel workbook (*.xlsx)")
        if not path:
            return
        try:
            shutil.copy2(self.store.path, path)
            self.statusBar().showMessage("Workbook backup created", 6000)
        except Exception as exc:
            QMessageBox.critical(self, "Backup failed", str(exc))

    def open_data_folder(self):
        try:
            os.startfile(str(self.store.folder))
        except (AttributeError, OSError) as exc:
            QMessageBox.warning(self, "Could not open folder", str(exc))

    def reload_workbook(self):
        if QMessageBox.question(self, "Reload workbook", "Reload data from disk and discard any external changes not saved by this app?") != QMessageBox.StandardButton.Yes:
            return
        try:
            self.store.load()
            self.refresh_all()
            self.statusBar().showMessage("Workbook reloaded", 5000)
        except Exception as exc:
            QMessageBox.critical(self, "Reload failed", str(exc))


STYLESHEET = """
QMainWindow, QWidget { background: #F3F7FC; color: #1E3550; font-family: 'Segoe UI'; font-size: 10pt; }
#sidebar { background: #102A43; color: #D7E7F7; }
#sidebar QLabel { color: #D1DEEA; background: transparent; font-size: 9pt; }
QLabel#brand { color: #FFFFFF; background: #1B4E82; border: 1px solid #2B659A; border-radius: 9px; font-size: 14pt; font-weight: 750; padding: 10px 9px; }
QLabel#sideTagline { color: #D4E3F1; background: transparent; font-size: 8pt; letter-spacing: 1px; padding: 5px 4px; }
QLabel#sidebarCount { color: #A8D7FF; background: #173F6B; border-radius: 7px; font-size: 8pt; font-weight: 700; padding: 10px; }
QLabel#sidebarFootnote { color: #BFD0E0; background: transparent; font-size: 8pt; padding: 4px 2px; }
QPushButton#navButton { text-align: left; color: #C4D5E5; background: transparent; border: 0; border-radius: 8px; padding: 12px 13px; font-size: 10pt; }
QPushButton#navButton:hover { background: #173F6B; color: white; }
QPushButton#navButton:checked { background: #2563EB; color: white; font-weight: 650; }
#pageTitle { color: #102A43; font-size: 21pt; font-weight: 700; }
#pageSubtitle, #hint { color: #71839A; font-size: 9pt; }
#panel, #statCard, #toolCard { background: white; border: 1px solid #E2EAF3; border-radius: 11px; }
#sectionTitle { color: #173F6B; font-size: 12pt; font-weight: 700; }
#statLabel { color: #71839A; font-size: 8pt; font-weight: 700; letter-spacing: .7px; }
QLabel[objectName^="statValue"] { font-size: 23pt; font-weight: 750; }
#statValue_blue { color: #2563EB; } #statValue_purple { color: #7256C4; }
#statValue_green { color: #16805D; } #statValue_orange { color: #C77B19; }
#primaryButton, #secondaryButton, #dangerButton, #stageButton { border-radius: 7px; padding: 9px 14px; font-weight: 600; }
#primaryButton { background: #2563EB; color: white; border: 1px solid #2563EB; }
#primaryButton:hover { background: #1D4ED8; }
#secondaryButton { color: #24547D; background: #F5F9FE; border: 1px solid #D6E3F1; }
#secondaryButton:hover { background: #EAF2FC; }
#dangerButton { color: #B43B49; background: #FFF7F7; border: 1px solid #F0D3D6; }
#dangerButton:hover { background: #FDECEE; }
#stageButton { text-align: left; color: #244A70; background: #F5F9FE; border: 1px solid #E2EAF3; padding: 8px 12px; }
#stageButton:hover { background: #EAF2FC; border-color: #B8D0EC; }
QLineEdit, QTextEdit { background: #FBFDFF; border: 1px solid #D8E3EF; border-radius: 7px; padding: 8px 10px; selection-background-color: #2563EB; }
QLineEdit:focus, QTextEdit:focus { border: 1px solid #4387ED; background: white; }
QDateEdit#deadlinePicker { color: #24425F; background: #FFFFFF; border: 1px solid #D3E1F0; border-radius: 8px; padding: 6px 32px 6px 11px; }
QDateEdit#deadlinePicker:hover { border-color: #AFC8E3; }
QDateEdit#deadlinePicker:focus { border: 1px solid #4387ED; }
QDateEdit::drop-down { subcontrol-origin: padding; subcontrol-position: top right; width: 29px; border-left: 1px solid #E5ECF4; border-top-right-radius: 7px; border-bottom-right-radius: 7px; }
QDateEdit::drop-down:hover { background: #F0F6FE; }
QDateEdit::down-arrow { width: 9px; height: 9px; }
QPushButton#clearDeadlineButton { color: #58718A; background: #F5F9FE; border: 1px solid #D6E3F1; border-radius: 7px; padding: 7px 10px; }
QPushButton#clearDeadlineButton:hover { color: #173F6B; background: #EAF2FC; }
QCalendarWidget#deadlineCalendar { color: #24425F; background: white; border: 1px solid #C9D9EA; border-radius: 10px; }
QCalendarWidget QWidget#qt_calendar_navigationbar { background: #EAF2FC; border-top-left-radius: 9px; border-top-right-radius: 9px; }
QCalendarWidget QToolButton { color: #173F6B; background: transparent; border: 0; border-radius: 5px; padding: 6px; font-weight: 650; }
QCalendarWidget QToolButton:hover { background: #DCEBFF; }
QCalendarWidget QSpinBox { color: #173F6B; background: white; border: 1px solid #D3E1F0; border-radius: 5px; padding: 3px 5px; }
QCalendarWidget QHeaderView::section { color: #71839A; background: #F7FAFE; border: 0; padding: 6px 2px; font-size: 8pt; font-weight: 700; }
QCalendarWidget QTableView { color: #334E68; background: white; alternate-background-color: #F8FAFD; selection-background-color: #2563EB; selection-color: white; border: 0; outline: 0; gridline-color: transparent; }
QCalendarWidget QTableView::item:hover { color: #173F6B; background: #EAF2FC; }
QCalendarWidget QTableView::item:selected { color: white; background: #2563EB; border-radius: 5px; }
QComboBox { color: #24425F; background: #FFFFFF; border: 1px solid #D3E1F0; border-radius: 8px; padding: 6px 32px 6px 11px; selection-background-color: #2563EB; }
QComboBox:hover { border-color: #AFC8E3; background: #FCFDFF; }
QComboBox:focus, QComboBox:on { border: 1px solid #4387ED; background: #FFFFFF; }
QComboBox::drop-down { subcontrol-origin: padding; subcontrol-position: top right; width: 31px; border: 0; }
QComboBox::down-arrow { image: none; width: 0; height: 0; }
QToolButton#comboArrow { color: #60778F; background: #F6F9FD; border: 0; border-left: 1px solid #E5ECF4; border-top-right-radius: 7px; border-bottom-right-radius: 7px; padding: 0; }
QToolButton#comboArrow:hover { background: #EAF2FC; }
QToolButton#comboArrow:pressed { background: #DCEBFF; }
QToolButton#comboArrow:disabled { background: #F7F9FC; }
QComboBox QAbstractItemView { color: #24425F; background: #FFFFFF; border: 1px solid #C9D9EA; border-radius: 11px; padding: 6px; outline: 0; selection-background-color: #DCEBFF; selection-color: #123A63; }
QComboBox QAbstractItemView::item { min-height: 30px; padding: 8px 12px; margin: 1px 2px; border-radius: 6px; }
QComboBox QAbstractItemView::item:hover { color: #173F6B; background: #F0F6FE; }
QComboBox QAbstractItemView::item:selected { color: #123A63; background: #DCEBFF; border-left: 3px solid #2563EB; }
QTableWidget { background: white; alternate-background-color: #F7FAFE; border: 0; border-radius: 6px; selection-background-color: #DCEBFF; selection-color: #123A63; }
QListWidget#compareList { color: #24425F; background: #FFFFFF; alternate-background-color: #F3F7FC; selection-background-color: #DCEBFF; selection-color: #123A63; border: 0; border-radius: 6px; outline: 0; }
QListWidget#compareList::item { padding: 10px 12px; border-bottom: 1px solid #EEF2F7; border-radius: 5px; }
QListWidget#compareList::item:hover { color: #173F5F; background: #F0F6FE; }
QListWidget#compareList::item:selected { color: #123A63; background: #DCEBFF; border-left: 3px solid #2563EB; }
QListWidget#compareList::item:selected:hover { color: #123A63; background: #E7F0FC; }
QHeaderView::section { background: #F1F6FC; color: #58718A; border: 0; border-bottom: 1px solid #DEE8F2; padding: 10px; font-size: 8pt; font-weight: 700; }
QTableWidget::item { padding: 7px; border-bottom: 1px solid #EEF2F7; }
QTabWidget::pane { border: 1px solid #E2EAF3; border-radius: 8px; background: white; }
QTabBar::tab { color: #59718A; background: #EEF4FB; border: 0; padding: 9px 15px; margin-right: 3px; }
QTabBar::tab:selected { color: white; background: #2563EB; border-radius: 5px; }
QLabel#dialogIntro { color: #71839A; padding: 4px; }
QLabel#detailTitle, QLabel#toolTitle { color: #173F6B; font-size: 13pt; font-weight: 700; }
QLabel#detailBody { color: #344F69; padding-top: 8px; }
QLabel#pathNote { color: #71839A; background: #EAF2FC; border-radius: 7px; padding: 12px; }
QStatusBar { background: #F3F7FC; color: #71839A; border-top: 1px solid #E2EAF3; }
QScrollArea { border: 0; }
QScrollBar:vertical { background: #F0F4F9; width: 11px; margin: 4px 3px 4px 3px; border-radius: 5px; }
QScrollBar::handle:vertical { background: #B9CBE0; min-height: 34px; border-radius: 4px; margin: 1px; }
QScrollBar::handle:vertical:hover { background: #82A8D3; }
QScrollBar::handle:vertical:pressed { background: #2563EB; }
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical { height: 0; border: 0; background: transparent; }
QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical { background: transparent; }
QScrollBar:horizontal { background: #F0F4F9; height: 11px; margin: 3px 4px 3px 4px; border-radius: 5px; }
QScrollBar::handle:horizontal { background: #B9CBE0; min-width: 34px; border-radius: 4px; margin: 1px; }
QScrollBar::handle:horizontal:hover { background: #82A8D3; }
QScrollBar::handle:horizontal:pressed { background: #2563EB; }
QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal { width: 0; border: 0; background: transparent; }
QScrollBar::add-page:horizontal, QScrollBar::sub-page:horizontal { background: transparent; }
QMenu { color: #24425F; background: #FFFFFF; border: 1px solid #CAD9E9; border-radius: 10px; padding: 6px; }
QMenu::item { padding: 9px 28px 9px 12px; margin: 2px; border-radius: 6px; }
QMenu::item:selected { background: #EAF2FC; color: #173F6B; }
QMenu::separator { height: 1px; background: #E8EEF5; margin: 5px 8px; }
QMenu::item:disabled { color: #AAB7C5; }
"""


def main():
    app = QApplication(sys.argv)
    app.setApplicationName(APP_NAME)
    app.setApplicationVersion(APP_VERSION)
    app.setStyle("Fusion")
    try:
        window = TrackerWindow()
    except Exception as exc:
        QMessageBox.critical(None, f"{APP_NAME} could not start", str(exc))
        return 1
    window.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
