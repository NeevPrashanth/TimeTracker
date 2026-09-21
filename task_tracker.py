"""Task Tracker - a notepad-style completed-task log plus a todo list.

Completed tasks are written to dated text files under ``task_logs``.  Todos
are persisted in ``todo_sheet.csv``.  The Reports tab groups log entries by
todo id and can save the displayed report as a dependency-free PDF.
"""

import csv
import calendar
import os
import re
import sys
import textwrap
import uuid
import tkinter as tk
from tkinter import messagebox, ttk
from datetime import date, datetime, timedelta


if getattr(sys, "frozen", False) and sys.platform == "darwin":
    # A macOS .app keeps its executable inside the bundle; store user data outside it.
    BASE_DIR = os.path.join(os.path.expanduser("~"), "Library", "Application Support", "MyTracker")
elif getattr(sys, "frozen", False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Bundled resources are extracted to _MEIPASS by PyInstaller, while user data
# should continue to be stored beside the executable.
RESOURCE_DIR = getattr(sys, "_MEIPASS", BASE_DIR)
APP_ICON = os.path.join(RESOURCE_DIR, "TimeTracker.ico")
APP_ICON_PNG = os.path.join(RESOURCE_DIR, "TimeTracker.png")

LOG_DIR = os.path.join(BASE_DIR, "task_logs")
TODO_FILE = os.path.join(BASE_DIR, "todo_sheet.csv")
TODO_FIELDS = ["id", "task", "start_date", "due_date", "priority", "status"]

PRIORITY_ORDER = {"P1": 1, "P2": 2, "P3": 3, "P4": 4}
PRIORITY_COLORS = {"P1": "red", "P2": "#e08600", "P3": "blue", "P4": "gray"}
STATUS_COLORS = {"completed": "#2e7d32", "inprogress": "#6b6b6b", "open": "#6b6b6b"}
STATUS_LABELS = {"completed": "COMPLETED", "inprogress": "IN PROGRESS", "open": "OPEN"}
UUID_LINE_RE = re.compile(
    r"^([0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}) (.*)$",
    re.DOTALL)
LOG_LINE_RE = re.compile(r"^(\d{2}:\d{2}:\d{2}) - (.*)$")
TIME_SUFFIX_RE = re.compile(r"\s*###((\d{1,2}):(\d{2}))\s*$")
RENDERED_TIME_RE = re.compile(r"Time Spent:\s*(\d{1,}):([0-5]\d)\s*$")
TIME_SPENT_INPUT_RE = re.compile(r"^(\d{1,2}):(\d{2})$")
DEFAULT_TIME_SPENT = "00:00"
DATE_INPUT_FORMATS = ("%Y-%m-%d", "%d-%m-%Y", "%d/%m/%Y", "%m/%d/%Y", "%d-%m-%y")


def today_str():
    return date.today().isoformat()


def format_date_display(value):
    if not value:
        return "-"
    try:
        return datetime.strptime(value, "%Y-%m-%d").strftime("%d-%m-%Y")
    except ValueError:
        return value


def normalize_date(value):
    """Return a user-entered date as YYYY-MM-DD, or the original if invalid."""
    value = (value or "").strip()
    if not value:
        return ""
    for fmt in DATE_INPUT_FORMATS:
        try:
            return datetime.strptime(value, fmt).strftime("%Y-%m-%d")
        except ValueError:
            continue
    return value


def days_left_banner(due_date_str):
    if not due_date_str:
        return None, None
    try:
        due = datetime.strptime(due_date_str, "%Y-%m-%d").date()
    except ValueError:
        return None, None
    delta = (due - date.today()).days
    if delta < 0:
        return f"Overdue {abs(delta)}d", "#c0392b"
    if delta == 0:
        return "Due today", "#e08600"
    if delta <= 2:
        return f"{delta}d left", "#e08600"
    return f"{delta}d left", "#2e7d32"


def strip_uuid_for_display(line):
    stripped = line.rstrip("\n")
    match = LOG_LINE_RE.match(stripped)
    if not match:
        return line
    timestamp, rest = match.groups()
    uuid_match = UUID_LINE_RE.match(rest)
    if uuid_match:
        rest = uuid_match.group(2)
    rest = render_time_suffix(rest)
    return f"{timestamp} - {rest}\n"


def normalize_time_spent(value):
    match = TIME_SPENT_INPUT_RE.match((value or "").strip())
    if not match:
        return None
    hours, minutes = match.groups()
    if not (0 <= int(hours) <= 24 and 1 <= int(minutes) <= 59):
        return None
    return f"{int(hours):02d}:{int(minutes):02d}"


def extract_time_suffix(text):
    match = TIME_SUFFIX_RE.search(text)
    if not match:
        return text, None
    return text[:match.start()], match.group(1)


def render_time_suffix(text):
    clean_text, spent = extract_time_suffix(text)
    if spent is None:
        return text
    return f"{clean_text}, Time Spent: {spent}"


def time_spent_to_minutes(hhmm):
    hours, minutes = hhmm.split(":")
    return int(hours) * 60 + int(minutes)


def minutes_to_time_spent(total_minutes):
    return f"{total_minutes // 60:02d}:{total_minutes % 60:02d}"


def report_time_summary(start_dt, end_dt, groups, unplanned_entries):
    todo_minutes = sum(
        time_spent_to_minutes(todo.get("total_time", "00:00"))
        for priority_todos in groups.values() for todo in priority_todos)
    unplanned_minutes = 0
    for _day_str, _time_str, text in unplanned_entries:
        match = RENDERED_TIME_RE.search(text)
        if match:
            unplanned_minutes += int(match.group(1)) * 60 + int(match.group(2))
    number_of_days = max(1, (end_dt - start_dt).days + 1)
    total_minutes = todo_minutes + unplanned_minutes
    average_minutes = (total_minutes + number_of_days // 2) // number_of_days
    return (minutes_to_time_spent(todo_minutes),
            minutes_to_time_spent(unplanned_minutes),
            minutes_to_time_spent(average_minutes))


def file_for_date(date_str):
    return os.path.join(LOG_DIR, f"{date_str}.txt")


def ensure_file(filepath, date_str):
    os.makedirs(LOG_DIR, exist_ok=True)
    if not os.path.exists(filepath):
        with open(filepath, "w", encoding="utf-8") as f:
            f.write(f"----- {date_str} -----\n\n")


def read_file(filepath):
    with open(filepath, "r", encoding="utf-8") as f:
        return f.read()


def read_log_records(filepath):
    """Return timestamped records, preserving any multiline continuation text."""
    records = []
    current_time, current_lines = None, []
    with open(filepath, "r", encoding="utf-8") as f:
        for raw_line in f:
            line = raw_line.rstrip("\r\n")
            match = LOG_LINE_RE.match(line)
            if match:
                if current_time is not None:
                    records.append((current_time, "\n".join(current_lines)))
                current_time, first_line = match.groups()
                current_lines = [first_line]
            elif current_time is not None and not line.startswith("-----"):
                current_lines.append(line)
    if current_time is not None:
        records.append((current_time, "\n".join(current_lines)))
    return records


def load_todos():
    if not os.path.exists(TODO_FILE):
        return []
    with open(TODO_FILE, "r", newline="", encoding="utf-8") as f:
        todos = list(csv.DictReader(f))
    for todo in todos:
        todo.setdefault("status", "open")
        if not todo.get("status"):
            todo["status"] = "open"
    return todos


def save_todos(todos):
    os.makedirs(BASE_DIR, exist_ok=True)
    with open(TODO_FILE, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=TODO_FIELDS)
        writer.writeheader()
        for todo in todos:
            writer.writerow({field: todo.get(field, "") for field in TODO_FIELDS})


def _pdf_escape(text):
    return text.replace("\\", "\\\\").replace("(", "\\(").replace(")", "\\)")


NAMED_COLORS = {
    "red": "#ff0000", "blue": "#0000ff", "gray": "#808080",
    "white": "#ffffff", "black": "#000000",
}


def _color_rgb(color):
    color = NAMED_COLORS.get(color, color).lstrip("#")
    if len(color) == 3:
        color = "".join(c * 2 for c in color)
    return tuple(int(color[i:i + 2], 16) / 255 for i in (0, 2, 4))


def _wrap_pdf_text(text, size, max_width, indent_extra=0):
    max_chars = max(20, int((max_width - indent_extra) / (size * 0.5)))
    return textwrap.wrap(text, max_chars) or [""]


class PdfBuilder:
    """Minimal dependency-free PDF page/box/text builder."""

    PAGE_WIDTH = 612
    PAGE_HEIGHT = 792
    MARGIN = 40

    def __init__(self):
        self.pages = []
        self.y = 0
        self._card_start_y = None
        self._card_start_page = None
        self._new_page()

    def _new_page(self):
        self.pages.append([])
        self.y = self.PAGE_HEIGHT - self.MARGIN

    def ensure_space(self, height):
        if self.y - height < self.MARGIN:
            self._new_page()

    def draw_rect(self, x, y, w, h, stroke=None, fill=None, line_width=1):
        ops = self.pages[-1]
        if fill:
            r, g, b = _color_rgb(fill)
            ops.append(f"{r:.3f} {g:.3f} {b:.3f} rg")
            ops.append(f"{x:.2f} {y:.2f} {w:.2f} {h:.2f} re f")
        if stroke:
            r, g, b = _color_rgb(stroke)
            ops.append(f"{r:.3f} {g:.3f} {b:.3f} RG")
            ops.append(f"{line_width} w")
            ops.append(f"{x:.2f} {y:.2f} {w:.2f} {h:.2f} re S")

    def add_line(self, text, font="F1", size=9, color="black", indent=0, leading=13):
        self.ensure_space(leading)
        r, g, b = _color_rgb(color)
        x = self.MARGIN + indent
        baseline = self.y - size
        safe = _pdf_escape(text)
        self.pages[-1].append(
            f"{r:.3f} {g:.3f} {b:.3f} rg\nBT /{font} {size} Tf {x:.2f} {baseline:.2f} Td ({safe}) Tj ET"
        )
        self.y -= leading

    def add_spacer(self, height):
        self.ensure_space(height)
        self.y -= height

    def start_card(self, min_height=44):
        self.ensure_space(min_height)
        self._card_start_y = self.y
        self._card_start_page = len(self.pages) - 1

    def end_card(self, border_color="#cccccc", padding=8):
        if self._card_start_page == len(self.pages) - 1:
            top = self._card_start_y + padding
            bottom = self.y - 4
            self.draw_rect(self.MARGIN - 6, bottom,
                           self.PAGE_WIDTH - 2 * self.MARGIN + 12,
                           top - bottom, stroke=border_color, line_width=1.2)
        self.y -= padding

    def write(self, filepath):
        num_pages = len(self.pages)
        page_obj_start = 3
        content_obj_start = 3 + num_pages
        font1_obj = 3 + 2 * num_pages
        font2_obj = font1_obj + 1
        objects = {}
        kids_refs = " ".join(f"{page_obj_start + i} 0 R" for i in range(num_pages))
        objects[1] = "<< /Type /Catalog /Pages 2 0 R >>"
        objects[2] = f"<< /Type /Pages /Kids [{kids_refs}] /Count {num_pages} >>"
        for i, ops in enumerate(self.pages):
            page_obj_num = page_obj_start + i
            content_obj_num = content_obj_start + i
            objects[page_obj_num] = (
                f"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 {self.PAGE_WIDTH} {self.PAGE_HEIGHT}] "
                f"/Resources << /Font << /F1 {font1_obj} 0 R /F2 {font2_obj} 0 R >> >> "
                f"/Contents {content_obj_num} 0 R >>"
            )
            stream_data = "\n".join(ops)
            objects[content_obj_num] = f"<< /Length {len(stream_data)} >>\nstream\n{stream_data}\nendstream"
        objects[font1_obj] = "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>"
        objects[font2_obj] = "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold >>"
        buffer = bytearray(b"%PDF-1.4\n")
        offsets = {}
        max_obj_num = font2_obj
        for obj_num in range(1, max_obj_num + 1):
            offsets[obj_num] = len(buffer)
            body = objects.get(obj_num, "")
            buffer += f"{obj_num} 0 obj\n{body}\nendobj\n".encode("latin-1", errors="replace")
        xref_offset = len(buffer)
        buffer += f"xref\n0 {max_obj_num + 1}\n".encode()
        buffer += b"0000000000 65535 f \n"
        for obj_num in range(1, max_obj_num + 1):
            buffer += f"{offsets[obj_num]:010d} 00000 n \n".encode()
        buffer += (f"trailer\n<< /Size {max_obj_num + 1} /Root 1 0 R >>\n"
                   f"startxref\n{xref_offset}\n%%EOF").encode()
        with open(filepath, "wb") as f:
            f.write(buffer)


def generate_report_pdf(filepath, title, start_dt, end_dt, groups, unplanned_entries):
    pdf = PdfBuilder()
    content_width = pdf.PAGE_WIDTH - 2 * pdf.MARGIN
    todo_total, unplanned_total, daily_average = report_time_summary(
        start_dt, end_dt, groups, unplanned_entries)
    pdf.add_line(title, font="F2", size=14, leading=22)
    pdf.add_line(f"Report: {start_dt.isoformat()} to {end_dt.isoformat()}",
                 font="F1", size=10, color="#555555", leading=22)
    pdf.add_line(f"Todo time: {todo_total}   |   Unplanned time: {unplanned_total}   |   "
                 f"Average per day: {daily_average}",
                 font="F2", size=10, color="#1f4f78", leading=20)
    pdf.add_spacer(6)
    for priority in ("P1", "P2", "P3", "P4"):
        color = PRIORITY_COLORS.get(priority, "gray")
        pdf.add_line(priority, font="F2", size=13, color=color, leading=20)
        pdf.add_spacer(8)
        todos = groups[priority]
        if not todos:
            pdf.add_line("No activity in this range.", color="gray", indent=12, leading=16)
            pdf.add_spacer(10)
            continue
        for todo in todos:
            pdf.start_card()
            status = todo["status"]
            status_color = STATUS_COLORS.get(status, "#6b6b6b")
            status_label = STATUS_LABELS.get(status, status.upper())
            pdf.add_line(f"{priority}  [{status_label}]", font="F2", size=10,
                         color=status_color, indent=8, leading=16)
            pdf.add_line(todo["task"], font="F2", size=10, indent=8, leading=16)
            pdf.add_line(f"Id: {todo['id']}  | Time spent: {todo.get('total_time', '00:00')}",
                         size=7, color="gray", indent=8, leading=14)
            for day_str, time_str, text in todo["entries"]:
                for wrapped in _wrap_pdf_text(f"{day_str} {time_str} - {text}", 9, content_width, 24):
                    pdf.add_line(wrapped, size=9, indent=16, leading=14)
            pdf.end_card(border_color=color, padding=8)
            pdf.add_spacer(8)
    pdf.add_line("Unplanned Tasks", font="F2", size=13, color="#333333", leading=20)
    pdf.add_spacer(8)
    if not unplanned_entries:
        pdf.add_line("No unplanned tasks in this range.", color="gray", indent=12, leading=16)
    else:
        pdf.start_card()
        for day_str, time_str, text in unplanned_entries:
            for wrapped in _wrap_pdf_text(f"{day_str} {time_str} - {text}", 9, content_width, 16):
                pdf.add_line(wrapped, indent=8, leading=14)
        pdf.end_card(border_color="#333333", padding=8)
    pdf.write(filepath)


class TaskTrackerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Task-Tracker-Pad")
        self._app_icon_image = None
        if os.path.exists(APP_ICON_PNG):
            try:
                self._app_icon_image = tk.PhotoImage(file=APP_ICON_PNG)
                self.root.iconphoto(True, self._app_icon_image)
            except tk.TclError:
                self._app_icon_image = None
        if os.path.exists(APP_ICON):
            try:
                self.root.iconbitmap(APP_ICON)
            except tk.TclError:
                pass
        self.root.geometry("1100x650")
        self.current_date = today_str()
        self.current_file = file_for_date(self.current_date)
        ensure_file(self.current_file, self.current_date)
        self.todos = load_todos()
        self.todo_editing_id = None
        self.todo_details_editing_id = None
        self._report_wrap_labels = []
        self._report_canvas_width = 0
        self._report_groups = None
        self._report_unplanned = []
        self._report_start_dt = None
        self._report_end_dt = None
        self._scroll_regions = []
        self._build_ui()
        self.root.bind_all("<MouseWheel>", self._on_mousewheel, add="+")
        self._load_today_into_display()
        self._refresh_todo_list()
        self.root.after(30000, self._check_date_rollover)

    def _configure_styles(self):
        style = ttk.Style(self.root)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure("TLabel", font=("Segoe UI", 9, "bold"))
        style.configure("TLabelframe.Label", font=("Segoe UI", 9, "bold"))
        style.configure("Header.TLabel", font=("Segoe UI", 11, "bold"))
        style.configure("Hint.TLabel", font=("Segoe UI", 8), foreground="gray")
        style.configure("Footer.TLabel", font=("Segoe UI", 9), foreground="#555555",
                        background="#eeeeee", padding=(8, 5))
        style.configure("Time.TEntry", font=("Segoe UI", 14, "bold"), padding=(5, 4))
        style.configure("Calendar.TButton", font=("Segoe UI Emoji", 12), padding=(4, 2))
        icon_colors = {
            "EditIcon.TButton": ("#8fd19e", "#74bb85"),
            "DeleteIcon.TButton": ("#c0392b", "#a53125"),
            "ViewIcon.TButton": ("#2878b5", "#1f6397"),
        }
        for style_name, (background, active_background) in icon_colors.items():
            style.configure(style_name, font=("Segoe UI Symbol", 12, "bold"),
                            padding=(5, 3), background=background, foreground="white",
                            borderwidth=0)
            style.map(style_name,
                      background=[("active", active_background), ("disabled", "#b8b8b8")],
                      foreground=[("disabled", "#eeeeee")])
        # A larger diagonal pencil without increasing the button padding or width.
        style.configure("EditIcon.TButton", font=("Segoe UI Symbol", 17, "bold"),
                        foreground="#173d24", padding=(2, 0))
        style.configure("TNotebook.Tab", font=("Segoe UI", 10, "bold"), padding=(16, 8))
        style.map("TNotebook.Tab", background=[("selected", "#1f6feb")],
                  foreground=[("selected", "white")])
        style.configure("TButton", font=("Segoe UI", 10, "bold"), padding=(10, 6))
        button_colors = {
            "Success.TButton": ("#2e7d32", "#256428"),
            "Warning.TButton": ("#e08600", "#c07200"),
            "Danger.TButton": ("#c0392b", "#a53125"),
            "Info.TButton": ("#1f6feb", "#1a5bc4"),
        }
        for style_name, (bg, active_bg) in button_colors.items():
            style.configure(style_name, background=bg, foreground="white",
                            font=("Segoe UI", 10, "bold"), padding=(10, 6), borderwidth=0)
            style.map(style_name, background=[("active", active_bg), ("disabled", "#a5a5a5")],
                      foreground=[("disabled", "#e6e6e6")])

    def _validate_time_part(self, proposed, maximum):
        """Allow only a partially typed, two-digit time value in range."""
        if proposed == "":
            return True
        if not proposed.isdigit() or len(proposed) > 2:
            return False
        if proposed == "00":  # Visible placeholder; never accepted for submission.
            return True
        value = int(proposed)
        return value <= maximum and (len(proposed) == 1 or value >= 1)

    def _finish_time_part(self, variable, maximum):
        value = variable.get()
        if value.isdigit() and 1 <= int(value) <= maximum:
            variable.set(f"{int(value):02d}")
        else:
            variable.set("00")

    def _make_time_inputs(self, parent, on_change):
        hour_var = tk.StringVar(value="00")
        minute_var = tk.StringVar(value="00")
        hour_check = (self.root.register(
            lambda proposed: self._validate_time_part(proposed, 24)), "%P")
        minute_check = (self.root.register(
            lambda proposed: self._validate_time_part(proposed, 59)), "%P")

        hour_entry = ttk.Entry(parent, width=4, justify="center", textvariable=hour_var,
                               validate="key", validatecommand=hour_check, style="Time.TEntry")
        hour_entry.pack(side="left", padx=(6, 2))
        ttk.Label(parent, text=":", font=("Segoe UI", 14, "bold")).pack(side="left")
        minute_entry = ttk.Entry(parent, width=4, justify="center", textvariable=minute_var,
                                 validate="key", validatecommand=minute_check, style="Time.TEntry")
        minute_entry.pack(side="left", padx=(2, 0))
        hour_entry.bind("<FocusIn>", lambda _e: hour_entry.select_range(0, "end"))
        minute_entry.bind("<FocusIn>", lambda _e: minute_entry.select_range(0, "end"))
        hour_entry.bind("<FocusOut>", lambda _e: self._finish_time_part(hour_var, 24))
        minute_entry.bind("<FocusOut>", lambda _e: self._finish_time_part(minute_var, 59))
        hour_var.trace_add("write", on_change)
        minute_var.trace_add("write", on_change)
        return hour_var, minute_var

    @staticmethod
    def _time_from_parts(hour_var, minute_var):
        return normalize_time_spent(f"{hour_var.get()}:{minute_var.get()}")

    @staticmethod
    def _validate_date_typing(proposed):
        """Accept only a partial YYYY-MM-DD value while the user is typing."""
        if len(proposed) > 10:
            return False
        for index, character in enumerate(proposed):
            if index in (4, 7):
                if character != "-":
                    return False
            elif not character.isdigit():
                return False
        return True

    @staticmethod
    def _parse_iso_date(value):
        if len(value) != 10:
            return None
        try:
            return datetime.strptime(value, "%Y-%m-%d").date()
        except ValueError:
            return None

    def _open_date_picker(self, entry):
        """Show a small, dependency-free calendar and write the selection to entry."""
        selected = self._parse_iso_date(entry.get().strip()) or date.today()
        picker = tk.Toplevel(self.root)
        picker.title("Choose date")
        picker.resizable(False, False)
        picker.transient(self.root)

        content = ttk.Frame(picker, padding=8)
        content.pack(fill="both", expand=True)
        header = ttk.Frame(content)
        header.pack(fill="x", pady=(0, 6))
        header.columnconfigure(1, weight=1)
        month_label = ttk.Label(header, anchor="center", style="Header.TLabel")
        month_label.grid(row=0, column=1, sticky="ew")
        days_frame = ttk.Frame(content)
        days_frame.pack()
        shown = [selected.year, selected.month]

        def choose(day_number):
            chosen = date(shown[0], shown[1], day_number)
            entry.delete(0, "end")
            entry.insert(0, chosen.isoformat())
            self._update_add_todo_button_state()
            picker.destroy()

        def draw_calendar():
            for widget in days_frame.winfo_children():
                widget.destroy()
            month_label.configure(text=f"{calendar.month_name[shown[1]]} {shown[0]}")
            for column, name in enumerate(("Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun")):
                ttk.Label(days_frame, text=name, width=4, anchor="center").grid(
                    row=0, column=column, padx=1, pady=1)
            for row, week in enumerate(calendar.Calendar(firstweekday=0).monthdayscalendar(
                    shown[0], shown[1]), start=1):
                for column, day_number in enumerate(week):
                    if day_number:
                        ttk.Button(days_frame, text=str(day_number), width=3,
                                   command=lambda day=day_number: choose(day)).grid(
                                       row=row, column=column, padx=1, pady=1)

        def move_month(offset):
            month_index = shown[0] * 12 + shown[1] - 1 + offset
            shown[0], zero_based_month = divmod(month_index, 12)
            shown[1] = zero_based_month + 1
            draw_calendar()

        ttk.Button(header, text="<", width=3, command=lambda: move_month(-1)).grid(row=0, column=0)
        ttk.Button(header, text=">", width=3, command=lambda: move_month(1)).grid(row=0, column=2)
        draw_calendar()
        picker.update_idletasks()
        x = entry.winfo_rootx()
        y = entry.winfo_rooty() + entry.winfo_height()
        picker.geometry(f"+{x}+{y}")
        picker.grab_set()

    def _build_ui(self):
        self._configure_styles()
        ttk.Label(self.root, text="@copywrite NeevinfraLtd 2026", anchor="center",
                  style="Footer.TLabel").pack(side="bottom", fill="x")
        paned = ttk.PanedWindow(self.root, orient="horizontal")
        paned.pack(fill="both", expand=True)
        left_frame, right_frame = ttk.Frame(paned), ttk.Frame(paned)
        paned.add(left_frame, weight=1)
        paned.add(right_frame, weight=1)
        self._build_left_pane(left_frame)
        self._build_right_pane(right_frame)

    @staticmethod
    def _is_descendant(widget, ancestor):
        while widget is not None:
            if widget == ancestor:
                return True
            widget = widget.master
        return False

    def _on_mousewheel(self, event):
        """Scroll the panel currently underneath the mouse pointer."""
        pointed_widget = self.root.winfo_containing(
            self.root.winfo_pointerx(), self.root.winfo_pointery())
        if pointed_widget is None or not event.delta:
            return None
        for region, scroll_target in reversed(self._scroll_regions):
            if self._is_descendant(pointed_widget, region):
                direction = -1 if event.delta > 0 else 1
                steps = max(1, abs(event.delta) // 120)
                scroll_target.yview_scroll(direction * steps, "units")
                return "break"
        return None

    def _build_left_pane(self, parent):
        mono_font = ("Consolas", 11)
        top_frame = ttk.Frame(parent, padding=8)
        top_frame.pack(fill="x")
        ttk.Label(top_frame, text="Enter completed task:").pack(anchor="w")
        input_row = ttk.Frame(top_frame)
        input_row.pack(fill="x", pady=(4, 0))
        self.input_text = tk.Text(input_row, height=3, font=mono_font, bg="white", wrap="word", undo=True)
        self.input_text.pack(side="left", fill="x", expand=True)
        self.input_text.bind("<Return>", self._on_return)
        self.input_text.focus_set()
        self.add_task_button = ttk.Button(input_row, text="Add Task", command=self.add_task,
                                          style="Success.TButton")
        self.add_task_button.pack(side="left", padx=(6, 0), fill="y")
        time_row = ttk.Frame(top_frame)
        time_row.pack(fill="x", pady=(4, 0))
        ttk.Label(time_row, text="Add spent on unplanned tasks:").pack(side="left")
        self.time_hour_var, self.time_minute_var = self._make_time_inputs(
            time_row, self._update_add_task_button_state)
        ttk.Label(time_row, text="HH (00-24)   MM (01-59)", style="Hint.TLabel").pack(
            side="left", padx=(8, 0))
        ttk.Label(top_frame, text="(Press Enter to add, Shift+Enter for a new line)",
                  style="Hint.TLabel").pack(anchor="w", pady=(2, 0))
        self._update_add_task_button_state()
        bottom_frame = ttk.Frame(parent, padding=8)
        bottom_frame.pack(fill="both", expand=True)
        ttk.Label(bottom_frame, text="Today's completed tasks", style="Header.TLabel").pack(anchor="w")
        display_container = ttk.Frame(bottom_frame)
        display_container.pack(fill="both", expand=True, pady=(4, 0))
        scrollbar = ttk.Scrollbar(display_container)
        scrollbar.pack(side="right", fill="y")
        self.display_text = tk.Text(display_container, font=mono_font, bg="white", wrap="word",
                                    yscrollcommand=scrollbar.set, state="disabled")
        self.display_text.pack(side="left", fill="both", expand=True)
        scrollbar.config(command=self.display_text.yview)
        self._scroll_regions.append((display_container, self.display_text))

    def _on_return(self, event):
        if event.state & 0x0001:
            return None
        self.add_task()
        return "break"

    def _load_today_into_display(self):
        lines = read_file(self.current_file).splitlines(keepends=True)
        display_content = "".join(strip_uuid_for_display(line) for line in lines)
        self.display_text.configure(state="normal")
        self.display_text.delete("1.0", "end")
        self.display_text.insert("end", display_content)
        self.display_text.see("end")
        self.display_text.configure(state="disabled")

    def add_task(self):
        task = self.input_text.get("1.0", "end").strip()
        time_spent = self._time_from_parts(self.time_hour_var, self.time_minute_var)
        if not task or not time_spent or time_spent == DEFAULT_TIME_SPENT:
            return
        self._log_completed(f"{task} ###{time_spent}")
        self.input_text.delete("1.0", "end")
        self.time_hour_var.set("00")
        self.time_minute_var.set("00")

    def _update_add_task_button_state(self, *_args):
        time_spent = self._time_from_parts(self.time_hour_var, self.time_minute_var)
        valid = bool(time_spent) and time_spent != DEFAULT_TIME_SPENT
        self.add_task_button.configure(state="normal" if valid else "disabled")

    def _log_completed(self, description):
        timestamp = datetime.now().strftime("%H:%M:%S")
        line = f"{timestamp} - {description}\n"
        with open(self.current_file, "a", encoding="utf-8") as f:
            f.write(line)
        self.display_text.configure(state="normal")
        self.display_text.insert("end", strip_uuid_for_display(line))
        self.display_text.see("end")
        self.display_text.configure(state="disabled")

    def _check_date_rollover(self):
        new_date = today_str()
        if new_date != self.current_date:
            self.current_date = new_date
            self.current_file = file_for_date(self.current_date)
            ensure_file(self.current_file, self.current_date)
            self._load_today_into_display()
        self.root.after(30000, self._check_date_rollover)

    def _build_right_pane(self, parent):
        notebook = ttk.Notebook(parent)
        notebook.pack(fill="both", expand=True)
        newtodo_tab, reports_tab = ttk.Frame(notebook), ttk.Frame(notebook)
        notebook.add(newtodo_tab, text="New Todo")
        notebook.add(reports_tab, text="Reports")
        self._build_newtodo_tab(newtodo_tab)
        self._build_reports_tab(reports_tab)

    def _build_newtodo_tab(self, parent):
        form_frame = ttk.LabelFrame(parent, text="New Todo", padding=8)
        form_frame.pack(fill="x", padx=8, pady=8)
        form_frame.columnconfigure(1, weight=1)
        ttk.Label(form_frame, text="Task:").grid(row=0, column=0, sticky="w", pady=2)
        self.todo_task_entry = ttk.Entry(form_frame)
        self.todo_task_entry.grid(row=0, column=1, columnspan=3, sticky="ew", pady=2)
        self.todo_task_entry.bind("<Return>", lambda e: self.add_todo())
        self.todo_task_entry.bind("<KeyRelease>", self._update_add_todo_button_state)
        ttk.Label(form_frame, text="Start Date (YYYY-MM-DD):").grid(row=1, column=0, sticky="w", pady=2)
        date_check = (self.root.register(self._validate_date_typing), "%P")
        start_date_frame = ttk.Frame(form_frame)
        start_date_frame.grid(row=1, column=1, sticky="w", pady=2)
        self.todo_start_entry = ttk.Entry(start_date_frame, width=12, validate="key",
                                          validatecommand=date_check)
        self.todo_start_entry.pack(side="left")
        ttk.Button(start_date_frame, text="📅", width=3, style="Calendar.TButton",
                   command=lambda: self._open_date_picker(self.todo_start_entry)).pack(
                       side="left", padx=(4, 0))
        self.todo_start_entry.insert(0, today_str())
        ttk.Label(form_frame, text="Due Date (YYYY-MM-DD):").grid(row=2, column=0, sticky="w", pady=2)
        due_date_frame = ttk.Frame(form_frame)
        due_date_frame.grid(row=2, column=1, sticky="w", pady=2)
        self.todo_due_entry = ttk.Entry(due_date_frame, width=12, validate="key",
                                        validatecommand=date_check)
        self.todo_due_entry.pack(side="left")
        ttk.Button(due_date_frame, text="📅", width=3, style="Calendar.TButton",
                   command=lambda: self._open_date_picker(self.todo_due_entry)).pack(
                       side="left", padx=(4, 0))
        self.todo_due_entry.insert(0, today_str())
        self.todo_start_entry.bind("<KeyRelease>", self._update_add_todo_button_state)
        self.todo_due_entry.bind("<KeyRelease>", self._update_add_todo_button_state)
        ttk.Label(form_frame, text="Priority:").grid(row=2, column=2, sticky="w", padx=(8, 4))
        self.todo_priority_var = tk.StringVar(value="P4")
        ttk.Combobox(form_frame, textvariable=self.todo_priority_var,
                     values=["P1", "P2", "P3", "P4"], width=5, state="readonly").grid(
                         row=2, column=3, sticky="w", pady=2)
        self.add_todo_button = ttk.Button(form_frame, text="Add Todo", command=self.add_todo,
                                          style="Success.TButton")
        self.add_todo_button.grid(row=3, column=0, columnspan=4, sticky="w", pady=(6, 0))
        todo_header = ttk.Frame(parent, padding=(8, 0))
        todo_header.pack(fill="x")
        ttk.Label(todo_header, text="Open Todo Tasks (by priority)",
                  style="Header.TLabel").pack(side="left")
        self.todo_search_hint = "by description"
        self.todo_search_var = tk.StringVar(value=self.todo_search_hint)
        ttk.Label(todo_header, text="Search:").pack(side="left", padx=(16, 4))
        todo_search_entry = ttk.Entry(todo_header, textvariable=self.todo_search_var, width=24,
                                      foreground="gray")
        todo_search_entry.pack(side="left", fill="x", expand=True)
        def clear_search_hint(_event):
            if self.todo_search_var.get() == self.todo_search_hint:
                self.todo_search_var.set("")
                todo_search_entry.configure(foreground="black")
        def restore_search_hint(_event):
            if not self.todo_search_var.get().strip():
                self.todo_search_var.set(self.todo_search_hint)
                todo_search_entry.configure(foreground="gray")
        todo_search_entry.bind("<FocusIn>", clear_search_hint)
        todo_search_entry.bind("<FocusOut>", restore_search_hint)
        self.todo_search_var.trace_add("write", self._refresh_todo_list)
        list_container = ttk.Frame(parent)
        list_container.pack(fill="both", expand=True, padx=8, pady=(4, 8))
        canvas = tk.Canvas(list_container, bg="white", highlightthickness=0)
        list_scrollbar = ttk.Scrollbar(list_container, orient="vertical", command=canvas.yview)
        self.todo_inner_frame = ttk.Frame(canvas)
        self.todo_inner_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        window_id = canvas.create_window((0, 0), window=self.todo_inner_frame, anchor="nw")
        canvas.bind("<Configure>", lambda e: canvas.itemconfig(window_id, width=e.width))
        canvas.configure(yscrollcommand=list_scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        list_scrollbar.pack(side="right", fill="y")
        self._scroll_regions.append((list_container, canvas))
        self._update_add_todo_button_state()

    def _update_add_todo_button_state(self, *_args):
        task_filled = bool(self.todo_task_entry.get().strip())
        start_date = self._parse_iso_date(self.todo_start_entry.get().strip())
        due_date = self._parse_iso_date(self.todo_due_entry.get().strip())
        dates_valid = start_date is not None and due_date is not None and start_date <= due_date
        state = "normal" if task_filled and dates_valid else "disabled"
        self.add_todo_button.configure(state=state)

    def add_todo(self):
        task = self.todo_task_entry.get().strip()
        start_date = normalize_date(self.todo_start_entry.get().strip())
        due_date = normalize_date(self.todo_due_entry.get().strip())
        priority = self.todo_priority_var.get().strip().upper()
        if not task or not start_date or not due_date or priority not in PRIORITY_ORDER:
            return
        try:
            start_dt = datetime.strptime(start_date, "%Y-%m-%d").date()
            due_dt = datetime.strptime(due_date, "%Y-%m-%d").date()
        except ValueError:
            messagebox.showerror("Invalid Date", "Enter valid dates in YYYY-MM-DD format.")
            return
        if start_dt > due_dt:
            messagebox.showerror("Invalid Date Range",
                                 "Start Date must be earlier than or equal to Due Date.")
            return
        self.todos.append({"id": str(uuid.uuid4()), "task": task, "start_date": start_date,
                           "due_date": due_date, "priority": priority, "status": "open"})
        save_todos(self.todos)
        self._refresh_todo_list()
        self.todo_task_entry.delete(0, "end")
        self.todo_start_entry.delete(0, "end")
        self.todo_start_entry.insert(0, today_str())
        self.todo_due_entry.delete(0, "end")
        self.todo_due_entry.insert(0, today_str())
        self.todo_priority_var.set("P4")
        self._update_add_todo_button_state()
        self.todo_task_entry.focus_set()

    def complete_todo(self, todo_id):
        todo = next((t for t in self.todos if t["id"] == todo_id), None)
        if todo is None:
            return
        priority = todo.get("priority", "P4").lower()
        summary = (f"{todo_id} {priority} startdate: {format_date_display(todo.get('start_date'))} "
                   f"end date: {format_date_display(today_str())} due date: "
                   f"{format_date_display(todo.get('due_date'))} {todo.get('task', '')}")
        todo["status"] = "completed"
        save_todos(self.todos)
        self._log_completed(summary)
        if self.todo_editing_id == todo_id:
            self.todo_editing_id = None
        if self.todo_details_editing_id == todo_id:
            self.todo_details_editing_id = None
        self._refresh_todo_list()

    def toggle_update_mode(self, todo_id):
        self.todo_details_editing_id = None
        self.todo_editing_id = None if self.todo_editing_id == todo_id else todo_id
        self._refresh_todo_list()

    def cancel_update(self):
        self.todo_editing_id = None
        self._refresh_todo_list()

    def _get_todo_updates(self, todo_id):
        updates = []
        if os.path.isdir(LOG_DIR):
            for filename in sorted(os.listdir(LOG_DIR)):
                if not re.fullmatch(r"\d{4}-\d{2}-\d{2}\.txt", filename):
                    continue
                day_str = filename[:-4]
                for time_str, rest in read_log_records(os.path.join(LOG_DIR, filename)):
                    match = UUID_LINE_RE.match(rest)
                    if match and match.group(1) == todo_id:
                        text, spent = extract_time_suffix(match.group(2))
                        if spent:
                            text = f"{text}, Time Spent: {spent}"
                        updates.append((day_str, time_str, text))
        return updates

    def show_todo_updates(self, todo_id):
        todo = next((item for item in self.todos if item["id"] == todo_id), None)
        updates = self._get_todo_updates(todo_id)

        window = tk.Toplevel(self.root)
        window.title("Todo update history")
        window.geometry("700x420")
        window.transient(self.root)
        frame = ttk.Frame(window, padding=10)
        frame.pack(fill="both", expand=True)
        ttk.Label(frame, text=todo.get("task", "Todo updates") if todo else "Todo updates",
                  style="Header.TLabel", wraplength=650).pack(anchor="w", pady=(0, 8))
        text_box = tk.Text(frame, font=("Segoe UI", 10), wrap="word", state="normal")
        scrollbar = ttk.Scrollbar(frame, command=text_box.yview)
        text_box.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        text_box.pack(side="left", fill="both", expand=True)
        if updates:
            for day_str, time_str, text in updates:
                text_box.insert("end", f"{day_str} {time_str}\n{text}\n\n")
        else:
            text_box.insert("end", "No updates recorded for this task yet.")
        text_box.configure(state="disabled")

    def toggle_todo_details(self, todo_id):
        self.todo_editing_id = None
        self.todo_details_editing_id = (
            None if self.todo_details_editing_id == todo_id else todo_id)
        self._refresh_todo_list()

    def cancel_todo_details(self):
        self.todo_details_editing_id = None
        self._refresh_todo_list()

    def save_todo_details(self, todo_id):
        todo = next((item for item in self.todos if item["id"] == todo_id), None)
        if todo is None:
            return
        description = self.todo_edit_task_var.get().strip()
        due_date = self._parse_iso_date(self.todo_edit_due_var.get().strip())
        start_date = self._parse_iso_date(todo.get("start_date", ""))
        if not description:
            messagebox.showerror("Invalid Task", "Task description cannot be empty.")
            return
        if due_date is None:
            messagebox.showerror("Invalid Date", "Enter a valid due date in YYYY-MM-DD format.")
            return
        if start_date is not None and due_date < start_date:
            messagebox.showerror("Invalid Date Range",
                                 "Due Date must be later than or equal to Start Date.")
            return
        todo["task"] = description
        todo["due_date"] = due_date.isoformat()
        save_todos(self.todos)
        self.todo_details_editing_id = None
        self._refresh_todo_list()

    def delete_todo(self, todo_id):
        todo = next((item for item in self.todos if item["id"] == todo_id), None)
        if todo is None:
            return
        if not messagebox.askyesno(
                "Delete Task",
                f"Are you sure you want to delete this task?\n\n{todo.get('task', '')}"):
            return
        self.todos = [item for item in self.todos if item["id"] != todo_id]
        save_todos(self.todos)
        if self.todo_editing_id == todo_id:
            self.todo_editing_id = None
        if self.todo_details_editing_id == todo_id:
            self.todo_details_editing_id = None
        self._refresh_todo_list()

    def submit_update(self, todo_id):
        todo = next((t for t in self.todos if t["id"] == todo_id), None)
        if todo is None:
            return
        comment = self.todo_update_text.get("1.0", "end").strip()
        time_spent = self._time_from_parts(self.todo_time_hour_var, self.todo_time_minute_var)
        if not comment or not time_spent or time_spent == DEFAULT_TIME_SPENT:
            return
        self._log_completed(f"{todo_id} {comment} ###{time_spent}")
        if todo.get("status") == "open":
            todo["status"] = "inprogress"
            save_todos(self.todos)
        self.todo_editing_id = None
        self._refresh_todo_list()

    def _update_submit_update_button_state(self, *_args):
        time_spent = self._time_from_parts(self.todo_time_hour_var, self.todo_time_minute_var)
        valid = bool(time_spent) and time_spent != DEFAULT_TIME_SPENT
        self.submit_update_button.configure(state="normal" if valid else "disabled")

    def _refresh_todo_list(self, *_args):
        for widget in self.todo_inner_frame.winfo_children():
            widget.destroy()
        raw_search = self.todo_search_var.get().strip()
        search_text = "" if raw_search == self.todo_search_hint else raw_search.casefold()
        open_todos = [
            todo for todo in self.todos
            if todo.get("status", "open") != "completed"
            and (not search_text or search_text in todo.get("task", "").casefold())
        ]
        sorted_todos = sorted(open_todos, key=lambda t: PRIORITY_ORDER.get(t.get("priority", "P4"), 4))
        if not sorted_todos:
            empty_message = "No matching open tasks." if search_text else "No open todo tasks."
            ttk.Label(self.todo_inner_frame, text=empty_message, foreground="gray").pack(
                anchor="w", padx=6, pady=6)
            return
        for todo in sorted_todos:
            priority = todo.get("priority", "P4")
            color = PRIORITY_COLORS.get(priority, "gray")
            banner_text, banner_color = days_left_banner(todo.get("due_date"))
            is_editing = self.todo_editing_id == todo["id"]
            is_details_editing = self.todo_details_editing_id == todo["id"]
            updates = self._get_todo_updates(todo["id"])
            row = tk.Frame(self.todo_inner_frame, bg="white", bd=1, relief="solid")
            row.pack(fill="x", padx=4, pady=3)
            row.columnconfigure(3, weight=1)
            tk.Label(row, text=priority, fg=color, bg="white", font=("Segoe UI", 10, "bold")).grid(
                row=0, column=0, sticky="w", padx=(6, 4), pady=4)
            tk.Label(row, text=f"Due: {todo.get('due_date') or '-'}", bg="white",
                     font=("Segoe UI", 9), width=14).grid(row=0, column=1, sticky="w", pady=4)
            if banner_text:
                tk.Label(row, text=banner_text, bg=banner_color, fg="white",
                         font=("Segoe UI", 8, "bold"), padx=6, pady=1).grid(
                             row=0, column=2, sticky="w", pady=4)
            if todo.get("status") == "inprogress":
                tk.Label(row, text="In Progress", bg="white", fg="#6b6b6b",
                         font=("Segoe UI", 8, "italic")).grid(row=1, column=0, columnspan=2,
                                                               sticky="w", padx=(6, 4))
            if is_details_editing:
                edit_frame = tk.Frame(row, bg="white")
                edit_frame.grid(row=0, column=3, rowspan=2, sticky="ew", padx=4, pady=4)
                edit_frame.columnconfigure(1, weight=1)
                tk.Label(edit_frame, text="Description:", bg="white",
                         font=("Segoe UI", 9, "bold")).grid(row=0, column=0, sticky="w", padx=(0, 4))
                self.todo_edit_task_var = tk.StringVar(value=todo.get("task", ""))
                ttk.Entry(edit_frame, textvariable=self.todo_edit_task_var).grid(
                    row=0, column=1, columnspan=2, sticky="ew")
                tk.Label(edit_frame, text="Due Date:", bg="white",
                         font=("Segoe UI", 9, "bold")).grid(row=1, column=0, sticky="w",
                                                              padx=(0, 4), pady=(4, 0))
                self.todo_edit_due_var = tk.StringVar(value=todo.get("due_date", ""))
                date_check = (self.root.register(self._validate_date_typing), "%P")
                due_entry = ttk.Entry(edit_frame, width=12, textvariable=self.todo_edit_due_var,
                                      validate="key", validatecommand=date_check)
                due_entry.grid(row=1, column=1, sticky="w", pady=(4, 0))
                ttk.Button(edit_frame, text="\U0001f4c5", width=3, style="Calendar.TButton",
                           command=lambda entry=due_entry: self._open_date_picker(entry)).grid(
                               row=1, column=2, sticky="w", padx=(4, 0), pady=(4, 0))
                actions = tk.Frame(edit_frame, bg="white")
                actions.grid(row=2, column=0, columnspan=3, sticky="w", pady=(6, 0))
                ttk.Button(actions, text="Save", style="Success.TButton",
                           command=lambda tid=todo["id"]: self.save_todo_details(tid)).pack(side="left")
                ttk.Button(actions, text="Cancel", style="Danger.TButton",
                           command=self.cancel_todo_details).pack(side="left", padx=(4, 0))
            elif is_editing:
                edit_frame = tk.Frame(row, bg="white")
                edit_frame.grid(row=0, column=3, rowspan=2, sticky="ew", padx=(4, 4), pady=(4, 4))
                edit_frame.columnconfigure(0, weight=1)
                tk.Label(edit_frame, text=todo.get("task", ""), bg="white",
                         font=("Segoe UI", 10), anchor="w", justify="left", wraplength=260).pack(fill="x")
                self.todo_update_text = tk.Text(edit_frame, height=5, font=("Segoe UI", 10), wrap="word", bg="white")
                self.todo_update_text.pack(fill="x", pady=(4, 0))
                self.todo_update_text.focus_set()
                time_row = tk.Frame(edit_frame, bg="white")
                time_row.pack(fill="x", pady=(4, 0))
                tk.Label(time_row, text="Time spent:", bg="white", font=("Segoe UI", 9)).pack(side="left")
                self.todo_time_hour_var, self.todo_time_minute_var = self._make_time_inputs(
                    time_row, self._update_submit_update_button_state)
                tk.Label(time_row, text="HH 00-24 / MM 01-59", bg="white", fg="gray",
                         font=("Segoe UI", 8)).pack(side="left", padx=(8, 0))
                update_btn_row = tk.Frame(edit_frame, bg="white")
                update_btn_row.pack(fill="x", pady=(4, 0))
                self.submit_update_button = ttk.Button(update_btn_row, text="Submit Update",
                    command=lambda tid=todo["id"]: self.submit_update(tid), style="Info.TButton")
                self.submit_update_button.pack(side="left")
                ttk.Button(update_btn_row, text="Cancel", command=self.cancel_update,
                           style="Danger.TButton").pack(side="left", padx=(4, 0))
                self._update_submit_update_button_state()
            else:
                tk.Label(row, text=todo.get("task", ""), bg="white", font=("Segoe UI", 10),
                         anchor="w", justify="left", wraplength=260).grid(
                             row=0, column=3, sticky="ew", padx=(4, 4), pady=(4, 0))
                tk.Label(row, text=f"Start: {todo.get('start_date') or '-'}", bg="white",
                         font=("Segoe UI", 8), fg="gray", anchor="w").grid(
                             row=1, column=3, sticky="w", padx=(4, 4), pady=(0, 4))
                btns = tk.Frame(row, bg="white")
                btns.grid(row=0, column=4, rowspan=2, sticky="ne", padx=(4, 6), pady=4)
                icon_row = tk.Frame(btns, bg="white")
                icon_row.pack(fill="x", pady=(0, 3))
                ttk.Button(icon_row, text="\u270f", width=3, style="EditIcon.TButton",
                           command=lambda tid=todo["id"]: self.toggle_todo_details(tid)).pack(
                               side="left", padx=(0, 2))
                ttk.Button(icon_row, text="\U0001f5d1", width=3, style="DeleteIcon.TButton",
                           command=lambda tid=todo["id"]: self.delete_todo(tid)).pack(
                               side="left", padx=(0, 2))
                ttk.Button(icon_row, text="\U0001f441", width=3, style="ViewIcon.TButton",
                           state="normal" if updates else "disabled",
                           command=lambda tid=todo["id"]: self.show_todo_updates(tid)).pack(side="left")
                ttk.Button(btns, text="Update", width=9,
                           command=lambda tid=todo["id"]: self.toggle_update_mode(tid),
                           style="Warning.TButton").pack(fill="x", pady=(0, 2))
                ttk.Button(btns, text="Complete", width=9,
                           command=lambda tid=todo["id"]: self.complete_todo(tid),
                           style="Success.TButton").pack(fill="x")

    def _build_reports_tab(self, parent):
        form = ttk.Frame(parent, padding=8)
        form.pack(fill="x")
        ttk.Label(form, text="Start Date (YYYY-MM-DD):").grid(row=0, column=0, sticky="w")
        self.report_start_entry = ttk.Entry(form, width=14)
        self.report_start_entry.grid(row=0, column=1, sticky="w", padx=(4, 12))
        self.report_start_entry.insert(0, today_str())
        ttk.Label(form, text="End Date (YYYY-MM-DD):").grid(row=0, column=2, sticky="w")
        self.report_end_entry = ttk.Entry(form, width=14)
        self.report_end_entry.grid(row=0, column=3, sticky="w", padx=(4, 12))
        self.report_end_entry.insert(0, today_str())
        ttk.Button(form, text="Generate Report", command=self.generate_report,
                   style="Info.TButton").grid(row=0, column=4, padx=(0, 4))
        self.download_pdf_button = ttk.Button(form, text="Download as PDF",
                                              command=self.download_report_pdf,
                                              state="disabled", style="Info.TButton")
        self.download_pdf_button.grid(row=0, column=5)
        report_container = ttk.Frame(parent)
        report_container.pack(fill="both", expand=True, padx=8, pady=(4, 8))
        canvas = tk.Canvas(report_container, bg="#f0f0f0", highlightthickness=0)
        report_scrollbar = ttk.Scrollbar(report_container, orient="vertical", command=canvas.yview)
        self.report_inner_frame = ttk.Frame(canvas)
        self.report_inner_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        window_id = canvas.create_window((0, 0), window=self.report_inner_frame, anchor="nw")
        def _on_canvas_configure(event):
            canvas.itemconfig(window_id, width=event.width)
            self._update_report_wraplengths(event.width)
        canvas.bind("<Configure>", _on_canvas_configure)
        canvas.configure(yscrollcommand=report_scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        report_scrollbar.pack(side="right", fill="y")
        self._scroll_regions.append((report_container, canvas))
        ttk.Label(self.report_inner_frame, text="Pick a date range and click Generate Report.",
                  foreground="gray").pack(anchor="w", padx=6, pady=6)

    def _build_report_data(self, start_dt, end_dt):
        todos_by_id = {t["id"]: t for t in self.todos}
        uuid_entries, minutes_by_todo, unplanned_entries = {}, {}, []
        current = start_dt
        while current <= end_dt:
            day_str = current.isoformat()
            filepath = file_for_date(day_str)
            if os.path.exists(filepath):
                for time_str, rest in read_log_records(filepath):
                    uuid_match = UUID_LINE_RE.match(rest)
                    if uuid_match:
                        todo_id, text = uuid_match.groups()
                        clean_text, spent = extract_time_suffix(text)
                        display_text = f"{clean_text}, Time Spent: {spent}" if spent else clean_text
                        uuid_entries.setdefault(todo_id, []).append((day_str, time_str, display_text))
                        if spent:
                            minutes_by_todo[todo_id] = minutes_by_todo.get(todo_id, 0) + time_spent_to_minutes(spent)
                    else:
                        unplanned_entries.append((day_str, time_str, render_time_suffix(rest)))
            current += timedelta(days=1)
        groups = {p: [] for p in ("P1", "P2", "P3", "P4")}
        for todo_id, entries in uuid_entries.items():
            todo = todos_by_id.get(todo_id, {})
            priority = todo.get("priority", "P4")
            if priority not in groups:
                priority = "P4"
            groups[priority].append({"id": todo_id, "task": todo.get("task", "(deleted todo)"),
                                     "status": todo.get("status", "unknown"), "entries": entries,
                                     "total_time": minutes_to_time_spent(minutes_by_todo.get(todo_id, 0))})
        return groups, unplanned_entries

    def generate_report(self):
        start_str = normalize_date(self.report_start_entry.get().strip()) or today_str()
        end_str = normalize_date(self.report_end_entry.get().strip()) or today_str()
        try:
            start_dt = datetime.strptime(start_str, "%Y-%m-%d").date()
            end_dt = datetime.strptime(end_str, "%Y-%m-%d").date()
        except ValueError:
            self._render_report_error("Invalid start/end date. Use YYYY-MM-DD.")
            return
        if start_dt > end_dt:
            start_dt, end_dt = end_dt, start_dt
        groups, unplanned_entries = self._build_report_data(start_dt, end_dt)
        self._report_groups, self._report_unplanned = groups, unplanned_entries
        self._report_start_dt, self._report_end_dt = start_dt, end_dt
        self._render_report(start_dt, end_dt, groups, unplanned_entries)
        has_data = any(groups[p] for p in groups) or bool(unplanned_entries)
        self.download_pdf_button.configure(state="normal" if has_data else "disabled")

    def _register_wrap_label(self, label, padding):
        self._report_wrap_labels.append((label, padding))
        if self._report_canvas_width:
            label.configure(wraplength=max(150, self._report_canvas_width - padding))

    def _update_report_wraplengths(self, width):
        self._report_canvas_width = width
        for label, padding in self._report_wrap_labels:
            try:
                label.configure(wraplength=max(150, width - padding))
            except tk.TclError:
                pass

    def _render_report_error(self, message):
        self._report_wrap_labels = []
        for widget in self.report_inner_frame.winfo_children():
            widget.destroy()
        tk.Label(self.report_inner_frame, text=message, fg="#c0392b", bg="#f0f0f0",
                 font=("Segoe UI", 10, "bold")).pack(anchor="w", padx=6, pady=6)
        self._report_groups = None
        self._report_unplanned = []
        self.download_pdf_button.configure(state="disabled")

    def _render_report(self, start_dt, end_dt, groups, unplanned_entries):
        self._report_wrap_labels = []
        for widget in self.report_inner_frame.winfo_children():
            widget.destroy()
        tk.Label(self.report_inner_frame, text=f"Report: {start_dt.isoformat()} to {end_dt.isoformat()}",
                 bg="#f0f0f0", font=("Segoe UI", 11, "bold")).pack(anchor="w", padx=6, pady=(6, 8))
        todo_total, unplanned_total, daily_average = report_time_summary(
            start_dt, end_dt, groups, unplanned_entries)
        summary = tk.Frame(self.report_inner_frame, bg="#e8f1f8", bd=1, relief="solid")
        summary.pack(fill="x", padx=6, pady=(0, 8))
        summary.columnconfigure((0, 1, 2), weight=1, uniform="summary")
        for column, (label_text, value, color) in enumerate((
                ("Todo Tasks", todo_total, "#2e7d32"),
                ("Unplanned Tasks", unplanned_total, "#d07800"),
                ("Average per Day", daily_average, "#1f6feb"))):
            box = tk.Frame(summary, bg="#e8f1f8")
            box.grid(row=0, column=column, sticky="ew", padx=8, pady=7)
            tk.Label(box, text=label_text, bg="#e8f1f8", fg="#555555",
                     font=("Segoe UI", 8, "bold")).pack()
            tk.Label(box, text=value, bg="#e8f1f8", fg=color,
                     font=("Segoe UI", 14, "bold")).pack()
        for priority in ("P1", "P2", "P3", "P4"):
            color = PRIORITY_COLORS.get(priority, "gray")
            tk.Label(self.report_inner_frame, text=priority, fg=color, bg="#f0f0f0",
                     font=("Segoe UI", 13, "bold")).pack(anchor="w", padx=6, pady=(6, 2))
            todos = groups[priority]
            if not todos:
                tk.Label(self.report_inner_frame, text="No activity in this range.", fg="gray",
                         bg="#f0f0f0").pack(anchor="w", padx=16, pady=(0, 4))
                continue
            for todo in todos:
                self._render_report_card(priority, color, todo)
        tk.Label(self.report_inner_frame, text="Unplanned Tasks", fg="#333333", bg="#f0f0f0",
                 font=("Segoe UI", 13, "bold")).pack(anchor="w", padx=6, pady=(10, 2))
        if not unplanned_entries:
            tk.Label(self.report_inner_frame, text="No unplanned tasks in this range.", fg="gray",
                     bg="#f0f0f0").pack(anchor="w", padx=16, pady=(0, 6))
        else:
            card = tk.Frame(self.report_inner_frame, bg="white", bd=1, relief="solid")
            card.pack(fill="x", padx=6, pady=(0, 6))
            for day_str, time_str, text in unplanned_entries:
                label = tk.Label(card, text=f"{day_str} {time_str} - {text}", bg="white",
                                 font=("Consolas", 9), anchor="w", justify="left")
                label.pack(anchor="w", padx=8, pady=2, fill="x")
                self._register_wrap_label(label, padding=40)

    def _render_report_card(self, priority, priority_color, todo):
        card = tk.Frame(self.report_inner_frame, bg="white", bd=1, relief="solid")
        card.pack(fill="x", padx=6, pady=4)
        header = tk.Frame(card, bg="white")
        header.pack(fill="x", padx=8, pady=(6, 2))
        tk.Label(header, text=priority, fg=priority_color, bg="white",
                 font=("Segoe UI", 10, "bold")).pack(side="left")
        status = todo["status"]
        status_color = STATUS_COLORS.get(status, "#6b6b6b")
        status_label = STATUS_LABELS.get(status, status.upper())
        tk.Label(header, text=status_label, fg=status_color, bg="white",
                 font=("Segoe UI", 8, "bold"), padx=6, pady=1).pack(side="left", padx=(8, 0))
        task_label = tk.Label(header, text=todo["task"], bg="white", font=("Segoe UI", 10, "bold"),
                              anchor="w", justify="left")
        task_label.pack(side="left", padx=(8, 0), fill="x", expand=True)
        self._register_wrap_label(task_label, padding=160)
        tk.Label(card, text=f"Id: {todo['id']}  | Time spent: {todo.get('total_time', '00:00')}",
                 bg="white", fg="gray", font=("Segoe UI", 7)).pack(anchor="w", padx=8)
        log_frame = tk.Frame(card, bg="white")
        log_frame.pack(fill="x", padx=8, pady=(2, 8))
        for day_str, time_str, text in todo["entries"]:
            label = tk.Label(log_frame, text=f"{day_str} {time_str} - {text}", bg="white",
                             font=("Consolas", 9), anchor="w", justify="left")
            label.pack(anchor="w", fill="x")
            self._register_wrap_label(label, padding=40)

    def download_report_pdf(self):
        if not self._report_groups:
            return
        start_str = self.report_start_entry.get().strip().replace("/", "-")
        end_str = self.report_end_entry.get().strip().replace("/", "-")
        filename = f"report_{start_str}_to_{end_str}.pdf"
        filepath = os.path.join(BASE_DIR, filename)
        try:
            generate_report_pdf(filepath, "Task Tracker Report", self._report_start_dt,
                                self._report_end_dt, self._report_groups, self._report_unplanned)
        except OSError as exc:
            messagebox.showerror("Download Failed", f"Could not write PDF:\n{exc}")
            return
        messagebox.showinfo("Report Saved", f"Report saved to:\n{filepath}")






































































































































































































































































































































































































































def main():
    root = tk.Tk()
    TaskTrackerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()

 
