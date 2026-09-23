"""
main.py - DLD Lab Word Template Generator

A single-window Tkinter (ttk) desktop app that walks the user through three steps:
    Step 1: choose how many headings the report needs (1-100)
    Step 2: type a title for each heading, one at a time
    Step 3: generate a .docx by copying cover_page/cover.docx and appending, for every
            heading: a "Heading 1" paragraph, 5 empty screenshot-placeholder lines and a
            5x3 truth-table (A | B | X).

Dependencies: python-docx, tkinter (with ttk).
"""

from __future__ import annotations

import re
import shutil
import sys
from pathlib import Path

# --------------------------------------------------------------------------------------
# Imports that can fail. tkinter is required to show anything at all, so if it is
# missing we can only report on stderr. python-docx is imported defensively so that the
# GUI can display a friendly error screen instead of a raw traceback.
# --------------------------------------------------------------------------------------
try:
    import tkinter as tk
    from tkinter import filedialog, ttk
except ImportError as exc:  # pragma: no cover - depends on the Python installation
    sys.stderr.write(f"ERROR: tkinter is not available in this Python installation: {exc}\n")
    sys.exit(1)

try:
    from docx import Document
    from docx.enum.style import WD_STYLE_TYPE
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Pt
except ImportError as exc:
    Document = None
    DOCX_IMPORT_ERROR = exc          # remembered so the GUI can show it
else:
    DOCX_IMPORT_ERROR = None


# ======================================================================================
# CONFIGURATION - everything you are likely to want to customise lives here.
# (See the "Customization Guide" in README.md.)
# ======================================================================================

# Folder of this script; used so the app works no matter which directory you launch from.
BASE_DIR = Path(__file__).resolve().parent

# Location of the cover-page template. Never auto-created.
COVER_PATH = BASE_DIR / "cover_page" / "cover.docx"
COVER_DISPLAY = "cover_page/cover.docx"      # short form used in messages

# Allowed range for the number of headings.
MIN_HEADINGS = 1
MAX_HEADINGS = 100

# Paragraph style used for each heading.
HEADING_STYLE = "Heading 1"
# Optional font overrides for heading text (None = keep whatever the style defines).
HEADING_FONT_NAME = None          # e.g. "Times New Roman"
HEADING_FONT_SIZE_PT = None       # e.g. 14

# Number of empty paragraphs inserted under each heading for screenshots.
SCREENSHOT_PLACEHOLDER_COUNT = 5

# Table style name. Must exist in cover.docx; if it does not, plain black borders are
# drawn instead. Use None for "no style, no borders".
TABLE_STYLE = "Table Grid"
# Optional font overrides for table text (None = keep document default).
TABLE_FONT_NAME = None
TABLE_FONT_SIZE_PT = None

# Table contents: first row is the header, remaining rows are the A/B presets with an
# empty X column to be filled in by the student. Table dimensions are derived from this
# list (5 rows x 3 columns by default). Every row must have the same number of cells.
TABLE_ROWS = [
    ["A", "B", "X"],
    ["0", "0", ""],
    ["0", "1", ""],
    ["1", "0", ""],
    ["1", "1", ""],
]

# Colours for inline warnings / status messages.
ERROR_COLOR = "#b00020"
OK_COLOR = "#1b5e20"


# ======================================================================================
# DOCUMENT GENERATION (no GUI code in this section)
# ======================================================================================

class GenerationError(Exception):
    """Raised for failures whose message is already written for the end user."""


def missing_cover_message() -> str:
    """Text shown whenever cover_page/cover.docx cannot be found."""
    return (
        f"Missing required file: {COVER_DISPLAY} (expected at {COVER_PATH}). "
        "Create the 'cover_page' folder next to main.py, put 'cover.docx' inside it, "
        "and start the app again."
    )


def permission_message(path: str) -> str:
    """Text shown when Windows/OS refuses to write the output file."""
    return (
        f"Permission denied while writing '{path}'. The file may be open in Word, or "
        "the file/folder is read-only. Close the file in Word or choose another "
        "location or file name."
    )


def ensure_heading_style(doc) -> None:
    """
    Make sure the paragraph style named HEADING_STYLE exists in the document.

    Word only stores a style like "Heading 1" in a file once it has been used or
    modified, so a plain cover.docx may not contain it. If it is missing we create a
    simple built-in-style replacement (bold, 16 pt, outline level 1) so headings still
    appear in Word's Navigation Pane. If it exists, nothing is changed.
    """
    styles = doc.styles
    try:
        styles[HEADING_STYLE]            # KeyError if the style is not defined
        return
    except KeyError:
        pass

    style = styles.add_style(HEADING_STYLE, WD_STYLE_TYPE.PARAGRAPH, builtin=True)
    try:
        style.base_style = styles["Normal"]
    except KeyError:
        pass                              # no Normal style; leave base unset
    style.font.bold = True
    style.font.size = Pt(16)
    style.paragraph_format.space_before = Pt(12)
    style.paragraph_format.space_after = Pt(6)
    style.paragraph_format.keep_with_next = True

    # Outline level 0 = "Heading 1" level; appended last to respect XML element order.
    outline = OxmlElement("w:outlineLvl")
    outline.set(qn("w:val"), "0")
    style.element.get_or_add_pPr().append(outline)


def add_manual_borders(table) -> None:
    """
    Draw single black borders around every cell by editing the table XML directly.
    Used as a fallback when TABLE_STYLE does not exist in cover.docx.
    """
    tbl_pr = table._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        element = OxmlElement(f"w:{edge}")
        element.set(qn("w:val"), "single")
        element.set(qn("w:sz"), "4")          # 1/8 pt units -> 0.5 pt line
        element.set(qn("w:space"), "0")
        element.set(qn("w:color"), "000000")
        borders.append(element)
    # tblBorders must come before tblLook in the schema; insert accordingly.
    tbl_look = tbl_pr.find(qn("w:tblLook"))
    if tbl_look is not None:
        tbl_look.addprevious(borders)
    else:
        tbl_pr.append(borders)


def apply_table_style(table) -> None:
    """Apply TABLE_STYLE to the table; fall back to manual borders if it is missing."""
    if not TABLE_STYLE:
        return                                # None/"" -> leave the table unstyled
    try:
        table.style = TABLE_STYLE
    except KeyError:
        add_manual_borders(table)


def apply_font(paragraph, font_name, size_pt) -> None:
    """Override font name/size on every run of a paragraph (only when values are set)."""
    for run in paragraph.runs:
        if font_name:
            run.font.name = font_name
        if size_pt:
            run.font.size = Pt(size_pt)


def add_heading_block(doc, title: str) -> None:
    """
    Append one complete section to the document:
        1. heading paragraph
        2. SCREENSHOT_PLACEHOLDER_COUNT empty paragraphs
        3. a table built from TABLE_ROWS
    """
    # 1. Heading paragraph in the configured style.
    heading = doc.add_paragraph(title, style=HEADING_STYLE)
    apply_font(heading, HEADING_FONT_NAME, HEADING_FONT_SIZE_PT)

    # 2. Empty lines where the student pastes screenshots.
    for _ in range(SCREENSHOT_PLACEHOLDER_COUNT):
        doc.add_paragraph("")

    # 3. Truth table.
    table = doc.add_table(rows=len(TABLE_ROWS), cols=len(TABLE_ROWS[0]))
    apply_table_style(table)
    for r_index, row_values in enumerate(TABLE_ROWS):
        for c_index, value in enumerate(row_values):
            cell = table.cell(r_index, c_index)
            cell.text = value
            for paragraph in cell.paragraphs:
                apply_font(paragraph, TABLE_FONT_NAME, TABLE_FONT_SIZE_PT)


def build_document_content(doc, titles: list[str]) -> None:
    """Validate the table configuration, then append one block per heading, in order."""
    columns = len(TABLE_ROWS[0]) if TABLE_ROWS else 0
    if columns == 0 or any(len(row) != columns for row in TABLE_ROWS):
        raise GenerationError(
            "TABLE_ROWS in main.py is misconfigured: it must be non-empty and every "
            "row must have the same number of cells."
        )
    ensure_heading_style(doc)
    for title in titles:
        add_heading_block(doc, title)


def _remove_quietly(path: str) -> None:
    """Best-effort delete of a partially created output file; never raises."""
    try:
        Path(path).unlink()
    except OSError:
        pass


def generate_document(out_path: str, titles: list[str]) -> None:
    """
    Create the report template at out_path.

    Steps: check the cover exists -> shutil.copy it to out_path -> open the copy with
    python-docx -> append the heading blocks -> save.
    Raises GenerationError (with a user-ready message) for every anticipated failure.
    Any unexpected exception is allowed to propagate to the caller.
    """
    # Re-check here: the file could have been deleted since start-up.
    if not COVER_PATH.is_file():
        raise GenerationError(missing_cover_message())

    # --- Copy the cover page to the chosen destination ---------------------------------
    # NOTE: SameFileError and PermissionError are subclasses of OSError, so they must be
    # caught before the generic OSError handler.
    try:
        shutil.copy(COVER_PATH, out_path)
    except shutil.SameFileError:
        raise GenerationError(
            "Cannot save over the template itself. Choose a different file name or folder."
        ) from None
    except PermissionError:
        raise GenerationError(permission_message(out_path)) from None
    except OSError as exc:
        raise GenerationError(f"Could not write '{out_path}': {exc.strerror or exc}") from exc

    # --- From here on a file exists at out_path; if anything fails we delete it so the
    # --- user is not left with a half-made document.
    try:
        # Open the copy. Any failure other than a permission problem means the
        # cover file is not a readable .docx.
        try:
            doc = Document(out_path)
        except PermissionError:
            raise GenerationError(permission_message(out_path)) from None
        except Exception as exc:
            raise GenerationError(
                f"{COVER_DISPLAY} could not be read - it may be corrupt or not a real "
                f".docx file ({type(exc).__name__}: {exc})."
            ) from exc

        # Append the per-heading content.
        build_document_content(doc, titles)

        # Save in place.
        try:
            doc.save(out_path)
        except PermissionError:
            raise GenerationError(permission_message(out_path)) from None
        except OSError as exc:
            raise GenerationError(f"Could not save '{out_path}': {exc.strerror or exc}") from exc
    except Exception:
        _remove_quietly(out_path)
        raise


# ======================================================================================
# INPUT VALIDATION (pure functions, no GUI)
# ======================================================================================

def parse_heading_count(text: str):
    """
    Validate the Step 1 input.
    Returns (value, None) when valid, or (None, warning_message) when invalid.
    Only plain ASCII digits are accepted: no signs, decimals, spaces inside, or letters.
    """
    cleaned = text.strip()
    if not cleaned:
        return None, f"Enter a whole number from {MIN_HEADINGS} to {MAX_HEADINGS}."
    if not re.fullmatch(r"[0-9]+", cleaned):
        return None, "Digits only - no letters, decimals or signs."
    if len(cleaned) > 9:                       # avoid absurdly long numbers
        return None, f"Number must be between {MIN_HEADINGS} and {MAX_HEADINGS}."
    value = int(cleaned)
    if value < MIN_HEADINGS or value > MAX_HEADINGS:
        return None, f"Number must be between {MIN_HEADINGS} and {MAX_HEADINGS}."
    return value, None


# ======================================================================================
# GUI
# ======================================================================================

class App:
    """
    Single-window wizard. The top area ("content") is cleared and rebuilt for each step;
    the bottom status label is created once and lives for the whole session.
    """

    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.heading_count = None          # int once Step 1 has been passed
        self.titles: list[str] = []        # one entry per heading ("" = not entered yet)
        self.current_index = 0             # 0-based heading being edited in Step 2
        self.fatal = False                 # True when a start-up error stopped the app

        # Window setup: resizable, with a sensible minimum size.
        root.title("DLD Lab Word Template Generator")
        root.geometry("560x280")
        root.minsize(440, 240)
        root.columnconfigure(0, weight=1)
        root.rowconfigure(0, weight=1)     # content area grows, status stays at the bottom

        # Content area (swapped per step).
        self.content = ttk.Frame(root, padding=16)
        self.content.grid(row=0, column=0, sticky="nsew")
        self.content.columnconfigure(0, weight=1)
        self.content.rowconfigure(9, weight=1)   # spacer row pushes buttons (row 10) down

        # Separator + single status label at the bottom of the window.
        ttk.Separator(root, orient="horizontal").grid(row=1, column=0, sticky="ew")
        self.status_label = ttk.Label(root, text="", anchor="w", justify="left", padding=(16, 6))
        self.status_label.grid(row=2, column=0, sticky="ew")
        # Keep the status text wrapped to the current window width (long paths/errors).
        self.status_label.bind(
            "<Configure>",
            lambda event: self.status_label.configure(wraplength=max(event.width - 32, 100)),
        )

        # Start-up checks: fail early with a clear message instead of a traceback.
        problem = self.startup_problem()
        if problem:
            self.show_fatal(problem)
        else:
            self.show_step1()

    # ---------------------------------------------------------------- helpers ---------

    def startup_problem(self):
        """Return an error message if the app cannot run, else None."""
        if DOCX_IMPORT_ERROR is not None:
            return (
                "The 'python-docx' package is not installed (or failed to import). "
                "Run:  pip install python-docx   "
                "(install 'python-docx', not 'docx').  "
                f"Details: {DOCX_IMPORT_ERROR}"
            )
        if not COVER_PATH.is_file():
            return missing_cover_message()
        return None

    def set_status(self, text: str, kind: str = "") -> None:
        """Update the bottom status label. kind: 'error', 'ok' or '' (neutral)."""
        color = {"error": ERROR_COLOR, "ok": OK_COLOR}.get(kind, "")
        self.status_label.configure(text=text, foreground=color)

    def _reset_content(self) -> None:
        """Destroy the widgets of the current step and clear the status line."""
        for child in self.content.winfo_children():
            child.destroy()
        self.set_status("")

    def _add_warning_label(self, row: int) -> tk.StringVar:
        """Create an (initially empty) red inline-warning label and return its variable."""
        var = tk.StringVar(value="")
        label = ttk.Label(self.content, textvariable=var, foreground=ERROR_COLOR, justify="left")
        label.grid(row=row, column=0, sticky="w", pady=(4, 0))
        return var

    def _add_buttons(self, specs):
        """Place right-aligned buttons at the bottom of the content area. specs = [(text, command)]."""
        frame = ttk.Frame(self.content)
        frame.grid(row=10, column=0, sticky="e", pady=(12, 0))
        buttons = []
        for column, (text, command) in enumerate(specs):
            button = ttk.Button(frame, text=text, command=command)
            button.grid(row=0, column=column, padx=(6, 0))
            buttons.append(button)
        return buttons

    # ------------------------------------------------------------- fatal screen -------

    def show_fatal(self, message: str) -> None:
        """
        Show an unrecoverable start-up error inside the main window (no popup) with an
        Exit button. The message is also written to stderr, and the process exits with
        status 1 once the window is closed.
        """
        self.fatal = True
        sys.stderr.write(f"ERROR: {message}\n")
        self._reset_content()
        label = ttk.Label(self.content, text=message, foreground=ERROR_COLOR, justify="left")
        label.grid(row=0, column=0, sticky="ew")
        label.bind("<Configure>", lambda event: label.configure(wraplength=max(event.width - 10, 100)))
        exit_button, = self._add_buttons([("Exit", self.root.destroy)])
        exit_button.focus_set()
        self.set_status("Fatal error - the application cannot continue. Click Exit to close.", "error")

    # ------------------------------------------------------------------ Step 1 -------

    def show_step1(self) -> None:
        """Step 1: ask for the number of headings (integer 1-100)."""
        self._reset_content()
        ttk.Label(self.content, text="Number of headings:").grid(row=0, column=0, sticky="w")

        # Pre-fill with the stored count when returning from Step 2 ("preserve if valid").
        initial = str(self.heading_count) if self.heading_count else str(MIN_HEADINGS)
        self.count_var = tk.StringVar(value=initial)

        # Warning label is created BEFORE the spinbox because the key-validation
        # callbacks below write to it.
        self.count_warning = self._add_warning_label(row=2)

        # Key-level validation: only up to 3 ASCII digits can be typed/pasted.
        vcmd = (self.root.register(self._on_count_key), "%P")
        icmd = (self.root.register(self._on_count_key_rejected),)
        self.count_spin = ttk.Spinbox(
            self.content, from_=MIN_HEADINGS, to=MAX_HEADINGS, increment=1, width=8,
            textvariable=self.count_var,
            validate="key", validatecommand=vcmd, invalidcommand=icmd,
        )
        self.count_spin.grid(row=1, column=0, sticky="w", pady=(6, 0))
        self.count_spin.bind("<Return>", lambda event: self.on_step1_next())
        self.count_spin.focus_set()

        self._add_buttons([("Next", self.on_step1_next)])    # no Back button on Step 1

    def _on_count_key(self, proposed: str) -> bool:
        """validatecommand: allow the edit only if the result is empty or 1-3 ASCII digits."""
        allowed = re.fullmatch(r"[0-9]{0,3}", proposed) is not None
        if allowed:
            self.count_warning.set("")
        return allowed

    def _on_count_key_rejected(self) -> None:
        """invalidcommand: tell the user why the keystroke was ignored."""
        self.count_warning.set("Only whole numbers are allowed (digits 0-9, up to 3 digits).")

    def on_step1_next(self) -> None:
        """Validate the count; advance to Step 2 only if valid."""
        value, warning = parse_heading_count(self.count_var.get())
        if warning:
            self.count_warning.set(warning)
            return
        self.count_warning.set("")

        # Resize the title list to the new count, keeping titles already typed.
        self.titles = self.titles[:value] + [""] * (value - len(self.titles))
        self.heading_count = value
        self.current_index = 0
        self.show_step2()

    # ------------------------------------------------------------------ Step 2 -------

    def show_step2(self) -> None:
        """Step 2: ask for the title of heading number current_index+1."""
        self._reset_content()
        total = self.heading_count
        index = self.current_index
        ttk.Label(
            self.content, text=f"Heading {index + 1} of {total} \u2014 enter title:"
        ).grid(row=0, column=0, sticky="w")

        # Pre-fill with any previously entered title for this heading.
        self.title_var = tk.StringVar(value=self.titles[index])
        self.title_entry = ttk.Entry(self.content, textvariable=self.title_var)
        self.title_entry.grid(row=1, column=0, sticky="ew", pady=(6, 0))
        self.title_entry.bind("<Return>", lambda event: self.on_step2_next())
        self.title_entry.focus_set()
        self.title_entry.icursor("end")

        self.title_warning = self._add_warning_label(row=2)
        self._add_buttons([("Back", self.on_step2_back), ("Next", self.on_step2_next)])

    def on_step2_back(self) -> None:
        """Return to Step 1. A non-blank title being typed is kept for later."""
        typed = self.title_var.get().strip()
        if typed:
            self.titles[self.current_index] = typed
        self.show_step1()

    def on_step2_next(self) -> None:
        """Validate the title (non-empty), store it, and move on (Step 3 after the last one)."""
        title = self.title_var.get().strip()
        if not title:
            self.title_warning.set("Title cannot be empty.")
            return
        self.titles[self.current_index] = title

        if self.current_index < self.heading_count - 1:
            self.current_index += 1
            self.show_step2()
        else:
            self.show_step3()

    # ------------------------------------------------------------------ Step 3 -------

    def show_step3(self) -> None:
        """Step 3: summary + Generate button."""
        self._reset_content()
        ttk.Label(
            self.content, text=f"Number of headings: {self.heading_count}"
        ).grid(row=0, column=0, sticky="w")

        _back, generate_button = self._add_buttons(
            [("Back", self.on_step3_back), ("Generate", self.on_generate)]
        )
        generate_button.focus_set()

    def on_step3_back(self) -> None:
        """Return to Step 2 at the last heading."""
        self.current_index = self.heading_count - 1
        self.show_step2()

    def on_generate(self) -> None:
        """Ask where to save, build the document, and report the result in the status label."""
        out_path = filedialog.asksaveasfilename(
            parent=self.root,
            title="Save Word document",
            defaultextension=".docx",
            filetypes=[("Word document", "*.docx")],
        )
        # Cancelling returns "" (or an empty tuple on some platforms).
        if not out_path:
            self.set_status("Save cancelled - no file was created.")
            return

        # Some platforms do not append the default extension; make sure it is .docx.
        if not out_path.lower().endswith(".docx"):
            out_path += ".docx"

        try:
            generate_document(out_path, self.titles)
        except GenerationError as exc:
            self.set_status(str(exc), "error")
        except Exception as exc:               # anything we did not anticipate
            self.set_status(f"Unexpected error ({type(exc).__name__}): {exc}", "error")
        else:
            self.set_status(f"Saved: {out_path}", "ok")


# ======================================================================================
# ENTRY POINT
# ======================================================================================

def main() -> int:
    """Create the window, run the event loop, and return the process exit code."""
    root = tk.Tk()
    app = App(root)
    root.mainloop()
    return 1 if app.fatal else 0


if __name__ == "__main__":
    sys.exit(main())