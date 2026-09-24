# DLD Lab Word Template Generator

A small desktop app that builds a Word template for your DLD lab report. It starts
with your own cover page, then adds three fixed headings (**Critical Analysis**,
**Summary of the skills learned**, **Conclusion**) followed by the extra headings you
type in. Under every heading you get 5 empty lines for screenshots and a small
A / B / X table. The first heading starts on page 2, so your cover page stays clean.

---

## 1. Install & Run (Normal Users)

### Before you start (Windows and Linux)

Your project folder must look like this:

```
project/
  main.py
  README.md
  cover_page/
    cover.docx      <- your own cover page. You must put it here yourself.
```

The app does **not** create `cover.docx` for you. If it is missing, the app shows an
error and closes.

### Windows

1. **Install Python.** Go to [python.org](https://www.python.org/downloads/), download
   Python and run the installer. On the first screen, tick **"Add python.exe to
   PATH"**, then click **Install Now**.
2. **Open a command window in the project folder.** Open the `project` folder in File
   Explorer, click the address bar at the top, type `cmd` and press **Enter**. A black
   window opens.
3. **Install the one extra piece the app needs** (only the first time). Type this and
   press **Enter**, then wait until it finishes:

```
   pip install python-docx
```

4. **Start the app.** Type this and press **Enter**:

```
   python main.py
```

   If Windows says `python` is not recognized, type `py main.py` instead.

Next time you only need steps 2 and 4.

### Linux (Ubuntu, Fedora)

1. **Open a Terminal in the project folder.** In your file manager, open the `project`
   folder, right-click an empty area and choose **Open in Terminal**.
2. **Install what the app needs** (only the first time). Type the line for your system
   and press **Enter**. You will be asked for your password.

   **Ubuntu:**

```
   sudo apt update
   sudo apt install python3 python3-tk python3-docx
```

   **Fedora:**

```
   sudo dnf install python3 python3-tkinter python3-docx
```

3. **Start the app.** Type this and press **Enter**:

```
   python3 main.py
```

Next time you only need steps 1 and 3.

If something goes wrong, look up the message in the **Error Cheat-Sheet** below.

---

## 2. App Walkthrough

*Screenshots to be added. Put your images in a `screenshots/` folder next to
`main.py`; the links below are placeholders.*

### Step 1 — Number of additional headings

![Step 1](screenshots/step1.png)

### Step 1 — Warning on invalid input

![Step 1 warning](screenshots/step1-warning.png)

### Step 2 — Heading titles

![Step 2](screenshots/step2.png)

### Step 2 — Warning on empty title

![Step 2 warning](screenshots/step2-warning.png)

### Step 3 — Generate

![Step 3](screenshots/step3.png)

### Save dialog

![Save dialog](screenshots/save-dialog.png)

### Success (the app closes by itself)

![Success](screenshots/success.png)

### Error (the window stays open)

![Error](screenshots/error.png)

### Generated document — cover page and first heading

![Generated document](screenshots/document-overview.png)

### Filling in the screenshots and the table

![Filling in the document](screenshots/document-filling.png)

---

## 3. Error Cheat-Sheet

Messages appear either on the start-up error screen (with an **Exit** button), as a red
inline warning under an input, or in the status line at the bottom of the window.

| # | Situation | Where shown | Message begins with | Meaning | Fix |
|---|-----------|-------------|---------------------|---------|-----|
| 1 | Missing `cover_page/cover.docx` | Start-up error screen (app exits); or status line if deleted while running | `Missing required file: cover_page/cover.docx` | The template is not where the app expects it | Create `cover_page/` next to `main.py`, put `cover.docx` inside, restart |
| 2 | `python-docx` not installed | Start-up error screen (app exits) | `The 'python-docx' package is not installed` | The library cannot be imported | `pip install python-docx` (not `docx`), run with the same Python you installed it for |
| 3 | Invalid heading count | Inline warning, Step 1 | `Enter a whole number from 1 to 100.` / `Digits only ...` / `Number must be between 1 and 100.` | Empty, non-numeric, decimal, negative, or outside 1-100 | Type a whole number from 1 to 100 |
| 4 | Empty heading title | Inline warning, Step 2 | `Title cannot be empty.` | The field is blank or spaces only | Type a title |
| 5 | Save dialog cancelled | Status line, Step 3 | `Save cancelled - no file was created.` | You closed the dialog without choosing a file | Click **Generate** again and pick a location |
| 6 | File write permission denied | Status line, Step 3 | `Permission denied while writing ...` | The OS blocked writing (read-only folder, protected location, or read-only file) | Save somewhere you can write (Documents, Desktop); check the file/folder is not read-only |
| 7 | Path already open in Word | Status line, Step 3 | `Permission denied while writing ...` (same message as #6) | Word locks open files, so they cannot be overwritten. A hidden `~$name.docx` file next to it confirms this | Close the file in Word and generate again, or choose a different file name |
| 8 | Corrupt/unreadable `cover.docx` | Status line, Step 3 | `cover_page/cover.docx could not be read ...` | The file is damaged, empty, or not a real `.docx` (for example a renamed `.doc` or `.pdf`) | Open it in Word and re-save as `.docx`, or replace it with a good copy. The partial output file is deleted automatically |
| 9 | Any unexpected exception | Status line, Step 3 | `Unexpected error (<ExceptionName>): ...` | Something the app did not anticipate | Note the full message; retry with a different save location; check `cover.docx` opens in Word. The partial output file is deleted automatically |
| 10 | Save path is the template itself | Status line, Step 3 | `Cannot save over the template itself.` | You chose `cover_page/cover.docx` as the destination | Choose a different file name or folder |
| 11 | Folder missing / disk full / other OS error | Status line, Step 3 | `Could not write '<path>': <reason>` (or `Could not save`) | The operating system reported a problem other than permissions | Fix the reason shown (choose an existing folder, free disk space) |
| 12 | `TABLE_ROWS` misconfigured | Status line, Step 3 | `TABLE_ROWS in main.py is misconfigured` | Rows have different lengths, or the list is empty | Make every row the same length (see the Customization Guide) |
| 13 | `tkinter` not available | Terminal (no window can open) | `tkinter is not available in this Python installation` | Python was installed without Tk support | Linux: `sudo apt install python3-tk`; otherwise reinstall Python from python.org |

---

## 4. Customization Guide

**⚠️ WARNING: For advanced users / developers only. Changing these may break the app.**

All options are constants near the top of `main.py`, in the block labelled
`CONFIGURATION`. Edit the value, save, and restart the app.

| What | Constant(s) | Default |
|------|-------------|---------|
| Fixed headings | `FIXED_HEADINGS` | `Critical Analysis`, `Summary of the skills learned`, `Conclusion` |
| Limits for additional headings | `MIN_HEADINGS`, `MAX_HEADINGS` | `1`, `100` |
| Auto-close delay | `AUTO_CLOSE_DELAY_MS` | `1500` |
| Heading style | `HEADING_STYLE` | `"Heading 1"` |
| Heading font | `HEADING_FONT_NAME`, `HEADING_FONT_SIZE_PT` | `None` (use the style's font) |
| Table style | `TABLE_STYLE` | `"Table Grid"` |
| Table font | `TABLE_FONT_NAME`, `TABLE_FONT_SIZE_PT` | `None` |
| Screenshot placeholders | `SCREENSHOT_PLACEHOLDER_COUNT` | `5` |
| Table size and A/B presets | `TABLE_ROWS` | 5 rows × 3 columns |

### Fixed headings

`FIXED_HEADINGS` is the list of headings placed at the start of every document, in
order, before the headings typed by the user. They get exactly the same treatment
(heading style, screenshot lines, table). Example:

```python
FIXED_HEADINGS = [
    "Critical Analysis",
    "Summary of the skills learned",
    "Conclusion",
]
```

The number typed in Step 1 counts only the additional headings, so the final document
contains `len(FIXED_HEADINGS)` + your count. The Step 1 label and the Step 3 summary
adapt automatically if you change the list.

### Cover page break

The first heading of the document is given Word's "Page break before" property inside
`build_document_content()` (the `start_on_new_page` argument of `add_heading_block()`).
This keeps everything off the cover page and starts the report on page 2. It is not
controlled by a constant.

### Auto-close delay

After a successful save the window closes after `AUTO_CLOSE_DELAY_MS` milliseconds (long
enough to see the `Saved: <path>` message). Set it to `0` to close immediately. On
failure the window always stays open.

### Heading style

`HEADING_STYLE = "Heading 1"` is looked up in `cover.docx`. Its appearance comes from
that file: in Word, right-click **Heading 1** in the Styles gallery → **Modify**, then
save `cover.docx`. If the style is not stored in `cover.docx`, the app creates a simple
one (bold, 16 pt) so headings still work. You can also force a font without touching
Word:

```python
HEADING_FONT_NAME = "Times New Roman"
HEADING_FONT_SIZE_PT = 14
```

### Table style

`TABLE_STYLE = "Table Grid"` gives visible borders. Any other style name only works if
that style exists inside `cover.docx`. If the named style is missing, the app draws plain
black single-line borders instead. `TABLE_STYLE = None` leaves the table completely
unstyled (usually no borders).

### Screenshot placeholder count

```python
SCREENSHOT_PLACEHOLDER_COUNT = 8   # 8 empty lines under each heading
```

`0` is allowed and inserts none.

### Table dimensions and A/B presets

The table is built directly from `TABLE_ROWS`: row 1 is the header, and rows/columns are
counted from the list. **Every row must have the same number of cells**, otherwise
generation stops with a "TABLE_ROWS is misconfigured" error.

Change the A/B input pattern (for example Gray-code order):

```python
TABLE_ROWS = [
    ["A", "B", "X"],
    ["0", "0", ""],
    ["0", "1", ""],
    ["1", "1", ""],
    ["1", "0", ""],
]
```

Three inputs (9 rows × 4 columns):

```python
TABLE_ROWS = [
    ["A", "B", "C", "X"],
    ["0", "0", "0", ""],
    ["0", "0", "1", ""],
    ["0", "1", "0", ""],
    ["0", "1", "1", ""],
    ["1", "0", "0", ""],
    ["1", "0", "1", ""],
    ["1", "1", "0", ""],
    ["1", "1", "1", ""],
]
```

Rename the column labels (for example `Y` instead of `X`) by editing the first row.

### Fonts

Use the four `*_FONT_NAME` / `*_FONT_SIZE_PT` constants. The font must be installed on the
computer that opens the document. `None` means "do not override".

---

## 5. Advanced / Dependency Reference

**For Arch Linux and advanced/custom setups.**
Raw dependencies for people who want to manage them manually. This is **not** part of the
normal install path — normal users should skip this section and use Section 1.

### Raw dependencies

| Dependency | Notes |
|------------|-------|
| Python 3.9 or newer | Standard-library modules used: `re`, `shutil`, `sys`, `pathlib`, `tkinter` |
| Tcl/Tk with the `tkinter` module (including `ttk`) | Usually a separate system package on Linux |
| `python-docx` (PyPI name `python-docx`) | Written against the 1.x API. Pulls in `lxml` and `typing_extensions` |

> Do not confuse `python-docx` with the unrelated PyPI package `docx`. The app imports
> `docx`, but the package to install is `python-docx`.

### Arch Linux (system packages)

```
sudo pacman -S python tk python-docx
python main.py
```

### Any system (virtual environment, pip)

Useful on systems that refuse `pip install` outside a virtual environment:

```
python -m venv .venv
source .venv/bin/activate        # Windows: .venv\Scripts\activate
pip install python-docx
python main.py
```

`tkinter` cannot be installed with pip; it must come from your Python installation or
your system's package manager.