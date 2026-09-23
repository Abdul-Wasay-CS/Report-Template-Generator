# DLD Lab Word Template Generator

## 1. Overview

A small desktop app (Python + Tkinter) that builds a Word template for Digital Logic
Design (DLD) lab reports. You tell it how many headings you need and give each a title.
It then creates a `.docx` that starts with your own cover page (`cover_page/cover.docx`)
and, for every heading, appends:

1. the heading (Word style **Heading 1**),
2. 5 empty lines where you paste screenshots,
3. a 5 × 3 truth table:

| A | B | X |
|---|---|---|
| 0 | 0 |   |
| 0 | 1 |   |
| 1 | 0 |   |
| 1 | 1 |   |

Project layout:

```
project/
  main.py
  cover_page/cover.docx   <- you provide this yourself
  README.md
```

The app never creates `cover_page/` or `cover.docx`. If the cover file is missing, the
app shows an error and exits.

## 2. Requirements

- **Python 3.9 or newer**
- **tkinter** with `ttk` (bundled with the python.org installers for Windows and macOS;
  on Debian/Ubuntu install it with `sudo apt install python3-tk`)
- **python-docx** (PyPI package name: `python-docx`)
- A cover page file at `cover_page/cover.docx`

## 3. Install

1. Put `main.py` and `README.md` in a folder (e.g. `project/`).
2. Create the folder `project/cover_page/` and place your `cover.docx` inside it.
3. Install the dependency (optionally inside a virtual environment):

```
pip install python-docx
```

> Install `python-docx`, **not** `docx` (a different, unrelated package).

## 4. Run

From inside the project folder:

```
python main.py
```

On Windows you can also use `py main.py`. On Linux/macOS you may need `python3 main.py`.
The app finds `cover_page/cover.docx` relative to `main.py`, so it works from any
current directory.

## 5. User Manual

### 5.1 Operating the app

**Step 1 - Number of headings**

1. Type a whole number from 1 to 100 in the box (or use the arrows).
2. Click **Next**. If the value is empty, not a whole number, or outside 1-100, a red
   warning appears under the box and you stay on this step.

**Step 2 - Heading titles** (one screen per heading)

1. The label reads `Heading X of N - enter title:`. Type the title.
2. Click **Next**. An empty or spaces-only title shows a warning and does not advance.
3. After the last heading you land on Step 3.
4. **Back** returns to Step 1. The count is kept, and titles you already entered stay
   filled in (if you change the count, extra titles are dropped or new blanks are added).

**Step 3 - Generate**

1. The window shows the number of headings.
2. Click **Generate** and choose where to save the `.docx` file.
3. The status line at the bottom shows `Saved: <path>` on success, or the error reason.
4. **Back** returns to Step 2 at the last heading.

Cancelling the save dialog creates no file and just shows a status message.

### 5.2 Filling in the generated document

1. Open the saved `.docx` in Word. Your cover page is first, followed by one section
   per heading, in the order you entered them.
2. Under each heading, click into the empty lines and paste your screenshots
   (5 blank lines are provided; Word will push the table down as needed).
3. In each table, columns **A** and **B** already hold the input combinations. Fill
   the empty **X** cells in rows 2-5 with the output (0 or 1) you observed.
4. Save the document from Word as usual.

## 6. Customization Guide

All options are constants near the top of `main.py`, in the block labelled
`CONFIGURATION`. Edit the value, save, and restart the app.

| What | Constant(s) | Default |
|------|-------------|---------|
| Heading style | `HEADING_STYLE` | `"Heading 1"` |
| Heading font | `HEADING_FONT_NAME`, `HEADING_FONT_SIZE_PT` | `None` (use the style's font) |
| Table style | `TABLE_STYLE` | `"Table Grid"` |
| Table font | `TABLE_FONT_NAME`, `TABLE_FONT_SIZE_PT` | `None` |
| Screenshot placeholders | `SCREENSHOT_PLACEHOLDER_COUNT` | `5` |
| Table size and A/B presets | `TABLE_ROWS` | 5 rows × 3 columns |
| Heading count limits | `MIN_HEADINGS`, `MAX_HEADINGS` | `1`, `100` |

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

### Optional: page break after the cover

Not included by default. To add one, insert this line as the first statement of
`build_document_content()` in `main.py`:

```python
doc.add_page_break()
```

## 7. Error Cheat-Sheet

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
| 12 | `TABLE_ROWS` misconfigured | Status line, Step 3 | `TABLE_ROWS in main.py is misconfigured` | Rows have different lengths, or the list is empty | Make every row the same length (see Section 6) |
| 13 | `tkinter` not available | Terminal (no window can open) | `tkinter is not available in this Python installation` | Python was installed without Tk support | Linux: `sudo apt install python3-tk`; otherwise reinstall Python from python.org |