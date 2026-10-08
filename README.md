# DLD Lab Word Template Generator

A small desktop app that builds a Word template for your DLD lab report. It starts
with your own cover page, optionally adds three fixed headings, and lets you add any
number of your own "gate" headings - each of which can get an automatically generated
truth table. The first heading starts on page 2, so your cover page stays clean.

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

The window shows a colored "Step X of N" progress bar at the top at every stage (N is
2 if you don't add any truth tables, 3 if you add at least one), and a status message
at the bottom of the window at all times.

*Screenshots to be added. Put your images in a `screenshots/` folder next to
`main.py`; the links below are placeholders.*

### Step 1 — Headings: fixed-headings checkbox and the "add a heading" form

![Step 1 - add heading form](screenshots/step1-add-heading.png)

### Step 1 — Warning on an empty title or missing/invalid input count

![Step 1 - warning](screenshots/step1-warning.png)

### Step 1 — Live list of headings added so far

![Step 1 - live list](screenshots/step1-heading-list.png)

### Step 1 — "Back" removing the last added heading

![Step 1 - remove last](screenshots/step1-remove-last.png)

### Step 1 — "Finish Headings" (small button)

![Step 1 - finish headings](screenshots/step1-finish-headings.png)

### Step 2 — Choosing All 1s / All 0s for a heading's F column (shown only if you added a truth table)

![Step 2 - F value selection](screenshots/step2-f-value.png)

### Step 3 — Generate: summary and Generate button

![Step 3 - generate](screenshots/step3-generate.png)

### Save dialog

![Save dialog](screenshots/save-dialog.png)

### Success (the app closes by itself)

![Success](screenshots/success.png)

### Error (the window stays open)

![Error](screenshots/error.png)

### Generated document — cover page and first heading

![Generated document](screenshots/document-overview.png)

### Generated document — a gate heading with its truth table

![Truth table in the document](screenshots/document-truth-table.png)

---

## 3. Formatting Reference

This is what the generated `.docx` looks like once opened in Word. Nothing here needs
to be set by hand - it is applied automatically during generation.

| Element | Font | Size | Style | Notes |
|---------|------|------|-------|-------|
| Main headings (the fixed ones: Summary of the skills learned, Conclusion, Critical Analysis) | Aptos Slab Extrabold | 18 pt | — | Text only. **No table** follows a fixed heading. |
| Body text (the 5 screenshot placeholder lines under every heading) | Calibri (Body) | 12 pt | — | Empty lines - paste your screenshots here. |
| Gate names (your user-typed headings) | Berlin Sans FB Demi | 28 pt | Bold | Followed by 5 screenshot lines, then a truth table if you requested one. |
| Table content (header row and value rows) | Times New Roman | 20 pt | Bold | Only appears under a gate heading that had "Include truth table" ticked. |

### How a truth table is built

- You choose the number of inputs for that heading (1 to 7).
- The table always has one header row plus 2^(number of inputs) data rows - for
  example 3 inputs gives 8 data rows, 7 inputs gives the maximum of 128 data rows.
- Input columns are labeled with letters in order (A, B, C, D, E, then G, H, ... -
  **the letter F is always skipped** as an input label), and are pre-filled with every
  binary combination in ascending order.
- The last column is always labeled **F** and is the output column.
- After you add all your headings, the app asks you - once per truth-table heading -
  to fill that heading's entire F column with either **all 1s** or **all 0s**. Every
  data row in that column gets the same value you chose; you edit individual rows
  yourself afterwards in Word if needed.

Example header and first rows for a 3-input table where **All 1s** was chosen:

```
A | B | C | F
0 | 0 | 0 | 1
0 | 0 | 1 | 1
0 | 1 | 0 | 1
...
```

> If a listed font (e.g. **Aptos Slab Extrabold** or **Berlin Sans FB Demi**) is not
> installed on the computer that opens the document, Word will substitute a similar
> font automatically; the size and bold settings still apply.

---

## 4. Error Cheat-Sheet

Messages appear either on the start-up error screen (with an **Exit** button), as a red
inline warning under an input, or in the status line at the bottom of the window.

| # | Situation | Where shown | Message begins with | Meaning | Fix |
|---|-----------|-------------|---------------------|---------|-----|
| 1 | Missing `cover_page/cover.docx` | Start-up error screen (app exits); or status line if deleted while running | `Missing required file: cover_page/cover.docx` | The template is not where the app expects it | Create `cover_page/` next to `main.py`, put `cover.docx` inside, restart |
| 2 | `python-docx` not installed | Start-up error screen (app exits) | `The 'python-docx' package is not installed` | The library cannot be imported | `pip install python-docx` (not `docx`), run with the same Python you installed it for |
| 3 | Empty heading title | Inline warning, Step 1 | `Title cannot be empty.` | "Heading title" is blank or spaces only when you click **Add Heading** | Type a title before adding |
| 4 | Missing or invalid number of inputs | Inline warning, Step 1 | `Enter a whole number of inputs from 1 to 7.` / `Digits only ...` / `Number of inputs must be between 1 and 7.` | "Include truth table" is ticked but the inputs field is empty, non-numeric, or outside 1-7 | Type a whole number from 1 to 7, or untick "Include truth table" |
| 5 | No headings added | Inline warning, Step 1 | `Add at least one heading before continuing.` | You clicked **Finish Headings** with an empty list | Add at least one heading first |
| 6 | F-value not selected | Inline warning, Step 2 | `Select All 1s or All 0s before continuing.` | You clicked **Next** on the F-value screen without choosing an option | Click **All 1s** or **All 0s**, then **Next** |
| 7 | Save dialog cancelled | Status line, Step 3 | `Save cancelled - no file was created.` | You closed the dialog without choosing a file | Click **Generate** again and pick a location |
| 8 | File write permission denied | Status line, Step 3 | `Permission denied while writing ...` | The OS blocked writing (read-only folder, protected location, or read-only file) | Save somewhere you can write (Documents, Desktop); check the file/folder is not read-only |
| 9 | Path already open in Word | Status line, Step 3 | `Permission denied while writing ...` (same message as #8) | Word locks open files, so they cannot be overwritten. A hidden `~$name.docx` file next to it confirms this | Close the file in Word and generate again, or choose a different file name |
| 10 | Corrupt/unreadable `cover.docx` | Status line, Step 3 | `cover_page/cover.docx could not be read ...` | The file is damaged, empty, or not a real `.docx` (for example a renamed `.doc` or `.pdf`) | Open it in Word and re-save as `.docx`, or replace it with a good copy. The partial output file is deleted automatically |
| 11 | Any unexpected exception | Status line, Step 3 | `Unexpected error (<ExceptionName>): ...` | Something the app did not anticipate | Note the full message; retry with a different save location; check `cover.docx` opens in Word. The partial output file is deleted automatically |
| 12 | Save path is the template itself | Status line, Step 3 | `Cannot save over the template itself.` | You chose `cover_page/cover.docx` as the destination | Choose a different file name or folder |
| 13 | Folder missing / disk full / other OS error | Status line, Step 3 | `Could not write '<path>': <reason>` (or `Could not save`) | The operating system reported a problem other than permissions | Fix the reason shown (choose an existing folder, free disk space) |
| 14 | `tkinter` not available | Terminal (no window can open) | `tkinter is not available in this Python installation` | Python was installed without Tk support | Linux: `sudo apt install python3-tk`; otherwise reinstall Python from python.org |

---

## 5. Customization Guide

**⚠️ WARNING: For advanced users / developers only. Changing these may break the app.**

All options are constants near the top of `main.py`, in the block labelled
`CONFIGURATION`. Edit the value, save, and restart the app.

| What | Constant(s) | Default |
|------|-------------|---------|
| Fixed heading titles | `FIXED_HEADING_SUMMARY`, `FIXED_HEADING_CONCLUSION`, `FIXED_HEADING_CRITICAL` | "Summary of the skills learned", "Conclusion", "Critical Analysis" |
| Limits for truth-table inputs | `MIN_TRUTH_TABLE_INPUTS`, `MAX_TRUTH_TABLE_INPUTS` | `1`, `7` |
| Input column letters / output letter | `INPUT_LETTER_POOL`, `OUTPUT_COLUMN_LETTER` | `"ABCDEGHIJKLMNOPQRSTUVWXYZ"` (skips `F`), `"F"` |
| Auto-close delay | `AUTO_CLOSE_DELAY_MS` | `1500` |
| Shared heading paragraph style | `HEADING_STYLE` | `"Heading 1"` |
| Fixed-heading font | `FIXED_HEADING_FONT_NAME`, `FIXED_HEADING_FONT_SIZE_PT`, `FIXED_HEADING_BOLD` | Aptos Slab Extrabold, 18, `None` |
| Gate-heading font | `GATE_HEADING_FONT_NAME`, `GATE_HEADING_FONT_SIZE_PT`, `GATE_HEADING_BOLD` | Berlin Sans FB Demi, 28, `True` |
| Body text font | `BODY_FONT_NAME`, `BODY_FONT_SIZE_PT` | Calibri (Body), 12 |
| Table font | `TABLE_FONT_NAME`, `TABLE_FONT_SIZE_PT`, `TABLE_BOLD` | Times New Roman, 20, `True` |
| Table style | `TABLE_STYLE` | `"Table Grid"` |
| Screenshot placeholders | `SCREENSHOT_PLACEHOLDER_COUNT` | `5` |
| GUI color palette | `COLOR_BG`, `COLOR_CARD_BG`, `COLOR_TEXT`, `COLOR_PRIMARY`, `COLOR_PRIMARY_DARK`, `COLOR_PRIMARY_TEXT` | soft blue/steel theme |

### Fixed heading titles and placement

`FIXED_HEADING_SUMMARY` and `FIXED_HEADING_CONCLUSION` are always inserted before the
user's own headings (in that order); `FIXED_HEADING_CRITICAL` is always inserted at the
very end of the document, after every user heading - this ordering is not
configurable by a constant, only the text of each title is. All three get a heading
paragraph plus the screenshot placeholder lines, and never a table. This logic lives in
`build_document_content()`.

### Maximum truth-table inputs

`MAX_TRUTH_TABLE_INPUTS = 7` caps the table at 2^7 = 128 data rows, which is already a
very long table in Word. Raising this constant is possible but will make documents with
large input counts slow to generate and unwieldy to scroll through.

### Input letters and the reserved output letter "F"

`INPUT_LETTER_POOL` lists, in order, which letters are used for input columns. The
letter in `OUTPUT_COLUMN_LETTER` ("F" by default) is always skipped when choosing input
letters, and is always used for the final (output) column, however many inputs there
are. If you change `OUTPUT_COLUMN_LETTER` to a different letter, remove that same
letter from `INPUT_LETTER_POOL` to avoid it appearing twice.

### F-column fill behavior

`build_truth_table_rows(input_count, f_value)` fills every data row's last column with
the single value (`"1"` or `"0"`) chosen on the F-value screen for that heading. There
is no per-row override in the GUI; to set individual rows differently, edit the table
directly in Word after generation.

### Fonts

Each of the four text roles (fixed heading, gate heading, body, table) has its own
`*_FONT_NAME` / `*_FONT_SIZE_PT` constant, matching the **Formatting Reference** table
above. `GATE_HEADING_BOLD`, `FIXED_HEADING_BOLD`, and `TABLE_BOLD` control bold
specifically; setting any of them to `None` leaves bold exactly as the underlying Word
style defines it. The chosen font must be installed on whichever computer later opens
the document, or Word will substitute a similar one.

### Cover page break

The first heading of the document is given Word's "Page break before" property inside
`build_document_content()` (the `start_on_new_page` argument of `add_heading_block()`).
This keeps everything off the cover page and starts the report on page 2. It is not
controlled by a constant.

### Auto-close delay

After a successful save the window closes after `AUTO_CLOSE_DELAY_MS` milliseconds (long
enough to see the `Saved: <path>` message). Set it to `0` to close immediately. On
failure the window always stays open.

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

### GUI color theme

`COLOR_BG`, `COLOR_CARD_BG`, `COLOR_TEXT`, `COLOR_PRIMARY`, `COLOR_PRIMARY_DARK`, and
`COLOR_PRIMARY_TEXT` are applied once, in `App._apply_theme()`, across the whole window
(background, buttons, entries/spinboxes/listbox, radio/check buttons, and the
"Step X of N" progress bar). Edit the hex values to change the palette; `ERROR_COLOR`
and `OK_COLOR` for warnings/success stay separate on purpose, so problems remain
readable regardless of the chosen theme.

---

## 6. Advanced / Dependency Reference

**For Arch Linux and advanced/custom setups.**
Raw dependencies for people who want to manage them manually. This is **not** part of the
normal install path — normal users should skip this section and use Section 1.

### Raw dependencies

| Dependency | Notes |
|------------|-------|
| Python 3.9 or newer | Standard-library modules used: `re`, `shutil`, `sys`, `pathlib`, `tkinter` |
| Tcl/Tk with the `tkinter` module (including `ttk`) | Usually a separate system package on Linux. The app uses ttk's built-in `"clam"` theme as the base for its color palette |
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