# Academic Document Format Engine (Word VBA Add-in)

A simple, deterministic Microsoft Word (Windows) Add-in built to help students format essays, reports, and theses while reducing the time spent fighting Word's built-in styles and finicky Table of Contents tools. 

It doesn't use AI inside Word, so it runs strictly on hard-coded rules, regexes, and bookmarks. If your text follows the syntax rules in INSTRUCTION.md, it formats everything predictably in one click. If something looks ambiguous, it leaves it alone or highlights it so you can check it yourself.

## What It Actually Does

- **Consistent Essay & Thesis Styling:** Applies your chosen font, size, line spacing, and first-line indent to body paragraphs, while formatting Headings (Levels 1–4), Table/Figure captions, and Source/Note lines according to strict structural rules.
- **Deterministic TOC, LOT & LOF:** Generates Table of Contents, List of Tables, and List of Figures using internal bookmarks and hyperlinks, syncing page numbers (including Roman numerals for front matter). Includes a **Quick Update** button to refresh lists later if you need to edit the content.
- **LaTeX-to-Word Math Converter:** Turns raw LaTeX (`$...$`, `$$...$$`, `\(...\)`, `\[...\]`, basic matrices, `align`, `cases`, etc.) copied from chatbots into native Word equations (`OMath`). It shades converted math in light colors so you can review them, and lets you toggle back to raw LaTeX to fix typos.
- **Single-Step Undo:** Everything runs inside one Word Undo record—if you don't like the result, hit **`Ctrl + Z` once** to revert the whole document.
- **Automatic Crash Backup:** Before modifying anything, the tool automatically creates a physical backup copy (`_crash_backup_...`) in the same folder. If everything finishes cleanly, the backup deletes itself; if Word crashes or hits an unexpected error, the backup stays untouched on your drive.

## How to Install & Uninstall (Windows Only)

1. Go to the **[Releases](../../releases)** page and download **`AcademicFormatter_v1.0.2.zip`**.
2. Extract the zip folder anywhere on your computer.
3. Close all open Microsoft Word windows.
4. Double-click **`install-or-update.bat`** (this checks GitHub for a newer `.dotm` when online, falls back to the local file if offline/blocked, unblocks Windows security flags, and copies `AcademicFormatter.dotm` into your Word `STARTUP` folder).
5. Open Word. You'll see a new **Formatter Macro** tab on the Ribbon.

To remove the add-in at any time, close all Word windows and double-click **`Uninstall.bat`**.

*(Note: Only works on Microsoft Word Desktop for Windows. Won't work on Mac or Word Online).*

## Using AI to Prep Your Draft

Because this tool is strictly rule-based (e.g., `1.1. Heading` needs that trailing dot after the numbers, and markers like `=contentStart()` must sit on their own lines), feeding it messy formatting will make it skip things.

To save time, feed **`INSTRUCTION.md`** (included in the zip) to ChatGPT, Claude, or Gemini first, paste your draft, and ask it to clean up the headings, captions, and LaTeX to match the rules. When pasting the result into Word, always use **`Ctrl + Shift + V`** (*Keep Text Only*) so Word doesn't secretly escape `$` signs into `\$`.

## Sample Files Included

If you want to see how it works before touching your actual paper, test it on the two sample files in the release folder:
* `Sample_Thesis.docx` — A sample essay/thesis layout with markers, 4 heading levels, tables, figures, and citations.
* `LaTeX_Math_Benchmark.docx` — A test sheet with 74 test cases (80+ math formulas, currency edge cases, and unsupported syntax checks).

## Shout-Outs

Massive shout-out to **Gemini** and **Antigravity** for helping architect, write, debug, and stress-test the entire VBA codebase from scratch.

## License & Caveats

Shared free and open-source under the **MIT License**. This is a personal hobby project shared as-is—I don't do regular maintenance or tech support.