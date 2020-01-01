<!-- AF-INSTRUCTION-SIGNATURE: AcademicFormatter -->
# Academic Document Format Engine, LaTeX Math & AI Citation Patcher (v1.0.7) — Deterministic Reference Manual & AI System Prompt

> **Target Audience:**
> 1. **Large-Context AI Assistants (ChatGPT, Claude, Gemini, Copilot):** Use this document as your authoritative system prompt and deterministic grammar specification when helping users prepare, normalize, audit APA citations/acronyms, or troubleshoot academic manuscripts for the **Formatter Macro (`AcademicFormatter.dotm`)** in Microsoft Word.
> 2. **End-Users (Authors, Researchers, Students):** Use this manual as a complete self-service reference for installation, manuscript preparation, Ribbon controls, color-coded visual diagnostics, and instant troubleshooting.

---

## Part 0: AI Assistant System Directive & Guardrails

If an end-user provides you with this document and asks you to format, normalize, or troubleshoot their draft for the **Formatter Macro**, you MUST obey these strict operational invariants:

1. **Your Role:** You are a **Deterministic Pre-Processor & Technical Troubleshooter** for the Word VBA **Academic Document Format Engine (v1.0.7)**. The macro formats paragraphs, headings, tables, figures, lists, citations/acronyms, custom `@STYLE` profiles, `@TAG` classifications, and LaTeX equations by scanning **plain-text structural patterns**, **explicit `@TAG` bookmarks**, and **boundary markers**—NOT Word Styles.
2. **Strictly Forbidden Recommendations:**
   - **NEVER** tell the user to apply Word's built-in Heading Styles (`Heading 1`, `Heading 2`, etc.).
   - **NEVER** tell the user to use Word's `References -> Insert Caption` or `References -> Table of Contents` (built-in TOC fields).
   - **NEVER** tell the user to use Word's `Insert -> Equation` manually when they can write standard LaTeX delimiters (`$...$`, `\( ... \)`, `$$...$$`, `\[...\]`).
   - **NEVER** output escaped dollar signs (`\$`) when writing LaTeX math, and **NEVER** wrap inline formulas in backticks inside prose meant to be pasted into Word.
3. **Mandatory Output Format for Draft Normalization:**
   - Always output the normalized manuscript or section inside a **single fenced `text` code block** (` ```text ... ``` `) so the user can click **Copy** without inheriting Markdown Rich-Text formatting or backslash-escaped `\$` characters.
   - Always remind the user at the end of your response:
     > **Paste Instruction:** In Microsoft Word, always paste using **`Ctrl + Shift + V`** (*Keep Text Only* / *Chỉ giữ văn bản*) so browser HTML/Markdown styling does not corrupt paragraph breaks or escape `$` signs into `\$`.
4. **Strict Grammar Enforcement:**
   - Ensure every Level 2–4 subheading ends its numeric prefix with a **mandatory trailing dot + space** (e.g., `1.1. Research Scope`, `2.3.1. Data Collection` — NEVER `1.1 Research Scope`).
   - Ensure every Table and Figure caption has a **mandatory separator** (`:`, `. `, or ` - `) between the number token and the title text (e.g., `Table 1.1: Revenue Growth` or `Hình 2.1. Sơ đồ hệ thống`).
   - Ensure every Source/Note line starts with an exact keyword followed immediately by a **colon** (`Source:`, `Note:`, `Nguồn:`, `Ghi chú:`).
   - Ensure every LaTeX expression uses only the supported environments, commands, and symbols listed in **Part 2.4**. Rewrite unsupported LaTeX packages/macros (e.g., `\text{}` inside math is supported, but custom `\newcommand`, `\frac12` without braces, or `array` with vertical borders `{|c|c|}`) into supported equivalents before handing the text to the user.

---

## Part 1: Pre-Flight Requirements & Data Safety

### 1.1. Installation, Auto-Update & Uninstallation (One-Click or Manual)
- **Supported Platform:** Microsoft Word for Windows Desktop (**Word 2013 / v15.0 or newer**, including Word 2016, 2019, 2021, 2024, and Microsoft 365; both 32-bit and 64-bit). *Not supported on macOS, Word Online (Web), or mobile.*
- **Method A — Smart Auto-Installer & Updater (`install-or-update.bat`):**
  1. **Close ALL Microsoft Word windows** (the script explicitly checks for running `WINWORD.EXE` and exits with code `4` if open, preventing file locks and preventing Word from re-writing `DisabledItems` on exit).
  2. Double-click **`install-or-update.bat`** (a standalone self-contained bilingual `.bat` + PowerShell script; can sit alongside `AcademicFormatter.dotm` or run directly when online).
  3. **What the installer automatically verifies & heals:**
     - **Online Update + Verified Offline Fallback:** Checks GitHub (`releases/latest` -> `raw/main`, TLS 1.2, 5 s timeout) and compares **SHA-256** against the local `.dotm`. If offline, firewalled, or security-blocked, cleanly falls back to the local `AcademicFormatter.dotm`.
     - **OOXML & ZIP Integrity Gate (`Test-DotmIntegrity`, 50 KB – 10 MB):** Validates `PK` ZIP magic bytes, checks `word/vbaProject.bin`, and parses `docProps/core.xml` + `customUI/customUI14.xml` (`XmlResolver = $null`) so corrupt or HTML-intercepted downloads never overwrite a working local file.
     - **Custom `STARTUP` Detection & Post-Copy Verification:** Resolves Word `>= v15`, reads any custom Office Registry `STARTUP-PATH` (falling back to `%APPDATA%\Microsoft\Word\STARTUP`), copies `AcademicFormatter.dotm`, strips Windows *Mark of the Web* (`Zone.Identifier` via `Unblock-File`), and re-verifies destination **SHA-256 + XML integrity** (`Exit Code 6` alerts immediately if Windows Defender quarantined the file).
     - **Self-Healing Word `Resiliency\DisabledItems` & Macro Check:** Automatically purges `AcademicFormatter.dotm` entries from `HKCU\...\Word\Resiliency\{DisabledItems,StartupItems}` across Office `12.0–16.0` before and after install, and warns if Word Trust Center `VBAWarnings = 4` (*Disable all macros without notification*) is active.
     - **Automatic `INSTRUCTION.md` Sync (with Foreign-File Protection):** Right after installing `.dotm`, the installer checks GitHub for the latest `INSTRUCTION.md` (validated via `Test-InstructionIntegrity`: `5 KB–500 KB`, non-HTML, and verified `AF-INSTRUCTION-SIGNATURE: AcademicFormatter` / legacy signature) in the installer folder (`$ScriptDir` only). If no `INSTRUCTION.md` exists or it belongs to AcademicFormatter, it atomically updates `INSTRUCTION.md` via UTF-8 no-BOM bytes + SHA-256 check; if an unrelated `INSTRUCTION.md` from another project is present, it leaves that file untouched and writes `AcademicFormatter_INSTRUCTION.md` instead (printing the exact full path to share with your AI; fails safely offline without failing installation).
- **Method B — Manual Copy:**
  1. Close Microsoft Word.
  2. Press `Win + R`, type `%APPDATA%\Microsoft\Word\STARTUP`, and press `Enter`.
  3. Copy `AcademicFormatter.dotm` into that folder.
  4. Right-click `AcademicFormatter.dotm` -> **Properties** -> if **Unblock** appears at the bottom of the *General* tab, check **Unblock** -> **OK**.
- **Uninstallation (`Uninstall.bat`):**
  1. Close all open Microsoft Word windows.
  2. Double-click **`Uninstall.bat`** to scan every Registry `STARTUP-PATH` and `%APPDATA%\Microsoft\Word\STARTUP`, delete `AcademicFormatter.dotm` / `.dotm.bak`, and clean `Resiliency\{DisabledItems,StartupItems}` keys (or delete `AcademicFormatter.dotm` manually from `%APPDATA%\Microsoft\Word\STARTUP`).

### 1.2. The 5 Pre-Flight Safety Gates
Before `Format All`, `Update Lists`, `Build Math`, or `Apply AI Patch` modifies a single character in your document, it enforces 5 deterministic pre-flight checks:

| Gate # | Condition Checked                                          | Why It Exists                                                                                                                                                               | Exact User Action Required if Blocked                                                                                                              |
| :----- | :--------------------------------------------------------- | :-------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | :------------------------------------------------------------------------------------------------------------------------------------------------- |
| **1**  | **Saved to Disk (`doc.Path <> ""`)**                       | Unsaved documents (`Document1`) have no folder for physical crash backup.                                                                                                   | Press **`Ctrl + S`** and save the `.docx` file to a local folder first.                                                                            |
| **2**  | **Local File Path (No `http://` or `https://` Cloud URL)** | When OneDrive or SharePoint opens a file as a remote URL moniker (`https://d.docs.live.net/...`), Windows FileSystemObject cannot write a local atomic backup alongside it. | Use `File -> Save a Copy` to a local disk folder (e.g., your local `Documents` or `Desktop` folder) or pause OneDrive cloud-only path redirection. |
| **3**  | **Not Read-Only (`doc.ReadOnly = False`)**                 | A Read-Only document cannot be formatted or saved.                                                                                                                          | Close the file, remove the Read-Only attribute (or click *Enable Editing*), and reopen.                                                            |
| **4**  | **Not Protected (`doc.ProtectionType = -1`)**              | Password/form-protected documents reject DOM range mutations.                                                                                                               | Go to `Review -> Protect -> Restrict Editing` and click **Stop Protection**.                                                                       |
| **5**  | **No Unsaved Edits (`doc.Saved = True`)**                  | Ensures the disk file copied for the physical backup is 100% identical to the in-memory document state before formatting starts.                                            | Press **`Ctrl + S`** right before clicking `Format All` or `Update Lists`.                                                                         |

### 1.3. Track Changes Handling, Physical Crash Backup & Master Undo
- **Track Changes (`TrackRevisions`):** If *Track Changes* is active, formatting thousands of paragraphs generates massive revision bloat and slows Word COM by 10x–50x. The engine prompts you to temporarily pause *Track Changes* during execution and **automatically restores** its state upon completion.
- **Ephemeral Physical Crash Backup (`Format All`, `Build Math`, `Apply AI Patch`):**
  - Before execution touches the document, it creates a physical copy on disk in the same folder:
    `<OriginalFileName>_crash_backup_YYYYMMDD_HHNNSS.docx`
  - **Normal Completion (Zero Silent Errors):** The temporary backup file is **automatically deleted** as soon as formatting finishes cleanly, leaving your folder clutter-free.
  - **Runtime Error or Minor Warning (`g_SilentErrors > 0`):** The backup file is **preserved on disk** and its exact path is displayed in the completion/error dialog.
- **Single-Step Master Undo (`Ctrl + Z`):**
  - Every major action (`Format All`, `Update Lists`, `Build Math`, `Finalize Math`, `Clean Markers`, `Apply AI Patch`) wraps its entire execution inside a single Word `UndoRecord`. Pressing **`Ctrl + Z` once** reverts the entire operation in one step.
  - *`@STYLE` Persistence Note:* `Ctrl + Z` reverts all document text formatting and `@TAG` bookmarks, but Word `Document.Variables` (where `@STYLE` profiles are stored) live outside Word's Undo stack — to clear a `@STYLE` profile after testing, click **`AI Patch -> Reset Custom Style (Về mặc định)`** on the Ribbon.

---

## Part 2: Deterministic Formatting & Syntax Grammar ("How to Write for 100% First-Pass Accuracy")

The engine does **not** guess intent from fonts or visual styling. It classifies every paragraph using strict, case-insensitive regular expressions. Follow the grammar below for 100% recognition on the first run.

### 2.1. Boundary & List Markers (The 6 System Markers)
Markers tell the engine where your main body starts/ends and where to generate lists.

| Marker Token (Case-Insensitive) | Required / Optional | Function                                                                                                                                                                             |
| :------------------------------ | :------------------ | :----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `=listContent()`                | Optional            | Generates the **Table of Contents (TOC / MỤC LỤC)** above this marker.                                                                                                               |
| `=listTableHere()`              | Optional            | Generates the **List of Tables (LOT / DANH MỤC BẢNG BIỂU)** above this marker (auto-skipped if the document has 0 tables).                                                           |
| `=listFigureHere()`             | Optional            | Generates the **List of Figures (LOF / DANH MỤC HÌNH ẢNH)** above this marker (auto-skipped if the document has 0 figures).                                                          |
| `=listAcronyms()`               | Optional            | Anchor where **`Apply AI Patch`** generates/rebuilds the **List of Acronyms (LOA / DANH MỤC TỪ VIẾT TẮT)** immediately below this marker (bookmark `bm_loa_header` + `bm_zone_loa`). |
| `=contentStart()`               | **MANDATORY**       | Marks the **beginning of the Main Body**. Placed immediately before your first Chapter / Level 1 heading of the body.                                                                |
| `=contentEnd()`                 | **MANDATORY**       | Marks the **end of the Main Body**. Placed immediately after your final body paragraph (typically right before `References` / `Tài liệu tham khảo` or `Appendix` / `Phụ lục`).       |

#### Strict Marker Rules:
1. **Standalone Paragraph Rule:** Every marker MUST sit alone on its own paragraph (created with a hard `Enter`, not `Shift + Enter`). Nothing else may appear on that line.
2. **No Table Rule:** Markers MUST NOT be placed inside a Word Table cell.
3. **No Duplicates:** `=contentStart()`, `=contentEnd()`, and each `=list...()` marker must appear **at most once** (`=contentStart()` and `=contentEnd()` are mandatory).
4. **Strict Ordering Rule:** All list markers (`=listContent()`, `=listTableHere()`, `=listFigureHere()`, `=listAcronyms()`) MUST be placed **before** `=contentStart()`. Placing any list marker between `=contentStart()` and `=contentEnd()` triggers a **Marker Warning** and safely aborts execution:
   ```text
   [Cover Page / Title Page]
   [Acknowledgements / Abstract]
   =listContent()
   =listTableHere()
   =listFigureHere()
   =listAcronyms()
   =contentStart()
   [Main Body: Chapters 1..N, Tables, Figures, Equations]
   =contentEnd()
   [References / Tài liệu tham khảo]
   [Appendices / Phụ lục]
   ```

---

### 2.2. Headings Grammar (Levels 1 to 4)

#### A. Level 1 Headings (`ElementTypeLevel1`)
A paragraph is recognized as **Level 1** (Formatted as **UPPERCASE, Bold, Centered**, `FirstLineIndent = 0`, `LeftIndent = 0`, `KeepWithNext = True`; preserves the document's existing `SpaceBefore` / `SpaceAfter`) if its trimmed text matches any of the following **case-insensitive** patterns:

| Category                                                                         | Exact English Keywords / Patterns                                                                                                                                                                                                                | Exact Vietnamese Keywords / Patterns                                                                                                                                                                                                       |
| :------------------------------------------------------------------------------- | :----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | :----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| **Front Matter**                                                                 | `Abstract`<br>`Acknowledgment`, `Acknowledgments`, `Acknowledgement`, `Acknowledgements`<br>`Declaration`, `Declaration of Originality`<br>`Acronym`, `Acronyms`, `List of Acronyms`<br>`Abbreviation`, `Abbreviations`, `List of Abbreviations` | `Lời cảm ơn`, `Lời cám ơn`<br>`Lời cam đoan`<br>`Tóm tắt`<br>`Danh mục từ viết tắt`, `Danh mục các từ viết tắt`, `Danh mục chữ viết tắt`, `Danh mục các chữ viết tắt`, `Từ viết tắt`, `Chữ viết tắt`<br>`Danh mục ký hiệu và chữ viết tắt` |
| **Intro & Conclusion**                                                           | `Introduction`<br>`Conclusion`, `Conclusions`                                                                                                                                                                                                    | `Mở đầu`, `Lời mở đầu`, `Giới thiệu`<br>`Kết luận`, `Kết luận và kiến nghị`                                                                                                                                                                |
| **Numbered Chapters**                                                            | `Chapter <ID>: <Title>` or `Chapter <ID>. <Title>`<br>Upper Roman Level 1 (`I. <Title>`..`XXXIX. <Title>`)<br>Example: `Chapter 1: Literature Review`, `II. Methodology`                                                                         | `Chương <ID>: <Title>`, `Chương <ID>. <Title>`, `Phần <ID>: <Title>`<br>Số La Mã viết hoa (`I. Đặt vấn đề`, `II. Nội dung`)<br>Example: `Chương 1: Cơ sở lý luận`, `Phần II: Thực nghiệm`                                                  |
| **References**                                                                   | `Reference`, `References`<br>`Bibliography`                                                                                                                                                                                                      | `Tài liệu tham khảo`<br>`Danh mục tài liệu tham khảo`                                                                                                                                                                                      |
| **Appendices**                                                                   | `Appendix`, `Appendices`, `Appendixes`<br>`Appendix <ID>` (e.g., `Appendix A`, `Appendix 1`)<br>`Appendix <ID>: <Title>` or `Appendix <ID>. <Title>`                                                                                             | `Phụ lục`<br>`Phụ lục <ID>` (e.g., `Phụ lục A`, `Phụ lục 1`)<br>`Phụ lục <ID>: <Title>` or `Phụ lục <ID>. <Title>`                                                                                                                         |
| **Manual List Headers** *(when the corresponding `=list...()` marker is absent)* | `Table of Contents`, `Tables of Contents`<br>`List of Table`, `List of Tables`<br>`List of Figure`, `List of Figures`                                                                                                                            | `Mục lục`<br>`Danh mục bảng`, `Danh mục bảng biểu`, `Danh sách bảng`, `Danh sách bảng biểu`<br>`Danh mục hình`, `Danh mục hình ảnh`, `Danh mục hình vẽ`, `Danh sách hình`, `Danh sách hình ảnh`, `Danh sách hình vẽ`, `Danh mục sơ đồ`     |

> **Manual List Header Safety Rule:** If your document contains a manual heading like `MỤC LỤC` or `TABLE OF CONTENTS` **and** also has the `=listContent()` marker, the engine highlights the manual header in **Bright Red (`wdRed`)** and excludes it from the TOC so you don't get duplicate headers.

#### B. Levels 2, 3, and 4 Subheadings (`ElementTypeLevel2..4` — Main Body Only)
Subheadings are recognized **only inside the Main Body** (between `=contentStart()` and `=contentEnd()`).

> **CRITICAL RULE — THE TRAILING DOT:** Every numeric subheading prefix **MUST end with a dot (`.`) followed by at least one space and the title text**. Missing the final dot is the #1 user mistake!

| Level       | Exact Regex Pattern               | Formatting Applied                                                                                                                                                  | Valid Examples (`PASS`)                                     | Invalid Examples (`FAIL` -> Treated as Normal Paragraph)                                            |
| :---------- | :-------------------------------- | :------------------------------------------------------------------------------------------------------------------------------------------------------------------ | :---------------------------------------------------------- | :-------------------------------------------------------------------------------------------------- |
| **Level 2** | `^(\d+\.\d+\.)\s+(.+)$`           | **Bold**, Left-aligned, First-line indent (`1.27 cm` default), **`LeftIndent = 0.5 cm`**, `KeepWithNext = True`, preserves existing spacing, Auto Title-Case letter  | `1.1. Research Background`<br>`2.10. Kết quả thực nghiệm`   | `1.1 Research Background` *(missing dot after `.1`)*<br>`1.1.` *(no title text)*                    |
| **Level 3** | `^(\d+\.\d+\.\d+\.)\s+(.+)$`      | **Bold + Italic**, Left-aligned, First-line indent (`1.27 cm` default), **`LeftIndent = 0.75 cm`**, `KeepWithNext = True`, preserves existing spacing               | `1.1.1. Theoretical Framework`<br>`3.2.1. Mô hình đề xuất`  | `1.1.1 Theoretical Framework` *(missing final dot)*                                                 |
| **Level 4** | `^(\d+\.\d+\.\d+\.\d+\.)\s+(.+)$` | *Italic* (Not Bold), Left-aligned, First-line indent (`1.27 cm` default), **`LeftIndent = 1.0 cm`**, `KeepWithNext = True`, preserves existing spacing              | `1.1.1.1. Primary Survey`<br>`2.1.3.1. Phân tích định tính` | `1.1.1.1 Primary Survey` *(missing final dot)*<br>`1.1.1.1.1. Too Deep` *(>4 levels not supported)* |

> **Orphan Prevention (`KeepWithNext`) & Paragraph Spacing:** `KeepWithNext = True` is automatically enabled on Headings Level 1–4, Table Captions, Inline Pictures, and generated List Headers (`TOC`/`LOT`/`LOF`) so titles are never orphaned at the bottom of a page (`KeepWithNext = False` on Normal Paragraphs, Source/Note lines, and Figure Captions). Existing `SpaceBefore` and `SpaceAfter` values of your document are intentionally preserved rather than overwritten.

> **Word Navigation Pane & PDF Bookmarks (`ParagraphFormat.OutlineLevel`):** Without replacing paragraph styles with Word built-in `Heading 1..4` styles (keeping your document's clean typography intact), the engine automatically sets `OutlineLevel = 1..4` on recognized headings and `OutlineLevel = 1` on generated list headers (`MỤC LỤC`, `DANH MỤC BẢNG BIỂU`, `DANH MỤC HÌNH ẢNH`, `DANH MỤC TỪ VIẾT TẮT`) — syncing Level 4 to Body Text (`10`) when Ribbon **Depth** is set to `Level 1-3`. Normal paragraphs, captions, source/note lines, duplicate manual headers, and TOC entries are explicitly reset to `OutlineLevel = 10` (`Body Text`) so phantom outlines from copy-paste never pollute Word's Navigation Pane or exported PDF bookmarks.

#### C. Soft Line Breaks (`Shift + Enter` / `^l`) Inside Headings & Captions
Before scanning, `NormalizeContentBreaks` inspects every paragraph containing a soft line break (`Shift + Enter` / `Chr(11)`):
- **Case 1 — Split Prefix (`Chapter 1:` + `Shift+Enter` + `Introduction`):** If the first line before `Shift + Enter` is **only a prefix** (e.g., `Chapter 1:`, `Chương 1.`, `1.1.`, `Table 2.1.`, `Figure 3.1:`, `Source:`), the engine intentionally **keeps the clean multi-line visual break intact in the body document** and automatically joins the lines with a single space inside generated TOC / LOT / LOF entries (`Chapter 1: Introduction`).
- **Case 2 — Heading + Body Merged by `Shift + Enter` (`1.1. Overview` + `Shift+Enter` + `This section discusses...`):** If the first line is already a complete heading/caption and the second line is body prose, the engine converts `Shift + Enter` into a **hard paragraph break (`Enter` / `vbCr`)** so the heading and body paragraph are separated and formatted independently!

---

### 2.3. Table Captions, Figure Captions & Source/Note Lines

Recognized inside both the **Main Body** (`=contentStart()` to `=contentEnd()`) and the **Post-Content / Appendix Zone** (after `=contentEnd()`).

#### A. Table & Figure Captions (`ElementTypeTableTitle`, `ElementTypeFigureTitle`)
A paragraph is recognized as a Table or Figure caption **only if** it matches:
`^<Keyword>\s+<NumToken><Separator><TitleText>$`

1. **`<Keyword>` (Case-Insensitive):**
   - **Table:** `Table`, `Bảng`, `Bang` (accepts unaccented `Bang`).
   - **Figure:** `Figure`, `Fig.`, `Fig`, `Chart`, `Diagram`, `Hình`, `Hinh`, `Hình ảnh`, `Hinh ảnh`, `Hình vẽ`, `Hinh vẽ`, `Sơ đồ`, `Biểu đồ`, `Đồ thị`.
2. **`<NumToken>`:** Alphanumeric segments joined by `.`, `-`, or `_` (Regex: `[A-Za-z0-9]+(?:[\.\-_][A-Za-z0-9]+)*`).
   - Valid: `1`, `1.1`, `2.14`, `A.1`, `B-2`, `3_1`, `4.2a`.
3. **`<Separator>` (MANDATORY):** Must be one of:
   - A **colon** `:` with optional spaces around it (`: `, ` : `).
   - A **period** `.` followed by **at least one space** (`. `).
   - A **hyphen / en-dash / em-dash** (`-`, `–`, `—`) surrounded by **at least one space on both sides** (` - `, ` – `, ` — `).
4. **Why This Separator Rule Matters:**
   - A body sentence like `"Table 4.2 presents the regression results."` has **no** `:`, `. `, or ` - ` after `4.2`. The engine deterministically classifies it as a **Normal Paragraph**, never a false-positive Table Caption!
5. **Formatting Applied:**
   - **Table Captions (`ElementTypeTableTitle`):** **Bold** (not Italic), Left-aligned with First-line indent (`1.27 cm` default), `KeepWithNext = True` (stays with table below), preserves existing spacing, first letter of `<TitleText>` auto-capitalized.
   - **Figure Captions (`ElementTypeFigureTitle`):** **Bold** (not Italic), Centered (`FirstLineIndent = 0`), `KeepWithNext = False`, preserves existing spacing, first letter of `<TitleText>` auto-capitalized.
6. **Layout-Table Exception for Figures:** Normally, paragraphs inside Word Table cells are formatted as `TableCell` and excluded from caption lists. **Exception:** If you place side-by-side images inside a layout table and put a valid Figure caption (`Figure 1.1: ...`) inside a cell of that same table (where `Table.Range.InlineShapes.Count > 0`), the engine recognizes and indexes the Figure caption into the List of Figures!
7. **Length Limit in Quick Update:** Captions up to `350` raw characters are scanned in high-speed Quick Update mode. Keep captions concise and put source citations on the next paragraph (`Source:` / `Nguồn:`).

#### B. Source & Note Lines (`ElementTypeSourceNote`)
- **Exact Regex:** `^(Source|Nguồn|Note|Ghi chú):\s*(.*)$` (Case-insensitive; **colon `:` is mandatory**).
- **Formatting Applied:**
  - Whole line: Same font size as body text (`BodySize`), Left-aligned, no indent (`FirstLineIndent = 0`, `LeftIndent = 0`), `KeepWithNext = False`, preserves existing spacing.
  - Prefix (`Source:`, `Nguồn:`, `Note:`, `Ghi chú:`): **Bold** (not Italic); content after colon: *Italic* (not Bold), first letter after colon auto-capitalized.
- **Common Mistake:** Writing `Nguồn (2025)` or `Note - Author calculation` without a colon `:` right after the keyword will cause it to be formatted as a Normal Paragraph. Always write `Nguồn: ...` or `Note: ...`.

---

### 2.4. LaTeX Math Engine Grammar (v1.0.7)

The built-in pure-VBA LaTeX-to-OMML compiler (`Build Math` / `Dựng Toán`) converts raw LaTeX expressions directly into native Microsoft Word `OMath` equations (`Cambria Math`) without external tools.

#### A. Supported Delimiters
| Mode                                | Delimiters                                                                                                                                           | Behavior & Placement                                                                                                                                                                                       |
| :---------------------------------- | :--------------------------------------------------------------------------------------------------------------------------------------------------- | :--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| **Inline Math**                     | `$...$` or `\(...\)`                                                                                                                                 | Embedded inside a paragraph (up to `250` chars per `$...$` span). Rendered inline and shaded **Pastel Cyan** (`#DCEBF5`) for review.                                                                       |
| **Display Block (Single-Line)**     | `$$...$$` or `\[...\]` or `\begin{env}...\end{env}`                                                                                                  | Occupies an entire paragraph (or a table cell). Rendered as a centered display equation and shaded **Pastel Mint** (`#DCF5E1`).                                                                            |
| **Display Block (Multi-Paragraph)** | Opening `$$`, `\[`, or `\begin{env}` on one paragraph, closing `$$`, `\]`, or `\end{env}` up to **40 paragraphs** later (`MATH_MERGE_MAX_SPAN = 40`) | Automatically merges the paragraphs into a single centered display `OMath` block. Never crosses `=contentEnd()` or swallows subsequent closed single-line `$$...$$` blocks if a closing tag was forgotten. |

#### B. Smart `$...$` Disambiguation (Pandoc Dollar Rules & Currency Guard)
To prevent prose dollar signs and financial figures from being corrupted into equations, the scanner classifies every `$...$` span into one of 3 states:

1. **Ignored as Plain Text / Currency (`Class 0` — Untouched, No Shading):**
   - Leading/trailing spaces inside `$ ... $` when no math tokens (`\`, `_`, `^`, `=`) are present.
   - Closing `$` immediately followed by a digit (`$a$5`).
   - **Recognized Currency / Unit Expressions (`IsCurrencyText`):** Starts with a number (`0–9`, `,`, `.`, spaces) followed by any supported currency or rate unit (and contains no `\`, `=`, `{`, `}`, `^`, `_`):
     * *Currency/Magnitude Units:* `usd`, `vnd`, `eur`, `gbp`, `jpy`, `cny`, `chf`, `k`, `m`, `b`, `tr`, `bn`, `dong`, `d`, `đồng`, `đ`, `nghin`, `ngan`, `nghìn`, `ngàn`, `trieu`, `triệu`, `ty`, `tỷ`, `dollar`, `dollars`, `cent`, `cents`, `thousand`, `million`, `billion`.
     * *Period/Rate Units (after `/`, e.g., `$1,200/month`, `$15.50/mo`, or `$500 đồng/đơn vị`):* `month`, `mo`, `year`, `yr`, `day`, `unit`, `thang`, `tháng`, `nam`, `năm`, `nguoi`, `người`, `don vi`, `đơn vị`.
     * *Examples left 100% untouched:* `$500 triệu`, `$10 tỷ`, `$250 nghìn`, `$1,200/tháng`, `$50 million`, `$100 USD`.
2. **Valid Math (`Class 1` — Built into Word OMath):**
   - Balanced braces/parentheses, `<= 250` chars, not pure digits, and either has no spaces after a leading digit or contains explicit math proof (`\`, `=`, `^`, `_`).
   - *Examples:* `$x^2 + y^2 = z^2$`, `$100 \times 25\% = 25$`, `$g = \frac{Y_t - Y_{t-1}}{Y_{t-1}}$`.
3. **Ambiguous / Malformed (`Class 2` — Highlighted Pastel Pink `#FFE6E6`, Not Converted):**
   - Pure numbers in dollar signs (`$500$`, `$12.5%$`), number + unknown word with spaces and no math operators (`$20 MT$`), unbalanced `{}`/`()`, or length `> 250` chars.
   - *Tip:* If `$500$` is genuinely a math constant, place the cursor inside the Pink highlight and click **`Toggle / Edit (Đổi / Sửa)`** once to force-build it into an `OMath`!

#### C. Supported LaTeX Environments (`ModTexParserEnvs.bas`)
| Category                     | Exact Environment Names                                                                                                                                        | Rendering Details                                                                                                                                                                                                                                                  |
| :--------------------------- | :------------------------------------------------------------------------------------------------------------------------------------------------------------- | :----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| **Piecewise / Cases**        | `cases`                                                                                                                                                        | Left curly brace `{` with 2 left-aligned columns (`&` separates columns, `\\` separates rows).                                                                                                                                                                     |
| **Multi-Line Alignment**     | `align`, `align*`, `aligned`, `split`                                                                                                                          | Alternating `right, left, right, left` column alignment across `&` separators.                                                                                                                                                                                     |
| **Centered Multi-Line**      | `gathered`, `multline`, `multline*`, `equation`, `equation*`                                                                                                   | Centered multi-line or single-equation display blocks.                                                                                                                                                                                                             |
| **Matrices & Determinants**  | `matrix` *(no brackets)*<br>`pmatrix` `( ... )`<br>`bmatrix` `[ ... ]`<br>`Bmatrix` `{ ... }`<br>`vmatrix` `\| ... \|`<br>`Vmatrix` `‖ ... ‖`<br>`smallmatrix` | Up to arbitrary rows/columns separated by `&` and `\\`. Supports nested matrices inside `align`/`cases` and matrices >10 rows.                                                                                                                                     |
| **Arrays (Borderless Only)** | `array` with `{l}`, `{c}`, `{r}` column specs (e.g., `\begin{array}{lcc} ... \end{array}`)                                                                     | **Fail-Closed Rule:** If the column spec contains vertical borders `\|` (e.g., `{c\|c}`) or the body contains `\hline`, the engine refuses conversion and shades the block **Pastel Pink** so you can convert it to a real Word Table or remove `\|` and `\hline`. |

#### D. Supported LaTeX Commands & Macros (`ModTexParserCmds.bas`, `ModTexSymbols.bas`)
- **Fractions & Binomials:** `\frac{a}{b}`, `\dfrac{a}{b}`, `\tfrac{a}{b}`, `\cfrac{a}{b}`, `\binom{n}{k}`, `\dbinom{n}{k}`, `\tbinom{n}{k}`.
- **Radicals:** `\sqrt{x}`, `\sqrt[n]{x}`.
- **Subscripts, Superscripts & Primes:** `x_i`, `x_{i,j}`, `x^2`, `x^{a+b}`, `x_i^2`, `f'(x)`, `f''(x)`, `f'_i`, `x'_1`, `f''_n`, `x^\prime`.
- **Big Operators (N-ary with `_` and `^` limits):** `\sum`, `\prod`, `\coprod`, `\int`, `\iint`, `\iiint`, `\oint` (including `\limits` / `\nolimits`, e.g., `\sum\limits_{i=1}^n`).
- **Delimiters (Auto-Scaling):** `\left`, `\right`, `\big`, `\bigl`, `\bigr`, `\Big`, `\Bigl`, `\Bigr`, `\bigg`, `\biggl`, `\biggr`, `\Bigg`, `\Biggl`, `\Biggr` paired with `(`, `)`, `[`, `]`, `\{`, `\}`, `|`, `\|`, `\Vert`, `\langle`, `\rangle`, `\lfloor`, `\rfloor`, `\lceil`, `\rceil`, or `.` (invisible delimiter).
- **Accents, Bars & Over-Vectors:** `\hat{x}`, `\widehat{ABC}`, `\bar{x}`, `\overline{AB}`, `\underline{xy}`, `\vec{v}`, `\overrightarrow{AB}`, `\tilde{x}`, `\dot{x}`, `\ddot{x}`.
- **Underbrace, Overset, Boxed, Labels & Formatting Directives:** `\underbrace{expr}_{note}`, `\overset{top}{base}`, `\xrightarrow{label}`, `\xrightarrow[sub]{sup}`, `\boxed{expr}`, `\tag{1.1}`, `\label{...}` *(safely consumed without breaking conversion)*, `\displaystyle`, `\textstyle`, `\nonumber`, `\notag`.
- **Fonts & Text Spans:** `\text{...}`, `\mbox{...}`, `\textrm{...}`, `\mathrm{...}`, `\mathbf{...}`, `\bm{...}`, `\boldsymbol{...}`, `\mathbb{...}` (e.g., `\mathbb{R}, \mathbb{E}`), `\mathcal{...}` (e.g., `\mathcal{L}, \mathcal{N}`).
- **Named Operators & Limit Functions:** `\sin`, `\cos`, `\tan`, `\cot`, `\sec`, `\csc`, `\arcsin`, `\arccos`, `\arctan`, `\sinh`, `\cosh`, `\tanh`, `\ln`, `\log`, `\exp`, `\lim`, `\min`, `\max`, `\inf`, `\sup`, `\deg`, `\dim`, `\det`, `\ker`, `\hom`, `\arg`, `\argmin`, `\argmax`, `\operatorname{name}`, `\operatorname*{name}_{sub}` (places limits directly underneath).
- **Negations (`\not`):** `\not\in` (`∉`), `\not=` (`≠`), `\not\equiv` (`≢`), `\not\subset` (`⊄`), `\not\subseteq` (`⊈`), `\not\approx` (`≉`), `\not\sim` (`≁`), `\not<` (`≮`), `\not>` (`≯`), `\not\leq` / `\not\le` (`≰`), `\not\geq` / `\not\ge` (`≱`). *(Any unknown `\not\foo` fails closed to Pastel Pink).*
- **Greek Letters:** All lowercase (`\alpha` through `\omega`, plus `\varepsilon`, `\vartheta`, `\varpi`, `\varrho`, `\varsigma`, `\varphi`) and uppercase (`\Gamma`, `\Delta`, `\Theta`, `\Lambda`, `\Xi`, `\Pi`, `\Sigma`, `\Upsilon`, `\Phi`, `\Psi`, `\Omega`).
- **Relations, Arrows & Logic/Set Symbols:** `\leq`, `\le`, `\geq`, `\ge`, `\neq`, `\ne`, `\approx`, `\sim`, `\simeq`, `\cong`, `\equiv`, `\propto`, `\times`, `\div`, `\cdot`, `\pm`, `\mp`, `\in`, `\notin`, `\subset`, `\subseteq`, `\supset`, `\supseteq`, `\cup`, `\cap`, `\setminus`, `\emptyset`, `\varnothing`, `\forall`, `\exists`, `\nexists`, `\nabla`, `\partial`, `\infty`, `\triangle`, `\perp`, `\parallel`, `\Vert`, `\|`, `\mid`, `\circ`, `\oplus`, `\otimes`, `\to`, `\rightarrow`, `\leftarrow`, `\gets`, `\Rightarrow`, `\Leftarrow`, `\Leftrightarrow`, `\ldots`, `\dots`, `\cdots`, `\vdots`, `\ddots`, `\,`, `\:`, `\;`, `\ `, `\!`, `\quad`, `\qquad`.

---

### 2.5. AI-Assisted Citation Hyperlinking, Acronym Normalization, Custom Profile & `AF-PATCH v2` Protocol

This feature follows the **"AI Web Brain + Deterministic VBA Hands"** architecture: the Word add-in never makes external API calls or requires API keys. Instead, you collaborate with any AI Web Chat and apply the deterministic `=== AF-PATCH v2 ===` patch in **1 Click** (`=== AF-PATCH v1 ===` remains 100% backward-compatible).

#### A. Flexible 3-Step AI Web Chat Workflow (Reference Guide — Not Rigid)

> **Flexibility Note:** The 3-step flow below is a **recommended reference workflow**, NOT a rigid gate. Adapt freely to what the user actually asks for:
> - If the user already has clean `References` and only wants `@STYLE` (custom typography), `@TAG` (explicit paragraph classification), or `@STRUCT` fixes, you may skip directly to the relevant step or emit `=== AF-PATCH v2 ===` immediately once any real ambiguity is clarified.
> - When doing a full manuscript review (citations + acronyms + structure + custom style), walking through Steps 1 -> 2 -> 3 prevents silent mismatches and wasted rework.

1. **Step 1 — Content, Structure & Sources Audit (Recommended First Pass):**
   - Scan structural grammar across front matter, main body, and appendices: flag missing trailing dots (`1.1 Research Scope`), missing `:` on chapters/appendices/tables/figures/sources (`Chương 1 Mở đầu`, `Table 1 Revenue`, `Appendix A Data`, `Nguồn (2025)`), or stray colons on standalone headers (`Lời cảm ơn:`) for **`@STRUCT`**.
   - Identify **non-standard headings/captions** that do not match built-in regexes (e.g., `"DANH SÁCH THÀNH VIÊN"`, `"Phần đầu tiên 1- Giới thiệu"`, `"Biểu số 1 - Thống kê"`) or **false-positive lines** that look like headings/captions (e.g., a quoted `"1.1. Sample quote..."` inside prose) for **`@TAG`**.
   - Count and cross-check every in-text citation (body prose, body data tables, and `Source:`/`Nguồn:` lines) against `References` / `Tài liệu tham khảo`. Compare with the user's `Pre-Audit Scan` numbers if supplied. Flag uncited references, missing entries, spelling/year mismatches, or duplicate/scattered per-chapter `References` blocks.
   - List proposed acronym pairs (`ABBR | Full Form`) and ask whether the user wants a **2-column** or **3-column** (`Definition & Source`) table.
   - **Strict Name, Language & Prose Guardrails:**
     - NEVER alter academic prose, sentence structure, wording, or factual content outside the scoped patch operations (`@STYLE`, `@TAG`, `@STRUCT`, `@FIX`, `@CITE`, `@ACRONYMS`). Under `@STRUCT`, ONLY fix the structural prefix/separator (`1.1` -> `1.1.`, `Table 1` -> `Table 1:`, `Nguồn` -> `Nguồn:`); never rewrite the user's prose without asking first.
     - NEVER forcibly Westernize Vietnamese names (`Nguyễn` -> `Nguyen`, stripping diacritics) or Vietnamize English names unless fixing a real mismatch against `References` or explicitly approved by the user.
     - Preserve the manuscript's dominant Vietnamese APA tokens (`và cs.`, `và nnk.`, `k.n.`, `không rõ năm`, `đang in`...); NEVER rewrite them into `et al.` / `n.d.` unless the draft contradicts itself internally and the user confirmed the convention.
     - In `@ACRONYMS` Column 3, NEVER fabricate or add a `Nguồn:` / `Source:` entry that does not exist in `References`.
   - **Strict `Source:` / `Nguồn:` Under Tables & Figures Rule:**
     - **Case 1 — Self-Compilation (NO external reference):** Lines such as `Source: Author's compilation`, `Source: Author's calculation`, `Nguồn: Tác giả tự tổng hợp`, `Nguồn: Nhóm nghiên cứu`, or citing the paper's own author(s) do **NOT** correspond to any `References` entry -> **NEVER** generate `@CITE` or attach a hyperlink.
     - **Case 2 — External Source in `Source:` / `Nguồn:` (IN `References`):** Lines such as `Source: Adapted from Vaswani et al. (2017)` or `Nguồn: World Health Organization (2020)` cite an entry that exists in `References` -> this IS a full APA citation -> **MUST** generate `@CITE` and hyperlink it normally (consistent with `ModCrossRef.bas`, which also supports links inside `ElementTypeSourceNote`).

2. **Step 2 — Clean `References` (If Needed) & Confirm `@STYLE` / `@TAG` Profile:**
   - **If `References` needs repair** (duplicate entries causing `Ambiguous ref (2+)`, multiple per-chapter `References` blocks inside the body that trigger `inRefZone` early and skip later chapters, unsorted entries, or missing `a`/`b` year suffixes):
     - Merge, deduplicate (keep the richest entry with DOI/URL/pages/publisher), sync `a`/`b` suffixes, and sort strictly **A–Z** preserving APA *italics*.
     - Instruct the user in 3 quick steps: (1) delete all scattered per-chapter `References` blocks leaving **only ONE heading** after `=contentEnd()`, (2) paste the cleaned list (`Ctrl + V` + unify font/size + `Ctrl + T` for hanging indent, or `Ctrl + Shift + V` for plain text), and (3) press **`Ctrl + S`** before applying the patch.
   - **Default Engine Typography Summary (Ask if Custom `@STYLE` is Desired):**
     - `L1`: UPPERCASE, Bold, Centered, `LeftIndent=0`, `FirstIndent=0`, `KeepWithNext=1`.
     - `L2`: Bold, Left, `LeftIndent=0.5cm`, `FirstIndent=Ribbon`, `KeepWithNext=1`.
     - `L3`: Bold + Italic, Left, `LeftIndent=0.75cm`, `FirstIndent=Ribbon`, `KeepWithNext=1`.
     - `L4`: Italic, Left, `LeftIndent=1.0cm`, `FirstIndent=Ribbon`, `KeepWithNext=1`.
     - `TBL`: Bold, Left, `FirstIndent=Ribbon`, `KeepWithNext=1`. `FIG`: Bold, Centered, `FirstIndent=0`, `KeepWithNext=0`.
     - `SRC`: Bold prefix + Italic body, Left, `LeftIndent=0`, `FirstIndent=0`, `KeepWithNext=0`.
     - `BODY`: Justified, `FirstIndent=Ribbon`, `KeepWithNext=0` (Font/Size/Spacing from Ribbon; existing `SpaceBefore`/`SpaceAfter` preserved unless overridden).
     - Ask if the user wants a university/journal house style via `@STYLE` (e.g., colored headings, `upper=0`, `leftIndent=0`, `spaceAfter=6`) or explicit `@TAG` overrides.

3. **Step 3 — Generate `AF-PATCH v2` Block & Recommended Execution Order:**
   - Output ONLY inside a **single fenced `text` code block** using the exact flat syntax below (never JSON):

```text
=== AF-PATCH v2 ===
@STYLE
L1 | color=#1E6B3A | upper=0 | size=14
L2 | leftIndent=0
TBL | align=center | italic=1
BODY | spaceAfter=6
@TAG
TYPE: L1
CTX: DANH SÁCH THÀNH VIÊN

TYPE: TBL
CTX: Biểu số 1 - Thống kê doanh thu

TYPE: SKIP
CTX: 1.1. Câu trích dẫn mẫu không phải tiêu đề
@STRUCT
CTX: 1.1 Cơ sở lý thuyết
FIND: 1.1
NEW: 1.1.

CTX: Table A.1 Revenue
FIND: Table A.1
NEW: Table A.1:
@REFS
ref_smith18 | Smith, J. (2018)
ref_jones19a | Jones, M. (2019a)
ref_jones19b | Jones, M. (2019b)
ref_smith | Smith, J. (2020)
ref_nguyen | Nguyễn, V. A. (2021)
@FIX
CTX: hệ thống IT và VIT trong
FIND: IT
NEW: CNTT
@CITE
REF: ref_smith
CTX: trong giai đoạn 2020–2021 (Smith, 2020) và
FIND: Smith, 2020

REF: ref_nguyen
CTX: bài viết (Nguyễn 2021).
FIND: Nguyễn 2021
FIX: Nguyễn, 2021

REF: ref_smith18
CTX: (Smith, 2018; Jones, 2019a, b)
FIND: Smith, 2018

REF: ref_jones19a
CTX: (Smith, 2018; Jones, 2019a, b)
FIND: Jones, 2019a

REF: ref_jones19b
CTX: (Smith, 2018; Jones, 2019a, b)
FIND: b
@ACRONYMS
CNTT | Công nghệ thông tin |
LLM | Large Language Model | Mô hình ngôn ngữ lớn\nNguồn: Vaswani et al. (2017)
=== END ===
```
   - **Recommended Execution Order after Copying the Patch:** `Ctrl + S` -> `AI Patch` (1-Click Clipboard) -> **`Ctrl + S`** -> `Format All` (or `Update Lists` if only headings/captions/TOC changed) -> `Finalize Math` / `Clean Markers` -> **`Export PDF (Xuất PDF)`**. *(Remember to press `Ctrl + S` between `AI Patch` and `Format All` to pass Pre-Flight Gate 5.)*

#### B. Deterministic Grammar & Execution Guarantees (`AF-PATCH v2`)
- **`@STYLE` lines (`SCOPE | prop=val | ...` or `RESET | ALL` / `RESET | SCOPE`):**
  - **8 Valid Scopes:** `L1`, `L2`, `L3`, `L4`, `TBL`, `FIG`, `SRC`, `BODY`.
  - **13 Whitelisted Properties:**
    - `font` *(1–40 safe chars, e.g., `Arial`, `Times New Roman`)*
    - `size` *(`8` – `36` pt)*
    - `color` *(`#RRGGBB` hex e.g. `#1E6B3A`, or `auto`)*
    - `bold`, `italic`, `underline`, `upper`, `keepWithNext` *(`0` or `1`; `upper=0` on `L1` disables automatic uppercase while keeping `OMath` `<m:nor/>` safe)*
    - `align` *(`left`, `center`, `right`, `justify`)*
    - `firstIndent`, `leftIndent` *(`0` – `5` cm)*
    - `spaceBefore`, `spaceAfter` *(`0` – `72` pt)*
  - **Persistence & Reset:** Stored per-document in `doc.Variables("af_style_<SCOPE>")` (subsequent `@STYLE` directives on the same scope **merge** keys rather than wiping unspecified properties). Subsequent `AF-PATCH` runs that omit `@STYLE` **preserve** your active profile. Emit `RESET | ALL` (or `RESET | L1`, etc.) inside `@STYLE` to disable overrides for future runs *(note: `RESET` only clears the saved override profile; to reset already-colored headings back to automatic black, emit `L1 | color=auto` before or instead of `RESET`)*. Active custom scopes are summarized in the `Format All` completion report.
- **`@TAG` blocks (`TYPE:` + `CTX:`):**
  - **9 Valid `TYPE:` Values:** `L1`, `L2`, `L3`, `L4`, `TBL`, `FIG`, `SRC`, `SKIP` *(forces paragraph to be treated as normal body text and excluded from TOC/LOT/LOF even if it looks like a heading/caption)*, and `CLEAR` *(removes any prior `@TAG` bookmark on that paragraph)*.
  - **`CTX:`** must match **1 unique paragraph verbatim** (fail-closed if `0` or `2+` matches; skips `References` and generated list zones). Stored as hidden `bm_af_<type>_<nnnn>` bookmarks that **survive `Clean Markers`** and take priority over Regex classification in both `Format All` and `Update Lists` (setting proper `OutlineLevel` and applying any `@STYLE` overrides).
- **`@REFS` lines (`key_id | prefix`):** Provide an exact starting prefix (`>= 12` chars recommended) of each target paragraph inside `References` / `Tài liệu tham khảo` after `=contentEnd()`. The engine requires **1 exact unique match** (if duplicate `References` lines exist, `@REFS` fails closed as `Ambiguous ref` -> clean & deduplicate `References` first in Step 2). The VBA engine reuses existing `bm_ref_*` bookmarks or allocates `bm_ref_0001..` cleanly.
- **`@STRUCT` blocks (`CTX:` + `FIND:` + `NEW:`):** Structural grammar fixer (`1.1 Title` -> `1.1. Title`, `Chapter 1 Intro` -> `Chapter 1: Intro`, `Table A.1 Revenue` -> `Table A.1: Revenue`, `Nguồn (2025)` -> `Nguồn: (2025)`, `Lời cảm ơn:` -> `Lời cảm ơn`). **`CTX:` should be the entire heading/caption/source line verbatim** and `FIND:`/`NEW:` only change the prefix/separator. Unlike `@FIX`, `@STRUCT` scans **Pre-Content, Main Body (including existing `Heading` styles), and Post-Content / Appendices**, while strictly skipping `References`, generated list zones (`bm_zone_toc`/`lot`/`lof`/`loa`), `OMath`, `Fields`, and `Hyperlinks` (enforcing word boundaries and fail-closed `1`-hit uniqueness). **Workflow Tip:** Run **`AI Patch`** with `@STRUCT` / `@TAG` / `@STYLE` **before `Format All`** so headings and captions are fixed cleanly on the original draft while preserving embedded figures and tables.
- **`@FIX` blocks (`CTX:` + `FIND:` + `NEW:`):** Separated by blank lines. **`CTX:` is mandatory** and must contain `FIND:` verbatim so blind global replacements never occur. Scoped strictly to Main Body prose & body table cells (never touches Headings or Pre-Content). Word-boundary checks prevent corrupting substrings (`IT` will NOT match `VIT` or `ITV`).
- **`@CITE` blocks (`REF:` + `CTX:` + `FIND:` + optional `FIX:`):** **`CTX:` is mandatory.**
  - *Outer Punctuation Protection:* If `FIND:`/`FIX:` includes outer parentheses `(Smith, 2020)` or trailing `;`, the engine automatically trims the outer `(...)` / `;` so only `Smith, 2020` becomes the clickable hyperlink (ScreenTip `Ref -> <Reference Preview>`, distinct from `Ctrl+Click ->` used by Cross-Refs).
  - *Fail-Closed Ambiguity & DOM Guards:* `0` matches or `> 1` matches in the body cause that specific block to be **safely skipped** and reported in the Dry-Run summary (`never guessed`). Spans overlapping `OMath` equations, Word `Fields`, or `CrossRef` links (`bm_toc_`, `bm_tbl_`, `bm_fig_`) are never overwritten. Citations inside body data table cells (`=contentStart()..=contentEnd()`) ARE supported (`Chr(7)` protected).
- **`@ACRONYMS` lines (`ABBR | Full Form | Optional Column 3`):**
  - Deduplicates case-insensitively and sorts strictly **A–Z** (`2` columns if Column 3 is empty on all rows; `3` columns otherwise).
  - Uses literal `\n` inside Column 3 for line breaks; each sub-line is trimmed (`Trim$`), and leading labels `Ý nghĩa:`, `Nguồn:`, `Definition:`, `Meaning:`, `Source:` are automatically formatted **bold**.
  - Formatted using current Ribbon `Font (Phông)` & `Size (Cỡ)` (`13 pt` default), white header background (`HeadingFormat = -1`), vertical center alignment across all cells, Row 1 & Col 1 centered, Col 2 & Col 3 left-aligned.
  - **Incremental Patch Safety:** If a subsequent patch omits `@ACRONYMS` (`0` items), the existing LOA table is **100% preserved untouched**. If `>= 1` items are present, `bm_zone_loa` is cleanly rebuilt.
- **Pre-Audit Scan (`Đếm trích dẫn`):** Read-only scan over body paragraphs + body table cells supporting English (`n.d.`, `in press`, `forthcoming`, `et al.`, `&`, `and`) and Vietnamese APA forms (`k.n.`, `k.n`, `kn`, `k.n.x.b.`, `k.nxb`, `không rõ năm`, `chưa rõ năm`, `đang in`, `đang xuất bản`, `sắp xuất bản`, `và cs.`, `và cộng sự`, `& cs.`, `và nnk.`, `và những người khác`, `cùng cộng sự`), reporting Parenthetical (clusters & split items), Narrative, References entries, and Chapter breakdown.

#### C. `@STYLE` & `@TAG` Patching Principles & Guardrails (AI Reference Rules)

When advising users or generating `@STYLE` / `@TAG` directives, AI Assistants MUST follow these 5 deterministic operating principles:

1. **3-Tier Precedence Hierarchy (Ribbon vs. `@STYLE` vs. `@TAG`):**
   - **Tier 1 — Ribbon (`g_Config`):** Document-wide baseline (`Font`, `Size`, `Spacing`, `Indent`, `Lang`, `Depth`, `Ref Link`).
   - **Tier 2 — `@STYLE` (`doc.Variables("af_style_<SCOPE>")`):** Sparse per-scope overrides (`L1..L4`, `TBL`, `FIG`, `SRC`, `BODY`); any omitted property inherits the Ribbon/engine default automatically. Stored strictly inside the current `.docx` file (never affects other Word documents).
   - **Tier 3 — `@TAG` (`bm_af_*` bookmarks):** Explicit structural classification that takes priority **before** all built-in regexes (even on paragraphs with Word Auto-Numbering `ListType <> 0`) and stays synchronized across `Format All`, `Update Lists`, `Pre-Audit Scan`, `@CITE`, `@FIX`, Word Navigation Pane, and PDF Bookmarks.
2. **Strict Scope Boundaries (What `@STYLE` Touches vs. Strictly Protects):**
   - **`L1`** applies uniformly to Chapters, Front/Back matter headings, `@TAG L1` lines, **and all 4 auto-generated list headers** (`MỤC LỤC`, `DANH MỤC BẢNG BIỂU`, `DANH MỤC HÌNH ẢNH`, `DANH MỤC TỪ VIẾT TẮT`).
   - **`BODY`** applies ONLY to body prose paragraphs (`ElementTypeNormalParagraph`) between `=contentStart()..=contentEnd()` and **only runs during `Format All`** (`Update Lists` leaves body paragraphs 100% untouched). `BODY` **NEVER** touches the Cover Page / Front Matter, data table cells (`ElementTypeTableCell`), `References` entries after `=contentEnd()`, or TOC/LOT/LOF entry lines (which always inherit Ribbon Font/Size and black text).
   - **`SRC` Override:** By default, `Source:`/`Nguồn:` splits **Bold** prefix + *Italic* content; specifying `bold=...` or `italic=...` in `@STYLE SRC` applies that style uniformly to the entire `Source:`/`Nguồn:` line.
   - **`L1 | upper=0` Behavior:** Only disables the engine's automatic uppercase conversion; it does **NOT** lowercase headings that the author originally typed in ALL CAPS (e.g., `CHƯƠNG I: ...` stays uppercase unless the author edits the raw text).
3. **Incremental Merge, `Ctrl + Z` Caveat & Proper Reset:**
   - Sending additional `@STYLE` directives for the same scope **merges** new keys into the existing profile without wiping previously set properties. Subsequent `AI Patch` runs without `@STYLE` preserve the active profile untouched.
   - Because `Document.Variables` live outside Word's `Undo` stack, pressing `Ctrl + Z` after experimenting with `@STYLE` will NOT remove saved `af_style_*` variables.
   - **Standard Reset Path:** Recommend clicking **`AI Patch -> Reset Custom Style (Về mặc định)`** on the Ribbon — it lets the user choose **`[Yes]`** (clear `@STYLE` only, keep `@TAG`) or **`[No]`** (clear both `@STYLE` and `@TAG`), and offers a clean **`Format All`** pass that strips leftover custom colors, underlines, `AllCaps`, italics, and paragraph spacing (`SpaceBefore`/`SpaceAfter`) back to default academic black. Distinguish this from `RESET | ALL` inside a patch block (which only clears saved variables for future runs without recoloring already-colored headings unless `color=auto` or the Ribbon reset button is used).
4. **Vietnamese Unicode Font Safety Rule (Mandatory for `@STYLE font=...`):**
   - When suggesting `font=...` for Vietnamese manuscripts, AI MUST prioritize full-coverage Unicode fonts supporting **Latin Extended Additional (`U+1EA0–U+1EF9`, `ơ`, `ư`)** across Regular/Bold/Italic faces: `Times New Roman`, `Arial`, `Calibri`, `Cambria`, `Verdana`, `Tahoma`, `Segoe UI`, `Consolas`, `Trebuchet MS`.
   - Warn users to avoid decorative fonts or italic faces missing extended Vietnamese glyphs (such as `Georgia Italic` or non-Unicode fonts), which cause Windows/Word to render hollow boxes (`☐`) or uneven fallback glyphs.
5. **`@TAG` Writing Rules, Synergy & Cross-Ref Boundaries:**
   - **Exact-Line Priority:** The `@TAG` planner prioritizes exact whole-paragraph matches (`StrComp`) before falling back to substring matches (`InStr`) -> AI should **always copy the entire heading/caption line verbatim** into `CTX:` to avoid `Ambiguous (2+ hits)` when a heading phrase is also mentioned in prose.
   - **`@TAG` + `@STRUCT` Synergy:** Both `@STRUCT` and `@TAG` can target the same paragraph in a single `AF-PATCH v2` block (the engine attaches `bm_af_*` *after* `@STRUCT` finishes editing the text, so the tag is never lost).
   - **`@TAG SKIP`:** Removes false-positive headings cleanly from TOC, Word Navigation Pane, and PDF Bookmarks (even if the paragraph previously carried a built-in `Heading 1..9` style), preserves table cells as `TableCell`, and unlocks `Pre-Audit Scan`, `@CITE`, and `@FIX` to scan citations inside that paragraph.
   - **Persistence & Cross-Ref Limit:** `bm_af_*` bookmarks survive `Clean Markers` and `Update Lists` (editing text inside a tagged line keeps the tag, whereas newly typed non-standard headings later need a new `@TAG` patch). Custom unnumbered `@TAG` headings appear in TOC/LOT/LOF, Navigation Pane, and PDF Bookmarks, while prose auto-cross-referencing (`ModCrossRef`) still matches standard numbered tokens (`Bảng 1.1`, `Table 1.1`, `Hình 2.1`, `Đồ thị 1.1`, `Chương 1`, `Mục 1.1`, `Phụ lục A`...).

---

## Part 3: Document Zones & No-Fly Zones

### 3.1. Three Structural Document Zones (`Format All`)
The formatting engine divides your document into 3 zones relative to `=contentStart()` and `=contentEnd()`:

| Zone                                                | Location                                                     | What Is Formatted                                                                                                                                                                                                                                                                                                                                                                                                                          | What Is Strictly Untouched (Protected)                                                                                                                                                                                                                                                                                                                                                                                                             |
| :-------------------------------------------------- | :----------------------------------------------------------- | :----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | :------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| **1. Pre-Content Zone (Front Matter)**              | Paragraphs **before** `=contentStart()`                      | - **Level 1 Headings** (`Abstract`, `Acknowledgements`, `Lời cảm ơn`, `Danh mục từ viết tắt`, etc.)<br>- Generated **TOC / LOT / LOF** zones                                                                                                                                                                                                                                                                                               | - **Cover Page / Title Page** text, borders, and layout tables<br>- **Acronym / Abbreviation tables** and front-matter prose<br>- Line spacing and indentation of non-heading paragraphs                                                                                                                                                                                                                                                           |
| **2. Main Body Zone**                               | Paragraphs **between** `=contentStart()` and `=contentEnd()` | - **Headings Level 1–4** (`LeftIndent`: `0.5 cm` L2, `0.75 cm` L3, `1.0 cm` L4; `KeepWithNext = True`)<br>- **Normal Paragraphs** (Font, Size, Justified, First-Line Indent, `KeepWithNext = False`, optional Line Spacing; preserves existing `SpaceBefore`/`SpaceAfter`)<br>- **Table & Figure Captions** (`KeepWithNext = True` on Table Captions) + **Source/Note lines**<br>- **Inline Pictures** (Centered, `FirstLineIndent=0`, `KeepWithNext = True` with following caption)<br>- **Data Table Cells** (`Size = BodySize`, Centered horizontally & vertically; preserves existing `SpaceBefore`/`SpaceAfter`) | - **Native Word `OMath` equations** (keeps inner `<m:nor/>` `\text{...}` font intact while syncing `Font.Size`)<br>- **Raw LaTeX Display Blocks** (`$$...$$`, `\[...\]`)<br>- **Bulleted & Numbered Lists** (`ListFormat.ListType <> 0`)<br>- **Floating Shapes / TextBoxes / SmartArt / Charts**<br>- **Existing built-in Word TOC fields** (`doc.TablesOfContents`)<br>- **User Highlights** (Green, Turquoise, Pink, Gray, etc. — only system `wdYellow` and `wdRed` inside body text are cleared) |
| **3. Post-Content Zone (Back Matter & Appendices)** | Paragraphs **after** `=contentEnd()`                         | - **Level 1 Headings** (`References`, `Tài liệu tham khảo`, `Appendix A`, `Phụ lục 1`, etc.)<br>- **Appendix Table & Figure Captions** (`Table A.1: ...`, `Hình A.1: ...` — indexed into LOT/LOF!)<br>- **Appendix Source/Note lines**, **Appendix Table Cells**, and **Appendix Inline Pictures**                                                                                                                                         | - **Reference / Bibliography list entries** (hanging indents, APA/IEEE italics, and line spacing are 100% preserved)<br>- Normal prose paragraphs in Appendices (line spacing not overwritten)                                                                                                                                                                                                                                                     |

### 3.2. Math Engine No-Fly Zones (`Build Math`)
When you click **`Build Math (Dựng Toán)`**, the LaTeX scanner automatically skips:
1. **Generated List Zones:** `bm_zone_toc` (TOC), `bm_zone_lot` (LOT), `bm_zone_lof` (LOF), and `bm_zone_loa` (List of Acronyms).
2. **References / Bibliography Section:** As soon as the scanner encounters a standalone heading matching `References`, `Bibliography`, `Tài liệu tham khảo`, or `Danh mục tài liệu tham khảo` (`<= 160` chars, without `(`), **math scanning pauses** so `$` signs or URLs/DOIs in bibliography citations are never touched.
3. **Automatic Resume at Appendices:** If an `Appendix` / `Appendices` / `Phụ lục` heading (`<= 160` chars, without `(`) appears after `References`, **math scanning automatically resumes** so equations in your Appendices (`$$\sigma^2 = ...$$`) are built normally!

---

## Part 4: Ribbon Controls & Standard 5-Step Workflow

### 4.1. Map of the `Formatter Macro` Ribbon Tab
| Group                            | Control Label                                                                                                                         | Options / Action                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                      |
| :------------------------------- | :------------------------------------------------------------------------------------------------------------------------------------ | :---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| **1. Markers (Thẻ Đánh Dấu)**    | `Insert Marker (Chèn Thẻ)`                                                                                                            | Inserts standalone markers at cursor (blocks duplicates and table-cell insertion):<br>• `TOC =listContent() (Mục lục)`<br>• `LOT =listTableHere() (DM Bảng)`<br>• `LOF =listFigureHere() (DM Hình)`<br>• `LOA =listAcronyms() (DM Viết tắt)`<br>• `Start =contentStart() (Bắt đầu)`<br>• `End =contentEnd() (Kết thúc)`                                                                                                                                                                                                                                                                                                                                                                                                                                                                               |
| **2. Typography (Kiểu chữ)**     | `Font (Phông):`<br>`Size (Cỡ):`<br>`Spacing (Giãn):`<br>`Indent (Thụt):`                                                              | • `Times New Roman` *(default)*, `Arial`, `Calibri`<br>• `12 pt`, `13 pt (Standard / Chuẩn)` *(default)*, `14 pt`<br>• `Keep (Giữ nguyên)` *(default)*, `1.15 lines`, `1.3 lines`, `1.5 lines`<br>• `1.27 cm (0.5 in)` *(default)*, `1.00 cm`, `0.00 cm`                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                              |
| **3. List Config (Cấu hình DM)** | `Lang (Ngôn ngữ):`<br>`Bold All TOC (In đậm TOC)`<br>`Footer Page # (Số trang Footer)`<br>`Depth (Độ sâu):`<br>`Ref Link (Màu link):` | • `Vietnamese (Tiếng Việt)` -> `MỤC LỤC`, `DANH MỤC BẢNG BIỂU`, `DANH MỤC HÌNH ẢNH`, `DANH MỤC TỪ VIẾT TẮT`<br>• `English (Tiếng Anh)` -> `TABLE OF CONTENTS`, `LIST OF TABLES`, `LIST OF FIGURES`, `LIST OF ACRONYMS`<br>• `Bold All TOC` *(Checked = bold all TOC entries; Unchecked = only Level 1 bold in TOC)*<br>• `Footer Page #` *(Checked = reads formatted footer page numbers like `i, ii, 1, 2`; Unchecked = physical PDF page index)*<br>• `Level 1-4 (Cấp 1-4)` *(default)* or `Level 1-3 (Cấp 1-3)`<br>• `Ref Link`: `Blue` *(default)*, `Orange`, `Red`, `Black`, `Off` *(auto-links & syncs Cross-Refs and `bm_ref_*` citations; `Off` clears Cross-Refs only and preserves existing `bm_ref_*` citations; applying a new AI Patch while `Off` defaults new citation links to Blue)* |
| **4. Math (Toán LaTeX)**         | `Build Math (Dựng Toán)`<br>`Toggle / Edit (Đổi / Sửa)`<br>`Finalize Math (Chốt Toán)`                                                | • **Build Math:** Batch-converts all LaTeX in the document (and Appendix after `=contentEnd()`) into Word `OMath` with Pastel review shading, wrapped in Pre-flight & Crash Backup protection.<br>• **Toggle / Edit:** Context-sensitive 4-state switch at the current cursor/selection.<br>• **Finalize Math:** Strips all Pastel shading, temporary `doc.Variables`, and `bm_math_*` bookmarks (warns if unbuilt Pink/Amber items remain).                                                                                                                                                                                                                                                                                                                                                          |
| **5. Actions (Thực thi)**        | `Format All (Định dạng)`<br>`Update Lists (Cập nhật DM)`<br>`Clean Markers (Xóa Thẻ)`<br>`AI Patch` *(SplitButton)*<br>`Export PDF (Xuất PDF)` | • **Format All:** Full 5-phase formatting of body text, headings, captions, tables, figures (applying active `@STYLE` & `@TAG`) + rebuilds TOC/LOT/LOF & syncs link colors.<br>• **Update Lists:** High-speed refresh of headings/captions (including `@STYLE`/`@TAG`), link colors, and TOC/LOT/LOF page numbers **without touching body paragraphs**.<br>• **Clean Markers:** Removes all `=...()` paragraphs (keeping `bm_loa_header` and `bm_af_*` `@TAG` bookmarks) and finalizes math shading before PDF export.<br>• **AI Patch:** 1-Click applies `AF-PATCH v1/v2` (`@STYLE`, `@TAG`, `@STRUCT`, `@REFS`, `@FIX`, `@CITE`, `@ACRONYMS`) from **Clipboard**, or dropdown **`Apply from File...`**, **`Pre-Audit Scan (Đếm trích dẫn)`**, or **`Reset Custom Style (Về mặc định)`** *(clears `@STYLE` and optionally `@TAG`, with optional clean `Format All` pass to restore default black typography)*.<br>• **Export PDF (Xuất PDF):** 1-Click PDF export with clean **Heading Bookmarks** (`CreateBookmarks:=1` — never leaks internal `bm_*` bookmarks), warns if active markers or unfinalized math remain, safely resolves local save folder (fallback to `Documents` for unsaved/OneDrive docs), and detects locked PDF files cleanly. |

### 4.2. The 4-State `Toggle / Edit (Đổi / Sửa)` Cycle
Place your cursor inside an equation (or select raw text) and click **`Toggle / Edit (Đổi / Sửa)`**:
1. **Built Equation (Pastel Cyan / Mint `OMath`) -> Raw LaTeX (Pastel Amber `#FFF2CC`):** Restores the exact original LaTeX source string so you can edit a typo.
2. **Edited Raw LaTeX (Pastel Amber) -> Rebuilt Equation (Pastel Cyan / Mint `OMath`):** Re-parses your edited LaTeX and builds a fresh Word `OMath`.
3. **Ambiguous / Flagged Span (Pastel Pink `#FFE6E6`, e.g., `$500$`) -> Forced `OMath` (Pastel Cyan / Mint):** Forces conversion of an ambiguous span into a Word `OMath` (if syntax is valid).
4. **Unmarked Raw Selection (User highlights plain text `x^2 + 1` or `$x^2 + 1$`) -> Built `OMath`:** Converts the selected text on the spot and registers a `bm_math_omath_` bookmark so it can be toggled later.

### 4.3. How `Update Lists (Cập nhật DM)` Works — Including After Markers Are Deleted!
`Update Lists` is engineered for late-stage editing when you have tweaked custom spacing/styling in your body paragraphs and only want to refresh headings, captions, and TOC/LOT/LOF page numbers (with **zero backup file overhead**):
- **3-Tier List Zone Discovery (`DiscoverAndPrepareQuickZones`):** Even if you already ran `Clean Markers` (deleting `=listContent()`, `=listTableHere()`, `=listFigureHere()`), `Update Lists` still finds and refreshes your lists using a 3-tier fallback:
  - **Tier 1 (System Bookmarks):** Finds existing `bm_zone_toc`, `bm_zone_lot`, `bm_zone_lof`.
  - **Tier 2 (Active Markers):** Finds `=listContent()`, `=listTableHere()`, `=listFigureHere()`.
  - **Tier 3 (Markerless Header Recognition):** Finds standalone header paragraphs (`MỤC LỤC` / `TABLE OF CONTENTS`, `DANH MỤC BẢNG BIỂU` / `LIST OF TABLES`, `DANH MỤC HÌNH ẢNH` / `LIST OF FIGURES`) located before `=contentStart()`.
- *Note:* `=contentStart()` and `=contentEnd()` must still be present for `Update Lists` to know the body boundaries. Always run `Clean Markers (Xóa Thẻ)` as your **very last step** before exporting to PDF.

### 4.4. Recommended End-to-End Workflow
1. **Step 1 — Insert Markers & Configure Ribbon:** Place `=listContent()`, `=listTableHere()`, `=listFigureHere()`, `=listAcronyms()` in front matter, `=contentStart()` before Chapter 1, and `=contentEnd()` after your last body paragraph. Choose your Font, Size, Line Spacing, Ref Link Color, and List Language (`Vietnamese` or `English`) on the **`Formatter Macro`** tab.
2. **Step 2 — Build & Review Math (`Build Math`):** Click **`Build Math (Dựng Toán)`**. Inspect any **Pastel Pink** equations; click inside them and use **`Toggle / Edit (Đổi / Sửa)`** to fix typos.
3. **Step 3 — Pre-Audit Scan & Apply AI Patch (`AI Patch` — Recommended Before `Format All`):** Click **`AI Patch -> Pre-Audit Scan`** to count EN/VI APA citations and References entries (copy summary to Clipboard). Collaborate with your AI using the flexible 3-Step workflow (`INSTRUCTION.md` Part 2.5), copy the `=== AF-PATCH v2 ===` block, and click **`AI Patch`** (1-Click Clipboard) to set custom `@STYLE` profiles, classify non-standard lines (`@TAG`), fix heading/caption prefixes (`@STRUCT`), link citations (`@CITE`), normalize acronyms (`@FIX`), and build the List of Acronyms (`@ACRONYMS`) directly on your draft without losing images or tables.
4. **Step 4 — Run Full Formatting (`Format All` / `Update Lists`):** Save your file (`Ctrl + S`), then click **`Format All (Định dạng)`** (assigns `OutlineLevel` for Word Navigation Pane & PDF Bookmarks, applies `@STYLE` and `@TAG`, formats lists/styles, and syncs `bm_ref_*` colors with `Ref Link`). Later edits to headings/captions/colors can use **`Update Lists (Cập nhật DM)`**.
5. **Step 5 — Finalize for Submission & 1-Click PDF Export (`Finalize Math` + `Clean Markers` + `Export PDF`):** Click **`Finalize Math (Chốt Toán)`** and **`Clean Markers (Xóa Thẻ)`** to strip Pastel shading and `=...()` markers (`bm_loa_header`, `bm_af_*` `@TAG` bookmarks, and citation links stay intact). Then click **`Export PDF (Xuất PDF)`** on the Ribbon to export a submission-ready PDF with clean **Heading Bookmarks** (`CreateBookmarks:=1`) in 1 click *(no manual `File -> Save As -> Options -> Create bookmarks using: Headings` dialog needed!)*.

---

## Part 5: Highlight/Shading Color Decoder & Troubleshooting Matrix

### 5.1. Complete Color Decoder (6 Diagnostic Colors)
The engine uses **Word Highlight (`Range.HighlightColorIndex`)** for document structure markers and **Paragraph/Character Background Shading (`Range.Shading.BackgroundPatternColor`)** for the Math Engine:

| Visual Color                | Technical Constant / RGB                       | Where It Appears                                                                                                    | Exact Meaning & Required Action                                                                                                                                                                                                                                        |
| :-------------------------- | :--------------------------------------------- | :------------------------------------------------------------------------------------------------------------------ | :--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| **Bright Yellow Highlight** | `wdYellow` (`HighlightColorIndex = 7`)         | On active `=contentStart()`, `=contentEnd()`, `=list...()` marker paragraphs                                        | **Active System Marker:** Shows active `=...()` markers in the document (automatically removed when you click `Clean Markers`).                                                                                                                                        |
| **Bright Red Highlight**    | `wdRed` (`HighlightColorIndex = 6`)            | On a duplicate manual list heading (`MỤC LỤC` when `=listContent()` is also present)                                | **Skipped Duplicate Manual Header:** Delete the redundant manual heading if `=list...()` already generates it *(unclosed `$$`/`\[` blocks are safely skipped within their 40-paragraph window without corrupting body text, and counted in the summary dialog)*. |
| **Pastel Cyan Shading**     | RGB `(220, 235, 245)` / `#DCEBF5` (`16119260`) | On inline `OMath` equations (`$...$`, `\(...\)`)                                                                    | **Success (Inline Math):** Inline LaTeX was cleanly transpiled to Word `OMath`. Cleared automatically when you click `Finalize Math` or `Clean Markers`.                                                                                                               |
| **Pastel Mint Shading**     | RGB `(220, 245, 225)` / `#DCF5E1` (`14808540`) | On display `OMath` equations (`$$...$$`, `\[...\]`, `\begin{...}`)                                                  | **Success (Display Block Math):** Display LaTeX was cleanly transpiled to centered Word `OMath`. Cleared automatically by `Finalize Math` or `Clean Markers`.                                                                                                          |
| **Pastel Pink Shading**     | RGB `(255, 230, 230)` / `#FFE6E6` (`15132415`) | On raw `$...$` or `$$...$$` text                                                                                    | **Ambiguous or Unsupported LaTeX (Fail-Closed):** Either an ambiguous number/unit (`$500$`, `$20 MT$`), unbalanced braces, bordered `array` (`{\|c\|c\|}` / `\hline`), mismatched `\begin`/`\end`, or unknown LaTeX command. Edit the LaTeX and click `Toggle / Edit`. |
| **Pastel Amber Shading**    | RGB `(255, 242, 204)` / `#FFF2CC` (`13497343`) | On raw LaTeX text after clicking `Toggle / Edit`                                                                    | **Edit Mode (Raw LaTeX Source):** You toggled a built `OMath` back to raw LaTeX to edit it. Click `Toggle / Edit` again when done to rebuild the `OMath`.                                                                                                              |

---

### 5.2. Master Troubleshooting Matrix (Symptom -> Root Cause -> Deterministic Fix)

| #      | Symptom Observed                                                                                                        | Exact Root Cause                                                                                                                                                                                         | Deterministic Fix                                                                                                                                                                                                                                                                                      |
| :----- | :---------------------------------------------------------------------------------------------------------------------- | :------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | :----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| **1**  | Inline formula `\$x^2 + y^2\$` is completely ignored by `Build Math`                                                    | Pasting from ChatGPT/Claude/Notion with normal `Ctrl + V` copied escaped `\$` or invisible HTML spans. Escaped `\$` is treated as a literal dollar character.                                            | Replace `\$` with `$` in Word (`Ctrl + H`: Find `\$`, Replace with `$`), then run `Build Math`. Next time, copy from a fenced `text` block and paste with **`Ctrl + Shift + V`**.                                                                                                                      |
| **2**  | Subheading `1.1 Research Scope` or `2.1.3 Methodology` is formatted as a normal body paragraph and missing from the TOC | Missing the **mandatory trailing dot** after the final digit (`^(\d+\.\d+\.)\s+(.+)$`).                                                                                                                  | Add a period after the number: `1.1. Research Scope` and `2.1.3. Methodology`, then click `Update Lists`.                                                                                                                                                                                              |
| **3**  | Table or Figure caption `Table 1 Revenue` is missing from the List of Tables / Figures                                  | Missing the **mandatory separator** (`:`, `. `, or ` - `) after the number token (`1`).                                                                                                                  | Add a colon or dot+space: `Table 1: Revenue` or `Table 1. Revenue`, then click `Update Lists`.                                                                                                                                                                                                         |
| **4**  | `Source` or `Nguồn` line under a table/figure is formatted as normal body paragraph                                     | Missing the **mandatory colon (`:`)** immediately after `Source`, `Note`, `Nguồn`, or `Ghi chú` (`Source:` should be Bold prefix + Italic text, `0 cm` indent).                                          | Change `Nguồn (2025)` to `Nguồn: (2025)` or `Source: Author's calculation`, then click `Update Lists`.                                                                                                                                                                                                 |
| **5**  | TOC / LOT / LOF header titles are generated in Vietnamese (`MỤC LỤC`) inside an English thesis                          | The Ribbon dropdown `Lang (Ngôn ngữ)` defaults to `Vietnamese (Tiếng Việt)`.                                                                                                                             | On the **`Formatter Macro`** tab, switch **`Lang (Ngôn ngữ)`** to **`English (Tiếng Anh)`** and click **`Update Lists (Cập nhật DM)`** (takes ~1 second).                                                                                                                                              |
| **6**  | `$500$` or `$20 MT$` is shaded **Pastel Pink** instead of being ignored or built                                        | Pure numbers (`$500$`) or number + unknown unit with space (`$20 MT$`) are flagged as ambiguous (`Class 2`) so the engine doesn't silently guess whether you meant $500 USD or the constant $500$.       | • If it is **currency/text**: write `$500 USD` or `500 USD`, or leave it—`Finalize Math` will remove the Pink shading and keep `$500$` as plain text.<br>• If it is a **math constant**: click inside `$500$` and click **`Toggle / Edit (Đổi / Sửa)`** to force-build it into an `OMath`.             |
| **7**  | A matrix or `array` block is shaded **Pastel Pink**                                                                     | Word OMML `<m:m>` does not support internal vertical/horizontal table borders (`\begin{array}{\|c\|c\|}` or `\hline`), or `\begin{pmatrix}...\end{matrix}` tags are mismatched.                          | Remove `\|` from `{cc}` and remove `\hline` (or use `pmatrix` / `bmatrix`), then click `Toggle / Edit`. For bordered data grids, use a real Word Table.                                                                                                                                                |
| **8**  | TOC page numbers show `1, 2, 3...` when the footer shows `i, ii, iii...` (or vice versa)                                | `Footer Page # (Số trang Footer)` checkbox on the Ribbon does not match your desired page numbering mode, or your document has no section break between front matter and Chapter 1.                      | Keep **`Footer Page #`** checked to read Word's formatted section page numbers (`i, ii, 1, 2`), or uncheck it to use physical PDF page numbers (`1..N`), then click `Update Lists`.                                                                                                                    |
| **9**  | Clicking `Finalize Math (Chốt Toán)` shows a warning: `Note: X Pink/Amber item(s) kept as raw text`                     | You had `X` equations still in **Pastel Pink** (unbuilt/ambiguous) or **Pastel Amber** (toggled to raw edit mode) when finalizing.                                                                       | If unintentional, press **`Ctrl + Z`** once to undo `Finalize Math`, review the Pink/Amber items, convert them with `Toggle / Edit`, and click `Finalize Math` again.                                                                                                                                  |
| **10** | Clicking `Format All`, `Update Lists`, or `Apply AI Patch` shows `Pre-Flight Failed: Unsaved changes detected`          | The document has unsaved edits in memory (`doc.Saved = False`).                                                                                                                                          | Press **`Ctrl + S`** to save the document, then click the button again.                                                                                                                                                                                                                                |
| **11** | AI Patch Dry-Run skips an entry with `Ambiguous (2+ hits)` or `Not found in body`                                       | Either `CTX:` was too short and matched 2+ places in the body, or `CTX:`/`FIND:` did not match the exact paragraph text.                                                                                 | Give that `@CITE` or `@FIX` block a longer unique `CTX:` snippet copied verbatim from Word (or split multi-citations onto separate `@CITE` blocks sharing one `CTX:`).                                                                                                                                 |
| **12** | Clicking `AI Patch` shows `Clipboard does not contain Unicode text` or `Invalid AF-PATCH`                               | Clipboard is empty or missing the `=== AF-PATCH v1 ===` / `=== AF-PATCH v2 ===` and `=== END ===` lines.                                                                                                 | Copy the fenced `text` patch block from your AI chat cleanly and click `AI Patch`, or save it as UTF-8 `.txt` and use `AI Patch -> Apply from File...`.                                                                                                                                                |
| **13** | AI Patch Dry-Run warns `không tìm thấy =listAcronyms()` and skips `@ACRONYMS`                                           | Neither `=listAcronyms()`, an existing `bm_loa_header` bookmark, nor a manual LOA heading exists before `=contentStart()`.                                                                               | Insert `=listAcronyms()` on a standalone line in front matter (before `=contentStart()`), then re-run `AI Patch`.                                                                                                                                                                                      |
| **14** | Exported PDF bookmark sidebar shows `bm_toc_0001` / `bm_ref_0001` or has no heading bookmarks                           | Manual Word PDF export **Options** used `"Word bookmarks"` instead of `"Headings"` (which reads `ParagraphFormat.OutlineLevel`).                                                                         | Click **`Export PDF (Xuất PDF)`** on the `Formatter Macro` tab (exports with `CreateBookmarks:=1` Headings automatically). If exporting manually via `File -> Save As -> PDF -> Options...`, tick **`Create bookmarks using: Headings`** (never *"Word bookmarks"*).                                  |
| **15** | `install-or-update.bat` exits with code `4`, `6`, `7`, or `[BLOCKED]`                                                   | Code `4`: Word (`WINWORD.EXE`) was still running.<br>Code `6` / `[BLOCKED]`: Windows Defender/antivirus quarantined or blocked download/copy.<br>Code `7`: Local `.dotm` failed ZIP/XML integrity check. | • Close ALL Word windows (check Task Manager for stray `WINWORD.EXE`) and re-run.<br>• Re-download clean release zip from GitHub Releases if `Code 7`, or exclude/allow the extracted `.dotm` in Windows Security (`Protection History`) and re-run.                                                   |
| **16** | `Formatter Macro` tab is missing in Word even though `AcademicFormatter.dotm` is in `STARTUP`                           | Either Word quarantined the template in `Resiliency\DisabledItems` after a prior crash, or Trust Center **Macro Settings** has `VBAWarnings = 4` (*Disable all macros without notification*).                                                  | Close ALL Word windows and run **`install-or-update.bat`** to auto-clear `DisabledItems` + strip `Zone.Identifier`. Also verify **`File -> Options -> Trust Center -> Trust Center Settings -> Macro Settings`** allows VBA macros (and check `File -> Options -> Add-ins -> Manage: Disabled Items`). |
| **17** | AI Patch Dry-Run skips `@REFS` with `Ambiguous ref (2+)` or skips entire chapters after merging multiple authors' drafts | The document contains duplicate entries in `References`, or multiple per-chapter `References` headings left inside `=contentStart()..=contentEnd()` (which triggers the `References` No-Fly Zone early and blocks subsequent chapters). | Ask the AI to run **Step 2** (`INSTRUCTION.md` Part 2.5.A) to merge, deduplicate, and A–Z sort all references into a single list placed **after `=contentEnd()`**, delete all per-chapter `References` blocks in Word, press **`Ctrl + S`**, then proceed to Step 3 and run **`AI Patch`**.               |
| **18** | Pressed `Ctrl + Z` after testing `@STYLE`, but the next `Format All` / `Update Lists` still reapplies custom colors/fonts | Word `Document.Variables` (`af_style_*`) live outside Word's `UndoRecord` stack, so `Ctrl + Z` reverts document text without clearing the saved `@STYLE` profile. | Click **`AI Patch -> Reset Custom Style (Về mặc định)`** on the Ribbon (choose `[Yes]` to keep `@TAG` or `[No]` to clear both `@STYLE` & `@TAG`) and confirm the clean `Format All` pass to restore default black typography. |
| **19** | Vietnamese accented letters (`ơ, ư, ế, ệ, ộ, ả...`) render as hollow boxes (`☐`) or mismatched glyphs after setting `@STYLE font=...` | The chosen font or its Italic face lacks full **Latin Extended Additional (`U+1EA0–U+1EF9`)** Vietnamese glyphs (e.g., `Georgia Italic` or legacy non-Unicode fonts). | Switch `@STYLE font=...` to a full Vietnamese Unicode font (`Times New Roman`, `Arial`, `Calibri`, `Cambria`, `Verdana`, `Tahoma`, `Segoe UI`, `Consolas`, `Trebuchet MS`), then run **`AI Patch` -> `Ctrl + S` -> `Format All`**. |
