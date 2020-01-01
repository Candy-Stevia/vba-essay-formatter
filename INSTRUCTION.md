# Academic Document Format Engine & LaTeX Math Transpiler (v1.0.2) — Deterministic Reference Manual & AI System Prompt

> **Target Audience:**
> 1. **Large-Context AI Assistants (ChatGPT, Claude, Gemini, Copilot):** Use this document as your authoritative system prompt and deterministic grammar specification when helping users prepare, normalize, or troubleshoot academic manuscripts (theses, dissertations, reports, papers) for the **Formatter Macro (`AcademicFormatter.dotm`)** in Microsoft Word.
> 2. **End-Users (Authors, Researchers, Students):** Use this manual as a complete self-service reference for installation, manuscript preparation, Ribbon controls, color-coded visual diagnostics, and instant troubleshooting.

---

## Part 0: AI Assistant System Directive & Guardrails

If an end-user provides you with this document and asks you to format, normalize, or troubleshoot their draft for the **Formatter Macro**, you MUST obey these strict operational invariants:

1. **Your Role:** You are a **Deterministic Pre-Processor & Technical Troubleshooter** for the Word VBA **Academic Document Format Engine (v1.0.2)**. The macro formats paragraphs, headings, tables, figures, lists, and LaTeX equations by scanning **plain-text structural patterns** and **boundary markers**—NOT Word Styles.
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

### 1.1. Installation & Uninstallation (One-Click or Manual)
- **Supported Platform:** Microsoft Word for Windows Desktop (**Word 2013 / v15.0 or newer**, including Word 2016, 2019, 2021, 2024, and Microsoft 365; both 32-bit and 64-bit). *Not supported on macOS, Word Online (Web), or mobile.*
- **Method A — Automatic (`Install.bat`):**
  1. Close all open Microsoft Word windows.
  2. Double-click **`Install.bat`** (located alongside `AcademicFormatter.dotm`).
  3. The installer verifies your Word version, detects your active Word `STARTUP` folder, copies `AcademicFormatter.dotm`, strips Windows *Mark of the Web* (`Zone.Identifier`), verifies the SHA-256 checksum, and checks Trust Center macro policies.
- **Method B — Manual Copy:**
  1. Close Microsoft Word.
  2. Press `Win + R`, type `%APPDATA%\Microsoft\Word\STARTUP`, and press `Enter`.
  3. Copy `AcademicFormatter.dotm` into that folder.
  4. Right-click `AcademicFormatter.dotm` -> **Properties** -> if **Unblock** appears at the bottom of the *General* tab, check **Unblock** -> **OK**.
- **Uninstallation (`Uninstall.bat`):**
  1. Close all open Microsoft Word windows.
  2. Double-click **`Uninstall.bat`** to automatically locate and remove `AcademicFormatter.dotm` from your Word `STARTUP` folder (or delete `AcademicFormatter.dotm` manually from `%APPDATA%\Microsoft\Word\STARTUP`).

### 1.2. The 5 Pre-Flight Safety Gates
Before `Format All` or `Update Lists` modifies a single character in your document, it enforces 5 deterministic pre-flight checks:

| Gate # | Condition Checked | Why It Exists | Exact User Action Required if Blocked |
| :--- | :--- | :--- | :--- |
| **1** | **Saved to Disk (`doc.Path <> ""`)** | Unsaved documents (`Document1`) have no folder for physical crash backup. | Press **`Ctrl + S`** and save the `.docx` file to a local folder first. |
| **2** | **Local File Path (No `http://` or `https://` Cloud URL)** | When OneDrive or SharePoint opens a file as a remote URL moniker (`https://d.docs.live.net/...`), Windows FileSystemObject cannot write a local atomic backup alongside it. | Use `File -> Save a Copy` to a local disk folder (e.g., your local `Documents` or `Desktop` folder) or pause OneDrive cloud-only path redirection. |
| **3** | **Not Read-Only (`doc.ReadOnly = False`)** | A Read-Only document cannot be formatted or saved. | Close the file, remove the Read-Only attribute (or click *Enable Editing*), and reopen. |
| **4** | **Not Protected (`doc.ProtectionType = -1`)** | Password/form-protected documents reject DOM range mutations. | Go to `Review -> Protect -> Restrict Editing` and click **Stop Protection**. |
| **5** | **No Unsaved Edits (`doc.Saved = True`)** | Ensures the disk file copied for the physical backup is 100% identical to the in-memory document state before formatting starts. | Press **`Ctrl + S`** right before clicking `Format All` or `Update Lists`. |

### 1.3. Track Changes Handling, Physical Crash Backup & Master Undo
- **Track Changes (`TrackRevisions`):** If *Track Changes* is active, formatting thousands of paragraphs generates massive revision bloat and slows Word COM by 10x–50x. The engine prompts you to temporarily pause *Track Changes* during execution and **automatically restores** its state upon completion.
- **Ephemeral Physical Crash Backup (`Full Run`):**
  - Before `Format All` touches the document, it creates a physical copy on disk in the same folder:
    `<OriginalFileName>_crash_backup_YYYYMMDD_HHNNSS.docx`
  - **Normal Completion (Zero Silent Errors):** The temporary backup file is **automatically deleted** as soon as formatting finishes cleanly, leaving your folder clutter-free.
  - **Runtime Error or Minor Warning (`g_SilentErrors > 0`):** The backup file is **preserved on disk** and its exact path is displayed in the completion/error dialog.
- **Single-Step Master Undo (`Ctrl + Z`):**
  - Every major action (`Format All`, `Update Lists`, `Build Math`, `Finalize Math`, `Clean Markers`) wraps its entire execution inside a single Word `UndoRecord`. Pressing **`Ctrl + Z` once** reverts the entire operation in one step.

---

## Part 2: Deterministic Formatting & Syntax Grammar ("How to Write for 100% First-Pass Accuracy")

The engine does **not** guess intent from fonts or visual styling. It classifies every paragraph using strict, case-insensitive regular expressions. Follow the grammar below for 100% recognition on the first run.

### 2.1. Boundary & List Markers (The 5 System Markers)
Markers tell the engine where your main body starts/ends and where to generate lists.

| Marker Token (Case-Insensitive) | Required / Optional | Function |
| :--- | :--- | :--- |
| `=listContent()` | Optional | Generates the **Table of Contents (TOC / MỤC LỤC)** above this marker. |
| `=listTableHere()` | Optional | Generates the **List of Tables (LOT / DANH MỤC BẢNG BIỂU)** above this marker (auto-skipped if the document has 0 tables). |
| `=listFigureHere()` | Optional | Generates the **List of Figures (LOF / DANH MỤC HÌNH ẢNH)** above this marker (auto-skipped if the document has 0 figures). |
| `=contentStart()` | **MANDATORY** | Marks the **beginning of the Main Body**. Placed immediately before your first Chapter / Level 1 heading of the body. |
| `=contentEnd()` | **MANDATORY** | Marks the **end of the Main Body**. Placed immediately after your final body paragraph (typically right before `References` / `Tài liệu tham khảo` or `Appendix` / `Phụ lục`). |

#### Strict Marker Rules:
1. **Standalone Paragraph Rule:** Every marker MUST sit alone on its own paragraph (created with a hard `Enter`, not `Shift + Enter`). Nothing else may appear on that line.
2. **No Table Rule:** Markers MUST NOT be placed inside a Word Table cell.
3. **No Duplicates:** `=contentStart()` and `=contentEnd()` must each appear **exactly once**.
4. **Strict Ordering Rule:** All list markers (`=listContent()`, `=listTableHere()`, `=listFigureHere()`) MUST be placed **before** `=contentStart()`. Placing any list marker between `=contentStart()` and `=contentEnd()` triggers a **Marker Warning** and safely aborts execution:
   ```text
   [Cover Page / Title Page]
   [Acknowledgements / Abstract / Acronyms]
   =listContent()
   =listTableHere()
   =listFigureHere()
   =contentStart()
   [Main Body: Chapters 1..N, Tables, Figures, Equations]
   =contentEnd()
   [References / Tài liệu tham khảo]
   [Appendices / Phụ lục]
   ```

---

### 2.2. Headings Grammar (Levels 1 to 4)

#### A. Level 1 Headings (`ElementTypeLevel1`)
A paragraph is recognized as **Level 1** (Formatted as **UPPERCASE, Bold, Centered**, `SpaceBefore = 6pt`, `SpaceAfter = 12pt`, `FirstLineIndent = 0`) if its trimmed text matches any of the following **case-insensitive** patterns:

| Category | Exact English Keywords / Patterns | Exact Vietnamese Keywords / Patterns |
| :--- | :--- | :--- |
| **Front Matter** | `Abstract`<br>`Acknowledgment`, `Acknowledgments`, `Acknowledgement`, `Acknowledgements`<br>`Declaration`, `Declaration of Originality`<br>`Acronym`, `Acronyms`, `List of Acronyms`<br>`Abbreviation`, `Abbreviations`, `List of Abbreviations` | `Lời cảm ơn`, `Lời cám ơn`<br>`Lời cam đoan`<br>`Tóm tắt`<br>`Danh mục từ viết tắt`, `Danh mục các từ viết tắt`, `Danh mục chữ viết tắt`, `Danh mục các chữ viết tắt`, `Từ viết tắt`, `Chữ viết tắt`<br>`Danh mục ký hiệu và chữ viết tắt` |
| **Intro & Conclusion** | `Introduction`<br>`Conclusion`, `Conclusions` | `Mở đầu`, `Lời mở đầu`, `Giới thiệu`<br>`Kết luận`, `Kết luận và kiến nghị` |
| **Numbered Chapters & Parts** | `Chapter <ID>: <Title>` or `Chapter <ID>. <Title>`<br>*(where `<ID>` is alphanumeric: `1`, `2`, `I`, `II`, `A`, etc., followed by `:` or `.`, space, and title text)*<br>Example: `Chapter 1: Literature Review`<br>**Uppercase Roman Numerals (case-sensitive):** `<ROMAN>. <Title>` or `<ROMAN>: <Title>` (`I`–`XXXIX`, must be UPPERCASE, e.g. `I. Introduction`, `IV. Results`; lowercase `i.`, `iv.` are ignored) | `Chương <ID>: <Title>` or `Chương <ID>. <Title>`<br>`Phần <ID>: <Title>` or `Phần <ID>. <Title>`<br>Example: `Chương 1: Cơ sở lý luận`, `Phần I. Mở đầu`<br>**Số La Mã viết hoa (phân biệt hoa/thường):** `I. Tổng quan`, `IV. Kết quả` *(chữ thường `i.`, `iv.` không được tính)* |
| **References** | `Reference`, `References`<br>`Bibliography` | `Tài liệu tham khảo`<br>`Danh mục tài liệu tham khảo` |
| **Appendices** | `Appendix`, `Appendices`, `Appendixes`<br>`Appendix <ID>` (e.g., `Appendix A`, `Appendix 1`)<br>`Appendix <ID>: <Title>` or `Appendix <ID>. <Title>` | `Phụ lục`<br>`Phụ lục <ID>` (e.g., `Phụ lục A`, `Phụ lục 1`)<br>`Phụ lục <ID>: <Title>` or `Phụ lục <ID>. <Title>` |
| **Manual List Headers** *(when the corresponding `=list...()` marker is absent)* | `Table of Contents`, `Tables of Contents`<br>`List of Table`, `List of Tables`<br>`List of Figure`, `List of Figures` | `Mục lục`<br>`Danh mục bảng`, `Danh mục bảng biểu`, `Danh sách bảng`, `Danh sách bảng biểu`<br>`Danh mục hình`, `Danh mục hình ảnh`, `Danh mục hình vẽ`, `Danh sách hình`, `Danh sách hình ảnh`, `Danh sách hình vẽ`, `Danh mục sơ đồ` |

> **Manual List Header Safety Rule:** If your document contains a manual heading like `MỤC LỤC` or `TABLE OF CONTENTS` **and** also has the `=listContent()` marker, the engine highlights the manual header in **Yellow (`wdYellow`)** and skips it in the TOC so you don't get duplicate headers.

#### B. Levels 2, 3, and 4 Subheadings (`ElementTypeLevel2..4` — Main Body Only)
Subheadings are recognized **only inside the Main Body** (between `=contentStart()` and `=contentEnd()`).

> **CRITICAL RULE — THE TRAILING DOT:** Every numeric subheading prefix **MUST end with a dot (`.`) followed by at least one space and the title text**. Missing the final dot is the #1 user mistake!

| Level | Exact Regex Pattern | Formatting Applied | Valid Examples (`PASS`) | Invalid Examples (`FAIL` -> Treated as Normal Paragraph) |
| :--- | :--- | :--- | :--- | :--- |
| **Level 2** | `^(\d+\.\d+\.)\s+(.+)$` | **Bold**, Left-aligned, First-line indent (`1.27 cm` default), `SpaceBefore = 6pt`, `SpaceAfter = 6pt`, Auto Title-Case first letter | `1.1. Research Background`<br>`2.10. Kết quả thực nghiệm` | `1.1 Research Background` *(missing dot after `.1`)*<br>`1.1.` *(no title text)* |
| **Level 3** | `^(\d+\.\d+\.\d+\.)\s+(.+)$` | **Bold + Italic**, Left-aligned, First-line indent, `SpaceBefore = 6pt`, `SpaceAfter = 6pt` | `1.1.1. Theoretical Framework`<br>`3.2.1. Mô hình đề xuất` | `1.1.1 Theoretical Framework` *(missing final dot)* |
| **Level 4** | `^(\d+\.\d+\.\d+\.\d+\.)\s+(.+)$` | *Italic* (Not Bold), Left-aligned, First-line indent, `SpaceBefore = 6pt`, `SpaceAfter = 6pt` | `1.1.1.1. Primary Survey`<br>`2.1.3.1. Phân tích định tính` | `1.1.1.1 Primary Survey` *(missing final dot)*<br>`1.1.1.1.1. Too Deep` *(>4 levels not supported)* |

#### C. Soft Line Breaks (`Shift + Enter` / `^l`) Inside Headings & Captions
Before scanning, `NormalizeContentBreaks` inspects every paragraph containing a soft line break (`Shift + Enter` / `Chr(11)`):
- **Case 1 — Split Prefix (`Chapter 1:` + `Shift+Enter` + `Introduction`):** If the first line before `Shift + Enter` is **only a prefix** (e.g., `Chapter 1:`, `Chương 1.`, `1.1.`, `Table 2.1.`, `Figure 3.1:`, `Source:`), the engine automatically replaces `Shift + Enter` with a single space (`" "`) to merge them into one clean heading/caption line (`Chapter 1: Introduction`).
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
5. **Formatting Applied:** **Bold + Italic, Centered**, `SpaceBefore = 6pt`, `SpaceAfter = 6pt`, `FirstLineIndent = 0`, first letter of `<TitleText>` auto-capitalized.
6. **Layout-Table Exception for Figures:** Normally, paragraphs inside Word Table cells are formatted as `TableCell` and excluded from caption lists. **Exception:** If you place side-by-side images inside a layout table and put a valid Figure caption (`Figure 1.1: ...`) inside a cell of that same table (where `Table.Range.InlineShapes.Count > 0`), the engine recognizes and indexes the Figure caption into the List of Figures!
7. **Length Limit in Quick Update:** Captions up to `550` raw characters (`500` clean characters) are scanned in high-speed Quick Update mode. Keep captions concise and put source citations on the next paragraph (`Source:` / `Nguồn:`).

#### B. Source & Note Lines (`ElementTypeSourceNote`)
- **Exact Regex:** `^(Source|Nguồn|Note|Ghi chú):\s*(.*)$` (Case-insensitive; **colon `:` is mandatory**).
- **Formatting Applied:**
  - Whole line: *Italic* (Not Bold), Left-aligned, First-line indent (`1.27 cm` default), `Size = BodySize - 2 pt` (e.g., `11 pt` when body is `13 pt`), `SpaceBefore = 6pt`, `SpaceAfter = 12pt`.
  - Prefix (`Source:`, `Nguồn:`, `Note:`, `Ghi chú:`): **Bold + Italic**, first letter after colon auto-capitalized.
- **Common Mistake:** Writing `Nguồn (2025)` or `Note - Author calculation` without a colon `:` right after the keyword will cause it to be formatted as a Normal Paragraph. Always write `Nguồn: ...` or `Note: ...`.

---

### 2.4. LaTeX Math Engine Grammar (v1.0.0)

The built-in pure-VBA LaTeX-to-OMML compiler (`Build Math` / `Dựng Toán`) converts raw LaTeX expressions directly into native Microsoft Word `OMath` equations (`Cambria Math`) without external tools.

#### A. Supported Delimiters
| Mode | Delimiters | Behavior & Placement |
| :--- | :--- | :--- |
| **Inline Math** | `$...$` or `\(...\)` | Embedded inside a paragraph (up to `250` chars per `$...$` span). Rendered inline and shaded **Pastel Cyan** (`#DCEBF5`) for review. |
| **Display Block (Single-Line)** | `$$...$$` or `\[...\]` or `\begin{env}...\end{env}` | Occupies an entire paragraph (or a table cell). Rendered as a centered display equation and shaded **Pastel Mint** (`#DCF5E1`). |
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
| Category | Exact Environment Names | Rendering Details |
| :--- | :--- | :--- |
| **Piecewise / Cases** | `cases` | Left curly brace `{` with 2 left-aligned columns (`&` separates columns, `\\` separates rows). |
| **Multi-Line Alignment** | `align`, `align*`, `aligned`, `split` | Alternating `right, left, right, left` column alignment across `&` separators. |
| **Centered Multi-Line** | `gathered`, `multline`, `multline*`, `equation`, `equation*` | Centered multi-line or single-equation display blocks. |
| **Matrices & Determinants** | `matrix` *(no brackets)*<br>`pmatrix` `( ... )`<br>`bmatrix` `[ ... ]`<br>`Bmatrix` `{ ... }`<br>`vmatrix` `\| ... \|`<br>`Vmatrix` `‖ ... ‖`<br>`smallmatrix` | Up to arbitrary rows/columns separated by `&` and `\\`. Supports nested matrices inside `align`/`cases` and matrices >10 rows. |
| **Arrays (Borderless Only)** | `array` with `{l}`, `{c}`, `{r}` column specs (e.g., `\begin{array}{lcc} ... \end{array}`) | **Fail-Closed Rule:** If the column spec contains vertical borders `\|` (e.g., `{c\|c}`) or the body contains `\hline`, the engine refuses conversion and shades the block **Pastel Pink** so you can convert it to a real Word Table or remove `\|` and `\hline`. |

#### D. Supported LaTeX Commands & Macros (`ModTexParserCmds.bas`, `ModTexSymbols.bas`)
- **Fractions & Binomials:** `\frac{a}{b}`, `\dfrac{a}{b}`, `\tfrac{a}{b}`, `\cfrac{a}{b}`, `\binom{n}{k}`, `\dbinom{n}{k}`, `\tbinom{n}{k}`.
- **Radicals:** `\sqrt{x}`, `\sqrt[n]{x}`.
- **Subscripts, Superscripts & Primes:** `x_i`, `x_{i,j}`, `x^2`, `x^{a+b}`, `x_i^2`, `f'(x)`, `f''(x)`, `x^\prime`.
- **Big Operators (N-ary with `_` and `^` limits):** `\sum`, `\prod`, `\coprod`, `\int`, `\iint`, `\iiint`, `\oint`.
- **Delimiters (Auto-Scaling):** `\left`, `\right`, `\big`, `\Big`, `\bigg`, `\Bigg` paired with `(`, `)`, `[`, `]`, `\{`, `\}`, `|`, `\|`, `\Vert`, `\langle`, `\rangle`, `\lfloor`, `\rfloor`, `\lceil`, `\rceil`, or `.` (invisible delimiter).
- **Accents, Bars & Over-Vectors:** `\hat{x}`, `\widehat{ABC}`, `\bar{x}`, `\overline{AB}`, `\underline{xy}`, `\vec{v}`, `\overrightarrow{AB}`, `\tilde{x}`, `\dot{x}`, `\ddot{x}`.
- **Underbrace, Overset, Boxed & Labeled Arrows:** `\underbrace{expr}_{note}`, `\overset{top}{base}`, `\xrightarrow{label}`, `\xrightarrow[sub]{sup}`, `\boxed{expr}`, `\tag{1.1}`, `\displaystyle`.
- **Fonts & Text Spans:** `\text{...}`, `\mbox{...}`, `\textrm{...}`, `\mathrm{...}`, `\mathbf{...}`, `\bm{...}`, `\boldsymbol{...}`, `\mathbb{...}` (e.g., `\mathbb{R}, \mathbb{E}`), `\mathcal{...}` (e.g., `\mathcal{L}, \mathcal{N}`).
- **Named Operators & Limit Functions:** `\sin`, `\cos`, `\tan`, `\cot`, `\sec`, `\csc`, `\arcsin`, `\arccos`, `\arctan`, `\sinh`, `\cosh`, `\tanh`, `\ln`, `\log`, `\exp`, `\lim`, `\min`, `\max`, `\inf`, `\sup`, `\deg`, `\dim`, `\det`, `\ker`, `\hom`, `\arg`, `\argmin`, `\argmax`, `\operatorname{name}`, `\operatorname*{name}_{sub}` (places limits directly underneath).
- **Negations (`\not`):** `\not\in` (`∉`), `\not=` (`≠`), `\not\equiv` (`≢`), `\not\subset` (`⊄`), `\not\subseteq` (`⊈`), `\not\approx` (`≉`), `\not\sim` (`≁`), `\not<` (`≮`), `\not>` (`≯`), `\not\leq` / `\not\le` (`≰`), `\not\geq` / `\not\ge` (`≱`). *(Any unknown `\not\foo` fails closed to Pastel Pink).*
- **Greek Letters:** All lowercase (`\alpha` through `\omega`, plus `\varepsilon`, `\vartheta`, `\varpi`, `\varrho`, `\varsigma`, `\varphi`) and uppercase (`\Gamma`, `\Delta`, `\Theta`, `\Lambda`, `\Xi`, `\Pi`, `\Sigma`, `\Upsilon`, `\Phi`, `\Psi`, `\Omega`).
- **Relations, Arrows & Logic/Set Symbols:** `\leq`, `\le`, `\geq`, `\ge`, `\neq`, `\ne`, `\approx`, `\sim`, `\simeq`, `\cong`, `\equiv`, `\propto`, `\times`, `\div`, `\cdot`, `\pm`, `\mp`, `\in`, `\notin`, `\subset`, `\subseteq`, `\supset`, `\supseteq`, `\cup`, `\cap`, `\setminus`, `\emptyset`, `\varnothing`, `\forall`, `\exists`, `\nexists`, `\nabla`, `\partial`, `\infty`, `\triangle`, `\perp`, `\parallel`, `\Vert`, `\|`, `\mid`, `\circ`, `\oplus`, `\otimes`, `\to`, `\rightarrow`, `\leftarrow`, `\gets`, `\Rightarrow`, `\Leftarrow`, `\Leftrightarrow`, `\ldots`, `\dots`, `\cdots`, `\vdots`, `\ddots`, `\,`, `\:`, `\;`, `\ `, `\!`, `\quad`, `\qquad`.

---

### 2.5. Inline Cross-Reference Auto-Linking (v1.0.1)
During **`Format All (Định dạng)`**, the engine automatically scans body paragraphs, source/note lines, table cells, and Appendix prose to convert valid inline references into clickable internal hyperlinks (working in both Word via `Ctrl + Click` and exported PDFs):

- **Supported Reference Groups (Case-Insensitive):**
  - **Tables (`bm_tbl_*`):** `Table`, `Bảng`, `Bang` + ID (e.g., `Table 1.1`, `Bảng 2.1`, `Table A.1`)
  - **Figures (`bm_fig_*`):** `Figure`, `Fig.`, `Fig`, `Chart`, `Diagram`, `Hình`, `Sơ đồ`, `Biểu đồ`, `Đồ thị` + ID
  - **Headings & Appendices (`bm_toc_*`):** `Chapter`, `Chương`, `Section`, `Mục`, `Phần`, `Appendix`, `Phụ lục` + number or letter ID (plus standalone `Appendix` / `Phụ lục` when a matching standalone Level 1 heading exists)
- **Automatic Normalization & Styling:**
  - Capitalizes the first letter of the keyword (e.g., `table 1.1` -> `Table 1.1`, `bảng 2.1` -> `Bảng 2.1`) and strips stray trailing dots after numbers when followed by lowercase text or punctuation (while preserving sentence-ending periods).
  - Applies a clean, non-underlined link color configurable via **`Ref Link (Màu link)`** on the Ribbon (`Blue` default, `Orange`, `Red`, `Black`, or `Off`), plus a native Word `ScreenTip` showing the target title.
- **Safety Guards:** References pointing to non-existent targets are left untouched. Paragraphs inside `References` / `Tài liệu tham khảo` and TOC/LOT/LOF zones are never linked.

---

## Part 3: Document Zones & No-Fly Zones

### 3.1. Three Structural Document Zones (`Format All`)
The formatting engine divides your document into 3 zones relative to `=contentStart()` and `=contentEnd()`:

| Zone | Location | What Is Formatted | What Is Strictly Untouched (Protected) |
| :--- | :--- | :--- | :--- |
| **1. Pre-Content Zone (Front Matter)** | Paragraphs **before** `=contentStart()` | - **Level 1 Headings** (`Abstract`, `Acknowledgements`, `Lời cảm ơn`, `Danh mục từ viết tắt`, etc.)<br>- Generated **TOC / LOT / LOF** zones | - **Cover Page / Title Page** text, borders, and layout tables<br>- **Acronym / Abbreviation tables** and front-matter prose<br>- Line spacing and indentation of non-heading paragraphs |
| **2. Main Body Zone** | Paragraphs **between** `=contentStart()` and `=contentEnd()` | - **Headings Level 1–4**<br>- **Normal Paragraphs** (Font, Size, Justified, First-Line Indent, `SpaceBefore=6pt`, `SpaceAfter=6pt`, optional Line Spacing)<br>- **Table & Figure Captions** + **Source/Note lines**<br>- **Inline Pictures** (Centered, `FirstLineIndent=0`, `KeepWithNext` with following caption)<br>- **Data Table Cells** (`Size = BodySize - 2 pt`, `SpaceBefore=3pt`, `SpaceAfter=3pt`) | - **Native Word `OMath` equations** (kept in `Cambria Math`)<br>- **Raw LaTeX Display Blocks** (`$$...$$`, `\[...\]`)<br>- **Bulleted & Numbered Lists** (`ListFormat.ListType <> 0`)<br>- **Floating Shapes / TextBoxes / SmartArt / Charts**<br>- **Existing built-in Word TOC fields** (`doc.TablesOfContents`)<br>- **User Highlights** (Green, Turquoise, Pink, Gray, etc. — only system `wdYellow` and `wdRed` inside body text are cleared) |
| **3. Post-Content Zone (Back Matter & Appendices)** | Paragraphs **after** `=contentEnd()` | - **Level 1 Headings** (`References`, `Tài liệu tham khảo`, `Appendix A`, `Phụ lục 1`, etc.)<br>- **Appendix Table & Figure Captions** (`Table A.1: ...`, `Hình A.1: ...` — indexed into LOT/LOF!)<br>- **Appendix Source/Note lines**, **Appendix Table Cells**, and **Appendix Inline Pictures** | - **Reference / Bibliography list entries** (hanging indents, APA/IEEE italics, and line spacing are 100% preserved)<br>- Normal prose paragraphs in Appendices (line spacing not overwritten) |

### 3.2. Math Engine No-Fly Zones (`Build Math`)
When you click **`Build Math (Dựng Toán)`**, the LaTeX scanner automatically skips:
1. **Generated List Zones:** `bm_zone_toc` (TOC), `bm_zone_lot` (LOT), and `bm_zone_lof` (LOF).
2. **References / Bibliography Section:** As soon as the scanner encounters a standalone heading matching `References`, `Bibliography`, `Tài liệu tham khảo`, or `Danh mục tài liệu tham khảo` (`<= 45` chars), **math scanning pauses** so `$` signs or URLs/DOIs in bibliography citations are never touched.
3. **Automatic Resume at Appendices:** If an `Appendix` / `Appendices` / `Phụ lục` heading appears after `References`, **math scanning automatically resumes** so equations in your Appendices (`$$\sigma^2 = ...$$`) are built normally!

---

## Part 4: Ribbon Controls & Standard 4-Step Workflow

### 4.1. Map of the `Formatter Macro` Ribbon Tab
| Group | Control Label | Options / Action |
| :--- | :--- | :--- |
| **1. Markers (Thẻ Đánh Dấu)** | `Insert Marker (Chèn Thẻ)` | Inserts standalone markers at cursor (blocks duplicates and table-cell insertion):<br>• `TOC =listContent() (Mục lục)`<br>• `LOT =listTableHere() (DM Bảng)`<br>• `LOF =listFigureHere() (DM Hình)`<br>• `Start =contentStart() (Bắt đầu)`<br>• `End =contentEnd() (Kết thúc)` |
| **2. Typography (Kiểu chữ)** | `Font (Phông):`<br>`Size (Cỡ):`<br>`Spacing (Giãn):`<br>`Indent (Thụt):` | • `Times New Roman` *(default)*, `Arial`, `Calibri`<br>• `12 pt`, `13 pt (Standard / Chuẩn)` *(default)*, `14 pt`<br>• `Keep (Giữ nguyên)` *(default)*, `1.15 lines`, `1.3 lines`, `1.5 lines`<br>• `1.27 cm (0.5 in)` *(default)*, `1.00 cm`, `0.00 cm` |
| **3. List Config (Cấu hình DM)** | `Lang (Ngôn ngữ):`<br>`Bold All TOC (In đậm TOC)`<br>`Footer Page # (Số trang Footer)`<br>`Depth (Độ sâu):`<br>`Ref Link (Màu link):` | • `Vietnamese (Tiếng Việt)` -> `MỤC LỤC`, `DANH MỤC BẢNG BIỂU`, `DANH MỤC HÌNH ẢNH`<br>• `English (Tiếng Anh)` -> `TABLE OF CONTENTS`, `LIST OF TABLES`, `LIST OF FIGURES`<br>• `Bold All TOC` *(Checked = bold all TOC entries; Unchecked = only Level 1 bold in TOC)*<br>• `Footer Page #` *(Checked = reads formatted footer page numbers like `i, ii, 1, 2`; Unchecked = physical PDF page index)*<br>• `Level 1-4 (Cấp 1-4)` *(default)* or `Level 1-3 (Cấp 1-3)`<br>• `Blue` *(default)*, `Orange`, `Red`, `Black`, or `Off` *(controls inline cross-reference hyperlink color)* |
| **4. Math (Toán LaTeX)** | `Build Math (Dựng Toán)`<br>`Toggle / Edit (Đổi / Sửa)`<br>`Finalize Math (Chốt Toán)` | • **Build Math:** Batch-converts all LaTeX in the document to Word `OMath` with Pastel review shading.<br>• **Toggle / Edit:** Context-sensitive 4-state switch at the current cursor/selection.<br>• **Finalize Math:** Strips all Pastel shading and `bm_math_*` bookmarks (warns if unbuilt Pink/Amber items remain). |
| **5. Actions (Thực thi)** | `Format All (Định dạng)`<br>`Update Lists (Cập nhật DM)`<br>`Clean Markers (Xóa Thẻ)` | • **Format All:** Full formatting of body text, headings, captions, tables, figures + rebuilds TOC/LOT/LOF + auto-links inline cross-references.<br>• **Update Lists:** Fast (<2s) refresh of headings/captions and TOC/LOT/LOF page numbers **without touching body paragraphs**.<br>• **Clean Markers:** Removes all `=...()` paragraphs and finalizes math shading before PDF export. |

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

### 4.4. Recommended 4-Step End-to-End Workflow
1. **Step 1 — Insert Markers & Configure Ribbon:** Place `=listContent()`, `=listTableHere()`, `=listFigureHere()` in front matter, `=contentStart()` before Chapter 1, and `=contentEnd()` after your last body paragraph. Choose your Font, Size, Line Spacing, and List Language (`Vietnamese` or `English`) on the **`Formatter Macro`** tab.
2. **Step 2 — Build & Review Math (`Build Math`):** Click **`Build Math (Dựng Toán)`**. Inspect any **Pastel Pink** equations; click inside them and use **`Toggle / Edit (Đổi / Sửa)`** to fix typos.
3. **Step 3 — Run Full Formatting (`Format All`):** Save your file (`Ctrl + S`), then click **`Format All (Định dạng)`**. If you later add/edit headings or captions and want to keep manual body adjustments, use **`Update Lists (Cập nhật DM)`**.
4. **Step 4 — Finalize for Submission (`Finalize Math` + `Clean Markers`):** Click **`Finalize Math (Chốt Toán)`** and **`Clean Markers (Xóa Thẻ)`** to strip all Pastel review colors and `=...()` marker lines. Export to PDF!

---

## Part 5: Highlight/Shading Color Decoder & Troubleshooting Matrix

### 5.1. Complete Color Decoder (6 Diagnostic Colors)
The engine uses **Word Highlight (`Range.HighlightColorIndex`)** for document structure markers and **Paragraph/Character Background Shading (`Range.Shading.BackgroundPatternColor`)** for the Math Engine:

| Visual Color | Technical Constant / RGB | Where It Appears | Exact Meaning & Required Action |
| :--- | :--- | :--- | :--- |
| **Bright Yellow Highlight** | `wdYellow` (`HighlightColorIndex = 7`) | On `=contentStart()`, `=listContent()`, or a manual `MỤC LỤC` heading | **Active Marker / Skipped Duplicate Manual Header:** Shows active `=...()` markers (removed when you click `Clean Markers`), OR warns that a manual `MỤC LỤC` heading was skipped because `=listContent()` is also present. |
| **Bright Red Highlight** | `wdRed` (`HighlightColorIndex = 6`) | On an unclosed `$$` or `\[` paragraph during `Format All` | **Unclosed Legacy LaTeX Block:** `ModLatex` detected an opening `$$` or `\[` that never closed before the end of the document. Add the missing closing `$$` or `\]`. |
| **Pastel Cyan Shading** | RGB `(220, 235, 245)` / `#DCEBF5` (`16119260`) | On inline `OMath` equations (`$...$`, `\(...\)`) | **Success (Inline Math):** Inline LaTeX was cleanly transpiled to Word `OMath`. Cleared automatically when you click `Finalize Math` or `Clean Markers`. |
| **Pastel Mint Shading** | RGB `(220, 245, 225)` / `#DCF5E1` (`14808540`) | On display `OMath` equations (`$$...$$`, `\[...\]`, `\begin{...}`) | **Success (Display Block Math):** Display LaTeX was cleanly transpiled to centered Word `OMath`. Cleared automatically by `Finalize Math` or `Clean Markers`. |
| **Pastel Pink Shading** | RGB `(255, 230, 230)` / `#FFE6E6` (`15132415`) | On raw `$...$` or `$$...$$` text | **Ambiguous or Unsupported LaTeX (Fail-Closed):** Either an ambiguous number/unit (`$500$`, `$20 MT$`), unbalanced braces, bordered `array` (`{\|c\|c\|}` / `\hline`), mismatched `\begin`/`\end`, or unknown LaTeX command. Edit the LaTeX and click `Toggle / Edit`. |
| **Pastel Amber Shading** | RGB `(255, 242, 204)` / `#FFF2CC` (`13497343`) | On raw LaTeX text after clicking `Toggle / Edit` | **Edit Mode (Raw LaTeX Source):** You toggled a built `OMath` back to raw LaTeX to edit it. Click `Toggle / Edit` again when done to rebuild the `OMath`. |

---

### 5.2. Master Troubleshooting Matrix (Symptom -> Root Cause -> Deterministic Fix)

| # | Symptom Observed | Exact Root Cause | Deterministic Fix |
| :--- | :--- | :--- | :--- |
| **1** | Inline formula `\$x^2 + y^2\$` is completely ignored by `Build Math` | Pasting from ChatGPT/Claude/Notion with normal `Ctrl + V` copied escaped `\$` or invisible HTML spans. Escaped `\$` is treated as a literal dollar character. | Replace `\$` with `$` in Word (`Ctrl + H`: Find `\$`, Replace with `$`), then run `Build Math`. Next time, copy from a fenced `text` block and paste with **`Ctrl + Shift + V`**. |
| **2** | Subheading `1.1 Research Scope` or `2.1.3 Methodology` is formatted as a normal body paragraph and missing from the TOC | Missing the **mandatory trailing dot** after the final digit (`^(\d+\.\d+\.)\s+(.+)$`). | Add a period after the number: `1.1. Research Scope` and `2.1.3. Methodology`, then click `Update Lists`. |
| **3** | Table or Figure caption `Table 1 Revenue` is missing from the List of Tables / Figures | Missing the **mandatory separator** (`:`, `. `, or ` - `) after the number token (`1`). | Add a colon or dot+space: `Table 1: Revenue` or `Table 1. Revenue`, then click `Update Lists`. |
| **4** | `Source` or `Nguồn` line under a table/figure is formatted as normal 13pt text instead of 11pt italic | Missing the **mandatory colon (`:`)** immediately after `Source`, `Note`, `Nguồn`, or `Ghi chú`. | Change `Nguồn (2025)` to `Nguồn: (2025)` or `Source: Author's calculation`, then click `Update Lists`. |
| **5** | TOC / LOT / LOF header titles are generated in Vietnamese (`MỤC LỤC`) inside an English thesis | The Ribbon dropdown `Lang (Ngôn ngữ)` defaults to `Vietnamese (Tiếng Việt)`. | On the **`Formatter Macro`** tab, switch **`Lang (Ngôn ngữ)`** to **`English (Tiếng Anh)`** and click **`Update Lists (Cập nhật DM)`** (takes ~1 second). |
| **6** | `$500$` or `$20 MT$` is shaded **Pastel Pink** instead of being ignored or built | Pure numbers (`$500$`) or number + unknown unit with space (`$20 MT$`) are flagged as ambiguous (`Class 2`) so the engine doesn't silently guess whether you meant $500 USD or the constant $500$. | • If it is **currency/text**: write `$500 USD` or `500 USD`, or leave it—`Finalize Math` will remove the Pink shading and keep `$500$` as plain text.<br>• If it is a **math constant**: click inside `$500$` and click **`Toggle / Edit (Đổi / Sửa)`** to force-build it into an `OMath`. |
| **7** | A matrix or `array` block is shaded **Pastel Pink** | Word OMML `<m:m>` does not support internal vertical/horizontal table borders (`\begin{array}{\|c\|c\|}` or `\hline`), or `\begin{pmatrix}...\end{matrix}` tags are mismatched. | Remove `\|` from `{cc}` and remove `\hline` (or use `pmatrix` / `bmatrix`), then click `Toggle / Edit`. For bordered data grids, use a real Word Table. |
| **8** | TOC page numbers show `1, 2, 3...` when the footer shows `i, ii, iii...` (or vice versa) | `Footer Page # (Số trang Footer)` checkbox on the Ribbon does not match your desired page numbering mode, or your document has no section break between front matter and Chapter 1. | Keep **`Footer Page #`** checked to read Word's formatted section page numbers (`i, ii, 1, 2`), or uncheck it to use physical PDF page numbers (`1..N`), then click `Update Lists`. |
| **9** | Clicking `Finalize Math (Chốt Toán)` shows a warning: `Note: X Pink/Amber item(s) kept as raw text` | You had `X` equations still in **Pastel Pink** (unbuilt/ambiguous) or **Pastel Amber** (toggled to raw edit mode) when finalizing. | If unintentional, press **`Ctrl + Z`** once to undo `Finalize Math`, review the Pink/Amber items, convert them with `Toggle / Edit`, and click `Finalize Math` again. |
| **10** | Clicking `Format All` or `Update Lists` shows `Pre-Flight Failed: Unsaved changes detected` | The document has unsaved edits in memory (`doc.Saved = False`). | Press **`Ctrl + S`** to save the document, then click the button again. |

---

## Part 6: Ready-to-Use AI Prompt Templates for End-Users

Copy and paste any of the 3 prompts below into ChatGPT, Claude, or Gemini (**along with this `INSTRUCTION.md` file attached**) to automate draft preparation and troubleshooting.

### Prompt 1: Normalize a Raw Thesis/Report Draft for 100% First-Pass Accuracy
```text
I am using the Microsoft Word "Formatter Macro" (Academic Document Format Engine v1.0.0). I have attached its deterministic specification (`INSTRUCTION.md`).

Please normalize my raw draft below so it passes the macro with 100% accuracy on the first run:
1. Do NOT change my academic meaning, citations, or data.
2. Place `=listContent()`, `=listTableHere()`, `=listFigureHere()`, `=contentStart()`, and `=contentEnd()` on standalone lines in the exact valid order if I asked for a full document layout.
3. Normalize all Chapter / Level 1 headings and ensure EVERY Level 2, 3, and 4 subheading ends its numeric prefix with a mandatory trailing dot + space (e.g., `1.1. Title`, `1.1.1. Subtitle`).
4. Normalize all Table and Figure captions to use a mandatory colon or dot+space separator (`Table 1.1: Title` or `Hình 1.1: Tên hình`) and ensure all Source/Note lines start with `Source:`, `Note:`, `Nguồn:`, or `Ghi chú:`.
5. Normalize all mathematical expressions into clean LaTeX (`$...$` for inline, `$$...$$` for display blocks) using ONLY the commands and environments supported in Part 2.4 of `INSTRUCTION.md`. Never escape dollar signs as `\$`.
6. Output the entire result inside a single fenced `text` code block (` ```text ... ``` `) so I can copy and paste it into Word using `Ctrl + Shift + V`.

Here is my raw draft:
[PASTE YOUR DRAFT HERE]
```

### Prompt 2: Diagnose Why a Heading, Caption, or Formula Failed
```text
I am using the Word "Formatter Macro" (v1.0.0, specification attached in `INSTRUCTION.md`).
The following paragraph(s) or equation(s) in my Word document were either:
- Not recognized as a Heading / Table Caption / Figure Caption / Source line, OR
- Highlighted in Pastel Pink / Bright Red / Bright Yellow.

Inspect my text character-by-character against Part 2 and Part 5 of `INSTRUCTION.md`, explain the exact root cause in one sentence, and give me the corrected line(s) inside a fenced `text` block ready to paste via `Ctrl + Shift + V`:

[PASTE THE PROBLEMATIC PARAGRAPH OR LATEX HERE]
```

### Prompt 3: Convert Complex / Unsupported LaTeX into Engine-Compliant LaTeX
```text
I need to compile the following LaTeX equation(s) inside Microsoft Word using the pure-VBA LaTeX Math Engine v1.0.0 (specification attached in `INSTRUCTION.md`, Part 2.4).
Please rewrite the LaTeX below so it uses ONLY supported v1.0.0 environments, commands, and symbols (e.g., replace custom macros, `\frac12` shorthand, `\mbox` quirks, or bordered `array` constructs with supported `align`, `cases`, `pmatrix`, `bmatrix`, `\dfrac`, `\operatorname*`, etc.), without changing the mathematical meaning.
Output ONLY the corrected LaTeX inside a fenced `text` block with unescaped `$` / `$$` delimiters:

[PASTE YOUR LATEX EQUATIONS HERE]
```
