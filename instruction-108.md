<!-- AF-INSTRUCTION-SIGNATURE: AcademicFormatter -->
# AcademicFormatter v1.0.8 ("Admonition Update") — Reference Manual & AI System Prompt

> **What this tool is (and isn't):** AcademicFormatter is a local, rule-based Microsoft Word Desktop (Windows) add-in for Vietnamese and English essays, reports, and theses. Everything runs offline inside Word with no cloud uploads or API keys. It does **not** replace every formatting workflow or citation style on earth. Instead, it gives you predictable, composable building blocks so you can format, link, convert supported LaTeX math, or audit a manuscript without fighting Word manually.

---

## Part 0: AI Assistant Contract & User-First Guidance

When a user shares this manual with you (ChatGPT, Claude, Gemini, Copilot, etc.), treat it as your exact behavioral and grammar contract for the **`Formatter Macro`** tab in Word.

1. **Assume a First-Time User (Never Assume They Read This Manual):**
   - Users usually drop `instruction-108.md` into the chat without reading it and may not know the Ribbon layout yet.
   - Avoid unexplained jargon (`@STYLE`, `Sparse`, `OutlineLevel`, `Dry-Run`); explain what each action changes in plain language.
   - Whenever you ask the user to do something in Word, give exact click-by-click steps (e.g., *open the **Formatter Macro** tab -> click the arrow next to **AI Patch** -> select **Pre-Audit X-Ray & Lint** -> click **Yes** to copy -> paste the report back here*), move one clean step at a time, and remind them to press **`Ctrl + S`** before actions that require a saved file.
2. **Required Inputs Gate (NEVER Guess Without the Manuscript):**
   - If the user only sends `instruction-108.md` (or only an X-Ray report) and asks you to format a draft or generate any `AF-PATCH` block (`@REFS`, `@CITE`, `@FIX`, `@STRUCT`, `@TAG`, `@ACRONYMS`), **NEVER guess or invent a placeholder patch**. Politely ask them to **attach their `.docx` file (or paste the text/chapter they want worked on)** first, because `AF-PATCH` requires verbatim `CTX:` snippets and `References` prefixes from the real document.
   - Once they share the `.docx` (especially a merged team draft or pre-styled paper), briefly explain in one sentence that you can read the text but cannot reliably see visual font sizes (`14/13/12`), colors, or Word's built-in TOC—then give click-by-click steps to run **`Pre-Audit X-Ray & Lint`** and paste the Clipboard report back before finalizing a patch.
   - If the user only asks about installation, what a button does, or troubleshooting, answer directly without asking for their file.
3. **Composable Capabilities (No Rigid Pipeline):**
   - Do **not** force a fixed workflow or forbid Word's built-in Heading styles / native Word TOC fields. Match the tools to what the user actually has: fresh draft -> normalize text or `Format All`; merged team draft with Word TOC & stepped fonts (`14/13/12`) -> **`Pre-Audit X-Ray & Lint`**, **`Capture Style`**, **`AI Patch`** (no markers needed), and **`Link Cross-Refs Only`** (zero font/TOC/Outline changes); pre-submission review -> **`Pre-Audit X-Ray & Lint`** then targeted approved fixes.
4. **Mandatory Output & Copy-Paste Format:**
   - Output any normalized draft or `=== AF-PATCH v2 ===` block inside a **single fenced `text` code block** (` ```text ... ``` `) so the user can copy cleanly without Markdown HTML or escaped `\$`.
   - Remind the user to paste plain drafts into Word with **`Ctrl + Shift + V`** (*Keep Text Only*). Never output `\$` for LaTeX math or wrap inline math in backticks.
5. **Content, Language & Citation Guardrails:**
   - **Never** rewrite academic prose, alter meaning, or fabricate sources. Under `@STRUCT`, only fix structural prefixes/separators (`1.1` -> `1.1.`, `Table 1` -> `Table 1:`, `Nguồn` -> `Nguồn:`).
   - **Never** strip Vietnamese diacritics or Westernize Vietnamese names (`Nguyễn` -> `Nguyen`), and preserve the draft's dominant Vietnamese APA tokens (`và cs.`, `và nnk.`, `k.n.`, `không rõ năm`, `đang in`) unless fixing an internal mismatch approved by the user.
   - **Vietnamese Unicode Font Rule:** For `@STYLE font=...`, only suggest full-coverage Unicode fonts supporting extended Vietnamese (`U+1EA0–U+1EF9`, `ơ`, `ư`) across Regular/Bold/Italic: `Times New Roman`, `Arial`, `Calibri`, `Cambria`, `Verdana`, `Tahoma`, `Segoe UI`, `Consolas`, `Trebuchet MS`. Warn against fonts with incomplete Vietnamese italics (e.g., `Georgia Italic`).
   - **`Source:` / `Nguồn:` Rule:** Self-compiled lines (`Source: Author's compilation`, `Nguồn: Tác giả tự tổng hợp`, `Nguồn: Nhóm nghiên cứu`) have no bibliography entry -> **never** emit `@CITE`. External sources (`Source: Adapted from Vaswani et al. (2017)`, `Nguồn: WHO (2020)`) that exist in `References` **should** receive `@CITE`.

---

## Part 1: Installation, Ribbon Controls & Safety Nets

### 1.1. Installation & Removal
- **Platform:** Microsoft Word Desktop for Windows (**Word 2013 or newer**, 32/64-bit). *Not supported on macOS, Word Online, or mobile.*
- **Install / Update:** Close all Word windows, then double-click **`install-or-update.bat`** (downloads latest `.dotm` and `instruction-108.md` when online, falls back to local files offline, unblocks Windows flags, and copies to Word `STARTUP`).
- **Uninstall:** Close all Word windows and double-click **`Uninstall.bat`**.

### 1.2. Complete Map of the `Formatter Macro` Ribbon Tab

| Group           | Control                                                      | What It Does & Changes                                                                                                                                                                                                         | Preserved / Marker Requirement                                                                                                                                                 |
| :-------------- | :----------------------------------------------------------- | :----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | :----------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| **Markers**     | `Insert Marker (Chèn Thẻ)`                                   | Inserts standalone `=listContent()`, `=listTableHere()`, `=listFigureHere()`, `=listAcronyms()`, `=contentStart()`, `=contentEnd()` at cursor (blocks duplicates & table cells).                                               | Highlights markers bright yellow; removes cleanly later with `Clean Markers`.                                                                                                  |
| **Typography**  | `Font`, `Size`, `Spacing`, `Indent`                          | Sets baseline defaults (`Times New Roman`, `13 pt`, `Keep` / `1.15` / `1.3` / `1.5`, `1.27` / `1.00` / `0.00 cm`) for `Format All`, `Update Lists`, and `@ACRONYMS`.                                                           | Overridden per scope by any active `@STYLE` saved in the file.                                                                                                                 |
| **List Config** | `Lang`, `Bold All TOC`, `Footer Page #`, `Depth`, `Ref Link` | Sets list header language (`Vietnamese` / `English`), TOC bolding, footer vs physical page numbers, TOC depth (`Level 1-4` or `1-3`), and link color (`Blue`, `Orange`, `Red`, `Black`, `Off`).                                | `Ref Link = Off` clears cross-refs only and preserves existing `@CITE` links.                                                                                                  |
| **Math**        | `Build Math (Dựng Toán)`                                     | Converts supported LaTeX (`$...$`, `\(...\)`, `$$...$$`, `\[...\]`, `\begin{...}`) into native Word equations with Pastel review shading.                                                                                      | **No markers required** (stops at `References`, auto-resumes at `Appendix`/`Phụ lục`).                                                                                         |
| **Math**        | `Toggle / Edit (Đổi / Sửa)`                                  | 4-state switch at cursor/selection: (1) Built equation -> raw LaTeX (Amber); (2) Edited Amber -> rebuilt equation; (3) Pink ambiguous span -> forced equation; (4) Plain selected LaTeX -> built equation (`bm_math_<nnnnn>`). | Works anywhere in main body without markers.                                                                                                                                   |
| **Math**        | `Next Pink`                                                  | Jumps to and selects the next Pastel Pink math item needing review (wraps around).                                                                                                                                             | Read-only navigation; no markers required.                                                                                                                                     |
| **Math**        | `Finalize Math (Chốt Toán)`                                  | Removes all Pastel math shading and math bookmarks (warns if Pink/Amber items remain).                                                                                                                                         | Keeps internal undo data so `Ctrl + Z` can revert; no markers required.                                                                                                        |
| **Actions**     | `Format All (Định dạng)`                                     | Formats body text, headings (L1–L4), captions, source notes, tables, figures (applying `@STYLE`/`@TAG`), links cross-refs, syncs citation colors, and rebuilds TOC/LOT/LOF.                                                    | **Requires `=contentStart()` & `=contentEnd()`**. Preserves cover page, native Word TOC, bullets/lists, equations, user highlights (non-yellow/red), and `References` entries. |
| **Actions**     | `Update Lists (Cập nhật DM)`                                 | Refreshes headings, captions, cross-refs, link colors, and TOC/LOT/LOF page numbers (works even after `=list...()` markers were deleted).                                                                                      | **Requires `=contentStart()` & `=contentEnd()`**. Leaves body paragraphs 100% untouched; never overwrites a native Word TOC field.                                             |
| **Actions**     | `Link Cross-Refs Only` *(under `Update Lists`)*              | Adds clickable cross-reference links (`Bảng`, `Hình`, `Sơ đồ`, `Chương`, `Mục`, `Phụ lục`...) across body & appendices.                                                                                                        | **No markers required**. **Zero** font, size, OutlineLevel, or native Word TOC mutation.                                                                                       |
| **Actions**     | `Clean Markers (Xóa Thẻ)`                                    | Deletes all `=...()` paragraphs, finalizes math shading, and purges temporary math draft data before submission.                                                                                                               | Preserves `@TAG` bookmarks, list headers, headings, and all clickable links.                                                                                                   |
| **Actions**     | `AI Patch` *(Clipboard / From File...)*                      | Shows a Dry-Run preview, then applies `AF-PATCH v1/v2` (`@STYLE`, `@TAG`, `@STRUCT`, `@REFS`, `@FIX`, `@CITE`, `@ACRONYMS`).                                                                                                   | **No markers required** (infers body start automatically when markers are absent). Preserves images, tables, and equations.                                                    |
| **Actions**     | `Pre-Audit X-Ray & Lint` *(under `AI Patch`)*                | Read-only scan producing Document Topology, 8-scope Style Consistency Linter, Unmatched Visual Suspects, APA citation counts, and suggested Sparse `@STYLE`. Copies **100% full report** to Clipboard.                         | **Read-only, zero edits, no markers required.**                                                                                                                                |
| **Actions**     | `Capture Style` *(under `AI Patch`)*                         | Learns the majority font/size/style per scope from the current document and saves it as a Sparse `@STYLE` profile (automatically skips `TIE` conflict scopes).                                                                 | **No markers required.** Does not edit document text; preserves stepped heading sizes.                                                                                         |
| **Actions**     | `Reset Custom Style` *(under `AI Patch`)*                    | Clears saved `@STYLE` (and optionally `@TAG`), with an optional clean `Format All` pass to restore default black typography.                                                                                                   | Only way to clear saved `@STYLE` (since `@STYLE` persists outside `Ctrl + Z`).                                                                                                 |
| **Actions**     | `Export PDF (Xuất PDF)`                                      | 1-Click PDF export with clean **Heading Bookmarks** only (warns if markers or unfinalized math shading remain).                                                                                                                | Read-only export; falls back to local `Documents` folder for unsaved/OneDrive paths.                                                                                           |

### 1.3. Safety Nets & Pre-Flight Gates
- **5 Pre-Flight Gates** (checked before `Format All`, `Update Lists`, `Build Math`, and `AI Patch`; `Link Cross-Refs Only` only requires a writable document that is not Read-Only or Protected): (1) File is saved on disk; (2) Local disk path (not an `http://`/`https://` OneDrive/SharePoint URL); (3) Not Read-Only; (4) Not Protected/Restricted; (5) **No unsaved edits** — always press **`Ctrl + S`** right before running (including between `AI Patch` and `Format All`).
- **Track Changes, Crash Backup, Undo & Cancel:** Active *Track Changes* is paused and automatically restored when done. `Format All`, `Build Math`, and `AI Patch` create a local `<Name>_crash_backup_YYYYMMDD_HHNNSS.docx` before editing and auto-delete it on clean completion (preserved on error/warning). Every major action wraps in a **single-step `Ctrl + Z`** Undo record, and pressing **`ESC`** during long runs cancels safely and rolls back changes. *Note:* Saved `@STYLE` profiles live in the document outside Word's Undo stack—use **`AI Patch -> Reset Custom Style`** to remove them.

---

## Part 2: Deterministic Grammar (Structure, Zones & LaTeX Math)

### 2.1. The 6 System Markers
Required for **`Format All`** and **`Update Lists`** (`=contentStart()` and `=contentEnd()`). Optional for **`AI Patch`**, **`Pre-Audit X-Ray & Lint`**, **`Link Cross-Refs Only`**, and **`Build Math`**.

| Marker (Case-Insensitive) | Role                                                             | Placement Rules                                                                                                                                                                                      |
| :------------------------ | :--------------------------------------------------------------- | :--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `=listContent()`          | Generates TOC (`MỤC LỤC` / `TABLE OF CONTENTS`) above marker     | Must sit alone on its own paragraph (`Enter`, not `Shift + Enter`), never inside a table, no duplicates. List markers may sit before `=contentStart()` or after `=contentEnd()`, never between them. |
| `=listTableHere()`        | Generates LOT (auto-skipped if 0 tables)                         | Same as above                                                                                                                                                                                        |
| `=listFigureHere()`       | Generates LOF (auto-skipped if 0 figures)                        | Same as above                                                                                                                                                                                        |
| `=listAcronyms()`         | Anchor where `AI Patch` builds the List of Acronyms below marker | Same as above                                                                                                                                                                                        |
| `=contentStart()`         | Start of Main Body (place immediately before first body heading) | Required for `Format All` & `Update Lists`                                                                                                                                                           |
| `=contentEnd()`           | End of Main Body (place before `References` / `Appendices`)      | Required for `Format All` & `Update Lists`                                                                                                                                                           |

### 2.2. Headings (Levels 1–4), Captions & Source Lines

- **Level 1 Headings** (default: UPPERCASE, Bold, Centered, `LeftIndent=0`, `FirstIndent=0`, `KeepWithNext=1`, `OutlineLevel=1`):
  - *Front Matter:* `Abstract`, `Acknowledgment(s)`, `Acknowledgement(s)`, `Declaration`, `Declaration of Originality`, `Acronym(s)`, `List of Acronyms`, `Abbreviation(s)`, `List of Abbreviations`, `Lời cảm ơn`, `Lời cám ơn`, `Lời cam đoan`, `Tóm tắt`, `Danh mục (các) từ/chữ viết tắt`, `Từ/Chữ viết tắt`, `Danh mục ký hiệu và chữ viết tắt`.
  - *Intro & Conclusion:* `Introduction`, `Conclusion(s)`, `Mở đầu`, `Lời mở đầu`, `Giới thiệu`, `Kết luận`, `Kết luận và kiến nghị`.
  - *Chapters & Roman:* `Chapter <ID>:/. <Title>`, `Chương <ID>:/. <Title>`, `Phần <ID>:/. <Title>`, Uppercase Roman `I. <Title>`..`XXXIX. <Title>`.
  - *Single-Digit L1:* `1. MỞ ĐẦU`, `2. PHƯƠNG PHÁP` (recognized when carrying `OutlineLevel = 1` / `Heading 1` or `@TAG L1`, and supported by Cross-Ref via `mục 1`, `mục 2`).
  - *References & Subheadings:* `Reference(s)`, `Bibliography`, `(Danh mục) Tài liệu tham khảo`. Recognized subheadings inside References (`Tài liệu tiếng Việt`, `Tài liệu tiếng Anh`, `Tài liệu nước ngoài`, `Văn bản pháp luật`, `Nguồn Internet`, `Nguồn trực tuyến`, `Vietnamese/English References`, `Primary/Secondary Sources`, `Books`, `Journal Articles`, `Online Sources`) or `@TAG REF_SUB` stay inside the References zone without being counted as bibliography entries.
  - *Appendices:* `Appendix(es)`, `Appendix <ID>[:/. <Title>]`, `Phụ lục`, `Phụ lục <ID>[:/. <Title>]`. *(Sub-appendices like `Phụ lục 1..9` styled at `Heading 2` / `OutlineLevel = 2` are kept at L2 and still link via Cross-Ref).*
  - *Manual List Headers:* `Table(s) of Contents`, `List of Table(s)`, `List of Figure(s)`, `Mục lục`, `Danh mục/Danh sách bảng (biểu)`, `Danh mục/Danh sách hình (ảnh/vẽ)`, `Danh mục sơ đồ` — treated as L1 if the matching `=list...()` marker is absent; flagged **Red** and skipped if the marker is present.
- **Level 2–4 Subheadings (Main Body Only — Trailing Dot Mandatory):**
  - **L2:** `^(\d+\.\d+\.)\s+(.+)$` (Bold, Left, `LeftIndent=0.5cm`) — e.g., `1.1. Research Scope` (**NOT** `1.1 Research Scope`).
  - **L3:** `^(\d+\.\d+\.\d+\.)\s+(.+)$` (Bold + Italic, Left, `LeftIndent=0.75cm`) — e.g., `1.1.1. Framework`.
  - **L4:** `^(\d+\.\d+\.\d+\.\d+\.)\s+(.+)$` (Italic, Left, `LeftIndent=1.0cm`) — e.g., `1.1.1.1. Survey`.
  - *Word Multilevel Lists:* Headings numbered with Word's automatic Heading list numbering (`OutlineLevel <= 4`) have their list prefix read automatically.
  - *`Shift + Enter` Soft Breaks:* Prefix-only first lines (`Chapter 1:` + `Shift+Enter` + `Title`) keep the visual 2-line break in body and join cleanly in TOC; merged `Heading` + `Shift+Enter` + `Body prose` is split into two paragraphs.
- **Table & Figure Captions (Main Body & Appendices):**
  - Pattern: `^<Keyword>\s+<NumToken><Separator><TitleText>$`
  - *Table keywords:* `Table`, `Bảng`, `Bang`. (Bold, Left, `FirstIndent=Ribbon`, `KeepWithNext=1`).
  - *Figure keywords:* `Figure`, `Fig.`, `Fig`, `Chart`, `Diagram`, `Hình`, `Hinh`, `Hình ảnh`, `Hình vẽ`, `Sơ đồ`, `Biểu đồ`, `Đồ thị`. (Bold, Centered, `FirstIndent=0`, `KeepWithNext=0`; also recognized inside a small layout table containing inline images).
  - *Mandatory `<Separator>`:* `:`, `. ` (dot + space), or ` - ` / ` – ` / ` — ` (spaced dash). Prose like `"Table 4.2 shows..."` has no separator and is never misclassified.
- **Source & Note Lines:** `^(Source|Nguồn|Note|Ghi chú):\s*(.*)$` (colon `:` mandatory; Bold prefix + Italic body, Left, `Indent=0`).

### 2.3. Three Document Zones & Math No-Fly Zones

| Zone                                                  | Formatted by `Format All`                                                         | Protected (Untouched)                                                                                                                             |
| :---------------------------------------------------- | :-------------------------------------------------------------------------------- | :------------------------------------------------------------------------------------------------------------------------------------------------ |
| **1. Pre-Content** *(before `=contentStart()`)*       | L1 headings, generated TOC/LOT/LOF                                                | Cover page, layout tables, front-matter prose, acronym tables                                                                                     |
| **2. Main Body** *(`=contentStart()..=contentEnd()`)* | L1–L4, body paragraphs, captions, source notes, inline pictures, table cells      | Word `OMath`, raw LaTeX display blocks, bulleted/numbered lists, floating shapes/charts, native Word TOC fields, user highlights (non-yellow/red) |
| **3. Post-Content** *(after `=contentEnd()`)*         | L1 headings, `@TAG` headings, appendix captions/source notes/table cells/pictures | `References` list entries (hanging indents & italics preserved), appendix prose spacing                                                           |

- **Math No-Fly Zones (`Build Math`):** Skips generated list zones and automatically pauses at `References` / `Bibliography` / `Tài liệu tham khảo`, then **automatically resumes** if an `Appendix` / `Phụ lục` heading appears afterward.

### 2.4. Supported LaTeX Math Grammar (`Build Math`)
- **Delimiters:** Inline `$...$` (up to 250 chars) or `\(...\)`; single-line display `$$...$$`, `\[...\]`, `\begin{env}...\end{env}`; multi-paragraph display up to 40 paragraphs.
- **Smart `$...$` Currency Guard (3 States):**
  1. *Ignored as Plain Text:* Spaces inside `$ ... $` with no math tokens; `$a$5`; recognized currency/rate units (`usd`, `vnd`, `eur`, `gbp`, `jpy`, `cny`, `chf`, `k`, `m`, `b`, `tr`, `bn`, `dong`, `d`, `đồng`, `đ`, `nghin`/`nghìn`/`ngàn`, `trieu`/`triệu`, `ty`/`tỷ`, `dollar(s)`, `cent(s)`, `thousand`, `million`, `billion`, plus `/month`, `/mo`, `/year`, `/yr`, `/day`, `/unit`, `/tháng`, `/năm`, `/người`, `/đơn vị`) — e.g. `$500 triệu`, `$1,200/tháng`, `$100 USD`.
  2. *Valid Math (Built):* Balanced braces/parens, `<= 250` chars, not pure digits, with math proof (`\`, `=`, `<`, `>`, `^`, `_`, or variable + operator `+ - * /` such as `$0 < p < 0.05$`, `$1 + 2x$`) or no spaces after a leading digit.
  3. *Ambiguous / Unsupported (Pastel Pink):* Pure numbers (`$500$`), unknown units (`$20 MT$`), unbalanced delimiters, or unsupported commands. Click inside Pink text and press `Toggle / Edit` to force-build if valid.
- **Supported Environments:** `cases`, `align`, `align*`, `aligned`, `split`, `gathered`, `multline`, `multline*`, `equation`, `equation*`, `matrix`, `pmatrix`, `bmatrix`, `Bmatrix`, `vmatrix`, `Vmatrix`, `smallmatrix`, and **borderless** `array` (`{l}`, `{c}`, `{r}`). *Fail-Closed:* `array` with `|` borders (`{c|c}`) or `\hline` is shaded Pink—remove borders or use a Word Table.
- **Supported Commands & Symbols:**
  - *Structures:* `\frac`, `\dfrac`, `\tfrac`, `\cfrac`, `\binom`, `\dbinom`, `\tbinom`, `\sqrt`, `\sqrt[n]{x}`, `_`, `^`, primes (`'`, `''`, `\prime`), `\sum`, `\prod`, `\coprod`, `\int`, `\iint`, `\iiint`, `\oint` (with `\limits`/`\nolimits`).
  - *Delimiters:* `\left`/`\right`, `\big`/`\Big`/`\bigg`/`\Bigg` (`l`/`r`) with `()[]\{\}|`, `\|`, `\Vert`, `\langle\rangle`, `\lfloor\rfloor`, `\lceil\rceil`, `.`.
  - *Accents & Annotations:* `\hat`, `\widehat`, `\bar`, `\overline`, `\underline`, `\vec`, `\overrightarrow`, `\tilde`, `\dot`, `\ddot`, `\underbrace{e}_{n}`, `\overset{t}{b}`, `\xrightarrow[sub]{sup}`, `\boxed`, `\tag{...}`, `\label{...}`, `\displaystyle`, `\textstyle`, `\nonumber`, `\notag`.
  - *Fonts & Operators:* `\text`, `\mbox`, `\textrm`, `\mathrm`, `\mathbf`, `\bm`, `\boldsymbol`, `\mathbb`, `\mathcal`; `\sin`, `\cos`, `\tan`, `\cot`, `\sec`, `\csc`, `\arcsin`, `\arccos`, `\arctan`, `\sinh`, `\cosh`, `\tanh`, `\ln`, `\log`, `\exp`, `\lim`, `\min`, `\max`, `\inf`, `\sup`, `\deg`, `\dim`, `\det`, `\ker`, `\hom`, `\arg`, `\argmin`, `\argmax`, `\operatorname`, `\operatorname*`.
  - *Negations:* `\not\in`, `\not=`, `\not\equiv`, `\not\subset`, `\not\subseteq`, `\not\approx`, `\not\sim`, `\not<`, `\not>`, `\not\leq`/`\le`, `\not\geq`/`\ge`.
  - *Greek & Symbols:* Full lowercase/uppercase Greek (`\alpha..\omega`, `\varepsilon`, `\vartheta`, `\varpi`, `\varrho`, `\varsigma`, `\varphi`, `\Gamma..\Omega`); `\leq`, `\le`, `\geq`, `\ge`, `\neq`, `\ne`, `\approx`, `\sim`, `\simeq`, `\cong`, `\equiv`, `\propto`, `\times`, `\div`, `\cdot`, `\pm`, `\mp`, `\in`, `\notin`, `\subset`, `\subseteq`, `\supset`, `\supseteq`, `\cup`, `\cap`, `\setminus`, `\emptyset`, `\varnothing`, `\forall`, `\exists`, `\nexists`, `\nabla`, `\partial`, `\infty`, `\triangle`, `\perp`, `\parallel`, `\mid`, `\circ`, `\oplus`, `\otimes`, `\to`, `\rightarrow`, `\leftarrow`, `\gets`, `\Rightarrow`, `\Leftarrow`, `\Leftrightarrow`, `\ldots`, `\dots`, `\cdots`, `\vdots`, `\ddots`, `\,`, `\:`, `\;`, `\ `, `\!`, `\quad`, `\qquad`.

---

## Part 3: `AF-PATCH v2.1` Syntax & X-Ray / Style Lint Workflow

### 3.1. `Pre-Audit X-Ray & Lint` & `Capture Style`
- **`Pre-Audit X-Ray & Lint` (Read-Only):** Scans the file without making edits. Shows a compact summary dialog and offers to copy the **100% untruncated report** to your Clipboard:
  1. *Document Topology:* Explicit markers (or inferred body start), native Word TOC/TOF counts, table/equation counts, active `@TAG` / `@STYLE` counts, and **Unmatched Visual Suspects** (short bold/larger/outlined lines not yet matched by heading rules or `@TAG`).
  2. *8-Scope Style Consistency Linter:* Votes the dominant style for each scope (`L1..L4`, `TBL`, `FIG`, `SRC`, `BODY`), flags `TIE` conflicts, and lists every deviating paragraph with page number and exact property mismatch.
  3. *APA Citation & Reference Audit:* Counts parenthetical, narrative, and appendix citations plus reference entries (excluding reference subheadings) with a per-chapter breakdown.
  4. *Suggested Sparse `@STYLE`:* Ready-to-review `@STYLE` block omitting properties that already match defaults.
- **`Capture Style`:** Runs the same majority vote and saves clean scopes directly as a Sparse `@STYLE` profile inside the `.docx` file, automatically skipping any `TIE` scope so conflicts are never locked in by mistake.

### 3.2. `AF-PATCH v2.1` Flat Protocol (7 Sections)
Always wrap in `=== AF-PATCH v2 ===` ... `=== END ===` inside one `text` code block (never JSON). Omitting any section leaves that aspect untouched.

1. **`@STYLE`** (`SCOPE | prop=val | ...` or `RESET | ALL` / `RESET | SCOPE`):
   - **8 Scopes:** `L1`, `L2`, `L3`, `L4`, `TBL`, `FIG`, `SRC`, `BODY`.
   - **13 Properties:** `font` (1–40 chars), `size` (`8`–`36` pt), `color` (`#RRGGBB` or `auto`), `bold`, `italic`, `underline`, `upper`, `keepWithNext` (`0`/`1`), `align` (`left`/`center`/`right`/`justify`), `firstIndent`, `leftIndent` (`0`–`5` cm), `spaceBefore`, `spaceAfter` (`0`–`72` pt).
   - **Precedence:** Ribbon < `@STYLE` (merges incrementally per `.docx`; `L1` also styles generated list headers; `BODY` only runs during `Format All`) < `@TAG`. `L1 | upper=0` disables auto-uppercase without lowercasing text typed in ALL CAPS.
2. **`@TAG`** (`TYPE:` + `CTX:` on separate lines):
   - **12 Valid Types:** `L1`, `L2`, `L3`, `L4`, `TBL`, `FIG`, `SRC`, `REF` *(starts References zone, L1 in TOC)*, `REF_SUB` *(subheading inside References, L2 in TOC)*, `APP` *(starts Appendix zone after References, L1 in TOC)*, `SKIP` *(demotes false-positive heading to normal text / table cell)*, `CLEAR` *(removes prior tag)*.
   - Prioritizes exact whole-paragraph match (`CTX:` verbatim). Tags survive `Clean Markers` and can classify free-form headings before running `X-Ray & Lint`.
3. **`@STRUCT`** (`CTX:` + `FIND:` + `NEW:`): Fixes heading/caption/source prefixes across Pre-Content, Main Body, and Appendices (skips `References`, list zones, equations, fields, and hyperlinks). Requires 1 unique match.
4. **`@REFS`** (`key_id | prefix`): Exact starting prefix (`>= 12` chars recommended) of each entry inside `References`. Fails closed if duplicated (`Ambiguous ref (2+)` -> deduplicate `References` first).
5. **`@FIX`** (`CTX:` + `FIND:` + `NEW:`): Word-boundary scoped text fix inside Main Body, body tables, and Appendices (never touches Headings, Pre-Content, or `References`). Requires 1 unique match.
6. **`@CITE`** (`REF:` + `CTX:` + `FIND:` + optional `FIX:` / `MATCH:` / `NEAR:`):
   - Works across Main Body prose, data table cells, `Source:`/`Nguồn:` lines, and **Appendices after `References`**. Automatically trims outer `(...)` or `;` so only the citation text is linked.
   - **`MATCH: EXACT_ALL`** *(for repeated identical table cells)*: Links every table cell whose entire text equals `CTX:` verbatim (no ambiguity skip).
   - **`NEAR: <anchor>`** *(disambiguation)*: Restricts a repeated `CTX:` to the paragraph/cell immediately following a line or cell containing `<anchor>`.
7. **`@ACRONYMS`** (`ABBR | Full Form | Optional Col 3`): Deduplicates and sorts A–Z (2 cols if Col 3 empty; 3 cols otherwise; `\n` for line breaks; bolds `Ý nghĩa:`, `Nguồn:`, `Definition:`, `Meaning:`, `Source:`). Omitting `@ACRONYMS` preserves any existing acronym table.

### 3.3. Compact `AF-PATCH v2.1` Example
```text
=== AF-PATCH v2 ===
@STYLE
L1 | color=#1E6B3A | upper=0 | size=14
L2 | size=13 | leftIndent=0
BODY | size=12 | spaceAfter=6
@TAG
TYPE: L1
CTX: DANH SÁCH THÀNH VIÊN

TYPE: REF_SUB
CTX: Nhóm tài liệu nội địa

TYPE: SKIP
CTX: 1.1. Câu trích dẫn mẫu không phải tiêu đề
@STRUCT
CTX: 1.1 Cơ sở lý thuyết
FIND: 1.1
NEW: 1.1.
@REFS
ref_smith20 | Smith, J. (2020)
ref_load21 | The Loadstar. (2021)
@FIX
CTX: hệ thống IT và VIT trong
FIND: IT
NEW: CNTT
@CITE
REF: ref_smith20
CTX: bài viết (Smith 2020).
FIND: Smith 2020
FIX: Smith, 2020

REF: ref_load21
MATCH: EXACT_ALL
CTX: The Loadstar (2021)
FIND: The Loadstar (2021)

REF: ref_load21
NEAR: Kho bãi Rotterdam
CTX: Theo báo cáo The Loadstar (2021)
FIND: The Loadstar (2021)
@ACRONYMS
CNTT | Công nghệ thông tin |
LLM | Large Language Model | Mô hình ngôn ngữ lớn\nNguồn: Vaswani et al. (2017)
=== END ===
```

---

## Part 4: Diagnostic Color Decoder & Troubleshooting

### 4.1. Six Diagnostic Colors

| Color                           | Where It Appears                                                    | Meaning & Action                                                                                                                                                           |
| :------------------------------ | :------------------------------------------------------------------ | :------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| **Bright Yellow** *(Highlight)* | `=contentStart()`, `=contentEnd()`, `=list...()`                    | Active system marker. Removed automatically by `Clean Markers`.                                                                                                            |
| **Bright Red** *(Highlight)*    | Duplicate manual `MỤC LỤC` / `TABLE OF CONTENTS` when marker exists | Skipped duplicate header. Delete the redundant manual line.                                                                                                                |
| **Pastel Cyan** *(Shading)*     | Inline Word `OMath` (`$...$`, `\(...\)`)                            | Built inline math (`bm_math_<nnnnn>`). Cleared by `Finalize Math` or `Clean Markers`.                                                                                      |
| **Pastel Mint** *(Shading)*     | Display Word `OMath` (`$$...$$`, `\[...\]`, `\begin{...}`)          | Built display math (`bm_math_<nnnnn>`). Cleared by `Finalize Math` or `Clean Markers`.                                                                                     |
| **Pastel Pink** *(Shading)*     | Raw `$...$` or `$$...$$` text                                       | Ambiguous (`$500$`, `$20 MT$`) or unsupported LaTeX (bordered `array`, unknown macro). Edit or press `Toggle / Edit` to force-build; use `Next Pink` to jump between them. |
| **Pastel Amber** *(Shading)*    | Raw LaTeX toggled from a built equation                             | Edit mode. Fix the LaTeX and click `Toggle / Edit` again to rebuild.                                                                                                       |

*(Note: Interactive `Finalize Math` clears Pastel shading & math bookmarks while keeping undo data for `Ctrl + Z`; `Clean Markers` additionally purges all temporary math draft data for final submission.)*

### 4.2. Concise Troubleshooting Matrix

| #      | Symptom                                                         | Cause                                                               | Fix                                                                                                                      |
| :----- | :-------------------------------------------------------------- | :------------------------------------------------------------------ | :----------------------------------------------------------------------------------------------------------------------- |
| **1**  | `\$x^2\$` ignored by `Build Math`                               | Normal `Ctrl + V` pasted escaped `\$` from chat UI                  | Replace `\$` -> `$` (`Ctrl + H`), re-run `Build Math`. Next time paste with **`Ctrl + Shift + V`**.                      |
| **2**  | `1.1 Scope` treated as body text / missing from TOC             | Missing trailing dot after number                                   | Change to `1.1. Scope` (or fix via `@STRUCT`), then run `Update Lists`.                                                  |
| **3**  | `Table 1 Revenue` or `Nguồn (2025)` not recognized              | Missing mandatory separator (`:`, `. `, ` - `) or `:` after `Nguồn` | Write `Table 1: Revenue` and `Nguồn: (2025)`, then run `Update Lists`.                                                   |
| **4**  | `$500$` or bordered `array` shaded **Pink**                     | Ambiguous pure number or unsupported `\|` / `\hline` in `array`     | For currency, write `$500 USD`; for math constant, click `Toggle / Edit`; for bordered grids, use a Word Table.          |
| **5**  | `Pre-Flight Failed: Unsaved changes`                            | Document has unsaved edits in memory                                | Press **`Ctrl + S`**, then click the Ribbon button again.                                                                |
| **6**  | Ribbon `Font`/`Size` ignored by `Format All`                    | An active `@STYLE` profile is saved in the `.docx`                  | Click **`AI Patch -> Reset Custom Style`**, or update via `@STYLE`.                                                      |
| **7**  | `Capture Style` skips a scope (`TIE`)                           | Two different styles tied in frequency for that scope               | Check the `Pre-Audit X-Ray & Lint` report to unify the minority paragraphs, or set the scope explicitly via `@STYLE`.    |
| **8**  | `AI Patch` skips `@CITE`/`@FIX` (`Ambiguous 2+ hits`)           | `CTX:` matched multiple places in the document                      | Use a longer verbatim `CTX:`, add `NEAR: <preceding anchor>`, or use `MATCH: EXACT_ALL` for identical whole table cells. |
| **9**  | `AI Patch` skips `@REFS` (`Ambiguous ref 2+`)                   | Duplicate entries or multiple per-chapter `References` blocks       | Merge & deduplicate into one A–Z `References` list at the end, press `Ctrl + S`, then re-run `AI Patch`.                 |
| **10** | TOC shows Vietnamese headers in an English paper                | Ribbon `Lang` defaults to `Vietnamese`                              | Switch `Lang` to `English` on the Ribbon and click `Update Lists`.                                                       |
| **11** | TOC page numbers don't match Roman footer (`i, ii`)             | `Footer Page #` checkbox unchecked                                  | Check **`Footer Page #`** on the Ribbon and click `Update Lists`.                                                        |
| **12** | Need cross-refs on a teammate's file without changing fonts/TOC | `Format All` would re-format body and rebuild TOC                   | Use **`Update Lists -> Link Cross-Refs Only`** (zero font/TOC/Outline changes, no markers needed).                       |
| **13** | Exported PDF missing sidebar outline bookmarks                  | Saved via standard `Save As PDF` without heading option             | Click **`Export PDF (Xuất PDF)`** on the Ribbon (exports clean Heading Bookmarks only).                                  |

---

## Final Guidance for AI Assistants

Before emitting any `=== AF-PATCH v2 ===` block, confirm you have **both** the user's real `.docx`/text (they can send it multiple times to keep you updated) **and** their actual goal (light teamwork pass vs. full standardization). Guide them click-by-click in plain language, align on audit findings and formatting rules first, and only then output the patch inside one `text` block.
