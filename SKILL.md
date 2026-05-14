---
name: nwafu-paper
description: Apply NWAFU (西北农林科技大学) undergraduate thesis formatting standards to .docx files. Use whenever the user needs to format a paper, thesis, report, or academic document according to 西北农林科技大学本科毕业论文撰写规范, or when the user mentions NWAFU, 西农, 论文格式, or undergraduate thesis formatting. Also trigger when the user asks to check if a document follows NWAFU standards, or wants to create a new document that must comply with these standards.
---

# 西北农林科技大学本科毕业论文（设计）撰写规范

Based on the official NWAFU document `关于印发《西北农林科技大学本科毕业论文（设计）撰写规范与编排要求》的通知` (教务〔2026〕3号), referencing GB/T 7713.1-2025 and GB/T 7714-2025.

## Overview

Apply these formatting rules when creating or modifying .docx files for NWAFU undergraduate theses. When editing an existing document, unpack it first with `python scripts/office/unpack.py`, edit the XML directly, then repack with `python scripts/office/pack.py`. When creating a new document, use python-docx with the specifications below.

## 1. Page Setup

| Setting | Value |
|---------|-------|
| Paper size | A4 (21.0 cm x 29.7 cm) |
| Top margin | 3.0 cm (1701 DXA) |
| Bottom margin | 2.5 cm (1418 DXA) |
| Left margin | 2.6 cm (1474 DXA) |
| Right margin | 2.6 cm (1474 DXA) |
| Gutter | 0 cm |
| Gutter position | Left |
| Header distance | 2.0 cm from top (1134 DXA) |
| Footer distance | 1.75 cm from bottom (992 DXA) |
| Printing | Double-sided |

```python
# python-docx example
from docx.shared import Cm
section = doc.sections[0]
section.page_width = Cm(21.0)
section.page_height = Cm(29.7)
section.top_margin = Cm(3.0)
section.bottom_margin = Cm(2.5)
section.left_margin = Cm(2.6)
section.right_margin = Cm(2.6)
section.header_distance = Cm(2.0)
section.footer_distance = Cm(1.75)
```

## 2. Fonts

- **Chinese text**: 宋体 (SimSun) for body, 黑体 (SimHei) for headings
- **Western text and Arabic numerals**: Times New Roman
- **All text in the document** must use these font pairings — set both `w:ascii`/`w:hAnsi` (Times New Roman) and `w:eastAsia` (SimSun or SimHei) on every run.

```python
# Setting fonts on a run in python-docx
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml

def set_run_font(run, cn_font='SimSun', en_font='Times New Roman'):
    rPr = run._element.get_or_add_rPr()
    rFonts = parse_xml(
        f'<w:rFonts {nsdecls("w")} w:ascii="{en_font}" w:hAnsi="{en_font}" w:eastAsia="{cn_font}"/>'
    )
    rPr.insert(0, rFonts)
```

## 3. Title (题名)

- **Font**: 黑体 (SimHei), 二号字 (22pt), bold, centered
- **Line spacing**: 单倍行距 (single, 240 auto)
- **Limit**: No more than 25 Chinese characters
- Subtitle: use em-dash between main title and subtitle
- English title: no more than 10 content words, consistent with Chinese title

## 4. Abstract (摘要)

**Chinese abstract heading "摘要":**
- 黑体 (SimHei), 三号字 (16pt, sz=32), centered
- 段前空2行 (before=800), 段后空1行 (after=400)
- Fixed line spacing 20pt (400 exact)

**Chinese abstract body:**
- 宋体 (SimSun), 小四号字 (12pt, sz=24)
- 两端对齐 (justified), NO first-line indent
- Fixed line spacing 20pt (400 exact)
- 300-500 characters

**Keywords "关键词：":**
- Below abstract body, one blank line gap
- Label "关键词：" in 宋体 (SimSun), 五号字 (10.5pt, sz=21), **bold**, 左对齐顶格
- Keywords in 宋体 (SimSun), 五号字 (10.5pt), not bold
- 3-5 keywords (max 8), separated by English semicolons `;`, no punctuation after last keyword

**English abstract (if applicable):**
- Heading "ABSTRACT" — same formatting as Chinese abstract heading
- Body — 小四号字 (12pt), justified, fixed 20pt
- "KEY WORDS:" — 五号字 (10.5pt), bold, left-aligned
- English abstract and keywords must match Chinese content

## 5. Table of Contents (目次)

- Page starts on odd page (right side)
- Chapter titles: 黑体 小四号 (12pt), fixed 20pt, before 6pt
- Other entries: 宋体 小四号 (12pt), fixed 20pt, before 0pt after 0pt
- List only primary and secondary headings
- Chapter titles 顶格, first-level headings indent 1 char, second-level headings indent 2 chars
- Title and page number connected by dot leaders `……`, page number right-aligned

## 6. Body Text — Heading Hierarchy

Numbering format: 第一章 / 1.1 / 1.1.1 / 1.1.1.1

| Level | Numbering | Font | Size | Align | Before | After | Line Spacing |
|-------|-----------|------|------|-------|--------|-------|-------------|
| Chapter title (章标题) | 第一章 | 黑体 SimHei | 三号 16pt (sz=32) | Center | 2 lines (40pt) | 1 line (20pt) | Fixed 20pt |
| First-level (一级) | 1.1 | 黑体 SimHei | 小三号 15pt (sz=30) | Left | 1 line (20pt) | 0.5 line (10pt) | Fixed 20pt |
| Second-level (二级) | 1.1.1 | 黑体 SimHei | 四号 14pt (sz=28) | Left | 0.5 line (10pt) | 0pt | Fixed 20pt |
| Third-level (三级) | 1.1.1.1 | 黑体 SimHei | 小四号 12pt (sz=24) | Left | 0pt | 0pt | Fixed 20pt |

- Chapter title and number separated by 1 Chinese character space
- The following sections are treated as chapter-level headings: 原创性声明, 摘要, 目次, 参考文献, 致谢, 附录, 作者简历
- In practical use, if the paper uses 一、二、三… numbering instead of 第一章, these are treated as first-level headings
- "参考文献" heading: treat as chapter-level (before=40pt, after=20pt, same as 章标题)

## 7. Body Text — Paragraphs

- **Font**: 宋体 (SimSun), 小四号字 (12pt, sz=24)
- **Alignment**: 两端对齐 (justified)
- **First-line indent**: 2 Chinese characters (~480 twips for 12pt)
- **Line spacing**: Fixed 20pt (400 exact)
- **Before/after**: 0pt (段前段后不空行)
- **No grid** (无网格)

Exception: if a paragraph contains formulas, line spacing may be adjusted to accommodate.

## 8. Figures, Tables, and Formulas

**General rules:**
- Number using Arabic numerals per chapter: 图1-1, 图1-2, 表2-1, 表2-2, (3-1), (3-2)
- 0.5 line space between figure/table/formula and surrounding text
- Cross-reference by number — never use "如上表" or "如左图"
- Cite source if adapted from literature

**Figures (图):**
- Centered. Figure caption (图题) centered below the figure
- Caption: 宋体 五号 (10.5pt), centered, before 0.5 line, after 0
- Figure text: 宋体 五号 (10.5pt), fixed 18pt line spacing
- Natural science disciplines: use Chinese-English bilingual captions
- Multi-part figures: label (a), (b), (c)…
- Must use vector graphics for diagrams, flowcharts, data plots (not raster)

**Tables (表):**
- Must use 三线表 (three-line table), width matches text
- Top and bottom borders: 1.5pt; inner lines: 1pt
- Table caption (表题) centered above the table
- Table text: 五号字 (10.5pt)
- Data source note below table: 宋体 五号 (10.5pt), fixed 18pt
- Columns must specify "quantity/test item, symbol, unit"
- Numbers in same column: decimal-aligned, consistent precision
- Continued tables: add "（续）" after number, repeat header row
- Multi-part tables: (a), (b), (c)… left-to-right, top-to-bottom

**Formulas (公式):**
- Centered on separate line
- Break long formulas after `=` or `+`/`-`
- Formula number right-aligned, in full-width parentheses (3-1)
- "式中" explanation: indent 2 chars, semicolons between terms, period at end
- Explanation order: left-to-right, top-to-bottom

## 9. Footnotes

- Use ① ② ③… superscript numbering, per-page counting (reset each page)
- Footnote text: 宋体 小五号 (9pt), first-line indent 2 chars, single line spacing
- Before/after: 0pt

## 10. References (参考文献)

**Citation style**: 著者-出版年制 (Author-Year)

**Reference list formatting:**
- Font: 宋体 (SimSun), 五号字 (10.5pt, sz=21)
- **No numbering** — do not prefix with [1], [2], etc.
- Hanging indent: 2 Chinese characters (~480 twips)
- Line spacing: Fixed 16pt (320 exact)
- Before: 3pt (60), After: 0pt
- Punctuation: half-width (English) punctuation

**Sorting order:**
1. Sort by language: Chinese → Japanese → Western → Russian → other
2. Within each language: sort by first author surname (Chinese: pinyin alphabetical; Western: letter-by-letter), then by publication year

**Common entry formats (GB/T 7714-2015):**
```
Monograph:    Author. Title (Edition)[M]. Place: Publisher, Year.
Journal:      Author. Title[J]. Journal, Year, Vol(Issue): Pages.
Dissertation: Author. Title[D]. Place: Institution, Year.
Report:       Author. Title[R]. Place: Institution, Year.
Online:       Author. Title[EB/OL]. URL, Date.
Standard:     Issuing body. Title[S]. Place: Publisher, Year.
Government:   Issuing body. Title[Z]. Place: Issuing body, Year.
```

**In-text citation format:**
- Single author: (Author, Year)
- Two authors: (Author1 and Author2, Year)
- Three+ authors: (Author1 et al., Year)
- Multiple citations: (Author1, Year; Author2, Year)

## 11. Appendix (附录)

- Numbered as 附录 A, 附录 B, 附录 C… (always include A even if only one)
- Each appendix has a title
- Appendix content format same as body text
- Figures/tables/formulas renumbered with appendix prefix: 图A.1, 表B.2, 式(C-3)

## 12. Acknowledgments (致谢) and Author Biography (作者简历)

- "致谢" heading: chapter-title level formatting
- "作者简历" heading: chapter-title level formatting
- Content format same as body text

## 13. Headers and Footers

**Header (页眉):**
- Present on every page from abstract onward (except blank pages)
- Single-line header rule
- Odd pages (正文): corresponding chapter/section name
- Even pages: "西北农林科技大学本科毕业论文" (or "西北农林科技大学本科毕业设计")
- Font: 宋体 (SimSun), 五号字 (10.5pt), centered

**Page numbers (页码):**
- Pre-content pages (摘要 through 主要符号注释表): uppercase Roman numerals Ⅰ, Ⅱ, Ⅲ…
- Body text onward: Arabic numerals 1, 2, 3…
- Position: page bottom, 五号字 (10.5pt), centered
- No decorative lines around numbers

## 14. Binding

**Order:**
1. Cover (封面)
2. Title page (题名页)
3. Funding statement (资助声明页, optional)
4. Originality statement and usage authorization (原创性声明和使用授权说明)
5. Chinese/English abstract pages (中英文摘要页)
6. Table of contents (目次页)
7. List of figures and tables (插图和附表清单, if applicable)
8. List of symbols (主要符号注释表, if applicable)
9. Glossary (术语表, if applicable)
10. Body text (正文)
11. References (参考文献)
12. Appendix (附录, optional)
13. Acknowledgments (致谢)
14. Author biography (作者简历)

**Spine:** Required if printed pages >= 30. Text in 宋体 五号字, containing: thesis title, author name, university name.

## 15. Quick Implementation Notes

When creating a DOCX with python-docx:

```python
# Line spacing: fixed 20pt = 400 twips
def set_line_exact(para, pt_val):
    pPr = para._element.get_or_add_pPr()
    sp = pPr.find(qn('w:spacing'))
    if sp is None:
        sp = parse_xml(f'<w:spacing {nsdecls("w")}/>')
        pPr.insert(0, sp)
    sp.set(qn('w:line'), str(pt_val * 20))
    sp.set(qn('w:lineRule'), 'exact')

# First-line indent: 2 chars = font_size * 2 * 20 twips
# For 12pt: 12 * 2 * 20 = 480
def set_first_indent(para, chars=2, pt=12):
    val = int(chars * pt * 20)
    pPr = para._element.get_or_add_pPr()
    ind = pPr.find(qn('w:ind'))
    if ind is None:
        ind = parse_xml(f'<w:ind {nsdecls("w")}/>')
        pPr.insert(1, ind)
    ind.set(qn('w:firstLine'), str(val))

# Hanging indent for references
def set_hanging(para, chars=2, pt=10.5):
    val = int(chars * pt * 20)
    pPr = para._element.get_or_add_pPr()
    ind = pPr.find(qn('w:ind'))
    if ind is None:
        ind = parse_xml(f'<w:ind {nsdecls("w")}/>')
        pPr.insert(1, ind)
    ind.set(qn('w:left'), str(val))
    ind.set(qn('w:hanging'), str(val))

# Spacing before/after in points (1pt = 20 twips)
def set_spacing(para, before_pt, after_pt):
    pPr = para._element.get_or_add_pPr()
    sp = pPr.find(qn('w:spacing'))
    if sp is None:
        sp = parse_xml(f'<w:spacing {nsdecls("w")}/>')
        pPr.insert(0, sp)
    if before_pt is not None:
        sp.set(qn('w:before'), str(before_pt * 20))
    if after_pt is not None:
        sp.set(qn('w:after'), str(after_pt * 20))
```

**Font size reference:**
| Name | Points | sz (half-pt) | Usage |
|------|--------|-------------|-------|
| 二号 | 22pt | 44 | Title |
| 三号 | 16pt | 32 | Chapter title, Abstract heading |
| 小三号 | 15pt | 30 | First-level heading |
| 四号 | 14pt | 28 | Second-level heading |
| 小四号 | 12pt | 24 | Body text, Third-level heading |
| 五号 | 10.5pt | 21 | References, Keywords, Captions |
| 小五号 | 9pt | 18 | Footnotes |

**DXA conversion reference:**
- 1 inch = 1440 DXA
- 1 cm ≈ 567 DXA
- 1 pt = 20 DXA (for spacing values)
