---
name: ms-office-com-skill
description: "Use when editing MS Word files (.docx/.doc) on Windows via the COM interface (win32com.client). Specialized for Chinese government documents (公文): GB/T 9704-2012 formatting (仿宋_GB2312 + Times New Roman fonts, alignment, tables, page numbers), GB/T 7714 references, semantic superscript/subscript, PDF export with heading bookmarks, auto font install, dialog-hang prevention, and the real pitfalls. ONLY valid for Windows + MS Word COM — NOT for Linux/macOS, WPS, Excel/PowerPoint, or pure-library (python-docx) approaches."
version: 2.0.0
author: zergfff
license: MIT
platforms: [windows]
metadata:
  tags: [Word, COM, win32com, Office, docx, editing, windows-only, chinese-documents, 公文]
---

# Word COM Automation — 中文公文文档处理 (Chinese Government Documents)

## ⚠️ SCOPE — 本技能仅适用于 Windows + Microsoft Word COM（中文文档）

**本技能只适用于以下组合：Windows 操作系统 + Microsoft Word + `win32com.client` COM 接口，专用于中文公文/报告文档（.docx/.doc）。其他任何组合一律不得使用本技能：**

- ❌ **Linux / macOS** — COM 是 Windows 专属，本技能所有代码在这些平台不可用。
- ❌ **WPS Office** — ProgID 不同（`Kwps.Application` 等），对象模型行为有差异，本技能不覆盖。
- ❌ **Excel / PowerPoint** — 本技能只覆盖 Word；Excel/PPT 的 COM 操作不在范围内。
- ❌ **python-docx 等纯库** — 本技能是 COM 路线；除非用户明确点名用这些库，否则禁止使用它们替代 COM。
- ❌ **Microsoft Graph / Office 365 云端 API** — 那是不同的接口体系（OAuth + Graph），与本技能无关。

**适用判断：** 用户在 Windows 上要求操作本地 Word 文件（.docx/.doc），且机器装有 MS Word → 用本技能。Excel/PPT 或其他平台 → 告知用户本技能不适用，不硬套。

On Windows, drive MS Office applications directly through their COM object model with
`win32com.client` (pywin32). This is the **only** path that preserves complex structure —
merged tables, styles, outline levels, headers/footers, fields, shapes — that pure-library
approaches (python-docx) flatten or lose.

Do NOT use python-docx unless the user explicitly names it.
If COM fails, report the native Windows error — do not silently fall back.

## When to Use

- Read, edit, audit, or mass-edit an existing `.docx` / `.doc`.
- Fix formatting, numbering, styles, outline levels, tables, budgets, or cross-document consistency.
- Generate a new document with exact structure (tables, headings, merged cells, images).

## Golden Rules (each cost real time)

1. **Backup before every edit round** — `cp file.docx file_bak.docx` in the same folder. Restoration is the only reliable undo after corruption.
2. **Never run a Find.Execute replace-loop whose replacement text CONTAINS the search text.** The loop re-matches its own output → infinite duplication (one incident inserted 2,282 copies of a TOC block). Use paragraph-level iteration when `new` contains `old`.
3. **After any COM timeout/crash, kill the Office process before reopening.** A stale WINWORD.EXE keeps the corrupted doc in memory; the next `Documents.Open` reuses that instance and `Save()` writes the corruption to disk. Kill with `cmd //c "taskkill /F /IM WINWORD.EXE"` (from git-bash) or `taskkill //F //IM WINWORD.EXE`.
4. **Verify by re-reading, never trust "saved ok".** Reopen read-only, count key phrases (`t.count('x')`), check char/table counts vs. baseline, confirm headings appear exactly once.
5. Set `DisplayAlerts = 0`, always `Close(False)` + `Quit()` in a finally-style flow.

## Environment

- **Windows 11 + Office 2024 (LTSC)** — primary target. All patterns below work unchanged on Office 2016–2024; the Word COM object model has been stable across these versions.
- Python 3.x + pywin32 (`pip install pywin32`). Use `pythoncom.CoInitialize()` in worker threads before touching COM.
- `gencache.EnsureDispatch('Word.Application')` gives early binding (typed properties, faster, more reliable constants) at the cost of a one-time type-cache build; plain `Dispatch` is late binding. Use `EnsureDispatch` for complex scripts.

## Core Patterns

### Word

```python
import win32com.client
w = win32com.client.DispatchEx('Word.Application'); w.Visible = False; w.DisplayAlerts = 0
d = w.Documents.Open(path, ReadOnly=True, AddToRecentFiles=False)   # ReadOnly=True to inspect
text = d.Content.Text                    # full text incl. tables (\x07 cell marks)
# write full text to a cache .txt for big docs; terminal output truncates ~72K chars
d.Close(False); w.Quit()
```

Edit: `ReadOnly=False`, work, then `d.Save(); d.Close(False); w.Quit()`.

> **一致性说明：** 全 skill 统一用 `DispatchEx`（独立进程，防弹窗卡死、防误连已开实例）。只有明确要复用已打开实例的场景才用 `Dispatch`；复杂脚本用 `gencache.EnsureDispatch`。三者的详细区别见「Dispatch vs DispatchEx vs EnsureDispatch」章节。

### Word — export to PDF (公文交付)

```python
d.ExportAsFixedFormat(OutputFileName=pdf_path, ExportFormat=17)   # 17 = wdExportFormatPDF
```

**Rule: 另存为 PDF 时默认勾选"创建书签时使用 → 标题"（CreateBookmarks=1, wdExportCreateHeadingBookmarks）。** 这样 PDF 左侧导航栏会按文档标题生成书签层级，长文档翻阅方便。实测 Office 2024 可用（生成 PDF 含 `/Outlines` + 每个标题一个 `/Title` + `/Dest`）：

```python
d.ExportAsFixedFormat(OutputFileName=pdf_path, ExportFormat=17, CreateBookmarks=1)  # 1=按标题建书签
# CreateBookmarks: 0=不创建书签, 1=wdExportCreateHeadingBookmarks(按标题), 2=wdExportCreateNoHeadings(按文档结构)
```

- **前提：文档标题必须用 Heading 样式**（`para.Style = -2/-3/-4`，OutlineLevel 1-9），书签才按标题生成；正文样式（OutlineLevel=10）不会成为书签。
- 导出前先 `for toc in d.TablesOfContents: toc.Update()` + `d.Fields.Update()` + `d.Repaginate()`，保证 PDF 页码和书签正确。
- 只读检查完的文档（ReadOnly=True）也可以导出 PDF，无需写回。

### Word — refresh TOC, update fields, headers/footers

After any heading/paragraph edits, the TOC page numbers and cross-reference fields are stale. Refresh them before saving/exporting:

```python
for toc in d.TablesOfContents: toc.Update()          # refresh each TOC
d.Fields.Update()                                     # update all fields (page numbers, cross-refs)
d.Repaginate()                                        # recalc page layout
```

Headers/footers are per-section (公文常有多节，页码从正文重起):

```python
sec = d.Sections(1)
hdr = sec.Headers(1)                                  # 1 = wdHeaderFooterPrimary; 2 = even, 3 = first
hdr.Range.Text = '晋城市生态环境局'
ftr = sec.Footers(1)
ftr.Range.Text = '第 '                     # then insert PAGE field:
ftr.Range.Fields.Add(ftr.Range, -1, 'PAGE \\* MERGEFORMAT')
```

### Word — comments (批注) — expert-review workflows

Read all comments (专家意见常以批注形式给):

```python
for c in d.Comments:                     # 1-based collection
    print(c.Index, c.Author, c.Range.Text)
```

Add a comment anchored to a range; delete comments:

```python
rng = d.Range(start_pos, end_pos)
d.Comments.Add(rng, '修改意见文本')
d.Comments(i).Delete()
```

### Word — track changes / accept revisions

```python
d.TrackRevisions = True                   # record edits as revisions
# ... make edits ...
d.AcceptAllRevisions()                    # accept everything (or d.RejectAllRevisions())
d.TrackRevisions = False
```
Remember: `d.Content.Text` includes revision marks (`InsertedText`/`DeletedText` markers) when revisions are pending — accept/reject before text-count verification.

## Find & Replace (Word) — the danger zone

- `rng.Find.Execute(old)` errors with **"字符串参量过长" (string parameter too long)** on search strings over ~255 chars. Match a short unique prefix instead.
- **Safe paragraph-level replacement (also safe when new contains old):**

```python
for para in d.Paragraphs:
    s = para.Range.Text
    if s.startswith(old):
        para.Range.Text = new + s[len(old):]
        break
```

- Deleting a paragraph: `para.Range.Delete()` (removes the ¶ mark too). **Do NOT cache `list(d.Paragraphs)` and mutate by index — COM ranges go stale after deletion.** Re-scan the live collection per delete, or anchor by text.
- Find loops are fine only when `new` does NOT contain `old`:
  `rng = d.Content; rng.Find.Execute(old); rng.Text = new; rng.Collapse(0); rng = d.Range(rng.Start, d.Content.End)`

## Styles & Outline (Word)

- `para.OutlineLevel`: 1–9 = heading levels, **10 = body text**.
- Builtin style constants: **-1 = wdStyleNormal (正文), -2 = wdStyleHeading1, -3 = wdStyleHeading2, -4 = wdStyleHeading3**. Setting -2 gives Heading1, not Heading2 — classic off-by-one.
- Fix body text wrongly tagged as heading: `para.Style = -1` (resets OutlineLevel to 10). Verify with `para.Style.NameLocal`.

## Fonts — 中文字体 vs 西文字体 (critical COM split)

**Rule: 生成 Word 时若未指定字体，默认中文字体 = 仿宋_GB2312，默认西文字体 = Times New Roman**（公文规范，GB/T 9704-2012）。

COM 里中英文字体是**两个独立属性**，只设一个会漏掉另一半：

```python
# 同时设置西文 + 中文，缺一不可
rng.Font.Name = 'Times New Roman'        # 西文/数字字体
rng.Font.NameFarEast = '仿宋_GB2312'     # 中文/东亚字体
```

- 只设 `Font.Name` → 中文不变（仍继承原字体或默认）；只设 `NameFarEast` → 西文/数字不变。**两个都要设。**
- **例外：页脚/页眉 Range 上不要设 `NameFarEast`**（实测报 OLE error 0x800a16d4）——页码是数字，`Font.Name` 足够（见「Page numbers」章节）。正文、表格、普通文本一律双属性。
- 对一段文字设置字体前先选中整个 Range：`para.Range.Font.Name = ...` 而不是 `para.Font`。
- 表格单元格同理：`t.Cell(r, c).Range.Font.Name = 'Times New Roman'; t.Cell(r, c).Range.Font.NameFarEast = '仿宋_GB2312'`。
- 新建文档时，用 `d.Content.Font` 设文档默认字体；或直接改 Normal 样式：`d.Styles(-1).Font.NameFarEast = '仿宋_GB2312'`。
- 常见坑：从模板继承的字体与期望不一致 → 打开后先 `d.Content.Font.Reset()`（若允许）或逐段覆盖。

## Font substitution — 缺字自动替换 (生僻字回退)

**Rule: 当指定字体中不存在的文字（缺字形），Word 自动用相似字体替换显示。例如"溇"在仿宋_GB2312 中没有该字形，默认自动替换为仿宋显示/打印。** 这是显示/渲染层行为，文档内字体名不变。

COM 里可用 `Application.SubstituteFont(UnavailableFont, SubstituteFont)` 显式设置"不可用字体 → 替换字体"映射（等价于 文件→选项→字体替换 对话框）：

```python
w.SubstituteFont('仿宋_GB2312', '仿宋')        # 缺字/缺字体时用仿宋回退（实测 Office 2024 可用）
w.SubstituteFont('楷体_GB2312', '楷体')
```

- ✅ 实测：`Application.SubstituteFont` 在 Office 2024 (Word) 下正常；参数顺序 = (不可用字体, 替换字体)。
- ❌ `Document.FontSubstitutions` 集合在 win32com 下访问常报错（`<unknown>.FontSubstitutions`）——不要依赖它，用 SubstituteFont 方法。
- 这种替换只影响显示/打印/导出 PDF 的观感，**不修改文档中保存的字体名**；导 PDF 前若担心生僻字观感，先调用 SubstituteFont 设置好回退映射。
- 生僻字检测：若需确认某字在某字体中是否有字形，可在脚本里用 PIL/fontTools 查字体 cmap，或直接接受 Word 的回退行为（公文常见做法）。

## GB/T 9704-2012 公文格式速查 (Chinese government documents)

| 元素 | 字体 | 字号/其他 |
|---|---|---|
| 标题 | 方正小标宋简体 | 2号，居中 |
| 一级标题 | 黑体 | 3号 |
| 二级标题 | 楷体_GB2312 | 3号 |
| 正文 | 仿宋_GB2312 | 3号，两端对齐，首行缩进2字符，行距28磅 |
| 表格表头 | 黑体 | 小四加粗，浅灰底纹，跨页重复表头 |
| 表格内容 | 仿宋_GB2312 | 小四，按内容居左/居右/居中 |

每处设置字体时都按上面的 Name + NameFarEast 双属性写，避免中文回退到宋体/等线。

## Paragraph / Image / Table alignment (排版默认规则)

**Rule: 正文默认两端对齐 + 首行缩进2字符；图片和表格默认居中（无缩进）。**
**Rule 2: 确认是文章标题 / 表标题（表题）/ 图题（图题）的段落，默认居中且不缩进2字符。**

Word 对齐常量：0=左对齐, 1=居中, 2=右对齐, 3=两端对齐(Justify), 4=分散对齐。

```python
# 正文：两端对齐 + 首行缩进2字符（按字符单位缩进，公文标准）
para.Alignment = 3                                  # wdAlignParagraphJustify
para.CharacterUnitFirstLineIndent = 2               # 首行缩进2字符（不是 FirstLineIndent 磅值！）
para.LineSpacingRule = 4                            # wdLineSpaceExactly 精确行距
para.LineSpacing = 28                               # 28磅（配合 3号仿宋）

# 标题 / 表题 / 图题：居中 + 无缩进
para.Alignment = 1                                  # wdAlignParagraphCenter
para.CharacterUnitFirstLineIndent = 0               # 清字符缩进
para.FirstLineIndent = 0                            # 清磅值缩进
para.LeftIndent = 0; para.RightIndent = 0           # 清左右缩进

# 图片：所在段落居中，无缩进
shp.Range.ParagraphFormat.Alignment = 1             # wdAlignParagraphCenter
shp.Range.ParagraphFormat.CharacterUnitFirstLineIndent = 0
shp.Range.ParagraphFormat.LeftIndent = 0
shp.Range.ParagraphFormat.LineSpacingRule = 0       # 图片段默认单倍行距

# 表格：整表在页面居中（行对齐），单元格内再按内容对齐
t.Rows.Alignment = 1                                # wdRowAlignCenter（表格整体居中）
```

- **标题类段落判定**：文章大标题（文档/章节标题）、表题（"表4-1 设备清单"）、图题（"图1 流程图"）→ 居中 + 清缩进；正文段落 → 两端对齐 + 缩进2字符。**批量排版时按此规则逐段分类处理。**
- **首行缩进必须用 `CharacterUnitFirstLineIndent = 2`（按字符），不要用 `FirstLineIndent`（磅值）**——公文要求"2字符"，磅值会随字号漂移。
- 标题/表题/图题清缩进时**同时清字符缩进和磅值缩进**（`CharacterUnitFirstLineIndent=0` + `FirstLineIndent=0`），否则旧文档残留磅值缩进会盖住居中效果。
- 图片插入后，它所在段落默认可能带缩进或左对齐 → 显式设 `Alignment = 1` + 清缩进。
- 表格默认靠左 → `t.Rows.Alignment = 1` 整体居中；表格内单元格对齐用 `t.Cell(r,c).VerticalAlignment`（1=上, 2=中, 3=下）和 `Range.ParagraphFormat.Alignment`（0/1/2/3）。

### Table cell defaults — 表格内文字格式 (默认规则)

**Rule: 表格里文字默认上下居中 + 左右居中 + 无缩进 + 单倍行距；表头加粗 + 重复标题行（跨页重复）。** 全部实测 Office 2024 可用：

```python
# 遍历所有单元格：上下居中 + 左右居中 + 无缩进 + 单倍行距
for i in range(1, t.Rows.Count + 1):
    for j in range(1, t.Columns.Count + 1):
        c = t.Cell(i, j)
        c.Range.ParagraphFormat.Alignment = 1              # 左右居中
        c.VerticalAlignment = 1                            # 上下居中 (wdCellAlignVerticalCenter)
        c.Range.ParagraphFormat.CharacterUnitFirstLineIndent = 0
        c.Range.ParagraphFormat.FirstLineIndent = 0        # 无缩进
        c.Range.ParagraphFormat.LineSpacingRule = 0        # 单倍行距 (wdLineSpaceSingle)

# 表头：加粗 + 重复标题行（跨页自动重复表头）
t.Rows(1).Range.Font.Bold = True
t.Rows(1).HeadingFormat = True                            # 跨页重复标题行（实测返回 -1）
```

- 垂直对齐常量：**1 = 居中**（wdCellAlignVerticalCenter），0=上，3=下——注意居中不是 2！
- `HeadingFormat = True` 使表头行在表格跨页时自动重复（等于 表格工具→布局→重复标题行）。
- 若表头不止一行，重复标题行要设 `t.Rows(1)` 和 `t.Rows(2)` 一起 `HeadingFormat = True`。
- 若个别列需要左对齐（如长文本描述列），再单独覆盖 `t.Cell(r,c).Range.ParagraphFormat.Alignment = 0`。

## References — GB/T 7714-2015 参考文献格式 (顺序编码制)

**Rule: 公文/报告参考文献按 GB/T 7714-2015 顺序编码制著录，正文引用上标 [1][2] 标注，文末按出现顺序编号。**

### 文献类型标识

| 标识 | 类型 | 标识 | 类型 |
|---|---|---|---|
| [J] | 期刊文章 | [M] | 专著/图书 |
| [D] | 学位论文 | [C] | 会议论文 |
| [R] | 报告 | [S] | 标准 |
| [P] | 专利 | [EB/OL] | 电子资源/网页 |
| [N] | 报纸文章 | [Z] | 其他未说明类型 |

### 各类型著录格式（顺序编码制）

```
期刊文章：[序号] 主要责任者. 文献题名[J]. 刊名, 出版年, 卷(期): 起止页码.
    例：[1] 张三, 李四, 王五, 等. 地表水水质监测技术研究[J]. 环境监测管理与技术, 2021, 33(2): 12-16.
专著：  [序号] 主要责任者. 书名[M]. 版本项(第2版等，第1版可省略). 出版地: 出版者, 出版年: 引文页码.
    例：[2] 中华人民共和国生态环境部. 环境监测分析方法标准制修订技术导则[M]. 北京: 中国环境出版社, 2020: 45-50.
学位论文：[序号] 主要责任者. 题名[D]. 保存地点: 保存单位, 年份.
    例：[3] 李四. 沁河流域水环境质量评价研究[D]. 太原: 山西大学, 2022.
会议论文：[序号] 主要责任者. 题名[C]//论文集名. 出版地: 出版者, 出版年: 页码.
标准：  [序号] 标准起草单位. 标准名称: 标准编号[S]. 出版地: 出版者, 出版年.
    例：[4] 国家市场监督管理总局. 地表水环境质量标准: GB 3838-2002[S]. 北京: 中国环境科学出版社, 2002.
报告：  [序号] 主要责任者. 题名[R]. 出版地: 出版者, 出版年.
网页：  [序号] 主要责任者. 题名[EB/OL]. (更新日期)[引用日期]. 获取路径.
    例：[5] 山西省生态环境厅. 山西省水污染防治工作方案[EB/OL]. (2023-05-10)[2026-08-12]. https://xxx.gov.cn/xxx.html.
```

### 规则要点

- **作者 3 人及以内全部列出，超过 3 人列出前 3 人加"，等"（中文）或", et al"（英文）**。
- **正文引用标注为上标方括号** `[1]`，按出现顺序编号；文末参考文献列表与正文编号一一对应。
- 多个文献合并标注：`[1-3]` 或 `[1,5]`。
- 标点一律半角；题名后的文献类型标识 [J]/[M] 等用方括号。
- 版本项：第 2 版写"2版"，第 1 版省略不写。
- 电子资源必须给出（更新日期）和 [引用日期]。
- 标准类先写标准名称再写标准号：`地表水环境质量标准: GB 3838-2002[S]`。

### COM 设置正文上标引用

```python
# 正文中 [1] 设为上标（先清该段上下标再设）
rng = d.Range(start, end)                 # 定位到 "[1]" 所在范围
rng.Font.Superscript = False; rng.Font.Subscript = False   # 先清
rng.Font.Superscript = True                                # 再设上标
```

## Superscript / Subscript — 上下角标语义判定 (单位/化学式/离子)

**Rule: 先判断文字意思，再决定上下标——不能机械把所有数字变角标。** 例如 "m3" 确定是立方米后 3 才作上标；"O2/CO2/NH3" 是化学式数字作下标；"Ca2+" 是离子 2+ 作上标。普通数字（2.0mg/L 中的 2.0、年份、编号）一律不动。

### 常见语义→角标对照表

| 原文 | 语义 | 处理 |
|---|---|---|
| m3 / m2 | 立方米/平方米（单位） | 3、2 上标 → m³ m² |
| hm2 / km2 | 公顷/平方千米 | 2 上标 → hm² km² |
| cm2 / mm2 | 平方厘米/平方毫米 | 2 上标 |
| O2 / CO2 / SO2 / NO2 / H2O / H2S | 化学式（分子） | 数字下标 → O₂ CO₂ SO₂ NO₂ H₂O H₂S |
| NH3 / NH3-N | 氨/氨氮（化学式） | 3 下标 → NH₃ NH₃-N（-N 不上标不下标） |
| CH4 / N2O / H2O2 | 甲烷/一氧化二氮/过氧化氢 | 数字下标 |
| Ca2+ / Mg2+ / Na+ / Fe3+ / Cl- / SO42- | 离子（带电荷） | 数字和正负号上标 → Ca²⁺ Mg²⁺ Na⁺ Fe³⁺ Cl⁻ SO₄²⁻ |
| 5mg/L / 2.0mg/L / 0.5μg/L | 浓度数值 | **不动**（数字是数值不是角标） |
| 2023年 / 第2期 / 编号123 | 年份/编号 | **不动** |

### 判定规则（先语义后角标）

1. **单位**：`字母+数字` 且数字∈{2,3} 且整体是面积/体积单位词（m² m³ hm² km² cm² mm²）→ 数字上标。
2. **化学式**：元素符号组合（含大写字母开头）+ 紧跟数字 → 数字下标（O₂ CO₂ NH₃ H₂O CH₄ N₂O）。
3. **离子**：元素/原子团 + 数字+正负号（2+ 3+ - 2-）→ **数字和正负号都上标**（Ca²⁺ SO₄²⁻）；正负号前无数字（Na+ Cl-）→ 仅正负号上标。
4. **氨氮 NH3-N**：NH₃ 部分数字下标，-N 正常（不上标不下标）。
5. **数值/年份/编号**：上下文是数字、浓度、金额、日期、序号 → 一律不动。
6. **歧义时**：如 "m3" 可能指立方米也可能指"3号监测点 m3"——按上下文判断，无法确定就不动。

### COM 实现（先清后设，实测 Office 2024 可用）

```python
def apply_script(doc, needle, rel_start, rel_end, kind):
    """在 doc 中查找 needle，对其 [rel_start, rel_end) 相对偏移的字符设置上下标。
    kind: 'sup' 上标 / 'sub' 下标。先清整段再设目标字符。"""
    f = doc.Content.Find; f.ClearFormatting()
    if f.Execute(needle):
        base = f.Parent.Start
        rr = doc.Range(base + rel_start, base + rel_end)
        rr.Font.Superscript = False; rr.Font.Subscript = False   # 先清
        rr.Font.Superscript = (kind == 'sup')
        rr.Font.Subscript = (kind == 'sub')
        return True
    return False

# 应用示例（每个模式调用一次）：
apply_script(d, 'hm2', 2, 3, 'sup')       # hm² 公顷：2 上标
apply_script(d, 'm3/d', 1, 2, 'sup')      # m³/d 立方米：3 上标
apply_script(d, 'NH3-N', 2, 3, 'sub')     # NH₃-N 氨氮：3 下标
apply_script(d, 'Ca2+', 2, 4, 'sup')      # Ca²⁺ 钙离子：2+ 上标
apply_script(d, 'CO2', 2, 3, 'sub')       # CO₂ 二氧化碳：2 下标
```

- **必须先清再设**：`rr.Font.Superscript = False; rr.Font.Subscript = False` 再设目标——否则残留角标格式会叠加。
- 实测返回值：`Superscript=True` 在 COM 返回 `-1`（True），`Subscript=True` 同理；验证时用 `int(rr.Font.Superscript)` 判断。
- 批量处理时按"模式列表"循环（每个模式是一个 `(needle, rel_start, rel_end, kind)` 元组），先处理长模式（如 NH3-N、SO42-）再处理短模式（如 CO2、m3），避免子串误匹配。

## Tables (Word)

- Access any cell flat: `t.Range.Cells(i).Range.Text = 'x\r'` — works on merged tables where `t.Rows(r).Cells` raises.
- **Add a column:** `t.Columns.Add(t.Columns(1))` then fill `t.Cell(r, 1).Range.Text` per row (header first). Verify `t.Columns.Count` went up.
- **Delete a trailing column ONLY if every cell is empty:** inspect all `t.Columns(c).Cells`; headers/merged cells count as content.
- **Append a table at end:** collapse range to end, `d.Tables.Add(r, rows, cols)`, `t.Borders.Enable = True`, fill `t.Cell(i,j).Range.Text`.
- Insert a heading before a position: `r = d.Range(pos, pos); r.InsertBefore(text + '\r')`, then set style on the inserted range.

### Table auto-fit — 表格超出页边距的修复 (AI-generated tables)

**Rule: AI 生成的表格常超出页边距，手动修复顺序 = 表格布局 → 根据窗口自动调整 → 根据内容自动调整。COM 里对应 `AllowAutoFit` + `AutoFitBehavior()`（实测 Office 2024 可用）：**

```python
t.AllowAutoFit = True
t.AutoFitBehavior(1)          # wdAutoFitWindow — 根据窗口自动调整（首选，把表收进页边距）
# 若仍需微调：t.AutoFitBehavior(2)  wdAutoFitContent — 根据内容自动调整
# 固定列宽：t.AutoFitBehavior(0)     wdAutoFitFixed
```

- 实测效果：`AutoFitBehavior(1)` 后 `t.PreferredWidthType` 变为 1（百分比，整表占页面宽度）；`AutoFitBehavior(2)` 后变为 2 且 `PreferredWidth=100`。
- **新表生成后默认就应调用一次** `AutoFitBehavior(1)`，不要等用户发现超边距。
- 若个别列仍超宽，可再设 `t.Columns(j).Width = pt` 或 `t.Cell(r,c).Width` 微调。
- 表题（表上方"表4-1 …"段落）居中，与表格整体宽度无关。

## Page numbers (页码)

**Rule: AI 生成 Word 默认页面底端（页脚）加入页码，小五号（9pt）、Times New Roman。** 实测 Office 2024 可用：

```python
sec = d.Sections(1)
ftr = sec.Footers(1)
fr = ftr.Range; fr.Text = ''
fr.Fields.Add(fr, -1, 'PAGE \\* MERGEFORMAT')     # PAGE 域（自动页码）
fr.Font.Name = 'Times New Roman'                   # 只设 Name 即可（页码是数字）
fr.Font.Size = 9                                   # 小五号 = 9pt
fr.ParagraphFormat.Alignment = 1                   # 居中（也可 0 左 / 2 右）
```

- 注意：页脚 Range 上**不要设 `NameFarEast`**（实测报 OLE error 0x800a16d4）——页码为数字，`Font.Name` 足够。
- 页脚内容会含 `\r` 结尾（`'1\r'`），读回验证时注意 strip。
- 公文多节时每节分别设；正文节若要求"第 X 页 共 Y 页"，用两个域：`PAGE` + `NUMPAGES`（`'PAGE \\* MERGEFORMAT / NUMPAGES \\* MERGEFORMAT'` 形式或分开 Add）。
- 导出 PDF 前先 `d.Fields.Update()` + `d.Repaginate()` 让页码正确。

## COM Object Discovery (探索对象模型)

When you don't know the exact property/method names of a COM object (the "复杂结构" pain), enumerate them at runtime instead of guessing — this is how the active Office MCP projects were built:

```python
import win32com.client
w = win32com.client.DispatchEx('Word.Application')   # 只读探索，独立进程
props = getattr(x, '_prop_map_get_', {})
print('PROPERTIES:', sorted(props.keys()))
print('METHODS:', sorted([m for m in dir(x) if not m.startswith('_') and 'method' in str(type(getattr(x, m, None))).lower()]))
```

For Word, walk the object model programmatically: `w.ActiveDocument.Sections.Count`, `d.Sections(1).Headers.Count`, `d.Tables.Count`, `d.Shapes.Count`, `d.StoryRanges.Count` — inspect counts before assuming structure.

## Dialog-box hang prevention (防弹窗卡死) — 后台操作被前台弹窗卡住的根治方案

**Symptom: AI 后台操作 Word 卡住，前台弹出"是否保存"、"是否打开"、"文件正在使用"等对话框。** COM 调用是同步阻塞的——对话框一出现，Python 脚本就永远挂起直到有人手动点掉（或超时被杀）。

**Root cause: `DisplayAlerts=0` 只抑制部分警告（保存提示等），管不了这些弹窗：文件正在使用/只读打开、受保护视图、宏安全、转换确认。** 需要组合拳（全部实测 Office 2024 可用）：

### 1. 启动时的防弹窗设置（每次 COM 脚本必做）

```python
w = win32com.client.DispatchEx('Word.Application')  # 独立进程，不碰用户已打开的 Word
w.Visible = False
w.DisplayAlerts = 0                                  # 抑制保存等警告
w.AutomationSecurity = 3                             # msoAutomationSecurityForceDisable 禁宏弹窗
```

### 2. 打开文件前检测占用（避免"文件正在使用"弹窗）

```python
def is_locked(path):
    try:
        fh = open(path, 'r+b')
        try:
            msvcrt.locking(fh.fileno(), msvcrt.LK_NBLCK, 1)
            locked = False
            try: msvcrt.locking(fh.fileno(), msvcrt.LK_UNLCK, 1)
            except: pass
        except OSError:
            locked = True
        fh.close()
        return locked
    except (PermissionError, OSError):
        return True

if is_locked(path):
    raise RuntimeError(f'文件被占用，拒绝打开：{path}')   # 报错而不是让 Word 弹窗卡住
```

### 3. Open / Close 全参数（堵住对话框入口）

```python
d = w.Documents.Open(path,
    ConfirmConversions=False,     # 不弹格式转换确认
    ReadOnly=False,
    AddToRecentFiles=False,
    Revert=False)                 # 不弹"已打开，是否恢复"

# 保存/关闭时显式传参，杜绝"是否保存更改"弹窗
d.Save()
d.Close(SaveChanges=False)        # 或 SaveChanges=wdSaveChanges(-1)
w.Quit()
```

### 4. 受保护视图检查（打开网络/他人文件时）

```python
if w.ProtectedViewWindows.Count > 0:
    # 受保护视图下 COM 操作受限，需先退出保护视图
    for i in range(1, w.ProtectedViewWindows.Count + 1):
        w.ProtectedViewWindows(i).Edit()     # 或 .Activate() 后操作
```

### 5. 卡死兜底：超时 + 强杀（最后防线）

即使做了上面所有防护，仍可能遇到未预料的弹窗。用后台线程 + 超时 + taskkill 兜底：

```python
import threading, subprocess, time

result = {}
def worker():
    try:
        pythoncom.CoInitialize()
        w = win32com.client.DispatchEx('Word.Application')
        w.Visible = False; w.DisplayAlerts = 0; w.AutomationSecurity = 3
        # ... 你的 COM 操作 ...
        w.Quit(); pythoncom.CoUninitialize()
        result['ok'] = True
    except Exception as e:
        result['err'] = e

t = threading.Thread(target=worker); t.start()
t.join(timeout=120)                    # 120s 无响应视为卡死
if t.is_alive():
    subprocess.run('taskkill /F /IM WINWORD.EXE', shell=True)   # 强杀，别让弹窗永远挂着
    raise RuntimeError('Word 操作超时，已强杀进程（可能遇到未抑制的弹窗）')
if 'err' in result:
    raise result['err']
```

**规则总结：**
1. `DispatchEx`（独立进程）+ `DisplayAlerts=0` + `AutomationSecurity=3` 是基础三件套，**每次脚本都写**。
2. 打开前 `is_locked()` 检测，占用就报错，绝不裸 Open。
3. `Open`/`Close` 永远显式传全参数。
4. 复杂脚本套线程超时 + taskkill 兜底，超时即杀不留僵尸进程。
5. 记得备用方案：事后 `cmd //c "taskkill /F /IM WINWORD.EXE"` 清理残留。

## Dispatch vs DispatchEx vs EnsureDispatch (choose the right one)

The single most common source of "why is my script controlling someone else's Word?" confusion:

| Call | Behavior | Use when |
|---|---|---|
| `win32com.client.Dispatch('Word.Application')` | Reuses an **already-running instance** (Running Object Table); starts one only if none exists | Interactive single-session work; safe when no other instance is open |
| `win32com.client.DispatchEx('Word.Application')` | **Always spawns a fresh independent process** | Server/background/batch work; concurrent scripts; guaranteed clean shutdown (`Quit()` really exits, process doesn't linger) |
| `gencache.EnsureDispatch('Word.Application')` | Early-bound static proxy; builds/uses `%TEMP%\gen_py` cache; typed properties, faster, stable constants | Complex scripts where you rely on typed members / want discovery via `_prop_map_get_` |

**Pitfalls (community-verified):**
- `gencache` cache corruption: `EnsureDispatch` raises `AttributeError: module 'win32com.gen_py...' has no attribute 'CLSIDToClassMap'` (pywin32 #1923). Fix: **delete `%TEMP%\gen_py`** and retry — it does NOT rebuild a corrupted cache automatically.
- 64-bit Office + `EnsureDispatch` can fail with `TypeError: This COM object can not automate the makepy process` (pywin32 #1568) when the type library doesn't register cleanly. Run `makepy` manually or clear the cache.
- `DispatchEx` does not see documents opened in a pre-existing interactive Office instance; `Dispatch` may attach to one unexpectedly. Pick deliberately per script.
- Word may linger after `Quit()` when launched via `Dispatch` — prefer `DispatchEx` for guaranteed cleanup, or `taskkill /F /IM WINWORD.EXE` after.

## Multi-threaded COM (threading)

win32com + threads is flaky; community pattern (博客园-verified):

- One `Application` object per process; open/close **separate Documents per thread**.
- In each worker thread, call `pythoncom.CoInitialize()` before touching COM and `pythoncom.CoUninitialize()` after closing the doc. Without this, random COM errors in threads are near-certain.

```python
import pythoncom, win32com.client

def worker(path):
    pythoncom.CoInitialize()
    try:
        w = win32com.client.DispatchEx('Word.Application')   # fresh instance per worker
        w.Visible = False; w.DisplayAlerts = 0
        d = w.Documents.Open(path, ReadOnly=True)
        # ... work ...
        d.Close(False); w.Quit()
    finally:
        pythoncom.CoUninitialize()
```

## Image-Container Paragraphs (Word) — data-loss incident

In converted/OCR'd docs, paragraphs whose text is just `/` are often **image containers** — the InlineShape is anchored inside them. `para.Range.Delete()` deletes paragraph + image.

**Rules:** before deleting any paragraph, check `len(para.Range.InlineShapes) == 0`. Snapshot `d.InlineShapes.Count` before edits and verify it never dropped afterward. Insert images via Range, not `Selection.TypeParagraph()/TypeText()` (leaves stray `/` chars).

```python
for para in list(d.Paragraphs):
    s = para.Range.Text.replace('\r','').replace('\x07','').strip()
    if s == '/' and len(para.Range.InlineShapes) == 0:
        para.Range.Delete()
```

## Document Health Check (run after every significant edit)

```python
d = w.Documents.Open(path, ReadOnly=True, AddToRecentFiles=False)
t = d.Content.Text
print(len(t), d.Paragraphs.Count, d.Tables.Count)      # compare vs pre-edit baseline
for k in ['<known-pattern>']:
    print(k, t.count(k))                                # expect 1, not 2282
```

## Budget / Data-Cascade Edits

When a quantity changes (e.g. 6 sets → 4 sets), recompute and update EVERY linked number: subtotals, grand total, funding split, annual split, per-appendix tables, remarks text, %-based figures. Verify at the end: `a+b==total`, funding rows sum to total, annual rows sum to total, line items sum to subtotal.

## Verification Checklist

1. Reopen read-only; char/paragraph/table counts sane (not 3× inflated).
2. `t.count('关键旧短语') == 0` and `t.count('关键新短语') >= 1` for each edit.
3. Appendix headings appear exactly once; TOC entries match body titles.
4. Budget sums reconcile.
5. No leftover heading-styled body lines (`OutlineLevel != 10` review).
6. No process left running — `tasklist | findstr /i "WINWORD"` clean.

## References

- `references/win32com-cheatsheet.md` — full verified snippets: safe bulk replace, style constants table, table column add/delete, merged-cell reads, image insert/resize, fonts, PDF export, page numbers, references, superscript/subscript.
- `references/pitfalls.md` — incident transcripts: infinite TOC duplication, stale-process corruption, long-string Find error, stale paragraph lists, image-container paragraph loss, style off-by-one.
