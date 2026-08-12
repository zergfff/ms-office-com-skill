# win32com Cheatsheet — verified working snippets

All patterns verified in real sessions editing Chinese government proposal documents (公文实施方案) on Windows. Paths use `C:\...` but patterns are generic.

## Word — safe bulk replace

Short patterns — Find.Execute is fine (limit ~255 chars):

```python
def replace_all(doc, old, new):
    rng = doc.Content; n = 0
    while True:
        if not rng.Find.Execute(old): break
        rng.Text = new; n += 1
        rng.Collapse(0)
        rng = doc.Range(rng.Start, doc.Content.End)   # DANGER: only safe if new does NOT contain old
    return n
```

**Long patterns / paragraph-level (recommended for anything over ~200 chars or any risk of self-match):**

```python
def del_para_containing(doc, key):
    for para in list(doc.Paragraphs):
        if key in para.Range.Text:
            para.Range.Delete()
```

```python
# rewrite paragraph prefix (e.g. renumbering （3）->（2）)
maps = {'（3）水质状况：': '（2）水质状况：', ...}
for para in d.Paragraphs:
    s = para.Range.Text
    for old, new in maps.items():
        if s.startswith(old):
            para.Range.Text = new + s[len(old):]
            break
```

## Word — style constants (WdBuiltinStyle)

| Value | Style | Chinese |
|-------|-------|---------|
| -1 | wdStyleNormal | 正文 |
| -2 | wdStyleHeading1 | 标题1 |
| -3 | wdStyleHeading2 | 标题2 |
| -4 | wdStyleHeading3 | 标题3 |

- Set: `para.Style = -1` (body), `para.Style = -3` (heading2).
- Verify: `para.OutlineLevel` — 10 = body, 1-9 = heading levels.
- `para.Style.NameLocal` returns localized name ('正文', '标题 2', ...).

## Word — table column add / delete / read

```python
t = d.Tables(idx)                      # 1-based; inventory first: t.Range.Text[:40]

# add leading column (e.g. device-name column)
t.Columns.Add(t.Columns(1))
for r in range(1, t.Rows.Count + 1):
    cell = t.Cell(r, 1)
    cell.Range.Text = names[r-1]
    cell.Range.Font.Size = 9
    cell.Range.Font.Name = '仿宋_GB2312'
    if r == 1: cell.Range.Font.Bold = True

# read every cell flat (works on merged tables, unlike Rows(r).Cells)
cells = t.Range.Cells
for i in range(1, cells.Count + 1):
    s = cells(i).Range.Text.replace('\r','').replace('\x07','').strip()

# delete trailing column ONLY if every cell is empty
vals = [c.Range.Text.replace('\r','').replace('\x07','').strip() for c in t.Columns(t.Columns.Count).Cells]
if all(v == '' for v in vals):
    t.Columns(t.Columns.Count).Delete()
```

## Word — insert heading before a position

```python
pos = d.Tables(24).Range.Start        # or a paragraph's Range.Start
r = d.Range(pos, pos)
r.InsertBefore(text + '\r')
r2 = d.Range(pos, pos + len(text) + 1)
r2.Style = -3                          # heading2
r2.Font.Bold = True
```

## Word — append a table at end

```python
r = d.Range(); r.Collapse(0)
t = d.Tables.Add(r, rows, cols)
t.Borders.Enable = True
for i in range(rows):
    for j in range(cols):
        t.Cell(i+1, j+1).Range.Text = data[i][j]
        t.Cell(i+1, j+1).Range.Font.Size = 9
```

## Word — insert an image (with caption) before an anchor paragraph

```python
WIDTH_PT = 14.5 * 28.35          # cm→pt: w.CentimetersToPoints is Application-only; multiply works everywhere
for para in d.Paragraphs:
    if para.Range.Text.strip().startswith('1.3.4 污染源与重金属分布'):
        pos = para.Range.Start
        r = d.Range(pos, pos); r.InsertBefore(caption + '\r')
        r2 = d.Range(pos, pos); r2.InsertBefore('\r')
        r3 = d.Range(pos - 1, pos - 1)
        shp = r3.InlineShapes.AddPicture(FileName=img, LinkToFile=False, SaveWithDocument=True)
        shp.LockAspectRatio = -1; shp.Width = WIDTH_PT
        break
```

Pre-resize large JPGs with PIL before inserting (7000×4000 px source will bloat the docx / may fail):

```python
from PIL import Image
im = Image.open(src)
if im.mode == 'CMYK': im = im.convert('RGB')
w, h = im.size; nh = int(h * 1600 / w)
im.resize((1600, nh), Image.LANCZOS).save(dst, 'JPEG', quality=80)
```

Avoid `Selection.TypeParagraph()/TypeText()` for image insertion — it leaves stray `/` chars. Prefer the Range-based recipe above.

## Word — full-document read (1M-token context rule)

```python
d = w.Documents.Open(p, ReadOnly=True, AddToRecentFiles=False)
text = d.Content.Text                     # includes table text (\x07 cell marks)
open(out, 'w', encoding='utf-8').write(text)
print(len(text), d.Paragraphs.Count, d.Tables.Count, d.Sections.Count)
```
Save to a cache .txt and read it with read_file for full-document analysis; grep/search works on it.

## Word — export to PDF / refresh TOC / headers-footers

```python
# export PDF (公文交付) — 默认创建书签使用标题 (CreateBookmarks=1)
d.ExportAsFixedFormat(OutputFileName=r'C:\path\out.pdf', ExportFormat=17, CreateBookmarks=1)
# 前提：标题用 Heading 样式（Style=-2/-3/-4），正文不会成为书签
# CreateBookmarks: 0=不建, 1=按标题, 2=按文档结构

# refresh TOC + fields + repaginate BEFORE exporting (页码才会对)
for toc in d.TablesOfContents: toc.Update()
d.Fields.Update()
d.Repaginate()

# per-section headers/footers (公文多节：封面无页码、正文重起)
sec = d.Sections(1)
sec.Headers(1).Range.Text = '晋城市生态环境局'          # 1=primary, 2=even, 3=first
ftr = sec.Footers(1)
ftr.Range.Text = '第 '
ftr.Range.Fields.Add(ftr.Range, -1, 'PAGE \\* MERGEFORMAT')
```

## Word — comments & track changes (专家意见/批注工作流)

```python
# read all comments
for c in d.Comments:
    print(c.Index, c.Author, c.Range.Text)

# add comment at a range
d.Comments.Add(d.Range(start, end), '意见内容')

# track changes
d.TrackRevisions = True
# ...edits...
d.AcceptAllRevisions(); d.TrackRevisions = False
```
注意：有未接受的修订时 `d.Content.Text` 会含修订标记，先 Accept 再做文本计数验证。

## COM Object Discovery (探索对象模型)

```python
import win32com.client
w = win32com.client.Dispatch('Word.Application')   # 只读探索，Dispatch 即可
props = getattr(x, '_prop_map_get_', {})
print('PROPERTIES:', sorted(props.keys()))
print('METHODS:', sorted([m for m in dir(x) if not m.startswith('_') and 'method' in str(type(getattr(x, m, None))).lower()]))
# Word: 先数再动 —— d.Sections.Count, d.Sections(1).Headers.Count, d.Tables.Count, d.Shapes.Count
```

## Word — 字体设置 (中英文分开，COM 经典坑)

**规则：未指定字体时，默认中文字体 = 仿宋_GB2312，默认西文字体 = Times New Roman**（GB/T 9704-2012 公文规范）。

```python
# 两个属性必须同时设，否则漏掉一半
rng.Font.Name = 'Times New Roman'        # 西文/数字
rng.Font.NameFarEast = '仿宋_GB2312'     # 中文/东亚

# 表格单元格同理
cell.Range.Font.Name = 'Times New Roman'
cell.Range.Font.NameFarEast = '仿宋_GB2312'

# 文档级默认：改 Normal 样式 (wdStyleNormal = -1)
d.Styles(-1).Font.NameFarEast = '仿宋_GB2312'
d.Styles(-1).Font.Name = 'Times New Roman'
```

GB/T 9704-2012 速查：标题 方正小标宋简体 2号居中；一级标题 黑体 3号；二级标题 楷体_GB2312 3号；正文 仿宋_GB2312 3号两端对齐首行缩进2字符行距28磅；表头 黑体小四加粗浅灰底纹跨页重复；表内容 仿宋_GB2312 小四。

## Word — 段落/图片/表格对齐 (排版默认规则)

**规则：正文默认两端对齐 + 首行缩进2字符；图片和表格默认居中（无缩进）。**
**规则2：确认是文章标题 / 表标题（表题）/ 图题（图题）的段落，默认居中且不缩进2字符。**

```python
# 正文：两端对齐(3) + 首行缩进2字符（按字符单位！）
para.Alignment = 3                                # wdAlignParagraphJustify
para.CharacterUnitFirstLineIndent = 2             # 首行缩进2字符
para.LineSpacingRule = 4; para.LineSpacing = 28   # 精确行距28磅

# 标题 / 表题 / 图题：居中 + 无缩进（同时清字符缩进和磅值缩进）
para.Alignment = 1                                # wdAlignParagraphCenter
para.CharacterUnitFirstLineIndent = 0
para.FirstLineIndent = 0
para.LeftIndent = 0; para.RightIndent = 0

# 图片：段落居中、无缩进、单倍行距
shp.Range.ParagraphFormat.Alignment = 1           # wdAlignParagraphCenter
shp.Range.ParagraphFormat.CharacterUnitFirstLineIndent = 0
shp.Range.ParagraphFormat.LeftIndent = 0
shp.Range.ParagraphFormat.LineSpacingRule = 0     # 单倍行距

# 表格：整表居中
t.Rows.Alignment = 1                              # wdRowAlignCenter
# 单元格垂直对齐：t.Cell(r,c).VerticalAlignment  (1=上 2=中 3=下)
```

对齐常量：0=左, 1=中, 2=右, 3=两端, 4=分散。**缩进用 `CharacterUnitFirstLineIndent`（字符单位），不要用 `FirstLineIndent`（磅值，随字号漂移）。** 标题类段落（文章大标题/表题"表4-1 …"/图题"图1 …"）判定为居中类，正文判定为两端对齐+缩进类，批量排版逐段分类处理。

## Word — 表格内文字格式 (默认规则)

**规则：表格里文字默认上下居中 + 左右居中 + 无缩进 + 单倍行距；表头加粗 + 重复标题行（跨页重复）。** 实测 Office 2024：

```python
for i in range(1, t.Rows.Count + 1):
    for j in range(1, t.Columns.Count + 1):
        c = t.Cell(i, j)
        c.Range.ParagraphFormat.Alignment = 1         # 左右居中
        c.VerticalAlignment = 1                        # 上下居中（注意：居中=1，不是2！）
        c.Range.ParagraphFormat.CharacterUnitFirstLineIndent = 0
        c.Range.ParagraphFormat.FirstLineIndent = 0    # 无缩进
        c.Range.ParagraphFormat.LineSpacingRule = 0    # 单倍行距

t.Rows(1).Range.Font.Bold = True                       # 表头加粗
t.Rows(1).HeadingFormat = True                         # 跨页重复标题行
```
垂直对齐：0=上, **1=居中**, 3=下。多行表头则 Rows(1)+Rows(2) 都设 HeadingFormat=True。长文本列可单独改 Alignment=0 左对齐。

## Word — 缺字自动替换 (生僻字回退)

**规则：字体中不存在的文字（缺字形）自动用相似字体替换显示。例："溇"在仿宋_GB2312 无字形 → 默认回退为仿宋显示/打印。** 只影响显示/PDF 观感，不修改文档内字体名。

```python
# 显式设置回退映射（等价 文件→选项→字体替换；实测 Office 2024 可用）
w.SubstituteFont('仿宋_GB2312', '仿宋')     # 参数顺序 = (不可用字体, 替换字体)
w.SubstituteFont('楷体_GB2312', '楷体')
```
注意：`Document.FontSubstitutions` 集合在 win32com 下访问常报错，用 SubstituteFont 方法代替。生僻字检测可用 fontTools 查字体 cmap。

## Word — 表格自动调整 (AI 表格超页边距修复)

**规则：AI 生成的表格常超出页边距，修复顺序 = 表格布局 → 根据窗口自动调整 → 根据内容自动调整。**

```python
t.AllowAutoFit = True
t.AutoFitBehavior(1)          # wdAutoFitWindow 根据窗口自动调整（首选）
# t.AutoFitBehavior(2)        # wdAutoFitContent 根据内容自动调整
# t.AutoFitBehavior(0)        # wdAutoFitFixed 固定列宽
```
实测 Office 2024：`AutoFitBehavior(1)` → PreferredWidthType=1（百分比），`AutoFitBehavior(2)` → PreferredWidthType=2, PreferredWidth=100。**新表生成后默认调用一次 AutoFitBehavior(1)。**

## Word — 页脚页码 (默认规则)

**规则：AI 生成 Word 默认页面底端加页码，小五号（9pt）、Times New Roman，居中。**

```python
sec = d.Sections(1)
ftr = sec.Footers(1)
fr = ftr.Range; fr.Text = ''
fr.Fields.Add(fr, -1, 'PAGE \\* MERGEFORMAT')     # PAGE 域
fr.Font.Name = 'Times New Roman'                   # 页码为数字，只设 Name（设 NameFarEast 会报 0x800a16d4）
fr.Font.Size = 9                                   # 小五号
fr.ParagraphFormat.Alignment = 1                   # 居中
```
"第 X 页 共 Y 页"：PAGE + NUMPAGES 两个域。导出 PDF 前 `d.Fields.Update()` + `d.Repaginate()`。

## Word — 防弹窗卡死 (Dialog-box hang prevention)

**症状：AI 后台操作 Word 卡住，前台弹出"是否保存/是否打开/文件正在使用"对话框。** COM 同步阻塞——弹窗出现脚本就永远挂起。`DisplayAlerts=0` 管不了文件占用、受保护视图、宏安全等弹窗。组合拳（实测 Office 2024）：

```python
# 1) 基础三件套
w = win32com.client.DispatchEx('Word.Application')   # 独立进程
w.Visible = False; w.DisplayAlerts = 0
w.AutomationSecurity = 3                             # 禁宏弹窗

# 2) 打开前检测文件占用（占用就报错，不裸 Open）
import msvcrt
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
        fh.close(); return locked
    except (PermissionError, OSError):
        return True

# 3) Open/Close 全参数
d = w.Documents.Open(path, ConfirmConversions=False, ReadOnly=False,
                     AddToRecentFiles=False, Revert=False)
d.Close(SaveChanges=False)    # 杜绝"是否保存"弹窗

# 4) 受保护视图检查
for i in range(1, w.ProtectedViewWindows.Count + 1):
    w.ProtectedViewWindows(i).Edit()

# 5) 复杂脚本套线程超时兜底：threading.Thread + join(timeout) + taskkill /F /IM WINWORD.EXE
```

## Word — 参考文献 (GB/T 7714-2015 顺序编码制)

**规则：参考文献按 GB/T 7714-2015 顺序编码制，正文引用上标 [1][2]，文末按出现顺序编号。**

类型标识：期刊[J] 专著[M] 学位论文[D] 会议[C] 报告[R] 标准[S] 专利[P] 网页[EB/OL] 报纸[N] 其他[Z]

```
期刊：[1] 张三, 李四, 王五, 等. 题名[J]. 刊名, 2021, 33(2): 12-16.
专著：[2] 作者. 书名[M]. 2版. 出版地: 出版者, 2020: 45-50.
标准：[3] 起草单位. 标准名称: GB 3838-2002[S]. 出版地: 出版者, 2002.
网页：[4] 作者. 题名[EB/OL]. (2023-05-10)[2026-08-12]. https://xxx.gov.cn/xxx.html.
```
作者 >3 人：前3人+，等 / et al。正文引用：`[1]` 上标，合并 `[1-3]`。

## Word — 上下角标语义判定 (单位/化学式/离子)

**规则：先判断文字意思再决定上下标，不机械处理。** 实测 Office 2024：

```python
def apply_script(doc, needle, rel_start, rel_end, kind):
    f = doc.Content.Find; f.ClearFormatting()
    if f.Execute(needle):
        base = f.Parent.Start
        rr = doc.Range(base + rel_start, base + rel_end)
        rr.Font.Superscript = False; rr.Font.Subscript = False   # 先清
        rr.Font.Superscript = (kind == 'sup'); rr.Font.Subscript = (kind == 'sub')
        return True
    return False

apply_script(d, 'hm2', 2, 3, 'sup')       # hm² 公顷
apply_script(d, 'm3/d', 1, 2, 'sup')      # m³/d 立方米
apply_script(d, 'NH3-N', 2, 3, 'sub')     # NH₃-N 氨氮
apply_script(d, 'Ca2+', 2, 4, 'sup')      # Ca²⁺ 钙离子
apply_script(d, 'CO2', 2, 3, 'sub')       # CO₂ 二氧化碳
```

语义对照：m3/m2/hm2/km2（面积体积单位）→ 数字上标；O2/CO2/SO2/NH3/H2O/CH4（化学式）→ 数字下标；Ca2+/Mg2+/Na+/Fe3+/SO42-（离子）→ 数字+正负号上标；NH3-N → 3 下标、-N 不动；浓度数值（2.0mg/L）、年份、编号 → 一律不动。**必须先清再设**（Superscript=False + Subscript=False），COM 返回 -1 表示 True。先处理长模式（NH3-N、SO42-）再短模式（CO2、m3）。

## Common cleanup (git-bash)

```bash
cmd //c "taskkill /F /IM WINWORD.EXE"

tasklist | findstr /i "WINWORD"    # verify none remain
```
