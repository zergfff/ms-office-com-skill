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

## Excel — open / read / write

```python
import win32com.client
x = win32com.client.Dispatch('Excel.Application'); x.Visible = False; x.DisplayAlerts = 0
wb = x.Workbooks.Open(r'C:\path\book.xlsx')
ws = wb.Worksheets(1)                      # or wb.Worksheets('预算表')

val = ws.Cells(3, 5).Value                 # read a cell
ws.Cells(4, 5).Value = 1234.5              # write a cell
ws.Range('B2:D10').Value = [ [1,2,3], [4,5,6], ... ]   # batch write 2D list
ws.Cells(5, 1).Formula = '=SUM(B2:B10)'    # formula

# merge / unmerge
ws.Range('A1:C1').Merge(); ws.Range('A1:C1').UnMerge()

# formatting
ws.Range('A1').Font.Bold = True
ws.Range('A1').Font.Size = 12
ws.Cells(1, 1).NumberFormat = '#,##0.00'
ws.Columns(1).ColumnWidth = 22
ws.Range('A1:D1').Interior.Color = 0xFFFF00   # BGR! yellow

# row/col insert/delete
ws.Rows(5).Insert(); ws.Rows(5).Delete()
ws.Columns(3).Insert()

# new workbook + rename sheet
wb2 = x.Workbooks.Add(); ws2 = wb2.Worksheets(1); ws2.Name = '新表'

wb.Save(); wb.Close(False); x.Quit()
```

SaveAs format constants: xlsx=51, xlsm=52, csv=6, xls=56. Use `wb.SaveAs(path, FileFormat=51)`.

## Excel — value gotchas

- Empty cell reads as `None`, not ''.
- Dates come back as `datetime.datetime` — format with `.strftime()`.
- Batch write a column as a 1-column list of lists: `[[v] for v in values]`.
- Setting `.Value` on a merged range raises — set the top-left cell only.
- COM quirk: Excel sometimes keeps a hidden instance alive — always `x.Quit()` in finally, and `taskkill /F /IM EXCEL.EXE` if it lingers.

## PowerPoint — open / edit slides / shapes

```python
import win32com.client
p = win32com.client.Dispatch('PowerPoint.Application')
p.Visible = True                            # PPT needs Visible=True for most ops
prs = p.Presentations.Open(r'C:\path\deck.pptx')

# read all text
for i, slide in enumerate(prs.Slides, 1):
    for shp in slide.Shapes:
        if shp.HasTextFrame:
            t = shp.TextFrame.TextRange.Text
            if t.strip(): print(i, t)

# add blank slide + textbox
s = prs.Slides.Add(prs.Slides.Count + 1, 12)     # 12 = ppLayoutBlank
tb = s.Shapes.AddTextbox(1, 50, 50, 500, 60)     # 1 = msoTextOrientationHorizontal
tr = tb.TextFrame.TextRange
tr.Text = '标题文字'
tr.Font.Size = 32; tr.Font.Bold = True

# add table
st = s.Shapes.AddTable(3, 4, 50, 150, 600, 120)
for r in range(1, 4):
    for c in range(1, 5):
        st.Table.Cell(r, c).Shape.TextFrame.TextRange.Text = f'{r},{c}'

# add picture
s.Shapes.AddPicture(r'C:\path\img.png', False, True, 50, 300, 400, 250)

prs.Save(); prs.Close(); p.Quit()
```

Save format: pptx=24, ppt=1. `prs.SaveAs(path, 24)`.

## Word — export to PDF / refresh TOC / headers-footers

```python
# export PDF (公文交付)
d.ExportAsFixedFormat(OutputFileName=r'C:\path\out.pdf', ExportFormat=17)  # 17=wdExportFormatPDF

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

## Excel — advanced (PivotTable / Chart / conditional formatting / validation / freeze / filter)

```python
x.ActiveWindow.SplitRow = 1; x.ActiveWindow.FreezePanes = True      # 冻结首行
ws.Range('A1:D50').AutoFilter(1, '晋城')                             # 自动筛选
wb.Names.Add('设备清单', ws.Range('A2:A20'))                         # 命名区域

rng = ws.Range('C2:C50')
rng.FormatConditions.Add(2, 3, '>100')                              # 2=xlCellValue, 3=xlGreater
rng.FormatConditions(1).Interior.Color = 0x00FF00                   # BGR 浅绿

v = ws.Range('B2:B10').Validation; v.Delete()
v.Add(1, 1, 1, '1', '100')                                          # 整数 1..100 校验

cht = ws.Shapes.AddChart().Chart; cht.SetSourceData(ws.Range('A1:B10')); cht.ChartType = 51

pc = x.ActiveWorkbook.PivotCaches().Create(1, ws.Range('A1:D50'))
pt = pc.CreatePivotTable('Pivot!R3C1', '汇总')
pt.PivotFields('产品').Orientation = 1; pt.PivotFields('金额').Orientation = 4

# sheet → PDF (视觉验证/交付)
ws.ExportAsFixedFormat(0, r'C:\path\sheet.pdf')
```

## PowerPoint — export slides to images / notes / master / transitions

```python
for i, slide in enumerate(prs.Slides, 1):
    slide.Export(rf'C:\path\slide_{i}.png', 'PNG', 1280, 720)   # 视觉验证

slide.NotesPage.Shapes(2).TextFrame.TextRange.Text = '讲解要点...'
prs.ApplyTemplate(r'C:\path\template.potx')
prs.SlideMaster.Shapes.Title.TextFrame.TextRange.Font.Size = 36
slide.SlideShowTransition.EntryEffect = 33                        # ppEffectFade
slide.SlideShowTransition.Duration = 1.0
```

## COM Object Discovery (探索对象模型)

```python
import win32com.client
x = win32com.client.Dispatch('Excel.Application')
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

```python
# 正文：两端对齐(3) + 首行缩进2字符（按字符单位！）
para.Alignment = 3                                # wdAlignParagraphJustify
para.CharacterUnitFirstLineIndent = 2             # 首行缩进2字符
para.LineSpacingRule = 4; para.LineSpacing = 28   # 精确行距28磅

# 图片：段落居中、无缩进
shp.Range.ParagraphFormat.Alignment = 1           # wdAlignParagraphCenter
shp.Range.ParagraphFormat.CharacterUnitFirstLineIndent = 0
shp.Range.ParagraphFormat.LeftIndent = 0

# 表格：整表居中
t.Rows.Alignment = 1                              # wdRowAlignCenter
# 单元格垂直对齐：t.Cell(r,c).VerticalAlignment  (1=上 2=中 3=下)
```

对齐常量：0=左, 1=中, 2=右, 3=两端, 4=分散。**缩进用 `CharacterUnitFirstLineIndent`（字符单位），不要用 `FirstLineIndent`（磅值，随字号漂移）。**

## Word — 缺字自动替换 (生僻字回退)

**规则：字体中不存在的文字（缺字形）自动用相似字体替换显示。例："溇"在仿宋_GB2312 无字形 → 默认回退为仿宋显示/打印。** 只影响显示/PDF 观感，不修改文档内字体名。

```python
# 显式设置回退映射（等价 文件→选项→字体替换；实测 Office 2024 可用）
w.SubstituteFont('仿宋_GB2312', '仿宋')     # 参数顺序 = (不可用字体, 替换字体)
w.SubstituteFont('楷体_GB2312', '楷体')
```
注意：`Document.FontSubstitutions` 集合在 win32com 下访问常报错，用 SubstituteFont 方法代替。生僻字检测可用 fontTools 查字体 cmap。

## Common cleanup (git-bash)

```bash
cmd //c "taskkill /F /IM WINWORD.EXE"
cmd //c "taskkill /F /IM EXCEL.EXE"
cmd //c "taskkill /F /IM POWERPNT.EXE"
tasklist | findstr /i "WINWORD EXCEL POWERPNT"    # verify none remain
```
