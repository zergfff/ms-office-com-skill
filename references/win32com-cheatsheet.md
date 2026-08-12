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

## Common cleanup (git-bash)

```bash
cmd //c "taskkill /F /IM WINWORD.EXE"
cmd //c "taskkill /F /IM EXCEL.EXE"
cmd //c "taskkill /F /IM POWERPNT.EXE"
tasklist | findstr /i "WINWORD EXCEL POWERPNT"    # verify none remain
```
