# MS Office COM — Pitfall Incident Transcripts

Real incidents from editing Chinese government proposal documents (晋城市水生态环境监测监管能力建设项目实施方案, multi-round edits, 2026-08). Each cost real time; each is now a Golden Rule.

## Incident 1: Infinite TOC duplication (document corruption)

**Symptom:** char count jumped 72,518 → 211,601; paragraph count 1,864 → 9,234; a 3-line TOC block repeated 2,282 times.

**Root cause:** a `replace_all` Find loop where the replacement contained the search string:

```python
def replace_all(doc, old, new):        # BAD when new contains old
    rng = doc.Content; n = 0
    while True:
        if not rng.Find.Execute(old): break
        rng.Text = new                 # inserting new re-introduces old text after the cursor
        n += 1
        rng.Collapse(0)
        rng = doc.Range(rng.Start, doc.Content.End)
```

Called as `replace_all(d, '附表3\t124', '附表3\t124\r附表4\t...')` — after each replacement the search text `附表3\t124` sits right after the collapsed cursor, so the loop re-finds it forever.

**Detection:** `t.count('附表4')` returned 2,282 instead of 1. A sane doc's appendix heading appears once.

**Fix:** restore from pre-round backup, then do TOC edits at paragraph level:

```python
for para in d.Paragraphs:
    s = para.Range.Text
    if s.startswith('附表3\t124'):
        para.Range.Text = '附表3\t124\r附表4\t晋城市地表水考核断面基本信息表\r附表5\t...'
        break
```

## Incident 2: Stale Word process writes corruption to disk

After the infinite loop, the first script **timed out** and was killed — `w.Quit()` never ran. A second script's `Documents.Open` silently **reused the corrupted in-memory document** (Word keeps it open), executed a tiny successful edit, and `Save()` wrote 211K chars to disk. The file on disk was permanently corrupted.

**Rule:** after any COM timeout/crash, kill Word before reopening:
- `cmd //c "taskkill /F /IM WINWORD.EXE"` (from git-bash; `taskkill //F //IM` fails under MSYS arg mangling)
- then restore from backup and redo the edits.

## Incident 3: Long-string Find error

`rng.Find.Execute(old)` raised:
`pywintypes.com_error: (-2147352567, '发生意外。', (0, 'Microsoft Word', '字符串参量过长。', 'wdmain11.chm', 25334, -2146822434), None)`

Trigger: search string > ~255 chars (a full "数据来源：…" parenthetical). Fix: match on a short unique prefix and replace the whole paragraph, or split into several shorter Find calls.

## Incident 4: Style constant off-by-one

`para.Style = -2` produced **标题1 (Heading1)**, not 标题2. Word's WdBuiltinStyle:
- `-1` wdStyleNormal (正文)
- `-2` wdStyleHeading1
- `-3` wdStyleHeading2
- `-4` wdStyleHeading3

Also: `para.Style = -1` resets OutlineLevel to 10 (body). Verify with `para.Style.NameLocal` and `para.OutlineLevel` after setting.

## Incident 5: Cached Paragraph list goes stale after deletion

`paras = list(d.Paragraphs)` then deleting `paras[i+1].Range` shifted every later index; the next cached index deleted the WRONG paragraph (附表1-3标题 were lost entirely). Rule: never hold a Paragraph list across mutations — re-scan `d.Paragraphs` fresh for each delete/insert, or anchor by text/position, not by cached index.

## Incident 6: Image-container paragraphs — data loss

Symptom: after "cleaning" stray `/` paragraphs, `d.InlineShapes.Count` dropped 22 → 2. Root cause: in converted/OCR'd Word docs, placeholder paragraphs whose text is just `/` are **image containers** — the InlineShape is anchored inside that paragraph. `para.Range.Delete()` deleted paragraph + image.

Rules:
- Before deleting any paragraph, check `len(para.Range.InlineShapes)`. Only delete if 0.
- Snapshot `d.InlineShapes.Count` before edits; verify it after (should never decrease from an edit that isn't an intentional image removal).
- When a clean-up sweep looks suspicious, roll back from the pre-edit backup and re-apply edits with the guard, rather than continuing to mutate a damaged doc.

```python
# safe sweep — deletes only text-only '/' paragraphs
for para in list(d.Paragraphs):
    s = para.Range.Text.replace('\r','').replace('\x07','').strip()
    if s == '/' and len(para.Range.InlineShapes) == 0:
        para.Range.Delete()
```

## Incident 7: Budget cascade (6→4 sets example)

Original 7028万 total → after -66万: 6962万. Every linked number updated: sub-project 5253→5187, funding 4920/1054/1054→4874/1044/1044, annual 2811/2460/1757→2785/2437/1740, appendix subtotal 1323→1257, lab package 1018→952. Reconciliation checks: `5187+1775==6962`, `4874+1044+1044==6962`, `2785+2437+1740==6962`, line items sum to subtotal.

**Lesson:** when a quantity changes, recompute and update EVERY linked number and verify sums at the end — never trust a single find/replace to catch the cascade.

## Incident 8: gen_py cache corruption — EnsureDispatch fails (community, pywin32 #1923)

Symptom: `gencache.EnsureDispatch('Excel.Application')` raises
`AttributeError: module 'win32com.gen_py.00020813-...' has no attribute 'CLSIDToClassMap'`
(or `TypeError: This COM object can not automate the makepy process` on 64-bit Office, #1568).

Fix: **delete `%TEMP%\gen_py`** and retry — the cache is NOT rebuilt automatically when corrupted.
Related: PyInstaller-frozen apps break win32com the same way (pyinstaller #6257/#7898) — pre-generate the cache or use `dynamic.Dispatch`.

## Incident 9: Dispatch attached to the WRONG (already-running) instance

`Dispatch('Word.Application')` reuses whatever instance is already running (Running Object Table).
If a user has Word open interactively, your script silently drives THAT instance — documents opened in
the script appear in their window, and `Quit()` closes their Word. Community-verified pattern:
use `DispatchEx` for independent/batch work (fresh process, clean shutdown), keep `Dispatch` only for
single-session interactive work when you KNOW no other instance is running.

## Incident 10: Threads + win32com = random COM errors (community, 博客园)

Per-thread COM without initialization fails sporadically. Pattern: one Application per process,
separate Documents per thread, and in each thread:
```python
pythoncom.CoInitialize()
try: ... work ...
finally: pythoncom.CoUninitialize()
```

## Incident 11: Excel merged range reads return only top-left (community, CSDN)

`ws.Range('B2:D4').Value` on a merged range returns only the top-left cell's value.
Walk the span with `rng.Cells(r, c)` / `rng.Offset(r, c)`; when writing, set the top-left cell only.

## User conventions for Chinese government proposal documents (公文)

- **No source citations.** Reports written in a bureau's name must NOT carry "来源于官方网站/数据来源/网络检索" annotations or URLs — state facts directly. Remove `（…数据来源：…）` parentheticals entirely.
- **Numbering must be gapless**: no `(1)` followed by `(3)`, no `①` then `③`. Renumber paragraph prefixes; verify by scanning paragraphs for starts-with patterns.
- **Appendix/附表 titles need names** (e.g. `附表1 晋城市水生态环境监测监管能力建设项目清单`), consistent across body AND TOC. Multi-device parameter tables need a first-column 设备名称 (device name) mapping each row to its equipment.
- **Content numbers must cross-check**: investment subtotals = grand total, funding shares sum to 100%, year splits sum to total, equipment counts match the checklist tables. Report remaining inconsistencies rather than silently choosing.
