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

Symptom: `gencache.EnsureDispatch('Word.Application')` raises
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

## LLM-generation pitfalls (A类: Claude DOCX Skill 15 Critical Rules 提炼, 生成类)

These are the "LLM tries to write a docx and gets it subtly wrong" class — the official
Claude DOCX Skill's 15 Critical Rules, distilled for the COM workflow (they apply to content
LLMs generate, which then lands in Word via COM):

- **A1 `\n` 换行无效** — LLM 常写 `"第一行\n第二行"` 想换行；Word 里 `\n` 不产生段落，必须新建 Paragraph（COM：`d.Content.InsertAfter('\r' + text)` 或逐段 `Range.InsertBefore`）。
- **A2 页面尺寸默认错** — 生成文档默认 A4 是常态，但 LLM 常按 US Letter 想象。公文必须显式确认页面尺寸：`d.PageSetup.PaperSize = 7`（wdPaperA4），`d.PageSetup.Orientation = 0`（纵向）。
- **A3 TOC 无法索引标题** — 标题样式缺 `outlineLevel` → 目录抓不到。COM 对应：标题段落必须 `para.OutlineLevel` 为 1-9 且 `para.Style` 是 Heading 系列（见 Styles & Outline）。
- **A4 表格"双宽度"要求** — 表格总宽必须等于各列宽之和。COM 生成表后调用 `AutoFitBehavior(1)`（见 Table auto-fit）强制收进页面；或显式 `t.Columns(j).Width` 设置匹配列宽。
- **A5 底纹类型用错 → 纯黑背景** — ShadingType.SOLID 会变纯黑，必须用 CLEAR。COM 对应：单元格底纹 `t.Cell(r,c).Shading.BackgroundPatternColor = 0xD9D9D9`（浅灰），不要用 0（黑）。
- **A6 列表用手打符号** — LLM 直接插 `•`/`①` 模拟列表 → 后续无法自动重排编号。COM：公文场景列表少；若需列表用 `ListFormat.ApplyListTemplate` 或手动编号，且保持编号无缺口（见公文约定）。
- **A7 横版传竖版尺寸** — 设横向页面时若传了横向尺寸，库内部会再交换一次又变回竖版。COM 不涉及 docx-js 交换，但设横向时仍要显式：`d.PageSetup.Orientation = 1`（横向）+ 手动交换 PageSetup.PageWidth/PageHeight。
- **A8 删除段落留空段落** — 修订模式下删除段落要连段落标记一起删，否则留空段。COM 对应：`para.Range.Delete()`（含 ¶ 标记），不要只删文字。
- **A9 修订标记用错元素** — 删除文本必须用 `w:delText` 而非 `w:t`。COM 侧：用 `TrackRevisions` 模式操作即可，不要手工改 XML。
- **A10 特殊字符/引号** — 智能引号、中文引号在 XML 需实体。COM 侧：直接 `Range.Text = '中文"引号"'` 即可，COM 自动处理实体，无需手工转义。

## File-corruption pitfalls (B类: 文档损坏类, 社区/GitHub 高频)

- **B1 python-docx 生成的文件 MS Office 报损坏**（python-docx #446）— Google Docs/LibreOffice 打开正常但 MS Office 报"检测到损坏"。COM 路线天然规避（Word 自己写出的文件不会打包错），但若用户给的是第三方库生成的文件：打开后先 `d.SaveAs(path, 16)`（wdFormatXMLDocument）重存一遍，让 Word 修复打包。
- **B2 兼容模式 + 修订跟踪 = 保存崩溃**（Microsoft Q&A）— 开启兼容模式时带修订保存会报"文件错误"。规避：操作前检查 `d.CompatibilityMode`；如 >15（Word 2013+ 兼容模式）且要保留修订，先 `d.SaveAs2(path, 16)` 提升格式再改。修复损坏文档：新建文档→关修订→全选复制→贴入→重新打开修订→另存新名。
- **B3 多级列表编号"千古难题"**（Reddit/La Trobe guide）— AI 生成的多级标题没绑定到样式/多级列表，编号不连续、重启乱序。COM 侧：多级标题用 `para.Style`（Heading1/2/3）驱动编号，不要手打"1.1"数字；验证时扫描段落前缀看编号是否连续（与公文"编号无缺口"规则一致）。

## Incident 12: 生僻字缺字形被 Word 回退成微软雅黑 (大屯海"薸"事故, 2026-08-13)

生成"大屯海水生态现场调研与检测方案.docx"时，"大薸"的"薸"字没按仿宋_GB2312 渲染，而是微软雅黑。

根因：**仿宋_GB2312 / 楷体_GB2312 是 GB2312 字符集（6763 汉字），缺"薸溇垚犇"等生僻字**。Word 缺字形时回退到系统 UI 字体（微软雅黑），不是相似字体。`SubstituteFont` 只能映射"字体不存在"，管不了"字形缺失"。

修复（见 SKILL.md「Font coverage」）：fontTools 检查 cmap → 缺字段落 NameFarEast 换"仿宋"（FangSong 全覆盖 GBK）。**生成时直接指定"仿宋"而非"仿宋_GB2312"可根治**。

**补充实测（2026-08-13）：尝试过 FontLink 注册表方案（给仿宋_GB2312 添加 SystemLink → FangSong），导出 PDF 后"薸"字仍用 MicrosoftYaHei 子集——现代 Word 的 DirectWrite 渲染不走 FontLink 回退链，此方案无效。唯一可靠方案 = 文档内直接使用覆盖字体（"仿宋"/FangSong），用 pymupdf `page.get_fonts()` 验证 PDF 无 MicrosoftYaHei。**

## Incident 13: 表格默认段落属性未设置 (大屯海, 2026-08-13)

生成表格时未显式设置：段前段后 0 磅、取消"对齐到网格"（DisableLineHeightGrid=True）、取消"自动调整右缩进"（AutoAdjustRightIndent=False）。文档开"文档网格"排版时表格行距异常。

修复（见 SKILL.md「Table cell defaults」）：遍历单元格设 `SpaceBefore=0; SpaceAfter=0; DisableLineHeightGrid=True; AutoAdjustRightIndent=False`。实测默认值 DisableLineHeightGrid=0、AutoAdjustRightIndent=-1（即默认勾选），必须显式覆盖。

## Incident 14: 表格自适应单次不够，需三连 (大屯海, 2026-08-13)

单次 `AutoFitBehavior(1)`（窗口自适应）后列宽有时不合理。用户确认标准流程：**窗口→内容→窗口三连**（AutoFitBehavior(1)→(2)→(1)），三连后 PreferredWidthType=1（百分比整表占页面），效果最稳定。见 SKILL.md「Table auto-fit」。

## Incident 15: 单元格文本多加 \r → 每格空行 (2026-08-16)

`t.Cell(i,j).Range.Text = 'x\r'` 导致每个单元格多一个空段落。根因：`Cell.Range` 自带段落标记，追加 `\r` 等于再插一个空段。修复：直接 `= 'x'`。**SKILL.md 主文旧示例 `'x\r'` 与 cheatsheet `cell.Range.Text = names[r-1]` 互相矛盾——本事故就是照抄主文踩的坑，已统一为不带 \r。**

## Incident 16: fsType=2 受限嵌入 → PDF 丢字形 (2026-08-16)

标题"方正小标宋简体"在 PDF 中全部中文字形消失，只剩西文 "XX"；docx 文字完好、字体未内嵌。根因链：① GitHub 镜像 Mac 字体内部家族名是英文（`FZXiaoBiaoSong-B05S`）且 Windows 名字记录缺 nameID 16 → Word（DirectWrite）按内部名匹配解析不到，注册表显示名是摆设；② 即使能匹配，`OS/2.fsType=2`（Restricted 受限嵌入）→ Word 导出 PDF 直接丢弃字形而非替换。修复：fontTools 改中文家族名（补 nameID 16）+ fsType 2→8 + 用户级重装（RemoveFontResourceW→os.replace→AddFontResourceW→WM_FONTCHANGE）。详见 SKILL.md「Font embedding」。

## Incident 17: 双数字引用 [10]+ 偏移量漏 ] (2026-08-16)

`apply_script(d, '[12]', 0, 3, 'sup')` 只覆盖前 3 字符，`]` 落空——`"[12]"` 是 **4 字符**。单数字 `[1]`-`[9]`（3 字符）侥幸正确，双数字全错。**更隐蔽的是验证脚本用了同样的 `range(3)`，与生成代码共享盲区，输出"✓ 完整上标"假阳性**——用户在 PDF 渲染里发现 `]` 是正常字号。修复：生成与验证的偏移/长度一律用 `len(needle)`；教训 = **验证代码不能抄生成代码的同一份偏移量**。用户原话："两位数及以上很容易又出问题"——按长度/偏移处理文本的代码必须对 10+ 的编号做长度泛化。

## Incident 18: Heading 样式颜色残留 → 标题深蓝 (2026-08-16)

`para.Style = -2`（Heading1）后段落呈深蓝色——Heading 样式自带颜色/西文字体，生成脚本只覆盖了 Name/NameFarEast/Size，没覆盖颜色。修复：样式化后显式 `rng.Font.Color = 0`（黑色）。验证：遍历全部段落断言 `Font.Color == 0`（用户要求公文全文纯黑）。

## Incident 19: 表题与表格分页 (2026-08-16)

12 表长文档中部分表题孤立在页底、表格从下一页开始。修复：表题段 `KeepWithNext = True`（与下段同页）。验证：表题段 `Range.Information(3)`（wdActiveEndPageNumber）与表格首格页码相等。

## Incident 20: GB/T 7714 顺序编号乱序 (2026-08-16)

正文第一个引用是 [11]、第二个是 [12]，[1] 反而在中段——编号没按正文首次出现顺序。用户当场抓包（"[11] 应该是 [1] 吧"）。顺序编码制铁律：**先收集正文引用出现顺序，再据此编号并排列表**。验证：从正文按出现顺序提取 `[N]`，断言严格递增且与列表一一对应。

## Incident 21: 化学式多段角标偏移越界 → 波及后续字符 (SO42-, 2026-08-16)

**现象：** 正文"硫酸盐（SO42-）"中 '4' 不是下标、'2-' 不是上标，末尾 ')' 反而变成上标。用户原话："'4'应该是下标'2-'应该是上标，最后的'）'是正常大小"。

**根因链：**
1. `'SO42-'` 是 5 字符（S0 O1 4-2 2-3 --4）：'4' 下标应在索引 2，'2-' 上标应在索引 3-4。旧代码写成 `("SO42-", 4, 6, "sup")`——从 '-' 开始（漏掉 '4' 下标和 '2' 上标），且 **rel_end=6 越过 needle 末尾**，`d.Range(base+4, base+6)` 把 needle 之后的 ')' 也包进上标范围。
2. 验证函数照抄同一份偏移量（共享盲区，同 Incident 17），且 **sub 检查读的是 `Font.Superscript` 而不是 `Font.Subscript`**（对 sub 检查"不是上标"就误判为"是下标"），target 还用了 0（Word True 读回 -1）——三重盲区叠加导致"自检通过"。

**修复：** 拆两段调用 `("SO42-", 2, 3, "sub")` + `("SO42-", 3, 5, "sup")`；验证按 kind 读对应属性（sup→Superscript、sub→Subscript），target=-1。PDF 渲染层独立确认：'4' → 10.56pt bbox 向下伸（下标）、'2-' → 10.56pt bbox 向上收（上标）、'）' → 15.96pt 正常。

**教训：** ① rel_end 必须 ≤ len(needle)，偏移量基于 needle 自身索引；② 多段角标（酸根离子 SO42-/NO3-/PO43-）拆多次调用；③ 验证必须读与被验证属性一致的属性，target 统一 -1；④ 渲染层（PDF span 字号/bbox）是最终裁决，COM 自检只是中间层。

## Incident 22: 多级列表 %N 占位符引用错级别 → 第三层变"一." (2026-08-16)

**现象：** 用 Word 多级列表做公文四层编号（一、/（一）/1./（1）），第一、二层正确，第三、四层显示"一."/"（一）"而不是"1."/"（1）"。

**根因链：**
1. **`NumberFormat` 占位符 `%N` 引用的是「第 N 级」的编号值，不是本级**。Level3 写 `'%1.'` → %1 引用 Level1 的编号（中文样式 一、二）→ "一."；Level4 写 `'（%1）'` → "（一）"。Level2 的 `'（%1）'` 碰巧正确（Level1/2 都是中文计数），掩盖了问题。
2. 另一个坑：NumberStyle 用 `38` 得到"壹、"（中文大写）；正确的中文小写计数（一、二、三）是 **39**。

**修复：** 本级编号引用写 `%N`（N=当前级别）：Level2 `'（%2）'`、Level3 `'%3.'`、Level4 `'（%4）'`；层级式（1.1.1）每级写 `%1.%2...%N`（实测 `'%1'`/`'%1.%2'`/`'%1.%2.%3'`）。套用三步：`d.ListTemplates.Add(True)` 建模板 → 整段区域一次 `ApplyListTemplate(lt, False, 1)` → 逐段 `ListLevelNumber = level`（逐段 Apply 会全部变成第一级）。验证：遍历每段 `ListFormat.ListString` 断言编号序列。

**教训：** ① `%N` = "第 N 级的编号值"；② 验证必须检查**每一级**的实际编号，不能只看前两级；③ NumberStyle 39=中文小写、38=中文大写；④ 编号字体跟随段落字体（PDF span 实测同字体），无需单独设 ListLevel.Font。

## User conventions for Chinese government proposal documents (公文)

- **No source citations.** Reports written in a bureau's name must NOT carry "来源于官方网站/数据来源/网络检索" annotations or URLs — state facts directly. Remove `（…数据来源：…）` parentheticals entirely.
- **Numbering must be gapless**: no `(1)` followed by `(3)`, no `①` then `③`. Renumber paragraph prefixes; verify by scanning paragraphs for starts-with patterns.
- **Appendix/附表 titles need names** (e.g. `附表1 晋城市水生态环境监测监管能力建设项目清单`), consistent across body AND TOC. Multi-device parameter tables need a first-column 设备名称 (device name) mapping each row to its equipment.
- **Content numbers must cross-check**: investment subtotals = grand total, funding shares sum to 100%, year splits sum to total, equipment counts match the checklist tables. Report remaining inconsistencies rather than silently choosing.
