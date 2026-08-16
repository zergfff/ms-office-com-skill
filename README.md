# word-com-chinese-skill

**中文公文 Word COM 自动化技能** —— 专用于 Windows + Microsoft Word 的中文公文/报告文档（.docx/.doc）生成与修改，通过 `win32com.client` (pywin32) 驱动真实 Word。

> ⚠️ **适用范围：仅 Windows + Microsoft Word + COM (win32com.client)，专用于中文文档（公文、报告、方案）。**
> ❌ 不适用于 Linux/macOS、WPS Office、Excel/PowerPoint、python-docx 纯库路线、Microsoft Graph 云端 API。

> 💡 **Token 成本：** SKILL.md 约 620 行 / ~18K tokens，命中时整份注入 agent 主上下文（1M 窗口下约占 1.8%）。references/ 和 scripts/ 按需加载，不常驻。若对 token 敏感，可用精简模式（见 SKILL.md「Token 优化」章节）。

为 AI agent（Hermes / Claude Code / Codex / OpenClaw / Cursor 等）提供中文公文处理能力：GB/T 9704-2012 格式、GB/T 7714 参考文献、上下角标语义判定、PDF 书签导出、公文字体自动安装、防弹窗卡死等。

## 为什么用 COM？

纯库（python-docx）会丢失或扁平化复杂结构：合并表格、样式、大纲级别、页眉页脚、域、图片锚点。COM 驱动真实 Word 应用，保留一切——本技能即为此设计。

## 目录结构

- `SKILL.md` — 技能主体（Golden Rules、Word 核心模式、查找替换、样式大纲、字体与替换、Font embedding(fsType)、**模板优先规则（有模板用模板/无模板用默认）**、GB/T 9704-2012 公文格式、对齐与表格默认、GB/T 7714 参考文献、上下角标语义、表格手术与自适应、页码、**多级列表自动编号（一、/1.1.1）**、PDF 书签导出、防弹窗卡死、Dispatch 选型、健康检查、预算级联）
- `references/win32com-cheatsheet.md` — 已验证可运行的 Word 代码片段（安全批量替换、样式常量、表格列增删、合并单元格读取、图片插入/缩放、字体、PDF 导出、页码、参考文献、上下角标）
- `references/pitfalls.md` — 真实事故档案（无限目录复制、陈旧进程损坏、长字符串 Find 错误、图片容器段落丢失、样式 off-by-one、LLM 生成坑 A1-A10、文档损坏坑 B1-B3、Incident 15-23：单元格 \\r 空行 / fsType 丢字形 / 双数字引用偏移 / Heading 颜色 / 表题同页 / 引用顺序编号 / 化学式多段角标偏移越界 SO42- / **多级列表 %N 占位符引用错级别** / **ListLevelNumber 非列表段落默认返回 1**）
- `scripts/ensure_fonts.py` — 检测本机是否缺少公文字体，缺失则自动下载安装（用户级，无需管理员；GitHub 源超时快速失败）

## 安装方式（各 Agent）

### 方式一：Prompt 安装（最快，把下面这段发给你的 agent）

> 请从 https://github.com/zergfff/word-com-chinese-skill 下载安装 `word-com-chinese-skill` 技能：
> 1. 克隆或下载该仓库到你的 skills 目录（各 agent 的 skills 目录见下）；
> 2. 把 `SKILL.md`、`README.md`、`references/`、`scripts/` 整个文件夹放入 skills 目录；
> 3. 之后处理中文 Word 公文时自动加载此技能。

### Hermes Agent

```bash
# 复制到当前 profile 的 skills 目录（productivity 分类）
cp -r word-com-chinese-skill ~/AppData/Local/hermes/profiles/<profile>/skills/productivity/word-com-chinese-skill
# 或直接用 hermes skills install（若配置了 skills hub）
```

### Claude Code

```bash
mkdir -p ~/.claude/skills && cp -r word-com-chinese-skill ~/.claude/skills/
# 或
npx skills add zergfff/word-com-chinese-skill
```

### Codex (OpenAI)

```bash
mkdir -p ~/.codex/skills && cp -r word-com-chinese-skill ~/.codex/skills/
```

### OpenClaw / Cursor / 其他

任何支持 Anthropic 风格 SKILL.md（YAML frontmatter：`name` / `description`）的 agent：
把整个文件夹复制到该 agent 的 skills 目录即可。

## 环境要求

- **Windows 11 + Microsoft Word 2024 (LTSC)** — 首选；Word 2016+ 均可（Word COM 对象模型稳定）
- 已安装 Microsoft Word
- Python + pywin32：`pip install pywin32`

## 功能一览

| 模块 | 能力 |
|---|---|
| **格式** | GB/T 9704-2012 公文格式速查：标题 方正小标宋简体2号居中 / 一级标题 黑体3号 / 二级标题 楷体_GB2312 3号 / 正文 仿宋_GB2312 3号两端对齐首行缩进2字符行距28磅；字体双属性（中文字体 NameFarEast + 西文字体 Name）；缺字自动回退（SubstituteFont）；**生僻字缺字形自动换覆盖字体（fontTools 检测 cmap，"薸溇垚犇"在 GB2312 字库缺失 → 换 FangSong 全覆盖）**；标题/表题/图题居中无缩进；图片段单倍行距；表格内文字上下/左右居中+无缩进+单倍行距+段前段后0磅+取消对齐网格/自动右缩进；表头加粗+跨页重复标题行 |
| **内容** | GB/T 7714-2015 参考文献（顺序编码制、各文献类型、正文上标[1]引用）；上下角标语义判定（m3→m³上标、CO2→CO₂下标、Ca2+→Ca²⁺上标、数值不动） |
| **表格** | 表格手术（加列/删空列/合并单元格安全读写/尾部追加表）；AI 表格超页边距修复（**窗口→内容→窗口三连 AutoFitBehavior**）；表头重复 |
| **交付** | 页脚页码（PAGE 域，小五 9pt Times New Roman）；PDF 导出（CreateBookmarks=1 按标题建书签，先刷新TOC/域）；目录自动刷新 |
| **安全** | 防弹窗卡死（DispatchEx + DisplayAlerts + AutomationSecurity + is_locked + Open/Close 全参数 + 线程超时强杀）；崩溃恢复（taskkill）；每次编辑后健康检查；预算级联核对 |
| **字体** | `scripts/ensure_fonts.py` 自动检测+下载安装缺失的公文字体（仿宋_GB2312/楷体_GB2312/方正小标宋简体/方正楷体_GBK/黑体等，用户级安装无需管理员）；**GitHub 源快速失败策略（连接10s/总限时30s 超时即跳过，国内被墙不傻等；失败可设 HTTPS_PROXY 代理重试）** |

## License

MIT
