---
name: NJU-MBA-thesis-formatter
description: 将 .md 或 .docx 文件格式化为符合南京大学商学院MBA学位论文规范的 .docx 文档。当用户提到格式化文档、论文排版、MBA论文格式、作业格式刷、md转docx、docx格式整理时应立即使用此技能。支持论文模式（含封面/声明/摘要/目录/参考文献/致谢）和作业模式（仅核心样式）。
---

# MBA 论文/作业 Docx 格式化

将 `.md` 或 `.docx` 输入文件转换为符合《南京大学商学院MBA学位论文写作要求与规范》的 `.docx` 文件。

## 参数提取

| 参数 | 必填 | 说明 |
|------|------|------|
| 输入文件路径 | 是 | `.md`、`.txt` 或 `.docx` 文件 |
| 输出路径 | 否 | 默认放在**输入文件的同级目录**，文件名同输入但后缀 `.docx` |
| 模式 | 否 | `thesis`（学位论文）、`coursework`（课程作业）或不指定（纯格式化）。用户说"论文"/"thesis"/"学位论文"时启用论文模式；说"课程作业"/"作业"/"小论文"时启用课程作业模式 |

**三种模式对比**：

| 要素 | 论文模式 `thesis` | 课程作业 `coursework` | 纯格式化(默认) |
|------|-------------------|----------------------|----------------|
| 封面 | 学位论文标准封面 | 课程作业封面（logo+标题+姓名/班级/学号） | ❌ |
| 原创性声明 | ✅ | ❌ | ❌ |
| 中英文摘要 | ✅ | ❌ | ❌ |
| 目录 | ✅ | ❌ | ❌ |
| 章/节/正文排版 | ✅ | ✅ | ✅ |
| 图表题格式化 | ✅ | ✅ | ✅ |
| 参考文献 | ✅ | ❌ | ❌ |
| 附录 | ✅ | ❌ | ❌ |
| 致谢 | ✅ | ❌ | ❌ |
| 页眉/页码 | ✅ | ❌ | ❌ |

## 23E 模板模式（整篇装配）

以 `assets/MBA学位论文模板格式（23E自制）.docx` 为**样式库**，从带单元标记的合并稿整篇装配学位论文 docx。**与旧 thesis 模式并存，旧模式行为不变。**

```bash
python3 scripts/build_thesis_docx.py <合并稿.md> \
    --template assets/MBA学位论文模板格式（23E自制）.docx \
    [--cover assets/cover-template.docx] -o <输出.docx>
```

**合并稿单元约定**：全文按论文顺序拼接为一个 md，每个单元用 `<!-- @unit:ID -->` 与 `<!-- @/unit:ID -->` 注释包裹。单元 ID：`abstract_cn`（中文摘要）、`abstract_en`（英文摘要，目录域自动插在其后）、`ch1`…`chN`（各章，章内 `#` 标题照排为章标题）、`references`（参考文献）、`appendix_a`/`appendix_b`/…（附录）、`acknowledgement`（致谢）。摘要/参考文献/致谢单元内的 `#` 标题行由引擎以模板样式重渲染，不写进 md 也可以。让 Agent 把分章 md 拼成此格式即可（逐章读取、按上表顺序包裹标记）。

- `scripts/thesis23e_engine.py`（Thesis23EEngine，FormatEngine 子类）：模板只作样式库（剥离正文与节残留），按单元重建节结构——摘要/Abstract/目录罗马页码（自摘要起 I）、第1章起阿拉伯页码重新计 1、页眉＝单元标题；目录为 Word TOC 域（`\o "1-3"`，打开后 F9 更新）；章/节/图表题全部**照排原文**（md 已带编号，不自动重编）
- `content_parser.parse_merged_thesis()`：解析合并稿单元（`<!-- @unit:ID -->`），后处理再分类图表题／来源注／列表项／关键词（旧 `parse_markdown` 默认不启用）
- 封面自成一节（merge_cover.py 改）：封面节无页眉无页码，复用封面模板自身页面设置；图片关系改写覆盖 r:embed／r:id（VML）／r:link 三种形态
- 摘要/Abstract/目录标题用克隆样式「MBA论文前置标题」（无 outline，不进 TOC）；参考文献/附录/致谢进 TOC
- 章级增量更新用 `scripts/section_replacer.py`（见该文件 docstring）
- **文本一致性核查 `scripts/check_docx_text.py`**：`check_docx_text.py <docx> <合并稿.md>` 逐单元比对 docx 与合并稿文本（内容级归一：全角/半角、中英文间空格、`<br>`/换行、合并单元格免疫），冻结表差异单列
- **版式配置 `.docx-layout.json`**（放 docx 同目录）：`frozen`＝用户手工定版禁止重渲染的表（replacer 按题注锚点原样回置）；`landscape`＝横排宽表题注前缀（引擎渲染到该表自动开/收横排节，题注+表+注在横排节内）
- **后处理 `scripts/docx_postprocess.py`**（装配与增量更新末尾自动执行）：①作者元数据（dc:creator／cp:lastModifiedBy）改写为系统用户——python-docx 默认模板写 "python-docx"、模板可能残留原作者，产物每次生成/编辑后都必须过一遍；②清理 word/media 中未被引用的图片（模板残留、增量替换废弃图），校验范围含页眉页脚等全部 XML part，不会误删在引图；③`--fix-styles` 版式归一（样式层，不动内容）：图表题/注样式统一加粗居中 10.5pt、参考文献样式去自动编号（numPr 与 md 自带 [N] 叠加会序号重复）、各节页眉加 0.75pt 下边框——23E 模板（WPS 存盘）样式缺规格且带自动编号，装配时由 `fix_layout_styles_doc` 自动归一

## 输入方式

**自然语言**：
```
帮我把这篇论文格式化为MBA论文格式
论文
输入：/path/to/paper.md
```

**参数简写**：
```
-d /path/to/input.md -m coursework -o /path/to/output.docx
```
参数：`-d` 输入文件，`-m` 模式(thesis/coursework)，`-o` 输出路径(可选，未指定绝对路径时，输出至输入文件相同目录)

## 处理流程

### 1. 解析输入
- 读取输入文件，调用 `scripts/content_parser.py` 解析为结构化元素
- `.md` / `.txt` 文件：按标题层级（# → 章、## → 节、### → 小节）、段落、引用块、表格、图片解析；无 `#` 标记的 txt 纯文本全部成为正文，应先补标题层级
- `.docx` 文件：按 Word 样式映射到对应元素类型；全文无 Heading 样式时按文本模式兜底（第X章→章、X.Y→节、X.Y.Z→小节）

### 2. 格式化
调用 `scripts/format_engine.py` 应用 NJU MBA 论文格式。引擎自动处理：
- `**粗体**` → docx 粗体格式
- `\(LaTeX\)` → 行内 OMML 数学公式
- `\[LaTeX\]` → 独立行 OMML 数学公式（居中）
- ` ```mermaid ` 代码块 → 通过 mermaid.ink 渲染为 PNG 图片插入

```python
from scripts.format_engine import FormatEngine

engine = FormatEngine(mode="thesis")  # or "coursework"

for el in elements:
    if el['type'] == 'chapter': engine.add_chapter(el['text'])
    elif el['type'] == 'h1': engine.add_h1(el['text'])
    elif el['type'] == 'body': engine.add_body(el['text'])  # auto-parses **bold** & \(LaTeX\)
    elif el['type'] == 'display_math': engine.add_display_math(el['text'])
    elif el['type'] == 'mermaid_image': engine.add_image(el['src'])
    elif el['type'] == 'table': engine.add_three_line_table(el['headers'], el['rows'])

engine.save(output_path)
```

### 3. OOXML 后处理（论文模式）
若内容含脚注标注，调用 `scripts/footnote_inserter.py` 通过 OOXML 插入页下注：
```python
from scripts.footnote_inserter import insert_footnotes
insert_footnotes(input_docx, output_docx, footnotes, positions)
```

### 4. 封面插入
- **论文模式**：使用 `assets/cover-template.docx`（南京大学专业型硕士标准封面，含学校 logo、论文题目、作者导师信息、原创性声明等）
- **课程作业模式**：使用 `assets/assignment-cover.docx`（课程作业封面，含学校 logo、课程作业标题、姓名、班级、学号）
- 封面模板作为文档首页插入，正文从第二页开始

## 核心格式规范

详细规范见 `references/`：`format-spec.md` 速查表，及学院《南京大学商学院MBA学位论文写作要求与规范》原文 PDF。关键规则：

- **章标题**：黑体 16pt 加粗居中，段前 24pt 段后 18pt，每章另起一页
- **一级节**：黑体 14pt 顶左，段前 24pt 段后 6pt
- **二级节**：黑体 13pt 左缩 2 字，段前 12pt 段后 6pt
- **正文**：宋体/Times New Roman 12pt，两端对齐，首行缩进 2 字，行距 20pt
- **图表**：宋体 10.5pt 加粗居中，三线表
- **页下注**：①②③ 连续编号
- **编号规则**：图/表分章连续编号（图1.1、表2.2）

## 输出前检查
- 字体是否正确（中文宋体/黑体，英文 Times New Roman）
- 行距、段前段后是否符合规范
- 三线表格式是否正确
- 图表编号是否符合"章.序号"规则
- 论文模式：封面、声明、摘要、目录、参考文献、致谢是否齐全
