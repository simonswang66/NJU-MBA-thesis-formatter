# NJU MBA Thesis Formatter · 南京大学MBA论文格式化 Skill

一个适配 Claude Code 等 Agent 环境的技能：把 Markdown / docx 草稿一键刷成**符合《南京大学商学院MBA学位论文写作要求与规范》**的 Word 文档，支持从分章 md 整篇装配学位论文（封面、声明、摘要、目录、正文、参考文献、附录、致谢一次成型）。

核心思路：**格式规则全部沉淀在脚本与模板里，Agent 只做搬运**——你专心写内容，排版交给 skill。

## 30 秒开始

把下面这段话直接发给有 shell 权限的 AI Agent（Claude Code / Codex / Cursor 等）：

```text
帮我安装 NJU-MBA-thesis-formatter 这个 skill。请按下面步骤做：
1. 按你所在的平台选择安装目录：Claude Code → ~/.claude/skills/
   （Windows 为 %USERPROFILE%\.claude\skills\）；WorkBuddy → ~/.workbuddy/skills/
   （Windows 为 %USERPROFILE%\.workbuddy\skills\，装完需重启 WorkBuddy）；
   Codex → ~/.codex/skills/；没有 skills 机制的平台跳过复制，记下原目录路径即可
2. 把 NJU-MBA-thesis-formatter 文件夹整体复制到安装目录下
3. 验证：NJU-MBA-thesis-formatter/ 下应看到 SKILL.md、scripts/、assets/、references/
4. 安装依赖：pip3 install python-docx lxml defusedxml
5. 告诉我安装好了
```

平台没有 skills 机制也可以零安装使用——直接把这句话发给 Agent（替换为实际路径）：

```text
读 <本目录绝对路径>/SKILL.md，按里面的说明帮我把论文格式化为南大 MBA 学位论文格式
```

装好（或读完 SKILL.md）后对 Agent 说：

```text
帮我把这篇论文 md 格式化为南大 MBA 学位论文格式
```

或者整篇装配（推荐，效果见下）：

```text
把我的分章 md 按 摘要→Abstract→第1-7章→参考文献→附录→致谢 拼成带 @unit 标记的合并稿，
用 23E 模板模式装配成 docx
```

## 能力

- 🎓 **三种模式**：学位论文 `thesis`（封面/原创性声明/中英摘要/目录/参考文献/附录/致谢）、课程作业 `coursework`（作业封面+核心样式）、纯格式化（只有样式）
- 🏗 **23E 模板模式·整篇装配**：以南大 MBA 论文模板为样式库重建节结构——摘要/Abstract/目录罗马页码、正文阿拉伯页码重新计 1、每章页眉自动换成章标题；目录为 Word 域（打开后 F9 更新）
- 📐 **规范样式全覆盖**：章标题黑体 16pt 居中、节标题黑体 14/13pt、正文宋体小四首行缩进 2 字符行距 20pt、图表题五号加粗居中、三线表、页下注 ①②③
- 🧮 **LaTeX 公式 → OMML**（经 pandoc，Word 原生可编辑，非图片）
- 📊 **Mermaid 图 → PNG** 自动渲染插入（需 Playwright，可选）
- 🔁 **章级增量更新**：改了一章只重渲染该章，其余部分（含你在 Word 里手工精调的格式）原样保留；表格默认"留骨架、只换文字"
- 🧊 **冻结表/横排表配置**：手工调好的表锁定不动；超宽表自动开横排节
- ✅ **文本一致性核查**：docx 与 md 源逐单元比对，差异归一化报告
- 🧹 **自动后处理**：清理未引用残留图、作者元数据改写为当前系统用户、版式样式归一（页眉下边框、图表注缩进、参考文献去自动编号）
- 📥 **多种输入**：`.md` 直接解析；`.txt` 按 md 规则解析；`.docx` 按 Word 标题样式识别层级（全文无标题样式时自动按"第X章 / X.Y / X.Y.Z"文本模式兜底）

## 适合 / 不适合

**✅ 合适**：南京大学 MBA 学位论文整篇装配与日常更新；MBA 课程作业排版；任何需要"md 写完直接出规范 Word"的场景

**❌ 不合适**：非南大格式规范（可换模板尝试，样式名需一致）；依赖 Word 自动交叉引用/自动编号管理图表的写作流（本 skill 照排 md 里的编号，不自动重编）；PDF 输入（请先转成文字稿）

## 平台支持

能力全部在 `scripts/` 的 Python 脚本里，SKILL.md 只是说明书——所以任何**能读文件、能跑 shell** 的 Agent 都能用：

| 平台 | 用法 |
|------|------|
| Claude Code | 原生支持：装到 `~/.claude/skills/` 后自动发现，对话中直接触发 |
| WorkBuddy（腾讯） | 结构完全兼容（SKILL.md＋scripts/＋assets/＋references/）：复制到 `~/.workbuddy/skills/`（Windows 为 `%USERPROFILE%\.workbuddy\skills\`）并重启 WorkBuddy；或在客户端「技能市场 → 上传技能」直接上传本目录 zip |
| Codex | 新版 Codex 支持同一套 Agent Skills 标准：装到 `~/.codex/skills/`；旧版按下方通用方式 |
| OpenCode / Cursor 等 | 通用方式：把本目录放在任意位置，对 Agent 说"读 `<路径>/SKILL.md`，按里面的说明格式化我的论文" |
| Coze（扣子）等云端 Bot 平台 | ❌ 不适用：代码节点是受限沙箱（禁 pip、仅标准库＋少量内置库、无文件系统持久化），装不了 python-docx/pandoc，也没有本地文件落盘——请改用本地 Agent |

判断标准一句话：平台得有**本地文件系统读写＋能跑本地 Python（可 pip 装包）**；纯云端沙箱平台（Coze、Dify Cloud 类）不满足，不必尝试。

## 三种模式用法

```bash
# 学位论文（单文件 md，旧模式）
python3 scripts/format_engine.py -d 论文.md -m thesis

# 课程作业
python3 scripts/format_engine.py -d 作业.md -m coursework
```

通常不用记命令——直接对 Agent 说"论文"或"课程作业"，它会按 SKILL.md 选择模式。

## 23E 模板模式（学位论文整篇装配，推荐）

```bash
python3 scripts/build_thesis_docx.py 合并稿.md \
    --template assets/MBA学位论文模板格式（23E自制）.docx \
    --cover assets/cover-template.docx \
    -o 论文全文.docx
```

### 合并稿单元约定

把全文按论文顺序拼成一个 md，每个单元用注释标记包裹：

```markdown
<!-- @unit:abstract_cn -->
# 摘要
（正文……）
**关键词：** xxx；xxx
<!-- @/unit:abstract_cn -->

<!-- @unit:ch1 -->
# 第1章 绪论
## 1.1 研究背景
……
<!-- @/unit:ch1 -->
```

| 单元 ID | 内容 | 说明 |
|---------|------|------|
| `abstract_cn` / `abstract_en` | 中/英文摘要 | 目录域自动插在英文摘要之后 |
| `ch1` … `chN` | 各章 | `#` 标题照排为章标题 |
| `references` | 参考文献 | md 里自带 [1] [2] 编号 |
| `appendix_a` / `appendix_b` … | 附录 | 标题写 `附录A：xxx` |
| `acknowledgement` | 致谢 | |

分章写作的话，让 Agent 逐章读取并加上标记即可，一步完成。

### 章级增量更新与核查

```bash
# 只更新第3章（自动时间戳备份 docx 到同目录 bak/）
python3 scripts/section_replacer.py 论文全文.docx 合并稿.md --unit ch3

# 核查 docx 与合并稿文本一致性
python3 scripts/check_docx_text.py 论文全文.docx 合并稿.md
```

### 版式配置 `.docx-layout.json`

放在 docx 同目录，两项均可选：

```json
{
  "frozen": ["表5.2"],
  "landscape": ["表3.3"]
}
```

- `frozen`：在 Word 里手工调好的表，增量更新时整张原样保留（文字也不动）
- `landscape`：题注前缀匹配的宽表，渲染时自动开横排节（题注+表+资料来源注都在横排节内）

## 输入格式

| 输入 | 支持方式 |
|------|---------|
| `.md` | 最佳路径：`#`/`##`/`###` 映射章/节/小节 |
| `.txt` | 按 md 规则解析。无 `#` 标记的纯文本会全部成为正文——先让 Agent 按内容补标题层级再格式化 |
| `.docx` | 按 Word 标题样式（Heading 1–4）映射层级；全文无标题样式时自动按"第X章 / X.Y / X.Y.Z"文本模式识别 |
| PDF | 不支持，先转成文字稿 |

注意：23E 整篇装配模式只接受带 `@unit` 标记的合并 md；docx/txt 草稿让 Agent 先转成该格式即可。

## 依赖

| 依赖 | 用途 | 安装 |
|------|------|------|
| Python 3.10+ | 运行脚本 | — |
| python-docx / lxml / defusedxml | docx 读写 | `pip3 install python-docx lxml defusedxml` |
| pandoc | LaTeX → OMML 公式 | macOS `brew install pandoc`，Windows 用官网安装包（无公式可省） |
| Playwright | Mermaid → PNG | `pip3 install playwright && playwright install chromium`（无 Mermaid 可省） |

## 常见问题

**Q：打开生成的 docx，目录是空的？**
A：目录是 Word 域。全选（Ctrl/Cmd+A）后按 F9，或右键目录→"更新域"。

**Q：Word 提示"发现无法读取的内容"？**
A：先用 python-docx 重开文件定位（`python3 -c "from docx import Document; Document('文件.docx')"`）；Claude Code 用户装了 Anthropic document-skills 插件的话，可跑其 `office/validate.py` 做 schema 级校验。

**Q：某张表我在 Word 里精调过格式，更新章节时不想被碰？**
A：把表号加进 `.docx-layout.json` 的 `frozen` 清单。

**Q：生成的 docx 作者信息是谁？**
A：装配与更新末尾会自动把作者元数据改写为当前系统用户名，不会残留模板原作者信息。

**Q：模板样式名对不上？**
A：23E 模式依赖模板里的 `MBA论文*` 系列样式。换用自己的模板时，确保样式名一致（或让 Agent 参照 `scripts/thesis23e_engine.py` 的 STYLE 映射调整）。

## 目录结构

```
NJU-MBA-thesis-formatter/
├── SKILL.md                 # 技能说明（Agent 读取）
├── README.md                # 本文件
├── scripts/                 # 解析、渲染、装配、增量更新、核查、后处理
├── assets/                  # 23E 论文模板、学位论文封面、课程作业封面
└── references/              # 南大 MBA 格式规范速查表
```

## 说明

- 模板与封面均按学院公开规范自制，封面字段为占位符，请在 Word 中自行填写。
- 仅供同学间学习交流；格式以学院当年发布的《南京大学商学院MBA学位论文写作要求与规范》为准。

## 免责声明

本 skill 只是提升排版效率的工具，其产出**不构成可直接提交的最终稿**：

- 每次生成或增量更新后，务必在 Word 中通篇人工核对：目录（F9 更新后）、页码、图表编号与题注、表格线型、参考文献格式、页眉页脚等；
- 格式要求以学院当年发布的《南京大学商学院MBA学位论文写作要求与规范》为准；规范更新或模板差异导致的偏差，请以人工判断修正；
- 论文内容与格式的最终责任在作者本人。请勿完全依赖本 skill 的产出而不做检查。
