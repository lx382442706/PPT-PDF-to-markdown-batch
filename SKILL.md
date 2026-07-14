---
name: PPT&PDF-to-markdown-batch
description: >
  批量将文件夹内的 PDF/PPT/PPTX 转换为 Markdown。PPT/PPTX 先通过 PowerPoint COM 导出为 PDF，再走
  PyMuPDF 快照 + 文本提取管线；普通 PDF 采用 MinerU API 精准解析。自动清理人名和版本号，日期优先从文件名提取。
  默认使用 DeepSeek Chat LLM 后处理排版（带3次重试保护），标题从 H1 起步。
  Use when user says: PDF转markdown、批量PDF转md、PPT转md、PPTX转markdown、pdf to markdown、/PPT&PDF-to-markdown-batch
agent_created: true
---

# Skill: 文档 → Markdown 批量转换（PPT→PDF→快照 + 普通PDF MinerU）

双管线批量文档转 Markdown 工具：
- **PPT/PPTX**：PowerPoint COM 导出为 PDF → PyMuPDF (fitz) 逐页快照 + 文本提取 → LLM 智能排版
- **PPT 导出 PDF**：与 PPT/PPTX 共用同一管线（自动检测 16:9/A4横向/4:3 比例分流）
- **普通 PDF**：MinerU API 精准解析 → 后处理管线
- 自动清理输出名称：去除人名"来潇"、版本号 V1/V2/V3 等
- 日期优先从文件名提取（支持嵌入位置如"报告-2025.7.29"），无日期时使用文件修改时间
- 标题从 H1 起步

## 支持格式

| 格式 | 管线 | 后处理特性 |
|------|------|-----------|
| PPT/PPTX | **PowerPoint COM → PDF** → PyMuPDF 快照 + 文本→EasyOCR降级 | PPT→PDF导出 → 逐页截图+文本提取（不足时EasyOCR降级）→ LLM排版 |
| PDF（PPT导出） | **PyMuPDF 快照** + 文本提取→EasyOCR降级（与 PPT 共用管线） | 自动检测 16:9/A4横向/4:3 分流，文本不足即用EasyOCR → LLM排版 |
| PDF（普通） | MinerU API (vlm) | 正文截取、图片宽度100%、标题规范化（H1 起步）、页码删除 |

## 核心能力

| 步骤 | 说明 |
|------|------|
| **PPT→PDF导出** | PowerPoint COM 将 PPT/PPTX 导出为 PDF（ppSaveAsPDF） |
| **PDF快照** | PyMuPDF (fitz) 逐页渲染 1920px 宽 PNG 截图（PPT/PPTX/PPT导出PDF 通用） |
| **PDF文本** | PyMuPDF 文本提取 + EasyOCR 降级（文字过少时自动切换） |
| **LLM排版** | DeepSeek Chat 智能整合（默认）：去冗余、章节提取、段落优化、表格保留，带 3 次指数退避重试 |
| **名称清理** | 自动去人名"来潇"、版本号 V1/V2/V3 |
| **日期提取** | 优先文件名（含嵌入位置如"报告-2025.7.29"），无日期用文件修改时间 |
| **PDF解析** | MinerU vlm/pipeline 模型解析 → ZIP(MD+图片+JSON)（仅普通PDF） |
| **正文截取** | 普通 PDF 自动定位「摘要」或第一章 |
| **图片宽度** | `![](path)` → `<img src="path" width="100%">` |
| **页码删除** | 多格式页码清除（仅普通PDF） |
| **标题规范** | 自动识别编号标题，从 H1 (#) 起步 |
| **唯一编码** | 自动生成 8 位唯一编码，插入 ES 链接并重命名源文件 |
| **降级保护** | PPT→PDF 导出失败跳过该文件；PPT导出PDF 截图失败回退 MinerU |
| **EasyOCR 降级** | PyMuPDF 提取文字过少（<30字符）时自动调用 EasyOCR 对截图进行文字识别，提升图文混合场景识别率 |
| **中断恢复** | LLM 调用 3 次指数退避重试，分段失败保留原文，单文件失败不影响后续 |

## 唯一编码集成

每个文档转换前通过 `生成唯一编码` 技能生成 8 位唯一编码，转换后：

1. **源文件重命名**：`核心名.pdf` → `核心名【UID】.pdf`
2. **ES 链接**：在 `.md` 文件顶部插入引用行

```
> 原始文件: [006T9600](es:006T9600) | 核心名【006T9600】.pdf
```

> 文件夹和 .md 文件名不含 UID、日期、版本号。

生成命令（在 vault 根目录执行）：

```powershell
.\.claude\scripts\unique-code.ps1
```

前置依赖：`生成唯一编码` 技能已安装（`.claude/skills/生成唯一编码/SKILL.md`）。

---

## Inputs

| 参数 | 必需 | 说明 |
|------|:---:|------|
| `source_dir` | 是 | 包含 PDF/PPT/PPTX 文件的源文件夹路径 |
| `-o / --output` | 是 | 输出根目录 |
| `--model` | 否 | 模型版本：`vlm`(默认/推荐)、`pipeline`、`MinerU-HTML` |
| `--no-formula` | 否 | 关闭公式识别 |
| `--no-table` | 否 | 关闭表格识别 |
| `--language` | 否 | 文档语言，默认 `ch` |
| `--batch-size` | 否 | 单批次文件数，默认 50（最大 50） |
| `--no-postprocess` | 否 | 跳过后处理（仅原始转换） |
| `--no-llm` | 否 | 跳过 LLM 智能排版（保留其他后处理） |
| `--no-chapter-reorg` | 否 | 跳过 PPT 章节去重重组 |
| `--remove-lines` | 否 | 手动指定要删除的冗余行关键词，多个用逗号分隔（如：`中核华兴,NIC 2026`） |

## Output

子文件夹命名格式为 **YYYY-MM-DD 文件核心名**：
- 优先从文件名提取日期（支持 `YYYY-MM-DD`、`YYYYMMDD`、`YYYY.MM.DD`、`YYYY年MM月DD日`、`YYMMDD` 等格式）
- 文件名无日期时，自动使用文件最后修改日期
- 文件夹和 .md 文件名不含 UID、日期、版本号

```
output_dir/
├── 2025-09-30 核电混凝土智能高频振捣棒研发实施方案/
│   ├── 核电混凝土智能高频振捣棒研发实施方案.md  ← 含 ES 链接
│   ├── attachments/
│   │   ├── xxx.jpg
│   │   └── ...
├── 2025-10-15 另一个文件名/
│   ├── 另一个文件名.md
│   ├── attachments/
│   │   └── ...
└── ...
```

> 源文件留在原位置，就地重命名为 `核心名【UID】.ext`，不移动、不复制到输出目录。

## Process

### PPT/PPTX 管线（导出 PDF → 截图）

```
PPTX/PPT 文件
  ├── 1. PowerPoint COM 导出为 PDF (ppSaveAsPDF)
  ├── 2. PyMuPDF (fitz) 逐页渲染 PNG 截图 → 1920px 宽
  ├── 3. PyMuPDF 原生文本提取 → 无需 OCR
  ├── 4. 构建 Markdown: # 第N页 + 文本 + <img ...>
  └── 5. DeepSeek LLM 智能排版 → 最终 MD
```

输出示例：
```markdown
# 第1页
- 项目背景介绍
- 核心目标与实施路径

<img src="attachments/slide_01.png" width="100%">

# 第2页
技术方案详解

<img src="attachments/slide_02.png" width="100%">
```

### PPT 导出 PDF 管线（截图 + 文本提取，与 PPT/PPTX 共用）

```
PPT 导出的 PDF 文件
  ├── 1. 自动检测 16:9 / A4横向 / 4:3 页面比例 → 确认 PPT 导出，分流至截图管线
  ├── 2. PyMuPDF (fitz) 逐页渲染 PNG 截图 → 1920px 宽
  ├── 3. PyMuPDF 原生文本提取 → 无需 OCR
  ├── 4. 构建 Markdown: # 第N页 + 文本 + <img ...>
  └── 5. DeepSeek LLM 智能排版 → 最终 MD
```

### 普通 PDF 管线（MinerU API）

### 前置条件

1. Python 库：`requests`, `openai`（均已预装）
2. MinerU API Key（需设置环境变量 `MINERU_API_KEY`）
3. DeepSeek API Key（需设置环境变量 `PPT_LLM_API_KEY`）

### 标准批量转换

```bash
SKILL_DIR="C:/Users/LEGION/.workbuddy/skills/PPT&PDF-to-markdown-batch"
python "$SKILL_DIR/scripts/pdf_to_markdown_batch.py" \
  "/path/to/source_folder/" \
  -o "/path/to/output/"
```

脚本自动执行：
1. 扫描源文件夹，列出所有 `.pdf` / `.ppt` / `.pptx` 文件
2. 识别 PPT 导出 PDF（16:9/A4横向/4:3 比例），统一分流至截图管线
3. **Phase 1** — PPT/PPTX：PowerPoint COM 导出为 PDF → PyMuPDF 截图 + 文本 → LLM 排版
4. **Phase 1.5** — PPT 导出 PDF：PyMuPDF 截图 + 文本 → LLM 排版（失败时自动降级）
5. **Phase 2** — 普通 PDF：MinerU API 上传 → 解析 → 下载 ZIP → 解压
6. **智能后处理**（普通 PDF）：
   - 截取正文（去封面/目录/修订记录）
   - 图片宽度100%
   - 页码删除
   - 标题规范化
7. 所有输出自动清理人名"来潇"、版本号 V1/V2/V3
8. 源文件复制到 `original file/`
9. 输出进度摘要与最终统计

### 跳过 LLM 排版（保留其他后处理）

```bash
python "$SKILL_DIR/scripts/pdf_to_markdown_batch.py" \
  "/path/to/source_folder/" \
  -o "/path/to/output/" \
  --no-llm
```

### 完全跳过所有后处理（原始 MinerU 输出）

```bash
python "$SKILL_DIR/scripts/pdf_to_markdown_batch.py" \
  "/path/to/source_folder/" \
  -o "/path/to/output/" \
  --no-postprocess
```

## PDF 正文截取规则

| 优先级 | 识别目标 | 正则匹配 |
|--------|---------|---------|
| 1 | 摘要 | `^#{1,3}\s*摘要$`, `ABSTRACT`, `Abstract` |
| 2 | 第一章 | `前言`, `引言`, `绪论`, `第X章`, `第一章` |
| 3 | 兜底 | 跳过前3行以上空白/图片区域 |

## PPT 页眉页脚清理规则

PPT 幻灯片上方常包含公司 logo 文字、会议/比赛名称，下方常含页码。这些冗余信息在 MinerU 解析后变成重复出现的文本行。

### 自动检测（基于频率统计）

| 检测条件 | 行为 |
|---------|------|
| 非标题/非图片/非表格短行（2~50 字符） | 统计出现次数 |
| 出现 ≥ 2 次 | 判定为页眉/页脚冗余 |
| 含有正文关键词（摘要/背景/方案等） | 排除，不予删除 |

### 手动指定（推荐）

使用 `--remove-lines` 参数精确删除指定关键词的行：

```bash
python "$SKILL_DIR/scripts/pdf_to_markdown_batch.py" \
  "E:/ppt_files/" -o "E:/output/" \
  --remove-lines "中核华兴,中核创科,NIC 2026,技术交流会"
```

> 匹配规则：包含指定关键词的短行（< 60 字符）将被删除。

### 页码格式支持

| 格式 | 示例 |
|------|------|
| 中文页码 | `第5页`、`第5页/共20页` |
| 短横线页码 | `- 5 -`、`— 5 —` |
| 斜杠页码 | `1/20`、`5 / 20` |
| 括号页码 | `[1/20]`、`(5/20)` |
| 英文页码 | `Slide 5`、`Page 5` |

## PPT 章节标题重组规则

PPT 中同一个章节标题（如"项目背景"、"技术方案"）可能在多张幻灯片上重复出现。章节重组功能自动：

| 步骤 | 说明 |
|------|------|
| 检测重复 | 扫描所有 `##`/`###` 标题，找出出现 ≥ 2 次的标题 |
| 保留首次 | 保留首次出现的标题行 |
| 合并内容 | 后续相同标题下的内容合并到首次标题后，标题行不重复 |
| 非重复标题 | 不影响非重复标题的正常输出 |

**示例**：
```
原始：                   重组后：
## 项目背景              ## 项目背景
  - 内容A                  - 内容A
## 技术方案                - 内容B（来自后续同名标题）
  - 内容C                ## 技术方案
## 项目背景（重复）         - 内容C
  - 内容B                  - 内容D（来自后续同名标题）
## 技术方案（重复）
  - 内容D
```

使用 `--no-chapter-reorg` 可跳过章节重组。

## LLM 中断恢复与重试机制

批量转换多文件时，LLM API 可能因网络波动、并发限流等原因中断。脚本内置以下保护：

| 保护层级 | 触发条件 | 行为 |
|---------|---------|------|
| **第 0 层：客户端超时** | LLM API 调用超过 180s 无响应 | 触发异常进入重试，防止无限等待卡死 |
| **第 1 层：指数退避重试** | LLM API 调用失败/超时 | 自动重试 3 次，间隔 2s → 4s → 8s |
| **第 2 层：分段保留原文** | 长文档分段处理时某段失败 | 该段保留原文并输出 ⚠ 警告，其余段正常排版 |
| **第 3 层：整体回退** | 短文档 LLM 全部重试失败 | 保留原始排版，继续处理下一文件 |
| **第 4 层：批量保护** | 任一文件转换抛异常 | 记录失败但继续处理后续文件 |

> 最终汇总报告会显示 成功/失败/跳过 计数，所有文件均有明确状态。

## PPT 导出 PDF 自动识别与分流

部分 PDF 实际由 PPT 导出（保留 16:9、A4横向或 4:3 幻灯片比例），上方含公司 logo/会议名称、下方含页码。脚本自动通过**页面长宽比**判定并分流至截图管线（无需 MinerU）。

### 检测逻辑

| 检测条件 | 判定结果 |
|---------|---------|
| 前 3 页多数宽高比在 1.72~1.82 (16:9) | PPT 导出 → 截图管线（PyMuPDF 快照+文本） |
| 前 3 页多数宽高比在 1.39~1.43 (A4横向) | PPT 导出 → 截图管线（PyMuPDF 快照+文本） |
| 前 3 页多数宽高比在 1.28~1.38 (4:3) | PPT 导出 → 截图管线（PyMuPDF 快照+文本） |
| 其他比例 | 普通文档 → MinerU API 管线 |

### PPT 导出 PDF 的处理管线

```
页面比例检测 → PyMuPDF 逐页截图(1920px) + 文本提取 → 构建逐页Markdown → LLM排版
```

> 转换失败时自动降级：PyMuPDF 失败 → PyPDF2 文本提取；完全失败 → 回退到 MinerU 管线。
> 无需额外参数，完全自动化。

## PPTX 智能排版（LLM）

默认使用 **DeepSeek Chat**，排版效果优于 Kimi K2.5。可通过环境变量切换为其他模型。
所有 LLM 调用均带有 **3 次指数退避重试**（2s/4s/8s），分段处理时单段失败保留原文，不影响整体输出。

| 动作 | 示例 |
|------|------|
| 合并罗列 | `受限空间出入口\n受限空间内\n人脸识别` → `- 受限空间出入口监控与人脸识别` |
| 优化段落 | 零散短句 → 连贯段落，去除冗余换行 |
| 去页眉页脚 | 删除重复出现的公司名称/logo/会议名称/页码等冗余信息 |
| 章节整合 | 合并重复出现的章节标题下的内容 |
| 保留结构 | 标题 `##`/`###`、表格、图片引用原样保留 |
| 不添加信息 | 严格保持原文事实，仅优化排版格式 |
| **中断恢复** | 分段处理中任一段 LLM 失败 → 保留该段原文；整体失败 → 保留原始排版 |

## API Key 管理

| 用途 | 环境变量 | 说明 |
|------|---------|------|
| MinerU | `MINERU_API_KEY` | **必须**设置，否则 MinerU 管线不可用 |
| LLM 排版 | `PPT_LLM_API_BASE` + `PPT_LLM_API_KEY` + `PPT_LLM_MODEL` | **必须**设置，否则 LLM 排版不可用；默认 API Base 为 DeepSeek |

### 运行前配置（必须）

使用前必须先设置环境变量：

```bash
# 必须：设置你的 API Key
export MINERU_API_KEY="your-mineru-jwt-token"
export PPT_LLM_API_KEY="sk-your-deepseek-api-key"

# 可选：使用其他 LLM 模型（默认 DeepSeek Chat）
export PPT_LLM_API_BASE="https://api.deepseek.com/v1"
export PPT_LLM_MODEL="deepseek-chat"

# 然后直接运行
python "$SKILL_DIR/scripts/pdf_to_markdown_batch.py" \
  "E:/files/" -o "E:/output/"
```

### 使用其他模型（如 Kimi）

```bash
export PPT_LLM_API_BASE="https://api.moonshot.cn/v1"
export PPT_LLM_API_KEY="sk-your-kimi-key"
export PPT_LLM_MODEL="kimi-k2.5"
```

## Examples

**用户说：** "把 E:/files 下所有 PDF 和 PPT 转成 markdown"

```bash
SKILL_DIR="C:/Users/LEGION/.workbuddy/skills/PPT&PDF-to-markdown-batch"
python "$SKILL_DIR/scripts/pdf_to_markdown_batch.py" \
  "E:/files/" -o "E:/output/"
```

---

**用户说：** "只转换不后处理，我要原始 MinerU 输出"

```bash
python "$SKILL_DIR/scripts/pdf_to_markdown_batch.py" \
  "E:/files/" -o "E:/output/" --no-postprocess
```

---

**用户说：** "PPTX 排版我自己调，不要 LLM"

```bash
python "$SKILL_DIR/scripts/pdf_to_markdown_batch.py" \
  "E:/files/" -o "E:/output/" --no-llm
```

---

**用户说：** "这些PPT上方有公司logo和会议名称，帮我清理掉"

```bash
python "$SKILL_DIR/scripts/pdf_to_markdown_batch.py" \
  "E:/ppt_files/" -o "E:/output/" \
  --remove-lines "中核华兴,第十四届核电建造技术交流会,NIC 2026"
```

---

**用户说：** "PPT里面的章节标题重复出现了，合并一下"

```bash
# 默认已启用章节重组，无需额外参数
python "$SKILL_DIR/scripts/pdf_to_markdown_batch.py" \
  "E:/ppt_files/" -o "E:/output/"
```

## Troubleshooting

| 错误现象 | 原因 | 解决方法 |
|---------|------|---------|
| `PPT→PDF 导出失败` | PowerPoint 未安装或文件损坏 | 检查 PowerPoint 安装，手动将 PPT 另存为 PDF 后放入源目录 |
| `pages exceeds limit (200)` | PDF 超过 200 页 | 拆分 PDF 为多个小文件 |
| `LLM 排版失败` | DeepSeek API 网络或额度问题 | 自动重试 3 次（2s/4s/8s 退避）；全部失败则保留原文继续 |
| `LLM 分段排版部分失败` | 某段内容过长或 API 异常 | 该段保留原文，其余段正常排版，不影响整体输出 |
| `轮询超时` | 大文件 MinerU 解析慢 | 默认 30 分钟超时 |
| `openai not found` | 缺少依赖 | `pip install openai` |
| `PPT导出PDF截图失败` | 文件损坏或加密 | 自动降级回退到 MinerU 管线 |
| `人名/版本号未清除` | 非标准格式或嵌入太深 | 手动重命名源文件后重新转换 |
| `FileNotFoundError on output dir` | Python 路径中的 `《》""` 等特殊字符导致 `os.makedirs` 失败，或沙箱限制写入外部目录 | 1) 优先使用 Python venv + `PYTHONUTF8=1` 环境变量 2) 如需跨目录移动结果，使用 PowerShell `Copy-Item -LiteralPath`（`-LiteralPath` 避免 `[]` 被解释为通配符） |
| `Permission denied` 写入外部目录 | WorkBuddy 沙箱安全策略 | 使用 PowerShell `Copy-Item` 并确认权限弹窗，或先输出到 workspace 目录再手动移动 |

## 注意事项

1. **PPT 依赖 PowerPoint**：PPT/PPTX 需本机安装 Microsoft PowerPoint 以导出 PDF。若未安装，请先手动将 PPT 另存为 PDF 再放入源目录。
2. **批次限制**：普通 PDF 单次 ≤ 50 个文件（MinerU 限制），PPT/PPTX/PPT导出PDF 无此限制。
3. **文件大小**：单文件 ≤ 200MB，页数 ≤ 200 页。
4. **名称清理**：自动删除输出中的人名"来潇"和版本号 V1/V2/V3 等，保留文件核心名称。
5. **日期提取**：优先从文件名提取日期（支持开头和嵌入位置如"方案-2025.7.29"），无日期时使用文件修改时间。
6. **PDF 正文截取**：仅对普通 PDF 生效；PPT/PPTX/PPT导出PDF 由 LLM 统一排版。
7. **输出覆盖**：同名子文件夹会被覆盖。
8. **LLM 中断保护**：所有 LLM 调用均带有 3 次指数退避重试（2s/4s/8s），分段模式单段失败保留原文继续，确保批量转换不因个别 API 问题中断。
