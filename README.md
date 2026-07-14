# PPT & PDF to Markdown Batch Converter

> 双管线批量文档转 Markdown 工具 — 同时支持 **PowerPoint (PPT/PPTX)** 和 **PDF** 文件

![Python](https://img.shields.io/badge/Python-3.8%2B-blue)
![License](https://img.shields.io/badge/license-MIT-green)

---

## ✨ 功能特性

| 特性 | 说明 |
|------|------|
| 🎯 **双管线架构** | PPT/PPTX 走截图管线，普通 PDF 走 MinerU 解析管线 |
| 🖼️ **PPT 截图** | PowerPoint COM 导出 PDF → PyMuPDF 逐页高清截图 (1920px) |
| 🤖 **LLM 智能排版** | 默认 DeepSeek Chat 去冗余、合并段落、优化表格，带 3 次重试保护 |
| 🔍 **PPT 导出 PDF 自动识别** | 通过页面比例(16:9/A4/4:3)自动分流 |
| 🧹 **自动清理** | 去人名、版本号、页码、页眉页脚 |
| 📅 **智能命名** | 从文件名提取日期，输出 `YYYY-MM-DD 标题/` 结构 |
| 🔤 **EasyOCR 降级** | PyMuPDF 文字提取不足时自动用 EasyOCR 识别截图，图文混合场景必备 |
| 🔑 **唯一编码集成** | 自动生成 8 位唯一编码，重命名源文件并插入 ES 引用链接 |
| 🛡️ **降级保护** | 截图失败→回退 MinerU；LLM 超时→保留原文；单文件失败不阻塞后续 |

## 📋 支持格式

| 格式 | 处理方式 | 特点 |
|------|---------|------|
| **PPT / PPTX** | COM → PDF → 截图 + 文本→EasyOCR降级 → LLM 排版 | 保留幻灯片视觉效果 + 双保险文字提取 |
| **PPT 导出的 PDF** | 自动检测比例 → 截图 + 文本 → LLM 排版 | 与 PPT 同管线 |
| **普通 PDF** | MinerU API 解析 → 后处理 | 高精度文本/表格/公式提取 |

## 🚀 快速开始

### 前置条件

```bash
# Python 依赖（均已预装）
pip install requests openai PyMuPDF PyPDF2 python-docx pythoncom easyocr
```

### 设置环境变量

```bash
# 必须：MinerU API Key（用于普通 PDF 解析）
export MINERU_API_KEY='your-mineru-jwt-token'

# 必须：LLM API Key（用于 PPT 智能排版）
export PPT_LLM_API_KEY='sk-your-deepseek-api-key'

# 可选：切换 LLM 模型
export PPT_LLM_API_BASE="https://api.deepseek.com/v1"
export PPT_LLM_MODEL="deepseek-chat"
```

### 基本使用

```bash
python scripts/pdf_to_markdown_batch.py \
  "/path/to/your/files/" \
  -o "/path/to/output/"
```

### 常用参数

```bash
# 跳过 LLM 排版（保留截图 + 文本）
python scripts/pdf_to_markdown_batch.py ... --no-llm

# 跳过所有后处理（原始 MinerU 输出）
python scripts/pdf_to_markdown_batch.py ... --no-postprocess

# 清理特定页眉页脚
python scripts/pdf_to_markdown_batch.py ... --remove-lines "公司名称,会议名称,页码"

# 跳过章节去重重组
python scripts/pdf_to_markdown_batch.py ... --no-chapter-reorg
```

## 🔑 唯一编码集成

每个文档转换前自动生成 **8 位唯一编码**（基于计算机时间精确到秒）：

1. **源文件重命名**：核心名.pdf → 核心名【UID】.pdf
2. **ES 链接**：在 .md 文件顶部插入引用行

`
> 原始文件: [006T9600](es:006T9600) | 核心名【006T9600】.pdf
`

> 文件夹和 .md 文件名不含 UID、日期、版本号。

前置依赖：生成唯一编码 技能。

## 📁 输出结构

```
output_dir/
├── 2025-09-30 项目方案/
│   ├── 2025-09-30 项目方案.md
│   ├── attachments/
│   │   ├── slide_01.png
│   │   └── ...
└── ...
```

## 🔧 技术细节

### PPT 管线
```
PPTX → PowerPoint COM → PDF → PyMuPDF 逐页截图(1920px) → 文本提取(不足→EasyOCR降级) → LLM 排版
```

### 普通 PDF 管线
```
PDF → MinerU API → ZIP(MD+图片+JSON) → 正文截取 → 标题规范化 → 图片宽度100% → 页码删除
```

### LLM 中断保护
| 层级 | 保护机制 |
|------|---------|
| 指数退避 | 失败自动重试 3 次 (2s→4s→8s) |
| 分段保留 | 长文档某段失败→保留原文，其余正常 |
| 整体回退 | 整文件 LLM 失败→保留原始排版继续下一个 |
| 批量保护 | 任一文件异常→记录失败继续处理 |

## 📄 License

MIT
