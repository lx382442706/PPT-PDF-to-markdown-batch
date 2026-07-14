#!/usr/bin/env python
"""
MinerU 文档 → Markdown 批量转换脚本
基于 MinerU 精准解析 API (/api/v4/file-urls/batch + /api/v4/extract-results/batch)
支持格式: PDF, PPT, PPTX
支持批量上传、异步轮询、ZIP下载解压、Markdown提取、智能后处理
"""

import os
import sys
import re
import time
import json
import shutil
import zipfile
import argparse
import requests
import urllib3
urllib3.disable_warnings()
from pathlib import Path
from datetime import datetime
from openai import OpenAI
from PyPDF2 import PdfReader
import fitz  # PyMuPDF for PDF page rendering
import pythoncom


# --- EasyOCR fallback (lazy loaded) ---
_EASYOCR_READER = None

def _get_easyocr_reader():
    global _EASYOCR_READER
    if _EASYOCR_READER is None:
        try:
            import easyocr
            print('  [EasyOCR] initializing reader (first load ~2GB model)...', flush=True)
            _EASYOCR_READER = easyocr.Reader(['ch_sim', 'en'], gpu=True, verbose=False)
            print('  [EasyOCR] ready', flush=True)
        except Exception as e:
            print(f'  [EasyOCR] init failed: {e}', flush=True)
            _EASYOCR_READER = '__UNAVAILABLE__'
    return _EASYOCR_READER if _EASYOCR_READER != '__UNAVAILABLE__' else None

def _extract_text_via_easyocr(img_path):
    reader = _get_easyocr_reader()
    if reader is None:
        return None
    try:
        results = reader.readtext(img_path, detail=0, paragraph=True)
        text = '\n'.join(results).strip()
        return text if text else None
    except Exception as e:
        print(f'  [EasyOCR] recognition failed: {e}', flush=True)
        return None
def _count_substantive_text(slides_text):
    from collections import Counter
    import re as _re
    all_lines = []
    for _pn, txt in slides_text:
        if not txt: continue
        for line in txt.split(chr(10)):
            line = line.strip()
            if not line: continue
            if _re.match(r"^\d{1,4}$", line): continue
            if _re.match(r"^第\d+页(?:/共\d+页)?$", line): continue
            if _re.match(r"^(?:Page|Slide)\s+\d+\s+of\s+\d+$", line, _re.I): continue
            all_lines.append(line)
    counter = Counter(all_lines)
    num_pages = max(len(slides_text), 1)
    threshold = max(num_pages * 0.6, 2)
    filtered = [l for l in all_lines if counter[l] < threshold or len(l) > 30]
    return len(''.join(filtered))



# ─── 跨平台安全目录创建 ────────────────────────────
def safe_makedirs(path, exist_ok=True):
    """
    安全创建目录，兼容跨平台路径。
    """
    os.makedirs(str(path), exist_ok=exist_ok)


# ─── 配置 ───────────────────────────────────────────
API_BASE = "https://mineru.net"
BATCH_URL = f"{API_BASE}/api/v4/file-urls/batch"
BATCH_RESULT_URL = f"{API_BASE}/api/v4/extract-results/batch"
BATCH_SIZE = 50
POLL_INTERVAL = 10
MAX_POLL_TIME = 1800
UPLOAD_TIMEOUT = 600

DEFAULT_API_KEY = os.environ.get("MINERU_API_KEY", "")
if not DEFAULT_API_KEY:
    print("⚠️  警告: 未设置 MINERU_API_KEY 环境变量，MinerU 管线将无法使用")
    print("   export MINERU_API_KEY='your-mineru-jwt-token'")

# Default LLM config (DeepSeek) — overridable via PPT_LLM_* env vars
DEFAULT_LLM_API_BASE = "https://api.deepseek.com/v1"
DEFAULT_LLM_API_KEY = os.environ.get("PPT_LLM_API_KEY", "")
if not DEFAULT_LLM_API_KEY:
    print("⚠️  警告: 未设置 PPT_LLM_API_KEY 环境变量，LLM 排版功能将无法使用")
    print("   export PPT_LLM_API_KEY='your-deepseek-api-key'")
DEFAULT_LLM_MODEL = "deepseek-chat"

# Legacy Kimi config (kept for fallback reference, not used by default)
KIMI_API_BASE = "https://api.moonshot.cn/v1"
KIMI_API_KEY = os.environ.get("KIMI_API_KEY", "")
# if you want to use Kimi as LLM provider: export PPT_LLM_API_KEY=$KIMI_API_KEY
KIMI_MODEL = "kimi-k2.5"


def get_llm_config():
    """Resolve LLM config: env vars take priority, fallback to DeepSeek defaults."""
    return {
        "api_base": os.environ.get("PPT_LLM_API_BASE", DEFAULT_LLM_API_BASE),
        "api_key": os.environ.get("PPT_LLM_API_KEY", DEFAULT_LLM_API_KEY),
        "model": os.environ.get("PPT_LLM_MODEL", DEFAULT_LLM_MODEL),
    }


# ─── LLM 调用增强 ────────────────────────────────────
LLM_MAX_RETRIES = 3
LLM_RETRY_BACKOFF = [2, 4, 8]  # 秒


def call_llm_with_retry(messages, system_prompt, max_tokens=16000, temperature=1.0):
    """
    带重试机制的 LLM 调用。
    指数退避重试 3 次，失败后返回 None（由调用方决定回退策略）。
    """
    llm = get_llm_config()
    client = OpenAI(api_key=llm["api_key"], base_url=llm["api_base"])

    for attempt in range(LLM_MAX_RETRIES):
        try:
            resp = client.chat.completions.create(
                model=llm["model"],
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": messages}
                ],
                temperature=temperature,
                max_tokens=max_tokens,
                timeout=180,
            )
            return resp.choices[0].message.content
        except Exception as e:
            err_msg = str(e)
            if attempt < LLM_MAX_RETRIES - 1:
                wait = LLM_RETRY_BACKOFF[attempt]
                print(f"(重试 {attempt+1}/{LLM_MAX_RETRIES}, {wait}s后) ", end="", flush=True)
                time.sleep(wait)
            else:
                print(f"失败: {err_msg}")
                return None


def get_api_key():
    return os.environ.get("MINERU_API_KEY", DEFAULT_API_KEY)


def get_headers():
    return {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {get_api_key()}",
    }


def collect_files(source_dir):
    """扫描目录，收集所有 PDF/PPT/PPTX 文件（去重）"""
    src = Path(source_dir)
    if not src.exists():
        print(f"[ERROR] 源目录不存在: {source_dir}")
        sys.exit(1)
    exts = {".pdf", ".ppt", ".pptx"}
    files = []
    for ext in exts:
        files.extend(src.glob(f"*{ext}"))
        files.extend(src.glob(f"*{ext.upper()}"))
    files = sorted(set(files), key=lambda p: p.name)
    if not files:
        print(f"[WARN] 未找到支持的文件 (PDF/PPT/PPTX): {source_dir}")
        sys.exit(0)
    return files


def chunk_list(lst, n):
    for i in range(0, len(lst), n):
        yield lst[i : i + n]


def submit_batch(files_chunk, model_version, enable_formula, enable_table, language):
    payload = {
        "files": [{"name": f.name, "data_id": f.stem.strip()[:120]} for f in files_chunk],
        "model_version": model_version,
        "enable_formula": enable_formula,
        "enable_table": enable_table,
        "language": language,
    }
    resp = requests.post(BATCH_URL, headers=get_headers(), json=payload, timeout=30, verify=False)
    data = resp.json()
    if data.get("code") != 0:
        raise RuntimeError(f"申请上传链接失败: {data.get('msg')} (trace_id={data.get('trace_id')})")
    return data["data"]["batch_id"], data["data"]["file_urls"]


def upload_files(file_paths, file_urls):
    for fp, url in zip(file_paths, file_urls):
        fsize_mb = fp.stat().st_size / 1024 / 1024
        print(f"  [上传] {fp.name} ({fsize_mb:.1f} MB) ... ", end="", flush=True)
        with open(fp, "rb") as f:
            resp = requests.put(url, data=f, timeout=UPLOAD_TIMEOUT, verify=False)
            if resp.status_code == 200:
                print("OK")
            else:
                print(f"FAIL (HTTP {resp.status_code})")


def poll_batch_results(batch_id):
    url = f"{BATCH_RESULT_URL}/{batch_id}"
    start = time.time()
    last_print = 0
    while True:
        elapsed = time.time() - start
        if elapsed > MAX_POLL_TIME:
            raise TimeoutError(f"轮询超时 ({MAX_POLL_TIME}s)")
        resp = requests.get(url, headers=get_headers(), timeout=30, verify=False)
        data = resp.json()
        if data.get("code") != 0:
            raise RuntimeError(f"查询失败: {data.get('msg')}")
        results = data["data"].get("extract_result", [])
        states = {r.get("state") for r in results}
        if elapsed - last_print >= 30 or set(states) == {"done"} or "failed" in states:
            done = sum(1 for r in results if r.get("state") == "done")
            running = sum(1 for r in results if r.get("state") == "running")
            pending = sum(1 for r in results if r.get("state") in ("pending", "waiting-file"))
            failed = sum(1 for r in results if r.get("state") == "failed")
            converting = sum(1 for r in results if r.get("state") == "converting")
            print(f"  [进度] done={done} running={running} pending={pending} failed={failed} converting={converting} | {elapsed:.0f}s")
            last_print = elapsed
        if all(s in ("done", "failed") for s in states):
            return results
        time.sleep(POLL_INTERVAL)


# ═══════════════════════════════════════════════════════
#  日期提取与文件夹命名
# ═══════════════════════════════════════════════════════

def extract_date_from_filename(filename):
    """
    从文件名中提取日期，返回 (date_str_YYYY-MM-DD, core_name) 或 (None, original_stem)
    支持格式: YYYY-MM-DD、YYYYMMDD、YYYY.MM.DD、YYYY.M.DD、YYYY年MM月DD日、YYMMDD
    支持日期位于文件名开头或任意位置（如 "报告名称-2025.7.29"、"2025年7月29日方案"）
    """
    name = Path(filename).stem.strip()

    # 模式1: 日期开头 + 空格 + 核心名称
    #   "2025-09-30 核电混凝土智能高频振捣棒研发实施方案"
    #   "2025.09.30 报告名称"
    m = re.match(r'^(\d{4}[-.]\d{1,2}[-.]\d{1,2})\s+(.+)$', name)
    if m:
        date_part = m.group(1).replace('.', '-')
        core_name = m.group(2).strip()
        try:
            dt = datetime.strptime(date_part, '%Y-%m-%d')
            return dt.strftime('%Y-%m-%d'), core_name
        except ValueError:
            pass

    # 模式2: YYYYMMDD 开头 + 核心名称 (可能无空格)
    #   "20250930报告名称" or "20250930 报告名称"
    m = re.match(r'^(\d{8})\s*(.+)$', name)
    if m:
        date_part = m.group(1)
        core_name = m.group(2).strip()
        try:
            dt = datetime.strptime(date_part, '%Y%m%d')
            return dt.strftime('%Y-%m-%d'), core_name
        except ValueError:
            pass

    # 模式3: 中文日期格式 "YYYY年MM月DD日 核心名称"
    m = re.match(r'^(\d{4})年(\d{1,2})月(\d{1,2})日\s*(.+)$', name)
    if m:
        y, mo, d = m.groups()
        core_name = m.group(4).strip()
        try:
            return f"{int(y):04d}-{int(mo):02d}-{int(d):02d}", core_name
        except ValueError:
            pass

    # 模式4: YYMMDD 开头 (6位)
    #   "250930 报告名称"
    m = re.match(r'^(\d{6})\s+(.+)$', name)
    if m:
        date_part = m.group(1)
        core_name = m.group(2).strip()
        try:
            dt = datetime.strptime(date_part, '%y%m%d')
            # YY → 20YY (假设2000年后)
            return dt.strftime('%Y-%m-%d'), core_name
        except ValueError:
            pass

    # 纯日期文件名（无核心名称）
    pure_patterns = [
        (r'^(\d{4}[-.]\d{2}[-.]\d{2})$', '%Y-%m-%d'),
        (r'^(\d{8})$', '%Y%m%d'),
        (r'^(\d{6})$', '%y%m%d'),
    ]
    for pat, fmt in pure_patterns:
        m = re.match(pat, name)
        if m:
            date_part = m.group(1)
            try:
                dt = datetime.strptime(date_part.replace('.', '-'), fmt)
                return dt.strftime('%Y-%m-%d'), name
            except ValueError:
                continue

    # 中文纯日期
    m = re.match(r'^(\d{4})年(\d{1,2})月(\d{1,2})日$', name)
    if m:
        y, mo, d = m.groups()
        try:
            return f"{int(y):04d}-{int(mo):02d}-{int(d):02d}", name
        except ValueError:
            pass

    # ── 嵌入日期模式：日期出现在文件名任意位置 ──
    # 模式5: YYYY.M.DD / YYYY-MM-DD 在文件名任意位置
    #   "来潇-振捣棒技术方案研讨会-2025.7.29" → date=2025-07-29, core=来潇-振捣棒技术方案研讨会
    #   "报告-2025.03.15-终版"
    m = re.search(r'(\d{4})[-.](\d{1,2})[-.](\d{1,2})', name)
    if m:
        y, mo, d = m.groups()
        try:
            date_str = f"{int(y):04d}-{int(mo):02d}-{int(d):02d}"
            # 从文件名中移除日期部分，剩下的为核心名
            core_name = re.sub(r'[-_\s]*\d{4}[-.]\d{1,2}[-.]\d{1,2}[-_\s]*', '', name).strip()
            core_name = re.sub(r'[-_]{2,}', '-', core_name).strip('-').strip()
            return date_str, core_name
        except ValueError:
            pass

    # 模式6: YYYYMMDD (8位连写) 在文件名任意位置
    #   "报告_20250729_终版"
    m = re.search(r'(\d{8})', name)
    if m:
        date_part = m.group(1)
        try:
            dt = datetime.strptime(date_part, '%Y%m%d')
            core_name = re.sub(r'[-_\s]*\d{8}[-_\s]*', '', name).strip()
            core_name = re.sub(r'[-_]{2,}', '-', core_name).strip('-').strip()
            return dt.strftime('%Y-%m-%d'), core_name
        except ValueError:
            pass

    # 模式7: 中文日期在文件名任意位置
    #   "振捣棒技术方案研讨会2025年7月29日"
    m = re.search(r'(\d{4})年(\d{1,2})月(\d{1,2})日', name)
    if m:
        y, mo, d = m.groups()
        try:
            date_str = f"{int(y):04d}-{int(mo):02d}-{int(d):02d}"
            core_name = re.sub(r'[-_\s]*\d{4}年\d{1,2}月\d{1,2}日[-_\s]*', '', name).strip()
            core_name = re.sub(r'[-_]{2,}', '-', core_name).strip('-').strip()
            return date_str, core_name
        except ValueError:
            pass

    return None, name


def clean_output_name(name):
    """
    清理输出名称：删除人名、版本号等冗余信息。
    1. 删除人名 "来潇" 及其变体（来潇-、来潇_ 等）
    2. 删除版本信息（V1、V2、-V3、_V1、版本1 等）
    3. 清理多余空白和分隔符
    """
    # 1. 删除人名 "来潇" 及周围的连字符/空格
    name = re.sub(r'来潇[-_\s]*', '', name)

    # 2. 删除版本信息: -V1, -V2, _V3, (V1), V1, 版本1 等
    name = re.sub(r'[-_\s]*[Vv]\d+(\.\d+)*', '', name)
    name = re.sub(r'[-_\s]*版本\d+', '', name)
    name = re.sub(r'\s*[\[\(][Vv]\d+[\]\)]', '', name)

    # 3. 删/替换可能导致 Windows 目录创建失败的特殊字符
    # 《 》 " " → 替换为安全字符或删除
    name = name.replace('\u300a', '[').replace('\u300b', ']')  # 《》→ []
    name = name.replace('\u201c', '').replace('\u201d', '')     # "" → 删除

    # 4. 清理多余空白和分隔符
    name = re.sub(r'\s+', ' ', name)
    name = re.sub(r'[-_]{2,}', '-', name)
    name = re.sub(r'^\s*[-_]\s*', '', name)
    name = re.sub(r'\s*[-_]\s*$', '', name)
    name = name.strip()

    return name if name else "未命名"


def format_output_dirname(source_path, stem):
    """
    生成输出子文件夹名称：YYYY-MM-DD 文件核心名
    优先从文件名提取日期，无日期时使用文件最后修改日期
    应用 clean_output_name 去人名和版本号
    """
    if source_path:
        date_str, core_name = extract_date_from_filename(Path(source_path).name)
        core_name = clean_output_name(core_name)
        if date_str:
            return f"{date_str} {core_name}"
        # 使用文件最后修改日期
        try:
            mtime = Path(source_path).stat().st_mtime
            date_str = datetime.fromtimestamp(mtime).strftime('%Y-%m-%d')
        except Exception:
            date_str = datetime.now().strftime('%Y-%m-%d')
        stem_clean = clean_output_name(stem)
        return f"{date_str} {stem_clean}"
    return clean_output_name(stem)

# ═══════════════════════════════════════════════════════
#  下载与解压
# ═══════════════════════════════════════════════════════

def download_and_extract(result_item, output_dir, source_path=None):
    """下载 ZIP 并解压到独立子文件夹，仅保留 .md + attachments/"""
    file_name = result_item.get("file_name", "unknown")
    full_zip_url = result_item.get("full_zip_url", "")
    if result_item.get("state") != "done":
        print(f"  [SKIP] {file_name}: state={result_item.get('state')} err={result_item.get('err_msg', '')}")
        return None
    if not full_zip_url:
        print(f"  [SKIP] {file_name}: 无下载链接")
        return None

    stem = Path(file_name).stem.strip()
    dirname = format_output_dirname(source_path, stem)
    sub_dir = Path(output_dir) / dirname
    safe_makedirs(str(sub_dir), exist_ok=True)

    print(f"  [下载] {file_name} ... ", end="", flush=True)
    zip_resp = requests.get(full_zip_url, timeout=300, verify=False)
    if zip_resp.status_code != 200:
        print(f"失败 HTTP {zip_resp.status_code}")
        return None

    tmp_dir = Path(output_dir) / f".tmp_{stem}"
    safe_makedirs(str(tmp_dir), exist_ok=True)

    zip_path = tmp_dir / "result.zip"
    zip_path.write_bytes(zip_resp.content)
    zsize_mb = len(zip_resp.content) / 1024 / 1024
    print(f"OK ({zsize_mb:.1f} MB)")

    md_files = []
    with zipfile.ZipFile(zip_path, "r") as zf:
        for info in zf.infolist():
            target = tmp_dir / info.filename
            if info.is_dir():
                safe_makedirs(str(target))
            else:
                safe_makedirs(str(target.parent))
                target.write_bytes(zf.read(info))
                if target.suffix.lower() == ".md":
                    md_files.append(target)

    if md_files:
        root_mds = [m for m in md_files if m.parent == tmp_dir]
        main_md = (root_mds or md_files)[0]
        target_md = sub_dir / f"{dirname}.md"
        content = main_md.read_text(encoding="utf-8")
        content = content.replace("](images/", "](attachments/")
        content = content.replace('src="images/', 'src="attachments/')
        target_md.write_text(content, encoding="utf-8")
        print(f"  [MD] {target_md.name}")

    images_src = tmp_dir / "images"
    if images_src.exists() and images_src.is_dir():
        images_dst = sub_dir / "attachments"
        if images_dst.exists():
            shutil.rmtree(images_dst)
        images_src.rename(images_dst)

    shutil.rmtree(tmp_dir, ignore_errors=True)
    return sub_dir / f"{dirname}.md" if md_files else None


# ═══════════════════════════════════════════════════════
#  后处理模块
# ═══════════════════════════════════════════════════════

def is_ppt_file(filename):
    return Path(filename).suffix.lower() in (".ppt", ".pptx")


def is_pdf_file(filename):
    return Path(filename).suffix.lower() == ".pdf"


def detect_ppt_like_pdf(pdf_path):
    """
    通过页面长宽比判断 PDF 是否由 PPT 导出：
    - 读取前 3 页尺寸
    - 若多数页宽高比在 16:9 (1.72-1.82) 或 4:3 (1.28-1.38) 范围 → 判定为 PPT 导出
    - 用于自动启用 PPT 清洗策略（页眉页脚清理、章节重组等）
    
    返回: True（PPT-like） 或 False（普通文档PDF）
    """
    try:
        reader = PdfReader(pdf_path)
        # 读取前 3 页（或全部，取较小值）
        pages_to_check = min(3, len(reader.pages))
        if pages_to_check == 0:
            return False
        
        ppt_like_count = 0
        for i in range(pages_to_check):
            page = reader.pages[i]
            # 获取页面尺寸（单位：点，1pt = 1/72英寸）
            width = float(page.mediabox.width)
            height = float(page.mediabox.height)
            
            if width <= 0 or height <= 0:
                continue
            
            # 计算宽高比（始终 ≥ 1，即横向/纵向统一处理）
            # PPT 通常为横向，但保险起见取 max/min
            ar = max(width, height) / min(width, height)
            
            # 16:9 ≈ 1.78（允许 ±0.04 浮动）
            # A4横向 ≈ 1.41（PPT 导出为 A4 横向时常见）
            # 4:3 ≈ 1.33（允许 ±0.05 浮动）
            if 1.72 <= ar <= 1.82:       # 16:9
                ppt_like_count += 1
            elif 1.39 <= ar <= 1.43:     # A4横向 (PPT导出)
                ppt_like_count += 1
            elif 1.28 <= ar <= 1.38:     # 4:3
                ppt_like_count += 1
        
        # 超过半数页面匹配 → PPT-like
        is_ppt_like = ppt_like_count > pages_to_check / 2
        
        if is_ppt_like:
            print(f"  [检测] {Path(pdf_path).name}: 页面比例符合PPT特征 (16:9/A4横向/4:3)，启用PPT清洗策略")
        else:
            print(f"  [检测] {Path(pdf_path).name}: 普通文档PDF，标准处理")
        
        return is_ppt_like
    except Exception as e:
        # 读取失败时默认为普通PDF，安全起见不误清
        print(f"  [检测] {Path(pdf_path).name}: 无法读取页面尺寸 ({e})，按普通PDF处理")
        return False


def trim_pdf_front_matter(content):
    """
    PDF 正文智能截取：
    - 有"摘要" → 删除摘要前全部文字（含"摘要"行保留）
    - 无"摘要" → 删除正文前封面/目录等前端内容
    """
    # 尝试找"摘要"/"Abstract"作为正文起点
    abstract_patterns = [
        r'^#{1,3}\s*摘\s*要\s*$',
        r'^#{1,3}\s*ABSTRACT\s*$',
        r'^#{1,3}\s*Abstract\s*$',
        r'^\*\*摘\s*要\*\*',
        r'^\*\*ABSTRACT\*\*',
    ]
    for pat in abstract_patterns:
        m = re.search(pat, content, re.MULTILINE | re.IGNORECASE)
        if m:
            content = content[m.start():]
            print("  [后处理] 从「摘要」处截取正文")
            return content

    # 无摘要 → 检查是否有"目录"
    toc_patterns = [
        r'^#{1,3}\s*.*(?:目录|目次)\s*$',
        r'^\*\*.*(?:目录|目次).*\*\*\s*$',
    ]
    for pat in toc_patterns:
        m = re.search(pat, content, re.MULTILINE)
        if m:
            content = content[m.start():]
            print("  [后处理] 从「目录」处截取（保留目录及正文）")
            return content

    # 无目录 → 找正文起始标志
    body_markers = [
        r'^#{1,3}\s*(前言|引言|绪论|第[一二三四五六七八九十\d]+章|第一章)',
        r'^\*\*(前言|引言|绪论)\*\*',
        r'^#{1,3}\s*\d+[\.\s、]',
    ]
    for pat in body_markers:
        m = re.search(pat, content, re.MULTILINE)
        if m:
            content = content[m.start():]
            print("  [后处理] 从正文第一章处截取")
            return content

    # 兜底：跳过前3行空行后的内容
    lines = content.split('\n')
    start = 0
    blank_count = 0
    for i, line in enumerate(lines):
        stripped = line.strip()
        if not stripped or stripped.startswith('![') or stripped.startswith('<img'):
            blank_count += 1
        else:
            if blank_count >= 3:
                start = i
                break
            blank_count = 0
    if start > 0:
        content = '\n'.join(lines[start:])
        print("  [后处理] 跳过前端空白/图片区域")
    return content


def trim_before_toc(content):
    """
    如果 MD 文件包含「目录」章节标题，删除目录之前的所有内容。
    保留「目录」及其后的全部正文。
    支持格式：# 目录、# 工艺卡目录、# 目次、**目录** 等变体。
    """
    toc_patterns = [
        r'^#{1,3}\s*.*(?:目录|目次)\s*$',
        r'^\*\*.*(?:目录|目次).*\*\*\s*$',
    ]
    for pat in toc_patterns:
        m = re.search(pat, content, re.MULTILINE)
        if m:
            content = content[m.start():]
            print("  [后处理] 从「目录」处截取（保留目录及正文）")
            return content
    return content


def fix_image_width(content):
    """
    将 Markdown 图片语法 ![](path) 改为 HTML <img width="100%">
    确保图片宽度与页面同宽
    """
    # 匹配 ![](attachments/xxx) 和 ![alt](attachments/xxx)
    def replace_img(m):
        alt = m.group(1) or ""
        path = m.group(2)
        return f'<img src="{path}" width="100%">'

    content = re.sub(r'!\[([^\]]*)\]\(([^)]+)\)', replace_img, content)

    # 也处理已有的 <img> 标签，添加 width
    def add_width(m):
        tag = m.group(0)
        if 'width=' not in tag:
            tag = tag.replace('<img ', '<img width="100%" ')
        return tag
    content = re.sub(r'<img\s[^>]*>', add_width, content)

    return content


# ═══════════════════════════════════════════════════════
#  PPTX 快照转换（PowerPoint COM + python-pptx）
# ═══════════════════════════════════════════════════════

def convert_pptx_via_pdf(pptx_path, output_dir):
    """
    PPT/PPTX → PDF 导出 → 截图管线：
    1. PowerPoint COM 将 PPT/PPTX 导出为 PDF
    2. 将 PDF 送入 convert_ppt_pdf_via_slides 截图+文本提取管线
    """
    import tempfile
    from win32com.client import Dispatch

    src = Path(pptx_path)
    ppt_name = src.stem.strip()

    # ── 1. 导出 PDF ──
    temp_pdf = Path(tempfile.gettempdir()) / f"_wb_ppt_{src.stem}.pdf"
    print(f"  [PPT→PDF] {src.name} ... ", end="", flush=True)

    pythoncom.CoInitialize()
    ppt_app = None
    presentation = None
    try:
        ppt_app = Dispatch("PowerPoint.Application")
        ppt_app.Visible = True
        ppt_app.DisplayAlerts = True
        presentation = ppt_app.Presentations.Open(str(src.resolve()), ReadOnly=True, WithWindow=False)
        # 32 = ppSaveAsPDF
        presentation.SaveAs(str(temp_pdf), 32)
        presentation.Close()
        presentation = None
        ppt_app.Quit()
        ppt_app = None
        pdf_size_mb = temp_pdf.stat().st_size / 1024 / 1024
        print(f"OK ({pdf_size_mb:.1f} MB)")
    except Exception as e:
        print(f"失败: {e}")
        return None
    finally:
        if presentation:
            try:
                presentation.Close()
            except:
                pass
        if ppt_app:
            try:
                ppt_app.Quit()
            except:
                pass
        try:
            pythoncom.CoUninitialize()
        except:
            pass

    # ── 2. 送入截图管线（传递原始 pptx 路径用于命名）──
    md_path = convert_ppt_pdf_via_slides(temp_pdf, output_dir, source_path=pptx_path)

    # ── 3. 清理临时 PDF ──
    try:
        temp_pdf.unlink()
    except Exception:
        pass

    return md_path


def convert_ppt_pdf_via_slides(pdf_path, output_dir, source_path=None):
    """
    PPT 导出 PDF → Markdown（快照 + 文本提取方式）：
    1. 使用 PyMuPDF (fitz) 渲染每页为 PNG 截图
    2. 使用 PyMuPDF 提取每页文本（原生文本提取，准确率远高于 OCR）
    3. 按页构建 Markdown：文本 + 截图
    4. 输出到指定子文件夹

    source_path: 可选原始源文件路径（如 PPTX），用于文件夹命名
    """
    src = Path(pdf_path)
    # 如果提供了 source_path，用它来命名（保留原始文件名信息）
    naming_path = source_path if source_path else pdf_path
    naming_path_obj = Path(naming_path)
    display_name = naming_path_obj.stem.strip()
    dirname = format_output_dirname(naming_path, display_name)
    sub_dir = Path(output_dir) / dirname
    safe_makedirs(str(sub_dir), exist_ok=True)
    att_dir = sub_dir / "attachments"
    safe_makedirs(str(att_dir), exist_ok=True)

    md_path = sub_dir / f"{dirname}.md"

    # ── 1. 打开 PDF ──
    print(f"  [PDF打开] {display_name} ... ", end="", flush=True)
    doc = None
    try:
        doc = fitz.open(str(src))
        total_pages = doc.page_count
        print(f"OK ({total_pages} 页)")
    except Exception as e:
        print(f"失败: {e}")
        # 降级：尝试 PyPDF2
        try:
            from PyPDF2 import PdfReader
            reader = PdfReader(str(src))
            total_pages = len(reader.pages)
            print(f"  降级成功 ({total_pages} 页, PyPDF2)")
            doc = reader  # 注意：PyPDF2 无法渲染截图
        except Exception as e2:
            print(f"  完全失败: {e2}")
            return None

    use_fitz = isinstance(doc, fitz.Document)

    # ── 2. 逐页截图 + 文本提取 ──
    slides_images = []
    slides_text = []
    slides_failed = 0

    for page_num in range(total_pages):
        # ── 截图 ──
        img_name = f"slide_{page_num+1:02d}.png"
        img_path = str(att_dir / img_name)

        if use_fitz:
            page = doc[page_num]
            try:
                # 高分辨率渲染 (1920px 宽, 保持比例)
                scale = 1920.0 / page.rect.width
                mat = fitz.Matrix(scale, scale)
                pix = page.get_pixmap(matrix=mat)
                pix.save(img_path)
                slides_images.append(img_name)
            except Exception:
                # 降级：低分辨率重试
                try:
                    scale = 1280.0 / page.rect.width
                    mat = fitz.Matrix(scale, scale)
                    pix = page.get_pixmap(matrix=mat)
                    pix.save(img_path)
                    slides_images.append(img_name)
                except Exception:
                    slides_failed += 1

            # ── 文本提取（PyMuPDF）──
            try:
                text = page.get_text("text")
                slides_text.append((page_num + 1, (text or "").strip()))
            except Exception:
                slides_text.append((page_num + 1, ""))
        else:
            # PyPDF2 降级模式：仅提取文本，无法截图
            try:
                text = doc.pages[page_num].extract_text() or ""
                slides_text.append((page_num + 1, text.strip()))
            except Exception:
                slides_text.append((page_num + 1, ""))

    # 关闭文档
    if use_fitz:
        try:
            doc.close()
        except Exception:
            pass


    # ── 3. Quality Check: merge -> filter -> EasyOCR batch ──
    if use_fitz and slides_images and slides_text:
        st = _count_substantive_text(slides_text)
        if st < 200:
            print(f"  [QC] substantive {st}c<200, triggering EasyOCR")
            import os as _o
            for idx, (pn, _) in enumerate(slides_text):
                if idx < len(slides_images):
                    imgf = _o.path.join(str(att_dir), slides_images[idx])
                    if _o.path.exists(imgf):
                        ocr_res = _extract_text_via_easyocr(imgf)
                        if ocr_res: slides_text[idx] = (pn, ocr_res)
            print(f"  [QC] EasyOCR done for {len(slides_images)} pages")
        else: print(f"  [QC] substantive {st}c>=200, keep PyMuPDF")

    # ── 3. 打印截图统计 ──
    if use_fitz:
        if slides_failed > 0:
            print(f"  [PDF截图] 部分OK ({len(slides_images)}/{total_pages} 张, {slides_failed} 失败)")
        else:
            print(f"  [PDF截图] OK ({len(slides_images)} 张)")
    else:
        print(f"  [PDF截图] 跳过（降级模式，不支持截图）")

    # ── 5. 构建 Markdown ──
    print(f"  [MD构建] 正在生成 {md_path.name} ... ", end="", flush=True)

    has_images = len(slides_images) > 0
    md_lines = []

    for slide_num in range(1, total_pages + 1):
        md_lines.append(f"# 第{slide_num}页")
        md_lines.append("")

        # 文本内容
        if slide_num <= len(slides_text):
            text = slides_text[slide_num - 1][1]
            if text:
                md_lines.append(text)
                md_lines.append("")

        # 截图
        if has_images and slide_num <= len(slides_images):
            img_file = slides_images[slide_num - 1]
            md_lines.append(f'<img src="attachments/{img_file}" width="100%">')
            md_lines.append("")

        md_lines.append("")

    md_content = '\n'.join(md_lines)
    md_path.write_text(md_content, encoding="utf-8")
    print(f"OK ({total_pages} 页, {len(md_content)} 字符)")

    return md_path


def llm_cleanup_slide_md(md_path, source_filename):
    """
    对 PPT 快照生成的 Markdown 进行 LLM 排版优化（带重试与中断恢复）：
    - 合并相邻页面的文本内容
    - 识别并去除页眉页脚冗余
    - 重新组织章节结构
    - 优化标题层级（从 # 起步）
    - 分段失败时保留该段原文，不影响整体输出
    """
    content = Path(md_path).read_text(encoding="utf-8")

    if len(content) < 200:
        print(f"  [LLM排版] 内容过短，跳过")
        return

    # 分段处理大文件
    max_chars = 32000
    if len(content) > max_chars:
        print(f"  [LLM排版] 内容过长 ({len(content)} chars)，分段处理 ({LLM_MAX_RETRIES}次重试保护) ...")
        lines = content.split('\n')
        chunks = []
        curr = []
        curr_len = 0
        for line in lines:
            if curr_len + len(line) > max_chars and curr:
                chunks.append('\n'.join(curr))
                curr = []
                curr_len = 0
            curr.append(line)
            curr_len += len(line) + 1
        if curr:
            chunks.append('\n'.join(curr))

        results = []
        failed_chunks = 0
        for i, chunk in enumerate(chunks):
            print(f"    [{i+1}/{len(chunks)}] ... ", end="", flush=True)
            result = call_llm_with_retry(
                f"优化以下PPT内容排版（第{i+1}段）：\n\n{chunk}",
                PPT_SLIDE_CLEANUP_PROMPT,
                max_tokens=16000,
                temperature=1.0,
            )
            if result is not None:
                results.append(result)
                print("OK")
            else:
                failed_chunks += 1
                results.append(chunk)
                print("(保留原文)")
        if failed_chunks > 0:
            print(f"  [LLM排版] ⚠ {failed_chunks}/{len(chunks)} 段排版失败，该段保留原文排版")
        content = '\n\n'.join(results)
    else:
        print(f"  [LLM排版] ({len(content)} chars, {LLM_MAX_RETRIES}次重试保护) ... ", end="", flush=True)
        result = call_llm_with_retry(
            f"优化以下PPT内容排版：\n\n{content}",
            PPT_SLIDE_CLEANUP_PROMPT,
            max_tokens=16000,
            temperature=1.0,
        )
        if result is not None:
            content = result
            print("OK")
        else:
            print("LLM排版失败，保留原文排版")
            return

    # 自动清理文档抬头表格
    content = remove_document_header_tables(content)

    # 如果存在目录，删除目录前内容
    content = trim_before_toc(content)

    Path(md_path).write_text(content, encoding="utf-8")


# PPT 快照后处理 LLM Prompt
PPT_SLIDE_CLEANUP_PROMPT = """你是一个专业的文档排版专家。对从PPT幻灯片逐页提取的Markdown内容进行智能整合排版。

规则：
1. **合并相邻页**：如果连续多页的文本内容属于同一话题，将它们合并为连贯的段落。去除页面间的重复信息。

2. **删除冗余页眉**：删除重复出现的公司名称、会议名称、比赛名称等（如"中核华兴建设有限公司"、"NIC 2026"），保留正文内容。

3. **章节标题提取**：从内容中识别真正的章节标题（如"项目背景"、"技术方案"、"需求分析"等），为它们添加 # 标题标记。使用 H1 (#) 作为最高级标题。

4. **保留图片引用**：`<img src="..." width="100%">` 标签必须原样保留，不要删除或修改。

5. **保留表格**：表格内容原样保留。

6. **保留工艺流程步骤**：包含"工艺流程步骤"、"施工步骤"、"操作步骤"、"工艺流程"、"施工流程"等关键词的步骤型内容必须完整保留，不得省略或精简任何步骤编号和内容。

7. **优化段落结构**：
   - 将零散的罗列项合并为列表或段落
   - 去除无意义的空行和重复换行
   - 确保句子通顺连贯

8. **删除页码**：删除所有形式的页码信息。

9. **禁止**：不要添加原文没有的事实信息。图片引用绝对不能删除。

请直接输出优化后的Markdown，不要添加任何解释。"""


def remove_page_numbers(content, is_ppt=False):
    """
    删除页码信息：
    - "第X页/共Y页", "第X页"
    - "- X -", "— X —" (常见页码格式)
    - PPT常见: "1/20", "Slide 5", "Page 5", "[1/20]"
    - 独立行上的纯数字页码
    """
    lines = content.split('\n')
    new_lines = []
    for line in lines:
        stripped = line.strip()
        # 检查是否为纯页码行（含 PPT 格式）
        is_page_num = False
        page_num_patterns = [
            r'^第\s*\d+\s*页(?:\s*/\s*共\s*\d+\s*页)?$',
            r'^第\s*\d+\s*/\s*\d+\s*页$',
            r'^[-–]\s*\d+\s*[-–]$',                     # "- 5 -" / "– 5 –"
            r'^—\s*\d+\s*—$',                           # "— 5 —"
            r'^\d{1,3}\s*/\s*\d{1,3}$',                 # "1/20", "5/20"
            r'^[Ss]lide\s+\d+$',                        # "Slide 5"
            r'^[Pp]age\s+\d+$',                         # "Page 5"
            r'^\[\d{1,3}/\d{1,3}\]$',                   # "[1/20]"
            r'^\(\d{1,3}/\d{1,3}\)$',                   # "(1/20)"
        ]
        for pat in page_num_patterns:
            if re.match(pat, stripped):
                is_page_num = True
                break
        if is_page_num:
            continue

        # 行内页码替换
        line = re.sub(r'第\s*\d+\s*页(?:\s*/\s*共\s*\d+\s*页)?', '', line)
        line = re.sub(r'第\s*\d+\s*/\s*\d+\s*页', '', line)
        # PPT 内联页码清理（如 "标题 — 5 —" → "标题"）
        if is_ppt:
            line = re.sub(r'\s*[—–-]\s*\d+\s*[—–-]\s*$', '', line)
            line = re.sub(r'\s*\d{1,3}\s*/\s*\d{1,3}\s*$', '', line)
        new_lines.append(line)

    # 清理多余空行
    result = '\n'.join(new_lines)
    result = re.sub(r'\n{3,}', '\n\n', result)
    result = result.strip() + '\n'
    return result


def clean_ppt_redundant_lines(content, manual_keywords=None):
    """
    PPT/PPTX 冗余行清理：
    1. 统计所有独立非标题/非图片/非表格短行（< 40 字符）的出现次数
    2. 出现 >= 2 次且匹配特定模式的 → 判定为页眉/页脚冗余，删除
    3. 手动指定关键词精确删除（支持模糊匹配）
    
    目标：去除PPT上方重复的公司logo文字、会议名称、比赛名称等
    """
    lines = content.split('\n')
    
    # ── 收集候选短行（非标题、非图片、非表格、非空行）──
    heading_pattern = re.compile(r'^#{1,6}\s+')
    img_pattern = re.compile(r'^<img\s|^!\[|^\s*<img\s')
    table_pattern = re.compile(r'^\s*\|')
    
    candidate_count = {}
    candidate_indices = {}  # 记录每个候选行首次出现的位置
    for i, line in enumerate(lines):
        stripped = line.strip()
        if not stripped:
            continue
        if heading_pattern.match(stripped):
            continue
        if img_pattern.match(stripped):
            continue
        if table_pattern.match(stripped):
            continue
        if len(stripped) < 2:
            continue  # 太短的行（如单独数字）
        if len(stripped) > 50:
            continue  # 太长的行，不是页眉
        # 记录候选
        norm = stripped.lower()
        candidate_count[norm] = candidate_count.get(norm, 0) + 1
        if norm not in candidate_indices:
            candidate_indices[norm] = i
    
    # ── 判定冗余行 ──
    # 出现 >= 2 次的短行，排除明显是正文内容关键词的行
    body_keywords = [
        '摘要', 'abstract', '引言', '前言', '结论', '参考', '总结', '致谢',
        '目录', '背景', '目标', '方案', '设计', '实施', '分析', '评估',
        '建议', '措施', '要求', '标准', '规范', '参数', '性能', '指标'
    ]
    
    redundant_norms = set()
    for norm, count in candidate_count.items():
        if count < 2:
            continue
        # 排除正文关键词
        is_body = any(kw in norm for kw in body_keywords)
        if is_body:
            continue
        redundant_norms.add(norm)
    
    # ── 手动关键词 ──
    if manual_keywords:
        manual_kws = [kw.strip().lower() for kw in manual_keywords]
        if manual_kws:
            print(f"  [冗余清理] 手动关键词: {manual_keywords}")
    
    # ── 执行删除 ──
    removed_count = 0
    new_lines = []
    for i, line in enumerate(lines):
        stripped = line.strip()
        should_remove = False
        
        # 模式1：频率统计判定
        norm = stripped.lower()
        if norm in redundant_norms:
            should_remove = True
        
        # 模式2：手动关键词匹配
        if manual_keywords and stripped:
            for kw in manual_keywords:
                if kw in stripped.lower():
                    # 确保不是正文重要内容（检查行长度和上下文）
                    if len(stripped) < 60:
                        should_remove = True
                        break
        
        if should_remove:
            removed_count += 1
            continue
        new_lines.append(line)
    
    if removed_count > 0:
        print(f"  [冗余清理] 删除 {removed_count} 行页眉/页脚冗余内容")
    
    result = '\n'.join(new_lines)
    result = re.sub(r'\n{3,}', '\n\n', result)
    return result.strip() + '\n'


def reorganize_ppt_chapters(content):
    """
    PPT 章节标题去重重组：
    PPT 中同一章节标题（如"项目背景"、"技术方案"）可能在多页重复出现，
    本函数检测重复标题，将同一标题下的内容合并为逻辑连贯的章节结构。
    
    处理逻辑：
    1. 扫描所有 ## 标题，统计重复出现者
    2. 对于重复标题：保留首次出现，后续相同标题下的内容合并到首次标题后
    3. 删除后续的重复标题行，其内容跟随首个标题
    
    示例：
    ## 项目背景         ──→    ## 项目背景
    内容A                       内容A
    内容B                       - 内容B
    ## 技术方案                  - 内容C
    内容C                       ## 技术方案
    ## 项目背景（重复）          内容D
    内容D                       内容E
    ## 技术方案（重复）          ...（内容合并）
    内容E
    """
    from collections import defaultdict
    
    lines = content.split('\n')
    
    # ── 解析文档为 section 列表 ──
    # 每个 section = (heading_line_idx, heading_text, heading_level, body_lines)
    sections = []
    current_heading = None
    current_body = []
    in_preamble = True  # 第一个标题之前的内容
    
    for i, line in enumerate(lines):
        stripped = line.strip()
        heading_match = re.match(r'^(#{2,4})\s+(.+)', stripped)
        
        if heading_match:
            level = len(heading_match.group(1))
            title = heading_match.group(2).strip()
            
            # 保存前一个 section
            if in_preamble and current_body:
                sections.append({
                    'type': 'preamble',
                    'heading': None,
                    'level': 0,
                    'title': '',
                    'body': current_body,
                    'first_idx': -1
                })
            elif current_heading is not None:
                sections.append({
                    'type': 'section',
                    'heading': current_heading,
                    'level': current_heading['level'],
                    'title': current_heading['title'],
                    'body': current_body,
                    'first_idx': current_heading['idx']
                })
            
            in_preamble = False
            current_heading = {'idx': i, 'level': level, 'title': title, 'line': line}
            current_body = []
        else:
            current_body.append(line)
    
    # 最后一个 section
    if current_heading is not None:
        sections.append({
            'type': 'section',
            'heading': current_heading,
            'level': current_heading['level'],
            'title': current_heading['title'],
            'body': current_body,
            'first_idx': current_heading['idx']
        })
    elif in_preamble and current_body:
        sections.append({
            'type': 'preamble',
            'heading': None,
            'level': 0,
            'title': '',
            'body': current_body,
            'first_idx': -1
        })
    
    # ── 检测重复标题 ──
    title_count = defaultdict(list)
    for sec in sections:
        if sec['type'] == 'section':
            norm_title = sec['title'].strip().lower()
            title_count[norm_title].append(sec)
    
    duplicate_titles = {t for t, secs in title_count.items() if len(secs) >= 2}
    
    if not duplicate_titles:
        return content  # 无重复标题，直接返回
    
    print(f"  [章节重组] 检测到 {len(duplicate_titles)} 个重复标题，正在合并章节...")
    
    # ── 合并：首次出现保留标题，后续相同标题合并内容 ──
    seen_titles = {}
    merged_output = []
    
    for sec in sections:
        if sec['type'] == 'preamble':
            merged_output.extend(sec['body'])
            merged_output.append('')
            continue
        
        norm_title = sec['title'].strip().lower()
        
        if norm_title in duplicate_titles:
            if norm_title in seen_titles:
                # 重复标题：不添加标题行，只合并内容
                body = '\n'.join(sec['body']).strip()
                if body:
                    merged_output.append('')
                    merged_output.append(body)
                    merged_output.append('')
            else:
                # 首次出现：保留标题
                seen_titles[norm_title] = True
                merged_output.append(sec['heading']['line'])
                merged_output.extend(sec['body'])
                merged_output.append('')
        else:
            # 非重复标题：正常输出
            merged_output.append(sec['heading']['line'])
            merged_output.extend(sec['body'])
            merged_output.append('')
    
    result = '\n'.join(merged_output)
    result = re.sub(r'\n{3,}', '\n\n', result)
    return result.strip() + '\n'


def normalize_headings(content):
    """
    自动识别并规范化标题级别：
    - 确保 # 从 # 开始（H1，PPT和PDF统一）
    - 同一级别标题使用一致的 # 数量
    - 修复编号标题（如 "1.1 xxx" 不加 # 的情况）
    """
    lines = content.split('\n')
    result = []
    in_table = False

    for line in lines:
        stripped = line.strip()

        # 跳过表格内的处理
        if stripped.startswith('|') or stripped.startswith('<table'):
            in_table = True
        if in_table:
            result.append(line)
            if stripped.endswith('|') or '</table>' in stripped:
                in_table = False
            continue

        # 已有的 Markdown 标题：规范化到 H1 起步
        heading_match = re.match(r'^(#{1,6})\s+(.+)', stripped)
        if heading_match:
            hashes = heading_match.group(1)
            title_text = heading_match.group(2)
            # 至少 # 起步，最多 ######
            level = min(max(len(hashes), 1), 6)
            result.append(f"{'#' * level} {title_text}")
            continue

        # 检测段落中嵌入的编号标题模式（如 "1.1 标题内容" 或 "（一）标题"）
        heading_patterns = [
            (r'^(\d+(?:\.\d+)*)\s+(.+)$', 2),  # "1.1 xxx" → ##
            (r'^（([一二三四五六七八九十]+)）\s*(.+)$', 2),  # "（一）xxx" → ##
            (r'^([一二三四五六七八九十]+)[、，]\s*(.+)$', 2),  # "一、xxx" → ##
        ]
        matched = False
        for pat, level in heading_patterns:
            m = re.match(pat, stripped)
            if m and len(stripped) < 80:  # 短行更可能是标题
                title = m.group(0)
                result.append(f"{'#' * level} {title}")
                matched = True
                break
        if not matched:
            result.append(line)

    return '\n'.join(result)


def llm_reorganize_pptx(content, source_filename):
    """
    利用 LLM 对 PPT/PPTX 内容进行智能整合排版（默认 DeepSeek，可通过 PPT_LLM_* 环境变量切换）：
    - 合并罗列信息
    - 优化段落结构
    - 保持原意不变
    - 带重试与中断恢复：单段失败保留原文，不影响整体
    """
    fname = Path(source_filename).name

    # 如果内容太小，跳过 LLM 处理
    if len(content) < 200:
        print(f"  [LLM] 内容过短，跳过智能排版")
        return content

    # 截断过长内容（避免 token 超限）
    max_chars = 30000
    if len(content) > max_chars:
        return llm_reorganize_chunked(content, source_filename, max_chars)

    print(f"  [LLM] 正在智能排版 ({len(content)} chars, {LLM_MAX_RETRIES}次重试保护) ... ", end="", flush=True)

    system_prompt = """你是一个专业的文档排版专家。你的任务是对从PPT/PPTX提取的Markdown内容进行智能排版优化。

规则：
1. **合并罗列信息**：将分散的、短促的罗列项整合为流畅的段落或结构化列表。
   例如原始内容：
   ```
   受限空间出入口
   受限空间内
   人形报警
   人脸识别
   ```
   应整合为：
   ```
   - 受限空间出入口监控与人脸识别
   - 受限空间内人形报警
   ```

2. **删除页眉页脚冗余**：删除每页幻灯片顶部重复出现的公司logo文字/公司名称/会议名称/比赛名称（如"中核华兴"、"NIC 2026"、"第十四届核电建造技术交流会"），删除底部重复出现的页码信息。保留正文内容不变。

3. **保留标题层级**：保持原有的 # 标题结构不变。

4. **保留表格**：表格内容原样保留，不做修改。

5. **保留图片引用**：`<img ...>` 标签原样保留。

6. **保留工艺流程步骤**：包含"工艺流程步骤"、"施工步骤"、"操作步骤"、"工艺流程"、"施工流程"等关键词的步骤型内容必须完整保留，不得省略或精简任何步骤编号和内容。

7. **优化段落**：将零散的短句合并为连贯的段落，去除冗余换行。

8. **章节内容整合**：如果同一章节标题（如"项目背景"、"技术方案"）在内容中多次出现，将相同标题下的内容合并到一个章节中，避免标题重复。

9. **禁止**：不要添加原文没有的信息，不要改变事实内容。

请直接输出优化后的Markdown，不要添加任何解释。"""

    result = call_llm_with_retry(
        f"请对以下PPT/PPTX提取的Markdown内容进行智能排版优化：\n\n{content}",
        system_prompt,
        max_tokens=16000,
        temperature=1.0,
    )
    if result is not None:
        print("OK")
        return result
    else:
        print("(保留原文)")
        return content


def llm_reorganize_chunked(content, source_filename, max_chars):
    """分段处理长文档（带重试与中断恢复）"""
    lines = content.split('\n')
    chunks = []
    current_chunk = []
    current_len = 0

    for line in lines:
        line_len = len(line) + 1
        if current_len + line_len > max_chars and current_chunk:
            chunks.append('\n'.join(current_chunk))
            current_chunk = []
            current_len = 0
        current_chunk.append(line)
        current_len += line_len
    if current_chunk:
        chunks.append('\n'.join(current_chunk))

    print(f"  [LLM] 分段处理 ({len(chunks)} 段, {LLM_MAX_RETRIES}次重试保护) ...")
    results = []
    failed_chunks = 0
    system_prompt = """你是一个专业的文档排版专家。对PPT提取内容进行智能排版：
- 合并罗列信息为流畅列表/段落
- 删除页眉页脚冗余（公司logo/名称、会议/比赛名称、页码）
- 保留标题、表格、图片引用
- 完整保留工艺流程步骤、施工步骤等步骤型内容，不得省略任何步骤编号和内容
- 优化段落结构
- 合并重复出现的章节标题下的内容
- 不添加原文没有的信息
直接输出优化后的Markdown，不加解释。"""

    for i, chunk in enumerate(chunks):
        print(f"    [{i+1}/{len(chunks)}] ... ", end="", flush=True)
        result = call_llm_with_retry(
            f"请优化以下Markdown内容（第{i+1}段）：\n\n{chunk}",
            system_prompt,
            max_tokens=16000,
            temperature=1.0,
        )
        if result is not None:
            results.append(result)
            print("OK")
        else:
            failed_chunks += 1
            results.append(chunk)
            print("(保留原文)")

    if failed_chunks > 0:
        print(f"  [LLM] {failed_chunks}/{len(chunks)} 段排版失败，保留原文")

    return '\n\n'.join(results)


def save_original_file(source_path, output_sub_dir):
    """将源文件复制到输出子文件夹的 'original file' 目录中"""
    src = Path(source_path)
    dst_dir = Path(output_sub_dir) / "original file"
    safe_makedirs(str(dst_dir), exist_ok=True)
    dst = dst_dir / src.name
    shutil.copy2(src, dst)
    print(f"  [存档] 源文件 → original file/{src.name}")


def remove_document_header_tables(content):
    """
    自动检测并删除 MinerU 解析残留的文档抬头 HTML 表格。

    识别特征：
    - <table> 标签，第一行第一格含 rowspan="2"
    - 包含机构/单位名称（不限于特定单位）
    - 包含"文件编码"或"文件名称"
    - 包含"状态"和"版次"
    - 文件编码格式：字母+空格+数字（如 "CW 1 XXX15 ..."）
    """
    import re

    # 匹配：<table> 开头，rowspan="2" 的第一格含单位名，然后文件编码/名称/状态/版次
    # 由于可能有 <img> 嵌入，需要用宽松匹配
    pattern = re.compile(
        r'<table>'
        r'<tr>'
        r'<td\s+rowspan="2">'
        r'(?:<img[^>]*/>)?'      # 可能嵌入的图片
        r'[^<]*'                  # 单位名称
        r'</td>'
        r'<td>[^<]*文件编码[^<]*</td>'
        r'<td>[A-Z]+[\s\d]+[^<]*</td>'  # 文件编码（字母+空格+数字格式）
        r'<td>[^<]*状态[^<]*</td>'
        r'<td>[^<]*版次[^<]*</td>'
        r'</tr>'
        r'<tr>'
        r'<td>[^<]*文件名称[^<]*</td>'
        r'<td>[^<]*</td>'          # 文件名称内容
        r'<td>[^<]*(?:CFC|PRE|DES|SS)[^<]*</td>'  # 状态值（常用缩写）
        r'<td>[^<]*</td>'          # 版次值
        r'</tr>'
        r'</table>',
        re.IGNORECASE
    )

    before = content.count('<table>')
    content = pattern.sub('', content)
    removed = before - content.count('<table>')

    # 清理产生的连续空行
    content = re.sub(r'\n{3,}', '\n\n', content)

    if removed > 0:
        print(f"  [抬头清理] 自动删除 {removed} 处文档抬头表格")

    return content


def post_process_md(md_path, source_path, manual_remove_keywords=None, no_chapter_reorg=False, is_ppt_like=False):
    """
    PDF MinerU 输出后处理管线：
    1. 截取正文（去封面/目录/修订记录）
    2. 图片宽度100%
    3. 自动删除文档抬头 HTML 表格
    4. PPT-like PDF: 冗余行清理
    5. 删除页码（PPT-like 使用加强模式）
    6. PPT-like PDF: 章节标题去重重组
    7. 规范化标题（从 H1 起步）
    8. PPT-like PDF: LLM智能排版
    """
    content = Path(md_path).read_text(encoding="utf-8")
    needs_ppt_cleanup = is_ppt_like

    # 1. 正文截取
    content = trim_pdf_front_matter(content)

    # 2. 图片宽度
    content = fix_image_width(content)

    # 3. 自动删除文档抬头表格
    content = remove_document_header_tables(content)

    # 4. PPT-like PDF: 冗余行清理
    if needs_ppt_cleanup:
        kws = manual_remove_keywords.split(',') if manual_remove_keywords else None
        content = clean_ppt_redundant_lines(content, manual_keywords=kws)

    # 5. 删除页码
    content = remove_page_numbers(content, is_ppt=needs_ppt_cleanup)

    # 6. PPT-like PDF: 章节重组
    if needs_ppt_cleanup and not no_chapter_reorg:
        content = reorganize_ppt_chapters(content)

    # 7. 标题规范化（H1 起步）
    content = normalize_headings(content)

    # 8. PPT-like PDF: LLM 排版
    if needs_ppt_cleanup:
        content = llm_reorganize_pptx(content, source_path)

    # 9. 目录截取：如果存在目录，删除目录前内容
    content = trim_before_toc(content)

    Path(md_path).write_text(content, encoding="utf-8")
    return md_path


def main():
    parser = argparse.ArgumentParser(description="文档 → Markdown 批量转换 (PDF/PPT/PPTX)")
    parser.add_argument("source_dir", help="包含文档文件的源文件夹路径")
    parser.add_argument("-o", "--output", required=True, help="输出根目录")
    parser.add_argument("--model", default="vlm", choices=["pipeline", "vlm", "MinerU-HTML"],
                        help="MinerU 模型版本 (仅 PDF，默认: vlm)")
    parser.add_argument("--no-formula", action="store_true", help="关闭 PDF 公式识别")
    parser.add_argument("--no-table", action="store_true", help="关闭 PDF 表格识别")
    parser.add_argument("--language", default="ch", help="文档语言 (默认: ch)")
    parser.add_argument("--batch-size", type=int, default=BATCH_SIZE,
                        help=f"单批次文件数 (仅 PDF，默认: {BATCH_SIZE}, 最大: 50)")
    parser.add_argument("--no-postprocess", action="store_true", help="跳过后处理")
    parser.add_argument("--no-llm", action="store_true", help="跳过 LLM 排版")
    parser.add_argument("--no-chapter-reorg", action="store_true", help="跳过 PPT-like PDF 章节重组")
    parser.add_argument("--no-archive", action="store_true", help="跳过源文件存档（不复制到 original file/）")
    parser.add_argument("--remove-lines", default=None,
                        help="手动指定要删除的冗余行关键词（如：中核华兴,NIC 2026）")
    args = parser.parse_args()

    # ── 初始化计数器 ──
    success_count = 0
    fail_count = 0
    skip_count = 0
    llm_failed_count = 0  # LLM 排版部分失败的文档数

    files = collect_files(args.source_dir)
    pdf_files = [f for f in files if is_pdf_file(f.name)]
    ppt_files = [f for f in files if is_ppt_file(f.name)]

    # ── 检测 PPT 导出的 PDF（优先使用截图管线）──
    ppt_like_pdfs = []
    normal_pdfs = []
    for pdf_file in pdf_files:
        if detect_ppt_like_pdf(pdf_file):
            ppt_like_pdfs.append(pdf_file)
        else:
            normal_pdfs.append(pdf_file)

    total_size = sum(f.stat().st_size for f in files) / 1024 / 1024

    # ── 显示 LLM 配置 ──
    llm_cfg = get_llm_config()

    print(f"\n{'='*60}")
    print(f"文档 → Markdown 批量转换")
    print(f"源文件夹: {args.source_dir}")
    print(f"输出目录: {args.output}")
    print(f"文件总数: {len(files)} (PPT/PPTX: {len(ppt_files)}, PPT导出PDF: {len(ppt_like_pdfs)}, 普通PDF: {len(normal_pdfs)})")
    print(f"总大小: {total_size:.1f} MB")
    print(f"LLM 排版: {llm_cfg['model']} @ {llm_cfg['api_base']}")
    print(f"PPT/PPTX 管线: PowerPoint COM 导出PDF → PyMuPDF 快照 + 文本提取")
    print(f"PPT导出PDF 管线: PyMuPDF 快照 + 文本提取 (无MinerU)")
    print(f"普通PDF 管线: MinerU API (模型: {args.model})")
    print(f"后处理: {'关' if args.no_postprocess else '开'}")
    print(f"LLM排版: {'关' if args.no_llm else '开'}")
    if args.remove_lines:
        print(f"冗余行清理: {args.remove_lines}")
    print(f"{'='*60}\n")

    safe_makedirs(args.output, exist_ok=True)

    # ── Phase 1: PPT/PPTX 转换（快照管线）──
    for ppt_file in ppt_files:
        print(f"── PPT/PPTX: {ppt_file.name} ──")
        print(f"  [大小] {ppt_file.stat().st_size / 1024 / 1024:.1f} MB")
        try:
            md_path = convert_pptx_via_pdf(ppt_file, args.output)
            if md_path is None:
                print(f"  [WARN] {ppt_file.name}: PPT→PDF 导出失败，跳过")
                skip_count += 1
                continue
            # 保存源文件
            if not args.no_archive:
                save_original_file(ppt_file, md_path.parent)
            # 后处理
            if not args.no_postprocess:
                if not args.no_llm:
                    print(f"  [后处理] {md_path.name} ...")
                    llm_cleanup_slide_md(md_path, ppt_file.name)
                else:
                    # 至少做图片宽度 + 标题规范化
                    content = Path(md_path).read_text(encoding="utf-8")
                    content = fix_image_width(content)
                    content = normalize_headings(content)
                    Path(md_path).write_text(content, encoding="utf-8")
            success_count += 1
        except Exception as e:
            print(f"  [FAIL] {ppt_file.name}: {e}")
            import traceback
            traceback.print_exc()
            fail_count += 1

    # ── Phase 1.5: PPT 导出 PDF 转换（截图管线，无 MinerU）──
    for pdf_file in ppt_like_pdfs:
        print(f"── PPT导出PDF: {pdf_file.name} ──")
        print(f"  [大小] {pdf_file.stat().st_size / 1024 / 1024:.1f} MB")
        try:
            md_path = convert_ppt_pdf_via_slides(pdf_file, args.output)
            if md_path is None:
                print(f"  [WARN] {pdf_file.name}: 转换失败，降级到 MinerU 管线 ...")
                normal_pdfs.append(pdf_file)
                skip_count += 1
                continue
            # 保存源文件
            if not args.no_archive:
                save_original_file(pdf_file, md_path.parent)
            # 后处理：LLM 智能排版（复用 PPT 管线逻辑）
            if not args.no_postprocess:
                if not args.no_llm:
                    print(f"  [后处理] {md_path.name} ...")
                    llm_cleanup_slide_md(md_path, pdf_file.name)
                else:
                    content = Path(md_path).read_text(encoding="utf-8")
                    content = fix_image_width(content)
                    content = normalize_headings(content)
                    Path(md_path).write_text(content, encoding="utf-8")
            success_count += 1
        except Exception as e:
            print(f"  [FAIL] {pdf_file.name}: {e}")
            import traceback
            traceback.print_exc()
            # 降级到 MinerU
            print(f"  [降级] 回退到 MinerU 管线 ...")
            normal_pdfs.append(pdf_file)
            fail_count += 1

    # ── Phase 2: 普通 PDF 转换（MinerU 管线）──
    if normal_pdfs:
        batch_size = min(args.batch_size, 50)
        enable_formula = not args.no_formula
        enable_table = not args.no_table

        chunks = list(chunk_list(normal_pdfs, batch_size))
        for ci, chunk in enumerate(chunks):
            chunk_num = ci + 1
            chunk_sz = sum(f.stat().st_size for f in chunk) / 1024 / 1024
            print(f"\n── PDF 第 {chunk_num}/{len(chunks)} 批 ({len(chunk)} 个, {chunk_sz:.1f} MB) ──")

            print("[1/3] 申请上传链接...")
            try:
                batch_id, file_urls = submit_batch(
                    chunk, args.model, enable_formula, enable_table, args.language
                )
                print(f"  batch_id: {batch_id}")
            except Exception as e:
                print(f"  [ERROR] {e}")
                fail_count += len(chunk)
                continue

            print("[2/3] 上传文件到 OSS...")
            try:
                upload_files(chunk, file_urls)
            except Exception as e:
                print(f"  [WARN] 部分上传失败: {e}")

            print("[3/3] 等待解析完成...")
            try:
                extract_results = poll_batch_results(batch_id)
            except Exception as e:
                print(f"  [ERROR] {e}")
                fail_count += len(chunk)
                continue

            file_map = {f.name: f for f in chunk}
            print("[4/4] 下载结果...")
            for r in extract_results:
                fname = r.get("file_name", "unknown")
                src_path = file_map.get(fname)
                if r.get("state") == "done":
                    md_path = download_and_extract(r, args.output, src_path)
                    if md_path and src_path:
                        sub_dir = md_path.parent
                        if not args.no_archive:
                            save_original_file(src_path, sub_dir)

                        if not args.no_postprocess:
                            print(f"  [后处理] {md_path.name} ...")
                            if not args.no_llm:
                                post_process_md(md_path, src_path,
                                                manual_remove_keywords=args.remove_lines,
                                                no_chapter_reorg=args.no_chapter_reorg,
                                                is_ppt_like=False)
                            else:
                                # 仅基础后处理
                                content = Path(md_path).read_text(encoding="utf-8")
                                content = trim_pdf_front_matter(content)
                                content = fix_image_width(content)
                                content = remove_page_numbers(content, is_ppt=False)
                                content = normalize_headings(content)
                                Path(md_path).write_text(content, encoding="utf-8")
                        success_count += 1
                    else:
                        fail_count += 1
                else:
                    print(f"  [FAIL] {fname}: {r.get('err_msg', r.get('state'))}")
                    fail_count += 1

    # ── 汇总报告 ──
    md_outputs = sorted(Path(args.output).rglob("*.md"))
    md_outputs = [m for m in md_outputs if '.tmp_' not in str(m)]

    print(f"\n{'='*60}")
    total = success_count + fail_count + skip_count
    print(f"✅ 转换完成 (成功: {success_count}, 失败: {fail_count}, 跳过: {skip_count}, 共计: {total})")
    print(f"输出目录: {args.output}")
    print(f"{'='*60}")

    if md_outputs:
        print(f"\n生成的 Markdown 文件 ({len(md_outputs)}):")
        for m in md_outputs:
            size_kb = m.stat().st_size / 1024
            rel = m.relative_to(args.output)
            print(f"  {rel}  ({size_kb:.1f} KB)")

    if not args.no_archive:
        orig_dirs = sorted(Path(args.output).rglob("original file"))
        if orig_dirs:
            print(f"\n存档源文件目录 ({len(orig_dirs)}):")
            for d in orig_dirs:
                files_in = list(d.iterdir())
                if files_in:
                    rel = d.relative_to(args.output)
                    print(f"  {rel}/ ({len(files_in)} 个文件)")


if __name__ == "__main__":
    main()
