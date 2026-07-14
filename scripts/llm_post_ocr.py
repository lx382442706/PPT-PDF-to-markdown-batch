#!/usr/bin/env python
"""对 OCR 提取的 Markdown 进行 LLM 智能排版（独立脚本）"""
import os, sys, time, argparse
from openai import OpenAI

sys.stdout.reconfigure(encoding='utf-8')

parser = argparse.ArgumentParser(description="LLM post-process OCR Markdown")
parser.add_argument("md_path", help="Path to the markdown file")
args = parser.parse_args()

md_path = args.md_path
content = open(md_path, 'r', encoding='utf-8').read()
print(f'原始大小: {len(content)} chars')

# Keep the ES link line
lines = content.split('\n')
es_link = lines[0] if lines[0].startswith('>') else ''
body = '\n'.join(lines[1:]) if es_link else content

if len(body.strip()) < 200:
    print('内容过短，跳过LLM排版')
    sys.exit(0)

SYSTEM_PROMPT = """你是一个专业的核电文档排版专家。对从PPT幻灯片OCR提取的Markdown内容进行智能整合排版。

规则：
1. **修复OCR错误**：识别并修正OCR识别错误，保持专业术语准确。常见错误示例：
   - 集回→集团、中核O0CGN→中核CGN、桉nq→堆芯、眺百→的
   - 豉W门汇捉→高层汇报、娑罖益降笨增竣亚铤系→协同降本增效全产业链体系
   - 皴发→研发、筻童→第6章、课丝/课跎/课甄/课足/课噩→课题X
   - G0CGN→CGN、Hualong Tech。→Hualong Tech
   - 页码标记（如4157→无、5157→无）应删除

2. **合并相邻页**：相邻页面同属一个章节的内容合并为连贯段落，去除跨页重复的页眉信息。

3. **删除页面冗余**：
   - 删除每页重复的公司logo/名称行（"中核集团 华龙国际 中广核 CGN CNNC Hualong Tech"）
   - 删除页面底部的页码标记
   - 删除重复出现的目录页——保留第一份即可

4. **章节标题提取**：从内容中识别真正的章节标题，使用合适的Markdown标题层级：
   - H1 (#) 作为文档标题（封面信息）
   - H2 (##) 作为大章节（如"背景及必要性"、"总体技术方案"等）
   - H3 (###) 作为子章节

5. **保留图片引用**：所有 `<img src="..." width="100%">` 标签必须原样保留，不可删除或修改。每张图片紧跟对应章节内容之后。

6. **保留表格**：技术指标对比表、进度表等表格型内容转换为Markdown表格或结构化列表。

7. **完整保留课题清单和步骤**：课题设置情况、研发需求列表、进度节点等必须完整保留，不省略任何条目。

8. **优化段落结构**：
   - 零散短句合并为连贯段落
   - 去除无意义空行和重复换行
   - 相同话题合并到一个章节下

9. **禁止**：不要添加原文没有的事实信息。不要删除以 `> 原始文件:` 开头的ES链接行。

请直接输出优化后的Markdown，不要添加任何解释。"""

# Split by page markers and group
parts = body.split('\n# 第')
preamble = parts[0]
pages = ['# 第' + p for p in parts[1:]]
print(f'共 {len(pages)} 页')

# Group pages into ~6-page chunks
chunk_size = 5
chunked = []
for i in range(0, len(pages), chunk_size):
    chunk_pages = pages[i:i+chunk_size]
    if i == 0:
        chunked.append(preamble + '\n\n' + '\n\n'.join(chunk_pages))
    else:
        chunked.append('\n\n'.join(chunk_pages))

print(f'分 {len(chunked)} 段处理')

# API config
api_key = os.environ.get('PPT_LLM_API_KEY', 'sk-abc9b230c13b4cae923925d1b5bad2e0')
api_base = os.environ.get('PPT_LLM_API_BASE', 'https://api.deepseek.com/v1')
model = os.environ.get('PPT_LLM_MODEL', 'deepseek-chat')
client = OpenAI(api_key=api_key, base_url=api_base)

results = []
failed = 0
for idx, chunk in enumerate(chunked):
    print(f'  [{idx+1}/{len(chunked)}] ({len(chunk)} chars) ... ', end='', flush=True)
    success = False
    for attempt in range(3):
        try:
            resp = client.chat.completions.create(
                model=model,
                messages=[
                    {'role': 'system', 'content': SYSTEM_PROMPT},
                    {'role': 'user', 'content': f'优化以下PPT OCR内容（第{idx+1}段）：\n\n{chunk}'}
                ],
                temperature=1.0,
                max_tokens=16000,
                timeout=180,
            )
            results.append(resp.choices[0].message.content)
            print('OK')
            success = True
            break
        except Exception as e:
            if attempt < 2:
                wait = [2, 4][attempt]
                print(f'(重试 {attempt+1}/3, {wait}s)... ', end='', flush=True)
                time.sleep(wait)
            else:
                print(f'FAIL: {e}')
                failed += 1
                results.append(chunk)

# Assemble final result
final = es_link + '\n\n' + '\n\n'.join(results) if es_link else '\n\n'.join(results)

# Write back
open(md_path, 'w', encoding='utf-8').write(final)
size_kb = os.path.getsize(md_path) / 1024
print(f'\n=== LLM排版完成 ===')
print(f'文件: {md_path}')
print(f'大小: {size_kb:.1f} KB')
if failed > 0:
    print(f'警告: {failed}/{len(chunked)} 段排版失败，保留原文')
