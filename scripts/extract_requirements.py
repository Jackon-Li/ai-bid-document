#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
extract_requirements.py — 从招标文件中抽取可追溯的要求条目台账。

用法：
    python extract_requirements.py --input 招标文件.md
    python extract_requirements.py --input 招标文件.txt --output requirements.csv
    python extract_requirements.py --input 招标文件.docx          # 需安装 python-docx

支持的输入格式：.txt / .md / .docx（可选依赖）

输出：requirements.csv（UTF-8-BOM，Excel 可直接打开）

设计原则：
  1. 只做“候选条目”抽取，不替人做最终判定 —— 抽取结果必须人工过一遍；
  2. source_quote 保留原文，供复核时逐字对照；
  3. 强制词判定采用“强词优先”：强词与弱词同时出现时，按更严者执行；
  4. 用正则而非纯子串匹配，避免“投标人须知”的“须”、“故障响应”的“应”被误判。
"""

import argparse
import csv
import re
import sys
from pathlib import Path

# ---------------------------------------------------------------- 词表

# 强强制词（正则）。命中即 mandatory=Y。
STRONG_PATTERNS = [
    r"必须",
    r"(?<![响对适供相反呼感理])应(?![答用急变收商])",   # 排除“响应/适应/供应/相应”等
    r"须(?!知)",                                        # 排除“须知”
    r"不得", r"严禁", r"禁止", r"不允许", r"不能",
    r"不低于", r"不少于", r"不高于", r"不超过", r"不小于", r"不大于",
    r"至少", r"最少", r"最多", r"以上", r"以内",
    r"(?:需|须|应)(?:提供|提交|附|出具|加盖|签署|填写)",
]

# 弱强制词（命中则 mandatory=N，除非同时命中强词）
WEAK_PATTERNS = [r"建议", r"宜", r"鼓励", r"优先考虑", r"择优",
                 r"原则上", r"如有", r"若", r"可考虑", r"推荐", r"参照"]

# 类型判定：按顺序匹配，先匹配到的生效（顺序即优先级）
TYPE_RULES = [
    ("资格", ["营业执照", "独立法人", "法人资格", "资质证书", "资质等级", "财务报告",
              "审计报告", "社会保险", "社保", "纳税", "税收", "失信", "违法记录",
              "联合体", "注册资本", "开户许可", "信用中国", "信用证明",
              "安全生产许可", "管理体系", "ISO", "体系认证"]),
    ("格式", ["正本", "副本", "密封", "封套", "装订", "骑缝章", "盖章", "签字",
              "份数", "格式要求", "页码", "目录", "电子版", "U盘", "光盘",
              "投标函", "授权委托书", "标记", "封装"]),
    ("评分", ["得分", "评分", "分值", "加分", "计分", "满分", "扣分", "评审标准",
              "业绩", "低价优先", "评标办法"]),
    ("商务", ["工期", "交货", "交付期", "付款", "支付", "质保", "保修", "保证金",
              "违约", "赔偿", "知识产权", "保密", "验收", "合同", "报价", "价格",
              "质保金"]),
    ("服务", ["售后", "响应时间", "培训", "驻场", "备件", "运维", "巡检",
              "技术支持", "服务网点", "故障", "维护", "升级服务", "服务承诺"]),
    ("技术", ["技术", "性能", "参数", "接口", "协议", "容量", "功能", "兼容",
              "支持", "接入", "带宽", "存储", "并发", "时延", "分辨率",
              "系统", "设备", "部署", "架构", "标准", "算法", "加密", "冗余"]),
]

# 需要向招标方澄清的模糊表述
AMBIGUOUS_WORDS = ["或同等", "等同", "原则上", "一般情况下", "视情况", "如有必要",
                   "由招标人认定", "择优", "满足使用要求", "满足需要", "国内领先",
                   "根据实际情况", "适时"]

# 结构识别
CHAPTER_RE = re.compile(r"^\s*(第[一二三四五六七八九十百零\d]+[章部分篇])\s*[、.．:：]?\s*(.*)$")
CLAUSE_RE = re.compile(r"^\s*(\d+(?:\.\d+){1,3})\s*[、.．]?\s+(.+)$")
CLAUSE_CN_RE = re.compile(r"^\s*[（(]\s*([一二三四五六七八九十]+)\s*[)）]\s*(.+)$")
MD_HEADING_RE = re.compile(r"^\s*#{1,6}\s+(.+)$")

MIN_SENTENCE_LEN = 8       # 短于此长度的句子不作为要求条目
MAX_HEADING_LEN = 40       # 章节标题的最大长度，超过则视为正文
MAX_REQUIREMENT_LEN = 400


# ---------------------------------------------------------------- 读取

def read_text(path: Path) -> str:
    suffix = path.suffix.lower()
    if suffix == ".docx":
        try:
            import docx  # type: ignore
        except ImportError:
            sys.exit(
                "[错误] 读取 .docx 需要 python-docx。请先安装：\n"
                "    pip install python-docx\n"
                "或把招标文件另存为 .txt / .md 后重试。"
            )
        d = docx.Document(str(path))
        blocks = [p.text for p in d.paragraphs]
        for table in d.tables:
            for row in table.rows:
                blocks.append(" | ".join(c.text.strip() for c in row.cells))
        return "\n".join(blocks)

    for enc in ("utf-8-sig", "utf-8", "gb18030", "gbk"):
        try:
            return path.read_text(encoding=enc)
        except UnicodeDecodeError:
            continue
    sys.exit(f"[错误] 无法识别文件编码：{path}（请另存为 UTF-8）")


# ---------------------------------------------------------------- 判定

def detect_mandatory(sentence: str) -> str:
    if any(re.search(p, sentence) for p in STRONG_PATTERNS):
        return "Y"
    if any(re.search(p, sentence) for p in WEAK_PATTERNS):
        return "N"
    return "N"


def detect_type(sentence: str, section: str) -> str:
    """先看章节名（更准），再看句内关键词。"""
    for type_name, words in TYPE_RULES:
        if any(w in section for w in words):
            return type_name
    for type_name, words in TYPE_RULES:
        if any(w in sentence for w in words):
            return type_name
    return "技术"


def detect_response_type(sentence: str, req_type: str) -> str:
    if req_type == "资格" or re.search(r"(?:提供|提交|附|出具|加盖|签署)", sentence):
        return "证明"
    if re.search(r"(?:承诺|保证|不低于|不少于|不超过|不高于)", sentence):
        return "承诺"
    return "响应"


def detect_clarify(sentence: str) -> str:
    return "Y" if any(w in sentence for w in AMBIGUOUS_WORDS) else "N"


def split_sentences(text: str):
    return [p.strip() for p in re.split(r"(?<=[。；;！？])", text) if p.strip()]


# ---------------------------------------------------------------- 结构解析

def parse_line(line: str, current_chapter: str):
    """
    返回 (chapter, clause_no, content)。
    - 章标题行：chapter 更新，content 为空（跳过）
    - 条款行：clause_no 为编号，content 为条款正文
    - 普通行：沿用当前 chapter，clause_no 为空
    """
    stripped = line.strip()

    # markdown 标题
    m = MD_HEADING_RE.match(stripped)
    if m:
        stripped = m.group(1).strip()
        if len(stripped) <= MAX_HEADING_LEN and not stripped.endswith("。"):
            chap = CHAPTER_RE.match(stripped)
            if chap:
                return stripped, "", ""
            # markdown 章节标题（如 "## 第二章 技术要求" 或 "## 3.2 存储设计"）
            cl = CLAUSE_RE.match(stripped)
            if cl:
                return current_chapter, cl.group(1), cl.group(2)
            return current_chapter, "", ""   # 纯标题，跳过

    # 第X章 标题
    chap = CHAPTER_RE.match(stripped)
    if chap:
        rest = chap.group(2).strip()
        if len(rest) <= MAX_HEADING_LEN:
            return stripped, "", ""          # 章标题行，不是要求
        return stripped, "", rest            # 形如“第二章 系统应支持…”的合并行

    # 数字条款 1.1 / 2.3.1
    cl = CLAUSE_RE.match(stripped)
    if cl and len(cl.group(2)) >= 4:
        return current_chapter, cl.group(1), cl.group(2)

    # 中文括号编号 （一）（二）
    cn = CLAUSE_CN_RE.match(stripped)
    if cn and len(cn.group(2)) >= 4:
        return current_chapter, cn.group(1), cn.group(2)

    return current_chapter, "", stripped


def extract(text: str, min_len: int = MIN_SENTENCE_LEN):
    rows = []
    current_chapter = "（未识别章节）"
    total_candidates = 0

    for raw_line in text.splitlines():
        if not raw_line.strip():
            continue

        current_chapter, clause_no, content = parse_line(raw_line, current_chapter)
        if not content:
            continue

        section = f"{current_chapter} {clause_no}".strip() if clause_no else current_chapter

        for sent in split_sentences(content):
            sent = re.sub(r"^[\-\*\u2022]+\s*", "", sent).strip()
            if len(sent) < min_len:
                continue
            total_candidates += 1
            if len(sent) > MAX_REQUIREMENT_LEN:
                sent = sent[:MAX_REQUIREMENT_LEN] + "…"

            req_type = detect_type(sent, section)
            rows.append({
                "req_id": f"R-{len(rows) + 1:03d}",
                "section": section,
                "requirement": sent,
                "type": req_type,
                "mandatory": detect_mandatory(sent),
                "source_quote": sent,
                "response_type": detect_response_type(sent, req_type),
                "need_clarify": detect_clarify(sent),
                "owner": "",
                "status": "待响应",
                "note": "",
            })

    return rows, total_candidates


def write_csv(rows, out_path: Path):
    fields = ["req_id", "section", "requirement", "type", "mandatory",
              "source_quote", "response_type", "need_clarify", "owner", "status", "note"]
    with out_path.open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=fields)
        w.writeheader()
        w.writerows(rows)


def main():
    ap = argparse.ArgumentParser(
        description="从招标文件抽取可追溯的要求条目台账（requirements.csv）")
    ap.add_argument("--input", required=True, help="招标文件路径（.txt/.md/.docx）")
    ap.add_argument("--output", default="requirements.csv", help="输出 CSV 路径")
    ap.add_argument("--min-len", type=int, default=MIN_SENTENCE_LEN,
                    help=f"句子最短长度，默认 {MIN_SENTENCE_LEN}")
    args = ap.parse_args()

    src = Path(args.input)
    if not src.exists():
        sys.exit(f"[错误] 文件不存在：{src}")

    text = read_text(src)
    rows, candidates = extract(text, args.min_len)
    out = Path(args.output)
    write_csv(rows, out)

    mandatory_y = sum(1 for r in rows if r["mandatory"] == "Y")
    clarify_y = sum(1 for r in rows if r["need_clarify"] == "Y")
    by_type = {}
    for r in rows:
        by_type[r["type"]] = by_type.get(r["type"], 0) + 1

    print("=" * 56)
    print("要求抽取完成")
    print("=" * 56)
    print(f"输入文件   : {src}")
    print(f"输出文件   : {out.resolve()}")
    print(f"候选句总数 : {candidates}")
    print(f"生成条目数 : {len(rows)}")
    print(f"必答(Y)    : {mandatory_y}   非必答(N): {len(rows) - mandatory_y}")
    print(f"需澄清     : {clarify_y}")
    print("-" * 56)
    print("按类型分布：")
    for t in ["资格", "格式", "技术", "商务", "服务", "评分"]:
        if by_type.get(t):
            print(f"  {t}: {by_type[t]}")
    print("-" * 56)
    print("下一步（必须人工完成，脚本不做最终判定）：")
    print("  1. 逐条核对 source_quote 是否为原文照抄；")
    print("  2. 修正 mandatory —— 陈述句形式的强制条款（如“合同签订后60日内交货”）")
    print("     不含情态动词，脚本抓不到，必须人工补标 Y；")
    print("  3. 补齐 owner（责任人），把 need_clarify=Y 的整理成澄清函；")
    print(f"  4. 以「总条目 {len(rows)} / 必答 {mandatory_y}」作为后续覆盖率基准。")


if __name__ == "__main__":
    main()
