#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
validate_bid.py — 投标稿交付前合规校验。

用法：
    python validate_bid.py --bid ./bid --req requirements.csv
    python validate_bid.py --bid ./bid --req requirements.csv --map outline-map.csv
    python validate_bid.py --bid ./bid --list-todo        # 只列待确认项

检查项：
    [BLOCKER] 占位符残留       —— 存在未清理的 [待确认]/TODO/XXX/待补充 等
    [BLOCKER] 必答覆盖率       —— 存在 mandatory=Y 但未被任何章节引用的条目
    [BLOCKER] 灰色响应表述     —— 出现“基本响应/部分响应”等易被判不响应的表述
    [WARN]    章节编号连续性   —— 一级章节缺号、跳号、重号
    [WARN]    绝对化/注水词    —— 出现“100%/零故障/卓越/全方位”等风险表述
    [INFO]    篇幅统计

退出码：0 = 全部通过（可能有 WARN）；1 = 存在 BLOCKER；2 = 输入错误
"""

import argparse
import csv
import re
import sys
from pathlib import Path

TEXT_SUFFIXES = {".md", ".txt", ".markdown"}
SKIP_DIRS = {".git", "__pycache__", "node_modules", ".idea", ".vscode"}

# ---------------------------------------------------------------- 规则表

TODO_PATTERNS = [
    (r"\[待确认", "待确认标记"),
    (r"\[待商务确认", "待商务确认标记"),
    (r"\[待决策人批准", "待决策批准标记"),
    (r"\[通识表述", "通识表述待技术确认"),
    (r"\bTODO\b", "TODO"),
    (r"\bTBD\b", "TBD"),
    (r"\bXXX\b", "XXX 占位"),
    (r"待补充", "待补充"),
    (r"待填写", "待填写"),
    (r"_{4,}", "下划线空位"),
    (r"[？?]{2,}", "连续问号占位"),
]

GRAY_WORDS = [
    "基本响应", "基本满足", "部分响应", "原则上满足", "原则上响应",
    "大体满足", "视情况响应", "酌情满足", "基本符合", "勉强满足",
]

RISK_WORDS = [
    "100%", "零故障", "永不", "终身免费", "无限次", "彻底解决",
    "绝对安全", "万无一失", "卓越", "一流", "完美", "极致",
    "全方位", "深度融合", "业界领先", "国内领先",
]

DIFFUSION_WORDS = ["随着", "众所周知", "不言而喻", "势在必行", "日益凸显", "深入贯彻"]

NUM_UNIT_HINT = re.compile(r"\d+(?:\.\d+)?\s*(Gbps|Mbps|Tbps|路|台|个|套|年|天|小时|分钟|ms|%|TB|GB|PB)")


# ---------------------------------------------------------------- 读取

def read_any(path: Path) -> str:
    suffix = path.suffix.lower()
    if suffix == ".docx":
        try:
            import docx  # type: ignore
        except ImportError:
            return f"[跳过] {path.name}：读取 docx 需 python-docx\n"
        d = docx.Document(str(path))
        parts = [p.text for p in d.paragraphs]
        for t in d.tables:
            for row in t.rows:
                parts.append(" | ".join(c.text.strip() for c in row.cells))
        return "\n".join(parts)

    for enc in ("utf-8-sig", "utf-8", "gb18030", "gbk"):
        try:
            return path.read_text(encoding=enc)
        except UnicodeDecodeError:
            continue
    return ""


def collect(bid_path: Path):
    """返回 {文件名: 文本}"""
    files = {}
    if bid_path.is_file():
        files[bid_path.name] = read_any(bid_path)
        return files

    for p in sorted(bid_path.rglob("*")):
        if any(part in SKIP_DIRS for part in p.parts):
            continue
        if p.is_file() and p.suffix.lower() in (TEXT_SUFFIXES | {".docx"}):
            files[str(p.relative_to(bid_path))] = read_any(p)
    return files


def load_req(req_path: Path):
    for enc in ("utf-8-sig", "utf-8", "gb18030"):
        try:
            with req_path.open(encoding=enc, newline="") as f:
                rows = list(csv.DictReader(f))
            break
        except UnicodeDecodeError:
            continue
    else:
        sys.exit(f"[错误] 无法读取 {req_path}")

    if not rows:
        sys.exit(f"[错误] {req_path} 为空或缺少表头")
    need = {"req_id", "mandatory"}
    missing = need - set(rows[0].keys())
    if missing:
        sys.exit(f"[错误] requirements.csv 缺少列：{sorted(missing)}")
    return rows


# ---------------------------------------------------------------- 检查

def scan_patterns(files, patterns, label):
    """返回 [(文件名, 行号, 命中片段, 名称)]"""
    hits = []
    for fname, text in files.items():
        for lineno, line in enumerate(text.splitlines(), 1):
            for pat, name in patterns:
                if re.search(pat, line, flags=re.IGNORECASE):
                    snippet = line.strip()
                    if len(snippet) > 90:
                        idx = max(0, re.search(pat, line, flags=re.IGNORECASE).start() - 30)
                        snippet = line[idx:idx + 90].strip()
                    hits.append((fname, lineno, snippet, name))
    return hits


def check_coverage(files, req_rows):
    full_text = "\n".join(files.values())
    cited = set(m.upper() for m in re.findall(r"\bR-\d{3,}\b", full_text, flags=re.IGNORECASE))

    mandatory = [r for r in req_rows if (r.get("mandatory") or "").strip().upper() == "Y"]
    uncovered = [r for r in mandatory if r["req_id"].strip().upper() not in cited]
    extra = [c for c in cited if c not in {r["req_id"].strip().upper() for r in req_rows}]
    return cited, mandatory, uncovered, extra


def check_numbering(files):
    """检查一级章节编号连续性。
    只识别一级标题（单个 #），支持 `# 第3章 xxx` 与 `# 3. xxx` 两种写法；
    二级及以下标题（1.1 / 2.3）属条款级，不参与连续性判定。
    """
    issues = []
    for fname, text in files.items():
        nums = []
        for line in text.splitlines():
            s = line.strip()
            m = re.match(r"^#\s+(?:第\s*(\d+)\s*[章部篇]|(\d+)\s*[\.、]?(?:\s|$))", s)
            if m:
                nums.append(int(m.group(1) or m.group(2)))
        if len(nums) < 2:
            continue
        seen, prev = set(), 0
        for n in nums:
            if n in seen:
                issues.append((fname, f"一级章节编号重复：第 {n} 章"))
            elif prev > 0 and n > prev + 1:
                issues.append((fname, f"一级章节跳号：第 {prev} 章 → 第 {n} 章（缺 {n - prev - 1} 章）"))
            elif prev > 0 and n < prev:
                issues.append((fname, f"一级章节编号倒退：第 {prev} 章 → 第 {n} 章"))
            seen.add(n)
            prev = max(prev, n)
    return issues


def check_map(map_path: Path, req_rows):
    """校验 outline-map.csv：必答条目是否都有映射"""
    try:
        with map_path.open(encoding="utf-8-sig", newline="") as f:
            mrows = list(csv.DictReader(f))
    except Exception as e:  # noqa: BLE001
        return [f"无法读取映射表：{e}"]
    mapped = {r.get("req_id", "").strip().upper() for r in mrows}
    issues = []
    for r in req_rows:
        if (r.get("mandatory") or "").strip().upper() == "Y":
            if r["req_id"].strip().upper() not in mapped:
                issues.append(f"{r['req_id']}（必答）未在 outline-map.csv 中映射章节")
    for r in mrows:
        if not (r.get("chapter") or "").strip():
            issues.append(f"{r.get('req_id', '?')} 的 chapter 列（章节号）为空")
    return issues


# ---------------------------------------------------------------- 报告

def hr(ch="─", n=64):
    return ch * n


def main():
    ap = argparse.ArgumentParser(description="投标稿交付前合规校验")
    ap.add_argument("--bid", required=True, help="投标稿文件或目录")
    ap.add_argument("--req", help="requirements.csv 路径")
    ap.add_argument("--map", dest="map_file", help="outline-map.csv 路径（可选）")
    ap.add_argument("--list-todo", action="store_true", help="只列出待确认项")
    ap.add_argument("--strict", action="store_true", help="把 WARN 也视为失败")
    args = ap.parse_args()

    bid_path = Path(args.bid)
    if not bid_path.exists():
        sys.exit(f"[错误] 路径不存在：{bid_path}")

    files = collect(bid_path)
    if not files:
        sys.exit(f"[错误] 在 {bid_path} 下未找到可读的 .md/.txt/.docx 文件")

    # ---- 只列 TODO
    if args.list_todo:
        hits = scan_patterns(files, TODO_PATTERNS, "TODO")
        print(hr("="))
        print(f"待确认项清单（共 {len(hits)} 处）")
        print(hr("="))
        if not hits:
            print("✓ 未发现待确认标记")
        for fname, lineno, snippet, name in hits:
            print(f"[{name}] {fname}:{lineno}")
            print(f"    {snippet}")
        print(hr("="))
        return 0 if not hits else 1

    if not args.req:
        sys.exit("[错误] 需要 --req requirements.csv 才能做覆盖率检查")

    req_rows = load_req(Path(args.req))
    blockers, warns, infos = [], [], []

    # 1. 占位符
    todo_hits = scan_patterns(files, TODO_PATTERNS, "TODO")
    if todo_hits:
        blockers.append(("占位符残留", len(todo_hits), todo_hits))
    else:
        infos.append("占位符检查：通过（0 处）")

    # 2. 覆盖率
    cited, mandatory, uncovered, extra = check_coverage(files, req_rows)
    cov = (len(mandatory) - len(uncovered)) / len(mandatory) * 100 if mandatory else 100.0
    infos.append(f"必答覆盖率：{cov:.1f}%（{len(mandatory) - len(uncovered)}/{len(mandatory)}）")
    if uncovered:
        blockers.append(("必答条目未覆盖", len(uncovered),
                         [(r["req_id"], r.get("type", ""), r.get("requirement", "")[:70])
                          for r in uncovered]))
    if extra:
        warns.append(("引用了台账中不存在的条目号", len(extra), [(x,) for x in sorted(extra)]))

    # 3. 灰色表述
    gray_hits = scan_patterns(files, [(w, w) for w in GRAY_WORDS], "灰色表述")
    if gray_hits:
        blockers.append(("灰色响应表述（易被判不响应）", len(gray_hits), gray_hits))

    # 4. 章节编号
    num_issues = check_numbering(files)
    if num_issues:
        warns.append(("章节编号异常", len(num_issues), num_issues))

    # 5. 风险词 & 注水词
    risk_hits = scan_patterns(files, [(w, w) for w in RISK_WORDS], "风险表述")
    if risk_hits:
        warns.append(("绝对化/宣传性表述", len(risk_hits), risk_hits))
    diff_hits = scan_patterns(files, [(w, w) for w in DIFFUSION_WORDS], "注水词")
    if diff_hits:
        warns.append(("注水/套话表述", len(diff_hits), diff_hits))

    # 6. 映射表
    if args.map_file:
        map_issues = check_map(Path(args.map_file), req_rows)
        if map_issues:
            blockers.append(("章节映射缺失", len(map_issues), [(x,) for x in map_issues]))
        else:
            infos.append("章节映射：通过（必答条目全部已映射）")

    # 7. 篇幅
    total_chars = sum(len(t) for t in files.values())
    total_lines = sum(len(t.splitlines()) for t in files.values())
    num_count = len(NUM_UNIT_HINT.findall("\n".join(files.values())))
    infos.append(f"篇幅：{len(files)} 个文件 / {total_lines} 行 / {total_chars} 字符 / 量化表述 {num_count} 处")

    # ---- 输出
    print(hr("="))
    print("投标文件交付前校验报告")
    print(hr("="))
    print(f"稿件路径 : {bid_path.resolve()}")
    print(f"要求台账 : {Path(args.req).resolve()}（{len(req_rows)} 条）")
    print(hr())

    print("【统计】")
    for line in infos:
        print(f"  · {line}")

    if blockers:
        print(hr())
        print(f"【阻断项 BLOCKER】共 {sum(b[1] for b in blockers)} 处 —— 必须清零后才能交付")
        for title, count, items in blockers:
            print(f"\n  ▼ {title}（{count} 处）")
            for it in items[:30]:
                if len(it) == 4:      # (file, line, snippet, name)
                    print(f"    - [{it[3]}] {it[0]}:{it[1]}  {it[2]}")
                elif len(it) == 3:    # (req_id, type, requirement)
                    print(f"    - {it[0]} [{it[1]}] {it[2]}")
                else:
                    print(f"    - {it[0]}")
            if len(items) > 30:
                print(f"    … 另有 {len(items) - 30} 处未列出")

    if warns:
        print(hr())
        print(f"【提示 WARN】共 {sum(w[1] for w in warns)} 处 —— 建议人工确认")
        for title, count, items in warns:
            print(f"\n  ▼ {title}（{count} 处）")
            for it in items[:20]:
                if len(it) == 4:
                    print(f"    - [{it[3]}] {it[0]}:{it[1]}  {it[2]}")
                else:
                    print(f"    - {it[0] if len(it) == 1 else ' '.join(str(x) for x in it)}")
            if len(items) > 20:
                print(f"    … 另有 {len(items) - 20} 处未列出")

    print(hr("="))
    failed = bool(blockers) or (args.strict and bool(warns))
    if failed:
        print("结论：不通过 ✗")
        print("  请先清零全部 BLOCKER，再重跑本脚本。")
        if warns:
            print("  WARN 项不阻断，但需人工判断是否保留（绝对化承诺须决策人批准）。")
    else:
        print("结论：通过 ✓")
        print("  脚本已过，但脚本只查形式。仍需按 compliance-checklist.md 做人工复核——")
        print("  尤其要确认「答非所问」与「数值方向」，这两类脚本查不出来。")
    print(hr("="))
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
