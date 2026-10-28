#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
解析 PaperPass「Word 标红版 AIGC 检测报告」，输出每段颜色与字数统计。

用法:
    python3 parse_report.py 报告.docx              # 汇总 + 逐段
    python3 parse_report.py 报告.docx --only-flag  # 只看被标色的段
    python3 parse_report.py 报告.docx --csv out.csv

颜色含义（PaperPass）:
    F12828  红  疑似度 ≥70%
    F39800  橙  60% ~ 70%
    9D91E9  紫  50% ~ 60%
    B0B0B0  灰  过短片段 / 标题 / 参考文献，不计入检测
    无颜色       <50%，通过
"""
import sys
import csv
import argparse
import collections

try:
    import docx
    from docx.oxml.ns import qn
except ImportError:
    sys.exit("需要 python-docx：pip install python-docx")

COLORS = {
    "F12828": "红 ≥70%",
    "F39800": "橙 60-70%",
    "9D91E9": "紫 50-60%",
    "B0B0B0": "灰 不计",
    "NONE": "通过 <50%",
}


def parse(path):
    doc = docx.Document(path)
    rows = []
    for para in doc.paragraphs:
        text = para.text.strip()
        if len(text) < 15:
            continue
        counter = collections.Counter()
        for run in para.runs:
            rpr = run._element.rPr
            if rpr is None:
                continue
            node = rpr.find(qn("w:color"))
            if node is None:
                continue
            val = node.get(qn("w:val"))
            if val in (None, "auto", "000000"):
                continue
            counter[val] += len(run.text)
        color = counter.most_common(1)[0][0] if counter else "NONE"
        rows.append({"color": color, "chars": len(text), "text": text,
                     "colored": sum(counter.values())})
    return rows


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("report")
    ap.add_argument("--only-flag", action="store_true", help="只输出被标色的段落")
    ap.add_argument("--csv", help="导出 CSV")
    args = ap.parse_args()

    rows = parse(args.report)
    by_color = collections.Counter(r["color"] for r in rows)
    by_chars = collections.Counter()
    for r in rows:
        by_chars[r["color"]] += r["chars"]

    print("=" * 60)
    print("段落统计")
    for c, n in by_color.most_common():
        print(f"  {COLORS.get(c, c):<10} {n:>3} 段   {by_chars[c]:>5} 字")

    total = sum(by_chars.values())
    flagged = sum(v for k, v in by_chars.items() if k not in ("NONE", "B0B0B0"))
    print(f"\n  正文合计 {total} 字，被标色 {flagged} 字（{flagged/total*100:.1f}%）")

    # 段落长度规律：这是本项目最重要的经验之一
    flag_len = [r["chars"] for r in rows if r["color"] in ("F12828", "F39800", "9D91E9")]
    ok_len = [r["chars"] for r in rows if r["color"] == "NONE"]
    if flag_len:
        print(f"  被标色段平均 {sum(flag_len)/len(flag_len):.0f} 字"
              f"（{min(flag_len)}~{max(flag_len)}）")
    if ok_len:
        print(f"  通过段平均   {sum(ok_len)/len(ok_len):.0f} 字"
              f"（{min(ok_len)}~{max(ok_len)}）")
        if sum(flag_len)/max(len(flag_len), 1) < sum(ok_len)/len(ok_len):
            print("  ⚠ 被标色的段明显更短 → 考虑把短段合并到 150 字以上")

    print("\n" + "=" * 60)
    print("逐段明细")
    for i, r in enumerate(rows):
        if args.only_flag and r["color"] in ("NONE", "B0B0B0"):
            continue
        tag = COLORS.get(r["color"], r["color"])
        print(f"[{i:02d}] {tag:<10} {r['chars']:>4}字  {r['text'][:60]}")

    if args.csv:
        with open(args.csv, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.DictWriter(f, fieldnames=["color", "chars", "text"])
            w.writeheader()
            for r in rows:
                w.writerow({k: r[k] for k in ("color", "chars", "text")})
        print(f"\nCSV 已写入 {args.csv}")


if __name__ == "__main__":
    main()
