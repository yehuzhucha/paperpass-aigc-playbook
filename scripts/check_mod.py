#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
计算两版文本的字符级修改率。

用法:
    python3 check_mod.py 原文.txt 改后.docx
    python3 check_mod.py 旧版.docx 新版.docx

经验值（本项目实测）:
    <20%   基本无效，AIGC 分数不会动
    20-40% 有效，但不如预期
    >40%   大幅变化（注意：整篇重写反而可能涨分，见 README）
"""
import sys
import re
import difflib


def load(path):
    if path.endswith(".docx"):
        import docx
        return "".join(p.text for p in docx.Document(path).paragraphs)
    return open(path, encoding="utf-8").read()


def clean(text):
    return re.sub(r"\s", "", text)


def main():
    if len(sys.argv) < 3:
        sys.exit(__doc__)
    a, b = clean(load(sys.argv[1])), clean(load(sys.argv[2]))
    ratio = difflib.SequenceMatcher(None, a, b).ratio()
    mod = (1 - ratio) * 100

    if mod < 20:
        verdict = "❌ 太低，检测器不会认，别浪费钱去查"
    elif mod < 40:
        verdict = "⚠ 有效但不强，可能只降 5~10 个点"
    else:
        verdict = "✅ 修改率足够（但注意：整篇重写可能反而涨分）"

    print(f"原文字符 {len(a)}  改后字符 {len(b)}")
    print(f"相似度 {ratio*100:.1f}%   修改率 {mod:.1f}%")
    print(verdict)


if __name__ == "__main__":
    main()
