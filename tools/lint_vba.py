# -*- coding: utf-8 -*-
"""Статический контроль VBA: баланс блоков, скобок, кавычек, меток."""
import os
import re
import sys

SRC = os.path.join(os.path.dirname(__file__), "..", "gms-source")


def logical_lines(text):
    """склейка строк с переносом ' _', удаление строковых литералов и комментариев"""
    raw = text.replace("\r\n", "\n").split("\n")
    joined, buf = [], ""
    for ln in raw:
        if buf:
            ln = buf + " " + ln.strip()
            buf = ""
        if ln.rstrip().endswith(" _"):
            buf = ln.rstrip()[:-2].rstrip()
            continue
        joined.append(ln)
    if buf:
        joined.append(buf)

    out = []
    for ln in joined:
        # вырезаем строковые литералы ("" -> один литерал)
        code = re.sub(r'"(?:""|[^"])*"', '""', ln)
        # комментарии (апостроф), кроме случаев внутри строк - строки уже вырезаны
        idx = code.find("'")
        if idx >= 0:
            code = code[:idx]
        out.append(code)
    return out


fail = False
# Вызовы без аргументов на отдельной строке должны ссылаться на процедуру
# проекта. Это не полноценный компилятор VBA и не проверяет вызовы с аргументами.
modules = [fn for fn in sorted(os.listdir(SRC)) if fn.endswith('.bas')]
procedures = set()
for fn in modules:
    text = open(os.path.join(SRC, fn), encoding='utf-8').read()
    procedures.update(name.lower() for name in re.findall(
        r'^\s*(?:(?:Private|Public|Friend)\s+)?(?:Sub|Function)\s+(\w+)',
        text, re.I | re.M))
standalone_keywords = {'else', 'loop', 'wend', 'next', 'do', 'stop', 'return'}

for fn in sorted(os.listdir(SRC)):
    if not fn.endswith(".bas"):
        continue
    text = open(os.path.join(SRC, fn), encoding="utf-8").read()
    lines = logical_lines(text)
    print(f"== {fn}: {len(lines)} логических строк")

    def count(pat, flags=re.I):
        rx = re.compile(pat, flags)
        return sum(len(rx.findall(ln)) for ln in lines)

    pairs = [
        (r"(?<!End )(?<!Exit )\bSub\b", r"\bEnd Sub\b", "Sub"),
        (r"(?<!End )(?<!Exit )\bFunction\b", r"\bEnd Function\b", "Function"),
        (r"\bFor (Each )?", r"(?<!Resume )\bNext\b", "For"),
        (r"^\s*Do\b|\bDo\b", r"\bLoop\b", "Do"),
        (r"\bSelect Case\b", r"\bEnd Select\b", "Select"),
        (r"(?<!End )\bWith\b", r"\bEnd With\b", "With"),
    ]
    for op, cl, name in pairs:
        no, nc = count(op), count(cl)
        if no != nc:
            print(f"  ДИСБАЛАНС {name}: {no} <> {nc}")
            fail = True

    # блочные If (оканчиваются на Then) против End If
    no = sum(1 for ln in lines if re.search(r"\bIf\b.*\bThen\s*$", ln, re.I))
    nc = count(r"\bEnd If\b")
    if no != nc:
        print(f"  ДИСБАЛАНС If: блочных {no} <> End If {nc}")
        fail = True

    for i, ln in enumerate(lines, 1):
        standalone = re.fullmatch(r'\s*([A-Za-z]\w*)\s*', ln)
        if standalone:
            name = standalone.group(1).lower()
            if name not in procedures and name not in standalone_keywords:
                print(f'  строка {i}: возможный неопределённый вызов: {standalone.group(1)}')
                fail = True
        if re.search(r':\s*\w+\s*=\s*[^:]+,\s*\w+\s+As\s+\w+', ln, re.I):
            print(f'  строка {i}: объявление после присваивания: {ln.strip()}')
            fail = True
        if ln.count("(") != ln.count(")"):
            print(f"  строка {i}: скобки: {ln.strip()[:90]}")
            fail = True
        if ln.count('"') % 2 != 0:
            print(f"  строка {i}: кавычки: {ln.strip()[:90]}")
            fail = True

    gotos = re.findall(r"GoTo\s+(\w+)", text)
    labels = set(re.findall(r"^\s*([A-Za-z]\w*):\s*$", text, re.M))
    for g in gotos:
        if g != "0" and g not in labels:
            print(f"  переход на несуществующую метку: {g}")
            fail = True

print("ИТОГ:", "ОШИБКИ" if fail else "OK")
sys.exit(1 if fail else 0)
