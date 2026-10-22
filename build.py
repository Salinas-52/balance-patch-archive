"""xlsx -> index.html 빌드 스크립트.

사용법:  python build.py [xlsx경로]
- 원본 xlsx 의 값은 가공하지 않고 그대로 JSON 으로 옮긴다.
- 속성 표기 이상(공통, -, 빈 값 등)은 변환하지 않고 flag 로만 표시해 화면에서 처리한다.
"""
import json
import re
import sys
from pathlib import Path

import openpyxl

ROOT = Path(__file__).parent
SRC = Path(sys.argv[1]) if len(sys.argv) > 1 else ROOT / "Balance_patch_notes_archive_v2609.xlsx"

TITLE_RE = re.compile(r"밸런스\s*패치")
ATTRS = ["물", "불", "바람", "빛", "어둠"]


def s(v):
    return None if v is None else str(v)


def split_attr_rows(mon, attr, skill):
    """속성/몬스터/스킬이 줄바꿈으로 여러 개 묶인 행(2024.9 뱀파이어)을 분리한다.
    세 칸의 줄 수가 같고 속성 줄이 2개 이상일 때만 분리하고, 아니면 원본 그대로 둔다."""
    if not (mon and attr and skill and "\n" in attr):
        return None
    m, a, k = mon.split("\n"), attr.split("\n"), skill.split("\n")
    if len(m) == len(a) == len(k) and len(a) >= 2:
        return list(zip(m, a, k))
    return None


def load(path):
    wb = openpyxl.load_workbook(path)

    collab = []
    for r in wb["Collab"].iter_rows(min_row=1):
        c, d = r[2].value, r[3].value
        if c and d:
            collab.append([str(c), str(d)])

    tabs = []
    report = []
    for ws in wb.worksheets[3:]:
        groups = []

        def new_group(title):
            g = {"title": title, "sections": []}
            groups.append(g)
            return g

        group = None
        section = None

        def cur_section(name=""):
            nonlocal group, section
            if group is None:
                group = new_group(None)
            section = {"name": name, "rows": []}
            group["sections"].append(section)
            return section

        for r in ws.iter_rows(min_row=1):
            a, b, c = s(r[0].value), s(r[1].value), s(r[2].value)
            i, j = s(r[8].value), s(r[9].value)
            if a == "일정":
                continue
            row_no = r[0].row
            if a and b and (c is not None or i or j or a[:2] == "20"):
                # 데이터 행
                if group is None or section is None:
                    cur_section("")
                parts = split_attr_rows(b, c, i)
                if parts:
                    for mm, aa, kk in parts:
                        section["rows"].append(dict(d=a, m=mm, a=aa, k=kk, t=j, src=f"{ws.title}!{row_no}", split=True))
                    report.append(f"분리: {ws.title} {row_no}행 -> {len(parts)}개 몬스터")
                else:
                    section["rows"].append(dict(d=a, m=b, a=c, k=i, t=j, src=f"{ws.title}!{row_no}"))
                continue
            text = a or b
            if text and c is None and not i and not j:
                if TITLE_RE.search(text):
                    group = new_group(text)
                    section = None
                else:
                    cur_section(text)
                continue
            if any(x is not None for x in (a, b, c, i, j)):
                report.append(f"미분류 행: {ws.title} {row_no}행 {a!r} {b!r} {c!r}")
        # 제목 없는 첫 그룹은 비어 있으면 제거
        groups = [g for g in groups if any(sec["rows"] for sec in g["sections"])]
        tabs.append({"id": ws.title, "state": ws.sheet_state, "groups": groups})
    return collab, tabs, report


def main():
    collab, tabs, report = load(SRC)
    # 탭 id 정규화 없이 원본 이름 유지, 정렬키만 별도 계산
    for t in tabs:
        y, m = t["id"].split(".")
        t["key"] = f"{int(y):04d}.{int(m):02d}"
    tabs.sort(key=lambda t: t["key"], reverse=True)

    n = sum(len(sec["rows"]) for t in tabs for g in t["groups"] for sec in g["sections"])
    data = {"collab": collab, "tabs": tabs}
    blob = json.dumps(data, ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")
    tpl = (ROOT / "template.html").read_text(encoding="utf-8")
    out = tpl.replace("/*__DATA__*/null", blob)
    (ROOT / "index.html").write_text(out, encoding="utf-8")
    print(f"탭 {len(tabs)}개, 변경 행 {n}개, 콜라보 {len(collab)}쌍 -> index.html ({len(out)//1024}KB)")
    for line in report:
        print(" ", line)


if __name__ == "__main__":
    main()
