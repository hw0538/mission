"""SAR System Check worst 선정 도구 (sar_worst.html 과 같은 로직).

시스템 체크 엑셀(들)을 읽어 정해진 양식의 새 엑셀을 만들고,
Serial 번호가 같은 측정 건끼리 비교해 worst 의 Plot No. 칸에 번호를 붙인다.

결과 양식 (측정 건마다 1g·10g 줄, 원본에 8g·APD 가 있으면 그 줄도 같은 양식으로 추가, 날짜는 yyyy-mm-dd)
  Date Tested | System Dipole(Type, Serial #) | T.S. Liquid(Head, 1g/10g) |
  Measured Results(Zoom Scan to 100 mW, Normalize to 1 W) | Target | Delta ±10 % | Plot No.

읽을 수 있는 원본
  - 측정 건 = '1g' 행부터 라벨(8g/10g/APD…)이 이어지는 행까지 → 2줄·4줄 양식 모두 가능
  - 열은 헤더 이름으로 찾고, Delta 는 (A-B)/B*100 계산이 맞는 열로 찾음
  - 여러 파일·양식이 섞여도 하나의 표로 모으고, worst 는 전체를 합쳐서 Serial 별로 판정
    (Serial 은 대소문자·공백 무시, 비어 있으면 판정 제외)

worst 기준
  1) |1g Delta %| 가 가장 큰 데이터
  2) 1g 가 동률이면 |10g Delta %| 가 큰 데이터
  3) 둘 다 동률이면 먼저 측정된 데이터 + 동률 메모
  (표에 보이는 소수 둘째 자리 기준, 부호 무시)

사용법
  python sar_worst.py                      # ./input 폴더의 엑셀 전부 → ./output/SystemCheck_Worst.xlsx
  python sar_worst.py a.xlsx b.xlsx        # 파일 지정
  python sar_worst.py input -o result.xlsx # 폴더 지정 + 출력 경로 지정
  python sar_worst.py --order freq         # 번호를 주파수 오름차순으로
"""

import argparse
import datetime as dt
import json
import re
import sys
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

HERE = Path(__file__).resolve().parent
WARN = []


# ================================================================ 값 정리

def text(v):
    if v is None:
        return ""
    if isinstance(v, (dt.datetime, dt.date)):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip()


def low(v):
    return text(v).lower()


def label(v):
    return re.sub(r"\s", "", low(v))


def num(v):
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    if v is None or isinstance(v, (dt.datetime, dt.date)):
        return None
    s = text(v).replace("%", "")
    return float(s) if re.fullmatch(r"[+-]?\d*\.?\d+(e[+-]?\d+)?", s, re.I) else None


def parse_date(v):
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    if isinstance(v, (int, float)) and 20000 < v < 80000:
        return dt.date(1899, 12, 30) + dt.timedelta(days=int(v))
    s = text(v)
    for pat, order in ((r"(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})", (3, 1, 2)),
                       (r"(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})", (1, 2, 3))):
        m = re.fullmatch(pat, s)
        if m:
            try:
                return dt.date(int(m[order[0]]), int(m[order[1]]), int(m[order[2]]))
            except ValueError:
                return None
    return None


def fmt_date(b):
    return b["date"].strftime("%Y-%m-%d") if b["date"] else text(b["date_raw"])  # yyyy-mm-dd


def date_key(b):
    return b["date"].toordinal() if b["date"] else 10 ** 9


def r2(x):
    return round(abs(x) * 100 + 1e-9) / 100


def band_key(t):
    """파일마다 표기가 조금 달라도(D835V2 / D835 V2 / d835-v2) 같은 대역으로 묶음"""
    return re.sub(r"[\s\-_]", "", text(t)).upper()


def serial_key(s):
    """Serial 비교용: 대소문자·공백 무시 (4d174 = 4D174), 숫자로 저장된 1205.0 → 1205"""
    return re.sub(r"\s", "", re.sub(r"\.0+$", "", text(s))).upper()


def freq_of(t):
    s = t.upper()
    m = re.search(r"(\d+(?:\.\d+)?)\s*GHZ", s)
    if m:
        return float(m[1]) * 1000
    m = re.search(r"(\d+(?:\.\d+)?)", s)
    return float(m[1]) if m else float("inf")


# ================================================================ 읽기

def read_sheet(ws, file):
    source = f"{file}/{ws.title}"
    max_r, max_c = ws.max_row, ws.max_column
    master = {}
    for m in ws.merged_cells.ranges:
        for r in range(m.min_row, m.max_row + 1):
            for c in range(m.min_col, m.max_col + 1):
                master[(r, c)] = (m.min_row, m.min_col)

    def get(r, c):
        if not c:
            return None
        rr, cc = master.get((r, c), (r, c))
        return ws.cell(rr, cc).value

    # 1) 라벨 열과 첫 측정 행
    first = label_col = None
    for r in range(1, max_r + 1):
        for c in range(1, max_c + 1):
            if label(get(r, c)) == "1g":
                first, label_col = r, c
                break
        if first:
            break
    if not first:
        return []

    def lab(r):
        return label(get(r, label_col))

    # 2) block 나누기
    spans, r = [], first
    while r <= max_r:
        if lab(r) != "1g":
            r += 1
            continue
        end = r
        while end + 1 <= max_r and lab(end + 1) and lab(end + 1) != "1g":
            end += 1
        spans.append((r, end))
        r = end + 1

    def row10(s):
        return next((x for x in range(s[0], s[1] + 1) if lab(x) == "10g"), None)

    # 3) 헤더 (첫 측정 행 바로 위 최대 8줄)
    top = max(1, first - 8)

    def head(c):
        return " ".join(low(get(r, c)) for r in range(top, first))

    col, delta_cand = {}, []
    for c in range(1, max_c + 1):
        if c == label_col:
            continue
        h = head(c)
        checks = {
            "date": "date" in h,
            "type": re.search(r"\btype\b", h),
            "serial": "serial" in h,
            "zoom": "zoom" in h and not re.search(r"ratio|est\.", h),
            "norm": "normaliz" in h,
            "target": "target" in h,
            "liquid": "liquid" in h or re.search(r"\beli\b", h),
        }
        for k, ok in checks.items():
            if k not in col and ok:
                col[k] = c
        if "delta" in h:
            delta_cand.append(c)
    r1 = [s[0] for s in spans]
    if "date" not in col:
        col["date"] = next((c for c in range(1, max_c + 1) if all(parse_date(get(r, c)) for r in r1)), None)
    if "type" not in col:
        col["type"] = next((c for c in range(1, max_c + 1)
                            if all(re.fullmatch(r"D\s*\d[\w.\s-]*", text(get(r, c)), re.I) for r in r1)), None)
    if "liquid" not in col and label_col > 1 and all(re.match(r"[a-z]", text(get(r, label_col - 1)), re.I) for r in r1):
        col["liquid"] = label_col - 1

    # 4) Delta 열: (A-B)/B*100 계산이 맞는 열. 맞는 A,B 는 Normalize·Target 으로도 사용
    check_rows = [x for s in spans for x in (s[0], row10(s)) if x]

    def fit_pair(r, c):
        v = num(get(r, c))
        if v is None:
            return None
        nums = [(k, num(get(r, k))) for k in range(1, max_c + 1) if k not in (c, label_col)]
        nums = [(k, x) for k, x in nums if x is not None]
        for ka, a in nums:
            for kb, b in nums:
                if ka != kb and b != 0 and abs((a - b) / b * 100 - v) <= 0.015 + abs(v) * 0.002:
                    return (ka, kb)
        return None

    def score(c):
        pairs, ok = {}, 0
        for r in check_rows:
            p = fit_pair(r, c)
            if p:
                ok += 1
                pairs[p] = pairs.get(p, 0) + 1
        best = max(pairs.items(), key=lambda kv: kv[1])[0] if pairs else None
        return {"c": c, "ok": ok, "pair": best}

    cands = sorted((score(c) for c in delta_cand), key=lambda d: -d["ok"])
    delta = cands[0] if cands else None
    if not delta or delta["ok"] == 0:
        alls = sorted((score(c) for c in range(1, max_c + 1) if c != label_col), key=lambda d: -d["ok"])
        if alls and alls[0]["ok"] >= -(-len(check_rows) // 2):
            delta = alls[0]
    missing = [n for n, ok in (("날짜", col.get("date")), ("Type", col.get("type")),
                               ("Delta", delta and (delta["ok"] or delta_cand))) if not ok]
    if missing:
        WARN.append(f"{source}: {', '.join(missing)} 열을 찾지 못해 건너뜀")
        return []
    col["delta"] = delta["c"]
    if delta["pair"]:
        col.setdefault("norm", delta["pair"][0])
        col.setdefault("target", delta["pair"][1])

    def vals(r):
        o = {k: num(get(r, col.get(k))) for k in ("zoom", "norm", "target", "delta")}
        if o["delta"] is None and o["norm"] is not None and o["target"]:
            o["delta"] = (o["norm"] - o["target"]) / o["target"] * 100
        return o

    out = []
    for s in spans:
        rt = row10(s)
        if not rt:
            WARN.append(f"{source} {s[0]}행: 10g 행이 없어 건너뜀")
            continue
        b = {"file": file, "source": source, "row": s[0],
             "date_raw": get(s[0], col["date"]), "date": parse_date(get(s[0], col["date"])),
             "type": text(get(s[0], col["type"])), "serial": text(get(s[0], col.get("serial"))),
             "liquid": text(get(s[0], col.get("liquid"))), "g1": vals(s[0]), "g10": vals(rt),
             # 1g·8g·10g·APD… 원본 순서 그대로 (결과에 모두 표시)
             "rows": [{"lab": text(get(x, label_col)), **vals(x)} for x in range(s[0], s[1] + 1)]}
        if not b["type"]:
            WARN.append(f"{source} {s[0]}행: Type 이 비어 있어 건너뜀")
            continue
        if not col.get("serial"):
            if not out:
                WARN.append(f"{source}: Serial 열을 찾지 못해 worst 판정에서 제외 (표에는 표시)")
        elif not b["serial"]:
            WARN.append(f"{source} {s[0]}행: Serial 이 비어 있어 worst 판정에서 제외 (표에는 표시)")
        if b["g1"]["delta"] is None or b["g10"]["delta"] is None:
            WARN.append(f"{source} {s[0]}행: Delta 값을 구할 수 없어 건너뜀")
            continue
        b["d1"], b["d10"] = b["g1"]["delta"], b["g10"]["delta"]
        b["sig"] = json.dumps([date_key(b), band_key(b["type"]), b["serial"], b["rows"]])
        out.append(b)
    return out


def collect_files(paths):
    files = []
    for p in map(Path, paths):
        if p.is_dir():
            files += sorted(f for f in p.iterdir()
                            if f.suffix.lower() in (".xlsx", ".xlsm") and not f.name.startswith("~$"))
        elif p.is_file():
            files.append(p)
        else:
            WARN.append(f"경로 없음: {p}")
    return files


# ================================================================ worst 선정

def analyze(records, order):
    records = sorted(records, key=date_key)  # 안정 정렬: 같은 날은 읽은 순서
    pos = {id(b): i for i, b in enumerate(records)}
    # Serial 번호가 같은 것끼리 묶어서 worst 선정
    groups = {}
    for b in records:
        k = serial_key(b["serial"])
        if k:  # Serial 이 비어 있으면 판정 제외 (읽을 때 경고)
            groups.setdefault(k, []).append(b)
    worst = []
    for serial, recs in groups.items():
        best = recs[0]
        for x in recs:  # 완전 동률이면 먼저 측정된 것 유지
            if (r2(x["d1"]), r2(x["d10"])) > (r2(best["d1"]), r2(best["d10"])):
                best = x
        key = (r2(best["d1"]), r2(best["d10"]))
        tied = [x for x in recs if x is not best and (r2(x["d1"]), r2(x["d10"])) == key]
        same1g = [x for x in recs if x is not best and r2(x["d1"]) == key[0]]
        why = ("1g·10g 완전 동률 → 먼저 측정된 데이터" if tied else
               "1g 동률 → 10g 로 결정" if same1g else
               "측정 1회" if len(recs) == 1 else "1g 최대")
        worst.append({"serial": serial, "rec": best, "tied": tied, "count": len(recs),
                      "n_files": len({x["file"] for x in recs}), "why": why})
    if order == "freq":
        worst.sort(key=lambda w: (freq_of(w["rec"]["type"]), w["serial"]))
    else:
        worst.sort(key=lambda w: pos[id(w["rec"])])
    no = {}
    for k, w in enumerate(worst, 1):
        w["no"] = k
        no[id(w["rec"])] = w
    return records, worst, no


# ================================================================ 쓰기 (사진 양식)

THIN = Side(style="thin", color="000000")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HEAD_FILL = PatternFill("solid", fgColor="D8CFB0")
FONT = Font(name="Arial", size=10)
BOLD = Font(name="Arial", size=10, bold=True)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)


def box(ws, rng, value=None, font=FONT, fill=None, fmt=None):
    if ":" in rng:
        ws.merge_cells(rng)
    for row in ws[rng] if ":" in rng else [[ws[rng]]]:
        for c in row:
            c.border, c.alignment, c.font = BOX, CENTER, font
            if fill:
                c.fill = fill
            if fmt:
                c.number_format = fmt
    ws[rng.split(":")[0]].value = value


def write_main(ws, records, no):
    for rng, v in [("A1:A2", "Date Tested"), ("B1:C1", "System Dipole"), ("B2", "Type"), ("C2", "Serial #"),
                   ("D1:E2", "T.S.\nLiquid"), ("F1:G1", "Measured Results"),
                   ("F2", "Zoom Scan to\n100 mW"), ("G2", "Normalize\nto 1 W"),
                   ("H1:H2", "Target\n(Ref. Value)"), ("I1:I2", "Delta\n±10 %"), ("J1:J2", "Plot No.")]:
        box(ws, rng, v, fill=HEAD_FILL)
    ws.row_dimensions[1].height = 20
    ws.row_dimensions[2].height = 30
    for col, w in zip("ABCDEFGHIJ", (16, 13, 11, 9, 7, 12, 12, 12, 10, 12)):
        ws.column_dimensions[col].width = w

    r = 3
    for rec in records:
        w = no.get(id(rec))
        n = len(rec["rows"])

        def span(col):
            return f"{col}{r}:{col}{r + n - 1}" if n > 1 else f"{col}{r}"

        box(ws, span("A"), fmt_date(rec))
        box(ws, span("B"), rec["type"])
        box(ws, span("C"), rec["serial"])
        box(ws, span("D"), rec["liquid"])
        for i, g in enumerate(rec["rows"]):  # 1g·8g·10g·APD… 행 모두 같은 양식으로
            rr = r + i
            box(ws, f"E{rr}", g["lab"])
            box(ws, f"F{rr}", g["zoom"], fmt="0.00")
            box(ws, f"G{rr}", g["norm"], fmt="0.0")
            box(ws, f"H{rr}", g["target"], fmt="0.00")
            box(ws, f"I{rr}", None if g["delta"] is None else round(g["delta"], 2), fmt="0.00")
        box(ws, span("J"), w["no"] if w else None, font=BOLD if w else FONT)
        if w and w["tied"]:
            ws[f"J{r}"].comment = Comment("1g·10g Delta 완전 동률: " + ", ".join(fmt_date(x) for x in w["tied"])
                                          + " 데이터와 같음 (먼저 측정된 데이터 선택)", "sar_worst")
        r += n
    ws.freeze_panes = "A3"


def write_summary(ws, worst):
    heads = ["Plot No.", "Serial #", "Type", "Date Tested", "1g Delta (%)", "10g Delta (%)",
             "측정 횟수", "비교 파일 수", "선정 근거", "원본"]
    for c, h in enumerate(heads, 1):
        box(ws, ws.cell(1, c).coordinate, h, font=BOLD, fill=HEAD_FILL)
    for i, w in enumerate(worst, 2):
        rec = w["rec"]
        vals = [w["no"], rec["serial"], rec["type"], fmt_date(rec), round(rec["d1"], 2), round(rec["d10"], 2),
                w["count"], w["n_files"], w["why"], f'{rec["source"]} ({rec["row"]}행)']
        for c, v in enumerate(vals, 1):
            box(ws, ws.cell(i, c).coordinate, v, fmt="0.00" if c in (5, 6) else None)
    for col, w in zip("ABCDEFGHIJ", (9, 10, 12, 13, 12, 13, 10, 11, 30, 40)):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A2"


# ================================================================ main

def main():
    ap = argparse.ArgumentParser(description="SAR System Check Serial별 worst 넘버링")
    ap.add_argument("inputs", nargs="*", default=[str(HERE / "input")], help="엑셀 파일 또는 폴더")
    ap.add_argument("-o", "--output", default=str(HERE / "output" / "SystemCheck_Worst.xlsx"))
    ap.add_argument("--order", choices=["table", "freq"], default="table",
                    help="번호 순서: table=표 순서(날짜순), freq=주파수 오름차순")
    args = ap.parse_args()

    files = collect_files(args.inputs)
    if not files:
        sys.exit("읽을 엑셀 파일이 없습니다. input 폴더에 넣거나 파일 경로를 지정하세요.")

    records, seen = [], set()
    for f in files:
        print(f"읽는 중: {f.name}")
        try:
            wb = load_workbook(f, data_only=True)
        except Exception as e:  # noqa: BLE001
            WARN.append(f"{f.name}: 파일을 열 수 없어요 ({e})")
            continue
        added = dup = 0
        for ws in wb.worksheets:
            for b in read_sheet(ws, f.name):
                if b["sig"] in seen:
                    dup += 1
                    continue
                seen.add(b["sig"])
                records.append(b)
                added += 1
        if not added and not dup:
            WARN.append(f"{f.name}: 시스템 체크 데이터를 찾지 못했어요 ('1g' 행과 날짜·Type·Delta 열 확인)")
        if dup:
            WARN.append(f"{f.name}: 이미 불러온 데이터와 같은 {dup}건은 한 번만 반영했어요")

    for w in WARN:
        print("  ! " + w)
    if not records:
        sys.exit("시스템 체크 데이터를 찾지 못했습니다.")

    records, worst, no = analyze(records, args.order)
    wb = Workbook()
    ws = wb.active
    ws.title = "System Check"
    write_main(ws, records, no)
    write_summary(wb.create_sheet("Worst Summary"), worst)

    out = Path(args.output)
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out)

    print(f"\n파일 {len(files)}개의 데이터 {len(records)}건을 합쳐서 비교 → Serial {len(worst)}개")
    for w in worst:
        rec = w["rec"]
        print(f"  Plot {w['no']:>2}: Serial {rec['serial']:<8} {rec['type']:<11} {fmt_date(rec):<10} "
              f"1g {rec['d1']:+6.2f}%  10g {rec['d10']:+6.2f}%  "
              f"({w['n_files']}개 파일 {w['count']}건 중, 원본 {rec['source']})"
              + ("  ※완전 동률" if w["tied"] else ""))
    print(f"\n저장: {out}")


if __name__ == "__main__":
    main()
