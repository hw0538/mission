"""SAR System Check worst 선정 도구 (sar_worst.html 과 같은 로직).

시스템 체크 엑셀(들)을 읽어 파일마다 정해진 양식의 시트로 정리한 새 엑셀을 만들고,
Serial 번호가 같은 측정 건끼리 비교해 worst 의 Plot No. 칸에 번호를 붙인다.

결과 양식 (측정 건마다 1g·10g 줄, 원본에 8g·APD 가 있으면 그 줄도 같은 양식으로 추가, 날짜는 yyyy-mm-dd)
  Date Tested | System Dipole(Type, Serial #) | T.S. Liquid(Head, 1g/10g) |
  Measured Results(Zoom Scan to 100 mW, Normalize to 1 W) | Target | Delta ±10 % | Plot No.

읽을 수 있는 원본
  - 측정 건 = '1g' 행부터 라벨(8g/10g/APD…)이 이어지는 행까지 → 2줄·4줄 양식 모두 가능
  - 열은 헤더 이름으로 찾고, Delta 는 (A-B)/B*100 계산이 맞는 열로 찾음
  - 엑셀은 합치지 않고 파일마다 결과 시트를 하나씩 만들고, worst 는 모든 파일의 데이터 값을 비교해 Serial 별로 판정
    (여러 파일에 같은 측정 건이 있으면 각 시트에는 표시, 비교는 한 번만)
    (Serial 은 대소문자·공백 무시, 비어 있으면 판정 제외)

worst 기준
  1) |1g Delta %| 가 가장 큰 데이터
  2) 1g 가 동률이면 |10g Delta %| 가 큰 데이터
  3) 둘 다 동률이면 먼저 측정된 데이터 + 동률 메모
  (표에 보이는 소수 둘째 자리 기준, 부호 무시)
  4) 정리 후 번호가 하나도 없는 파일은 그 파일 전체(Serial 무관)에서 worst 하나를 골라 번호 추가

사용법
  python sar_worst.py                      # ./input 폴더의 엑셀 전부 → ./output/SystemCheck_Worst.xlsx
  python sar_worst.py a.xlsx b.xlsx        # 파일 지정
  python sar_worst.py input -o result.xlsx # 폴더 지정 + 출력 경로 지정
  python sar_worst.py --from 2024-02-27 --to 2024-05-01   # 검토 기간 지정 (두 날짜 포함)

기준 엑셀 (System Check Date 표, 예: input/systemcheckdate.xlsx)
  입력에 섞여 있으면 자동 인식 → 날짜·대역이 같은 측정 건만 정리.
  측정 건에 'SAR 4' 처럼 Lab 이 적혀 있으면 Lab 번호도 같아야 함.
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
SHEET_WARN = []  # 한 파일을 읽는 동안의 시트 경고 (파일 전체에서 데이터를 못 찾았을 때만 WARN 으로 옮김)


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


def iso_date(s):
    """명령줄 날짜 인자 'YYYY-MM-DD' → date"""
    try:
        return dt.datetime.strptime(s.strip(), "%Y-%m-%d").date()
    except ValueError:
        raise argparse.ArgumentTypeError(f"날짜는 YYYY-MM-DD 형식이어야 해요: {s}")


def serial_key(s):
    """Serial 비교용: 대소문자·공백 무시 (4d174 = 4D174), 숫자로 저장된 1205.0 → 1205"""
    return re.sub(r"\s", "", re.sub(r"\.0+$", "", text(s))).upper()


# ---- 기준 엑셀(System Check Date)과 맞추기: 날짜 · 대역 · SAR Lab ----
SYSTEM_BANDS = [13.5, 750, 835, 1750, 1900, 2300, 2450, 2600, 3500, 5250, 5600, 5750, 5800, 6000, 8000]


def nearest_bands(f):
    """가장 가까운 시스템 대역 (정확히 중간이면 둘 다)"""
    d = min(abs(b - f) for b in SYSTEM_BANDS)
    return [b for b in SYSTEM_BANDS if abs(abs(b - f) - d) < 1e-9]


def bands_of(freq, typ):
    """측정 건의 대역: Frequency 열 값 → 없으면 Type (D2450V2 → 2450, D5GHzV2 → 5 GHz 대역 전부)"""
    if freq is not None and freq > 0:
        return nearest_bands(freq)
    m = re.match(r"D\s*(\d+(?:\.\d+)?)\s*GHz", text(typ), re.I)
    if m:
        g = float(m[1])
        lo, hi = g * 1000, (int(g) + 1) * 1000
        in_ghz = [b for b in SYSTEM_BANDS if lo <= b < hi]
        return in_ghz or nearest_bands(lo)
    m = re.match(r"D\s*(\d+(?:\.\d+)?)", text(typ), re.I)
    return nearest_bands(float(m[1])) if m else []


def lab_no(v, allow_bare=False):
    """'SAR 4' / 'SAR4' / 'SAR-4' → 4 (allow_bare 이면 '4' 도 4)"""
    t = text(v)
    m = re.search(r"SAR\s*[-#.]?\s*(\d+)", t, re.I) or (re.fullmatch(r"(\d+)", t) if allow_bare else None)
    return int(m[1]) if m else None


def band_text(bs):
    return "/".join(text(b) for b in bs)


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
            "freq": re.search(r"freq|mhz|ghz", h),
            "lab": re.search(r"\blab\b|sar\s*system", h),
        }
        for k, ok in checks.items():
            if k not in col and ok:
                col[k] = c
                if k == "freq":
                    col["freq_ghz"] = "ghz" in h and "mhz" not in h
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
    if missing:  # 시스템 체크 표가 아닌 시트(Dipole 등): 파일 전체에서 데이터를 못 찾았을 때만 경고 (main 에서 판단)
        SHEET_WARN.append(f"{source}: {', '.join(missing)} 열을 찾지 못해 건너뜀")
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
        # 대역(기준 엑셀 비교용): Frequency 열 → 없으면 Type
        fv = num(get(s[0], col.get("freq")))
        b["bands"] = bands_of(None if fv is None else (fv * 1000 if col.get("freq_ghz") else fv), b["type"])
        # SAR Lab: Lab/SAR System 열 → 없으면 측정 건 안에서 'SAR 4' 같은 글자 (원본 Plot 칸 등)
        b["lab"] = lab_no(get(s[0], col["lab"]), True) if col.get("lab") else None
        for x in range(s[0], s[1] + 1):
            for c in range(1, max_c + 1):
                if b["lab"] is None and c not in (label_col, col["type"]) \
                        and re.fullmatch(r"SAR\s*[-#.]?\s*\d+", text(get(x, c)), re.I):
                    b["lab"] = lab_no(get(x, c))
        if not b["type"]:
            WARN.append(f"{source} {s[0]}행: Type 이 비어 있어 건너뜀")
            continue
        if not col.get("serial"):
            if not out:
                WARN.append(f"{source}: Serial 열을 찾지 못해 worst 판정에서 제외 (표에는 표시)")
        elif not b["serial"]:
            WARN.append(f"{source} {s[0]}행: Serial 이 비어 있어 worst 판정에서 제외 (표에는 표시)")
        if b["g1"]["delta"] is None or b["g10"]["delta"] is None:
            continue  # 값이 비어 있는 행(빈 양식 줄)은 조용히 건너뜀
        b["d1"], b["d10"] = b["g1"]["delta"], b["g10"]["delta"]
        b["sig"] = json.dumps([date_key(b), band_key(b["type"]), b["serial"], b["rows"]])
        out.append(b)
    return out


def read_ref(wb, file):
    """기준 엑셀: 'System Check Date' 헤더가 있는 표 (test_date 도구 결과 systemcheckdate.xlsx 등).
    헤더 행에서 날짜·대역(MHz/대역/Band/Freq)·SAR Lab 열을 찾음. 표가 없으면 None"""
    found, rows = False, []
    for ws in wb.worksheets:
        master = {}
        for m in ws.merged_cells.ranges:
            for r in range(m.min_row, m.max_row + 1):
                for c in range(m.min_col, m.max_col + 1):
                    master[(r, c)] = (m.min_row, m.min_col)

        def get(r, c):
            rr, cc = master.get((r, c), (r, c))
            return ws.cell(rr, cc).value

        max_r, max_c = ws.max_row, ws.max_column
        for hr in range(1, min(max_r, 30) + 1):
            dc = bc = lc = None
            for c in range(1, max_c + 1):
                h = low(get(hr, c))
                if not dc and "system check date" in h:
                    dc = c
                elif not bc and re.search(r"mhz|대역|band|freq", h):
                    bc = c
                elif not lc and re.search(r"\blab\b", h):
                    lc = c
            if not dc or not bc:
                continue
            found = True
            for r in range(hr + 1, max_r + 1):
                d = parse_date(get(r, dc))
                if not d:
                    continue
                bands = []
                for x in re.findall(r"\d+(?:\.\d+)?", text(get(r, bc))):
                    bands += [b for b in nearest_bands(float(x)) if b not in bands]
                if bands:
                    rows.append({"file": file, "date": d, "bands": bands,
                                 "lab": lab_no(get(r, lc), True) if lc else None})
            break
    return rows if found else None


def ref_hit(x, b):
    """같은 날짜 + 같은 대역 + (측정 건에 'SAR 4' 가 있으면) 같은 Lab 번호"""
    return (x["date"] == b["date"]
            and (not b["bands"] or any(v in b["bands"] for v in x["bands"]))
            and (b["lab"] is None or x["lab"] is None or x["lab"] == b["lab"]))


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

def sheet_base(n):
    """시트 이름의 기본: 파일 이름에 'SAR숫자'가 있으면 그 부분만
    (예: 'Liquid and System Check Table SAR11.xlsx' → 'SAR11'), 없으면 파일 이름(확장자 제외) 그대로"""
    s = re.sub(r"\.(xlsx|xlsm)$", "", n, flags=re.I)
    m = re.search(r"SAR\s*-?\s*\d+", s, flags=re.I)
    return m.group(0) if m else s


def sheet_names(names):
    """엑셀 시트 이름 규칙: 31자 이하, [ ] : * ? / 역슬래시 사용 불가, 중복 불가"""
    used, out = {"worst summary"}, {}
    for n in names:
        base = re.sub(r"[\[\]:*?/\\]", "_", sheet_base(n)).strip("'").strip() or "Sheet"
        base = base[:31]
        name, k = base, 2
        while name.lower() in used:
            suf = f" ({k})"
            name, k = base[:31 - len(suf)] + suf, k + 1
        used.add(name.lower())
        out[n] = name
    return out


def file_numbers(name):
    return [int(x) for x in re.findall(r"\d+", re.sub(r"\.(xlsx|xlsm)$", "", name, flags=re.I))]


def sort_files(names):
    """파일 순서: 파일 이름에 번호가 있으면 그 번호 순서대로.
    - 이름에서 처음 나오는 숫자로 비교 (3_A, 1_B, 2_C → 1_B, 2_C, 3_A / A_3, B_1, C_2 → B_1, C_2, A_3)
    - 처음 숫자가 같으면 이어지는 숫자까지 비교 (SC_2024-02-26 < SC_2024-03-05)
    - 번호가 없는 파일은 뒤에, 불러온 순서대로"""
    numbered = [(file_numbers(n), n, i) for i, n in enumerate(names) if file_numbers(n)]
    plain = [n for n in names if not file_numbers(n)]
    width = max((len(x[0]) for x in numbered), default=0)
    numbered.sort(key=lambda x: (x[0] + [-1] * (width - len(x[0])), x[1], x[2]))
    return [n for _, n, _ in numbered] + plain


def analyze(all_records, file_names):
    """all_records: 모든 파일의 측정 건 (파일 간 중복 포함).
    반환: (파일별 표, 비교에 쓴 측정 건, worst 목록, sig→worst)"""
    file_names = sort_files(file_names)  # 파일 이름의 번호 순서 → 시트 순서·넘버링 순서
    names = sheet_names(file_names)
    idx = {id(b): i for i, b in enumerate(all_records)}
    # 1) 파일별 표 (엑셀은 합치지 않음). 파일 안에서는 날짜순
    by_file = [{"name": n, "sheet": names[n],
                # 시트 이름 규칙(31자·금지 문자·중복) 때문에 기본 이름(sheet_base)과 달라졌는지
                "renamed": names[n] != sheet_base(n),
                "records": sorted((b for b in all_records if b["file"] == n), key=lambda b: (date_key(b), idx[id(b)]))}
               for n in file_names]
    # 표 순서: 파일 순서 → 파일 안 위→아래
    pos = {}
    for f in by_file:
        for b in f["records"]:
            pos.setdefault(b["sig"], len(pos))
    # 2) worst 비교용: 모든 파일의 측정 건, 같은 측정 건은 한 번만
    seen, records = set(), []
    for f in by_file:
        for b in f["records"]:
            if b["sig"] not in seen:
                seen.add(b["sig"])
                records.append(b)
    records.sort(key=lambda b: (date_key(b), pos[b["sig"]]))
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
        n_files = len({b["file"] for b in all_records if serial_key(b["serial"]) == serial})
        worst.append({"serial": serial, "rec": best, "tied": tied, "count": len(recs),
                      "n_files": n_files, "why": why})
    # 정리 후 번호가 하나도 없는 파일(시트) → 그 파일 전체(Serial 무관)에서 worst 하나를 골라 번호 추가
    marked = {w["rec"]["sig"] for w in worst}
    for f in by_file:
        recs = f["records"]
        if any(b["sig"] in marked for b in recs):
            continue
        best = recs[0]  # 파일 안은 날짜순 → 완전 동률이면 먼저 측정된 것
        for x in recs:
            if (r2(x["d1"]), r2(x["d10"])) > (r2(best["d1"]), r2(best["d10"])):
                best = x
        key = (r2(best["d1"]), r2(best["d10"]))
        tied = [x for x in recs if x is not best and x["sig"] != best["sig"] and (r2(x["d1"]), r2(x["d10"])) == key]
        same1g = [x for x in recs if x is not best and x["sig"] != best["sig"] and r2(x["d1"]) == key[0]]
        why = "번호 없는 파일 → 파일 전체 worst" + (
            " (1g·10g 완전 동률 → 먼저 측정된 데이터)" if tied else " (1g 동률 → 10g 로 결정)" if same1g else "")
        worst.append({"serial": serial_key(best["serial"]), "rec": best, "tied": tied, "count": len(recs),
                      "n_files": 1, "why": why, "file_worst": True})
        marked.add(best["sig"])
    # 넘버링 순서 고정: 파일 번호 순서 → 각 시트 안 위→아래
    worst.sort(key=lambda w: pos[w["rec"]["sig"]])
    # 번호는 측정 건(sig) 기준 → 같은 측정 건이 여러 파일 시트에 있으면 모두 같은 번호
    no = {}
    for k, w in enumerate(worst, 1):
        w["no"] = k
        no[w["rec"]["sig"]] = w
    return by_file, records, worst, no


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


def write_main(ws, records, no, source_name=None):
    """source_name 이 있으면(시트 이름이 파일 이름과 달라진 경우) 표 위에 원본 파일 이름 한 줄"""
    o = 1 if source_name else 0
    if o:
        ws.merge_cells("A1:J1")
        ws["A1"].value = f"원본 파일: {source_name}"
        ws["A1"].font = BOLD
        ws["A1"].alignment = Alignment(horizontal="left", vertical="center")
        ws.row_dimensions[1].height = 20
    h1, h2 = 1 + o, 2 + o
    for rng, v in [(f"A{h1}:A{h2}", "Date Tested"), (f"B{h1}:C{h1}", "System Dipole"), (f"B{h2}", "Type"),
                   (f"C{h2}", "Serial #"), (f"D{h1}:E{h2}", "T.S.\nLiquid"), (f"F{h1}:G{h1}", "Measured Results"),
                   (f"F{h2}", "Zoom Scan to\n100 mW"), (f"G{h2}", "Normalize\nto 1 W"),
                   (f"H{h1}:H{h2}", "Target\n(Ref. Value)"), (f"I{h1}:I{h2}", "Delta\n±10 %"), (f"J{h1}:J{h2}", "Plot No.")]:
        box(ws, rng, v, fill=HEAD_FILL)
    ws.row_dimensions[h1].height = 20
    ws.row_dimensions[h2].height = 30
    for col, w in zip("ABCDEFGHIJ", (16, 13, 11, 9, 7, 12, 12, 12, 10, 12)):
        ws.column_dimensions[col].width = w

    r = h2 + 1
    for rec in records:
        w = no.get(rec["sig"])
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
            box(ws, f"G{rr}", g["norm"], fmt="0.00")
            box(ws, f"H{rr}", g["target"], fmt="0.00")
            box(ws, f"I{rr}", None if g["delta"] is None else round(g["delta"], 2), fmt="0.00")
        box(ws, span("J"), w["no"] if w else None, font=BOLD if w else FONT)
        if w and w["tied"]:
            ws[f"J{r}"].comment = Comment("1g·10g Delta 완전 동률: " + ", ".join(fmt_date(x) for x in w["tied"])
                                          + " 데이터와 같음 (먼저 측정된 데이터 선택)", "sar_worst")
        r += n
    ws.freeze_panes = f"A{h2 + 1}"


def write_summary(ws, worst, by_file):
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
    # 시트 이름 ↔ 원본 파일 이름 대응표: Worst 표 아래 한 줄 띄우고, 넓은 I~L 열에
    t0, c0 = len(worst) + 3, 9
    for i, h in enumerate(["시트 이름", "원본 파일 이름", "측정 건수", "비고"]):
        box(ws, ws.cell(t0, c0 + i).coordinate, h, font=BOLD, fill=HEAD_FILL)
    for k, f in enumerate(by_file, 1):
        vals = [f["sheet"], f["name"], len(f["records"]),
                "시트 이름 규칙(31자·금지 문자·중복) 때문에 바뀜" if f["renamed"] else None]
        for i, v in enumerate(vals):
            box(ws, ws.cell(t0 + k, c0 + i).coordinate, v)
    ws.column_dimensions["K"].width = 10
    ws.column_dimensions["L"].width = 40
    ws.freeze_panes = "A2"


# ================================================================ main

def main():
    ap = argparse.ArgumentParser(description="SAR System Check Serial별 worst 넘버링")
    ap.add_argument("inputs", nargs="*", default=[str(HERE / "input")], help="엑셀 파일 또는 폴더")
    ap.add_argument("-o", "--output", default=None,
                    help="출력 경로 (기본: output/SystemCheck_Worst[_시작일~종료일].xlsx)")
    ap.add_argument("--from", dest="date_from", type=iso_date, help="검토 시작일 YYYY-MM-DD (포함)")
    ap.add_argument("--to", dest="date_to", type=iso_date, help="검토 종료일 YYYY-MM-DD (포함)")
    args = ap.parse_args()
    if args.date_from and args.date_to and args.date_from > args.date_to:
        sys.exit("시작일이 종료일보다 늦어요. --from / --to 를 확인하세요.")
    period_on = bool(args.date_from or args.date_to)

    files = collect_files(args.inputs)
    if not files:
        sys.exit("읽을 엑셀 파일이 없습니다. input 폴더에 넣거나 파일 경로를 지정하세요.")

    records, seen, file_names, refs, ref_files = [], set(), [], [], []
    for f in files:
        if f.name in file_names or f.name in ref_files:
            WARN.append(f"{f.name}: 이미 불러온 파일이라 건너뛰었어요")
            continue
        print(f"읽는 중: {f.name}")
        try:
            wb = load_workbook(f, data_only=True)
        except Exception as e:  # noqa: BLE001
            WARN.append(f"{f.name}: 파일을 열 수 없어요 ({e})")
            continue
        SHEET_WARN.clear()
        recs = [b for ws in wb.worksheets for b in read_sheet(ws, f.name)]
        if not recs:
            WARN.extend(SHEET_WARN)
        # 시스템 체크 데이터가 없고 'System Check Date' 표가 있으면 → 기준 엑셀
        ref = None if recs else read_ref(wb, f.name)
        if ref is not None:
            refs += ref
            ref_files.append(f.name)
            print(f"  · 기준 엑셀: {f.name} ({len(ref)}건)")
            if not ref:
                WARN.append(f"{f.name}: 기준 엑셀(System Check Date)에 날짜·대역이 있는 행이 없어요")
            continue
        if not recs:
            WARN.append(f"{f.name}: 시스템 체크 데이터를 찾지 못했어요 ('1g' 행과 날짜·Type·Delta 열 확인)")
            continue
        # 파일은 합치지 않음: 파일 간 중복 측정 건도 그 파일 시트에 그대로 표시하고, worst 비교 때만 한 번으로 셈
        dup = 0
        for b in recs:
            if b["sig"] in seen:
                dup += 1
            seen.add(b["sig"])
            records.append(b)
        if dup:
            WARN.append(f"{f.name}: 다른 파일과 같은 측정 건 {dup}건은 이 파일 시트에는 표시하고, worst 비교에서는 한 번만 셌어요")
        file_names.append(f.name)

    for w in WARN:
        print("  ! " + w)
    if not records:
        sys.exit("시스템 체크 데이터를 찾지 못했습니다.")
    all_records = records

    # 검토 기간: 기간 밖(과 날짜를 못 읽은) 측정 건은 비교·시트에서 모두 뺌
    if period_on:
        total = len(records)
        records = [b for b in records if b["date"]
                   and (not args.date_from or b["date"] >= args.date_from)
                   and (not args.date_to or b["date"] <= args.date_to)]
        print(f"  · 검토 기간 {args.date_from or '처음'} ~ {args.date_to or '끝'} → {total}건 중 {len(records)}건 사용")
        if not records:
            sys.exit("선택한 검토 기간에 해당하는 측정 건이 없어요.")
        skipped = [n for n in file_names if not any(b["file"] == n for b in records)]
        if skipped:
            print(f"  · 기간 안에 측정 건이 없어 시트를 만들지 않는 파일: {', '.join(skipped)}")
        file_names = [n for n in file_names if n not in skipped]

    # 기준 엑셀: 같은 날짜·대역(·SAR Lab) 행이 있는 측정 건만 비교·시트에 넣음
    if refs:
        in_p = len(records)
        records = [b for b in records if b["date"] and any(ref_hit(x, b) for x in refs)]
        print(f"  · 기준 엑셀({', '.join(ref_files)}) 비교 → {in_p}건 중 {len(records)}건 사용, "
              f"{in_p - len(records)}건은 날짜·대역·SAR Lab 이 안 맞아 뺌")
        missing = [x for x in refs
                   if (not args.date_from or x["date"] >= args.date_from)
                   and (not args.date_to or x["date"] <= args.date_to)
                   and not any(b["date"] and ref_hit(x, b) for b in all_records)]
        if missing:  # 오류(빨간색): 기준 엑셀에 날짜가 있는데 그 날짜·대역·Lab 의 시스템 체크가 없음
            red, end = ("\033[91m", "\033[0m") if sys.stdout.isatty() else ("", "")
            print(f"{red}  [오류] 기준 엑셀에는 있지만 시스템 체크 데이터가 없어요 ({len(missing)}건){end}")
            for x in missing:
                print(f"{red}    - {x['date']} · {band_text(x['bands'])} MHz"
                      + (f" · SAR {x['lab']}" if x["lab"] is not None else "") + f" — 시스템 체크 없음{end}")
        if not records:
            sys.exit("기준 엑셀과 날짜·대역·SAR Lab 이 같은 측정 건이 없어요.")
        skipped = [n for n in file_names if not any(b["file"] == n for b in records)]
        if skipped:
            print(f"  · 기준 엑셀에 맞는 측정 건이 없어 시트를 만들지 않는 파일: {', '.join(skipped)}")
        file_names = [n for n in file_names if n not in skipped]

    by_file, uniq, worst, no = analyze(records, file_names)
    wb = Workbook()
    wb.remove(wb.active)
    for f in by_file:  # 파일 하나 = 시트 하나
        write_main(wb.create_sheet(f["sheet"]), f["records"], no, f["name"] if f["renamed"] else None)
        print(f"  · 시트 '{f['sheet']}': {f['name']} 측정 {len(f['records'])}건, "
              f"worst 번호 {sum(1 for b in f['records'] if b['sig'] in no)}개")
    write_summary(wb.create_sheet("Worst Summary"), worst, by_file)

    if args.output:
        out = Path(args.output)
    else:
        label = f"_{args.date_from or '처음'}~{args.date_to or '끝'}" if period_on else ""
        out = HERE / "output" / f"SystemCheck_Worst{label}.xlsx"
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out)

    n_fw = sum(1 for w in worst if w.get("file_worst"))
    print(f"\n파일 {len(by_file)}개(파일별 시트)의 측정 {len(uniq)}건(중복 제외)을 비교 → Serial {len(worst) - n_fw}개"
          + (f" + 번호 없던 파일에 파일 전체 worst {n_fw}개" if n_fw else ""))
    for w in worst:
        rec = w["rec"]
        print(f"  Plot {w['no']:>2}: Serial {rec['serial']:<8} {rec['type']:<11} {fmt_date(rec):<10} "
              f"1g {rec['d1']:+6.2f}%  10g {rec['d10']:+6.2f}%  "
              f"({w['n_files']}개 파일 {w['count']}건 중, 원본 {rec['source']})"
              + ("  ※완전 동률" if w["tied"] else ""))
    print(f"\n저장: {out}")


if __name__ == "__main__":
    main()
