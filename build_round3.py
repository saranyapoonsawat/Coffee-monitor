#!/usr/bin/env python3
"""สร้างต้นแบบ ชบา Monitor (prototype/) จากไฟล์ใน data/ — รันซ้ำได้

ใช้เฉพาะไลบรารีมาตรฐานของ Python (อ่าน xlsx/docx ด้วย zipfile + XML)
    python3 build.py
"""
import csv
import html
import json
import re
import zipfile
import xml.etree.ElementTree as ET
from collections import Counter, defaultdict
from pathlib import Path

# WCAG สคริปต์นี้อยู่ใน prototype_round3/ · อ่าน data/ จากโฟลเดอร์แม่ และเขียนหน้าลงโฟลเดอร์นี้
ROOT = Path(__file__).resolve().parent.parent
DATA = ROOT / "data"
OUT = Path(__file__).resolve().parent
XLSX = DATA / "IS757_ข้อมูลจำลอง_ชบาคอฟฟี่_v2.xlsx"
CSV = DATA / "IS757_ตัวเลขรายชั่วโมงสมมติ.csv"

DATA_AS_OF = "จ. 26 ต.ค. 69 · 14:00 น."
CONF_REVIEW = 0.75
PH_INTERVIEW = "[ตัวเลขจากการสัมภาษณ์]"
PH_CONFIRM = "[รอยืนยัน]"
PH_EVIDENCE = "[รอใส่หลักฐานจากบันทึกสัมภาษณ์]"

# F04 ชื่อระดับความรุนแรงชุดเดียว ใช้ทุกหน้าเว็บและมือถือ (ทีมกำหนด รอบ 1)
#     "วิกฤต" เป็นระดับของเคสเท่านั้น · ระดับของโพสต์ใช้ "ความเร่งด่วน (ระบบ)" จาก xlsx = สูง/กลาง/ต่ำ
SEV_LEVELS = ["วิกฤต", "สูง", "กลาง", "ต่ำ"]
SEV_CRIT = SEV_LEVELS[0]
SEV_TAG = {"วิกฤต": "crit", "สูง": "warn", "กลาง": "pri", "ต่ำ": "plain"}
# F10 F11 ผู้จัดการเขต = ผู้รับสำรอง ส่งเมื่อผู้รับหลักไม่รับทราบใน 5 นาที (ค่าที่ทีมกำหนด ไม่ได้มาจาก data/)
ESCALATE_MIN = 5
BACKUP_RULE = f"ถ้าผู้รับหลักไม่รับทราบใน {ESCALATE_MIN} นาที ส่งต่อผู้รับสำรอง"
# F08 เวลาที่กด "รับเรื่องนี้" ในต้นแบบ (นาทีถัดจาก push 22:42)
ACCEPT_TIME = "22:43"
# F22 ชื่อเมนูชุดเดียวกันทั้งเว็บและแท็บมือถือ
# R3 รายการปรับปรุงคำ: ชื่อเมนูใหม่ ใช้ร่วมกันทั้งเว็บ มือถือ และ FigCraft
NAV_LABEL = {"overview": "Dashboard", "posts": "Social Feed", "cases": "Alerts & Cases"}

# ---------------------------------------------------------------- tokens
T = {
    "bg": "#F3F2F8", "card": "#FFFFFF",
    "ink": "#16151F", "ink2": "#4A4858", "ink3": "#5E5B72",
    "primary": "#4F3FD1", "lime": "#D9F26B",
    "crit": "#B42318", "warn": "#8A4B00", "ok": "#1F6B3A",
    "neg": "#D06A1E", "neu": "#C4C1D3", "mix": "#7E5BD0", "pos": "#1D3F9E",
    # ค่าเสริมที่ build.py ใช้ (ตรวจความต่างสีด้วย)
    "white": "#FFFFFF", "line": "#E2E0EC", "control": "#767289",
    "critTint": "#FDECEA", "warnTint": "#FFF3DF", "okTint": "#E7F3EA",
    "primaryTint": "#ECEAFB", "band": "#ECEAF4",
}
SENT_COLOR = {"ลบ": T["neg"], "เป็นกลาง": T["neu"], "ผสม": T["mix"], "บวก": T["pos"]}
SENT_ORDER = ["ลบ", "เป็นกลาง", "ผสม", "บวก"]
NEU_STROKE = T["ink3"]  # เส้นขอบช่วยให้สีเป็นกลางผ่าน 3:1


def _lum(hexcolor):
    h = hexcolor.lstrip("#")
    out = []
    for i in (0, 2, 4):
        c = int(h[i:i + 2], 16) / 255
        out.append(c / 12.92 if c <= 0.04045 else ((c + 0.055) / 1.055) ** 2.4)
    r, g, b = out
    return 0.2126 * r + 0.7152 * g + 0.0722 * b


def contrast(a, b):
    la, lb = sorted((_lum(a), _lum(b)), reverse=True)
    return (la + 0.05) / (lb + 0.05)


# (ชื่อ, หน้า, พื้น, ประเภท text/graphic, ใช้ที่ไหน, ทางช่วยถ้าไม่ผ่าน)
COLOR_PAIRS = [
    ("ตัวอักษรหลัก", "ink", "bg", "text", "ข้อความทั่วไปบนพื้นหลัง", None),
    ("ตัวอักษรหลัก", "ink", "card", "text", "ข้อความบนการ์ด", None),
    ("ตัวอักษรรอง", "ink2", "bg", "text", "คำอธิบาย", None),
    ("ตัวอักษรรอง", "ink2", "card", "text", "คำอธิบายบนการ์ด", None),
    ("ตัวอักษรเบา", "ink3", "bg", "text", "ป้ายกำกับ แกนกราฟ", None),
    ("ตัวอักษรเบา", "ink3", "card", "text", "ป้ายกำกับบนการ์ด", None),
    ("ลิงก์ primary", "primary", "card", "text", "ลิงก์ ปุ่มรอง", None),
    ("ลิงก์ primary", "primary", "bg", "text", "ลิงก์บนพื้นหลัง", None),
    ("ปุ่มหลัก", "white", "primary", "text", "ตัวอักษรบนปุ่มหลัก", None),
    ("primary บนพื้นอ่อน", "primary", "primaryTint", "text", "เมนูที่เลือก ชิป", None),
    ("ไฮไลต์ lime", "ink", "lime", "text", "คำหลักฐาน ป้ายเด่น", None),
    ("วิกฤต", "crit", "card", "text", "ข้อความวิกฤต", None),
    ("วิกฤต บนพื้นอ่อน", "crit", "critTint", "text", "แถบแจ้งเตือน", None),
    ("ป้ายวิกฤต", "white", "crit", "text", "ป้าย วิกฤต", None),
    ("เตือน", "warn", "card", "text", "ป้าย รอตรวจสอบ", None),
    ("เตือน บนพื้นอ่อน", "warn", "warnTint", "text", "กล่องเตือน", None),
    ("สำเร็จ", "ok", "card", "text", "สถานะสำเร็จ", None),
    ("สำเร็จ บนพื้นอ่อน", "ok", "okTint", "text", "ป้าย ยืนยันแล้ว · ตรวจสอบแล้ว (F35)", None),
    ("ขอบการ์ดชั่วโมงพุ่ง (F26)", "primary", "card", "graphic", "เส้นขอบการ์ด KPI ชั่วโมงพุ่งสูงสุด", None),
    ("ขอบแถบเคส (F26)", "crit", "critTint", "graphic", "เส้นขอบซ้ายแถบเคสที่ยังไม่ปิด", None),
    ("ป้ายระดับสูง (F04)", "warn", "warnTint", "text", "ป้าย ความเร่งด่วน สูง", None),
    ("ป้ายระดับกลาง (F04)", "primary", "primaryTint", "text", "ป้าย ความเร่งด่วน กลาง", None),
    ("ป้ายระดับต่ำ (F04)", "ink2", "bg", "text", "ป้าย ความเร่งด่วน ต่ำ", None),
    ("ตัวอักษรบนแผงสาธิต (F16)", "ink", "warnTint", "text", "แผงโหมดทดลองใช้งานนอกกรอบโทรศัพท์", None),
    ("พื้นเข้ม (จอล็อก แถบเลิกทำ)", "white", "ink", "text", "ข้อความบนพื้นเข้ม", None),
    ("lime บนพื้นเข้ม", "lime", "ink", "graphic", "ปุ่มเลิกทำบนแถบเข้ม", None),
    ("กราฟ ลบ", "neg", "card", "graphic", "แท่ง/เส้นเชิงลบ", None),
    ("กราฟ ลบ บนแถบนอกเวลา", "neg", "band", "graphic", "เส้นเชิงลบในช่วงนอกเวลา", None),
    ("กราฟ ผสม", "mix", "card", "graphic", "แท่งผสม", None),
    ("กราฟ บวก", "pos", "card", "graphic", "แท่งบวก", None),
    ("กราฟ เป็นกลาง", "neu", "card", "graphic", "แท่งเป็นกลาง",
     "ไม่ผ่านเมื่อใช้ลำพัง: วาดเส้นขอบ #5E5B72 รอบแท่ง และมีตัวเลขกำกับ"),
    ("เส้นขอบแท่งเป็นกลาง", "ink3", "card", "graphic", "เส้นขอบรอบแท่งเป็นกลาง", None),
    ("เส้นทั้งหมด (primary)", "primary", "card", "graphic", "เส้นทั้งหมดในกราฟ", None),
    ("ขอบช่องกรอก", "control", "card", "graphic", "ขอบ select/input/checkbox", None),
    ("วงโฟกัส", "primary", "card", "graphic", "เส้นโฟกัสคีย์บอร์ด", None),
]


def contrast_rows():
    rows = []
    for name, fg, bg, kind, where, mitig in COLOR_PAIRS:
        r = contrast(T[fg], T[bg])
        need = 4.5 if kind == "text" else 3.0
        rows.append(dict(name=name, fg=fg, bg=bg, fgv=T[fg], bgv=T[bg], kind=kind,
                         where=where, ratio=round(r, 2), need=need, ok=r >= need, mitig=mitig))
    return rows


def label_color(fill):
    """เลือกสีตัวอักษรบนแท่งกราฟที่ผ่าน 4.5:1 (หรือ None = วางนอกแท่ง)"""
    best = max((T["ink"], T["white"]), key=lambda c: contrast(c, fill))
    return best if contrast(best, fill) >= 4.5 else None


# ---------------------------------------------------------------- read data
NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}


def read_xlsx(path):
    z = zipfile.ZipFile(path)
    ss = [''.join(t.text or '' for t in si.iter('{%s}t' % NS['m']))
          for si in ET.fromstring(z.read('xl/sharedStrings.xml')).findall('m:si', NS)]
    wb = ET.fromstring(z.read('xl/workbook.xml'))
    sheets = {}
    for i, s in enumerate(wb.find('m:sheets', NS), 1):
        rows = []
        for row in ET.fromstring(z.read(f'xl/worksheets/sheet{i}.xml')).iter('{%s}row' % NS['m']):
            d = {}
            for c in row.findall('m:c', NS):
                col = re.match(r'[A-Z]+', c.get('r')).group()
                v = c.find('m:v', NS)
                t = c.get('t')
                if t == 's':
                    d[col] = ss[int(v.text)]
                elif t == 'inlineStr':
                    d[col] = ''.join(x.text or '' for x in c.iter('{%s}t' % NS['m']))
                else:
                    d[col] = v.text if v is not None else ''
            rows.append(d)
        sheets[s.get('name')] = rows
    return sheets


def load_posts():
    sh = read_xlsx(XLSX)
    rows = sh["ข้อความ"]
    head = rows[0]
    posts = {}
    for r in rows[1:]:
        p = {head[k]: (v or '').strip() for k, v in r.items() if k in head}
        posts[p["ID"]] = p
    notes = {r.get("A", ""): r.get("B", "") for r in sh["คำอธิบายและการแก้ไข"][1:]}
    return posts, notes


def load_csv():
    with open(CSV, encoding="utf-8-sig") as f:
        rows = list(csv.DictReader(f))
    for r in rows:
        r["n"] = int(r["จำนวนการกล่าวถึง"])
    return rows


POSTS, NOTES = load_posts()
CSVROWS = load_csv()

# ---------------------------------------------------------------- stats
DAYS = ["2026-10-19", "2026-10-20", "2026-10-21", "2026-10-22",
        "2026-10-23", "2026-10-24", "2026-10-25"]
DAY_TH = {}
for r in CSVROWS:
    DAY_TH.setdefault(r["ชั่วโมงเริ่ม"][:10], r["วันที่ (ไทย)"].split(" · ")[0])
HOUR_TH = {r["ชั่วโมงเริ่ม"]: r["วันที่ (ไทย)"] for r in CSVROWS}
WEEK = [r for r in CSVROWS if r["ชั่วโมงเริ่ม"][:10] in DAYS]  # 19–25 ต.ค. เท่านั้น


def hkey(day, h):
    return f"{day} {h:02d}:00"


def is_off(h):
    return h >= 20 or h < 7


S = {}
S["total"] = sum(r["n"] for r in WEEK)
S["neg"] = sum(r["n"] for r in WEEK if r["ความรู้สึก"] == "ลบ")
S["neg_pct"] = S["neg"] / S["total"] * 100
S["by_sent"] = Counter()
for r in WEEK:
    S["by_sent"][r["ความรู้สึก"]] += r["n"]
S["platforms"] = sorted({r["แพลตฟอร์ม"] for r in CSVROWS})
S["n_platforms"] = len(S["platforms"])
S["off_total"] = sum(r["n"] for r in WEEK if r["นอกเวลาทำการ"] == "ใช่")

hour_tot, hour_neg = Counter(), Counter()
for r in WEEK:
    hour_tot[r["ชั่วโมงเริ่ม"]] += r["n"]
    if r["ความรู้สึก"] == "ลบ":
        hour_neg[r["ชั่วโมงเริ่ม"]] += r["n"]
HOURS = [hkey(d, h) for d in DAYS for h in range(24)]
peak_key = max(HOURS, key=lambda k: hour_tot[k])
S["peak_key"] = peak_key
S["peak"] = hour_tot[peak_key]
S["peak_neg"] = hour_neg[peak_key]
S["peak_th"] = HOUR_TH[peak_key]
second = sorted(HOURS, key=lambda k: -hour_tot[k])[1]
S["second_key"], S["second"], S["second_th"] = second, hour_tot[second], HOUR_TH[second]


def baseline(hour):
    vals = [hour_tot[hkey(d, hour)] for d in DAYS[:4]]  # จ.–พฤ.
    return sum(vals) / len(vals)


S["base_peak_hour"] = baseline(int(peak_key[11:13]))
S["base_22"] = baseline(22)

S["daily"] = {d: sum(hour_tot[hkey(d, h)] for h in range(24)) for d in DAYS}
S["daily_neg"] = {d: sum(hour_neg[hkey(d, h)] for h in range(24)) for d in DAYS}
S["all_total_incl26"] = sum(r["n"] for r in CSVROWS)
S["total_26"] = sum(r["n"] for r in CSVROWS if r["ชั่วโมงเริ่ม"].startswith("2026-10-26"))
S["last_hour_26"] = max(r["ชั่วโมงเริ่ม"] for r in CSVROWS)

plat = defaultdict(Counter)
for r in WEEK:
    plat[r["แพลตฟอร์ม"]][r["ความรู้สึก"]] += r["n"]
S["plat"] = sorted(((p, c, sum(c.values())) for p, c in plat.items()), key=lambda x: -x[2])

br_tot, br_neg = Counter(), Counter()
for r in WEEK:
    br_tot[r["สาขา"]] += r["n"]
    if r["ความรู้สึก"] == "ลบ":
        br_neg[r["สาขา"]] += r["n"]
S["branches"] = sorted(((b, br_neg[b], br_tot[b]) for b in br_tot), key=lambda x: -x[1])
BRANCHES5 = ["ลาดพร้าว", "รังสิต", "อารีย์", "บางนา", "พระราม 2"]

CAMPAIGN = "โปร 1 แถม 1 สมาชิก"
camp_daily = Counter()
camp_sent = Counter()
for r in WEEK:
    if r["แคมเปญ"] == CAMPAIGN:
        camp_daily[r["ชั่วโมงเริ่ม"][:10]] += r["n"]
        camp_sent[r["ความรู้สึก"]] += r["n"]
S["camp_days"] = [d for d in DAYS if camp_daily[d]]
S["camp_daily"] = camp_daily
S["camp_total"] = sum(camp_daily.values())
S["camp_neg"] = camp_sent["ลบ"]
S["camp_sent"] = camp_sent
S["camp_26"] = sum(r["n"] for r in CSVROWS if r["แคมเปญ"] == CAMPAIGN and r["ชั่วโมงเริ่ม"].startswith("2026-10-26"))

# คืนวันศุกร์ (สำหรับ m3)
FRI = "2026-10-23"
S["fri_hours"] = [(h, hour_tot[hkey(FRI, h)], hour_neg[hkey(FRI, h)]) for h in range(18, 23)]

# ชุดตัวกรองที่ไม่มีผลลัพธ์จริงใน CSV (ใช้ในสถานะ "ว่าง")
_combo = Counter()
for r in WEEK:
    if r["แคมเปญ"] == CAMPAIGN:
        _combo[(r["แพลตฟอร์ม"], r["สาขา"])] += r["n"]
EMPTY_COMBO = None
for p in S["platforms"]:
    for b in BRANCHES5:
        if _combo[(p, b)] == 0:
            EMPTY_COMBO = (p, b)
            break
    if EMPTY_COMBO:
        break

# F24 สถิติแยกสาขา: คำนวณจาก CSV ชุดเดียวกัน เพื่อให้ตัวกรองสาขาใน w1 เปลี่ยนตัวเลขและกราฟจริง
def branch_stats(branch=None):
    rows = [r for r in WEEK if branch is None or r["สาขา"] == branch]
    st = {"branch": branch, "hour_tot": Counter(), "hour_neg": Counter(), "camp_daily": Counter(), "camp_sent": Counter()}
    pl = defaultdict(Counter)
    for r in rows:
        st["hour_tot"][r["ชั่วโมงเริ่ม"]] += r["n"]
        if r["ความรู้สึก"] == "ลบ":
            st["hour_neg"][r["ชั่วโมงเริ่ม"]] += r["n"]
        pl[r["แพลตฟอร์ม"]][r["ความรู้สึก"]] += r["n"]
        if r["แคมเปญ"] == CAMPAIGN:
            st["camp_daily"][r["ชั่วโมงเริ่ม"][:10]] += r["n"]
            st["camp_sent"][r["ความรู้สึก"]] += r["n"]
    st["total"] = sum(r["n"] for r in rows)
    st["neg"] = sum(r["n"] for r in rows if r["ความรู้สึก"] == "ลบ")
    st["neg_pct"] = st["neg"] / st["total"] * 100 if st["total"] else 0
    st["n_platforms"] = len(pl)
    st["plat"] = sorted(((p, c, sum(c.values())) for p, c in pl.items()), key=lambda x: -x[2])
    pk = max(HOURS, key=lambda k: st["hour_tot"][k])
    st["peak_key"], st["peak"], st["peak_neg"], st["peak_th"] = pk, st["hour_tot"][pk], st["hour_neg"][pk], HOUR_TH[pk]
    vals = [st["hour_tot"][hkey(d, int(pk[11:13]))] for d in DAYS[:4]]
    st["base_peak_hour"] = sum(vals) / len(vals)
    st["camp_days"] = S["camp_days"]
    st["camp_total"] = sum(st["camp_daily"].values())
    st["camp_neg"] = st["camp_sent"]["ลบ"]
    return st


ST_ALL = branch_stats()
ST_BRANCH = {b: branch_stats(b) for b in BRANCHES5}

# ตรวจค่าที่โจทย์กำหนดก่อนวาด
assert ST_ALL["total"] == S["total"] and ST_ALL["peak"] == S["peak"]  # F24 ฟังก์ชันใหม่ต้องได้ค่าเดียวกับของเดิม
assert S["total"] == 2122, S["total"]
assert S["neg"] == 909, S["neg"]
assert S["peak"] == 174 and peak_key == "2026-10-23 23:00", (S["peak"], peak_key)


def fmt(n):
    return f"{n:,}"


# ---------------------------------------------------------------- post helpers
EMOJI_LABEL = {"😭": "อีโมจิร้องไห้", "🥲": "อีโมจิยิ้มทั้งน้ำตา"}
EMOJI_RE = re.compile("[\U0001F300-\U0001FAFF☀-➿]")


def esc(s):
    return html.escape(str(s), quote=True)


def thai_dt(s):
    """'ศ. 23 ต.ค. 69 21:40' -> 'ศ. 23 ต.ค. 69 · 21:40 น.'"""
    m = re.match(r"^(.*\d{2}) (\d{1,2}:\d{2})$", s)
    return f"{m.group(1)} · {m.group(2)} น." if m else s


def sort_key(p):
    m = re.search(r"(\d+) ต\.ค\. \d+ (\d+):(\d+)", p["วันเวลา"])
    return (int(m.group(1)), int(m.group(2)), int(m.group(3)))


def post_time(p):
    return re.search(r"(\d{1,2}:\d{2})$", p["วันเวลา"]).group(1) + " น."


def keywords(p):
    return [k.strip() for k in p["คำหลักฐานที่ระบบใช้"].split(",") if k.strip()]


# ตามที่ผู้ออกแบบแก้ใน FigCraft (24 พ.ย. 69): ข้อความโพสต์ที่ยกมาในมือถือ m2/m3 ใช้คำ "พลาสติกแข็ง"
# มีผลเฉพาะ m2 (หลักฐาน C18) และ m3 (C18, C19) · หน้าเว็บและไฟล์ data/ ยังเป็นข้อความเดิม
MOBILE_QUOTE_EDIT = {"C18": ("ชิ้นใสแข็ง", "พลาสติกแข็ง"), "C19": ("ชิ้นแข็ง", "พลาสติกแข็ง")}


def mobile_quote(p):
    old, new = MOBILE_QUOTE_EDIT.get(p["ID"], ("", ""))
    if not old:
        return render_text(p["ข้อความ"], keywords(p))
    return render_text(p["ข้อความ"].replace(old, new), [k.replace(old, new) for k in keywords(p)])


def render_text(text, kws=(), mark=True):
    ranges = []
    if mark:
        for k in kws:
            i = text.find(k)
            if i >= 0:
                ranges.append((i, i + len(k)))
    ranges.sort()
    merged = []
    for a, b in ranges:
        if merged and a <= merged[-1][1]:
            merged[-1] = (merged[-1][0], max(b, merged[-1][1]))
        else:
            merged.append((a, b))
    out, pos = [], 0
    for a, b in merged:
        out.append(esc(text[pos:a]))
        out.append(f"<mark>{esc(text[a:b])}</mark>")
        pos = b
    out.append(esc(text[pos:]))
    s = "".join(out)

    def emo(m):
        lab = EMOJI_LABEL.get(m.group(0), "อีโมจิ")
        return f'<span class="emo">[{lab}]</span>'
    return EMOJI_RE.sub(emo, s)


def unmatched_kws(p):
    return [k for k in keywords(p) if k not in p["ข้อความ"]]


def is_sensitive(p):
    return "ข้อมูลสุขภาพ" in p["สิ่งที่แก้ / ข้อเสนอให้ทบทวน"]


def engagement(p):
    v = p["มีส่วนร่วม"]
    return "ไม่มียอดมีส่วนร่วม" if v in ("—", "") else f"มีส่วนร่วม {v}"


def conf_pct(p):
    return round(float(p["ความมั่นใจระบบ"]) * 100)


def needs_review(p):
    return float(p["ความมั่นใจระบบ"]) < CONF_REVIEW


def branch_parts(p):
    b = p["สาขา"]
    m = re.match(r"^(.*?) \((.*)\)$", b)
    if m:
        return m.group(1), m.group(2)
    return b, ""


def excluded(p):
    return p["ความรู้สึก (ควร)"] == "ไม่เกี่ยวข้อง"


def owner_status(p):
    """ผู้รับผิดชอบ/สถานะในต้นแบบ — ชื่อบทบาทคัดจากภาพ wireframe ใน data/ (ค่าสมมติ รอยืนยัน)"""
    pid, case = p["ID"], p["เคส (เสนอ)"]
    if pid == "C09":
        return "แอดมินเพจ", "ใหม่"
    if pid == "C29":
        return "ผู้ดูแลข้อมูลส่วนบุคคล", "ใหม่"
    if case == "K10":
        return "CS ส่วนกลาง", "อยู่ระหว่างดำเนินการ"
    if case:
        return "CS ส่วนกลาง", "ใหม่"
    return "ยังไม่มอบหมาย", "ใหม่"


REVIEW_IDS = sorted([pid for pid, p in POSTS.items() if needs_review(p)],
                    key=lambda i: float(POSTS[i]["ความมั่นใจระบบ"]))
K10 = sorted([p for p in POSTS.values() if p["เคส (เสนอ)"] == "K10"], key=sort_key)
ISSUES = []
for p in POSTS.values():
    if p["ประเด็น"] not in ISSUES and p["ประเด็น"] != "ไม่เกี่ยวข้อง":
        ISSUES.append(p["ประเด็น"])

C09, C18, C19, C33 = POSTS["C09"], POSTS["C18"], POSTS["C19"], POSTS["C33"]


def minutes(t):
    h, m = t.replace(" น.", "").split(":")
    return int(h) * 60 + int(m)


PUSH_TIME = "22:42"
M2_ELAPSED = minutes(PUSH_TIME) - minutes(post_time(C18))
SPREAD_MIN = minutes(post_time(C33)) - minutes(post_time(C18))
SPREAD = [C18, C19, C33]

# ---------------------------------------------------------------- icons
ICONS = {
    "grid": '<rect x="3" y="3" width="7" height="7" rx="1"/><rect x="14" y="3" width="7" height="7" rx="1"/><rect x="3" y="14" width="7" height="7" rx="1"/><rect x="14" y="14" width="7" height="7" rx="1"/>',
    "list": '<path d="M8 6h13M8 12h13M8 18h13M3 6h.01M3 12h.01M3 18h.01"/>',
    "bell": '<path d="M6 8a6 6 0 0 1 12 0c0 7 3 9 3 9H3s3-2 3-9"/><path d="M10.3 21a1.94 1.94 0 0 0 3.4 0"/>',
    "sliders": '<path d="M4 21v-7M4 10V3M12 21v-9M12 8V3M20 21v-5M20 12V3M1 14h6M9 8h6M17 16h6"/>',
    "search": '<circle cx="11" cy="11" r="7"/><path d="m21 21-4.3-4.3"/>',
    "alert": '<path d="M10.3 3.9 1.8 18a2 2 0 0 0 1.7 3h17a2 2 0 0 0 1.7-3L13.7 3.9a2 2 0 0 0-3.4 0z"/><path d="M12 9v4M12 17h.01"/>',
    "check": '<path d="M20 6 9 17l-5-5"/>',
    "clock": '<circle cx="12" cy="12" r="9"/><path d="M12 7v5l3 2"/>',
    "left": '<path d="M19 12H5M12 19l-7-7 7-7"/>',
    "right": '<path d="M5 12h14M12 5l7 7-7 7"/>',
    "send": '<path d="M22 2 11 13M22 2l-7 20-4-9-9-4 20-7z"/>',
    "eyeoff": '<path d="M17.9 17.9A10.1 10.1 0 0 1 12 20c-7 0-10-8-10-8a18.5 18.5 0 0 1 5.1-5.9M9.9 4.2A9.1 9.1 0 0 1 12 4c7 0 10 8 10 8a18.5 18.5 0 0 1-2.2 3.2M14.1 14.1a3 3 0 1 1-4.2-4.2M1 1l22 22"/>',
    "external": '<path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6M15 3h6v6M10 14 21 3"/>',
    "undo": '<path d="M9 14 4 9l5-5"/><path d="M4 9h10.5a5.5 5.5 0 0 1 0 11H11"/>',
    "user": '<circle cx="12" cy="8" r="4"/><path d="M4 21a8 8 0 0 1 16 0"/>',
    "refresh": '<path d="M21 12a9 9 0 1 1-3-6.7L21 8"/><path d="M21 3v5h-5"/>',
    "x": '<path d="M18 6 6 18M6 6l12 12"/>',
    "star": '<path d="m12 2 3.1 6.3 6.9 1-5 4.9 1.2 6.8L12 17.8 5.8 21l1.2-6.8-5-4.9 6.9-1z"/>',
    "pulse": '<path d="M22 12h-4l-3 9L9 3l-3 9H2"/>',
    "phone": '<path d="M22 16.9v3a2 2 0 0 1-2.2 2 19.8 19.8 0 0 1-8.6-3.1 19.5 19.5 0 0 1-6-6A19.8 19.8 0 0 1 2.1 4.2 2 2 0 0 1 4.1 2h3a2 2 0 0 1 2 1.7c.1 1 .4 1.9.7 2.8a2 2 0 0 1-.5 2.1L8 9.9a16 16 0 0 0 6 6l1.3-1.3a2 2 0 0 1 2.1-.4c.9.3 1.8.6 2.8.7a2 2 0 0 1 1.7 2z"/>',
    "lock": '<rect x="4" y="11" width="16" height="10" rx="2"/><path d="M8 11V7a4 4 0 0 1 8 0v4"/>',
    "filter": '<path d="M22 3H2l8 9.5V19l4 2v-8.5z"/>',
    "wifioff": '<path d="M1 1l22 22M16.7 11.1A11 11 0 0 1 19 12.6M5 12.6a11 11 0 0 1 5.2-2.5M10.7 5.1A16 16 0 0 1 22.6 9M1.4 9a16 16 0 0 1 4.4-2.8M8.5 16.1a5 5 0 0 1 7 0M12 20h.01"/>',
}


def icon(name, label=None, size=20, cls="ic"):
    a11y = f'role="img" aria-label="{esc(label)}"' if label else 'aria-hidden="true" focusable="false"'
    return (f'<svg class="{cls}" {a11y} width="{size}" height="{size}" viewBox="0 0 24 24" fill="none" '
            f'stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">'
            f'{ICONS[name]}</svg>')


def sent_badge(s, extra=""):
    color = SENT_COLOR.get(s)
    dot = (f'<span class="dot" style="background:{color};{"box-shadow:inset 0 0 0 1px " + NEU_STROKE if s == "เป็นกลาง" else ""}"></span>'
           if color else "")
    return f'<span class="sent {extra}" data-sent="{esc(s)}">{dot}{esc(s)}</span>'


def conf_meter(p):
    pct = conf_pct(p)
    low = needs_review(p)
    return (f'<span class="conf{" low" if low else ""}"><span class="meter" aria-hidden="true">'
            f'<span style="width:{pct}%"></span></span><span class="num">{pct}%</span>'
            f'{"<span class=\"tag warn\">รอตรวจสอบ</span>" if low else ""}</span>')


def sev_tag(level, with_icon=False):
    """F04 ป้ายระดับความรุนแรงชุดเดียว (ข้อความ + สี ไม่พึ่งสีอย่างเดียว)"""
    ic = icon('alert', size=14) if with_icon else ""  # F01 ไอคอนเตือนในป้ายวิกฤต
    return f'<span class="tag {SEV_TAG.get(level, "plain")}" data-sev="{esc(level)}">{ic}{esc(level)}</span>'


STATUS_TAG = {"ใหม่": "plain", "อยู่ระหว่างดำเนินการ": "pri", "ตรวจสอบแล้ว": "ok"}  # F35


def status_tag(s):
    ic = icon('check', size=13) if s == "ตรวจสอบแล้ว" else ""  # F35 ป้ายตรวจสอบแล้ว มีไอคอน + ข้อความ
    return f'<span class="tag {STATUS_TAG.get(s, "plain")}">{ic}{esc(s)}</span>'


def case_posts(case):
    return sorted([p for p in POSTS.values() if p["เคส (เสนอ)"] == case], key=sort_key)


def case_level(case):
    """F04 ระดับเคส: K10 = วิกฤต ตามสถานการณ์กรณีที่ 1 · เคสอื่นใช้ความเร่งด่วนสูงสุดของโพสต์ในเคส (จาก xlsx)"""
    if case == "K10":
        return SEV_CRIT
    lv = [p["ความเร่งด่วน (ระบบ)"] for p in case_posts(case)]
    return min(lv, key=SEV_LEVELS.index) if lv else ""


# ---------------------------------------------------------------- SVG charts
def svg_hourly(st=None):
    st = st or ST_ALL  # F24 วาดได้ทั้งภาพรวมและรายสาขา
    ht, hn = st["hour_tot"], st["hour_neg"]
    W, H, pl, pr, pt, pb = 1000, 300, 44, 16, 56, 34
    n = len(HOURS)
    step = (W - pl - pr) / (n - 1)
    ymax = 180 if st["branch"] is None else max(40, -(-st["peak"] * 12 // 10 // 20) * 20)
    ticks = (0, 50, 100, 150) if st["branch"] is None else tuple(round(ymax * i / 4) for i in range(4))

    def x(i):
        return pl + i * step

    def y(v):
        return pt + (1 - v / ymax) * (H - pt - pb)
    parts = []
    # แถบนอกเวลาทำการ
    i = 0
    while i < n:
        if is_off(i % 24):
            j = i
            while j + 1 < n and is_off((j + 1) % 24):
                j += 1
            x0 = max(pl, x(i) - step / 2)
            x1 = min(W - pr, x(j) + step / 2)
            parts.append(f'<rect x="{x0:.1f}" y="{pt}" width="{x1 - x0:.1f}" height="{H - pt - pb}" fill="{T["band"]}"/>')
            i = j + 1
        else:
            i += 1
    for v in ticks:
        parts.append(f'<line x1="{pl}" x2="{W - pr}" y1="{y(v):.1f}" y2="{y(v):.1f}" stroke="{T["line"]}"/>')
        parts.append(f'<text x="{pl - 8}" y="{y(v) + 4:.1f}" text-anchor="end" class="ax">{v}</text>')
    for di, d in enumerate(DAYS):
        parts.append(f'<text x="{x(di * 24 + 12):.1f}" y="{H - 10}" text-anchor="middle" class="ax">{esc(DAY_TH[d].rsplit(" ", 1)[0])}</text>')
    tot = " ".join(f"{x(i):.1f},{y(ht[k]):.1f}" for i, k in enumerate(HOURS))
    neg = " ".join(f"{x(i):.1f},{y(hn[k]):.1f}" for i, k in enumerate(HOURS))
    parts.append(f'<polyline points="{tot}" fill="none" stroke="{T["primary"]}" stroke-width="2" stroke-linejoin="round"/>')
    parts.append(f'<polyline points="{neg}" fill="none" stroke="{T["neg"]}" stroke-width="2" stroke-linejoin="round"/>')
    # เส้น 21:40 (C18)
    hh, mm = map(int, post_time(C18).replace(" น.", "").split(":"))
    mx = x(DAYS.index(FRI) * 24 + hh + mm / 60)
    parts.append(f'<line x1="{mx:.1f}" x2="{mx:.1f}" y1="{pt - 26}" y2="{H - pb}" stroke="{T["ink"]}" stroke-dasharray="4 3" stroke-width="1.5"/>')
    parts.append(f'<text x="{mx - 6:.1f}" y="{pt - 30}" text-anchor="end" class="anno">{esc(post_time(C18))} รีวิวแรกบน {esc(C18["แพลตฟอร์ม"])} (C18)</text>')
    pi = HOURS.index(st["peak_key"])
    anchor = "end" if pi > n * 0.8 else "start"
    dx = -10 if anchor == "end" else 10
    parts.append(f'<circle cx="{x(pi):.1f}" cy="{y(st["peak"]):.1f}" r="4.5" fill="{T["primary"]}" stroke="#fff" stroke-width="2"/>')
    parts.append(f'<text x="{x(pi) + dx:.1f}" y="{y(st["peak"]) - 2:.1f}" text-anchor="{anchor}" class="anno">{esc(st["peak_th"])} · {st["peak"]} ครั้ง (ลบ {st["peak_neg"]})</text>')
    parts.append(f'<text x="{x(pi) + dx:.1f}" y="{y(st["peak"]) + 14:.1f}" text-anchor="{anchor}" class="ax">ปกติชั่วโมงนี้ (จ.–พฤ.) เฉลี่ย {st["base_peak_hour"]:.1f} ครั้ง</text>')
    scope = "ทุกสาขา" if st["branch"] is None else branch_name(st["branch"])
    label = (f"กราฟเส้นการกล่าวถึงรายชั่วโมง 19–25 ต.ค. 69 {scope} รวม {fmt(st['total'])} ครั้ง เชิงลบ {fmt(st['neg'])} ครั้ง "
             f"จุดสูงสุด {st['peak_th']} {st['peak']} ครั้ง (ลบ {st['peak_neg']}) "
             f"เส้นประบอกเวลา {post_time(C18)} ที่รีวิวแรก C18 บน Google Maps")
    return (f'<svg class="chart" role="img" aria-label="{esc(label)}" viewBox="0 0 {W} {H}" preserveAspectRatio="xMidYMid meet">'
            + "".join(parts) + "</svg>")


def svg_platform(st=None):
    """F28 ทุกส่วนของแท่งโฟกัสได้และมีตัวเลขเมื่อชี้/โฟกัส (data-tip) · F24 วาดรายสาขาได้"""
    st = st or ST_ALL
    # WCAG B7 2.5.8 แท่งที่โฟกัสได้สูง 28 หน่วย (ไม่ต่ำกว่า 24px ที่ 1280 ขึ้นไป)
    rowh, W, lw, bw, sh = 34, 460, 96, 300, 28
    H = rowh * len(st["plat"]) + 8
    parts = []
    for i, (p, c, tot) in enumerate(st["plat"]):
        y0 = 4 + i * rowh
        parts.append(f'<text x="0" y="{y0 + 22}" class="lab">{esc(p)}</text>')
        xx = lw
        for s in SENT_ORDER:
            v = c[s]
            w = bw * v / tot
            fill = SENT_COLOR[s]
            stroke = f' stroke="{NEU_STROKE}" stroke-width="1"' if s == "เป็นกลาง" else ' stroke="#fff" stroke-width="1"'
            tip = f"{p} · {s} {fmt(v)} จาก {fmt(tot)} ({round(v / tot * 100)}%)"
            parts.append(f'<rect class="seg-hit" tabindex="0" role="img" aria-label="{esc(tip)}" data-tip="{esc(tip)}" x="{xx:.1f}" y="{y0 + 3}" width="{max(w, 0.1):.1f}" height="{sh}"{stroke} fill="{fill}"/>')
            if w >= 36:
                lc = label_color(fill)
                if lc:
                    parts.append(f'<text x="{xx + w / 2:.1f}" y="{y0 + 22}" text-anchor="middle" class="seg" fill="{lc}" pointer-events="none">{round(v / tot * 100)}%</text>')
            xx += w
        parts.append(f'<text x="{W}" y="{y0 + 22}" text-anchor="end" class="lab num">{fmt(tot)}</text>')
    label = "กราฟแท่งซ้อน ความรู้สึกแยกตามแพลตฟอร์ม 19–25 ต.ค. 69 กดแท็บเพื่ออ่านตัวเลขทีละส่วน หรือใช้ปุ่มดูตารางข้อมูล"
    return (f'<svg class="chart" role="group" aria-label="{esc(label)}" viewBox="0 0 {W} {H}">' + "".join(parts) + "</svg>")


def platform_table(st=None):
    """F28 ตารางข้อมูลของกราฟแพลตฟอร์ม"""
    st = st or ST_ALL
    rows = "".join(f'<tr><th scope="row">{esc(p)}</th>' + "".join(f'<td class="num">{fmt(c[s])}</td>' for s in SENT_ORDER) + f'<td class="num"><b>{fmt(t)}</b></td></tr>' for p, c, t in st["plat"])
    return (f'<table class="data-tbl"><caption class="sr-only">ความรู้สึกแยกตามแพลตฟอร์ม</caption><thead><tr><th scope="col">แพลตฟอร์ม</th>'
            + "".join(f'<th scope="col">{s}</th>' for s in SENT_ORDER) + f'<th scope="col">รวม</th></tr></thead><tbody>{rows}</tbody></table>')


def branch_name(b):
    return "ไม่ระบุสาขา" if b == "ไม่ระบุ" else f"สาขา{b}"


def svg_branch():
    rowh, W, lw, bw = 30, 460, 100, 250
    mx = max(t for _, _, t in S["branches"])
    H = rowh * len(S["branches"]) + 8
    parts = []
    for i, (b, ng, tot) in enumerate(S["branches"]):
        y0 = 4 + i * rowh
        parts.append(f'<text x="0" y="{y0 + 20}" class="lab">{esc(branch_name(b))}</text>')
        parts.append(f'<rect x="{lw}" y="{y0 + 6}" width="{bw * tot / mx:.1f}" height="20" rx="3" fill="{T["band"]}"/>')
        parts.append(f'<rect x="{lw}" y="{y0 + 6}" width="{bw * ng / mx:.1f}" height="20" rx="3" fill="{T["neg"]}"/>')
        parts.append(f'<text x="{W}" y="{y0 + 20}" text-anchor="end" class="lab num"><tspan font-weight="700">{fmt(ng)}</tspan> / {fmt(tot)}</text>')
    label = "กราฟแท่ง เชิงลบแยกตามสาขา 19–25 ต.ค. 69: " + "; ".join(
        f"{branch_name(b)} ลบ {ng} จากทั้งหมด {tot}" for b, ng, tot in S["branches"])
    return (f'<svg class="chart" role="img" aria-label="{esc(label)}" viewBox="0 0 {W} {H}">' + "".join(parts) + "</svg>")


def svg_campaign(st=None):
    st = st or ST_ALL  # F24
    days = st["camp_days"]
    W, H, pb, pt = 420, 170, 26, 22
    mx = max(max(st["camp_daily"][d] for d in days), 1)
    bw = 34
    gap = (W - bw * len(days)) / (len(days))
    parts = []
    for i, d in enumerate(days):
        v = st["camp_daily"][d]
        h = (H - pb - pt) * v / mx
        x0 = gap / 2 + i * (bw + gap)
        parts.append(f'<rect x="{x0:.1f}" y="{H - pb - h:.1f}" width="{bw}" height="{h:.1f}" rx="4" fill="{T["primary"]}"/>')
        parts.append(f'<text x="{x0 + bw / 2:.1f}" y="{H - pb - h - 6:.1f}" text-anchor="middle" class="lab num">{v}</text>')
        parts.append(f'<text x="{x0 + bw / 2:.1f}" y="{H - 6}" text-anchor="middle" class="ax">{esc(DAY_TH[d].split(" ")[0] + " " + DAY_TH[d].split(" ")[1])}</text>')
    label = "กราฟแท่ง การกล่าวถึงแคมเปญโปร 1 แถม 1 สมาชิกรายวัน: " + "; ".join(
        f"{DAY_TH[d]} {st['camp_daily'][d]}" for d in days)
    return (f'<svg class="chart" role="img" aria-label="{esc(label)}" viewBox="0 0 {W} {H}">' + "".join(parts) + "</svg>")


def svg_fri_bars():
    W, H, pb, pt = 326, 168, 26, 24
    data = S["fri_hours"]
    mx = max(t for _, t, _ in data)
    bw = 44
    gap = (W - bw * len(data)) / len(data)
    parts = []
    for i, (h, tot, ng) in enumerate(data):
        x0 = gap / 2 + i * (bw + gap)
        ht = (H - pb - pt) * tot / mx
        hn = (H - pb - pt) * ng / mx
        parts.append(f'<rect x="{x0:.1f}" y="{H - pb - ht:.1f}" width="{bw}" height="{ht:.1f}" fill="{T["neu"]}" stroke="{NEU_STROKE}" stroke-width="1"/>')
        parts.append(f'<rect x="{x0:.1f}" y="{H - pb - hn:.1f}" width="{bw}" height="{hn:.1f}" fill="{T["neg"]}"/>')
        parts.append(f'<text x="{x0 + bw / 2:.1f}" y="{H - pb - ht - 6:.1f}" text-anchor="middle" class="lab num"><tspan font-weight="700">{tot}</tspan> / {ng}</text>')
        parts.append(f'<text x="{x0 + bw / 2:.1f}" y="{H - 6}" text-anchor="middle" class="ax">{h:02d}:00</text>')
    label = "กราฟแท่ง การกล่าวถึงรายชั่วโมงคืนวันศุกร์ 23 ต.ค. 69: " + "; ".join(
        f"{h:02d}:00 น. ทั้งหมด {t} ลบ {n}" for h, t, n in data)
    return (f'<svg class="chart" role="img" aria-label="{esc(label)}" viewBox="0 0 {W} {H}">' + "".join(parts) + "</svg>")


def svg_spark():
    W, H = 260, 70
    vals = [S["daily"][d] for d in DAYS]
    mx = max(vals)
    pts = " ".join(f"{8 + i * (W - 16) / 6:.1f},{H - 8 - (H - 16) * v / mx:.1f}" for i, v in enumerate(vals))
    label = "กราฟเส้นย่อ ยอดรายวัน 19–25 ต.ค. 69: " + ", ".join(f"{DAY_TH[d]} {S['daily'][d]}" for d in DAYS)
    return (f'<svg class="chart" role="img" aria-label="{esc(label)}" viewBox="0 0 {W} {H}">'
            f'<polyline points="{pts}" fill="none" stroke="{T["primary"]}" stroke-width="2"/></svg>')


# ---------------------------------------------------------------- page shells
FONTS = ('<link rel="preconnect" href="https://fonts.googleapis.com">'
         '<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>'
         '<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=IBM+Plex+Sans+Thai:wght@400;500;600;700&display=swap">')


def page(title, body, pid, kind):
    return f"""<!DOCTYPE html>
<html lang="th">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{esc(title)} · ชบา Monitor (ต้นแบบ)</title>
{FONTS}
<link rel="stylesheet" href="tokens.css">
<script src="app.js" defer></script>
</head>
<body class="k-{kind}" data-page="{pid}">
<a class="skip" href="#main">ข้ามไปเนื้อหาหลัก</a>
<nav aria-label="นำทางต้นแบบ"><a class="home-link" href="index.html">{icon('left', size=14)}กลับหน้าเริ่ม</a></nav><!-- WCAG A1 ลิงก์อยู่ใน landmark -->
{body}
</body>
</html>
"""


# F22 เมนูเว็บใช้ชื่อชุดเดียวกับแท็บมือถือ (NAV_LABEL) · Alerts & Cases เปิดได้แล้ว
NAV = [("w1", NAV_LABEL["overview"], "grid", "w1.html"), ("w2", NAV_LABEL["posts"], "list", "w2.html"),
       ("cases", NAV_LABEL["cases"], "bell", "w-cases.html"), ("settings", "ตั้งค่า", "sliders", None)]


def web_shell(pid, crumbs, content, active):
    nav = []
    for key, label, ic, href in NAV:
        if key == "settings":
            nav.append('<li class="sep" aria-hidden="true"></li>')
        cur = ' aria-current="page"' if key == active else ""
        badge = f'<span class="count" aria-label="เคสระดับ{SEV_CRIT}ที่ยังไม่มีผู้รับทราบ">K10</span>' if key == "cases" else ""  # F04
        if href:
            nav.append(f'<li><a class="nav-item" href="{href}"{cur}>{icon(ic)}<span>{label}</span>{badge}</a></li>')
        else:
            nav.append(f'<li><span class="nav-item static" aria-disabled="true">{icon(ic)}<span>{label}<small>ภาพนิ่ง</small></span>{badge}</span></li>')
    nav.append('<li class="nav-note">เมนูรอง แสดงตามสิทธิ์</li>')
    crumb_html = []
    for i, (label, href) in enumerate(crumbs):
        if href and i < len(crumbs) - 1:
            crumb_html.append(f'<li><a href="{href}">{esc(label)}</a></li>')
        else:
            crumb_html.append(f'<li aria-current="page">{esc(label)}</li>')
    return f"""
<div class="web">
  <aside class="side">
    <div class="brand"><span class="logo">{icon('pulse', size=20)}</span><span><b>ชบา Monitor</b><small>Social Listening & Insights</small></span></div>
    <nav aria-label="เมนูหลัก"><ul>{''.join(nav)}</ul></nav>
  </aside>
  <div class="web-main">
    <header class="topbar">
      <nav aria-label="เส้นทางบอกตำแหน่ง"><ol class="crumbs">{''.join(crumb_html)}</ol></nav>
      <div class="top-right">
        <span class="asof">{icon('refresh', size=16)}ข้อมูลถึง {DATA_AS_OF}</span>
        <form class="search" role="search" onsubmit="return false">
          <label for="q" class="sr-only">ค้นหาโพสต์หรือเคส</label>
          {icon('search', size=18)}<input id="q" type="search" placeholder="ค้นหาโพสต์หรือเคส เช่น C09, K10">
        </form>
        <button class="icon-btn" type="button" aria-label="การแจ้งเตือน (ภาพนิ่ง)">{icon('bell', size=20)}</button>
        <span class="me"><span class="avatar">{icon('user', label='ผู้ใช้', size=18)}</span><span><b>ทีม Customer Support</b><small>บทบาท: CS · เวลางาน</small></span></span>
      </div>
    </header>
    <main id="main" class="content">{content}</main>
  </div>
</div>"""


def demo_panel():
    """F16 แผง "โหมดทดลองใช้งาน" นอกกรอบโทรศัพท์ (แทน "สาธิต: กด E" บนจอ) · F03 สวิตช์ซ่อนรายละเอียดบนหน้าจอล็อก"""
    return f"""<aside class="demo" aria-label="โหมดทดลองใช้งาน"><!-- WCAG A1 แผงอยู่ใน landmark -->
  <h2 id="demo-h">{icon('sliders', size=16)}โหมดทดลองใช้งาน <small>สำหรับผู้สังเกตการณ์ · ไม่ใช่ส่วนของแอป</small></h2>
  <label class="sw-row"><input type="checkbox" role="switch" id="demo-fail"><span><b>จำลองส่งไม่สำเร็จ</b><small>การส่งครั้งถัดไปไม่สำเร็จ แล้วปิดสวิตช์เอง · ปุ่มลัด E</small></span></label>
  <label class="sw-row"><input type="checkbox" role="switch" id="demo-hide"><span><b>ซ่อนรายละเอียดบนหน้าจอล็อก</b><small>push แสดงแค่ระดับและจำนวนเรื่อง (m1)</small></span></label>
  <label class="sw-row sel">ขนาดจอ<select id="demo-size"><option value="390x844">390 × 844</option><option value="360x740">360 × 740</option></select></label>
  <p class="demo-out" id="demo-out" role="status"></p>
</aside>"""


def phone_shell(pid, clock, content, note, header=None, actions="", tabbar="", noscroll=False, extra=""):
    head = ""
    if header:
        back, title = header
        bk = (f'<a class="m-back" href="{back}" aria-label="ย้อนกลับ">{icon("left", size=22)}</a>' if back else '<span class="m-back-sp"></span>')
        head = f'<header class="m-head">{bk}<h1>{esc(title)}</h1><span class="m-back-sp"></span></header>'
    return f"""
<main id="main" class="stage">
  <div class="phone" role="region" aria-label="หน้าจอมือถือ 390 × 844">
    <div class="statusbar"><span>{clock}</span><span class="sb-r">ชบา Monitor</span></div>
    {head}
    <div class="m-body{' no-scroll' if noscroll else ''}" id="m-body">{content}</div>
    {f'<div class="m-actions">{actions}</div>' if actions else ''}
    {tabbar}
    {extra}
  </div>
  <aside class="phone-note" aria-label="คำอธิบายหน้าจอ">{demo_panel()}{note}</aside>
</main>"""


def tabbar(active):
    # F22 ชื่อแท็บมือถือ = ชื่อเมนูเว็บ
    items = [(NAV_LABEL["overview"], "grid", "m3.html"), (NAV_LABEL["posts"], "list", "m3.html#key-posts"), (NAV_LABEL["cases"], "bell", "m2.html")]
    out = []
    for label, ic, href in items:
        cur = ' aria-current="page"' if label == active else ""
        out.append(f'<a href="{href}"{cur}>{icon(ic, size=22)}<span>{label}</span></a>')
    return f'<nav class="m-tabs" aria-label="เมนูหลักมือถือ">{"".join(out)}</nav>'


STEPS = ["เปิดแจ้งเตือน", "ประเมิน", "ตัดสินความเร่งด่วน", "ตรวจสาขา", "ตรวจผู้รับ", "ทบทวน", "ยืนยันส่ง", "ติดตาม"]
STEP_PAGE = {"เปิดแจ้งเตือน": "m1, m2", "ประเมิน": "m3 (ข้อมูลตัดสินใจเบื้องต้นอยู่ใน m2)",
             "ตัดสินความเร่งด่วน": "m2", "ตรวจสาขา": "m4 (ส่วนบน)", "ตรวจผู้รับ": "m4",
             "ทบทวน": "m5", "ยืนยันส่ง": "m5", "ติดตาม": "m6"}


def step_note(here, desc):
    lis = []
    for i, s in enumerate(STEPS, 1):
        on = s in here
        lis.append(f'<li{" class=\"on\" aria-current=\"step\"" if on else ""}><span>{i}</span>{s}</li>')
    return (f'<p class="pn-k">กรณีที่ 1 · ขั้นใน User Flow</p><ol class="flow">{"".join(lis)}</ol>'
            f'<div class="pn-d">{desc}</div>')


# ---------------------------------------------------------------- CSS
CSS_VARS = ":root{" + "".join(f"--{k}:{v};" for k, v in T.items()) + "}\n"
CSS = CSS_VARS + r"""
*{box-sizing:border-box}
[hidden]{display:none!important}
html{-webkit-text-size-adjust:100%}
body{margin:0;font-family:"IBM Plex Sans Thai",system-ui,sans-serif;font-size:15px;line-height:1.6;color:var(--ink);background:var(--bg);overflow-wrap:break-word}
h1,h2,h3,h4{line-height:1.5;margin:0;font-weight:700}
h1{font-size:26px}h2{font-size:18px}h3{font-size:16px}
p{margin:0}
a{color:var(--primary)}
small{font-size:13px}
button,input,select{font:inherit;color:inherit}
:focus-visible{outline:3px solid var(--primary);outline-offset:2px;border-radius:6px}
.sr-only{position:absolute;width:1px;height:1px;overflow:hidden;clip:rect(0 0 0 0);white-space:nowrap}
.skip{position:absolute;left:-9999px}.skip:focus{left:12px;top:12px;z-index:99;background:#fff;padding:8px 12px}
.ic{flex:none;vertical-align:middle}
mark{background:var(--lime);color:var(--ink);border-radius:3px;padding:0 2px}
.emo{font-size:12px;color:var(--ink3);background:var(--bg);border-radius:4px;padding:0 4px;white-space:nowrap}
.num{font-variant-numeric:tabular-nums}
.muted{color:var(--ink3)}.sub{color:var(--ink2)}
.meta{display:block;font-size:12px;color:var(--ink3);line-height:1.5}
.home-link{position:fixed;left:12px;bottom:12px;z-index:50;display:inline-flex;align-items:center;gap:4px;min-height:28px;padding:2px 10px;font-size:13px;background:#fff;border:1px solid var(--line);border-radius:999px;text-decoration:none}

/* buttons */
.btn{display:inline-flex;align-items:center;justify-content:center;gap:8px;min-height:40px;padding:8px 16px;border-radius:10px;border:1px solid var(--control);background:#fff;color:var(--ink);font-weight:600;text-decoration:none;cursor:pointer;white-space:nowrap}
.btn.primary{background:var(--primary);border-color:var(--primary);color:#fff}
.btn.soft{background:var(--primaryTint);border-color:transparent;color:var(--primary)}
.btn.ghost{border-color:transparent;background:transparent;color:var(--primary)}
.btn.lime{background:var(--lime);border-color:var(--lime);color:var(--ink)}
.btn[disabled],.btn[aria-disabled=true]{opacity:.55;cursor:not-allowed}
.btn.sm{min-height:32px;padding:4px 12px;font-size:14px}
.icon-btn{display:inline-flex;align-items:center;justify-content:center;width:40px;height:40px;border-radius:10px;border:1px solid var(--line);background:#fff;cursor:pointer}
.tag{display:inline-flex;align-items:center;gap:4px;font-size:13px;font-weight:600;padding:1px 8px;border-radius:999px;white-space:nowrap;line-height:1.6}
.tag.warn{color:var(--warn);background:var(--warnTint)}
.tag.crit{color:#fff;background:var(--crit)}
.tag.ok{color:var(--ok);background:var(--okTint)}
.tag.pri{color:var(--primary);background:var(--primaryTint)}
.tag.plain{color:var(--ink2);background:var(--bg)}
.sent{display:inline-flex;align-items:center;gap:6px;font-size:14px;font-weight:600;padding:1px 10px;border:1px solid var(--line);border-radius:999px;background:#fff;white-space:nowrap}
.sent .dot{width:10px;height:10px;border-radius:50%}
.conf{display:inline-flex;align-items:center;gap:8px;flex-wrap:wrap}
.conf .meter{width:56px;height:6px;border-radius:3px;background:var(--line);overflow:hidden}
.conf .meter span{display:block;height:100%;background:var(--primary)}
.conf.low .meter span{background:var(--warn)}
.card{background:var(--card);border:1px solid var(--line);border-radius:16px;padding:20px}
.card h2+p.sub,.card h2+.sub{margin-top:2px}
.legend{display:flex;flex-wrap:wrap;gap:4px 16px;font-size:13px;color:var(--ink2);margin:8px 0}
.legend span{display:inline-flex;align-items:center;gap:6px}
.legend i{display:inline-block;width:12px;height:12px;border-radius:2px}
.legend i.line{height:3px;width:16px;border-radius:2px}
.chart{display:block;width:100%;height:auto;font-family:inherit}
.chart .ax{font-size:11px;fill:var(--ink3)}
.chart .lab{font-size:13px;fill:var(--ink)}
.chart .seg{font-size:12px;font-weight:600}
.chart .anno{font-size:12px;font-weight:600;fill:var(--ink)}
.live{min-height:0}
.placeholder{font-weight:600;color:var(--warn);background:var(--warnTint);border-radius:4px;padding:0 4px}

/* ---------- web ---------- */
.k-web{min-width:1280px}
.web{display:grid;grid-template-columns:232px 1fr;min-height:100vh;background:linear-gradient(90deg,#fff 231px,var(--line) 231px,var(--line) 232px,transparent 232px)}
.side{background:#fff;padding:20px 14px;position:sticky;top:0;height:100vh}
.brand{display:flex;gap:10px;align-items:center;margin-bottom:24px}
.brand b{display:block;font-size:15px;line-height:1.4}.brand small{display:block;color:var(--ink3);line-height:1.4}
.logo{display:inline-flex;width:36px;height:36px;border-radius:10px;background:var(--primary);color:var(--lime);align-items:center;justify-content:center}
.side ul{list-style:none;margin:0;padding:0;display:flex;flex-direction:column;gap:4px}
.nav-item{display:flex;align-items:center;gap:10px;min-height:44px;padding:8px 12px;border-radius:10px;color:var(--ink2);text-decoration:none;font-weight:500}
.nav-item[aria-current=page]{background:var(--primaryTint);color:var(--primary);font-weight:700}
a.nav-item:hover{background:var(--bg)}
.nav-item.static small{display:block;font-size:12px;color:var(--ink3);font-weight:400;line-height:1.4}
.nav-item .count{margin-left:auto;font-size:12px;font-weight:700;color:#fff;background:var(--crit);border-radius:999px;padding:0 7px}
.side .sep{border-top:1px solid var(--line);margin:8px 6px}
.nav-note{font-size:12px;color:var(--ink3);padding:4px 12px}
.web-main{min-width:0}
.topbar{display:flex;align-items:center;justify-content:space-between;gap:16px;padding:12px 28px;background:#fff;border-bottom:1px solid var(--line);position:sticky;top:0;z-index:5}
.crumbs{display:flex;gap:6px;list-style:none;margin:0;padding:0;font-size:14px;color:var(--ink2)}
.crumbs li+li::before{content:"›";margin-right:6px;color:var(--ink3)}
.top-right{display:flex;align-items:center;gap:12px}
.asof{display:inline-flex;align-items:center;gap:6px;font-size:13px;color:var(--ink2);white-space:nowrap}
.search{display:flex;align-items:center;gap:8px;background:var(--bg);border:1px solid var(--control);border-radius:10px;padding:0 10px;min-height:40px}
.search input{border:0;background:transparent;min-width:270px;min-height:36px}
.search input:focus{outline:none}.search:focus-within{outline:3px solid var(--primary);outline-offset:2px}
.me{display:flex;align-items:center;gap:8px}.me b{display:block;font-size:14px;line-height:1.4}.me small{display:block;color:var(--ink3);line-height:1.4;font-size:12px}
.avatar{display:inline-flex;width:36px;height:36px;border-radius:50%;background:var(--lime);align-items:center;justify-content:center}
.content{padding:24px 28px 72px;display:flex;flex-direction:column;gap:20px;max-width:1360px}
.page-head{display:flex;justify-content:space-between;align-items:flex-end;gap:16px}
.tabs{display:flex;gap:4px;border-bottom:1px solid var(--line)}
.tabs a{display:inline-flex;align-items:center;min-height:44px;padding:8px 14px;color:var(--ink2);text-decoration:none;border-bottom:3px solid transparent;font-weight:500}
.tabs a[aria-current=page]{color:var(--primary);border-bottom-color:var(--primary);font-weight:700}
.tabs .static{display:inline-flex;align-items:center;min-height:44px;padding:8px 14px;color:var(--ink3)}
.filters{display:flex;flex-wrap:wrap;gap:12px;align-items:flex-end}
.field{display:flex;flex-direction:column;gap:4px;font-size:13px;color:var(--ink2)}
.field select,.field input[type=text]{min-height:40px;border:1px solid var(--control);border-radius:10px;background:#fff;padding:6px 10px;font-size:14px;color:var(--ink);min-width:150px}
.check{display:inline-flex;align-items:center;gap:8px;min-height:40px;font-size:14px;color:var(--ink)}
.check input{width:24px;height:24px;accent-color:var(--primary);flex:none}
.saved{display:flex;flex-wrap:wrap;align-items:center;gap:8px;margin-top:14px;padding-top:14px;border-top:1px solid var(--line);font-size:13px;color:var(--ink2)}
.chip{display:inline-flex;align-items:center;gap:6px;min-height:32px;padding:4px 12px;border:1px solid var(--control);border-radius:999px;background:#fff;font-size:14px;cursor:pointer}
/* WCAG A3 2.5.8 เป้าหมายอย่างน้อย 24×24 */
input[type=checkbox],input[type=radio]{min-width:24px;min-height:24px}
.solo-link,.crumbs a{display:inline-flex;align-items:center;min-height:24px}
/* WCAG B6 ชิปตัวกรองที่ใช้อยู่ (span) ใช้คลาส .on แทน aria-pressed */
.chip.on,.chip[aria-pressed=true]{background:var(--primaryTint);border-color:var(--primary);color:var(--primary);font-weight:600}
/* F25 F26 แถบเคสอยู่บนสุดและเด่นที่สุดในหน้า */
.alertbar{display:flex;align-items:flex-start;gap:16px;background:var(--critTint);border:1px solid #F4C7C1;border-left:6px solid var(--crit);border-radius:14px;padding:16px 20px;color:var(--crit)}
.alertbar>b{white-space:nowrap;font-size:17px}
.alertbar ul{list-style:none;margin:0;padding:0;flex:1;display:flex;flex-direction:column;gap:6px;color:var(--ink);font-size:15px}
.alertbar ul a{font-weight:700}
.alertbar .more{white-space:nowrap;font-size:14px;display:inline-flex;align-items:center;gap:4px;min-height:32px}
.kpis{display:grid;grid-template-columns:repeat(4,1fr);gap:16px}
.kpi{background:#fff;border:1px solid var(--line);border-radius:16px;padding:16px 18px}
.kpi .k{font-size:14px;color:var(--ink2)}.kpi .v{font-size:32px;font-weight:700;line-height:1.3}.kpi .d{font-size:13px;color:var(--ink2)}
.kpi.hot{background:#fff;border:2px solid var(--primary);padding:15px 17px}.kpi.hot .k{color:var(--primary);font-weight:600} /* F26 พื้นขาว + เส้นขอบสีเน้น */
.grid2{display:grid;grid-template-columns:1fr 1fr;gap:20px}
.chart-foot{display:flex;align-items:center;gap:16px;flex-wrap:wrap;margin-top:10px;font-size:13px;color:var(--ink2)}
.cov{list-style:none;margin:0;padding:0}
.cov li{display:grid;grid-template-columns:150px 1fr;gap:12px;padding:12px 0;border-top:1px solid var(--line);font-size:14px}
.cov li:first-child{border-top:0}
.foot-note{font-size:13px;color:var(--ink3)}
/* table */
/* R3 ตารางไม่ถูกตัดขอบ: เลื่อนแนวนอนได้เมื่อจอแคบ */
.tbl-wrap{background:#fff;border:1px solid var(--line);border-radius:16px;overflow-x:auto}
table{border-collapse:collapse;width:100%}
th,td{text-align:left;vertical-align:top;padding:12px 8px;border-top:1px solid var(--line);font-size:14px}
thead th{border-top:0;font-size:13px;color:var(--ink2);font-weight:600;background:#FBFAFE}
td .meta{display:block;font-size:12px;color:var(--ink3)}
td.msg{min-width:240px}
td.msg .tag{white-space:normal;text-align:left}
td.idcell{min-width:96px;max-width:120px}
td.confcell .conf{flex-wrap:wrap;max-width:118px}
td.ownercell{min-width:84px}
td.platcell{min-width:130px;max-width:170px}
td.nowrap{white-space:nowrap}
td.brcell{min-width:84px;word-break:keep-all}
td .conf{flex-wrap:nowrap}
.postid{font-weight:700}
a.postid{display:inline-flex;align-items:center;min-height:24px}
tr.is-hidden{display:none}
.legend-line{display:flex;flex-wrap:wrap;gap:6px 18px;font-size:13px;color:var(--ink2);align-items:center}
.kwchips{display:flex;flex-wrap:wrap;gap:4px;margin-top:4px}
.kwchip{font-size:12px;border:1px dashed var(--control);border-radius:4px;padding:0 6px;color:var(--ink2)}
.ex-row{display:grid;grid-template-columns:1.4fr 1fr auto;gap:20px;padding:18px 0;border-top:1px solid var(--line);align-items:center}
.ex-row:first-of-type{border-top:0}
.ex-row .q{font-size:17px;margin:4px 0}
.ex-row.restored{background:var(--okTint);margin:0 -20px;padding:18px 20px}
/* w3 */
.detail{display:grid;grid-template-columns:1fr 400px;gap:20px;align-items:start}
.detail-col{display:flex;flex-direction:column;gap:20px;min-width:0}
.quote{font-size:24px;font-weight:600;line-height:1.6;margin:10px 0}
.res-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin-top:12px}
.res-grid dt{font-size:13px;color:var(--ink2)}.res-grid dd{margin:4px 0 0}
.explain{margin-top:14px;background:var(--bg);border-radius:10px;padding:10px 14px;font-size:14px;color:var(--ink2)}
.sim{list-style:none;margin:10px 0 0;padding:0}
.sim li{display:grid;grid-template-columns:52px 1fr auto;gap:12px;padding:12px 0;border-top:1px solid var(--line);align-items:center;font-size:14px}
.hist{list-style:none;margin:10px 0 0;padding:0}
.hist li{display:grid;grid-template-columns:200px 1fr;gap:12px;padding:6px 0;font-size:14px}
.hist li time{color:var(--ink2)}
.hist li.new{background:var(--primaryTint);border-radius:8px;padding:6px 8px}
/* F42 ลดระยะห่างในแผงแก้ผล */
.panel{position:sticky;top:84px;background:#fff;border:2px solid var(--primary);border-radius:18px;padding:18px;display:flex;flex-direction:column;gap:12px}
.panel fieldset{border:0;margin:0;padding:0;display:flex;flex-direction:column;gap:6px}
.panel legend{font-weight:700;font-size:15px;padding:0;margin-bottom:6px}
.panel .btnrow{display:grid;grid-template-columns:2fr 1fr;gap:8px} /* F42 บันทึก + ยกเลิก แถวเดียว */
.panel .btnrow .btn{min-height:48px}
.saved-sum{display:flex;flex-direction:column;gap:10px} /* F40 F41 สรุปหลังบันทึก */
.saved-sum .chg{display:flex;align-items:center;gap:8px;flex-wrap:wrap;font-size:15px;font-weight:700}
.saved-sum dl{display:grid;grid-template-columns:96px 1fr;gap:4px 8px;margin:0;font-size:14px}
.saved-sum dt{color:var(--ink2)}.saved-sum dd{margin:0}
.panel legend small{font-weight:400;color:var(--ink2)}
.seg-opts{display:flex;flex-wrap:wrap;gap:8px}
.seg-opts label{position:relative}
.seg-opts input{position:absolute;opacity:0;inset:0;width:100%;height:100%;margin:0;cursor:pointer}
.seg-opts span{display:inline-flex;align-items:center;min-height:40px;padding:6px 14px;border:1px solid var(--control);border-radius:999px;background:#fff;font-size:14px;font-weight:500}
.seg-opts input:checked+span{background:var(--primary);border-color:var(--primary);color:#fff;font-weight:700}
.seg-opts input:focus-visible+span{outline:3px solid var(--primary);outline-offset:2px}
.diff{background:var(--bg);border-radius:12px;padding:10px 14px;font-size:14px;display:flex;flex-direction:column;gap:4px}
.diff b{font-weight:700}
.panel .hint{font-size:13px;color:var(--ink2)}
.panel .row{display:flex;gap:10px;flex-wrap:wrap}
.notice{border-radius:12px;padding:10px 14px;font-size:14px;display:flex;gap:10px;align-items:flex-start}
.notice.ok{background:var(--okTint);color:var(--ok)}
.notice.err{background:var(--critTint);color:var(--crit)}
.notice.info{background:var(--primaryTint);color:var(--primary)}
.notice.warn{background:var(--warnTint);color:var(--warn)}
.notice .grow{flex:1}
.notice p{color:inherit}
.demo-flag{font-size:12px;color:var(--warn)}
.demo-hint{font-size:13px;color:var(--ink3);margin-top:8px}
/* F37 ตั้งค่าโพสต์นี้: หน้าตาเป็นค่าของโพสต์ ไม่ใช่ตัวกรอง */
.setup{display:flex;align-items:center;gap:20px;flex-wrap:wrap;padding:14px 20px}
.setup h2{font-size:15px;margin-right:4px}
.setup .opt{display:flex;align-items:center;gap:8px;font-size:14px;color:var(--ink2)}
.setup select{min-height:40px;border:1px solid transparent;border-bottom:2px solid var(--control);border-radius:8px 8px 0 0;background:var(--bg);padding:6px 10px;font-weight:700;color:var(--ink)}
.setup .muted{margin-left:auto;font-size:13px}
/* F33 แถบคำสั่งกลุ่ม */
.bulk{display:flex;align-items:center;gap:12px;flex-wrap:wrap;background:var(--primaryTint);border:1px solid var(--primary);border-radius:14px;padding:10px 16px;position:sticky;top:66px;z-index:4}
.bulk b{color:var(--primary)}
.bulk .field{flex-direction:row;align-items:center;gap:6px}
.bulk .field select{min-width:0}
.fchips{display:flex;flex-wrap:wrap;gap:8px;align-items:center}
.fchip{display:inline-flex;align-items:center;gap:4px;min-height:32px;padding:2px 4px 2px 12px;border-radius:999px;background:var(--primaryTint);color:var(--primary);font-size:14px;font-weight:600}
.fchip button{display:inline-flex;align-items:center;justify-content:center;width:28px;height:28px;border:0;border-radius:50%;background:transparent;color:var(--primary);cursor:pointer}
.fchip button:hover{background:#fff}
/* F28 tooltip + ตาราง */
.seg-hit{cursor:pointer}
.seg-hit:focus{outline:none}.seg-hit:focus-visible{stroke:var(--ink);stroke-width:3}
.tip{position:fixed;z-index:20;pointer-events:none;background:var(--ink);color:#fff;font-size:13px;border-radius:8px;padding:4px 10px;white-space:nowrap}
.data-tbl{margin-top:10px;border:1px solid var(--line);border-radius:10px;overflow:hidden}
.data-tbl th,.data-tbl td{padding:6px 10px;font-size:13px}
.card-acts{display:flex;align-items:center;gap:12px;flex-wrap:wrap;margin-top:10px}
/* F22 Alerts & Cases */
.case-list{display:flex;flex-direction:column;gap:12px}
.case-row{background:#fff;border:1px solid var(--line);border-radius:16px}
.case-row.crit{border-left:6px solid var(--crit)}
.case-btn{display:grid;grid-template-columns:70px 90px 1fr auto;gap:16px;align-items:center;width:100%;padding:16px 20px;border:0;background:transparent;text-align:left;cursor:pointer;border-radius:16px}
.case-btn .cid{font-size:20px;font-weight:700}
.case-btn .chev{transition:transform .15s}
.case-btn[aria-expanded=true] .chev{transform:rotate(90deg)}
.case-detail{padding:0 20px 20px;display:flex;flex-direction:column;gap:12px}
.case-detail dl{display:grid;grid-template-columns:140px 1fr;gap:6px 12px;margin:0;font-size:14px}
.case-detail dt{color:var(--ink2)}.case-detail dd{margin:0}
/* states */
.states{display:grid;grid-template-columns:160px repeat(4,1fr);gap:14px;align-items:stretch}
.states .rowh{font-weight:700;padding-top:8px}
.states .colh{font-weight:700;display:flex;align-items:center;gap:8px}
.states .colh i{width:10px;height:10px;border-radius:50%;display:inline-block}
.st{background:#fff;border:1px solid var(--line);border-radius:14px;padding:14px;display:flex;flex-direction:column;gap:10px;font-size:14px}
.st .mock{background:var(--bg);border-radius:10px;padding:12px;min-height:110px;display:flex;flex-direction:column;gap:8px;justify-content:center}
.st dl{display:grid;grid-template-columns:64px 1fr;gap:4px 8px;margin:0;font-size:13px}
.st dt{color:var(--ink3)}.st dd{margin:0}
.skel{height:12px;border-radius:6px;background:linear-gradient(90deg,#E6E4EF,#F4F3F9,#E6E4EF);background-size:200% 100%;animation:sk 1.4s infinite}
@keyframes sk{to{background-position:-200% 0}}
@media (prefers-reduced-motion:reduce){.skel{animation:none}}

/* ---------- mobile ---------- */
.k-mobile{min-height:100vh}
.stage{min-height:100vh;display:flex;align-items:center;justify-content:center;gap:40px;padding:24px}
.phone{position:relative;width:390px;height:844px;flex:none;background:var(--bg);border-radius:40px;box-shadow:0 0 0 10px #1c1b26,0 24px 60px rgba(22,21,31,.25);overflow:hidden;display:flex;flex-direction:column}
.statusbar{display:flex;justify-content:space-between;padding:10px 26px 4px;font-size:13px;font-weight:600;flex:none;background:#fff}
.statusbar .sb-r{color:var(--ink3);font-weight:500}
.m-head{display:flex;align-items:center;justify-content:space-between;height:52px;padding:0 6px;background:#fff;border-bottom:1px solid var(--line);flex:none}
.m-head h1{font-size:18px}
.m-back,.m-back-sp{display:inline-flex;width:44px;height:44px;align-items:center;justify-content:center;color:var(--ink);border-radius:10px}
.m-body{flex:1;overflow:auto;padding:14px 16px;display:flex;flex-direction:column;gap:12px}
.m-body.no-scroll{overflow:hidden}
.m-actions{flex:none;background:#fff;border-top:1px solid var(--line);padding:12px 16px 14px;display:flex;flex-direction:column;gap:8px}
.m-actions .btn{min-height:44px;border-radius:12px;font-size:16px}
.m-actions .btn.main{min-height:56px;font-size:17px}
.m-actions .pair{display:grid;grid-template-columns:1fr 1fr;gap:8px}
.mcard{background:#fff;border:1px solid var(--line);border-radius:16px;padding:12px 14px}
.m-tabs{flex:none;display:grid;grid-template-columns:repeat(3,1fr);background:#fff;border-top:1px solid var(--line);padding:4px 0 10px}
.m-tabs a{display:flex;flex-direction:column;align-items:center;gap:2px;min-height:52px;justify-content:center;font-size:12px;color:var(--ink2);text-decoration:none}
.m-tabs a[aria-current=page]{color:var(--primary);font-weight:700}
.phone-note{width:300px;font-size:14px;color:var(--ink2)}
.pn-k{font-size:13px;font-weight:700;color:var(--ink);margin-bottom:6px}
.flow{list-style:none;margin:0 0 14px;padding:0;display:flex;flex-direction:column;gap:2px}
.flow li{display:flex;gap:8px;align-items:center;font-size:13px;color:var(--ink3)}
.flow li span{display:inline-flex;width:20px;height:20px;border-radius:50%;background:#fff;border:1px solid var(--line);font-size:11px;align-items:center;justify-content:center}
.flow li.on{color:var(--primary);font-weight:700}
.flow li.on span{background:var(--primary);color:#fff;border-color:var(--primary)}
.pn-d p+p{margin-top:8px}
.kbd{display:inline-block;border:1px solid var(--control);border-bottom-width:2px;border-radius:4px;padding:0 6px;font-size:12px;background:#fff;color:var(--ink)}
/* m1 lock */
.lock{flex:1;background:linear-gradient(180deg,#1c1b2b,#2b2552);color:#fff;display:flex;flex-direction:column;align-items:center;padding:40px 14px 20px}
.lock .lk-date{font-size:16px;opacity:1}
.lock .lk-time{font-size:76px;font-weight:600;line-height:1.2}
.push{margin-top:40px;width:100%;display:block;text-decoration:none;color:var(--ink);background:rgba(255,255,255,.96);border-radius:20px;padding:12px 14px;min-height:44px}
.push .ph{display:flex;align-items:center;gap:8px;font-size:13px;color:var(--ink2)}
.push .ph .logo{width:22px;height:22px;border-radius:6px}
.push .ph time{margin-left:auto}
.push b{display:block;margin-top:4px;font-size:15px;line-height:1.5}
.push p{font-size:14px;color:var(--ink2);line-height:1.5}
.lock .hint{margin-top:auto;font-size:13px;color:#fff;display:flex;align-items:center;gap:6px}
/* m2 */
.sev{display:flex;align-items:center;gap:8px;flex-wrap:wrap}
.sev .tag.crit{font-size:14px;padding:2px 10px}
.issue{font-size:15px;font-weight:700;color:var(--crit);line-height:1.5}
.m-title{font-size:21px;line-height:1.45}
.kv{display:flex;align-items:center;gap:8px;flex-wrap:wrap}
.kv b{font-size:17px}
.spread{display:grid;grid-template-columns:repeat(3,1fr);gap:6px;margin-top:8px}
.spread div{background:var(--bg);border-radius:10px;padding:6px 8px;font-size:13px;line-height:1.45}
.spread b{display:block;font-size:14px}
.spread .arrow{display:none}
.confrow{display:flex;align-items:center;gap:8px;font-size:14px;flex-wrap:wrap}
.confrow .meter{flex:1;height:6px;border-radius:3px;background:var(--line);min-width:80px}
.confrow .meter span{display:block;height:100%;border-radius:3px;background:var(--primary)}
.ev{font-size:15px;line-height:1.55}
.ev small{display:flex;align-items:center;gap:4px;color:var(--ink3);margin-top:2px}
.m-status{font-size:13px;color:var(--ink2)}
dialog{border:0;border-radius:20px;padding:20px;width:340px;max-width:calc(100vw - 32px);box-shadow:0 20px 60px rgba(22,21,31,.35)}
dialog::backdrop{background:rgba(22,21,31,.55)}
dialog h2{font-size:18px;margin-bottom:6px}
dialog fieldset{border:0;padding:0;margin:12px 0;display:flex;flex-direction:column;gap:4px}
dialog legend{font-weight:600;font-size:14px;margin-bottom:4px}
.radio{display:flex;align-items:center;gap:10px;min-height:44px;font-size:15px}
.radio input{width:24px;height:24px;accent-color:var(--primary);flex:none}
.dlg-actions{display:grid;grid-template-columns:1fr 1fr;gap:8px;margin-top:8px}
.dlg-actions .btn{min-height:48px}
.toast{position:absolute;left:12px;right:12px;bottom:12px;background:var(--ink);color:#fff;border-radius:14px;padding:10px 12px;display:flex;align-items:center;gap:10px;font-size:14px;z-index:5}
.toast .grow{flex:1}
.toast .btn{min-height:44px}
/* m4–m6 */
.stepper{display:flex;gap:12px;align-items:center;font-size:14px;color:var(--ink2)}
.stepper span{display:inline-flex;align-items:center;gap:6px}
.stepper i{font-style:normal;display:inline-flex;width:26px;height:26px;border-radius:50%;background:#E6E4EF;align-items:center;justify-content:center;font-size:13px;font-weight:700}
.stepper .on{color:var(--primary);font-weight:700}.stepper .on i,.stepper .done i{background:var(--primary);color:#fff}
.rcpt{display:flex;flex-direction:column;gap:8px}
.rcpt label{display:flex;align-items:center;gap:12px;min-height:60px;background:#fff;border:1px solid var(--line);border-radius:14px;padding:8px 14px}
.rcpt label:has(input:checked){border-color:var(--primary)}
.rcpt input{width:24px;height:24px;accent-color:var(--primary);flex:none}
.rcpt b{display:flex;gap:6px;align-items:center;flex-wrap:wrap;font-size:15px}
.rcpt small{display:block;color:var(--ink2);line-height:1.45}
.branchbox{background:#fff;border:1px solid var(--line);border-radius:14px;padding:10px 14px}
.branchbox fieldset{border:0;margin:0;padding:0}
.branchbox legend{font-weight:700;font-size:15px;padding:0}
.branchbox select{min-height:44px;border:1px solid var(--control);border-radius:10px;padding:6px 10px;width:100%;background:#fff}
.sec-k{font-size:13px;color:var(--ink2);font-weight:600}
.review dl{display:grid;grid-template-columns:72px 1fr;gap:8px 12px;margin:0;font-size:14px}
.review dt{color:var(--ink2)}.review dd{margin:0}
.review dd ul{margin:0;padding-left:18px}
.done-ic{display:flex;width:72px;height:72px;border-radius:50%;background:var(--lime);align-items:center;justify-content:center;margin:4px auto 0}
.undo-bar{display:flex;align-items:center;gap:10px;background:var(--ink);color:#fff;border-radius:16px;padding:8px 8px 8px 14px;font-size:14px}
.undo-bar .grow{flex:1}
.undo-bar .btn{min-height:44px}
.ack{list-style:none;margin:0;padding:0}
.ack li{display:flex;align-items:center;gap:10px;padding:10px 0;border-top:1px solid var(--line);font-size:14px}
.ack li:first-child{border-top:0}
.ack .tag{margin-left:auto}
.ack li.stack .who{display:flex;flex-direction:column;align-items:flex-start;gap:4px}
.ack li.stack .tag{margin-left:0}
/* F16 แผงโหมดทดลองใช้งาน (นอกกรอบโทรศัพท์) */
.demo{background:var(--warnTint);border:1px dashed var(--warn);border-radius:14px;padding:12px 14px;margin-bottom:18px;display:flex;flex-direction:column;gap:8px;color:var(--ink)}
.demo h2{font-size:14px;display:flex;align-items:center;gap:6px;flex-wrap:wrap}
.demo h2 small{font-weight:400;color:var(--ink2);font-size:12px}
.sw-row{display:flex;gap:10px;align-items:flex-start;font-size:14px;min-height:32px;cursor:pointer}
.sw-row input{width:24px;height:24px;accent-color:var(--primary);margin-top:2px;flex:none}
.sw-row small{display:block;font-size:12px;color:var(--ink2);line-height:1.45}
.sw-row.sel{align-items:center;justify-content:space-between}
.sw-row select{min-height:32px;border:1px solid var(--control);border-radius:8px;background:#fff;padding:2px 6px}
.demo-out{font-size:12px;color:var(--warn);font-weight:600}
.demo-out:empty{display:none}
/* F05 รูปหลักฐานจำลอง */
.ev{display:grid;grid-template-columns:var(--evimg,88px) 1fr;gap:10px;align-items:start}
.ev-img{width:var(--evimg,88px);height:var(--evimg,88px);border-radius:10px;background:repeating-linear-gradient(135deg,#E6E4EF 0 8px,#F3F2F8 8px 16px);border:1px solid var(--control);display:flex;align-items:center;justify-content:center;text-align:center;font-size:11px;line-height:1.35;color:var(--ink2);padding:4px}
.ev-src{display:flex;align-items:center;justify-content:space-between;gap:6px;grid-column:1/-1;margin-top:-2px}
.ev-src small{margin:0}
.ev-src .btn{min-height:44px;padding:4px 10px;font-size:14px}
/* F08 รับเรื่อง */
.accept{display:flex;align-items:center;justify-content:space-between;gap:8px}
.accept .btn{min-height:44px}
.accept.done .m-status{color:var(--ok);font-weight:600}
/* F09 ผู้รับที่บังคับตามกฎ */
.rcpt .locked{display:flex;align-items:center;gap:12px;min-height:60px;background:#fff;border:1px solid var(--primary);border-radius:14px;padding:8px 14px}
.rcpt .locked .lk{display:inline-flex;width:24px;height:24px;align-items:center;justify-content:center;color:var(--primary);flex:none}
/* F14 F15 */
.msgbox textarea,dialog textarea{width:100%;min-height:96px;border:1px solid var(--control);border-radius:10px;padding:8px 10px;font:inherit;font-size:14px;line-height:1.5}
.msg-acts{display:flex;gap:8px;margin-top:6px}
.msg-acts .btn{min-height:44px;flex:1}
/* F18 ตัวนับเวลา */
.elapsed{display:flex;align-items:baseline;gap:6px;font-size:14px;color:var(--ink)}
.elapsed b{font-size:22px}
/* F21 pop-up ส่งไม่สำเร็จ */
.fail-dlg .cause{display:flex;gap:8px;align-items:center;color:var(--crit);font-weight:700;margin:6px 0}
.fail-dlg .dlg-actions{grid-template-columns:1fr}
.phone.s360{width:360px;height:740px}
/* m2 ต้องไม่เลื่อน: จัดระยะให้กระชับโดยไม่ตัดข้อมูล · .tight ใช้เมื่อย่อรูปแล้วยังล้น (fitM2) */
[data-page=m2] .m-body{gap:9px;padding:12px 16px}
[data-page=m2] .mcard{padding:10px 12px}
[data-page=m2] .spread b{font-size:13px;white-space:nowrap}
[data-page=m2] .spread div{padding:5px 8px}
[data-page=m2] .ev p{font-size:14px;line-height:1.5}
.phone.tight .m-body{gap:6px;padding:8px 14px;line-height:1.45}
.phone.tight .m-title{font-size:18px;line-height:1.4}
.phone.tight .issue{font-size:14px}
.phone.tight .kv b{font-size:15px}
.phone.tight .mcard .sub,.phone.tight .mcard p{font-size:13px}
.phone.tight .spread{margin-top:4px}
.phone.tight .spread div{line-height:1.35}
.phone.tight .m-head{height:46px}
.phone.tight .m-actions{padding:8px 14px 10px;gap:6px}
.phone.tight .m-actions .btn.main{min-height:52px}
/* index & doc pages */
.k-doc .wrap{max-width:1080px;margin:0 auto;padding:40px 28px 80px;display:flex;flex-direction:column;gap:20px}
.k-doc table{background:#fff;border:1px solid var(--line);border-radius:12px;overflow:hidden}
.k-doc td,.k-doc th{font-size:14px}
.swatches{display:grid;grid-template-columns:repeat(auto-fill,minmax(150px,1fr));gap:12px}
.sw{background:#fff;border:1px solid var(--line);border-radius:12px;overflow:hidden;font-size:13px}
.sw i{display:block;height:56px}
.sw div{padding:8px 10px}
.pass{color:var(--ok);font-weight:700}.fail{color:var(--crit);font-weight:700}
.pill-sw{display:inline-block;width:14px;height:14px;border-radius:3px;vertical-align:-2px;border:1px solid var(--line)}
.cases{display:grid;grid-template-columns:1fr 1fr;gap:20px}
.case{display:flex;flex-direction:column;gap:10px;text-decoration:none;color:var(--ink);background:#fff;border:1px solid var(--line);border-radius:20px;padding:24px;min-height:44px}
.case:hover{border-color:var(--primary)}
.case h2{font-size:21px}
.case .go{color:var(--primary);font-weight:700;display:inline-flex;gap:6px;align-items:center;margin-top:auto}
.links{display:flex;flex-wrap:wrap;gap:8px}
.disclaimer{font-size:14px;color:var(--ink2);border-top:1px solid var(--line);padding-top:16px}
.draft{display:inline-block;font-weight:700;color:var(--warn);background:var(--warnTint);border-radius:6px;padding:0 8px;margin-left:8px;font-size:14px;vertical-align:middle}
.prose p+p{margin-top:10px}
.prose ul{margin:8px 0;padding-left:20px}
"""

# ---------------------------------------------------------------- pages
pages = {}


def coverage_items():
    cov = NOTES["ความครอบคลุมที่ต้องแสดงในต้นแบบ"]
    items = []
    for part in cov.split(" · "):
        k, v = part.split(":", 1)
        items.append((k.strip(), v.strip()))
    return items


def filter_bar(compact=False):
    plats = "".join(f"<option>{esc(p)}</option>" for p in S["platforms"])
    brs = "".join(f'<option value="{esc(b)}">สาขา{esc(b)}</option>' for b in BRANCHES5)  # F24 ค่าใช้กรองจริง
    iss = "".join(f"<option>{esc(i)}</option>" for i in ISSUES)
    return f"""
<section class="card" aria-labelledby="flt-h">
  <h2 id="flt-h" class="sr-only">ตัวกรอง</h2>
  <div class="filters">
    <label class="field">ช่วงเวลา<select id="f-range"><option>19–25 ต.ค. 69 (7 วัน)</option><option>20–26 ต.ค. 69</option></select></label>
    <label class="field">แพลตฟอร์ม<select id="f-plat"><option>ทั้งหมด ({S['n_platforms']})</option>{plats}</select></label>
    <label class="field">สาขา<select id="f-branch"><option value="">ทั้งหมด ({len(BRANCHES5)} สาขา)</option>{brs}</select></label>
    <label class="field">ประเด็น<select id="f-issue"><option>ทั้งหมด</option>{iss}</select></label>
    <label class="field">แคมเปญ<select id="f-camp"><option>ไม่กรอง</option><option>{CAMPAIGN}</option></select></label>
    <label class="check"><input type="checkbox" checked> ซ่อนข้อความไม่เกี่ยวข้อง (2)</label>
    <button class="btn" type="button" style="margin-left:auto">{icon('star', size=18)}บันทึกชุดตัวกรอง</button>
  </div>
  <div class="saved" role="group" aria-label="ชุดตัวกรองที่บันทึกไว้">
    <span>ตัวกรองที่บันทึกไว้</span>
    <button class="chip" type="button" aria-pressed="false" data-saved="camp">{icon('star', size=14)}{CAMPAIGN} · 20–26 ต.ค. 69</button>
    <button class="chip" type="button" aria-pressed="false" data-branch="รังสิต">{icon('star', size=14)}สาขารังสิต · ความปลอดภัยอาหาร</button>
    <button class="chip" type="button" aria-pressed="false">{icon('star', size=14)}รีวิวเดลิเวอรี (Grab, LINE MAN)</button>
    <span class="muted" id="saved-msg" role="status"></span>
  </div>
  <p class="foot-note" style="margin-top:8px" id="flt-note">ตัวกรองสาขาเปลี่ยนตัวเลขและกราฟจริงจาก CSV · ตัวกรองอื่นยังเป็นภาพนิ่ง</p>
</section>"""


def kpi_html(st):
    """F24 ตัวเลขหลักของภาพรวมหรือรายสาขา"""
    scope = f"{st['n_platforms']} แพลตฟอร์ม" if st["branch"] is None else f"{branch_name(st['branch'])} · {st['n_platforms']} แพลตฟอร์ม"
    first_d, last_d = st["camp_days"][0], st["camp_days"][-1]
    camp_range = f"{DAY_TH[first_d].split(' ')[1]}–{DAY_TH[last_d].split(' ')[1]} ต.ค. 69"
    camp_pct = f" ({st['camp_neg'] / st['camp_total'] * 100:.1f}%)" if st["camp_total"] else ""
    return f"""
  <div class="kpi"><p class="k">การกล่าวถึงทั้งหมด</p><p class="v num">{fmt(st['total'])}</p><p class="d">19–25 ต.ค. 69 · {esc(scope)}</p></div>
  <div class="kpi"><p class="k">เชิงลบ</p><p class="v num">{fmt(st['neg'])}</p><p class="d">{st['neg_pct']:.1f}% ของทั้งหมด</p></div>
  <div class="kpi hot"><p class="k">ชั่วโมงที่พุ่งสูงสุด</p><p class="v num">{st['peak']} ครั้ง</p><p class="d">{esc(st['peak_th'])} · ลบ {st['peak_neg']}</p></div>
  <div class="kpi"><p class="k">แคมเปญ {CAMPAIGN}</p><p class="v num">{st['camp_total']}</p><p class="d">{camp_range} · ลบ {st['camp_neg']}{camp_pct} · 26 ต.ค. ยังไม่ครบวัน</p></div>"""


def build_w1():
    sp = S["camp_sent"]
    # F30 ทุกแพลตฟอร์มใช้ข้อความจากชีต "คำอธิบายและการแก้ไข" · ไม่มีข้อความ = "ยังไม่ทราบข้อจำกัด"
    cov = coverage_items()
    cov_rows = []
    for plat in S["platforms"]:
        txt = next((v for k, v in cov if plat in [x.strip() for x in k.split("/")]), None)
        cov_rows.append(f'<li><b>{esc(plat)}</b>' + (f'<span>{esc(txt)}</span>' if txt else '<span class="muted">ยังไม่ทราบข้อจำกัด</span>') + '</li>')
    for k, v in cov:  # ช่องทางที่ชีตระบุว่านอกขอบเขต (ไม่ใช่แพลตฟอร์มใน CSV)
        if not any(pl in [x.strip() for x in k.split("/")] for pl in S["platforms"]):
            cov_rows.append(f'<li><b>{esc(k)}</b><span>{esc(v)}</span></li>')
    cov_html = "".join(cov_rows)
    first_d, last_d = S["camp_days"][0], S["camp_days"][-1]
    camp_range = f"{DAY_TH[first_d].split(' ')[1]}–{DAY_TH[last_d].split(' ')[1]} ต.ค. 69"
    c04, c14 = POSTS["C04"], POSTS["C14"]
    unk = next(ng for b, ng, _ in S["branches"] if b == "ไม่ระบุ")  # F29 เชิงลบที่ไม่ระบุสาขา = 310 (CSV)
    # F24 แม่แบบของแต่ละสาขา ให้ JavaScript สลับเมื่อเลือกตัวกรองสาขา (ตัวเลขคำนวณจาก CSV ใน build.py)
    variants = [("", ST_ALL)] + [(b, ST_BRANCH[b]) for b in BRANCHES5]
    tpls = "".join(
        f'<template data-branch="{esc(k)}" data-part="kpi">{kpi_html(st)}</template>'
        f'<template data-branch="{esc(k)}" data-part="hourly">{svg_hourly(st)}</template>'
        f'<template data-branch="{esc(k)}" data-part="plat">{svg_platform(st)}</template>'
        f'<template data-branch="{esc(k)}" data-part="plattbl">{platform_table(st)}</template>'
        f'<template data-branch="{esc(k)}" data-part="camp">{svg_campaign(st)}</template>'
        for k, st in variants)
    k10n, k14n = len(case_posts("K10")), len(case_posts("K14"))
    content = f"""
<div class="page-head"><h1>Dashboard</h1></div>
<!-- F23 ซ่อนแท็บที่ยังไม่ทำ (ประเด็นร้องเรียน / รายสาขา / แคมเปญ) -->
<!-- F25 แถบเคสที่ยังไม่ปิดอยู่เหนือตัวกรอง · F04 ใช้ชื่อระดับชุดเดียว -->
<section class="alertbar" aria-labelledby="open-cases-h">{icon('alert', label='คำเตือน', size=24)}<b id="open-cases-h">เคสที่ยังไม่ปิด 2 เคส</b>
  <ul>
    <li>{sev_tag(case_level('K10'), True)} <a href="w2.html?case=K10">K10</a> ความปลอดภัยอาหาร · สาขารังสิต · {k10n} โพสต์ · ต้นเรื่อง C18 · ผู้โพสต์ต้นเรื่องถามความคืบหน้า (C28 {esc(thai_dt(POSTS['C28']['วันเวลา']))})</li>
    <li>{sev_tag(case_level('K14'))} <a href="w2.html?case=K14">K14</a> การชำระเงิน · สาขาลาดพร้าว · {k14n} โพสต์ · เสี่ยงเปิดเผยข้อมูลส่วนบุคคล (C29)</li>
  </ul>
  <a class="more" href="w-cases.html">ดูทั้งหมดใน {NAV_LABEL['cases']} {icon('right', size=16)}</a>
</section>
{filter_bar()}
<p class="sr-only" id="w1-live" role="status" aria-live="polite"></p>
<section class="kpis" id="kpis" aria-label="ตัวเลขหลัก 19–25 ต.ค. 69">{kpi_html(ST_ALL)}
</section>
<section class="card" aria-labelledby="h-hourly">
  <h2 id="h-hourly">การกล่าวถึงรายชั่วโมง 7 วัน</h2>
  <p class="sub">19–25 ต.ค. 69 · <span class="scope">ทุกสาขา</span> · พื้นหลังเทาอ่อน = นอกเวลาทำการ (20:00–07:00 น.)</p>
  <div class="legend"><span><i class="line" style="background:{T['primary']}"></i>ทั้งหมด</span><span><i class="line" style="background:{T['neg']}"></i>เชิงลบ</span><span><i style="background:{T['band']}"></i>นอกเวลาทำการ</span><span><i class="line" style="background:{T['ink']}"></i>เส้นประ = รีวิวแรก C18</span></div>
  <div id="slot-hourly">{svg_hourly()}</div>
  <div class="chart-foot">
    <span>เปิด Social Feed ของช่วงนี้ โดยส่งตัวกรองปัจจุบันไปด้วย</span>
    <a class="btn soft" href="w2.html">{icon('right', size=18)}ดูโพสต์ ศ. 23 ต.ค. 21:00 – ส. 24 ต.ค. 03:00 น.</a>
  </div>
</section>
<div class="grid2">
  <section class="card" aria-labelledby="h-plat">
    <h2 id="h-plat">ความรู้สึกแยกตามแพลตฟอร์ม</h2>
    <p class="sub">สัดส่วนต่อแพลตฟอร์ม · ตัวเลขขวา = จำนวนการกล่าวถึง · ชี้หรือกดแท็บที่แท่งเพื่อดูตัวเลข</p>
    <div class="legend">{''.join(f'<span><i style="background:{SENT_COLOR[s]};{"box-shadow:inset 0 0 0 1px " + NEU_STROKE if s == "เป็นกลาง" else ""}"></i>{s}</span>' for s in SENT_ORDER)}</div>
    <p class="foot-note">ผสม = มีทั้งชมและติ</p><!-- F28 -->
    <div id="slot-plat">{svg_platform()}</div>
    <div class="card-acts"><button class="btn sm" type="button" id="plat-tbl-btn" aria-expanded="false" aria-controls="slot-plattbl">{icon('list', size=16)}ดูตารางข้อมูล</button></div><!-- F28 -->
    <div id="slot-plattbl" hidden>{platform_table()}</div>
  </section>
  <section class="card" aria-labelledby="h-br">
    <h2 id="h-br">เชิงลบแยกตามสาขา</h2>
    <p class="sub">ตัวเลขหนา = เชิงลบ · ตัวเลขจาง = ทั้งหมดของสาขา</p>
    <div class="legend"><span><i style="background:{T['neg']}"></i>เชิงลบ</span><span><i style="background:{T['band']}"></i>ทั้งหมด</span></div>
    {svg_branch()}
    <p class="explain">"ไม่ระบุสาขา" คือข้อความที่ระบบหาสาขาไม่ได้ ส่วนสาขาที่ระบบเดาจะแสดงความมั่นใจใน Social Feed และต้องให้คนยืนยันก่อนส่งต่อ</p>
    <div class="card-acts"><a class="btn soft" href="w2.html?branch=ไม่ระบุ">{icon('right', size=18)}ระบุสาขา {fmt(unk)} รายการ</a><span class="foot-note">เชิงลบที่ไม่ระบุสาขา</span></div><!-- F29 -->
  </section>
</div>
<div class="grid2">
  <section class="card" aria-labelledby="h-cov">
    <h2 id="h-cov">ความครอบคลุมของข้อมูล</h2>
    <p class="sub">ตัวเลขบนหน้านี้นับเฉพาะข้อมูลที่ระบบเข้าถึงได้ (ข้อความจากชีต "คำอธิบายและการแก้ไข")</p>
    <ul class="cov">{cov_html}</ul>
  </section>
  <section class="card" aria-labelledby="h-camp">
    <h2 id="h-camp">แคมเปญ: {CAMPAIGN}</h2>
    <p class="sub">ชุดตัวกรองที่บันทึกไว้ · 20–26 ต.ค. 69 · กราฟแสดง {camp_range} เพราะ 26 ต.ค. ยังไม่ครบวัน · <span class="scope">ทุกสาขา</span></p>
    <div id="slot-camp">{svg_campaign()}</div>
    <p class="sub" style="margin-top:6px">ความรู้สึก (ทุกสาขา): บวก {sp['บวก']} · ลบ {sp['ลบ']} · เป็นกลาง {sp['เป็นกลาง']} · ผสม {sp['ผสม']}</p>
    <h3 style="margin-top:14px">ตัวอย่างข้อความในแคมเปญ</h3>
    <ul class="sim">
      <li><span class="postid">C04</span><span>"{render_text(c04['ข้อความ'])}"<span class="meta muted">{esc(c04['แพลตฟอร์ม'])} · {esc(thai_dt(c04['วันเวลา']))}</span></span>{sent_badge(c04['ความรู้สึก (ระบบ)'])}</li>
      <li><span class="postid">C14</span><span>"{render_text(c14['ข้อความ'])}"<span class="meta muted">{esc(c14['แพลตฟอร์ม'])} · {esc(thai_dt(c14['วันเวลา']))}</span></span>{sent_badge(c14['ความรู้สึก (ระบบ)'])}</li>
    </ul>
  </section>
</div>
<p class="foot-note">ตัวเลขทั้งหมดคำนวณจาก IS757_ตัวเลขรายชั่วโมงสมมติ.csv ช่วง 19–25 ต.ค. 69 (ไม่รวมยอด 26 ต.ค.) · <a class="solo-link" href="data.html">ที่มาของตัวเลข</a> · ชบา คอฟฟี่ เป็นแบรนด์สมมติ</p>
<div hidden id="branch-tpls">{tpls}</div>
<div class="tip" id="tip" hidden></div>
"""
    return page("Dashboard", web_shell("w1", [("Dashboard", "w1.html"), ("ภาพรวม", None)], content, "w1"), "w1", "web")


def post_row(p):
    pid = p["ID"]
    owner, status = owner_status(p)
    br, br_note = branch_parts(p)
    sens = is_sensitive(p)
    if sens:
        msg = '<span class="tag warn">ข้อความมีข้อมูลสุขภาพ (PDPA) · ซ่อนไว้ แสดงตามสิทธิ์</span>'
    else:
        msg = render_text(p["ข้อความ"], keywords(p))
        um = unmatched_kws(p)
        if um:
            msg += '<span class="kwchips">' + "".join(f'<span class="kwchip">คำหลักฐาน: {esc(k)}</span>' for k in um) + "</span>"
    extra = []
    if p["เคส (เสนอ)"]:
        extra.append(f'<span class="tag plain">เคส {esc(p["เคส (เสนอ)"])}</span>')
    if pid == "C29":
        extra.append('<span class="tag crit">เสี่ยงเปิดเผยข้อมูลส่วนบุคคล</span>')
    idcell = (f'<a class="postid" href="w3.html">{pid}</a>' if pid == "C09" else f'<span class="postid">{pid}</span>')
    br_html = esc(br)
    if br_note:
        br_html += f'<span class="meta" style="color:var(--warn)">{esc(br_note)} · รอยืนยัน</span>'
    # data-case / data-branch ใช้กับ ?case= และ ?branch= จาก Dashboard
    return f"""<tr data-id="{pid}" data-review="{'1' if needs_review(p) else '0'}" data-case="{esc(p['เคส (เสนอ)'])}" data-branch="{esc(br)}">
<td><input type="checkbox" class="rowsel" aria-label="เลือก {pid}" style="width:24px;height:24px;accent-color:var(--primary)"></td>
<td class="idcell">{idcell}<span class="meta">{esc(thai_dt(p['วันเวลา']))}</span></td>
<td class="platcell">{esc(p['แพลตฟอร์ม'])}<span class="meta">{esc(engagement(p))}</span><span class="meta">{icon('eyeoff', label='ชื่อผู้โพสต์ถูกปกปิด', size=13)} {esc(p['ผู้โพสต์ (ปกปิด)'])}</span></td>
<td class="msg">{msg}{('<span class="kwchips">' + ''.join(extra) + '</span>') if extra else ''}</td>
<td class="brcell">{br_html}</td>
<td class="sentcell">{sent_badge(p['ความรู้สึก (ระบบ)'])}</td>
<td class="confcell">{conf_meter(p)}</td>
<td class="nowrap urgcell">{sev_tag(p['ความเร่งด่วน (ระบบ)'])}</td>
<td class="ownercell">{esc(owner)}</td>
<td class="statuscell">{status_tag(status)}</td>
</tr>"""


def build_w2():
    inc = sorted([p for p in POSTS.values() if not excluded(p)], key=sort_key, reverse=True)
    n_ex = sum(1 for p in POSTS.values() if excluded(p))
    rows = "".join(post_row(p) for p in inc)
    owners = ["CS ส่วนกลาง", "แอดมินเพจ", "ผู้ดูแลข้อมูลส่วนบุคคล", "ยังไม่มอบหมาย"]
    statuses = list(STATUS_TAG)
    content = f"""
<div class="page-head"><h1>Social Feed</h1></div>
<div class="tabs" role="presentation">
  <a href="w2.html" id="tab-all" aria-current="page">ทั้งหมด ({len(inc)})</a>
  <a href="w2.html#review" id="tab-review">รอตรวจสอบ · ความมั่นใจต่ำ ({len(REVIEW_IDS)})</a>
  <a href="w2-excluded.html">คัดออก ({n_ex})</a>
</div>
<section class="card" aria-label="ตัวกรองรายการ">
  <div class="filters">
    <label class="field">ความรู้สึก (ระบบ)<select><option>ทั้งหมด</option>{''.join(f'<option>{s}</option>' for s in SENT_ORDER)}</select></label>
    <label class="field">ความเร่งด่วน<select><option>ทั้งหมด</option>{''.join(f'<option>{l}</option>' for l in SEV_LEVELS[1:])}</select></label>
    <label class="field">สถานะ<select><option>ทั้งหมด</option>{''.join(f'<option>{x}</option>' for x in statuses)}</select></label>
    <label class="field">ผู้รับผิดชอบ<select><option>ทั้งหมด</option>{''.join(f'<option>{o}</option>' for o in owners)}</select></label>
    <label class="check"><input type="checkbox" id="only-low"> เฉพาะความมั่นใจต่ำกว่า 75%</label>
  </div>
  <div class="saved">{icon('filter', size=16)}<span>ตัวกรองที่ใช้อยู่</span>
    <div class="fchips" id="fchips"><span class="chip on">19–25 ต.ค. 69</span></div><!-- F31 ช่วงวันตรงกับ w1 -->
    <span class="muted">(ตัวกรองแบบเลื่อนลงเป็นภาพนิ่ง)</span></div>
</section>
<p class="legend-line">{icon('eyeoff', size=16)}<span>ชื่อผู้โพสต์ถูกปกปิด แสดงเฉพาะประเภทบัญชีและเลขท้าย 2 หลัก</span><span><mark>คำไฮไลต์</mark> = คำหลักฐานที่ระบบใช้ตัดสิน</span><span>ความมั่นใจต่ำกว่า 75% ขึ้นป้าย <span class="tag warn">รอตรวจสอบ</span></span></p>
<p class="muted" id="list-status" role="status">แสดง {len(inc)} รายการ</p>
<!-- F33 แถบคำสั่งกลุ่ม ขึ้นเมื่อเลือกอย่างน้อย 1 แถว -->
<div class="bulk" id="bulk" hidden role="region" aria-label="คำสั่งกับรายการที่เลือก">
  <b id="bulk-n">เลือก 0 รายการ</b>
  <label class="field">มอบหมาย<select id="bulk-owner"><option value="">—</option>{''.join(f'<option>{o}</option>' for o in owners[:3])}</select></label>
  <label class="field">เปลี่ยนสถานะ<select id="bulk-status"><option value="">—</option>{''.join(f'<option>{x}</option>' for x in statuses)}</select></label>
  <button class="btn primary sm" type="button" id="bulk-apply">นำไปใช้</button>
  <button class="btn sm" type="button" id="bulk-edit">แก้ผลพร้อมกัน</button>
  <button class="btn ghost sm" type="button" id="bulk-clear">ยกเลิกการเลือก</button>
  <span class="muted" id="bulk-msg" role="status"></span>
</div>
<div class="tbl-wrap">
<table>
<caption class="sr-only">Social Feed เรียงจากใหม่ไปเก่า</caption>
<thead><tr><th scope="col"><label class="check" style="min-height:0;font-size:13px;color:var(--ink2);gap:6px"><input type="checkbox" id="sel-all" aria-label="เลือกทุกแถวที่แสดง">เลือก</label></th><th scope="col">โพสต์ · เวลา</th><th scope="col">แพลตฟอร์ม · ผู้โพสต์</th><th scope="col">ข้อความ · คำหลักฐานที่ระบบใช้</th><th scope="col">สาขา</th><th scope="col">ความรู้สึก (ระบบ)</th><th scope="col">ความมั่นใจ</th><th scope="col">ความเร่งด่วน (ระบบ)</th><th scope="col">ผู้รับผิดชอบ</th><th scope="col">สถานะ</th></tr></thead>
<tbody>{rows}</tbody>
</table>
</div>
<p class="foot-note">คลิก C09 เพื่อเปิดหน้ารายละเอียดและแก้ผล · ผู้รับผิดชอบและสถานะเป็นค่าตัวอย่างในต้นแบบ {PH_CONFIRM} · รายการนี้มีเฉพาะ {len(POSTS)} โพสต์ตัวอย่างจาก xlsx ส่วนตัวเลขบน Dashboard นับจาก CSV</p>
"""
    return page("Social Feed", web_shell("w2", [("Dashboard", "w1.html"), ("Social Feed", None)], content, "w2"), "w2", "web")


def build_w2_excluded():
    inc_n = sum(1 for p in POSTS.values() if not excluded(p))
    ex = sorted([p for p in POSTS.values() if excluded(p)], key=sort_key)
    items = []
    for p in ex:
        pid = p["ID"]
        items.append(f"""<div class="ex-row" data-id="{pid}">
  <div><p><b>{pid}</b> · {esc(p['แพลตฟอร์ม'])} · {esc(thai_dt(p['วันเวลา']))} · {icon('eyeoff', label='ชื่อผู้โพสต์ถูกปกปิด', size=14)} {esc(p['ผู้โพสต์ (ปกปิด)'])}</p>
  <p class="q">"{render_text(p['ข้อความ'], keywords(p))}"</p>
  <p class="sub">ระบบจัดเป็น {sent_badge(p['ความรู้สึก (ระบบ)'])} ความมั่นใจ {conf_pct(p)}% · คนทำเครื่องหมาย "ไม่เกี่ยวข้อง"</p></div>
  <div><p class="muted">เหตุผลที่คัดออก</p><p><b>{esc(p['ประเภทความผิดพลาด'])}</b></p></div>
  <div class="row" style="display:flex;flex-direction:column;gap:8px;align-items:stretch">
    <button class="btn" type="button" data-restore="{pid}">{icon('undo', size=18)}<span>คืนสู่รายการ</span></button>
    <p class="muted ex-status" role="status" style="font-size:13px"></p>
  </div>
</div>""")
    content = f"""
<div class="page-head"><h1>Social Feed</h1></div>
<div class="tabs" role="presentation">
  <a href="w2.html">ทั้งหมด ({inc_n})</a>
  <a href="w2.html#review">รอตรวจสอบ · ความมั่นใจต่ำ ({len(REVIEW_IDS)})</a>
  <a href="w2-excluded.html" aria-current="page">คัดออก ({len(ex)})</a>
</div>
<section class="card" aria-labelledby="ex-h">
  <h2 id="ex-h">คัดออก · ข้อความที่ไม่เกี่ยวข้องกับแบรนด์ ({len(ex)})</h2>
  <p class="sub">ไม่นับใน Dashboard และการแจ้งเตือน แต่ยังเก็บไว้ให้ตรวจย้อนหลังและคืนสู่รายการได้</p>
  <div style="margin-top:10px">{''.join(items)}</div>
  <div class="notice warn" style="margin-top:14px">{icon('alert', label='ข้อสังเกต', size=18)}<p>ทั้งสองข้อความระบบให้ความมั่นใจสูง ({conf_pct(POSTS['C12'])}% และ {conf_pct(POSTS['C27'])}%) แต่ผิด ความมั่นใจสูงจึงไม่ได้แปลว่าถูกเสมอ ตัวกรองนี้ต้องให้คนทำเครื่องหมายได้ ไม่ใช่พึ่งค่าความมั่นใจอย่างเดียว</p></div>
</section>
"""
    return page("Social Feed · คัดออก", web_shell("w2x", [("Dashboard", "w1.html"), ("Social Feed", "w2.html"), ("คัดออก", None)], content, "w2"), "w2x", "web")


def radio_group(name, options, checked=None):
    out = []
    for o in options:
        c = " checked" if o == checked else ""
        out.append(f'<label><input type="radio" name="{name}" value="{esc(o)}"{c}><span>{esc(o)}</span></label>')
    return f'<div class="seg-opts">{"".join(out)}</div>'


def build_w3():
    p = C09
    sims = []
    for sid in ("C31", "C32"):
        s = POSTS[sid]
        sims.append(f'<li data-sim="{sid}" data-plat="{esc(s["แพลตฟอร์ม"])}" data-orig="{esc(s["ความรู้สึก (ระบบ)"])}"><b>{sid}</b><span>"{render_text(s["ข้อความ"], keywords(s))}"</span><span class="muted simres">{esc(s["แพลตฟอร์ม"])} · ระบบ {sent_badge(s["ความรู้สึก (ระบบ)"])} {conf_pct(s)}%</span></li>')
    c15 = POSTS["C15"]
    kw = keywords(p)[0]
    content = f"""
<div class="page-head"><h1>C09 · รายละเอียดโพสต์</h1></div>
<!-- F37 ส่วนตั้งค่าของโพสต์ มีหัวข้อและหน้าตาต่างจากแถบตัวกรอง -->
<section class="card setup" aria-labelledby="setup-h">
  <h2 id="setup-h">{icon('sliders', size=18)} ตั้งค่าโพสต์นี้</h2>
  <label class="opt">สถานะ<select id="set-status">{''.join(f'<option{" selected" if x == "ใหม่" else ""}>{x}</option>' for x in STATUS_TAG)}</select></label>
  <label class="opt">ผู้รับผิดชอบ<select id="set-owner"><option>แอดมินเพจ</option><option>CS ส่วนกลาง</option></select></label>
  <label class="opt">เคส<select id="set-case"><option>ไม่ผูกเคส</option></select></label>
  <span class="muted" id="setup-msg" role="status">เปลี่ยนแล้วบันทึกทันที</span>
</section>
<div class="detail">
  <div class="detail-col">
    <section class="card" aria-labelledby="p-h">
      <h2 id="p-h" class="sr-only">ข้อความต้นทาง</h2>
      <div style="display:flex;justify-content:space-between;gap:12px;align-items:flex-start">
        <p class="sub">{icon('eyeoff', label='ชื่อผู้โพสต์ถูกปกปิด', size=16)} {esc(p['ผู้โพสต์ (ปกปิด)'])} · {esc(p['แพลตฟอร์ม'])} · {esc(thai_dt(p['วันเวลา']))} · {esc(engagement(p))}</p>
        <button class="btn sm" type="button" aria-label="เปิดโพสต์ต้นทาง (ภาพนิ่ง)">{icon('external', size=16)}เปิดโพสต์ต้นทาง</button>
      </div>
      <p class="quote">"{render_text(p['ข้อความ'], kw and keywords(p))}"</p>
      <p class="sub">สาขา: <b>{esc(p['สาขา'])}</b> (ระบุในโพสต์) · ประเด็น: <b>{esc(p['ประเด็น'])}</b></p>
    </section>
    <section class="card" aria-labelledby="r-h">
      <h2 id="r-h">ผลจากระบบ</h2>
      <p class="sub">อ่านก่อนตัดสินใจแก้ · คำไฮไลต์คือคำที่ระบบใช้เป็นหลักฐาน</p>
      <dl class="res-grid">
        <div><dt>ความรู้สึก</dt><dd id="cur-sent">{sent_badge(p['ความรู้สึก (ระบบ)'])}</dd></div>
        <div><dt>ความมั่นใจ</dt><dd id="cur-conf">{conf_meter(p)}</dd></div>
        <div><dt>คำหลักฐาน</dt><dd><mark>{esc(kw)}</mark></dd></div>
        <div><dt>ความเร่งด่วน</dt><dd id="cur-urg">{sev_tag(p['ความเร่งด่วน (ระบบ)'])}</dd></div>
      </dl>
      <p class="explain">ระบบให้น้ำหนักกับคำว่า "{esc(kw)}" แต่คำว่า "ชงอร่อยจน..." และ "555" ในประโยคเดียวกันไม่ได้ถูกใช้เป็นหลักฐาน</p>
      <p class="sub" style="margin-top:10px">การเรียนรู้ของระบบ: <span id="learn-status">—</span></p>
    </section>
    <section class="card" aria-labelledby="s-h">
      <h2 id="s-h">ข้อความที่ระบบอ่านผิดแบบเดียวกัน</h2>
      <p class="sub">โพสต์ขบขันที่ระบบจัดเป็นลบ (กรณีที่ 2)</p>
      <ul class="sim">{''.join(sims)}</ul>
      <!-- F38 ใช้ผลเดียวกันกับ C31, C32 มีกล่องยืนยันและเลิกทำได้ -->
      <div class="card-acts"><button class="btn" type="button" id="apply-sim" disabled aria-describedby="apply-sim-hint">ใช้ผลเดียวกันกับ 2 รายการนี้</button>
        <span class="hint muted" id="apply-sim-hint" style="font-size:13px">บันทึกผลของ C09 ก่อน แล้วใช้ผลเดียวกันได้</span></div>
      <div id="sim-msg" role="status" aria-live="polite"></div>
      <p class="foot-note" style="margin-top:8px">เทียบ: C15 "{render_text(c15['ข้อความ'])}" ระบบอ่านสแลงถูกเป็น{esc(c15['ความรู้สึก (ระบบ)'])}</p>
    </section>
    <section class="card" aria-labelledby="h-h">
      <h2 id="h-h">ประวัติ</h2>
      <ul class="hist" id="history">
        <li><time>พฤ. 22 ต.ค. 69 · 08:21 น.</time><span>ระบบจัดเป็น {esc(p['ความรู้สึก (ระบบ)'])} ({conf_pct(p)}%) · ความเร่งด่วน{esc(p['ความเร่งด่วน (ระบบ)'])}</span></li>
        <li><time>จ. 26 ต.ค. 69 · 13:58 น.</time><span>มอบหมายให้แอดมินเพจ · สถานะ ใหม่</span></li>
      </ul>
    </section>
  </div>
  <form class="panel" id="edit" aria-labelledby="e-h" novalidate>
    <h2 id="e-h">แก้ผลวิเคราะห์</h2>
    <div class="edit-fields" id="edit-fields" style="display:flex;flex-direction:column;gap:12px">
    <div class="diff" aria-live="polite" id="diff">
      <span>ความรู้สึก: <span class="o">{esc(p['ความรู้สึก (ระบบ)'])}</span> → <b class="n">{esc(p['ความรู้สึก (ระบบ)'])}</b></span>
      <span>ความเร่งด่วน: <span class="o">{esc(p['ความเร่งด่วน (ระบบ)'])}</span> → <b class="n">{esc(p['ความเร่งด่วน (ระบบ)'])}</b></span>
    </div>
    <fieldset><legend>1. ความรู้สึกที่ถูกต้อง</legend>{radio_group('sent', SENT_ORDER[::-1] + ['ไม่เกี่ยวข้อง'], p['ความรู้สึก (ระบบ)'])}</fieldset>
    <fieldset><legend>2. เหตุผลที่ระบบอ่านผิด <small>(เลือก 1 · จำเป็น)</small></legend>{radio_group('reason', ['ประชด', 'ขบขัน', 'ความเห็นผสม', 'อื่นๆ'])}</fieldset>
    <fieldset><legend>3. ความเร่งด่วน <small>(แก้แยกจากความรู้สึกได้)</small></legend>{radio_group('urg', ['ต่ำ', 'กลาง', 'สูง'], p['ความเร่งด่วน (ระบบ)'])}</fieldset>
    <label class="check" style="align-items:flex-start"><input type="checkbox" id="learn" checked style="margin-top:3px"><span><b>ส่งให้ระบบเรียนรู้</b><br><span class="hint">ส่งเฉพาะข้อความและป้ายที่แก้ ไม่ส่งชื่อหรือรหัสผู้โพสต์</span></span></label>
    <div id="save-msg" class="live" role="status" aria-live="polite"></div>
    <div class="btnrow"><button class="btn primary" type="submit" id="save" disabled>บันทึก</button><button class="btn" type="button" id="cancel">ยกเลิก</button></div><!-- F42 แถวเดียว -->
    <p class="hint" id="save-hint">เลือกเหตุผลก่อน ปุ่มบันทึกจึงกดได้ · ยกเลิก = คืนค่าเดิม</p>
    </div>
    <!-- F40 F41 หลังบันทึก: ซ่อนฟอร์ม แสดงสรุป + เลิกทำ -->
    <div class="saved-sum" id="saved-sum" hidden tabindex="-1"></div>
  </form>
</div>
<p class="demo-hint">สาธิตสำหรับผู้สังเกตการณ์: กด <span class="kbd">E</span> เพื่อให้การบันทึกครั้งถัดไปไม่สำเร็จ <span class="demo-flag" id="demo-flag" role="status"></span></p>
<dialog id="sim-dlg" aria-labelledby="sim-dlg-h">
  <form method="dialog">
    <h2 id="sim-dlg-h">ใช้ผลเดียวกันกับ C31 และ C32?</h2>
    <p class="sub" style="font-size:14px" id="sim-dlg-t"></p>
    <p class="sub" style="font-size:14px;margin-top:6px">เลิกทำได้หลังยืนยัน</p>
    <div class="dlg-actions"><button class="btn" value="cancel" type="submit">ยกเลิก</button><button class="btn primary" value="ok" type="submit">ใช้ผลกับ 2 รายการ</button></div>
  </form>
</dialog>
"""
    return page("C09 · รายละเอียดโพสต์", web_shell("w3", [("Dashboard", "w1.html"), ("Social Feed", "w2.html"), ("C09", None)], content, "w2"), "w3", "web")


def build_w_cases():
    """F22 หน้า Alerts & Cases (ตาม sitemap) · แสดง K10 K14 จาก xlsx · คลิกเพื่อเปิดรายละเอียด"""
    info = {"K10": ("ความปลอดภัยอาหาร", "สาขารังสิต", "C18"), "K14": ("การชำระเงิน", "สาขาลาดพร้าว", "C29")}
    rows = []
    for cid in ("K10", "K14"):
        ps = case_posts(cid)
        issue, br, src = info[cid]
        lv = case_level(cid)
        owner, status = owner_status(POSTS[src])
        plats = []
        for x in ps:
            if x["แพลตฟอร์ม"].split(" (")[0] not in plats:
                plats.append(x["แพลตฟอร์ม"].split(" (")[0])
        items = []
        for x in ps:
            txt = '<span class="tag warn">ข้อความมีข้อมูลสุขภาพ (PDPA) · ซ่อนไว้</span>' if is_sensitive(x) else f'"{render_text(x["ข้อความ"], keywords(x))}"'
            items.append(f'<li><span class="postid">{x["ID"]}</span><span>{txt}<span class="meta">{esc(x["แพลตฟอร์ม"])} · {esc(thai_dt(x["วันเวลา"]))} · {icon("eyeoff", label="ชื่อผู้โพสต์ถูกปกปิด", size=13)} {esc(x["ผู้โพสต์ (ปกปิด)"])}</span></span>{sev_tag(x["ความเร่งด่วน (ระบบ)"])}</li>')
        extra = (f'<dt>ข้อควรระวัง</dt><dd><span class="tag crit">เสี่ยงเปิดเผยข้อมูลส่วนบุคคล</span> ระบบไม่แสดงข้อมูลส่วนบุคคลที่อาจตามมา</dd>' if cid == "K14" else
                 f'<dt>การลาม</dt><dd>{" → ".join(esc(p["แพลตฟอร์ม"]) for p in SPREAD)} ใน {SPREAD_MIN} นาที · นอกเวลาทำการ</dd>')
        rows.append(f"""<div class="case-row{' crit' if lv == SEV_CRIT else ''}" id="case-{cid}">
  <button class="case-btn" type="button" aria-expanded="false" aria-controls="cd-{cid}" data-case="{cid}">
    <span class="cid">{cid}</span>{sev_tag(lv, lv == SEV_CRIT)}
    <span><b>{esc(issue)} · {esc(br)}</b><span class="meta">{len(ps)} โพสต์ · {esc(' · '.join(plats))} · โพสต์แรก {esc(thai_dt(ps[0]['วันเวลา']))}</span></span>
    <span style="display:flex;gap:10px;align-items:center">{status_tag(status)}<span class="chev">{icon('right', size=18)}</span></span>
  </button>
  <div class="case-detail" id="cd-{cid}" hidden>
    <dl><dt>ระดับเคส</dt><dd>{sev_tag(lv)}</dd><dt>ต้นเรื่อง</dt><dd>{src}</dd><dt>ผู้รับผิดชอบ</dt><dd>{esc(owner)} {PH_CONFIRM}</dd>{extra}</dl>
    <ul class="sim">{''.join(items)}</ul>
    <div class="card-acts"><a class="btn soft" href="w2.html?case={cid}">{icon('list', size=18)}เปิดใน {NAV_LABEL['posts']}</a>{'<a class="btn" href="m2.html">ดูขั้นยกระดับบนมือถือ (กรณีที่ 1)</a>' if cid == 'K10' else ''}</div>
  </div>
</div>""")
    legend = " · ".join(f"{sev_tag(l)}" for l in SEV_LEVELS)
    content = f"""
<div class="page-head"><h1>{NAV_LABEL['cases']}</h1></div>
<p class="sub">เคสที่ยังไม่ปิด 2 เคส · ระดับความรุนแรง: {legend} · "{SEV_CRIT}" ใช้กับระดับเคส ส่วนโพสต์ใช้ความเร่งด่วนจากระบบ (สูง / กลาง / ต่ำ)</p>
<div class="case-list">{''.join(rows)}</div>
<p class="foot-note">ข้อมูลเคสจากคอลัมน์ "เคส (เสนอ)" ใน xlsx · ผู้รับผิดชอบและสถานะเป็นค่าตัวอย่างในต้นแบบ</p>
"""
    return page(NAV_LABEL["cases"], web_shell("wcases", [("Dashboard", "w1.html"), (NAV_LABEL["cases"], None)], content, "cases"), "wcases", "web")


def state_cell(mock, when, sees, next_):
    return (f'<div class="st"><div class="mock">{mock}</div><dl><dt>เกิดเมื่อ</dt><dd>{when}</dd>'
            f'<dt>ผู้ใช้เห็น</dt><dd>{sees}</dd><dt>ทำต่อได้</dt><dd>{next_}</dd></dl></div>')


def build_w_states():
    ep, eb = EMPTY_COMBO if EMPTY_COMBO else ("YouTube", "บางนา")
    sel = lambda lab, val, dis="": f'<label class="field">{lab}<select{dis}><option>{val}</option></select></label>'
    rows = [
        ("ตัวกรอง", [
            state_cell(sel("สาขา", f"ทั้งหมด ({len(BRANCHES5)} สาขา)") + '<span class="chip on">' + CAMPAIGN + ' · 20–26 ต.ค. 69</span>',
                       "โหลดตัวเลือกครบ", "ค่าที่เลือกอยู่และชุดตัวกรองที่บันทึกไว้", "เปลี่ยนค่า บันทึกชุดตัวกรอง"),
            state_cell(sel("สาขา", "กำลังโหลดตัวเลือก…", " disabled aria-busy=\"true\"") + '<div class="skel" style="width:70%"></div>',
                       "เปิดหน้า หรือเปลี่ยนช่วงเวลา", "ช่องที่ยังเลือกไม่ได้ พร้อมคำว่า กำลังโหลด", "ใช้ตัวกรองอื่นที่โหลดแล้วได้"),
            state_cell(f'<p><b>ไม่พบโพสต์ตามตัวกรองนี้</b></p><p class="sub">{esc(ep)} · สาขา{esc(eb)} · {CAMPAIGN} · 19–25 ต.ค. 69 = 0 ครั้ง</p><button class="btn sm" type="button">ล้างตัวกรอง</button>',
                       "ชุดตัวกรองแคบจนไม่มีข้อมูล (ตัวอย่างนี้เป็นผลจริงจาก CSV)", "ตัวกรองที่ใช้อยู่ และบอกว่า 0 ไม่ใช่ระบบเสีย", "ล้างตัวกรอง หรือเอาเงื่อนไขออกทีละตัว"),
            state_cell(f'<div class="notice err">{icon("alert", label="ผิดพลาด", size=18)}<p>โหลดรายการสาขาไม่สำเร็จ · ตัวกรองเดิมยังใช้อยู่</p></div><button class="btn sm" type="button">{icon("refresh", size=16)}ลองใหม่</button>',
                       "เรียกตัวเลือกจากเซิร์ฟเวอร์ไม่สำเร็จ", "สาเหตุเท่าที่ทราบ และค่าที่ยังใช้อยู่", "ลองใหม่ หรือทำงานต่อด้วยตัวกรองเดิม"),
        ]),
        ("แผนภูมิ", [
            state_cell(svg_spark() + f'<p class="sub">19–25 ต.ค. 69 · รวม {fmt(S["total"])}</p>',
                       "มีข้อมูลในช่วงที่เลือก", "เส้นแนวโน้ม ตัวเลขกำกับ และแถบนอกเวลา", "คลิกช่วงเวลาเพื่อเปิด Social Feed"),
            state_cell('<div class="skel"></div><div class="skel" style="width:80%"></div><div class="skel" style="width:60%"></div><p class="sub">กำลังโหลดข้อมูล 19–25 ต.ค. 69…</p>',
                       "เปลี่ยนตัวกรองหรือช่วงเวลา", "โครงกราฟและช่วงเวลาที่กำลังโหลด", "รอ หรือเปลี่ยนตัวกรองต่อได้"),
            state_cell('<p><b>ไม่มีการกล่าวถึงในช่วงนี้</b></p><p class="sub">แกนยังแสดงที่ 0 เพื่อไม่ให้ดูเหมือนกราฟหาย</p>',
                       "ตัวกรองไม่มีข้อมูล", "แกนกราฟที่ 0 และข้อความอธิบาย", "ขยายช่วงเวลา หรือล้างตัวกรอง"),
            state_cell(f'<div class="notice err">{icon("alert", label="ผิดพลาด", size=18)}<p>โหลดกราฟไม่สำเร็จ · แสดงข้อมูลล่าสุดถึง {DATA_AS_OF}</p></div><button class="btn sm" type="button">{icon("refresh", size=16)}ลองใหม่</button>',
                       "ดึงข้อมูลรายชั่วโมงไม่สำเร็จ", "เวลาของข้อมูลล่าสุดที่แสดงอยู่", "ลองใหม่ หรือเปิด Social Feed แทน"),
        ]),
        ("ตัวควบคุมแก้ผล", [
            state_cell('<div class="diff"><span>ความรู้สึก: ลบ → <b>บวก</b></span><span>เหตุผล: ขบขัน</span></div><button class="btn primary sm" type="button">บันทึก</button>',
                       "เลือกค่าใหม่และเหตุผลแล้ว", "ค่าเดิมเทียบค่าใหม่ ปุ่มบันทึกกดได้", "บันทึก หรือยกเลิกเพื่อคืนค่าเดิม"),
            state_cell('<button class="btn primary sm" type="button" disabled aria-busy="true">กำลังบันทึก…</button><p class="sub">ปุ่มถูกล็อกกันกดซ้ำ ยังไม่แสดงว่าสำเร็จ</p>',
                       "กดบันทึกแล้ว รอเซิร์ฟเวอร์ยืนยัน", "สถานะกำลังบันทึก", "รอ (ต้นแบบใช้ 1.2 วินาที)"),
            state_cell('<button class="btn primary sm" type="button" disabled>บันทึก</button><p class="sub">เลือกเหตุผลก่อน ปุ่มบันทึกจึงกดได้</p>',
                       "เปิดแผงแล้วแต่ยังไม่ได้เลือกเหตุผล", "ปุ่มบันทึกที่ยังกดไม่ได้ พร้อมบอกสิ่งที่ขาด", "เลือกเหตุผล หรือปิดแผง"),
            state_cell(f'<div class="notice err">{icon("alert", label="ผิดพลาด", size=18)}<p>บันทึกไม่สำเร็จ · ยังไม่ได้บันทึกค่าใหม่ ร่างยังอยู่ครบ</p></div><button class="btn sm" type="button">{icon("refresh", size=16)}ลองใหม่</button>',
                       "บันทึกไม่สำเร็จ (สาธิตด้วยปุ่ม E ใน w3)", "ค่าที่แสดงยังเป็นค่าเดิม ร่างไม่หาย", "ลองใหม่โดยไม่ต้องกรอกเหตุผลซ้ำ"),
        ]),
    ]
    colors = [T["ok"], T["primary"], T["ink3"], T["crit"]]
    heads = "".join(f'<div class="colh"><i style="background:{c}"></i>{i + 1}. {n}</div>' for i, (n, c) in enumerate(zip(["ปกติ", "กำลังโหลด", "ว่าง", "ผิดพลาด"], colors)))
    grid = '<div></div>' + heads + "".join(f'<div class="rowh">{name}</div>' + "".join(cells) for name, cells in rows)
    content = f"""
<div class="page-head"><h1>สถานะของตัวกรอง แผนภูมิ และตัวควบคุมแก้ผล</h1></div>
<p class="sub">ใช้ชื่อสถานะชุดเดียวกันทุกส่วน (ปกติ / กำลังโหลด / ว่าง / ผิดพลาด) · ภาพนิ่ง ยกเว้นสถานะของตัวควบคุมแก้ผลที่ลองได้จริงใน <a href="w3.html">w3</a></p>
<div class="states">{grid}</div>
"""
    return page("สถานะของส่วนประกอบเว็บ", web_shell("wst", [("Dashboard", "w1.html"), ("สถานะของส่วนประกอบ", None)], content, ""), "wst", "web")


# ---------------- mobile
def build_m1():
    first = f"{C18['แพลตฟอร์ม']} {post_time(C18)}"
    ph = f'<span class="ph"><span class="logo" style="display:inline-flex;align-items:center;justify-content:center;color:var(--lime)">{icon("pulse", size=14)}</span>ชบา Monitor<time>{PUSH_TIME} น.</time></span>'
    content = f"""
<div class="lock">
  <h1 class="sr-only">การแจ้งเตือนบนหน้าจอล็อก</h1><!-- WCAG หน้าต้องมี h1 -->
  <p class="lk-date">{esc(DAY_TH[FRI])}</p>
  <p class="lk-time num">{PUSH_TIME}</p>
  <a class="push" id="push-full" href="m2.html"><!-- WCAG 2.5.3 ใช้ข้อความที่มองเห็นเป็นชื่อ -->
    {ph}
    <b>{sev_tag(SEV_CRIT, True)} K10 ความปลอดภัยอาหาร · สาขารังสิต</b><!-- F01 ไอคอนเตือนในป้ายวิกฤต -->
    <p>เชิงลบพุ่งนอกเวลาทำการ พบโพสต์ใน {len(SPREAD)} แพลตฟอร์มใน {SPREAD_MIN} นาที · เริ่มจาก {esc(first)}</p>
  </a>
  <!-- F03 push แบบซ่อนรายละเอียด สลับได้จากแผงโหมดทดลองใช้งาน -->
  <a class="push" id="push-private" href="m2.html" hidden>
    {ph}
    <b>{sev_tag(SEV_CRIT, True)} มีเหตุระดับ{SEV_CRIT} 1 เรื่อง · แตะเพื่อดู</b>
  </a>
  <p class="hint">{icon('lock', size=16)}แตะการแจ้งเตือนเพื่อเปิดเคส</p>
</div>"""
    note = step_note(["เปิดแจ้งเตือน"], f"<p>Push ถึง CS เวรตามผังเวร {esc(DAY_TH[FRI])} · {PUSH_TIME} น.</p><p>แตะการ์ดเพื่อเปิดหน้าสรุปเคสโดยตรง (ไม่ผ่านหน้าแรก)</p><p>สวิตช์ \"ซ่อนรายละเอียดบนหน้าจอล็อก\" ในแผงโหมดทดลองใช้งานแสดง push แบบไม่บอกประเด็นและสาขา</p>")
    body = f"""
<main id="main" class="stage">
  <div class="phone" role="region" aria-label="หน้าจอมือถือ 390 × 844 จอล็อก">{content}</div>
  <aside class="phone-note" aria-label="คำอธิบายหน้าจอ">{demo_panel()}{note}</aside>
</main>"""
    return page("m1 Push บนหน้าจอล็อก", body, "m1", "mobile")


def m2_core():
    br_guess = branch_parts(C19)[1]
    spread = "".join(f'<div><b>{esc(p["แพลตฟอร์ม"])}</b>{esc(post_time(p))}<br><span class="muted">{esc(engagement(p))}</span></div>' for p in SPREAD)
    return f"""
<div class="sev">{sev_tag(SEV_CRIT, True)}<span class="sr-only">ระดับเคส</span><p class="issue">{esc(C18['ประเด็น'])}</p></div><!-- F04 ไม่ใช้ "ความรุนแรงสูง" ซ้อนกับ "วิกฤต" · รวมแถวเพื่อให้ m2 ไม่ต้องเลื่อน -->
<h2 class="m-title">ลูกค้าแจ้งเจอชิ้นพลาสติกแข็งในลาเต้</h2>
<div class="mcard">
  <div class="kv"><b>สาขา{esc(C18['สาขา'])}</b><span class="tag ok">{icon('check', size=14)}ยืนยันจากรีวิวสาขา (C18)</span></div>
  <p class="sub" style="font-size:14px;margin-top:4px">{icon('clock', label='เวลา', size=16)} โพสต์แรก {esc(post_time(C18))} · {M2_ELAPSED} นาทีที่แล้ว · นอกเวลาทำการ</p>
</div>
<div class="mcard">
  <p style="font-size:14px"><b>พบโพสต์ใน {len(SPREAD)} แพลตฟอร์มใน {SPREAD_MIN} นาที</b> · ยังไม่พบเพจข่าว</p>
  <div class="spread">{spread}</div>
</div>
<!-- F07 บอกว่าระบบประเมินอะไร และมั่นใจเท่าไร -->
<div class="confrow"><span>ระบบประเมินว่า{SEV_CRIT} · มั่นใจ <b class="num">{conf_pct(C18)}%</b></span><span class="meter" aria-hidden="true"><span style="width:{conf_pct(C18)}%"></span></span></div>
<p style="font-size:13px;color:var(--warn);margin-top:-6px">สาขาในโพสต์ X (C19): {esc(br_guess)}</p>
<!-- F05 กรอบรูปหลักฐานจำลอง (ไม่ใช้รูปคนจริง) + ปุ่มเปิดโพสต์ต้นทาง -->
<div class="mcard ev" id="ev">
  <div class="ev-img" role="img" aria-label="กรอบรูปหลักฐานจำลอง: รูปจากโพสต์ต้นทาง">รูปจากโพสต์ต้นทาง</div>
  <p>"{mobile_quote(C18)}"</p>
  <div class="ev-src"><small class="muted" style="display:flex;align-items:center;gap:4px">{icon('eyeoff', label='ชื่อผู้โพสต์ถูกปกปิด', size=14)}{esc(C18['ผู้โพสต์ (ปกปิด)'])} · {esc(C18['แพลตฟอร์ม'])}</small>
    <button class="btn ghost" type="button" id="open-src">{icon('external', size=16)}เปิดโพสต์ต้นทาง</button></div>
</div>
<!-- F08 รับเรื่อง -->
<div class="accept" id="accept"><p class="m-status" id="accept-t">ยังไม่มีผู้รับเรื่อง · ยังไม่ตอบกลับสาธารณะ</p><button class="btn" type="button" id="accept-btn" data-time="{ACCEPT_TIME}">{icon('user', size=18)}รับเรื่องนี้</button></div>"""


def build_m2():
    actions = f"""
<a class="btn primary main" href="m4.html">{icon('send', size=20)}ยกระดับและแจ้งทีม</a>
<div class="pair"><a class="btn" href="m3.html">ดูภาพรวม</a><button class="btn ghost" type="button" id="not-urgent">ไม่ใช่เรื่องด่วน</button></div>"""
    dialog = f"""
<dialog id="dlg" aria-labelledby="dlg-h">
  <form method="dialog" id="dlg-form">
    <h2 id="dlg-h">ยืนยันว่าไม่ใช่เรื่องด่วน?</h2>
    <p class="sub" style="font-size:14px">แจ้งเตือนจะปิด และเคสเข้ากระบวนการเคสทั่วไปตามประเภทปัญหา · เลิกทำได้</p>
    <fieldset><legend>เหตุผล (จำเป็น)</legend>
      <label class="radio"><input type="radio" name="nu" value="ไม่เกี่ยวข้องกับแบรนด์">ไม่เกี่ยวข้องกับแบรนด์</label>
      <label class="radio"><input type="radio" name="nu" value="เกี่ยวกับเคสเดิม ไม่ใช่เหตุใหม่">เกี่ยวกับเคสเดิม ไม่ใช่เหตุใหม่</label>
      <label class="radio"><input type="radio" name="nu" value="อื่นๆ">อื่นๆ</label>
      <!-- F20 อื่นๆ มีช่องพิมพ์ -->
      <div id="nu-other-box" hidden><label class="sr-only" for="nu-other">ระบุเหตุผลอื่น</label><textarea id="nu-other" placeholder="ระบุเหตุผล" rows="2" style="min-height:64px"></textarea></div>
    </fieldset>
    <p class="hint muted" id="dlg-hint" style="font-size:13px">เลือกเหตุผลก่อน ปุ่มยืนยันจึงกดได้</p>
    <div class="dlg-actions"><button class="btn" value="cancel" type="submit" formnovalidate>กลับ</button><button class="btn primary" value="ok" type="submit" id="dlg-ok" disabled>ยืนยัน</button></div>
  </form>
</dialog>
<div class="toast" id="m2-toast" hidden role="status"><span class="grow" id="m2-toast-t"></span><button class="btn lime" type="button" id="m2-undo">{icon('undo', size=18)}เลิกทำ</button></div>"""
    note = step_note(["เปิดแจ้งเตือน", "ตัดสินความเร่งด่วน"],
                     "<p>ตัดสินใจได้ใน 10 วินาที ไม่ต้องเลื่อน (ตรวจที่ 390×844 และ 360×740 · ถ้าล้น ย่อรูปหลักฐานก่อน) ลำดับ: ความรุนแรง → ประเด็น → สาขา → เวลา → การลาม → ความเชื่อมั่น → หลักฐาน</p>"
                     "<p>ปุ่ม: ยกระดับ (ไป m4) · ดูภาพรวม (ไป m3 = ขั้นประเมิน) · ไม่ใช่เรื่องด่วน (กล่องยืนยัน + เลิกทำ) · รับเรื่องนี้</p>"
                     '<p><a class="solo-link" href="m-states.html">ดู 4 สถานะของตัวควบคุมนี้</a></p>')
    return page("m2 แจ้งเตือนเคส K10", phone_shell("m2", PUSH_TIME, m2_core(), note, header=(None, "เคส K10"), actions=actions, noscroll=True, extra=dialog), "m2", "mobile")


def build_m3():
    posts = "".join(f'<div class="mcard"><p class="muted" style="font-size:13px">{p["ID"]} · {esc(p["แพลตฟอร์ม"])} · {esc(post_time(p))} · {esc(engagement(p))}</p><p>"{mobile_quote(p)}"</p></div>' for p in SPREAD)
    content = f"""
<section class="mcard" aria-labelledby="m3h">
  <h2 id="m3h" style="font-size:15px">การกล่าวถึงรายชั่วโมง ตั้งแต่ 18:00 น. <span class="muted" style="font-weight:400">ทั้งหมด / ลบ</span></h2>
  <div class="legend"><span><i style="background:{T['neg']}"></i>ลบ</span><span><i style="background:{T['neu']};box-shadow:inset 0 0 0 1px {NEU_STROKE}"></i>อื่น ๆ</span></div>
  {svg_fri_bars()}
  <p class="sub" style="font-size:13px">แท่ง 22:00 น. เป็นยอดทั้งชั่วโมงจาก CSV · ปกติชั่วโมง 22:00 น. ของ จ.–พฤ. เฉลี่ย {S['base_22']:.1f} ครั้ง</p>
</section>
<p class="sub" style="font-size:14px">แพลตฟอร์มที่ลาม: {' → '.join(esc(p['แพลตฟอร์ม']) for p in SPREAD)} · เพจข่าว: ยังไม่พบ</p>
<h2 id="key-posts" style="font-size:16px">โพสต์สำคัญ {len(SPREAD)} รายการ</h2>
{posts}
<a href="w2.html?case=K10" style="display:inline-flex;align-items:center;gap:6px;min-height:44px">เปิดบนเว็บเพื่ออ่านความเห็นทั้งหมด {icon('external', size=14)}</a>"""
    actions = f'<a class="btn primary main" href="m4.html">{icon("right", size=20)}เลือกผู้รับและยกระดับ</a>'
    note = step_note(["ประเมิน"], "<p>ประเมินจาก Dashboard แบบย่อ: กราฟรายชั่วโมง แพลตฟอร์มที่ลาม เพจข่าว และโพสต์สำคัญ พร้อมทางเลือกเปิดบนเว็บ</p>")
    return page("m3 ภาพรวมเคส K10", phone_shell("m3", "22:43", content, note, header=("m2.html", "ภาพรวมเคส K10"), actions=actions, tabbar=tabbar(NAV_LABEL["overview"])), "m3", "mobile")


RECIPIENTS = [
    ("mgr", "ผู้จัดการสาขารังสิต", "ผู้รับหลัก · push + โทรอัตโนมัติ", True, "หลัก", False),
    ("qa", "QA เวรกลางคืน", "ตามตารางเวร · push", True, "", False),
    ("area", "ผู้จัดการเขต", f"ผู้รับสำรอง · ส่งเมื่อผู้รับหลักไม่รับทราบใน {ESCALATE_MIN} นาที", True, "สำรอง", False),  # F10
    ("exec", "ผู้บริหาร", "", True, "จำเป็น", True),  # ตาม FigCraft: ตัดคำอธิบายใต้ชื่อผู้บริหารออก
    ("pr", "ฝ่ายสื่อสารองค์กร", "สำหรับร่างคำตอบต่อสาธารณะ", False, "", False),
]


def build_m4():
    opts = "".join(f'<option{" selected" if b == "รังสิต" else ""}>{b}</option>' for b in BRANCHES5)
    rc = []
    for key, name, desc, checked, badge, locked in RECIPIENTS:
        if key == "pr":
            rc.append('<p class="sec-k">เพิ่มเติม</p>')
        if locked:
            # F09 ผู้บริหาร: ไอคอนกุญแจ + "บังคับตามกฎ" แทนช่องติ๊กสีเทา (ใบงานกำหนด ห้ามตัดออก)
            rc.append(f'<div class="locked rc-locked" data-name="{esc(name)}" data-role="{esc(badge)}"><span class="lk">{icon("lock", label="บังคับตามกฎ", size=22)}</span>'
                      f'<span><b><span class="rc-name">{esc(name)}</span><span class="tag pri">บังคับตามกฎ</span></b>{f"<small>{esc(desc)}</small>" if desc else ""}</span></div>')
            continue
        tag = f'<span class="tag pri">{"ผู้รับ" + badge}</span>' if badge else ""  # F10 ป้าย ผู้รับหลัก / ผู้รับสำรอง
        attrs = " checked" if checked else ""
        rc.append(f'<label><input type="checkbox" name="rc" value="{key}" data-name="{esc(name)}" data-role="{esc(badge)}"{attrs}><span><b><span class="rc-name">{esc(name)}</span>{tag}</b><small>{esc(desc)}</small></span></label>')
    content = f"""
<div class="stepper" aria-label="ขั้นตอน"><span class="on" aria-current="step"><i>1</i>ผู้รับ</span><span><i>2</i>ตรวจทาน</span><span><i>3</i>ส่ง</span></div>
<div class="notice info" id="undone" hidden role="status"></div>
<div class="branchbox">
  <fieldset><legend>ตรวจสาขา</legend>
    <p class="sub" style="font-size:13px">ผู้โพสต์ต้นเรื่องระบุ: สาขา{esc(C18['สาขา'])} (C18) · ในโพสต์ X (C19): {esc(branch_parts(C19)[1])}</p>
    <label class="radio"><input type="radio" name="br" value="confirm" checked>ยืนยันสาขารังสิต</label>
    <label class="radio"><input type="radio" name="br" value="change">เปลี่ยนสาขา</label>
    <div id="br-pick" hidden><label class="sr-only" for="br-sel">เลือกสาขา</label><select id="br-sel">{opts}</select></div>
    <label class="radio"><input type="radio" name="br" value="unknown">ยังไม่ทราบสาขา</label>
    <p class="notice warn" id="br-unknown" hidden style="font-size:13px">ส่งให้ทีมกลางหรือ QA เริ่มประเมินความเสี่ยงก่อน พร้อมป้าย "รอยืนยันสาขา"</p>
  </fieldset>
</div>
<p class="sec-k">แนะนำตามกฎ: ความปลอดภัยอาหาร + {SEV_CRIT}</p>
<div class="rcpt" id="rcpt">{''.join(rc)}</div>
<div class="mcard" style="border-style:dashed;font-size:13px;color:var(--ink2)">{icon('clock', label='เวลา', size=16)} {BACKUP_RULE} (ผู้จัดการเขต) · ปรับได้ในตั้งค่า</div><!-- F11 -->"""
    actions = '<a class="btn primary main" href="m5.html" id="to-review">ถัดไป: ตรวจทาน (<span id="rc-count">4</span> คน)</a>'
    note = step_note(["ตรวจสาขา", "ตรวจผู้รับ"], f"<p>ส่วนบน = ตรวจสาขา (ยืนยัน / เปลี่ยน / ยังไม่ทราบ) · ส่วนล่าง = ตรวจผู้รับ</p><p>ผู้บริหารแสดงเป็น \"บังคับตามกฎ\" ไม่ใช่ช่องติ๊ก เพราะกรณีที่ 1 ต้องยกระดับถึงผู้บริหาร</p><p>ผู้จัดการเขต = ผู้รับสำรอง ได้รับเมื่อผู้รับหลักไม่รับทราบใน {ESCALATE_MIN} นาที (ทีมกำหนด รอบ 1)</p>")
    return page("m4 เลือกผู้รับ", phone_shell("m4", "22:44", content, note, header=("m2.html", "เลือกผู้รับ"), actions=actions), "m4", "mobile")


RCPT_LABEL = {"หลัก": " (ผู้รับหลัก)", "สำรอง": f" (ผู้รับสำรอง · ส่งเมื่อผู้รับหลักไม่รับทราบใน {ESCALATE_MIN} นาที)", "จำเป็น": " (จำเป็น)"}  # F10 F13


def default_rcpt_list():
    return "".join(f"<li>{esc(n)}{RCPT_LABEL.get(b, '')}</li>" for _, n, _, c, b, _ in RECIPIENTS if c)


K10_MSG = "ลูกค้าแจ้งเจอชิ้นพลาสติกแข็งในลาเต้ที่สาขารังสิต พบโพสต์ใน X และ TikTok โปรดรับทราบและตรวจสอบที่สาขา"


def build_m5():
    content = f"""
<div class="stepper" aria-label="ขั้นตอน"><span class="done"><i>1</i>ผู้รับ</span><span class="on" aria-current="step"><i>2</i>ตรวจทาน</span><span><i>3</i>ส่ง</span></div>
<div class="mcard review">
  <p style="display:flex;gap:8px;align-items:center;flex-wrap:wrap;margin-bottom:10px">{sev_tag(SEV_CRIT, True)}<b>K10 · ความปลอดภัยอาหาร · <span id="rv-branch">สาขารังสิต</span></b></p>
  <dl>
    <dt>ส่งถึง</dt><dd><ul id="rv-list">{default_rcpt_list()}</ul></dd>
    <dt>ช่องทาง</dt><dd>push + โทรอัตโนมัติถึงผู้รับหลัก · ผู้รับสำรองได้ push เมื่อครบ {ESCALATE_MIN} นาที</dd>
    <dt>ข้อความ</dt><dd class="msgbox">
      <span class="tag warn" style="margin-bottom:4px">{icon('alert', size=13)}ยังไม่ยืนยันข้อเท็จจริง</span><!-- F15 -->
      <p id="msg-text">{K10_MSG}</p>
      <div id="msg-edit" hidden><label class="sr-only" for="msg-input">แก้ไขข้อความที่จะส่ง</label><textarea id="msg-input">{K10_MSG}</textarea>
        <div class="msg-acts"><button class="btn primary" type="button" id="msg-save">ใช้ข้อความนี้</button><button class="btn" type="button" id="msg-cancel">ยกเลิก</button></div></div>
      <button class="btn ghost" type="button" id="msg-edit-btn" style="min-height:44px;padding:4px 0">แก้ไขข้อความ</button><!-- F14 -->
    </dd>
    <dt>แนบ</dt><dd>ลิงก์ {len(SPREAD)} โพสต์ ({', '.join(p['ID'] for p in SPREAD)}) · ชื่อผู้โพสต์ปกปิด</dd>
  </dl>
</div>
<div class="notice info" style="flex-direction:column;gap:4px"><b>หลังกดส่ง</b>
  <ul style="margin:0;padding-left:18px;font-size:14px">
    <li>สถานะเคสเปลี่ยนเป็น "รอรับทราบ"</li>
    <li>เริ่มนับเวลารับทราบของแต่ละคน · {BACKUP_RULE}</li>
    <li>เลิกทำได้ภายใน <span class="placeholder">{PH_INTERVIEW}</span> วินาที (ต้นแบบใช้ 5 วินาที · เวลาสาธิต)</li>
  </ul>
</div>
<div id="send-msg" role="status" aria-live="polite"></div>"""  # F16 ไม่มี "สาธิต: กด E" บนจอแล้ว
    actions = f'<button class="btn primary main" type="button" id="send">{icon("send", size=20)}<span>ยืนยันส่ง</span></button><a class="btn" href="m4.html">กลับไปแก้ผู้รับ</a>'
    # F21 pop-up ส่งไม่สำเร็จ: สาเหตุ + ลองใหม่ + โทรหาผู้รับหลัก
    extra = f"""
<dialog id="fail-dlg" class="fail-dlg" aria-labelledby="fail-h" aria-describedby="fail-d">
  <h2 id="fail-h">ส่งไม่สำเร็จ</h2>
  <p class="cause">{icon('wifioff', label='ไม่มีสัญญาณ', size=20)}ไม่มีสัญญาณอินเทอร์เน็ต</p>
  <p class="sub" style="font-size:14px" id="fail-d">ยังไม่ได้ส่งถึงใคร · ผู้รับและข้อความยังอยู่ครบ · ไม่สร้างเคสซ้ำ</p>
  <p class="sub" style="font-size:14px;margin-top:6px" id="call-msg" role="status"></p>
  <div class="dlg-actions">
    <button class="btn primary" type="button" id="retry">{icon('refresh', size=18)}ลองใหม่</button>
    <button class="btn" type="button" id="call">{icon('phone', size=18)}<span>โทรหาผู้รับหลัก · <span id="call-name">ผู้จัดการสาขารังสิต</span></span></button>
    <button class="btn ghost" type="button" id="fail-close">ปิด</button>
  </div>
</dialog>"""
    note = step_note(["ทบทวน", "ยืนยันส่ง"], "<p>ทบทวนครบ: ความรุนแรง สาขา ผู้รับ ช่องทาง ข้อความ หลักฐาน</p><p>ยืนยันส่ง → กำลังส่ง… 1.2 วินาที → m6 · เปิดสวิตช์ \"จำลองส่งไม่สำเร็จ\" ในแผงโหมดทดลองใช้งาน = pop-up บอกสาเหตุ + ลองใหม่ + โทรหาผู้รับหลัก</p>")
    return page("m5 ตรวจทานก่อนส่ง", phone_shell("m5", "22:45", content, note, header=("m4.html", "ตรวจทานก่อนส่ง"), actions=actions, extra=extra), "m5", "mobile")


SENT_TIME = "22:46"
FIRST_RESP_ELAPSED = minutes(SENT_TIME) - minutes(post_time(C18))  # F18 = 66 นาที คำนวณจากเวลาโพสต์ C18


def ack_tag(role):
    # F10 ผู้รับสำรองยังไม่ได้รับจนกว่าผู้รับหลักจะไม่รับทราบใน 5 นาที
    if role == "สำรอง":
        return f'<span class="tag plain">ผู้รับสำรอง · ส่งถ้าผู้รับหลักไม่รับทราบใน {ESCALATE_MIN} นาที</span>'
    return '<span class="tag pri">ส่งถึงแล้ว · รอรับทราบ</span>'


def build_m6():
    # R3 ผู้รับสำรอง: ป้ายยาวอยู่ใต้ชื่อ ไม่บีบชื่อ (เหมือน FigCraft)
    acks = "".join((f'<li class="stack">{icon("user", label="ผู้รับ", size=20)}<span class="who"><span>{esc(n)}</span>{ack_tag(b)}</span></li>' if b == "สำรอง" else
                    f'<li>{icon("user", label="ผู้รับ", size=20)}<span>{esc(n)}</span>{ack_tag(b)}</li>') for _, n, _, c, b, _ in RECIPIENTS if c)
    content = f"""
<div class="done-ic">{icon('check', label='สำเร็จ', size=34)}</div>
<h2 style="text-align:center;font-size:24px">ส่งแล้ว {SENT_TIME} น.</h2>
<p class="sub" style="text-align:center">เคส K10 สถานะ "รอรับทราบ"</p>
<div class="undo-bar" id="undo-bar" role="status"><span class="grow" id="undo-text">เลิกทำได้อีก <b class="num" id="undo-sec">5</b> วินาที (เวลาสาธิต)</span><button class="btn lime" type="button" id="undo">{icon('undo', size=18)}เลิกทำ</button></div>
<section class="mcard" aria-labelledby="ack-h"><h3 id="ack-h" style="font-size:15px">การรับทราบรายคน</h3><ul class="ack" id="ack">{acks}</ul></section>
<div class="mcard" style="font-size:13px;color:var(--ink2)">เป้าหมายรับทราบ: ผู้จัดการสาขา <span class="placeholder">{PH_INTERVIEW}</span> นาที · QA <span class="placeholder">{PH_INTERVIEW}</span> นาที · {BACKUP_RULE} (ผู้จัดการเขต)</div>
<div class="notice warn" style="flex-direction:column;gap:2px"><b>คำตอบแรกต่อสาธารณะ: ยังไม่มี</b>
  <p class="elapsed" role="timer">ผ่านมา <b class="num">{FIRST_RESP_ELAPSED}</b> นาที นับจากโพสต์แรก {esc(post_time(C18))}</p><!-- F18 -->
  <span style="font-size:14px">เป้าหมายภายใน <span class="placeholder">{PH_INTERVIEW}</span> นาที</span></div>"""
    actions = '<div class="pair"><button class="btn" type="button" aria-label="ถอน/เปลี่ยนผู้รับ (ภาพนิ่ง)">ถอน/เปลี่ยนผู้รับ</button><a class="btn primary" href="m2.html">กลับไปที่เคส</a></div>'
    note = step_note(["ติดตาม"], f"<p>ส่งสำเร็จ ≠ รับทราบ ≠ ปิดเคส · แสดงสถานะแยกรายคน · ผู้รับสำรองยังไม่ได้รับจนครบ {ESCALATE_MIN} นาที</p><p>แถบเลิกทำนับถอยหลัง 5 วินาที (เวลาสาธิต; ค่าจริงรอผลสัมภาษณ์) · เลิกทำแล้วกลับ m4</p>")
    return page("m6 ส่งสำเร็จ", phone_shell("m6", SENT_TIME, content, note, header=(None, "ยกระดับเคส K10"), actions=actions, tabbar=tabbar(NAV_LABEL["cases"])), "m6", "mobile")


def build_m_states():
    head = f'<p style="display:flex;gap:6px;align-items:center">{sev_tag(SEV_CRIT, True)}<b>K10 · สาขารังสิต</b></p><p><b>ลูกค้าแจ้งเจอชิ้นพลาสติกแข็งในลาเต้</b></p>'  # F04
    sending = "".join(f'<li style="display:flex;justify-content:space-between;padding:4px 0;font-size:13px"><span>{n}</span><span class="{c}">{s}</span></li>'
                      for n, s, c in [("ผจก.สาขารังสิต", "ส่งถึงแล้ว", "pass"), ("QA เวรกลางคืน", "กำลังส่ง…", ""), ("ผู้จัดการเขต", f"สำรอง · รอ {ESCALATE_MIN} นาที", "muted"), ("ผู้บริหาร", "รอคิว", "muted")])
    cells = [
        ("ปกติ", T["ok"], head + '<p class="sub" style="font-size:13px">ยังไม่มีผู้รับเรื่อง · ยังไม่ตอบกลับสาธารณะ</p><a class="btn primary" href="m4.html" style="min-height:52px">ยกระดับและแจ้งทีม</a><div class="pair" style="display:grid;grid-template-columns:1fr 1fr;gap:6px"><a class="btn" href="m3.html" style="min-height:44px">ดูภาพรวม</a><span class="btn ghost" style="min-height:44px">ไม่ใช่เรื่องด่วน</span></div>',
         "เคสใหม่ที่ยังไม่มีผู้รับ", "ปุ่มหลัก 1 ปุ่ม + ปุ่มรอง 2 ปุ่ม อยู่ในระยะนิ้วโป้ง", "ส่งแจ้งเตือนในทีม ดูภาพรวม หรือบอกว่าไม่ด่วน"),
        ("กำลังโหลด", T["primary"], head + f'<ul style="list-style:none;margin:0;padding:0">{sending}</ul><button class="btn primary" type="button" disabled aria-busy="true" style="min-height:52px">กำลังส่ง…</button><p class="sub" style="font-size:12px">ปิดแอปได้ ระบบส่งต่อให้จนครบ</p>',
         "กดยืนยันส่งแล้ว ระบบกำลังส่งถึงผู้รับ", "สถานะรายคน และปุ่มหลักถูกล็อกกันกดซ้ำ", f"รอ หรือปิดแอป · ส่งครบแล้วไปหน้าส่งสำเร็จ (m6) ที่เลิกทำได้ภายใน {PH_INTERVIEW} วินาที"),
        ("ว่าง", T["ink3"], f'<div style="text-align:center">{icon("check", label="ไม่มีเคสค้าง", size=28)}<p><b>ไม่มีเคสที่รอคุณรับทราบ</b></p><p class="sub" style="font-size:13px">ระบบยังเฝ้าดู {S["n_platforms"]} แพลตฟอร์มอยู่</p></div><a class="btn" href="m3.html" style="min-height:44px">ดูภาพรวม</a>',
         "ไม่มีเคสที่รอผู้ใช้คนนี้", "ยืนยันว่าระบบยังทำงาน ไม่ใช่หน้าที่โหลดไม่ขึ้น", "ไปดูภาพรวม"),
        ("ผิดพลาด", T["crit"], head + f'<div class="notice err" style="font-size:13px">{icon("wifioff", label="ไม่มีสัญญาณ", size=18)}<p><b>ส่งไม่สำเร็จ</b> · ไม่มีสัญญาณอินเทอร์เน็ต · ยังไม่ได้ส่งถึงใคร ผู้รับและข้อความยังอยู่ครบ</p></div><a class="btn primary" href="m5.html" style="min-height:52px">{icon("refresh", size=18)}ลองใหม่</a><span class="btn" style="min-height:44px">{icon("phone", size=18)}โทรหาผู้รับหลัก</span>',
         "ส่งไม่สำเร็จ เช่น ไม่มีสัญญาณ", "สาเหตุ และยืนยันว่าผู้รับกับข้อความไม่หาย", "ลองอีกครั้ง หรือโทรตรงเป็นทางสำรอง"),
    ]
    out = []
    for i, (name, c, mock, w, s, n) in enumerate(cells, 1):
        out.append(f'<div class="st"><p class="colh"><i style="background:{c}"></i>{i}. {name}</p><div class="mock" style="background:#fff;border:1px solid var(--line);min-height:330px;justify-content:flex-start">{mock}</div>'
                   f'<dl><dt>เกิดเมื่อ</dt><dd>{w}</dd><dt>ผู้ใช้เห็น</dt><dd>{s}</dd><dt>ทำต่อได้</dt><dd>{n}</dd></dl></div>')
    body = f"""<main id="main"><div class="wrap" style="max-width:1320px;margin:0 auto;padding:40px 28px 80px;display:flex;flex-direction:column;gap:16px">
<h1>4 สถานะของตัวควบคุมการแจ้งเตือน (มือถือ)</h1>
<p class="sub">ใช้ชื่อสถานะชุดเดียวกับเว็บ · ตัวควบคุมคือส่วนปุ่มดำเนินการท้ายหน้าแจ้งเตือน (<a href="m2.html">m2</a>) · สำหรับตัวนี้ "กำลังโหลด" คือกำลังส่ง · กว้างการ์ดละประมาณจอมือถือ</p>
<div style="display:grid;grid-template-columns:repeat(4,1fr);gap:16px">{''.join(out)}</div>
<p class="foot-note">ทดลองสถานะ "กำลังโหลด" และ "ผิดพลาด" ได้จริงที่ <a href="m5.html">m5</a> (เปิดสวิตช์ "จำลองส่งไม่สำเร็จ" ในแผงโหมดทดลองใช้งาน ก่อนยืนยันส่ง)</p>
</div></main>"""
    return page("m-states สถานะตัวควบคุมการแจ้งเตือน", body, "mst", "doc")


# ---------------- other pages
def build_index():
    links = [("w-cases.html", NAV_LABEL["cases"] + " (เว็บ)"), ("w-states.html", "สถานะของส่วนประกอบเว็บ"), ("m-states.html", "สถานะของตัวควบคุมการแจ้งเตือน"),
             ("rationale.html", "เหตุผลลำดับข้อมูลมือถือ vs เว็บ"), ("foundations.html", "พื้นฐานการออกแบบ"), ("data.html", "ที่มาของตัวเลข")]
    body = f"""<main id="main" class="wrap">
<div style="display:flex;gap:14px;align-items:center"><span class="logo" style="width:52px;height:52px;border-radius:14px">{icon('pulse', size=26)}</span>
<div><h1>ทดลองใช้งาน ชบา Monitor</h1><p class="sub">IS757 · แบรนด์สมมติ ชบา คอฟฟี่ ({len(BRANCHES5)} สาขา: {' '.join(BRANCHES5)}) · ข้อมูลทั้งหมดเป็นข้อมูลจำลอง</p></div></div>
<div class="cases">
  <a class="case" href="m1.html"><span><span class="tag pri">กรณีที่ 1</span> <span class="muted">มือถือ</span></span>
    <h2>ได้รับแจ้งเตือนนอกเวลา แล้วตัดสินใจว่าจะทำอะไรต่อ</h2>
    <p class="sub">สมมติว่าคุณเป็นเจ้าหน้าที่เวร {esc(DAY_TH[FRI])} · {PUSH_TIME} น. โทรศัพท์มีการแจ้งเตือนเข้ามา ให้ทำเหมือนที่คุณจะทำจริง</p>
    <span class="go">เริ่ม {icon('right', size=18)}</span></a>
  <a class="case" href="w1.html"><span><span class="tag pri">กรณีที่ 2</span> <span class="muted">เว็บ</span></span>
    <h2>ตรวจโพสต์ที่ระบบวิเคราะห์ความรู้สึกผิด แล้วแก้ไข</h2>
    <p class="sub">คุณเปิดระบบในเวลางาน เริ่มจาก Dashboard ลองหาโพสต์ที่ระบบเข้าใจผิดและแก้ผลให้ถูก</p>
    <span class="go">เริ่ม {icon('right', size=18)}</span></a>
</div>
<div class="notice warn">{icon('alert', label='หมายเหตุ', size=18)}<p>ต้นแบบนี้ยังไม่สมบูรณ์ บางปุ่ม เช่น ตัวกรองแบบเลื่อนลงใน Social Feed ยังกดไม่ได้ ถ้ากดแล้วไม่เกิดอะไร ให้บอกผู้สังเกตการณ์ว่าคุณคาดว่าจะเกิดอะไร</p></div>
<nav aria-label="หน้าประกอบ"><p class="sec-k" style="margin-bottom:8px">หน้าประกอบสำหรับผู้ตรวจ</p><div class="links">{''.join(f'<a class="btn" href="{h}">{t}</a>' for h, t in links)}</div></nav>
<p class="disclaimer">ร่างโดยใช้ AI ช่วย ข้อมูลทั้งหมดเป็นข้อมูลสมมติ</p>
</main>"""
    return page("หน้าเริ่ม", body, "index", "doc")


def build_rationale():
    rows = [
        ("1", "ความรุนแรง (ป้าย วิกฤต)", "อยู่ในแถบเคสที่ยังไม่ปิด รองจากตัวกรองและตัวเลขหลัก",
         f"ผู้ใช้ต้องรู้ก่อนว่าต้องรีบหรือไม่ จึงอยู่บนสุด ใช้ข้อความ \"{SEV_CRIT}\" คู่สีแดงและไอคอนเตือน ไม่พึ่งสีอย่างเดียว · ชื่อระดับชุดเดียวทุกหน้า: {' · '.join(SEV_LEVELS)} (\"{SEV_CRIT}\" ใช้กับระดับเคส ส่วนโพสต์ใช้ความเร่งด่วนจาก xlsx)",
         f"ลำดับชั้นทางสายตา (docx) · WCAG 2.2 1.4.1 Use of Color · หลักฐานสัมภาษณ์ {PH_EVIDENCE}"),
        ("2", "ประเด็น", "ตัวกรอง \"ประเด็น\" และแท็บประเด็นร้องเรียน",
         "ประเภทปัญหากำหนดผู้รับ (ความปลอดภัยอาหาร → ผู้จัดการสาขาและ QA ตามตารางระดับใน docx)",
         "Nielsen ข้อ 2 match between system and real world (ใช้คำเดียวกับเกณฑ์ความรุนแรง)"),
        ("3", "สาขา + สถานะยืนยัน", "ตัวกรองสาขา · กราฟเชิงลบแยกสาขา",
         "ต้องรู้ว่าส่งถึงผู้จัดการสาขาไหน และสาขามาจากผู้โพสต์หรือระบบเดา",
         "docx: \"P3: ยืนยันสาขาและค้นหาผู้ติดต่อใช้เวลา\" (ข้อค้นพบจากบันทึกสัมภาษณ์)"),
        ("4", "เวลา (โพสต์แรก · ผ่านไปกี่นาที · นอกเวลา)", "แกนเวลาของกราฟ 7 วัน",
         "บนมือถือแปลงเป็นเวลาที่ผ่านไป เพื่อไม่ต้องคำนวณเอง",
         "Shneiderman: reduce short-term memory load"),
        ("5", "การลาม (แพลตฟอร์ม · เวลา · มีส่วนร่วม)", "กราฟความรู้สึกแยกแพลตฟอร์ม",
         "ผู้ให้สัมภาษณ์ระบุว่าต้องรู้ยอดการมีส่วนร่วมหรือการลามก่อนโทรแจ้งหัวหน้า",
         "docx อ้าง P2 และ P3 ข้อ 9"),
        ("6", "ความเชื่อมั่น", "คอลัมน์ความมั่นใจใน Social Feed (ป้าย รอตรวจสอบ < 75%)",
         "ใบงานกำหนดให้ค่าจาก AI แสดงความเชื่อมั่นหรือหลักฐาน วางก่อนหลักฐานเพื่อบอกว่าควรเชื่อผลแค่ไหน",
         "ข้อกำหนดใบงาน (docx) · Nielsen ข้อ 1 visibility of system status"),
        ("7", "หลักฐาน (ข้อความต้นเรื่อง + คำไฮไลต์)", "หน้ารายละเอียดโพสต์ (w3)",
         "บนมือถือแสดงแค่ข้อความต้นเรื่อง 1 ข้อความ ส่วนความเห็นทั้งหมดเปิดบนเว็บ",
         "การเปิดเผยข้อมูลทีละขั้น (progressive disclosure) · Nielsen ข้อ 8"),
        ("8", "ปุ่ม 3 ทาง", "ปุ่มกระจายตามหน้า",
         "ปุ่มหลัก 1 ปุ่ม (56 px) อยู่ล่างสุดในระยะนิ้วโป้ง ปุ่มรอง 2 ปุ่ม (≥ 44 px) · \"ไม่ใช่เรื่องด่วน\" ต้องยืนยันและเลิกทำได้",
         "กฎของ Hick (ทางเลือกน้อย) · กฎของ Fitts (เป้าใหญ่ ใกล้นิ้ว) · WCAG 2.2 2.5.8 Target Size (Minimum) · Nielsen ข้อ 5 · Shneiderman: permit easy reversal of actions"),
    ]
    trs = "".join(f"<tr><td>{a}</td><td><b>{b}</b></td><td>{c}</td><td>{d}</td><td>{e}</td></tr>" for a, b, c, d, e in rows)
    steps = "".join(f"<tr><td>{i}</td><td>{s}</td><td>{STEP_PAGE[s]}</td></tr>" for i, s in enumerate(STEPS, 1))
    body = f"""<main id="main" class="wrap prose">
<h1>เหตุผล: ลำดับข้อมูลมือถือต่างจากเว็บ <span class="draft">ร่าง</span></h1>
<p class="sub">เว็บใช้ลำดับ <b>overview first, zoom and filter, then details-on-demand</b> (Shneiderman) เพราะเจ้าหน้าที่ตรวจหลายรายการในเวลางาน: ตัวกรอง → ตัวเลขหลัก → กราฟ → แยกแพลตฟอร์ม/สาขา → Social Feed → รายละเอียด
ส่วนมือถือ (m2) ใช้ลำดับ <b>ตัดสินใจก่อน</b> เพราะผู้ใช้ถูกปลุกด้วย push นอกเวลาทำการ และต้องตัดสินใจเบื้องต้นได้ใน 10 วินาทีโดยไม่เลื่อนจอ (เกณฑ์ในใบงาน ตามที่ docx อ้าง)</p>
<table>
<caption class="sr-only">เทียบลำดับข้อมูลบนมือถือกับตำแหน่งบนเว็บ</caption>
<thead><tr><th>ลำดับบนมือถือ</th><th>ข้อมูล</th><th>บนเว็บอยู่ที่</th><th>เหตุผล</th><th>หลักที่อ้าง</th></tr></thead>
<tbody>{trs}</tbody></table>
<h2>สิ่งที่มือถือตัดออกจากจอแรก</h2>
<ul><li>ตัวกรองหลายมิติ ตัวเลขหลัก 4 ตัว และกราฟ 7 วัน → ย้ายไป m3 (เปิดเมื่อต้องการ) ใช้ชิปสำเร็จรูปแทนตัวกรอง ตามกฎของ Hick (docx)</li>
<li>การแก้ผลวิเคราะห์ → ทำบนเว็บเท่านั้น เพราะต้องอ่านบริบทและลดโอกาสแตะพลาด (docx)</li></ul>
<h2>เกณฑ์การเข้าถึงที่ใช้ (WCAG 2.2)</h2>
<ul><li>1.4.3 Contrast (Minimum): ตัวอักษร ≥ 4.5:1 · 1.4.11 Non-text Contrast: กราฟิกและขอบช่องกรอก ≥ 3:1 (ผลคำนวณอยู่ใน <a href="foundations.html">foundations</a>)</li>
<li>2.5.8 Target Size (Minimum): ≥ 24×24 px · ต้นแบบใช้มือถือ ≥ 44 px และปุ่มหลัก 52–56 px</li>
<li>1.4.1 Use of Color: สถานะทุกตัวมีข้อความกำกับ · 2.4.7 Focus Visible: วงโฟกัส 3 px สี primary · 2.4.11 Focus Not Obscured (Minimum): แถบปุ่มล่างไม่ทับส่วนที่โฟกัส (ต้องทดสอบเพิ่ม)</li></ul>
<h2>ขั้นใน User Flow กรณีที่ 1 (docx) กับหน้าจอ</h2>
<table><thead><tr><th>#</th><th>ขั้น</th><th>หน้าจอ</th></tr></thead><tbody>{steps}</tbody></table>
<h2>สิ่งที่ยังต้องยืนยัน</h2>
<ul><li>เวลาเป้าหมายรับทราบและคำตอบแรก 5/30/60 นาที: <span class="placeholder">{PH_INTERVIEW}</span> (docx ระบุว่ายังไม่ใช่ SLA ที่องค์กรอนุมัติ)</li>
<li>เวลาส่งต่อผู้รับสำรอง: {ESCALATE_MIN} นาที (ทีมกำหนดหลังทดสอบรอบ 1 ยังไม่ใช่ SLA ที่องค์กรอนุมัติ) · ระยะเลิกทำหลังส่ง: <span class="placeholder">{PH_INTERVIEW}</span></li>
<li>รายละเอียดกฎการส่งถึงผู้บริหาร (ใครมีอำนาจยกระดับ ช่องทาง): <span class="placeholder">{PH_CONFIRM}</span> · ผู้บริหารอยู่ในเส้นทางกรณีที่ 1 เสมอตามใบงาน</li>
<li>หลักฐานว่าทำไมความรุนแรงต้องมาก่อนประเด็น: <span class="placeholder">{PH_EVIDENCE}</span></li></ul>
</main>"""
    return page("เหตุผลลำดับข้อมูล", body, "rationale", "doc")


def build_foundations():
    rows = contrast_rows()
    trs = "".join(
        f'<tr><td>{esc(r["name"])}</td><td><span class="pill-sw" style="background:{r["fgv"]}"></span> {r["fgv"]}</td>'
        f'<td><span class="pill-sw" style="background:{r["bgv"]}"></span> {r["bgv"]}</td><td>{"ตัวอักษร" if r["kind"] == "text" else "กราฟิก"}</td>'
        f'<td class="num">{r["ratio"]:.2f}:1</td><td>≥ {r["need"]}:1</td>'
        f'<td class="{"pass" if r["ok"] else "fail"}">{"ผ่าน" if r["ok"] else "ไม่ผ่าน"}</td><td>{esc(r["where"])}{("<br><small>" + esc(r["mitig"]) + "</small>") if r["mitig"] else ""}</td></tr>'
        for r in rows)
    sw = [("พื้นหลัง", "bg"), ("การ์ด", "card"), ("ตัวอักษรหลัก", "ink"), ("ตัวอักษรรอง", "ink2"), ("ตัวอักษรเบา", "ink3"),
          ("primary", "primary"), ("lime (คู่ตัวอักษรเข้มเท่านั้น)", "lime"), ("วิกฤต", "crit"), ("เตือน", "warn"), ("สำเร็จ", "ok"),
          ("กราฟ ลบ", "neg"), ("กราฟ เป็นกลาง", "neu"), ("กราฟ ผสม", "mix"), ("กราฟ บวก", "pos"), ("ขอบช่องกรอก", "control")]
    sws = "".join(f'<div class="sw"><i style="background:{T[k]};{"box-shadow:inset 0 -1px 0 " + T["line"] if k in ("card",) else ""}"></i><div><b>{n}</b><br>{T[k]}</div></div>' for n, k in sw)
    body = f"""<main id="main" class="wrap">
<h1>พื้นฐานการออกแบบ</h1>
<p class="sub">ค่าทั้งหมดอยู่ใน tokens.css (สร้างโดย build.py) · ใช้ร่วมกันทุกหน้า</p>
<section class="card"><h2>สี</h2><div class="swatches" style="margin-top:12px">{sws}</div></section>
<section class="card"><h2>ความต่างสีตาม WCAG (คำนวณใน build.py)</h2>
<p class="sub">ตัวอักษร ≥ 4.5:1 (1.4.3) · กราฟิกและขอบส่วนควบคุม ≥ 3:1 (1.4.11) · ตัวอักษรบนแท่งกราฟ build.py เลือกสีเข้มหรือขาวที่ผ่าน 4.5:1 ให้อัตโนมัติ</p>
<table style="margin-top:12px"><thead><tr><th>คู่สี</th><th>หน้า</th><th>พื้น</th><th>ประเภท</th><th>ค่า</th><th>เกณฑ์</th><th>ผล</th><th>ใช้ที่</th></tr></thead><tbody>{trs}</tbody></table></section>
<section class="card"><h2>ตัวอักษร</h2>
<p>ฟอนต์ IBM Plex Sans Thai (Google Fonts) · ระยะบรรทัด 1.6 สำหรับเนื้อความ และ 1.5 สำหรับหัวข้อ (ไม่ต่ำกว่า 1.5)</p>
<p style="font-size:26px;font-weight:700;line-height:1.5">หัวเรื่อง 26 · Dashboard</p><p style="font-size:18px;font-weight:700;line-height:1.5">หัวข้อ 18 · การกล่าวถึงรายชั่วโมง</p><p>เนื้อความ 15 · {esc(C09['ข้อความ'])}</p><p style="font-size:13px" class="sub">คำอธิบาย 13 · ข้อมูลถึง {DATA_AS_OF}</p></section>
<section class="card"><h2>ระยะ เป้าสัมผัส และรูปแบบ</h2><ul>
<li>ระยะ: 4 · 8 · 12 · 16 · 20 · 24 px · มุมโค้งการ์ด 16 px</li>
<li>เป้าสัมผัส: ขั้นต่ำ 24×24 px · เว็บ 32–44 px · มือถือ ≥ 44 px · ปุ่มหลักมือถือ 52–56 px</li>
<li>วันเวลาไทย: {esc(thai_dt(C18['วันเวลา']))} · ใช้ พ.ศ. ย่อ และนาฬิกา 24 ชั่วโมง</li>
<li>ส่วนควบคุมใช้ &lt;button&gt; &lt;a href&gt; &lt;input&gt;+&lt;label&gt; ของจริง · ไอคอนที่สื่อความหมายเองมี aria-label · ไอคอนที่อยู่คู่ข้อความซ่อนจากโปรแกรมอ่านจอ (aria-hidden) เพื่อไม่ให้อ่านซ้ำ</li>
<li>ไม่ใช้อีโมจิ: อีโมจิในข้อความลูกค้า (C31, C32) แสดงเป็นป้ายข้อความ เช่น <span class="emo">[อีโมจิร้องไห้]</span></li>
<li>ข้อความที่มีข้อมูลสุขภาพ (C16) ซ่อนในรายการตามหมายเหตุ PDPA ในชุดข้อมูล</li></ul></section>
</main>"""
    return page("พื้นฐานการออกแบบ", body, "foundations", "doc")


def build_data():
    daily = "".join(f'<tr><td>{esc(DAY_TH[d])}</td><td class="num">{fmt(S["daily"][d])}</td><td class="num">{fmt(S["daily_neg"][d])}</td><td class="num">{S["camp_daily"][d] or "—"}</td></tr>' for d in DAYS)
    plat = "".join(f'<tr><td>{esc(p)}</td>' + "".join(f'<td class="num">{c[s]}</td>' for s in SENT_ORDER) + f'<td class="num"><b>{t}</b></td></tr>' for p, c, t in S["plat"])
    br = "".join(f'<tr><td>{esc(branch_name(b))}</td><td class="num">{ng}</td><td class="num">{t}</td><td class="num">{ng / t * 100:.1f}%</td></tr>' for b, ng, t in S["branches"])
    sents = " · ".join(f"{s} {fmt(S['by_sent'][s])}" for s in SENT_ORDER)
    body = f"""<main id="main" class="wrap prose">
<h1>ที่มาของตัวเลข</h1>
<p class="sub">ทุกตัวเลขบน Dashboard คำนวณใน build.py จาก <b>data/IS757_ตัวเลขรายชั่วโมงสมมติ.csv</b> ({fmt(len(CSVROWS))} แถว · หนึ่งแถว = ชั่วโมง × แพลตฟอร์ม × สาขา × ความรู้สึก × แคมเปญ · คอลัมน์ "จำนวนการกล่าวถึง" คือค่าที่นำมารวม)</p>
<section class="card"><h2>ค่าที่ตรวจก่อนวาด (ช่วง 19–25 ต.ค. 69)</h2>
<table style="margin-top:10px"><thead><tr><th>รายการ</th><th>ค่าคำนวณ</th><th>ค่าที่กำหนด</th><th>ผล</th></tr></thead><tbody>
<tr><td>ทั้งหมด</td><td class="num">{fmt(S['total'])}</td><td class="num">2,122</td><td class="pass">ตรง</td></tr>
<tr><td>เชิงลบ</td><td class="num">{fmt(S['neg'])}</td><td class="num">909</td><td class="pass">ตรง</td></tr>
<tr><td>ชั่วโมงสูงสุด</td><td>{esc(S['peak_th'])} = {S['peak']} ครั้ง</td><td>ศ. 23 ต.ค. 23:00 น. = 174</td><td class="pass">ตรง</td></tr>
</tbody></table>
<p class="foot-note" style="margin-top:8px">build.py หยุดทำงาน (assert) ถ้าค่าใดไม่ตรง</p></section>
<section class="card"><h2>กติกาการนับ</h2><ul>
<li>Dashboard ใช้ช่วง 19–25 ต.ค. 69 เท่านั้น · ยอดวันที่ 26 ต.ค. ({fmt(S['total_26'])} ครั้ง ถึงชั่วโมง {S['last_hour_26'][11:]} น.) ไม่นำมารวม เพราะยังไม่ครบวัน · ข้อมูลถึง {DATA_AS_OF}</li>
<li>สัดส่วนเชิงลบ = {fmt(S['neg'])} ÷ {fmt(S['total'])} = {S['neg_pct']:.1f}% · แยกความรู้สึก: {sents}</li>
<li>นอกเวลาทำการ = แถวที่คอลัมน์ "นอกเวลาทำการ" เป็น "ใช่" (20:00–07:00 น.) รวม {fmt(S['off_total'])} ครั้ง</li>
<li>ชั่วโมงที่สูงรองลงมา: {esc(S['second_th'])} = {S['second']} ครั้ง</li>
<li>"ปกติชั่วโมงนี้" = ค่าเฉลี่ยชั่วโมงเดียวกันของ จ.–พฤ. (19–22 ต.ค.): 23:00 น. เฉลี่ย {S['base_peak_hour']:.1f} · 22:00 น. เฉลี่ย {S['base_22']:.1f}</li>
<li>แคมเปญ {CAMPAIGN}: นับแถวที่คอลัมน์แคมเปญตรงชื่อ · {S['camp_days'][0][8:]}–{S['camp_days'][-1][8:]} ต.ค. รวม {S['camp_total']} (ลบ {S['camp_neg']}) · วันที่ 26 ต.ค. ไม่ครบวัน ({S['camp_26']} ครั้ง) จึงไม่แสดงในกราฟ แม้ชุดตัวกรองที่บันทึกไว้ชื่อ 20–26 ต.ค. 69</li>
<li>หน้า m3 แท่ง 22:00 น. เป็นยอดทั้งชั่วโมงจากไฟล์ (CSV เป็นรายชั่วโมง จึงแยกยอดถึง 22:42 น. ไม่ได้)</li>
<li>ข้อความ ความมั่นใจ คำหลักฐาน และยอดมีส่วนร่วมรายโพสต์มาจาก xlsx ไม่ได้มาจาก CSV</li></ul></section>
<section class="card"><h2>รายวัน</h2><table style="margin-top:10px"><thead><tr><th>วัน</th><th>ทั้งหมด</th><th>เชิงลบ</th><th>แคมเปญ</th></tr></thead><tbody>{daily}</tbody></table></section>
<section class="card"><h2>แยกแพลตฟอร์ม</h2><table style="margin-top:10px"><thead><tr><th>แพลตฟอร์ม</th>{''.join(f'<th>{s}</th>' for s in SENT_ORDER)}<th>รวม</th></tr></thead><tbody>{plat}</tbody></table></section>
<section class="card"><h2>เชิงลบแยกสาขา</h2><table style="margin-top:10px"><thead><tr><th>สาขา</th><th>เชิงลบ</th><th>ทั้งหมด</th><th>สัดส่วน</th></tr></thead><tbody>{br}</tbody></table></section>
</main>"""
    return page("ที่มาของตัวเลข", body, "data", "doc")


# ---------------------------------------------------------------- app.js
APP_JS = r"""/* ชบา Monitor ต้นแบบ — JavaScript ล้วน ไม่ใช้ไลบรารี */
(function () {
  'use strict';
  var page = document.body.dataset.page;
  var $ = function (s, r) { return (r || document).querySelector(s); };
  var $$ = function (s, r) { return Array.prototype.slice.call((r || document).querySelectorAll(s)); };
  var DELAY = 1200;
  var failNext = false;
  var qs = new URLSearchParams(location.search);
  var RCPT_LABEL = __RCPT_LABEL__;
  var STATUS_HTML = __STATUS_HTML__;
  var ESC_MIN = __ESC_MIN__;

  function typing(e) {
    var t = e.target;
    return t && ((t.tagName === 'INPUT' && !/^(checkbox|radio)$/.test(t.type)) || t.tagName === 'TEXTAREA' || t.tagName === 'SELECT');
  }
  function onDemoKey(cb) {
    document.addEventListener('keydown', function (e) {
      if (e.metaKey || e.ctrlKey || e.altKey || typing(e)) return;
      if (e.code === 'KeyE' || e.key === 'e' || e.key === 'E') { failNext = !failNext; cb(failNext); }
    });
  }
  function flag(on) {
    var f = $('#demo-flag');
    if (f) f.textContent = on ? 'โหมดทดลองใช้งาน: ครั้งถัดไปจะไม่สำเร็จ' : '';
  }
  function esc(s) { var d = document.createElement('div'); d.textContent = s; return d.innerHTML; }
  function store(k, v) { try { if (v === undefined) return JSON.parse(sessionStorage.getItem(k)); sessionStorage.setItem(k, JSON.stringify(v)); } catch (e) { return null; } }
  function unstore(k) { try { sessionStorage.removeItem(k); } catch (e) {} }
  var colors = { 'บวก': 'var(--pos)', 'ลบ': 'var(--neg)', 'ผสม': 'var(--mix)', 'เป็นกลาง': 'var(--neu)' };
  function badge(s) {
    var c = colors[s];
    var dot = c ? '<span class="dot" style="background:' + c + (s === 'เป็นกลาง' ? ';box-shadow:inset 0 0 0 1px var(--ink3)' : '') + '"></span>' : '';
    return '<span class="sent" data-sent="' + esc(s) + '">' + dot + esc(s) + '</span>';
  }

  /* ---------- F16 แผงโหมดทดลองใช้งาน (มือถือ) ---------- */
  var demoFail = $('#demo-fail'), demoHide = $('#demo-hide'), demoSize = $('#demo-size'), demoOut = $('#demo-out');
  function consumeFail() {
    failNext = false;
    if (demoFail) { demoFail.checked = false; store('demo-fail', false); demoOut.textContent = 'จำลองไม่สำเร็จไปแล้ว 1 ครั้ง · สวิตช์ปิดเอง'; }
  }
  if (demoFail) {
    if (qs.get('fail') === '1') store('demo-fail', true);
    demoFail.checked = !!store('demo-fail');
    failNext = demoFail.checked;
    var onFail = function () {
      failNext = demoFail.checked; store('demo-fail', failNext);
      demoOut.textContent = failNext ? 'การส่งครั้งถัดไปจะไม่สำเร็จ' : '';
    };
    demoFail.addEventListener('change', onFail);
    document.addEventListener('keydown', function (e) {
      if (e.metaKey || e.ctrlKey || e.altKey || typing(e)) return;
      if (e.code === 'KeyE' || e.key === 'e' || e.key === 'E') { demoFail.checked = !demoFail.checked; onFail(); }
    });
    if (failNext) demoOut.textContent = 'การส่งครั้งถัดไปจะไม่สำเร็จ';
  }
  function applyHide() {
    var on = !!store('demo-hide');
    var full = $('#push-full'), priv = $('#push-private');
    if (full && priv) { full.hidden = on; priv.hidden = !on; }
  }
  if (demoHide) {
    if (qs.get('hide') === '1') store('demo-hide', true);
    demoHide.checked = !!store('demo-hide');
    demoHide.addEventListener('change', function () { store('demo-hide', demoHide.checked); applyHide(); });
    applyHide();
  }
  function applySize() {
    var v = store('demo-size') || '390x844';
    var ph = $('.phone');
    if (!ph) return;
    ph.classList.toggle('s360', v === '360x740');
    ph.setAttribute('aria-label', 'หน้าจอมือถือ ' + v.replace('x', ' × '));
    if (demoSize) demoSize.value = v;
    if (window.fitM2) window.fitM2();
  }
  if (demoSize) {
    if (qs.get('size')) store('demo-size', qs.get('size'));
    demoSize.addEventListener('change', function () { store('demo-size', demoSize.value); applySize(); });
  }

  /* ---------- web: ชุดตัวกรองที่บันทึกไว้ ---------- */
  $$('.saved button.chip').forEach(function (b) {
    b.addEventListener('click', function () {
      var on = b.getAttribute('aria-pressed') !== 'true';
      $$('.saved button.chip').forEach(function (x) { x.setAttribute('aria-pressed', 'false'); });
      b.setAttribute('aria-pressed', on ? 'true' : 'false');
      var m = $('#saved-msg');
      if (b.dataset.branch !== undefined && $('#f-branch')) {
        $('#f-branch').value = on ? b.dataset.branch : '';
        $('#f-branch').dispatchEvent(new Event('change'));
        if (m) m.textContent = on ? 'ใช้ตัวกรองสาขาแล้ว (ประเด็นยังเป็นภาพนิ่ง)' : '';
        return;
      }
      if (m) m.textContent = on ? 'เลือกชุดตัวกรองแล้ว (ต้นแบบ: ชุดนี้ยังไม่เปลี่ยนตัวเลข)' : '';
    });
  });

  /* ---------- w1: F24 ตัวกรองสาขา · F28 tooltip + ตาราง ---------- */
  if (page === 'w1') {
    var sel = $('#f-branch'), tpls = $$('#branch-tpls template');
    var slots = { kpi: '#kpis', hourly: '#slot-hourly', plat: '#slot-plat', plattbl: '#slot-plattbl', camp: '#slot-camp' };
    var swap = function () {
      var b = sel.value;
      Object.keys(slots).forEach(function (part) {
        var t = tpls.filter(function (x) { return x.dataset.branch === b && x.dataset.part === part; })[0];
        if (t) $(slots[part]).innerHTML = t.innerHTML;
      });
      var label = b ? 'สาขา' + b : 'ทุกสาขา';
      $$('.scope').forEach(function (x) { x.textContent = label; });
      var tot = $('#kpis .kpi .v');
      $('#w1-live').textContent = 'แสดงข้อมูล ' + label + ' · ทั้งหมด ' + (tot ? tot.textContent : '') + ' ครั้ง';
    };
    sel.addEventListener('change', swap);
    var tip = $('#tip');
    var showTip = function (el) {
      var r = el.getBoundingClientRect();
      tip.textContent = el.dataset.tip; tip.hidden = false;
      var w = tip.offsetWidth;
      tip.style.left = Math.max(8, Math.min(window.innerWidth - w - 8, r.left + r.width / 2 - w / 2)) + 'px';
      tip.style.top = (r.top - 34) + 'px';
    };
    var hideTip = function () { tip.hidden = true; };
    ['mouseover', 'focusin'].forEach(function (ev) {
      document.addEventListener(ev, function (e) { var t = e.target.closest && e.target.closest('.seg-hit'); if (t) showTip(t); });
    });
    ['mouseout', 'focusout'].forEach(function (ev) {
      document.addEventListener(ev, function (e) { if (e.target.closest && e.target.closest('.seg-hit')) hideTip(); });
    });
    window.addEventListener('scroll', function () {
      var f = document.activeElement;
      if (f && f.classList && f.classList.contains('seg-hit')) showTip(f); else hideTip();
    }, { passive: true });
    var tb = $('#plat-tbl-btn');
    tb.addEventListener('click', function () {
      var open = tb.getAttribute('aria-expanded') !== 'true';
      tb.setAttribute('aria-expanded', open ? 'true' : 'false');
      $('#slot-plattbl').hidden = !open;
      tb.lastChild.textContent = open ? 'ซ่อนตารางข้อมูล' : 'ดูตารางข้อมูล';
    });
  }

  /* ---------- w2: แท็บรอตรวจสอบ · ?case= ?branch= · คำสั่งกลุ่ม ---------- */
  if (page === 'w2') {
    var low = $('#only-low');
    var FLT = { 'case': 'เคส', 'branch': 'สาขา' };
    var url = function () { var q = qs.toString(); return 'w2.html' + (q ? '?' + q : ''); };
    var chips = function () {
      $$('#fchips .fchip').forEach(function (c) { c.remove(); });
      Object.keys(FLT).forEach(function (k) {
        var v = qs.get(k);
        if (!v) return;
        var c = document.createElement('span');
        c.className = 'fchip';
        c.innerHTML = esc(FLT[k] + ': ' + v) + '<button type="button" aria-label="ลบตัวกรอง ' + esc(FLT[k] + ' ' + v) + '"><svg class="ic" aria-hidden="true" focusable="false" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M18 6 6 18M6 6l12 12"/></svg></button>';
        $('button', c).addEventListener('click', function () {
          qs.delete(k);
          history.replaceState(null, '', url() + location.hash);
          apply();
          $('#list-status').focus && $('#list-status').setAttribute('tabindex', '-1');
          $('#list-status').focus();
        });
        $('#fchips').appendChild(c);
      });
    };
    var apply = function () {
      var review = location.hash === '#review' || (low && low.checked);
      if (low) low.checked = review;
      var fc = qs.get('case'), fb = qs.get('branch');
      var n = 0;
      $$('tbody tr').forEach(function (tr) {
        var hide = (review && tr.dataset.review !== '1') || (fc && tr.dataset.case !== fc) || (fb && tr.dataset.branch !== fb);
        tr.classList.toggle('is-hidden', !!hide);
        if (hide) { var cb = $('.rowsel', tr); if (cb) cb.checked = false; }
        if (!hide) n++;
      });
      if (location.hash !== '#review') $('#tab-all').setAttribute('aria-current', 'page'); else $('#tab-all').removeAttribute('aria-current');
      var trv = $('#tab-review');
      if (location.hash === '#review') trv.setAttribute('aria-current', 'page'); else trv.removeAttribute('aria-current');
      var parts = [];
      if (review) parts.push('เฉพาะความมั่นใจต่ำกว่า 75%');
      if (fc) parts.push('เคส ' + fc);
      if (fb) parts.push('สาขา ' + fb);
      $('#list-status').textContent = 'แสดง ' + n + ' รายการ' + (parts.length ? ' · ' + parts.join(' · ') : '');
      chips(); bulk();
    };
    window.addEventListener('hashchange', apply);
    if (low) low.addEventListener('change', function () {
      if (low.checked) location.hash = 'review';
      else { history.replaceState(null, '', url()); apply(); }
    });
    /* F33 เลือกแถว + แถบคำสั่งกลุ่ม */
    var selAll = $('#sel-all');
    var picked = function () { return $$('tbody tr:not(.is-hidden) .rowsel:checked').map(function (c) { return c.closest('tr'); }); };
    var bulk = function () {
      var rows = picked(), vis = $$('tbody tr:not(.is-hidden) .rowsel');
      $('#bulk').hidden = rows.length === 0;
      $('#bulk-n').textContent = 'เลือก ' + rows.length + ' รายการ';
      selAll.checked = rows.length > 0 && rows.length === vis.length;
      selAll.indeterminate = rows.length > 0 && rows.length < vis.length;
    };
    selAll.addEventListener('change', function () {
      $$('tbody tr:not(.is-hidden) .rowsel').forEach(function (c) { c.checked = selAll.checked; });
      $('#bulk-msg').textContent = '';
      bulk();
    });
    $$('.rowsel').forEach(function (c) { c.addEventListener('change', function () { $('#bulk-msg').textContent = ''; bulk(); }); });
    $('#bulk-apply').addEventListener('click', function () {
      var o = $('#bulk-owner').value, st = $('#bulk-status').value, rows = picked();
      if (!o && !st) { $('#bulk-msg').textContent = 'เลือกผู้รับผิดชอบหรือสถานะก่อน'; return; }
      rows.forEach(function (tr) {
        if (o) $('.ownercell', tr).textContent = o;
        if (st) $('.statuscell', tr).innerHTML = STATUS_HTML[st];
      });
      $('#bulk-msg').textContent = 'ใช้กับ ' + rows.length + ' รายการแล้ว' + (o ? ' · มอบหมาย ' + o : '') + (st ? ' · สถานะ ' + st : '');
    });
    $('#bulk-edit').addEventListener('click', function () {
      $('#bulk-msg').textContent = 'แก้ผลพร้อมกัน ' + picked().length + ' รายการ: ต้นแบบนี้แก้พร้อมกันได้จริงที่หน้า C09 (ใช้ผลเดียวกันกับ C31, C32)';
    });
    $('#bulk-clear').addEventListener('click', function () {
      $$('.rowsel').forEach(function (c) { c.checked = false; });
      $('#bulk-msg').textContent = '';
      bulk(); selAll.focus();
    });
    /* ผลที่บันทึกจาก w3 · F35 ป้ายตรวจสอบแล้ว */
    var mark = function (id, sent) {
      var tr = $('tr[data-id="' + id + '"]');
      if (!tr) return;
      $('.sentcell', tr).innerHTML = badge(sent) + '<span class="meta">ตรวจทานโดยเจ้าหน้าที่</span>';
      $('.statuscell', tr).innerHTML = STATUS_HTML['ตรวจสอบแล้ว'];
    };
    var c09 = store('c09');
    if (c09) mark('C09', c09.sent);
    var sims = store('sims');
    if (sims) sims.ids.forEach(function (id) { mark(id, sims.sent); });
    apply();
  }

  /* ---------- w-cases: F22 เปิดรายละเอียดเคส ---------- */
  if (page === 'wcases') {
    var toggle = function (b, open) {
      b.setAttribute('aria-expanded', open ? 'true' : 'false');
      $('#' + b.getAttribute('aria-controls')).hidden = !open;
    };
    $$('.case-btn').forEach(function (b) {
      b.addEventListener('click', function () { toggle(b, b.getAttribute('aria-expanded') !== 'true'); });
    });
    var want = qs.get('case');
    if (want) { var cb = $('.case-btn[data-case="' + want + '"]'); if (cb) { toggle(cb, true); cb.focus(); } }
  }

  /* ---------- w2-excluded: คืนสู่รายการ ---------- */
  if (page === 'w2x') {
    $$('[data-restore]').forEach(function (b) {
      var row = b.closest('.ex-row'), msg = $('.ex-status', row), label = $('span', b), done = false;
      b.addEventListener('click', function () {
        done = !done;
        row.classList.toggle('restored', done);
        label.textContent = done ? 'เลิกทำ' : 'คืนสู่รายการ';
        msg.textContent = done ? 'คืนสู่รายการแล้ว · นับใน Dashboard รอบถัดไป' : 'เลิกทำแล้ว · ยังอยู่ในคัดออก';
      });
    });
  }

  /* ---------- w3: แผงแก้ผล ---------- */
  if (page === 'w3') {
    var form = $('#edit'), save = $('#save'), cancel = $('#cancel'), msg = $('#save-msg');
    var fields = $('#edit-fields'), sum = $('#saved-sum');
    var hist = $('#history'), learnStatus = $('#learn-status');
    var setStatus = $('#set-status');
    var original = { sent: val('sent'), urg: val('urg') };
    var saved = { sent: original.sent, urg: original.urg, reviewed: false, status: setStatus.value };
    var prev = null, busy = false, last = null;
    var origConf = $('#cur-conf').innerHTML;
    var urgHtml = { 'สูง': 'warn', 'กลาง': 'pri', 'ต่ำ': 'plain' };

    function val(n) { var c = $('input[name="' + n + '"]:checked', form); return c ? c.value : ''; }
    function setVal(n, v) { $$('input[name="' + n + '"]', form).forEach(function (i) { i.checked = (i.value === v); }); }
    function urgTag(u) { return '<span class="tag ' + (urgHtml[u] || 'plain') + '">' + esc(u) + '</span>'; }
    function refresh() {
      var d = $('#diff');
      d.innerHTML = '<span>ความรู้สึก: <span class="o">' + esc(saved.sent) + '</span> → <b>' + esc(val('sent')) + '</b></span>' +
        '<span>ความเร่งด่วน: <span class="o">' + esc(saved.urg) + '</span> → <b>' + esc(val('urg')) + '</b></span>';
      save.disabled = !val('reason') || busy;
      $('#save-hint').textContent = val('reason') ? 'พร้อมบันทึก · ยกเลิก = คืนค่าเดิม' : 'เลือกเหตุผลก่อน ปุ่มบันทึกจึงกดได้ · ยกเลิก = คืนค่าเดิม';
    }
    function render() {
      $('#cur-sent').innerHTML = badge(saved.sent);
      $('#cur-urg').innerHTML = urgTag(saved.urg);
      $('#cur-conf').innerHTML = saved.reviewed ? '<span class="tag ok">ตรวจทานโดยเจ้าหน้าที่</span>' : origConf;
      setStatus.value = saved.status;
      var ap = $('#apply-sim');
      ap.disabled = !saved.reviewed;
      $('#apply-sim-hint').textContent = saved.reviewed ? 'ใช้ความรู้สึก ' + saved.sent + ' · เหตุผล ' + (last ? last.reason : '') + ' กับ C31 และ C32' : 'บันทึกผลของ C09 ก่อน แล้วใช้ผลเดียวกันได้';
    }
    function addHist(text) {
      var li = document.createElement('li');
      li.className = 'new';
      li.innerHTML = '<time>จ. 26 ต.ค. 69 · 14:05 น.</time><span>' + text + '</span>';
      hist.appendChild(li);
    }
    function resetForm(keepReason) {
      setVal('sent', saved.sent); setVal('urg', saved.urg);
      if (keepReason && last) setVal('reason', last.reason);
      else $$('input[name="reason"]', form).forEach(function (i) { i.checked = false; });
      refresh();
    }
    function showMsg(kind, html) { msg.innerHTML = html ? '<div class="notice ' + kind + '">' + html + '</div>' : ''; }
    /* F40 F41 สรุปหลังบันทึก: ซ่อนฟอร์ม · แสดง ลบ → บวก · บันทึกแล้ว + เลิกทำ */
    function showSummary(from) {
      fields.hidden = true;
      showMsg('', '');
      sum.innerHTML = '<p class="chg">' + badge(from.sent) + '<span aria-hidden="true">→</span><span class="sr-only">เป็น</span>' + badge(saved.sent) + '<span class="tag ok">บันทึกแล้ว</span></p>' +
        '<dl><dt>ความเร่งด่วน</dt><dd>' + esc(from.urg) + ' → ' + esc(saved.urg) + '</dd><dt>เหตุผล</dt><dd>' + esc(last.reason) + '</dd><dt>สถานะ</dt><dd>' + STATUS_HTML['ตรวจสอบแล้ว'] + '</dd><dt>การเรียนรู้</dt><dd>' + esc(learnStatus.textContent) + '</dd></dl>' +
        '<div class="btnrow"><button class="btn" type="button" id="undo">เลิกทำ</button><button class="btn ghost" type="button" id="again">แก้อีกครั้ง</button></div>';
      sum.hidden = false;
      $('#undo').addEventListener('click', undo);
      $('#again').addEventListener('click', function () { sum.hidden = true; fields.hidden = false; resetForm(true); $('input[name="sent"]:checked', form).focus(); });
      sum.focus();
    }

    form.addEventListener('change', refresh);
    cancel.addEventListener('click', function () {
      resetForm();
      showMsg('info', '<p class="grow">ยกเลิกแล้ว · คืนค่าเดิม ' + esc(saved.sent) + ' / ความเร่งด่วน' + esc(saved.urg) + '</p>');
      if (saved.reviewed) { showSummary(prev || original); }
    });
    function doSave() {
      if (busy) return;
      busy = true;
      save.textContent = 'กำลังบันทึก…';
      save.setAttribute('aria-busy', 'true');
      save.disabled = true;
      showMsg('info', '<p class="grow">กำลังบันทึก…</p>');
      var next = { sent: val('sent'), urg: val('urg'), reason: val('reason'), learn: $('#learn').checked };
      setTimeout(function () {
        busy = false;
        save.textContent = 'บันทึก';
        save.removeAttribute('aria-busy');
        if (failNext) {
          failNext = false; flag(false);
          showMsg('err', '<p class="grow"><b>บันทึกไม่สำเร็จ</b> · การเชื่อมต่อขาด ยังไม่ได้บันทึกค่าใหม่ ร่างการแก้ไขยังอยู่ครบ</p><button class="btn sm" type="button" id="retry">ลองใหม่</button>');
          $('#retry').addEventListener('click', doSave);
          refresh();
          return;
        }
        prev = { sent: saved.sent, urg: saved.urg, reviewed: saved.reviewed, status: saved.status };
        last = next;
        saved = { sent: next.sent, urg: next.urg, reviewed: true, status: 'ตรวจสอบแล้ว' };
        learnStatus.textContent = next.learn ? 'อยู่ในคิวเรียนรู้รอบถัดไป (ยังไม่ได้เรียนรู้ทันที)' : 'ไม่ส่งให้ระบบเรียนรู้';
        render();
        addHist('แก้ความรู้สึก ' + esc(prev.sent) + ' → ' + esc(next.sent) + ' · ความเร่งด่วน ' + esc(prev.urg) + ' → ' + esc(next.urg) + ' · เหตุผล ' + esc(next.reason) + ' · สถานะ ตรวจสอบแล้ว · ผู้แก้ ทีม Customer Support');
        store('c09', { sent: saved.sent });
        showSummary(prev);
      }, DELAY);
    }
    function undo() {
      if (!prev) return;
      var was = saved.sent;
      saved = prev; prev = null;
      render();
      addHist('ย้อนการแก้ไข · คืนค่า ' + esc(was) + ' → ' + esc(saved.sent));
      learnStatus.textContent = 'ถอนออกจากคิวเรียนรู้แล้ว';
      unstore('c09');
      sum.hidden = true; fields.hidden = false;
      resetForm();
      showMsg('info', '<p class="grow">เลิกทำแล้ว · คืนค่าก่อนแก้ (' + esc(saved.sent) + ')</p>');
      $('input[name="sent"]:checked', form).focus();
    }
    form.addEventListener('submit', function (e) { e.preventDefault(); if (val('reason')) doSave(); });
    /* F37 ตั้งค่าโพสต์นี้ */
    $$('.setup select').forEach(function (s) {
      s.addEventListener('change', function () {
        if (s === setStatus) saved.status = s.value;
        $('#setup-msg').textContent = 'บันทึกแล้ว · ' + s.closest('label').firstChild.textContent.trim() + ' ' + s.value;
        addHist('เปลี่ยน' + esc(s.closest('label').firstChild.textContent.trim()) + 'เป็น ' + esc(s.value));
      });
    });
    /* F38 ใช้ผลเดียวกันกับ C31 C32 */
    var simDlg = $('#sim-dlg'), simMsg = $('#sim-msg'), simApplied = false;
    $('#apply-sim').addEventListener('click', function () {
      $('#sim-dlg-t').textContent = 'C31 และ C32 จะเปลี่ยนความรู้สึกจาก ลบ เป็น ' + saved.sent + ' · เหตุผล ' + last.reason + ' · สถานะเป็น ตรวจสอบแล้ว';
      simDlg.showModal();
    });
    function setSims(on) {
      $$('[data-sim]').forEach(function (li) {
        var res = $('.simres', li);
        if (on) {
          li.dataset.html = res.innerHTML;
          res.innerHTML = esc(li.dataset.plat) + ' · ' + badge(saved.sent) + ' <span class="tag ok">แก้แล้ว · เดิมระบบ ' + esc(li.dataset.orig) + '</span>';
        } else if (li.dataset.html) { res.innerHTML = li.dataset.html; }
      });
    }
    simDlg.addEventListener('close', function () {
      if (simDlg.returnValue !== 'ok') { $('#apply-sim').focus(); return; }
      simApplied = true; setSims(true);
      store('sims', { ids: ['C31', 'C32'], sent: saved.sent });
      addHist('ใช้ผลเดียวกันกับ C31, C32 · ความรู้สึก ' + esc(saved.sent) + ' · เหตุผล ' + esc(last.reason));
      simMsg.innerHTML = '<div class="notice ok" style="margin-top:10px"><p class="grow"><b>ใช้ผลกับ 2 รายการแล้ว</b> · C31, C32 เป็น ' + esc(saved.sent) + '</p><button class="btn sm" type="button" id="sim-undo">เลิกทำ</button></div>';
      $('#sim-undo').addEventListener('click', function () {
        if (!simApplied) return;
        simApplied = false; setSims(false); unstore('sims');
        addHist('ย้อนการใช้ผลกับ C31, C32');
        simMsg.innerHTML = '<div class="notice info" style="margin-top:10px"><p class="grow">เลิกทำแล้ว · C31, C32 กลับเป็นผลเดิมของระบบ</p></div>';
        $('#apply-sim').focus();
      });
      simDlg.returnValue = '';
      $('#sim-undo').focus();
    });
    onDemoKey(flag);
    refresh(); render();
  }

  /* ---------- m1: F03 ---------- */
  /* (สลับ push อยู่ใน applyHide ด้านบน) */

  /* ---------- m2 ---------- */
  if (page === 'm2') {
    var dlg = $('#dlg'), ok = $('#dlg-ok'), toast = $('#m2-toast');
    var other = $('#nu-other'), otherBox = $('#nu-other-box');
    var toastMsg = function (t, undo) { $('#m2-toast-t').textContent = t; $('#m2-undo').hidden = !undo; toast.hidden = false; };
    $('#not-urgent').addEventListener('click', function () { dlg.showModal(); });
    /* F20 ยืนยันได้เมื่อเลือกเหตุผล · อื่นๆ ต้องพิมพ์ */
    var check = function () {
      var r = $('input[name="nu"]:checked', dlg);
      var isOther = r && r.value === 'อื่นๆ';
      otherBox.hidden = !isOther;
      var good = r && (!isOther || other.value.trim());
      ok.disabled = !good;
      $('#dlg-hint').textContent = !r ? 'เลือกเหตุผลก่อน ปุ่มยืนยันจึงกดได้' : (isOther && !other.value.trim() ? 'พิมพ์เหตุผลก่อน ปุ่มยืนยันจึงกดได้' : 'พร้อมยืนยัน');
    };
    dlg.addEventListener('change', check);
    other.addEventListener('input', check);
    $$('input[name="nu"]', dlg).forEach(function (i) { i.addEventListener('change', function () { if (i.value === 'อื่นๆ') other.focus(); }); });
    dlg.addEventListener('close', function () {
      if (dlg.returnValue === 'ok') {
        var r = $('input[name="nu"]:checked', dlg).value;
        if (r === 'อื่นๆ') r = 'อื่นๆ: ' + other.value.trim();
        toastMsg('ปิดแจ้งเตือนแล้ว · ' + r, true);
      }
      $$('input[name="nu"]', dlg).forEach(function (i) { i.checked = false; });
      other.value = ''; check();
      dlg.returnValue = '';
    });
    $('#m2-undo').addEventListener('click', function () {
      toastMsg('เลิกทำแล้ว · เปิดแจ้งเตือนกลับมา', false);
      setTimeout(function () { toast.hidden = true; }, 2500);
    });
    /* F05 */
    $('#open-src').addEventListener('click', function () {
      toastMsg('ต้นแบบ: จะเปิดโพสต์ C18 บน Google Maps (ภาพนิ่ง)', false);
      setTimeout(function () { toast.hidden = true; }, 2500);
    });
    /* F08 รับเรื่อง */
    var acc = $('#accept-btn');
    var accepted = function () {
      $('#accept-t').textContent = 'รับเรื่องแล้วโดยคุณ · ' + acc.dataset.time + ' น.';
      $('#accept').classList.add('done');
      acc.hidden = true;
    };
    acc.addEventListener('click', function () { store('k10-accepted', true); accepted(); $('#accept-t').setAttribute('tabindex', '-1'); $('#accept-t').focus(); });
    if (store('k10-accepted')) accepted();
    /* หน้า m2 ต้องไม่เลื่อน: ถ้าล้น ย่อรูปหลักฐานก่อน (88 → 48 px) ไม่ตัดข้อมูล */
    /* ขั้น 1 ย่อรูป 88 → 56 px (เล็กกว่านี้อ่านป้ายในกรอบไม่ได้) · ขั้น 2 จัดระยะกระชับ (.tight) แล้วขยายรูปคืนเท่าที่พอ · ไม่ตัดข้อมูล */
    window.fitM2 = function () {
      var body = $('#m-body'), ph = $('.phone'), sz;
      var over = function () { return body.scrollHeight > body.clientHeight + 1; };
      var shrink = function () { sz = 88; body.style.setProperty('--evimg', sz + 'px'); while (over() && sz > 56) { sz -= 8; body.style.setProperty('--evimg', sz + 'px'); } };
      ph.classList.remove('tight');
      shrink();
      if (over()) { ph.classList.add('tight'); shrink(); }
      body.dataset.evimg = sz;
      body.dataset.tight = ph.classList.contains('tight') ? '1' : '0';
      body.dataset.overflow = Math.max(0, body.scrollHeight - body.clientHeight);
    };
  }
  applySize();
  if (window.fitM2 && document.fonts) document.fonts.ready.then(window.fitM2);

  /* ---------- m4: ตรวจสาขาและผู้รับ ---------- */
  if (page === 'm4') {
    var undone = store('k10-undone');
    if (undone) {
      var u = $('#undone');
      u.textContent = 'เลิกทำการส่งแล้ว · ผู้รับที่เปิดอ่านแล้วจะเห็นสถานะ "ถอน" · แก้ผู้รับแล้วส่งใหม่ได้';
      u.hidden = false;
      unstore('k10-undone');
    }
    var pick = $('#br-pick'), bsel = $('#br-sel'), unk = $('#br-unknown');
    var mgrName = $('input[value="mgr"]').closest('label').querySelector('.rc-name');
    var branch = function () {
      var v = $('input[name="br"]:checked').value;
      pick.hidden = v !== 'change';
      unk.hidden = v !== 'unknown';
      var b = v === 'change' ? bsel.value : (v === 'unknown' ? '' : 'รังสิต');
      var name = b ? 'ผู้จัดการสาขา' + b : 'ทีมกลาง / QA (รอยืนยันสาขา)';
      mgrName.textContent = name;
      $('input[value="mgr"]').dataset.name = name;
      return b;
    };
    /* F09 ผู้รับที่บังคับตามกฎนับรวมเสมอ */
    var chosen = function () { return $$('input[name="rc"]:checked, .rc-locked'); };
    var count = function () { $('#rc-count').textContent = chosen().length; };
    $$('input[name="br"]').forEach(function (i) { i.addEventListener('change', branch); });
    bsel.addEventListener('change', branch);
    $$('input[name="rc"]').forEach(function (i) { i.addEventListener('change', count); });
    $('#to-review').addEventListener('click', function () {
      var b = branch();
      var list = chosen().map(function (i) { return { name: i.dataset.name, role: i.dataset.role }; });
      var prevK = store('k10') || {};
      store('k10', { branch: b, list: list, msg: prevK.msg });
    });
    branch(); count();
  }

  function rcptLabel(r) { return r.name + (RCPT_LABEL[r.role] || ''); }

  /* ---------- m5: ยืนยันส่ง ---------- */
  if (page === 'm5') {
    var k = store('k10') || {};
    if (k.list) {
      $('#rv-list').innerHTML = k.list.map(function (r) { return '<li>' + esc(rcptLabel(r)) + '</li>'; }).join('');
      $('#rv-branch').textContent = k.branch ? 'สาขา' + k.branch : 'รอยืนยันสาขา';
      var main = k.list.filter(function (r) { return r.role === 'หลัก'; })[0];
      if (main) $('#call-name').textContent = main.name;
    }
    /* F14 แก้ไขข้อความ */
    var mt = $('#msg-text'), me = $('#msg-edit'), mi = $('#msg-input'), mb = $('#msg-edit-btn');
    if (k.msg) { mt.textContent = k.msg; mi.value = k.msg; }
    mb.addEventListener('click', function () { me.hidden = false; mt.hidden = true; mb.hidden = true; mi.value = mt.textContent; mi.focus(); });
    $('#msg-cancel').addEventListener('click', function () { me.hidden = true; mt.hidden = false; mb.hidden = false; mb.focus(); });
    $('#msg-save').addEventListener('click', function () {
      var v = mi.value.trim();
      if (!v) { mi.setAttribute('aria-invalid', 'true'); mi.placeholder = 'ข้อความว่างไม่ได้'; mi.focus(); return; }
      mi.removeAttribute('aria-invalid');
      mt.textContent = v; k.msg = v; store('k10', k);
      me.hidden = true; mt.hidden = false; mb.hidden = false; mb.textContent = 'แก้ไขข้อความ (แก้แล้ว)'; mb.focus();
    });
    var btn = $('#send'), lab = $('span', btn), out = $('#send-msg'), sending = false;
    var fd = $('#fail-dlg');
    var send = function () {
      if (sending) return;
      sending = true;
      btn.disabled = true; btn.setAttribute('aria-busy', 'true'); lab.textContent = 'กำลังส่ง…';
      out.innerHTML = '';
      setTimeout(function () {
        sending = false;
        if (failNext) {
          consumeFail();
          btn.disabled = false; btn.removeAttribute('aria-busy'); lab.textContent = 'ยืนยันส่ง';
          /* F21 pop-up บอกสาเหตุ + ลองใหม่ + โทรหาผู้รับหลัก */
          $('#call-msg').textContent = '';
          out.innerHTML = '<div class="notice err"><p class="grow"><b>ยังไม่ได้ส่ง</b> · ไม่มีสัญญาณอินเทอร์เน็ต · ร่างยังอยู่ครบ</p></div>';
          fd.showModal();
          return;
        }
        location.href = 'm6.html';
      }, DELAY);
    };
    btn.addEventListener('click', send);
    $('#retry').addEventListener('click', function () { fd.close(); send(); });
    $('#call').addEventListener('click', function () { $('#call-msg').textContent = 'กำลังโทรหา ' + $('#call-name').textContent + ' (จำลอง · ต้นแบบไม่โทรจริง)'; });
    $('#fail-close').addEventListener('click', function () { fd.close(); btn.focus(); });
  }

  /* ---------- m6: เลิกทำนับถอยหลัง ---------- */
  if (page === 'm6') {
    var k6 = store('k10');
    if (k6 && k6.list) {
      $('#ack').innerHTML = k6.list.map(function (r) {
        var tag = r.role === 'สำรอง' ? '<span class="tag plain">ผู้รับสำรอง · ส่งถ้าผู้รับหลักไม่รับทราบใน ' + ESC_MIN + ' นาที</span>' : '<span class="tag pri">ส่งถึงแล้ว · รอรับทราบ</span>';
        return '<li><svg class="ic" role="img" aria-label="ผู้รับ" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="8" r="4"/><path d="M4 21a8 8 0 0 1 16 0"/></svg><span>' + esc(r.name) + '</span>' + tag + '</li>';
      }).join('');
    }
    var sec = 5, el = $('#undo-sec'), ub = $('#undo');
    var t = setInterval(function () {
      sec -= 1;
      if (sec > 0) { el.textContent = sec; return; }
      clearInterval(t);
      $('#undo-text').textContent = 'หมดเวลาเลิกทำ (เวลาสาธิต) · ถอนหรือเปลี่ยนผู้รับได้จากปุ่มด้านล่าง';
      ub.hidden = true;
    }, 1000);
    ub.addEventListener('click', function () {
      clearInterval(t);
      store('k10-undone', true);
      location.href = 'm4.html';
    });
  }
})();
"""
APP_JS = (APP_JS.replace("__RCPT_LABEL__", json.dumps(RCPT_LABEL, ensure_ascii=False))
          .replace("__STATUS_HTML__", json.dumps({k: status_tag(k) for k in STATUS_TAG}, ensure_ascii=False))
          .replace("__ESC_MIN__", str(ESCALATE_MIN)))


# ---------------------------------------------------------------- write
def main():
    OUT.mkdir(exist_ok=True)
    (OUT / "tokens.css").write_text(CSS, encoding="utf-8")
    (OUT / "app.js").write_text(APP_JS, encoding="utf-8")
    files = {
        "index.html": build_index(), "w1.html": build_w1(), "w2.html": build_w2(),
        "w2-excluded.html": build_w2_excluded(), "w3.html": build_w3(), "w-cases.html": build_w_cases(), "w-states.html": build_w_states(),
        "m1.html": build_m1(), "m2.html": build_m2(), "m3.html": build_m3(), "m4.html": build_m4(),
        "m5.html": build_m5(), "m6.html": build_m6(), "m-states.html": build_m_states(),
        "rationale.html": build_rationale(), "foundations.html": build_foundations(), "data.html": build_data(),
    }
    for name, content in files.items():
        (OUT / name).write_text(content, encoding="utf-8")
    summary = {
        "total": S["total"], "neg": S["neg"], "peak": S["peak"], "peak_th": S["peak_th"],
        "review_ids": REVIEW_IDS, "empty_combo": EMPTY_COMBO,
        "contrast": contrast_rows(), "files": sorted(list(files) + ["tokens.css", "app.js"]),
    }
    (OUT / "build-summary.json").write_text(json.dumps(summary, ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"สร้าง {len(files)} หน้า + tokens.css + app.js ใน {OUT}")
    print(f"ทั้งหมด {fmt(S['total'])} · เชิงลบ {fmt(S['neg'])} · สูงสุด {S['peak_th']} = {S['peak']}")
    print("รอตรวจสอบ (<75%):", ", ".join(REVIEW_IDS), f"({len(REVIEW_IDS)} โพสต์)")


if __name__ == "__main__":
    main()
