#!/usr/bin/env python3
"""รอบ 2: จับภาพทุกหน้า → screenshots/round2/ · ภาพก่อน–หลัง → screenshots/compare/
ตรวจข้อความล้นกรอบ · m2 ไม่ต้องเลื่อนที่ 390×844 และ 360×740 · ลองกด 2 กรณีตั้งแต่ต้นจนจบ (รวมทางผิดพลาด)
    .venv/bin/python shots.py
"""
import sys
from pathlib import Path
from playwright.sync_api import sync_playwright

ROOT = Path(__file__).resolve().parent
P = ROOT / "prototype"
# R3 ภาพหลังปรับคำ → round3 · เทียบกับรอบ 1 ทุกภาพที่มีทั้งสองรอบ
OUT = ROOT / "screenshots" / "round3"
BEFORE = ROOT / "screenshots" / "round1"
CMP = ROOT / "screenshots" / "compare_r1_r3"
OUT.mkdir(parents=True, exist_ok=True)
CMP.mkdir(parents=True, exist_ok=True)
sys.path.insert(0, str(ROOT))
import build  # noqa: E402

MOBILE = {"m1", "m2", "m3", "m4", "m5", "m6"}
# หน้าที่แก้ในรอบ 2 และมีภาพรอบ 1 ให้เทียบ
CHANGED = ["w1", "w2", "w3", "w3-saved", "m1", "m2", "m2-dialog", "m3", "m4", "m5", "m5-error", "m6", "m-states", "rationale", "index"]

OVERFLOW_JS = r"""
() => {
  const bad = [];
  const boxSel = '.card,.mcard,.st,.kpi,.phone,.btn,.chip,.tag,.sent,.nav-item,.push,.notice,.undo-bar,.toast,.case,.rcpt label,.rcpt .locked,.m-tabs a,td,.spread div,dialog,.demo,.fchip,.bulk';
  const walker = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT);
  let n;
  while ((n = walker.nextNode())) {
    if (!n.textContent.trim()) continue;
    const el = n.parentElement;
    if (!el || el.closest('[hidden],.sr-only,.skip,svg,template')) continue;
    const cs = getComputedStyle(el);
    if (cs.visibility === 'hidden' || cs.display === 'none') continue;
    const box = el.closest(boxSel);
    if (!box) continue;
    const r = document.createRange(); r.selectNodeContents(n);
    const tr = r.getBoundingClientRect(), br = box.getBoundingClientRect();
    if (!tr.width) continue;
    if (tr.right > br.right + 1.5 || tr.left < br.left - 1.5) {
      bad.push({text: n.textContent.trim().slice(0, 40), box: box.className || box.tagName, over: Math.round(Math.max(tr.right - br.right, br.left - tr.left))});
    }
  }
  if (document.documentElement.scrollWidth > window.innerWidth + 1) bad.push({text: 'หน้าเลื่อนแนวนอน', box: 'html', over: document.documentElement.scrollWidth - window.innerWidth});
  return bad;
}
"""

SMALL_MOBILE_JS = """() => [...document.querySelectorAll('.phone a,.phone button,.phone input,.phone select,.phone textarea')]
  .filter(e => e.offsetParent && !e.closest('.push,dialog:not([open])')).map(e => {const r=e.closest('label')?e.closest('label').getBoundingClientRect():e.getBoundingClientRect(); return [e.textContent.trim().slice(0,20)||e.getAttribute('aria-label')||e.type, Math.round(r.width), Math.round(r.height)]})
  .filter(x => x[2] < 44 || x[1] < 24)"""
SMALL_WEB_JS = """() => [...document.querySelectorAll('a,button,input,select')]
  .filter(e => e.offsetParent && !e.closest('.crumbs,p,li,td,.prose,.chart-foot,.foot-note'))
  .map(e => {const r=(e.closest('label')||e).getBoundingClientRect(); return [e.textContent.trim().slice(0,20)||e.getAttribute('aria-label')||e.type, Math.round(r.width), Math.round(r.height)]})
  .filter(x => x[1] < 24 || x[2] < 24)"""


def main():
    problems, log = [], []
    with sync_playwright() as pw:
        br = pw.chromium.launch()
        ctx = br.new_context(viewport={"width": 1440, "height": 960}, device_scale_factor=1)
        pg = ctx.new_page()
        errors = []
        pg.on("pageerror", lambda e: errors.append(str(e)))

        def load(name, query=""):
            pg.goto((P / name).as_uri() + query)
            pg.wait_for_load_state("networkidle")
            pg.evaluate("document.fonts.ready")
            return pg.evaluate("document.fonts.check('16px \"IBM Plex Sans Thai\"')")

        def phone(name):
            pg.locator(".phone").screenshot(path=str(OUT / f"{name}.png"))

        def stage(name):
            pg.locator(".stage").screenshot(path=str(OUT / f"{name}.png"))

        def full(name):
            pg.evaluate("window.scrollTo(0, 0)"); pg.wait_for_timeout(100)
            pg.screenshot(path=str(OUT / f"{name}.png"), full_page=True)

        def check(cond, msg):
            if not cond:
                problems.append(msg)
            return cond

        # ---------- 1) ทุกหน้า ----------
        for f in sorted(P.glob("*.html")):
            stem = f.stem
            pg.evaluate("sessionStorage.clear()") if pg.url.startswith("file:") else None
            if not load(f.name):
                problems.append(f"{f.name}: ฟอนต์ IBM Plex Sans Thai ไม่โหลด")
            for b in pg.evaluate(OVERFLOW_JS):
                problems.append(f"{f.name}: ข้อความล้นกรอบ '{b['text']}' ({b['box']}) {b['over']}px")
            if stem in MOBILE:
                phone(stem)
                stage(f"{stem}-with-demo-panel") if stem in ("m1", "m5") else None
                for s in pg.evaluate(SMALL_MOBILE_JS):
                    problems.append(f"{f.name}: เป้าสัมผัสเล็กกว่า 44px '{s[0]}' {s[1]}×{s[2]}")
            else:
                full(stem)
            for s in pg.evaluate(SMALL_WEB_JS):
                problems.append(f"{f.name}: เป้าสัมผัสเล็กกว่า 24px '{s[0]}' {s[1]}×{s[2]}")

        # ---------- 2) m2 ต้องไม่เลื่อน ที่ 2 ขนาด ----------
        for size in ("390x844", "360x740"):
            load("m2.html", f"?size={size}")
            pg.wait_for_timeout(200)
            sh, ch, img, tight = pg.evaluate("(() => {const b=document.querySelector('#m-body'); return [b.scrollHeight, b.clientHeight, b.dataset.evimg, b.dataset.tight];})()")
            log.append(f"m2 {size}: เนื้อหา {sh}px / พื้นที่ {ch}px · รูปหลักฐาน {img}px · ระยะกระชับ {'ใช้' if tight == '1' else 'ไม่ใช้'}")
            check(sh <= ch + 1, f"m2 {size}: ต้องเลื่อน ({sh} > {ch}) แม้ย่อรูปถึง {img}px")
            for b in pg.evaluate(OVERFLOW_JS):
                problems.append(f"m2 {size}: ข้อความล้นกรอบ '{b['text']}' {b['over']}px")
            phone(f"m2-{size}")
        pg.evaluate("sessionStorage.clear()")

        # ---------- 3) กรณีที่ 1 (มือถือ) ตั้งแต่ต้นจนจบ ----------
        load("m1.html")
        check(pg.locator("#push-full .tag.crit svg").count() == 1, "F01: ป้ายวิกฤตใน push ไม่มีไอคอน")
        pg.locator("#demo-hide").check()                                   # F03
        check(pg.locator("#push-private").is_visible() and pg.locator("#push-full").is_hidden(), "F03: สลับ push แบบซ่อนไม่ได้")
        check("มีเหตุระดับวิกฤต 1 เรื่อง · แตะเพื่อดู" in pg.locator("#push-private").inner_text(), "F03: ข้อความ push แบบซ่อนไม่ตรง")
        stage("m1-private")
        pg.locator("#push-private").click(); pg.wait_for_url("**/m2.html")
        check(pg.locator("#demo-hide").is_checked(), "F03: สวิตช์ไม่จำค่าข้ามหน้า")
        pg.locator("#demo-hide").uncheck()
        check("ระบบประเมินว่าวิกฤต · มั่นใจ 96%" in pg.locator(".confrow").inner_text(), "F07")
        check(pg.locator(".ev-img").inner_text().strip() == "รูปจากโพสต์ต้นทาง" and pg.locator("#open-src").is_visible(), "F05")
        pg.locator("#accept-btn").click()                                  # F08
        check(pg.locator("#accept-t").inner_text() == "รับเรื่องแล้วโดยคุณ · 22:43 น.", "F08: " + pg.locator("#accept-t").inner_text())
        phone("m2-accepted")
        pg.locator("#not-urgent").click()                                  # F20
        check(pg.locator("#dlg-ok").is_disabled(), "F20: ยืนยันกดได้ก่อนเลือกเหตุผล")
        pg.locator("dialog label.radio:has-text('อื่นๆ')").click()
        check(pg.locator("#nu-other").is_visible() and pg.locator("#dlg-ok").is_disabled(), "F20: อื่นๆ ไม่มีช่องพิมพ์ หรือยืนยันได้โดยไม่พิมพ์")
        pg.locator("#nu-other").fill("ลูกค้าแจ้งผิดสาขา")
        check(pg.locator("#dlg-ok").is_enabled(), "F20: พิมพ์แล้วยังกดยืนยันไม่ได้")
        pg.screenshot(path=str(OUT / "m2-dialog.png"))
        pg.locator("#dlg-ok").click()
        pg.wait_for_selector("#m2-toast", state="visible", timeout=2000)
        check("อื่นๆ: ลูกค้าแจ้งผิดสาขา" in pg.locator("#m2-toast-t").inner_text(), "F20: เหตุผลที่พิมพ์ไม่แสดง")
        phone("m2-not-urgent")
        pg.locator("#m2-undo").click()
        pg.locator(".m-actions a:has-text('ดูภาพรวม')").click(); pg.wait_for_url("**/m3.html")
        check(pg.locator("a:has-text('เปิดบนเว็บ')").get_attribute("href") == "w2.html?case=K10", "F19")
        check([t.strip() for t in pg.locator(".m-tabs a span").all_inner_texts()] == ["Dashboard", "Social Feed", "Alerts & Cases"], "F22: ชื่อแท็บมือถือไม่ตรงเมนูเว็บ")
        pg.locator(".m-actions a").click(); pg.wait_for_url("**/m4.html")
        check(pg.locator("input[value=exec]").count() == 0 and "บังคับตามกฎ" in pg.locator(".rc-locked").inner_text(), "F09")
        check("ผู้รับสำรอง" in pg.locator("label:has(input[value=area])").inner_text(), "F10: m4")
        check("ถ้าผู้รับหลักไม่รับทราบใน 5 นาที ส่งต่อผู้รับสำรอง" in pg.locator("#m-body").inner_text() and "[รอยืนยัน]" not in pg.locator(".mcard[style*=dashed]").inner_text(), "F11")
        check(pg.locator("#rc-count").inner_text() == "4", "m4: จำนวนผู้รับ")
        pg.locator("#to-review").click(); pg.wait_for_url("**/m5.html")
        rv = pg.locator("#rv-list").inner_text()
        check("ผู้บริหาร (จำเป็น)" in rv and "ผู้บริหาร (ผู้บริหาร" not in rv, "F13: " + rv)
        check("ผู้รับสำรอง" in rv, "F10: m5")
        check("ยังไม่ยืนยันข้อเท็จจริง" in pg.locator(".msgbox").inner_text(), "F15")
        pg.locator("#msg-edit-btn").click()                                # F14
        check(pg.locator("#msg-input").is_visible(), "F14: ไม่เปิดช่องพิมพ์")
        pg.locator("#msg-input").fill("ลูกค้าแจ้งเจอชิ้นใสแข็งในลาเต้ที่สาขารังสิต โปรดตรวจสอบที่สาขาก่อนตอบลูกค้า")
        pg.locator("#msg-save").click()
        check("โปรดตรวจสอบที่สาขาก่อนตอบลูกค้า" in pg.locator("#msg-text").inner_text(), "F14: ข้อความที่แก้ไม่แสดง")
        phone("m5-edited")
        pg.locator("#demo-fail").check()                                   # F16 F21 ทางผิดพลาดจากแผงสาธิต
        stage("m5-with-demo-panel-on")
        pg.locator("#send").click()
        phone("m5-sending")
        pg.wait_for_selector("#fail-dlg[open]", timeout=3000)
        check("ไม่มีสัญญาณอินเทอร์เน็ต" in pg.locator("#fail-dlg").inner_text(), "F21: ไม่บอกสาเหตุ")
        check(not pg.locator("#demo-fail").is_checked(), "F16: สวิตช์ไม่ปิดเองหลังจำลอง")
        pg.locator("#call").click()
        check("กำลังโทรหา ผู้จัดการสาขารังสิต" in pg.locator("#call-msg").inner_text(), "F21: ปุ่มโทร")
        pg.screenshot(path=str(OUT / "m5-error.png"))
        pg.locator("#retry").click(); pg.wait_for_url("**/m6.html", timeout=4000)
        ack = pg.locator("#ack").inner_text()
        check("ผู้บริหาร" in ack and "ผู้รับสำรอง · ส่งถ้าผู้รับหลักไม่รับทราบใน 5 นาที" in ack, "F10: m6 " + ack)
        check("ผ่านมา 66 นาที นับจากโพสต์แรก 21:40 น." in pg.locator(".elapsed").inner_text().replace("\n", " "), "F18")
        phone("m6-countdown")
        pg.locator("#undo").click(); pg.wait_for_url("**/m4.html")
        check(pg.locator("#undone").is_visible(), "m6 เลิกทำ → m4")
        phone("m4-after-undo")
        log.append("กรณีที่ 1: m1 (push ซ่อน) → m2 (รับเรื่อง · ไม่ใช่เรื่องด่วน+อื่นๆ+เลิกทำ) → m3 → m4 → m5 (แก้ข้อความ · จำลองส่งไม่สำเร็จ → pop-up → โทร → ลองใหม่) → m6 → เลิกทำ → m4")

        # ---------- 4) กรณีที่ 2 (เว็บ) ตั้งแต่ต้นจนจบ ----------
        pg.evaluate("sessionStorage.clear()")
        load("w1.html")
        check(pg.locator(".tabs").count() == 0, "F23")
        order = pg.evaluate("[...document.querySelectorAll('.content > *')].map(e => e.className || e.tagName).slice(0, 3)")
        check(order[1] == "alertbar" and order[2] == "card", f"F25: ลำดับ {order}")
        bg = pg.evaluate("getComputedStyle(document.querySelector('.kpi.hot')).backgroundColor")
        check(bg == "rgb(255, 255, 255)", f"F26: {bg}")
        cov = pg.locator(".cov").inner_text()
        check(all(p in cov for p in build.S["platforms"]) and "ยังไม่ทราบข้อจำกัด" in cov, "F30")
        check(pg.locator("a:has-text('ระบุสาขา 310 รายการ')").get_attribute("href") == "w2.html?branch=ไม่ระบุ", "F29")
        pg.locator(".seg-hit").first.focus()                               # F28
        check(pg.locator("#tip").is_visible() and "Facebook · ลบ 235" in pg.locator("#tip").inner_text(), "F28: tooltip " + pg.locator("#tip").inner_text())
        pg.locator("#plat-tbl-btn").click()
        check(pg.locator("#slot-plattbl table").is_visible(), "F28: ตาราง")
        pg.locator(".seg-hit").nth(1).hover()
        full("w1-tooltip-table")
        pg.select_option("#f-branch", "รังสิต")                             # F24
        rs = build.ST_BRANCH["รังสิต"]
        kp = pg.locator("#kpis").inner_text()
        check(build.fmt(rs["total"]) in kp and str(rs["peak"]) + " ครั้ง" in kp, f"F24: KPI {kp[:80]}")
        check("สาขารังสิต" in pg.locator("#slot-hourly svg").get_attribute("aria-label"), "F24: กราฟรายชั่วโมงไม่เปลี่ยน")
        full("w1-rangsit")
        log.append(f"w1 สาขารังสิต: ทั้งหมด {build.fmt(rs['total'])} · ลบ {rs['neg']} · สูงสุด {rs['peak_th']} = {rs['peak']}")
        pg.select_option("#f-branch", "")
        pg.locator(".alertbar a:has-text('K10')").click(); pg.wait_for_url("**/w2.html?case=K10")   # F25
        n = pg.locator("tbody tr:not(.is-hidden)").count()
        k10 = len([p for p in build.case_posts("K10") if not build.excluded(p)])
        check(n == k10, f"?case=K10 แสดง {n} ไม่ใช่ {k10}")
        check("19–25 ต.ค. 69" in pg.locator("#fchips").inner_text(), "F31")
        check(pg.locator("thead th").count() == 10 and "ความเร่งด่วน" in pg.locator("thead").inner_text(), "F32 F33: หัวคอลัมน์")
        full("w2-case-K10")
        pg.locator(".fchip button").click()
        check(pg.locator("tbody tr:not(.is-hidden)").count() == 34, "ลบชิปตัวกรองไม่ได้")
        check(pg.locator("#bulk").is_hidden(), "F33: แถบคำสั่งขึ้นก่อนเลือก")
        pg.locator("tr[data-id=C34] .rowsel").check(); pg.locator("tr[data-id=C35] .rowsel").check()
        check(pg.locator("#bulk").is_visible() and "เลือก 2 รายการ" in pg.locator("#bulk-n").inner_text(), "F33: แถบคำสั่งไม่ขึ้น")
        pg.select_option("#bulk-status", "ตรวจสอบแล้ว"); pg.locator("#bulk-apply").click()
        check("ตรวจสอบแล้ว" in pg.locator("tr[data-id=C34] .statuscell").inner_text(), "F35")
        full("w2-bulk")
        pg.goto((P / "w2.html").as_uri() + "?branch=ไม่ระบุ"); pg.wait_for_load_state("networkidle")
        nb = pg.locator("tbody tr:not(.is-hidden)").count()
        log.append(f"w2?branch=ไม่ระบุ: {nb} โพสต์ (xlsx) · w2?case=K10: {k10} โพสต์")
        full("w2-branch-unknown")
        pg.goto((P / "w2.html").as_uri() + "?case=K14"); pg.wait_for_load_state("networkidle")
        check(pg.locator("tbody tr:not(.is-hidden)").count() == 1, "?case=K14")
        pg.locator("a.postid:has-text('C09')").count() or pg.goto((P / "w2.html").as_uri())
        pg.goto((P / "w2.html").as_uri()); pg.locator("a.postid:has-text('C09')").click(); pg.wait_for_url("**/w3.html")
        check(pg.locator("#setup-h").inner_text().strip() == "ตั้งค่าโพสต์นี้", "F37")
        check(pg.locator("#apply-sim").is_disabled(), "F38: กดได้ก่อนบันทึก")
        btn_y = pg.evaluate("[document.querySelector('#save').getBoundingClientRect().top, document.querySelector('#cancel').getBoundingClientRect().top]")
        check(abs(btn_y[0] - btn_y[1]) < 2, "F42: ปุ่มไม่อยู่แถวเดียว")
        pg.locator("label:has(input[name=sent][value='บวก'])").click()
        pg.locator("label:has(input[name=reason][value='ขบขัน'])").click()
        pg.keyboard.press("e")
        pg.locator("#save").click()
        pg.screenshot(path=str(OUT / "w3-saving.png"))
        pg.wait_for_selector("#retry", timeout=3000)
        check("ลบ" in pg.locator("#cur-sent").inner_text(), "บันทึกไม่สำเร็จต้องไม่เปลี่ยนค่า")
        full("w3-error")
        pg.locator("#retry").click()
        pg.wait_for_selector("#saved-sum:not([hidden])", timeout=3000)
        sm = pg.locator("#saved-sum").inner_text().replace("\n", " ")
        check("ลบ" in sm and "บวก" in sm and "บันทึกแล้ว" in sm, "F40: " + sm)
        check(pg.locator("#edit-fields").is_hidden() and not pg.locator("#save").is_visible() and "เลือกเหตุผลก่อน" not in pg.locator("#edit").inner_text(), "F41")
        check(pg.locator("#set-status").input_value() == "ตรวจสอบแล้ว", "F35: w3")
        full("w3-saved")
        pg.locator("#apply-sim").click()                                    # F38
        check(pg.locator("#sim-dlg").is_visible(), "F38: ไม่มีกล่องยืนยัน")
        pg.screenshot(path=str(OUT / "w3-apply-sim-dialog.png"))
        pg.locator("#sim-dlg button[value=ok]").click()
        pg.wait_for_selector("#sim-undo", timeout=2000)
        check(pg.locator("[data-sim] .tag.ok").count() == 2, "F38: ไม่เปลี่ยน C31 C32")
        full("w3-apply-sim-done")
        pg.locator("#sim-undo").click()
        check(pg.locator("[data-sim] .tag.ok").count() == 0, "F38: เลิกทำไม่ได้")
        pg.locator("#undo").click()
        check("ลบ" in pg.locator("#cur-sent").inner_text() and pg.locator("#edit-fields").is_visible(), "w3 เลิกทำ")
        full("w3-undone")
        pg.locator("label:has(input[name=urg][value='กลาง'])").click()
        pg.locator("#cancel").click()
        check(pg.locator("input[name=urg][value='ต่ำ']").is_checked(), "ยกเลิกต้องคืนค่าเดิม")
        # บันทึกอีกครั้งแล้วดูใน w2 + w-cases
        pg.locator("label:has(input[name=reason][value='ขบขัน'])").click(); pg.locator("label:has(input[name=sent][value='บวก'])").click()
        pg.locator("#save").click(); pg.wait_for_selector("#saved-sum:not([hidden])", timeout=3000)
        pg.locator(".nav-item:has-text('Social Feed')").click(); pg.wait_for_url("**/w2.html")
        check("ตรวจสอบแล้ว" in pg.locator("tr[data-id=C09] .statuscell").inner_text(), "F35: w2 ไม่เห็นผลจาก w3")
        pg.locator(".nav-item:has-text('Alerts & Cases')").click(); pg.wait_for_url("**/w-cases.html")
        pg.locator(".case-btn[data-case=K10]").click()
        check(pg.locator("#cd-K10").is_visible(), "F22: เปิดรายละเอียด K10 ไม่ได้")
        full("w-cases-K10-open")
        log.append("กรณีที่ 2: w1 (tooltip · ตาราง · สาขารังสิต) → K10 → w2?case=K10 (ลบชิป · เลือก 2 แถว · เปลี่ยนสถานะ) → C09 → w3 (E ไม่สำเร็จ → ลองใหม่ → สรุป → ใช้กับ C31 C32 → เลิกทำ ×2 → ยกเลิก) → w2 → w-cases K10")

        # R3 สถานะที่รอบ 1 มีภาพ ใช้เทียบก่อน–หลังให้ครบทุกภาพ
        load("w2.html", "#review"); pg.wait_for_timeout(300)
        check("รอตรวจสอบ" in pg.locator("#tab-review").inner_text(), "R3: แท็บรอตรวจสอบ")
        full("w2-review")
        load("m6.html"); pg.wait_for_timeout(6500)  # เลยช่วงเลิกทำ 5 วินาที (เวลาสาธิต)
        phone("m6-expired")

        # ---------- 5) ภาพก่อน–หลัง ----------
        made = 0
        for stem in sorted(x.stem for x in BEFORE.glob("*.png")):
            a, b = BEFORE / f"{stem}.png", OUT / f"{stem}.png"
            if not (a.exists() and b.exists()):
                log.append(f"ก่อน–หลัง: ข้าม {stem} (ไม่มีภาพ {'รอบ 1' if not a.exists() else 'รอบล่าสุด'})")
                continue
            html = (f'<html><body style="margin:0;font:600 20px sans-serif;background:#fff">'
                    f'<div style="display:flex;gap:24px;padding:16px;align-items:flex-start">'
                    f'<figure style="margin:0"><figcaption>รอบ 1 · {stem}</figcaption><img src="{a.as_uri()}" style="max-width:1440px;border:1px solid #ccc"></figure>'
                    f'<figure style="margin:0"><figcaption>รอบล่าสุด (ปรับคำ) · {stem}</figcaption><img src="{b.as_uri()}" style="max-width:1440px;border:1px solid #ccc"></figure>'
                    f'</div></body></html>')
            tmp = CMP / "_tmp.html"
            tmp.write_text(html, encoding="utf-8")
            cp = ctx.new_page()
            cp.set_viewport_size({"width": 3000 if stem not in MOBILE and not stem.startswith("m") else 1000, "height": 800})
            cp.goto(tmp.as_uri()); cp.wait_for_load_state("load")
            cp.screenshot(path=str(CMP / f"{stem}.png"), full_page=True)
            cp.close()
            made += 1
        (CMP / "_tmp.html").unlink(missing_ok=True)
        log.append(f"ก่อน–หลัง: {made} ภาพใน screenshots/compare_r1_r3/")

        if errors:
            problems += [f"JS error: {e}" for e in errors]
        br.close()

    for l in log:
        print(l)
    print(f"ภาพหน้าจอ {len(list(OUT.glob('*.png')))} ไฟล์ใน screenshots/round3/")
    if problems:
        print(f"พบปัญหา {len(problems)} ข้อ")
        for p in problems:
            print(" -", p)
        sys.exit(1)
    print("ผ่านทุกข้อ")


if __name__ == "__main__":
    main()
