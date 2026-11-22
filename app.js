/* Coffee Cafe Monitor ต้นแบบ — JavaScript ล้วน ไม่ใช้ไลบรารี */
(function () {
  'use strict';
  var page = document.body.dataset.page;
  var $ = function (s, r) { return (r || document).querySelector(s); };
  var $$ = function (s, r) { return Array.prototype.slice.call((r || document).querySelectorAll(s)); };
  var DELAY = 1200;
  var failNext = false;
  var qs = new URLSearchParams(location.search);
  var RCPT_LABEL = {"หลัก": " (ผู้รับหลัก)", "สำรอง": " (ผู้รับสำรอง · ส่งเมื่อผู้รับหลักไม่รับทราบใน 5 นาที)", "จำเป็น": " (จำเป็น)"};
  var STATUS_HTML = {"ใหม่": "<span class=\"tag plain\">ใหม่</span>", "อยู่ระหว่างดำเนินการ": "<span class=\"tag pri\">อยู่ระหว่างดำเนินการ</span>", "ตรวจสอบแล้ว": "<span class=\"tag ok\"><svg class=\"ic\" aria-hidden=\"true\" focusable=\"false\" width=\"13\" height=\"13\" viewBox=\"0 0 24 24\" fill=\"none\" stroke=\"currentColor\" stroke-width=\"2\" stroke-linecap=\"round\" stroke-linejoin=\"round\"><path d=\"M20 6 9 17l-5-5\"/></svg>ตรวจสอบแล้ว</span>"};
  var ESC_MIN = 5;

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
