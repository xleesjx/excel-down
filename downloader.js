/*!
 * 통합예약 엑셀 일괄 받기
 * 울산광역시교육청 통합예약 관리자 > 통합예약 목록 화면에서
 * '접수중'인 수업의 신청자 엑셀을 한 번에 내려받는다.
 *
 * 동작 방식: 각 수업의 [신청 관리] 버튼에 들어 있는 번호(opViewApplyList)를 읽고,
 * 신청자 목록 화면의 [엑셀출력](opExcel)이 보내는 것과 같은 요청을 보낸다.
 *  - 기본(1.3.0~): 받은 엑셀을 파일로 저장하지 않고 점검 웹앱 창으로 바로 보낸다 (postMessage)
 *  - 폴더 저장: 처음 한 번 고른 폴더(예: C:\통합예약명단)에 서버가 주는 파일 이름으로 직접 저장
 *          (같은 이름이 있으면 "이름 (1).xls"처럼 번호를 붙인다)
 *  - 예비: 크롬 다운로드 폴더로 받기 (숨은 iframe으로 [엑셀출력]과 똑같이 요청)
 */
(function () {
  'use strict';

  var VERSION = '1.3.0';
  var NAME = '통합예약 엑셀 일괄 받기';
  var LIST_PAGE = 'BD_selectReservationMngList.do';
  var EXCEL_URL = '/component/excel/download/ND_excelDownload.do';
  var EXCEL_KEY = 'reservationNoContsCceAplyList';
  var STATUSES = ['접수중', '접수마감']; // 이 상태 중 운영이 끝나지 않은 수업이 대상
  var FOLDER_HINT = 'C:\\통합예약명단';
  var ROW_PER_PAGE = '500'; // 엑셀에 인원이 잘리지 않도록 넉넉하게
  var GAP_MS = 1500; // 요청 사이 간격
  var IFRAME_GAP_MS = 2500; // 예비 방식에서 다운로드 사이 간격
  var ERROR_WAIT_MS = 6000; // 예비 방식에서 오류 화면이 돌아오는지 기다리는 시간
  // 점검 웹앱 (명단을 보내는 곳). 이 주소의 창에만 보낸다.
  // (시험할 때만 내 PC의 시험용 주소로 바꿀 수 있다)
  var WEBAPP_URL = /^http:\/\/localhost(:\d+)?\//.test(window.__excelDownWebapp || '') ? window.__excelDownWebapp : 'https://xleesjx.github.io/weekend-checker/';
  var WEBAPP_ORIGIN = new URL(WEBAPP_URL).origin;
  var WEBAPP_WINDOW = 'weekendChecker';
  var READY_WAIT_MS = 5 * 60 * 1000; // 웹앱 로그인을 기다리는 최대 시간

  if (window.__excelDownBusy) {
    alert(NAME + '\n\n이미 실행 중입니다. 잠시만 기다려 주세요.');
    return;
  }

  // ---------- 화면 확인 ----------
  if (location.pathname.indexOf(LIST_PAGE) === -1) {
    alert(
      NAME + '\n\n통합예약 목록 화면에서 눌러 주세요.\n\n' +
      '관리자 페이지 > 통합예약 관리 > 통합예약\n' +
      '(수업 목록과 [신청 관리] 버튼이 보이는 화면)'
    );
    return;
  }

  // ---------- 공통 ----------
  function clean(s) {
    return (s || '').replace(/\s+/g, ' ').trim();
  }

  function esc(s) {
    return String(s).replace(/[&<>"']/g, function (ch) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[ch];
    });
  }

  function sleep(ms) {
    return new Promise(function (r) { setTimeout(r, ms); });
  }

  // ---------- 목록 읽기 ----------
  var TODAY = (function () {
    var d = new Date();
    return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
  })();

  function headerIndex(table, label) {
    var ths = table.querySelectorAll('thead th');
    for (var i = 0; i < ths.length; i++) {
      if (clean(ths[i].textContent).indexOf(label) === 0) return i;
    }
    return -1;
  }

  function readClasses() {
    var found = [];
    var buttons = document.querySelectorAll('[onclick*="opViewApplyList"]');
    for (var i = 0; i < buttons.length; i++) {
      var btn = buttons[i];
      var m = /opViewApplyList\s*\(\s*'([^']*)'\s*,\s*'([^']*)'\s*,\s*'([^']*)'/.exec(
        btn.getAttribute('onclick') || ''
      );
      var tr = btn.closest('tr');
      if (!m || !tr) continue;

      var cells = tr.querySelectorAll('td');
      var table = tr.closest('table');
      var statusIdx = table ? headerIndex(table, '접수상태') : -1;
      var titleIdx = table ? headerIndex(table, '신청명') : -1;
      var dateIdx = table ? headerIndex(table, '운영일시') : -1;
      var countIdx = table ? headerIndex(table, '신청현황') : -1;

      var status = '';
      if (statusIdx >= 0 && cells[statusIdx]) {
        status = clean(cells[statusIdx].textContent);
      } else {
        for (var c = 0; c < cells.length; c++) {
          if (STATUSES.indexOf(clean(cells[c].textContent)) >= 0) { status = clean(cells[c].textContent); break; }
        }
      }
      if (STATUSES.indexOf(status) < 0) continue;

      // 운영일시 "2026-11-07 ~ 2026-12-05 09:20 ~ 12:20" → 시작일, 종료일
      var dateText = dateIdx >= 0 && cells[dateIdx] ? clean(cells[dateIdx].innerText) : '';
      var days = dateText.match(/\d{4}-\d{2}-\d{2}/g) || [];
      var startDay = days[0] || '';
      var endDay = days[days.length - 1] || '';
      if (endDay && endDay < TODAY) continue; // 이미 끝난 수업
      if (!endDay && status !== '접수중') continue; // 날짜를 못 읽으면 예전처럼 접수중만

      var title = '';
      if (titleIdx >= 0 && cells[titleIdx]) title = clean(cells[titleIdx].textContent);
      if (!title) {
        var a = tr.querySelector('a');
        title = a ? clean(a.textContent) : '수업 ' + m[1];
      }

      found.push({
        rsvtRegSn: m[1],
        rType: m[2],
        siteSn: m[3],
        title: title,
        status: status,
        startDay: startDay,
        endDay: endDay,
        date: dateText,
        count: countIdx >= 0 && cells[countIdx] ? clean(cells[countIdx].innerText) : ''
      });
    }
    // 기본 체크: 가장 가까운 수업일(오늘 이후 시작일 중 가장 빠른 날)에 운영하는 수업만
    var next = '';
    found.forEach(function (f) {
      if (f.startDay && f.startDay >= TODAY && (!next || f.startDay < next)) next = f.startDay;
    });
    if (!next) next = TODAY;
    found.forEach(function (f) {
      f.checked = !f.startDay || (f.startDay <= next && next <= f.endDay);
    });
    return found;
  }

  // ---------- 엑셀 요청 값 (opExcel과 같은 값) ----------
  function excelParams(item) {
    return {
      excelKey: EXCEL_KEY,
      q_rowPerPage: ROW_PER_PAGE,
      q_currPage: '1',
      q_sortName: '',
      q_sortOrder: '',
      q_searchSiteSn: item.siteSn,
      q_aplyRegSn: '',
      q_pagePerPage: '10',
      q_pagingEndNum: '10',
      q_rType: item.rType,
      q_rsvtRegSn: item.rsvtRegSn,
      q_pagingStartNum: '0',
      q_siteSn: item.siteSn,
      q_searchStartDate: '',
      q_searchEndDate: '',
      q_searchContsAprvSeCd: '',
      q_searchVal: ''
    };
  }

  // ---------- 저장 폴더 (처음 한 번 고르고 기억) ----------
  var canPickFolder = typeof window.showDirectoryPicker === 'function' && window.isSecureContext;

  function idb() {
    return new Promise(function (resolve, reject) {
      var req = indexedDB.open('excelDown', 1);
      req.onupgradeneeded = function () { req.result.createObjectStore('kv'); };
      req.onsuccess = function () { resolve(req.result); };
      req.onerror = function () { reject(req.error); };
    });
  }
  function idbGet(key) {
    return idb().then(function (db) {
      return new Promise(function (resolve) {
        var r = db.transaction('kv').objectStore('kv').get(key);
        r.onsuccess = function () { resolve(r.result || null); };
        r.onerror = function () { resolve(null); };
      });
    }).catch(function () { return null; });
  }
  function idbSet(key, val) {
    return idb().then(function (db) {
      return new Promise(function (resolve) {
        var tx = db.transaction('kv', 'readwrite');
        tx.objectStore('kv').put(val, key);
        tx.oncomplete = function () { resolve(true); };
        tx.onerror = function () { resolve(false); };
      });
    }).catch(function () { return false; });
  }

  function pickFolder() {
    return window.showDirectoryPicker({ id: 'excelDown', mode: 'readwrite', startIn: 'documents' })
      .then(function (h) { return idbSet('dir', h).then(function () { return h; }); });
  }

  // 반드시 버튼 클릭 직후에 부른다 (크롬이 허용 창을 띄우려면 클릭이 필요)
  function ensureFolder(handle) {
    if (!handle) return pickFolder();
    return handle.queryPermission({ mode: 'readwrite' }).then(function (p) {
      if (p === 'granted') return handle;
      return handle.requestPermission({ mode: 'readwrite' }).then(function (p2) {
        if (p2 === 'granted') return handle;
        throw new Error('폴더 저장을 허용하지 않았습니다.');
      });
    });
  }

  // ---------- 파일 이름 (서버가 주는 이름 그대로) ----------
  function bytesToText(s) {
    // fetch가 돌려주는 헤더 값은 바이트를 한 글자씩 담고 있다 → UTF-8, 안 되면 EUC-KR로 읽기
    if (!/[\u0080-\u00ff]/.test(s)) return s;
    var bytes = new Uint8Array(s.length);
    for (var i = 0; i < s.length; i++) bytes[i] = s.charCodeAt(i) & 0xff;
    try { return new TextDecoder('utf-8', { fatal: true }).decode(bytes); } catch (e) {}
    try { return new TextDecoder('euc-kr').decode(bytes); } catch (e) {}
    return s;
  }

  function fileNameFrom(disposition) {
    if (!disposition) return '';
    var m = /filename\*\s*=\s*([^']*)'[^']*'([^;]+)/i.exec(disposition);
    if (m) {
      try { return decodeURIComponent(m[2].trim().replace(/^"|"$/g, '')); } catch (e) {}
    }
    m = /filename\s*=\s*("([^"]*)"|[^;]+)/i.exec(disposition);
    if (!m) return '';
    var name = (m[2] !== undefined ? m[2] : m[1]).trim();
    if (/%[0-9a-f]{2}/i.test(name)) {
      try { name = decodeURIComponent(name); } catch (e) {}
    }
    return bytesToText(name);
  }

  function fallbackName() {
    var d = new Date();
    function p(n, w) { return String(n).padStart(w || 2, '0'); }
    return '통합예약신청자목록 - ' + d.getFullYear() + '-' + p(d.getMonth() + 1) + '-' + p(d.getDate()) +
      'T' + p(d.getHours()) + p(d.getMinutes()) + p(d.getSeconds()) + '.' + p(d.getMilliseconds(), 3) + '.xls';
  }

  function safeName(name) {
    return name.replace(/[\\/:*?"<>|\u0000-\u001f]/g, '_').replace(/^\.+/, '').trim() || fallbackName();
  }

  function exists(dir, name) {
    return dir.getFileHandle(name).then(function () { return true; }, function () { return false; });
  }

  // 같은 이름이 있으면 "이름 (1).xls", "이름 (2).xls" …
  function uniqueName(dir, name) {
    var dot = name.lastIndexOf('.');
    var stem = dot > 0 ? name.slice(0, dot) : name;
    var ext = dot > 0 ? name.slice(dot) : '';
    function attempt(n) {
      var candidate = n === 0 ? name : stem + ' (' + n + ')' + ext;
      return exists(dir, candidate).then(function (taken) {
        return taken && n < 500 ? attempt(n + 1) : candidate;
      });
    }
    return attempt(0);
  }

  // ---------- 엑셀 받기 (폴더 저장·웹앱 보내기 공통) → { name, blob } ----------
  function fetchExcel(item) {
    var body = new URLSearchParams(excelParams(item));
    return fetch(EXCEL_URL, { method: 'POST', body: body, credentials: 'same-origin' })
      .then(function (res) {
        var type = res.headers.get('content-type') || '';
        var disp = res.headers.get('content-disposition') || '';
        if (!res.ok) throw new Error('서버 오류 ' + res.status);
        if (/text\/html/i.test(type) && !/attachment/i.test(disp)) {
          return res.text().then(function (html) {
            var t = clean(new DOMParser().parseFromString(html, 'text/html').body.innerText).slice(0, 120);
            throw new Error(t || '엑셀 대신 화면이 돌아왔습니다 (로그인이 끊겼을 수 있음)');
          });
        }
        return res.blob().then(function (blob) {
          if (!blob.size) throw new Error('빈 파일이 돌아왔습니다');
          return { name: safeName(fileNameFrom(disp) || fallbackName()), blob: blob };
        });
      });
  }

  // ---------- 폴더에 직접 저장 ----------
  function saveToFolder(dir, item) {
    return fetchExcel(item).then(function (got) {
      return uniqueName(dir, got.name).then(function (finalName) {
        return dir.getFileHandle(finalName, { create: true }).then(function (fh) {
          return fh.createWritable().then(function (w) {
            return w.write(got.blob).then(function () { return w.close(); });
          });
        }).then(function () { return finalName; });
      });
    });
  }

  // ---------- 점검 웹앱으로 바로 보내기 ----------
  // 웹앱이 로그인하고 명단을 다 불러오면 'wc-ready'로 답한다. 그때까지 1초마다 물어본다.
  function waitMessage(win, type, ms, onTick) {
    return new Promise(function (resolve, reject) {
      var t0 = Date.now();
      function onMsg(e) {
        if (e.origin !== WEBAPP_ORIGIN || e.source !== win || !e.data || e.data.type !== type) return;
        cleanup(); resolve(e.data);
      }
      var timer = setInterval(function () {
        if (win.closed) { cleanup(); reject(new Error('점검 웹앱 창이 닫혔습니다.')); return; }
        if (Date.now() - t0 > ms) { cleanup(); reject(new Error('점검 웹앱에서 답이 없습니다.')); return; }
        if (onTick) onTick();
      }, 1000);
      function cleanup() { clearInterval(timer); window.removeEventListener('message', onMsg); }
      window.addEventListener('message', onMsg);
      if (onTick) onTick();
    });
  }

  // 웹앱 결과(반영·건너뜀)는 웹앱이 보낸 것을 그대로 보여 준다
  function waitDone(win) {
    return new Promise(function (resolve) {
      function onMsg(e) {
        if (e.origin !== WEBAPP_ORIGIN || e.source !== win || !e.data) return;
        if (e.data.type === 'wc-done' || e.data.type === 'wc-cancel') { window.removeEventListener('message', onMsg); resolve(e.data); }
      }
      window.addEventListener('message', onMsg);
    });
  }

  // ---------- 예비 방식: 크롬 다운로드 폴더로 받기 ----------
  function downloadViaChrome(item, index) {
    return new Promise(function (resolve, reject) {
      var frameName = 'excelDownFrame_' + Date.now() + '_' + index;
      var iframe = document.createElement('iframe');
      iframe.name = frameName;
      iframe.style.display = 'none';
      document.body.appendChild(iframe);

      var form = document.createElement('form');
      form.method = 'post';
      form.action = EXCEL_URL;
      form.target = frameName;
      form.style.display = 'none';
      var params = excelParams(item);
      Object.keys(params).forEach(function (k) {
        var input = document.createElement('input');
        input.type = 'hidden';
        input.name = k;
        input.value = params[k];
        form.appendChild(input);
      });
      document.body.appendChild(form);

      var done = false;
      function finish(err) {
        if (done) return;
        done = true;
        form.remove();
        setTimeout(function () { iframe.remove(); }, 60000);
        if (err) reject(err); else resolve('다운로드 폴더');
      }

      // 파일 응답은 iframe 화면을 바꾸지 않는다. 화면(HTML)이 들어오면 로그인 만료나 서버 오류다.
      iframe.addEventListener('load', function () {
        var text = '';
        try {
          var doc = iframe.contentDocument;
          if (!doc || doc.URL === 'about:blank') return;
          text = doc.body ? clean(doc.body.innerText).slice(0, 120) : '';
        } catch (e) {}
        finish(new Error(text || '엑셀 대신 화면이 돌아왔습니다'));
      });

      form.submit();
      setTimeout(function () { finish(null); }, ERROR_WAIT_MS);
    });
  }

  // ---------- 화면(패널) ----------
  var classes = readClasses();

  var old = document.getElementById('excelDownPanel');
  if (old) old.remove();

  var panel = document.createElement('div');
  panel.id = 'excelDownPanel';
  panel.setAttribute('role', 'dialog');
  panel.setAttribute('aria-label', NAME);
  panel.innerHTML =
    '<style>' +
    '#excelDownPanel{position:fixed;top:24px;right:24px;z-index:2147483647;width:780px;max-width:calc(100vw - 32px);' +
    'max-height:calc(100vh - 48px);overflow:auto;background:#fff;color:#0f172a;border:1px solid #cbd5e1;border-radius:12px;' +
    'box-shadow:0 12px 40px rgba(15,23,42,.25);font:14px/1.5 "Malgun Gothic","Noto Sans KR",system-ui,sans-serif;text-align:left}' +
    '#excelDownPanel *{box-sizing:border-box}' +
    '#excelDownPanel .xd-head{display:flex;justify-content:space-between;align-items:center;padding:14px 16px;border-bottom:1px solid #e2e8f0}' +
    '#excelDownPanel .xd-head strong{font-size:16px}' +
    '#excelDownPanel .xd-x{border:none;background:none;font-size:22px;line-height:1;cursor:pointer;color:#475569;padding:4px 8px;min-width:36px;min-height:36px}' +
    '#excelDownPanel .xd-body{padding:14px 16px;display:flex;flex-direction:column;gap:10px}' +
    '#excelDownPanel ul{list-style:none;margin:0;padding:0;display:flex;flex-direction:column;gap:6px}' +
    '#excelDownPanel li{display:flex;gap:10px;align-items:flex-start;border:1px solid #e2e8f0;border-radius:8px;padding:6px 10px}' +
    '#excelDownPanel li .xd-mark{min-width:56px;font-weight:700;color:#475569}' +
    '#excelDownPanel li .xd-chk{width:20px;height:20px;margin:2px 0 0;flex:none;cursor:pointer}' +
    '#excelDownPanel li label.xd-t{cursor:pointer}' +
    '#excelDownPanel .xd-badge{display:inline-block;border-radius:999px;padding:0 8px;font-size:12px;font-weight:700;margin-right:2px}' +
    '#excelDownPanel li .xd-t{display:flex;flex-direction:column;min-width:0;flex:1}' +
    '#excelDownPanel li .xd-t > span{white-space:nowrap;overflow:hidden;text-overflow:ellipsis}' +
    '#excelDownPanel li .xd-s{color:#475569;font-size:12px}' +
    '#excelDownPanel .xd-note{color:#475569;font-size:13px;margin:0}' +
    '#excelDownPanel .xd-box{background:#f8fafc;border:1px solid #e2e8f0;border-radius:8px;padding:8px 10px;display:flex;gap:8px;align-items:center;justify-content:space-between;flex-wrap:wrap}' +
    '#excelDownPanel .xd-warn{background:#fff7ed;border:1px solid #fed7aa;border-radius:8px;padding:8px 10px;font-size:13px;margin:0}' +
    '#excelDownPanel .xd-foot{display:flex;gap:8px;justify-content:flex-end;flex-wrap:wrap;padding:12px 16px;border-top:1px solid #e2e8f0}' +
    '#excelDownPanel button.xd-btn{font:inherit;font-size:15px;border-radius:8px;padding:8px 16px;min-height:44px;cursor:pointer;border:1px solid #cbd5e1;background:#fff;color:#0f172a}' +
    '#excelDownPanel button.xd-small{font:inherit;font-size:13px;border-radius:6px;padding:4px 10px;min-height:32px;cursor:pointer;border:1px solid #cbd5e1;background:#fff;color:#0f172a}' +
    '#excelDownPanel button.xd-primary{background:#1d4ed8;border-color:#1d4ed8;color:#fff;font-weight:700}' +
    '#excelDownPanel button.xd-link{border:none;background:none;color:#1d4ed8;text-decoration:underline;font-size:13px;min-height:44px;padding:0 4px}' +
    '#excelDownPanel button:disabled{opacity:.5;cursor:default}' +
    '</style>' +
    '<div class="xd-head"><strong>' + NAME + '</strong>' +
    '<button type="button" class="xd-x" aria-label="닫기">×</button></div>' +
    '<div class="xd-body"></div>' +
    '<div class="xd-foot"></div>';
  document.body.appendChild(panel);

  var body = panel.querySelector('.xd-body');
  var foot = panel.querySelector('.xd-foot');
  var running = false;

  function close() {
    if (running && !confirm('받는 중입니다. 창을 닫을까요?')) return;
    panel.remove();
  }
  panel.querySelector('.xd-x').addEventListener('click', close);

  function button(label, cls, onClick, parent) {
    var b = document.createElement('button');
    b.type = 'button';
    b.className = cls;
    b.textContent = label;
    b.addEventListener('click', onClick);
    (parent || foot).appendChild(b);
    return b;
  }

  if (!classes.length) {
    body.innerHTML =
      '<p style="margin:0">이 화면에 <strong>운영이 끝나지 않은 접수중·접수마감 수업</strong>이 없습니다.</p>' +
      '<p class="xd-note">누리집이 「울산창의융합교육센터」인지, 검색 조건이 걸려 있지 않은지 확인해 주세요.</p>' +
      '<p class="xd-note">버전 ' + VERSION + '</p>';
    button('닫기', 'xd-btn xd-primary', close);
    return;
  }

  body.innerHTML =
    '<p style="margin:0">받을 수업을 확인해 주세요. <span class="xd-note">(가장 가까운 수업일 수업이 미리 체크되어 있습니다)</span></p>' +
    '<ul>' + classes.map(function (c, i) {
      var badge = c.status === '접수중'
        ? '<span class="xd-badge" style="background:#ecfdf3;color:#027a48">접수중</span>'
        : '<span class="xd-badge" style="background:#fef3f2;color:#b42318">접수마감</span>';
      return '<li data-i="' + i + '"><input type="checkbox" class="xd-chk" id="xdChk' + i + '"' + (c.checked ? ' checked' : '') + '>' +
        '<span class="xd-mark">' + (c.checked ? '대기' : '제외') + '</span><label class="xd-t" for="xdChk' + i + '" title="' + esc(c.title) + '"><span>' +
        badge + ' ' + esc(c.title) + '</span><span class="xd-s">' + esc([c.date, c.count].filter(Boolean).join(' · ')) +
        '</span></label></li>';
    }).join('') + '</ul>' +
    '<div class="xd-box"><span class="xd-folder"></span></div>' +
    '<p class="xd-warn xd-tip"></p>' +
    '<p class="xd-note">버전 ' + VERSION + '</p>';

  var folderLabel = body.querySelector('.xd-folder');
  var folderBox = body.querySelector('.xd-box');
  var tip = body.querySelector('.xd-tip');
  var dirHandle = null;

  function showFolder() {
    if (!canPickFolder) {
      folderLabel.innerHTML = '이 크롬에서는 폴더 저장을 쓸 수 없어 <strong>크롬 다운로드 폴더</strong>로 받습니다.';
      tip.innerHTML = '크롬 위쪽에 <strong>「여러 파일 다운로드」</strong> 허용 창이 뜨면 <strong>[허용]</strong>을 눌러 주세요.';
      return;
    }
    folderLabel.innerHTML = dirHandle
      ? '저장 폴더: <strong>' + esc(dirHandle.name) + '</strong>'
      : '저장 폴더: <strong>아직 고르지 않음</strong> ([폴더에 저장]을 누르면 고르는 창이 뜹니다)';
    tip.innerHTML = dirHandle
      ? '크롬이 <strong>「폴더에 저장 허용」</strong>을 물으면 <strong>[허용]</strong>(가능하면 「방문할 때마다 허용」)을 눌러 주세요.'
      : '폴더 고르는 창에서 <strong>' + esc(FOLDER_HINT) + '</strong>을 골라 주세요. (없으면 그 창에서 새 폴더로 만들 수 있습니다)';
  }

  function mark(i, text, color, detail) {
    var li = body.querySelector('li[data-i="' + i + '"]');
    if (!li) return;
    var m = li.querySelector('.xd-mark');
    m.textContent = text;
    m.style.color = color;
    if (detail) li.querySelector('.xd-s').textContent = detail;
  }

  var changeBtn = null;
  if (canPickFolder) {
    changeBtn = button('폴더 바꾸기', 'xd-small', function () {
      pickFolder().then(function (h) { dirHandle = h; showFolder(); }, function () {});
    }, folderBox);
  }

  var fallbackBtn = canPickFolder ? button('다운로드 폴더로 받기', 'xd-link', function () { start('chrome'); }) : null;
  var cancelBtn = button('취소', 'xd-btn', close);
  var startBtn = button('받기', 'xd-btn', function () {
    start(canPickFolder ? 'folder' : 'chrome');
  });
  var sendBtn = button('점검 웹앱으로 보내기', 'xd-btn xd-primary', startSend);

  function selected() {
    return classes.filter(function (c, i) {
      var chk = body.querySelector('#xdChk' + i);
      return chk && chk.checked;
    });
  }
  function refreshCount() {
    var n = selected().length;
    startBtn.textContent = n + '개 ' + (canPickFolder ? '폴더에 저장' : '받기');
    startBtn.disabled = running || n === 0;
    sendBtn.textContent = n + '개 점검 웹앱으로 보내기';
    sendBtn.disabled = running || n === 0;
    if (fallbackBtn) fallbackBtn.disabled = running || n === 0;
    classes.forEach(function (c, i) {
      var chk = body.querySelector('#xdChk' + i);
      if (chk) mark(i, chk.checked ? '대기' : '제외', '#475569');
    });
  }
  body.addEventListener('change', function (e) {
    if (e.target && e.target.classList.contains('xd-chk')) refreshCount();
  });
  refreshCount();

  // 기억한 폴더 읽기. 혹시 읽다가 탭이 멈춘 적이 있으면(표시가 남아 있으면) 건너뛰고 다시 고르게 한다.
  var loadFolder = !canPickFolder ? Promise.resolve(null) : idbGet('reading').then(function (crashedBefore) {
    if (crashedBefore) {
      return idbSet('dir', null).then(function () { return idbSet('reading', null); }).then(function () { return null; });
    }
    return idbSet('reading', true).then(function () { return idbGet('dir'); }).then(function (h) {
      return idbSet('reading', null).then(function () { return h; });
    });
  });
  loadFolder.then(function (h) {
    if (h && h.kind === 'directory') dirHandle = h;
    showFolder();
    sendBtn.focus();
  });

  function setBusy(b) {
    running = b;
    window.__excelDownBusy = b;
    [startBtn, sendBtn, cancelBtn, changeBtn, fallbackBtn].forEach(function (x) { if (x) x.disabled = b; });
    body.querySelectorAll('.xd-chk').forEach(function (x) { x.disabled = b; });
  }

  function start(mode) {
    if (running || !selected().length) return;
    var ready = mode === 'folder'
      ? ensureFolder(dirHandle).then(function (h) { dirHandle = h; showFolder(); return h; })
      : Promise.resolve(null);

    setBusy(true);
    ready.then(function (dir) {
      if (mode === 'chrome') {
        tip.innerHTML = '크롬 위쪽에 <strong>「여러 파일 다운로드」</strong> 허용 창이 뜨면 <strong>[허용]</strong>을 눌러 주세요.';
      }
      return run(mode, dir);
    }, function (err) {
      setBusy(false);
      if (err && err.name === 'AbortError') return; // 폴더 고르기를 취소함
      alert(NAME + '\n\n' + (err && err.message ? err.message : '폴더를 열 수 없습니다.') +
        '\n\n다시 누르거나 [다운로드 폴더로 받기]를 써 주세요.');
    });
  }

  // 반드시 버튼 클릭 직후에 창을 연다 (나중에 열면 크롬이 팝업으로 막는다)
  function startSend() {
    if (running || !selected().length) return;
    var win = window.open(WEBAPP_URL + '#from=cms', WEBAPP_WINDOW);
    if (!win) {
      alert(NAME + '\n\n크롬이 점검 웹앱 창을 막았습니다.\n주소창 오른쪽의 팝업 차단 표시를 눌러 「항상 허용」을 고른 뒤 다시 눌러 주세요.');
      return;
    }
    setBusy(true);
    runSend(win);
  }

  function runSend(win) {
    var failed = [], got = [];
    var todo = selected();
    var link = '점검 웹앱 창: <strong>연결 중…</strong> (처음이면 웹앱에서 로그인해 주세요)';
    function showLink(t) { link = t; tip.innerHTML = link; }
    showLink(link);
    folderBox.style.display = 'none';
    // 엑셀 받기와 웹앱 연결을 함께 진행한다
    var ready = waitMessage(win, 'wc-ready', READY_WAIT_MS, function () {
      try { win.postMessage({ type: 'xd-hello', version: VERSION }, WEBAPP_ORIGIN); } catch (e) {}
    }).then(function (r) { showLink('점검 웹앱 창: <strong style="color:#027a48">연결됨</strong>'); return r; });
    ready.catch(function () {});
    var chain = Promise.resolve();
    todo.forEach(function (c, n) {
      var i = classes.indexOf(c);
      chain = chain.then(function () {
        mark(i, '받는 중', '#1d4ed8');
        return fetchExcel(c).then(function (r) {
          return r.blob.arrayBuffer().then(function (buf) {
            got.push({ name: r.name, at: Date.now(), title: c.title, bytes: buf });
            mark(i, '받음', '#027a48');
          });
        }, function (err) {
          failed.push({ item: c, detail: err && err.message ? err.message : '' });
          mark(i, '실패', '#b42318');
        }).then(function () { if (n < todo.length - 1) return sleep(GAP_MS); });
      });
    });

    var msg = document.createElement('div');
    msg.className = 'xd-result';
    function finish(html) {
      setBusy(false);
      msg.innerHTML = html + failed.map(function (f) {
        return '<p class="xd-note">· 받지 못함: ' + esc(f.item.title) + (f.detail ? ' — ' + esc(f.detail) : '') + '</p>';
      }).join('');
      if (!msg.parentNode) body.appendChild(msg);
      foot.innerHTML = '';
      button('닫기', 'xd-btn xd-primary', close).focus();
    }

    return chain.then(function () {
      if (!got.length) {
        finish('<p style="margin:0;font-weight:700;color:#b42318">명단을 하나도 받지 못했습니다.</p><p class="xd-note">로그인이 끊겼을 수 있습니다. 다시 로그인한 뒤 버튼을 다시 눌러 주세요.</p>');
        return;
      }
      msg.innerHTML = '<p style="margin:0;font-weight:700">' + got.length + '개 받음 · 점검 웹앱 연결을 기다리는 중…</p>';
      body.appendChild(msg);
      return ready.then(function () {
        var done = waitDone(win);
        win.postMessage({ type: 'xd-files', version: VERSION, files: got }, WEBAPP_ORIGIN, got.map(function (g) { return g.bytes; }));
        msg.innerHTML = '<p style="margin:0;font-weight:700;color:#1d4ed8">' + got.length + '개를 점검 웹앱으로 보냈습니다.</p>' +
          '<p class="xd-note">점검 웹앱 창에서 <strong>[올리기]</strong>를 눌러 주세요.</p>';
        try { win.focus(); } catch (e) {}
        // 웹앱이 한참 답이 없어도 이 창을 닫을 수 있도록 버튼을 풀어 둔다
        setBusy(false);
        foot.innerHTML = '';
        button('닫기', 'xd-btn', close);
        return done.then(function (r) {
          if (r.type === 'wc-cancel') { finish('<p style="margin:0;font-weight:700">점검 웹앱에서 올리기를 취소했습니다.</p><p class="xd-note">필요하면 버튼을 다시 눌러 주세요.</p>'); return; }
          if (!r.ok) { finish('<p style="margin:0;font-weight:700;color:#b42318">점검 웹앱에서 올리지 못했습니다.</p><p class="xd-note">' + esc(r.message || '') + '</p>'); return; }
          var res = r.results || [];
          var applied = res.filter(function (x) { return x.result === '반영'; }).length;
          finish('<p style="margin:0;font-weight:700;color:#027a48">점검 웹앱에 올렸습니다 · 반영 ' + applied + '개' +
            (res.length - applied ? ' · 그 밖 ' + (res.length - applied) + '개' : '') + (r.skipped ? ' · 접수 기간 밖 ' + r.skipped + '개' : '') + '</p>' +
            res.map(function (x) { return '<p class="xd-note">· <strong>' + esc(x.result) + '</strong> ' + esc(x.title) + (x.detail ? ' — ' + esc(x.detail) : '') + '</p>'; }).join(''));
        });
      }, function (err) {
        finish('<p style="margin:0;font-weight:700;color:#b42318">' + esc(err.message) + '</p>' +
          '<p class="xd-note">점검 웹앱 창에서 로그인했는지 확인한 뒤 버튼을 다시 누르거나, [폴더에 저장]으로 받아 주세요.</p>');
      });
    });
  }

  function run(mode, dir) {
    var failed = [];
    var chain = Promise.resolve();
    var todo = selected();
    todo.forEach(function (c, n) {
      var i = classes.indexOf(c);
      chain = chain.then(function () {
        mark(i, '받는 중', '#1d4ed8');
        var job = mode === 'folder' ? saveToFolder(dir, c) : downloadViaChrome(c, i);
        return job.then(function (savedAs) {
          mark(i, '완료', '#027a48', mode === 'folder' ? savedAs : null);
        }, function (err) {
          failed.push({ item: c, detail: err && err.message ? err.message : '' });
          mark(i, '실패', '#b42318');
        }).then(function () {
          if (n < todo.length - 1) return sleep(mode === 'folder' ? GAP_MS : IFRAME_GAP_MS);
        });
      });
    });

    return chain.then(function () {
      setBusy(false);
      var ok = todo.length - failed.length;
      var where = mode === 'folder' ? '「' + esc(dir.name) + '」 폴더' : '다운로드 폴더';
      var msg = document.createElement('div');
      msg.className = 'xd-result';
      if (!failed.length) {
        msg.innerHTML = '<p style="margin:0;font-weight:700;color:#027a48">' + ok +
          '개 모두 받았습니다.</p><p class="xd-note">' + where + '를 확인한 뒤 점검 웹앱에서 [명단 올리기]를 눌러 주세요.</p>';
      } else {
        msg.innerHTML = '<p style="margin:0;font-weight:700;color:#b42318">' + ok + '개 받음 · ' + failed.length +
          '개 실패</p><p class="xd-note">로그인이 끊겼을 수 있습니다. 다시 로그인한 뒤 버튼을 다시 누르거나, 실패한 수업은 직접 받아 주세요.</p>' +
          failed.map(function (f) {
            return '<p class="xd-note">· ' + esc(f.item.title) + (f.detail ? ' — ' + esc(f.detail) : '') + '</p>';
          }).join('');
      }
      body.appendChild(msg);
      foot.innerHTML = '';
      button('닫기', 'xd-btn xd-primary', close).focus();
    });
  }
})();
