/*!
 * 통합예약 엑셀 일괄 받기
 * 울산광역시교육청 통합예약 관리자 > 통합예약 목록 화면에서
 * '접수중'인 수업의 신청자 엑셀을 한 번에 내려받는다.
 *
 * 동작 방식: 각 수업의 [신청 관리] 버튼에 들어 있는 번호(opViewApplyList)를 읽고,
 * 신청자 목록 화면의 [엑셀출력](opExcel)이 보내는 것과 같은 요청을
 * 숨은 iframe으로 보내 브라우저가 서버가 주는 파일 이름 그대로 저장하게 한다.
 */
(function () {
  'use strict';

  var VERSION = '1.0.0';
  var NAME = '통합예약 엑셀 일괄 받기';
  var LIST_PAGE = 'BD_selectReservationMngList.do';
  var EXCEL_URL = '/component/excel/download/ND_excelDownload.do';
  var EXCEL_KEY = 'reservationNoContsCceAplyList';
  var STATUS_TEXT = '접수중';
  var ROW_PER_PAGE = '500'; // 엑셀에 인원이 잘리지 않도록 넉넉하게
  var GAP_MS = 2500; // 다운로드 사이 간격
  var ERROR_WAIT_MS = 6000; // 오류 화면이 돌아오는지 기다리는 시간

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

  // ---------- 목록 읽기 ----------
  function clean(s) {
    return (s || '').replace(/\s+/g, ' ').trim();
  }

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

      var isOpen = false;
      if (statusIdx >= 0 && cells[statusIdx]) {
        isOpen = clean(cells[statusIdx].textContent) === STATUS_TEXT;
      } else {
        for (var c = 0; c < cells.length; c++) {
          if (clean(cells[c].textContent) === STATUS_TEXT) { isOpen = true; break; }
        }
      }
      if (!isOpen) continue;

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
        date: dateIdx >= 0 && cells[dateIdx] ? clean(cells[dateIdx].innerText) : '',
        count: countIdx >= 0 && cells[countIdx] ? clean(cells[countIdx].innerText) : ''
      });
    }
    return found;
  }

  // ---------- 엑셀 요청 (opExcel과 같은 값) ----------
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

  // 결과: 'sent'(파일로 내려옴 → 오류 화면 없음) 또는 'error'(화면이 돌아옴 = 실패)
  function requestExcel(item, index) {
    return new Promise(function (resolve) {
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
      function finish(result, detail) {
        if (done) return;
        done = true;
        form.remove();
        // 다운로드가 끝까지 받아지도록 iframe은 잠시 뒤에 지운다
        setTimeout(function () { iframe.remove(); }, 60000);
        resolve({ result: result, detail: detail || '' });
      }

      // 파일 다운로드 응답은 iframe 화면을 바꾸지 않는다.
      // 화면(HTML)이 들어왔다면 로그인 만료나 서버 오류다.
      iframe.addEventListener('load', function () {
        var text = '';
        try {
          var doc = iframe.contentDocument;
          text = doc && doc.body ? clean(doc.body.innerText).slice(0, 120) : '';
          if (!doc || doc.URL === 'about:blank') return;
        } catch (e) {
          text = '';
        }
        finish('error', text);
      });

      form.submit();
      setTimeout(function () { finish('sent'); }, ERROR_WAIT_MS);
    });
  }

  function sleep(ms) {
    return new Promise(function (r) { setTimeout(r, ms); });
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
    '#excelDownPanel{position:fixed;top:24px;right:24px;z-index:2147483647;width:440px;max-width:calc(100vw - 32px);' +
    'max-height:calc(100vh - 48px);overflow:auto;background:#fff;color:#0f172a;border:1px solid #cbd5e1;border-radius:12px;' +
    'box-shadow:0 12px 40px rgba(15,23,42,.25);font:14px/1.5 "Malgun Gothic","Noto Sans KR",system-ui,sans-serif;text-align:left}' +
    '#excelDownPanel *{box-sizing:border-box}' +
    '#excelDownPanel .xd-head{display:flex;justify-content:space-between;align-items:center;padding:14px 16px;border-bottom:1px solid #e2e8f0}' +
    '#excelDownPanel .xd-head strong{font-size:16px}' +
    '#excelDownPanel .xd-x{border:none;background:none;font-size:22px;line-height:1;cursor:pointer;color:#475569;padding:4px 8px;min-width:36px;min-height:36px}' +
    '#excelDownPanel .xd-body{padding:14px 16px;display:flex;flex-direction:column;gap:10px}' +
    '#excelDownPanel ul{list-style:none;margin:0;padding:0;display:flex;flex-direction:column;gap:6px}' +
    '#excelDownPanel li{display:flex;gap:10px;align-items:flex-start;border:1px solid #e2e8f0;border-radius:8px;padding:8px 10px}' +
    '#excelDownPanel li .xd-mark{min-width:56px;font-weight:700;color:#475569}' +
    '#excelDownPanel li .xd-t{display:flex;flex-direction:column}' +
    '#excelDownPanel li .xd-s{color:#475569;font-size:12px}' +
    '#excelDownPanel .xd-note{color:#475569;font-size:13px;margin:0}' +
    '#excelDownPanel .xd-warn{background:#fff7ed;border:1px solid #fed7aa;border-radius:8px;padding:8px 10px;font-size:13px;margin:0}' +
    '#excelDownPanel .xd-foot{display:flex;gap:8px;justify-content:flex-end;padding:12px 16px;border-top:1px solid #e2e8f0}' +
    '#excelDownPanel button.xd-btn{font:inherit;font-size:15px;border-radius:8px;padding:8px 16px;min-height:44px;cursor:pointer;border:1px solid #cbd5e1;background:#fff;color:#0f172a}' +
    '#excelDownPanel button.xd-primary{background:#1d4ed8;border-color:#1d4ed8;color:#fff;font-weight:700}' +
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
    if (running && !confirm('다운로드 중입니다. 창을 닫을까요? (이미 시작한 파일은 계속 받아집니다)')) return;
    panel.remove();
  }
  panel.querySelector('.xd-x').addEventListener('click', close);

  function esc(s) {
    return String(s).replace(/[&<>"']/g, function (ch) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[ch];
    });
  }

  function button(label, primary, onClick) {
    var b = document.createElement('button');
    b.type = 'button';
    b.className = 'xd-btn' + (primary ? ' xd-primary' : '');
    b.textContent = label;
    b.addEventListener('click', onClick);
    foot.appendChild(b);
    return b;
  }

  if (!classes.length) {
    body.innerHTML =
      '<p style="margin:0">이 화면에 <strong>' + STATUS_TEXT + '</strong>인 수업이 없습니다.</p>' +
      '<p class="xd-note">누리집이 「울산창의융합교육센터」인지, 검색 조건이 걸려 있지 않은지 확인해 주세요. ' +
      '접수마감 수업은 직접 받아 주세요.</p>' +
      '<p class="xd-note">버전 ' + VERSION + '</p>';
    button('닫기', true, close);
    return;
  }

  body.innerHTML =
    '<p style="margin:0"><strong>' + STATUS_TEXT + '</strong>인 수업 <strong>' + classes.length +
    '개</strong>의 신청자 엑셀을 받을까요?</p>' +
    '<ul>' + classes.map(function (c, i) {
      return '<li data-i="' + i + '"><span class="xd-mark">대기</span><span class="xd-t"><span>' +
        esc(c.title) + '</span><span class="xd-s">' + esc([c.date, c.count].filter(Boolean).join(' · ')) +
        '</span></span></li>';
    }).join('') + '</ul>' +
    '<p class="xd-warn">크롬 위쪽에 <strong>「여러 파일 다운로드」</strong> 허용 창이 뜨면 <strong>[허용]</strong>을 눌러 주세요.</p>' +
    '<p class="xd-note">버전 ' + VERSION + '</p>';

  function mark(i, text, color) {
    var li = body.querySelector('li[data-i="' + i + '"] .xd-mark');
    if (li) { li.textContent = text; li.style.color = color; }
  }

  var cancelBtn = button('취소', false, close);
  var startBtn = button(classes.length + '개 받기', true, function () {
    start();
  });
  startBtn.focus();

  function start() {
    running = true;
    window.__excelDownBusy = true;
    startBtn.disabled = true;
    cancelBtn.disabled = true;

    var failed = [];
    var chain = Promise.resolve();
    classes.forEach(function (c, i) {
      chain = chain.then(function () {
        mark(i, '받는 중', '#1d4ed8');
        return requestExcel(c, i).then(function (r) {
          if (r.result === 'error') {
            failed.push({ item: c, detail: r.detail });
            mark(i, '실패', '#b42318');
          } else {
            mark(i, '완료', '#027a48');
          }
          if (i < classes.length - 1) return sleep(GAP_MS);
        });
      });
    });

    chain.then(function () {
      running = false;
      window.__excelDownBusy = false;
      var ok = classes.length - failed.length;
      var msg = document.createElement('div');
      if (!failed.length) {
        msg.innerHTML = '<p style="margin:0;font-weight:700;color:#027a48">' + ok +
          '개 모두 받았습니다.</p><p class="xd-note">다운로드 폴더를 확인한 뒤 점검 웹앱에서 [명단 올리기]를 눌러 주세요.</p>';
      } else {
        msg.innerHTML = '<p style="margin:0;font-weight:700;color:#b42318">' + ok + '개 받음 · ' + failed.length +
          '개 실패</p><p class="xd-note">로그인이 끊겼을 수 있습니다. 다시 로그인한 뒤 버튼을 다시 누르거나, 실패한 수업은 직접 받아 주세요.</p>' +
          failed.map(function (f) {
            return '<p class="xd-note">· ' + esc(f.item.title) + (f.detail ? ' — ' + esc(f.detail) : '') + '</p>';
          }).join('');
      }
      body.appendChild(msg);
      foot.innerHTML = '';
      button('닫기', true, close).focus();
    });
  }
})();
