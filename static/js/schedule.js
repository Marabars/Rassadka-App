window.App = window.App || {};

App.schedule = (function () {
  'use strict';

  var U = App.U;

  /* ── Module state (survives tab switches) ─────────────────── */
  var _container  = null;
  var _tableWrap  = null;
  var _state      = { nameFilter: '', dateFrom: '', dateTo: '', collapsedWeeks: {} };
  var _prefs      = {};   // { employeeName: ['16.38', '16.40', ...] }
  var _assignments = null;
  var _cachedMonth = null;
  var _debounce   = {};

  /* ── Layout constants ─────────────────────────────────────── */
  var COL1_W = 260;  // ФИО, px
  var COL2_W = 210;  // Preferred seats, px

  /* Week palette: alternating background pairs */
  var WEEKS_BG  = ['#121027', '#0C1021'];
  var WEEKS_HDR = ['#1B1938', '#101828'];
  var WEEK_SEP  = '2px solid #7549E8';

  /* ── Entry point ──────────────────────────────────────────── */
  function render(container) {
    _container = container;
    container.innerHTML = '';
    var month = App.state.getMonth();

    /* ── Top bar ── */
    var bar = U.el('div', {
      style: 'display:flex;align-items:center;justify-content:space-between;margin-bottom:20px;flex-wrap:wrap;gap:14px'
    });
    bar.appendChild(U.el('div', {}, [
      U.el('h2', { text: 'График рассадки', style: 'margin:0 0 4px' }),
      U.el('div', { class: 'muted', text: month ? _fmtMonth(month) : 'Месяц не выбран' })
    ]));
    var exportBtn = U.el('button', { class: 'btn btn-secondary', text: '↓ Скачать xlsx' });
    if (!month) exportBtn.disabled = true;
    exportBtn.addEventListener('click', function () {
      if (month) window.location.href = '/api/seating/export?month=' + encodeURIComponent(month);
    });
    bar.appendChild(exportBtn);
    container.appendChild(bar);

    if (!month) {
      container.appendChild(_emptyState('Сначала выберите месяц и сгенерируйте рассадку.'));
      return;
    }

    /* ── Filter bar ── */
    container.appendChild(_buildFilters());

    /* ── Table container (scrollable) ── */
    _tableWrap = U.el('div', {
      style: 'overflow:auto;max-height:72vh;border-radius:8px;border:1px solid #2E2A4A'
    });
    container.appendChild(_tableWrap);

    var loadEl = U.el('div', { class: 'muted', style: 'padding:28px', text: 'Загружаю…' });
    _tableWrap.appendChild(loadEl);

    /* ── Fetch data ── */
    var dataPromise = (month !== _cachedMonth)
      ? App.api.getSeating(month).then(function (d) {
          _assignments = d.assignments || [];
          _cachedMonth = month;
        })
      : Promise.resolve();

    App.api.getPreferences().then(function (d) {
      _prefs = {};
      (d.preferences || []).forEach(function (p) { _prefs[p.name] = p.seats; });
    }).catch(function () {});

    dataPromise.then(function () {
      _tableWrap.removeChild(loadEl);
      _buildTable();
    }).catch(function (err) {
      loadEl.textContent = 'Ошибка: ' + err.message;
    });
  }

  /* ── Filter bar ──────────────────────────────────────────── */
  function _buildFilters() {
    var bar = U.el('div', {
      style: 'display:flex;gap:12px;align-items:center;flex-wrap:wrap;margin-bottom:14px'
    });

    var nameInp = document.createElement('input');
    nameInp.type = 'text';
    nameInp.placeholder = '🔍 Поиск по ФИО…';
    nameInp.value = _state.nameFilter;
    nameInp.style.cssText = _inpStyle('flex:1;min-width:200px;max-width:320px');
    nameInp.addEventListener('input', function () {
      _state.nameFilter = nameInp.value;
      _buildTable();
    });
    bar.appendChild(nameInp);

    bar.appendChild(U.el('span', {
      text: 'Период:',
      style: 'color:#9994BB;font-size:20px;white-space:nowrap'
    }));

    var fromInp = _dateInp(_state.dateFrom, function (v) { _state.dateFrom = v; _buildTable(); });
    bar.appendChild(fromInp);
    bar.appendChild(U.el('span', { text: '—', style: 'color:#4A4668;font-size:22px' }));
    var toInp = _dateInp(_state.dateTo, function (v) { _state.dateTo = v; _buildTable(); });
    bar.appendChild(toInp);

    var clearBtn = U.el('button', {
      class: 'btn btn-secondary',
      text: '✕',
      style: 'padding:8px 14px;font-size:20px',
      title: 'Сбросить период'
    });
    clearBtn.addEventListener('click', function () {
      _state.dateFrom = '';
      _state.dateTo   = '';
      fromInp.value   = '';
      toInp.value     = '';
      _buildTable();
    });
    bar.appendChild(clearBtn);

    return bar;
  }

  function _dateInp(val, cb) {
    var inp = document.createElement('input');
    inp.type = 'date';
    inp.value = val;
    inp.style.cssText = _inpStyle('width:168px');
    inp.addEventListener('change', function () { cb(inp.value); });
    return inp;
  }

  function _inpStyle(extra) {
    return [
      'padding:8px 12px',
      'font-size:20px',
      'background:#13112A',
      'border:1px solid #2E2A4A',
      'border-radius:6px',
      'color:#E5E3EB',
      'outline:none',
      extra
    ].join(';');
  }

  /* ── Table builder ───────────────────────────────────────── */
  function _buildTable() {
    if (!_assignments) return;
    _tableWrap.innerHTML = '';

    /* Pivot */
    var empData  = {};
    var allDates = {};
    _assignments.forEach(function (a) {
      allDates[a.date] = true;
      if (!empData[a.employee_name]) empData[a.employee_name] = {};
      if (a.status === 'OFFICE') empData[a.employee_name][a.date] = a.seat_id || '';
    });

    /* Apply filters */
    var nameFilter = _state.nameFilter.toLowerCase();
    var sortedDates = Object.keys(allDates).sort().filter(function (d) {
      if (_state.dateFrom && d < _state.dateFrom) return false;
      if (_state.dateTo   && d > _state.dateTo)   return false;
      return true;
    });
    var sortedNames = Object.keys(empData).sort().filter(function (n) {
      return !nameFilter || n.toLowerCase().indexOf(nameFilter) >= 0;
    });

    if (!sortedDates.length || !sortedNames.length) {
      _tableWrap.appendChild(U.el('div', {
        style: 'padding:36px;text-align:center;color:#7E7A9A;font-size:22px',
        text: 'Нет данных для выбранных фильтров.'
      }));
      return;
    }

    /* Office counts per date */
    var officeCounts = {};
    sortedDates.forEach(function (d) {
      var cnt = 0; sortedNames.forEach(function (n) { if (empData[n][d]) cnt++; });
      officeCounts[d] = cnt;
    });

    /* Group dates by week */
    var weeks = _groupByWeek(sortedDates);

    var table = document.createElement('table');
    table.style.cssText = 'border-collapse:collapse;font-size:22px;min-width:max-content;width:100%';

    /* ── THEAD ── */
    var thead = document.createElement('thead');
    thead.style.cssText = 'position:sticky;top:0;z-index:10';
    table.appendChild(thead);

    /* Row 1 — week groups */
    var trW = document.createElement('tr');
    thead.appendChild(trW);

    /* ФИО (rowspan=2) */
    var thFio = document.createElement('th');
    thFio.textContent = 'ФИО';
    thFio.rowSpan = 2;
    thFio.style.cssText = _stickyHdrStyle(0, COL1_W);
    trW.appendChild(thFio);

    /* Preferred seats (rowspan=2) */
    var thPref = document.createElement('th');
    thPref.textContent = 'Предп. места';
    thPref.rowSpan = 2;
    thPref.style.cssText = _stickyHdrStyle(COL1_W, COL2_W) + ';border-right:' + WEEK_SEP;
    trW.appendChild(thPref);

    /* Week group headers */
    weeks.forEach(function (wk, wi) {
      var col   = _state.collapsedWeeks[wk.key];
      var label = _weekRange(wk.dates);
      var th    = document.createElement('th');
      th.textContent = (col ? '▶ ' : '▼ ') + label;
      th.title  = col ? 'Развернуть' : 'Свернуть';
      th.colSpan = col ? 1 : wk.dates.length;
      th.rowSpan = col ? 2 : 1;
      th.style.cssText = [
        'background:' + WEEKS_HDR[wi % 2],
        'color:#FFFFFF',
        'font-size:19px',
        'font-weight:700',
        'padding:10px 14px',
        'border-bottom:1px solid #2E2A4A',
        'border-right:' + WEEK_SEP,
        'text-align:center',
        'white-space:nowrap',
        'cursor:pointer',
        'user-select:none'
      ].join(';');
      th.addEventListener('click', (function (key) {
        return function () {
          _state.collapsedWeeks[key] = !_state.collapsedWeeks[key];
          _buildTable();
        };
      })(wk.key));
      trW.appendChild(th);
    });

    /* Row 2 — individual dates (expanded weeks only) */
    var trD = document.createElement('tr');
    thead.appendChild(trD);
    weeks.forEach(function (wk, wi) {
      if (_state.collapsedWeeks[wk.key]) return;
      wk.dates.forEach(function (date, di) {
        var parts = date.split('-');
        var th = document.createElement('th');
        th.textContent = parts[2] + '.' + parts[1];
        var isLast = (di === wk.dates.length - 1);
        th.style.cssText = [
          'background:' + WEEKS_HDR[wi % 2],
          'color:#FFFFFF',
          'font-size:18px',
          'font-weight:600',
          'padding:8px 5px',
          'border-bottom:2px solid #433E67',
          'min-width:64px',
          'text-align:center',
          isLast ? 'border-right:' + WEEK_SEP : ''
        ].join(';');
        trD.appendChild(th);
      });
    });

    /* ── TBODY ── */
    var tbody = document.createElement('tbody');
    table.appendChild(tbody);

    /* Count row */
    var trCnt = document.createElement('tr');
    var tdCL = _stickyTd('В офисе', 0, COL1_W, '#0A0814');
    tdCL.style.cssText += ';font-size:18px;font-weight:600;letter-spacing:.06em;text-transform:uppercase;color:#4A4668';
    trCnt.appendChild(tdCL);
    var tdCP = _stickyTd('', COL1_W, COL2_W, '#0A0814');
    tdCP.style.borderRight = WEEK_SEP;
    trCnt.appendChild(tdCP);
    weeks.forEach(function (wk, wi) {
      var bg = WEEKS_BG[wi % 2];
      if (_state.collapsedWeeks[wk.key]) {
        var td = _simpleTd('', bg, true);
        trCnt.appendChild(td);
      } else {
        wk.dates.forEach(function (date, di) {
          var cnt = officeCounts[date];
          var td = _simpleTd(cnt > 0 ? String(cnt) : '—', bg, di === wk.dates.length - 1);
          td.style.fontSize = '22px';
          td.style.fontWeight = '600';
          td.style.color = cnt > 0 ? '#82D6CC' : '#2E2A4A';
          trCnt.appendChild(td);
        });
      }
    });
    tbody.appendChild(trCnt);

    /* Employee rows */
    sortedNames.forEach(function (name, i) {
      var rowBg = i % 2 === 0 ? '#13112A' : '#0F0D20';
      var tr = document.createElement('tr');

      /* Col 1: ФИО */
      var tdN = _stickyTd(name, 0, COL1_W, rowBg);
      tdN.style.cssText += ';color:#E5E3EB;white-space:nowrap';
      tr.appendChild(tdN);

      /* Col 2: Preferred seats */
      var tdP = _stickyTd('', COL1_W, COL2_W, rowBg);
      tdP.style.borderRight = WEEK_SEP;
      tdP.appendChild(_prefInput(name, rowBg));
      tr.appendChild(tdP);

      /* Data cells per week */
      weeks.forEach(function (wk, wi) {
        var bg = WEEKS_BG[wi % 2];
        if (_state.collapsedWeeks[wk.key]) {
          var days = wk.dates.filter(function (d) { return empData[name][d]; }).length;
          var td = _simpleTd(days > 0 ? days + 'д' : '', bg, true);
          td.style.color = days > 0 ? '#82D6CC' : '#2E2A4A';
          td.style.fontSize = '19px';
          tr.appendChild(td);
        } else {
          wk.dates.forEach(function (date, di) {
            var seat = empData[name][date];
            var td = _simpleTd('', bg, di === wk.dates.length - 1);
            if (seat) {
              var chip = document.createElement('span');
              chip.textContent = seat;
              chip.style.cssText = [
                'display:inline-block',
                'padding:2px 7px',
                'border-radius:4px',
                'background:rgba(130,214,204,.1)',
                'color:#82D6CC',
                'font-family:Consolas,"Courier New",monospace',
                'font-size:19px',
                'font-weight:700'
              ].join(';');
              td.appendChild(chip);
            }
            tr.appendChild(td);
          });
        }
      });

      tbody.appendChild(tr);
    });

    _tableWrap.appendChild(table);
  }

  /* ── Preferred seats input ──────────────────────────────── */
  function _prefInput(name, rowBg) {
    var inp = document.createElement('input');
    inp.type = 'text';
    inp.value = (_prefs[name] || []).join(', ');
    inp.placeholder = 'напр. 16.38, 502';
    inp.title = 'До 3 мест через запятую. Используется при генерации.';
    inp.style.cssText = [
      'width:100%',
      'box-sizing:border-box',
      'background:transparent',
      'border:1px solid transparent',
      'border-radius:4px',
      'color:#E5E3EB',
      'font-size:19px',
      'padding:4px 8px',
      'outline:none',
      'font-family:Consolas,"Courier New",monospace'
    ].join(';');
    inp.addEventListener('focus', function () {
      inp.style.border = '1px solid #7549E8';
      inp.style.background = '#1C1A35';
    });
    inp.addEventListener('blur', function () {
      inp.style.border = '1px solid transparent';
      inp.style.background = 'transparent';
      _debounceSave(name, inp.value);
    });
    inp.addEventListener('keydown', function (e) {
      if (e.key === 'Enter') inp.blur();
    });
    return inp;
  }

  function _debounceSave(name, value) {
    var seats = value.split(',')
      .map(function (s) { return s.trim(); })
      .filter(Boolean)
      .slice(0, 3);
    _prefs[name] = seats;
    clearTimeout(_debounce[name]);
    _debounce[name] = setTimeout(function () {
      App.api.setPreference(name, seats).catch(function () {});
    }, 700);
  }

  /* ── Week helpers ────────────────────────────────────────── */
  function _weekKey(dateStr) {
    var d   = new Date(dateStr + 'T00:00:00');
    var day = d.getDay();
    var diff = (day === 0) ? -6 : 1 - day;
    var mon = new Date(d);
    mon.setDate(d.getDate() + diff);
    return mon.toISOString().slice(0, 10);
  }

  function _weekRange(dates) {
    function fmt(s) { var p = s.split('-'); return p[2] + '.' + p[1]; }
    return dates.length === 1
      ? fmt(dates[0])
      : fmt(dates[0]) + '–' + fmt(dates[dates.length - 1]);
  }

  function _groupByWeek(dates) {
    var groups = {}, order = [];
    dates.forEach(function (d) {
      var wk = _weekKey(d);
      if (!groups[wk]) { groups[wk] = []; order.push(wk); }
      groups[wk].push(d);
    });
    return order.map(function (wk) { return { key: wk, dates: groups[wk] }; });
  }

  /* ── Cell factories ──────────────────────────────────────── */
  function _stickyHdrStyle(leftPx, widthPx) {
    return [
      'position:sticky',
      'left:' + leftPx + 'px',
      'z-index:6',
      'background:#1C1A38',
      'color:#E5E3EB',
      'font-size:19px',
      'font-weight:700',
      'letter-spacing:.04em',
      'text-transform:uppercase',
      'padding:12px 16px',
      'border-bottom:2px solid #2E2A4A',
      'border-right:1px solid #2E2A4A',
      'width:' + widthPx + 'px',
      'min-width:' + widthPx + 'px',
      'white-space:nowrap',
      'text-align:left'
    ].join(';');
  }

  function _stickyTd(text, leftPx, widthPx, bg) {
    var td = document.createElement('td');
    td.textContent = text;
    td.style.cssText = [
      'position:sticky',
      'left:' + leftPx + 'px',
      'z-index:2',
      'background:' + bg,
      'padding:9px 16px',
      'border-bottom:1px solid #1C1A35',
      'border-right:1px solid #2E2A4A',
      'width:' + widthPx + 'px',
      'min-width:' + widthPx + 'px',
      'font-size:22px'
    ].join(';');
    return td;
  }

  function _simpleTd(text, bg, isLastInWeek) {
    var td = document.createElement('td');
    td.textContent = text;
    td.style.cssText = [
      'background:' + bg,
      'padding:7px 5px',
      'text-align:center',
      'border-bottom:1px solid #1C1A35',
      isLastInWeek ? 'border-right:' + WEEK_SEP : ''
    ].join(';');
    return td;
  }

  /* ── Utils ──────────────────────────────────────────────── */
  function _emptyState(text) {
    return U.el('div', {
      class: 'card',
      style: 'color:var(--text-muted);text-align:center;padding:56px 28px',
      text: text
    });
  }

  function _fmtMonth(m) {
    var months = ['Январь','Февраль','Март','Апрель','Май','Июнь',
                  'Июль','Август','Сентябрь','Октябрь','Ноябрь','Декабрь'];
    var parts = m.split('-');
    return (months[parseInt(parts[1], 10) - 1] || m) + ' ' + parts[0];
  }

  return { render: render };
})();
