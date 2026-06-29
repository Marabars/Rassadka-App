window.App = window.App || {};

App.schedule = (function () {
  'use strict';

  var U = App.U;

  function render(container) {
    container.innerHTML = '';

    var month = App.state.getMonth();

    var header = U.el('div', {
      style: 'display:flex;align-items:center;justify-content:space-between;margin-bottom:24px;flex-wrap:wrap;gap:16px'
    });

    header.appendChild(U.el('div', {}, [
      U.el('h2', { text: 'График рассадки', style: 'margin:0 0 4px' }),
      U.el('div', { class: 'muted', text: month ? _formatMonth(month) : 'Месяц не выбран' })
    ]));

    var exportBtn = U.el('button', { class: 'btn btn-secondary', text: '↓ Скачать xlsx' });
    if (!month) exportBtn.disabled = true;
    exportBtn.addEventListener('click', function () {
      if (month) window.location.href = '/api/seating/export?month=' + encodeURIComponent(month);
    });
    header.appendChild(exportBtn);
    container.appendChild(header);

    if (!month) {
      container.appendChild(_emptyState('Сначала выберите месяц и сгенерируйте рассадку на вкладке «Загрузка файлов».'));
      return;
    }

    var loadingEl = U.el('div', { class: 'muted', style: 'padding:28px 0', text: 'Загружаю…' });
    container.appendChild(loadingEl);

    App.api.getSeating(month).then(function (data) {
      container.removeChild(loadingEl);
      _renderTable(container, data.assignments);
    }).catch(function (err) {
      loadingEl.textContent = 'Ошибка: ' + err.message;
    });
  }

  function _renderTable(container, assignments) {
    if (!assignments || assignments.length === 0) {
      container.appendChild(_emptyState('Нет данных для этого месяца. Сгенерируйте рассадку на вкладке «Загрузка файлов».'));
      return;
    }

    var empData = {};
    var allDates = {};
    assignments.forEach(function (a) {
      allDates[a.date] = true;
      if (!empData[a.employee_name]) empData[a.employee_name] = {};
      if (a.status === 'OFFICE') {
        empData[a.employee_name][a.date] = a.seat_id || '';
      }
    });

    var sortedDates = Object.keys(allDates).sort();
    var sortedNames = Object.keys(empData).sort();

    var officeCounts = {};
    sortedDates.forEach(function (d) {
      var cnt = 0;
      sortedNames.forEach(function (n) { if (empData[n][d]) cnt++; });
      officeCounts[d] = cnt;
    });

    var wrap = U.el('div', { style: 'overflow-x:auto' });
    var table = document.createElement('table');
    table.style.cssText = 'border-collapse:collapse;font-size:24px;min-width:max-content;width:100%';

    // === THEAD ===
    var thead = document.createElement('thead');

    var trHead = document.createElement('tr');
    trHead.appendChild(_th('ФИО', true, 'left', '280px'));
    sortedDates.forEach(function (date) {
      var parts = date.split('-');
      trHead.appendChild(_th(parts[2] + '.' + parts[1], false, 'center', '72px', date));
    });
    thead.appendChild(trHead);

    var trCounts = document.createElement('tr');
    var tdLabel = document.createElement('td');
    tdLabel.textContent = 'В офисе';
    tdLabel.style.cssText = [
      'background:#0A0814',
      'color:#4A4668',
      'font-size:18px',
      'font-weight:600',
      'letter-spacing:.06em',
      'text-transform:uppercase',
      'padding:6px 20px',
      'border-bottom:2px solid #2E2A4A',
      'position:sticky',
      'left:0',
      'z-index:3'
    ].join(';');
    trCounts.appendChild(tdLabel);
    sortedDates.forEach(function (date) {
      var cnt = officeCounts[date];
      var td = document.createElement('td');
      td.style.cssText = 'background:#0A0814;text-align:center;padding:6px 5px;border-bottom:2px solid #2E2A4A;font-size:22px;font-weight:600';
      td.style.color = cnt > 0 ? '#82D6CC' : '#2E2A4A';
      td.textContent = cnt > 0 ? String(cnt) : '—';
      trCounts.appendChild(td);
    });
    thead.appendChild(trCounts);
    table.appendChild(thead);

    // === TBODY ===
    var tbody = document.createElement('tbody');
    sortedNames.forEach(function (name, i) {
      var rowBg = i % 2 === 0 ? '#13112A' : '#0F0D20';
      var tr = document.createElement('tr');

      var tdName = document.createElement('td');
      tdName.textContent = name;
      tdName.style.cssText = [
        'padding:12px 20px',
        'color:#E5E3EB',
        'white-space:nowrap',
        'border-bottom:1px solid #1C1A35',
        'background:' + rowBg,
        'position:sticky',
        'left:0',
        'z-index:1'
      ].join(';');
      tr.appendChild(tdName);

      sortedDates.forEach(function (date) {
        var seat = empData[name][date];
        var td = document.createElement('td');
        td.style.cssText = 'padding:8px 5px;text-align:center;border-bottom:1px solid #1C1A35;background:' + rowBg;
        if (seat) {
          var chip = document.createElement('span');
          chip.textContent = seat;
          chip.style.cssText = [
            'display:inline-block',
            'padding:3px 8px',
            'border-radius:4px',
            'background:rgba(130,214,204,.1)',
            'color:#82D6CC',
            'font-family:Consolas,"Courier New",monospace',
            'font-size:22px',
            'font-weight:700',
            'letter-spacing:.02em'
          ].join(';');
          td.appendChild(chip);
        }
        tr.appendChild(td);
      });

      tbody.appendChild(tr);
    });
    table.appendChild(tbody);
    wrap.appendChild(table);
    container.appendChild(wrap);
  }

  function _th(text, isFirst, align, minWidth, title) {
    var th = document.createElement('th');
    th.textContent = text;
    if (title) th.title = title;
    var sticky = isFirst ? 'position:sticky;left:0;z-index:3' : 'z-index:2';
    th.style.cssText = [
      'background:#13112A',
      'color:#7E7A9A',
      'font-size:20px',
      'font-weight:600',
      'letter-spacing:.05em',
      'text-transform:uppercase',
      'border-bottom:1px solid #2E2A4A',
      'padding:12px ' + (isFirst ? '20px' : '5px'),
      'min-width:' + minWidth,
      'text-align:' + align,
      sticky,
      'position:sticky',
      'top:0'
    ].join(';');
    return th;
  }

  function _emptyState(text) {
    return U.el('div', {
      class: 'card',
      style: 'color:var(--text-muted);text-align:center;padding:56px 28px',
      text: text
    });
  }

  function _formatMonth(m) {
    var months = ['январь','февраль','март','апрель','май','июнь','июль','август','сентябрь','октябрь','ноябрь','декабрь'];
    var parts = m.split('-');
    var mon = months[parseInt(parts[1], 10) - 1] || m;
    return mon.charAt(0).toUpperCase() + mon.slice(1) + ' ' + parts[0];
  }

  return { render: render };
})();