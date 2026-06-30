window.App = window.App || {};

App.seatingList = (function () {
  'use strict';

  var U = App.U;
  var _container = null;
  var _onSelect = null;
  var _state = {
    search: '',
    statuses: { OFFICE: true, REMOTE: true, VACATION: true, DAY_OFF: true },
    collapsed: { office: false, other: false }
  };

  var STATUS_LABEL = { OFFICE: 'В офисе', REMOTE: 'Удалённо', VACATION: 'Отпуск', DAY_OFF: 'Выходной' };
  var STATUS_ICON  = { OFFICE: '🏢', REMOTE: '🏠', VACATION: '🌴', DAY_OFF: '💤' };
  var STATUS_COLOR = { OFFICE: '#82D6CC', REMOTE: '#9B8EC4', VACATION: '#BD9375', DAY_OFF: '#5A5580' };

  function render(container, onEmployeeSelect) {
    _container = container;
    _onSelect = onEmployeeSelect;
    _draw();
  }

  function refresh() {
    if (_container) _draw();
  }

  function scrollTo(employeeName) {
    if (!_container || !employeeName) return;
    var chips = _container.querySelectorAll('.emp-chip');
    for (var i = 0; i < chips.length; i++) {
      if (chips[i].getAttribute('data-employee') === employeeName) {
        chips[i].scrollIntoView({ behavior: 'smooth', block: 'nearest' });
        break;
      }
    }
  }

  function _draw() {
    _container.innerHTML = '';
    var day = App.state.getDay();

    if (!day) {
      _container.appendChild(U.el('p', { class: 'muted', text: 'Выберите день на календаре.' }));
      return;
    }

    var employees = App.state.get().employees;
    if (!employees || employees.length === 0) {
      _container.appendChild(U.el('p', { class: 'muted', text: 'Нет данных для этого дня. Сгенерируйте рассадку.' }));
      return;
    }

    _container.appendChild(_toolbar());

    var search = _state.search.toLowerCase();

    function matchesFilter(e) {
      if (!_state.statuses[e.status]) return false;
      if (search && e.name.toLowerCase().indexOf(search) === -1) return false;
      return true;
    }

    var officeEmps = employees.filter(function (e) { return e.status === 'OFFICE' && matchesFilter(e); });
    var otherEmps  = employees.filter(function (e) { return e.status !== 'OFFICE' && matchesFilter(e); });

    if (officeEmps.length || !search) {
      _container.appendChild(_section('office', 'В офисе', officeEmps));
    }
    if (otherEmps.length || !search) {
      _container.appendChild(_section('other', 'Вне офиса', otherEmps));
    }
  }

  function _toolbar() {
    var bar = U.el('div', { class: 'sl-toolbar' });

    var searchEl = U.el('input', {
      class: 'sl-search',
      type: 'text',
      placeholder: 'Поиск по ФИО…',
      value: _state.search
    });
    searchEl.addEventListener('input', function (e) {
      _state.search = e.target.value;
      _draw();
    });
    bar.appendChild(searchEl);

    var filterRow = U.el('div', { class: 'sl-filters' });
    ['OFFICE', 'REMOTE', 'VACATION', 'DAY_OFF'].forEach(function (st) {
      var btn = U.el('button', {
        class: 'sl-filter-btn' + (_state.statuses[st] ? ' active' : ''),
        text: STATUS_LABEL[st],
        style: _state.statuses[st] ? 'border-color:' + STATUS_COLOR[st] + ';color:' + STATUS_COLOR[st] : ''
      });
      btn.addEventListener('click', function () {
        _state.statuses[st] = !_state.statuses[st];
        _draw();
      });
      filterRow.appendChild(btn);
    });
    bar.appendChild(filterRow);
    return bar;
  }

  function _section(key, title, emps) {
    var wrap = U.el('div', { class: 'sl-section' });

    var hdr = U.el('div', { class: 'sl-section-hdr' });
    var collapsed = _state.collapsed[key];
    var arrow = U.el('span', { class: 'sl-arrow', text: collapsed ? '▶' : '▼' });
    var label = U.el('span', { text: title + ' (' + emps.length + ')' });
    hdr.appendChild(arrow);
    hdr.appendChild(label);
    hdr.addEventListener('click', function () {
      _state.collapsed[key] = !_state.collapsed[key];
      _draw();
    });
    wrap.appendChild(hdr);

    if (!collapsed) {
      if (emps.length === 0) {
        wrap.appendChild(U.el('p', { class: 'muted', style: 'padding:4px 0 4px 8px;font-size:var(--fs-sm)', text: 'Нет совпадений' }));
      } else {
        emps.forEach(function (e) { wrap.appendChild(_chip(e)); });
      }
    }
    return wrap;
  }

  function _chip(emp) {
    var st = App.state.get();
    var isSelected = (st.selectedEmployee === emp.name) ||
                     (emp.seat_id && st.selectedSeat === emp.seat_id);

    var children = [
      U.el('span', { class: 'emp-status', text: STATUS_ICON[emp.status] || '?' }),
      U.el('span', { class: 'emp-name', text: emp.name })
    ];
    if (emp.seat_id) {
      children.push(U.el('span', { class: 'emp-seat', text: emp.seat_id }));
    }

    var chip = U.el('div', {
      class: 'emp-chip' + (isSelected ? ' emp-chip--selected' : ''),
      'data-employee': emp.name,
      'data-from-seat': emp.seat_id || '',
      draggable: emp.status === 'OFFICE' ? 'true' : 'false'
    }, children);

    chip.addEventListener('click', function () {
      App.state.set({ selectedEmployee: emp.name, selectedSeat: emp.seat_id });
      App.floorPlan.refresh();
      if (_onSelect) _onSelect(emp);
    });
    return chip;
  }

  return { render: render, refresh: refresh, scrollTo: scrollTo };
})();
