window.App = window.App || {};

App.seatingList = (function () {
  'use strict';

  var U = App.U;
  var _container = null;
  var _onSelect = null;

  function render(container, onEmployeeSelect) {
    _container = container;
    _onSelect = onEmployeeSelect;
    _draw();
  }

  function refresh() {
    if (_container) _draw();
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

    var officeEmps = employees.filter(function (e) { return e.status === 'OFFICE'; });
    var remoteEmps = employees.filter(function (e) { return e.status !== 'OFFICE'; });

    if (officeEmps.length) {
      _container.appendChild(U.el('div', { class: 'list-section-title', text: 'В офисе (' + officeEmps.length + ')' }));
      officeEmps.forEach(function (e) { _container.appendChild(_chip(e)); });
    }
    if (remoteEmps.length) {
      _container.appendChild(U.el('div', { class: 'list-section-title', text: 'Вне офиса (' + remoteEmps.length + ')' }));
      remoteEmps.forEach(function (e) { _container.appendChild(_chip(e)); });
    }
  }

  function _chip(emp) {
    var statusLabel = { OFFICE: '🏢', REMOTE: '🏠', VACATION: '🌴', DAY_OFF: '💤' };
    var seatBadge = emp.seat_id ? ' → ' + emp.seat_id : '';
    var chip = U.el('div', {
      class: 'emp-chip',
      'data-employee': emp.name,
      'data-from-seat': emp.seat_id || '',
      draggable: emp.status === 'OFFICE' ? 'true' : 'false'
    }, [
      U.el('span', { class: 'emp-status', text: statusLabel[emp.status] || '?' }),
      U.el('span', { class: 'emp-name', text: emp.name }),
      U.el('span', { class: 'emp-seat muted', text: seatBadge })
    ]);

    chip.addEventListener('click', function () {
      App.state.set({ selectedEmployee: emp.name, selectedSeat: emp.seat_id });
      App.floorPlan.refresh();
      if (_onSelect) _onSelect(emp);
    });
    return chip;
  }

  return { render: render, refresh: refresh };
})();
