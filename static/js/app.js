window.App = window.App || {};

App.seatingView = (function () {
  'use strict';
  var U = App.U;

  function render(container) {
    container.innerHTML = '';

    var calRow = U.el('div', { class: 'seating-top-row' });
    App.calendar.render(calRow, function (date) {
      _onDaySelected(date);
    });
    container.appendChild(calRow);

    var legend = U.el('div', { class: 'floor-legend' });
    [['В офисе', '#82D6CC'], ['Удалённо', '#3B3760'],
     ['Отпуск', '#BD9375'], ['Свободно', '#1C1A35']].forEach(function (pair) {
      legend.appendChild(U.el('div', { class: 'legend-item' }, [
        U.el('div', { class: 'legend-dot', style: 'background:' + pair[1] }),
        U.el('span', { text: pair[0] })
      ]));
    });
    container.appendChild(legend);

    var mainArea = U.el('div', { class: 'seating-main-area' });

    var planContainer = U.el('div', { class: 'floor-plan-container' });
    App.floorPlan.render(planContainer, {
      onSeatClick: function (seatId, empName) {
        App.state.set({ selectedSeat: seatId, selectedEmployee: empName || null });
        App.seatingList.refresh();
        if (empName) App.seatingList.scrollTo(empName);
      }
    });
    mainArea.appendChild(planContainer);

    var listContainer = U.el('div', { class: 'emp-list-container card' });
    listContainer.appendChild(U.el('h3', { text: 'Сотрудники', style: 'margin-bottom:10px' }));
    var listBody = U.el('div', { class: 'emp-list-body' });
    App.seatingList.render(listBody, null);
    listContainer.appendChild(listBody);
    mainArea.appendChild(listContainer);

    container.appendChild(mainArea);
  }

  function _onDaySelected(date) {
    App.state.set({ day: date });
    App.api.getSeating(App.state.getMonth(), date).then(function (data) {
      App.state.set({ assignments: data.assignments });
      return App.api.getEmployees(date);
    }).then(function (data) {
      App.state.set({ employees: data.employees });
      App.floorPlan.refresh();
      App.seatingList.refresh();
    });
    var calRow = U.qs('.seating-top-row');
    if (calRow) App.calendar.render(calRow, _onDaySelected);
  }

  return { render: render };
})();

(function () {
  'use strict';

  var U = App.U;

  function switchTab(tabId) {
    App.state.set({ tab: tabId });
    U.qsa('.nav-tab').forEach(function (btn) {
      btn.classList.toggle('active', btn.dataset.tab === tabId);
    });
    render();
  }

  function render() {
    var container = U.qs('#app-main');
    container.innerHTML = '';
    var tab = App.state.get().tab;
    if (tab === 'upload') {
      App.upload.render(container);
    } else if (tab === 'seating') {
      App.seatingView.render(container);
    } else if (tab === 'schedule') {
      App.schedule.render(container);
    } else if (tab === 'layout') {
      App.layoutEditor.render(container);
    }
  }

  function init() {
    U.qsa('.nav-tab').forEach(function (btn) {
      btn.addEventListener('click', function () { switchTab(btn.dataset.tab); });
    });
    App.api.getLayout().then(function (data) {
      App.state.set({ layout: data.layout });
      render();
    }).catch(function () {
      App.state.set({ layout: App.C.DEFAULT_FLOOR_LAYOUT });
      render();
    });
  }

  App.app = { switchTab: switchTab, render: render };
  document.addEventListener('DOMContentLoaded', init);
})();
