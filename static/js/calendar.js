window.App = window.App || {};

App.calendar = (function () {
  'use strict';

  var U = App.U;

  function render(container, onDaySelect) {
    container.innerHTML = '';
    var wrap = U.el('div', { class: 'calendar-wrap' });

    var now = new Date();
    var currentMonth = new Date(now.getFullYear(), now.getMonth(), 1);
    var nextMonth = new Date(now.getFullYear(), now.getMonth() + 1, 1);

    wrap.appendChild(_renderMonth(currentMonth, onDaySelect));
    wrap.appendChild(_renderMonth(nextMonth, onDaySelect));

    container.appendChild(wrap);
  }

  function _renderMonth(firstDay, onDaySelect) {
    var DAYS = ['Пн','Вт','Ср','Чт','Пт','Сб','Вс'];
    var MONTHS = ['Январь','Февраль','Март','Апрель','Май','Июнь',
                  'Июль','Август','Сентябрь','Октябрь','Ноябрь','Декабрь'];

    var card = U.el('div', { class: 'cal-card' });
    var title = MONTHS[firstDay.getMonth()] + ' ' + firstDay.getFullYear();
    card.appendChild(U.el('div', { class: 'cal-title', text: title }));

    var grid = U.el('div', { class: 'cal-grid' });
    DAYS.forEach(function (d) {
      grid.appendChild(U.el('div', { class: 'cal-header', text: d }));
    });

    var startDow = firstDay.getDay();
    var offset = (startDow === 0 ? 6 : startDow - 1);
    for (var i = 0; i < offset; i++) {
      grid.appendChild(U.el('div', { class: 'cal-cell empty' }));
    }

    var today = new Date();
    var todayStr = _isoDate(today);
    var selectedDay = App.state.getDay();

    var daysInMonth = new Date(firstDay.getFullYear(), firstDay.getMonth() + 1, 0).getDate();
    for (var d = 1; d <= daysInMonth; d++) {
      var dateObj = new Date(firstDay.getFullYear(), firstDay.getMonth(), d);
      var dateStr = _isoDate(dateObj);
      var isWeekend = (dateObj.getDay() === 0 || dateObj.getDay() === 6);
      var cls = 'cal-cell' + (isWeekend ? ' weekend' : '') +
                (dateStr === todayStr ? ' today' : '') +
                (dateStr === selectedDay ? ' selected' : '');
      var cell = U.el('div', { class: cls, text: String(d), 'data-date': dateStr });
      cell.addEventListener('click', function () {
        App.state.set({ day: this.dataset.date });
        if (onDaySelect) onDaySelect(this.dataset.date);
      });
      grid.appendChild(cell);
    }
    card.appendChild(grid);
    return card;
  }

  function _isoDate(d) {
    return d.getFullYear() + '-' +
           String(d.getMonth() + 1).padStart(2, '0') + '-' +
           String(d.getDate()).padStart(2, '0');
  }

  return { render: render };
})();
