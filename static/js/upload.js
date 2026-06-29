window.App = window.App || {};

App.upload = (function () {
  'use strict';

  var U = App.U;
  var api = App.api;
  var state = App.state;

  var _choicesId = null;
  var _templateId = null;

  function render(container) {
    container.innerHTML = '';
    var card = U.el('div', { class: 'card upload-card' });

    card.appendChild(U.el('h2', { text: 'Загрузка файлов', style: 'margin-bottom:16px' }));

    var monthRow = U.el('div', { class: 'form-row' });
    monthRow.appendChild(U.el('label', { text: 'Месяц рассадки:', for: 'month-input' }));
    var monthInput = U.el('input', { type: 'month', id: 'month-input', value: state.getMonth() });
    monthInput.addEventListener('change', function () {
      state.set({ month: monthInput.value });
    });
    monthRow.appendChild(monthInput);
    card.appendChild(monthRow);

    card.appendChild(_fileRow('choices', 'Файл выборов сотрудников (.xlsx):',
      function (file) {
        return api.uploadFile('/api/upload/choices', file).then(function (r) {
          _choicesId = r.upload_id;
          return 'Загружен: ' + r.filename;
        });
      }
    ));

    card.appendChild(_fileRow('template', 'Шаблон прошлой рассадки (.xlsx):',
      function (file) {
        return api.uploadFile('/api/upload/template', file).then(function (r) {
          _templateId = r.upload_id;
          return 'Загружен: ' + r.filename;
        });
      }
    ));

    var genBtn = U.el('button', { class: 'btn btn-primary', text: 'Сгенерировать рассадку',
                                   style: 'margin-top:16px' });
    var statusEl = U.el('div', { class: 'upload-status', style: 'margin-top:12px' });
    genBtn.addEventListener('click', function () {
      if (!_choicesId || !_templateId) {
        statusEl.textContent = 'Загрузите оба файла перед генерацией.';
        statusEl.className = 'upload-status badge-warning';
        return;
      }
      statusEl.textContent = 'Генерирую…';
      statusEl.className = 'upload-status';
      api.generate(state.getMonth(), _choicesId, _templateId).then(function (res) {
        state.set({ lastGenResult: res });
        statusEl.innerHTML = 'Готово. Назначено мест: <b>' + res.assigned_count + '</b>. ' +
          'Ошибок: <b>' + res.error_count + '</b>. ' +
          'Всего проблем: <b>' + res.issues_count + '</b>.';
        statusEl.className = 'upload-status badge-info';
      }).catch(function (e) {
        statusEl.textContent = 'Ошибка: ' + e.message;
        statusEl.className = 'upload-status badge-error';
      });
    });
    card.appendChild(genBtn);
    card.appendChild(statusEl);
    container.appendChild(card);
  }

  function _fileRow(id, labelText, onUpload) {
    var row = U.el('div', { class: 'form-row', style: 'margin-top:16px' });
    row.appendChild(U.el('label', { text: labelText, for: id + '-file' }));
    var input = U.el('input', { type: 'file', id: id + '-file', accept: '.xlsx' });
    var status = U.el('span', { class: 'muted', style: 'margin-left:8px' });
    input.addEventListener('change', function () {
      if (!input.files[0]) return;
      status.textContent = 'Загружаю…';
      onUpload(input.files[0]).then(function (msg) {
        status.textContent = msg;
      }).catch(function (e) {
        status.textContent = 'Ошибка: ' + e.message;
      });
    });
    var wrap = U.el('div', {}, [input, status]);
    row.appendChild(wrap);
    return row;
  }

  return { render: render };
})();
