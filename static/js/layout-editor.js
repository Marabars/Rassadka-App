window.App = window.App || {};

App.layoutEditor = (function () {
  'use strict';

  var U = App.U;
  var C = App.C;

  var _layout = [];
  var _svgRoot = null;
  var _container = null;
  var _dragging = null;

  var DESK_W = C.DESK_W;
  var DESK_H = C.DESK_H;
  var SVG_W  = C.SVG_W;
  var SVG_H  = C.SVG_H;

  function render(container) {
    _container = container;
    container.innerHTML = '';
    _layout = JSON.parse(JSON.stringify(
      App.state.get().layout || C.DEFAULT_FLOOR_LAYOUT
    ));
    _drawUI();
  }

  function _drawUI() {
    _container.innerHTML = '';

    var header = U.el('div', { class: 'editor-header' }, [
      U.el('h2', { text: 'Редактор схемы офиса', style: 'margin:0' }),
      U.el('span', { class: 'muted', text: 'Перетаскивайте рабочие места. Двойной клик — удалить. Кнопка «+» — добавить.' })
    ]);
    _container.appendChild(header);

    var toolbar = U.el('div', { class: 'editor-toolbar' });

    var addForm = U.el('form', { class: 'add-desk-form' }, [
      U.el('input', { type: 'text', id: 'new-desk-id', placeholder: 'ID рабочего места',
                      required: 'true', style: 'width:140px' }),
      _zoneSelect('new-desk-zone'),
      U.el('button', { type: 'submit', class: 'btn btn-primary', text: '+ Добавить место' })
    ]);
    addForm.addEventListener('submit', function (e) {
      e.preventDefault();
      var idInput = U.qs('#new-desk-id');
      var zoneInput = U.qs('#new-desk-zone');
      var newId = idInput.value.trim();
      if (!newId) return;
      if (_layout.find(function (d) { return d.id === newId; })) {
        alert('Место с таким ID уже существует.');
        return;
      }
      _layout.push({ id: newId, x: 50, y: 50, zone: zoneInput.value });
      idInput.value = '';
      _redraw();
    });
    toolbar.appendChild(addForm);

    var saveBtn = U.el('button', { class: 'btn btn-primary', text: 'Сохранить схему',
                                   style: 'margin-left:auto' });
    var saveStatus = U.el('span', { class: 'muted save-status' });
    saveBtn.addEventListener('click', function () {
      saveStatus.textContent = 'Сохраняю…';
      App.api.saveLayout(_layout).then(function (res) {
        App.state.set({ layout: JSON.parse(JSON.stringify(_layout)) });
        saveStatus.textContent = 'Сохранено (' + res.count + ' мест).';
      }).catch(function (e) {
        saveStatus.textContent = 'Ошибка: ' + e.message;
      });
    });

    var resetBtn = U.el('button', { class: 'btn btn-secondary', text: 'Сбросить к умолчанию',
                                    style: 'margin-left:8px' });
    resetBtn.addEventListener('click', function () {
      if (!confirm('Сбросить схему к исходному варианту из фото офиса?')) return;
      _layout = JSON.parse(JSON.stringify(C.DEFAULT_FLOOR_LAYOUT));
      _redraw();
    });

    toolbar.appendChild(saveBtn);
    toolbar.appendChild(saveStatus);
    toolbar.appendChild(resetBtn);
    _container.appendChild(toolbar);

    _container.appendChild(U.el('p', { class: 'muted editor-hint',
      text: 'Двойной клик по месту — удалить его. Перетащите место на новую позицию. Изменения сохраняются кнопкой «Сохранить схему».' }));

    var canvasWrap = U.el('div', { class: 'editor-canvas-wrap' });
    var svg = U.svgEl('svg', {
      width: SVG_W, height: SVG_H,
      viewBox: '0 0 ' + SVG_W + ' ' + SVG_H,
      class: 'floor-svg editor-svg'
    });
    _svgRoot = svg;
    svg.addEventListener('mousemove', _onMouseMove);
    svg.addEventListener('mouseup', _onMouseUp);
    svg.addEventListener('mouseleave', _onMouseUp);
    _drawDesks();
    canvasWrap.appendChild(svg);
    _container.appendChild(canvasWrap);
  }

  function _redraw() {
    if (!_svgRoot) return;
    _svgRoot.innerHTML = '';
    _drawDesks();
  }

  function _drawDesks() {
    _layout.forEach(function (desk, idx) {
      var g = U.svgEl('g', { class: 'desk-group editor-desk', 'data-idx': String(idx) });

      var rect = U.svgEl('rect', {
        x: desk.x, y: desk.y, width: DESK_W, height: DESK_H,
        rx: 4, ry: 4, fill: '#e8f5e9', stroke: '#43a047', 'stroke-width': 1.5
      });
      g.appendChild(rect);

      var label = U.svgEl('text', {
        x: desk.x + DESK_W / 2, y: desk.y + DESK_H / 2 + 5,
        'text-anchor': 'middle', 'font-size': '12',
        fill: '#212121', 'font-family': 'sans-serif', 'font-weight': '600'
      });
      label.textContent = desk.id;
      g.appendChild(label);

      var zoneLabel = U.svgEl('text', {
        x: desk.x + DESK_W / 2, y: desk.y + DESK_H - 6,
        'text-anchor': 'middle', 'font-size': '9',
        fill: '#757575', 'font-family': 'sans-serif'
      });
      zoneLabel.textContent = desk.zone;
      g.appendChild(zoneLabel);

      g.style.cursor = 'move';
      g.addEventListener('mousedown', function (e) {
        e.preventDefault();
        _dragging = { idx: idx, startX: e.clientX, startY: e.clientY,
                      origX: desk.x, origY: desk.y };
      });

      g.addEventListener('dblclick', function () {
        if (!confirm('Удалить рабочее место «' + desk.id + '»?')) return;
        _layout.splice(idx, 1);
        _dirty = true;
        _redraw();
      });

      _svgRoot.appendChild(g);
    });
  }

  function _onMouseMove(e) {
    if (!_dragging) return;
    var svgRect = _svgRoot.getBoundingClientRect();
    var scaleX = SVG_W / svgRect.width;
    var scaleY = SVG_H / svgRect.height;
    var dx = (e.clientX - _dragging.startX) * scaleX;
    var dy = (e.clientY - _dragging.startY) * scaleY;
    var newX = Math.max(0, Math.min(SVG_W - DESK_W, _dragging.origX + dx));
    var newY = Math.max(0, Math.min(SVG_H - DESK_H, _dragging.origY + dy));
    _layout[_dragging.idx].x = Math.round(newX);
    _layout[_dragging.idx].y = Math.round(newY);
    _redraw();
  }

  function _onMouseUp() {
    if (_dragging) {
      _dragging = null;
    }
  }

  function _zoneSelect(id) {
    var sel = U.el('select', { id: id, style: 'margin:0 8px; padding:6px; border:1px solid #ccc; border-radius:4px' });
    ['top', 'mid-top', 'mid', 'right', 'bottom', 'custom'].forEach(function (z) {
      sel.appendChild(U.el('option', { value: z, text: z }));
    });
    return sel;
  }

  return { render: render };
})();
