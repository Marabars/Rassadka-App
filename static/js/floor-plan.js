window.App = window.App || {};

App.floorPlan = (function () {
  'use strict';

  var U = App.U;
  var C = App.C;
  var _opts = {};
  var _svgRoot = null;

  function render(container, opts) {
    _opts = opts || {};
    container.innerHTML = '';

    var svg = U.svgEl('svg', {
      width: C.SVG_W, height: C.SVG_H,
      viewBox: '0 0 ' + C.SVG_W + ' ' + C.SVG_H,
      class: 'floor-svg'
    });
    _svgRoot = svg;
    _drawDesks(svg);
    container.appendChild(svg);
    setTimeout(function () { if (App.dragDrop) App.dragDrop.init(); }, 0);
  }

  function refresh() {
    if (!_svgRoot) return;
    _svgRoot.innerHTML = '';
    _drawDesks(_svgRoot);
    setTimeout(function () { if (App.dragDrop) App.dragDrop.init(); }, 0);
  }

  function _drawDesks(svg) {
    var layout = App.state.get().layout || C.DEFAULT_FLOOR_LAYOUT;
    var byId = App.state.assignmentsBySeat();
    var selected = App.state.get().selectedSeat;
    var W = C.DESK_W, H = C.DESK_H;

    layout.forEach(function (desk) {
      var assignment = byId[desk.id];
      var fill = assignment ? C.STATUS_COLOR[assignment.status] || C.STATUS_COLOR.FREE
                            : C.STATUS_COLOR.FREE;
      var isSelected = desk.id === selected;

      var g = U.svgEl('g', {
        class: 'desk-group' + (isSelected ? ' selected' : ''),
        'data-seat': desk.id
      });

      var isLight = (fill === C.STATUS_COLOR.OFFICE || fill === C.STATUS_COLOR.VACATION);
      var idFill   = isLight ? '#1A3030' : '#7E7A9A';
      var nameFill = isLight ? '#0E2424' : '#E5E3EB';

      var rect = U.svgEl('rect', {
        x: desk.x, y: desk.y, width: W, height: H,
        rx: 4, ry: 4,
        fill: fill,
        stroke: isSelected ? '#7549E8' : '#2E2A4A',
        'stroke-width': isSelected ? 2.5 : 1
      });
      g.appendChild(rect);

      var idLabel = U.svgEl('text', {
        x: desk.x + W / 2, y: desk.y + 16,
        'text-anchor': 'middle', 'font-size': '11',
        fill: idFill, 'font-family': 'sans-serif'
      });
      idLabel.textContent = desk.id;
      g.appendChild(idLabel);

      if (assignment && assignment.employee_name) {
        var nameLabel = U.svgEl('text', {
          x: desk.x + W / 2, y: desk.y + 34,
          'text-anchor': 'middle', 'font-size': '10',
          fill: nameFill, 'font-weight': '600', 'font-family': 'sans-serif'
        });
        nameLabel.textContent = _shorten(assignment.employee_name);
        g.appendChild(nameLabel);
      }

      var title = U.svgEl('title', {});
      title.textContent = desk.id + (assignment ? ': ' + assignment.employee_name : ': свободно');
      g.appendChild(title);

      g.style.cursor = 'pointer';
      g.addEventListener('click', function () {
        App.state.set({ selectedSeat: desk.id });
        if (_opts.onSeatClick) _opts.onSeatClick(desk.id);
        refresh();
      });

      svg.appendChild(g);
    });
  }

  function _shorten(fullName) {
    var parts = fullName.split(' ');
    if (parts.length < 2) return fullName.slice(0, 12);
    return parts[0] + ' ' + (parts[1] ? parts[1][0] + '.' : '');
  }

  return { render: render, refresh: refresh };
})();
