window.App = window.App || {};

App.dragDrop = (function () {
  'use strict';

  var _draggedEmployee = null;   // from list chip HTML5 drag
  var _ghost = null;             // custom ghost for SVG-seat drag
  var _fromSeatDragging = false; // true while mouse-drag from SVG seat is active

  /* ---- public ---- */
  function fromSeatDragging() { return _fromSeatDragging; }

  function init() {
    _wireEmployeeChips();
    _wireSeatTargets();
    _wireDropZones();
    _wireSeatMouseDrag();
  }

  /* ------------------------------------------------------------------ */
  /* HTML5 drag: list chip → SVG seat                                    */
  /* ------------------------------------------------------------------ */
  function _wireEmployeeChips() {
    document.querySelectorAll('[data-employee]').forEach(function (chip) {
      chip.draggable = true;
      chip.addEventListener('dragstart', function (e) {
        _draggedEmployee = chip.dataset.employee;
        e.dataTransfer.effectAllowed = 'move';
        e.dataTransfer.setData('text/plain', _draggedEmployee);
        chip.classList.add('dragging');
        document.querySelectorAll('[data-drop-zone]').forEach(function (z) {
          z.classList.add('drop-zone-visible');
        });
      });
      chip.addEventListener('dragend', function () {
        chip.classList.remove('dragging');
        document.querySelectorAll('[data-drop-zone]').forEach(function (z) {
          z.classList.remove('drop-zone-visible', 'drag-over');
        });
      });
    });
  }

  function _wireSeatTargets() {
    document.querySelectorAll('[data-seat]').forEach(function (desk) {
      desk.addEventListener('dragover', function (e) {
        e.preventDefault();
        e.dataTransfer.dropEffect = 'move';
        desk.classList.add('drag-over');
      });
      desk.addEventListener('dragleave', function () {
        desk.classList.remove('drag-over');
      });
      desk.addEventListener('drop', function (e) {
        e.preventDefault();
        desk.classList.remove('drag-over');
        var seatId = desk.dataset.seat;
        var employee = _draggedEmployee;
        if (!employee || !seatId) return;
        var date = App.state.getDay();
        if (!date) { alert('Выберите день на календаре перед перемещением.'); return; }
        App.api.overrideSeat(date, seatId, employee).then(function () {
          return App.api.getSeating(App.state.getMonth(), date);
        }).then(function (data) {
          App.state.set({ assignments: data.assignments });
          App.floorPlan.refresh();
          App.seatingList.refresh();
        }).catch(function (err) { console.error('Drop failed:', err); });
        _draggedEmployee = null;
      });
    });
  }

  function _wireDropZones() {
    document.querySelectorAll('[data-drop-zone]').forEach(function (zone) {
      zone.addEventListener('dragover', function (e) {
        if (!_draggedEmployee) return;
        e.preventDefault();
        e.dataTransfer.dropEffect = 'move';
        zone.classList.add('drag-over');
      });
      zone.addEventListener('dragleave', function (e) {
        if (!zone.contains(e.relatedTarget)) zone.classList.remove('drag-over');
      });
      zone.addEventListener('drop', function (e) {
        e.preventDefault();
        zone.classList.remove('drag-over');
        var employee = _draggedEmployee;
        if (!employee) return;
        var date = App.state.getDay();
        if (!date) return;
        var targetStatus = zone.dataset.dropZone === 'other' ? 'REMOTE' : null;
        if (!targetStatus) return;
        App.api.setEmployeeStatus(date, employee, targetStatus).then(function () {
          return App.api.getSeating(App.state.getMonth(), date);
        }).then(function (data) {
          App.state.set({ assignments: data.assignments });
          return App.api.getEmployees(date);
        }).then(function (data) {
          App.state.set({ employees: data.employees });
          App.floorPlan.refresh();
          App.seatingList.refresh();
        }).catch(function (err) { console.error('Drop zone failed:', err); });
        _draggedEmployee = null;
      });
    });
  }

  /* ------------------------------------------------------------------ */
  /* Custom mouse drag: SVG seat → drop zone / another seat              */
  /* ------------------------------------------------------------------ */
  function _wireSeatMouseDrag() {
    document.querySelectorAll('.desk-group[data-seat]').forEach(function (desk) {
      desk.addEventListener('mousedown', function (e) {
        var sid = desk.dataset.seat;
        var byId = App.state.assignmentsBySeat();
        var asgn = byId[sid];
        if (!asgn || !asgn.employee_name || asgn.status !== 'OFFICE') return;

        var startX = e.clientX, startY = e.clientY;
        var empName = asgn.employee_name;
        var dragging = false;

        function onMove(e) {
          var dx = e.clientX - startX, dy = e.clientY - startY;
          if (!dragging && Math.sqrt(dx * dx + dy * dy) > 6) {
            dragging = true;
            _ghostCreate(empName, e.clientX, e.clientY);
            document.querySelectorAll('[data-drop-zone]').forEach(function (z) {
              z.classList.add('drop-zone-visible');
            });
          }
          if (dragging) {
            _ghostMove(e.clientX, e.clientY);
            _highlightZones(e.clientX, e.clientY);
          }
        }

        function onUp(e) {
          document.removeEventListener('mousemove', onMove);
          document.removeEventListener('mouseup', onUp);

          if (!dragging) return;

          _fromSeatDragging = true;
          _ghost.style.display = 'none';
          var el = document.elementFromPoint(e.clientX, e.clientY);
          _ghost.style.display = '';
          _ghostDestroy();

          document.querySelectorAll('[data-drop-zone]').forEach(function (z) {
            z.classList.remove('drop-zone-visible', 'drag-over');
          });

          var date = App.state.getDay();
          if (!date) { _fromSeatDragging = false; return; }

          var zone = el && el.closest('[data-drop-zone]');
          var seatEl = el && el.closest('[data-seat]');

          var p;
          if (zone && zone.dataset.dropZone === 'other') {
            p = App.api.setEmployeeStatus(date, empName, 'REMOTE').then(function () {
              return App.api.getSeating(App.state.getMonth(), date);
            }).then(function (data) {
              App.state.set({ assignments: data.assignments });
              return App.api.getEmployees(date);
            }).then(function (data) {
              App.state.set({ employees: data.employees });
              App.floorPlan.refresh();
              App.seatingList.refresh();
            });
          } else if (seatEl && seatEl !== desk) {
            var targetSeat = seatEl.dataset.seat;
            p = App.api.overrideSeat(date, targetSeat, empName).then(function () {
              return App.api.getSeating(App.state.getMonth(), date);
            }).then(function (data) {
              App.state.set({ assignments: data.assignments });
              App.floorPlan.refresh();
              App.seatingList.refresh();
            });
          }
          if (p) p.catch(function (err) { console.error('Seat drag failed:', err); });

          setTimeout(function () { _fromSeatDragging = false; }, 0);
        }

        document.addEventListener('mousemove', onMove);
        document.addEventListener('mouseup', onUp);
      });
    });
  }

  /* ---- ghost helpers ---- */
  function _ghostCreate(name, x, y) {
    _ghost = document.createElement('div');
    _ghost.style.cssText = [
      'position:fixed', 'z-index:10000', 'pointer-events:none',
      'background:#7549E8', 'color:#fff', 'border-radius:8px',
      'padding:6px 16px', 'font-size:18px',
      'font-family:Calibri,"Helvetica Neue",Helvetica,sans-serif',
      'box-shadow:0 4px 16px rgba(0,0,0,.6)', 'white-space:nowrap', 'opacity:.9'
    ].join(';');
    _ghost.textContent = name;
    document.body.appendChild(_ghost);
    _ghostMove(x, y);
  }

  function _ghostMove(x, y) {
    if (!_ghost) return;
    _ghost.style.left = (x + 14) + 'px';
    _ghost.style.top  = (y + 14) + 'px';
  }

  function _ghostDestroy() {
    if (_ghost) { _ghost.parentNode && _ghost.parentNode.removeChild(_ghost); _ghost = null; }
  }

  function _highlightZones(x, y) {
    if (!_ghost) return;
    _ghost.style.display = 'none';
    var el = document.elementFromPoint(x, y);
    _ghost.style.display = '';
    document.querySelectorAll('[data-drop-zone]').forEach(function (z) {
      z.classList.toggle('drag-over', !!(el && z.contains(el)));
    });
  }

  return { init: init, fromSeatDragging: fromSeatDragging };
})();
