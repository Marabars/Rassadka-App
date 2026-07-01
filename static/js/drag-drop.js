window.App = window.App || {};

App.dragDrop = (function () {
  'use strict';

  var _draggedEmployee = null;
  var _ghost = null;
  var _fromSeatDragging = false;
  var _activeDrag = null;   // tracks the one live mouse-drag from SVG

  /* ---- public ---- */
  function fromSeatDragging() { return _fromSeatDragging; }

  /* Debounce init so parallel setTimeout(init,0) calls collapse into one */
  var _initTimer = null;
  function init() {
    if (_initTimer) clearTimeout(_initTimer);
    _initTimer = setTimeout(_doInit, 0);
  }

  function _doInit() {
    _initTimer = null;
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
        _showDropZones();
      });
      chip.addEventListener('dragend', function () {
        chip.classList.remove('dragging');
        _hideDropZones();
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
        if (!date) { alert('Выберите день на календаре.'); return; }
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
        _dropToStatus(date, employee, targetStatus);
        _draggedEmployee = null;
      });
    });
  }

  /* ------------------------------------------------------------------ */
  /* Custom mouse drag: SVG seat → drop zone / another seat              */
  /* ------------------------------------------------------------------ */
  function _wireSeatMouseDrag() {
    document.querySelectorAll('.desk-group[data-seat]').forEach(function (desk) {
      /* Prevent duplicate listeners after multiple init() calls */
      if (desk._seatDragBound) return;
      desk._seatDragBound = true;

      desk.addEventListener('mousedown', function (e) {
        /* Only left button */
        if (e.button !== 0) return;

        var sid = desk.dataset.seat;
        var byId = App.state.assignmentsBySeat();
        var asgn = byId[sid];
        if (!asgn || !asgn.employee_name || asgn.status !== 'OFFICE') return;

        /* Cancel any leftover drag */
        _cancelActiveDrag();

        var startX = e.clientX, startY = e.clientY;
        var empName = asgn.employee_name;
        var dragging = false;

        function onMove(e) {
          var dx = e.clientX - startX, dy = e.clientY - startY;
          if (!dragging && Math.sqrt(dx * dx + dy * dy) > 5) {
            dragging = true;
            _ghostCreate(empName, e.clientX, e.clientY);
            _showDropZones();
          }
          if (dragging) {
            _ghostMove(e.clientX, e.clientY);
            _highlightZones(e.clientX, e.clientY);
          }
        }

        function onUp(e) {
          cleanup();
          if (!dragging) return;

          _fromSeatDragging = true;
          var el = _elementUnderGhost(e.clientX, e.clientY);
          _ghostDestroy();
          _hideDropZones();

          var date = App.state.getDay();
          if (!date) { _fromSeatDragging = false; return; }

          var zone   = el && el.closest('[data-drop-zone]');
          var seatEl = el && el.closest('[data-seat]');

          var p;
          if (zone && zone.dataset.dropZone === 'other') {
            p = _dropToStatus(date, empName, 'REMOTE');
          } else if (seatEl && seatEl !== desk) {
            p = _dropToSeat(date, seatEl.dataset.seat, empName);
          }
          if (p) p.catch(function (err) { console.error('Seat drag failed:', err); });

          setTimeout(function () { _fromSeatDragging = false; }, 0);
        }

        function cleanup() {
          document.removeEventListener('mousemove', onMove);
          document.removeEventListener('mouseup',   onUp);
          document.removeEventListener('mouseleave', onLeave);
          _activeDrag = null;
        }

        function onLeave(e) {
          /* Mouse left the browser window — cancel drag */
          if (e.target === document.documentElement) {
            cleanup();
            _ghostDestroy();
            _hideDropZones();
          }
        }

        document.addEventListener('mousemove',  onMove);
        document.addEventListener('mouseup',    onUp);
        document.addEventListener('mouseleave', onLeave);
        _activeDrag = { cancel: cleanup };
      });
    });
  }

  /* ------------------------------------------------------------------ */
  /* Shared drop actions                                                  */
  /* ------------------------------------------------------------------ */
  function _dropToStatus(date, empName, status) {
    return App.api.setEmployeeStatus(date, empName, status).then(function () {
      return App.api.getSeating(App.state.getMonth(), date);
    }).then(function (data) {
      App.state.set({ assignments: data.assignments });
      return App.api.getEmployees(date);
    }).then(function (data) {
      App.state.set({ employees: data.employees });
      App.floorPlan.refresh();
      App.seatingList.refresh();
    });
  }

  function _dropToSeat(date, seatId, empName) {
    return App.api.overrideSeat(date, seatId, empName).then(function () {
      return App.api.getSeating(App.state.getMonth(), date);
    }).then(function (data) {
      App.state.set({ assignments: data.assignments });
      App.floorPlan.refresh();
      App.seatingList.refresh();
    });
  }

  /* ------------------------------------------------------------------ */
  /* Ghost helpers                                                        */
  /* ------------------------------------------------------------------ */
  function _ghostCreate(name, x, y) {
    _ghostDestroy();   // always clean up before creating
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
    if (_ghost) {
      try { _ghost.parentNode && _ghost.parentNode.removeChild(_ghost); } catch (_) {}
      _ghost = null;
    }
  }

  /* Returns the element under cursor, ignoring the ghost (pointer-events:none already) */
  function _elementUnderGhost(x, y) {
    return document.elementFromPoint(x, y);
  }

  /* ------------------------------------------------------------------ */
  /* Drop zone visual helpers                                             */
  /* ------------------------------------------------------------------ */
  function _showDropZones() {
    document.querySelectorAll('[data-drop-zone]').forEach(function (z) {
      z.classList.add('drop-zone-visible');
    });
  }

  function _hideDropZones() {
    document.querySelectorAll('[data-drop-zone]').forEach(function (z) {
      z.classList.remove('drop-zone-visible', 'drag-over');
    });
  }

  function _highlightZones(x, y) {
    var el = document.elementFromPoint(x, y);
    document.querySelectorAll('[data-drop-zone]').forEach(function (z) {
      z.classList.toggle('drag-over', !!(el && z.contains(el)));
    });
  }

  function _cancelActiveDrag() {
    if (_activeDrag) {
      _activeDrag.cancel();
      _activeDrag = null;
    }
    _ghostDestroy();
    _hideDropZones();
  }

  return { init: init, fromSeatDragging: fromSeatDragging };
})();