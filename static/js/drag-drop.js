window.App = window.App || {};

App.dragDrop = (function () {
  'use strict';

  var _draggedEmployee = null;
  var _draggedFromSeat = null;

  function init() {
    _wireEmployeeChips();
    _wireSeatTargets();
    _wireDropZones();
  }

  function _wireEmployeeChips() {
    var chips = document.querySelectorAll('[data-employee]');
    chips.forEach(function (chip) {
      chip.draggable = true;
      chip.addEventListener('dragstart', function (e) {
        _draggedEmployee = chip.dataset.employee;
        _draggedFromSeat = chip.dataset.fromSeat || null;
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
    var desks = document.querySelectorAll('[data-seat]');
    desks.forEach(function (desk) {
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
        if (!date) {
          alert('Выберите день на календаре перед перемещением.');
          return;
        }

        App.api.overrideSeat(date, seatId, employee).then(function () {
          return App.api.getSeating(App.state.getMonth(), date);
        }).then(function (data) {
          App.state.set({ assignments: data.assignments });
          App.floorPlan.refresh();
          App.seatingList.refresh();
        }).catch(function (err) {
          console.error('Drop failed:', err);
        });

        _draggedEmployee = null;
        _draggedFromSeat = null;
      });
    });
  }

  function _wireDropZones() {
    var zones = document.querySelectorAll('[data-drop-zone]');
    zones.forEach(function (zone) {
      zone.addEventListener('dragover', function (e) {
        if (!_draggedEmployee) return;
        e.preventDefault();
        e.dataTransfer.dropEffect = 'move';
        zone.classList.add('drag-over');
      });
      zone.addEventListener('dragleave', function (e) {
        if (!zone.contains(e.relatedTarget)) {
          zone.classList.remove('drag-over');
        }
      });
      zone.addEventListener('drop', function (e) {
        e.preventDefault();
        zone.classList.remove('drag-over');
        var employee = _draggedEmployee;
        var fromSeat = _draggedFromSeat;
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
        }).catch(function (err) {
          console.error('Drop zone failed:', err);
        });

        _draggedEmployee = null;
        _draggedFromSeat = null;
      });
    });
  }

  return { init: init };
})();