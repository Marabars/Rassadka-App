window.App = window.App || {};

App.dragDrop = (function () {
  'use strict';

  var _draggedEmployee = null;
  var _draggedFromSeat = null;

  function init() {
    _wireEmployeeChips();
    _wireSeatTargets();
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
      });
      chip.addEventListener('dragend', function () {
        chip.classList.remove('dragging');
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

  return { init: init };
})();
