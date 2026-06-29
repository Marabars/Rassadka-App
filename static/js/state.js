window.App = window.App || {};

App.state = (function () {
  'use strict';

  var now = new Date();
  var _state = {
    tab: 'upload',
    month: now.getFullYear() + '-' + String(now.getMonth() + 1).padStart(2, '0'),
    day: null,
    assignments: [],
    employees: [],
    layout: null,
    selectedEmployee: null,
    selectedSeat: null,
    uploadStatus: null,
    lastGenResult: null,
  };

  function get() { return _state; }
  function set(patch) { Object.assign(_state, patch); }
  function getMonth() { return _state.month; }
  function getDay() { return _state.day; }

  function assignmentsBySeat() {
    var map = {};
    _state.assignments.forEach(function (a) {
      if (a.seat_id) map[a.seat_id] = a;
    });
    return map;
  }

  return { get: get, set: set, getMonth: getMonth, getDay: getDay,
           assignmentsBySeat: assignmentsBySeat };
})();
