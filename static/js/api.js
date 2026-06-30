window.App = window.App || {};

App.api = (function () {
  'use strict';

  var BASE = '';

  async function uploadFile(endpoint, file) {
    var fd = new FormData();
    fd.append('file', file);
    var r = await fetch(BASE + endpoint, { method: 'POST', body: fd });
    if (!r.ok) throw new Error(await r.text());
    return r.json();
  }

  async function generate(month, choicesId, templateId) {
    var r = await fetch(BASE + '/api/generate', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ month: month, choices_id: choicesId, template_id: templateId })
    });
    if (!r.ok) throw new Error(await r.text());
    return r.json();
  }

  async function getSeating(month, date) {
    var params = 'month=' + encodeURIComponent(month);
    if (date) params += '&date=' + encodeURIComponent(date);
    var r = await fetch(BASE + '/api/seating?' + params);
    if (!r.ok) throw new Error(await r.text());
    return r.json();
  }

  async function overrideSeat(date, seatId, employeeName) {
    var r = await fetch(BASE + '/api/seating/' + encodeURIComponent(date) + '/' + encodeURIComponent(seatId), {
      method: 'PUT',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ employee_name: employeeName || null })
    });
    if (!r.ok) throw new Error(await r.text());
    return r.json();
  }

  async function getEmployees(date) {
    var r = await fetch(BASE + '/api/employees?date=' + encodeURIComponent(date));
    if (!r.ok) throw new Error(await r.text());
    return r.json();
  }

  async function getLayout() {
    var r = await fetch(BASE + '/api/layout');
    if (!r.ok) throw new Error(await r.text());
    return r.json();
  }

  async function saveLayout(layout) {
    var r = await fetch(BASE + '/api/layout', {
      method: 'PUT',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ layout: layout })
    });
    if (!r.ok) throw new Error(await r.text());
    return r.json();
  }

  async function getPreferences() {
    var r = await fetch(BASE + '/api/preferences');
    if (!r.ok) throw new Error(await r.text());
    return r.json();
  }

  async function setPreference(employeeName, seats) {
    var r = await fetch(BASE + '/api/preferences/' + encodeURIComponent(employeeName), {
      method: 'PUT',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ seats: seats })
    });
    if (!r.ok) throw new Error(await r.text());
    return r.json();
  }

  async function getLogs(lines) {
    var n = lines || 300;
    var r = await fetch(BASE + '/api/logs?lines=' + n);
    if (!r.ok) throw new Error(await r.text());
    return r.json();
  }

  async function postClientError(message, stack, url) {
    try {
      await fetch(BASE + '/api/log-error', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ message: message, stack: stack || '', url: url || location.href })
      });
    } catch (_) {}
  }

  return {
    uploadFile: uploadFile, generate: generate,
    getSeating: getSeating, overrideSeat: overrideSeat,
    getEmployees: getEmployees, getLayout: getLayout, saveLayout: saveLayout,
    getPreferences: getPreferences, setPreference: setPreference,
    getLogs: getLogs, postClientError: postClientError
  };
})();
