window.App = window.App || {};

App.C = (function () {
  'use strict';

  var DEFAULT_FLOOR_LAYOUT = [
    { id: '16.38', x: 50,  y: 80,  zone: 'top' },
    { id: '16.39', x: 160, y: 80,  zone: 'top' },
    { id: '16.40', x: 270, y: 80,  zone: 'top' },
    { id: '640',   x: 380, y: 80,  zone: 'top' },
    { id: '16.37', x: 50,  y: 150, zone: 'top' },
    { id: '16.36', x: 160, y: 150, zone: 'top' },
    { id: '16.35', x: 270, y: 150, zone: 'top' },
    { id: '502',   x: 380, y: 150, zone: 'top' },
    { id: '16.32', x: 50,  y: 250, zone: 'mid-top' },
    { id: '16.33', x: 160, y: 250, zone: 'mid-top' },
    { id: '16.34', x: 270, y: 250, zone: 'mid-top' },
    { id: '504',   x: 380, y: 250, zone: 'mid-top' },
    { id: '638',   x: 50,  y: 320, zone: 'mid-top' },
    { id: '16.30', x: 160, y: 320, zone: 'mid-top' },
    { id: '16.29', x: 270, y: 320, zone: 'mid-top' },
    { id: '505',   x: 380, y: 320, zone: 'mid-top' },
    { id: '16.27', x: 50,  y: 420, zone: 'mid' },
    { id: '16.26', x: 160, y: 420, zone: 'mid' },
    { id: '16.25', x: 270, y: 420, zone: 'mid' },
    { id: '16.24', x: 270, y: 490, zone: 'mid' },
    { id: '16.23', x: 560, y: 430, zone: 'right' },
    { id: '636',   x: 560, y: 520, zone: 'right' },
    { id: '777',   x: 560, y: 610, zone: 'right' },
    { id: '16.18', x: 50,  y: 640, zone: 'bottom' },
    { id: '16.19', x: 160, y: 640, zone: 'bottom' },
    { id: '16.20', x: 270, y: 640, zone: 'bottom' },
    { id: '16.21', x: 380, y: 640, zone: 'bottom' },
    { id: '16.17', x: 50,  y: 710, zone: 'bottom' },
    { id: '16.16', x: 160, y: 710, zone: 'bottom' },
    { id: '16.15', x: 270, y: 710, zone: 'bottom' },
    { id: '634',   x: 380, y: 710, zone: 'bottom' },
    { id: '16.10', x: 50,  y: 780, zone: 'bottom' },
    { id: '632',   x: 160, y: 780, zone: 'bottom' },
    { id: '630',   x: 270, y: 780, zone: 'bottom' },
    { id: '628',   x: 380, y: 780, zone: 'bottom' }
  ];

  var STATUS_COLOR = {
    OFFICE:   '#82D6CC',
    REMOTE:   '#3B3760',
    VACATION: '#BD9375',
    DAY_OFF:  '#1F1D34',
    FREE:     '#1C1A35'
  };

  var DESK_W = 90;
  var DESK_H = 50;
  var SVG_W  = 680;
  var SVG_H  = 870;

  return {
    DEFAULT_FLOOR_LAYOUT: DEFAULT_FLOOR_LAYOUT,
    STATUS_COLOR: STATUS_COLOR,
    DESK_W: DESK_W, DESK_H: DESK_H, SVG_W: SVG_W, SVG_H: SVG_H
  };
})();
