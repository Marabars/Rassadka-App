window.App = window.App || {};

App.U = (function () {
  'use strict';

  function el(tag, attrs, children) {
    var node = document.createElement(tag);
    attrs = attrs || {};
    Object.keys(attrs).forEach(function (k) {
      if (k === 'text') { node.textContent = attrs[k]; }
      else if (k === 'html') { node.innerHTML = attrs[k]; }
      else { node.setAttribute(k, attrs[k]); }
    });
    (children || []).forEach(function (c) { if (c) node.appendChild(c); });
    return node;
  }

  function debounce(fn, ms) {
    var t;
    return function () {
      clearTimeout(t);
      var args = arguments;
      var ctx = this;
      t = setTimeout(function () { fn.apply(ctx, args); }, ms);
    };
  }

  function qs(selector, root) { return (root || document).querySelector(selector); }
  function qsa(selector, root) { return Array.from((root || document).querySelectorAll(selector)); }

  function svgEl(tag, attrs) {
    var NS = 'http://www.w3.org/2000/svg';
    var node = document.createElementNS(NS, tag);
    attrs = attrs || {};
    Object.keys(attrs).forEach(function (k) { node.setAttribute(k, attrs[k]); });
    return node;
  }

  return { el: el, debounce: debounce, qs: qs, qsa: qsa, svgEl: svgEl };
})();
