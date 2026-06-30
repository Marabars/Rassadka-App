App.logs = (function () {
    var _container = null;
    var _autoTimer = null;
    var _state = { level: 'ALL', search: '', autoRefresh: false };

    var LEVEL_COLORS = {
        'ERROR':    '#FF6B6B',
        'CRITICAL': '#FF4444',
        'WARNING':  '#FFD166',
        'WARN':     '#FFD166',
        'INFO':     '#82D6CC',
        'DEBUG':    '#9B8EC4',
        '[CLIENT]': '#BD9375',
    };

    function _levelOf(line) {
        var parts = ['CRITICAL','ERROR','WARNING','WARN','INFO','DEBUG','[CLIENT]'];
        for (var i = 0; i < parts.length; i++) {
            if (line.indexOf(parts[i]) !== -1) return parts[i];
        }
        return 'DEBUG';
    }

    function _colorLine(line) {
        var lvl = _levelOf(line);
        var color = LEVEL_COLORS[lvl] || '#C8C4DE';
        var esc = line.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
        return '<span style="color:' + color + '">' + esc + '</span>';
    }

    function _render(lines) {
        if (!_container) return;
        var search = _state.search.toLowerCase();
        var level  = _state.level;

        var filtered = lines.filter(function (l) {
            if (!l.trim()) return false;
            var lvl = _levelOf(l);
            if (level !== 'ALL') {
                if (level === 'ERROR'   && lvl !== 'ERROR'   && lvl !== 'CRITICAL') return false;
                if (level === 'WARNING' && lvl !== 'WARNING' && lvl !== 'WARN' &&
                                          lvl !== 'ERROR'   && lvl !== 'CRITICAL') return false;
                if (level === 'INFO'    && lvl === 'DEBUG') return false;
            }
            if (search && l.toLowerCase().indexOf(search) === -1) return false;
            return true;
        });

        var pre = _container.querySelector('.logs-pre');
        if (!pre) return;
        if (!filtered.length) {
            pre.innerHTML = '<span style="color:#5A5580">Нет записей</span>';
            return;
        }
        pre.innerHTML = filtered.map(_colorLine).join('\n');
        pre.scrollTop = pre.scrollHeight;
    }

    function _load() {
        App.api.getLogs(500).then(function (data) {
            _render(data.lines || []);
        }).catch(function (err) {
            var pre = _container && _container.querySelector('.logs-pre');
            if (pre) pre.innerHTML = '<span style="color:#FF6B6B">Ошибка загрузки: ' + err.message + '</span>';
        });
    }

    function _startAuto() {
        if (_autoTimer) return;
        _autoTimer = setInterval(_load, 5000);
    }

    function _stopAuto() {
        if (_autoTimer) { clearInterval(_autoTimer); _autoTimer = null; }
    }

    function init(container) {
        _container = container;
        _container.innerHTML = [
            '<div class="logs-toolbar">',
            '  <select class="logs-level-sel" title="Уровень">',
            '    <option value="ALL">Все уровни</option>',
            '    <option value="INFO">INFO+</option>',
            '    <option value="WARNING">WARNING+</option>',
            '    <option value="ERROR">ERROR</option>',
            '  </select>',
            '  <input class="logs-search" type="text" placeholder="Поиск..." />',
            '  <label class="logs-auto-lbl">',
            '    <input type="checkbox" class="logs-auto-chk" /> Авто (5с)',
            '  </label>',
            '  <button class="logs-refresh-btn btn-secondary">Обновить</button>',
            '</div>',
            '<pre class="logs-pre">Загрузка…</pre>',
        ].join('');

        _container.querySelector('.logs-level-sel').addEventListener('change', function (e) {
            _state.level = e.target.value;
            _load();
        });
        _container.querySelector('.logs-search').addEventListener('input', function (e) {
            _state.search = e.target.value;
            _load();
        });
        _container.querySelector('.logs-refresh-btn').addEventListener('click', _load);
        _container.querySelector('.logs-auto-chk').addEventListener('change', function (e) {
            _state.autoRefresh = e.target.checked;
            if (_state.autoRefresh) _startAuto(); else _stopAuto();
        });

        _load();
    }

    function destroy() {
        _stopAuto();
        _container = null;
    }

    return { init: init, destroy: destroy, refresh: _load };
})();
