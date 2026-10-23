'use strict';

(function () {
  // ------------------------------------------------------------------
  // Configuration
  // ------------------------------------------------------------------
  var CONFIG = {
    worksheetName: 'AttendanceData', // name of the worksheet that feeds this extension
    pageSize: 10000,                 // rows per page when reading worksheet data
    // Campuses left out by default (non-physical campuses). Matching ignores case, spaces, and punctuation.
    excludedCampuses: ['Beechmont Center', 'Multination Fellowship', 'SE Online', 'SE Recovery']
  };

  var DAY = 86400000;
  var SHIFT = 364 * DAY; // prior period sits exactly 52 weeks back so Sundays line up
  var MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

  var AREA_DEFS = {
    Adult: { label: 'Adult', types: ['Adult'] },
    Kids: { label: 'Kids', types: ['EarlyChildhood', 'Elementary'] },
    EarlyChildhood: { label: 'Early Childhood', types: ['EarlyChildhood'] },
    Elementary: { label: 'Elementary', types: ['Elementary'] },
    MSM: { label: 'Middle School (MSM)', types: ['MSM'] },
    All: { label: 'All areas', types: ['Adult', 'EarlyChildhood', 'Elementary', 'MSM'] }
  };
  var TABS = [['All', 'All areas'], ['Adult', 'Adult'], ['Kids', 'Kids'], ['EarlyChildhood', 'Early Childhood'], ['Elementary', 'Elementary'], ['MSM', 'MSM']];
  var PRESETS = [['4w', 'Last 4 weeks'], ['6m', 'Last 6 months'], ['ytd', 'Year to date'], ['12m', 'Last 12 months']];

  var state = {
    loading: true,
    error: null,
    rows: [],
    campuses: [],
    selected: new Set(),
    minWeek: null,
    maxWeek: null,
    check: null,
    preset: '4w',
    start: null,
    end: null,
    kidsOpen: false,
    campusOpen: false,
    tab: 'All',
    message: null
  };

  var app = document.getElementById('app');

  // ------------------------------------------------------------------
  // Small helpers
  // ------------------------------------------------------------------
  function esc(s) {
    return String(s).replace(/[&<>"']/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
    });
  }
  function pad(n) { return (n < 10 ? '0' : '') + n; }
  function normName(s) { return String(s).toLowerCase().replace(/[^a-z0-9]/g, ''); }
  function defaultCampusSet() {
    var ex = new Set(CONFIG.excludedCampuses.map(normName));
    var set = new Set(state.campuses.filter(function (c) { return !ex.has(normName(c)); }));
    return set.size ? set : new Set(state.campuses);
  }
  function sameSet(a, b) {
    if (a.size !== b.size) return false;
    var same = true;
    a.forEach(function (v) { if (!b.has(v)) same = false; });
    return same;
  }
  function missingExclusions() {
    var present = new Set(state.campuses.map(normName));
    return CONFIG.excludedCampuses.filter(function (n) { return !present.has(normName(n)); });
  }
  function iso(ms) {
    var d = new Date(ms);
    return d.getUTCFullYear() + '-' + pad(d.getUTCMonth() + 1) + '-' + pad(d.getUTCDate());
  }
  function parseISO(s) {
    var m = /^(\d{4})-(\d{2})-(\d{2})/.exec(String(s));
    return m ? Date.UTC(+m[1], +m[2] - 1, +m[3]) : NaN;
  }
  function fmtDate(ms, withYear) {
    var d = new Date(ms);
    return MONTHS[d.getUTCMonth()] + ' ' + d.getUTCDate() + (withYear ? ', ' + d.getUTCFullYear() : '');
  }
  function fmtRange(a, b) {
    var sameYear = new Date(a).getUTCFullYear() === new Date(b).getUTCFullYear();
    return fmtDate(a, !sameYear) + ' to ' + fmtDate(b, true);
  }
  // Sunday on or after the given date. Weeks run Monday to Sunday and are dated by their Sunday.
  function weekEnding(ms) {
    var dow = new Date(ms).getUTCDay();
    return ms + ((7 - dow) % 7) * DAY;
  }
  // Latest Sunday that is safe to include: last completed Sunday, capped at the newest week in the data.
  function latestEnd() {
    var t = new Date();
    var today = Date.UTC(t.getFullYear(), t.getMonth(), t.getDate());
    var dow = new Date(today).getUTCDay();
    var end = today - (dow === 0 ? 7 : dow) * DAY;
    if (state.maxWeek !== null && end > state.maxWeek) end = state.maxWeek;
    return end;
  }

  var nf = new Intl.NumberFormat('en-US', { maximumFractionDigits: 0 });
  function fmtInt(n) { return n === null || n === undefined ? 'n/a' : nf.format(Math.round(n)); }
  function fmtSigned(n) {
    if (n === null || n === undefined) return 'n/a';
    return (n > 0 ? '+' : n < 0 ? '-' : '') + nf.format(Math.abs(n));
  }
  function pctPts(p) { return Math.round(p * 1000) / 10; }
  function fmtPct(p) { return p === null ? 'n/a' : Math.abs(pctPts(p)).toFixed(1) + '%'; }
  function dir(p) {
    if (p === null) return 'na';
    var v = pctPts(p);
    return v > 0 ? 'up' : v < 0 ? 'down' : 'flat';
  }
  var ARROW = { up: '\u25B2 ', down: '\u25BC ', flat: '', na: '' };

  // ------------------------------------------------------------------
  // Data loading
  // ------------------------------------------------------------------
  function resolveColumns(names) {
    var lower = names.map(function (n) { return String(n).toLowerCase().replace(/[ _]/g, ''); });
    function find(test) {
      for (var i = 0; i < lower.length; i++) { if (test(lower[i])) return i; }
      return -1;
    }
    var idx = {
      key: find(function (n) { return n.indexOf('sundaykey') !== -1; }),
      campus: find(function (n) { return n.indexOf('campus') !== -1; }),
      type: find(function (n) { return n.indexOf('attendancetype') !== -1; }),
      val: find(function (n) { return n.indexOf('attendance') !== -1 && n.indexOf('type') === -1; })
    };
    var missing = [];
    if (idx.key < 0) missing.push('SundayKey');
    if (idx.campus < 0) missing.push('CampusName');
    if (idx.type < 0) missing.push('AttendanceType');
    if (idx.val < 0) missing.push('SUM(Attendance)');
    if (missing.length) {
      throw new Error('The worksheet "' + CONFIG.worksheetName + '" is missing: ' + missing.join(', ') +
        '.\nColumns the extension received: ' + names.join(', ') +
        '.\nFix the pills on that worksheet and click Refresh data.');
    }
    return idx;
  }

  async function loadData() {
    var dashboard = tableau.extensions.dashboardContent.dashboard;
    var ws = dashboard.worksheets.find(function (w) { return w.name === CONFIG.worksheetName; });
    if (!ws) {
      throw new Error('Worksheet "' + CONFIG.worksheetName + '" was not found on this dashboard.\nWorksheets on this dashboard: ' +
        (dashboard.worksheets.map(function (w) { return w.name; }).join(', ') || '(none)') +
        '.\nAdd that worksheet to the dashboard (it can be tiny and floating) and click Refresh data.');
    }

    var reader = await ws.getSummaryDataReaderAsync(CONFIG.pageSize, { ignoreSelection: true });
    var rows = [];
    var idx = null;
    var badDates = 0;
    var nonSunday = 0;
    var total = 0;
    try {
      for (var p = 0; p < reader.pageCount; p++) {
        var page = await reader.getPageAsync(p);
        if (!idx) idx = resolveColumns(page.columns.map(function (c) { return c.fieldName; }));
        for (var i = 0; i < page.data.length; i++) {
          var r = page.data[i];
          total++;
          var week = parseISO(r[idx.key].value);
          if (isNaN(week)) { badDates++; continue; }
          if (new Date(week).getUTCDay() !== 0) nonSunday++;
          var raw = r[idx.val].nativeValue !== undefined && r[idx.val].nativeValue !== null ? r[idx.val].nativeValue : r[idx.val].value;
          var v = Number(raw);
          if (!isFinite(v)) v = 0;
          rows.push({
            week: week,
            campus: String(r[idx.campus].value),
            type: String(r[idx.type].value),
            value: v
          });
        }
      }
    } finally {
      await reader.releaseAsync();
    }

    if (!rows.length) {
      throw new Error('The worksheet "' + CONFIG.worksheetName + '" returned no usable rows (rows read: ' + total + ', unreadable dates: ' + badDates + ').');
    }

    var campusSet = new Set();
    var minW = Infinity;
    var maxW = -Infinity;
    rows.forEach(function (r) {
      campusSet.add(r.campus);
      if (r.week < minW) minW = r.week;
      if (r.week > maxW) maxW = r.week;
    });
    var campuses = Array.from(campusSet).sort(function (a, b) { return a.localeCompare(b); });

    var previous = state.selected;
    state.rows = rows;
    state.campuses = campuses;
    var kept = new Set(campuses.filter(function (c) { return previous.has(c); }));
    state.selected = kept.size ? kept : defaultCampusSet();
    state.minWeek = minW;
    state.maxWeek = maxW;
    state.check = { rows: rows.length, badDates: badDates, nonSunday: nonSunday };
  }

  // ------------------------------------------------------------------
  // Period selection
  // ------------------------------------------------------------------
  function applyPreset(key) {
    var end = latestEnd();
    var start;
    if (key === '4w') start = end - 3 * 7 * DAY;
    else if (key === '12m') start = end - 51 * 7 * DAY;
    else if (key === '6m') {
      var d = new Date(end);
      d.setUTCMonth(d.getUTCMonth() - 6);
      start = weekEnding(d.getTime());
    } else {
      var year = new Date(end).getUTCFullYear();
      start = weekEnding(Date.UTC(year, 0, 1));
      if (start > end) start = weekEnding(Date.UTC(year - 1, 0, 1));
    }
    state.preset = key;
    state.start = start;
    state.end = end;
    state.message = null;
  }

  // ------------------------------------------------------------------
  // Calculations
  // ------------------------------------------------------------------
  function weeklyMap() {
    var m = new Map();
    state.rows.forEach(function (r) {
      if (!state.selected.has(r.campus)) return;
      var t = m.get(r.week);
      if (!t) { t = {}; m.set(r.week, t); }
      t[r.type] = (t[r.type] || 0) + r.value;
    });
    return m;
  }
  function areaValue(map, week, key) {
    var t = map.get(week);
    if (!t) return null;
    var sum = 0;
    var any = false;
    AREA_DEFS[key].types.forEach(function (ty) {
      if (t[ty] !== undefined) { sum += t[ty]; any = true; }
    });
    return any ? sum : null;
  }

  function buildModel() {
    var n = Math.floor((state.end - state.start) / (7 * DAY)) + 1;
    var weeks = [];
    for (var i = 0; i < n; i++) weeks.push(state.start + i * 7 * DAY);
    var map = weeklyMap();

    function agg(arr) {
      var total = 0;
      var count = 0;
      arr.forEach(function (v) { if (v !== null && v > 0) { total += v; count++; } });
      return { total: total, count: count, avg: count ? total / count : null };
    }
    function stats(key) {
      var cur = weeks.map(function (w) { return areaValue(map, w, key); });
      var pri = weeks.map(function (w) { return areaValue(map, w - SHIFT, key); });
      var c = agg(cur);
      var p = agg(pri);
      var pct = (c.avg !== null && p.avg !== null && p.avg > 0) ? (c.avg - p.avg) / p.avg : null;
      var change = (c.avg !== null && p.avg !== null) ? Math.round(c.avg) - Math.round(p.avg) : null;
      return { key: key, label: AREA_DEFS[key].label, cur: cur, pri: pri, c: c, p: p, pct: pct, change: change };
    }

    var s = {};
    ['Adult', 'Kids', 'EarlyChildhood', 'Elementary', 'MSM', 'All'].forEach(function (k) { s[k] = stats(k); });
    return { n: n, weeks: weeks, s: s };
  }

  // ------------------------------------------------------------------
  // Rendering
  // ------------------------------------------------------------------
  function barHtml(pct, sub) {
    var w = 0;
    if (pct !== null) w = Math.min(100, Math.round(Math.abs(pct) * 100 * 10));
    var neg = pct !== null && pctPts(pct) < 0 && w > 0 ? '<span class="fill down" style="width:' + w + 'px"></span>' : '';
    var pos = pct !== null && pctPts(pct) > 0 && w > 0 ? '<span class="fill up" style="width:' + w + 'px"></span>' : '';
    return '<div class="bar" aria-hidden="true"><div class="neg">' + neg + '</div><div class="pos">' + pos + '</div></div>';
  }

  function rowHtml(s, o) {
    var d = dir(s.pct);
    var cls = o.sub ? 'row sub' : (o.total ? 'row total' : 'row');
    var nameCell;
    if (o.toggle) {
      nameCell = '<div class="area-cell"><div class="area-name">' + esc(s.label) + '</div>' +
        '<button class="btn small" data-action="kids-toggle" data-fid="kids-toggle" aria-expanded="' + state.kidsOpen + '">' +
        (state.kidsOpen ? 'Hide detail' : 'Show Early Childhood and Elementary') + '</button></div>';
    } else {
      nameCell = '<div class="area-name">' + esc(s.label) + '</div>';
    }
    return '<div class="' + cls + '">' + nameCell +
      '<div class="num avg">' + fmtInt(s.c.avg) + '</div>' +
      '<div class="num avg prior">' + fmtInt(s.p.avg) + '</div>' +
      '<div class="num chg ' + d + '">' + fmtSigned(s.change) + '</div>' +
      '<div class="num pct ' + d + '">' + ARROW[d] + fmtPct(s.pct) + '</div>' +
      barHtml(s.pct, o.sub) +
      '<div class="num tot">' + fmtInt(s.c.total) + ' vs ' + fmtInt(s.p.total) + '</div>' +
      '</div>';
  }

  function kpiHtml(model) {
    var a = model.s.All;
    var d = dir(a.pct);
    var up = 0;
    var down = 0;
    ['Adult', 'Kids', 'MSM'].forEach(function (k) {
      var dd = dir(model.s[k].pct);
      if (dd === 'up') up++;
      else if (dd === 'down') down++;
    });
    var areaText = up + (up === 1 ? ' area growing, ' : ' areas growing, ') + down + ' declining';
    var changeText = a.change === null ? '' : (a.change > 0 ? 'more people per week than prior' : a.change < 0 ? 'fewer people per week than prior' : 'no change per week');
    var big = a.pct === null ? 'n/a' : ARROW[d] + (d === 'up' ? '+' : d === 'down' ? '-' : '') + fmtPct(a.pct);
    return '<div class="kpis">' +
      '<div class="card kpi"><div class="k-label">Avg weekly attendance, current</div><div class="k-big">' + fmtInt(a.c.avg) + '</div><div class="k-small">' + fmtInt(a.c.total) + ' total over ' + a.c.count + ' weeks</div></div>' +
      '<div class="card kpi"><div class="k-label">Avg weekly attendance, prior</div><div class="k-big prior">' + fmtInt(a.p.avg) + '</div><div class="k-small">' + fmtInt(a.p.total) + ' total over ' + a.p.count + ' weeks</div></div>' +
      '<div class="card kpi"><div class="k-label">Change in weekly average</div><div class="k-big ' + d + '">' + fmtSigned(a.change) + '</div><div class="k-small">' + changeText + '</div></div>' +
      '<div class="card kpi hero ' + d + '"><div class="k-label">Growth vs prior period</div><div class="k-big">' + big + '</div><div class="k-small">' + areaText + '</div></div>' +
      '</div>';
  }

  function niceTicks(lo, hi, count) {
    var raw = (hi - lo) / count;
    var mag = Math.pow(10, Math.floor(Math.log10(raw)));
    var norm = raw / mag;
    var step = (norm <= 1 ? 1 : norm <= 2 ? 2 : norm <= 5 ? 5 : 10) * mag;
    var first = Math.floor(lo / step) * step;
    var last = Math.ceil(hi / step) * step;
    var ticks = [];
    for (var v = first; v <= last + step * 0.001; v += step) ticks.push(v);
    return ticks;
  }

  function chartHtml(model) {
    var s = model.s[state.tab];
    var n = model.n;
    var vals = [];
    s.cur.concat(s.pri).forEach(function (v) { if (v !== null && v > 0) vals.push(v); });
    if (!vals.length) return '<div class="empty">No data for the selected campuses and period.</div>';

    var lo = Math.min.apply(null, vals);
    var hi = Math.max.apply(null, vals);
    var span = hi - lo;
    var lo2 = span === 0 ? Math.max(0, lo * 0.9) : Math.max(0, lo - span * 0.1);
    var hi2 = span === 0 ? hi * 1.1 + 1 : hi + span * 0.1;
    var ticks = niceTicks(lo2, hi2, 4);
    var yMin = ticks[0];
    var yMax = ticks[ticks.length - 1];
    function y(v) { return 220 - ((v - yMin) / (yMax - yMin)) * 204; }
    var L = 72;
    function x(i) { return n > 1 ? L + i * ((740 - L) / (n - 1)) : (L + 740) / 2; }
    function pathFor(arr) {
      var d = '';
      var pen = false;
      arr.forEach(function (v, i) {
        if (v === null || v <= 0) { pen = false; return; }
        d += (pen ? 'L' : 'M') + x(i).toFixed(1) + ' ' + y(v).toFixed(1) + ' ';
        pen = true;
      });
      return d.trim();
    }

    var out = '<svg viewBox="0 0 860 270" role="img" aria-label="Line chart of weekly attendance for ' + esc(s.label) + ', current period versus the same weeks one year earlier, with flat lines at each period weekly average">';
    ticks.forEach(function (t, i) {
      var yy = y(t).toFixed(1);
      out += '<line x1="' + L + '" y1="' + yy + '" x2="745" y2="' + yy + '" stroke="' + (i === 0 ? '#9CA3AF' : '#E5E7EB') + '" stroke-width="1"/>';
      out += '<text x="' + (L - 8) + '" y="' + (y(t) + 4).toFixed(1) + '" font-size="12" fill="#4B5563" text-anchor="end">' + fmtInt(t) + '</text>';
    });

    var idxSet = {};
    [0, Math.round((n - 1) / 4), Math.round((n - 1) / 2), Math.round(3 * (n - 1) / 4), n - 1].forEach(function (i) { idxSet[i] = true; });
    Object.keys(idxSet).forEach(function (k) {
      var i = +k;
      out += '<text x="' + x(i).toFixed(1) + '" y="242" font-size="12" fill="#4B5563" text-anchor="middle">Wk ' + (i + 1) + '</text>';
      out += '<text x="' + x(i).toFixed(1) + '" y="259" font-size="12" fill="#4B5563" text-anchor="middle">' + fmtDate(model.weeks[i], false) + '</text>';
    });

    out += '<path d="' + pathFor(s.pri) + '" fill="none" stroke="#6B7280" stroke-width="2" stroke-dasharray="7 5" stroke-linejoin="round"/>';
    out += '<path d="' + pathFor(s.cur) + '" fill="none" stroke="#1E4F8F" stroke-width="3" stroke-linejoin="round"/>';

    s.pri.forEach(function (v, i) {
      if (v === null || v <= 0) return;
      out += '<circle cx="' + x(i).toFixed(1) + '" cy="' + y(v).toFixed(1) + '" r="3" fill="#6B7280"><title>Prior, week ' + (i + 1) + ' (' + fmtDate(model.weeks[i] - SHIFT, true) + '): ' + fmtInt(v) + '</title></circle>';
    });
    s.cur.forEach(function (v, i) {
      if (v === null || v <= 0) return;
      out += '<circle cx="' + x(i).toFixed(1) + '" cy="' + y(v).toFixed(1) + '" r="3.5" fill="#1E4F8F"><title>Current, week ' + (i + 1) + ' (' + fmtDate(model.weeks[i], true) + '): ' + fmtInt(v) + '</title></circle>';
    });

    if (s.c.avg !== null && s.p.avg !== null) {
      var cy = y(s.c.avg);
      var py = y(s.p.avg);
      out += '<line x1="' + L + '" y1="' + cy.toFixed(1) + '" x2="745" y2="' + cy.toFixed(1) + '" stroke="#1E4F8F" stroke-width="1.5" stroke-opacity="0.55"/>';
      out += '<line x1="' + L + '" y1="' + py.toFixed(1) + '" x2="745" y2="' + py.toFixed(1) + '" stroke="#6B7280" stroke-width="1.5" stroke-opacity="0.7"/>';
      var curAbove = cy <= py;
      out += '<text x="752" y="' + (curAbove ? cy - 5 : cy + 14).toFixed(1) + '" font-size="12" font-weight="600" fill="#1E4F8F">Avg ' + fmtInt(s.c.avg) + '</text>';
      out += '<text x="752" y="' + (curAbove ? py + 14 : py - 5).toFixed(1) + '" font-size="12" font-weight="600" fill="#4B5563">Avg ' + fmtInt(s.p.avg) + '</text>';
    }
    out += '</svg>';
    return out;
  }

  function filtersHtml() {
    var total = state.campuses.length;
    var count = state.selected.size;
    var isDefault = sameSet(state.selected, defaultCampusSet());
    var summary = (isDefault && count !== total) ? 'Physical campuses (' + count + ')' : (count === total ? 'All campuses (' + total + ')' : (count === 0 ? 'No campuses selected' : count + ' of ' + total + ' campuses'));
    var panel = '';
    if (state.campusOpen) {
      var items = state.campuses.map(function (c, i) {
        var on = state.selected.has(c);
        return '<button class="campus-item" data-action="campus-toggle" data-index="' + i + '" data-fid="campus-' + i + '" aria-pressed="' + on + '">' +
          '<span class="box' + (on ? ' on' : '') + '">' + (on ? '<svg width="12" height="12" viewBox="0 0 12 12" aria-hidden="true"><path d="M2 6.5l2.5 2.5L10 3.5" fill="none" stroke="#FFFFFF" stroke-width="2"/></svg>' : '') + '</span>' +
          '<span>' + esc(c) + '</span></button>';
      }).join('');
      panel = '<div class="campus-panel"><div class="tools">' +
        '<button class="btn link" data-action="campus-all" data-fid="campus-all">Select all</button>' +
        '<button class="btn link" data-action="campus-physical" data-fid="campus-physical">Physical only</button>' +
        '<button class="btn link" data-action="campus-none" data-fid="campus-none">Clear</button></div>' +
        '<div class="campus-list">' + items + '</div></div>';
    }
    var presets = PRESETS.map(function (p) {
      var on = state.preset === p[0];
      return '<button class="btn' + (on ? ' on' : '') + '" data-action="preset" data-value="' + p[0] + '" data-fid="preset-' + p[0] + '" aria-pressed="' + on + '">' + p[1] + '</button>';
    }).join('');

    var valid = state.start <= state.end;
    var periodHtml = '';
    if (valid) {
      var n = Math.floor((state.end - state.start) / (7 * DAY)) + 1;
      periodHtml =
        '<div class="item"><svg width="28" height="6" aria-hidden="true"><line x1="0" y1="3" x2="28" y2="3" stroke="#1E4F8F" stroke-width="4"/></svg>' +
        '<div><strong>Current:</strong> ' + fmtRange(state.start, state.end) + ' <span class="muted">(' + n + ' weeks)</span></div></div>' +
        '<div class="item"><svg width="28" height="6" aria-hidden="true"><line x1="0" y1="3" x2="28" y2="3" stroke="#6B7280" stroke-width="3" stroke-dasharray="6 4"/></svg>' +
        '<div><strong>Prior:</strong> ' + fmtRange(state.start - SHIFT, state.end - SHIFT) + ' <span class="muted">(same ' + n + ' weeks, set automatically 52 weeks back)</span></div></div>';
    }
    return '<div class="card filters"><div class="filter-row">' +
      '<div class="campus-wrap"><div class="label">Campuses</div>' +
      '<button class="campus-btn" data-action="campus-menu" data-fid="campus-menu" aria-expanded="' + state.campusOpen + '"><span>' + esc(summary) + '</span>' +
      '<svg width="14" height="14" viewBox="0 0 14 14" aria-hidden="true"><path d="M2 5l5 5 5-5" fill="none" stroke="#374151" stroke-width="2"/></svg></button>' + panel + '</div>' +
      '<div class="field"><label class="label" for="start-date">Start date</label><input class="input" id="start-date" data-fid="start-date" data-action="start" type="date" value="' + iso(state.start) + '"></div>' +
      '<div class="field"><label class="label" for="end-date">End date</label><input class="input" id="end-date" data-fid="end-date" data-action="end" type="date" value="' + iso(state.end) + '"></div>' +
      '<div class="field grow"><div class="label">Quick periods</div><div class="btn-row nowrap">' + presets + '</div></div>' +
      '</div><div class="period">' + periodHtml +
      '<div class="muted">Dates snap to the Sunday of their week. All holiday weekends included.</div>' +
      (state.message ? '<div class="warn-msg">' + esc(state.message) + '</div>' : '') +
      '</div></div>';
  }

  function render() {
    var activeFid = document.activeElement && document.activeElement.getAttribute ? document.activeElement.getAttribute('data-fid') : null;

    if (state.error) {
      app.innerHTML = '<div class="error"><h2>The dashboard could not load its data</h2>' + esc(state.error) + '</div>';
      return;
    }
    if (state.loading) {
      app.innerHTML = '<div class="status">Loading attendance data...</div>';
      return;
    }

    var html = '<div class="header"><div><h1>Average Weekly Attendance, Year over Year</h1>' +
      '<div class="sub">Total service attendance per week, by ministry area. A person is counted once per service attended. Volunteers excluded.</div></div>' +
      '<div class="meta"><span>Latest complete week: ' + fmtDate(latestEnd(), true) + '</span>' +
      '<button class="btn link" data-action="refresh" data-fid="refresh">Refresh data</button></div></div>';

    html += filtersHtml();

    var valid = state.start <= state.end;
    if (!valid) {
      html += '<div class="card panel"><div class="empty">The start date is after the end date. Adjust the dates or pick a quick period.</div></div>';
    } else if (state.selected.size === 0) {
      html += '<div class="card panel"><div class="empty">Select at least one campus.</div></div>';
    } else {
      var model = buildModel();
      html += kpiHtml(model);

      html += '<div class="card panel"><div class="panel-head"><div><h2>Average weekly attendance by ministry area</h2></div>' +
        '<div class="hint">Bars show % change in weekly average, scaled from -10% to +10%</div></div>' +
        '<div class="table-scroll"><div class="table">' +
        '<div class="thead"><div>Area</div><div class="r">Avg weekly, current</div><div class="r">Avg weekly, prior</div><div class="r">Change</div><div class="r">% change</div><div class="c">Growth / decline</div><div class="r">Total, current vs prior</div></div>' +
        rowHtml(model.s.Adult, {}) +
        rowHtml(model.s.Kids, { toggle: true }) +
        (state.kidsOpen ? rowHtml(model.s.EarlyChildhood, { sub: true }) + rowHtml(model.s.Elementary, { sub: true }) : '') +
        rowHtml(model.s.MSM, {}) +
        rowHtml(model.s.All, { total: true }) +
        '</div></div>';

      var notes = [];
      ['Adult', 'Kids', 'EarlyChildhood', 'Elementary', 'MSM'].forEach(function (k) {
        var s = model.s[k];
        if (s.c.count !== model.n || s.p.count !== model.n) {
          notes.push(s.label + ': attendance in ' + s.c.count + ' of ' + model.n + ' weeks (current), ' + s.p.count + ' of ' + model.n + ' (prior).');
        }
      });
      if (notes.length) {
        html += '<div class="notes"><span class="coverage"><strong>Coverage:</strong> ' + esc(notes.join(' ')) + ' Averages use only weeks with attendance.</span></div>';
      }
      html += '</div>';

      html += '<div class="card panel"><div class="panel-head"><div><h2>Weekly attendance, aligned week by week</h2>' +
        '<div class="hint">Each week of the current period against the same week one year earlier. Flat lines mark each period weekly average.</div></div>' +
        '<div class="btn-row">' + TABS.map(function (t) {
          var on = state.tab === t[0];
          return '<button class="btn' + (on ? ' on' : '') + '" data-action="tab" data-value="' + t[0] + '" data-fid="tab-' + t[0] + '" aria-pressed="' + on + '">' + t[1] + '</button>';
        }).join('') + '</div></div>' +
        '<div class="chart-wrap">' + chartHtml(model) + '</div></div>';
    }

    var ck = state.check;
    var warn = ck && (ck.badDates > 0 || ck.nonSunday > 0);
    html += '<div class="footer">Weekly average = total attendance for each Sunday-dated week, averaged across the weeks in the period that have attendance. ' +
      '% change = (current weekly average minus prior weekly average) divided by prior weekly average. Monday through Saturday services roll to the following Sunday. Kids is Early Childhood plus Elementary.<br>' +
      'Data check: ' + (ck ? nf.format(ck.rows) : '0') + ' rows, ' + state.campuses.length + ' campuses, weeks ' + fmtDate(state.minWeek, true) + ' to ' + fmtDate(state.maxWeek, true) + '. Physical campuses exclude: ' + esc(CONFIG.excludedCampuses.join(', ')) + '.' +
      (missingExclusions().length ? ' <span class="warn">Not found in the data: ' + esc(missingExclusions().join(', ')) + '.</span>' : '') +
      (warn ? ' <span class="warn">' + ck.nonSunday + ' dates are not Sundays and ' + ck.badDates + ' dates could not be read. Check the SundayKey calculation.</span>' : '') +
      '</div>';

    app.innerHTML = html;

    if (activeFid) {
      var el = app.querySelector('[data-fid="' + activeFid + '"]');
      if (el) el.focus();
    }
  }

  // ------------------------------------------------------------------
  // Events
  // ------------------------------------------------------------------
  app.addEventListener('click', function (e) {
    var t = e.target.closest('[data-action]');
    if (!t) return;
    var a = t.getAttribute('data-action');
    if (a === 'campus-menu') state.campusOpen = !state.campusOpen;
    else if (a === 'campus-toggle') {
      var name = state.campuses[+t.getAttribute('data-index')];
      if (state.selected.has(name)) state.selected.delete(name); else state.selected.add(name);
    } else if (a === 'campus-all') state.selected = new Set(state.campuses);
    else if (a === 'campus-physical') state.selected = defaultCampusSet();
    else if (a === 'campus-none') state.selected = new Set();
    else if (a === 'preset') applyPreset(t.getAttribute('data-value'));
    else if (a === 'kids-toggle') state.kidsOpen = !state.kidsOpen;
    else if (a === 'tab') state.tab = t.getAttribute('data-value');
    else if (a === 'refresh') { refresh(); return; }
    else return;
    render();
  });

  app.addEventListener('change', function (e) {
    var t = e.target;
    var a = t.getAttribute && t.getAttribute('data-action');
    if (a !== 'start' && a !== 'end') return;
    var ms = parseISO(t.value);
    if (isNaN(ms)) return;
    if (a === 'start') {
      state.start = weekEnding(ms);
    } else {
      var snapped = ms - new Date(ms).getUTCDay() * DAY;
      var cap = latestEnd();
      state.end = snapped > cap ? cap : snapped;
    }
    state.preset = null;
    state.message = state.start > state.end ? 'Start date is after end date.' : null;
    render();
  });

  document.addEventListener('click', function (e) {
    if (state.campusOpen && !e.target.closest('.campus-wrap')) {
      state.campusOpen = false;
      render();
    }
  });
  document.addEventListener('keydown', function (e) {
    if (e.key === 'Escape' && state.campusOpen) {
      state.campusOpen = false;
      render();
      var b = app.querySelector('[data-fid="campus-menu"]');
      if (b) b.focus();
    }
  });

  // ------------------------------------------------------------------
  // Startup
  // ------------------------------------------------------------------
  async function refresh() {
    state.loading = true;
    state.error = null;
    render();
    try {
      await loadData();
      state.loading = false;
      if (state.preset) applyPreset(state.preset);
    } catch (err) {
      state.error = err && err.message ? err.message : String(err);
    }
    render();
  }

  async function init() {
    try {
      if (typeof tableau === 'undefined') {
        throw new Error('The Tableau Extensions API library was not found.\nPut tableau.extensions.1.latest.min.js in the lib folder next to index.html.');
      }
      await tableau.extensions.initializeAsync();
      await loadData();
      state.loading = false;
      applyPreset('4w');
    } catch (err) {
      state.loading = false;
      state.error = err && err.message ? err.message : String(err);
    }
    render();
  }

  init();
})();
