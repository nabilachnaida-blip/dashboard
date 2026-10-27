(function () {
  "use strict";

  var SHEET_ARRIVALS = "integration Usine";
  var US_BLUE = "#0B3F91";
  var US_OCEAN = "#0D6FA3";
  var TOP_DEPT_COUNT = 4; // table/chart keep only the most significant departments

  // Window this section reports on: "Depuis août jusqu'à aujourd'hui" —
  // from 1 August 2026 up to the visitor's today, recalculated on every load.
  var PERIOD_FROM = new Date(Date.UTC(2026, 7, 1)); // 1er août 2026
  function todayUTC() {
    var now = new Date();
    return new Date(Date.UTC(now.getFullYear(), now.getMonth(), now.getDate()));
  }

  function clean(v) { return v === null || v === undefined ? "" : String(v).trim(); }
  function norm(v) {
    var s = clean(v).toLowerCase().normalize("NFKD");
    var out = "";
    for (var i = 0; i < s.length; i++) {
      var code = s.charCodeAt(i);
      if (code < 0x0300 || code > 0x036f) out += s[i];
    }
    return out;
  }
  function fmt(n) { return (n || 0).toLocaleString("fr-FR"); }

  function parseDate(value) {
    if (value === null || value === undefined || value === "") return null;
    if (value instanceof Date) return new Date(Date.UTC(value.getFullYear(), value.getMonth(), value.getDate()));
    return null;
  }

  function sheetRows(ws) {
    if (!ws) return [];
    var range = ws["!ref"] ? XLSX.utils.decode_range(ws["!ref"]) : { s: { r: 0, c: 0 }, e: { r: 0, c: 0 } };
    return XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, range: { s: { r: 0, c: 0 }, e: range.e } });
  }

  function isoWeekInfo(date) {
    var d = new Date(Date.UTC(date.getUTCFullYear(), date.getUTCMonth(), date.getUTCDate()));
    var dayNum = (d.getUTCDay() + 6) % 7;
    d.setUTCDate(d.getUTCDate() - dayNum + 3);
    var firstThursday = new Date(Date.UTC(d.getUTCFullYear(), 0, 4));
    var firstDayNum = (firstThursday.getUTCDay() + 6) % 7;
    firstThursday.setUTCDate(firstThursday.getUTCDate() - firstDayNum + 3);
    var week = 1 + Math.round((d - firstThursday) / (7 * 24 * 3600 * 1000));
    return week;
  }

  function parseArrivals(rows) {
    var headerRow = -1, colMap = {};
    for (var r = 0; r < Math.min(rows.length, 5) && headerRow === -1; r++) {
      var row = rows[r] || [];
      for (var c = 0; c < row.length; c++) {
        var label = norm(row[c]);
        if (!label) continue;
        if (label.indexOf("departement") !== -1) { colMap.departement = c; headerRow = r; }
        else if (label.indexOf("effectif") !== -1) colMap.effectif = c;
        else if (label.indexOf("arriv") !== -1 && label.indexOf("usine") !== -1) colMap.date_arrivee = c;
      }
    }
    if (headerRow === -1 || colMap.departement === undefined || colMap.effectif === undefined) return [];
    var out = [];
    for (var i = headerRow + 1; i < rows.length; i++) {
      var rr = rows[i] || [];
      var dep = rr[colMap.departement], eff = rr[colMap.effectif];
      if (!clean(dep) || eff === null || eff === undefined || eff === "") continue;
      var dateArr = colMap.date_arrivee !== undefined ? parseDate(rr[colMap.date_arrivee]) : null;
      if (!dateArr) continue;
      var effInt = parseInt(eff, 10);
      if (isNaN(effInt)) continue;
      // Upper-cased so spelling variants ("Bird Cage" / "BIRD CAGE") merge — same as js/dashboard.js.
      out.push({ departement: clean(dep).replace(/\s+/g, " ").toUpperCase(), effectif: effInt, date_arrivee_usine: dateArr });
    }
    return out;
  }

  function topDepts(byDep) {
    var rows = Object.keys(byDep).map(function (d) { return { departement: d, effectif: byDep[d] }; });
    rows.sort(function (a, b) { return b.effectif - a.effectif; });
    return rows.slice(0, TOP_DEPT_COUNT);
  }

  function renderDeptTable(rows) {
    var tbody = document.getElementById("us-aout-dept-tbl-body");
    if (!tbody) return;
    tbody.innerHTML = rows.map(function (r) {
      return "<tr><td>" + r.departement + "</td><td>" + fmt(r.effectif) + "</td></tr>";
    }).join("");
  }

  function buildDeptChart(rows) {
    var dom = document.getElementById("us-chart-aout-dept");
    if (!dom || typeof echarts === "undefined") return;
    var sorted = rows.slice().sort(function (a, b) { return a.effectif - b.effectif; }); // ascending for horizontal bars
    var labels = sorted.map(function (r) { return r.departement; });
    var values = sorted.map(function (r) { return r.effectif; });
    var chart = echarts.init(dom);
    chart.setOption({
      color: [US_OCEAN],
      tooltip: { trigger: "axis", axisPointer: { type: "shadow" } },
      grid: { left: "3%", right: "10%", bottom: "3%", top: "5%", containLabel: true },
      xAxis: { type: "value", axisLabel: { fontSize: 11, color: "#1e293b" } },
      yAxis: { type: "category", data: labels, axisLabel: { fontSize: 12, color: "#1e293b", fontWeight: 600 } },
      series: [{
        name: "Intégrations", type: "bar", data: values, barMaxWidth: 28,
        itemStyle: { borderRadius: [0, 4, 4, 0] },
        label: { show: true, position: "right", fontSize: 12, color: "#1e293b", fontWeight: 700 }
      }]
    });
    window.addEventListener("resize", function () { chart.resize(); });
  }

  function buildWeekChart(byWeek) {
    var dom = document.getElementById("us-chart-aout-weekly");
    if (!dom || typeof echarts === "undefined") return;
    var weeks = Object.keys(byWeek).map(Number).sort(function (a, b) { return a - b; });
    var labels = weeks.map(function (w) { return "S" + w; });
    var values = weeks.map(function (w) { return byWeek[w]; });
    var chart = echarts.init(dom);
    chart.setOption({
      color: [US_BLUE],
      tooltip: { trigger: "axis", axisPointer: { type: "shadow" } },
      grid: { left: "3%", right: "4%", bottom: "3%", top: "10%", containLabel: true },
      xAxis: { type: "category", data: labels, axisLabel: { fontSize: 11, color: "#1e293b" } },
      yAxis: { type: "value", axisLabel: { fontSize: 11, color: "#1e293b" } },
      series: [{
        name: "Intégrations", type: "bar", data: values, barMaxWidth: 40,
        itemStyle: { borderRadius: [4, 4, 0, 0] },
        label: { show: true, position: "top", fontSize: 11, color: "#1e293b", fontWeight: 700 }
      }]
    });
    window.addEventListener("resize", function () { chart.resize(); });
  }

  // Monthly arrivals stacked by department: the top departments keep their own
  // colour, every other department is grouped under "Autres"; the month total sits on top.
  var DEPT_COLORS = [US_BLUE, US_OCEAN, "#3987e5", "#9ec5f4"];
  var OTHER_COLOR = "#cbd5e1";
  var MONTH_NAMES = ["Janv.", "Févr.", "Mars", "Avr.", "Mai", "Juin", "Juil.", "Août", "Sept.", "Oct.", "Nov.", "Déc."];
  function buildMonthDeptChart(byMonthDep, topNames) {
    var dom = document.getElementById("us-chart-aout-monthly");
    if (!dom || typeof echarts === "undefined") return;
    var months = Object.keys(byMonthDep).sort(); // "YYYY-MM"
    var labels = months.map(function (m) { return MONTH_NAMES[+m.slice(5) - 1] + " " + m.slice(0, 4); });
    var valueFor = function (w, keep) {
      var deps = byMonthDep[w], v = 0;
      Object.keys(deps).forEach(function (d) { if (keep(d)) v += deps[d]; });
      return v || null;
    };
    var series = topNames.map(function (name) {
      return { name: name, keep: function (d) { return d === name; } };
    }).concat([{ name: "Autres", keep: function (d) { return topNames.indexOf(d) === -1; } }])
      .map(function (s) {
        return {
          name: s.name, type: "bar", stack: "dept", barMaxWidth: 40,
          data: months.map(function (w) { return valueFor(w, s.keep); }),
          label: { show: true, fontSize: 10, color: "#fff", fontWeight: 700,
                   formatter: function (p) { return p.value >= 5 ? p.value : ""; } }
        };
      })
      .filter(function (s) { return s.data.some(function (v) { return v; }); });
    // Invisible series carrying the month total as a label above each stack
    series.push({
      name: "Total", type: "bar", stack: "dept", data: months.map(function () { return 0; }),
      itemStyle: { color: "transparent" }, tooltip: { show: false },
      label: { show: true, position: "top", fontSize: 12, color: "#1e293b", fontWeight: 700,
               formatter: function (p) { return fmt(valueFor(months[p.dataIndex], function () { return true; })); } }
    });
    var chart = echarts.init(dom);
    chart.setOption({
      color: DEPT_COLORS.slice(0, topNames.length).concat([OTHER_COLOR]),
      tooltip: { trigger: "axis", axisPointer: { type: "shadow" } },
      legend: { top: 0, data: series.map(function (s) { return s.name; }).filter(function (n) { return n !== "Total"; }), textStyle: { fontSize: 11, color: "#1e293b" } },
      grid: { left: "3%", right: "4%", bottom: "3%", top: 40, containLabel: true },
      xAxis: { type: "category", data: labels, axisLabel: { fontSize: 11, color: "#1e293b" } },
      yAxis: { type: "value", axisLabel: { fontSize: 11, color: "#1e293b" } },
      series: series
    });
    window.addEventListener("resize", function () { chart.resize(); });
  }

  // suivi.xlsx servi à côté de la page ; sinon (double-clic, file://) la copie intégrée à la page
  function loadSuiviBuffer() {
    return fetch("suivi.xlsx", { cache: "no-store" }) // always read the latest saved file, never a cached copy
      .then(function (r) {
        if (!r.ok) throw new Error("Impossible de charger suivi.xlsx (HTTP " + r.status + ")");
        return r.arrayBuffer();
      })
      .catch(function (e) {
        if (!window.US_SUIVI_B64) throw e;
        var bin = atob(window.US_SUIVI_B64), buf = new Uint8Array(bin.length);
        for (var i = 0; i < bin.length; i++) buf[i] = bin.charCodeAt(i);
        return buf.buffer;
      });
  }

  function loadData() {
    loadSuiviBuffer()
      .then(function (buf) {
        var wb = XLSX.read(buf, { type: "array", cellDates: true });
        var arrivals = parseArrivals(sheetRows(wb.Sheets[SHEET_ARRIVALS]));
        var periodTo = todayUTC();
        var period = arrivals.filter(function (a) {
          return a.date_arrivee_usine >= PERIOD_FROM && a.date_arrivee_usine <= periodTo;
        });

        var byDep = {}, byWeek = {}, byMonthDep = {}, total = 0;
        period.forEach(function (a) {
          byDep[a.departement] = (byDep[a.departement] || 0) + a.effectif;
          var w = isoWeekInfo(a.date_arrivee_usine);
          byWeek[w] = (byWeek[w] || 0) + a.effectif;
          var m = a.date_arrivee_usine.toISOString().slice(0, 7);
          byMonthDep[m] = byMonthDep[m] || {};
          byMonthDep[m][a.departement] = (byMonthDep[m][a.departement] || 0) + a.effectif;
          total += a.effectif;
        });

        var totalEl = document.getElementById("us-aout-verified-total");
        if (totalEl) totalEl.textContent = fmt(total);

        // Headline total — animated by reveal.js; if the card was already
        // revealed before the file finished loading, animate it now.
        var kpiEl = document.getElementById("us-aout-kpi-total");
        if (kpiEl) {
          kpiEl.setAttribute("data-count-to", String(total));
          var group = kpiEl.closest(".us-reveal-stagger");
          if (group && group.classList.contains("us-visible") && window.usAnimateCount) window.usAnimateCount(kpiEl);
        }

        var weeks = Object.keys(byWeek).map(Number).sort(function (a, b) { return a - b; });
        var rangeEl = document.getElementById("us-aout-week-range");
        if (rangeEl && weeks.length) rangeEl.textContent = "Progression S" + weeks[0] + " → S" + weeks[weeks.length - 1];

        var top = topDepts(byDep);
        renderDeptTable(top);
        buildDeptChart(top);
        buildWeekChart(byWeek);
        buildMonthDeptChart(byMonthDep, top.map(function (r) { return r.departement; }));
      })
      .catch(function (e) {
        var box = document.getElementById("us-aout-data-error");
        if (box) { box.style.display = "block"; box.textContent = "Données détaillées indisponibles : " + e.message; }
      });
  }

  document.addEventListener("DOMContentLoaded", function () {
    if (!document.getElementById("us-chart-aout-weekly")) return;
    if (!document.body.classList.contains("us-locked")) { loadData(); return; }
    var observer = new MutationObserver(function () {
      if (!document.body.classList.contains("us-locked")) { observer.disconnect(); loadData(); }
    });
    observer.observe(document.body, { attributes: true, attributeFilter: ["class"] });
  });
})();
