(function () {
  "use strict";

  var SHEET_ARRIVALS = "integration Usine";
  var US_BLUE = "#0B3F91";
  var US_OCEAN = "#0D6FA3";
  var TOP_DEPT_COUNT = 4; // table/chart keep only the most significant departments

  // Campaign window this bulletin reports on — fixed, not "today minus N days":
  // it matches the 622 figure exactly (arrivals already logged through the
  // end of the arrêt annuel campaign, a few days past the publish date).
  var PERIOD_FROM = new Date(Date.UTC(2026, 6, 30)); // 30 juillet 2026
  var PERIOD_TO = new Date(Date.UTC(2026, 7, 27)); // 27 août 2026

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
      out.push({ departement: clean(dep), effectif: effInt, date_arrivee_usine: dateArr });
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

  function loadData() {
    fetch("suivi.xlsx")
      .then(function (r) {
        if (!r.ok) throw new Error("HTTP " + r.status);
        return r.arrayBuffer();
      })
      .then(function (buf) {
        var wb = XLSX.read(buf, { type: "array", cellDates: true });
        var arrivals = parseArrivals(sheetRows(wb.Sheets[SHEET_ARRIVALS]));
        var period = arrivals.filter(function (a) {
          return a.date_arrivee_usine >= PERIOD_FROM && a.date_arrivee_usine <= PERIOD_TO;
        });

        var byDep = {}, byWeek = {}, total = 0;
        period.forEach(function (a) {
          byDep[a.departement] = (byDep[a.departement] || 0) + a.effectif;
          var w = isoWeekInfo(a.date_arrivee_usine);
          byWeek[w] = (byWeek[w] || 0) + a.effectif;
          total += a.effectif;
        });

        var totalEl = document.getElementById("us-aout-verified-total");
        if (totalEl) totalEl.textContent = fmt(total);

        var top = topDepts(byDep);
        renderDeptTable(top);
        buildDeptChart(top);
        buildWeekChart(byWeek);
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
