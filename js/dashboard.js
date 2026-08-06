/* Suivi Usine — fully static dashboard.
 * Reads suivi.xlsx client-side (SheetJS), computes the same KPIs as the
 * server versions, and renders with ECharts. No backend: update the data
 * by replacing suivi.xlsx in this folder and pushing to GitHub. */
(function () {
  "use strict";

  var SHEET_ARRIVALS = "integration Usine";
  var SHEET_FORMATION_WEEKLY = "Feuil3";
  var SHEET_FORMATION_IFMIA = "En formation IFMIA";
  var SHEET_MANUAL_KPIS = "Indicateurs Manuels";
  var ATELIER_MAP = { ferrage: "FERRAGE", montage: "MONTAGE", peinture: "PEINTURE" };

  var US_BLUE = "#0B3F91";
  var US_ORANGE = "#FF5A1F";
  var US_CAT_COLORS = ["#0B3F91", "#0D6FA3", "#00B3B8", "#38D6C4", "#6EE7D2", "#9FEDE2", "#FF5A1F", "#4DA8DA"];

  // ── text helpers ─────────────────────────────────────────────────────
  function clean(v) { return v === null || v === undefined ? "" : String(v).trim(); }
  function norm(v) {
    var s = clean(v).toLowerCase().normalize("NFKD");
    var out = "";
    for (var i = 0; i < s.length; i++) {
      var code = s.charCodeAt(i);
      if (code < 0x0300 || code > 0x036f) out += s[i]; // skip combining diacritics
    }
    return out;
  }

  // ── date helpers — everything is UTC-midnight internally, always read
  //    with UTC getters, so results never drift with the visitor's timezone ──
  function parseDate(value) {
    if (value === null || value === undefined || value === "") return null;
    // SheetJS builds cellDates using local-time semantics, so the calendar
    // day it intended must be read back with local (not UTC) getters —
    // verified empirically against the vendored xlsx.full.min.js.
    if (value instanceof Date) return new Date(Date.UTC(value.getFullYear(), value.getMonth(), value.getDate()));
    var s = clean(value);
    if (!s || s === "N/A") return null;
    var m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (m) return new Date(Date.UTC(+m[1], +m[2] - 1, +m[3]));
    m = s.match(/^(\d{2})\/(\d{2})\/(\d{4})/);
    if (m) return new Date(Date.UTC(+m[3], +m[2] - 1, +m[1]));
    return null;
  }

  function todayUTC() {
    var now = new Date();
    return new Date(Date.UTC(now.getFullYear(), now.getMonth(), now.getDate()));
  }

  function isoWeekInfo(date) {
    var d = new Date(Date.UTC(date.getUTCFullYear(), date.getUTCMonth(), date.getUTCDate()));
    var dayNum = (d.getUTCDay() + 6) % 7;
    d.setUTCDate(d.getUTCDate() - dayNum + 3);
    var firstThursday = new Date(Date.UTC(d.getUTCFullYear(), 0, 4));
    var firstDayNum = (firstThursday.getUTCDay() + 6) % 7;
    firstThursday.setUTCDate(firstThursday.getUTCDate() - firstDayNum + 3);
    var week = 1 + Math.round((d - firstThursday) / (7 * 24 * 3600 * 1000));
    return { isoYear: d.getUTCFullYear(), isoWeek: week };
  }

  function isoWeekMonday(year, week) {
    var jan4 = new Date(Date.UTC(year, 0, 4));
    var jan4DayNum = (jan4.getUTCDay() + 6) % 7;
    var week1Monday = new Date(jan4);
    week1Monday.setUTCDate(jan4.getUTCDate() - jan4DayNum);
    var target = new Date(week1Monday);
    target.setUTCDate(week1Monday.getUTCDate() + (week - 1) * 7);
    return target;
  }

  function fmtDate(d) {
    var y = d.getUTCFullYear(), m = String(d.getUTCMonth() + 1).padStart(2, "0"), day = String(d.getUTCDate()).padStart(2, "0");
    return y + "-" + m + "-" + day;
  }

  // ── Excel parsing (mirrors excel_parser.py) ─────────────────────────
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
        else if (label.indexOf("contrat") !== -1) colMap.date_contrat = c;
      }
    }
    if (headerRow === -1 || colMap.departement === undefined || colMap.effectif === undefined) {
      throw new Error("Feuille '" + SHEET_ARRIVALS + "' : en-têtes introuvables");
    }
    var out = [];
    for (var i = headerRow + 1; i < rows.length; i++) {
      var rr = rows[i] || [];
      var dep = rr[colMap.departement], eff = rr[colMap.effectif];
      if (!clean(dep) || eff === null || eff === undefined || eff === "") continue;
      var dateArr = colMap.date_arrivee !== undefined ? parseDate(rr[colMap.date_arrivee]) : null;
      var dateCtr = colMap.date_contrat !== undefined ? parseDate(rr[colMap.date_contrat]) : null;
      if (!dateArr) continue;
      var effInt = parseInt(eff, 10);
      if (isNaN(effInt)) continue;
      out.push({ departement: clean(dep), effectif: effInt, date_arrivee_usine: dateArr, date_debut_contrat: dateCtr });
    }
    return out;
  }

  function parseFormationSemaine(rows, anchorYear) {
    var weekRow = -1, weekCols = {};
    for (var r = 0; r < Math.min(rows.length, 5) && weekRow === -1; r++) {
      var row = rows[r] || [], cols = {};
      for (var c = 0; c < row.length; c++) {
        var v = norm(row[c]).replace(/\s+/g, "");
        var m = v.match(/^s(\d{1,2})$/);
        if (m) cols[c] = parseInt(m[1], 10);
      }
      if (Object.keys(cols).length) { weekRow = r; weekCols = cols; }
    }
    if (weekRow === -1) return [];

    var merged = {}, currentAtelier = null;
    for (var i = weekRow + 1; i < rows.length; i++) {
      var rr = rows[i] || [];
      var labelA = norm(rr[0]), labelB = norm(rr[1]);
      if (ATELIER_MAP[labelA]) currentAtelier = ATELIER_MAP[labelA];
      if (!currentAtelier) continue;
      var kind = null;
      if (labelB.indexOf("estim") === 0) kind = "estimation";
      else if (labelB.indexOf("reel") === 0) kind = "reel";
      else continue;
      Object.keys(weekCols).forEach(function (colStr) {
        var col = Number(colStr), week = weekCols[colStr];
        var val = rr[col];
        if (val === null || val === undefined || val === "") return;
        var valInt = parseInt(val, 10);
        if (isNaN(valInt)) return;
        var key = week + "_" + currentAtelier;
        if (!merged[key]) merged[key] = { iso_year: anchorYear, iso_week: week, atelier: currentAtelier, estimation: 0, reel: null };
        merged[key][kind] = valInt;
      });
    }
    return Object.keys(merged).map(function (k) { return merged[k]; });
  }

  function parseFormationIfmia(rows) {
    var headerRow = -1;
    for (var r = 0; r < Math.min(rows.length, 10); r++) {
      if (norm((rows[r] || [])[0]) === "ur") { headerRow = r; break; }
    }
    if (headerRow === -1) return [];
    var out = [];
    for (var i = headerRow + 1; i < rows.length; i++) {
      var rr = rows[i] || [];
      var ur = rr[0], eff = rr[1];
      if (!clean(ur) || eff === null || eff === undefined || eff === "") continue;
      var dateIfmia = parseDate(rr[2]), dateUsine = parseDate(rr[3]);
      if (!dateIfmia) continue;
      var effInt = parseInt(eff, 10);
      if (isNaN(effInt)) continue;
      out.push({ ur: clean(ur), effectif: effInt, date_integration_ifmia: dateIfmia, date_integration_usine: dateUsine });
    }
    return out;
  }

  // Reads the "Indicateurs Manuels" sheet: one row per indicator (Appels
  // téléphoniques, Visite médicale), values entered by hand per week —
  // same S30/S31/... week-header layout as Feuil3, but flat (no
  // Estimation/Réel split). Optional sheet: older workbooks without it
  // simply yield no manual KPIs.
  function parseManualKpis(rows, anchorYear) {
    var weekRow = -1, weekCols = {};
    for (var r = 0; r < Math.min(rows.length, 5) && weekRow === -1; r++) {
      var row = rows[r] || [], cols = {};
      for (var c = 0; c < row.length; c++) {
        var v = norm(row[c]).replace(/\s+/g, "");
        var m = v.match(/^s(\d{1,2})$/);
        if (m) cols[c] = parseInt(m[1], 10);
      }
      if (Object.keys(cols).length) { weekRow = r; weekCols = cols; }
    }
    if (weekRow === -1) return {};

    var byWeek = {};
    function ensure(week) {
      var key = anchorYear + "_" + week;
      if (!byWeek[key]) byWeek[key] = { iso_year: anchorYear, iso_week: week, appels: null, visite: null };
      return byWeek[key];
    }

    for (var i = weekRow + 1; i < rows.length; i++) {
      var rr = rows[i] || [];
      var label = norm(rr[0]);
      if (!label) continue;
      var field = null;
      if (label.indexOf("appel") !== -1) field = "appels";
      else if (label.indexOf("visite") !== -1 && label.indexOf("medic") !== -1) field = "visite";
      else continue;
      Object.keys(weekCols).forEach(function (colStr) {
        var col = Number(colStr), week = weekCols[colStr];
        var val = rr[col];
        if (val === null || val === undefined || val === "") return;
        var valInt = parseInt(val, 10);
        if (isNaN(valInt)) return;
        ensure(week)[field] = valInt;
      });
    }
    return byWeek;
  }

  function parseSuiviWorkbook(arrayBuffer) {
    var wb = XLSX.read(arrayBuffer, { type: "array", cellDates: true });
    var missing = [SHEET_ARRIVALS, SHEET_FORMATION_WEEKLY, SHEET_FORMATION_IFMIA].filter(function (s) {
      return wb.SheetNames.indexOf(s) === -1;
    });
    if (missing.length) throw new Error("Feuille(s) manquante(s) dans le classeur : " + missing.join(", "));

    var arrivalsRows = XLSX.utils.sheet_to_json(wb.Sheets[SHEET_ARRIVALS], { header: 1, raw: true, defval: null });
    var arrivals = parseArrivals(arrivalsRows);
    if (!arrivals.length) throw new Error("Aucune donnée exploitable dans la feuille '" + SHEET_ARRIVALS + "'");

    var minDate = arrivals[0].date_arrivee_usine;
    arrivals.forEach(function (a) { if (a.date_arrivee_usine < minDate) minDate = a.date_arrivee_usine; });
    var anchorYear = isoWeekInfo(minDate).isoYear;

    var fsRows = XLSX.utils.sheet_to_json(wb.Sheets[SHEET_FORMATION_WEEKLY], { header: 1, raw: true, defval: null });
    var ifmiaRows = XLSX.utils.sheet_to_json(wb.Sheets[SHEET_FORMATION_IFMIA], { header: 1, raw: true, defval: null });

    var manualKpis = {};
    if (wb.SheetNames.indexOf(SHEET_MANUAL_KPIS) !== -1) {
      var mkRows = XLSX.utils.sheet_to_json(wb.Sheets[SHEET_MANUAL_KPIS], { header: 1, raw: true, defval: null });
      manualKpis = parseManualKpis(mkRows, anchorYear);
    }

    return {
      arrivals: arrivals,
      formation_semaine: parseFormationSemaine(fsRows, anchorYear),
      formation_ifmia: parseFormationIfmia(ifmiaRows),
      manual_kpis: manualKpis,
    };
  }

  // ── KPI computation (mirrors app.py's /api/data) ────────────────────
  function computeDashboardData(parsed, filters) {
    var presentCodes = Array.from(new Set(parsed.arrivals.map(function (a) { return a.departement; }))).sort();
    if (!presentCodes.length) {
      return { empty: true, departments: [] };
    }

    var depParam = filters.departement && presentCodes.indexOf(filters.departement) !== -1 ? filters.departement : "";
    var dateFrom = filters.date_from ? parseDate(filters.date_from) : null;
    var dateTo = filters.date_to ? parseDate(filters.date_to) : null;

    var scopeCodes = depParam ? [depParam] : presentCodes;

    var arrivals = parsed.arrivals.filter(function (a) { return scopeCodes.indexOf(a.departement) !== -1; });
    if (dateFrom) arrivals = arrivals.filter(function (a) { return a.date_arrivee_usine >= dateFrom; });
    if (dateTo) arrivals = arrivals.filter(function (a) { return a.date_arrivee_usine <= dateTo; });

    var minRelevantDate = null;
    if (!dateFrom && parsed.formation_semaine.length) {
      var earliest = parsed.formation_semaine.reduce(function (acc, r) {
        if (!acc) return r;
        return (r.iso_year < acc.iso_year || (r.iso_year === acc.iso_year && r.iso_week < acc.iso_week)) ? r : acc;
      }, null);
      if (earliest) minRelevantDate = isoWeekMonday(earliest.iso_year, earliest.iso_week);
    }

    var weekly = {};
    arrivals.forEach(function (a) {
      if (minRelevantDate && a.date_arrivee_usine < minRelevantDate) return;
      var info = isoWeekInfo(a.date_arrivee_usine);
      var key = info.isoYear + "_" + info.isoWeek;
      if (!weekly[key]) weekly[key] = { year: info.isoYear, week: info.isoWeek, value: 0 };
      weekly[key].value += a.effectif;
    });
    var weeklySorted = Object.keys(weekly).map(function (k) { return weekly[k]; })
      .sort(function (x, y) { return x.year - y.year || x.week - y.week; });
    var weeklyLabels = weeklySorted.map(function (w) { return "S" + w.week; });
    var weeklyValues = weeklySorted.map(function (w) { return w.value; });

    var today = todayUTC();
    var cur = isoWeekInfo(today);
    var thisWeekEntry = weeklySorted.filter(function (w) { return w.year === cur.isoYear && w.week === cur.isoWeek; })[0];
    var thisWeekTotal = thisWeekEntry ? thisWeekEntry.value : 0;

    var refDate = dateTo || today;
    var monthTotal = 0;
    arrivals.forEach(function (a) {
      if (a.date_arrivee_usine.getUTCFullYear() === refDate.getUTCFullYear() && a.date_arrivee_usine.getUTCMonth() === refDate.getUTCMonth()) {
        monthTotal += a.effectif;
      }
    });

    var fsByWeek = {};
    parsed.formation_semaine.forEach(function (fs) {
      var key = fs.iso_year + "_" + fs.iso_week;
      if (!fsByWeek[key]) fsByWeek[key] = { year: fs.iso_year, week: fs.iso_week, estimation: 0, reel: 0, reelKnown: false };
      fsByWeek[key].estimation += fs.estimation || 0;
      if (fs.reel !== null && fs.reel !== undefined) {
        fsByWeek[key].reel += fs.reel;
        fsByWeek[key].reelKnown = true;
      }
    });
    var fsSorted = Object.keys(fsByWeek).map(function (k) { return fsByWeek[k]; })
      .sort(function (x, y) { return x.year - y.year || x.week - y.week; });
    var formationLabels = fsSorted.map(function (w) { return "S" + w.week; });
    var formationEstimation = fsSorted.map(function (w) { return w.estimation; });
    var formationReel = fsSorted.map(function (w) { return w.reelKnown ? w.reel : null; });

    var totalFormationReel = 0;
    var knownWeeks = fsSorted.filter(function (w) { return w.reelKnown; });
    if (knownWeeks.length) totalFormationReel = knownWeeks[knownWeeks.length - 1].reel;

    var nextWeekNum = cur.isoWeek < 52 ? cur.isoWeek + 1 : 1;
    var nextYear = cur.isoWeek < 52 ? cur.isoYear : cur.isoYear + 1;
    var nextEntry = fsByWeek[nextYear + "_" + nextWeekNum];
    var nextWeek = {
      iso_year: nextYear, iso_week: nextWeekNum, label: "S" + nextWeekNum,
      estimation: nextEntry ? nextEntry.estimation : null,
    };

    var upcomingByDept = {};
    arrivals.forEach(function (a) {
      if (a.date_debut_contrat && a.date_debut_contrat > today) {
        upcomingByDept[a.departement] = (upcomingByDept[a.departement] || 0) + a.effectif;
      }
    });
    var upcomingTotal = 0;
    Object.keys(upcomingByDept).forEach(function (k) { upcomingTotal += upcomingByDept[k]; });
    var upcomingList = Object.keys(upcomingByDept).map(function (d) { return { departement: d, label: d, effectif: upcomingByDept[d] }; })
      .sort(function (a, b) { return b.effectif - a.effectif; });

    var ifmia = parsed.formation_ifmia;
    if (depParam) ifmia = ifmia.filter(function (i) { return i.ur === depParam; });
    ifmia = ifmia.slice().sort(function (a, b) { return b.date_integration_ifmia - a.date_integration_ifmia; }).slice(0, 50);

    // Manual weekly indicators (Appels téléphoniques, Visite médicale) —
    // plant-wide counts entered by hand in Excel, not department-scoped.
    var manualEntry = (parsed.manual_kpis || {})[cur.isoYear + "_" + cur.isoWeek];
    var appelsThisWeek = manualEntry && manualEntry.appels !== null ? manualEntry.appels : null;
    var visiteThisWeek = manualEntry && manualEntry.visite !== null ? manualEntry.visite : null;

    return {
      empty: false,
      departments: presentCodes.map(function (c) { return { code: c, label: c }; }),
      selected_departement: depParam,
      weekly_arrivals: { labels: weeklyLabels, values: weeklyValues },
      this_week_total: thisWeekTotal,
      this_week_label: "S" + cur.isoWeek,
      month_total: monthTotal,
      formation: { labels: formationLabels, estimation: formationEstimation, reel: formationReel },
      total_formation_reel: totalFormationReel,
      next_week: nextWeek,
      appels_this_week: appelsThisWeek,
      visite_this_week: visiteThisWeek,
      upcoming_contracts: { total: upcomingTotal, by_departement: upcomingList },
      formation_ifmia_detail: ifmia.map(function (i) {
        return {
          ur: i.ur, effectif: i.effectif,
          date_integration_ifmia: fmtDate(i.date_integration_ifmia),
          date_integration_usine: i.date_integration_usine ? fmtDate(i.date_integration_usine) : null,
        };
      }),
    };
  }

  // ── rendering ────────────────────────────────────────────────────────
  var US_CHARTS = {};
  var parsedWorkbook = null;
  var deptsLoaded = false;

  function fmt(n) { return (n || 0).toLocaleString("fr-FR"); }

  function buildWeeklyChart(labels, values, thisWeekLabel) {
    var dom = document.getElementById("us-chart-weekly");
    if (!dom) return;
    if (US_CHARTS.weekly) US_CHARTS.weekly.dispose();
    var chart = echarts.init(dom);
    US_CHARTS.weekly = chart;
    chart.setOption({
      color: [US_BLUE],
      tooltip: { trigger: "axis", axisPointer: { type: "shadow" } },
      grid: { left: "3%", right: "4%", bottom: "3%", top: "10%", containLabel: true },
      xAxis: { type: "category", data: labels, axisLabel: { fontSize: 11 } },
      yAxis: { type: "value", axisLabel: { fontSize: 11 } },
      series: [{
        name: "Entrées", type: "bar", data: values, barMaxWidth: 32,
        itemStyle: { borderRadius: [4, 4, 0, 0], color: function (p) { return labels[p.dataIndex] === thisWeekLabel ? US_ORANGE : US_BLUE; } }
      }]
    });
  }

  function buildFormationChart(labels, estimation, reel) {
    var dom = document.getElementById("us-chart-formation");
    if (!dom) return;
    if (US_CHARTS.formation) US_CHARTS.formation.dispose();
    var chart = echarts.init(dom);
    US_CHARTS.formation = chart;

    // Écart = Réel - Estimation, shown for the most recent week that has a known Réel.
    var lastKnownIdx = -1;
    for (var i = reel.length - 1; i >= 0; i--) {
      if (reel[i] !== null && reel[i] !== undefined) { lastKnownIdx = i; break; }
    }
    var reelSeries = { name: "Réel", type: "bar", data: reel, barMaxWidth: 28, itemStyle: { color: US_BLUE, borderRadius: [4, 4, 0, 0] } };
    if (lastKnownIdx !== -1) {
      var ecart = reel[lastKnownIdx] - estimation[lastKnownIdx];
      reelSeries.markPoint = {
        symbol: "circle", symbolSize: 0,
        data: [{
          coord: [lastKnownIdx, reel[lastKnownIdx]],
          label: {
            show: true, position: "top", distance: 14,
            formatter: function () { return "Écart (" + labels[lastKnownIdx] + ") : " + (ecart > 0 ? "+" : "") + ecart; },
            color: US_ORANGE, fontWeight: "bold", fontSize: 12,
            backgroundColor: "#FFF3EF", borderColor: US_ORANGE, borderWidth: 1, borderRadius: 4, padding: [4, 8]
          }
        }]
      };
    }

    chart.setOption({
      tooltip: { trigger: "axis" },
      legend: { data: ["Estimation", "Réel"], top: 0, textStyle: { fontSize: 12 } },
      grid: { left: "3%", right: "4%", bottom: "3%", top: "18%", containLabel: true },
      xAxis: { type: "category", data: labels, axisLabel: { fontSize: 11 } },
      yAxis: { type: "value", axisLabel: { fontSize: 11 } },
      series: [
        reelSeries,
        { name: "Estimation", type: "line", data: estimation, smooth: false, symbol: "circle", symbolSize: 6, lineStyle: { width: 3, color: US_ORANGE }, itemStyle: { color: US_ORANGE } }
      ]
    });
  }

  function buildUpcomingChart(byDept) {
    var dom = document.getElementById("us-chart-upcoming");
    if (!dom) return;
    if (US_CHARTS.upcoming) US_CHARTS.upcoming.dispose();
    var chart = echarts.init(dom);
    US_CHARTS.upcoming = chart;
    var labels = byDept.map(function (d) { return d.label; });
    var values = byDept.map(function (d) { return d.effectif; });
    chart.setOption({
      tooltip: { trigger: "axis", axisPointer: { type: "shadow" } },
      grid: { left: "3%", right: "8%", bottom: "3%", top: "5%", containLabel: true },
      xAxis: { type: "value", axisLabel: { fontSize: 11 } },
      yAxis: { type: "category", data: labels, axisLabel: { fontSize: 11 }, inverse: true },
      series: [{
        name: "Contrats à venir", type: "bar", data: values, barMaxWidth: 22,
        itemStyle: { borderRadius: [0, 4, 4, 0], color: function (p) { return US_CAT_COLORS[p.dataIndex % US_CAT_COLORS.length]; } },
        label: { show: true, position: "right", fontSize: 11, color: "#1e293b" }
      }]
    });
  }

  function renderIfmiaTable(rows) {
    var tbody = document.getElementById("us-ifmia-tbl-body");
    if (!tbody) return;
    if (!rows.length) {
      tbody.innerHTML = '<tr><td colspan="4" style="text-align:center;color:#94a3b8;padding:16px;">Aucune donnée</td></tr>';
      return;
    }
    tbody.innerHTML = rows.map(function (r) {
      return "<tr><td>" + r.ur + "</td><td>" + fmt(r.effectif) + "</td><td>" + r.date_integration_ifmia + "</td><td>" + (r.date_integration_usine || "—") + "</td></tr>";
    }).join("");
  }

  function populateDepartments(depts, selected) {
    var sel = document.getElementById("us-dept-select");
    if (!sel || deptsLoaded) return;
    depts.forEach(function (d) {
      var opt = document.createElement("option");
      opt.value = d.code; opt.textContent = d.label;
      sel.appendChild(opt);
    });
    if (selected) sel.value = selected;
    deptsLoaded = true;
  }

  function render(d) {
    document.getElementById("us-content").style.display = "block";
    document.getElementById("us-empty").style.display = "none";

    populateDepartments(d.departments, d.selected_departement);

    document.getElementById("us-this-week-label").textContent = d.this_week_label || "cette semaine";
    document.getElementById("us-kpi-week").textContent = fmt(d.this_week_total);
    document.getElementById("us-kpi-month").textContent = fmt(d.month_total);
    document.getElementById("us-kpi-formation").textContent = fmt(d.total_formation_reel);
    document.getElementById("us-kpi-upcoming").textContent = fmt(d.upcoming_contracts.total);

    document.getElementById("us-next-week-label").textContent = d.next_week ? d.next_week.label : "S+1";
    var nextEl = document.getElementById("us-kpi-next-week");
    if (d.next_week && d.next_week.estimation !== null && d.next_week.estimation !== undefined) {
      nextEl.textContent = fmt(d.next_week.estimation);
    } else {
      nextEl.textContent = "Non renseigné";
    }

    document.getElementById("us-appels-week-label").textContent = d.this_week_label || "cette semaine";
    document.getElementById("us-kpi-appels").textContent = d.appels_this_week !== null ? fmt(d.appels_this_week) : "Non renseigné";
    document.getElementById("us-visite-week-label").textContent = d.this_week_label || "cette semaine";
    document.getElementById("us-kpi-visite").textContent = d.visite_this_week !== null ? fmt(d.visite_this_week) : "Non renseigné";

    setTimeout(function () {
      buildWeeklyChart(d.weekly_arrivals.labels, d.weekly_arrivals.values, d.this_week_label);
      buildFormationChart(d.formation.labels, d.formation.estimation, d.formation.reel);
      buildUpcomingChart(d.upcoming_contracts.by_departement);
    }, 50);

    renderIfmiaTable(d.formation_ifmia_detail);
  }

  function refresh() {
    if (!parsedWorkbook) return;
    var errEl = document.getElementById("us-error");
    if (errEl) errEl.style.display = "none";
    try {
      var filters = {
        departement: document.getElementById("us-dept-select").value,
        date_from: document.getElementById("us-date-from").value,
        date_to: document.getElementById("us-date-to").value,
      };
      var d = computeDashboardData(parsedWorkbook, filters);
      if (d.empty) {
        document.getElementById("us-content").style.display = "none";
        document.getElementById("us-empty").style.display = "block";
        return;
      }
      render(d);
    } catch (e) {
      if (errEl) { errEl.style.display = "block"; errEl.textContent = "Erreur : " + e.message; }
    }
  }

  window.usResetFilters = function () {
    document.getElementById("us-dept-select").value = "";
    document.getElementById("us-date-from").value = "";
    document.getElementById("us-date-to").value = "";
    refresh();
  };

  function init() {
    var errEl = document.getElementById("us-error");
    fetch("suivi.xlsx")
      .then(function (r) {
        if (!r.ok) throw new Error("Impossible de charger suivi.xlsx (HTTP " + r.status + ")");
        return r.arrayBuffer();
      })
      .then(function (buf) {
        parsedWorkbook = parseSuiviWorkbook(buf);
        refresh();
      })
      .catch(function (e) {
        if (errEl) { errEl.style.display = "block"; errEl.textContent = "Erreur de chargement : " + e.message; }
      });

    document.getElementById("us-dept-select").addEventListener("change", refresh);
    document.getElementById("us-date-from").addEventListener("change", refresh);
    document.getElementById("us-date-to").addEventListener("change", refresh);
  }

  window.addEventListener("resize", function () {
    Object.keys(US_CHARTS).forEach(function (k) { try { US_CHARTS[k].resize(); } catch (e) {} });
  });

  document.addEventListener("DOMContentLoaded", init);
})();
