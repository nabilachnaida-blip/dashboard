/* Suivi Usine — fully static dashboard.
 * Reads suivi.xlsx client-side (SheetJS), computes the same KPIs as the
 * server versions, and renders with ECharts. No backend: update the data
 * by replacing suivi.xlsx in this folder and pushing to GitHub. */
(function () {
  "use strict";

  var SHEET_ARRIVALS = "integration Usine";
  var SHEET_FORMATION_IFMIA = "En formation IFMIA";
  var SHEET_WEEKLY_INDICATORS = "Indicateurs Hebdomadaires";
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

  // Reads the single "Indicateurs Hebdomadaires" sheet: every manually
  // entered weekly static value lives here as one row (Ferrage/Montage/
  // Peinture × Estimation/Réel, Appels téléphoniques, Visite médicale),
  // S30/S31/... as columns. Returns both the per-atelier training series
  // and the flat manual-KPI series in one pass.
  function parseIndicateursHebdomadaires(rows, anchorYear) {
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
    if (weekRow === -1) return { formation_semaine: [], manual_kpis: {} };

    var fsMerged = {};
    var manualByWeek = {};
    function ensureManual(week) {
      var key = anchorYear + "_" + week;
      if (!manualByWeek[key]) {
        manualByWeek[key] = {
          iso_year: anchorYear, iso_week: week,
          appels: null, visite: null,
          ifmia_diplomes: null, ifmia_non_diplomes: null,
        };
      }
      return manualByWeek[key];
    }

    for (var i = weekRow + 1; i < rows.length; i++) {
      var rr = rows[i] || [];
      var label = norm(rr[0]);
      if (!label) continue;

      var atelier = null, kind = null, manualField = null;
      Object.keys(ATELIER_MAP).forEach(function (k) { if (label.indexOf(k) !== -1) atelier = ATELIER_MAP[k]; });
      if (atelier) {
        if (label.indexOf("estim") !== -1) kind = "estimation";
        else if (label.indexOf("reel") !== -1) kind = "reel";
      } else if (label.indexOf("ifmia") !== -1 && label.indexOf("non") !== -1 && label.indexOf("diplom") !== -1) {
        manualField = "ifmia_non_diplomes";
      } else if (label.indexOf("ifmia") !== -1 && label.indexOf("diplom") !== -1) {
        manualField = "ifmia_diplomes";
      } else if (label.indexOf("appel") !== -1) {
        manualField = "appels";
      } else if (label.indexOf("visite") !== -1 && label.indexOf("medic") !== -1) {
        manualField = "visite";
      }
      if (!kind && !manualField) continue;

      Object.keys(weekCols).forEach(function (colStr) {
        var col = Number(colStr), week = weekCols[colStr];
        var val = rr[col];
        if (val === null || val === undefined || val === "") return;
        var valInt = parseInt(val, 10);
        if (isNaN(valInt)) return;
        if (kind) {
          var key = week + "_" + atelier;
          if (!fsMerged[key]) fsMerged[key] = { iso_year: anchorYear, iso_week: week, atelier: atelier, estimation: 0, reel: null };
          fsMerged[key][kind] = valInt;
        } else {
          ensureManual(week)[manualField] = valInt;
        }
      });
    }

    return {
      formation_semaine: Object.keys(fsMerged).map(function (k) { return fsMerged[k]; }),
      manual_kpis: manualByWeek,
    };
  }

  // Diplômés table (col A: UR/Effectifs/Date IFMIA/Date Usine).
  function parseFormationIfmiaTable(rows) {
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

  // Non-diplômés grid (col F onward: UR row × S30/S31/... week columns) —
  // one row per department, values entered by hand per week.
  function parseNonDiplomesGrid(rows, anchorYear) {
    var colOffset = 5; // column F, 0-indexed
    var headerRow = -1, weekCols = {};
    for (var r = 0; r < Math.min(rows.length, 10) && headerRow === -1; r++) {
      if (norm((rows[r] || [])[colOffset]) === "ur") {
        headerRow = r;
        var row = rows[r] || [];
        for (var c = colOffset + 1; c < row.length; c++) {
          var m = norm(row[c]).replace(/\s+/g, "").match(/^s(\d{1,2})$/);
          if (m) weekCols[c] = parseInt(m[1], 10);
        }
      }
    }
    if (headerRow === -1) return [];

    var out = [];
    for (var i = headerRow + 1; i < rows.length; i++) {
      var rr = rows[i] || [];
      var ur = rr[colOffset];
      if (!clean(ur)) continue;
      Object.keys(weekCols).forEach(function (colStr) {
        var col = Number(colStr), week = weekCols[colStr];
        var val = rr[col];
        if (val === null || val === undefined || val === "") return;
        var valInt = parseInt(val, 10);
        if (isNaN(valInt)) return;
        out.push({ ur: clean(ur), iso_year: anchorYear, iso_week: week, effectif: valInt });
      });
    }
    return out;
  }

  function parseSuiviWorkbook(arrayBuffer) {
    var wb = XLSX.read(arrayBuffer, { type: "array", cellDates: true });
    var missing = [SHEET_ARRIVALS, SHEET_WEEKLY_INDICATORS, SHEET_FORMATION_IFMIA].filter(function (s) {
      return wb.SheetNames.indexOf(s) === -1;
    });
    if (missing.length) throw new Error("Feuille(s) manquante(s) dans le classeur : " + missing.join(", "));

    var arrivalsRows = XLSX.utils.sheet_to_json(wb.Sheets[SHEET_ARRIVALS], { header: 1, raw: true, defval: null });
    var arrivals = parseArrivals(arrivalsRows);
    if (!arrivals.length) throw new Error("Aucune donnée exploitable dans la feuille '" + SHEET_ARRIVALS + "'");

    var minDate = arrivals[0].date_arrivee_usine;
    arrivals.forEach(function (a) { if (a.date_arrivee_usine < minDate) minDate = a.date_arrivee_usine; });
    var anchorYear = isoWeekInfo(minDate).isoYear;

    var wiRows = XLSX.utils.sheet_to_json(wb.Sheets[SHEET_WEEKLY_INDICATORS], { header: 1, raw: true, defval: null });
    var ifmiaRows = XLSX.utils.sheet_to_json(wb.Sheets[SHEET_FORMATION_IFMIA], { header: 1, raw: true, defval: null });
    var weekly = parseIndicateursHebdomadaires(wiRows, anchorYear);

    return {
      arrivals: arrivals,
      formation_semaine: weekly.formation_semaine,
      formation_ifmia: parseFormationIfmiaTable(ifmiaRows),
      non_diplomes_grid: parseNonDiplomesGrid(ifmiaRows, anchorYear),
      manual_kpis: weekly.manual_kpis,
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

    // "Total du mois" targets the current month once it's mostly over (last
    // week), otherwise the current month barely has any data yet — show the
    // last fully-elapsed month instead. An explicit date_to filter overrides
    // this and always wins.
    var refDate;
    if (dateTo) {
      refDate = dateTo;
    } else {
      var daysInCurMonth = new Date(Date.UTC(today.getUTCFullYear(), today.getUTCMonth() + 1, 0)).getUTCDate();
      var inLastWeekOfMonth = today.getUTCDate() > daysInCurMonth - 7;
      refDate = inLastWeekOfMonth ? today : new Date(Date.UTC(today.getUTCFullYear(), today.getUTCMonth() - 1, 1));
    }
    var monthTotal = 0;
    arrivals.forEach(function (a) {
      if (a.date_arrivee_usine.getUTCFullYear() === refDate.getUTCFullYear() && a.date_arrivee_usine.getUTCMonth() === refDate.getUTCMonth()) {
        monthTotal += a.effectif;
      }
    });
    var MONTH_NAMES_FR = ["Janvier", "Février", "Mars", "Avril", "Mai", "Juin", "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre"];
    var monthLabel = MONTH_NAMES_FR[refDate.getUTCMonth()];

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

    // Diplômés — total headcount from the arrival table (department-scoped),
    // not week-specific: just the running total of everyone tracked.
    var ifmiaDip = parsed.formation_ifmia;
    if (depParam) ifmiaDip = ifmiaDip.filter(function (i) { return i.ur === depParam; });
    var ifmiaDiplomesTotal = null;
    if (ifmiaDip.length) {
      ifmiaDiplomesTotal = 0;
      ifmiaDip.forEach(function (i) { ifmiaDiplomesTotal += i.effectif; });
    }

    // Non-diplômés — sum across departments (or just the selected one) for
    // the current week, from the UR × semaine grid.
    var nonDipRows = parsed.non_diplomes_grid || [];
    if (depParam) nonDipRows = nonDipRows.filter(function (r) { return r.ur === depParam; });
    var ifmiaNonDiplomesThisWeek = null;
    nonDipRows.forEach(function (r) {
      if (r.iso_year === cur.isoYear && r.iso_week === cur.isoWeek) {
        ifmiaNonDiplomesThisWeek = (ifmiaNonDiplomesThisWeek || 0) + r.effectif;
      }
    });

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
      month_label: monthLabel,
      formation: { labels: formationLabels, estimation: formationEstimation, reel: formationReel },
      appels_this_week: appelsThisWeek,
      visite_this_week: visiteThisWeek,
      ifmia_diplomes_total: ifmiaDiplomesTotal,
      ifmia_non_diplomes_this_week: ifmiaNonDiplomesThisWeek,
      upcoming_contracts: { total: upcomingTotal, by_departement: upcomingList },
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

  function renderFormationTable(labels, estimation, reel) {
    var tbody = document.getElementById("us-formation-tbl-body");
    if (!tbody) return;
    if (!labels.length) {
      tbody.innerHTML = '<tr><td colspan="4" style="text-align:center;color:#94a3b8;padding:16px;">Aucune donnée</td></tr>';
      return;
    }
    tbody.innerHTML = labels.map(function (label, i) {
      var est = estimation[i];
      var reelVal = reel[i];
      var known = reelVal !== null && reelVal !== undefined;
      var ecart = known ? reelVal - est : null;
      var ecartCell = known
        ? '<td style="color:' + (ecart > 0 ? "#0a7c55" : ecart < 0 ? "#c2410c" : "#1e293b") + ';font-weight:700;">' + (ecart > 0 ? "+" : "") + ecart + "</td>"
        : '<td style="color:#94a3b8;">—</td>';
      return "<tr><td>" + label + "</td><td>" + fmt(est) + "</td><td>" + (known ? fmt(reelVal) : '<span style="color:#94a3b8;">—</span>') + "</td>" + ecartCell + "</tr>";
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
    document.getElementById("us-kpi-month-label").textContent = d.month_label || "";
    document.getElementById("us-kpi-upcoming").textContent = fmt(d.upcoming_contracts.total);

    setKpiValue("us-kpi-formation", d.ifmia_diplomes_total);
    document.getElementById("us-ifmia-nondip-week-label").textContent = d.this_week_label || "cette semaine";
    setKpiValue("us-kpi-next-week", d.ifmia_non_diplomes_this_week);

    // Entrées usine — delta vs the previous week in the trend.
    var weekSub = document.getElementById("us-kpi-week-sub");
    if (weekSub) {
      var wLabels = d.weekly_arrivals.labels, wValues = d.weekly_arrivals.values;
      var curIdx = wLabels.indexOf(d.this_week_label);
      if (curIdx > 0) {
        var delta = wValues[curIdx] - wValues[curIdx - 1];
        weekSub.textContent = (delta >= 0 ? "+" : "") + delta + " vs " + wLabels[curIdx - 1];
        weekSub.className = "us-kpi-sub " + (delta > 0 ? "us-sub-up" : delta < 0 ? "us-sub-down" : "");
      } else {
        weekSub.textContent = "";
      }
    }


    // Habilitation — top department contributing to upcoming contracts.
    var upcomingSub = document.getElementById("us-kpi-upcoming-sub");
    if (upcomingSub) {
      var top = d.upcoming_contracts.by_departement[0];
      upcomingSub.textContent = top ? ("Top : " + top.label + " (" + fmt(top.effectif) + ")") : "";
    }

    function setKpiValue(elId, value) {
      var el = document.getElementById(elId);
      if (!el) return;
      if (value === null || value === undefined) {
        el.textContent = "Non renseigné";
        el.classList.add("us-empty-val");
      } else {
        el.textContent = fmt(value);
        el.classList.remove("us-empty-val");
      }
    }

    document.getElementById("us-appels-week-label").textContent = d.this_week_label || "cette semaine";
    setKpiValue("us-kpi-appels", d.appels_this_week);
    document.getElementById("us-visite-week-label").textContent = d.this_week_label || "cette semaine";
    setKpiValue("us-kpi-visite", d.visite_this_week);

    renderFormationTable(d.formation.labels, d.formation.estimation, d.formation.reel);

    setTimeout(function () {
      buildWeeklyChart(d.weekly_arrivals.labels, d.weekly_arrivals.values, d.this_week_label);
      buildFormationChart(d.formation.labels, d.formation.estimation, d.formation.reel);
      buildUpcomingChart(d.upcoming_contracts.by_departement);
    }, 50);
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

  var WELCOME_SEEN_KEY = "us_welcome_seen";

  window.usCloseWelcome = function () {
    var overlay = document.getElementById("us-welcome-overlay");
    if (overlay) overlay.classList.remove("us-show");
    try { localStorage.setItem(WELCOME_SEEN_KEY, "1"); } catch (e) {}
  };

  function showWelcomeIfFirstVisit() {
    var seen = false;
    try { seen = localStorage.getItem(WELCOME_SEEN_KEY) === "1"; } catch (e) {}
    if (seen) return;
    var overlay = document.getElementById("us-welcome-overlay");
    if (overlay) overlay.classList.add("us-show");
  }

  document.addEventListener("DOMContentLoaded", function () {
    init();
    showWelcomeIfFirstVisit();
  });
})();
