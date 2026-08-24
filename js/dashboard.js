/* Suivi Usine — fully static dashboard.
 * Reads suivi.xlsx client-side (SheetJS), computes the same KPIs as the
 * server versions, and renders with ECharts. No backend: update the data
 * by replacing suivi.xlsx in this folder and pushing to GitHub. */
(function () {
  "use strict";

  var SHEET_ARRIVALS = "integration Usine";
  var SHEET_FORMATION_IFMIA = "En formation IFMIA";
  var SHEET_NON_DIPLOMES_IFMIA = "Non Diplomes IFMIA";
  var SHEET_WEEKLY_INDICATORS = "Indicateurs Hebdomadaires";
  var SHEET_PROBLEMATIQUES = "Problematiques";
  var SHEET_ARRIVEES_PREVUES = "Arrivees Prevues";
  var SHEET_REPARTITION_ARRIVEE_USINE = "Repartition Arrivee Usine";
  var ATELIER_MAP = { ferrage: "FERRAGE", montage: "MONTAGE", peinture: "PEINTURE" };

  var US_BLUE = "#0B3F91";
  var US_ORANGE = "#FF5A1F";
  var US_CAT_COLORS = ["#0B3F91", "#0D6FA3", "#00B3B8", "#38D6C4", "#6EE7D2", "#9FEDE2", "#FF5A1F", "#4DA8DA"];

  // Reads a sheet into a row-array where index i is ALWAYS actual sheet
  // row i (0-indexed), matching XLSX.utils.encode_cell({r: i, ...}). Plain
  // sheet_to_json indexes from the sheet's used-range start instead (e.g.
  // row 0 of the array is row 4 of the sheet if rows 1-3 are empty), which
  // silently breaks every row/col computed elsewhere in this file — always
  // read sheets through this helper, never call sheet_to_json directly.
  function sheetRows(ws) {
    if (!ws) return [];
    var range = ws["!ref"] ? XLSX.utils.decode_range(ws["!ref"]) : { s: { r: 0, c: 0 }, e: { r: 0, c: 0 } };
    return XLSX.utils.sheet_to_json(ws, {
      header: 1, raw: true, defval: null,
      range: { s: { r: 0, c: 0 }, e: range.e },
    });
  }

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
          depart_ifmia: null,
          vivier:null,
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
      } else if (label.indexOf("depart") !== -1 && label.indexOf("ifmia") !== -1) {
        manualField = "depart_ifmia";

      } else if (label.indexOf("vivier") !== -1) {
  manualField = "vivier";
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

  // Locates the "UR" header cell in column A and the S30..S40 week columns
  // right after it.
  function findUrWeekHeader(rows) {
    for (var r = 0; r < Math.min(rows.length, 10); r++) {
      if (norm((rows[r] || [])[0]) === "ur") {
        var weekCols = {};
        var row = rows[r] || [];
        for (var c = 1; c <= 11; c++) {
          var m = norm(row[c]).replace(/\s+/g, "").match(/^s(\d{1,2})$/);
          if (m) weekCols[c] = parseInt(m[1], 10);
        }
        return { headerRow: r, weekCols: weekCols };
      }
    }
    return { headerRow: -1, weekCols: {} };
  }

  // Generic UR × semaine grid parser — used for both the "En formation
  // IFMIA" (diplômés) and "Non Diplomes IFMIA" sheets, each its own sheet
  // of this exact shape. Each week is its own column so historical weeks
  // are never overwritten.
  function parseUrWeekGrid(rows, anchorYear) {
    var header = findUrWeekHeader(rows);
    var headerRow = header.headerRow, weekCols = header.weekCols;
    if (headerRow === -1) return [];

    var out = [];
    for (var i = headerRow + 1; i < rows.length; i++) {
      var rr = rows[i] || [];
      var ur = rr[0];
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

  function parseFormationIfmiaGrid(rows, anchorYear) {
    return parseUrWeekGrid(rows, anchorYear);
  }

  function parseNonDiplomesGrid(rows, anchorYear) {
    return parseUrWeekGrid(rows, anchorYear);
  }

  // Optional "Problematiques" sheet: an intro line, a numbered list of
  // "Raison N" rows, and two drop-off percentages — key/value rows (col A /
  // col B), entered by hand.
  function parseProblematiques(rows) {
    var result = { intro: "", reasons: [], pct_visite_formation: null, pct_parcours_formation: null };
    for (var i = 1; i < rows.length; i++) {
      var rr = rows[i] || [];
      var label = norm(rr[0]);
      var val = rr[1];
      if (!label) continue;
      if (label.indexOf("raison") !== -1) {
        var reason = clean(val);
        if (reason) result.reasons.push(reason);
      } else if (label.indexOf("intro") !== -1 || label === "description") {
        result.intro = clean(val);
      } else if (label.indexOf("visite") !== -1) {
        var n1 = parseInt(val, 10);
        if (!isNaN(n1)) result.pct_visite_formation = n1;
      } else if (label.indexOf("parcours") !== -1) {
        var n2 = parseInt(val, 10);
        if (!isNaN(n2)) result.pct_parcours_formation = n2;
      }
    }
    return result;
  }

  // Optional "Arrivees Prevues" sheet: a hand-maintained list of upcoming
  // arrivals not yet reflected in the IFMIA/arrivals tables — one row per
  // announcement (Date / Categorie / Valeur).
  function parseArriveesPrevues(rows) {
    var out = [];
    for (var i = 1; i < rows.length; i++) {
      var rr = rows[i] || [];
      var date = parseDate(rr[0]);
      var categorie = clean(rr[1]);
      var valeur = rr[2];
      if (!date || valeur === null || valeur === undefined || valeur === "") continue;
      var valInt = parseInt(valeur, 10);
      if (isNaN(valInt)) continue;
      out.push({ date: date, categorie: categorie, valeur: valInt, row: i });
    }
    return out.sort(function (a, b) { return a.date - b.date; });
  }

  // Optional "Repartition Arrivee Usine" sheet: a hand-maintained snapshot
  // of headcount by department (Departement / Effectif), plant-wide — not
  // scoped by the département/semaine filters, same as "Arrivees Prevues".
  function parseRepartitionArriveeUsine(rows) {
    var out = [];
    for (var i = 1; i < rows.length; i++) {
      var rr = rows[i] || [];
      var dep = clean(rr[0]);
      var eff = rr[1];
      if (!dep || eff === null || eff === undefined || eff === "") continue;
      var effInt = parseInt(eff, 10);
      if (isNaN(effInt)) continue;
      out.push({ departement: dep, effectif: effInt });
    }
    return out;
  }

  function parseSuiviWorkbook(arrayBuffer) {
    return parseSuiviWorkbookData(XLSX.read(arrayBuffer, { type: "array", cellDates: true }));
  }

  function parseSuiviWorkbookData(wb) {
    var missing = [SHEET_ARRIVALS, SHEET_WEEKLY_INDICATORS, SHEET_FORMATION_IFMIA].filter(function (s) {
      return wb.SheetNames.indexOf(s) === -1;
    });
    if (missing.length) throw new Error("Feuille(s) manquante(s) dans le classeur : " + missing.join(", "));

    var arrivalsRows = sheetRows(wb.Sheets[SHEET_ARRIVALS]);
    var arrivals = parseArrivals(arrivalsRows);
    if (!arrivals.length) throw new Error("Aucune donnée exploitable dans la feuille '" + SHEET_ARRIVALS + "'");

    var minDate = arrivals[0].date_arrivee_usine;
    arrivals.forEach(function (a) { if (a.date_arrivee_usine < minDate) minDate = a.date_arrivee_usine; });
    var anchorYear = isoWeekInfo(minDate).isoYear;

    var wiRows = sheetRows(wb.Sheets[SHEET_WEEKLY_INDICATORS]);
    var ifmiaRows = sheetRows(wb.Sheets[SHEET_FORMATION_IFMIA]);
    var weekly = parseIndicateursHebdomadaires(wiRows, anchorYear);

    var problematiques = { intro: "", reasons: [], pct_visite_formation: null, pct_parcours_formation: null };
    if (wb.SheetNames.indexOf(SHEET_PROBLEMATIQUES) !== -1) {
      var probRows = sheetRows(wb.Sheets[SHEET_PROBLEMATIQUES]);
      problematiques = parseProblematiques(probRows);
    }

    var arriveesPrevues = [];
    if (wb.SheetNames.indexOf(SHEET_ARRIVEES_PREVUES) !== -1) {
      var apRows = sheetRows(wb.Sheets[SHEET_ARRIVEES_PREVUES]);
      arriveesPrevues = parseArriveesPrevues(apRows);
    }

    var repartitionArriveeUsine = [];
    if (wb.SheetNames.indexOf(SHEET_REPARTITION_ARRIVEE_USINE) !== -1) {
      var rauRows = sheetRows(wb.Sheets[SHEET_REPARTITION_ARRIVEE_USINE]);
      repartitionArriveeUsine = parseRepartitionArriveeUsine(rauRows);
    }

    var nonDiplomesRows = [];
    if (wb.SheetNames.indexOf(SHEET_NON_DIPLOMES_IFMIA) !== -1) {
      nonDiplomesRows = sheetRows(wb.Sheets[SHEET_NON_DIPLOMES_IFMIA]);
    }

    return {
      arrivals: arrivals,
      anchor_year: anchorYear,
      formation_semaine: weekly.formation_semaine,
      formation_ifmia: parseFormationIfmiaGrid(ifmiaRows, anchorYear),
      non_diplomes_grid: parseNonDiplomesGrid(nonDiplomesRows, anchorYear),
      manual_kpis: weekly.manual_kpis,
      problematiques: problematiques,
      arrivees_prevues: arriveesPrevues,
      repartition_arrivee_usine: repartitionArriveeUsine,
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

    // Department-scoped but NOT narrowed by the Semaine filter — used for
    // the trend charts (weekly arrivals, atelier Réel), which must always
    // show every week regardless of which week is currently selected.
    var arrivalsAllWeeks = parsed.arrivals.filter(function (a) { return scopeCodes.indexOf(a.departement) !== -1; });

    var arrivals = arrivalsAllWeeks;
    if (dateFrom) arrivals = arrivals.filter(function (a) { return a.date_arrivee_usine >= dateFrom; });
    if (dateTo) arrivals = arrivals.filter(function (a) { return a.date_arrivee_usine <= dateTo; });

    var minRelevantDate = null;
    if (parsed.formation_semaine.length) {
      var earliest = parsed.formation_semaine.reduce(function (acc, r) {
        if (!acc) return r;
        return (r.iso_year < acc.iso_year || (r.iso_year === acc.iso_year && r.iso_week < acc.iso_week)) ? r : acc;
      }, null);
      if (earliest) minRelevantDate = isoWeekMonday(earliest.iso_year, earliest.iso_week);
    }

    var weekly = {};
    arrivalsAllWeeks.forEach(function (a) {
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
    // "cette semaine" follows the selected Semaine filter (dateTo is its
    // Sunday) when one is picked — showing that week's real numbers, i.e.
    // the historique — and falls back to today's actual week otherwise.
    var refPoint = dateTo || today;
    var cur = isoWeekInfo(refPoint);
    var thisWeekEntry = weeklySorted.filter(function (w) { return w.year === cur.isoYear && w.week === cur.isoWeek; })[0];
    var thisWeekTotal = thisWeekEntry ? thisWeekEntry.value : 0;

    // Réel per atelier/semaine is calculated live from the "integration
    // Usine" arrivals (matched by département FER/MON/PEI) — never typed
    // by hand. Estimation stays the manually-entered plan from Indicateurs
    // Hebdomadaires. Uses arrivalsAllWeeks so the Task Force table and the
    // Formation chart always show every week, unaffected by the Semaine filter.
    var ATELIER_DEPT = { FERRAGE: "FER", MONTAGE: "MON", PEINTURE: "PEI" };
    var arrivalsByAtelierWeek = {};
    arrivalsAllWeeks.forEach(function (a) {
      var atelier = null;
      Object.keys(ATELIER_DEPT).forEach(function (k) { if (ATELIER_DEPT[k] === a.departement) atelier = k; });
      if (!atelier) return;
      var info = isoWeekInfo(a.date_arrivee_usine);
      var key = atelier + "_" + info.isoYear + "_" + info.isoWeek;
      arrivalsByAtelierWeek[key] = (arrivalsByAtelierWeek[key] || 0) + a.effectif;
    });

    var atelierByWeek = {}; // atelier -> [{iso_year, iso_week, estimation, reel}]
    var fsByWeek = {};
    parsed.formation_semaine.forEach(function (fs) {
      var reelVal = arrivalsByAtelierWeek[fs.atelier + "_" + fs.iso_year + "_" + fs.iso_week] || 0;

      if (!atelierByWeek[fs.atelier]) atelierByWeek[fs.atelier] = [];
      atelierByWeek[fs.atelier].push({ iso_year: fs.iso_year, iso_week: fs.iso_week, estimation: fs.estimation || 0, reel: reelVal });

      var key = fs.iso_year + "_" + fs.iso_week;
      if (!fsByWeek[key]) fsByWeek[key] = { year: fs.iso_year, week: fs.iso_week, estimation: 0, reel: 0 };
      fsByWeek[key].estimation += fs.estimation || 0;
      fsByWeek[key].reel += reelVal;
    });
    Object.keys(atelierByWeek).forEach(function (k) {
      atelierByWeek[k].sort(function (x, y) { return x.iso_year - y.iso_year || x.iso_week - y.iso_week; });
    });

    var fsSorted = Object.keys(fsByWeek).map(function (k) { return fsByWeek[k]; })
      .sort(function (x, y) { return x.year - y.year || x.week - y.week; });
    var formationLabels = fsSorted.map(function (w) { return "S" + w.week; });
    var formationEstimation = fsSorted.map(function (w) { return w.estimation; });
    var formationReel = fsSorted.map(function (w) { return w.reel; });

    // Habilitation — contrats à venir. Always computed from the full
    // arrivals table (department-scoped only), never from the
    // Semaine-filtered subset: it answers "who has a contract starting
    // soon", which has nothing to do with which arrival week is browsed.
    var upcomingByDept = {};
    parsed.arrivals.forEach(function (a) {
      if (scopeCodes.indexOf(a.departement) === -1) return;
      if (a.date_debut_contrat && a.date_debut_contrat > today) {
        upcomingByDept[a.departement] = (upcomingByDept[a.departement] || 0) + a.effectif;
      }
    });
    var upcomingTotal = 0;
    Object.keys(upcomingByDept).forEach(function (k) { upcomingTotal += upcomingByDept[k]; });
    var upcomingList = Object.keys(upcomingByDept).map(function (d) { return { departement: d, label: d, effectif: upcomingByDept[d] }; })
      .sort(function (a, b) { return b.effectif - a.effectif; });

    // Répartition par département du 1er juillet à aujourd'hui — a fixed
    // reporting window (not affected by the Semaine filter), calculated
    // live from the "integration Usine" arrivals, department-scoped.
    var julyStart = new Date(Date.UTC(today.getUTCFullYear(), 6, 1));
    var julyDeptSums = {};
    parsed.arrivals.forEach(function (a) {
      if (scopeCodes.indexOf(a.departement) === -1) return;
      if (a.date_arrivee_usine < julyStart || a.date_arrivee_usine > today) return;
      julyDeptSums[a.departement] = (julyDeptSums[a.departement] || 0) + a.effectif;
    });
    var julyDeptList = Object.keys(julyDeptSums).sort().map(function (d) { return { departement: d, effectif: julyDeptSums[d] }; });
    var julyDeptTotal = julyDeptList.reduce(function (s, r) { return s + r.effectif; }, 0);

    // Diplômés — from the UR × semaine grid, scoped to the currently
    // viewed week (same historique pattern as the other weekly cards):
    // each week is its own column, so past weeks are never overwritten.
    var ifmiaDipRows = parsed.formation_ifmia || [];
    if (depParam) ifmiaDipRows = ifmiaDipRows.filter(function (i) { return i.ur === depParam; });
    var ifmiaDipThisWeek = ifmiaDipRows.filter(function (r) { return r.iso_year === cur.isoYear && r.iso_week === cur.isoWeek; });
    var ifmiaFerrageTotal = null, ifmiaMontageTotal = null, ifmiaBirdCageTotal = null;
    var ferHas = false, monHas = false, bcHas = false;
    ifmiaDipThisWeek.forEach(function (i) {
      if (i.ur === "FER") { ifmiaFerrageTotal = (ifmiaFerrageTotal || 0) + i.effectif; ferHas = true; }
      else if (i.ur === "MON") { ifmiaMontageTotal = (ifmiaMontageTotal || 0) + i.effectif; monHas = true; }
      else if (i.ur === "BC") { ifmiaBirdCageTotal = (ifmiaBirdCageTotal || 0) + i.effectif; bcHas = true; }
    });
    if (!ferHas) ifmiaFerrageTotal = null;
    if (!monHas) ifmiaMontageTotal = null;
    if (!bcHas) ifmiaBirdCageTotal = null;

    // Total à IFMIA is its own manually-entered "TOTAL" row (not the sum
    // of Ferrage + Montage + other lots) — plant-wide, so not filtered by
    // département, same as Appels/Visite/Départ IFMIA.
    var ifmiaDiplomesTotal = null;
    (parsed.formation_ifmia || []).forEach(function (r) {
      if (r.ur === "TOTAL" && r.iso_year === cur.isoYear && r.iso_week === cur.isoWeek) {
        ifmiaDiplomesTotal = r.effectif;
      }
    });

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
    var departIfmiaThisWeek = manualEntry && manualEntry.depart_ifmia !== null ? manualEntry.depart_ifmia : null;
    var vivierThisWeek = manualEntry && manualEntry.vivier !== null ? manualEntry.vivier : null;

    // Arrivées prévues — a hand-maintained list of upcoming arrivals (see
    // the "Arrivees Prevues" sheet). Always shown regardless of the
    // département/semaine filters — only expires against the real
    // "today", never against whatever filter is currently selected.
    var upcomingIntegrations = (parsed.arrivees_prevues || [])
      .filter(function (r) { return r.date >= today; })
      .map(function (r) { return { date_label: fmtDateFr(r.date), categorie: r.categorie, value: r.valeur, row: r.row }; });

    // Répartition arrivée usine — hand-maintained snapshot, plant-wide
    // (same as arrivées prévues: not filtered by département/semaine).
    var repartitionArriveeUsineRows = parsed.repartition_arrivee_usine || [];
    var repartitionArriveeUsineTotal = repartitionArriveeUsineRows.reduce(function (s, r) { return s + r.effectif; }, 0);

    return {
      empty: false,
      departments: presentCodes.map(function (c) { return { code: c, label: c }; }),
      selected_departement: depParam,
      weekly_arrivals: { labels: weeklyLabels, values: weeklyValues },
      this_week_total: thisWeekTotal,
      this_week_label: "S" + cur.isoWeek,
      formation: { labels: formationLabels, estimation: formationEstimation, reel: formationReel },
      appels_this_week: appelsThisWeek,
      vivier_this_week: vivierThisWeek,
      visite_this_week: visiteThisWeek,
      depart_ifmia_this_week: departIfmiaThisWeek,
      ifmia_diplomes_total: ifmiaDiplomesTotal,
      ifmia_ferrage_total: ifmiaFerrageTotal,
      ifmia_montage_total: ifmiaMontageTotal,
      ifmia_birdcage_total: ifmiaBirdCageTotal,
      upcoming_integrations: upcomingIntegrations,
      ifmia_non_diplomes_this_week: ifmiaNonDiplomesThisWeek,
      upcoming_contracts: { total: upcomingTotal, by_departement: upcomingList },
      problematiques: parsed.problematiques,
      atelier_formation: atelierByWeek,
      july_to_today: { rows: julyDeptList, total: julyDeptTotal, from: julyStart, to: today },
      repartition_arrivee_usine: { rows: repartitionArriveeUsineRows, total: repartitionArriveeUsineTotal },
    };
  }

  // ── rendering ────────────────────────────────────────────────────────
  var US_CHARTS = {};
  var parsedWorkbook = null;
  var rawWorkbook = null;
  var lastRenderedWeekLabel = null;
  var hasPendingEdits = false;
  var deptsLoaded = false;
  var weeksLoaded = false;

  function fmt(n) { return (n || 0).toLocaleString("fr-FR"); }

  var CATEGORIE_LABELS = { "diplomes": "Diplômés IFMIA", "non diplomes": "Non diplômés" };
  function categorieLabel(cat) {
    return CATEGORIE_LABELS[norm(cat).replace(/\s+/g, " ").trim()] || cat || "";
  }

  // Builds a single-quoted JS string literal safe to embed inside a
  // double-quoted HTML attribute (e.g. onclick="fn('...')").
  function jsStr(s) {
    return "'" + String(s).replace(/\\/g, "\\\\").replace(/'/g, "\\'") + "'";
  }

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

    chart.setOption({
      tooltip: { trigger: "axis" },
      legend: { data: ["Estimation", "Réel"], top: 0, textStyle: { fontSize: 12 } },
      grid: { left: "3%", right: "4%", bottom: "3%", top: "18%", containLabel: true },
      xAxis: { type: "category", data: labels, axisLabel: { fontSize: 11 } },
      yAxis: { type: "value", axisLabel: { fontSize: 11 } },
      series: [
        { name: "Réel", type: "bar", data: reel, barMaxWidth: 28, itemStyle: { color: US_BLUE, borderRadius: [4, 4, 0, 0] } },
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

  var ATELIER_ORDER = [
    { key: "FERRAGE", label: "Ferrage" },
    { key: "MONTAGE", label: "Montage" },
    { key: "PEINTURE", label: "Peinture" }
  ];

  function renderAtelierTable(atelierByWeek) {
    var thead = document.getElementById("us-atelier-tbl-head");
    var tbody = document.getElementById("us-atelier-tbl-body");
    if (!thead || !tbody) return;

    var weeksRef = null;
    ATELIER_ORDER.forEach(function (o) {
      if (!weeksRef && atelierByWeek[o.key] && atelierByWeek[o.key].length) weeksRef = atelierByWeek[o.key];
    });
    if (!weeksRef) {
      thead.innerHTML = "";
      tbody.innerHTML = '<tr><td style="text-align:center;color:#94a3b8;padding:16px;">Aucune donnée</td></tr>';
      return;
    }

    var weekLabels = weeksRef.map(function (w) { return "S" + w.iso_week; });
    thead.innerHTML = "<tr><th>Atelier</th><th>Type</th>" + weekLabels.map(function (l) { return "<th>" + l + "</th>"; }).join("") + "</tr>";

    var rowsHtml = "";
    ATELIER_ORDER.forEach(function (o) {
      var rows = atelierByWeek[o.key] || [];
      var estCells = rows.map(function (r) { return "<td>" + fmt(r.estimation) + "</td>"; }).join("");
      var reelCells = rows.map(function (r) { return "<td>" + fmt(r.reel) + "</td>"; }).join("");
      var cumul = 0;
      var ecartCells = rows.map(function (r) {
        var e = r.reel - r.estimation;
        var color = e > 0 ? "#0a7c55" : e < 0 ? "#c2410c" : "#1e293b";
        return '<td style="color:' + color + ';font-weight:700;">' + (e > 0 ? "+" : "") + e + "</td>";
      }).join("");
      var ecartCumulCells = rows.map(function (r) {
        cumul += r.reel - r.estimation;
        var color = cumul > 0 ? "#0a7c55" : cumul < 0 ? "#c2410c" : "#1e293b";
        return '<td style="color:' + color + ';font-weight:700;">' + (cumul > 0 ? "+" : "") + cumul + "</td>";
      }).join("");
      rowsHtml += "<tr><td rowspan=\"4\" style=\"font-weight:700;\">" + o.label + "</td><td>Estimation</td>" + estCells + "</tr>";
      rowsHtml += "<tr><td>Réel</td>" + reelCells + "</tr>";
      rowsHtml += "<tr><td>Écart</td>" + ecartCells + "</tr>";
      rowsHtml += "<tr><td>Écart cumulé</td>" + ecartCumulCells + "</tr>";
    });
    tbody.innerHTML = rowsHtml;
  }

  function fmtDateFr(d) {
    var day = String(d.getUTCDate()).padStart(2, "0");
    var month = String(d.getUTCMonth() + 1).padStart(2, "0");
    return day + "/" + month;
  }

  function renderJulyTable(julyData) {
    var title = document.getElementById("us-july-title");
    var tbody = document.getElementById("us-july-tbl-body");
    if (!tbody) return;
    if (title) title.textContent = "Répartition par département — du " + fmtDateFr(julyData.from) + " au " + fmtDateFr(julyData.to);
    if (!julyData.rows.length) {
      tbody.innerHTML = '<tr><td colspan="2" style="text-align:center;color:#94a3b8;padding:16px;">Aucune donnée</td></tr>';
      return;
    }
    var rowsHtml = julyData.rows.map(function (r) {
      return "<tr><td>" + r.departement + "</td><td>" + fmt(r.effectif) + "</td></tr>";
    }).join("");
    rowsHtml += '<tr style="font-weight:800;background:#f8fafc;"><td>Total général</td><td>' + fmt(julyData.total) + "</td></tr>";
    tbody.innerHTML = rowsHtml;
  }

  function renderRepartitionArriveeUsineTable(data) {
    var card = document.getElementById("us-rau-card");
    var tbody = document.getElementById("us-rau-tbl-body");
    if (!card || !tbody) return;
    if (!data.rows.length) { card.style.display = "none"; return; }
    card.style.display = "block";
    var rowsHtml = data.rows.map(function (r) {
      return "<tr><td>" + r.departement + "</td><td>" + fmt(r.effectif) + "</td></tr>";
    }).join("");
    rowsHtml += '<tr style="font-weight:800;background:#f8fafc;"><td>Total général</td><td>' + fmt(data.total) + "</td></tr>";
    tbody.innerHTML = rowsHtml;
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

  function populateWeeks(weekLabels) {
    var sel = document.getElementById("us-week-select");
    if (!sel || weeksLoaded || !weekLabels || !weekLabels.length) return;
    weekLabels.forEach(function (label) {
      var opt = document.createElement("option");
      opt.value = label; opt.textContent = label;
      sel.appendChild(opt);
    });
    weeksLoaded = true;
  }

  function render(d) {
    document.getElementById("us-content").style.display = "block";
    document.getElementById("us-empty").style.display = "none";

    populateDepartments(d.departments, d.selected_departement);
    populateWeeks(d.formation.labels);

    lastRenderedWeekLabel = d.this_week_label || null;
    document.getElementById("us-this-week-label").textContent = d.this_week_label || "cette semaine";
    document.getElementById("us-kpi-week").textContent = fmt(d.this_week_total);
    document.getElementById("us-kpi-month").textContent = fmt(d.july_to_today.total);
    document.getElementById("us-kpi-upcoming").textContent = fmt(d.upcoming_contracts.total);

    document.getElementById("us-formation-total-week-label").textContent = d.this_week_label || "cette semaine";
    setKpiValue("us-kpi-formation", d.ifmia_diplomes_total);
    document.getElementById("us-formation-ferrage-week-label").textContent = d.this_week_label || "cette semaine";
    setDeptKpiValue("us-kpi-formation-ferrage", d.ifmia_ferrage_total);
    document.getElementById("us-formation-montage-week-label").textContent = d.this_week_label || "cette semaine";
    setDeptKpiValue("us-kpi-formation-montage", d.ifmia_montage_total);
    document.getElementById("us-formation-birdcage-week-label").textContent = d.this_week_label || "cette semaine";
    setDeptKpiValue("us-kpi-formation-birdcage", d.ifmia_birdcage_total);
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

    // Per-département formation-IFMIA KPIs (Ferrage/Montage/Bird Cage) are
    // hidden entirely when the value is 0 — a card showing "0" for a
    // département that just isn't training anyone this week is noise.
    function setDeptKpiValue(elId, value) {
      var el = document.getElementById(elId);
      if (!el) return;
      var card = el.closest(".us-kpi-card");
      if (value === 0) {
        if (card) card.style.display = "none";
        return;
      }
      if (card) card.style.display = "";
      setKpiValue(elId, value);
    }
    document.getElementById("us-vivier-week-label").textContent = d.this_week_label || "cette semaine";
    setKpiValue("us-kpi-vivier", d.vivier_this_week);

    document.getElementById("us-appels-week-label").textContent = d.this_week_label || "cette semaine";
    setKpiValue("us-kpi-appels", d.appels_this_week);
    document.getElementById("us-visite-week-label").textContent = d.this_week_label || "cette semaine";
    setKpiValue("us-kpi-visite", d.visite_this_week);
    document.getElementById("us-depart-ifmia-week-label").textContent = d.this_week_label || "cette semaine";
    setKpiValue("us-kpi-depart-ifmia", d.depart_ifmia_this_week);

    renderAtelierTable(d.atelier_formation);
    renderJulyTable(d.july_to_today);
    renderRepartitionArriveeUsineTable(d.repartition_arrivee_usine);

    var announceCard = document.getElementById("us-announce-card");
    var announceList = document.getElementById("us-announce-list");
    if (announceCard && announceList) {
      var integrations = d.upcoming_integrations || [];
      announceCard.style.display = integrations.length ? "block" : "none";
      announceList.innerHTML = integrations.map(function (r) {
        var catLabel = categorieLabel(r.categorie);
        return '<div class="us-announce-item">' +
          '<span>Arrivée prévue le <strong>' + r.date_label + "</strong>" + (catLabel ? " — " + catLabel : "") + " : <strong>" + fmt(r.value) + "</strong></span>" +
          '<button type="button" class="us-announce-edit-btn" onclick="usOpenAnnounceEditModal(' + r.row + "," + jsStr(r.date_label) + "," + jsStr(catLabel) + "," + r.value + ')" title="Modifier cette valeur">' +
          '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>' +
          "</button></div>";
      }).join("");
    }

    var prob = d.problematiques;
    var insightCard = document.getElementById("us-insight-card");
    if (insightCard) {
      var hasContent = prob && (prob.intro || prob.reasons.length || prob.pct_visite_formation !== null || prob.pct_parcours_formation !== null);
      insightCard.style.display = hasContent ? "block" : "none";
      if (hasContent) {
        document.getElementById("us-insight-intro").textContent = prob.intro || "";
        var reasonsEl = document.getElementById("us-insight-reasons");
        reasonsEl.innerHTML = prob.reasons.map(function (r) { return "<li>" + r + "</li>"; }).join("");
        document.getElementById("us-insight-visite").textContent = prob.pct_visite_formation !== null ? prob.pct_visite_formation + "%" : "Non renseigné";
        document.getElementById("us-insight-parcours").textContent = prob.pct_parcours_formation !== null ? prob.pct_parcours_formation + "%" : "Non renseigné";
      }
    }

    setTimeout(function () {
      buildWeeklyChart(d.weekly_arrivals.labels, d.weekly_arrivals.values, d.this_week_label);
      buildFormationChart(d.formation.labels, d.formation.estimation, d.formation.reel);
      buildUpcomingChart(d.upcoming_contracts.by_departement);
    }, 50);
  }

  // ── Inline KPI editing ──────────────────────────────────────────────
  // Edits write directly into the in-memory workbook (rawWorkbook) and the
  // dashboard re-renders immediately. Nothing is persisted anywhere until
  // the user clicks "Télécharger" and re-uploads the file — there is no
  // backend and no automatic write-back to the repo.
  var EDIT_LABELS = {
    appels: "Appels téléphoniques",
    visite: "Visite médicale",
    "non-diplomes": "En formation (non diplômés)",
    "formation-ferrage": "En formation à IFMIA — Ferrage",
    "formation-montage": "En formation à IFMIA — Montage",
    "formation-birdcage": "En formation à IFMIA — Bird Cage",
    "formation-total": "En formation total à IFMIA",
    "depart-ifmia": "Départ IFMIA",
  };
  var WEEK_SCOPED_EDIT = { appels: true, visite: true, "non-diplomes": true, "depart-ifmia": true, "formation-ferrage": true, "formation-montage": true, "formation-birdcage": true, "formation-total": true };

  function cellAddr(r, c) { return XLSX.utils.encode_cell({ r: r, c: c }); }

  function setCell(ws, r, c, value, type) {
    ws[cellAddr(r, c)] = { t: type, v: value };
    var range = ws["!ref"] ? XLSX.utils.decode_range(ws["!ref"]) : { s: { r: r, c: c }, e: { r: r, c: c } };
    range.s.r = Math.min(range.s.r, r); range.s.c = Math.min(range.s.c, c);
    range.e.r = Math.max(range.e.r, r); range.e.c = Math.max(range.e.c, c);
    ws["!ref"] = XLSX.utils.encode_range(range);
  }

  function resolveEditTarget(kpiKey, weekLabel, dept) {
    var weekNum = parseInt((weekLabel || "").replace(/^S/i, ""), 10);

    if (kpiKey === "appels" || kpiKey === "visite" || kpiKey === "depart-ifmia") {
      var ws = rawWorkbook.Sheets[SHEET_WEEKLY_INDICATORS];
      var rows = sheetRows(ws);
      var weekRow = -1, weekCols = {};
      for (var r = 0; r < Math.min(rows.length, 5) && weekRow === -1; r++) {
        var row = rows[r] || [], cols = {};
        for (var c = 0; c < row.length; c++) {
          var m = norm(row[c]).replace(/\s+/g, "").match(/^s(\d{1,2})$/);
          if (m) cols[c] = parseInt(m[1], 10);
        }
        if (Object.keys(cols).length) { weekRow = r; weekCols = cols; }
      }
      if (weekRow === -1) return { ok: false, reason: "Grille des semaines introuvable dans 'Indicateurs Hebdomadaires'." };
      var col = -1;
      Object.keys(weekCols).forEach(function (c) { if (weekCols[c] === weekNum) col = Number(c); });
      if (col === -1) return { ok: false, reason: "La semaine " + weekLabel + " n'existe pas dans le fichier." };

      var targetRow = -1;
      for (var i = weekRow + 1; i < rows.length; i++) {
        var label = norm((rows[i] || [])[0]);
        if (kpiKey === "appels" && label.indexOf("appel") !== -1) { targetRow = i; break; }
        if (kpiKey === "visite" && label.indexOf("visite") !== -1 && label.indexOf("medic") !== -1) { targetRow = i; break; }
        if (kpiKey === "depart-ifmia" && label.indexOf("depart") !== -1 && label.indexOf("ifmia") !== -1) { targetRow = i; break; }
      }
      if (targetRow === -1) return { ok: false, reason: "Ligne introuvable dans 'Indicateurs Hebdomadaires'." };
      return { ok: true, sheet: SHEET_WEEKLY_INDICATORS, row: targetRow, col: col };
    }

    if (kpiKey === "non-diplomes") {
      var ws2 = rawWorkbook.Sheets[SHEET_NON_DIPLOMES_IFMIA];
      var rows2 = sheetRows(ws2);
      var colOffset = 0; // column A, 0-indexed — its own sheet since the split
      var header2 = findUrWeekHeader(rows2);
      var headerRow = header2.headerRow, weekCols2 = header2.weekCols;
      if (headerRow === -1) return { ok: false, reason: "Grille non-diplômés introuvable." };
      var col2 = -1;
      Object.keys(weekCols2).forEach(function (c) { if (weekCols2[c] === weekNum) col2 = Number(c); });
      if (col2 === -1) return { ok: false, reason: "La semaine " + weekLabel + " n'existe pas dans la grille non-diplômés." };

      var urTarget = (dept || "TOTAL").toUpperCase();
      var targetRow2 = -1;
      var lastRow = headerRow;
      for (var i2 = headerRow + 1; i2 < rows2.length; i2++) {
        var urVal = clean((rows2[i2] || [])[colOffset]);
        if (urVal) lastRow = i2;
        if (urVal.toUpperCase() === urTarget) { targetRow2 = i2; break; }
      }
      if (targetRow2 === -1) {
        targetRow2 = lastRow + 1;
        return { ok: true, sheet: SHEET_NON_DIPLOMES_IFMIA, row: targetRow2, col: col2, ensureLabel: { col: colOffset, value: urTarget } };
      }
      return { ok: true, sheet: SHEET_NON_DIPLOMES_IFMIA, row: targetRow2, col: col2 };
    }

    if (kpiKey === "formation-ferrage" || kpiKey === "formation-montage" || kpiKey === "formation-birdcage" || kpiKey === "formation-total") {
      var urCode = kpiKey === "formation-ferrage" ? "FER" : kpiKey === "formation-montage" ? "MON" : kpiKey === "formation-birdcage" ? "BC" : "TOTAL";
      var ws3 = rawWorkbook.Sheets[SHEET_FORMATION_IFMIA];
      var rows3 = sheetRows(ws3);
      var dipColOffset = 0; // column A, 0-indexed — see parseFormationIfmiaGrid
      var header3 = findUrWeekHeader(rows3);
      var headerRow3 = header3.headerRow, weekCols3 = header3.weekCols;
      if (headerRow3 === -1) return { ok: false, reason: "Tableau des diplômés introuvable." };
      var col3 = -1;
      Object.keys(weekCols3).forEach(function (c) { if (weekCols3[c] === weekNum) col3 = Number(c); });
      if (col3 === -1) return { ok: false, reason: "La semaine " + weekLabel + " n'existe pas dans le tableau des diplômés." };

      var targetRow3 = -1;
      var lastRow3 = headerRow3;
      for (var i3 = headerRow3 + 1; i3 < rows3.length; i3++) {
        var ur3 = clean((rows3[i3] || [])[dipColOffset]);
        if (ur3) lastRow3 = i3;
        if (ur3.toUpperCase() === urCode) { targetRow3 = i3; break; }
      }
      if (targetRow3 === -1) {
        targetRow3 = lastRow3 + 1;
        return { ok: true, sheet: SHEET_FORMATION_IFMIA, row: targetRow3, col: col3, ensureLabel: { col: dipColOffset, value: urCode } };
      }
      return { ok: true, sheet: SHEET_FORMATION_IFMIA, row: targetRow3, col: col3 };
    }

    return { ok: false, reason: "Champ non modifiable." };
  }

  function applyKpiEdit(kpiKey, weekLabel, dept, newValue) {
    var target = resolveEditTarget(kpiKey, weekLabel, dept);
    if (!target.ok) return target;
    var ws = rawWorkbook.Sheets[target.sheet];
    if (target.ensureLabel) setCell(ws, target.row, target.ensureLabel.col, target.ensureLabel.value, "s");
    setCell(ws, target.row, target.col, newValue, "n");

    parsedWorkbook = parseSuiviWorkbookData(rawWorkbook);
    hasPendingEdits = true;
    updateEditBar();
    refresh();
    return { ok: true };
  }

  function updateEditBar() {
    var bar = document.getElementById("us-edit-bar");
    if (bar) bar.style.display = hasPendingEdits ? "flex" : "none";
  }

  function applyAnnounceEdit(row, newValue) {
    var ws = rawWorkbook.Sheets[SHEET_ARRIVEES_PREVUES];
    if (!ws) return { ok: false, reason: "Feuille '" + SHEET_ARRIVEES_PREVUES + "' introuvable." };
    setCell(ws, row, 2, newValue, "n");
    parsedWorkbook = parseSuiviWorkbookData(rawWorkbook);
    hasPendingEdits = true;
    updateEditBar();
    refresh();
    return { ok: true };
  }

  var currentEditKey = null;
  var currentAnnounceRow = null;

  window.usOpenEditModal = function (kpiKey) {
    if (!rawWorkbook) return;
    currentEditKey = kpiKey;
    currentAnnounceRow = null;
    var weekLabel = lastRenderedWeekLabel || "S" + isoWeekInfo(todayUTC()).isoWeek;
    var dept = (document.getElementById("us-dept-select") || {}).value || "";
    var weekScoped = !!WEEK_SCOPED_EDIT[kpiKey];

    document.getElementById("us-edit-modal-title").textContent = EDIT_LABELS[kpiKey] || "Modifier";
    var weekEl = document.getElementById("us-edit-modal-week");
    weekEl.style.display = weekScoped ? "" : "none";
    weekEl.textContent = weekScoped ? "Semaine " + weekLabel : "";

    var target = resolveEditTarget(kpiKey, weekLabel, dept);
    var errEl = document.getElementById("us-edit-modal-error");
    var input = document.getElementById("us-edit-modal-input");
    if (!target.ok) {
      errEl.style.display = "block";
      errEl.textContent = target.reason;
      input.style.display = "none";
    } else {
      errEl.style.display = "none";
      input.style.display = "";
      var ws = rawWorkbook.Sheets[target.sheet];
      var cell = ws[cellAddr(target.row, target.col)];
      input.value = cell && typeof cell.v === "number" ? cell.v : "";
    }
    document.getElementById("us-edit-modal-overlay").style.display = "flex";
    if (target.ok) input.focus();
  };

  window.usOpenAnnounceEditModal = function (row, dateLabel, categorieText, currentValue) {
    if (!rawWorkbook) return;
    currentEditKey = null;
    currentAnnounceRow = row;

    document.getElementById("us-edit-modal-title").textContent = "Arrivée prévue — " + categorieText;
    var weekEl = document.getElementById("us-edit-modal-week");
    weekEl.style.display = "";
    weekEl.textContent = "Le " + dateLabel;

    document.getElementById("us-edit-modal-error").style.display = "none";
    var input = document.getElementById("us-edit-modal-input");
    input.style.display = "";
    input.value = currentValue;
    document.getElementById("us-edit-modal-overlay").style.display = "flex";
    input.focus();
  };

  window.usCloseEditModal = function () {
    document.getElementById("us-edit-modal-overlay").style.display = "none";
    currentEditKey = null;
    currentAnnounceRow = null;
  };

  window.usSaveEditModal = function () {
    var input = document.getElementById("us-edit-modal-input");
    var val = parseInt(input.value, 10);
    var errEl = document.getElementById("us-edit-modal-error");
    if (isNaN(val) || val < 0) {
      errEl.style.display = "block";
      errEl.textContent = "Merci d'entrer un nombre valide (≥ 0).";
      return;
    }
    var result;
    if (currentAnnounceRow !== null) {
      result = applyAnnounceEdit(currentAnnounceRow, val);
    } else {
      var weekLabel = lastRenderedWeekLabel || "S" + isoWeekInfo(todayUTC()).isoWeek;
      var dept = (document.getElementById("us-dept-select") || {}).value || "";
      result = applyKpiEdit(currentEditKey, weekLabel, dept, val);
    }
    if (!result.ok) {
      errEl.style.display = "block";
      errEl.textContent = result.reason;
      return;
    }
    window.usCloseEditModal();
  };

  window.usDownloadUpdatedWorkbook = function () {
    if (!rawWorkbook) return;
    var out = XLSX.write(rawWorkbook, { bookType: "xlsx", type: "array" });
    var blob = new Blob([out], { type: "application/octet-stream" });
    var url = URL.createObjectURL(blob);
    var a = document.createElement("a");
    a.href = url; a.download = "suivi.xlsx";
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  function refresh() {
    if (!parsedWorkbook) return;
    var errEl = document.getElementById("us-error");
    if (errEl) errEl.style.display = "none";
    try {
      var filters = { departement: document.getElementById("us-dept-select").value };
      var weekVal = document.getElementById("us-week-select").value; // e.g. "S32"
      if (weekVal && parsedWorkbook.anchor_year) {
        var weekNum = parseInt(weekVal.replace(/^S/i, ""), 10);
        if (!isNaN(weekNum)) {
          var monday = isoWeekMonday(parsedWorkbook.anchor_year, weekNum);
          var sunday = new Date(monday);
          sunday.setUTCDate(sunday.getUTCDate() + 6);
          filters.date_from = monday.toISOString().slice(0, 10);
          filters.date_to = sunday.toISOString().slice(0, 10);
        }
      }
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
    document.getElementById("us-week-select").value = "";
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
        rawWorkbook = XLSX.read(buf, { type: "array", cellDates: true });
        parsedWorkbook = parseSuiviWorkbookData(rawWorkbook);
        refresh();
      })
      .catch(function (e) {
        if (errEl) { errEl.style.display = "block"; errEl.textContent = "Erreur de chargement : " + e.message; }
      });

    document.getElementById("us-dept-select").addEventListener("change", refresh);
    document.getElementById("us-week-select").addEventListener("change", refresh);
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

  // ── Access gate — a light client-side deterrent only. This is NOT real
  // security: the hash and comparison run in the browser and are visible
  // via view-source, and suivi.xlsx remains directly fetchable at its own
  // URL regardless of this screen. It just keeps casual/accidental
  // visitors from landing straight on the dashboard.
  var GATE_UNLOCK_KEY = "us_gate_unlocked";
  var GATE_HASH = "97e195bd33d5466dbdfd768218ac94c80d869b3c1630640edfdda3454dc76d72"; // sha256("AOUT2026")

  function sha256Hex(text) {
    var enc = new TextEncoder().encode(text);
    return crypto.subtle.digest("SHA-256", enc).then(function (buf) {
      return Array.from(new Uint8Array(buf)).map(function (b) { return b.toString(16).padStart(2, "0"); }).join("");
    });
  }

  function unlockApp() {
    document.body.classList.remove("us-locked");
    document.body.classList.add("us-unlocked");
    var overlay = document.getElementById("us-gate-overlay");
    if (overlay) overlay.style.display = "none";
    init();
    showWelcomeIfFirstVisit();
  }

  function initGate() {
    var unlocked = false;
    try { unlocked = sessionStorage.getItem(GATE_UNLOCK_KEY) === "1"; } catch (e) {}
    if (unlocked) { unlockApp(); return; }

    var form = document.getElementById("us-gate-form");
    if (!form) { unlockApp(); return; } // no gate markup on the page — fail open

    form.addEventListener("submit", function (e) {
      e.preventDefault();
      var input = document.getElementById("us-gate-input");
      var errorEl = document.getElementById("us-gate-error");
      var value = (input.value || "").trim().toUpperCase();
      sha256Hex(value).then(function (hash) {
        if (hash === GATE_HASH) {
          try { sessionStorage.setItem(GATE_UNLOCK_KEY, "1"); } catch (e) {}
          unlockApp();
        } else {
          if (errorEl) errorEl.style.display = "block";
          input.value = "";
          input.focus();
        }
      });
    });
  }

  document.addEventListener("DOMContentLoaded", initGate);
})();
