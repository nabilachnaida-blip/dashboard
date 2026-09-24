(function () {
  /* Pour ajouter une journée : ajouter une valeur à la fin de chacun des
     trois tableaux. Les libellés, les indicateurs du bandeau et le nombre
     de journées affiché se recalculent tout seuls. */
  var CONVOQUES = [230, 203, 207, 177, 184, 184, 180, 179, 227, 124, 240,169,116,184,192,161,230,100,178, 207,161,202,186,102,158,157,115,115];
  var PRESENTS = [111, 97, 109, 83, 90, 90, 75, 99, 107, 49,103,65,56,106,98,81,100,58,88,118,88,100,97,63,98,119,70,70];
  var RETENUS = [109, 95, 106, 78, 83, 83, 70, 86, 102, 48, 95,63,52,101,95,77,100,55,84,118,84,98,95,59,98,115,70,70];

  var LABELS = CONVOQUES.map(function (_, i) { return "Journée " + (i + 1); });

  if (PRESENTS.length !== CONVOQUES.length || RETENUS.length !== CONVOQUES.length) {
    console.warn("bilan-forum : les tableaux CONVOQUES / PRESENTS / RETENUS n'ont pas la même longueur ("
      + CONVOQUES.length + " / " + PRESENTS.length + " / " + RETENUS.length + ")");
  }

  var COLOR_CONVOQUES = "#0D6FA3";
  var COLOR_PRESENTS = "#00B3B8";
  var COLOR_RETENUS = "#FF5A1F";

  /* Fenêtre affichée : par défaut les dernières journées ; les boutons
     « Précédentes / Suivantes » font glisser la fenêtre, « Tout » l'ouvre
     sur l'ensemble des journées. */
  var PAGE_SIZE = 7;
  var chart = null;
  var start = Math.max(0, LABELS.length - PAGE_SIZE);
  var showAll = false;

  function sum(arr) {
    return arr.reduce(function (a, b) { return a + b; }, 0);
  }

  function fmtPct(v) {
    return v.toLocaleString("fr-FR", { minimumFractionDigits: 1, maximumFractionDigits: 1 }) + " %";
  }

  /* Indicateurs du bandeau : totaux et taux de perte calculés sur toutes
     les journées. On écrit data-count-to, que reveal.js anime à l'apparition. */
  function renderKpis() {
    var conv = sum(CONVOQUES), pres = sum(PRESENTS), ret = sum(RETENUS);
    var pertePresents = conv ? (conv - pres) / conv * 100 : 0;
    var perteRetenus = pres ? (pres - ret) / pres * 100 : 0;
    var perteGlobale = conv ? (conv - ret) / conv * 100 : 0;

    function setCount(id, value) {
      var el = document.getElementById(id);
      if (el) el.setAttribute("data-count-to", String(value));
    }
    function setText(id, text) {
      var el = document.getElementById(id);
      if (el) el.textContent = text;
    }

    setText("us-forum-days-count", String(LABELS.length));
    setCount("us-forum-kpi-convoques", conv);
    setCount("us-forum-kpi-presents", pres);
    setText("us-forum-kpi-presents-label", "Présents — " + fmtPct(pertePresents) + " de perte");
    setCount("us-forum-kpi-retenus", ret);
    setText("us-forum-kpi-retenus-label", "Retenus — " + fmtPct(perteRetenus) + " de perte vs présents");
    setCount("us-forum-kpi-perte", perteGlobale.toFixed(1));
  }

  function render() {
    if (!chart) return;
    var from = showAll ? 0 : start;
    var to = showAll ? LABELS.length : Math.min(LABELS.length, start + PAGE_SIZE);

    chart.setOption({
      xAxis: {
        data: LABELS.slice(from, to),
        /* sur la vue complète, on incline les libellés pour qu'ils tiennent tous */
        axisLabel: { fontSize: showAll ? 10 : 11, color: "#1e293b", interval: 0, rotate: showAll ? 45 : 0 }
      },
      series: [
        { data: CONVOQUES.slice(from, to) },
        { data: PRESENTS.slice(from, to) },
        { data: RETENUS.slice(from, to) }
      ]
    });

    var range = document.getElementById("us-forum-range");
    if (range) {
      range.textContent = showAll
        ? "Journées 1 – " + LABELS.length
        : "Journées " + (from + 1) + " – " + to;
    }
    var prev = document.getElementById("us-forum-prev");
    var next = document.getElementById("us-forum-next");
    var all = document.getElementById("us-forum-all");
    if (prev) prev.disabled = showAll || from === 0;
    if (next) next.disabled = showAll || to >= LABELS.length;
    if (all) all.textContent = showAll ? "Dernières journées" : "Tout afficher";
  }

  function bindControls() {
    var prev = document.getElementById("us-forum-prev");
    var next = document.getElementById("us-forum-next");
    var all = document.getElementById("us-forum-all");

    if (prev) prev.addEventListener("click", function () {
      start = Math.max(0, start - PAGE_SIZE);
      render();
    });
    if (next) next.addEventListener("click", function () {
      start = Math.min(Math.max(0, LABELS.length - PAGE_SIZE), start + PAGE_SIZE);
      render();
    });
    if (all) all.addEventListener("click", function () {
      showAll = !showAll;
      if (!showAll) start = Math.max(0, LABELS.length - PAGE_SIZE);
      render();
    });
  }

  function buildChart() {
    var dom = document.getElementById("us-chart-forum-days");
    if (!dom || typeof echarts === "undefined") return;
    chart = echarts.init(dom);
    chart.setOption({
      color: [COLOR_CONVOQUES, COLOR_PRESENTS, COLOR_RETENUS],
      tooltip: { trigger: "axis", axisPointer: { type: "shadow" } },
      legend: { data: ["Convoqués", "Présents", "Retenus"], top: 0, textStyle: { fontSize: 12, color: "#1e293b" } },
      grid: { left: "3%", right: "4%", bottom: "3%", top: "16%", containLabel: true },
      xAxis: { type: "category", data: [], axisLabel: { fontSize: 11, color: "#1e293b" } },
      yAxis: { type: "value", axisLabel: { fontSize: 11, color: "#1e293b" } },
      series: [
        { name: "Convoqués", type: "bar", data: [], barMaxWidth: 18, itemStyle: { borderRadius: [3, 3, 0, 0] } },
        { name: "Présents", type: "bar", data: [], barMaxWidth: 18, itemStyle: { borderRadius: [3, 3, 0, 0] } },
        { name: "Retenus", type: "bar", data: [], barMaxWidth: 18, itemStyle: { borderRadius: [3, 3, 0, 0] } }
      ]
    });
    bindControls();
    render();
    window.addEventListener("resize", function () { chart.resize(); });
  }

  document.addEventListener("DOMContentLoaded", function () {
    /* Les KPI sont remplis tout de suite : reveal.js (chargé après) lit
       data-count-to au moment où le bandeau apparaît. */
    renderKpis();
    var dom = document.getElementById("us-chart-forum-days");
    if (!dom) return;
    if (!document.body.classList.contains("us-locked")) { buildChart(); return; }
    var observer = new MutationObserver(function () {
      if (!document.body.classList.contains("us-locked")) { observer.disconnect(); buildChart(); }
    });
    observer.observe(document.body, { attributes: true, attributeFilter: ["class"] });
  });
})();
