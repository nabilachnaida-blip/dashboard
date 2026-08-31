(function () {
  var LABELS = ["Journée 1", "Journée 2", "Journée 3", "Journée 4", "Journée 5", "Journée 6", "Journée 7", "Journée 8", "Journée 9", "Journée 10", "Journée 11", "Journée 12", "Journée 13", "Journée 14", "Journée 15", "Journée 16", "Journée 17"];
  var CONVOQUES = [230, 203, 207, 177, 184, 184, 180, 179, 227, 124, 240,169,116,184,192,161,230];
  var PRESENTS = [111, 97, 109, 83, 90, 90, 75, 99, 107, 49,103,65,56,106,98,81,100];
  var RETENUS = [109, 95, 106, 78, 83, 83, 70, 86, 102, 48, 95,63,52,101,95,77,100];

  var COLOR_CONVOQUES = "#0D6FA3";
  var COLOR_PRESENTS = "#00B3B8";
  var COLOR_RETENUS = "#FF5A1F";

  function buildChart() {
    var dom = document.getElementById("us-chart-forum-days");
    if (!dom || typeof echarts === "undefined") return;
    var chart = echarts.init(dom);
    chart.setOption({
      color: [COLOR_CONVOQUES, COLOR_PRESENTS, COLOR_RETENUS],
      tooltip: { trigger: "axis", axisPointer: { type: "shadow" } },
      legend: { data: ["Convoqués", "Présents", "Retenus"], top: 0, textStyle: { fontSize: 12, color: "#1e293b" } },
      grid: { left: "3%", right: "4%", bottom: "3%", top: "16%", containLabel: true },
      xAxis: { type: "category", data: LABELS, axisLabel: { fontSize: 11, color: "#1e293b" } },
      yAxis: { type: "value", axisLabel: { fontSize: 11, color: "#1e293b" } },
      series: [
        { name: "Convoqués", type: "bar", data: CONVOQUES, barMaxWidth: 18, itemStyle: { borderRadius: [3, 3, 0, 0] } },
        { name: "Présents", type: "bar", data: PRESENTS, barMaxWidth: 18, itemStyle: { borderRadius: [3, 3, 0, 0] } },
        { name: "Retenus", type: "bar", data: RETENUS, barMaxWidth: 18, itemStyle: { borderRadius: [3, 3, 0, 0] } }
      ]
    });
    window.addEventListener("resize", function () { chart.resize(); });
  }

  document.addEventListener("DOMContentLoaded", function () {
    var dom = document.getElementById("us-chart-forum-days");
    if (!dom) return;
    if (!document.body.classList.contains("us-locked")) { buildChart(); return; }
    var observer = new MutationObserver(function () {
      if (!document.body.classList.contains("us-locked")) { observer.disconnect(); buildChart(); }
    });
    observer.observe(document.body, { attributes: true, attributeFilter: ["class"] });
  });
})();
