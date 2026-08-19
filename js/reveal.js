(function () {
  "use strict";

  var reduceMotion = window.matchMedia && window.matchMedia("(prefers-reduced-motion: reduce)").matches;

  function animateCount(el) {
    var to = parseFloat(el.getAttribute("data-count-to"));
    if (isNaN(to)) return;
    var decimals = parseInt(el.getAttribute("data-count-decimals") || "0", 10);
    var suffix = el.getAttribute("data-count-suffix") || "";
    if (reduceMotion) {
      el.textContent = to.toLocaleString("fr-FR", { minimumFractionDigits: decimals, maximumFractionDigits: decimals }) + suffix;
      return;
    }
    var duration = 1300;
    var start = null;
    function step(ts) {
      if (!start) start = ts;
      var progress = Math.min((ts - start) / duration, 1);
      var eased = 1 - Math.pow(1 - progress, 3);
      var value = to * eased;
      el.textContent = value.toLocaleString("fr-FR", { minimumFractionDigits: decimals, maximumFractionDigits: decimals }) + suffix;
      if (progress < 1) requestAnimationFrame(step);
    }
    requestAnimationFrame(step);
  }

  function reveal(el) {
    el.classList.add("us-visible");
    if (el.hasAttribute("data-count-to")) animateCount(el);
    el.querySelectorAll("[data-count-to]").forEach(animateCount);
  }

  function init() {
    var targets = document.querySelectorAll(".us-reveal, .us-reveal-stagger");
    if (!targets.length) return;

    if (reduceMotion || typeof IntersectionObserver === "undefined") {
      targets.forEach(reveal);
      return;
    }

    var observer = new IntersectionObserver(function (entries) {
      entries.forEach(function (entry) {
        if (entry.isIntersecting) {
          reveal(entry.target);
          observer.unobserve(entry.target);
        }
      });
    }, { threshold: 0.15, rootMargin: "0px 0px -40px 0px" });

    targets.forEach(function (el) { observer.observe(el); });
  }

  // Observing elements under a display:none ancestor is safe — the
  // IntersectionObserver just fires once the access gate unlocks and
  // <main> becomes visible, no separate gate-aware rescan needed.
  document.addEventListener("DOMContentLoaded", init);
})();
