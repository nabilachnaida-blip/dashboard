(function () {
  // Same light client-side gate as the main dashboard (js/dashboard.js) —
  // shares the passcode via sessionStorage so unlocking once covers every page.
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
