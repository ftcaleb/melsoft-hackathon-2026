/* Shared behaviour for the event pages (agenda, judges, what's needed, RSVP): cells that
 * rise into place as they scroll into view ([data-reveal] > *).
 * Loaded with defer, so it runs after each page's own script has built its lists.
 */
(function () {
  "use strict";
  var calm = window.matchMedia("(prefers-reduced-motion: reduce)").matches;

  // Scroll reveal, one after another within each group; skipped for reduced motion.
  if (calm || !("IntersectionObserver" in window)) return;
  var seen = new IntersectionObserver(function (entries) {
    entries.forEach(function (e) {
      if (!e.isIntersecting) return;
      var cell = e.target;
      cell.classList.add("in");
      seen.unobserve(cell);
      // Drop the stagger once it has played, so hover effects respond straight away.
      setTimeout(function () { cell.style.transitionDelay = ""; }, 1000);
    });
  }, { rootMargin: "0px 0px -8% 0px" });
  Array.prototype.forEach.call(document.querySelectorAll("[data-reveal]"), function (group) {
    Array.prototype.forEach.call(group.children, function (cell, i) {
      cell.style.transitionDelay = (i * 60) + "ms";
      cell.classList.add("rv");
      seen.observe(cell);
    });
  });
})();
