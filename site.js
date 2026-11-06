/* Shared behaviour for the event pages (agenda, judges, what's needed, RSVP):
 *   - the live backdrop in the page head: Vanta's NET effect in the event's plum, or pink
 *     on midnight in a dark head (vantajs.com), desktop only and loaded on demand, so phones
 *     and reduced-motion visitors never download three.js and keep the plain grid instead;
 *   - cells that rise into place as they scroll into view ([data-reveal] > *).
 * Loaded with defer, so it runs after each page's own script has built its lists.
 */
(function () {
  "use strict";
  var calm = window.matchMedia("(prefers-reduced-motion: reduce)").matches;

  function load(src) {
    return new Promise(function (done, fail) {
      var s = document.createElement("script");
      s.src = src;
      s.onload = done;
      s.onerror = fail;
      document.head.appendChild(s);
    });
  }

  var net = document.querySelector(".page-head-net");
  // Only once site.css is in: without it the canvas would sit in the flow over the headline.
  var styled = net && window.getComputedStyle(net).position === "absolute";
  // The dark head (rsvp.html) gets pink lines on midnight; the others plum on white.
  var dark = net && !!net.closest(".page-head-dark");
  if (styled && !calm && window.matchMedia("(min-width: 761px)").matches) {
    load("https://cdnjs.cloudflare.com/ajax/libs/three.js/r134/three.min.js")
      .then(function () { return load("https://cdn.jsdelivr.net/npm/vanta@0.5.24/dist/vanta.net.min.js"); })
      .then(function () {
        window.VANTA.NET({
          el: net,
          mouseControls: true, touchControls: false, gyroControls: false,
          minHeight: 200, minWidth: 200, scale: 1, scaleMobile: 1,
          color: dark ? 0xe48cc4 : 0x5e0743, backgroundColor: dark ? 0x070f24 : 0xffffff,
          points: 9, maxDistance: 21, spacing: 17, showDots: true
        });
        net.classList.add("on");
      })
      .catch(function () { /* No WebGL, or offline: the plain grid stays. */ });
  }

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
