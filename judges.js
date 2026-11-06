/* The judges: one record per person, and the bio sheet that shows it.
 *
 * Read by judges.html (the cards) and agenda.html (the judge row on day two).
 * Exposes window.Judges:
 *   .list            -> every judge, in the order they appear on the site
 *   .hosts           -> the co-hosts, shown above the judges
 *   .initials(name)  -> "MS" for "Matshepo Soto"
 *   .avatar(j, size) -> a round photo, or initials if the photo is missing
 *   .open(slug)      -> show that judge's bio sheet
 *
 * Bios are drawn from public profiles (company sites, press). Anyone without
 * a confirmed bio shows a short "coming soon" line instead.
 */
(function () {
  "use strict";

  var JUDGING = [
    { day: "Day 2 · Saturday 14 November", time: "12:00", what: "Live online presentations", href: "agenda.html#day2" },
    { day: "Day 2 · Saturday 14 November", time: "13:00", what: "Winner announcement", href: "agenda.html#day2" }
  ];

  var JUDGES = [
    {
      slug: "francois-van-der-merwe",
      name: "Francois van der Merwe",
      title: "Founder and CEO",
      org: "Otinga.io",
      photo: "assets/judges/francois.webp",
      bio: [
        "Francois van der Merwe is the founder and CEO of Otinga.io, a South African company that runs AI-powered hackathons and innovation programmes for large enterprises, including banks and insurers in South Africa and abroad.",
        "With a hands-on background in AI and technology, Francois has written about why AI projects stall when organisations skip the fundamentals and jump straight to choosing a model or vendor."
      ]
    },
    {
      slug: "sizwe-mabanga",
      name: "Sizwe Mabanga",
      title: "Chief Technology Officer",
      org: "Mpilo Technologies",
      photo: "assets/judges/sizwe.webp",
      bio: [
        "Sizwe Mabanga is the Chief Technology Officer of Mpilo Technologies, a South African software development and digital solutions company, appointed in 2025 as the company marked its tenth year.",
        "Sizwe brings experience in cloud technology, software engineering and digital transformation, with deep skills in agile delivery, multi-cloud environments such as AWS and Azure, and digital banking platforms including Volt MX and Temenos."
      ]
    },
    {
      slug: "sanjay-soni",
      name: "Sanjay Soni",
      title: "Executive Director",
      org: "Black Lite Group",
      photo: "assets/judges/sanjay.webp",
      bio: [
        "Sanjay Soni is a chartered accountant, CA(SA), and Executive Director of Black Lite Group, a B-BBEE strategy and business advisory firm that works with South African corporates, multinationals and growing businesses.",
        "Sanjay brings around 25 years of advisory experience across corporate finance, ownership structuring and high-impact transactions."
      ]
    },
    {
      slug: "chido-dzinotyiwei",
      name: "Chido Dzinotyiwei",
      title: "Co-founder and CEO",
      org: "Vambo AI",
      photo: "assets/judges/chido.webp",
      bio: [
        "Chido Dzinotyiwei is the co-founder and CEO of Vambo AI, a multilingual generative AI company building language technology for African and global languages, from translation and transcription to content generation and AI-powered search.",
        "Chido’s work has been recognised with the Mandela Rhodes Foundation’s Äänit Prize, and Vambo AI received a Meta Llama Impact Grant in 2025. Chido holds a Master’s in Development Finance from the University of Cape Town and previously worked in banking and venture development at Standard Bank."
      ]
    },
    {
      slug: "tshegofhatso",
      name: "Tshegofhatso",
      title: "Software Engineer",
      org: "",
      photo: "assets/judges/tshego.webp",
      bio: []
    },
    {
      slug: "matshepo-soto",
      name: "Matshepo Soto",
      title: "Software Engineer",
      org: "Nedbank",
      photo: "assets/judges/matshepo.webp",
      bio: [
        "Matshepo Soto is a software engineer at Nedbank, working as a backend developer on the bank’s payment systems. Matshepo trained as a software engineer at WeThinkCode.",
        "In 2024 Matshepo won seven coding competitions and the Dynamism Youth in ICT and Rising Star in Coding awards, and has twice been named among the 100 Influential Women in Tech. Matshepo also leads Google Developer Group Pretoria."
      ]
    }
  ];

  // The co-hosts: shown above the judges, without a bio sheet.
  var HOSTS = [
    { name: "Buhle Dlomo", title: "Co-host", photo: "assets/judges/buhle.webp" },
    { name: "Lhoza Horo", title: "Co-host", photo: "assets/judges/lhoza.webp" }
  ];

  JUDGES.forEach(function (j) { j.on = JUDGING; });

  var bySlug = {};
  JUDGES.forEach(function (j) { bySlug[j.slug] = j; });

  function initials(name) {
    var parts = String(name).split(/\s+/).filter(Boolean);
    var first = parts[0] ? parts[0][0] : "?";
    var last = parts.length > 1 ? parts[parts.length - 1][0] : "";
    return (first + last).toUpperCase();
  }

  function avatar(j, size) {
    var box = document.createElement("span");
    box.className = "av";
    box.setAttribute("aria-hidden", "true");
    box.style.width = size + "px";
    box.style.height = size + "px";
    box.style.fontSize = Math.round(size * 0.34) + "px";
    function fallback() { box.textContent = initials(j.name); }
    if (j.photo) {
      var img = document.createElement("img");
      img.src = j.photo;
      img.alt = "";
      img.width = size;
      img.height = size;
      // A missing file falls back to initials rather than a blank circle.
      img.addEventListener("error", function () { box.removeChild(img); fallback(); });
      box.appendChild(img);
    } else {
      fallback();
    }
    return box;
  }

  function el(tag, cls, text) {
    var n = document.createElement(tag);
    if (cls) n.className = cls;
    if (text != null) n.textContent = text;
    return n;
  }

  var back = null, lastFocus = null;

  function close() {
    if (!back) return;
    back.remove();
    back = null;
    document.body.classList.remove("bio-open");
    document.removeEventListener("keydown", onKey);
    if (lastFocus && lastFocus.focus) lastFocus.focus();
  }

  function onKey(e) {
    if (e.key === "Escape") close();
    // Keep keyboard focus inside the sheet while it is open.
    if (e.key === "Tab" && back) {
      var f = back.querySelectorAll("button, a[href]");
      var first = f[0], last = f[f.length - 1];
      if (e.shiftKey && document.activeElement === first) { e.preventDefault(); last.focus(); }
      else if (!e.shiftKey && document.activeElement === last) { e.preventDefault(); first.focus(); }
    }
  }

  function open(slug) {
    var j = bySlug[slug];
    if (!j) return;
    close();
    lastFocus = document.activeElement;

    back = el("div", "bio-back");
    back.addEventListener("click", function (e) { if (e.target === back) close(); });

    var sheet = el("div", "bio-sheet");
    sheet.setAttribute("role", "dialog");
    sheet.setAttribute("aria-modal", "true");
    sheet.setAttribute("aria-labelledby", "bio-name");

    var grab = el("div", "bio-grab");
    grab.appendChild(el("span"));
    sheet.appendChild(grab);

    var x = el("button", "bio-close");
    x.type = "button";
    x.setAttribute("aria-label", "Close");
    x.innerHTML = '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" aria-hidden="true"><path d="M18 6 6 18M6 6l12 12"/></svg>';
    x.addEventListener("click", close);
    sheet.appendChild(x);

    var head = el("div", "bio-head");
    head.appendChild(avatar(j, 76));
    var who = el("div");
    var h = el("h2", "bio-name", j.name);
    h.id = "bio-name";
    who.appendChild(h);
    who.appendChild(el("p", "bio-title", j.title));
    if (j.org) who.appendChild(el("p", "bio-org", j.org));
    head.appendChild(who);
    sheet.appendChild(head);

    var chips = el("div", "bio-chips");
    chips.appendChild(el("span", "chip solid", "Judge"));
    sheet.appendChild(chips);

    if (j.bio.length) {
      j.bio.forEach(function (p) { sheet.appendChild(el("p", "bio-text", p)); });
    } else {
      sheet.appendChild(el("p", "bio-text soon", "Full biography coming soon."));
    }

    var when = el("div", "bio-when");
    when.appendChild(el("h4", "label", "Where to find them"));
    j.on.forEach(function (o) {
      var a = el("a", "bio-slot");
      a.href = o.href;
      a.appendChild(el("b", null, o.day + " · " + o.time));
      a.appendChild(el("span", null, o.what));
      // Already on the agenda: close the sheet and let the hash open day two.
      a.addEventListener("click", function () { close(); });
      when.appendChild(a);
    });
    sheet.appendChild(when);

    back.appendChild(sheet);
    document.body.appendChild(back);
    document.body.classList.add("bio-open");
    document.addEventListener("keydown", onKey);
    x.focus();
  }

  window.Judges = { list: JUDGES, hosts: HOSTS, initials: initials, avatar: avatar, open: open };
})();
