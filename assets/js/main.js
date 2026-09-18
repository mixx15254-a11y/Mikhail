/* ДФК «Победа» · Самара */
(function () {
  "use strict";

  /* ---------- burger menu ---------- */
  var burger = document.getElementById("burger");
  var nav = document.getElementById("site-nav");
  if (burger && nav) {
    burger.addEventListener("click", function () {
      var open = nav.classList.toggle("open");
      burger.setAttribute("aria-expanded", String(open));
      burger.setAttribute("aria-label", open ? "Закрыть меню" : "Открыть меню");
    });
    nav.addEventListener("click", function (e) {
      if (e.target.tagName === "A") {
        nav.classList.remove("open");
        burger.setAttribute("aria-expanded", "false");
      }
    });
  }

  /* ---------- schedule filter ---------- */
  var filters = Array.prototype.slice.call(document.querySelectorAll(".sched-filter"));
  var cards = Array.prototype.slice.call(document.querySelectorAll(".sched-card"));
  filters.forEach(function (btn) {
    btn.addEventListener("click", function () {
      filters.forEach(function (b) {
        b.classList.toggle("is-active", b === btn);
        b.setAttribute("aria-pressed", String(b === btn));
      });
      var f = btn.getAttribute("data-filter");
      cards.forEach(function (card) {
        var show = f === "all" || card.getAttribute("data-loc") === f;
        card.classList.toggle("is-hidden", !show);
      });
    });
  });

  /* ---------- reveal on scroll ---------- */
  if ("IntersectionObserver" in window) {
    var targets = document.querySelectorAll(
      ".feature, .principle, .coach-card, .sched-card, .price-card, .loc-card, .event-card, .review-card, .faq-item, .about-art, .hero-art"
    );
    var io = new IntersectionObserver(
      function (entries) {
        entries.forEach(function (entry) {
          if (entry.isIntersecting) {
            entry.target.classList.add("is-visible");
            io.unobserve(entry.target);
          }
        });
      },
      { threshold: 0.12 }
    );
    targets.forEach(function (t) {
      t.classList.add("reveal");
      io.observe(t);
    });
  }

  /* ---------- footer year ---------- */
  var yearEl = document.getElementById("year");
  if (yearEl) yearEl.textContent = String(new Date().getFullYear());

  /* ---------- signup form (без бэкенда: готовит письмо / ссылку в VK) ---------- */
  var form = document.getElementById("signup-form");
  if (form) {
    form.addEventListener("submit", function (e) {
      e.preventDefault();
      var hint = document.getElementById("form-hint");
      var alt = document.getElementById("form-alt");
      var parent = document.getElementById("f-parent").value.trim();
      var phone = document.getElementById("f-phone").value.trim();
      var year = document.getElementById("f-year").value.trim();
      var loc = document.getElementById("f-location").value;
      var comment = document.getElementById("f-comment").value.trim();
      var consent = document.getElementById("f-consent").checked;

      if (!parent || !phone) {
        hint.className = "form-hint is-error";
        hint.textContent = "Пожалуйста, укажите имя и телефон — без них мы не сможем связаться с вами.";
        return;
      }
      if (!consent) {
        hint.className = "form-hint is-error";
        hint.textContent = "Нужно согласие на обработку персональных данных.";
        return;
      }

      var body =
        "Заявка с сайта ДФК «Победа»%0A%0A" +
        "Родитель: " + encodeURIComponent(parent) + "%0A" +
        "Телефон: " + encodeURIComponent(phone) + "%0A" +
        (year ? "Год рождения ребёнка: " + encodeURIComponent(year) + "%0A" : "") +
        (loc ? "Площадка: " + encodeURIComponent(loc) + "%0A" : "") +
        (comment ? "Комментарий: " + encodeURIComponent(comment) : "");

      var mail = document.getElementById("alt-mail");
      mail.href =
        "mailto:dshpobeda@mail.ru?subject=" +
        encodeURIComponent("Заявка на бесплатную неделю — " + parent) +
        "&body=" + body;

      hint.className = "form-hint";
      hint.textContent = "Спасибо, " + parent + "! Заявка сформирована.";
      alt.hidden = false;
    });
  }
})();
