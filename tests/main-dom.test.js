/* Функциональная проверка assets/js/main.js без браузера.
 * Запуск: node tests/main-dom.test.js
 * Даёт скрипту минимальный DOM и вызывает его же обработчики событий. */
"use strict";
const fs = require("fs");
const path = require("path");
const assert = require("assert");

function makeClassList(el) {
  const set = new Set(el.initialClasses || []);
  return {
    add: (c) => set.add(c),
    remove: (c) => set.delete(c),
    toggle: (c, force) => {
      const want = force === undefined ? !set.has(c) : force;
      want ? set.add(c) : set.delete(c);
      return want;
    },
    contains: (c) => set.has(c),
  };
}

function makeEl(opts) {
  const el = {
    tagName: (opts.tagName || "div").toUpperCase(),
    id: opts.id || "",
    attrs: Object.assign({}, opts.attrs || {}),
    initialClasses: opts.classes || [],
    listeners: {},
    value: opts.value !== undefined ? opts.value : "",
    checked: !!opts.checked,
    hidden: !!opts.hidden,
    textContent: "",
    href: opts.href || "",
    className: "",
  };
  el.classList = makeClassList(el);
  el.addEventListener = (type, fn) => {
    (el.listeners[type] = el.listeners[type] || []).push(fn);
  };
  el.setAttribute = (k, v) => {
    el.attrs[k] = v;
  };
  el.getAttribute = (k) => (k in el.attrs ? el.attrs[k] : null);
  el.fire = (type, ev) => {
    (el.listeners[type] || []).forEach((fn) => fn(ev || { preventDefault() {}, target: el }));
  };
  return el;
}

/* --- DOM-фикстура, соответствующая index.html --- */
const byId = {};
["burger", "site-nav", "year", "signup-form", "form-hint", "form-alt",
 "f-parent", "f-phone", "f-year", "f-location", "f-comment", "f-consent", "alt-mail"]
  .forEach((id) => { byId[id] = makeEl({ id, tagName: id.startsWith("f-") ? "input" : "div" }); });
byId["f-consent"].tagName = "INPUT";
byId["alt-mail"].tagName = "A";
byId["site-nav"].tagName = "NAV";

const navLink = makeEl({ tagName: "a" });
byId["site-nav"].containsLink = true;

const filters = ["all", "zd", "intel", "sputnik"].map((f, i) =>
  makeEl({ tagName: "button", attrs: { "data-filter": f }, classes: i === 0 ? ["is-active"] : [] }));

const cardData = [
  ["zd"], ["zd"], ["zd"], ["intel"], ["intel"], ["sputnik"], ["sputnik"], ["intel"],
];
const cards = cardData.map(([loc]) => makeEl({ tagName: "article", attrs: { "data-loc": loc } }));

global.window = {}; // без IntersectionObserver — ветка reveal пропускается
global.document = {
  getElementById: (id) => byId[id] || null,
  querySelectorAll: (sel) => {
    if (sel === ".sched-filter") return filters;
    if (sel === ".sched-card") return cards;
    return [];
  },
};

/* --- загружаем проверяемый скрипт --- */
const src = fs.readFileSync(path.join(__dirname, "..", "assets", "js", "main.js"), "utf8");
eval(src);

/* 1. бургер открывает/закрывает меню */
byId["burger"].fire("click");
assert.strictEqual(byId["site-nav"].classList.contains("open"), true, "меню открылось");
assert.strictEqual(byId["burger"].attrs["aria-expanded"], "true", "aria-expanded=true");
byId["burger"].fire("click");
assert.strictEqual(byId["site-nav"].classList.contains("open"), false, "меню закрылось");
console.log("PASS burger menu");

/* 2. фильтр расписания скрывает чужие площадки */
filters[1].fire("click"); // «ЖД»
const visibleAfterZd = cards.filter((c) => !c.classList.contains("is-hidden"));
assert.strictEqual(visibleAfterZd.length, 3, "видны 3 группы «ЖД», а не " + visibleAfterZd.length);
assert.strictEqual(filters[1].attrs["aria-pressed"], "true", "aria-pressed у активной");
assert.strictEqual(filters[0].attrs["aria-pressed"], "false", "у «все» сброшен");
filters[0].fire("click"); // все
assert.strictEqual(cards.filter((c) => !c.classList.contains("is-hidden")).length, 8, "снова 8 карточек");
console.log("PASS schedule filter");

/* 3. форма: пустые поля -> ошибка */
const form = byId["signup-form"];
form.fire("submit");
assert.ok(byId["form-hint"].className.includes("is-error"), "ошибка при пустых полях");
console.log("PASS form validation (empty)");

/* 4. форма: без согласия -> ошибка */
byId["f-parent"].value = "Иван";
byId["f-phone"].value = "+7 927 000-00-00";
byId["f-consent"].checked = false;
form.fire("submit");
assert.ok(byId["form-hint"].className.includes("is-error"), "ошибка без согласия");
console.log("PASS form validation (consent)");

/* 5. форма: валидная заявка -> mailto собран */
byId["f-consent"].checked = true;
byId["f-year"].value = "2018";
byId["f-location"].value = "«ЖД» — ул. Свободы, 2в";
byId["f-comment"].value = "Хотим к Пильщикову";
form.fire("submit");
assert.strictEqual(byId["form-alt"].hidden, false, "панель отправки показана");
assert.ok(byId["alt-mail"].href.startsWith("mailto:dshpobeda@mail.ru"), "mailto на почту клуба");
assert.ok(byId["alt-mail"].href.includes(encodeURIComponent("Иван")), "имя в письме");
assert.ok(byId["alt-mail"].href.includes(encodeURIComponent("2018")), "год в письме");
assert.ok(decodeURIComponent(byId["alt-mail"].href).includes("Хотим к Пильщикову"), "комментарий в письме");
assert.ok(!byId["form-hint"].className.includes("is-error"), "без ошибки");
console.log("PASS form submit (mailto)");

/* 6. год в подвале */
assert.strictEqual(byId["year"].textContent, String(new Date().getFullYear()), "год подставлен");
console.log("PASS footer year");

console.log("ALL TESTS PASSED");
