/* التعافي المالي — خطة الإيراد والتسليم */
(function () {
  "use strict";

  const STORAGE_KEY = "financialRecovery.rows.v1";
  const META_KEY = "financialRecovery.meta.v1";

  /*
   * أعمدة الجدول (وهي نفسها أعمدة ملف Excel عند التصدير والاستيراد)
   * edit: طريقة التعديل عند الضغط على الخانة
   *   text | select | number | plan (رقم أو ملاحظة) | textarea | date | dialog (يفتح نافذة التعديل) | none
   */
  const COLUMNS = [
    { key: "projectNo", label: "رقم المشروع", edit: "text", width: 12 },
    { key: "status", label: "حالة العقد", edit: "select", options: ["جديد", "قائم"], width: 10 },
    { key: "segment", label: "Segment", edit: "text", list: "segmentList", width: 12 },
    { key: "entity", label: "الجهة", edit: "text", list: "entityList", width: 12 },
    { key: "portfolio", label: "المحفظة", edit: "text", list: "portfolioList", width: 24 },
    { key: "project", label: "Project Name", edit: "dialog", width: 45 },
    { key: "projectManager", label: "مدير المشروع", edit: "text", list: "pmList", width: 20 },
    { key: "revenue", label: "الإيراد - المالية", money: true, edit: "number", width: 18 },
    { key: "totalPlan", label: "اجمالي الخطط للتسليم", money: true, edit: "plan", width: 18 },
    { key: "plannedToNov", label: "التسليم المخطط حتى نوفمبر", money: true, edit: "number", width: 18 },
    { key: "plannedDec", label: "التسليم المخطط ديسمبر", money: true, edit: "number", width: 18 },
    { key: "delivered", label: "مسلم", money: true, edit: "number", width: 18 },
    { key: "remaining", label: "المتبقي", money: true, edit: "none", width: 18 },
    { key: "action", label: "الإجراء المطلوب", edit: "textarea", width: 38 },
    { key: "owner", label: "المسؤول", edit: "textarea", width: 35 },
    { key: "nextStep", label: "الخطوة القادمة", edit: "textarea", width: 35 },
    { key: "dueDate", label: "تاريخ الاستحقاق", edit: "date", width: 14 },
  ];
  const COL = Object.fromEntries(COLUMNS.map((c) => [c.key, c]));
  const MONEY_KEYS = ["revenue", "totalPlan", "plannedToNov", "plannedDec", "delivered"];
  const TEXT_KEYS = COLUMNS.filter((c) => !c.money && c.key !== "dueDate").map((c) => c.key);
  const EDITABLE = COLUMNS.filter((c) => c.edit !== "none" && c.edit !== "dialog").map((c) => c.key);
  const HEAD_CLASS = { project: "col-project", projectManager: "col-pm", portfolio: "col-pf", action: "col-text", owner: "col-text", nextStep: "col-text" };

  // ---------- الحالة ----------
  let rows = load();
  let sort = { key: null, dir: 1 };
  let editingId = null;
  let activeEdit = null; // الخانة التي يجري تعديلها الآن
  const $ = (sel) => document.querySelector(sel);

  function load() {
    try {
      const saved = JSON.parse(localStorage.getItem(STORAGE_KEY));
      if (Array.isArray(saved)) return migrate(saved);
    } catch (e) { console.warn("تعذر قراءة البيانات المحفوظة", e); }
    return JSON.parse(JSON.stringify(window.SEED_DATA || []));
  }
  function norm(t) {
    return String(t || "").trim().replace(/[إأآ]/g, "ا").replace(/ة/g, "ه").replace(/ى/g, "ي").replace(/\s+/g, " ");
  }
  // يكمل الحقول التي أضيفت لاحقاً (مدير المشروع، رقم المشروع، المحفظة) في البيانات المحفوظة سابقاً
  function migrate(list) {
    const byName = new Map((window.SEED_DATA || []).map((r) => [norm(r.project), r]));
    list.forEach((r) => {
      const seed = byName.get(norm(r.project)) || {};
      ["projectManager", "projectNo", "portfolio"].forEach((k) => {
        if (r[k] === undefined) r[k] = seed[k] || "";
      });
    });
    return list;
  }
  function save() {
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(rows));
      localStorage.setItem(META_KEY, JSON.stringify({ updatedAt: new Date().toISOString() }));
    } catch (e) { toast("تعذر الحفظ في المتصفح"); }
    renderLastUpdate();
  }
  const rowById = (id) => rows.find((x) => x.id === id);

  // ---------- الحسابات ----------
  const isNum = (v) => typeof v === "number" && isFinite(v);
  const n = (v) => (isNum(v) ? v : 0);
  const round2 = (v) => Math.round(v * 100) / 100;
  function remaining(r) {
    return isNum(r.totalPlan) ? round2(r.totalPlan - n(r.delivered)) : null;
  }
  function deliveredOk(r) {
    return isNum(r.delivered) && r.delivered > 0 && isNum(r.plannedToNov) && r.delivered >= r.plannedToNov;
  }
  function today() {
    const d = new Date(); d.setHours(0, 0, 0, 0); return d;
  }
  function daysUntil(r) {
    return Math.round((new Date(r.dueDate + "T00:00:00") - today()) / 86400000);
  }
  function dueState(r) {
    if (!r.dueDate) return r.dueRaw ? "invalid" : "none";
    const diff = daysUntil(r);
    if (diff < 0) return "overdue";
    if (diff <= 7) return "soon";
    return "later";
  }
  function daysLabel(r) {
    const diff = daysUntil(r);
    if (diff === 0) return "اليوم";
    if (diff < 0) return `متأخر ${-diff} يوم`;
    return `بعد ${diff} يوم`;
  }

  // ---------- التنسيق ----------
  const fmt2 = new Intl.NumberFormat("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  const fmt0 = new Intl.NumberFormat("en-US", { maximumFractionDigits: 0 });
  const fmtM = new Intl.NumberFormat("en-US", { minimumFractionDigits: 1, maximumFractionDigits: 1 });
  function money(v) {
    if (!isNum(v)) return "";
    if (v === 0) return "-";
    return v < 0 ? `(${fmt2.format(-v)})` : fmt2.format(v);
  }
  const millions = (v) => `${fmtM.format(v / 1e6)} مليون`;
  function esc(s) {
    return String(s == null ? "" : s).replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  }
  function fmtDate(iso) {
    const [y, m, d] = iso.split("-");
    return `${d}/${m}/${y}`;
  }

  // ---------- الفلاتر والترتيب ----------
  const FILTERS = ["#searchInput", "#fStatus", "#fSegment", "#fEntity", "#fPortfolio", "#fPM", "#fDue"];
  const collator = new Intl.Collator("ar", { numeric: true, sensitivity: "base" });

  function filtered() {
    const q = $("#searchInput").value.trim().toLowerCase();
    const fs = $("#fStatus").value, fg = $("#fSegment").value, fe = $("#fEntity").value,
      ff = $("#fPortfolio").value, fp = $("#fPM").value, fd = $("#fDue").value;
    let out = rows.filter((r) => {
      if (fs && r.status !== fs) return false;
      if (fg && r.segment !== fg) return false;
      if (fe && r.entity !== fe) return false;
      if (ff && (ff === "__none__" ? r.portfolio : r.portfolio !== ff)) return false;
      if (fp && (fp === "__none__" ? r.projectManager : r.projectManager !== fp)) return false;
      if (fd) {
        const s = dueState(r);
        if (fd === "none" ? !(s === "none" || s === "invalid") : s !== fd) return false;
      }
      if (q) {
        const hay = [r.projectNo, r.project, r.portfolio, r.projectManager, r.owner, r.action, r.nextStep, r.segment, r.entity, r.status, r.planNote].join(" ").toLowerCase();
        if (!hay.includes(q)) return false;
      }
      return true;
    });
    if (sort.key) {
      const k = sort.key;
      const val = (r) => (k === "remaining" ? remaining(r) : k === "totalPlan" && !isNum(r.totalPlan) ? r.planNote : r[k]);
      out = out.slice().sort((a, b) => {
        const va = val(a), vb = val(b);
        const ea = va == null || va === "", eb = vb == null || vb === "";
        if (ea && eb) return 0;
        if (ea) return 1; // الخانات الفارغة دائماً في الأسفل
        if (eb) return -1;
        if (isNum(va) && isNum(vb)) return (va - vb) * sort.dir;
        if (isNum(va) !== isNum(vb)) return isNum(va) ? -1 : 1;
        return collator.compare(String(va), String(vb)) * sort.dir;
      });
    }
    return out;
  }

  function fillSelect(sel, values, placeholder, withNone) {
    const cur = sel.value;
    sel.innerHTML = `<option value="">${placeholder}</option>` +
      values.map((v) => `<option value="${esc(v)}">${esc(v)}</option>`).join("") +
      (withNone ? `<option value="__none__">غير محدد</option>` : "");
    if (values.includes(cur) || (withNone && cur === "__none__")) sel.value = cur;
  }
  const uniq = (key) => [...new Set(rows.map((r) => r[key]).filter(Boolean))].sort(collator.compare);
  const fillList = (id, values) => { $("#" + id).innerHTML = values.map((v) => `<option value="${esc(v)}">`).join(""); };

  function refreshFilterOptions() {
    fillSelect($("#fStatus"), uniq("status"), "حالة العقد: الكل");
    fillSelect($("#fSegment"), uniq("segment"), "Segment: الكل");
    fillSelect($("#fEntity"), uniq("entity"), "الجهة: الكل");
    fillSelect($("#fPortfolio"), uniq("portfolio"), "المحفظة: الكل", true);
    fillSelect($("#fPM"), uniq("projectManager"), "مدير المشروع: الكل", true);
    fillList("segmentList", uniq("segment"));
    fillList("entityList", uniq("entity"));
    fillList("portfolioList", uniq("portfolio"));
    fillList("pmList", uniq("projectManager"));
  }

  // ---------- رأس الجدول ----------
  const SORT_ICON = `<svg class="sort-icon" viewBox="0 0 10 14" width="9" height="13" aria-hidden="true"><path class="up" d="M5 1l4 5H1z"/><path class="down" d="M5 13L1 8h8z"/></svg>`;
  function renderHead() {
    const th = (key, label, cls = "") => {
      const active = sort.key === key;
      const state = active ? (sort.dir === 1 ? "asc" : "desc") : "";
      const aria = active ? (sort.dir === 1 ? "ascending" : "descending") : "none";
      const title = active ? (sort.dir === 1 ? "مرتب تصاعدياً (أ → ي)، اضغط للتنازلي" : "مرتب تنازلياً (ي → أ)، اضغط لإلغاء الترتيب") : "اضغط للترتيب أبجدياً";
      return `<th class="${cls} ${state ? "sorted " + state : ""}" aria-sort="${aria}">
        <button type="button" class="th-sort" data-sort="${key}" title="${title}"><span>${esc(label)}</span>${SORT_ICON}</button></th>`;
    };
    $("#gridHead").innerHTML = `<tr>${th("id", "#", "col-num")}${COLUMNS.map((c) =>
      th(c.key, c.label, [c.money ? "money" : "", HEAD_CLASS[c.key] || ""].join(" ").trim())).join("")}</tr>`;
  }

  // ---------- العرض ----------
  function render() {
    const data = filtered();
    renderHead();
    renderTable(data);
    renderSummary(data);
  }

  // المؤشرات والإجماليات وعدد الصفوف، دون إعادة رسم الصفوف
  function renderSummary(data = filtered()) {
    renderKpis(data);
    renderFoot(data);
    $("#rowCount").textContent = data.length === rows.length ? `${rows.length} مشروع` : `${data.length} من ${rows.length}`;
    $("#emptyState").hidden = data.length > 0;
  }

  function sum(data, key) {
    return round2(data.reduce((s, r) => s + n(key === "remaining" ? remaining(r) : r[key]), 0));
  }

  function renderKpis(data) {
    const revenue = sum(data, "revenue");
    const plan = sum(data, "totalPlan");
    const toNov = sum(data, "plannedToNov");
    const delivered = sum(data, "delivered");
    const rem = sum(data, "remaining");
    const pct = plan > 0 ? Math.min(100, (delivered / plan) * 100) : 0;
    const novPct = toNov > 0 ? Math.min(100, (delivered / toNov) * 100) : 0;
    const noContract = data.filter((r) => r.planNote).length;
    const overdue = data.filter((r) => dueState(r) === "overdue").length;
    const soon = data.filter((r) => dueState(r) === "soon").length;

    const card = (label, value, sub, color, bar) => `
      <div class="kpi" style="--kpi-color:${color}">
        <div class="kpi-label">${label}</div>
        <div class="kpi-value">${value}</div>
        <div class="kpi-sub">${sub}</div>
        ${bar != null ? `<div class="progress"><span style="width:${bar.toFixed(1)}%"></span></div>` : ""}
      </div>`;

    $("#kpis").innerHTML = [
      card("الإيراد - المالية", `${millions(revenue)}<small>ر.س</small>`, fmt0.format(revenue), "var(--accent)"),
      card("اجمالي الخطط للتسليم", `${millions(plan)}<small>ر.س</small>`, fmt0.format(plan), "var(--primary)"),
      card("المسلَّم", `${millions(delivered)}<small>ر.س</small>`, `${pct.toFixed(1)}% من إجمالي الخطط`, "var(--success)", pct),
      card("التسليم مقابل المخطط حتى نوفمبر", `${novPct.toFixed(1)}%`, `${millions(delivered)} من ${millions(toNov)}`, "#2E7DB8", novPct),
      card("المتبقي", `${millions(rem)}<small>ر.س</small>`, fmt0.format(rem), "var(--warning)"),
      card("المشاريع", `${data.length}`, `${noContract} لم يصل العقد`, "var(--primary-600)"),
      card("الاستحقاقات", `${overdue}<small>متأخر</small>`, `${soon} تستحق خلال 7 أيام`, "var(--danger)"),
    ].join("");
  }

  function renderFoot(data) {
    if (!data.length) { $("#gridFoot").innerHTML = ""; return; }
    const first = COLUMNS.findIndex((c) => c.money);
    const last = COLUMNS.length - 1 - [...COLUMNS].reverse().findIndex((c) => c.money);
    $("#gridFoot").innerHTML = `<tr>
      <td colspan="${first + 1}">الإجمالي</td>
      ${COLUMNS.slice(first, last + 1).map((c) => `<td class="money">${money(sum(data, c.key))}</td>`).join("")}
      <td colspan="${COLUMNS.length - 1 - last}"></td></tr>`;
  }

  const muted = (t) => `<span class="badge badge-muted">${t}</span>`;

  // محتوى كل خانة: { cls, html }
  function cellView(r, key) {
    const v = r[key];
    switch (key) {
      case "projectNo":
        return { cls: "col-no", html: v ? `<span class="proj-no">${esc(v)}</span>` : "" };
      case "status":
        return { cls: "cell-tinted", html: v ? `<span class="badge ${v === "قائم" ? "badge-active" : "badge-new"}">${esc(v)}</span>` : muted("غير محدد") };
      case "segment":
        return { cls: "cell-lav", html: v ? `<span class="badge badge-seg">${esc(v)}</span>` : "" };
      case "entity":
        return { cls: "cell-lav", html: v ? `<span class="badge ${v === "MOE" ? "badge-moe" : "badge-nonmoe"}">${esc(v)}</span>` : "" };
      case "portfolio":
        return { cls: "col-pf", html: v ? `<span class="pf">${esc(v)}</span>` : muted("غير محدد") };
      case "project":
        return { cls: "col-project cell-lav", html: `<button type="button" class="project-link" data-open="${r.id}" title="فتح شاشة التعديل">${esc(v)}</button>` };
      case "projectManager":
        return { cls: "col-pm", html: v ? `<span class="pm"><span class="pm-avatar">${esc(v.trim().charAt(0))}</span>${esc(v)}</span>` : muted("غير محدد") };
      case "totalPlan":
        if (!isNum(v)) return { cls: "money", html: r.planNote ? `<span class="badge badge-note">${esc(r.planNote)}</span>` : "" };
        return moneyView(v);
      case "plannedToNov":
      case "delivered":
        return moneyView(v, deliveredOk(r) ? "ok" : "");
      case "revenue":
        return moneyView(v, "revenue cell-tinted");
      case "plannedDec":
        return moneyView(v);
      case "remaining": {
        const rem = remaining(r);
        const pct = isNum(r.totalPlan) && r.totalPlan > 0 ? Math.min(100, (n(r.delivered) / r.totalPlan) * 100) : null;
        const view = moneyView(rem, "computed");
        if (pct != null) view.html += `<div class="mini-bar" title="نسبة التسليم ${pct.toFixed(1)}%"><span style="width:${pct.toFixed(1)}%"></span></div>`;
        return view;
      }
      case "dueDate": {
        const ds = dueState(r);
        const cls = ds === "overdue" ? "due-overdue" : ds === "soon" ? "due-soon" : "";
        if (ds === "invalid") return { cls, html: `<span class="due invalid" title="تاريخ غير صالح في الملف الأصلي"><span class="due-date">${esc(r.dueRaw)}</span><small>تاريخ غير صالح</small></span>` };
        if (!r.dueDate) return { cls, html: "" };
        return { cls, html: `<span class="due ${ds}"><span class="due-date">${fmtDate(r.dueDate)}</span><small>${daysLabel(r)}</small></span>` };
      }
      default:
        return { cls: "col-text", html: esc(v) };
    }
  }
  function moneyView(v, extra = "") {
    const cls = ["money", extra];
    if (isNum(v) && v === 0) cls.push("zero");
    if (isNum(v) && v < 0) cls.push("neg");
    return { cls: cls.join(" ").trim(), html: money(v) };
  }

  function rowHtml(r, num) {
    return `<tr data-id="${r.id}"><td class="col-num">${num}</td>${COLUMNS.map((c) => {
      const { cls, html } = cellView(r, c.key);
      const editable = EDITABLE.includes(c.key);
      return `<td class="${cls}${editable ? " editable" : ""}"${editable ? ` data-field="${c.key}" title="اضغط للتعديل"` : ""}>${html}</td>`;
    }).join("")}</tr>`;
  }

  function renderTable(data) {
    activeEdit = null;
    $("#gridBody").innerHTML = data.map((r, i) => rowHtml(r, i + 1)).join("");
  }

  // يعيد رسم صف واحد فقط في مكانه (بعد تعديل خانة)
  function updateRow(id) {
    const tr = $(`#gridBody tr[data-id="${id}"]`);
    const r = rowById(id);
    if (!tr || !r) return null;
    const tpl = document.createElement("tbody");
    tpl.innerHTML = rowHtml(r, tr.firstElementChild.textContent);
    const fresh = tpl.firstElementChild;
    tr.replaceWith(fresh);
    return fresh;
  }

  function renderLastUpdate() {
    let txt = "";
    try {
      const meta = JSON.parse(localStorage.getItem(META_KEY));
      if (meta && meta.updatedAt) {
        txt = "آخر تحديث: " + new Date(meta.updatedAt).toLocaleString("ar-SA-u-ca-gregory-nu-latn", { dateStyle: "medium", timeStyle: "short" });
      }
    } catch (e) { /* تجاهل */ }
    $("#lastUpdate").textContent = txt;
  }

  // ---------- التعديل المباشر داخل الخانة ----------
  function editorValue(r, key) {
    if (key === "totalPlan") return isNum(r.totalPlan) ? String(r.totalPlan) : (r.planNote || "");
    const v = r[key];
    return v == null ? "" : String(v);
  }

  function startEdit(td) {
    if (activeEdit) commitEdit();
    const tr = td.closest("tr");
    if (!tr || !td.isConnected) return;
    const id = Number(tr.dataset.id), key = td.dataset.field, col = COL[key], r = rowById(id);
    if (!r || !col) return;
    const value = editorValue(r, key);
    let el;
    if (col.edit === "select") {
      el = document.createElement("select");
      const opts = [...new Set(["", ...col.options, ...uniq(key)])];
      el.innerHTML = opts.map((o) => `<option value="${esc(o)}">${o ? esc(o) : "غير محدد"}</option>`).join("");
      el.value = value;
      el.addEventListener("change", () => commitEdit());
    } else if (col.edit === "textarea") {
      el = document.createElement("textarea");
      el.rows = 3;
      el.value = value;
    } else {
      el = document.createElement("input");
      if (col.edit === "date") el.type = "date";
      else {
        el.type = "text";
        if (col.edit === "number" || col.edit === "plan") { el.inputMode = "decimal"; el.dir = "ltr"; }
        if (col.edit === "plan") el.placeholder = "رقم أو ملاحظة";
        if (col.list) el.setAttribute("list", col.list);
      }
      el.value = value;
    }
    el.className = "cell-input";
    el.setAttribute("aria-label", col.label);
    el.addEventListener("keydown", (e) => {
      if (e.key === "Escape") { e.preventDefault(); commitEdit(true); }
      else if (e.key === "Enter" && !(col.edit === "textarea" && e.shiftKey)) { e.preventDefault(); commitEdit(false, 0, true); }
      else if (e.key === "Tab") { e.preventDefault(); commitEdit(false, e.shiftKey ? -1 : 1); }
    });
    el.addEventListener("blur", () => { if (activeEdit && activeEdit.el === el) commitEdit(); });

    activeEdit = { id, key, el, td };
    td.classList.add("editing");
    td.innerHTML = "";
    td.appendChild(el);
    el.focus();
    if (el.select && el.type === "text") el.select();
  }

  // يحوّل ما كتبه المستخدم إلى قيمة؛ يرجع undefined إذا كانت القيمة غير صالحة
  function parseValue(key, raw) {
    const col = COL[key];
    const t = raw.trim();
    if (col.edit === "number") {
      if (!t) return null;
      const num = Number(t.replace(/[,\s]/g, ""));
      return isFinite(num) ? round2(num) : undefined;
    }
    if (col.edit === "date") return t || null;
    return t;
  }

  function commitEdit(cancel = false, move = 0, focusCell = false) {
    const a = activeEdit;
    if (!a) return;
    activeEdit = null;
    const r = rowById(a.id);
    let changed = false;
    if (!cancel && r) {
      const raw = a.el.value;
      if (a.key === "totalPlan") {
        const t = raw.trim();
        const num = Number(t.replace(/[,\s]/g, ""));
        const next = !t ? { totalPlan: null, planNote: "" }
          : isFinite(num) ? { totalPlan: round2(num), planNote: "" }
          : { totalPlan: null, planNote: t };
        changed = next.totalPlan !== r.totalPlan || next.planNote !== (r.planNote || "");
        Object.assign(r, next);
      } else {
        const v = parseValue(a.key, raw);
        if (v === undefined) toast("القيمة يجب أن تكون رقماً");
        else if (v !== (r[a.key] == null ? (COL[a.key].edit === "number" || COL[a.key].edit === "date" ? null : "") : r[a.key])) {
          r[a.key] = v;
          if (a.key === "dueDate") r.dueRaw = "";
          changed = true;
        }
      }
    }
    const fresh = updateRow(a.id);
    if (changed) {
      save();
      refreshFilterOptions();
      renderSummary();
      toast("تم الحفظ");
    }
    if (!fresh) return;
    if (move) {
      const cells = [...fresh.querySelectorAll("td[data-field]")];
      const i = cells.findIndex((c) => c.dataset.field === a.key) + move;
      if (cells[i]) startEdit(cells[i]);
    } else if (focusCell) {
      const td = fresh.querySelector(`td[data-field="${a.key}"]`);
      if (td) { td.tabIndex = -1; td.focus({ preventScroll: true }); }
    }
  }

  // ---------- نافذة الإضافة / التعديل ----------
  const dlg = $("#editDialog");
  const form = $("#editForm");

  function openEditor(id) {
    if (activeEdit) commitEdit();
    editingId = id;
    const r = id != null ? rowById(id) : {};
    $("#dialogTitle").textContent = id != null ? "تعديل مشروع" : "مشروع جديد";
    $("#deleteBtn").hidden = id == null;
    for (const el of form.elements) {
      if (!el.name) continue;
      const v = r[el.name];
      el.value = v == null ? "" : v;
    }
    dlg.showModal();
    form.elements.project.focus();
  }

  form.addEventListener("submit", (e) => {
    e.preventDefault();
    const data = {};
    for (const el of form.elements) {
      if (!el.name) continue;
      if (MONEY_KEYS.includes(el.name)) data[el.name] = el.value === "" ? null : round2(parseFloat(el.value));
      else data[el.name] = el.value.trim();
    }
    data.dueDate = data.dueDate || null;
    if (isNum(data.totalPlan)) data.planNote = "";
    if (editingId != null) {
      const r = rowById(editingId);
      Object.assign(r, data);
      if (data.dueDate) r.dueRaw = "";
    } else {
      const id = rows.reduce((m, r) => Math.max(m, r.id), 0) + 1;
      rows.push(Object.assign({ id, dueRaw: "" }, data));
    }
    save();
    refreshFilterOptions();
    render();
    dlg.close();
    toast(editingId != null ? "تم حفظ التعديلات" : "تمت إضافة المشروع");
  });

  $("#deleteBtn").addEventListener("click", async () => {
    const r = rowById(editingId);
    if (!r || !(await askConfirm(`حذف المشروع "${r.project}"؟`, "حذف"))) return;
    rows = rows.filter((x) => x.id !== editingId);
    save(); refreshFilterOptions(); render();
    dlg.close();
    toast("تم حذف المشروع");
  });
  dlg.querySelectorAll("[data-close]").forEach((b) => b.addEventListener("click", () => dlg.close()));

  // ---------- Excel ----------
  // داخل عارض claude.ai يتم الحفظ عبر خاصية التنزيلات، وخارجه عبر المتصفح مباشرة
  const inViewer = !!(window.claude && typeof window.claude.use === "function");
  let downloadsApi = null;
  if (inViewer) {
    $("#exportBtn").hidden = true;
    window.claude.use("downloads").then((d) => {
      downloadsApi = d;
      $("#exportBtn").hidden = !d;
    }).catch(() => {});
  }

  // ألوان ملف Excel — نفس ألوان الجدول في النظام (بصيغة ARGB)
  const XL = {
    primary: "FF17355C", primary600: "FF1F4475", primary50: "FFEAEFF7", primary25: "FFF4F7FB",
    lavender: "FFE9EBF5", accent: "FF215C98", white: "FFFFFFFF", zebra: "FFFAFBFD",
    border: "FFDFE4EC", text: "FF1B2330", text2: "FF4B5668", muted: "FF7A8597",
    success: "FF3F8F2E", successBg: "FFE6F4DC", warning: "FFA86A00", warningBg: "FFFFF3D6",
    danger: "FFB3261E", dangerBg: "FFFCE8E6",
    newBg: "FFE7F0FB", moeBg: "FFE8F1EC", moe: "FF1E6B45", nonMoeBg: "FFF1ECF7", nonMoe: "FF5B3D8A",
  };
  const XL_FONT = "Arial";
  const XL_MONEY = '#,##0.00;(#,##0.00);"-"';
  const fill = (argb) => ({ type: "pattern", pattern: "solid", fgColor: { argb } });
  const font = (o = {}) => Object.assign({ name: XL_FONT, size: 10, color: { argb: XL.text } }, o);
  const thin = (argb = XL.border) => ({ style: "thin", color: { argb } });

  function activeFiltersText() {
    const parts = [];
    const q = $("#searchInput").value.trim();
    if (q) parts.push(`بحث: ${q}`);
    [["#fStatus", "حالة العقد"], ["#fSegment", "Segment"], ["#fEntity", "الجهة"], ["#fPortfolio", "المحفظة"], ["#fPM", "مدير المشروع"], ["#fDue", "الاستحقاق"]].forEach(([sel, label]) => {
      const el = $(sel);
      if (el.value) parts.push(`${label}: ${el.options[el.selectedIndex].text}`);
    });
    if (sort.key) {
      const name = sort.key === "id" ? "#" : COL[sort.key].label;
      parts.push(`مرتب حسب ${name} ${sort.dir === 1 ? "تصاعدياً" : "تنازلياً"}`);
    }
    return parts.length ? parts.join("  •  ") : "بدون فلاتر";
  }

  async function buildStyledWorkbook(data) {
    const wb = new ExcelJS.Workbook();
    wb.creator = "التعافي المالي";
    wb.created = new Date();
    const cols = [{ key: "_num", label: "#", width: 5 }, ...COLUMNS];
    const N = cols.length;
    const HEADER_ROW = 8;
    const ws = wb.addWorksheet("خطة الايراد والتسليم", {
      views: [{ rightToLeft: true, state: "frozen", ySplit: HEADER_ROW, showGridLines: false, zoomScale: 90 }],
      pageSetup: { orientation: "landscape", paperSize: 9, fitToPage: true, fitToWidth: 1, fitToHeight: 0, margins: { left: 0.3, right: 0.3, top: 0.4, bottom: 0.4, header: 0.2, footer: 0.2 } },
      properties: { defaultRowHeight: 18 },
    });
    ws.pageSetup.printTitlesRow = `${HEADER_ROW}:${HEADER_ROW}`;
    ws.columns = cols.map((c) => ({ width: c.width || 14 }));

    // 1) العنوان
    ws.mergeCells(1, 1, 1, N);
    const title = ws.getCell(1, 1);
    title.value = "التعافي المالي  |  خطة الإيراد والتسليم";
    title.font = font({ size: 16, bold: true, color: { argb: XL.white } });
    title.fill = fill(XL.primary);
    title.alignment = { vertical: "middle", horizontal: "right", indent: 1 };
    ws.getRow(1).height = 36;

    // 2) سطر المعلومات
    ws.mergeCells(2, 1, 2, N);
    const info = ws.getCell(2, 1);
    const now = new Date();
    const stampTxt = now.toLocaleString("ar-SA-u-ca-gregory-nu-latn", { dateStyle: "medium", timeStyle: "short" });
    info.value = `تاريخ التصدير: ${stampTxt}     عدد المشاريع: ${data.length} من ${rows.length}     ${activeFiltersText()}`;
    info.font = font({ size: 10, color: { argb: XL.text2 } });
    info.fill = fill(XL.primary25);
    info.alignment = { vertical: "middle", horizontal: "right", indent: 1 };
    info.border = { bottom: thin() };
    ws.getRow(2).height = 22;
    ws.getRow(3).height = 8;

    // 3) بطاقات المؤشرات (صفوف 4–6)
    const revenue = sum(data, "revenue"), plan = sum(data, "totalPlan"), toNov = sum(data, "plannedToNov");
    const delivered = sum(data, "delivered"), rem = sum(data, "remaining");
    const overdue = data.filter((r) => dueState(r) === "overdue").length;
    const soon = data.filter((r) => dueState(r) === "soon").length;
    const noContract = data.filter((r) => r.planNote).length;
    const kpis = [
      { label: "الإيراد - المالية", value: revenue, fmt: '#,##0 "ر.س"', sub: `${millions(revenue)} ريال`, color: XL.accent },
      { label: "اجمالي الخطط للتسليم", value: plan, fmt: '#,##0 "ر.س"', sub: `${millions(plan)} ريال`, color: XL.primary },
      { label: "المسلَّم", value: delivered, fmt: '#,##0 "ر.س"', sub: `${plan > 0 ? ((delivered / plan) * 100).toFixed(1) : "0.0"}% من إجمالي الخطط`, color: XL.success },
      { label: "التسليم مقابل المخطط حتى نوفمبر", value: toNov > 0 ? delivered / toNov : 0, fmt: "0.0%", sub: `${millions(delivered)} من ${millions(toNov)}`, color: "FF2E7DB8" },
      { label: "المتبقي", value: rem, fmt: '#,##0 "ر.س"', sub: `${millions(rem)} ريال`, color: XL.warning },
      { label: "المشاريع", value: data.length, fmt: "0", sub: `${noContract} لم يصل العقد`, color: XL.primary600 },
      { label: "الاستحقاقات المتأخرة", value: overdue, fmt: "0", sub: `${soon} تستحق خلال 7 أيام`, color: XL.danger },
    ];
    // توزيع الأعمدة على البطاقات بحسب العرض لتتقارب أحجامها
    const widths = cols.map((c) => c.width || 14);
    const totalW = widths.reduce((a, b) => a + b, 0);
    let start = 1;
    kpis.forEach((k, i) => {
      let end = start, acc = widths[start - 1];
      const target = (totalW / kpis.length) * (i + 1);
      const before = widths.slice(0, start - 1).reduce((a, b) => a + b, 0);
      while (end < N - (kpis.length - 1 - i) && before + acc + widths[end] / 2 < target) { acc += widths[end]; end++; }
      if (i === kpis.length - 1) end = N;
      [[4, k.label, font({ size: 9.5, color: { argb: XL.text2 } })],
       [5, k.value, font({ size: 15, bold: true, color: { argb: XL.primary } })],
       [6, k.sub, font({ size: 9, color: { argb: XL.muted } })]].forEach(([r, v, f]) => {
        ws.mergeCells(r, start, r, end);
        const c = ws.getCell(r, start);
        c.value = v;
        c.font = f;
        c.alignment = { vertical: "middle", horizontal: "right", indent: 1 };
        if (r === 5) c.numFmt = k.fmt;
        for (let x = start; x <= end; x++) {
          const cell = ws.getCell(r, x);
          cell.fill = fill(XL.white);
          cell.border = {
            top: r === 4 ? { style: "thick", color: { argb: k.color } } : undefined,
            bottom: r === 6 ? thin() : undefined,
            left: x === end ? thin() : undefined,
            right: x === start ? thin() : undefined,
          };
        }
      });
      start = end + 1;
    });
    ws.getRow(4).height = 20; ws.getRow(5).height = 26; ws.getRow(6).height = 18; ws.getRow(7).height = 10;

    // 4) رأس الجدول
    const head = ws.getRow(HEADER_ROW);
    cols.forEach((c, i) => {
      const cell = head.getCell(i + 1);
      cell.value = c.label;
      cell.font = font({ size: 10.5, bold: true, color: { argb: XL.white } });
      cell.fill = fill(XL.primary);
      cell.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
      cell.border = { left: thin("FF2B4A72"), right: thin("FF2B4A72"), bottom: { style: "medium", color: { argb: XL.primary } } };
    });
    head.height = 32;

    // 5) الصفوف
    const LAV = ["segment", "entity", "portfolio", "project"];
    data.forEach((r, i) => {
      const row = ws.getRow(HEADER_ROW + 1 + i);
      const base = i % 2 ? XL.zebra : XL.white;
      const ok = deliveredOk(r);
      const ds = dueState(r);
      cols.forEach((c, ci) => {
        const cell = row.getCell(ci + 1);
        let value = null;
        let f = font();
        let bg = base;
        let align = { vertical: "top", horizontal: "right", wrapText: true };
        if (LAV.includes(c.key)) bg = XL.lavender;
        if (c.key === "_num") { value = i + 1; f = font({ size: 9, color: { argb: XL.muted } }); align = { vertical: "top", horizontal: "center" }; }
        else if (c.money) {
          value = c.key === "remaining" ? remaining(r) : r[c.key];
          align = { vertical: "top", horizontal: "left" };
          cell.numFmt = XL_MONEY;
          if (c.key === "totalPlan" && !isNum(r.totalPlan) && r.planNote) {
            value = r.planNote; bg = XL.warningBg; f = font({ bold: true, color: { argb: XL.warning } }); align = { vertical: "top", horizontal: "center" };
          }
          if (c.key === "revenue") { bg = XL.primary50; f = font({ bold: true, color: { argb: XL.accent } }); }
          if ((c.key === "plannedToNov" || c.key === "delivered") && ok) { bg = XL.successBg; f = font({ bold: true, color: { argb: XL.success } }); }
          if (isNum(value) && value === 0) f = font({ color: { argb: XL.muted } });
          if (isNum(value) && value < 0) f = font({ color: { argb: XL.danger } });
        } else if (c.key === "dueDate") {
          align = { vertical: "top", horizontal: "center" };
          if (r.dueDate) {
            const [y, m, d] = r.dueDate.split("-").map(Number);
            value = new Date(Date.UTC(y, m - 1, d));
            cell.numFmt = "dd/mm/yyyy";
          } else value = r.dueRaw || null;
          if (ds === "overdue") { bg = XL.dangerBg; f = font({ bold: true, color: { argb: XL.danger } }); }
          else if (ds === "soon") { bg = XL.warningBg; f = font({ bold: true, color: { argb: XL.warning } }); }
          else if (ds === "invalid") f = font({ color: { argb: XL.danger }, italic: true });
        } else {
          value = r[c.key] || null;
          switch (c.key) {
            case "projectNo": f = font({ bold: true, color: { argb: XL.primary } }); align = { vertical: "top", horizontal: "center" }; break;
            case "status":
              align = { vertical: "top", horizontal: "center" };
              if (value === "قائم") { bg = XL.primary; f = font({ bold: true, color: { argb: XL.white } }); }
              else if (value) { bg = XL.newBg; f = font({ bold: true, color: { argb: XL.accent } }); }
              else { value = "غير محدد"; bg = XL.primary50; f = font({ color: { argb: XL.muted } }); }
              break;
            case "segment": f = font({ bold: true, color: { argb: XL.primary } }); align = { vertical: "top", horizontal: "center" }; break;
            case "entity":
              align = { vertical: "top", horizontal: "center" };
              if (value === "MOE") { bg = XL.moeBg; f = font({ bold: true, color: { argb: XL.moe } }); }
              else if (value) { bg = XL.nonMoeBg; f = font({ bold: true, color: { argb: XL.nonMoe } }); }
              break;
            case "portfolio": f = font({ color: { argb: XL.text2 } }); if (!value) { value = "غير محدد"; f = font({ color: { argb: XL.muted } }); } break;
            case "project": f = font({ bold: true, color: { argb: XL.accent } }); break;
            case "projectManager": f = font({ bold: true, color: { argb: XL.primary } }); if (!value) { value = "غير محدد"; f = font({ color: { argb: XL.muted } }); } break;
            default: f = font({ color: { argb: XL.text2 } });
          }
        }
        cell.value = value;
        cell.font = f;
        cell.fill = fill(bg);
        cell.alignment = Object.assign({ indent: align.horizontal === "center" ? 0 : 1 }, align);
        cell.border = { bottom: thin(), left: thin("FFEEF1F5"), right: thin("FFEEF1F5") };
      });
      // Excel لا يضبط ارتفاع الصفوف ذات النص الطويل تلقائياً، فنقدّره من طول النص وعرض العمود
      let lines = 1;
      cols.forEach((c, ci) => {
        if (c.money || c.key === "_num" || c.key === "dueDate") return;
        const text = String(row.getCell(ci + 1).value || "");
        if (!text) return;
        const perLine = Math.max(4, (c.width || 14) * 1.35 - 2);
        const l = text.split("\n").reduce((acc, part) => acc + Math.max(1, Math.ceil(part.length / perLine)), 0);
        lines = Math.max(lines, l);
      });
      row.height = Math.max(22, lines * 14 + 8);
    });

    // 6) سطر الإجمالي
    const totalRowNo = HEADER_ROW + data.length + 1;
    const totalRow = ws.getRow(totalRowNo);
    const firstMoney = cols.findIndex((c) => c.money);
    ws.mergeCells(totalRowNo, 1, totalRowNo, firstMoney);
    cols.forEach((c, ci) => {
      const cell = totalRow.getCell(ci + 1);
      if (ci === 0) cell.value = "الإجمالي";
      else if (c.money) { cell.value = sum(data, c.key); cell.numFmt = XL_MONEY; }
      cell.font = font({ size: 11, bold: true, color: { argb: XL.primary } });
      cell.fill = fill(XL.primary50);
      cell.alignment = { vertical: "middle", horizontal: c.money ? "left" : "right", indent: c.money ? 0 : 1 };
      cell.border = { top: { style: "medium", color: { argb: XL.primary } }, bottom: { style: "medium", color: { argb: XL.primary } } };
    });
    totalRow.height = 26;

    // 7) دليل الألوان
    const legendRowNo = totalRowNo + 2;
    const legend = [
      ["تم تسليم المخطط", XL.successBg, XL.success],
      ["متأخر", XL.dangerBg, XL.danger],
      ["يستحق خلال 7 أيام", XL.warningBg, XL.warning],
      ["لم يصل العقد", XL.warningBg, XL.warning],
    ];
    const lc = ws.getCell(legendRowNo, 2);
    lc.value = "دليل الألوان:";
    lc.font = font({ bold: true, color: { argb: XL.text2 } });
    legend.forEach(([label, bg, fg], i) => {
      const cell = ws.getCell(legendRowNo, 3 + i);
      cell.value = label;
      cell.font = font({ size: 9, bold: true, color: { argb: fg } });
      cell.fill = fill(bg);
      cell.alignment = { horizontal: "center", vertical: "middle" };
      cell.border = { top: thin(fg), bottom: thin(fg), left: thin(fg), right: thin(fg) };
    });

    ws.autoFilter = { from: { row: HEADER_ROW, column: 1 }, to: { row: Math.max(HEADER_ROW, totalRowNo - 1), column: N } };
    return wb;
  }

  async function exportExcel() {
    if (activeEdit) commitEdit();
    if (!window.ExcelJS) return toast("مكتبة Excel لم تُحمَّل بعد، حاول مرة أخرى");
    const data = filtered();
    const btn = $("#exportBtn");
    btn.disabled = true;
    try {
      const wb = await buildStyledWorkbook(data);
      const buffer = await wb.xlsx.writeBuffer();
      const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      const stamp = new Date().toISOString().slice(0, 10);
      const filename = `التعافي_المالي_${stamp}.xlsx`;
      const label = data.length === rows.length ? `${data.length} مشروع` : `${data.length} مشروع (حسب الفلاتر)`;
      if (!inViewer) {
        const url = URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = url; a.download = filename;
        document.body.appendChild(a); a.click(); a.remove();
        setTimeout(() => URL.revokeObjectURL(url), 2000);
        return toast(`تم تصدير ${label}`);
      }
      if (!downloadsApi) return toast("التصدير غير متاح في هذا العرض");
      await downloadsApi.save({ filename, data: blob });
      toast(`تم تصدير ${label}`);
    } catch (err) {
      const code = err && err.code;
      if (code === "declined") toast("تم إلغاء التصدير");
      else if (code === "rate_limited") toast("هناك نافذة حفظ مفتوحة بالفعل، حاول بعد قليل");
      else if (code === "host_unavailable") toast("هذا التطبيق لا يدعم حفظ الملفات، افتح الصفحة من المتصفح");
      else { console.error(err); toast("تعذر التصدير"); }
    } finally {
      btn.disabled = false;
    }
  }

  function importExcel(file) {
    if (!window.XLSX) return toast("مكتبة Excel غير متاحة");
    const reader = new FileReader();
    reader.onload = async (e) => {
      try {
        const wb = XLSX.read(e.target.result, { type: "array", cellDates: true });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
        // صف العناوين قد لا يكون الأول (الملف المصدَّر من النظام يبدأ بعنوان ومؤشرات)
        let headerAt = aoa.slice(0, 30).findIndex((line) => (line || []).some((h) => String(h || "").trim() === "Project Name"));
        if (headerAt < 0) headerAt = 0;
        const header = (aoa[headerAt] || []).map((h) => String(h || "").trim());
        const idx = {};
        COLUMNS.forEach((c) => { idx[c.key] = header.indexOf(c.label); });
        if (idx.project < 0) throw new Error("لم يتم العثور على عمود Project Name");
        const out = [];
        aoa.slice(headerAt + 1).forEach((line) => {
          const get = (k) => (idx[k] >= 0 ? line[idx[k]] : null);
          const project = String(get("project") || "").trim();
          if (!project) return; // يتجاوز صف الإجمالي والصفوف الفارغة
          const r = { id: out.length + 1 };
          TEXT_KEYS.forEach((k) => { r[k] = String(get(k) == null ? "" : get(k)).trim(); });
          MONEY_KEYS.forEach((k) => {
            const v = get(k);
            r[k] = typeof v === "number" ? round2(v) : null;
            if (k === "totalPlan") r.planNote = typeof v === "string" ? v.trim() : "";
          });
          const d = get("dueDate");
          r.dueDate = null; r.dueRaw = "";
          if (d instanceof Date && !isNaN(d)) {
            const local = new Date(d.getTime() - d.getTimezoneOffset() * 60000);
            r.dueDate = local.toISOString().slice(0, 10);
          } else if (d) r.dueRaw = String(d);
          out.push(r);
        });
        if (!out.length) throw new Error("الملف لا يحتوي على مشاريع");
        if (!(await askConfirm(`سيتم استبدال البيانات الحالية بـ ${out.length} مشروع من الملف. متابعة؟`, "استيراد"))) return;
        rows = migrate(out);
        save(); refreshFilterOptions(); render();
        toast(`تم استيراد ${out.length} مشروع`);
      } catch (err) {
        toast("تعذر قراءة الملف: " + err.message);
      }
    };
    reader.readAsArrayBuffer(file);
  }

  // ---------- نافذة التأكيد ----------
  function askConfirm(message, okLabel) {
    const d = $("#confirmDialog");
    $("#confirmMsg").textContent = message;
    $("#confirmYes").textContent = okLabel || "تأكيد";
    return new Promise((resolve) => {
      const done = (v) => { d.removeEventListener("close", onClose); if (d.open) d.close(); resolve(v); };
      const onClose = () => done(false);
      $("#confirmYes").onclick = () => done(true);
      $("#confirmNo").onclick = () => done(false);
      d.addEventListener("close", onClose);
      d.showModal();
    });
  }

  // ---------- تنبيه ----------
  let toastTimer;
  function toast(msg) {
    const t = $("#toast");
    t.textContent = msg;
    t.classList.add("show");
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => t.classList.remove("show"), 2200);
  }

  // ---------- الأحداث ----------
  FILTERS.forEach((s) => $(s).addEventListener("input", render));
  $("#clearFilters").addEventListener("click", () => {
    FILTERS.forEach((s) => { $(s).value = ""; });
    sort = { key: null, dir: 1 };
    render();
  });
  // الترتيب: تصاعدي ← تنازلي ← بدون ترتيب
  $("#gridHead").addEventListener("click", (e) => {
    const b = e.target.closest("[data-sort]");
    if (!b) return;
    const k = b.dataset.sort;
    if (sort.key === k) {
      if (sort.dir === 1) sort.dir = -1; else sort = { key: null, dir: 1 };
    } else sort = { key: k, dir: 1 };
    render();
    const again = $(`#gridHead [data-sort="${k}"]`);
    if (again) again.focus();
  });
  $("#gridBody").addEventListener("click", (e) => {
    const open = e.target.closest("[data-open]");
    if (open) return openEditor(Number(open.dataset.open));
    const td = e.target.closest("td[data-field]");
    if (td && !td.classList.contains("editing")) startEdit(td);
  });
  // Enter على خانة محددة يبدأ تعديلها
  $("#gridBody").addEventListener("keydown", (e) => {
    const td = e.target.closest && e.target.closest("td[data-field]");
    if (td && e.target === td && e.key === "Enter") { e.preventDefault(); startEdit(td); }
  });
  $("#addBtn").addEventListener("click", () => openEditor(null));
  $("#exportBtn").addEventListener("click", exportExcel);
  $("#importFile").addEventListener("change", (e) => {
    if (e.target.files[0]) importExcel(e.target.files[0]);
    e.target.value = "";
  });

  refreshFilterOptions();
  renderLastUpdate();
  render();
})();
