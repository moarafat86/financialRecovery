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

  async function exportExcel() {
    if (activeEdit) commitEdit();
    if (!window.XLSX) return toast("مكتبة Excel لم تُحمَّل بعد، حاول مرة أخرى");
    const data = filtered();
    const aoa = [COLUMNS.map((c) => c.label)];
    data.forEach((r) => {
      aoa.push(COLUMNS.map((c) => {
        if (c.key === "remaining") return remaining(r);
        if (c.key === "totalPlan") return isNum(r.totalPlan) ? r.totalPlan : (r.planNote || null);
        if (c.key === "dueDate") return r.dueDate ? new Date(r.dueDate + "T00:00:00") : (r.dueRaw || null);
        return r[c.key] == null || r[c.key] === "" ? null : r[c.key];
      }));
    });
    aoa.push(COLUMNS.map((c, i) => (i === 0 ? "الإجمالي" : c.money ? sum(data, c.key) : null)));
    const ws = XLSX.utils.aoa_to_sheet(aoa, { cellDates: true });
    ws["!cols"] = COLUMNS.map((c) => ({ wch: c.width || 14 }));
    const range = XLSX.utils.decode_range(ws["!ref"]);
    const dateCol = COLUMNS.findIndex((c) => c.key === "dueDate");
    for (let R = 1; R <= range.e.r; R++) {
      COLUMNS.forEach((c, C) => {
        if (!c.money) return;
        const cell = ws[XLSX.utils.encode_cell({ r: R, c: C })];
        if (cell && cell.t === "n") cell.z = "#,##0.00";
      });
      const dc = ws[XLSX.utils.encode_cell({ r: R, c: dateCol })];
      if (dc && dc.t === "d") dc.z = "yyyy-mm-dd";
    }
    const wb = XLSX.utils.book_new();
    wb.Workbook = { Views: [{ RTL: true }] };
    XLSX.utils.book_append_sheet(wb, ws, "خطة الايراد والتسليم");
    const stamp = new Date().toISOString().slice(0, 10);
    const filename = `التعافي_المالي_${stamp}.xlsx`;
    const label = data.length === rows.length ? `${data.length} مشروع` : `${data.length} مشروع (حسب الفلاتر)`;
    if (!inViewer) {
      XLSX.writeFile(wb, filename);
      return toast(`تم تصدير ${label}`);
    }
    if (!downloadsApi) return toast("التصدير غير متاح في هذا العرض");
    const bytes = XLSX.write(wb, { bookType: "xlsx", type: "array" });
    const btn = $("#exportBtn");
    btn.disabled = true;
    try {
      await downloadsApi.save({ filename, data: new Blob([bytes]) });
      toast(`تم تصدير ${label}`);
    } catch (err) {
      const code = err && err.code;
      if (code === "declined") toast("تم إلغاء التصدير");
      else if (code === "rate_limited") toast("هناك نافذة حفظ مفتوحة بالفعل، حاول بعد قليل");
      else if (code === "host_unavailable") toast("هذا التطبيق لا يدعم حفظ الملفات، افتح الصفحة من المتصفح");
      else toast("تعذر التصدير في هذا العرض");
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
        const header = (aoa[0] || []).map((h) => String(h || "").trim());
        const idx = {};
        COLUMNS.forEach((c) => { idx[c.key] = header.indexOf(c.label); });
        if (idx.project < 0) throw new Error("لم يتم العثور على عمود Project Name");
        const out = [];
        aoa.slice(1).forEach((line) => {
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
