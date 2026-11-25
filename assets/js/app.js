/* التعافي المالي — خطة الإيراد والتسليم */
(function () {
  "use strict";

  const STORAGE_KEY = "financialRecovery.rows.v1";
  const META_KEY = "financialRecovery.meta.v1";

  // الأعمدة بنفس ترتيب ملف Excel
  const COLUMNS = [
    { key: "status", label: "حالة العقد" },
    { key: "segment", label: "Segment" },
    { key: "entity", label: "الجهة" },
    { key: "project", label: "Project Name" },
    { key: "revenue", label: "الإيراد - المالية", money: true },
    { key: "totalPlan", label: "اجمالي الخطط للتسليم", money: true },
    { key: "plannedToNov", label: "التسليم المخطط حتى نوفمبر", money: true },
    { key: "plannedDec", label: "التسليم المخطط ديسمبر", money: true },
    { key: "delivered", label: "مسلم", money: true },
    { key: "remaining", label: "المتبقي", money: true, computed: true },
    { key: "action", label: "الإجراء المطلوب" },
    { key: "owner", label: "المسؤول" },
    { key: "nextStep", label: "الخطوة القادمة" },
    { key: "dueDate", label: "تاريخ الاستحقاق" },
  ];
  const MONEY_KEYS = ["revenue", "totalPlan", "plannedToNov", "plannedDec", "delivered"];

  // ---------- الحالة ----------
  let rows = load();
  let sort = { key: null, dir: 1 };
  let editingId = null;
  const $ = (sel) => document.querySelector(sel);

  function load() {
    try {
      const saved = JSON.parse(localStorage.getItem(STORAGE_KEY));
      if (Array.isArray(saved)) return saved;
    } catch (e) { /* تجاهل */ }
    return JSON.parse(JSON.stringify(window.SEED_DATA || []));
  }
  function save() {
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(rows));
      localStorage.setItem(META_KEY, JSON.stringify({ updatedAt: new Date().toISOString() }));
    } catch (e) { toast("تعذر الحفظ في المتصفح"); }
    renderLastUpdate();
  }

  // ---------- الحسابات ----------
  const isNum = (v) => typeof v === "number" && isFinite(v);
  const n = (v) => (isNum(v) ? v : 0);
  function remaining(r) {
    return isNum(r.totalPlan) ? round2(r.totalPlan - n(r.delivered)) : null;
  }
  const round2 = (v) => Math.round(v * 100) / 100;
  function deliveredOk(r) {
    return isNum(r.delivered) && r.delivered > 0 && isNum(r.plannedToNov) && r.delivered >= r.plannedToNov;
  }
  function today() {
    const d = new Date(); d.setHours(0, 0, 0, 0); return d;
  }
  function dueState(r) {
    if (!r.dueDate) return r.dueRaw ? "invalid" : "none";
    const d = new Date(r.dueDate + "T00:00:00");
    const diff = Math.round((d - today()) / 86400000);
    if (diff < 0) return "overdue";
    if (diff <= 7) return "soon";
    return "later";
  }
  function daysLabel(r) {
    const d = new Date(r.dueDate + "T00:00:00");
    const diff = Math.round((d - today()) / 86400000);
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

  // ---------- الفلاتر ----------
  function filtered() {
    const q = $("#searchInput").value.trim().toLowerCase();
    const fs = $("#fStatus").value, fg = $("#fSegment").value, fe = $("#fEntity").value, fd = $("#fDue").value;
    let out = rows.filter((r) => {
      if (fs && r.status !== fs) return false;
      if (fg && r.segment !== fg) return false;
      if (fe && r.entity !== fe) return false;
      if (fd) {
        const s = dueState(r);
        if (fd === "none" ? !(s === "none" || s === "invalid") : s !== fd) return false;
      }
      if (q) {
        const hay = [r.project, r.owner, r.action, r.nextStep, r.segment, r.entity, r.status, r.planNote].join(" ").toLowerCase();
        if (!hay.includes(q)) return false;
      }
      return true;
    });
    if (sort.key) {
      const k = sort.key;
      const val = (r) => (k === "remaining" ? remaining(r) : r[k]);
      out = out.slice().sort((a, b) => {
        const va = val(a), vb = val(b);
        const ea = va == null || va === "", eb = vb == null || vb === "";
        if (ea && eb) return 0;
        if (ea) return 1;
        if (eb) return -1;
        if (isNum(va) && isNum(vb)) return (va - vb) * sort.dir;
        return String(va).localeCompare(String(vb), "ar") * sort.dir;
      });
    }
    return out;
  }

  function fillSelect(sel, values, placeholder) {
    const cur = sel.value;
    sel.innerHTML = `<option value="">${placeholder}</option>` +
      values.map((v) => `<option value="${esc(v)}">${esc(v)}</option>`).join("");
    if (values.includes(cur)) sel.value = cur;
  }
  const uniq = (key) => [...new Set(rows.map((r) => r[key]).filter(Boolean))].sort((a, b) => a.localeCompare(b, "ar"));

  function refreshFilterOptions() {
    fillSelect($("#fStatus"), uniq("status"), "حالة العقد: الكل");
    fillSelect($("#fSegment"), uniq("segment"), "Segment: الكل");
    fillSelect($("#fEntity"), uniq("entity"), "الجهة: الكل");
    $("#segmentList").innerHTML = uniq("segment").map((v) => `<option value="${esc(v)}">`).join("");
    $("#entityList").innerHTML = uniq("entity").map((v) => `<option value="${esc(v)}">`).join("");
  }

  // ---------- العرض ----------
  function render() {
    const data = filtered();
    renderKpis(data);
    renderTable(data);
    $("#rowCount").textContent = data.length === rows.length ? `${rows.length} مشروع` : `${data.length} من ${rows.length}`;
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

  function badgeStatus(s) {
    if (!s) return `<span class="badge badge-muted">غير محدد</span>`;
    return `<span class="badge ${s === "قائم" ? "badge-active" : "badge-new"}">${esc(s)}</span>`;
  }
  function badgeEntity(s) {
    if (!s) return "";
    return `<span class="badge ${s === "MOE" ? "badge-moe" : "badge-nonmoe"}">${esc(s)}</span>`;
  }

  function moneyCell(v, extra = "") {
    const cls = ["money", extra];
    if (isNum(v) && v === 0) cls.push("zero");
    if (isNum(v) && v < 0) cls.push("neg");
    return `<td class="${cls.join(" ").trim()}">${money(v)}</td>`;
  }

  function renderTable(data) {
    const body = data.map((r, i) => {
      const rem = remaining(r);
      const ok = deliveredOk(r);
      const ds = dueState(r);
      const pct = isNum(r.totalPlan) && r.totalPlan > 0 ? Math.min(100, (n(r.delivered) / r.totalPlan) * 100) : null;

      let due = "";
      if (ds === "invalid") due = `<span class="due invalid" title="تاريخ غير صالح في الملف الأصلي"><span class="due-date">${esc(r.dueRaw)}</span><small>تاريخ غير صالح</small></span>`;
      else if (r.dueDate) due = `<span class="due ${ds}"><span class="due-date">${fmtDate(r.dueDate)}</span><small>${daysLabel(r)}</small></span>`;

      const planCell = isNum(r.totalPlan)
        ? moneyCell(r.totalPlan)
        : `<td class="money">${r.planNote ? `<span class="badge badge-note">${esc(r.planNote)}</span>` : ""}</td>`;

      return `<tr data-id="${r.id}">
        <td class="col-num">${i + 1}</td>
        <td class="cell-tinted">${badgeStatus(r.status)}</td>
        <td class="cell-lav"><span class="badge badge-seg">${esc(r.segment)}</span></td>
        <td class="cell-lav">${badgeEntity(r.entity)}</td>
        <td class="col-project cell-lav">${esc(r.project)}</td>
        ${moneyCell(r.revenue, "revenue cell-tinted")}
        ${planCell}
        ${moneyCell(r.plannedToNov, ok ? "ok" : "")}
        ${moneyCell(r.plannedDec)}
        ${moneyCell(r.delivered, ok ? "ok" : "")}
        <td class="money${isNum(rem) && rem < 0 ? " neg" : ""}${isNum(rem) && rem === 0 ? " zero" : ""}">${money(rem)}${pct != null ? `<div class="mini-bar" title="نسبة التسليم ${pct.toFixed(1)}%"><span style="width:${pct.toFixed(1)}%"></span></div>` : ""}</td>
        <td class="col-text">${esc(r.action)}</td>
        <td class="col-text">${esc(r.owner)}</td>
        <td class="col-text">${esc(r.nextStep)}</td>
        <td class="${ds === "overdue" ? "due-overdue" : ds === "soon" ? "due-soon" : ""}">${due}</td>
        <td class="col-actions"><button class="icon-btn" data-edit="${r.id}" title="تعديل" aria-label="تعديل">
          <svg viewBox="0 0 24 24" width="16" height="16"><path d="M4 20h4L19 9l-4-4L4 16v4z" fill="none" stroke="currentColor" stroke-width="2" stroke-linejoin="round"/></svg>
        </button></td>
      </tr>`;
    }).join("");
    $("#gridBody").innerHTML = body;
    $("#emptyState").hidden = data.length > 0;

    $("#gridFoot").innerHTML = data.length ? `<tr>
      <td></td><td colspan="4">الإجمالي</td>
      ${["revenue", "totalPlan", "plannedToNov", "plannedDec", "delivered", "remaining"].map((k) => `<td class="money">${money(sum(data, k))}</td>`).join("")}
      <td colspan="5"></td></tr>` : "";

    document.querySelectorAll(".grid thead th[data-sort]").forEach((th) => {
      th.classList.toggle("sorted", th.dataset.sort === sort.key);
      th.classList.toggle("desc", th.dataset.sort === sort.key && sort.dir === -1);
    });
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

  // ---------- الإضافة / التعديل ----------
  const dlg = $("#editDialog");
  const form = $("#editForm");

  function openEditor(id) {
    editingId = id;
    const r = id != null ? rows.find((x) => x.id === id) : {};
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
    if (isNum(data.totalPlan)) data.planNote = data.planNote || "";
    if (editingId != null) {
      const r = rows.find((x) => x.id === editingId);
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
    const r = rows.find((x) => x.id === editingId);
    if (!r || !(await askConfirm(`حذف المشروع "${r.project}"؟`, "حذف"))) return;
    rows = rows.filter((x) => x.id !== editingId);
    save(); refreshFilterOptions(); render();
    dlg.close();
    toast("تم حذف المشروع");
  });
  dlg.querySelectorAll("[data-close]").forEach((b) => b.addEventListener("click", () => dlg.close()));

  // ---------- Excel ----------
  function exportExcel() {
    if (!window.XLSX) return toast("مكتبة Excel غير متاحة");
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
    ws["!cols"] = [10, 12, 12, 45, 18, 18, 18, 18, 18, 18, 38, 35, 35, 14].map((w) => ({ wch: w }));
    const range = XLSX.utils.decode_range(ws["!ref"]);
    for (let R = 1; R <= range.e.r; R++) {
      for (let C = 4; C <= 9; C++) {
        const cell = ws[XLSX.utils.encode_cell({ r: R, c: C })];
        if (cell && cell.t === "n") cell.z = "#,##0.00";
      }
      const dc = ws[XLSX.utils.encode_cell({ r: R, c: 13 })];
      if (dc && dc.t === "d") dc.z = "yyyy-mm-dd";
    }
    const wb = XLSX.utils.book_new();
    wb.Workbook = { Views: [{ RTL: true }] };
    XLSX.utils.book_append_sheet(wb, ws, "خطة الايراد والتسليم");
    const stamp = new Date().toISOString().slice(0, 10);
    XLSX.writeFile(wb, `التعافي_المالي_${stamp}.xlsx`);
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
          ["status", "segment", "entity", "action", "owner", "nextStep"].forEach((k) => { r[k] = String(get(k) == null ? "" : get(k)).trim(); });
          r.project = project;
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
        rows = out;
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
    toastTimer = setTimeout(() => t.classList.remove("show"), 2800);
  }

  // ---------- الأحداث ----------
  ["#searchInput", "#fStatus", "#fSegment", "#fEntity", "#fDue"].forEach((s) =>
    $(s).addEventListener("input", render));
  $("#clearFilters").addEventListener("click", () => {
    ["#searchInput", "#fStatus", "#fSegment", "#fEntity", "#fDue"].forEach((s) => { $(s).value = ""; });
    sort = { key: null, dir: 1 };
    render();
  });
  document.querySelectorAll(".grid thead th[data-sort]").forEach((th) =>
    th.addEventListener("click", () => {
      const k = th.dataset.sort;
      if (sort.key === k) {
        if (sort.dir === 1) sort.dir = -1; else sort = { key: null, dir: 1 };
      } else sort = { key: k, dir: 1 };
      render();
    }));
  $("#gridBody").addEventListener("click", (e) => {
    const b = e.target.closest("[data-edit]");
    if (b) return openEditor(Number(b.dataset.edit));
  });
  $("#gridBody").addEventListener("dblclick", (e) => {
    const tr = e.target.closest("tr[data-id]");
    if (tr) openEditor(Number(tr.dataset.id));
  });
  $("#addBtn").addEventListener("click", () => openEditor(null));
  if ($("#exportBtn")) $("#exportBtn").addEventListener("click", exportExcel);
  $("#importFile").addEventListener("change", (e) => {
    if (e.target.files[0]) importExcel(e.target.files[0]);
    e.target.value = "";
  });

  refreshFilterOptions();
  renderLastUpdate();
  render();
})();
