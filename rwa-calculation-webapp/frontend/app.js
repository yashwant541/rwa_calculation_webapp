/* RWA Calculation webapp - front end (vanilla JS, no build step). */
(function () {
  "use strict";

  function apiBase() {
    try { if (typeof getWebAppBackendUrl === "function") return getWebAppBackendUrl(""); }
    catch (e) { /* not in Dataiku */ }
    return "";
  }
  const API = apiBase().replace(/\/$/, "");

  let ADMIN_TOKEN = null;
  try { ADMIN_TOKEN = sessionStorage.getItem("rwa_admin_token"); } catch (e) {}

  async function call(path, opts) {
    opts = opts || {};
    const headers = opts.headers || {};
    if (ADMIN_TOKEN && path.indexOf("/api/admin") === 0) headers["X-Admin-Token"] = ADMIN_TOKEN;
    if (opts.json !== undefined) { headers["Content-Type"] = "application/json"; opts.body = JSON.stringify(opts.json); }
    const res = await fetch(API + path, { method: opts.method || "GET", headers, body: opts.body });
    let data = null;
    try { data = await res.json(); } catch (e) {}
    if (!res.ok || (data && data.success === false)) {
      let msg = (data && (data.message || data.error)) || ("HTTP " + res.status);
      if (data && data.details) msg += " — " + (typeof data.details === "string" ? data.details : JSON.stringify(data.details));
      if (data && Array.isArray(data.errors) && data.errors.length) msg = data.errors.join("; ");
      const err = new Error(msg); err.data = data; throw err;
    }
    return data;
  }
  const D = (r) => (r && r.data !== undefined ? r.data : r);

  // ---------------------------------------------------------------- dom helpers
  const $ = (s, r) => (r || document).querySelector(s);
  const $$ = (s, r) => Array.prototype.slice.call((r || document).querySelectorAll(s));
  const SVGNS = "http://www.w3.org/2000/svg";
  const el = (tag, attrs, kids) => {
    const svg = ["svg", "path", "rect", "g", "text", "line"].indexOf(tag) >= 0;
    const n = svg ? document.createElementNS(SVGNS, tag) : document.createElement(tag);
    Object.entries(attrs || {}).forEach(([k, v]) => {
      if (v == null) return;
      if (k === "class") n.setAttribute("class", v);
      else if (k === "text") n.textContent = v;
      else if (k === "html") n.innerHTML = v;
      else if (k.slice(0, 2) === "on") n.addEventListener(k.slice(2), v);
      else n.setAttribute(k, v);
    });
    (kids || []).forEach((c) => c != null && n.appendChild(typeof c === "string" ? document.createTextNode(c) : c));
    return n;
  };
  function toast(msg, kind) {
    const t = el("div", { class: "toast " + (kind || ""), text: msg });
    $("#toast-wrap").appendChild(t);
    setTimeout(() => t.remove(), 4200);
  }
  function confirmModal(title, body) {
    return new Promise((resolve) => {
      $("#modal-title").textContent = title;
      $("#modal-body").textContent = body;
      $("#modal").classList.remove("hidden");
      const done = (v) => { $("#modal").classList.add("hidden"); resolve(v); };
      const ok = $("#modal-ok").cloneNode(true), cancel = $("#modal-cancel").cloneNode(true);
      $("#modal-ok").replaceWith(ok); $("#modal-cancel").replaceWith(cancel);
      ok.addEventListener("click", () => done(true));
      cancel.addEventListener("click", () => done(false));
    });
  }
  function show(view) {
    $$(".view").forEach((v) => v.classList.add("hidden"));
    $("#view-" + view).classList.remove("hidden");
    window.scrollTo(0, 0);
  }
  const fmt = (v) => (v === null || v === undefined) ? "∅" : (typeof v === "number" ? trimNum(v) : String(v));
  function trimNum(n) {
    if (typeof n !== "number") return String(n);
    if (!isFinite(n)) return n > 0 ? "∞" : "-∞";
    return Number.isInteger(n) ? String(n) : String(Math.round(n * 1e10) / 1e10);
  }
  const bytes = (n) => n == null ? "" : n < 1024 ? n + " B" : n < 1048576 ? (n / 1024).toFixed(0) + " KB" : (n / 1048576).toFixed(1) + " MB";
  const when = (s) => s ? String(s).replace("T", " ").slice(0, 16) : "";

  // ---------------------------------------------------------------- router
  function go(hash) { if (location.hash === hash) route(); else location.hash = hash; }
  function route() {
    closeDrawer();
    const raw = (location.hash || "#/").replace(/^#\/?/, "");
    const [pathPart, queryPart] = raw.split("?");
    const parts = pathPart.split("/").filter(Boolean);
    const q = {};
    (queryPart || "").split("&").forEach((kv) => { const [k, v] = kv.split("="); if (k) q[k] = decodeURIComponent(v || ""); });
    if (parts[0] === "c" && parts[1] && parts[2] === "trace") return openTracePage(decodeURIComponent(parts[1]), q);
    if (parts[0] === "c" && parts[1]) return openCategory(decodeURIComponent(parts[1]));
    if (parts[0] === "admin") return openAdmin(parts[1] ? decodeURIComponent(parts[1]) : null);
    return loadHome();
  }

  const ICONS = { landmark: "\u{1F3DB}", building: "\u{1F3E2}", bank: "\u{1F3E6}", briefcase: "\u{1F4BC}", layers: "\u{1F5C2}" };
  const STATUS_PILL = {
    READY: ["ok", "Ready"], RECIPE_REQUIRED: ["warn", "Recipe required"],
    MAPPING_REQUIRED: ["warn", "Mapping required"], DATA_REQUIRED: ["neutral", "Awaiting data"],
    NOT_READY: ["neutral", "Not ready"],
  };

  // ================================================================ HOME
  async function loadHome() {
    show("home");
    const grid = $("#cat-grid");
    grid.innerHTML = "<p class='muted'>Loading…</p>";
    try {
      const r = await call("/api/categories");
      const d = D(r);
      $("#brand-text").textContent = d.app_title || "RWA Calculation";
      grid.innerHTML = "";
      (d.categories || []).forEach((c) => {
        const [pk, pl] = STATUS_PILL[c.status] || STATUS_PILL.NOT_READY;
        grid.appendChild(el("div", { class: "cat-card", onclick: () => go("#/c/" + encodeURIComponent(c.key)) }, [
          el("div", { class: "cat-badge", text: ICONS[c.icon] || ICONS.layers }),
          el("div", { class: "cat-name", text: c.display_name || c.key }),
          el("div", { class: "cat-desc", text: c.description || "" }),
          el("div", { class: "cat-meta" }, [
            el("span", { text: `${c.data_file_count} data` }),
            el("span", { text: `${c.mapping_file_count} mapping` }),
            el("span", { text: `${c.output_file_count} output` }),
          ]),
          el("div", {}, [el("span", { class: "pill " + pk, text: pl })]),
        ]));
      });
      if (!(d.categories || []).length) grid.innerHTML = "<p class='muted'>No categories configured.</p>";
    } catch (e) { grid.innerHTML = "<p class='errbox'>" + e.message + "</p>"; }
  }

  // ================================================================ CATEGORY
  let CUR = null;                // { key, display_name, ... }
  let DATA_FILES = [];
  const PICK = new Set();
  let CMP_CONFIG = { rows: [] }; // remembered for the trace page
  let OUTPUT_XLSX = [];

  async function openCategory(key) {
    show("category");
    $("#cat-title").textContent = key;
    try {
      const r = await call("/api/categories");
      CUR = (D(r).categories || []).find((c) => c.key === key) || { key, display_name: key };
      $("#cat-title").textContent = CUR.display_name || key;
      $("#cat-sub").textContent = CUR.description || "";
    } catch (e) { toast(e.message, "err"); }
    SPLIT = {
      enabled: !!CUR.split_enabled,
      rules: (CUR.split_rules || []).map((r) => Object.assign({}, r)),
      defaults: (CUR.split_rules || []).map((r) => Object.assign({}, r)),
      columns: [],
    };
    setStep(1);
    refreshData();
  }

  function setStep(n) {
    $$("#stepper li").forEach((li) => {
      const s = +li.dataset.step;
      li.classList.toggle("active", s === n);
      li.classList.toggle("done", s < n);
    });
    $$(".step-panel").forEach((p) => p.classList.add("hidden"));
    $("#step-" + n).classList.remove("hidden");
    if (n === 2) renderSplit();
    if (n === 3) renderPick();
    if (n === 5) refreshOutputs();
    if (n === 6) renderCompare();
    if (n === 7) renderTraceLauncher();
  }

  // ---- step 2: split
  let SPLIT = { enabled: false, rules: [], defaults: [], columns: [] };
  const splitPayload = () => SPLIT.enabled
    ? { enabled: true, rules: SPLIT.rules.map((r) => ({ name: r.name, column: r.column, value: r.value })) }
    : { enabled: false };

  async function renderSplit() {
    $("#split-enabled").value = SPLIT.enabled ? "yes" : "no";
    $("#split-body").classList.toggle("hidden", !SPLIT.enabled);
    $("#split-preview-result").innerHTML = "";
    if (!SPLIT.enabled) return;
    drawSplitRules();
    await loadSplitColumns();
  }

  async function loadSplitColumns() {
    try {
      const r = await call("/api/split-preview", { method: "POST", json: { category: CUR.key, selected_paths: [...PICK] } });
      SPLIT.columns = D(r).columns || [];
    } catch (e) { SPLIT.columns = []; toast(e.message, "err"); }
    drawSplitRules();
  }

  function drawSplitRules() {
    const box = $("#split-rules"); box.innerHTML = "";
    SPLIT.rules.forEach((rule) => {
      const col = el("select", {});
      const cols = SPLIT.columns.slice();
      if (rule.column && cols.indexOf(rule.column) < 0) cols.unshift(rule.column);
      cols.forEach((c) => col.appendChild(el("option", { value: c, text: c })));
      col.value = rule.column;
      col.addEventListener("change", () => { rule.column = col.value; });
      const val = el("input", { type: "text", value: rule.value || "", placeholder: "value to keep" });
      val.addEventListener("input", () => { rule.value = val.value; });
      const missing = SPLIT.columns.length && SPLIT.columns.indexOf(rule.column) < 0;
      box.appendChild(el("div", { class: "split-row" }, [
        el("div", {}, [el("label", { text: "dataset" }), el("div", { class: "split-name", text: rule.name })]),
        el("div", {}, [el("label", { text: "column" }), col]),
        el("div", {}, [el("label", { text: "value equals" }), val]),
        missing ? el("div", { class: "split-warn", text: "column not in the data files" }) : null,
      ]));
    });
  }

  async function previewSplit() {
    const box = $("#split-preview-result");
    box.innerHTML = "<p class='muted'>Counting rows…</p>";
    try {
      const r = await call("/api/split-preview", { method: "POST", json: {
        category: CUR.key, selected_paths: [...PICK], rules: splitPayload().rules } });
      const d = D(r);
      box.innerHTML = "";
      if (d.rules_error) return box.appendChild(el("div", { class: "warnbox", text: "⚠ " + d.rules_error }));
      if (d.missing_columns.length) box.appendChild(el("div", { class: "warnbox", text: "⚠ Column(s) not found: " + d.missing_columns.join(", ") }));
      const head = el("tr", {}, [el("th", { text: "File" }), el("th", { text: "Rows" })]
        .concat(SPLIT.rules.map((r) => el("th", { text: r.name }))));
      const tb = el("tbody", {});
      (d.files || []).forEach((f) => {
        if (f.error) return tb.appendChild(el("tr", {}, [el("td", { text: f.filename }), el("td", { class: "err", colspan: SPLIT.rules.length + 1, text: f.error })]));
        tb.appendChild(el("tr", {}, [el("td", { text: f.filename }), el("td", { class: "num", text: f.rows })]
          .concat((f.datasets || []).map((x) => el("td", { class: "num", text: x.rows == null ? "—" : x.rows })))));
      });
      box.appendChild(el("div", { class: "table-wrap" }, [el("table", {}, [el("thead", {}, [head]), tb])]));
      box.appendChild(el("p", { class: "muted small", text: PICK.size ? "Counts for the files you selected." : "Counts for all uploaded data files (the files are chosen in the next step)." }));
    } catch (e) { box.innerHTML = "<div class='errbox'>" + e.message + "</div>"; }
  }

  function dl(folderType, path) {
    window.open(`${API}/api/download?category=${encodeURIComponent(CUR.key)}&folder_type=${folderType}&path=${encodeURIComponent(path)}`, "_blank");
  }

  // ---- step 1: data + mapping
  async function refreshData() {
    const tb = $("#data-table tbody"); tb.innerHTML = "<tr><td colspan='4' class='muted'>Loading…</td></tr>";
    try {
      const r = await call(`/api/files?category=${encodeURIComponent(CUR.key)}&folder_type=data`);
      DATA_FILES = D(r).files || [];
      tb.innerHTML = "";
      if (!DATA_FILES.length) tb.innerHTML = "<tr><td colspan='4' class='muted'>No data files uploaded.</td></tr>";
      DATA_FILES.forEach((f) => tb.appendChild(el("tr", {}, [
        el("td", { text: f.filename }),
        el("td", { text: bytes(f.size) }),
        el("td", { class: "muted small", text: when(f.last_modified) }),
        el("td", {}, [
          el("button", { class: "icon-btn dl", title: "Download", text: "↓", onclick: () => dl("data", f.path) }),
          el("button", { class: "icon-btn", title: "Delete", text: "✕", onclick: () => delFile("data", f.path, f.filename) }),
        ]),
      ])));
    } catch (e) { tb.innerHTML = "<tr><td colspan='4' class='errbox'>" + e.message + "</td></tr>"; }
    try {
      const m = await call(`/api/files?category=${encodeURIComponent(CUR.key)}&folder_type=mapping`);
      const mf = D(m).files || [];
      $("#mapping-current").textContent = mf.length
        ? `Current: ${mf[0].filename}${mf.length > 1 ? "  ⚠ more than one mapping file present" : ""}`
        : "No mapping file yet.";
    } catch (e) { $("#mapping-current").textContent = ""; }
  }

  async function uploadData(fileList) {
    if (!fileList || !fileList.length) return;
    const fd = new FormData();
    fd.append("category", CUR.key); fd.append("folder_type", "data");
    Array.prototype.forEach.call(fileList, (f) => fd.append("files", f));
    try {
      const r = await call("/api/upload", { method: "POST", body: fd });
      toast(`Uploaded ${(D(r).uploaded_files || []).length} file(s)`, "ok");
      refreshData();
    } catch (e) { toast(e.message, "err"); }
  }

  async function uploadMapping() {
    const f = $("#mapping-input").files[0];
    if (!f) return toast("Choose a mapping file first", "err");
    const fd = new FormData();
    fd.append("category", CUR.key); fd.append("folder_type", "mapping");
    fd.append("files", f); fd.append("replace_mapping", "true");
    try {
      await call("/api/upload", { method: "POST", body: fd });
      toast("Mapping file updated", "ok"); $("#mapping-input").value = ""; refreshData();
    } catch (e) { toast(e.message, "err"); }
  }

  async function delFile(folderType, path, name) {
    if (!(await confirmModal("Delete file", `Delete "${name}" from ${folderType}? This cannot be undone.`))) return;
    try {
      await call("/api/delete", { method: "POST", json: { category: CUR.key, folder_type: folderType, path } });
      toast("Deleted " + name, "ok");
      folderType === "data" ? refreshData() : refreshOutputs();
    } catch (e) { toast(e.message, "err"); }
  }

  async function clearData() {
    if (!DATA_FILES.length) return;
    if (!(await confirmModal("Clear data files", `Delete all ${DATA_FILES.length} data file(s) for ${CUR.display_name}?`))) return;
    for (const f of DATA_FILES.slice()) {
      try { await call("/api/delete", { method: "POST", json: { category: CUR.key, folder_type: "data", path: f.path } }); } catch (e) {}
    }
    toast("Cleared", "ok"); refreshData();
  }

  // ---- step 2: select + validate
  function renderPick() {
    const box = $("#pick-list"); box.innerHTML = "";
    if (!DATA_FILES.length) { box.innerHTML = "<p class='muted'>No data files — upload some in step 1.</p>"; return; }
    if (!PICK.size) DATA_FILES.forEach((f) => PICK.add(f.path));
    DATA_FILES.forEach((f) => {
      const cb = el("input", { type: "checkbox" });
      cb.checked = PICK.has(f.path);
      cb.addEventListener("change", () => { cb.checked ? PICK.add(f.path) : PICK.delete(f.path); });
      box.appendChild(el("label", { class: "pick-row" }, [
        cb, el("span", { class: "fn", text: f.filename }), el("span", { class: "fm", text: bytes(f.size) }),
      ]));
    });
  }

  async function runValidate() {
    const box = $("#validate-result");
    if (!PICK.size) return toast("Select at least one file", "err");
    box.innerHTML = "<p class='muted'>Validating…</p>";
    try {
      const r = await call("/api/validate-inputs", { method: "POST", json: { category: CUR.key, selected_paths: [...PICK], split: splitPayload() } });
      const v = D(r);
      box.innerHTML = "";
      box.appendChild(el("div", { class: v.valid ? "okbox" : "errbox", text: v.valid ? "Validation passed." : "Validation found problems." }));
      const kv = el("dl", { class: "kv" }, [
        el("dt", { text: "Input rows" }), el("dd", { text: v.input_row_count }),
        el("dt", { text: "Input columns" }), el("dd", { text: v.input_column_count }),
        el("dt", { text: "Mapping file" }), el("dd", { text: (v.mapping_file && v.mapping_file.filename) || "–" }),
        el("dt", { text: "Mapping rows" }), el("dd", { text: v.mapping_row_count }),
      ]);
      box.appendChild(kv);
      (v.errors || []).forEach((e) => box.appendChild(el("div", { class: "warnbox", text: "⚠ " + e })));
      if (v.split_preview && SPLIT.enabled) {
        (v.split_preview.files || []).filter((f) => f.datasets).forEach((f) => box.appendChild(el("p", { class: "muted small",
          text: `${f.filename} will split into: ` + f.datasets.map((x) => `${x.name} (${x.rows == null ? "—" : x.rows} rows)`).join(", ") })));
      }
      (v.input_files || []).forEach((f) => box.appendChild(el("p", { class: "muted small",
        text: `${f.filename}: ${f.row_count} rows, header on row ${f.detected_header_row}` })));
    } catch (e) { box.innerHTML = "<div class='errbox'>" + e.message + "</div>"; }
  }

  // ---- step 3: run
  async function runCalc() {
    if (!PICK.size) { setStep(3); return toast("Select files first", "err"); }
    $("#run-status").textContent = "Running the Dataiku recipe… this can take a while.";
    $("#run-result").innerHTML = "";
    try {
      const r = await call("/api/run", { method: "POST", json: { category: CUR.key, selected_paths: [...PICK], split: splitPayload() } });
      const s = D(r);
      $("#run-status").innerHTML = "<span class='pill ok'>" + (s.status || "COMPLETED") + "</span>";
      const box = $("#run-result");
      box.appendChild(el("dl", { class: "kv" }, [
        el("dt", { text: "Run id" }), el("dd", { text: s.run_id }),
        el("dt", { text: "Recipe" }), el("dd", { text: s.recipe_id || "–" }),
        el("dt", { text: "Output file" }), el("dd", { text: s.output_file || "–" }),
        el("dt", { text: "Started" }), el("dd", { text: when(s.started_at) }),
        el("dt", { text: "Completed" }), el("dd", { text: when(s.completed_at) }),
      ]));
      if (s.split && (s.split.datasets || []).length) {
        box.appendChild(el("h3", { text: `Split into ${s.split.datasets.length} dataset(s)`, style: "margin-top:14px;font-size:14px" }));
        const tb = el("tbody", {});
        s.split.datasets.forEach((d) => tb.appendChild(el("tr", {}, [
          el("td", { text: d.source_filename }), el("td", { text: d.split_name }),
          el("td", { class: "muted small", text: `${d.column} = ${d.value}` }), el("td", { class: "num", text: d.row_count })])));
        box.appendChild(el("div", { class: "table-wrap" }, [el("table", {}, [
          el("thead", {}, [el("tr", {}, ["Source file", "Dataset", "Filter", "Rows"].map((h) => el("th", { text: h })))]), tb])]));
        (s.split.skipped || []).forEach((m) => box.appendChild(el("div", { class: "warnbox", text: "⚠ " + m })));
      }
      const pf = s.produced_files || [];
      if (pf.length) {
        box.appendChild(el("h3", { text: "Produced files", style: "margin-top:14px;font-size:14px" }));
        box.appendChild(el("div", { class: "chip-list" }, pf.map((f) => el("span", { class: "chip", text: f.filename }))));
      }
    } catch (e) { $("#run-status").innerHTML = "<span class='errbox'>" + e.message + "</span>"; }
  }

  // ---- step 4: outputs
  let TEMPLATE_XLSX = [];
  async function refreshOutputs() {
    OUTPUT_XLSX = []; TEMPLATE_XLSX = [];
    for (const which of ["output", "template"]) {
      const tb = $(`#${which}-table tbody`); tb.innerHTML = "";
      try {
        const r = await call(`/api/files?category=${encodeURIComponent(CUR.key)}&folder_type=${which}`);
        const files = D(r).files || [];
        if (!files.length) { tb.innerHTML = "<tr><td colspan='2' class='muted'>Empty.</td></tr>"; continue; }
        files.forEach((f) => {
          const traceable = [".xlsx", ".xlsm"].indexOf(f.extension) >= 0 && !/_COMPARISON_/i.test(f.filename);
          if (traceable && which === "output") OUTPUT_XLSX.push(f);
          if (traceable && which === "template") TEMPLATE_XLSX.push(f);
          tb.appendChild(el("tr", {}, [
            el("td", { text: f.filename }),
            el("td", {}, [
              el("button", { class: "icon-btn dl", title: "Download", text: "↓", onclick: () => dl(which, f.path) }),
              traceable ? el("button", { class: "icon-btn", title: "Trace", text: "🔬", onclick: () => go(`#/c/${encodeURIComponent(CUR.key)}/trace?folder=${which}&file=${encodeURIComponent(f.path)}`) }) : null,
              el("button", { class: "icon-btn", title: "Delete", text: "✕", onclick: () => delFile(which, f.path, f.filename) }),
            ]),
          ]));
        });
      } catch (e) { tb.innerHTML = "<tr><td colspan='2' class='errbox'>" + e.message + "</td></tr>"; }
    }
  }
  // the trace needs the formula-bearing workbooks -> prefer template, fall back to output
  const traceFiles = () => OUTPUT_XLSX.length ? { which: "output", files: OUTPUT_XLSX } : { which: "template", files: TEMPLATE_XLSX };

  // ---- step 5: compare
  let CMP_COLS = [];
  let CMP_SOURCE = "output";
  async function renderCompare() {
    const box = $("#cmp-rows"); box.innerHTML = "<p class='muted'>Loading column schema…</p>";
    try {
      const r = await call(`/api/output-schema?category=${encodeURIComponent(CUR.key)}`);
      CMP_COLS = D(r).columns || [];
      CMP_SOURCE = D(r).source_folder || "output";
    } catch (e) { CMP_COLS = []; }
    box.innerHTML = "";
    if (!CMP_COLS.length) { box.innerHTML = "<div class='warnbox'>No produced workbooks yet — run the calculation first.</div>"; return; }
    box.appendChild(el("p", { class: "muted small", text: `Comparing the workbooks in the ${CMP_SOURCE} folder — every formula column is evaluated, so a mismatched row can be walked in Trace.` }));
    if (!CMP_CONFIG.rows.length) CMP_CONFIG.rows = [{ left: "", right: "", type: "numeric" }];
    CMP_CONFIG.rows.forEach((rule, i) => box.appendChild(cmpRow(rule, i)));
  }
  function cmpRow(rule, i) {
    const opt = (v) => el("option", { value: v, text: v });
    const left = el("select", {}, [el("option", { value: "", text: "— left column —" })].concat(CMP_COLS.map(opt)));
    const right = el("select", {}, [el("option", { value: "", text: "— right column —" })].concat(CMP_COLS.map(opt)));
    const type = el("select", {}, ["numeric", "text", "date", "exact"].map(opt));
    left.value = rule.left; right.value = rule.right; type.value = rule.type || "numeric";
    left.addEventListener("change", () => rule.left = left.value);
    right.addEventListener("change", () => rule.right = right.value);
    type.addEventListener("change", () => rule.type = type.value);
    return el("div", { class: "cmp-row" }, [
      el("div", {}, [el("label", { text: "left" }), left]),
      el("div", {}, [el("label", { text: "right" }), right]),
      el("div", {}, [el("label", { text: "type" }), type]),
      el("div", { class: "cmp-rm" }, [el("button", { class: "icon-btn", text: "✕", title: "Remove",
        onclick: () => { CMP_CONFIG.rows.splice(i, 1); renderCompare(); } })]),
    ]);
  }
  async function runCompare() {
    const rules = CMP_CONFIG.rows.filter((r) => r.left && r.right);
    if (!rules.length) return toast("Pick at least one column pair", "err");
    const box = $("#cmp-result"); box.innerHTML = "<p class='muted'>Comparing…</p>";
    try {
      const threshold = parseFloat($("#cmp-threshold").value) || 0;
      const r = await call("/api/compare-all-outputs", { method: "POST", json: {
        category: CUR.key,
        comparisons: rules.map((x) => ({ left_column: x.left, right_column: x.right, comparison_type: x.type })),
        threshold: threshold,
      } });
      const s = D(r);
      box.innerHTML = "";
      (s.warnings || []).forEach((w) => box.appendChild(el("div", { class: "warnbox", text: "⚠ " + w })));
      box.appendChild(el("dl", { class: "kv" }, [
        el("dt", { text: "Compared from" }), el("dd", { text: (s.source_folder || "template") + " folder" }),
        el("dt", { text: "Files compared" }), el("dd", { text: s.files_compared }),
        el("dt", { text: "Rows" }), el("dd", { text: s.total_rows }),
        el("dt", { text: "Matched" }), el("dd", { text: s.matched_rows }),
        el("dt", { text: "Mismatched" }), el("dd", { text: s.mismatched_rows }),
        el("dt", { text: "Match rate" }), el("dd", { text: s.match_rate != null ? s.match_rate + " %" : "–" }),
      ]));
      box.appendChild(el("div", { class: "row-actions", style: "justify-content:flex-start" }, [
        el("button", { class: "btn ghost small", text: "Download " + s.comparison_file, onclick: () => dl("output", "/" + s.comparison_file) }),
        el("button", { class: "btn small", text: "Trace mismatched rows →", onclick: async () => {
          let tf = traceFiles();
          if (!tf.files.length) { await refreshOutputs(); tf = traceFiles(); }
          if (!tf.files.length) { setStep(5); return toast("No formula workbook to trace — run the calculation first", "err"); }
          go(`#/c/${encodeURIComponent(CUR.key)}/trace?folder=${tf.which}&file=${encodeURIComponent(tf.files[0].path)}`);
        } }),
      ]));
      CMP_CONFIG.lastRules = rules;
    } catch (e) { box.innerHTML = "<div class='errbox'>" + e.message + "</div>"; }
  }

  // ---- step 6: trace launcher
  function fillTraceSelect() {
    const sel = $("#trace-file"); sel.innerHTML = "";
    const tf = traceFiles();
    tf.files.forEach((f) => sel.appendChild(el("option", { value: tf.which + "|" + f.path, text: f.filename })));
    if (!tf.files.length) sel.appendChild(el("option", { value: "", text: "no formula workbooks — run the calculation first" }));
  }
  function renderTraceLauncher() {
    fillTraceSelect();
    if (!traceFiles().files.length) refreshOutputs().then(fillTraceSelect);
    const sel = $("#trace-file");
    $("#trace-open").onclick = () => {
      const v = sel.value;
      if (!v) return toast("No workbook to trace", "err");
      const [which, p] = v.split("|");
      go(`#/c/${encodeURIComponent(CUR.key)}/trace?folder=${which}&file=${encodeURIComponent(p)}`);
    };
  }

  // ================================================================ TRACE PAGE
  let TRACE = null;   // { sid, file, category, summary }

  async function openTracePage(key, q) {
    show("trace");
    if (!CUR || CUR.key !== key) {
      try { const r = await call("/api/categories"); CUR = (D(r).categories || []).find((c) => c.key === key) || { key, display_name: key }; }
      catch (e) { CUR = { key, display_name: key }; }
    }
    $("#trace-back").onclick = () => go("#/c/" + encodeURIComponent(key));
    $("#trace-title").textContent = "Trace — " + (CUR.display_name || key);
    $("#trace-sub").textContent = "Opening " + (q.file || "").split("/").pop() + " …";
    $("#tr-journey").innerHTML = "";
    $("#tr-findings").innerHTML = ""; $("#tr-mismatch-list").innerHTML = "";
    try {
      const r = await call("/excel-trace/open-output", { method: "POST", json: {
        category: key, path: q.file, folder_type: q.folder || "output" } });
      TRACE = { sid: r.session_id, file: r.filename, category: key, folder: q.folder || "output", summary: r.summary };
      const s = r.summary;
      $("#trace-sub").textContent = r.filename;
      if (!s.engine_build) {
        $("#tr-summary").innerHTML = "<span class='errbox'>An old <code>excel_deep_trace.py</code> is on the project libraries — "
          + "replace it with the one from <code>rwa-calculation-webapp/python-lib/</code> and restart the backend.</span>";
        return;
      }
      if (!s.total_formula_cells) {
        $("#tr-summary").innerHTML = "<span class='warnbox'>This workbook has <b>no live Excel formulas</b> — "
          + "the calculation recipe writes values only. Nothing to trace here.</span>";
      } else {
        $("#tr-summary").textContent =
          `${s.total_formula_cells} formula cells · ${s.sheet_names.length} sheet(s) · `
          + `${s.conditional_cells} conditional · ${s.lookup_cells} lookup · ${s.findings.total} finding(s)`;
      }
      const sheetSel = $("#tr-sheet"); sheetSel.innerHTML = "";
      (s.sheet_names || []).forEach((n) => sheetSel.appendChild(el("option", { value: n, text: n })));
      // pick the sheet with the most formula cells
      const busiest = (s.sheets || []).slice().sort((a, b) => b.formula_cells - a.formula_cells)[0];
      if (busiest) sheetSel.value = busiest.name;
      loadFindings();
      renderMismatchBox();
    } catch (e) {
      $("#tr-summary").innerHTML = "<span class='errbox'>" + e.message + "</span>";
    }
  }

  async function loadFindings() {
    const box = $("#tr-findings");
    try {
      const r = await call(`/excel-trace/findings?session_id=${TRACE.sid}`);
      const fs = r.findings || [];
      box.innerHTML = "";
      if (!fs.length) { box.innerHTML = "<p class='muted small'>No findings.</p>"; return; }
      fs.slice(0, 40).forEach((f) => {
        const sev = { HIGH: "err", MEDIUM: "warn", LOW: "neutral", INFO: "neutral" }[f.severity] || "neutral";
        box.appendChild(el("div", { class: "rj-narr", style: "margin-bottom:6px;cursor:pointer",
          onclick: () => explainCell(f.cell) }, [
          el("span", { class: "pill " + sev, text: f.severity }), " ",
          el("span", { class: "mono small", text: f.cell }),
          el("p", { class: "muted small", text: f.message }),
        ]));
      });
    } catch (e) { box.innerHTML = "<p class='errbox'>" + e.message + "</p>"; }
  }

  async function renderMismatchBox() {
    const box = $("#tr-mismatch-list");
    const rules = (CMP_CONFIG.lastRules || []).filter((r) => r.type === "numeric");
    const rule = rules[0];
    box.innerHTML = "";
    if (!rule) {
      box.innerHTML = "<p class='muted small'>Run a numeric comparison first, or pick a row manually above.</p>";
      return;
    }
    const btn = el("button", { class: "btn ghost small", text: `Scan ${rule.left} vs ${rule.right}` });
    btn.onclick = async () => {
      btn.disabled = true; btn.textContent = "Scanning…";
      try {
        const r = await call("/excel-trace/mismatch-rows", { method: "POST", json: {
          session_id: TRACE.sid, sheet: $("#tr-sheet").value, left: rule.left, right: rule.right, tolerance: 0.01 } });
        box.innerHTML = `<p class='muted small'>${r.total} mismatch row(s)</p>`;
        (r.rows || []).slice(0, 60).forEach((row) => box.appendChild(el("div", { class: "rj-narr", style: "cursor:pointer",
          onclick: () => { $("#tr-row").value = row.row; loadRowJourney(); } }, [
          el("span", { class: "mono small", text: "row " + row.row }), " ",
          el("span", { class: "muted small", text: `${fmt(row.left)} vs ${fmt(row.right)}  (Δ ${row.diff == null ? "—" : trimNum(row.diff)})` }),
        ])));
      } catch (e) { box.innerHTML = "<p class='errbox'>" + e.message + "</p>"; btn.disabled = false; btn.textContent = "Retry scan"; }
    };
    box.appendChild(btn);
  }

  async function loadRowJourney() {
    const sheet = $("#tr-sheet").value;
    const row = parseInt($("#tr-row").value, 10);
    if (!row) return toast("Enter a row number", "err");
    const box = $("#tr-journey"); box.innerHTML = "<p class='muted'>Building row journey…</p>";
    try {
      const r = await call("/excel-trace/row-journey", { method: "POST", json: { session_id: TRACE.sid, sheet, row } });
      const j = r.journey;
      box.innerHTML = "";
      box.appendChild(el("div", { class: "view-head" }, [
        el("h1", { style: "font-size:17px", text: `${sheet} · row ${row}` }),
        el("p", { class: "muted small", text: `${j.formula_cells} computed cell(s), ${j.input_cells} input(s)` }),
      ]));
      if ((j.inputs || []).length) {
        box.appendChild(el("div", { class: "chip-list", style: "margin-bottom:14px" },
          j.inputs.map((c) => el("span", { class: "chip", text: `${c.header || c.cell}: ${fmt(c.value)}` }))));
      }
      (j.cells || []).forEach((c) => {
        const card = el("div", { class: "rj-cell" + (c.tracer_reproduced_excel === false ? " mismatch" : "") }, [
          el("div", { class: "rj-head" }, [
            el("span", { class: "rj-name", text: c.header || c.cell }),
            el("span", { class: "rj-key", text: c.key }),
            el("span", { class: "rj-val", text: c.value_repr }),
          ]),
          el("div", { class: "rj-formula", text: c.formula }),
          el("div", { class: "rj-narr" }, (c.narrative || []).map((n) => el("p", { text: n }))),
          el("div", { class: "rj-actions" }, [
            el("button", { class: "btn ghost small", text: "Open full trace →", onclick: () => explainCell(c.key) }),
          ]),
        ]);
        box.appendChild(card);
      });
      if (!(j.cells || []).length) box.appendChild(el("p", { class: "muted", text: "No formula cells on this row." }));
    } catch (e) { box.innerHTML = "<div class='errbox'>" + e.message + "</div>"; }
  }

  // ================================================================ TRACE DRAWER
  const DSTACK = [];
  async function explainCell(cell, push) {
    try {
      const [treeR, explainR] = await Promise.all([
        call("/excel-trace/expression-tree", { method: "POST", json: { session_id: TRACE.sid, cell } }),
        call("/excel-trace/explain-cell", { method: "POST", json: { session_id: TRACE.sid, cell } }).catch(() => null),
      ]);
      const tree = treeR.tree;
      if (push) DSTACK.push(cell); else DSTACK.length = 0, DSTACK.push(cell);
      DRAWER = { cell, tree, explain: explainR };
      $("#drawer").classList.remove("hidden"); $("#drawer-scrim").classList.remove("hidden");
      renderDrawer();
      showTab("tree");
    } catch (e) { toast(e.message, "err"); }
  }
  let DRAWER = null;
  function closeDrawer() { $("#drawer").classList.add("hidden"); $("#drawer-scrim").classList.add("hidden"); }

  function renderDrawer() {
    const t = DRAWER.tree;
    const cr = $("#d-crumbs"); cr.innerHTML = "";
    DSTACK.forEach((k, i) => {
      if (i) cr.appendChild(el("span", { class: "sep", text: " ▸ " }));
      if (i === DSTACK.length - 1) cr.appendChild(el("span", { class: "cur", text: k }));
      else cr.appendChild(el("button", { text: k, onclick: () => { DSTACK.length = i + 1; explainCell(k); } }));
    });
    $("#d-col").textContent = t.key;
    $("#d-val").textContent = t.value_repr;
    $("#d-dtype").textContent = (t.root && t.root.dtype) || "";
    $("#d-sub").textContent = t.formula || "";
    const flags = $("#d-flags"); flags.innerHTML = "";
    if (t.is_formula && t.no_excel_value) {
      flags.appendChild(el("div", { class: "flag info" }, [el("span", { text: "•" }),
        el("span", { text: "This workbook stored no computed value for this cell — the value shown is the tracer's own evaluation of the formula." })]));
    } else if (t.is_formula && t.tracer_reproduced_excel === false) {
      flags.appendChild(el("div", { class: "flag warn" }, [el("span", { text: "⚠" }),
        el("span", { html: `The tracer couldn't fully reproduce Excel's value (unsupported or volatile function). Headline shows Excel's saved value <b>${t.excel_value_repr}</b>; the breakdown below is partial.` })]));
    }
    (t.notes || []).forEach((n) => flags.appendChild(el("div", { class: "flag info" }, [el("span", { text: "•" }), el("span", { text: n })])));
    renderTree(); renderFlow(); renderExplain(); renderGraph(); renderSource();
  }
  function showTab(name) {
    $$("#d-tabs button").forEach((b) => b.classList.toggle("active", b.dataset.tab === name));
    $$("#drawer .tabpane").forEach((p) => p.classList.add("hidden"));
    $("#d-tab-" + name).classList.remove("hidden");
  }

  const OPLABEL = { if: "IF", op: "op", compare: "cmp", concat: "&", logical: "bool", func: "fn", lookup: "lookup",
    agg: "agg", cell: "cell", name: "name", number: "num", text: "txt", bool: "bool", empty: "∅", error: "err",
    range: "range", unary: "±", skipped: "—", unsupported: "?" };

  function renderTree() {
    const box = $("#d-tab-tree"); box.innerHTML = "";
    if (!DRAWER.tree.root) { box.innerHTML = "<p class='muted'>" + (DRAWER.tree.parse_error ? "Formula could not be parsed." : "No breakdown.") + "</p>"; return; }
    box.appendChild(nodeEl(DRAWER.tree.root, true));
  }
  function nodeEl(n, open) {
    const kids = n.children || [];
    const wrap = el("div", { class: "tnode" + (n.taken === true ? " branch-taken" : n.taken === false ? " branch-skip" : "") + (n.kind === "error" ? " is-error" : "") });
    const has = kids.length > 0;
    const tog = el("span", { class: "tn-toggle", text: has ? (open ? "▾" : "▸") : "·" });
    const kb = el("div", { class: "tn-kids" });
    if (has) {
      kids.forEach((c) => kb.appendChild(nodeEl(c, c.evaluated !== false)));
      if (!open) kb.classList.add("hidden");
      tog.style.cursor = "pointer";
      tog.addEventListener("click", () => { const h = kb.classList.toggle("hidden"); tog.textContent = h ? "▸" : "▾"; });
    }
    const opc = n.kind === "if" ? "op-if" : n.kind === "lookup" ? "op-lookup" : n.kind === "cell" ? "op-cell" : n.kind === "error" ? "op-error" : "";
    const expr = el("span", { class: "tn-expr" });
    if (n.kind === "cell" && n.detail && n.detail.is_formula && n.detail.key) {
      expr.appendChild(el("span", { class: "cell-link", text: n.excel, title: "trace " + n.detail.key, onclick: () => explainCell(n.detail.key, true) }));
    } else expr.textContent = n.excel;
    const row = el("div", { class: "tn-row" }, [tog, el("span", { class: "tn-op " + opc, text: OPLABEL[n.kind] || n.kind }), expr]);
    if (n.evaluated !== false) {
      row.appendChild(el("span", { class: "tn-arrow", text: "=" }));
      row.appendChild(el("span", { class: "tn-val" + (n.value === null ? " empty" : ""), text: n.value_repr }));
    } else row.appendChild(el("span", { class: "tn-tag skip", text: "not taken" }));
    if (n.taken === true && n.role !== "root") row.appendChild(el("span", { class: "tn-tag taken", text: "taken" }));
    if (n.role && ["cond", "then", "else", "key", "value", "fallback"].indexOf(n.role) >= 0) row.appendChild(el("span", { class: "tn-role", text: n.role }));
    wrap.appendChild(row);
    const det = detailEl(n);
    if (det) wrap.appendChild(det);
    if (has) wrap.appendChild(kb);
    return wrap;
  }
  function detailEl(n) {
    const d = n.detail || {};
    if (n.kind === "lookup" && d.table_preview) {
      const tbl = el("table", { class: "lk-table" });
      (d.table_preview).forEach((r) => tbl.appendChild(el("tr", { class: r.matched ? "match" : "" },
        (r.cells || []).map((c) => el("td", { text: fmt(c) })))));
      return el("div", { class: "tn-detail" }, [
        el("div", { class: "muted small", text: `${d.function} key ${d.lookup_repr} · ${d.match_found ? "matched row " + (d.matched_row_index + 1) : "no match"}` }),
        tbl,
      ]);
    }
    if (n.kind === "error") return el("div", { class: "tn-detail muted small", text: d.error || "" });
    if (d.short_circuit_at != null) return el("div", { class: "tn-detail muted small", text: `${d.operator} short-circuited at argument ${d.short_circuit_at + 1}` });
    if (n.kind === "if" && d.cond_empty) return el("div", { class: "tn-detail muted small", text: "condition empty → treated as TRUE" });
    return null;
  }

  function renderFlow() {
    const box = $("#d-tab-flow"); box.innerHTML = "";
    if (!DRAWER.tree.root) return;
    const steps = [];
    (function walk(n) {
      (n.children || []).forEach((c) => { if (c.evaluated !== false) walk(c); });
      if (n.evaluated === false) return;
      if (["number", "text", "bool", "empty"].indexOf(n.kind) >= 0) return;
      steps.push(n);
    })(DRAWER.tree.root);
    box.appendChild(el("p", { class: "muted small", text: "Every evaluated sub-expression, innermost first." }));
    const list = el("div", { class: "flow" });
    steps.forEach((n, i) => list.appendChild(el("div", { class: "flow-step" + (i === steps.length - 1 ? " final" : "") }, [
      el("span", { class: "fi", text: i + 1 }), el("span", { class: "fe", text: n.excel }), el("span", { class: "fv", text: n.value_repr }),
    ])));
    box.appendChild(list);
  }

  function renderExplain() {
    const box = $("#d-tab-explain"); box.innerHTML = "";
    (DRAWER.tree.narrative || []).forEach((line) => box.appendChild(el("p", { style: "margin:0 0 10px;font-size:13.5px", text: line })));
    if ((DRAWER.tree.notes || []).length) {
      box.appendChild(el("p", { class: "muted small", text: "Notes:" }));
      const ul = el("ul", {});
      DRAWER.tree.notes.forEach((n) => ul.appendChild(el("li", { class: "muted small", text: n })));
      box.appendChild(ul);
    }
  }

  function renderGraph() {
    const box = $("#d-tab-graph"); box.innerHTML = "";
    const g = DRAWER.explain && DRAWER.explain.precedent_graph;
    if (!g || !g.nodes || !g.nodes.length) { box.innerHTML = "<p class='muted'>No precedent graph.</p>"; return; }
    const NW = 150, NH = 26, GX = 40, GY = 12, PAD = 12;
    const byDepth = {};
    g.nodes.forEach((n) => { (byDepth[n.depth] = byDepth[n.depth] || []).push(n); });
    const depths = Object.keys(byDepth).map(Number).sort((a, b) => b - a);  // deepest precedents left
    const pos = {}; let maxRows = 0;
    depths.forEach((d, di) => {
      byDepth[d].forEach((n, ri) => { pos[n.id] = { x: PAD + di * (NW + GX), y: PAD + ri * (NH + GY) }; });
      maxRows = Math.max(maxRows, byDepth[d].length);
    });
    const W = PAD * 2 + depths.length * (NW + GX) - GX, H = PAD * 2 + maxRows * (NH + GY) - GY;
    const svg = el("svg", { viewBox: `0 0 ${W} ${H}`, width: W, height: H });
    (g.edges || []).forEach((e) => {
      const a = pos[e.from], b = pos[e.to]; if (!a || !b) return;
      const x1 = a.x + NW, y1 = a.y + NH / 2, x2 = b.x, y2 = b.y + NH / 2, mx = (x1 + x2) / 2;
      svg.appendChild(el("path", { class: "gedge", d: `M${x1},${y1} C${mx},${y1} ${mx},${y2} ${x2},${y2}` }));
    });
    g.nodes.forEach((n) => {
      const p = pos[n.id];
      const cls = n.id === g.root ? "root" : n.is_formula ? "formula" : "input";
      const grp = el("g", { class: "gnode " + cls, transform: `translate(${p.x},${p.y})`,
        onclick: () => n.is_formula && explainCell(n.id, true) }, [
        el("rect", { width: NW, height: NH, rx: 7 }),
        el("text", { x: 8, y: NH / 2 + 3, text: `${n.cell}  ${n.value_repr}`.slice(0, 24) }),
      ]);
      grp.appendChild(el("title", { text: n.id + (n.formula ? "\n" + n.formula : "") }));
      svg.appendChild(grp);
    });
    box.appendChild(svg);
  }

  function renderSource() {
    const box = $("#d-tab-source"); box.innerHTML = "";
    const t = DRAWER.tree, x = DRAWER.explain || {};
    box.appendChild(el("h4", { text: "Formula" }));
    box.appendChild(el("pre", { class: "code", text: t.formula || "—" }));
    const meta = x.formula_metadata || {};
    box.appendChild(el("dl", { class: "kv" }, [
      el("dt", { text: "Family" }), el("dd", { text: t.family || meta.formula_family || "–" }),
      el("dt", { text: "Functions" }), el("dd", { text: (t.functions || meta.functions || []).join(", ") || "–" }),
      el("dt", { text: "Dependencies" }), el("dd", { text: meta.dependency_count != null ? meta.dependency_count : "–" }),
      el("dt", { text: "Excel value" }), el("dd", { text: t.excel_value_repr }),
      el("dt", { text: "Tracer value" }), el("dd", { text: t.value_repr }),
    ]));
  }

  // ================================================================ ADMIN
  let ADMIN_CATS = [], ADMIN_ACTIVE = null;
  async function openAdmin(catKey) {
    show("admin");
    if (!ADMIN_TOKEN) { $("#admin-login").classList.remove("hidden"); $("#admin-panel").classList.add("hidden"); return; }
    $("#admin-login").classList.add("hidden"); $("#admin-panel").classList.remove("hidden");
    try {
      const r = await call("/api/admin/config");
      ADMIN_CATS = D(r).categories || [];
      const tabs = $("#admin-tabs"); tabs.innerHTML = "";
      ADMIN_CATS.forEach((c) => tabs.appendChild(el("button", { text: c.display_name || c.key, "data-tab": c.key,
        onclick: () => selectAdminTab(c.key) })));
      tabs.appendChild(el("button", { text: "⚙ Settings", "data-tab": "__settings__", onclick: () => selectAdminTab("__settings__") }));
      const want = catKey && ADMIN_CATS.some((c) => c.key === catKey) ? catKey : (ADMIN_CATS[0] && ADMIN_CATS[0].key);
      selectAdminTab(want || "__settings__");
    } catch (e) {
      if (e.data && e.data.message && /authentication/i.test(e.data.message)) { ADMIN_TOKEN = null; try { sessionStorage.removeItem("rwa_admin_token"); } catch (x) {} return openAdmin(); }
      toast(e.message, "err");
    }
  }
  function selectAdminTab(key) {
    ADMIN_ACTIVE = key;
    $$("#admin-tabs button").forEach((b) => b.classList.toggle("active", b.dataset.tab === key));
    try { history.replaceState(null, "", "#/admin" + (key === "__settings__" ? "" : "/" + encodeURIComponent(key))); } catch (e) {}
    key === "__settings__" ? renderSettings() : renderCategoryForm(key);
  }

  function renderCategoryForm(key) {
    const body = $("#admin-tab-body"); body.innerHTML = "";
    const c = ADMIN_CATS.find((x) => x.key === key);
    if (!c) return;
    const nm = el("input", { type: "text", value: c.display_name || "" });
    const rn = el("input", { type: "text", value: c.recipe_name || "", placeholder: "friendly recipe name (shown in run history)" });
    const rid = el("input", { type: "text", value: c.recipe_id || "", placeholder: "compute_XXXXXXXX" });
    const folders = {};
    ["data", "mapping", "output", "template"].forEach((k) => {
      folders[k] = el("input", { type: "text", value: (c.folders || {})[k] || "", placeholder: "managed folder id" });
    });
    body.appendChild(el("div", { class: "admin-section" }, [
      el("h3", { text: "Identity" }),
      el("p", { class: "hint", text: "Key: " + c.key }),
      el("label", { text: "Display name" }), nm,
      el("label", { text: "Recipe name" }), rn,
      el("label", { text: "Recipe id (from the Flow)" }), rid,
    ]));
    body.appendChild(el("div", { class: "admin-section" }, [
      el("h3", { text: "Managed folders" }),
      el("p", { class: "hint", text: "The 8-character id in the folder's URL / Settings tab." }),
      el("div", { class: "grid-2" }, ["data", "mapping", "output", "template"].map((k) =>
        el("div", { class: "field" }, [el("label", { text: k }), folders[k]]))),
    ]));
    body.appendChild(el("div", { class: "row-actions" }, [
      el("button", { class: "btn", text: "Save", onclick: async () => {
        try {
          await call(`/api/admin/category/${encodeURIComponent(key)}`, { method: "POST", json: {
            display_name: nm.value, recipe_name: rn.value, recipe_id: rid.value,
            folders: { data: folders.data.value, mapping: folders.mapping.value, output: folders.output.value, template: folders.template.value },
          } });
          toast("Saved", "ok"); openAdmin(key);
        } catch (e) { toast(e.message, "err"); }
      } }),
    ]));
  }

  function renderSettings() {
    const body = $("#admin-tab-body"); body.innerHTML = "";
    const u = el("input", { type: "text", placeholder: "new username (optional)" });
    const cur = el("input", { type: "password", placeholder: "current password" });
    const np = el("input", { type: "password", placeholder: "new password (min 6)" });
    body.appendChild(el("div", { class: "admin-section" }, [
      el("h3", { text: "Change admin credentials" }),
      el("label", { text: "New username" }), u,
      el("label", { text: "Current password" }), cur,
      el("label", { text: "New password" }), np,
      el("div", { class: "row-actions" }, [el("button", { class: "btn", text: "Update", onclick: async () => {
        try {
          await call("/api/admin/change-credentials", { method: "POST", json: {
            new_username: u.value || undefined, current_password: cur.value, new_password: np.value } });
          toast("Updated — sign in again", "ok");
          ADMIN_TOKEN = null; try { sessionStorage.removeItem("rwa_admin_token"); } catch (e) {}
          openAdmin();
        } catch (e) { toast(e.message, "err"); }
      } })]),
    ]));
    body.appendChild(el("div", { class: "admin-section" }, [
      el("h3", { text: "Session" }),
      el("button", { class: "btn ghost", text: "Sign out", onclick: () => {
        ADMIN_TOKEN = null; try { sessionStorage.removeItem("rwa_admin_token"); } catch (e) {}
        openAdmin();
      } }),
    ]));
  }

  async function doLogin() {
    $("#login-err").textContent = "";
    try {
      const r = await call("/api/login", { method: "POST", json: { username: $("#login-user").value, password: $("#login-pass").value } });
      ADMIN_TOKEN = D(r).token;
      try { sessionStorage.setItem("rwa_admin_token", ADMIN_TOKEN); } catch (e) {}
      openAdmin(D(r).must_change ? "__settings__" : null);
      if (D(r).must_change) toast("Please change the default credentials.", "err");
    } catch (e) { $("#login-err").textContent = e.message; }
  }

  // ================================================================ wire up
  function init() {
    $("#nav-home").addEventListener("click", () => go("#/"));
    $("#brand-home").addEventListener("click", () => go("#/"));
    $("#nav-admin").addEventListener("click", () => go("#/admin"));
    $("#cat-back").addEventListener("click", () => go("#/"));
    $("#admin-back").addEventListener("click", () => go("#/"));
    $("#login-btn").addEventListener("click", doLogin);
    $("#login-pass").addEventListener("keydown", (e) => { if (e.key === "Enter") doLogin(); });

    $("#browse-btn").addEventListener("click", () => $("#file-input").click());
    $("#file-input").addEventListener("change", (e) => uploadData(e.target.files));
    const dz = $("#dropzone");
    ["dragover", "dragenter"].forEach((ev) => dz.addEventListener(ev, (e) => { e.preventDefault(); dz.classList.add("drag"); }));
    ["dragleave", "drop"].forEach((ev) => dz.addEventListener(ev, (e) => { e.preventDefault(); dz.classList.remove("drag"); }));
    dz.addEventListener("drop", (e) => uploadData(e.dataTransfer.files));
    $("#mapping-upload").addEventListener("click", uploadMapping);
    $("#clear-data").addEventListener("click", clearData);

    $("#to-split").addEventListener("click", () => setStep(2));
    $("#split-enabled").addEventListener("change", (e) => { SPLIT.enabled = e.target.value === "yes"; renderSplit(); });
    $("#split-preview").addEventListener("click", previewSplit);
    $("#split-reset").addEventListener("click", () => { SPLIT.rules = SPLIT.defaults.map((r) => Object.assign({}, r)); drawSplitRules(); $("#split-preview-result").innerHTML = ""; });
    $("#pick-all").addEventListener("click", () => { DATA_FILES.forEach((f) => PICK.add(f.path)); renderPick(); });
    $("#pick-none").addEventListener("click", () => { PICK.clear(); renderPick(); });
    $("#run-validate").addEventListener("click", runValidate);
    $("#to-run").addEventListener("click", () => setStep(4));
    $("#run-go").addEventListener("click", runCalc);
    $("#cmp-add").addEventListener("click", () => { CMP_CONFIG.rows.push({ left: "", right: "", type: "numeric" }); renderCompare(); });
    $("#cmp-run").addEventListener("click", runCompare);
    $("#tr-go").addEventListener("click", loadRowJourney);
    $$("[data-goto]").forEach((b) => b.addEventListener("click", () => setStep(+b.dataset.goto)));
    $$("#stepper li").forEach((li) => li.addEventListener("click", () => setStep(+li.dataset.step)));

    $("#d-close").addEventListener("click", closeDrawer);
    $("#drawer-scrim").addEventListener("click", closeDrawer);
    $$("#d-tabs button").forEach((b) => b.addEventListener("click", () => showTab(b.dataset.tab)));
    document.addEventListener("keydown", (e) => { if (e.key === "Escape") closeDrawer(); });

    window.addEventListener("hashchange", route);
    window.addEventListener("load", route);
    call("/api/whoami").then((r) => { const d = D(r); if (d && d.app_title) $("#brand-text").textContent = d.app_title; if (d && d.user) $("#who").textContent = d.user; }).catch(() => {});
    route();
  }

  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", init);
  else init();
})();
