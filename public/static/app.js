const $ = (id) => document.getElementById(id);
const today = new Date().toISOString().slice(0, 10);
const month = today.slice(0, 7);

const tableHeaders = {
  salespeople: ["业务员", "今日销售金额", "本月销售金额", "巅峰目标", "每日目标", "完成率"],
  products: ["排名", "商品编码", "商品名称", "销售数量", "销售金额", "销售业务员数"],
  tasks: ["月份", "业务员", "商品编码", "商品名称", "当日销量", "月累计销量", "月任务数量", "完成率", "剩余任务数量", "是否达标"],
  taskRank: ["业务员", "金砖商品目标", "金砖商品每日目标", "每日完成件数", "今日完成率", "累计完成", "累计完成率"]
};
const PAGE_SIZE = 50;
const tableStates = {};
const editTableStates = {};

function escapeHtml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function fmtNumber(value) {
  const num = Number(value || 0);
  return Number.isInteger(num) ? num.toLocaleString("zh-CN") : num.toLocaleString("zh-CN", { maximumFractionDigits: 2 });
}

function fmtCell(header, value) {
  if (header.includes("完成率")) return `${(Number(value || 0) * 100).toFixed(1)}%`;
  if (header === "是否达标") {
    const cls = value === "已达标" ? "good" : "warn";
    return `<span class="tag ${cls}">${value}</span>`;
  }
  if (typeof value === "number") return fmtNumber(value);
  return escapeHtml(value ?? "");
}

async function api(url, options) {
  const res = await fetch(url, options);
  const contentType = res.headers.get("content-type") || "";
  if (!res.ok) {
    if (contentType.includes("json")) {
      const data = await res.json();
      const error = new Error(data.error || data.message || "请求失败");
      error.data = data;
      error.status = res.status;
      throw error;
    }
    throw new Error("请求失败");
  }
  return contentType.includes("json") ? res.json() : res.blob();
}

function rawCellText(row, header) {
  const value = row[header];
  if (header.includes("完成率")) return `${(Number(value || 0) * 100).toFixed(1)}%`;
  return String(value ?? "");
}

function renderTable(el, headers, rows) {
  if (!el) return;
  const id = el.id || `table-${Math.random().toString(36).slice(2)}`;
  if (!el.id) el.id = id;
  const state = tableStates[id] || { page: 1, filters: {} };
  tableStates[id] = state;
  if (!rows || rows.length === 0) {
    el.innerHTML = '<div class="empty">暂无数据</div>';
    return;
  }
  const filtered = rows.filter((row) => headers.every((header) => {
    const keyword = String(state.filters[header] || "").trim().toLowerCase();
    if (!keyword) return true;
    return rawCellText(row, header).toLowerCase().includes(keyword);
  }));
  const pageCount = Math.max(1, Math.ceil(filtered.length / PAGE_SIZE));
  state.page = Math.min(Math.max(1, state.page || 1), pageCount);
  const start = (state.page - 1) * PAGE_SIZE;
  const pageRows = filtered.slice(start, start + PAGE_SIZE);
  const head = headers.map((h) => `<th>${escapeHtml(h)}</th>`).join("");
  const search = headers.map((h) => `<th><input class="column-filter" data-field-filter="${escapeHtml(h)}" placeholder="搜索${escapeHtml(h)}" value="${escapeHtml(state.filters[h] || "")}"></th>`).join("");
  const body = pageRows.map((row) => `<tr>${headers.map((h) => `<td>${fmtCell(h, row[h])}</td>`).join("")}</tr>`).join("");
  const from = filtered.length ? start + 1 : 0;
  const to = Math.min(start + PAGE_SIZE, filtered.length);
  el.innerHTML = `
    <div class="table-meta">共 ${filtered.length} 条，显示 ${from}-${to} 条，每页最多 ${PAGE_SIZE} 条</div>
    <table><thead><tr>${head}</tr><tr class="filter-row">${search}</tr></thead><tbody>${body || `<tr><td colspan="${headers.length}"><div class="empty">暂无匹配数据</div></td></tr>`}</tbody></table>
    <div class="pager">
      <button data-page-prev ${state.page <= 1 ? "disabled" : ""}>上一页</button>
      <span>第 ${state.page} / ${pageCount} 页</span>
      <button data-page-next ${state.page >= pageCount ? "disabled" : ""}>下一页</button>
    </div>
  `;
  el.querySelectorAll("[data-field-filter]").forEach((input) => {
    input.addEventListener("input", () => {
      const field = input.dataset.fieldFilter;
      state.filters[field] = input.value;
      state.page = 1;
      renderTable(el, headers, rows);
      const nextInput = el.querySelector(`[data-field-filter="${CSS.escape(field)}"]`);
      if (nextInput) {
        nextInput.focus();
        nextInput.setSelectionRange(nextInput.value.length, nextInput.value.length);
      }
    });
  });
  const prev = el.querySelector("[data-page-prev]");
  const next = el.querySelector("[data-page-next]");
  if (prev) prev.addEventListener("click", () => { state.page -= 1; renderTable(el, headers, rows); });
  if (next) next.addEventListener("click", () => { state.page += 1; renderTable(el, headers, rows); });
}

function params(obj) {
  const p = new URLSearchParams();
  Object.entries(obj).forEach(([k, v]) => {
    if (v !== undefined && v !== null && v !== "") p.set(k, v);
  });
  return p.toString();
}

function download(url) {
  window.location.href = url;
}

function initNavigation() {
  const goPage = (page) => {
    document.querySelectorAll(".nav, .page").forEach((x) => x.classList.remove("active"));
    const nav = document.querySelector(`.nav[data-page="${page}"]`);
    if (nav) nav.classList.add("active");
    $(page).classList.add("active");
    if (["base", "salesTargets", "goldProducts"].includes(page)) loadMaintenance();
    if (page === "database") loadDatabaseStatus();
  };
  document.querySelectorAll(".nav").forEach((btn) => {
    btn.addEventListener("click", () => {
      goPage(btn.dataset.page);
    });
  });
  document.querySelectorAll(".goto").forEach((btn) => {
    btn.addEventListener("click", () => goPage(btn.dataset.page));
  });
  document.querySelectorAll(".tab").forEach((btn) => {
    btn.addEventListener("click", () => {
      document.querySelectorAll(".tab, .tab-page").forEach((x) => x.classList.remove("active"));
      btn.classList.add("active");
      $(btn.dataset.tab).classList.add("active");
    });
  });
}

async function loadDashboard() {
  const data = await api(`/api/dashboard?${params({ date: $("dashDate").value })}`);
  const overview = [
    ["今日销售额", `¥${fmtNumber(data.overview.total_amount)}`],
    ["本月销售额", `¥${fmtNumber(data.overview.month_amount)}`],
    ["本月完成率", `${(Number(data.overview.month_completion_rate || 0) * 100).toFixed(1)}%`]
  ];
  $("overview").innerHTML = overview.map(([label, value]) => `<div class="metric"><div class="label">${label}</div><div class="value">${value}</div></div>`).join("");
  renderTable($("dashSalesTop"), tableHeaders.salespeople, data.salespeopleRank);
  renderTable($("dashGoldTop"), tableHeaders.taskRank, data.taskRank);
}

async function loadSalespeople() {
  const endDate = $("spEndDate").value || today;
  const q = params({ date: endDate, startDate: $("spStartDate").value, endDate, salesperson: $("spName").value.trim() });
  renderTable($("spTable"), tableHeaders.salespeople, await api(`/api/sales-ranking?${q}`));
}

async function loadProducts() {
  const endDate = $("pdEndDate").value || today;
  const q = params({ period: $("pdPeriod").value, date: endDate, startDate: $("pdStartDate").value, endDate, product: $("pdKeyword").value.trim(), keyOnly: $("pdKeyOnly").checked ? "1" : "" });
  renderTable($("pdTable"), tableHeaders.products, await api(`/api/product-stats?${q}`));
}

async function loadTasks() {
  const q = taskQuery();
  renderTable($("taskTable"), tableHeaders.taskRank, await api(`/api/gold-rank?${q}`));
}

async function loadDatabaseStatus() {
  const data = await api("/api/db-status");
  const overview = [
    ["数据库连接状态", data.connectionStatus || (data.ok ? "已连接" : "异常")],
    ["销售记录数", fmtNumber(data.salesCount)],
    ["业务员数量", fmtNumber(data.salespersonCount)],
    ["商品数量", fmtNumber(data.productCount)]
  ];
  $("dbOverview").innerHTML = overview.map(([label, value]) => `<div class="metric"><div class="label">${label}</div><div class="value">${value}</div></div>`).join("");
}

async function runDatabaseAction(action) {
  const config = {
    clear: {
      url: "/api/clear-sales-data",
      title: "清空销售数据",
      first: "确认清空销售数据？此操作只删除销售明细表，不删除业务员资料、商品资料、目标任务和管理员账号。",
      second: "请再次确认：销售明细删除后不可恢复，数据库表结构会保留。"
    },
    reinit: {
      url: "/api/reinitialize-database",
      title: "重新初始化数据库",
      first: "确认重新初始化数据库？此操作会清空所有业务数据，但保留表结构、系统配置和管理员账号。",
      second: "请再次确认：业务数据删除后不可恢复。"
    }
  }[action];
  if (!confirm(config.first)) return;
  if (!confirm(config.second)) return;
  $("dbActionResult").innerHTML = `<span class="hint">${config.title}执行中...</span>`;
  try {
    const data = await api(config.url, { method: "POST" });
    $("dbActionResult").innerHTML = `<span class="ok">${data.message}</span>`;
    await loadDatabaseStatus();
    loadDashboard();
    loadSalespeople();
    loadProducts();
    loadTasks();
  } catch (err) {
    $("dbActionResult").innerHTML = `<span class="error">${err.message}</span>`;
  }
}

function taskQuery() {
  const endDate = $("taskEndDate").value || today;
  return params({
    month: $("taskMonth").value,
    date: endDate,
    startDate: $("taskStartDate").value,
    endDate,
    salesperson: $("taskSalesperson").value.trim()
  });
}

async function uploadSales(mode = "") {
  const file = $("uploadFile").files[0];
  if (!file) {
    $("uploadResult").innerHTML = '<span class="error">请先选择 Excel 文件。</span>';
    return;
  }
  const form = new FormData();
  form.append("file", file);
  try {
    const q = mode ? `?mode=${mode}` : "";
    const data = await api(`/api/upload${q}`, { method: "POST", body: form });
    $("uploadResult").innerHTML = `<span class="ok">${data.message}</span><br>日期：${data.dates.join("、")}`;
    loadDashboard();
  } catch (err) {
    if (err.status === 409 && err.data && err.data.duplicate) {
      const dates = err.data.dates.map((d) => `${d.sale_date}（${d.c}条）`).join("、");
      const choice = confirm(`检测到这些日期已有数据：${dates}\n\n点击“确定”覆盖这些日期；点击“取消”则合并导入。`);
      uploadSales(choice ? "replace" : "merge");
      return;
    }
    const details = err.data && err.data.errors ? `<br>${err.data.errors.join("<br>")}` : "";
    $("uploadResult").innerHTML = `<span class="error">${err.message}</span>${details}`;
  }
}

let maintenanceCache = null;

async function loadMaintenance() {
  maintenanceCache = await api("/api/maintenance");
  renderSalespeopleEdit(maintenanceCache.salespeople);
  renderProductsEdit(maintenanceCache.products);
  renderSalesTargetsEdit(maintenanceCache.salesTargets);
  renderGoldTargetsEdit(maintenanceCache.goldTargets);
  renderPendingEdit(maintenanceCache.pendingItems);
}

function editShell(id, table, headers, rowsHtml) {
  const target = $(id);
  if (!target) return;
  $(id).innerHTML = `
    <div class="edit-actions drop-zone" data-drop="${table}">
      <a class="link-btn" href="/api/base-template/${table}" download>下载导入模板</a>
      <label class="link-btn file-action">导入Excel<input type="file" data-import="${table}" accept=".xlsx,.xls" hidden></label>
      <button data-export="${table}">导出当前数据</button>
      <button data-add5="${table}">批量新增5行</button>
      <button data-delete-selected="${table}" class="danger">批量删除选中</button>
      <button data-save="${table}">保存</button>
      <span class="hint">可拖拽 Excel 到此区域导入</span>
    </div>
    <div class="result" data-import-result="${table}"></div>
    <div class="table-meta" data-edit-meta></div>
    <table class="edit-table">
      <thead>
        <tr><th>选择</th>${headers.map((h) => `<th>${h}</th>`).join("")}<th>操作</th></tr>
        <tr class="filter-row"><th></th>${headers.map((h, i) => `<th><input class="column-filter" data-edit-filter="${i}" placeholder="搜索${escapeHtml(h)}"></th>`).join("")}<th></th></tr>
      </thead>
      <tbody>${rowsHtml}</tbody>
    </table>
    <div class="pager" data-edit-pager></div>
  `;
  $(id).querySelector(`[data-add5="${table}"]`).addEventListener("click", () => addEditRowsIn(id, table, 5));
  $(id).querySelector(`[data-save="${table}"]`).addEventListener("click", () => saveEditTableFrom(id, table));
  $(id).querySelector(`[data-export="${table}"]`).addEventListener("click", () => download(`/api/export-maintenance/${table}`));
  $(id).querySelector(`[data-delete-selected="${table}"]`).addEventListener("click", () => deleteSelectedRows(id));
  $(id).querySelector(`[data-import="${table}"]`).addEventListener("change", (e) => importMaintenanceTo(id, table, e.target.files[0]));
  const dropZone = $(id).querySelector(`[data-drop="${table}"]`);
  dropZone.addEventListener("dragover", (e) => { e.preventDefault(); dropZone.classList.add("dragging"); });
  dropZone.addEventListener("dragleave", () => dropZone.classList.remove("dragging"));
  dropZone.addEventListener("drop", (e) => {
    e.preventDefault();
    dropZone.classList.remove("dragging");
    importMaintenanceTo(id, table, e.dataTransfer.files[0]);
  });
  $(id).querySelectorAll("[data-remove]").forEach((btn) => btn.addEventListener("click", () => {
    btn.closest("tr").remove();
    applyEditTableFilters(id);
  }));
  $(id).querySelectorAll("[data-edit-filter]").forEach((input) => {
    input.addEventListener("input", () => {
      const state = editTableStates[id] || { page: 1, filters: {} };
      state.filters[input.dataset.editFilter] = input.value;
      state.page = 1;
      editTableStates[id] = state;
      applyEditTableFilters(id);
      const nextInput = $(id).querySelector(`[data-edit-filter="${input.dataset.editFilter}"]`);
      if (nextInput) {
        nextInput.focus();
        nextInput.setSelectionRange(nextInput.value.length, nextInput.value.length);
      }
    });
  });
  applyEditTableFilters(id);
}

function editCellText(tr, index) {
  const td = tr.children[index + 1];
  if (!td) return "";
  const field = td.querySelector("input, select");
  return field ? field.value : td.textContent;
}

function applyEditTableFilters(panelId) {
  const panel = $(panelId);
  if (!panel) return;
  const state = editTableStates[panelId] || { page: 1, filters: {} };
  editTableStates[panelId] = state;
  const rows = Array.from(panel.querySelectorAll("tbody tr"));
  const filtered = rows.filter((tr) => Object.entries(state.filters).every(([index, keyword]) => {
    const value = String(keyword || "").trim().toLowerCase();
    if (!value) return true;
    return editCellText(tr, Number(index)).toLowerCase().includes(value);
  }));
  const pageCount = Math.max(1, Math.ceil(filtered.length / PAGE_SIZE));
  state.page = Math.min(Math.max(1, state.page || 1), pageCount);
  const start = (state.page - 1) * PAGE_SIZE;
  const visible = new Set(filtered.slice(start, start + PAGE_SIZE));
  rows.forEach((tr) => {
    tr.style.display = visible.has(tr) ? "" : "none";
  });
  const meta = panel.querySelector("[data-edit-meta]");
  const pager = panel.querySelector("[data-edit-pager]");
  const from = filtered.length ? start + 1 : 0;
  const to = Math.min(start + PAGE_SIZE, filtered.length);
  if (meta) meta.textContent = `共 ${filtered.length} 条，显示 ${from}-${to} 条，每页最多 ${PAGE_SIZE} 条`;
  if (pager) {
    pager.innerHTML = `
      <button data-edit-prev ${state.page <= 1 ? "disabled" : ""}>上一页</button>
      <span>第 ${state.page} / ${pageCount} 页</span>
      <button data-edit-next ${state.page >= pageCount ? "disabled" : ""}>下一页</button>
    `;
    pager.querySelector("[data-edit-prev]")?.addEventListener("click", () => { state.page -= 1; applyEditTableFilters(panelId); });
    pager.querySelector("[data-edit-next]")?.addEventListener("click", () => { state.page += 1; applyEditTableFilters(panelId); });
  }
}

function renderSalespeopleEdit(rows) {
  const html = rows.map((r) => salespeopleRow(r)).join("");
  editShell("salespeopleEdit", "salespeople", ["业务员编码", "业务员姓名", "所属区域", "所属部门", "入职日期", "状态", "待确认"], html);
}

function salespeopleRow(r = {}) {
  return `<tr>
    <td><input type="checkbox" data-select-row></td>
    <td><input data-field="code" value="${r.code || ""}"></td>
    <td><input data-field="name" value="${r.name || ""}"></td>
    <td><input data-field="region" value="${r.region || ""}"></td>
    <td><input data-field="department" value="${r.department || ""}"></td>
    <td><input type="date" data-field="entry_date" value="${r.entry_date || ""}"></td>
    <td><select data-field="active"><option value="1" ${r.active !== 0 ? "selected" : ""}>启用</option><option value="0" ${r.active === 0 ? "selected" : ""}>停用</option></select></td>
    <td><select data-field="pending_confirm"><option value="1" ${r.pending_confirm === 1 ? "selected" : ""}>是</option><option value="0" ${r.pending_confirm !== 1 ? "selected" : ""}>否</option></select></td>
    <td><button data-remove>删除</button></td>
  </tr>`;
}

function renderProductsEdit(rows) {
  const html = rows.map((r) => productsRow(r)).join("");
  editShell("productsEdit", "products", ["商品编码", "商品名称", "商品规格", "商品分类", "是否金砖商品", "状态", "待确认"], html);
}

function productsRow(r = {}) {
  return `<tr>
    <td><input type="checkbox" data-select-row></td>
    <td><input data-field="code" value="${r.code || ""}"></td>
    <td><input data-field="name" value="${r.name || ""}"></td>
    <td><input data-field="spec" value="${r.spec || ""}"></td>
    <td><input data-field="category" value="${r.category || ""}"></td>
    <td><select data-field="is_key"><option value="1" ${r.is_key === 1 ? "selected" : ""}>是</option><option value="0" ${r.is_key !== 1 ? "selected" : ""}>否</option></select></td>
    <td><select data-field="status"><option value="启用" ${(r.status || "启用") === "启用" ? "selected" : ""}>启用</option><option value="停用" ${r.status === "停用" ? "selected" : ""}>停用</option></select></td>
    <td><select data-field="pending_confirm"><option value="1" ${r.pending_confirm === 1 ? "selected" : ""}>是</option><option value="0" ${r.pending_confirm !== 1 ? "selected" : ""}>否</option></select></td>
    <td><button data-remove>删除</button></td>
  </tr>`;
}

function renderSalesTargetsEdit(rows) {
  const html = rows.map((r) => salesTargetsRow(r)).join("");
  editShell("salesTargetsEdit", "salesTargets", ["月份", "业务员", "巅峰目标", "每日目标"], html);
  editShell("salesTargetsPageEdit", "salesTargets", ["月份", "业务员", "巅峰目标", "每日目标"], html);
}

function salesTargetsRow(r = {}) {
  return `<tr>
    <td><input type="checkbox" data-select-row></td>
    <td><input type="month" data-field="month" value="${r.month || month}"></td>
    <td><input data-field="salesperson_name" value="${r.salesperson_name || ""}"></td>
    <td><input type="number" min="0" step="1" data-field="peak_target" value="${r.peak_target || 0}"></td>
    <td><input type="number" min="0" step="1" data-field="daily_target" value="${r.daily_target || 0}"></td>
    <td><button data-remove>删除</button></td>
  </tr>`;
}

function renderGoldTargetsEdit(rows) {
  const html = rows.map((r) => goldTargetsRow(r)).join("");
  editShell("goldTargetsEdit", "goldProducts", ["商品编码", "商品名称", "是否金砖商品"], html);
  editShell("goldProductsPageEdit", "goldProducts", ["商品编码", "商品名称", "是否金砖商品"], html);
}

function goldTargetsRow(r = {}) {
  return `<tr>
    <td><input type="checkbox" data-select-row></td>
    <td><input data-field="code" value="${r.code || ""}"></td>
    <td><input data-field="name" value="${r.name || ""}"></td>
    <td><select data-field="is_key"><option value="1" ${r.is_key === 1 ? "selected" : ""}>是</option><option value="0" ${r.is_key !== 1 ? "selected" : ""}>否</option></select></td>
    <td><button data-remove>删除</button></td>
  </tr>`;
}

function renderPendingEdit(rows) {
  const pending = rows.filter((r) => r.status === "待确认");
  if (pending.length === 0) {
    $("pendingEdit").innerHTML = '<div class="empty">暂无待确认数据</div>';
    return;
  }
  const body = pending.map((r) => `
    <tr>
      <td>${r.item_type}</td>
      <td><input data-pending-key="${r.id}" value="${r.ref_key}"></td>
      <td><input data-pending-name="${r.id}" value="${r.name}"></td>
      <td>${r.created_at}</td>
      <td>
        <button data-modify="${r.id}">修改</button>
        <button data-confirm="${r.id}">确认</button>
        <button class="danger" data-delete-pending="${r.id}">删除</button>
      </td>
    </tr>
  `).join("");
  $("pendingEdit").innerHTML = `<table><thead><tr><th>类型</th><th>编码/姓名</th><th>名称</th><th>发现时间</th><th>操作</th></tr></thead><tbody>${body}</tbody></table>`;
  $("pendingEdit").querySelectorAll("[data-modify]").forEach((btn) => btn.addEventListener("click", () => pendingAction(btn.dataset.modify, "modify")));
  $("pendingEdit").querySelectorAll("[data-confirm]").forEach((btn) => btn.addEventListener("click", () => pendingAction(btn.dataset.confirm, "confirm")));
  $("pendingEdit").querySelectorAll("[data-delete-pending]").forEach((btn) => btn.addEventListener("click", () => pendingAction(btn.dataset.deletePending, "delete")));
}

async function pendingAction(id, action) {
  const text = action === "confirm" ? "确认这条数据？" : action === "modify" ? "保存这条待确认数据的修改？" : "删除这条待确认数据及对应订单/档案？";
  if (!confirm(text)) return;
  const payload = { id, action };
  if (action === "modify") {
    payload.ref_key = $(`pendingEdit`).querySelector(`[data-pending-key="${id}"]`).value.trim();
    payload.name = $(`pendingEdit`).querySelector(`[data-pending-name="${id}"]`).value.trim();
  }
  const data = await api("/api/pending/action", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload)
  });
  alert(data.message);
  loadMaintenance();
}

function panelIdFor(table) {
  return table === "salespeople" ? "salespeopleEdit" : table === "products" ? "productsEdit" : table === "salesTargets" ? "salesTargetsEdit" : "goldTargetsEdit";
}

function addEditRowsIn(panel, table, count = 1) {
  const map = { salespeople: salespeopleRow, products: productsRow, salesTargets: salesTargetsRow, goldProducts: goldTargetsRow };
  for (let i = 0; i < count; i += 1) {
    $(panel).querySelector("tbody").insertAdjacentHTML("beforeend", map[table]({}));
    $(panel).querySelector("tbody tr:last-child [data-remove]").addEventListener("click", (e) => {
      e.target.closest("tr").remove();
      applyEditTableFilters(panel);
    });
  }
  applyEditTableFilters(panel);
}

function addEditRows(table, count = 1) {
  addEditRowsIn(panelIdFor(table), table, count);
}

function deleteSelectedRows(panelId) {
  const rows = Array.from($(panelId).querySelectorAll("tbody tr")).filter((tr) => tr.querySelector("[data-select-row]")?.checked);
  if (rows.length === 0) {
    alert("请先勾选需要删除的行。");
    return;
  }
  if (!confirm(`确认删除选中的 ${rows.length} 行？删除后需要点击“保存”才会写入数据库。`)) return;
  rows.forEach((tr) => tr.remove());
  applyEditTableFilters(panelId);
}

async function importMaintenanceTo(panel, table, file) {
  if (!file) return;
  const result = $(panel).querySelector(`[data-import-result="${table}"]`);
  const form = new FormData();
  form.append("file", file);
  try {
    const data = await api(`/api/import-maintenance/${table}`, { method: "POST", body: form });
    const errorLink = data.errorUrl ? `，<a href="${data.errorUrl}" download>下载错误明细</a>` : "";
    result.innerHTML = `<span class="ok">导入完成：成功 ${data.success} 条，失败 ${data.failed} 条</span>${errorLink}`;
    loadMaintenance();
  } catch (err) {
    result.innerHTML = `<span class="error">${err.message}</span>`;
  }
}

async function importMaintenance(table, file) {
  return importMaintenanceTo(panelIdFor(table), table, file);
}

async function saveEditTableFrom(panel, table) {
  const records = Array.from($(panel).querySelectorAll("tbody tr")).map((tr) => {
    const obj = {};
    tr.querySelectorAll("[data-field]").forEach((input) => {
      let value = input.value;
      if (["active", "is_key", "pending_confirm"].includes(input.dataset.field)) value = value === "1";
      obj[input.dataset.field] = value;
    });
    return obj;
  });
  const data = await api("/api/maintenance", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ table, records })
  });
  alert(data.message);
  loadMaintenance();
}

async function saveEditTable(table) {
  return saveEditTableFrom(panelIdFor(table), table);
}

function initEvents() {
  $("dashDate").value = today;
  $("spStartDate").value = today;
  $("spEndDate").value = today;
  $("pdStartDate").value = today;
  $("pdEndDate").value = today;
  $("taskStartDate").value = today;
  $("taskEndDate").value = today;
  $("taskMonth").value = month;

  $("dashDate").addEventListener("change", loadDashboard);
  $("uploadBtn").addEventListener("click", () => uploadSales());
  $("spSearch").addEventListener("click", loadSalespeople);
  $("pdSearch").addEventListener("click", loadProducts);
  $("taskSearch").addEventListener("click", loadTasks);
  $("dbRefresh").addEventListener("click", loadDatabaseStatus);
  $("clearSalesData").addEventListener("click", () => runDatabaseAction("clear"));
  $("reinitDatabase").addEventListener("click", () => runDatabaseAction("reinit"));

  $("spExport").addEventListener("click", () => download(`/api/export/sales-ranking?${params({ date: $("spEndDate").value || today, startDate: $("spStartDate").value, endDate: $("spEndDate").value, salesperson: $("spName").value.trim() })}`));
  $("pdExport").addEventListener("click", () => download(`/api/export/products?${params({ period: $("pdPeriod").value, date: $("pdEndDate").value || today, startDate: $("pdStartDate").value, endDate: $("pdEndDate").value, product: $("pdKeyword").value.trim(), keyOnly: $("pdKeyOnly").checked ? "1" : "" })}`));
  $("taskExport").addEventListener("click", () => download(`/api/export/gold-rank?${taskQuery()}`));

  $("resetDemo").addEventListener("click", async () => {
    if (!confirm("确认重置为系统模拟数据？当前销售明细和维护数据会被覆盖。")) return;
    const data = await api("/api/reset-demo", { method: "POST" });
    alert(data.message);
    loadMaintenance();
    loadDashboard();
  });
}

initNavigation();
initEvents();
const initialPage = new URLSearchParams(window.location.search).get("page");
if (initialPage && document.querySelector(`.nav[data-page="${initialPage}"]`)) {
  document.querySelector(`.nav[data-page="${initialPage}"]`).click();
}
loadDashboard();
loadSalespeople();
loadProducts();
loadTasks();
loadDatabaseStatus();
