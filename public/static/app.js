const $ = (id) => document.getElementById(id);
const today = new Date().toISOString().slice(0, 10);
const month = today.slice(0, 7);

const tableHeaders = {
  salespeople: ["业务员", "今日销售金额", "本月销售金额", "巅峰目标", "每日目标", "完成率"],
  products: ["排名", "商品编码", "商品名称", "销售数量", "销售金额", "销售业务员数"],
  tasks: ["月份", "业务员", "商品编码", "商品名称", "当日销量", "月累计销量", "月任务数量", "完成率", "剩余任务数量", "是否达标"],
  taskRank: ["业务员", "金砖商品目标", "金砖商品每日目标", "每日完成件数", "今日完成率", "累计完成", "累计完成率"]
};

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
  return value ?? "";
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

function renderTable(el, headers, rows) {
  if (!rows || rows.length === 0) {
    el.innerHTML = '<div class="empty">暂无数据</div>';
    return;
  }
  const head = headers.map((h) => `<th>${h}</th>`).join("");
  const body = rows.map((row) => `<tr>${headers.map((h) => `<td>${fmtCell(h, row[h])}</td>`).join("")}</tr>`).join("");
  el.innerHTML = `<table><thead><tr>${head}</tr></thead><tbody>${body}</tbody></table>`;
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
    if (page === "base") loadMaintenance();
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
    ["今日总销售金额", `¥${fmtNumber(data.overview.total_amount)}`],
    ["今日总销售数量", fmtNumber(data.overview.total_qty)],
    ["今日成交客户数", fmtNumber(data.overview.customer_count)],
    ["今日销售商品数", fmtNumber(data.overview.product_count)],
    ["今日有销售业务员", fmtNumber(data.overview.salesperson_count)]
  ];
  $("overview").innerHTML = overview.map(([label, value]) => `<div class="metric"><div class="label">${label}</div><div class="value">${value}</div></div>`).join("");
  renderTable($("dashUnmet"), tableHeaders.taskRank, data.unmet);
}

async function loadSalespeople() {
  const q = params({ date: $("spDate").value, salesperson: $("spName").value.trim() });
  renderTable($("spTable"), tableHeaders.salespeople, await api(`/api/sales-ranking?${q}`));
}

async function loadProducts() {
  const q = params({ period: $("pdPeriod").value, date: $("pdDate").value, product: $("pdKeyword").value.trim(), keyOnly: $("pdKeyOnly").checked ? "1" : "" });
  renderTable($("pdTable"), tableHeaders.products, await api(`/api/product-stats?${q}`));
}

async function loadTasks() {
  const q = taskQuery();
  renderTable($("taskTable"), tableHeaders.taskRank, await api(`/api/gold-rank?${q}`));
}

function taskQuery() {
  return params({
    month: $("taskMonth").value,
    date: $("taskDate").value,
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
    <table class="edit-table"><thead><tr><th>选择</th>${headers.map((h) => `<th>${h}</th>`).join("")}<th>操作</th></tr></thead><tbody>${rowsHtml}</tbody></table>
  `;
  $(id).querySelector(`[data-add5="${table}"]`).addEventListener("click", () => addEditRows(table, 5));
  $(id).querySelector(`[data-save="${table}"]`).addEventListener("click", () => saveEditTable(table));
  $(id).querySelector(`[data-export="${table}"]`).addEventListener("click", () => download(`/api/export-maintenance/${table}`));
  $(id).querySelector(`[data-delete-selected="${table}"]`).addEventListener("click", () => deleteSelectedRows(id));
  $(id).querySelector(`[data-import="${table}"]`).addEventListener("change", (e) => importMaintenance(table, e.target.files[0]));
  const dropZone = $(id).querySelector(`[data-drop="${table}"]`);
  dropZone.addEventListener("dragover", (e) => { e.preventDefault(); dropZone.classList.add("dragging"); });
  dropZone.addEventListener("dragleave", () => dropZone.classList.remove("dragging"));
  dropZone.addEventListener("drop", (e) => {
    e.preventDefault();
    dropZone.classList.remove("dragging");
    importMaintenance(table, e.dataTransfer.files[0]);
  });
  $(id).querySelectorAll("[data-remove]").forEach((btn) => btn.addEventListener("click", () => btn.closest("tr").remove()));
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
  editShell("productsEdit", "products", ["商品编码", "商品名称", "商品规格", "商品分类", "是否重点商品", "状态", "待确认"], html);
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
  editShell("goldTargetsEdit", "goldProducts", ["商品编码", "商品名称", "金砖商品目标", "每日目标"], html);
}

function goldTargetsRow(r = {}) {
  return `<tr>
    <td><input type="checkbox" data-select-row></td>
    <td><input data-field="code" value="${r.code || ""}"></td>
    <td><input data-field="name" value="${r.name || ""}"></td>
    <td><input type="number" min="0" step="1" data-field="gold_target" value="${r.gold_target || 0}"></td>
    <td><input type="number" min="0" step="0.01" data-field="gold_daily_target" value="${r.gold_daily_target || 0}"></td>
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

function addEditRows(table, count = 1) {
  const map = { salespeople: salespeopleRow, products: productsRow, salesTargets: salesTargetsRow, goldProducts: goldTargetsRow };
  const panel = panelIdFor(table);
  for (let i = 0; i < count; i += 1) {
    $(panel).querySelector("tbody").insertAdjacentHTML("beforeend", map[table]({}));
    $(panel).querySelector("tbody tr:last-child [data-remove]").addEventListener("click", (e) => e.target.closest("tr").remove());
  }
}

function deleteSelectedRows(panelId) {
  const rows = Array.from($(panelId).querySelectorAll("tbody tr")).filter((tr) => tr.querySelector("[data-select-row]")?.checked);
  if (rows.length === 0) {
    alert("请先勾选需要删除的行。");
    return;
  }
  if (!confirm(`确认删除选中的 ${rows.length} 行？删除后需要点击“保存”才会写入数据库。`)) return;
  rows.forEach((tr) => tr.remove());
}

async function importMaintenance(table, file) {
  if (!file) return;
  const panel = panelIdFor(table);
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

async function saveEditTable(table) {
  const panel = panelIdFor(table);
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

function initEvents() {
  $("dashDate").value = today;
  $("spDate").value = today;
  $("pdDate").value = today;
  $("taskDate").value = today;
  $("taskMonth").value = month;

  $("dashDate").addEventListener("change", loadDashboard);
  $("uploadBtn").addEventListener("click", () => uploadSales());
  $("spSearch").addEventListener("click", loadSalespeople);
  $("pdSearch").addEventListener("click", loadProducts);
  $("taskSearch").addEventListener("click", loadTasks);

  $("spExport").addEventListener("click", () => download(`/api/export/sales-ranking?${params({ date: $("spDate").value, salesperson: $("spName").value.trim() })}`));
  $("pdExport").addEventListener("click", () => download(`/api/export/products?${params({ period: $("pdPeriod").value, date: $("pdDate").value, product: $("pdKeyword").value.trim(), keyOnly: $("pdKeyOnly").checked ? "1" : "" })}`));
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
