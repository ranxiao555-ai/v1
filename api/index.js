const { Pool } = require("pg");
const formidable = require("formidable");
const XLSX = require("xlsx");

module.exports.config = { api: { bodyParser: false } };

let pool;

function getDatabaseUrl() {
  return process.env.DATABASE_URL ||
    process.env.POSTGRES_URL ||
    process.env.POSTGRES_URL_NON_POOLING ||
    process.env.POSTGRES_PRISMA_URL ||
    "";
}

function getDatabaseUrlSource() {
  for (const key of ["DATABASE_URL", "POSTGRES_URL", "POSTGRES_URL_NON_POOLING", "POSTGRES_PRISMA_URL"]) {
    if (process.env[key]) return key;
  }
  return "";
}

function getPool() {
  const connectionString = getDatabaseUrl();
  if (!connectionString) {
    throw new Error("缺少 Neon PostgreSQL 连接字符串。请在 Vercel 环境变量中配置 DATABASE_URL，或使用 Neon/Vercel 集成自动生成的 POSTGRES_URL。");
  }
  if (!pool) {
    pool = new Pool({
      connectionString,
      ssl: process.env.PGSSLMODE === "disable" ? false : { rejectUnauthorized: false }
    });
  }
  return pool;
}

const SALES_HEADERS = ["日期", "业务员", "客户名称", "商品编码", "商品名称", "销售数量", "单价", "销售金额"];
const BASE_CONFIG = {
  salespeople: { title: "业务员基础资料", headers: ["业务员编码", "业务员姓名", "所属区域", "所属部门", "入职日期", "状态"] },
  products: { title: "商品基础资料", headers: ["商品编码", "商品名称", "商品规格", "商品分类", "是否重点商品", "状态"] },
  goldProducts: { title: "重点商品维护", headers: ["商品编码", "商品名称", "金砖商品目标", "每日目标"] },
  salesTargets: { title: "销售目标维护", headers: ["业务员", "巅峰目标", "每日目标"] }
};

let schemaReady = false;

async function ensureSchema() {
  if (schemaReady) return;
  if (!getDatabaseUrl()) {
    throw new Error("缺少 Neon PostgreSQL 连接字符串。请在 Vercel 环境变量中配置 DATABASE_URL，或使用 Neon/Vercel 集成自动生成的 POSTGRES_URL。");
  }
  await getPool().query(`
    CREATE TABLE IF NOT EXISTS salespeople (
      id SERIAL PRIMARY KEY,
      code TEXT,
      name TEXT NOT NULL UNIQUE,
      region TEXT NOT NULL DEFAULT '',
      department TEXT NOT NULL DEFAULT '',
      entry_date TEXT,
      active INTEGER NOT NULL DEFAULT 1,
      pending_confirm INTEGER NOT NULL DEFAULT 0
    );

    CREATE TABLE IF NOT EXISTS products (
      id SERIAL PRIMARY KEY,
      code TEXT NOT NULL UNIQUE,
      name TEXT NOT NULL,
      spec TEXT NOT NULL DEFAULT '',
      category TEXT NOT NULL DEFAULT '',
      is_key INTEGER NOT NULL DEFAULT 0,
      gold_target NUMERIC NOT NULL DEFAULT 0,
      gold_daily_target NUMERIC NOT NULL DEFAULT 0,
      status TEXT NOT NULL DEFAULT '启用',
      pending_confirm INTEGER NOT NULL DEFAULT 0
    );

    CREATE TABLE IF NOT EXISTS salesperson_targets (
      id SERIAL PRIMARY KEY,
      month TEXT NOT NULL,
      salesperson_name TEXT NOT NULL,
      peak_target NUMERIC NOT NULL DEFAULT 0,
      daily_target NUMERIC NOT NULL DEFAULT 0,
      UNIQUE(month, salesperson_name)
    );

    CREATE TABLE IF NOT EXISTS gold_targets (
      id SERIAL PRIMARY KEY,
      month TEXT NOT NULL,
      salesperson_name TEXT NOT NULL,
      target_qty NUMERIC NOT NULL DEFAULT 0,
      UNIQUE(month, salesperson_name)
    );

    CREATE TABLE IF NOT EXISTS targets (
      id SERIAL PRIMARY KEY,
      month TEXT NOT NULL,
      salesperson_name TEXT NOT NULL,
      product_code TEXT NOT NULL,
      product_name TEXT NOT NULL,
      target_qty NUMERIC NOT NULL DEFAULT 0,
      UNIQUE(month, salesperson_name, product_code)
    );

    CREATE TABLE IF NOT EXISTS sales (
      id SERIAL PRIMARY KEY,
      sale_date TEXT NOT NULL,
      salesperson_name TEXT NOT NULL,
      customer_name TEXT NOT NULL,
      product_code TEXT NOT NULL,
      product_name TEXT NOT NULL,
      quantity NUMERIC NOT NULL,
      unit_price NUMERIC NOT NULL,
      amount NUMERIC NOT NULL,
      uploaded_at TEXT NOT NULL,
      source_file TEXT
    );

    CREATE TABLE IF NOT EXISTS pending_items (
      id SERIAL PRIMARY KEY,
      item_type TEXT NOT NULL,
      ref_key TEXT NOT NULL,
      name TEXT NOT NULL,
      payload JSONB,
      status TEXT NOT NULL DEFAULT '待确认',
      created_at TEXT NOT NULL
    );

    CREATE INDEX IF NOT EXISTS idx_sales_sale_date ON sales(sale_date);
    CREATE INDEX IF NOT EXISTS idx_sales_salesperson ON sales(salesperson_name);
    CREATE INDEX IF NOT EXISTS idx_sales_product_code ON sales(product_code);
    CREATE INDEX IF NOT EXISTS idx_products_is_key ON products(is_key);
  `);
  schemaReady = true;
}

function json(res, status, data) {
  res.statusCode = status;
  res.setHeader("Content-Type", "application/json; charset=utf-8");
  res.setHeader("Cache-Control", "no-store");
  res.end(JSON.stringify(data));
}

function xlsxResponse(res, filename, headers, rows) {
  const wb = XLSX.utils.book_new();
  const data = [headers, ...rows.map((row) => headers.map((header) => row[header] ?? ""))];
  const ws = XLSX.utils.aoa_to_sheet(data);
  XLSX.utils.book_append_sheet(wb, ws, "数据");
  const buffer = XLSX.write(wb, { bookType: "xlsx", type: "buffer" });
  res.statusCode = 200;
  res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
  res.setHeader("Content-Disposition", `attachment; filename="download.xlsx"; filename*=UTF-8''${encodeURIComponent(filename)}`);
  res.setHeader("Cache-Control", "no-store");
  res.end(buffer);
}

function parseUrl(req) {
  return new URL(req.url, `http://${req.headers.host || "localhost"}`);
}

function monthDays(month) {
  const [year, m] = month.split("-").map(Number);
  return new Date(year, m, 0).getDate();
}

function monthStart(dateText) {
  return `${dateText.slice(0, 7)}-01`;
}

function toNum(value) {
  if (value === null || value === undefined || value === "") return 0;
  const n = Number(String(value).replace(/,/g, "").trim());
  return Number.isFinite(n) ? n : 0;
}

function parseDateCell(value) {
  if (!value) return "";
  if (value instanceof Date) return value.toISOString().slice(0, 10);
  if (typeof value === "number") {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (parsed) return `${parsed.y}-${String(parsed.m).padStart(2, "0")}-${String(parsed.d).padStart(2, "0")}`;
  }
  return String(value).trim().replace(/\//g, "-").slice(0, 10);
}

function truthy(value) {
  return ["是", "启用", "1", "true", "True", "Y", "y"].includes(String(value || "").trim());
}

function activeStatus(value) {
  return ["停用", "离职", "否", "0"].includes(String(value || "").trim()) ? 0 : 1;
}

async function parseJson(req) {
  const chunks = [];
  for await (const chunk of req) chunks.push(chunk);
  if (!chunks.length) return {};
  return JSON.parse(Buffer.concat(chunks).toString("utf8"));
}

function parseForm(req) {
  return new Promise((resolve, reject) => {
    const form = formidable.formidable
      ? formidable.formidable({ multiples: false, maxFileSize: 20 * 1024 * 1024 })
      : new formidable.IncomingForm({ multiples: false, maxFileSize: 20 * 1024 * 1024 });
    form.parse(req, (err, fields, files) => {
      if (err) reject(err);
      else resolve({ fields, files });
    });
  });
}

function workbookRows(buffer, headers) {
  const wb = XLSX.read(buffer, { type: "buffer", cellDates: true });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { defval: "" });
  const missing = headers.filter((header) => !Object.prototype.hasOwnProperty.call(rows[0] || {}, header));
  if (missing.length) throw new Error(`缺少必要字段：${missing.join("、")}`);
  return rows;
}

async function readUploadFile(req) {
  const { files } = await parseForm(req);
  const file = Array.isArray(files.file) ? files.file[0] : files.file;
  if (!file) throw new Error("请选择 Excel 文件");
  const fs = require("fs/promises");
  return { filename: file.originalFilename || file.newFilename || "upload.xlsx", buffer: await fs.readFile(file.filepath) };
}

async function query(sql, params = []) {
  const result = await getPool().query(sql, params);
  return result.rows;
}

async function addPending(client, itemType, refKey, name, payload) {
  const existing = await client.query(
    "SELECT id FROM pending_items WHERE item_type=$1 AND ref_key=$2 AND status='待确认'",
    [itemType, refKey]
  );
  if (existing.rowCount) return;
  await client.query(
    "INSERT INTO pending_items(item_type, ref_key, name, payload, status, created_at) VALUES ($1,$2,$3,$4,'待确认',$5)",
    [itemType, refKey, name, JSON.stringify(payload), new Date().toISOString()]
  );
}

async function apiDashboard(req, res, url) {
  const selected = url.searchParams.get("date") || new Date().toISOString().slice(0, 10);
  const overview = (await query(`
    SELECT ROUND(COALESCE(SUM(amount),0)::numeric,2)::float AS total_amount,
           COALESCE(SUM(quantity),0)::float AS total_qty,
           COUNT(DISTINCT customer_name)::int AS customer_count,
           COUNT(DISTINCT product_code)::int AS product_count,
           COUNT(DISTINCT salesperson_name)::int AS salesperson_count
    FROM sales WHERE sale_date=$1
  `, [selected]))[0];
  const gold = await goldRank(selected.slice(0, 7), selected, "");
  json(res, 200, {
    date: selected,
    overview,
    salespeopleRank: await salesRank(selected, ""),
    productsRank: await productStats("day", selected, "", "").then((r) => r.slice(0, 8)),
    taskRank: gold.slice(0, 8),
    unmet: gold.filter((r) => r["累计完成率"] < 1).slice(0, 10)
  });
}

async function salesRank(selected, salesperson) {
  const month = selected.slice(0, 7);
  const params = [selected, month, monthStart(selected), selected];
  let where = "";
  if (salesperson) {
    params.push(salesperson);
    where = `WHERE sp.name=$${params.length}`;
  }
  return query(`
    SELECT sp.name AS "业务员",
           ROUND(COALESCE(day_sales.amount,0)::numeric,2)::float AS "今日销售金额",
           ROUND(COALESCE(month_sales.amount,0)::numeric,2)::float AS "本月销售金额",
           COALESCE(st.peak_target,0)::float AS "巅峰目标",
           COALESCE(st.daily_target,0)::float AS "每日目标",
           CASE WHEN COALESCE(st.daily_target,0)=0 THEN 0 ELSE (COALESCE(day_sales.amount,0)/st.daily_target)::float END AS "完成率"
    FROM salespeople sp
    LEFT JOIN salesperson_targets st ON st.salesperson_name=sp.name AND st.month=$2
    LEFT JOIN (SELECT salesperson_name, SUM(amount) amount FROM sales WHERE sale_date=$1 GROUP BY salesperson_name) day_sales ON day_sales.salesperson_name=sp.name
    LEFT JOIN (SELECT salesperson_name, SUM(amount) amount FROM sales WHERE sale_date BETWEEN $3 AND $4 GROUP BY salesperson_name) month_sales ON month_sales.salesperson_name=sp.name
    ${where}
    ORDER BY "完成率" DESC, "今日销售金额" DESC
  `, params);
}

async function productStats(period, selected, keyword, keyOnly) {
  let start = selected, end = selected;
  if (period === "month") start = monthStart(selected);
  if (period === "week") {
    const d = new Date(selected);
    const day = d.getDay() || 7;
    d.setDate(d.getDate() - day + 1);
    start = d.toISOString().slice(0, 10);
  }
  const params = [start, end];
  let extra = "";
  if (keyword) {
    params.push(keyword, `%${keyword}%`);
    extra += ` AND (s.product_code=$${params.length - 1} OR s.product_name LIKE $${params.length})`;
  }
  if (keyOnly === "1") extra += " AND COALESCE(p.is_key,0)=1 AND COALESCE(p.status,'启用')='启用'";
  const rows = await query(`
    SELECT s.product_code AS "商品编码", s.product_name AS "商品名称",
           SUM(s.quantity)::float AS "销售数量",
           ROUND(SUM(s.amount)::numeric,2)::float AS "销售金额",
           COUNT(DISTINCT s.salesperson_name)::int AS "销售业务员数"
    FROM sales s
    LEFT JOIN products p ON p.code=s.product_code
    WHERE s.sale_date BETWEEN $1 AND $2 ${extra}
    GROUP BY s.product_code, s.product_name
    ORDER BY "销售数量" DESC
  `, params);
  return rows.map((row, index) => ({ ...row, "排名": index + 1 }));
}

async function goldRank(month, selected, salesperson) {
  const days = monthDays(month);
  const params = [selected, `${month}-01`, selected];
  let where = "";
  if (salesperson) {
    params.push(salesperson);
    where = `WHERE sp.name=$${params.length}`;
  }
  return query(`
    SELECT sp.name AS "业务员",
           COALESCE(gt.qty,0)::float AS "金砖商品目标",
           ROUND((COALESCE(gt.qty,0)/${days})::numeric,2)::float AS "金砖商品每日目标",
           COALESCE(day_sales.qty,0)::float AS "每日完成件数",
           CASE WHEN COALESCE(gt.qty,0)=0 THEN 0 ELSE (COALESCE(day_sales.qty,0)/(gt.qty/${days}))::float END AS "今日完成率",
           COALESCE(month_sales.qty,0)::float AS "累计完成",
           CASE WHEN COALESCE(gt.qty,0)=0 THEN 0 ELSE (COALESCE(month_sales.qty,0)/gt.qty)::float END AS "累计完成率"
    FROM salespeople sp
    LEFT JOIN (
      SELECT COALESCE(SUM(gold_target),0) qty
      FROM products
      WHERE is_key=1 AND COALESCE(status,'启用')='启用'
    ) gt ON TRUE
    LEFT JOIN (
      SELECT s.salesperson_name, SUM(s.quantity) qty FROM sales s
      JOIN products p ON p.code=s.product_code AND p.is_key=1 AND COALESCE(p.status,'启用')='启用'
      WHERE s.sale_date=$1 GROUP BY s.salesperson_name
    ) day_sales ON day_sales.salesperson_name=sp.name
    LEFT JOIN (
      SELECT s.salesperson_name, SUM(s.quantity) qty FROM sales s
      JOIN products p ON p.code=s.product_code AND p.is_key=1 AND COALESCE(p.status,'启用')='启用'
      WHERE s.sale_date BETWEEN $2 AND $3 GROUP BY s.salesperson_name
    ) month_sales ON month_sales.salesperson_name=sp.name
    ${where}
    ORDER BY "今日完成率" DESC, "累计完成率" DESC
  `, params);
}

async function dbCheck() {
  const tables = ["salespeople", "products", "salesperson_targets", "gold_targets", "targets", "sales", "pending_items"];
  const tableStatus = await query(`
    SELECT table_name
    FROM information_schema.tables
    WHERE table_schema='public' AND table_name = ANY($1)
  `, [tables]);
  const existing = tableStatus.map((row) => row.table_name);
  const counts = {};
  for (const table of tables) {
    if (existing.includes(table)) {
      counts[table] = Number((await query(`SELECT COUNT(*)::int AS count FROM ${table}`))[0].count);
    }
  }
  return {
    ok: true,
    database: "connected",
    provider: "Neon PostgreSQL",
    env: getDatabaseUrlSource(),
    schemaReady: tables.every((table) => existing.includes(table)),
    tables: existing,
    counts
  };
}

async function maintenanceData() {
  const [salespeople, products, salesTargets, goldTargets, pendingItems] = await Promise.all([
    query("SELECT id, code, name, region, department, entry_date, active, pending_confirm FROM salespeople ORDER BY id"),
    query("SELECT id, code, name, spec, category, is_key, gold_target, gold_daily_target, status, pending_confirm FROM products ORDER BY code"),
    query("SELECT id, month, salesperson_name, peak_target, daily_target FROM salesperson_targets ORDER BY month DESC, salesperson_name"),
    query("SELECT id, code, name, gold_target, gold_daily_target FROM products WHERE is_key=1 ORDER BY code"),
    query("SELECT id, item_type, ref_key, name, payload, status, created_at FROM pending_items ORDER BY id DESC")
  ]);
  return { salespeople, products, salesTargets, goldTargets, pendingItems, targets: [] };
}

async function saveMaintenance(req, res) {
  const body = await parseJson(req);
  const table = body.table;
  const records = body.records || [];
  const client = await getPool().connect();
  try {
    await client.query("BEGIN");
    if (table === "salespeople") {
      await client.query("DELETE FROM salespeople");
      for (const r of records) {
        if (!r.name) continue;
        await client.query("INSERT INTO salespeople(code,name,region,department,entry_date,active,pending_confirm) VALUES ($1,$2,$3,$4,$5,$6,$7)", [r.code || "", r.name || "", r.region || "", r.department || "", r.entry_date || null, r.active ? 1 : 0, r.pending_confirm ? 1 : 0]);
      }
    } else if (table === "products") {
      await client.query("DELETE FROM products");
      for (const r of records) {
        if (!r.code || !r.name) continue;
        await client.query("INSERT INTO products(code,name,spec,category,is_key,gold_target,gold_daily_target,status,pending_confirm) VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9)", [r.code, r.name, r.spec || "", r.category || "", r.is_key ? 1 : 0, toNum(r.gold_target), toNum(r.gold_daily_target), r.status || "启用", r.pending_confirm ? 1 : 0]);
      }
    } else if (table === "salesTargets") {
      await client.query("DELETE FROM salesperson_targets");
      for (const r of records) {
        if (!r.salesperson_name) continue;
        await client.query("INSERT INTO salesperson_targets(month,salesperson_name,peak_target,daily_target) VALUES ($1,$2,$3,$4)", [r.month || new Date().toISOString().slice(0, 7), r.salesperson_name, toNum(r.peak_target), toNum(r.daily_target)]);
      }
    } else if (table === "goldProducts") {
      for (const r of records) {
        if (!r.code || !r.name) continue;
        await client.query(`
          INSERT INTO products(code,name,is_key,gold_target,gold_daily_target,status) VALUES ($1,$2,1,$3,$4,'启用')
          ON CONFLICT(code) DO UPDATE SET name=EXCLUDED.name,is_key=1,gold_target=EXCLUDED.gold_target,gold_daily_target=EXCLUDED.gold_daily_target
        `, [r.code, r.name, toNum(r.gold_target), toNum(r.gold_daily_target)]);
      }
    }
    await client.query("COMMIT");
    json(res, 200, { ok: true, message: "保存成功" });
  } catch (e) {
    await client.query("ROLLBACK");
    throw e;
  } finally {
    client.release();
  }
}

function rowsForTemplate(type) {
  const today = new Date().toISOString().slice(0, 10);
  if (type === "sales") return [{ "日期": today, "业务员": "张三", "客户名称": "示例客户", "商品编码": "SP001", "商品名称": "示例商品", "销售数量": 10, "单价": 25.5, "销售金额": "" }];
  if (type === "salespeople") return [{ "业务员编码": "YW001", "业务员姓名": "张三", "所属区域": "华东区", "所属部门": "销售部", "入职日期": today, "状态": "启用" }];
  if (type === "products") return [{ "商品编码": "SP001", "商品名称": "牛肉卷", "商品规格": "500g/袋", "商品分类": "肉类", "是否重点商品": "是", "状态": "启用" }];
  if (type === "goldProducts") return [{ "商品编码": "SP001", "商品名称": "牛肉卷", "金砖商品目标": 300, "每日目标": 10 }];
  if (type === "salesTargets") return [{ "业务员": "张三", "巅峰目标": 200000, "每日目标": 8000 }];
  return [];
}

async function importMaintenance(req, res, type) {
  const config = BASE_CONFIG[type];
  if (!config) return json(res, 404, { error: "未知导入类型" });
  const { filename, buffer } = await readUploadFile(req);
  const errors = [];
  let success = 0;
  try {
    const rows = workbookRows(buffer, config.headers);
    const client = await getPool().connect();
    try {
      for (let i = 0; i < rows.length; i++) {
        const row = rows[i];
        try {
          if (type === "salespeople") {
            const name = String(row["业务员姓名"] || "").trim();
            if (!name) throw new Error("业务员姓名不能为空");
            await client.query(`
              INSERT INTO salespeople(code,name,region,department,entry_date,active,pending_confirm) VALUES ($1,$2,$3,$4,$5,$6,0)
              ON CONFLICT(name) DO UPDATE SET code=EXCLUDED.code,region=EXCLUDED.region,department=EXCLUDED.department,entry_date=EXCLUDED.entry_date,active=EXCLUDED.active
            `, [row["业务员编码"] || "", name, row["所属区域"] || "", row["所属部门"] || "", parseDateCell(row["入职日期"]) || null, activeStatus(row["状态"])]);
          } else if (type === "products") {
            const code = String(row["商品编码"] || "").trim();
            const name = String(row["商品名称"] || "").trim();
            if (!code || !name) throw new Error("商品编码、商品名称不能为空");
            await client.query(`
              INSERT INTO products(code,name,spec,category,is_key,status,pending_confirm) VALUES ($1,$2,$3,$4,$5,$6,0)
              ON CONFLICT(code) DO UPDATE SET name=EXCLUDED.name,spec=EXCLUDED.spec,category=EXCLUDED.category,is_key=EXCLUDED.is_key,status=EXCLUDED.status
            `, [code, name, row["商品规格"] || "", row["商品分类"] || "", truthy(row["是否重点商品"]) ? 1 : 0, row["状态"] || "启用"]);
          } else if (type === "goldProducts") {
            const code = String(row["商品编码"] || "").trim();
            const name = String(row["商品名称"] || "").trim();
            if (!code || !name) throw new Error("商品编码、商品名称不能为空");
            await client.query(`
              INSERT INTO products(code,name,is_key,gold_target,gold_daily_target,status,pending_confirm) VALUES ($1,$2,1,$3,$4,'启用',0)
              ON CONFLICT(code) DO UPDATE SET name=EXCLUDED.name,is_key=1,gold_target=EXCLUDED.gold_target,gold_daily_target=EXCLUDED.gold_daily_target
            `, [code, name, toNum(row["金砖商品目标"]), toNum(row["每日目标"])]);
          } else if (type === "salesTargets") {
            const salesperson = String(row["业务员"] || "").trim();
            if (!salesperson) throw new Error("业务员不能为空");
            const month = new Date().toISOString().slice(0, 7);
            await client.query(`
              INSERT INTO salesperson_targets(month,salesperson_name,peak_target,daily_target) VALUES ($1,$2,$3,$4)
              ON CONFLICT(month,salesperson_name) DO UPDATE SET peak_target=EXCLUDED.peak_target,daily_target=EXCLUDED.daily_target
            `, [month, salesperson, toNum(row["巅峰目标"]), toNum(row["每日目标"])]);
          }
          success++;
        } catch (e) {
          errors.push({ "行号": i + 2, "错误原因": e.message });
        }
      }
    } finally {
      client.release();
    }
  } catch (e) {
    errors.push({ "行号": "", "错误原因": filename.endsWith(".xls") ? "老式 .xls 文件请另存为 .xlsx 后导入" : e.message });
  }
  if (errors.length) {
    const errRows = errors;
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(errRows), "错误明细");
    const buffer = XLSX.write(wb, { bookType: "xlsx", type: "buffer" });
    const key = `error_${Date.now()}`;
    global.__errorFiles = global.__errorFiles || {};
    global.__errorFiles[key] = buffer;
    return json(res, 200, { ok: true, success, failed: errors.length, errorUrl: `/api/error-file/${key}` });
  }
  json(res, 200, { ok: true, success, failed: 0, errorUrl: "" });
}

async function uploadSales(req, res, url) {
  const { filename, buffer } = await readUploadFile(req);
  const rows = workbookRows(buffer, SALES_HEADERS).map((row) => {
    const qty = toNum(row["销售数量"]);
    const price = toNum(row["单价"]);
    return {
      sale_date: parseDateCell(row["日期"]),
      salesperson_name: String(row["业务员"] || "").trim(),
      customer_name: String(row["客户名称"] || "").trim(),
      product_code: String(row["商品编码"] || "").trim(),
      product_name: String(row["商品名称"] || "").trim(),
      quantity: qty,
      unit_price: price,
      amount: row["销售金额"] === "" ? qty * price : toNum(row["销售金额"])
    };
  });
  const dates = [...new Set(rows.map((r) => r.sale_date))].sort();
  const mode = url.searchParams.get("mode") || "";
  const existing = await query("SELECT sale_date, COUNT(*)::int c FROM sales WHERE sale_date = ANY($1) GROUP BY sale_date", [dates]);
  if (existing.length && !["merge", "replace"].includes(mode)) {
    return json(res, 409, { duplicate: true, message: "检测到相同日期已有销售数据，请选择覆盖或合并。", dates: existing });
  }
  const client = await getPool().connect();
  const newSalespeople = [];
  const newProducts = [];
  try {
    await client.query("BEGIN");
    if (mode === "replace" && dates.length) await client.query("DELETE FROM sales WHERE sale_date = ANY($1)", [dates]);
    for (const row of rows) {
      if (!row.sale_date || !row.salesperson_name || !row.product_code || !row.product_name) continue;
      const sp = await client.query("SELECT id FROM salespeople WHERE name=$1", [row.salesperson_name]);
      if (!sp.rowCount) {
        await client.query("INSERT INTO salespeople(name,active,pending_confirm) VALUES ($1,1,1)", [row.salesperson_name]);
        await addPending(client, "业务员", row.salesperson_name, row.salesperson_name, { "业务员姓名": row.salesperson_name });
        newSalespeople.push(row.salesperson_name);
      }
      const pr = await client.query("SELECT id FROM products WHERE code=$1", [row.product_code]);
      if (!pr.rowCount) {
        await client.query("INSERT INTO products(code,name,status,pending_confirm) VALUES ($1,$2,'启用',1)", [row.product_code, row.product_name]);
        await addPending(client, "商品", row.product_code, row.product_name, { "商品编码": row.product_code, "商品名称": row.product_name });
        newProducts.push(row.product_code);
      }
      await client.query(`
        INSERT INTO sales(sale_date,salesperson_name,customer_name,product_code,product_name,quantity,unit_price,amount,uploaded_at,source_file)
        VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10)
      `, [row.sale_date, row.salesperson_name, row.customer_name, row.product_code, row.product_name, row.quantity, row.unit_price, row.amount, new Date().toISOString(), filename]);
    }
    await client.query("COMMIT");
  } catch (e) {
    await client.query("ROLLBACK");
    throw e;
  } finally {
    client.release();
  }
  let message = `成功导入 ${rows.length} 条销售明细`;
  if (newSalespeople.length || newProducts.length) message += `，自动识别新业务员 ${newSalespeople.length} 个、新商品 ${newProducts.length} 个，已加入待确认数据池`;
  json(res, 200, { ok: true, rows: rows.length, dates, newSalespeople, newProducts, message });
}

async function main(req, res) {
  try {
    await ensureSchema();
    const url = parseUrl(req);
    const path = url.pathname.replace(/^\/api/, "") || "/";
    if (req.method === "GET" && path === "/db-check") return json(res, 200, await dbCheck());
    if (req.method === "GET" && path === "/dashboard") return apiDashboard(req, res, url);
    if (req.method === "GET" && path === "/sales-ranking") return json(res, 200, await salesRank(url.searchParams.get("date") || new Date().toISOString().slice(0, 10), url.searchParams.get("salesperson") || ""));
    if (req.method === "GET" && path === "/gold-rank") return json(res, 200, await goldRank(url.searchParams.get("month") || new Date().toISOString().slice(0, 7), url.searchParams.get("date") || new Date().toISOString().slice(0, 10), url.searchParams.get("salesperson") || ""));
    if (req.method === "GET" && path === "/product-stats") return json(res, 200, await productStats(url.searchParams.get("period") || "day", url.searchParams.get("date") || new Date().toISOString().slice(0, 10), url.searchParams.get("product") || "", url.searchParams.get("keyOnly") || ""));
    if (req.method === "GET" && path === "/maintenance") return json(res, 200, await maintenanceData());
    if (req.method === "POST" && path === "/maintenance") return saveMaintenance(req, res);
    if (req.method === "POST" && path === "/upload") return uploadSales(req, res, url);
    if (req.method === "GET" && path === "/sales-template") return xlsxResponse(res, "销售明细导入模板.xlsx", SALES_HEADERS, rowsForTemplate("sales"));
    if (req.method === "GET" && path === "/demo-sales") return xlsxResponse(res, "模拟销售明细.xlsx", SALES_HEADERS, rowsForTemplate("sales"));
    if (req.method === "GET" && path.startsWith("/base-template/")) {
      const type = path.split("/").pop();
      const config = BASE_CONFIG[type];
      if (!config) return json(res, 404, { error: "未知模板类型" });
      return xlsxResponse(res, `${config.title}导入模板.xlsx`, config.headers, rowsForTemplate(type));
    }
    if (req.method === "POST" && path.startsWith("/import-maintenance/")) return importMaintenance(req, res, path.split("/").pop());
    if (req.method === "GET" && path.startsWith("/export-maintenance/")) {
      const type = path.split("/").pop();
      const data = await maintenanceData();
      if (type === "salespeople") return xlsxResponse(res, "业务员基础资料_当前数据.xlsx", BASE_CONFIG.salespeople.headers, data.salespeople.map((r) => ({ "业务员编码": r.code, "业务员姓名": r.name, "所属区域": r.region, "所属部门": r.department, "入职日期": r.entry_date, "状态": r.active ? "启用" : "停用" })));
      if (type === "products") return xlsxResponse(res, "商品基础资料_当前数据.xlsx", BASE_CONFIG.products.headers, data.products.map((r) => ({ "商品编码": r.code, "商品名称": r.name, "商品规格": r.spec, "商品分类": r.category, "是否重点商品": r.is_key ? "是" : "否", "状态": r.status })));
      if (type === "goldProducts") return xlsxResponse(res, "重点商品维护_当前数据.xlsx", BASE_CONFIG.goldProducts.headers, data.goldTargets.map((r) => ({ "商品编码": r.code, "商品名称": r.name, "金砖商品目标": r.gold_target, "每日目标": r.gold_daily_target })));
      if (type === "salesTargets") return xlsxResponse(res, "销售目标维护_当前数据.xlsx", BASE_CONFIG.salesTargets.headers, data.salesTargets.map((r) => ({ "业务员": r.salesperson_name, "巅峰目标": r.peak_target, "每日目标": r.daily_target })));
    }
    if (req.method === "GET" && path.startsWith("/error-file/")) {
      const key = path.split("/").pop();
      const buffer = global.__errorFiles && global.__errorFiles[key];
      if (!buffer) return json(res, 404, { error: "错误明细已过期，请重新导入生成。" });
      res.statusCode = 200;
      res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
      res.setHeader("Content-Disposition", "attachment; filename=errors.xlsx");
      return res.end(buffer);
    }
    if (req.method === "GET" && path.startsWith("/export/")) {
      const name = path.split("/").pop();
      if (name === "sales-ranking") {
        const rows = await salesRank(url.searchParams.get("date") || new Date().toISOString().slice(0, 10), url.searchParams.get("salesperson") || "");
        return xlsxResponse(res, "业务销售排行.xlsx", ["业务员", "今日销售金额", "本月销售金额", "巅峰目标", "每日目标", "完成率"], rows);
      }
      if (name === "products") {
        const rows = await productStats(url.searchParams.get("period") || "day", url.searchParams.get("date") || new Date().toISOString().slice(0, 10), url.searchParams.get("product") || "", url.searchParams.get("keyOnly") || "");
        return xlsxResponse(res, "商品销售统计.xlsx", ["排名", "商品编码", "商品名称", "销售数量", "销售金额", "销售业务员数"], rows);
      }
      if (name === "gold-rank") {
        const rows = await goldRank(url.searchParams.get("month") || new Date().toISOString().slice(0, 7), url.searchParams.get("date") || new Date().toISOString().slice(0, 10), url.searchParams.get("salesperson") || "");
        return xlsxResponse(res, "金砖商品完成率排行.xlsx", ["业务员", "金砖商品目标", "金砖商品每日目标", "每日完成件数", "今日完成率", "累计完成", "累计完成率"], rows);
      }
    }
    return json(res, 404, { error: "接口不存在" });
  } catch (err) {
    console.error(err);
    return json(res, 500, { error: err.message || "服务器错误" });
  }
}

module.exports = main;
