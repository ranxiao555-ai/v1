# 销售管理系统 V1.1

这是一个可部署到 Vercel 的销售管理系统。前端使用静态页面，后端使用 Vercel Serverless Function，数据永久保存到 Neon PostgreSQL。

## 当前功能

- 上传销售明细 Excel
- 自动识别新业务员、新商品，并加入待确认数据池
- 老板驾驶舱：首页显示今日销售额、本月销售额、本月完成率、业务员 TOP10、金砖商品完成率 TOP10
- 业务销售排行
- 商品销售排行
- 金砖商品完成率排行
- 业务员目标维护：业务员、巅峰目标、每日目标，支持 Excel 导入导出
- 金砖商品维护：商品编码、商品名称、是否金砖商品，支持 Excel 导入导出
- 基础数据维护：业务员、商品、金砖商品、销售目标
- 基础数据 Excel 模板下载、批量导入、导出当前数据
- 各统计表 Excel 导出
- 统计表和维护表每页最多显示 50 条，超过自动分页
- 表格支持每个字段单独搜索
- 业务销售、商品销售、金砖排行支持开始日期/结束日期筛选
- 自动创建 PostgreSQL 数据表
- 手机端访问适配
- 系统维护页：显示连接状态、销售记录数、业务员数量、商品数量
- 测试阶段数据库工具：清空销售数据、重新初始化数据库

## 技术架构

- 前端：`public/index.html` + `public/static/app.js` + `public/static/styles.css`
- 后端：`api/index.js`
- 部署：Vercel
- 数据库：Neon PostgreSQL
- Excel：`xlsx`
- 数据库驱动：`pg`

系统启动任意 API 时会自动执行数据库初始化，自动创建以下数据表：

- `salespeople`：业务员资料
- `products`：商品资料、金砖商品标记、金砖目标
- `salesperson_targets`：业务销售目标
- `gold_targets`：预留的按月金砖目标表
- `targets`：预留的业务员 + 商品 + 月份任务表
- `sales`：销售明细上传数据
- `pending_items`：待确认数据池

## 本地运行

先安装依赖：

```bash
npm install
```

本地需要配置 Neon PostgreSQL 连接字符串：

```bash
export DATABASE_URL="postgresql://用户名:密码@主机/dbname?sslmode=require"
npm run dev
```

打开：

```text
http://localhost:3000
```

检查数据库连接：

```text
http://localhost:3000/api/db-check
```

如果返回 `ok: true`，说明数据库连接成功，并且数据表已经自动创建。

## 数据库初始化工具

页面入口：`系统维护`。

- 清空销售数据：调用当前 Vercel 环境中的 `DATABASE_URL`，仅删除 `sales` 销售明细数据，不删除数据库表结构、业务员资料、商品资料、目标任务和管理员账号。
- 重新初始化数据库：清空业务表数据，保留数据库表结构、系统配置和管理员账号。
- 两个操作都会二次确认。
- 执行成功后显示：`数据库已清空，共删除X条记录`。

接口：

```text
GET /api/db-check
GET /api/db-status
POST /api/clear-sales-data
POST /api/reinitialize-database
```

命令行清空销售测试数据：

```bash
DATABASE_URL="你的 Neon 连接字符串" npm run clear:sales-data
```

该命令只执行：

```sql
DELETE FROM sales;
```

不会删除业务员目标、商品资料、管理员账号、数据库表结构和系统功能。

## Neon 创建步骤

1. 打开 [Neon 控制台](https://console.neon.tech/)。
2. 点击 `New Project`。
3. 填写项目名称，例如 `sales-report-v1`。
4. 选择 PostgreSQL 版本，默认即可。
5. 选择数据库区域，建议选择离用户近的区域。
6. 创建完成后，进入项目的 `Dashboard`。
7. 找到 `Connection string`。
8. 选择 `Pooled connection`。
9. 复制形如下面的连接字符串：

```text
postgresql://user:password@ep-xxxx-pooler.region.aws.neon.tech/neondb?sslmode=require
```

Neon 只需要创建一个 Project 和一个 Database。表不需要手动创建，系统会在第一次访问 API 时自动创建。

## Vercel 环境变量

在 Vercel 项目中配置：

| 变量名 | 必填 | 说明 |
| --- | --- | --- |
| `DATABASE_URL` | 是 | Neon PostgreSQL 连接字符串，推荐使用 Pooled connection |
| `PGSSLMODE` | 否 | 默认不需要填。代码默认启用 SSL；只有本地特殊情况才设为 `disable` |

代码也兼容 Neon/Vercel 集成自动生成的这些变量：

- `POSTGRES_URL`
- `POSTGRES_URL_NON_POOLING`
- `POSTGRES_PRISMA_URL`

但推荐统一手动配置 `DATABASE_URL`，最清楚。

## Vercel 配置步骤

1. 打开 [Vercel Dashboard](https://vercel.com/dashboard)。
2. 点击你的销售系统项目。
3. 点击顶部或左侧的 `Settings`。
4. 点击 `Environment Variables`。
5. 在 `Name` 输入 `DATABASE_URL`。
6. 在 `Value` 粘贴 Neon 的 Pooled connection string。
7. 环境选择 `Production`、`Preview`、`Development`，建议三个都勾选。
8. 点击 `Save`。
9. 点击 `Deployments`。
10. 找到最新一次部署，点击右侧三个点。
11. 点击 `Redeploy`。
12. 确认重新部署。

如果你的 Vercel 项目已经绑定 GitHub，本仓库推送后也会自动触发重新部署。

## 验证数据库连接成功

部署完成后，访问：

```text
https://你的域名/api/db-check
```

成功时会看到类似：

```json
{
  "ok": true,
  "database": "connected",
  "provider": "Neon PostgreSQL",
  "env": "DATABASE_URL",
  "schemaReady": true,
  "tables": ["salespeople", "products", "salesperson_targets", "gold_targets", "targets", "sales", "pending_items"],
  "counts": {
    "salespeople": 0,
    "products": 0,
    "sales": 0
  }
}
```

如果仍然提示缺少连接字符串，请检查：

- Vercel 环境变量名称必须是 `DATABASE_URL`
- 是否粘贴到了当前项目，而不是其他项目
- 是否勾选了 `Production`
- 修改环境变量后是否重新部署
- Neon 连接字符串是否包含 `sslmode=require`

## 上传销售 Excel

销售明细上传模板地址：

```text
/api/sales-template
```

上传文件必须包含以下列：

- 日期
- 业务员
- 客户名称
- 商品编码
- 商品名称
- 销售数量
- 单价
- 销售金额

如果销售金额为空，系统会自动计算：

```text
销售金额 = 销售数量 × 单价
```

导入成功后，数据会写入 Neon PostgreSQL 的 `sales` 表，刷新页面、换电脑、换手机访问都会看到同一套数据。

## 数据读取说明

- 业务销售排行读取 `sales`、`salespeople`、`salesperson_targets`
- 商品销售排行读取 `sales`、`products`
- 金砖商品排行读取 `sales`、`salespeople`、`products`
- 基础数据维护读取 `salespeople`、`products`、`salesperson_targets`、`pending_items`
- 业务员目标维护读取 `salesperson_targets`
- 金砖商品维护读取 `products`

金砖商品排行只统计 `products.is_key = 1` 且状态为 `启用` 的商品。

## 常用命令

语法检查：

```bash
npm run check
```

本地开发：

```bash
npm run dev
```

提交代码：

```bash
git add .
git commit -m "Use Neon PostgreSQL for Vercel deployment"
git push origin main
```
