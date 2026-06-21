# 销售管理系统 V1 云端部署说明

本目录已经新增 Vercel 部署版本：

- `public/`：Vercel 前端页面
- `api/index.js`：Vercel Serverless API
- `package.json`：Node 依赖
- `vercel.json`：Vercel 路由配置

## 数据存储方式

云端版本不使用本地 SQLite 保存共享数据。

部署后所有业务数据写入 Neon PostgreSQL 云数据库，由环境变量 `DATABASE_URL` 指定。

推荐使用 Neon 的 `Pooled connection` 连接字符串。

## 必填环境变量

在 Vercel 项目设置中添加：

```text
DATABASE_URL=你的 PostgreSQL 连接字符串
```

示例：

```text
postgresql://user:password@host:5432/dbname?sslmode=require
```

## 部署步骤

```bash
cd /Users/mac/Documents/Codex/2026-06-20/zai/outputs/sales-report-v1
pnpm install
vercel
```

首次访问任意 `/api/...` 接口时，系统会自动创建这些 PostgreSQL 表：

- `salespeople`
- `products`
- `salesperson_targets`
- `gold_targets`
- `targets`
- `sales`
- `pending_items`

可以访问 `/api/db-check` 验证数据库连接和建表状态。

## 多人共享结论

只要所有人访问同一个 Vercel 域名，并且 Vercel 项目配置的是同一个 `DATABASE_URL`：

- 手机端上传的数据，电脑端能看到
- A 用户修改业务员资料，B 用户刷新后能看到
- 所有人共享同一套 PostgreSQL 数据

## 注意事项

- Vercel 版使用 `public/index.html`，不是 `static/index.html`
- Vercel 版接口在 `/api/...`
- Excel 导入模板由接口即时生成，不依赖本地 `exports/`
- `.xlsx` 完整支持
- 老式二进制 `.xls` 建议先另存为 `.xlsx`
