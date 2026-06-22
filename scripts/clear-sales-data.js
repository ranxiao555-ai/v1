const { Pool } = require("pg");

function getDatabaseUrl() {
  return process.env.DATABASE_URL ||
    process.env.POSTGRES_URL ||
    process.env.POSTGRES_URL_NON_POOLING ||
    process.env.POSTGRES_PRISMA_URL ||
    "";
}

async function main() {
  const connectionString = getDatabaseUrl();
  if (!connectionString) {
    console.error("缺少 DATABASE_URL。请先设置 Neon PostgreSQL 连接字符串。");
    process.exitCode = 1;
    return;
  }

  const pool = new Pool({
    connectionString,
    ssl: process.env.PGSSLMODE === "disable" ? false : { rejectUnauthorized: false }
  });

  try {
    const result = await pool.query("DELETE FROM sales");
    const deleted = result.rowCount || 0;
    console.log(`数据库已清空，共删除${deleted}条记录`);
    console.log("已保留数据库结构、系统功能、管理员账号、业务员目标表结构和商品表结构。");
  } finally {
    await pool.end();
  }
}

main().catch((error) => {
  console.error(error.message || error);
  process.exitCode = 1;
});
