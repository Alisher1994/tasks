/**
 * Последовательное применение SQL из server/migrations/*.sql
 */
import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

export async function runMigrations(pool) {
  await pool.query(`
    CREATE TABLE IF NOT EXISTS schema_migrations (
      id TEXT PRIMARY KEY,
      applied_at TIMESTAMPTZ DEFAULT NOW()
    )
  `);

  const migDir = path.join(__dirname, "migrations");
  if (!fs.existsSync(migDir)) {
    return;
  }
  const files = fs.readdirSync(migDir).filter((f) => f.endsWith(".sql")).sort();

  for (const file of files) {
    const id = file.replace(/\.sql$/i, "");
    const { rows } = await pool.query("SELECT 1 FROM schema_migrations WHERE id = $1", [id]);
    if (rows.length > 0) {
      continue;
    }

    const sql = fs.readFileSync(path.join(migDir, file), "utf8");
    const client = await pool.connect();
    try {
      await client.query("BEGIN");
      await client.query(sql);
      await client.query("INSERT INTO schema_migrations (id) VALUES ($1)", [id]);
      await client.query("COMMIT");
      console.log(`Миграция применена: ${file}`);
    } catch (e) {
      await client.query("ROLLBACK");
      throw e;
    } finally {
      client.release();
    }
  }
}
