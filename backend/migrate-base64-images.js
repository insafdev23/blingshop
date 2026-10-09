const path = require("path");
const fs = require("fs");

const backendEnv = path.join(__dirname, ".env");
if (fs.existsSync(backendEnv)) {
  require("dotenv").config({ path: backendEnv });
} else {
  require("dotenv").config();
}

const mysql = require("mysql2/promise");
const { uploadImage, DRIVER } = require("./storage");

const MIME_EXT = {
  "image/jpeg": ".jpg",
  "image/jpg": ".jpg",
  "image/png": ".png",
  "image/webp": ".webp",
  "image/gif": ".gif",
};

function parseDataUrl(dataUrl) {
  if (!dataUrl || typeof dataUrl !== "string") return null;
  const match = /^data:([^;]+);base64,([\s\S]+)$/.exec(dataUrl);
  if (!match) return null;
  const mime = match[1].toLowerCase();
  return {
    buffer: Buffer.from(match[2], "base64"),
    ext: MIME_EXT[mime] || ".jpg",
  };
}

async function runMigration() {
  console.log("──────────────────────────────────────────────────");
  console.log("  Base64 -> Storage URL Migration Script");
  console.log(`  Driver: ${DRIVER}`);
  console.log(`  Database: ${process.env.DB_HOST}:${process.env.DB_PORT}/${process.env.DB_NAME}`);
  console.log("──────────────────────────────────────────────────");

  const pool = mysql.createPool({
    host: process.env.DB_HOST || "localhost",
    port: Number(process.env.DB_PORT) || 28727,
    user: process.env.DB_USER || "avnadmin",
    password: process.env.DB_PASSWORD || "",
    database: process.env.DB_NAME || "defaultdb",
    ssl: { rejectUnauthorized: false },
    waitForConnections: true,
    connectionLimit: 3,
    connectTimeout: 30000,
  });

  const tables = ["products", "pending_products"];
  let totalMigrated = 0;
  let totalFailed = 0;

  for (const table of tables) {
    try {
      // 1. Fetch ONLY the IDs of rows that need migration (lightweight query to avoid ECONNRESET)
      const [idRows] = await pool.query(
        `SELECT id FROM ${table} WHERE image LIKE 'data:%'`
      );
      console.log(`\nScanning "${table}" table: found ${idRows.length} base64 images to convert.`);

      for (let i = 0; i < idRows.length; i++) {
        const id = idRows[i].id;
        try {
          // 2. Fetch one row at a time
          const [rows] = await pool.query(
            `SELECT image FROM ${table} WHERE id = ?`,
            [id]
          );
          if (!rows || rows.length === 0 || !rows[0].image) continue;

          const parsed = parseDataUrl(rows[0].image);
          if (!parsed) {
            console.warn(`  [${i + 1}/${idRows.length}] ID ${id}: Skipping (invalid data URL)`);
            totalFailed++;
            continue;
          }

          const filename = `${table}-${id}${parsed.ext}`;
          const url = await uploadImage(parsed.buffer, filename);

          // 3. Update database row with clean URL
          await pool.query(`UPDATE ${table} SET image = ? WHERE id = ?`, [url, id]);
          console.log(`  [ok] [${i + 1}/${idRows.length}] ${table} id=${id} -> ${url}`);
          totalMigrated++;
        } catch (itemErr) {
          console.error(`  [fail] [${i + 1}/${idRows.length}] ${table} id=${id}:`, itemErr.message);
          totalFailed++;
        }
      }
    } catch (err) {
      console.error(`Error scanning table "${table}":`, err.message);
    }
  }

  console.log("\n──────────────────────────────────────────────────");
  console.log(`  Migration Complete!`);
  console.log(`  Migrated: ${totalMigrated}, Failed: ${totalFailed}`);
  console.log("──────────────────────────────────────────────────");

  await pool.end();
}

runMigration().catch(err => {
  console.error("Migration failed with error:", err);
  process.exit(1);
});
