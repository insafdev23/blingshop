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
    connectionLimit: 5,
  });

  const tables = ["products", "pending_products"];
  let totalMigrated = 0;

  for (const table of tables) {
    try {
      const [rows] = await pool.query(
        `SELECT id, image FROM ${table} WHERE image LIKE 'data:image/%'`
      );
      console.log(`\nScanning "${table}" table: found ${rows.length} base64 images to convert.`);

      for (let i = 0; i < rows.length; i++) {
        const row = rows[i];
        const match = row.image.match(/^data:image\/([a-zA-Z0-9+]+);base64,(.+)$/s);
        if (!match) {
          console.warn(`  [${i + 1}/${rows.length}] ID ${row.id}: Skipping (unrecognized data URL format)`);
          continue;
        }

        const ext = match[1].toLowerCase() === "jpeg" ? "jpg" : match[1].toLowerCase();
        const buffer = Buffer.from(match[2], "base64");
        const filename = `${table}-${row.id}.${ext}`;

        try {
          const url = await uploadImage(buffer, filename);
          await pool.query(`UPDATE ${table} SET image = ? WHERE id = ?`, [url, row.id]);
          console.log(`  [${i + 1}/${rows.length}] ID ${row.id} -> ${url}`);
          totalMigrated++;
        } catch (err) {
          console.error(`  [${i + 1}/${rows.length}] ID ${row.id} failed:`, err.message);
        }
      }
    } catch (err) {
      console.error(`Error processing table "${table}":`, err.message);
    }
  }

  console.log("\n──────────────────────────────────────────────────");
  console.log(`  Migration Complete! Successfully migrated ${totalMigrated} image(s).`);
  console.log("──────────────────────────────────────────────────");

  await pool.end();
}

runMigration().catch(err => {
  console.error("Migration failed with error:", err);
  process.exit(1);
});
