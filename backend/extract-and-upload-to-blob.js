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

async function extractAndMigrate() {
  console.log("──────────────────────────────────────────────────");
  console.log("  Step 1: Extract Base64 & Upload to Vercel Blob");
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
  const cleanData = {};

  for (const table of tables) {
    cleanData[table] = [];
    try {
      // Read rows (SELECT works even if DB is in read-only mode!)
      const [idRows] = await pool.query(`SELECT id FROM ${table}`);
      console.log(`\nTable "${table}": found ${idRows.length} rows to process.`);

      for (let i = 0; i < idRows.length; i++) {
        const id = idRows[i].id;
        const [rows] = await pool.query(`SELECT * FROM ${table} WHERE id = ?`, [id]);
        if (!rows || rows.length === 0) continue;

        const row = { ...rows[0] };
        if (row.image && row.image.startsWith("data:")) {
          const parsed = parseDataUrl(row.image);
          if (parsed) {
            const filename = `${table}-${id}${parsed.ext}`;
            try {
              const url = await uploadImage(parsed.buffer, filename);
              row.image = url;
              console.log(`  [ok] [${i + 1}/${idRows.length}] ${table} id=${id} -> ${url}`);
            } catch (err) {
              console.error(`  [fail] [${i + 1}/${idRows.length}] ${table} id=${id} upload failed:`, err.message);
            }
          }
        } else {
          console.log(`  [ok] [${i + 1}/${idRows.length}] ${table} id=${id} (already clean URL or empty)`);
        }
        cleanData[table].push(row);
      }
    } catch (err) {
      console.error(`Error reading table "${table}":`, err.message);
    }
  }

  // Save clean data backup locally to disk
  const backupPath = path.join(__dirname, "clean-data-backup.json");
  fs.writeFileSync(backupPath, JSON.stringify(cleanData, null, 2));
  console.log(`\n✅ Saved complete clean backup to: ${backupPath}`);

  // Generate SQL restore script
  const sqlPath = path.join(__dirname, "clean-restore.sql");
  let sqlContent = `-- Clean POS Data Restore (Images on Vercel Blob)\n\n`;

  for (const table of tables) {
    const rows = cleanData[table];
    if (!rows || rows.length === 0) continue;
    sqlContent += `-- ${table}\n`;
    sqlContent += `TRUNCATE TABLE \`${table}\`;\n`;
    for (const row of rows) {
      const keys = Object.keys(row).map(k => `\`${k}\``).join(", ");
      const values = Object.values(row).map(v => {
        if (v === null || v === undefined) return "NULL";
        if (typeof v === "number") return v;
        return `'${String(v).replace(/'/g, "''").replace(/\\/g, "\\\\")}'`;
      }).join(", ");
      sqlContent += `INSERT INTO \`${table}\` (${keys}) VALUES (${values});\n`;
    }
    sqlContent += `\n`;
  }

  fs.writeFileSync(sqlPath, sqlContent);
  console.log(`✅ Generated clean SQL restore file: ${sqlPath}`);
  console.log(`   (File size: ${(fs.statSync(sqlPath).size / 1024).toFixed(1)} KB)`);

  await pool.end();
  console.log("\n──────────────────────────────────────────────────");
  console.log("  Extraction & Blob Upload Complete!");
  console.log("──────────────────────────────────────────────────");
}

extractAndMigrate().catch(err => {
  console.error("Extraction failed:", err);
  process.exit(1);
});
