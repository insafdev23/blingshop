const path = require("path");
const fs = require("fs");

const backendEnv = path.join(__dirname, ".env");
if (fs.existsSync(backendEnv)) {
  require("dotenv").config({ path: backendEnv });
} else {
  require("dotenv").config();
}

const mysql = require("mysql2/promise");
const bcrypt = require("bcryptjs");

function generateId() {
  return "ST" + Date.now().toString().slice(-6) + Math.floor(Math.random() * 90 + 10);
}

async function restoreClean() {
  console.log("──────────────────────────────────────────────────");
  console.log("  Step 2: Initialize Tables & Restore Clean Data");
  console.log(`  Database: ${process.env.DB_HOST}:${process.env.DB_PORT}/${process.env.DB_NAME}`);
  console.log("──────────────────────────────────────────────────");

  const backupPath = path.join(__dirname, "clean-data-backup.json");
  if (!fs.existsSync(backupPath)) {
    console.error("❌ Error: clean-data-backup.json not found.");
    process.exit(1);
  }

  const cleanData = JSON.parse(fs.readFileSync(backupPath, "utf8"));

  const pool = mysql.createPool({
    host: process.env.DB_HOST || "localhost",
    port: Number(process.env.DB_PORT) || 28727,
    user: process.env.DB_USER || "avnadmin",
    password: process.env.DB_PASSWORD || "",
    database: process.env.DB_NAME || "blingshop",
    ssl: { rejectUnauthorized: false },
    waitForConnections: true,
    connectionLimit: 3,
  });

  const conn = await pool.getConnection();

  try {
    console.log("\n1. Creating database tables if they do not exist...");

    await conn.query(`
      CREATE TABLE IF NOT EXISTS staff (
        id VARCHAR(32) PRIMARY KEY,
        name VARCHAR(100) NOT NULL,
        username VARCHAR(50) NOT NULL UNIQUE,
        password VARCHAR(255) NOT NULL,
        role VARCHAR(20) NOT NULL DEFAULT 'Cashier',
        active TINYINT(1) NOT NULL DEFAULT 1
      )
    `);

    await conn.query(`
      CREATE TABLE IF NOT EXISTS products (
        id VARCHAR(32) PRIMARY KEY,
        name VARCHAR(200) NOT NULL,
        category VARCHAR(50) NOT NULL DEFAULT 'Other',
        business VARCHAR(30) NOT NULL DEFAULT 'Blingshop',
        price DECIMAL(10,2) NOT NULL DEFAULT 0,
        cost DECIMAL(10,2) NOT NULL DEFAULT 0,
        stock INT NOT NULL DEFAULT 0,
        sku VARCHAR(50),
        image TEXT,
        created_at DATETIME
      )
    `);

    await conn.query(`
      CREATE TABLE IF NOT EXISTS pending_products (
        id VARCHAR(32) PRIMARY KEY,
        business VARCHAR(30) NOT NULL DEFAULT 'Blingshop',
        category VARCHAR(50) NOT NULL DEFAULT 'Other',
        sku VARCHAR(50) NOT NULL,
        name VARCHAR(200) NOT NULL,
        costThb DECIMAL(10,2) NOT NULL DEFAULT 0,
        cost DECIMAL(10,2) NOT NULL DEFAULT 0,
        price DECIMAL(10,2) NOT NULL DEFAULT 0,
        stock INT NOT NULL DEFAULT 1,
        image TEXT,
        costUnknown TINYINT(1) NOT NULL DEFAULT 0,
        createdAt DATETIME,
        staffId VARCHAR(32),
        staffName VARCHAR(100)
      )
    `);

    await conn.query(`
      CREATE TABLE IF NOT EXISTS sales (
        id VARCHAR(32) PRIMARY KEY,
        date DATETIME NOT NULL,
        items JSON NOT NULL,
        subtotal DECIMAL(10,2) NOT NULL,
        discount DECIMAL(10,2) NOT NULL DEFAULT 0,
        total DECIMAL(10,2) NOT NULL,
        paymentMethod VARCHAR(20) NOT NULL DEFAULT 'Cash',
        cashAmount DECIMAL(10,2) NOT NULL DEFAULT 0,
        cardAmount DECIMAL(10,2) NOT NULL DEFAULT 0,
        bankAmount DECIMAL(10,2) NOT NULL DEFAULT 0,
        deliveryMethod VARCHAR(20) NOT NULL DEFAULT 'In Store',
        deliveryFee DECIMAL(10,2) NOT NULL DEFAULT 0,
        deliveryPaidTo VARCHAR(10) DEFAULT NULL,
        business VARCHAR(30) NOT NULL DEFAULT 'Blingshop',
        staffId VARCHAR(32),
        staffName VARCHAR(100)
      )
    `);

    await conn.query(`
      CREATE TABLE IF NOT EXISTS refunds (
        id VARCHAR(32) PRIMARY KEY,
        sale_id VARCHAR(32) NOT NULL,
        date DATETIME NOT NULL,
        items JSON NOT NULL,
        refund_amount DECIMAL(10,2) NOT NULL,
        type VARCHAR(16) NOT NULL DEFAULT 'refund',
        notes TEXT,
        staffId VARCHAR(32),
        staffName VARCHAR(100)
      )
    `);

    await conn.query(`
      CREATE TABLE IF NOT EXISTS customers (
        id VARCHAR(32) PRIMARY KEY,
        name VARCHAR(100) NOT NULL,
        phone VARCHAR(50),
        email VARCHAR(100),
        points INT DEFAULT 0,
        notes TEXT,
        created_at DATETIME
      )
    `);

    await conn.query(`
      CREATE TABLE IF NOT EXISTS settings (
        \`key\` VARCHAR(50) PRIMARY KEY,
        value TEXT
      )
    `);

    await conn.query(`INSERT IGNORE INTO settings (\`key\`, value) VALUES ('zebra_ip', '')`);

    // Seed default owner account if no staff exist
    const [staffRows] = await conn.query("SELECT COUNT(*) as c FROM staff");
    if (staffRows[0].c === 0) {
      const defaultPassword = "owner123";
      const hash = bcrypt.hashSync(defaultPassword, 10);
      await conn.query(
        "INSERT INTO staff (id, name, username, password, role, active) VALUES (?, ?, ?, ?, ?, 1)",
        [generateId(), "Owner", "owner", hash, "Owner"]
      );
      console.log("  Created default login: username=owner, password=owner123");
    }

    console.log("✅ All tables created successfully.");

    // 2. Insert clean data
    for (const table of Object.keys(cleanData)) {
      const rows = cleanData[table];
      console.log(`\nRestoring table "${table}" (${rows.length} rows)...`);

      // Clear existing records in case of re-run
      await conn.query(`DELETE FROM \`${table}\``);

      let count = 0;
      for (const row of rows) {
        // Format dates properly for MySQL DATETIME
        const formattedRow = { ...row };
        if (formattedRow.created_at) {
          formattedRow.created_at = new Date(formattedRow.created_at).toISOString().slice(0, 19).replace('T', ' ');
        }
        if (formattedRow.createdAt) {
          formattedRow.createdAt = new Date(formattedRow.createdAt).toISOString().slice(0, 19).replace('T', ' ');
        }

        const keys = Object.keys(formattedRow);
        const cols = keys.map(k => `\`${k}\``).join(", ");
        const placeholders = keys.map(() => "?").join(", ");
        const values = keys.map(k => formattedRow[k]);

        await conn.query(
          `INSERT INTO \`${table}\` (${cols}) VALUES (${placeholders})`,
          values
        );
        count++;
      }
      console.log(`  Inserted ${count} clean records into "${table}".`);
    }

    console.log("\n──────────────────────────────────────────────────");
    console.log("  ✅ SUCCESS! Complete POS Database Restored Cleanly.");
    console.log("──────────────────────────────────────────────────");
  } finally {
    conn.release();
    await pool.end();
  }
}

restoreClean().catch(err => {
  console.error("Restore failed:", err);
  process.exit(1);
});
