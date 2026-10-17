const path = require("path");
const fs = require("fs");

const backendEnv = path.join(__dirname, ".env");
if (fs.existsSync(backendEnv)) {
  require("dotenv").config({ path: backendEnv });
} else {
  require("dotenv").config();
}

const mysql = require("mysql2/promise");

async function fixColumns() {
  console.log("──────────────────────────────────────────────────");
  console.log("  Database Schema & Missing Columns Fix");
  console.log(`  Database: ${process.env.DB_HOST}:${process.env.DB_PORT}/${process.env.DB_NAME}`);
  console.log("──────────────────────────────────────────────────");

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
    const checks = [
      {
        table: "sales",
        column: "customerId",
        sql: "ALTER TABLE sales ADD COLUMN customerId VARCHAR(32) DEFAULT NULL AFTER total",
      },
      {
        table: "sales",
        column: "customerName",
        sql: "ALTER TABLE sales ADD COLUMN customerName VARCHAR(100) DEFAULT NULL AFTER customerId",
      },
      {
        table: "sales",
        column: "staffId",
        sql: "ALTER TABLE sales ADD COLUMN staffId VARCHAR(32) DEFAULT NULL AFTER customerName",
      },
      {
        table: "sales",
        column: "staffName",
        sql: "ALTER TABLE sales ADD COLUMN staffName VARCHAR(100) DEFAULT NULL AFTER staffId",
      },
      {
        table: "sales",
        column: "paymentMethod",
        sql: "ALTER TABLE sales ADD COLUMN paymentMethod VARCHAR(20) NOT NULL DEFAULT 'Cash' AFTER total",
      },
      {
        table: "sales",
        column: "cashAmount",
        sql: "ALTER TABLE sales ADD COLUMN cashAmount DECIMAL(10,2) NOT NULL DEFAULT 0 AFTER paymentMethod",
      },
      {
        table: "sales",
        column: "cardAmount",
        sql: "ALTER TABLE sales ADD COLUMN cardAmount DECIMAL(10,2) NOT NULL DEFAULT 0 AFTER cashAmount",
      },
      {
        table: "sales",
        column: "bankAmount",
        sql: "ALTER TABLE sales ADD COLUMN bankAmount DECIMAL(10,2) NOT NULL DEFAULT 0 AFTER cardAmount",
      },
      {
        table: "sales",
        column: "deliveryMethod",
        sql: "ALTER TABLE sales ADD COLUMN deliveryMethod VARCHAR(20) NOT NULL DEFAULT 'In Store' AFTER bankAmount",
      },
      {
        table: "sales",
        column: "deliveryFee",
        sql: "ALTER TABLE sales ADD COLUMN deliveryFee DECIMAL(10,2) NOT NULL DEFAULT 0 AFTER deliveryMethod",
      },
      {
        table: "sales",
        column: "deliveryPaidTo",
        sql: "ALTER TABLE sales ADD COLUMN deliveryPaidTo VARCHAR(10) DEFAULT NULL AFTER deliveryFee",
      },
      {
        table: "customers",
        column: "birthday",
        sql: "ALTER TABLE customers ADD COLUMN birthday VARCHAR(20) DEFAULT NULL AFTER email",
      },
      {
        table: "customers",
        column: "preferredBiz",
        sql: "ALTER TABLE customers ADD COLUMN preferredBiz VARCHAR(30) NOT NULL DEFAULT 'Both' AFTER birthday",
      },
      {
        table: "customers",
        column: "createdAt",
        sql: "ALTER TABLE customers ADD COLUMN createdAt DATETIME DEFAULT CURRENT_TIMESTAMP AFTER notes",
      },
      {
        table: "products",
        column: "cost",
        sql: "ALTER TABLE products ADD COLUMN cost DECIMAL(10,2) NOT NULL DEFAULT 0 AFTER price",
      },
      {
        table: "sale_items",
        column: "paidPrice",
        sql: "ALTER TABLE sale_items ADD COLUMN paidPrice DECIMAL(10,2) DEFAULT NULL AFTER price",
      },
      {
        table: "pending_products",
        column: "costUnknown",
        sql: "ALTER TABLE pending_products ADD COLUMN costUnknown TINYINT(1) NOT NULL DEFAULT 0 AFTER image",
      },
    ];

    for (const item of checks) {
      const [rows] = await conn.query(
        "SELECT COUNT(*) as c FROM information_schema.columns WHERE table_schema = DATABASE() AND table_name = ? AND column_name = ?",
        [item.table, item.column]
      );
      if (rows[0].c === 0) {
        console.log(`Adding missing column \`${item.column}\` to table \`${item.table}\`...`);
        await conn.query(item.sql);
        console.log(`  ✅ Added \`${item.column}\` to \`${item.table}\`.`);
      } else {
        console.log(`  ✓ \`${item.table}.${item.column}\` already exists.`);
      }
    }

    try {
      console.log("Ensuring `sales.items` is nullable...");
      await conn.query("ALTER TABLE sales MODIFY COLUMN items JSON NULL");
      console.log("  ✅ Made `sales.items` nullable.");
    } catch (err) {
      console.warn("  Notice for `sales.items`:", err.message);
    }

    console.log("\n──────────────────────────────────────────────────");
    console.log("  ✅ All table columns verified and up to date!");
    console.log("──────────────────────────────────────────────────");
  } finally {
    conn.release();
    await pool.end();
  }
}

fixColumns().catch(err => {
  console.error("❌ Migration failed:", err.message);
  process.exit(1);
});
