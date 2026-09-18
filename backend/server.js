require("dotenv").config();

const express  = require("express");
const cors     = require("cors");
const mysql    = require("mysql2/promise");
const https    = require("https");
const http     = require("http");
const fs       = require("fs");
const path     = require("path");
const net      = require("net");
const bcrypt   = require("bcryptjs");

const app  = express();
const PORT = process.env.PORT || 3001;

// Strip leading /api prefix so routes match seamlessly on Vercel
app.use((req, res, next) => {
  if (req.url.startsWith('/api')) {
    req.url = req.url.replace(/^\/api/, '') || '/';
  }
  next();
});

// ... your route definitions come below here (e.g. app.post('/auth/login', ...))

// ── SSL certs ────────────────────────────────────────────────────
let sslOptions = null;
const certFile = path.join(__dirname, "localhost+1.pem");
const keyFile  = path.join(__dirname, "localhost+1-key.pem");
if (fs.existsSync(certFile) && fs.existsSync(keyFile)) {
  sslOptions = { cert: fs.readFileSync(certFile), key: fs.readFileSync(keyFile) };
  console.log("SSL certificates found — running HTTPS");
} else {
  console.warn("SSL certs not found — running HTTP only");
}

// ── MySQL connection pool ────────────────────────────────────────
const pool = mysql.createPool({
  host:               process.env.DB_HOST     || "localhost",
  port:               Number(process.env.DB_PORT) || 28727,
  user:               process.env.DB_USER     || "avnadmin",
  password:           process.env.DB_PASSWORD || "",
  database:           process.env.DB_NAME     || "defaultdb",
  ssl: {
    rejectUnauthorized: false
  },
  waitForConnections: true,
  connectionLimit:    5,
  queueLimit:         0,
  connectTimeout:     15000,
});

app.use(cors());
app.use(express.json({ limit: "10mb" }));

function generateId() { return "ST" + Date.now().toString().slice(-6) + Math.floor(Math.random()*90+10); }

// ── Auth ─────────────────────────────────────────────────────────
app.post("/auth/login", async (req, res) => {
  const { username, password } = req.body;
  if (!username || !password) return res.json({ success:false, error:"Username and password required" });
  try {
    const [rows] = await pool.query("SELECT * FROM staff WHERE username = ?", [username.trim()]);
    if (rows.length === 0) return res.json({ success:false, error:"Invalid username or password" });
    const staff = rows[0];
    if (!staff.active) return res.json({ success:false, error:"This account has been deactivated" });
    const match = bcrypt.compareSync(password, staff.password);
    if (!match) return res.json({ success:false, error:"Invalid username or password" });
    res.json({ success:true, staff: { id: staff.id, name: staff.name, username: staff.username, role: staff.role } });
  } catch (e) { res.status(500).json({ success:false, error: e.message }); }
});

// ── Staff Management ─────────────────────────────────────────────
app.get("/staff", async (req, res) => {
  try {
    const [rows] = await pool.query("SELECT id, name, username, role, active, createdAt FROM staff ORDER BY name");
    res.json(rows);
  } catch (e) { res.status(500).json({ error: e.message }); }
});

app.post("/staff", async (req, res) => {
  const { id, name, username, password, role } = req.body;
  try {
    const hash = bcrypt.hashSync(password, 10);
    await pool.query(
      "INSERT INTO staff (id, name, username, password, role, active) VALUES (?, ?, ?, ?, ?, 1)",
      [id, name, username.trim(), hash, role || "Salesperson"]
    );
    res.json({ success: true });
  } catch (e) { res.status(400).json({ error: e.code === "ER_DUP_ENTRY" ? "Username already exists" : e.message }); }
});

app.put("/staff/:id", async (req, res) => {
  const { name, username, role, active, password } = req.body;
  try {
    if (password) {
      const hash = bcrypt.hashSync(password, 10);
      await pool.query("UPDATE staff SET name=?, username=?, role=?, active=?, password=? WHERE id=?", [name, username.trim(), role, active ? 1 : 0, hash, req.params.id]);
    } else {
      await pool.query("UPDATE staff SET name=?, username=?, role=?, active=? WHERE id=?", [name, username.trim(), role, active ? 1 : 0, req.params.id]);
    }
    res.json({ success: true });
  } catch (e) { res.status(400).json({ error: e.code === "ER_DUP_ENTRY" ? "Username already exists" : e.message }); }
});

app.delete("/staff/:id", async (req, res) => {
  try {
    await pool.query("DELETE FROM staff WHERE id=?", [req.params.id]);
    res.json({ success: true });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

// ── Products ─────────────────────────────────────────────────────
app.get("/products", async (req, res) => {
  try {
    const [rows] = await pool.query("SELECT * FROM products ORDER BY name");
    res.json(rows);
  } catch (e) { res.status(500).json({ error: e.message }); }
});

app.post("/products", async (req, res) => {
  const { id, name, category, business, price, cost, stock, sku, image } = req.body;
  try {
    await pool.query(
      "INSERT INTO products (id, name, category, business, price, cost, stock, sku, image) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)",
      [id, name, category, business || "Blingshop", price, cost || 0, stock, sku, image || null]
    );
    res.json({ success: true });
  } catch (e) { res.status(400).json({ error: e.message }); }
});

app.put("/products/:id", async (req, res) => {
  const { name, category, business, price, cost, stock, sku, image } = req.body;
  try {
    await pool.query(
      "UPDATE products SET name=?, category=?, business=?, price=?, cost=?, stock=?, sku=?, image=? WHERE id=?",
      [name, category, business || "Blingshop", price, cost || 0, stock, sku, image || null, req.params.id]
    );
    res.json({ success: true });
  } catch (e) { res.status(400).json({ error: e.message }); }
});

app.delete("/products/:id", async (req, res) => {
  try {
    await pool.query("DELETE FROM products WHERE id=?", [req.params.id]);
    res.json({ success: true });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

// ── Pending Products (Bulk/Quick Add sourcing queue) ────────────────
app.get("/pending-products", async (req, res) => {
  try {
    const [rows] = await pool.query("SELECT * FROM pending_products ORDER BY createdAt DESC");
    res.json(rows);
  } catch (e) { res.status(500).json({ error: e.message }); }
});

app.post("/pending-products", async (req, res) => {
  const { id, business, category, sku, name, costThb, cost, price, stock, image, staffId, staffName } = req.body;
  try {
    await pool.query(
      "INSERT INTO pending_products (id, business, category, sku, name, costThb, cost, price, stock, image, createdAt, staffId, staffName) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, NOW(), ?, ?)",
      [id, business || "Blingshop", category || "Other", sku, name, costThb || 0, cost || 0, price || 0, stock || 1, image || null, staffId || null, staffName || null]
    );
    res.json({ success: true });
  } catch (e) { res.status(400).json({ error: e.message }); }
});

app.put("/pending-products/:id", async (req, res) => {
  const { business, category, sku, name, costThb, cost, price, stock, image } = req.body;
  try {
    await pool.query(
      "UPDATE pending_products SET business=?, category=?, sku=?, name=?, costThb=?, cost=?, price=?, stock=?, image=? WHERE id=?",
      [business || "Blingshop", category || "Other", sku, name, costThb || 0, cost || 0, price || 0, stock || 1, image || null, req.params.id]
    );
    res.json({ success: true });
  } catch (e) { res.status(400).json({ error: e.message }); }
});

app.delete("/pending-products/:id", async (req, res) => {
  try {
    await pool.query("DELETE FROM pending_products WHERE id=?", [req.params.id]);
    res.json({ success: true });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

// ── Sales ─────────────────────────────────────────────────────────
app.get("/sales", async (req, res) => {
  try {
    const [sales] = await pool.query("SELECT * FROM sales ORDER BY date DESC");
    const [items] = await pool.query("SELECT * FROM sale_items");
    res.json(sales.map(s => ({ ...s, items: items.filter(i => i.sale_id === s.id) })));
  } catch (e) { res.status(500).json({ error: e.message }); }
});

app.post("/sales", async (req, res) => {
  const { id, date, items, subtotal, discount, total, customerId, customerName, staffId, staffName, paymentMethod, cashAmount, cardAmount, bankAmount, deliveryMethod, deliveryFee, deliveryPaidTo } = req.body;
  const mysqlDate = new Date(date).toISOString().slice(0, 19).replace("T", " ");
  const conn = await pool.getConnection();
  try {
    await conn.beginTransaction();
    await conn.query(
      "INSERT INTO sales (id, date, subtotal, discount, total, customerId, customerName, staffId, staffName, paymentMethod, cashAmount, cardAmount, bankAmount, deliveryMethod, deliveryFee, deliveryPaidTo) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
      [id, mysqlDate, subtotal, discount, total, customerId || null, customerName || null, staffId || null, staffName || null, paymentMethod || "Cash", cashAmount || 0, cardAmount || 0, bankAmount || 0, deliveryMethod || "In Store", deliveryFee || 0, deliveryPaidTo || null]
    );
    for (const item of items) {
      await conn.query(
        "INSERT INTO sale_items (sale_id, product_id, name, price, qty, business) VALUES (?, ?, ?, ?, ?, ?)",
        [id, item.id, item.name, item.price, item.qty, item.business || "Blingshop"]
      );
      await conn.query(
        "UPDATE products SET stock = stock - ? WHERE id = ?",
        [item.qty, item.id]
      );
    }
    await conn.commit();
    res.json({ success: true });
  } catch (e) {
    await conn.rollback();
    res.status(400).json({ error: e.message });
  } finally {
    conn.release();
  }
});

// ── Refunds ───────────────────────────────────────────────────────
app.get("/refunds", async (req, res) => {
  try {
    const [rows] = await pool.query("SELECT * FROM refunds ORDER BY date DESC");
    res.json(rows.map(r => ({ ...r, items: typeof r.items === "string" ? JSON.parse(r.items) : r.items })));
  } catch (e) { res.status(500).json({ error: e.message }); }
});

app.post("/refunds", async (req, res) => {
  const { id, sale_id, date, items, refund_amount, type, notes, staffId, staffName } = req.body;
  const mysqlDate = new Date(date).toISOString().slice(0, 19).replace("T", " ");
  const conn = await pool.getConnection();
  try {
    await conn.beginTransaction();
    await conn.query(
      "INSERT INTO refunds (id, sale_id, date, items, refund_amount, type, notes, staffId, staffName) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)",
      [id, sale_id, mysqlDate, JSON.stringify(items), refund_amount, type || "refund", notes || null, staffId || null, staffName || null]
    );
    for (const item of items) {
      await conn.query("UPDATE products SET stock = stock + ? WHERE id = ?", [item.qty, item.product_id]);
    }
    await conn.commit();
    res.json({ success: true });
  } catch (e) {
    await conn.rollback();
    res.status(400).json({ error: e.message });
  } finally {
    conn.release();
  }
});

// ── Customers ─────────────────────────────────────────────────────
app.get("/customers", async (req, res) => {
  try {
    const [rows] = await pool.query("SELECT * FROM customers ORDER BY name");
    res.json(rows);
  } catch (e) { res.status(500).json({ error: e.message }); }
});

app.post("/customers", async (req, res) => {
  const { id, name, phone, email, birthday, preferredBiz, notes, createdAt } = req.body;
  const mysqlDate = new Date(createdAt || Date.now()).toISOString().slice(0, 19).replace("T", " ");
  try {
    await pool.query(
      "INSERT INTO customers (id, name, phone, email, birthday, preferredBiz, notes, createdAt) VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
      [id, name, phone || "", email || "", birthday || null, preferredBiz || "Both", notes || null, mysqlDate]
    );
    res.json({ success: true });
  } catch (e) { res.status(400).json({ error: e.message }); }
});

app.put("/customers/:id", async (req, res) => {
  const { name, phone, email, birthday, preferredBiz, notes } = req.body;
  try {
    await pool.query(
      "UPDATE customers SET name=?, phone=?, email=?, birthday=?, preferredBiz=?, notes=? WHERE id=?",
      [name, phone || "", email || "", birthday || null, preferredBiz || "Both", notes || null, req.params.id]
    );
    res.json({ success: true });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

app.delete("/customers/:id", async (req, res) => {
  try {
    await pool.query("DELETE FROM customers WHERE id=?", [req.params.id]);
    res.json({ success: true });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

// ── Settings ──────────────────────────────────────────────────────
app.get("/settings", async (req, res) => {
  try {
    const [rows] = await pool.query("SELECT * FROM settings");
    res.json(Object.fromEntries(rows.map(r => [r.key, r.value])));
  } catch (e) { res.status(500).json({ error: e.message }); }
});

app.put("/settings/:key", async (req, res) => {
  try {
    await pool.query(
      "INSERT INTO settings (`key`, value) VALUES (?, ?) ON DUPLICATE KEY UPDATE value=?",
      [req.params.key, req.body.value, req.body.value]
    );
    res.json({ success: true });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

// ── Zebra ZPL Proxy ───────────────────────────────────────────────
app.post("/print-zpl", express.text({ type: "*/*" }), (req, res) => {
  const { ip, port = 9100 } = req.query;
  const zpl = req.body;
  if (!ip) return res.status(400).json({ error: "Missing printer IP." });
  const client = new net.Socket();
  client.connect(Number(port), ip, () => { client.write(zpl); client.end(); });
  client.on("close", () => res.json({ success: true }));
  client.on("error", err => res.status(500).json({ error: err.message }));
});

// ── Health check ──────────────────────────────────────────────────
app.get("/health", async (req, res) => {
  try {
    await pool.query("SELECT 1");
    res.json({ status: "ok", database: "mysql", connected: true });
  } catch (e) {
    res.status(500).json({ status: "error", message: e.message });
  }
});

// ── Start ─────────────────────────────────────────────────────────
async function initDB() {
  const conn = await pool.getConnection();
  try {
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
      CREATE TABLE IF NOT EXISTS pending_products (
        id VARCHAR(32) PRIMARY KEY,
        business VARCHAR(20) NOT NULL DEFAULT 'Blingshop',
        category VARCHAR(30) NOT NULL DEFAULT 'Other',
        sku VARCHAR(20) NOT NULL,
        name VARCHAR(200) NOT NULL,
        costThb DECIMAL(10,2) NOT NULL DEFAULT 0,
        cost DECIMAL(10,2) NOT NULL DEFAULT 0,
        price DECIMAL(10,2) NOT NULL DEFAULT 0,
        stock INT NOT NULL DEFAULT 1,
        image LONGTEXT,
        createdAt DATETIME NOT NULL,
        staffId VARCHAR(32),
        staffName VARCHAR(100)
      )
    `);
    await conn.query(`INSERT IGNORE INTO settings (\`key\`, value) VALUES ('zebra_ip', '')`);

    // Add cost column to products if it doesn't exist yet (for stock-worth reporting)
    const [costCol] = await conn.query(
      "SELECT COUNT(*) as c FROM information_schema.columns WHERE table_schema = DATABASE() AND table_name = 'products' AND column_name = 'cost'"
    );
    if (costCol[0].c === 0) {
      await conn.query("ALTER TABLE products ADD COLUMN cost DECIMAL(10,2) NOT NULL DEFAULT 0 AFTER price");
      console.log("Added 'cost' column to products table.");
    }

    // Add payment-tracking columns to sales if they don't exist yet (Cash / Card / Bank Transfer / Split)
    const [payCol] = await conn.query(
      "SELECT COUNT(*) as c FROM information_schema.columns WHERE table_schema = DATABASE() AND table_name = 'sales' AND column_name = 'paymentMethod'"
    );
    if (payCol[0].c === 0) {
      await conn.query("ALTER TABLE sales ADD COLUMN paymentMethod VARCHAR(20) NOT NULL DEFAULT 'Cash' AFTER total");
      await conn.query("ALTER TABLE sales ADD COLUMN cashAmount DECIMAL(10,2) NOT NULL DEFAULT 0 AFTER paymentMethod");
      await conn.query("ALTER TABLE sales ADD COLUMN cardAmount DECIMAL(10,2) NOT NULL DEFAULT 0 AFTER cashAmount");
      await conn.query("ALTER TABLE sales ADD COLUMN bankAmount DECIMAL(10,2) NOT NULL DEFAULT 0 AFTER cardAmount");
      console.log("Added payment-tracking columns to sales table.");
    }

    // Add delivery-tracking columns to sales if they don't exist yet (In Store / Courier / Flash Delivery)
    const [delCol] = await conn.query(
      "SELECT COUNT(*) as c FROM information_schema.columns WHERE table_schema = DATABASE() AND table_name = 'sales' AND column_name = 'deliveryMethod'"
    );
    if (delCol[0].c === 0) {
      await conn.query("ALTER TABLE sales ADD COLUMN deliveryMethod VARCHAR(20) NOT NULL DEFAULT 'In Store' AFTER bankAmount");
      await conn.query("ALTER TABLE sales ADD COLUMN deliveryFee DECIMAL(10,2) NOT NULL DEFAULT 0 AFTER deliveryMethod");
      await conn.query("ALTER TABLE sales ADD COLUMN deliveryPaidTo VARCHAR(10) DEFAULT NULL AFTER deliveryFee");
      console.log("Added delivery-tracking columns to sales table.");
    }

    // Seed a default Owner account if no staff exist yet
    const [existing] = await conn.query("SELECT COUNT(*) as c FROM staff");
    if (existing[0].c === 0) {
      const defaultPassword = "owner123";
      const hash = bcrypt.hashSync(defaultPassword, 10);
      await conn.query(
        "INSERT INTO staff (id, name, username, password, role, active) VALUES (?, ?, ?, ?, ?, 1)",
        [generateId(), "Owner", "owner", hash, "Owner"]
      );
      console.log("─────────────────────────────────────────────");
      console.log("  No staff accounts found — created default login:");
      console.log("  Username: owner");
      console.log("  Password: owner123");
      console.log("  Please log in and change this immediately.");
      console.log("─────────────────────────────────────────────");
    }
    console.log("Database connection verified.");
  } finally {
    conn.release();
  }
}

// Initialize database connection
initDB().catch(err => {
  console.error("Failed to connect to database:", err.message);
});

// Only bind a network port when running locally (not inside Vercel serverless)
if (!process.env.VERCEL) {
  const server = sslOptions
    ? https.createServer(sslOptions, app)
    : http.createServer(app);

  server.listen(PORT, "0.0.0.0", () => {
    console.log("─────────────────────────────────────────────");
    console.log("  Blingshop + RC Boutique POS Backend");
    console.log(`  Protocol : ${sslOptions ? "HTTPS" : "HTTP"}`);
    console.log(`  Local    : ${sslOptions ? "https" : "http"}://localhost:${PORT}`);
    console.log(`  Database : MySQL (${process.env.DB_NAME || "blingshop"})`);
    console.log("─────────────────────────────────────────────");
  });
}

module.exports = app;