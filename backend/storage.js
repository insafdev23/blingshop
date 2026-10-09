// Image storage adapter.
//
// Two interchangeable drivers, picked by the STORAGE_DRIVER env var:
//   "local"       - writes files to backend/uploads/products and serves them
//                    via the /uploads static route in server.js. Works on any
//                    machine/server with a persistent filesystem (this dev
//                    machine, a VPS, Railway, etc). Default.
//   "vercel-blob"  - uploads to Vercel Blob storage over HTTP. Use this when
//                    the backend runs as a Vercel serverless function, which
//                    has no persistent local disk. Requires BLOB_READ_WRITE_TOKEN
//                    to be set (Vercel sets this automatically once a Blob
//                    store is attached to the project).
//
// Both drivers return a public URL string that can be stored directly in the
// products/pending_products "image" column in place of a base64 data URL.

const fs = require("fs");
const path = require("path");

const DRIVER = process.env.STORAGE_DRIVER || "local";
const LOCAL_UPLOAD_DIR = path.join(__dirname, "uploads", "products");
// Set PUBLIC_BASE_URL only if the backend serving /uploads is on a different
// origin than the one building the image URL would otherwise resolve to.
const PUBLIC_BASE_URL = process.env.PUBLIC_BASE_URL || "";

function randomFilename(originalName) {
  const ext = (path.extname(originalName || "") || ".jpg").toLowerCase();
  return `${Date.now()}-${Math.random().toString(36).slice(2, 10)}${ext}`;
}

function uploadLocal(buffer, originalName) {
  fs.mkdirSync(LOCAL_UPLOAD_DIR, { recursive: true });
  const filename = randomFilename(originalName);
  fs.writeFileSync(path.join(LOCAL_UPLOAD_DIR, filename), buffer);
  return `${PUBLIC_BASE_URL}/uploads/products/${filename}`;
}

async function uploadVercelBlob(buffer, originalName) {
  const { put } = require("@vercel/blob");
  const filename = `products/${randomFilename(originalName)}`;
  const blob = await put(filename, buffer, { access: "public" });
  return blob.url;
}

async function uploadImage(buffer, originalName) {
  if (DRIVER === "vercel-blob") return uploadVercelBlob(buffer, originalName);
  return uploadLocal(buffer, originalName);
}

module.exports = { uploadImage, DRIVER };
