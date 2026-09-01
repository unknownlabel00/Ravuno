import React, { useState, useEffect, useMemo, useRef } from "react";
import {
  Check, Copy, Ruler, Plus, Trash2, ChevronRight, ChevronLeft,
  Package, TrendingUp, Clock, Image as ImageIcon, X, Loader2,
  Search, Download, Star, Tag, Receipt, ShoppingBag, Percent, DollarSign
} from "lucide-react";
import * as XLSX from "xlsx";
import Papa from "papaparse";
import { LineChart, Line, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer } from "recharts";

// RAVUNO Vault Ledger tokens — bg: #0B0B0C  surface: #17171A  hairline: #2A2A2F  brass: #C9A567  ivory: #F2F0EA  muted: #8C8A85
// border: #2A2A2F  accent: #C9A567  destructive: #C9695E
const PLATFORMS = [
  { id: "carousell", label: "Carousell" },
  { id: "threads", label: "Threads" },
  { id: "grailed", label: "Grailed" },
  { id: "depop", label: "Depop" },
  { id: "shopee", label: "Shopee" },
];
const MAX_PHOTOS = 10;

// ---------- Supabase backend (plain REST — the supabase-js SDK isn't available in this sandbox) ----------
const SUPABASE_URL = "https://qixrcxwinvcwqddnvowq.supabase.co";
const SUPABASE_ANON_KEY = "sb_publishable_2LdyldgK6bdhssGp_Kz3JQ_7f4GcRAg";
const AUTH_URL = `${SUPABASE_URL}/auth/v1`;
const REST_URL = `${SUPABASE_URL}/rest/v1`;
const STORAGE_URL = `${SUPABASE_URL}/storage/v1`;
const SESSION_KEY = "ravuno_session";

function loadStoredSession() {
  try { return JSON.parse(localStorage.getItem(SESSION_KEY)); } catch (e) { return null; }
}
function saveStoredSession(session) {
  if (session) localStorage.setItem(SESSION_KEY, JSON.stringify(session));
  else localStorage.removeItem(SESSION_KEY);
}

async function authRequest(path, body) {
  const res = await fetch(`${AUTH_URL}${path}`, {
    method: "POST",
    headers: { "Content-Type": "application/json", apikey: SUPABASE_ANON_KEY },
    body: JSON.stringify(body),
  });
  const data = await res.json().catch(() => ({}));
  if (!res.ok) throw new Error(data.error_description || data.msg || data.error || "Something went wrong. Please try again.");
  return data;
}
const signUpRequest = (email, password) => authRequest("/signup", { email, password });
const signInRequest = (email, password) => authRequest("/token?grant_type=password", { email, password });
const refreshSessionRequest = (refresh_token) => authRequest("/token?grant_type=refresh_token", { refresh_token });
async function signOutRequest(session) {
  try {
    await fetch(`${AUTH_URL}/logout`, { method: "POST", headers: { apikey: SUPABASE_ANON_KEY, Authorization: `Bearer ${session.access_token}` } });
  } catch (e) { /* best effort — clearing local session is what actually matters */ }
}

// generic REST call to a Postgres table, with a single retry after refreshing an expired session
async function apiRequest(path, options, session, onSession) {
  const run = (s) => fetch(`${REST_URL}${path}`, {
    ...options,
    headers: {
      "Content-Type": "application/json",
      apikey: SUPABASE_ANON_KEY,
      Authorization: `Bearer ${s.access_token}`,
      Prefer: "return=representation",
      ...(options?.headers || {}),
    },
  });
  let res = await run(session);
  if (res.status === 401 && session?.refresh_token) {
    const refreshed = await refreshSessionRequest(session.refresh_token);
    onSession?.(refreshed);
    session = refreshed;
    res = await run(session);
  }
  if (!res.ok) {
    const err = await res.json().catch(() => ({}));
    throw new Error(err.message || "Request failed");
  }
  if (res.status === 204) return null;
  const text = await res.text();
  return text ? JSON.parse(text) : null;
}

async function uploadPhotoToStorage(dataUrl, path, session) {
  const blob = await (await fetch(dataUrl)).blob();
  const res = await fetch(`${STORAGE_URL}/object/product-photos/${path}`, {
    method: "POST",
    headers: { apikey: SUPABASE_ANON_KEY, Authorization: `Bearer ${session.access_token}`, "Content-Type": blob.type || "image/jpeg", "x-upsert": "true" },
    body: blob,
  });
  if (!res.ok) throw new Error("Photo upload failed");
  return `${STORAGE_URL}/object/public/product-photos/${path}`;
}

// row <-> app-model mapping (DB uses snake_case; the app's components use camelCase)
function rowToItem(row) {
  return {
    id: row.id, title: row.title || "", category: row.category || "", itemType: row.item_type || "",
    brand: row.brand || "", condition: row.condition || "", colour: row.colour || "", colourCondition: row.colour_condition || "", remark: row.remark || "",
    cost: row.cost ?? "", price: row.price ?? "", sku: row.sku || "", status: row.status || "listed",
    attributes: row.attributes || {}, customAttributes: row.custom_attributes || [], photos: row.photos || [],
    addedOn: row.added_on, listedOn: row.listed_on, soldOn: row.sold_on,
  };
}
function itemToRow(item, userId) {
  return {
    user_id: userId, title: item.title || null, category: item.category || null, item_type: item.itemType || null,
    brand: item.brand || null, condition: item.condition || null, colour: item.colour || null, colour_condition: item.colourCondition || null, remark: item.remark || null,
    cost: item.cost === "" ? null : item.cost, price: item.price === "" ? null : item.price, sku: item.sku || null,
    status: item.status || "listed", attributes: item.attributes || {}, custom_attributes: item.customAttributes || [],
    photos: item.photos || [], added_on: item.addedOn || null, listed_on: item.listedOn || null, sold_on: item.soldOn || null,
  };
}
function rowToTxn(row) {
  return { id: row.id, timestamp: row.occurred_at, method: row.method, subtotal: row.subtotal, discount: row.discount, total: row.total, lines: row.lines || [] };
}
function txnToRow(txn, userId) {
  return { user_id: userId, occurred_at: txn.timestamp, method: txn.method, subtotal: txn.subtotal, discount: txn.discount, total: txn.total, lines: txn.lines };
}


// ---------- category architecture ----------
// Product → Category → Item Type → Category-Specific Attributes → Common Selling Fields
const CATEGORY_LIST = [
  { id: "apparel", label: "Apparel" },
  { id: "footwear", label: "Footwear" },
  { id: "bags", label: "Bags" },
  { id: "watches", label: "Watches" },
  { id: "electronics", label: "Electronics" },
  { id: "home", label: "Home & Living" },
  { id: "books", label: "Books" },
  { id: "games", label: "Games & Collectibles" },
  { id: "other", label: "Other" },
];
const CATEGORY_LABELS = Object.fromEntries(CATEGORY_LIST.map((c) => [c.id, c.label]));

// matches a loosely-typed CSV category value (id or label, case-insensitive) to a real category id
function matchCategoryId(raw) {
  const v = (raw || "").trim().toLowerCase();
  const found = CATEGORY_LIST.find((c) => c.id.toLowerCase() === v || c.label.toLowerCase() === v);
  return found ? found.id : "other";
}

const CSV_IMPORT_COLUMNS = ["title", "category", "itemType", "brand", "condition", "colour", "colourCondition", "remark", "cost", "price", "sku"];

// attribute definitions per category (apparel and "other" are handled specially)
const CATEGORY_DEFS = {
  footwear: {
    attributes: [
      { id: "shoeType", label: "Shoe Type" },
      { id: "size", label: "Size" },
      { id: "usSize", label: "US Size" },
      { id: "ukSize", label: "UK Size" },
      { id: "euSize", label: "EU Size" },
      { id: "footLength", label: "Foot Length (cm)", type: "number" },
      { id: "material", label: "Material" },
    ],
  },
  bags: {
    attributes: [
      { id: "bagType", label: "Bag Type" },
      { id: "material", label: "Material" },
      { id: "width", label: "Width (cm)", type: "number" },
      { id: "height", label: "Height (cm)", type: "number" },
      { id: "depth", label: "Depth (cm)", type: "number" },
      { id: "strapLength", label: "Strap Length (cm)", type: "number" },
      { id: "closureType", label: "Closure Type" },
    ],
  },
  watches: {
    attributes: [
      { id: "watchType", label: "Watch Type" },
      { id: "model", label: "Model" },
      { id: "movement", label: "Movement" },
      { id: "caseSize", label: "Case Size (mm)", type: "number" },
      { id: "caseMaterial", label: "Case Material" },
      { id: "strapMaterial", label: "Strap Material" },
      { id: "dialColour", label: "Dial Colour" },
      { id: "waterResistance", label: "Water Resistance" },
      { id: "serialNumber", label: "Serial / Model Number" },
    ],
  },
  electronics: {
    attributes: [
      { id: "deviceType", label: "Device Type" },
      { id: "model", label: "Model" },
      { id: "serialNumber", label: "Serial Number" },
      { id: "workingCondition", label: "Function / Working Condition" },
      { id: "accessories", label: "Accessories Included" },
      { id: "storage", label: "Storage / Capacity" },
      { id: "batteryCondition", label: "Battery Condition" },
      { id: "connectivity", label: "Connectivity" },
    ],
  },
  home: {
    attributes: [
      { id: "material", label: "Material" },
      { id: "dimensions", label: "Dimensions" },
      { id: "weight", label: "Weight" },
      { id: "includedItems", label: "Included Items" },
    ],
  },
  books: {
    attributes: [
      { id: "author", label: "Author" },
      { id: "publisher", label: "Publisher" },
      { id: "isbn", label: "ISBN" },
      { id: "edition", label: "Edition" },
      { id: "language", label: "Language" },
      { id: "publicationYear", label: "Publication Year", type: "number" },
      { id: "format", label: "Format" },
    ],
  },
  games: {
    attributes: [
      { id: "platform", label: "Platform" },
      { id: "edition", label: "Edition" },
      { id: "manufacturer", label: "Manufacturer" },
      { id: "year", label: "Year", type: "number" },
      { id: "region", label: "Region" },
      { id: "includedItems", label: "Included Items" },
      { id: "authenticity", label: "Authenticity / Notes" },
    ],
  },
};

const APPAREL_TOP_TYPES = ["T-Shirt", "Shirt", "Polo", "Hoodie", "Jacket", "Sweater", "Other topwear"];
const APPAREL_BOTTOM_TYPES = ["Jeans", "Pants", "Shorts", "Skirt", "Dress", "Other bottomwear"];

// unclear/other apparel item types default to "top" unless explicitly a bottom type
function getApparelSubcategory(itemType) {
  return APPAREL_BOTTOM_TYPES.includes(itemType) ? "bottom" : "top";
}

// approximate vintage/second-hand sizing — a suggestion only, never overwrites a manually entered size
function computeSuggestedSize(a, sub) {
  if (sub === "bottom") {
    const w = parseFloat(a.waist);
    return w ? `W${w}` : "";
  }
  const c = parseFloat(a.chest);
  if (!c) return "";
  if (c <= 19) return "XS";
  if (c <= 20) return "S";
  if (c <= 22) return "M";
  if (c <= 24) return "L";
  if (c <= 26) return "XL";
  if (c <= 28) return "XXL";
  return "XXXL";
}

function estimateSizeFlag(chest, length) {
  const c = parseFloat(chest), l = parseFloat(length);
  if (!c || !l) return null;
  const bands = [
    { max: 20, expected: [26, 28] }, { max: 22, expected: [27, 29] },
    { max: 24, expected: [28, 30] }, { max: 26, expected: [29, 31] },
    { max: Infinity, expected: [30, 33] },
  ];
  const [lo] = bands.find((b) => c <= b.max).expected;
  if (l < lo - 1.5) return "Length runs short for this chest — possibly cropped or shortened. Flag in listing.";
  return null;
}

function productTitle(p) {
  if (p.title && p.title.trim()) return p.title.trim();
  const auto = `${p.brand || ""} ${p.itemType || ""}`.trim();
  return auto || "[Product]";
}

function getFilledAttrLines(p) {
  if (p.category === "other") {
    return (p.customAttributes || []).filter((c) => c.label && c.value).map((c) => `${c.label}: ${c.value}`);
  }
  const defs = CATEGORY_DEFS[p.category]?.attributes || [];
  return defs.filter((d) => p.attributes?.[d.id]).map((d) => `${d.label}: ${p.attributes[d.id]}`);
}

function apparelMeasureParts(a, sub) {
  const parts = sub === "bottom"
    ? [["Waist", a.waist], ["Length", a.length], ["Front Rise", a.frontRise], ["Thigh", a.thigh], ["Leg Opening", a.legOpening]]
    : [["Chest", a.chest], ["Length", a.length], ["Shoulder", a.shoulder], ["Sleeve", a.sleeve]];
  return parts.filter(([, v]) => v).map(([label, v]) => `${label}: ${v}"`);
}

// ---------- template engine ----------
// Category → Platform → Template. Apparel gets the advanced treatment; every other
// category shares one generic, category-aware template generator.
function buildApparelTemplates(p) {
  const a = p.attributes || {};
  const sub = getApparelSubcategory(p.itemType);
  const suggested = computeSuggestedSize(a, sub);
  const size = a.size || suggested || "";
  const sizeFlag = sub === "top" ? estimateSizeFlag(a.chest, a.length) : null;
  const title = productTitle(p);
  const measureParts = apparelMeasureParts(a, sub);
  const measureLine = measureParts.join(" | ") || "[measurements]";
  const measureBullets = measureParts.map((m) => `• ${m}`).join("\n") || "• [measurements]";
  const condLine = p.colourCondition ? `${p.condition || "?"}/10 (Colour ${p.colourCondition}/10)` : `${p.condition || "?"}/10`;

  const carousell = `${title} — Size ${size || "[Size]"}

Condition: ${condLine}
${measureLine}
Remark: ${p.remark || "[key detail]"}

Price: RM${p.price || "?"}
SKU: ${p.sku || "[XXX]"}

Thanks for checking out my shop!`;

  const threads = `${title}

Condition: ${condLine}
${measureLine}
Remark: ${p.remark || "[key detail]"}

Price: RM${p.price || "?"}
SKU: ${p.sku || "?"}

Free shipping 🚚`;

  const grailed = `PLEASE READ BEFORE PURCHASING

Thank you for your interest in my small shop!
Feel free to ask any questions — I reply fast 😊

━━━━━━━━━━━━━━━━━━━━

📌 Item Condition
• Condition: ${p.condition || "__"}/10
• Notes: Used & vintage item. Signs of wear may be present.
• Remark: ${p.remark || "[Insert condition details / flaws / stains / tears]"}

━━━━━━━━━━━━━━━━━━━━

📏 Size & Measurements

• Size: ${size || "[Size]"}
${measureBullets}

⚠️ Please compare the measurements with your own garments to ensure proper fit.

━━━━━━━━━━━━━━━━━━━━

🚚 Shipping Information
• Tracking number will be provided.

📞 Buyer must provide a phone number after purchase for courier delivery purposes.

━━━━━━━━━━━━━━━━━━━━

❗ Important Notes
• No returns / refunds / exchanges.
• Please double-check measurements and ask any questions before purchasing.
• Buyer is responsible for any customs fees or import taxes, if applicable.

━━━━━━━━━━━━━━━━━━━━

🏷 Additional Info
• SKU: ${p.sku || "[SKU]"}

━━━━━━━━━━━━━━━━━━━━

Thank you very much for viewing this item 🙏`;

  const depop = `${p.itemType || "item"} by ${p.brand || "brand"} 🧵
size ${size || "?"} · condition ${condLine}
${p.remark ? p.remark + " " : ""}💌 bundle for a deal, ships worldwide`;

  const shopee = `${title}
• Saiz: ${size || "-"}
${measureBullets}
• Kondisi: ${condLine}
${p.remark ? `• Catatan: ${p.remark}\n` : ""}SKU: ${p.sku || "-"}`;

  return { carousell, threads, grailed, depop, shopee, sizeFlag };
}

function buildGenericTemplates(p) {
  const title = productTitle(p);
  const attrLines = getFilledAttrLines(p);
  const condLine = p.colourCondition ? `${p.condition || "?"}/10 (Colour ${p.colourCondition}/10)` : `${p.condition || "?"}/10`;

  const carousell = `${title}

Condition: ${condLine}
${attrLines.length ? attrLines.join("\n") + "\n" : ""}${p.colour ? `Colour: ${p.colour}\n` : ""}Remark: ${p.remark || "[key detail]"}

Price: RM${p.price || "?"}
SKU: ${p.sku || "[XXX]"}

Thanks for checking out my shop!`;

  const threads = `${title} 🏷️ condition ${condLine}${p.remark ? ` — ${p.remark}` : ""}
${attrLines.map((l) => `• ${l}`).join("\n")}

RM${p.price || "?"}
DM to grab it before it's gone 👀`;

  const grailed = `${title}
Condition: ${condLine}
${attrLines.join("\n")}
${p.remark ? `Notes: ${p.remark}` : ""}

SKU: ${p.sku || "-"}
Message with any questions before buying.`;

  const depop = `${title} 🧵
condition ${condLine}${attrLines.length ? " · " + attrLines.join(" · ") : ""}
${p.remark ? p.remark + " " : ""}💌 bundle for a deal`;

  const shopee = `${title}
${attrLines.map((l) => "• " + l).join("\n")}
• Kondisi: ${condLine}
${p.remark ? `• Catatan: ${p.remark}\n` : ""}SKU: ${p.sku || "-"}`;

  return { carousell, threads, grailed, depop, shopee, sizeFlag: null };
}

// tokens available to a custom template: {title} {brand} {category} {condition} {colour} {remark} {price} {sku} {size} {measurements}
function getDisplaySize(p) {
  if (p.category !== "apparel") return "";
  const a = p.attributes || {};
  const sub = getApparelSubcategory(p.itemType);
  return a.size || computeSuggestedSize(a, sub) || "";
}
function getMeasurementsString(p) {
  if (p.category === "apparel") {
    const a = p.attributes || {};
    const sub = getApparelSubcategory(p.itemType);
    return apparelMeasureParts(a, sub).join(" | ");
  }
  return getFilledAttrLines(p).join(" | ");
}
const TEMPLATE_TOKENS = ["title", "brand", "category", "condition", "colour", "colourCondition", "remark", "price", "sku", "size", "measurements", "chest", "length", "waist", "frontRise", "thigh", "shoulder", "sleeve", "legOpening"];
function renderCustomTemplate(str, p) {
  const a = p.attributes || {};
  const tokens = {
    title: productTitle(p), brand: p.brand || "", category: CATEGORY_LABELS[p.category] || "",
    condition: `${p.condition || "?"}/10`, colour: p.colour || "", colourCondition: p.colourCondition || "", remark: p.remark || "",
    price: p.price || "", sku: p.sku || "", size: getDisplaySize(p), measurements: getMeasurementsString(p),
    chest: a.chest || "", length: a.length || "", waist: a.waist || "", frontRise: a.frontRise || "",
    thigh: a.thigh || "", shoulder: a.shoulder || "", sleeve: a.sleeve || "", legOpening: a.legOpening || "",
  };
  return str.replace(/\{(\w+)\}/g, (match, key) => (key in tokens ? tokens[key] : match));
}

// reverse direction: given a reference item's real values and text the user already wrote about it,
// finds each value inside the text and swaps it for the matching {token} — best-effort, meant to be
// reviewed and touched up before saving, not guaranteed to catch every occurrence.
// individual measurement numbers are matched on their own (not just as one combined string), since
// people naturally write measurements in their own phrasing rather than Ravuno's exact format
function generateTemplateFromExample(pastedText, item) {
  let result = pastedText;
  const a = item.attributes || {};
  const size = getDisplaySize(item);
  const title = productTitle(item);
  const catLabel = CATEGORY_LABELS[item.category];

  // longer, more specific forms are tried before shorter/riskier ones
  const candidates = [];
  if (item.condition) candidates.push([`${item.condition}/10`, "condition"]);
  if (item.colourCondition) candidates.push([`${item.colourCondition}/10`, "colourCondition"]);
  if (title) candidates.push([title, "title"]);
  if (item.remark) candidates.push([item.remark, "remark"]);
  if (item.sku) candidates.push([String(item.sku), "sku"]);
  if (item.brand) candidates.push([item.brand, "brand"]);
  if (item.colour) candidates.push([item.colour, "colour"]);
  if (catLabel) candidates.push([catLabel, "category"]);
  if (size) candidates.push([size, "size"]);
  // individual measurement values — matched as standalone numbers regardless of surrounding wording
  if (a.chest) candidates.push([String(a.chest), "chest"]);
  if (a.length) candidates.push([String(a.length), "length"]);
  if (a.waist) candidates.push([String(a.waist), "waist"]);
  if (a.frontRise) candidates.push([String(a.frontRise), "frontRise"]);
  if (a.thigh) candidates.push([String(a.thigh), "thigh"]);
  if (a.shoulder) candidates.push([String(a.shoulder), "shoulder"]);
  if (a.sleeve) candidates.push([String(a.sleeve), "sleeve"]);
  if (a.legOpening) candidates.push([String(a.legOpening), "legOpening"]);
  if (item.price !== undefined && item.price !== "" && item.price !== null) candidates.push([String(item.price), "price"]);
  if (item.condition) candidates.push([String(item.condition), "condition"]); // bare-digit fallback, tried last
  if (item.colourCondition) candidates.push([String(item.colourCondition), "colourCondition"]); // bare-digit fallback, tried last

  candidates.sort((a, b) => b[0].length - a[0].length);

  const usedTokens = new Set();
  candidates.forEach(([value, token]) => {
    if (!value || usedTokens.has(token)) return;
    const escaped = value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    const re = new RegExp(`\\b${escaped}\\b`, "g");
    if (re.test(result)) {
      result = result.replace(re, `{${token}}`);
      usedTokens.add(token);
    }
  });
  return result;
}

function buildTemplates(p, customTemplates) {
  const defaults = p.category === "apparel" ? buildApparelTemplates(p) : buildGenericTemplates(p);
  if (!customTemplates) return defaults;
  const merged = { ...defaults };
  PLATFORMS.forEach((pl) => {
    const custom = customTemplates[pl.id];
    if (custom && custom.trim()) merged[pl.id] = renderCustomTemplate(custom, p);
  });
  return merged;
}

function daysSince(dateStr) {
  if (!dateStr) return 0;
  return Math.floor((new Date() - new Date(dateStr)) / (1000 * 60 * 60 * 24));
}

// compress an uploaded image to keep stored size reasonable
function compressImage(file, maxWidth = 500, quality = 0.7) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      const img = new Image();
      img.onload = () => {
        const scale = Math.min(1, maxWidth / img.width);
        const canvas = document.createElement("canvas");
        canvas.width = img.width * scale;
        canvas.height = img.height * scale;
        const ctx = canvas.getContext("2d");
        ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
        resolve(canvas.toDataURL("image/jpeg", quality));
      };
      img.onerror = reject;
      img.src = e.target.result;
    };
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}

// converts any uploaded logo to true black & white so it always matches the receipt design,
// regardless of what color image the user actually uploads
function compressImageGrayscale(file, maxWidth = 300, quality = 0.85) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      const img = new Image();
      img.onload = () => {
        const scale = Math.min(1, maxWidth / img.width);
        const canvas = document.createElement("canvas");
        canvas.width = img.width * scale;
        canvas.height = img.height * scale;
        const ctx = canvas.getContext("2d");
        ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
        const imageData = ctx.getImageData(0, 0, canvas.width, canvas.height);
        const d = imageData.data;
        for (let i = 0; i < d.length; i += 4) {
          const gray = d[i] * 0.299 + d[i + 1] * 0.587 + d[i + 2] * 0.114;
          d[i] = d[i + 1] = d[i + 2] = gray;
        }
        ctx.putImageData(imageData, 0, 0);
        resolve(canvas.toDataURL("image/png", quality));
      };
      img.onerror = reject;
      img.src = e.target.result;
    };
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}

function downloadDataUrl(dataUrl, filename) {
  const a = document.createElement("a");
  a.href = dataUrl;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
}
function sanitizeForFilename(s) {
  return (s || "item").replace(/[^a-zA-Z0-9-_]+/g, "_").replace(/^_+|_+$/g, "") || "item";
}
// downloads every photo for a listing as individually named files (no zip library available in this environment)
function downloadAllPhotos(item) {
  const photos = item.photos || [];
  const skuPart = sanitizeForFilename(item.sku);
  photos.forEach((p, idx) => {
    setTimeout(() => downloadDataUrl(p, `${skuPart}_${String(idx + 1).padStart(2, "0")}.jpg`), idx * 350);
  });
}

// ---------- small UI atoms ----------
function CopyButton({ text }) {
  const [copied, setCopied] = useState(false);
  return (
    <button
      onClick={async () => {
        try { await navigator.clipboard.writeText(text); setCopied(true); setTimeout(() => setCopied(false), 1500); }
        catch (e) { setCopied(false); }
      }}
      className="flex items-center gap-1.5 text-xs font-medium px-3 py-1.5 rounded-lg border transition-colors"
      style={{ borderColor: copied ? "#8C8A85" : "#F2F0EA", color: copied ? "#8C8A85" : "#F2F0EA", background: copied ? "#1C2420" : "transparent", fontFamily: "'IBM Plex Mono', monospace" }}
    >
      {copied ? <Check size={13} /> : <Copy size={13} />}
      {copied ? "Copied" : "Copy"}
    </button>
  );
}
function Field({ label, value, onChange, placeholder, type = "text" }) {
  return (
    <label className="block">
      <span className="block text-[11px] mb-1 tracking-wide uppercase" style={{ color: "#8C8A85", fontFamily: "'IBM Plex Mono', monospace" }}>{label}</span>
      <input type={type} value={value} onChange={onChange} placeholder={placeholder}
        className="w-full px-3 py-2 rounded-lg border text-sm outline-none"
        style={{ borderColor: "#2A2A2F", background: "#17171A", color: "#F2F0EA" }} />
    </label>
  );
}
function ReadOnlyField({ label, value, placeholder }) {
  return (
    <label className="block">
      <span className="block text-[11px] mb-1 tracking-wide uppercase" style={{ color: "#8C8A85", fontFamily: "'IBM Plex Mono', monospace" }}>{label}</span>
      <div className="w-full px-3 py-2 rounded-lg border text-sm"
        style={{ borderColor: "#2A2A2F", background: "#1E1E22", color: value ? "#F2F0EA" : "#75726C" }}>
        {value || placeholder}
      </div>
    </label>
  );
}
function Select({ label, value, onChange, children }) {
  return (
    <label className="block">
      <span className="block text-[11px] mb-1 tracking-wide uppercase" style={{ color: "#8C8A85", fontFamily: "'IBM Plex Mono', monospace" }}>{label}</span>
      <select value={value} onChange={onChange}
        className="w-full px-3 py-2 rounded-lg border text-sm outline-none"
        style={{ borderColor: "#2A2A2F", background: "#17171A", color: "#F2F0EA" }}>
        {children}
      </select>
    </label>
  );
}
function StatCard({ icon: Icon, label, value, tone }) {
  return (
    <div className="rounded-xl border px-4 py-3 flex items-center gap-3" style={{ borderColor: "#2A2A2F", background: "#17171A" }}>
      <div className="w-9 h-9 rounded-full flex items-center justify-center shrink-0" style={{ background: tone, color: "#FFFFFF" }}>
        <Icon size={16} />
      </div>
      <div>
        <p className="text-lg font-semibold leading-none" style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{value}</p>
        <p className="text-[11px] uppercase tracking-wide mt-1" style={{ color: "#8C8A85" }}>{label}</p>
      </div>
    </div>
  );
}
function DetailRow({ label, value }) {
  return (
    <div className="flex items-start justify-between gap-4 py-1 border-b" style={{ borderColor: "#232327" }}>
      <dt className="text-xs uppercase tracking-wide shrink-0" style={{ color: "#8C8A85", fontFamily: "'IBM Plex Mono', monospace" }}>{label}</dt>
      <dd className="text-right" style={{ color: "#F2F0EA" }}>{value}</dd>
    </div>
  );
}

// fullscreen photo viewer with prev/next and touch-swipe support
function Lightbox({ photos, index, onIndexChange, onClose, onDownload }) {
  const touchStartX = useRef(null);
  if (!photos || photos.length === 0) return null;
  const next = () => onIndexChange((index + 1) % photos.length);
  const prev = () => onIndexChange((index - 1 + photos.length) % photos.length);
  return (
    <div className="fixed inset-0 z-[70] flex items-center justify-center p-4" style={{ background: "rgba(20,17,14,0.9)" }} onClick={onClose}>
      <button onClick={onClose} className="absolute top-4 right-4 w-9 h-9 rounded-full flex items-center justify-center" style={{ background: "#17171A", color: "#F2F0EA" }} aria-label="Close">
        <X size={16} />
      </button>
      {photos.length > 1 && (
        <button onClick={(e) => { e.stopPropagation(); prev(); }} className="absolute left-2 sm:left-6 w-9 h-9 rounded-full flex items-center justify-center" style={{ background: "#17171A", color: "#F2F0EA" }} aria-label="Previous photo">
          <ChevronLeft size={18} />
        </button>
      )}
      <img
        src={photos[index]}
        alt=""
        onClick={(e) => e.stopPropagation()}
        onTouchStart={(e) => { touchStartX.current = e.touches[0].clientX; }}
        onTouchEnd={(e) => {
          if (touchStartX.current == null) return;
          const dx = e.changedTouches[0].clientX - touchStartX.current;
          if (dx > 50) prev(); else if (dx < -50) next();
          touchStartX.current = null;
        }}
        className="max-h-[80vh] max-w-full object-contain rounded-lg"
      />
      {photos.length > 1 && (
        <button onClick={(e) => { e.stopPropagation(); next(); }} className="absolute right-2 sm:right-6 w-9 h-9 rounded-full flex items-center justify-center" style={{ background: "#17171A", color: "#F2F0EA" }} aria-label="Next photo">
          <ChevronRight size={18} />
        </button>
      )}
      {onDownload && (
        <button onClick={(e) => { e.stopPropagation(); onDownload(index); }}
          className="absolute bottom-6 flex items-center gap-1.5 text-xs font-semibold px-3 py-2 rounded-lg"
          style={{ background: "#17171A", color: "#F2F0EA", fontFamily: "'IBM Plex Mono', monospace" }}>
          <Download size={13} /> Download this photo
        </button>
      )}
    </div>
  );
}

function PhotoGalleryPicker({ photos, onChange, onPreview }) {
  const fileRef = useRef(null);
  const remaining = MAX_PHOTOS - photos.length;
  async function handleFiles(fileList) {
    const files = Array.from(fileList).slice(0, remaining);
    if (files.length === 0) return;
    const compressed = await Promise.all(files.map((f) => compressImage(f)));
    onChange([...photos, ...compressed]);
  }
  function removeAt(idx) { onChange(photos.filter((_, i) => i !== idx)); }
  function makePrimary(idx) {
    if (idx === 0) return;
    const next = [...photos];
    const [chosen] = next.splice(idx, 1);
    next.unshift(chosen);
    onChange(next);
  }
  return (
    <div>
      <span className="block text-[11px] mb-1 tracking-wide uppercase" style={{ color: "#8C8A85", fontFamily: "'IBM Plex Mono', monospace" }}>
        Photos ({photos.length}/{MAX_PHOTOS})
      </span>
      <div className="flex flex-wrap gap-2">
        {photos.map((p, idx) => (
          <div key={idx} className="relative w-16 h-16">
            <img src={p} alt="" onClick={() => onPreview && onPreview(idx)}
              className="w-16 h-16 object-cover rounded-lg border-2 cursor-pointer"
              style={{ borderColor: idx === 0 ? "#C9A567" : "#2A2A2F" }} />
            {idx === 0 ? (
              <span className="absolute -top-1.5 -left-1.5 text-[8px] font-bold px-1 py-0.5 rounded" style={{ background: "#C9A567", color: "#0B0B0C" }}>MAIN</span>
            ) : (
              <button onClick={() => makePrimary(idx)} title="Set as primary photo"
                className="absolute -bottom-1.5 -left-1.5 w-5 h-5 rounded-full flex items-center justify-center" style={{ background: "#17171A", color: "#8C8A85", border: "1px solid #2A2A2F" }}>
                <Star size={10} />
              </button>
            )}
            <button onClick={() => removeAt(idx)} className="absolute -top-1.5 -right-1.5 w-5 h-5 rounded-full flex items-center justify-center" style={{ background: "#C9695E", color: "#FFFFFF" }}>
              <X size={11} />
            </button>
          </div>
        ))}
        {photos.length < MAX_PHOTOS && (
          <button onClick={() => fileRef.current?.click()}
            className="w-16 h-16 rounded-lg border-2 border-dashed flex flex-col items-center justify-center gap-0.5 text-[10px]"
            style={{ borderColor: "#2A2A2F", color: "#75726C" }}>
            <ImageIcon size={16} />Add
          </button>
        )}
      </div>
      <input ref={fileRef} type="file" accept="image/*" multiple className="hidden"
        onChange={(e) => { handleFiles(e.target.files); e.target.value = ""; }} />
    </div>
  );
}

// ---------- category-aware form sections ----------
function ApparelFields({ form, setForm }) {
  const sub = getApparelSubcategory(form.itemType);
  const a = form.attributes;
  const suggested = computeSuggestedSize(a, sub);
  const sizeFlag = sub === "top" ? estimateSizeFlag(a.chest, a.length) : null;

  function handleItemTypeChange(e) {
    const nextType = e.target.value;
    const nextSub = getApparelSubcategory(nextType);
    const prevSub = getApparelSubcategory(form.itemType);
    setForm((f) => (nextSub === prevSub ? { ...f, itemType: nextType } : { ...f, itemType: nextType, attributes: {} }));
  }
  function setAttr(id, val) { setForm((f) => ({ ...f, attributes: { ...f.attributes, [id]: val } })); }

  return (
    <>
      <Select label="Item type" value={form.itemType} onChange={handleItemTypeChange}>
        <option value="">Select type…</option>
        <optgroup label="Tops">{APPAREL_TOP_TYPES.map((t) => <option key={t} value={t}>{t}</option>)}</optgroup>
        <optgroup label="Bottoms">{APPAREL_BOTTOM_TYPES.map((t) => <option key={t} value={t}>{t}</option>)}</optgroup>
      </Select>

      <div className="grid grid-cols-2 gap-3">
        <Field label="Size" value={a.size || ""} onChange={(e) => setAttr("size", e.target.value)} placeholder={suggested || "M"} />
        <ReadOnlyField label="Suggested size" value={suggested} placeholder="from measurements" />
      </div>

      {sub === "top" ? (
        <>
          <div className="grid grid-cols-2 gap-3">
            <Field label="Chest (in)" value={a.chest || ""} onChange={(e) => setAttr("chest", e.target.value)} placeholder="21" type="number" />
            <Field label="Length (in)" value={a.length || ""} onChange={(e) => setAttr("length", e.target.value)} placeholder="27" type="number" />
          </div>
          <div className="grid grid-cols-2 gap-3">
            <Field label="Shoulder (in)" value={a.shoulder || ""} onChange={(e) => setAttr("shoulder", e.target.value)} placeholder="optional" type="number" />
            <Field label="Sleeve (in)" value={a.sleeve || ""} onChange={(e) => setAttr("sleeve", e.target.value)} placeholder="optional" type="number" />
          </div>
        </>
      ) : (
        <>
          <div className="grid grid-cols-2 gap-3">
            <Field label="Waist (in)" value={a.waist || ""} onChange={(e) => setAttr("waist", e.target.value)} placeholder="32" type="number" />
            <Field label="Length (in)" value={a.length || ""} onChange={(e) => setAttr("length", e.target.value)} placeholder="40" type="number" />
          </div>
          <div className="grid grid-cols-2 gap-3">
            <Field label="Front rise (in)" value={a.frontRise || ""} onChange={(e) => setAttr("frontRise", e.target.value)} placeholder="12" type="number" />
            <Field label="Thigh (in)" value={a.thigh || ""} onChange={(e) => setAttr("thigh", e.target.value)} placeholder="12" type="number" />
          </div>
          <Field label="Leg opening (in)" value={a.legOpening || ""} onChange={(e) => setAttr("legOpening", e.target.value)} placeholder="optional" type="number" />
        </>
      )}

      {sizeFlag && (
        <div className="text-xs px-3 py-2 rounded-lg flex items-start gap-2" style={{ background: "#2A1715", color: "#C9695E" }}>
          <Ruler size={14} className="mt-0.5 shrink-0" />{sizeFlag}
        </div>
      )}
    </>
  );
}

function DynamicAttributeFields({ defs, values, onChange }) {
  return (
    <div className="grid grid-cols-2 gap-3">
      {defs.map((d) => (
        <Field key={d.id} label={d.label} value={values[d.id] || ""} onChange={(e) => onChange(d.id, e.target.value)} type={d.type === "number" ? "number" : "text"} placeholder="" />
      ))}
    </div>
  );
}

function CustomAttributesEditor({ value, onChange }) {
  function update(idx, key, val) { onChange(value.map((c, i) => (i === idx ? { ...c, [key]: val } : c))); }
  function add() { onChange([...value, { label: "", value: "" }]); }
  function remove(idx) { onChange(value.filter((_, i) => i !== idx)); }
  return (
    <div className="space-y-2">
      <span className="block text-[11px] mb-1 tracking-wide uppercase" style={{ color: "#8C8A85", fontFamily: "'IBM Plex Mono', monospace" }}>Custom attributes</span>
      {value.map((c, idx) => (
        <div key={idx} className="grid grid-cols-[1fr_1fr_auto] gap-2">
          <input value={c.label} onChange={(e) => update(idx, "label", e.target.value)} placeholder="Field name"
            className="px-3 py-2 rounded-lg border text-sm outline-none" style={{ borderColor: "#2A2A2F", background: "#17171A", color: "#F2F0EA" }} />
          <input value={c.value} onChange={(e) => update(idx, "value", e.target.value)} placeholder="Value"
            className="px-3 py-2 rounded-lg border text-sm outline-none" style={{ borderColor: "#2A2A2F", background: "#17171A", color: "#F2F0EA" }} />
          <button onClick={() => remove(idx)}><X size={14} style={{ color: "#C9695E" }} /></button>
        </div>
      ))}
      <button onClick={add} className="text-xs font-medium px-3 py-1.5 rounded-lg border flex items-center gap-1"
        style={{ borderColor: "#2A2A2F", color: "#F2F0EA", fontFamily: "'IBM Plex Mono', monospace" }}>
        <Plus size={12} /> Add field
      </button>
    </div>
  );
}

function TemplatesTab({ customTemplates, onSave, saving, logoUrl, onLogoUpload, onLogoRemove, logoUploading, items }) {
  const [draft, setDraft] = useState(customTemplates || {});
  useEffect(() => { setDraft(customTemplates || {}); }, [customTemplates]);
  const logoInputRef = useRef(null);
  const [refItemId, setRefItemId] = useState("");
  const [pasteOpenFor, setPasteOpenFor] = useState(null); // platform id currently showing the paste box
  const [pasteText, setPasteText] = useState("");

  const refItem = items?.find((i) => i.id === refItemId) || null;

  function handleConvert(platformId) {
    if (!refItem || !pasteText.trim()) return;
    const generated = generateTemplateFromExample(pasteText, refItem);
    setDraft((d) => ({ ...d, [platformId]: generated }));
    setPasteOpenFor(null);
    setPasteText("");
  }

  const sampleProduct = {
    title: "Uniqlo Fleece Jacket", brand: "Uniqlo", category: "apparel", itemType: "Jacket",
    condition: "9", colour: "Navy", remark: "small pilling near cuff", price: "45", sku: "TZ-014",
    attributes: { chest: "21", length: "27" },
  };

  return (
    <div className="max-w-2xl mx-auto space-y-6">
      <div className="rounded-xl border p-5" style={{ background: "#17171A", borderColor: "#2A2A2F" }}>
        <h2 className="text-sm font-semibold mb-2" style={{ color: "#F2F0EA", fontFamily: "'Fraunces', serif" }}>Business logo</h2>
        <p className="text-xs mb-3" style={{ color: "#8C8A85" }}>
          Upload a logo to show at the top of your printed receipts. Automatically converted to black &amp; white to
          keep receipts clean and print-friendly, regardless of the original image.
        </p>
        <div className="flex items-center gap-4">
          {logoUrl ? (
            <img src={logoUrl} alt="Your logo" className="w-16 h-16 rounded-lg object-contain p-1" style={{ background: "#FFFFFF" }} />
          ) : (
            <div className="w-16 h-16 rounded-lg border-2 border-dashed flex items-center justify-center" style={{ borderColor: "#2A2A2F", color: "#75726C" }}>
              <ImageIcon size={20} />
            </div>
          )}
          <div className="flex items-center gap-2">
            <button onClick={() => logoInputRef.current?.click()} disabled={logoUploading}
              className="text-xs font-semibold px-3 py-2 rounded-lg" style={{ background: "#C9A567", color: "#0B0B0C", opacity: logoUploading ? 0.6 : 1 }}>
              {logoUploading ? "Uploading…" : logoUrl ? "Replace logo" : "Upload logo"}
            </button>
            {logoUrl && (
              <button onClick={onLogoRemove} disabled={logoUploading} className="text-xs font-medium px-3 py-2 rounded-lg border" style={{ borderColor: "#2A2A2F", color: "#C9695E" }}>
                Remove
              </button>
            )}
          </div>
        </div>
        <input ref={logoInputRef} type="file" accept="image/*" className="hidden"
          onChange={(e) => { const f = e.target.files?.[0]; if (f) onLogoUpload(f); e.target.value = ""; }} />
      </div>

      <div className="rounded-xl border p-5" style={{ background: "#17171A", borderColor: "#2A2A2F" }}>
        <h2 className="text-sm font-semibold mb-2" style={{ color: "#F2F0EA", fontFamily: "'Fraunces', serif" }}>Your templates</h2>
        <p className="text-xs mb-3" style={{ color: "#8C8A85" }}>
          Write your own listing text for each platform. Leave a box empty to use Ravuno's built-in default instead.
        </p>
        <div className="flex flex-wrap gap-1.5 mb-4">
          {TEMPLATE_TOKENS.map((t) => (
            <span key={t} className="text-[10px] font-medium px-2 py-1 rounded-md" style={{ background: "#0B0B0C", color: "#8C8A85", fontFamily: "'IBM Plex Mono', monospace" }}>{`{${t}}`}</span>
          ))}
        </div>
        <label className="block">
          <span className="block text-[11px] mb-1 tracking-wide uppercase" style={{ color: "#75726C", fontFamily: "'IBM Plex Mono', monospace" }}>
            Reference item (for auto-convert below)
          </span>
          <select value={refItemId} onChange={(e) => setRefItemId(e.target.value)}
            className="w-full px-3 py-2 rounded-lg border text-sm outline-none" style={{ borderColor: "#2A2A2F", background: "#0B0B0C", color: "#F2F0EA" }}>
            <option value="">Select a saved item…</option>
            {(items || []).map((i) => <option key={i.id} value={i.id}>{productTitle(i)} {i.sku ? `(${i.sku})` : ""}</option>)}
          </select>
        </label>
      </div>

      {PLATFORMS.map((p) => (
        <div key={p.id} className="rounded-xl border p-5" style={{ background: "#17171A", borderColor: "#2A2A2F" }}>
          <div className="flex items-center justify-between mb-2">
            <span className="text-sm font-semibold" style={{ color: "#F2F0EA" }}>{p.label}</span>
            {draft[p.id] && (
              <button onClick={() => setDraft((d) => ({ ...d, [p.id]: "" }))} className="text-xs font-medium" style={{ color: "#C9695E" }}>
                Reset to default
              </button>
            )}
          </div>
          <textarea
            value={draft[p.id] || ""}
            onChange={(e) => setDraft((d) => ({ ...d, [p.id]: e.target.value }))}
            placeholder={`{title}\n\nCondition: {condition}\n{measurements}\nRemark: {remark}\n\nPrice: RM{price}\nSKU: {sku}`}
            rows={6}
            className="w-full px-3 py-2 rounded-lg border text-xs outline-none resize-y"
            style={{ borderColor: "#2A2A2F", background: "#17171A", color: "#F2F0EA", fontFamily: "'IBM Plex Mono', monospace" }}
          />

          {pasteOpenFor === p.id ? (
            <div className="mt-2 rounded-lg p-3" style={{ background: "#0B0B0C", border: "1px solid #2A2A2F" }}>
              <p className="text-[11px] mb-2" style={{ color: "#75726C" }}>
                {refItem ? `Paste how you already wrote up "${productTitle(refItem)}" — we'll try to detect and swap in the tokens automatically.` : "Pick a reference item above first."}
              </p>
              <textarea
                value={pasteText}
                onChange={(e) => setPasteText(e.target.value)}
                rows={4}
                placeholder="Paste your own listing text here…"
                className="w-full px-3 py-2 rounded-lg border text-xs outline-none resize-y"
                style={{ borderColor: "#2A2A2F", background: "#17171A", color: "#F2F0EA" }}
              />
              <div className="flex gap-2 mt-2">
                <button onClick={() => handleConvert(p.id)} disabled={!refItem || !pasteText.trim()}
                  className="text-xs font-semibold px-3 py-1.5 rounded-lg" style={{ background: "#C9A567", color: "#0B0B0C", opacity: (!refItem || !pasteText.trim()) ? 0.5 : 1 }}>
                  Convert &amp; fill in above
                </button>
                <button onClick={() => { setPasteOpenFor(null); setPasteText(""); }} className="text-xs font-medium px-3 py-1.5 rounded-lg border" style={{ borderColor: "#2A2A2F", color: "#8C8A85" }}>
                  Cancel
                </button>
              </div>
            </div>
          ) : (
            <button onClick={() => setPasteOpenFor(p.id)} className="text-xs font-medium mt-2" style={{ color: "#C9A567" }}>
              ✨ Convert from my own text
            </button>
          )}

          {draft[p.id] && draft[p.id].trim() && (
            <div className="mt-2 rounded-lg p-3" style={{ background: "#0B0B0C" }}>
              <p className="text-[10px] uppercase tracking-wide mb-1" style={{ color: "#75726C" }}>Preview</p>
              <pre className="text-[11px] whitespace-pre-wrap" style={{ fontFamily: "'IBM Plex Mono', monospace", color: "#F2F0EA" }}>
                {renderCustomTemplate(draft[p.id], sampleProduct)}
              </pre>
            </div>
          )}
        </div>
      ))}

      <button onClick={() => onSave(draft)} disabled={saving}
        className="w-full flex items-center justify-center gap-2 py-2.5 rounded-xl text-sm font-semibold transition-opacity"
        style={{ background: "#C9A567", color: "#0B0B0C", opacity: saving ? 0.6 : 1, fontFamily: "'IBM Plex Mono', monospace" }}>
        {saving ? "Saving…" : "Save templates"}
      </button>
    </div>
  );
}

const emptyForm = {
  title: "", titleTouched: false,
  category: "", itemType: "",
  brand: "", condition: "", colour: "", colourCondition: "", remark: "", cost: "", price: "", sku: "",
  photos: [], attributes: {}, customAttributes: [],
};

export default function App() {
  const [session, setSession] = useState(undefined); // undefined = still checking, null = signed out
  const [authMode, setAuthMode] = useState("signin"); // "signin" | "signup"
  const [authEmail, setAuthEmail] = useState("");
  const [authPassword, setAuthPassword] = useState("");
  const [authBusy, setAuthBusy] = useState(false);
  const [authError, setAuthError] = useState("");
  const [authNotice, setAuthNotice] = useState("");
  const [profile, setProfile] = useState(null);
  const [customTemplates, setCustomTemplates] = useState({});
  const [templatesSaving, setTemplatesSaving] = useState(false);
  const [feedbackText, setFeedbackText] = useState("");
  const [feedbackStatus, setFeedbackStatus] = useState("idle"); // idle | sending | sent | error
  const [receiptTxn, setReceiptTxn] = useState(null);
  const [logoUrl, setLogoUrl] = useState(null);
  const [logoUploading, setLogoUploading] = useState(false);
  const [upgradeBusy, setUpgradeBusy] = useState(false);
  const [showAuthForm, setShowAuthForm] = useState(false);
  const [landingPreview, setLandingPreview] = useState(null); // "crosslist" | "inventory" | "pos" | null

  const [tab, setTab] = useState("crosslist");
  const [items, setItems] = useState([]);
  const [txns, setTxns] = useState([]);
  const [loading, setLoading] = useState(true);
  const [saveError, setSaveError] = useState(false);
  const [form, setForm] = useState(emptyForm);
  const [activeId, setActiveId] = useState(null);
  const [editingId, setEditingId] = useState(null);
  const [skuSearch, setSkuSearch] = useState("");
  const [cart, setCart] = useState([]);
  const [payMethod, setPayMethod] = useState("Cash");
  const [ledgerSearch, setLedgerSearch] = useState("");
  const [ledgerStatusFilter, setLedgerStatusFilter] = useState("all"); // all | listed | sold
  const [csvImporting, setCsvImporting] = useState(false);
  const [csvImportResult, setCsvImportResult] = useState(null); // { added, skipped } | null
  const csvInputRef = useRef(null);
  const [crosslistSearch, setCrosslistSearch] = useState("");
  const [detailId, setDetailId] = useState(null);
  const [galleryIndex, setGalleryIndex] = useState(0);
  const [lightbox, setLightbox] = useState(null); // { photos, index }
  const now = new Date();
  const [salesPeriodType, setSalesPeriodType] = useState("monthly"); // "monthly" | "yearly"
  const [salesMonth, setSalesMonth] = useState(now.getMonth());
  const [salesYear, setSalesYear] = useState(now.getFullYear());
  const [selectedInvoiceId, setSelectedInvoiceId] = useState(null);

  // check for an existing session on load
  useEffect(() => {
    const stored = loadStoredSession();
    setSession(stored || null);
  }, []);

  // once signed in, load this user's items + transactions from Supabase
  useEffect(() => {
    if (!session) { setLoading(false); return; }
    (async () => {
      setLoading(true);
      try {
        const [itemRows, txnRows, profileRows, templateRows, brandingRows] = await Promise.all([
          apiRequest("/items?select=*&order=created_at.desc", { method: "GET" }, session, updateSession),
          apiRequest("/transactions?select=*&order=occurred_at.desc", { method: "GET" }, session, updateSession),
          apiRequest("/profiles?select=*", { method: "GET" }, session, updateSession).catch(() => []),
          apiRequest("/templates?select=*", { method: "GET" }, session, updateSession).catch(() => []),
          apiRequest("/branding?select=*", { method: "GET" }, session, updateSession).catch(() => []),
        ]);
        setItems((itemRows || []).map(rowToItem));
        setTxns((txnRows || []).map(rowToTxn));
        setProfile((profileRows && profileRows[0]) || null);
        const tRow = templateRows && templateRows[0];
        setCustomTemplates(tRow ? { carousell: tRow.carousell, threads: tRow.threads, grailed: tRow.grailed, depop: tRow.depop, shopee: tRow.shopee } : {});
        setLogoUrl((brandingRows && brandingRows[0]?.logo_url) || null);
      } catch (e) {
        setSaveError(true);
      } finally {
        setLoading(false);
      }
    })();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [session?.access_token]);

  async function saveTemplates(next) {
    if (!session) return;
    setTemplatesSaving(true);
    try {
      const row = { user_id: session.user.id, ...next, updated_at: new Date().toISOString() };
      await apiRequest("/templates?on_conflict=user_id", {
        method: "POST", body: JSON.stringify(row), headers: { Prefer: "resolution=merge-duplicates,return=representation" },
      }, session, updateSession);
      setCustomTemplates(next);
    } catch (e) {
      setSaveError(true);
    } finally {
      setTemplatesSaving(false);
    }
  }

  async function handleLogoUpload(file) {
    if (!session || !file) return;
    setLogoUploading(true);
    try {
      const grayscaleDataUrl = await compressImageGrayscale(file);
      const publicUrl = await uploadPhotoToStorage(grayscaleDataUrl, `${session.user.id}/logo.png`, session);
      const row = { user_id: session.user.id, logo_url: publicUrl, updated_at: new Date().toISOString() };
      await apiRequest("/branding?on_conflict=user_id", {
        method: "POST", body: JSON.stringify(row), headers: { Prefer: "resolution=merge-duplicates,return=representation" },
      }, session, updateSession);
      setLogoUrl(publicUrl);
    } catch (e) {
      setSaveError(true);
    } finally {
      setLogoUploading(false);
    }
  }

  async function handleLogoRemove() {
    if (!session) return;
    setLogoUploading(true);
    try {
      await apiRequest("/branding?on_conflict=user_id", {
        method: "POST",
        body: JSON.stringify({ user_id: session.user.id, logo_url: null, updated_at: new Date().toISOString() }),
        headers: { Prefer: "resolution=merge-duplicates,return=representation" },
      }, session, updateSession);
      setLogoUrl(null);
    } catch (e) {
      setSaveError(true);
    } finally {
      setLogoUploading(false);
    }
  }

  async function handleSendFeedback() {
    if (!feedbackText.trim() || !session) return;
    setFeedbackStatus("sending");
    try {
      const res = await fetch("/api/send-feedback", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ message: feedbackText, email: session.user.email }),
      });
      const data = await res.json().catch(() => ({}));
      if (data.success) { setFeedbackStatus("sent"); setFeedbackText(""); }
      else { setFeedbackStatus("error"); }
    } catch (e) {
      setFeedbackStatus("error");
    }
  }

  async function handleSupport(amount) {
    if (!session) return;
    setUpgradeBusy(true);
    try {
      const res = await fetch("/api/create-checkout", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ amount, userId: session.user.id, email: session.user.email }),
      });
      const data = await res.json().catch(() => ({}));
      if (data.url) window.location.href = data.url;
      else { alert(`Checkout failed: ${data.error || "unknown error"}`); setSaveError(true); }
    } catch (e) {
      alert(`Checkout failed: ${e.message}`);
      setSaveError(true);
    } finally {
      setUpgradeBusy(false);
    }
  }

  function updateSession(next) {
    setSession(next);
    saveStoredSession(next);
  }

  async function handleAuthSubmit(e) {
    e.preventDefault();
    setAuthError(""); setAuthNotice(""); setAuthBusy(true);
    try {
      if (authMode === "signup") {
        const data = await signUpRequest(authEmail, authPassword);
        if (data.access_token) { updateSession(data); }
        else { setAuthNotice("Check your email to confirm your account, then sign in."); setAuthMode("signin"); }
      } else {
        const data = await signInRequest(authEmail, authPassword);
        updateSession(data);
      }
      setAuthPassword("");
    } catch (err) {
      setAuthError(err.message);
    } finally {
      setAuthBusy(false);
    }
  }

  async function handleSignOut() {
    if (session) await signOutRequest(session);
    updateSession(null);
    setItems([]); setTxns([]); setForm(emptyForm); setCart([]); setActiveId(null);
    setShowAuthForm(false);
  }

  // keep the product title in sync with brand + item type until the user edits it directly
  useEffect(() => {
    if (!form.titleTouched) {
      const auto = `${form.brand || ""} ${form.itemType || ""}`.trim();
      if (auto !== form.title) setForm((f) => ({ ...f, title: auto }));
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [form.brand, form.itemType]);

  const set = (k) => (e) => setForm((f) => ({ ...f, [k]: e.target.value }));

  function handleCategoryChange(e) {
    setForm((f) => ({ ...f, category: e.target.value, itemType: "", attributes: {}, customAttributes: [] }));
  }

  const preview = useMemo(() => buildTemplates(form, customTemplates), [form, customTemplates]);
  const activeListing = items.find((l) => l.id === activeId) || null;
  const activeTemplates = activeListing ? buildTemplates(activeListing, customTemplates) : preview;
  const isFormUsable = form.category.trim() && form.itemType.trim();

  async function saveListing() {
    if (!isFormUsable || !session) return;
    setSaveError(false);
    const isEditing = !!editingId;
    const draft = isEditing
      ? { ...form }
      : { ...form, status: "listed", addedOn: new Date().toISOString().slice(0, 10), listedOn: new Date().toISOString().slice(0, 10) };
    try {
      // upload any photos first, so the saved row stores real URLs rather than base64 blobs
      const uploadedPhotos = await Promise.all(
        (draft.photos || []).map((p, idx) => (p.startsWith("data:") ? uploadPhotoToStorage(p, `${session.user.id}/${Date.now()}_${idx}.jpg`, session) : p))
      );
      if (isEditing) {
        // partial update: only editable fields, never touching status/dates so a sold item can't be silently reverted
        const editRow = {
          title: draft.title || null, category: draft.category || null, item_type: draft.itemType || null,
          brand: draft.brand || null, condition: draft.condition || null, colour: draft.colour || null, colour_condition: draft.colourCondition || null, remark: draft.remark || null,
          cost: draft.cost === "" ? null : draft.cost, price: draft.price === "" ? null : draft.price, sku: draft.sku || null,
          attributes: draft.attributes || {}, custom_attributes: draft.customAttributes || [], photos: uploadedPhotos,
        };
        const [updated] = await apiRequest(`/items?id=eq.${editingId}`, { method: "PATCH", body: JSON.stringify(editRow) }, session, updateSession);
        const updatedItem = rowToItem(updated);
        setItems((prev) => prev.map((i) => (i.id === editingId ? updatedItem : i)));
        setActiveId(updatedItem.id);
        setEditingId(null);
      } else {
        const row = itemToRow({ ...draft, photos: uploadedPhotos }, session.user.id);
        const [inserted] = await apiRequest("/items", { method: "POST", body: JSON.stringify(row) }, session, updateSession);
        const newItem = rowToItem(inserted);
        setItems((prev) => [newItem, ...prev]);
        setActiveId(newItem.id);
      }
      setForm(emptyForm);
    } catch (e) {
      setSaveError(true);
    }
  }

  function startEditItem(item) {
    setForm({
      title: item.title || "", titleTouched: true,
      category: item.category || "", itemType: item.itemType || "",
      brand: item.brand || "", condition: item.condition || "", colour: item.colour || "", colourCondition: item.colourCondition || "",
      remark: item.remark || "", cost: item.cost ?? "", price: item.price ?? "", sku: item.sku || "",
      photos: item.photos || [], attributes: item.attributes || {}, customAttributes: item.customAttributes || [],
    });
    setEditingId(item.id);
    setTab("crosslist");
  }

  function cancelEdit() {
    setEditingId(null);
    setForm(emptyForm);
  }

  async function removeItem(id) {
    setItems((prev) => prev.filter((l) => l.id !== id));
    if (activeId === id) setActiveId(null);
    try { await apiRequest(`/items?id=eq.${id}`, { method: "DELETE" }, session, updateSession); }
    catch (e) { setSaveError(true); }
  }

  // --- POS: search, cart, checkout ---
  const availableStock = items.filter((i) => i.status !== "sold");
  const searchResults = useMemo(() => {
    const q = skuSearch.trim().toLowerCase();
    if (!q) return []; // search-to-reveal: nothing shown until the user searches
    return availableStock.filter((i) =>
      (i.sku || "").toLowerCase().includes(q) ||
      (i.brand || "").toLowerCase().includes(q) ||
      (i.itemType || "").toLowerCase().includes(q) ||
      productTitle(i).toLowerCase().includes(q)
    );
  }, [skuSearch, items]);

  function addToCart(item) { setCart((prev) => (prev.some((c) => c.id === item.id) ? prev : [...prev, { id: item.id, discount: 0 }])); }
  function removeFromCart(id) { setCart((prev) => prev.filter((c) => c.id !== id)); }
  function updateDiscount(id, discount) {
    const d = Math.max(0, parseFloat(discount) || 0);
    setCart((prev) => prev.map((c) => (c.id === id ? { ...c, discount: d } : c)));
  }
  // discount is transaction-only — never written back to the product's stored price
  const cartLines = cart.map((c) => {
    const item = items.find((i) => i.id === c.id);
    if (!item) return null;
    const originalPrice = Number(item.price || 0);
    const discount = Math.min(c.discount || 0, originalPrice);
    const finalPrice = originalPrice - discount;
    return { ...item, discount, originalPrice, finalPrice };
  }).filter(Boolean);
  const cartSubtotal = cartLines.reduce((sum, l) => sum + l.originalPrice, 0);
  const cartDiscount = cartLines.reduce((sum, l) => sum + l.discount, 0);
  const cartTotal = cartSubtotal - cartDiscount;

  async function completeSale() {
    if (cartLines.length === 0 || !session) return;
    setSaveError(false);
    const soldOn = new Date().toISOString().slice(0, 10);
    const soldIds = cartLines.map((l) => l.id);
    const txnRow = txnToRow({
      timestamp: new Date().toISOString(), method: payMethod, total: cartTotal, subtotal: cartSubtotal, discount: cartDiscount,
      lines: cartLines.map((l) => ({ title: productTitle(l), sku: l.sku, price: l.price, cost: Number(l.cost || 0), discount: l.discount, finalPrice: l.finalPrice })),
    }, session.user.id);
    try {
      await apiRequest(`/items?id=in.(${soldIds.join(",")})`, { method: "PATCH", body: JSON.stringify({ status: "sold", sold_on: soldOn }) }, session, updateSession);
      const [insertedTxn] = await apiRequest("/transactions", { method: "POST", body: JSON.stringify(txnRow) }, session, updateSession);
      setItems((prev) => prev.map((i) => (soldIds.includes(i.id) ? { ...i, status: "sold", soldOn } : i)));
      setTxns((prev) => [rowToTxn(insertedTxn), ...prev]);
      setReceiptTxn({ ...rowToTxn(insertedTxn), invoiceNumber: `INV-${String(txns.length + 1).padStart(4, "0")}` });
      setCart([]);
    } catch (e) {
      setSaveError(true);
    }
  }

  // --- Sales: invoice numbering (stable, assigned by chronological order across all sales) ---
  const invoicedTxns = useMemo(() => {
    const chronological = [...txns].sort((a, b) => new Date(a.timestamp) - new Date(b.timestamp));
    const numberOf = new Map();
    chronological.forEach((t, idx) => numberOf.set(t.id, `INV-${String(idx + 1).padStart(4, "0")}`));
    return txns.map((t) => ({ ...t, invoiceNumber: numberOf.get(t.id) })); // preserves newest-first order
  }, [txns]);

  const availableSalesYears = useMemo(() => {
    const years = new Set(invoicedTxns.map((t) => new Date(t.timestamp).getFullYear()));
    years.add(now.getFullYear());
    return Array.from(years).sort((a, b) => b - a);
  }, [invoicedTxns]);

  const periodTxns = useMemo(() => {
    return invoicedTxns.filter((t) => {
      const d = new Date(t.timestamp);
      if (salesPeriodType === "monthly") return d.getFullYear() === salesYear && d.getMonth() === salesMonth;
      return d.getFullYear() === salesYear;
    });
  }, [invoicedTxns, salesPeriodType, salesYear, salesMonth]);

  const salesSummary = useMemo(() => {
    const totalSales = periodTxns.reduce((sum, t) => sum + Number(t.total || 0), 0);
    const orders = periodTxns.length;
    const unitsSold = periodTxns.reduce((sum, t) => sum + t.lines.length, 0);
    const discounts = periodTxns.reduce((sum, t) => sum + Number(t.discount || 0), 0);
    // net profit = final sale price minus cost, per line — sales made before this feature default cost to 0
    const netProfit = periodTxns.reduce((sum, t) => sum + t.lines.reduce((s, l) => s + (Number(l.finalPrice || 0) - Number(l.cost || 0)), 0), 0);
    return { totalSales, orders, unitsSold, discounts, netProfit };
  }, [periodTxns]);

  const salesGraphData = useMemo(() => {
    if (salesPeriodType === "monthly") {
      const daysInMonth = new Date(salesYear, salesMonth + 1, 0).getDate();
      const byDay = new Array(daysInMonth).fill(0);
      periodTxns.forEach((t) => { byDay[new Date(t.timestamp).getDate() - 1] += Number(t.total || 0); });
      return byDay.map((revenue, idx) => ({ label: String(idx + 1), revenue }));
    }
    const monthLabels = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
    const byMonth = new Array(12).fill(0);
    periodTxns.forEach((t) => { byMonth[new Date(t.timestamp).getMonth()] += Number(t.total || 0); });
    return byMonth.map((revenue, idx) => ({ label: monthLabels[idx], revenue }));
  }, [periodTxns, salesPeriodType, salesYear, salesMonth]);

  const selectedInvoice = invoicedTxns.find((t) => t.id === selectedInvoiceId) || null;
  const MONTH_NAMES = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];

  const stats = useMemo(() => {
    const listed = items.filter((i) => i.status === "listed");
    const sold = items.filter((i) => i.status === "sold");
    const profit = sold.reduce((sum, i) => sum + (Number(i.price || 0) - Number(i.cost || 0)), 0);
    const stale = listed.filter((i) => daysSince(i.addedOn) >= 21);
    return { listedCount: listed.length, soldCount: sold.length, profit, staleCount: stale.length };
  }, [items]);

  // --- Stock Ledger: search + detail + export ---
  const ledgerItems = useMemo(() => {
    const q = ledgerSearch.trim().toLowerCase();
    let base = items;
    if (ledgerStatusFilter === "listed") base = base.filter((i) => i.status !== "sold");
    if (ledgerStatusFilter === "sold") base = base.filter((i) => i.status === "sold");
    if (!q) return base;
    return base.filter((i) =>
      (i.sku || "").toLowerCase().includes(q) ||
      (i.brand || "").toLowerCase().includes(q) ||
      (i.itemType || "").toLowerCase().includes(q) ||
      (CATEGORY_LABELS[i.category] || "").toLowerCase().includes(q) ||
      productTitle(i).toLowerCase().includes(q) ||
      (i.status || "").toLowerCase().includes(q)
    );
  }, [items, ledgerSearch, ledgerStatusFilter]);

  const crosslistItems = useMemo(() => {
    const q = crosslistSearch.trim().toLowerCase();
    if (!q) return items;
    return items.filter((i) =>
      (i.sku || "").toLowerCase().includes(q) ||
      (i.brand || "").toLowerCase().includes(q) ||
      (CATEGORY_LABELS[i.category] || "").toLowerCase().includes(q) ||
      productTitle(i).toLowerCase().includes(q)
    );
  }, [items, crosslistSearch]);

  const detailItem = items.find((i) => i.id === detailId) || null;

  function exportStockTake(scope) {
    const list = scope === "listed" ? items.filter((i) => i.status !== "sold") : items;
    const dateStr = new Date().toLocaleDateString();
    const rows = [
      ["I SELL EVERYTHING"], ["STOCK TAKE"], [`Date: ${dateStr}`], [],
      ["No.", "SKU", "Title", "Category", "Item Type", "Cost (RM)", "Price (RM)", "Status", "COUNT"],
      ...list.map((i, idx) => [
        idx + 1, i.sku || "-", productTitle(i), CATEGORY_LABELS[i.category] || "-", i.itemType || "-",
        Number(i.cost || 0), Number(i.price || 0), (i.status || "listed").toUpperCase(), "",
      ]),
    ];
    const ws = XLSX.utils.aoa_to_sheet(rows);
    ws["!cols"] = [{ wch: 5 }, { wch: 10 }, { wch: 22 }, { wch: 14 }, { wch: 16 }, { wch: 11 }, { wch: 11 }, { wch: 10 }, { wch: 8 }];
    ws["!rows"] = rows.map((_, idx) => (idx < 4 ? { hpx: 18 } : idx === 4 ? { hpx: 22 } : { hpx: 24 }));
    ws["!merges"] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 8 } }, { s: { r: 1, c: 0 }, e: { r: 1, c: 8 } }, { s: { r: 2, c: 0 }, e: { r: 2, c: 8 } }];
    ws["!margins"] = { left: 0.4, right: 0.4, top: 0.5, bottom: 0.5, header: 0.2, footer: 0.2 };
    ws["!freeze"] = { xSplit: 0, ySplit: 5, topLeftCell: "A6", state: "frozen" };
    ws["!pageSetup"] = { orientation: "landscape", fitToWidth: 1, fitToHeight: 0 };
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Stock Take");
    XLSX.writeFile(wb, `stock-take-${scope}-${new Date().toISOString().slice(0, 10)}.xlsx`);
  }

  function downloadSampleImportCsv() {
    const rows = [
      CSV_IMPORT_COLUMNS,
      ["Uniqlo Fleece Jacket", "Apparel", "Jacket", "Uniqlo", "9", "Navy", "8", "small pilling near cuff", "15", "45", "TZ-014"],
      ["Casio Watch", "Watches", "Digital Watch", "Casio", "8", "Black", "", "", "60", "150", "TZ-015"],
    ];
    const csv = rows.map((r) => r.map((v) => `"${String(v).replace(/"/g, '""')}"`).join(",")).join("\n");
    const blob = new Blob([csv], { type: "text/csv" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url; a.download = "ravuno-import-template.csv";
    document.body.appendChild(a); a.click(); document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }

  function handleCsvImport(file) {
    if (!file || !session) return;
    setCsvImporting(true);
    setCsvImportResult(null);
    Papa.parse(file, {
      header: true,
      skipEmptyLines: true,
      complete: async (results) => {
        try {
          const rows = results.data;
          const today = new Date().toISOString().slice(0, 10);
          const validRows = rows.filter((r) => (r.title || r.Title || "").trim());
          const skipped = rows.length - validRows.length;
          const newRows = validRows.map((r) => {
            const get = (key) => r[key] ?? r[key.charAt(0).toUpperCase() + key.slice(1)] ?? "";
            const title = String(get("title")).trim();
            return {
              user_id: session.user.id,
              title,
              category: matchCategoryId(get("category")),
              item_type: String(get("itemType") || get("item_type") || "").trim() || null,
              brand: String(get("brand")).trim() || null,
              condition: String(get("condition")).trim() || null,
              colour: String(get("colour") || get("color")).trim() || null,
              colour_condition: String(get("colourCondition") || get("colour_condition")).trim() || null,
              remark: String(get("remark")).trim() || null,
              cost: get("cost") === "" ? null : Number(get("cost")) || null,
              price: get("price") === "" ? null : Number(get("price")) || null,
              sku: String(get("sku")).trim() || null,
              status: "listed", attributes: {}, custom_attributes: [], photos: [],
              added_on: today, listed_on: today,
            };
          });
          if (newRows.length > 0) {
            const inserted = await apiRequest("/items", { method: "POST", body: JSON.stringify(newRows) }, session, updateSession);
            setItems((prev) => [...(inserted || []).map(rowToItem), ...prev]);
          }
          setCsvImportResult({ added: newRows.length, skipped });
        } catch (e) {
          setCsvImportResult({ added: 0, skipped: 0, error: true });
        } finally {
          setCsvImporting(false);
        }
      },
      error: () => { setCsvImporting(false); setCsvImportResult({ added: 0, skipped: 0, error: true }); },
    });
  }

  if (session === undefined) {
    return (
      <div className="min-h-screen w-full flex items-center justify-center" style={{ background: "#0B0B0C", fontFamily: "'Inter', sans-serif" }}>
        <style>{`@import url('https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,500;9..144,600&family=Space+Grotesk:wght@500;600;700&family=Inter:wght@400;500;600;700;800&display=swap');`}</style>
        <Loader2 size={20} className="animate-spin" style={{ color: "#75726C" }} />
      </div>
    );
  }

  if (!session && !showAuthForm) {
    return (
      <div className="min-h-screen w-full" style={{ background: "#0B0B0C", fontFamily: "'Inter', sans-serif", color: "#F2F0EA" }}>
        <style>{`@import url('https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,500;9..144,600&family=Space+Grotesk:wght@500;600;700&family=Inter:wght@400;500;600;700;800&family=IBM+Plex+Mono:wght@400;500;600&display=swap');`}</style>

        <header style={{ background: "#17171A", borderBottom: "1px solid #2A2A2F" }}>
          <div className="max-w-5xl mx-auto px-6 py-5 flex items-center justify-between">
            <div className="flex items-baseline gap-2">
              <span className="text-xl font-semibold tracking-tight" style={{ fontFamily: "'Fraunces', serif" }}>RAVUNO</span>
              <span className="text-[10px] font-medium tracking-wider" style={{ color: "#75726C" }}>V1.4</span>
            </div>
            <button onClick={() => { setAuthMode("signin"); setShowAuthForm(true); }} className="text-sm font-medium px-4 py-2 rounded-lg border" style={{ borderColor: "#2A2A2F" }}>
              Sign in
            </button>
          </div>
        </header>

        <main className="max-w-5xl mx-auto px-6">
          {/* hero */}
          <section className="py-16 text-center max-w-2xl mx-auto">
            <h1 className="text-3xl sm:text-4xl font-semibold tracking-tight mb-4" style={{ fontFamily: "'Fraunces', serif" }}>One place to list, track, and sell everything.</h1>
            <p className="text-base mb-8" style={{ color: "#8C8A85" }}>
              Ravuno helps solo resellers generate ready-to-paste listings for Carousell, Grailed, Depop, Threads and Shopee,
              manage stock in one ledger, run a simple POS with negotiated discounts, and track sales performance — all in one login.
            </p>
            <button onClick={() => { setAuthMode("signup"); setShowAuthForm(true); }}
              className="inline-flex items-center gap-2 px-6 py-3 rounded-xl text-sm font-semibold"
              style={{ background: "#C9A567", color: "#0B0B0C", fontFamily: "'IBM Plex Mono', monospace" }}>
              Get started
            </button>
          </section>

          {/* features */}
          <section className="grid sm:grid-cols-3 gap-4 pb-16">
            {[
              { id: "crosslist", title: "Cross-List", desc: "One product form, five ready-to-paste platform listings." },
              { id: "inventory", title: "Stock Ledger", desc: "Search, track, and export your inventory for stock takes." },
              { id: "pos", title: "POS + Sales", desc: "Search-to-sell checkout with discounts, invoices, and reports." },
            ].map((f) => (
              <button key={f.id} onClick={() => setLandingPreview(f.id)}
                className="rounded-xl border p-5 text-left transition-colors" style={{ background: "#17171A", borderColor: "#2A2A2F" }}>
                <h3 className="text-sm font-semibold mb-1" style={{ fontFamily: "'Fraunces', serif" }}>{f.title}</h3>
                <p className="text-xs mb-3" style={{ color: "#8C8A85" }}>{f.desc}</p>
                <span className="text-xs font-semibold" style={{ color: "#C9A567" }}>See how it looks →</span>
              </button>
            ))}
          </section>

          {/* free + support */}
          <section className="pb-16">
            <div className="rounded-xl border p-8 text-center max-w-xl mx-auto" style={{ background: "#17171A", borderColor: "#C9A567" }}>
              <span className="text-xs font-semibold px-2 py-0.5 rounded-full" style={{ background: "#241F13", color: "#C9A567" }}>Free, always</span>
              <p className="text-2xl font-bold mt-3 mb-2">RM0<span className="text-sm font-normal" style={{ color: "#75726C" }}> / forever</span></p>
              <p className="text-sm mb-6" style={{ color: "#8C8A85" }}>
                Full access to Cross-List, Stock Ledger, POS, and Sales — no paywall, no trial limits.
                If it helps your business, an optional tip is always appreciated once you're signed in.
              </p>
              <button onClick={() => { setAuthMode("signup"); setShowAuthForm(true); }} className="w-full py-2.5 rounded-lg text-sm font-semibold" style={{ background: "#C9A567", color: "#0B0B0C" }}>
                Get started — it's free
              </button>
            </div>
          </section>
        </main>

        <footer className="border-t" style={{ borderColor: "#2A2A2F" }}>
          <div className="max-w-5xl mx-auto px-6 py-8 text-center">
            <p className="text-xs mb-1" style={{ color: "#75726C" }}>Questions, support, or refund requests:</p>
            <a href="mailto:Unknownlable00@gmail.com" className="text-sm font-medium" style={{ color: "#C9A567" }}>Unknownlable00@gmail.com</a>
          </div>
        </footer>

        {landingPreview && (
          <div className="fixed inset-0 z-50 flex items-center justify-center p-4" style={{ background: "rgba(23,22,26,0.6)" }} onClick={() => setLandingPreview(null)}>
            <div className="w-full max-w-lg rounded-xl border overflow-hidden max-h-[85vh] overflow-y-auto" style={{ background: "#0B0B0C", borderColor: "#2A2A2F" }} onClick={(e) => e.stopPropagation()}>
              <div className="flex items-center justify-between px-5 py-4 border-b" style={{ background: "#17171A", borderColor: "#2A2A2F" }}>
                <span className="text-sm font-semibold" style={{ color: "#F2F0EA" }}>
                  {landingPreview === "crosslist" && "Cross-List preview"}
                  {landingPreview === "inventory" && "Stock Ledger preview"}
                  {landingPreview === "pos" && "POS + Sales preview"}
                </span>
                <button onClick={() => setLandingPreview(null)} className="w-7 h-7 rounded-full flex items-center justify-center" style={{ background: "#0B0B0C", color: "#F2F0EA" }} aria-label="Close">
                  <X size={15} />
                </button>
              </div>

              <div className="p-5">
                {landingPreview === "crosslist" && (
                  <div className="space-y-4">
                    <div className="rounded-lg border p-4" style={{ background: "#17171A", borderColor: "#2A2A2F" }}>
                      <p className="text-xs font-semibold mb-3" style={{ color: "#F2F0EA" }}>Uniqlo Fleece Jacket · Apparel</p>
                      <div className="grid grid-cols-3 gap-2 text-xs" style={{ color: "#8C8A85" }}>
                        <div><span className="block text-[10px] uppercase" style={{ color: "#75726C" }}>Size</span>M</div>
                        <div><span className="block text-[10px] uppercase" style={{ color: "#75726C" }}>Condition</span>9/10</div>
                        <div><span className="block text-[10px] uppercase" style={{ color: "#75726C" }}>Price</span>RM45</div>
                      </div>
                    </div>
                    <div className="rounded-lg border overflow-hidden" style={{ background: "#17171A", borderColor: "#2A2A2F" }}>
                      <div className="px-4 py-2 border-b flex items-center justify-between" style={{ borderColor: "#232327" }}>
                        <span className="text-xs font-semibold px-2 py-0.5 rounded" style={{ background: "#0B0B0C" }}>Carousell</span>
                        <span className="text-[10px] font-medium" style={{ color: "#C9A567" }}>Copy</span>
                      </div>
                      <pre className="text-[11px] px-4 py-3 whitespace-pre-wrap" style={{ fontFamily: "'IBM Plex Mono', monospace", color: "#F2F0EA" }}>
{`Uniqlo Fleece Jacket — Size M

Condition: 9/10
Chest: 21" | Length: 27"
Remark: small pilling near cuff

Price: RM45
SKU: TZ-014

Thanks for checking out my shop!`}
                      </pre>
                    </div>
                    <p className="text-xs" style={{ color: "#75726C" }}>Same product also generates ready-to-paste versions for Threads, Grailed, Depop, and Shopee.</p>
                  </div>
                )}

                {landingPreview === "inventory" && (
                  <div className="space-y-4">
                    <div className="grid grid-cols-3 gap-2">
                      <StatCard icon={Package} label="Listed" value={12} tone="#4A4A52" />
                      <StatCard icon={TrendingUp} label="Sold" value={5} tone="#8C8A85" />
                      <StatCard icon={Clock} label="Stale 21d+" value={2} tone="#C9695E" />
                    </div>
                    <div className="rounded-lg border overflow-hidden" style={{ background: "#17171A", borderColor: "#2A2A2F" }}>
                      {[
                        { sku: "TZ-014", title: "Uniqlo Fleece Jacket", price: 45, status: "LISTED" },
                        { sku: "TZ-009", title: "Nike Windbreaker", price: 65, status: "SOLD" },
                        { sku: "TZ-021", title: "Carhartt Work Pants", price: 38, status: "LISTED" },
                      ].map((row, idx) => (
                        <div key={row.sku} className="flex items-center justify-between px-4 py-3 text-xs" style={{ borderTop: idx > 0 ? "1px solid #232327" : "none" }}>
                          <div>
                            <p className="font-medium" style={{ color: "#F2F0EA" }}>{row.title}</p>
                            <p style={{ color: "#75726C", fontFamily: "'IBM Plex Mono', monospace" }}>{row.sku} · RM{row.price}</p>
                          </div>
                          <span className="text-[10px] font-semibold px-2 py-0.5 rounded-full" style={{ background: row.status === "SOLD" ? "#1C2420" : "#241F13", color: row.status === "SOLD" ? "#8C8A85" : "#F2F0EA" }}>{row.status}</span>
                        </div>
                      ))}
                    </div>
                    <p className="text-xs" style={{ color: "#75726C" }}>Search by SKU, brand, or category, and export a print-ready stock take to Excel anytime.</p>
                  </div>
                )}

                {landingPreview === "pos" && (
                  <div className="space-y-4">
                    <div className="rounded-lg border p-4" style={{ background: "#17171A", borderColor: "#2A2A2F" }}>
                      <p className="text-xs font-semibold mb-3" style={{ color: "#F2F0EA" }}>Sale</p>
                      {[{ name: "Vintage Shirt", price: 80, discount: 10 }, { name: "Casio Watch", price: 150, discount: 20 }].map((l) => (
                        <div key={l.name} className="flex items-center justify-between text-xs py-1.5 border-b" style={{ borderColor: "#232327" }}>
                          <span style={{ color: "#F2F0EA" }}>{l.name}</span>
                          <span style={{ fontFamily: "'IBM Plex Mono', monospace", color: "#C9A567" }}>RM{l.price - l.discount}</span>
                        </div>
                      ))}
                      <div className="flex items-center justify-between text-sm font-semibold pt-2" style={{ color: "#F2F0EA" }}>
                        <span>Total</span><span style={{ fontFamily: "'IBM Plex Mono', monospace" }}>RM200</span>
                      </div>
                    </div>
                    <div className="grid grid-cols-2 sm:grid-cols-4 gap-2">
                      <StatCard icon={TrendingUp} label="Sales (RM)" value={4850} tone="#C9A567" />
                      <StatCard icon={Receipt} label="Orders" value={37} tone="#4A4A52" />
                      <StatCard icon={ShoppingBag} label="Units" value={42} tone="#8C8A85" />
                      <StatCard icon={Percent} label="Discounts" value={320} tone="#C9695E" />
                    </div>
                    <p className="text-xs" style={{ color: "#75726C" }}>Search a product, negotiate a discount, complete the sale — invoice and stock update automatically.</p>
                  </div>
                )}
              </div>
            </div>
          </div>
        )}
      </div>
    );
  }

  if (!session) {
    return (
      <div className="min-h-screen w-full flex items-center justify-center px-6" style={{ background: "#0B0B0C", fontFamily: "'Inter', sans-serif" }}>
        <style>{`@import url('https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,500;9..144,600&family=Space+Grotesk:wght@500;600;700&family=Inter:wght@400;500;600;700;800&display=swap');`}</style>
        <div className="w-full max-w-sm">
          <button onClick={() => setShowAuthForm(false)} className="text-xs font-medium mb-6" style={{ color: "#75726C" }}>← Back</button>
          <div className="flex items-baseline gap-2 justify-center mb-8">
            <span className="text-2xl font-semibold tracking-tight" style={{ color: "#F2F0EA", fontFamily: "'Fraunces', serif" }}>RAVUNO</span>
            <span className="text-[10px] font-medium tracking-wider" style={{ color: "#75726C" }}>V1.4</span>
          </div>
          <div className="rounded-xl border p-6" style={{ background: "#17171A", borderColor: "#2A2A2F" }}>
            <h1 className="text-lg font-semibold mb-1" style={{ color: "#F2F0EA", fontFamily: "'Fraunces', serif" }}>
              {authMode === "signin" ? "Sign in" : "Create your account"}
            </h1>
            <p className="text-xs mb-5" style={{ color: "#75726C" }}>
              {authMode === "signin" ? "Access your inventory from any device." : "Your data is private to your account only."}
            </p>
            {authNotice && <p className="text-xs px-3 py-2 rounded-lg mb-4" style={{ background: "#241F13", color: "#C9A567" }}>{authNotice}</p>}
            {authError && <p className="text-xs px-3 py-2 rounded-lg mb-4" style={{ background: "#2A1715", color: "#C9695E" }}>{authError}</p>}
            <form onSubmit={handleAuthSubmit} className="space-y-3">
              <Field label="Email" type="email" value={authEmail} onChange={(e) => setAuthEmail(e.target.value)} placeholder="you@example.com" />
              <Field label="Password" type="password" value={authPassword} onChange={(e) => setAuthPassword(e.target.value)} placeholder="••••••••" />
              <button type="submit" disabled={authBusy}
                className="w-full flex items-center justify-center gap-2 py-2.5 rounded-xl text-sm font-semibold transition-opacity"
                style={{ background: "#C9A567", color: "#0B0B0C", opacity: authBusy ? 0.6 : 1, fontFamily: "'IBM Plex Mono', monospace" }}>
                {authBusy ? <Loader2 size={14} className="animate-spin" /> : null}
                {authMode === "signin" ? "Sign in" : "Sign up"}
              </button>
            </form>
            <button
              onClick={() => { setAuthMode(authMode === "signin" ? "signup" : "signin"); setAuthError(""); setAuthNotice(""); }}
              className="w-full text-center text-xs mt-4" style={{ color: "#8C8A85" }}>
              {authMode === "signin" ? "New here? Create an account" : "Already have an account? Sign in"}
            </button>
          </div>
        </div>
      </div>
    );
  }

  return (
    <div className="min-h-screen w-full" style={{ background: "#0B0B0C", color: "#F2F0EA", fontFamily: "'Inter', sans-serif" }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,500;9..144,600&family=Space+Grotesk:wght@500;600;700&family=Inter:wght@400;500;600;700;800&family=IBM+Plex+Mono:wght@400;500;600&display=swap');
        .no-scrollbar { scrollbar-width: none; -ms-overflow-style: none; }
        .no-scrollbar::-webkit-scrollbar { display: none; }
      `}</style>

      <header style={{ background: "#17171A", borderBottom: "1px solid #2A2A2F" }}>
        <div className="max-w-6xl mx-auto px-6 py-6 relative flex flex-col items-center gap-5">
          <div className="absolute right-6 top-6 flex items-center gap-3">
            {profile?.subscription_status === "active" && (
              <span className="text-xs font-semibold px-2 py-1 rounded-full" style={{ background: "#241F13", color: "#C9A567" }}>❤️ Supporter</span>
            )}
            <button onClick={handleSignOut} className="text-xs font-medium" style={{ color: "#75726C" }}>
              Sign out
            </button>
          </div>
          <div className="flex items-baseline gap-2">
            <span className="text-2xl font-semibold tracking-tight" style={{ color: "#F2F0EA", fontFamily: "'Fraunces', serif" }}>RAVUNO</span>
            <span className="text-[10px] font-medium tracking-wider" style={{ color: "#75726C", fontFamily: "'IBM Plex Mono', monospace" }}>V1.4</span>
          </div>
          <nav className="flex gap-1 p-1 rounded-lg overflow-x-auto no-scrollbar max-w-full" style={{ background: "#0B0B0C", border: "1px solid #2A2A2F", WebkitOverflowScrolling: "touch", scrollSnapType: "x proximity" }}>
            {[{ id: "crosslist", label: "Cross-List" }, { id: "inventory", label: "Stock Ledger" }, { id: "pos", label: "POS" }, { id: "sales", label: "Sales" }, { id: "templates", label: "Templates" }, { id: "support", label: "Support" }].map((t) => (
              <button key={t.id} onClick={() => setTab(t.id)}
                className="px-4 py-1.5 rounded-md text-sm font-medium transition-colors shrink-0 whitespace-nowrap"
                style={{
                  background: tab === t.id ? "#1E1E22" : "transparent",
                  color: tab === t.id ? "#C9A567" : "#8C8A85",
                  boxShadow: tab === t.id ? "inset 0 0 0 1px #2A2A2F" : "none",
                  scrollSnapAlign: "start",
                  fontFamily: "'Space Grotesk', sans-serif",
                }}>
                {t.label}
              </button>
            ))}
          </nav>
        </div>
      </header>

      {loading ? (
        <div className="max-w-6xl mx-auto px-6 py-20 flex items-center justify-center gap-2" style={{ color: "#8C8A85" }}>
          <Loader2 size={16} className="animate-spin" /> Loading your saved items…
        </div>
      ) : (
        <main className="max-w-6xl mx-auto px-6 py-8">
          {saveError && (
            <div className="mb-4 text-xs px-3 py-2 rounded-lg" style={{ background: "#2A1715", color: "#C9695E" }}>
              Couldn't save changes — they may not persist. Try again shortly.
            </div>
          )}

          {tab === "crosslist" && (
            <div className="grid lg:grid-cols-[380px_1fr] gap-8">
              <section>
                <div className="rounded-xl border p-6" style={{ background: "#17171A", borderColor: "#2A2A2F" }}>
                  <div className="flex items-center justify-between mb-4">
                    <h2 className="text-lg font-semibold" style={{ color: "#F2F0EA", fontFamily: "'Fraunces', serif" }}>
                      {editingId ? "Edit product" : "Add a product"}
                    </h2>
                    {editingId && (
                      <button onClick={cancelEdit} className="text-xs font-medium" style={{ color: "#75726C" }}>Cancel</button>
                    )}
                  </div>
                  <div className="space-y-3">
                    <Select label="Category" value={form.category} onChange={handleCategoryChange}>
                      <option value="">Select category…</option>
                      {CATEGORY_LIST.map((c) => <option key={c.id} value={c.id}>{c.label}</option>)}
                    </Select>

                    {form.category === "apparel" && <ApparelFields form={form} setForm={setForm} />}

                    {form.category && form.category !== "apparel" && form.category !== "other" && (
                      <>
                        <Field label="Item type" value={form.itemType} onChange={set("itemType")} placeholder="e.g. Sneakers, Digital Watch" />
                        <DynamicAttributeFields
                          defs={CATEGORY_DEFS[form.category]?.attributes || []}
                          values={form.attributes}
                          onChange={(id, val) => setForm((f) => ({ ...f, attributes: { ...f.attributes, [id]: val } }))}
                        />
                      </>
                    )}

                    {form.category === "other" && (
                      <>
                        <Field label="Item type" value={form.itemType} onChange={set("itemType")} placeholder="e.g. Vinyl record" />
                        <CustomAttributesEditor value={form.customAttributes} onChange={(next) => setForm((f) => ({ ...f, customAttributes: next }))} />
                      </>
                    )}

                    {!form.category && (
                      <p className="text-xs px-3 py-2 rounded-lg" style={{ background: "#241F13", color: "#8C8A85" }}>
                        Select a category to see relevant fields.
                      </p>
                    )}

                    {form.category && (
                      <>
                        <div className="border-t pt-3 mt-1" style={{ borderColor: "#232327" }}>
                          <Field label="Product title" value={form.title} onChange={(e) => setForm((f) => ({ ...f, title: e.target.value, titleTouched: true }))} placeholder="Auto from brand + item type" />
                        </div>
                        <div className="grid grid-cols-2 gap-3">
                          <Field label="Brand" value={form.brand} onChange={set("brand")} placeholder="Uniqlo" />
                          <Field label="Condition (/10)" value={form.condition} onChange={set("condition")} placeholder="9" type="number" />
                        </div>
                        <div className="grid grid-cols-2 gap-3">
                          <Field label="Colour" value={form.colour} onChange={set("colour")} placeholder="Black" />
                          <Field label="Colour condition (/10)" value={form.colourCondition} onChange={set("colourCondition")} placeholder="optional" type="number" />
                        </div>
                        <Field label="Remark" value={form.remark} onChange={set("remark")} placeholder="small pilling near cuff" />
                        <div className="grid grid-cols-3 gap-3">
                          <Field label="Cost (RM)" value={form.cost} onChange={set("cost")} placeholder="15" type="number" />
                          <Field label="Price (RM)" value={form.price} onChange={set("price")} placeholder="45" type="number" />
                          <Field label="SKU" value={form.sku} onChange={set("sku")} placeholder="TZ-014" />
                        </div>
                        <PhotoGalleryPicker
                          photos={form.photos}
                          onChange={(next) => setForm((f) => ({ ...f, photos: next }))}
                          onPreview={(idx) => setLightbox({ photos: form.photos, index: idx })}
                        />
                      </>
                    )}
                  </div>
                  <button onClick={saveListing} disabled={!isFormUsable}
                    className="mt-5 w-full flex items-center justify-center gap-2 py-2.5 rounded-xl text-sm font-semibold transition-opacity"
                    style={{ background: "#C9A567", color: "#0B0B0C", opacity: isFormUsable ? 1 : 0.35, fontFamily: "'IBM Plex Mono', monospace" }}>
                    <Plus size={15} /> {editingId ? "Update listing" : "Save listing & generate"}
                  </button>
                </div>

                {items.length > 0 && (
                  <div className="mt-6">
                    <div className="flex items-center justify-between mb-2">
                      <p className="text-xs tracking-[0.2em] uppercase" style={{ color: "#8C8A85", fontFamily: "'IBM Plex Mono', monospace" }}>Saved ({items.length})</p>
                    </div>
                    <div className="relative mb-2">
                      <Search size={13} className="absolute left-3 top-1/2 -translate-y-1/2" style={{ color: "#75726C" }} />
                      <input
                        value={crosslistSearch}
                        onChange={(e) => setCrosslistSearch(e.target.value)}
                        placeholder="Search saved items…"
                        className="w-full pl-8 pr-3 py-2 rounded-lg border text-xs outline-none"
                        style={{ borderColor: "#2A2A2F", background: "#0B0B0C", color: "#F2F0EA" }}
                      />
                    </div>
                    <div className="space-y-2 max-h-80 overflow-y-auto pr-1">
                      {crosslistItems.map((l) => (
                        <button key={l.id} onClick={() => setActiveId(l.id)}
                          className="w-full flex items-center justify-between px-3 py-2 rounded-xl border text-left transition-colors"
                          style={{ borderColor: activeId === l.id ? "#C9A567" : "#2A2A2F", background: activeId === l.id ? "#1E1E22" : "transparent" }}>
                          <span className="flex items-center gap-2 min-w-0">
                            {l.photos && l.photos[0] ? (
                              <img src={l.photos[0]} alt="" className="w-8 h-8 rounded object-cover shrink-0" />
                            ) : (
                              <ChevronRight size={14} className="shrink-0" style={{ color: "#8C8A85" }} />
                            )}
                            <span className="truncate text-sm font-medium">{productTitle(l)}</span>
                          </span>
                          <span className="flex items-center gap-2 shrink-0">
                            <span className="text-xs" style={{ color: "#8C8A85", fontFamily: "'IBM Plex Mono', monospace" }}>{CATEGORY_LABELS[l.category]}</span>
                            <Trash2 size={14} style={{ color: "#C9695E" }} onClick={(e) => { e.stopPropagation(); removeItem(l.id); }} />
                          </span>
                        </button>
                      ))}
                      {crosslistItems.length === 0 && (
                        <p className="text-xs text-center py-4" style={{ color: "#75726C" }}>No items match your search.</p>
                      )}
                    </div>
                  </div>
                )}
              </section>

              <section>
                <p className="text-xs tracking-wide uppercase mb-3 font-medium" style={{ color: "#75726C" }}>
                  {activeListing ? `Ready to paste — ${productTitle(activeListing)}` : "Live preview — fill in the product on the left"}
                </p>
                <div className="grid sm:grid-cols-2 gap-4">
                  {PLATFORMS.map((p) => (
                    <div key={p.id} className="rounded-xl border overflow-hidden" style={{ background: "#17171A", borderColor: "#2A2A2F" }}>
                      <div className="flex items-center justify-between px-4 py-3 border-b" style={{ borderColor: "#232327" }}>
                        <span className="text-xs font-semibold px-2 py-1 rounded-md" style={{ background: "#0B0B0C", color: "#F2F0EA" }}>{p.label}</span>
                        <CopyButton text={activeTemplates[p.id]} />
                      </div>
                      <pre className="text-xs px-4 py-3 whitespace-pre-wrap leading-relaxed" style={{ fontFamily: "'IBM Plex Mono', monospace", color: "#F2F0EA" }}>
                        {activeTemplates[p.id]}
                      </pre>
                    </div>
                  ))}
                </div>
              </section>
            </div>
          )}

          {tab === "inventory" && (
            <div className="space-y-8">
              <div className="grid grid-cols-3 gap-3">
                <StatCard icon={Package} label="Listed" value={stats.listedCount} tone="#4A4A52" />
                <StatCard icon={TrendingUp} label="Sold" value={stats.soldCount} tone="#8C8A85" />
                <StatCard icon={Clock} label="Stale 21d+" value={stats.staleCount} tone="#C9695E" />
              </div>

              <div>
                <div className="flex items-center justify-between mb-3 flex-wrap gap-2">
                  <h2 className="text-base font-semibold" style={{ color: "#F2F0EA", fontFamily: "'Fraunces', serif" }}>Stock Items</h2>
                  <div className="flex items-center gap-2 flex-wrap">
                    <button onClick={() => exportStockTake("listed")}
                      className="flex items-center gap-1.5 text-xs font-semibold px-3 py-2 rounded-lg border"
                      style={{ borderColor: "#2A2A2F", color: "#F2F0EA", fontFamily: "'IBM Plex Mono', monospace" }}>
                      <Download size={13} /> Export Listed Only
                    </button>
                    <button onClick={() => exportStockTake("all")}
                      className="flex items-center gap-1.5 text-xs font-semibold px-3 py-2 rounded-lg"
                      style={{ background: "#C9A567", color: "#0B0B0C", fontFamily: "'IBM Plex Mono', monospace" }}>
                      <Download size={13} /> Export All Stock
                    </button>
                  </div>
                </div>

                <div className="rounded-lg border p-3 mb-4" style={{ borderColor: "#2A2A2F", background: "#17171A" }}>
                  <div className="flex items-center justify-between flex-wrap gap-2">
                    <div>
                      <p className="text-xs font-semibold" style={{ color: "#F2F0EA" }}>Bulk import from CSV</p>
                      <p className="text-[11px]" style={{ color: "#75726C" }}>Text fields only — add photos afterward via Edit.</p>
                    </div>
                    <div className="flex items-center gap-2">
                      <button onClick={downloadSampleImportCsv} className="text-xs font-medium" style={{ color: "#C9A567" }}>
                        Download sample
                      </button>
                      <button onClick={() => csvInputRef.current?.click()} disabled={csvImporting}
                        className="text-xs font-semibold px-3 py-1.5 rounded-lg border" style={{ borderColor: "#2A2A2F", color: "#F2F0EA", opacity: csvImporting ? 0.6 : 1 }}>
                        {csvImporting ? "Importing…" : "Import CSV"}
                      </button>
                    </div>
                  </div>
                  {csvImportResult && (
                    <p className="text-[11px] mt-2" style={{ color: csvImportResult.error ? "#C9695E" : "#C9A567" }}>
                      {csvImportResult.error ? "Import failed — check the file and try again."
                        : `Imported ${csvImportResult.added} item${csvImportResult.added === 1 ? "" : "s"}${csvImportResult.skipped ? `, skipped ${csvImportResult.skipped} row(s) missing a title` : ""}.`}
                    </p>
                  )}
                  <input ref={csvInputRef} type="file" accept=".csv" className="hidden"
                    onChange={(e) => { const f = e.target.files?.[0]; if (f) handleCsvImport(f); e.target.value = ""; }} />
                </div>

                <div className="flex gap-1 p-1 rounded-lg mb-3 w-fit" style={{ background: "#0B0B0C", border: "1px solid #2A2A2F" }}>
                  {[{ id: "all", label: "All" }, { id: "listed", label: "Listed" }, { id: "sold", label: "Sold" }].map((f) => (
                    <button key={f.id} onClick={() => setLedgerStatusFilter(f.id)}
                      className="px-3 py-1.5 rounded-md text-xs font-medium transition-colors"
                      style={{ background: ledgerStatusFilter === f.id ? "#1E1E22" : "transparent", color: ledgerStatusFilter === f.id ? "#C9A567" : "#75726C" }}>
                      {f.label}
                    </button>
                  ))}
                </div>

                <div className="relative mb-4">
                  <Search size={15} className="absolute left-3 top-1/2 -translate-y-1/2" style={{ color: "#75726C" }} />
                  <input value={ledgerSearch} onChange={(e) => setLedgerSearch(e.target.value)} placeholder="Search SKU, brand, item…"
                    className="w-full pl-9 pr-9 py-2.5 rounded-lg border text-sm outline-none" style={{ borderColor: "#2A2A2F", background: "#17171A", color: "#F2F0EA" }} />
                  {ledgerSearch && (
                    <button onClick={() => setLedgerSearch("")} className="absolute right-3 top-1/2 -translate-y-1/2" aria-label="Clear search">
                      <X size={15} style={{ color: "#75726C" }} />
                    </button>
                  )}
                </div>

                <div className="hidden sm:grid gap-2 px-4 mb-1 text-[11px] uppercase tracking-wide"
                  style={{ color: "#8C8A85", fontFamily: "'IBM Plex Mono', monospace", gridTemplateColumns: "1fr 1.6fr 1.1fr 1.1fr 0.7fr 0.7fr 0.8fr" }}>
                  <span>SKU</span><span>Title</span><span>Category</span><span>Item Type</span><span>Cost</span><span>Price</span><span>Status</span>
                </div>

                <div className="space-y-2">
                  {ledgerItems.map((i) => {
                    const stale = i.status === "listed" && daysSince(i.addedOn) >= 21;
                    return (
                      <div key={i.id} onClick={() => { setDetailId(i.id); setGalleryIndex(0); }}
                        className="rounded-xl border px-4 py-3 cursor-pointer transition-colors"
                        style={{ borderColor: i.status === "sold" ? "#8C8A85" : stale ? "#C9695E" : "#2A2A2F", background: "#17171A" }}>

                        <div className="hidden sm:grid items-center gap-2 text-sm" style={{ gridTemplateColumns: "1fr 1.6fr 1.1fr 1.1fr 0.7fr 0.7fr 0.8fr" }}>
                          <span className="truncate" style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{i.sku || "-"}</span>
                          <span className="truncate font-medium">{productTitle(i)}</span>
                          <span className="truncate">{CATEGORY_LABELS[i.category] || "-"}</span>
                          <span className="truncate">{i.itemType || "-"}</span>
                          <span style={{ fontFamily: "'IBM Plex Mono', monospace" }}>RM{i.cost || 0}</span>
                          <span style={{ fontFamily: "'IBM Plex Mono', monospace" }}>RM{i.price || 0}</span>
                          <span className="text-xs font-semibold px-2 py-0.5 rounded-full w-fit"
                            style={{ background: i.status === "sold" ? "#1C2420" : "#241F13", color: i.status === "sold" ? "#8C8A85" : "#F2F0EA", fontFamily: "'IBM Plex Mono', monospace" }}>
                            {(i.status || "listed").toUpperCase()}
                          </span>
                        </div>

                        <div className="sm:hidden flex items-center gap-3">
                          {i.photos && i.photos[0] ? (
                            <img src={i.photos[0]} alt="" className="w-10 h-10 rounded-lg object-cover shrink-0" />
                          ) : (
                            <div className="w-10 h-10 rounded-lg shrink-0 flex items-center justify-center border-2 border-dashed" style={{ borderColor: "#2A2A2F", color: "#2A2A2F" }}>
                              <ImageIcon size={14} />
                            </div>
                          )}
                          <div className="min-w-0 flex-1">
                            <p className="text-sm font-semibold truncate">{productTitle(i)}</p>
                            <p className="text-xs" style={{ color: "#8C8A85", fontFamily: "'IBM Plex Mono', monospace" }}>
                              {i.sku || "-"} · {CATEGORY_LABELS[i.category] || "-"} · RM{i.price || 0}
                            </p>
                          </div>
                          <span className="text-xs font-semibold px-2 py-0.5 rounded-full shrink-0"
                            style={{ background: i.status === "sold" ? "#1C2420" : "#241F13", color: i.status === "sold" ? "#8C8A85" : "#F2F0EA", fontFamily: "'IBM Plex Mono', monospace" }}>
                            {(i.status || "listed").toUpperCase()}
                          </span>
                        </div>

                        <div className="flex items-center gap-2 mt-2 pt-2 border-t" style={{ borderColor: "#232327" }} onClick={(e) => e.stopPropagation()}>
                          <Trash2 size={14} style={{ color: "#C9695E", cursor: "pointer" }} onClick={() => removeItem(i.id)} />
                        </div>
                      </div>
                    );
                  })}
                  {ledgerItems.length === 0 && (
                    <p className="text-sm text-center py-8" style={{ color: "#75726C" }}>
                      {items.length === 0 ? "No items yet — add one from the Cross-List tab." : "No items match your search."}
                    </p>
                  )}
                </div>
              </div>
            </div>
          )}

          {tab === "pos" && (
            <div className="grid lg:grid-cols-[1fr_380px] gap-8">
              <section>
                <div className="rounded-xl border p-6" style={{ background: "#17171A", borderColor: "#2A2A2F" }}>
                  <div className="relative">
                    <Search size={17} className="absolute left-4 top-1/2 -translate-y-1/2" style={{ color: "#75726C" }} />
                    <input
                      value={skuSearch}
                      onChange={(e) => setSkuSearch(e.target.value)}
                      placeholder="Search product name or SKU…"
                      autoFocus
                      className="w-full pl-11 pr-4 py-3 rounded-lg border text-base outline-none"
                      style={{ borderColor: "#2A2A2F", background: "#17171A", color: "#F2F0EA" }}
                    />
                  </div>
                </div>

                {skuSearch.trim() === "" ? (
                  <p className="text-sm text-center py-10" style={{ color: "#75726C" }}>
                    Search by product name or SKU to find an item to sell.
                  </p>
                ) : (
                  <div className="mt-4 space-y-2">
                    {searchResults.map((i) => (
                      <button key={i.id} onClick={() => addToCart(i)}
                        className="w-full flex items-center justify-between gap-3 px-4 py-3 rounded-xl border text-left" style={{ borderColor: "#2A2A2F", background: "#17171A" }}>
                        <div className="flex items-center gap-3 min-w-0">
                          {i.photos && i.photos[0] ? (
                            <img src={i.photos[0]} alt="" className="w-10 h-10 rounded-lg object-cover shrink-0" />
                          ) : (
                            <div className="w-10 h-10 rounded-lg shrink-0 flex items-center justify-center border-2 border-dashed" style={{ borderColor: "#2A2A2F", color: "#2A2A2F" }}>
                              <ImageIcon size={14} />
                            </div>
                          )}
                          <div className="min-w-0">
                            <p className="text-sm font-semibold truncate">{productTitle(i)}</p>
                            <p className="text-xs" style={{ color: "#8C8A85", fontFamily: "'IBM Plex Mono', monospace" }}>{i.sku || "no sku"} · RM{i.price || 0}</p>
                          </div>
                        </div>
                        <Plus size={16} style={{ color: "#F2F0EA" }} className="shrink-0" />
                      </button>
                    ))}
                    {searchResults.length === 0 && <p className="text-sm text-center py-8" style={{ color: "#75726C" }}>No products found.</p>}
                  </div>
                )}
              </section>

              <section>
                <div className="rounded-xl border p-5 sticky top-4" style={{ background: "#17171A", borderColor: "#2A2A2F" }}>
                  <h2 className="text-base font-semibold mb-3" style={{ color: "#F2F0EA", fontFamily: "'Fraunces', serif" }}>Sale</h2>
                  <div className="space-y-3 mb-4">
                    {cartLines.map((l) => (
                      <div key={l.id} className="pb-3 border-b" style={{ borderColor: "#232327" }}>
                        <div className="flex items-center justify-between gap-2 mb-2">
                          <p className="truncate text-sm font-medium">{productTitle(l)}</p>
                          <X size={14} style={{ color: "#C9695E", cursor: "pointer" }} className="shrink-0" onClick={() => removeFromCart(l.id)} />
                        </div>
                        <div className="grid grid-cols-2 gap-2 items-end">
                          <label className="block">
                            <span className="block text-[10px] mb-1 uppercase tracking-wide" style={{ color: "#75726C" }}>Discount (RM)</span>
                            <input type="number" min="0" value={l.discount || ""} onChange={(e) => updateDiscount(l.id, e.target.value)} placeholder="0"
                              className="w-full px-2 py-1.5 rounded-md border text-xs text-center outline-none" style={{ borderColor: "#2A2A2F", background: "#17171A", color: "#F2F0EA", fontFamily: "'IBM Plex Mono', monospace" }} />
                          </label>
                          <div>
                            <span className="block text-[10px] mb-1 uppercase tracking-wide" style={{ color: "#75726C" }}>Final</span>
                            <p className="text-xs font-semibold text-center py-1.5" style={{ fontFamily: "'IBM Plex Mono', monospace", color: "#C9A567" }}>RM{l.finalPrice}</p>
                          </div>
                        </div>
                        {l.discount > 0 && (
                          <p className="text-[11px] mt-1 text-right" style={{ color: "#75726C", fontFamily: "'IBM Plex Mono', monospace" }}>
                            RM{l.originalPrice} − RM{l.discount}
                          </p>
                        )}
                      </div>
                    ))}
                    {cartLines.length === 0 && <p className="text-xs" style={{ color: "#75726C" }}>No items yet — search and select a product to begin a sale.</p>}
                  </div>

                  <div className="space-y-1 pt-1 pb-3 border-b mb-3 text-sm" style={{ borderColor: "#2A2A2F" }}>
                    <div className="flex items-center justify-between" style={{ color: "#8C8A85" }}>
                      <span>Subtotal</span><span style={{ fontFamily: "'IBM Plex Mono', monospace" }}>RM{cartSubtotal}</span>
                    </div>
                    <div className="flex items-center justify-between" style={{ color: "#8C8A85" }}>
                      <span>Discount</span><span style={{ fontFamily: "'IBM Plex Mono', monospace" }}>− RM{cartDiscount}</span>
                    </div>
                    <div className="flex items-center justify-between font-semibold pt-1" style={{ color: "#F2F0EA" }}>
                      <span>Total</span><span style={{ fontFamily: "'IBM Plex Mono', monospace" }}>RM{cartTotal}</span>
                    </div>
                  </div>

                  <Select label="Payment method" value={payMethod} onChange={(e) => setPayMethod(e.target.value)}>
                    <option>Cash</option><option>Bank Transfer</option><option>Online</option>
                  </Select>
                  <button onClick={completeSale} disabled={cartLines.length === 0}
                    className="mt-4 w-full flex items-center justify-center gap-2 py-2.5 rounded-xl text-sm font-semibold transition-opacity"
                    style={{ background: "#C9A567", color: "#0B0B0C", opacity: cartLines.length === 0 ? 0.35 : 1, fontFamily: "'IBM Plex Mono', monospace" }}>
                    Complete Sale
                  </button>
                </div>
              </section>
            </div>
          )}

          {tab === "sales" && (
            <div className="space-y-6">
              {invoicedTxns.length === 0 ? (
                <div className="rounded-xl border p-10 text-center" style={{ background: "#17171A", borderColor: "#2A2A2F" }}>
                  <Receipt size={28} className="mx-auto mb-3" style={{ color: "#75726C" }} />
                  <p className="text-sm font-semibold mb-1" style={{ color: "#F2F0EA" }}>No sales data yet</p>
                  <p className="text-xs" style={{ color: "#75726C" }}>Complete your first sale to start tracking performance.</p>
                </div>
              ) : (
                <>
                  {/* period selector */}
                  <div className="flex items-center gap-2 flex-wrap">
                    <div className="flex gap-1 p-1 rounded-lg" style={{ background: "#17171A", border: "1px solid #2A2A2F" }}>
                      {[{ id: "monthly", label: "Monthly" }, { id: "yearly", label: "Yearly" }].map((p) => (
                        <button key={p.id} onClick={() => setSalesPeriodType(p.id)}
                          className="px-3 py-1.5 rounded-md text-xs font-medium transition-colors"
                          style={{ background: salesPeriodType === p.id ? "#C9A567" : "transparent", color: salesPeriodType === p.id ? "#0B0B0C" : "#8C8A85" }}>
                          {p.label}
                        </button>
                      ))}
                    </div>
                    {salesPeriodType === "monthly" && (
                      <select value={salesMonth} onChange={(e) => setSalesMonth(Number(e.target.value))}
                        className="px-3 py-2 rounded-lg border text-sm outline-none" style={{ borderColor: "#2A2A2F", background: "#17171A", color: "#F2F0EA" }}>
                        {MONTH_NAMES.map((m, idx) => <option key={m} value={idx}>{m}</option>)}
                      </select>
                    )}
                    <select value={salesYear} onChange={(e) => setSalesYear(Number(e.target.value))}
                      className="px-3 py-2 rounded-lg border text-sm outline-none" style={{ borderColor: "#2A2A2F", background: "#17171A", color: "#F2F0EA" }}>
                      {availableSalesYears.map((y) => <option key={y} value={y}>{y}</option>)}
                    </select>
                  </div>

                  {/* summary metrics */}
                  <div className="grid grid-cols-2 sm:grid-cols-5 gap-3">
                    <StatCard icon={TrendingUp} label="Total Sales (RM)" value={salesSummary.totalSales} tone="#C9A567" />
                    <StatCard icon={DollarSign} label="Net Profit (RM)" value={salesSummary.netProfit} tone="#4A4A52" />
                    <StatCard icon={Receipt} label="Orders" value={salesSummary.orders} tone="#4A4A52" />
                    <StatCard icon={ShoppingBag} label="Units Sold" value={salesSummary.unitsSold} tone="#8C8A85" />
                    <StatCard icon={Percent} label="Discounts (RM)" value={salesSummary.discounts} tone="#C9695E" />
                  </div>

                  {/* revenue graph */}
                  <div className="rounded-xl border p-5" style={{ background: "#17171A", borderColor: "#2A2A2F" }}>
                    <h2 className="text-sm font-semibold mb-4" style={{ color: "#F2F0EA", fontFamily: "'Fraunces', serif" }}>
                      Revenue — {salesPeriodType === "monthly" ? `${MONTH_NAMES[salesMonth]} ${salesYear}` : salesYear}
                    </h2>
                    <div style={{ width: "100%", height: 220 }}>
                      <ResponsiveContainer>
                        <LineChart data={salesGraphData} margin={{ top: 5, right: 10, bottom: 0, left: 0 }}>
                          <CartesianGrid stroke="#232327" vertical={false} />
                          <XAxis dataKey="label" tick={{ fontSize: 11, fill: "#75726C" }} axisLine={{ stroke: "#2A2A2F" }} tickLine={false} interval={salesPeriodType === "monthly" ? 2 : 0} />
                          <YAxis tick={{ fontSize: 11, fill: "#75726C" }} axisLine={false} tickLine={false} width={44} />
                          <Tooltip
                            contentStyle={{ background: "#1E1E22", border: "1px solid #2A2A2F", borderRadius: 8, fontSize: 12 }}
                            labelStyle={{ color: "#FFFFFF" }}
                            itemStyle={{ color: "#FFFFFF" }}
                            formatter={(value) => [`RM${value}`, "Revenue"]}
                          />
                          <Line type="monotone" dataKey="revenue" stroke="#C9A567" strokeWidth={2} dot={false} activeDot={{ r: 4 }} />
                        </LineChart>
                      </ResponsiveContainer>
                    </div>
                  </div>

                  {/* invoice list */}
                  <div>
                    <h2 className="text-sm font-semibold mb-3" style={{ color: "#F2F0EA", fontFamily: "'Fraunces', serif" }}>Invoices</h2>
                    <div className="space-y-2">
                      {periodTxns.map((t) => (
                        <button key={t.id} onClick={() => setSelectedInvoiceId(t.id)}
                          className="w-full flex items-center justify-between gap-3 px-4 py-3 rounded-xl border text-left" style={{ borderColor: "#2A2A2F", background: "#17171A" }}>
                          <div className="min-w-0">
                            <div className="flex items-center gap-2 mb-0.5">
                              <span className="text-xs font-semibold" style={{ fontFamily: "'IBM Plex Mono', monospace", color: "#F2F0EA" }}>{t.invoiceNumber}</span>
                              <span className="text-xs" style={{ color: "#75726C" }}>
                                {new Date(t.timestamp).toLocaleDateString(undefined, { day: "2-digit", month: "short", year: "numeric" })} · {new Date(t.timestamp).toLocaleTimeString(undefined, { hour: "2-digit", minute: "2-digit" })}
                              </span>
                            </div>
                            <p className="text-sm truncate" style={{ color: "#8C8A85" }}>{t.lines.map((l) => l.title).join(", ")}</p>
                          </div>
                          <div className="text-right shrink-0" style={{ fontFamily: "'IBM Plex Mono', monospace" }}>
                            <p className="text-sm font-semibold" style={{ color: "#C9A567" }}>RM{t.total}</p>
                            {t.discount > 0 && <p className="text-[11px]" style={{ color: "#75726C" }}>−RM{t.discount} disc.</p>}
                          </div>
                        </button>
                      ))}
                      {periodTxns.length === 0 && (
                        <p className="text-sm text-center py-8" style={{ color: "#75726C" }}>No invoices in this period.</p>
                      )}
                    </div>
                  </div>
                </>
              )}
            </div>
          )}

          {tab === "templates" && (
            <TemplatesTab customTemplates={customTemplates} onSave={saveTemplates} saving={templatesSaving}
              logoUrl={logoUrl} onLogoUpload={handleLogoUpload} onLogoRemove={handleLogoRemove} logoUploading={logoUploading} items={items} />
          )}

          {tab === "support" && (
            <div className="max-w-xl mx-auto">
              <div className="rounded-xl border p-8 text-center" style={{ background: "#17171A", borderColor: "#2A2A2F" }}>
                <p className="text-3xl mb-4">👋</p>
                <h2 className="text-lg font-semibold mb-3" style={{ color: "#F2F0EA", fontFamily: "'Fraunces', serif" }}>Why Ravuno is free</h2>
                <p className="text-sm leading-relaxed mb-2" style={{ color: "#8C8A85" }}>
                  I built Ravuno because I know what it's like juggling Carousell, Shopee, and a messy spreadsheet just to
                  keep a small resale business running. I wanted a tool that actually helps sellers like us — without a
                  paywall in the way.
                </p>
                <p className="text-sm leading-relaxed mb-6" style={{ color: "#8C8A85" }}>
                  Ravuno is free to use, fully and always. If it's genuinely helped your business and you'd like to support
                  the time that goes into building and maintaining it, that's completely optional — and appreciated.
                </p>

                <div className="text-left mb-6 pb-6 border-b" style={{ borderColor: "#232327" }}>
                  <p className="text-sm font-semibold mb-1" style={{ color: "#F2F0EA" }}>Got feedback or a feature idea?</p>
                  <p className="text-xs mb-3" style={{ color: "#75726C" }}>This goes straight to me — I read every one.</p>
                  <textarea
                    value={feedbackText}
                    onChange={(e) => { setFeedbackText(e.target.value); if (feedbackStatus !== "idle") setFeedbackStatus("idle"); }}
                    placeholder="Something you'd like to see, or something that's not working right..."
                    rows={3}
                    className="w-full px-3 py-2 rounded-lg border text-sm outline-none resize-y"
                    style={{ borderColor: "#2A2A2F", background: "#0B0B0C", color: "#F2F0EA" }}
                  />
                  <div className="flex items-center gap-3 mt-2">
                    <button onClick={handleSendFeedback} disabled={!feedbackText.trim() || feedbackStatus === "sending"}
                      className="text-xs font-semibold px-4 py-2 rounded-lg transition-opacity"
                      style={{ background: "#C9A567", color: "#0B0B0C", opacity: (!feedbackText.trim() || feedbackStatus === "sending") ? 0.5 : 1, fontFamily: "'IBM Plex Mono', monospace" }}>
                      {feedbackStatus === "sending" ? "Sending…" : "Send feedback"}
                    </button>
                    {feedbackStatus === "sent" && <span className="text-xs" style={{ color: "#C9A567" }}>Sent — thank you!</span>}
                    {feedbackStatus === "error" && <span className="text-xs" style={{ color: "#C9695E" }}>Couldn't send — try again shortly.</span>}
                  </div>
                </div>

                <div className="grid grid-cols-3 gap-3 mb-4">
                  {[5, 10, 20].map((amt) => (
                    <button key={amt} onClick={() => handleSupport(amt)} disabled={upgradeBusy}
                      className="py-3 rounded-lg text-sm font-semibold border transition-opacity"
                      style={{ borderColor: "#C9A567", color: "#C9A567", opacity: upgradeBusy ? 0.6 : 1, fontFamily: "'IBM Plex Mono', monospace" }}>
                      RM{amt}
                    </button>
                  ))}
                </div>
                <p className="text-xs" style={{ color: "#75726C" }}>One-time, no account changes, no subscription — just a thank-you.</p>

                <div className="mt-6 pt-6 border-t" style={{ borderColor: "#232327" }}>
                  <p className="text-xs mb-1" style={{ color: "#75726C" }}>Questions or feedback:</p>
                  <a href="mailto:Unknownlable00@gmail.com" className="text-sm font-medium" style={{ color: "#C9A567" }}>Unknownlable00@gmail.com</a>
                </div>
              </div>
            </div>
          )}
        </main>
      )}

      {detailItem && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4" style={{ background: "rgba(36,31,26,0.55)" }} onClick={() => setDetailId(null)}>
          <div className="max-w-md w-full rounded-xl border p-6 relative max-h-[85vh] overflow-y-auto shadow-sm" style={{ background: "#17171A", borderColor: "#2A2A2F" }} onClick={(e) => e.stopPropagation()}>
            <button onClick={() => setDetailId(null)} className="absolute top-4 right-4 w-7 h-7 rounded-full flex items-center justify-center" style={{ background: "#1E1E22", color: "#F2F0EA" }} aria-label="Close">
              <X size={15} />
            </button>

            {(() => {
              const photos = detailItem.photos && detailItem.photos.length > 0 ? detailItem.photos : [];
              const touchRef = { current: null };
              return photos.length > 0 ? (
                <div className="mb-4">
                  <img
                    src={photos[Math.min(galleryIndex, photos.length - 1)]}
                    alt=""
                    onClick={() => setLightbox({ photos, index: Math.min(galleryIndex, photos.length - 1) })}
                    onTouchStart={(e) => { touchRef.current = e.touches[0].clientX; }}
                    onTouchEnd={(e) => {
                      if (touchRef.current == null) return;
                      const dx = e.changedTouches[0].clientX - touchRef.current;
                      if (dx > 50) setGalleryIndex((g) => (g - 1 + photos.length) % photos.length);
                      else if (dx < -50) setGalleryIndex((g) => (g + 1) % photos.length);
                      touchRef.current = null;
                    }}
                    className="w-full h-56 object-cover rounded-xl border-2 cursor-pointer" style={{ borderColor: "#2A2A2F" }}
                  />
                  {photos.length > 1 && (
                    <div className="flex gap-2 mt-2 overflow-x-auto pb-1">
                      {photos.map((p, idx) => (
                        <img key={idx} src={p} alt="" onClick={() => setGalleryIndex(idx)}
                          className="w-12 h-12 rounded-lg object-cover shrink-0 cursor-pointer border-2"
                          style={{ borderColor: idx === galleryIndex ? "#C9A567" : "#2A2A2F" }} />
                      ))}
                    </div>
                  )}
                  <div className="flex items-center gap-2 mt-3 flex-wrap">
                    <button onClick={() => downloadDataUrl(photos[galleryIndex], `${sanitizeForFilename(detailItem.sku)}_${String(galleryIndex + 1).padStart(2, "0")}.jpg`)}
                      className="flex items-center gap-1.5 text-xs font-medium px-3 py-1.5 rounded-lg border" style={{ borderColor: "#2A2A2F", color: "#F2F0EA", fontFamily: "'IBM Plex Mono', monospace" }}>
                      <Download size={13} /> Download photo
                    </button>
                    {photos.length > 1 && (
                      <button onClick={() => downloadAllPhotos(detailItem)}
                        className="flex items-center gap-1.5 text-xs font-medium px-3 py-1.5 rounded-lg" style={{ background: "#C9A567", color: "#0B0B0C", fontFamily: "'IBM Plex Mono', monospace" }}>
                        <Download size={13} /> Download all ({photos.length})
                      </button>
                    )}
                  </div>
                </div>
              ) : (
                <div className="w-full h-40 rounded-xl border-2 border-dashed flex items-center justify-center mb-4" style={{ borderColor: "#2A2A2F", color: "#2A2A2F" }}>
                  <ImageIcon size={28} />
                </div>
              );
            })()}

            <h3 className="text-xl font-semibold mb-1" style={{ color: "#F2F0EA", fontFamily: "'Fraunces', serif" }}>{productTitle(detailItem)}</h3>
            <div className="flex items-center gap-2 mb-4 flex-wrap">
              <span className="inline-flex items-center gap-1 text-xs font-semibold px-2 py-0.5 rounded-full" style={{ background: "#241F13", color: "#F2F0EA", fontFamily: "'IBM Plex Mono', monospace" }}>
                <Tag size={11} /> {CATEGORY_LABELS[detailItem.category] || "—"}
              </span>
              <span className="inline-block text-xs font-semibold px-2 py-0.5 rounded-full"
                style={{ background: detailItem.status === "sold" ? "#1C2420" : "#241F13", color: detailItem.status === "sold" ? "#8C8A85" : "#F2F0EA", fontFamily: "'IBM Plex Mono', monospace" }}>
                {(detailItem.status || "listed").toUpperCase()}
              </span>
            </div>

            <dl className="space-y-2 text-sm">
              <DetailRow label="Item type" value={detailItem.itemType || "—"} />
              <DetailRow label="Brand" value={detailItem.brand || "—"} />
              <DetailRow label="Colour" value={detailItem.colour || "—"} />
              <DetailRow label="Colour condition" value={detailItem.colourCondition ? `${detailItem.colourCondition}/10` : "—"} />

              {detailItem.category === "apparel" ? (
                (() => {
                  const sub = getApparelSubcategory(detailItem.itemType);
                  const a = detailItem.attributes || {};
                  const suggested = computeSuggestedSize(a, sub);
                  return (
                    <>
                      <DetailRow label="Size" value={a.size || suggested || "—"} />
                      {apparelMeasureParts(a, sub).map((line) => {
                        const [label, value] = line.split(": ");
                        return <DetailRow key={label} label={label} value={value} />;
                      })}
                    </>
                  );
                })()
              ) : detailItem.category === "other" ? (
                (detailItem.customAttributes || []).filter((c) => c.label && c.value).map((c) => <DetailRow key={c.label} label={c.label} value={c.value} />)
              ) : (
                (CATEGORY_DEFS[detailItem.category]?.attributes || [])
                  .filter((d) => detailItem.attributes?.[d.id])
                  .map((d) => <DetailRow key={d.id} label={d.label} value={detailItem.attributes[d.id]} />)
              )}

              <DetailRow label="Condition" value={detailItem.condition ? `${detailItem.condition}/10` : "—"} />
              <DetailRow label="Remark" value={detailItem.remark || "—"} />
              <DetailRow label="Cost" value={`RM${detailItem.cost || 0}`} />
              <DetailRow label="Price" value={`RM${detailItem.price || 0}`} />
              <DetailRow label="SKU" value={detailItem.sku || "—"} />
              <DetailRow label="Date added" value={detailItem.addedOn || "—"} />
              {detailItem.status === "sold" && <DetailRow label="Date sold" value={detailItem.soldOn || "—"} />}
            </dl>

            <div className="flex gap-2 mt-4">
              <button
                onClick={() => { startEditItem(detailItem); setDetailId(null); }}
                className="flex-1 flex items-center justify-center gap-2 py-2.5 rounded-xl text-sm font-semibold border"
                style={{ borderColor: "#2A2A2F", color: "#F2F0EA", fontFamily: "'IBM Plex Mono', monospace" }}>
                Edit
              </button>
              <button
                onClick={() => { setActiveId(detailItem.id); setTab("crosslist"); setDetailId(null); }}
                className="flex-1 flex items-center justify-center gap-2 py-2.5 rounded-xl text-sm font-semibold"
                style={{ background: "#C9A567", color: "#0B0B0C", fontFamily: "'IBM Plex Mono', monospace" }}>
                Open in Cross-List
              </button>
            </div>
          </div>
        </div>
      )}

      {selectedInvoice && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4" style={{ background: "rgba(23,22,26,0.55)" }} onClick={() => setSelectedInvoiceId(null)}>
          <div className="max-w-md w-full rounded-xl border p-6 relative max-h-[85vh] overflow-y-auto shadow-sm" style={{ background: "#17171A", borderColor: "#2A2A2F" }} onClick={(e) => e.stopPropagation()}>
            <button onClick={() => setSelectedInvoiceId(null)} className="absolute top-4 right-4 w-7 h-7 rounded-full flex items-center justify-center" style={{ background: "#1E1E22", color: "#F2F0EA" }} aria-label="Close">
              <X size={15} />
            </button>

            <p className="text-xs font-semibold mb-1" style={{ fontFamily: "'IBM Plex Mono', monospace", color: "#C9A567" }}>{selectedInvoice.invoiceNumber}</p>
            <h3 className="text-lg font-semibold mb-1" style={{ color: "#F2F0EA", fontFamily: "'Fraunces', serif" }}>
              {new Date(selectedInvoice.timestamp).toLocaleDateString(undefined, { day: "2-digit", month: "long", year: "numeric" })}
            </h3>
            <p className="text-xs mb-4" style={{ color: "#75726C" }}>
              {new Date(selectedInvoice.timestamp).toLocaleTimeString(undefined, { hour: "2-digit", minute: "2-digit" })} · {selectedInvoice.method}
            </p>

            <div className="space-y-3 mb-4">
              {selectedInvoice.lines.map((l, idx) => {
                const original = Number(l.price || 0);
                return (
                  <div key={idx} className="pb-3 border-b" style={{ borderColor: "#232327" }}>
                    <p className="text-sm font-medium mb-1">{l.title}</p>
                    <div className="grid grid-cols-3 gap-2 text-xs" style={{ color: "#8C8A85" }}>
                      <div><span className="block text-[10px] uppercase" style={{ color: "#75726C" }}>Original</span>RM{original}</div>
                      <div><span className="block text-[10px] uppercase" style={{ color: "#75726C" }}>Discount</span>RM{l.discount || 0}</div>
                      <div><span className="block text-[10px] uppercase" style={{ color: "#75726C" }}>Final</span><span style={{ color: "#C9A567", fontWeight: 600 }}>RM{l.finalPrice}</span></div>
                    </div>
                  </div>
                );
              })}
            </div>

            <dl className="space-y-2 text-sm">
              <DetailRow label="Subtotal" value={`RM${selectedInvoice.subtotal ?? selectedInvoice.total}`} />
              <DetailRow label="Total discount" value={`RM${selectedInvoice.discount || 0}`} />
              <DetailRow label="Final total" value={`RM${selectedInvoice.total}`} />
            </dl>
          </div>
        </div>
      )}

      {receiptTxn && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 receipt-overlay" style={{ background: "rgba(23,22,26,0.6)" }} onClick={() => setReceiptTxn(null)}>
          <style>{`
            @media print {
              body * { visibility: hidden; }
              #print-receipt, #print-receipt * { visibility: visible; }
              #print-receipt { position: fixed; top: 0; left: 0; width: 100%; padding: 0; margin: 0; }
              .receipt-overlay { position: static !important; background: none !important; padding: 0 !important; }
              .no-print { display: none !important; }
            }
          `}</style>
          <div className="w-full max-w-xs rounded-xl border overflow-hidden" style={{ background: "#FFFFFF", borderColor: "#2A2A2F" }} onClick={(e) => e.stopPropagation()}>
            <div id="print-receipt" className="p-6" style={{ fontFamily: "'IBM Plex Mono', monospace", color: "#17161A" }}>
              {logoUrl ? (
                <img src={logoUrl} alt="" className="mx-auto mb-2 h-12 object-contain" style={{ filter: "grayscale(100%)" }} />
              ) : (
                <p className="text-center text-sm font-bold mb-1">RAVUNO</p>
              )}
              <p className="text-center text-[10px] mb-4" style={{ color: "#555" }}>Sale Receipt</p>
              <div className="text-xs mb-3 pb-3 border-b border-dashed" style={{ borderColor: "#999" }}>
                <div className="flex justify-between"><span>Invoice</span><span>{receiptTxn.invoiceNumber}</span></div>
                <div className="flex justify-between"><span>Date</span><span>{new Date(receiptTxn.timestamp).toLocaleDateString()}</span></div>
                <div className="flex justify-between"><span>Time</span><span>{new Date(receiptTxn.timestamp).toLocaleTimeString(undefined, { hour: "2-digit", minute: "2-digit" })}</span></div>
                <div className="flex justify-between"><span>Payment</span><span>{receiptTxn.method}</span></div>
              </div>
              <div className="text-xs mb-3 pb-3 border-b border-dashed" style={{ borderColor: "#999" }}>
                {receiptTxn.lines.map((l, idx) => (
                  <div key={idx} className="flex justify-between mb-1">
                    <span className="truncate pr-2">{l.title}</span>
                    <span className="shrink-0">RM{l.finalPrice}</span>
                  </div>
                ))}
              </div>
              <div className="text-xs mb-1"><div className="flex justify-between"><span>Subtotal</span><span>RM{receiptTxn.subtotal}</span></div></div>
              {receiptTxn.discount > 0 && (
                <div className="text-xs mb-1"><div className="flex justify-between"><span>Discount</span><span>− RM{receiptTxn.discount}</span></div></div>
              )}
              <div className="text-sm font-bold flex justify-between pt-2 mt-2 border-t" style={{ borderColor: "#17161A" }}>
                <span>TOTAL</span><span>RM{receiptTxn.total}</span>
              </div>
              <p className="text-center text-[10px] mt-5" style={{ color: "#555" }}>Thank you for your purchase!</p>
            </div>
            <div className="no-print flex gap-2 p-4 border-t" style={{ borderColor: "#EDEAE2" }}>
              <button onClick={() => window.print()}
                className="flex-1 py-2.5 rounded-lg text-sm font-semibold" style={{ background: "#C9A567", color: "#0B0B0C" }}>
                Print / Save as PDF
              </button>
              <button onClick={() => setReceiptTxn(null)}
                className="px-4 py-2.5 rounded-lg text-sm font-medium border" style={{ borderColor: "#E4E1D9", color: "#17161A" }}>
                Done
              </button>
            </div>
          </div>
        </div>
      )}

      {lightbox && (
        <Lightbox
          photos={lightbox.photos}
          index={lightbox.index}
          onIndexChange={(idx) => { setLightbox((lb) => ({ ...lb, index: idx })); if (detailItem) setGalleryIndex(idx); }}
          onClose={() => setLightbox(null)}
          onDownload={(idx) => downloadDataUrl(lightbox.photos[idx], `${sanitizeForFilename(detailItem?.sku)}_${String(idx + 1).padStart(2, "0")}.jpg`)}
        />
      )}
    </div>
  );
}
