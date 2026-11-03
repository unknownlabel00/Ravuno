import React, { useState, useEffect, useMemo, useRef } from "react";
import {
  Check, Copy, Ruler, Plus, Trash2, ChevronRight, ChevronLeft,
  Package, TrendingUp, Clock, Image as ImageIcon, X, Loader2,
  Search, Download, Star, Tag, Receipt, ShoppingBag, Percent, DollarSign, Eye, MessageCircle, Users, Link2
} from "lucide-react";
import * as XLSX from "xlsx";
import Papa from "papaparse";
import { LineChart, Line, BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer } from "recharts";

// RAVUNO Vault Ledger tokens — bg: #0B0B0C  surface: #17171A  hairline: #2A2A2F  brass: #C9A567  ivory: #F2F0EA  muted: #8C8A85
// border: #2A2A2F  accent: #C9A567  destructive: #C9695E
const PLATFORMS = [
  { id: "carousell", label: "Carousell" },
  { id: "facebook", label: "Facebook Marketplace" },
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
const signInRequest = (email, password) => authRequest("/token?grant_type=password", { email, password });
const refreshSessionRequest = (refresh_token) => authRequest("/token?grant_type=refresh_token", { refresh_token });
const recoverPasswordRequest = (email) => authRequest("/recover", { email });
async function updatePasswordRequest(accessToken, newPassword) {
  const res = await fetch(`${AUTH_URL}/user`, {
    method: "PUT",
    headers: { "Content-Type": "application/json", apikey: SUPABASE_ANON_KEY, Authorization: `Bearer ${accessToken}` },
    body: JSON.stringify({ password: newPassword }),
  });
  const data = await res.json().catch(() => ({}));
  if (!res.ok) throw new Error(data.error_description || data.msg || data.error || "Couldn't update password.");
  return data;
}
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

// for the public storefront — no logged-in session exists, relies entirely on
// the public RLS policies on storefronts/items for anonymous visitors
async function publicApiRequest(path) {
  const res = await fetch(`${REST_URL}${path}`, { headers: { apikey: SUPABASE_ANON_KEY } });
  if (!res.ok) throw new Error("Request failed");
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
    id: row.id, title: row.title || "", category: row.category || "", itemType: normalizeItemType(row.item_type),
    brand: row.brand || "", condition: row.condition || "", colour: row.colour || "", colourCondition: row.colour_condition || "", remark: row.remark || "",
    cost: row.cost ?? "", price: row.price ?? "", sku: row.sku || "", status: row.status || "listed",
    attributes: row.attributes || {}, customAttributes: row.custom_attributes || [], photos: row.photos || [],
    addedOn: row.added_on, listedOn: row.listed_on, soldOn: row.sold_on,
    showOnStorefront: row.show_on_storefront ?? false,
    featured: row.featured ?? false,
    sale: row.sale || null,
  };
}
function itemToRow(item, userId) {
  return {
    user_id: userId, title: item.title || null, category: item.category || null, item_type: item.itemType || null,
    brand: item.brand || null, condition: item.condition || null, colour: item.colour || null, colour_condition: item.colourCondition || null, remark: item.remark || null,
    cost: item.cost === "" ? null : item.cost, price: item.price === "" ? null : item.price, sku: item.sku || null,
    status: item.status || "listed", attributes: item.attributes || {}, custom_attributes: item.customAttributes || [],
    photos: item.photos || [], added_on: item.addedOn || null, listed_on: item.listedOn || null, sold_on: item.soldOn || null,
    show_on_storefront: item.showOnStorefront ?? false,
    featured: item.featured ?? false,
    sale: item.sale ?? null,
  };
}
function rowToTxn(row) {
  return { id: row.id, timestamp: row.occurred_at, method: row.method, subtotal: row.subtotal, discount: row.discount, total: row.total, postage: Number(row.postage || 0), lines: row.lines || [] };
}
function txnToRow(txn, userId) {
  return { user_id: userId, occurred_at: txn.timestamp, method: txn.method, subtotal: txn.subtotal, discount: txn.discount, total: txn.total, postage: Number(txn.postage || 0), lines: txn.lines };
}

function rowToPurchase(row) {
  return {
    id: row.id, sku: row.sku || "", cost: row.cost ?? "", place: row.place || "",
    photo: row.photo || "", receiptPhoto: row.receipt_photo || "",
    purchasedOn: row.purchased_on, createdAt: row.created_at,
  };
}
function purchaseToRow(p, userId) {
  return {
    user_id: userId, sku: p.sku || null, cost: p.cost === "" ? null : p.cost, place: p.place || null,
    photo: p.photo || null, receipt_photo: p.receiptPhoto || null,
    purchased_on: p.purchasedOn || new Date().toISOString().slice(0, 10),
  };
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

// ITEM-TYPES-START
const APPAREL_TOP_TYPES = ["T-Shirt", "Shirt", "Polo", "Hoodie", "Jacket", "Sweatshirt", "Crewneck", "Other topwear"];
const APPAREL_BOTTOM_TYPES = ["Jeans", "Pants", "Shorts", "Skirt", "Dress", "Other bottomwear"];
// "Sweater" was renamed to "Sweatshirt"; anything still saved or imported under the old name is read as the new one
const normalizeItemType = (t) => (String(t || "").trim().toLowerCase() === "sweater" ? "Sweatshirt" : String(t || "").trim());
// ITEM-TYPES-END

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

Shipping Semenanjung RM7.50 | Sabah Sarawak RM15
SKU: ${p.sku || "[XXX]"}

Thanks for checking out my shop!`;

  const threads = `${title}

Condition: ${condLine}
${measureLine}
Remark: ${p.remark || "[key detail]"}

Price: RM${p.price || "?"}
SKU: ${p.sku || "?"}

Free shipping 🚚`;

  const grailed = `${title} — Size ${size || "[Size]"}

PLEASE READ BEFORE PURCHASING

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

  const shopee = `${title} — Size ${size || "[Size]"}

🛍️ PRE-OWNED / VINTAGE CLOTHING

Thank you for checking out my small shop! ♡
Feel free to message me if you have any questions. I'll reply as soon as possible!

━━━━━━━━━━━━━━━━━━

🏷️ ITEM DETAILS

👕 Item: ${p.itemType || "[Hoodie / Sweatshirt / Pants / etc.]"}
🏷️ Brand: ${p.brand || "[Brand Name]"}
📌 Size on Tag: ${size || "[Size]"}
🎨 Colour: ${p.colour || "[Colour]"}
⭐ Condition: ${condLine}

📝 Remarks: ${p.remark || "[Faded / Minor Flaws / No Major Flaws]"}

━━━━━━━━━━━━━━━━━━

📏 MEASUREMENTS

${measureBullets}

⚠️ Please compare the measurements with your own garments to ensure a proper fit. Sizing may vary between brands and vintage pieces.

━━━━━━━━━━━━━━━━━━

🔎 CONDITION GUIDE

♻️ All items are pre-owned unless stated otherwise.
🧵 Signs of wear may be present due to previous use and age.
📸 Any noticeable flaws will be mentioned in the remarks and shown in the photos where possible.

━━━━━━━━━━━━━━━━━━

🚚 SHIPPING INFORMATION

📦 Shipping will be arranged through Shopee.
📍 Please ensure your shipping address and contact details are correct before placing your order.

━━━━━━━━━━━━━━━━━━

❗ IMPORTANT NOTES

• Please review all photos, measurements, and condition details before purchasing.
• Colours may vary slightly due to lighting and screen settings.
• Feel free to message me for additional photos or measurements.

━━━━━━━━━━━━━━━━━━

🗂️ ADDITIONAL INFO

🔖 SKU: ${p.sku || "[Item Code]"}

━━━━━━━━━━━━━━━━━━

🤍 Thank you for supporting my small shop!
Happy shopping!`;

  return { carousell, facebook: carousell, threads, grailed, depop, shopee, sizeFlag };
}

function buildGenericTemplates(p) {
  const title = productTitle(p);
  const attrLines = getFilledAttrLines(p);
  const condLine = p.colourCondition ? `${p.condition || "?"}/10 (Colour ${p.colourCondition}/10)` : `${p.condition || "?"}/10`;

  const carousell = `${title}

Condition: ${condLine}
${attrLines.length ? attrLines.join("\n") + "\n" : ""}${p.colour ? `Colour: ${p.colour}\n` : ""}Remark: ${p.remark || "[key detail]"}

Shipping Semenanjung RM7.50 | Sabah Sarawak RM15
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

  return { carousell, facebook: carousell, threads, grailed, depop, shopee, sizeFlag: null };
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
function CopyButton({ text, label = "Copy" }) {
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
      {copied ? "Copied" : label}
    </button>
  );
}
// splits a generated listing into Title (its first line) and Description (everything after the
// first blank line) — works reliably for the built-in templates, and for custom ones as long as
// the title is kept on its own first line
function splitTitleDescription(text) {
  const str = text || "";
  const firstBreak = str.indexOf("\n");
  if (firstBreak === -1) return { title: str, description: "" };
  const title = str.slice(0, firstBreak).trim();
  const description = str.slice(firstBreak).replace(/^\s+/, "");
  return { title, description };
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
// SALE-LOGIC-START
// a sale rule is { kind: "percent" | "off" | "price", value }. "off" and "price" only touch pieces that cost more than the value,
// so a piece that is already at or below the sale price is left alone.
function applySale(price, rule) {
  const p = Number(price) || 0, v = rule ? Number(rule.value) : 0;
  if (!(p > 0) || !(v > 0)) return p;
  if (rule.kind === "percent") return v < 100 ? Math.round(p * (100 - v)) / 100 : p;
  if (rule.kind === "off") return p > v ? Math.round((p - v) * 100) / 100 : p;
  if (rule.kind === "price") return p > v ? v : p;
  return p;
}
// which rule wins: the piece's own, else the best collection it is in, else the store-wide one (the last two only for pieces showing on the storefront)
function saleFor(item, ctx) {
  const price = Number((item && item.price) || 0);
  const none = { rule: null, price, off: 0, badge: "", offText: "" };
  if (!item || !(price > 0)) return none;
  let rule = null;
  if (item.sale) rule = item.sale;
  else if (item.showOnStorefront) {
    const key = String(item.id);
    const inCollections = ((ctx && ctx.collections) || []).filter((c) => c.sale && (c.itemIds || []).includes(key));
    if (inCollections.length) rule = inCollections.map((c) => c.sale).sort((x, y) => applySale(price, x) - applySale(price, y))[0];
    else if (ctx && ctx.store) rule = ctx.store;
  }
  if (!rule) return none;
  const next = applySale(price, rule);
  const off = Math.round((price - next) * 100) / 100;
  if (!(off > 0)) return { ...none, rule };
  const isPct = rule.kind === "percent";
  return { rule, price: next, off, badge: isPct ? `-${Number(rule.value)}%` : `-${fmtRM(off)}`, offText: isPct ? `${Number(rule.value)}% off` : `${fmtRM(off)} off` };
}
const fmtRM = (n) => `RM${Number.isInteger(n) ? n : n.toFixed(2)}`;
const ruleText = (r) => (r.kind === "percent" ? `${r.value}% off` : r.kind === "off" ? `RM${r.value} off` : `RM${r.value} each`);
const ruleLong = (r) => (r.kind === "percent" ? `${r.value}% off everything` : r.kind === "off" ? `RM${r.value} off every piece that costs more than RM${r.value}` : `every piece above RM${r.value} drops to RM${r.value}`);
const ruleToDraft = (r) => (r ? { kind: r.kind, value: String(r.value) } : { kind: "percent", value: "" });
function parseRule(d) {
  const v = Number(d.value);
  if (!(v > 0)) return null;
  if (d.kind === "percent") { const n = Math.round(v); return n >= 1 && n <= 99 ? { kind: "percent", value: n } : null; }
  return { kind: d.kind, value: Math.round(v * 100) / 100 };
}
// SALE-LOGIC-END

// columns the public storefront is allowed to read — deliberately excludes cost and SKU
const STOREFRONT_ITEM_COLUMNS = "id,user_id,title,category,item_type,brand,condition,colour,colour_condition,remark,price,status,attributes,custom_attributes,photos,show_on_storefront,featured,sale,created_at";

// ---------- storefront v5: routes, banner slideshow, and the owner's editors for banners / FAQ / collections ----------
function parseStoreHash() {
  const h = typeof window !== "undefined" ? window.location.hash : "";
  let m;
  if ((m = h.match(/^#item=([A-Za-z0-9-]+)/))) return { page: "home", item: m[1], collection: null };
  if ((m = h.match(/^#collection=([A-Za-z0-9-]+)/))) return { page: "collection", item: null, collection: m[1] };
  if (h === "#sale") return { page: "sale", item: null, collection: null };
  if (h === "#faq") return { page: "faq", item: null, collection: null };
  if (h === "#collections") return { page: "collections", item: null, collection: null };
  return { page: "home", item: null, collection: null }; // includes #shop
}
function rowToCollection(r) {
  return { id: r.id, sale: r.sale || null, title: r.title || "", description: r.description || "", coverUrl: r.cover_url || "", itemIds: Array.isArray(r.item_ids) ? r.item_ids.map(String) : [] };
}

function Marquee({ text, background, color }) {
  const wrapRef = useRef(null), trackRef = useRef(null), unitRef = useRef(null);
  const [reps, setReps] = useState(6);
  // how many copies of the text the screen needs: enough to cover its width plus one spare
  useEffect(() => {
    const measure = () => {
      const unit = unitRef.current ? unitRef.current.offsetWidth : 0;
      const w = wrapRef.current ? wrapRef.current.offsetWidth : (typeof window !== "undefined" ? window.innerWidth : 400);
      if (unit > 0) setReps(Math.max(3, Math.ceil(w / unit) + 2));
    };
    measure();
    const t = setTimeout(measure, 400); // web fonts finishing loading change the text width
    window.addEventListener("resize", measure);
    return () => { clearTimeout(t); window.removeEventListener("resize", measure); };
  }, [text]);
  // slide left one copy at a time, then loop; repainted on every frame so it can't stall or go blank
  useEffect(() => {
    if (typeof window !== "undefined" && window.matchMedia && window.matchMedia("(prefers-reduced-motion: reduce)").matches) return;
    let raf, last = null, x = 0;
    const tick = (now) => {
      if (last == null) last = now;
      const unit = unitRef.current ? unitRef.current.offsetWidth : 0;
      if (unit > 0 && trackRef.current) {
        x = (x + ((now - last) / 1000) * 60) % unit;
        trackRef.current.style.transform = `translate3d(${-x}px, 0, 0)`;
      }
      last = now;
      raf = requestAnimationFrame(tick);
    };
    raf = requestAnimationFrame(tick);
    return () => cancelAnimationFrame(raf);
  }, [text]);
  return (
    <div ref={wrapRef} data-marquee="" aria-label={text} style={{ background, color, overflow: "hidden" }}>
      <div ref={trackRef} data-marquee-track="" aria-hidden="true"
        style={{ display: "flex", width: "max-content", padding: "8px 0", whiteSpace: "nowrap", fontSize: 11, letterSpacing: "0.2em", textTransform: "uppercase", willChange: "transform" }}>
        {Array.from({ length: reps }).map((_, i) => <span key={i} ref={i === 0 ? unitRef : null} style={{ padding: "0 2rem", flexShrink: 0 }}>{text}</span>)}
      </div>
    </div>
  );
}

function BannerCarousel({ slides, target }) {
  const [idx, setIdx] = useState(0);
  const [paused, setPaused] = useState(false);
  const touchX = useRef(null);
  const n = slides.length;
  useEffect(() => {
    if (n < 2 || paused) return;
    if (typeof window !== "undefined" && window.matchMedia && window.matchMedia("(prefers-reduced-motion: reduce)").matches) return;
    const t = setInterval(() => setIdx((i) => (i + 1) % n), 5000);
    return () => clearInterval(t);
  }, [n, paused]);
  useEffect(() => { if (idx >= n) setIdx(0); }, [n, idx]);
  if (n === 0) return null;
  const step = (d) => setIdx((i) => (i + d + n) % n);
  return (
    <div className="w-full mx-auto relative overflow-hidden" style={{ maxWidth: 1200 }}
      onMouseEnter={() => setPaused(true)} onMouseLeave={() => setPaused(false)}
      onTouchStart={(e) => { touchX.current = e.touches[0].clientX; setPaused(true); }}
      onTouchEnd={(e) => {
        const dx = touchX.current == null ? 0 : e.changedTouches[0].clientX - touchX.current;
        touchX.current = null; setPaused(false);
        if (Math.abs(dx) > 40) step(dx < 0 ? 1 : -1);
      }}>
      <div className="flex" style={{ transform: `translateX(-${idx * 100}%)`, transition: "transform 600ms ease" }}>
        {slides.map((s, i) => {
          const t = target(s.link);
          const img = <img src={s.url} alt="" className="w-full block object-cover" style={{ aspectRatio: "3 / 2", objectPosition: "center" }} draggable={false} />;
          return (
            <div key={i} className="w-full shrink-0" aria-hidden={i !== idx}>
              {t ? (
                <a href={t.href} onClick={t.onClick} target={t.external ? "_blank" : undefined} rel={t.external ? "noreferrer" : undefined} tabIndex={i === idx ? 0 : -1} className="block">{img}</a>
              ) : img}
            </div>
          );
        })}
      </div>
      {n > 1 && (
        <>
          <div className="absolute bottom-3 left-0 right-0 flex justify-center gap-2">
            {slides.map((_, i) => (
              <button key={i} onClick={() => setIdx(i)} aria-label={`Slide ${i + 1}`} className="w-2 h-2 rounded-full" style={{ background: i === idx ? "#FFFFFF" : "rgba(255,255,255,0.5)" }} />
            ))}
          </div>
          <button onClick={() => step(-1)} aria-label="Previous slide" className="hidden md:flex absolute left-3 top-1/2 -translate-y-1/2 w-9 h-9 rounded-full items-center justify-center" style={{ background: "rgba(20,20,20,0.5)", color: "#FFFFFF" }}><ChevronLeft size={18} /></button>
          <button onClick={() => step(1)} aria-label="Next slide" className="hidden md:flex absolute right-3 top-1/2 -translate-y-1/2 w-9 h-9 rounded-full items-center justify-center" style={{ background: "rgba(20,20,20,0.5)", color: "#FFFFFF" }}><ChevronRight size={18} /></button>
        </>
      )}
    </div>
  );
}

const OWNER_CARD = { background: "#17171A", borderColor: "#2A2A2F" };
const OWNER_INPUT = { borderColor: "#2A2A2F", background: "#0B0B0C", color: "#F2F0EA" };
const OWNER_H2 = { color: "#F2F0EA", fontFamily: "'Fraunces', serif" };

const RULE_KINDS = [["percent", "% off"], ["off", "RM off"], ["price", "Set price"]];
function RuleEditor({ rule, onChange }) {
  const hint = { percent: "Takes this % off each piece.", off: "Takes this many RM off each piece that costs more than that.", price: "Every piece above this price drops to it. Cheaper pieces stay as they are." }[rule.kind];
  return (
    <div>
      <div className="flex gap-1.5 mb-2 flex-wrap">
        {RULE_KINDS.map(([k, label]) => (
          <button key={k} type="button" data-kind={k} onClick={() => onChange({ ...rule, kind: k })} className="text-xs font-semibold px-3 py-1.5 rounded-lg border"
            style={{ borderColor: rule.kind === k ? "#C9A567" : "#2A2A2F", color: rule.kind === k ? "#C9A567" : "#8C8A85", background: rule.kind === k ? "#241F13" : "transparent" }}>{label}</button>
        ))}
      </div>
      <div className="flex items-center gap-2">
        <span className="text-sm" style={{ color: "#8C8A85" }}>{rule.kind === "percent" ? "%" : "RM"}</span>
        <input type="number" inputMode="decimal" min="0" value={rule.value} onChange={(e) => onChange({ ...rule, value: e.target.value })}
          placeholder={rule.kind === "off" ? "5" : "10"} className="w-24 px-3 py-2 rounded-lg border text-sm outline-none" style={{ borderColor: "#2A2A2F", background: "#0B0B0C", color: "#F2F0EA" }} />
      </div>
      <p className="text-[11px] mt-1.5" style={{ color: "#75726C" }}>{hint}</p>
    </div>
  );
}

function AnnouncementCard({ storefront, onSave }) {
  const [text, setText] = useState(storefront.announcement || "");
  const [saved, setSaved] = useState(false);
  useEffect(() => { setText(storefront.announcement || ""); }, [storefront.announcement]);
  return (
    <div className="rounded-xl border p-5" style={OWNER_CARD}>
      <h2 className="text-sm font-semibold mb-2" style={OWNER_H2}>Moving label</h2>
      <p className="text-xs mb-3" style={{ color: "#8C8A85" }}>A strip of text that scrolls across your storefront, below your logo and above the banner. Leave it empty to hide it.</p>
      <input value={text} onChange={(e) => setText(e.target.value)} maxLength={160} placeholder="Free shipping min. RM30"
        className="w-full px-3 py-2 rounded-lg border text-sm outline-none mb-3" style={OWNER_INPUT} />
      <button onClick={async () => { const ok = await onSave({ announcement: text.trim() || null }, { announcement: text.trim() }); setSaved(!!ok); setTimeout(() => setSaved(false), 1500); }}
        className="text-xs font-semibold px-4 py-2 rounded-lg" style={{ background: "#C9A567", color: "#0B0B0C" }}>{saved ? "Saved ✓" : "Save"}</button>
    </div>
  );
}

const LINK_CHOICES = [["", "No link"], ["shop", "Shop (all pieces)"], ["sale", "Sale page"], ["faq", "FAQ page"], ["collections", "All collections"]];
function BannerSlidesCard({ storefront, items, collections, onSaveBanners, onUploadImage }) {
  const slides = storefront.banners && storefront.banners.length ? storefront.banners : storefront.bannerUrl ? [{ url: storefront.bannerUrl, link: "" }] : [];
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");
  const [urlDraft, setUrlDraft] = useState({});
  const fileRef = useRef(null);
  const showing = (items || []).filter((x) => x.status !== "sold" && x.showOnStorefront);
  async function addFiles(fileList) {
    setBusy(true); setError("");
    try {
      const next = [...slides];
      for (const f of Array.from(fileList)) {
        if (next.length >= 6) break;
        next.push({ url: await onUploadImage(f), link: "" });
      }
      await onSaveBanners(next);
    } catch (e) { setError("Couldn't upload that picture. Try a smaller one."); } finally { setBusy(false); }
  }
  const move = (i, d) => { const next = [...slides]; const j = i + d; if (j < 0 || j >= next.length) return; [next[i], next[j]] = [next[j], next[i]]; onSaveBanners(next); };
  const remove = (i) => onSaveBanners(slides.filter((_, k) => k !== i));
  const setLink = (i, link) => onSaveBanners(slides.map((s, k) => (k === i ? { ...s, link } : s)));
  return (
    <div className="rounded-xl border p-5" style={OWNER_CARD}>
      <h2 className="text-sm font-semibold mb-2" style={OWNER_H2}>Top banner slides</h2>
      <p className="text-xs mb-3" style={{ color: "#8C8A85" }}>Up to 6 pictures that slide every few seconds. Choose where each one takes customers when tapped. Best size: 1200 × 800 px (3:2), with the main subject near the centre.</p>
      {slides.length === 0 && (
        <div className="w-full rounded-lg border-2 border-dashed flex items-center justify-center mb-3" style={{ borderColor: "#2A2A2F", aspectRatio: "3 / 2" }}><ImageIcon size={20} style={{ color: "#75726C" }} /></div>
      )}
      <div className="space-y-4 mb-3">
        {slides.map((s, i) => {
          const link = s.link || "";
          const kind = link.startsWith("url:") ? "url" : link;
          return (
            <div key={`${s.url}-${i}`} className="rounded-lg border p-3" style={{ borderColor: "#2A2A2F", background: "#0B0B0C" }}>
              <img src={s.url} alt="" className="w-full rounded mb-3 object-cover" style={{ aspectRatio: "3 / 2" }} />
              <p className="text-[10px] uppercase tracking-wide mb-1" style={{ color: "#75726C" }}>Tapping slide {i + 1} opens</p>
              <select value={kind} onChange={(e) => setLink(i, e.target.value === "url" ? "url:" : e.target.value)} className="w-full px-3 py-2 rounded-lg border text-xs outline-none mb-2" style={OWNER_INPUT}>
                {LINK_CHOICES.map(([v, label]) => <option key={v} value={v}>{label}</option>)}
                <option value="url">A web address…</option>
                {collections.length > 0 && <optgroup label="A collection">{collections.map((c) => <option key={c.id} value={`collection:${c.id}`}>{c.title}</option>)}</optgroup>}
                {showing.length > 0 && <optgroup label="A piece">{showing.map((x) => <option key={x.id} value={`item:${x.id}`}>{productTitle(x)}</option>)}</optgroup>}
              </select>
              {kind === "url" && (
                <input value={urlDraft[i] != null ? urlDraft[i] : link.slice(4)} onChange={(e) => setUrlDraft((d) => ({ ...d, [i]: e.target.value }))}
                  onBlur={() => { const v = (urlDraft[i] != null ? urlDraft[i] : link.slice(4)).trim(); if (/^https:\/\/\S+$/.test(v)) setLink(i, `url:${v}`); }}
                  placeholder="https://…" className="w-full px-3 py-2 rounded-lg border text-xs outline-none mb-2" style={OWNER_INPUT} />
              )}
              <div className="flex items-center gap-2">
                <button onClick={() => move(i, -1)} disabled={i === 0} className="text-xs px-3 py-1.5 rounded-lg border" style={{ borderColor: "#2A2A2F", color: "#F2F0EA", opacity: i === 0 ? 0.35 : 1 }}>↑</button>
                <button onClick={() => move(i, 1)} disabled={i === slides.length - 1} className="text-xs px-3 py-1.5 rounded-lg border" style={{ borderColor: "#2A2A2F", color: "#F2F0EA", opacity: i === slides.length - 1 ? 0.35 : 1 }}>↓</button>
                <button onClick={() => remove(i)} className="text-xs px-3 py-1.5 rounded-lg border ml-auto" style={{ borderColor: "#2A2A2F", color: "#C9695E" }}>Remove</button>
              </div>
            </div>
          );
        })}
      </div>
      {error && <p className="text-xs mb-2" style={{ color: "#C9695E" }}>{error}</p>}
      <button onClick={() => fileRef.current && fileRef.current.click()} disabled={busy || slides.length >= 6}
        className="text-xs font-semibold px-4 py-2 rounded-lg" style={{ background: "#C9A567", color: "#0B0B0C", opacity: (busy || slides.length >= 6) ? 0.5 : 1 }}>
        {busy ? "Uploading…" : slides.length >= 6 ? "6 slides is the maximum" : "Add slide pictures"}
      </button>
      <input ref={fileRef} type="file" accept="image/*" multiple className="hidden" onChange={(e) => { if (e.target.files && e.target.files.length) addFiles(e.target.files); e.target.value = ""; }} />
    </div>
  );
}

function FaqCard({ storefront, onSave }) {
  const [rows, setRows] = useState(storefront.faq || []);
  const [saved, setSaved] = useState(false);
  useEffect(() => { setRows(storefront.faq || []); }, [storefront.faq]);
  const edit = (i, key, v) => setRows((r) => r.map((x, k) => (k === i ? { ...x, [key]: v } : x)));
  const move = (i, d) => setRows((r) => { const next = [...r]; const j = i + d; if (j < 0 || j >= next.length) return r; [next[i], next[j]] = [next[j], next[i]]; return next; });
  async function save() {
    const clean = rows.map((r) => ({ q: (r.q || "").trim(), a: (r.a || "").trim() })).filter((r) => r.q && r.a);
    const ok = await onSave({ faq: clean }, { faq: clean });
    setSaved(!!ok); setTimeout(() => setSaved(false), 1500);
  }
  return (
    <div className="rounded-xl border p-5" style={OWNER_CARD}>
      <h2 className="text-sm font-semibold mb-2" style={OWNER_H2}>FAQ</h2>
      <p className="text-xs mb-3" style={{ color: "#8C8A85" }}>Questions customers ask most. They show on your storefront's FAQ tab. Check the pre-filled answers match how you actually work.</p>
      <div className="space-y-3 mb-3">
        {rows.map((r, i) => (
          <div key={i} className="rounded-lg border p-3" style={{ borderColor: "#2A2A2F", background: "#0B0B0C" }}>
            <input value={r.q || ""} onChange={(e) => edit(i, "q", e.target.value)} maxLength={140} placeholder="Question" className="w-full px-3 py-2 rounded-lg border text-sm outline-none mb-2" style={OWNER_INPUT} />
            <textarea value={r.a || ""} onChange={(e) => edit(i, "a", e.target.value)} maxLength={800} rows={3} placeholder="Answer" className="w-full px-3 py-2 rounded-lg border text-sm outline-none mb-2 resize-y" style={OWNER_INPUT} />
            <div className="flex items-center gap-2">
              <button onClick={() => move(i, -1)} disabled={i === 0} className="text-xs px-3 py-1.5 rounded-lg border" style={{ borderColor: "#2A2A2F", color: "#F2F0EA", opacity: i === 0 ? 0.35 : 1 }}>↑</button>
              <button onClick={() => move(i, 1)} disabled={i === rows.length - 1} className="text-xs px-3 py-1.5 rounded-lg border" style={{ borderColor: "#2A2A2F", color: "#F2F0EA", opacity: i === rows.length - 1 ? 0.35 : 1 }}>↓</button>
              <button onClick={() => setRows((x) => x.filter((_, k) => k !== i))} className="text-xs px-3 py-1.5 rounded-lg border ml-auto" style={{ borderColor: "#2A2A2F", color: "#C9695E" }}>Remove</button>
            </div>
          </div>
        ))}
        {rows.length === 0 && <p className="text-[11px]" style={{ color: "#75726C" }}>No questions yet.</p>}
      </div>
      <div className="flex items-center gap-2">
        <button onClick={() => setRows((r) => (r.length >= 30 ? r : [...r, { q: "", a: "" }]))} className="text-xs font-semibold px-4 py-2 rounded-lg border" style={{ borderColor: "#2A2A2F", color: "#F2F0EA" }}>Add question</button>
        <button onClick={save} className="text-xs font-semibold px-4 py-2 rounded-lg" style={{ background: "#C9A567", color: "#0B0B0C" }}>{saved ? "Saved ✓" : "Save FAQ"}</button>
      </div>
    </div>
  );
}

function CollectionsCard({ collections, items, onSave, onDelete, onUploadImage }) {
  const [editing, setEditing] = useState(null); // { id?, title, description, coverUrl, itemIds }
  const [search, setSearch] = useState("");
  const [busy, setBusy] = useState(false);
  const [confirmId, setConfirmId] = useState(null);
  const coverRef = useRef(null);
  const pickable = (items || []).filter((x) => x.status !== "sold" && (!search.trim() || productTitle(x).toLowerCase().includes(search.trim().toLowerCase())));
  const toggle = (id) => setEditing((e) => ({ ...e, itemIds: e.itemIds.includes(String(id)) ? e.itemIds.filter((x) => x !== String(id)) : [...e.itemIds, String(id)] }));
  async function save() {
    if (!editing.title.trim()) return;
    setBusy(true);
    const ok = await onSave(editing);
    setBusy(false);
    if (ok) setEditing(null);
  }
  return (
    <div className="rounded-xl border p-5" style={OWNER_CARD}>
      <h2 className="text-sm font-semibold mb-2" style={OWNER_H2}>Collections</h2>
      <p className="text-xs mb-3" style={{ color: "#8C8A85" }}>Hand-pick pieces into a themed landing page (for example "Y2K picks" or "Under RM20"). Customers find them under the Collections tab, and you can link a banner slide to one.</p>
      {!editing && (
        <>
          <div className="space-y-2 mb-3">
            {collections.map((c) => (
              <div key={c.id} className="flex items-center justify-between gap-3 px-3 py-2.5 rounded-lg border" style={{ borderColor: "#2A2A2F", background: "#0B0B0C" }}>
                <div className="min-w-0">
                  <p className="text-sm truncate" style={{ color: "#F2F0EA" }}>{c.title}</p>
                  <p className="text-[11px]" style={{ color: "#75726C" }}>{c.itemIds.length} piece{c.itemIds.length === 1 ? "" : "s"}</p>
                </div>
                <div className="flex items-center gap-2 shrink-0">
                  <button onClick={() => { setEditing({ ...c }); setSearch(""); }} className="text-xs px-3 py-1.5 rounded-lg border" style={{ borderColor: "#2A2A2F", color: "#F2F0EA" }}>Edit</button>
                  <button onClick={() => { if (confirmId === c.id) { onDelete(c.id); setConfirmId(null); } else setConfirmId(c.id); }} className="text-xs px-3 py-1.5 rounded-lg border" style={{ borderColor: "#2A2A2F", color: "#C9695E" }}>{confirmId === c.id ? "Tap again to delete" : "Delete"}</button>
                </div>
              </div>
            ))}
            {collections.length === 0 && <p className="text-[11px]" style={{ color: "#75726C" }}>No collections yet.</p>}
          </div>
          <button onClick={() => { setEditing({ title: "", description: "", coverUrl: "", itemIds: [] }); setSearch(""); }} className="text-xs font-semibold px-4 py-2 rounded-lg" style={{ background: "#C9A567", color: "#0B0B0C" }}>New collection</button>
        </>
      )}
      {editing && (
        <div>
          <input value={editing.title} onChange={(e) => setEditing({ ...editing, title: e.target.value })} maxLength={80} placeholder="Collection name" className="w-full px-3 py-2 rounded-lg border text-sm outline-none mb-2" style={OWNER_INPUT} />
          <textarea value={editing.description} onChange={(e) => setEditing({ ...editing, description: e.target.value })} maxLength={500} rows={2} placeholder="Short description (optional)" className="w-full px-3 py-2 rounded-lg border text-sm outline-none mb-3 resize-y" style={OWNER_INPUT} />
          <p className="text-[10px] uppercase tracking-wide mb-1.5" style={{ color: "#75726C" }}>Cover picture (optional)</p>
          {editing.coverUrl && <img src={editing.coverUrl} alt="" className="w-full rounded-lg mb-2 object-cover" style={{ aspectRatio: "3 / 2" }} />}
          <div className="flex items-center gap-2 mb-3">
            <button onClick={() => coverRef.current && coverRef.current.click()} className="text-xs font-semibold px-3 py-2 rounded-lg border" style={{ borderColor: "#2A2A2F", color: "#F2F0EA" }}>{editing.coverUrl ? "Replace cover" : "Upload cover"}</button>
            {editing.coverUrl && <button onClick={() => setEditing({ ...editing, coverUrl: "" })} className="text-xs px-3 py-2 rounded-lg border" style={{ borderColor: "#2A2A2F", color: "#C9695E" }}>Remove</button>}
            <span className="text-[11px]" style={{ color: "#75726C" }}>Without one, the first piece's photo is used.</span>
          </div>
          <input ref={coverRef} type="file" accept="image/*" className="hidden" onChange={async (e) => { const f = e.target.files && e.target.files[0]; e.target.value = ""; if (f) { try { const url = await onUploadImage(f); setEditing((x) => ({ ...x, coverUrl: url })); } catch (err) { /* keep editing */ } } }} />
          <p className="text-[10px] uppercase tracking-wide mb-1.5" style={{ color: "#75726C" }}>Pieces in this collection ({editing.itemIds.length} chosen)</p>
          <input value={search} onChange={(e) => setSearch(e.target.value)} placeholder="Search your pieces…" className="w-full px-3 py-2 rounded-lg border text-xs outline-none mb-2" style={OWNER_INPUT} />
          <div className="rounded-lg border overflow-y-auto mb-3" style={{ borderColor: "#2A2A2F", maxHeight: 288 }}>
            {pickable.map((x) => {
              const on = editing.itemIds.includes(String(x.id));
              return (
                <button key={x.id} onClick={() => toggle(x.id)} className="w-full flex items-center gap-3 px-3 py-2 text-left border-b" style={{ borderColor: "#232327", background: on ? "#241F13" : "transparent" }}>
                  <span className="w-5 h-5 rounded border-2 shrink-0 flex items-center justify-center" style={{ borderColor: on ? "#C9A567" : "#2A2A2F", background: on ? "#C9A567" : "transparent" }}>{on && <Check size={12} style={{ color: "#0B0B0C" }} />}</span>
                  {x.photos && x.photos[0] ? <img src={x.photos[0]} alt="" className="w-9 h-9 rounded object-cover shrink-0" /> : <span className="w-9 h-9 rounded shrink-0" style={{ background: "#0B0B0C" }} />}
                  <span className="min-w-0 flex-1">
                    <span className="block text-xs truncate" style={{ color: "#F2F0EA" }}>{productTitle(x)}</span>
                    <span className="block text-[11px]" style={{ color: "#75726C" }}>RM{x.price || "?"}{x.showOnStorefront ? "" : " · Draft (hidden until Active)"}</span>
                  </span>
                </button>
              );
            })}
            {pickable.length === 0 && <p className="text-[11px] p-3" style={{ color: "#75726C" }}>No pieces found.</p>}
          </div>
          <div className="flex items-center gap-2">
            <button onClick={save} disabled={busy || !editing.title.trim()} className="text-xs font-semibold px-4 py-2 rounded-lg" style={{ background: "#C9A567", color: "#0B0B0C", opacity: (busy || !editing.title.trim()) ? 0.5 : 1 }}>{busy ? "Saving…" : "Save collection"}</button>
            <button onClick={() => setEditing(null)} className="text-xs font-medium px-4 py-2 rounded-lg border" style={{ borderColor: "#2A2A2F", color: "#8C8A85" }}>Cancel</button>
          </div>
        </div>
      )}
    </div>
  );
}

function PublicStorefront({ slug }) {
  const [state, setState] = useState("loading"); // loading | ready | not-found | error
  const [store, setStore] = useState({ whatsapp: "", banner: "", intro: "", logo: "", sale: null, banners: [], announcement: "", faq: [] });
  const [storeItems, setStoreItems] = useState([]);
  const [lightbox, setLightbox] = useState(null);
  const [search, setSearch] = useState("");
  const [categoryFilter, setCategoryFilter] = useState("all");
  const [sort, setSort] = useState("newest");
  const [photoIdx, setPhotoIdx] = useState(0);
  const [copied, setCopied] = useState(false);
  const gridRef = useRef(null);
  const trackRef = useRef({ userId: null, enabled: false });
  const track = (type, item) => { if (trackRef.current.enabled) trackStorefrontEvent(trackRef.current.userId, type, item); };
  const [route, setRoute] = useState(parseStoreHash);
  const itemId = route.item, page = route.page, collectionId = route.collection;
  const prevHashRef = useRef("");
  const [faqOpen, setFaqOpen] = useState(-1);
  const [collections, setCollections] = useState([]);
  const saleCtx = { store: store.sale, collections };

  useEffect(() => { document.title = `@${slug}`; }, [slug]);

  useEffect(() => {
    (async () => {
      try {
        const sfRows = await publicApiRequest(`/storefronts?select=user_id,whatsapp_number,banner_url,intro,sale,banners,announcement,faq&slug=eq.${encodeURIComponent(slug)}`);
        const sf = sfRows && sfRows[0];
        if (!sf) { setState("not-found"); return; }
        const sess = loadStoredSession(); // your own signed-in visits are never counted
        trackRef.current = { userId: sf.user_id, enabled: !(sess && sess.user && sess.user.id === sf.user_id) };
        const [itemRows, brandingRows, collectionRows] = await Promise.all([
          publicApiRequest(`/items?select=${STOREFRONT_ITEM_COLUMNS}&user_id=eq.${sf.user_id}&show_on_storefront=eq.true&order=created_at.desc`),
          publicApiRequest(`/branding?select=logo_url&user_id=eq.${sf.user_id}`).catch(() => []),
          publicApiRequest(`/collections?select=id,title,description,cover_url,item_ids,position,sale&user_id=eq.${sf.user_id}&order=position.asc`).catch(() => []),
        ]);
        setStore({ whatsapp: sf.whatsapp_number || "", banner: sf.banner_url || "", intro: sf.intro || "", logo: (brandingRows && brandingRows[0]?.logo_url) || "", sale: sf.sale || null, banners: Array.isArray(sf.banners) ? sf.banners : [], announcement: sf.announcement || "", faq: Array.isArray(sf.faq) ? sf.faq : [] });
        setCollections((collectionRows || []).map(rowToCollection));
        setStoreItems((itemRows || []).map(rowToItem));
        try {
          const viewKey = `rvn_view_${slug}`;
          if (!sessionStorage.getItem(viewKey)) { sessionStorage.setItem(viewKey, "1"); track("view", null); }
        } catch (e) { /* storage blocked: skip */ }
        setState("ready");
      } catch (e) {
        setState("error");
      }
    })();
  }, [slug]);

  // keep the open product in sync with the URL hash, so product links are shareable and the back button works
  useEffect(() => {
    const sync = () => setRoute(parseStoreHash());
    window.addEventListener("popstate", sync);
    window.addEventListener("hashchange", sync);
    return () => { window.removeEventListener("popstate", sync); window.removeEventListener("hashchange", sync); };
  }, []);
  useEffect(() => { setPhotoIdx(0); }, [itemId]);

  const go = (hash) => {
    window.history.pushState(null, "", window.location.pathname + window.location.search + hash);
    setRoute(parseStoreHash());
    if (hash === "#shop") setTimeout(() => gridRef.current && gridRef.current.scrollIntoView({ behavior: "smooth", block: "start" }), 60);
    else window.scrollTo(0, 0);
  };
  const openItem = (id) => { if (!window.location.hash.startsWith("#item=")) prevHashRef.current = window.location.hash; go(`#item=${id}`); };
  const closeItem = () => go(prevHashRef.current || "");
  const copyLink = async () => {
    try { await navigator.clipboard.writeText(window.location.href); setCopied(true); setTimeout(() => setCopied(false), 1500); } catch (e) { /* ignore */ }
  };

  const waDigits = (store.whatsapp || "").replace(/[^0-9]/g, "");
  const waLink = (item) => {
    const sale = item && item.status !== "sold" ? saleFor(item, saleCtx) : null;
    const priceText = item ? (sale && sale.off > 0 ? `${fmtRM(sale.price)}, ${sale.offText}` : `RM${item.price || "?"}`) : "";
    const msg = item ? `Hi! I'm interested in "${productTitle(item)}" (${priceText}) from your Ravuno storefront.` : "Hi! I have a question about your store.";
    return `https://wa.me/${waDigits}?text=${encodeURIComponent(msg)}`;
  };

  const available = storeItems.filter((i) => i.status !== "sold");
  const featured = available.filter((i) => i.featured).slice(0, 4);
  const newArrivals = available.slice(0, 8);
  const categories = Array.from(new Set(storeItems.map((i) => i.category).filter(Boolean)));
  const q = search.trim().toLowerCase();
  const isFiltering = !!q || categoryFilter !== "all";
  let shown = storeItems.filter((i) => {
    if (categoryFilter !== "all" && i.category !== categoryFilter) return false;
    if (!q) return true;
    return [productTitle(i), i.brand, i.colour, i.itemType].some((v) => (v || "").toLowerCase().includes(q));
  });
  if (sort === "price-asc") shown = [...shown].sort((a, b) => saleFor(a, saleCtx).price - saleFor(b, saleCtx).price);
  if (sort === "price-desc") shown = [...shown].sort((a, b) => saleFor(b, saleCtx).price - saleFor(a, saleCtx).price);
  shown = [...shown.filter((i) => i.status !== "sold"), ...shown.filter((i) => i.status === "sold")];

  const selected = itemId ? storeItems.find((i) => String(i.id) === String(itemId)) : null;
  const selectedPhotos = (selected && selected.photos) || [];
  const specRows = selected ? [
    ["Size", getDisplaySize(selected)],
    ["Condition", selected.condition ? `${selected.condition}/10` : ""],
    ["Colour", [selected.colour, selected.colourCondition ? `(${selected.colourCondition}/10)` : ""].filter(Boolean).join(" ")],
    ["Brand", selected.brand],
    ["Category", [CATEGORY_LABELS[selected.category], selected.itemType].filter(Boolean).join(" · ")],
  ].filter(([, v]) => v) : [];
  const measureLines = selected
    ? (selected.category === "apparel" ? apparelMeasureParts(selected.attributes || {}, getApparelSubcategory(selected.itemType)) : getFilledAttrLines(selected))
    : [];
  const related = selected ? available.filter((i) => i.category === selected.category && i.id !== selected.id).slice(0, 4) : [];
  const saleItems = available.map((i) => [i, saleFor(i, saleCtx)]).filter(([, x]) => x.off > 0).sort((x, y) => y[1].off / (Number(y[0].price) || 1) - x[1].off / (Number(x[0].price) || 1)).map(([i]) => i);
  const itemById = new Map(storeItems.map((i) => [String(i.id), i]));
  const collectionPieces = (c) => {
    const list = (c.itemIds || []).map((id) => itemById.get(String(id))).filter(Boolean);
    return [...list.filter((i) => i.status !== "sold"), ...list.filter((i) => i.status === "sold")];
  };
  const collectionCover = (c) => c.coverUrl || ((collectionPieces(c).find((i) => i.photos && i.photos[0]) || {}).photos || [])[0] || "";
  const activeCollection = collectionId ? collections.find((c) => String(c.id) === String(collectionId)) : null;
  const slides = store.banners.length ? store.banners : store.banner ? [{ url: store.banner, link: "" }] : [];
  const slideTarget = (link) => {
    const l = String(link || "");
    let hash = null;
    if (l === "shop") hash = "#shop"; else if (l === "sale") hash = "#sale"; else if (l === "faq") hash = "#faq"; else if (l === "collections") hash = "#collections";
    else if (l.startsWith("collection:") && /^[A-Za-z0-9-]+$/.test(l.slice(11))) hash = `#collection=${l.slice(11)}`;
    else if (l.startsWith("item:") && /^[A-Za-z0-9-]+$/.test(l.slice(5))) hash = `#item=${l.slice(5)}`;
    if (hash) return { href: hash, onClick: (e) => { e.preventDefault(); if (hash.startsWith("#item=")) openItem(hash.slice(6)); else go(hash); } };
    if (l.startsWith("url:") && /^https:\/\/\S+$/.test(l.slice(4))) return { href: l.slice(4), external: true };
    return null;
  };
  const activeNav = itemId ? "" : page === "collection" ? "collections" : page;
  const navItems = [
    { key: "home", label: "Shop", hash: "" },
    ...(collections.length ? [{ key: "collections", label: "Collections", hash: "#collections" }] : []),
    { key: "sale", label: "Sale", hash: "#sale", hot: saleItems.length > 0 },
    { key: "faq", label: "FAQ", hash: "#faq" },
  ];

  useEffect(() => {
    if (state !== "ready" || !selected) return;
    try {
      const key = `rvn_p_${selected.id}`;
      if (sessionStorage.getItem(key)) return;
      sessionStorage.setItem(key, "1");
    } catch (e) { /* storage blocked: still count it */ }
    track("product", selected);
  }, [state, itemId]);

  const selSale = selected && selected.status !== "sold" ? saleFor(selected, saleCtx) : null;
  const selOnSale = !!selSale && selSale.off > 0;
  const C = { bg: "#F7F4EE", ink: "#141414", muted: "#7A756B", line: "#E4DED2", tile: "#EDE8DD" };
  const serif = { fontFamily: "'Fraunces', serif" };

  const renderCard = (item) => {
    const sold = item.status === "sold";
    const sale = sold ? null : saleFor(item, saleCtx);
    const onSale = !!sale && sale.off > 0;
    return (
      <a key={item.id} href={`#item=${item.id}`} onClick={(e) => { e.preventDefault(); openItem(item.id); }}
        className="block w-full" style={{ color: "inherit", textDecoration: "none" }}>
        <div className="relative overflow-hidden w-full" style={{ background: C.tile, aspectRatio: "1 / 1" }}>
          {item.photos && item.photos[0] ? (
            <img src={item.photos[0]} alt={productTitle(item)} decoding="async" className="absolute inset-0 w-full h-full object-cover" style={{ objectPosition: "center", opacity: sold ? 0.55 : 1 }} />
          ) : (
            <div className="absolute inset-0 flex items-center justify-center"><ImageIcon size={22} style={{ color: "#B5AE9F" }} /></div>
          )}
          {sold && <span className="absolute top-2 left-2 text-[10px] tracking-[0.14em] uppercase px-2 py-1" style={{ background: C.ink, color: C.bg }}>Sold</span>}
          {onSale && <span className="absolute top-2 right-2 text-[10px] tracking-[0.12em] uppercase px-2 py-1" style={{ background: "#B5472F", color: "#FFFFFF" }}>{onSale ? sale.badge : ""}</span>}
        </div>
        <div className="pt-3">
          <p className="text-[10px] tracking-[0.16em] uppercase mb-1 truncate" style={{ color: C.muted, minHeight: "1.5em" }}>{item.brand || "\u00A0"}</p>
          <p className="text-sm mb-1" style={{ color: C.ink, lineHeight: 1.35, minHeight: "2.7em", display: "-webkit-box", WebkitLineClamp: 2, WebkitBoxOrient: "vertical", overflow: "hidden" }}>{productTitle(item)}</p>
          {onSale ? (
            <p className="text-sm"><span style={{ color: "#B5472F" }}>{fmtRM(sale.price)}</span> <span className="ml-1" style={{ color: C.muted, textDecoration: "line-through" }}>RM{item.price}</span></p>
          ) : (
            <p className="text-sm" style={{ color: sold ? C.muted : C.ink, textDecoration: sold ? "line-through" : "none" }}>RM{item.price || "?"}</p>
          )}
        </div>
      </a>
    );
  };

  const renderCollectionTile = (c) => (
    <a key={c.id} href={`#collection=${c.id}`} onClick={(e) => { e.preventDefault(); go(`#collection=${c.id}`); }} aria-label={c.title}
      className="relative overflow-hidden block w-full" style={{ display: "block", aspectRatio: "3 / 2", background: C.tile, color: "inherit", textDecoration: "none" }}>
      {collectionCover(c) && <img src={collectionCover(c)} alt={c.title} decoding="async" className="absolute inset-0 w-full h-full object-cover" />}
    </a>
  );

  const sectionTitle = (text, aside) => (
    <div className="flex items-baseline justify-between mb-5">
      <h2 className="text-2xl" style={serif}>{text}</h2>
      {aside && <span className="text-[11px] tracking-[0.14em] uppercase" style={{ color: C.muted }}>{aside}</span>}
    </div>
  );

  return (
    <div className="min-h-screen w-full" style={{ background: C.bg, color: C.ink, fontFamily: "'Inter', sans-serif" }}>
      <style>{`@import url('https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,400;9..144,500;9..144,600&family=Inter:wght@400;500;600&display=swap');`}</style>

      <header className="sticky top-0 z-30 border-b" style={{ borderColor: C.line, background: C.bg }}>
        <div className="max-w-6xl mx-auto px-5 pt-3 flex justify-center">
          <a href={window.location.pathname} onClick={(e) => { e.preventDefault(); go(""); }} className="block" style={{ color: "inherit", textDecoration: "none" }}>
            {store.logo ? <img src={store.logo} alt={slug} className="h-9 object-contain" /> : <span className="text-lg" style={serif}>@{slug}</span>}
          </a>
        </div>
        <nav className="max-w-6xl mx-auto px-5 py-2.5 flex items-center justify-center gap-6 overflow-x-auto text-[11px] tracking-[0.16em] uppercase whitespace-nowrap">
          {navItems.map((n) => (
            <a key={n.key} href={n.hash || window.location.pathname} onClick={(e) => { e.preventDefault(); go(n.hash); }}
              style={{ color: n.hot ? "#B5472F" : "inherit", textDecoration: "none", paddingBottom: 2, borderBottom: `1px solid ${activeNav === n.key ? (n.hot ? "#B5472F" : C.ink) : "transparent"}` }}>{n.label}</a>
          ))}
          {waDigits && <a href={waLink(null)} onClick={() => track("whatsapp", null)} target="_blank" rel="noreferrer" style={{ color: "inherit", textDecoration: "none", paddingBottom: 2, borderBottom: "1px solid transparent" }}>Contact</a>}
        </nav>
      </header>

      {store.announcement && <Marquee text={store.announcement} background={C.ink} color={C.bg} />}

      {state === "loading" && <p className="text-center text-sm py-32" style={{ color: C.muted }}>Loading…</p>}
      {state === "not-found" && <p className="text-center text-sm py-32" style={{ color: C.muted }}>This storefront doesn't exist.</p>}
      {state === "error" && <p className="text-center text-sm py-32" style={{ color: C.muted }}>Something went wrong loading this page.</p>}

      {state === "ready" && !itemId && page === "home" && (
        <main>
          <BannerCarousel slides={slides} target={slideTarget} />

          {!isFiltering && (
            <section className="max-w-3xl mx-auto px-6 py-14 text-center">
              <p className="text-xl sm:text-2xl leading-relaxed mb-8" style={serif}>
                {store.intro || "Hand-picked second-hand pieces. Once they're gone, they're gone."}
              </p>
              {waDigits && (
                <a href={waLink(null)} onClick={() => track("whatsapp", null)} target="_blank" rel="noreferrer" className="inline-block text-[11px] tracking-[0.18em] uppercase px-7 py-3.5" style={{ background: C.ink, color: C.bg }}>
                  Enquire on WhatsApp
                </a>
              )}
            </section>
          )}

          {!isFiltering && newArrivals.length > 0 && (
            <section className="max-w-6xl mx-auto px-5 pb-14">
              {sectionTitle("New arrivals")}
              <div className="flex gap-4 overflow-x-auto pb-2">
                {newArrivals.map((item) => <div key={item.id} className="w-40 sm:w-56 shrink-0">{renderCard(item)}</div>)}
              </div>
            </section>
          )}

          {!isFiltering && collections.length > 0 && (
            <section className="max-w-6xl mx-auto px-5 pb-14">
              {sectionTitle("Collections")}
              <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-3">{collections.slice(0, 6).map(renderCollectionTile)}</div>
            </section>
          )}

          {!isFiltering && categories.length > 1 && (
            <section className="max-w-6xl mx-auto px-5 pb-14">
              {sectionTitle("Shop by category")}
              <div className="grid grid-cols-2 md:grid-cols-3 gap-3">
                {categories.map((cat) => {
                  const sample = storeItems.find((i) => i.category === cat && i.status !== "sold" && i.photos && i.photos[0]) || storeItems.find((i) => i.category === cat && i.photos && i.photos[0]);
                  const count = available.filter((i) => i.category === cat).length;
                  return (
                    <button key={cat} onClick={() => { setCategoryFilter(cat); setTimeout(() => gridRef.current && gridRef.current.scrollIntoView({ behavior: "smooth", block: "start" }), 50); }}
                      className="relative overflow-hidden text-left" style={{ aspectRatio: "4 / 3", background: C.tile }}>
                      {sample && <img src={sample.photos[0]} alt="" className="absolute inset-0 w-full h-full object-cover" />}
                      <div className="absolute inset-0" style={{ background: "linear-gradient(to top, rgba(20,20,20,0.65), rgba(20,20,20,0.05))" }} />
                      <div className="absolute bottom-3 left-3 right-3">
                        <p className="text-lg leading-tight" style={{ ...serif, color: "#FFFFFF" }}>{CATEGORY_LABELS[cat] || cat}</p>
                        <p className="text-[10px] tracking-[0.14em] uppercase" style={{ color: "rgba(255,255,255,0.8)" }}>{count} available</p>
                      </div>
                    </button>
                  );
                })}
              </div>
            </section>
          )}

          {!isFiltering && featured.length > 0 && (
            <section className="max-w-6xl mx-auto px-5 pb-14">
              {sectionTitle("Featured")}
              <div className="grid grid-cols-2 md:grid-cols-4 gap-x-4 gap-y-8">{featured.map(renderCard)}</div>
            </section>
          )}

          
          <section ref={gridRef} className="max-w-6xl mx-auto px-5 pb-20" style={{ scrollMarginTop: "80px" }}>
            {sectionTitle("All pieces", `${shown.length} item${shown.length === 1 ? "" : "s"}`)}
            <div className="relative mb-4">
              <Search size={15} className="absolute left-3 top-1/2 -translate-y-1/2" style={{ color: "#A8A296" }} />
              <input value={search} onChange={(e) => setSearch(e.target.value)} placeholder="Search by name, brand or colour…"
                className="w-full pl-9 pr-3 py-3 text-sm outline-none" style={{ border: `1px solid ${C.line}`, background: "#FBF9F4", color: C.ink }} />
            </div>
            <div className="flex items-center gap-2 overflow-x-auto pb-2 mb-4">
              {["all", ...categories].map((cat) => {
                const on = categoryFilter === cat;
                return (
                  <button key={cat} onClick={() => setCategoryFilter(cat)} className="shrink-0 text-[11px] tracking-[0.12em] uppercase px-3.5 py-2"
                    style={{ border: `1px solid ${on ? C.ink : C.line}`, background: on ? C.ink : "transparent", color: on ? C.bg : C.ink }}>
                    {cat === "all" ? "All" : (CATEGORY_LABELS[cat] || cat)}
                  </button>
                );
              })}
            </div>
            <div className="flex items-center justify-end gap-3 mb-6 flex-wrap">
              <select value={sort} onChange={(e) => setSort(e.target.value)} className="text-[11px] tracking-[0.12em] uppercase px-3 py-2 outline-none" style={{ border: `1px solid ${C.line}`, background: "#FBF9F4", color: C.ink }}>
                <option value="newest">Newest</option>
                <option value="price-asc">Price: low to high</option>
                <option value="price-desc">Price: high to low</option>
              </select>
            </div>

            {storeItems.length === 0 && <p className="text-center text-sm py-16" style={{ color: C.muted }}>New pieces arriving soon.</p>}
            {storeItems.length > 0 && shown.length === 0 && <p className="text-center text-sm py-16" style={{ color: C.muted }}>No pieces match your search.</p>}
            {shown.length > 0 && <div className="grid grid-cols-2 md:grid-cols-3 lg:grid-cols-4 gap-x-4 gap-y-8">{shown.map(renderCard)}</div>}
          </section>
        </main>
      )}

      {state === "ready" && itemId && (
        <main className="max-w-5xl mx-auto px-5 py-8">
          <button onClick={closeItem} className="text-[11px] tracking-[0.16em] uppercase mb-6">← Back</button>
          {!selected ? (
            <p className="text-center text-sm py-24" style={{ color: C.muted }}>This piece is no longer available.</p>
          ) : (
            <>
              <div className="grid md:grid-cols-2 gap-8 md:gap-12">
                <div>
                  <div className="relative w-full overflow-hidden flex items-center justify-center" style={{ background: C.tile, aspectRatio: "1 / 1" }}>
                    {selOnSale && <span className="absolute top-3 right-3 z-10 text-[11px] tracking-[0.14em] uppercase px-3 py-1.5" style={{ background: "#B5472F", color: "#FFFFFF" }}>{selSale.badge} sale</span>}
                    {selectedPhotos.length > 0 ? (
                      <img src={selectedPhotos[photoIdx] || selectedPhotos[0]} alt={productTitle(selected)} onClick={() => setLightbox({ photos: selectedPhotos, index: photoIdx })}
                        className="w-full h-full object-contain cursor-pointer" />
                    ) : (
                      <ImageIcon size={28} style={{ color: "#B5AE9F" }} />
                    )}
                  </div>
                  {selectedPhotos.length > 1 && (
                    <div className="flex gap-2 mt-3 overflow-x-auto">
                      {selectedPhotos.map((p, idx) => (
                        <img key={idx} src={p} alt="" onClick={() => setPhotoIdx(idx)} className="w-16 h-16 object-cover shrink-0 cursor-pointer"
                          style={{ border: `1px solid ${idx === photoIdx ? C.ink : C.line}`, opacity: idx === photoIdx ? 1 : 0.7 }} />
                      ))}
                    </div>
                  )}
                </div>

                <div>
                  {selected.brand && <p className="text-[11px] tracking-[0.18em] uppercase mb-2" style={{ color: C.muted }}>{selected.brand}</p>}
                  <h1 className="text-3xl leading-tight mb-3" style={serif}>{productTitle(selected)}</h1>
                  <p className="text-xl mb-6" style={{ textDecoration: selected.status === "sold" ? "line-through" : "none", color: selected.status === "sold" ? C.muted : C.ink }}>
                    {selOnSale ? (
                      <><span style={{ color: "#B5472F" }}>{fmtRM(selSale.price)}</span> <span className="text-base ml-1" style={{ color: C.muted, textDecoration: "line-through" }}>RM{selected.price}</span></>
                    ) : (<>RM{selected.price || "?"}</>)}
                  </p>

                  <dl className="mb-6 text-sm">
                    {specRows.map(([k, v]) => (
                      <div key={k} className="flex justify-between gap-4 py-2.5 border-b" style={{ borderColor: C.line }}>
                        <dt style={{ color: C.muted }}>{k}</dt><dd className="text-right">{v}</dd>
                      </div>
                    ))}
                  </dl>

                  {measureLines.length > 0 && (
                    <div className="mb-6">
                      <p className="text-[11px] tracking-[0.16em] uppercase mb-2" style={{ color: C.muted }}>Measurements</p>
                      <ul className="text-sm space-y-1">{measureLines.map((l, idx) => <li key={idx}>{l}</li>)}</ul>
                    </div>
                  )}

                  {selected.remark && (
                    <div className="mb-8">
                      <p className="text-[11px] tracking-[0.16em] uppercase mb-2" style={{ color: C.muted }}>Remarks</p>
                      <p className="text-sm leading-relaxed">{selected.remark}</p>
                    </div>
                  )}

                  {selected.status === "sold" ? (
                    <div className="py-4 text-center text-[11px] tracking-[0.18em] uppercase" style={{ border: `1px solid ${C.line}`, color: C.muted }}>This piece has sold</div>
                  ) : waDigits ? (
                    <a href={waLink(selected)} onClick={() => track("whatsapp", selected)} target="_blank" rel="noreferrer" className="block text-center text-[11px] tracking-[0.18em] uppercase py-4" style={{ background: C.ink, color: C.bg }}>
                      Enquire on WhatsApp
                    </a>
                  ) : (
                    <p className="text-xs text-center" style={{ color: C.muted }}>Contact details unavailable</p>
                  )}
                  <button onClick={copyLink} className="block w-full text-center text-[11px] tracking-[0.16em] uppercase mt-4" style={{ color: C.muted }}>
                    {copied ? "Link copied" : "Copy link to this piece"}
                  </button>
                </div>
              </div>

              {related.length > 0 && (
                <section className="mt-16">
                  {sectionTitle("You may also like")}
                  <div className="grid grid-cols-2 md:grid-cols-4 gap-x-4 gap-y-8">{related.map(renderCard)}</div>
                </section>
              )}
            </>
          )}
        </main>
      )}

      {state === "ready" && !itemId && page === "sale" && (
        <main className="max-w-6xl mx-auto px-5 py-10">
          <h1 className="text-3xl mb-1" style={serif}>Sale</h1>
          <p className="text-sm mb-8" style={{ color: C.muted }}>
            {saleItems.length > 0 ? `${saleItems.length} piece${saleItems.length === 1 ? "" : "s"} on sale${store.sale ? ` · ${ruleText(store.sale)} storewide` : ""}` : "Nothing on sale right now. Check back soon."}
          </p>
          {saleItems.length > 0 && <div className="grid grid-cols-2 md:grid-cols-3 lg:grid-cols-4 gap-x-4 gap-y-8">{saleItems.map(renderCard)}</div>}
        </main>
      )}

      {state === "ready" && !itemId && page === "faq" && (
        <main className="max-w-2xl mx-auto px-5 py-10">
          <h1 className="text-3xl mb-6" style={serif}>FAQ</h1>
          {store.faq.length === 0 && <p className="text-sm" style={{ color: C.muted }}>No questions yet.</p>}
          {store.faq.map((f, i) => (
            <div key={i} className="border-b" style={{ borderColor: C.line }}>
              <button onClick={() => setFaqOpen(faqOpen === i ? -1 : i)} aria-expanded={faqOpen === i} className="w-full flex items-center justify-between gap-4 py-4 text-left">
                <span className="text-sm font-medium">{f.q}</span>
                <span aria-hidden="true" className="text-lg leading-none" style={{ color: C.muted }}>{faqOpen === i ? "–" : "+"}</span>
              </button>
              {faqOpen === i && <p className="pb-5 text-sm leading-relaxed" style={{ color: C.muted, whiteSpace: "pre-line" }}>{f.a}</p>}
            </div>
          ))}
          {waDigits && (
            <a href={waLink(null)} onClick={() => track("whatsapp", null)} target="_blank" rel="noreferrer" className="block text-center text-[11px] tracking-[0.18em] uppercase py-4 mt-10" style={{ background: C.ink, color: C.bg }}>
              Still have a question? Message us
            </a>
          )}
        </main>
      )}

      {state === "ready" && !itemId && page === "collections" && (
        <main className="max-w-6xl mx-auto px-5 py-10">
          <h1 className="text-3xl mb-6" style={serif}>Collections</h1>
          {collections.length === 0 && <p className="text-sm" style={{ color: C.muted }}>No collections yet.</p>}
          <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-3">{collections.map(renderCollectionTile)}</div>
        </main>
      )}

      {state === "ready" && !itemId && page === "collection" && (
        <main>
          {!activeCollection ? (
            <div className="max-w-3xl mx-auto px-6 py-24 text-center">
              <p className="text-sm mb-4" style={{ color: C.muted }}>This collection isn't available.</p>
              <button onClick={() => go("#collections")} className="text-[11px] tracking-[0.16em] uppercase">← All collections</button>
            </div>
          ) : (
            <>
              {activeCollection.coverUrl && (
                <div className="w-full mx-auto" style={{ maxWidth: 1200 }}>
                  <img src={activeCollection.coverUrl} alt="" className="w-full block object-cover" style={{ aspectRatio: "3 / 2", objectPosition: "center" }} />
                </div>
              )}
              <section className="max-w-3xl mx-auto px-6 py-12 text-center">
                <p className="text-[11px] tracking-[0.2em] uppercase mb-3" style={{ color: C.muted }}>Collection</p>
                <h1 className="text-4xl mb-4" style={serif}>{activeCollection.title}</h1>
                {activeCollection.description && <p className="text-base leading-relaxed" style={{ color: C.muted, whiteSpace: "pre-line" }}>{activeCollection.description}</p>}
              </section>
              <section className="max-w-6xl mx-auto px-5 pb-20">
                {collectionPieces(activeCollection).length === 0
                  ? <p className="text-center text-sm py-12" style={{ color: C.muted }}>No pieces in this collection right now.</p>
                  : <div className="grid grid-cols-2 md:grid-cols-3 lg:grid-cols-4 gap-x-4 gap-y-8">{collectionPieces(activeCollection).map(renderCard)}</div>}
                <div className="text-center mt-12"><button onClick={() => go("#collections")} className="text-[11px] tracking-[0.16em] uppercase" style={{ color: C.muted }}>← All collections</button></div>
              </section>
            </>
          )}
        </main>
      )}

      <footer className="border-t py-12 text-center" style={{ borderColor: C.line }}>
        {waDigits && (
          <a href={waLink(null)} onClick={() => track("whatsapp", null)} target="_blank" rel="noreferrer" className="block text-[11px] tracking-[0.16em] uppercase mb-4">Message us on WhatsApp</a>
        )}
        <p className="text-[11px]" style={{ color: C.muted }}>Powered by Ravuno</p>
      </footer>

      {lightbox && (
        <Lightbox photos={lightbox.photos} index={lightbox.index} onIndexChange={(idx) => { setLightbox((lb) => ({ ...lb, index: idx })); setPhotoIdx(idx); }} onClose={() => setLightbox(null)} />
      )}
    </div>
  );
}

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
    const compressed = await Promise.all(files.map((f) => compressImage(f, 1200, 0.85)));
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

// ---------- storefront analytics: anonymous visit tracking (no names, no contact details) ----------
function getVisitorId() {
  try {
    let id = localStorage.getItem("rvn_vid");
    if (!id) {
      id = (typeof crypto !== "undefined" && crypto.randomUUID) ? crypto.randomUUID() : `${Date.now()}-${Math.random().toString(36).slice(2)}`;
      localStorage.setItem("rvn_vid", id);
    }
    return id;
  } catch (e) { return null; }
}
function detectTrafficSource() {
  try {
    const params = new URLSearchParams(window.location.search);
    const tagged = (params.get("src") || params.get("utm_source") || "").toLowerCase().replace(/[^a-z0-9_-]/g, "").slice(0, 30);
    if (tagged) return tagged;
    const ua = navigator.userAgent || "";
    if (/Barcelona/i.test(ua)) return "threads"; // Threads' in-app browser
    if (/Instagram/i.test(ua)) return "instagram";
    if (/FBAN|FBAV|FB_IAB/i.test(ua)) return "facebook";
    if (/WhatsApp/i.test(ua)) return "whatsapp";
    if (/TikTok|musical_ly|Bytedance/i.test(ua)) return "tiktok";
    const ref = document.referrer ? new URL(document.referrer).hostname.replace(/^www\./, "") : "";
    if (ref && ref !== window.location.hostname) {
      if (/threads\.(net|com)$/.test(ref)) return "threads";
      if (/instagram\.com$/.test(ref)) return "instagram";
      if (/(facebook|fb)\.com$/.test(ref)) return "facebook";
      if (/(^|\.)google\./.test(ref)) return "google";
      if (/tiktok\.com$/.test(ref)) return "tiktok";
      if (/carousell\./.test(ref)) return "carousell";
      return ref.slice(0, 40);
    }
  } catch (e) { /* fall through to direct */ }
  return "direct";
}
function detectDevice() {
  const ua = (typeof navigator !== "undefined" && navigator.userAgent) || "";
  if (/iPhone|iPad|iPod/i.test(ua)) return "ios";
  if (/Android/i.test(ua)) return "android";
  return "desktop";
}
function trackStorefrontEvent(storeUserId, type, item) {
  try {
    fetch(`${REST_URL}/storefront_events`, {
      method: "POST", keepalive: true,
      headers: { "Content-Type": "application/json", apikey: SUPABASE_ANON_KEY, Prefer: "return=minimal" },
      body: JSON.stringify({
        store_user_id: storeUserId, event_type: type,
        item_id: item ? String(item.id) : null, item_title: item ? productTitle(item).slice(0, 200) : null,
        source: detectTrafficSource(), device: detectDevice(), visitor_id: getVisitorId(),
      }),
    }).catch(() => {});
  } catch (e) { /* tracking must never break the page */ }
}

// ANALYTICS-LOGIC-START
const ANALYTICS_MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
const ANALYTICS_DOW = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
function periodRange(mode, cursor) {
  if (mode === "year") return { start: new Date(cursor.year, 0, 1), end: new Date(cursor.year + 1, 0, 1) };
  return { start: new Date(cursor.year, cursor.month, 1), end: new Date(cursor.year, cursor.month + 1, 1) };
}
function parseDay(str) {
  if (!str) return null;
  const m = String(str).slice(0, 10).match(/^(\d{4})-(\d{2})-(\d{2})$/);
  return m ? new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3])) : null;
}
function buildAnalytics({ events, txns, items, purchases, mode, cursor }) {
  const { start, end } = periodRange(mode, cursor);
  const inRange = (d) => !!d && !isNaN(d.getTime()) && d >= start && d < end;
  const labels = mode === "year" ? ANALYTICS_MONTHS : Array.from({ length: new Date(cursor.year, cursor.month + 1, 0).getDate() }, (_, i) => String(i + 1));
  const rows = labels.map((label) => ({ label, visits: 0, products: 0, whatsapp: 0, revenue: 0, profit: 0, sold: 0, added: 0 }));
  const idx = (d) => (mode === "year" ? d.getMonth() : d.getDate() - 1);
  const pretty = (s) => (s ? s.charAt(0).toUpperCase() + s.slice(1) : "Direct");
  const deviceNames = { ios: "iPhone / iPad", android: "Android", desktop: "Desktop" };

  // storefront traffic
  const visitors = new Set(), openers = new Set(), tappers = new Set();
  const sources = {}, devices = {}, products = {};
  const dow = Array(7).fill(0), hours = Array(24).fill(0);
  let visits = 0, productViews = 0, whatsapp = 0;
  const productEntry = (e) => (products[e.item_id] = products[e.item_id] || { name: e.item_title || "Item", views: 0, taps: 0 });
  for (const e of events || []) {
    const d = new Date(e.created_at);
    if (!inRange(d)) continue;
    const row = rows[idx(d)];
    if (e.event_type === "view") {
      visits++; row.visits++;
      if (e.visitor_id) visitors.add(e.visitor_id);
      const src = e.source || "direct"; sources[src] = (sources[src] || 0) + 1;
      const dev = e.device || "other"; devices[dev] = (devices[dev] || 0) + 1;
      dow[d.getDay()]++; hours[d.getHours()]++;
    } else if (e.event_type === "product") {
      productViews++; row.products++;
      if (e.visitor_id) openers.add(e.visitor_id);
      if (e.item_id) productEntry(e).views++;
    } else if (e.event_type === "whatsapp") {
      whatsapp++; row.whatsapp++;
      if (e.visitor_id) tappers.add(e.visitor_id);
      if (e.item_id) productEntry(e).taps++;
    }
  }

  // business performance
  const itemsBySku = new Map(), itemsByTitle = new Map();
  for (const it of items || []) {
    const k = String(it.sku || "").trim().toLowerCase();
    if (k && !itemsBySku.has(k)) itemsBySku.set(k, it);
    const t = productTitle(it);
    if (t && !itemsByTitle.has(t)) itemsByTitle.set(t, it);
  }
  let revenue = 0, profit = 0, units = 0, orders = 0, postage = 0;
  const cats = {}, brands = {};
  for (const t of txns || []) {
    const d = new Date(t.timestamp);
    if (!inRange(d)) continue;
    const row = rows[idx(d)];
    orders++; postage += Number(t.postage || 0); revenue += Number(t.total || 0); row.revenue += Number(t.total || 0);
    for (const l of t.lines || []) {
      const fp = Number(l.finalPrice || 0), cost = Number(l.cost || 0);
      profit += fp - cost; row.profit += fp - cost; units++; row.sold++;
      const it = itemsBySku.get(String(l.sku || "").trim().toLowerCase()) || itemsByTitle.get(l.title);
      const cat = (it && it.category) || "other";
      const brand = (it && it.brand && String(it.brand).trim()) || "Unbranded";
      const c = (cats[cat] = cats[cat] || { revenue: 0, units: 0 }); c.revenue += fp; c.units++;
      const b = (brands[brand] = brands[brand] || { revenue: 0, units: 0 }); b.revenue += fp; b.units++;
    }
  }
  let added = 0;
  for (const it of items || []) {
    const d = parseDay(it.addedOn);
    if (inRange(d)) { added++; rows[idx(d)].added++; }
  }
  let spend = 0;
  for (const p of purchases || []) {
    const d = parseDay(p.purchasedOn);
    if (inRange(d)) spend += Number(p.cost || 0);
  }
  const dayCounts = [];
  for (const it of items || []) {
    if (it.status !== "sold") continue;
    const sold = parseDay(it.soldOn), listed = parseDay(it.listedOn) || parseDay(it.addedOn);
    if (inRange(sold) && listed) dayCounts.push(Math.max(0, Math.round((sold - listed) / 86400000)));
  }
  const r2 = (n) => Math.round(n * 100) / 100;
  rows.forEach((r) => { r.revenue = r2(r.revenue); r.profit = r2(r.profit); });
  const byValue = (obj) => Object.entries(obj).map(([name, value]) => ({ name, value })).sort((a, b) => b.value - a.value);
  const pct = (n, of) => (of ? Math.round((n / of) * 100) : 0);

  return {
    rows,
    hasEvents: visits + productViews + whatsapp > 0,
    totals: {
      visits, uniqueVisitors: visitors.size, productViews, whatsapp,
      enquiryRate: Math.min(100, pct(tappers.size, visitors.size)),
      revenue: r2(revenue), profit: r2(profit), margin: revenue ? Math.round((profit / revenue) * 100) : 0,
      units, orders, avgSale: units ? r2(revenue / units) : 0, spend: r2(spend), added, postage: r2(postage),
      avgDaysToSell: dayCounts.length ? Math.round(dayCounts.reduce((a, b) => a + b, 0) / dayCounts.length) : null,
    },
    funnel: [
      { name: "Visitors", value: visitors.size, label: `${visitors.size}` },
      { name: "Opened a piece", value: openers.size, label: `${openers.size} · ${pct(openers.size, visitors.size)}%` },
      { name: "Tapped WhatsApp", value: tappers.size, label: `${tappers.size} · ${pct(tappers.size, visitors.size)}%` },
    ],
    sources: byValue(sources).map((s) => ({ name: pretty(s.name), value: s.value, label: `${s.value} · ${pct(s.value, visits)}%` })),
    devices: byValue(devices).map((s) => ({ name: deviceNames[s.name] || pretty(s.name), value: s.value, label: `${s.value} · ${pct(s.value, visits)}%` })),
    dow: ANALYTICS_DOW.map((label, i) => ({ label, count: dow[i] })),
    hours: hours.map((count, h) => ({ label: String(h), count })),
    topProducts: Object.values(products).sort((a, b) => (b.views + b.taps * 3) - (a.views + a.taps * 3)).slice(0, 8),
    categories: Object.entries(cats).map(([name, v]) => ({ name, revenue: r2(v.revenue), units: v.units })).sort((a, b) => b.revenue - a.revenue),
    brands: Object.entries(brands).map(([name, v]) => ({ name, revenue: r2(v.revenue), units: v.units })).sort((a, b) => b.revenue - a.revenue).slice(0, 8),
  };
}
// ANALYTICS-LOGIC-END

const CHART_AXIS = { tick: { fontSize: 11, fill: "#75726C" }, axisLine: { stroke: "#2A2A2F" }, tickLine: false };
const CHART_TOOLTIP = {
  contentStyle: { background: "#1E1E22", border: "1px solid #2A2A2F", borderRadius: 8, fontSize: 12 },
  labelStyle: { color: "#FFFFFF" }, itemStyle: { color: "#FFFFFF" }, cursor: { fill: "rgba(255,255,255,0.04)" },
};
const TRACKED_CHANNELS = [["threads", "Threads"], ["instagram", "Instagram"], ["facebook", "Facebook"], ["whatsapp", "WhatsApp"], ["tiktok", "TikTok"], ["carousell", "Carousell"]];

function AnalyticsCard({ title, hint, children }) {
  return (
    <div className="rounded-xl border p-5" style={{ background: "#17171A", borderColor: "#2A2A2F" }}>
      <h2 className="text-sm font-semibold" style={{ color: "#F2F0EA", fontFamily: "'Fraunces', serif" }}>{title}</h2>
      {hint && <p className="text-[11px] mt-1" style={{ color: "#75726C" }}>{hint}</p>}
      <div className="mt-4">{children}</div>
    </div>
  );
}
function BarList({ rows, color = "#C9A567", empty = "Nothing yet for this period." }) {
  if (!rows || rows.length === 0) return <p className="text-xs" style={{ color: "#75726C" }}>{empty}</p>;
  const max = Math.max(1, ...rows.map((r) => r.value));
  return (
    <div className="space-y-3">
      {rows.map((r) => (
        <div key={r.name}>
          <div className="flex items-baseline justify-between gap-3 mb-1">
            <span className="text-xs truncate" style={{ color: "#F2F0EA" }}>{r.name}</span>
            <span className="text-xs shrink-0" style={{ color: "#8C8A85", fontFamily: "'IBM Plex Mono', monospace" }}>{r.label != null ? r.label : r.value}</span>
          </div>
          <div className="h-1.5 rounded-full" style={{ background: "#232327" }}>
            <div className="h-1.5 rounded-full" style={{ width: `${Math.max(2, (r.value / max) * 100)}%`, background: color }} />
          </div>
        </div>
      ))}
    </div>
  );
}
function SectionLabel({ children }) {
  return <p className="text-[10px] uppercase tracking-[0.22em] pt-2" style={{ color: "#C9A567", fontFamily: "'IBM Plex Mono', monospace" }}>{children}</p>;
}

function AnalyticsTab({ session, updateSession, storefront, items, txns, purchases }) {
  const today = new Date();
  const [mode, setMode] = useState("month"); // month | year
  const [cursor, setCursor] = useState({ year: today.getFullYear(), month: today.getMonth() });
  const [events, setEvents] = useState([]);
  const [loading, setLoading] = useState(false);
  const [loadError, setLoadError] = useState(false);
  const [copiedKey, setCopiedKey] = useState("");

  const range = periodRange(mode, cursor);
  const startMs = range.start.getTime(), endMs = range.end.getTime();
  const userId = session && session.user ? session.user.id : null;

  useEffect(() => {
    if (!userId) return;
    let cancelled = false;
    (async () => {
      setLoading(true); setLoadError(false);
      try {
        const all = [], PAGE = 1000;
        for (let page = 0; page < 30; page++) {
          const rows = await apiRequest(
            `/storefront_events?select=event_type,item_id,item_title,source,device,visitor_id,created_at&store_user_id=eq.${userId}&created_at=gte.${new Date(startMs).toISOString()}&created_at=lt.${new Date(endMs).toISOString()}&order=created_at.asc&limit=${PAGE}&offset=${page * PAGE}`,
            { method: "GET" }, session, updateSession);
          all.push(...(rows || []));
          if (!rows || rows.length < PAGE) break;
        }
        if (!cancelled) setEvents(all);
      } catch (e) {
        if (!cancelled) setLoadError(true);
      } finally {
        if (!cancelled) setLoading(false);
      }
    })();
    return () => { cancelled = true; };
  }, [userId, startMs, endMs]);

  const data = useMemo(() => buildAnalytics({ events, txns, items, purchases, mode, cursor }), [events, txns, items, purchases, mode, cursor]);
  const t = data.totals;
  const step = (dir) => setCursor((c) => {
    if (mode === "year") return { ...c, year: c.year + dir };
    const d = new Date(c.year, c.month + dir, 1);
    return { year: d.getFullYear(), month: d.getMonth() };
  });
  const periodLabel = mode === "year" ? String(cursor.year) : new Date(cursor.year, cursor.month, 1).toLocaleDateString(undefined, { month: "long", year: "numeric" });
  const origin = typeof window !== "undefined" ? window.location.origin : "";
  const copyTracked = async (key) => {
    try {
      await navigator.clipboard.writeText(`${origin}/store/${storefront.slug}?src=${key}`);
      setCopiedKey(key); setTimeout(() => setCopiedKey(""), 1500);
    } catch (e) { /* ignore */ }
  };
  const catRows = data.categories.map((c) => ({ name: CATEGORY_LABELS[c.name] || c.name, value: c.revenue, label: `RM${c.revenue} · ${c.units} sold` }));
  const brandRows = data.brands.map((b) => ({ name: b.name, value: b.revenue, label: `RM${b.revenue} · ${b.units} sold` }));
  const productRows = data.topProducts.map((p) => ({ name: p.name, value: p.views + p.taps * 3, label: `${p.views} views · ${p.taps} taps` }));

  return (
    <div className="max-w-3xl mx-auto space-y-5">
      <div className="flex items-center justify-between gap-3 flex-wrap">
        <div className="flex rounded-lg border overflow-hidden" style={{ borderColor: "#2A2A2F" }}>
          {["month", "year"].map((m) => (
            <button key={m} onClick={() => setMode(m)} className="text-xs font-semibold px-4 py-2"
              style={{ background: mode === m ? "#C9A567" : "transparent", color: mode === m ? "#0B0B0C" : "#8C8A85" }}>
              {m === "month" ? "Monthly" : "Yearly"}
            </button>
          ))}
        </div>
        <div className="flex items-center gap-2">
          <button onClick={() => step(-1)} className="p-1.5 rounded-lg border" style={{ borderColor: "#2A2A2F" }}><ChevronLeft size={14} style={{ color: "#8C8A85" }} /></button>
          <span className="text-xs font-medium min-w-[110px] text-center" style={{ color: "#F2F0EA", fontFamily: "'IBM Plex Mono', monospace" }}>{periodLabel}</span>
          <button onClick={() => step(1)} className="p-1.5 rounded-lg border" style={{ borderColor: "#2A2A2F" }}><ChevronRight size={14} style={{ color: "#8C8A85" }} /></button>
        </div>
      </div>

      <SectionLabel>Storefront · marketing</SectionLabel>

      {!storefront && (
        <AnalyticsCard title="No storefront yet" hint="Create your storefront in the Storefront tab, then visits show up here." />
      )}
      {loadError && <p className="text-xs px-3 py-2 rounded-lg" style={{ background: "#2A1715", color: "#C9695E" }}>Couldn't load visit data — try again in a moment.</p>}
      {storefront && !loading && !loadError && !data.hasEvents && (
        <p className="text-xs px-3 py-2.5 rounded-lg leading-relaxed" style={{ background: "#241F13", color: "#C9A567" }}>
          No visits recorded for {periodLabel} yet. Tracking only starts once this version is live — earlier visits weren't recorded.
          Use the tracked links below in your posts to see which one brings people in.
        </p>
      )}

      <div className="grid grid-cols-2 gap-3">
        <StatCard icon={Eye} label="Visits" value={t.visits} tone="#C9A567" />
        <StatCard icon={Users} label="Visitors" value={t.uniqueVisitors} tone="#4A4A52" />
        <StatCard icon={MessageCircle} label="WhatsApp taps" value={t.whatsapp} tone="#6FA287" />
        <StatCard icon={TrendingUp} label="Enquiry rate" value={`${t.enquiryRate}%`} tone="#8C8A85" />
      </div>

      <AnalyticsCard title="Visits over time" hint={mode === "month" ? "Per day" : "Per month"}>
        <div style={{ width: "100%", height: 220 }}>
          <ResponsiveContainer>
            <LineChart data={data.rows} margin={{ top: 5, right: 10, bottom: 0, left: 0 }}>
              <CartesianGrid stroke="#232327" vertical={false} />
              <XAxis dataKey="label" {...CHART_AXIS} interval={mode === "month" ? 2 : 0} />
              <YAxis {...CHART_AXIS} axisLine={false} width={30} allowDecimals={false} />
              <Tooltip {...CHART_TOOLTIP} />
              <Legend wrapperStyle={{ fontSize: 11, color: "#8C8A85" }} />
              <Line type="monotone" dataKey="visits" name="Visits" stroke="#C9A567" strokeWidth={2} dot={false} />
              <Line type="monotone" dataKey="products" name="Product views" stroke="#F2F0EA" strokeWidth={1.5} dot={false} />
              <Line type="monotone" dataKey="whatsapp" name="WhatsApp taps" stroke="#6FA287" strokeWidth={2} dot={false} />
            </LineChart>
          </ResponsiveContainer>
        </div>
      </AnalyticsCard>

      <AnalyticsCard title="Visitor funnel" hint="Of the people who visited, how many opened a piece and how many reached out.">
        <BarList rows={data.funnel} />
      </AnalyticsCard>

      <AnalyticsCard title="Where visitors come from" hint="Detected automatically from Threads, Instagram, Facebook and so on, or from a tracked link.">
        <BarList rows={data.sources} />
      </AnalyticsCard>

      {storefront && (
        <AnalyticsCard title="Tracked links" hint="Use a different link in each post or channel. Visits through it are labelled with that source, even if the app can't detect it.">
          <div className="flex flex-wrap gap-2">
            {TRACKED_CHANNELS.map(([key, name]) => (
              <button key={key} onClick={() => copyTracked(key)} className="flex items-center gap-1.5 text-xs font-medium px-3 py-2 rounded-lg border"
                style={{ borderColor: copiedKey === key ? "#C9A567" : "#2A2A2F", color: copiedKey === key ? "#C9A567" : "#F2F0EA" }}>
                {copiedKey === key ? <Check size={12} /> : <Link2 size={12} />} {copiedKey === key ? "Copied" : name}
              </button>
            ))}
          </div>
          <p className="text-[11px] mt-3 break-all" style={{ color: "#75726C", fontFamily: "'IBM Plex Mono', monospace" }}>{origin}/store/{storefront.slug}?src=threads</p>
        </AnalyticsCard>
      )}

      <AnalyticsCard title="Most viewed pieces" hint="Ranked by views, with WhatsApp taps counted extra.">
        <BarList rows={productRows} color="#F2F0EA" />
      </AnalyticsCard>

      <AnalyticsCard title="Best day to post" hint="Visits by day of the week.">
        <div style={{ width: "100%", height: 170 }}>
          <ResponsiveContainer>
            <BarChart data={data.dow} margin={{ top: 5, right: 10, bottom: 0, left: 0 }}>
              <CartesianGrid stroke="#232327" vertical={false} />
              <XAxis dataKey="label" {...CHART_AXIS} />
              <YAxis {...CHART_AXIS} axisLine={false} width={30} allowDecimals={false} />
              <Tooltip {...CHART_TOOLTIP} />
              <Bar dataKey="count" name="Visits" fill="#C9A567" radius={[3, 3, 0, 0]} />
            </BarChart>
          </ResponsiveContainer>
        </div>
      </AnalyticsCard>

      <AnalyticsCard title="Best time to post" hint="Visits by hour of the day, in your local time.">
        <div style={{ width: "100%", height: 170 }}>
          <ResponsiveContainer>
            <BarChart data={data.hours} margin={{ top: 5, right: 10, bottom: 0, left: 0 }}>
              <CartesianGrid stroke="#232327" vertical={false} />
              <XAxis dataKey="label" {...CHART_AXIS} interval={2} />
              <YAxis {...CHART_AXIS} axisLine={false} width={30} allowDecimals={false} />
              <Tooltip {...CHART_TOOLTIP} />
              <Bar dataKey="count" name="Visits" fill="#C9A567" radius={[3, 3, 0, 0]} />
            </BarChart>
          </ResponsiveContainer>
        </div>
      </AnalyticsCard>

      <AnalyticsCard title="Devices">
        <BarList rows={data.devices} color="#8C8A85" />
      </AnalyticsCard>

      <SectionLabel>Business performance</SectionLabel>

      <div className="grid grid-cols-2 gap-3">
        <StatCard icon={DollarSign} label="Revenue" value={`RM${t.revenue}`} tone="#C9A567" />
        <StatCard icon={TrendingUp} label="Profit" value={`RM${t.profit}`} tone="#6FA287" />
        <StatCard icon={Percent} label="Margin" value={`${t.margin}%`} tone="#4A4A52" />
        <StatCard icon={ShoppingBag} label="Units sold" value={t.units} tone="#8C8A85" />
        <StatCard icon={Tag} label="Avg sale price" value={`RM${t.avgSale}`} tone="#4A4A52" />
        <StatCard icon={Clock} label="Avg days to sell" value={t.avgDaysToSell == null ? "—" : t.avgDaysToSell} tone="#C9695E" />
      </div>

      <AnalyticsCard title="Revenue & profit" hint={`${mode === "month" ? "Per day" : "Per month"} · excludes RM${t.postage} postage collected`}>
        <div style={{ width: "100%", height: 220 }}>
          <ResponsiveContainer>
            <BarChart data={data.rows} margin={{ top: 5, right: 10, bottom: 0, left: 0 }}>
              <CartesianGrid stroke="#232327" vertical={false} />
              <XAxis dataKey="label" {...CHART_AXIS} interval={mode === "month" ? 2 : 0} />
              <YAxis {...CHART_AXIS} axisLine={false} width={40} />
              <Tooltip {...CHART_TOOLTIP} formatter={(v, n) => [`RM${v}`, n]} />
              <Legend wrapperStyle={{ fontSize: 11, color: "#8C8A85" }} />
              <Bar dataKey="revenue" name="Revenue" fill="#C9A567" radius={[3, 3, 0, 0]} />
              <Bar dataKey="profit" name="Profit" fill="#6FA287" radius={[3, 3, 0, 0]} />
            </BarChart>
          </ResponsiveContainer>
        </div>
      </AnalyticsCard>

      <AnalyticsCard title="Stock added vs sold" hint={`You spent RM${t.spend} on purchases in this period.`}>
        <div style={{ width: "100%", height: 200 }}>
          <ResponsiveContainer>
            <BarChart data={data.rows} margin={{ top: 5, right: 10, bottom: 0, left: 0 }}>
              <CartesianGrid stroke="#232327" vertical={false} />
              <XAxis dataKey="label" {...CHART_AXIS} interval={mode === "month" ? 2 : 0} />
              <YAxis {...CHART_AXIS} axisLine={false} width={30} allowDecimals={false} />
              <Tooltip {...CHART_TOOLTIP} />
              <Legend wrapperStyle={{ fontSize: 11, color: "#8C8A85" }} />
              <Bar dataKey="added" name="Added to stock" fill="#8C8A85" radius={[3, 3, 0, 0]} />
              <Bar dataKey="sold" name="Sold" fill="#C9A567" radius={[3, 3, 0, 0]} />
            </BarChart>
          </ResponsiveContainer>
        </div>
      </AnalyticsCard>

      <AnalyticsCard title="What sells: by category" hint="Revenue from sales in this period.">
        <BarList rows={catRows} />
      </AnalyticsCard>

      <AnalyticsCard title="What sells: top brands">
        <BarList rows={brandRows} color="#F2F0EA" />
      </AnalyticsCard>

      <p className="text-[11px] text-center pb-4" style={{ color: "#75726C" }}>
        Visits from your own signed-in browser aren't counted. Counts are anonymous and can be inflated by anyone deliberately refreshing, so treat them as a guide.
      </p>
    </div>
  );
}

function StorefrontTab({ storefront, onSaveStorefront, storefrontSaving, storefrontError, collections, onSaveFields, onSaveBanners, onUploadImage, onSaveCollection, onDeleteCollection, logoUrl, onLogoUpload, onLogoRemove, logoUploading, items, onSaveStoreSale, onSetItemSale, onSetCollectionSale }) {
  const [slugDraft, setSlugDraft] = useState(storefront?.slug || "");
  const [whatsappDraft, setWhatsappDraft] = useState(storefront?.whatsappNumber || "");
  const [introDraft, setIntroDraft] = useState(storefront?.intro || "");
  useEffect(() => { setSlugDraft(storefront?.slug || ""); setWhatsappDraft(storefront?.whatsappNumber || ""); setIntroDraft(storefront?.intro || ""); }, [storefront?.slug, storefront?.whatsappNumber, storefront?.intro]);
  const logoInputRef = useRef(null);
  const [storeRule, setStoreRule] = useState(ruleToDraft(storefront && storefront.sale));
  const [pieceRule, setPieceRule] = useState({ kind: "percent", value: "" });
  const [collRule, setCollRule] = useState({ kind: "price", value: "" });
  const [pickId, setPickId] = useState("");
  const [pickCollection, setPickCollection] = useState("");
  useEffect(() => { setStoreRule(ruleToDraft(storefront && storefront.sale)); }, [JSON.stringify(storefront && storefront.sale)]);
  const pickable = (items || []).filter((i) => i.status !== "sold" && !i.sale).sort((a, b) => productTitle(a).localeCompare(productTitle(b)));
  const onSaleItems = (items || []).filter((i) => i.sale && i.status !== "sold");
  const collectionsOnSale = (collections || []).filter((c) => c.sale);

  return (
    <div className="max-w-2xl mx-auto space-y-6">
      <div className="rounded-xl border p-5" style={{ background: "#17171A", borderColor: "#2A2A2F" }}>
        <h2 className="text-sm font-semibold mb-2" style={{ color: "#F2F0EA", fontFamily: "'Fraunces', serif" }}>Public storefront</h2>
        <p className="text-xs mb-3" style={{ color: "#8C8A85" }}>
          A browse-only page showing your Active items — buyers tap through to message you on WhatsApp. No checkout, no payments.
        </p>
        <label className="block mb-3">
          <span className="block text-[10px] mb-1 uppercase tracking-wide" style={{ color: "#75726C" }}>Your link</span>
          <div className="flex items-center rounded-lg border overflow-hidden" style={{ borderColor: "#2A2A2F", background: "#0B0B0C" }}>
            <span className="pl-3 text-xs" style={{ color: "#75726C", fontFamily: "'IBM Plex Mono', monospace" }}>ravuno.netlify.app/store/</span>
            <input value={slugDraft} onChange={(e) => setSlugDraft(e.target.value)} placeholder="iselleverything"
              className="flex-1 min-w-0 px-1 py-2 text-xs outline-none bg-transparent" style={{ color: "#F2F0EA", fontFamily: "'IBM Plex Mono', monospace" }} />
          </div>
        </label>
        <label className="block mb-3">
          <span className="block text-[10px] mb-1 uppercase tracking-wide" style={{ color: "#75726C" }}>WhatsApp number (with country code)</span>
          <input value={whatsappDraft} onChange={(e) => setWhatsappDraft(e.target.value)} placeholder="60123456789"
            className="w-full px-3 py-2 rounded-lg border text-sm outline-none" style={{ borderColor: "#2A2A2F", background: "#0B0B0C", color: "#F2F0EA" }} />
        </label>
        <label className="block mb-3">
          <span className="block text-[10px] mb-1 uppercase tracking-wide" style={{ color: "#75726C" }}>Store introduction (shown on your storefront homepage)</span>
          <textarea value={introDraft} onChange={(e) => setIntroDraft(e.target.value)} rows={3} maxLength={280}
            placeholder="Hand-picked second-hand pieces. Once they're gone, they're gone."
            className="w-full px-3 py-2 rounded-lg border text-sm outline-none resize-y" style={{ borderColor: "#2A2A2F", background: "#0B0B0C", color: "#F2F0EA" }} />
        </label>
        {storefrontError && <p className="text-xs mb-2" style={{ color: "#C9695E" }}>{storefrontError}</p>}
        <div className="flex items-center gap-3">
          <button onClick={() => onSaveStorefront(slugDraft, whatsappDraft, introDraft)} disabled={storefrontSaving || !slugDraft.trim()}
            className="text-xs font-semibold px-4 py-2 rounded-lg" style={{ background: "#C9A567", color: "#0B0B0C", opacity: (storefrontSaving || !slugDraft.trim()) ? 0.5 : 1 }}>
            {storefrontSaving ? "Saving…" : storefront ? "Update storefront" : "Create storefront"}
          </button>
          {storefront && (
            <a href={`/store/${storefront.slug}`} target="_blank" rel="noreferrer" className="text-xs font-medium" style={{ color: "#C9A567" }}>
              View your storefront →
            </a>
          )}
        </div>
        <p className="text-[11px] mt-2" style={{ color: "#75726C" }}>
          Control which items show up per-item from each item's detail view in Stock Ledger (Active / Draft toggle). Your
          logo (below) is shown at the top of your storefront automatically.
        </p>
      </div>

      {storefront && (
        <>
          <AnnouncementCard storefront={storefront} onSave={onSaveFields} />
          <BannerSlidesCard storefront={storefront} items={items} collections={collections} onSaveBanners={onSaveBanners} onUploadImage={onUploadImage} />
          <CollectionsCard collections={collections} items={items} onSave={onSaveCollection} onDelete={onDeleteCollection} onUploadImage={onUploadImage} />
          <FaqCard storefront={storefront} onSave={onSaveFields} />
        </>
      )}

      {storefront && (
        <div className="rounded-xl border p-5" style={{ background: "#17171A", borderColor: "#2A2A2F" }}>
          <h2 className="text-sm font-semibold mb-2" style={{ color: "#F2F0EA", fontFamily: "'Fraunces', serif" }}>Sale</h2>
          <p className="text-xs mb-4" style={{ color: "#8C8A85" }}>
            Customers see a corner badge on sale pieces, with the new price and the old one crossed out. POS pre-fills the matching discount at checkout. A piece that already costs less than the sale price is never touched.
          </p>

          <div data-sale-section="store" className="mb-6">
            <p className="text-[10px] uppercase tracking-wide mb-1.5" style={{ color: "#75726C" }}>Everything on your storefront</p>
            <RuleEditor rule={storeRule} onChange={setStoreRule} />
            <div className="flex items-center gap-2 mt-2.5 flex-wrap">
              <button onClick={() => onSaveStoreSale(parseRule(storeRule))} disabled={!parseRule(storeRule)}
                className="text-xs font-semibold px-3 py-2 rounded-lg" style={{ background: "#C9A567", color: "#0B0B0C", opacity: parseRule(storeRule) ? 1 : 0.4 }}>
                {storefront.sale ? "Update" : "Start sale"}
              </button>
              {storefront.sale ? (
                <button onClick={() => { onSaveStoreSale(null); setStoreRule({ kind: "percent", value: "" }); }} className="text-xs font-medium px-3 py-2 rounded-lg border" style={{ borderColor: "#2A2A2F", color: "#C9695E" }}>End sale</button>
              ) : null}
            </div>
            <p className="text-[11px] mt-1.5" style={{ color: storefront.sale ? "#C9A567" : "#75726C" }}>
              {storefront.sale ? `Live now: ${ruleLong(storefront.sale)} (pieces showing on your storefront).` : "No store-wide sale running."}
            </p>
          </div>

          <div data-sale-section="collection" className="mb-6">
            <p className="text-[10px] uppercase tracking-wide mb-1.5" style={{ color: "#75726C" }}>A whole collection</p>
            {(collections || []).length === 0 ? (
              <p className="text-[11px]" style={{ color: "#75726C" }}>Create a collection first (see Collections above).</p>
            ) : (
              <>
                <select value={pickCollection} onChange={(e) => setPickCollection(e.target.value)} className="w-full px-3 py-2 rounded-lg border text-xs outline-none mb-2"
                  style={{ borderColor: "#2A2A2F", background: "#0B0B0C", color: "#F2F0EA" }}>
                  <option value="">Choose a collection…</option>
                  {(collections || []).filter((c) => !c.sale).map((c) => <option key={c.id} value={c.id}>{c.title}</option>)}
                </select>
                <RuleEditor rule={collRule} onChange={setCollRule} />
                <button onClick={() => { onSetCollectionSale(pickCollection, parseRule(collRule)); setPickCollection(""); setCollRule({ kind: "price", value: "" }); }} disabled={!pickCollection || !parseRule(collRule)}
                  className="text-xs font-semibold px-3 py-2 rounded-lg mt-2.5" style={{ background: "#C9A567", color: "#0B0B0C", opacity: (!pickCollection || !parseRule(collRule)) ? 0.4 : 1 }}>Add</button>
              </>
            )}
            {collectionsOnSale.length > 0 && (
              <div className="space-y-2 mt-3">
                {collectionsOnSale.map((c) => (
                  <div key={c.id} className="flex items-center justify-between gap-3 px-3 py-2 rounded-lg border" style={{ borderColor: "#2A2A2F", background: "#0B0B0C" }}>
                    <div className="min-w-0">
                      <p className="text-xs truncate" style={{ color: "#F2F0EA" }}>{c.title}</p>
                      <p className="text-[11px]" style={{ color: "#C9A567" }}>{ruleLong(c.sale)}</p>
                    </div>
                    <button onClick={() => onSetCollectionSale(c.id, null)} className="shrink-0 text-xs px-3 py-1.5 rounded-lg border" style={{ borderColor: "#2A2A2F", color: "#C9695E" }}>End</button>
                  </div>
                ))}
              </div>
            )}
          </div>

          <div data-sale-section="piece">
            <p className="text-[10px] uppercase tracking-wide mb-1.5" style={{ color: "#75726C" }}>Specific pieces</p>
            <select value={pickId} onChange={(e) => setPickId(e.target.value)} className="w-full px-3 py-2 rounded-lg border text-xs outline-none mb-2"
              style={{ borderColor: "#2A2A2F", background: "#0B0B0C", color: "#F2F0EA" }}>
              <option value="">Choose a piece…</option>
              {pickable.map((i) => <option key={i.id} value={i.id}>{productTitle(i)} · RM{i.price || "?"}</option>)}
            </select>
            <RuleEditor rule={pieceRule} onChange={setPieceRule} />
            <button onClick={() => { onSetItemSale(pickId, parseRule(pieceRule)); setPickId(""); setPieceRule({ kind: "percent", value: "" }); }} disabled={!pickId || !parseRule(pieceRule)}
              className="text-xs font-semibold px-3 py-2 rounded-lg mt-2.5" style={{ background: "#C9A567", color: "#0B0B0C", opacity: (!pickId || !parseRule(pieceRule)) ? 0.4 : 1 }}>Add</button>
            {onSaleItems.length > 0 && (
              <div className="space-y-2 mt-3">
                {onSaleItems.map((i) => (
                  <div key={i.id} className="flex items-center justify-between gap-3 px-3 py-2 rounded-lg border" style={{ borderColor: "#2A2A2F", background: "#0B0B0C" }}>
                    <div className="min-w-0">
                      <p className="text-xs truncate" style={{ color: "#F2F0EA" }}>{productTitle(i)}</p>
                      <p className="text-[11px]" style={{ color: "#8C8A85", fontFamily: "'IBM Plex Mono', monospace" }}>
                        {ruleText(i.sale)} · {fmtRM(applySale(i.price, i.sale))} <span style={{ textDecoration: "line-through" }}>RM{i.price}</span>
                      </p>
                    </div>
                    <button onClick={() => onSetItemSale(i.id, null)} className="shrink-0 p-1" aria-label="Remove sale"><X size={14} style={{ color: "#C9695E" }} /></button>
                  </div>
                ))}
              </div>
            )}
          </div>
          <p className="text-[11px] mt-4" style={{ color: "#75726C" }}>A piece's own sale beats its collection's, and a collection's beats the store-wide one. There's no end date, so end each sale yourself when it's over.</p>
        </div>
      )}

      <div className="rounded-xl border p-5" style={{ background: "#17171A", borderColor: "#2A2A2F" }}>
        <h2 className="text-sm font-semibold mb-2" style={{ color: "#F2F0EA", fontFamily: "'Fraunces', serif" }}>Business logo</h2>
        <p className="text-xs mb-3" style={{ color: "#8C8A85" }}>
          Shown at the top of your storefront in full colour, and on printed receipts in black &amp; white.
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
            <button onClick={() => logoInputRef.current && logoInputRef.current.click()} disabled={logoUploading}
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
          onChange={(e) => { const f = e.target.files && e.target.files[0]; if (f) onLogoUpload(f); e.target.value = ""; }} />
      </div>
    </div>
  );
}

const emptyForm = {
  title: "", titleTouched: false,
  category: "", itemType: "",
  brand: "", condition: "", colour: "", colourCondition: "", remark: "", cost: "", price: "", sku: "",
  photos: [], attributes: {}, customAttributes: [], showOnStorefront: false,
};

export default function App() {
  const [session, setSession] = useState(undefined); // undefined = still checking, null = signed out
  const [authEmail, setAuthEmail] = useState("");
  const [authPassword, setAuthPassword] = useState("");
  const [authBusy, setAuthBusy] = useState(false);
  const [authError, setAuthError] = useState("");
  const [showForgotPassword, setShowForgotPassword] = useState(false);
  const [forgotEmail, setForgotEmail] = useState("");
  const [forgotStatus, setForgotStatus] = useState("idle"); // idle | sending | sent | error
  const [recovery, setRecovery] = useState(null); // { accessToken, refreshToken } | null — set when the URL carries a password-recovery link
  const [newPassword, setNewPassword] = useState("");
  const [recoveryStatus, setRecoveryStatus] = useState("idle"); // idle | saving | done | error
  const [recoveryError, setRecoveryError] = useState("");
  const [authNotice, setAuthNotice] = useState("");
  const [profile, setProfile] = useState(null);
  const [customTemplates, setCustomTemplates] = useState({});
  const [templatesSaving, setTemplatesSaving] = useState(false);
  const [feedbackText, setFeedbackText] = useState("");
  const [feedbackStatus, setFeedbackStatus] = useState("idle"); // idle | sending | sent | error
  const [receiptTxn, setReceiptTxn] = useState(null);
  const [logoUrl, setLogoUrl] = useState(null);
  const [purchases, setPurchases] = useState([]);
  const [purchaseForm, setPurchaseForm] = useState({ sku: "", cost: "", place: "", photo: "", receiptPhoto: "", purchasedOn: new Date().toISOString().slice(0, 10) });
  const [purchaseSaving, setPurchaseSaving] = useState(false);
  const [purchasePeriod, setPurchasePeriod] = useState(() => { const d = new Date(); return { month: d.getMonth(), year: d.getFullYear() }; });
  const [purchaseDetailId, setPurchaseDetailId] = useState(null);
  const [storefront, setStorefront] = useState(null); // { slug, whatsappNumber, ... } | null
  const [collections, setCollections] = useState([]);
  const [storefrontSaving, setStorefrontSaving] = useState(false);
  const [storefrontError, setStorefrontError] = useState("");
  const [bannerUploading, setBannerUploading] = useState(false);
  const [logoUploading, setLogoUploading] = useState(false);
  const [upgradeBusy, setUpgradeBusy] = useState(false);

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
  const [postage, setPostage] = useState("");
  const [ledgerSearch, setLedgerSearch] = useState("");
  const [ledgerStatusFilter, setLedgerStatusFilter] = useState("all"); // all | listed | sold
  const [bulkSelectMode, setBulkSelectMode] = useState(false);
  const [bulkSelectedIds, setBulkSelectedIds] = useState(() => new Set());
  const [bulkUpdating, setBulkUpdating] = useState(false);
  const [activePlatform, setActivePlatform] = useState("carousell");
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

  // detect a password-recovery link (Supabase redirects back here with tokens in the URL hash)
  useEffect(() => {
    if (typeof window === "undefined") return;
    const hash = window.location.hash;
    if (hash.includes("type=recovery")) {
      const params = new URLSearchParams(hash.replace(/^#/, ""));
      const accessToken = params.get("access_token");
      const refreshToken = params.get("refresh_token");
      if (accessToken) {
        setRecovery({ accessToken, refreshToken });
        window.history.replaceState(null, "", window.location.pathname);
      }
    }
  }, []);

  // once signed in, load this user's items + transactions from Supabase
  useEffect(() => {
    if (!session) { setLoading(false); return; }
    (async () => {
      setLoading(true);
      try {
        // null (not []) marks "this one request failed" — distinct from a genuinely empty result,
        // so a transient network hiccup (common right after the phone reconnects) never gets
        // treated the same as "you have no data here" and wipes out what's already on screen
        const [itemRows, txnRows, profileRows, templateRows, brandingRows, purchaseRows, storefrontRows, collectionRows] = await Promise.all([
          apiRequest("/items?select=*&order=created_at.desc", { method: "GET" }, session, updateSession),
          apiRequest("/transactions?select=*&order=occurred_at.desc", { method: "GET" }, session, updateSession),
          apiRequest("/profiles?select=*", { method: "GET" }, session, updateSession).catch(() => null),
          apiRequest("/templates?select=*", { method: "GET" }, session, updateSession).catch(() => null),
          apiRequest("/branding?select=*", { method: "GET" }, session, updateSession).catch(() => null),
          apiRequest("/purchases?select=*&order=purchased_on.desc", { method: "GET" }, session, updateSession).catch(() => null),
          apiRequest("/storefronts?select=*", { method: "GET" }, session, updateSession).catch(() => null),
          apiRequest("/collections?select=*&order=position.asc,created_at.asc", { method: "GET" }, session, updateSession).catch(() => null),
        ]);
        setItems((itemRows || []).map(rowToItem));
        setTxns((txnRows || []).map(rowToTxn));
        if (profileRows !== null) setProfile((profileRows && profileRows[0]) || null);
        if (templateRows !== null) {
          const tRow = templateRows && templateRows[0];
          setCustomTemplates(tRow ? { carousell: tRow.carousell, threads: tRow.threads, grailed: tRow.grailed, depop: tRow.depop, shopee: tRow.shopee } : {});
        }
        if (brandingRows !== null) setLogoUrl((brandingRows && brandingRows[0]?.logo_url) || null);
        if (purchaseRows !== null) setPurchases((purchaseRows || []).map(rowToPurchase));
        if (storefrontRows !== null) {
          const sfRow = storefrontRows && storefrontRows[0];
          setStorefront(sfRow ? { slug: sfRow.slug, whatsappNumber: sfRow.whatsapp_number || "", bannerUrl: sfRow.banner_url || "", intro: sfRow.intro || "", sale: sfRow.sale || null, banners: Array.isArray(sfRow.banners) ? sfRow.banners : [], announcement: sfRow.announcement || "", faq: Array.isArray(sfRow.faq) ? sfRow.faq : [] } : null);
        }
        if (collectionRows !== null) setCollections((collectionRows || []).map(rowToCollection));
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
      const dataUrl = await compressImage(file, 300, 0.85);
      const publicUrl = await uploadPhotoToStorage(dataUrl, `${session.user.id}/logo.png`, session);
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
      const data = await signInRequest(authEmail, authPassword);
      updateSession(data);
      setAuthPassword("");
    } catch (err) {
      setAuthError(err.message);
    } finally {
      setAuthBusy(false);
    }
  }

  async function handleForgotPasswordSubmit(e) {
    e.preventDefault();
    if (!forgotEmail.trim()) return;
    setForgotStatus("sending");
    try {
      await recoverPasswordRequest(forgotEmail.trim());
      setForgotStatus("sent");
    } catch (err) {
      setForgotStatus("error");
    }
  }

  async function handleSetNewPassword(e) {
    e.preventDefault();
    if (!recovery || newPassword.length < 10) { setRecoveryError("Use at least 10 characters."); return; }
    setRecoveryStatus("saving");
    setRecoveryError("");
    try {
      const user = await updatePasswordRequest(recovery.accessToken, newPassword);
      updateSession({ access_token: recovery.accessToken, refresh_token: recovery.refreshToken, user });
      setRecoveryStatus("done");
      setTimeout(() => setRecovery(null), 1200);
    } catch (err) {
      setRecoveryError(err.message);
      setRecoveryStatus("error");
    }
  }

  async function handleSignOut() {
    if (session) await signOutRequest(session);
    updateSession(null);
    // clear every piece of account-specific state — otherwise, if the next account's first
    // fetch has even a brief hiccup, the previous account's data could still be sitting here
    setItems([]); setTxns([]); setForm(emptyForm); setCart([]); setActiveId(null);
    setProfile(null); setCustomTemplates({}); setLogoUrl(null); setPurchases([]); setStorefront(null); setCollections([]);
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
          show_on_storefront: draft.showOnStorefront ?? false,
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
      showOnStorefront: item.showOnStorefront ?? false,
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

  async function saveStorefrontSettings(slug, whatsappNumber, intro) {
    if (!session) return;
    const cleanSlug = slug.trim().toLowerCase().replace(/[^a-z0-9-]/g, "");
    if (!cleanSlug) { setStorefrontError("Pick a slug using only letters, numbers, and dashes."); return; }
    setStorefrontSaving(true);
    setStorefrontError("");
    try {
      const row = { user_id: session.user.id, slug: cleanSlug, whatsapp_number: whatsappNumber.trim() || null, intro: (intro || "").trim() || null, updated_at: new Date().toISOString() };
      await apiRequest("/storefronts?on_conflict=user_id", {
        method: "POST", body: JSON.stringify(row), headers: { Prefer: "resolution=merge-duplicates,return=representation" },
      }, session, updateSession);
      setStorefront((prev) => ({ ...prev, slug: cleanSlug, whatsappNumber: whatsappNumber.trim(), intro: (intro || "").trim() }));
    } catch (e) {
      setStorefrontError(e.message && e.message.includes("duplicate") ? "That link is already taken — try another." : "Couldn't save — try again.");
    } finally {
      setStorefrontSaving(false);
    }
  }

  async function handleBannerUpload(file) {
    if (!session || !file || !storefront?.slug) return;
    setBannerUploading(true);
    try {
      const dataUrl = await compressImage(file, 1200, 0.8);
      const publicUrl = await uploadPhotoToStorage(dataUrl, `${session.user.id}/banner_${Date.now()}.jpg`, session);
      const row = { user_id: session.user.id, slug: storefront.slug, banner_url: publicUrl, updated_at: new Date().toISOString() };
      await apiRequest("/storefronts?on_conflict=user_id", {
        method: "POST", body: JSON.stringify(row), headers: { Prefer: "resolution=merge-duplicates,return=representation" },
      }, session, updateSession);
      setStorefront((prev) => ({ ...prev, bannerUrl: publicUrl }));
    } catch (e) {
      setSaveError(true);
    } finally {
      setBannerUploading(false);
    }
  }

  async function handleBannerRemove() {
    if (!session || !storefront?.slug) return;
    setBannerUploading(true);
    try {
      const row = { user_id: session.user.id, slug: storefront.slug, banner_url: null, updated_at: new Date().toISOString() };
      await apiRequest("/storefronts?on_conflict=user_id", {
        method: "POST", body: JSON.stringify(row), headers: { Prefer: "resolution=merge-duplicates,return=representation" },
      }, session, updateSession);
      setStorefront((prev) => ({ ...prev, bannerUrl: "" }));
    } catch (e) {
      setSaveError(true);
    } finally {
      setBannerUploading(false);
    }
  }

  async function saveStoreFields(fields, local) {
    if (!session || !storefront || !storefront.slug) return false;
    const row = { user_id: session.user.id, slug: storefront.slug, ...fields, updated_at: new Date().toISOString() };
    try {
      await apiRequest("/storefronts?on_conflict=user_id", { method: "POST", body: JSON.stringify(row), headers: { Prefer: "resolution=merge-duplicates,return=representation" } }, session, updateSession);
      setStorefront((prev) => ({ ...prev, ...local }));
      return true;
    } catch (e) { setSaveError(true); return false; }
  }
  const saveBanners = (arr) => saveStoreFields({ banners: arr, banner_url: (arr[0] && arr[0].url) || null }, { banners: arr, bannerUrl: (arr[0] && arr[0].url) || "" });
  async function uploadStoreImage(file) {
    const dataUrl = await compressImage(file, 1200, 0.8);
    return uploadPhotoToStorage(dataUrl, `${session.user.id}/store_${Date.now()}_${Math.random().toString(36).slice(2, 6)}.jpg`, session);
  }
  async function saveCollection(c) {
    const row = { user_id: session.user.id, title: c.title.trim(), description: (c.description || "").trim() || null, cover_url: c.coverUrl || null, item_ids: c.itemIds };
    try {
      if (c.id) {
        const [u] = await apiRequest(`/collections?id=eq.${c.id}`, { method: "PATCH", body: JSON.stringify(row) }, session, updateSession);
        setCollections((prev) => prev.map((x) => (x.id === c.id ? rowToCollection(u) : x)));
      } else {
        const [i] = await apiRequest("/collections", { method: "POST", body: JSON.stringify({ ...row, position: collections.length }) }, session, updateSession);
        setCollections((prev) => [...prev, rowToCollection(i)]);
      }
      return true;
    } catch (e) { setSaveError(true); return false; }
  }
  async function deleteCollection(id) {
    setCollections((prev) => prev.filter((x) => x.id !== id));
    try { await apiRequest(`/collections?id=eq.${id}`, { method: "DELETE" }, session, updateSession); }
    catch (e) { setSaveError(true); }
  }

  const saveStoreSale = (rule) => saveStoreFields({ sale: rule }, { sale: rule });

  async function setItemSale(id, rule) {
    setItems((prev) => prev.map((i) => (i.id === id ? { ...i, sale: rule } : i)));
    try { await apiRequest(`/items?id=eq.${id}`, { method: "PATCH", body: JSON.stringify({ sale: rule }) }, session, updateSession); }
    catch (e) { setSaveError(true); }
  }

  async function setCollectionSale(id, rule) {
    setCollections((prev) => prev.map((c) => (c.id === id ? { ...c, sale: rule } : c)));
    try { await apiRequest(`/collections?id=eq.${id}`, { method: "PATCH", body: JSON.stringify({ sale: rule }) }, session, updateSession); }
    catch (e) { setSaveError(true); }
  }

  async function toggleItemFeatured(id, next) {
    setItems((prev) => prev.map((i) => (i.id === id ? { ...i, featured: next } : i)));
    try { await apiRequest(`/items?id=eq.${id}`, { method: "PATCH", body: JSON.stringify({ featured: next }) }, session, updateSession); }
    catch (e) { setSaveError(true); }
  }

  async function toggleItemStorefront(id, next) {
    setItems((prev) => prev.map((i) => (i.id === id ? { ...i, showOnStorefront: next } : i)));
    try { await apiRequest(`/items?id=eq.${id}`, { method: "PATCH", body: JSON.stringify({ show_on_storefront: next }) }, session, updateSession); }
    catch (e) { setSaveError(true); }
  }

  async function bulkSetStorefront(next) {
    if (bulkSelectedIds.size === 0) return;
    const ids = Array.from(bulkSelectedIds);
    setBulkUpdating(true);
    setItems((prev) => prev.map((i) => (bulkSelectedIds.has(i.id) ? { ...i, showOnStorefront: next } : i)));
    try {
      await apiRequest(`/items?id=in.(${ids.join(",")})`, { method: "PATCH", body: JSON.stringify({ show_on_storefront: next }) }, session, updateSession);
      setBulkSelectedIds(new Set());
      setBulkSelectMode(false);
    } catch (e) {
      setSaveError(true);
    } finally {
      setBulkUpdating(false);
    }
  }

  async function savePurchase() {
    if (!session || purchaseSaving) return;
    if (!purchaseForm.cost || !purchaseForm.place.trim()) return;
    setPurchaseSaving(true);
    try {
      let photo = purchaseForm.photo;
      let receiptPhoto = purchaseForm.receiptPhoto;
      if (photo && photo.startsWith("data:")) photo = await uploadPhotoToStorage(photo, `${session.user.id}/purchase_${Date.now()}.jpg`, session);
      if (receiptPhoto && receiptPhoto.startsWith("data:")) receiptPhoto = await uploadPhotoToStorage(receiptPhoto, `${session.user.id}/receipt_${Date.now()}.jpg`, session);
      const row = purchaseToRow({ ...purchaseForm, photo, receiptPhoto }, session.user.id);
      const [inserted] = await apiRequest("/purchases", { method: "POST", body: JSON.stringify(row) }, session, updateSession);
      setPurchases((prev) => [rowToPurchase(inserted), ...prev]);
      setPurchaseForm({ sku: "", cost: "", place: "", photo: "", receiptPhoto: "", purchasedOn: new Date().toISOString().slice(0, 10) });
    } catch (e) {
      setSaveError(true);
    } finally {
      setPurchaseSaving(false);
    }
  }

  async function removePurchase(id) {
    setPurchases((prev) => prev.filter((p) => p.id !== id));
    try { await apiRequest(`/purchases?id=eq.${id}`, { method: "DELETE" }, session, updateSession); }
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

  function addToCart(item) {
    const sale = saleFor(item, { store: storefront && storefront.sale, collections });
    const saleDiscount = sale.off > 0 ? sale.off : 0;
    setCart((prev) => (prev.some((c) => c.id === item.id) ? prev : [...prev, { id: item.id, discount: saleDiscount }]));
  }
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
    const onSale = saleFor(item, { store: storefront && storefront.sale, collections }).off > 0;
    return { ...item, discount, originalPrice, finalPrice, onSale };
  }).filter(Boolean);
  const cartSubtotal = cartLines.reduce((sum, l) => sum + l.originalPrice, 0);
  const cartDiscount = cartLines.reduce((sum, l) => sum + l.discount, 0);
  const cartTotal = cartSubtotal - cartDiscount;
  const postageAmount = Math.max(0, Number(postage) || 0); // 0 = DFOD (customer pays the courier on delivery)
  const amountToCollect = Math.round((cartTotal + postageAmount) * 100) / 100;

  async function completeSale() {
    if (cartLines.length === 0 || !session) return;
    setSaveError(false);
    const soldOn = new Date().toISOString().slice(0, 10);
    const soldIds = cartLines.map((l) => l.id);
    const txnRow = txnToRow({
      timestamp: new Date().toISOString(), method: payMethod, total: cartTotal, subtotal: cartSubtotal, discount: cartDiscount, postage: postageAmount,
      lines: cartLines.map((l) => ({ title: productTitle(l), sku: l.sku, price: l.price, cost: Number(l.cost || 0), discount: l.discount, finalPrice: l.finalPrice })),
    }, session.user.id);
    try {
      await apiRequest(`/items?id=in.(${soldIds.join(",")})`, { method: "PATCH", body: JSON.stringify({ status: "sold", sold_on: soldOn }) }, session, updateSession);
      const [insertedTxn] = await apiRequest("/transactions", { method: "POST", body: JSON.stringify(txnRow) }, session, updateSession);
      setItems((prev) => prev.map((i) => (soldIds.includes(i.id) ? { ...i, status: "sold", soldOn } : i)));
      setTxns((prev) => [rowToTxn(insertedTxn), ...prev]);
      setReceiptTxn({ ...rowToTxn(insertedTxn), invoiceNumber: `INV-${String(txns.length + 1).padStart(4, "0")}` });
      setCart([]);
      setPostage("");
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
    const round2 = (n) => Math.round(n * 100) / 100; // avoids floating-point artifacts like 41.269999999999996
    const totalSales = round2(periodTxns.reduce((sum, t) => sum + Number(t.total || 0), 0));
    const orders = periodTxns.length;
    const unitsSold = periodTxns.reduce((sum, t) => sum + t.lines.length, 0);
    const discounts = round2(periodTxns.reduce((sum, t) => sum + Number(t.discount || 0), 0));
    const postageCollected = round2(periodTxns.reduce((sum, t) => sum + Number(t.postage || 0), 0));
    // net profit = final sale price minus cost, per line — sales made before this feature default cost to 0
    const netProfit = round2(periodTxns.reduce((sum, t) => sum + t.lines.reduce((s, l) => s + (Number(l.finalPrice || 0) - Number(l.cost || 0)), 0), 0));
    return { totalSales, orders, unitsSold, discounts, netProfit, postage: postageCollected };
  }, [periodTxns]);

  const purchasesInPeriod = useMemo(() => {
    return purchases.filter((p) => {
      if (!p.purchasedOn) return false;
      const d = new Date(p.purchasedOn);
      return d.getMonth() === purchasePeriod.month && d.getFullYear() === purchasePeriod.year;
    });
  }, [purchases, purchasePeriod]);

  const purchaseSummary = useMemo(() => {
    const raw = purchasesInPeriod.reduce((sum, p) => sum + Number(p.cost || 0), 0);
    const totalCost = Math.round(raw * 100) / 100; // avoids floating-point artifacts like 41.269999999999996
    return { totalCost, count: purchasesInPeriod.length };
  }, [purchasesInPeriod]);

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
    return { listedCount: listed.length, soldCount: sold.length, profit };
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
  const purchaseDetail = purchases.find((p) => p.id === purchaseDetailId) || null;

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
              item_type: normalizeItemType(get("itemType") || get("item_type") || "") || null,
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

  const storefrontMatch = typeof window !== "undefined" && window.location.pathname.match(/^\/store\/([a-zA-Z0-9-]+)/);
  if (storefrontMatch) {
    return <PublicStorefront slug={storefrontMatch[1]} />;
  }

  if (recovery) {
    return (
      <div className="min-h-screen w-full flex items-center justify-center p-6" style={{ background: "#0B0B0C", fontFamily: "'Inter', sans-serif" }}>
        <style>{`@import url('https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,500;9..144,600&family=Inter:wght@400;500;600;700;800&display=swap');`}</style>
        <div className="w-full max-w-sm rounded-xl border p-6" style={{ background: "#17171A", borderColor: "#2A2A2F" }}>
          <h1 className="text-xl font-semibold mb-1 text-center" style={{ color: "#F2F0EA", fontFamily: "'Fraunces', serif" }}>Set a new password</h1>
          {recoveryStatus === "done" ? (
            <p className="text-sm text-center mt-4" style={{ color: "#C9A567" }}>Password updated — signing you in…</p>
          ) : (
            <form onSubmit={handleSetNewPassword} className="mt-4">
              <input type="password" value={newPassword} onChange={(e) => setNewPassword(e.target.value)} placeholder="New password" autoFocus
                className="w-full px-3 py-2.5 rounded-lg border text-sm outline-none mb-3" style={{ borderColor: "#2A2A2F", background: "#0B0B0C", color: "#F2F0EA" }} />
              {recoveryError && <p className="text-xs mb-3" style={{ color: "#C9695E" }}>{recoveryError}</p>}
              <button type="submit" disabled={recoveryStatus === "saving"}
                className="w-full py-2.5 rounded-xl text-sm font-semibold" style={{ background: "#C9A567", color: "#0B0B0C", opacity: recoveryStatus === "saving" ? 0.6 : 1 }}>
                {recoveryStatus === "saving" ? "Saving…" : "Update password"}
              </button>
            </form>
          )}
        </div>
      </div>
    );
  }

  if (session === undefined) {
    return (
      <div className="min-h-screen w-full flex items-center justify-center" style={{ background: "#0B0B0C", fontFamily: "'Inter', sans-serif" }}>
        <style>{`@import url('https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,500;9..144,600&family=Space+Grotesk:wght@500;600;700&family=Inter:wght@400;500;600;700;800&display=swap');`}</style>
        <Loader2 size={20} className="animate-spin" style={{ color: "#75726C" }} />
      </div>
    );
  }

  if (!session) {
    return (
      <div className="min-h-screen w-full flex items-center justify-center px-6" style={{ background: "#0B0B0C", fontFamily: "'Inter', sans-serif" }}>
        <style>{`@import url('https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,500;9..144,600&family=Space+Grotesk:wght@500;600;700&family=Inter:wght@400;500;600;700;800&display=swap');`}</style>
        <div className="w-full max-w-sm">
          <div className="flex items-baseline gap-2 justify-center mb-8">
            <span className="text-2xl font-semibold tracking-tight" style={{ color: "#F2F0EA", fontFamily: "'Fraunces', serif" }}>RAVUNO</span>
            <span className="text-[10px] font-medium tracking-wider" style={{ color: "#75726C" }}>V4.9</span>
          </div>
          <div className="rounded-xl border p-6" style={{ background: "#17171A", borderColor: "#2A2A2F" }}>
            <h1 className="text-lg font-semibold mb-1" style={{ color: "#F2F0EA", fontFamily: "'Fraunces', serif" }}>
              Sign in
            </h1>
            <p className="text-xs mb-5" style={{ color: "#75726C" }}>
              Private inventory. Sign in to continue.
            </p>
            {authNotice && <p className="text-xs px-3 py-2 rounded-lg mb-4" style={{ background: "#241F13", color: "#C9A567" }}>{authNotice}</p>}
            {authError && <p className="text-xs px-3 py-2 rounded-lg mb-4" style={{ background: "#2A1715", color: "#C9695E" }}>{authError}</p>}

            {showForgotPassword ? (
              <>
                <p className="text-xs mb-3" style={{ color: "#8C8A85" }}>Enter your email and we'll send a password reset link.</p>
                {forgotStatus === "sent" ? (
                  <p className="text-xs px-3 py-2 rounded-lg mb-3" style={{ background: "#241F13", color: "#C9A567" }}>Check your email for a reset link.</p>
                ) : (
                  <form onSubmit={handleForgotPasswordSubmit} className="space-y-3">
                    <Field label="Email" type="email" value={forgotEmail} onChange={(e) => setForgotEmail(e.target.value)} placeholder="you@example.com" />
                    {forgotStatus === "error" && <p className="text-xs" style={{ color: "#C9695E" }}>Couldn't send the reset email — try again.</p>}
                    <button type="submit" disabled={forgotStatus === "sending"}
                      className="w-full py-2.5 rounded-xl text-sm font-semibold" style={{ background: "#C9A567", color: "#0B0B0C", opacity: forgotStatus === "sending" ? 0.6 : 1 }}>
                      {forgotStatus === "sending" ? "Sending…" : "Send reset link"}
                    </button>
                  </form>
                )}
                <button onClick={() => { setShowForgotPassword(false); setForgotStatus("idle"); }} className="w-full text-center text-xs mt-4" style={{ color: "#8C8A85" }}>
                  ← Back to sign in
                </button>
              </>
            ) : (
              <>
                <form onSubmit={handleAuthSubmit} className="space-y-3">
                  <Field label="Email" type="email" value={authEmail} onChange={(e) => setAuthEmail(e.target.value)} placeholder="you@example.com" />
                  <Field label="Password" type="password" value={authPassword} onChange={(e) => setAuthPassword(e.target.value)} placeholder="••••••••" />
                  <button type="submit" disabled={authBusy}
                    className="w-full flex items-center justify-center gap-2 py-2.5 rounded-xl text-sm font-semibold transition-opacity"
                    style={{ background: "#C9A567", color: "#0B0B0C", opacity: authBusy ? 0.6 : 1, fontFamily: "'IBM Plex Mono', monospace" }}>
                    {authBusy ? <Loader2 size={14} className="animate-spin" /> : null}
                    Sign in
                  </button>
                </form>
                <button onClick={() => { setShowForgotPassword(true); setForgotEmail(authEmail); }} className="w-full text-center text-xs mt-3" style={{ color: "#8C8A85" }}>
                  Forgot password?
                </button>
              </>
            )}
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
            <span className="text-[10px] font-medium tracking-wider" style={{ color: "#75726C", fontFamily: "'IBM Plex Mono', monospace" }}>V4.9</span>
          </div>
          <nav className="flex gap-1 p-1 rounded-lg overflow-x-auto no-scrollbar max-w-full" style={{ background: "#0B0B0C", border: "1px solid #2A2A2F", WebkitOverflowScrolling: "touch", scrollSnapType: "x proximity" }}>
            {[{ id: "crosslist", label: "Cross-List" }, { id: "inventory", label: "Stock Ledger" }, { id: "pos", label: "POS" }, { id: "purchases", label: "Purchases" }, { id: "sales", label: "Sales" }, { id: "analytics", label: "Analytics" }, { id: "storefront", label: "Storefront" }, { id: "support", label: "Support" }].map((t) => (
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
              <section className="min-w-0">
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
                          <p className="text-[10px] uppercase tracking-[0.2em]" style={{ color: "#C9A567", fontFamily: "'IBM Plex Mono', monospace" }}>Details</p>
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
                        <p className="text-[10px] uppercase tracking-[0.2em]" style={{ color: "#C9A567", fontFamily: "'IBM Plex Mono', monospace" }}>Pricing & stock</p>
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

                  <div className="flex items-center justify-between mt-5 px-3 py-2.5 rounded-xl border" style={{ borderColor: "#2A2A2F", background: "#0B0B0C" }}>
                    <div>
                      <p className="text-xs font-semibold" style={{ color: "#F2F0EA" }}>Show on storefront</p>
                      <p className="text-[11px]" style={{ color: "#75726C" }}>{form.showOnStorefront ? "Active — visible to visitors" : "Draft — hidden from your storefront"}</p>
                    </div>
                    <button onClick={() => setForm((f) => ({ ...f, showOnStorefront: !f.showOnStorefront }))}
                      className="relative w-11 h-6 rounded-full shrink-0 transition-colors"
                      style={{ background: form.showOnStorefront ? "#C9A567" : "#2A2A2F" }}>
                      <span className="absolute top-0.5 w-5 h-5 rounded-full transition-transform" style={{ background: "#0B0B0C", left: form.showOnStorefront ? "22px" : "2px" }} />
                    </button>
                  </div>

                  <button onClick={saveListing} disabled={!isFormUsable}
                    className="mt-3 w-full flex items-center justify-center gap-2 py-2.5 rounded-xl text-sm font-semibold transition-opacity"
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
                            <span className="text-xs font-semibold" style={{ color: "#C9A567", fontFamily: "'IBM Plex Mono', monospace" }}>RM{l.price || "?"}</span>
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

              <section className="min-w-0">
                <div className="flex items-center justify-between gap-2 mb-3 flex-wrap">
                  <p className="text-xs tracking-wide uppercase font-medium break-words" style={{ color: "#75726C" }}>
                    {activeListing ? `Ready to paste — ${productTitle(activeListing)}` : "Live preview — fill in the product on the left"}
                  </p>
                  {activeListing && (
                    <span className="text-sm font-semibold shrink-0" style={{ color: "#C9A567", fontFamily: "'IBM Plex Mono', monospace" }}>RM{activeListing.price || "?"}</span>
                  )}
                </div>
                <div className="flex gap-2 overflow-x-auto pb-2 mb-3">
                  {PLATFORMS.map((p) => {
                    const on = activePlatform === p.id;
                    return (
                      <button key={p.id} onClick={() => setActivePlatform(p.id)}
                        className="shrink-0 text-xs font-semibold px-3.5 py-2 rounded-lg border transition-colors"
                        style={{ borderColor: on ? "#C9A567" : "#2A2A2F", color: on ? "#C9A567" : "#8C8A85", background: on ? "#241F13" : "transparent" }}>
                        {p.label}
                      </button>
                    );
                  })}
                </div>
                <div className="grid gap-4">
                  {PLATFORMS.filter((p) => p.id === activePlatform).map((p) => {
                    const { title, description } = splitTitleDescription(activeTemplates[p.id]);
                    return (
                      <div key={p.id} className="rounded-xl border overflow-hidden" style={{ background: "#17171A", borderColor: "#2A2A2F" }}>
                        <div className="px-4 py-3 border-b" style={{ borderColor: "#232327" }}>
                          <span className="text-xs font-semibold px-2 py-1 rounded-md" style={{ background: "#0B0B0C", color: "#F2F0EA" }}>{p.label}</span>
                        </div>
                        <div className="px-4 py-3 border-b" style={{ borderColor: "#232327" }}>
                          <div className="flex items-start justify-between gap-2 mb-1.5">
                            <span className="text-[10px] uppercase tracking-wide" style={{ color: "#75726C" }}>Title</span>
                            <CopyButton text={title} label="Copy" />
                          </div>
                          <p className="text-xs leading-relaxed" style={{ fontFamily: "'IBM Plex Mono', monospace", color: "#F2F0EA" }}>{title}</p>
                        </div>
                        <div className="px-4 py-3">
                          <div className="flex items-start justify-between gap-2 mb-1.5">
                            <span className="text-[10px] uppercase tracking-wide" style={{ color: "#75726C" }}>Description</span>
                            <CopyButton text={description} label="Copy" />
                          </div>
                          <pre className="text-xs whitespace-pre-wrap leading-relaxed" style={{ fontFamily: "'IBM Plex Mono', monospace", color: "#F2F0EA" }}>{description}</pre>
                        </div>
                      </div>
                    );
                  })}
                </div>
              </section>
            </div>
          )}

          {tab === "inventory" && (
            <div className="space-y-8">
              <div className="grid grid-cols-2 gap-3">
                <StatCard icon={Package} label="Listed" value={stats.listedCount} tone="#4A4A52" />
                <StatCard icon={TrendingUp} label="Sold" value={stats.soldCount} tone="#8C8A85" />
              </div>

              <div>
                <div className="flex items-center justify-between mb-3 flex-wrap gap-2">
                  <h2 className="text-base font-semibold" style={{ color: "#F2F0EA", fontFamily: "'Fraunces', serif" }}>Stock Items</h2>
                  <div className="flex items-center gap-2 flex-wrap">
                    <button onClick={() => { setBulkSelectMode((v) => !v); setBulkSelectedIds(new Set()); }}
                      className="text-xs font-semibold px-3 py-2 rounded-lg border"
                      style={{ borderColor: bulkSelectMode ? "#C9A567" : "#2A2A2F", color: bulkSelectMode ? "#C9A567" : "#F2F0EA" }}>
                      {bulkSelectMode ? "Cancel" : "Select"}
                    </button>
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

                {bulkSelectMode && (
                  <div className="flex items-center justify-between gap-2 mb-3 px-3 py-2.5 rounded-xl border flex-wrap" style={{ borderColor: "#2A2A2F", background: "#17171A" }}>
                    <div className="flex items-center gap-3">
                      <span className="text-xs font-medium" style={{ color: "#8C8A85" }}>{bulkSelectedIds.size} selected</span>
                      <button onClick={() => setBulkSelectedIds(new Set(ledgerItems.map((i) => i.id)))} className="text-xs font-semibold" style={{ color: "#C9A567" }}>Select all</button>
                    </div>
                    <div className="flex items-center gap-2">
                      <button onClick={() => bulkSetStorefront(true)} disabled={bulkUpdating || bulkSelectedIds.size === 0}
                        className="text-xs font-semibold px-3 py-1.5 rounded-lg" style={{ background: "#C9A567", color: "#0B0B0C", opacity: (bulkUpdating || bulkSelectedIds.size === 0) ? 0.5 : 1 }}>
                        Show on Storefront
                      </button>
                      <button onClick={() => bulkSetStorefront(false)} disabled={bulkUpdating || bulkSelectedIds.size === 0}
                        className="text-xs font-semibold px-3 py-1.5 rounded-lg border" style={{ borderColor: "#2A2A2F", color: "#F2F0EA", opacity: (bulkUpdating || bulkSelectedIds.size === 0) ? 0.5 : 1 }}>
                        Hide from Storefront
                      </button>
                    </div>
                  </div>
                )}

                <div className="hidden sm:grid gap-2 px-4 mb-1 text-[11px] uppercase tracking-wide"
                  style={{ color: "#8C8A85", fontFamily: "'IBM Plex Mono', monospace", gridTemplateColumns: "1fr 1.6fr 1.1fr 1.1fr 0.7fr 0.7fr 0.8fr" }}>
                  <span>SKU</span><span>Title</span><span>Category</span><span>Item Type</span><span>Cost</span><span>Price</span><span>Status</span>
                </div>

                <div className="space-y-2">
                  {ledgerItems.map((i) => {
                    const selected = bulkSelectedIds.has(i.id);
                    function handleRowClick() {
                      if (bulkSelectMode) {
                        setBulkSelectedIds((prev) => {
                          const next = new Set(prev);
                          if (next.has(i.id)) next.delete(i.id); else next.add(i.id);
                          return next;
                        });
                      } else {
                        setDetailId(i.id); setGalleryIndex(0);
                      }
                    }
                    return (
                      <div key={i.id} onClick={handleRowClick}
                        className="rounded-xl border px-4 py-3 cursor-pointer transition-colors flex items-center gap-3"
                        style={{ borderColor: selected ? "#C9A567" : i.status === "sold" ? "#8C8A85" : "#2A2A2F", background: "#17171A" }}>

                        {bulkSelectMode && (
                          <div className="w-5 h-5 rounded border-2 shrink-0 flex items-center justify-center"
                            style={{ borderColor: selected ? "#C9A567" : "#2A2A2F", background: selected ? "#C9A567" : "transparent" }}>
                            {selected && <Check size={13} style={{ color: "#0B0B0C" }} />}
                          </div>
                        )}

                        <div className="flex-1 min-w-0">
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

                        {!bulkSelectMode && (
                          <div className="flex items-center gap-2 mt-2 pt-2 border-t" style={{ borderColor: "#232327" }} onClick={(e) => e.stopPropagation()}>
                            <Trash2 size={14} style={{ color: "#C9695E", cursor: "pointer" }} onClick={() => removeItem(i.id)} />
                          </div>
                        )}
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
              <section className="min-w-0">
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

              <section className="min-w-0">
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
                    {cartLines.some((l) => l.onSale) && <p className="text-[11px] pb-1" style={{ color: "#C9A567" }}>Sale discount pre-filled. Edit a line's discount to change it.</p>}
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

                  <div className="mb-4">
                    <p className="text-[10px] uppercase tracking-wide mb-1.5" style={{ color: "#75726C", fontFamily: "'IBM Plex Mono', monospace" }}>Postage</p>
                    <div className="flex flex-wrap items-center gap-2">
                      {[["DFOD", 0], ["West RM7.50", 7.5], ["East RM15", 15]].map(([label, amt]) => {
                        const on = postageAmount === amt;
                        return (
                          <button key={label} onClick={() => setPostage(amt === 0 ? "" : String(amt))}
                            className="text-xs font-semibold px-3 py-2 rounded-lg border"
                            style={{ borderColor: on ? "#C9A567" : "#2A2A2F", color: on ? "#C9A567" : "#8C8A85", background: on ? "#241F13" : "transparent" }}>
                            {label}
                          </button>
                        );
                      })}
                      <input type="number" inputMode="decimal" min="0" value={postage} onChange={(e) => setPostage(e.target.value)} placeholder="Other"
                        className="w-20 px-3 py-2 rounded-lg border text-xs outline-none" style={{ borderColor: "#2A2A2F", background: "#0B0B0C", color: "#F2F0EA" }} />
                    </div>
                    <p className="text-[11px] mt-1.5" style={{ color: "#75726C" }}>DFOD = customer pays the courier on delivery, so nothing is added here.</p>
                    {postageAmount > 0 && (
                      <div className="flex items-center justify-between font-semibold pt-2 text-sm" style={{ color: "#F2F0EA" }}>
                        <span>Customer pays</span><span style={{ fontFamily: "'IBM Plex Mono', monospace" }}>RM{amountToCollect}</span>
                      </div>
                    )}
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

          {tab === "purchases" && (
            <div className="max-w-2xl mx-auto space-y-6">
              <div className="rounded-xl border p-5" style={{ background: "#17171A", borderColor: "#2A2A2F" }}>
                <h2 className="text-lg font-semibold mb-1" style={{ color: "#F2F0EA", fontFamily: "'Fraunces', serif" }}>Log a purchase</h2>
                <p className="text-xs mb-4" style={{ color: "#8C8A85" }}>
                  Record what you actually paid for stock. This is for your own records and monthly cost totals.
                </p>

                <div className="flex gap-3 mb-4">
                  <label className="w-20 h-20 rounded-lg border-2 border-dashed flex items-center justify-center shrink-0 cursor-pointer overflow-hidden" style={{ borderColor: "#2A2A2F" }}>
                    {purchaseForm.photo ? (
                      <img src={purchaseForm.photo} alt="" className="w-full h-full object-cover" />
                    ) : (
                      <ImageIcon size={18} style={{ color: "#75726C" }} />
                    )}
                    <input type="file" accept="image/*" className="hidden"
                      onChange={async (e) => { const f = e.target.files?.[0]; if (f) { const url = await compressImage(f); setPurchaseForm((p) => ({ ...p, photo: url })); } e.target.value = ""; }} />
                  </label>
                  <label className="w-20 h-20 rounded-lg border-2 border-dashed flex items-center justify-center shrink-0 cursor-pointer overflow-hidden" style={{ borderColor: "#2A2A2F" }}>
                    {purchaseForm.receiptPhoto ? (
                      <img src={purchaseForm.receiptPhoto} alt="" className="w-full h-full object-cover" />
                    ) : (
                      <div className="text-center">
                        <Receipt size={16} style={{ color: "#75726C" }} className="mx-auto" />
                        <span className="text-[9px]" style={{ color: "#75726C" }}>Receipt</span>
                      </div>
                    )}
                    <input type="file" accept="image/*" className="hidden"
                      onChange={async (e) => { const f = e.target.files?.[0]; if (f) { const url = await compressImage(f); setPurchaseForm((p) => ({ ...p, receiptPhoto: url })); } e.target.value = ""; }} />
                  </label>
                </div>

                <div className="space-y-3 mb-4">
                  <label className="block">
                    <span className="block text-[10px] mb-1 uppercase tracking-wide" style={{ color: "#75726C" }}>Code / SKU</span>
                    <input value={purchaseForm.sku} onChange={(e) => setPurchaseForm((p) => ({ ...p, sku: e.target.value }))} placeholder="TZ-014"
                      className="w-full px-3 py-2 rounded-lg border text-sm outline-none" style={{ borderColor: "#2A2A2F", background: "#0B0B0C", color: "#F2F0EA" }} />
                  </label>
                  <label className="block">
                    <span className="block text-[10px] mb-1 uppercase tracking-wide" style={{ color: "#75726C" }}>Cost (RM)</span>
                    <input type="number" value={purchaseForm.cost} onChange={(e) => setPurchaseForm((p) => ({ ...p, cost: e.target.value }))} placeholder="15"
                      className="w-full px-3 py-2 rounded-lg border text-sm outline-none" style={{ borderColor: "#2A2A2F", background: "#0B0B0C", color: "#F2F0EA" }} />
                  </label>
                  <label className="block">
                    <span className="block text-[10px] mb-1 uppercase tracking-wide" style={{ color: "#75726C" }}>Bought at</span>
                    <input value={purchaseForm.place} onChange={(e) => setPurchaseForm((p) => ({ ...p, place: e.target.value }))} placeholder="Pasar malam"
                      className="w-full px-3 py-2 rounded-lg border text-sm outline-none" style={{ borderColor: "#2A2A2F", background: "#0B0B0C", color: "#F2F0EA" }} />
                  </label>
                  <label className="block">
                    <span className="block text-[10px] mb-1 uppercase tracking-wide" style={{ color: "#75726C" }}>Date</span>
                    <input type="date" value={purchaseForm.purchasedOn} onChange={(e) => setPurchaseForm((p) => ({ ...p, purchasedOn: e.target.value }))}
                      className="w-full px-3 py-2 rounded-lg border text-sm outline-none"
                      style={{
                        borderColor: "#2A2A2F", background: "#0B0B0C", color: "#F2F0EA",
                        WebkitAppearance: "none", appearance: "none",
                        width: "100%", maxWidth: "100%", display: "block", overflow: "hidden", boxSizing: "border-box",
                      }} />
                  </label>
                </div>

                <button onClick={savePurchase} disabled={purchaseSaving || !purchaseForm.cost || !purchaseForm.place.trim()}
                  className="w-full flex items-center justify-center gap-2 py-2.5 rounded-xl text-sm font-semibold transition-opacity"
                  style={{ background: "#C9A567", color: "#0B0B0C", opacity: (purchaseSaving || !purchaseForm.cost || !purchaseForm.place.trim()) ? 0.4 : 1, fontFamily: "'IBM Plex Mono', monospace" }}>
                  <Plus size={15} /> {purchaseSaving ? "Saving…" : "Save purchase"}
                </button>
              </div>

              <div>
                <div className="flex items-center justify-between mb-3">
                  <h2 className="text-base font-semibold" style={{ color: "#F2F0EA", fontFamily: "'Fraunces', serif" }}>This period</h2>
                  <div className="flex items-center gap-2">
                    <button onClick={() => setPurchasePeriod((p) => { const m = p.month - 1; return m < 0 ? { month: 11, year: p.year - 1 } : { ...p, month: m }; })}
                      className="p-1.5 rounded-lg border" style={{ borderColor: "#2A2A2F" }}><ChevronLeft size={14} style={{ color: "#8C8A85" }} /></button>
                    <span className="text-xs font-medium" style={{ color: "#F2F0EA", fontFamily: "'IBM Plex Mono', monospace" }}>
                      {new Date(purchasePeriod.year, purchasePeriod.month).toLocaleDateString(undefined, { month: "short", year: "numeric" })}
                    </span>
                    <button onClick={() => setPurchasePeriod((p) => { const m = p.month + 1; return m > 11 ? { month: 0, year: p.year + 1 } : { ...p, month: m }; })}
                      className="p-1.5 rounded-lg border" style={{ borderColor: "#2A2A2F" }}><ChevronRight size={14} style={{ color: "#8C8A85" }} /></button>
                  </div>
                </div>
                <div className="grid grid-cols-2 gap-3 mb-5">
                  <StatCard icon={DollarSign} label="Total Cost (RM)" value={purchaseSummary.totalCost} tone="#C9A567" />
                  <StatCard icon={Receipt} label="Purchases" value={purchaseSummary.count} tone="#4A4A52" />
                </div>

                <div className="space-y-2">
                  {purchasesInPeriod.map((p) => (
                    <button key={p.id} onClick={() => setPurchaseDetailId(p.id)}
                      className="w-full flex items-center justify-between px-3 py-2.5 rounded-xl border text-left transition-colors"
                      style={{ borderColor: "#2A2A2F", background: "#17171A" }}>
                      <div className="flex items-center gap-3 min-w-0">
                        {p.photo ? <img src={p.photo} alt="" className="w-10 h-10 rounded-lg object-cover shrink-0" /> : <div className="w-10 h-10 rounded-lg shrink-0" style={{ background: "#0B0B0C" }} />}
                        <div className="min-w-0">
                          <p className="text-sm font-medium truncate" style={{ color: "#F2F0EA" }}>{p.sku || "No code"}</p>
                          <p className="text-xs truncate" style={{ color: "#75726C" }}>{p.place} · {new Date(p.purchasedOn).toLocaleDateString()}</p>
                        </div>
                      </div>
                      <div className="flex items-center gap-2 shrink-0">
                        <span className="text-sm font-semibold" style={{ color: "#C9A567", fontFamily: "'IBM Plex Mono', monospace" }}>RM{p.cost}</span>
                        <Trash2 size={14} style={{ color: "#C9695E" }} onClick={(e) => { e.stopPropagation(); removePurchase(p.id); }} />
                      </div>
                    </button>
                  ))}
                  {purchasesInPeriod.length === 0 && (
                    <p className="text-xs text-center py-6" style={{ color: "#75726C" }}>No purchases logged for this period.</p>
                  )}
                </div>
              </div>
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
                    <StatCard icon={Receipt} label="Postage (RM)" value={salesSummary.postage} tone="#4A4A52" />
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

          {tab === "analytics" && (
            <AnalyticsTab session={session} updateSession={updateSession} storefront={storefront} items={items} txns={txns} purchases={purchases} />
          )}

          {tab === "storefront" && (
            <StorefrontTab storefront={storefront} onSaveStorefront={saveStorefrontSettings} storefrontSaving={storefrontSaving} storefrontError={storefrontError}
              collections={collections} onSaveFields={saveStoreFields} onSaveBanners={saveBanners} onUploadImage={uploadStoreImage} onSaveCollection={saveCollection} onDeleteCollection={deleteCollection}
              logoUrl={logoUrl} onLogoUpload={handleLogoUpload} onLogoRemove={handleLogoRemove} logoUploading={logoUploading}
              items={items} onSaveStoreSale={saveStoreSale} onSetItemSale={setItemSale} onSetCollectionSale={setCollectionSale} />
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

            <div className="flex items-center justify-between mt-4 px-3 py-2.5 rounded-xl border" style={{ borderColor: "#2A2A2F", background: "#0B0B0C" }}>
              <div>
                <p className="text-xs font-semibold" style={{ color: "#F2F0EA" }}>Show on storefront</p>
                <p className="text-[11px]" style={{ color: "#75726C" }}>{detailItem.showOnStorefront ? "Active — visible to visitors" : "Draft — hidden from your storefront"}</p>
              </div>
              <button onClick={() => toggleItemStorefront(detailItem.id, !detailItem.showOnStorefront)}
                className="relative w-11 h-6 rounded-full shrink-0 transition-colors"
                style={{ background: detailItem.showOnStorefront ? "#C9A567" : "#2A2A2F" }}>
                <span className="absolute top-0.5 w-5 h-5 rounded-full transition-transform" style={{ background: "#0B0B0C", left: detailItem.showOnStorefront ? "22px" : "2px" }} />
              </button>
            </div>

            <div className="flex items-center justify-between mt-2 px-3 py-2.5 rounded-xl border" style={{ borderColor: "#2A2A2F", background: "#0B0B0C" }}>
              <div>
                <p className="text-xs font-semibold" style={{ color: "#F2F0EA" }}>Featured</p>
                <p className="text-[11px]" style={{ color: "#75726C" }}>{detailItem.featured ? "Highlighted on your storefront homepage" : "Not featured"}</p>
              </div>
              <button onClick={() => toggleItemFeatured(detailItem.id, !detailItem.featured)}
                className="relative w-11 h-6 rounded-full shrink-0 transition-colors"
                style={{ background: detailItem.featured ? "#C9A567" : "#2A2A2F" }}>
                <span className="absolute top-0.5 w-5 h-5 rounded-full transition-transform" style={{ background: "#0B0B0C", left: detailItem.featured ? "22px" : "2px" }} />
              </button>
            </div>

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
              {selectedInvoice.postage > 0 && <DetailRow label="Postage" value={`RM${selectedInvoice.postage}`} />}
              {selectedInvoice.postage > 0 && <DetailRow label="Customer paid" value={`RM${Math.round((Number(selectedInvoice.total) + Number(selectedInvoice.postage)) * 100) / 100}`} />}
            </dl>
          </div>
        </div>
      )}

      {purchaseDetail && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4" style={{ background: "rgba(23,22,26,0.55)" }} onClick={() => setPurchaseDetailId(null)}>
          <div className="max-w-md w-full rounded-xl border p-6 relative max-h-[85vh] overflow-y-auto shadow-sm" style={{ background: "#17171A", borderColor: "#2A2A2F" }} onClick={(e) => e.stopPropagation()}>
            <button onClick={() => setPurchaseDetailId(null)} className="absolute top-4 right-4 w-7 h-7 rounded-full flex items-center justify-center" style={{ background: "#1E1E22", color: "#F2F0EA" }} aria-label="Close">
              <X size={15} />
            </button>

            <div className="grid grid-cols-2 gap-3 mb-5">
              <div>
                <p className="text-[10px] uppercase tracking-wide mb-1.5" style={{ color: "#75726C" }}>Item photo</p>
                {purchaseDetail.photo ? (
                  <img src={purchaseDetail.photo} alt="" onClick={() => setLightbox({ photos: [purchaseDetail.photo], index: 0 })}
                    className="w-full h-32 object-cover rounded-lg cursor-pointer" style={{ border: "1px solid #2A2A2F" }} />
                ) : (
                  <div className="w-full h-32 rounded-lg flex items-center justify-center" style={{ background: "#0B0B0C", border: "1px dashed #2A2A2F" }}>
                    <ImageIcon size={18} style={{ color: "#75726C" }} />
                  </div>
                )}
              </div>
              <div>
                <p className="text-[10px] uppercase tracking-wide mb-1.5" style={{ color: "#75726C" }}>Receipt</p>
                {purchaseDetail.receiptPhoto ? (
                  <img src={purchaseDetail.receiptPhoto} alt="" onClick={() => setLightbox({ photos: [purchaseDetail.receiptPhoto], index: 0 })}
                    className="w-full h-32 object-cover rounded-lg cursor-pointer" style={{ border: "1px solid #2A2A2F" }} />
                ) : (
                  <div className="w-full h-32 rounded-lg flex items-center justify-center" style={{ background: "#0B0B0C", border: "1px dashed #2A2A2F" }}>
                    <Receipt size={18} style={{ color: "#75726C" }} />
                  </div>
                )}
              </div>
            </div>

            <dl>
              <DetailRow label="Code / SKU" value={purchaseDetail.sku || "—"} />
              <DetailRow label="Cost" value={`RM${purchaseDetail.cost}`} />
              <DetailRow label="Bought at" value={purchaseDetail.place || "—"} />
              <DetailRow label="Date" value={new Date(purchaseDetail.purchasedOn).toLocaleDateString()} />
            </dl>

            <button onClick={() => { removePurchase(purchaseDetail.id); setPurchaseDetailId(null); }}
              className="mt-5 w-full flex items-center justify-center gap-2 py-2.5 rounded-xl text-sm font-semibold border" style={{ borderColor: "#2A2A2F", color: "#C9695E" }}>
              <Trash2 size={14} /> Delete purchase
            </button>
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
              {receiptTxn.postage > 0 && (
                <div className="text-xs mb-1"><div className="flex justify-between"><span>Postage</span><span>RM{receiptTxn.postage}</span></div></div>
              )}
              <div className="text-sm font-bold flex justify-between pt-2 mt-2 border-t" style={{ borderColor: "#17161A" }}>
                <span>TOTAL</span><span>RM{Math.round((Number(receiptTxn.total) + Number(receiptTxn.postage || 0)) * 100) / 100}</span>
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
