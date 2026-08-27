import React, { useState, useEffect, useMemo, useRef } from "react";
import {
  Check, Copy, Ruler, Plus, Trash2, ChevronRight, ChevronLeft,
  Package, TrendingUp, Clock, Image as ImageIcon, X, Loader2,
  Search, Download, Star, Tag, Receipt, ShoppingBag, Percent
} from "lucide-react";
import * as XLSX from "xlsx";
import { LineChart, Line, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer } from "recharts";

// RAVUNO design tokens — bg: #F5F4F0  surface: #FFFFFF  ink: #17161A  muted: #726F68
// border: #E4E1D9  accent: #146658  destructive: #B3261E
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
    brand: row.brand || "", condition: row.condition || "", colour: row.colour || "", remark: row.remark || "",
    cost: row.cost ?? "", price: row.price ?? "", sku: row.sku || "", status: row.status || "listed",
    attributes: row.attributes || {}, customAttributes: row.custom_attributes || [], photos: row.photos || [],
    addedOn: row.added_on, listedOn: row.listed_on, soldOn: row.sold_on,
  };
}
function itemToRow(item, userId) {
  return {
    user_id: userId, title: item.title || null, category: item.category || null, item_type: item.itemType || null,
    brand: item.brand || null, condition: item.condition || null, colour: item.colour || null, remark: item.remark || null,
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
  const condLine = `${p.condition || "?"}/10`;

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
  const condLine = `${p.condition || "?"}/10`;

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
const TEMPLATE_TOKENS = ["title", "brand", "category", "condition", "colour", "remark", "price", "sku", "size", "measurements"];
function renderCustomTemplate(str, p) {
  const tokens = {
    title: productTitle(p), brand: p.brand || "", category: CATEGORY_LABELS[p.category] || "",
    condition: `${p.condition || "?"}/10`, colour: p.colour || "", remark: p.remark || "",
    price: p.price || "", sku: p.sku || "", size: getDisplaySize(p), measurements: getMeasurementsString(p),
  };
  return str.replace(/\{(\w+)\}/g, (match, key) => (key in tokens ? tokens[key] : match));
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
      style={{ borderColor: copied ? "#726F68" : "#17161A", color: copied ? "#726F68" : "#17161A", background: copied ? "#E4EFEA" : "transparent", fontFamily: "'IBM Plex Mono', monospace" }}
    >
      {copied ? <Check size={13} /> : <Copy size={13} />}
      {copied ? "Copied" : "Copy"}
    </button>
  );
}
function Field({ label, value, onChange, placeholder, type = "text" }) {
  return (
    <label className="block">
      <span className="block text-[11px] mb-1 tracking-wide uppercase" style={{ color: "#726F68", fontFamily: "'IBM Plex Mono', monospace" }}>{label}</span>
      <input type={type} value={value} onChange={onChange} placeholder={placeholder}
        className="w-full px-3 py-2 rounded-lg border text-sm outline-none"
        style={{ borderColor: "#E4E1D9", background: "#FFFFFF", color: "#17161A" }} />
    </label>
  );
}
function ReadOnlyField({ label, value, placeholder }) {
  return (
    <label className="block">
      <span className="block text-[11px] mb-1 tracking-wide uppercase" style={{ color: "#726F68", fontFamily: "'IBM Plex Mono', monospace" }}>{label}</span>
      <div className="w-full px-3 py-2 rounded-lg border text-sm"
        style={{ borderColor: "#E4E1D9", background: "#F7F6F2", color: value ? "#17161A" : "#8B887F" }}>
        {value || placeholder}
      </div>
    </label>
  );
}
function Select({ label, value, onChange, children }) {
  return (
    <label className="block">
      <span className="block text-[11px] mb-1 tracking-wide uppercase" style={{ color: "#726F68", fontFamily: "'IBM Plex Mono', monospace" }}>{label}</span>
      <select value={value} onChange={onChange}
        className="w-full px-3 py-2 rounded-lg border text-sm outline-none"
        style={{ borderColor: "#E4E1D9", background: "#FFFFFF", color: "#17161A" }}>
        {children}
      </select>
    </label>
  );
}
function StatCard({ icon: Icon, label, value, tone }) {
  return (
    <div className="rounded-xl border px-4 py-3 flex items-center gap-3" style={{ borderColor: "#E4E1D9", background: "#FFFFFF" }}>
      <div className="w-9 h-9 rounded-full flex items-center justify-center shrink-0" style={{ background: tone, color: "#FFFFFF" }}>
        <Icon size={16} />
      </div>
      <div>
        <p className="text-lg font-semibold leading-none" style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{value}</p>
        <p className="text-[11px] uppercase tracking-wide mt-1" style={{ color: "#726F68" }}>{label}</p>
      </div>
    </div>
  );
}
function DetailRow({ label, value }) {
  return (
    <div className="flex items-start justify-between gap-4 py-1 border-b" style={{ borderColor: "#EDEAE2" }}>
      <dt className="text-xs uppercase tracking-wide shrink-0" style={{ color: "#726F68", fontFamily: "'IBM Plex Mono', monospace" }}>{label}</dt>
      <dd className="text-right" style={{ color: "#17161A" }}>{value}</dd>
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
      <button onClick={onClose} className="absolute top-4 right-4 w-9 h-9 rounded-full flex items-center justify-center" style={{ background: "#FFFFFF", color: "#17161A" }} aria-label="Close">
        <X size={16} />
      </button>
      {photos.length > 1 && (
        <button onClick={(e) => { e.stopPropagation(); prev(); }} className="absolute left-2 sm:left-6 w-9 h-9 rounded-full flex items-center justify-center" style={{ background: "#FFFFFF", color: "#17161A" }} aria-label="Previous photo">
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
        <button onClick={(e) => { e.stopPropagation(); next(); }} className="absolute right-2 sm:right-6 w-9 h-9 rounded-full flex items-center justify-center" style={{ background: "#FFFFFF", color: "#17161A" }} aria-label="Next photo">
          <ChevronRight size={18} />
        </button>
      )}
      {onDownload && (
        <button onClick={(e) => { e.stopPropagation(); onDownload(index); }}
          className="absolute bottom-6 flex items-center gap-1.5 text-xs font-semibold px-3 py-2 rounded-lg"
          style={{ background: "#FFFFFF", color: "#17161A", fontFamily: "'IBM Plex Mono', monospace" }}>
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
      <span className="block text-[11px] mb-1 tracking-wide uppercase" style={{ color: "#726F68", fontFamily: "'IBM Plex Mono', monospace" }}>
        Photos ({photos.length}/{MAX_PHOTOS})
      </span>
      <div className="flex flex-wrap gap-2">
        {photos.map((p, idx) => (
          <div key={idx} className="relative w-16 h-16">
            <img src={p} alt="" onClick={() => onPreview && onPreview(idx)}
              className="w-16 h-16 object-cover rounded-lg border-2 cursor-pointer"
              style={{ borderColor: idx === 0 ? "#146658" : "#17161A" }} />
            {idx === 0 ? (
              <span className="absolute -top-1.5 -left-1.5 text-[8px] font-bold px-1 py-0.5 rounded" style={{ background: "#146658", color: "#FFFFFF" }}>MAIN</span>
            ) : (
              <button onClick={() => makePrimary(idx)} title="Set as primary photo"
                className="absolute -bottom-1.5 -left-1.5 w-5 h-5 rounded-full flex items-center justify-center" style={{ background: "#FFFFFF", color: "#726F68", border: "1px solid #E4E1D9" }}>
                <Star size={10} />
              </button>
            )}
            <button onClick={() => removeAt(idx)} className="absolute -top-1.5 -right-1.5 w-5 h-5 rounded-full flex items-center justify-center" style={{ background: "#B3261E", color: "#FFFFFF" }}>
              <X size={11} />
            </button>
          </div>
        ))}
        {photos.length < MAX_PHOTOS && (
          <button onClick={() => fileRef.current?.click()}
            className="w-16 h-16 rounded-lg border-2 border-dashed flex flex-col items-center justify-center gap-0.5 text-[10px]"
            style={{ borderColor: "#E4E1D9", color: "#8B887F" }}>
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
        <div className="text-xs px-3 py-2 rounded-lg flex items-start gap-2" style={{ background: "#FBEEEC", color: "#B3261E" }}>
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
      <span className="block text-[11px] mb-1 tracking-wide uppercase" style={{ color: "#726F68", fontFamily: "'IBM Plex Mono', monospace" }}>Custom attributes</span>
      {value.map((c, idx) => (
        <div key={idx} className="grid grid-cols-[1fr_1fr_auto] gap-2">
          <input value={c.label} onChange={(e) => update(idx, "label", e.target.value)} placeholder="Field name"
            className="px-3 py-2 rounded-lg border text-sm outline-none" style={{ borderColor: "#E4E1D9", background: "#FFFFFF", color: "#17161A" }} />
          <input value={c.value} onChange={(e) => update(idx, "value", e.target.value)} placeholder="Value"
            className="px-3 py-2 rounded-lg border text-sm outline-none" style={{ borderColor: "#E4E1D9", background: "#FFFFFF", color: "#17161A" }} />
          <button onClick={() => remove(idx)}><X size={14} style={{ color: "#B3261E" }} /></button>
        </div>
      ))}
      <button onClick={add} className="text-xs font-medium px-3 py-1.5 rounded-lg border flex items-center gap-1"
        style={{ borderColor: "#17161A", color: "#17161A", fontFamily: "'IBM Plex Mono', monospace" }}>
        <Plus size={12} /> Add field
      </button>
    </div>
  );
}

function TemplatesTab({ customTemplates, onSave, saving }) {
  const [draft, setDraft] = useState(customTemplates || {});
  useEffect(() => { setDraft(customTemplates || {}); }, [customTemplates]);

  const sampleProduct = {
    title: "Uniqlo Fleece Jacket", brand: "Uniqlo", category: "apparel", itemType: "Jacket",
    condition: "9", colour: "Navy", remark: "small pilling near cuff", price: "45", sku: "TZ-014",
    attributes: { chest: "21", length: "27" },
  };

  return (
    <div className="max-w-2xl mx-auto space-y-6">
      <div className="rounded-xl border p-5" style={{ background: "#FFFFFF", borderColor: "#E4E1D9" }}>
        <h2 className="text-sm font-semibold mb-2" style={{ color: "#17161A" }}>Your templates</h2>
        <p className="text-xs mb-3" style={{ color: "#726F68" }}>
          Write your own listing text for each platform. Leave a box empty to use Ravuno's built-in default instead.
        </p>
        <div className="flex flex-wrap gap-1.5">
          {TEMPLATE_TOKENS.map((t) => (
            <span key={t} className="text-[10px] font-medium px-2 py-1 rounded-md" style={{ background: "#F5F4F0", color: "#726F68", fontFamily: "'IBM Plex Mono', monospace" }}>{`{${t}}`}</span>
          ))}
        </div>
      </div>

      {PLATFORMS.map((p) => (
        <div key={p.id} className="rounded-xl border p-5" style={{ background: "#FFFFFF", borderColor: "#E4E1D9" }}>
          <div className="flex items-center justify-between mb-2">
            <span className="text-sm font-semibold" style={{ color: "#17161A" }}>{p.label}</span>
            {draft[p.id] && (
              <button onClick={() => setDraft((d) => ({ ...d, [p.id]: "" }))} className="text-xs font-medium" style={{ color: "#B3261E" }}>
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
            style={{ borderColor: "#E4E1D9", background: "#FFFFFF", color: "#17161A", fontFamily: "'IBM Plex Mono', monospace" }}
          />
          {draft[p.id] && draft[p.id].trim() && (
            <div className="mt-2 rounded-lg p-3" style={{ background: "#F5F4F0" }}>
              <p className="text-[10px] uppercase tracking-wide mb-1" style={{ color: "#8B887F" }}>Preview</p>
              <pre className="text-[11px] whitespace-pre-wrap" style={{ fontFamily: "'IBM Plex Mono', monospace", color: "#17161A" }}>
                {renderCustomTemplate(draft[p.id], sampleProduct)}
              </pre>
            </div>
          )}
        </div>
      ))}

      <button onClick={() => onSave(draft)} disabled={saving}
        className="w-full flex items-center justify-center gap-2 py-2.5 rounded-xl text-sm font-semibold transition-opacity"
        style={{ background: "#146658", color: "#FFFFFF", opacity: saving ? 0.6 : 1, fontFamily: "'IBM Plex Mono', monospace" }}>
        {saving ? "Saving…" : "Save templates"}
      </button>
    </div>
  );
}

const emptyForm = {
  title: "", titleTouched: false,
  category: "", itemType: "",
  brand: "", condition: "", colour: "", remark: "", cost: "", price: "", sku: "",
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
  const [skuSearch, setSkuSearch] = useState("");
  const [cart, setCart] = useState([]);
  const [payMethod, setPayMethod] = useState("Cash");
  const [ledgerSearch, setLedgerSearch] = useState("");
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
        const [itemRows, txnRows, profileRows, templateRows] = await Promise.all([
          apiRequest("/items?select=*&order=created_at.desc", { method: "GET" }, session, updateSession),
          apiRequest("/transactions?select=*&order=occurred_at.desc", { method: "GET" }, session, updateSession),
          apiRequest("/profiles?select=*", { method: "GET" }, session, updateSession).catch(() => []),
          apiRequest("/templates?select=*", { method: "GET" }, session, updateSession).catch(() => []),
        ]);
        setItems((itemRows || []).map(rowToItem));
        setTxns((txnRows || []).map(rowToTxn));
        setProfile((profileRows && profileRows[0]) || null);
        const tRow = templateRows && templateRows[0];
        setCustomTemplates(tRow ? { carousell: tRow.carousell, threads: tRow.threads, grailed: tRow.grailed, depop: tRow.depop, shopee: tRow.shopee } : {});
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
    const draft = { ...form, status: "listed", addedOn: new Date().toISOString().slice(0, 10), listedOn: new Date().toISOString().slice(0, 10) };
    try {
      // upload any photos first, so the saved row stores real URLs rather than base64 blobs
      const uploadedPhotos = await Promise.all(
        (draft.photos || []).map((p, idx) => (p.startsWith("data:") ? uploadPhotoToStorage(p, `${session.user.id}/${Date.now()}_${idx}.jpg`, session) : p))
      );
      const row = itemToRow({ ...draft, photos: uploadedPhotos }, session.user.id);
      const [inserted] = await apiRequest("/items", { method: "POST", body: JSON.stringify(row) }, session, updateSession);
      const newItem = rowToItem(inserted);
      setItems((prev) => [newItem, ...prev]);
      setActiveId(newItem.id);
      setForm(emptyForm);
    } catch (e) {
      setSaveError(true);
    }
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

  function addToCart(item) { setCart((prev) => (prev.some((c) => c.id === item.id) ? prev : [...prev, { id: item.id, qty: 1, discount: 0 }])); }
  function removeFromCart(id) { setCart((prev) => prev.filter((c) => c.id !== id)); }
  function updateQty(id, qty) {
    const n = Math.max(1, parseInt(qty) || 1);
    setCart((prev) => prev.map((c) => (c.id === id ? { ...c, qty: n } : c)));
  }
  function updateDiscount(id, discount) {
    const d = Math.max(0, parseFloat(discount) || 0);
    setCart((prev) => prev.map((c) => (c.id === id ? { ...c, discount: d } : c)));
  }
  // discount is transaction-only — never written back to the product's stored price
  const cartLines = cart.map((c) => {
    const item = items.find((i) => i.id === c.id);
    if (!item) return null;
    const originalPrice = Number(item.price || 0) * c.qty;
    const discount = Math.min(c.discount || 0, originalPrice);
    const finalPrice = originalPrice - discount;
    return { ...item, qty: c.qty, discount, originalPrice, finalPrice };
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
      lines: cartLines.map((l) => ({ title: productTitle(l), sku: l.sku, qty: l.qty, price: l.price, discount: l.discount, finalPrice: l.finalPrice })),
    }, session.user.id);
    try {
      await apiRequest(`/items?id=in.(${soldIds.join(",")})`, { method: "PATCH", body: JSON.stringify({ status: "sold", sold_on: soldOn }) }, session, updateSession);
      const [insertedTxn] = await apiRequest("/transactions", { method: "POST", body: JSON.stringify(txnRow) }, session, updateSession);
      setItems((prev) => prev.map((i) => (soldIds.includes(i.id) ? { ...i, status: "sold", soldOn } : i)));
      setTxns((prev) => [rowToTxn(insertedTxn), ...prev]);
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
    const unitsSold = periodTxns.reduce((sum, t) => sum + t.lines.reduce((s, l) => s + Number(l.qty || 0), 0), 0);
    const discounts = periodTxns.reduce((sum, t) => sum + Number(t.discount || 0), 0);
    return { totalSales, orders, unitsSold, discounts };
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
    if (!q) return items;
    return items.filter((i) =>
      (i.sku || "").toLowerCase().includes(q) ||
      (i.brand || "").toLowerCase().includes(q) ||
      (i.itemType || "").toLowerCase().includes(q) ||
      (CATEGORY_LABELS[i.category] || "").toLowerCase().includes(q) ||
      productTitle(i).toLowerCase().includes(q) ||
      (i.status || "").toLowerCase().includes(q)
    );
  }, [items, ledgerSearch]);

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

  if (session === undefined) {
    return (
      <div className="min-h-screen w-full flex items-center justify-center" style={{ background: "#F5F4F0", fontFamily: "'Inter', sans-serif" }}>
        <style>{`@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');`}</style>
        <Loader2 size={20} className="animate-spin" style={{ color: "#8B887F" }} />
      </div>
    );
  }

  if (!session && !showAuthForm) {
    return (
      <div className="min-h-screen w-full" style={{ background: "#F5F4F0", fontFamily: "'Inter', sans-serif", color: "#17161A" }}>
        <style>{`@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&family=IBM+Plex+Mono:wght@400;500;600&display=swap');`}</style>

        <header style={{ background: "#FFFFFF", borderBottom: "1px solid #E4E1D9" }}>
          <div className="max-w-5xl mx-auto px-6 py-5 flex items-center justify-between">
            <div className="flex items-baseline gap-2">
              <span className="text-xl font-extrabold tracking-tight">RAVUNO</span>
              <span className="text-[10px] font-medium tracking-wider" style={{ color: "#8B887F" }}>V1</span>
            </div>
            <button onClick={() => { setAuthMode("signin"); setShowAuthForm(true); }} className="text-sm font-medium px-4 py-2 rounded-lg border" style={{ borderColor: "#E4E1D9" }}>
              Sign in
            </button>
          </div>
        </header>

        <main className="max-w-5xl mx-auto px-6">
          {/* hero */}
          <section className="py-16 text-center max-w-2xl mx-auto">
            <h1 className="text-3xl sm:text-4xl font-bold tracking-tight mb-4">One place to list, track, and sell everything.</h1>
            <p className="text-base mb-8" style={{ color: "#726F68" }}>
              Ravuno helps solo resellers generate ready-to-paste listings for Carousell, Grailed, Depop, Threads and Shopee,
              manage stock in one ledger, run a simple POS with negotiated discounts, and track sales performance — all in one login.
            </p>
            <button onClick={() => { setAuthMode("signup"); setShowAuthForm(true); }}
              className="inline-flex items-center gap-2 px-6 py-3 rounded-xl text-sm font-semibold"
              style={{ background: "#146658", color: "#FFFFFF", fontFamily: "'IBM Plex Mono', monospace" }}>
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
                className="rounded-xl border p-5 text-left transition-colors" style={{ background: "#FFFFFF", borderColor: "#E4E1D9" }}>
                <h3 className="text-sm font-semibold mb-1">{f.title}</h3>
                <p className="text-xs mb-3" style={{ color: "#726F68" }}>{f.desc}</p>
                <span className="text-xs font-semibold" style={{ color: "#146658" }}>See how it looks →</span>
              </button>
            ))}
          </section>

          {/* free + support */}
          <section className="pb-16">
            <div className="rounded-xl border p-8 text-center max-w-xl mx-auto" style={{ background: "#FFFFFF", borderColor: "#146658" }}>
              <span className="text-xs font-semibold px-2 py-0.5 rounded-full" style={{ background: "#EEF3F1", color: "#146658" }}>Free, always</span>
              <p className="text-2xl font-bold mt-3 mb-2">RM0<span className="text-sm font-normal" style={{ color: "#8B887F" }}> / forever</span></p>
              <p className="text-sm mb-6" style={{ color: "#726F68" }}>
                Full access to Cross-List, Stock Ledger, POS, and Sales — no paywall, no trial limits.
                If it helps your business, an optional tip is always appreciated once you're signed in.
              </p>
              <button onClick={() => { setAuthMode("signup"); setShowAuthForm(true); }} className="w-full py-2.5 rounded-lg text-sm font-semibold" style={{ background: "#146658", color: "#FFFFFF" }}>
                Get started — it's free
              </button>
            </div>
          </section>
        </main>

        <footer className="border-t" style={{ borderColor: "#E4E1D9" }}>
          <div className="max-w-5xl mx-auto px-6 py-8 text-center">
            <p className="text-xs mb-1" style={{ color: "#8B887F" }}>Questions, support, or refund requests:</p>
            <a href="mailto:Unknownlable00@gmail.com" className="text-sm font-medium" style={{ color: "#146658" }}>Unknownlable00@gmail.com</a>
          </div>
        </footer>

        {landingPreview && (
          <div className="fixed inset-0 z-50 flex items-center justify-center p-4" style={{ background: "rgba(23,22,26,0.6)" }} onClick={() => setLandingPreview(null)}>
            <div className="w-full max-w-lg rounded-xl border overflow-hidden max-h-[85vh] overflow-y-auto" style={{ background: "#F5F4F0", borderColor: "#E4E1D9" }} onClick={(e) => e.stopPropagation()}>
              <div className="flex items-center justify-between px-5 py-4 border-b" style={{ background: "#FFFFFF", borderColor: "#E4E1D9" }}>
                <span className="text-sm font-semibold" style={{ color: "#17161A" }}>
                  {landingPreview === "crosslist" && "Cross-List preview"}
                  {landingPreview === "inventory" && "Stock Ledger preview"}
                  {landingPreview === "pos" && "POS + Sales preview"}
                </span>
                <button onClick={() => setLandingPreview(null)} className="w-7 h-7 rounded-full flex items-center justify-center" style={{ background: "#F5F4F0", color: "#17161A" }} aria-label="Close">
                  <X size={15} />
                </button>
              </div>

              <div className="p-5">
                {landingPreview === "crosslist" && (
                  <div className="space-y-4">
                    <div className="rounded-lg border p-4" style={{ background: "#FFFFFF", borderColor: "#E4E1D9" }}>
                      <p className="text-xs font-semibold mb-3" style={{ color: "#17161A" }}>Uniqlo Fleece Jacket · Apparel</p>
                      <div className="grid grid-cols-3 gap-2 text-xs" style={{ color: "#726F68" }}>
                        <div><span className="block text-[10px] uppercase" style={{ color: "#8B887F" }}>Size</span>M</div>
                        <div><span className="block text-[10px] uppercase" style={{ color: "#8B887F" }}>Condition</span>9/10</div>
                        <div><span className="block text-[10px] uppercase" style={{ color: "#8B887F" }}>Price</span>RM45</div>
                      </div>
                    </div>
                    <div className="rounded-lg border overflow-hidden" style={{ background: "#FFFFFF", borderColor: "#E4E1D9" }}>
                      <div className="px-4 py-2 border-b flex items-center justify-between" style={{ borderColor: "#EDEAE2" }}>
                        <span className="text-xs font-semibold px-2 py-0.5 rounded" style={{ background: "#F5F4F0" }}>Carousell</span>
                        <span className="text-[10px] font-medium" style={{ color: "#146658" }}>Copy</span>
                      </div>
                      <pre className="text-[11px] px-4 py-3 whitespace-pre-wrap" style={{ fontFamily: "'IBM Plex Mono', monospace", color: "#17161A" }}>
{`Uniqlo Fleece Jacket — Size M

Condition: 9/10
Chest: 21" | Length: 27"
Remark: small pilling near cuff

Price: RM45
SKU: TZ-014

Thanks for checking out my shop!`}
                      </pre>
                    </div>
                    <p className="text-xs" style={{ color: "#8B887F" }}>Same product also generates ready-to-paste versions for Threads, Grailed, Depop, and Shopee.</p>
                  </div>
                )}

                {landingPreview === "inventory" && (
                  <div className="space-y-4">
                    <div className="grid grid-cols-3 gap-2">
                      <StatCard icon={Package} label="Listed" value={12} tone="#17161A" />
                      <StatCard icon={TrendingUp} label="Sold" value={5} tone="#726F68" />
                      <StatCard icon={Clock} label="Stale 21d+" value={2} tone="#B3261E" />
                    </div>
                    <div className="rounded-lg border overflow-hidden" style={{ background: "#FFFFFF", borderColor: "#E4E1D9" }}>
                      {[
                        { sku: "TZ-014", title: "Uniqlo Fleece Jacket", price: 45, status: "LISTED" },
                        { sku: "TZ-009", title: "Nike Windbreaker", price: 65, status: "SOLD" },
                        { sku: "TZ-021", title: "Carhartt Work Pants", price: 38, status: "LISTED" },
                      ].map((row, idx) => (
                        <div key={row.sku} className="flex items-center justify-between px-4 py-3 text-xs" style={{ borderTop: idx > 0 ? "1px solid #EDEAE2" : "none" }}>
                          <div>
                            <p className="font-medium" style={{ color: "#17161A" }}>{row.title}</p>
                            <p style={{ color: "#8B887F", fontFamily: "'IBM Plex Mono', monospace" }}>{row.sku} · RM{row.price}</p>
                          </div>
                          <span className="text-[10px] font-semibold px-2 py-0.5 rounded-full" style={{ background: row.status === "SOLD" ? "#E4EFEA" : "#EEF3F1", color: row.status === "SOLD" ? "#726F68" : "#17161A" }}>{row.status}</span>
                        </div>
                      ))}
                    </div>
                    <p className="text-xs" style={{ color: "#8B887F" }}>Search by SKU, brand, or category, and export a print-ready stock take to Excel anytime.</p>
                  </div>
                )}

                {landingPreview === "pos" && (
                  <div className="space-y-4">
                    <div className="rounded-lg border p-4" style={{ background: "#FFFFFF", borderColor: "#E4E1D9" }}>
                      <p className="text-xs font-semibold mb-3" style={{ color: "#17161A" }}>Sale</p>
                      {[{ name: "Vintage Shirt", price: 80, discount: 10 }, { name: "Casio Watch", price: 150, discount: 20 }].map((l) => (
                        <div key={l.name} className="flex items-center justify-between text-xs py-1.5 border-b" style={{ borderColor: "#EDEAE2" }}>
                          <span style={{ color: "#17161A" }}>{l.name}</span>
                          <span style={{ fontFamily: "'IBM Plex Mono', monospace", color: "#146658" }}>RM{l.price - l.discount}</span>
                        </div>
                      ))}
                      <div className="flex items-center justify-between text-sm font-semibold pt-2" style={{ color: "#17161A" }}>
                        <span>Total</span><span style={{ fontFamily: "'IBM Plex Mono', monospace" }}>RM200</span>
                      </div>
                    </div>
                    <div className="grid grid-cols-2 sm:grid-cols-4 gap-2">
                      <StatCard icon={TrendingUp} label="Sales (RM)" value={4850} tone="#146658" />
                      <StatCard icon={Receipt} label="Orders" value={37} tone="#17161A" />
                      <StatCard icon={ShoppingBag} label="Units" value={42} tone="#726F68" />
                      <StatCard icon={Percent} label="Discounts" value={320} tone="#B3261E" />
                    </div>
                    <p className="text-xs" style={{ color: "#8B887F" }}>Search a product, negotiate a discount, complete the sale — invoice and stock update automatically.</p>
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
      <div className="min-h-screen w-full flex items-center justify-center px-6" style={{ background: "#F5F4F0", fontFamily: "'Inter', sans-serif" }}>
        <style>{`@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');`}</style>
        <div className="w-full max-w-sm">
          <button onClick={() => setShowAuthForm(false)} className="text-xs font-medium mb-6" style={{ color: "#8B887F" }}>← Back</button>
          <div className="flex items-baseline gap-2 justify-center mb-8">
            <span className="text-2xl font-extrabold tracking-tight" style={{ color: "#17161A" }}>RAVUNO</span>
            <span className="text-[10px] font-medium tracking-wider" style={{ color: "#8B887F" }}>V1</span>
          </div>
          <div className="rounded-xl border p-6" style={{ background: "#FFFFFF", borderColor: "#E4E1D9" }}>
            <h1 className="text-lg font-semibold mb-1" style={{ color: "#17161A" }}>
              {authMode === "signin" ? "Sign in" : "Create your account"}
            </h1>
            <p className="text-xs mb-5" style={{ color: "#8B887F" }}>
              {authMode === "signin" ? "Access your inventory from any device." : "Your data is private to your account only."}
            </p>
            {authNotice && <p className="text-xs px-3 py-2 rounded-lg mb-4" style={{ background: "#EEF3F1", color: "#146658" }}>{authNotice}</p>}
            {authError && <p className="text-xs px-3 py-2 rounded-lg mb-4" style={{ background: "#FBEEEC", color: "#B3261E" }}>{authError}</p>}
            <form onSubmit={handleAuthSubmit} className="space-y-3">
              <Field label="Email" type="email" value={authEmail} onChange={(e) => setAuthEmail(e.target.value)} placeholder="you@example.com" />
              <Field label="Password" type="password" value={authPassword} onChange={(e) => setAuthPassword(e.target.value)} placeholder="••••••••" />
              <button type="submit" disabled={authBusy}
                className="w-full flex items-center justify-center gap-2 py-2.5 rounded-xl text-sm font-semibold transition-opacity"
                style={{ background: "#146658", color: "#FFFFFF", opacity: authBusy ? 0.6 : 1, fontFamily: "'IBM Plex Mono', monospace" }}>
                {authBusy ? <Loader2 size={14} className="animate-spin" /> : null}
                {authMode === "signin" ? "Sign in" : "Sign up"}
              </button>
            </form>
            <button
              onClick={() => { setAuthMode(authMode === "signin" ? "signup" : "signin"); setAuthError(""); setAuthNotice(""); }}
              className="w-full text-center text-xs mt-4" style={{ color: "#726F68" }}>
              {authMode === "signin" ? "New here? Create an account" : "Already have an account? Sign in"}
            </button>
          </div>
        </div>
      </div>
    );
  }

  return (
    <div className="min-h-screen w-full" style={{ background: "#F5F4F0", color: "#17161A", fontFamily: "'Inter', sans-serif" }}>
      <style>{`@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&family=IBM+Plex+Mono:wght@400;500;600&display=swap');`}</style>

      <header style={{ background: "#FFFFFF", borderBottom: "1px solid #E4E1D9" }}>
        <div className="max-w-6xl mx-auto px-6 py-6 relative flex flex-col items-center gap-5">
          <div className="absolute right-6 top-6 flex items-center gap-3">
            {profile?.subscription_status === "active" && (
              <span className="text-xs font-semibold px-2 py-1 rounded-full" style={{ background: "#EEF3F1", color: "#146658" }}>❤️ Supporter</span>
            )}
            <button onClick={handleSignOut} className="text-xs font-medium" style={{ color: "#8B887F" }}>
              Sign out
            </button>
          </div>
          <div className="flex items-baseline gap-2">
            <span className="text-2xl font-extrabold tracking-tight" style={{ color: "#17161A" }}>RAVUNO</span>
            <span className="text-[10px] font-medium tracking-wider" style={{ color: "#8B887F", fontFamily: "'IBM Plex Mono', monospace" }}>V1</span>
          </div>
          <nav className="flex gap-1 p-1 rounded-lg" style={{ background: "#F5F4F0", border: "1px solid #E4E1D9" }}>
            {[{ id: "crosslist", label: "Cross-List" }, { id: "inventory", label: "Stock Ledger" }, { id: "pos", label: "POS" }, { id: "sales", label: "Sales" }, { id: "templates", label: "Templates" }, { id: "support", label: "Support" }].map((t) => (
              <button key={t.id} onClick={() => setTab(t.id)}
                className="px-4 py-1.5 rounded-md text-sm font-medium transition-colors"
                style={{
                  background: tab === t.id ? "#FFFFFF" : "transparent",
                  color: tab === t.id ? "#17161A" : "#726F68",
                  boxShadow: tab === t.id ? "0 1px 2px rgba(23,22,26,0.08)" : "none",
                }}>
                {t.label}
              </button>
            ))}
          </nav>
        </div>
      </header>

      {loading ? (
        <div className="max-w-6xl mx-auto px-6 py-20 flex items-center justify-center gap-2" style={{ color: "#726F68" }}>
          <Loader2 size={16} className="animate-spin" /> Loading your saved items…
        </div>
      ) : (
        <main className="max-w-6xl mx-auto px-6 py-8">
          {saveError && (
            <div className="mb-4 text-xs px-3 py-2 rounded-lg" style={{ background: "#FBEEEC", color: "#B3261E" }}>
              Couldn't save changes — they may not persist. Try again shortly.
            </div>
          )}

          {tab === "crosslist" && (
            <div className="grid lg:grid-cols-[380px_1fr] gap-8">
              <section>
                <div className="rounded-xl border p-6" style={{ background: "#FFFFFF", borderColor: "#E4E1D9" }}>
                  <h2 className="text-lg font-semibold mb-4" style={{ color: "#17161A" }}>Add a product</h2>
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
                      <p className="text-xs px-3 py-2 rounded-lg" style={{ background: "#EEF3F1", color: "#726F68" }}>
                        Select a category to see relevant fields.
                      </p>
                    )}

                    {form.category && (
                      <>
                        <div className="border-t pt-3 mt-1" style={{ borderColor: "#EDEAE2" }}>
                          <Field label="Product title" value={form.title} onChange={(e) => setForm((f) => ({ ...f, title: e.target.value, titleTouched: true }))} placeholder="Auto from brand + item type" />
                        </div>
                        <div className="grid grid-cols-2 gap-3">
                          <Field label="Brand" value={form.brand} onChange={set("brand")} placeholder="Uniqlo" />
                          <Field label="Condition (/10)" value={form.condition} onChange={set("condition")} placeholder="9" type="number" />
                        </div>
                        <Field label="Colour" value={form.colour} onChange={set("colour")} placeholder="Black" />
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
                    style={{ background: "#146658", color: "#FFFFFF", opacity: isFormUsable ? 1 : 0.35, fontFamily: "'IBM Plex Mono', monospace" }}>
                    <Plus size={15} /> Save listing & generate
                  </button>
                </div>

                {items.length > 0 && (
                  <div className="mt-6">
                    <p className="text-xs tracking-[0.2em] uppercase mb-2" style={{ color: "#726F68", fontFamily: "'IBM Plex Mono', monospace" }}>Saved</p>
                    <div className="space-y-2">
                      {items.map((l) => (
                        <button key={l.id} onClick={() => setActiveId(l.id)}
                          className="w-full flex items-center justify-between px-3 py-2 rounded-xl border text-left transition-colors"
                          style={{ borderColor: activeId === l.id ? "#146658" : "#E4E1D9", background: activeId === l.id ? "#FFFFFF" : "transparent" }}>
                          <span className="flex items-center gap-2 min-w-0">
                            {l.photos && l.photos[0] ? (
                              <img src={l.photos[0]} alt="" className="w-8 h-8 rounded object-cover shrink-0" />
                            ) : (
                              <ChevronRight size={14} className="shrink-0" style={{ color: "#726F68" }} />
                            )}
                            <span className="truncate text-sm font-medium">{productTitle(l)}</span>
                          </span>
                          <span className="flex items-center gap-2 shrink-0">
                            <span className="text-xs" style={{ color: "#726F68", fontFamily: "'IBM Plex Mono', monospace" }}>{CATEGORY_LABELS[l.category]}</span>
                            <Trash2 size={14} style={{ color: "#B3261E" }} onClick={(e) => { e.stopPropagation(); removeItem(l.id); }} />
                          </span>
                        </button>
                      ))}
                    </div>
                  </div>
                )}
              </section>

              <section>
                <p className="text-xs tracking-wide uppercase mb-3 font-medium" style={{ color: "#8B887F" }}>
                  {activeListing ? `Ready to paste — ${productTitle(activeListing)}` : "Live preview — fill in the product on the left"}
                </p>
                <div className="grid sm:grid-cols-2 gap-4">
                  {PLATFORMS.map((p) => (
                    <div key={p.id} className="rounded-xl border overflow-hidden bg-white" style={{ borderColor: "#E4E1D9" }}>
                      <div className="flex items-center justify-between px-4 py-3 border-b" style={{ borderColor: "#EDEAE2" }}>
                        <span className="text-xs font-semibold px-2 py-1 rounded-md" style={{ background: "#F5F4F0", color: "#17161A" }}>{p.label}</span>
                        <CopyButton text={activeTemplates[p.id]} />
                      </div>
                      <pre className="text-xs px-4 py-3 whitespace-pre-wrap leading-relaxed" style={{ fontFamily: "'IBM Plex Mono', monospace", color: "#17161A" }}>
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
                <StatCard icon={Package} label="Listed" value={stats.listedCount} tone="#17161A" />
                <StatCard icon={TrendingUp} label="Sold" value={stats.soldCount} tone="#726F68" />
                <StatCard icon={Clock} label="Stale 21d+" value={stats.staleCount} tone="#B3261E" />
              </div>

              <div>
                <div className="flex items-center justify-between mb-3 flex-wrap gap-2">
                  <h2 className="text-base font-semibold" style={{ color: "#17161A" }}>Stock Items</h2>
                  <div className="flex items-center gap-2">
                    <button onClick={() => exportStockTake("listed")}
                      className="flex items-center gap-1.5 text-xs font-semibold px-3 py-2 rounded-lg border"
                      style={{ borderColor: "#17161A", color: "#17161A", fontFamily: "'IBM Plex Mono', monospace" }}>
                      <Download size={13} /> Export Listed Only
                    </button>
                    <button onClick={() => exportStockTake("all")}
                      className="flex items-center gap-1.5 text-xs font-semibold px-3 py-2 rounded-lg"
                      style={{ background: "#146658", color: "#FFFFFF", fontFamily: "'IBM Plex Mono', monospace" }}>
                      <Download size={13} /> Export All Stock
                    </button>
                  </div>
                </div>

                <div className="relative mb-4">
                  <Search size={15} className="absolute left-3 top-1/2 -translate-y-1/2" style={{ color: "#8B887F" }} />
                  <input value={ledgerSearch} onChange={(e) => setLedgerSearch(e.target.value)} placeholder="Search SKU, brand, item…"
                    className="w-full pl-9 pr-9 py-2.5 rounded-lg border text-sm outline-none" style={{ borderColor: "#E4E1D9", background: "#FFFFFF", color: "#17161A" }} />
                  {ledgerSearch && (
                    <button onClick={() => setLedgerSearch("")} className="absolute right-3 top-1/2 -translate-y-1/2" aria-label="Clear search">
                      <X size={15} style={{ color: "#8B887F" }} />
                    </button>
                  )}
                </div>

                <div className="hidden sm:grid gap-2 px-4 mb-1 text-[11px] uppercase tracking-wide"
                  style={{ color: "#726F68", fontFamily: "'IBM Plex Mono', monospace", gridTemplateColumns: "1fr 1.6fr 1.1fr 1.1fr 0.7fr 0.7fr 0.8fr" }}>
                  <span>SKU</span><span>Title</span><span>Category</span><span>Item Type</span><span>Cost</span><span>Price</span><span>Status</span>
                </div>

                <div className="space-y-2">
                  {ledgerItems.map((i) => {
                    const stale = i.status === "listed" && daysSince(i.addedOn) >= 21;
                    return (
                      <div key={i.id} onClick={() => { setDetailId(i.id); setGalleryIndex(0); }}
                        className="rounded-xl border px-4 py-3 cursor-pointer transition-colors"
                        style={{ borderColor: i.status === "sold" ? "#726F68" : stale ? "#B3261E" : "#E4E1D9", background: "#FFFFFF" }}>

                        <div className="hidden sm:grid items-center gap-2 text-sm" style={{ gridTemplateColumns: "1fr 1.6fr 1.1fr 1.1fr 0.7fr 0.7fr 0.8fr" }}>
                          <span className="truncate" style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{i.sku || "-"}</span>
                          <span className="truncate font-medium">{productTitle(i)}</span>
                          <span className="truncate">{CATEGORY_LABELS[i.category] || "-"}</span>
                          <span className="truncate">{i.itemType || "-"}</span>
                          <span style={{ fontFamily: "'IBM Plex Mono', monospace" }}>RM{i.cost || 0}</span>
                          <span style={{ fontFamily: "'IBM Plex Mono', monospace" }}>RM{i.price || 0}</span>
                          <span className="text-xs font-semibold px-2 py-0.5 rounded-full w-fit"
                            style={{ background: i.status === "sold" ? "#E4EFEA" : "#EEF3F1", color: i.status === "sold" ? "#726F68" : "#17161A", fontFamily: "'IBM Plex Mono', monospace" }}>
                            {(i.status || "listed").toUpperCase()}
                          </span>
                        </div>

                        <div className="sm:hidden flex items-center gap-3">
                          {i.photos && i.photos[0] ? (
                            <img src={i.photos[0]} alt="" className="w-10 h-10 rounded-lg object-cover shrink-0" />
                          ) : (
                            <div className="w-10 h-10 rounded-lg shrink-0 flex items-center justify-center border-2 border-dashed" style={{ borderColor: "#E4E1D9", color: "#E4E1D9" }}>
                              <ImageIcon size={14} />
                            </div>
                          )}
                          <div className="min-w-0 flex-1">
                            <p className="text-sm font-semibold truncate">{productTitle(i)}</p>
                            <p className="text-xs" style={{ color: "#726F68", fontFamily: "'IBM Plex Mono', monospace" }}>
                              {i.sku || "-"} · {CATEGORY_LABELS[i.category] || "-"} · RM{i.price || 0}
                            </p>
                          </div>
                          <span className="text-xs font-semibold px-2 py-0.5 rounded-full shrink-0"
                            style={{ background: i.status === "sold" ? "#E4EFEA" : "#EEF3F1", color: i.status === "sold" ? "#726F68" : "#17161A", fontFamily: "'IBM Plex Mono', monospace" }}>
                            {(i.status || "listed").toUpperCase()}
                          </span>
                        </div>

                        <div className="flex items-center gap-2 mt-2 pt-2 border-t" style={{ borderColor: "#EDEAE2" }} onClick={(e) => e.stopPropagation()}>
                          <Trash2 size={14} style={{ color: "#B3261E", cursor: "pointer" }} onClick={() => removeItem(i.id)} />
                        </div>
                      </div>
                    );
                  })}
                  {ledgerItems.length === 0 && (
                    <p className="text-sm text-center py-8" style={{ color: "#8B887F" }}>
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
                <div className="rounded-xl border p-6" style={{ background: "#FFFFFF", borderColor: "#E4E1D9" }}>
                  <div className="relative">
                    <Search size={17} className="absolute left-4 top-1/2 -translate-y-1/2" style={{ color: "#8B887F" }} />
                    <input
                      value={skuSearch}
                      onChange={(e) => setSkuSearch(e.target.value)}
                      placeholder="Search product name or SKU…"
                      autoFocus
                      className="w-full pl-11 pr-4 py-3 rounded-lg border text-base outline-none"
                      style={{ borderColor: "#E4E1D9", background: "#FFFFFF", color: "#17161A" }}
                    />
                  </div>
                </div>

                {skuSearch.trim() === "" ? (
                  <p className="text-sm text-center py-10" style={{ color: "#8B887F" }}>
                    Search by product name or SKU to find an item to sell.
                  </p>
                ) : (
                  <div className="mt-4 space-y-2">
                    {searchResults.map((i) => (
                      <button key={i.id} onClick={() => addToCart(i)}
                        className="w-full flex items-center justify-between gap-3 px-4 py-3 rounded-xl border text-left" style={{ borderColor: "#E4E1D9", background: "#FFFFFF" }}>
                        <div className="flex items-center gap-3 min-w-0">
                          {i.photos && i.photos[0] ? (
                            <img src={i.photos[0]} alt="" className="w-10 h-10 rounded-lg object-cover shrink-0" />
                          ) : (
                            <div className="w-10 h-10 rounded-lg shrink-0 flex items-center justify-center border-2 border-dashed" style={{ borderColor: "#E4E1D9", color: "#E4E1D9" }}>
                              <ImageIcon size={14} />
                            </div>
                          )}
                          <div className="min-w-0">
                            <p className="text-sm font-semibold truncate">{productTitle(i)}</p>
                            <p className="text-xs" style={{ color: "#726F68", fontFamily: "'IBM Plex Mono', monospace" }}>{i.sku || "no sku"} · RM{i.price || 0}</p>
                          </div>
                        </div>
                        <Plus size={16} style={{ color: "#17161A" }} className="shrink-0" />
                      </button>
                    ))}
                    {searchResults.length === 0 && <p className="text-sm text-center py-8" style={{ color: "#8B887F" }}>No products found.</p>}
                  </div>
                )}
              </section>

              <section>
                <div className="rounded-xl border p-5 sticky top-4" style={{ background: "#FFFFFF", borderColor: "#E4E1D9" }}>
                  <h2 className="text-base font-semibold mb-3" style={{ color: "#17161A" }}>Sale</h2>
                  <div className="space-y-3 mb-4">
                    {cartLines.map((l) => (
                      <div key={l.id} className="pb-3 border-b" style={{ borderColor: "#EDEAE2" }}>
                        <div className="flex items-center justify-between gap-2 mb-2">
                          <p className="truncate text-sm font-medium">{productTitle(l)}</p>
                          <X size={14} style={{ color: "#B3261E", cursor: "pointer" }} className="shrink-0" onClick={() => removeFromCart(l.id)} />
                        </div>
                        <div className="grid grid-cols-3 gap-2 items-end">
                          <label className="block">
                            <span className="block text-[10px] mb-1 uppercase tracking-wide" style={{ color: "#8B887F" }}>Qty</span>
                            <input type="number" min="1" value={l.qty} onChange={(e) => updateQty(l.id, e.target.value)}
                              className="w-full px-2 py-1.5 rounded-md border text-xs text-center outline-none" style={{ borderColor: "#E4E1D9", fontFamily: "'IBM Plex Mono', monospace" }} />
                          </label>
                          <label className="block">
                            <span className="block text-[10px] mb-1 uppercase tracking-wide" style={{ color: "#8B887F" }}>Discount (RM)</span>
                            <input type="number" min="0" value={l.discount || ""} onChange={(e) => updateDiscount(l.id, e.target.value)} placeholder="0"
                              className="w-full px-2 py-1.5 rounded-md border text-xs text-center outline-none" style={{ borderColor: "#E4E1D9", fontFamily: "'IBM Plex Mono', monospace" }} />
                          </label>
                          <div>
                            <span className="block text-[10px] mb-1 uppercase tracking-wide" style={{ color: "#8B887F" }}>Final</span>
                            <p className="text-xs font-semibold text-center py-1.5" style={{ fontFamily: "'IBM Plex Mono', monospace", color: "#146658" }}>RM{l.finalPrice}</p>
                          </div>
                        </div>
                        {l.discount > 0 && (
                          <p className="text-[11px] mt-1 text-right" style={{ color: "#8B887F", fontFamily: "'IBM Plex Mono', monospace" }}>
                            RM{l.originalPrice} − RM{l.discount}
                          </p>
                        )}
                      </div>
                    ))}
                    {cartLines.length === 0 && <p className="text-xs" style={{ color: "#8B887F" }}>No items yet — search and select a product to begin a sale.</p>}
                  </div>

                  <div className="space-y-1 pt-1 pb-3 border-b mb-3 text-sm" style={{ borderColor: "#E4E1D9" }}>
                    <div className="flex items-center justify-between" style={{ color: "#726F68" }}>
                      <span>Subtotal</span><span style={{ fontFamily: "'IBM Plex Mono', monospace" }}>RM{cartSubtotal}</span>
                    </div>
                    <div className="flex items-center justify-between" style={{ color: "#726F68" }}>
                      <span>Discount</span><span style={{ fontFamily: "'IBM Plex Mono', monospace" }}>− RM{cartDiscount}</span>
                    </div>
                    <div className="flex items-center justify-between font-semibold pt-1" style={{ color: "#17161A" }}>
                      <span>Total</span><span style={{ fontFamily: "'IBM Plex Mono', monospace" }}>RM{cartTotal}</span>
                    </div>
                  </div>

                  <Select label="Payment method" value={payMethod} onChange={(e) => setPayMethod(e.target.value)}>
                    <option>Cash</option><option>Bank Transfer</option><option>Online</option>
                  </Select>
                  <button onClick={completeSale} disabled={cartLines.length === 0}
                    className="mt-4 w-full flex items-center justify-center gap-2 py-2.5 rounded-xl text-sm font-semibold transition-opacity"
                    style={{ background: "#146658", color: "#FFFFFF", opacity: cartLines.length === 0 ? 0.35 : 1, fontFamily: "'IBM Plex Mono', monospace" }}>
                    Complete Sale
                  </button>
                </div>
              </section>
            </div>
          )}

          {tab === "sales" && (
            <div className="space-y-6">
              {invoicedTxns.length === 0 ? (
                <div className="rounded-xl border p-10 text-center" style={{ background: "#FFFFFF", borderColor: "#E4E1D9" }}>
                  <Receipt size={28} className="mx-auto mb-3" style={{ color: "#8B887F" }} />
                  <p className="text-sm font-semibold mb-1" style={{ color: "#17161A" }}>No sales data yet</p>
                  <p className="text-xs" style={{ color: "#8B887F" }}>Complete your first sale to start tracking performance.</p>
                </div>
              ) : (
                <>
                  {/* period selector */}
                  <div className="flex items-center gap-2 flex-wrap">
                    <div className="flex gap-1 p-1 rounded-lg" style={{ background: "#FFFFFF", border: "1px solid #E4E1D9" }}>
                      {[{ id: "monthly", label: "Monthly" }, { id: "yearly", label: "Yearly" }].map((p) => (
                        <button key={p.id} onClick={() => setSalesPeriodType(p.id)}
                          className="px-3 py-1.5 rounded-md text-xs font-medium transition-colors"
                          style={{ background: salesPeriodType === p.id ? "#146658" : "transparent", color: salesPeriodType === p.id ? "#FFFFFF" : "#726F68" }}>
                          {p.label}
                        </button>
                      ))}
                    </div>
                    {salesPeriodType === "monthly" && (
                      <select value={salesMonth} onChange={(e) => setSalesMonth(Number(e.target.value))}
                        className="px-3 py-2 rounded-lg border text-sm outline-none" style={{ borderColor: "#E4E1D9", background: "#FFFFFF", color: "#17161A" }}>
                        {MONTH_NAMES.map((m, idx) => <option key={m} value={idx}>{m}</option>)}
                      </select>
                    )}
                    <select value={salesYear} onChange={(e) => setSalesYear(Number(e.target.value))}
                      className="px-3 py-2 rounded-lg border text-sm outline-none" style={{ borderColor: "#E4E1D9", background: "#FFFFFF", color: "#17161A" }}>
                      {availableSalesYears.map((y) => <option key={y} value={y}>{y}</option>)}
                    </select>
                  </div>

                  {/* summary metrics */}
                  <div className="grid grid-cols-2 sm:grid-cols-4 gap-3">
                    <StatCard icon={TrendingUp} label="Total Sales (RM)" value={salesSummary.totalSales} tone="#146658" />
                    <StatCard icon={Receipt} label="Orders" value={salesSummary.orders} tone="#17161A" />
                    <StatCard icon={ShoppingBag} label="Units Sold" value={salesSummary.unitsSold} tone="#726F68" />
                    <StatCard icon={Percent} label="Discounts (RM)" value={salesSummary.discounts} tone="#B3261E" />
                  </div>

                  {/* revenue graph */}
                  <div className="rounded-xl border p-5" style={{ background: "#FFFFFF", borderColor: "#E4E1D9" }}>
                    <h2 className="text-sm font-semibold mb-4" style={{ color: "#17161A" }}>
                      Revenue — {salesPeriodType === "monthly" ? `${MONTH_NAMES[salesMonth]} ${salesYear}` : salesYear}
                    </h2>
                    <div style={{ width: "100%", height: 220 }}>
                      <ResponsiveContainer>
                        <LineChart data={salesGraphData} margin={{ top: 5, right: 10, bottom: 0, left: 0 }}>
                          <CartesianGrid stroke="#EDEAE2" vertical={false} />
                          <XAxis dataKey="label" tick={{ fontSize: 11, fill: "#8B887F" }} axisLine={{ stroke: "#E4E1D9" }} tickLine={false} interval={salesPeriodType === "monthly" ? 2 : 0} />
                          <YAxis tick={{ fontSize: 11, fill: "#8B887F" }} axisLine={false} tickLine={false} width={44} />
                          <Tooltip
                            contentStyle={{ background: "#17161A", border: "none", borderRadius: 8, fontSize: 12 }}
                            labelStyle={{ color: "#FFFFFF" }}
                            itemStyle={{ color: "#FFFFFF" }}
                            formatter={(value) => [`RM${value}`, "Revenue"]}
                          />
                          <Line type="monotone" dataKey="revenue" stroke="#146658" strokeWidth={2} dot={false} activeDot={{ r: 4 }} />
                        </LineChart>
                      </ResponsiveContainer>
                    </div>
                  </div>

                  {/* invoice list */}
                  <div>
                    <h2 className="text-sm font-semibold mb-3" style={{ color: "#17161A" }}>Invoices</h2>
                    <div className="space-y-2">
                      {periodTxns.map((t) => (
                        <button key={t.id} onClick={() => setSelectedInvoiceId(t.id)}
                          className="w-full flex items-center justify-between gap-3 px-4 py-3 rounded-xl border text-left" style={{ borderColor: "#E4E1D9", background: "#FFFFFF" }}>
                          <div className="min-w-0">
                            <div className="flex items-center gap-2 mb-0.5">
                              <span className="text-xs font-semibold" style={{ fontFamily: "'IBM Plex Mono', monospace", color: "#17161A" }}>{t.invoiceNumber}</span>
                              <span className="text-xs" style={{ color: "#8B887F" }}>
                                {new Date(t.timestamp).toLocaleDateString(undefined, { day: "2-digit", month: "short", year: "numeric" })} · {new Date(t.timestamp).toLocaleTimeString(undefined, { hour: "2-digit", minute: "2-digit" })}
                              </span>
                            </div>
                            <p className="text-sm truncate" style={{ color: "#726F68" }}>{t.lines.map((l) => l.title).join(", ")}</p>
                          </div>
                          <div className="text-right shrink-0" style={{ fontFamily: "'IBM Plex Mono', monospace" }}>
                            <p className="text-sm font-semibold" style={{ color: "#146658" }}>RM{t.total}</p>
                            {t.discount > 0 && <p className="text-[11px]" style={{ color: "#8B887F" }}>−RM{t.discount} disc.</p>}
                          </div>
                        </button>
                      ))}
                      {periodTxns.length === 0 && (
                        <p className="text-sm text-center py-8" style={{ color: "#8B887F" }}>No invoices in this period.</p>
                      )}
                    </div>
                  </div>
                </>
              )}
            </div>
          )}

          {tab === "templates" && (
            <TemplatesTab customTemplates={customTemplates} onSave={saveTemplates} saving={templatesSaving} />
          )}

          {tab === "support" && (
            <div className="max-w-xl mx-auto">
              <div className="rounded-xl border p-8 text-center" style={{ background: "#FFFFFF", borderColor: "#E4E1D9" }}>
                <p className="text-3xl mb-4">👋</p>
                <h2 className="text-lg font-semibold mb-3" style={{ color: "#17161A" }}>Why Ravuno is free</h2>
                <p className="text-sm leading-relaxed mb-2" style={{ color: "#726F68" }}>
                  I built Ravuno because I know what it's like juggling Carousell, Shopee, and a messy spreadsheet just to
                  keep a small resale business running. I wanted a tool that actually helps sellers like us — without a
                  paywall in the way.
                </p>
                <p className="text-sm leading-relaxed mb-6" style={{ color: "#726F68" }}>
                  Ravuno is free to use, fully and always. If it's genuinely helped your business and you'd like to support
                  the time that goes into building and maintaining it, that's completely optional — and appreciated.
                </p>

                <div className="grid grid-cols-3 gap-3 mb-4">
                  {[5, 10, 20].map((amt) => (
                    <button key={amt} onClick={() => handleSupport(amt)} disabled={upgradeBusy}
                      className="py-3 rounded-lg text-sm font-semibold border transition-opacity"
                      style={{ borderColor: "#146658", color: "#146658", opacity: upgradeBusy ? 0.6 : 1, fontFamily: "'IBM Plex Mono', monospace" }}>
                      RM{amt}
                    </button>
                  ))}
                </div>
                <p className="text-xs" style={{ color: "#8B887F" }}>One-time, no account changes, no subscription — just a thank-you.</p>

                <div className="mt-6 pt-6 border-t" style={{ borderColor: "#EDEAE2" }}>
                  <p className="text-xs mb-1" style={{ color: "#8B887F" }}>Questions or feedback:</p>
                  <a href="mailto:Unknownlable00@gmail.com" className="text-sm font-medium" style={{ color: "#146658" }}>Unknownlable00@gmail.com</a>
                </div>
              </div>
            </div>
          )}
        </main>
      )}

      {detailItem && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4" style={{ background: "rgba(36,31,26,0.55)" }} onClick={() => setDetailId(null)}>
          <div className="max-w-md w-full rounded-xl border p-6 relative max-h-[85vh] overflow-y-auto shadow-sm" style={{ background: "#FFFFFF", borderColor: "#E4E1D9" }} onClick={(e) => e.stopPropagation()}>
            <button onClick={() => setDetailId(null)} className="absolute top-4 right-4 w-7 h-7 rounded-full flex items-center justify-center" style={{ background: "#F7F6F2", color: "#17161A" }} aria-label="Close">
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
                    className="w-full h-56 object-cover rounded-xl border-2 cursor-pointer" style={{ borderColor: "#17161A" }}
                  />
                  {photos.length > 1 && (
                    <div className="flex gap-2 mt-2 overflow-x-auto pb-1">
                      {photos.map((p, idx) => (
                        <img key={idx} src={p} alt="" onClick={() => setGalleryIndex(idx)}
                          className="w-12 h-12 rounded-lg object-cover shrink-0 cursor-pointer border-2"
                          style={{ borderColor: idx === galleryIndex ? "#146658" : "#E4E1D9" }} />
                      ))}
                    </div>
                  )}
                  <div className="flex items-center gap-2 mt-3 flex-wrap">
                    <button onClick={() => downloadDataUrl(photos[galleryIndex], `${sanitizeForFilename(detailItem.sku)}_${String(galleryIndex + 1).padStart(2, "0")}.jpg`)}
                      className="flex items-center gap-1.5 text-xs font-medium px-3 py-1.5 rounded-lg border" style={{ borderColor: "#17161A", color: "#17161A", fontFamily: "'IBM Plex Mono', monospace" }}>
                      <Download size={13} /> Download photo
                    </button>
                    {photos.length > 1 && (
                      <button onClick={() => downloadAllPhotos(detailItem)}
                        className="flex items-center gap-1.5 text-xs font-medium px-3 py-1.5 rounded-lg" style={{ background: "#146658", color: "#FFFFFF", fontFamily: "'IBM Plex Mono', monospace" }}>
                        <Download size={13} /> Download all ({photos.length})
                      </button>
                    )}
                  </div>
                </div>
              ) : (
                <div className="w-full h-40 rounded-xl border-2 border-dashed flex items-center justify-center mb-4" style={{ borderColor: "#E4E1D9", color: "#E4E1D9" }}>
                  <ImageIcon size={28} />
                </div>
              );
            })()}

            <h3 className="text-xl font-semibold mb-1" style={{ color: "#17161A" }}>{productTitle(detailItem)}</h3>
            <div className="flex items-center gap-2 mb-4 flex-wrap">
              <span className="inline-flex items-center gap-1 text-xs font-semibold px-2 py-0.5 rounded-full" style={{ background: "#EEF3F1", color: "#17161A", fontFamily: "'IBM Plex Mono', monospace" }}>
                <Tag size={11} /> {CATEGORY_LABELS[detailItem.category] || "—"}
              </span>
              <span className="inline-block text-xs font-semibold px-2 py-0.5 rounded-full"
                style={{ background: detailItem.status === "sold" ? "#E4EFEA" : "#EEF3F1", color: detailItem.status === "sold" ? "#726F68" : "#17161A", fontFamily: "'IBM Plex Mono', monospace" }}>
                {(detailItem.status || "listed").toUpperCase()}
              </span>
            </div>

            <dl className="space-y-2 text-sm">
              <DetailRow label="Item type" value={detailItem.itemType || "—"} />
              <DetailRow label="Brand" value={detailItem.brand || "—"} />
              <DetailRow label="Colour" value={detailItem.colour || "—"} />

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

            <button
              onClick={() => { setActiveId(detailItem.id); setTab("crosslist"); setDetailId(null); }}
              className="mt-4 w-full flex items-center justify-center gap-2 py-2.5 rounded-xl text-sm font-semibold"
              style={{ background: "#146658", color: "#FFFFFF", fontFamily: "'IBM Plex Mono', monospace" }}>
              Open in Cross-List
            </button>
          </div>
        </div>
      )}

      {selectedInvoice && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4" style={{ background: "rgba(23,22,26,0.55)" }} onClick={() => setSelectedInvoiceId(null)}>
          <div className="max-w-md w-full rounded-xl border p-6 relative max-h-[85vh] overflow-y-auto shadow-sm" style={{ background: "#FFFFFF", borderColor: "#E4E1D9" }} onClick={(e) => e.stopPropagation()}>
            <button onClick={() => setSelectedInvoiceId(null)} className="absolute top-4 right-4 w-7 h-7 rounded-full flex items-center justify-center" style={{ background: "#F7F6F2", color: "#17161A" }} aria-label="Close">
              <X size={15} />
            </button>

            <p className="text-xs font-semibold mb-1" style={{ fontFamily: "'IBM Plex Mono', monospace", color: "#146658" }}>{selectedInvoice.invoiceNumber}</p>
            <h3 className="text-lg font-semibold mb-1" style={{ color: "#17161A" }}>
              {new Date(selectedInvoice.timestamp).toLocaleDateString(undefined, { day: "2-digit", month: "long", year: "numeric" })}
            </h3>
            <p className="text-xs mb-4" style={{ color: "#8B887F" }}>
              {new Date(selectedInvoice.timestamp).toLocaleTimeString(undefined, { hour: "2-digit", minute: "2-digit" })} · {selectedInvoice.method}
            </p>

            <div className="space-y-3 mb-4">
              {selectedInvoice.lines.map((l, idx) => {
                const original = Number(l.price || 0) * Number(l.qty || 0);
                return (
                  <div key={idx} className="pb-3 border-b" style={{ borderColor: "#EDEAE2" }}>
                    <p className="text-sm font-medium mb-1">{l.title}</p>
                    <div className="grid grid-cols-4 gap-2 text-xs" style={{ color: "#726F68" }}>
                      <div><span className="block text-[10px] uppercase" style={{ color: "#8B887F" }}>Qty</span>{l.qty}</div>
                      <div><span className="block text-[10px] uppercase" style={{ color: "#8B887F" }}>Original</span>RM{original}</div>
                      <div><span className="block text-[10px] uppercase" style={{ color: "#8B887F" }}>Discount</span>RM{l.discount || 0}</div>
                      <div><span className="block text-[10px] uppercase" style={{ color: "#8B887F" }}>Final</span><span style={{ color: "#146658", fontWeight: 600 }}>RM{l.finalPrice}</span></div>
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
