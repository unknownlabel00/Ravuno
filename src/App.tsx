import React, { useState, useEffect, useMemo, useRef } from "react";
import {
  Check, Copy, Ruler, Plus, Trash2, ChevronRight, ChevronLeft,
  Package, TrendingUp, Clock, Image as ImageIcon, X, Loader2,
  Search, Download, Star, Tag, Receipt, ShoppingBag, Percent, DollarSign, Eye, MessageCircle, Users, Link2,
  User, ArrowRight, Shirt, Recycle, ShieldCheck, Truck, SlidersHorizontal
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

// customers (shoppers) get their own session slot, so a customer login can never overwrite or become an admin login
const SHOPPER_SESSION_KEY = "ravuno_shopper_session";
function loadShopperSession() {
  try { return JSON.parse(localStorage.getItem(SHOPPER_SESSION_KEY)); } catch (e) { return null; }
}
function saveShopperSession(session) {
  if (session) localStorage.setItem(SHOPPER_SESSION_KEY, JSON.stringify(session));
  else localStorage.removeItem(SHOPPER_SESSION_KEY);
}
const shopperSessionFrom = (d) => ({ access_token: d.access_token, refresh_token: d.refresh_token, user: d.user });
// are customer accounts open yet? (a yes/no flag the database exposes without revealing any settings)
async function shopperAccountsOpen() {
  try {
    const res = await fetch(`${REST_URL}/rpc/shopper_accounts_open`, { method: "POST", headers: { "Content-Type": "application/json", apikey: SUPABASE_ANON_KEY }, body: "{}" });
    if (!res.ok) return false;
    return (await res.json()) === true;
  } catch (e) { return false; }
}
async function checkoutIsOpen() {
  try {
    const res = await fetch(`${REST_URL}/rpc/checkout_is_open`, { method: "POST", headers: { "Content-Type": "application/json", apikey: SUPABASE_ANON_KEY }, body: "{}" });
    if (!res.ok) return false;
    return (await res.json()) === true;
  } catch (e) { return false; }
}
async function authUserRequest(accessToken) {
  const res = await fetch(`${AUTH_URL}/user`, { headers: { apikey: SUPABASE_ANON_KEY, Authorization: `Bearer ${accessToken}` } });
  const data = await res.json().catch(() => ({}));
  if (!res.ok) throw new Error(data.msg || "invalid link");
  return data;
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

// loads a logo (a picked file or an existing web address) and returns a PNG data URL, with the background made transparent when it can be
async function logoToPng(source, { transparent = true, strength = "normal", maxWidth = 600 } = {}) {
  const src = typeof source === "string" ? (source.startsWith("data:") ? source : source + (source.includes("?") ? "&" : "?") + "cb=" + Date.now()) : await new Promise((res, rej) => { const r = new FileReader(); r.onload = (e) => res(e.target.result); r.onerror = rej; r.readAsDataURL(source); });
  const img = await new Promise((res, rej) => { const i = new Image(); if (typeof source === "string") i.crossOrigin = "anonymous"; i.onload = () => res(i); i.onerror = () => rej(new Error("Could not load the logo")); i.src = src; });
  const scale = Math.min(1, maxWidth / img.width);
  const canvas = document.createElement("canvas");
  canvas.width = Math.max(1, Math.round(img.width * scale)); canvas.height = Math.max(1, Math.round(img.height * scale));
  const ctx = canvas.getContext("2d");
  ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
  let result = { ok: false, removed: 0 };
  if (transparent) {
    const data = ctx.getImageData(0, 0, canvas.width, canvas.height);
    result = removeBackground(data.data, canvas.width, canvas.height, LOGO_STRENGTH[strength] || LOGO_STRENGTH.normal);
    if (result.ok) ctx.putImageData(data, 0, 0);
  }
  return { dataUrl: canvas.toDataURL("image/png"), result };
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
    sale: row.sale || null, reservedUntil: row.reserved_until || null, reservedOrder: row.reserved_order || null,
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
  return { id: row.id, timestamp: row.occurred_at, method: row.method, subtotal: row.subtotal, discount: row.discount, total: row.total, postage: Number(row.postage || 0), lines: row.lines || [], customerId: row.customer_id || null, trackingNo: row.tracking_no || "", shippedAt: row.shipped_at || null, postageBy: row.postage_by === "me" ? "me" : "customer", postageCost: row.postage_cost == null ? null : Number(row.postage_cost) };
}
function txnToRow(txn, userId) {
  return { user_id: userId, occurred_at: txn.timestamp, method: txn.method, subtotal: txn.subtotal, discount: txn.discount, total: txn.total, postage: Number(txn.postage || 0), lines: txn.lines, customer_id: txn.customerId || null, postage_by: txn.postageBy === "me" ? "me" : "customer" };
}

function describeSkuItems(list) {
  const shown = list.slice(0, 2).map((i) => `${i.sku} (“${productTitle(i)}”${i.status === "sold" ? ", sold" : ""})`).join(", ");
  return list.length > 2 ? `${shown} +${list.length - 2} more` : shown;
}
function rowToCustomer(r) {
  return { id: r.id, name: r.name || "", phone: r.phone, createdAt: r.created_at };
}
function rowToPurchase(row) {
  return {
    id: row.id, cost: row.cost ?? "", place: row.place || "",
    photo: row.photo || "", receiptPhoto: row.receipt_photo || "",
    purchasedOn: row.purchased_on, createdAt: row.created_at,
  };
}
function purchaseToRow(p, userId) {
  return {
    user_id: userId, cost: p.cost === "" ? null : p.cost, place: p.place || null,
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
const APPAREL_TOP_TYPES = ["T-Shirt", "Shirt", "Polo", "Hoodie", "Jacket", "Sweatshirt", "Sweater", "Other topwear"];
const APPAREL_BOTTOM_TYPES = ["Jeans", "Pants", "Shorts", "Skirt", "Dress", "Other bottomwear"];
// "Crewneck" is just a "Sweatshirt"; anything still saved or imported under the old name is read as the new one ("Sweater" is its own type again)
const normalizeItemType = (t) => (String(t || "").trim().toLowerCase() === "crewneck" ? "Sweatshirt" : String(t || "").trim());
// ITEM-TYPES-END

// CATEGORY-LOGIC-START
// Customers browse by "shop category". Apparel stays the internal parent category: each apparel piece's shop category
// is worked out from its item type, so nothing stored in the database has to change.
const SHOP_CATEGORIES = [
  { id: "tshirts", label: "T-Shirts", parent: "apparel", types: ["T-Shirt"] },
  { id: "shirts", label: "Shirts", parent: "apparel", types: ["Shirt", "Polo"] },
  { id: "hoodies", label: "Hoodies", parent: "apparel", types: ["Hoodie"] },
  { id: "sweatshirts", label: "Sweatshirts", parent: "apparel", types: ["Sweatshirt"] },
  { id: "jackets", label: "Jackets", parent: "apparel", types: ["Jacket"] },
  { id: "sweaters", label: "Sweaters", parent: "apparel", types: ["Sweater"] },
  { id: "pants", label: "Pants", parent: "apparel", types: ["Pants"] },
  { id: "jeans", label: "Jeans", parent: "apparel", types: ["Jeans"] },
  { id: "shorts", label: "Shorts", parent: "apparel", types: ["Shorts"] },
  { id: "other-clothing", label: "Other Clothing", parent: "apparel", types: [] }, // everything else that is apparel
  { id: "footwear", label: "Footwear", parent: "footwear" },
  { id: "bags", label: "Bags", parent: "bags" },
  { id: "watches", label: "Watches", parent: "watches" },
  { id: "electronics", label: "Electronics", parent: "electronics" },
  { id: "home", label: "Home & Living", parent: "home" },
  { id: "books", label: "Books", parent: "books" },
  { id: "games", label: "Games & Collectibles", parent: "games" },
  { id: "other", label: "Other", parent: "other" },
];
const shopKey = (t) => String(t || "").toLowerCase().replace(/[^a-z0-9]/g, "");
function shopCategoryOf(item) {
  const cat = String((item && item.category) || "").toLowerCase();
  if (cat === "apparel") {
    const type = shopKey(normalizeItemType(item.itemType));
    const hit = SHOP_CATEGORIES.find((c) => c.parent === "apparel" && (c.types || []).some((t) => shopKey(t) === type));
    return hit ? hit.id : "other-clothing";
  }
  const hit = SHOP_CATEGORIES.find((c) => c.id === cat && c.parent !== "apparel");
  return hit ? hit.id : "other";
}
const shopCategoryLabel = (id) => { const c = SHOP_CATEGORIES.find((x) => x.id === id); return c ? c.label : ""; };
// the shop categories that have at least one piece, in the order above
function shopCategoriesPresent(items) {
  const seen = new Set((items || []).map(shopCategoryOf));
  return SHOP_CATEGORIES.map((c) => c.id).filter((id) => seen.has(id));
}
// what the product page shows: "Hoodies", or "Shirts · Polo" when the exact type adds something
function categoryLine(item) {
  const id = shopCategoryOf(item);
  const c = SHOP_CATEGORIES.find((x) => x.id === id);
  const type = normalizeItemType(item.itemType);
  const parts = [c ? c.label : ""];
  if (type && String(item.category).toLowerCase() === "apparel" && !(c && (c.types || [])[0] === type)) parts.push(type);
  else if (type && String(item.category).toLowerCase() !== "apparel") parts.push(type);
  return parts.filter(Boolean).join(" · ");
}
// CSV import: "T-Shirts", "Hoodie", "Other Clothing" … in the category column mean Apparel + that type; the old category names still work
function resolveImportCategory(rawCategory, rawType) {
  const key = shopKey(rawCategory);
  const type = normalizeItemType(rawType);
  const plain = CATEGORY_LIST.find((c) => shopKey(c.id) === key || shopKey(c.label) === key);
  if (plain) return { category: plain.id, itemType: type };
  const clothing = SHOP_CATEGORIES.find((c) => c.parent === "apparel" && (shopKey(c.label) === key || (c.types || []).some((t) => shopKey(t) === key || shopKey(t) + "s" === key)));
  if (clothing) return { category: "apparel", itemType: type || (clothing.types || [])[0] || "" };
  return { category: "other", itemType: type };
}
// CATEGORY-LOGIC-END

// SKU-LOGIC-START
// "35J" and "35K" share the number 35, which is usually a mix-up; leading zeros are ignored ("035J" is 35)
const skuNumber = (s) => { const m = String(s || "").trim().match(/^0*(\d+)/); return m ? m[1] : ""; };
const skuKey = (s) => { const t = String(s || "").trim(); return t ? (skuNumber(t) ? "n" + skuNumber(t) : "s" + t.toLowerCase()) : ""; };
function findSkuClashes(items, sku, excludeId) {
  const text = String(sku || "").trim().toLowerCase();
  const num = skuNumber(sku);
  if (!text) return { exact: [], sameNumber: [] };
  const others = (items || []).filter((i) => i.id !== excludeId && String(i.sku || "").trim());
  const exact = others.filter((i) => String(i.sku).trim().toLowerCase() === text);
  const sameNumber = num ? others.filter((i) => !exact.includes(i) && skuNumber(i.sku) === num) : [];
  return { exact, sameNumber };
}
// every item whose SKU number (or whole SKU, when it has no number) is also used by another item
function skuDuplicateIds(items) {
  const groups = new Map();
  for (const i of items || []) {
    const k = skuKey(i.sku);
    if (!k) continue;
    if (!groups.has(k)) groups.set(k, []);
    groups.get(k).push(i.id);
  }
  const out = new Set();
  for (const ids of groups.values()) if (ids.length > 1) ids.forEach((id) => out.add(id));
  return out;
}
// the SKU is a number plus an optional letter typed in its own box: 100 + M = "100M"
const splitSku = (s) => { const m = String(s || "").trim().match(/^(\d*)(.*)$/); return { num: m[1], suffix: m[2] }; };
const joinSku = (num, suffix) => String(num || "").replace(/\D/g, "") + String(suffix || "").replace(/\s+/g, "").toUpperCase();
// the next free number: one more than the highest number in use ("49G" and "102" both count)
function nextSkuNumber(items) {
  let max = 0;
  for (const i of items || []) { const n = Number(skuNumber(i.sku)); if (n > max) max = n; }
  return max + 1;
}
// SKU-LOGIC-END

// SHOP-FILTER-START
const PRICE_BANDS = [
  { id: "", label: "Any price" },
  { id: "u30", label: "Under RM30", min: 0, max: 30 },
  { id: "30-60", label: "RM30 to RM60", min: 30, max: 60 },
  { id: "60-100", label: "RM60 to RM100", min: 60, max: 100 },
  { id: "100", label: "RM100 and up", min: 100, max: Infinity },
];
const priceInBand = (price, bandId) => {
  const b = PRICE_BANDS.find((x) => x.id === bandId);
  if (!b || !b.id) return true;
  const n = Number(price);
  return Number.isFinite(n) && n >= b.min && n < b.max;
};
const SIZE_ORDER = ["XXS", "XS", "S", "M", "L", "XL", "XXL", "XXXL"];
const sizeRank = (s) => { const i = SIZE_ORDER.indexOf(String(s).toUpperCase()); return i === -1 ? 100 : i; };
const sortSizes = (list) => [...list].sort((a, b) => sizeRank(a) - sizeRank(b) || String(a).localeCompare(String(b), undefined, { numeric: true }));
// the filters on top of category/search: sizes (any of), one brand, one price band
function passesShopFilters(item, f, effectivePrice, size) {
  if (f.sizes && f.sizes.length && !f.sizes.some((x) => String(x).toLowerCase() === String(size || "").toLowerCase())) return false;
  if (f.brand && String(item.brand || "").trim().toLowerCase() !== String(f.brand).toLowerCase()) return false;
  if (!priceInBand(effectivePrice, f.price)) return false;
  return true;
}
// SHOP-FILTER-END

// BGREMOVE-START
// Makes a flat background transparent. Works from the picture's edges inward, so white or cream inside the logo itself is kept.
// d = RGBA pixel data, tol = how different from the background a pixel may be and still count as background.
function removeBackground(d, w, h, tol) {
  const at = (x, y) => (y * w + x) * 4;
  const corners = [[0, 0], [w - 1, 0], [0, h - 1], [w - 1, h - 1]].map(([x, y]) => { const i = at(x, y); return [d[i], d[i + 1], d[i + 2], d[i + 3]]; });
  if (corners.filter((c) => c[3] < 30).length >= 3) return { ok: false, removed: 0, reason: "already-transparent" };
  const far = (a, b) => Math.sqrt((a[0] - b[0]) ** 2 + (a[1] - b[1]) ** 2 + (a[2] - b[2]) ** 2);
  const base = corners.find((c) => corners.filter((o) => far(o, c) <= tol * 1.5).length >= 3);
  if (!base) return { ok: false, removed: 0, reason: "no-plain-background" };
  const agree = corners.filter((o) => far(o, base) <= tol * 1.5);
  const bg = [0, 1, 2].map((k) => agree.reduce((t, c) => t + c[k], 0) / agree.length);
  const dist = (i) => Math.sqrt((d[i] - bg[0]) ** 2 + (d[i + 1] - bg[1]) ** 2 + (d[i + 2] - bg[2]) ** 2);
  const gone = new Uint8Array(w * h);
  const stack = [];
  const push = (x, y) => { const k = y * w + x; if (!gone[k] && dist(k * 4) <= tol) { gone[k] = 1; stack.push(k); } };
  for (let x = 0; x < w; x++) { push(x, 0); push(x, h - 1); }
  for (let y = 0; y < h; y++) { push(0, y); push(w - 1, y); }
  while (stack.length) {
    const k = stack.pop(), x = k % w, y = (k - x) / w;
    if (x > 0) push(x - 1, y); if (x < w - 1) push(x + 1, y); if (y > 0) push(x, y - 1); if (y < h - 1) push(x, y + 1);
  }
  let removed = 0;
  for (let k = 0; k < w * h; k++) if (gone[k]) { d[k * 4 + 3] = 0; removed++; }
  // soften the edge so the outline does not look cut out
  for (let y = 0; y < h; y++) for (let x = 0; x < w; x++) {
    const k = y * w + x;
    if (gone[k]) continue;
    const touches = (x > 0 && gone[k - 1]) || (x < w - 1 && gone[k + 1]) || (y > 0 && gone[k - w]) || (y < h - 1 && gone[k + w]);
    if (!touches) continue;
    const dd = dist(k * 4);
    if (dd < tol * 1.25) d[k * 4 + 3] = Math.min(d[k * 4 + 3], Math.round(255 * Math.max(0, Math.min(1, (dd - tol) / (tol * 0.25)))));
  }
  return { ok: removed > w * h * 0.01, removed, share: removed / (w * h) };
}
const LOGO_STRENGTH = { gentle: 14, normal: 28, strong: 48 };
// BGREMOVE-END

// POS-LOGIC-START
// an extra discount on the whole order is shared across the lines in proportion to what each costs after its own discount,
// so every line's final price still adds up to the order total (profit and analytics stay correct)
function allocateOrderDiscount(lines, amount) {
  const total = lines.reduce((sum, l) => sum + l.finalPrice, 0);
  const amt = Math.min(Math.max(0, Number(amount) || 0), total);
  if (!(amt > 0) || !(total > 0)) return lines.map((l) => ({ ...l, orderShare: 0 }));
  let left = Math.round(amt * 100) / 100;
  return lines.map((l, i) => {
    const want = i === lines.length - 1 ? left : Math.round(amt * (l.finalPrice / total) * 100) / 100;
    const share = Math.min(left, want, l.finalPrice);
    left = Math.round((left - share) * 100) / 100;
    return { ...l, orderShare: share, finalPrice: Math.round((l.finalPrice - share) * 100) / 100 };
  });
}
// POS-LOGIC-END

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

// CUSTOMER-LOGIC-START
// Phone numbers are stored as digits with the country code (Malaysia by default): "011-5635 8946" -> "601156358946"
function normalizePhone(input) {
  let d = String(input || "").replace(/\D/g, "");
  if (!d) return "";
  if (d.startsWith("00")) d = d.slice(2);
  else if (d.startsWith("0")) d = "60" + d.slice(1);
  else if (d.startsWith("1") && d.length <= 10) d = "60" + d;
  return /^[0-9]{9,15}$/.test(d) ? d : "";
}
const formatPhone = (d) => (d ? `+${d}` : "");
const DEFAULT_SHIP_MESSAGE = "Hello {name} 👋 Your parcel has been shipped out! Thank you so much for your purchase ✨\n\nTracking Number: {tracking}\n\nIn case you're interested in other pieces: {store}";
const DEFAULT_PROMO_MESSAGE = "Hi {name}! 👋 New pieces just dropped ✨ Have a look: {store}";
// fills {name}, {tracking} and {store}; with no name, "Hello {name} 👋" simply becomes "Hello 👋"
function renderMessage(tpl, vars) {
  const v = vars || {};
  const name = String(v.name || "").trim();
  let out = String(tpl || "");
  out = name ? out.split("{name}").join(name) : out.replace(/ ?\{name\}/g, "");
  out = out.split("{tracking}").join(String(v.tracking || "").trim());
  out = out.split("{store}").join(String(v.store || ""));
  return out;
}
const waChatLink = (phone, text) => `https://wa.me/${phone}${text ? `?text=${encodeURIComponent(text)}` : ""}`;
function customerStats(customers, txns) {
  const out = {};
  for (const c of customers || []) out[c.id] = { orders: 0, spent: 0, last: null };
  for (const t of txns || []) {
    const st = t.customerId && out[t.customerId];
    if (!st) continue;
    st.orders += 1;
    st.spent = Math.round((st.spent + Number(t.total || 0)) * 100) / 100;
    if (!st.last || new Date(t.timestamp) > new Date(st.last)) st.last = t.timestamp;
  }
  return out;
}
// spreadsheet-safe: a cell that starts with = + - @ would be run as a formula, so it gets a leading '
function csvCell(v) {
  let x = String(v == null ? "" : v);
  if (/^[=+\-@\t\r]/.test(x)) x = "'" + x;
  return /[",\n\r]/.test(x) ? `"${x.replace(/"/g, '""')}"` : x;
}
function customersToCsv(customers, stats) {
  const rows = [["Name", "Phone", "Orders", "Spent (RM)", "Last order"]];
  for (const c of customers || []) {
    const st = (stats && stats[c.id]) || { orders: 0, spent: 0, last: null };
    rows.push([c.name, c.phone, st.orders, st.spent, st.last ? String(st.last).slice(0, 10) : ""]);
  }
  return rows.map((r) => r.map(csvCell).join(",")).join("\r\n");
}
// CUSTOMER-LOGIC-END

// SHOPPER-LOGIC-START
const MY_STATES = ["Johor", "Kedah", "Kelantan", "Melaka", "Negeri Sembilan", "Pahang", "Perak", "Perlis", "Pulau Pinang", "Sabah", "Sarawak", "Selangor", "Terengganu", "Kuala Lumpur", "Labuan", "Putrajaya"];
const EMAIL_RE = /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/;
function passwordProblem(pw) {
  const p = String(pw || "");
  if (p.length < 10) return "Use at least 10 characters.";
  if (!/[A-Za-z]/.test(p) || !/[0-9]/.test(p)) return "Include at least one letter and one number.";
  return "";
}
function validateAddress(f) {
  const e = {};
  const line1 = String(f.address1 || "").trim(), country = String(f.country || "").trim(), pc = String(f.postcode || "").trim();
  if (!line1) e.address1 = "Enter your street address."; else if (line1.length > 120) e.address1 = "That address line is too long.";
  if (String(f.address2 || "").trim().length > 120) e.address2 = "That address line is too long.";
  if (!String(f.city || "").trim()) e.city = "Enter your city.";
  if (!String(f.state || "").trim()) e.state = "Select your state.";
  if (!country) e.country = "Enter your country.";
  const isMY = country.toLowerCase() === "malaysia";
  if (!pc) e.postcode = "Enter your postcode.";
  else if (isMY ? !/^[0-9]{5}$/.test(pc) : !/^[A-Za-z0-9 -]{3,10}$/.test(pc)) e.postcode = isMY ? "Malaysian postcodes have 5 digits." : "Enter a valid postcode.";
  return e;
}
function validateSignup(f) {
  const e = {};
  const name = String(f.fullName || "").trim();
  if (name.length < 2) e.fullName = "Enter your full name."; else if (name.length > 80) e.fullName = "That name is too long.";
  if (!normalizePhone(f.phone)) e.phone = "Enter a valid phone number, e.g. 011-2345 6789.";
  if (!EMAIL_RE.test(String(f.email || "").trim())) e.email = "Enter a valid email address.";
  const pp = passwordProblem(f.password);
  if (pp) e.password = pp;
  if (!f.confirmPassword) e.confirmPassword = "Confirm your password.";
  else if (!pp && f.password !== f.confirmPassword) e.confirmPassword = "The passwords don't match.";
  Object.assign(e, validateAddress(f));
  if (!f.agree) e.agree = "Please agree to continue.";
  return e;
}
// what the sign-up form sends to Supabase Auth; a database trigger turns it into the shopper's profile
function signupPayload(f) {
  return {
    full_name: String(f.fullName || "").trim(), phone: normalizePhone(f.phone),
    address_line1: String(f.address1 || "").trim(), address_line2: String(f.address2 || "").trim(),
    city: String(f.city || "").trim(), state: String(f.state || "").trim(), postcode: String(f.postcode || "").trim(),
    country: String(f.country || "").trim(), marketing_opt_in: !!f.marketing,
  };
}
function friendlySignupError(msg) {
  const m = String(msg || "").toLowerCase();
  if (m.includes("already registered") || m.includes("already been registered")) return "We couldn't create that account. If you already have one, log in or reset your password.";
  if (m.includes("database error") || m.includes("not allowed") || m.includes("closed") || m.includes("disabled")) return "Sign-ups aren't open right now. Please try again later.";
  if (m.includes("rate limit") || m.includes("too many") || m.includes("seconds")) return "Too many attempts. Please wait a few minutes and try again.";
  if (m.includes("password")) return "That password isn't accepted. Try a longer one with letters and numbers.";
  return "We couldn't create your account. Please check your details and try again.";
}
function friendlyLoginError(msg) {
  const m = String(msg || "").toLowerCase();
  if (m.includes("invalid login") || m.includes("invalid credentials")) return "Email or password is incorrect.";
  if (m.includes("not confirmed")) return "Please confirm your email first (check your inbox), then log in.";
  if (m.includes("rate limit") || m.includes("too many")) return "Too many attempts. Please wait a few minutes and try again.";
  return "We couldn't log you in. Please try again.";
}
// Supabase email links come back as #access_token=…&type=recovery|signup, or #error=…
function parseAuthHash(hash) {
  const h = String(hash || "").replace(/^#/, "");
  if (!h.includes("=")) return {};
  const p = new URLSearchParams(h);
  return { type: p.get("type") || "", accessToken: p.get("access_token") || "", refreshToken: p.get("refresh_token") || "", error: p.get("error") || p.get("error_description") ? "bad-link" : "" };
}
// a piece another customer is holding for an unpaid order
const isReserved = (item, now) => !!(item && item.reservedUntil && new Date(item.reservedUntil).getTime() > (now == null ? Date.now() : now));
function orderLabel(o) {
  if (!o) return "";
  if (o.status === "cancelled") return o.payment_status === "expired" ? "Expired" : "Cancelled";
  if (o.payment_status === "pending") return "Awaiting payment";
  if (o.payment_status === "claimed") return "Payment being checked";
  if (o.payment_status === "refunded") return "Refunded";
  return { placed: "Paid", processing: "Processing", shipped: "Shipped", delivered: "Delivered", completed: "Completed" }[o.status] || "Order placed";
}
function timeLeft(expiresAt, now) {
  const ms = new Date(expiresAt).getTime() - (now == null ? Date.now() : now);
  if (!(ms > 0)) return "Expired";
  const mins = Math.floor(ms / 60000), h = Math.floor(mins / 60), m = mins % 60;
  return h > 0 ? `${h}h ${m}m` : `${m}m`;
}
function validateCheckout(f) {
  const e = validateAddress(f);
  const name = String(f.fullName || "").trim();
  if (name.length < 2) e.fullName = "Enter your full name."; else if (name.length > 80) e.fullName = "That name is too long.";
  if (!normalizePhone(f.phone)) e.phone = "Enter a valid phone number, e.g. 011-2345 6789.";
  if (String(f.country || "").trim() && String(f.country).trim().toLowerCase() !== "malaysia") e.country = "We only ship within Malaysia for now.";
  if (String(f.note || "").length > 300) e.note = "Your note is too long.";
  return e;
}
// SHOPPER-LOGIC-END

// ORDERS-LOGIC-START
// which list an order belongs in, on the seller's Orders tab
function orderBucket(o, now) {
  const t = now == null ? Date.now() : now;
  if (!o || o.status === "cancelled" || o.status === "completed") return "done";
  if (o.payment_status === "claimed") return "todo";
  if (o.payment_status === "pending") return new Date(o.expires_at).getTime() > t ? "waiting" : "done";
  if (o.payment_status === "paid" && (o.status === "processing" || o.status === "placed")) return "todo";
  if (o.status === "shipped" || o.status === "delivered") return "shipped";
  return "done";
}
// one plain sentence on what the shipping did to your profit
function shippingNote(o, cost) {
  const fee = Number(o.shipping_fee || 0);
  if (cost === "" || cost == null || isNaN(Number(cost))) {
    return o.free_shipping || fee === 0 ? "Free shipping order: enter what the courier charged so your profit stays right." : `Customer paid ${fmtRM(fee)}. Leave the cost blank if the courier charged the same.`;
  }
  const c = Number(cost), gap = Math.round((fee - c) * 100) / 100;
  if (gap === 0) return `Customer paid ${fmtRM(fee)}, courier ${fmtRM(c)}: no difference.`;
  return gap > 0 ? `Customer paid ${fmtRM(fee)}, courier ${fmtRM(c)}: +${fmtRM(gap)} for you.` : `Customer paid ${fmtRM(fee)}, courier ${fmtRM(c)}: ${fmtRM(-gap)} out of your pocket.`;
}
function validateShipping(f) {
  const e = {}, c = String(f.courier || "").trim(), t = String(f.tracking || "").trim(), cost = String(f.cost == null ? "" : f.cost).trim();
  if (c.length < 2) e.courier = "Enter the courier, e.g. J&T.";
  if (t.length < 4) e.tracking = "Enter the tracking number.";
  if (cost !== "" && (isNaN(Number(cost)) || Number(cost) < 0 || Number(cost) > 1000)) e.cost = "Enter the courier cost in RM, or leave it blank.";
  return e;
}
function validateSettings(f) {
  const e = {}, num = (v) => (String(v).trim() === "" ? NaN : Number(v));
  if (!(num(f.west) >= 0 && num(f.west) <= 100)) e.west = "Enter a rate between RM0 and RM100.";
  if (!(num(f.east) >= 0 && num(f.east) <= 100)) e.east = "Enter a rate between RM0 and RM100.";
  if (String(f.freeMin).trim() !== "" && !(num(f.freeMin) >= 0 && num(f.freeMin) <= 10000)) e.freeMin = "Enter an amount, or leave it blank for no free shipping.";
  if (!(Number.isInteger(num(f.hold)) && num(f.hold) >= 1 && num(f.hold) <= 168)) e.hold = "Enter whole hours, between 1 and 168.";
  if (String(f.payment || "").length > 1500) e.payment = "That's too long (1500 characters max).";
  return e;
}
// ORDERS-LOGIC-END

// columns the public storefront is allowed to read — deliberately excludes cost and SKU
const STOREFRONT_ITEM_COLUMNS = "id,user_id,title,category,item_type,brand,condition,colour,colour_condition,remark,price,status,attributes,custom_attributes,photos,show_on_storefront,featured,sale,reserved_until,created_at";

// ---------- storefront v5: routes, banner slideshow, and the owner's editors for banners / FAQ / collections ----------
function parseStoreHash() {
  const h = typeof window !== "undefined" ? window.location.hash : "";
  let m;
  if ((m = h.match(/^#item=([A-Za-z0-9-]+)/))) return { page: "home", item: m[1], collection: null };
  if ((m = h.match(/^#collection=([A-Za-z0-9-]+)/))) return { page: "collection", item: null, collection: m[1] };
  if (h === "#sale") return { page: "sale", item: null, collection: null };
  if (h === "#faq") return { page: "faq", item: null, collection: null };
  if (h === "#collections") return { page: "collections", item: null, collection: null };
  if (["#login", "#signup", "#forgot", "#reset", "#account", "#bag", "#checkout"].includes(h)) return { page: h.slice(1), item: null, collection: null };
  if ((m = h.match(/^#order=([A-Za-z0-9-]+)/))) return { page: "order", item: null, collection: null, order: m[1] };
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
    <div className="w-full relative overflow-hidden"
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
          const img = <img src={s.url} alt="" className="w-full block object-cover aspect-[3/2] md:aspect-[21/9]" style={{ objectPosition: "center" }} draggable={false} />;
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
          <div data-slide-counter="" className="hidden md:flex absolute bottom-5 left-5 items-center gap-3 text-[12px] font-semibold tracking-[0.1em]" style={{ color: "#FFFFFF" }}>
            <span>{String(idx + 1).padStart(2, "0")}</span><span style={{ opacity: 0.6 }}>/ {String(n).padStart(2, "0")}</span>
          </div>
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

const COURIERS = ["J&T", "Pos Laju", "Ninja Van", "DHL eCommerce", "Flash Express", "City-Link", "SPX"];
const OD = { card: "#17171A", deep: "#0B0B0C", line: "#2A2A2F", ink: "#F2F0EA", mute: "#8C8A85", dim: "#75726C", gold: "#C9A567", red: "#C9695E" };

// first tap arms it, second tap does it: for anything you wouldn't want to hit by accident
function ConfirmButton({ label, confirmLabel, onConfirm, disabled, style, className, ...rest }) {
  const [armed, setArmed] = useState(false);
  useEffect(() => { if (!armed) return; const t = setTimeout(() => setArmed(false), 5000); return () => clearTimeout(t); }, [armed]);
  return (
    <button type="button" disabled={disabled} className={className} style={style} {...rest} onClick={() => { if (armed) { setArmed(false); onConfirm(); } else setArmed(true); }}>
      {armed ? confirmLabel : label}
    </button>
  );
}

function OrderDetail({ o, priv, shipTemplate, storeLink, onAct, onSaveNote }) {
  const [busy, setBusy] = useState(false);
  const [msg, setMsg] = useState(null);
  const [ship, setShip] = useState({ courier: "", tracking: "", cost: "" });
  const [errors, setErrors] = useState({});
  const [reason, setReason] = useState("");
  const [note, setNote] = useState((priv && priv.admin_note) || "");
  useEffect(() => {
    setShip({ courier: o.courier || "", tracking: o.tracking_no || "", cost: priv && priv.actual_shipping_cost != null ? String(priv.actual_shipping_cost) : (o.free_shipping || Number(o.shipping_fee) === 0 ? "" : String(o.shipping_fee)) });
    setErrors({});
  }, [o.id, o.status, o.payment_status, priv && priv.actual_shipping_cost]);
  const run = async (name, args, okText) => {
    setBusy(true); setMsg(null);
    const r = await onAct(name, args);
    setBusy(false);
    setMsg(r.ok ? { text: okText } : { text: r.error || "Something went wrong.", bad: true });
  };
  const lapsed = o.payment_status === "pending" && new Date(o.expires_at).getTime() <= Date.now();
  const expiredLate = (o.status === "cancelled" && o.payment_status === "expired") || lapsed;
  const unpaid = (o.payment_status === "pending" && !lapsed) || o.payment_status === "claimed";
  const paid = o.payment_status === "paid";
  const phone = normalizePhone(o.customer_phone) || o.customer_phone;
  const field = { borderColor: OD.line, background: OD.deep, color: OD.ink };
  const label = (t) => <p className="text-[10px] uppercase tracking-wide mb-1.5" style={{ color: OD.dim }}>{t}</p>;
  const events = [...(o.order_events || [])].sort((x, y) => new Date(x.created_at) - new Date(y.created_at));
  const trackingText = o.tracking_no ? renderMessage(shipTemplate, { name: o.customer_name, tracking: o.tracking_no, store: storeLink }) : "";
  return (
    <div data-order-detail={o.id} className="pt-4 mt-3 border-t space-y-5" style={{ borderColor: OD.line }}>
      <div>
        {label("Customer")}
        <p className="text-sm" style={{ color: OD.ink }}>{o.customer_name}</p>
        <p className="text-xs mt-0.5" style={{ color: OD.mute }}>{formatPhone(phone)} · {o.customer_email}</p>
        <p className="text-xs mt-2 leading-relaxed" style={{ color: OD.mute }}>{[o.ship_line1, o.ship_line2, `${o.ship_postcode} ${o.ship_city}`, o.ship_state, o.ship_country].filter(Boolean).join(", ")}</p>
        {o.customer_note && <p data-order-note-from-customer="" className="text-xs mt-2 italic" style={{ color: OD.gold }}>"{o.customer_note}"</p>}
        <a data-whatsapp-customer="" href={waChatLink(phone, "")} target="_blank" rel="noopener noreferrer" className="inline-block mt-2 text-xs underline" style={{ color: OD.gold }}>WhatsApp {o.customer_name.split(" ")[0]}</a>
      </div>

      <div>
        {label("Pieces")}
        {(o.order_items || []).map((li) => (
          <div key={li.id} className="flex justify-between gap-3 py-1.5 text-sm" style={{ color: OD.ink }}>
            <span className="min-w-0 truncate">{li.title}{li.size ? ` · ${li.size}` : ""}</span><span className="shrink-0" style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{rm(li.price)}</span>
          </div>
        ))}
        <div className="flex justify-between text-xs pt-2 mt-1 border-t" style={{ borderColor: OD.line, color: OD.mute }}><span>Shipping ({o.shipping_zone === "east" ? "East" : "West"} Malaysia{o.free_shipping ? ", free" : ""})</span><span>{Number(o.shipping_fee) === 0 ? "Free" : rm(o.shipping_fee)}</span></div>
        <div className="flex justify-between text-sm pt-1.5" style={{ color: OD.ink }}><span>Customer pays</span><strong data-order-total="">{rm(o.total)}</strong></div>
      </div>

      {(unpaid || expiredLate) && (
        <div data-payment-block="" className="rounded-lg border p-4" style={{ borderColor: OD.gold, background: "#1C1810" }}>
          {label("Payment")}
          {o.payment_status === "claimed" && <p className="text-sm mb-3" style={{ color: OD.ink }}>The customer says they've paid. Check your bank for <strong>{rm(o.total)}</strong> with reference <strong>{o.order_number}</strong> before you confirm.</p>}
          {o.payment_status === "pending" && !lapsed && <p className="text-sm mb-3" style={{ color: OD.ink }}>Waiting for <strong>{rm(o.total)}</strong>, reference <strong>{o.order_number}</strong>. The hold runs out in {timeLeft(o.expires_at)}. If the money has already arrived, you can confirm it now.</p>}
          {expiredLate && <p className="text-sm mb-3" style={{ color: OD.mute }}>This order expired. If the customer paid late, you can reopen it. It works if the pieces are still free.</p>}
          <ConfirmButton data-action="mark-paid" disabled={busy} label={expiredLate ? "Reopen and mark as paid" : "Mark as paid"} confirmLabel="Tap again: I've checked my bank" onConfirm={() => run("admin_mark_paid", { p_order: o.id }, "Marked as paid. The sale is now in Sales and the customer is saved.")}
            className="w-full text-sm font-semibold py-2.5 rounded-lg" style={{ background: OD.gold, color: OD.deep, opacity: busy ? 0.5 : 1 }} />
        </div>
      )}
      {paid && o.paid_at && <p data-paid-note="" className="text-xs" style={{ color: OD.gold }}>Paid on {fmtWhen(o.paid_at)}. The sale is in your Sales tab.</p>}

      {paid && ["processing", "shipped", "delivered"].includes(o.status) && (
        <div data-ship-block="">
          {label(o.status === "processing" ? "Ship it" : "Shipping details")}
          <div className="grid grid-cols-2 gap-2">
            <div>
              <input list="rv-couriers" name="courier" value={ship.courier} onChange={(e) => setShip({ ...ship, courier: e.target.value })} placeholder="Courier" aria-label="Courier" className="w-full px-3 py-2 rounded-lg border text-xs outline-none" style={field} />
              <datalist id="rv-couriers">{COURIERS.map((c) => <option key={c} value={c} />)}</datalist>
              {errors.courier && <p role="alert" className="text-[11px] mt-1" style={{ color: OD.red }}>{errors.courier}</p>}
            </div>
            <div>
              <input name="tracking" value={ship.tracking} maxLength={60} onChange={(e) => setShip({ ...ship, tracking: e.target.value })} placeholder="Tracking number" aria-label="Tracking number" className="w-full px-3 py-2 rounded-lg border text-xs outline-none" style={{ ...field, fontFamily: "'IBM Plex Mono', monospace" }} />
              {errors.tracking && <p role="alert" className="text-[11px] mt-1" style={{ color: OD.red }}>{errors.tracking}</p>}
            </div>
          </div>
          <div className="flex items-center gap-2 mt-2">
            <span className="text-xs" style={{ color: OD.mute }}>Courier charged you RM</span>
            <input name="cost" type="number" inputMode="decimal" min="0" value={ship.cost} onChange={(e) => setShip({ ...ship, cost: e.target.value })} placeholder="optional" aria-label="Courier cost" className="w-24 px-3 py-2 rounded-lg border text-xs outline-none" style={field} />
          </div>
          {errors.cost && <p role="alert" className="text-[11px] mt-1" style={{ color: OD.red }}>{errors.cost}</p>}
          <p data-shipping-note="" className="text-[11px] mt-1.5" style={{ color: OD.dim }}>{shippingNote(o, ship.cost)}</p>
          <button data-action="ship" disabled={busy} onClick={() => {
            const er = validateShipping(ship); setErrors(er);
            if (Object.keys(er).length) return;
            run("admin_ship_order", { p_order: o.id, p_courier: ship.courier.trim(), p_tracking: ship.tracking.trim(), p_actual_cost: String(ship.cost).trim() === "" ? null : Number(ship.cost) }, o.status === "processing" ? "Marked as shipped. The customer can see the tracking number." : "Shipping details saved.");
          }} className="mt-3 w-full text-sm font-semibold py-2.5 rounded-lg" style={{ background: OD.gold, color: OD.deep, opacity: busy ? 0.5 : 1 }}>{o.status === "processing" ? "Mark as shipped" : "Save shipping details"}</button>
          {o.tracking_no && (
            <a data-whatsapp-tracking="" href={waChatLink(phone, trackingText)} target="_blank" rel="noopener noreferrer" className="block text-center mt-2 text-sm py-2.5 rounded-lg border" style={{ borderColor: OD.line, color: OD.gold, textDecoration: "none" }}>Send tracking on WhatsApp</a>
          )}
        </div>
      )}

      {paid && (o.status === "shipped" || o.status === "delivered") && (
        <div className="flex gap-2">
          {o.status === "shipped" && <button data-action="delivered" disabled={busy} onClick={() => run("admin_advance_order", { p_order: o.id, p_to: "delivered" }, "Marked as delivered.")} className="flex-1 text-sm py-2.5 rounded-lg border" style={{ borderColor: OD.line, color: OD.ink }}>Mark delivered</button>}
          <button data-action="completed" disabled={busy} onClick={() => run("admin_advance_order", { p_order: o.id, p_to: "completed" }, "Order completed.")} className="flex-1 text-sm py-2.5 rounded-lg border" style={{ borderColor: OD.line, color: OD.ink }}>Mark completed</button>
        </div>
      )}

      {["placed", "processing"].includes(o.status) && (
        <div>
          {label(paid ? "Cancel and refund" : "Cancel this order")}
          <p className="text-[11px] mb-2" style={{ color: OD.dim }}>{paid ? "The pieces go back on sale and the sale is removed from Sales. Refund the customer yourself." : "The pieces are released straight away."}</p>
          <input name="reason" value={reason} maxLength={200} onChange={(e) => setReason(e.target.value)} placeholder="Reason (optional, shown to you only)" aria-label="Reason for cancelling" className="w-full px-3 py-2 rounded-lg border text-xs outline-none mb-2" style={field} />
          <ConfirmButton data-action="cancel" disabled={busy} label={paid ? "Cancel and refund" : "Cancel order"} confirmLabel="Tap again to confirm" onConfirm={() => run("admin_cancel_order", { p_order: o.id, p_reason: reason.trim() || null }, paid ? "Cancelled. The pieces are back on sale and the sale is removed." : "Order cancelled and the pieces released.")}
            className="w-full text-sm py-2.5 rounded-lg border" style={{ borderColor: OD.line, color: OD.red }} />
        </div>
      )}

      {msg && <p data-order-msg="" role="status" className="text-sm" style={{ color: msg.bad ? OD.red : OD.gold }}>{msg.bad ? "" : "✓ "}{msg.text}</p>}

      <div>
        {label("Your private note")}
        <textarea name="note" rows={2} maxLength={500} value={note} onChange={(e) => setNote(e.target.value)} aria-label="Private note" className="w-full px-3 py-2 rounded-lg border text-xs outline-none" style={field} />
        <button data-action="save-note" onClick={() => onSaveNote(o.id, note)} className="mt-1.5 text-xs px-3 py-1.5 rounded-lg border" style={{ borderColor: OD.line, color: OD.gold }}>Save note</button>
      </div>

      <div>
        {label("History")}
        <ol data-order-history="">
          {events.map((ev) => (
            <li key={ev.id} className="flex justify-between gap-3 py-1.5 text-xs border-b" style={{ borderColor: OD.line, color: OD.mute }}><span>{ev.label}</span><span className="shrink-0">{fmtWhen(ev.created_at)}</span></li>
          ))}
        </ol>
      </div>
    </div>
  );
}

function CheckoutSettings({ settings, checkoutOpen, onSave, onToggle }) {
  const init = (st) => ({ payment: (st && st.payment_instructions) || "", west: st ? String(st.ship_west) : "7.5", east: st ? String(st.ship_east) : "15", freeMin: st ? (st.free_ship_min != null ? String(st.free_ship_min) : "") : "30", hold: st ? String(st.hold_hours) : "24" });
  const hasPayment = !!settings && String(settings.payment_instructions || "").trim() !== "";
  const [f, setF] = useState(init(settings));
  const [errors, setErrors] = useState({});
  const [msg, setMsg] = useState(null);
  const [open, setOpen] = useState(!hasPayment);
  const [busy, setBusy] = useState(false);
  useEffect(() => { setF(init(settings)); }, [settings && settings.updated_at, !!settings]);
  const field = { borderColor: OD.line, background: OD.deep, color: OD.ink };
  const set = (k) => (e) => setF((p) => ({ ...p, [k]: e.target.value }));
  async function save() {
    const er = validateSettings(f); setErrors(er);
    if (Object.keys(er).length) return;
    setBusy(true); const r = await onSave(f); setBusy(false);
    setMsg(r.ok ? { text: "Saved." } : { text: r.error || "Couldn't save.", bad: true });
  }
  async function toggle(next) {
    setBusy(true); const r = await onToggle(next); setBusy(false);
    setMsg(r.ok ? { text: next ? "Checkout is open. Customers can now order." : "Checkout is closed." } : { text: r.error || "Couldn't change that.", bad: true });
  }
  const num = (k, label, ph) => (
    <label className="block"><span className="block text-[10px] uppercase tracking-wide mb-1" style={{ color: OD.dim }}>{label}</span>
      <input name={k} type="number" inputMode="decimal" min="0" value={f[k]} onChange={set(k)} placeholder={ph} className="w-full px-3 py-2 rounded-lg border text-xs outline-none" style={field} />
      {errors[k] && <span role="alert" className="block text-[11px] mt-1" style={{ color: OD.red }}>{errors[k]}</span>}</label>
  );
  return (
    <div data-checkout-settings="" className="rounded-xl border p-5" style={{ background: OD.card, borderColor: OD.line }}>
      <div className="flex items-center justify-between gap-3">
        <div className="min-w-0">
          <h2 className="text-sm font-semibold" style={{ color: OD.ink, fontFamily: "'Fraunces', serif" }}>Checkout</h2>
          <p data-checkout-state="" className="text-xs mt-0.5" style={{ color: checkoutOpen ? OD.gold : OD.mute }}>{checkoutOpen ? "Open: customers can order" : "Closed: customers can only browse and WhatsApp you"}</p>
        </div>
        <button data-settings-toggle="" onClick={() => setOpen(!open)} className="text-xs px-3 py-1.5 rounded-lg border shrink-0" style={{ borderColor: OD.line, color: OD.gold }}>{open ? "Hide settings" : "Settings"}</button>
      </div>
      {open && (
        <div className="mt-4 space-y-4">
          <label className="block"><span className="block text-[10px] uppercase tracking-wide mb-1" style={{ color: OD.dim }}>How customers pay (shown after they place an order)</span>
            <textarea name="payment" rows={4} maxLength={1500} value={f.payment} onChange={set("payment")} placeholder={"DuitNow ID: 01x-xxxxxxx\nBank: Maybank 1234 5678 9012\nName: Your Name"} className="w-full px-3 py-2 rounded-lg border text-xs outline-none" style={field} />
            {errors.payment && <span role="alert" className="block text-[11px] mt-1" style={{ color: OD.red }}>{errors.payment}</span>}</label>
          <div className="grid grid-cols-2 gap-3">
            {num("west", "West Malaysia shipping (RM)", "7.5")}{num("east", "East Malaysia shipping (RM)", "15")}
            {num("freeMin", "Free shipping from (RM)", "blank = never")}{num("hold", "Hold unpaid pieces (hours)", "24")}
          </div>
          <button data-action="save-settings" disabled={busy} onClick={save} className="text-xs font-semibold px-4 py-2 rounded-lg" style={{ background: OD.gold, color: OD.deep, opacity: busy ? 0.5 : 1 }}>Save settings</button>
        </div>
      )}
      <div className="mt-4 pt-4 border-t" style={{ borderColor: OD.line }}>
        {checkoutOpen ? (
          <button data-action="close-checkout" disabled={busy} onClick={() => toggle(false)} className="text-xs px-4 py-2 rounded-lg border" style={{ borderColor: OD.line, color: OD.red }}>Close checkout</button>
        ) : (
          <ConfirmButton data-action="open-checkout" disabled={busy || !hasPayment} label="Open checkout to customers" confirmLabel="Tap again to open it to customers" onConfirm={() => toggle(true)} className="text-xs font-semibold px-4 py-2 rounded-lg" style={{ background: hasPayment ? OD.gold : "#2A2A2F", color: hasPayment ? OD.deep : OD.dim }} />
        )}
        {!checkoutOpen && !hasPayment && <p data-open-blocked="" className="text-[11px] mt-2" style={{ color: OD.dim }}>Save your payment details first. Customers need to know where to pay.</p>}
        {msg && <p data-settings-msg="" role="status" className="text-xs mt-2" style={{ color: msg.bad ? OD.red : OD.gold }}>{msg.bad ? "" : "✓ "}{msg.text}</p>}
      </div>
    </div>
  );
}

function OrdersTab({ orders, privateById, settings, checkoutOpen, shipTemplate, storeLink, onRefresh, onAct, onSaveNote, onSaveSettings, onToggleCheckout }) {
  const [filter, setFilter] = useState("todo");
  const [openId, setOpenId] = useState(null);
  const list = orders || [];
  const counts = { todo: 0, waiting: 0, shipped: 0, done: 0, all: list.length };
  list.forEach((o) => { counts[orderBucket(o)] += 1; });
  const base = filter === "all" ? list : list.filter((o) => orderBucket(o) === filter);
  // the order you have open stays on screen even after an action moves it to another list, so you can read the confirmation and use the next button
  const pinned = openId ? list.find((o) => o.id === openId) : null;
  const shown = pinned && !base.some((o) => o.id === openId) ? [pinned, ...base] : base;
  const chips = [["todo", "To do"], ["waiting", "Awaiting payment"], ["shipped", "Shipped"], ["done", "Done"], ["all", "All"]];
  const empty = { todo: "Nothing needs you right now.", waiting: "No orders are waiting for payment.", shipped: "Nothing is on its way.", done: "No finished orders yet.", all: "No orders yet." }[filter];
  return (
    <div className="max-w-2xl mx-auto space-y-6">
      <CheckoutSettings settings={settings} checkoutOpen={checkoutOpen} onSave={onSaveSettings} onToggle={onToggleCheckout} />
      <div className="rounded-xl border p-5" style={{ background: OD.card, borderColor: OD.line }}>
        <div className="flex items-center justify-between gap-3 mb-3">
          <h2 className="text-sm font-semibold" style={{ color: OD.ink, fontFamily: "'Fraunces', serif" }}>Orders</h2>
          <button data-action="refresh-orders" onClick={onRefresh} className="text-xs px-3 py-1.5 rounded-lg border" style={{ borderColor: OD.line, color: OD.gold }}>Refresh</button>
        </div>
        <div className="flex gap-2 overflow-x-auto pb-1 mb-4">
          {chips.map(([k, l]) => (
            <button key={k} data-filter={k} onClick={() => setFilter(k)} className="shrink-0 text-xs px-3 py-1.5 rounded-lg border whitespace-nowrap"
              style={{ borderColor: filter === k ? OD.gold : OD.line, color: filter === k ? OD.gold : OD.mute, background: filter === k ? "#241F13" : "transparent" }}>{l} ({counts[k]})</button>
          ))}
        </div>
        {orders === null && <p className="text-xs py-6 text-center" style={{ color: OD.dim }}>Loading…</p>}
        {orders !== null && shown.length === 0 && <p data-orders-empty="" className="text-xs py-6 text-center" style={{ color: OD.dim }}>{empty}</p>}
        <div className="space-y-2">
          {shown.map((o) => {
            const bucket = orderBucket(o), isOpen = openId === o.id, lapsed = o.payment_status === "pending" && new Date(o.expires_at).getTime() <= Date.now();
            return (
              <div key={o.id} data-order-card={o.id} className="rounded-xl border p-3" style={{ borderColor: bucket === "todo" ? OD.gold : OD.line, background: OD.deep }}>
                <button data-order-toggle="" onClick={() => setOpenId(isOpen ? null : o.id)} className="w-full text-left">
                  <div className="flex justify-between gap-3 text-sm" style={{ color: OD.ink }}>
                    <span style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{o.order_number} · {o.customer_name}</span><span>{rm(o.total)}</span>
                  </div>
                  <div className="flex justify-between gap-3 text-xs mt-1" style={{ color: OD.mute }}>
                    <span className="truncate">{(o.order_items || []).map((i) => i.title).join(", ")}</span>
                    <span data-order-label="" className="shrink-0" style={{ color: bucket === "todo" ? OD.gold : OD.mute }}>{lapsed ? "Expired" : orderLabel(o)}</span>
                  </div>
                  <p className="text-[11px] mt-1" style={{ color: OD.dim }}>{fmtWhen(o.created_at)}{o.tracking_no ? ` · ${o.courier || ""} ${o.tracking_no}` : ""}</p>
                </button>
                {isOpen && <OrderDetail o={o} priv={privateById[o.id]} shipTemplate={shipTemplate} storeLink={storeLink} onAct={onAct} onSaveNote={onSaveNote} />}
              </div>
            );
          })}
        </div>
      </div>
    </div>
  );
}

function CustomersTab({ customers, txns, messages, storeLink, onRename, onDelete, onSaveMessages, onAdd }) {
  const [search, setSearch] = useState("");
  const [adding, setAdding] = useState(false);
  const [addPhone, setAddPhone] = useState("");
  const [addName, setAddName] = useState("");
  const [addMsg, setAddMsg] = useState(null);
  const [editId, setEditId] = useState(null);
  const [editName, setEditName] = useState("");
  const [confirmId, setConfirmId] = useState(null);
  const [shipDraft, setShipDraft] = useState(messages.ship || DEFAULT_SHIP_MESSAGE);
  const [promoDraft, setPromoDraft] = useState(messages.promo || DEFAULT_PROMO_MESSAGE);
  const [savedAt, setSavedAt] = useState(0);
  useEffect(() => { setShipDraft(messages.ship || DEFAULT_SHIP_MESSAGE); setPromoDraft(messages.promo || DEFAULT_PROMO_MESSAGE); }, [messages.ship, messages.promo]);
  const stats = useMemo(() => customerStats(customers, txns), [customers, txns]);
  const q = search.trim().toLowerCase();
  const qd = search.replace(/\D/g, "");
  const shown = customers.filter((c) => !q || c.name.toLowerCase().includes(q) || (qd.length >= 3 && c.phone.includes(qd.startsWith("0") ? "60" + qd.slice(1) : qd)));
  const promoTemplate = messages.promo || DEFAULT_PROMO_MESSAGE;
  const card = { background: "#17171A", borderColor: "#2A2A2F" };
  const input = { borderColor: "#2A2A2F", background: "#0B0B0C", color: "#F2F0EA" };
  function exportCsv() {
    const blob = new Blob(["\uFEFF", customersToCsv(customers, stats)], { type: "text/csv;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url; a.download = "ravuno-customers.csv";
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 2000);
  }
  async function addOne() {
    const phone = normalizePhone(addPhone);
    if (!phone) { setAddMsg({ bad: true, text: "That number looks incomplete." }); return; }
    const r = await onAdd(phone, addName);
    if (!r.ok) { setAddMsg({ bad: true, text: "Couldn't save, try again." }); return; }
    setAddMsg({ text: r.existed ? "Already in your list, details updated." : `Saved${addName.trim() ? " " + addName.trim() : ""}.` });
    setAddPhone(""); setAddName("");
  }
  async function save() {
    const ok = await onSaveMessages(shipDraft.trim(), promoDraft.trim());
    if (ok) setSavedAt(Date.now());
  }
  return (
    <div className="max-w-2xl mx-auto space-y-6">
      <div className="rounded-xl border p-5" style={card} data-customers-messages="">
        <h2 className="text-sm font-semibold mb-1" style={{ color: "#F2F0EA", fontFamily: "'Fraunces', serif" }}>Messages you send</h2>
        <p className="text-xs mb-4" style={{ color: "#8C8A85" }}>Use {"{name}"} for the buyer's name, {"{tracking}"} for the tracking number and {"{store}"} for your storefront link.</p>
        <label className="block mb-3">
          <span className="block text-[10px] mb-1 uppercase tracking-wide" style={{ color: "#75726C" }}>Shipping message</span>
          <textarea value={shipDraft} onChange={(e) => setShipDraft(e.target.value)} maxLength={1000} rows={6} aria-label="Shipping message"
            className="w-full px-3 py-2 rounded-lg border text-xs outline-none" style={input} />
        </label>
        <label className="block mb-3">
          <span className="block text-[10px] mb-1 uppercase tracking-wide" style={{ color: "#75726C" }}>Promo message</span>
          <textarea value={promoDraft} onChange={(e) => setPromoDraft(e.target.value)} maxLength={1000} rows={3} aria-label="Promo message"
            className="w-full px-3 py-2 rounded-lg border text-xs outline-none" style={input} />
        </label>
        <div className="flex items-center gap-2 flex-wrap">
          <button onClick={save} className="text-xs font-semibold px-3 py-2 rounded-lg" style={{ background: "#C9A567", color: "#0B0B0C" }}>Save messages</button>
          <button onClick={() => { setShipDraft(DEFAULT_SHIP_MESSAGE); setPromoDraft(DEFAULT_PROMO_MESSAGE); }} className="text-xs font-medium px-3 py-2 rounded-lg border" style={{ borderColor: "#2A2A2F", color: "#8C8A85" }}>Reset to default</button>
          {savedAt > 0 && <span className="text-[11px]" style={{ color: "#C9A567" }}>Saved</span>}
        </div>
      </div>

      <div className="rounded-xl border p-5" style={card}>
        <div className="flex items-center justify-between gap-3 mb-3">
          <h2 className="text-sm font-semibold" style={{ color: "#F2F0EA", fontFamily: "'Fraunces', serif" }}>Customers <span className="text-xs font-normal" style={{ color: "#75726C" }}>{customers.length}</span></h2>
          <div className="flex items-center gap-2">
            <button onClick={() => { setAdding(!adding); setAddMsg(null); }} data-add-customer-toggle="" className="text-xs font-semibold px-3 py-1.5 rounded-lg" style={{ background: "#C9A567", color: "#0B0B0C" }}>{adding ? "Close" : "Add customer"}</button>
            <button onClick={exportCsv} disabled={customers.length === 0} className="text-xs font-medium px-3 py-1.5 rounded-lg border" style={{ borderColor: "#2A2A2F", color: "#C9A567", opacity: customers.length === 0 ? 0.4 : 1 }}>Export CSV</button>
          </div>
        </div>
        <p className="text-[11px] mb-3" style={{ color: "#75726C" }}>Only message people who bought from you or said they're happy to hear from you.</p>
        {adding && (
          <div data-add-customer="" className="rounded-xl border p-3 mb-3" style={{ borderColor: "#2A2A2F", background: "#0B0B0C" }}>
            <div className="grid grid-cols-2 gap-2">
              <input type="tel" inputMode="tel" value={addPhone} onChange={(e) => { setAddPhone(e.target.value); setAddMsg(null); }} aria-label="New customer WhatsApp number" placeholder="WhatsApp 011-2345 6789"
                className="w-full px-3 py-2 rounded-lg border text-xs outline-none" style={{ borderColor: addMsg && addMsg.bad ? "#C9695E" : "#2A2A2F", background: "#17171A", color: "#F2F0EA" }} />
              <input value={addName} maxLength={80} onChange={(e) => { setAddName(e.target.value); setAddMsg(null); }} aria-label="New customer name" placeholder="Name"
                className="w-full px-3 py-2 rounded-lg border text-xs outline-none" style={input} />
            </div>
            <div className="flex items-center gap-3 mt-2.5">
              <button onClick={addOne} disabled={!addPhone.trim()} className="text-xs font-semibold px-3 py-2 rounded-lg" style={{ background: "#C9A567", color: "#0B0B0C", opacity: addPhone.trim() ? 1 : 0.4 }}>Save customer</button>
              {addMsg && <span className="text-[11px]" style={{ color: addMsg.bad ? "#C9695E" : "#C9A567" }}>{addMsg.text}</span>}
            </div>
          </div>
        )}
        {customers.length > 0 && (
          <input value={search} onChange={(e) => setSearch(e.target.value)} placeholder="Search name or number" aria-label="Search customers"
            className="w-full px-3 py-2 rounded-lg border text-xs outline-none mb-3" style={input} />
        )}
        <div className="space-y-2">
          {shown.map((c) => {
            const st = stats[c.id] || { orders: 0, spent: 0, last: null };
            const link = waChatLink(c.phone, renderMessage(promoTemplate, { name: c.name, store: storeLink }));
            return (
              <div key={c.id} data-customer-row="" className="rounded-xl border p-3" style={{ borderColor: "#2A2A2F", background: "#0B0B0C" }}>
                <div className="flex items-start justify-between gap-3">
                  <div className="min-w-0">
                    <p className="text-sm font-medium truncate" style={{ color: "#F2F0EA" }}>{c.name || "No name"}</p>
                    <p className="text-xs" style={{ color: "#8C8A85", fontFamily: "'IBM Plex Mono', monospace" }}>{formatPhone(c.phone)}</p>
                    <p className="text-[11px] mt-0.5" style={{ color: "#75726C" }}>
                      {st.orders} order{st.orders === 1 ? "" : "s"} · RM{st.spent}{st.last ? ` · last ${new Date(st.last).toLocaleDateString(undefined, { day: "2-digit", month: "short" })}` : ""}
                    </p>
                  </div>
                  <div className="flex flex-col gap-1.5 shrink-0">
                    <a data-customer-message="" href={link} target="_blank" rel="noopener noreferrer" className="text-xs font-semibold px-3 py-1.5 rounded-lg text-center" style={{ background: "#C9A567", color: "#0B0B0C", textDecoration: "none" }}>Message</a>
                    <button onClick={() => { setEditId(editId === c.id ? null : c.id); setEditName(c.name); setConfirmId(null); }} className="text-xs px-3 py-1.5 rounded-lg border" style={{ borderColor: "#2A2A2F", color: "#8C8A85" }}>Edit</button>
                  </div>
                </div>
                {editId === c.id && (
                  <div className="mt-3 pt-3 border-t flex flex-wrap items-center gap-2" style={{ borderColor: "#232327" }}>
                    <input value={editName} maxLength={80} onChange={(e) => setEditName(e.target.value)} aria-label="Edit customer name" placeholder="Name" className="flex-1 min-w-[8rem] px-3 py-2 rounded-lg border text-xs outline-none" style={input} />
                    <button onClick={() => { onRename(c.id, editName); setEditId(null); }} className="text-xs font-semibold px-3 py-2 rounded-lg" style={{ background: "#C9A567", color: "#0B0B0C" }}>Save</button>
                    <button onClick={() => { if (confirmId === c.id) { onDelete(c.id); setEditId(null); setConfirmId(null); } else setConfirmId(c.id); }}
                      className="text-xs font-medium px-3 py-2 rounded-lg border" style={{ borderColor: "#2A2A2F", color: "#C9695E" }}>{confirmId === c.id ? "Tap again to delete" : "Delete"}</button>
                  </div>
                )}
              </div>
            );
          })}
          {customers.length === 0 && (
            <p className="text-xs text-center py-6" style={{ color: "#75726C" }}>No customers yet. Tap Add customer, or add a name and WhatsApp number at POS when you sell.</p>
          )}
          {customers.length > 0 && shown.length === 0 && <p className="text-xs text-center py-4" style={{ color: "#75726C" }}>No match.</p>}
        </div>
      </div>
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

function ShopperField({ label, error, ui, children }) {
  return (
    <label className="block">
      <span className="block text-[11px] tracking-[0.14em] uppercase mb-1.5" style={{ color: ui.C.muted }}>{label}</span>
      {children}
      {error && <span role="alert" className="block text-xs mt-1.5" style={{ color: "#B5472F" }}>{error}</span>}
    </label>
  );
}
function ShopperShell({ ui, page, title, intro, wide, children }) {
  return (
    <main data-shopper-page={page} className={`${wide ? "max-w-xl" : "max-w-md"} mx-auto px-5 py-10 sm:py-14`}>
      <h1 className="text-3xl mb-2" style={ui.serif}>{title}</h1>
      {intro && <p className="text-sm mb-7 leading-relaxed" style={{ color: ui.C.muted }}>{intro}</p>}
      {children}
    </main>
  );
}
function ShopperLink({ ui, hash, go, children }) {
  return <a href={hash} onClick={(e) => { e.preventDefault(); go(hash); }} className="text-sm" style={{ color: ui.C.ink, textDecoration: "underline", textUnderlineOffset: 3 }}>{children}</a>;
}
function ShopperButton({ ui, busy, children }) {
  return <button type="submit" disabled={busy} className="w-full text-[11px] tracking-[0.18em] uppercase" style={{ background: ui.C.ink, color: ui.C.bg, padding: "15px 0", opacity: busy ? 0.6 : 1 }}>{children}</button>;
}

function ShopperLogin({ ui, go, onSession, note, returnTo }) {
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [busy, setBusy] = useState(false);
  const [err, setErr] = useState("");
  async function submit(e) {
    e.preventDefault(); setErr("");
    if (!EMAIL_RE.test(email.trim()) || !password) { setErr("Enter your email and password."); return; }
    setBusy(true);
    try { const d = await signInRequest(email.trim(), password); onSession(shopperSessionFrom(d)); go(returnTo && returnTo !== "#login" ? returnTo : "#account"); }
    catch (ex) { setErr(friendlyLoginError(ex.message)); }
    finally { setBusy(false); }
  }
  return (
    <ShopperShell ui={ui} page="login" title="Log in" intro="Welcome back.">
      {note && <p data-shopper-note="" className="text-sm mb-5" style={{ color: ui.C.ink }}>{note}</p>}
      <form onSubmit={submit} noValidate className="space-y-4">
        <ShopperField ui={ui} label="Email"><input name="email" type="email" autoComplete="email" value={email} onChange={(e) => setEmail(e.target.value)} style={ui.input} /></ShopperField>
        <ShopperField ui={ui} label="Password"><input name="password" type="password" autoComplete="current-password" value={password} onChange={(e) => setPassword(e.target.value)} style={ui.input} /></ShopperField>
        {err && <p role="alert" data-form-error="" className="text-sm" style={{ color: "#B5472F" }}>{err}</p>}
        <ShopperButton ui={ui} busy={busy}>{busy ? "Logging in…" : "Log in"}</ShopperButton>
      </form>
      <p className="flex items-center justify-between mt-6">
        <ShopperLink ui={ui} hash="#forgot" go={go}>Forgot password?</ShopperLink>
        <ShopperLink ui={ui} hash="#signup" go={go}>Create an account</ShopperLink>
      </p>
    </ShopperShell>
  );
}

function ShopperSignup({ ui, slug, go, onSession }) {
  const [f, setF] = useState({ fullName: "", phone: "", email: "", password: "", confirmPassword: "", address1: "", address2: "", city: "", state: "", postcode: "", country: "Malaysia", agree: false, marketing: false });
  const [errors, setErrors] = useState({});
  const [busy, setBusy] = useState(false);
  const [formError, setFormError] = useState("");
  const [sent, setSent] = useState(false);
  const set = (k) => (e) => setF((p) => ({ ...p, [k]: e.target.type === "checkbox" ? e.target.checked : e.target.value }));
  const isMY = f.country.trim().toLowerCase() === "malaysia";
  async function submit(e) {
    e.preventDefault(); setFormError("");
    const errs = validateSignup(f);
    setErrors(errs);
    if (Object.keys(errs).length) {
      const first = document.querySelector(`[data-shopper-page="signup"] [name="${Object.keys(errs)[0]}"]`);
      if (first && first.focus) first.focus();
      return;
    }
    setBusy(true);
    try {
      const redirect = `${window.location.origin}/store/${slug}`;
      const d = await authRequest(`/signup?redirect_to=${encodeURIComponent(redirect)}`, { email: f.email.trim(), password: f.password, data: signupPayload(f) });
      if (d.access_token && d.user) { onSession(shopperSessionFrom(d)); go("#account"); }
      else setSent(true);
    } catch (ex) { setFormError(friendlySignupError(ex.message)); }
    finally { setBusy(false); }
  }
  if (sent) {
    return (
      <ShopperShell ui={ui} page="signup-sent" title="Check your email" intro={`We've sent a confirmation link to ${f.email.trim()}. Open it to activate your account, then log in.`}>
        <p className="text-sm mb-6" style={{ color: ui.C.muted }}>Can't find it? Check your spam folder.</p>
        <ShopperLink ui={ui} hash="#login" go={go}>Go to log in</ShopperLink>
      </ShopperShell>
    );
  }
  const inp = (name, extra = {}) => <input name={name} value={f[name]} onChange={set(name)} style={ui.input} {...extra} />;
  const sectionLabel = (t) => <p className="text-[11px] tracking-[0.18em] uppercase pt-3 pb-1 border-b" style={{ color: ui.C.ink, borderColor: ui.C.line }}>{t}</p>;
  return (
    <ShopperShell ui={ui} page="signup" wide title="Create an account" intro="Save your details for faster checkout and track your orders.">
      <form onSubmit={submit} noValidate className="space-y-4">
        {sectionLabel("Your details")}
        <ShopperField ui={ui} label="Full name" error={errors.fullName}>{inp("fullName", { autoComplete: "name" })}</ShopperField>
        <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
          <ShopperField ui={ui} label="Phone number" error={errors.phone}>{inp("phone", { type: "tel", inputMode: "tel", autoComplete: "tel", placeholder: "011-2345 6789" })}</ShopperField>
          <ShopperField ui={ui} label="Email" error={errors.email}>{inp("email", { type: "email", autoComplete: "email" })}</ShopperField>
        </div>
        <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
          <ShopperField ui={ui} label="Password" error={errors.password}>{inp("password", { type: "password", autoComplete: "new-password" })}</ShopperField>
          <ShopperField ui={ui} label="Confirm password" error={errors.confirmPassword}>{inp("confirmPassword", { type: "password", autoComplete: "new-password" })}</ShopperField>
        </div>
        <p className="text-xs -mt-2" style={{ color: ui.C.muted }}>At least 10 characters, with a letter and a number.</p>

        {sectionLabel("Delivery address")}
        <ShopperField ui={ui} label="Address line 1" error={errors.address1}>{inp("address1", { autoComplete: "address-line1" })}</ShopperField>
        <ShopperField ui={ui} label="Address line 2 (optional)" error={errors.address2}>{inp("address2", { autoComplete: "address-line2" })}</ShopperField>
        <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
          <ShopperField ui={ui} label="City" error={errors.city}>{inp("city", { autoComplete: "address-level2" })}</ShopperField>
          <ShopperField ui={ui} label="State" error={errors.state}>
            {isMY ? (
              <select name="state" value={f.state} onChange={set("state")} style={ui.input}>
                <option value="">Select state</option>
                {MY_STATES.map((st) => <option key={st} value={st}>{st}</option>)}
              </select>
            ) : inp("state", { autoComplete: "address-level1" })}
          </ShopperField>
        </div>
        <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
          <ShopperField ui={ui} label="Postcode" error={errors.postcode}>{inp("postcode", { inputMode: "numeric", autoComplete: "postal-code" })}</ShopperField>
          <ShopperField ui={ui} label="Country" error={errors.country}>{inp("country", { autoComplete: "country-name" })}</ShopperField>
        </div>

        <label className="flex items-start gap-3 text-sm pt-2" style={{ color: ui.C.ink }}>
          <input name="agree" type="checkbox" checked={f.agree} onChange={set("agree")} className="mt-1 shrink-0" style={{ width: 18, height: 18 }} />
          <span>I agree that my details are used to process my orders and deliveries.</span>
        </label>
        {errors.agree && <span role="alert" className="block text-xs -mt-2" style={{ color: "#B5472F" }}>{errors.agree}</span>}
        <label className="flex items-start gap-3 text-sm" style={{ color: ui.C.ink }}>
          <input name="marketing" type="checkbox" checked={f.marketing} onChange={set("marketing")} className="mt-1 shrink-0" style={{ width: 18, height: 18 }} />
          <span>Send me news about new drops (optional).</span>
        </label>
        {formError && <p role="alert" data-form-error="" className="text-sm" style={{ color: "#B5472F" }}>{formError}</p>}
        <ShopperButton ui={ui} busy={busy}>{busy ? "Creating account…" : "Create account"}</ShopperButton>
      </form>
      <p className="mt-6 text-sm" style={{ color: ui.C.muted }}>Already have an account? <ShopperLink ui={ui} hash="#login" go={go}>Log in</ShopperLink></p>
    </ShopperShell>
  );
}

function ShopperForgot({ ui, slug, go }) {
  const [email, setEmail] = useState("");
  const [busy, setBusy] = useState(false);
  const [done, setDone] = useState(false);
  const [err, setErr] = useState("");
  async function submit(e) {
    e.preventDefault(); setErr("");
    if (!EMAIL_RE.test(email.trim())) { setErr("Enter a valid email address."); return; }
    setBusy(true);
    try { await authRequest(`/recover?redirect_to=${encodeURIComponent(`${window.location.origin}/store/${slug}`)}`, { email: email.trim() }); }
    catch (ex) { /* never reveal whether an account exists */ }
    setBusy(false); setDone(true);
  }
  return (
    <ShopperShell ui={ui} page="forgot" title="Reset your password" intro={done ? "" : "Enter your email and we'll send you a link to choose a new password."}>
      {done ? (
        <p data-shopper-note="" className="text-sm mb-6" style={{ color: ui.C.ink }}>If an account exists for that email, we've sent a reset link. It can take a minute to arrive.</p>
      ) : (
        <form onSubmit={submit} noValidate className="space-y-4">
          <ShopperField ui={ui} label="Email" error={err}><input name="email" type="email" autoComplete="email" value={email} onChange={(e) => setEmail(e.target.value)} style={ui.input} /></ShopperField>
          <ShopperButton ui={ui} busy={busy}>{busy ? "Sending…" : "Send reset link"}</ShopperButton>
        </form>
      )}
      <p className="mt-6"><ShopperLink ui={ui} hash="#login" go={go}>Back to log in</ShopperLink></p>
    </ShopperShell>
  );
}

function ShopperReset({ ui, recovery, go, onSession, onDone }) {
  const [pw, setPw] = useState("");
  const [pw2, setPw2] = useState("");
  const [busy, setBusy] = useState(false);
  const [err, setErr] = useState("");
  if (!recovery) {
    return (
      <ShopperShell ui={ui} page="reset" title="Link expired" intro="This reset link isn't valid any more. Request a new one.">
        <ShopperLink ui={ui} hash="#forgot" go={go}>Reset my password</ShopperLink>
      </ShopperShell>
    );
  }
  async function submit(e) {
    e.preventDefault(); setErr("");
    const problem = passwordProblem(pw);
    if (problem) { setErr(problem); return; }
    if (pw !== pw2) { setErr("The passwords don't match."); return; }
    setBusy(true);
    try {
      const user = await updatePasswordRequest(recovery.accessToken, pw);
      onSession({ access_token: recovery.accessToken, refresh_token: recovery.refreshToken, user });
      onDone(); go("#account");
    } catch (ex) { setErr("We couldn't update your password. The link may have expired, so request a new one."); }
    finally { setBusy(false); }
  }
  return (
    <ShopperShell ui={ui} page="reset" title="Choose a new password" intro="At least 10 characters, with a letter and a number.">
      <form onSubmit={submit} noValidate className="space-y-4">
        <ShopperField ui={ui} label="New password"><input name="password" type="password" autoComplete="new-password" value={pw} onChange={(e) => setPw(e.target.value)} style={ui.input} /></ShopperField>
        <ShopperField ui={ui} label="Confirm new password"><input name="confirmPassword" type="password" autoComplete="new-password" value={pw2} onChange={(e) => setPw2(e.target.value)} style={ui.input} /></ShopperField>
        {err && <p role="alert" data-form-error="" className="text-sm" style={{ color: "#B5472F" }}>{err}</p>}
        <ShopperButton ui={ui} busy={busy}>{busy ? "Saving…" : "Update password"}</ShopperButton>
      </form>
    </ShopperShell>
  );
}

function ShopperRow({ ui, label, value }) {
  return (
    <div className="flex justify-between gap-4 py-2.5 border-b text-sm" style={{ borderColor: ui.C.line }}>
      <span className="shrink-0" style={{ color: ui.C.muted }}>{label}</span>
      <span className="text-right min-w-0 break-words" style={{ color: ui.C.ink }}>{value || "-"}</span>
    </div>
  );
}
function ShopperEditButton({ ui, kind, onClick }) {
  return <button type="button" data-edit={kind} onClick={onClick} className="mt-5 text-[11px] tracking-[0.18em] uppercase border" style={{ borderColor: ui.C.ink, color: ui.C.ink, padding: "12px 24px" }}>Edit</button>;
}
function ShopperActions({ ui, busy, onCancel, saveLabel }) {
  return (
    <div className="flex gap-3 pt-2">
      <button type="submit" disabled={busy} className="flex-1 text-[11px] tracking-[0.18em] uppercase" style={{ background: ui.C.ink, color: ui.C.bg, padding: "14px 0", opacity: busy ? 0.6 : 1 }}>{busy ? "Saving…" : saveLabel || "Save"}</button>
      {onCancel && <button type="button" onClick={onCancel} className="flex-1 text-[11px] tracking-[0.18em] uppercase border" style={{ borderColor: ui.C.line, color: ui.C.ink, padding: "14px 0" }}>Cancel</button>}
    </div>
  );
}

function ShopperProfileSection({ ui, profile, email, onSave, onFlash }) {
  const [editing, setEditing] = useState(false);
  const [f, setF] = useState({ fullName: "", phone: "", marketing: false });
  const [errors, setErrors] = useState({});
  const [busy, setBusy] = useState(false);
  const [formError, setFormError] = useState("");
  const start = () => { setF({ fullName: profile.full_name || "", phone: profile.phone ? formatPhone(profile.phone) : "", marketing: !!profile.marketing_opt_in }); setErrors({}); setFormError(""); setEditing(true); onFlash(""); };
  async function submit(e) {
    e.preventDefault(); setFormError("");
    const er = {}, name = f.fullName.trim(), phone = normalizePhone(f.phone);
    if (name.length < 2) er.fullName = "Enter your full name."; else if (name.length > 80) er.fullName = "That name is too long.";
    if (!phone) er.phone = "Enter a valid phone number, e.g. 011-2345 6789.";
    setErrors(er);
    if (Object.keys(er).length) return;
    setBusy(true);
    try { await onSave({ full_name: name, phone, marketing_opt_in: !!f.marketing }); setEditing(false); onFlash("Profile saved."); }
    catch (ex) { setFormError("We couldn't save your changes. Please try again."); }
    finally { setBusy(false); }
  }
  if (!editing) {
    return (
      <div data-profile="">
        <ShopperRow ui={ui} label="Name" value={profile.full_name} />
        <ShopperRow ui={ui} label="Email" value={email} />
        <ShopperRow ui={ui} label="Phone" value={profile.phone ? formatPhone(profile.phone) : ""} />
        <ShopperRow ui={ui} label="News about new drops" value={profile.marketing_opt_in ? "Yes" : "No"} />
        <ShopperEditButton ui={ui} kind="profile" onClick={start} />
      </div>
    );
  }
  return (
    <form data-edit-form="profile" onSubmit={submit} noValidate className="space-y-4">
      <ShopperField ui={ui} label="Full name" error={errors.fullName}><input name="fullName" autoComplete="name" value={f.fullName} onChange={(e) => setF({ ...f, fullName: e.target.value })} style={ui.input} /></ShopperField>
      <ShopperField ui={ui} label="Phone number" error={errors.phone}><input name="phone" type="tel" inputMode="tel" autoComplete="tel" value={f.phone} onChange={(e) => setF({ ...f, phone: e.target.value })} style={ui.input} /></ShopperField>
      <p className="text-xs" style={{ color: ui.C.muted }}>Your email ({email}) can't be changed here.</p>
      <label className="flex items-start gap-3 text-sm" style={{ color: ui.C.ink }}>
        <input name="marketing" type="checkbox" checked={f.marketing} onChange={(e) => setF({ ...f, marketing: e.target.checked })} className="mt-1 shrink-0" style={{ width: 18, height: 18 }} />
        <span>Send me news about new drops.</span>
      </label>
      {formError && <p role="alert" data-form-error="" className="text-sm" style={{ color: "#B5472F" }}>{formError}</p>}
      <ShopperActions ui={ui} busy={busy} onCancel={() => setEditing(false)} />
    </form>
  );
}

function ShopperAddressSection({ ui, profile, onSave, onFlash }) {
  const [editing, setEditing] = useState(false);
  const [f, setF] = useState({});
  const [errors, setErrors] = useState({});
  const [busy, setBusy] = useState(false);
  const [formError, setFormError] = useState("");
  const set = (k) => (e) => setF((p) => ({ ...p, [k]: e.target.value }));
  const start = () => { setF({ address1: profile.address_line1 || "", address2: profile.address_line2 || "", city: profile.city || "", state: profile.state || "", postcode: profile.postcode || "", country: profile.country || "Malaysia" }); setErrors({}); setFormError(""); setEditing(true); onFlash(""); };
  async function submit(e) {
    e.preventDefault(); setFormError("");
    const er = validateAddress(f);
    setErrors(er);
    if (Object.keys(er).length) return;
    setBusy(true);
    try {
      await onSave({ address_line1: f.address1.trim(), address_line2: f.address2.trim() || null, city: f.city.trim(), state: f.state.trim(), postcode: f.postcode.trim(), country: f.country.trim() });
      setEditing(false); onFlash("Address saved.");
    } catch (ex) { setFormError("We couldn't save your address. Please try again."); }
    finally { setBusy(false); }
  }
  if (!editing) {
    const lines = [profile.address_line1, profile.address_line2, [profile.postcode, profile.city].filter(Boolean).join(" "), profile.state, profile.country].filter(Boolean);
    return (
      <div data-address="">
        <p className="text-[11px] tracking-[0.14em] uppercase mb-2" style={{ color: ui.C.muted }}>Delivery address</p>
        {lines.length ? <p className="text-sm leading-relaxed" style={{ color: ui.C.ink }}>{lines.map((l, i) => <span key={i} className="block">{l}</span>)}</p> : <p className="text-sm" style={{ color: ui.C.muted }}>No address saved yet.</p>}
        <ShopperEditButton ui={ui} kind="address" onClick={start} />
      </div>
    );
  }
  const isMY = String(f.country || "").trim().toLowerCase() === "malaysia";
  const inp = (name, extra = {}) => <input name={name} value={f[name] || ""} onChange={set(name)} style={ui.input} {...extra} />;
  return (
    <form data-edit-form="address" onSubmit={submit} noValidate className="space-y-4">
      <ShopperField ui={ui} label="Address line 1" error={errors.address1}>{inp("address1", { autoComplete: "address-line1" })}</ShopperField>
      <ShopperField ui={ui} label="Address line 2 (optional)" error={errors.address2}>{inp("address2", { autoComplete: "address-line2" })}</ShopperField>
      <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
        <ShopperField ui={ui} label="City" error={errors.city}>{inp("city", { autoComplete: "address-level2" })}</ShopperField>
        <ShopperField ui={ui} label="State" error={errors.state}>
          {isMY ? (
            <select name="state" value={f.state || ""} onChange={set("state")} style={ui.input}>
              <option value="">Select state</option>
              {MY_STATES.map((st) => <option key={st} value={st}>{st}</option>)}
            </select>
          ) : inp("state", { autoComplete: "address-level1" })}
        </ShopperField>
      </div>
      <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
        <ShopperField ui={ui} label="Postcode" error={errors.postcode}>{inp("postcode", { inputMode: "numeric", autoComplete: "postal-code" })}</ShopperField>
        <ShopperField ui={ui} label="Country" error={errors.country}>{inp("country", { autoComplete: "country-name" })}</ShopperField>
      </div>
      {formError && <p role="alert" data-form-error="" className="text-sm" style={{ color: "#B5472F" }}>{formError}</p>}
      <ShopperActions ui={ui} busy={busy} onCancel={() => setEditing(false)} />
    </form>
  );
}

function ShopperPasswordSection({ ui, shopper, onSession, onFlash }) {
  const [f, setF] = useState({ current: "", next: "", confirm: "" });
  const [errors, setErrors] = useState({});
  const [busy, setBusy] = useState(false);
  const [formError, setFormError] = useState("");
  const set = (k) => (e) => setF((p) => ({ ...p, [k]: e.target.value }));
  async function submit(e) {
    e.preventDefault(); setFormError(""); onFlash("");
    const er = {}, problem = passwordProblem(f.next);
    if (!f.current) er.current = "Enter your current password.";
    if (problem) er.next = problem; else if (f.next === f.current) er.next = "Choose a different password from your current one.";
    if (!f.confirm) er.confirm = "Confirm your new password."; else if (!problem && f.next !== f.confirm) er.confirm = "The passwords don't match.";
    setErrors(er);
    if (Object.keys(er).length) return;
    setBusy(true);
    try {
      let d;
      try { d = await signInRequest(shopper.user.email, f.current); } // proves they know the current password
      catch (ex) { setErrors({ current: String(ex.message).toLowerCase().includes("invalid login") ? "Your current password is incorrect." : friendlyLoginError(ex.message) }); return; }
      const user = await updatePasswordRequest(d.access_token, f.next);
      onSession({ access_token: d.access_token, refresh_token: d.refresh_token, user });
      setF({ current: "", next: "", confirm: "" }); setErrors({});
      onFlash("Password updated.");
    } catch (ex) { setFormError("We couldn't update your password. Please try again."); }
    finally { setBusy(false); }
  }
  return (
    <form data-edit-form="password" onSubmit={submit} noValidate className="space-y-4">
      <p className="text-[11px] tracking-[0.14em] uppercase" style={{ color: ui.C.muted }}>Change password</p>
      <ShopperField ui={ui} label="Current password" error={errors.current}><input name="current" type="password" autoComplete="current-password" value={f.current} onChange={set("current")} style={ui.input} /></ShopperField>
      <ShopperField ui={ui} label="New password" error={errors.next}><input name="next" type="password" autoComplete="new-password" value={f.next} onChange={set("next")} style={ui.input} /></ShopperField>
      <ShopperField ui={ui} label="Confirm new password" error={errors.confirm}><input name="confirm" type="password" autoComplete="new-password" value={f.confirm} onChange={set("confirm")} style={ui.input} /></ShopperField>
      <p className="text-xs" style={{ color: ui.C.muted }}>At least 10 characters, with a letter and a number.</p>
      {formError && <p role="alert" data-form-error="" className="text-sm" style={{ color: "#B5472F" }}>{formError}</p>}
      <ShopperActions ui={ui} busy={busy} saveLabel="Update password" />
    </form>
  );
}

function ShopperAccount({ ui, shopper, onSession, onSignOut, note, checkoutOpen, go }) {
  const [profile, setProfile] = useState(undefined); // undefined = loading, null = no profile row
  const [err, setErr] = useState("");
  const [section, setSection] = useState("profile");
  const [flash, setFlash] = useState("");
  useEffect(() => {
    let on = true;
    (async () => {
      try {
        const rows = await apiRequest("/shopper_profiles?select=*", { method: "GET" }, shopper, onSession);
        if (on) setProfile(rows && rows[0] ? rows[0] : null);
      } catch (e) { if (on) setErr("We couldn't load your details. Please try again."); }
    })();
    return () => { on = false; };
  }, [shopper && shopper.access_token]);
  // saves only the columns passed in; the database lets a shopper change nothing else
  const save = async (patch) => {
    const rows = await apiRequest(`/shopper_profiles?user_id=eq.${shopper.user.id}`, { method: "PATCH", body: JSON.stringify(patch) }, shopper, onSession);
    setProfile((rows && rows[0]) || { ...profile, ...patch });
  };
  const tabs = [["profile", "Profile"], ["address", "Address"], ...(checkoutOpen ? [["orders", "Orders"]] : []), ["settings", "Settings"]];
  return (
    <ShopperShell ui={ui} page="account" wide title="My account" intro={shopper.user && shopper.user.email ? `Signed in as ${shopper.user.email}` : ""}>
      {note && <p data-shopper-note="" className="text-sm mb-5" style={{ color: ui.C.ink }}>{note}</p>}
      {err && <p role="alert" className="text-sm mb-4" style={{ color: ui.C.red }}>{err}</p>}
      {profile === undefined && !err && <p className="text-sm" style={{ color: ui.C.muted }}>Loading…</p>}
      {profile === null && <p data-no-profile="" className="text-sm leading-relaxed mb-4" style={{ color: ui.C.muted }}>This account doesn't have a shopper profile. If you manage this store, use the admin app instead.</p>}
      {profile && (
        <>
          <div role="tablist" className="flex gap-7 border-b mb-6 overflow-x-auto" style={{ borderColor: ui.C.line }}>
            {tabs.map(([k, label]) => {
              const on = section === k;
              return (
                <button key={k} role="tab" aria-selected={on} data-account-tab={k} onClick={() => { setSection(k); setFlash(""); }}
                  className="text-[11px] tracking-[0.18em] uppercase pb-3 whitespace-nowrap"
                  style={{ color: on ? ui.C.ink : ui.C.muted, borderBottom: on ? `2px solid ${ui.C.ink}` : "2px solid transparent", marginBottom: -1 }}>{label}</button>
              );
            })}
          </div>
          {flash && <p data-flash="" role="status" className="text-sm mb-4" style={{ color: ui.C.ink }}>✓ {flash}</p>}
          {section === "profile" && <ShopperProfileSection ui={ui} profile={profile} email={shopper.user && shopper.user.email} onSave={save} onFlash={setFlash} />}
          {section === "orders" && <ShopperOrdersSection ui={ui} shopper={shopper} onSession={onSession} go={go} />}
          {section === "address" && <ShopperAddressSection ui={ui} profile={profile} onSave={save} onFlash={setFlash} />}
          {section === "settings" && <ShopperPasswordSection ui={ui} shopper={shopper} onSession={onSession} onFlash={setFlash} />}
        </>
      )}
      <button onClick={onSignOut} className="mt-10 w-full text-[11px] tracking-[0.18em] uppercase border" style={{ borderColor: ui.C.ink, color: ui.C.ink, padding: "14px 0" }}>Log out</button>
    </ShopperShell>
  );
}

const rm = (n) => fmtRM(Number(n || 0));
const fmtWhen = (d) => new Date(d).toLocaleString(undefined, { day: "2-digit", month: "short", hour: "2-digit", minute: "2-digit" });
function ShopperThumb({ ui, src, size }) {
  return (
    <div className="relative shrink-0 overflow-hidden" style={{ width: size, height: size, background: ui.C.tile }}>
      {src ? <img src={src} alt="" decoding="async" className="absolute inset-0 w-full h-full object-cover" /> : null}
    </div>
  );
}
function ShopperTotals({ ui, quote, fallbackSubtotal, shippingNote }) {
  // every figure comes from the shop's database; until it answers we show "…" rather than a total that might be wrong
  const sub = quote ? quote.subtotal : fallbackSubtotal;
  const ship = quote ? (quote.free_shipping ? "Free" : rm(quote.shipping_fee)) : "…";
  const total = quote ? rm(quote.total) : "…";
  const remaining = quote && !quote.free_shipping && quote.free_ship_min != null ? Math.round((Number(quote.free_ship_min) - Number(quote.subtotal)) * 100) / 100 : 0;
  const row = (key, label, value, strong) => (
    <div className={`flex justify-between ${strong ? "pt-3 mt-1 border-t text-base" : ""}`} style={{ borderColor: ui.C.line, color: strong ? ui.C.ink : ui.C.muted }}>
      <span>{label}</span><span data-total-row={key} style={{ color: ui.C.ink }}>{value}</span>
    </div>
  );
  return (
    <div data-bag-summary="" className="space-y-2 text-sm">
      {row("Subtotal", "Subtotal", rm(sub))}
      {row("Shipping", shippingNote || "Shipping", ship)}
      {row("Total", "Total", total, true)}
      {remaining > 0 && <p data-free-ship-hint="" className="text-xs pt-1" style={{ color: ui.C.muted }}>Spend {rm(remaining)} more for free shipping.</p>}
    </div>
  );
}

function ShopperBag({ ui, shopper, bag, onReload, onRemove, onSession, go }) {
  const [quote, setQuote] = useState(null);
  const [profileState, setProfileState] = useState(null); // null = not loaded yet
  useEffect(() => { onReload(); }, []);
  useEffect(() => {
    let on = true;
    (async () => {
      try { const rows = await apiRequest("/shopper_profiles?select=state", { method: "GET" }, shopper, onSession); if (on) setProfileState((rows && rows[0] && rows[0].state) || ""); }
      catch (e) { if (on) setProfileState(""); }
    })();
    return () => { on = false; };
  }, []);
  const key = bag ? bag.map((b) => b.item_id + (b.available ? "1" : "0")).join(",") : "";
  useEffect(() => {
    if (profileState === null || !bag || bag.length === 0) return;
    let on = true;
    apiRequest("/rpc/checkout_quote", { method: "POST", body: JSON.stringify({ p_state: profileState || null }) }, shopper, onSession).then((q) => { if (on) setQuote(q); }).catch(() => {});
    return () => { on = false; };
  }, [profileState, key]);
  if (!bag) return <ShopperShell ui={ui} page="bag" title="Your bag"><p className="text-sm" style={{ color: ui.C.muted }}>Loading…</p></ShopperShell>;
  if (bag.length === 0) {
    return (
      <ShopperShell ui={ui} page="bag" title="Your bag" intro="Your bag is empty.">
        <ShopperLink ui={ui} hash="" go={go}>Continue shopping</ShopperLink>
      </ShopperShell>
    );
  }
  const okAll = bag.every((b) => b.available);
  const fallback = bag.filter((b) => b.available).reduce((sum, b) => sum + Number(b.final_price || 0), 0);
  return (
    <ShopperShell ui={ui} page="bag" wide title="Your bag" intro={`${bag.length} piece${bag.length === 1 ? "" : "s"}. Every piece is one of a kind, and it's only held for you once you place your order.`}>
      <div className="mb-6">
        {bag.map((b) => (
          <div key={b.item_id} data-bag-line={b.item_id} className="flex gap-4 py-4 border-b" style={{ borderColor: ui.C.line }}>
            <ShopperThumb ui={ui} src={b.photo} size={84} />
            <div className="min-w-0 flex-1">
              {b.brand && <p className="text-[10px] tracking-[0.16em] uppercase truncate" style={{ color: ui.C.muted }}>{b.brand}</p>}
              <p className="text-sm leading-snug" style={{ color: ui.C.ink }}>{b.title}</p>
              {(b.size || b.colour) && <p className="text-xs mt-0.5" style={{ color: ui.C.muted }}>{[b.size && `Size ${b.size}`, b.colour].filter(Boolean).join(" · ")}</p>}
              {b.available ? (
                <p className="text-sm mt-1.5">
                  {Number(b.final_price) !== Number(b.price) ? (<><span style={{ color: ui.C.red }}>{rm(b.final_price)}</span> <span className="ml-1" style={{ color: ui.C.muted, textDecoration: "line-through" }}>{rm(b.price)}</span></>) : <span style={{ color: ui.C.ink }}>{rm(b.price)}</span>}
                </p>
              ) : (
                <p data-unavailable="" className="text-xs mt-1.5" style={{ color: ui.C.red }}>{b.reason === "sold" ? "Sorry, this piece has sold." : b.reason === "reserved" ? "Another customer is holding this piece." : "This piece is no longer available."}</p>
              )}
            </div>
            <button data-remove="" onClick={() => onRemove(b.item_id)} aria-label={`Remove ${b.title}`} className="self-start text-[11px] tracking-[0.14em] uppercase" style={{ color: ui.C.muted }}>Remove</button>
          </div>
        ))}
      </div>
      <ShopperTotals ui={ui} quote={quote} fallbackSubtotal={fallback} shippingNote={profileState ? "Shipping" : "Shipping (West Malaysia rate; confirmed at checkout)"} />
      {!okAll && <p role="alert" data-bag-blocked="" className="text-sm mt-5" style={{ color: ui.C.red }}>Please remove the unavailable pieces to continue.</p>}
      <button data-checkout-continue="" onClick={() => go("#checkout")} disabled={!okAll} className="mt-6 w-full text-[11px] tracking-[0.18em] uppercase" style={{ background: ui.C.ink, color: ui.C.bg, padding: "16px 0", opacity: okAll ? 1 : 0.4 }}>Checkout</button>
      <p className="mt-5 text-center"><ShopperLink ui={ui} hash="" go={go}>Continue shopping</ShopperLink></p>
    </ShopperShell>
  );
}

function ShopperCheckout({ ui, shopper, bag, onReload, onSession, go }) {
  const [f, setF] = useState(null); // null until the saved profile has loaded
  const [step, setStep] = useState("details");
  const [errors, setErrors] = useState({});
  const [quote, setQuote] = useState(null);
  const [busy, setBusy] = useState(false);
  const [formError, setFormError] = useState("");
  useEffect(() => { onReload(); }, []);
  useEffect(() => {
    let on = true;
    (async () => {
      let pr = null;
      try { const rows = await apiRequest("/shopper_profiles?select=*", { method: "GET" }, shopper, onSession); pr = rows && rows[0]; } catch (e) {}
      if (on) setF({ fullName: (pr && pr.full_name) || "", phone: pr && pr.phone ? formatPhone(pr.phone) : "", address1: (pr && pr.address_line1) || "", address2: (pr && pr.address_line2) || "", city: (pr && pr.city) || "", state: (pr && pr.state) || "", postcode: (pr && pr.postcode) || "", country: (pr && pr.country) || "Malaysia", note: "" });
    })();
    return () => { on = false; };
  }, []);
  const bagKey = bag ? bag.map((b) => b.item_id + (b.available ? "1" : "0")).join(",") : "";
  useEffect(() => {
    if (!f || !bag || bag.length === 0) return;
    let on = true;
    apiRequest("/rpc/checkout_quote", { method: "POST", body: JSON.stringify({ p_state: f.state || null }) }, shopper, onSession).then((q) => { if (on) setQuote(q); }).catch(() => {});
    return () => { on = false; };
  }, [f && f.state, bagKey]);
  if (!f || !bag) return <ShopperShell ui={ui} page="checkout" title="Checkout"><p className="text-sm" style={{ color: ui.C.muted }}>Loading…</p></ShopperShell>;
  if (bag.length === 0) {
    return (
      <ShopperShell ui={ui} page="checkout" title="Checkout" intro="Your bag is empty.">
        <ShopperLink ui={ui} hash="" go={go}>Continue shopping</ShopperLink>
      </ShopperShell>
    );
  }
  const set = (k) => (e) => setF((p) => ({ ...p, [k]: e.target.value }));
  const isMY = f.country.trim().toLowerCase() === "malaysia";
  const inp = (name, extra = {}) => <input name={name} value={f[name]} onChange={set(name)} style={ui.input} {...extra} />;
  const allOk = !!quote && quote.all_available && bag.every((b) => b.available);
  function review(e) {
    e.preventDefault(); setFormError("");
    const er = validateCheckout(f);
    setErrors(er);
    if (Object.keys(er).length) { const first = document.querySelector(`[data-shopper-page="checkout"] [name="${Object.keys(er)[0]}"]`); if (first && first.focus) first.focus(); return; }
    setStep("review"); window.scrollTo(0, 0);
  }
  async function place() {
    setBusy(true); setFormError("");
    try {
      const res = await apiRequest("/rpc/place_order", { method: "POST", body: JSON.stringify({ p_name: f.fullName.trim(), p_phone: normalizePhone(f.phone), p_line1: f.address1.trim(), p_line2: f.address2.trim() || null, p_city: f.city.trim(), p_state: f.state.trim(), p_postcode: f.postcode.trim(), p_country: f.country.trim(), p_note: f.note.trim() || null }) }, shopper, onSession);
      try { sessionStorage.setItem("ravuno_just_placed", res.order_id); } catch (e) {}
      await onReload();
      go(`#order=${res.order_id}`);
    } catch (ex) {
      const msg = ex.message || "We couldn't place your order.";
      setFormError(msg);
      if (/no longer available/i.test(msg)) await onReload();
    } finally { setBusy(false); }
  }
  const lines = (
    <div className="mb-5">
      {bag.map((b) => (
        <div key={b.item_id} data-bag-line={b.item_id} className="flex gap-3 py-3 border-b items-center" style={{ borderColor: ui.C.line }}>
          <ShopperThumb ui={ui} src={b.photo} size={56} />
          <div className="min-w-0 flex-1"><p className="text-sm leading-snug truncate" style={{ color: ui.C.ink }}>{b.title}</p>{b.size && <p className="text-xs" style={{ color: ui.C.muted }}>Size {b.size}</p>}</div>
          <span className="text-sm" style={{ color: b.available ? ui.C.ink : "#B5472F" }}>{b.available ? rm(b.final_price) : "Unavailable"}</span>
        </div>
      ))}
    </div>
  );
  if (step === "review") {
    return (
      <ShopperShell ui={ui} page="checkout-review" wide title="Review your order" intro="Check everything before you place your order.">
        {lines}
        <ShopperTotals ui={ui} quote={quote} fallbackSubtotal={bag.filter((b) => b.available).reduce((sum, b) => sum + Number(b.final_price || 0), 0)} />
        <div className="mt-7 space-y-5 text-sm" style={{ color: ui.C.ink }}>
          <div data-review-contact=""><p className="text-[11px] tracking-[0.14em] uppercase mb-1.5" style={{ color: ui.C.muted }}>Contact</p>{f.fullName.trim()}<br />{formatPhone(normalizePhone(f.phone))}<br />{shopper.user && shopper.user.email}</div>
          <div data-review-address=""><p className="text-[11px] tracking-[0.14em] uppercase mb-1.5" style={{ color: ui.C.muted }}>Deliver to</p>{[f.address1.trim(), f.address2.trim(), `${f.postcode.trim()} ${f.city.trim()}`, f.state.trim(), f.country.trim()].filter(Boolean).map((l, i) => <span key={i} className="block">{l}</span>)}</div>
          {f.note.trim() && <div><p className="text-[11px] tracking-[0.14em] uppercase mb-1.5" style={{ color: ui.C.muted }}>Your note</p>{f.note.trim()}</div>}
          <p className="text-xs leading-relaxed" style={{ color: ui.C.muted }}>After you place your order we'll show you how to pay by DuitNow or bank transfer, and your pieces are held for you while you do.</p>
        </div>
        {formError && <p role="alert" data-form-error="" className="text-sm mt-5" style={{ color: ui.C.red }}>{formError}{/no longer available/i.test(formError) && <> <ShopperLink ui={ui} hash="#bag" go={go}>Review your bag</ShopperLink></>}</p>}
        {!allOk && !formError && <p role="alert" className="text-sm mt-5" style={{ color: ui.C.red }}>Some pieces are unavailable. <ShopperLink ui={ui} hash="#bag" go={go}>Review your bag</ShopperLink></p>}
        <div className="flex gap-3 pt-6">
          <button data-place-order="" onClick={place} disabled={busy || !allOk} className="flex-1 text-[11px] tracking-[0.18em] uppercase" style={{ background: ui.C.ink, color: ui.C.bg, padding: "16px 0", opacity: busy || !allOk ? 0.5 : 1 }}>{busy ? "Placing order…" : "Place order"}</button>
          <button type="button" onClick={() => setStep("details")} className="flex-1 text-[11px] tracking-[0.18em] uppercase border" style={{ borderColor: ui.C.line, color: ui.C.ink, padding: "16px 0" }}>Edit details</button>
        </div>
      </ShopperShell>
    );
  }
  return (
    <ShopperShell ui={ui} page="checkout" wide title="Checkout" intro="Delivery within Malaysia.">
      {lines}
      <ShopperTotals ui={ui} quote={quote} fallbackSubtotal={bag.filter((b) => b.available).reduce((sum, b) => sum + Number(b.final_price || 0), 0)} />
      <form onSubmit={review} noValidate className="space-y-4 mt-8">
        <p className="text-[11px] tracking-[0.18em] uppercase pt-1 pb-1 border-b" style={{ color: ui.C.ink, borderColor: ui.C.line }}>Contact</p>
        <ShopperField ui={ui} label="Full name" error={errors.fullName}>{inp("fullName", { autoComplete: "name" })}</ShopperField>
        <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
          <ShopperField ui={ui} label="Phone number" error={errors.phone}>{inp("phone", { type: "tel", inputMode: "tel", autoComplete: "tel" })}</ShopperField>
          <ShopperField ui={ui} label="Email"><input value={(shopper.user && shopper.user.email) || ""} readOnly style={{ ...ui.input, color: ui.C.muted }} /></ShopperField>
        </div>
        <p className="text-[11px] tracking-[0.18em] uppercase pt-3 pb-1 border-b" style={{ color: ui.C.ink, borderColor: ui.C.line }}>Delivery address</p>
        <ShopperField ui={ui} label="Address line 1" error={errors.address1}>{inp("address1", { autoComplete: "address-line1" })}</ShopperField>
        <ShopperField ui={ui} label="Address line 2 (optional)" error={errors.address2}>{inp("address2", { autoComplete: "address-line2" })}</ShopperField>
        <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
          <ShopperField ui={ui} label="City" error={errors.city}>{inp("city", { autoComplete: "address-level2" })}</ShopperField>
          <ShopperField ui={ui} label="State" error={errors.state}>
            {isMY ? (<select name="state" value={f.state} onChange={set("state")} style={ui.input}><option value="">Select state</option>{MY_STATES.map((st) => <option key={st} value={st}>{st}</option>)}</select>) : inp("state")}
          </ShopperField>
        </div>
        <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
          <ShopperField ui={ui} label="Postcode" error={errors.postcode}>{inp("postcode", { inputMode: "numeric", autoComplete: "postal-code" })}</ShopperField>
          <ShopperField ui={ui} label="Country" error={errors.country}>{inp("country", { autoComplete: "country-name" })}</ShopperField>
        </div>
        <ShopperField ui={ui} label="Note for the seller (optional)" error={errors.note}><textarea name="note" rows={2} maxLength={300} value={f.note} onChange={set("note")} style={ui.input} /></ShopperField>
        <ShopperButton ui={ui} busy={false}>Review order</ShopperButton>
      </form>
      <p className="mt-5 text-center"><ShopperLink ui={ui} hash="#bag" go={go}>Back to bag</ShopperLink></p>
    </ShopperShell>
  );
}

function ShopperOrder({ ui, shopper, orderId, onSession, go }) {
  const [order, setOrder] = useState(undefined); // undefined = loading, null = not found
  const [err, setErr] = useState("");
  const [busy, setBusy] = useState(false);
  const [confirmCancel, setConfirmCancel] = useState(false);
  const [now, setNow] = useState(Date.now());
  const [justPlaced, setJustPlaced] = useState(false);
  const load = async () => {
    try {
      await apiRequest("/rpc/refresh_orders", { method: "POST", body: "{}" }, shopper, onSession).catch(() => {});
      const rows = await apiRequest(`/orders?id=eq.${orderId}&select=*,order_items(*),order_events(*)`, { method: "GET" }, shopper, onSession);
      setOrder(rows && rows[0] ? rows[0] : null);
    } catch (e) { setErr("We couldn't load this order."); }
  };
  useEffect(() => {
    setOrder(undefined); setErr(""); setConfirmCancel(false); setJustPlaced(false);   // never show the previous order while the next one loads
    try { if (sessionStorage.getItem("ravuno_just_placed") === orderId) { setJustPlaced(true); sessionStorage.removeItem("ravuno_just_placed"); } } catch (e) {}
    load();
  }, [orderId]);
  useEffect(() => { const t = setInterval(() => setNow(Date.now()), 30000); return () => clearInterval(t); }, []);
  async function act(fn) {
    setBusy(true); setErr("");
    try { await apiRequest(`/rpc/${fn}`, { method: "POST", body: JSON.stringify({ p_order: orderId }) }, shopper, onSession); setConfirmCancel(false); await load(); }
    catch (e) { setErr(e.message || "Something went wrong."); }
    finally { setBusy(false); }
  }
  if (order === undefined && !err) return <ShopperShell ui={ui} page="order" title="Your order"><p className="text-sm" style={{ color: ui.C.muted }}>Loading…</p></ShopperShell>;
  if (!order) {
    return (
      <ShopperShell ui={ui} page="order" title="Order not found" intro={err || "We couldn't find this order."}>
        <ShopperLink ui={ui} hash="#account" go={go}>Back to my account</ShopperLink>
      </ShopperShell>
    );
  }
  const lapsed = order.payment_status === "pending" && new Date(order.expires_at).getTime() <= now;
  const expired = lapsed || order.payment_status === "expired";
  const label = expired ? "Expired" : orderLabel(order);
  const awaiting = order.payment_status === "pending" && !lapsed && order.status === "placed";
  const events = [...(order.order_events || [])].sort((x, y) => new Date(x.created_at) - new Date(y.created_at));
  const lineItems = order.order_items || [];
  return (
    <ShopperShell ui={ui} page="order" wide title={`Order ${order.order_number}`} intro={`Placed ${fmtWhen(order.created_at)}`}>
      {justPlaced && <p data-shopper-note="" className="text-sm mb-5" style={{ color: ui.C.ink }}>Thank you! Your order is placed and your pieces are being held for you.</p>}
      {err && <p role="alert" className="text-sm mb-4" style={{ color: ui.C.red }}>{err}</p>}
      <p className="text-[11px] tracking-[0.18em] uppercase mb-5" style={{ color: ui.C.ink }}>Status: <span data-order-status="">{label}</span></p>

      {(awaiting || order.payment_status === "claimed") && (
        <div data-pay-box="" className="border p-5 mb-7" style={{ borderColor: ui.C.ink }}>
          {awaiting ? (
            <>
              <p className="text-sm mb-1" style={{ color: ui.C.ink }}>Please pay <strong data-pay-amount="">{rm(order.total)}</strong> within <strong data-time-left="">{timeLeft(order.expires_at, now)}</strong> to keep your pieces.</p>
              <p className="text-xs mb-4" style={{ color: ui.C.muted }}>Use <strong>{order.order_number}</strong> as the payment reference.</p>
              <pre data-pay-instructions="" className="whitespace-pre-wrap text-sm p-4 mb-4" style={{ background: ui.C.tile, color: ui.C.ink, fontFamily: "inherit" }}>{order.payment_instructions}</pre>
              <button data-ive-paid="" onClick={() => act("order_claim_paid")} disabled={busy} className="w-full text-[11px] tracking-[0.18em] uppercase" style={{ background: ui.C.ink, color: ui.C.bg, padding: "15px 0", opacity: busy ? 0.6 : 1 }}>I've paid</button>
              <p className="text-xs mt-3" style={{ color: ui.C.muted }}>Tap this once your transfer is done. We check our bank before processing your order.</p>
            </>
          ) : (
            <p data-claimed-note="" className="text-sm" style={{ color: ui.C.ink }}>Thanks! We'll check your payment of {rm(order.total)} and update this order. You'll see it change here.</p>
          )}
        </div>
      )}
      {expired && <p data-expired-note="" className="text-sm mb-6" style={{ color: ui.C.red }}>This order has expired because payment wasn't received in time, and the pieces went back on sale.</p>}

      <div className="mb-6">
        {lineItems.map((li) => (
          <div key={li.id} data-order-line="" className="flex gap-3 py-3 border-b items-center" style={{ borderColor: ui.C.line }}>
            <ShopperThumb ui={ui} src={li.photo_url} size={56} />
            <div className="min-w-0 flex-1"><p className="text-sm leading-snug" style={{ color: ui.C.ink }}>{li.title}</p>{li.size && <p className="text-xs" style={{ color: ui.C.muted }}>Size {li.size}</p>}</div>
            <span className="text-sm" style={{ color: ui.C.ink }}>{rm(li.price)}</span>
          </div>
        ))}
      </div>
      <div className="space-y-2 text-sm mb-7" data-order-totals="">
        <div className="flex justify-between" style={{ color: ui.C.muted }}><span>Subtotal</span><span style={{ color: ui.C.ink }}>{rm(order.subtotal)}</span></div>
        <div className="flex justify-between" style={{ color: ui.C.muted }}><span>Shipping</span><span style={{ color: ui.C.ink }}>{Number(order.shipping_fee) === 0 ? "Free" : rm(order.shipping_fee)}</span></div>
        <div className="flex justify-between pt-3 mt-1 border-t text-base" style={{ borderColor: ui.C.line, color: ui.C.ink }}><span>Total</span><span>{rm(order.total)}</span></div>
      </div>
      <div className="space-y-5 text-sm mb-8" style={{ color: ui.C.ink }}>
        <div data-order-address=""><p className="text-[11px] tracking-[0.14em] uppercase mb-1.5" style={{ color: ui.C.muted }}>Deliver to</p>
          {order.customer_name}<br />{[order.ship_line1, order.ship_line2, `${order.ship_postcode} ${order.ship_city}`, order.ship_state, order.ship_country].filter(Boolean).map((l, i) => <span key={i} className="block">{l}</span>)}</div>
        {(order.tracking_no || order.courier) && (
          <div data-tracking=""><p className="text-[11px] tracking-[0.14em] uppercase mb-1.5" style={{ color: ui.C.muted }}>Tracking</p>
            {order.courier ? <span className="block">Courier: {order.courier}</span> : null}{order.tracking_no ? <span className="block">Tracking number: <strong>{order.tracking_no}</strong></span> : null}</div>
        )}
        {order.customer_note && <div><p className="text-[11px] tracking-[0.14em] uppercase mb-1.5" style={{ color: ui.C.muted }}>Your note</p>{order.customer_note}</div>}
      </div>
      <p className="text-[11px] tracking-[0.18em] uppercase pb-2 border-b" style={{ color: ui.C.ink, borderColor: ui.C.line }}>Order history</p>
      <ol data-timeline="" className="mb-8">
        {events.map((ev) => (
          <li key={ev.id} className="flex justify-between gap-4 py-2.5 border-b text-sm" style={{ borderColor: ui.C.line }}>
            <span style={{ color: ui.C.ink }}>{ev.label}</span><span className="shrink-0 text-xs" style={{ color: ui.C.muted }}>{fmtWhen(ev.created_at)}</span>
          </li>
        ))}
      </ol>
      {awaiting && (
        <button data-cancel-order="" onClick={() => (confirmCancel ? act("order_cancel") : setConfirmCancel(true))} disabled={busy} className="w-full text-[11px] tracking-[0.18em] uppercase border mb-3" style={{ borderColor: ui.C.line, color: ui.C.red, padding: "14px 0" }}>{confirmCancel ? "Tap again to cancel this order" : "Cancel order"}</button>
      )}
      <p className="text-center"><ShopperLink ui={ui} hash="#account" go={go}>Back to my account</ShopperLink></p>
    </ShopperShell>
  );
}

function ShopperOrdersSection({ ui, shopper, onSession, go }) {
  const [orders, setOrders] = useState(undefined);
  const [err, setErr] = useState("");
  useEffect(() => {
    let on = true;
    (async () => {
      try {
        await apiRequest("/rpc/refresh_orders", { method: "POST", body: "{}" }, shopper, onSession).catch(() => {});
        const rows = await apiRequest("/orders?select=id,order_number,status,payment_status,total,created_at,expires_at,courier,tracking_no,order_items(title)&order=created_at.desc", { method: "GET" }, shopper, onSession);
        if (on) setOrders(rows || []);
      } catch (e) { if (on) setErr("We couldn't load your orders."); }
    })();
    return () => { on = false; };
  }, []);
  if (err) return <p role="alert" className="text-sm" style={{ color: ui.C.red }}>{err}</p>;
  if (orders === undefined) return <p className="text-sm" style={{ color: ui.C.muted }}>Loading…</p>;
  if (orders.length === 0) return <div data-orders-empty=""><p className="text-sm mb-4" style={{ color: ui.C.muted }}>You haven't placed any orders yet.</p><ShopperLink ui={ui} hash="" go={go}>Start shopping</ShopperLink></div>;
  return (
    <div data-orders="">
      {orders.map((o) => (
        <a key={o.id} data-order-row={o.id} href={`#order=${o.id}`} onClick={(e) => { e.preventDefault(); go(`#order=${o.id}`); }} className="block py-4 border-b" style={{ borderColor: ui.C.line, color: "inherit", textDecoration: "none" }}>
          <div className="flex justify-between gap-4 text-sm"><span style={{ color: ui.C.ink }}>{o.order_number}</span><span style={{ color: ui.C.ink }}>{rm(o.total)}</span></div>
          <div className="flex justify-between gap-4 text-xs mt-1" style={{ color: ui.C.muted }}><span>{fmtWhen(o.created_at)} · {(o.order_items || []).length} piece{(o.order_items || []).length === 1 ? "" : "s"}</span><span data-order-label="">{o.payment_status === "pending" && o.status === "placed" && new Date(o.expires_at).getTime() <= Date.now() ? "Expired" : orderLabel(o)}</span></div>
          {o.tracking_no && <p className="text-xs mt-1" style={{ color: ui.C.muted }}>Tracking: {o.courier ? `${o.courier} ` : ""}{o.tracking_no}</p>}
        </a>
      ))}
    </div>
  );
}

function ShopperPages({ page, slug, C, serif, shopper, open, checkoutOpen, bag, onReloadBag, onRemoveFromBag, orderId, recovery, note, onSession, onSignOut, onRecoveryDone, go }) {
  const ui = { C, serif, input: { border: `1px solid ${C.line}`, background: C.panel, color: C.ink, fontSize: 16, padding: "12px 14px", width: "100%", outline: "none", borderRadius: 0 } };
  if (!open && !shopper && !(page === "reset" && recovery)) {
    return (
      <ShopperShell ui={ui} page="closed" title="Accounts are opening soon" intro="Customer accounts aren't available just yet. You can still browse the shop and message us on WhatsApp to order.">
        <ShopperLink ui={ui} hash="" go={go}>Back to the shop</ShopperLink>
      </ShopperShell>
    );
  }
  if (page === "bag" || page === "checkout" || page === "order") {
    if (!checkoutOpen) {
      return (
        <ShopperShell ui={ui} page="closed" title="Checkout isn't open yet" intro="You can still browse the shop and message us on WhatsApp to order.">
          <ShopperLink ui={ui} hash="" go={go}>Back to the shop</ShopperLink>
        </ShopperShell>
      );
    }
    if (!shopper) return <ShopperLogin ui={ui} go={go} onSession={onSession} note="Log in to continue." returnTo={window.location.hash} />;
    if (page === "bag") return <ShopperBag ui={ui} shopper={shopper} bag={bag} onReload={onReloadBag} onRemove={onRemoveFromBag} onSession={onSession} go={go} />;
    if (page === "checkout") return <ShopperCheckout ui={ui} shopper={shopper} bag={bag} onReload={onReloadBag} onSession={onSession} go={go} />;
    return <ShopperOrder ui={ui} shopper={shopper} orderId={orderId} onSession={onSession} go={go} />;
  }
  const account = <ShopperAccount ui={ui} shopper={shopper} onSession={onSession} onSignOut={onSignOut} note={note} checkoutOpen={checkoutOpen} go={go} />;
  if (page === "account") return shopper ? account : <ShopperLogin ui={ui} go={go} onSession={onSession} note="Log in to see your account." returnTo="#account" />;
  if (page === "signup") return shopper ? account : <ShopperSignup ui={ui} slug={slug} go={go} onSession={onSession} />;
  if (page === "forgot") return <ShopperForgot ui={ui} slug={slug} go={go} />;
  if (page === "reset") return <ShopperReset ui={ui} recovery={recovery} go={go} onSession={onSession} onDone={onRecoveryDone} />;
  return shopper ? account : <ShopperLogin ui={ui} go={go} onSession={onSession} note={note} />;
}

function PublicStorefront({ slug }) {
  const [state, setState] = useState("loading"); // loading | ready | not-found | error
  const [store, setStore] = useState({ whatsapp: "", banner: "", intro: "", logo: "", sale: null, banners: [], announcement: "", faq: [] });
  const [storeItems, setStoreItems] = useState([]);
  const [lightbox, setLightbox] = useState(null);
  const [search, setSearch] = useState("");
  const [categoryFilter, setCategoryFilter] = useState("all");
  const [sort, setSort] = useState("newest");
  const [fSizes, setFSizes] = useState([]);
  const [fBrand, setFBrand] = useState("");
  const [fPrice, setFPrice] = useState("");
  const [showFilters, setShowFilters] = useState(false);
  const touchX = useRef(null);
  const [photoIdx, setPhotoIdx] = useState(0);
  const [copied, setCopied] = useState(false);
  const gridRef = useRef(null);
  const searchRef = useRef(null);
  const trackRef = useRef({ userId: null, enabled: false });
  const track = (type, item) => { if (trackRef.current.enabled) trackStorefrontEvent(trackRef.current.userId, type, item); };
  const [route, setRoute] = useState(parseStoreHash);
  const itemId = route.item, page = route.page, collectionId = route.collection;
  const prevHashRef = useRef("");
  const [faqOpen, setFaqOpen] = useState(-1);
  const [shopper, setShopper] = useState(loadShopperSession);
  const [accountsOpen, setAccountsOpen] = useState(false);
  const [recovery, setRecovery] = useState(null);
  const [authNote, setAuthNote] = useState("");
  const [checkoutOpen, setCheckoutOpen] = useState(false);
  const [bag, setBag] = useState(null); // null = not loaded yet
  const [bagMsg, setBagMsg] = useState("");
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
  const onShopperSession = (next) => { saveShopperSession(next); setShopper(next); };
  const onShopperSignOut = async () => {
    const s = shopper;
    saveShopperSession(null); setShopper(null); setRecovery(null); setAuthNote("");
    go("");
    if (s) await signOutRequest(s);
  };
  const rpc = (name, args) => apiRequest(`/rpc/${name}`, { method: "POST", body: JSON.stringify(args || {}) }, shopper, onShopperSession);
  const loadBag = async () => {
    if (!shopper) { setBag(null); return null; }
    try { const items = await rpc("cart_get"); setBag(items || []); return items; } catch (e) { return null; }
  };
  const removeFromBag = async (id) => { try { setBag(await rpc("cart_remove", { p_item: id })); } catch (e) {} };
  const addToBag = async (item) => {
    setBagMsg("");
    if (!shopper) { try { sessionStorage.setItem("ravuno_pending_item", item.id); } catch (e) {} go("#login"); return; }
    try { setBag(await rpc("cart_add", { p_item: item.id })); } catch (e) { setBagMsg(e.message || "We couldn't add that piece."); }
  };
  const shopperId = shopper && shopper.user ? shopper.user.id : null;
  useEffect(() => { if (shopperId && checkoutOpen) loadBag(); else if (!shopperId) setBag(null); }, [shopperId, checkoutOpen]);
  // a piece they tapped "Log in to add to bag" on goes into the bag as soon as they are logged in
  useEffect(() => {
    if (!shopperId || !checkoutOpen) return;
    let pending = null;
    try { pending = sessionStorage.getItem("ravuno_pending_item"); } catch (e) {}
    if (!pending) return;
    try { sessionStorage.removeItem("ravuno_pending_item"); } catch (e) {}
    (async () => {
      try { setBag(await rpc("cart_add", { p_item: pending })); go("#bag"); }
      catch (e) { setBagMsg(e.message || "We couldn't add that piece."); go(`#item=${pending}`); }
    })();
  }, [shopperId, checkoutOpen]);
  // are accounts open? and did the visitor arrive from an email link (confirm / reset password / expired)?
  useEffect(() => {
    shopperAccountsOpen().then(setAccountsOpen);
    checkoutIsOpen().then(setCheckoutOpen);
    const h = parseAuthHash(window.location.hash);
    const clean = (hash) => { window.history.replaceState(null, "", window.location.pathname + window.location.search + hash); setRoute(parseStoreHash()); };
    if (h.error) { setAuthNote("That link has expired or isn't valid. Please request a new one."); clean("#login"); return; }
    if (h.type === "recovery" && h.accessToken) { setRecovery({ accessToken: h.accessToken, refreshToken: h.refreshToken }); clean("#reset"); return; }
    if (h.accessToken && ["signup", "magiclink", "email_change", "invite"].includes(h.type)) {
      (async () => {
        try {
          const user = await authUserRequest(h.accessToken);
          const sess = { access_token: h.accessToken, refresh_token: h.refreshToken, user };
          saveShopperSession(sess); setShopper(sess); setAuthNote("Your email is confirmed. Welcome!");
        } catch (e) { setAuthNote("That link has expired. Please log in."); clean("#login"); return; }
        clean("#account");
      })();
    }
  }, []);
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
  const categories = shopCategoriesPresent(storeItems);
  const q = search.trim().toLowerCase();
  const isFiltering = !!q || categoryFilter !== "all";
  const sizeOptions = sortSizes(Array.from(new Set(available.map(getDisplaySize).filter(Boolean))));
  const brandOptions = Array.from(new Map(available.filter((i) => String(i.brand || "").trim()).map((i) => [String(i.brand).trim().toLowerCase(), String(i.brand).trim()])).values()).sort((a, b) => a.localeCompare(b));
  const activeFilters = fSizes.length + (fBrand ? 1 : 0) + (fPrice ? 1 : 0);
  const clearFilters = () => { setFSizes([]); setFBrand(""); setFPrice(""); };
  let shown = storeItems.filter((i) => {
    if (categoryFilter !== "all" && shopCategoryOf(i) !== categoryFilter) return false;
    if (!passesShopFilters(i, { sizes: fSizes, brand: fBrand, price: fPrice }, saleFor(i, saleCtx).price, getDisplaySize(i))) return false;
    if (!q) return true;
    return [productTitle(i), i.brand, i.colour, i.itemType, shopCategoryLabel(shopCategoryOf(i))].some((v) => (v || "").toLowerCase().includes(q));
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
    ["Category", categoryLine(selected)],
  ].filter(([, v]) => v) : [];
  const measureLines = selected
    ? (selected.category === "apparel" ? apparelMeasureParts(selected.attributes || {}, getApparelSubcategory(selected.itemType)) : getFilledAttrLines(selected))
    : [];
  const related = selected ? [...available.filter((i) => i.id !== selected.id && shopCategoryOf(i) === shopCategoryOf(selected)), ...available.filter((i) => i.id !== selected.id && shopCategoryOf(i) !== shopCategoryOf(selected) && i.category === selected.category)].slice(0, 4) : [];
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
  const selReserved = !!selected && selected.status !== "sold" && isReserved(selected);
  const selInBag = !!selected && !!bag && bag.some((b) => b.item_id === selected.id);
  const C = { bg: "#0C0C0C", ink: "#F2F0EA", muted: "#9C988F", line: "#2B2B2B", tile: "#171717", panel: "#141414", red: "#E8694F" };
  const serif = { fontFamily: "'Barlow Condensed', 'Inter', sans-serif", fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.02em" };

  const renderCard = (item) => {
    const sold = item.status === "sold";
    const reserved = !sold && isReserved(item);
    const sale = sold ? null : saleFor(item, saleCtx);
    const onSale = !!sale && sale.off > 0;
    return (
      <a key={item.id} href={`#item=${item.id}`} onClick={(e) => { e.preventDefault(); openItem(item.id); }}
        className="block w-full" style={{ color: "inherit", textDecoration: "none" }}>
        <div className="relative overflow-hidden w-full" style={{ background: C.tile, aspectRatio: "1 / 1", border: `1px solid ${C.line}`, borderRadius: 6 }}>
          {item.photos && item.photos[0] ? (
            <img src={item.photos[0]} alt={productTitle(item)} decoding="async" className="absolute inset-0 w-full h-full object-cover" style={{ objectPosition: "center", opacity: sold || reserved ? 0.55 : 1 }} />
          ) : (
            <div className="absolute inset-0 flex items-center justify-center"><ImageIcon size={22} style={{ color: C.muted }} /></div>
          )}
          {sold && <span className="absolute top-2 left-2 text-[10px] font-semibold tracking-[0.12em] uppercase px-2 py-1" style={{ background: C.ink, color: C.bg, borderRadius: 3 }}>Sold</span>}
          {reserved && <span data-reserved-badge="" className="absolute top-2 left-2 text-[10px] font-semibold tracking-[0.12em] uppercase px-2 py-1" style={{ background: C.muted, color: C.bg, borderRadius: 3 }}>Reserved</span>}
          {onSale && <span className="absolute top-2 right-2 text-[10px] font-semibold tracking-[0.1em] uppercase px-2 py-1" style={{ background: "#B5472F", color: "#FFFFFF", borderRadius: 3 }}>{onSale ? sale.badge : ""}</span>}
        </div>
        <div className="pt-3">
          <p className="text-[10px] tracking-[0.16em] uppercase mb-1 truncate" style={{ color: C.muted, minHeight: "1.5em" }}>{item.brand || "\u00A0"}</p>
          <p className="text-sm font-medium mb-1" style={{ color: C.ink, lineHeight: 1.35, minHeight: "2.7em", display: "-webkit-box", WebkitLineClamp: 2, WebkitBoxOrient: "vertical", overflow: "hidden" }}>{productTitle(item)}</p>
          {onSale ? (
            <p className="text-sm font-semibold"><span style={{ color: C.red }}>{fmtRM(sale.price)}</span> <span className="ml-1 font-normal" style={{ color: C.muted, textDecoration: "line-through" }}>RM{item.price}</span></p>
          ) : (
            <p className="text-sm font-semibold" style={{ color: sold ? C.muted : C.ink, textDecoration: sold ? "line-through" : "none" }}>RM{item.price || "?"}</p>
          )}
        </div>
      </a>
    );
  };

  const renderCollectionTile = (c) => (
    <a key={c.id} href={`#collection=${c.id}`} onClick={(e) => { e.preventDefault(); go(`#collection=${c.id}`); }} aria-label={c.title}
      className="relative overflow-hidden block w-full" style={{ display: "block", aspectRatio: "3 / 2", background: C.tile, color: "inherit", textDecoration: "none", border: `1px solid ${C.line}`, borderRadius: 6 }}>
      {collectionCover(c) && <img src={collectionCover(c)} alt={c.title} decoding="async" className="absolute inset-0 w-full h-full object-cover" />}
    </a>
  );

  const sectionTitle = (text, aside, action) => (
    <div className="flex items-end justify-between mb-5 gap-4">
      <h2 className="text-3xl leading-none" style={serif}>{text}</h2>
      {action ? (
        <button onClick={action.onClick} className="inline-flex items-center gap-1.5 text-[11px] font-semibold tracking-[0.12em] uppercase whitespace-nowrap">{action.label} <ArrowRight size={13} /></button>
      ) : aside ? <span className="text-[11px] tracking-[0.14em] uppercase" style={{ color: C.muted }}>{aside}</span> : null}
    </div>
  );
  const scrollToGrid = () => setTimeout(() => gridRef.current && gridRef.current.scrollIntoView({ behavior: "smooth", block: "start" }), 50);
  const pickCategory = (cat) => { setCategoryFilter(cat); scrollToGrid(); };

  return (
    <div className="min-h-screen w-full" style={{ background: C.bg, color: C.ink, fontFamily: "'Inter', sans-serif" }}>
      <style>{`@import url('https://fonts.googleapis.com/css2?family=Barlow+Condensed:wght@600;700&family=Inter:wght@400;500;600;700&display=swap');
        .rv-bag::after{content:attr(data-count);position:absolute;top:-8px;right:-8px;min-width:16px;height:16px;padding:0 4px;border-radius:9999px;background:#F2F0EA;color:#0C0C0C;font-size:9px;font-weight:700;display:flex;align-items:center;justify-content:center}
        @media (min-width:640px){.rv-bag::after{display:none}}`}</style>

      {store.announcement && <Marquee text={store.announcement} background="#000000" color="#F2F0EA" />}

      <header className="sticky top-0 z-30 border-b" style={{ borderColor: C.line, background: C.bg }}>
        <div className="max-w-7xl mx-auto px-5 flex flex-wrap items-center">
          <a href={window.location.pathname} onClick={(e) => { e.preventDefault(); go(""); }} className="order-1 block py-3" style={{ color: "inherit", textDecoration: "none" }}>
            {store.logo ? <img src={store.logo} alt={slug} className="h-9 object-contain" /> : <span className="text-2xl" style={serif}>@{slug}</span>}
          </a>
          <div className="order-2 ml-auto md:order-3 flex items-center gap-5">
            <button data-nav-search="" aria-label="Search" onClick={() => { if (itemId || page !== "home") go(""); setTimeout(() => { scrollToGrid(); if (searchRef.current) searchRef.current.focus(); }, 120); }}><Search size={18} /></button>
            {(accountsOpen || shopper) && (
              <a data-account-link="" href={shopper ? "#account" : "#login"} aria-label={shopper ? "My Account" : "Login / Sign up"} onClick={(e) => { e.preventDefault(); go(shopper ? "#account" : "#login"); }}
                className="inline-flex items-center gap-2 text-[11px] tracking-[0.12em] uppercase whitespace-nowrap" style={{ color: "inherit", textDecoration: "none" }}>
                <User size={18} /><span className="hidden sm:inline">{shopper ? "My Account" : "Login / Sign up"}</span>
              </a>
            )}
            {checkoutOpen && shopper && (
              <a data-bag-link="" href="#bag" aria-label={`Bag, ${bag ? bag.length : 0} items`} onClick={(e) => { e.preventDefault(); go("#bag"); }}
                data-count={bag ? bag.length : 0} className="rv-bag relative inline-flex items-center gap-2 text-[11px] tracking-[0.12em] uppercase whitespace-nowrap" style={{ color: "inherit", textDecoration: "none" }}>
                <ShoppingBag size={18} />
                <span className="hidden sm:inline">Bag ({bag ? bag.length : 0})</span>
              </a>
            )}
          </div>
          <nav className="order-3 w-full md:order-2 md:w-auto md:flex-1 md:justify-center flex items-center gap-7 overflow-x-auto pb-3 md:py-0 text-[12px] font-medium tracking-[0.06em] whitespace-nowrap">
            {navItems.map((n) => (
              <a key={n.key} href={n.hash || window.location.pathname} onClick={(e) => { e.preventDefault(); go(n.hash); }}
                style={{ color: n.hot ? C.red : "inherit", textDecoration: "none", paddingBottom: 3, borderBottom: `1px solid ${activeNav === n.key ? (n.hot ? C.red : C.ink) : "transparent"}` }}>{n.label}</a>
            ))}
            {waDigits && <a href={waLink(null)} onClick={() => track("whatsapp", null)} target="_blank" rel="noreferrer" style={{ color: "inherit", textDecoration: "none", paddingBottom: 3, borderBottom: "1px solid transparent" }}>Contact</a>}
          </nav>
        </div>
      </header>

      {state === "loading" && <p className="text-center text-sm py-32" style={{ color: C.muted }}>Loading…</p>}
      {state === "not-found" && <p className="text-center text-sm py-32" style={{ color: C.muted }}>This storefront doesn't exist.</p>}
      {state === "error" && <p className="text-center text-sm py-32" style={{ color: C.muted }}>Something went wrong loading this page.</p>}

      {state === "ready" && !itemId && ["login", "signup", "forgot", "reset", "account", "bag", "checkout", "order"].includes(page) && (
        <ShopperPages page={page} slug={slug} C={C} serif={serif} shopper={shopper} open={accountsOpen} recovery={recovery} note={authNote}
          checkoutOpen={checkoutOpen} bag={bag} onReloadBag={loadBag} onRemoveFromBag={removeFromBag} orderId={route.order}
          onSession={onShopperSession} onSignOut={onShopperSignOut} onRecoveryDone={() => setRecovery(null)} go={go} />
      )}

      {state === "ready" && !itemId && page === "home" && (
        <main>
          <BannerCarousel slides={slides} target={slideTarget} />
          {slides.length === 0 && (
            <section className="px-5 py-16 md:py-24 border-b" style={{ borderColor: C.line, background: "linear-gradient(135deg, #1b1b1b, #0c0c0c)" }}>
              <div className="max-w-7xl mx-auto">
                <h1 className="text-5xl md:text-7xl leading-none mb-4" style={serif}>{store.logo ? slug : `@${slug}`}</h1>
                <p className="text-base mb-6 max-w-md" style={{ color: C.muted }}>{store.intro || "Hand-picked second-hand pieces. Once they're gone, they're gone."}</p>
                <button onClick={scrollToGrid} className="inline-flex items-center gap-2 text-[12px] font-semibold tracking-[0.12em] uppercase px-6 py-3.5" style={{ background: C.ink, color: C.bg, borderRadius: 3 }}>Shop now <ArrowRight size={14} /></button>
              </div>
            </section>
          )}

          {!isFiltering && categories.length > 1 && (
            <section data-category-strip="" className="max-w-7xl mx-auto px-5 pt-8 pb-12">
              <div className="grid grid-flow-col gap-3 overflow-x-auto pb-1" style={{ gridAutoColumns: "minmax(138px, 1fr)" }}>
                {categories.map((cat) => {
                  const sample = storeItems.find((i) => shopCategoryOf(i) === cat && i.status !== "sold" && i.photos && i.photos[0]) || storeItems.find((i) => shopCategoryOf(i) === cat && i.photos && i.photos[0]);
                  return (
                    <button key={cat} data-shop-category={cat} onClick={() => pickCategory(cat)} aria-label={`Shop ${shopCategoryLabel(cat)}`}
                      className="relative overflow-hidden text-left" style={{ aspectRatio: "1 / 1.05", background: C.tile, border: `1px solid ${C.line}`, borderRadius: 6 }}>
                      {sample && <img src={sample.photos[0]} alt="" decoding="async" className="absolute inset-0 w-full h-full object-cover" />}
                      <div className="absolute inset-x-0 bottom-0 flex items-center justify-between gap-2 px-3 py-2.5" style={{ background: "rgba(12,12,12,0.8)" }}>
                        <span className="text-[11px] font-semibold tracking-[0.1em] uppercase truncate">{shopCategoryLabel(cat)}</span>
                        <ArrowRight size={12} className="shrink-0" />
                      </div>
                    </button>
                  );
                })}
              </div>
            </section>
          )}

          {!isFiltering && (
            <section className="max-w-3xl mx-auto px-6 pb-14 text-center">
              <p className="text-lg sm:text-xl leading-relaxed font-light mb-7">
                {store.intro || "Hand-picked second-hand pieces. Once they're gone, they're gone."}
              </p>
              {waDigits && (
                <a href={waLink(null)} onClick={() => track("whatsapp", null)} target="_blank" rel="noreferrer" className="inline-block text-[11px] font-semibold tracking-[0.16em] uppercase px-7 py-3.5" style={{ border: `1px solid ${C.ink}`, borderRadius: 3 }}>
                  Enquire on WhatsApp
                </a>
              )}
            </section>
          )}

          {!isFiltering && featured.length > 0 && (
            <section className="max-w-7xl mx-auto px-5 pb-14">
              {sectionTitle("Featured drops", null, { label: "View all products", onClick: () => pickCategory("all") })}
              <div className="grid grid-cols-2 md:grid-cols-4 gap-x-4 gap-y-8">{featured.map(renderCard)}</div>
            </section>
          )}

          {!isFiltering && newArrivals.length > 0 && (
            <section className="max-w-7xl mx-auto px-5 pb-14">
              {sectionTitle("New arrivals")}
              <div className="flex gap-4 overflow-x-auto pb-2">
                {newArrivals.map((item) => <div key={item.id} className="w-40 sm:w-56 shrink-0">{renderCard(item)}</div>)}
              </div>
            </section>
          )}

          {!isFiltering && collections.length > 0 && (
            <section className="max-w-7xl mx-auto px-5 pb-14">
              {sectionTitle("Collections", null, { label: "All collections", onClick: () => go("#collections") })}
              <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-3">{collections.slice(0, 6).map(renderCollectionTile)}</div>
            </section>
          )}

          {!isFiltering && (
            <section data-promo-tiles="" className="max-w-7xl mx-auto px-5 pb-14 grid md:grid-cols-2 gap-4">
              {[
                accountsOpen && !shopper
                  ? { title: "Join the community", text: "Create an account to check out faster and track your orders.", cta: "Sign up", on: () => go("#signup"), pic: available[3] }
                  : saleItems.length > 0
                    ? { title: "On sale now", text: `${saleItems.length} piece${saleItems.length === 1 ? "" : "s"} marked down. Once they're gone, they're gone.`, cta: "Shop the sale", on: () => go("#sale"), pic: saleItems[0] }
                    : { title: "Browse everything", text: "Every piece is one of a kind.", cta: "Shop now", on: () => pickCategory("all"), pic: available[3] },
                waDigits
                  ? { title: "Looking for something?", text: "Tell us what you're after and we'll help you find it.", cta: "Message us", href: waLink(null), pic: available[5] }
                  : collections.length > 0 ? { title: "Collections", text: "Curated sets of pieces that go together.", cta: "View collections", on: () => go("#collections"), pic: available[5] } : null,
              ].filter(Boolean).map((t, i) => (
                <div key={i} className="relative overflow-hidden flex flex-col justify-end p-6" style={{ minHeight: 170, background: C.tile, border: `1px solid ${C.line}`, borderRadius: 6 }}>
                  {t.pic && t.pic.photos && t.pic.photos[0] && <img src={t.pic.photos[0]} alt="" decoding="async" className="absolute inset-0 w-full h-full object-cover" style={{ opacity: 0.35 }} />}
                  <div className="absolute inset-0" style={{ background: "linear-gradient(to top, rgba(12,12,12,0.85), rgba(12,12,12,0.25))" }} />
                  <div className="relative">
                    <h3 className="text-3xl leading-none mb-1.5" style={serif}>{t.title}</h3>
                    <p className="text-sm mb-4" style={{ color: "rgba(242,240,234,0.8)" }}>{t.text}</p>
                    {t.href
                      ? <a href={t.href} onClick={() => track("whatsapp", null)} target="_blank" rel="noreferrer" className="inline-flex items-center gap-2 text-[11px] font-semibold tracking-[0.14em] uppercase px-5 py-3" style={{ background: C.ink, color: C.bg, borderRadius: 3, textDecoration: "none" }}>{t.cta} <ArrowRight size={13} /></a>
                      : <button onClick={t.on} className="inline-flex items-center gap-2 text-[11px] font-semibold tracking-[0.14em] uppercase px-5 py-3" style={{ background: C.ink, color: C.bg, borderRadius: 3 }}>{t.cta} <ArrowRight size={13} /></button>}
                  </div>
                </div>
              ))}
            </section>
          )}

          
          <section ref={gridRef} className="max-w-6xl mx-auto px-5 pb-20" style={{ scrollMarginTop: "80px" }}>
            {sectionTitle("All pieces", `${shown.length} item${shown.length === 1 ? "" : "s"}`)}
            <div className="relative mb-4">
              <Search size={15} className="absolute left-3 top-1/2 -translate-y-1/2" style={{ color: C.muted }} />
              <input ref={searchRef} value={search} onChange={(e) => setSearch(e.target.value)} placeholder="Search by name, brand or colour…"
                className="w-full pl-9 pr-3 py-3 text-sm outline-none" style={{ border: `1px solid ${C.line}`, background: C.panel, color: C.ink }} />
            </div>
            <div className="flex items-center gap-2 overflow-x-auto pb-2 mb-4">
              {["all", ...categories].map((cat) => {
                const on = categoryFilter === cat;
                return (
                  <button key={cat} onClick={() => setCategoryFilter(cat)} className="shrink-0 text-[11px] font-semibold tracking-[0.1em] uppercase px-3.5 py-2"
                    style={{ border: `1px solid ${on ? C.ink : C.line}`, background: on ? C.ink : "transparent", color: on ? C.bg : C.ink, borderRadius: 3 }}>
                    {cat === "all" ? "All" : shopCategoryLabel(cat)}
                  </button>
                );
              })}
            </div>
            <div className="flex items-center justify-between gap-3 mb-4 flex-wrap">
              <button data-filters-toggle="" onClick={() => setShowFilters((v) => !v)} aria-expanded={showFilters} className="inline-flex items-center gap-2 text-[11px] font-semibold tracking-[0.12em] uppercase px-3.5 py-2" style={{ border: `1px solid ${activeFilters ? C.ink : C.line}`, borderRadius: 3 }}>
                <SlidersHorizontal size={14} /> Filters{activeFilters ? ` (${activeFilters})` : ""}
              </button>
              <select data-sort="" value={sort} onChange={(e) => setSort(e.target.value)} aria-label="Sort" className="text-[11px] tracking-[0.12em] uppercase px-3 py-2 outline-none" style={{ border: `1px solid ${C.line}`, background: C.panel, color: C.ink, borderRadius: 3 }}>
                <option value="newest">Newest</option>
                <option value="price-asc">Price: low to high</option>
                <option value="price-desc">Price: high to low</option>
              </select>
            </div>
            {showFilters && (
              <div data-filters-panel="" className="mb-6 p-4 space-y-5" style={{ border: `1px solid ${C.line}`, background: C.panel, borderRadius: 6 }}>
                {sizeOptions.length > 0 && (
                  <div>
                    <p className="text-[11px] tracking-[0.16em] uppercase mb-2" style={{ color: C.muted }}>Size</p>
                    <div className="flex flex-wrap gap-2">
                      {sizeOptions.map((sz) => {
                        const on = fSizes.includes(sz);
                        return <button key={sz} data-filter-size={sz} aria-pressed={on} onClick={() => setFSizes((cur) => (cur.includes(sz) ? cur.filter((x) => x !== sz) : [...cur, sz]))}
                          className="text-xs font-semibold px-3.5 py-2 min-w-[44px]" style={{ border: `1px solid ${on ? C.ink : C.line}`, background: on ? C.ink : "transparent", color: on ? C.bg : C.ink, borderRadius: 3 }}>{sz}</button>;
                      })}
                    </div>
                  </div>
                )}
                {brandOptions.length > 0 && (
                  <div>
                    <p className="text-[11px] tracking-[0.16em] uppercase mb-2" style={{ color: C.muted }}>Brand</p>
                    <select data-filter-brand="" value={fBrand} onChange={(e) => setFBrand(e.target.value)} aria-label="Brand" className="w-full text-sm px-3 py-3 outline-none" style={{ border: `1px solid ${C.line}`, background: C.bg, color: C.ink, borderRadius: 3 }}>
                      <option value="">All brands</option>
                      {brandOptions.map((b) => <option key={b} value={b}>{b}</option>)}
                    </select>
                  </div>
                )}
                <div>
                  <p className="text-[11px] tracking-[0.16em] uppercase mb-2" style={{ color: C.muted }}>Price</p>
                  <div className="flex flex-wrap gap-2">
                    {PRICE_BANDS.map((b) => {
                      const on = fPrice === b.id;
                      return <button key={b.id || "any"} data-filter-price={b.id || "any"} aria-pressed={on} onClick={() => setFPrice(b.id)}
                        className="text-xs font-semibold px-3.5 py-2" style={{ border: `1px solid ${on ? C.ink : C.line}`, background: on ? C.ink : "transparent", color: on ? C.bg : C.ink, borderRadius: 3 }}>{b.label}</button>;
                    })}
                  </div>
                </div>
                {activeFilters > 0 && <button data-filters-clear="" onClick={clearFilters} className="text-[11px] font-semibold tracking-[0.12em] uppercase underline">Clear filters</button>}
              </div>
            )}

            {storeItems.length === 0 && <p className="text-center text-sm py-16" style={{ color: C.muted }}>New pieces arriving soon.</p>}
            {storeItems.length > 0 && shown.length === 0 && (
              <div data-no-match="" className="text-center py-16">
                <p className="text-sm mb-4" style={{ color: C.muted }}>No pieces match your search.</p>
                {(activeFilters > 0 || q || categoryFilter !== "all") && <button onClick={() => { clearFilters(); setSearch(""); setCategoryFilter("all"); }} className="text-[11px] font-semibold tracking-[0.14em] uppercase px-5 py-3" style={{ border: `1px solid ${C.ink}`, borderRadius: 3 }}>Clear all</button>}
              </div>
            )}
            {shown.length > 0 && <div className="grid grid-cols-2 md:grid-cols-3 lg:grid-cols-4 gap-x-4 gap-y-8">{shown.map(renderCard)}</div>}
          </section>
        </main>
      )}

      {state === "ready" && itemId && (
        <main className="max-w-5xl mx-auto px-5 py-8 pb-28 md:pb-8">
          <button onClick={closeItem} className="text-[11px] tracking-[0.16em] uppercase mb-6">← Back</button>
          {!selected ? (
            <p className="text-center text-sm py-24" style={{ color: C.muted }}>This piece is no longer available.</p>
          ) : (
            <>
              <div className="grid md:grid-cols-2 gap-8 md:gap-12">
                <div>
                  <div data-gallery="" className="relative w-full overflow-hidden flex items-center justify-center" style={{ background: C.tile, aspectRatio: "1 / 1", border: `1px solid ${C.line}`, borderRadius: 6, touchAction: "pan-y" }}
                    onTouchStart={(e) => { touchX.current = e.touches[0].clientX; }}
                    onTouchEnd={(e) => { if (touchX.current === null || selectedPhotos.length < 2) return; const dx = e.changedTouches[0].clientX - touchX.current; touchX.current = null; if (Math.abs(dx) > 40) setPhotoIdx((i) => (dx < 0 ? (i + 1) % selectedPhotos.length : (i - 1 + selectedPhotos.length) % selectedPhotos.length)); }}>
                    {selOnSale && <span className="absolute top-3 right-3 z-10 text-[11px] tracking-[0.14em] uppercase px-3 py-1.5" style={{ background: "#B5472F", color: "#FFFFFF", borderRadius: 3 }}>{selSale.badge} sale</span>}
                    {selectedPhotos.length > 0 ? (
                      <img src={selectedPhotos[photoIdx] || selectedPhotos[0]} alt={productTitle(selected)} onClick={() => setLightbox({ photos: selectedPhotos, index: photoIdx })}
                        className="w-full h-full object-contain cursor-pointer" />
                    ) : (
                      <ImageIcon size={28} style={{ color: C.muted }} />
                    )}
                    {selectedPhotos.length > 1 && (
                      <>
                        <button data-photo-prev="" aria-label="Previous photo" onClick={() => setPhotoIdx((i) => (i - 1 + selectedPhotos.length) % selectedPhotos.length)} className="absolute left-2 top-1/2 -translate-y-1/2 w-9 h-9 rounded-full flex items-center justify-center" style={{ background: "rgba(12,12,12,0.6)", color: "#FFFFFF" }}><ChevronLeft size={18} /></button>
                        <button data-photo-next="" aria-label="Next photo" onClick={() => setPhotoIdx((i) => (i + 1) % selectedPhotos.length)} className="absolute right-2 top-1/2 -translate-y-1/2 w-9 h-9 rounded-full flex items-center justify-center" style={{ background: "rgba(12,12,12,0.6)", color: "#FFFFFF" }}><ChevronRight size={18} /></button>
                        <span data-photo-count="" className="absolute bottom-2 right-2 text-[11px] font-semibold tracking-[0.08em] px-2 py-1" style={{ background: "rgba(12,12,12,0.65)", color: "#FFFFFF", borderRadius: 3 }}>{(photoIdx % selectedPhotos.length) + 1} / {selectedPhotos.length}</span>
                      </>
                    )}
                  </div>
                  {selectedPhotos.length > 1 && (
                    <div className="flex gap-2 mt-3 overflow-x-auto pb-1">
                      {selectedPhotos.map((p, idx) => (
                        <button key={idx} data-thumb={idx} aria-label={`Photo ${idx + 1}`} onClick={() => setPhotoIdx(idx)} className="shrink-0 overflow-hidden" style={{ width: 72, height: 72, border: `1px solid ${idx === photoIdx ? C.ink : C.line}`, opacity: idx === photoIdx ? 1 : 0.65, borderRadius: 4 }}>
                          <img src={p} alt="" className="w-full h-full object-cover" />
                        </button>
                      ))}
                    </div>
                  )}
                </div>

                <div>
                  {selected.brand && <p className="text-[11px] tracking-[0.18em] uppercase mb-2" style={{ color: C.muted }}>{selected.brand}</p>}
                  <h1 className="text-3xl md:text-4xl leading-tight mb-3" style={serif}>{productTitle(selected)}</h1>
                  {getDisplaySize(selected) && <span data-size-chip="" className="inline-block text-[11px] font-semibold tracking-[0.14em] uppercase px-3 py-1.5 mb-4" style={{ border: `1px solid ${C.ink}`, borderRadius: 3 }}>Size {getDisplaySize(selected)}</span>}
                  <p className="text-xl mb-6" style={{ textDecoration: selected.status === "sold" ? "line-through" : "none", color: selected.status === "sold" ? C.muted : C.ink }}>
                    {selOnSale ? (
                      <><span style={{ color: C.red }}>{fmtRM(selSale.price)}</span> <span className="text-base ml-1" style={{ color: C.muted, textDecoration: "line-through" }}>RM{selected.price}</span></>
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
                      <table data-size-table="" className="w-full text-sm" style={{ borderCollapse: "collapse" }}>
                        <tbody>
                          {measureLines.map((l, idx) => { const at = l.indexOf(": "); const k = at > 0 ? l.slice(0, at) : l; const v = at > 0 ? l.slice(at + 2) : ""; return (
                            <tr key={idx} style={{ borderBottom: `1px solid ${C.line}` }}><td className="py-2" style={{ color: C.muted }}>{k}</td><td className="py-2 text-right">{v}</td></tr>
                          ); })}
                        </tbody>
                      </table>
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
                  ) : selReserved ? (
                    <div data-reserved-note="" className="py-4 px-3 text-center text-[11px] tracking-[0.18em] uppercase leading-relaxed" style={{ border: `1px solid ${C.line}`, color: C.muted }}>Reserved. Another customer is holding this piece. It may come back if they don't pay.</div>
                  ) : (
                    <>
                      {checkoutOpen && (
                        <>
                          <button data-add-to-bag="" onClick={() => (selInBag ? go("#bag") : addToBag(selected))} className="block w-full text-center text-[11px] tracking-[0.18em] uppercase py-4 mb-3" style={{ background: C.ink, color: C.bg }}>
                            {selInBag ? "In your bag · View bag" : shopper ? "Add to bag" : "Log in to add to bag"}
                          </button>
                          {bagMsg && <p role="alert" data-bag-error="" className="text-xs mb-3 text-center" style={{ color: C.red }}>{bagMsg}</p>}
                        </>
                      )}
                      {waDigits ? (
                        <a href={waLink(selected)} onClick={() => track("whatsapp", selected)} target="_blank" rel="noreferrer" className="block text-center text-[11px] tracking-[0.18em] uppercase py-4" style={checkoutOpen ? { border: `1px solid ${C.ink}`, color: C.ink } : { background: C.ink, color: C.bg }}>
                          Enquire on WhatsApp
                        </a>
                      ) : (
                        <p className="text-xs text-center" style={{ color: C.muted }}>Contact details unavailable</p>
                      )}
                    </>
                  )}
                  <button onClick={copyLink} className="block w-full text-center text-[11px] tracking-[0.16em] uppercase mt-4" style={{ color: C.muted }}>
                    {copied ? "Link copied" : "Copy link to this piece"}
                  </button>
                </div>
              </div>

              {selected.status !== "sold" && !selReserved && (checkoutOpen || waDigits) && (
                <div data-sticky-buy="" className="md:hidden fixed bottom-0 inset-x-0 z-40 flex items-center gap-4 px-5 py-3 border-t" style={{ background: C.bg, borderColor: C.line }}>
                  <div className="shrink-0">
                    <p className="text-[10px] tracking-[0.14em] uppercase" style={{ color: C.muted }}>Price</p>
                    <p className="text-base font-semibold" style={{ color: selOnSale ? C.red : C.ink }}>{selOnSale ? fmtRM(selSale.price) : `RM${selected.price || "?"}`}</p>
                  </div>
                  {checkoutOpen ? (
                    <button data-sticky-add="" onClick={() => (selInBag ? go("#bag") : addToBag(selected))} className="flex-1 text-center text-[11px] font-semibold tracking-[0.16em] uppercase py-3.5" style={{ background: C.ink, color: C.bg, borderRadius: 3 }}>
                      {selInBag ? "View bag" : shopper ? "Add to bag" : "Log in to add"}
                    </button>
                  ) : (
                    <a href={waLink(selected)} onClick={() => track("whatsapp", selected)} target="_blank" rel="noreferrer" className="flex-1 text-center text-[11px] font-semibold tracking-[0.16em] uppercase py-3.5" style={{ background: C.ink, color: C.bg, borderRadius: 3, textDecoration: "none" }}>Enquire</a>
                  )}
                </div>
              )}

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

      {state === "ready" && !["login", "signup", "forgot", "reset", "account", "bag", "checkout", "order"].includes(page) && (
        <section data-feature-row="" className="border-t" style={{ borderColor: C.line }}>
          <div className="max-w-7xl mx-auto px-5 py-8 grid grid-cols-2 md:grid-cols-4 gap-y-6">
            {[
              [Shirt, "Hand-picked", "Chosen piece by piece"],
              [Recycle, "Pre-loved", "Give clothes a longer life"],
              [ShieldCheck, "One of one", "Once it's gone, it's gone"],
              [Truck, "Tracked delivery", "Follow your order"],
            ].map(([Icon, h, t], i) => (
              <div key={h} className="flex items-center gap-3 px-3 md:px-6" style={{ borderLeft: i > 0 ? `1px solid ${C.line}` : "none" }}>
                <Icon size={26} strokeWidth={1.4} className="shrink-0" />
                <div><p className="text-[12px] font-semibold tracking-[0.1em] uppercase">{h}</p><p className="text-xs" style={{ color: C.muted }}>{t}</p></div>
              </div>
            ))}
          </div>
        </section>
      )}

      <footer className="border-t" style={{ borderColor: C.line }}>
        <div className="max-w-7xl mx-auto px-5 py-12 grid gap-10 md:grid-cols-[1.3fr_1fr_1fr_1fr]">
          <div>
            {store.logo ? <img src={store.logo} alt={slug} className="h-9 object-contain mb-3" /> : <p className="text-3xl mb-3" style={serif}>@{slug}</p>}
            <p className="text-sm max-w-[16rem]" style={{ color: C.muted }}>{store.intro || "Hand-picked second-hand pieces."}</p>
          </div>
          <div>
            <p className="text-[12px] font-semibold tracking-[0.12em] uppercase mb-4">Quick links</p>
            <ul className="space-y-2.5 text-sm" style={{ color: C.muted }}>
              <li><a href={window.location.pathname} onClick={(e) => { e.preventDefault(); go(""); }} style={{ color: "inherit", textDecoration: "none" }}>Home</a></li>
              {collections.length > 0 && <li><a href="#collections" onClick={(e) => { e.preventDefault(); go("#collections"); }} style={{ color: "inherit", textDecoration: "none" }}>Collections</a></li>}
              <li><a href="#faq" onClick={(e) => { e.preventDefault(); go("#faq"); }} style={{ color: "inherit", textDecoration: "none" }}>FAQ</a></li>
              {waDigits && <li><a href={waLink(null)} onClick={() => track("whatsapp", null)} target="_blank" rel="noreferrer" style={{ color: "inherit", textDecoration: "none" }}>Message us on WhatsApp</a></li>}
            </ul>
          </div>
          <div>
            <p className="text-[12px] font-semibold tracking-[0.12em] uppercase mb-4">Shop categories</p>
            <ul className="space-y-2.5 text-sm" style={{ color: C.muted }}>
              {categories.slice(0, 8).map((cat) => (
                <li key={cat}><button data-footer-category={cat} onClick={() => { if (itemId || page !== "home") go(""); pickCategory(cat); }}>{shopCategoryLabel(cat)}</button></li>
              ))}
            </ul>
          </div>
          <div>
            <p className="text-[12px] font-semibold tracking-[0.12em] uppercase mb-4">Account</p>
            <ul className="space-y-2.5 text-sm" style={{ color: C.muted }}>
              {(accountsOpen || shopper) && <li><a href={shopper ? "#account" : "#login"} onClick={(e) => { e.preventDefault(); go(shopper ? "#account" : "#login"); }} style={{ color: "inherit", textDecoration: "none" }}>{shopper ? "My orders" : "Log in"}</a></li>}
              {accountsOpen && !shopper && <li><a href="#signup" onClick={(e) => { e.preventDefault(); go("#signup"); }} style={{ color: "inherit", textDecoration: "none" }}>Create an account</a></li>}
            </ul>
          </div>
        </div>
        <div className="border-t" style={{ borderColor: C.line }}>
          <p className="max-w-7xl mx-auto px-5 py-5 text-[11px]" style={{ color: C.muted }}>Powered by Ravuno</p>
        </div>
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
// what a sale's postage did to your money: collected from the customer, minus what the courier actually charged.
// When the courier cost isn't known it is treated as "the same as the customer paid", so nothing changes until you enter it.
function postageMath(t) {
  const postage = Number(t.postage || 0);
  const collected = t.postageBy === "me" ? 0 : postage;
  const known = t.postageBy === "me" ? postage : (t.postageCost == null ? null : Number(t.postageCost));
  const cost = known == null ? collected : known;
  return { collected, cost, known: known == null ? 0 : known, profit: Math.round((collected - cost) * 100) / 100 };
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
  let revenue = 0, profit = 0, units = 0, orders = 0, postage = 0, postagePaid = 0;
  const cats = {}, brands = {};
  for (const t of txns || []) {
    const d = new Date(t.timestamp);
    if (!inRange(d)) continue;
    const row = rows[idx(d)];
    orders++; revenue += Number(t.total || 0); row.revenue += Number(t.total || 0);
    const pm = postageMath(t); postage += pm.collected; postagePaid += pm.known; profit += pm.profit; row.profit += pm.profit;
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
      units, orders, avgSale: units ? r2(revenue / units) : 0, spend: r2(spend), added, postage: r2(postage), postagePaid: r2(postagePaid),
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

      <AnalyticsCard title="Revenue & profit" hint={`${mode === "month" ? "Per day" : "Per month"} · excludes RM${t.postage} postage collected${t.postagePaid > 0 ? ` · RM${t.postagePaid} courier cost recorded, included in profit` : ""}`}>
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

function StorefrontTab({ storefront, onSaveStorefront, storefrontSaving, storefrontError, collections, onSaveFields, onSaveBanners, onUploadImage, onSaveCollection, onDeleteCollection, logoUrl, onLogoUpload, onLogoRemove, onLogoTransparent, logoMsg, logoUploading, items, onSaveStoreSale, onSetItemSale, onSetCollectionSale }) {
  const [slugDraft, setSlugDraft] = useState(storefront?.slug || "");
  const [whatsappDraft, setWhatsappDraft] = useState(storefront?.whatsappNumber || "");
  const [introDraft, setIntroDraft] = useState(storefront?.intro || "");
  useEffect(() => { setSlugDraft(storefront?.slug || ""); setWhatsappDraft(storefront?.whatsappNumber || ""); setIntroDraft(storefront?.intro || ""); }, [storefront?.slug, storefront?.whatsappNumber, storefront?.intro]);
  const logoInputRef = useRef(null);
  const [logoStrength, setLogoStrength] = useState("normal");
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
            <img src={logoUrl} alt="Your logo" data-logo-preview="" className="w-16 h-16 rounded-lg object-contain p-1" style={{ background: "repeating-conic-gradient(#2A2A2F 0% 25%, #1C1C20 0% 50%) 50% / 12px 12px" }} />
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
        {logoUrl && (
          <div className="flex items-center gap-2 mt-3 flex-wrap">
            <button data-logo-transparent="" onClick={() => onLogoTransparent(logoStrength)} disabled={logoUploading} className="text-xs font-semibold px-3 py-2 rounded-lg border" style={{ borderColor: "#C9A567", color: "#C9A567", opacity: logoUploading ? 0.6 : 1 }}>
              Make background transparent
            </button>
            <select value={logoStrength} onChange={(e) => setLogoStrength(e.target.value)} aria-label="Strength" className="text-xs px-2 py-2 rounded-lg border outline-none" style={{ borderColor: "#2A2A2F", background: "#17171A", color: "#F2F0EA" }}>
              <option value="gentle">Gentle</option><option value="normal">Normal</option><option value="strong">Strong</option>
            </select>
          </div>
        )}
        {logoMsg && <p data-logo-msg="" className="text-xs mt-3" style={{ color: "#C9A567" }}>{logoMsg}</p>}
        <p className="text-[11px] mt-2" style={{ color: "#75726C" }}>The checkerboard shows what is see-through. For the cleanest edges, upload the original file (a PNG) rather than a screenshot.</p>
        <input ref={logoInputRef} type="file" accept="image/*" className="hidden"
          onChange={(e) => { const f = e.target.files && e.target.files[0]; if (f) onLogoUpload(f); e.target.value = ""; }} />
      </div>
    </div>
  );
}

const emptyForm = {
  title: "", titleTouched: false,
  category: "", itemType: "",
  brand: "", condition: "", colour: "", colourCondition: "", remark: "", cost: "", price: "", sku: "", skuAuto: true,
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
  const [receiptTxn, setReceiptTxn] = useState(null);
  const [logoUrl, setLogoUrl] = useState(null);
  const [purchases, setPurchases] = useState([]);
  const [customers, setCustomers] = useState([]);
  const [messages, setMessages] = useState({ ship: "", promo: "" });
  const [saleCustomer, setSaleCustomer] = useState({ name: "", phone: "" });
  const [shipPhone, setShipPhone] = useState("");
  const [shipName, setShipName] = useState("");
  const [shipTracking, setShipTracking] = useState("");
  const [orders, setOrders] = useState(null); // null = not loaded yet
  const [orderPrivate, setOrderPrivate] = useState({});
  const [shopSettings, setShopSettings] = useState(null);
  const [checkoutOpenAdmin, setCheckoutOpenAdmin] = useState(false);
  const [purchaseForm, setPurchaseForm] = useState({ cost: "", place: "", photo: "", receiptPhoto: "", purchasedOn: new Date().toISOString().slice(0, 10) });
  const [purchaseSaving, setPurchaseSaving] = useState(false);
  const [purchasePeriod, setPurchasePeriod] = useState(() => { const d = new Date(); return { month: d.getMonth(), year: d.getFullYear() }; });
  const [purchaseDetailId, setPurchaseDetailId] = useState(null);
  const [storefront, setStorefront] = useState(null); // { slug, whatsappNumber, ... } | null
  const [collections, setCollections] = useState([]);
  const [storefrontSaving, setStorefrontSaving] = useState(false);
  const [storefrontError, setStorefrontError] = useState("");
  const [bannerUploading, setBannerUploading] = useState(false);
  const [logoUploading, setLogoUploading] = useState(false);
  const [logoMsg, setLogoMsg] = useState("");

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
  const [postageBy, setPostageBy] = useState("customer");
  const [orderDiscount, setOrderDiscount] = useState("");
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
        const [itemRows, txnRows, profileRows, templateRows, brandingRows, purchaseRows, storefrontRows, collectionRows, customerRows, messageRows] = await Promise.all([
          apiRequest("/items?select=*&order=created_at.desc", { method: "GET" }, session, updateSession),
          apiRequest("/transactions?select=*&order=occurred_at.desc", { method: "GET" }, session, updateSession),
          apiRequest("/profiles?select=*", { method: "GET" }, session, updateSession).catch(() => null),
          apiRequest("/templates?select=*", { method: "GET" }, session, updateSession).catch(() => null),
          apiRequest("/branding?select=*", { method: "GET" }, session, updateSession).catch(() => null),
          apiRequest("/purchases?select=*&order=purchased_on.desc", { method: "GET" }, session, updateSession).catch(() => null),
          apiRequest("/storefronts?select=*", { method: "GET" }, session, updateSession).catch(() => null),
          apiRequest("/collections?select=*&order=position.asc,created_at.asc", { method: "GET" }, session, updateSession).catch(() => null),
          apiRequest("/customers?select=*&order=created_at.desc", { method: "GET" }, session, updateSession).catch(() => null),
          apiRequest("/customer_messages?select=*", { method: "GET" }, session, updateSession).catch(() => null),
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
        if (customerRows !== null) setCustomers((customerRows || []).map(rowToCustomer));
        if (messageRows !== null) { const m = messageRows && messageRows[0]; setMessages({ ship: (m && m.ship_message) || "", promo: (m && m.promo_message) || "" }); }
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
      const { dataUrl, result } = await logoToPng(file, { transparent: true, strength: "normal" });
      const publicUrl = (await uploadPhotoToStorage(dataUrl, `${session.user.id}/logo.png`, session)) + `?v=${Date.now()}`;
      setLogoMsg(result.ok ? "Logo uploaded with a transparent background." : result.reason === "already-transparent" ? "Logo uploaded. It was already transparent." : "Logo uploaded. Couldn't find a plain background to remove.");
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

  async function handleLogoTransparent(strength) {
    if (!session || !logoUrl) return;
    setLogoUploading(true); setLogoMsg("");
    try {
      const { dataUrl, result } = await logoToPng(logoUrl, { transparent: true, strength });
      if (!result.ok) { setLogoMsg(result.reason === "already-transparent" ? "This logo is already transparent." : "Couldn't find a plain background to remove. Try Strong, or upload the original file."); return; }
      const publicUrl = (await uploadPhotoToStorage(dataUrl, `${session.user.id}/logo.png`, session)) + `?v=${Date.now()}`;
      await apiRequest("/branding?on_conflict=user_id", {
        method: "POST", body: JSON.stringify({ user_id: session.user.id, logo_url: publicUrl, updated_at: new Date().toISOString() }), headers: { Prefer: "resolution=merge-duplicates,return=representation" },
      }, session, updateSession);
      setLogoUrl(publicUrl);
      setLogoMsg(`Background removed (${Math.round(result.share * 100)}% of the picture). Not right? Try another strength, or upload the original again.`);
    } catch (e) {
      setLogoMsg("Couldn't process the logo. Try uploading the original file again.");
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
    setItems([]); setTxns([]); setCustomers([]); setMessages({ ship: "", promo: "" }); setForm(emptyForm); setCart([]); setActiveId(null);
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
      remark: item.remark || "", cost: item.cost ?? "", price: item.price ?? "", sku: item.sku || "", skuAuto: false,
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
      setPurchaseForm({ cost: "", place: "", photo: "", receiptPhoto: "", purchasedOn: new Date().toISOString().slice(0, 10) });
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
  const availableStock = items.filter((i) => i.status !== "sold" && !isReserved(i));
  const heldCount = items.filter((i) => i.status !== "sold" && isReserved(i)).length;
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
  const baseLines = cart.map((c) => {
    const item = items.find((i) => i.id === c.id);
    if (!item) return null;
    const originalPrice = Number(item.price || 0);
    const discount = Math.min(c.discount || 0, originalPrice);
    const finalPrice = originalPrice - discount;
    const onSale = saleFor(item, { store: storefront && storefront.sale, collections }).off > 0;
    return { ...item, discount, originalPrice, finalPrice, onSale };
  }).filter(Boolean);
  const cartSubtotal = baseLines.reduce((sum, l) => sum + l.originalPrice, 0);
  const itemDiscountTotal = baseLines.reduce((sum, l) => sum + l.discount, 0);
  const afterItemDiscounts = Math.round((cartSubtotal - itemDiscountTotal) * 100) / 100;
  const cartLines = allocateOrderDiscount(baseLines, Number(orderDiscount) || 0);
  const orderDiscountApplied = Math.round(cartLines.reduce((sum, l) => sum + l.orderShare, 0) * 100) / 100;
  const orderDiscountCapped = (Number(orderDiscount) || 0) > afterItemDiscounts + 0.004;
  const cartDiscount = Math.round((itemDiscountTotal + orderDiscountApplied) * 100) / 100;
  const cartTotal = Math.round((cartSubtotal - cartDiscount) * 100) / 100;
  const postageAmount = Math.max(0, Number(postage) || 0); // 0 = DFOD (customer pays the courier on delivery)
  const amountToCollect = Math.round((cartTotal + (postageBy === "customer" ? postageAmount : 0)) * 100) / 100;
  const salePhoneBad = !!saleCustomer.phone.trim() && !normalizePhone(saleCustomer.phone);
  const saleMatches = (() => {
    const digits = saleCustomer.phone.replace(/\D/g, "");
    const q = digits.startsWith("0") ? "60" + digits.slice(1) : digits;
    const n = saleCustomer.name.trim().toLowerCase();
    if (digits.length < 3 && n.length < 2) return [];
    const exact = normalizePhone(saleCustomer.phone);
    return customers.filter((c) => c.phone !== exact && ((digits.length >= 3 && (c.phone.includes(q) || c.phone.includes(digits))) || (n.length >= 2 && c.name.toLowerCase().includes(n)))).slice(0, 4);
  })();

  // finds the customer by phone (or creates them) so a sale can be linked to them
  async function ensureCustomer(phone, name) {
    const existing = customers.find((c) => c.phone === phone);
    const cleanName = String(name || "").trim();
    if (existing) {
      if (cleanName && cleanName !== existing.name) {
        await apiRequest(`/customers?id=eq.${existing.id}`, { method: "PATCH", body: JSON.stringify({ name: cleanName }) }, session, updateSession);
        setCustomers((prev) => prev.map((c) => (c.id === existing.id ? { ...c, name: cleanName } : c)));
      }
      return existing.id;
    }
    const [inserted] = await apiRequest("/customers", { method: "POST", body: JSON.stringify({ user_id: session.user.id, name: cleanName || null, phone }) }, session, updateSession);
    setCustomers((prev) => [rowToCustomer(inserted), ...prev]);
    return inserted.id;
  }

  async function completeSale() {
    if (cartLines.length === 0 || !session || salePhoneBad) return;
    setSaveError(false);
    const soldOn = new Date().toISOString().slice(0, 10);
    const soldIds = cartLines.map((l) => l.id);
    const salePhone = normalizePhone(saleCustomer.phone);
    try {
      const customerId = salePhone ? await ensureCustomer(salePhone, saleCustomer.name) : null;
      const txnRow = txnToRow({
        timestamp: new Date().toISOString(), method: payMethod, total: cartTotal, subtotal: cartSubtotal, discount: cartDiscount, postage: postageAmount, postageBy: postageAmount > 0 ? postageBy : "customer", customerId,
        lines: cartLines.map((l) => ({ title: productTitle(l), sku: l.sku, price: l.price, cost: Number(l.cost || 0), discount: Math.round((l.discount + l.orderShare) * 100) / 100, finalPrice: l.finalPrice })),
      }, session.user.id);
      await apiRequest(`/items?id=in.(${soldIds.join(",")})`, { method: "PATCH", body: JSON.stringify({ status: "sold", sold_on: soldOn }) }, session, updateSession);
      const [insertedTxn] = await apiRequest("/transactions", { method: "POST", body: JSON.stringify(txnRow) }, session, updateSession);
      setItems((prev) => prev.map((i) => (soldIds.includes(i.id) ? { ...i, status: "sold", soldOn } : i)));
      setTxns((prev) => [rowToTxn(insertedTxn), ...prev]);
      setReceiptTxn({ ...rowToTxn(insertedTxn), invoiceNumber: `INV-${String(txns.length + 1).padStart(4, "0")}` });
      setCart([]);
      setPostage("");
      setPostageBy("customer");
      setOrderDiscount("");
      setSaleCustomer({ name: "", phone: "" });
    } catch (e) {
      setSaveError(true);
    }
  }

  // the one-tap "send tracking" button calls this as WhatsApp opens: it remembers the tracking number and that it was sent
  async function markShipped(txnId, { phone, name, tracking }) {
    if (!session) return;
    setSaveError(false);
    try {
      const customerId = await ensureCustomer(phone, name);
      const shippedAt = new Date().toISOString();
      const trackingNo = String(tracking || "").trim();
      await apiRequest(`/transactions?id=eq.${txnId}`, { method: "PATCH", body: JSON.stringify({ customer_id: customerId, tracking_no: trackingNo, shipped_at: shippedAt }) }, session, updateSession);
      setTxns((prev) => prev.map((t) => (t.id === txnId ? { ...t, customerId, trackingNo, shippedAt } : t)));
    } catch (e) { setSaveError(true); }
  }

  async function refreshOrders() {
    if (!session) return;
    try {
      await apiRequest("/rpc/refresh_orders", { method: "POST", body: "{}" }, session, updateSession).catch(() => {});
      const [os, ps, ss] = await Promise.all([
        apiRequest("/orders?select=*,order_items(*),order_events(*)&order=created_at.desc&limit=200", { method: "GET" }, session, updateSession),
        apiRequest("/order_private?select=*", { method: "GET" }, session, updateSession).catch(() => []),
        apiRequest("/shop_settings?select=*", { method: "GET" }, session, updateSession).catch(() => []),
      ]);
      setOrders(os || []);
      const m = {}; (ps || []).forEach((x) => { m[x.order_id] = x; }); setOrderPrivate(m);
      setShopSettings((ss && ss[0]) || null);
      setCheckoutOpenAdmin(await checkoutIsOpen());
    } catch (e) { setOrders((prev) => prev || []); }
  }
  // Mark paid and refunds change pieces, sales and customers, so bring those up to date too
  async function reloadSales() {
    try {
      const [itemRows, txnRows, customerRows] = await Promise.all([
        apiRequest("/items?select=*&order=created_at.desc", { method: "GET" }, session, updateSession),
        apiRequest("/transactions?select=*&order=occurred_at.desc", { method: "GET" }, session, updateSession),
        apiRequest("/customers?select=*&order=created_at.desc", { method: "GET" }, session, updateSession).catch(() => null),
      ]);
      setItems((itemRows || []).map(rowToItem)); setTxns((txnRows || []).map(rowToTxn));
      if (customerRows !== null) setCustomers((customerRows || []).map(rowToCustomer));
    } catch (e) {}
  }
  async function actOnOrder(name, args) {
    try {
      await apiRequest(`/rpc/${name}`, { method: "POST", body: JSON.stringify(args) }, session, updateSession);
      await refreshOrders();
      if (name === "admin_mark_paid" || name === "admin_cancel_order") await reloadSales();
      return { ok: true };
    } catch (e) { return { ok: false, error: e.message }; }
  }
  async function saveOrderNote(orderId, note) {
    try {
      await apiRequest("/order_private?on_conflict=order_id", { method: "POST", body: JSON.stringify({ order_id: orderId, admin_note: String(note || "").trim() || null }), headers: { Prefer: "resolution=merge-duplicates,return=representation" } }, session, updateSession);
      setOrderPrivate((prev) => ({ ...prev, [orderId]: { ...(prev[orderId] || { order_id: orderId }), admin_note: String(note || "").trim() || null } }));
    } catch (e) { setSaveError(true); }
  }
  async function saveShopSettings(f) {
    try {
      const row = { user_id: session.user.id, ship_west: Number(f.west), ship_east: Number(f.east), free_ship_min: String(f.freeMin).trim() === "" ? null : Number(f.freeMin), hold_hours: Number(f.hold), payment_instructions: String(f.payment || "").trim() || null, updated_at: new Date().toISOString() };
      const rows = await apiRequest("/shop_settings?on_conflict=user_id", { method: "POST", body: JSON.stringify(row), headers: { Prefer: "resolution=merge-duplicates,return=representation" } }, session, updateSession);
      setShopSettings((rows && rows[0]) || row);
      return { ok: true };
    } catch (e) { return { ok: false, error: e.message }; }
  }
  async function toggleCheckout(open) {
    try {
      await apiRequest("/rpc/admin_set_checkout_open", { method: "POST", body: JSON.stringify({ p_open: !!open }) }, session, updateSession);
      setCheckoutOpenAdmin(!!open);
      return { ok: true };
    } catch (e) { return { ok: false, error: e.message }; }
  }
  const sessionUserId = session && session.user ? session.user.id : null;
  useEffect(() => { if (sessionUserId) refreshOrders(); }, [sessionUserId]);
  useEffect(() => {
    if (tab !== "orders" || !sessionUserId) return;
    refreshOrders();
    const t = setInterval(refreshOrders, 60000);
    return () => clearInterval(t);
  }, [tab, sessionUserId]);

  async function addCustomer(phone, name) {
    try {
      const existed = customers.some((c) => c.phone === phone);
      await ensureCustomer(phone, name);
      return { ok: true, existed };
    } catch (e) { setSaveError(true); return { ok: false }; }
  }

  async function renameCustomer(id, name) {
    const clean = String(name || "").trim();
    setCustomers((prev) => prev.map((c) => (c.id === id ? { ...c, name: clean } : c)));
    try { await apiRequest(`/customers?id=eq.${id}`, { method: "PATCH", body: JSON.stringify({ name: clean || null }) }, session, updateSession); }
    catch (e) { setSaveError(true); }
  }

  async function deleteCustomer(id) {
    setCustomers((prev) => prev.filter((c) => c.id !== id));
    setTxns((prev) => prev.map((t) => (t.customerId === id ? { ...t, customerId: null } : t)));
    try { await apiRequest(`/customers?id=eq.${id}`, { method: "DELETE" }, session, updateSession); }
    catch (e) { setSaveError(true); }
  }

  async function saveMessages(ship, promo) {
    const row = { user_id: session.user.id, ship_message: ship || null, promo_message: promo || null, updated_at: new Date().toISOString() };
    try {
      await apiRequest("/customer_messages?on_conflict=user_id", { method: "POST", body: JSON.stringify(row), headers: { Prefer: "resolution=merge-duplicates,return=representation" } }, session, updateSession);
      setMessages({ ship: ship || "", promo: promo || "" });
      return true;
    } catch (e) { setSaveError(true); return false; }
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
    const postageCollected = round2(periodTxns.reduce((sum, t) => sum + postageMath(t).collected, 0));
    const postagePaid = round2(periodTxns.reduce((sum, t) => sum + postageMath(t).known, 0));
    // net profit = final sale price minus cost, per line, plus any gap between the postage collected and what the courier charged (sales made before cost-tracking default cost to 0)
    const netProfit = round2(periodTxns.reduce((sum, t) => sum + t.lines.reduce((s, l) => s + (Number(l.finalPrice || 0) - Number(l.cost || 0)), 0) + postageMath(t).profit, 0));
    return { totalSales, orders, unitsSold, discounts, netProfit, postage: postageCollected, postagePaid };
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
  const storeLink = storefront && storefront.slug ? `${typeof window !== "undefined" ? window.location.origin : ""}/store/${storefront.slug}` : "";
  const shipTemplate = messages.ship || DEFAULT_SHIP_MESSAGE;
  const ordersTodo = (orders || []).filter((o) => orderBucket(o) === "todo").length;
  const ordersTabLabel = ordersTodo > 0 ? `Orders (${ordersTodo})` : "Orders";
  const shipPhoneNorm = normalizePhone(shipPhone);
  const shipReady = !!selectedInvoice && !!shipPhoneNorm && shipTracking.trim().length > 0;
  const shipText = renderMessage(shipTemplate, { name: shipName, tracking: shipTracking, store: storeLink });
  const shipLink = waChatLink(shipPhoneNorm, shipText);
  const customerById = useMemo(() => { const m = new Map(); customers.forEach((c) => m.set(c.id, c)); return m; }, [customers]);
  // opening an invoice pre-fills the buyer and any tracking number already saved on it
  useEffect(() => {
    const t = selectedInvoiceId ? txns.find((x) => x.id === selectedInvoiceId) : null;
    const c = t && t.customerId ? customers.find((x) => x.id === t.customerId) : null;
    setShipPhone(c ? formatPhone(c.phone) : "");
    setShipName(c ? c.name : "");
    setShipTracking(t ? t.trackingNo || "" : "");
  }, [selectedInvoiceId]);
  const MONTH_NAMES = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];

  const stats = useMemo(() => {
    const listed = items.filter((i) => i.status === "listed");
    const sold = items.filter((i) => i.status === "sold");
    const profit = sold.reduce((sum, i) => sum + (Number(i.price || 0) - Number(i.cost || 0)), 0);
    return { listedCount: listed.length, soldCount: sold.length, profit };
  }, [items]);

  // --- Stock Ledger: search + detail + export ---
  const skuDupIds = useMemo(() => skuDuplicateIds(items), [items]);
  const skuClash = useMemo(() => findSkuClashes(items, form.sku, editingId), [items, form.sku, editingId]);
  const nextSkuNo = useMemo(() => nextSkuNumber(items), [items]);
  // a new listing starts on the next free number; it stops following once you type your own
  useEffect(() => {
    if (editingId || !form.skuAuto) return;
    setForm((f) => { const want = joinSku(nextSkuNo, splitSku(f.sku).suffix); return !f.skuAuto || f.sku === want ? f : { ...f, sku: want }; });
  }, [nextSkuNo, editingId, form.skuAuto, form.sku]);
  const ledgerItems = useMemo(() => {
    const q = ledgerSearch.trim().toLowerCase();
    let base = items;
    const f = ledgerStatusFilter === "dupes" && skuDupIds.size === 0 ? "all" : ledgerStatusFilter;
    if (f === "listed") base = base.filter((i) => i.status !== "sold");
    if (f === "sold") base = base.filter((i) => i.status === "sold");
    if (f === "dupes") base = base.filter((i) => skuDupIds.has(i.id)).sort((a, b) => (Number(skuNumber(a.sku)) || 0) - (Number(skuNumber(b.sku)) || 0) || String(a.sku).localeCompare(String(b.sku)));
    if (!q) return base;
    return base.filter((i) =>
      (i.sku || "").toLowerCase().includes(q) ||
      (i.brand || "").toLowerCase().includes(q) ||
      (i.itemType || "").toLowerCase().includes(q) ||
      (CATEGORY_LABELS[i.category] || "").toLowerCase().includes(q) || shopCategoryLabel(shopCategoryOf(i)).toLowerCase().includes(q) ||
      productTitle(i).toLowerCase().includes(q) ||
      (i.status || "").toLowerCase().includes(q)
    );
  }, [items, ledgerSearch, ledgerStatusFilter, skuDupIds]);

  const crosslistItems = useMemo(() => {
    const q = crosslistSearch.trim().toLowerCase();
    if (!q) return items;
    return items.filter((i) =>
      (i.sku || "").toLowerCase().includes(q) ||
      (i.brand || "").toLowerCase().includes(q) ||
      (CATEGORY_LABELS[i.category] || "").toLowerCase().includes(q) || shopCategoryLabel(shopCategoryOf(i)).toLowerCase().includes(q) ||
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
              ...(() => { const r = resolveImportCategory(get("category"), get("itemType") || get("item_type") || ""); return { category: r.category, item_type: r.itemType || null }; })(),
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
            <span className="text-[10px] font-medium tracking-wider" style={{ color: "#75726C" }}>V4.18</span>
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
            <button onClick={handleSignOut} className="text-xs font-medium" style={{ color: "#75726C" }}>
              Sign out
            </button>
          </div>
          <div className="flex items-baseline gap-2">
            <span className="text-2xl font-semibold tracking-tight" style={{ color: "#F2F0EA", fontFamily: "'Fraunces', serif" }}>RAVUNO</span>
            <span className="text-[10px] font-medium tracking-wider" style={{ color: "#75726C", fontFamily: "'IBM Plex Mono', monospace" }}>V4.18</span>
          </div>
          <nav className="flex gap-1 p-1 rounded-lg overflow-x-auto no-scrollbar max-w-full" style={{ background: "#0B0B0C", border: "1px solid #2A2A2F", WebkitOverflowScrolling: "touch", scrollSnapType: "x proximity" }}>
            {[{ id: "crosslist", label: "Cross-List" }, { id: "inventory", label: "Stock Ledger" }, { id: "pos", label: "POS" }, { id: "orders", label: ordersTabLabel }, { id: "purchases", label: "Purchases" }, { id: "sales", label: "Sales" }, { id: "customers", label: "Customers" }, { id: "analytics", label: "Analytics" }, { id: "storefront", label: "Storefront" }].map((t) => (
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
                          <div className="grid grid-cols-2 gap-3">
                          <Field label="Cost (RM)" value={form.cost} onChange={set("cost")} placeholder="15" type="number" />
                          <Field label="Price (RM)" value={form.price} onChange={set("price")} placeholder="45" type="number" />
                        </div>
                        <div className="grid grid-cols-2 gap-3" data-sku-boxes="">
                          <Field label={editingId ? "SKU number" : `SKU number (next: ${nextSkuNo})`} value={splitSku(form.sku).num} placeholder={String(nextSkuNo)}
                            onChange={(e) => { const v = e.target.value; setForm((f) => ({ ...f, sku: joinSku(v, splitSku(f.sku).suffix), skuAuto: false })); }} />
                          <Field label="Letter (optional)" value={splitSku(form.sku).suffix} placeholder="M"
                            onChange={(e) => { const v = e.target.value; setForm((f) => ({ ...f, sku: joinSku(splitSku(f.sku).num, v) })); }} />
                        </div>
                        <p data-sku-preview="" className="text-[11px]" style={{ color: "#8C8A85", fontFamily: "'IBM Plex Mono', monospace" }}>SKU will be saved as <span style={{ color: "#F2F0EA" }}>{form.sku || "(none)"}</span></p>
                        {(skuClash.exact.length > 0 || skuClash.sameNumber.length > 0) && (
                          <div data-sku-note="" className="rounded-lg px-3 py-2 text-[11px] leading-relaxed space-y-1" style={{ background: "#2A1F14", border: "1px solid #C9A567", color: "#E8C98A" }}>
                            {skuClash.exact.length > 0 && <p>⚠ SKU {form.sku.trim()} is already used by {describeSkuItems(skuClash.exact)}.</p>}
                            {skuClash.sameNumber.length > 0 && <p>⚠ Number {skuNumber(form.sku)} is already used by {describeSkuItems(skuClash.sameNumber)}. Same number with a different letter, so double-check it isn't a mix-up.</p>}
                          </div>
                        )}
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
                  {[{ id: "all", label: "All" }, { id: "listed", label: "Listed" }, { id: "sold", label: "Sold" }, ...(skuDupIds.size ? [{ id: "dupes", label: `⚠ Duplicates ${skuDupIds.size}` }] : [])].map((f) => (
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

                {ledgerStatusFilter === "dupes" && skuDupIds.size > 0 && (
                  <p data-dupes-note="" className="text-[11px] mb-3" style={{ color: "#C9A567" }}>These items share a SKU number with another item. Open one to change its SKU.</p>
                )}
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
                          <span className="truncate" style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{i.sku || "-"}{skuDupIds.has(i.id) && <span title="Same SKU number as another item" style={{ color: "#C9A567" }}> ⚠</span>}{isReserved(i) && <span data-held-tag="" title="Held for a customer order" style={{ color: "#C9A567" }}> · Reserved</span>}</span>
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
                              {i.sku || "-"}{skuDupIds.has(i.id) ? " ⚠" : ""}{isReserved(i) ? " · Reserved" : ""} · {CATEGORY_LABELS[i.category] || "-"} · RM{i.price || 0}
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

                {heldCount > 0 && (
                  <p data-held-note="" className="text-xs text-center -mt-2 mb-4" style={{ color: "#C9A567" }}>{heldCount} piece{heldCount === 1 ? " is" : "s are"} held for customer orders and can't be sold here until paid or released. Find them under Orders.</p>
                )}
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
                        {(l.discount > 0 || l.orderShare > 0) && (
                          <p className="text-[11px] mt-1 text-right" style={{ color: "#75726C", fontFamily: "'IBM Plex Mono', monospace" }}>
                            RM{l.originalPrice}{l.discount > 0 ? ` − RM${l.discount}` : ""}{l.orderShare > 0 ? ` − RM${l.orderShare} (order discount)` : ""}
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
                    {itemDiscountTotal > 0 && (
                      <div className="flex items-center justify-between" style={{ color: "#8C8A85" }}>
                        <span>Item discounts</span><span style={{ fontFamily: "'IBM Plex Mono', monospace" }}>− RM{itemDiscountTotal}</span>
                      </div>
                    )}
                    <div className="flex items-center justify-between gap-3 py-1" style={{ color: "#8C8A85" }} data-order-discount="">
                      <span>Discount on whole order (RM)</span>
                      <input type="number" inputMode="decimal" min="0" value={orderDiscount} onChange={(e) => setOrderDiscount(e.target.value)} placeholder="0" aria-label="Discount on whole order"
                        className="w-20 px-2 py-1.5 rounded-md border text-xs text-center outline-none" style={{ borderColor: "#2A2A2F", background: "#17171A", color: "#F2F0EA", fontFamily: "'IBM Plex Mono', monospace" }} />
                    </div>
                    {orderDiscountCapped && <p className="text-[11px]" style={{ color: "#C9695E" }}>Capped at RM{afterItemDiscounts}, what's left after item discounts.</p>}
                    <div className="flex items-center justify-between" style={{ color: "#8C8A85" }} data-total-discount="">
                      <span>Total discount</span><span style={{ fontFamily: "'IBM Plex Mono', monospace" }}>− RM{cartDiscount}</span>
                    </div>
                    <div className="flex items-center justify-between font-semibold pt-1" style={{ color: "#F2F0EA" }}>
                      <span>Total</span><span style={{ fontFamily: "'IBM Plex Mono', monospace" }}>RM{cartTotal}</span>
                    </div>
                  </div>

                  <div className="mb-4" data-pos-postage="">
                    <p className="text-[10px] uppercase tracking-wide mb-1.5" style={{ color: "#75726C", fontFamily: "'IBM Plex Mono', monospace" }}>Postage</p>
                    <div className="flex flex-wrap items-center gap-2">
                      <button onClick={() => setPostage("")} className="text-xs font-semibold px-3 py-2 rounded-lg border"
                        style={{ borderColor: postageAmount === 0 ? "#C9A567" : "#2A2A2F", color: postageAmount === 0 ? "#C9A567" : "#8C8A85", background: postageAmount === 0 ? "#241F13" : "transparent" }}>
                        DFOD
                      </button>
                      <input type="number" inputMode="decimal" min="0" value={postage} onChange={(e) => setPostage(e.target.value)} placeholder="Custom RM" aria-label="Custom postage"
                        className="w-28 px-3 py-2 rounded-lg border text-xs outline-none" style={{ borderColor: postageAmount > 0 ? "#C9A567" : "#2A2A2F", background: "#0B0B0C", color: "#F2F0EA" }} />
                    </div>
                    <p className="text-[11px] mt-1.5" style={{ color: "#75726C" }}>DFOD = customer pays the courier on delivery, so nothing is added here.</p>
                    {postageAmount > 0 && (
                      <div className="mt-3" data-postage-by="">
                        <p className="text-[10px] uppercase tracking-wide mb-1.5" style={{ color: "#75726C", fontFamily: "'IBM Plex Mono', monospace" }}>Postage paid by</p>
                        <div className="flex gap-2">
                          {[["customer", "Customer"], ["me", "Me"]].map(([k, label]) => {
                            const on = postageBy === k;
                            return (
                              <button key={k} data-paid-by={k} onClick={() => setPostageBy(k)} className="text-xs font-semibold px-4 py-2 rounded-lg border"
                                style={{ borderColor: on ? "#C9A567" : "#2A2A2F", color: on ? "#C9A567" : "#8C8A85", background: on ? "#241F13" : "transparent" }}>{label}</button>
                            );
                          })}
                        </div>
                        <p className="text-[11px] mt-1.5" style={{ color: "#75726C" }}>
                          {postageBy === "customer" ? "Customer pays the postage on top of the items." : "You cover the postage. It comes off your profit and the customer only pays for the items."}
                        </p>
                        <div className="flex items-center justify-between font-semibold pt-2 text-sm" style={{ color: "#F2F0EA" }}>
                          <span>Customer pays</span><span style={{ fontFamily: "'IBM Plex Mono', monospace" }}>RM{amountToCollect}</span>
                        </div>
                      </div>
                    )}
                  </div>

                  <div className="mb-4" data-pos-customer="">
                    <p className="text-[10px] uppercase tracking-wide mb-1.5" style={{ color: "#75726C", fontFamily: "'IBM Plex Mono', monospace" }}>Customer (optional)</p>
                    <div className="grid grid-cols-2 gap-2">
                      <input type="tel" inputMode="tel" value={saleCustomer.phone} aria-label="Customer WhatsApp number" placeholder="WhatsApp 011-2345 6789"
                        onChange={(e) => { const v = e.target.value; const norm = normalizePhone(v); const ex = norm && customers.find((c) => c.phone === norm); setSaleCustomer((p) => ({ phone: v, name: p.name || (ex ? ex.name : "") })); }}
                        className="w-full px-3 py-2 rounded-lg border text-xs outline-none" style={{ borderColor: salePhoneBad ? "#C9695E" : "#2A2A2F", background: "#0B0B0C", color: "#F2F0EA" }} />
                      <input value={saleCustomer.name} maxLength={80} aria-label="Customer name" placeholder="Name" onChange={(e) => setSaleCustomer((p) => ({ ...p, name: e.target.value }))}
                        className="w-full px-3 py-2 rounded-lg border text-xs outline-none" style={{ borderColor: "#2A2A2F", background: "#0B0B0C", color: "#F2F0EA" }} />
                    </div>
                    {salePhoneBad && <p className="text-[11px] mt-1.5" style={{ color: "#C9695E" }}>That number looks incomplete. Fix it or clear it to continue.</p>}
                    {saleMatches.length > 0 && (
                      <div className="flex flex-wrap gap-1.5 mt-2">
                        {saleMatches.map((c) => (
                          <button key={c.id} data-customer-chip="" onClick={() => setSaleCustomer({ name: c.name, phone: formatPhone(c.phone) })}
                            className="text-[11px] px-2.5 py-1.5 rounded-lg border" style={{ borderColor: "#2A2A2F", color: "#C9A567", background: "#17171A" }}>
                            {c.name || "No name"} · {formatPhone(c.phone)}
                          </button>
                        ))}
                      </div>
                    )}
                    <p className="text-[11px] mt-1.5" style={{ color: "#75726C" }}>Saved so you can send tracking in one tap later.</p>
                  </div>

                  <Select label="Payment method" value={payMethod} onChange={(e) => setPayMethod(e.target.value)}>
                    <option>Cash</option><option>Bank Transfer</option><option>Online</option>
                  </Select>
                  <button onClick={completeSale} disabled={cartLines.length === 0 || salePhoneBad}
                    className="mt-4 w-full flex items-center justify-center gap-2 py-2.5 rounded-xl text-sm font-semibold transition-opacity"
                    style={{ background: "#C9A567", color: "#0B0B0C", opacity: (cartLines.length === 0 || salePhoneBad) ? 0.35 : 1, fontFamily: "'IBM Plex Mono', monospace" }}>
                    Complete Sale
                  </button>
                </div>
              </section>
            </div>
          )}

          {tab === "orders" && (
            <OrdersTab orders={orders} privateById={orderPrivate} settings={shopSettings} checkoutOpen={checkoutOpenAdmin} shipTemplate={shipTemplate} storeLink={storeLink}
              onRefresh={refreshOrders} onAct={actOnOrder} onSaveNote={saveOrderNote} onSaveSettings={saveShopSettings} onToggleCheckout={toggleCheckout} />
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
                          <p className="text-sm font-medium truncate" style={{ color: "#F2F0EA" }}>{p.place || "Purchase"}</p>
                          <p className="text-xs truncate" style={{ color: "#75726C" }}>{new Date(p.purchasedOn).toLocaleDateString()}</p>
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
                    {salesSummary.postagePaid > 0 && <StatCard icon={Receipt} label="Postage you paid (RM)" value={salesSummary.postagePaid} tone="#C9695E" />}
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
                            {(t.customerId || t.shippedAt) && (
                              <p className="text-[11px] truncate" style={{ color: "#C9A567" }}>
                                {t.customerId && customerById.get(t.customerId) ? (customerById.get(t.customerId).name || formatPhone(customerById.get(t.customerId).phone)) : ""}
                                {t.shippedAt ? `${t.customerId ? " · " : ""}Tracking sent ✓` : ""}
                              </p>
                            )}
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

          {tab === "customers" && (
            <CustomersTab customers={customers} txns={txns} messages={messages} storeLink={storeLink}
              onRename={renameCustomer} onDelete={deleteCustomer} onSaveMessages={saveMessages} onAdd={addCustomer} />
          )}

          {tab === "analytics" && (
            <AnalyticsTab session={session} updateSession={updateSession} storefront={storefront} items={items} txns={txns} purchases={purchases} />
          )}

          {tab === "storefront" && (
            <StorefrontTab storefront={storefront} onSaveStorefront={saveStorefrontSettings} storefrontSaving={storefrontSaving} storefrontError={storefrontError}
              collections={collections} onSaveFields={saveStoreFields} onSaveBanners={saveBanners} onUploadImage={uploadStoreImage} onSaveCollection={saveCollection} onDeleteCollection={deleteCollection}
              logoUrl={logoUrl} onLogoUpload={handleLogoUpload} onLogoRemove={handleLogoRemove} onLogoTransparent={handleLogoTransparent} logoMsg={logoMsg} logoUploading={logoUploading}
              items={items} onSaveStoreSale={saveStoreSale} onSetItemSale={setItemSale} onSetCollectionSale={setCollectionSale} />
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
              {selectedInvoice.postage > 0 && <DetailRow label={selectedInvoice.postageBy === "me" ? "Postage (paid by you)" : "Postage"} value={`RM${selectedInvoice.postage}`} />}
              {selectedInvoice.postageCost != null && selectedInvoice.postageBy !== "me" && <DetailRow label="Courier charged you" value={`RM${selectedInvoice.postageCost}`} />}
              {selectedInvoice.postage > 0 && <DetailRow label="Customer paid" value={`RM${Math.round((Number(selectedInvoice.total) + (selectedInvoice.postageBy === "me" ? 0 : Number(selectedInvoice.postage))) * 100) / 100}`} />}
            </dl>

            <div data-ship-section="" className="mt-5 pt-4 border-t" style={{ borderColor: "#2A2A2F" }}>
              <p className="text-[10px] uppercase tracking-wide mb-2" style={{ color: "#75726C" }}>Customer & shipping</p>
              <div className="grid grid-cols-2 gap-2 mb-2">
                <input type="tel" inputMode="tel" value={shipPhone} onChange={(e) => setShipPhone(e.target.value)} aria-label="Customer WhatsApp number" placeholder="WhatsApp 011-2345 6789"
                  className="w-full px-3 py-2 rounded-lg border text-xs outline-none" style={{ borderColor: shipPhone.trim() && !shipPhoneNorm ? "#C9695E" : "#2A2A2F", background: "#0B0B0C", color: "#F2F0EA" }} />
                <input value={shipName} maxLength={80} onChange={(e) => setShipName(e.target.value)} aria-label="Customer name" placeholder="Name"
                  className="w-full px-3 py-2 rounded-lg border text-xs outline-none" style={{ borderColor: "#2A2A2F", background: "#0B0B0C", color: "#F2F0EA" }} />
              </div>
              {shipPhone.trim() && !shipPhoneNorm && <p className="text-[11px] mb-2" style={{ color: "#C9695E" }}>That number looks incomplete.</p>}
              <input value={shipTracking} maxLength={60} onChange={(e) => setShipTracking(e.target.value)} aria-label="Tracking number" placeholder="Tracking number"
                className="w-full px-3 py-2 rounded-lg border text-xs outline-none" style={{ borderColor: "#2A2A2F", background: "#0B0B0C", color: "#F2F0EA", fontFamily: "'IBM Plex Mono', monospace" }} />
              {shipReady ? (
                <>
                  <p className="text-[11px] mt-3 mb-1" style={{ color: "#75726C" }}>Message preview</p>
                  <pre data-ship-preview="" className="whitespace-pre-wrap text-xs rounded-lg p-3" style={{ background: "#0B0B0C", color: "#C9C7C0", fontFamily: "inherit", border: "1px solid #2A2A2F" }}>{shipText}</pre>
                  <a data-ship-send="" href={shipLink} target="_blank" rel="noopener noreferrer"
                    onClick={() => markShipped(selectedInvoice.id, { phone: shipPhoneNorm, name: shipName, tracking: shipTracking })}
                    className="mt-3 w-full flex items-center justify-center py-2.5 rounded-xl text-sm font-semibold" style={{ background: "#C9A567", color: "#0B0B0C", textDecoration: "none" }}>
                    {selectedInvoice.shippedAt ? "Resend on WhatsApp" : "Send on WhatsApp"}
                  </a>
                </>
              ) : (
                <p className="text-[11px] mt-2" style={{ color: "#75726C" }}>Add the buyer's WhatsApp number and the tracking number to send it in one tap.</p>
              )}
              {selectedInvoice.shippedAt && (
                <p className="text-[11px] mt-2" style={{ color: "#C9A567" }}>Tracking sent on {new Date(selectedInvoice.shippedAt).toLocaleDateString(undefined, { day: "2-digit", month: "short", year: "numeric" })}.</p>
              )}
            </div>
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
                <div className="text-xs mb-1"><div className="flex justify-between"><span>Postage</span><span>{receiptTxn.postageBy === "me" ? "FREE" : `RM${receiptTxn.postage}`}</span></div></div>
              )}
              <div className="text-sm font-bold flex justify-between pt-2 mt-2 border-t" style={{ borderColor: "#17161A" }}>
                <span>TOTAL</span><span>RM{Math.round((Number(receiptTxn.total) + (receiptTxn.postageBy === "me" ? 0 : Number(receiptTxn.postage || 0))) * 100) / 100}</span>
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
