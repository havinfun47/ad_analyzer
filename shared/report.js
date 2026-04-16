/* ============================================================
   Scale Science — Meta Ads Report Engine
   ============================================================ */

/* ── Constants ─────────────────────────────────────────────── */
const PURCHASE_ACTION  = "offsite_conversion.fb_pixel_purchase";
const CART_ACTION      = "offsite_conversion.fb_pixel_add_to_cart";
const CHECKOUT_ACTION  = "offsite_conversion.fb_pixel_initiate_checkout";
const OUTBOUND_CLICK   = "link_click";

/* ── State ─────────────────────────────────────────────────── */
let currentClient   = null;
let currentRange    = "last_7d";
let currentDates    = null;
let currentCompare  = null;
let sortState       = {};

/* ── Date Helpers ──────────────────────────────────────────── */
function formatDate(d) {
  return d.toISOString().split("T")[0];
}

function parseDateStr(s) {
  const [y, m, d] = s.split("-").map(Number);
  return new Date(y, m - 1, d);
}

function displayDate(s) {
  const d = parseDateStr(s);
  return d.toLocaleDateString("en-CA", { month: "short", day: "numeric", year: "numeric" });
}

function getRangeDates(rangeKey, customStart, customEnd) {
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const yesterday = new Date(today); yesterday.setDate(today.getDate() - 1);

  let since, until, compSince, compUntil;

  const daysAgo = (n) => { const d = new Date(today); d.setDate(today.getDate() - n); return d; };

  switch (rangeKey) {
    case "last_7d":
      since = daysAgo(7); until = yesterday;
      compSince = daysAgo(14); compUntil = daysAgo(8);
      break;
    case "last_14d":
      since = daysAgo(14); until = yesterday;
      compSince = daysAgo(28); compUntil = daysAgo(15);
      break;
    case "last_30d":
      since = daysAgo(30); until = yesterday;
      compSince = daysAgo(60); compUntil = daysAgo(31);
      break;
    case "last_90d":
      since = daysAgo(90); until = yesterday;
      compSince = daysAgo(180); compUntil = daysAgo(91);
      break;
    case "this_month": {
      since = new Date(today.getFullYear(), today.getMonth(), 1);
      until = yesterday;
      const prevEnd = new Date(since); prevEnd.setDate(prevEnd.getDate() - 1);
      const prevStart = new Date(prevEnd.getFullYear(), prevEnd.getMonth(), 1);
      compSince = prevStart; compUntil = prevEnd;
      break;
    }
    case "custom":
      since = parseDateStr(customStart); until = parseDateStr(customEnd);
      const days = Math.round((until - since) / 86400000) + 1;
      compUntil = new Date(since); compUntil.setDate(compUntil.getDate() - 1);
      compSince = new Date(compUntil); compSince.setDate(compUntil.getDate() - days + 1);
      break;
    default:
      since = daysAgo(7); until = yesterday;
      compSince = daysAgo(14); compUntil = daysAgo(8);
  }

  return {
    since:     formatDate(since),
    until:     formatDate(until),
    compSince: formatDate(compSince),
    compUntil: formatDate(compUntil)
  };
}

/* ── Currency Formatter ────────────────────────────────────── */
function formatCurrency(val, currency = "CAD") {
  if (val == null || isNaN(val)) return "—";
  return new Intl.NumberFormat("en-CA", {
    style: "currency",
    currency,
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  }).format(val);
}

function formatNum(val, decimals = 2) {
  if (val == null || isNaN(val)) return "—";
  return Number(val).toLocaleString("en-CA", { minimumFractionDigits: decimals, maximumFractionDigits: decimals });
}

function formatPct(val) {
  if (val == null || isNaN(val)) return "—";
  return (Number(val) * 1).toFixed(2) + "%";
}

function formatRoas(val) {
  if (val == null || isNaN(val)) return "—";
  return Number(val).toFixed(2) + "x";
}

/* ── API Helpers ───────────────────────────────────────────── */
function getToken() {
  return localStorage.getItem(TOKEN_KEY) || "";
}

async function apiGet(path, params = {}) {
  const token = getToken();
  const url   = new URL(`${META_API_BASE}/${path}`);
  url.searchParams.set("access_token", token);
  for (const [k, v] of Object.entries(params)) {
    url.searchParams.set(k, typeof v === "object" ? JSON.stringify(v) : v);
  }

  const res = await fetch(url.toString());
  const data = await res.json();

  if (data.error) {
    throw new Error(data.error.message || "Meta API error");
  }
  return data;
}

async function fetchInsights(adAccountId, level, timeRange, extraFields = []) {
  const baseFields = [
    "spend", "impressions", "cpm", "frequency",
    "outbound_clicks", "outbound_clicks_ctr",
    "action_values", "actions", "cost_per_action_type",
    "purchase_roas"
  ];
  const fields = [...new Set([...baseFields, ...extraFields])].join(",");

  let all = [];
  let nextUrl = null;

  const params = {
    fields,
    time_range: JSON.stringify({ since: timeRange.since, until: timeRange.until }),
    level,
    limit: 100
  };

  do {
    let data;
    if (nextUrl) {
      const res = await fetch(nextUrl);
      data = await res.json();
      if (data.error) throw new Error(data.error.message);
    } else {
      data = await apiGet(`${adAccountId}/insights`, params);
    }

    if (data.data) all = all.concat(data.data);
    nextUrl = data.paging?.next || null;
  } while (nextUrl);

  return all;
}

/* ── Metric Extractors ─────────────────────────────────────── */
function getAction(row, actionType) {
  const actions = row.actions || [];
  const found = actions.find(a => a.action_type === actionType);
  return found ? parseFloat(found.value) : 0;
}

function getActionValue(row, actionType) {
  const vals = row.action_values || [];
  const found = vals.find(a => a.action_type === actionType);
  return found ? parseFloat(found.value) : 0;
}

function getCostPerAction(row, actionType) {
  const cpas = row.cost_per_action_type || [];
  const found = cpas.find(a => a.action_type === actionType);
  return found ? parseFloat(found.value) : null;
}

function getOutboundClicks(row) {
  // Try outbound_clicks array — Meta returns {action_type, value} pairs
  const oc = row.outbound_clicks;
  if (Array.isArray(oc) && oc.length > 0) {
    const found = oc.find(a =>
      a.action_type === "link_click" || a.action_type === "outbound_click"
    );
    if (found) return parseFloat(found.value) || 0;
    // Only one entry with a different action_type — use it anyway
    if (oc.length === 1) return parseFloat(oc[0].value) || 0;
  }
  // Flat number fallback (some API versions / breakdowns return a scalar)
  if (oc != null && !Array.isArray(oc) && !isNaN(parseFloat(oc))) {
    return parseFloat(oc);
  }
  // Last resort: pull link_click from the generic actions array
  const actions = row.actions || [];
  const lc = actions.find(a => a.action_type === "link_click");
  if (lc) return parseFloat(lc.value) || 0;
  return 0;
}

function getOutboundCtr(row) {
  // Try the dedicated outbound_clicks_ctr field first
  const ctrField = row.outbound_clicks_ctr;
  if (Array.isArray(ctrField) && ctrField.length > 0) {
    const found = ctrField.find(a =>
      a.action_type === "link_click" || a.action_type === "outbound_click"
    );
    if (found) return parseFloat(found.value) || 0;
    if (ctrField.length === 1) return parseFloat(ctrField[0].value) || 0;
  }
  if (ctrField != null && !Array.isArray(ctrField) && !isNaN(parseFloat(ctrField))) {
    return parseFloat(ctrField);
  }
  // Calculate manually: outbound_clicks / impressions * 100
  const clicks = getOutboundClicks(row);
  const impressions = parseFloat(row.impressions || 0);
  return impressions > 0 ? (clicks / impressions) * 100 : 0;
}

function parseRoas(row) {
  if (!row.purchase_roas || !row.purchase_roas.length) return 0;
  return parseFloat(row.purchase_roas[0].value) || 0;
}

function buildAccountMetrics(rows, currency) {
  if (!rows || rows.length === 0) return null;

  // Sum over all rows (usually one at account level)
  let spend = 0, impressions = 0, purchases = 0, revenue = 0, outboundClicks = 0, freq = 0, cpmSum = 0, rows_count = 0;

  for (const r of rows) {
    spend         += parseFloat(r.spend || 0);
    impressions   += parseFloat(r.impressions || 0);
    purchases     += getAction(r, PURCHASE_ACTION);
    revenue       += getActionValue(r, PURCHASE_ACTION);
    outboundClicks += getOutboundClicks(r);
    cpmSum        += parseFloat(r.cpm || 0) * parseFloat(r.impressions || 0);
    freq          += parseFloat(r.frequency || 0) * parseFloat(r.impressions || 0);
    rows_count    += 1;
  }

  const roas       = spend > 0 ? revenue / spend : 0;
  const cpa        = purchases > 0 ? spend / purchases : null;
  const aov        = purchases > 0 ? revenue / purchases : null;
  const cpm        = impressions > 0 ? cpmSum / impressions : 0;
  const frequency  = impressions > 0 ? freq / impressions : 0;
  const ctr        = outboundClicks > 0 && impressions > 0 ? (outboundClicks / impressions) * 100 : 0;
  const convRate   = outboundClicks > 0 ? (purchases / outboundClicks) * 100 : 0;

  return { spend, impressions, purchases, revenue, outboundClicks, roas, cpa, aov, cpm, frequency, ctr, convRate };
}

/* ── % Change Badge ────────────────────────────────────────── */
function pctChange(current, previous) {
  if (!previous || previous === 0) return null;
  return ((current - previous) / Math.abs(previous)) * 100;
}

function changeBadge(current, previous, lowerIsBetter = false) {
  const pct = pctChange(current, previous);
  if (pct === null) return `<span class="kpi-badge neutral">—</span>`;

  const isPositive = lowerIsBetter ? pct < 0 : pct > 0;
  const cls        = isPositive ? "positive" : "negative";
  const arrow      = pct > 0 ? "↑" : "↓";
  const label      = `${arrow} ${Math.abs(pct).toFixed(1)}%`;

  return `<span class="kpi-badge ${cls}">${label}</span>`;
}

/* ── Render KPI Cards ──────────────────────────────────────── */
function renderKPIs(metrics, prevMetrics, currency) {
  const curr = metrics;
  const prev = prevMetrics;
  const c = currency || "CAD";

  const cards = [
    {
      label: "Amount Spent",
      value: formatCurrency(curr.spend, c),
      badge: changeBadge(curr.spend, prev?.spend, true),
      compare: prev ? `vs ${formatCurrency(prev.spend, c)}` : ""
    },
    {
      label: "Revenue",
      value: formatCurrency(curr.revenue, c),
      badge: changeBadge(curr.revenue, prev?.revenue, false),
      compare: prev ? `vs ${formatCurrency(prev.revenue, c)}` : ""
    },
    {
      label: "Purchase ROAS",
      value: formatRoas(curr.roas),
      badge: changeBadge(curr.roas, prev?.roas, false),
      compare: prev ? `vs ${formatRoas(prev.roas)}` : ""
    },
    {
      label: "Cost per Purchase",
      value: curr.cpa != null ? formatCurrency(curr.cpa, c) : "—",
      badge: prev?.cpa != null ? changeBadge(curr.cpa, prev?.cpa, true) : `<span class="kpi-badge neutral">—</span>`,
      compare: prev?.cpa != null ? `vs ${formatCurrency(prev.cpa, c)}` : ""
    },
    {
      label: "Avg. Order Value",
      value: curr.aov != null ? formatCurrency(curr.aov, c) : "—",
      badge: prev?.aov != null ? changeBadge(curr.aov, prev?.aov, false) : `<span class="kpi-badge neutral">—</span>`,
      compare: prev?.aov != null ? `vs ${formatCurrency(prev.aov, c)}` : ""
    },
    {
      label: "CTR (Outbound)",
      value: formatPct(curr.ctr),
      badge: changeBadge(curr.ctr, prev?.ctr, false),
      compare: prev ? `vs ${formatPct(prev.ctr)}` : ""
    },
    {
      label: "CPM",
      value: formatCurrency(curr.cpm, c),
      badge: changeBadge(curr.cpm, prev?.cpm, true),
      compare: prev ? `vs ${formatCurrency(prev.cpm, c)}` : ""
    },
    {
      label: "Frequency",
      value: formatNum(curr.frequency, 2),
      badge: changeBadge(curr.frequency, prev?.frequency, true),
      compare: prev ? `vs ${formatNum(prev.frequency, 2)}` : ""
    },
    {
      label: "Conversion Rate",
      value: formatPct(curr.convRate),
      badge: changeBadge(curr.convRate, prev?.convRate, false),
      compare: prev ? `vs ${formatPct(prev.convRate)}` : ""
    }
  ];

  return cards.map(card => `
    <div class="kpi-card">
      <div class="kpi-label">${card.label}</div>
      <div class="kpi-value">${card.value}</div>
      <div style="display:flex;align-items:center;gap:8px;flex-wrap:wrap;">
        ${card.badge}
        <span class="kpi-compare">${card.compare}</span>
      </div>
    </div>
  `).join("");
}

/* ── Render Campaign Table ─────────────────────────────────── */
function buildCampaignRows(rows, currency) {
  return rows.map(r => {
    const spend     = parseFloat(r.spend || 0);
    const roas      = parseRoas(r);
    const outClicks = getOutboundClicks(r);
    const ctr       = getOutboundCtr(r);        // uses full fallback chain
    const cpm       = parseFloat(r.cpm || 0);
    const frequency = parseFloat(r.frequency || 0);
    const purchases = getAction(r, PURCHASE_ACTION);
    const convRate  = outClicks > 0 ? (purchases / outClicks) * 100 : 0;
    return { name: r.campaign_name || "—", spend, roas, ctr, cpm, frequency, convRate };
  });
}

function renderCampaignTable(rows, currency) {
  const c = currency || "CAD";
  const tableId = "tbl-campaign";
  const cols = [
    { key: "name",     label: "Campaign",        numeric: false },
    { key: "spend",    label: "Amount Spent",     numeric: true,  fmt: v => formatCurrency(v, c), lowerBetter: true },
    { key: "roas",     label: "Purchase ROAS",    numeric: true,  fmt: formatRoas },
    { key: "ctr",      label: "CTR",              numeric: true,  fmt: formatPct },
    { key: "cpm",      label: "CPM",              numeric: true,  fmt: v => formatCurrency(v, c), lowerBetter: true },
    { key: "frequency",label: "Frequency",        numeric: true,  fmt: v => formatNum(v, 2) },
    { key: "convRate", label: "Conv. Rate",       numeric: true,  fmt: formatPct }
  ];

  if (!rows.length) return `<div class="table-empty">No campaign data for this period.</div>`;

  // Totals
  const totals = {
    spend: rows.reduce((s, r) => s + r.spend, 0),
    roas:  rows.reduce((s, r) => s + r.roas * r.spend, 0) / (rows.reduce((s, r) => s + r.spend, 0) || 1),
    ctr:   rows.reduce((s, r) => s + r.ctr, 0) / (rows.length || 1),
    cpm:   rows.reduce((s, r) => s + r.cpm, 0) / (rows.length || 1),
    frequency: rows.reduce((s, r) => s + r.frequency, 0) / (rows.length || 1),
    convRate: rows.reduce((s, r) => s + r.convRate, 0) / (rows.length || 1)
  };

  return renderSortableTable(tableId, cols, rows, totals, c);
}

/* ── Render Ad Set Table ───────────────────────────────────── */
function buildAdSetRows(rows, currency) {
  return rows.map(r => {
    const spend     = parseFloat(r.spend || 0);
    const revenue   = getActionValue(r, PURCHASE_ACTION);
    const roas      = parseRoas(r);
    const ctr       = getOutboundCtr(r);        // uses full fallback chain
    const cpm       = parseFloat(r.cpm || 0);
    const frequency = parseFloat(r.frequency || 0);
    return { name: r.adset_name || "—", spend, revenue, roas, ctr, cpm, frequency };
  });
}

function renderAdSetTable(rows, currency) {
  const c = currency || "CAD";
  const tableId = "tbl-adset";
  const cols = [
    { key: "name",     label: "Ad Set",           numeric: false },
    { key: "spend",    label: "Amount Spent",     numeric: true, fmt: v => formatCurrency(v, c) },
    { key: "revenue",  label: "Revenue",          numeric: true, fmt: v => formatCurrency(v, c) },
    { key: "roas",     label: "Purchase ROAS",    numeric: true, fmt: formatRoas },
    { key: "ctr",      label: "Outbound CTR",     numeric: true, fmt: formatPct },
    { key: "cpm",      label: "CPM",              numeric: true, fmt: v => formatCurrency(v, c), lowerBetter: true },
    { key: "frequency",label: "Frequency",        numeric: true, fmt: v => formatNum(v, 2) }
  ];

  if (!rows.length) return `<div class="table-empty">No ad set data for this period.</div>`;

  const totals = {
    spend:   rows.reduce((s, r) => s + r.spend, 0),
    revenue: rows.reduce((s, r) => s + r.revenue, 0),
    roas:    rows.reduce((s, r) => s + r.spend, 0) > 0
               ? rows.reduce((s, r) => s + r.revenue, 0) / rows.reduce((s, r) => s + r.spend, 0)
               : 0,
    ctr:      rows.reduce((s, r) => s + r.ctr, 0) / (rows.length || 1),
    cpm:      rows.reduce((s, r) => s + r.cpm, 0) / (rows.length || 1),
    frequency: rows.reduce((s, r) => s + r.frequency, 0) / (rows.length || 1)
  };

  return renderSortableTable(tableId, cols, rows, totals, c);
}

/* ── Fetch Ad Thumbnails ───────────────────────────────────── */
async function fetchAdThumbnails(adAccountId) {
  try {
    let all = [];
    let next = null;
    do {
      let data;
      if (next) {
        const res = await fetch(next);
        data = await res.json();
        if (data.error) throw new Error(data.error.message);
      } else {
        data = await apiGet(`${adAccountId}/ads`, {
          fields: "id,name,creative{thumbnail_url,image_url,object_type,object_story_spec{link_data{picture},video_data{thumbnail_url}}}",
          limit: 200
        });
      }
      if (data.data) all = all.concat(data.data);
      next = data.paging?.next || null;
    } while (next);

    const map = {};
    for (const ad of all) {
      const cr = ad.creative || {};
      const objectType = (cr.object_type || "").toUpperCase();
      // Try multiple fields — different ad formats populate different ones
      const thumbnailUrl =
        cr.thumbnail_url ||
        cr.object_story_spec?.video_data?.thumbnail_url ||
        cr.object_story_spec?.link_data?.picture ||
        cr.image_url ||
        null;
      map[ad.id] = {
        thumbnailUrl,
        isVideo: objectType === "VIDEO"
      };
    }
    return map;
  } catch (e) {
    console.warn("Thumbnail fetch failed (non-fatal):", e.message);
    return {};
  }
}

/* ── Render Ad Table ───────────────────────────────────────── */
function buildAdRows(rows, thumbnails, currency) {
  return rows.map(r => {
    const spend       = parseFloat(r.spend || 0);
    const purchases   = getAction(r, PURCHASE_ACTION);
    const checkouts   = getAction(r, CHECKOUT_ACTION);
    const cartAdds    = getAction(r, CART_ACTION);
    const outClicks   = getOutboundClicks(r);
    const roas        = parseRoas(r);

    // Use API cost_per_action_type first; fall back to manual spend / conversions
    const cpPurchase  = getCostPerAction(r, PURCHASE_ACTION)
                        ?? (purchases > 0 ? spend / purchases : null);
    const cpCheckout  = getCostPerAction(r, CHECKOUT_ACTION)
                        ?? (checkouts > 0 ? spend / checkouts : null);
    const cpCart      = getCostPerAction(r, CART_ACTION)
                        ?? (cartAdds > 0 ? spend / cartAdds : null);
    const cpClick     = getCostPerAction(r, "link_click")
                        ?? (outClicks > 0 ? spend / outClicks : null);

    const cpm         = parseFloat(r.cpm || 0);
    const frequency   = parseFloat(r.frequency || 0);

    // Thumbnail from separate ads endpoint, matched by ad_id
    const thumb       = thumbnails?.[r.ad_id] || {};

    return {
      name: r.ad_name || "—",
      spend, purchases, roas,
      cpPurchase, cpCheckout, cpCart, cpClick,
      cpm, frequency,
      thumbnailUrl: thumb.thumbnailUrl || null,
      isVideo:      thumb.isVideo      || false
    };
  });
}

function renderAdTable(rows, currency) {
  const c = currency || "CAD";
  const tableId = "tbl-ad";
  const nullFmt = (fn) => v => v != null ? fn(v) : "—";
  const cols = [
    {
      key: "name",
      label: "Ad",
      numeric: false,
      render(val, row) {
        const placeholder = `<div class="ad-thumb-placeholder">
          <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5">
            <rect x="3" y="3" width="18" height="18" rx="2"/><circle cx="8.5" cy="8.5" r="1.5"/>
            <polyline points="21 15 16 10 5 21"/>
          </svg>
        </div>`;

        let thumbEl;
        if (row.thumbnailUrl) {
          if (row.isVideo) {
            // Video: wrap in badge div so the ::after play overlay applies
            thumbEl = `<div class="ad-thumb-video-badge">
              <img class="ad-thumb" src="${row.thumbnailUrl}" loading="lazy"
                   onerror="this.closest('.ad-thumb-video-badge').style.display='none'">
            </div>`;
          } else {
            // Image ad: plain img, hide on error
            thumbEl = `<img class="ad-thumb" src="${row.thumbnailUrl}" loading="lazy"
                            onerror="this.style.display='none'">`;
          }
        } else {
          thumbEl = placeholder;
        }

        return `<div class="ad-thumb-cell">${thumbEl}<span class="ad-name-text">${val}</span></div>`;
      }
    },
    { key: "spend",      label: "Amount Spent",        numeric: true, fmt: v => formatCurrency(v, c) },
    { key: "purchases",  label: "Purchases",           numeric: true, fmt: v => formatNum(v, 0) },
    { key: "roas",       label: "Purchase ROAS",       numeric: true, fmt: formatRoas },
    { key: "cpPurchase", label: "Cost / Purchase",     numeric: true, fmt: nullFmt(v => formatCurrency(v, c)), lowerBetter: true },
    { key: "cpCheckout", label: "Cost / Checkout",     numeric: true, fmt: nullFmt(v => formatCurrency(v, c)), lowerBetter: true },
    { key: "cpCart",     label: "Cost / Add to Cart",  numeric: true, fmt: nullFmt(v => formatCurrency(v, c)), lowerBetter: true },
    { key: "cpClick",    label: "Cost / Click",        numeric: true, fmt: nullFmt(v => formatCurrency(v, c)), lowerBetter: true },
    { key: "cpm",        label: "CPM",                 numeric: true, fmt: v => formatCurrency(v, c), lowerBetter: true },
    { key: "frequency",  label: "Frequency",           numeric: true, fmt: v => formatNum(v, 2) }
  ];

  if (!rows.length) return `<div class="table-empty">No ad data for this period.</div>`;

  const totals = {
    spend:     rows.reduce((s, r) => s + r.spend, 0),
    purchases: rows.reduce((s, r) => s + r.purchases, 0),
    roas:      null,
    cpPurchase: null, cpCheckout: null, cpCart: null, cpClick: null,
    cpm:       rows.reduce((s, r) => s + r.cpm, 0) / (rows.length || 1),
    frequency: rows.reduce((s, r) => s + r.frequency, 0) / (rows.length || 1)
  };

  return renderSortableTable(tableId, cols, rows, totals, c);
}

/* ── Generic Sortable Table ────────────────────────────────── */
function renderSortableTable(tableId, cols, rows, totals, currency) {
  const state = sortState[tableId] || { col: null, dir: 1 };

  let sorted = [...rows];
  if (state.col) {
    sorted.sort((a, b) => {
      const av = a[state.col] ?? -Infinity;
      const bv = b[state.col] ?? -Infinity;
      if (typeof av === "string") return av.localeCompare(bv) * state.dir;
      return (av - bv) * state.dir;
    });
  }

  const headers = cols.map(col => {
    const isSorted = state.col === col.key;
    const icon = isSorted ? (state.dir === 1 ? "↑" : "↓") : "↕";
    const cls = `${col.numeric ? "th-num" : ""} ${isSorted ? "sorted" : ""}`;
    return `<th class="${cls}" data-table="${tableId}" data-col="${col.key}" style="cursor:pointer;">
      ${col.label}<i class="sort-icon">${icon}</i>
    </th>`;
  }).join("");

  const bodyRows = sorted.map(row => {
    const cells = cols.map(col => {
      const val  = row[col.key];
      const cell = col.render
        ? col.render(val, row)
        : (col.fmt || (v => v))(val);
      const cls  = col.numeric ? "td-num" : "td-name";
      return `<td class="${cls}">${cell}</td>`;
    }).join("");
    return `<tr>${cells}</tr>`;
  }).join("");

  const footCells = cols.map(col => {
    const val = totals?.[col.key];
    const fmt = col.fmt || (v => v);
    const cls = col.numeric ? "td-num" : "td-name";
    if (col.key === "name") return `<td class="${cls}">Totals</td>`;
    if (val == null) return `<td class="${cls} td-num">—</td>`;
    return `<td class="${cls}">${fmt(val)}</td>`;
  }).join("");

  return `
    <div class="table-scroll">
      <table class="data-table" id="${tableId}">
        <thead><tr>${headers}</tr></thead>
        <tbody>${bodyRows}</tbody>
        <tfoot><tr>${footCells}</tr></tfoot>
      </table>
    </div>`;
}

/* ── Attach sort listeners ─────────────────────────────────── */
function attachSortListeners() {
  document.querySelectorAll("[data-table][data-col]").forEach(th => {
    th.addEventListener("click", () => {
      const tableId = th.dataset.table;
      const col     = th.dataset.col;
      const state   = sortState[tableId] || { col: null, dir: 1 };

      if (state.col === col) {
        sortState[tableId] = { col, dir: state.dir * -1 };
      } else {
        sortState[tableId] = { col, dir: -1 }; // desc by default on first click
      }

      refreshTables();
    });
  });
}

/* ── LocalStorage for editable sections ───────────────────── */
function editableKey(type) {
  const clientKey = currentClient?.key || "unknown";
  const range     = currentRange === "custom"
    ? `${currentDates?.since}_${currentDates?.until}`
    : currentRange;
  return `editorial_${clientKey}_${range}_${type}`;
}

function saveEditorial(type, html) {
  localStorage.setItem(editableKey(type), html);
}

function loadEditorial(type) {
  return localStorage.getItem(editableKey(type)) || "";
}

function initEditorialBlocks() {
  ["summary", "nextsteps"].forEach(type => {
    const el = document.getElementById(`editorial-${type}`);
    if (!el) return;

    const saved = loadEditorial(type);
    if (saved) el.innerHTML = saved;

    el.addEventListener("input", () => saveEditorial(type, el.innerHTML));

    // Handle paste — strip color styles, handle images
    el.addEventListener("paste", e => {
      e.preventDefault();
      const cd = e.clipboardData;
      if (!cd) return;

      // Image paste takes priority
      for (const item of cd.items || []) {
        if (item.type.startsWith("image/")) {
          const file = item.getAsFile();
          const reader = new FileReader();
          reader.onload = ev => document.execCommand("insertImage", false, ev.target.result);
          reader.readAsDataURL(file);
          return;
        }
      }

      // HTML paste — strip color/background/font styles so text is visible on dark bg
      const html = cd.getData("text/html");
      if (html) {
        const doc = new DOMParser().parseFromString(html, "text/html");
        doc.querySelectorAll("[style]").forEach(node => {
          let s = node.getAttribute("style");
          s = s.replace(/\bcolor\s*:[^;]+;?/gi, "")
               .replace(/\bbackground(-color)?\s*:[^;]+;?/gi, "")
               .replace(/\bfont-family\s*:[^;]+;?/gi, "")
               .replace(/\bfont-size\s*:[^;]+;?/gi, "");
          s.trim() ? node.setAttribute("style", s) : node.removeAttribute("style");
        });
        // Remove legacy color/bgcolor attributes
        doc.querySelectorAll("[color],[bgcolor]").forEach(node => {
          node.removeAttribute("color");
          node.removeAttribute("bgcolor");
        });
        document.execCommand("insertHTML", false, doc.body.innerHTML);
        return;
      }

      // Plain text fallback
      const text = cd.getData("text/plain");
      if (text) document.execCommand("insertText", false, text);
    });

    // Tab = indent bullet, Shift+Tab = outdent bullet
    el.addEventListener("keydown", e => {
      if (e.key === "Tab") {
        e.preventDefault();
        document.execCommand(e.shiftKey ? "outdent" : "indent");
      }
    });

    // Handle image drag-drop
    el.addEventListener("dragover", e => e.preventDefault());
    el.addEventListener("drop", e => {
      e.preventDefault();
      const file = e.dataTransfer?.files?.[0];
      if (file && file.type.startsWith("image/")) {
        const reader = new FileReader();
        reader.onload = ev => {
          document.execCommand("insertImage", false, ev.target.result);
        };
        reader.readAsDataURL(file);
      }
    });
  });

  // Toolbar buttons
  document.querySelectorAll(".toolbar-btn").forEach(btn => {
    btn.addEventListener("mousedown", e => {
      e.preventDefault();
      const cmd   = btn.dataset.cmd;
      const value = btn.dataset.value || null;

      if (cmd === "createLink") {
        const url = prompt("Enter URL:");
        if (url) document.execCommand("createLink", false, url);
      } else {
        document.execCommand(cmd, false, value);
      }
    });
  });
}

/* ── Render Period Label ───────────────────────────────────── */
function renderPeriodLabel(dates) {
  if (!dates) return "";
  return `${displayDate(dates.since)} – ${displayDate(dates.until)}`;
}

function renderComparePeriodLabel(dates) {
  if (!dates) return "";
  return `compared to ${displayDate(dates.compSince)} – ${displayDate(dates.compUntil)}`;
}

/* ── Main Load ─────────────────────────────────────────────── */
let _accountData   = null;
let _accountPrev   = null;
let _campaignData  = null;
let _adsetData     = null;
let _adData        = null;

async function loadReport() {
  const token = getToken();
  if (!token) {
    showTokenScreen();
    return;
  }

  showLoading(true);
  showError(null);

  const dates = currentDates;
  const client = currentClient;

  try {
    const [accCurr, accPrev, campaigns, adsets, ads, thumbnails] = await Promise.all([
      fetchInsights(client.adAccountId, "account",  { since: dates.since,     until: dates.until }),
      fetchInsights(client.adAccountId, "account",  { since: dates.compSince, until: dates.compUntil }),
      fetchInsights(client.adAccountId, "campaign", { since: dates.since,     until: dates.until }, ["campaign_name"]),
      fetchInsights(client.adAccountId, "adset",    { since: dates.since,     until: dates.until }, ["adset_name"]),
      fetchInsights(client.adAccountId, "ad",       { since: dates.since,     until: dates.until }, ["ad_name", "ad_id"]),
      fetchAdThumbnails(client.adAccountId)
    ]);

    _accountData  = buildAccountMetrics(accCurr,    client.currency);
    _accountPrev  = buildAccountMetrics(accPrev,     client.currency);
    _campaignData = buildCampaignRows(campaigns,     client.currency);
    _adsetData    = buildAdSetRows(adsets,            client.currency);
    _adData       = buildAdRows(ads, thumbnails,      client.currency);

    renderReport();
  } catch (err) {
    console.error(err);
    showError(err.message || "Failed to load data.");
  } finally {
    showLoading(false);
  }
}

function renderReport() {
  const client  = currentClient;
  const dates   = currentDates;
  const currency = client.currency;

  // Period labels
  document.getElementById("period-label").textContent  = renderPeriodLabel(dates);
  document.getElementById("compare-label").textContent = renderComparePeriodLabel(dates);

  // Print header
  const ph = document.getElementById("print-period");
  if (ph) ph.textContent = `${renderPeriodLabel(dates)} ${renderComparePeriodLabel(dates)}`;

  // KPI cards
  const kpiEl = document.getElementById("kpi-grid");
  if (kpiEl && _accountData) {
    kpiEl.innerHTML = renderKPIs(_accountData, _accountPrev, currency);
  } else if (kpiEl) {
    kpiEl.innerHTML = `<div class="empty-state"><h3>No data</h3><p>No metrics found for this period.</p></div>`;
  }

  // Default sort: spend descending (only set if user hasn't manually sorted)
  if (!sortState["tbl-campaign"]) sortState["tbl-campaign"] = { col: "spend", dir: -1 };
  if (!sortState["tbl-adset"])    sortState["tbl-adset"]    = { col: "spend", dir: -1 };
  if (!sortState["tbl-ad"])       sortState["tbl-ad"]       = { col: "spend", dir: -1 };

  // Tables
  refreshTables();

  // Re-attach sort listeners
  attachSortListeners();

  // Show report
  document.getElementById("report-content").style.display = "";
}

function refreshTables() {
  const currency = currentClient?.currency || "CAD";

  const campEl = document.getElementById("campaign-table-container");
  if (campEl && _campaignData) campEl.innerHTML = renderCampaignTable(_campaignData, currency);

  const adsetEl = document.getElementById("adset-table-container");
  if (adsetEl && _adsetData) adsetEl.innerHTML = renderAdSetTable(_adsetData, currency);

  const adEl = document.getElementById("ad-table-container");
  if (adEl && _adData) adEl.innerHTML = renderAdTable(_adData, currency);

  attachSortListeners();
}

/* ── UI State ──────────────────────────────────────────────── */
function showLoading(on) {
  const el = document.getElementById("loading-state");
  if (!el) return;
  el.classList.toggle("active", on);
  // report-content visibility is controlled by renderReport(), not here
  if (on) {
    const rc = document.getElementById("report-content");
    if (rc) rc.style.display = "none";
  }
}

function showError(msg) {
  const el  = document.getElementById("error-state");
  const txt = document.getElementById("error-message");
  if (!el) return;
  if (msg) {
    el.classList.add("active");
    if (txt) txt.textContent = msg;
  } else {
    el.classList.remove("active");
  }
}

function showTokenScreen() {
  const ts = document.getElementById("token-screen");
  if (ts) ts.classList.add("active");
}

function hideTokenScreen() {
  const ts = document.getElementById("token-screen");
  if (ts) ts.classList.remove("active");
}

/* ── Date Range Controls ───────────────────────────────────── */
function initDateControls() {
  const sel       = document.getElementById("date-range-select");
  const customRow = document.getElementById("custom-date-row");
  const applyBtn  = document.getElementById("apply-custom");

  if (!sel) return;

  sel.addEventListener("change", () => {
    currentRange = sel.value;
    if (currentRange === "custom") {
      customRow?.classList.add("active");
    } else {
      customRow?.classList.remove("active");
      currentDates = getRangeDates(currentRange);
      loadReport();
    }
  });

  applyBtn?.addEventListener("click", () => {
    const s = document.getElementById("custom-since")?.value;
    const u = document.getElementById("custom-until")?.value;
    if (!s || !u) return;
    currentDates = getRangeDates("custom", s, u);
    loadReport();
  });
}

/* ── Token Management ──────────────────────────────────────── */
function initTokenControls() {
  // Submit token
  document.getElementById("token-submit")?.addEventListener("click", () => {
    const val = document.getElementById("token-input")?.value?.trim();
    if (!val) return;
    localStorage.setItem(TOKEN_KEY, val);
    hideTokenScreen();
    loadReport();
  });

  // Disconnect
  document.getElementById("btn-disconnect")?.addEventListener("click", () => {
    localStorage.removeItem(TOKEN_KEY);
    location.reload();
  });

  // Re-enter after error
  document.getElementById("btn-reenter")?.addEventListener("click", () => {
    showError(null);
    showTokenScreen();
  });
}

/* ============================================================
   TRENDS PAGE
   ============================================================ */

let _trendsMode  = "weekly";
let _trendsCache = {};   // { weekly: rows[], monthly: rows[] }

/* ── Inject tab bar + trends section into DOM ─────────────── */
function injectTrendsUI() {
  const rc = document.getElementById("report-content");
  if (!rc || document.getElementById("tab-bar")) return;

  // Tab bar (before report-content, inside <main>)
  const tabBar     = document.createElement("div");
  tabBar.id        = "tab-bar";
  tabBar.className = "no-print";
  tabBar.innerHTML = `
    <button class="tab-btn active" data-tab="report">
      <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2">
        <rect x="3" y="3" width="7" height="7" rx="1.5"/>
        <rect x="14" y="3" width="7" height="7" rx="1.5"/>
        <rect x="3" y="14" width="7" height="7" rx="1.5"/>
        <rect x="14" y="14" width="7" height="7" rx="1.5"/>
      </svg>
      Report
    </button>
    <button class="tab-btn" data-tab="trends">
      <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2">
        <polyline points="22 12 18 12 15 21 9 3 6 12 2 12"/>
      </svg>
      Trends
    </button>`;
  rc.before(tabBar);

  // Trends content (after report-content, inside <main>)
  const trendsDiv       = document.createElement("div");
  trendsDiv.id          = "trends-content";
  trendsDiv.className   = "no-print";
  trendsDiv.style.display = "none";
  trendsDiv.innerHTML   = `
    <div class="trends-header">
      <div class="trends-title-row">
        <div class="section-label">Performance Trends</div>
        <div class="trends-toggle">
          <button class="toggle-btn active" data-mode="weekly">Weekly</button>
          <button class="toggle-btn" data-mode="monthly">Monthly</button>
        </div>
      </div>
    </div>
    <div id="trends-table-wrap"></div>`;
  rc.after(trendsDiv);
}

/* ── Tab + toggle event handling ──────────────────────────── */
function initTabs() {
  document.addEventListener("click", e => {
    // Main tabs
    const tabBtn = e.target.closest(".tab-btn[data-tab]");
    if (tabBtn) {
      const tab = tabBtn.dataset.tab;
      document.querySelectorAll(".tab-btn[data-tab]").forEach(b =>
        b.classList.toggle("active", b === tabBtn));
      document.getElementById("report-content").style.display =
        tab === "report" ? "" : "none";
      const tc = document.getElementById("trends-content");
      if (tc) tc.style.display = tab === "trends" ? "" : "none";
      if (tab === "trends" && !_trendsCache[_trendsMode]) loadTrends(_trendsMode);
      return;
    }

    // Weekly/Monthly toggle
    const toggleBtn = e.target.closest(".toggle-btn[data-mode]");
    if (toggleBtn) {
      const mode = toggleBtn.dataset.mode;
      _trendsMode = mode;
      document.querySelectorAll(".toggle-btn[data-mode]").forEach(b =>
        b.classList.toggle("active", b === toggleBtn));
      if (_trendsCache[mode]) {
        renderTrendsFromCache();
      } else {
        loadTrends(mode);
      }
    }
  });
}

/* ── Fetch time-series insights ───────────────────────────── */
async function fetchTrendsInsights(mode) {
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const yesterday = new Date(today);
  yesterday.setDate(today.getDate() - 1);

  const since = new Date(today);
  if (mode === "weekly") {
    since.setDate(today.getDate() - 91);   // ~13 weeks
  } else {
    since.setFullYear(today.getFullYear() - 1); // 12 months
  }

  const fields = [
    "spend", "impressions", "cpm", "frequency",
    "outbound_clicks", "action_values", "actions",
    "date_start", "date_stop"
  ].join(",");

  const data = await apiGet(`${currentClient.adAccountId}/insights`, {
    fields,
    time_range:     JSON.stringify({ since: formatDate(since), until: formatDate(yesterday) }),
    time_increment: mode === "weekly" ? 7 : "monthly",
    level:          "account",
    limit:          100
  });

  return data.data || [];
}

/* ── Build display rows from API response ─────────────────── */
function buildTrendsRows(rows) {
  return rows
    .map(r => {
      const spend      = parseFloat(r.spend || 0);
      const revenue    = getActionValue(r, PURCHASE_ACTION);
      const purchases  = getAction(r, PURCHASE_ACTION);
      const outClicks  = getOutboundClicks(r);
      const impressions = parseFloat(r.impressions || 0);

      return {
        dateStart:  r.date_start,
        dateStop:   r.date_stop,
        spend,
        revenue,
        roas:       spend > 0 ? revenue / spend : 0,
        cpa:        purchases > 0 ? spend / purchases : null,
        ctr:        impressions > 0 ? (outClicks / impressions) * 100 : 0,
        cpm:        parseFloat(r.cpm || 0),
        frequency:  parseFloat(r.frequency || 0)
      };
    })
    .sort((a, b) => b.dateStart.localeCompare(a.dateStart)); // most recent first
}

/* ── Format period label ──────────────────────────────────── */
function formatTrendsPeriod(dateStart, dateStop, mode) {
  if (mode === "monthly") {
    return parseDateStr(dateStart)
      .toLocaleDateString("en-CA", { month: "long", year: "numeric" });
  }
  const sf = parseDateStr(dateStart).toLocaleDateString("en-CA", { month: "short", day: "numeric" });
  const ef = parseDateStr(dateStop).toLocaleDateString("en-CA",  { month: "short", day: "numeric" });
  return `${sf} – ${ef}`;
}

/* ── Render trends table HTML ─────────────────────────────── */
function renderTrendsTable(rows, mode) {
  const c = currentClient?.currency || "CAD";
  if (!rows.length) return `<div class="table-empty">No data for this period.</div>`;

  const cols = [
    { label: "Period",           key: "period",    num: false },
    { label: "Amount Spent",     key: "spend",     num: true,  fmt: v => formatCurrency(v, c) },
    { label: "Revenue",          key: "revenue",   num: true,  fmt: v => formatCurrency(v, c) },
    { label: "ROAS",             key: "roas",      num: true,  fmt: formatRoas },
    { label: "Cost / Purchase",  key: "cpa",       num: true,  fmt: v => v != null ? formatCurrency(v, c) : "—" },
    { label: "Outbound CTR",     key: "ctr",       num: true,  fmt: formatPct },
    { label: "CPM",              key: "cpm",       num: true,  fmt: v => formatCurrency(v, c) },
    { label: "Frequency",        key: "frequency", num: true,  fmt: v => formatNum(v, 2) }
  ];

  const headers = cols.map(col =>
    `<th class="${col.num ? "th-num" : ""}">${col.label}</th>`
  ).join("");

  const bodyRows = rows.map(row => {
    const cells = cols.map(col => {
      if (col.key === "period") {
        return `<td class="td-name trends-period">${formatTrendsPeriod(row.dateStart, row.dateStop, mode)}</td>`;
      }
      return `<td class="td-num">${col.fmt(row[col.key])}</td>`;
    }).join("");
    return `<tr>${cells}</tr>`;
  }).join("");

  // Totals / averages
  const totalSpend   = rows.reduce((s, r) => s + r.spend, 0);
  const totalRevenue = rows.reduce((s, r) => s + r.revenue, 0);
  const totalPurch   = rows.reduce((s, r) => s + (r.cpa != null ? r.spend / r.cpa : 0), 0);
  const foot = {
    spend:     totalSpend,
    revenue:   totalRevenue,
    roas:      totalSpend > 0 ? totalRevenue / totalSpend : 0,
    cpa:       totalPurch > 0 ? totalSpend / totalPurch : null,
    ctr:       rows.reduce((s, r) => s + r.ctr, 0)       / (rows.length || 1),
    cpm:       rows.reduce((s, r) => s + r.cpm, 0)       / (rows.length || 1),
    frequency: rows.reduce((s, r) => s + r.frequency, 0) / (rows.length || 1)
  };

  const footCells = cols.map(col => {
    if (col.key === "period") return `<td class="td-name">Total / Avg</td>`;
    return `<td class="td-num">${col.fmt(foot[col.key])}</td>`;
  }).join("");

  return `
    <div class="table-scroll">
      <table class="data-table">
        <thead><tr>${headers}</tr></thead>
        <tbody>${bodyRows}</tbody>
        <tfoot><tr>${footCells}</tr></tfoot>
      </table>
    </div>`;
}

/* ── Load + render orchestrator ───────────────────────────── */
async function loadTrends(mode) {
  const wrap = document.getElementById("trends-table-wrap");
  if (!wrap) return;

  wrap.innerHTML = `<div class="table-wrapper">
    <div class="skeleton skeleton-table" style="height:240px;border-radius:0;"></div>
  </div>`;

  try {
    const raw  = await fetchTrendsInsights(mode);
    _trendsCache[mode] = buildTrendsRows(raw);
    renderTrendsFromCache();
  } catch (err) {
    wrap.innerHTML = `<div class="table-wrapper" style="padding:32px;text-align:center;">
      <p style="color:var(--red);font-size:13px;">${err.message}</p>
    </div>`;
  }
}

function renderTrendsFromCache() {
  const wrap = document.getElementById("trends-table-wrap");
  if (!wrap) return;
  const rows = _trendsCache[_trendsMode] || [];
  wrap.innerHTML = `<div class="table-wrapper">${renderTrendsTable(rows, _trendsMode)}</div>`;
}

/* ── Bootstrap ─────────────────────────────────────────────── */
function initReport(clientKey) {
  if (!CLIENTS || !CLIENTS[clientKey]) {
    console.error("Unknown client key:", clientKey);
    return;
  }

  currentClient = { key: clientKey, ...CLIENTS[clientKey] };
  currentRange  = "last_7d";
  currentDates  = getRangeDates(currentRange);

  // Set client name in UI
  document.querySelectorAll(".client-name").forEach(el => {
    el.textContent = currentClient.name;
  });

  // Set page title
  document.title = `${currentClient.name} — Meta Ads Report | Scale Science`;

  initTokenControls();
  initDateControls();
  initEditorialBlocks();
  injectTrendsUI();
  initTabs();

  if (!getToken()) {
    showTokenScreen();
  } else {
    loadReport();
  }
}
