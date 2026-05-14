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
    // Only fields verified against Meta Marketing API v25 docs.
    // Kept narrow — Meta returns 500 "reduce the amount of data" if expansion is too heavy.
    const creativeFields = [
      "thumbnail_url",
      "image_url",
      "image_hash",
      "video_id",
      "object_type",
      "effective_object_story_id"
    ].join(",");

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
          fields: `id,name,creative{${creativeFields}}`,
          limit: 50
        });
      }
      if (data.data) all = all.concat(data.data);
      next = data.paging?.next || null;
    } while (next);

    // Pass 1: extract whatever URL we can immediately, collect unresolved hashes / story-ids / video-ids
    const map = {};
    const neededHashes = new Set();
    const neededStoryIds = new Set();
    const neededVideoIds = new Set();

    for (const ad of all) {
      const cr  = ad.creative || {};
      const objectType = (cr.object_type || "").toUpperCase();
      const isVideo = objectType === "VIDEO" || !!cr.video_id;

      const url     = cr.thumbnail_url || cr.image_url || null;
      // For thumbnail resolution fallbacks, only chase hash/video/story when no URL yet
      const hash    = !url ? cr.image_hash || null : null;
      const videoId = !url ? cr.video_id   || null : null;
      const storyId = !url ? cr.effective_object_story_id || null : null;

      map[ad.id] = {
        thumbnailUrl: url,
        isVideo,
        videoId:    cr.video_id    || null,  // kept for creative download
        imageHash:  cr.image_hash  || null,  // always stored — used to fetch full-res via /adimages
        _hash: hash, _videoId: videoId, _storyId: storyId
      };

      if (!url) {
        if (hash)    neededHashes.add(hash);
        if (videoId) neededVideoIds.add(videoId);
        if (storyId) neededStoryIds.add(storyId);
      }
    }

    // Pass 2: resolve image_hash → URL via /adimages (batched)
    if (neededHashes.size) {
      try {
        const hashes = [...neededHashes];
        const res = await apiGet(`${adAccountId}/adimages`, {
          hashes: JSON.stringify(hashes),
          fields: "hash,permalink_url,url_128"
        });
        const byHash = {};
        for (const im of (res.data || [])) {
          byHash[im.hash] = im.url_128 || im.permalink_url;
        }
        for (const ad of Object.values(map)) {
          if (!ad.thumbnailUrl && ad._hash && byHash[ad._hash]) {
            ad.thumbnailUrl = byHash[ad._hash];
          }
        }
      } catch (e) { console.warn("adimages lookup failed:", e.message); }
    }

    // Pass 3: resolve video_id → picture via batch
    if (neededVideoIds.size) {
      try {
        const ids = [...neededVideoIds];
        const res = await apiGet("", { ids: ids.join(","), fields: "picture" });
        for (const ad of Object.values(map)) {
          if (!ad.thumbnailUrl && ad._videoId && res[ad._videoId]?.picture) {
            ad.thumbnailUrl = res[ad._videoId].picture;
          }
        }
      } catch (e) { console.warn("video lookup failed:", e.message); }
    }

    // Pass 4: resolve effective_object_story_id → full_picture via batch
    if (neededStoryIds.size) {
      try {
        const ids = [...neededStoryIds];
        // Batch in chunks of 50 (Meta limit)
        const byId = {};
        for (let i = 0; i < ids.length; i += 50) {
          const chunk = ids.slice(i, i + 50);
          const res = await apiGet("", { ids: chunk.join(","), fields: "full_picture,picture" });
          Object.assign(byId, res);
        }
        for (const ad of Object.values(map)) {
          if (!ad.thumbnailUrl && ad._storyId && byId[ad._storyId]) {
            ad.thumbnailUrl = byId[ad._storyId].full_picture || byId[ad._storyId].picture;
          }
        }
      } catch (e) { console.warn("story lookup failed:", e.message); }
    }

    // Promote storyId to public field, strip resolver-only private fields
    for (const ad of Object.values(map)) {
      ad.storyId = ad._storyId || null;
      delete ad._hash; delete ad._videoId; delete ad._storyId;
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
      adId:      r.ad_id || null,
      storyId:   thumb.storyId   || null,
      videoId:   thumb.videoId   || null,
      imageHash: thumb.imageHash || null,
      name:      r.ad_name || "—",
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
              <img class="ad-thumb" src="${row.thumbnailUrl}" loading="lazy" referrerpolicy="no-referrer"
                   onerror="this.closest('.ad-thumb-video-badge').style.display='none'">
            </div>`;
          } else {
            // Image ad: plain img, hide on error
            thumbEl = `<img class="ad-thumb" src="${row.thumbnailUrl}" loading="lazy" referrerpolicy="no-referrer"
                            onerror="this.style.display='none'">`;
          }
        } else {
          thumbEl = placeholder;
        }

        const adId = row.adId ? `data-ad-id="${row.adId}" data-ad-name="${val.replace(/"/g, '&quot;')}"` : "";
        return `<div class="ad-thumb-cell ad-preview-trigger" ${adId} style="cursor:${row.adId ? 'pointer' : 'default'}">${thumbEl}<span class="ad-name-text">${val}</span></div>`;
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
    const [accInfo, accCurr, accPrev, campaigns, adsets, ads, thumbnails] = await Promise.all([
      apiGet(client.adAccountId, { fields: "currency" }),
      fetchInsights(client.adAccountId, "account",  { since: dates.since,     until: dates.until }),
      fetchInsights(client.adAccountId, "account",  { since: dates.compSince, until: dates.compUntil }),
      fetchInsights(client.adAccountId, "campaign", { since: dates.since,     until: dates.until }, ["campaign_name"]),
      fetchInsights(client.adAccountId, "adset",    { since: dates.since,     until: dates.until }, ["adset_name"]),
      fetchInsights(client.adAccountId, "ad",       { since: dates.since,     until: dates.until }, ["ad_name", "ad_id"]),
      fetchAdThumbnails(client.adAccountId)
    ]);

    // Use the currency the ad account is actually billed in
    const currency = accInfo.currency || client.currency;
    client.currency = currency;

    _accountData  = buildAccountMetrics(accCurr,    currency);
    _accountPrev  = buildAccountMetrics(accPrev,     currency);
    _campaignData = buildCampaignRows(campaigns,     currency);
    _adsetData    = buildAdSetRows(adsets,            currency);
    _adData       = buildAdRows(ads, thumbnails,      currency);

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

function fbPostUrl(storyId) {
  if (!storyId) return "";
  const idx = storyId.indexOf("_");
  if (idx < 0) return "";
  const pageId = storyId.slice(0, idx);
  const postId = storyId.slice(idx + 1);
  return `https://www.facebook.com/permalink.php?story_fbid=${postId}&id=${pageId}`;
}

function injectAdCsvButton(adEl) {
  const section = adEl.closest(".section");
  if (!section) return;
  const header = section.querySelector(".section-header");
  if (!header || header.querySelector(".btn-csv-ad")) return;

  const csvBtn = document.createElement("button");
  csvBtn.className = "btn-ghost btn-csv-ad no-print";
  csvBtn.innerHTML = `<svg width="11" height="11" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><path d="M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4M7 10l5 5 5-5M12 15V3"/></svg> Export CSV`;
  csvBtn.addEventListener("click", exportAdPreviewCSV);

  const zipBtn = document.createElement("button");
  zipBtn.className = "btn-ghost btn-creatives-zip no-print";
  zipBtn.innerHTML = `<svg width="11" height="11" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><path d="M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4M7 10l5 5 5-5M12 15V3"/></svg> Download Creatives`;
  zipBtn.addEventListener("click", downloadAdCreatives);

  header.appendChild(csvBtn);
  header.appendChild(zipBtn);
}

function exportAdPreviewCSV() {
  if (!_adData || !_adData.length) return;
  const c = currentClient?.currency || "CAD";
  const fmt = (v, fn) => v != null ? fn(v) : "";

  const headers = [
    "Ad Name",
    "Facebook Post Link",
    "Amount Spent",
    "Purchases",
    "Purchase ROAS",
    "Cost / Purchase",
    "Cost / Checkout",
    "Cost / Add to Cart",
    "Cost / Click",
    "CPM",
    "Frequency"
  ];

  const rows = _adData.map(r => [
    r.name,
    fbPostUrl(r.storyId),
    fmt(r.spend,      v => v.toFixed(2)),
    fmt(r.purchases,  v => v.toFixed(0)),
    fmt(r.roas,       v => v.toFixed(2)),
    fmt(r.cpPurchase, v => v.toFixed(2)),
    fmt(r.cpCheckout, v => v.toFixed(2)),
    fmt(r.cpCart,     v => v.toFixed(2)),
    fmt(r.cpClick,    v => v.toFixed(2)),
    fmt(r.cpm,        v => v.toFixed(2)),
    fmt(r.frequency,  v => v.toFixed(2))
  ]);

  const csvContent = [headers, ...rows]
    .map(row => row.map(cell => `"${String(cell).replace(/"/g, '""')}"`).join(","))
    .join("\n");

  const client     = currentClient?.name?.replace(/[^a-z0-9]/gi, "_") || "Client";
  const since      = currentDates?.since || "";
  const until      = currentDates?.until || "";
  const blob = new Blob([csvContent], { type: "text/csv;charset=utf-8;" });
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement("a");
  a.href = url;
  a.download = `${client}_AdPreview_${since}_${until}.csv`;
  a.click();
  URL.revokeObjectURL(url);
}

/* ── Creative zip download ─────────────────────────────────── */
async function loadJSZip() {
  if (window.JSZip) return;
  await new Promise((resolve, reject) => {
    const s = document.createElement("script");
    s.src = "https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js";
    s.onload = resolve; s.onerror = reject;
    document.head.appendChild(s);
  });
}

async function tryFetchBlob(url) {
  try {
    const res = await fetch(url);
    if (!res.ok) return null;
    return await res.blob();
  } catch { return null; }
}

async function downloadAdCreatives() {
  if (!_adData || !_adData.length) return;
  const btn = document.querySelector(".btn-creatives-zip");
  const originalLabel = btn?.innerHTML;

  try {
    if (btn) { btn.disabled = true; btn.textContent = "Loading…"; }
    await loadJSZip();
    const zip = new window.JSZip();

    const clientName = currentClient?.name?.replace(/[^a-z0-9]/gi, "_") || "Client";
    const adAccountId = currentClient?.adAccountId;
    const since       = currentDates?.since || "";
    const until       = currentDates?.until || "";
    const sanitize    = str => (str || "ad").replace(/[/\\:*?"<>|]/g, "_").trim().slice(0, 60);

    // Deduplicate by adId
    const seen = new Set();
    const ads  = _adData.filter(r => {
      if (!r.adId || seen.has(r.adId)) return false;
      seen.add(r.adId); return true;
    });

    if (btn) btn.textContent = "Resolving images…";

    // Batch-resolve image hashes → full-res URL via /adimages
    const hashToUrl = {};
    const hashes = [...new Set(ads.map(r => r.imageHash).filter(Boolean))];
    if (hashes.length && adAccountId) {
      try {
        // /adimages accepts up to 50 hashes at a time
        for (let i = 0; i < hashes.length; i += 50) {
          const chunk = hashes.slice(i, i + 50);
          const res = await apiGet(`${adAccountId}/adimages`, {
            hashes: JSON.stringify(chunk),
            fields: "hash,url"
          });
          for (const img of (res.data || [])) {
            if (img.hash && img.url) hashToUrl[img.hash] = img.url;
          }
        }
      } catch (e) { console.warn("adimages batch failed:", e.message); }
    }

    const urlLinks = [];
    let done = 0;

    for (const row of ads) {
      done++;
      if (btn) btn.textContent = `Downloading ${done} / ${ads.length}…`;
      const name = sanitize(row.name) || `ad_${done}`;

      try {
        if (row.videoId) {
          // Video: get the original MP4 source URL from Meta
          const vData = await apiGet(row.videoId, { fields: "source" });
          const sourceUrl = vData.source;
          if (sourceUrl) {
            const blob = await tryFetchBlob(sourceUrl);
            if (blob && blob.size > 10000) {
              zip.file(`${name}.mp4`, blob);
            } else {
              urlLinks.push({ name, url: sourceUrl, type: "video" });
            }
          }
        } else {
          // Image: prefer full-res from /adimages, fall back to whatever URL we have
          const imgUrl = (row.imageHash && hashToUrl[row.imageHash]) || row.thumbnailUrl;
          if (imgUrl) {
            const blob = await tryFetchBlob(imgUrl);
            const ext  = /\.png(\?|$)/i.test(imgUrl) ? "png" : /\.gif(\?|$)/i.test(imgUrl) ? "gif" : "jpg";
            if (blob && blob.size > 1000) {
              zip.file(`${name}.${ext}`, blob);
            } else {
              urlLinks.push({ name, url: imgUrl, type: "image" });
            }
          }
        }
      } catch (e) {
        console.warn(`Skipped ${row.name}:`, e.message);
      }
    }

    // Always include a gallery HTML — images/videos viewable even if binary fetch was blocked
    const galleryRows = ads.map((r, i) => {
      const n      = sanitize(r.name) || `ad_${i + 1}`;
      const thumb  = r.thumbnailUrl || "";
      const fullImg = (r.imageHash && hashToUrl[r.imageHash]) || thumb;
      const preview = r.videoId
        ? `<video controls style="max-width:100%;border-radius:6px;" poster="${thumb}">
             <source src="" data-src-video="${r.videoId}">
           </video>`
        : fullImg
          ? `<a href="${fullImg}" target="_blank"><img src="${fullImg}" referrerpolicy="no-referrer" style="max-width:100%;border-radius:6px;display:block;"></a>`
          : `<div style="color:#888;padding:20px 0">No preview available</div>`;
      return `<div style="background:#1a1a1a;border:1px solid #333;border-radius:8px;padding:16px">
        <div style="font-weight:600;margin-bottom:10px;font-size:13px;word-break:break-word">${r.name}</div>
        ${preview}
      </div>`;
    }).join("\n");

    zip.file("gallery.html", `<!DOCTYPE html><html><head><meta charset="UTF-8">
<title>${clientName} Ad Creatives ${since}–${until}</title>
<style>body{background:#0a0a0a;color:#f0ede8;font-family:sans-serif;padding:24px;margin:0}
h1{font-size:18px;margin-bottom:20px}.grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(300px,1fr));gap:16px}</style>
</head><body>
<h1>${clientName} — Ad Creatives &nbsp;<small style="color:#666;font-weight:400">${since} to ${until}</small></h1>
<div class="grid">${galleryRows}</div>
</body></html>`);

    if (urlLinks.length) {
      zip.file("_manual_download_links.txt",
        "These creatives could not be auto-downloaded (browser CORS restriction).\n" +
        "Open each URL in your browser and use Save As to download it manually.\n\n" +
        urlLinks.map(l => `[${l.type.toUpperCase()}] ${l.name}\n${l.url}`).join("\n\n")
      );
    }

    const zipBlob = await zip.generateAsync({ type: "blob", compression: "DEFLATE" });
    const dlUrl   = URL.createObjectURL(zipBlob);
    const a       = document.createElement("a");
    a.href = dlUrl;
    a.download = `${clientName}_AdCreatives_${since}_${until}.zip`;
    a.click();
    URL.revokeObjectURL(dlUrl);

  } catch (e) {
    alert(`Export failed: ${e.message}`);
  } finally {
    if (btn) { btn.disabled = false; btn.innerHTML = originalLabel; }
  }
}

function refreshTables() {
  const currency = currentClient?.currency || "CAD";

  const campEl = document.getElementById("campaign-table-container");
  if (campEl && _campaignData) campEl.innerHTML = renderCampaignTable(_campaignData, currency);

  const adsetEl = document.getElementById("adset-table-container");
  if (adsetEl && _adsetData) adsetEl.innerHTML = renderAdSetTable(_adsetData, currency);

  const adEl = document.getElementById("ad-table-container");
  if (adEl && _adData) {
    adEl.innerHTML = renderAdTable(_adData, currency);
    injectAdCsvButton(adEl);
  }

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

/* ── Ad Preview Modal ─────────────────────────────────────── */
const AD_PREVIEW_FORMATS = [
  { value: "DESKTOP_FEED_STANDARD",  label: "Facebook Feed" },
  { value: "MOBILE_FEED_STANDARD",   label: "Facebook Mobile" },
  { value: "INSTAGRAM_STANDARD",     label: "Instagram Feed" },
  { value: "INSTAGRAM_STORY",        label: "Instagram Story" },
  { value: "INSTAGRAM_REELS",        label: "Instagram Reels" },
];

function injectAdPreviewModal() {
  if (document.getElementById("ad-preview-modal")) return;
  const formatOptions = AD_PREVIEW_FORMATS.map(f =>
    `<option value="${f.value}">${f.label}</option>`
  ).join("");
  const el = document.createElement("div");
  el.id = "ad-preview-modal";
  el.innerHTML = `
    <div class="ad-preview-backdrop"></div>
    <div class="ad-preview-dialog">
      <div class="ad-preview-header">
        <span class="ad-preview-title" id="ad-preview-title"></span>
        <div class="ad-preview-controls">
          <select class="date-select" id="ad-preview-format">${formatOptions}</select>
          <button class="btn-ghost ad-preview-close" id="ad-preview-close">✕</button>
        </div>
      </div>
      <div class="ad-preview-body" id="ad-preview-body">
        <div class="ad-preview-loading">Loading preview…</div>
      </div>
    </div>`;
  document.body.appendChild(el);

  document.getElementById("ad-preview-close").addEventListener("click", closeAdPreview);
  document.getElementById("ad-preview-backdrop") ||
    el.querySelector(".ad-preview-backdrop").addEventListener("click", closeAdPreview);
  document.getElementById("ad-preview-format").addEventListener("change", () => {
    const adId = el.dataset.adId;
    if (adId) loadAdPreviewFrame(adId, document.getElementById("ad-preview-format").value);
  });

  document.addEventListener("keydown", e => {
    if (e.key === "Escape" && el.classList.contains("open")) closeAdPreview();
  });

  // Event delegation for ad-preview-trigger clicks
  document.addEventListener("click", e => {
    const trigger = e.target.closest(".ad-preview-trigger[data-ad-id]");
    if (trigger) openAdPreview(trigger.dataset.adId, trigger.dataset.adName);
  });
}

function openAdPreview(adId, adName) {
  const modal = document.getElementById("ad-preview-modal");
  if (!modal) return;
  modal.dataset.adId = adId;
  document.getElementById("ad-preview-title").textContent = adName || "Ad Preview";
  document.getElementById("ad-preview-format").value = AD_PREVIEW_FORMATS[0].value;
  modal.classList.add("open");
  document.body.style.overflow = "hidden";
  loadAdPreviewFrame(adId, AD_PREVIEW_FORMATS[0].value);
}

function closeAdPreview() {
  const modal = document.getElementById("ad-preview-modal");
  if (!modal) return;
  modal.classList.remove("open");
  document.body.style.overflow = "";
  document.getElementById("ad-preview-body").innerHTML =
    `<div class="ad-preview-loading">Loading preview…</div>`;
}

async function loadAdPreviewFrame(adId, format) {
  const body = document.getElementById("ad-preview-body");
  body.innerHTML = `<div class="ad-preview-loading">Loading preview…</div>`;
  try {
    const data = await apiGet(`${adId}/previews`, { ad_format: format });
    const html = data.data?.[0]?.body;
    if (!html) { body.innerHTML = `<div class="ad-preview-loading">No preview available for this format.</div>`; return; }
    // Meta returns an escaped iframe string — unescape it
    const unescaped = html.replace(/&lt;/g,"<").replace(/&gt;/g,">").replace(/&amp;/g,"&").replace(/&quot;/g,'"');
    body.innerHTML = `<div class="ad-preview-frame">${unescaped}</div>`;
  } catch(e) {
    body.innerHTML = `<div class="ad-preview-loading" style="color:var(--red)">Failed to load preview: ${e.message}</div>`;
  }
}

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
    </button>
    <button class="tab-btn" data-tab="analysis">
      <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2">
        <rect x="3" y="3" width="18" height="18" rx="2"/>
        <line x1="3" y1="9" x2="21" y2="9"/>
        <line x1="9" y1="9" x2="9" y2="21"/>
      </svg>
      Analysis
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

  // Analysis content (after trends, inside <main>)
  const analysisDiv         = document.createElement("div");
  analysisDiv.id            = "analysis-content";
  analysisDiv.className     = "no-print";
  analysisDiv.style.display = "none";
  analysisDiv.innerHTML     = `
    <div class="trends-header">
      <div class="trends-title-row">
        <div class="section-label">Performance Analysis</div>
        <div class="trends-toggle" data-toggle-group="analysis">
          <button class="toggle-btn"       data-analysis-mode="weekly">Weekly</button>
          <button class="toggle-btn active" data-analysis-mode="monthly">Monthly</button>
          <button class="toggle-btn"       data-analysis-mode="quarterly">Quarterly</button>
        </div>
      </div>
    </div>
    <div id="analysis-table-wrap"></div>`;
  trendsDiv.after(analysisDiv);
}

/* ── Tab + toggle event handling ──────────────────────────── */
function initTabs() {
  document.addEventListener("click", e => {
    // Main tabs (Report / Trends / Analysis)
    const tabBtn = e.target.closest(".tab-btn[data-tab]");
    if (tabBtn) {
      const tab = tabBtn.dataset.tab;
      document.querySelectorAll(".tab-btn[data-tab]").forEach(b =>
        b.classList.toggle("active", b === tabBtn));
      document.getElementById("report-content").style.display =
        tab === "report" ? "" : "none";
      const tc = document.getElementById("trends-content");
      if (tc) tc.style.display = tab === "trends" ? "" : "none";
      const ac = document.getElementById("analysis-content");
      if (ac) ac.style.display = tab === "analysis" ? "" : "none";

      if (tab === "trends"   && !_trendsCache[_trendsMode])     loadTrends(_trendsMode);
      if (tab === "analysis" && !_analysisCache[_analysisMode]) loadAnalysis(_analysisMode);
      return;
    }

    // Trends weekly/monthly toggle
    const trendsToggle = e.target.closest(".toggle-btn[data-mode]");
    if (trendsToggle) {
      const mode = trendsToggle.dataset.mode;
      _trendsMode = mode;
      document.querySelectorAll(".toggle-btn[data-mode]").forEach(b =>
        b.classList.toggle("active", b === trendsToggle));
      if (_trendsCache[mode]) renderTrendsFromCache();
      else loadTrends(mode);
      return;
    }

    // Analysis weekly/monthly/quarterly toggle
    const analysisToggle = e.target.closest(".toggle-btn[data-analysis-mode]");
    if (analysisToggle) {
      const mode = analysisToggle.dataset.analysisMode;
      _analysisMode = mode;
      document.querySelectorAll(".toggle-btn[data-analysis-mode]").forEach(b =>
        b.classList.toggle("active", b === analysisToggle));
      if (_analysisCache[mode]) renderAnalysisFromCache();
      else loadAnalysis(mode);
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

      const days = Math.max(1, Math.round(
        (parseDateStr(r.date_stop) - parseDateStr(r.date_start)) / 86400000
      ) + 1);

      return {
        dateStart:     r.date_start,
        dateStop:      r.date_stop,
        spend,
        avgDailySpend: spend / days,
        revenue,
        roas:          spend > 0 ? revenue / spend : 0,
        cpa:           purchases > 0 ? spend / purchases : null,
        ctr:           impressions > 0 ? (outClicks / impressions) * 100 : 0,
        cpm:           parseFloat(r.cpm || 0),
        frequency:     parseFloat(r.frequency || 0),
        cvr:           outClicks > 0 ? (purchases / outClicks) * 100 : 0
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

/* ── Trends cell colour helper ────────────────────────────── */
function trendCellBg(value, min, max, lowerIsBetter) {
  if (value == null || min == null || max == null || min === max) return "";
  const range = max - min;
  const ratio = (value - min) / range; // 0 = worst end, 1 = best end
  const goodRatio = lowerIsBetter ? 1 - ratio : ratio;
  // Only colour cells clearly in the top or bottom third
  if (goodRatio >= 0.67) {
    const alpha = 0.06 + (goodRatio - 0.67) / 0.33 * 0.14; // 0.06–0.20
    return `background-color:rgba(0,255,106,${alpha.toFixed(3)});`;
  } else if (goodRatio <= 0.33) {
    const alpha = 0.06 + (0.33 - goodRatio) / 0.33 * 0.14;
    return `background-color:rgba(255,59,59,${alpha.toFixed(3)});`;
  }
  return "";
}

/* ── Render trends table HTML ─────────────────────────────── */
function renderTrendsTable(rows, mode) {
  const c = currentClient?.currency || "CAD";
  if (!rows.length) return `<div class="table-empty">No data for this period.</div>`;

  const cols = [
    { label: "Period",           key: "period",    num: false },
    { label: "Amount Spent",     key: "spend",         num: true,  fmt: v => formatCurrency(v, c),          neutral: true },
    { label: "Avg Daily Spend",  key: "avgDailySpend", num: true,  fmt: v => formatCurrency(v, c),          neutral: true },
    { label: "Revenue",          key: "revenue",       num: true,  fmt: v => formatCurrency(v, c),          lowerIsBetter: false },
    { label: "ROAS",             key: "roas",      num: true,  fmt: formatRoas,                         lowerIsBetter: false },
    { label: "Cost / Purchase",  key: "cpa",       num: true,  fmt: v => v != null ? formatCurrency(v, c) : "—", lowerIsBetter: true },
    { label: "Outbound CTR",     key: "ctr",       num: true,  fmt: formatPct,                          lowerIsBetter: false },
    { label: "CPM",              key: "cpm",       num: true,  fmt: v => formatCurrency(v, c),          lowerIsBetter: true },
    { label: "Frequency",        key: "frequency", num: true,  fmt: v => formatNum(v, 2),               lowerIsBetter: true },
    { label: "Conv. Rate",       key: "cvr",       num: true,  fmt: formatPct,                          lowerIsBetter: false }
  ];

  // Pre-compute per-column min/max for numeric, non-neutral columns
  const colRanges = {};
  cols.forEach(col => {
    if (!col.num || col.neutral || col.key === "period") return;
    const vals = rows.map(r => r[col.key]).filter(v => v != null && isFinite(v));
    if (vals.length < 2) return;
    colRanges[col.key] = { min: Math.min(...vals), max: Math.max(...vals) };
  });

  const headers = cols.map(col =>
    `<th class="${col.num ? "th-num" : ""}">${col.label}</th>`
  ).join("");

  const bodyRows = rows.map(row => {
    const cells = cols.map(col => {
      if (col.key === "period") {
        return `<td class="td-name trends-period">${formatTrendsPeriod(row.dateStart, row.dateStop, mode)}</td>`;
      }
      const range = colRanges[col.key];
      const bg = range ? trendCellBg(row[col.key], range.min, range.max, col.lowerIsBetter) : "";
      return `<td class="td-num"${bg ? ` style="${bg}"` : ""}>${col.fmt(row[col.key])}</td>`;
    }).join("");
    return `<tr>${cells}</tr>`;
  }).join("");

  // Totals / averages
  const totalSpend   = rows.reduce((s, r) => s + r.spend, 0);
  const totalRevenue = rows.reduce((s, r) => s + r.revenue, 0);
  const totalPurch   = rows.reduce((s, r) => s + (r.cpa != null ? r.spend / r.cpa : 0), 0);
  const foot = {
    spend:         totalSpend,
    avgDailySpend: rows.reduce((s, r) => s + r.avgDailySpend, 0) / (rows.length || 1),
    revenue:       totalRevenue,
    roas:      totalSpend > 0 ? totalRevenue / totalSpend : 0,
    cpa:       totalPurch > 0 ? totalSpend / totalPurch : null,
    ctr:       rows.reduce((s, r) => s + r.ctr, 0)       / (rows.length || 1),
    cpm:       rows.reduce((s, r) => s + r.cpm, 0)       / (rows.length || 1),
    frequency: rows.reduce((s, r) => s + r.frequency, 0) / (rows.length || 1),
    cvr:       rows.reduce((s, r) => s + r.cvr, 0)       / (rows.length || 1)
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

/* =========================================================
   Performance Analysis Tab
   ========================================================= */

let _analysisMode  = "monthly";
let _analysisCache = {}; // { quarterly: cols, monthly: cols, weekly: cols }

const TARGET_ROAS_DEFAULTS = { root: 2.9, toothpod: 1.0, mycosoul: 1.0 };

function getTargetRoas() {
  const key = `analysis_target_roas::${currentClient?.key || "default"}`;
  const stored = localStorage.getItem(key);
  if (stored != null && !isNaN(parseFloat(stored))) return parseFloat(stored);
  return TARGET_ROAS_DEFAULTS[currentClient?.key] ?? 1.0;
}
function setTargetRoas(value) {
  const key = `analysis_target_roas::${currentClient?.key || "default"}`;
  localStorage.setItem(key, String(value));
}

/* ── Period generation ────────────────────────────────────── */
function getAnalysisPeriods(mode) {
  const periods = [];
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const monthShort = d => d.toLocaleDateString("en-CA", { month: "short" });

  if (mode === "quarterly") {
    // Current quarter (partial) + last 4 complete quarters
    const curQ = Math.floor(today.getMonth() / 3);
    const curY = today.getFullYear();
    const curQStart = new Date(curY, curQ * 3, 1);
    periods.push({
      since: formatDate(curQStart),
      until: formatDate(today),
      label: `${monthShort(curQStart)} - ${monthShort(today)} ${today.getDate()}`,
      key:   `cur-q-${curY}-${curQ}`
    });
    for (let i = 1; i <= 4; i++) {
      let q = curQ - i, y = curY;
      while (q < 0) { q += 4; y -= 1; }
      const qStart = new Date(y, q * 3, 1);
      const qEnd   = new Date(y, q * 3 + 3, 0);
      periods.push({
        since: formatDate(qStart),
        until: formatDate(qEnd),
        label: `Q${q + 1} '${String(y).slice(2)}`,
        key:   `q-${y}-${q}`
      });
    }
  } else if (mode === "monthly") {
    // Current month (partial) + last 4 complete months
    const curM = today.getMonth();
    const curY = today.getFullYear();
    const curMStart = new Date(curY, curM, 1);
    periods.push({
      since: formatDate(curMStart),
      until: formatDate(today),
      label: `${monthShort(curMStart)} 1 - ${today.getDate()}`,
      key:   `cur-m-${curY}-${curM}`
    });
    for (let i = 1; i <= 4; i++) {
      let m = curM - i, y = curY;
      while (m < 0) { m += 12; y -= 1; }
      const mStart = new Date(y, m, 1);
      const mEnd   = new Date(y, m + 1, 0);
      periods.push({
        since: formatDate(mStart),
        until: formatDate(mEnd),
        label: `${monthShort(mStart)} '${String(y).slice(2)}`,
        key:   `m-${y}-${m}`
      });
    }
  } else { // weekly: current week (partial) + last 4 complete weeks
    const dow = today.getDay() === 0 ? 7 : today.getDay(); // Monday=1..Sunday=7
    const curWStart = new Date(today);
    curWStart.setDate(today.getDate() - (dow - 1));
    periods.push({
      since: formatDate(curWStart),
      until: formatDate(today),
      label: `Wk ${monthShort(curWStart)} ${curWStart.getDate()}`,
      key:   `cur-w-${formatDate(curWStart)}`
    });
    for (let i = 1; i <= 4; i++) {
      const wStart = new Date(curWStart);
      wStart.setDate(curWStart.getDate() - 7 * i);
      const wEnd = new Date(wStart);
      wEnd.setDate(wStart.getDate() + 6);
      periods.push({
        since: formatDate(wStart),
        until: formatDate(wEnd),
        label: `Wk ${monthShort(wStart)} ${wStart.getDate()}`,
        key:   `w-${formatDate(wStart)}`
      });
    }
  }
  return periods;
}

/* ── Action / video helpers ───────────────────────────────── */
function sumActionArray(arr) {
  if (!Array.isArray(arr)) return 0;
  return arr.reduce((s, x) => s + parseFloat(x.value || 0), 0);
}
function getPurchaseAction(row) {
  // Prefer fb_pixel_purchase, fall back to omni_purchase
  const arr = row?.actions || [];
  return arr.find(x => x.action_type === PURCHASE_ACTION)
      || arr.find(x => x.action_type === "omni_purchase")
      || arr.find(x => x.action_type === "purchase");
}
function getViewThroughPurchases(row) {
  const a = getPurchaseAction(row);
  if (!a) return 0;
  return parseFloat(a["1d_view"] || 0);
}

/* ── Aggregation helpers ──────────────────────────────────── */
function sumField(rows, field) {
  return rows.reduce((s, r) => s + parseFloat(r[field] || 0), 0);
}
function sumActionsByType(rows, field) {
  // Preserve all per-action fields (value + attribution-window breakdowns like 1d_view)
  const byType = {};
  for (const r of rows) {
    for (const a of (r[field] || [])) {
      if (!byType[a.action_type]) byType[a.action_type] = {};
      for (const [k, v] of Object.entries(a)) {
        if (k === "action_type") continue;
        byType[a.action_type][k] = (byType[a.action_type][k] || 0) + parseFloat(v || 0);
      }
    }
  }
  return Object.entries(byType).map(([action_type, vals]) => {
    const obj = { action_type };
    for (const [k, v] of Object.entries(vals)) obj[k] = String(v);
    return obj;
  });
}
function weightedAvgActionsByType(rows, field, weightField) {
  // weighted by the per-period weightField (e.g. video_view count)
  const sumVal = {}, sumWt = {};
  for (const r of rows) {
    const w = parseFloat(r[weightField] || 0) || 1;
    for (const a of (r[field] || [])) {
      sumVal[a.action_type] = (sumVal[a.action_type] || 0) + parseFloat(a.value || 0) * w;
      sumWt [a.action_type] = (sumWt [a.action_type] || 0) + w;
    }
  }
  return Object.entries(sumVal).map(([action_type, totalNum]) => ({
    action_type, value: String(totalNum / (sumWt[action_type] || 1))
  }));
}

function aggregateRows(rows) {
  if (!rows.length) return {};
  if (rows.length === 1) return rows[0];

  const spend       = sumField(rows, "spend");
  const impressions = sumField(rows, "impressions");
  const reach       = sumField(rows, "reach"); // sum is an over-estimate for multi-period reach; close enough
  const linkClicks  = sumField(rows, "inline_link_clicks");
  const outClicks   = rows.reduce((s, r) => s + getOutboundClicks(r), 0);

  const actions      = sumActionsByType(rows, "actions");
  const actionVals   = sumActionsByType(rows, "action_values");
  const thruplays    = sumActionsByType(rows, "video_thruplay_watched_actions");

  // Look up purchase value from aggregated action_values to recompute ROAS
  const purchaseVal  = parseFloat(
    actionVals.find(a => a.action_type === PURCHASE_ACTION)?.value
    || actionVals.find(a => a.action_type === "omni_purchase")?.value
    || 0
  );

  // For weighted avg watch time, weight by per-period video_view count from actions array
  const videoViewByRow = rows.map(r =>
    (r.actions || []).find(a => a.action_type === "video_view")?.value || 0
  );
  const rowsWithVw = rows.map((r, i) => ({ ...r, _vw: parseFloat(videoViewByRow[i]) || 0 }));
  const avgWatch   = weightedAvgActionsByType(rowsWithVw, "video_avg_time_watched_actions", "_vw");

  // Re-compute cost_per_action_type for the purchase/cart actions we care about
  const cpat = actions.map(a => {
    const cnt = parseFloat(a.value || 0);
    return { action_type: a.action_type, value: cnt > 0 ? String(spend / cnt) : "0" };
  });

  return {
    spend:                              String(spend),
    impressions:                        String(impressions),
    reach:                              String(reach),
    cpm:                                String(impressions > 0 ? (spend / impressions) * 1000 : 0),
    frequency:                          String(reach > 0 ? impressions / reach : 0),
    inline_link_clicks:                 String(linkClicks),
    cost_per_inline_link_click:         String(linkClicks > 0 ? spend / linkClicks : 0),
    inline_link_click_ctr:              String(impressions > 0 ? (linkClicks / impressions) * 100 : 0),
    outbound_clicks:                    String(outClicks),
    actions,
    action_values:                      actionVals,
    cost_per_action_type:               cpat,
    video_thruplay_watched_actions:     thruplays,
    video_avg_time_watched_actions:     avgWatch,
    purchase_roas: [{ action_type: "omni_purchase", value: String(spend > 0 ? purchaseVal / spend : 0) }]
  };
}

function rowsWithin(rows, since, until) {
  return rows.filter(r => r.date_start >= since && r.date_start <= until);
}

/* ── Fetch insights — ONE total API call ──────────────────── */
async function fetchAnalysisInsights(mode) {
  const periods = getAnalysisPeriods(mode);
  // Periods are currently sorted current-first; we need the overall date range
  const allSince = periods[periods.length - 1].since; // oldest period start
  const allUntil = periods[0].until;                   // newest (today)
  const incr     = mode === "weekly" ? 7 : "monthly"; // quarterly aggregates from monthly rows

  const fields = [
    "spend", "impressions", "reach", "cpm", "frequency",
    "actions", "action_values", "cost_per_action_type",
    "outbound_clicks", "inline_link_clicks",
    "cost_per_inline_link_click", "inline_link_click_ctr",
    "video_thruplay_watched_actions",
    "video_avg_time_watched_actions",
    "purchase_roas",
    "date_start", "date_stop"
  ].join(",");

  // Single call requesting both default and 1d_view attribution.
  // Meta returns the actions array with a `value` (= default) and a separate
  // `1d_view` field per action, so we can derive % view conversion in one pass.
  const res = await apiGet(`${currentClient.adAccountId}/insights`, {
    fields,
    time_range:                 JSON.stringify({ since: allSince, until: allUntil }),
    time_increment:             incr,
    action_attribution_windows: JSON.stringify(["default", "1d_view"]),
    level:                      "account",
    limit:                      100
  });

  const rows = res.data || [];
  return periods.map(p => ({
    period: p,
    row:    aggregateRows(rowsWithin(rows, p.since, p.until))
  }));
}

/* ── Compute metrics for a single period ──────────────────── */
function computePeriodMetrics(period, row) {
  const spend       = parseFloat(row.spend || 0);
  const impressions = parseFloat(row.impressions || 0);
  const reach       = parseFloat(row.reach || 0);
  const cpm         = parseFloat(row.cpm || 0);
  const freq        = parseFloat(row.frequency || (reach > 0 ? impressions / reach : 0));
  const cpc         = parseFloat(row.cost_per_inline_link_click || 0);
  const ctr         = parseFloat(row.inline_link_click_ctr || 0);

  const purchase    = getPurchaseAction(row);
  const purchases   = purchase ? parseFloat(purchase.value || 0) : 0;
  const purchaseVal = getActionValue(row, PURCHASE_ACTION) || getActionValue(row, "omni_purchase");
  const cpa         = purchases > 0 ? spend / purchases : null;
  const roas        = parseRoas(row) ?? (spend > 0 ? purchaseVal / spend : 0);

  const addToCart   = getAction(row, "offsite_conversion.fb_pixel_add_to_cart") || getAction(row, "add_to_cart");
  const cATC        = addToCart > 0 ? spend / addToCart : null;

  const outClicks   = getOutboundClicks(row);
  const cr          = outClicks > 0 ? (purchases / outClicks) * 100 : 0;
  const aov         = purchases > 0 ? purchaseVal / purchases : null;

  // Meta deprecated `video_3_sec_watched_actions` — 3-second plays now live in the
  // `actions` array as action_type=video_view (Meta defines a video view as ≥3s)
  const v3sec       = getAction(row, "video_view");
  const thruplays   = sumActionArray(row.video_thruplay_watched_actions);
  const avgWatch    = sumActionArray(row.video_avg_time_watched_actions); // already an avg in seconds
  const thumbStop   = impressions > 0 ? (v3sec / impressions) * 100 : 0;
  const holdRate    = v3sec       > 0 ? (thruplays / v3sec) * 100 : 0;

  // % view conversion: view-only purchases / total purchases.
  // Total comes from purchase.value (default attribution). View-only comes from
  // the purchase action's `1d_view` breakdown field (returned when the request
  // includes action_attribution_windows=['default','1d_view']).
  const viewPurch   = purchase ? parseFloat(purchase["1d_view"] || 0) : 0;
  const viewPct     = purchases > 0 ? Math.min(100, (viewPurch / purchases) * 100) : 0;

  // Days in period — for "daily ad spent"
  const dStart = parseDateStr(period.since);
  const dStop  = parseDateStr(period.until);
  const days   = Math.max(1, Math.round((dStop - dStart) / 86400000) + 1);

  return {
    period,
    dailySpend:   spend / days,
    spend,
    thumbStop, holdRate, avgWatch,
    cpm, cpc, ctr,
    cATC, purchases, cpa, purchaseVal, roas,
    aov, cr, freq, viewPct
  };
}

/* ── Cell colour: row-level heatmap ───────────────────────── */
function analysisCellBg(value, min, max, lowerIsBetter) {
  if (value == null || min == null || max == null || min === max) return "";
  const range = max - min;
  const ratio = (value - min) / range; // 0=worst, 1=best when higher-is-better
  const good  = lowerIsBetter ? 1 - ratio : ratio;
  // Brighter scale than trends — closer to the spreadsheet look
  if (good >= 0.5) {
    const alpha = 0.10 + (good - 0.5) * 0.50; // 0.10–0.35
    return `background-color:rgba(0,255,106,${alpha.toFixed(3)});`;
  } else {
    const alpha = 0.10 + (0.5 - good) * 0.50;
    return `background-color:rgba(255,82,82,${alpha.toFixed(3)});`;
  }
}

/* ── Metric row definitions ───────────────────────────────── */
function getAnalysisMetricRows() {
  const c = currentClient?.currency || "CAD";
  const sec = (v) => {
    if (!v) return "0:00";
    const s = Math.round(v);
    return `${Math.floor(s / 60)}:${String(s % 60).padStart(2, "0")}`;
  };
  return [
    { label: "Daily Ad spent",   key: "dailySpend",  fmt: v => formatCurrency(v, c),         direction: "neutral" },
    { label: "Ad spent",         key: "spend",       fmt: v => formatCurrency(v, c),         direction: "neutral" },
    { label: "thumb-stop ratio", key: "thumbStop",   fmt: v => formatPct(v),                 direction: "higher" },
    { label: "hold rate (thru)", key: "holdRate",    fmt: v => formatPct(v),                 direction: "higher" },
    { label: "average watch time", key: "avgWatch",  fmt: sec,                               direction: "higher" },
    { label: "Average CPM",      key: "cpm",         fmt: v => formatCurrency(v, c),         direction: "lower"  },
    { label: "Average CPC",      key: "cpc",         fmt: v => formatCurrency(v, c),         direction: "lower"  },
    { label: "Average CTR",      key: "ctr",         fmt: v => formatPct(v),                 direction: "higher" },
    { label: "cATC",             key: "cATC",        fmt: v => v != null ? formatCurrency(v, c) : "—", direction: "lower"  },
    { label: "purchases",        key: "purchases",   fmt: v => formatNum(v, 0),              direction: "higher" },
    { label: "cpa",              key: "cpa",         fmt: v => v != null ? formatCurrency(v, c) : "—", direction: "lower"  },
    { label: "purchase value",   key: "purchaseVal", fmt: v => formatCurrency(v, c),         direction: "higher" },
    { label: "ROAS current (FB)",key: "roas",        fmt: v => formatRoas(v),                direction: "higher" },
    { label: "target ROAS FB",   key: "__target",                                            direction: "target" },
    { label: "AOV",              key: "aov",         fmt: v => v != null ? formatCurrency(v, c) : "—", direction: "higher" },
    { label: "CR",               key: "cr",          fmt: v => formatPct(v),                 direction: "higher" },
    { label: "Frequency",        key: "freq",        fmt: v => formatNum(v, 2),              direction: "lower"  },
    { label: "% view conversion",key: "viewPct",     fmt: v => formatPct(v),                 direction: "higher" },
  ];
}

/* ── Render the Analysis table ────────────────────────────── */
function renderAnalysisTable(cols, mode) {
  if (!cols.length) return `<div class="table-empty">No data.</div>`;
  const metrics = getAnalysisMetricRows();
  const target  = getTargetRoas();

  // Header row
  const headerCells = cols.map(c => `<th class="th-num">${c.period.label}</th>`).join("");

  // Per-metric row with heatmap colours
  const bodyRows = metrics.map(m => {
    if (m.direction === "target") {
      // Special row: editable target ROAS input in the first data column, others blank
      const inputCell = `<td class="td-num analysis-target">
        <input type="number" step="0.1" min="0" id="analysis-target-roas" value="${target}" />
      </td>`;
      const blankCells = cols.slice(1).map(() => `<td class="td-num"></td>`).join("");
      return `<tr class="analysis-row-target">
        <td class="td-name">${m.label}</td>
        ${inputCell}${blankCells}
      </tr>`;
    }

    const vals = cols.map(c => c.metrics[m.key]).filter(v => v != null && isFinite(v));
    let mn = null, mx = null;
    if (m.direction !== "neutral" && vals.length >= 2) {
      mn = Math.min(...vals);
      mx = Math.max(...vals);
    }

    const cells = cols.map(c => {
      const v   = c.metrics[m.key];
      const bg  = (m.direction === "higher" || m.direction === "lower")
                  ? analysisCellBg(v, mn, mx, m.direction === "lower")
                  : "";
      const txt = (v == null || (typeof v === "number" && !isFinite(v))) ? "—" : m.fmt(v);
      return `<td class="td-num"${bg ? ` style="${bg}"` : ""}>${txt}</td>`;
    }).join("");

    return `<tr><td class="td-name">${m.label}</td>${cells}</tr>`;
  }).join("");

  return `
    <div class="table-scroll">
      <table class="data-table analysis-table">
        <thead><tr><th class="th-name">Metric</th>${headerCells}</tr></thead>
        <tbody>${bodyRows}</tbody>
      </table>
    </div>`;
}

/* ── Orchestration ────────────────────────────────────────── */
async function loadAnalysis(mode) {
  const wrap = document.getElementById("analysis-table-wrap");
  if (!wrap) return;
  wrap.innerHTML = `<div class="table-wrapper">
    <div class="skeleton skeleton-table" style="height:480px;border-radius:0;"></div>
  </div>`;
  try {
    const raw  = await fetchAnalysisInsights(mode);
    const cols = raw.map(({ period, row }) => ({
      period,
      metrics: computePeriodMetrics(period, row)
    }));
    _analysisCache[mode] = cols;
    renderAnalysisFromCache();
  } catch (err) {
    console.error("Analysis fetch failed:", err);
    wrap.innerHTML = `<div class="table-wrapper" style="padding:32px;text-align:center;">
      <p style="color:var(--red);font-size:13px;">${err.message}</p>
    </div>`;
  }
}

function renderAnalysisFromCache() {
  const wrap = document.getElementById("analysis-table-wrap");
  if (!wrap) return;
  const cols = _analysisCache[_analysisMode] || [];
  wrap.innerHTML = `<div class="table-wrapper">${renderAnalysisTable(cols, _analysisMode)}</div>`;
  // Wire target ROAS input
  const input = document.getElementById("analysis-target-roas");
  if (input) {
    input.addEventListener("change", e => {
      const v = parseFloat(e.target.value);
      if (!isNaN(v) && v >= 0) setTargetRoas(v);
    });
  }
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
  injectAdPreviewModal();
  initTabs();

  if (!getToken()) {
    showTokenScreen();
  } else {
    loadReport();
  }
}
