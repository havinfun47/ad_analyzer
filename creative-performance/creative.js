/* ============================================================
   Scale Science — Creative Performance (Phase 1)
   Motion-style visual creative analytics for 3 client accounts
   ============================================================ */

const API_BASE   = "https://graph.facebook.com/v25.0";
const TOKEN_KEY  = "meta_access_token";
const CACHE_KEY  = "cp_cache_v1";
const CACHE_TTL  = 15 * 60 * 1000; // 15 minutes

const PURCHASE_ACTION = "offsite_conversion.fb_pixel_purchase";
const OMNI_PURCHASE   = "omni_purchase";

const CLIENTS = {
  anvytech: { name: "Anvy Tech", adAccountId: "act_575199276244807" },
  toothpod: { name: "Toothpod",  adAccountId: "act_727374130071249" },
  mycosoul: { name: "myco:soul", adAccountId: "act_521928957251738" }
};

/* ── State ─────────────────────────────────────────────────── */
const state = {
  clientKey: "anvytech",
  range:     "last_14d",
  status:    "all",
  format:    "all",
  minSpend:  0,
  sortBy:    "spend",
  sortDir:   "desc",
  currency:  "CAD",
  rawAds:    [],   // merged ad+insight rows
  lastFetched: null
};

/* ── Token ─────────────────────────────────────────────────── */
const getToken = () => localStorage.getItem(TOKEN_KEY);
const setToken = (t) => localStorage.setItem(TOKEN_KEY, t);
const clearToken = () => localStorage.removeItem(TOKEN_KEY);

/* ── Date range helpers ────────────────────────────────────── */
function pad(n) { return String(n).padStart(2, "0"); }
function ymd(d) { return `${d.getFullYear()}-${pad(d.getMonth()+1)}-${pad(d.getDate())}`; }

function rangeToDates(range) {
  const now = new Date();
  const until = ymd(now);
  if (range === "lifetime") return { since: "2015-01-01", until };
  const days = { last_7d: 7, last_14d: 14, last_30d: 30, last_90d: 90 }[range] || 14;
  const start = new Date(now);
  start.setDate(start.getDate() - (days - 1));
  return { since: ymd(start), until };
}

/* ── API wrapper ───────────────────────────────────────────── */
async function api(path, params = {}) {
  const token = getToken();
  if (!token) throw new Error("No access token");
  const url = new URL(`${API_BASE}/${path}`);
  url.searchParams.set("access_token", token);
  for (const [k, v] of Object.entries(params)) {
    url.searchParams.set(k, typeof v === "object" ? JSON.stringify(v) : v);
  }
  const res = await fetch(url.toString());
  const data = await res.json();
  if (data.error) throw new Error(data.error.message || "Meta API error");
  return data;
}

async function apiPaginated(path, params) {
  let all = [];
  let next = null;
  do {
    let data;
    if (next) {
      const res = await fetch(next);
      data = await res.json();
      if (data.error) throw new Error(data.error.message);
    } else {
      data = await api(path, params);
    }
    if (data.data) all = all.concat(data.data);
    next = data.paging?.next || null;
  } while (next);
  return all;
}

/* ── Cache (localStorage) ──────────────────────────────────── */
function cacheKey(accountId, since, until) {
  return `${CACHE_KEY}::${accountId}::${since}::${until}`;
}
function readCache(key) {
  try {
    const raw = localStorage.getItem(key);
    if (!raw) return null;
    const entry = JSON.parse(raw);
    if (Date.now() - entry.timestamp > CACHE_TTL) return null;
    return entry;
  } catch { return null; }
}
function writeCache(key, payload) {
  try {
    localStorage.setItem(key, JSON.stringify({ timestamp: Date.now(), ...payload }));
  } catch { /* quota exceeded — ignore */ }
}
function bustCache(key) { localStorage.removeItem(key); }

/* ── Formatters ────────────────────────────────────────────── */
function formatCurrency(val, currency) {
  return new Intl.NumberFormat("en-CA", {
    style: "currency", currency: currency || state.currency,
    minimumFractionDigits: val >= 1000 ? 0 : 2, maximumFractionDigits: 2
  }).format(val || 0);
}
function formatNum(val, decimals = 0) {
  return new Intl.NumberFormat("en-CA", {
    minimumFractionDigits: decimals, maximumFractionDigits: decimals
  }).format(val || 0);
}
function formatPct(val) { return `${(val || 0).toFixed(2)}%`; }
function formatRoas(val) { return val != null ? `${val.toFixed(2)}×` : "—"; }
function relativeTime(ts) {
  if (!ts) return "Never";
  const diff = Math.floor((Date.now() - ts) / 1000);
  if (diff < 5)   return "Just now";
  if (diff < 60)  return `${diff}s ago`;
  if (diff < 3600) return `${Math.floor(diff/60)}m ago`;
  return `${Math.floor(diff/3600)}h ago`;
}

/* ── Action / insight extractors ───────────────────────────── */
function getAction(row, type) {
  const arr = row?.actions || [];
  const a = arr.find(x => x.action_type === type);
  return a ? parseFloat(a.value || 0) : 0;
}
function getActionValue(row, type) {
  const arr = row?.action_values || [];
  const a = arr.find(x => x.action_type === type);
  return a ? parseFloat(a.value || 0) : 0;
}
function getPurchases(row) {
  return getAction(row, PURCHASE_ACTION) || getAction(row, OMNI_PURCHASE);
}
function getRevenue(row) {
  return getActionValue(row, PURCHASE_ACTION) || getActionValue(row, OMNI_PURCHASE);
}
function parseRoas(row) {
  if (!row.purchase_roas) return null;
  if (Array.isArray(row.purchase_roas) && row.purchase_roas[0]) {
    return parseFloat(row.purchase_roas[0].value);
  }
  return null;
}

/* ── Creative type detection ───────────────────────────────── */
function detectFormat(creative) {
  if (!creative) return "image";
  if (creative.asset_feed_spec) return "dynamic";
  const oss = creative.object_story_spec || {};
  if (oss.link_data?.child_attachments?.length > 0) return "carousel";
  if (creative.video_id || (creative.object_type || "").toUpperCase() === "VIDEO") return "video";
  return "image";
}

function pickThumbnail(creative) {
  if (!creative) return null;
  return creative.thumbnail_url
      || creative.image_url
      || creative.object_story_spec?.video_data?.image_url
      || creative.object_story_spec?.link_data?.picture
      || creative.object_story_spec?.link_data?.child_attachments?.[0]?.picture
      || null;
}

/* ── Data fetch ────────────────────────────────────────────── */
async function fetchAccountCurrency(adAccountId) {
  try {
    const info = await api(adAccountId, { fields: "currency" });
    return info.currency || "CAD";
  } catch { return "CAD"; }
}

async function fetchAdsMetadata(adAccountId) {
  // Light creative fields only — heavy expansion triggers Meta's "reduce data" 500
  const creativeFields = [
    "thumbnail_url",
    "image_url",
    "video_id",
    "object_type"
  ].join(",");
  return apiPaginated(`${adAccountId}/ads`, {
    fields: `id,name,effective_status,created_time,creative{${creativeFields}}`,
    filtering: [{ field: "ad.effective_status", operator: "IN", value: ["ACTIVE", "PAUSED"] }],
    limit: 50
  });
}

async function fetchInsights(adAccountId, since, until) {
  const fields = [
    "ad_id", "spend", "impressions", "clicks", "ctr", "cpc", "cpm",
    "actions", "action_values", "purchase_roas", "cost_per_action_type"
  ].join(",");
  return apiPaginated(`${adAccountId}/insights`, {
    level: "ad",
    time_range: { since, until },
    fields,
    limit: 500
  });
}

/* ── Merge ads + insights ──────────────────────────────────── */
function mergeAdData(ads, insights) {
  const insightsById = {};
  for (const i of insights) insightsById[i.ad_id] = i;

  return ads.map(a => {
    const ins = insightsById[a.id] || {};
    const spend     = parseFloat(ins.spend || 0);
    const impr      = parseFloat(ins.impressions || 0);
    const clicks    = parseFloat(ins.clicks || 0);
    const ctr       = parseFloat(ins.ctr || 0);
    const cpc       = parseFloat(ins.cpc || 0);
    const cpm       = parseFloat(ins.cpm || 0);
    const purchases = getPurchases(ins);
    const revenue   = getRevenue(ins);
    const roasApi   = parseRoas(ins);
    const roas      = roasApi != null ? roasApi : (spend > 0 ? revenue / spend : null);
    const cpa       = purchases > 0 ? spend / purchases : null;

    return {
      adId:          a.id,
      name:          a.name || "—",
      status:        (a.effective_status || "").toUpperCase(),
      launchDate:    a.created_time || null,
      format:        detectFormat(a.creative),
      thumbnailUrl:  pickThumbnail(a.creative),
      videoId:       a.creative?.video_id || null,
      spend, impressions: impr, clicks, ctr, cpc, cpm,
      purchases, revenue, roas, cpa
    };
  });
}

/* ── Filtering + sorting ───────────────────────────────────── */
function applyFilters(rows) {
  return rows.filter(r => {
    if (state.status !== "all" && r.status !== state.status) return false;
    if (state.format !== "all" && r.format !== state.format) return false;
    if (state.minSpend > 0 && r.spend < state.minSpend) return false;
    return true;
  });
}
function applySort(rows) {
  const sign = state.sortDir === "asc" ? 1 : -1;
  const key  = state.sortBy;
  return rows.slice().sort((a, b) => {
    let av = a[key], bv = b[key];
    if (key === "launchDate") { av = av ? Date.parse(av) : 0; bv = bv ? Date.parse(bv) : 0; }
    // Nulls sort last regardless of direction
    if (av == null && bv == null) return 0;
    if (av == null) return 1;
    if (bv == null) return -1;
    return (av - bv) * sign;
  });
}

/* ── Ad preview fetch + cache (Facebook feed iframe) ───────── */
const PREVIEW_TTL    = 24 * 60 * 60 * 1000; // 24h
const PREVIEW_FORMAT = "MOBILE_FEED_STANDARD";

function previewCacheKey(adId) { return `cp_preview_v1::${adId}::${PREVIEW_FORMAT}`; }

function readPreviewCache(adId) {
  try {
    const raw = localStorage.getItem(previewCacheKey(adId));
    if (!raw) return null;
    const entry = JSON.parse(raw);
    if (Date.now() - entry.timestamp > PREVIEW_TTL) return null;
    return entry.html;
  } catch { return null; }
}
function writePreviewCache(adId, html) {
  try {
    localStorage.setItem(previewCacheKey(adId), JSON.stringify({ timestamp: Date.now(), html }));
  } catch { /* quota — ignore */ }
}

async function fetchAdPreview(adId) {
  const cached = readPreviewCache(adId);
  if (cached) return cached;
  const data = await api(`${adId}/previews`, { ad_format: PREVIEW_FORMAT });
  const raw  = data.data?.[0]?.body;
  if (!raw) return null;
  const html = raw
    .replace(/&lt;/g, "<").replace(/&gt;/g, ">")
    .replace(/&amp;/g, "&").replace(/&quot;/g, '"');
  writePreviewCache(adId, html);
  return html;
}

/* IntersectionObserver lazy loader for preview iframes */
let _previewObserver = null;
function setupPreviewLazyLoad() {
  if (_previewObserver) _previewObserver.disconnect();
  _previewObserver = new IntersectionObserver(async (entries) => {
    for (const entry of entries) {
      if (!entry.isIntersecting) continue;
      const slot = entry.target;
      _previewObserver.unobserve(slot);
      const adId = slot.dataset.adId;
      try {
        const html = await fetchAdPreview(adId);
        if (html) {
          slot.innerHTML = html;
          slot.classList.add("loaded");
        } else {
          slot.innerHTML = renderFallbackThumb(slot.dataset.thumb, slot.dataset.format);
        }
      } catch (err) {
        console.warn(`Preview failed for ${adId}:`, err.message);
        slot.innerHTML = renderFallbackThumb(slot.dataset.thumb, slot.dataset.format);
      }
    }
  }, { rootMargin: "300px 0px" }); // pre-load 300px before scroll into view

  document.querySelectorAll(".cp-preview[data-ad-id]").forEach(el => _previewObserver.observe(el));
}

function renderFallbackThumb(thumbUrl, format) {
  if (!thumbUrl) {
    return `<div class="cp-thumb-empty">
      <svg width="32" height="32" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5">
        <rect x="3" y="3" width="18" height="18" rx="2"/><circle cx="8.5" cy="8.5" r="1.5"/><polyline points="21 15 16 10 5 21"/>
      </svg>
    </div>`;
  }
  const play = format === "video"
    ? `<div class="cp-play"><svg width="48" height="48" viewBox="0 0 24 24" fill="white"><circle cx="12" cy="12" r="11" fill="rgba(0,0,0,0.5)" stroke="white" stroke-width="1"/><polygon points="10,8 16,12 10,16"/></svg></div>`
    : "";
  return `<img src="${thumbUrl}" referrerpolicy="no-referrer" loading="lazy" onerror="this.style.display='none'">${play}`;
}

/* ── Render ────────────────────────────────────────────────── */
function renderCard(row) {
  const statusClass = row.status === "ACTIVE" ? "active" : "paused";
  const statusLabel = row.status === "ACTIVE" ? "Active" : "Paused";
  const safeThumb   = (row.thumbnailUrl || "").replace(/"/g, "&quot;");

  return `
    <div class="cp-card" data-ad-id="${row.adId}">
      <div class="cp-preview" data-ad-id="${row.adId}" data-thumb="${safeThumb}" data-format="${row.format}">
        <div class="cp-preview-loading">
          <div class="cp-spinner"></div>
          <div>Loading post…</div>
        </div>
      </div>
      <div class="cp-badges">
        <span class="cp-badge ${row.format}">${row.format}</span>
        <span class="cp-badge ${statusClass}">${statusLabel}</span>
      </div>
      <div class="cp-card-body">
        <div class="cp-name" title="${row.name.replace(/"/g, '&quot;')}">${row.name}</div>
        <div class="cp-metrics">
          <div>
            <div class="cp-metric-label">Spend</div>
            <div class="cp-metric-value">${formatCurrency(row.spend)}</div>
          </div>
          <div>
            <div class="cp-metric-label">ROAS</div>
            <div class="cp-metric-value ${row.roas == null ? 'muted' : ''}">${formatRoas(row.roas)}</div>
          </div>
          <div>
            <div class="cp-metric-label">CPA</div>
            <div class="cp-metric-value ${row.cpa == null ? 'muted' : ''}">${row.cpa != null ? formatCurrency(row.cpa) : "—"}</div>
          </div>
          <div>
            <div class="cp-metric-label">CTR</div>
            <div class="cp-metric-value">${formatPct(row.ctr)}</div>
          </div>
        </div>
      </div>
    </div>`;
}

function renderGrid() {
  const filtered = applySort(applyFilters(state.rawAds));
  const grid    = document.getElementById("cp-grid");
  const empty   = document.getElementById("cp-empty");

  if (!filtered.length) {
    grid.innerHTML = "";
    empty.style.display = "block";
  } else {
    empty.style.display = "none";
    grid.innerHTML = filtered.map(renderCard).join("");
    setupPreviewLazyLoad();
  }

  // Summary line
  const totalSpend = filtered.reduce((s, r) => s + r.spend, 0);
  const totalRev   = filtered.reduce((s, r) => s + r.revenue, 0);
  const blendedRoas = totalSpend > 0 ? totalRev / totalSpend : 0;
  document.getElementById("summary").innerHTML = `
    <span><strong>${filtered.length}</strong> ads</span>
    <span><strong>${formatCurrency(totalSpend)}</strong> spend</span>
    <span><strong>${formatCurrency(totalRev)}</strong> revenue</span>
    <span><strong>${formatRoas(blendedRoas)}</strong> blended ROAS</span>
  `;
}

function setPageHeader() {
  const client = CLIENTS[state.clientKey];
  const { since, until } = rangeToDates(state.range);
  document.getElementById("page-title").textContent = `${client.name} — Creative Performance`;
  document.getElementById("page-sub").textContent   = `${since} → ${until}`;
}

/* ── Loading orchestration ─────────────────────────────────── */
function showLoading(on) {
  document.getElementById("cp-loading").style.display = on ? "grid" : "none";
  document.getElementById("cp-grid").style.display    = on ? "none" : "grid";
  document.getElementById("summary").style.display    = on ? "none" : "flex";
}
function showError(msg) {
  document.getElementById("cp-error-message").textContent = msg;
  document.getElementById("cp-error").style.display = "block";
  document.getElementById("cp-loading").style.display = "none";
  document.getElementById("cp-grid").style.display = "none";
  document.getElementById("summary").style.display = "none";
}
function hideError() {
  document.getElementById("cp-error").style.display = "none";
}

async function loadData({ force = false } = {}) {
  const client = CLIENTS[state.clientKey];
  if (!client) return;
  hideError();
  setPageHeader();

  const { since, until } = rangeToDates(state.range);
  const key = cacheKey(client.adAccountId, since, until);

  if (!force) {
    const cached = readCache(key);
    if (cached) {
      state.currency    = cached.currency;
      state.rawAds      = cached.ads;
      state.lastFetched = cached.timestamp;
      updateLastUpdatedLabel();
      renderGrid();
      return;
    }
  } else {
    bustCache(key);
  }

  showLoading(true);
  document.getElementById("last-updated").textContent = "Fetching…";

  try {
    const [currency, ads, insights] = await Promise.all([
      fetchAccountCurrency(client.adAccountId),
      fetchAdsMetadata(client.adAccountId),
      fetchInsights(client.adAccountId, since, until)
    ]);
    state.currency = currency;
    state.rawAds   = mergeAdData(ads, insights);
    state.lastFetched = Date.now();
    writeCache(key, { currency, ads: state.rawAds });
    showLoading(false);
    updateLastUpdatedLabel();
    renderGrid();
  } catch (err) {
    console.error(err);
    showLoading(false);
    showError(err.message || "Failed to load.");
  }
}

function updateLastUpdatedLabel() {
  document.getElementById("last-updated").textContent = `Last updated ${relativeTime(state.lastFetched)}`;
}
// Keep "last updated" fresh while page is open
setInterval(() => {
  if (state.lastFetched) updateLastUpdatedLabel();
}, 30 * 1000);

/* ── URL state sync ────────────────────────────────────────── */
function readUrlState() {
  const p = new URLSearchParams(window.location.search);
  if (p.get("client") && CLIENTS[p.get("client")]) state.clientKey = p.get("client");
  if (p.get("range"))    state.range    = p.get("range");
  if (p.get("status"))   state.status   = p.get("status");
  if (p.get("format"))   state.format   = p.get("format");
  if (p.get("minSpend")) state.minSpend = parseFloat(p.get("minSpend")) || 0;
  if (p.get("sortBy"))   state.sortBy   = p.get("sortBy");
  if (p.get("sortDir"))  state.sortDir  = p.get("sortDir");
}
function writeUrlState() {
  const p = new URLSearchParams();
  p.set("client",   state.clientKey);
  p.set("range",    state.range);
  if (state.status   !== "all") p.set("status",   state.status);
  if (state.format   !== "all") p.set("format",   state.format);
  if (state.minSpend > 0)       p.set("minSpend", state.minSpend);
  if (state.sortBy   !== "spend") p.set("sortBy",  state.sortBy);
  if (state.sortDir  !== "desc")  p.set("sortDir", state.sortDir);
  const newUrl = `${window.location.pathname}?${p.toString()}`;
  window.history.replaceState(null, "", newUrl);
}

/* ── UI wiring ─────────────────────────────────────────────── */
function reflectStateToControls() {
  document.getElementById("client-select").value     = state.clientKey;
  document.getElementById("range-select").value      = state.range;
  document.getElementById("filter-status").value     = state.status;
  document.getElementById("filter-format").value     = state.format;
  document.getElementById("filter-min-spend").value  = state.minSpend || "";
  document.getElementById("sort-by").value           = state.sortBy;
  document.getElementById("sort-dir").value          = state.sortDir;
}

function attachListeners() {
  document.getElementById("client-select").addEventListener("change", e => {
    state.clientKey = e.target.value;
    writeUrlState();
    loadData();
  });
  document.getElementById("range-select").addEventListener("change", e => {
    state.range = e.target.value;
    writeUrlState();
    loadData();
  });
  document.getElementById("btn-refresh").addEventListener("click", () => loadData({ force: true }));

  ["filter-status", "filter-format", "filter-min-spend", "sort-by", "sort-dir"].forEach(id => {
    document.getElementById(id).addEventListener(id === "filter-min-spend" ? "input" : "change", e => {
      if (id === "filter-status")    state.status   = e.target.value;
      if (id === "filter-format")    state.format   = e.target.value;
      if (id === "filter-min-spend") state.minSpend = parseFloat(e.target.value) || 0;
      if (id === "sort-by")          state.sortBy   = e.target.value;
      if (id === "sort-dir")         state.sortDir  = e.target.value;
      writeUrlState();
      renderGrid();
    });
  });

  document.getElementById("btn-disconnect").addEventListener("click", () => {
    if (!confirm("Disconnect and clear your saved token?")) return;
    clearToken();
    showTokenScreen();
  });

  document.getElementById("btn-retry").addEventListener("click", () => loadData({ force: true }));

  document.getElementById("token-submit").addEventListener("click", () => {
    const val = document.getElementById("token-input").value.trim();
    if (!val) return;
    setToken(val);
    hideTokenScreen();
    loadData();
  });
}

/* ── Token screen ──────────────────────────────────────────── */
function showTokenScreen() {
  document.getElementById("token-screen").style.display = "flex";
  document.body.style.overflow = "hidden";
}
function hideTokenScreen() {
  document.getElementById("token-screen").style.display = "none";
  document.body.style.overflow = "";
}

/* ── Bootstrap ─────────────────────────────────────────────── */
document.addEventListener("DOMContentLoaded", () => {
  readUrlState();
  reflectStateToControls();
  attachListeners();

  if (!getToken()) {
    showTokenScreen();
  } else {
    hideTokenScreen();
    loadData();
  }
});
