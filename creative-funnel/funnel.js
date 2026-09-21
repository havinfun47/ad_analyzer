/* ============================================================
   Scale Science — Creative Funnel
   TOF / MOF / BOF by relative ad frequency.
   ============================================================ */

const API_BASE   = "https://graph.facebook.com/v25.0";
const TOKEN_KEY  = "meta_access_token";
const CACHE_KEY  = "cf_cache_v1";
const CACHE_TTL  = 15 * 60 * 1000; // 15 min

const PREVIEW_CACHE_KEY = "cf_preview_v1";
const PREVIEW_TTL       = 24 * 60 * 60 * 1000; // 24h
const PREVIEW_FORMAT    = "MOBILE_FEED_STANDARD";
const PREVIEW_IFRAME_HEIGHT = 820;

const PURCHASE_ACTION = "offsite_conversion.fb_pixel_purchase";
const OMNI_PURCHASE   = "omni_purchase";

const CLIENTS = {
  anvytech: { name: "Anvy Tech", adAccountId: "act_575199276244807" },
  toothpod: { name: "Toothpod",  adAccountId: "act_727374130071249" },
  naturebee: { name: "NatureBee", adAccountId: "act_666605770419715" }
};

const STAGES = [
  { key: "tof", label: "Top of Funnel",    desc: "Cold prospecting · lowest ad frequency" },
  { key: "mof", label: "Middle of Funnel", desc: "Engaged audiences · mid frequency" },
  { key: "bof", label: "Bottom of Funnel", desc: "Warm/retargeting · highest ad frequency" }
];

/* ── State ─────────────────────────────────────────────────── */
const state = {
  clientKey: "anvytech",
  range:     "last_30d",
  minSpend:  100,
  currency:  "CAD",
  stages:    { tof: [], mof: [], bof: [] },
  lastFetched: null
};

/* ── Token ─────────────────────────────────────────────────── */
const getToken   = () => localStorage.getItem(TOKEN_KEY);
const setToken   = (t) => localStorage.setItem(TOKEN_KEY, t);
const clearToken = () => localStorage.removeItem(TOKEN_KEY);

/* ── Dates ─────────────────────────────────────────────────── */
const pad  = n => String(n).padStart(2, "0");
const ymd  = d => `${d.getFullYear()}-${pad(d.getMonth()+1)}-${pad(d.getDate())}`;

function rangeToDates(range) {
  const now = new Date();
  const until = ymd(now);
  if (range === "this_month") {
    return { since: ymd(new Date(now.getFullYear(), now.getMonth(), 1)), until };
  }
  const days = { last_7d: 7, last_14d: 14, last_30d: 30, last_90d: 90 }[range] || 30;
  const start = new Date(now);
  start.setDate(start.getDate() - (days - 1));
  return { since: ymd(start), until };
}

/* ── API ───────────────────────────────────────────────────── */
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

/* ── Cache ─────────────────────────────────────────────────── */
function cacheKey(accountId, since, until, minSpend) {
  return `${CACHE_KEY}::${accountId}::${since}::${until}::${minSpend}`;
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
  try { localStorage.setItem(key, JSON.stringify({ timestamp: Date.now(), ...payload })); }
  catch { /* quota */ }
}
function bustCache(key) { localStorage.removeItem(key); }

/* ── Formatters ────────────────────────────────────────────── */
function formatCurrency(val) {
  return new Intl.NumberFormat("en-CA", {
    style: "currency", currency: state.currency || "CAD",
    minimumFractionDigits: val >= 1000 ? 0 : 2, maximumFractionDigits: 2
  }).format(val || 0);
}
function formatRoas(val) { return val != null ? `${val.toFixed(2)}×` : "—"; }
function relativeTime(ts) {
  if (!ts) return "Never";
  const diff = Math.floor((Date.now() - ts) / 1000);
  if (diff < 5)   return "Just now";
  if (diff < 60)  return `${diff}s ago`;
  if (diff < 3600) return `${Math.floor(diff/60)}m ago`;
  return `${Math.floor(diff/3600)}h ago`;
}

/* ── Extractors ────────────────────────────────────────────── */
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

/* ── Data fetch ────────────────────────────────────────────── */
async function fetchAccountCurrency(adAccountId) {
  try {
    const info = await api(adAccountId, { fields: "currency" });
    return info.currency || "CAD";
  } catch { return "CAD"; }
}

async function fetchAdInsights(adAccountId, since, until) {
  const fields = [
    "ad_id", "ad_name", "spend", "impressions", "reach", "frequency",
    "actions", "action_values", "purchase_roas"
  ].join(",");
  return apiPaginated(`${adAccountId}/insights`, {
    fields,
    level: "ad",
    time_range: { since, until },
    limit: 500
  });
}

/* ── Bucket by frequency into TOF / MOF / BOF ─────────────── */
function bucketByFrequency(ads) {
  if (!ads.length) return { tof: [], mof: [], bof: [] };
  const sorted = ads.slice().sort((a, b) => a.frequency - b.frequency); // asc: low freq first (TOF)
  const n = sorted.length;
  const tofEnd = Math.ceil(n / 3);
  const mofEnd = Math.ceil((n * 2) / 3);
  return {
    tof: sorted.slice(0, tofEnd),
    mof: sorted.slice(tofEnd, mofEnd),
    bof: sorted.slice(mofEnd)
  };
}

function buildAdsFromInsights(insightsRows) {
  return insightsRows.map(r => {
    const spend       = parseFloat(r.spend || 0);
    const impressions = parseFloat(r.impressions || 0);
    const reach       = parseFloat(r.reach || 0);
    const frequency   = parseFloat(r.frequency || (reach > 0 ? impressions / reach : 0));
    const purchases   = getPurchases(r);
    const revenue     = getRevenue(r);
    const roasApi     = parseRoas(r);
    const roas        = roasApi != null ? roasApi : (spend > 0 ? revenue / spend : null);
    return {
      adId:     r.ad_id,
      name:     r.ad_name || "—",
      spend, impressions, reach, frequency,
      purchases, revenue, roas
    };
  });
}

/* ── Ad preview iframe (Facebook mobile feed) ─────────────── */
function previewCacheKey(adId) { return `${PREVIEW_CACHE_KEY}::${adId}::${PREVIEW_FORMAT}`; }
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
  try { localStorage.setItem(previewCacheKey(adId), JSON.stringify({ timestamp: Date.now(), html })); }
  catch { /* quota */ }
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

function injectPreviewHtml(slot, html) {
  slot.innerHTML = html;
  const iframe = slot.querySelector("iframe");
  if (!iframe) return;
  iframe.setAttribute("height", PREVIEW_IFRAME_HEIGHT);
  iframe.style.height = `${PREVIEW_IFRAME_HEIGHT}px`;
  iframe.style.width  = "100%";
  iframe.setAttribute("scrolling", "auto");
  slot.style.height = `${PREVIEW_IFRAME_HEIGHT}px`;
  slot.classList.add("loaded");
}

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
        if (html) injectPreviewHtml(slot, html);
        else slot.innerHTML = `<div class="cf-preview-loading">Preview unavailable</div>`;
      } catch (err) {
        console.warn(`Preview failed for ${adId}:`, err.message);
        slot.innerHTML = `<div class="cf-preview-loading" style="color:#FF6B6B">Preview failed</div>`;
      }
    }
  }, { rootMargin: "300px 0px" });

  document.querySelectorAll(".cf-preview[data-ad-id]").forEach(el => _previewObserver.observe(el));
}

/* ── Render ────────────────────────────────────────────────── */
function stageTotals(ads) {
  const spend   = ads.reduce((s, a) => s + a.spend, 0);
  const revenue = ads.reduce((s, a) => s + a.revenue, 0);
  return {
    count:   ads.length,
    spend,
    revenue,
    roas:    spend > 0 ? revenue / spend : 0
  };
}

function renderStage(stage, ads) {
  const totals = stageTotals(ads);
  if (!ads.length) {
    return `
      <div class="cf-stage">
        <div class="cf-stage-info">
          <div class="cf-stage-label">${stage.label}</div>
          <div class="cf-stat" style="color:var(--text-dim)">No ads in this stage</div>
        </div>
        <div></div>
      </div>`;
  }

  const cards = ads.map(ad => `
    <div class="cf-ad-card">
      <div class="cf-preview" data-ad-id="${ad.adId}">
        <div class="cf-preview-loading">
          <div class="cf-spinner"></div>
          <div>Loading post…</div>
        </div>
      </div>
      <div class="cf-ad-foot">
        <span><strong>${formatCurrency(ad.spend)}</strong></span>
        <span>ROAS <strong>${formatRoas(ad.roas)}</strong></span>
      </div>
    </div>`).join("");

  return `
    <div class="cf-stage">
      <div class="cf-stage-info">
        <div class="cf-stage-label">${stage.label}</div>
        <div class="cf-stat">Amount Spent: <strong>${formatCurrency(totals.spend)}</strong></div>
        <div class="cf-stat">ROAS: <strong>${formatRoas(totals.roas)}</strong></div>
        <div class="cf-stat" style="font-size:12px;color:var(--text-dim);margin-top:14px">
          ${totals.count} ad${totals.count === 1 ? "" : "s"}
        </div>
      </div>
      <div class="cf-stage-scroll">${cards}</div>
    </div>`;
}

function renderStages() {
  const container = document.getElementById("cf-stages");
  const anyAds = STAGES.some(s => state.stages[s.key].length);
  document.getElementById("cf-empty").style.display = anyAds ? "none" : "block";
  container.innerHTML = anyAds ? STAGES.map(s => renderStage(s, state.stages[s.key])).join("") : "";
  if (anyAds) setupPreviewLazyLoad();
}

function setPageHeader() {
  const client = CLIENTS[state.clientKey];
  const { since, until } = rangeToDates(state.range);
  document.getElementById("page-title").textContent = `${client.name} — Creative Funnel`;
  document.getElementById("page-sub").textContent   =
    `${since} → ${until} · Min spend ${formatCurrency(state.minSpend)} · TOF = lowest freq, BOF = highest`;
}

/* ── Orchestration ─────────────────────────────────────────── */
function showLoading(on) {
  document.getElementById("cf-loading").style.display = on ? "block" : "none";
  document.getElementById("cf-stages").style.display  = on ? "none" : "block";
}
function showError(msg) {
  document.getElementById("cf-error-message").textContent = msg;
  document.getElementById("cf-error").style.display = "block";
  document.getElementById("cf-loading").style.display = "none";
  document.getElementById("cf-stages").style.display = "none";
}
function hideError() { document.getElementById("cf-error").style.display = "none"; }

function updateLastUpdatedLabel() {
  document.getElementById("last-updated").textContent = `Last updated ${relativeTime(state.lastFetched)}`;
}
setInterval(() => { if (state.lastFetched) updateLastUpdatedLabel(); }, 30 * 1000);

async function loadData({ force = false } = {}) {
  const client = CLIENTS[state.clientKey];
  if (!client) return;
  hideError();
  setPageHeader();

  const { since, until } = rangeToDates(state.range);
  const key = cacheKey(client.adAccountId, since, until, state.minSpend);

  if (!force) {
    const cached = readCache(key);
    if (cached) {
      state.currency    = cached.currency;
      state.stages      = cached.stages;
      state.lastFetched = cached.timestamp;
      updateLastUpdatedLabel();
      renderStages();
      return;
    }
  } else {
    bustCache(key);
  }

  showLoading(true);
  document.getElementById("last-updated").textContent = "Fetching…";

  try {
    const [currency, insights] = await Promise.all([
      fetchAccountCurrency(client.adAccountId),
      fetchAdInsights(client.adAccountId, since, until)
    ]);
    state.currency = currency;

    // Aggregate any duplicate ad_id rows (Meta occasionally splits by attribution)
    const byAd = {};
    for (const r of insights) {
      if (!r.ad_id) continue;
      if (!byAd[r.ad_id]) byAd[r.ad_id] = r;
      else {
        // Sum spend / impressions / actions
        byAd[r.ad_id].spend       = String(parseFloat(byAd[r.ad_id].spend||0) + parseFloat(r.spend||0));
        byAd[r.ad_id].impressions = String(parseFloat(byAd[r.ad_id].impressions||0) + parseFloat(r.impressions||0));
      }
    }
    const merged = Object.values(byAd);

    const ads = buildAdsFromInsights(merged)
      .filter(a => a.spend >= state.minSpend && a.frequency > 0);

    state.stages      = bucketByFrequency(ads);
    state.lastFetched = Date.now();
    writeCache(key, { currency, stages: state.stages });

    showLoading(false);
    updateLastUpdatedLabel();
    renderStages();
  } catch (err) {
    console.error(err);
    showLoading(false);
    showError(err.message || "Failed to load.");
  }
}

/* ── URL state ─────────────────────────────────────────────── */
function readUrlState() {
  const p = new URLSearchParams(window.location.search);
  if (p.get("client") && CLIENTS[p.get("client")]) state.clientKey = p.get("client");
  if (p.get("range"))    state.range    = p.get("range");
  if (p.get("minSpend")) state.minSpend = parseFloat(p.get("minSpend")) || 100;
}
function writeUrlState() {
  const p = new URLSearchParams();
  p.set("client", state.clientKey);
  p.set("range",  state.range);
  if (state.minSpend > 0) p.set("minSpend", state.minSpend);
  window.history.replaceState(null, "", `${window.location.pathname}?${p.toString()}`);
}

/* ── UI wiring ─────────────────────────────────────────────── */
function reflectStateToControls() {
  document.getElementById("client-select").value = state.clientKey;
  document.getElementById("range-select").value  = state.range;
  document.getElementById("min-spend").value     = state.minSpend;
}

function attachListeners() {
  document.getElementById("client-select").addEventListener("change", e => {
    state.clientKey = e.target.value; writeUrlState(); loadData();
  });
  document.getElementById("range-select").addEventListener("change", e => {
    state.range = e.target.value; writeUrlState(); loadData();
  });
  document.getElementById("min-spend").addEventListener("change", e => {
    state.minSpend = parseFloat(e.target.value) || 0; writeUrlState(); loadData();
  });
  document.getElementById("btn-refresh").addEventListener("click", () => loadData({ force: true }));
  document.getElementById("btn-retry").addEventListener("click", () => loadData({ force: true }));

  document.getElementById("btn-disconnect").addEventListener("click", () => {
    if (!confirm("Disconnect and clear your saved token?")) return;
    clearToken();
    showTokenScreen();
  });

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

  if (!getToken()) showTokenScreen();
  else { hideTokenScreen(); loadData(); }
});
