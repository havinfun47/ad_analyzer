// Shopify one-time install (OAuth authorization code grant) for the
// "Scale Science Reporting" app.
//
//   GET /shopify-auth            App URL. Shopify sends the merchant here on install
//                                (and whenever the app is opened from the admin).
//                                Verifies the request and redirects to Shopify's
//                                consent screen.
//   GET /shopify-auth/callback   Redirect URL. Verifies the request, exchanges the
//                                code for a non-expiring offline Admin API token and
//                                stores it in public.shopify_stores (service role only).
//
// Secrets (Supabase > Edge Functions > Secrets): SHOPIFY_CLIENT_ID, SHOPIFY_CLIENT_SECRET.
// Deployed with verify_jwt = false: Shopify can't send a Supabase JWT, so every
// request is authenticated with Shopify's HMAC signature instead.

import { createClient } from "jsr:@supabase/supabase-js@2";

const SCOPES = "read_reports,read_orders";
// Only stores we have agreed to report on can install. Use the store's permanent
// .myshopify.com domain (what Shopify sends), not its vanity name.
const ALLOWED_SHOPS = new Set([
  "aqxm0v-y0.myshopify.com", // Lucille Health
]);
const STATE_MAX_AGE_MS = 10 * 60 * 1000;

const CLIENT_ID = Deno.env.get("SHOPIFY_CLIENT_ID") ?? "";
const CLIENT_SECRET = Deno.env.get("SHOPIFY_CLIENT_SECRET") ?? "";
const SUPABASE_URL = Deno.env.get("SUPABASE_URL") ?? "";
const SERVICE_KEY = Deno.env.get("SUPABASE_SERVICE_ROLE_KEY") ?? "";
const CALLBACK_URL = `${SUPABASE_URL}/functions/v1/shopify-auth/callback`;

const enc = new TextEncoder();

async function hmacHex(message: string): Promise<string> {
  const key = await crypto.subtle.importKey(
    "raw", enc.encode(CLIENT_SECRET), { name: "HMAC", hash: "SHA-256" }, false, ["sign"],
  );
  const sig = await crypto.subtle.sign("HMAC", key, enc.encode(message));
  return [...new Uint8Array(sig)].map((b) => b.toString(16).padStart(2, "0")).join("");
}

function safeEqual(a: string, b: string): boolean {
  if (a.length !== b.length) return false;
  let diff = 0;
  for (let i = 0; i < a.length; i++) diff |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return diff === 0;
}

// Shopify signs every query param except hmac: sorted, "k=v" joined with "&".
async function validShopifyHmac(params: URLSearchParams): Promise<boolean> {
  const hmac = params.get("hmac");
  if (!hmac) return false;
  const message = [...params.entries()]
    .filter(([k]) => k !== "hmac" && k !== "signature")
    .sort(([a], [b]) => (a < b ? -1 : a > b ? 1 : 0))
    .map(([k, v]) => `${k}=${v}`)
    .join("&");
  return safeEqual(await hmacHex(message), hmac);
}

function validShop(shop: string | null): shop is string {
  return !!shop && /^[a-z0-9][a-z0-9-]*\.myshopify\.com$/.test(shop) && ALLOWED_SHOPS.has(shop);
}

// Stateless CSRF state: "<issuedAt>~<nonce>~<shop>~<sig>", signed with the client secret.
async function makeState(shop: string): Promise<string> {
  const body = `${Date.now()}~${crypto.randomUUID()}~${shop}`;
  return `${body}~${await hmacHex(`state:${body}`)}`;
}

async function validState(state: string | null, shop: string): Promise<boolean> {
  const parts = (state ?? "").split("~");
  if (parts.length !== 4) return false;
  const [issuedAt, nonce, stateShop, sig] = parts;
  if (!safeEqual(await hmacHex(`state:${issuedAt}~${nonce}~${stateShop}`), sig)) return false;
  return stateShop === shop && Date.now() - Number(issuedAt) < STATE_MAX_AGE_MS;
}

// Plain text on purpose: Supabase serves function responses on *.supabase.co as
// text/plain, so HTML would show up as raw markup.
function page(title: string, message: string, status = 200): Response {
  return new Response(`${title}\n\n${message}\n`, {
    status, headers: { "content-type": "text/plain; charset=utf-8" },
  });
}

function db() {
  return createClient(SUPABASE_URL, SERVICE_KEY, { auth: { persistSession: false } });
}

Deno.serve(async (req) => {
  if (!CLIENT_ID || !CLIENT_SECRET) {
    return page("Not configured", "SHOPIFY_CLIENT_ID and SHOPIFY_CLIENT_SECRET are not set in Supabase.", 500);
  }
  const url = new URL(req.url);
  const params = url.searchParams;
  const shop = params.get("shop");

  if (!validShop(shop)) {
    console.warn("shop not in ALLOWED_SHOPS:", shop);
    return page("Not allowed", "This store isn't set up for Scale Science reporting.", 403);
  }
  if (!(await validShopifyHmac(params))) return page("Invalid request", "The request signature didn't verify.", 401);

  // ── Callback: exchange the code for an offline token ──────────────
  if (url.pathname.endsWith("/callback")) {
    if (!(await validState(params.get("state"), shop))) {
      return page("Expired", "This install link expired. Open the app again from Shopify to retry.", 400);
    }
    const code = params.get("code");
    if (!code) return page("Invalid request", "Missing authorization code.", 400);

    const res = await fetch(`https://${shop}/admin/oauth/access_token`, {
      method: "POST",
      headers: { "content-type": "application/json", accept: "application/json" },
      body: JSON.stringify({ client_id: CLIENT_ID, client_secret: CLIENT_SECRET, code }),
    });
    if (!res.ok) {
      console.error("token exchange failed", res.status, await res.text());
      return page("Couldn't connect", "Shopify refused the token exchange. Try installing again.", 502);
    }
    const { access_token, scope } = await res.json();

    const { error } = await db().from("shopify_stores").upsert({
      shop, access_token, scope, installed_at: new Date().toISOString(),
    });
    if (error) {
      console.error("token save failed", error);
      return page("Couldn't save", "Connected to Shopify but the token couldn't be saved.", 500);
    }
    return page("Connected", `${shop} is connected to the Scale Science ad analyzer (${scope}). You can close this tab.`);
  }

  // ── App URL: already installed? Otherwise start the consent flow ──
  const { data } = await db().from("shopify_stores").select("scope").eq("shop", shop).maybeSingle();
  const missingScope = SCOPES.split(",").some((s) => !(data?.scope ?? "").split(",").includes(s));
  if (data && !missingScope) {
    return page("Connected", `${shop} is connected to the Scale Science ad analyzer. Nothing to do here.`);
  }

  const authorize = new URL(`https://${shop}/admin/oauth/authorize`);
  authorize.searchParams.set("client_id", CLIENT_ID);
  authorize.searchParams.set("scope", SCOPES);
  authorize.searchParams.set("redirect_uri", CALLBACK_URL);
  authorize.searchParams.set("state", await makeState(shop));
  // No grant_options[]=per-user, so Shopify issues an offline (non-expiring) token.
  return Response.redirect(authorize.toString(), 302);
});
