// Shopify KPIs for the ad analyzer's Performance Overview.
//
//   POST /shopify-metrics
//   body: { client, metaToken, since, until, compSince, compUntil }   (dates yyyy-mm-dd)
//
// Access: the caller must already be able to see the client's Meta ad account, so we
// ask Meta with their token (the same one the report page uses) before touching
// Shopify. The token travels in the body, not a header, so it never lands in
// Supabase's request logs. Returns aggregates only.
//
// Deployed with verify_jwt = false: authentication is the Meta check above.

import { createClient } from "jsr:@supabase/supabase-js@2";

const SHOPIFY_API_VERSION = "2026-10";
const META_API = "https://graph.facebook.com/v25.0";

// Client key (as in the site's config.js) -> store + ad account. Server-side on
// purpose: callers name a client, never a store.
const CLIENTS: Record<string, { shop: string; adAccountId: string }> = {
  lucillehealth: { shop: "aqxm0v-y0.myshopify.com", adAccountId: "act_1685611612853923" },
};

const ALLOWED_ORIGINS = new Set([
  "https://accounts.scalescientist.com",
  "https://havinfun47.github.io",
]);

const DATE_RE = /^\d{4}-\d{2}-\d{2}$/;
const MAX_RANGE_DAYS = 400;

function cors(req: Request): Record<string, string> {
  const origin = req.headers.get("origin") ?? "";
  return {
    "Access-Control-Allow-Origin": ALLOWED_ORIGINS.has(origin) ? origin : "https://accounts.scalescientist.com",
    "Access-Control-Allow-Methods": "POST, OPTIONS",
    "Access-Control-Allow-Headers": "content-type",
    "Access-Control-Max-Age": "86400",
    "Vary": "Origin",
  };
}

function json(req: Request, body: unknown, status = 200): Response {
  return new Response(JSON.stringify(body), {
    status, headers: { ...cors(req), "content-type": "application/json" },
  });
}

function validRange(since: unknown, until: unknown): boolean {
  if (typeof since !== "string" || typeof until !== "string") return false;
  if (!DATE_RE.test(since) || !DATE_RE.test(until)) return false;
  const days = (Date.parse(until) - Date.parse(since)) / 86400000;
  return Number.isFinite(days) && days >= 0 && days <= MAX_RANGE_DAYS;
}

async function canSeeAdAccount(metaToken: string, adAccountId: string): Promise<boolean> {
  const res = await fetch(`${META_API}/${adAccountId}?fields=id`, {
    headers: { Authorization: `Bearer ${metaToken}` },
  });
  if (!res.ok) return false;
  const body = await res.json().catch(() => ({}));
  return body.id === adAccountId;
}

async function shopifyql(shop: string, token: string, query: string) {
  const res = await fetch(`https://${shop}/admin/api/${SHOPIFY_API_VERSION}/graphql.json`, {
    method: "POST",
    headers: { "content-type": "application/json", "X-Shopify-Access-Token": token },
    body: JSON.stringify({
      query: `query($q: String!) {
        shop { currencyCode }
        shopifyqlQuery(query: $q) { tableData { rows } parseErrors }
      }`,
      variables: { q: query },
    }),
  });
  const body = await res.json().catch(() => null);
  const result = body?.data?.shopifyqlQuery;
  if (!res.ok || body?.errors || !result || result.parseErrors?.length) {
    console.error("shopifyql failed", res.status, JSON.stringify(body?.errors ?? result?.parseErrors ?? body));
    throw new Error("Shopify query failed");
  }
  return { currency: body.data.shop.currencyCode as string, rows: (result.tableData?.rows ?? []) as Record<string, string | null>[] };
}

const num = (v: string | null | undefined) => (v == null || v === "" ? 0 : Number(v));

async function periodMetrics(shop: string, token: string, since: string, until: string) {
  const range = `SINCE ${since} UNTIL ${until}`;
  const [split, totals] = await Promise.all([
    shopifyql(shop, token,
      `FROM sales SHOW orders, total_sales GROUP BY subscription_or_one_time, new_or_returning_customer ${range}`),
    shopifyql(shop, token,
      `FROM sales SHOW orders, customers, new_customers, returning_customers, returning_customer_rate ${range}`),
  ]);

  const m = {
    subscriptionOrders: 0, subscriptionSales: 0,
    oneTimeOrders: 0, oneTimeSales: 0,
    subscriptionOrdersNew: 0, subscriptionOrdersReturning: 0,
  };
  // Rows with no subscription_or_one_time carry refunds/adjustments, not orders: skip them.
  for (const r of split.rows) {
    const orders = num(r.orders), sales = num(r.total_sales);
    if (r.subscription_or_one_time === "subscription") {
      m.subscriptionOrders += orders; m.subscriptionSales += sales;
      if (r.new_or_returning_customer === "New") m.subscriptionOrdersNew += orders;
      if (r.new_or_returning_customer === "Returning") m.subscriptionOrdersReturning += orders;
    } else if (r.subscription_or_one_time === "one_time") {
      m.oneTimeOrders += orders; m.oneTimeSales += sales;
    }
  }

  const t = totals.rows[0] ?? {};
  const rate = t.returning_customer_rate;
  return {
    currency: split.currency,
    metrics: {
      ...m,
      orders: num(t.orders),
      customers: num(t.customers),
      newCustomers: num(t.new_customers),
      returningCustomers: num(t.returning_customers),
      returningCustomerRate: rate == null || rate === "" ? null : Number(rate),
    },
  };
}

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") return new Response(null, { status: 204, headers: cors(req) });
  if (req.method !== "POST") return json(req, { error: "Use POST" }, 405);

  let body: Record<string, unknown>;
  try { body = await req.json(); } catch { return json(req, { error: "Invalid JSON" }, 400); }

  const client = CLIENTS[String(body.client ?? "")];
  if (!client) return json(req, { error: "No Shopify store for this client" }, 404);
  if (typeof body.metaToken !== "string" || !body.metaToken) return json(req, { error: "Missing Meta token" }, 401);
  if (!validRange(body.since, body.until)) return json(req, { error: "Invalid date range" }, 400);
  const hasCompare = body.compSince != null || body.compUntil != null;
  if (hasCompare && !validRange(body.compSince, body.compUntil)) {
    return json(req, { error: "Invalid comparison range" }, 400);
  }

  if (!(await canSeeAdAccount(body.metaToken, client.adAccountId))) {
    return json(req, { error: "This Meta token can't access this client's ad account" }, 403);
  }

  const db = createClient(Deno.env.get("SUPABASE_URL")!, Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!, {
    auth: { persistSession: false },
  });
  const { data: store } = await db.from("shopify_stores").select("access_token").eq("shop", client.shop).maybeSingle();
  if (!store) return json(req, { error: "Shopify store not connected" }, 409);

  try {
    const [current, previous] = await Promise.all([
      periodMetrics(client.shop, store.access_token, body.since as string, body.until as string),
      hasCompare
        ? periodMetrics(client.shop, store.access_token, body.compSince as string, body.compUntil as string)
        : Promise.resolve(null),
    ]);
    return json(req, {
      currency: current.currency,
      current: current.metrics,
      previous: previous?.metrics ?? null,
    });
  } catch (err) {
    return json(req, { error: (err as Error).message }, 502);
  }
});
