/* ============================================================
   Scale Science — Client Configuration
   ============================================================ */

const CLIENTS = {
  root: {
    name: "Root",
    adAccountId: "act_575199276244807",
    currency: "CAD",
    logo: null
  },
  toothpod: {
    name: "Toothpod",
    adAccountId: "act_727374130071249",
    currency: "CAD",
    logo: null
  },
  naturebee: {
    name: "NatureBee",
    adAccountId: "act_666605770419715",
    currency: "CAD",
    logo: null
  },
  lucillehealth: {
    name: "Lucille Health",
    adAccountId: "act_1685611612853923",
    currency: "USD",
    logo: null,
    shopify: true   // Shopify KPIs in Performance Overview (store set in supabase/functions/shopify-metrics)
  }
};

const META_API_VERSION = "v25.0";
const META_API_BASE    = `https://graph.facebook.com/${META_API_VERSION}`;
const TOKEN_KEY        = "meta_access_token";
const SHOPIFY_METRICS_URL = "https://erbdmcgjevejbrpvxgge.supabase.co/functions/v1/shopify-metrics";
