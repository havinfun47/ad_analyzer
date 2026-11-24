// Diagnostic: reports the public IP this project's Edge Functions call out from,
// so it can be put on the Shopify app's IP allowlist. Requires the project's anon key.
Deno.serve(async () => {
  const r = await fetch("https://ipinfo.io/json");
  const info = await r.json();
  return Response.json({ ip: info.ip, org: info.org, region: info.region, country: info.country, edge_region: Deno.env.get("SB_REGION") ?? null });
});
