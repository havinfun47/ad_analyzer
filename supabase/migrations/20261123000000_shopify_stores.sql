-- Offline Admin API tokens for stores that installed "Scale Science Reporting".
-- Only Edge Functions (service role) may read or write it: RLS is on with no
-- policies, and the public API roles have no grants.
create table public.shopify_stores (
  shop          text primary key check (shop ~ '^[a-z0-9][a-z0-9-]*\.myshopify\.com$'),
  access_token  text not null,
  scope         text,
  installed_at  timestamptz not null default now()
);

alter table public.shopify_stores enable row level security;
revoke all on public.shopify_stores from anon, authenticated;
