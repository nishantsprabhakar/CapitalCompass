# CapitalCompass Monetized Deployment

CapitalCompass needs a backend host because login, admin controls, sessions, premium gating, Office exports, and payment-link settings run through `server.mjs`. Do not deploy the monetized app as a static GitHub Pages site.

## Recommended Host

Use Render Web Service via the included `render.yaml` Blueprint.

Render will:

- Run the Node server with `npm start`
- Expose the app over HTTPS
- Health check `/api/health`
- Auto-deploy from the `main` branch
- Preserve `work/` on a persistent disk for local auth store, uploads, and generated runtime files

## One-Time Setup

1. Create a Stripe Payment Link for the premium plan.
2. Open Render and create a new Blueprint from:
   `https://github.com/nishantsprabhakar/CapitalCompass`
3. Confirm the `capital-compass` web service from `render.yaml`.
4. Set these Render environment variables:

   - `CAPITAL_COMPASS_PAYMENT_LINK`: live Stripe Payment Link
   - `CAPITAL_COMPASS_OWNER_EMAIL`: your admin email
   - `CAPITAL_COMPASS_OWNER_PASSWORD`: a strong admin password
   - `CAPITAL_COMPASS_SECURE_COOKIE`: `1`

5. Deploy.
6. Open the Render URL.
7. Login with the owner credentials.
8. Open `/admin` to manage users, plans, premium feature access, promo codes, and payment link.

## Monetization Flow

1. User signs up.
2. User clicks premium checkout.
3. Stripe handles international payment on its hosted checkout page.
4. Admin verifies payment in Stripe.
5. Admin upgrades the user to `premium` or `enterprise`, or grants specific features:
   - Deep-dive IC memo and financial model
   - AI enrichment
   - Premium exports

## Production Notes

- Use HTTPS only.
- Keep `CAPITAL_COMPASS_OWNER_PASSWORD` private.
- For institutional scale, migrate `work/auth-store.json` to managed Postgres and add Stripe webhooks for automatic entitlement upgrades.
- For PE client data, add a formal privacy policy, terms of service, deletion policy, and enterprise security documentation before broad commercialization.
