# CapitalCompass Account Setup Runbook

Purpose: set up the accounts needed to run a same-day sale process for CapitalCompass at a USD 2M+ target.

Important: these accounts must be created by Nishant or an authorized representative because they require identity details, email/phone verification, payment details, and acceptance of platform terms.

## Priority 1: Sale Process Accounts

### 1. Acquisition Marketplace

Primary:

- Acquire.com: https://acquire.com/
- Flippa: https://flippa.com/sell

Use for:

- Creating a confidential listing
- Reaching buyers actively searching for SaaS / AI tools / digital assets
- Running buyer Q&A

Setup data needed:

- Founder name
- Business/app name: CapitalCompass
- Contact email
- Asking price: anchor USD 7.5M to USD 10M; floor USD 2M
- Category: SaaS / AI Apps & Tools / FinTech / Private Markets Workflow
- Website: https://nishantsprabhakar.github.io/CapitalCompass/acquire.html
- GitHub repo: https://github.com/nishantsprabhakar/CapitalCompass
- Summary from `sales/01_ACQUISITION_TEASER.md`

Recommendation:

- Create marketplace accounts, but do not rely on them alone. The USD 2M+ case is stronger through direct strategic outreach than open marketplace bidding.

### 2. Data Room / Teaser Tracking

Recommended:

- DocSend: https://www.docsend.com/

Use for:

- Controlled teaser sharing
- Data-room links
- Page-by-page buyer engagement analytics
- NDA gating and watermarking where available

Upload first:

- `sales/01_ACQUISITION_TEASER.md`
- `sales/04_VALUATION_AND_DEAL_TERMS.md`
- `sales/05_DATA_ROOM_CHECKLIST.md`
- Product screenshots
- Link to acquisition page

### 3. E-Signature / NDA

Recommended:

- DocuSign: https://www.docusign.com/products/electronic-signature

Use for:

- NDA
- LOI
- Asset purchase agreement routing

Setup data needed:

- Legal name/entity
- Email
- Billing method if paid plan is required
- NDA template or counsel-prepared NDA

### 4. Meeting Scheduler

Recommended:

- Calendly: https://calendly.com/

Use for:

- 20-minute acquisition intro calls
- 45-minute product demo
- Buyer diligence slots

Suggested event names:

- CapitalCompass Strategic Acquisition Intro - 20 min
- CapitalCompass Product Demo - 45 min

### 5. Lightweight CRM

Recommended:

- HubSpot CRM: https://www.hubspot.com/products/crm

Use for:

- Tracking buyer outreach
- Logging replies
- Follow-up reminders
- Pipeline stages: Targeted, Contacted, Interested, NDA, Demo, LOI, Closed / Lost

Import:

- `sales/02_BUYER_TARGETS.csv`

## Priority 2: Monetization / Live Product Accounts

### 6. Stripe

Recommended:

- Stripe Payment Links: https://stripe.com/payments/payment-links

Use for:

- Premium access payments
- Pilot payments
- Annual private deployment retainers

Products to create:

- CapitalCompass Per-Deal Pack: USD 1,500
- CapitalCompass Team License: USD 2,500/month
- CapitalCompass Private Deployment: USD 25,000/year

After creation:

- Copy the live Payment Link
- Add it to Render env var `CAPITAL_COMPASS_PAYMENT_LINK`
- Add it in the CapitalCompass admin console

### 7. Render

Recommended:

- Render Blueprint deploy link: https://render.com/deploy?repo=https://github.com/nishantsprabhakar/CapitalCompass

Use for:

- Monetized hosted backend
- Login/admin/premium gating
- Payment link configuration

Required env vars:

- `CAPITAL_COMPASS_PAYMENT_LINK`
- `CAPITAL_COMPASS_OWNER_EMAIL`
- `CAPITAL_COMPASS_OWNER_PASSWORD`
- `CAPITAL_COMPASS_SECURE_COOKIE=1`

## Priority 3: Outreach Accounts

### 8. LinkedIn

Use for:

- Direct corporate development / product leadership outreach
- Warm intros through network

Search targets:

- Corporate Development
- Head of Product
- Private Markets
- M&A Workflow
- Strategy
- CEO / Founder for smaller targets

### 9. Email Domain

Use:

- A professional email alias for sale process, e.g. `acquire@capitalcompass.ai` or `nishant.p@skegen.com`

Minimum setup:

- Dedicated folder/label: CapitalCompass Sale
- Email signature
- Calendar link
- Acquisition page link

## Same-Day Setup Order

1. DocSend
2. Calendly
3. HubSpot CRM
4. Stripe
5. Render
6. Acquire.com
7. Flippa
8. DocuSign
9. LinkedIn outbound

## Security Rules

- Do not upload private local auth stores.
- Do not upload user/deal files unless buyer has signed NDA.
- Do not share admin credentials.
- Share source repo only after NDA or serious buyer qualification.
- Use viewer-specific tracked links for teaser/data room.

## Listing Copy

Use this headline:

CapitalCompass - PE diligence AI/workflow cockpit for IC memo, sponsor model, and deal pipeline automation

Use this short description:

CapitalCompass is a local-first private equity diligence workbench that turns company decks, documents, financial models, and source URLs into screening notes, diligence questions, IC memo drafts, sponsor models, source-quality checks, and deal pipeline analytics. Strategic acquisition target for data-room, PE CRM, private-market intelligence, and research platforms.

Use this valuation line:

Seeking strategic acquisition. USD 2M minimum floor; preference for USD 5M to USD 10M strategic outcome or structured transaction with meaningful upfront consideration.
