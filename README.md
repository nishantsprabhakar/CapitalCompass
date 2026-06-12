# CapitalCompass

CapitalCompass is a local private-equity diligence workbench inspired by *Capital in the Shadows* by Nishant Prabhakar.

It ingests company decks, information documents, financial models, and source URLs to produce:

- Initial screening note with IC gates and diligence questions (`.docx`)
- PE-style IC memo (`.pptx`)
- Detailed sponsor financial model (`.xlsx`)

## Run Locally

```powershell
npm start
```

Then open:

```text
http://localhost:4186
```

## GitHub Pages

The static product page lives in `docs/` and is published to the `gh-pages` branch.

Use this repository setting:

- Settings -> Pages
- Build and deployment source: Deploy from a branch
- Branch: `gh-pages`
- Folder: `/root`

The site URL is expected to be:

```text
https://nishantsprabhakar.github.io/CapitalCompass/
```

## Core Workflow

1. Upload files or provide a local diligence folder path.
2. Add reputable source URLs for research.
3. Run initial screening or full deep-dive outputs.
4. Download the generated Word, PowerPoint, and Excel artifacts.

## Proprietary Scoring

CapitalCompass uses the Capital Compass IC Readiness Score v1.0, a rules-based PE underwriting algorithm. It scores five 20-point pillars: market quality, commercial proof, financial quality, operations/moat, and valuation/exit fit. Each pillar is built from granular subfactors with evidence tiers, then adjusted for missing evidence, critical risks, high risks, external-source confidence, and unresolved IC gates.

The score is intentionally conservative: management claims are not treated as proof unless the materials include source-backed evidence such as customer-level revenue, signed contracts, audited financials, margin bridges, downside cases, market sources, exit comps, or value-creation execution details.

## Strategic Buyer Readiness

CapitalCompass now includes an S&P-style acquisition readiness layer. Each diligence run produces a 100/100 platform architecture readiness score, a separate deal evidence quality score, source-quality score, audit record hash, competitive benchmark against leading market-intelligence and diligence tools, and a downloadable strategic acquisition readiness report.

The score separation is deliberate: platform architecture readiness measures what the product can evidence in local mode, while deal evidence quality and commercial proof remain separate diligence questions for a real strategic process.

## Optional AI API

CapitalCompass can optionally call an AI API during a diligence run. Supported modes are OpenAI-compatible chat completions, Groq chat completions, Anthropic Messages, Gemini `generateContent`, and custom OpenAI-compatible JSON endpoints. API keys are sent only from the browser to the local server for the current run and are not stored in generated outputs.

## Commercial Packaging

CapitalCompass is structured for three monetization paths:

- Per-deal pack: screening note, risk register, diligence questions, IC memo draft, and sponsor model.
- Team license: repeatable local deployment for funds, family offices, search funds, and advisory teams.
- Private deployment: firm-specific templates, sector playbooks, controlled data handling, and optional human review.

## Secure Access and Payments

The app includes local signup/login, salted `scrypt` password hashing, httpOnly session cookies, CSRF protection for mutations, admin-only user controls, premium feature gating, and a configurable hosted payment link. Set `CAPITAL_COMPASS_PAYMENT_LINK` to your live Stripe Payment Link or update it from the admin console. The first user to sign up becomes the local admin owner; alternatively set `CAPITAL_COMPASS_OWNER_EMAIL` and `CAPITAL_COMPASS_OWNER_PASSWORD` before startup to bootstrap an owner account.

## Publish for Monetization

Deploy the monetized app as a backend web service, not GitHub Pages. The repo includes a Render Blueprint:

[Deploy CapitalCompass on Render](https://render.com/deploy?repo=https://github.com/nishantsprabhakar/CapitalCompass)

See [DEPLOYMENT.md](DEPLOYMENT.md) for the production setup, required environment variables, payment-link setup, and admin monetization workflow.

## Security Positioning

CapitalCompass is local-first. It is designed for confidential deal materials that should not be uploaded into a generic hosted AI workspace. Uploaded IC templates are kept separate from company diligence materials and are used only for presentation styling.

## Disclaimer

CapitalCompass provides analytical workpapers and diligence workflow support. It does not provide legal, tax, accounting, or investment advice.

## Notes

The Pixxel template file is used locally as a style reference when present in `work/`.
