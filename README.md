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

The static product page lives in `docs/`.

If the GitHub Actions Pages workflow fails because Pages is not enabled, use this repository setting:

- Settings -> Pages
- Build and deployment source: Deploy from a branch
- Branch: `main`
- Folder: `/docs`

The site URL is expected to be:

```text
https://nishantsprabhakar.github.io/CapitalCompass/
```

## Core Workflow

1. Upload files or provide a local diligence folder path.
2. Add reputable source URLs for research.
3. Run initial screening or full deep-dive outputs.
4. Download the generated Word, PowerPoint, and Excel artifacts.

## Commercial Packaging

CapitalCompass is structured for three monetization paths:

- Per-deal pack: screening note, risk register, diligence questions, IC memo draft, and sponsor model.
- Team license: repeatable local deployment for funds, family offices, search funds, and advisory teams.
- Private deployment: firm-specific templates, sector playbooks, controlled data handling, and optional human review.

## Security Positioning

CapitalCompass is local-first. It is designed for confidential deal materials that should not be uploaded into a generic hosted AI workspace. Uploaded IC templates are kept separate from company diligence materials and are used only for presentation styling.

## Disclaimer

CapitalCompass provides analytical workpapers and diligence workflow support. It does not provide legal, tax, accounting, or investment advice.

## Notes

The Pixxel template file is used locally as a style reference when present in `work/`.
