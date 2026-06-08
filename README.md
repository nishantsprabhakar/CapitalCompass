# Capital Compass

Capital Compass is a local private-equity diligence workbench inspired by *Capital in the Shadows* by Nishant Prabhakar.

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

## Core Workflow

1. Upload files or provide a local diligence folder path.
2. Add reputable source URLs for research.
3. Run initial screening or full deep-dive outputs.
4. Download the generated Word, PowerPoint, and Excel artifacts.

## Notes

The Pixxel template file is used locally as a style reference when present in `work/`.
