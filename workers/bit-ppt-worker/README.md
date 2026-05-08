# BIT PPT Worker Spike

This is a Cloudflare Worker feasibility spike for the future online service.

Current scope:

- `GET /health`
- `POST /check`

Out of scope for this spike:

- PPTX generation
- OMML formula conversion
- image file existence checks

The worker intentionally reports:

```json
{
  "capabilities": {
    "check": true,
    "generate": false,
    "omml": false
  }
}
```

## Run

```powershell
npm install
npm run check
npm run dev
```

Example `/check` request:

```powershell
Invoke-RestMethod `
  -Method Post `
  -Uri http://localhost:8787/check `
  -ContentType "application/json" `
  -Body '{"deckYaml":"slides:\n  - layout: imageText\n    title: Worker 占位图\n    image:\n      mode: placeholder\n      prompt: 一张流程图。\n    text:\n      - 先在线校验。\n"}'
```

## Deploy

After `npx wrangler login`:

```powershell
npm run deploy
```

Then test the deployed URL:

```powershell
Invoke-RestMethod https://<worker-url>/health
```

## Notes

This worker does not import `src/generate.mjs` yet because the current Node
generator imports `fs`, `path`, `pptxgenjs`, and `latex-to-omml`. This spike is
only meant to confirm that the online check/preflight API shape is viable on
Cloudflare Workers before a larger core split.
