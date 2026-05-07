# Codex Handoff

This directory is the active project root. The parent directory still contains
legacy LaTeX template files and the original BIT PPTX reference, but current
development should happen here.

## Project Goal

Build a LaTeX-free Beijing Institute of Technology style PPTX generator that is
friendly to AI agents and later MCP integration.

The generator reads YAML and writes editable PPTX files. It should avoid
Pandoc, LaTeX, Beamer, and screenshot-only slides.

## Current Stack

- Node.js ESM
- YAML input via `yaml`
- PPTX generation via `pptxgenjs`
- PPTX OpenXML post-processing via `jszip`
- Native Office Math formulas via `latex-to-omml`
- CLI entry: `bin/bit-ppt.mjs`
- Core generator: `src/generate.mjs`

## Important Commands

Run from this directory:

```powershell
npm install
npm run check:ppt
npm run build:ppt
npm run check:body-layouts
npm run build:body-layouts
npm run check:charts
npm run build:charts
```

Direct CLI:

```powershell
node bin/bit-ppt.mjs generate content/example.yaml output/example.pptx
node bin/bit-ppt.mjs check content/example.yaml --json
node bin/bit-ppt.mjs list-layouts
```

Font override example:

```powershell
node bin/bit-ppt.mjs generate content/body-layout-test.yaml output/font-test.pptx --font-cn "Noto Sans CJK SC" --font-code "Cascadia Mono"
```

Packaging check:

```powershell
npm pack --dry-run
```

## Current Capabilities

Supported slide layouts include title, agenda, section, bullets, claim,
twoColumn, cards, table, comparison, timeline, process, metrics, matrix, quote,
formula, references, imageText, closing, plus semantic/research layouts:

- problemSolution
- painOpportunity
- experimentDesign
- resultAnalysis
- riskMitigation
- contribution
- summary
- architecture
- ablation
- caseStudy
- imageGrid
- code
- appendix
- flowchart
- chart

Formula support:

- Display formulas and inline `$...$` / `\(...\)` formulas are converted to
  OMML and injected into the PPTX.
- Table-cell formula support was tested visually in WPS.
- Avoid reverting to SVG/image formulas except as an explicit fallback.

Chart and diagram support:

- `flowchart` uses editable PowerPoint shapes and arrow lines.
- `chart` uses native PowerPoint chart XML for bar, line, pie, doughnut,
  scatter, and area charts.

Validation:

- `check` returns `validation.errors`, `validation.warnings`, `repairPrompt`,
  and preflight `actions`.
- Generation stops on validation errors.
- `content/invalid-deck-test.yaml` is a fixture for repair prompt behavior.

Fonts:

- Defaults: `微软雅黑`, `微软雅黑 Light`, `Arial`, `SimSun`, `Consolas`.
- YAML `meta.fonts` and CLI font options can replace fonts.
- Do not bundle Microsoft YaHei TTF files due to font licensing concerns.

## Verified Outputs

Known test decks:

- `content/example.yaml`
- `content/body-layout-test.yaml`
- `content/chart-flow-test.yaml`
- `content/formula-test.yaml`
- `content/inline-formula-test.yaml`
- `content/table-formula-test.yaml`

Recent checks were clean:

- `npm run check:ppt`
- `npm run check:body-layouts`
- `npm run check:charts`

`npm pack --dry-run` produced a package around 260 KB and did not include
`output/`.

## User Preferences

- The user wants a local-agent-friendly solution, likely MCP for Codex/Claude.
- Output must stay editable PPTX whenever practical.
- Formula output should be native OMML, not images.
- Keep work pragmatic and test in WPS screenshots when visual quality matters.
- The BIT style need not be pixel-perfect, but should stay consistent and clean.

## Likely Next Step

Implement MCP integration inside this directory.

Suggested MCP tools:

- `list_layouts`
- `validate_deck`
- `preflight_deck`
- `generate_pptx`
- `get_repair_prompt`

Reuse exported functions from `src/generate.mjs`:

- `listLayouts`
- `checkDeck`
- `checkDeckFile`
- `generateDeckFile`
- `validateDeck`

Keep CLI and MCP sharing the same implementation. Do not fork generator logic.
