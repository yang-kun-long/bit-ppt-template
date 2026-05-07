# BIT PPT Template Generator

This directory contains a LaTeX-free Beijing Institute of Technology PPTX generator.
It reads a small YAML deck description and writes an editable PowerPoint file.

## Usage

```powershell
npm install
npm run build:ppt
npm run build:body-layouts
npm run build:charts
```

The sample output is written to:

```text
output/example.pptx
output/body-layout-test.pptx
output/chart-flow-test.pptx
```

You can also pass custom input and output paths:

```powershell
node bin/bit-ppt.mjs generate my-slides.yaml output/my-slides.pptx
```

Run preflight only:

```powershell
npm run check:ppt
npm run check:body-layouts
npm run check:charts
node src/generate.mjs content/overflow-test.yaml output/overflow-test.pptx --check
```

After linking the package locally, the same commands become shorter:

```powershell
npm link
bit-ppt generate content/example.yaml output/example.pptx
bit-ppt check content/example.yaml --json
bit-ppt list-layouts
```

When published to npm, users can run:

```powershell
npx bit-ppt-template generate input.yaml output.pptx
```

## Input Schema

The input file has two top-level keys:

- `meta`: deck-level title, subtitle, author, advisor, date.
- `slides`: ordered slide definitions.

Optional font replacement can be set in `meta.fonts`:

```yaml
meta:
  title: Demo
  fonts:
    cn: 微软雅黑
    cnLight: 微软雅黑 Light
    en: Arial
    serif: SimSun
    code: Consolas
```

CLI font options override `meta.fonts`:

```powershell
bit-ppt generate input.yaml output.pptx --font-cn "Noto Sans CJK SC" --font-en Arial --font-code Consolas
```

Supported layouts:

- `title`
- `agenda`
- `section`
- `bullets`
- `claim`
- `twoColumn`
- `cards`
- `table`
- `comparison`
- `timeline`
- `process`
- `architecture`
- `ablation`
- `caseStudy`
- `imageGrid`
- `code`
- `appendix`
- `flowchart`
- `chart`
- `problemSolution`
- `painOpportunity`
- `experimentDesign`
- `resultAnalysis`
- `riskMitigation`
- `contribution`
- `summary`
- `metrics`
- `matrix`
- `quote`
- `formula`
- `references`
- `imageText`
- `closing`

See `AI_CONTENT_GUIDE.md` for the full schema and writing constraints.

The intended AI workflow is:

```text
AI writes YAML content
        ↓
generate.mjs applies BIT visual rules
        ↓
editable .pptx
```

## Design Notes

The first template style uses extracted visual assets from the provided BIT PPTX:

- BIT green: `#006C39`
- accent red: `#A13F3D`
- Chinese font target: `微软雅黑`
- reusable school assets in `assets/`

The generator deliberately creates native PowerPoint text boxes, tables, and shapes
instead of rendering slides as images.

Formula slides are the exception to the "PptxGenJS only" path: `formula`
converts LaTeX-style math to Office Math / OMML and patches the generated PPTX
OpenXML. This keeps the project LaTeX-free while making formulas native Office
math objects instead of pictures.

Common text layouts also support inline `$...$` and `\(...\)` math. Table-cell
math is experimentally supported; references currently keep math syntax as plain
text.

## Preflight

Before writing the PPTX, the generator validates deck structure and estimates
text wrapping / content height for the layouts most likely to overflow.

`--check` returns JSON with:

- `validation.errors`: hard schema errors that stop generation.
- `validation.warnings`: length or density warnings that may still generate.
- `repairPrompt`: concise instructions that can be sent back to an AI model.
- `actions`: automatic preflight actions such as splitting long slides.

Normal generation stops on validation errors. Warnings are printed but do not
block output.

Currently handled:

- `bullets`: splits long bullet lists into continuation slides.
- `table`: splits long tables by rows.
- `references`: splits long reference lists into continuation slides.

The check is intentionally conservative. It is based on font size, textbox
width/height, line height, and approximate mixed Chinese/English character
width. It will not match PowerPoint's renderer perfectly, but it is good enough
to prevent most obvious overlaps and to produce repair signals for AI-generated
content.
