# Changelog

## 0.2.0 - 2026-05-08

### Added

- Added real PowerPoint/WPS speaker notes via `speakerNotes`, with CLI guide and MCP guide support.
- Added editable image placeholder mode for `imageText`, `caseStudy`, and `imageGrid`.
- Added automatic `imageText` placeholder variants when image aspect ratio is unknown.
- Added `bit-ppt-mcp`, a stdio MCP server adapter that reuses the same generator implementation as the CLI.
- Added MCP tools: `list_layouts`, `get_guide`, `validate_deck`, `preflight_deck`, `get_repair_prompt`, and `generate_pptx`.
- Added `bit-ppt-http`, a Node HTTP API with `/health`, `/check`, and `/generate` endpoints.
- Added demo decks for speaker notes and image placeholders.

### Changed

- Extended progressive guide output with `speaker-notes` and `image-placeholder`.
- Extended `doctor` dependency checks to include the MCP SDK and `zod`.

### Notes

- CLI remains the primary local interface.
- MCP is now available as a second release target.
- Node HTTP is available as the third release target for YAML upload and PPTX download.

## 0.1.0 - 2026-05-07

### Added

- Initial LaTeX-free BIT-style PPTX generator.
- YAML input, editable PPTX output, validation, repair prompts, preflight splitting, native charts, flowcharts, and OMML formula post-processing.
