# AI Generative Patch Template

Local web MVP for generating patch templates from imported DXF outlines or high-contrast scan images.

## Rules Implemented

- Generates 4 template layers.
- Each board defaults to `150mm x 180mm`.
- Bottom board corners are curved.
- Needle positioning slot defaults to `30mm x 11mm`.
- Layer 1 uses the imported patch DXF outline only.
- Scan/image import auto-traces the largest foreground shape into a vector outline.
- Layers 2, 3, and 4 include an additional outward offset line, default `7mm`.
- Needle slot top edge is `50mm` from board top edge.
- Needle slot right edge is `11.5mm` from board right edge.

## Run

From the project folder:

```bash
python3 -m http.server 8769
```

Windows alternative:

```powershell
py -m http.server 8769
```

Open:

```text
http://127.0.0.1:8769
```

Image tracing uses browser-side Potrace from the normal web app and does not require Python or the native Potrace command.

## Export

- Export uses the browser download API, so it works in current Chrome and Edge on Windows.
- Downloaded files are named from the imported DXF, for example `Patch 5-patch-template.dxf`.
- The export excludes the original DXF preview/source layer and includes only generated template layers.

## Scan/Image Import

- Use `Import Scan` for PNG, JPEG, WebP, BMP, PDF, SVG, and other browser-readable image files.
- The tracer expects a high-contrast scan/photo with the patch as the largest foreground shape.
- `Scan trace width, mm` controls the real-world width assigned to the traced outline before template generation.
- After import, the original raster/PDF first page is shown under the traced vector so PIC can compare and adjust the trace.
- `Trace sensitivity`, `Trace detail`, `Smooth`, and `Original opacity` can be adjusted before applying the patch rules. Use `Detail` for real shape/corner fidelity and raise `Smooth` only when PIC needs more scan noise cleanup before the offset is generated.
- The scan tracer runs Potrace inside Chrome/Edge from the bundled `assets/potrace.js` file. Potrace preview updates are debounced and traced on a smaller working raster for smoother browser performance.
- Layer 2-4 offset is generated from one clean outer envelope for multi-part scan artwork, not every logo dot/hole/internal detail. The visible pink patch artwork still keeps the traced details.
- After PIC clicks `Apply Patch Rules`, the traced vector uses the same Predict, Apply Patch Rules, and Export DXF flow as a DXF import.

## Tracing Tools

- Potrace is the recommended production upgrade for black/white patch scans because it fits smoother curves before the DXF offset is generated.
- Browser-side Potrace is bundled locally in `assets/potrace.js`, adapted from the original JavaScript Potrace port. It is GPL licensed, same as Potrace.
- ImageMagick and Autotrace are best kept as local preprocessing/server tools, not browser-only dependencies.
- diffVG is not wired into the app yet. It can be explored later in the existing `pytorch/` area, but it is heavier than Potrace and needs per-image optimization settings before it is suitable for PIC workflow.

## Power Automate Setup

This public repo does not include live Power Automate URLs.

For an internal MA Portal deployment, define the URLs before `app.js` loads:

```html
<script>
  window.PATCH_TEMPLATE_CONFIG = {
    powerAutomateTrainingUrl: "YOUR_SAVE_FLOW_URL",
    powerAutomateReadUrl: "YOUR_READ_FLOW_URL"
  };
</script>
<script src="app.js"></script>
```

Use `config.example.js` as the shape only. Do not commit live PA URLs to a public repo.
