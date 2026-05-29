# AI Generative Patch Template

Local web MVP for generating patch templates from imported DXF outlines.

## Rules Implemented

- Generates 4 template layers.
- Each board defaults to `150mm x 180mm`.
- Bottom board corners are curved.
- Needle positioning slot defaults to `30mm x 11mm`.
- Layer 1 uses the imported patch DXF outline only.
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

## Export

- Export uses the browser download API, so it works in current Chrome and Edge on Windows.
- Downloaded files are named from the imported DXF, for example `Patch 5-patch-template.dxf`.
- The export excludes the original DXF preview/source layer and includes only generated template layers.

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
