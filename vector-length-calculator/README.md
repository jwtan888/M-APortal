# DXF Vector Length Calculator

A precise tool for calculating total vector path length from DXF files, designed for garment manufacturing machine runtime estimation and material consumption planning.

## Features

### 🎯 Core Functionality
- **Import DXF files** - Drag & drop or click to upload
- **Calculate total vector length** - Precise measurement in cm (primary unit)
- **Multi-unit display** - Shows results in cm, mm, m, inches, feet, and yards
- **Entity breakdown** - See length contribution by entity type (LINE, ARC, CIRCLE, POLYLINE, etc.)
- **Machine runtime estimation** - Calculate estimated run time based on machine speed
- **Export results** - Download as JSON or CSV for audit trails

### 🔍 Audit & Verification
- **Independent Python audit script** - Cross-verify calculations with `dxf-length-audit.py`
- **Detailed entity tracking** - Every entity's length is recorded
- **Unit detection** - Automatically detects DXF drawing units
- **Transparent formulas** - All calculation methods documented in code

## Files Included

| File | Purpose |
|------|---------|
| `index.html` | Main web interface |
| `dxf-parser.js` | Pure JavaScript DXF file parser |
| `app.js` | Application logic and length calculations |
| `dxf-length-audit.py` | Python audit/verification script |

## Usage

### Web Interface (Browser)

1. Open `index.html` in any modern browser
2. Drag & drop your DXF file or click to browse
3. Click "Calculate Vector Length"
4. View results in centimeters and other units
5. Optionally calculate machine runtime
6. Export results as JSON or CSV

### Command Line (Python Audit)

```bash
# Basic usage
python dxf-length-audit.py your-file.dxf

# With JSON output
python dxf-length-audit.py your-file.dxf --output results.json
```

## Supported DXF Entities

The calculator handles these entity types:

| Entity | Calculation Method |
|--------|-------------------|
| LINE | Euclidean distance: √[(x₂-x₁)² + (y₂-y₁)² + (z₂-z₁)²] |
| CIRCLE | Circumference: 2πr |
| ARC | Arc length: r × θ (θ in radians) |
| POLYLINE / LWPOLYLINE | Sum of segment lengths (including bulge arcs) |
| SPLINE | Numerical integration with adaptive sampling |
| ELLIPSE | Ramanujan's approximation formula |
| POINT | 0 (no length) |

## Accuracy Notes

### Critical Measurements
This tool is designed for production-critical measurements. Key accuracy features:

1. **Double precision** - All calculations use 64-bit floating point
2. **4 decimal places** - Results shown to 0.0001 precision
3. **Bulge handling** - Properly calculates arc segments in polylines
4. **Angle normalization** - Correctly handles arcs crossing 0/2π boundary
5. **Unit conversion** - Accurate conversion between metric and imperial

### Audit Trail
For critical production decisions:
1. Run the web calculator
2. Verify with the Python audit script
3. Compare results - they should match exactly
4. Export JSON/CSV for documentation

### Known Limitations
- SPLINE entities use numerical approximation (100 samples per span)
- 3D polylines are projected to 2D for length calculation
- XREF blocks are not expanded
- Binary DXF format not supported (ASCII only)

## Machine Runtime Estimation

Enter your machine speed in cm/min to estimate total runtime:

```
Runtime = Total Length (cm) ÷ Machine Speed (cm/min)
```

Example:
- Total vector length: 500 cm
- Machine speed: 50 cm/min
- Estimated runtime: 10 minutes

## Example Output

```
============================================================
DXF VECTOR LENGTH CALCULATION - AUDIT REPORT
============================================================

File: pattern.dxf
Units Detected: Millimeters
Total Entities Processed: 156

--- TOTAL VECTOR LENGTH ---
Centimeters: 1234.5678 cm
Millimeters: 12345.6780 mm
Meters:      12.3457 m
Inches:      486.0503"
Feet:        40.5042'
Yards:       13.5014 yd

--- ENTITY BREAKDOWN ---
LINE: 89 entities, 456.78 cm
ARC: 34 entities, 234.56 cm
CIRCLE: 12 entities, 123.45 cm
LWPOLYLINE: 21 entities, 419.78 cm
============================================================
```

## Technical Details

### DXF Parsing
- Parses ASCII DXF format (versions R12 through 2018)
- Extracts ENTITIES section
- Reads HEADER for unit detection ($INSUNITS variable)
- Handles group codes according to Autodesk DXF specification

### Unit Detection
The tool reads the `$INSUNITS` header variable:
- 1 = Inches
- 4 = Millimeters (default if not specified)
- 5 = Centimeters
- 6 = Meters
- And more...

### Calculation Formulas

All formulas are implemented identically in both JavaScript and Python for cross-verification:

**Line:**
```javascript
length = Math.sqrt(dx*dx + dy*dy + dz*dz)
```

**Circle:**
```javascript
length = 2 * Math.PI * radius
```

**Arc:**
```javascript
length = radius * angleDifference  // angle in radians
```

**Polyline with Bulge:**
```javascript
bulge = tan(θ/4)
includedAngle = 4 * atan(|bulge|)
radius = chordLength / (2 * sin(includedAngle/2))
arcLength = radius * includedAngle
```

**Ellipse (Ramanujan):**
```javascript
h = ((a-b)²) / ((a+b)²)
circumference ≈ π(a+b)(1 + 3h/(10+√(4-3h)))
```

## License

MIT License - Free for commercial and personal use

## Support

For issues or feature requests, please check the repository issues or contact the development team.
