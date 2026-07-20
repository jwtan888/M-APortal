/**
 * Vector Length Calculator Application
 * Batch processing for DXF, PDF, SVG files — outputs only cm
 */

// Global state
let uploadedFiles = [];
let calculationResults = [];
window.calculationResults = calculationResults;

// DOM Elements
const dropZone = document.getElementById('dropZone');
const fileInput = document.getElementById('fileInput');
const fileList = document.getElementById('fileList');
const calculateBtn = document.getElementById('calculateBtn');
const loading = document.getElementById('loading');
const loadingText = document.getElementById('loadingText');
const results = document.getElementById('results');
const errorDiv = document.getElementById('error');
const errorMessage = document.getElementById('errorMessage');

// Initialize event listeners
function init() {
    dropZone.addEventListener('click', () => fileInput.click());
    fileInput.addEventListener('change', handleFileSelect);

    dropZone.addEventListener('dragover', (e) => {
        e.preventDefault();
        dropZone.classList.add('dragover');
    });

    dropZone.addEventListener('dragleave', () => {
        dropZone.classList.remove('dragover');
    });

    dropZone.addEventListener('drop', handleDrop);

    calculateBtn.addEventListener('click', calculateAll);

    document.getElementById('clearBtn').addEventListener('click', clearAll);

    // Global click-to-copy for .copy-val elements
    document.body.addEventListener('click', (e) => {
        const el = e.target.closest('.copy-val');
        if (!el) return;
        const val = el.dataset.copy;
        if (!val) return;
        navigator.clipboard.writeText(val).then(() => {
            const orig = el.textContent;
            el.textContent = '✓ copied';
            setTimeout(() => { el.textContent = orig; }, 800);
        });
    });
}

/**
 * Handle file drop
 */
function handleDrop(e) {
    e.preventDefault();
    dropZone.classList.remove('dragover');

    const files = Array.from(e.dataTransfer.files);
    addFiles(files);
}

/**
 * Handle file selection via input
 */
function handleFileSelect(e) {
    const files = Array.from(e.target.files);
    addFiles(files);
    fileInput.value = ''; // Reset so same file can be re-selected
}

/**
 * Add files to the queue
 */
function addFiles(files) {
    const allowedExtensions = ['.dxf', '.pdf', '.svg'];

    files.forEach(file => {
        const ext = '.' + file.name.split('.').pop().toLowerCase();
        if (!allowedExtensions.includes(ext)) {
            return;
        }
        // Avoid duplicates
        if (!uploadedFiles.find(f => f.name === file.name && f.size === file.size)) {
            uploadedFiles.push(file);
        }
    });

    renderFileList();
}

/**
 * Render the file list
 */
function renderFileList() {
    if (uploadedFiles.length === 0) {
        fileList.classList.add('hidden');
        calculateBtn.classList.add('hidden');
        return;
    }

    fileList.classList.remove('hidden');
    calculateBtn.classList.remove('hidden');
    results.classList.add('hidden');

    fileList.innerHTML = uploadedFiles.map((file, i) => {
        const ext = file.name.split('.').pop().toUpperCase();
        return `
            <div class="flex items-center justify-between p-3 bg-gray-50 rounded-lg file-card">
                <div class="flex items-center gap-3">
                    <span class="text-xs font-bold px-2 py-1 rounded bg-blue-100 text-blue-700">${ext}</span>
                    <span class="text-sm font-medium">${file.name}</span>
                </div>
                <div class="flex items-center gap-3">
                    <span class="text-xs text-gray-500">${formatFileSize(file.size)}</span>
                    <button onclick="removeFile(${i})" class="text-red-500 hover:text-red-700 text-sm">✕</button>
                </div>
            </div>
        `;
    }).join('');
}

/**
 * Remove a file from the queue
 */
function removeFile(index) {
    uploadedFiles.splice(index, 1);
    renderFileList();
}

/**
 * Calculate vector lengths for all files
 */
async function calculateAll() {
    if (uploadedFiles.length === 0) return;

    calculationResults = [];
    window.calculationResults = calculationResults;
    showLoading(true, 'Processing files...');
    hideError();

    try {
        for (let i = 0; i < uploadedFiles.length; i++) {
            const file = uploadedFiles[i];
            loadingText.textContent = `Processing ${i + 1}/${uploadedFiles.length}: ${file.name}`;

            const result = await processFile(file);
            calculationResults.push(result);
        }

        showLoading(false);
        displayResults();
    } catch (err) {
        console.error('Error processing files:', err);
        showError('Failed to process files: ' + err.message);
        showLoading(false);
    }
}

/**
 * Process a single file based on its type
 */
async function processFile(file) {
    const ext = file.name.split('.').pop().toLowerCase();
    const content = await readFileAsTextOrArrayBuffer(file, ext);

    let entities = [];
    let units = 'mm';

    if (ext === 'dxf') {
        const parser = new DXFParser();
        const dxfData = parser.parse(content);
        // Filter out TEXT, MTEXT, and other non-vector entities
        entities = dxfData.entities.filter(e => !isTextEntity(e.type));
        units = parser.units;
        // Inject layer colors into entities for ByLayer color resolution
        const layerColors = parser.layerColors || {};
        entities.forEach(e => {
            if (e.layer && layerColors[e.layer] && (e.color === 0 || e.color === 256 || e.color === undefined)) {
                e.color = layerColors[e.layer].color;
                if (layerColors[e.layer].trueColor && !e.trueColor) {
                    e.trueColor = layerColors[e.layer].trueColor;
                }
            }
        });
    } else if (ext === 'pdf') {
        const parser = new PDFVectorParser();
        entities = await parser.parse(content);
        units = 'pt'; // PDF uses points
        if (entities.length === 0 && parser._opsCount > 0) {
            // Diagnostic: PDF had operators but no vectors extracted
            // This means vectors are encoded as outlined text or form XObjects
            // Fallback: try raw content stream parsing
            entities = await parser.parseRawContent(content);
            if (entities.length === 0 && parser._opsCount > 0) {
                // Still no vectors — the PDF has operators but they're not path-based
                // Suggest SVG export from CorelDRAW
            }
        }
    } else if (ext === 'svg') {
        const parser = new SVGParser();
        entities = parser.parse(content);
        units = parser.units || 'px';
    }

    const lengths = calculateEntityLengths(entities, units);

    return {
        fileName: file.name,
        fileType: ext.toUpperCase(),
        totalLengthCm: lengths.totalCm,
        entityCount: lengths.count,
        breakdown: lengths.breakdown,
        entities: lengths.entities,
        rawEntities: entities,
        units: units,
        error: lengths.error
    };
}

/**
 * Check if entity type is text (to be excluded)
 */
function isTextEntity(type) {
    const textTypes = ['TEXT', 'MTEXT', 'ATTRIB', 'ATTDEF'];
    return textTypes.includes(type.toUpperCase());
}

/**
 * Calculate lengths for entities, converting to cm
 */
function calculateEntityLengths(entities, sourceUnit) {
    const result = {
        totalCm: 0,
        count: 0,
        breakdown: {},
        entities: [],
        error: null
    };

    if (!entities || entities.length === 0) {
        // Build diagnostic info
        const allTypes = [];
        if (ext === 'dxf') {
            const parser2 = new DXFParser();
            const d2 = parser2.parse(content);
            d2.entities.forEach(e => allTypes.push(e.type));
        }
        const uniqueTypes = [...new Set(allTypes)];
        result.error = uniqueTypes.length > 0
            ? `No measurable vectors found. Entity types in file: ${uniqueTypes.join(', ')}`
            : 'No vector entities found — file may contain only text/raster content';
        return result;
    }

    let idx = 0;
    entities.forEach((entity, rawIdx) => {
        const lengthInSourceUnit = getEntityLength(entity);

        if (lengthInSourceUnit > 0) {
            idx++;
            const lengthCm = convertToCm(lengthInSourceUnit, sourceUnit);
            result.totalCm += lengthCm;
            result.count++;

            if (!result.breakdown[entity.type]) {
                result.breakdown[entity.type] = { count: 0, lengthCm: 0 };
            }
            result.breakdown[entity.type].count++;
            result.breakdown[entity.type].lengthCm += lengthCm;

            result.entities.push({
                index: idx,
                rawIndex: rawIdx,
                type: entity.type,
                layer: entity.layer || entity.layerName || '-',
                lengthCm: lengthCm,
                color: entity.color !== undefined ? entity.color : null,
                trueColor: entity.trueColor !== undefined ? entity.trueColor : null
            });
        }
    });

    return result;
}

/**
 * Convert a length value from source unit to cm
 */
function convertToCm(length, unit) {
    // Conversion factors to cm
    const toCm = {
        'mm': 0.1,
        'cm': 1,
        'm': 100,
        'in': 2.54,
        'inches': 2.54,
        'ft': 30.48,
        'feet': 30.48,
        'yd': 91.44,
        'pt': 2.54 / 72,    // 1 point = 1/72 inch
        'px': 2.54 / 96,    // 1 px = 1/96 inch (CSS px)
        'Unitless': 0.1,    // Assume mm
    };

    const factor = toCm[unit] || toCm['Unitless'];
    return length * factor;
}

/**
 * Get length of a single entity in source units
 */
function getEntityLength(entity) {
    switch (entity.type) {
        case 'LINE':
            return lineLength(entity);
        case 'CIRCLE':
            return 2 * Math.PI * (entity.radius || 0);
        case 'ARC':
            return arcLength(entity);
        case 'POLYLINE':
        case 'LWPOLYLINE':
            return polylineLength(entity);
        case 'SPLINE':
            return splineLength(entity);
        case 'ELLIPSE':
            return ellipseLength(entity);
        case 'RECT':
            return (entity.width || 0) * 2 + (entity.height || 0) * 2;
        case 'PATH':
            return pathLength(entity);
        case 'HATCH':
            return hatchLength(entity);
        case 'POINT':
            return 0;
        default:
            if (entity.x !== undefined && entity.y !== undefined &&
                entity.x1 !== undefined && entity.y1 !== undefined) {
                return lineLength(entity);
            }
            return 0;
    }
}

function lineLength(e) {
    const dx = (e.x1 || 0) - (e.x || 0);
    const dy = (e.y1 || 0) - (e.y || 0);
    const dz = (e.z1 || 0) - (e.z || 0);
    return Math.sqrt(dx * dx + dy * dy + dz * dz);
}

function arcLength(e) {
    const r = e.radius || 0;
    let sa = e.startAngle || 0;
    let ea = e.endAngle || 0;
    while (sa < 0) sa += 2 * Math.PI;
    while (ea < 0) ea += 2 * Math.PI;
    let diff = ea - sa;
    if (diff < 0) diff += 2 * Math.PI;
    return r * diff;
}

function polylineLength(e) {
    const verts = e.vertices || [];
    if (verts.length < 2) return 0;

    let total = 0;
    for (let i = 0; i < verts.length - 1; i++) {
        const v1 = verts[i], v2 = verts[i + 1];
        if (v1.bulge && Math.abs(v1.bulge) > 0.0001) {
            total += bulgeArcLength(v1, v2);
        } else {
            total += lineLength({ x: v1.x, y: v1.y, x1: v2.x, y1: v2.y });
        }
    }

    // Check if closed
    const isClosed = e.closed || (verts.length > 2 &&
        Math.abs(verts[0].x - verts[verts.length - 1].x) < 0.0001 &&
        Math.abs(verts[0].y - verts[verts.length - 1].y) < 0.0001);

    if (isClosed && verts.length > 2) {
        const v1 = verts[verts.length - 1], v2 = verts[0];
        if (v1.bulge && Math.abs(v1.bulge) > 0.0001) {
            total += bulgeArcLength(v1, v2);
        } else {
            total += lineLength({ x: v1.x, y: v1.y, x1: v2.x, y1: v2.y });
        }
    }
    return total;
}

function bulgeArcLength(v1, v2) {
    const bulge = v1.bulge;
    const chord = Math.sqrt((v2.x - v1.x) ** 2 + (v2.y - v1.y) ** 2);
    if (chord < 0.0001) return 0;
    const angle = 4 * Math.atan(Math.abs(bulge));
    const r = chord / (2 * Math.sin(angle / 2));
    return r * angle;
}

function splineLength(e) {
    const pts = e.controlPoints || [];
    if (pts.length < 2) return 0;
    let total = 0;
    const samples = 100;
    for (let i = 0; i < pts.length - 1; i++) {
        for (let j = 0; j < samples; j++) {
            const t1 = j / samples, t2 = (j + 1) / samples;
            const p1 = pts[i], p2 = pts[i + 1];
            const x1 = p1.x + (p2.x - p1.x) * t1;
            const y1 = p1.y + (p2.y - p1.y) * t1;
            const x2 = p1.x + (p2.x - p1.x) * t2;
            const y2 = p1.y + (p2.y - p1.y) * t2;
            total += Math.sqrt((x2 - x1) ** 2 + (y2 - y1) ** 2);
        }
    }
    return total;
}

function ellipseLength(e) {
    const mx = e.majorAxisX || 0, my = e.majorAxisY || 0;
    const majorLen = Math.sqrt(mx * mx + my * my);
    const ratio = e.axisRatio || 1;
    const minorLen = majorLen * ratio;
    const a = majorLen / 2, b = minorLen / 2;
    const h = ((a - b) ** 2) / ((a + b) ** 2);
    const circ = Math.PI * (a + b) * (1 + (3 * h) / (10 + Math.sqrt(4 - 3 * h)));

    const sp = e.startParam || 0, ep = e.endParam || 2 * Math.PI;
    const range = ep - sp;
    if (range < 2 * Math.PI - 0.0001) {
        return circ * (range / (2 * Math.PI));
    }
    return circ;
}

/**
 * Calculate HATCH boundary length
 * HATCH entities contain boundary edges (lines, arcs, ellipses, splines)
 */
function hatchLength(e) {
    const edges = e.boundaryEdges || [];
    if (edges.length === 0) return 0;

    let total = 0;
    edges.forEach(edge => {
        switch (edge.edgeType) {
            case 1: // LINE edge
                total += lineLength({ x: edge.x, y: edge.y, x1: edge.x1, y1: edge.y1 });
                break;
            case 2: // ARC edge
                total += arcLength({
                    radius: edge.radius,
                    startAngle: (edge.startAngle || 0) * Math.PI / 180, // DXF degrees → radians
                    endAngle: (edge.endAngle || 360) * Math.PI / 180,
                    cx: edge.x, cy: edge.y
                });
                break;
            case 3: // ELLIPSE edge
                total += ellipseLength({
                    majorAxisX: edge.majorAxisX,
                    majorAxisY: edge.majorAxisY,
                    axisRatio: edge.axisRatio || 1,
                    cx: edge.x, cy: edge.y
                });
                break;
            case 4: // SPLINE edge — approximate as polyline through fit points
                // DXF hatch spline edges store fit points at group 10/20 after degree/knots
                // Approximate: treat as line through available points
                break;
        }
    });
    return total;
}

/**
 * Calculate length of an SVG path's 'd' attribute
 */
function pathLength(entity) {
    if (!entity.pathSegments || entity.pathSegments.length === 0) return 0;
    let total = 0;
    let cx = entity.pathSegments[0].x || 0;
    let cy = entity.pathSegments[0].y || 0;

    for (const seg of entity.pathSegments) {
        switch (seg.command) {
            case 'M':
                cx = seg.x; cy = seg.y;
                break;
            case 'L':
            case 'H':
            case 'V': {
                const nx = seg.x !== undefined ? seg.x : cx;
                const ny = seg.y !== undefined ? seg.y : cy;
                total += Math.sqrt((nx - cx) ** 2 + (ny - cy) ** 2);
                cx = nx; cy = ny;
                break;
            }
            case 'C': {
                // Cubic bezier — approximate with line segments
                const pts = bezierPoints(cx, cy, seg.x1, seg.y1, seg.x2, seg.y2, seg.x, seg.y, 20);
                for (let i = 0; i < pts.length - 1; i++) {
                    total += Math.sqrt((pts[i+1].x - pts[i].x) ** 2 + (pts[i+1].y - pts[i].y) ** 2);
                }
                cx = seg.x; cy = seg.y;
                break;
            }
            case 'Q': {
                // Quadratic bezier
                const pts = quadBezierPoints(cx, cy, seg.x1, seg.y1, seg.x, seg.y, 20);
                for (let i = 0; i < pts.length - 1; i++) {
                    total += Math.sqrt((pts[i+1].x - pts[i].x) ** 2 + (pts[i+1].y - pts[i].y) ** 2);
                }
                cx = seg.x; cy = seg.y;
                break;
            }
            case 'A': {
                // Arc
                total += arcSegmentLength(cx, cy, seg);
                cx = seg.x; cy = seg.y;
                break;
            }
            case 'Z':
                // Close path — back to start
                break;
        }
    }
    return total;
}

function bezierPoints(x0, y0, x1, y1, x2, y2, x3, y3, samples) {
    const pts = [];
    for (let i = 0; i <= samples; i++) {
        const t = i / samples;
        const mt = 1 - t;
        pts.push({
            x: mt*mt*mt*x0 + 3*mt*mt*t*x1 + 3*mt*t*t*x2 + t*t*t*x3,
            y: mt*mt*mt*y0 + 3*mt*mt*t*y1 + 3*mt*t*t*y2 + t*t*t*y3
        });
    }
    return pts;
}

function quadBezierPoints(x0, y0, x1, y1, x2, y2, samples) {
    const pts = [];
    for (let i = 0; i <= samples; i++) {
        const t = i / samples;
        const mt = 1 - t;
        pts.push({
            x: mt*mt*x0 + 2*mt*t*x1 + t*t*x2,
            y: mt*mt*y0 + 2*mt*t*y1 + t*t*y2
        });
    }
    return pts;
}

function arcSegmentLength(cx, cy, seg) {
    const rx = seg.rx || 0, ry = seg.ry || 0;
    // Approximate ellipse arc with straight line
    const dist = Math.sqrt((seg.x - cx) ** 2 + (seg.y - cy) ** 2);
    return dist;
}

/**
 * Display results
 */
function displayResults() {
    const fileResultsDiv = document.getElementById('fileResults');
    fileResultsDiv.innerHTML = '';

    let grandTotalCm = 0;
    let grandTotalEntities = 0;
    const validResults = calculationResults.filter(r => !r.error);

    // --- Summary Table ---
    if (validResults.length > 0) {
        let summaryHTML = '<div class="bg-white rounded-xl shadow-lg overflow-hidden mb-4"><div class="p-4 bg-gradient-to-r from-blue-600 to-purple-600 text-white"><h3 class="text-lg font-bold">All Files Summary</h3></div>';
        summaryHTML += '<div class="overflow-x-auto"><table class="w-full text-sm"><thead class="bg-gray-50 border-b"><tr><th class="text-left p-3 font-semibold">File</th><th class="text-left p-3 font-semibold">Type</th><th class="text-right p-3 font-semibold">Vectors</th><th class="text-right p-3 font-semibold">Length (cm)</th></tr></thead><tbody>';

        for (const r of validResults) {
            summaryHTML += `<tr class="border-b border-gray-100 hover:bg-blue-50">
                <td class="p-3 font-medium text-gray-800 cursor-pointer copy-val hover:underline" data-copy="${escHtml(r.fileName)}">${escHtml(r.fileName)}</td>
                <td class="p-3"><span class="text-xs font-bold px-2 py-0.5 rounded bg-blue-100 text-blue-700">${r.fileType}</span></td>
                <td class="p-3 text-right text-gray-600">${r.entityCount}</td>
                <td class="p-3 text-right font-semibold text-purple-700 cursor-pointer copy-val hover:underline" data-copy="${r.totalLengthCm.toFixed(4)}">${r.totalLengthCm.toFixed(4)}</td>
            </tr>`;
            grandTotalCm += r.totalLengthCm;
            grandTotalEntities += r.entityCount;
        }

        summaryHTML += `<tr class="bg-purple-50 font-bold"><td class="p-3" colspan="3">Grand Total</td><td class="p-3 text-right text-purple-700 cursor-pointer copy-val hover:underline" data-copy="${grandTotalCm.toFixed(4)}">${grandTotalCm.toFixed(4)}</td></tr>`;
        summaryHTML += '</tbody></table></div></div>';
        fileResultsDiv.innerHTML += summaryHTML;
    }

    // --- Per-File Detail Cards ---
    const colorPalette = [
        'border-l-blue-400', 'border-l-green-400', 'border-l-purple-400',
        'border-l-orange-400', 'border-l-teal-400', 'border-l-pink-400'
    ];

    calculationResults.forEach((result, index) => {
        if (result.error) {
            fileResultsDiv.innerHTML += `
                <div class="bg-white rounded-xl shadow-lg p-6 border-l-4 border-l-red-400">
                    <div class="flex justify-between items-start">
                        <div>
                            <h3 class="text-lg font-bold text-gray-800 cursor-pointer copy-val hover:underline" data-copy="${escHtml(result.fileName)}">${escHtml(result.fileName)}</h3>
                            <p class="text-red-600 text-sm mt-1">${result.error}</p>
                        </div>
                        <span class="text-xs font-bold px-2 py-1 rounded bg-gray-100 text-gray-600">${result.fileType}</span>
                    </div>
                </div>
            `;
            return;
        }

        const card = document.createElement('div');
        const borderColor = colorPalette[index % colorPalette.length];
        card.className = `bg-white rounded-xl shadow-lg p-6 file-card border-l-4 ${borderColor}`;

        const canvasId = `preview-${index}`;

        let breakdownHTML = '';
        result.entities.forEach((entity, ei) => {
            const colorHex = entityColorHex(entity);
            const onclickHandler = `blinkHighlightEntity('preview-${index}', window.calculationResults[${index}].rawEntities, ${entity.rawIndex !== undefined ? entity.rawIndex : ei})`;
            breakdownHTML += `
                <div class="flex justify-between items-center py-1.5 border-b border-gray-100 last:border-0 hover:bg-blue-50">
                    <span class="text-sm text-gray-600 flex items-center gap-2 cursor-pointer" onclick="${onclickHandler}">
                        <span class="inline-block w-4 h-4 rounded-full border border-gray-300 flex-shrink-0" style="background:${colorHex};box-shadow:0 0 2px ${colorHex}"></span>
                        #${entity.index} <span class="text-gray-400">${entity.type}</span> <span class="text-gray-400 text-xs">(${escHtml(entity.layer)})</span>
                    </span>
                    <span class="text-sm font-semibold text-blue-600 cursor-pointer copy-val hover:underline" data-copy="${entity.lengthCm.toFixed(4)}">${entity.lengthCm.toFixed(4)} cm</span>
                </div>
            `;
        });
        // File subtotal row
        breakdownHTML += `
            <div class="flex justify-between items-center pt-2 mt-2 border-t-2 border-gray-300">
                <span class="text-sm font-bold text-gray-800">File Total</span>
                <span class="text-sm font-bold text-purple-700 cursor-pointer copy-val hover:underline" data-copy="${result.totalLengthCm.toFixed(4)}">${result.totalLengthCm.toFixed(4)} cm</span>
            </div>
        `;

        card.innerHTML = `
            <div class="flex justify-between items-start mb-4">
                <div>
                    <h3 class="text-lg font-bold text-gray-800 cursor-pointer copy-val hover:underline" data-copy="${escHtml(result.fileName)}">${escHtml(result.fileName)}</h3>
                    <p class="text-sm text-gray-500">${result.entityCount} vectors processed</p>
                </div>
                <div class="text-right">
                    <span class="text-xs font-bold px-2 py-1 rounded bg-blue-100 text-blue-700">${result.fileType}</span>
                    <p class="text-2xl font-bold text-purple-700 mt-2 cursor-pointer copy-val hover:underline" data-copy="${result.totalLengthCm.toFixed(4)}">${result.totalLengthCm.toFixed(4)} cm</p>
                </div>
            </div>
            <div class="flex flex-col md:flex-row gap-4 mb-4">
                <div class="flex-shrink-0 bg-gray-100 rounded-lg p-1 flex items-center justify-center" style="min-width:200px;min-height:180px;">
                    <canvas id="${canvasId}" class="w-full h-full" style="max-width:200px;max-height:180px;"></canvas>
                </div>
                <div class="flex-1 bg-gray-50 rounded-lg p-4 max-h-96 overflow-y-auto">
                    <h4 class="text-sm font-semibold text-gray-700 mb-2">Entity Breakdown</h4>
                    ${breakdownHTML}
                </div>
            </div>
        `;

        fileResultsDiv.appendChild(card);

        setTimeout(() => {
            renderVectorPreview(canvasId, result.rawEntities);
        }, 50);
    });

    // Process errors for grand total
    if (validResults.length === 0 && calculationResults.every(r => r.error)) {
        grandTotalCm = 0;
        grandTotalEntities = 0;
    }

    // Grand total
    const grandTotalEl = document.getElementById('grandTotalCm');
    grandTotalEl.textContent = grandTotalCm.toFixed(4) + ' cm';
    grandTotalEl.classList.add('cursor-pointer', 'copy-val', 'hover:underline');
    grandTotalEl.dataset.copy = grandTotalCm.toFixed(4);
    document.getElementById('grandTotalSummary').textContent =
        `${validResults.length} file(s) · ${grandTotalEntities} total vectors`;

    results.classList.remove('hidden');
}

/**
 * Escape HTML to prevent XSS in file names
 */
function escHtml(str) {
    const div = document.createElement('div');
    div.textContent = str;
    return div.innerHTML;
}

/**
 * Resolve entity color to CSS hex — prefers True Color (RGB24) over ACI index
 */
function entityColorHex(entity) {
    // True Color (group 420) — 24-bit RGB: R*65536 + G*256 + B
    if (entity.trueColor) {
        const tc = Math.abs(entity.trueColor);
        const r = (tc >> 16) & 0xFF;
        const g = (tc >> 8) & 0xFF;
        const b = tc & 0xFF;
        return '#' + r.toString(16).padStart(2,'0') + g.toString(16).padStart(2,'0') + b.toString(16).padStart(2,'0');
    }
    // ACI color index (group 62)
    return aciColor(entity.color);
}

/**
 * AutoCAD Color Index (ACI) lookup — full 256 color table
 */
function aciColor(num) {
    if (num === null || num === undefined || num === 0 || num === 256) return '#6b7280'; // ByLayer/default
    const C = {
        1:'#ff0000',2:'#ffff00',3:'#00ff00',4:'#00ffff',5:'#0000ff',6:'#ff00ff',7:'#ffffff',
        8:'#414141',9:'#808080',
        10:'#ff0000',11:'#ffaaaa',12:'#cc0000',13:'#cc7777',14:'#990000',15:'#994c4c',
        16:'#7f0000',17:'#7f3f3f',18:'#4c0000',19:'#4c2626',
        20:'#ff7f00',21:'#ffbf7f',22:'#cc6600',23:'#cc9966',24:'#994c00',25:'#99724c',
        26:'#7f3f00',27:'#7f5f3f',28:'#4c2600',29:'#4c3926',
        30:'#ffff00',31:'#ffff7f',32:'#cccc00',33:'#cccc66',34:'#999900',35:'#99994c',
        36:'#7f7f00',37:'#7f7f3f',38:'#4c4c00',39:'#4c4c26',
        40:'#00ff00',41:'#7fff7f',42:'#00cc00',43:'#66cc66',44:'#009900',45:'#4c994c',
        46:'#007f00',47:'#3f7f3f',48:'#004c00',49:'#264c26',
        50:'#00ffff',51:'#7fffff',52:'#00cccc',53:'#66cccc',54:'#009999',55:'#4c9999',
        56:'#007f7f',57:'#3f7f7f',58:'#004c4c',59:'#264c26',
        60:'#0000ff',61:'#7f7fff',62:'#0000cc',63:'#6666cc',64:'#000099',65:'#4c4c99',
        66:'#00007f',67:'#3f3f7f',68:'#00004c',69:'#26264c',
        70:'#ff00ff',71:'#ff7fff',72:'#cc00cc',73:'#cc66cc',74:'#990099',75:'#994c99',
        76:'#7f007f',77:'#7f3f7f',78:'#4c004c',79:'#4c264c',
        80:'#7f0000',81:'#bf7f7f',82:'#cc0000',83:'#cc7f7f',84:'#990000',85:'#994c4c',
        86:'#7f0000',87:'#7f3f3f',88:'#4c0000',89:'#4c2626',
        90:'#ff0000',91:'#ff7f7f',92:'#cc0000',93:'#cc6666',94:'#990000',95:'#994c4c',
        96:'#7f0000',97:'#7f3f3f',98:'#4c0000',99:'#4c2626',
        140:'#00ccff',
        240:'#414141',241:'#808080',242:'#b3b3b3',243:'#cccccc',244:'#e6e6e6',
        245:'#f2f2f2',246:'#ffffff',247:'#e6e6e6',248:'#cccccc',249:'#b3b3b3',
        250:'#808080',251:'#595959',252:'#414141',253:'#2c2c2c',254:'#1a1a1a',255:'#000000'
    };
    if (C[num]) return C[num];
    // Compute remaining ACI colors 100-239 algorithmically
    // These are interleaved shades from the 6 hue families
    if (num >= 100 && num <= 239) {
        const idx = num - 100;
        const hueIdx = idx % 6;
        const shadeIdx = Math.floor(idx / 6);
        const hues = [
            [255,0,0],[255,127,0],[255,255,0],[0,255,0],[0,255,255],[0,0,255]
        ];
        const h = hues[hueIdx];
        const factor = 1 - (shadeIdx * 0.1);
        const r = Math.round(h[0] * factor), g = Math.round(h[1] * factor), b = Math.round(h[2] * factor);
        return '#' + r.toString(16).padStart(2,'0') + g.toString(16).padStart(2,'0') + b.toString(16).padStart(2,'0');
    }
    return '#6b7280';
}

/**
 * Render vector artwork preview onto a canvas
 */
function renderVectorPreview(canvasId, entities, highlightIdx) {
    const canvas = document.getElementById(canvasId);
    if (!canvas || !entities || entities.length === 0) return;

    const ctx = canvas.getContext('2d');
    const dpr = window.devicePixelRatio || 1;
    const w = 200, h = 180;
    canvas.width = w * dpr;
    canvas.height = h * dpr;
    canvas.style.width = w + 'px';
    canvas.style.height = h + 'px';
    ctx.scale(dpr, dpr);

    const bb = getEntityBounds(entities);
    if (!bb) {
        ctx.fillStyle = '#999';
        ctx.font = '12px sans-serif';
        ctx.textAlign = 'center';
        ctx.fillText('(no geometry)', w / 2, h / 2);
        return;
    }

    // Crop to vector content — use percentile bounds to ignore artboard/outliers
    const cropped = cropToBounds(entities, bb, 0.95);

    const padding = 12;
    const drawW = w - padding * 2;
    const drawH = h - padding * 2;
    const dataW = cropped.maxX - cropped.minX || 1;
    const dataH = cropped.maxY - cropped.minY || 1;

    const scale = Math.min(drawW / dataW, drawH / dataH);
    const offsetX = padding + (drawW - dataW * scale) / 2 - cropped.minX * scale;
    const offsetY = padding + (drawH - dataH * scale) / 2 + cropped.maxY * scale;

    function tx(x) { return x * scale + offsetX; }
    function ty(y) { return offsetY - y * scale; }

    const typeColors = {
        LINE: '#2563eb', CIRCLE: '#dc2626', ARC: '#d97706',
        POLYLINE: '#059669', LWPOLYLINE: '#059669',
        SPLINE: '#7c3aed', ELLIPSE: '#db2777',
        RECT: '#0891b2', PATH: '#4f46e5', HATCH: '#ea580c'
    };

    ctx.clearRect(0, 0, w, h);

    entities.forEach((e, idx) => {
        const isHighlighted = idx === highlightIdx;
        ctx.strokeStyle = isHighlighted ? '#ff0000' : entityColorHex(e);
        ctx.lineWidth = isHighlighted ? Math.max(2, 3 / scale) : Math.max(0.6, 1.2 / scale);
        ctx.beginPath();
        drawEntity(ctx, e, tx, ty);
        ctx.stroke();
    });
}

/**
 * Blink highlight a specific entity in the preview
 */
function blinkHighlightEntity(canvasId, entities, entityIdx) {
    let blinkCount = 0;
    const maxBlinks = 6;
    const interval = setInterval(() => {
        const showHighlight = blinkCount % 2 === 0;
        renderVectorPreview(canvasId, entities, showHighlight ? entityIdx : -1);
        blinkCount++;
        if (blinkCount >= maxBlinks) {
            clearInterval(interval);
            renderVectorPreview(canvasId, entities, -1);
        }
    }, 200);
}

function getEntityBounds(entities) {
    let minX = Infinity, minY = Infinity, maxX = -Infinity, maxY = -Infinity;
    function add(x, y) {
        if (x < minX) minX = x;
        if (y < minY) minY = y;
        if (x > maxX) maxX = x;
        if (y > maxY) maxY = y;
    }
    entities.forEach(e => entityBounds(e, add));
    if (!isFinite(minX)) return null;
    return { minX, minY, maxX, maxY };
}

/**
 * Crop bounding box to vector content using percentile — ignores outlier coords
 * (e.g. artboard origins at -100000 while drawing is at 10000-50000)
 */
function cropToBounds(entities, bb, percentile) {
    // Collect all coordinate points
    const xs = [], ys = [];
    entities.forEach(e => collectCoords(e, xs, ys));

    if (xs.length === 0 || ys.length === 0) return bb;

    // Sort and find percentile bounds
    xs.sort((a, b) => a - b);
    ys.sort((a, b) => a - b);

    const lo = Math.floor(xs.length * (1 - percentile) / 2);
    const hi = Math.ceil(xs.length * (1 + percentile) / 2) - 1;

    const minX = xs[lo];
    const maxX = xs[Math.min(hi, xs.length - 1)];
    const minY = ys[lo];
    const maxY = ys[Math.min(hi, ys.length - 1)];

    // If the cropped area is too small relative to full bounds, use full bounds
    const croppedW = maxX - minX;
    const croppedH = maxY - minY;
    const fullW = bb.maxX - bb.minX;
    const fullH = bb.maxY - bb.minY;

    if (croppedW < fullW * 0.3 || croppedH < fullH * 0.3) {
        // Cropped area is too small — likely all entities are clustered
        // Use the full bounding box but add 10% margin
        return {
            minX: minX - croppedW * 0.1,
            minY: minY - croppedH * 0.1,
            maxX: maxX + croppedW * 0.1,
            maxY: maxY + croppedH * 0.1
        };
    }

    return { minX, minY, maxX, maxY };
}

function collectCoords(e, xs, ys) {
    function push(x, y) { xs.push(x); ys.push(y); }
    switch (e.type) {
        case 'LINE': push(e.x || 0, e.y || 0); push(e.x1 || 0, e.y1 || 0); break;
        case 'CIRCLE': push(e.cx || 0, e.cy || 0); break;
        case 'ARC': push(e.cx || 0, e.cy || 0); break;
        case 'POLYLINE': case 'LWPOLYLINE': (e.vertices || []).forEach(v => push(v.x, v.y)); break;
        case 'SPLINE': (e.controlPoints || []).forEach(p => push(p.x, p.y)); break;
        case 'ELLIPSE': push(e.cx || 0, e.cy || 0); break;
        case 'RECT': push(e.x || 0, e.y || 0); push((e.x || 0) + (e.width || 0), (e.y || 0) + (e.height || 0)); break;
        case 'HATCH': (e.boundaryEdges || []).forEach(edge => {
            if (edge.edgeType === 1) { push(edge.x, edge.y); push(edge.x1, edge.y1); }
            else if (edge.edgeType === 2) { push(edge.x, edge.y); }
            else if (edge.edgeType === 3) { push(edge.x, edge.y); }
        }); break;
        case 'PATH': (e.pathSegments || []).forEach(seg => {
            if (seg.x !== undefined) push(seg.x, seg.y || 0);
        }); break;
        default:
            if (e.x !== undefined) push(e.x, e.y || 0);
            if (e.x1 !== undefined) push(e.x1, e.y1 || 0);
            break;
    }
}

function entityBounds(e, add) {
    switch (e.type) {
        case 'LINE':
            add(e.x || 0, e.y || 0);
            add(e.x1 || 0, e.y1 || 0);
            break;
        case 'CIRCLE':
            add((e.cx || 0) - (e.radius || 0), (e.cy || 0) - (e.radius || 0));
            add((e.cx || 0) + (e.radius || 0), (e.cy || 0) + (e.radius || 0));
            break;
        case 'ARC':
            add((e.cx || 0) - (e.radius || 0), (e.cy || 0) - (e.radius || 0));
            add((e.cx || 0) + (e.radius || 0), (e.cy || 0) + (e.radius || 0));
            break;
        case 'POLYLINE':
        case 'LWPOLYLINE':
            (e.vertices || []).forEach(v => add(v.x, v.y));
            break;
        case 'SPLINE':
            (e.controlPoints || []).forEach(p => add(p.x, p.y));
            break;
        case 'ELLIPSE': {
            const cx = e.cx || 0, cy = e.cy || 0;
            const mx = e.majorAxisX || 0, my = e.majorAxisY || 0;
            const r = Math.sqrt(mx * mx + my * my) / 2;
            add(cx - r, cy - r);
            add(cx + r, cy + r);
            break;
        }
        case 'RECT':
            add(e.x || 0, e.y || 0);
            add((e.x || 0) + (e.width || 0), (e.y || 0) + (e.height || 0));
            break;
        case 'HATCH':
            (e.boundaryEdges || []).forEach(edge => {
                if (edge.edgeType === 1) { add(edge.x, edge.y); add(edge.x1, edge.y1); }
                else if (edge.edgeType === 2) { add(edge.x, edge.y); }
                else if (edge.edgeType === 3) { add(edge.x, edge.y); }
            });
            break;
        case 'PATH':
            (e.pathSegments || []).forEach(seg => {
                if (seg.x !== undefined) add(seg.x, seg.y);
                if (seg.x1 !== undefined) add(seg.x1, seg.y1);
                if (seg.x2 !== undefined) add(seg.x2, seg.y2);
            });
            break;
        default:
            if (e.x !== undefined) add(e.x, e.y || 0);
            if (e.x1 !== undefined) add(e.x1, e.y1 || 0);
            break;
    }
}

function drawEntity(ctx, e, tx, ty) {
    switch (e.type) {
        case 'LINE':
            ctx.moveTo(tx(e.x || 0), ty(e.y || 0));
            ctx.lineTo(tx(e.x1 || 0), ty(e.y1 || 0));
            break;
        case 'CIRCLE': {
            const cx = tx(e.cx || 0), cy = ty(e.cy || 0);
            ctx.arc(cx, cy, (e.radius || 0) * ((tx(1) - tx(0)) || 1), 0, Math.PI * 2);
            break;
        }
        case 'ARC': {
            const cx = tx(e.cx || 0), cy = ty(e.cy || 0);
            const r = (e.radius || 0) * ((tx(1) - tx(0)) || 1);
            let sa = e.startAngle || 0, ea = e.endAngle || 0;
            while (sa < 0) sa += 2 * Math.PI;
            while (ea < 0) ea += 2 * Math.PI;
            ctx.arc(cx, cy, r, sa, ea, ea < sa);
            break;
        }
        case 'POLYLINE':
        case 'LWPOLYLINE': {
            const verts = e.vertices || [];
            if (verts.length === 0) break;
            ctx.moveTo(tx(verts[0].x), ty(verts[0].y));
            for (let i = 1; i < verts.length; i++) {
                const v = verts[i];
                if (v.bulge && Math.abs(v.bulge) > 0.0001) {
                    drawBulgeArc(ctx, verts[i - 1], v, tx, ty);
                } else {
                    ctx.lineTo(tx(v.x), ty(v.y));
                }
            }
            break;
        }
        case 'SPLINE': {
            const verts = typeof MarkerNesting !== 'undefined' && MarkerNesting.getEntityVerts
                ? MarkerNesting.getEntityVerts(e)
                : (e.controlPoints || []).map(p => ({ x: p.x, y: p.y }));
            if (verts.length === 0) break;
            ctx.moveTo(tx(verts[0].x), ty(verts[0].y));
            for (let i = 1; i < verts.length; i++) {
                ctx.lineTo(tx(verts[i].x), ty(verts[i].y));
            }
            break;
        }
        case 'ELLIPSE': {
            const cx = tx(e.cx || 0), cy = ty(e.cy || 0);
            const mx = e.majorAxisX || 0, my = e.majorAxisY || 0;
            const ratio = e.axisRatio || 1;
            const majorLen = Math.sqrt(mx * mx + my * my) / 2 * ((tx(1) - tx(0)) || 1);
            const minorLen = majorLen * ratio;
            const angle = Math.atan2(my, mx);
            ctx.ellipse(cx, cy, majorLen, minorLen, angle, 0, Math.PI * 2);
            break;
        }
        case 'RECT':
            ctx.rect(tx(e.x || 0), ty(e.y || 0), (e.width || 0) * ((tx(1) - tx(0)) || 1), (e.height || 0) * ((tx(1) - tx(0)) || 1));
            break;
        case 'HATCH':
            (e.boundaryEdges || []).forEach(edge => {
                if (edge.edgeType === 1) {
                    ctx.moveTo(tx(edge.x || 0), ty(edge.y || 0));
                    ctx.lineTo(tx(edge.x1 || 0), ty(edge.y1 || 0));
                } else if (edge.edgeType === 2) {
                    const cx = tx(edge.x || 0), cy = ty(edge.y || 0);
                    const r = (edge.radius || 0) * ((tx(1) - tx(0)) || 1);
                    let sa = (edge.startAngle || 0) * Math.PI / 180;
                    let ea = (edge.endAngle || 360) * Math.PI / 180;
                    ctx.arc(cx, cy, r, sa, ea, false);
                } else if (edge.edgeType === 3) {
                    const cx = tx(edge.x || 0), cy = ty(edge.y || 0);
                    const mx = edge.majorAxisX || 0, my = edge.majorAxisY || 0;
                    const ratio = edge.axisRatio || 1;
                    const majorLen = Math.sqrt(mx * mx + my * my) / 2 * ((tx(1) - tx(0)) || 1);
                    const minorLen = majorLen * ratio;
                    const angle = Math.atan2(my, mx);
                    ctx.ellipse(cx, cy, majorLen, minorLen, angle, 0, Math.PI * 2);
                }
            });
            break;
        case 'PATH':
            (e.pathSegments || []).forEach((seg, i) => {
                if (i === 0 || seg.command === 'M') {
                    ctx.moveTo(tx(seg.x || 0), ty(seg.y || 0));
                } else if (seg.command === 'L') {
                    ctx.lineTo(tx(seg.x), ty(seg.y));
                } else if (seg.command === 'C') {
                    ctx.bezierCurveTo(tx(seg.x1), ty(seg.y1), tx(seg.x2), ty(seg.y2), tx(seg.x), ty(seg.y));
                } else if (seg.command === 'Q') {
                    ctx.quadraticCurveTo(tx(seg.x1), ty(seg.y1), tx(seg.x), ty(seg.y));
                } else if (seg.command === 'Z') {
                    ctx.closePath();
                }
            });
            break;
        default:
            if (e.x !== undefined && e.y !== undefined && e.x1 !== undefined && e.y1 !== undefined) {
                ctx.moveTo(tx(e.x), ty(e.y));
                ctx.lineTo(tx(e.x1), ty(e.y1));
            }
            break;
    }
}

function drawBulgeArc(ctx, v1, v2, tx, ty) {
    const bulge = v1.bulge;
    const chord = Math.sqrt((v2.x - v1.x) ** 2 + (v2.y - v1.y) ** 2);
    if (chord < 0.0001) return;
    const angle = 4 * Math.atan(Math.abs(bulge));
    const r = chord / (2 * Math.sin(angle / 2));
    const mx = (v1.x + v2.x) / 2, my = (v1.y + v2.y) / 2;
    const dx = v2.x - v1.x, dy = v2.y - v1.y;
    const h = r * Math.cos(angle / 2);
    const px = mx + (-dy / chord) * h * Math.sign(bulge);
    const py = my + (dx / chord) * h * Math.sign(bulge);
    ctx.quadraticCurveTo(tx(px), ty(py), tx(v2.x), ty(v2.y));
}

/**
 * Export all results to CSV
 */
function exportToCSV() {
    if (calculationResults.length === 0) return;

    let csv = 'File,Entity #,Entity Type,Layer,Length (cm)\n';

    calculationResults.forEach(result => {
        result.entities.forEach(entity => {
            csv += `"${result.fileName}",${entity.index},"${entity.type}","${entity.layer}",${entity.lengthCm.toFixed(4)}\n`;
        });
        csv += `"${result.fileName}",,"TOTAL","—","${result.totalLengthCm.toFixed(4)}"\n`;
    });

    const grandTotal = calculationResults.reduce((sum, r) => sum + r.totalLengthCm, 0);
    csv += `\n"GRAND TOTAL",,"—","—","${grandTotal.toFixed(4)}"\n`;

    const blob = new Blob([csv], { type: 'text/csv' });
    downloadBlob(blob, 'vector-lengths-batch.csv');
}

/**
 * Clear all
 */
function clearAll() {
    uploadedFiles = [];
    calculationResults = [];
    renderFileList();
    results.classList.add('hidden');
}

/**
 * Read file — text for DXF/SVG, ArrayBuffer for PDF
 */
function readFileAsTextOrArrayBuffer(file, ext) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        if (ext === 'pdf') {
            reader.onload = (e) => resolve(e.target.result);
            reader.readAsArrayBuffer(file);
        } else {
            reader.onload = (e) => resolve(e.target.result);
            reader.readAsText(file);
        }
        reader.onerror = (e) => reject(e);
    });
}

/**
 * Format file size
 */
function formatFileSize(bytes) {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}

/**
 * Show/hide loading
 */
function showLoading(show, text) {
    if (show) {
        loading.classList.remove('hidden');
        calculateBtn.disabled = true;
        if (text) loadingText.textContent = text;
    } else {
        loading.classList.add('hidden');
        calculateBtn.disabled = false;
    }
}

/**
 * Show error
 */
function showError(message) {
    errorMessage.textContent = message;
    errorDiv.classList.remove('hidden');
}

function hideError() {
    errorDiv.classList.add('hidden');
}

/**
 * Download blob
 */
function downloadBlob(blob, filename) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
}

// Initialize
if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
} else {
    init();
}
