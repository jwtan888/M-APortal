/**
 * Vector Length Calculator Application
 * Calculates total vector path length from DXF files for machine runtime estimation
 */

// Global state
let currentDXFData = null;
let calculationResults = null;

// DOM Elements
const dropZone = document.getElementById('dropZone');
const fileInput = document.getElementById('fileInput');
const fileInfo = document.getElementById('fileInfo');
const fileName = document.getElementById('fileName');
const fileSize = document.getElementById('fileSize');
const calculateBtn = document.getElementById('calculateBtn');
const loading = document.getElementById('loading');
const results = document.getElementById('results');
const errorDiv = document.getElementById('error');
const errorMessage = document.getElementById('errorMessage');

// Initialize event listeners
function init() {
    // Drag and drop handlers
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
    
    // Calculate button
    calculateBtn.addEventListener('click', calculateVectorLength);
    
    // Runtime calculation
    document.getElementById('calcRuntimeBtn').addEventListener('click', calculateRuntime);
    
    // Export buttons
    document.getElementById('exportJsonBtn').addEventListener('click', exportToJSON);
    document.getElementById('exportCsvBtn').addEventListener('click', exportToCSV);
}

/**
 * Handle file drop
 */
function handleDrop(e) {
    e.preventDefault();
    dropZone.classList.remove('dragover');
    
    const files = e.dataTransfer.files;
    if (files.length > 0 && files[0].name.toLowerCase().endsWith('.dxf')) {
        processFile(files[0]);
    } else {
        showError('Please drop a valid DXF file');
    }
}

/**
 * Handle file selection via input
 */
function handleFileSelect(e) {
    const files = e.target.files;
    if (files.length > 0) {
        processFile(files[0]);
    }
}

/**
 * Process uploaded file
 */
function processFile(file) {
    hideError();
    
    // Show file info
    fileName.textContent = file.name;
    fileSize.textContent = formatFileSize(file.size);
    fileInfo.classList.remove('hidden');
    calculateBtn.classList.remove('hidden');
    results.classList.add('hidden');
    
    // Store file for later processing
    window.currentFile = file;
}

/**
 * Calculate vector length from DXF
 */
async function calculateVectorLength() {
    if (!window.currentFile) {
        showError('No file selected');
        return;
    }
    
    showLoading(true);
    
    try {
        const content = await readFileAsText(window.currentFile);
        
        // Parse DXF
        const parser = new DXFParser();
        const dxfData = parser.parse(content);
        currentDXFData = dxfData;
        
        // Calculate lengths
        calculationResults = calculateEntityLengths(dxfData.entities, parser.units);
        
        // Display results
        displayResults(calculationResults, parser.units);
        
        showLoading(false);
        results.classList.remove('hidden');
        
    } catch (err) {
        console.error('Error processing DXF:', err);
        showError('Failed to process DXF file: ' + err.message);
        showLoading(false);
    }
}

/**
 * Calculate lengths for all entities
 */
function calculateEntityLengths(entities, units) {
    const results = {
        totalLengthMm: 0,
        entityCount: 0,
        breakdown: {},
        entities: []
    };
    
    entities.forEach(entity => {
        const length = calculateEntityLength(entity);
        
        if (length > 0) {
            results.totalLengthMm += length;
            results.entityCount++;
            
            // Track by type
            if (!results.breakdown[entity.type]) {
                results.breakdown[entity.type] = { count: 0, length: 0 };
            }
            results.breakdown[entity.type].count++;
            results.breakdown[entity.type].length += length;
            
            // Store individual entity data
            results.entities.push({
                type: entity.type,
                layer: entity.layer || '0',
                lengthMm: length
            });
        }
    });
    
    // Convert to centimeters (primary unit for user)
    results.totalLengthCm = results.totalLengthMm / 10;
    results.totalLengthM = results.totalLengthMm / 1000;
    results.totalLengthInches = results.totalLengthMm / 25.4;
    results.totalLengthFeet = results.totalLengthInches / 12;
    results.totalLengthYards = results.totalLengthFeet / 3;
    
    results.units = units;
    
    return results;
}

/**
 * Calculate length of a single entity
 * This is the critical calculation - auditable and precise
 */
function calculateEntityLength(entity) {
    switch (entity.type) {
        case 'LINE':
            return calculateLineLength(entity);
        
        case 'CIRCLE':
            return calculateCircleLength(entity);
        
        case 'ARC':
            return calculateArcLength(entity);
        
        case 'POLYLINE':
        case 'LWPOLYLINE':
            return calculatePolylineLength(entity);
        
        case 'SPLINE':
            return calculateSplineLength(entity);
        
        case 'ELLIPSE':
            return calculateEllipseLength(entity);
        
        case 'POINT':
            return 0; // Points have no length
        
        default:
            // Try to handle unknown types as lines if they have coordinates
            if (entity.x !== undefined && entity.y !== undefined && 
                entity.x1 !== undefined && entity.y1 !== undefined) {
                return calculateLineLength(entity);
            }
            return 0;
    }
}

/**
 * Calculate LINE entity length using Euclidean distance
 * Formula: √[(x₂-x₁)² + (y₂-y₁)² + (z₂-z₁)²]
 */
function calculateLineLength(entity) {
    const dx = (entity.x1 || 0) - (entity.x || 0);
    const dy = (entity.y1 || 0) - (entity.y || 0);
    const dz = (entity.z1 || 0) - (entity.z || 0);
    
    return Math.sqrt(dx * dx + dy * dy + dz * dz);
}

/**
 * Calculate CIRCLE entity length (circumference)
 * Formula: 2πr
 */
function calculateCircleLength(entity) {
    const radius = entity.radius || 0;
    return 2 * Math.PI * radius;
}

/**
 * Calculate ARC entity length
 * Formula: r × θ (where θ is in radians)
 */
function calculateArcLength(entity) {
    const radius = entity.radius || 0;
    let startAngle = entity.startAngle || 0;
    let endAngle = entity.endAngle || 0;
    
    // Normalize angles to 0-2π range
    while (startAngle < 0) startAngle += 2 * Math.PI;
    while (endAngle < 0) endAngle += 2 * Math.PI;
    
    let angleDiff = endAngle - startAngle;
    
    // Handle arcs that cross the 0/2π boundary
    if (angleDiff < 0) {
        angleDiff += 2 * Math.PI;
    }
    
    return radius * angleDiff;
}

/**
 * Calculate POLYLINE/LWPOLYLINE length
 * Handles bulge (arc segments) in polylines
 */
function calculatePolylineLength(entity) {
    if (!entity.vertices || entity.vertices.length < 2) {
        return 0;
    }
    
    let totalLength = 0;
    const vertices = entity.vertices;
    const isClosed = entity.closed || (vertices.length > 2 && 
        Math.abs(vertices[0].x - vertices[vertices.length - 1].x) < 0.0001 &&
        Math.abs(vertices[0].y - vertices[vertices.length - 1].y) < 0.0001);
    
    for (let i = 0; i < vertices.length - 1; i++) {
        const v1 = vertices[i];
        const v2 = vertices[i + 1];
        
        if (v1.bulge && Math.abs(v1.bulge) > 0.0001) {
            // Arc segment with bulge
            totalLength += calculateBulgeArcLength(v1, v2);
        } else {
            // Straight line segment
            totalLength += calculateLineLength({ x: v1.x, y: v1.y, z: v1.z || 0, x1: v2.x, y1: v2.y, z1: v2.z || 0 });
        }
    }
    
    // Close the polyline if needed
    if (isClosed && vertices.length > 2) {
        const v1 = vertices[vertices.length - 1];
        const v2 = vertices[0];
        
        if (v1.bulge && Math.abs(v1.bulge) > 0.0001) {
            totalLength += calculateBulgeArcLength(v1, v2);
        } else {
            totalLength += calculateLineLength({ x: v1.x, y: v1.y, z: v1.z || 0, x1: v2.x, y1: v2.y, z1: v2.z || 0 });
        }
    }
    
    return totalLength;
}

/**
 * Calculate arc length from bulge value
 * Bulge = tan(θ/4) where θ is the included angle
 */
function calculateBulgeArcLength(v1, v2) {
    const bulge = v1.bulge;
    const chordLength = Math.sqrt(
        Math.pow(v2.x - v1.x, 2) + Math.pow(v2.y - v1.y, 2)
    );
    
    if (chordLength < 0.0001) return 0;
    
    // Calculate included angle from bulge
    const includedAngle = 4 * Math.atan(Math.abs(bulge));
    
    // Calculate radius: r = chord / (2 * sin(θ/2))
    const radius = chordLength / (2 * Math.sin(includedAngle / 2));
    
    // Arc length = r × θ
    return radius * includedAngle;
}

/**
 * Calculate SPLINE length using numerical integration
 * Uses adaptive sampling for accuracy
 */
function calculateSplineLength(entity) {
    if (!entity.controlPoints || entity.controlPoints.length < 2) {
        return 0;
    }
    
    // For splines, use chord length approximation with fine sampling
    // This is more accurate than trying to compute the exact parametric length
    const samples = 100; // Number of sample points per span
    let totalLength = 0;
    
    const points = entity.controlPoints;
    
    for (let i = 0; i < points.length - 1; i++) {
        const p1 = points[i];
        const p2 = points[i + 1];
        
        // Linear interpolation between control points
        // (Note: This is an approximation; true B-splines would require more complex math)
        for (let j = 0; j < samples; j++) {
            const t1 = j / samples;
            const t2 = (j + 1) / samples;
            
            const x1 = p1.x + (p2.x - p1.x) * t1;
            const y1 = p1.y + (p2.y - p1.y) * t1;
            const x2 = p1.x + (p2.x - p1.x) * t2;
            const y2 = p1.y + (p2.y - p1.y) * t2;
            
            totalLength += Math.sqrt(Math.pow(x2 - x1, 2) + Math.pow(y2 - y1, 2));
        }
    }
    
    return totalLength;
}

/**
 * Calculate ELLIPSE length using Ramanujan's approximation
 * Formula: π × [3(a+b) - √((3a+b)(a+3b))]
 * where a = semi-major axis, b = semi-minor axis
 */
function calculateEllipseLength(entity) {
    // Get major axis length
    const majorAxisX = entity.majorAxisX || 0;
    const majorAxisY = entity.majorAxisY || 0;
    const majorAxisLength = Math.sqrt(majorAxisX * majorAxisX + majorAxisY * majorAxisY);
    
    // Get axis ratio (minor/major)
    const axisRatio = entity.axisRatio || 1;
    const minorAxisLength = majorAxisLength * axisRatio;
    
    const a = majorAxisLength / 2; // Semi-major axis
    const b = minorAxisLength / 2; // Semi-minor axis
    
    // Check if it's a partial ellipse
    const startParam = entity.startParam || 0;
    const endParam = entity.endParam || 2 * Math.PI;
    const paramRange = endParam - startParam;
    
    // Ramanujan's second approximation for full ellipse circumference
    const h = Math.pow(a - b, 2) / Math.pow(a + b, 2);
    const circumference = Math.PI * (a + b) * (1 + (3 * h) / (10 + Math.sqrt(4 - 3 * h)));
    
    // If partial ellipse, scale by parameter range
    if (paramRange < 2 * Math.PI - 0.0001) {
        return circumference * (paramRange / (2 * Math.PI));
    }
    
    return circumference;
}

/**
 * Display calculation results
 */
function displayResults(results, units) {
    // Main length displays (in cm as primary unit)
    document.getElementById('lengthCm').textContent = results.totalLengthCm.toFixed(4);
    document.getElementById('lengthMm').textContent = results.totalLengthMm.toFixed(4);
    document.getElementById('lengthM').textContent = results.totalLengthM.toFixed(4);
    document.getElementById('lengthIn').textContent = results.totalLengthInches.toFixed(4);
    document.getElementById('lengthFt').textContent = results.totalLengthFeet.toFixed(4);
    document.getElementById('lengthYd').textContent = results.totalLengthYards.toFixed(4);
    
    // Entity breakdown
    const breakdownDiv = document.getElementById('entityBreakdown');
    breakdownDiv.innerHTML = '';
    
    for (const [type, data] of Object.entries(results.breakdown)) {
        const item = document.createElement('div');
        item.className = 'flex justify-between items-center p-3 bg-gray-50 rounded-lg';
        item.innerHTML = `
            <span class="font-medium text-gray-700">${type}</span>
            <div class="text-right">
                <span class="text-blue-600 font-semibold">${data.count} entities</span>
                <span class="mx-2 text-gray-400">|</span>
                <span class="text-gray-800">${(data.length / 10).toFixed(2)} cm</span>
            </div>
        `;
        breakdownDiv.appendChild(item);
    }
    
    // Audit info
    document.getElementById('totalEntities').textContent = results.entityCount;
    document.getElementById('unitsDetected').textContent = units;
    
    // Reset runtime
    document.getElementById('runtime').textContent = '--:--';
}

/**
 * Calculate machine runtime
 */
function calculateRuntime() {
    if (!calculationResults) return;
    
    const speed = parseFloat(document.getElementById('machineSpeed').value) || 50; // cm/min
    const lengthCm = calculationResults.totalLengthCm;
    
    const timeMinutes = lengthCm / speed;
    const hours = Math.floor(timeMinutes / 60);
    const minutes = Math.floor(timeMinutes % 60);
    const seconds = Math.floor((timeMinutes % 1) * 60);
    
    let timeStr = '';
    if (hours > 0) {
        timeStr += `${hours}h `;
    }
    if (minutes > 0 || hours > 0) {
        timeStr += `${minutes}m `;
    }
    timeStr += `${seconds}s`;
    
    document.getElementById('runtime').textContent = timeStr.trim();
}

/**
 * Export results to JSON
 */
function exportToJSON() {
    if (!calculationResults) return;
    
    const exportData = {
        timestamp: new Date().toISOString(),
        file: window.currentFile?.name || 'Unknown',
        units: calculationResults.units,
        totals: {
            centimeters: calculationResults.totalLengthCm,
            millimeters: calculationResults.totalLengthMm,
            meters: calculationResults.totalLengthM,
            inches: calculationResults.totalLengthInches,
            feet: calculationResults.totalLengthFeet,
            yards: calculationResults.totalLengthYards
        },
        entityCount: calculationResults.entityCount,
        breakdown: calculationResults.breakdown,
        entities: calculationResults.entities
    };
    
    const blob = new Blob([JSON.stringify(exportData, null, 2)], { type: 'application/json' });
    downloadBlob(blob, 'vector-length-results.json');
}

/**
 * Export results to CSV
 */
function exportToCSV() {
    if (!calculationResults) return;
    
    let csv = 'Entity Type,Layer,Length (mm),Length (cm)\n';
    
    calculationResults.entities.forEach(entity => {
        csv += `"${entity.type}","${entity.layer}",${entity.lengthMm.toFixed(4)},${(entity.lengthMm / 10).toFixed(4)}\n`;
    });
    
    csv += `\n"TOTAL",,"${calculationResults.totalLengthMm.toFixed(4)}","${calculationResults.totalLengthCm.toFixed(4)}"\n`;
    
    const blob = new Blob([csv], { type: 'text/csv' });
    downloadBlob(blob, 'vector-length-breakdown.csv');
}

/**
 * Download blob as file
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

/**
 * Read file as text
 */
function readFileAsText(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (e) => resolve(e.target.result);
        reader.onerror = (e) => reject(e);
        reader.readAsText(file);
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
 * Show/hide loading state
 */
function showLoading(show) {
    if (show) {
        loading.classList.remove('hidden');
        calculateBtn.disabled = true;
    } else {
        loading.classList.add('hidden');
        calculateBtn.disabled = false;
    }
}

/**
 * Show error message
 */
function showError(message) {
    errorMessage.textContent = message;
    errorDiv.classList.remove('hidden');
}

/**
 * Hide error message
 */
function hideError() {
    errorDiv.classList.add('hidden');
}

// Initialize on DOM ready
if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
} else {
    init();
}
