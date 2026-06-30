/**
 * SVG Vector Parser — extracts vector paths from SVG files
 * Supports: line, rect, circle, ellipse, polygon, polyline, path
 * Ignores: text, tspan, image, and other non-vector elements
 */

class SVGParser {
    constructor() {
        this.entities = [];
        this.units = 'px'; // Default SVG unit
    }

    /**
     * Parse SVG content string
     * @param {string} content
     * @returns {Array} Array of vector entities
     */
    parse(content) {
        this.entities = [];

        const parser = new DOMParser();
        const doc = parser.parseFromString(content, 'image/svg+xml');

        // Check for parse errors
        const parseError = doc.querySelector('parsererror');
        if (parseError) {
            throw new Error('Invalid SVG: ' + parseError.textContent.substring(0, 200));
        }

        const svg = doc.documentElement;
        if (svg.tagName.toLowerCase() !== 'svg') {
            throw new Error('Not a valid SVG file');
        }

        // Detect units from SVG attributes
        this._detectUnits(svg);

        // Process all vector elements
        this._processElement(svg);

        return this.entities;
    }

    /**
     * Detect units from SVG root element
     */
    _detectUnits(svg) {
        const width = svg.getAttribute('width');
        const height = svg.getAttribute('height');
        const viewBox = svg.getAttribute('viewBox');

        // Check width/height for unit suffix
        const unitMatch = (width || '').match(/(px|pt|mm|cm|in|em|rem)$/);
        if (unitMatch) {
            this.units = unitMatch[1];
        }

        // Check viewBox for scale hint
        if (viewBox && !unitMatch) {
            // SVGs with viewBox typically use user units (px equivalent)
            this.units = 'px';
        }
    }

    /**
     * Process an SVG element and its children
     */
    _processElement(element) {
        const tag = element.tagName.toLowerCase();

        // Skip text and non-vector elements
        const skipTags = ['text', 'tspan', 'textpath', 'image', 'use', 'defs', 'style', 'script', 'metadata', 'title', 'desc', 'foreignobject'];
        if (skipTags.includes(tag)) {
            return;
        }

        // Extract vector entity based on tag
        const entity = this._extractEntity(element);
        if (entity) {
            this.entities.push(entity);
        }

        // Process children (but not <defs> contents — they're referenced, not rendered directly)
        if (tag !== 'defs') {
            Array.from(element.children).forEach(child => this._processElement(child));
        }
    }

    /**
     * Extract a vector entity from an SVG element
     */
    _extractEntity(element) {
        const tag = element.tagName.toLowerCase();
        const layer = this._getLayer(element);

        switch (tag) {
            case 'line':
                return {
                    type: 'LINE',
                    layer,
                    x: this._parseNum(element.getAttribute('x1')),
                    y: this._parseNum(element.getAttribute('y1')),
                    x1: this._parseNum(element.getAttribute('x2')),
                    y1: this._parseNum(element.getAttribute('y2'))
                };

            case 'rect': {
                const x = this._parseNum(element.getAttribute('x')) || 0;
                const y = this._parseNum(element.getAttribute('y')) || 0;
                const w = this._parseNum(element.getAttribute('width')) || 0;
                const h = this._parseNum(element.getAttribute('height')) || 0;
                const rx = this._parseNum(element.getAttribute('rx')) || 0;
                const ry = this._parseNum(element.getAttribute('ry')) || 0;

                if (rx > 0 || ry > 0) {
                    // Rounded rect — approximate as path
                    return {
                        type: 'PATH',
                        layer,
                        pathSegments: this._roundedRectPath(x, y, w, h, rx || ry, ry || rx)
                    };
                }
                return {
                    type: 'RECT',
                    layer,
                    x, y, width: w, height: h
                };
            }

            case 'circle':
                return {
                    type: 'CIRCLE',
                    layer,
                    cx: this._parseNum(element.getAttribute('cx')),
                    cy: this._parseNum(element.getAttribute('cy')),
                    radius: this._parseNum(element.getAttribute('r'))
                };

            case 'ellipse':
                // Approximate ellipse as ellipse entity
                const cx = this._parseNum(element.getAttribute('cx')) || 0;
                const cy = this._parseNum(element.getAttribute('cy')) || 0;
                const rx = this._parseNum(element.getAttribute('rx')) || 0;
                const ry = this._parseNum(element.getAttribute('ry')) || 0;
                return {
                    type: 'ELLIPSE',
                    layer,
                    cx, cy,
                    majorAxisX: rx * 2,
                    majorAxisY: 0,
                    axisRatio: ry / rx || 1
                };

            case 'polygon':
            case 'polyline': {
                const points = element.getAttribute('points');
                if (!points) return null;
                const coords = points.trim().split(/[\s,]+/).map(Number);
                const vertices = [];
                for (let i = 0; i < coords.length; i += 2) {
                    vertices.push({ x: coords[i], y: coords[i + 1] });
                }
                if (vertices.length < 2) return null;
                return {
                    type: 'POLYLINE',
                    layer,
                    vertices,
                    closed: tag === 'polygon'
                };
            }

            case 'path': {
                const d = element.getAttribute('d');
                if (!d) return null;
                const segments = this._parsePathD(d);
                if (segments.length === 0) return null;
                return {
                    type: 'PATH',
                    layer,
                    pathSegments: segments
                };
            }

            default:
                return null;
        }
    }

    /**
     * Parse SVG path 'd' attribute into segments
     */
    _parsePathD(d) {
        const segments = [];
        // Tokenize path data
        const tokens = d.match(/[a-zA-Z][^a-zA-Z]*/g);
        if (!tokens) return segments;

        let cx = 0, cy = 0; // current point
        let startX = 0, startY = 0; // start of subpath

        for (const token of tokens) {
            const cmd = token[0];
            const args = token.substring(1).trim().split(/[\s,]+/).filter(s => s !== '').map(Number);

            switch (cmd.toUpperCase()) {
                case 'M':
                    for (let i = 0; i < args.length; i += 2) {
                        const x = cmd === 'm' && segments.length > 0 ? cx + args[i] : args[i];
                        const y = cmd === 'm' && segments.length > 0 ? cy + args[i + 1] : args[i + 1];
                        segments.push({ command: 'M', x, y });
                        cx = x; cy = y;
                        startX = x; startY = y;
                    }
                    break;

                case 'L':
                    for (let i = 0; i < args.length; i += 2) {
                        const x = cmd === 'l' ? cx + args[i] : args[i];
                        const y = cmd === 'l' ? cy + args[i + 1] : args[i + 1];
                        segments.push({ command: 'L', x, y });
                        cx = x; cy = y;
                    }
                    break;

                case 'H': {
                    const x = cmd === 'h' ? cx + args[0] : args[0];
                    segments.push({ command: 'L', x, y: cy });
                    cx = x;
                    break;
                }

                case 'V': {
                    const y = cmd === 'v' ? cy + args[0] : args[0];
                    segments.push({ command: 'L', x: cx, y });
                    cy = y;
                    break;
                }

                case 'C':
                    for (let i = 0; i < args.length; i += 6) {
                        const x1 = cmd === 'c' ? cx + args[i] : args[i];
                        const y1 = cmd === 'c' ? cy + args[i + 1] : args[i + 1];
                        const x2 = cmd === 'c' ? cx + args[i + 2] : args[i + 2];
                        const y2 = cmd === 'c' ? cy + args[i + 3] : args[i + 3];
                        const x = cmd === 'c' ? cx + args[i + 4] : args[i + 4];
                        const y = cmd === 'c' ? cy + args[i + 5] : args[i + 5];
                        segments.push({ command: 'C', x1, y1, x2, y2, x, y });
                        cx = x; cy = y;
                    }
                    break;

                case 'S':
                    for (let i = 0; i < args.length; i += 4) {
                        const x = cmd === 's' ? cx + args[i + 2] : args[i + 2];
                        const y = cmd === 's' ? cy + args[i + 3] : args[i + 3];
                        segments.push({ command: 'C', x1: cx, y1: cy, x2: cmd === 's' ? cx + args[i] : args[i], y2: cmd === 's' ? cy + args[i + 1] : args[i + 1], x, y });
                        cx = x; cy = y;
                    }
                    break;

                case 'Q':
                    for (let i = 0; i < args.length; i += 4) {
                        const x1 = cmd === 'q' ? cx + args[i] : args[i];
                        const y1 = cmd === 'q' ? cy + args[i + 1] : args[i + 1];
                        const x = cmd === 'q' ? cx + args[i + 2] : args[i + 2];
                        const y = cmd === 'q' ? cy + args[i + 3] : args[i + 3];
                        segments.push({ command: 'Q', x1, y1, x, y });
                        cx = x; cy = y;
                    }
                    break;

                case 'A':
                    for (let i = 0; i < args.length; i += 7) {
                        const x = cmd === 'a' ? cx + args[i + 5] : args[i + 5];
                        const y = cmd === 'a' ? cy + args[i + 6] : args[i + 6];
                        segments.push({
                            command: 'A',
                            rx: args[i], ry: args[i + 1],
                            rotation: args[i + 2],
                            largeArc: args[i + 3],
                            sweep: args[i + 4],
                            x, y
                        });
                        cx = x; cy = y;
                    }
                    break;

                case 'Z':
                    segments.push({ command: 'Z' });
                    cx = startX; cy = startY;
                    break;
            }
        }

        return segments;
    }

    /**
     * Generate path segments for a rounded rectangle
     */
    _roundedRectPath(x, y, w, h, rx, ry) {
        return [
            { command: 'M', x: x + rx, y },
            { command: 'L', x: x + w - rx, y },
            { command: 'A', rx, ry, xAxisRotation: 0, largeArcFlag: 0, sweepFlag: 1, x: x + w, y: y + ry },
            { command: 'L', x: x + w, y: y + h - ry },
            { command: 'A', rx, ry, xAxisRotation: 0, largeArcFlag: 0, sweepFlag: 1, x: x + w - rx, y: y + h },
            { command: 'L', x: x + rx, y: y + h },
            { command: 'A', rx, ry, xAxisRotation: 0, largeArcFlag: 0, sweepFlag: 1, x, y: y + h - ry },
            { command: 'L', x, y: y + ry },
            { command: 'A', rx, ry, xAxisRotation: 0, largeArcFlag: 0, sweepFlag: 1, x: x + rx, y },
            { command: 'Z' }
        ];
    }

    /**
     * Get layer/group name for an element
     */
    _getLayer(element) {
        // Check for group id or class
        const parent = this._findGroupParent(element);
        if (parent) {
            return parent.getAttribute('id') || parent.getAttribute('class') || '-';
        }
        return element.getAttribute('id') || element.getAttribute('class') || '-';
    }

    /**
     * Find nearest group parent
     */
    _findGroupParent(element) {
        let parent = element.parentElement;
        while (parent) {
            if (parent.tagName.toLowerCase() === 'g') {
                return parent;
            }
            parent = parent.parentElement;
        }
        return null;
    }

    /**
     * Parse a numeric value, handling unit suffixes
     */
    _parseNum(val) {
        if (!val) return 0;
        const num = parseFloat(val);
        return isNaN(num) ? 0 : num;
    }
}

// Export for browser
if (typeof module !== 'undefined' && module.exports) {
    module.exports = SVGParser;
}
