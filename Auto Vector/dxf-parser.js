/**
 * DXF Parser - Pure JavaScript DXF file parser
 * Parses DXF files and extracts entities for vector length calculation
 */

class DXFParser {
    constructor() {
        this.data = null;
        this.entities = [];
        this.units = 'mm'; // Default to millimeters
        this.layerColors = {}; // layer name → {color: aci, trueColor: rgb24}
    }

    /**
     * Parse DXF file content (as text)
     * @param {string} content - DXF file content as string
     * @returns {Object} Parsed DXF data
     */
    parse(content) {
        const lines = content.split(/\r\n|\n|\r/);
        this.data = { header: {}, entities: [], blocks: [] };
        
        let i = 0;
        let currentSection = null;
        let currentEntity = null;
        let inEntities = false;
        let inBlocks = false;
        
        while (i < lines.length) {
            const groupCode = parseInt(lines[i].trim(), 10);
            const value = lines[i + 1];
            
            if (isNaN(groupCode)) {
                i++;
                continue;
            }
            
            // Section handling
            if (groupCode === 0) {
                if (value === 'SECTION') {
                    currentSection = this.parseSection(lines, i);
                    if (currentSection.name === 'HEADER') {
                        this.data.header = currentSection.data;
                        // Detect units from header
                        if (currentSection.data.$INSUNITS) {
                            this.units = this.getUnits(currentSection.data.$INSUNITS);
                        }
                    } else if (currentSection.name === 'ENTITIES') {
                        this.data.entities = currentSection.entities || [];
                    } else if (currentSection.name === 'BLOCKS') {
                        this.data.blocks = currentSection.blocks || [];
                    } else if (currentSection.name === 'TABLES') {
                        // Extract layer colors from LAYER table
                        this.layerColors = currentSection.layerColors || {};
                    }
                    i = currentSection.endIndex;
                    continue;
                }
            }
            
            i += 2;
        }
        
        // Expand INSERT (block reference) entities into actual geometry
        this.expandInserts();
        this.data.units = this.units;

        return this.data;
    }

    /**
     * Expand INSERT entities — replace block references with actual block geometry
     */
    expandInserts() {
        const blocks = this.data.blocks || [];
        const blockMap = {};
        blocks.forEach(b => {
            if (b.name && b.entities && b.entities.length > 0) {
                blockMap[b.name] = b.entities;
            }
        });

        let i = 0;
        while (i < this.data.entities.length) {
            const entity = this.data.entities[i];
            if (entity.type === 'INSERT' && entity.blockName) {
                const blockEntities = blockMap[entity.blockName];
                if (blockEntities) {
                    const sx = entity.scaleX || 1;
                    const sy = entity.scaleY || 1;
                    const ox = entity.insertX || 0;
                    const oy = entity.insertY || 0;
                    const rot = (entity.rotation || 0) * Math.PI / 180;
                    const cosR = Math.cos(rot), sinR = Math.sin(rot);

                    const expanded = blockEntities
                        .filter(e => e.type !== 'POINT')
                        .map(be => this.transformEntity(be, sx, sy, ox, oy, cosR, sinR));

                    expanded.forEach(e => {
                        if (!e.layer || e.layer === '0') e.layer = entity.layer || '0';
                        if (e.color === undefined || e.color === 0 || e.color === 256) e.color = entity.color;
                        if (!e.trueColor && entity.trueColor) e.trueColor = entity.trueColor;
                    });

                    this.data.entities.splice(i, 1, ...expanded);
                    i += expanded.length;
                } else {
                    i++;
                }
            } else {
                i++;
            }
        }
    }

    /**
     * Transform an entity's coordinates (scale, rotate, translate)
     */
    transformEntity(e, sx, sy, ox, oy, cosR, sinR) {
        function tx(x, y) {
            const rx = x * sx, ry = y * sy;
            return { x: rx * cosR - ry * sinR + ox, y: rx * sinR + ry * cosR + oy };
        }

        const out = { type: e.type, layer: e.layer, color: e.color, trueColor: e.trueColor };

        if (e.type === 'LINE') {
            const p1 = tx(e.x || 0, e.y || 0), p2 = tx(e.x1 || 0, e.y1 || 0);
            out.x = p1.x; out.y = p1.y; out.x1 = p2.x; out.y1 = p2.y;
        } else if (e.type === 'CIRCLE') {
            const c = tx(e.cx || 0, e.cy || 0);
            out.cx = c.x; out.cy = c.y; out.radius = (e.radius || 0) * Math.max(sx, sy);
        } else if (e.type === 'ARC') {
            const c = tx(e.cx || 0, e.cy || 0);
            out.cx = c.x; out.cy = c.y; out.radius = (e.radius || 0) * Math.max(sx, sy);
            out.startAngle = e.startAngle;
            out.endAngle = e.endAngle;
        } else if (e.type === 'POLYLINE' || e.type === 'LWPOLYLINE') {
            out.vertices = (e.vertices || []).map(v => {
                const p = tx(v.x, v.y);
                return { x: p.x, y: p.y, bulge: v.bulge };
            });
            out.closed = e.closed;
        } else if (e.type === 'ELLIPSE') {
            const c = tx(e.cx || 0, e.cy || 0);
            out.cx = c.x; out.cy = c.y;
            out.majorAxisX = (e.majorAxisX || 0) * sx;
            out.majorAxisY = (e.majorAxisY || 0) * sy;
            out.axisRatio = e.axisRatio || 1;
            out.startParam = e.startParam;
            out.endParam = e.endParam;
        } else if (e.type === 'SPLINE') {
            out.controlPoints = (e.controlPoints || []).map(p => {
                const t = tx(p.x, p.y);
                return { x: t.x, y: t.y };
            });
        } else if (e.type === 'HATCH') {
            out.boundaryEdges = (e.boundaryEdges || []).map(edge => {
                const ne = { edgeType: edge.edgeType };
                if (edge.edgeType === 1) {
                    const p1 = tx(edge.x, edge.y), p2 = tx(edge.x1, edge.y1);
                    ne.x = p1.x; ne.y = p1.y; ne.x1 = p2.x; ne.y1 = p2.y;
                } else if (edge.edgeType === 2) {
                    const c = tx(edge.x, edge.y);
                    ne.x = c.x; ne.y = c.y; ne.radius = (edge.radius || 0) * Math.max(sx, sy);
                    ne.startAngle = edge.startAngle;
                    ne.endAngle = edge.endAngle;
                }
                return ne;
            });
        } else {
            Object.assign(out, e);
        }

        return out;
    }

    /**
     * Parse a DXF section
     */
    parseSection(lines, startIndex) {
        let i = startIndex + 2; // Skip SECTION line
        let sectionName = null;
        let data = {};
        let entities = [];
        let blocks = [];
        
        // Get section name
        while (i < lines.length) {
            const groupCode = parseInt(lines[i].trim(), 10);
            const value = lines[i + 1];
            
            if (groupCode === 2 && !sectionName) {
                sectionName = value.trim();
                i += 2;
                break;
            }
            i += 2;
        }
        
        // Parse section content
        while (i < lines.length) {
            const groupCode = parseInt(lines[i].trim(), 10);
            const value = lines[i + 1];
            
            if (groupCode === 0 && value === 'ENDSEC') {
                return {
                    name: sectionName,
                    data: data,
                    entities: entities,
                    blocks: blocks,
                    endIndex: i + 2
                };
            }
            
            if (sectionName === 'HEADER' && groupCode !== 0) {
                const varName = value ? value.trim() : '';
                if (varName.startsWith('$')) {
                    i += 2;
                    const nextCode = parseInt(lines[i]?.trim(), 10);
                    const nextValue = lines[i + 1];
                    if (!isNaN(nextCode)) {
                        data[varName] = this.convertValue(nextCode, nextValue);
                        i += 2;
                    }
                    continue;
                }
            }
            
            if (sectionName === 'ENTITIES' && groupCode === 0) {
                const entity = this.parseEntity(lines, i);
                if (entity) {
                    entities.push(entity);
                    i = entity.endIndex;
                    continue;
                }
            }
            
            if (sectionName === 'BLOCKS' && groupCode === 0) {
                const block = this.parseBlock(lines, i);
                if (block) {
                    blocks.push(block);
                    i = block.endIndex;
                    continue;
                }
            }

            // Parse TABLES section — extract LAYER colors
            if (sectionName === 'TABLES' && groupCode === 0) {
                const item = lines[i + 1]?.trim();
                if (item === 'LAYER') {
                    const layerInfo = this.parseLayer(lines, i);
                    if (layerInfo) {
                        data._layerColors = data._layerColors || {};
                        data._layerColors[layerInfo.name] = { color: layerInfo.color, trueColor: layerInfo.trueColor };
                        i = layerInfo.endIndex;
                        continue;
                    }
                }
                // Skip other table entries (STYLE, LTYPE, etc.)
                if (item === 'TABLE' || item === 'ENDTAB') {
                    i += 2;
                    continue;
                }
            }
            
            i += 2;
        }
        
        return {
            name: sectionName,
            data: data,
            entities: entities,
            blocks: blocks,
            layerColors: data._layerColors || {},
            endIndex: i
        };
    }
    
    /**
     * Parse a LAYER table entry
     */
    parseLayer(lines, startIndex) {
        // startIndex points to group code 0, value "LAYER"
        let i = startIndex + 2; // Skip the "0 LAYER" line
        const layer = { name: '', color: 7, trueColor: null };

        while (i < lines.length) {
            const groupCode = parseInt(lines[i]?.trim(), 10);
            const value = lines[i + 1];

            if (isNaN(groupCode)) { i++; continue; }

            if (groupCode === 0) {
                // End of this LAYER entry
                layer.endIndex = i;
                return layer;
            }

            const parsed = this.convertValue(groupCode, value);
            if (groupCode === 2) layer.name = parsed;     // Layer name
            if (groupCode === 62) layer.color = parsed;    // ACI color
            if (groupCode === 420) layer.trueColor = parsed; // True Color RGB24

            i += 2;
        }

        layer.endIndex = i;
        return layer;
    }

    /**
     * Parse a DXF entity
     */
    parseEntity(lines, startIndex) {
        const entityType = lines[startIndex + 1]?.trim();
        if (!entityType || entityType === 'ENDSEC' || entityType === 'EOF') {
            return null;
        }
        
        const entity = { type: entityType };
        let i = startIndex + 2;
        
        while (i < lines.length) {
            const groupCode = parseInt(lines[i]?.trim(), 10);
            const value = lines[i + 1];
            
            if (groupCode === 0) {
                entity.endIndex = i;
                break;
            }
            
            this.parseEntityProperty(entity, groupCode, value);
            i += 2;
        }
        
        if (!entity.endIndex) {
            entity.endIndex = i;
        }
        
        return entity;
    }
    
    /**
     * Parse entity property based on group code
     */
    parseEntityProperty(entity, groupCode, value) {
        const parsedValue = this.convertValue(groupCode, value);
        
        // Common group codes for coordinates
        if (groupCode === 10) entity.x = parsedValue;
        else if (groupCode === 20) entity.y = parsedValue;
        else if (groupCode === 30) entity.z = parsedValue;
        else if (groupCode === 11) entity.x1 = parsedValue;
        else if (groupCode === 21) entity.y1 = parsedValue;
        else if (groupCode === 31) entity.z1 = parsedValue;
        else if (groupCode === 12) entity.x2 = parsedValue;
        else if (groupCode === 22) entity.y2 = parsedValue;
        else if (groupCode === 32) entity.z2 = parsedValue;
        else if (groupCode === 13) entity.x3 = parsedValue;
        else if (groupCode === 23) entity.y3 = parsedValue;
        else if (groupCode === 33) entity.z3 = parsedValue;
        else if (groupCode === 40) entity.radius = parsedValue;
        else if (groupCode === 41) entity.startAngle = parsedValue;
        else if (groupCode === 42) entity.endAngle = parsedValue;
        else if (groupCode === 50) entity.startAngle = parsedValue;
        else if (groupCode === 51) entity.endAngle = parsedValue;
        else if (groupCode === 8) entity.layer = parsedValue;
        else if (groupCode === 6) entity.linetype = parsedValue;
        else if (groupCode === 62) entity.color = parsedValue;
        else if (groupCode === 420) entity.trueColor = parsedValue;
        else if (groupCode === 430) entity.colorName = parsedValue;
        else if (groupCode === 370) entity.lineweight = parsedValue;
        else if (groupCode === 48) entity.dashScale = parsedValue;
        else if (groupCode === 73) entity.direction = parsedValue;
        
        // Handle INSERT (block reference) — store block name and insertion point
        if (entity.type === 'INSERT') {
            if (groupCode === 2) entity.blockName = parsedValue;
            if (groupCode === 10) entity.insertX = parsedValue;
            else if (groupCode === 20) entity.insertY = parsedValue;
            else if (groupCode === 30) entity.insertZ = parsedValue;
            if (groupCode === 41) entity.scaleX = parsedValue;
            else if (groupCode === 42) entity.scaleY = parsedValue;
            else if (groupCode === 43) entity.scaleZ = parsedValue;
            if (groupCode === 50) entity.rotation = parsedValue;
        }

        // Handle arc center point (same as circle)
        if (entity.type === 'ARC' || entity.type === 'CIRCLE') {
            if (groupCode === 10) entity.cx = parsedValue;
            else if (groupCode === 20) entity.cy = parsedValue;
            else if (groupCode === 30) entity.cz = parsedValue;
        }
        
        // Handle ellipse
        if (entity.type === 'ELLIPSE') {
            if (groupCode === 10) entity.cx = parsedValue;
            else if (groupCode === 20) entity.cy = parsedValue;
            else if (groupCode === 30) entity.cz = parsedValue;
            else if (groupCode === 11) entity.majorAxisX = parsedValue;
            else if (groupCode === 21) entity.majorAxisY = parsedValue;
            else if (groupCode === 22) entity.majorAxisZ = parsedValue;
            else if (groupCode === 40) entity.axisRatio = parsedValue;
            else if (groupCode === 41) entity.startParam = parsedValue;
            else if (groupCode === 42) entity.endParam = parsedValue;
        }
        
        // Handle spline control points
        if (entity.type === 'SPLINE') {
            if (!entity.controlPoints) entity.controlPoints = [];
            if (groupCode === 10) {
                if (!entity.currentPoint) entity.currentPoint = {};
                entity.currentPoint.x = parsedValue;
            } else if (groupCode === 20) {
                if (!entity.currentPoint) entity.currentPoint = {};
                entity.currentPoint.y = parsedValue;
                entity.controlPoints.push({...entity.currentPoint});
                entity.currentPoint = null;
            }
            
            if (!entity.knots) entity.knots = [];
            if (groupCode === 40) entity.knots.push(parsedValue);
            
            if (!entity.weights) entity.weights = [];
            if (groupCode === 41) entity.weights.push(parsedValue);
            
            if (groupCode === 71) entity.degree = parsedValue;
            if (groupCode === 72) entity.numKnots = parsedValue;
            if (groupCode === 73) entity.numControlPoints = parsedValue;
        }
        
        // Handle polyline vertices
        if (entity.type === 'POLYLINE' || entity.type === 'LWPOLYLINE') {
            if (!entity.vertices) entity.vertices = [];
            if (groupCode === 10) {
                if (!entity.currentVertex) entity.currentVertex = {};
                entity.currentVertex.x = parsedValue;
            } else if (groupCode === 20) {
                if (!entity.currentVertex) entity.currentVertex = {};
                entity.currentVertex.y = parsedValue;
                entity.vertices.push({...entity.currentVertex});
                entity.currentVertex = null;
            } else if (groupCode === 40) {
                if (entity.vertices.length > 0) {
                    entity.vertices[entity.vertices.length - 1].bulge = parsedValue;
                }
            }
        }

        // Handle HATCH boundary edges
        if (entity.type === 'HATCH') {
            if (!entity.boundaryEdges) entity.boundaryEdges = [];
            if (groupCode === 91) entity.numPaths = parsedValue;
            if (groupCode === 92) entity.currentPathType = parsedValue;
            if (groupCode === 93) entity.currentNumEdges = parsedValue;
            if (groupCode === 72) {
                // Edge type: 1=LINE, 2=ARC, 3=ELLIPSE, 4=SPLINE
                entity.currentEdge = { edgeType: parsedValue };
            }
            if (entity.currentEdge) {
                if (groupCode === 10) entity.currentEdge.x = parsedValue;
                if (groupCode === 20) entity.currentEdge.y = parsedValue;
                if (groupCode === 11) entity.currentEdge.x1 = parsedValue;
                if (groupCode === 21) entity.currentEdge.y1 = parsedValue;
                if (groupCode === 40) entity.currentEdge.radius = parsedValue;
                if (groupCode === 50) entity.currentEdge.startAngle = parsedValue;
                if (groupCode === 51) entity.currentEdge.endAngle = parsedValue;
                if (groupCode === 73) entity.currentEdge.ccw = parsedValue;
                // Ellipse edge specifics
                if (groupCode === 12) entity.currentEdge.majorAxisX = parsedValue;
                if (groupCode === 22) entity.currentEdge.majorAxisY = parsedValue;
                if (groupCode === 42) entity.currentEdge.axisRatio = parsedValue;
                // Spline edge
                if (groupCode === 94) entity.currentEdge.degree = parsedValue;
                if (groupCode === 95) entity.currentEdge.numFitPoints = parsedValue;
                if (groupCode === 96) entity.currentEdge.numKnots = parsedValue;

                // Push edge when we see the next edge type marker or another group 72
                if (groupCode === 72 && entity.currentEdge.edgeType !== parsedValue) {
                    entity.boundaryEdges.push({...entity.currentEdge});
                    entity.currentEdge = { edgeType: parsedValue };
                }
                // Also push on boundary count or end of edges
                if (groupCode === 93 || groupCode === 97) {
                    if (entity.currentEdge.edgeType) {
                        entity.boundaryEdges.push({...entity.currentEdge});
                        entity.currentEdge = null;
                    }
                }
            }
        }
    }
    
    /**
     * Parse a BLOCK section
     */
    parseBlock(lines, startIndex) {
        const block = { name: '', entities: [] };
        let i = startIndex + 2;
        
        while (i < lines.length) {
            const groupCode = parseInt(lines[i]?.trim(), 10);
            const value = lines[i + 1];
            
            if (groupCode === 0 && value === 'ENDBLK') {
                block.endIndex = i + 2;
                break;
            }

            if (groupCode === 2) {
                block.name = value ? value.trim() : '';
            }

            if (groupCode === 0 && value !== 'BLOCK') {
                const entity = this.parseEntity(lines, i);
                if (entity) {
                    block.entities.push(entity);
                    i = entity.endIndex;
                    continue;
                }
            }
            
            i += 2;
        }
        
        if (!block.endIndex) {
            block.endIndex = i;
        }
        
        return block;
    }
    
    /**
     * Convert value based on group code type
     */
    convertValue(groupCode, value) {
        if (value === undefined || value === null) return null;
        
        const str = value.trim();
        
        // String codes
        if ((groupCode >= 0 && groupCode <= 9) || 
            (groupCode >= 100 && groupCode <= 102) ||
            (groupCode >= 300 && groupCode <= 369)) {
            return str;
        }
        
        // Integer codes
        if ((groupCode >= 60 && groupCode <= 99) ||
            (groupCode >= 160 && groupCode <= 179) ||
            (groupCode >= 270 && groupCode <= 289) ||
            (groupCode >= 370 && groupCode <= 389) ||
            (groupCode >= 390 && groupCode <= 399)) {
            return parseInt(str, 10);
        }
        
        // Real/float codes
        if ((groupCode >= 10 && groupCode <= 59) ||
            (groupCode >= 110 && groupCode <= 149) ||
            (groupCode >= 210 && groupCode <= 239) ||
            (groupCode >= 260 && groupCode <= 269) ||
            (groupCode >= 280 && groupCode <= 289) ||
            (groupCode >= 360 && groupCode <= 369) ||
            (groupCode >= 400 && groupCode <= 409) ||
            (groupCode >= 410 && groupCode <= 419) ||
            (groupCode >= 420 && groupCode <= 429) ||
            (groupCode >= 430 && groupCode <= 439) ||
            (groupCode >= 440 && groupCode <= 459) ||
            (groupCode >= 460 && groupCode <= 469) ||
            (groupCode >= 470 && groupCode <= 481)) {
            return parseFloat(str);
        }
        
        // Default: try float first, then int, then string
        const floatVal = parseFloat(str);
        if (!isNaN(floatVal)) return floatVal;
        return str;
    }
    
    /**
     * Get unit name from INSUNITS code — returns lowercase key for convertToCm lookup
     */
    getUnits(insunits) {
        const unitMap = {
            0: 'mm',
            1: 'in',
            2: 'ft',
            3: 'mm',
            4: 'mm',
            5: 'cm',
            6: 'm',
            7: 'm',
            8: 'in',
            9: 'in',
            10: 'yd',
            11: 'mm',
            12: 'mm',
            13: 'mm',
            14: 'mm',
            15: 'm',
            16: 'm',
            17: 'm',
            18: 'm',
            19: 'm',
            20: 'm'
        };
        return unitMap[insunits] || 'mm';
    }
    
    /**
     * Convert length from detected units to millimeters (legacy — not used by new app.js)
     */
    toMillimeters(length, unit) {
        const conversions = {
            'Unitless': 1,
            'Inches': 25.4,
            'Feet': 304.8,
            'Miles': 1609344,
            'Millimeters': 1,
            'Centimeters': 10,
            'Meters': 1000,
            'Kilometers': 1000000,
            'Microinches': 0.0000254,
            'Mils': 0.0254,
            'Yards': 914.4,
            'Angstroms': 0.0000001,
            'Nanometers': 0.000001,
            'Microns': 0.001,
            'Decimeters': 100,
            'Tens of Meters': 10000,
            'Hectometers': 100000,
            'Gigameters': 1e12,
            'Astronomical Units': 149597870700000,
            'Light Years': 9.4607e18,
            'Parsecs': 3.0857e19
        };
        return length * (conversions[unit] || 1);
    }
}

// Export for browser and Node.js
if (typeof module !== 'undefined' && module.exports) {
    module.exports = DXFParser;
}
