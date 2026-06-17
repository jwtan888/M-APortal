/**
 * DXF Parser - Pure JavaScript DXF file parser
 * Parses DXF files and extracts entities for vector length calculation
 */

class DXFParser {
    constructor() {
        this.data = null;
        this.entities = [];
        this.units = 'mm'; // Default to millimeters
    }

    /**
     * Parse DXF file content (as text)
     * @param {string} content - DXF file content as string
     * @returns {Object} Parsed DXF data
     */
    parse(content) {
        const lines = content.split(/\r?\n/);
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
                    }
                    i = currentSection.endIndex;
                    continue;
                }
            }
            
            i += 2;
        }
        
        return this.data;
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
            
            i += 2;
        }
        
        return {
            name: sectionName,
            data: data,
            entities: entities,
            blocks: blocks,
            endIndex: i
        };
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
        else if (groupCode === 370) entity.lineweight = parsedValue;
        else if (groupCode === 48) entity.dashScale = parsedValue;
        else if (groupCode === 73) entity.direction = parsedValue;
        
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
    }
    
    /**
     * Parse a BLOCK section
     */
    parseBlock(lines, startIndex) {
        const block = { entities: [] };
        let i = startIndex + 2;
        
        while (i < lines.length) {
            const groupCode = parseInt(lines[i]?.trim(), 10);
            const value = lines[i + 1];
            
            if (groupCode === 0 && value === 'ENDBLK') {
                block.endIndex = i + 2;
                break;
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
     * Get unit name from INSUNITS code
     */
    getUnits(insunits) {
        const unitMap = {
            0: 'Unitless',
            1: 'Inches',
            2: 'Feet',
            3: 'Miles',
            4: 'Millimeters',
            5: 'Centimeters',
            6: 'Meters',
            7: 'Kilometers',
            8: 'Microinches',
            9: 'Mils',
            10: 'Yards',
            11: 'Angstroms',
            12: 'Nanometers',
            13: 'Microns',
            14: 'Decimeters',
            15: 'Tens of Meters',
            16: 'Hectometers',
            17: 'Gigameters',
            18: 'Astronomical Units',
            19: 'Light Years',
            20: 'Parsecs'
        };
        return unitMap[insunits] || 'Millimeters';
    }
    
    /**
     * Convert length from detected units to millimeters
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
