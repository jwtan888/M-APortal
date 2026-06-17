#!/usr/bin/env python3
"""
DXF Vector Length Calculator - Audit/Verification Script
This script provides an independent verification of vector length calculations
for critical production measurements.

Usage:
    python dxf-length-audit.py <input.dxf> [--output results.json]
"""

import math
import json
import re
import sys
from typing import Dict, List, Tuple, Optional
from dataclasses import dataclass, asdict
from datetime import datetime


@dataclass
class EntityResult:
    """Result for a single entity"""
    type: str
    layer: str
    length_mm: float
    length_cm: float


@dataclass
class CalculationResult:
    """Complete calculation result"""
    file: str
    timestamp: str
    units_detected: str
    total_length_mm: float
    total_length_cm: float
    total_length_m: float
    total_length_inches: float
    total_length_feet: float
    total_length_yards: float
    entity_count: int
    breakdown: Dict[str, Dict]
    entities: List[EntityResult]


class DXFAuditParser:
    """
    Independent DXF parser for audit verification
    Implements the same algorithms as the JavaScript version for cross-verification
    """
    
    def __init__(self):
        self.entities = []
        self.units = 'Millimeters'
        self.header_vars = {}
    
    def parse_file(self, filepath: str) -> 'DXFAuditParser':
        """Parse DXF file"""
        with open(filepath, 'r', encoding='utf-8', errors='ignore') as f:
            content = f.read()
        
        return self.parse_content(content)
    
    def parse_content(self, content: str) -> 'DXFAuditParser':
        """Parse DXF content string"""
        lines = content.split('\n')
        i = 0
        
        in_entities = False
        
        while i < len(lines) - 1:
            try:
                group_code = int(lines[i].strip())
                value = lines[i + 1].strip()
            except (ValueError, IndexError):
                i += 1
                continue
            
            # Check for section start
            if group_code == 0 and value == 'SECTION':
                i += 2
                section_name = self._read_value(lines, i, 2)
                
                if section_name == 'HEADER':
                    i = self._parse_header_section(lines, i)
                    continue
                elif section_name == 'ENTITIES':
                    in_entities = True
                    i += 2
                    continue
                elif section_name == 'ENDSEC':
                    in_entities = False
                    i += 2
                    continue
            
            # Parse entities
            if in_entities and group_code == 0:
                entity = self._parse_entity(lines, i)
                if entity:
                    self.entities.append(entity)
                    i = entity.get('end_index', i + 2)
                    continue
            
            i += 2
        
        return self
    
    def _read_value(self, lines: List[str], index: int, offset: int = 0) -> str:
        """Safely read a value from lines"""
        idx = index + offset
        if idx < len(lines):
            return lines[idx].strip()
        return ''
    
    def _parse_header_section(self, lines: List[str], start: int) -> int:
        """Parse HEADER section to detect units"""
        i = start + 2
        
        while i < len(lines) - 1:
            try:
                gc = int(lines[i].strip())
                val = lines[i + 1].strip()
            except (ValueError, IndexError):
                i += 1
                continue
            
            if gc == 0 and val == 'ENDSEC':
                return i + 2
            
            if gc == 9:  # Variable name
                var_name = val
                i += 2
                
                if i < len(lines) - 1:
                    try:
                        var_gc = int(lines[i].strip())
                        var_val = self._convert_value(var_gc, lines[i + 1])
                        self.header_vars[var_name] = var_val
                        
                        # Detect units
                        if var_name == '$INSUNITS' and isinstance(var_val, int):
                            self.units = self._get_unit_name(var_val)
                    except ValueError:
                        pass
            
            i += 2
        
        return i
    
    def _parse_entity(self, lines: List[str], start: int) -> Optional[Dict]:
        """Parse a single entity"""
        try:
            entity_type = lines[start + 1].strip()
        except IndexError:
            return None
        
        if entity_type in ('ENDSEC', 'EOF', 'BLOCK', 'ENDBLK'):
            return None
        
        entity = {
            'type': entity_type,
            'layer': '0'
        }
        
        i = start + 2
        current_vertex = None
        vertices = []
        control_points = []
        current_point = None
        
        while i < len(lines) - 1:
            try:
                gc = int(lines[i].strip())
                val_str = lines[i + 1].strip()
            except (ValueError, IndexError):
                i += 1
                continue
            
            if gc == 0:
                # End of entity
                if vertices:
                    entity['vertices'] = vertices
                if control_points:
                    entity['control_points'] = control_points
                entity['end_index'] = i
                return entity
            
            val = self._convert_value(gc, val_str)
            
            # Parse common properties
            if gc == 8:
                entity['layer'] = val
            elif gc == 10:
                entity['x'] = val
            elif gc == 20:
                entity['y'] = val
            elif gc == 30:
                entity['z'] = val
            elif gc == 11:
                entity['x1'] = val
            elif gc == 21:
                entity['y1'] = val
            elif gc == 31:
                entity['z1'] = val
            elif gc == 12:
                entity['x2'] = val
            elif gc == 22:
                entity['y2'] = val
            elif gc == 40:
                entity['radius'] = val
            elif gc == 41:
                entity['start_angle'] = val
            elif gc == 42:
                entity['end_angle'] = val
            elif gc == 50:
                entity['start_angle'] = val
            elif gc == 51:
                entity['end_angle'] = val
            
            # Handle circle/arc center
            if entity_type in ('CIRCLE', 'ARC'):
                if gc == 10:
                    entity['cx'] = val
                elif gc == 20:
                    entity['cy'] = val
            
            # Handle ellipse
            if entity_type == 'ELLIPSE':
                if gc == 10:
                    entity['cx'] = val
                elif gc == 20:
                    entity['cy'] = val
                elif gc == 11:
                    entity['major_axis_x'] = val
                elif gc == 21:
                    entity['major_axis_y'] = val
                elif gc == 40:
                    entity['axis_ratio'] = val
            
            # Handle polyline vertices
            if entity_type in ('POLYLINE', 'LWPOLYLINE'):
                if gc == 10:
                    current_vertex = {'x': val}
                elif gc == 20:
                    if current_vertex:
                        current_vertex['y'] = val
                        vertices.append(current_vertex)
                        current_vertex = None
                elif gc == 40 and vertices:
                    vertices[-1]['bulge'] = val
            
            # Handle spline control points
            if entity_type == 'SPLINE':
                if gc == 10:
                    current_point = {'x': val}
                elif gc == 20:
                    if current_point:
                        current_point['y'] = val
                        control_points.append(current_point)
                        current_point = None
            
            i += 2
        
        # End of file reached
        if vertices:
            entity['vertices'] = vertices
        if control_points:
            entity['control_points'] = control_points
        entity['end_index'] = i
        return entity
    
    def _convert_value(self, group_code: int, value: str):
        """Convert string value based on group code"""
        try:
            # String codes
            if (0 <= group_code <= 9) or (100 <= group_code <= 102) or (300 <= group_code <= 369):
                return value
            
            # Integer codes
            if (60 <= group_code <= 99) or (160 <= group_code <= 179) or \
               (270 <= group_code <= 289) or (370 <= group_code <= 389):
                return int(value)
            
            # Float codes (default for most numeric values)
            return float(value)
        except ValueError:
            return value
    
    def _get_unit_name(self, insunits: int) -> str:
        """Get unit name from INSUNITS code"""
        unit_map = {
            0: 'Unitless',
            1: 'Inches',
            2: 'Feet',
            3: 'Miles',
            4: 'Millimeters',
            5: 'Centimeters',
            6: 'Meters',
            7: 'Kilometers',
            10: 'Yards'
        }
        return unit_map.get(insunits, 'Millimeters')


class VectorLengthCalculator:
    """
    Calculate vector lengths from parsed DXF entities
    Uses the same mathematical formulas as the JavaScript version
    """
    
    @staticmethod
    def calculate_line_length(entity: Dict) -> float:
        """Calculate LINE entity length using Euclidean distance"""
        dx = entity.get('x1', 0) - entity.get('x', 0)
        dy = entity.get('y1', 0) - entity.get('y', 0)
        dz = entity.get('z1', 0) - entity.get('z', 0)
        return math.sqrt(dx*dx + dy*dy + dz*dz)
    
    @staticmethod
    def calculate_circle_length(entity: Dict) -> float:
        """Calculate CIRCLE circumference: 2πr"""
        radius = entity.get('radius', 0)
        return 2 * math.pi * radius
    
    @staticmethod
    def calculate_arc_length(entity: Dict) -> float:
        """Calculate ARC length: r × θ (θ in radians)"""
        radius = entity.get('radius', 0)
        start_angle = entity.get('start_angle', 0) or 0
        end_angle = entity.get('end_angle', 0) or 0
        
        # Normalize angles
        while start_angle < 0:
            start_angle += 2 * math.pi
        while end_angle < 0:
            end_angle += 2 * math.pi
        
        angle_diff = end_angle - start_angle
        if angle_diff < 0:
            angle_diff += 2 * math.pi
        
        return radius * angle_diff
    
    @staticmethod
    def calculate_bulge_arc_length(v1: Dict, v2: Dict) -> float:
        """Calculate arc length from bulge value"""
        bulge = v1.get('bulge', 0)
        if not bulge or abs(bulge) < 0.0001:
            return VectorLengthCalculator.calculate_line_length({
                'x': v1['x'], 'y': v1['y'], 'z': 0,
                'x1': v2['x'], 'y1': v2['y'], 'z1': 0
            })
        
        chord_length = math.sqrt(
            (v2['x'] - v1['x'])**2 + (v2['y'] - v1['y'])**2
        )
        
        if chord_length < 0.0001:
            return 0
        
        included_angle = 4 * math.atan(abs(bulge))
        radius = chord_length / (2 * math.sin(included_angle / 2))
        
        return radius * included_angle
    
    @classmethod
    def calculate_polyline_length(cls, entity: Dict) -> float:
        """Calculate POLYLINE/LWPOLYLINE length"""
        vertices = entity.get('vertices', [])
        if len(vertices) < 2:
            return 0
        
        total_length = 0
        
        for i in range(len(vertices) - 1):
            v1 = vertices[i]
            v2 = vertices[i + 1]
            
            if v1.get('bulge') and abs(v1['bulge']) > 0.0001:
                total_length += cls.calculate_bulge_arc_length(v1, v2)
            else:
                total_length += cls.calculate_line_length({
                    'x': v1['x'], 'y': v1['y'], 'z': v1.get('z', 0),
                    'x1': v2['x'], 'y1': v2['y'], 'z1': v2.get('z', 0)
                })
        
        # Check if closed
        is_closed = entity.get('closed', False) or (
            len(vertices) > 2 and
            abs(vertices[0]['x'] - vertices[-1]['x']) < 0.0001 and
            abs(vertices[0]['y'] - vertices[-1]['y']) < 0.0001
        )
        
        if is_closed and len(vertices) > 2:
            v1 = vertices[-1]
            v2 = vertices[0]
            
            if v1.get('bulge') and abs(v1['bulge']) > 0.0001:
                total_length += cls.calculate_bulge_arc_length(v1, v2)
            else:
                total_length += cls.calculate_line_length({
                    'x': v1['x'], 'y': v1['y'], 'z': v1.get('z', 0),
                    'x1': v2['x'], 'y1': v2['y'], 'z1': v2.get('z', 0)
                })
        
        return total_length
    
    @staticmethod
    def calculate_spline_length(entity: Dict, samples: int = 100) -> float:
        """Calculate SPLINE length using numerical integration"""
        control_points = entity.get('control_points', [])
        if len(control_points) < 2:
            return 0
        
        total_length = 0
        
        for i in range(len(control_points) - 1):
            p1 = control_points[i]
            p2 = control_points[i + 1]
            
            for j in range(samples):
                t1 = j / samples
                t2 = (j + 1) / samples
                
                x1 = p1['x'] + (p2['x'] - p1['x']) * t1
                y1 = p1['y'] + (p2['y'] - p1['y']) * t1
                x2 = p1['x'] + (p2['x'] - p1['x']) * t2
                y2 = p1['y'] + (p2['y'] - p1['y']) * t2
                
                total_length += math.sqrt((x2 - x1)**2 + (y2 - y1)**2)
        
        return total_length
    
    @staticmethod
    def calculate_ellipse_length(entity: Dict) -> float:
        """Calculate ELLIPSE length using Ramanujan's approximation"""
        major_axis_x = entity.get('major_axis_x', 0) or 0
        major_axis_y = entity.get('major_axis_y', 0) or 0
        major_axis_length = math.sqrt(major_axis_x**2 + major_axis_y**2)
        
        axis_ratio = entity.get('axis_ratio', 1) or 1
        minor_axis_length = major_axis_length * axis_ratio
        
        a = major_axis_length / 2  # Semi-major axis
        b = minor_axis_length / 2  # Semi-minor axis
        
        # Ramanujan's second approximation
        h = ((a - b) ** 2) / ((a + b) ** 2)
        circumference = math.pi * (a + b) * (1 + (3 * h) / (10 + math.sqrt(4 - 3 * h)))
        
        # Check for partial ellipse
        start_param = entity.get('start_param', 0) or 0
        end_param = entity.get('end_param', 2 * math.pi) or (2 * math.pi)
        param_range = end_param - start_param
        
        if param_range < 2 * math.pi - 0.0001:
            return circumference * (param_range / (2 * math.pi))
        
        return circumference
    
    @classmethod
    def calculate_entity_length(cls, entity: Dict) -> float:
        """Calculate length for any entity type"""
        entity_type = entity.get('type', '')
        
        if entity_type == 'LINE':
            return cls.calculate_line_length(entity)
        elif entity_type == 'CIRCLE':
            return cls.calculate_circle_length(entity)
        elif entity_type == 'ARC':
            return cls.calculate_arc_length(entity)
        elif entity_type in ('POLYLINE', 'LWPOLYLINE'):
            return cls.calculate_polyline_length(entity)
        elif entity_type == 'SPLINE':
            return cls.calculate_spline_length(entity)
        elif entity_type == 'ELLIPSE':
            return cls.calculate_ellipse_length(entity)
        elif entity_type == 'POINT':
            return 0
        else:
            # Try as line if coordinates exist
            if all(k in entity for k in ['x', 'y', 'x1', 'y1']):
                return cls.calculate_line_length(entity)
            return 0


def calculate_dxf_length(filepath: str, output_file: Optional[str] = None) -> CalculationResult:
    """
    Main function to calculate DXF vector length
    Returns auditable results
    """
    # Parse DXF
    parser = DXFAuditParser()
    parser.parse_file(filepath)
    
    # Calculate lengths
    total_length_mm = 0
    entity_results = []
    breakdown = {}
    
    calculator = VectorLengthCalculator()
    
    for entity in parser.entities:
        length = calculator.calculate_entity_length(entity)
        
        if length > 0:
            total_length_mm += length
            
            # Track by type
            etype = entity.get('type', 'UNKNOWN')
            if etype not in breakdown:
                breakdown[etype] = {'count': 0, 'length_mm': 0}
            breakdown[etype]['count'] += 1
            breakdown[etype]['length_mm'] += length
            
            # Store individual result
            entity_results.append(EntityResult(
                type=etype,
                layer=entity.get('layer', '0'),
                length_mm=round(length, 4),
                length_cm=round(length / 10, 4)
            ))
    
    # Create result object
    result = CalculationResult(
        file=filepath,
        timestamp=datetime.now().isoformat(),
        units_detected=parser.units,
        total_length_mm=round(total_length_mm, 4),
        total_length_cm=round(total_length_mm / 10, 4),
        total_length_m=round(total_length_mm / 1000, 4),
        total_length_inches=round(total_length_mm / 25.4, 4),
        total_length_feet=round(total_length_mm / 25.4 / 12, 4),
        total_length_yards=round(total_length_mm / 25.4 / 12 / 3, 4),
        entity_count=len(entity_results),
        breakdown={k: {'count': v['count'], 'length_mm': round(v['length_mm'], 4)} 
                   for k, v in breakdown.items()},
        entities=entity_results
    )
    
    # Output results
    print("\n" + "="*60)
    print("DXF VECTOR LENGTH CALCULATION - AUDIT REPORT")
    print("="*60)
    print(f"\nFile: {filepath}")
    print(f"Units Detected: {parser.units}")
    print(f"Total Entities Processed: {result.entity_count}")
    print("\n--- TOTAL VECTOR LENGTH ---")
    print(f"Centimeters: {result.total_length_cm:.4f} cm")
    print(f"Millimeters: {result.total_length_mm:.4f} mm")
    print(f"Meters:      {result.total_length_m:.4f} m")
    print(f"Inches:      {result.total_length_inches:.4f}\"")
    print(f"Feet:        {result.total_length_feet:.4f}'")
    print(f"Yards:       {result.total_length_yards:.4f} yd")
    
    print("\n--- ENTITY BREAKDOWN ---")
    for etype, data in breakdown.items():
        print(f"{etype}: {data['count']} entities, {data['length_mm']/10:.2f} cm")
    
    # Save to JSON if requested
    if output_file:
        result_dict = {
            'file': result.file,
            'timestamp': result.timestamp,
            'units_detected': result.units_detected,
            'totals': {
                'centimeters': result.total_length_cm,
                'millimeters': result.total_length_mm,
                'meters': result.total_length_m,
                'inches': result.total_length_inches,
                'feet': result.total_length_feet,
                'yards': result.total_length_yards
            },
            'entity_count': result.entity_count,
            'breakdown': result.breakdown,
            'entities': [asdict(e) for e in result.entities]
        }
        
        with open(output_file, 'w') as f:
            json.dump(result_dict, f, indent=2)
        
        print(f"\n✓ Results saved to: {output_file}")
    
    print("\n" + "="*60)
    
    return result


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("Usage: python dxf-length-audit.py <input.dxf> [--output results.json]")
        sys.exit(1)
    
    input_file = sys.argv[1]
    output_file = None
    
    if '--output' in sys.argv:
        idx = sys.argv.index('--output')
        if idx + 1 < len(sys.argv):
            output_file = sys.argv[idx + 1]
    
    try:
        calculate_dxf_length(input_file, output_file)
    except FileNotFoundError:
        print(f"Error: File '{input_file}' not found")
        sys.exit(1)
    except Exception as e:
        print(f"Error: {e}")
        sys.exit(1)
