/* marker-nesting.js — DXF → polygon conversion, nesting engine, marker rendering */
(function() {
  const SEGMENTS_PER_FULL = 128;

  function vec(x, y) { return { x, y }; }

  function approxEquals(a, b, tol) {
    tol = tol || 1e-6;
    return Math.abs(a.x - b.x) < tol && Math.abs(a.y - b.y) < tol;
  }

  function cleanPolygon(verts) {
    if (verts.length < 3) return verts;
    const out = [verts[0]];
    for (let i = 1; i < verts.length; i++) {
      if (!approxEquals(verts[i], verts[i - 1])) out.push(verts[i]);
    }
    if (approxEquals(out[0], out[out.length - 1])) out.pop();
    return out;
  }

  function arcVertices(cx, cy, r, startAngle, endAngle, segments) {
    let sa = startAngle || 0, ea = endAngle || 0;
    while (sa < 0) sa += 2 * Math.PI;
    while (ea < 0) ea += 2 * Math.PI;
    let diff = ea - sa;
    if (diff < 0) diff += 2 * Math.PI;

    const n = Math.max(3, Math.ceil(segments * diff / (2 * Math.PI)));
    const verts = [];
    for (let i = 0; i <= n; i++) {
      const theta = sa + diff * i / n;
      verts.push(vec(cx + r * Math.cos(theta), cy + r * Math.sin(theta)));
    }
    return verts;
  }

  function ellipseVertices(cx, cy, majorAxisX, majorAxisY, axisRatio, startParam, endParam, segments) {
    const mx = majorAxisX || 0, my = majorAxisY || 0;
    const majorLen = Math.sqrt(mx * mx + my * my);
    const angle = Math.atan2(my, mx);
    const a = majorLen / 2;
    const b = (majorLen / 2) * (axisRatio || 1);

    let sp = startParam || 0, ep = endParam || 0;
    if (ep === 0 && sp === 0) ep = 2 * Math.PI;
    let diff = ep - sp;
    if (diff <= 0) diff += 2 * Math.PI;

    const n = Math.max(3, Math.ceil(segments * diff / (2 * Math.PI)));
    const verts = [];
    for (let i = 0; i <= n; i++) {
      const t = sp + diff * i / n;
      const cosT = Math.cos(t), sinT = Math.sin(t);
      const x = cx + a * cosT * Math.cos(angle) - b * sinT * Math.sin(angle);
      const y = cy + a * cosT * Math.sin(angle) + b * sinT * Math.cos(angle);
      verts.push(vec(x, y));
    }
    return verts;
  }

  // ── B-spline evaluation helpers (Cox-de Boor) ────────────
  function findSpan(n, p, u, knots) {
    if (u >= knots[n + 1]) return n;
    if (u <= knots[p]) return p;
    let low = p, high = n + 1;
    let mid = Math.floor((low + high) / 2);
    while (u < knots[mid] || u >= knots[mid + 1]) {
      if (u < knots[mid]) high = mid;
      else low = mid;
      mid = Math.floor((low + high) / 2);
    }
    return mid;
  }

  function basisFunctions(span, u, p, knots) {
    const N = new Array(p + 1);
    const left = new Array(p + 1);
    const right = new Array(p + 1);
    N[0] = 1;
    for (let j = 1; j <= p; j++) {
      left[j] = u - knots[span + 1 - j];
      right[j] = knots[span + j] - u;
      let saved = 0;
      for (let r = 0; r < j; r++) {
        const denom = right[r + 1] + left[j - r];
        const temp = (denom !== 0) ? N[r] / denom : 0;
        N[r] = saved + right[r + 1] * temp;
        saved = left[j - r] * temp;
      }
      N[j] = saved;
    }
    return N;
  }

  function evalBSpline(u, degree, knots, controlPoints, weights) {
    const n = controlPoints.length - 1;
    const span = findSpan(n, degree, u, knots);
    const N = basisFunctions(span, u, degree, knots);
    let x = 0, y = 0, wSum = 0;
    for (let i = 0; i <= degree; i++) {
      const idx = span - degree + i;
      if (idx >= 0 && idx < controlPoints.length) {
        const w = (weights && weights[idx] != null) ? weights[idx] : 1;
        const b = N[i] * w;
        x += controlPoints[idx].x * b;
        y += controlPoints[idx].y * b;
        wSum += b;
      }
    }
    if (wSum !== 0) { return { x: x / wSum, y: y / wSum }; }
    x = 0; y = 0;
    for (let i = 0; i <= degree; i++) {
      const idx = span - degree + i;
      if (idx >= 0 && idx < controlPoints.length) {
        x += controlPoints[idx].x * N[i];
        y += controlPoints[idx].y * N[i];
      }
    }
    return { x, y };
  }

  function generateClampedKnots(n, p) {
    const knots = [];
    const interior = n + 1 - p;
    for (let i = 0; i <= p; i++) knots.push(0);
    for (let i = 1; i < interior; i++) knots.push(i / interior);
    for (let i = 0; i <= p; i++) knots.push(1);
    return knots;
  }

  function splineVertices(controlPoints, segments, degree, knots, weights) {
    const pts = controlPoints || [];
    if (pts.length < 2) return [];
    const p = (degree != null && degree > 0 && degree < pts.length) ? degree : Math.min(3, pts.length - 1);
    const kv = (knots && knots.length >= pts.length + p + 1) ? knots : generateClampedKnots(pts.length - 1, p);
    const uStart = kv[p];
    const uEnd = kv[kv.length - p - 1];
    const uRange = uEnd - uStart;
    const totalSamples = Math.max(segments * Math.max(pts.length - p, 1), 32);
    const verts = [];
    if (uRange <= 0) {
      verts.push(evalBSpline(uStart, p, kv, pts, weights));
    } else {
      for (let i = 0; i <= totalSamples; i++) {
        const u = uStart + uRange * i / totalSamples;
        verts.push(evalBSpline(u, p, kv, pts, weights));
      }
    }
    return cleanPolygon(verts);
  }

  function bulgeArcVertices(v1, v2, segments) {
    const bulge = v1.bulge || 0;
    if (Math.abs(bulge) < 0.0001) return [vec(v2.x, v2.y)];

    const chord = Math.sqrt((v2.x - v1.x) ** 2 + (v2.y - v1.y) ** 2);
    if (chord < 0.0001) return [vec(v2.x, v2.y)];

    const angle = 4 * Math.atan(Math.abs(bulge));
    const r = chord / (2 * Math.sin(angle / 2));
    const mx = (v1.x + v2.x) / 2, my = (v1.y + v2.y) / 2;
    const dx = v2.x - v1.x, dy = v2.y - v1.y;
    const h = r * Math.cos(angle / 2);
    const sign = Math.sign(bulge);
    const px = mx + (-dy / chord) * h * sign;
    const py = my + (dx / chord) * h * sign;

    const n = Math.max(2, Math.ceil(segments * angle / (2 * Math.PI)));
    const sweepAngle = angle * sign;
    const startAngle = Math.atan2(v1.y - py, v1.x - px);
    const verts = [];
    for (let i = 1; i <= n; i++) {
      const theta = startAngle + sweepAngle * i / n;
      verts.push(vec(px + r * Math.cos(theta), py + r * Math.sin(theta)));
    }
    return verts;
  }

  function pathVertices(pathSegments) {
    const segs = pathSegments || [];
    if (segs.length === 0) return [];

    const verts = [];
    let cx = 0, cy = 0;

    for (let i = 0; i < segs.length; i++) {
      const seg = segs[i];
      switch (seg.command) {
        case 'M':
          cx = seg.x || 0; cy = seg.y || 0;
          verts.push(vec(cx, cy));
          break;
        case 'L':
          cx = seg.x || 0; cy = seg.y || 0;
          verts.push(vec(cx, cy));
          break;
        case 'C': {
          const x0 = cx, y0 = cy;
          const x1 = seg.x1 || 0, y1 = seg.y1 || 0;
          const x2 = seg.x2 || 0, y2 = seg.y2 || 0;
          const x3 = seg.x || 0, y3 = seg.y || 0;
          const n = 12;
          for (let j = 1; j <= n; j++) {
            const t = j / n;
            const mt = 1 - t;
            const x = mt * mt * mt * x0 + 3 * mt * mt * t * x1 + 3 * mt * t * t * x2 + t * t * t * x3;
            const y = mt * mt * mt * y0 + 3 * mt * mt * t * y1 + 3 * mt * t * t * y2 + t * t * t * y3;
            verts.push(vec(x, y));
          }
          cx = x3; cy = y3;
          break;
        }
        case 'Q': {
          const x0 = cx, y0 = cy;
          const x1 = seg.x1 || 0, y1 = seg.y1 || 0;
          const x2 = seg.x || 0, y2 = seg.y || 0;
          const n = 8;
          for (let j = 1; j <= n; j++) {
            const t = j / n;
            const mt = 1 - t;
            const x = mt * mt * x0 + 2 * mt * t * x1 + t * t * x2;
            const y = mt * mt * y0 + 2 * mt * t * y1 + t * t * y2;
            verts.push(vec(x, y));
          }
          cx = x2; cy = y2;
          break;
        }
        case 'A': {
          // Arc command from SVG
          break;
        }
        case 'Z':
          // Close path — handled by caller
          break;
      }
    }
    return verts;
  }

  function hatchVertices(entity, segsPerFull) {
    const segs = segsPerFull || SEGMENTS_PER_FULL;
    const edges = entity.boundaryEdges || [];
    const allVerts = [];
    edges.forEach(edge => {
      switch (edge.edgeType) {
        case 1:
          allVerts.push(vec(edge.x || 0, edge.y || 0));
          allVerts.push(vec(edge.x1 || 0, edge.y1 || 0));
          break;
        case 2:
          allVerts.push(...arcVertices(
            edge.x || 0, edge.y || 0, edge.radius || 0,
            (edge.startAngle || 0) * Math.PI / 180,
            (edge.endAngle || 360) * Math.PI / 180,
            segs
          ));
          break;
        case 3:
          allVerts.push(...ellipseVertices(
            edge.x || 0, edge.y || 0,
            edge.majorAxisX || 0, edge.majorAxisY || 0,
            edge.axisRatio || 1, 0, 2 * Math.PI,
            segs
          ));
          break;
        case 4: {
          const cp = [];
          if (edge.fitPointX != null) {
            cp.push(vec(edge.fitPointX, edge.fitPointY));
          }
          if (edge.fitPoints) {
            edge.fitPoints.forEach(fp => cp.push(vec(fp.x, fp.y)));
          }
          if (cp.length < 2 && edge.x != null && edge.x1 != null) {
            cp.push(vec(edge.x, edge.y));
            cp.push(vec(edge.x1, edge.y1));
          }
          if (cp.length >= 2) {
            allVerts.push(...splineVertices(cp, Math.max(segs / 8, 8), edge.degree || 3));
          }
          break;
        }
      }
    });
    return allVerts;
  }

  function entityToVerts(entity, segsPerFull) {
    const segs = segsPerFull || SEGMENTS_PER_FULL;
    switch (entity.type) {
      case 'LINE':
        return [vec(entity.x || 0, entity.y || 0), vec(entity.x1 || 0, entity.y1 || 0)];

      case 'CIRCLE':
        return arcVertices(entity.cx || 0, entity.cy || 0, entity.radius || 0, 0, 2 * Math.PI, segs);

      case 'ARC':
        return arcVertices(
          entity.cx || 0, entity.cy || 0, entity.radius || 0,
          entity.startAngle || 0, entity.endAngle || 0,
          segs
        );

      case 'POLYLINE':
      case 'LWPOLYLINE': {
        const verts = entity.vertices || [];
        if (verts.length < 2) return [];
        const out = [vec(verts[0].x, verts[0].y)];
        for (let i = 1; i < verts.length; i++) {
          if (verts[i - 1].bulge && Math.abs(verts[i - 1].bulge) > 0.0001) {
            out.push(...bulgeArcVertices(verts[i - 1], verts[i], segs));
          } else {
            out.push(vec(verts[i].x, verts[i].y));
          }
        }
        if (entity.closed) {
          const last = verts[verts.length - 1];
          const first = verts[0];
          if (last.bulge && Math.abs(last.bulge) > 0.0001) {
            out.push(...bulgeArcVertices(last, first, segs));
          } else if (!approxEquals(out[out.length - 1], out[0])) {
            out.push(vec(first.x, first.y));
          }
        }
        return out;
      }

      case 'SPLINE':
        return splineVertices(
          entity.controlPoints, Math.max(segs / 8, 8),
          entity.degree, entity.knots, entity.weights
        );

      case 'ELLIPSE':
        return ellipseVertices(
          entity.cx || 0, entity.cy || 0,
          entity.majorAxisX || 0, entity.majorAxisY || 0,
          entity.axisRatio || 1,
          entity.startParam || 0, entity.endParam || 0,
          segs
        );

      case 'HATCH':
        return hatchVertices(entity, segs);

      case 'PATH':
        return pathVertices(entity.pathSegments);

      default:
        return [];
    }
  }

  function mergeConnectedVerts(allVerts) {
    if (allVerts.length === 0) return [];
    let current = allVerts[0].slice();
    const merged = [];

    for (let i = 1; i < allVerts.length; i++) {
      const next = allVerts[i];
      if (next.length > 0 && current.length > 0) {
        if (approxEquals(current[current.length - 1], next[0])) {
          current.push(...next.slice(1));
        } else {
          merged.push(cleanPolygon(current));
          current = next.slice();
        }
      }
    }
    if (current.length > 0) merged.push(cleanPolygon(current));
    return merged;
  }

  function computeArea(verts) {
    if (verts.length < 3) return 0;
    let area = 0;
    for (let i = 0; i < verts.length; i++) {
      const j = (i + 1) % verts.length;
      area += verts[i].x * verts[j].y;
      area -= verts[j].x * verts[i].y;
    }
    return Math.abs(area) / 2;
  }

  function getUnitToMm(dxfUnit) {
    const map = { 'mm': 1, 'cm': 10, 'm': 1000, 'in': 25.4, 'ft': 304.8, 'yd': 914.4, 'px': 0.264583, 'pt': 0.352778 };
    return map[dxfUnit] || 1;
  }

  function forEachPoint(shape, cb) {
    if (!shape) return;
    if (Array.isArray(shape)) { shape.forEach(cb); return; }
    if (shape._sP) { shape._sP.forEach(function(sp) { sp.forEach(cb); }); return; }
  }

  function getShapeBB(shape) {
    var minX = Infinity, minY = Infinity, maxX = -Infinity, maxY = -Infinity;
    forEachPoint(shape, function(p) {
      if (p.x < minX) minX = p.x;
      if (p.y < minY) minY = p.y;
      if (p.x > maxX) maxX = p.x;
      if (p.y > maxY) maxY = p.y;
    });
    return isFinite(minX) ? { minX: minX, minY: minY, maxX: maxX, maxY: maxY, width: maxX - minX, height: maxY - minY } : null;
  }

  function shiftShape(shape, dx, dy) {
    if (!shape) return shape;
    if (Array.isArray(shape)) return shape.map(function(v) { return { x: v.x + dx, y: v.y + dy }; });
    if (shape._sP) return { _sP: shape._sP.map(function(sp) { return sp.map(function(v) { return { x: v.x + dx, y: v.y + dy }; }); }) };
    return shape;
  }

  function cloneShape(shape) {
    if (!shape) return shape;
    if (Array.isArray(shape)) return shape.map(function(v) { return { x: v.x, y: v.y }; });
    if (shape._sP) return { _sP: shape._sP.map(function(sp) { return sp.map(function(v) { return { x: v.x, y: v.y }; }); }) };
    return shape;
  }

  window.MarkerNesting = {

    convertEntitiesToParts: function(entities, options) {
      options = options || {};
      const unitToMm = getUnitToMm(options.units || 'mm');
      const segsPerFull = options.segments || SEGMENTS_PER_FULL;
      const entityColors = options.entityColors || [];
      const entityOrigIndices = options.entityOrigIndices || [];

      const allLoops = [];
      const loopEntityIdx = [];
      entities.forEach((entity, ei) => {
        const verts = entityToVerts(entity, segsPerFull);
        if (verts.length >= 3) {
          allLoops.push(verts);
          loopEntityIdx.push(ei);
        }
      });

      const polygons = mergeConnectedVerts(allLoops);

      let loopCursor = 0;
      const polyColors = [];
      const polyOrigIdx = [];
      for (let pi = 0; pi < polygons.length; pi++) {
        var fei = loopEntityIdx[Math.min(loopCursor, loopEntityIdx.length - 1)] || 0;
        polyColors.push(entityColors[fei] || '#6b7280');
        polyOrigIdx.push(entityOrigIndices[fei] != null ? entityOrigIndices[fei] : fei);
        var vNeeded = polygons[pi].length;
        var vGot = 0;
        while (loopCursor < allLoops.length && vGot < vNeeded) {
          vGot += vNeeded > vGot && loopCursor > 0 ? allLoops[loopCursor].length - 1 : allLoops[loopCursor].length;
          loopCursor++;
        }
      }

      return polygons.map((verts, idx) => {
        verts = cleanPolygon(verts);
        if (verts.length < 3) return null;

        const scaled = verts.map(v => vec(v.x * unitToMm, v.y * unitToMm));

        const area = computeArea(scaled);
        const bb = {
          minX: Math.min(...scaled.map(v => v.x)),
          minY: Math.min(...scaled.map(v => v.y)),
          maxX: Math.max(...scaled.map(v => v.x)),
          maxY: Math.max(...scaled.map(v => v.y))
        };

        return {
          id: idx + 1,
          points: scaled,
          area,
          boundingBox: bb,
          width: bb.maxX - bb.minX,
          height: bb.maxY - bb.minY,
          origColor: polyColors[idx] || '#6b7280',
          _origEntityIndex: polyOrigIdx[idx] != null ? polyOrigIdx[idx] : idx
        };
      }).filter(Boolean);
    },

    // ── Self-intersection repair for offset polygons ────────
    // Splits a self-intersecting polygon at the first intersection point
    // and returns the outer boundary loop. Returns null if repair fails.
    repairSelfIntersection: function(pts) {
      var n = pts.length;
      if (n < 4) return pts;
      var EPS = 1e-10;
      function _cross(a, b) { return a.x * b.y - a.y * b.x; }
      function _sub(a, b) { return { x: a.x - b.x, y: a.y - b.y }; }
      for (var i = 0; i < n; i++) {
        var ni = (i + 1) % n;
        for (var j = i + 2; j < n; j++) {
          if ((j + 1) % n === i) continue;
          var nj = (j + 1) % n;
          var p1 = pts[i], p2 = pts[ni];
          var p3 = pts[j], p4 = pts[nj];
          var denom = (p1.x - p2.x) * (p3.y - p4.y) - (p1.y - p2.y) * (p3.x - p4.x);
          if (Math.abs(denom) < EPS) continue;
          var t = ((p1.x - p3.x) * (p3.y - p4.y) - (p1.y - p3.y) * (p3.x - p4.x)) / denom;
          var u = -((p1.x - p2.x) * (p1.y - p3.y) - (p1.y - p2.y) * (p1.x - p3.x)) / denom;
          if (t > EPS && t < 1 - EPS && u > EPS && u < 1 - EPS) {
            var P = { x: p1.x + t * (p2.x - p1.x), y: p1.y + t * (p2.y - p1.y) };
            // Build loop A: P → ni → ... → j → P
            var loopA = [P];
            for (var k = ni; k !== j; k = (k + 1) % n) loopA.push(pts[k]);
            loopA.push(P);
            // Build loop B: P → nj → ... → i → P
            var loopB = [P];
            for (var k = nj; k !== i; k = (k + 1) % n) loopB.push(pts[k]);
            loopB.push(P);
            // Clean duplicate consecutive vertices
            function clean(l) {
              if (l.length < 2) return l;
              var out = [l[0]];
              for (var ci = 1; ci < l.length - 1; ci++) { if (!approxEquals(l[ci], l[ci - 1])) out.push(l[ci]); }
              if (out.length > 1 && approxEquals(out[0], out[out.length - 1])) out.pop();
              return out;
            }
            loopA = clean(loopA);
            loopB = clean(loopB);
            if (loopA.length < 3 || loopB.length < 3) return null;
            // Absolute area of each loop — outer loop has larger area
            function loopArea(l) {
              var a = 0;
              for (var ci = 0; ci < l.length; ci++) { var cj = (ci + 1) % l.length; a += l[ci].x * l[cj].y - l[cj].x * l[ci].y; }
              return Math.abs(a / 2);
            }
            var outer = loopArea(loopA) > loopArea(loopB) ? loopA : loopB;
            var repaired = MarkerNesting.repairSelfIntersection(outer);
            if (repaired && repaired.length >= 3) return repaired;
            return null;
          }
        }
      }
      return pts;
    },

    // ── Polygon offset ──────────────────────────────────────
    // Compute outward offset of a CCW polygon by distance `d` (mm).
    // Uses edge-offset + miter intersection; bevels reflex vertices.
    offsetPolygon: function(pts, d) {
      var n = pts.length;
      if (n < 3) return pts.slice();

      var area = 0;
      for (var i = 0; i < n; i++) { var j = (i + 1) % n; area += pts[i].x * pts[j].y - pts[j].x * pts[i].y; }
      var sign = area > 0 ? 1 : -1;
      var absArea = Math.abs(area / 2);

      function norm(v) { var l = Math.sqrt(v.x * v.x + v.y * v.y); return l > 1e-12 ? { x: v.x / l, y: v.y / l } : { x: 0, y: 0 }; }
      function sub(a, b) { return { x: a.x - b.x, y: a.y - b.y }; }
      function add(a, b) { return { x: a.x + b.x, y: a.y + b.y }; }
      function scale(v, s) { return { x: v.x * s, y: v.y * s }; }
      function cross(a, b) { return a.x * b.y - a.y * b.x; }

      var edgeNorms = [];
      for (var i = 0; i < n; i++) {
        var j = (i + 1) % n;
        var dir = norm(sub(pts[j], pts[i]));
        edgeNorms[i] = { x: sign * dir.y, y: -sign * dir.x };
      }

      var miterLimit = Math.abs(d) * 3;
      var result = [];

      for (var i = 0; i < n; i++) {
        var prev = (i - 1 + n) % n;
        var cur = i;

        var p1 = add(pts[prev], scale(edgeNorms[prev], d));
        var p2 = add(pts[cur], scale(edgeNorms[cur], d));
        var v1 = sub(pts[cur], pts[prev]);
        var v2 = sub(pts[(cur + 1) % n], pts[cur]);

        var denom = cross(v1, v2);
        if (Math.abs(denom) < 1e-12) {
          result.push(add(pts[cur], scale(add(edgeNorms[prev], edgeNorms[cur]), d / 2)));
        } else {
          var diff = sub(p2, p1);
          var t = cross(diff, v2) / denom;
          var pt = add(p1, scale(v1, t));
          var miterLen = Math.sqrt((pt.x - pts[cur].x) * (pt.x - pts[cur].x) + (pt.y - pts[cur].y) * (pt.y - pts[cur].y));
          if (miterLen > miterLimit) {
            result.push(add(pts[cur], scale(edgeNorms[prev], d)));
            result.push(add(pts[cur], scale(edgeNorms[cur], d)));
          } else {
            result.push(pt);
          }
        }
      }

      var offArea = 0;
      for (var i = 0; i < result.length; i++) { var j = (i + 1) % result.length; offArea += result[i].x * result[j].y - result[j].x * result[i].y; }
      offArea = Math.abs(offArea / 2);
      if (d > 0 && offArea < absArea * 0.5) {
        var repaired = MarkerNesting.repairSelfIntersection(result);
        if (repaired && repaired.length >= 3) return repaired;
        var hull = MarkerNesting.convexHull(pts);
        var hullOffset = MarkerNesting.offsetPolygon(hull, d);
        hullOffset._usedConvexHull = hull;
        return hullOffset;
      }
      var rn = result.length;
      for (var i = 0; i < rn; i++) {
        var ni = (i + 1) % rn;
        for (var j = i + 2; j < rn; j++) {
          if ((j + 1) % rn === i) continue;
          var nj = (j + 1) % rn;
          var o1 = cross(sub(result[ni], result[i]), sub(result[j], result[i]));
          var o2 = cross(sub(result[ni], result[i]), sub(result[nj], result[i]));
          var o3 = cross(sub(result[nj], result[j]), sub(result[i], result[j]));
          var o4 = cross(sub(result[nj], result[j]), sub(result[ni], result[j]));
          if (((o1 > 0) !== (o2 > 0)) && ((o3 > 0) !== (o4 > 0))) {
            var repaired = MarkerNesting.repairSelfIntersection(result);
            if (repaired && repaired.length >= 3) return repaired;
            var hull = MarkerNesting.convexHull(pts);
            var hullOffset = MarkerNesting.offsetPolygon(hull, d);
            hullOffset._usedConvexHull = hull;
            return hullOffset;
          }
        }
      }
      return result;
    },

    // ── Polygon Union ───────────────────────────────────────
    // Returns the union of two CCW simple polygons as a single polygon.
    // Returns null if polygons are disjoint (caller should fall back to hull).
    polygonUnion: function(polyA, polyB) {
      polyA = polyA.map(function(v) { return {x: v.x, y: v.y}; });
      polyB = polyB.map(function(v) { return {x: v.x, y: v.y}; });
      function _signedArea(p) { var a = 0; for (var i = 0, j = p.length - 1; i < p.length; j = i++) { a += p[j].x * p[i].y - p[i].x * p[j].y; } return a; }
      if (_signedArea(polyA) < 0) polyA.reverse();
      if (_signedArea(polyB) < 0) polyB.reverse();

      // Point in polygon via ray casting
      function _pip(pt, p) {
        var inside = false;
        for (var i = 0, j = p.length - 1; i < p.length; j = i++) {
          var yi = p[i].y, yj = p[j].y, xi = p[i].x, xj = p[j].x;
          if ((yi > pt.y) !== (yj > pt.y) && pt.x < (xj - xi) * (pt.y - yi) / (yj - yi) + xi) inside = !inside;
        }
        return inside;
      }

      // Segment intersection: returns intersection point + edge params, or null
      function _segInter(a, b, c, d) {
        var denom = (b.x - a.x) * (d.y - c.y) - (b.y - a.y) * (d.x - c.x);
        if (Math.abs(denom) < 1e-12) return null;
        var t = ((c.x - a.x) * (d.y - c.y) - (c.y - a.y) * (d.x - c.x)) / denom;
        var u = ((c.x - a.x) * (b.y - a.y) - (c.y - a.y) * (b.x - a.x)) / denom;
        if (t < -1e-10 || t > 1 + 1e-10 || u < -1e-10 || u > 1 + 1e-10) return null;
        var ot = t < 0 ? 0 : (t > 1 ? 1 : t);
        var ou = u < 0 ? 0 : (u > 1 ? 1 : u);
        return { pt: { x: a.x + ot * (b.x - a.x), y: a.y + ot * (b.y - a.y) }, t: ot, u: ou };
      }

      // Find all intersection points between polyA and polyB edges
      var rawInters = [];
      for (var ai = 0, aj = polyA.length - 1; ai < polyA.length; aj = ai++) {
        for (var bi = 0, bj = polyB.length - 1; bi < polyB.length; bj = bi++) {
          var inter = _segInter(polyA[aj], polyA[ai], polyB[bj], polyB[bi]);
          if (inter) rawInters.push({ pt: inter.pt, edgeA: aj, tA: inter.t, edgeB: bj, tB: inter.u });
        }
      }

      // Filter near-duplicate intersections
      var eps = 1e-6;
      var intersections = [];
      for (var i = 0; i < rawInters.length; i++) {
        var dup = false;
        for (var j = 0; j < intersections.length; j++) {
          if (Math.abs(rawInters[i].pt.x - intersections[j].pt.x) < eps &&
              Math.abs(rawInters[i].pt.y - intersections[j].pt.y) < eps) { dup = true; break; }
        }
        if (!dup) intersections.push(rawInters[i]);
      }

      // No intersections: check containment or disjoint
      if (intersections.length === 0) {
        if (_pip(polyB[0], polyA)) return polyA; // B inside A
        if (_pip(polyA[0], polyB)) return polyB; // A inside B
        return null;
      }

      // --- Build augmented polygon lists with intersections inserted ---
      // Each entry: { pt: {x,y}, isIntersection: bool, partnerIdx: int, entry: bool }
      
      function buildAugmented(poly, inters, edgeKey, tKey) {
        var aug = [];
        for (var i = 0; i < poly.length; i++) {
          var cur = poly[i];
          var next = poly[(i + 1) % poly.length];
          // Find intersections on this edge
          var onEdge = [];
          for (var k = 0; k < inters.length; k++) {
            if (inters[k][edgeKey] === i) {
              onEdge.push({ pt: inters[k].pt, idx: k, t: inters[k][tKey] });
            }
          }
          // Sort by t along edge
          onEdge.sort(function(a, b) { return a.t - b.t; });
          
          // Push current vertex
          aug.push({ x: cur.x, y: cur.y, isIntersection: false, polyIdx: i });
          
          // Push intersection points along this edge
          for (var m = 0; m < onEdge.length; m++) {
            aug.push({ x: onEdge[m].pt.x, y: onEdge[m].pt.y, isIntersection: true, interIdx: onEdge[m].idx });
          }
        }
        // Remove trailing duplicate (last vertex = first vertex for closed poly)
        if (aug.length > 1 && Math.abs(aug[0].x - aug[aug.length-1].x) < eps && Math.abs(aug[0].y - aug[aug.length-1].y) < eps) {
          if (!aug[aug.length-1].isIntersection) aug.pop();
        }
        return aug;
      }

      var augA = buildAugmented(polyA, intersections, 'edgeA', 'tA');
      var augB = buildAugmented(polyB, intersections, 'edgeB', 'tB');

      // --- Classify each intersection ---
      // For each intersection on polyA: is the edge just AFTER the intersection inside polyB?
      // If yes → A is entering B (ENTRY). If no → A is exiting B (EXIT).
      // For union: switch at ENTRY on current polygon.
      
      function classifyIntersections(aug, otherPoly, inters) {
        var result = {};
        for (var i = 0; i < aug.length; i++) {
          if (!aug[i].isIntersection) continue;
          // Find the next non-intersection vertex to determine edge direction
          var nextIdx = (i + 1) % aug.length;
          var dirX = aug[nextIdx].x - aug[i].x;
          var dirY = aug[nextIdx].y - aug[i].y;
          var len = Math.sqrt(dirX * dirX + dirY * dirY);
          if (len < eps) continue;
          var probeX = aug[i].x + (dirX / len) * eps * 10;
          var probeY = aug[i].y + (dirY / len) * eps * 10;
          var inside = _pip({ x: probeX, y: probeY }, otherPoly);
          result[aug[i].interIdx] = inside ? 'entry' : 'exit';
        }
        return result;
      }

      var classA = classifyIntersections(augA, polyB, intersections);
      var classB = classifyIntersections(augB, polyA, intersections);

      // Traverse to build union
      // Rule for UNION: switch at "entry" on current polygon
      var visited = {};
      var result = [];

      function augNextInter(aug, start) {
        for (var i = start + 1; i < aug.length; i++) { if (aug[i].isIntersection) return i; }
        for (var i = 0; i < start; i++) { if (aug[i].isIntersection) return i; }
        return -1;
      }

      function augAddRange(aug, from, to, out) {
        if (from <= to) {
          for (var i = from; i <= to; i++) out.push({ x: aug[i].x, y: aug[i].y });
        } else {
          for (var i = from; i < aug.length; i++) out.push({ x: aug[i].x, y: aug[i].y });
          for (var i = 0; i <= to; i++) out.push({ x: aug[i].x, y: aug[i].y });
        }
      }

      // Helper: if at an intersection, check classification and switch polygons if entry
      function processInter(atAug, atIdx, atClass) {
        if (!atAug[atIdx].isIntersection) return null;
        var interIdx = atAug[atIdx].interIdx;
        var cls = atClass[interIdx];
        if (cls === 'entry') {
          var otherAug = (atAug === augA) ? augB : augA;
          for (var oi = 0; oi < otherAug.length; oi++) {
            if (otherAug[oi].isIntersection && otherAug[oi].interIdx === interIdx) {
              visited[interIdx] = true;
              return { aug: otherAug, cls: (otherAug === augA) ? classA : classB, idx: oi };
            }
          }
        }
        return null;
      }

      // Find unvisited intersection on augA to start
      for (var si = 0; si < augA.length; si++) {
        if (!augA[si].isIntersection) continue;
        var startInterIdx = augA[si].interIdx;
        if (visited[startInterIdx]) continue;
        
        var currentAug = augA;
        var currentClass = classA;
        var currentIdx = si;
        var startIdx = si;
        visited[startInterIdx] = true;
        
        // Process the starting intersection: if entry, switch polygons immediately
        var sw = processInter(currentAug, currentIdx, currentClass);
        if (sw) { currentAug = sw.aug; currentClass = sw.cls; currentIdx = sw.idx; }
        
        var loopLen = 0;
        var MAX_LOOP = augA.length + augB.length + 100;
        
        while (loopLen < MAX_LOOP) {
          loopLen++;
          
          var nextInter = augNextInter(currentAug, currentIdx);
          if (nextInter < 0) break;
          
          augAddRange(currentAug, currentIdx, nextInter, result);
          
          if (nextInter === startIdx && currentAug === augA) break;
          
          var interIdx = currentAug[nextInter].interIdx;
          var classification = currentClass[interIdx];
          
          if (classification === 'entry') {
            var otherAug = (currentAug === augA) ? augB : augA;
            var otherInterIdx = -1;
            for (var oi = 0; oi < otherAug.length; oi++) {
              if (otherAug[oi].isIntersection && otherAug[oi].interIdx === interIdx) {
                otherInterIdx = oi; break;
              }
            }
            if (otherInterIdx < 0) break;
            visited[interIdx] = true;
            currentAug = otherAug;
            currentClass = (currentAug === augA) ? classA : classB;
            currentIdx = otherInterIdx;
          } else {
            visited[interIdx] = true;
            currentIdx = nextInter;
          }
        }
        
        if (loopLen > 1 && result.length >= 3) break;
      }

      if (result.length < 3) return null;

      // Clean result: remove consecutive duplicate points
      var cleaned = [];
      for (var i = 0; i < result.length; i++) {
        var prev = i === 0 ? result[result.length - 1] : result[i - 1];
        if (Math.abs(result[i].x - prev.x) > eps || Math.abs(result[i].y - prev.y) > eps) {
          cleaned.push({ x: result[i].x, y: result[i].y });
        }
      }
      // Ensure minimum 3 unique vertices
      if (cleaned.length < 3) return null;

      // Run poly area check: union area should be >= max of the two input areas
      // (Allow 5% tolerance for numerical issues)
      var unionArea = Math.abs(_signedArea(cleaned) / 2);
      var aArea = Math.abs(_signedArea(polyA) / 2);
      var bArea = Math.abs(_signedArea(polyB) / 2);
      if (unionArea < Math.max(aArea, bArea) * 0.95) return null; // union too small, likely winding issue

      return cleaned;
    },

    // ── Convex hull (monotone chain) ────────────────────────
    convexHull: function(points) {
      var ps = points.slice().sort(function(a, b) { return a.x !== b.x ? a.x - b.x : a.y - b.y; });
      if (ps.length <= 1) return ps.slice();
      function cross2(o, a, b) { return (a.x - o.x) * (b.y - o.y) - (a.y - o.y) * (b.x - o.x); }
      var lower = [];
      for (var i = 0; i < ps.length; i++) {
        while (lower.length >= 2 && cross2(lower[lower.length - 2], lower[lower.length - 1], ps[i]) <= 0) lower.pop();
        lower.push(ps[i]);
      }
      var upper = [];
      for (var i = ps.length - 1; i >= 0; i--) {
        while (upper.length >= 2 && cross2(upper[upper.length - 2], upper[upper.length - 1], ps[i]) <= 0) upper.pop();
        upper.push(ps[i]);
      }
      lower.pop(); upper.pop();
      return lower.concat(upper);
    },

    // ── Compute offset polygon for Auto role ────────────────
    // For single parts: offset the actual polygon.
    // For groups (composite _sP): compute convex hull of all children then offset hull.
    computeAutoOffset: function(part, d) {
      if (part._sP && Array.isArray(part._sP) && part._sP.length > 1) {
        // Compute individual offsets for each sub-part
        var childOffsets = part._sP.filter(function(sp) { return sp && sp.length >= 3; }).map(function(sp) { return MarkerNesting.offsetPolygon(sp, d); });
        if (childOffsets.length === 0) { return part.points ? part.points.slice() : []; }
        // Try pairwise polygon union to preserve concavities
        var unionResult = childOffsets[0];
        for (var ui = 1; ui < childOffsets.length; ui++) {
          var next = MarkerNesting.polygonUnion(unionResult, childOffsets[ui]);
          if (next) unionResult = next;
          else {
            // Union failed (disjoint sub-parts): fall back to convex hull
            var allVerts = [];
            part._sP.forEach(function(sp) { sp.forEach(function(v) { allVerts.push({ x: v.x, y: v.y }); }); });
            var hull = MarkerNesting.convexHull(allVerts);
            unionResult = MarkerNesting.offsetPolygon(hull, d);
            break;
          }
        }
        // Offset the merged result (if not already offset)
        var result = unionResult;
        // Store individual sub-part offsets for visual rendering
        result._childOffsets = childOffsets;
        return result;
      }
      if (part.points && part.points.length >= 3) {
        return MarkerNesting.offsetPolygon(part.points, d);
      }
      return part.points ? part.points.slice() : [];
    },

    createCompositeParts: function(parts) {
      var hasGroup = parts.some(function(p) { return p._groupId && p._groupId !== ''; });
      if (!hasGroup) return parts;
      var bySource = {};
      parts.forEach(function(p) {
        var key = p.sourceResultIndex + '__' + (p._groupId || '');
        if (!bySource[key]) bySource[key] = [];
        bySource[key].push(p);
      });
      var result = [];
      Object.keys(bySource).forEach(function(key) {
        var sourceParts = bySource[key];
        var gId = sourceParts[0]._groupId || '';
        var grouped = sourceParts.filter(function(p) { return p._groupId === gId && gId !== ''; });
        var ungrouped = sourceParts.filter(function(p) { return p._groupId !== gId || gId === ''; });
        if (grouped.length >= 2) {
          var _sP = grouped.map(function(p) { return p.points.map(function(v) { return { x: v.x, y: v.y }; }); });
          var bb = { minX: Infinity, maxX: -Infinity, minY: Infinity, maxY: -Infinity };
          _sP.forEach(function(sp) { sp.forEach(function(v) { if (v.x < bb.minX) bb.minX = v.x; if (v.x > bb.maxX) bb.maxX = v.x; if (v.y < bb.minY) bb.minY = v.y; if (v.y > bb.maxY) bb.maxY = v.y; }); });
          var comp = {
            id: grouped[0].id,
            sourceId: grouped[0].id,
            sourceFileName: grouped[0].sourceFileName,
            sourceResultIndex: parseInt(key.split('__')[0]),
            origColor: grouped[0].origColor,
            area: grouped.reduce(function(s, p) { return s + p.area; }, 0),
            width: bb.maxX - bb.minX,
            height: bb.maxY - bb.minY,
            boundingBox: bb,
            points: grouped[0].points.map(function(v) { return { x: v.x, y: v.y }; }),
            _sP: _sP,
            _pairFlag: grouped.some(function(p) { return p._pairFlag; }),
            _groupId: gId
          };
          result.push(comp);
        } else {
          grouped.forEach(function(p) { result.push(p); });
        }
        ungrouped.forEach(function(p) { result.push(p); });
      });
      return result;
    },

    expandCompositePlacements: function(placements) {
      var result = [];
      placements.forEach(function(p) {
        if (p._unionOffset) { result.push(p); return; }
        if (p._csIndividual) { result.push(p); return; } // already individual sub-part
        var pts = p.points;
        var nestedSP = pts && pts._sP && Array.isArray(pts._sP);
        var topSP = !nestedSP && p._sP && Array.isArray(p._sP);
        var spData = nestedSP ? pts._sP : (topSP ? p._sP : null);
        if (spData) {
          spData.forEach(function(sp, spi) {
            var expanded = {
              id: p.id ? (typeof p.id === 'number' ? p.id + spi / 100 : p.id + '_' + spi) : spi + 1,
              sourceId: p.sourceId,
              sourceFileName: p.sourceFileName,
              origColor: p.origColor,
              rotation: p.rotation || 0,
              _m: p._m || false,
              points: sp.map(function(v) { return { x: v.x, y: v.y }; })
            };
            // Sub-parts were never offset; _originalPoints = points for fill rendering
            expanded._originalPoints = expanded.points.map(function(v) { return { x: v.x, y: v.y }; });
            result.push(expanded);
          });
        } else {
          result.push(p);
        }
      });
      return result;
    },

    createPartsWithCopies: function(parts, count, options) {
      options = options || {};
      var skipMirrors = !!options.skipMirrors;
      const result = [];
      let id = 1;
      parts.forEach(part => {
        for (let i = 0; i < count; i++) {
          var copy = {
            id: id++,
            sourceId: part.id,
            sourceFileName: part.sourceFileName,
            area: part.area,
            width: part.width,
            height: part.height,
            boundingBox: part.boundingBox,
            origColor: part.origColor,
            _pairMirror: false,
            _groupId: part._groupId || ''
          };
          if (part._sP) {
            copy._sP = part._sP.map(function(sp) { return sp.map(function(v) { return { x: v.x, y: v.y }; }); });
            copy.points = [];
          } else {
            copy.points = part.points.map(function(p) { return { x: p.x, y: p.y }; });
          }
          result.push(copy);
          if (part._pairFlag && !skipMirrors) {
            const cx = (part.boundingBox.minX + part.boundingBox.maxX) / 2;
            var mirCopy = {
              id: id++,
              sourceId: part.id,
              sourceFileName: part.sourceFileName,
              area: part.area,
              origColor: part.origColor,
              _pairMirror: true,
              _groupId: part._groupId || ''
            };
            if (part._sP) {
              var mirSP = part._sP.map(function(sp) { return sp.map(function(p) { return { x: -(p.x - cx) + cx, y: p.y }; }); });
              var mirBb = { minX: Infinity, maxX: -Infinity, minY: Infinity, maxY: -Infinity };
              mirSP.forEach(function(sp) { sp.forEach(function(v) { if (v.x < mirBb.minX) mirBb.minX = v.x; if (v.x > mirBb.maxX) mirBb.maxX = v.x; if (v.y < mirBb.minY) mirBb.minY = v.y; if (v.y > mirBb.maxY) mirBb.maxY = v.y; }); });
              mirCopy._sP = mirSP;
              mirCopy.points = [];
              mirCopy.boundingBox = mirBb;
              mirCopy.width = mirBb.maxX - mirBb.minX;
              mirCopy.height = mirBb.maxY - mirBb.minY;
            } else {
              var mirPts = part.points.map(function(p) { return { x: -(p.x - cx) + cx, y: p.y }; });
              var mirBb2 = { minX: -(part.boundingBox.maxX - cx) + cx, maxX: -(part.boundingBox.minX - cx) + cx, minY: part.boundingBox.minY, maxY: part.boundingBox.maxY };
              mirCopy.points = mirPts;
              mirCopy.boundingBox = mirBb2;
              mirCopy.width = mirBb2.maxX - mirBb2.minX;
              mirCopy.height = mirBb2.maxY - mirBb2.minY;
            }
            result.push(mirCopy);
          }
        }
      });
      return result;
    },

    computeBounds: function(placements) {
      if (!placements || placements.length === 0) return { minX: 0, minY: 0, maxX: 0, maxY: 0 };
      let minX = Infinity, minY = Infinity, maxX = -Infinity, maxY = -Infinity;
      placements.forEach(p => {
        forEachPoint(p.points, function(pt) {
          if (pt.x < minX) minX = pt.x;
          if (pt.y < minY) minY = pt.y;
          if (pt.x > maxX) maxX = pt.x;
          if (pt.y > maxY) maxY = pt.y;
        });
      });
      return { minX, minY, maxX, maxY };
    },

    computeMarkerLength: function(placements) {
      const bounds = MarkerNesting.computeBounds(placements);
      return Math.max(0, bounds.maxY);
    },

    unitsToMm: getUnitToMm,
    forEachPoint: forEachPoint,
    getShapeBB: getShapeBB,
    shiftShape: shiftShape,
    cloneShape: cloneShape,

    _getPlacementBB: function(pl) {
      var mx = Infinity, nx = -Infinity, my = Infinity, ny = -Infinity;
      var pts = pl.points;
      if (!pts) return { minX: 0, maxX: 0, minY: 0, maxY: 0 };
      if (pts._sP) {
        pts._sP.forEach(function(sp) {
          if (!sp) return;
          sp.forEach(function(v) {
            if (v.x < mx) mx = v.x; if (v.x > nx) nx = v.x;
            if (v.y < my) my = v.y; if (v.y > ny) ny = v.y;
          });
        });
      } else if (Array.isArray(pts)) {
        pts.forEach(function(v) {
          if (v.x < mx) mx = v.x; if (v.x > nx) nx = v.x;
          if (v.y < my) my = v.y; if (v.y > ny) ny = v.y;
        });
      }
      return { minX: mx, maxX: nx, minY: my, maxY: ny };
    },

    _shiftPlacementPoints: function(pl, dx, dy) {
      if (pl.points && pl.points._sP) {
        pl.points._sP = pl.points._sP.map(function(sp) {
          return sp.map(function(v) { return { x: v.x + dx, y: v.y + dy }; });
        });
      } else if (Array.isArray(pl.points)) {
        pl.points = pl.points.map(function(v) { return { x: v.x + dx, y: v.y + dy }; });
      }
      if (pl._sP && Array.isArray(pl._sP)) {
        pl._sP = pl._sP.map(function(sp) {
          return sp.map(function(v) { return { x: v.x + dx, y: v.y + dy }; });
        });
      }
      if (pl._originalPoints && Array.isArray(pl._originalPoints)) {
        pl._originalPoints = pl._originalPoints.map(function(v) { return { x: v.x + dx, y: v.y + dy }; });
      }
    },

    _compactDown: function(placements, binWidth, bnd) {
      if (placements.length < 2) return;
      var bbs = placements.map(function(pl) { return MarkerNesting._getPlacementBB(pl); });
      var targetBnd = typeof bnd === 'number' ? bnd : 0;
      var idxs = bbs.map(function(_, i) { return i; }).sort(function(a, b) { return bbs[a].minY - bbs[b].minY; });
      for (var oi = 1; oi < idxs.length; oi++) {
        var i = idxs[oi], bi = bbs[i], targetY = targetBnd;
        for (var j = 0; j < placements.length; j++) {
          if (j === i) continue;
          var bj = bbs[j];
          if (bi.maxX > bj.minX && bi.minX < bj.maxX && bj.maxY > targetY) targetY = bj.maxY;
        }
        var shift = targetY - bi.minY;
        if (shift < -0.1) {
          MarkerNesting._shiftPlacementPoints(placements[i], 0, shift);
          bbs[i].minY += shift;
          bbs[i].maxY += shift;
        }
      }
    },

    _compactLeft: function(placements, binWidth, inset) {
      if (placements.length < 2) return;
      var bnd = typeof inset === 'number' ? inset : 0;
      var bbs = placements.map(function(pl) { return MarkerNesting._getPlacementBB(pl); });
      for (var pass = 0; pass < 3; pass++) {
        var moved = false;
        var idxsByY = bbs.map(function(_, i) { return i; }).sort(function(a, b) { return bbs[a].minY - bbs[b].minY; });
        // Leftward binary search
        for (var oi = 1; oi < idxsByY.length; oi++) {
          var i = idxsByY[oi], bi = bbs[i], lo = bnd, hi = bi.minX;
          while (hi - lo > 0.5) {
            var md = (lo + hi) / 2;
            var shift = md - bi.minX;
            var ok = true;
            for (var j = 0; j < placements.length; j++) {
              if (j === i) continue;
              var bj = bbs[j];
              if (bi.maxX + shift > bj.minX && bi.minX + shift < bj.maxX && bi.maxY > bj.minY && bi.minY < bj.maxY) {
                ok = false; break;
              }
            }
            if (ok) hi = md; else lo = md;
          }
          var finalShift = hi - bi.minX;
          if (finalShift < -0.1) {
            MarkerNesting._shiftPlacementPoints(placements[i], finalShift, 0);
            bbs[i].minX += finalShift; bbs[i].maxX += finalShift;
            moved = true;
          }
        }
        // Downward binary search
        for (var oi = 1; oi < idxsByY.length; oi++) {
          var i = idxsByY[oi], bi = bbs[i];
          var lo = bnd, hi = bi.minY;
          while (hi - lo > 0.5) {
            var md = (lo + hi) / 2;
            var shift = md - bi.minY;
            var ok = true;
            for (var j = 0; j < placements.length; j++) {
              if (j === i) continue;
              var bj = bbs[j];
              if (bi.maxX > bj.minX && bi.minX < bj.maxX && bi.maxY + shift > bj.minY && bi.minY + shift < bj.maxY) {
                ok = false; break;
              }
            }
            if (ok) hi = md; else lo = md;
          }
          var finalShift = hi - bi.minY;
          if (finalShift < -0.1) {
            MarkerNesting._shiftPlacementPoints(placements[i], 0, finalShift);
            bbs[i].minY += finalShift; bbs[i].maxY += finalShift;
            moved = true;
          }
        }
        if (!moved) break;
      }
    },

    _samplePoly: function(pts, maxN) {
      if (!pts || pts.length <= maxN) return pts;
      var step = (pts.length - 1) / (maxN - 1);
      var out = [pts[0]];
      for (var i = 1; i < maxN - 1; i++) out.push(pts[Math.round(i * step)]);
      out.push(pts[pts.length - 1]);
      return out;
    },

    _polyOverlap: function(a, b) {
      if (!a || !b || a.length < 3 || b.length < 3) return false;
      var aMinX = 1/0, aMaxX = -1/0, aMinY = 1/0, aMaxY = -1/0;
      for (var i = 0; i < a.length; i++) { if (a[i].x < aMinX) aMinX = a[i].x; if (a[i].x > aMaxX) aMaxX = a[i].x; if (a[i].y < aMinY) aMinY = a[i].y; if (a[i].y > aMaxY) aMaxY = a[i].y; }
      var bMinX = 1/0, bMaxX = -1/0, bMinY = 1/0, bMaxY = -1/0;
      for (var i = 0; i < b.length; i++) { if (b[i].x < bMinX) bMinX = b[i].x; if (b[i].x > bMaxX) bMaxX = b[i].x; if (b[i].y < bMinY) bMinY = b[i].y; if (b[i].y > bMaxY) bMaxY = b[i].y; }
      if (aMaxX < bMinX || bMaxX < aMinX || aMaxY < bMinY || bMaxY < aMinY) return false;
      function _ptIn(x, y, p) {
        var inside = false;
        for (var i = 0, j = p.length - 1; i < p.length; j = i++) {
          var yi = p[i].y, yj = p[j].y, xi = p[i].x, xj = p[j].x;
          if ((yi > y) !== (yj > y) && x < (xj - xi) * (y - yi) / (yj - yi) + xi) inside = !inside;
        }
        return inside;
      }
      for (var i = 0; i < a.length; i++) if (_ptIn(a[i].x, a[i].y, b)) return true;
      for (var i = 0; i < b.length; i++) if (_ptIn(b[i].x, b[i].y, a)) return true;
      function _cr(a, b, c) { return (b.x - a.x) * (c.y - a.y) - (b.y - a.y) * (c.x - a.x); }
      function _segX(a, b, c, d) {
        var o1 = _cr(a, b, c), o2 = _cr(a, b, d), o3 = _cr(c, d, a), o4 = _cr(c, d, b);
        if (((o1 > 0) !== (o2 > 0)) && ((o3 > 0) !== (o4 > 0))) return true;
        return false;
      }
      for (var i = 0, j = a.length - 1; i < a.length; j = i++) {
        for (var k = 0, l = b.length - 1; k < b.length; l = k++) {
          if (_segX(a[j], a[i], b[l], b[k])) return true;
        }
      }
      return false;
    },

    _plOverlap: function(plA, plB) {
      var subsA, subsB;
      if (plA._unionOffset && Array.isArray(plA.points) && plA.points.length >= 3) {
        subsA = [plA.points];
      } else if (plA._sP && Array.isArray(plA._sP)) {
        subsA = plA._sP;
      } else if (Array.isArray(plA.points)) {
        subsA = [plA.points];
      } else {
        subsA = [];
      }
      if (plB._unionOffset && Array.isArray(plB.points) && plB.points.length >= 3) {
        subsB = [plB.points];
      } else if (plB._sP && Array.isArray(plB._sP)) {
        subsB = plB._sP;
      } else if (Array.isArray(plB.points)) {
        subsB = [plB.points];
      } else {
        subsB = [];
      }
      for (var i = 0; i < subsA.length; i++) {
        if (!subsA[i] || subsA[i].length < 3) continue;
        for (var j = 0; j < subsB.length; j++) {
          if (!subsB[j] || subsB[j].length < 3) continue;
          if (MarkerNesting._polyOverlap(subsA[i], subsB[j])) return true;
        }
      }
      return false;
    },

    _compactPoly: function(placements, binWidth, inset) {
      if (placements.length < 2) return;
      var bnd = typeof inset === 'number' ? inset : 0;
      var SAMPLE_MAX = 40;
      function getSampledSubs(pl) {
        var subs;
        if (pl._unionOffset && Array.isArray(pl.points) && pl.points.length >= 3) {
          subs = [pl.points];
        } else if (pl._sP && Array.isArray(pl._sP)) {
          subs = pl._sP;
        } else if (Array.isArray(pl.points)) {
          subs = [pl.points];
        } else {
          subs = [];
        }
        return subs.map(function(s) { return MarkerNesting._samplePoly(s, SAMPLE_MAX); });
      }
      function subsBB(subs) {
        var mx = 1/0, nx = -1/0, my = 1/0, ny = -1/0;
        for (var si = 0; si < subs.length; si++) {
          if (!subs[si]) continue;
          for (var vi = 0; vi < subs[si].length; vi++) {
            var v = subs[si][vi];
            if (v.x < mx) mx = v.x; if (v.x > nx) nx = v.x;
            if (v.y < my) my = v.y; if (v.y > ny) ny = v.y;
          }
        }
        return { minX: mx, maxX: nx, minY: my, maxY: ny };
      }
      function shiftSubs(subs, dx, dy) {
        return subs.map(function(s) {
          if (!s) return s;
          return s.map(function(v) { return { x: v.x + dx, y: v.y + dy }; });
        });
      }
      function subsOverlap(subsA, subsB) {
        for (var i = 0; i < subsA.length; i++) {
          if (!subsA[i] || subsA[i].length < 3) continue;
          for (var j = 0; j < subsB.length; j++) {
            if (!subsB[j] || subsB[j].length < 3) continue;
            if (MarkerNesting._polyOverlap(subsA[i], subsB[j])) return true;
          }
        }
        return false;
      }
      var sampled = placements.map(function(pl) { return getSampledSubs(pl); });
      for (var pass = 0; pass < 2; pass++) {
        var moved = false;
        var bbs = sampled.map(function(s) { return subsBB(s); });
        var idxsByY = bbs.map(function(_, i) { return i; }).sort(function(a, b) { return bbs[a].minY - bbs[b].minY; });
        for (var oi = 1; oi < idxsByY.length; oi++) {
          var i = idxsByY[oi], bi = bbs[i];
          var lo = bnd, hi = bi.minX;
          if (hi - lo < 1.1) continue;
          while (hi - lo > 1.0) {
            var md = (lo + hi) / 2;
            var shift = md - bi.minX;
            var shifted = shiftSubs(sampled[i], shift, 0);
            var ok = true;
            for (var j = 0; j < placements.length; j++) {
              if (j === i) continue;
              if (subsOverlap(shifted, sampled[j])) { ok = false; break; }
            }
            if (ok) hi = md; else lo = md;
          }
          var finalShift = hi - bi.minX;
          if (finalShift < -0.1) {
            MarkerNesting._shiftPlacementPoints(placements[i], finalShift, 0);
            sampled[i] = shiftSubs(sampled[i], finalShift, 0);
            bbs[i] = subsBB(sampled[i]);
            moved = true;
          }
        }
        bbs = sampled.map(function(s) { return subsBB(s); });
        idxsByY = bbs.map(function(_, i) { return i; }).sort(function(a, b) { return bbs[a].minY - bbs[b].minY; });
        for (var oi = 1; oi < idxsByY.length; oi++) {
          var i = idxsByY[oi], bi = bbs[i];
          var lo = bnd, hi = bi.minY;
          if (hi - lo < 1.1) continue;
          while (hi - lo > 1.0) {
            var md = (lo + hi) / 2;
            var shift = md - bi.minY;
            var shifted = shiftSubs(sampled[i], 0, shift);
            var ok = true;
            for (var j = 0; j < placements.length; j++) {
              if (j === i) continue;
              if (subsOverlap(shifted, sampled[j])) { ok = false; break; }
            }
            if (ok) hi = md; else lo = md;
          }
          var finalShift = hi - bi.minY;
          if (finalShift < -0.1) {
            MarkerNesting._shiftPlacementPoints(placements[i], 0, finalShift);
            sampled[i] = shiftSubs(sampled[i], 0, finalShift);
            bbs[i] = subsBB(sampled[i]);
            moved = true;
          }
        }
        if (!moved) break;
      }
    },

    /* ── MaxRects packer with compaction ── */

    rectanglePack: function(parts, binWidthMm) {
      if (!parts || parts.length === 0) return { placements: [], unplaced: [], markerLength: 0 };

      var freeRects = [{ x: 0, y: 0, w: binWidthMm, h: 1e8 }];
      var placements = [];
      var sorted = parts.slice().sort(function(a, b) { return b.area - a.area; });

      function getBB(part) {
        var bb = part.boundingBox;
        if (bb) return bb;
        var b = getShapeBB(part.points || part);
        return b || { minX: 0, maxX: 0, minY: 0, maxY: 0, width: 0, height: 0 };
      }

      function rotatePts(pts, bb, rot) {
        var cx = (bb.minX + bb.maxX) / 2, cy = (bb.minY + bb.maxY) / 2;
        if (rot === '180') return pts.map(function(p) { return { x: -(p.x - cx) + cx, y: -(p.y - cy) + cy }; });
        if (rot === '270') return pts.map(function(p) { return { x: (p.y - cy) + cx, y: -(p.x - cx) + cy }; });
        return pts.map(function(p) { return { x: -(p.y - cy) + cx, y: (p.x - cx) + cy }; });
      }

      function findBL(w, h) {
        var best = null;
        for (var i = 0; i < freeRects.length; i++) {
          var r = freeRects[i];
          if (w <= r.w && h <= r.h) {
            if (!best || r.y < best.r.y || (r.y === best.r.y && r.x < best.r.x)) {
              best = { idx: i, r: r };
            }
          }
        }
        return best;
      }

      function splitAndPrune(rect, usedW, usedH) {
        var news = [];
        if (rect.w - usedW > 0) news.push({ x: rect.x + usedW, y: rect.y, w: rect.w - usedW, h: usedH });
        if (rect.h - usedH > 0) news.push({ x: rect.x, y: rect.y + usedH, w: rect.w, h: rect.h - usedH });
        news.forEach(function(n) { freeRects.push(n); });
        for (var i = 0; i < freeRects.length; i++) {
          for (var j = i + 1; j < freeRects.length; j++) {
            var a = freeRects[i], b = freeRects[j];
            if (a.x >= b.x && a.y >= b.y && a.x + a.w <= b.x + b.w && a.y + a.h <= b.y + b.h) {
              freeRects.splice(i--, 1); break;
            }
            if (b.x >= a.x && b.y >= a.y && b.x + b.w <= a.x + a.w && b.y + b.h <= a.y + a.h) {
              freeRects.splice(j--, 1);
            }
          }
        }
        freeRects.sort(function(a, b) { return a.y - b.y || a.x - b.x; });
      }

      sorted.forEach(function(part) {
        var bb = getBB(part);
        var w = bb.maxX - bb.minX, h = bb.maxY - bb.minY;

        function bestFit(w, h, rotated) {
          var f = findBL(w, h);
          return f ? { fit: f, pw: w, ph: h, rot: rotated } : null;
        }

        var opts = [bestFit(w, h, false)];
        if (w !== h) opts.push(bestFit(h, w, true));
        if (w !== h) opts.push(bestFit(w, h, '180'));
        if (w !== h) opts.push(bestFit(h, w, '270'));
        opts = opts.filter(Boolean);

        if (opts.length === 0) {
          var maxY = 0;
          freeRects.forEach(function(r) { var b = r.y + r.h; if (b > maxY) maxY = b; });
          freeRects.push({ x: 0, y: maxY, w: binWidthMm, h: 1e8 });
          opts = [bestFit(w, h, false)];
          if (w !== h) opts.push(bestFit(h, w, true));
          if (w !== h) opts.push(bestFit(w, h, '180'));
          if (w !== h) opts.push(bestFit(h, w, '270'));
          opts = opts.filter(Boolean);
        }

        if (opts.length > 0) {
          opts.sort(function(a, b) { return a.fit.r.y - b.fit.r.y || a.fit.r.x - b.fit.r.x; });
          var best = opts[0];
          var rect = best.fit.r;
          var rotKey = best.rot;
          var pts = rotKey ? rotatePts(part.points, bb, rotKey === true ? '90' : rotKey) : part.points;
          var minX = Infinity, minY = Infinity;
          pts.forEach(function(p) { if (p.x < minX) minX = p.x; if (p.y < minY) minY = p.y; });
          var translated = pts.map(function(p) { return { x: p.x - minX + rect.x, y: p.y - minY + rect.y }; });
          var rotAngle = rotKey === true ? 90 : (rotKey || 0);
          placements.push({
            id: part.id, sourceId: part.sourceId || part.id,
            points: translated, position: { x: rect.x, y: rect.y },
            rotation: rotAngle, width: best.pw, height: best.ph
          });
          freeRects.splice(best.fit.idx, 1);
          splitAndPrune(rect, best.pw, best.ph);
        }
      });

      var markerLen = 0;
      placements.forEach(function(p) {
        forEachPoint(p.points, function(pt) { if (pt.y > markerLen) markerLen = pt.y; });
      });

      return { placements: placements, unplaced: [], markerLengthMm: markerLen };
    },

    /* ── Multi-pass optimizer (Web Worker) ── */

    _workerBlobUrl: null,

    _getWorkerCode: function() {
      return [
        'function pointInPolygon(x,y,p){var i=0,j=p.length-1,xi,yi,xj,yj,inside=0;for(;i<p.length;j=i++){xi=p[i].x;yi=p[i].y;xj=p[j].x;yj=p[j].y;if((yi>y)!==(yj>y)&&x<(xj-xi)*(y-yi)/(yj-yi)+xi)inside=!inside}return inside}',
        'function _c(a,b,c){return(b.x-a.x)*(c.y-a.y)-(b.y-a.y)*(c.x-a.x)}',
        'function _oS(a,b,p){return p.x>=Math.min(a.x,b.x)&&p.x<=Math.max(a.x,b.x)&&p.y>=Math.min(a.y,b.y)&&p.y<=Math.max(a.y,b.y)}',
        'function _sX(a,b,c,d){var o1=_c(a,b,c),o2=_c(a,b,d),o3=_c(c,d,a),o4=_c(c,d,b);if((o1>0)!==(o2>0)&&(o3>0)!==(o4>0))return 1;if(o1===0&&_oS(a,b,c))return 1;if(o2===0&&_oS(a,b,d))return 1;if(o3===0&&_oS(c,d,a))return 1;if(o4===0&&_oS(c,d,b))return 1;return 0}',
        'function intersect(a,b,e){if(!a||!b||a.length<3||b.length<3)return 0;for(var i=0;i<a.length;i++){if(pointInPolygon(a[i].x,a[i].y,b))return 1}for(var i=0;i<b.length;i++){if(pointInPolygon(b[i].x,b[i].y,a))return 1}if(e!==0){for(var i=0,j=a.length-1;i<a.length;j=i++){for(var k=0,l=b.length-1;k<b.length;l=k++){if(_sX(a[j],a[i],b[l],b[k]))return 1}}}return 0}',
        'function _cOvl(a,b,e){if(!a||!b)return 0;var pa=a._sP||[a],pb=b._sP||[b];for(var i=0;i<pa.length;i++){if(typeof pa[i]==="object"&&!Array.isArray(pa[i]))continue;if(!pa[i]||!pa[i].length)continue;var mxa=1/0,mna=-mxa,mya=mxa,nna=-mxa;for(var _vi=0;_vi<pa[i].length;_vi++){var _v=pa[i][_vi];if(_v.x<mxa)mxa=_v.x;if(_v.x>mna)mna=_v.x;if(_v.y<mya)mya=_v.y;if(_v.y>nna)nna=_v.y}for(var j=0;j<pb.length;j++){if(typeof pb[j]==="object"&&!Array.isArray(pb[j]))continue;if(!pb[j]||!pb[j].length)continue;var mxb=1/0,mnb=-mxb,myb=mxb,nnb=-mxb;for(var _vj=0;_vj<pb[j].length;_vj++){var _w=pb[j][_vj];if(_w.x<mxb)mxb=_w.x;if(_w.x>mnb)mnb=_w.x;if(_w.y<myb)myb=_w.y;if(_w.y>nnb)nnb=_w.y}if(mxa>mnb||mna<mxb||mya>nnb||nna<myb)continue;if(intersect(pa[i],pb[j],e))return 1}}return 0}',
        'function getBB(p){if(p._sP){var mx=1/0,nx=-mx,my=mx,ny=-mx;for(var s=0;s<p._sP.length;s++){var sp=p._sP[s];for(var v=0;v<sp.length;v++){if(sp[v].x<mx)mx=sp[v].x;if(sp[v].x>nx)nx=sp[v].x;if(sp[v].y<my)my=sp[v].y;if(sp[v].y>ny)ny=sp[v].y}}return{minX:mx,maxX:nx,minY:my,maxY:ny,width:nx-mx,height:ny-my}}var I=1/0,i=-I,mx=I,my=I,n=i,y=i;(p.points||[]).forEach(function(v){if(v.x<mx)mx=v.x;if(v.x>n)n=v.x;if(v.y<my)my=v.y;if(v.y>y)y=v.y});return{minX:mx,maxX:n,minY:my,maxY:y,width:n-mx,height:y-my}}',
        'function rotPts(p,b,r){var cx=(b.minX+b.maxX)/2,cy=(b.minY+b.maxY)/2;if(r==="180")return p.map(function(v){return{x:-(v.x-cx)+cx,y:-(v.y-cy)+cy}});if(r==="270")return p.map(function(v){return{x:(v.y-cy)+cx,y:-(v.x-cx)+cy}});return p.map(function(v){return{x:-(v.y-cy)+cx,y:(v.x-cx)+cy}})}',
        'function pack(parts,w){var fr=[{x:0,y:0,w:w,h:1e8}],pl=[];function findBL(w,h){var b=null;for(var i=0;i<fr.length;i++){var r=fr[i];if(w<=r.w&&h<=r.h){if(!b||r.y<b.r.y||(r.y===b.r.y&&r.x<b.r.x))b={idx:i,r:r}}}return b}function split(r,uw,uh){if(r.w-uw>0)fr.push({x:r.x+uw,y:r.y,w:r.w-uw,h:uh});if(r.h-uh>0)fr.push({x:r.x,y:r.y+uh,w:r.w,h:r.h-uh});for(var i=0;i<fr.length;i++){for(var j=i+1;j<fr.length;j++){var a=fr[i],b=fr[j];if(a.x>=b.x&&a.y>=b.y&&a.x+a.w<=b.x+b.w&&a.y+a.h<=b.y+b.h){fr.splice(i--,1);break}if(b.x>=a.x&&b.y>=a.y&&b.x+b.w<=a.x+a.w&&b.y+b.h<=a.y+a.h){fr.splice(j--,1)}}}fr.sort(function(a,b){return a.y-b.y||a.x-b.x})}',
        'parts.slice().sort(function(a,b){return b.area-a.area}).forEach(function(part){var bb=getBB(part);var W=bb.width,H=bb.height;var opts=[(function(w,h){var f=findBL(w,h);return f?{fit:f,pw:w,ph:h,rot:0}:null})(W,H)];if(W!==H){opts.push((function(w,h){var f=findBL(w,h);return f?{fit:f,pw:w,ph:h,rot:1}:null})(H,W));opts.push((function(w,h){var f=findBL(w,h);return f?{fit:f,pw:w,ph:h,rot:"180"}:null})(W,H));opts.push((function(w,h){var f=findBL(w,h);return f?{fit:f,pw:w,ph:h,rot:"270"}:null})(H,W))}opts=opts.filter(Boolean);',
        'if(opts.length===0){var mY=0;fr.forEach(function(r){var b=r.y+r.h;if(b>mY)mY=b});fr.push({x:0,y:mY,w:w,h:1e8});opts=[(function(w,h){var f=findBL(w,h);return f?{fit:f,pw:w,ph:h,rot:0}:null})(W,H)];if(W!==H){opts.push((function(w,h){var f=findBL(w,h);return f?{fit:f,pw:w,ph:h,rot:1}:null})(H,W));opts.push((function(w,h){var f=findBL(w,h);return f?{fit:f,pw:w,ph:h,rot:"180"}:null})(W,H));opts.push((function(w,h){var f=findBL(w,h);return f?{fit:f,pw:w,ph:h,rot:"270"}:null})(H,W))}opts=opts.filter(Boolean)}',
        'if(opts.length>0){opts.sort(function(a,b){return a.fit.r.y-b.fit.r.y||a.fit.r.x-b.fit.r.x});var best=opts[0],rect=best.fit.r,rk=best.rot,pts=rk?rotPts(part.points,bb,rk===1?"90":rk):part.points,mnX=1/0,mnY=mnX;pts.forEach(function(p){if(p.x<mnX)mnX=p.x;if(p.y<mnY)mnY=p.y});pl.push({id:part.id,sourceId:part.sourceId||part.id,points:pts.map(function(p){return{x:p.x-mnX+rect.x,y:p.y-mnY+rect.y}}),position:{x:rect.x,y:rect.y},rotation:rk===1?90:(rk||0),width:best.pw,height:best.ph});fr.splice(best.fit.idx,1);split(rect,best.pw,best.ph)}});var len=0;pl.forEach(function(p){var _pts=p.points;if(_pts&&_pts._sP){_pts._sP.forEach(function(sp){sp.forEach(function(pt){if(pt.y>len)len=pt.y})})}else{(_pts||[]).forEach(function(pt){if(pt.y>len)len=pt.y})}});return{placements:pl,markerLengthMm:len}}',
        'function compact(p,w){if(p.length<2)return;var bbs=p.map(function(pl){var mx=1/0,nx=-mx,my=mx,ny=-mx;(pl.points||[]).forEach(function(v){if(v.x<mx)mx=v.x;if(v.x>nx)nx=v.x;if(v.y<my)my=v.y;if(v.y>ny)ny=v.y});return{minX:mx,maxX:nx,minY:my,maxY:ny}});for(var pass=0;pass<3;pass++){var moved=false;var idxs=bbs.map(function(_,i){return i}).sort(function(a,b){return bbs[a].minY-bbs[b].minY});for(var oi=1;oi<idxs.length;oi++){var i=idxs[oi],bi=bbs[i],lo=0,hi=bi.minX;while(hi-lo>0.5){var md=(lo+hi)/2,shift=md-bi.minX,ok=true;for(var j=0;j<p.length;j++){if(j===i)continue;var bj=bbs[j];if(bi.maxX+shift>bj.minX&&bi.minX+shift<bj.maxX&&bi.maxY>bj.minY&&bi.minY<bj.maxY){ok=false;break}}if(ok)hi=md;else lo=md}var shift=hi-bi.minX;if(shift<-0.1){p[i].points=p[i].points.map(function(v){return{x:v.x+shift,y:v.y}});bbs[i].minX+=shift;bbs[i].maxX+=shift;moved=true}}for(var oi=1;oi<idxs.length;oi++){var i=idxs[oi],bi=bbs[i],lo=0,hi=bi.minY;while(hi-lo>0.5){var md=(lo+hi)/2,shift=md-bi.minY,ok=true;for(var j=0;j<p.length;j++){if(j===i)continue;var bj=bbs[j];if(bi.maxX>bj.minX&&bi.minX<bj.maxX&&bi.maxY+shift>bj.minY&&bi.minY+shift<bj.maxY){ok=false;break}}if(ok)hi=md;else lo=md}var shift=hi-bi.minY;if(shift<-0.1){p[i].points=p[i].points.map(function(v){return{x:v.x,y:v.y+shift}});bbs[i].minY+=shift;bbs[i].maxY+=shift;moved=true}}if(!moved)break}}',
        'function runShuffle(parts,bw,it){var best=null,bestLen=1/0;for(var ind=0;ind<it;ind++){var order=parts.slice();if(ind>0){for(var j=order.length-1;j>0;j--){var k=Math.random()*(j+1)|0,t=order[j];order[j]=order[k];order[k]=t}}var r=pack(order,bw);if(r.placements.length>1)compact(r.placements,bw);if(r.markerLengthMm<bestLen){bestLen=r.markerLengthMm;best=r}self.postMessage({type:"progress",pass:ind+1,total:it,bestLen:bestLen,placements:best?best.placements:[]})}self.postMessage({type:"result",placements:best?best.placements:[],markerLengthMm:best?best.markerLengthMm:0})}',
        'function rotPtsArb(pts,a){if(a===0){if(pts._sP)return{_sP:pts._sP.map(function(sp){return sp.map(function(v){return{x:v.x,y:v.y}})})};return pts.map(function(v){return{x:v.x,y:v.y}})}if(pts._sP){var bb=getBB(pts),cx=(bb.minX+bb.maxX)/2,cy=(bb.minY+bb.maxY)/2,rad=a*Math.PI/180,cos=Math.cos(rad),sin=Math.sin(rad);return{_sP:pts._sP.map(function(sp){return sp.map(function(v){var dx=v.x-cx,dy=v.y-cy;return{x:dx*cos-dy*sin+cx,y:dx*sin+dy*cos+cy}})})}}var bb=getBB({points:pts}),cx=(bb.minX+bb.maxX)/2,cy=(bb.minY+bb.maxY)/2,rad=a*Math.PI/180,cos=Math.cos(rad),sin=Math.sin(rad);return pts.map(function(v){var dx=v.x-cx,dy=v.y-cy;return{x:dx*cos-dy*sin+cx,y:dx*sin+dy*cos+cy}})}',
        'function sg(pts,n){if(pts._sP)return{_sP:pts._sP.map(function(sp){return sg(sp,n)})};if(pts.length<=n||n<3)return pts;var step=(pts.length-1)/(n-1),out=[pts[0]];for(var i=1;i<n-1;i++)out.push(pts[Math.round(i*step)]);out.push(pts[pts.length-1]);return out}',
        'function pairMotif(pts,mir,bb,mbb,bw,n,cpt){var mv=[{pts:pts,bb:bb}];if(bb.maxX-bb.minX<=bw){var pts180=rotPtsArb(pts,180),bb180=getBB(pts180._sP?pts180:{points:pts180});if(bb180.maxX-bb180.minX<=bw)mv.push({pts:pts180,bb:bb180})}var mr=buildMotif(pts,bb,mv,bw,n,cpt);if(!mr)return null;if(mr.placements.length!==n)return null;var len=mr.markerLengthMm,ow=bb.maxX-bb.minX,mw=mbb.maxX-mbb.minX,oh=bb.maxY-bb.minY;var ex1=[];for(var pi=0;pi<mr.placements.length;pi++){var p=mr.placements[pi],pp=p.points;ex1.push({id:p.id||(pi+1),sourceId:p.sourceId||p.id||(pi+1),points:pp,rotation:0,_m:false});var pbb=getBB(pp._sP?pp:{points:pp}),pcx=(pbb.minX+pbb.maxX)/2;var mp=pp._sP?{_sP:pp._sP.map(function(sp){return sp.map(function(v){return{x:-(v.x-pcx)+pcx,y:v.y+len}})})}:pp.map(function(v){return{x:-(v.x-pcx)+pcx,y:v.y+len}});ex1.push({id:p.id||(pi+1),sourceId:p.sourceId||p.id||(pi+1),points:mp,rotation:0,_m:true})}var mlen1=0;ex1.forEach(function(p){var ps=p.points;if(ps._sP)ps._sP.forEach(function(sp){sp.forEach(function(pt){if(pt.y>mlen1)mlen1=pt.y})});else(ps||[]).forEach(function(pt){if(pt.y>mlen1)mlen1=pt.y})});if(mw<=bw-ow-1){var spc=ow+mw+1,pp2=Math.max(1,Math.floor((bw-0.5)/spc)),rs2=Math.ceil(n/pp2),ex2=[];for(var ri=0;ri<rs2;ri++){for(var ci=0;ci<pp2&&(ri*pp2+ci)<n;ci++){var pi2=ri*pp2+ci,p2=mr.placements[pi2],pp2_=p2.points,p2bb=getBB(pp2_._sP?pp2_:{points:pp2_});var ox=ci*spc-p2bb.minX,oy=ri*oh-p2bb.minY;ex2.push({id:p2.id||(pi2+1),sourceId:p2.sourceId||p2.id||(pi2+1),points:pp2_._sP?{_sP:pp2_._sP.map(function(s){return s.map(function(v){return{x:v.x+ox,y:v.y+oy}})})}:pp2_.map(function(v){return{x:v.x+ox,y:v.y+oy}}),rotation:0,_m:false});var pcx2=(p2bb.minX+p2bb.maxX)/2;var mp2_=pp2_._sP?{_sP:pp2_._sP.map(function(s){return s.map(function(v){return{x:-(v.x-pcx2)+pcx2+spc,y:v.y+oy}})})}:pp2_.map(function(v){return{x:-(v.x-pcx2)+pcx2+spc,y:v.y+oy}});ex2.push({id:p2.id||(pi2+1),sourceId:p2.sourceId||p2.id||(pi2+1),points:mp2_,rotation:0,_m:true})}}var mlen2=0;ex2.forEach(function(p){var ps=p.points;if(ps._sP)ps._sP.forEach(function(sp){sp.forEach(function(pt){if(pt.y>mlen2)mlen2=pt.y})});else(ps||[]).forEach(function(pt){if(pt.y>mlen2)mlen2=pt.y})});if(ex1.length!==n*2||ex2.length!==n*2)return null;if(mlen2<mlen1)return{placements:ex2,markerLengthMm:mlen2};}if(ex1.length!==n*2)return null;return{placements:ex1,markerLengthMm:mlen1}}',
        'function buildMotif(pts,bb,vars,bw,n,cpt,pairMode){var ap={_sP:[]},pl=[{pts:pts,dx:0,dy:0,bb:bb,_mFlag:false}],sp0=sg(pts,300);var mx=bb.maxX,mn=bb.minX,my=bb.maxY,nn=bb.minY;if(sp0._sP)sp0._sP.forEach(function(s){ap._sP.push(s.map(function(v){return{x:v.x,y:v.y}}))});else ap._sP.push(sp0.map(function(v){return{x:v.x,y:v.y}}));var step=Math.max(1,Math.round((my-nn)/12)),bi=cpt?6:4;for(var k=1;k<4;k++){var best=null,bestH=1/0;for(var vi=0;vi<vars.length;vi++){var v=vars[vi];if(pairMode&&v._mFlag!==(k%2===1))continue;var vb=v.bb,vw=vb.maxX-vb.minX,vh=vb.maxY-vb.minY,spi=v.sp||sg(v.pts,300),dy,lo,hi,md,found,mp,nh,nw;for(dy=-(vh|0)-step*2;dy<=(my-nn|0)+step*2;dy+=step){lo=0;hi=bw-vw;found=false;for(var i=0;i<bi;i++){md=(lo+hi)/2;mp=spi._sP?{_sP:spi._sP.map(function(s){return s.map(function(p){return{x:p.x+md,y:p.y+dy}})})}:spi.map(function(p){return{x:p.x+md,y:p.y+dy}});if(_cOvl(mp,ap,0))lo=md;else{hi=md;found=true}}if(found){nh=Math.max(my,vb.maxY+dy)-Math.min(nn,vb.minY+dy);nw=Math.max(mx,vb.maxX+hi)-Math.min(mn,vb.minX+hi);if(nh<bestH&&nw<=bw+0.5){bestH=nh;best={pts:v.pts,bb:vb,dx:hi,dy:dy,sp:spi,_mFlag:v._mFlag}}}lo=-(bw-vw);hi=0;found=false;for(var i=0;i<bi;i++){md=(lo+hi)/2;mp=spi._sP?{_sP:spi._sP.map(function(s){return s.map(function(p){return{x:p.x+md,y:p.y+dy}})})}:spi.map(function(p){return{x:p.x+md,y:p.y+dy}});if(_cOvl(mp,ap,0))hi=md;else{lo=md;found=true}}if(found){nh=Math.max(my,vb.maxY+dy)-Math.min(nn,vb.minY+dy);nw=Math.max(mx,vb.maxX+lo)-Math.min(mn,vb.minX+lo);if(nh<bestH&&nw<=bw+0.5){bestH=nh;best={pts:v.pts,bb:vb,dx:lo,dy:dy,sp:spi,_mFlag:v._mFlag}}}}}if(!best)break;var bsp=best.sp._sP||[best.sp];for(var pi=0;pi<bsp.length;pi++)ap._sP.push(bsp[pi].map(function(p){return{x:p.x+best.dx,y:p.y+best.dy}}));if(best.dx+best.bb.maxX>mx)mx=best.dx+best.bb.maxX;if(best.dx+best.bb.minX<mn)mn=best.dx+best.bb.minX;if(best.dy+best.bb.maxY>my)my=best.dy+best.bb.maxY;if(best.dy+best.bb.minY<nn)nn=best.dy+best.bb.minY;pl.push(best)}var cw=mx-mn,ch=my-nn,pp=Math.max(1,Math.floor((bw-0.5)/cw)),mps=pl.length,rs=Math.ceil(Math.ceil(n/mps)/pp),pl2=[],plc=0;if(cw>bw+0.5)return null;for(var r=0;r<rs&&plc<n;r++){for(var c=0;c<pp&&plc<n;c++){for(var pi=0;pi<mps&&plc<n;pi++){var ox=c*cw+pl[pi].dx-mn,oy=r*ch+pl[pi].dy-nn,pp2=pl[pi].pts._sP?{_sP:pl[pi].pts._sP.map(function(s){return s.map(function(v){return{x:v.x+ox,y:v.y+oy}})})}:pl[pi].pts.map(function(v){return{x:v.x+ox,y:v.y+oy}});pl2.push({id:plc+1,sourceId:plc+1,points:pp2,rotation:0,_m:pl[pi]._mFlag||false});plc++}}}var len=0;pl2.forEach(function(p){var ps=p.points;if(ps._sP)ps._sP.forEach(function(sp){sp.forEach(function(pt){if(pt.y>len)len=pt.y})});else(ps||[]).forEach(function(pt){if(pt.y>len)len=pt.y})});return{placements:pl2,markerLengthMm:len}}',
        'function doCell(p,pts,varB,bb,varBB,bw,pairMode,fine){var sw=bb.maxX-bb.minX,sh=bb.maxY-bb.minY,mw=varBB.maxX-varBB.minX,mh=varBB.maxY-varBB.minY,bestDx=sw,bestDy=0,bestH=1/0,bi=fine?6:4,ds=Math.max(1,Math.round((sh+mh)/(fine?30:18))),sp=sg(pts,300),sb=sg(varB,300),my,lo,hi,md,mp,i,ch,cw;for(my=-(mh|0);my<=(sh|0);my+=ds){lo=0;hi=sw+mw;for(i=0;i<bi;i++){md=(lo+hi)/2;mp=sp._sP?{_sP:sb._sP.map(function(sp){return sp.map(function(v){return{x:v.x+md,y:v.y+my}})})}:sb.map(function(v){return{x:v.x+md,y:v.y+my}});if(_cOvl(sp,mp,0))lo=md;else hi=md}ch=Math.max(bb.maxY,varBB.maxY+hi)-Math.min(bb.minY,varBB.minY+my);cw=Math.max(bb.maxX,varBB.maxX+hi)-Math.min(bb.minX,varBB.minX+my);if(ch<bestH&&cw<=bw){bestH=ch;bestDx=hi;bestDy=my}lo=-(sw+mw);hi=0;for(i=0;i<bi;i++){md=(lo+hi)/2;mp=sp._sP?{_sP:sb._sP.map(function(sp){return sp.map(function(v){return{x:v.x+md,y:v.y+my}})})}:sb.map(function(v){return{x:v.x+md,y:v.y+my}});if(_cOvl(sp,mp,0))hi=md;else lo=md}ch=Math.max(bb.maxY,varBB.maxY+lo)-Math.min(bb.minY,varBB.minY+my);cw=Math.max(bb.maxX,varBB.maxX+lo)-Math.min(bb.minX,varBB.minX+my);if(ch<bestH&&cw<=bw){bestH=ch;bestDx=lo;bestDy=my}}var cl=Math.min(bb.minX,varBB.minX+bestDx),ct=Math.min(bb.minY,varBB.minY+bestDy);cw=Math.max(bb.maxX,varBB.maxX+bestDx)-cl;bestH=Math.max(bb.maxY,varBB.maxY+bestDy)-ct;var ns=pairMode?p.length:Math.ceil(p.length/2);var pp=Math.max(1,Math.floor((bw-0.5)/cw)),rs=Math.ceil(ns/pp),pl=[],plc=0,r,c2,ox,oy;if(cw>bw+0.5)return{placements:[],markerLengthMm:1/0};for(r=0;r<rs&&plc<p.length;r++){for(c2=0;c2<pp&&plc<p.length;c2++){ox=c2*cw-cl;oy=r*bestH-ct;if(pts._sP){pl.push({id:p[plc].id,sourceId:p[plc].sourceId||p[plc].id,points:{_sP:pts._sP.map(function(sp){return sp.map(function(v){return{x:v.x+ox,y:v.y+oy}})})},rotation:0,_m:false})}else{pl.push({id:p[plc].id,sourceId:p[plc].sourceId||p[plc].id,points:pts.map(function(v){return{x:v.x+ox,y:v.y+oy}}),rotation:0,_m:false})}if(!pairMode){plc++;if(plc>=p.length)break}if(varB._sP){pl.push({id:p[plc].id,sourceId:p[plc].sourceId||p[plc].id,points:{_sP:varB._sP.map(function(sp){return sp.map(function(v){return{x:v.x+ox+bestDx,y:v.y+oy+bestDy}})})},rotation:0,_m:pairMode})}else{pl.push({id:p[plc].id,sourceId:p[plc].sourceId||p[plc].id,points:varB.map(function(v){return{x:v.x+ox+bestDx,y:v.y+oy+bestDy}}),rotation:0,_m:pairMode})}plc++}}var len=0;pl.forEach(function(pc){var ps=pc.points;if(ps._sP){ps._sP.forEach(function(sp){sp.forEach(function(pt){if(pt.y>len)len=pt.y})})}else{(ps||[]).forEach(function(pt){if(pt.y>len)len=pt.y})}});return{placements:pl,markerLengthMm:len}}',
        'function nestAngle(parts,bw,a,cpt,pairEn){var ref=parts[0],geo=ref._sP?ref:ref.points,pts=rotPtsArb(geo,a),bb=getBB(pts._sP?pts:{points:pts}),sw=bb.maxX-bb.minX,sh=bb.maxY-bb.minY;if(sw+0.5>bw)return{placements:[],markerLengthMm:1/0};var _s=sg(pts,300);var slo=0,shi=sw;for(var si=0;si<8;si++){var sm=(slo+shi)/2,sp=_s._sP?{_sP:_s._sP.map(function(sp){return sp.map(function(v){return{x:v.x+sm,y:v.y}})})}:_s.map(function(v){return{x:v.x+sm,y:v.y}});if(_cOvl(_s,sp,0))slo=sm;else shi=sm}var cx=(bb.minX+bb.maxX)/2,mir=pts._sP?{_sP:pts._sP.map(function(sp){return sp.map(function(v){return{x:-(v.x-cx)+cx,y:v.y}})})}:pts.map(function(v){return{x:-(v.x-cx)+cx,y:v.y}}),mbb=pts._sP?getBB(mir):getBB({points:mir}),mw=mbb.maxX-mbb.minX,mh=mbb.maxY-mbb.minY,mok=mw<=bw;if(mok&&pairEn){var pR=doCell(parts,pts,mir,bb,mbb,bw,true,cpt);var pR2=doCell(parts,mir,pts,mbb,bb,bw,true,cpt);if(pR2&&pR2.markerLengthMm<pR.markerLengthMm)pR=pR2;var mR;if(parts.length>=4&&parts.length<=24){mR=pairMotif(pts,mir,bb,mbb,bw,parts.length,cpt);if(mR&&mR.markerLengthMm<pR.markerLengthMm)return mR}if(cpt)console.log(\'pairDebug\',\'doCell\',pR.markerLengthMm,\'swap\',pR2.markerLengthMm,\'motif\',mR?mR.markerLengthMm:-1,\'n\',parts.length,\'a\',a);return pR}var n=parts.length,bestC=null,bestL=1/0;if(n>=2){var a180=a+180;if(a180>360)a180-=360;var pts180=rotPtsArb(geo,a180),bb180=getBB(pts180._sP?pts180:{points:pts180});var cr;if(mok){cr=doCell(parts,pts,mir,bb,mbb,bw,false,cpt);if(cr.markerLengthMm<bestL){bestC=cr;bestL=cr.markerLengthMm}if(cpt===0&&bestL<1/0)return bestC}cr=doCell(parts,pts,pts180,bb,bb180,bw,false,cpt);if(cr.markerLengthMm<bestL){bestC=cr;bestL=cr.markerLengthMm}cr=doCell(parts,pts,pts,bb,bb,bw,false,cpt);if(cr.markerLengthMm<bestL){bestC=cr;bestL=cr.markerLengthMm}}if(n>=4&&n<=24){var mv=[{pts:pts,bb:bb}];if(mok)mv.push({pts:mir,bb:mbb});if(bb180.maxX-bb180.minX<=bw)mv.push({pts:pts180,bb:bb180});if(mok){var cx180=(bb180.minX+bb180.maxX)/2,mir180=pts180._sP?{_sP:pts180._sP.map(function(sp){return sp.map(function(v){return{x:-(v.x-cx180)+cx180,y:v.y}})})}:pts180.map(function(v){return{x:-(v.x-cx180)+cx180,y:v.y}}),mbb180=pts180._sP?getBB(mir180):getBB({points:mir180});if(mbb180.maxX-mbb180.minX<=bw)mv.push({pts:mir180,bb:mbb180})}var mr=buildMotif(pts,bb,mv,bw,n,cpt);if(mr&&mr.markerLengthMm<bestL){bestC=mr;bestL=mr.markerLengthMm}}if(bestC&&bestL<1/0)return bestC;var rows=[],placed=0,y=0;var step=shi+1;if(mok&&pairEn){var sl=0,s2=sw;for(var i=0;i<8;i++){var md=(sl+s2)/2,mp=pts._sP?{_sP:mir._sP.map(function(sp){return sp.map(function(v){return{x:v.x+md,y:v.y}})})}:mir.map(function(v){return{x:v.x+md,y:v.y}});if(_cOvl(sp,mp,0))sl=md;else s2=md}if(s2+1>step)step=s2+1}var pts90=rotPtsArb(geo,a+90),bb90=getBB(pts90._sP?pts90:{points:pts90}),sw90=bb90.maxX-bb90.minX,sh90=bb90.maxY-bb90.minY,cx90=(bb90.minX+bb90.maxX)/2,mir90=pts90._sP?{_sP:pts90._sP.map(function(sp){return sp.map(function(v){return{x:-(v.x-cx90)+cx90,y:v.y}})})}:pts90.map(function(v){return{x:-(v.x-cx90)+cx90,y:v.y}}),mbb90=pts90._sP?getBB(mir90):getBB({points:mir90}),mw90=mbb90.maxX-mbb90.minX,mh90=mbb90.maxY-mbb90.minY,mok90=mw90<=bw;while(placed<n){var r90=mok90&&sw90<sw,rPts=r90?pts90:pts,rBb=r90?bb90:bb,rSh=r90?sh90:sh,rMir=r90?mir90:mir,rMbb=r90?mbb90:mbb,rMok=r90?mok90:mok,row=[],x=0;while(placed<n){var use=rMok&&row.length%2===1,cp=use?rMir:rPts,cb=use?rMbb:rBb,cw=cb.maxX-cb.minX;if(x+cw+0.5>bw){var _tv=false;if(cp._sP){for(var _vi=0;_vi<cp._sP.length;_vi++){for(var _vj=0;_vj<cp._sP[_vi].length;_vj++){if(x+cp._sP[_vi][_vj].x-cb.minX>bw){_tv=true;break}}if(_tv)break}}else{for(var _vi=0;_vi<cp.length;_vi++){if(x+cp[_vi].x-cb.minX>bw){_tv=true;break}}}if(_tv)break}var dx=x-cb.minX,dy=y-cb.minY,copy=cp._sP?{_sP:cp._sP.map(function(sp){return sp.map(function(v){return{x:v.x+dx,y:v.y+dy}})})}:cp.map(function(v){return{x:v.x+dx,y:v.y+dy}}),ol=false;for(var k=0;k<row.length&&!ol;k++){if(_cOvl(copy,row[k].points))ol=true}if(ol){if(use){var dx2=x-rBb.minX,copy2=rPts._sP?{_sP:rPts._sP.map(function(sp){return sp.map(function(v){return{x:v.x+dx2,y:v.y+dy}})})}:rPts.map(function(v){return{x:v.x+dx2,y:v.y+dy}}),ol2=false;for(var k=0;k<row.length&&!ol2;k++){if(_cOvl(copy2,row[k].points))ol2=true}if(!ol2){row.push({id:parts[placed].id,sourceId:parts[placed].sourceId||parts[placed].id,points:copy2,rotation:r90?a+90:a});x+=step;placed++;continue}}break}row.push({id:parts[placed].id,sourceId:parts[placed].sourceId||parts[placed].id,points:copy,rotation:r90?a+90:a,_m:use});x+=step;placed++}if(!row.length)break;rows.push(row);y+=Math.max(rSh,rMok?(r90?mh90:mh):0)}var pl=[];rows.forEach(function(r){r.forEach(function(p){pl.push(p)})});if(cpt&&pl.length>1){for(var ii=pl.length-1;ii>0;ii--){var pi=pl[ii],bi=getBB(pi.points._sP?pi.points:{points:pi.points}),lo=0,hi=bi.minY,orig=pi.points._sP?{_sP:pi.points._sP.map(function(sp){return sp.map(function(v){return{x:v.x,y:v.y}})})}:pi.points.map(function(v){return{x:v.x,y:v.y}});for(var it=0;it<8;it++){var md=(lo+hi)/2,sf=md-bi.minY,can=orig._sP?{_sP:orig._sP.map(function(sp){return sp.map(function(v){return{x:v.x,y:v.y+sf}})})}:orig.map(function(v){return{x:v.x,y:v.y+sf}}),cb2=getBB(can._sP?can:{points:can}),ol=false;for(var j=0;j<pl.length&&!ol;j++){if(j===ii)continue;var ob=getBB(pl[j].points._sP?pl[j].points:{points:pl[j].points});if(cb2.maxX>ob.minX&&cb2.minX<ob.maxX&&_cOvl(can,pl[j].points))ol=true}if(ol)lo=md;else hi=md}var bs=hi-bi.minY;if(bs<-0.1)pi.points=orig._sP?{_sP:orig._sP.map(function(sp){return sp.map(function(v){return{x:v.x,y:v.y+bs}})})}:orig.map(function(v){return{x:v.x,y:v.y+bs}})}}var len=0;pl.forEach(function(p){var _pts=p.points;if(_pts&&_pts._sP){_pts._sP.forEach(function(sp){sp.forEach(function(pt){if(pt.y>len)len=pt.y})})}else{(_pts||[]).forEach(function(pt){if(pt.y>len)len=pt.y})}});return{placements:pl,markerLengthMm:len}}',
        'function runNest(parts,bw,it,sa,ea,pairEn){var best=null,bestLen=1/0,pn=parts.length,cStep=pn>80?15:(it>100?(pn>20?5:8):(pn>20?10:8)),topN=pn>80?2:(it>100?3:2),fRng=pn>80?5:(it>100?8:5),tested={},cTotal=Math.ceil((ea-sa)/cStep),cIdx=0,coarse=[];for(var a=sa;a<ea;a+=cStep){tested[a]=1;cIdx++;var res=nestAngle(parts,bw,a,0,pairEn);coarse.push({a:a,len:res.markerLengthMm});if(res.markerLengthMm<bestLen){bestLen=res.markerLengthMm;best=res}self.postMessage({type:"progress",phase:"coarse",pPass:cIdx,pTotal:cTotal,bestLen:bestLen,placements:best?best.placements:[]})}coarse.sort(function(a,b){return a.len-b.len});var fAngles=[];for(var ci=0;ci<Math.min(topN,coarse.length);ci++){var ctr=coarse[ci].a,fs=Math.max(sa,ctr-fRng),fe=Math.min(ea,ctr+fRng);for(var a=fs;a<=fe;a++){if(tested[a])continue;fAngles.push(a);tested[a]=1}}if(!tested[0]&&0>=sa&&0<=ea){fAngles.push(0);tested[0]=1}var fTotal=fAngles.length;for(var fi=0;fi<fTotal;fi++){var res=nestAngle(parts,bw,fAngles[fi],it>4,pairEn);if(res.markerLengthMm<bestLen){bestLen=res.markerLengthMm;best=res}self.postMessage({type:"progress",phase:"fine",pPass:fi+1,pTotal:fTotal,bestLen:bestLen,placements:best?best.placements:[]})}self.postMessage({type:"result",placements:best?best.placements:[],markerLengthMm:best?best.markerLengthMm:0})}',
        'self.onmessage=function(e){var m=e.data;if(m.type==="start"){var parts=m.parts,bw=m.binWidth,md=m.mode||"shuffle",it=m.iterations||10,pe=m.pairEn;if(md==="strip")runNest(parts,bw,it,m.angleStart||0,m.angleEnd||360,pe);else runShuffle(parts,bw,it)}}'
      ].join('');
    },

    _isHomogeneous: function(parts) {
      if (!parts || parts.length <= 1) return true;
      var sigs = [];
      for (var i = 0; i < parts.length; i++) {
        var p = parts[i];
        var pts = p.points || [];
        var mx = Infinity, mn = -Infinity, my = Infinity, ny = -Infinity;
        for (var j = 0; j < pts.length; j++) {
          if (pts[j].x < mx) mx = pts[j].x;
          if (pts[j].x > mn) mn = pts[j].x;
          if (pts[j].y < my) my = pts[j].y;
          if (pts[j].y > ny) ny = pts[j].y;
        }
        sigs.push(pts.length + '|' + Math.round((mn - mx) * (ny - my)));
      }
      return sigs.every(function(s) { return s === sigs[0]; });
    },

    runNesting: function(parts, binWidthMm, options) {
      options = options || {};
      var population = options.iterations || 10;
      var onProgress = options.onProgress || null;
      var mode = options.mode || 'shuffle';

      // Strip mode uses first part's geometry for all -- falls back to shuffle when heterogeneous
      if (mode === 'strip' && !MarkerNesting._isHomogeneous(parts)) {
        console.warn('Heterogeneous parts detected with strip mode -- falling back to shuffle');
        mode = 'shuffle';
      }

      if (typeof Worker === 'undefined') {
        // Fallback: synchronous on main thread
        var best = null, bestLen = Infinity;
        for (var ind = 0; ind < population; ind++) {
          var order = parts.slice();
          if (ind > 0) {
            for (var j = order.length - 1; j > 0; j--) {
              var k = Math.floor(Math.random() * (j + 1));
              var tmp = order[j]; order[j] = order[k]; order[k] = tmp;
            }
          }
          var result = MarkerNesting.rectanglePack(order, binWidthMm);
          if (result.placements.length > 1) {
            MarkerNesting._compactDown(result.placements, binWidthMm);
            MarkerNesting._compactLeft(result.placements, binWidthMm);
          }
          if (result.markerLengthMm < bestLen) { bestLen = result.markerLengthMm; best = result; }
          if (onProgress) onProgress(ind + 1, population, bestLen, best ? best.placements : null, Math.round((ind + 1) / population * 100), 'pass');
        }
        if (best) MarkerNesting._propagateSP(best.placements, parts);
        MarkerNesting._mapColors(best ? best.placements : [], parts);
        return Promise.resolve({ placements: best ? best.placements : [], markerLengthMm: best ? best.markerLengthMm : 0 });
      }

      if (!MarkerNesting._workerBlobUrl) {
        MarkerNesting._workerBlobUrl = URL.createObjectURL(
          new Blob([MarkerNesting._getWorkerCode()], { type: 'application/javascript' })
        );
      }

      var blobUrl = MarkerNesting._workerBlobUrl;

      var legalInset = options.inset || 10;
      var pairEn = options.pairEn || false;
      if (pairEn && parts.some(function(p) { return p._pairMirror; })) {
        console.warn('Pre-built mirror parts detected — disabling pairEn to avoid double mirroring');
        pairEn = false;
      }
      var effWidth = binWidthMm - 2 * legalInset;
      var numRows = options.rows || 3;

      function shiftPlacements(pl) {
        if (legalInset <= 0 || !pl) return pl;
        var sx = legalInset, sy = legalInset;
        pl.forEach(function(p) {
          if (p.points) {
            if (p.points._sP) {
              p.points._sP = p.points._sP.map(function(sp) { return sp.map(function(v) { return { x: v.x + sx, y: v.y + sy }; }); });
            } else if (Array.isArray(p.points)) {
              p.points = p.points.map(function(v) { return { x: v.x + sx, y: v.y + sy }; });
            }
          }
          // Also shift _originalPoints to keep in sync
          if (p._originalPoints && Array.isArray(p._originalPoints)) {
            p._originalPoints = p._originalPoints.map(function(v) { return { x: v.x + sx, y: v.y + sy }; });
          }
          // Also shift top-level _sP to keep in sync
          if (p._sP && Array.isArray(p._sP)) {
            p._sP = p._sP.map(function(sp) {
              return sp.map(function(v) { return { x: v.x + sx, y: v.y + sy }; });
            });
          }
        });
        // Snap any placement that still crosses the boundary
        function snapPts(arr, d) { for (var i = 0; i < arr.length; i++) arr[i].x += d; }
        pl.forEach(function(p) {
          var pts = p.points || [];
          if (!pts || (Array.isArray(pts) && pts.length === 0)) return;
          var mx = Infinity, nx = -Infinity;
          if (pts._sP) {
            for (var si = 0; si < pts._sP.length; si++) {
              pts._sP[si].forEach(function(v) { if (v.x < mx) mx = v.x; if (v.x > nx) nx = v.x; });
            }
          } else {
            pts.forEach(function(v) { if (v.x < mx) mx = v.x; if (v.x > nx) nx = v.x; });
          }
          if (mx < legalInset) {
            var sd = legalInset - mx;
            if (pts._sP) { for (var si = 0; si < pts._sP.length; si++) snapPts(pts._sP[si], sd); }
            else { snapPts(pts, sd); }
            if (p._originalPoints && Array.isArray(p._originalPoints)) snapPts(p._originalPoints, sd);
            if (p._sP && Array.isArray(p._sP)) { for (var si = 0; si < p._sP.length; si++) snapPts(p._sP[si], sd); }
            mx += sd; nx += sd;
          }
          if (nx > binWidthMm - legalInset) {
            var sd2 = (binWidthMm - legalInset) - nx;
            if (pts._sP) { for (var si = 0; si < pts._sP.length; si++) snapPts(pts._sP[si], sd2); }
            else { snapPts(pts, sd2); }
            if (p._originalPoints && Array.isArray(p._originalPoints)) snapPts(p._originalPoints, sd2);
            if (p._sP && Array.isArray(p._sP)) { for (var si = 0; si < p._sP.length; si++) snapPts(p._sP[si], sd2); }
          }
        });
        return pl;
      }

      // Row mode: simple row arrangement (no worker needed)
      if (mode === 'row') {
        var partsCopy = parts.slice();
        var rowWidths = new Array(numRows).fill(0);
        var rowParts = [];
        for (var ri = 0; ri < numRows; ri++) rowParts[ri] = [];
        partsCopy.sort(function(a,b) { return b.height - a.height; });
        partsCopy.forEach(function(p) {
          var minR = 0, minW = rowWidths[0];
          for (var ri = 1; ri < numRows; ri++) {
            if (rowWidths[ri] < minW) { minW = rowWidths[ri]; minR = ri; }
          }
          rowParts[minR].push(p);
          rowWidths[minR] += p.width;
        });
        var by = 0, bpl = [];
        for (var ri = 0; ri < numRows; ri++) {
          var rp = rowParts[ri];
          if (rp.length === 0) continue;
          // Place original row
          var bx = 0, rowH = 0;
          for (var pi = 0; pi < rp.length; pi++) {
            var p = rp[pi], pb = p.boundingBox;
            var pw = pb.maxX - pb.minX, ph = pb.maxY - pb.minY;
            if (bx + pw > effWidth) continue;
            var dx = bx - pb.minX, dy = by - pb.minY;
            var copy = p.points.map(function(v) { return { x: v.x + dx, y: v.y + dy }; });
            bpl.push({ id: p.id, sourceId: p.sourceId || p.id, points: copy, rotation: 0, origColor: p.origColor });
            bx += pw;
            if (ph > rowH) rowH = ph;
          }
          by += rowH;
          // Place mirrored row below (if pair mode)
          if (pairEn) {
            var mbx = 0, mRowH = 0;
            for (var pi = 0; pi < rp.length; pi++) {
              var p = rp[pi], pb = p.boundingBox;
              var pw = pb.maxX - pb.minX, ph = pb.maxY - pb.minY;
              if (mbx + pw > effWidth) continue;
              var cx2 = (pb.minX + pb.maxX) / 2;
              var mirPts = p.points.map(function(v) { return { x: -(v.x - cx2) + cx2, y: v.y }; });
              var mirBb = { minX: -pb.maxX + 2*cx2, maxX: -pb.minX + 2*cx2, minY: pb.minY, maxY: pb.maxY };
              var dx = mbx - mirBb.minX, dy = by - mirBb.minY;
              var copy = mirPts.map(function(v) { return { x: v.x + dx, y: v.y + dy }; });
              bpl.push({ id: p.id, sourceId: p.sourceId || p.id, points: copy, rotation: 0, _m: true, origColor: p.origColor });
              mbx += pw;
              if (ph > mRowH) mRowH = ph;
            }
            by += mRowH;
          }
        }
        var rowLen = 0;
        bpl.forEach(function(p) { forEachPoint(p.points, function(pt) { if (pt.y > rowLen) rowLen = pt.y; }); });
        shiftPlacements(bpl);
        return Promise.resolve({ placements: bpl, markerLengthMm: rowLen });
      }

      // Parallel strip mode: split angle range across workers
      if (mode === 'strip') {
        var numWorkers = Math.min(options.maxWorkers || (typeof navigator !== 'undefined' && navigator.hardwareConcurrency) || 2, 4);
        var perWorker = Math.ceil(360 / numWorkers);
        var completed = 0;
        var combinedBest = null, combinedBestLen = Infinity;
        var workerStates = [];
        MarkerNesting._activeWorkers = [];
        MarkerNesting._resolveNesting = null;
        MarkerNesting._stopRequested = false;

        function computeProgress() {
          var tPass = 0, tTotal = 0, fine = 0, anyFine = false;
          for (var i = 0; i < numWorkers; i++) {
            var ws = workerStates[i];
            if (!ws) continue;
            var cPass, cTotal;
            if (ws.done) {
              cPass = (ws._coarseTotal || 0) + (ws._fineTotal || 1);
              cTotal = (ws._coarseTotal || 0) + (ws._fineTotal || 1);
            } else if (ws.phase === 'fine') {
              cPass = (ws._coarseTotal || 0) + (ws.pPass || 0);
              cTotal = (ws._coarseTotal || 0) + (ws.pTotal || 1);
            } else {
              cPass = ws.pPass || 0;
              cTotal = ws.pTotal || 1;
            }
            tPass += cPass;
            tTotal += cTotal;
            if (ws.phase === 'fine') anyFine = true;
            if (ws.done || ws.phase === 'fine') fine++;
          }
          var phase = fine === numWorkers ? 'fine' : (anyFine ? 'refine' : 'coarse');
          var pct = tTotal > 0 ? Math.round(tPass / tTotal * 100) : 0;
          return { pct: pct, phase: phase };
        }

        function sendProgress(bestLen, placements) {
          if (!onProgress) return;
          var displayPl = placements ? placements.map(function(p) {
            var pts = p.points;
            if (pts && pts._sP) { pts = pts._sP[0] || []; }
            return { id: p.id, sourceId: p.sourceId, rotation: p.rotation, _m: p._m, origColor: p.origColor, sourceFileName: p.sourceFileName, points: Array.isArray(pts) ? pts.map(function(v) { return {x:v.x,y:v.y}; }) : [] };
          }) : null;
          if (displayPl && displayPl.length > 0) shiftPlacements(displayPl);
          var p = computeProgress();
          onProgress(0, 0, bestLen, displayPl || null, p.pct, p.phase);
        }

        return new Promise(function(resolve) {
          MarkerNesting._resolveNesting = resolve;

          function spawnWorker(idx) {
            var startAngle = idx * perWorker;
            var endAngle = Math.min(startAngle + perWorker, 360);
            var worker = new Worker(blobUrl);
            MarkerNesting._activeWorkers.push(worker);
            workerStates[idx] = { phase: 'coarse', pPass: 0, pTotal: 0, done: false };

            worker._retries = 0;
            worker.onerror = function() {
              worker.terminate();
              if (worker._retries < 1) {
                worker._retries++;
                console.warn('Strip worker ' + idx + ' error — retrying');
                var w2 = new Worker(blobUrl);
                w2._retries = 1;
                w2.onerror = worker.onerror;
                w2.onmessage = worker.onmessage;
                workerStates[idx] = { phase: 'coarse', pPass: 0, pTotal: 0, done: false };
                MarkerNesting._activeWorkers[idx] = w2;
                worker = w2;
                w2.postMessage({type:'start', parts:parts, binWidth:effWidth, mode:'strip', iterations:population, angleStart:startAngle, angleEnd:endAngle, pairEn:pairEn});
              } else {
                workerDone(idx);
              }
            };
            worker.onmessage = function(e) {
              if (MarkerNesting._stopRequested) { worker.terminate(); workerDone(idx); return; }
              var msg = e.data;
              if (msg.type === 'progress') {
                if (msg.phase === 'coarse') workerStates[idx]._coarseTotal = msg.pTotal;
                if (msg.phase === 'fine') workerStates[idx]._fineTotal = msg.pTotal;
                workerStates[idx].pPass = msg.pPass || 0;
                workerStates[idx].pTotal = msg.pTotal || 0;
                workerStates[idx].phase = msg.phase || 'coarse';
                sendProgress(msg.bestLen, msg.placements);
              } else if (msg.type === 'result') {
                worker.terminate();
                workerStates[idx].done = true;
                if (!workerStates[idx]._fineTotal) workerStates[idx]._fineTotal = workerStates[idx]._coarseTotal || 1;
                if (msg.markerLengthMm < combinedBestLen) {
                  combinedBestLen = msg.markerLengthMm;
                  combinedBest = msg;
                }
                sendProgress(combinedBestLen, combinedBest ? combinedBest.placements : null);
                workerDone(idx);
              }
            };
            worker.postMessage({type:'start', parts:parts, binWidth:effWidth, mode:'strip', iterations:population, angleStart:startAngle, angleEnd:endAngle, pairEn:pairEn});
          }
          function workerDone(idx) {
            if (workerStates[idx]) workerStates[idx].done = true;
            completed++;
          if (completed === numWorkers) {
          MarkerNesting._activeWorkers = [];
          MarkerNesting._resolveNesting = null;
          MarkerNesting._stopRequested = false;
          var finPl = combinedBest ? shiftPlacements(combinedBest.placements) : [];
          var finLen = combinedBestLen;
          if (pairEn) {
            var rawLen = combinedBest ? combinedBest.placements.length : 0;
            if (rawLen !== parts.length * 2) {
              console.error('PAIR INVARIANT: expected ' + (parts.length * 2) + ' placements, got ' + rawLen);
              finPl = [];
              finLen = 0;
            }
          }
          if (finPl.length > 1) {
            MarkerNesting._compactDown(finPl, binWidthMm, legalInset);
            MarkerNesting._compactLeft(finPl, binWidthMm, legalInset);
            MarkerNesting._compactPoly(finPl, binWidthMm, legalInset);
            finLen = 0;
            finPl.forEach(function(p) {
              var ps = p.points;
              if (!ps) return;
              if (ps._sP) {
                ps._sP.forEach(function(sp) { sp.forEach(function(pt) { if (pt.y > finLen) finLen = pt.y; }); });
              } else {
                ps.forEach(function(pt) { if (pt.y > finLen) finLen = pt.y; });
              }
            });
          }
          // Cross-set snap candidate: flatten composites into individual sub-parts and pack row by row
          var csTriggered = false, csRejectReason = '', csWin = false;
          if (pairEn && parts.length >= 2 && parts[0] && parts[0]._sP && parts[0]._sP.length >= 2) {
            csTriggered = true;
            var csGap = 0.5;
            var csItems = [];
            var hasBuiltInMirrors = parts.some(function(p) { return p._pairMirror; });
            function csBB(pts) { var r={minX:Infinity,maxX:-Infinity,minY:Infinity,maxY:-Infinity}; pts.forEach(function(v){if(v.x<r.minX)r.minX=v.x;if(v.x>r.maxX)r.maxX=v.x;if(v.y<r.minY)r.minY=v.y;if(v.y>r.maxY)r.maxY=v.y}); return r; }
            for (var ci = 0; ci < parts.length; ci++) {
              var p = parts[ci];
              if (!p._sP || p._sP.length < 2) continue;
              var offPts = (p.points && p.points.length >= 3) ? p.points.map(function(v) { return {x: v.x, y: v.y}; }) : MarkerNesting.computeAutoOffset(p, 1);
              if (!offPts || offPts.length < 3) continue;
              var origSP = p._sP.map(function(sp) { return sp.map(function(v) { return {x: v.x, y: v.y}; }); });
              var bb = csBB(offPts);
              // _childOffsets may be on offPts (from computeAutoOffset) or on p.points (when pre-computed)
              var childOffSrc = offPts._childOffsets || (p.points && p.points._childOffsets) || null;
              var childOff = childOffSrc ? childOffSrc.map(function(co) { return co.map(function(v) { return {x:v.x,y:v.y}; }); }) : null;
              csItems.push({offPts:offPts, origSP:origSP, childOff:childOff, bb:bb, w:bb.maxX-bb.minX, h:bb.maxY-bb.minY, origId:p.id, sourceId:p.sourceId, isMirror:!!(p._pairMirror||p._m), origColor:p.origColor});
            }
            if (!hasBuiltInMirrors) {
              var mirItems = [];
              for (var ii = 0; ii < csItems.length; ii++) {
                var it = csItems[ii];
                var cx = (it.bb.minX + it.bb.maxX) / 2;
                var mirOff = it.offPts.map(function(v){return{x:-(v.x-cx)+cx,y:v.y}});
                var mirSP = it.origSP.map(function(sp) { return sp.map(function(v){return{x:-(v.x-cx)+cx,y:v.y};}); });
                var mirChildOff = it.childOff ? it.childOff.map(function(co) { return co.map(function(v){return{x:-(v.x-cx)+cx,y:v.y};}); }) : null;
                var mbb = csBB(mirOff);
                mirItems.push({offPts:mirOff, origSP:mirSP, childOff:mirChildOff, bb:mbb, w:mbb.maxX-mbb.minX, h:mbb.maxY-mbb.minY, origId:it.origId+1000, sourceId:it.sourceId, isMirror:true, origColor:it.origColor});
              }
              csItems = csItems.concat(mirItems);
            }
            if (csItems.length >= 4) {
              csItems.sort(function(a,b){return b.h-a.h});
              var csRows = [], csPl = [], csY = 0;
              for (var ii = 0; ii < csItems.length; ii++) {
                var it = csItems[ii];
                var placed = false;
                for (var ri = 0; ri < csRows.length; ri++) {
                  var row = csRows[ri];
                  if (row.cursor + it.w <= effWidth) {
                    var dx = row.cursor - it.bb.minX, dy = row.y - it.bb.minY;
                    var newOff = it.offPts.map(function(v){return{x:v.x+dx,y:v.y+dy}});
                    var newSP = it.origSP.map(function(sp){return sp.map(function(v){return{x:v.x+dx,y:v.y+dy};});});
                    var newChildOff = it.childOff ? it.childOff.map(function(co){return co.map(function(v){return{x:v.x+dx,y:v.y+dy};});}) : null;
                    csPl.push({id:it.origId,sourceId:it.sourceId,points:newOff,rotation:0,_m:it.isMirror,origColor:it.origColor,_sP:newSP,_childOffsets:newChildOff,_unionOffset:true});
                    row.cursor += it.w + csGap;
                    row.h = Math.max(row.h, it.h);
                    placed = true; break;
                  }
                }
                if (!placed) {
                  var dx = 0 - it.bb.minX, dy = csY - it.bb.minY;
                  var newOff = it.offPts.map(function(v){return{x:v.x+dx,y:v.y+dy}});
                  var newSP = it.origSP.map(function(sp){return sp.map(function(v){return{x:v.x+dx,y:v.y+dy};});});
                  var newChildOff = it.childOff ? it.childOff.map(function(co){return co.map(function(v){return{x:v.x+dx,y:v.y+dy};});}) : null;
                  csPl.push({id:it.origId,sourceId:it.sourceId,points:newOff,rotation:0,_m:it.isMirror,origColor:it.origColor,_sP:newSP,_childOffsets:newChildOff,_unionOffset:true});
                  csRows.push({y:csY, h:it.h, cursor:it.w+csGap});
                  csY += it.h + csGap;
                }
              }
              shiftPlacements(csPl);
              if (csPl.length > 1) {
                MarkerNesting._compactPoly(csPl, binWidthMm, legalInset);
              }
              var csLen = 0;
              csPl.forEach(function(p){
                if (p._sP && Array.isArray(p._sP)) { p._sP.forEach(function(sp){sp.forEach(function(pt){if(pt.y>csLen)csLen=pt.y});}); }
                else { (p.points||[]).forEach(function(pt){if(pt.y>csLen)csLen=pt.y}); }
              });
              if (csLen < finLen) { finPl = csPl; finLen = csLen; csWin = true; csRejectReason = ''; }
              else { csRejectReason = 'csLen='+csLen.toFixed(2)+' >= finLen='+finLen.toFixed(2); }
            } else {
              csRejectReason = 'csItems.length='+csItems.length+' < 4';
            }
          } else {
            csRejectReason = '!pairEn='+pairEn+' parts.length='+parts.length+' _sP='+(parts[0]&&parts[0]._sP?parts[0]._sP.length:'n/a');
          }
          MarkerNesting._propagateSP(finPl, parts);
          MarkerNesting._mapColors(finPl, parts);
          resolve({ placements: finPl, markerLengthMm: finLen, csTriggered: csTriggered, csWin: csWin, csRejectReason: csRejectReason });
            }
          }
          for (var w = 0; w < numWorkers; w++) spawnWorker(w);
        });
      }

      // Single worker for shuffle mode
      return new Promise(function(resolve) {
        var worker = new Worker(blobUrl);
        worker._retries = 0;
        worker.onerror = function(err) {
          worker.terminate();
          if (worker._retries < 1) {
            worker._retries++;
            console.warn('Nesting worker error — retrying (' + worker._retries + '):', err.message);
            var w2 = new Worker(blobUrl);
            w2._retries = 1;
            w2.onerror = worker.onerror;
            w2.onmessage = worker.onmessage;
            worker = w2;
            MarkerNesting._activeWorkers = [w2];
            w2.postMessage({type:'start', parts:parts, binWidth:effWidth, mode:'shuffle', iterations:population, pairEn:pairEn});
          } else {
            console.error('Nesting worker error (no retries left):', err.message, err.filename, err.lineno);
            resolve({ placements: [], markerLengthMm: 0 });
          }
        };
        worker.onmessage = function(e) {
          var msg = e.data;
          if (msg.type === 'progress') {
            var pct = msg.total > 0 ? Math.round(msg.pass / msg.total * 100) : 0;
            var progressPl = msg.placements ? msg.placements.map(function(p) {
              return { id: p.id, sourceId: p.sourceId, rotation: p.rotation, _m: p._m, origColor: p.origColor, points: p.points.map(function(v) { return {x:v.x,y:v.y}; }) };
            }) : null;
            if (progressPl && progressPl.length > 0) shiftPlacements(progressPl);
            if (onProgress) onProgress(msg.pass, msg.total, msg.bestLen, progressPl || null, pct, 'pass');
          } else if (msg.type === 'result') {
            worker.terminate();
            var shuffledPl = shiftPlacements(msg.placements || []);
            MarkerNesting._propagateSP(shuffledPl, parts);
            MarkerNesting._mapColors(shuffledPl, parts);
            resolve({ placements: shuffledPl, markerLengthMm: msg.markerLengthMm || 0 });
          }
        };
        worker.postMessage({
          type: 'start',
          parts: parts,
          binWidth: effWidth,
          mode: 'shuffle',
          iterations: population,
          pairEn: pairEn
        });
      });
    },

    stop: function() {
      MarkerNesting._stopRequested = true;
      if (MarkerNesting._activeWorkers) {
        MarkerNesting._activeWorkers.forEach(function(w) { try { w.terminate(); } catch(e) {} });
        MarkerNesting._activeWorkers = [];
      }
      if (MarkerNesting._resolveNesting) {
        // Resolve with current markerPlacements (set globally by onProgress)
        var cur = typeof markerPlacements !== 'undefined' ? markerPlacements : [];
        var len = 0;
        if (cur.length > 0) {
          for (var i = 0; i < cur.length; i++) {
            forEachPoint(cur[i].points, function(p) { if (p.y > len) len = p.y; });
          }
        }
        MarkerNesting._resolveNesting({ placements: cur, markerLengthMm: len });
        MarkerNesting._resolveNesting = null;
      }
      MarkerNesting._stopRequested = false;
    },

    _propagateSP: function(placements, parts) {
      if (!placements || !parts) return;
      var partById = {}, origBySrcId = {}, mirrBySrcId = {};
      parts.forEach(function(p) {
        partById[p.id] = p;
        var srcKey = p.sourceId;
        if (srcKey != null) {
          if (p._pairMirror) mirrBySrcId[srcKey] = p;
          else origBySrcId[srcKey] = p;
          partById[srcKey] = p;
        }
      });
      placements.forEach(function(pl) {
        if (pl._csIndividual) return;
        if (!pl._originalPoints) {
          var ref = partById[pl.id];
          if (!ref) {
            var srcKey = pl.sourceId != null ? pl.sourceId : pl.id;
            var isMirror = pl._m || false;
            ref = isMirror ? (mirrBySrcId[srcKey] || partById[srcKey]) : (origBySrcId[srcKey] || partById[srcKey]);
          }
          if (ref && ref._originalPoints && Array.isArray(ref._originalPoints) && ref.points && Array.isArray(ref.points)) {
            var pts = pl.points;
            if (pts && pts.length > 0 && Array.isArray(pts)) {
              function bbCenter(arr) { 
                if (!arr || arr.length === 0) return {cx: 0, cy: 0};
                var mx=Infinity,nx=-mx,my=mx,ny=-mx; 
                for(var i=0;i<arr.length;i++){
                  if(!isFinite(arr[i].x) || !isFinite(arr[i].y)) continue;
                  if(arr[i].x<mx)mx=arr[i].x;
                  if(arr[i].x>nx)nx=arr[i].x;
                  if(arr[i].y<my)my=arr[i].y;
                  if(arr[i].y>ny)ny=arr[i].y;
                }
                return {cx:(mx+nx)/2,cy:(my+ny)/2}; 
              }
              var plBBC = bbCenter(pts);
              var refPts = ref.points;
              var rot = pl.rotation || 0;
              // Compute BB center for rotation (matches rotPtsArb / rotPts which use BB center)
              var refBBC = bbCenter(refPts);
              var rcx = refBBC.cx, rcy = refBBC.cy;
              if (rot !== 0) {
                var rad = rot * Math.PI / 180;
                var cosR = Math.cos(rad), sinR = Math.sin(rad);
                // Rotate ref points around BB center, then align centers
                var rotatedRef = ref._originalPoints.map(function(v) {
                  var dvx = v.x - rcx, dvy = v.y - rcy;
                  return { x: rcx + dvx * cosR - dvy * sinR, y: rcy + dvx * sinR + dvy * cosR };
                });
                // Find center of rotated ref
                var rotRefBBC = bbCenter(rotatedRef);
                // Translate rotated ref to match placement center
                var dx = plBBC.cx - rotRefBBC.cx;
                var dy = plBBC.cy - rotRefBBC.cy;
                var newOrigPts = rotatedRef.map(function(v) {
                  return { x: v.x + dx, y: v.y + dy };
                });
                // Validate all points are finite before assigning
                var allValid = newOrigPts.every(function(v) { return isFinite(v.x) && isFinite(v.y); });
                if (allValid) {
                  pl._originalPoints = newOrigPts;
                } else {
                  // Log diagnostic: which points are invalid?
                  var invalidCount = newOrigPts.filter(function(v) { return !isFinite(v.x) || !isFinite(v.y); }).length;
                  console.warn('_propagateSP: ' + invalidCount + ' invalid points in rotated _originalPoints for placement ' + pl.id);
                  // Keep original reference points as fallback
                  pl._originalPoints = ref._originalPoints.map(function(v) { return {x: v.x, y: v.y}; });
                }
              } else {
                if (refPts.length >= 2 && pts.length >= 2) {
                  var refAngle = Math.atan2(refPts[1].y - refPts[0].y, refPts[1].x - refPts[0].x);
                  var plAngle = Math.atan2(pts[1].y - pts[0].y, pts[1].x - pts[0].x);
                  var rawDetected = ((plAngle - refAngle) * 180 / Math.PI + 360) % 360;
                  var isMirrored = pl._m || false;
                  var actualRot;
                  if (isMirrored) {
                    actualRot = ((180 - 2 * refAngle * 180 / Math.PI - rawDetected) % 360 + 360) % 360;
                  } else {
                    actualRot = rawDetected;
                  }
                  pl._detectedRot = actualRot;
                  // Use min-corner alignment to match worker positioning
                  var rotatePts = function(arr, ang) {
                    if (Math.abs(ang) < 1e-4) return arr.map(function(v) { return {x:v.x,y:v.y}; });
                    var rad2 = ang * Math.PI / 180;
                    var c = Math.cos(rad2), s = Math.sin(rad2);
                    return arr.map(function(v) { var dx=v.x-rcx, dy=v.y-rcy; return {x:rcx+dx*c-dy*s, y:rcy+dx*s+dy*c}; });
                  };
                  if (actualRot > 1e-4 && actualRot < 359.9999) {
                    var rad = actualRot * Math.PI / 180;
                    var cosR = Math.cos(rad), sinR = Math.sin(rad);
                    var transformed;
                    if (isMirrored) {
                      // Mirror original around its own center, then rotate, then translate by worker's offset delta
                      var origBB = bbCenter(ref._originalPoints);
                      var refOffBB = bbCenter(refPts);
                      
                      // Compute worker's transform on the offset polygon
                      var refOffMirrored = refPts.map(function(v) {
                        return { x: -(v.x - refOffBB.cx) + refOffBB.cx, y: v.y };
                      });
                      var refOffTransformed = refOffMirrored.map(function(v) {
                        var dvx = v.x - refOffBB.cx, dvy = v.y - refOffBB.cy;
                        return { x: refOffBB.cx + dvx * cosR - dvy * sinR, y: refOffBB.cy + dvx * sinR + dvy * cosR };
                      });
                      // Find min-corners for translation
                      var refMinX = Infinity, refMinY = Infinity;
                      for (var ri = 0; ri < refOffTransformed.length; ri++) {
                        if (refOffTransformed[ri].x < refMinX) refMinX = refOffTransformed[ri].x;
                        if (refOffTransformed[ri].y < refMinY) refMinY = refOffTransformed[ri].y;
                      }
                      var plMinX = Infinity, plMinY = Infinity;
                      for (var pi = 0; pi < pts.length; pi++) {
                        if (pts[pi].x < plMinX) plMinX = pts[pi].x;
                        if (pts[pi].y < plMinY) plMinY = pts[pi].y;
                      }
                      var dx = plMinX - refMinX;
                      var dy = plMinY - refMinY;
                      
                      // Apply same transform to original
                      transformed = ref._originalPoints.map(function(v) {
                        var rx = -(v.x - origBB.cx) + origBB.cx;
                        var ry = v.y;
                        var dvx = rx - origBB.cx, dvy = ry - origBB.cy;
                        return { x: origBB.cx + dvx * cosR - dvy * sinR, y: origBB.cy + dvx * sinR + dvy * cosR };
                      });
                      var newOrigPts = transformed.map(function(v) { return { x: v.x + dx, y: v.y + dy }; });
                      var allValid = newOrigPts.every(function(v) { return isFinite(v.x) && isFinite(v.y); });
                      if (allValid) {
                        pl._originalPoints = newOrigPts;
                      } else {
                        console.warn('_propagateSP: invalid points in mirror _originalPoints for ' + pl.id);
                        pl._originalPoints = ref._originalPoints.map(function(v) { return {x: v.x, y: v.y}; });
                      }
                    } else {
                      transformed = ref._originalPoints.map(function(v) {
                        var dvx = v.x - rcx, dvy = v.y - rcy;
                        return { x: rcx + dvx * cosR - dvy * sinR, y: rcy + dvx * sinR + dvy * cosR };
                      });
                    }
                    // Only process non-mirror case (mirror handled above)
                    if (!isMirrored) {
                      // Find center of transformed ref
                      var tBBC = bbCenter(transformed);
                      // Translate to match placement center
                      var dx = plBBC.cx - tBBC.cx;
                      var dy = plBBC.cy - tBBC.cy;
                      var newOrigPts = transformed.map(function(v) {
                        return { x: v.x + dx, y: v.y + dy };
                      });
                      var allValid = newOrigPts.every(function(v) { return isFinite(v.x) && isFinite(v.y); });
                      if (allValid) {
                        pl._originalPoints = newOrigPts;
                      } else {
                        var invalidCount = newOrigPts.filter(function(v) { return !isFinite(v.x) || !isFinite(v.y); }).length;
                        console.warn('_propagateSP: ' + invalidCount + ' invalid points in transformed _originalPoints for placement ' + pl.id);
                        pl._originalPoints = ref._originalPoints.map(function(v) { return {x: v.x, y: v.y}; });
                      }
                    }
                  } else {
                    // No rotation: align using min-corner (mirrored pieces need X-flip first)
                    if (isMirrored) {
                      var origBB2 = bbCenter(ref._originalPoints);
                      var mirrored = ref._originalPoints.map(function(v) {
                        return { x: -(v.x - origBB2.cx) + origBB2.cx, y: v.y };
                      });
                      // Compute translation from reference OFFSET (not original)
                      var refBBC = bbCenter(refPts);
                      var refMirrored = refPts.map(function(v) {
                        return { x: -(v.x - refBBC.cx) + refBBC.cx, y: v.y };
                      });
                      var refMirrMinX = Infinity, refMirrMinY = Infinity;
                      for (var rmi = 0; rmi < refMirrored.length; rmi++) {
                        if (refMirrored[rmi].x < refMirrMinX) refMirrMinX = refMirrored[rmi].x;
                        if (refMirrored[rmi].y < refMirrMinY) refMirrMinY = refMirrored[rmi].y;
                      }
                      var plMinX = Infinity, plMinY = Infinity;
                      for (var pi2 = 0; pi2 < pts.length; pi2++) {
                        if (pts[pi2].x < plMinX) plMinX = pts[pi2].x;
                        if (pts[pi2].y < plMinY) plMinY = pts[pi2].y;
                      }
                      var dx = plMinX - refMirrMinX;
                      var dy = plMinY - refMirrMinY;
                      var newOrigPts = mirrored.map(function(v) { return { x: v.x + dx, y: v.y + dy }; });
                      var allValid = newOrigPts.every(function(v) { return isFinite(v.x) && isFinite(v.y); });
                      if (allValid) {
                        pl._originalPoints = newOrigPts;
                      } else {
                        console.warn('_propagateSP: invalid points in mirrored _originalPoints for ' + pl.id);
                        pl._originalPoints = ref._originalPoints.map(function(v) { return {x: v.x, y: v.y}; });
                      }
                    } else {
                      // No rotation, not mirrored: translate original by same delta as offset polygon
                      var refOffMinX = Infinity, refOffMinY = Infinity;
                      for (var ri3 = 0; ri3 < refPts.length; ri3++) {
                        if (refPts[ri3].x < refOffMinX) refOffMinX = refPts[ri3].x;
                        if (refPts[ri3].y < refOffMinY) refOffMinY = refPts[ri3].y;
                      }
                      var plMinX2 = Infinity, plMinY2 = Infinity;
                      for (var pi3 = 0; pi3 < pts.length; pi3++) {
                        if (pts[pi3].x < plMinX2) plMinX2 = pts[pi3].x;
                        if (pts[pi3].y < plMinY2) plMinY2 = pts[pi3].y;
                      }
                      var dx2 = plMinX2 - refOffMinX;
                      var dy2 = plMinY2 - refOffMinY;
                      var newOrigPts2 = ref._originalPoints.map(function(v) { return { x: v.x + dx2, y: v.y + dy2 }; });
                      var allValid2 = newOrigPts2.every(function(v) { return isFinite(v.x) && isFinite(v.y); });
                      if (allValid2) {
                        pl._originalPoints = newOrigPts2;
                      } else {
                        console.warn('_propagateSP: invalid points in translated _originalPoints for ' + pl.id);
                        pl._originalPoints = ref._originalPoints.map(function(v) { return {x: v.x, y: v.y}; });
                      }
                    }
                  }
                } else {
                  var dx = plBBC.cx - rcx;
                  var dy = plBBC.cy - rcy;
                  var newOrigPts = ref._originalPoints.map(function(v) { return { x: v.x + dx, y: v.y + dy }; });
                  var allValid = newOrigPts.every(function(v) { return isFinite(v.x) && isFinite(v.y); });
                  if (allValid) {
                    pl._originalPoints = newOrigPts;
                  } else {
                    var invalidCount = newOrigPts.filter(function(v) { return !isFinite(v.x) || !isFinite(v.y); }).length;
                    console.warn('_propagateSP: ' + invalidCount + ' invalid points in fallback _originalPoints for placement ' + pl.id);
                    pl._originalPoints = ref._originalPoints.map(function(v) { return {x: v.x, y: v.y}; });
                  }
                }
              }
            }
          }
        }
        // Propagate _sP (sub-part children) with same transform used for _originalPoints
        if (!pl._sP && !(pl.points && pl.points._sP)) {
          if (ref && ref._sP && Array.isArray(ref._sP) && ref._sP.length > 0) {
            // Re-derive plArray for min-corner (same as _originalPoints logic above)
            var plPts = pl.points;
            var plArr = Array.isArray(plPts) ? plPts :
              (plPts && plPts._sP && Array.isArray(plPts._sP) && plPts._sP.length > 0 && Array.isArray(plPts._sP[0]) && plPts._sP[0].length >= 2)
                ? plPts._sP[0] : null;
            if (plArr && plArr.length >= 2 && ref.points && Array.isArray(ref.points) && ref.points.length >= 2) {
              // bbox center of placement polygon
              var _plBBC = {mx:Infinity,nx:-Infinity,my:Infinity,ny:-Infinity};
              for (var _pi = 0; _pi < plArr.length; _pi++) {
                if (plArr[_pi].x < _plBBC.mx) _plBBC.mx = plArr[_pi].x;
                if (plArr[_pi].x > _plBBC.nx) _plBBC.nx = plArr[_pi].x;
                if (plArr[_pi].y < _plBBC.my) _plBBC.my = plArr[_pi].y;
                if (plArr[_pi].y > _plBBC.ny) _plBBC.ny = plArr[_pi].y;
              }
              var _refPts = ref.points;
              // bbox center of reference polygon
              var _rmx=Infinity,_rMx=-Infinity,_rmy=Infinity,_rMy=-Infinity;
              for (var _rpi=0;_rpi<_refPts.length;_rpi++){
                if(_refPts[_rpi].x<_rmx)_rmx=_refPts[_rpi].x; if(_refPts[_rpi].x>_rMx)_rMx=_refPts[_rpi].x;
                if(_refPts[_rpi].y<_rmy)_rmy=_refPts[_rpi].y; if(_refPts[_rpi].y>_rMy)_rMy=_refPts[_rpi].y;
              }
              var _rcx = (_rmx+_rMx)/2, _rcy = (_rmy+_rMy)/2;
              // Detect rotation (same as _originalPoints path above)
              var _refAngle = Math.atan2(_refPts[1].y - _refPts[0].y, _refPts[1].x - _refPts[0].x);
              var _plAngle = Math.atan2(plArr[1].y - plArr[0].y, plArr[1].x - plArr[0].x);
              var _rawDet = ((_plAngle - _refAngle) * 180 / Math.PI + 360) % 360;
              var _isMir = pl._m || false;
              var _actRot = _isMir ? ((180 - 2 * _refAngle * 180 / Math.PI - _rawDet) % 360 + 360) % 360 : _rawDet;
              // center alignment: place _sP children centered within offset boundary
              var _plCx = (_plBBC.mx + _plBBC.nx) / 2, _plCy = (_plBBC.my + _plBBC.ny) / 2;
              var _refCx = (_rmx + _rMx) / 2, _refCy = (_rmy + _rMy) / 2;
              // Apply same transform to each _sP child (min-corner aligned to offset polygon)
              pl._sP = [];
              var _refOffMinX = Infinity, _refOffMinY = Infinity;
              for (var _rpi = 0; _rpi < _refPts.length; _rpi++) {
                if (_refPts[_rpi].x < _refOffMinX) _refOffMinX = _refPts[_rpi].x;
                if (_refPts[_rpi].y < _refOffMinY) _refOffMinY = _refPts[_rpi].y;
              }
              var _plOffMinX = _plBBC.mx, _plOffMinY = _plBBC.my;
              var _spDx = _plOffMinX - _refOffMinX;
              var _spDy = _plOffMinY - _refOffMinY;
              for (var _si = 0; _si < ref._sP.length; _si++) {
                var _ch = ref._sP[_si];
                if (typeof _ch === 'object' && !Array.isArray(_ch)) continue;
                if (!_ch || !_ch.length) continue;
                var _mapped;
                if (Math.abs(_actRot) > 1e-4 && Math.abs(_actRot - 360) > 1e-4) {
                  var _rad = _actRot * Math.PI / 180;
                  var _cR = Math.cos(_rad), _sR = Math.sin(_rad);
                  var _transformed;
                  if (_isMir) {
                    _transformed = _ch.map(function(v) {
                      var _rx = -(v.x - _rcx) + _rcx, _ry = v.y;
                      var _dvx = _rx - _rcx, _dvy = _ry - _rcy;
                      return { x: _rcx + _dvx * _cR - _dvy * _sR, y: _rcy + _dvx * _sR + _dvy * _cR };
                    });
                  } else {
                    _transformed = _ch.map(function(v) {
                      var _dvx = v.x - _rcx, _dvy = v.y - _rcy;
                      return { x: _rcx + _dvx * _cR - _dvy * _sR, y: _rcy + _dvx * _sR + _dvy * _cR };
                    });
                  }
                  // Find center of transformed child
                  var _tMx = -Infinity, _tMnx = Infinity, _tMy = -Infinity, _tMny = Infinity;
                  for (var _ti = 0; _ti < _transformed.length; _ti++) {
                    if (_transformed[_ti].x < _tMnx) _tMnx = _transformed[_ti].x;
                    if (_transformed[_ti].x > _tMx) _tMx = _transformed[_ti].x;
                    if (_transformed[_ti].y < _tMny) _tMny = _transformed[_ti].y;
                    if (_transformed[_ti].y > _tMy) _tMy = _transformed[_ti].y;
                  }
                  var _tCx = (_tMnx + _tMx) / 2, _tCy = (_tMny + _tMy) / 2;
                  // Translate to match placement center
                  var _dx = _plCx - _tCx;
                  var _dy = _plCy - _tCy;
                  _mapped = _transformed.map(function(v) {
                    return { x: v.x + _dx, y: v.y + _dy };
                  });
                  // Validate all points are finite
                  var allValid = _mapped.every(function(v) { return isFinite(v.x) && isFinite(v.y); });
                  if (!allValid) {
                    var invalidCount = _mapped.filter(function(v) { return !isFinite(v.x) || !isFinite(v.y); }).length;
                    console.warn('_propagateSP: ' + invalidCount + ' invalid points in _sP child ' + _si + ' for placement ' + pl.id);
                    // Keep original child geometry as fallback
                    _mapped = _ch.map(function(v) { return {x: v.x, y: v.y}; });
                  }
                } else {
                  // No rotation: same translation as offset polygon
                  if (_isMir) {
                    _mapped = _ch.map(function(v) {
                      return { x: -(v.x - _rcx) + _rcx + _spDx, y: v.y + _spDy };
                    });
                  } else {
                    _mapped = _ch.map(function(v) {
                      return { x: v.x + _spDx, y: v.y + _spDy };
                    });
                  }
                  // Validate all points are finite
                  var allValid = _mapped.every(function(v) { return isFinite(v.x) && isFinite(v.y); });
                  if (!allValid) {
                    var invalidCount = _mapped.filter(function(v) { return !isFinite(v.x) || !isFinite(v.y); }).length;
                    console.warn('_propagateSP: ' + invalidCount + ' invalid points in _sP child ' + _si + ' (no rotation) for placement ' + pl.id);
                    // Keep original child geometry as fallback
                    _mapped = _ch.map(function(v) { return {x: v.x, y: v.y}; });
                  }
                }
                pl._sP.push(_mapped);
              }
            }
          }
        }
      });
    },

    _mapColors: function(placements, parts) {
      if (!placements || !parts) return;
      var colorById = {}, fileById = {}, pairMirrorById = {};
      parts.forEach(function(p) {
        if (p.origColor) {
          colorById[p.id] = p.origColor;
          if (p.sourceId != null) colorById[p.sourceId] = p.origColor;
        }
        if (p.sourceFileName) {
          fileById[p.id] = p.sourceFileName;
          if (p.sourceId != null) fileById[p.sourceId] = p.sourceFileName;
        }
        if (p._pairMirror) {
          pairMirrorById[p.id] = true;
        }
      });
      placements.forEach(function(pl) {
        if (!pl.origColor) pl.origColor = colorById[pl.sourceId] || colorById[pl.id] || '#6b7280';
        if (!pl.sourceFileName) pl.sourceFileName = fileById[pl.sourceId] || fileById[pl.id] || '';
        if (!pl._m && pairMirrorById[pl.id]) { pl._m = true; }
      });
    },

    renderToCanvas: function(canvas, placements, binWidthMm, options) {
      options = options || {};
      const ctx = canvas.getContext('2d');
      const w = canvas.width;
      const h = canvas.height;
      const bounds = MarkerNesting.computeBounds(placements);
      var inset = options.inset || 10;
      var zoom = options.zoom || 1;
      var panX = options.panX || 0;
      var panY = options.panY || 0;
      var margin = options.margin || 40;
      const markerLen = Math.max(bounds.maxY || binWidthMm, 2) + inset;

      const drawWidth = w - margin * 2;
      const drawHeight = h - margin * 2;
      const base = Math.min(drawWidth / binWidthMm, drawHeight / markerLen);
      const scale = base * zoom;
      var ox = (w - binWidthMm * scale) / 2 + panX;
      var oy = (h - markerLen * scale) / 2 + panY;

      function tx(x) { return ox + x * scale; }
      function ty(y) { return oy + y * scale; }
      function worldX(cx) { return (cx - ox) / scale; }
      function worldY(cy) { return (cy - oy) / scale; }

      var tr = { scale: scale, ox: ox, oy: oy, tx: tx, ty: ty, worldX: worldX, worldY: worldY, inset: inset, markerLen: markerLen, margin: margin };

      ctx.clearRect(0, 0, w, h);

      ctx.fillStyle = '#f8f8f8';
      ctx.fillRect(tx(0), ty(0), binWidthMm * scale, markerLen * scale);

      ctx.strokeStyle = '#999';
      ctx.lineWidth = 1;
      ctx.strokeRect(tx(0), ty(0), binWidthMm * scale, markerLen * scale);

      if (inset > 0) {
        ctx.strokeStyle = '#2196F3';
        ctx.lineWidth = 1.5;
        ctx.strokeRect(tx(inset), ty(inset), (binWidthMm - 2 * inset) * scale, (markerLen - 2 * inset) * scale);
        ctx.fillStyle = '#2196F3';
        ctx.font = '12px monospace';
        ctx.textAlign = 'left';
      }

      ctx.fillStyle = '#333';
      ctx.font = '13px monospace';
      ctx.textAlign = 'center';
      ctx.fillText('Width: ' + (binWidthMm / 1000).toFixed(3) + ' m', tx(binWidthMm / 2), ty(0) - 10);

      // Boundary validation happens before render — no clipping
      if (placements && placements.length > 0) {
        placements.forEach(function(p, idx) {
          var color = p.origColor || '#6b7280';
          var isMirror = p._m || p._pairMirror;

          var rawPts = p.points;
          if (!rawPts) return;
          if (p._unionOffset && p._sP && Array.isArray(p._sP)) {
            p._sP.forEach(function(sp) {
              if (!sp || sp.length < 3) return;
              ctx.strokeStyle = color;
              ctx.lineWidth = 0.65;
              ctx.beginPath();
              ctx.moveTo(tx(sp[0].x), ty(sp[0].y));
              for (var spi = 1; spi < sp.length; spi++) ctx.lineTo(tx(sp[spi].x), ty(sp[spi].y));
              ctx.closePath();
              ctx.stroke();
            });
            var offsetDrawPts = (p._childOffsets && Array.isArray(p._childOffsets) && p._childOffsets.length > 0) ? p._childOffsets : null;
            var uc = (color || '#6b7280').toLowerCase();
            var uCyanLike = uc === '#00bcd4' || uc === '#00bcd5' || uc === '#00b0c8' || uc === '#2196f3' || (uc.indexOf('cyan') >= 0) || (uc.indexOf('teal') >= 0);
            ctx.strokeStyle = uCyanLike ? '#FF7043' : '#00BCD4';
            ctx.lineWidth = 0.4;
            if (offsetDrawPts) {
              // Draw individual child offsets for accurate visual boundary
              offsetDrawPts.forEach(function(co) {
                if (!co || co.length < 3) return;
                ctx.beginPath();
                ctx.moveTo(tx(co[0].x), ty(co[0].y));
                for (var cpi = 1; cpi < co.length; cpi++) ctx.lineTo(tx(co[cpi].x), ty(co[cpi].y));
                ctx.closePath();
                ctx.stroke();
              });
            } else if (Array.isArray(rawPts) && rawPts.length >= 3) {
              ctx.beginPath();
              ctx.moveTo(tx(rawPts[0].x), ty(rawPts[0].y));
              for (var upi = 1; upi < rawPts.length; upi++) ctx.lineTo(tx(rawPts[upi].x), ty(rawPts[upi].y));
              ctx.closePath();
              ctx.stroke();
            }
            return;
          }
          if (!Array.isArray(rawPts)) {
            if (rawPts._sP && Array.isArray(rawPts._sP) && rawPts._sP.length > 0) {
              // Expand composite for rendering: draw each sub-part individually
              rawPts._sP.forEach(function(sp, si) {
                if (!sp || sp.length < 3) return;
                ctx.beginPath();
                ctx.moveTo(tx(sp[0].x), ty(sp[0].y));
                for (var sii = 1; sii < sp.length; sii++) ctx.lineTo(tx(sp[sii].x), ty(sp[sii].y));
                ctx.closePath();
                ctx.strokeStyle = color;
                ctx.lineWidth = 0.65;
                ctx.stroke();
              });
            }
            return;
          }
          if (rawPts.length < 3) return;
          for (var _vi = 0; _vi < rawPts.length; _vi++) {
            if (!isFinite(rawPts[_vi].x) || !isFinite(rawPts[_vi].y)) return;
          }

          // Outline: draw original polygon (before offset) when available, else the placed shape
          var outlinePts = (p._originalPoints && Array.isArray(p._originalPoints) && p._originalPoints.length >= 3) ? p._originalPoints : rawPts;
          if (outlinePts.length >= 3) {
            ctx.strokeStyle = color;
            ctx.lineWidth = 0.65;
            ctx.beginPath();
            ctx.moveTo(tx(outlinePts[0].x), ty(outlinePts[0].y));
            for (var i = 1; i < outlinePts.length; i++) ctx.lineTo(tx(outlinePts[i].x), ty(outlinePts[i].y));
            ctx.closePath();
            ctx.stroke();
          }

          // If part has offset, draw offset polygon with cyan solid line (or fallback if DXF color is cyan/teal)
          if (p._originalPoints) {
            // Detect if DXF outline color is cyan/teal-like — use fallback for offset
            var c = (color || '#6b7280').toLowerCase();
            var isCyanLike = c === '#00bcd4' || c === '#00bcd5' || c === '#00b0c8' || c === '#2196f3' || (c.indexOf('cyan') >= 0) || (c.indexOf('teal') >= 0);
            var offsetColor = isCyanLike ? '#FF7043' : '#00BCD4';
            ctx.strokeStyle = offsetColor;
            ctx.lineWidth = 0.4;
            ctx.beginPath();
            ctx.moveTo(tx(rawPts[0].x), ty(rawPts[0].y));
            for (var i = 1; i < rawPts.length; i++) ctx.lineTo(tx(rawPts[i].x), ty(rawPts[i].y));
            ctx.closePath();
            ctx.stroke();
          }
        });
      }

      var markerLenM = markerLen / 1000;
      ctx.strokeStyle = '#e00';
      ctx.lineWidth = 2;
      ctx.beginPath();
      ctx.moveTo(tx(0), ty(markerLen));
      ctx.lineTo(tx(binWidthMm), ty(markerLen));
      ctx.stroke();

      ctx.fillStyle = '#e00';
      ctx.font = 'bold 14px monospace';
      ctx.textAlign = 'right';
      ctx.fillText('Marker length: ' + markerLenM.toFixed(4) + ' m', tx(binWidthMm) - 6, ty(markerLen) + 20);

      return {
        markerLengthMm: markerLen,
        markerLengthM: markerLenM,
        scale: scale,
        inset: inset,
        ox: ox, oy: oy,
        tx: tx, ty: ty,
        worldX: worldX, worldY: worldY,
        margin: margin, panX: panX, panY: panY,
        zoom: zoom, binWidthMm: binWidthMm, markerLen: markerLen
      };
    },

    // ── Shared geometry extraction (preview + marker + PDF) ──
    getEntityVerts: function(entity, segsPerFull) {
      return entityToVerts(entity, segsPerFull || SEGMENTS_PER_FULL);
    }
  };
})();
