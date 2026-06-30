/**
 * PDF Vector Parser — handles pdf.js constructPath compound operator
 * CorelDRAW PDFs use OPS.constructPath (91) which bundles all path ops into one call
 */
class PDFVectorParser {

    async parse(arrayBuffer) {
        if (typeof pdfjsLib === 'undefined') throw new Error('pdf.js not loaded');
        pdfjsLib.GlobalWorkerOptions.workerSrc =
            'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';

        const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
        const entities = [];
        for (let p = 1; p <= pdf.numPages; p++) {
            const page = await pdf.getPage(p);
            const pageEnts = await this._parsePage(page, p);
            entities.push(...pageEnts);
        }
        return entities;
    }

    async _parsePage(page, pageNum) {
        const ops = await page.getOperatorList();
        const OPS = pdfjsLib.OPS;
        const { fnArray, argsArray } = ops;
        const entities = [];

        for (let i = 0; i < fnArray.length; i++) {
            const fn = fnArray[i];
            const a  = argsArray[i];

            // constructPath: args = [opsArray, coordsArray, minmax?]
            // opsArray entries: 13=moveTo 14=lineTo 15=curveTo 16=curveTo2 17=curveTo3 18=closePath
            if (fn === OPS.constructPath) {
                const pathOps  = a[0]; // array of sub-op codes
                const coords   = a[1]; // flat coord array
                const len = this._measureConstructPath(pathOps, coords);
                if (len > 0.01) {
                    entities.push({
                        type: 'PATH',
                        layer: `page-${pageNum}`,
                        pathSegments: [], // not needed, length pre-computed
                        _precomputedLength: len
                    });
                }
                continue;
            }

            // rectangle shorthand (rare in CorelDRAW but handle it)
            if (fn === OPS.rectangle) {
                const len = 2 * (Math.abs(a[2]) + Math.abs(a[3]));
                if (len > 0.01) {
                    entities.push({
                        type: 'RECT',
                        layer: `page-${pageNum}`,
                        x: a[0], y: a[1], width: a[2], height: a[3],
                        _precomputedLength: len
                    });
                }
            }
            // All other ops (stroke, fill, setColor, etc.) are ignored —
            // constructPath already contains the complete geometry
        }

        return entities;
    }

    _measureConstructPath(pathOps, coords) {
        // Sub-op codes inside constructPath args[0]:
        // 13 = moveTo    (2 coords: x,y)
        // 14 = lineTo    (2 coords: x,y)
        // 15 = curveTo   (6 coords: x1,y1,x2,y2,x,y)
        // 16 = curveTo2  (4 coords: x2,y2,x,y)  cp1=current
        // 17 = curveTo3  (4 coords: x1,y1,x,y)  cp2=end
        // 18 = closePath (0 coords)

        let total = 0;
        let ci = 0; // coord index
        let cx = 0, cy = 0, sx = 0, sy = 0;

        for (const op of pathOps) {
            switch (op) {
                case 13: { // moveTo
                    cx = coords[ci++]; cy = coords[ci++];
                    sx = cx; sy = cy;
                    break;
                }
                case 14: { // lineTo
                    const nx = coords[ci++], ny = coords[ci++];
                    const dx = nx-cx, dy = ny-cy;
                    total += Math.sqrt(dx*dx + dy*dy);
                    cx = nx; cy = ny;
                    break;
                }
                case 15: { // curveTo (6 coords)
                    const x1=coords[ci++], y1=coords[ci++];
                    const x2=coords[ci++], y2=coords[ci++];
                    const x =coords[ci++], y =coords[ci++];
                    total += this._bezier(cx,cy, x1,y1, x2,y2, x,y);
                    cx=x; cy=y;
                    break;
                }
                case 16: { // curveTo2 (4 coords, cp1=current)
                    const x2=coords[ci++], y2=coords[ci++];
                    const x =coords[ci++], y =coords[ci++];
                    total += this._bezier(cx,cy, cx,cy, x2,y2, x,y);
                    cx=x; cy=y;
                    break;
                }
                case 17: { // curveTo3 (4 coords, cp2=end)
                    const x1=coords[ci++], y1=coords[ci++];
                    const x =coords[ci++], y =coords[ci++];
                    total += this._bezier(cx,cy, x1,y1, x,y, x,y);
                    cx=x; cy=y;
                    break;
                }
                case 18: { // closePath
                    const dx=sx-cx, dy=sy-cy;
                    total += Math.sqrt(dx*dx + dy*dy);
                    cx=sx; cy=sy;
                    break;
                }
            }
        }
        return total;
    }

    _bezier(x0,y0, x1,y1, x2,y2, x3,y3, n=16) {
        let len=0, px=x0, py=y0;
        for (let i=1; i<=n; i++) {
            const t=i/n, mt=1-t;
            const nx = mt*mt*mt*x0 + 3*mt*mt*t*x1 + 3*mt*t*t*x2 + t*t*t*x3;
            const ny = mt*mt*mt*y0 + 3*mt*mt*t*y1 + 3*mt*t*t*y2 + t*t*t*y3;
            const dx=nx-px, dy=ny-py;
            len += Math.sqrt(dx*dx+dy*dy);
            px=nx; py=ny;
        }
        return len;
    }
}

if (typeof module !== 'undefined' && module.exports) module.exports = PDFVectorParser;
