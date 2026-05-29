const state = {
  fileName: "",
  entities: [],
  generated: [],
  learnedProfile: loadLearnedProfile(),
  trainingExamples: loadTrainingExamples(),
  trainingMeta: loadTrainingMeta(),
  trainingStatus: "Loading training data...",
  trainingSyncing: false,
  baseViewBox: null,
  viewBox: null,
};

const APP_CONFIG = window.PATCH_TEMPLATE_CONFIG || {};
const POWER_AUTOMATE_TRAINING_URL = APP_CONFIG.powerAutomateTrainingUrl || "";
const POWER_AUTOMATE_READ_URL = APP_CONFIG.powerAutomateReadUrl || "";
const DEFAULT_FORMULA = {
  boardWidth: 150,
  boardHeight: 180,
  cornerRadius: 12,
  slotWidth: 30,
  slotHeight: 11,
  offset: 7,
  templateNumber: "AMS0000",
  patchRelX: 0,
  patchRelY: -32.4,
  patchRotation: 0,
};

const INITIAL_TEMPLATE_NUMBER = getTemplateNumberFromQuery() || DEFAULT_FORMULA.templateNumber;
DEFAULT_FORMULA.templateNumber = INITIAL_TEMPLATE_NUMBER;

const els = {
  dxfInput: document.querySelector("#dxfInput"),
  clearButton: document.querySelector("#clearButton"),
  exportButton: document.querySelector("#exportButton"),
  applyButton: document.querySelector("#applyButton"),
  preview: document.querySelector("#preview"),
  fileSummary: document.querySelector("#fileSummary"),
  entityCount: document.querySelector("#entityCount"),
  layerCount: document.querySelector("#layerCount"),
  dxfSize: document.querySelector("#dxfSize"),
  dxfSizeSummary: document.querySelector("#dxfSizeSummary"),
  learnedCount: document.querySelector("#learnedCount"),
  learnedSummary: document.querySelector("#learnedSummary"),
  trainingCount: document.querySelector("#trainingCount"),
  trainingSummary: document.querySelector("#trainingSummary"),
  saveTrainingButton: document.querySelector("#saveTrainingButton"),
  exportTrainingButton: document.querySelector("#exportTrainingButton"),
  importTrainingButton: document.querySelector("#importTrainingButton"),
  trainingInput: document.querySelector("#trainingInput"),
  exportModal: document.querySelector("#exportModal"),
  exportModalBackdrop: document.querySelector("#exportModalBackdrop"),
  exportCancelButton: document.querySelector("#exportCancelButton"),
  exportOnlyButton: document.querySelector("#exportOnlyButton"),
  exportAndRecordButton: document.querySelector("#exportAndRecordButton"),
  issueList: document.querySelector("#issueList"),
  controls: document.querySelector(".controls"),
  previewPanel: document.querySelector(".preview-panel"),
  boardWidth: document.querySelector("#boardWidth"),
  boardHeight: document.querySelector("#boardHeight"),
  cornerRadius: document.querySelector("#cornerRadius"),
  slotWidth: document.querySelector("#slotWidth"),
  slotHeight: document.querySelector("#slotHeight"),
  offset: document.querySelector("#offset"),
  templateNumber: document.querySelector("#templateNumber"),
  patchRelX: document.querySelector("#patchRelX"),
  patchRelY: document.querySelector("#patchRelY"),
  patchRotation: document.querySelector("#patchRotation"),
  zoomInButton: document.querySelector("#zoomInButton"),
  zoomOutButton: document.querySelector("#zoomOutButton"),
  zoomResetButton: document.querySelector("#zoomResetButton"),
};

function getTemplateNumberFromQuery() {
  const params = new URLSearchParams(window.location.search);
  const value = params.get("templateNo") || params.get("templateNumber") || params.get("template");
  return value ? value.trim() : "";
}

if (els.templateNumber) {
  els.templateNumber.value = INITIAL_TEMPLATE_NUMBER;
}

if (state.learnedProfile) applyLearnedProfileToControls(state.learnedProfile);
render();
loadTrainingMasterJsonFromPowerAutomate();
syncArtboardHeightToParameter();

els.dxfInput.addEventListener("change", async (event) => {
  const file = event.target.files[0];
  if (!file) return;
  loadDxf(await file.text(), file.name);
});

els.clearButton.addEventListener("click", clearArtboard);

els.applyButton.addEventListener("click", () => {
  state.generated = generatePatchTemplate(state.entities, getFormula());
  els.exportButton.disabled = state.generated.length === 0;
  resetView();
  render();
});

els.exportButton.addEventListener("click", () => {
  openExportModal();
});

els.exportCancelButton.addEventListener("click", closeExportModal);
els.exportModalBackdrop.addEventListener("click", closeExportModal);
els.exportOnlyButton.addEventListener("click", () => {
  closeExportModal();
  exportCurrentDxf();
});
els.exportAndRecordButton.addEventListener("click", async () => {
  await recordSatisfiedExportExample();
  closeExportModal();
  exportCurrentDxf();
});

function exportCurrentDxf() {
  const exportEntities = state.generated.filter((entity) => !entity.previewOnly && !/SOURCE/i.test(entity.layer || ""));
  const dxf = buildOutputDxf(exportEntities, getFormula());
  const blob = new Blob([dxf], { type: "application/dxf" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = state.fileName.replace(/\.dxf$/i, "") + "-patch-template.dxf";
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

function clearArtboard() {
  state.fileName = "";
  state.entities = [];
  state.generated = [];
  state.baseViewBox = null;
  state.viewBox = null;
  els.dxfInput.value = "";
  resetParameters();
  els.applyButton.disabled = true;
  els.clearButton.disabled = true;
  els.exportButton.disabled = true;
  closeExportModal();
  render();
}

function resetParameters() {
  els.boardWidth.value = DEFAULT_FORMULA.boardWidth;
  els.boardHeight.value = DEFAULT_FORMULA.boardHeight;
  els.cornerRadius.value = DEFAULT_FORMULA.cornerRadius;
  els.slotWidth.value = DEFAULT_FORMULA.slotWidth;
  els.slotHeight.value = DEFAULT_FORMULA.slotHeight;
  els.offset.value = DEFAULT_FORMULA.offset;
  els.templateNumber.value = DEFAULT_FORMULA.templateNumber;
  els.patchRelX.value = DEFAULT_FORMULA.patchRelX;
  els.patchRelY.value = DEFAULT_FORMULA.patchRelY;
  els.patchRotation.value = DEFAULT_FORMULA.patchRotation;
}

function openExportModal() {
  if (!state.generated.length) return;
  els.exportModal.classList.remove("hidden");
}

function closeExportModal() {
  els.exportModal.classList.add("hidden");
}

function syncArtboardHeightToParameter() {
  if (!els.controls || !els.previewPanel || !window.ResizeObserver) return;
  const resize = () => {
    if (window.matchMedia("(max-width: 1180px)").matches) {
      els.previewPanel.style.removeProperty("--parameter-card-height");
      return;
    }
    els.previewPanel.style.setProperty("--parameter-card-height", `${els.controls.offsetHeight}px`);
  };
  new ResizeObserver(resize).observe(els.controls);
  window.addEventListener("resize", resize);
  resize();
}

if (els.saveTrainingButton) {
  els.saveTrainingButton.addEventListener("click", async () => {
    const example = buildTrainingExample();
    if (!example) return;
    state.trainingExamples.push(example);
    saveTrainingExamples(state.trainingExamples);
    updateTrainingMasterJson();
    await syncTrainingMasterJson();
    render();
  });
}

if (els.exportTrainingButton) {
  els.exportTrainingButton.addEventListener("click", () => {
    const blob = new Blob([JSON.stringify(trainingPayload(), null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = "patch-template-training-data.json";
    document.body.appendChild(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(url);
  });
}

if (els.importTrainingButton && els.trainingInput) {
  els.importTrainingButton.addEventListener("click", () => els.trainingInput.click());

  els.trainingInput.addEventListener("change", async (event) => {
    const file = event.target.files[0];
    if (!file) return;
    const imported = JSON.parse(await file.text());
    const examples = Array.isArray(imported) ? imported : imported.examples;
    if (!Array.isArray(examples)) return;
    state.trainingExamples = examples;
    saveTrainingExamples(state.trainingExamples);
    updateTrainingMasterJson();
    await syncTrainingMasterJson();
    render();
    event.target.value = "";
  });
}

els.zoomInButton.addEventListener("click", () => zoomPreview(0.8));
els.zoomOutButton.addEventListener("click", () => zoomPreview(1.25));
els.zoomResetButton.addEventListener("click", () => {
  state.viewBox = state.baseViewBox ? { ...state.baseViewBox } : null;
  renderPreview();
});

[els.boardWidth, els.boardHeight, els.cornerRadius, els.slotWidth, els.slotHeight, els.offset, els.templateNumber, els.patchRelX, els.patchRelY, els.patchRotation].forEach((el) => {
  el.addEventListener("input", () => {
    if (!state.entities.length) return;
    state.generated = generatePatchTemplate(state.entities, getFormula());
    els.exportButton.disabled = state.generated.length === 0;
    resetView();
    render();
  });
});

function loadDxf(text, fileName) {
  state.fileName = fileName;
  state.entities = parseDxf(text);
  const learned = inferReferenceProfile(state.entities, getFormula());
  if (learned) {
    state.learnedProfile = learned;
    saveLearnedProfile(learned);
    applyLearnedProfileToControls(learned);
  }
  state.generated = [];
  resetView();
  els.applyButton.disabled = state.entities.length === 0;
  els.clearButton.disabled = state.entities.length === 0;
  els.exportButton.disabled = true;
  render();
}

function getFormula() {
  const learned = state.learnedProfile || {};
  return {
    boardWidth: Number(els.boardWidth.value) || DEFAULT_FORMULA.boardWidth,
    boardHeight: Number(els.boardHeight.value) || DEFAULT_FORMULA.boardHeight,
    cornerRadius: Number(els.cornerRadius.value) || DEFAULT_FORMULA.cornerRadius,
    slotWidth: Number(els.slotWidth.value) || DEFAULT_FORMULA.slotWidth,
    slotHeight: Number(els.slotHeight.value) || DEFAULT_FORMULA.slotHeight,
    offset: Number(els.offset.value) || DEFAULT_FORMULA.offset,
    templateNumber: (els.templateNumber.value || DEFAULT_FORMULA.templateNumber).trim() || DEFAULT_FORMULA.templateNumber,
    patchRelX: Number(els.patchRelX.value) || DEFAULT_FORMULA.patchRelX,
    patchRelY: Number(els.patchRelY.value) || DEFAULT_FORMULA.patchRelY,
    patchRotation: Number(els.patchRotation.value) || DEFAULT_FORMULA.patchRotation,
    slotRelX: learned.slotRelX ?? 48,
    slotRelY: learned.slotRelY ?? 28.8,
  };
}

function parseDxf(text) {
  const pairs = text.replace(/\r/g, "").split("\n").map((line) => line.trim());
  const entities = [];
  for (let i = 0; i < pairs.length - 1; i += 2) {
    if (pairs[i] !== "0") continue;
    const value = pairs[i + 1];

    if (value === "LWPOLYLINE") {
      const blockPairs = readRawEntityPairs(pairs, i + 2);
      const layer = valueFor(blockPairs, "8") || "0";
      const closed = valueFor(blockPairs, "70") === "1";
      const points = [];
      for (let p = 0; p < blockPairs.length; p += 1) {
        if (blockPairs[p].code === "10") {
          const y = blockPairs.slice(p + 1, p + 5).find((pair) => pair.code === "20");
          if (y) points.push({ x: num(blockPairs[p].value), y: num(y.value) });
        }
      }
      if (validPoints(points)) entities.push({ type: "LWPOLYLINE", layer, points: closed ? cleanClosingPoint(points) : points, closed });
    }

    if (value === "POLYLINE") {
      const blockPairs = readRawEntityPairs(pairs, i + 2);
      const layer = valueFor(blockPairs, "8") || "0";
      const flags = Number(valueFor(blockPairs, "70")) || 0;
      const points = collectPolylineVertices(pairs, i + 2);
      const closed = Boolean(flags & 1);
      if (validPoints(points)) entities.push({ type: "LWPOLYLINE", layer, points: closed ? cleanClosingPoint(points) : points, closed });
    }

    if (value === "SPLINE") {
      const blockPairs = readRawEntityPairs(pairs, i + 2);
      const layer = valueFor(blockPairs, "8") || "0";
      const flags = Number(valueFor(blockPairs, "70")) || 0;
      const points = collectSplinePoints(blockPairs);
      const closed = Boolean(flags & 1);
      if (validPoints(points)) entities.push({ type: "SPLINE", layer, points: closed ? cleanClosingPoint(points) : points, closed });
    }

    if (value === "CIRCLE") {
      const block = readEntityBlock(pairs, i + 2);
      const points = circlePoints(num(block["10"]), num(block["20"]), num(block["40"]));
      if (validPoints(points)) entities.push({ type: "CIRCLE", layer: block.layer, points, closed: true });
    }

    if (value === "ARC") {
      const block = readEntityBlock(pairs, i + 2);
      const points = arcPoints(num(block["10"]), num(block["20"]), num(block["40"]), degToRad(num(block["50"])), degToRad(num(block["51"])));
      if (validPoints(points)) entities.push({ type: "ARC", layer: block.layer, points, closed: false });
    }

    if (value === "ELLIPSE") {
      const blockPairs = readRawEntityPairs(pairs, i + 2);
      const layer = valueFor(blockPairs, "8") || "0";
      const points = ellipsePoints(blockPairs);
      const closed = isFullEllipse(num(valueFor(blockPairs, "41")), num(valueFor(blockPairs, "42")));
      if (validPoints(points)) entities.push({ type: "ELLIPSE", layer, points: closed ? cleanClosingPoint(points) : points, closed });
    }

    if (value === "HATCH") {
      const blockPairs = readRawEntityPairs(pairs, i + 2);
      const layer = valueFor(blockPairs, "8") || "0";
      hatchBoundaryLoops(blockPairs).forEach((points) => {
        if (validPoints(points)) entities.push({ type: "HATCH", layer, points: cleanClosingPoint(points), closed: true });
      });
    }

    if (value === "LINE") {
      const block = readEntityBlock(pairs, i + 2);
      const entity = {
        type: "LINE",
        layer: block.layer,
        points: [
          { x: num(block["10"]), y: num(block["20"]) },
          { x: num(block["11"]), y: num(block["21"]) },
        ],
      };
      if (validPoints(entity.points)) entities.push(entity);
    }
  }
  return entities;
}

function readRawEntityPairs(pairs, start) {
  const result = [];
  for (let i = start; i < pairs.length - 1; i += 2) {
    if (pairs[i] === "0") break;
    result.push({ code: pairs[i], value: pairs[i + 1] });
  }
  return result;
}

function readEntityBlock(pairs, start) {
  const raw = readRawEntityPairs(pairs, start);
  const block = { layer: "0" };
  raw.forEach((pair) => {
    if (pair.code === "8") block.layer = pair.value;
    block[pair.code] = pair.value;
  });
  return block;
}

function collectPolylineVertices(pairs, start) {
  const points = [];
  for (let i = start; i < pairs.length - 1; i += 2) {
    if (pairs[i] === "0" && pairs[i + 1] === "SEQEND") break;
    if (pairs[i] !== "0" || pairs[i + 1] !== "VERTEX") continue;
    const raw = readRawEntityPairs(pairs, i + 2);
    points.push({ x: num(valueFor(raw, "10")), y: num(valueFor(raw, "20")) });
  }
  return cleanCurvePoints(points);
}

function valueFor(pairs, code) {
  return pairs.find((pair) => pair.code === code)?.value;
}

function collectSplinePoints(pairs) {
  const controlPoints = [];
  const fitPoints = [];
  const knots = [];
  const degree = Number(valueFor(pairs, "71")) || 3;
  for (let i = 0; i < pairs.length; i += 1) {
    if (pairs[i].code === "40") knots.push(num(pairs[i].value));
    if (pairs[i].code !== "10" && pairs[i].code !== "11") continue;
    const target = pairs[i].code === "10" ? controlPoints : fitPoints;
    const yCode = pairs[i].code === "10" ? "20" : "21";
    const y = pairs.slice(i + 1, i + 5).find((pair) => pair.code === yCode);
    if (y) target.push({ x: num(pairs[i].value), y: num(y.value) });
  }
  const sampled = sampleBSpline(controlPoints, knots, degree);
  if (sampled.length) return sampled;
  return cleanCurvePoints(fitPoints.length ? fitPoints : controlPoints);
}

function sampleBSpline(controlPoints, knots, degree) {
  if (controlPoints.length <= degree || knots.length < controlPoints.length + degree + 1) return [];
  const n = controlPoints.length - 1;
  const start = knots[degree];
  const end = knots[knots.length - degree - 1];
  if (!Number.isFinite(start) || !Number.isFinite(end) || end <= start) return [];
  const unique = [...new Set(knots.filter((knot) => knot >= start && knot <= end))].sort((a, b) => a - b);
  const points = [];
  for (let i = 0; i < unique.length - 1; i += 1) {
    const a = unique[i];
    const b = unique[i + 1];
    if (b <= a) continue;
    const samples = Math.max(12, Math.ceil((b - a) * 12));
    for (let step = 0; step < samples; step += 1) {
      points.push(deBoorPoint(controlPoints, knots, degree, n, a + ((b - a) * step) / samples));
    }
  }
  points.push(deBoorPoint(controlPoints, knots, degree, n, end));
  return cleanCurvePoints(points.filter((point) => Number.isFinite(point.x) && Number.isFinite(point.y)));
}

function deBoorPoint(controlPoints, knots, degree, n, u) {
  const span = findKnotSpan(knots, degree, n, u);
  const d = [];
  for (let j = 0; j <= degree; j += 1) {
    d[j] = { ...controlPoints[span - degree + j] };
  }
  for (let r = 1; r <= degree; r += 1) {
    for (let j = degree; j >= r; j -= 1) {
      const left = knots[span - degree + j];
      const right = knots[span + 1 + j - r];
      const alpha = right === left ? 0 : (u - left) / (right - left);
      d[j] = {
        x: (1 - alpha) * d[j - 1].x + alpha * d[j].x,
        y: (1 - alpha) * d[j - 1].y + alpha * d[j].y,
      };
    }
  }
  return d[degree];
}

function findKnotSpan(knots, degree, n, u) {
  if (u >= knots[n + 1]) return n;
  if (u <= knots[degree]) return degree;
  let low = degree;
  let high = n + 1;
  let mid = Math.floor((low + high) / 2);
  while (u < knots[mid] || u >= knots[mid + 1]) {
    if (u < knots[mid]) high = mid;
    else low = mid;
    mid = Math.floor((low + high) / 2);
  }
  return mid;
}

function num(value) {
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : 0;
}

function degToRad(value) {
  return (value * Math.PI) / 180;
}

function circlePoints(cx, cy, radius, segments = 96) {
  if (!radius) return [];
  return Array.from({ length: segments }, (_, index) => {
    const angle = (Math.PI * 2 * index) / segments;
    return { x: cx + Math.cos(angle) * radius, y: cy + Math.sin(angle) * radius };
  });
}

function arcPoints(cx, cy, radius, start, end) {
  if (!radius) return [];
  let sweep = end - start;
  if (sweep <= 0) sweep += Math.PI * 2;
  const segments = Math.max(12, Math.ceil((sweep / (Math.PI * 2)) * 96));
  return Array.from({ length: segments + 1 }, (_, index) => {
    const angle = start + (sweep * index) / segments;
    return { x: cx + Math.cos(angle) * radius, y: cy + Math.sin(angle) * radius };
  });
}

function ellipsePoints(pairs) {
  const cx = num(valueFor(pairs, "10"));
  const cy = num(valueFor(pairs, "20"));
  const majorX = num(valueFor(pairs, "11"));
  const majorY = num(valueFor(pairs, "21"));
  const ratio = num(valueFor(pairs, "40")) || 1;
  const start = num(valueFor(pairs, "41"));
  const end = num(valueFor(pairs, "42"));
  const majorLength = Math.hypot(majorX, majorY);
  if (!majorLength) return [];
  const ux = majorX / majorLength;
  const uy = majorY / majorLength;
  const vx = -uy;
  const vy = ux;
  let sweep = end - start;
  if (sweep <= 0) sweep += Math.PI * 2;
  const segments = Math.max(24, Math.ceil((sweep / (Math.PI * 2)) * 128));
  return Array.from({ length: isFullEllipse(start, end) ? segments : segments + 1 }, (_, index) => {
    const angle = start + (sweep * index) / segments;
    const major = Math.cos(angle) * majorLength;
    const minor = Math.sin(angle) * majorLength * ratio;
    return { x: cx + ux * major + vx * minor, y: cy + uy * major + vy * minor };
  });
}

function isFullEllipse(start, end) {
  return Math.abs(Math.abs(end - start) - Math.PI * 2) < 0.001;
}

function hatchBoundaryLoops(pairs) {
  const loops = [];
  let i = pairs.findIndex((pair) => pair.code === "91");
  if (i === -1) return loops;
  const loopCount = Math.max(1, num(pairs[i].value));
  i += 1;
  for (let loop = 0; loop < loopCount && i < pairs.length; loop += 1) {
    while (i < pairs.length && pairs[i].code !== "92") i += 1;
    if (i >= pairs.length) break;
    i += 1; // 92
    if (pairs[i]?.code !== "93") {
      const fallback = [];
      while (i < pairs.length && pairs[i].code !== "97" && pairs[i].code !== "75" && pairs[i].code !== "92") {
        if (pairs[i].code === "10") {
          const y = pairs.slice(i + 1, i + 5).find((pair) => pair.code === "20");
          if (y) fallback.push({ x: num(pairs[i].value), y: num(y.value) });
        }
        i += 1;
      }
      const cleanedFallback = cleanCurvePoints(fallback).filter((point) => Math.abs(point.x) > 0.001 || Math.abs(point.y) > 0.001);
      if (cleanedFallback.length >= 3) loops.push(cleanedFallback);
      continue;
    }
    const edgeCount = Math.max(1, num(pairs[i].value));
    i += 1;
    const points = [];
    for (let edge = 0; edge < edgeCount && i < pairs.length; edge += 1) {
      while (i < pairs.length && pairs[i].code !== "72" && pairs[i].code !== "97" && pairs[i].code !== "75") i += 1;
      if (pairs[i]?.code !== "72") break;
      const parsed = parseHatchEdge(pairs, i);
      if (!parsed) break;
      points.push(...(points.length ? parsed.points.slice(1) : parsed.points));
      i = parsed.next;
    }
    const cleaned = cleanCurvePoints(points).filter((point) => Math.abs(point.x) > 0.001 || Math.abs(point.y) > 0.001);
    if (cleaned.length >= 3) loops.push(cleaned);
  }
  return loops;
}

function parseHatchEdge(pairs, start) {
  const edgeType = num(pairs[start].value);
  let i = start + 1;
  if (edgeType === 1) {
    const x1 = num(pairs[i]?.value);
    const y1 = num(pairs[i + 1]?.value);
    const x2 = num(pairs[i + 2]?.value);
    const y2 = num(pairs[i + 3]?.value);
    return { points: [{ x: x1, y: y1 }, { x: x2, y: y2 }], next: i + 4 };
  }
  if (edgeType === 2) {
    const cx = num(valueFor(pairs.slice(i, i + 12), "10"));
    const cy = num(valueFor(pairs.slice(i, i + 12), "20"));
    const radius = num(valueFor(pairs.slice(i, i + 12), "40"));
    const startAngle = degToRad(num(valueFor(pairs.slice(i, i + 12), "50")));
    const endAngle = degToRad(num(valueFor(pairs.slice(i, i + 12), "51")));
    while (i < pairs.length && !["72", "97", "75"].includes(pairs[i].code)) i += 1;
    return { points: arcPoints(cx, cy, radius, startAngle, endAngle), next: i };
  }
  if (edgeType === 3) {
    const block = [];
    while (i < pairs.length && !["72", "97", "75"].includes(pairs[i].code)) block.push(pairs[i++]);
    return { points: ellipsePoints(block), next: i };
  }
  if (edgeType === 4) {
    const degree = num(valueFor(pairs.slice(i, i + 10), "94")) || 3;
    while (i < pairs.length && pairs[i].code !== "95") i += 1;
    const knotCount = num(pairs[i]?.value);
    i += 1;
    while (i < pairs.length && pairs[i].code !== "96") i += 1;
    const controlCount = num(pairs[i]?.value);
    i += 1;
    const knots = [];
    while (i < pairs.length && knots.length < knotCount) {
      if (pairs[i].code === "40") knots.push(num(pairs[i].value));
      i += 1;
    }
    const controlPoints = [];
    while (i < pairs.length && controlPoints.length < controlCount) {
      if (pairs[i].code === "10") {
        const y = pairs.slice(i + 1, i + 5).find((pair) => pair.code === "20");
        if (y) controlPoints.push({ x: num(pairs[i].value), y: num(y.value) });
      }
      i += 1;
    }
    const points = sampleBSpline(controlPoints, knots, degree);
    while (i < pairs.length && !["72", "97", "75"].includes(pairs[i].code)) i += 1;
    return { points: points.length ? points : controlPoints, next: i };
  }
  return null;
}

function validPoints(points) {
  return points.length >= 2 && points.every((point) => Number.isFinite(point.x) && Number.isFinite(point.y));
}

function isShapeEntity(entity) {
  return ["LWPOLYLINE", "SPLINE", "CIRCLE", "ARC", "ELLIPSE", "HATCH"].includes(entity.type);
}

function generatePatchTemplate(entities, formula) {
  const outline = selectPatchOutline(entities, formula);
  if (!outline) return [];
  const basePatch = rotatePoints(centerPoints(outline.points, 0, 0), formula.patchRotation);
  const startX = 70;
  const gap = formula.boardWidth + 16;
  const startY = 0;
  const generated = [];
  const sourcePreviewY = startY + formula.boardHeight / 2 + 55;

  generated.push({
    type: "TEXT",
    layer: "AI_LABEL",
    text: "Patch DXF",
    point: { x: startX - formula.boardWidth / 2, y: sourcePreviewY + 18 },
    previewOnly: true,
  });
  generated.push({
    type: "LWPOLYLINE",
    layer: "AI_SOURCE_PATCH",
    points: movePoints(basePatch, startX, sourcePreviewY),
    closed: true,
    previewOnly: true,
  });

  for (let layer = 1; layer <= 4; layer += 1) {
    const origin = { x: startX + (layer - 1) * gap, y: startY };
    const board = templateBoard(origin, formula);
    const patch = movePoints(basePatch, origin.x + formula.patchRelX, origin.y + formula.patchRelY);
    const offsetPatch = offsetOrthogonalPolygon(patch, formula.offset, "outward");
    const slotCenter = needleSlotCenter(origin, formula);
    generated.push(
      { type: "TEXT", layer: "AI_LABEL", text: String(layer), point: { x: origin.x, y: origin.y + formula.boardHeight / 2 + 18 }, height: 8 },
      { type: "LWPOLYLINE", layer: `AI_LAYER_${layer}_BOARD`, points: board, closed: true },
      { type: "LWPOLYLINE", layer: `AI_LAYER_${layer}_PATCH`, points: patch, closed: true },
      { type: "LWPOLYLINE", layer: `AI_LAYER_${layer}_NEEDLE_SLOT`, points: pillPoints(slotCenter.x, slotCenter.y, formula.slotWidth, formula.slotHeight), closed: true },
    );
    if (layer === 1) {
      generated.push({
        type: "TEXT",
        layer: "AI_LAYER_1_TEMPLATE_NUMBER",
        text: formula.templateNumber,
        point: { x: (origin.x - formula.boardWidth / 2 + slotCenter.x - formula.slotWidth / 2) / 2, y: slotCenter.y },
        height: 3.5,
        anchor: "middle",
      });
    }
    if (layer >= 2 && offsetPatch.length >= 3) {
      generated.push({ type: "LWPOLYLINE", layer: `AI_LAYER_${layer}_PATCH_OFFSET_7MM`, points: offsetPatch, closed: true });
    }
  }
  return generated;
}

function selectPatchOutline(entities, formula) {
  const candidates = entities
    .filter((entity) => isShapeEntity(entity) && entity.points.length >= 3)
    .map((entity) => ({ entity, box: bounds([entity]), area: polygonArea(entity.points) }))
    .filter((item) => item.area > 100)
    .filter((item) => !isBoardSized(item.box, formula))
    .sort((a, b) => b.area - a.area);
  const unique = uniqueShapeCandidates(candidates);
  const hatchUnion = combineOverlappingHatches(unique);
  if (hatchUnion) return hatchUnion;
  const joined = joinConnectedPatchCandidates(unique);
  if (joined) return joined;
  return unique.filter((item) => !isNeedleSlotSized(item.box, formula))[0]?.entity || largestClosedShape(entities);
}

function uniqueShapeCandidates(candidates) {
  const seen = new Set();
  return candidates.filter((candidate) => {
    const box = candidate.box;
    const key = [
      Math.round(box.minX * 1000),
      Math.round(box.minY * 1000),
      Math.round(box.maxX * 1000),
      Math.round(box.maxY * 1000),
      Math.round(candidate.area * 1000),
      candidate.entity.points.length,
    ].join(":");
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

function combineOverlappingHatches(candidates) {
  const hatches = candidates.filter((candidate) => candidate.entity.type === "HATCH");
  if (hatches.length < 2) return null;
  const cluster = [hatches[0]];
  for (let i = 1; i < hatches.length; i += 1) {
    if (cluster.some((candidate) => boxesOverlap(candidate.box, hatches[i].box))) cluster.push(hatches[i]);
  }
  if (cluster.length < 2) return null;
  const union = polygonUnionBoundary(cluster.map((candidate) => candidate.entity.points));
  if (!union.length) return null;
  return { type: "LWPOLYLINE", layer: "AI_HATCH_UNION_PATCH_SOURCE", points: union, closed: true };
}

function boxesOverlap(a, b) {
  return a.minX <= b.maxX && a.maxX >= b.minX && a.minY <= b.maxY && a.maxY >= b.minY;
}

function polygonUnionBoundary(polygons) {
  const segments = [];
  polygons.forEach((polygon, polygonIndex) => {
    for (let i = 0; i < polygon.length; i += 1) {
      const a = polygon[i];
      const b = polygon[(i + 1) % polygon.length];
      const cuts = [{ t: 0, point: a }, { t: 1, point: b }];
      polygons.forEach((other, otherIndex) => {
        if (otherIndex === polygonIndex) return;
        for (let j = 0; j < other.length; j += 1) {
          const hit = segmentIntersectionParam(a, b, other[j], other[(j + 1) % other.length]);
          if (hit && hit.t > 0.0001 && hit.t < 0.9999) cuts.push({ t: hit.t, point: hit.point });
        }
      });
      cuts.sort((left, right) => left.t - right.t);
      for (let c = 0; c < cuts.length - 1; c += 1) {
        const p1 = cuts[c].point;
        const p2 = cuts[c + 1].point;
        if (Math.hypot(p2.x - p1.x, p2.y - p1.y) < 0.01) continue;
        const mid = { x: (p1.x + p2.x) / 2, y: (p1.y + p2.y) / 2 };
        const insideOther = polygons.some((other, otherIndex) => otherIndex !== polygonIndex && pointInPolygon(mid, other));
        if (!insideOther) segments.push({ p1, p2 });
      }
    }
  });
  return largestSegmentLoop(segments);
}

function segmentIntersectionParam(a1, a2, b1, b2) {
  const dax = a2.x - a1.x;
  const day = a2.y - a1.y;
  const dbx = b2.x - b1.x;
  const dby = b2.y - b1.y;
  const denominator = dax * dby - day * dbx;
  if (Math.abs(denominator) < 1e-9) return null;
  const t = ((b1.x - a1.x) * dby - (b1.y - a1.y) * dbx) / denominator;
  const u = ((b1.x - a1.x) * day - (b1.y - a1.y) * dax) / denominator;
  if (t < -0.0001 || t > 1.0001 || u < -0.0001 || u > 1.0001) return null;
  return { t, u, point: { x: a1.x + t * dax, y: a1.y + t * day } };
}

function pointInPolygon(point, polygon) {
  let inside = false;
  for (let i = 0, j = polygon.length - 1; i < polygon.length; j = i, i += 1) {
    const pi = polygon[i];
    const pj = polygon[j];
    const crosses = pi.y > point.y !== pj.y > point.y;
    if (crosses && point.x < ((pj.x - pi.x) * (point.y - pi.y)) / (pj.y - pi.y) + pi.x) inside = !inside;
  }
  return inside;
}

function largestSegmentLoop(segments) {
  const loops = [];
  const remaining = [...segments];
  while (remaining.length) {
    let current = remaining.shift();
    const loop = [current.p1, current.p2];
    let changed = true;
    while (changed && !samePoint(loop[0], loop[loop.length - 1])) {
      changed = false;
      const last = loop[loop.length - 1];
      for (let i = 0; i < remaining.length; i += 1) {
        const candidate = remaining[i];
        if (samePoint(last, candidate.p1)) {
          loop.push(candidate.p2);
        } else if (samePoint(last, candidate.p2)) {
          loop.push(candidate.p1);
        } else {
          continue;
        }
        remaining.splice(i, 1);
        changed = true;
        break;
      }
    }
    if (loop.length >= 4) loops.push(cleanClosingPoint(loop));
  }
  return loops.sort((a, b) => polygonArea(b) - polygonArea(a))[0] || [];
}

function joinConnectedPatchCandidates(candidates) {
  if (candidates.length < 2) return null;
  const remaining = candidates.map((candidate) => candidate.entity);
  let points = [...remaining.shift().points];
  let changed = true;
  while (changed) {
    changed = false;
    for (let i = 0; i < remaining.length; i += 1) {
      const candidate = remaining[i].points;
      const first = points[0];
      const last = points[points.length - 1];
      const cFirst = candidate[0];
      const cLast = candidate[candidate.length - 1];
      if (samePoint(last, cFirst)) {
        points = points.concat(candidate.slice(1));
      } else if (samePoint(last, cLast)) {
        points = points.concat([...candidate].reverse().slice(1));
      } else if (samePoint(first, cLast)) {
        points = candidate.slice(0, -1).concat(points);
      } else if (samePoint(first, cFirst)) {
        points = [...candidate].reverse().slice(0, -1).concat(points);
      } else {
        continue;
      }
      remaining.splice(i, 1);
      changed = true;
      break;
    }
  }
  if (points.length < 4 || !samePoint(points[0], points[points.length - 1])) return null;
  return { type: "LWPOLYLINE", layer: "AI_JOINED_PATCH_SOURCE", points: cleanClosingPoint(points), closed: true };
}

function samePoint(a, b) {
  return Math.hypot(a.x - b.x, a.y - b.y) <= 0.01;
}

function cleanClosingPoint(points) {
  if (points.length > 1 && samePoint(points[0], points[points.length - 1])) return points.slice(0, -1);
  return points;
}

function largestClosedShape(entities) {
  return entities
    .filter((entity) => isShapeEntity(entity) && entity.points.length >= 3)
    .sort((a, b) => polygonArea(b.points) - polygonArea(a.points))[0];
}

function isBoardSized(box, formula) {
  const width = box.maxX - box.minX;
  const height = box.maxY - box.minY;
  return Math.abs(width - formula.boardWidth) < 3 && Math.abs(height - formula.boardHeight) < 3;
}

function isNeedleSlotSized(box, formula) {
  const width = box.maxX - box.minX;
  const height = box.maxY - box.minY;
  const direct = Math.abs(width - formula.slotWidth) < 3 && Math.abs(height - formula.slotHeight) < 3;
  const rotated = Math.abs(width - formula.slotHeight) < 3 && Math.abs(height - formula.slotWidth) < 3;
  return direct || rotated;
}

function inferReferenceProfile(entities, fallback) {
  const shapes = entities
    .filter((entity) => isShapeEntity(entity) && entity.points.length >= 3)
    .map((entity) => {
      const box = bounds([entity]);
      return {
        entity,
        box,
        area: polygonArea(entity.points),
        width: box.maxX - box.minX,
        height: box.maxY - box.minY,
        cx: (box.minX + box.maxX) / 2,
        cy: (box.minY + box.maxY) / 2,
      };
    })
    .filter((shape) => shape.area > 100);
  const boards = shapes.filter((shape) => isBoardSized(shape.box, fallback));
  if (boards.length < 2) return null;

  const slots = shapes.filter((shape) => isNeedleSlotSized(shape.box, fallback));
  const patches = shapes
    .filter((shape) => !isBoardSized(shape.box, fallback))
    .filter((shape) => !isNeedleSlotSized(shape.box, fallback))
    .filter((shape) => shape.width < fallback.boardWidth * 0.85 && shape.height < fallback.boardHeight * 0.7)
    .sort((a, b) => b.area - a.area);
  if (!slots.length || !patches.length) return null;

  const samples = boards
    .map((board) => {
      const inside = (shape) =>
        shape.cx >= board.box.minX &&
        shape.cx <= board.box.maxX &&
        shape.cy >= board.box.minY &&
        shape.cy <= board.box.maxY;
      const slot = slots.find(inside);
      const patch = patches.find(inside);
      if (!slot || !patch) return null;
      return {
        boardWidth: board.width,
        boardHeight: board.height,
        slotWidth: Math.max(slot.width, slot.height),
        slotHeight: Math.min(slot.width, slot.height),
        slotRelX: slot.cx - board.cx,
        slotRelY: slot.cy - board.cy,
        patchRelX: patch.cx - board.cx,
        patchRelY: patch.cy - board.cy,
      };
    })
    .filter(Boolean);
  if (!samples.length) return null;

  return {
    boardWidth: round(avg(samples.map((sample) => sample.boardWidth))),
    boardHeight: round(avg(samples.map((sample) => sample.boardHeight))),
    slotWidth: round(avg(samples.map((sample) => sample.slotWidth))),
    slotHeight: round(avg(samples.map((sample) => sample.slotHeight))),
    slotRelX: round(avg(samples.map((sample) => sample.slotRelX))),
    slotRelY: round(avg(samples.map((sample) => sample.slotRelY))),
    patchRelX: round(avg(samples.map((sample) => sample.patchRelX))),
    patchRelY: round(avg(samples.map((sample) => sample.patchRelY))),
    sampleCount: samples.length,
  };
}

function avg(values) {
  return values.reduce((sum, value) => sum + value, 0) / values.length;
}

function round(value) {
  return Math.round(value * 1000) / 1000;
}

function loadLearnedProfile() {
  try {
    return JSON.parse(localStorage.getItem("patchTemplateLearnedProfile") || "null");
  } catch {
    return null;
  }
}

function saveLearnedProfile(profile) {
  try {
    localStorage.setItem("patchTemplateLearnedProfile", JSON.stringify(profile));
  } catch {
    // localStorage is unavailable in some automated checks.
  }
}

function loadTrainingExamples() {
  try {
    return JSON.parse(localStorage.getItem("patchTemplateTrainingExamples") || "[]");
  } catch {
    return [];
  }
}

function saveTrainingExamples(examples) {
  try {
    localStorage.setItem("patchTemplateTrainingExamples", JSON.stringify(examples));
  } catch {
    // localStorage is unavailable in some automated checks.
  }
}

function loadTrainingMeta() {
  try {
    return JSON.parse(localStorage.getItem("patchTemplateTrainingMeta") || "{\"revision\":0,\"updatedAt\":\"\"}");
  } catch {
    return { revision: 0, updatedAt: "" };
  }
}

function saveTrainingMeta(meta) {
  try {
    localStorage.setItem("patchTemplateTrainingMeta", JSON.stringify(meta));
  } catch {
    // localStorage is unavailable in some automated checks.
  }
}

async function recordSatisfiedExportExample() {
  if (!state.generated.length) return;
  const example = buildTrainingExample();
  if (!example) return;
  example.feedback = {
    satisfied: true,
    recordedAtExport: true,
    recordedAt: new Date().toISOString(),
  };
  state.trainingExamples.push(example);
  saveTrainingExamples(state.trainingExamples);
  updateTrainingMasterJson();
  await syncTrainingMasterJson();
  render();
}

function buildTrainingExample() {
  if (!state.entities.length) return null;
  const formula = getFormula();
  const outline = selectPatchOutline(state.entities, formula);
  if (!outline) return null;
  const outlineBox = bounds([outline]);
  const types = state.entities.reduce((acc, entity) => {
    acc[entity.type] = (acc[entity.type] || 0) + 1;
    return acc;
  }, {});
  return {
    schema: "patch-template-example-v1",
    savedAt: new Date().toISOString(),
    input: {
      fileName: state.fileName,
      entityCount: state.entities.length,
      entityTypes: types,
    },
    selectedPatch: {
      layer: outline.layer,
      type: outline.type,
      pointCount: outline.points.length,
      widthMm: round(outlineBox.maxX - outlineBox.minX),
      heightMm: round(outlineBox.maxY - outlineBox.minY),
      areaMm2: round(polygonArea(outline.points)),
    },
    formula: {
      boardWidthMm: formula.boardWidth,
      boardHeightMm: formula.boardHeight,
      cornerRadiusMm: formula.cornerRadius,
      needleSlotWidthMm: formula.slotWidth,
      needleSlotHeightMm: formula.slotHeight,
      outwardOffsetMm: formula.offset,
      patchRelX: formula.patchRelX,
      patchRelY: formula.patchRelY,
      patchRotationDeg: formula.patchRotation,
      templateNumber: formula.templateNumber,
    },
    output: {
      templateLayers: 4,
      generatedEntityCount: state.generated.length || generatePatchTemplate(state.entities, formula).length,
      ruleSet: "patch-template-4-layer-v1",
    },
  };
}

function trainingPayload() {
  const updatedAt = state.trainingMeta.updatedAt || new Date().toISOString();
  return {
    schema: "patch-template-ai-training-v1",
    revision: state.trainingMeta.revision || 0,
    updatedAt,
    exportedAt: updatedAt,
    exampleCount: state.trainingExamples.length,
    source: "patch-template-web",
    examples: state.trainingExamples,
    learnedProfile: state.learnedProfile,
  };
}

function updateTrainingMasterJson() {
  state.trainingMeta = {
    revision: (state.trainingMeta.revision || 0) + 1,
    updatedAt: new Date().toISOString(),
  };
  saveTrainingMeta(state.trainingMeta);
  try {
    localStorage.setItem("patchTemplateTrainingMasterJson", JSON.stringify(trainingPayload(), null, 2));
  } catch {
    // localStorage is unavailable in some automated checks.
  }
}

async function syncTrainingMasterJson() {
  if (!POWER_AUTOMATE_TRAINING_URL) {
    state.trainingStatus = "OneDrive save URL is not configured.";
    render();
    return false;
  }
  state.trainingSyncing = true;
  state.trainingStatus = "Uploading training data to OneDrive...";
  render();
  try {
    const response = await fetch(POWER_AUTOMATE_TRAINING_URL, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(trainingPayload()),
    });
    if (!response.ok) {
      const message = await response.text();
      throw new Error(message || `HTTP ${response.status}`);
    }
    state.trainingStatus = "Training data uploaded. Reloading OneDrive copy...";
    state.trainingSyncing = false;
    render();
    await loadTrainingMasterJsonFromPowerAutomate({ keepStatusOnSuccess: false });
    return true;
  } catch {
    state.trainingSyncing = false;
    state.trainingStatus = "Upload failed. Local training data is still kept in this browser.";
    render();
    return false;
  }
}

async function loadTrainingMasterJsonFromPowerAutomate(options = {}) {
  if (!POWER_AUTOMATE_READ_URL) {
    state.trainingStatus = "OneDrive read URL is not configured.";
    render();
    return false;
  }
  state.trainingSyncing = true;
  state.trainingStatus = "Loading training data from OneDrive...";
  render();
  try {
    const response = await fetch(POWER_AUTOMATE_READ_URL, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ request: "patch-template-training-data" }),
    });
    if (!response.ok) {
      const message = await response.text();
      throw new Error(message || `HTTP ${response.status}`);
    }
    const payload = normalizeTrainingPayload(await response.json());
    if (!payload || !Array.isArray(payload.examples)) {
      throw new Error("Training JSON does not contain an examples array.");
    }
    state.trainingExamples = payload.examples;
    state.trainingMeta = {
      revision: Number(payload.revision) || state.trainingMeta.revision || 0,
      updatedAt: payload.updatedAt || payload.exportedAt || state.trainingMeta.updatedAt || "",
    };
    state.learnedProfile = payload.learnedProfile || state.learnedProfile;
    saveTrainingExamples(state.trainingExamples);
    saveTrainingMeta(state.trainingMeta);
    saveLearnedProfile(state.learnedProfile);
    state.trainingSyncing = false;
    if (!options.keepStatusOnSuccess) {
      const updated = state.trainingMeta.updatedAt ? ` · ${formatTrainingDate(state.trainingMeta.updatedAt)}` : "";
      state.trainingStatus = `Loaded from OneDrive${updated}.`;
    }
    render();
    return true;
  } catch {
    state.trainingSyncing = false;
    state.trainingStatus = "Load failed. Showing local browser training data.";
    render();
    return false;
  }
}

function normalizeTrainingPayload(responseBody) {
  if (!responseBody) return null;
  if (typeof responseBody === "string") {
    try {
      return JSON.parse(responseBody);
    } catch {
      return null;
    }
  }
  if (Array.isArray(responseBody.examples)) return responseBody;
  const possibleContent = responseBody.body || responseBody.content || responseBody.fileContent || responseBody.value;
  if (typeof possibleContent === "string") {
    try {
      return JSON.parse(possibleContent);
    } catch {
      return null;
    }
  }
  if (possibleContent && Array.isArray(possibleContent.examples)) return possibleContent;
  return null;
}

function formatTrainingDate(value) {
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return String(value);
  return date.toLocaleString("en-MY", {
    day: "2-digit",
    month: "2-digit",
    year: "numeric",
    hour: "2-digit",
    minute: "2-digit",
  });
}

function applyLearnedProfileToControls(profile) {
  els.boardWidth.value = profile.boardWidth;
  els.boardHeight.value = profile.boardHeight;
  els.slotWidth.value = profile.slotWidth;
  els.slotHeight.value = profile.slotHeight;
  els.patchRelX.value = profile.patchRelX ?? 0;
  els.patchRelY.value = profile.patchRelY ?? -32.4;
  els.patchRotation.value = profile.patchRotation ?? 0;
}

function templateBoard(origin, formula) {
  const w = formula.boardWidth;
  const h = formula.boardHeight;
  const r = Math.min(formula.cornerRadius, w / 3, h / 3);
  const x = origin.x;
  const y = origin.y;
  const left = x - w / 2;
  const right = x + w / 2;
  const top = y + h / 2;
  const bottom = y - h / 2;
  const points = [
    { x: left, y: top },
    { x: right, y: top },
    { x: right, y: bottom + r },
  ];
  for (let i = 0; i <= 8; i += 1) {
    const angle = 0 - (Math.PI / 2) * (i / 8);
    points.push({ x: right - r + Math.cos(angle) * r, y: bottom + r + Math.sin(angle) * r });
  }
  points.push({ x: left + r, y: bottom });
  for (let i = 0; i <= 8; i += 1) {
    const angle = -Math.PI / 2 - (Math.PI / 2) * (i / 8);
    points.push({ x: left + r + Math.cos(angle) * r, y: bottom + r + Math.sin(angle) * r });
  }
  return points;
}

function pillPoints(cx, cy, width, height) {
  const radius = height / 2;
  const left = cx - width / 2 + radius;
  const right = cx + width / 2 - radius;
  const points = [];
  for (let i = 0; i <= 10; i += 1) {
    const angle = Math.PI / 2 - Math.PI * (i / 10);
    points.push({ x: right + Math.cos(angle) * radius, y: cy + Math.sin(angle) * radius });
  }
  for (let i = 0; i <= 10; i += 1) {
    const angle = -Math.PI / 2 - Math.PI * (i / 10);
    points.push({ x: left + Math.cos(angle) * radius, y: cy + Math.sin(angle) * radius });
  }
  return points;
}

function needleSlotCenter(origin, formula) {
  return {
    x: origin.x + formula.boardWidth / 2 - 11.5 - formula.slotWidth / 2,
    y: origin.y + formula.boardHeight / 2 - 50 - formula.slotHeight / 2,
  };
}

function offsetOrthogonalPolygon(points, amount, direction = "inward") {
  const circularOffset = offsetCircularShape(points, amount, direction);
  if (circularOffset) return circularOffset;
  const area = signedArea(points);
  const inwardSign = area >= 0 ? 1 : -1;
  const directionSign = direction === "outward" ? -1 : 1;
  const lines = [];
  for (let i = 0; i < points.length; i += 1) {
    const p1 = points[i];
    const p2 = points[(i + 1) % points.length];
    const dx = p2.x - p1.x;
    const dy = p2.y - p1.y;
    const len = Math.hypot(dx, dy) || 1;
    const nx = (-dy / len) * inwardSign * directionSign;
    const ny = (dx / len) * inwardSign * directionSign;
    lines.push({
      p1: { x: p1.x + nx * amount, y: p1.y + ny * amount },
      p2: { x: p2.x + nx * amount, y: p2.y + ny * amount },
    });
  }
  const result = [];
  for (let i = 0; i < lines.length; i += 1) {
    const prev = lines[(i - 1 + lines.length) % lines.length];
    const curr = lines[i];
    result.push(lineIntersection(prev.p1, prev.p2, curr.p1, curr.p2) || curr.p1);
  }
  return cleanSelfIntersectingOffset(result.filter((point) => Number.isFinite(point.x) && Number.isFinite(point.y)));
}

function offsetCircularShape(points, amount, direction) {
  if (points.length < 16) return null;
  const box = bounds([{ points }]);
  const width = box.maxX - box.minX;
  const height = box.maxY - box.minY;
  if (width <= 0 || height <= 0) return null;
  const ratio = Math.max(width, height) / Math.min(width, height);
  if (ratio > 1.08) return null;
  const cx = (box.minX + box.maxX) / 2;
  const cy = (box.minY + box.maxY) / 2;
  const distances = points.map((point) => Math.hypot(point.x - cx, point.y - cy));
  const average = avg(distances);
  const spread = Math.max(...distances) - Math.min(...distances);
  if (!average || spread / average > 0.08) return null;
  const signedAmount = direction === "outward" ? amount : -amount;
  return points.map((point) => {
    const dx = point.x - cx;
    const dy = point.y - cy;
    const distance = Math.hypot(dx, dy) || 1;
    const next = Math.max(0.1, distance + signedAmount);
    return { x: cx + (dx / distance) * next, y: cy + (dy / distance) * next };
  });
}

function lineIntersection(a1, a2, b1, b2) {
  const dax = a2.x - a1.x;
  const day = a2.y - a1.y;
  const dbx = b2.x - b1.x;
  const dby = b2.y - b1.y;
  const denominator = dax * dby - day * dbx;
  if (Math.abs(denominator) < 1e-9) return null;
  const t = ((b1.x - a1.x) * dby - (b1.y - a1.y) * dbx) / denominator;
  return { x: a1.x + t * dax, y: a1.y + t * day };
}

function cleanSelfIntersectingOffset(points) {
  let cleaned = cleanCurvePoints(points);
  for (let pass = 0; pass < 20; pass += 1) {
    const cut = firstSelfIntersection(cleaned);
    if (!cut) break;
    cleaned = keepLargerIntersectionLoop(cleaned, cut);
  }
  return cleanCurvePoints(cleaned);
}

function firstSelfIntersection(points) {
  const count = points.length;
  for (let i = 0; i < count; i += 1) {
    const a1 = points[i];
    const a2 = points[(i + 1) % count];
    for (let j = i + 2; j < count; j += 1) {
      if (i === 0 && j === count - 1) continue;
      const b1 = points[j];
      const b2 = points[(j + 1) % count];
      const intersection = segmentIntersection(a1, a2, b1, b2);
      if (intersection) return { i, j, point: intersection };
    }
  }
  return null;
}

function segmentIntersection(a1, a2, b1, b2) {
  const dax = a2.x - a1.x;
  const day = a2.y - a1.y;
  const dbx = b2.x - b1.x;
  const dby = b2.y - b1.y;
  const denominator = dax * dby - day * dbx;
  if (Math.abs(denominator) < 1e-9) return null;
  const t = ((b1.x - a1.x) * dby - (b1.y - a1.y) * dbx) / denominator;
  const u = ((b1.x - a1.x) * day - (b1.y - a1.y) * dax) / denominator;
  if (t <= 0.001 || t >= 0.999 || u <= 0.001 || u >= 0.999) return null;
  return { x: a1.x + t * dax, y: a1.y + t * day };
}

function keepLargerIntersectionLoop(points, cut) {
  const loopA = [cut.point, ...points.slice(cut.i + 1, cut.j + 1)];
  const loopB = [cut.point, ...points.slice(cut.j + 1), ...points.slice(0, cut.i + 1)];
  return polygonArea(loopA) >= polygonArea(loopB) ? loopA : loopB;
}

function cleanCurvePoints(points) {
  const cleaned = [];
  points.forEach((point) => {
    const prev = cleaned[cleaned.length - 1];
    if (!prev || Math.hypot(point.x - prev.x, point.y - prev.y) > 0.001) cleaned.push(point);
  });
  return cleaned;
}

function centerPoints(points, cx, cy) {
  const box = bounds([{ points }]);
  const currentCx = (box.minX + box.maxX) / 2;
  const currentCy = (box.minY + box.maxY) / 2;
  return points.map((point) => ({ x: point.x - currentCx + cx, y: point.y - currentCy + cy }));
}

function movePoints(points, dx, dy) {
  return points.map((point) => ({ x: point.x + dx, y: point.y + dy }));
}

function rotatePoints(points, degrees) {
  const angle = degToRad(degrees);
  const cos = Math.cos(angle);
  const sin = Math.sin(angle);
  return points.map((point) => ({
    x: point.x * cos - point.y * sin,
    y: point.x * sin + point.y * cos,
  }));
}

function signedArea(points) {
  let total = 0;
  for (let i = 0; i < points.length; i += 1) {
    const p1 = points[i];
    const p2 = points[(i + 1) % points.length];
    total += p1.x * p2.y - p2.x * p1.y;
  }
  return total / 2;
}

function polygonArea(points) {
  return Math.abs(signedArea(points));
}

function bounds(entities) {
  const points = [];
  entities.forEach((entity) => {
    if (entity.points) points.push(...entity.points);
    if (entity.point) points.push(entity.point);
  });
  if (!points.length) return { minX: 0, minY: 0, maxX: 100, maxY: 60 };
  return {
    minX: Math.min(...points.map((point) => point.x)),
    minY: Math.min(...points.map((point) => point.y)),
    maxX: Math.max(...points.map((point) => point.x)),
    maxY: Math.max(...points.map((point) => point.y)),
  };
}

function expandBox(box, amount) {
  return {
    minX: box.minX - amount,
    minY: box.minY - amount,
    maxX: box.maxX + amount,
    maxY: box.maxY + amount,
  };
}

function render() {
  if (els.entityCount) els.entityCount.textContent = String(state.entities.length);
  if (els.layerCount) els.layerCount.textContent = state.generated.length ? "4" : "0";
  renderDxfSize();
  const learnedCount = state.learnedProfile?.sampleCount || 0;
  if (els.learnedCount) els.learnedCount.textContent = String(learnedCount);
  if (els.learnedSummary) {
    els.learnedSummary.textContent = learnedCount
      ? `${learnedCount} reference template sample${learnedCount === 1 ? "" : "s"} saved in this browser.`
      : "No reference template learned yet.";
  }
  if (els.trainingCount) els.trainingCount.textContent = String(state.trainingExamples.length);
  if (els.trainingSummary) {
    const revision = state.trainingMeta.revision || 0;
    els.trainingSummary.textContent = state.trainingStatus || `Rev ${revision} master JSON, ${state.trainingExamples.length} approved.`;
  }
  if (els.saveTrainingButton) els.saveTrainingButton.disabled = !state.entities.length;
  if (els.exportTrainingButton) els.exportTrainingButton.disabled = state.trainingExamples.length === 0;
  els.fileSummary.textContent = state.entities.length ? `${state.fileName}: ${state.entities.length} entities parsed` : "No supported DXF entities found.";
  if (els.issueList) renderIssues();
  renderPreview();
}

function renderDxfSize() {
  const outline = state.entities.length ? selectPatchOutline(state.entities, getFormula()) : null;
  if (!outline) {
    els.dxfSize.textContent = "0 x 0 mm";
    els.dxfSizeSummary.textContent = "";
    return;
  }
  const box = bounds([outline]);
  const width = round(box.maxX - box.minX);
  const height = round(box.maxY - box.minY);
  els.dxfSize.textContent = `${width} x ${height} mm`;
  els.dxfSizeSummary.textContent = "";
}

function renderIssues() {
  els.issueList.innerHTML = "";
  const messages = state.entities.length
    ? [
        "Layer 1 uses original patch DXF outline.",
        "Layers 2, 3, and 4 add a 7mm outward offset line.",
        `Patch transform is shared by all 4 templates: X ${getFormula().patchRelX}mm, Y ${getFormula().patchRelY}mm, rotate ${getFormula().patchRotation}deg.`,
        "Template is generated around 150mm x 180mm unless changed.",
        "Needle slot top edge is 50mm from template top edge and right edge is 11.5mm from template right edge.",
        state.learnedProfile
          ? `Learned reference profile active from ${state.learnedProfile.sampleCount} board sample(s).`
          : "No reference profile learned yet. Import the PIC artwork once to learn placement.",
      ]
    : ["Waiting for patch DXF import."];
  messages.forEach((message) => {
    const li = document.createElement("li");
    li.textContent = message;
    els.issueList.appendChild(li);
  });
}

function renderPreview() {
  const sourcePreview = selectedSourcePreviewEntities();
  const visible = state.generated.length ? state.generated : sourcePreview;
  const box = expandBox(bounds(visible), 30);
  const width = Math.max(1, box.maxX - box.minX);
  const height = Math.max(1, box.maxY - box.minY);
  state.baseViewBox = { x: box.minX, y: -box.maxY, width, height };
  if (!state.viewBox) state.viewBox = { ...state.baseViewBox };
  els.preview.setAttribute("viewBox", `${state.viewBox.x} ${state.viewBox.y} ${state.viewBox.width} ${state.viewBox.height}`);
  setZoomButtons(Boolean(visible.length));
  els.preview.innerHTML = "";
  if (!state.generated.length) sourcePreview.forEach((entity) => drawEntity(entity, "source-line"));
  state.generated.forEach((entity) => drawEntity(entity, generatedClass(entity.layer)));
}

function selectedSourcePreviewEntities() {
  if (!state.entities.length) return [];
  const outline = selectPatchOutline(state.entities, getFormula());
  if (!outline) return state.entities;
  return [{ ...outline, layer: "AI_SELECTED_SOURCE_PREVIEW" }];
}

function generatedClass(layer) {
  if (/SOURCE/.test(layer)) return "source-line";
  if (/BOARD/.test(layer)) return "board-line";
  if (/OFFSET/.test(layer)) return "offset-line";
  if (/NEEDLE_SLOT/.test(layer)) return "slot-line";
  if (/PATCH/.test(layer)) return "patch-line";
  return "label-text";
}

function drawEntity(entity, className) {
  if (isShapeEntity(entity)) {
    const d = entity.points.map((point, index) => `${index === 0 ? "M" : "L"} ${point.x} ${-point.y}`).join(" ");
    const path = svgEl("path", { d: d + (entity.closed ? " Z" : ""), class: className });
    els.preview.appendChild(path);
  }
  if (entity.type === "LINE") {
    const line = svgEl("line", {
      x1: entity.points[0].x,
      y1: -entity.points[0].y,
      x2: entity.points[1].x,
      y2: -entity.points[1].y,
      class: className,
    });
    els.preview.appendChild(line);
  }
  if (entity.type === "TEXT") {
    const text = svgEl("text", { x: entity.point.x, y: -entity.point.y, class: "label-text", "text-anchor": entity.anchor || "start" });
    text.textContent = entity.text;
    els.preview.appendChild(text);
  }
}

function svgEl(tag, attrs) {
  const el = document.createElementNS("http://www.w3.org/2000/svg", tag);
  Object.entries(attrs).forEach(([key, value]) => el.setAttribute(key, value));
  return el;
}

function resetView() {
  state.baseViewBox = null;
  state.viewBox = null;
}

function zoomPreview(factor) {
  if (!state.viewBox) return;
  const cx = state.viewBox.x + state.viewBox.width / 2;
  const cy = state.viewBox.y + state.viewBox.height / 2;
  const width = state.viewBox.width * factor;
  const height = state.viewBox.height * factor;
  state.viewBox = { x: cx - width / 2, y: cy - height / 2, width, height };
  renderPreview();
}

function setZoomButtons(enabled) {
  els.zoomInButton.disabled = !enabled;
  els.zoomOutButton.disabled = !enabled;
  els.zoomResetButton.disabled = !enabled;
}

function buildOutputDxf(generated, formula) {
  const lines = [
    "0", "SECTION",
    "2", "HEADER",
    "9", "$INSUNITS",
    "70", "4",
    "0", "ENDSEC",
    "0", "SECTION",
    "2", "ENTITIES",
  ];
  generated.forEach((entity) => appendEntity(lines, entity));
  lines.push(
    "0", "ENDSEC",
    "0", "EOF",
    "",
  );
  return lines.join("\n");
}

function appendEntity(lines, entity) {
  if (entity.previewOnly) return;
  if (isShapeEntity(entity)) {
    lines.push("0", "LWPOLYLINE", "8", entity.layer, "90", String(entity.points.length), "70", entity.closed ? "1" : "0");
    entity.points.forEach((point) => lines.push("10", fmt(point.x), "20", fmt(point.y)));
  }
  if (entity.type === "TEXT") {
    lines.push("0", "TEXT", "8", entity.layer, "10", fmt(entity.point.x), "20", fmt(entity.point.y), "40", fmt(entity.height || 8), "1", entity.text);
    if (entity.anchor === "end" || entity.anchor === "middle") {
      lines.push("72", entity.anchor === "middle" ? "1" : "2", "11", fmt(entity.point.x), "21", fmt(entity.point.y));
    }
  }
}

function fmt(value) {
  return Number(value).toFixed(3).replace(/\.?0+$/, "");
}
