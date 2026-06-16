const state = {
  fileName: "",
  sourceType: "dxf",
  simpleShape: null,
  imageTrace: null,
  imagePreviewOpacity: 0.35,
  pendingScanCrop: null,
  importStatus: "",
  importError: "",
  rawEntityTypes: {},
  entities: [],
  generated: [],
  learnedProfile: loadLearnedProfile(),
  trainingExamples: loadTrainingExamples(),
  trainingMeta: loadTrainingMeta(),
  trainingStatus: "Loading training data...",
  trainingSyncing: false,
  patchModel: null,
  patchModelStatus: "Loading patch model...",
  traceRefreshSeq: 0,
  traceRefreshTimer: null,
  baseViewBox: null,
  viewBox: null,
  pan: null,
};

const APP_CONFIG = window.PATCH_TEMPLATE_CONFIG || {};
const POWER_AUTOMATE_TRAINING_URL = APP_CONFIG.powerAutomateTrainingUrl || "";
const POWER_AUTOMATE_READ_URL = APP_CONFIG.powerAutomateReadUrl || "";
const PATCH_MODEL_URL = APP_CONFIG.patchModelUrl || "pytorch/data/patch-template-training-data.json";
const APP_BASE_URL = appBaseUrl();
const DEFAULT_FORMULA = {
  boardWidth: 150,
  boardHeight: 180,
  cornerRadius: 12,
  slotWidth: 30,
  slotHeight: 11,
  slotRightEdge: 11.5,
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
  imageInput: document.querySelector("#imageInput"),
  predictButton: document.querySelector("#predictButton"),
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
  slotRightEdge: document.querySelector("#slotRightEdge"),
  offset: document.querySelector("#offset"),
  scanTraceWidth: document.querySelector("#scanTraceWidth"),
  scanTraceMode: document.querySelector("#scanTraceMode"),
  scanThresholdAdjust: document.querySelector("#scanThresholdAdjust"),
  scanTraceDetail: document.querySelector("#scanTraceDetail"),
  scanTraceSmooth: document.querySelector("#scanTraceSmooth"),
  scanImageOpacity: document.querySelector("#scanImageOpacity"),
  scanCropButton: document.querySelector("#scanCropButton"),
  scanConfirmCropButton: document.querySelector("#scanConfirmCropButton"),
  scanCancelCropButton: document.querySelector("#scanCancelCropButton"),
  scanResetCropButton: document.querySelector("#scanResetCropButton"),
  templateNumber: document.querySelector("#templateNumber"),
  patchRelX: document.querySelector("#patchRelX"),
  patchRelY: document.querySelector("#patchRelY"),
  patchRotation: document.querySelector("#patchRotation"),
  simpleShapeType: document.querySelector("#simpleShapeType"),
  simpleShapeWidth: document.querySelector("#simpleShapeWidth"),
  simpleShapeHeight: document.querySelector("#simpleShapeHeight"),
  simpleShapeWidthLabel: document.querySelector("#simpleShapeWidthLabel"),
  simpleShapeHeightWrap: document.querySelector("#simpleShapeHeightWrap"),
  createSimpleShapeButton: document.querySelector("#createSimpleShapeButton"),
  scanAdjustPanel: document.querySelector(".scan-adjust-panel"),
  zoomInButton: document.querySelector("#zoomInButton"),
  zoomOutButton: document.querySelector("#zoomOutButton"),
  zoomResetButton: document.querySelector("#zoomResetButton"),
};

function getTemplateNumberFromQuery() {
  const params = new URLSearchParams(window.location.search);
  const value = params.get("templateNo") || params.get("templateNumber") || params.get("template");
  return value ? value.trim() : "";
}

function appBaseUrl() {
  const scripts = [...document.querySelectorAll("script[src]")];
  const appScript = [...scripts].reverse().find((script) => /(?:^|\/)app\.js(?:[?#].*)?$/.test(script.getAttribute("src") || ""));
  return new URL(".", appScript?.src || window.location.href).href;
}

if (els.templateNumber) {
  els.templateNumber.value = INITIAL_TEMPLATE_NUMBER;
}

if (state.learnedProfile) applyLearnedProfileToControls(state.learnedProfile);
render();
loadPatchModel();
loadTrainingMasterJsonFromPowerAutomate();
syncArtboardHeightToParameter();

els.dxfInput.addEventListener("change", async (event) => {
  const file = event.target.files[0];
  if (!file) return;
  await handleDxfFile(file);
});

els.imageInput.addEventListener("change", async (event) => {
  const file = event.target.files[0];
  if (!file) return;
  await handleImageFile(file);
});

["dragenter", "dragover"].forEach((eventName) => {
  els.previewPanel.addEventListener(eventName, (event) => {
    event.preventDefault();
    event.dataTransfer.dropEffect = "copy";
    els.previewPanel.classList.add("drop-active");
  });
});

["dragleave", "drop"].forEach((eventName) => {
  els.previewPanel.addEventListener(eventName, () => {
    els.previewPanel.classList.remove("drop-active");
  });
});

els.previewPanel.addEventListener("drop", async (event) => {
  event.preventDefault();
  const file = [...event.dataTransfer.files].find((item) => isSupportedImportFile(item));
  if (!file) return;
  if (/\.dxf$/i.test(file.name)) await handleDxfFile(file);
  else await handleImageFile(file);
});

els.clearButton.addEventListener("click", clearArtboard);
els.predictButton.addEventListener("click", predictPatchTemplate);
els.createSimpleShapeButton.addEventListener("click", createSimplePatchTemplate);
els.simpleShapeType.addEventListener("change", updateSimpleShapeControls);
updateSimpleShapeControls();

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
  link.download = (state.fileName || "simple-patch").replace(/\.dxf$/i, "") + "-patch-template.dxf";
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

function clearArtboard() {
  window.clearTimeout(state.traceRefreshTimer);
  state.traceRefreshTimer = null;
  state.traceRefreshSeq += 1;
  state.fileName = "";
  state.sourceType = "dxf";
  state.simpleShape = null;
  state.imageTrace = null;
  state.imagePreviewOpacity = 0.35;
  state.pendingScanCrop = null;
  state.rawEntityTypes = {};
  state.entities = [];
  state.generated = [];
  state.baseViewBox = null;
  state.viewBox = null;
  els.dxfInput.value = "";
  els.imageInput.value = "";
  resetParameters();
  els.applyButton.disabled = true;
  els.clearButton.disabled = true;
  els.exportButton.disabled = true;
  updatePredictButton();
  closeExportModal();
  render();
}

async function handleDxfFile(file) {
  loadDxf(await file.text(), file.name);
  if (els.dxfInput) els.dxfInput.value = "";
}

async function handleImageFile(file) {
  window.clearTimeout(state.traceRefreshTimer);
  state.traceRefreshTimer = null;
  state.traceRefreshSeq += 1;
  state.fileName = file.name;
  state.sourceType = "image-trace";
  state.importStatus = `Importing ${file.name}...`;
  state.importError = "";
  state.simpleShape = null;
  state.imageTrace = null;
  state.pendingScanCrop = null;
  state.entities = [];
  state.generated = [];
  els.applyButton.disabled = true;
  els.clearButton.disabled = false;
  els.exportButton.disabled = true;
  resetView();
  render();
  try {
    const trace = await traceScanFile(file);
    loadImageTrace(trace, file.name);
  } catch (error) {
    loadImageImportError(file.name, error);
  } finally {
    if (els.imageInput) els.imageInput.value = "";
  }
}

async function loadPatchModel() {
  try {
    const response = await fetch(PATCH_MODEL_URL, { cache: "no-store" });
    if (!response.ok) throw new Error(`HTTP ${response.status}`);
    const payload = await response.json();
    const model = normalizePatchModelPayload(payload);
    validatePatchModel(model);
    state.patchModel = model;
    state.patchModelStatus = `Loaded ${model.exampleCount || 0} patch training example(s).`;
    loadEmbeddedTrainingData(payload);
  } catch (error) {
    state.patchModel = null;
    state.patchModelStatus = "Patch model JSON not loaded. Use a local server or upload the model JSON.";
  }
  updatePredictButton();
}

function normalizePatchModelPayload(payload) {
  return payload?.model && Array.isArray(payload.model.layers) ? payload.model : payload;
}

function loadEmbeddedTrainingData(payload) {
  const trainingPayload = payload?.trainingData || payload?.model?.trainingData || (Array.isArray(payload?.examples) ? payload : null);
  if (!trainingPayload || !Array.isArray(trainingPayload.examples)) return;
  state.trainingExamples = trainingPayload.examples.map(normalizeTrainingExample);
  state.trainingMeta = {
    revision: Number(trainingPayload.revision) || state.trainingMeta.revision || 0,
    updatedAt: trainingPayload.updatedAt || trainingPayload.exportedAt || state.trainingMeta.updatedAt || "",
  };
  state.learnedProfile = trainingPayload.learnedProfile || state.learnedProfile;
  saveTrainingExamples(state.trainingExamples);
  saveTrainingMeta(state.trainingMeta);
  saveLearnedProfile(state.learnedProfile);
  const updated = state.trainingMeta.updatedAt ? ` · ${formatTrainingDate(state.trainingMeta.updatedAt)}` : "";
  state.trainingStatus = `Loaded from training/model JSON${updated}.`;
}

function validatePatchModel(model) {
  if (!model || !Array.isArray(model.featureKeys) || !Array.isArray(model.targetKeys) || !Array.isArray(model.layers)) {
    throw new Error("Invalid patch model JSON");
  }
}

function updatePredictButton() {
  if (!els.predictButton) return;
  const enabled = Boolean(state.patchModel && state.entities.length);
  els.predictButton.disabled = !enabled;
  const status = state.patchModelStatus ? `\nStatus: ${state.patchModelStatus}` : "";
  els.predictButton.removeAttribute("title");
  els.predictButton.dataset.tooltip = `Use the MLP neural network model to predict template size, patch position, rotation, needle slot, and offset from the imported DXF.${status}`;
}

function predictPatchTemplate() {
  if (!state.patchModel || !state.entities.length) return;
  const currentFormula = getFormula();
  const outline = selectPatchOutline(state.entities, currentFormula);
  if (!outline) return;
  const patch = patchFeatureSource(outline);
  const prediction = runPatchModelPrediction(state.patchModel, patch);
  applyPredictionToControls(prediction);
  state.generated = generatePatchTemplate(state.entities, getFormula());
  els.applyButton.disabled = state.entities.length === 0;
  els.clearButton.disabled = false;
  els.exportButton.disabled = state.generated.length === 0;
  resetView();
  render();
}

function patchFeatureSource(outline) {
  const parts = outline.parts?.length ? outline.parts : [outline];
  const outlineBox = bounds(parts);
  const width = outlineBox.maxX - outlineBox.minX;
  const height = outlineBox.maxY - outlineBox.minY;
  const area = parts.reduce((sum, part) => sum + Math.abs(polygonArea(part.points)), 0);
  const pointCount = parts.reduce((sum, part) => sum + part.points.length, 0);
  const entityTypes = state.rawEntityTypes || {};
  return {
    entityCount: state.entities.length || 1,
    entityTypes,
    type: outline.type || "LWPOLYLINE",
    partCount: parts.length,
    pointCount,
    widthMm: width,
    heightMm: height,
    areaMm2: area,
    simpleShapeType: state.simpleShape?.type || "",
    sourceType: state.sourceType || "dxf",
    scanTrace: scanTraceFeatureInput(),
  };
}

function runPatchModelPrediction(model, patch) {
  let values = patchFeatureVector(patch, model.featureKeys);
  values = normalizeVector(values, model.xMean, model.xStd);
  model.layers.forEach((layer) => {
    if (layer.type === "linear") values = linearLayer(values, layer);
    if (layer.type === "relu") values = values.map((value) => Math.max(0, value));
  });
  values = denormalizeVector(values, model.yMean, model.yStd);
  return model.targetKeys.reduce((acc, key, index) => {
    acc[key] = values[index];
    return acc;
  }, {});
}

function patchFeatureVector(patch, keys) {
  const entityTypes = patch.entityTypes || {};
  const sourceType = patch.sourceType || "dxf";
  const scanTrace = patch.scanTrace || {};
  const traceMode = scanTrace.traceMode || "";
  const width = Number(patch.widthMm) || 0;
  const height = Number(patch.heightMm) || 0;
  const area = Number(patch.areaMm2) || 0;
  const values = {
    entity_count: Number(patch.entityCount) || 1,
    lwpolyline_count: Number(entityTypes.LWPOLYLINE) || (patch.type === "LWPOLYLINE" ? 1 : 0),
    spline_count: Number(entityTypes.SPLINE) || (patch.type === "SPLINE" ? 1 : 0),
    line_count: Number(entityTypes.LINE) || (patch.type === "LINE" ? 1 : 0),
    arc_count: Number(entityTypes.ARC) || (patch.type === "ARC" ? 1 : 0),
    circle_count: Number(entityTypes.CIRCLE) || (patch.type === "CIRCLE" ? 1 : 0),
    hatch_count: Number(entityTypes.HATCH) || (patch.type === "HATCH" ? 1 : 0),
    multipatch_count: patch.type === "MULTIPATCH" ? 1 : 0,
    part_count: Number(patch.partCount) || 1,
    point_count: Number(patch.pointCount) || 0,
    patch_width_mm: width,
    patch_height_mm: height,
    patch_area_mm2: area,
    patch_aspect_ratio: height > 0 ? width / height : 0,
    patch_fill_ratio: width > 0 && height > 0 ? area / (width * height) : 0,
    simple_rectangle: patch.simpleShapeType === "rectangle" ? 1 : 0,
    simple_circle_or_oval: patch.simpleShapeType === "circle" ? 1 : 0,
    source_dxf: sourceType === "dxf" ? 1 : 0,
    source_image_trace: sourceType === "image-trace" ? 1 : 0,
    source_pdf_trace: scanTrace.sourceType === "pdf" ? 1 : 0,
    source_bmp_trace: scanTrace.sourceType === "bmp" ? 1 : 0,
    trace_mode_auto: traceMode === "auto" ? 1 : 0,
    trace_mode_silhouette: traceMode === "silhouette" ? 1 : 0,
    trace_mode_dark: traceMode === "dark" ? 1 : 0,
    trace_sensitivity: Number(scanTrace.sensitivity) || 0,
    trace_detail: Number(scanTrace.detail) || 0,
    trace_threshold: Number(scanTrace.threshold) || 0,
    trace_actual_width_mm: Number(scanTrace.actualWidthMm) || 0,
    trace_raster_width_px: Number(scanTrace.rasterWidthPx) || 0,
    trace_raster_height_px: Number(scanTrace.rasterHeightPx) || 0,
  };
  return keys.map((key) => Number(values[key]) || 0);
}

function normalizeVector(values, means, stds) {
  return values.map((value, index) => (value - Number(means[index] || 0)) / (Number(stds[index]) || 1));
}

function denormalizeVector(values, means, stds) {
  return values.map((value, index) => value * (Number(stds[index]) || 1) + Number(means[index] || 0));
}

function linearLayer(values, layer) {
  return layer.bias.map((bias, rowIndex) => {
    const weights = layer.weight[rowIndex] || [];
    return weights.reduce((sum, weight, index) => sum + Number(weight) * values[index], Number(bias) || 0);
  });
}

function applyPredictionToControls(prediction) {
  if (Number.isFinite(prediction.boardWidthMm)) els.boardWidth.value = Math.max(80, round1(prediction.boardWidthMm));
  if (Number.isFinite(prediction.boardHeightMm)) els.boardHeight.value = Math.max(100, round1(prediction.boardHeightMm));
  if (Number.isFinite(prediction.cornerRadiusMm)) els.cornerRadius.value = Math.max(0, round1(prediction.cornerRadiusMm));
  if (Number.isFinite(prediction.needleSlotWidthMm)) els.slotWidth.value = Math.max(1, round1(prediction.needleSlotWidthMm));
  if (Number.isFinite(prediction.needleSlotHeightMm)) els.slotHeight.value = Math.max(1, round1(prediction.needleSlotHeightMm));
  if (Number.isFinite(prediction.needleSlotRightEdgeMm)) els.slotRightEdge.value = Math.max(0, round1(prediction.needleSlotRightEdgeMm));
  if (Number.isFinite(prediction.outwardOffsetMm)) els.offset.value = Math.max(0, round1(prediction.outwardOffsetMm));
  if (Number.isFinite(prediction.patchRelX)) els.patchRelX.value = round1(prediction.patchRelX);
  if (Number.isFinite(prediction.patchRelY)) els.patchRelY.value = round1(prediction.patchRelY);
  if (Number.isFinite(prediction.patchRotationDeg)) els.patchRotation.value = round1(prediction.patchRotationDeg);
}

function updateSimpleShapeControls() {
  const shape = getSimpleShapeType();
  if (shape === "circle") {
    els.simpleShapeWidthLabel.textContent = "Width, mm";
    els.simpleShapeHeightWrap.classList.remove("hidden");
    return;
  }
  els.simpleShapeWidthLabel.textContent = "Width, mm";
  els.simpleShapeHeightWrap.classList.remove("hidden");
}

function getSimpleShapeType() {
  return els.simpleShapeType.value || "rectangle";
}

function createSimplePatchTemplate() {
  const shapeType = getSimpleShapeType();
  const width = Math.max(1, Number(els.simpleShapeWidth.value) || 1);
  const height = Math.max(1, Number(els.simpleShapeHeight.value) || 1);
  const entity = simpleShapeEntity(shapeType, width, height);
  state.sourceType = "simple-shape";
  state.simpleShape = {
    type: shapeType,
    widthMm: round(width),
    heightMm: round(height),
    diameterMm: shapeType === "circle" && Math.abs(width - height) < 0.001 ? round(width) : null,
  };
  state.fileName = simpleShapeFileName(state.simpleShape);
  state.entities = [entity];
  state.generated = generatePatchTemplate(state.entities, getFormula());
  resetView();
  els.applyButton.disabled = false;
  els.clearButton.disabled = false;
  els.exportButton.disabled = state.generated.length === 0;
  updatePredictButton();
  render();
}

function simpleShapeEntity(shapeType, width, height) {
  if (shapeType === "circle") {
    return { type: "LWPOLYLINE", layer: "AI_SIMPLE_PATCH_SOURCE", points: ovalPoints(0, 0, width, height), closed: true };
  }
  return {
    type: "LWPOLYLINE",
    layer: "AI_SIMPLE_PATCH_SOURCE",
    points: rectanglePoints(width, height),
    closed: true,
  };
}

function rectanglePoints(width, height) {
  const left = -width / 2;
  const right = width / 2;
  const top = height / 2;
  const bottom = -height / 2;
  return [
    { x: left, y: top },
    { x: right, y: top },
    { x: right, y: bottom },
    { x: left, y: bottom },
  ];
}

function simpleShapeFileName(shape) {
  if (shape.type === "circle") {
    const prefix = shape.diameterMm ? "circle" : "oval";
    return `simple-${prefix}-${shape.widthMm}x${shape.heightMm}mm.dxf`;
  }
  return `simple-rectangle-${shape.widthMm}x${shape.heightMm}mm.dxf`;
}

function resetParameters() {
  els.boardWidth.value = DEFAULT_FORMULA.boardWidth;
  els.boardHeight.value = DEFAULT_FORMULA.boardHeight;
  els.cornerRadius.value = DEFAULT_FORMULA.cornerRadius;
  els.slotWidth.value = DEFAULT_FORMULA.slotWidth;
  els.slotHeight.value = DEFAULT_FORMULA.slotHeight;
  els.slotRightEdge.value = DEFAULT_FORMULA.slotRightEdge;
  els.offset.value = DEFAULT_FORMULA.offset;
  els.templateNumber.value = DEFAULT_FORMULA.templateNumber;
  els.scanTraceWidth.value = 50;
  els.scanTraceMode.value = "auto";
  els.scanThresholdAdjust.value = 0;
  els.scanTraceDetail.value = 10;
  els.scanTraceSmooth.value = 0;
  els.scanImageOpacity.value = 35;
  els.patchRelX.value = DEFAULT_FORMULA.patchRelX;
  els.patchRelY.value = DEFAULT_FORMULA.patchRelY;
  els.patchRotation.value = DEFAULT_FORMULA.patchRotation;
  els.simpleShapeType.value = "rectangle";
  els.simpleShapeWidth.value = 50;
  els.simpleShapeHeight.value = 35;
  updateSimpleShapeControls();
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
  const layout = document.querySelector(".layout");
  const resize = () => {
    if (window.matchMedia("(max-width: 1180px)").matches) {
      els.previewPanel.style.removeProperty("--parameter-card-height");
      layout?.style.removeProperty("--desktop-panel-height");
      return;
    }
    const layoutTop = layout?.getBoundingClientRect().top || 0;
    const viewportHeight = window.visualViewport?.height || window.innerHeight;
    const availableHeight = Math.max(420, viewportHeight - layoutTop - 16);
    const panelHeight = Math.min(els.controls.scrollHeight, availableHeight);
    els.previewPanel.style.setProperty("--parameter-card-height", `${panelHeight}px`);
    layout?.style.setProperty("--desktop-panel-height", `${availableHeight}px`);
  };
  new ResizeObserver(resize).observe(els.controls);
  window.addEventListener("resize", resize);
  window.visualViewport?.addEventListener("resize", resize);
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
    state.trainingExamples = examples.map(normalizeTrainingExample);
    applyModelFromTrainingPayload(imported);
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
els.preview.addEventListener("pointerdown", startPreviewPan);
els.preview.addEventListener("pointermove", movePreviewPan);
els.preview.addEventListener("pointerup", endPreviewPan);
els.preview.addEventListener("pointercancel", endPreviewPan);
els.preview.addEventListener("keydown", handlePreviewKeyPan);

[els.boardWidth, els.boardHeight, els.cornerRadius, els.slotWidth, els.slotHeight, els.slotRightEdge, els.offset, els.templateNumber, els.patchRelX, els.patchRelY, els.patchRotation].forEach((el) => {
  el.addEventListener("input", () => {
    if (!state.entities.length) return;
    state.generated = generatePatchTemplate(state.entities, getFormula());
    els.exportButton.disabled = state.generated.length === 0;
    resetView();
    render();
  });
});

els.scanTraceWidth.addEventListener("input", () => {
  if (!state.imageTrace) return;
  applyScanTraceWidthChange();
});

els.scanTraceMode.addEventListener("change", () => scheduleImageTracePreviewRefresh({ keepView: true, delay: 0 }));
els.scanThresholdAdjust.addEventListener("input", () => scheduleImageTracePreviewRefresh({ keepView: true }));
els.scanTraceDetail.addEventListener("input", () => scheduleImageTracePreviewRefresh({ keepView: true }));
els.scanTraceSmooth.addEventListener("input", () => scheduleImageTracePreviewRefresh({ keepView: true }));
els.scanImageOpacity.addEventListener("input", () => {
  state.imagePreviewOpacity = getScanImageOpacity();
  render();
});
els.scanCropButton?.addEventListener("click", previewScanCrop);
els.scanConfirmCropButton?.addEventListener("click", applyPreviewedScanCrop);
els.scanCancelCropButton?.addEventListener("click", cancelScanCropPreview);
els.scanResetCropButton?.addEventListener("click", resetScanCrop);

function isSupportedImportFile(file) {
  return /\.dxf$/i.test(file.name) || /\.pdf$/i.test(file.name) || /^image\//i.test(file.type) || /\.bmp$/i.test(file.name);
}

function loadDxf(text, fileName) {
  state.fileName = fileName;
  state.sourceType = "dxf";
  state.simpleShape = null;
  state.imageTrace = null;
  state.pendingScanCrop = null;
  state.importStatus = "";
  state.importError = "";
  state.imagePreviewOpacity = 0.35;
  state.rawEntityTypes = countDxfEntityTypes(text);
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
  els.clearButton.disabled = false;
  els.exportButton.disabled = true;
  updatePredictButton();
  render();
}

function loadImageTrace(trace, fileName) {
  state.fileName = fileName;
  state.sourceType = "image-trace";
  state.simpleShape = null;
  state.imageTrace = trace;
  state.pendingScanCrop = null;
  state.importStatus = "";
  state.importError = "";
  state.imagePreviewOpacity = getScanImageOpacity();
  state.rawEntityTypes = { IMAGE_TRACE: 1 };
  if (!state.imageTrace.originalRaster) state.imageTrace.originalRaster = state.imageTrace.raster;
  applyDetectedTraceWidth(trace);
  applyRasterFallbackWidth(trace);
  state.entities = traceHasParts(trace) ? [imageTraceEntity(trace, getScanTraceWidth())] : [];
  state.generated = [];
  resetView();
  els.applyButton.disabled = state.entities.length === 0;
  els.clearButton.disabled = false;
  els.exportButton.disabled = true;
  updatePredictButton();
  render();
}

function loadImageImportError(fileName, error) {
  state.fileName = fileName;
  state.sourceType = "image-trace";
  state.simpleShape = null;
  state.imageTrace = null;
  state.pendingScanCrop = null;
  state.importStatus = "";
  state.importError = readableImportError(error);
  state.rawEntityTypes = { IMPORT_ERROR: 1 };
  state.entities = [];
  state.generated = [];
  resetView();
  els.applyButton.disabled = true;
  els.clearButton.disabled = false;
  els.exportButton.disabled = true;
  updatePredictButton();
  render();
}

function readableImportError(error) {
  const message = error?.message || String(error || "Unknown import error");
  return message.replace(/^Error:\s*/i, "");
}

function getScanTraceWidth() {
  return Math.max(1, Number(els.scanTraceWidth.value) || 50);
}

function applyDetectedTraceWidth(trace) {
  if (!Number.isFinite(trace?.actualTraceWidthMm) || trace.actualTraceWidthMm <= 0) return;
  els.scanTraceWidth.value = round1(trace.actualTraceWidthMm);
}

function applyRasterFallbackWidth(trace) {
  if (Number.isFinite(trace?.actualTraceWidthMm) && trace.actualTraceWidthMm > 0) return;
  if (!trace?.raster?.width) return;
  els.scanTraceWidth.value = round1((trace.raster.width * 25.4) / 96);
}

function getTraceThresholdAdjust() {
  return Number(els.scanThresholdAdjust.value) || 0;
}

function getTraceMode() {
  return els.scanTraceMode?.value || "auto";
}

function getTraceSimplifyTolerance() {
  const detail = Math.max(1, Math.min(20, Number(els.scanTraceDetail.value) || 10));
  return 0.2 + (20 - detail) * 0.22;
}

function getTraceSmoothValue() {
  return Math.max(0, Math.min(12, Number(els.scanTraceSmooth?.value) || 0));
}

function getScanImageOpacity() {
  return Math.max(0, Math.min(1, (Number(els.scanImageOpacity.value) || 0) / 100));
}

function traceHasParts(trace) {
  return Boolean(trace?.parts?.some((points) => points.length >= 3) || trace?.points?.length >= 3);
}

function scanTraceFeatureInput() {
  if (state.sourceType !== "image-trace" || !state.imageTrace) return null;
  const trace = state.imageTrace;
  return {
    sourceType: trace.sourceType || "image",
    traceEngine: trace.engine || "potrace-browser",
    traceMode: getTraceMode(),
    threshold: round(Number(trace.threshold) || 0),
    sensitivity: round(Number(getTraceThresholdAdjust()) || 0),
    detail: Math.max(1, Math.min(20, Number(els.scanTraceDetail.value) || 10)),
    smooth: getTraceSmoothValue(),
    actualWidthMm: round(getScanTraceWidth()),
    physicalSource: trace.physicalSource || "",
    imageWidthPx: Number(trace.imageWidthPx) || Number(trace.raster?.width) || 0,
    imageHeightPx: Number(trace.imageHeightPx) || Number(trace.raster?.height) || 0,
    rasterWidthPx: Number(trace.raster?.width) || 0,
    rasterHeightPx: Number(trace.raster?.height) || 0,
    foreground: trace.foreground || "",
  };
}

function previewScanCrop() {
  const pending = currentViewRasterCrop();
  if (!pending) return;
  state.pendingScanCrop = pending;
  render();
}

async function applyPreviewedScanCrop() {
  const pending = state.pendingScanCrop || currentViewRasterCrop();
  if (!pending) return;
  await cropScanToRasterCrop(pending.crop);
}

function cancelScanCropPreview() {
  state.pendingScanCrop = null;
  render();
}

function currentViewRasterCrop() {
  const trace = state.imageTrace;
  if (!trace?.raster || !state.viewBox) return;
  const image = imagePreviewEntity(trace, getScanTraceWidth());
  if (!image?.width || !image?.height) return;
  const raster = trace.raster;
  const mmPerPixelX = image.width / raster.width;
  const mmPerPixelY = image.height / raster.height;
  const viewRight = state.viewBox.x + state.viewBox.width;
  const viewBottom = state.viewBox.y + state.viewBox.height;
  const imageRight = image.x + image.width;
  const imageBottom = image.y + image.height;
  const left = Math.max(state.viewBox.x, image.x);
  const top = Math.max(state.viewBox.y, image.y);
  const right = Math.min(viewRight, imageRight);
  const bottom = Math.min(viewBottom, imageBottom);
  if (right <= left || bottom <= top) return;
  const rawCrop = {
    x: Math.floor((left - image.x) / mmPerPixelX),
    y: Math.floor((top - image.y) / mmPerPixelY),
    width: Math.ceil((right - left) / mmPerPixelX),
    height: Math.ceil((bottom - top) / mmPerPixelY),
  };
  const pad = Math.max(18, Math.round(Math.max(rawCrop.width, rawCrop.height) * 0.08));
  const crop = clampRasterCrop({
    x: rawCrop.x - pad,
    y: rawCrop.y - pad,
    width: rawCrop.width + pad * 2,
    height: rawCrop.height + pad * 2,
  }, raster);
  if (crop.width < 8 || crop.height < 8) return;
  return { crop };
}

async function cropScanToRasterCrop(crop) {
  const trace = state.imageTrace;
  if (!trace?.raster || !crop) return;
  window.clearTimeout(state.traceRefreshTimer);
  state.traceRefreshTimer = null;
  const raster = trace.raster;
  const croppedRaster = cropRaster(raster, crop);
  const traced = await traceRasterWithSelectedEngine(croppedRaster);
  const nextTrace = {
    ...trace,
    ...withActualTraceWidth(traced, trace.mmPerRasterPixel),
    raster: croppedRaster,
    crop: {
      x: (trace.crop?.x || 0) + crop.x,
      y: (trace.crop?.y || 0) + crop.y,
      width: crop.width,
      height: crop.height,
    },
  };
  state.imageTrace = nextTrace;
  state.pendingScanCrop = null;
  applyDetectedTraceWidth(state.imageTrace);
  state.entities = traceHasParts(state.imageTrace) ? [imageTraceEntity(state.imageTrace, getScanTraceWidth())] : [];
  state.generated = [];
  els.applyButton.disabled = state.entities.length === 0;
  els.exportButton.disabled = true;
  resetView();
  updatePredictButton();
  render();
  zoomPreview(0.82);
}

async function resetScanCrop() {
  const trace = state.imageTrace;
  if (!trace?.originalRaster || trace.raster === trace.originalRaster) return;
  const traced = await traceRasterWithSelectedEngine(trace.originalRaster);
  state.imageTrace = {
    ...trace,
    ...withActualTraceWidth(traced, trace.mmPerRasterPixel),
    raster: trace.originalRaster,
    crop: null,
  };
  state.pendingScanCrop = null;
  applyDetectedTraceWidth(state.imageTrace);
  state.entities = traceHasParts(state.imageTrace) ? [imageTraceEntity(state.imageTrace, getScanTraceWidth())] : [];
  state.generated = [];
  els.applyButton.disabled = state.entities.length === 0;
  els.exportButton.disabled = true;
  resetView();
  updatePredictButton();
  render();
}

function applyScanTraceWidthChange() {
  if (!state.imageTrace) return;
  state.entities = traceHasParts(state.imageTrace) ? [imageTraceEntity(state.imageTrace, getScanTraceWidth())] : [];
  if (state.generated.length) {
    state.generated = generatePatchTemplate(state.entities, getFormula());
  }
  els.applyButton.disabled = state.entities.length === 0;
  els.exportButton.disabled = state.generated.length === 0;
  updatePredictButton();
  render();
}

function clampRasterCrop(crop, raster) {
  const x = Math.max(0, Math.min(raster.width - 1, crop.x));
  const y = Math.max(0, Math.min(raster.height - 1, crop.y));
  const maxWidth = raster.width - x;
  const maxHeight = raster.height - y;
  return {
    x,
    y,
    width: Math.max(1, Math.min(maxWidth, crop.width)),
    height: Math.max(1, Math.min(maxHeight, crop.height)),
  };
}

function cropRaster(raster, crop) {
  const canvas = document.createElement("canvas");
  canvas.width = crop.width;
  canvas.height = crop.height;
  const ctx = canvas.getContext("2d", { willReadFrequently: true });
  const imageData = ctx.createImageData(crop.width, crop.height);
  for (let y = 0; y < crop.height; y += 1) {
    const sourceStart = ((crop.y + y) * raster.width + crop.x) * 4;
    const sourceEnd = sourceStart + crop.width * 4;
    imageData.data.set(raster.rgba.slice(sourceStart, sourceEnd), y * crop.width * 4);
  }
  ctx.putImageData(imageData, 0, 0);
  const cropped = rasterFromCanvas(canvas);
  cropped.scale = raster.scale || 1;
  cropped.originalWidth = raster.originalWidth || raster.width;
  cropped.originalHeight = raster.originalHeight || raster.height;
  return cropped;
}

function scheduleImageTracePreviewRefresh(options = {}) {
  window.clearTimeout(state.traceRefreshTimer);
  state.traceRefreshTimer = window.setTimeout(() => {
    state.traceRefreshTimer = null;
    refreshImageTracePreview(options);
  }, options.delay ?? 260);
}

async function refreshImageTracePreview(options = {}) {
  if (!state.imageTrace?.raster) return;
  const seq = state.traceRefreshSeq + 1;
  state.traceRefreshSeq = seq;
  const traced = await traceRasterWithSelectedEngine(state.imageTrace.raster);
  if (seq !== state.traceRefreshSeq) return;
  state.imageTrace = { ...state.imageTrace, ...traced };
  state.pendingScanCrop = null;
  state.rawEntityTypes = { IMAGE_TRACE: 1 };
  state.entities = traceHasParts(state.imageTrace) ? [imageTraceEntity(state.imageTrace, getScanTraceWidth())] : [];
  state.generated = [];
  els.applyButton.disabled = state.entities.length === 0;
  els.exportButton.disabled = true;
  if (!options.keepView) resetView();
  updatePredictButton();
  render();
}

async function traceScanFile(file) {
  if (/\.pdf$/i.test(file.name) || file.type === "application/pdf") return tracePdfFile(file);
  if (/\.bmp$/i.test(file.name) || file.type === "image/bmp") return traceBmpFile(file);
  return traceImageFile(file);
}

async function traceImageFile(file) {
  const image = await loadRasterImage(file);
  const raster = rasterizeImage(image);
  const physical = await imagePhysicalScale(file, raster);
  const traced = await traceRasterWithSelectedEngine(raster);
  return {
    ...withActualTraceWidth(traced, physical?.mmPerRasterPixel),
    raster,
    originalRaster: raster,
    mmPerRasterPixel: physical?.mmPerRasterPixel || 0,
    sourceType: "image",
    physicalSource: physical?.source || "",
  };
}

function rasterizeImage(image) {
  const maxSide = 1800;
  const scale = Math.min(1, maxSide / Math.max(image.naturalWidth || image.width, image.naturalHeight || image.height));
  const width = Math.max(1, Math.round((image.naturalWidth || image.width) * scale));
  const height = Math.max(1, Math.round((image.naturalHeight || image.height) * scale));
  const canvas = document.createElement("canvas");
  canvas.width = width;
  canvas.height = height;
  const ctx = canvas.getContext("2d", { willReadFrequently: true });
  ctx.drawImage(image, 0, 0, width, height);
  const raster = rasterFromCanvas(canvas);
  raster.scale = scale;
  raster.originalWidth = image.naturalWidth || image.width;
  raster.originalHeight = image.naturalHeight || image.height;
  return raster;
}

function rasterFromCanvas(canvas) {
  const ctx = canvas.getContext("2d", { willReadFrequently: true });
  const imageData = ctx.getImageData(0, 0, canvas.width, canvas.height);
  return {
    width: canvas.width,
    height: canvas.height,
    gray: grayscalePixels(imageData.data),
    rgba: new Uint8ClampedArray(imageData.data),
    dataUrl: canvas.toDataURL("image/png"),
  };
}

function traceRaster(raster) {
  const width = raster.width;
  const height = raster.height;
  const gray = raster.gray;
  const threshold = otsuThreshold(gray);
  const adjustedThreshold = Math.max(0, Math.min(255, threshold + getTraceThresholdAdjust()));
  const foregroundIsDark = borderAverage(gray, width, height) > threshold;
  const mode = getTraceMode();
  const darkMask = () => foregroundMask(gray, width, height, adjustedThreshold, foregroundIsDark);
  let mask = mode === "dark" ? darkMask() : silhouetteMask(raster, width, height);
  let pixelParts = traceMaskParts(mask, width, height);
  if (mode === "auto" && !pixelParts.length) {
    mask = darkMask();
    pixelParts = traceMaskParts(mask, width, height);
  }
  const pixelLoop = pixelParts[0] || [];
  return {
    points: pixelLoop,
    parts: pixelParts,
    imageWidthPx: width,
    imageHeightPx: height,
    threshold: adjustedThreshold,
    foreground: foregroundIsDark ? "dark" : "light",
  };
}

async function traceRasterWithSelectedEngine(raster) {
  return traceRasterWithBrowserPotrace(raster);
}

async function traceRasterWithBrowserPotrace(raster) {
  await loadBrowserPotrace();
  const workRaster = downscaleRasterForPotrace(raster);
  const scaleBack = raster.width / workRaster.width;
  const mask = potracePreprocessMask(workRaster);
  const canvas = maskToCanvas(mask, workRaster.width, workRaster.height);
  const url = canvas.toDataURL("image/png");
  window.Potrace.setParameter({
    turnpolicy: "minority",
    turdsize: Math.max(2, Math.round(mask.length * 0.00003)),
    alphamax: 0.65 + Math.max(1, Math.min(20, Number(els.scanTraceDetail.value) || 10)) * 0.045,
    optcurve: true,
    opttolerance: 0.18,
  });
  window.Potrace.loadImageFromUrl(url);
  await new Promise((resolve) => window.Potrace.process(resolve));
  const parts = svgToTraceParts(window.Potrace.getSVG(1), getTraceSimplifyTolerance())
    .map((points) => scaleTracePoints(points, scaleBack))
    .filter((points) => points.length >= 3)
    .sort((a, b) => Math.abs(signedArea(b)) - Math.abs(signedArea(a)));
  if (!parts.length) return { ...traceRaster(raster), engine: "potrace-browser-empty" };
  return {
    imageWidthPx: raster.width,
    imageHeightPx: raster.height,
    threshold: 0,
    foreground: "potrace-browser",
    points: parts[0],
    parts,
    engine: "potrace-browser",
  };
}

function loadBrowserPotrace() {
  if (window.Potrace) return Promise.resolve(window.Potrace);
  return new Promise((resolve, reject) => {
    const script = document.createElement("script");
    script.src = new URL("potrace.js?v=20260616", APP_BASE_URL).href;
    script.onload = () => window.Potrace ? resolve(window.Potrace) : reject(new Error("Potrace browser script not loaded"));
    script.onerror = () => reject(new Error("Potrace browser script missing"));
    document.head.appendChild(script);
  });
}

function downscaleRasterForPotrace(raster) {
  const maxSide = 1000;
  const scale = Math.min(1, maxSide / Math.max(raster.width, raster.height));
  if (scale >= 1) return raster;
  const canvas = document.createElement("canvas");
  canvas.width = Math.max(1, Math.round(raster.width * scale));
  canvas.height = Math.max(1, Math.round(raster.height * scale));
  const source = document.createElement("canvas");
  source.width = raster.width;
  source.height = raster.height;
  const sourceCtx = source.getContext("2d", { willReadFrequently: true });
  const imageData = sourceCtx.createImageData(raster.width, raster.height);
  imageData.data.set(raster.rgba);
  sourceCtx.putImageData(imageData, 0, 0);
  const ctx = canvas.getContext("2d", { willReadFrequently: true });
  ctx.imageSmoothingEnabled = true;
  ctx.imageSmoothingQuality = "high";
  ctx.drawImage(source, 0, 0, canvas.width, canvas.height);
  const scaled = rasterFromCanvas(canvas);
  scaled.scale = (raster.scale || 1) * scale;
  scaled.originalWidth = raster.originalWidth || raster.width;
  scaled.originalHeight = raster.originalHeight || raster.height;
  return scaled;
}

function scaleTracePoints(points, scale) {
  if (Math.abs(scale - 1) < 0.0001) return points;
  return points.map((point) => ({ x: point.x * scale, y: point.y * scale }));
}

function maskToCanvas(mask, width, height) {
  const canvas = document.createElement("canvas");
  canvas.width = width;
  canvas.height = height;
  const ctx = canvas.getContext("2d", { willReadFrequently: true });
  const imageData = ctx.createImageData(width, height);
  for (let index = 0; index < mask.length; index += 1) {
    const value = mask[index] ? 0 : 255;
    const i = index * 4;
    imageData.data[i] = value;
    imageData.data[i + 1] = value;
    imageData.data[i + 2] = value;
    imageData.data[i + 3] = 255;
  }
  ctx.putImageData(imageData, 0, 0);
  return canvas;
}

function svgToTraceParts(svg, tolerance) {
  const parts = [];
  const matches = String(svg || "").matchAll(/<path[^>]*\sd="([^"]+)"/g);
  for (const match of matches) {
    parts.push(...svgPathToParts(match[1], tolerance));
  }
  return parts;
}

function svgPathToParts(pathData, tolerance) {
  const tokens = String(pathData || "").match(/[MmLlHhVvCcQqZz]|[-+]?(?:\d*\.\d+|\d+)(?:[eE][-+]?\d+)?/g) || [];
  const parts = [];
  let part = [];
  let x = 0;
  let y = 0;
  let startX = 0;
  let startY = 0;
  let command = "";
  let index = 0;
  const steps = Math.max(6, Math.min(24, Math.round(24 - tolerance * 2)));
  const isCommand = (token) => /^[A-Za-z]$/.test(token);
  const hasNumber = () => index < tokens.length && !isCommand(tokens[index]);
  const number = () => Number(tokens[index++]);
  const closePart = () => {
    const cleaned = cleanClosingPoint(cleanCurvePoints(part));
    if (cleaned.length >= 3) parts.push(cleaned);
    part = [];
  };
  while (index < tokens.length) {
    if (isCommand(tokens[index])) command = tokens[index++];
    if (command === "M" || command === "m") {
      let first = true;
      while (hasNumber()) {
        let nx = number();
        let ny = number();
        if (command === "m") {
          nx += x;
          ny += y;
        }
        x = nx;
        y = ny;
        if (first) {
          if (part.length >= 3) closePart();
          part = [{ x, y }];
          startX = x;
          startY = y;
          first = false;
        } else {
          part.push({ x, y });
        }
      }
      command = command === "M" ? "L" : "l";
    } else if (command === "L" || command === "l") {
      while (hasNumber()) {
        let nx = number();
        let ny = number();
        if (command === "l") {
          nx += x;
          ny += y;
        }
        x = nx;
        y = ny;
        part.push({ x, y });
      }
    } else if (command === "H" || command === "h") {
      while (hasNumber()) {
        let nx = number();
        if (command === "h") nx += x;
        x = nx;
        part.push({ x, y });
      }
    } else if (command === "V" || command === "v") {
      while (hasNumber()) {
        let ny = number();
        if (command === "v") ny += y;
        y = ny;
        part.push({ x, y });
      }
    } else if (command === "C" || command === "c") {
      while (hasNumber()) {
        let x1 = number();
        let y1 = number();
        let x2 = number();
        let y2 = number();
        let x3 = number();
        let y3 = number();
        if (command === "c") {
          x1 += x;
          y1 += y;
          x2 += x;
          y2 += y;
          x3 += x;
          y3 += y;
        }
        part.push(...sampleCubic(x, y, x1, y1, x2, y2, x3, y3, steps));
        x = x3;
        y = y3;
      }
    } else if (command === "Q" || command === "q") {
      while (hasNumber()) {
        let x1 = number();
        let y1 = number();
        let x2 = number();
        let y2 = number();
        if (command === "q") {
          x1 += x;
          y1 += y;
          x2 += x;
          y2 += y;
        }
        part.push(...sampleQuadratic(x, y, x1, y1, x2, y2, steps));
        x = x2;
        y = y2;
      }
    } else if (command === "Z" || command === "z") {
      x = startX;
      y = startY;
      closePart();
      command = "";
    } else {
      break;
    }
  }
  if (part.length >= 3) closePart();
  return parts;
}

function sampleCubic(x0, y0, x1, y1, x2, y2, x3, y3, steps) {
  const points = [];
  for (let step = 1; step <= steps; step += 1) {
    const t = step / steps;
    const mt = 1 - t;
    points.push({
      x: mt ** 3 * x0 + 3 * mt ** 2 * t * x1 + 3 * mt * t ** 2 * x2 + t ** 3 * x3,
      y: mt ** 3 * y0 + 3 * mt ** 2 * t * y1 + 3 * mt * t ** 2 * y2 + t ** 3 * y3,
    });
  }
  return points;
}

function sampleQuadratic(x0, y0, x1, y1, x2, y2, steps) {
  const points = [];
  for (let step = 1; step <= steps; step += 1) {
    const t = step / steps;
    const mt = 1 - t;
    points.push({
      x: mt ** 2 * x0 + 2 * mt * t * x1 + t ** 2 * x2,
      y: mt ** 2 * y0 + 2 * mt * t * y1 + t ** 2 * y2,
    });
  }
  return points;
}

function traceMaskForCurrentSettings(raster) {
  const width = raster.width;
  const height = raster.height;
  const gray = raster.gray;
  const threshold = Math.max(0, Math.min(255, otsuThreshold(gray) + getTraceThresholdAdjust()));
  const foregroundIsDark = borderAverage(gray, width, height) > threshold;
  const mode = getTraceMode();
  if (mode === "dark") return foregroundMask(gray, width, height, threshold, foregroundIsDark);
  return silhouetteMask(raster, width, height);
}

function potracePreprocessMask(raster) {
  const width = raster.width;
  const height = raster.height;
  const smooth = getTraceSmoothValue();
  const mode = getTraceMode();
  let score = mode === "dark"
    ? darkForegroundScore(raster)
    : silhouetteScore(raster);
  score = boxBlurGray(score, width, height, Math.max(1, Math.round(1 + smooth * 0.18)));
  const threshold = scoreThreshold(score, mode === "dark" ? 128 : Math.max(10, 20 - getTraceThresholdAdjust() * 0.12));
  let mask = thresholdScoreMask(score, threshold);
  const closeIterations = Math.max(1, Math.round(1 + smooth * 0.22));
  mask = closeMask(mask, width, height, closeIterations);
  if (smooth >= 5) mask = openMask(mask, width, height, Math.max(1, Math.round(smooth / 6)));
  return significantComponentMask(mask, width, height);
}

function significantComponentMask(mask, width, height) {
  const components = connectedBinaryComponents(mask, width, height);
  if (!components.length) return mask;
  const largest = components[0].length;
  const minArea = Math.max(30, Math.round(mask.length * 0.000025), Math.round(largest * 0.006));
  const result = new Uint8Array(mask.length);
  components
    .filter((component) => component.length >= minArea)
    .slice(0, 80)
    .forEach((component) => {
      component.forEach((index) => {
        result[index] = 1;
      });
    });
  return result;
}

function silhouetteScore(raster) {
  const width = raster.width;
  const height = raster.height;
  const bg = borderColorAverage(raster.rgba, width, height);
  const score = new Uint8Array(width * height);
  for (let index = 0; index < score.length; index += 1) {
    const i = index * 4;
    const r = raster.rgba[i];
    const g = raster.rgba[i + 1];
    const b = raster.rgba[i + 2];
    const diff = Math.hypot(r - bg.r, g - bg.g, b - bg.b);
    const dark = Math.max(0, bg.gray - raster.gray[index]);
    score[index] = Math.max(0, Math.min(255, Math.round(Math.max(diff, dark * 1.15))));
  }
  return score;
}

function darkForegroundScore(raster) {
  const score = new Uint8Array(raster.gray.length);
  const threshold = Math.max(0, Math.min(255, otsuThreshold(raster.gray) + getTraceThresholdAdjust()));
  const foregroundIsDark = borderAverage(raster.gray, raster.width, raster.height) > threshold;
  for (let index = 0; index < score.length; index += 1) {
    const value = foregroundIsDark ? threshold - raster.gray[index] : raster.gray[index] - threshold;
    score[index] = Math.max(0, Math.min(255, Math.round(128 + value)));
  }
  return score;
}

function boxBlurGray(values, width, height, radius) {
  if (radius <= 0) return values;
  const temp = new Uint8Array(values.length);
  const result = new Uint8Array(values.length);
  for (let y = 0; y < height; y += 1) {
    for (let x = 0; x < width; x += 1) {
      let total = 0;
      let count = 0;
      for (let dx = -radius; dx <= radius; dx += 1) {
        const nx = x + dx;
        if (nx < 0 || nx >= width) continue;
        total += values[y * width + nx];
        count += 1;
      }
      temp[y * width + x] = Math.round(total / count);
    }
  }
  for (let y = 0; y < height; y += 1) {
    for (let x = 0; x < width; x += 1) {
      let total = 0;
      let count = 0;
      for (let dy = -radius; dy <= radius; dy += 1) {
        const ny = y + dy;
        if (ny < 0 || ny >= height) continue;
        total += temp[ny * width + x];
        count += 1;
      }
      result[y * width + x] = Math.round(total / count);
    }
  }
  return result;
}

function scoreThreshold(score, fallback) {
  const auto = otsuThreshold(score);
  return Math.max(1, Math.min(254, Math.round((auto + fallback) / 2 + getTraceThresholdAdjust() * 0.08)));
}

function thresholdScoreMask(score, threshold) {
  const mask = new Uint8Array(score.length);
  for (let index = 0; index < score.length; index += 1) {
    mask[index] = score[index] >= threshold ? 1 : 0;
  }
  return mask;
}

function silhouetteMask(raster, width, height) {
  if (!raster.rgba?.length) return largestForegroundComponent(raster.gray, width, height, otsuThreshold(raster.gray), true);
  const bg = borderColorAverage(raster.rgba, width, height);
  const adjust = getTraceThresholdAdjust();
  const threshold = Math.max(5, Math.min(52, 18 - adjust * 0.16));
  const raw = new Uint8Array(width * height);
  for (let index = 0; index < raw.length; index += 1) {
    const i = index * 4;
    const r = raster.rgba[i];
    const g = raster.rgba[i + 1];
    const b = raster.rgba[i + 2];
    const diff = Math.hypot(r - bg.r, g - bg.g, b - bg.b);
    const chroma = Math.max(r, g, b) - Math.min(r, g, b);
    const gray = raster.gray[index];
    raw[index] = diff > threshold || (chroma > threshold * 0.9 && diff > threshold * 0.55) || gray < bg.gray - threshold * 1.2 ? 1 : 0;
  }
  return closeMask(raw, width, height, 2);
}

function largestBinaryComponent(mask, width, height) {
  const visited = new Uint8Array(mask.length);
  let best = [];
  const queue = [];
  for (let start = 0; start < mask.length; start += 1) {
    if (!mask[start] || visited[start]) continue;
    const component = [];
    queue.length = 0;
    queue.push(start);
    visited[start] = 1;
    for (let q = 0; q < queue.length; q += 1) {
      const index = queue[q];
      component.push(index);
      const x = index % width;
      const y = Math.floor(index / width);
      for (let dy = -1; dy <= 1; dy += 1) {
        for (let dx = -1; dx <= 1; dx += 1) {
          if (!dx && !dy) continue;
          const nx = x + dx;
          const ny = y + dy;
          if (nx < 0 || nx >= width || ny < 0 || ny >= height) continue;
          const next = ny * width + nx;
          if (mask[next] && !visited[next]) {
            visited[next] = 1;
            queue.push(next);
          }
        }
      }
    }
    if (component.length > best.length) best = component;
  }
  const result = new Uint8Array(mask.length);
  best.forEach((index) => {
    result[index] = 1;
  });
  return result;
}

function borderColorAverage(rgba, width, height) {
  let r = 0;
  let g = 0;
  let b = 0;
  let count = 0;
  const add = (x, y) => {
    const i = (y * width + x) * 4;
    r += rgba[i];
    g += rgba[i + 1];
    b += rgba[i + 2];
    count += 1;
  };
  for (let x = 0; x < width; x += 1) {
    add(x, 0);
    add(x, height - 1);
  }
  for (let y = 1; y < height - 1; y += 1) {
    add(0, y);
    add(width - 1, y);
  }
  r /= count || 1;
  g /= count || 1;
  b /= count || 1;
  return { r, g, b, gray: r * 0.299 + g * 0.587 + b * 0.114 };
}

function closeMask(mask, width, height, iterations) {
  let next = mask;
  for (let i = 0; i < iterations; i += 1) next = dilateMask(next, width, height);
  for (let i = 0; i < iterations; i += 1) next = erodeMask(next, width, height);
  return next;
}

function openMask(mask, width, height, iterations) {
  let next = mask;
  for (let i = 0; i < iterations; i += 1) next = erodeMask(next, width, height);
  for (let i = 0; i < iterations; i += 1) next = dilateMask(next, width, height);
  return next;
}

function dilateMask(mask, width, height) {
  const result = new Uint8Array(mask.length);
  for (let y = 0; y < height; y += 1) {
    for (let x = 0; x < width; x += 1) {
      let on = false;
      for (let dy = -1; dy <= 1 && !on; dy += 1) {
        for (let dx = -1; dx <= 1; dx += 1) {
          const nx = x + dx;
          const ny = y + dy;
          if (nx >= 0 && nx < width && ny >= 0 && ny < height && mask[ny * width + nx]) {
            on = true;
            break;
          }
        }
      }
      result[y * width + x] = on ? 1 : 0;
    }
  }
  return result;
}

function erodeMask(mask, width, height) {
  const result = new Uint8Array(mask.length);
  for (let y = 0; y < height; y += 1) {
    for (let x = 0; x < width; x += 1) {
      let on = true;
      for (let dy = -1; dy <= 1 && on; dy += 1) {
        for (let dx = -1; dx <= 1; dx += 1) {
          const nx = x + dx;
          const ny = y + dy;
          if (nx < 0 || nx >= width || ny < 0 || ny >= height || !mask[ny * width + nx]) {
            on = false;
            break;
          }
        }
      }
      result[y * width + x] = on ? 1 : 0;
    }
  }
  return result;
}

function withActualTraceWidth(trace, mmPerRasterPixel) {
  const pixelParts = trace.parts?.length ? trace.parts.filter((points) => points.length >= 3) : (trace.points?.length >= 3 ? [trace.points] : []);
  if (!Number.isFinite(mmPerRasterPixel) || mmPerRasterPixel <= 0 || !pixelParts.length) return trace;
  const box = bounds(pixelParts.map((points) => ({ points })));
  return {
    ...trace,
    actualTraceWidthMm: (box.maxX - box.minX) * mmPerRasterPixel,
  };
}

async function tracePdfFile(file) {
  const pdfjs = await loadPdfJs();
  const data = new Uint8Array(await file.arrayBuffer());
  const doc = await pdfjs.getDocument({ data }).promise;
  const page = await doc.getPage(1);
  const baseViewport = page.getViewport({ scale: 1 });
  const scale = Math.max(0.5, Math.min(2.5, 900 / Math.max(baseViewport.width, baseViewport.height)));
  const viewport = page.getViewport({ scale });
  const canvas = document.createElement("canvas");
  canvas.width = Math.ceil(viewport.width);
  canvas.height = Math.ceil(viewport.height);
  const ctx = canvas.getContext("2d", { willReadFrequently: true });
  ctx.fillStyle = "#fff";
  ctx.fillRect(0, 0, canvas.width, canvas.height);
  await page.render({ canvasContext: ctx, viewport }).promise;
  const raster = rasterFromCanvas(canvas);
  const traced = await traceRasterWithSelectedEngine(raster);
  const mmPerRasterPixel = 25.4 / 72 / scale;
  return {
    ...withActualTraceWidth(traced, mmPerRasterPixel),
    raster,
    originalRaster: raster,
    mmPerRasterPixel,
    sourceType: "pdf",
    pageNumber: 1,
    physicalSource: "pdf-page-units",
  };
}

async function imagePhysicalScale(file, raster) {
  const name = file.name || "";
  let physical = null;
  try {
    if (/\.png$/i.test(name) || file.type === "image/png") physical = pngPhysicalScale(await file.arrayBuffer(), raster);
    else if (/\.jpe?g$/i.test(name) || file.type === "image/jpeg") physical = jpegPhysicalScale(await file.arrayBuffer(), raster);
    else if (/\.svg$/i.test(name) || file.type === "image/svg+xml") physical = svgPhysicalScale(await file.text(), raster);
  } catch {
    // Fall back below to the browser/CSS image convention.
  }
  return physical || scaleFromOriginalPixel(25.4 / 96, raster, "96dpi-estimate");
}

function pngPhysicalScale(buffer, raster) {
  const view = new DataView(buffer);
  let offset = 8;
  while (offset + 12 <= view.byteLength) {
    const length = view.getUint32(offset, false);
    const type = asciiFromBuffer(buffer, offset + 4, 4);
    if (type === "pHYs" && offset + 21 <= view.byteLength) {
      const xPpm = view.getUint32(offset + 8, false);
      const unit = view.getUint8(offset + 16);
      if (unit === 1 && xPpm > 0) return scaleFromOriginalPixel(1000 / xPpm, raster, "png-physical-pixels");
    }
    offset += 12 + length;
  }
  return null;
}

function jpegPhysicalScale(buffer, raster) {
  const view = new DataView(buffer);
  let offset = 2;
  while (offset + 4 < view.byteLength) {
    if (view.getUint8(offset) !== 0xff) break;
    const marker = view.getUint8(offset + 1);
    const length = view.getUint16(offset + 2, false);
    if (marker === 0xe0 && asciiFromBuffer(buffer, offset + 4, 5) === "JFIF\0") {
      const unit = view.getUint8(offset + 11);
      const xDensity = view.getUint16(offset + 12, false);
      if (unit === 1 && xDensity > 0) return scaleFromOriginalPixel(25.4 / xDensity, raster, "jpeg-jfif-dpi");
      if (unit === 2 && xDensity > 0) return scaleFromOriginalPixel(10 / xDensity, raster, "jpeg-jfif-dpcm");
    }
    offset += 2 + length;
  }
  return null;
}

function svgPhysicalScale(text, raster) {
  const svg = new DOMParser().parseFromString(text, "image/svg+xml").documentElement;
  if (!svg || svg.nodeName.toLowerCase() !== "svg") return null;
  const widthMm = lengthToMm(svg.getAttribute("width"));
  if (!Number.isFinite(widthMm) || widthMm <= 0 || !raster.originalWidth) return null;
  return scaleFromOriginalPixel(widthMm / raster.originalWidth, raster, "svg-width");
}

function scaleFromOriginalPixel(mmPerOriginalPixel, raster, source) {
  if (!Number.isFinite(mmPerOriginalPixel) || mmPerOriginalPixel <= 0) return null;
  return {
    mmPerRasterPixel: mmPerOriginalPixel / (raster.scale || 1),
    source,
  };
}

function lengthToMm(value) {
  const match = String(value || "").trim().match(/^([+-]?\d*\.?\d+)\s*(mm|cm|in|pt|px)?$/i);
  if (!match) return null;
  const amount = Number(match[1]);
  const unit = (match[2] || "px").toLowerCase();
  if (!Number.isFinite(amount)) return null;
  if (unit === "mm") return amount;
  if (unit === "cm") return amount * 10;
  if (unit === "in") return amount * 25.4;
  if (unit === "pt") return (amount * 25.4) / 72;
  if (unit === "px") return (amount * 25.4) / 96;
  return null;
}

function asciiFromBuffer(buffer, offset, length) {
  return String.fromCharCode(...new Uint8Array(buffer, offset, length));
}

async function traceBmpFile(file) {
  const buffer = await file.arrayBuffer();
  const canvas = bmpToCanvas(buffer);
  const raster = rasterFromCanvas(canvas);
  const view = new DataView(buffer);
  const xPpm = view.getInt32(38, true);
  const traced = await traceRasterWithSelectedEngine(raster);
  const physical = xPpm > 0 ? { mmPerRasterPixel: 1000 / xPpm, source: "bmp-pixels-per-meter" } : scaleFromOriginalPixel(25.4 / 96, raster, "96dpi-estimate");
  return {
    ...withActualTraceWidth(traced, physical?.mmPerRasterPixel),
    raster,
    originalRaster: raster,
    mmPerRasterPixel: physical?.mmPerRasterPixel || 0,
    sourceType: "bmp",
    physicalSource: physical?.source || "",
  };
}

function bmpToCanvas(buffer) {
  const view = new DataView(buffer);
  if (view.getUint16(0, true) !== 0x4d42) throw new Error("Invalid BMP file");
  const pixelOffset = view.getUint32(10, true);
  const dibSize = view.getUint32(14, true);
  if (dibSize < 40) throw new Error("Unsupported BMP header");
  const width = view.getInt32(18, true);
  const rawHeight = view.getInt32(22, true);
  const height = Math.abs(rawHeight);
  const planes = view.getUint16(26, true);
  const bitDepth = view.getUint16(28, true);
  const compression = view.getUint32(30, true);
  if (planes !== 1 || compression !== 0 || ![24, 32].includes(bitDepth) || width <= 0 || height <= 0) {
    throw new Error("Only uncompressed 24-bit and 32-bit BMP files are supported");
  }
  const canvas = document.createElement("canvas");
  canvas.width = width;
  canvas.height = height;
  const ctx = canvas.getContext("2d", { willReadFrequently: true });
  const imageData = ctx.createImageData(width, height);
  const bytesPerPixel = bitDepth / 8;
  const rowSize = Math.floor((bitDepth * width + 31) / 32) * 4;
  const bottomUp = rawHeight > 0;
  for (let y = 0; y < height; y += 1) {
    const sourceY = bottomUp ? height - 1 - y : y;
    const row = pixelOffset + sourceY * rowSize;
    for (let x = 0; x < width; x += 1) {
      const source = row + x * bytesPerPixel;
      const target = (y * width + x) * 4;
      imageData.data[target] = view.getUint8(source + 2);
      imageData.data[target + 1] = view.getUint8(source + 1);
      imageData.data[target + 2] = view.getUint8(source);
      imageData.data[target + 3] = bitDepth === 32 ? view.getUint8(source + 3) : 255;
    }
  }
  ctx.putImageData(imageData, 0, 0);
  return canvas;
}

async function loadPdfJs() {
  if (window.pdfjsLib) return window.pdfjsLib;
  if (window.location.protocol === "file:") {
    throw new Error("PDF import requires the local server. Open http://127.0.0.1:8792/index.html instead of the HTML file directly.");
  }
  const localPdfjsUrl = new URL("assets/pdf.min.mjs", window.location.href).href;
  const localWorkerUrl = new URL("assets/pdf.worker.min.mjs", window.location.href).href;
  const cdnPdfjsUrl = "https://cdn.jsdelivr.net/npm/pdfjs-dist@4.10.38/build/pdf.min.mjs";
  const cdnWorkerUrl = "https://cdn.jsdelivr.net/npm/pdfjs-dist@4.10.38/build/pdf.worker.min.mjs";
  let pdfjs;
  try {
    pdfjs = await import(localPdfjsUrl);
    pdfjs.GlobalWorkerOptions.workerSrc = localWorkerUrl;
  } catch (error) {
    console.warn("Local PDF.js assets unavailable; loading PDF.js from CDN.", error);
    pdfjs = await import(cdnPdfjsUrl);
    pdfjs.GlobalWorkerOptions.workerSrc = cdnWorkerUrl;
  }
  window.pdfjsLib = pdfjs;
  return pdfjs;
}

function loadRasterImage(file) {
  return new Promise((resolve, reject) => {
    const image = new Image();
    const url = URL.createObjectURL(file);
    image.onload = () => {
      URL.revokeObjectURL(url);
      resolve(image);
    };
    image.onerror = () => {
      URL.revokeObjectURL(url);
      reject(new Error("Image could not be loaded"));
    };
    image.src = url;
  });
}

function grayscalePixels(data) {
  const gray = new Uint8Array(data.length / 4);
  for (let i = 0, p = 0; i < data.length; i += 4, p += 1) {
    const alpha = data[i + 3] / 255;
    const r = data[i] * alpha + 255 * (1 - alpha);
    const g = data[i + 1] * alpha + 255 * (1 - alpha);
    const b = data[i + 2] * alpha + 255 * (1 - alpha);
    gray[p] = Math.round(r * 0.299 + g * 0.587 + b * 0.114);
  }
  return gray;
}

function otsuThreshold(gray) {
  const hist = new Array(256).fill(0);
  gray.forEach((value) => {
    hist[value] += 1;
  });
  const total = gray.length;
  let sum = 0;
  hist.forEach((count, value) => {
    sum += value * count;
  });
  let sumB = 0;
  let weightB = 0;
  let best = 127;
  let bestVariance = -1;
  for (let value = 0; value < 256; value += 1) {
    weightB += hist[value];
    if (!weightB) continue;
    const weightF = total - weightB;
    if (!weightF) break;
    sumB += value * hist[value];
    const meanB = sumB / weightB;
    const meanF = (sum - sumB) / weightF;
    const variance = weightB * weightF * (meanB - meanF) ** 2;
    if (variance > bestVariance) {
      bestVariance = variance;
      best = value;
    }
  }
  return best;
}

function borderAverage(gray, width, height) {
  let total = 0;
  let count = 0;
  for (let x = 0; x < width; x += 1) {
    total += gray[x] + gray[(height - 1) * width + x];
    count += 2;
  }
  for (let y = 1; y < height - 1; y += 1) {
    total += gray[y * width] + gray[y * width + width - 1];
    count += 2;
  }
  return count ? total / count : 255;
}

function largestForegroundComponent(gray, width, height, threshold, foregroundIsDark) {
  return largestBinaryComponent(foregroundMask(gray, width, height, threshold, foregroundIsDark), width, height);
}

function foregroundMask(gray, width, height, threshold, foregroundIsDark) {
  const mask = new Uint8Array(width * height);
  for (let i = 0; i < gray.length; i += 1) {
    const isForeground = foregroundIsDark ? gray[i] <= threshold : gray[i] > threshold;
    mask[i] = isForeground ? 1 : 0;
  }
  return mask;
}

function traceMaskParts(mask, width, height) {
  const components = connectedBinaryComponents(mask, width, height);
  if (!components.length) return [];
  const largest = components[0].length;
  const minArea = Math.max(40, Math.round(mask.length * 0.00005), Math.round(largest * 0.015));
  const tolerance = getTraceSimplifyTolerance();
  return components
    .filter((component) => component.length >= minArea)
    .slice(0, 40)
    .flatMap((component) => boundaryLoops(maskFromComponent(component, mask.length), width, height))
    .filter((points) => Math.abs(signedArea(points)) >= minArea)
    .map((points) => simplifyPoints(points, tolerance))
    .filter((points) => points.length >= 3)
    .sort((a, b) => Math.abs(signedArea(b)) - Math.abs(signedArea(a)));
}

function maskFromComponent(component, size) {
  const result = new Uint8Array(size);
  component.forEach((index) => {
    result[index] = 1;
  });
  return result;
}

function connectedBinaryComponents(mask, width, height) {
  const visited = new Uint8Array(mask.length);
  const components = [];
  const queue = [];
  for (let start = 0; start < mask.length; start += 1) {
    if (!mask[start] || visited[start]) continue;
    const component = [];
    queue.length = 0;
    queue.push(start);
    visited[start] = 1;
    for (let q = 0; q < queue.length; q += 1) {
      const index = queue[q];
      component.push(index);
      const x = index % width;
      const y = Math.floor(index / width);
      const neighbors = [
        x > 0 ? index - 1 : -1,
        x < width - 1 ? index + 1 : -1,
        y > 0 ? index - width : -1,
        y < height - 1 ? index + width : -1,
      ];
      neighbors.forEach((next) => {
        if (next >= 0 && mask[next] && !visited[next]) {
          visited[next] = 1;
          queue.push(next);
        }
      });
    }
    components.push(component);
  }
  return components.sort((a, b) => b.length - a.length);
}

function boundaryLoop(mask, width, height) {
  return boundaryLoops(mask, width, height)[0] || [];
}

function boundaryLoops(mask, width, height) {
  const edges = [];
  const isSet = (x, y) => x >= 0 && x < width && y >= 0 && y < height && mask[y * width + x];
  for (let y = 0; y < height; y += 1) {
    for (let x = 0; x < width; x += 1) {
      if (!isSet(x, y)) continue;
      if (!isSet(x, y - 1)) edges.push(edgePoint(x, y, x + 1, y));
      if (!isSet(x + 1, y)) edges.push(edgePoint(x + 1, y, x + 1, y + 1));
      if (!isSet(x, y + 1)) edges.push(edgePoint(x + 1, y + 1, x, y + 1));
      if (!isSet(x - 1, y)) edges.push(edgePoint(x, y + 1, x, y));
    }
  }
  const loops = linkBoundaryEdges(edges);
  return loops.sort((a, b) => Math.abs(signedArea(b)) - Math.abs(signedArea(a)));
}

function edgePoint(x1, y1, x2, y2) {
  return { start: `${x1},${y1}`, end: `${x2},${y2}`, p1: { x: x1, y: y1 }, p2: { x: x2, y: y2 } };
}

function linkBoundaryEdges(edges) {
  const byStart = new Map();
  edges.forEach((edge) => {
    if (!byStart.has(edge.start)) byStart.set(edge.start, []);
    byStart.get(edge.start).push(edge);
  });
  const loops = [];
  while (edges.length) {
    const edge = edges.pop();
    removeStartEdge(byStart, edge);
    const loop = [edge.p1, edge.p2];
    let current = edge;
    while (current.end !== edge.start) {
      const next = byStart.get(current.end)?.pop();
      if (!next) break;
      const index = edges.indexOf(next);
      if (index !== -1) edges.splice(index, 1);
      loop.push(next.p2);
      current = next;
    }
    const cleaned = cleanClosingPoint(loop);
    if (cleaned.length >= 3) loops.push(cleaned);
  }
  return loops;
}

function removeStartEdge(map, edge) {
  const list = map.get(edge.start);
  if (!list) return;
  const index = list.indexOf(edge);
  if (index !== -1) list.splice(index, 1);
}

function imageTraceEntity(trace, widthMm) {
  const pixelParts = trace.parts?.length ? trace.parts.filter((points) => points.length >= 3) : (trace.points?.length >= 3 ? [trace.points] : []);
  if (!pixelParts.length) return { type: "LWPOLYLINE", layer: "AI_IMAGE_TRACE_SOURCE", points: [], closed: true };
  const box = bounds(pixelParts.map((points) => ({ points })));
  const pixelWidth = Math.max(1, box.maxX - box.minX);
  const mmPerPixel = widthMm / pixelWidth;
  const centerX = (box.minX + box.maxX) / 2;
  const centerY = (box.minY + box.maxY) / 2;
  const parts = pixelParts
    .map((points, index) => {
      const smoothed = smoothTracePart(points, getTraceSmoothWindow());
      return {
        type: "LWPOLYLINE",
        layer: `AI_IMAGE_TRACE_SOURCE_${index + 1}`,
        points: cleanCurvePoints(smoothed.map((point) => ({
          x: (point.x - centerX) * mmPerPixel,
          y: (centerY - point.y) * mmPerPixel,
        }))),
        closed: true,
      };
    })
    .filter((part) => part.points.length >= 3);
  if (parts.length === 1) return { ...parts[0], layer: "AI_IMAGE_TRACE_SOURCE" };
  return {
    type: "MULTIPATCH",
    layer: "AI_IMAGE_TRACE_SOURCE",
    parts,
    points: parts.flatMap((part) => part.points),
    closed: true,
  };
}

function imagePreviewEntity(trace, widthMm) {
  const raster = trace?.raster;
  if (!raster?.dataUrl) return null;
  const pixelParts = trace.parts?.length ? trace.parts.filter((points) => points.length >= 3) : (trace.points?.length >= 3 ? [trace.points] : []);
  const hasTrace = pixelParts.length > 0;
  const box = hasTrace
    ? bounds(pixelParts.map((points) => ({ points })))
    : { minX: 0, minY: 0, maxX: raster.width, maxY: raster.height };
  const pixelWidth = Math.max(1, hasTrace ? box.maxX - box.minX : raster.width);
  const mmPerPixel = widthMm / pixelWidth;
  const centerX = (box.minX + box.maxX) / 2;
  const centerY = (box.minY + box.maxY) / 2;
  return {
    type: "IMAGE",
    layer: "AI_SCAN_PREVIEW",
    href: raster.dataUrl,
    x: -centerX * mmPerPixel,
    y: -centerY * mmPerPixel,
    width: raster.width * mmPerPixel,
    height: raster.height * mmPerPixel,
    opacity: state.imagePreviewOpacity,
  };
}

function scanCropPreviewEntity(trace, crop, widthMm) {
  const image = imagePreviewEntity(trace, widthMm);
  const raster = trace?.raster;
  if (!image || !raster || !crop) return null;
  const mmPerPixelX = image.width / raster.width;
  const mmPerPixelY = image.height / raster.height;
  return {
    type: "RECT",
    layer: "AI_SCAN_CROP_PREVIEW",
    x: image.x + crop.x * mmPerPixelX,
    y: image.y + crop.y * mmPerPixelY,
    width: crop.width * mmPerPixelX,
    height: crop.height * mmPerPixelY,
  };
}

function simplifyPoints(points, tolerance) {
  const cleaned = cleanClosingPoint(cleanCurvePoints(points));
  if (cleaned.length <= 3) return cleaned;
  const simplified = douglasPeucker(cleaned.concat(cleaned[0]), tolerance);
  return cleanClosingPoint(simplified);
}

function getTraceSmoothWindow() {
  const smooth = getTraceSmoothValue();
  if (!smooth) return 0;
  return Math.max(1, Math.round(1 + smooth * 0.35));
}

function smoothTracePart(points, windowSize) {
  const cleaned = cleanClosingPoint(cleanCurvePoints(points));
  if (cleaned.length < 8 || windowSize <= 0) return cleaned;
  const passes = Math.max(1, Math.round(getTraceSmoothValue() / 3));
  let current = cleaned;
  for (let pass = 0; pass < passes; pass += 1) {
    current = smoothTracePass(current, windowSize);
  }
  return cleanCurvePoints(current);
}

function smoothTracePass(points, windowSize) {
  const result = [];
  for (let i = 0; i < points.length; i += 1) {
    const prev = points[(i - 1 + points.length) % points.length];
    const point = points[i];
    const next = points[(i + 1) % points.length];
    const angle = turnAngle(prev, point, next);
    if (angle < 128) {
      result.push(point);
      continue;
    }
    let x = 0;
    let y = 0;
    let weightTotal = 0;
    for (let offset = -windowSize; offset <= windowSize; offset += 1) {
      const neighbor = points[(i + offset + points.length) % points.length];
      const weight = windowSize + 1 - Math.abs(offset);
      x += neighbor.x * weight;
      y += neighbor.y * weight;
      weightTotal += weight;
    }
    result.push({ x: x / weightTotal, y: y / weightTotal });
  }
  return result;
}

function turnAngle(prev, point, next) {
  const ax = prev.x - point.x;
  const ay = prev.y - point.y;
  const bx = next.x - point.x;
  const by = next.y - point.y;
  const la = Math.hypot(ax, ay) || 1;
  const lb = Math.hypot(bx, by) || 1;
  const dot = Math.max(-1, Math.min(1, (ax * bx + ay * by) / (la * lb)));
  return (Math.acos(dot) * 180) / Math.PI;
}

function douglasPeucker(points, tolerance) {
  if (points.length <= 2) return points;
  let maxDistance = 0;
  let index = 0;
  const first = points[0];
  const last = points[points.length - 1];
  for (let i = 1; i < points.length - 1; i += 1) {
    const distance = pointLineDistance(points[i], first, last);
    if (distance > maxDistance) {
      maxDistance = distance;
      index = i;
    }
  }
  if (maxDistance <= tolerance) return [first, last];
  return douglasPeucker(points.slice(0, index + 1), tolerance).slice(0, -1).concat(douglasPeucker(points.slice(index), tolerance));
}

function pointLineDistance(point, a, b) {
  const dx = b.x - a.x;
  const dy = b.y - a.y;
  if (!dx && !dy) return Math.hypot(point.x - a.x, point.y - a.y);
  const t = Math.max(0, Math.min(1, ((point.x - a.x) * dx + (point.y - a.y) * dy) / (dx * dx + dy * dy)));
  return Math.hypot(point.x - (a.x + t * dx), point.y - (a.y + t * dy));
}

function countDxfEntityTypes(text) {
  const pairs = text.replace(/\r/g, "").split("\n").map((line) => line.trim());
  const counts = {};
  let inEntities = false;
  for (let i = 0; i < pairs.length - 3; i += 2) {
    if (pairs[i] === "0" && pairs[i + 1] === "SECTION" && pairs[i + 2] === "2" && pairs[i + 3] === "ENTITIES") {
      inEntities = true;
      i += 2;
      continue;
    }
    if (inEntities && pairs[i] === "0" && pairs[i + 1] === "ENDSEC") break;
    if (inEntities && pairs[i] === "0") counts[pairs[i + 1]] = (counts[pairs[i + 1]] || 0) + 1;
  }
  return counts;
}

function getFormula() {
  const learned = state.learnedProfile || {};
  return {
    boardWidth: Number(els.boardWidth.value) || DEFAULT_FORMULA.boardWidth,
    boardHeight: Number(els.boardHeight.value) || DEFAULT_FORMULA.boardHeight,
    cornerRadius: Number(els.cornerRadius.value) || DEFAULT_FORMULA.cornerRadius,
    slotWidth: Number(els.slotWidth.value) || DEFAULT_FORMULA.slotWidth,
    slotHeight: Number(els.slotHeight.value) || DEFAULT_FORMULA.slotHeight,
    slotRightEdge: Number(els.slotRightEdge.value) || DEFAULT_FORMULA.slotRightEdge,
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

function ovalPoints(cx, cy, width, height, segments = 96) {
  const rx = width / 2;
  const ry = height / 2;
  if (!rx || !ry) return [];
  return Array.from({ length: segments }, (_, index) => {
    const angle = (Math.PI * 2 * index) / segments;
    return { x: cx + Math.cos(angle) * rx, y: cy + Math.sin(angle) * ry };
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
  return ["LWPOLYLINE", "SPLINE", "CIRCLE", "ARC", "ELLIPSE", "HATCH", "MULTIPATCH"].includes(entity.type);
}

function generatePatchTemplate(entities, formula) {
  const outline = selectPatchOutline(entities, formula);
  if (!outline) return [];
  const basePatchParts = centeredPatchParts(outline, formula.patchRotation);
  const baseOffsetParts = centeredOffsetParts(outline, formula.patchRotation);
  const startX = 70;
  const gap = formula.boardWidth + 16;
  const startY = 0;
  const generated = [];
  const sourcePreviewY = startY + formula.boardHeight / 2 + 55;

  generated.push({
    type: "TEXT",
    layer: "AI_LABEL",
    text: state.sourceType === "image-trace" ? "Scan trace" : "Patch DXF",
    point: { x: startX - formula.boardWidth / 2, y: sourcePreviewY + 18 },
    previewOnly: true,
  });
  generated.push({
    type: "LWPOLYLINE",
    layer: "AI_SOURCE_PATCH",
    points: movePoints(basePatchParts[0], startX, sourcePreviewY),
    closed: true,
    previewOnly: true,
  });
  basePatchParts.slice(1).forEach((part) => {
    generated.push({
      type: "LWPOLYLINE",
      layer: "AI_SOURCE_PATCH",
      points: movePoints(part, startX, sourcePreviewY),
      closed: true,
      previewOnly: true,
    });
  });

  for (let layer = 1; layer <= 4; layer += 1) {
    const origin = { x: startX + (layer - 1) * gap, y: startY };
    const board = templateBoard(origin, formula);
    const patchParts = basePatchParts.map((part) => movePoints(part, origin.x + formula.patchRelX, origin.y + formula.patchRelY));
    const offsetSourceParts = baseOffsetParts.map((part) => movePoints(part, origin.x + formula.patchRelX, origin.y + formula.patchRelY));
    const offsetParts = unionOffsetParts(
      offsetSourceParts
        .map((part) => offsetOrthogonalPolygon(part, formula.offset, "outward"))
        .filter((part) => part.length >= 3),
    );
    const slotCenter = needleSlotCenter(origin, formula);
    generated.push(
      { type: "TEXT", layer: "AI_LABEL", text: String(layer), point: { x: origin.x, y: origin.y + formula.boardHeight / 2 + 18 }, height: 8 },
      { type: "LWPOLYLINE", layer: `AI_LAYER_${layer}_BOARD`, points: board, closed: true },
      { type: "LWPOLYLINE", layer: `AI_LAYER_${layer}_NEEDLE_SLOT`, points: pillPoints(slotCenter.x, slotCenter.y, formula.slotWidth, formula.slotHeight), closed: true },
    );
    patchParts.forEach((patch, index) => {
      generated.push({ type: "LWPOLYLINE", layer: `AI_LAYER_${layer}_PATCH${index ? `_${index + 1}` : ""}`, points: patch, closed: true });
    });
    if (layer === 4) {
      generated.push({
        type: "TEXT",
        layer: "AI_LAYER_4_TEMPLATE_NUMBER",
        text: formula.templateNumber,
        point: { x: (origin.x - formula.boardWidth / 2 + slotCenter.x - formula.slotWidth / 2) / 2, y: slotCenter.y },
        height: 3.5,
        anchor: "middle",
      });
    }
    if (layer >= 2) {
      offsetParts.forEach((offsetPatch, index) => {
        generated.push({ type: "LWPOLYLINE", layer: `AI_LAYER_${layer}_PATCH_OFFSET_7MM${index ? `_${index + 1}` : ""}`, points: offsetPatch, closed: true });
      });
    }
  }
  return generated;
}

function unionOffsetParts(parts) {
  if (parts.length < 2) return parts;
  const union = polygonUnionBoundary(parts);
  return union.length >= 3 ? [union] : parts;
}

function centeredPatchParts(outline, rotationDeg) {
  const parts = outline.parts?.length ? outline.parts.map((part) => part.points) : [outline.points];
  const box = bounds(parts.map((points) => ({ points })));
  const center = { x: (box.minX + box.maxX) / 2, y: (box.minY + box.maxY) / 2 };
  return parts.map((points) => rotatePoints(movePoints(points, -center.x, -center.y), rotationDeg));
}

function centeredOffsetParts(outline, rotationDeg) {
  const parts = centeredPatchParts(outline, rotationDeg).filter((points) => points.length >= 3);
  if (parts.length <= 1) return parts;
  const envelope = offsetEnvelopePart(parts);
  if (envelope.length >= 3) return [envelope];
  const ranked = parts
    .map((points) => ({ points, area: polygonArea(points), box: bounds([{ points }]) }))
    .sort((a, b) => b.area - a.area);
  const largestArea = ranked[0]?.area || 0;
  const totalArea = ranked.reduce((sum, part) => sum + part.area, 0);
  const minArea = Math.max(largestArea * 0.08, totalArea * 0.025, 4);
  const selected = ranked
    .filter((part) => part.area >= minArea)
    .filter((part) => {
      const width = part.box.maxX - part.box.minX;
      const height = part.box.maxY - part.box.minY;
      return width >= 2 && height >= 2;
    })
    .slice(0, 8)
    .map((part) => part.points);
  return selected.length ? selected : [ranked[0].points];
}

function offsetEnvelopePart(parts) {
  const points = parts.flat();
  if (points.length < 3) return [];
  const hull = convexHull(points);
  if (hull.length < 3) return [];
  return roundedHullPoints(hull, 2.5);
}

function roundedHullPoints(hull, cornerRadius) {
  const result = [];
  const count = hull.length;
  for (let i = 0; i < count; i += 1) {
    const prev = hull[(i - 1 + count) % count];
    const point = hull[i];
    const next = hull[(i + 1) % count];
    const d1 = Math.hypot(point.x - prev.x, point.y - prev.y) || 1;
    const d2 = Math.hypot(next.x - point.x, next.y - point.y) || 1;
    const r = Math.min(cornerRadius, d1 * 0.3, d2 * 0.3);
    const a = { x: point.x + ((prev.x - point.x) / d1) * r, y: point.y + ((prev.y - point.y) / d1) * r };
    const b = { x: point.x + ((next.x - point.x) / d2) * r, y: point.y + ((next.y - point.y) / d2) * r };
    result.push(a);
    for (let step = 1; step <= 3; step += 1) {
      const t = step / 4;
      result.push(quadraticBezierPoint(a, point, b, t));
    }
    result.push(b);
  }
  return cleanCurvePoints(result);
}

function quadraticBezierPoint(a, b, c, t) {
  const mt = 1 - t;
  return {
    x: mt * mt * a.x + 2 * mt * t * b.x + t * t * c.x,
    y: mt * mt * a.y + 2 * mt * t * b.y + t * t * c.y,
  };
}

function selectPatchOutline(entities, formula) {
  const simpleShape = entities.find((entity) => entity.layer === "AI_SIMPLE_PATCH_SOURCE" && isShapeEntity(entity) && entity.points.length >= 3);
  if (simpleShape) return simpleShape;
  const tracedMultiPatch = entities.find((entity) => entity.type === "MULTIPATCH" && entity.parts?.length);
  if (tracedMultiPatch) return tracedMultiPatch;
  const candidates = entities
    .filter((entity) => isShapeEntity(entity) && entity.points.length >= 3)
    .map((entity) => ({ entity, box: bounds([entity]), area: polygonArea(entity.points) }))
    .filter((item) => item.area > 100)
    .filter((item) => !isBoardSized(item.box, formula))
    .sort((a, b) => b.area - a.area);
  const unique = uniqueShapeCandidates(candidates);
  const group = selectPatchCandidateGroup(unique);
  if (group?.length > 1) {
    return {
      type: "MULTIPATCH",
      layer: "AI_MULTI_PATCH_SOURCE",
      parts: group.map((candidate) => candidate.entity),
      points: group.flatMap((candidate) => candidate.entity.points),
      closed: true,
    };
  }
  const hatchUnion = combineOverlappingHatches(unique);
  if (hatchUnion) return hatchUnion;
  const joined = joinConnectedPatchCandidates(unique);
  if (joined) return joined;
  return unique.filter((item) => !isNeedleSlotSized(item.box, formula))[0]?.entity || largestClosedShape(entities);
}

function selectPatchCandidateGroup(candidates) {
  if (candidates.length < 2) return null;
  const clusters = [];
  candidates.forEach((candidate) => {
    let cluster = clusters.find((items) => items.some((item) => candidateDistance(item, candidate) <= 12));
    if (!cluster) {
      cluster = [];
      clusters.push(cluster);
    }
    cluster.push(candidate);
  });
  const usable = clusters
    .filter((cluster) => cluster.length > 1)
    .map((cluster) => ({
      cluster,
      box: bounds(cluster.map((candidate) => candidate.entity)),
      area: cluster.reduce((sum, candidate) => sum + Math.abs(candidate.area), 0),
    }))
    .filter((item) => item.box.maxX - item.box.minX <= 120 && item.box.maxY - item.box.minY <= 120)
    .sort((a, b) => {
      const centerA = { x: (a.box.minX + a.box.maxX) / 2, y: (a.box.minY + a.box.maxY) / 2 };
      const centerB = { x: (b.box.minX + b.box.maxX) / 2, y: (b.box.minY + b.box.maxY) / 2 };
      return Math.hypot(centerA.x, centerA.y) - Math.hypot(centerB.x, centerB.y) || b.area - a.area;
    });
  return usable[0]?.cluster || null;
}

function candidateDistance(a, b) {
  const ac = { x: (a.box.minX + a.box.maxX) / 2, y: (a.box.minY + a.box.maxY) / 2 };
  const bc = { x: (b.box.minX + b.box.maxX) / 2, y: (b.box.minY + b.box.maxY) / 2 };
  const gapX = Math.max(0, Math.max(a.box.minX, b.box.minX) - Math.min(a.box.maxX, b.box.maxX));
  const gapY = Math.max(0, Math.max(a.box.minY, b.box.minY) - Math.min(a.box.maxY, b.box.maxY));
  return Math.hypot(gapX, gapY) || Math.hypot(ac.x - bc.x, ac.y - bc.y) * 0.05;
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
        slotRightEdge: board.box.maxX - slot.box.maxX,
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
    slotRightEdge: round(avg(samples.map((sample) => sample.slotRightEdge))),
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

function round1(value) {
  return Math.round(value * 10) / 10;
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
    const examples = JSON.parse(localStorage.getItem("patchTemplateTrainingExamples") || "[]");
    return Array.isArray(examples) ? examples.map(normalizeTrainingExample) : [];
  } catch {
    return [];
  }
}

function saveTrainingExamples(examples) {
  try {
    localStorage.setItem("patchTemplateTrainingExamples", JSON.stringify(examples.map(normalizeTrainingExample)));
  } catch {
    // localStorage is unavailable in some automated checks.
  }
}

function normalizeTrainingExample(example) {
  if (!example || typeof example !== "object") return example;
  const formula = example.formula || {};
  const output = example.output || {};
  const templateNumberLayer = Number(formula.templateNumberLayer || output.templateNumberLayer || 4);
  const offsetGeneration = output.offsetGeneration || {
    outwardOffsetMm: formula.outwardOffsetMm ?? formula.offset ?? 7,
    strategy: "legacy-unspecified",
  };
  return {
    ...example,
    formula: {
      ...formula,
      templateNumberLayer,
      templateNumberPlacement: formula.templateNumberPlacement || {
        layer: templateNumberLayer,
        reference: "between_template_left_edge_and_needle_slot",
      },
    },
    output: {
      ...output,
      templateNumberLayer,
      offsetGeneration,
    },
  };
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
  const outlineParts = outline.parts?.length ? outline.parts : [outline];
  const outlineBox = bounds(outlineParts);
  const offsetGeneration = describeOffsetGeneration(outline, formula);
  const types = state.entities.reduce((acc, entity) => {
    acc[entity.type] = (acc[entity.type] || 0) + 1;
    return acc;
  }, {});
  return {
    schema: "patch-template-example-v1",
    savedAt: new Date().toISOString(),
    input: {
      sourceType: state.sourceType || "dxf",
      fileName: state.fileName,
      entityCount: state.entities.length,
      entityTypes: types,
      simpleShape: state.simpleShape,
      scanTrace: scanTraceFeatureInput(),
    },
    selectedPatch: {
      layer: outline.layer,
      type: outline.type,
      partCount: outlineParts.length,
      pointCount: outlineParts.reduce((sum, part) => sum + part.points.length, 0),
      widthMm: round(outlineBox.maxX - outlineBox.minX),
      heightMm: round(outlineBox.maxY - outlineBox.minY),
      areaMm2: round(outlineParts.reduce((sum, part) => sum + Math.abs(polygonArea(part.points)), 0)),
    },
    formula: {
      boardWidthMm: formula.boardWidth,
      boardHeightMm: formula.boardHeight,
      cornerRadiusMm: formula.cornerRadius,
      needleSlotWidthMm: formula.slotWidth,
      needleSlotHeightMm: formula.slotHeight,
      needleSlotRightEdgeMm: formula.slotRightEdge,
      needleSlot: {
        referenceEdge: "right",
        edgeDistanceMm: formula.slotRightEdge,
        widthMm: formula.slotWidth,
        heightMm: formula.slotHeight,
      },
      outwardOffsetMm: formula.offset,
      patchRelX: formula.patchRelX,
      patchRelY: formula.patchRelY,
      patchRotationDeg: formula.patchRotation,
      templateNumber: formula.templateNumber,
      templateNumberLayer: 4,
      templateNumberPlacement: {
        layer: 4,
        reference: "between_template_left_edge_and_needle_slot",
      },
    },
    output: {
      templateLayers: 4,
      templateNumberLayer: 4,
      offsetGeneration,
      generatedEntityCount: state.generated.length || generatePatchTemplate(state.entities, formula).length,
      ruleSet: "patch-template-4-layer-v1",
    },
  };
}

function describeOffsetGeneration(outline, formula) {
  const baseParts = centeredOffsetParts(outline, formula.patchRotation);
  const partStrategies = baseParts.map((points, index) => {
    const box = bounds([{ points }]);
    const width = box.maxX - box.minX;
    const height = box.maxY - box.minY;
    const area = polygonArea(points);
    const compactRounded = Boolean(offsetCompactRoundedShape(points, formula.offset, "outward"));
    const circular = !compactRounded && Boolean(offsetCircularShape(points, formula.offset, "outward"));
    return {
      partIndex: index + 1,
      strategy: compactRounded ? "compact-rounded-convex-offset" : circular ? "circular-radial-offset" : "line-joined-offset",
      pointCount: points.length,
      widthMm: round(width),
      heightMm: round(height),
      areaMm2: round(area),
      fillRatio: width > 0 && height > 0 ? round(area / (width * height)) : 0,
    };
  });
  const rawOffsetParts = baseParts
    .map((points) => offsetOrthogonalPolygon(points, formula.offset, "outward"))
    .filter((points) => points.length >= 3);
  const unionedOffsetParts = unionOffsetParts(rawOffsetParts);
  return {
    outwardOffsetMm: formula.offset,
    sourcePartCount: baseParts.length,
    rawOffsetPartCount: rawOffsetParts.length,
    outputOffsetPartCount: unionedOffsetParts.length,
    unionApplied: rawOffsetParts.length > 1 && unionedOffsetParts.length < rawOffsetParts.length,
    partStrategies,
  };
}

function trainingPayload() {
  const updatedAt = state.trainingMeta.updatedAt || new Date().toISOString();
  const payload = {
    schema: "patch-template-ai-training-v1",
    revision: state.trainingMeta.revision || 0,
    updatedAt,
    exportedAt: updatedAt,
    exampleCount: state.trainingExamples.length,
    source: "patch-template-web",
    examples: state.trainingExamples,
    learnedProfile: state.learnedProfile,
  };
  if (state.patchModel) {
    payload.model = {
      ...serializablePatchModel(state.patchModel),
      trainingDataFileName: "patch-template-training-data.json",
      exampleCount: state.patchModel.exampleCount || state.trainingExamples.length,
    };
  }
  return payload;
}

function serializablePatchModel(model) {
  const { trainingData, ...modelOnly } = model;
  return modelOnly;
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
    state.trainingExamples = payload.examples.map(normalizeTrainingExample);
    state.trainingMeta = {
      revision: Number(payload.revision) || state.trainingMeta.revision || 0,
      updatedAt: payload.updatedAt || payload.exportedAt || state.trainingMeta.updatedAt || "",
    };
    state.learnedProfile = payload.learnedProfile || state.learnedProfile;
    applyModelFromTrainingPayload(payload);
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
  if (responseBody.trainingData && Array.isArray(responseBody.trainingData.examples)) return responseBody.trainingData;
  return null;
}

function applyModelFromTrainingPayload(payload) {
  const model = normalizePatchModelPayload(payload);
  try {
    validatePatchModel(model);
    state.patchModel = model;
    state.patchModelStatus = `Loaded ${model.exampleCount || payload.examples?.length || 0} patch training example(s).`;
  } catch {
    // The OneDrive JSON may be training-only before the first PyTorch export.
  }
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
  if (profile.slotRightEdge != null) els.slotRightEdge.value = profile.slotRightEdge;
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
    x: origin.x + formula.boardWidth / 2 - formula.slotRightEdge - formula.slotWidth / 2,
    y: origin.y + formula.boardHeight / 2 - 50 - formula.slotHeight / 2,
  };
}

function offsetOrthogonalPolygon(points, amount, direction = "inward") {
  const circularOffset = offsetCircularShape(points, amount, direction);
  if (circularOffset) return circularOffset;
  const compactOffset = offsetCompactRoundedShape(points, amount, direction);
  if (compactOffset) return compactOffset;
  return offsetLineJoinedPolygon(points, amount, direction);
}

function offsetLineJoinedPolygon(points, amount, direction = "inward") {
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

function offsetCompactRoundedShape(points, amount, direction) {
  if (direction !== "outward" || points.length < 7) return null;
  const box = bounds([{ points }]);
  const width = box.maxX - box.minX;
  const height = box.maxY - box.minY;
  if (width <= 0 || height <= 0) return null;
  const ratio = Math.max(width, height) / Math.min(width, height);
  const fillRatio = polygonArea(points) / (width * height);
  const largeOffsetForShape = amount / Math.min(width, height) > 0.18;
  if (ratio > (largeOffsetForShape ? 2.2 : 1.25)) return null;
  if (!largeOffsetForShape && fillRatio < 0.62) return null;
  const hull = convexHull(points);
  if (hull.length < 3 || hull.length > points.length * (largeOffsetForShape ? 0.95 : 0.7)) return null;
  return roundOffsetConvexPolygon(hull, amount, direction);
}

function convexHull(points) {
  const unique = [];
  const seen = new Set();
  points.forEach((point) => {
    const key = `${round(point.x)}:${round(point.y)}`;
    if (seen.has(key)) return;
    seen.add(key);
    unique.push(point);
  });
  unique.sort((a, b) => a.x - b.x || a.y - b.y);
  if (unique.length <= 3) return unique;
  const lower = [];
  unique.forEach((point) => {
    while (lower.length >= 2 && cross(lower[lower.length - 2], lower[lower.length - 1], point) <= 0) lower.pop();
    lower.push(point);
  });
  const upper = [];
  [...unique].reverse().forEach((point) => {
    while (upper.length >= 2 && cross(upper[upper.length - 2], upper[upper.length - 1], point) <= 0) upper.pop();
    upper.push(point);
  });
  lower.pop();
  upper.pop();
  return [...lower, ...upper];
}

function cross(a, b, c) {
  return (b.x - a.x) * (c.y - a.y) - (b.y - a.y) * (c.x - a.x);
}

function roundOffsetConvexPolygon(points, amount, direction) {
  const area = signedArea(points);
  const inwardSign = area >= 0 ? 1 : -1;
  const directionSign = direction === "outward" ? -1 : 1;
  const offsetEdges = [];
  for (let i = 0; i < points.length; i += 1) {
    const p1 = points[i];
    const p2 = points[(i + 1) % points.length];
    const dx = p2.x - p1.x;
    const dy = p2.y - p1.y;
    const len = Math.hypot(dx, dy) || 1;
    const nx = (-dy / len) * inwardSign * directionSign;
    const ny = (dx / len) * inwardSign * directionSign;
    offsetEdges.push({
      start: { x: p1.x + nx * amount, y: p1.y + ny * amount },
      end: { x: p2.x + nx * amount, y: p2.y + ny * amount },
    });
  }
  const result = [];
  for (let i = 0; i < points.length; i += 1) {
    const prev = offsetEdges[(i - 1 + offsetEdges.length) % offsetEdges.length];
    const curr = offsetEdges[i];
    result.push(...arcJoinPoints(points[i], prev.end, curr.start, amount));
  }
  return cleanCurvePoints(result);
}

function arcJoinPoints(center, start, end, radius) {
  let startAngle = Math.atan2(start.y - center.y, start.x - center.x);
  let endAngle = Math.atan2(end.y - center.y, end.x - center.x);
  let sweep = endAngle - startAngle;
  while (sweep > Math.PI) sweep -= Math.PI * 2;
  while (sweep < -Math.PI) sweep += Math.PI * 2;
  const steps = Math.max(3, Math.ceil(Math.abs(sweep) / (Math.PI / 18)));
  return Array.from({ length: steps + 1 }, (_, index) => {
    const angle = startAngle + (sweep * index) / steps;
    return { x: center.x + Math.cos(angle) * radius, y: center.y + Math.sin(angle) * radius };
  });
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
    if (entity.type === "RECT") {
      points.push(
        { x: entity.x, y: -entity.y },
        { x: entity.x + entity.width, y: -(entity.y + entity.height) },
      );
    }
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
  if (els.scanAdjustPanel) els.scanAdjustPanel.classList.toggle("hidden", state.sourceType !== "image-trace");
  const hasScanRaster = state.sourceType === "image-trace" && Boolean(state.imageTrace?.raster);
  if (els.scanCropButton) els.scanCropButton.disabled = !hasScanRaster || Boolean(state.pendingScanCrop);
  if (els.scanConfirmCropButton) els.scanConfirmCropButton.classList.toggle("hidden", !state.pendingScanCrop);
  if (els.scanCancelCropButton) els.scanCancelCropButton.classList.toggle("hidden", !state.pendingScanCrop);
  if (els.scanResetCropButton) els.scanResetCropButton.disabled = !state.imageTrace?.crop;
  if (state.importStatus) {
    els.fileSummary.textContent = state.importStatus;
  } else if (state.importError) {
    els.fileSummary.textContent = `${state.fileName}: import failed - ${state.importError}`;
  } else if (state.entities.length && state.sourceType === "simple-shape") {
    els.fileSummary.textContent = `${simpleShapeLabel(state.simpleShape)}: simple shape generated`;
  } else if (state.sourceType === "image-trace" && state.entities.length) {
    const pointCount = state.entities[0]?.points?.length || 0;
    const sizeNote = traceSizeStatus(state.imageTrace);
    const engineNote = traceEngineLabel(state.imageTrace);
    els.fileSummary.textContent = state.generated.length
      ? `${state.fileName}: template generated from ${pointCount} ${engineNote} trace points`
      : `${state.fileName}: compare original with ${pointCount} ${engineNote} trace points, ${sizeNote}, then apply`;
  } else if (state.sourceType === "image-trace" && state.fileName) {
    els.fileSummary.textContent = `${state.fileName}: original imported, no closed trace found; adjust sensitivity/detail`;
  } else if (state.fileName && !state.entities.length) {
    const unsupported = Object.entries(state.rawEntityTypes || {})
      .map(([type, count]) => `${type} ${count}`)
      .join(", ");
    els.fileSummary.textContent = unsupported
      ? `${state.fileName}: no patch vector found (${unsupported})`
      : `${state.fileName}: no supported patch vector found`;
  } else {
    els.fileSummary.textContent = state.entities.length
      ? `${state.fileName}: ${state.entities.length} entities parsed`
      : "Import patch DXF or scan image to auto generate template";
  }
  if (els.issueList) renderIssues();
  renderPreview();
}

function traceEngineLabel(trace) {
  if (trace?.engine === "potrace-browser") return "Potrace";
  return "built-in";
}

function traceSizeStatus(trace) {
  if (!trace?.actualTraceWidthMm) return "manual width ready";
  if (trace.physicalSource === "96dpi-estimate") return "size estimated from image pixels";
  return "actual size detected";
}

function simpleShapeLabel(shape) {
  if (!shape) return "Simple patch";
  if (shape.type === "circle") {
    const label = shape.diameterMm ? "Circle" : "Oval";
    return `${label} ${shape.widthMm} x ${shape.heightMm}mm`;
  }
  return `Rectangle ${shape.widthMm} x ${shape.heightMm}mm`;
}

function renderDxfSize() {
  const outline = state.entities.length ? selectPatchOutline(state.entities, getFormula()) : null;
  if (!outline) {
    if (state.sourceType === "image-trace" && state.imageTrace?.raster) {
      const image = imagePreviewEntity(state.imageTrace, getScanTraceWidth());
      els.dxfSize.textContent = image ? `${round(image.width)} x ${round(image.height)} mm` : "0 x 0 mm";
    } else {
      els.dxfSize.textContent = "0 x 0 mm";
    }
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
        `Needle slot top edge is 50mm from template top edge and right edge is ${getFormula().slotRightEdge}mm from template right edge.`,
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
  const fallbackPreview = drawableEntities(state.entities);
  const previewSource = sourcePreview.length ? sourcePreview : fallbackPreview;
  const visible = state.generated.length ? state.generated : previewSource;
  const box = expandBox(bounds(visible), 30);
  const width = Math.max(1, box.maxX - box.minX);
  const height = Math.max(1, box.maxY - box.minY);
  state.baseViewBox = { x: box.minX, y: -box.maxY, width, height };
  if (!state.viewBox) state.viewBox = { ...state.baseViewBox };
  els.preview.setAttribute("viewBox", `${state.viewBox.x} ${state.viewBox.y} ${state.viewBox.width} ${state.viewBox.height}`);
  setZoomButtons(Boolean(visible.length));
  els.preview.innerHTML = "";
  const rendered = new Set();
  if (state.importStatus) {
    drawPreviewMessage(state.importStatus);
    return;
  }
  if (state.importError) {
    drawPreviewMessage(`Import failed: ${state.importError}`);
    return;
  }
  if (!visible.length && state.fileName) {
    drawPreviewMessage(`${state.fileName}: no drawable vector found`);
    return;
  }
  if (!state.generated.length) previewSource.forEach((entity) => drawEntity(entity, "source-line", rendered));
  state.generated.forEach((entity) => drawEntity(entity, generatedClass(entity.layer), rendered));
}

function drawableEntities(entities) {
  return entities.filter((entity) => isShapeEntity(entity) || entity.type === "LINE" || entity.type === "TEXT" || entity.type === "IMAGE" || entity.type === "RECT");
}

function selectedSourcePreviewEntities() {
  if (state.sourceType === "image-trace" && state.imageTrace) {
    const image = imagePreviewEntity(state.imageTrace, getScanTraceWidth());
    const outline = state.entities.length ? selectPatchOutline(state.entities, getFormula()) : null;
    const crop = scanCropPreviewEntity(state.imageTrace, state.pendingScanCrop?.crop, getScanTraceWidth());
    return [image, outline && { ...outline, layer: "AI_SELECTED_SOURCE_PREVIEW" }, crop].filter(Boolean);
  }
  if (!state.entities.length) return [];
  const outline = selectPatchOutline(state.entities, getFormula());
  if (!outline) return state.entities;
  if (outline.parts?.length) {
    return outline.parts.map((part, index) => ({
      ...part,
      layer: `AI_SELECTED_SOURCE_PREVIEW_${index + 1}`,
    }));
  }
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

function drawEntity(entity, className, rendered = null) {
  if (entity.type === "MULTIPATCH" && entity.parts?.length) {
    entity.parts.forEach((part) => drawEntity(part, className, rendered));
    return;
  }
  const key = rendered ? previewEntityKey(entity, className) : "";
  if (key && rendered.has(key)) return;
  if (key) rendered.add(key);

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
  if (entity.type === "IMAGE") {
    const image = svgEl("image", {
      href: entity.href,
      x: entity.x,
      y: entity.y,
      width: entity.width,
      height: entity.height,
      opacity: entity.opacity,
      class: "scan-preview-image",
      preserveAspectRatio: "none",
    });
    els.preview.prepend(image);
  }
  if (entity.type === "RECT") {
    const rect = svgEl("rect", {
      x: entity.x,
      y: entity.y,
      width: entity.width,
      height: entity.height,
      class: "scan-crop-preview",
    });
    els.preview.appendChild(rect);
  }
}

function drawPreviewMessage(message) {
  const box = state.viewBox || { x: 0, y: 0, width: 100, height: 60 };
  const text = svgEl("text", {
    x: box.x + box.width / 2,
    y: box.y + box.height / 2,
    class: "preview-message",
    "text-anchor": "middle",
  });
  text.textContent = message;
  els.preview.appendChild(text);
}

function previewEntityKey(entity, className) {
  if (entity.points?.length) {
    return `${className}:${entity.closed ? 1 : 0}:${entity.points.map(previewPointKey).join("|")}`;
  }
  if (entity.type === "LINE" && entity.points?.length === 2) {
    return `${className}:line:${entity.points.map(previewPointKey).join("|")}`;
  }
  return "";
}

function previewPointKey(point) {
  return `${round(point.x)}:${round(point.y)}`;
}

function svgEl(tag, attrs) {
  const el = document.createElementNS("http://www.w3.org/2000/svg", tag);
  Object.entries(attrs).forEach(([key, value]) => el.setAttribute(key, value));
  return el;
}

function resetView() {
  state.baseViewBox = null;
  state.viewBox = null;
  state.pan = null;
  els.preview.classList.remove("is-panning");
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

function startPreviewPan(event) {
  if (!state.viewBox) return;
  els.preview.setPointerCapture?.(event.pointerId);
  state.pan = {
    pointerId: event.pointerId,
    clientX: event.clientX,
    clientY: event.clientY,
    viewBox: { ...state.viewBox },
  };
  els.preview.classList.add("is-panning");
}

function movePreviewPan(event) {
  if (!state.pan || state.pan.pointerId !== event.pointerId || !state.viewBox) return;
  const rect = els.preview.getBoundingClientRect();
  if (!rect.width || !rect.height) return;
  const dx = (event.clientX - state.pan.clientX) * (state.pan.viewBox.width / rect.width);
  const dy = (event.clientY - state.pan.clientY) * (state.pan.viewBox.height / rect.height);
  state.viewBox = clampViewBox({
    ...state.pan.viewBox,
    x: state.pan.viewBox.x - dx,
    y: state.pan.viewBox.y - dy,
  });
  els.preview.setAttribute("viewBox", `${state.viewBox.x} ${state.viewBox.y} ${state.viewBox.width} ${state.viewBox.height}`);
}

function endPreviewPan(event) {
  if (state.pan?.pointerId && state.pan.pointerId !== event.pointerId) return;
  state.pan = null;
  els.preview.classList.remove("is-panning");
}

function handlePreviewKeyPan(event) {
  if (!state.viewBox) return;
  const stepX = state.viewBox.width * 0.08;
  const stepY = state.viewBox.height * 0.08;
  const next = { ...state.viewBox };
  if (event.key === "ArrowLeft") next.x -= stepX;
  else if (event.key === "ArrowRight") next.x += stepX;
  else if (event.key === "ArrowUp") next.y -= stepY;
  else if (event.key === "ArrowDown") next.y += stepY;
  else return;
  event.preventDefault();
  state.viewBox = clampViewBox(next);
  renderPreview();
}

function clampViewBox(viewBox) {
  if (!state.baseViewBox) return viewBox;
  const minX = state.baseViewBox.x;
  const maxX = state.baseViewBox.x + state.baseViewBox.width - viewBox.width;
  const minY = state.baseViewBox.y;
  const maxY = state.baseViewBox.y + state.baseViewBox.height - viewBox.height;
  return {
    ...viewBox,
    x: maxX >= minX ? Math.min(maxX, Math.max(minX, viewBox.x)) : viewBox.x,
    y: maxY >= minY ? Math.min(maxY, Math.max(minY, viewBox.y)) : viewBox.y,
  };
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
