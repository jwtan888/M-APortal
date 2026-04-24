(function () {
  const seedData = window.DFM_SEED_DATA || { records: [], warnings: [], defectCatalog: [] };
  const STORAGE_KEY = "dfm-dashboard-records";
  const SYNC_META_KEY = "dfm-dashboard-sync-meta";
  const INVESTMENT_NOTES_KEY = "dfm-dashboard-investment-notes";
  const INVESTMENT_VISIBILITY_KEY = "dfm-dashboard-investment-visibility";
  const LEGACY_STORAGE_PREFIX = "dfm-dashboard-records-";
  const PENDING_SYNC_TTL_MS = 10 * 60 * 1000;
  const FLOW_ENDPOINTS = {
    add: "https://defaultb4f081a089004baaa6a8ff79312af2.61.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/d714eb96fab644dcbfd6c83f28d817b1/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=3MUeOaSxcE-uRBtPqIe2-_KlK8BZ96hU1e2tivu0HOQ",
    update:
      "https://defaultb4f081a089004baaa6a8ff79312af2.61.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/bf99bd66b5064455a0fef384e50953e2/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=1TzaVpi7HoqovWJL1iKsVlkk2QoX0x6pYuwNWUtMKwA",
    delete:
      "https://defaultb4f081a089004baaa6a8ff79312af2.61.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/52a76b61c7e84e09a8df6eee0ad2a029/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=QslAg0t1-QLN3UvBy_7jBzUaY7YxcfADrJIjm15i-rw",
  };
  const FETCH_DFM_CHART_URL =
    "https://defaultb4f081a089004baaa6a8ff79312af2.61.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/777a0c72d4684db68a350132def5fb37/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=4SIb_3sXxhX4jOUhTA6L8-kcgrr37CcmJwDkloVPDv4";
  const FETCH_DFM_SUMMARY_URL =
    "https://defaultb4f081a089004baaa6a8ff79312af2.61.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/ebef663b8a79414cb89579a4b4edf9d6/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=JSZZaLOTkfxWcoDSdb3J7w1rhRQ400g2exl7yrBdKbg";
  const DFM_SUMMARY_UPDATE_URL =
    "https://defaultb4f081a089004baaa6a8ff79312af2.61.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/8a1ee6b32b44477ab947ff85d3c38881/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=-cqYV3Ac1tBeoECt_zje-s9Yk1B0RZYvDxqV-fP24N8";
  const defectMap = new Map((seedData.defectCatalog || []).map((item) => [item.code, item]));
  const page = document.body.dataset.page || "dashboard";
  const AUTO_REFRESH_MS = page === "dashboard" ? 60000 : 0;

  const elements = {
    seasonDonut: document.getElementById("season-donut"),
    seasonDonutTotal: document.getElementById("season-donut-total"),
    seasonDonutLegend: document.getElementById("season-donut-legend"),
    categoryBars: document.getElementById("category-bars"),
    seasonTiles: document.getElementById("season-tiles"),
    kpiGrid: document.getElementById("kpi-grid"),
    benchmarkChart: document.getElementById("benchmark-chart"),
    investmentBoard: document.getElementById("investment-board"),
    investmentControls: document.getElementById("investment-controls"),
    filteredSummary: document.getElementById("filtered-summary"),
    seasonBars: document.getElementById("season-bars"),
    codeBars: document.getElementById("code-bars"),
    modSummary: document.getElementById("mod-summary"),
    styleCards: document.getElementById("style-cards"),
    codeDirectory: document.getElementById("code-directory"),
    dashboardDetailGrid: document.getElementById("dashboard-detail-grid"),
    syncSource: document.getElementById("sync-source"),
    syncStatus: document.getElementById("sync-status"),
    syncTime: document.getElementById("sync-time"),
    recordCards: document.getElementById("record-cards"),
    searchInput: document.getElementById("search-input"),
    seasonFilter: document.getElementById("season-filter"),
    categoryFilter: document.getElementById("category-filter"),
    modificationFilter: document.getElementById("modification-filter"),
    defectFilter: document.getElementById("defect-filter"),
    addRecordBtn: document.getElementById("add-record-btn"),
    refreshDataBtn: document.getElementById("refresh-data-btn"),
    resetDataBtn: document.getElementById("reset-data-btn"),
    dialog: document.getElementById("record-dialog"),
    dialogTitle: document.getElementById("dialog-title"),
    closeDialogBtn: document.getElementById("close-dialog-btn"),
    cancelDialogBtn: document.getElementById("cancel-dialog-btn"),
    form: document.getElementById("record-form"),
  };

  const state = {
    records: loadRecords(),
    syncMeta: loadSyncMeta(),
    investmentNotes: loadInvestmentNotes(),
    investmentVisibility: loadInvestmentVisibility(),
    editingId: null,
    rotation: {
      seconds: 30,
      timerId: null,
      tick: 0,
    },
    seedRefreshTimerId: null,
    dashboardFrameId: null,
    dashboardDeferredTimeout: null,
    lastDashboardAnalytics: null,
    latestSource: "Seed file",
    lastRefreshAt: null,
    syncStatus: "Waiting for refresh",
    isSaving: false,
    investmentEditingCode: null,
    filters: {
      search: "",
      season: "all",
      category: "all",
      modification: "all",
      defectOnly: false,
    },
  };

  function loadRecords() {
    const raw = window.localStorage.getItem(STORAGE_KEY) || loadLegacyStoredRecords();
    if (!raw) {
      return normalizeRecords(seedData.records || []);
    }

    try {
      return normalizeRecords(reconcileStoredRecords(JSON.parse(raw), seedData.records || []));
    } catch (error) {
      console.error("Failed to parse saved data", error);
      return normalizeRecords(seedData.records || []);
    }
  }

  function persistRecords() {
    window.localStorage.setItem(STORAGE_KEY, JSON.stringify(state.records));
  }

  function loadSyncMeta() {
    const raw = window.localStorage.getItem(SYNC_META_KEY);
    if (!raw) {
      return { pendingUpserts: {}, pendingDeletes: {} };
    }
    try {
      const parsed = JSON.parse(raw);
      return {
        pendingUpserts: parsed.pendingUpserts || {},
        pendingDeletes: parsed.pendingDeletes || {},
      };
    } catch (error) {
      console.error("Failed to parse sync metadata", error);
      return { pendingUpserts: {}, pendingDeletes: {} };
    }
  }

  function loadInvestmentNotes() {
    const raw = window.localStorage.getItem(INVESTMENT_NOTES_KEY);
    if (!raw) {
      return {};
    }
    try {
      return JSON.parse(raw);
    } catch (error) {
      console.error("Failed to parse investment notes", error);
      return {};
    }
  }

  function persistInvestmentNotes() {
    window.localStorage.setItem(INVESTMENT_NOTES_KEY, JSON.stringify(state.investmentNotes));
  }

  function loadInvestmentVisibility() {
    const defaults = {
      samImprovement: true,
      improvementType: true,
      improvementValue: true,
      investmentDecision: true,
    };
    const raw = window.localStorage.getItem(INVESTMENT_VISIBILITY_KEY);
    if (!raw) {
      return defaults;
    }
    try {
      return {
        ...defaults,
        ...JSON.parse(raw),
      };
    } catch (error) {
      console.error("Failed to parse investment visibility", error);
      return defaults;
    }
  }

  function persistInvestmentVisibility() {
    window.localStorage.setItem(INVESTMENT_VISIBILITY_KEY, JSON.stringify(state.investmentVisibility));
  }

  function pruneSyncMeta(now = Date.now()) {
    ["pendingUpserts", "pendingDeletes"].forEach((bucket) => {
      Object.keys(state.syncMeta[bucket] || {}).forEach((id) => {
        if (now - Number(state.syncMeta[bucket][id] || 0) > PENDING_SYNC_TTL_MS) {
          delete state.syncMeta[bucket][id];
        }
      });
    });
  }

  function persistSyncMeta() {
    pruneSyncMeta();
    window.localStorage.setItem(SYNC_META_KEY, JSON.stringify(state.syncMeta));
  }

  function markPendingUpsert(recordId) {
    if (!recordId) {
      return;
    }
    state.syncMeta.pendingUpserts[recordId] = Date.now();
    delete state.syncMeta.pendingDeletes[recordId];
    persistSyncMeta();
  }

  function markPendingDelete(recordId) {
    if (!recordId) {
      return;
    }
    state.syncMeta.pendingDeletes[recordId] = Date.now();
    delete state.syncMeta.pendingUpserts[recordId];
    persistSyncMeta();
  }

  function loadLegacyStoredRecords() {
    const keys = Object.keys(window.localStorage)
      .filter((key) => key.startsWith(LEGACY_STORAGE_PREFIX))
      .sort()
      .reverse();

    for (const key of keys) {
      const value = window.localStorage.getItem(key);
      if (value) {
        return value;
      }
    }
    return null;
  }

  function setFormBusy(isBusy) {
    state.isSaving = isBusy;
    if (elements.addRecordBtn) {
      elements.addRecordBtn.disabled = isBusy;
    }
    if (elements.refreshDataBtn) {
      elements.refreshDataBtn.disabled = isBusy;
    }
    if (elements.resetDataBtn) {
      elements.resetDataBtn.disabled = isBusy;
    }
    if (elements.closeDialogBtn) {
      elements.closeDialogBtn.disabled = isBusy;
    }
    if (elements.cancelDialogBtn) {
      elements.cancelDialogBtn.disabled = isBusy;
    }
    if (elements.form) {
      Array.from(elements.form.elements).forEach((field) => {
        field.disabled = isBusy;
      });
    }
  }

  function normalizeRecords(records) {
    return records
      .map((record, index) => normalizeRecord(record, index))
      .filter((record) => !isBlankRecord(record));
  }

  function mergeRemoteWithPending(remoteRecords, localRecords) {
    pruneSyncMeta();
    const remoteNormalized = normalizeRecords(remoteRecords || []);
    const localNormalized = normalizeRecords(localRecords || []);
    const remoteById = new Map(remoteNormalized.map((record) => [record.id, record]));
    const localById = new Map(localNormalized.map((record) => [record.id, record]));

    Object.keys(state.syncMeta.pendingUpserts || {}).forEach((id) => {
      if (remoteById.has(id)) {
        delete state.syncMeta.pendingUpserts[id];
        return;
      }
      const localRecord = localById.get(id);
      if (localRecord) {
        remoteById.set(id, localRecord);
      }
    });

    Object.keys(state.syncMeta.pendingDeletes || {}).forEach((id) => {
      remoteById.delete(id);
    });

    persistSyncMeta();
    return Array.from(remoteById.values());
  }

  function nextRowId() {
    let maxId = 0;
    state.records.forEach((record) => {
      const match = cleanText(record.rowId || record.id).match(/^dfm-(\d+)$/i);
      if (!match) {
        return;
      }
      maxId = Math.max(maxId, Number(match[1]));
    });
    return `dfm-${maxId + 1}`;
  }

  function reconcileStoredRecords(storedRecords, latestSeedRecords) {
    const seedBySourceRow = new Map();
    const seedByShape = new Map();

    latestSeedRecords.forEach((record) => {
      if (record.sourceRow !== null && record.sourceRow !== undefined) {
        seedBySourceRow.set(String(record.sourceRow), record);
      }
      seedByShape.set(recordShapeKey(record), record);
    });

    return storedRecords.map((record) => {
      const sourceRowKey =
        record && record.sourceRow !== null && record.sourceRow !== undefined
          ? String(record.sourceRow)
          : "";
      const seedMatch =
        seedBySourceRow.get(sourceRowKey) ||
        seedByShape.get(recordShapeKey(record)) ||
        null;

      if (!seedMatch) {
        return record;
      }

      const currentId = cleanText(record.id || record.no || record["No."]);
      const shouldRefreshKey = !currentId || /^row-\d+$/i.test(currentId);

      if (!shouldRefreshKey) {
        return {
          ...record,
          typeCode: seedMatch.typeCode || record.typeCode,
          modification: seedMatch.modification || record.modification,
          feature: seedMatch.feature || record.feature,
          description: seedMatch.description || record.description,
          productClass: seedMatch.productClass || record.productClass,
          defects: seedMatch.defects || record.defects,
          defectCount:
            typeof seedMatch.defectCount === "number" ? seedMatch.defectCount : record.defectCount,
          totalIntensity:
            typeof seedMatch.totalIntensity === "number" ? seedMatch.totalIntensity : record.totalIntensity,
          fgQty:
            record.fgQty === null || record.fgQty === undefined
              ? (seedMatch.fgQty ?? record.fgQty)
              : record.fgQty,
          sourceRow: record.sourceRow ?? seedMatch.sourceRow,
        };
      }

      return {
        ...record,
        id: seedMatch.id || record.id,
        no: seedMatch.no || seedMatch["No."] || record.no,
        typeCode: seedMatch.typeCode || record.typeCode,
        modification: seedMatch.modification || record.modification,
        feature: seedMatch.feature || record.feature,
        description: seedMatch.description || record.description,
        productClass: seedMatch.productClass || record.productClass,
        defects: seedMatch.defects || record.defects,
        defectCount:
          typeof seedMatch.defectCount === "number" ? seedMatch.defectCount : record.defectCount,
        totalIntensity:
          typeof seedMatch.totalIntensity === "number" ? seedMatch.totalIntensity : record.totalIntensity,
        fgQty:
          record.fgQty === null || record.fgQty === undefined
            ? (seedMatch.fgQty ?? record.fgQty)
            : record.fgQty,
        sourceRow: seedMatch.sourceRow ?? record.sourceRow,
      };
    });
  }

  function normalizeRecord(record, index) {
    const season = cleanText(record.season);
    const style = cleanText(record.style);
    const typeCode = cleanText(record.typeCode || record.constructionCode);
    const normalizedNo = normalizeExcelNo(record.no || record["No."], record.sourceRow);
    const normalizedRowId =
      cleanText(record.rowId || record.RowId || record.id) || `dfm-${Date.now() + index}`;
    const detail = defectMap.get(typeCode) || {};
    const defects = Array.isArray(record.defects)
      ? record.defects
      : Array.isArray(detail.defects)
        ? detail.defects
        : [];
    const totalIntensity =
      typeof record.totalIntensity === "number"
        ? record.totalIntensity
        : sum(defects.map((defect) => defect.intensity || 0));

    return {
      id: normalizedRowId,
      rowId: normalizedRowId,
      no: normalizedNo || "",
      sourceRow: record.sourceRow || null,
      season,
      category: cleanText(record.category),
      protoStage: cleanText(record.protoStage),
      style,
      styleKey: buildStyleKey(season, style),
      constructionCode: cleanText(record.constructionCode),
      typeCode,
      modification: cleanText(record.modification),
      remark: cleanText(record.remark),
      fgQty: numberOrNull(record.fgQty),
      fgAnchor: numberOrNull(record.fgAnchor),
      feature: cleanText(record.feature || detail.feature),
      description: cleanText(record.description || detail.description),
      productClass: cleanText(record.productClass || detail.productClass),
      defects,
      defectCount: defects.length,
      totalIntensity: round(totalIntensity, 2),
    };
  }

  function isBlankRecord(record) {
    return ![
      record.season,
      record.category,
      record.style,
      record.constructionCode,
      record.typeCode,
      record.modification,
      record.remark,
      record.fgQty,
    ].some((value) => String(value ?? "").trim() !== "");
  }

  function cleanText(value) {
    return String(value || "").replace(/\s+/g, " ").trim();
  }

  function normalizeExcelNo(value, sourceRow) {
    const text = cleanText(value);
    if (/^\d+$/.test(text)) {
      return text;
    }
    const rowMatch = text.match(/^row-(\d+)$/i);
    if (rowMatch) {
      return String(Math.max(1, Number(rowMatch[1]) - 1));
    }
    if (sourceRow !== null && sourceRow !== undefined && String(sourceRow).trim() !== "") {
      const sourceNumber = Number(sourceRow);
      if (Number.isFinite(sourceNumber)) {
        return String(Math.max(1, sourceNumber - 1));
      }
    }
    return text;
  }

  function buildStyleKey(season, style) {
    return cleanText(season) + "__" + cleanText(style);
  }

  function recordShapeKey(record) {
    return [
      cleanText(record?.season),
      cleanText(record?.style),
      cleanText(record?.constructionCode),
      cleanText(record?.typeCode),
      cleanText(record?.category),
    ].join("__");
  }

  function numberOrNull(value) {
    if (value === null || value === undefined || value === "") {
      return null;
    }
    const numeric = Number(value);
    return Number.isFinite(numeric) ? numeric : null;
  }

  function round(value, digits) {
    const factor = 10 ** digits;
    return Math.round(value * factor) / factor;
  }

  function sum(values) {
    return values.reduce((total, value) => total + Number(value || 0), 0);
  }

  function formatNumber(value) {
    return new Intl.NumberFormat("en-US").format(Number(value || 0));
  }

  function formatDecimal(value) {
    return Number(value || 0).toFixed(2);
  }

  function formatPercent(value) {
    return `${Math.round(Number(value || 0) * 100)}%`;
  }

  function formatCurrency(value) {
    return new Intl.NumberFormat("en-MY", {
      style: "currency",
      currency: "MYR",
      minimumFractionDigits: 2,
      maximumFractionDigits: 2,
    }).format(Number(value || 0));
  }

  function parseImprovementFactor(value) {
    const text = cleanText(value);
    if (!text) {
      return null;
    }
    const numericMatch = text.match(/-?\d+(?:\.\d+)?/);
    if (!numericMatch) {
      return null;
    }
    const numeric = Number(numericMatch[0]);
    if (!Number.isFinite(numeric)) {
      return null;
    }
    if (text.includes("%")) {
      return numeric / 100;
    }
    return numeric;
  }

  function calculateImprovementValue(samImprovement, totalVolume, factor, fallbackValue = "") {
    const sam = Number(samImprovement);
    const fallbackText = cleanText(fallbackValue);
    if (fallbackText === "0" || fallbackText === "0.00") {
      return "";
    }
    if (!Number.isFinite(sam) || sam === 0) {
      return fallbackText;
    }
    if (!Number.isFinite(factor)) {
      return fallbackText;
    }
    return formatCurrency(sam * Number(totalVolume || 0) * factor);
  }

  function extractImprovementValueFromSummaryResponse(payload) {
    if (!payload || typeof payload !== "object") {
      return "";
    }
    const candidates = [
      payload,
      payload.body,
      payload.data,
      payload.result,
    ].filter(Boolean);

    const findLooseMatch = (candidate, target) => {
      const targetKey = cleanText(target).toLowerCase();
      for (const [key, value] of Object.entries(candidate || {})) {
        if (cleanText(key).toLowerCase() === targetKey) {
          return value;
        }
      }
      return "";
    };

    for (const candidate of candidates) {
      const value =
        candidate?.improvementValue ??
        candidate?.["Improvement Value"] ??
        candidate?.ImprovementValue ??
        candidate?.["Improvement_x0020_Value"] ??
        findLooseMatch(candidate, "Improvement Value");
      const text = cleanText(value);
      if (text) {
        return text;
      }
    }

    return "";
  }

  function getInvestmentRowTotalVolume(code) {
    if (!code) {
      return 0;
    }
    const analytics = computeAnalytics(filterRecords(state.records));
    const row = analytics.investmentBoard.find((item) => item.label === code);
    return row ? Number(row.totalVolume || 0) : 0;
  }

  function codeMatchesSummaryLabel(recordCode, label) {
    const normalizedRecord = cleanText(recordCode).toUpperCase();
    const normalizedLabel = cleanText(label).toUpperCase();
    if (!normalizedRecord || !normalizedLabel) {
      return false;
    }
    if (normalizedRecord === normalizedLabel) {
      return true;
    }
    return normalizedRecord
      .split("/")
      .map((part) => cleanText(part).toUpperCase())
      .includes(normalizedLabel);
  }

  function formatInvestmentFieldValue(columnKey, value) {
    if (columnKey !== "improvementValue") {
      return cleanText(value) || "-";
    }
    const text = cleanText(value);
    if (!text) {
      return "-";
    }
    if (/^myr\b/i.test(text)) {
      return text.replace(/^myr\b/i, "MYR");
    }
    const numeric = Number(String(text).replace(/,/g, ""));
    if (Number.isFinite(numeric)) {
      return formatCurrency(numeric);
    }
    return text;
  }

  function seasonPriority(season) {
    const text = cleanText(season);
    const prefix = text.slice(0, 2).toUpperCase();
    const yearMatch = text.match(/(\d{2,4})/);
    const year = yearMatch ? Number(yearMatch[1]) : Number.MAX_SAFE_INTEGER;
    const order = { SP: 0, SU: 1, FA: 2, HO: 3 };
    const prefixRank = Object.prototype.hasOwnProperty.call(order, prefix) ? order[prefix] : 99;
    return [year, prefixRank, text];
  }

  function compareSeasons(a, b) {
    const left = seasonPriority(a);
    const right = seasonPriority(b);
    if (left[0] !== right[0]) {
      return left[0] - right[0];
    }
    if (left[1] !== right[1]) {
      return left[1] - right[1];
    }
    return left[2].localeCompare(right[2]);
  }

  function labelCount(value, singular, plural) {
    const count = Number(value || 0);
    return `${formatNumber(count)} ${count === 1 ? singular : plural}`;
  }

  function parseExcelLikeDate(value) {
    if (value === null || value === undefined || value === "") {
      return null;
    }
    if (value instanceof Date) {
      return Number.isNaN(value.getTime()) ? null : value;
    }
    if (typeof value === "number" && Number.isFinite(value) && value > 30000) {
      return new Date(Date.UTC(1899, 11, 30) + value * 24 * 60 * 60 * 1000);
    }
    const numeric = Number(value);
    if (typeof value === "string" && Number.isFinite(numeric) && numeric > 30000) {
      return new Date(Date.UTC(1899, 11, 30) + numeric * 24 * 60 * 60 * 1000);
    }
    const parsed = new Date(value);
    return Number.isNaN(parsed.getTime()) ? null : parsed;
  }

  function formatRefreshTime(value) {
    if (!value) {
      return "Not yet";
    }
    const parsed = parseExcelLikeDate(value);
    if (!parsed) {
      return "Not yet";
    }
    try {
      return new Intl.DateTimeFormat("en-MY", {
        hour: "2-digit",
        minute: "2-digit",
        second: "2-digit",
        hour12: true,
      }).format(parsed);
    } catch (error) {
      return "Not yet";
    }
  }

  function escapeHtml(value) {
    return String(value || "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#39;");
  }

  function filterRecords(records) {
    const search = state.filters.search.toLowerCase();
    return records.filter((record) => {
      if (state.filters.season !== "all" && record.season !== state.filters.season) {
        return false;
      }
      if (state.filters.category !== "all" && record.category !== state.filters.category) {
        return false;
      }
      if (state.filters.modification !== "all" && record.modification !== state.filters.modification) {
        return false;
      }
      if (state.filters.defectOnly && !record.defectCount) {
        return false;
      }
      if (!search) {
        return true;
      }
      const haystack = [
        record.season,
        record.category,
        record.style,
        record.constructionCode,
        record.typeCode,
        record.remark,
        record.feature,
        record.description,
      ]
        .join(" ")
        .toLowerCase();
      return haystack.includes(search);
    });
  }

  function computeWarnings(records) {
    const map = new Map();
    records.forEach((record) => {
      if (!record.styleKey || record.fgQty === null) {
        return;
      }
      if (!map.has(record.styleKey)) {
        map.set(record.styleKey, {
          season: record.season,
          style: record.style,
          values: [],
        });
      }
      const entry = map.get(record.styleKey);
      if (!entry.values.includes(record.fgQty)) {
        entry.values.push(record.fgQty);
      }
    });

    return Array.from(map.values())
      .filter((entry) => entry.values.length > 1)
      .map((entry) => ({
        message: `Conflicting FG Qty values for ${entry.season} / ${entry.style}: ${entry.values.join(", ")}`,
      }));
  }

  function buildStyleSummary(records) {
    const styles = new Map();

    records.forEach((record) => {
      if (!styles.has(record.styleKey)) {
        styles.set(record.styleKey, {
          season: record.season,
          style: record.style,
          category: record.category,
          fgQty: record.fgQty,
          constructionRows: 0,
          modifiedRows: 0,
          codes: new Set(),
        });
      }

      const style = styles.get(record.styleKey);
      style.constructionRows += 1;
      style.modifiedRows += record.modification === "M" ? 1 : 0;
      if ((style.fgQty === null || style.fgQty === undefined) && record.fgQty !== null && record.fgQty !== undefined) {
        style.fgQty = record.fgQty;
      }
      if (record.category && !style.category) {
        style.category = record.category;
      }
      if (record.constructionCode) {
        style.codes.add(record.constructionCode);
      }
    });

    return Array.from(styles.values()).map((style) => ({
      ...style,
      codeList: Array.from(style.codes).sort(),
    }));
  }

  function computeAnalytics(records) {
    const styleSummary = buildStyleSummary(records);
    const totalFg = sum(styleSummary.map((style) => style.fgQty || 0));
    const modificationRate = records.length
      ? records.filter((record) => record.modification === "M").length / records.length
      : 0;

    const seasonMap = new Map();
    const seasonStyleMap = new Map();
    const categoryMap = new Map();
    styleSummary.forEach((style) => {
      seasonMap.set(style.season, (seasonMap.get(style.season) || 0) + (style.fgQty || 0));
      seasonStyleMap.set(style.season, (seasonStyleMap.get(style.season) || 0) + 1);
      const categoryKey = style.category || "Unspecified";
      categoryMap.set(categoryKey, (categoryMap.get(categoryKey) || 0) + (style.fgQty || 0));
    });

    const codeMap = new Map();
    const seasonCodeMap = new Map();
    const codeSeasonMap = new Map();
    const utilizationTypeMap = new Map();
    const utilizationSeasonMap = new Map();
    records.forEach((record) => {
      const key = record.constructionCode || "Unknown";
      codeMap.set(key, (codeMap.get(key) || 0) + 1);
      const seasonKey = record.season || "Unknown";
      seasonCodeMap.set(seasonKey, (seasonCodeMap.get(seasonKey) || 0) + 1);
      if (!codeSeasonMap.has(key)) {
        codeSeasonMap.set(key, new Map());
      }
      const seasonCounts = codeSeasonMap.get(key);
      seasonCounts.set(seasonKey, (seasonCounts.get(seasonKey) || 0) + (record.fgQty || 0));

      const utilizationKey = record.typeCode || record.constructionCode || "Unknown";
      if (!utilizationTypeMap.has(utilizationKey)) {
        utilizationTypeMap.set(utilizationKey, 0);
      }
      utilizationTypeMap.set(utilizationKey, utilizationTypeMap.get(utilizationKey) + 1);
      if (!utilizationSeasonMap.has(utilizationKey)) {
        utilizationSeasonMap.set(utilizationKey, new Map());
      }
      const utilizationSeasons = utilizationSeasonMap.get(utilizationKey);
      utilizationSeasons.set(seasonKey, (utilizationSeasons.get(seasonKey) || 0) + (record.fgQty || 0));
    });

    const nonMRecords = records.filter((record) => record.modification === "Non-M");
    const nonMCodeMap = new Map();
    const nonMSeasonCodeMap = new Map();
    const nonMCodeSeasonMap = new Map();
    nonMRecords.forEach((record) => {
      const key = record.constructionCode || "Unknown";
      nonMCodeMap.set(key, (nonMCodeMap.get(key) || 0) + 1);
      const seasonKey = record.season || "Unknown";
      nonMSeasonCodeMap.set(seasonKey, (nonMSeasonCodeMap.get(seasonKey) || 0) + 1);
      if (!nonMCodeSeasonMap.has(key)) {
        nonMCodeSeasonMap.set(key, new Map());
      }
      const seasonCounts = nonMCodeSeasonMap.get(key);
      seasonCounts.set(seasonKey, (seasonCounts.get(seasonKey) || 0) + (record.fgQty || 0));
    });

    const investmentSeasonLabels = Array.from(nonMSeasonCodeMap.keys())
      .filter(Boolean)
      .sort(compareSeasons);

    const mTypeCodes = new Set();
    const nonMTypeCodes = new Set();
    const otherTypeCodes = new Set();
    records.forEach((record) => {
      const typeKey = cleanText(record.typeCode) || "Unknown";
      if (record.modification === "M") {
        mTypeCodes.add(typeKey);
      } else if (record.modification === "Non-M") {
        nonMTypeCodes.add(typeKey);
      } else {
        otherTypeCodes.add(typeKey);
      }
    });
    const totalModTypes = mTypeCodes.size + nonMTypeCodes.size + otherTypeCodes.size;

    const summaryDrivenRows = Object.entries(state.investmentNotes)
      .map(([label, manual]) => ({
        label,
        manual,
        rank: Number(manual.currentRank || 0),
        active: cleanText(manual.activeTop20).toLowerCase() === "yes",
      }))
      .filter((item) => item.active && item.rank > 0)
      .sort((a, b) => a.rank - b.rank || a.label.localeCompare(b.label))
      .slice(0, 20)
      .map(({ label, manual }) => {
        const seasonCounts = Array.from(nonMCodeSeasonMap.entries())
          .filter(([code]) => codeMatchesSummaryLabel(code, label))
          .reduce((acc, [, counts]) => {
            counts.forEach((value, season) => {
              acc.set(season, (acc.get(season) || 0) + value);
            });
            return acc;
          }, new Map());
        const seasonVolumes = investmentSeasonLabels.map((season) => ({
          season,
          value: seasonCounts.get(season) || 0,
        }));
        const fallbackTotalVolume = sum(seasonVolumes.map((item) => item.value));
        const totalVolume = Number(manual.currentTotalFgQty || 0) || fallbackTotalVolume;
        const factor = parseImprovementFactor(manual.improvementType);
        return {
          label,
          totalVolume,
          seasonVolumes,
          samImprovement: cleanText(manual.samImprovement),
          improvementType: cleanText(manual.improvementType),
          improvementValue: calculateImprovementValue(
            manual.samImprovement,
            totalVolume,
            factor,
            manual.improvementValue,
          ),
          investmentDecision: cleanText(manual.investmentDecision),
          updatedAt: manual.updatedAt || null,
        };
      });

    const investmentBoardRows = (summaryDrivenRows.length
      ? summaryDrivenRows
      : Array.from(nonMCodeMap.entries())
          .map(([label]) => {
            const seasonCounts = nonMCodeSeasonMap.get(label) || new Map();
            const manual = state.investmentNotes[label] || {};
            const seasonVolumes = investmentSeasonLabels.map((season) => ({
              season,
              value: seasonCounts.get(season) || 0,
            }));
            const totalVolume = sum(seasonVolumes.map((item) => item.value));
            const factor = parseImprovementFactor(manual.improvementType);
            return {
              label,
              totalVolume,
              seasonVolumes,
              samImprovement: cleanText(manual.samImprovement),
              improvementType: cleanText(manual.improvementType),
              improvementValue: calculateImprovementValue(
                manual.samImprovement,
                totalVolume,
                factor,
                manual.improvementValue,
              ),
              investmentDecision: cleanText(manual.investmentDecision),
              updatedAt: manual.updatedAt || null,
            };
          })
          .sort((a, b) => b.totalVolume - a.totalVolume || a.label.localeCompare(b.label))
          .slice(0, 20));

    return {
      styleSummary: styleSummary.sort((a, b) => b.fgQty - a.fgQty),
      kpis: [
        {
          label: "Total Style",
          value: formatNumber(styleSummary.length),
          subtext: "",
        },
        {
          label: "Total FG QTY",
          value: formatNumber(totalFg),
          subtext: "",
        },
        {
          label: "Total Unique Construction Code",
          value: formatNumber(codeMap.size),
          subtext: "",
        },
        {
          label: "construction occuracy",
          value: formatNumber(records.length),
          subtext: "",
        },
      ],
      seasonBars: Array.from(seasonMap.entries())
        .map(([label, value]) => ({ label, value }))
        .sort((a, b) => b.value - a.value),
      seasonDonut: Array.from(seasonCodeMap.entries())
        .map(([label, value]) => ({
          label,
          value,
          share: records.length ? value / records.length : 0,
        }))
        .sort((a, b) => b.value - a.value)
        .filter((item) => item.value > 0),
      categoryBars: Array.from(categoryMap.entries())
        .map(([label, value]) => ({ label, value }))
        .sort((a, b) => b.value - a.value)
        .slice(0, 5),
      seasonTiles: Array.from(seasonStyleMap.entries())
        .map(([label, value]) => ({
          label,
          value,
          fg: seasonMap.get(label) || 0,
        }))
        .sort((a, b) => compareSeasons(a.label, b.label))
        ,
      codeBars: investmentBoardRows.slice(0, 8).map((row) => ({
        label: row.label,
        value: row.totalVolume,
      })),
      modSummary: [
        {
          label: "Modified",
          value: mTypeCodes.size,
          share: totalModTypes ? mTypeCodes.size / totalModTypes : 0,
          note: "Unique type codes marked M",
        },
        {
          label: "Non-Modified",
          value: nonMTypeCodes.size,
          share: totalModTypes ? nonMTypeCodes.size / totalModTypes : 0,
          note: "Unique type codes marked Non-M",
        },
        {
          label: "Unspecified",
          value: otherTypeCodes.size,
          share: totalModTypes ? otherTypeCodes.size / totalModTypes : 0,
          note: "Type codes without modification value",
        },
      ].filter((item) => item.value > 0),
      codeDirectory: Array.from(codeMap.entries())
        .map(([label, value]) => ({ label, value }))
        .sort((a, b) => b.value - a.value || a.label.localeCompare(b.label))
        .slice(0, 40),
      benchmark: (() => {
        const activeNonMTypeCodes = new Set(
          records
            .filter((record) => record.modification === "Non-M")
            .map((record) => record.typeCode || record.constructionCode || "Unknown")
            .filter(Boolean),
        ).size;
        return {
          target: 641,
          active: activeNonMTypeCodes,
          inactive: Math.max(0, 641 - activeNonMTypeCodes),
          over: Math.max(0, activeNonMTypeCodes - 641),
        };
      })(),
      investmentSeasonLabels,
      investmentBoard: investmentBoardRows,
    };
  }

  function render() {
    populateFilterOptions();

    const filteredRecords = filterRecords(state.records);
    const computedAnalytics = computeAnalytics(filteredRecords);
    const hasDashboardContent =
      computedAnalytics.seasonDonut.length ||
      computedAnalytics.seasonBars.length ||
      computedAnalytics.codeBars.length ||
      computedAnalytics.investmentBoard.length ||
      computedAnalytics.categoryBars.length;
    const analytics =
      page === "dashboard" && !hasDashboardContent && state.lastDashboardAnalytics
        ? state.lastDashboardAnalytics
        : computedAnalytics;

    if (page === "dashboard") {
      if (hasDashboardContent) {
        state.lastDashboardAnalytics = computedAnalytics;
      }
      if (elements.dashboardDetailGrid) {
        elements.dashboardDetailGrid.hidden = !analytics.styleSummary.length && !analytics.codeDirectory.length;
      }
      renderKpis(analytics.kpis);
      renderBenchmarkChart(analytics.benchmark);
      renderInvestmentControls();
      renderInvestmentBoard(analytics.investmentBoard, analytics.investmentSeasonLabels, false);
      renderBars(elements.seasonBars, analytics.seasonBars, formatNumber);
      renderBars(elements.categoryBars, analytics.categoryBars, formatNumber);
      renderModSummary(analytics.modSummary);
      if (state.dashboardFrameId) {
        window.cancelAnimationFrame(state.dashboardFrameId);
      }
      if (state.dashboardDeferredTimeout) {
        window.clearTimeout(state.dashboardDeferredTimeout);
      }
      state.dashboardFrameId = window.requestAnimationFrame(() => {
        state.dashboardDeferredTimeout = window.setTimeout(() => {
          renderSeasonDonut(analytics.seasonDonut);
          renderSeasonTiles(analytics.seasonTiles);
          renderStyleCards(analytics.styleSummary);
          renderCodeDirectory(analytics.codeDirectory);
          state.dashboardDeferredTimeout = null;
        }, 80);
        state.dashboardFrameId = null;
      });
    }

    if (page === "data") {
      renderInvestmentControls();
      renderInvestmentBoard(analytics.investmentBoard, analytics.investmentSeasonLabels, true);
      renderRecordCards(filteredRecords);
      if (elements.filteredSummary) {
        elements.filteredSummary.textContent = `${labelCount(filteredRecords.length, "construction row", "construction rows")} shown`;
      }
      if (elements.syncSource) {
        elements.syncSource.textContent = `Data source: ${state.latestSource}`;
      }
      if (elements.syncStatus) {
        elements.syncStatus.textContent = `Sync status: ${state.syncStatus}`;
      }
      if (elements.syncTime) {
        elements.syncTime.textContent = `Last refreshed: ${formatRefreshTime(state.lastRefreshAt)}`;
      }
    }
  }

  function populateFilterOptions() {
    if (!elements.seasonFilter || !elements.categoryFilter || !elements.modificationFilter) {
      return;
    }
    state.filters.season = setSelectOptions(
      elements.seasonFilter,
      ["all"].concat(uniqueValues(state.records, "season")),
      state.filters.season,
      "All seasons",
    );
    state.filters.category = setSelectOptions(
      elements.categoryFilter,
      ["all"].concat(uniqueValues(state.records, "category")),
      state.filters.category,
      "All categories",
    );
    state.filters.modification = setSelectOptions(
      elements.modificationFilter,
      ["all"].concat(uniqueValues(state.records, "modification")),
      state.filters.modification,
      "All modifications",
    );
  }

  function uniqueValues(records, key) {
    return Array.from(new Set(records.map((record) => record[key]).filter(Boolean))).sort();
  }

  function setSelectOptions(select, values, activeValue, allLabel) {
    const current = values.includes(activeValue) ? activeValue : "all";
    select.innerHTML = values
      .map((value) => {
        const label = value === "all" ? allLabel : value;
        const selected = value === current ? " selected" : "";
        return `<option value="${escapeHtml(value)}"${selected}>${escapeHtml(label)}</option>`;
      })
      .join("");
    return current;
  }

  function renderKpis(kpis) {
    if (!elements.kpiGrid) {
      return;
    }
    elements.kpiGrid.innerHTML = kpis
      .map(
        (kpi) => `
          <article class="kpi-card">
            <p class="kpi-label">${escapeHtml(kpi.label)}</p>
            <p class="kpi-value">${escapeHtml(kpi.value)}</p>
            <p class="kpi-subtext">${escapeHtml(kpi.subtext)}</p>
          </article>
        `,
      )
      .join("");
  }

  function renderBenchmarkChart(benchmark) {
    if (!elements.benchmarkChart) {
      return;
    }
    const target = Number(benchmark?.target || 0);
    const active = Number(benchmark?.active || 0);
    const inactive = Number(benchmark?.inactive || 0);
    const over = Number(benchmark?.over || 0);
    const utilization = target > 0 ? Math.round((active / target) * 100) : 0;
    const denominator = Math.max(target, active, 1);
    const activeWidth = Math.max(0, (Math.min(active, denominator) / denominator) * 100);
    const inactiveWidth = Math.max(0, (inactive / denominator) * 100);

    elements.benchmarkChart.innerHTML = `
      <div class="benchmark-summary">
        <div class="benchmark-current">
          <div class="benchmark-label">Current Active Construction Code (Non-M)</div>
          <div class="benchmark-value">${escapeHtml(formatNumber(active))}<span class="benchmark-value-inline">${escapeHtml(`${utilization}%`)}</span></div>
          <div class="benchmark-meta">${over > 0 ? `${escapeHtml(formatNumber(over))} above benchmark` : ""}</div>
        </div>
        <div class="benchmark-current">
          <div class="benchmark-label">Nike Construction Code</div>
          <div class="benchmark-value">${escapeHtml(formatNumber(target))}</div>
          <div class="benchmark-meta"></div>
        </div>
      </div>
      <div class="benchmark-track">
        <div class="benchmark-fill" style="width:${activeWidth}%"></div>
        ${inactive > 0 ? `<div class="benchmark-fill benchmark-fill--inactive" style="width:${inactiveWidth}%"></div>` : ""}
      </div>
      <div class="benchmark-legend">
        <div class="benchmark-legend-item">
          <span class="benchmark-legend-label"><span class="benchmark-dot"></span>Active code qty</span>
          <span class="benchmark-legend-value">${escapeHtml(formatNumber(active))}</span>
        </div>
        <div class="benchmark-legend-item">
          <span class="benchmark-legend-label"><span class="benchmark-dot benchmark-dot--inactive"></span>Inactive code qty</span>
          <span class="benchmark-legend-value">${escapeHtml(formatNumber(inactive))}</span>
        </div>
      </div>
    `;
  }

  function getInvestmentManualColumns(editable = false) {
    const columns = [
      {
        key: "samImprovement",
        label: "SAM Improvement",
        role: "sam",
        placeholder: "SAM improvement",
      },
      {
        key: "improvementType",
        label: "Improvement Type",
        role: "improvement-type",
        placeholder: "Improvement type",
      },
      {
        key: "improvementValue",
        label: "Improvement Value",
        role: "improvement-value",
        displayOnly: true,
      },
      {
        key: "investmentDecision",
        label: "Investment Decision",
        role: "decision",
        type: "select",
      },
    ];

    if (editable) {
      return columns;
    }

    return columns.filter((column) => state.investmentVisibility[column.key] !== false);
  }

  function renderInvestmentControls() {
    if (!elements.investmentControls) {
      return;
    }

    const labels = {
      samImprovement: "SAM Improvement",
      improvementType: "Improvement Type",
      improvementValue: "Improvement Value",
      investmentDecision: "Investment Decision",
    };

    elements.investmentControls.innerHTML = `
      <span class="panel-note">Dashboard columns</span>
      ${Object.entries(labels)
        .map(
          ([key, label]) => `
            <button
              class="investment-toggle ${state.investmentVisibility[key] !== false ? "is-active" : ""}"
              type="button"
              data-action="toggle-investment-column"
              data-column="${escapeHtml(key)}"
            >
              ${escapeHtml(label)}
            </button>
          `,
        )
        .join("")}
    `;
  }

  function renderInvestmentBoard(rows, seasonLabels, editable = false) {
    if (!elements.investmentBoard) {
      return;
    }
    if (!rows.length) {
      elements.investmentBoard.innerHTML = '<div class="empty-state">No construction code volume available.</div>';
      return;
    }

    const manualColumns = getInvestmentManualColumns(editable);

    elements.investmentBoard.innerHTML = `
      <table class="investment-table ${editable ? "investment-table--editable" : "investment-table--readonly"}">
        <thead>
          <tr>
            <th rowspan="2">Construction Code</th>
            <th colspan="${Math.max(seasonLabels.length, 1)}" class="group-head">Volume by Season</th>
            <th rowspan="2">Total FG Volume</th>
            ${manualColumns.map((column) => `<th rowspan="2">${escapeHtml(column.label)}</th>`).join("")}
            ${editable ? '<th rowspan="2"></th>' : ""}
          </tr>
          <tr>
            ${
              seasonLabels.length
                ? seasonLabels.map((season) => `<th>${escapeHtml(season)}</th>`).join("")
                : "<th>Season</th>"
            }
          </tr>
        </thead>
        <tbody>
          ${rows
            .map(
              (row, index) => `
                <tr data-code="${escapeHtml(row.label)}" data-editing="${state.investmentEditingCode === row.label ? "true" : "false"}">
                  <td class="investment-code"><span class="investment-rank ${
                    index === 0
                      ? "investment-rank--1"
                      : index === 1
                        ? "investment-rank--2"
                        : index === 2
                          ? "investment-rank--3"
                          : ""
                  }">Top ${index + 1}</span>${escapeHtml(row.label)}</td>
                  ${
                    seasonLabels.length
                      ? row.seasonVolumes
                          .map(
                            (item) =>
                              `<td class="investment-volume">${item.value ? escapeHtml(formatNumber(item.value)) : "-"}</td>`,
                          )
                          .join("")
                      : '<td class="investment-volume">-</td>'
                  }
                  <td class="investment-total">${escapeHtml(formatNumber(row.totalVolume))}</td>
                  ${manualColumns
                    .map((column, columnIndex) => {
                      if (editable && column.type === "select") {
                        return `
                          <td class="investment-manual-cell">
                            <select class="investment-select" data-role="${escapeHtml(column.role)}" ${state.investmentEditingCode === row.label ? "" : "disabled"}>
                              ${["", "Yes", "No", "Review"]
                                .map((option) => {
                                  const selected = option === row[column.key] ? " selected" : "";
                                  const label = option || "Select decision";
                                  return `<option value="${escapeHtml(option)}"${selected}>${escapeHtml(label)}</option>`;
                                })
                                .join("")}
                            </select>
                            ${
                              editable && columnIndex === manualColumns.length - 1
                                ? `<span class="investment-meta">${row.updatedAt ? `Updated ${formatRefreshTime(row.updatedAt)}` : ""}</span>`
                                : ""
                            }
                          </td>
                        `;
                      }

                      return `
                        <td class="investment-manual-cell">
                          ${
                            editable && !column.displayOnly
                              ? `<input class="investment-input" data-role="${escapeHtml(column.role)}" type="text" ${state.investmentEditingCode === row.label ? "" : "disabled"} value="${escapeHtml(
                                  row[column.key] || "",
                                )}" placeholder="${escapeHtml(column.placeholder || "")}" />`
                              : `<span class="investment-display">${escapeHtml(formatInvestmentFieldValue(column.key, row[column.key]))}</span>`
                          }
                          ${
                            editable && columnIndex === manualColumns.length - 1
                              ? `<span class="investment-meta">${row.updatedAt ? `Updated ${formatRefreshTime(row.updatedAt)}` : ""}</span>`
                              : ""
                          }
                        </td>
                      `;
                    })
                    .join("")}
                  ${
                    editable
                      ? `<td class="investment-actions">
                    <button class="investment-save" type="button" data-action="${state.investmentEditingCode === row.label ? "save-investment" : "edit-investment"}" data-code="${escapeHtml(
                      row.label,
                    )}">${state.investmentEditingCode === row.label ? "Save" : "Edit"}</button>
                  </td>`
                      : ""
                  }
                </tr>
              `,
            )
            .join("")}
        </tbody>
      </table>
    `;
  }

  function renderSeasonDonut(items) {
    if (!elements.seasonDonut || !elements.seasonDonutLegend || !elements.seasonDonutTotal) {
      return;
    }
    const panel = elements.seasonDonut.closest(".panel");
    if (!items.length) {
      panel?.classList.add("panel--hidden");
      elements.seasonDonut.style.background = "rgba(30, 87, 216, 0.08)";
      elements.seasonDonutLegend.innerHTML = '<div class="empty-state">No season data available.</div>';
      elements.seasonDonutTotal.textContent = "0";
      return;
    }
    panel?.classList.remove("panel--hidden");

    const palette = ["#1e57d8", "#4f7de5", "#7ca4ef", "#aac4f5", "#c7f13b", "#151515", "#8d98b3"];
    const activeIndex = items.length ? state.rotation.tick % items.length : 0;
    let current = 0;
    const stops = items.map((item, index) => {
      const color = palette[index % palette.length];
      const start = current;
      current += item.share * 100;
      return {
        ...item,
        color,
        start,
        end: current,
      };
    });

    elements.seasonDonut.style.background = `conic-gradient(${stops
      .map((item) => `${item.color} ${item.start}% ${item.end}%`)
      .join(", ")})`;
    elements.seasonDonutTotal.textContent = formatNumber(sum(items.map((item) => item.value)));
    elements.seasonDonutLegend.innerHTML = stops
      .map(
        (item, index) => `
          <div class="legend-row">
            <span class="legend-dot" style="background:${item.color}"></span>
            <span class="legend-label ${index === activeIndex ? "is-active" : ""}">${escapeHtml(item.label || "Unknown")}</span>
            <span class="legend-value ${index === activeIndex ? "is-active" : ""}">${escapeHtml(`${formatPercent(item.share)} · ${formatNumber(item.value)}`)}</span>
          </div>
        `,
      )
      .join("");

    const activeSlice = stops[activeIndex];
    if (activeSlice) {
      elements.seasonDonut.style.boxShadow = `0 0 0 8px ${activeSlice.color}15`;
    }
    elements.seasonDonut.style.transform = `rotate(${state.rotation.tick * 18}deg)`;
  }

  function renderBars(target, items, formatter) {
    if (!target) {
      return;
    }
    const panel = target.closest(".panel");
    if (!items.length) {
      panel?.classList.add("panel--hidden");
      target.innerHTML = '<div class="empty-state">No data for the current filter set.</div>';
      return;
    }
    panel?.classList.remove("panel--hidden");
    const max = Math.max(...items.map((item) => item.value), 1);
    const activeIndex = items.length ? state.rotation.tick % items.length : 0;
    target.innerHTML = items
      .map(
        (item, index) => `
          <div class="bar-row">
            <div class="bar-head">
              <span class="${index === activeIndex ? "is-active" : ""}">${escapeHtml(item.label)}</span>
              <span class="${index === activeIndex ? "is-active" : ""}">${escapeHtml(formatter(item.value))}</span>
            </div>
            <div class="bar-track">
              <div class="bar-fill ${index === activeIndex ? "is-active" : ""}" style="width:${Math.max(8, (item.value / max) * 100)}%"></div>
            </div>
          </div>
        `,
      )
      .join("");
  }

  function renderModSummary(items) {
    if (!elements.modSummary) {
      return;
    }
    const panel = elements.modSummary.closest(".panel");
    if (!items.length) {
      panel?.classList.add("panel--hidden");
      elements.modSummary.innerHTML = '<div class="empty-state">No modification summary available.</div>';
      return;
    }
    panel?.classList.remove("panel--hidden");

    const activeIndex = items.length ? state.rotation.tick % items.length : 0;
    elements.modSummary.innerHTML = items
      .map(
        (item, index) => `
          <div class="stack-item ${index === activeIndex ? "is-active" : ""}">
            <div>${escapeHtml(item.label)}</div>
            <div class="stack-value">${escapeHtml(formatNumber(item.value))}${item.share !== undefined ? ` <span class="stack-value-inline">${escapeHtml(formatPercent(item.share))}</span>` : ""}</div>
            <div class="stack-meta">${escapeHtml(item.note)}</div>
          </div>
        `,
      )
      .join("");
  }

  function renderSeasonTiles(items) {
    if (!elements.seasonTiles) {
      return;
    }
    const panel = elements.seasonTiles.closest(".panel");
    if (!items.length) {
      panel?.classList.add("panel--hidden");
      elements.seasonTiles.innerHTML = '<div class="empty-state">No season density available.</div>';
      return;
    }
    panel?.classList.remove("panel--hidden");

    elements.seasonTiles.innerHTML = items
      .map(
        (item, index) => `
          <article class="metric-tile ${index === (items.length ? state.rotation.tick % items.length : 0) ? "is-active" : ""}">
            <div class="metric-tile-label">${escapeHtml(item.label || "Unknown")}</div>
            <div class="metric-tile-value">${escapeHtml(`${formatNumber(item.value)} style`)}</div>
          </article>
        `,
      )
      .join("");
  }

  function updateDashboardDetailVisibility() {
    if (!elements.dashboardDetailGrid) {
      return;
    }
    const panels = Array.from(elements.dashboardDetailGrid.querySelectorAll(".panel"));
    const hasVisiblePanel = panels.some((panel) => !panel.classList.contains("panel--hidden"));
    elements.dashboardDetailGrid.classList.toggle("panel-grid--collapsed", !hasVisiblePanel);
  }

  function renderStyleCards(styles) {
    if (!elements.styleCards) {
      return;
    }
    const panel = elements.styleCards.closest(".panel");
    if (!styles.length) {
      panel?.classList.add("panel--hidden");
      elements.styleCards.innerHTML = '<div class="empty-state">No styles match the current filters.</div>';
      updateDashboardDetailVisibility();
      return;
    }
    panel?.classList.remove("panel--hidden");
    updateDashboardDetailVisibility();

    elements.styleCards.innerHTML = styles
      .slice(0, 12)
      .map(
        (style, index) => `
          <article class="style-card ${index === (styles.length ? state.rotation.tick % styles.length : 0) ? "is-active" : ""}">
            <div class="style-head">
              <div>
                <div class="style-name">${escapeHtml(style.style)}</div>
                <div class="style-meta">${escapeHtml(style.season)} · ${escapeHtml(style.category || "No category")}</div>
              </div>
              <span class="tag">${escapeHtml(style.fgQty === null || style.fgQty === undefined ? "Pending FG" : `${formatNumber(style.fgQty)} FG`)}</span>
            </div>
            <div class="style-stats">
              <div class="mini-stat">
                <div class="mini-label">Codes</div>
                <div class="mini-value">${escapeHtml(labelCount(style.codeList.length, "code", "codes"))}</div>
              </div>
              <div class="mini-stat">
                <div class="mini-label">Modified</div>
                <div class="mini-value">${escapeHtml(labelCount(style.modifiedRows, "occurrence", "occurrences"))}</div>
              </div>
            </div>
            <div class="style-meta">${escapeHtml(style.codeList.slice(0, 5).join(", ") || "No construction code")}</div>
          </article>
        `,
      )
      .join("");
  }

  function renderCodeDirectory(items) {
    if (!elements.codeDirectory) {
      return;
    }
    const panel = elements.codeDirectory.closest(".panel");
    if (!items.length) {
      panel?.classList.add("panel--hidden");
      elements.codeDirectory.innerHTML = '<div class="empty-state">No construction codes in the current filter.</div>';
      updateDashboardDetailVisibility();
      return;
    }
    panel?.classList.remove("panel--hidden");
    updateDashboardDetailVisibility();

    elements.codeDirectory.innerHTML = items
      .map(
        (item, index) => `
          <article class="code-chip ${index === (items.length ? state.rotation.tick % items.length : 0) ? "is-active" : ""}">
            <div class="code-label">${escapeHtml(item.label)}</div>
            <div class="code-meta">${escapeHtml(labelCount(item.value, "occurrence", "occurrences"))}</div>
          </article>
        `,
      )
      .join("");
  }

  function renderRecordCards(records) {
    if (!elements.recordCards) {
      return;
    }
    if (!records.length) {
      elements.recordCards.innerHTML = '<div class="empty-state">No construction rows match the current filters.</div>';
      return;
    }

    elements.recordCards.innerHTML = records
      .map(
        (record) => `
          <article class="record-card">
            <div class="record-head">
              <div>
                <div class="record-style">${escapeHtml(record.style)}</div>
                <div class="record-meta">${escapeHtml(record.season)} · ${escapeHtml(record.category || "No category")}</div>
              </div>
              ${renderModificationTag(record.modification)}
            </div>
            <div class="record-meta">Construction Code: <strong>${escapeHtml(record.constructionCode || "-")}</strong></div>
            <div class="record-meta">Type Code: <strong>${escapeHtml(record.typeCode || "-")}</strong></div>
            <div class="record-meta">FG Qty: <strong>${escapeHtml(record.fgQty === null || record.fgQty === undefined ? "Pending update" : formatNumber(record.fgQty))}</strong></div>
            <div class="record-meta">${escapeHtml(record.remark || "No remark")}</div>
            <div class="row-actions">
              <button class="row-btn" data-action="edit" data-id="${escapeHtml(record.id)}">Edit</button>
              <button class="row-btn row-btn--danger" data-action="delete" data-id="${escapeHtml(record.id)}">Delete</button>
            </div>
          </article>
        `,
      )
      .join("");
  }

  function persistInvestmentNote(
    code,
    samImprovement,
    improvementType,
    improvementValue,
    investmentDecision,
    updatedAt,
  ) {
    if (!code) {
      return;
    }
    const nextSam = cleanText(samImprovement);
    const nextImprovementType = cleanText(improvementType);
    const nextImprovementValue = cleanText(improvementValue);
    const nextDecision = cleanText(investmentDecision);
    if (!nextSam && !nextImprovementType && !nextImprovementValue && !nextDecision) {
      delete state.investmentNotes[code];
    } else {
      const existing = state.investmentNotes[code] || {};
      state.investmentNotes[code] = {
        ...existing,
        samImprovement: nextSam,
        improvementType: nextImprovementType,
        improvementValue: nextImprovementValue,
        investmentDecision: nextDecision,
        updatedAt,
      };
    }
    persistInvestmentNotes();
    render();
  }

  function formatExcelFriendlyTimestamp(value) {
    const date = value instanceof Date ? value : new Date(value);
    if (Number.isNaN(date.getTime())) {
      return "";
    }
    const local = new Date(
      date.toLocaleString("en-US", { timeZone: "Asia/Kuala_Lumpur" }),
    );
    const year = local.getFullYear();
    const month = String(local.getMonth() + 1).padStart(2, "0");
    const day = String(local.getDate()).padStart(2, "0");
    const hour = String(local.getHours()).padStart(2, "0");
    const minute = String(local.getMinutes()).padStart(2, "0");
    const second = String(local.getSeconds()).padStart(2, "0");
    return `${year}-${month}-${day} ${hour}:${minute}:${second}`;
  }

  async function saveInvestmentRow(
    code,
    samImprovement,
    improvementType,
    investmentDecision,
  ) {
    if (!code) {
      return;
    }
    const nextSam = cleanText(samImprovement);
    const nextImprovementType = cleanText(improvementType);
    const nextDecision = cleanText(investmentDecision);
    const updatedAt = Date.now();
    const payload = {
      constructionCode: cleanText(code),
      samImprovement: nextSam,
      improvementType: nextImprovementType,
      investmentDecision: nextDecision,
      updatedBy: "Gavin",
      updatedAt: formatExcelFriendlyTimestamp(updatedAt),
    };

    try {
      const response = await window.fetch(DFM_SUMMARY_UPDATE_URL, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify(payload),
      });

      if (!response.ok) {
        const text = await response.text().catch(() => "");
        throw new Error(text || `DFM Summary update failed with ${response.status}.`);
      }

      let improvementValueFromResponse = "";
      try {
        const responseText = await response.text();
        if (responseText) {
          const parsed = JSON.parse(responseText);
          improvementValueFromResponse = extractImprovementValueFromSummaryResponse(parsed);
        }
      } catch (error) {
        improvementValueFromResponse = "";
      }

      persistInvestmentNote(
        code,
        nextSam,
        nextImprovementType,
        improvementValueFromResponse,
        nextDecision,
        updatedAt,
      );
    } catch (error) {
      if (isNoResponseError(error)) {
        persistInvestmentNote(
          code,
          nextSam,
          nextImprovementType,
          calculateImprovementValue(
            nextSam,
            getInvestmentRowTotalVolume(code),
            parseImprovementFactor(nextImprovementType),
            "",
          ),
          nextDecision,
          updatedAt,
        );
        return;
      }
      console.error(error);
      window.alert(`Unable to save DFM Summary decision.\n${error.message}`);
    }
  }

  function renderModificationTag(value) {
    const normalized = cleanText(value);
    if (normalized === "M") {
      return '<span class="tag">M</span>';
    }
    if (normalized === "Non-M") {
      return '<span class="tag tag--neutral">Non-M</span>';
    }
    return '<span class="tag tag--success">N/A</span>';
  }

  function openDialog(recordId) {
    if (!elements.dialog || !elements.form || !elements.dialogTitle) {
      return;
    }
    state.editingId = recordId || null;
    const record = state.records.find((item) => item.id === recordId);
    elements.dialogTitle.textContent = record ? "Edit Record" : "Add Record";
    elements.form.reset();

    const values = record || {
      season: "",
      category: "",
      protoStage: "",
      style: "",
      constructionCode: "",
      fgQty: "",
      remark: "",
    };

    Object.entries(values).forEach(([key, value]) => {
      const field = elements.form.elements.namedItem(key);
      if (!field) {
        return;
      }
      field.value = value === null ? "" : value;
    });

    elements.dialog.showModal();
  }

  function closeDialog() {
    if (!elements.dialog) {
      return;
    }
    elements.dialog.close();
    state.editingId = null;
  }

  function enrichRecordFromCatalog(record) {
    const detail = defectMap.get(record.typeCode) || {};
    const defects = Array.isArray(detail.defects) ? detail.defects : [];
    return {
      ...record,
      feature: detail.feature || "",
      description: detail.description || "",
      productClass: detail.productClass || "",
      defects,
      defectCount: defects.length,
      totalIntensity: round(sum(defects.map((defect) => defect.intensity || 0)), 2),
    };
  }

  function buildFlowPayload(record, action) {
    const noValue = cleanText(record.no);
    const rowIdValue = cleanText(record.rowId || record.id || record.no);
    return {
      "No.": cleanText(noValue),
      RowId: rowIdValue,
      rowId: rowIdValue,
      id: rowIdValue,
      no: cleanText(noValue),
      sourceRow: record.sourceRow ?? "",
      season: cleanText(record.season),
      category: cleanText(record.category),
      protoStage: cleanText(record.protoStage),
      style: cleanText(record.style),
      constructionCode: cleanText(record.constructionCode),
      typeCode: cleanText(record.typeCode),
      modification: cleanText(record.modification),
      styleKey: cleanText(record.styleKey),
      remark: cleanText(record.remark),
      fgQty: record.fgQty ?? "",
      fgAnchor: record.fgAnchor ?? "",
      action,
    };
  }

  async function callFlow(action, payload) {
    const url = FLOW_ENDPOINTS[action];
    if (!url) {
      throw new Error(`Missing Power Automate URL for ${action}.`);
    }

    const response = await window.fetch(url, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify(payload),
    });

    if (!response.ok) {
      const text = await response.text().catch(() => "");
      throw new Error(text || `Power Automate ${action} failed with ${response.status}.`);
    }

    return response;
  }

  function isNoResponseError(error) {
    return /NoResponse/i.test(String(error && error.message ? error.message : error));
  }

  function valueFromRow(row, keys) {
    for (const key of keys) {
      if (row && Object.prototype.hasOwnProperty.call(row, key)) {
        return row[key];
      }
    }
    return "";
  }

  function normalizeFieldName(value) {
    return String(value || "")
      .replace(/_x[0-9a-f]{4}_/gi, " ")
      .replace(/[^a-z0-9]+/gi, "")
      .toLowerCase();
  }

  function valueFromRowFuzzy(row, aliases) {
    if (!row || typeof row !== "object") {
      return "";
    }
    const aliasSet = new Set(aliases.map(normalizeFieldName));
    for (const [key, value] of Object.entries(row)) {
      if (aliasSet.has(normalizeFieldName(key))) {
        return value;
      }
    }
    return "";
  }

  function normalizeRemoteRecord(row) {
    return {
      rowId:
        valueFromRow(row, ["RowId", "rowId", "ROWID", "ROWID", "Row_x0020_Id"]) ||
        valueFromRowFuzzy(row, ["RowId", "rowId", "record id", "recordid"]),
      id:
        valueFromRow(row, ["RowId", "rowId", "ROWID", "ROWID", "Row_x0020_Id"]) ||
        valueFromRowFuzzy(row, ["RowId", "rowId", "record id", "recordid"]),
      no:
        valueFromRow(row, ["No.", "No", "NO", "no", "No_x002e_"]) ||
        valueFromRowFuzzy(row, ["No.", "No"]),
      season: valueFromRow(row, ["SEASON", "season"]) || valueFromRowFuzzy(row, ["SEASON", "season"]),
      category:
        valueFromRow(row, ["CATEGORY", "category"]) || valueFromRowFuzzy(row, ["CATEGORY", "category"]),
      protoStage:
        valueFromRow(
          row,
          ["PROTO STAGE", "PROTO\n STAGE", "protoStage", "PROTO_x0020_STAGE", "PROTO_x000a__x0020_STAGE"],
        ) || valueFromRowFuzzy(row, ["PROTO STAGE", "protoStage", "protostage"]),
      style: valueFromRow(row, ["STYLE", "style"]) || valueFromRowFuzzy(row, ["STYLE", "style"]),
      constructionCode:
        valueFromRow(
          row,
          ["CONSTRUCTION CODE", "CONSTRUCTION\n CODE", "constructionCode", "CONSTRUCTION_x0020_CODE"],
        ) || valueFromRowFuzzy(row, ["CONSTRUCTION CODE", "constructionCode", "constructioncode"]),
      typeCode: valueFromRow(row, ["TYPE", "typeCode"]) || valueFromRowFuzzy(row, ["TYPE", "typeCode"]),
      modification:
        valueFromRow(
          row,
          [
            "Construction Modification",
            "Construction \nModification",
            "constructionModification",
            "modification",
            "Construction_x0020_Modification",
          ],
        ) || valueFromRowFuzzy(row, ["Construction Modification", "constructionModification", "modification"]),
      remark: valueFromRow(row, ["REMARK", "remark"]) || valueFromRowFuzzy(row, ["REMARK", "remark"]),
      fgQty: valueFromRow(row, ["FG Qty", "fgQty", "FG_x0020_Qty"]) || valueFromRowFuzzy(row, ["FG Qty", "fgQty"]),
    };
  }

  function normalizeSummaryRow(row) {
    return {
      constructionCode:
        valueFromRow(row, ["Construction Code", "constructionCode"]) ||
        valueFromRowFuzzy(row, ["Construction Code", "constructionCode"]),
      currentTotalFgQty:
        valueFromRow(row, ["Current Total FG QTY", "currentTotalFgQty"]) ||
        valueFromRowFuzzy(row, ["Current Total FG QTY", "currentTotalFgQty"]),
      currentRank:
        valueFromRow(row, ["Current Rank", "currentRank"]) ||
        valueFromRowFuzzy(row, ["Current Rank", "currentRank"]),
      samImprovement:
        valueFromRow(row, ["SAM Improvement", "samImprovement"]) ||
        valueFromRowFuzzy(row, ["SAM Improvement", "samImprovement"]),
      improvementType:
        valueFromRow(row, ["Improvement Type", "improvementType"]) ||
        valueFromRowFuzzy(row, ["Improvement Type", "improvementType"]),
      improvementValue:
        valueFromRow(row, ["Improvement Value", "improvementValue", "Improvement_x0020_Value"]) ||
        valueFromRowFuzzy(row, ["Improvement Value", "improvementValue"]),
      investmentDecision:
        valueFromRow(row, ["Investment Decision", "investmentDecision"]) ||
        valueFromRowFuzzy(row, ["Investment Decision", "investmentDecision"]),
      updatedBy:
        valueFromRow(row, ["Updated By", "updatedBy"]) ||
        valueFromRowFuzzy(row, ["Updated By", "updatedBy"]),
      updatedAt:
        valueFromRow(row, ["Updated At", "updatedAt"]) ||
        valueFromRowFuzzy(row, ["Updated At", "updatedAt"]),
      activeTop20:
        valueFromRow(row, ["Active Top 20", "activeTop20"]) ||
        valueFromRowFuzzy(row, ["Active Top 20", "activeTop20"]),
    };
  }

  function buildInvestmentNotesFromSummaryRows(rows) {
    const nextNotes = {};
    rows
      .map(normalizeSummaryRow)
      .forEach((row) => {
        const code = cleanText(row.constructionCode);
        if (!code) {
          return;
        }
        nextNotes[code] = {
          currentTotalFgQty: cleanText(row.currentTotalFgQty),
          currentRank: cleanText(row.currentRank),
          samImprovement: cleanText(row.samImprovement),
          improvementType: cleanText(row.improvementType),
          improvementValue: cleanText(row.improvementValue),
          investmentDecision: cleanText(row.investmentDecision),
          updatedBy: cleanText(row.updatedBy),
          updatedAt: cleanText(row.updatedAt),
          activeTop20: cleanText(row.activeTop20),
        };
      });
    return nextNotes;
  }

  function extractRemoteRows(payload) {
    if (typeof payload === "string") {
      try {
        return extractRemoteRows(JSON.parse(payload));
      } catch (error) {
        return [];
      }
    }
    if (Array.isArray(payload)) {
      return payload;
    }
    if (payload && Array.isArray(payload.value)) {
      return payload.value;
    }
    if (payload && Array.isArray(payload.rows)) {
      return payload.rows;
    }
    if (payload && payload.body && Array.isArray(payload.body.value)) {
      return payload.body.value;
    }
    if (payload && payload.body && Array.isArray(payload.body.rows)) {
      return payload.body.rows;
    }
    if (payload && payload.body && typeof payload.body === "string") {
      return extractRemoteRows(payload.body);
    }
    if (payload && payload.data) {
      return extractRemoteRows(payload.data);
    }
    if (payload && payload.result) {
      return extractRemoteRows(payload.result);
    }
    if (payload && typeof payload === "object") {
      for (const value of Object.values(payload)) {
        const rows = extractRemoteRows(value);
        if (rows.length) {
          return rows;
        }
      }
    }
    return [];
  }

  function applySavedRecord(nextRecords, enriched, recordIndex) {
    if (recordIndex >= 0) {
      nextRecords[recordIndex] = enriched;
    } else {
      nextRecords.unshift(enriched);
    }

    nextRecords.forEach((record) => {
      if (record.styleKey === enriched.styleKey) {
        record.fgQty = enriched.fgQty;
      }
    });

    state.records = normalizeRecords(nextRecords);
    persistRecords();
    markPendingUpsert(enriched.id);
    closeDialog();
    render();
    window.setTimeout(() => {
      refreshFromLatestSeed({ silent: true });
    }, 1500);
  }

  function applyDeletedRecord(recordId) {
    state.records = state.records.filter((item) => item.id !== recordId);
    persistRecords();
    markPendingDelete(recordId);
    render();
    window.setTimeout(() => {
      refreshFromLatestSeed({ silent: true });
    }, 1500);
  }

  async function fetchLatestSeedData(options = {}) {
    const allowSeedFallback = options.allowSeedFallback !== false;
    let summaryRows = [];
    try {
      const [remoteResponse, summaryResponse] = await Promise.all([
        window.fetch(FETCH_DFM_CHART_URL, {
          method: "POST",
          headers: {
            "Content-Type": "application/json",
          },
          body: "{}",
          cache: "no-store",
        }),
        window.fetch(FETCH_DFM_SUMMARY_URL, {
          method: "POST",
          headers: {
            "Content-Type": "application/json",
          },
          body: "{}",
          cache: "no-store",
        }).catch(() => null),
      ]);
      if (summaryResponse && summaryResponse.ok) {
        const summaryText = await summaryResponse.text();
        let summaryPayload = null;
        try {
          summaryPayload = summaryText ? JSON.parse(summaryText) : null;
        } catch (error) {
          summaryPayload = summaryText;
        }
        summaryRows = extractRemoteRows(summaryPayload);
      }

      if (remoteResponse.ok) {
        const remoteText = await remoteResponse.text();
        let remotePayload = null;
        try {
          remotePayload = remoteText ? JSON.parse(remoteText) : null;
        } catch (error) {
          remotePayload = remoteText;
        }
        const remoteRows = extractRemoteRows(remotePayload);
        const normalizedRemoteRecords = normalizeRecords(remoteRows.map(normalizeRemoteRecord));
        if (normalizedRemoteRecords.length) {
          return {
            meta: {
              sourceFile: "Power Automate Fetch DFM chart",
              recordCount: normalizedRemoteRecords.length,
            },
            records: normalizedRemoteRecords,
            summaryRows,
          };
        }
      }
    } catch (error) {
      console.error("Power Automate fetch failed, falling back to seed file", error);
    }

    if (!allowSeedFallback) {
      throw new Error("Live Excel fetch returned no usable records.");
    }

    const response = await window.fetch(`./seed-data.js?t=${Date.now()}`, { cache: "no-store" });
    if (!response.ok) {
      throw new Error(`Latest seed fetch failed with ${response.status}.`);
    }

    const scriptText = await response.text();
    const match = scriptText.match(/window\.DFM_SEED_DATA\s*=\s*(.*);\s*$/s);
    if (!match) {
      throw new Error("Latest seed payload could not be parsed.");
    }

    return {
      ...JSON.parse(match[1]),
      summaryRows,
    };
  }

  async function refreshFromLatestSeed(options = {}) {
    try {
      const latestSeed = await fetchLatestSeedData();
      state.records = normalizeRecords(reconcileStoredRecords(state.records, latestSeed.records || []));
      state.records = normalizeRecords(mergeRemoteWithPending(state.records, state.records));
      if ((latestSeed.summaryRows || []).length) {
        state.investmentNotes = buildInvestmentNotesFromSummaryRows(latestSeed.summaryRows || []);
        persistInvestmentNotes();
      }
      state.latestSource = latestSeed?.meta?.sourceFile || "Seed file";
      state.lastRefreshAt = Date.now();
      state.syncStatus = "Synced";
      persistRecords();
      if (options.render !== false) {
        render();
      }
    } catch (error) {
      state.syncStatus = "Refresh failed";
      if (!options.silent) {
        console.error("Failed to refresh latest seed data", error);
      }
    }
  }

  async function hardRefreshFromLatestSeed(options = {}) {
    try {
      setFormBusy(true);
      state.syncStatus = "Refreshing";
      render();
      let latestSeed;
      try {
        latestSeed = await fetchLatestSeedData({ allowSeedFallback: false });
      } catch (liveError) {
        console.error("Live refresh returned no usable rows, falling back to seed data", liveError);
        latestSeed = await fetchLatestSeedData({ allowSeedFallback: true });
      }
      const mergedRecords = normalizeRecords(mergeRemoteWithPending(latestSeed.records || [], state.records));
      if (!mergedRecords.length) {
        const fallbackRecords = normalizeRecords(seedData.records || []);
        if (fallbackRecords.length) {
          state.records = fallbackRecords;
          state.latestSource = "Seed file";
          state.lastRefreshAt = Date.now();
          state.syncStatus = "Seed fallback";
          persistRecords();
          render();
          return;
        }
        throw new Error("No usable records were returned from live or seed data.");
      }
      state.records = mergedRecords;
      if ((latestSeed.summaryRows || []).length) {
        state.investmentNotes = buildInvestmentNotesFromSummaryRows(latestSeed.summaryRows || []);
        persistInvestmentNotes();
      }
      state.latestSource = latestSeed?.meta?.sourceFile || "Live Excel fetch";
      state.lastRefreshAt = Date.now();
      state.syncStatus = "Synced";
      persistRecords();
      render();
    } catch (error) {
      console.error("Failed to replace records from latest Excel-backed source", error);
      state.syncStatus = "Refresh failed";
      if (!options.silent) {
        window.alert(`Unable to refresh the latest Excel data.\n${error.message}`);
      }
    } finally {
      setFormBusy(false);
    }
  }

  async function saveRecord(event) {
    if (!elements.form) {
      return;
    }
    event.preventDefault();
    if (state.isSaving) {
      return;
    }
    const formData = new FormData(elements.form);
    const currentRecord = state.records.find((record) => record.id === state.editingId) || null;
    const constructionCode = cleanText(formData.get("constructionCode"));
    const draft = normalizeRecord(
      {
        id: state.editingId || nextRowId(),
        no: currentRecord?.no || "",
        sourceRow: currentRecord?.sourceRow || null,
        season: formData.get("season"),
        category: formData.get("category"),
        protoStage: formData.get("protoStage"),
        style: formData.get("style"),
        constructionCode,
        typeCode: currentRecord?.typeCode || constructionCode,
        modification: currentRecord?.modification || "",
        remark: formData.get("remark"),
        fgQty: formData.get("fgQty"),
      },
      0,
    );

    const enriched = enrichRecordFromCatalog(draft);
    const nextRecords = state.records.slice();
    const recordIndex = nextRecords.findIndex((record) => record.id === enriched.id);
    const existingRecord = recordIndex >= 0 ? nextRecords[recordIndex] : null;
    const flowAction = recordIndex >= 0 ? "update" : "add";

    try {
      setFormBusy(true);
      await callFlow(flowAction, buildFlowPayload(enriched, flowAction));
      applySavedRecord(nextRecords, enriched, recordIndex);
    } catch (error) {
      const allowRetryAsAdd =
        flowAction === "update" &&
        existingRecord &&
        (!existingRecord.sourceRow || /^row-\d+$/i.test(cleanText(existingRecord.id)));

      if (allowRetryAsAdd && /No row was found|NotFound/i.test(error.message || "")) {
        try {
          await callFlow("add", buildFlowPayload(enriched, "add"));
          applySavedRecord(nextRecords, enriched, recordIndex);
          return;
        } catch (retryError) {
          if (isNoResponseError(retryError)) {
            applySavedRecord(nextRecords, enriched, recordIndex);
            return;
          }
          console.error(retryError);
          window.alert(`Unable to save record to OneDrive Excel.\n${retryError.message}`);
          return;
        }
      }

      if (isNoResponseError(error)) {
        applySavedRecord(nextRecords, enriched, recordIndex);
        return;
      }

      console.error(error);
      window.alert(`Unable to save record to OneDrive Excel.\n${error.message}`);
    } finally {
      setFormBusy(false);
    }
  }

  async function deleteRecord(recordId) {
    const record = state.records.find((item) => item.id === recordId);
    if (!record) {
      return;
    }
    if (state.isSaving) {
      return;
    }
    const confirmed = window.confirm(`Delete ${record.season} / ${record.style} / ${record.constructionCode}?`);
    if (!confirmed) {
      return;
    }

    try {
      setFormBusy(true);
      const deletePayload = buildFlowPayload(record, "delete");
      await callFlow("delete", deletePayload);
      applyDeletedRecord(recordId);
    } catch (error) {
      if (isNoResponseError(error)) {
        applyDeletedRecord(recordId);
        return;
      }
      console.error(error);
      window.alert(`Unable to delete record from OneDrive Excel.\n${error.message}`);
    } finally {
      setFormBusy(false);
    }
  }

  function resetData() {
    const confirmed = window.confirm("Reset all edits and reload the original workbook seed data?");
    if (!confirmed) {
      return;
    }
    window.localStorage.removeItem(STORAGE_KEY);
    window.localStorage.removeItem(SYNC_META_KEY);
    window.localStorage.removeItem(INVESTMENT_NOTES_KEY);
    Object.keys(window.localStorage)
      .filter((key) => key.startsWith(LEGACY_STORAGE_PREFIX))
      .forEach((key) => window.localStorage.removeItem(key));
    state.records = normalizeRecords(seedData.records || []);
    state.syncMeta = { pendingUpserts: {}, pendingDeletes: {} };
    state.investmentNotes = {};
    render();
  }

  function bindEvents() {
    if (elements.searchInput) {
      elements.searchInput.addEventListener("input", (event) => {
        state.filters.search = event.target.value.trim();
        render();
      });
    }

    if (elements.seasonFilter) {
      elements.seasonFilter.addEventListener("change", (event) => {
        state.filters.season = event.target.value;
        render();
      });
    }

    if (elements.categoryFilter) {
      elements.categoryFilter.addEventListener("change", (event) => {
        state.filters.category = event.target.value;
        render();
      });
    }

    if (elements.modificationFilter) {
      elements.modificationFilter.addEventListener("change", (event) => {
        state.filters.modification = event.target.value;
        render();
      });
    }

    if (elements.defectFilter) {
      elements.defectFilter.addEventListener("change", (event) => {
        state.filters.defectOnly = event.target.checked;
        render();
      });
    }

    if (elements.addRecordBtn) {
      elements.addRecordBtn.addEventListener("click", () => openDialog());
    }
    if (elements.refreshDataBtn) {
      elements.refreshDataBtn.addEventListener("click", () => {
        hardRefreshFromLatestSeed();
      });
    }
    if (elements.resetDataBtn) {
      elements.resetDataBtn.addEventListener("click", resetData);
    }
    if (elements.closeDialogBtn) {
      elements.closeDialogBtn.addEventListener("click", closeDialog);
    }
    if (elements.cancelDialogBtn) {
      elements.cancelDialogBtn.addEventListener("click", closeDialog);
    }
    if (elements.form) {
      elements.form.addEventListener("submit", saveRecord);
    }

    if (elements.recordCards) {
      elements.recordCards.addEventListener("click", (event) => {
        const button = event.target.closest("button[data-action]");
        if (!button) {
          return;
        }
        const recordId = button.dataset.id;
        if (button.dataset.action === "edit") {
          openDialog(recordId);
        }
        if (button.dataset.action === "delete") {
          deleteRecord(recordId);
        }
      });
    }

    if (elements.investmentBoard) {
      elements.investmentBoard.addEventListener("click", (event) => {
        const button = event.target.closest('button[data-action]');
        if (!button) {
          return;
        }
        const row = button.closest("tr[data-code]");
        if (!row) {
          return;
        }
        const code = row.dataset.code || "";
        if (button.dataset.action === "edit-investment") {
          state.investmentEditingCode = code;
          render();
          return;
        }
        const samValue = row.querySelector('[data-role="sam"]')?.value || "";
        const improvementTypeValue = row.querySelector('[data-role="improvement-type"]')?.value || "";
        const decisionValue = row.querySelector('[data-role="decision"]')?.value || "";
        saveInvestmentRow(code, samValue, improvementTypeValue, decisionValue).finally(() => {
          state.investmentEditingCode = null;
          render();
        });
      });
    }

    if (elements.investmentControls) {
      elements.investmentControls.addEventListener("click", (event) => {
        const button = event.target.closest('button[data-action="toggle-investment-column"]');
        if (!button) {
          return;
        }
        const column = button.dataset.column;
        if (!column) {
          return;
        }
        state.investmentVisibility[column] = !(state.investmentVisibility[column] !== false);
        persistInvestmentVisibility();
        render();
      });
    }
  }

  bindEvents();
  if (page === "dashboard") {
    state.rotation.timerId = window.setInterval(() => {
      state.rotation.tick += 1;
      render();
    }, state.rotation.seconds * 1000);
  }
  if (AUTO_REFRESH_MS > 0) {
    state.seedRefreshTimerId = window.setInterval(() => {
      refreshFromLatestSeed({ silent: true });
    }, AUTO_REFRESH_MS);
  }
  render();
  refreshFromLatestSeed({ silent: true });
})();
