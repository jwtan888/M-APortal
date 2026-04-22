(function () {
  const seedData = window.DFM_SEED_DATA || { records: [], warnings: [], defectCatalog: [] };
  const seedVersion =
    seedData.meta?.generatedAt ||
    `${seedData.meta?.recordCount || 0}-${seedData.meta?.styleSeasonCount || 0}`;
  const STORAGE_KEY = `dfm-dashboard-records-${seedVersion}`;
  const defectMap = new Map((seedData.defectCatalog || []).map((item) => [item.code, item]));
  const page = document.body.dataset.page || "dashboard";

  const elements = {
    seasonDonut: document.getElementById("season-donut"),
    seasonDonutTotal: document.getElementById("season-donut-total"),
    seasonDonutLegend: document.getElementById("season-donut-legend"),
    categoryBars: document.getElementById("category-bars"),
    seasonTiles: document.getElementById("season-tiles"),
    kpiGrid: document.getElementById("kpi-grid"),
    filteredSummary: document.getElementById("filtered-summary"),
    seasonBars: document.getElementById("season-bars"),
    codeBars: document.getElementById("code-bars"),
    modSummary: document.getElementById("mod-summary"),
    styleCards: document.getElementById("style-cards"),
    codeDirectory: document.getElementById("code-directory"),
    recordCards: document.getElementById("record-cards"),
    searchInput: document.getElementById("search-input"),
    seasonFilter: document.getElementById("season-filter"),
    categoryFilter: document.getElementById("category-filter"),
    modificationFilter: document.getElementById("modification-filter"),
    defectFilter: document.getElementById("defect-filter"),
    addRecordBtn: document.getElementById("add-record-btn"),
    resetDataBtn: document.getElementById("reset-data-btn"),
    dialog: document.getElementById("record-dialog"),
    dialogTitle: document.getElementById("dialog-title"),
    closeDialogBtn: document.getElementById("close-dialog-btn"),
    cancelDialogBtn: document.getElementById("cancel-dialog-btn"),
    form: document.getElementById("record-form"),
  };

  const state = {
    records: loadRecords(),
    editingId: null,
    rotation: {
      seconds: 30,
      timerId: null,
      tick: 0,
    },
    filters: {
      search: "",
      season: "all",
      category: "all",
      modification: "all",
      defectOnly: false,
    },
  };

  function loadRecords() {
    const raw = window.localStorage.getItem(STORAGE_KEY);
    if (!raw) {
      return normalizeRecords(seedData.records || []);
    }

    try {
      return normalizeRecords(JSON.parse(raw));
    } catch (error) {
      console.error("Failed to parse saved data", error);
      return normalizeRecords(seedData.records || []);
    }
  }

  function persistRecords() {
    window.localStorage.setItem(STORAGE_KEY, JSON.stringify(state.records));
  }

  function normalizeRecords(records) {
    return records
      .map((record, index) => normalizeRecord(record, index))
      .filter((record) => !isBlankRecord(record));
  }

  function normalizeRecord(record, index) {
    const season = cleanText(record.season);
    const style = cleanText(record.style);
    const typeCode = cleanText(record.typeCode || record.constructionCode);
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
      id: cleanText(record.id) || "manual-" + (Date.now() + index),
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

  function buildStyleKey(season, style) {
    return cleanText(season) + "__" + cleanText(style);
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
          fgQty: record.fgQty || 0,
          constructionRows: 0,
          modifiedRows: 0,
          codes: new Set(),
        });
      }

      const style = styles.get(record.styleKey);
      style.constructionRows += 1;
      style.modifiedRows += record.modification === "M" ? 1 : 0;
      if (style.fgQty === 0 && record.fgQty) {
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
    records.forEach((record) => {
      const key = record.constructionCode || "Unknown";
      codeMap.set(key, (codeMap.get(key) || 0) + 1);
      const seasonKey = record.season || "Unknown";
      seasonCodeMap.set(seasonKey, (seasonCodeMap.get(seasonKey) || 0) + 1);
    });

    const mCount = records.filter((record) => record.modification === "M").length;
    const nonMCount = records.filter((record) => record.modification === "Non-M").length;
    const otherCount = Math.max(0, records.length - mCount - nonMCount);

    return {
      styleSummary: styleSummary.sort((a, b) => b.fgQty - a.fgQty),
      kpis: [
        {
          label: "Season + Style Count",
          value: formatNumber(styleSummary.length),
          subtext: "FG counted once per style key",
        },
        {
          label: "Unique FG Qty",
          value: formatNumber(totalFg),
          subtext: "No FG accumulation across construction rows",
        },
        {
          label: "construction occuracy",
          value: formatNumber(records.length),
          subtext: `${formatPercent(modificationRate)} modified rows`,
        },
        {
          label: "Construction Codes",
          value: formatNumber(codeMap.size),
          subtext: "Unique construction codes in filter",
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
        .slice(0, 5),
      codeBars: Array.from(codeMap.entries())
        .map(([label, value]) => ({ label, value }))
        .sort((a, b) => b.value - a.value)
        .slice(0, 8),
      modSummary: [
        { label: "Modified", value: mCount, note: "Rows marked M" },
        { label: "Non-Modified", value: nonMCount, note: "Rows marked Non-M" },
        { label: "Unspecified", value: otherCount, note: "Rows without modification value" },
      ].filter((item) => item.value > 0),
      codeDirectory: Array.from(codeMap.entries())
        .map(([label, value]) => ({ label, value }))
        .sort((a, b) => b.value - a.value || a.label.localeCompare(b.label))
        .slice(0, 40),
    };
  }

  function render() {
    populateFilterOptions();

    const filteredRecords = filterRecords(state.records);
    const analytics = computeAnalytics(filteredRecords);

    if (page === "dashboard") {
      renderSeasonDonut(analytics.seasonDonut);
      renderBars(elements.categoryBars, analytics.categoryBars, formatNumber);
      renderSeasonTiles(analytics.seasonTiles);
      renderKpis(analytics.kpis);
      renderBars(elements.seasonBars, analytics.seasonBars, formatNumber);
      renderBars(elements.codeBars, analytics.codeBars, formatNumber);
      renderModSummary(analytics.modSummary);
      renderStyleCards(analytics.styleSummary);
      renderCodeDirectory(analytics.codeDirectory);
    }

    if (page === "data") {
      renderRecordCards(filteredRecords);
      if (elements.filteredSummary) {
        elements.filteredSummary.textContent = `${labelCount(filteredRecords.length, "construction row", "construction rows")} shown`;
      }
    }
  }

  function populateFilterOptions() {
    if (!elements.seasonFilter || !elements.categoryFilter || !elements.modificationFilter) {
      return;
    }
    setSelectOptions(elements.seasonFilter, ["all"].concat(uniqueValues(state.records, "season")), state.filters.season, "All seasons");
    setSelectOptions(elements.categoryFilter, ["all"].concat(uniqueValues(state.records, "category")), state.filters.category, "All categories");
    setSelectOptions(
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

  function renderSeasonDonut(items) {
    if (!elements.seasonDonut || !elements.seasonDonutLegend || !elements.seasonDonutTotal) {
      return;
    }
    if (!items.length) {
      elements.seasonDonut.style.background = "rgba(30, 87, 216, 0.08)";
      elements.seasonDonutLegend.innerHTML = '<div class="empty-state">No season data available.</div>';
      elements.seasonDonutTotal.textContent = "0";
      return;
    }

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
            <span class="legend-value ${index === activeIndex ? "is-active" : ""}">${escapeHtml(formatPercent(item.share))}</span>
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
    if (!items.length) {
      target.innerHTML = '<div class="empty-state">No data for the current filter set.</div>';
      return;
    }
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
    if (!items.length) {
      elements.modSummary.innerHTML = '<div class="empty-state">No modification summary available.</div>';
      return;
    }

    const activeIndex = items.length ? state.rotation.tick % items.length : 0;
    elements.modSummary.innerHTML = items
      .map(
        (item, index) => `
          <div class="stack-item ${index === activeIndex ? "is-active" : ""}">
            <div>${escapeHtml(item.label)}</div>
            <div class="stack-value">${escapeHtml(formatNumber(item.value))}</div>
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
    if (!items.length) {
      elements.seasonTiles.innerHTML = '<div class="empty-state">No season density available.</div>';
      return;
    }

    elements.seasonTiles.innerHTML = items
      .map(
        (item, index) => `
          <article class="metric-tile ${index === (items.length ? state.rotation.tick % items.length : 0) ? "is-active" : ""}">
            <div class="metric-tile-label">${escapeHtml(item.label || "Unknown")}</div>
            <div class="metric-tile-value">${escapeHtml(formatNumber(item.value))}</div>
            <div class="metric-tile-meta">${escapeHtml(labelCount(item.value, "style", "styles"))} · ${escapeHtml(formatNumber(item.fg))} FG</div>
          </article>
        `,
      )
      .join("");
  }

  function renderStyleCards(styles) {
    if (!elements.styleCards) {
      return;
    }
    if (!styles.length) {
      elements.styleCards.innerHTML = '<div class="empty-state">No styles match the current filters.</div>';
      return;
    }

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
              <span class="tag">${escapeHtml(formatNumber(style.fgQty))} FG</span>
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
    if (!items.length) {
      elements.codeDirectory.innerHTML = '<div class="empty-state">No construction codes in the current filter.</div>';
      return;
    }

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
            <div class="record-meta">FG Qty: <strong>${escapeHtml(formatNumber(record.fgQty))}</strong></div>
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
      typeCode: "",
      modification: "",
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

  function saveRecord(event) {
    if (!elements.form) {
      return;
    }
    event.preventDefault();
    const formData = new FormData(elements.form);
    const draft = normalizeRecord(
      {
        id: state.editingId || `manual-${Date.now()}`,
        sourceRow: null,
        season: formData.get("season"),
        category: formData.get("category"),
        protoStage: formData.get("protoStage"),
        style: formData.get("style"),
        constructionCode: formData.get("constructionCode"),
        typeCode: formData.get("typeCode"),
        modification: formData.get("modification"),
        remark: formData.get("remark"),
        fgQty: formData.get("fgQty"),
      },
      0,
    );

    const enriched = enrichRecordFromCatalog(draft);
    const nextRecords = state.records.slice();
    const recordIndex = nextRecords.findIndex((record) => record.id === enriched.id);

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
    closeDialog();
    render();
  }

  function deleteRecord(recordId) {
    const record = state.records.find((item) => item.id === recordId);
    if (!record) {
      return;
    }
    const confirmed = window.confirm(`Delete ${record.season} / ${record.style} / ${record.constructionCode}?`);
    if (!confirmed) {
      return;
    }
    state.records = state.records.filter((item) => item.id !== recordId);
    persistRecords();
    render();
  }

  function resetData() {
    const confirmed = window.confirm("Reset all edits and reload the original workbook seed data?");
    if (!confirmed) {
      return;
    }
    window.localStorage.removeItem(STORAGE_KEY);
    state.records = normalizeRecords(seedData.records || []);
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
  }

  bindEvents();
  if (page === "dashboard") {
    state.rotation.timerId = window.setInterval(() => {
      state.rotation.tick += 1;
      render();
    }, state.rotation.seconds * 1000);
  }
  render();
})();
