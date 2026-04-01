function generateUtilityPlanFloorId() {
  return `utility_floor_${Math.random().toString(36).slice(2, 11)}`;
}

function generateUtilityPlanCalloutId() {
  return `utility_callout_${Math.random().toString(36).slice(2, 11)}`;
}

function normalizeUtilityPlanNumber(value, fallback = 0) {
  const number = Number(value);
  return Number.isFinite(number) ? number : fallback;
}

function normalizeUtilityPlanRect(rect = {}) {
  return {
    x: Math.max(0, normalizeUtilityPlanNumber(rect?.x, 0)),
    y: Math.max(0, normalizeUtilityPlanNumber(rect?.y, 0)),
    width: Math.max(0, normalizeUtilityPlanNumber(rect?.width, 0)),
    height: Math.max(0, normalizeUtilityPlanNumber(rect?.height, 0)),
  };
}

function normalizeUtilityPlanPoint(point = {}) {
  return {
    x: Math.max(0, normalizeUtilityPlanNumber(point?.x, 0)),
    y: Math.max(0, normalizeUtilityPlanNumber(point?.y, 0)),
  };
}

function getUtilityPlanTextMeasureContext() {
  if (!utilityPlanTextMeasureCanvas) {
    utilityPlanTextMeasureCanvas = document.createElement("canvas");
  }
  return utilityPlanTextMeasureCanvas.getContext("2d");
}

function getUtilityPlanTextLines(text) {
  return String(text || "")
    .replace(/\r\n/g, "\n")
    .replace(/\r/g, "\n")
    .split("\n")
    .map((line) => line.trim())
    .filter(Boolean);
}

function measureUtilityPlanText(text) {
  const ctx = getUtilityPlanTextMeasureContext();
  const lines = getUtilityPlanTextLines(text);
  const normalizedLines = lines.length ? lines : ["Callout"];
  ctx.font = `700 ${UTILITY_PLAN_LABEL_FONT_SIZE}px "Segoe UI", Arial, sans-serif`;
  const width = Math.max(
    72,
    ...normalizedLines.map((line) => ctx.measureText(line || " ").width)
  );
  const height = Math.max(
    32,
    normalizedLines.length * UTILITY_PLAN_LABEL_FONT_SIZE * UTILITY_PLAN_LINE_HEIGHT
  );
  return {
    width: Math.ceil(width + UTILITY_PLAN_LABEL_PADDING_X * 2),
    height: Math.ceil(height + UTILITY_PLAN_LABEL_PADDING_Y * 2),
    lines: normalizedLines,
  };
}

function clampUtilityPlanRectToPage(rect, pageWidth, pageHeight, { defaultFullPage = false } = {}) {
  const normalized = normalizeUtilityPlanRect(rect);
  if (
    defaultFullPage &&
    (normalized.width <= 0 || normalized.height <= 0)
  ) {
    return {
      x: 0,
      y: 0,
      width: Math.max(0, pageWidth),
      height: Math.max(0, pageHeight),
    };
  }
  const x = Math.min(Math.max(normalized.x, 0), Math.max(0, pageWidth));
  const y = Math.min(Math.max(normalized.y, 0), Math.max(0, pageHeight));
  const width = Math.min(
    Math.max(normalized.width, 0),
    Math.max(0, pageWidth - x)
  );
  const height = Math.min(
    Math.max(normalized.height, 0),
    Math.max(0, pageHeight - y)
  );
  return { x, y, width, height };
}

function clampUtilityPlanPointToPage(point, pageWidth, pageHeight) {
  const normalized = normalizeUtilityPlanPoint(point);
  return {
    x: Math.min(Math.max(normalized.x, 0), Math.max(0, pageWidth)),
    y: Math.min(Math.max(normalized.y, 0), Math.max(0, pageHeight)),
  };
}

function fitUtilityPlanLabelRectToPage(rect, pageWidth, pageHeight) {
  const normalized = normalizeUtilityPlanRect(rect);
  const width = Math.min(Math.max(normalized.width, 72), Math.max(72, pageWidth));
  const height = Math.min(Math.max(normalized.height, 32), Math.max(32, pageHeight));
  return {
    x: Math.min(Math.max(normalized.x, 0), Math.max(0, pageWidth - width)),
    y: Math.min(Math.max(normalized.y, 0), Math.max(0, pageHeight - height)),
    width,
    height,
  };
}

function createUtilityPlanLabelRect(text, targetPoint, pageWidth, pageHeight) {
  const size = measureUtilityPlanText(text);
  return fitUtilityPlanLabelRectToPage(
    {
      x: targetPoint.x + 18,
      y: targetPoint.y - size.height / 2,
      width: size.width,
      height: size.height,
    },
    pageWidth,
    pageHeight
  );
}

function createDefaultUtilityPlanCallout(seed = {}, pageWidth = 0, pageHeight = 0) {
  const text = String(seed?.text || "New Callout").trim() || "New Callout";
  const size = measureUtilityPlanText(text);
  const targetPoint = clampUtilityPlanPointToPage(
    seed?.targetPoint || { x: 80, y: 80 },
    pageWidth,
    pageHeight
  );
  const measuredRect = fitUtilityPlanLabelRectToPage(
    {
      ...normalizeUtilityPlanRect(seed?.labelRect),
      width: normalizeUtilityPlanNumber(seed?.labelRect?.width, size.width),
      height: normalizeUtilityPlanNumber(seed?.labelRect?.height, size.height),
    },
    pageWidth,
    pageHeight
  );
  return {
    id: String(seed?.id || generateUtilityPlanCalloutId()).trim() || generateUtilityPlanCalloutId(),
    text,
    labelRect:
      measuredRect.width > 0 && measuredRect.height > 0
        ? measuredRect
        : createUtilityPlanLabelRect(text, targetPoint, pageWidth, pageHeight),
    targetPoint,
  };
}

function createDefaultUtilityPlanFloor(seed = {}) {
  return {
    id: String(seed?.id || generateUtilityPlanFloorId()).trim() || generateUtilityPlanFloorId(),
    label: String(seed?.label || "Floor 1").trim() || "Floor 1",
    order: Math.max(0, Math.trunc(normalizeUtilityPlanNumber(seed?.order, 0))),
    pageNumber: Math.max(0, Math.trunc(normalizeUtilityPlanNumber(seed?.pageNumber, 0))),
    cropRect: normalizeUtilityPlanRect(seed?.cropRect),
    callouts: Array.isArray(seed?.callouts) ? [...seed.callouts] : [],
    exportPath: normalizeWindowsPath(seed?.exportPath || ""),
    lastExportedAtUtc: String(seed?.lastExportedAtUtc || "").trim(),
  };
}

function createDefaultSurveyReportDraft(projectId = "") {
  return {
    schemaVersion: UTILITY_PLAN_SCHEMA_VERSION,
    projectId: String(projectId || "").trim(),
    utilityPlan: {
      sourcePdfPath: "",
      exportFolderPath: "",
      floors: [createDefaultUtilityPlanFloor()],
    },
    findings: {},
    recommendations: { items: [] },
    photos: { items: [] },
  };
}

function normalizeUtilityPlanFloor(floor, pageWidth = 0, pageHeight = 0) {
  const nextFloor = createDefaultUtilityPlanFloor(floor);
  nextFloor.cropRect = clampUtilityPlanRectToPage(
    nextFloor.cropRect,
    pageWidth,
    pageHeight
  );
  nextFloor.callouts = (Array.isArray(floor?.callouts) ? floor.callouts : [])
    .filter(Boolean)
    .map((callout) =>
      createDefaultUtilityPlanCallout(callout, pageWidth, pageHeight)
    );
  return nextFloor;
}

function normalizeSurveyReportDraft(rawDraft, projectId = "") {
  const base = createDefaultSurveyReportDraft(projectId);
  const draft = rawDraft && typeof rawDraft === "object" ? rawDraft : {};
  const utilityPlan =
    draft.utilityPlan && typeof draft.utilityPlan === "object"
      ? draft.utilityPlan
      : {};
  const floors = (Array.isArray(utilityPlan.floors) ? utilityPlan.floors : [])
    .filter((floor) => floor && typeof floor === "object")
    .map((floor) => normalizeUtilityPlanFloor(floor))
    .sort((a, b) => {
      const orderDiff = Number(a.order || 0) - Number(b.order || 0);
      if (orderDiff !== 0) return orderDiff;
      return String(a.label || "").localeCompare(String(b.label || ""));
    });
  if (!floors.length) floors.push(createDefaultUtilityPlanFloor());
  floors.forEach((floor, index) => {
    floor.order = index;
  });
  return {
    ...base,
    schemaVersion:
      String(draft.schemaVersion || UTILITY_PLAN_SCHEMA_VERSION).trim() ||
      UTILITY_PLAN_SCHEMA_VERSION,
    projectId: String(draft.projectId || projectId || "").trim(),
    utilityPlan: {
      sourcePdfPath: normalizeWindowsPath(utilityPlan.sourcePdfPath || ""),
      exportFolderPath: normalizeWindowsPath(utilityPlan.exportFolderPath || ""),
      floors,
    },
    version: normalizeUtilityPlanNumber(draft.version, 0),
    updatedAtUtc: String(draft.updatedAtUtc || "").trim(),
    updatedBy: String(draft.updatedBy || "").trim(),
  };
}

function buildCanonicalSurveyReportDraft(draft) {
  const normalized = normalizeSurveyReportDraft(
    draft,
    draft?.projectId || getUtilityPlanProjectId(getActiveUtilityPlanProject())
  );
  return {
    schemaVersion: normalized.schemaVersion,
    projectId: normalized.projectId,
    utilityPlan: {
      sourcePdfPath: normalized.utilityPlan.sourcePdfPath,
      exportFolderPath: normalized.utilityPlan.exportFolderPath,
      floors: normalized.utilityPlan.floors.map((floor, index) => ({
        id: String(floor.id || "").trim() || generateUtilityPlanFloorId(),
        label:
          String(floor.label || `Floor ${index + 1}`).trim() || `Floor ${index + 1}`,
        order: index,
        pageNumber: Math.max(
          0,
          Math.trunc(normalizeUtilityPlanNumber(floor.pageNumber, 0))
        ),
        cropRect: normalizeUtilityPlanRect(floor.cropRect),
        callouts: (Array.isArray(floor.callouts) ? floor.callouts : []).map(
          (callout) => ({
            id:
              String(callout?.id || "").trim() || generateUtilityPlanCalloutId(),
            text: String(callout?.text || "Callout").trim() || "Callout",
            labelRect: normalizeUtilityPlanRect(callout?.labelRect),
            targetPoint: normalizeUtilityPlanPoint(callout?.targetPoint),
          })
        ),
        exportPath: normalizeWindowsPath(floor.exportPath || ""),
        lastExportedAtUtc: String(floor.lastExportedAtUtc || "").trim(),
      })),
    },
    findings: {},
    recommendations: { items: [] },
    photos: { items: [] },
  };
}

function getActiveUtilityPlanProject() {
  if (utilityPlanProjectIndex == null) return null;
  return db[utilityPlanProjectIndex] || null;
}

function getUtilityPlanProjectId(project) {
  return getLightingScheduleProjectId(
    project || {},
    utilityPlanState?.draft?.utilityPlan || null
  );
}

function getUtilityPlanProjectLabel(project, index) {
  if (!project) return `Project ${index + 1}`;
  const id = String(project.id || "").trim();
  const name = String(project.name || "").trim();
  if (id && name) return `${id} - ${name}`;
  return id || name || `Project ${index + 1}`;
}

function getActiveUtilityPlanFloor() {
  const floors = utilityPlanState?.draft?.utilityPlan?.floors;
  if (!Array.isArray(floors) || !floors.length) return null;
  return (
    floors.find((floor) => floor.id === utilityPlanState.activeFloorId) ||
    floors[0] ||
    null
  );
}

function getSelectedUtilityPlanCallout() {
  const floor = getActiveUtilityPlanFloor();
  if (!floor) return null;
  return (
    floor.callouts.find(
      (callout) => callout.id === utilityPlanState.activeCalloutId
    ) || null
  );
}

function setUtilityPlanStatus(message) {
  utilityPlanStatusMessage = String(message || "").trim();
  const statusEl = document.getElementById("utilityPlanStatus");
  if (statusEl) {
    statusEl.textContent = utilityPlanStatusMessage || "Ready.";
  }
}

function ensureUtilityPlanActiveFloor() {
  const floors = utilityPlanState?.draft?.utilityPlan?.floors;
  if (!Array.isArray(floors) || !floors.length) {
    utilityPlanState.activeFloorId = "";
    return null;
  }
  const activeFloor =
    floors.find((floor) => floor.id === utilityPlanState.activeFloorId) ||
    floors[0];
  utilityPlanState.activeFloorId = activeFloor.id;
  if (
    utilityPlanState.activeCalloutId &&
    !activeFloor.callouts.some(
      (callout) => callout.id === utilityPlanState.activeCalloutId
    )
  ) {
    utilityPlanState.activeCalloutId = "";
  }
  return activeFloor;
}

function ensureUtilityPlanFloorCropForPreview(
  floor,
  preview,
  { forceReset = false } = {}
) {
  if (!floor || !preview) return;
  floor.cropRect = clampUtilityPlanRectToPage(
    floor.cropRect,
    preview.pageWidth,
    preview.pageHeight,
    {
      defaultFullPage:
        forceReset || floor.cropRect.width <= 0 || floor.cropRect.height <= 0,
    }
  );
  floor.callouts = floor.callouts.map((callout) =>
    createDefaultUtilityPlanCallout(
      callout,
      preview.pageWidth,
      preview.pageHeight
    )
  );
}

function clearUtilityPlanPreview() {
  utilityPlanState.preview = null;
  const image = document.getElementById("utilityPlanPageImage");
  const overlay = document.getElementById("utilityPlanOverlay");
  if (image) image.removeAttribute("src");
  if (overlay) overlay.innerHTML = "";
}

function getUtilityPlanProjectPathCandidates(project) {
  return [project?.path, project?.localProjectPath, project?.workroomRootPath]
    .map((value) => normalizeProjectPath(value || ""))
    .filter(Boolean);
}

function findUtilityPlanProjectIndexFromLaunchContext(launchContext) {
  const context =
    launchContext && typeof launchContext === "object" ? launchContext : null;
  const pathKey = getProjectBaseKey(
    context?.rootProjectPath || context?.projectPath || ""
  );
  if (pathKey) {
    const matchedIndex = db.findIndex((project) =>
      getUtilityPlanProjectPathCandidates(project).some(
        (candidate) => getProjectBaseKey(candidate) === pathKey
      )
    );
    if (matchedIndex >= 0) return matchedIndex;
  }

  const contextProjectId = extractLightingScheduleProjectIdFromPath(
    context?.rootProjectPath || context?.projectPath || ""
  );
  if (contextProjectId) {
    const matchedIndex = db.findIndex(
      (project) => String(project?.id || "").trim() === contextProjectId
    );
    if (matchedIndex >= 0) return matchedIndex;
  }

  const activeIndex = Number(activeChecklistProject);
  if (Number.isInteger(activeIndex) && activeIndex >= 0 && db[activeIndex]) {
    return activeIndex;
  }
  return null;
}

function renderUtilityPlanProjectOptions(filterText = "") {
  const select = document.getElementById("utilityPlanProjectSelect");
  if (!select) return;
  select.innerHTML = "";
  const query = String(filterText || "").trim().toLowerCase();
  const matches = db
    .map((project, index) => ({ project, index }))
    .filter(({ project }) => {
      if (!query) return true;
      const haystack = [
        project?.id,
        project?.name,
        project?.nick,
        project?.path,
      ]
        .map((value) => String(value || "").toLowerCase())
        .join(" ");
      return haystack.includes(query);
    });

  matches.forEach(({ project, index }) => {
    select.appendChild(
      el("option", {
        value: String(index),
        textContent: getUtilityPlanProjectLabel(project, index),
      })
    );
  });

  select.disabled = matches.length === 0;
  if (!matches.length) {
    utilityPlanProjectIndex = null;
    utilityPlanState.draft = null;
    utilityPlanState.activeFloorId = "";
    utilityPlanState.activeCalloutId = "";
    clearUtilityPlanPreview();
    renderUtilityPlanEditorUi();
    return;
  }

  const availableIndices = matches.map(({ index }) => index);
  if (availableIndices.includes(utilityPlanProjectIndex)) {
    select.value = String(utilityPlanProjectIndex);
    return;
  }

  utilityPlanProjectIndex = availableIndices[0];
  select.value = String(utilityPlanProjectIndex);
}

function escapeUtilityPlanSvgText(text) {
  return String(text || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function getUtilityPlanCalloutRenderRect(callout, pageWidth, pageHeight) {
  const baseRect = normalizeUtilityPlanRect(callout?.labelRect);
  const measured = measureUtilityPlanText(callout?.text || "Callout");
  return fitUtilityPlanLabelRectToPage(
    {
      ...baseRect,
      width: Math.max(baseRect.width, measured.width),
      height: Math.max(baseRect.height, measured.height),
    },
    pageWidth,
    pageHeight
  );
}

function getUtilityPlanLeaderAnchor(rect, targetPoint) {
  const cx = rect.x + rect.width / 2;
  const cy = rect.y + rect.height / 2;
  const dx = targetPoint.x - cx;
  const dy = targetPoint.y - cy;
  if (Math.abs(dx) < 0.0001 && Math.abs(dy) < 0.0001) {
    return { x: cx, y: cy };
  }
  const scaleX = rect.width > 0 ? Math.abs(dx) / (rect.width / 2) : Infinity;
  const scaleY = rect.height > 0 ? Math.abs(dy) / (rect.height / 2) : Infinity;
  const scale = Math.max(scaleX, scaleY, 0.0001);
  return {
    x: cx + dx / scale,
    y: cy + dy / scale,
  };
}

function getUtilityPlanArrowPoints(start, target, size = 8) {
  const dx = target.x - start.x;
  const dy = target.y - start.y;
  const length = Math.hypot(dx, dy);
  if (length < 0.001) return "";
  const ux = dx / length;
  const uy = dy / length;
  const px = -uy;
  const py = ux;
  const baseX = target.x - ux * size;
  const baseY = target.y - uy * size;
  return [
    `${target.x},${target.y}`,
    `${baseX + px * size * 0.55},${baseY + py * size * 0.55}`,
    `${baseX - px * size * 0.55},${baseY - py * size * 0.55}`,
  ].join(" ");
}

function buildUtilityPlanTextMarkup(text, rect) {
  const lines = getUtilityPlanTextLines(text);
  const normalizedLines = lines.length ? lines : ["Callout"];
  const startX = rect.x + UTILITY_PLAN_LABEL_PADDING_X;
  const startY =
    rect.y + UTILITY_PLAN_LABEL_PADDING_Y + UTILITY_PLAN_LABEL_FONT_SIZE;
  return normalizedLines
    .map((line, index) => {
      const dy =
        index === 0 ? 0 : UTILITY_PLAN_LABEL_FONT_SIZE * UTILITY_PLAN_LINE_HEIGHT;
      return `<tspan x="${startX}" dy="${dy}">${escapeUtilityPlanSvgText(
        line
      )}</tspan>`;
    })
    .join("");
}

function renderUtilityPlanOverlay() {
  const overlay = document.getElementById("utilityPlanOverlay");
  const preview = utilityPlanState.preview;
  const floor = getActiveUtilityPlanFloor();
  if (!overlay) return;
  if (!preview || !floor) {
    overlay.innerHTML = "";
    return;
  }

  const pageWidth = preview.pageWidth;
  const pageHeight = preview.pageHeight;
  overlay.setAttribute("viewBox", `0 0 ${pageWidth} ${pageHeight}`);
  overlay.setAttribute("width", String(pageWidth));
  overlay.setAttribute("height", String(pageHeight));

  const cropRect = clampUtilityPlanRectToPage(
    floor.cropRect,
    pageWidth,
    pageHeight,
    { defaultFullPage: true }
  );
  floor.cropRect = cropRect;
  const handleSize = 10;
  const selectedCalloutId = utilityPlanState.activeCalloutId;
  const callouts = [...floor.callouts].sort((a, b) => {
    if (a.id === selectedCalloutId) return 1;
    if (b.id === selectedCalloutId) return -1;
    return 0;
  });

  const cropMarkup = `
    <g class="utility-plan-crop-group">
      <rect class="utility-plan-crop-rect" data-crop-body="true"
        x="${cropRect.x}" y="${cropRect.y}" width="${cropRect.width}" height="${cropRect.height}"></rect>
      <rect class="utility-plan-crop-handle" data-crop-handle="nw" data-handle="nw"
        x="${cropRect.x - handleSize / 2}" y="${cropRect.y - handleSize / 2}" width="${handleSize}" height="${handleSize}"></rect>
      <rect class="utility-plan-crop-handle" data-crop-handle="ne" data-handle="ne"
        x="${cropRect.x + cropRect.width - handleSize / 2}" y="${cropRect.y - handleSize / 2}" width="${handleSize}" height="${handleSize}"></rect>
      <rect class="utility-plan-crop-handle" data-crop-handle="sw" data-handle="sw"
        x="${cropRect.x - handleSize / 2}" y="${cropRect.y + cropRect.height - handleSize / 2}" width="${handleSize}" height="${handleSize}"></rect>
      <rect class="utility-plan-crop-handle" data-crop-handle="se" data-handle="se"
        x="${cropRect.x + cropRect.width - handleSize / 2}" y="${cropRect.y + cropRect.height - handleSize / 2}" width="${handleSize}" height="${handleSize}"></rect>
    </g>
  `;

  const calloutMarkup = callouts
    .map((callout) => {
      callout.targetPoint = clampUtilityPlanPointToPage(
        callout.targetPoint,
        pageWidth,
        pageHeight
      );
      callout.labelRect = getUtilityPlanCalloutRenderRect(
        callout,
        pageWidth,
        pageHeight
      );
      const anchor = getUtilityPlanLeaderAnchor(
        callout.labelRect,
        callout.targetPoint
      );
      const arrowPoints = getUtilityPlanArrowPoints(
        anchor,
        callout.targetPoint,
        10
      );
      const selectedClass =
        callout.id === selectedCalloutId ? " is-selected" : "";
      return `
        <g class="utility-plan-callout${selectedClass}" data-callout-id="${callout.id}">
          <line class="utility-plan-callout-line" x1="${anchor.x}" y1="${anchor.y}" x2="${callout.targetPoint.x}" y2="${callout.targetPoint.y}"></line>
          <polygon class="utility-plan-callout-arrow" points="${arrowPoints}"></polygon>
          <rect class="utility-plan-callout-label" data-callout-label="true" data-callout-id="${callout.id}"
            x="${callout.labelRect.x}" y="${callout.labelRect.y}" width="${callout.labelRect.width}" height="${callout.labelRect.height}" rx="4" ry="4"></rect>
          <text class="utility-plan-callout-text" font-size="${UTILITY_PLAN_LABEL_FONT_SIZE}">
            ${buildUtilityPlanTextMarkup(callout.text, callout.labelRect)}
          </text>
          <circle class="utility-plan-callout-target" data-callout-target="true" data-callout-id="${callout.id}"
            cx="${callout.targetPoint.x}" cy="${callout.targetPoint.y}" r="7"></circle>
        </g>
      `;
    })
    .join("");

  overlay.innerHTML = `${cropMarkup}${calloutMarkup}`;
}

function applyUtilityPlanViewportTransform() {
  const viewport = document.getElementById("utilityPlanViewport");
  const surface = document.getElementById("utilityPlanSurface");
  const preview = utilityPlanState.preview;
  if (!viewport || !surface || !preview) return;
  surface.style.width = `${preview.pageWidth}px`;
  surface.style.height = `${preview.pageHeight}px`;
  surface.style.transform = `translate(${utilityPlanState.panX}px, ${utilityPlanState.panY}px) scale(${utilityPlanState.zoom})`;
  viewport.classList.toggle("is-pan-mode", utilityPlanState.tool === "pan");
  viewport.classList.toggle(
    "is-panning",
    utilityPlanState.interaction?.type === "pan"
  );
}

function fitUtilityPlanViewToPage() {
  const viewport = document.getElementById("utilityPlanViewport");
  const preview = utilityPlanState.preview;
  if (!viewport || !preview) return;
  const availableWidth = Math.max(200, viewport.clientWidth - 32);
  const availableHeight = Math.max(200, viewport.clientHeight - 32);
  const zoom = Math.min(
    UTILITY_PLAN_ZOOM_MAX,
    Math.max(
      UTILITY_PLAN_ZOOM_MIN,
      Math.min(
        availableWidth / preview.pageWidth,
        availableHeight / preview.pageHeight
      )
    )
  );
  utilityPlanState.zoom = zoom;
  utilityPlanState.panX = (viewport.clientWidth - preview.pageWidth * zoom) / 2;
  utilityPlanState.panY =
    (viewport.clientHeight - preview.pageHeight * zoom) / 2;
  applyUtilityPlanViewportTransform();
}

function updateUtilityPlanZoom(nextZoom, centerClientX = null, centerClientY = null) {
  const viewport = document.getElementById("utilityPlanViewport");
  const preview = utilityPlanState.preview;
  if (!viewport || !preview) return;
  const zoom = Math.min(
    UTILITY_PLAN_ZOOM_MAX,
    Math.max(UTILITY_PLAN_ZOOM_MIN, nextZoom)
  );
  if (centerClientX == null || centerClientY == null) {
    utilityPlanState.zoom = zoom;
    applyUtilityPlanViewportTransform();
    return;
  }
  const rect = viewport.getBoundingClientRect();
  const pdfX =
    (centerClientX - rect.left - utilityPlanState.panX) / utilityPlanState.zoom;
  const pdfY =
    (centerClientY - rect.top - utilityPlanState.panY) / utilityPlanState.zoom;
  utilityPlanState.zoom = zoom;
  utilityPlanState.panX = centerClientX - rect.left - pdfX * zoom;
  utilityPlanState.panY = centerClientY - rect.top - pdfY * zoom;
  applyUtilityPlanViewportTransform();
}

async function loadUtilityPlanPdfInfo(pdfPath) {
  if (!pdfPath) {
    utilityPlanState.pdfInfo = null;
    return null;
  }
  if (
    utilityPlanState.pdfInfo &&
    normalizeWindowsPath(utilityPlanState.pdfInfo.path || "") ===
      normalizeWindowsPath(pdfPath)
  ) {
    return utilityPlanState.pdfInfo;
  }
  if (!window.pywebview?.api?.get_utility_plan_pdf_info) {
    throw new Error("Utility plan PDF info API is unavailable.");
  }
  const response = await window.pywebview.api.get_utility_plan_pdf_info(pdfPath);
  if (response?.status !== "success" || !response.data) {
    throw new Error(response?.message || "Unable to inspect PDF.");
  }
  utilityPlanState.pdfInfo = response.data;
  return response.data;
}

async function loadUtilityPlanPreviewForActiveFloor({ resetView = false } = {}) {
  const draft = utilityPlanState.draft;
  const floor = ensureUtilityPlanActiveFloor();
  const pdfPath = normalizeWindowsPath(draft?.utilityPlan?.sourcePdfPath || "");
  if (!draft || !floor || !pdfPath) {
    utilityPlanState.pdfInfo = null;
    clearUtilityPlanPreview();
    renderUtilityPlanEditorUi();
    return null;
  }

  const requestId = ++utilityPlanPreviewRequestId;
  const pdfInfo = await loadUtilityPlanPdfInfo(pdfPath);
  const pageCount = Math.max(1, Number(pdfInfo?.pageCount || 1));
  floor.pageNumber = Math.min(
    Math.max(Math.trunc(normalizeUtilityPlanNumber(floor.pageNumber, 0)), 0),
    pageCount - 1
  );

  if (!window.pywebview?.api?.get_utility_plan_page_preview) {
    throw new Error("Utility plan preview API is unavailable.");
  }
  const response = await window.pywebview.api.get_utility_plan_page_preview(
    pdfPath,
    floor.pageNumber,
    UTILITY_PLAN_PREVIEW_MAX_SIZE
  );
  if (requestId !== utilityPlanPreviewRequestId) return null;
  if (response?.status !== "success" || !response.data) {
    throw new Error(response?.message || "Unable to render PDF preview.");
  }
  utilityPlanState.preview = response.data;
  ensureUtilityPlanFloorCropForPreview(floor, response.data, {
    forceReset: resetView,
  });
  renderUtilityPlanEditorUi();
  if (resetView) {
    requestAnimationFrame(() => fitUtilityPlanViewToPage());
  }
  return response.data;
}

function applyLoadedUtilityPlanDraft(record, projectId) {
  utilityPlanState.draft = normalizeSurveyReportDraft(
    record || createDefaultSurveyReportDraft(projectId),
    projectId
  );
  const floor = ensureUtilityPlanActiveFloor();
  if (floor) utilityPlanState.activeFloorId = floor.id;
}

async function setUtilityPlanProject(index, { quiet = false } = {}) {
  if (!Number.isInteger(index) || !db[index]) {
    utilityPlanProjectIndex = null;
    utilityPlanState.draft = null;
    utilityPlanState.activeFloorId = "";
    utilityPlanState.activeCalloutId = "";
    utilityPlanState.pdfInfo = null;
    clearUtilityPlanPreview();
    renderUtilityPlanEditorUi();
    return;
  }

  utilityPlanProjectIndex = index;
  utilityPlanState.activeCalloutId = "";
  utilityPlanState.interaction = null;
  utilityPlanState.pdfInfo = null;
  const project = db[index];
  const projectId = getUtilityPlanProjectId(project);
  applyLoadedUtilityPlanDraft(null, projectId);
  renderUtilityPlanEditorUi();

  try {
    if (!window.pywebview?.api?.get_survey_report_draft) {
      throw new Error("Survey report draft API is unavailable.");
    }
    const response = await window.pywebview.api.get_survey_report_draft(projectId);
    if (response?.status !== "success") {
      throw new Error(
        response?.message || "Unable to load survey report draft."
      );
    }
    applyLoadedUtilityPlanDraft(response?.data || null, projectId);
    renderUtilityPlanEditorUi();
    if (utilityPlanState.draft?.utilityPlan?.sourcePdfPath) {
      await loadUtilityPlanPreviewForActiveFloor({ resetView: true });
      if (!quiet) setUtilityPlanStatus("Utility plan draft loaded.");
    } else {
      clearUtilityPlanPreview();
      renderUtilityPlanEditorUi();
      if (!quiet) setUtilityPlanStatus("Select the source PDF for this project.");
    }
  } catch (error) {
    reportClientError("Failed to load utility plan draft", error);
    clearUtilityPlanPreview();
    renderUtilityPlanEditorUi();
    setUtilityPlanStatus(error?.message || "Unable to load utility plan draft.");
  }
}

async function saveUtilityPlanDraft({ quiet = false } = {}) {
  const project = getActiveUtilityPlanProject();
  const draft = utilityPlanState.draft;
  if (!project || !draft) return null;
  if (!window.pywebview?.api?.save_survey_report_draft) return draft;

  if (utilityPlanSaving) {
    utilityPlanAutosaveQueued = true;
    return null;
  }

  try {
    utilityPlanSaving = true;
    const projectId = getUtilityPlanProjectId(project);
    const response = await window.pywebview.api.save_survey_report_draft(
      projectId,
      buildCanonicalSurveyReportDraft(draft)
    );
    if (response?.status !== "success" || !response.data) {
      throw new Error(
        response?.message || "Unable to save utility plan draft."
      );
    }
    const activeFloorId = utilityPlanState.activeFloorId;
    const activeCalloutId = utilityPlanState.activeCalloutId;
    applyLoadedUtilityPlanDraft(response.data, projectId);
    utilityPlanState.activeFloorId = activeFloorId;
    ensureUtilityPlanActiveFloor();
    utilityPlanState.activeCalloutId = activeCalloutId;
    if (!quiet) setUtilityPlanStatus("Utility plan draft saved.");
    renderUtilityPlanEditorUi();
    return response.data;
  } catch (error) {
    reportClientError("Failed to save utility plan draft", error);
    if (!quiet) {
      setUtilityPlanStatus(
        error?.message || "Unable to save utility plan draft."
      );
    }
    return null;
  } finally {
    utilityPlanSaving = false;
    if (utilityPlanAutosaveQueued) {
      utilityPlanAutosaveQueued = false;
      debouncedSaveUtilityPlanDraft();
    }
  }
}

const debouncedSaveUtilityPlanDraft = debounce(() => {
  void saveUtilityPlanDraft({ quiet: true });
}, 500);

function queueUtilityPlanSave() {
  debouncedSaveUtilityPlanDraft();
}

function renderUtilityPlanToolbar() {
  [
    ["utilityPlanToolPanBtn", "pan"],
    ["utilityPlanToolCropBtn", "crop"],
    ["utilityPlanToolCalloutBtn", "callout"],
  ].forEach(([id, tool]) => {
    const btn = document.getElementById(id);
    if (btn) btn.classList.toggle("is-active", utilityPlanState.tool === tool);
  });
}

function renderUtilityPlanProjectControls() {
  const projectSearch = document.getElementById("utilityPlanProjectSearch");
  if (projectSearch) projectSearch.value = utilityPlanProjectQuery;
  const projectSelect = document.getElementById("utilityPlanProjectSelect");
  if (projectSelect && utilityPlanProjectIndex != null) {
    projectSelect.value = String(utilityPlanProjectIndex);
  }
}

function renderUtilityPlanFloorList() {
  const list = document.getElementById("utilityPlanFloorList");
  if (!list) return;
  const floors = utilityPlanState?.draft?.utilityPlan?.floors || [];
  list.innerHTML = "";
  if (!floors.length) {
    list.appendChild(
      el("div", {
        className: "tiny muted",
        textContent: "No floors added yet.",
      })
    );
    return;
  }
  floors.forEach((floor, index) => {
    const item = el(
      "button",
      {
        className: `utility-plan-floor-item ${floor.id === utilityPlanState.activeFloorId ? "is-active" : ""}`.trim(),
        type: "button",
        onclick: async () => {
          utilityPlanState.activeFloorId = floor.id;
          utilityPlanState.activeCalloutId = "";
          renderUtilityPlanEditorUi();
          try {
            await loadUtilityPlanPreviewForActiveFloor({ resetView: true });
          } catch (error) {
            reportClientError("Failed to change utility plan floor", error);
            setUtilityPlanStatus(
              error?.message || "Unable to change utility plan floor."
            );
          }
        },
      },
      []
    );
    item.appendChild(
      el("div", { className: "utility-plan-floor-copy" }, [
        el("span", {
          className: "utility-plan-floor-label",
          textContent: String(floor.label || `Floor ${index + 1}`),
        }),
        el("span", {
          className: "utility-plan-floor-meta",
          textContent: `Page ${Number(floor.pageNumber || 0) + 1} · ${floor.callouts.length} callout${floor.callouts.length === 1 ? "" : "s"}`,
        }),
      ])
    );
    item.appendChild(
      el("span", {
        className: "utility-plan-floor-meta",
        textContent: floor.exportPath ? "Exported" : "Draft",
      })
    );
    list.appendChild(item);
  });
}

function renderUtilityPlanFormControls() {
  const draft = utilityPlanState.draft;
  const floor = getActiveUtilityPlanFloor();
  const callout = getSelectedUtilityPlanCallout();
  const pdfInfo = utilityPlanState.pdfInfo;
  const sourcePdfInput = document.getElementById("utilityPlanSourcePdfPath");
  const exportFolderInput = document.getElementById(
    "utilityPlanExportFolderPath"
  );
  const pdfMeta = document.getElementById("utilityPlanPdfMeta");
  const floorLabelInput = document.getElementById("utilityPlanFloorLabelInput");
  const pageSelect = document.getElementById("utilityPlanPageSelect");
  const calloutText = document.getElementById("utilityPlanCalloutText");
  const calloutMeta = document.getElementById("utilityPlanCalloutMeta");

  const hasDraft = !!draft;
  const hasSourcePdf = !!draft?.utilityPlan?.sourcePdfPath;
  const hasFloor = !!floor;
  if (sourcePdfInput) {
    sourcePdfInput.value = draft?.utilityPlan?.sourcePdfPath || "";
    sourcePdfInput.title = sourcePdfInput.value;
  }
  if (exportFolderInput) {
    exportFolderInput.value = draft?.utilityPlan?.exportFolderPath || "";
    exportFolderInput.title = exportFolderInput.value;
  }
  if (pdfMeta) {
    if (!hasSourcePdf) {
      pdfMeta.textContent = "No PDF selected.";
    } else if (pdfInfo?.pageCount) {
      pdfMeta.textContent = `${pdfInfo.pageCount} page${pdfInfo.pageCount === 1 ? "" : "s"} loaded.`;
    } else {
      pdfMeta.textContent = "Loading PDF details...";
    }
  }
  if (floorLabelInput) {
    floorLabelInput.value = floor?.label || "";
    floorLabelInput.disabled = !hasFloor;
  }
  if (pageSelect) {
    pageSelect.innerHTML = "";
    const pages = Array.isArray(pdfInfo?.pages) ? pdfInfo.pages : [];
    if (!pages.length) {
      pageSelect.appendChild(
        el("option", {
          value: "0",
          textContent: hasSourcePdf ? "Loading pages..." : "Select a PDF first",
        })
      );
    } else {
      pages.forEach((page) => {
        pageSelect.appendChild(
          el("option", {
            value: String(page.pageNumber),
            textContent: `Page ${page.pageNumber + 1} · ${Math.round(page.width)} x ${Math.round(page.height)}`,
          })
        );
      });
    }
    pageSelect.disabled = !hasFloor || !pages.length;
    pageSelect.value = hasFloor ? String(floor.pageNumber || 0) : "0";
  }
  if (calloutText) {
    calloutText.value = callout?.text || "";
    calloutText.disabled = !callout;
    if (typeof autoResizeTextarea === "function") {
      autoResizeTextarea(calloutText);
    }
  }
  if (calloutMeta) {
    calloutMeta.textContent = callout
      ? "Drag the box or arrow target on the plan."
      : "No callout selected.";
  }

  const disabledWithoutDraft = [
    "utilityPlanBrowsePdfBtn",
    "utilityPlanClearPdfBtn",
    "utilityPlanAddFloorBtn",
    "utilityPlanMoveFloorUpBtn",
    "utilityPlanMoveFloorDownBtn",
    "utilityPlanRemoveFloorBtn",
    "utilityPlanBrowseExportFolderBtn",
    "utilityPlanExportCurrentBtn",
    "utilityPlanExportAllBtn",
  ];
  disabledWithoutDraft.forEach((id) => {
    const btn = document.getElementById(id);
    if (btn) btn.disabled = !hasDraft;
  });
  const duplicateBtn = document.getElementById("utilityPlanDuplicateCalloutBtn");
  const deleteBtn = document.getElementById("utilityPlanDeleteCalloutBtn");
  if (duplicateBtn) duplicateBtn.disabled = !callout;
  if (deleteBtn) deleteBtn.disabled = !callout;
}

function renderUtilityPlanViewport() {
  const emptyState = document.getElementById("utilityPlanEmptyState");
  const viewport = document.getElementById("utilityPlanViewport");
  const image = document.getElementById("utilityPlanPageImage");
  const preview = utilityPlanState.preview;
  if (!emptyState || !viewport || !image) return;
  if (!preview) {
    viewport.hidden = true;
    emptyState.hidden = false;
    return;
  }
  emptyState.hidden = true;
  viewport.hidden = false;
  if (image.src !== preview.dataUrl) image.src = preview.dataUrl;
  applyUtilityPlanViewportTransform();
  renderUtilityPlanOverlay();
}

function renderUtilityPlanEditorUi() {
  renderUtilityPlanProjectControls();
  renderUtilityPlanFloorList();
  renderUtilityPlanToolbar();
  renderUtilityPlanFormControls();
  renderUtilityPlanViewport();

  if (!utilityPlanStatusMessage) {
    const project = getActiveUtilityPlanProject();
    const floor = getActiveUtilityPlanFloor();
    if (!project) {
      setUtilityPlanStatus("Select a project to begin.");
    } else if (!utilityPlanState.draft?.utilityPlan?.sourcePdfPath) {
      setUtilityPlanStatus("Select the source PDF for this project.");
    } else if (!floor) {
      setUtilityPlanStatus("Add a floor to begin marking up the utility plan.");
    } else {
      setUtilityPlanStatus(
        `${getUtilityPlanProjectLabel(project, utilityPlanProjectIndex || 0)} · ${floor.label}`
      );
    }
  } else {
    const statusEl = document.getElementById("utilityPlanStatus");
    if (statusEl) statusEl.textContent = utilityPlanStatusMessage;
  }
}

function updateSelectedUtilityPlanCalloutText(text) {
  const callout = getSelectedUtilityPlanCallout();
  const preview = utilityPlanState.preview;
  if (!callout || !preview) return;
  const nextText = String(text || "").trim() || "Callout";
  const measured = measureUtilityPlanText(nextText);
  callout.text = nextText;
  callout.labelRect = fitUtilityPlanLabelRectToPage(
    {
      ...callout.labelRect,
      width: Math.max(callout.labelRect.width, measured.width),
      height: Math.max(callout.labelRect.height, measured.height),
    },
    preview.pageWidth,
    preview.pageHeight
  );
  renderUtilityPlanEditorUi();
  queueUtilityPlanSave();
}

function duplicateSelectedUtilityPlanCallout() {
  const floor = getActiveUtilityPlanFloor();
  const callout = getSelectedUtilityPlanCallout();
  const preview = utilityPlanState.preview;
  if (!floor || !callout || !preview) return;
  const duplicated = createDefaultUtilityPlanCallout(
    {
      ...callout,
      id: generateUtilityPlanCalloutId(),
      labelRect: {
        ...callout.labelRect,
        x: callout.labelRect.x + 18,
        y: callout.labelRect.y + 18,
      },
      targetPoint: {
        x: callout.targetPoint.x + 18,
        y: callout.targetPoint.y + 18,
      },
    },
    preview.pageWidth,
    preview.pageHeight
  );
  floor.callouts.push(duplicated);
  utilityPlanState.activeCalloutId = duplicated.id;
  renderUtilityPlanEditorUi();
  queueUtilityPlanSave();
}

function deleteSelectedUtilityPlanCallout() {
  const floor = getActiveUtilityPlanFloor();
  const callout = getSelectedUtilityPlanCallout();
  if (!floor || !callout) return;
  floor.callouts = floor.callouts.filter(
    (candidate) => candidate.id !== callout.id
  );
  utilityPlanState.activeCalloutId = "";
  renderUtilityPlanEditorUi();
  queueUtilityPlanSave();
}

function getUtilityPlanViewportPoint(clientX, clientY) {
  const viewport = document.getElementById("utilityPlanViewport");
  const preview = utilityPlanState.preview;
  if (!viewport || !preview) return null;
  const rect = viewport.getBoundingClientRect();
  return clampUtilityPlanPointToPage(
    {
      x: (clientX - rect.left - utilityPlanState.panX) / utilityPlanState.zoom,
      y: (clientY - rect.top - utilityPlanState.panY) / utilityPlanState.zoom,
    },
    preview.pageWidth,
    preview.pageHeight
  );
}

function setUtilityPlanInteraction(interaction) {
  utilityPlanState.interaction = interaction;
  applyUtilityPlanViewportTransform();
}

function updateUtilityPlanCropResize(originalRect, handle, point, pageWidth, pageHeight) {
  const minSize = 24;
  let left = originalRect.x;
  let top = originalRect.y;
  let right = originalRect.x + originalRect.width;
  let bottom = originalRect.y + originalRect.height;

  if (handle.includes("w")) left = point.x;
  if (handle.includes("e")) right = point.x;
  if (handle.includes("n")) top = point.y;
  if (handle.includes("s")) bottom = point.y;

  left = Math.min(Math.max(left, 0), pageWidth);
  right = Math.min(Math.max(right, 0), pageWidth);
  top = Math.min(Math.max(top, 0), pageHeight);
  bottom = Math.min(Math.max(bottom, 0), pageHeight);

  if (right - left < minSize) {
    if (handle.includes("w")) left = right - minSize;
    else right = left + minSize;
  }
  if (bottom - top < minSize) {
    if (handle.includes("n")) top = bottom - minSize;
    else bottom = top + minSize;
  }

  return clampUtilityPlanRectToPage(
    { x: left, y: top, width: right - left, height: bottom - top },
    pageWidth,
    pageHeight
  );
}

function startUtilityPlanPointerInteraction(event) {
  const preview = utilityPlanState.preview;
  const floor = getActiveUtilityPlanFloor();
  const viewport = document.getElementById("utilityPlanViewport");
  if (!preview || !floor || !viewport) return;
  if (event.button !== 0 && event.button !== 1) return;

  const point = getUtilityPlanViewportPoint(event.clientX, event.clientY);
  if (!point) return;

  const targetEl = event.target.closest?.("[data-callout-target='true']");
  if (targetEl) {
    utilityPlanState.activeCalloutId =
      targetEl.getAttribute("data-callout-id") || "";
    setUtilityPlanInteraction({ type: "moveCalloutTarget" });
    renderUtilityPlanEditorUi();
    event.preventDefault();
    return;
  }

  const labelEl = event.target.closest?.("[data-callout-label='true']");
  if (labelEl) {
    utilityPlanState.activeCalloutId =
      labelEl.getAttribute("data-callout-id") || "";
    const callout = getSelectedUtilityPlanCallout();
    if (callout) {
      setUtilityPlanInteraction({
        type: "moveCalloutLabel",
        startPoint: point,
        originalRect: { ...callout.labelRect },
      });
      renderUtilityPlanEditorUi();
      event.preventDefault();
      return;
    }
  }

  const cropHandle = event.target.closest?.("[data-crop-handle]");
  if (cropHandle) {
    setUtilityPlanInteraction({
      type: "resizeCrop",
      handle: cropHandle.getAttribute("data-crop-handle"),
      originalRect: { ...floor.cropRect },
    });
    event.preventDefault();
    return;
  }

  if (event.target.closest?.("[data-crop-body='true']")) {
    setUtilityPlanInteraction({
      type: "moveCrop",
      startPoint: point,
      originalRect: { ...floor.cropRect },
    });
    event.preventDefault();
    return;
  }

  if (utilityPlanState.tool === "pan" || event.button === 1) {
    setUtilityPlanInteraction({
      type: "pan",
      startClientX: event.clientX,
      startClientY: event.clientY,
      originPanX: utilityPlanState.panX,
      originPanY: utilityPlanState.panY,
    });
    event.preventDefault();
    return;
  }

  if (utilityPlanState.tool === "crop") {
    floor.cropRect = { x: point.x, y: point.y, width: 1, height: 1 };
    setUtilityPlanInteraction({
      type: "drawCrop",
      startPoint: point,
    });
    renderUtilityPlanOverlay();
    event.preventDefault();
    return;
  }

  if (utilityPlanState.tool === "callout") {
    const callout = createDefaultUtilityPlanCallout(
      { targetPoint: point },
      preview.pageWidth,
      preview.pageHeight
    );
    floor.callouts.push(callout);
    utilityPlanState.activeCalloutId = callout.id;
    renderUtilityPlanEditorUi();
    queueUtilityPlanSave();
    event.preventDefault();
    return;
  }

  utilityPlanState.activeCalloutId = "";
  renderUtilityPlanEditorUi();
}

function updateUtilityPlanInteraction(event) {
  const interaction = utilityPlanState.interaction;
  const preview = utilityPlanState.preview;
  const floor = getActiveUtilityPlanFloor();
  if (!interaction || !preview || !floor) return;

  if (interaction.type === "pan") {
    utilityPlanState.panX =
      interaction.originPanX + (event.clientX - interaction.startClientX);
    utilityPlanState.panY =
      interaction.originPanY + (event.clientY - interaction.startClientY);
    applyUtilityPlanViewportTransform();
    return;
  }

  const point = getUtilityPlanViewportPoint(event.clientX, event.clientY);
  if (!point) return;

  if (interaction.type === "drawCrop") {
    const x1 = interaction.startPoint.x;
    const y1 = interaction.startPoint.y;
    floor.cropRect = clampUtilityPlanRectToPage(
      {
        x: Math.min(x1, point.x),
        y: Math.min(y1, point.y),
        width: Math.max(24, Math.abs(point.x - x1)),
        height: Math.max(24, Math.abs(point.y - y1)),
      },
      preview.pageWidth,
      preview.pageHeight
    );
    renderUtilityPlanOverlay();
    return;
  }

  if (interaction.type === "moveCrop") {
    const dx = point.x - interaction.startPoint.x;
    const dy = point.y - interaction.startPoint.y;
    floor.cropRect = clampUtilityPlanRectToPage(
      {
        x: interaction.originalRect.x + dx,
        y: interaction.originalRect.y + dy,
        width: interaction.originalRect.width,
        height: interaction.originalRect.height,
      },
      preview.pageWidth,
      preview.pageHeight
    );
    renderUtilityPlanOverlay();
    return;
  }

  if (interaction.type === "resizeCrop") {
    floor.cropRect = updateUtilityPlanCropResize(
      interaction.originalRect,
      String(interaction.handle || "se"),
      point,
      preview.pageWidth,
      preview.pageHeight
    );
    renderUtilityPlanOverlay();
    return;
  }

  const callout = getSelectedUtilityPlanCallout();
  if (!callout) return;

  if (interaction.type === "moveCalloutLabel") {
    const dx = point.x - interaction.startPoint.x;
    const dy = point.y - interaction.startPoint.y;
    callout.labelRect = fitUtilityPlanLabelRectToPage(
      {
        ...interaction.originalRect,
        x: interaction.originalRect.x + dx,
        y: interaction.originalRect.y + dy,
      },
      preview.pageWidth,
      preview.pageHeight
    );
    renderUtilityPlanOverlay();
    renderUtilityPlanFormControls();
    return;
  }

  if (interaction.type === "moveCalloutTarget") {
    callout.targetPoint = clampUtilityPlanPointToPage(
      point,
      preview.pageWidth,
      preview.pageHeight
    );
    renderUtilityPlanOverlay();
  }
}

function finishUtilityPlanInteraction({ persist = true } = {}) {
  if (!utilityPlanState.interaction) return;
  utilityPlanState.interaction = null;
  applyUtilityPlanViewportTransform();
  if (persist) queueUtilityPlanSave();
}

function handleUtilityPlanViewportWheel(event) {
  if (!utilityPlanState.preview) return;
  event.preventDefault();
  const factor = event.deltaY < 0 ? 1.1 : 0.9;
  updateUtilityPlanZoom(
    utilityPlanState.zoom * factor,
    event.clientX,
    event.clientY
  );
}

async function browseUtilityPlanSourcePdf() {
  const draft = utilityPlanState.draft;
  if (!draft) {
    toast("Select a project first.");
    return;
  }
  if (!window.pywebview?.api?.select_utility_plan_pdf) {
    toast("PDF picker is unavailable.");
    return;
  }
  try {
    const hasAnnotations = draft.utilityPlan.floors.some(
      (floor) =>
        floor.callouts.length > 0 ||
        (floor.cropRect.width > 0 && floor.cropRect.height > 0)
    );
    const response = await window.pywebview.api.select_utility_plan_pdf();
    if (response?.status !== "success" || !response.path) return;
    const nextPath = normalizeWindowsPath(response.path);
    const previousPath = normalizeWindowsPath(draft.utilityPlan.sourcePdfPath);
    if (hasAnnotations && previousPath && previousPath !== nextPath) {
      const proceed = confirm(
        "Changing the source PDF will reset crop boxes and callouts for all floors.\n\nPress OK to continue."
      );
      if (!proceed) return;
    }
    const pdfInfo = await loadUtilityPlanPdfInfo(nextPath);
    draft.utilityPlan.sourcePdfPath = nextPath;
    draft.utilityPlan.exportFolderPath =
      draft.utilityPlan.exportFolderPath || getFolderFromPath(nextPath);
    draft.utilityPlan.floors = draft.utilityPlan.floors.map((floor, index) => ({
      ...createDefaultUtilityPlanFloor({
        id: floor.id,
        label: floor.label || `Floor ${index + 1}`,
        order: index,
        pageNumber: Math.min(
          Math.max(Number(floor.pageNumber || 0), 0),
          Math.max(0, pdfInfo.pageCount - 1)
        ),
        cropRect: {
          x: 0,
          y: 0,
          width: pdfInfo.pages[0]?.width || 0,
          height: pdfInfo.pages[0]?.height || 0,
        },
        callouts: previousPath && previousPath !== nextPath ? [] : floor.callouts,
      }),
    }));
    utilityPlanState.pdfInfo = pdfInfo;
    ensureUtilityPlanActiveFloor();
    renderUtilityPlanEditorUi();
    await loadUtilityPlanPreviewForActiveFloor({ resetView: true });
    queueUtilityPlanSave();
    setUtilityPlanStatus("Source PDF updated.");
  } catch (error) {
    reportClientError("Failed to select utility plan PDF", error);
    setUtilityPlanStatus(error?.message || "Unable to select PDF.");
  }
}

function clearUtilityPlanSourcePdf() {
  const draft = utilityPlanState.draft;
  if (!draft) return;
  const hasAnnotations = draft.utilityPlan.floors.some(
    (floor) =>
      floor.callouts.length > 0 ||
      (floor.cropRect.width > 0 && floor.cropRect.height > 0)
  );
  if (
    hasAnnotations &&
    !confirm(
      "Clearing the source PDF will remove the floor crops and callouts for this draft.\n\nPress OK to continue."
    )
  ) {
    return;
  }
  draft.utilityPlan.sourcePdfPath = "";
  draft.utilityPlan.floors = draft.utilityPlan.floors.map((floor, index) =>
    createDefaultUtilityPlanFloor({
      id: floor.id,
      label: floor.label || `Floor ${index + 1}`,
      order: index,
    })
  );
  utilityPlanState.pdfInfo = null;
  utilityPlanState.activeCalloutId = "";
  clearUtilityPlanPreview();
  renderUtilityPlanEditorUi();
  queueUtilityPlanSave();
  setUtilityPlanStatus("Source PDF cleared.");
}

async function browseUtilityPlanExportFolder() {
  const draft = utilityPlanState.draft;
  if (!draft || !window.pywebview?.api?.select_template_output_folder) return;
  try {
    const defaultDir =
      draft.utilityPlan.exportFolderPath ||
      getFolderFromPath(draft.utilityPlan.sourcePdfPath || "") ||
      "";
    const response = await window.pywebview.api.select_template_output_folder(
      defaultDir
    );
    if (response?.status !== "success" || !response.path) return;
    draft.utilityPlan.exportFolderPath = normalizeWindowsPath(response.path);
    renderUtilityPlanEditorUi();
    queueUtilityPlanSave();
    setUtilityPlanStatus("Export folder updated.");
  } catch (error) {
    reportClientError("Failed to select export folder", error);
    setUtilityPlanStatus(error?.message || "Unable to select export folder.");
  }
}

async function addUtilityPlanFloor() {
  const draft = utilityPlanState.draft;
  if (!draft) return;
  const preview = utilityPlanState.preview;
  const nextIndex = draft.utilityPlan.floors.length + 1;
  const sourceFloor = getActiveUtilityPlanFloor();
  const newFloor = createDefaultUtilityPlanFloor({
    label: `Floor ${nextIndex}`,
    pageNumber: Number(sourceFloor?.pageNumber || 0),
    order: draft.utilityPlan.floors.length,
  });
  if (preview) {
    newFloor.cropRect = {
      x: 0,
      y: 0,
      width: preview.pageWidth,
      height: preview.pageHeight,
    };
  }
  draft.utilityPlan.floors.push(newFloor);
  utilityPlanState.activeFloorId = newFloor.id;
  utilityPlanState.activeCalloutId = "";
  renderUtilityPlanEditorUi();
  if (draft.utilityPlan.sourcePdfPath) {
    try {
      await loadUtilityPlanPreviewForActiveFloor({ resetView: true });
    } catch (error) {
      reportClientError("Failed to add floor", error);
    }
  }
  queueUtilityPlanSave();
}

function removeUtilityPlanFloor() {
  const draft = utilityPlanState.draft;
  const floor = getActiveUtilityPlanFloor();
  if (!draft || !floor) return;
  if (draft.utilityPlan.floors.length <= 1) {
    toast("At least one floor is required.");
    return;
  }
  draft.utilityPlan.floors = draft.utilityPlan.floors
    .filter((candidate) => candidate.id !== floor.id)
    .map((candidate, index) => ({ ...candidate, order: index }));
  utilityPlanState.activeFloorId = draft.utilityPlan.floors[0]?.id || "";
  utilityPlanState.activeCalloutId = "";
  renderUtilityPlanEditorUi();
  void loadUtilityPlanPreviewForActiveFloor({ resetView: true });
  queueUtilityPlanSave();
}

function moveUtilityPlanFloor(direction) {
  const draft = utilityPlanState.draft;
  const floor = getActiveUtilityPlanFloor();
  if (!draft || !floor) return;
  const currentIndex = draft.utilityPlan.floors.findIndex(
    (candidate) => candidate.id === floor.id
  );
  if (currentIndex < 0) return;
  const nextIndex = currentIndex + direction;
  if (nextIndex < 0 || nextIndex >= draft.utilityPlan.floors.length) return;
  const [moved] = draft.utilityPlan.floors.splice(currentIndex, 1);
  draft.utilityPlan.floors.splice(nextIndex, 0, moved);
  draft.utilityPlan.floors = draft.utilityPlan.floors.map((candidate, index) => ({
    ...candidate,
    order: index,
  }));
  renderUtilityPlanEditorUi();
  queueUtilityPlanSave();
}

async function updateUtilityPlanFloorPage(pageNumber) {
  const floor = getActiveUtilityPlanFloor();
  if (!floor) return;
  floor.pageNumber = Math.max(
    0,
    Math.trunc(normalizeUtilityPlanNumber(pageNumber, 0))
  );
  floor.callouts = [];
  floor.exportPath = "";
  floor.lastExportedAtUtc = "";
  utilityPlanState.activeCalloutId = "";
  renderUtilityPlanEditorUi();
  try {
    await loadUtilityPlanPreviewForActiveFloor({ resetView: true });
    queueUtilityPlanSave();
  } catch (error) {
    reportClientError("Failed to update utility plan page", error);
    setUtilityPlanStatus(error?.message || "Unable to load the selected page.");
  }
}

function updateUtilityPlanFloorLabel(label) {
  const floor = getActiveUtilityPlanFloor();
  if (!floor) return;
  floor.label = String(label || "").trim() || floor.label || "Floor";
  renderUtilityPlanFloorList();
  queueUtilityPlanSave();
}

async function exportUtilityPlanFloors(mode = "current") {
  const project = getActiveUtilityPlanProject();
  const draft = utilityPlanState.draft;
  const floor = getActiveUtilityPlanFloor();
  if (!project || !draft) {
    toast("Select a project first.");
    return;
  }
  if (!draft.utilityPlan.sourcePdfPath) {
    toast("Select a source PDF first.");
    return;
  }
  let exportFolderPath = draft.utilityPlan.exportFolderPath;
  if (!exportFolderPath) {
    await browseUtilityPlanExportFolder();
    exportFolderPath = draft.utilityPlan.exportFolderPath;
    if (!exportFolderPath) return;
  }

  try {
    setUtilityPlanStatus("Exporting utility plan PNGs...");
    const floorIds =
      mode === "all"
        ? draft.utilityPlan.floors.map((item) => item.id)
        : floor
          ? [floor.id]
          : [];
    const response = await window.pywebview.api.export_utility_plan_pngs(
      getUtilityPlanProjectId(project),
      {
        draft: buildCanonicalSurveyReportDraft(draft),
        floorIds,
        exportFolderPath,
      }
    );
    if (response?.status !== "success") {
      throw new Error(
        response?.message || "Unable to export utility plan PNGs."
      );
    }
    const previousFloorId = utilityPlanState.activeFloorId;
    const previousCalloutId = utilityPlanState.activeCalloutId;
    applyLoadedUtilityPlanDraft(response.data, getUtilityPlanProjectId(project));
    utilityPlanState.activeFloorId = previousFloorId;
    ensureUtilityPlanActiveFloor();
    utilityPlanState.activeCalloutId = previousCalloutId;
    renderUtilityPlanEditorUi();
    const exportCount = Array.isArray(response.exports)
      ? response.exports.length
      : 0;
    setUtilityPlanStatus(
      `Exported ${exportCount} utility plan PNG${exportCount === 1 ? "" : "s"}.`
    );
    toast(`Exported ${exportCount} utility plan PNG${exportCount === 1 ? "" : "s"}.`);
  } catch (error) {
    reportClientError("Failed to export utility plan PNGs", error);
    setUtilityPlanStatus(error?.message || "Unable to export utility plan PNGs.");
    toast("Utility plan export failed.");
  }
}

async function openUtilityPlanEditor(launchContext = null) {
  const dlg = document.getElementById("utilityPlanEditorDlg");
  if (!dlg) return;
  renderUtilityPlanProjectOptions(utilityPlanProjectQuery);
  const launchProjectIndex = findUtilityPlanProjectIndexFromLaunchContext(
    launchContext
  );
  if (launchProjectIndex != null && launchProjectIndex !== utilityPlanProjectIndex) {
    await setUtilityPlanProject(launchProjectIndex, { quiet: true });
  } else if (
    utilityPlanProjectIndex != null &&
    db[utilityPlanProjectIndex] &&
    !utilityPlanState.draft
  ) {
    await setUtilityPlanProject(utilityPlanProjectIndex, { quiet: true });
  } else {
    renderUtilityPlanEditorUi();
  }
  if (!dlg.open) dlg.showModal();
}

function closeUtilityPlanEditor() {
  finishUtilityPlanInteraction({ persist: false });
  void saveUtilityPlanDraft({ quiet: true });
  const dlg = document.getElementById("utilityPlanEditorDlg");
  if (dlg?.open) dlg.close();
}

const utilityPlanToolBtn = document.getElementById("toolUtilityPlanEditor");
if (utilityPlanToolBtn) {
  const openUtilityPlanHandler = async () => {
    await openUtilityPlanEditor(resolveCadLaunchContextForTool());
  };
  utilityPlanToolBtn.addEventListener("click", openUtilityPlanHandler);
  utilityPlanToolBtn.addEventListener("keydown", (event) => {
    if (event.key === "Enter" || event.key === " ") {
      event.preventDefault();
      openUtilityPlanHandler();
    }
  });
}

document.getElementById("utilityPlanCloseBtn")?.addEventListener("click", closeUtilityPlanEditor);
document.getElementById("utilityPlanEditorDlg")?.addEventListener("close", () => {
  finishUtilityPlanInteraction({ persist: false });
});
document.getElementById("utilityPlanProjectSearch")?.addEventListener("input", (event) => {
  utilityPlanProjectQuery = event.target.value;
  renderUtilityPlanProjectOptions(utilityPlanProjectQuery);
});
document.getElementById("utilityPlanProjectSelect")?.addEventListener("change", async (event) => {
  const index = Number(event.target.value);
  if (!Number.isNaN(index)) {
    await setUtilityPlanProject(index);
  }
});
document.getElementById("utilityPlanBrowsePdfBtn")?.addEventListener("click", browseUtilityPlanSourcePdf);
document.getElementById("utilityPlanClearPdfBtn")?.addEventListener("click", clearUtilityPlanSourcePdf);
document.getElementById("utilityPlanAddFloorBtn")?.addEventListener("click", () => {
  void addUtilityPlanFloor();
});
document.getElementById("utilityPlanMoveFloorUpBtn")?.addEventListener("click", () => moveUtilityPlanFloor(-1));
document.getElementById("utilityPlanMoveFloorDownBtn")?.addEventListener("click", () => moveUtilityPlanFloor(1));
document.getElementById("utilityPlanRemoveFloorBtn")?.addEventListener("click", removeUtilityPlanFloor);
document.getElementById("utilityPlanFloorLabelInput")?.addEventListener("input", (event) => {
  updateUtilityPlanFloorLabel(event.target.value);
});
document.getElementById("utilityPlanPageSelect")?.addEventListener("change", (event) => {
  void updateUtilityPlanFloorPage(event.target.value);
});
document.getElementById("utilityPlanBrowseExportFolderBtn")?.addEventListener("click", () => {
  void browseUtilityPlanExportFolder();
});
document.getElementById("utilityPlanExportCurrentBtn")?.addEventListener("click", () => {
  void exportUtilityPlanFloors("current");
});
document.getElementById("utilityPlanExportAllBtn")?.addEventListener("click", () => {
  void exportUtilityPlanFloors("all");
});
document.getElementById("utilityPlanCalloutText")?.addEventListener("input", (event) => {
  updateSelectedUtilityPlanCalloutText(event.target.value);
});
document.getElementById("utilityPlanDuplicateCalloutBtn")?.addEventListener("click", duplicateSelectedUtilityPlanCallout);
document.getElementById("utilityPlanDeleteCalloutBtn")?.addEventListener("click", deleteSelectedUtilityPlanCallout);
document.getElementById("utilityPlanToolPanBtn")?.addEventListener("click", () => {
  utilityPlanState.tool = "pan";
  renderUtilityPlanToolbar();
});
document.getElementById("utilityPlanToolCropBtn")?.addEventListener("click", () => {
  utilityPlanState.tool = "crop";
  renderUtilityPlanToolbar();
});
document.getElementById("utilityPlanToolCalloutBtn")?.addEventListener("click", () => {
  utilityPlanState.tool = "callout";
  renderUtilityPlanToolbar();
});
document.getElementById("utilityPlanZoomInBtn")?.addEventListener("click", () => {
  updateUtilityPlanZoom(utilityPlanState.zoom * 1.1);
});
document.getElementById("utilityPlanZoomOutBtn")?.addEventListener("click", () => {
  updateUtilityPlanZoom(utilityPlanState.zoom * 0.9);
});
document.getElementById("utilityPlanZoomResetBtn")?.addEventListener("click", fitUtilityPlanViewToPage);
document.getElementById("utilityPlanViewport")?.addEventListener("pointerdown", startUtilityPlanPointerInteraction);
document.getElementById("utilityPlanViewport")?.addEventListener("wheel", handleUtilityPlanViewportWheel, { passive: false });
window.addEventListener("pointermove", updateUtilityPlanInteraction);
window.addEventListener("pointerup", () => finishUtilityPlanInteraction({ persist: true }));
window.addEventListener("pointercancel", () => finishUtilityPlanInteraction({ persist: true }));

renderUtilityPlanEditorUi();
